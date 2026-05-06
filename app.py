"""
app.py — Amazon Toolkit Web 服务
运行方式: python app.py
访问地址: http://localhost:5000
"""

import io
import os
import queue
import sys
import threading
import time
import uuid
from pathlib import Path

from flask import Flask, Response, jsonify, render_template, request, send_from_directory

BASE_DIR = Path(__file__).resolve().parent

app = Flask(__name__, template_folder='web/templates', static_folder='web/static')
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB 上传限制

# 任务状态存储 { task_id: { status, queue, log } }
_tasks: dict = {}
_tasks_lock = threading.Lock()


# ── 工具注册表 ──────────────────────────────────────────

TOOLS = {
    'weekly_data_clean': {
        'name':    '周报数据清洗',
        'desc':    '领星 MSKU 产品表现数据清洗，生成标签汇总文件',
        'input_dir': BASE_DIR / 'data' / 'weekly_data_clean' / 'input',
        'output_dir': BASE_DIR / 'data' / 'weekly_data_clean' / 'output',
        'params': [
            {'key': 'week_number', 'label': '周次', 'placeholder': '如 26W17', 'required': True}
        ],
        'module': 'tools.weekly_data_clean',
        'note':   '输入文件须命名为 产品表现MSKU-{周次}.xlsx',
    },
    'weekly_report_append': {
        'name':    '周报数据写入',
        'desc':    '将清洗结果追加写入周销售统计主表，自动插入新数据行',
        'input_dir': BASE_DIR / 'data' / 'weekly_report_append' / 'input',
        'output_dir': BASE_DIR / 'data' / 'weekly_report_append' / 'output',
        'params': [
            {'key': 'week_number', 'label': '周次', 'placeholder': '如 26W17', 'required': True}
        ],
        'module': 'tools.weekly_report_append',
        'note':   '需先完成周报数据清洗（步骤一）。输入文件须命名为 周销售数据统计US-{周次}.xlsx',
        'depends': 'weekly_data_clean',
    },
    'ad_bulk_update': {
        'name':    '广告竞价更新',
        'desc':    '批量更新 Bulk Sheet 竞价，自动标注高 ACoS / 高点击不出单关键词',
        'input_dir': BASE_DIR / 'data' / 'ad_bulk_update' / 'input',
        'output_dir': BASE_DIR / 'data' / 'ad_bulk_update' / 'output',
        'params': [],
        'module': 'tools.ad_bulk_update',
        'note':   '输入目录须包含 BulkSheetExport*.xlsx 及 targeting_labels_*.xlsx 文件',
    },
}

# 确保所有目录存在
for t in TOOLS.values():
    t['input_dir'].mkdir(parents=True, exist_ok=True)
    t['output_dir'].mkdir(parents=True, exist_ok=True)


# ── 日志捕获 ─────────────────────────────────────────────

class _QueueWriter(io.TextIOBase):
    """将 print() 输出重定向到 queue"""
    def __init__(self, q: queue.Queue):
        self._q = q

    def write(self, s: str):
        if s.strip():
            self._q.put(('log', s.rstrip()))
        return len(s)

    def flush(self):
        pass


def _run_tool(task_id: str, tool_id: str, params: dict):
    """在子线程中执行工具，捕获输出写入队列"""
    q = _tasks[task_id]['queue']
    old_stdout = sys.stdout

    try:
        import importlib
        sys.stdout = _QueueWriter(q)

        tool = TOOLS[tool_id]
        module = importlib.import_module(tool['module'])

        # 注入参数（替代 input() 调用）
        if tool_id in ('weekly_data_clean', 'weekly_report_append'):
            week = params.get('week_number', '').strip()
            if not week:
                q.put(('log', '❌ 请填写周次'))
                q.put(('done', 'error'))
                return
            # monkey-patch builtins.input for this thread
            import builtins
            _orig_input = builtins.input
            builtins.input = lambda _prompt='': week
            try:
                module.run()
            finally:
                builtins.input = _orig_input
        else:
            module.run()

        q.put(('done', 'success'))

    except Exception as e:
        q.put(('log', f'❌ 执行出错: {e}'))
        q.put(('done', 'error'))
    finally:
        sys.stdout = old_stdout
        with _tasks_lock:
            _tasks[task_id]['status'] = 'done'


# ── 路由 ─────────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/tools')
def api_tools():
    result = {}
    for tid, t in TOOLS.items():
        result[tid] = {
            'name':    t['name'],
            'desc':    t['desc'],
            'params':  t['params'],
            'note':    t.get('note', ''),
            'depends': t.get('depends', None),
        }
    return jsonify(result)


@app.route('/api/files/<tool_id>')
def api_files(tool_id):
    if tool_id not in TOOLS:
        return jsonify({'error': '工具不存在'}), 404
    tool = TOOLS[tool_id]

    def list_dir(d: Path):
        if not d.exists():
            return []
        files = []
        for f in sorted(d.iterdir()):
            if f.is_file() and not f.name.startswith('.'):
                files.append({'name': f.name, 'size': f.stat().st_size, 'mtime': f.stat().st_mtime})
        return files

    return jsonify({
        'input':  list_dir(tool['input_dir']),
        'output': list_dir(tool['output_dir']),
    })


@app.route('/api/upload/<tool_id>', methods=['POST'])
def api_upload(tool_id):
    if tool_id not in TOOLS:
        return jsonify({'error': '工具不存在'}), 404
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': '没有文件'}), 400
    saved = []
    for f in files:
        dest = TOOLS[tool_id]['input_dir'] / f.filename
        f.save(dest)
        saved.append(f.filename)
    return jsonify({'saved': saved})


@app.route('/api/delete/<tool_id>', methods=['POST'])
def api_delete(tool_id):
    if tool_id not in TOOLS:
        return jsonify({'error': '工具不存在'}), 404
    data = request.json or {}
    fname = data.get('filename', '')
    zone  = data.get('zone', 'input')  # 'input' 或 'output'
    base  = TOOLS[tool_id]['input_dir'] if zone == 'input' else TOOLS[tool_id]['output_dir']
    target = base / fname
    if target.exists() and target.is_file():
        target.unlink()
        return jsonify({'ok': True})
    return jsonify({'error': '文件不存在'}), 404


@app.route('/api/download/<tool_id>/<filename>')
def api_download(tool_id, filename):
    if tool_id not in TOOLS:
        return jsonify({'error': '工具不存在'}), 404
    return send_from_directory(TOOLS[tool_id]['output_dir'], filename, as_attachment=True)


@app.route('/api/run/<tool_id>', methods=['POST'])
def api_run(tool_id):
    if tool_id not in TOOLS:
        return jsonify({'error': '工具不存在'}), 404
    params = request.json or {}
    task_id = str(uuid.uuid4())
    q: queue.Queue = queue.Queue()
    with _tasks_lock:
        _tasks[task_id] = {'status': 'running', 'queue': q, 'tool': tool_id}
    thread = threading.Thread(target=_run_tool, args=(task_id, tool_id, params), daemon=True)
    thread.start()
    return jsonify({'task_id': task_id})


@app.route('/api/stream/<task_id>')
def api_stream(task_id):
    """SSE 实时日志流"""
    def generate():
        if task_id not in _tasks:
            yield 'data: {"type":"error","msg":"任务不存在"}\n\n'
            return
        q = _tasks[task_id]['queue']
        while True:
            try:
                kind, payload = q.get(timeout=30)
                if kind == 'log':
                    import json
                    yield f'data: {json.dumps({"type":"log","msg":payload})}\n\n'
                elif kind == 'done':
                    import json
                    yield f'data: {json.dumps({"type":"done","status":payload})}\n\n'
                    break
            except queue.Empty:
                yield 'data: {"type":"ping"}\n\n'

    return Response(generate(), mimetype='text/event-stream',
                    headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'})


if __name__ == '__main__':
    print('=' * 48)
    print('  Amazon Toolkit — 本地 Web 服务')
    print('  访问地址: http://localhost:5000')
    print('=' * 48)
    app.run(host='0.0.0.0', port=5000, debug=False, threaded=True)