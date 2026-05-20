# CLAUDE.md

This file provides guidance to Claude Code when working with this repository.

## Project Overview

Amazon Toolkit — Amazon 电商自动化工具集，提供 Flask Web UI 和 CLI 两种使用方式。

- **运行入口**：`python app.py`（Web UI，端口 5002）
- **访问地址**：`http://localhost:5002`
- **部署脚本**：`./deploy.sh`（服务器拉取代码并重启）

## Architecture

### Tool Registration Pattern

每个工具在 `app.py` 的 `TOOLS` 字典中注册，包含以下字段：

```python
TOOLS = {
    'tool_id': {
        'name':       '工具显示名称',
        'desc':       '工具描述',
        'input_dir':  BASE_DIR / 'data' / 'tool_id' / 'input',
        'output_dir': BASE_DIR / 'data' / 'tool_id' / 'output',
        'params':     [],           # 前端参数配置，支持 input / select / textarea 类型
        'module':     'tools.tool_id',
        'note':       '前端提示文字',
        'guide':      [('标签', '说明'), ...],  # 可选，功能介绍展开面板
    },
}
```

### Tool Module Pattern

每个工具位于 `tools/` 目录，暴露 `run()` 函数：

```python
def run(**kwargs):
    # kwargs 由 app.py _run_tool 注入
    ...
```

### Param Types（前端参数类型）

| type | 渲染形式 | 说明 |
|------|---------|------|
| `input`（默认） | `<input>` | 普通文本输入 |
| `select` | `<select>` | 下拉选择，需配置 `options` 和 `default` |
| `textarea` | `<textarea>` | 多行文本，用于批量输入（如 ASIN 列表） |

### Backend Execution Flow

1. `POST /api/run/<tool_id>` → 创建后台线程
2. `_run_tool()` 捕获 stdout，写入 `queue.Queue`
3. `GET /api/stream/<task_id>` → SSE 实时推送日志到前端

### Frontend

- 单页应用，纯原生 JS，无框架依赖
- 深色 / 浅色主题切换，偏好持久化到 `localStorage`
- 日志窗口支持上下拖拽 resize
- 所有 API 路径使用相对路径（`api/...`），支持 Nginx 子路径代理

## Tools

| 模块 | 功能 |
|------|------|
| `tools/weekly_data_clean.py` | 领星 MSKU 数据清洗，促销折扣求和 |
| `tools/weekly_report_append.py` | 追加写入周报主表；Windows 用 xlwings，Linux 用 openpyxl |
| `tools/ad_bulk_update.py` | Bulk Sheet 广告竞价批量更新 |
| `tools/asin_variant_score.py` | 西柚 API 查询 ASIN 变体流量得分，支持 US/UK/DE/JP |

## Xiyou API（西柚 API）

- Base URL：`https://openapi.xiyouzhaoci.com`
- 签名方式：`SHA256(CLIENT_ID + timestamp + CLIENT_SECRET + json_body)`
- **注意**：json_body 序列化必须使用 `json.dumps(body, separators=(',',':'), sort_keys=True)`
- 变体接口：`POST /v1/asins/variations`
- 流量得分：`POST /v1/asins/traffic`（每批最多 100 个 ASIN）

## Development Conventions

- 修改代码后**不直接推送**，等用户明确指示后再 `git push`
- 数据文件（xlsx/xls）不提交到 git（见 `.gitignore`）
- 分支：日常开发在 `dev`，稳定后合并 `master`
- 端口：本地/服务器统一使用 `5002`
