# Amazon Toolkit

Amazon 电商自动化工具集，提供 Web UI 和 CLI 两种使用方式。

## 工具列表

| 工具 | 说明 |
|------|------|
| 周报数据清洗 | 领星 MSKU 产品表现数据清洗，生成标签汇总文件 |
| 周报数据写入 | 将清洗结果追加写入周销售统计主表 |
| 广告竞价更新 | 批量更新 Bulk Sheet 竞价，标注高 ACoS / 高点击不出单关键词 |
| ASIN变体得分查询 | 查询 ASIN 变体及西柚流量得分，输出标准 Excel 格式 |

---

## 本地运行

```bash
# 安装依赖
pip install -r requirements.txt

# 启动 Web UI（推荐）
python app.py
# 访问 http://localhost:5002

# 或使用 CLI
python main.py
```

> Windows 上 `weekly_report_append` 使用 xlwings（需要安装 Excel），Linux/Mac 自动切换为 openpyxl。

---

## 云服务器部署

### 首次部署

```bash
# 克隆仓库
git clone git@github.com:yssdnj/amazon_toolkit.git
cd amazon_toolkit

# 创建虚拟环境并安装依赖
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# 后台启动服务
nohup python3 app.py > app.log 2>&1 &
echo "服务已启动，访问 http://<服务器IP>:5002"
```

### 更新代码并重启服务

```bash
cd ~/amazon_toolkit

# 拉取最新代码
git pull origin dev

# 激活虚拟环境（如未激活）
source venv/bin/activate

# 更新依赖（如 requirements.txt 有变动）
pip install -r requirements.txt

# 重启服务
pkill -f "python3 app.py"
nohup python3 app.py > app.log 2>&1 &
echo "服务已重启，访问 http://<服务器IP>:5002"
```

### 查看运行日志

```bash
# 实时查看日志
tail -f ~/amazon_toolkit/app.log

# 查看服务是否在运行
ps aux | grep "python3 app.py"
```

### 停止服务

```bash
pkill -f "python3 app.py"
```
