# 船厂点表清洗向导

基于 Flask + Vue 的交互式工具，用于按照《船厂数据清洗需求.md》要求上传原始点表、选择表头和字段映射，在线补全后导出标准格式的多 Sheet Excel。

## 环境准备
- Python 3.10+
- 建议在虚拟环境中安装依赖

安装依赖：
```bash
python -m venv .venv
source .venv/bin/activate  # Windows 使用 .venv\\Scripts\\activate
pip install --upgrade pip
pip install -e .
```

## 运行步骤
1. 启动服务：
   ```bash
   python app.py
   ```
   默认监听 http://127.0.0.1:5000。

2. 浏览器访问主页后，按照页面向导完成以下操作：
   - **上传 Excel**：选择原始点表文件。
   - **选择 Sheet & 表头**：点击需要处理的 Sheet，选择实际表头所在的行。
   - **配置映射**：为“设备组/名”“检测点名称”等标准字段选择对应列（可使用自动匹配后再微调）。
   - **预览与补全**：检查清洗结果，批量补值或逐行编辑（如 IP、采集类型、业务类型等）。
   - **填写船只信息**：在导出前录入船名、船东、项目号等元数据。
   - **导出标准点表**：点击“导出标准点表”生成包含说明、分组/设备表和分组点表的 Excel。

## 访问与数据说明
- 上传文件会保存到 `uploads/temp.xlsx`，每次新上传会覆盖。
- 导出的标准 Excel 文件命名为 `standard_points.xlsx`，包含：
  - **说明**：船只信息、分组列表、设备列表。
  - **各分组 Sheet**：按设备分块排列的点位，保留数据类型、信号/业务类型、范围/单位等清洗结果。

## 常见问题
- 如果页面未显示 Tailwind 样式，请确保可以访问 `https://cdn.tailwindcss.com`。
- 如需修改监听端口，可在 `app.py` 末尾的 `app.run` 调整 `port` 参数。
