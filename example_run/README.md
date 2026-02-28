# docxtplrs 完整运行示例

这个目录包含一个完整的运行示例，包括：

- `template.docx` - Word模板文件（由 create_template.py 生成）
- `data.json` - JSON数据文件
- `run.py` - 运行脚本
- `create_template.py` - 模板生成脚本

## 快速开始

### 1. 确保 docxtplrs 已安装

```bash
cd /home/zhangyy/docxtplrs
source .venv/bin/activate
maturin develop --uv
```

### 2. 创建模板（如果还没有）

```bash
cd /home/zhangyy/docxtplrs/example_run
python3 create_template.py
```

### 3. 运行示例

```bash
python3 run.py
```

或者使用 uv：

```bash
uv run run.py
```

## 文件说明

### data.json

包含模板变量和数据的JSON文件：

```json
{
    "title": "文档标题",
    "date": "2024-03-15",
    "company": "公司名称",
    "contact_name": "联系人",
    "items": [
        {"name": "项目1", "quantity": 1, "price": "100", "subtotal": "100"}
    ],
    ...
}
```

### template.docx

Word模板文件，使用 Jinja2 模板语法：

- `{{ variable }}` - 变量替换
- `{% for item in items %}...{% endfor %}` - 循环

### run.py

主运行脚本，执行以下步骤：

1. 加载 JSON 数据
2. 加载 Word 模板
3. 检查模板变量
4. 渲染模板
5. 保存输出文件
6. 验证结果

## 输出

运行后会生成 `output.docx` 文件，可以在 Microsoft Word、WPS 或 LibreOffice 中打开查看。
