#!/usr/bin/env python3
"""
docxtplrs 演示示例 - 简化版

使用方式:
    uv run examples/demo.py
"""

import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path

# 导入docxtplrs
sys.path.insert(0, str(Path(__file__).parent.parent))
from docxtplrs import DocxTemplate, version


def create_demo_template(path: str):
    """创建演示用的docx模板"""
    with zipfile.ZipFile(path, "w") as zf:
        zf.writestr(
            "[Content_Types].xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''',
        )

        zf.writestr(
            "_rels/.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''',
        )

        zf.writestr(
            "word/_rels/document.xml.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''',
        )

        # 简化的文档内容
        zf.writestr(
            "word/document.xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:b/><w:sz w:val="48"/></w:rPr><w:t>{{ title }}</w:t></w:r>
        </w:p>
        <w:p><w:r><w:t>日期: {{ date }}</w:t></w:r></w:p>
        <w:p><w:r><w:t>公司: {{ company }}</w:t></w:r></w:p>
        <w:p><w:r><w:t>联系人: {{ contact_name }}</w:t></w:r></w:p>
        <w:p><w:r><w:t></w:t></w:r></w:p>
        <w:p><w:r><w:t>{{ greeting }}</w:t></w:r></w:p>
        <w:p><w:r><w:t></w:t></w:r></w:p>
        <w:p><w:r><w:t>项目列表:</w:t></w:r></w:p>
        <w:p><w:r><w:t>{% for item in items %}{{ item.name }}: ¥{{ item.price }} x {{ item.quantity }} = ¥{{ item.subtotal }}
{% endfor %}</w:t></w:r></w:p>
        <w:p><w:r><w:t></w:t></w:r></w:p>
        <w:p><w:r><w:t>小计: ¥{{ subtotal }}</w:t></w:r></w:p>
        <w:p><w:r><w:t>税率: {{ tax_rate }}%</w:t></w:r></w:p>
        <w:p><w:r><w:rPr><w:b/></w:rPr><w:t>总计: ¥{{ total }}</w:t></w:r></w:p>
        <w:p><w:r><w:t></w:t></w:r></w:p>
        <w:p><w:r><w:t>备注: {{ notes }}</w:t></w:r></w:p>
        <w:p><w:r><w:t></w:t></w:r></w:p>
        <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{{ author }} | {{ author_email }}</w:t></w:r></w:p>
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>''',
        )


def main():
    print("=" * 60)
    print("docxtplrs 演示")
    print(f"版本: {version()}")
    print("=" * 60)
    
    # 准备数据
    context = {
        "title": "项目报价单",
        "date": "2024-03-15",
        "company": "科技创新有限公司",
        "contact_name": "张经理",
        "greeting": "感谢您的咨询，以下是我们的报价：",
        "items": [
            {"name": "软件开发", "quantity": 1, "price": "50000", "subtotal": "50000"},
            {"name": "系统部署", "quantity": 2, "price": "8000", "subtotal": "16000"},
            {"name": "年度维护", "quantity": 1, "price": "12000", "subtotal": "12000"},
        ],
        "subtotal": "78000",
        "tax_rate": 6,
        "tax_amount": "4680",
        "total": "82680",
        "notes": "付款方式：合同签订后支付50%，验收后支付50%",
        "author": "销售部",
        "author_email": "sales@example.com"
    }
    
    # 创建临时文件
    with tempfile.TemporaryDirectory() as tmpdir:
        template_path = os.path.join(tmpdir, "template.docx")
        output_path = os.path.join(tmpdir, "output.docx")
        
        # 创建模板
        print("\n[1/4] 创建模板...")
        create_demo_template(template_path)
        print("      ✓ 模板创建成功")
        
        # 加载模板
        print("\n[2/4] 加载模板...")
        doc = DocxTemplate(template_path)
        print("      ✓ 模板加载成功")
        
        # 渲染
        print("\n[3/4] 渲染文档...")
        doc.render(context)
        print("      ✓ 渲染完成")
        
        # 保存
        print("\n[4/4] 保存文档...")
        doc.save(output_path)
        print(f"      ✓ 已保存: {output_path}")
        
        # 验证输出
        print("\n[验证] 检查输出内容...")
        with zipfile.ZipFile(output_path, 'r') as zf:
            content = zf.read("word/document.xml").decode('utf-8')
            
            checks = [
                ("科技创新有限公司" in content, "公司名称"),
                ("张经理" in content, "联系人"),
                ("82680" in content, "总计金额"),
            ]
            
            for passed, desc in checks:
                status = "✓" if passed else "✗"
                print(f"      {status} {desc}")
        
        # 复制到当前目录供查看
        final_output = "demo_output.docx"
        import shutil
        shutil.copy(output_path, final_output)
        print(f"\n输出文件已复制到: {os.path.abspath(final_output)}")
    
    print("\n" + "=" * 60)
    print("✓ 演示完成!")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    exit(main())
