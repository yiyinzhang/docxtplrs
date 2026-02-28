#!/usr/bin/env python3
"""创建示例docx模板文件"""

import zipfile
import os

TEMPLATE_PATH = "template.docx"

def create_template():
    """创建一个示例Word模板，包含Jinja2模板语法"""
    
    with zipfile.ZipFile(TEMPLATE_PATH, "w") as zf:
        # [Content_Types].xml
        zf.writestr(
            "[Content_Types].xml",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>''',
        )

        # _rels/.rels
        zf.writestr(
            "_rels/.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>''',
        )

        # word/_rels/document.xml.rels
        zf.writestr(
            "word/_rels/document.xml.rels",
            '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>''',
        )

        # word/document.xml - 文档内容（简化的XML，避免注释）
        document_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:body>
        <w:p>
            <w:pPr>
                <w:pStyle w:val="Title"/>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                    <w:sz w:val="48"/>
                    <w:szCs w:val="48"/>
                </w:rPr>
                <w:t>{{ title }}</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:pPr>
                <w:jc w:val="right"/>
            </w:pPr>
            <w:r>
                <w:t>日期: {{ date }}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:jc w:val="right"/>
            </w:pPr>
            <w:r>
                <w:t>致: {{ company }}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:r>
                <w:t>联系人: {{ contact_name }} ({{ contact_title }})</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t></w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t>尊敬的 {{ contact_name }}：</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t>{{ greeting }}</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:rPr>
                    <w:b/>
                </w:rPr>
                <w:t>项目详情：</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t>{% for item in items %}[{{ item.name }}] 数量: {{ item.quantity }}, 单价: ¥{{ item.price }}, 小计: ¥{{ item.subtotal }}
{% endfor %}</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t></w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:pPr>
                <w:jc w:val="right"/>
            </w:pPr>
            <w:r>
                <w:t>小计: ¥{{ subtotal }}</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:jc w:val="right"/>
            </w:pPr>
            <w:r>
                <w:t>税率: {{ tax_rate }}%</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:jc w:val="right"/>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:b/>
                </w:rPr>
                <w:t>总计: ¥{{ total }}</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t></w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t>备注: {{ notes }}</w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:r>
                <w:t></w:t>
            </w:r>
        </w:p>
        
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:t>-- 感谢您的合作 --</w:t>
            </w:r>
        </w:p>
        <w:p>
            <w:pPr>
                <w:jc w:val="center"/>
            </w:pPr>
            <w:r>
                <w:t>{{ author }} | {{ author_email }}</w:t>
            </w:r>
        </w:p>
        
        <w:sectPr>
            <w:pgSz w:w="12240" w:h="15840"/>
            <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
        </w:sectPr>
    </w:body>
</w:document>'''
        
        zf.writestr("word/document.xml", document_xml)
    
    print(f"✓ 模板已创建: {TEMPLATE_PATH}")
    return TEMPLATE_PATH

if __name__ == "__main__":
    create_template()
