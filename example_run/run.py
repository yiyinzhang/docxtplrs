#!/usr/bin/env python3
"""
docxtplrs 完整运行示例

使用方式:
    1. 确保已安装 docxtplrs: maturin develop --uv
    2. 创建模板: python3 create_template.py
    3. 运行: python3 run.py

文件说明:
    - template.docx    Word模板文件（包含Jinja2语法）
    - data.json        JSON数据文件
    - output.docx      生成的输出文件
"""

import json
import os
import sys
import zipfile
from pathlib import Path

# 添加父目录到路径以导入docxtplrs
sys.path.insert(0, str(Path(__file__).parent.parent))

from docxtplrs import DocxTemplate, RichText, version


def load_json(filepath: str) -> dict:
    """加载JSON文件"""
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def create_sample_template():
    """如果模板不存在，创建示例模板"""
    if os.path.exists("template.docx"):
        return
    
    print("创建示例模板...")
    import subprocess
    subprocess.run([sys.executable, "create_template.py"], check=True)


def preview_docx_content(docx_path: str) -> str:
    """预览docx文件的文本内容"""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            content = zf.read("word/document.xml").decode('utf-8')
            # 提取文本内容（简单处理）
            import re
            texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', content)
            return '\n'.join(t for t in texts if t.strip())
    except Exception as e:
        return f"无法读取: {e}"


def main():
    print("=" * 70)
    print("docxtplrs 运行示例")
    print(f"版本: {version()}")
    print("=" * 70)
    
    # 文件路径
    template_path = "template.docx"
    data_path = "data.json"
    output_path = "output.docx"
    
    # 检查文件
    if not os.path.exists(template_path):
        print(f"\n[!] 模板文件不存在: {template_path}")
        print("    正在创建示例模板...")
        create_sample_template()
    
    if not os.path.exists(data_path):
        print(f"\n[!] 数据文件不存在: {data_path}")
        print("    请确保 data.json 文件存在")
        return 1
    
    try:
        # 1. 加载数据
        print(f"\n[1/5] 加载JSON数据: {data_path}")
        context = load_json(data_path)
        print(f"      ✓ 已加载 {len(context)} 个变量")
        
        # 2. 加载模板
        print(f"\n[2/5] 加载Word模板: {template_path}")
        doc = DocxTemplate(template_path)
        print("      ✓ 模板加载成功")
        
        # 显示模板需要的变量
        vars_needed = doc.get_undeclared_template_variables()
        print(f"\n      模板需要的变量:")
        for var in sorted(vars_needed):
            status = "✓" if var in context else "✗"
            print(f"        {status} {var}")
        
        # 3. 渲染
        print(f"\n[3/5] 渲染模板...")
        doc.render(context)
        print("      ✓ 渲染完成")
        
        # 4. 保存
        print(f"\n[4/5] 保存结果: {output_path}")
        doc.save(output_path)
        print(f"      ✓ 文件已保存 ({os.path.getsize(output_path)} 字节)")
        
        # 5. 验证输出
        print(f"\n[5/5] 验证输出文件...")
        with zipfile.ZipFile(output_path, 'r') as zf:
            content = zf.read("word/document.xml").decode('utf-8')
            
            # 检查关键内容是否被替换
            checks = [
                (context.get('company', '') in content, "公司名称"),
                (context.get('contact_name', '') in content, "联系人"),
                (context.get('total', '') in content, "总计金额"),
            ]
            
            for passed, desc in checks:
                status = "✓" if passed else "✗"
                print(f"      {status} {desc}")
        
        # 显示预览
        print(f"\n[预览] 输出文件内容片段:")
        print("-" * 70)
        preview = preview_docx_content(output_path)
        # 只显示前500字符
        if len(preview) > 500:
            preview = preview[:500] + "..."
        print(preview)
        print("-" * 70)
        
        print(f"\n" + "=" * 70)
        print("✓ 成功生成文档!")
        print(f"  输出文件: {os.path.abspath(output_path)}")
        print("=" * 70)
        
        return 0
        
    except FileNotFoundError as e:
        print(f"\n✗ 错误: 文件未找到 - {e}")
        return 1
    except json.JSONDecodeError as e:
        print(f"\n✗ 错误: JSON格式错误 - {e}")
        return 1
    except Exception as e:
        print(f"\n✗ 错误: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    exit(main())
