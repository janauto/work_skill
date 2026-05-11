#!/usr/bin/env python3
"""
修复第七章结构：先需求参数再选型校核 (v2)
"""

import zipfile
import os
import shutil
from lxml import etree

# Namespaces
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def create_paragraph(text, style=None, font='宋体', size='24', bold=False,
                    alignment=None, indent=None, spacing_line='360'):
    """创建段落"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    # Paragraph properties
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')

    # Style
    if style:
        pStyle = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
        pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', style)

    # Alignment
    if alignment:
        jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
        jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', alignment)

    # Spacing
    spacing = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', spacing_line)
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'auto')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before', '0')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after', '0')

    # Indent
    if indent:
        ind = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')
        ind.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine', indent)

    # Run
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

    # Font
    rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', font)
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', font)
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', font)

    # Size
    sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', size)
    szCs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
    szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', size)

    # Bold
    if bold:
        b = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')

    # Text
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

    return p

def create_heading2(text):
    """创建二级标题"""
    return create_paragraph(text, style='2', font='黑体', size='28', bold=True,
                          alignment='left', spacing_line='360')

def create_body_text(text, indent='480'):
    """创建正文段落"""
    return create_paragraph(text, font='宋体', size='24', alignment='both', indent=indent)

def create_table_title(text):
    """创建表题"""
    return create_paragraph(text, font='宋体', size='22', bold=True, alignment='center')

def create_table(headers, rows):
    """创建表格"""
    tbl = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl')

    # Table properties
    tblPr = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr')

    # Table style
    tblStyle = etree.SubElement(tblPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblStyle')
    tblStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'TableGrid')

    # Table width
    tblW = etree.SubElement(tblPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblW')
    tblW.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', '5000')
    tblW.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'pct')

    # Table borders
    tblBorders = etree.SubElement(tblPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = etree.SubElement(tblBorders, f'{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{border_name}')
        border.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'single')
        border.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz', '4')
        border.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space', '0')
        border.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color', '000000')

    # Create header row
    header_row = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
    for header in headers:
        tc = etree.SubElement(header_row, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
        p = etree.SubElement(tc, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

        pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
        jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')

        r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

        rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', '宋体')
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
        rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', '宋体')

        sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
        sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')
        szCs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
        szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')

        b = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')

        t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        t.text = header

    # Create data rows
    for row_data in rows:
        tr = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
        for cell_text in row_data:
            tc = etree.SubElement(tr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
            p = etree.SubElement(tc, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

            r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
            rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

            rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
            rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', '宋体')
            rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
            rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', '宋体')

            sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
            sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')
            szCs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
            szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')

            t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            t.text = cell_text

    return tbl

def get_paragraph_text(p):
    """获取段落文本"""
    texts = p.findall('.//w:t', ns)
    return ''.join([t.text or '' for t in texts])

def get_paragraph_style(p):
    """获取段落样式"""
    pStyle = p.find('.//w:pStyle', ns)
    return pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if pStyle is not None else ''

def fix_chapter7(docx_path, output_path):
    """修复第七章"""
    # Parse the document
    tree = etree.parse(docx_path)
    root = tree.getroot()
    body = root.find('.//w:body', ns)

    # Find all paragraphs
    paragraphs = body.findall('.//w:p', ns)

    # Find chapter 7 boundaries (skip TOC entries)
    ch7_start = None
    ch7_end = None

    for i, p in enumerate(paragraphs):
        text = get_paragraph_text(p)
        style = get_paragraph_style(p)

        # Skip TOC entries
        if 'TOC' in style:
            continue

        # Find chapter 7 heading (style 1 = Heading 1)
        if style == '1' and '第七章' in text and '压力机' in text:
            ch7_start = i
            print(f'Found Chapter 7 at paragraph {i}: {text}')
            continue

        # Find conclusion (style 1 = Heading 1)
        if ch7_start and style == '1' and '结论' in text and len(text) < 10:
            ch7_end = i
            print(f'Found Conclusion at paragraph {i}: {text}')
            break

    if ch7_start is None or ch7_end is None:
        print("Could not find chapter 7 boundaries")
        return False

    print(f"Chapter 7 found: paragraphs {ch7_start} to {ch7_end}")
    print(f"Total paragraphs to remove: {ch7_end - ch7_start - 1}")

    # Remove old chapter 7 content (keep the title, remove everything after until conclusion)
    # We need to collect all elements between ch7_start and ch7_end
    elements_to_remove = []

    # Get all direct children of body
    all_children = list(body)

    # Find the index of ch7_start and ch7_end in all_children
    ch7_title_elem = paragraphs[ch7_start]
    conclusion_elem = paragraphs[ch7_end]

    ch7_title_idx = all_children.index(ch7_title_elem)
    conclusion_idx = all_children.index(conclusion_elem)

    print(f"Chapter 7 title at body index {ch7_title_idx}, conclusion at body index {conclusion_idx}")

    # Remove all elements between title and conclusion
    for i in range(conclusion_idx - 1, ch7_title_idx, -1):
        body.remove(all_children[i])

    # Find insertion point (after chapter 7 title)
    # Re-find the chapter 7 title paragraph in body's direct children
    all_children = list(body)
    insert_after_idx = None
    for i, child in enumerate(all_children):
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)
            if style == '1' and '第七章' in text and '压力机' in text:
                insert_after_idx = i
                break

    if insert_after_idx is None:
        print("Could not find chapter 7 title for insertion")
        return False

    print(f"Inserting new content after body index {insert_after_idx}")

    # Create new chapter 7 content with correct structure
    new_elements = []

    # 7.1 压力机选择
    new_elements.append(create_heading2('7.1 压力机选择'))

    new_elements.append(create_body_text(
        '根据第三章冲压力计算结果，本设计总冲压力F总=71.4kN。考虑到冲压成型加工的实际情况，'
        '并预留一定的安全裕度，需要选择公称压力大于1.3倍总冲压力的压力机。'
    ))

    new_elements.append(create_body_text(
        '计算所需公称压力：1.3×F总=1.3×71.4=92.8kN'
    ))

    new_elements.append(create_body_text(
        '综合考虑压力机的公称压力、闭合高度、工作台尺寸、滑块行程等参数，'
        '决定采用J23-160型开式可倾压力机。该压力机公称压力为160kN，满足安全系数要求。'
        '其相关参数见下表：'
    ))

    # Table 7.1
    new_elements.append(create_table_title('表7.1 压力机主要参数'))

    headers = ['名称', '数值', '名称', '数值']
    rows = [
        ['公称压力', '160kN', '连杆调节长度', '/'],
        ['滑块行程', '160mm', '工作台尺寸', '450mm×300mm'],
        ['行程次数', '55次/min', '封闭高度调节量', '45mm'],
        ['最大闭合高度', '220mm', '模柄孔尺寸', 'Φ40×60'],
    ]
    new_elements.append(create_table(headers, rows))

    # 7.2 压力机校核
    new_elements.append(create_heading2('7.2 压力机校核'))

    new_elements.append(create_body_text(
        '选定压力机后，需要对压力机的各项参数进行校核，确保其满足模具设计要求。'
    ))

    # 校核项目1：公称压力
    new_elements.append(create_body_text(
        '（1）公称压力校核'
    ))

    new_elements.append(create_body_text(
        '压力机的公称压力应大于冲压所需的总压力。本设计总冲压力F总=71.4kN，'
        '所选J23-160型压力机公称压力为160kN。由于160kN＞1.3×71.4kN=92.8kN，'
        '满足公称压力校核要求。'
    ))

    # 校核项目2：闭合高度
    new_elements.append(create_body_text(
        '（2）闭合高度校核'
    ))

    new_elements.append(create_body_text(
        '通过CAD软件测量，模具闭合高度H闭=214.2mm。压力机最大闭合高度为220mm，'
        '封闭高度调节量为45mm，因此压力机闭合高度范围为175～220mm。'
        '模具闭合高度214.2mm在此范围内，满足闭合高度校核要求。'
    ))

    # 校核项目3：工作台尺寸
    new_elements.append(create_body_text(
        '（3）工作台尺寸校核'
    ))

    new_elements.append(create_body_text(
        '冲模外形尺寸为405mm×285mm，压力机工作台尺寸为450mm×300mm。'
        '由于工作台尺寸大于模具外形尺寸，能够满足安装和工作要求。'
    ))

    # 校核项目4：滑块行程
    new_elements.append(create_body_text(
        '（4）滑块行程校核'
    ))

    new_elements.append(create_body_text(
        'J23-160型压力机滑块行程为160mm，大于工件所需的冲压行程，'
        '能够满足冲压工艺要求。'
    ))

    # 校核项目5：模柄孔尺寸
    new_elements.append(create_body_text(
        '（5）模柄孔尺寸校核'
    ))

    new_elements.append(create_body_text(
        '压力机模柄孔尺寸为Φ40×60mm，与所选模柄尺寸相匹配，满足安装要求。'
    ))

    # Insert new elements after chapter 7 title
    for i, elem in enumerate(new_elements):
        body.insert(insert_after_idx + 1 + i, elem)

    # Save the modified document
    tree.write(output_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    print(f"Successfully created new Chapter 7 with {len(new_elements)} elements")
    return True

def main():
    # File paths
    input_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真.docx'
    output_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-第七章修改版.docx'

    # Create backup
    backup_file = input_file + '.ch7_backup_v2'
    if not os.path.exists(backup_file):
        shutil.copy2(input_file, backup_file)
        print(f"Created backup: {backup_file}")

    # Extract docx
    temp_dir = '/tmp/docx_temp_ch7_v2'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    with zipfile.ZipFile(input_file, 'r') as z:
        z.extractall(temp_dir)

    # Fix the document
    document_xml = os.path.join(temp_dir, 'word', 'document.xml')
    fixed_xml = '/tmp/document_fixed_ch7_v2.xml'

    print("Fixing Chapter 7...")
    if fix_chapter7(document_xml, fixed_xml):
        # Replace the document.xml
        shutil.copy2(fixed_xml, document_xml)

        # Create new docx
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    z.write(file_path, arcname)

        print(f"\nFixed file saved to: {output_file}")
        print("\nChapter 7 structure has been reorganized:")
        print("7.1 压力机选择 (先明确需求参数)")
        print("    - 总冲压力计算结果引用")
        print("    - 所需公称压力计算")
        print("    - 压力机选型说明")
        print("    - 表7.1 压力机主要参数")
        print("7.2 压力机校核 (再验证选型)")
        print("    - (1) 公称压力校核")
        print("    - (2) 闭合高度校核")
        print("    - (3) 工作台尺寸校核")
        print("    - (4) 滑块行程校核")
        print("    - (5) 模柄孔尺寸校核")
    else:
        print("Failed to fix Chapter 7")

    # Cleanup
    shutil.rmtree(temp_dir)

if __name__ == '__main__':
    main()
