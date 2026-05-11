#!/usr/bin/env python3
"""
扩充第七章校核内容
"""

import zipfile
import os
import shutil
from lxml import etree

ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_paragraph_text(p):
    texts = p.findall('.//w:t', ns)
    return ''.join([t.text or '' for t in texts])

def get_paragraph_style(p):
    pStyle = p.find('.//w:pStyle', ns)
    return pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if pStyle is not None else ''

def create_body_paragraph(text, indent='480'):
    """创建正文段落"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    # Paragraph properties
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')

    # Alignment - justified
    jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'both')

    # Spacing - 1.5x line spacing
    spacing = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', '360')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'auto')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before', '0')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after', '0')

    # Indent
    ind = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')
    ind.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine', indent)

    # Run
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

    # Font
    rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', '宋体')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', '宋体')

    # Size - 小四 (24 half-points)
    sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '24')
    szCs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
    szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '24')

    # Text
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

    return p

def expand_chapter7(docx_path, output_path):
    """扩充第七章内容"""
    tree = etree.parse(docx_path)
    root = tree.getroot()
    body = root.find('.//w:body', ns)
    all_children = list(body)

    # Find chapter 7 boundaries
    ch7_title_idx = None
    conclusion_idx = None

    for i, child in enumerate(all_children):
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)

            if style == '1' and '第七章' in text and '压力机' in text:
                ch7_title_idx = i
            elif ch7_title_idx and style == '1' and '结论' in text:
                conclusion_idx = i
                break

    if ch7_title_idx is None or conclusion_idx is None:
        print("Could not find chapter 7 boundaries")
        return False

    print(f"Chapter 7: indices {ch7_title_idx} to {conclusion_idx}")

    # Find paragraphs to replace with expanded content
    # We need to find and expand the verification sections

    # Build new content for verification section
    new_verification_content = [
        # （1）公称压力校核
        '（1）公称压力校核',
        '压力机的公称压力是衡量其冲压能力的核心参数，直接决定了设备能否完成预定的冲压工序。根据第三章的计算结果，本设计总冲压力F总=71.4kN，包括冲裁力63.7kN、卸料力3.2kN、推件力4.0kN以及弯曲力0.92kN。按照冲压模具设计规范，所选压力机的公称压力应大于总冲压力的1.3倍，即:',
        'F公称 ≥ 1.3×F总 = 1.3×71.4 = 92.8kN',
        '所选J23-160型开式可倾压力机的公称压力为160kN，远大于92.8kN的要求值。公称压力裕量为160-92.8=67.2kN，安全系数实际达到160/71.4≈2.24，大于规范要求的1.3倍。这一裕量可以保证在冲压过程中，即使遇到材料性能波动或润滑条件变化等不利因素，压力机仍能稳定工作，不会出现过载现象。因此，公称压力校核满足要求。',

        # （2）闭合高度校核
        '（2）闭合高度校核',
        '模具闭合高度是指模具处于最低工作位置时，上模座上表面与下模座下表面之间的距离。这一参数必须与压力机的闭合高度范围相匹配，否则模具无法正确安装或无法完成冲压行程。',
        '通过CAD软件测量，本设计模具闭合高度H闭=214.2mm。J23-160型压力机的最大闭合高度为220mm，封闭高度调节量为45mm，因此压力机可调节的闭合高度范围为175～220mm。',
        '模具闭合高度214.2mm处于175～220mm范围内，且距离最大闭合高度220mm还有5.8mm的调节余量。这意味着在实际生产中，可以通过调节压力机的封闭高度来补偿模具磨损或调整冲压深度，保证产品质量的稳定性。因此，闭合高度校核满足要求。',

        # （3）工作台尺寸校核
        '（3）工作台尺寸校核',
        '压力机工作台的平面尺寸必须大于模具的外形尺寸，以确保模具能够正确安装在工作台上，并留有足够的空间进行送料、出料等操作。',
        '本设计冲模外形尺寸为405mm×285mm，J23-160型压力机工作台尺寸为450mm×300mm。工作台在长度方向比模具大45mm，宽度方向比模具大15mm。这些空间余量可以满足模具安装定位、紧固螺栓布置以及条料送进的操作空间要求。因此，工作台尺寸校核满足要求。',

        # （4）滑块行程校核
        '（4）滑块行程校核',
        '滑块行程是指压力机滑块从上止点到下止点的移动距离。对于冲裁工序，滑块行程应大于工件高度与凸模进入凹模深度之和，以保证冲压完成后工件能够顺利脱模。',
        '本设计工件为平板类冲压件，弯曲高度较小。J23-160型压力机滑块行程为160mm，远大于工件所需的冲压行程。较大的行程还有利于操作者观察模具工作状态和进行手工送料，提高操作安全性。因此，滑块行程校核满足要求。',

        # （5）模柄孔尺寸校核
        '（5）模柄孔尺寸校核',
        '模柄是连接上模与压力机滑块的重要零件，模柄直径和长度必须与压力机滑块上的模柄孔相匹配，否则模具无法正确安装或在冲压过程中产生偏移。',
        '本设计选用的模柄规格为Φ40×100mm，J23-160型压力机的模柄孔尺寸为Φ40×60mm。模柄直径与模柄孔直径一致，均为Φ40mm，可以实现紧密配合。模柄长度100mm大于模柄孔深度60mm，模柄可以完全插入模柄孔并通过压板紧固，保证连接的可靠性。因此，模柄孔尺寸校核满足要求。',

        # 综合结论
        '综合以上五项校核结果，J23-160型开式可倾压力机的各项参数均满足本设计模具的使用要求，设备选型合理可行。',
    ]

    # Find the start of verification section (after "选定压力机后...")
    verify_start_idx = None
    verify_end_idx = None

    for i in range(ch7_title_idx, conclusion_idx):
        child = all_children[i]
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)

            if style == '2' and '7.2 压力机校核' in text:
                verify_start_idx = i
            elif verify_start_idx and style == '1':
                verify_end_idx = i
                break

    if verify_start_idx is None:
        print("Could not find verification section")
        return False

    if verify_end_idx is None:
        verify_end_idx = conclusion_idx

    print(f"Verification section: {verify_start_idx} to {verify_end_idx}")

    # Remove old verification content (keep the heading)
    elements_to_remove = []
    for i in range(verify_start_idx + 1, verify_end_idx):
        if all_children[i].tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            elements_to_remove.append(all_children[i])
        elif all_children[i].tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl':
            # Keep the image and caption
            pass

    for elem in elements_to_remove:
        text = get_paragraph_text(elem)
        if '图' not in text and '闭合高度示意图' not in text:
            body.remove(elem)

    # Find the position to insert new content (after 7.2 heading)
    all_children = list(body)
    insert_after_idx = None
    for i, child in enumerate(all_children):
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)
            if style == '2' and '7.2 压力机校核' in text:
                insert_after_idx = i
                break

    if insert_after_idx is None:
        print("Could not find insertion point")
        return False

    # Insert new content
    for i, text in enumerate(new_verification_content):
        new_para = create_body_paragraph(text)
        body.insert(insert_after_idx + 1 + i, new_para)

    # Save
    tree.write(output_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    print(f"Successfully expanded verification section with {len(new_verification_content)} paragraphs")
    return True

def main():
    input_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-最终版-v2.docx'
    output_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-最终版-v3.docx'

    temp_dir = '/tmp/docx_temp_expand'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    with zipfile.ZipFile(input_file, 'r') as z:
        z.extractall(temp_dir)

    document_xml = os.path.join(temp_dir, 'word', 'document.xml')
    fixed_xml = '/tmp/document_fixed_expand.xml'

    print("Expanding verification content...")
    if expand_chapter7(document_xml, fixed_xml):
        shutil.copy2(fixed_xml, document_xml)

        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    z.write(file_path, arcname)

        print(f"\nExpanded file saved to: {output_file}")
    else:
        print("Failed to expand content")

    shutil.rmtree(temp_dir)

if __name__ == '__main__':
    main()
