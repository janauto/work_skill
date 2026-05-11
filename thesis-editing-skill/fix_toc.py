#!/usr/bin/env python3
"""
修复目录索引
"""

import zipfile
import os
import shutil
from lxml import etree

# Namespaces
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_paragraph_text(p):
    """获取段落文本"""
    texts = p.findall('.//w:t', ns)
    return ''.join([t.text or '' for t in texts])

def get_paragraph_style(p):
    """获取段落样式"""
    pStyle = p.find('.//w:pStyle', ns)
    return pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if pStyle is not None else ''

def create_toc_entry(text, style, bookmark_name, page_num=''):
    """创建目录条目"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    # Paragraph properties
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')

    # Style
    pStyle = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
    pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', style)

    # Tabs (for page number alignment)
    tabs = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tabs')
    tab = etree.SubElement(tabs, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab')
    tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'right')
    tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}leader', 'dot')
    tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pos', '9000')

    # Spacing
    spacing = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', '360')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'auto')

    # Hyperlink
    hyperlink = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')
    hyperlink.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor', bookmark_name)
    hyperlink.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}history', '1')

    # Run with text
    r = etree.SubElement(hyperlink, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

    sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')
    szCs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
    szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')

    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text

    # Tab run
    if page_num:
        r_tab = etree.SubElement(hyperlink, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        rPr_tab = etree.SubElement(r_tab, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        sz_tab = etree.SubElement(rPr_tab, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
        sz_tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')
        szCs_tab = etree.SubElement(rPr_tab, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
        szCs_tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')

        tab_run = etree.SubElement(r_tab, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab')

        r_page = etree.SubElement(hyperlink, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        rPr_page = etree.SubElement(r_page, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        sz_page = etree.SubElement(rPr_page, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
        sz_page.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')
        szCs_page = etree.SubElement(rPr_page, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
        szCs_page.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')

        t_page = etree.SubElement(r_page, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        t_page.text = page_num

    return p

def fix_toc(docx_path, output_path):
    """修复目录"""
    # Parse the document
    tree = etree.parse(docx_path)
    root = tree.getroot()
    body = root.find('.//w:body', ns)

    # Find all direct children
    all_children = list(body)

    # Find TOC boundaries
    toc_start = None
    toc_end = None

    for i, child in enumerate(all_children):
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)

            if style == 'TOC1' and '摘 要' in text:
                toc_start = i
            elif style == '1' and '第一章' in text:
                toc_end = i
                break

    if toc_start is None or toc_end is None:
        print("Could not find TOC boundaries")
        return False

    print(f"TOC found: indices {toc_start} to {toc_end}")

    # Collect old TOC entries and their bookmarks
    old_toc_entries = []
    for i in range(toc_start, toc_end):
        child = all_children[i]
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)

            if 'TOC' in style:
                # Get bookmark from hyperlink
                hyperlink = child.find('.//w:hyperlink', ns)
                bookmark = None
                if hyperlink is not None:
                    bookmark = hyperlink.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor')

                old_toc_entries.append({
                    'text': text,
                    'style': style,
                    'bookmark': bookmark,
                    'index': i
                })

    print(f"Found {len(old_toc_entries)} TOC entries")

    # Remove old TOC entries
    for i in range(toc_end - 1, toc_start - 1, -1):
        body.remove(all_children[i])

    # Create new TOC entries
    new_toc_entries = []

    # 摘要
    new_toc_entries.append(create_toc_entry('摘 要', 'TOC1', 'TOC_LINK_001', 'I'))

    # ABSTRACT
    new_toc_entries.append(create_toc_entry('ABSTRACT', 'TOC1', 'TOC_LINK_002', 'II'))

    # 第一章 绪论
    new_toc_entries.append(create_toc_entry('第一章 绪论', 'TOC1', 'TOC_LINK_003', '1'))
    new_toc_entries.append(create_toc_entry('1.1 研究背景', 'TOC2', 'TOC_LINK_004', '1'))
    new_toc_entries.append(create_toc_entry('1.2 国内外研究现状', 'TOC2', 'TOC_LINK_005', '1'))
    new_toc_entries.append(create_toc_entry('1.2.1 国内研究现状', 'TOC3', 'TOC_LINK_006', '1'))
    new_toc_entries.append(create_toc_entry('1.2.2 国外研究现状', 'TOC3', 'TOC_LINK_007', '2'))
    new_toc_entries.append(create_toc_entry('1.3 冲压特点与应用', 'TOC2', 'TOC_LINK_008', '2'))
    new_toc_entries.append(create_toc_entry('1.4 研究意义', 'TOC2', 'TOC_LINK_009', '3'))

    # 第二章
    new_toc_entries.append(create_toc_entry('第二章 零件工艺性分析、冲压方案与排样设计', 'TOC1', 'TOC_LINK_010', '4'))
    new_toc_entries.append(create_toc_entry('2.1 工件材料分析', 'TOC2', 'TOC_LINK_011', '4'))
    new_toc_entries.append(create_toc_entry('2.2 工件结构形状分析', 'TOC2', 'TOC_LINK_012', '5'))
    new_toc_entries.append(create_toc_entry('2.3 冲压工艺方案拟定', 'TOC2', 'TOC_LINK_013', '6'))
    new_toc_entries.append(create_toc_entry('2.3.1 工件展开', 'TOC3', 'TOC_LINK_014', '6'))
    new_toc_entries.append(create_toc_entry('2.3.2 冲裁工艺方法选择', 'TOC3', 'TOC_LINK_015', '6'))
    new_toc_entries.append(create_toc_entry('2.4 排样设计及材料利用率计算', 'TOC2', 'TOC_LINK_016', '7'))
    new_toc_entries.append(create_toc_entry('2.4.1 排样方式选择', 'TOC3', 'TOC_LINK_017', '8'))
    new_toc_entries.append(create_toc_entry('2.4.2 搭边值确定', 'TOC3', 'TOC_LINK_018', '9'))
    new_toc_entries.append(create_toc_entry('2.4.3 材料利用率计算', 'TOC3', 'TOC_LINK_019', '10'))

    # 第三章
    new_toc_entries.append(create_toc_entry('第三章 冲压力与压力中心计算', 'TOC1', 'TOC_LINK_020', '12'))
    new_toc_entries.append(create_toc_entry('3.1 冲压力计算', 'TOC2', 'TOC_LINK_021', '12'))
    new_toc_entries.append(create_toc_entry('3.2 初选压力机', 'TOC2', 'TOC_LINK_022', '14'))
    new_toc_entries.append(create_toc_entry('3.3 压力中心计算', 'TOC2', 'TOC_LINK_023', '14'))

    # 第四章
    new_toc_entries.append(create_toc_entry('第四章 模具刃口尺寸及弯曲工作部分计算', 'TOC1', 'TOC_LINK_024', '16'))
    new_toc_entries.append(create_toc_entry('4.1 冲裁间隙的确定', 'TOC2', 'TOC_LINK_025', '16'))
    new_toc_entries.append(create_toc_entry('4.2 弯曲工作部分尺寸计算', 'TOC2', 'TOC_LINK_026', '20'))
    new_toc_entries.append(create_toc_entry('4.2.1 弯曲凸、凹模间隙计算', 'TOC3', 'TOC_LINK_027', '20'))
    new_toc_entries.append(create_toc_entry('4.2.2 弯曲凸、凹模圆角半径的确定', 'TOC3', 'TOC_LINK_028', '20'))
    new_toc_entries.append(create_toc_entry('4.2.3 弯曲凹模工作部分深度', 'TOC3', 'TOC_LINK_029', '20'))

    # 第五章
    new_toc_entries.append(create_toc_entry('第五章 模具主要工作零件结构设计', 'TOC1', 'TOC_LINK_030', '21'))
    new_toc_entries.append(create_toc_entry('5.1 凹模设计', 'TOC2', 'TOC_LINK_031', '21'))
    new_toc_entries.append(create_toc_entry('5.1.1 凹模刃口结构形式的选择', 'TOC3', 'TOC_LINK_032', '21'))
    new_toc_entries.append(create_toc_entry('5.1.2 凹模精度与材料的确定', 'TOC3', 'TOC_LINK_033', '22'))
    new_toc_entries.append(create_toc_entry('5.1.3 凹模外形尺寸的确定', 'TOC3', 'TOC_LINK_034', '22'))
    new_toc_entries.append(create_toc_entry('5.2 凸模设计', 'TOC2', 'TOC_LINK_035', '23'))
    new_toc_entries.append(create_toc_entry('5.2.1 凸模结构的确定', 'TOC3', 'TOC_LINK_036', '23'))
    new_toc_entries.append(create_toc_entry('5.2.2 凸模高度的确定', 'TOC3', 'TOC_LINK_037', '23'))
    new_toc_entries.append(create_toc_entry('5.2.3 凸模材料的确定', 'TOC3', 'TOC_LINK_038', '24'))
    new_toc_entries.append(create_toc_entry('5.3 卸料装置设计', 'TOC2', 'TOC_LINK_039', '24'))
    new_toc_entries.append(create_toc_entry('5.3.1 卸料板外形设计', 'TOC3', 'TOC_LINK_040', '24'))
    new_toc_entries.append(create_toc_entry('5.3.2 卸料板材料选择', 'TOC3', 'TOC_LINK_041', '25'))
    new_toc_entries.append(create_toc_entry('5.3.3 弹性元件设计', 'TOC3', 'TOC_LINK_042', '25'))
    new_toc_entries.append(create_toc_entry('5.4 固定板和垫板设计', 'TOC2', 'TOC_LINK_043', '25'))
    new_toc_entries.append(create_toc_entry('5.5 定位零件设计', 'TOC2', 'TOC_LINK_044', '26'))

    # 第六章
    new_toc_entries.append(create_toc_entry('第六章 模具其他零件与总体结构设计', 'TOC1', 'TOC_LINK_045', '27'))
    new_toc_entries.append(create_toc_entry('6.1 模柄选择', 'TOC2', 'TOC_LINK_046', '27'))
    new_toc_entries.append(create_toc_entry('6.2 模架选择', 'TOC2', 'TOC_LINK_047', '27'))
    new_toc_entries.append(create_toc_entry('6.3 其他零件选择', 'TOC2', 'TOC_LINK_048', '29'))

    # 第七章 - 新结构
    new_toc_entries.append(create_toc_entry('第七章 压力机选择与校核', 'TOC1', 'TOC_LINK_049', '30'))
    new_toc_entries.append(create_toc_entry('7.1 压力机选择', 'TOC2', 'TOC_LINK_050', '30'))
    new_toc_entries.append(create_toc_entry('7.2 压力机校核', 'TOC2', 'TOC_LINK_051', '31'))

    # 结论
    new_toc_entries.append(create_toc_entry('结论', 'TOC1', 'TOC_LINK_052', '32'))

    # 参考文献
    new_toc_entries.append(create_toc_entry('参考文献', 'TOC1', 'TOC_LINK_053', '33'))

    # 致谢
    new_toc_entries.append(create_toc_entry('致 谢', 'TOC1', 'TOC_LINK_054', '35'))

    # Insert new TOC entries at the original position
    for i, entry in enumerate(new_toc_entries):
        body.insert(toc_start + i, entry)

    # Now we need to add bookmarks to the body paragraphs for the TOC links
    # Find chapter headings and add bookmarks
    all_children = list(body)
    bookmark_targets = {
        'TOC_LINK_003': '第一章',
        'TOC_LINK_010': '第二章',
        'TOC_LINK_020': '第三章',
        'TOC_LINK_024': '第四章',
        'TOC_LINK_030': '第五章',
        'TOC_LINK_045': '第六章',
        'TOC_LINK_049': '第七章',
        'TOC_LINK_052': '结论',
        'TOC_LINK_053': '参考文献',
        'TOC_LINK_054': '致 谢',
    }

    # Add bookmarks to body paragraphs
    for child in all_children:
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)

            # Check if this is a heading that needs a bookmark
            for bookmark, keyword in bookmark_targets.items():
                if keyword in text and style == '1':
                    # Check if bookmark already exists
                    existing_bookmark = child.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart')
                    if existing_bookmark is None:
                        # Add bookmark
                        pPr = child.find('w:pPr', ns)
                        if pPr is None:
                            pPr = etree.SubElement(child, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')

                        # Insert bookmark before pPr's next sibling
                        bookmarkStart = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart')
                        bookmarkStart.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', bookmark.replace('TOC_LINK_', ''))
                        bookmarkStart.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name', bookmark)

                        bookmarkEnd = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkEnd')
                        bookmarkEnd.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', bookmark.replace('TOC_LINK_', ''))

                        # Insert at the beginning of the paragraph
                        child.insert(0, bookmarkStart)
                        # Insert after first run
                        runs = child.findall('w:r', ns)
                        if runs:
                            child.insert(list(child).index(runs[0]) + 1, bookmarkEnd)
                        else:
                            child.append(bookmarkEnd)

                    break

    # Save the modified document
    tree.write(output_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    print(f"Successfully updated TOC with {len(new_toc_entries)} entries")
    return True

def main():
    # File paths
    input_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-第七章修改版.docx'
    output_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-第七章修改版-目录修正.docx'

    # Extract docx
    temp_dir = '/tmp/docx_temp_toc'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    with zipfile.ZipFile(input_file, 'r') as z:
        z.extractall(temp_dir)

    # Fix the document
    document_xml = os.path.join(temp_dir, 'word', 'document.xml')
    fixed_xml = '/tmp/document_fixed_toc.xml'

    print("Fixing TOC...")
    if fix_toc(document_xml, fixed_xml):
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
        print("\nTOC has been updated:")
        print("- 7.1 压力机选择 (replaced 7.1 冲压设备校核)")
        print("- 7.2 压力机校核 (replaced 7.2 压力机选择)")
    else:
        print("Failed to fix TOC")

    # Cleanup
    shutil.rmtree(temp_dir)

if __name__ == '__main__':
    main()
