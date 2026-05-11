#!/usr/bin/env python3
"""
修复第七章书签 - 确保位置正确
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

def fix_bookmarks(docx_path, output_path):
    """修复书签"""
    # Parse the document
    tree = etree.parse(docx_path)
    root = tree.getroot()
    body = root.find('.//w:body', ns)

    # Find all direct children
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

    # Find and fix bookmarks for 7.1 and 7.2
    bookmarks_to_fix = {
        '7.1 压力机选择': 'TOC_LINK_050',
        '7.2 压力机校核': 'TOC_LINK_051',
    }

    # Find existing bookmarks to get the max ID
    max_id = 0
    for child in all_children:
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            for bookmarkStart in child.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart'):
                bid = bookmarkStart.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                if bid and bid.isdigit():
                    max_id = max(max_id, int(bid))

    print(f"Max bookmark ID: {max_id}")

    fixed_bookmarks = []

    for i in range(ch7_title_idx, conclusion_idx):
        child = all_children[i]
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            style = get_paragraph_style(child)

            # Check if this is a heading we need to fix
            for heading_text, bookmark_name in bookmarks_to_fix.items():
                if heading_text in text and style == '2':
                    # Remove existing bookmark if any
                    for old_bookmark in child.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart'):
                        old_name = old_bookmark.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name')
                        if old_name == bookmark_name:
                            # Also remove the corresponding bookmarkEnd
                            old_id = old_bookmark.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
                            child.remove(old_bookmark)
                            for end in child.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkEnd'):
                                if end.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id') == old_id:
                                    child.remove(end)
                                    break

                    # Add new bookmark in correct position (after pPr, before first run)
                    max_id += 1
                    new_id = str(max_id)

                    # Create bookmarkStart
                    bookmarkStart = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart')
                    bookmarkStart.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', new_id)
                    bookmarkStart.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name', bookmark_name)

                    # Create bookmarkEnd
                    bookmarkEnd = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkEnd')
                    bookmarkEnd.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', new_id)

                    # Find pPr element
                    pPr = child.find('w:pPr', ns)
                    if pPr is not None:
                        # Insert bookmarkStart after pPr
                        pPr_index = list(child).index(pPr)
                        child.insert(pPr_index + 1, bookmarkStart)
                    else:
                        # Insert at beginning
                        child.insert(0, bookmarkStart)

                    # Find first run and insert bookmarkEnd after it
                    first_run = child.find('w:r', ns)
                    if first_run is not None:
                        run_index = list(child).index(first_run)
                        child.insert(run_index + 1, bookmarkEnd)
                    else:
                        child.append(bookmarkEnd)

                    fixed_bookmarks.append(bookmark_name)
                    print(f"Fixed bookmark {bookmark_name} (id={new_id}) on: {text}")

    # Save the modified document
    tree.write(output_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    print(f"\nFixed {len(fixed_bookmarks)} bookmarks: {fixed_bookmarks}")
    return True

def main():
    # File paths
    input_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-最终版-目录修复.docx'
    output_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-最终版-v2.docx'

    # Extract docx
    temp_dir = '/tmp/docx_temp_bookmarks_v2'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    with zipfile.ZipFile(input_file, 'r') as z:
        z.extractall(temp_dir)

    # Fix the document
    document_xml = os.path.join(temp_dir, 'word', 'document.xml')
    fixed_xml = '/tmp/document_fixed_bookmarks_v2.xml'

    print("Fixing bookmarks...")
    if fix_bookmarks(document_xml, fixed_xml):
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
    else:
        print("Failed to fix bookmarks")

    # Cleanup
    shutil.rmtree(temp_dir)

if __name__ == '__main__':
    main()
