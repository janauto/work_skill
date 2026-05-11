#!/usr/bin/env python3
"""
在第七章添加图片
"""

import zipfile
import os
import shutil
from lxml import etree

# Namespaces
ns = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}

def get_paragraph_text(p):
    """获取段落文本"""
    texts = p.findall('.//w:t', ns)
    return ''.join([t.text or '' for t in texts])

def get_paragraph_style(p):
    """获取段落样式"""
    pStyle = p.find('.//w:pStyle', ns)
    return pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if pStyle is not None else ''

def create_image_paragraph(image_rId, cx='5269230', cy='3494405'):
    """创建图片段落"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    # Paragraph properties - centered
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')

    # Spacing
    spacing = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before', '0')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after', '0')

    # Run
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')

    # Drawing
    drawing = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')

    # Inline
    inline = etree.SubElement(drawing, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')
    inline.set('distT', '0')
    inline.set('distB', '0')
    inline.set('distL', '0')
    inline.set('distR', '0')

    # Extent
    extent = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
    extent.set('cx', cx)
    extent.set('cy', cy)

    # EffectExtent
    effectExtent = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}effectExtent')
    effectExtent.set('l', '0')
    effectExtent.set('t', '0')
    effectExtent.set('r', '0')
    effectExtent.set('b', '0')

    # DocProperties
    docPr = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr')
    docPr.set('id', '1')
    docPr.set('name', 'Picture 1')

    # Graphic
    graphic = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
    graphicData = etree.SubElement(graphic, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')

    # Picture
    pic = etree.SubElement(graphicData, '{http://schemas.openxmlformats.org/drawingml/2006/picture}pic')

    # NonVisualPictureProperties
    nvPicPr = etree.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}nvPicPr')
    cNvPr = etree.SubElement(nvPicPr, '{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPr')
    cNvPr.set('id', '0')
    cNvPr.set('name', 'image62.png')
    cNvPicPr = etree.SubElement(nvPicPr, '{http://schemas.openxmlformats.org/drawingml/2006/picture}cNvPicPr')

    # BlipFill
    blipFill = etree.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill')
    blip = etree.SubElement(blipFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', image_rId)
    stretch = etree.SubElement(blipFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}stretch')
    fillRect = etree.SubElement(stretch, '{http://schemas.openxmlformats.org/drawingml/2006/main}fillRect')

    # ShapeProperties
    spPr = etree.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}spPr')
    xfrm = etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
    off = etree.SubElement(xfrm, '{http://schemas.openxmlformats.org/drawingml/2006/main}off')
    off.set('x', '0')
    off.set('y', '0')
    ext = etree.SubElement(xfrm, '{http://schemas.openxmlformats.org/drawingml/2006/main}ext')
    ext.set('cx', cx)
    ext.set('cy', cy)
    prstGeom = etree.SubElement(spPr, '{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom')
    prstGeom.set('prst', 'rect')

    return p

def create_caption_paragraph(text):
    """创建图题段落"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    # Paragraph properties - centered
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')

    # Spacing
    spacing = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}before', '0')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}after', '0')

    # Run
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')

    # Font
    rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', '宋体')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hAnsi', '宋体')

    # Size - 五号 (22 half-points)
    sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')
    szCs = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}szCs')
    szCs.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')

    # Bold
    b = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')

    # Text
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = text

    return p

def add_image_to_chapter7(docx_path, output_path):
    """在第七章添加图片"""
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

    # Find the position to insert image (before 闭合高度校核)
    insert_idx = None
    for i in range(ch7_title_idx, conclusion_idx):
        child = all_children[i]
        if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p':
            text = get_paragraph_text(child)
            if '（2）闭合高度校核' in text:
                # Insert before this paragraph
                insert_idx = i
                break

    if insert_idx is None:
        print("Could not find insertion point")
        return False

    print(f"Inserting image at index {insert_idx}")

    # Create image and caption paragraphs
    # Note: We need to add the relationship for the image
    image_para = create_image_paragraph('rId102')
    caption_para = create_caption_paragraph('图7.1 模具闭合高度示意图')

    # Insert image and caption
    body.insert(insert_idx, image_para)
    body.insert(insert_idx + 1, caption_para)

    # Save the modified document
    tree.write(output_path, xml_declaration=True, encoding='UTF-8', standalone=True)

    print("Successfully added image to Chapter 7")
    return True

def main():
    # File paths
    input_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-第七章修改版-目录修正.docx'
    output_file = '/Users/linxiansheng/Desktop/毕业论文/5月9日论文修改/E2220手机按键不锈钢片冲压模具设计与成型仿真-最终版.docx'

    # Extract docx
    temp_dir = '/tmp/docx_temp_image'
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)

    with zipfile.ZipFile(input_file, 'r') as z:
        z.extractall(temp_dir)

    # Check if image exists in the extracted file
    media_dir = os.path.join(temp_dir, 'word', 'media')
    if os.path.exists(media_dir):
        print(f"Media files: {os.listdir(media_dir)}")
    else:
        print("No media directory found")

    # Check relationships
    rels_file = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
    if os.path.exists(rels_file):
        with open(rels_file, 'r') as f:
            content = f.read()
            if 'rId102' in content:
                print("rId102 relationship exists")
            else:
                print("rId102 relationship NOT found")
                # Need to copy image from original and add relationship

    # Fix the document
    document_xml = os.path.join(temp_dir, 'word', 'document.xml')
    fixed_xml = '/tmp/document_fixed_image.xml'

    print("\\nAdding image to Chapter 7...")
    if add_image_to_chapter7(document_xml, fixed_xml):
        # Replace the document.xml
        shutil.copy2(fixed_xml, document_xml)

        # Create new docx
        with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    z.write(file_path, arcname)

        print(f"\\nFinal file saved to: {output_file}")
    else:
        print("Failed to add image")

    # Cleanup
    shutil.rmtree(temp_dir)

if __name__ == '__main__':
    main()
