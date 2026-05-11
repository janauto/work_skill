# 毕业设计论文修改 SKILL

> 适用于：Word文档程序化编辑、排版规范化、目录修复、内容扩充等场景

---

## 目录

- [1. 项目概述](#1-项目概述)
- [2. 核心能力](#2-核心能力)
- [3. 文件操作规范](#3-文件操作规范)
- [4. 排版规范](#4-排版规范)
- [5. 目录与书签](#5-目录与书签)
- [6. 表格处理](#6-表格处理)
- [7. 图片处理](#7-图片处理)
- [8. 公式处理](#8-公式处理)
- [9. 内容扩充](#9-内容扩充)
- [10. 常见问题](#10-常见问题)

---

## 1. 项目概述

### 1.1 背景

本SKILL记录了使用Python程序化编辑Word文档(.docx)的完整经验，特别针对毕业设计论文的排版规范化需求。

### 1.2 技术栈

- **python-docx**: Word文档读写（有限支持）
- **lxml**: XML直接操作（推荐）
- **zipfile**: docx文件解压/打包

### 1.3 文件结构

```
.docx文件
├── [Content_Types].xml
├── _rels/
├── word/
│   ├── document.xml          # 主要内容
│   ├── styles.xml            # 样式定义
│   ├── _rels/
│   │   └── document.xml.rels # 关系文件（图片引用等）
│   ├── media/                # 图片资源
│   └── embeddings/           # 嵌入对象
```

---

## 2. 核心能力

### 2.1 文档解析

```python
import zipfile
from lxml import etree

def parse_docx(docx_path):
    """解析docx文件，返回XML树"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        xml_content = z.read('word/document.xml')
    return etree.fromstring(xml_content)
```

### 2.2 段落文本提取

```python
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

def get_paragraph_text(p):
    """获取段落文本"""
    texts = p.findall('.//w:t', ns)
    return ''.join([t.text or '' for t in texts])

def get_paragraph_style(p):
    """获取段落样式"""
    pStyle = p.find('.//w:pStyle', ns)
    return pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') if pStyle is not None else ''
```

### 2.3 文档保存

```python
def save_docx(tree, temp_dir, output_path):
    """保存修改后的文档"""
    document_xml = os.path.join(temp_dir, 'word', 'document.xml')
    tree.write(document_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for root_dir, dirs, files in os.walk(temp_dir):
            for file in files:
                file_path = os.path.join(root_dir, file)
                arcname = os.path.relpath(file_path, temp_dir)
                z.write(file_path, arcname)
```

---

## 3. 文件操作规范

### 3.1 备份策略

```python
import shutil

def backup_file(input_file):
    """创建备份文件"""
    backup_file = input_file + '.backup'
    if not os.path.exists(backup_file):
        shutil.copy2(input_file, backup_file)
        print(f"Created backup: {backup_file}")
```

### 3.2 临时目录管理

```python
def extract_docx(input_file, temp_dir):
    """解压docx到临时目录"""
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    with zipfile.ZipFile(input_file, 'r') as z:
        z.extractall(temp_dir)
    return temp_dir
```

### 3.3 版本命名规范

```
原文件.docx
原文件-修改版.docx
原文件-修改版-v2.docx
原文件-最终版.docx
原文件-最终版-目录修复.docx
```

---

## 4. 排版规范

### 4.1 字体字号标准

| 元素 | 字体 | 字号 | 对齐 |
|------|------|------|------|
| 一级标题 | 黑体 | 三号(36) | 左对齐 |
| 二级标题 | 黑体 | 四号(28) | 左对齐 |
| 三级标题 | 黑体 | 小四(24) | 左对齐 |
| 正文 | 宋体 | 小四(24) | 两端对齐 |
| 图题/表题 | 宋体 | 五号(22) | 居中 |
| 参考文献 | 宋体 | 小四(24) | 左对齐 |

### 4.2 段落格式创建

```python
def create_body_paragraph(text, indent='480'):
    """创建正文段落 - 宋体小四，首行缩进2字符，两端对齐"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    # 段落属性
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    
    # 两端对齐
    jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'both')
    
    # 1.5倍行距
    spacing = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}spacing')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}line', '360')
    spacing.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}lineRule', 'auto')
    
    # 首行缩进
    ind = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind')
    ind.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}firstLine', indent)
    
    # 字体设置
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    
    rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ascii', '宋体')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
    
    sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '24')
    
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = text
    
    return p
```

### 4.3 标题格式创建

```python
def create_heading(text, level=2):
    """创建标题段落"""
    # 样式ID: 1=一级标题, 2=二级标题, 3=三级标题
    style_id = str(level)
    size = '36' if level == 1 else '28' if level == 2 else '24'
    
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    
    # 样式
    pStyle = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
    pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', style_id)
    
    # 左对齐
    jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'left')
    
    # 字体
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    
    rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '黑体')
    
    sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', size)
    
    b = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')
    
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = text
    
    return p
```

---

## 5. 目录与书签

### 5.1 目录结构

目录条目由以下部分组成：
- 样式（TOC1/TOC2/TOC3）
- 超链接（指向书签）
- 点引导线
- 页码

### 5.2 目录条目创建

```python
def create_toc_entry(text, style, bookmark_name, page_num=''):
    """创建目录条目"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    # 段落属性
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    
    # 样式
    pStyle = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pStyle')
    pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', style)
    
    # 制表位（页码右对齐）
    tabs = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tabs')
    tab = etree.SubElement(tabs, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab')
    tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'right')
    tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}leader', 'dot')
    tab.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pos', '9000')
    
    # 超链接
    hyperlink = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}hyperlink')
    hyperlink.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}anchor', bookmark_name)
    
    # 文本
    r = etree.SubElement(hyperlink, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = text
    
    # 页码
    if page_num:
        r_tab = etree.SubElement(hyperlink, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        tab_elem = etree.SubElement(r_tab, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tab')
        
        r_page = etree.SubElement(hyperlink, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        t_page = etree.SubElement(r_page, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        t_page.text = page_num
    
    return p
```

### 5.3 书签添加

**关键点：** 书签必须放在正确的位置（pPr之后，run之前）

```python
def add_bookmark_to_paragraph(p, bookmark_name, bookmark_id):
    """在段落添加书签"""
    # 创建bookmarkStart
    bookmarkStart = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkStart')
    bookmarkStart.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', bookmark_id)
    bookmarkStart.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}name', bookmark_name)
    
    # 创建bookmarkEnd
    bookmarkEnd = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bookmarkEnd')
    bookmarkEnd.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id', bookmark_id)
    
    # 找到pPr元素，在其后插入bookmarkStart
    pPr = p.find('w:pPr', ns)
    if pPr is not None:
        pPr_index = list(p).index(pPr)
        p.insert(pPr_index + 1, bookmarkStart)
    else:
        p.insert(0, bookmarkStart)
    
    # 找到第一个run，在其后插入bookmarkEnd
    first_run = p.find('w:r', ns)
    if first_run is not None:
        run_index = list(p).index(first_run)
        p.insert(run_index + 1, bookmarkEnd)
    else:
        p.append(bookmarkEnd)
```

### 5.4 常见问题

**问题：** 目录点击无法跳转
**原因：** 书签缺失或位置不正确
**解决：** 
1. 检查目录hyperlink的anchor属性
2. 确认正文中存在对应的bookmarkStart
3. 确保书签ID格式正确

---

## 6. 表格处理

### 6.1 创建Word表格

```python
def create_table(headers, rows, font='宋体', size='22'):
    """创建Word三线表"""
    tbl = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tbl')
    
    # 表格属性
    tblPr = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblPr')
    
    # 表格宽度
    tblW = etree.SubElement(tblPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblW')
    tblW.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w', '5000')
    tblW.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', 'pct')
    
    # 表格边框
    tblBorders = etree.SubElement(tblPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = etree.SubElement(tblBorders, f'{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{border_name}')
        border.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'single')
        border.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz', '4')
        border.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color', '000000')
    
    # 创建表头行
    header_row = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
    for header in headers:
        tc = etree.SubElement(header_row, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
        p = etree.SubElement(tc, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
        
        # 居中对齐
        pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
        jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
        jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')
        
        r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
        rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
        
        # 表头加粗
        b = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')
        
        t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
        t.text = header
    
    # 创建数据行
    for row_data in rows:
        tr = etree.SubElement(tbl, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr')
        for cell_text in row_data:
            tc = etree.SubElement(tr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tc')
            p = etree.SubElement(tc, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
            r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
            t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
            t.text = cell_text
    
    return tbl
```

### 6.2 图片表格替换

```python
def replace_image_table_with_real_table(body, table_title_text, headers, rows):
    """将图片表格替换为真正的Word表格"""
    # 1. 找到表题位置
    # 2. 删除表题后的图片段落
    # 3. 在表题后插入新表格
    pass
```

---

## 7. 图片处理

### 7.1 图片引用关系

```python
# 在document.xml.rels中
# rId102 -> media/image62.png
```

### 7.2 创建图片段落

```python
def create_image_paragraph(image_rId, cx='5269230', cy='3494405'):
    """创建图片段落"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    # 居中对齐
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')
    
    # Run包含Drawing
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    drawing = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}drawing')
    
    # Inline图片
    inline = etree.SubElement(drawing, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline')
    
    # 尺寸
    extent = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent')
    extent.set('cx', cx)
    extent.set('cy', cy)
    
    # 图片引用
    graphic = etree.SubElement(inline, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
    graphicData = etree.SubElement(graphic, '{http://schemas.openxmlformats.org/drawingml/2006/main}graphicData')
    graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
    
    pic = etree.SubElement(graphicData, '{http://schemas.openxmlformats.org/drawingml/2006/picture}pic')
    blipFill = etree.SubElement(pic, '{http://schemas.openxmlformats.org/drawingml/2006/picture}blipFill')
    blip = etree.SubElement(blipFill, '{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', image_rId)
    
    return p
```

### 7.3 图题创建

```python
def create_caption_paragraph(text):
    """创建图题段落 - 宋体五号加粗居中"""
    p = etree.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
    
    pPr = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr')
    jc = etree.SubElement(pPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}jc')
    jc.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'center')
    
    r = etree.SubElement(p, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    rPr = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    
    rFonts = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rFonts')
    rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '宋体')
    
    sz = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz')
    sz.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '22')
    
    b = etree.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b')
    
    t = etree.SubElement(r, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
    t.text = text
    
    return p
```

---

## 8. 公式处理

### 8.1 公式格式

公式使用三列表格实现：左空、中公式居中、右编号右对齐

### 8.2 公式编号

格式：`（章.序）`，如 `（2.1）`、`（3.5）`

---

## 9. 内容扩充

### 9.1 校核内容扩充原则

1. **概念先行**：先解释参数的含义和重要性
2. **数据支撑**：列出具体的计算过程和数值
3. **对比分析**：将设计值与标准值进行对比
4. **结论明确**：给出明确的校核结论
5. **避免AI味**：使用工程语言，不要过于模板化

### 9.2 示例：公称压力校核

```
（1）公称压力校核

压力机的公称压力是衡量其冲压能力的核心参数，直接决定了设备能否完成预定的冲压工序。
根据第三章的计算结果，本设计总冲压力F总=71.4kN，包括冲裁力63.7kN、卸料力3.2kN、
推件力4.0kN以及弯曲力0.92kN。按照冲压模具设计规范，所选压力机的公称压力应大于
总冲压力的1.3倍，即:

F公称 ≥ 1.3×F总 = 1.3×71.4 = 92.8kN

所选J23-160型开式可倾压力机的公称压力为160kN，远大于92.8kN的要求值。公称压力
裕量为160-92.8=67.2kN，安全系数实际达到160/71.4≈2.24，大于规范要求的1.3倍。
这一裕量可以保证在冲压过程中，即使遇到材料性能波动或润滑条件变化等不利因素，
压力机仍能稳定工作，不会出现过载现象。因此，公称压力校核满足要求。
```

---

## 10. 常见问题

### 10.1 目录无法跳转

**诊断步骤：**
1. 检查目录hyperlink的anchor属性值
2. 在正文中搜索对应的bookmarkStart
3. 确认bookmarkStart的name属性与anchor一致

**修复代码：**
```python
# 在正确的标题段落添加书签
add_bookmark_to_paragraph(heading_para, 'TOC_LINK_XXX', bookmark_id)
```

### 10.2 表格显示为图片

**原因：** 表格被保存为图片对象
**解决：** 删除图片段落，插入真正的Word表格

### 10.3 格式不统一

**解决：** 批量遍历段落，统一设置字体、字号、对齐、缩进

### 10.4 书签位置错误

**关键：** 书签必须在pPr之后、第一个run之前

```python
# 正确的插入位置
pPr_index = list(p).index(pPr)
p.insert(pPr_index + 1, bookmarkStart)  # 在pPr后插入
```

---

## 附录：常用XML命名空间

```python
ns = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}
```

---

## 版本记录

| 版本 | 日期 | 修改内容 |
|------|------|----------|
| v1.0 | 2026-05-11 | 初始版本，包含完整SKILL文档 |

---

## 许可证

MIT License
