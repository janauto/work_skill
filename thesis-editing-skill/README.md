# 毕业设计论文修改 SKILL

> 使用 Python 程序化编辑 Word 文档的完整解决方案

## 简介

本 SKILL 记录了使用 Python (lxml + zipfile) 程序化编辑 Word 文档(.docx)的完整经验，特别针对毕业设计论文的排版规范化需求。

## 功能特性

- **目录修复**：修复目录链接，确保点击可跳转
- **书签管理**：添加、修复书签，关联目录超链接
- **排版规范化**：统一字体、字号、对齐、缩进
- **表格处理**：将图片表格替换为真正的Word表格
- **图片插入**：按规范插入图片和图题
- **内容扩充**：扩充校核内容，使其充实、符合逻辑

## 文件结构

```
thesis-editing-skill/
├── README.md                           # 本文件
├── 毕业设计论文修改SKILL.md              # 完整SKILL文档
├── AGENT.md                            # 项目工作记录
├── fix_chapter7_v2.py                  # 第七章结构调整
├── fix_toc.py                          # 目录条目修复
├── fix_bookmarks_v2.py                 # 书签位置修复
├── expand_chapter7.py                  # 校核内容扩充
└── add_image_to_ch7.py                 # 图片插入
```

## 快速开始

### 1. 安装依赖

```bash
pip install lxml
```

### 2. 基本用法

```python
import zipfile
from lxml import etree

# 解析docx
with zipfile.ZipFile('input.docx', 'r') as z:
    xml_content = z.read('word/document.xml')
root = etree.fromstring(xml_content)

# 修改文档...

# 保存
tree.write('word/document.xml', xml_declaration=True, encoding='UTF-8')
```

### 3. 修复目录

```bash
python fix_toc.py
```

### 4. 扩充内容

```bash
python expand_chapter7.py
```

## 排版规范

| 元素 | 字体 | 字号 | 对齐 |
|------|------|------|------|
| 一级标题 | 黑体 | 三号(36) | 左对齐 |
| 二级标题 | 黑体 | 四号(28) | 左对齐 |
| 正文 | 宋体 | 小四(24) | 两端对齐 |
| 图题/表题 | 宋体 | 五号(22) | 居中 |

## 常见问题

### 目录无法跳转

**原因：** 书签缺失或位置不正确

**解决：** 使用 `fix_bookmarks_v2.py` 修复

### 表格显示为图片

**原因：** 表格被保存为图片对象

**解决：** 删除图片段落，插入真正的Word表格

## 相关项目

- [work_skill](https://github.com/janauto/work_skill) - 工作技能集合

## 许可证

MIT License
