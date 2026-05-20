# Image Handling Rules

产品调研中图片获取、验证和引用的完整规范。

## 图片来源优先级

| 优先级 | 来源 | 说明 |
|--------|------|------|
| 1 | 官方博客 CDN | 如 storage.googleapis.com, whoop.com/cdn |
| 2 | 官方产品页 | store.google.com, apple.com 等 |
| 3 | 媒体评测 | theverge.com, cnet.com, wired.com |
| 4 | 电商平台 | Amazon, Best Buy 产品图 |
| 5 | ASCII 图表 | 所有来源失败时的最终回退 |

## 禁止的图片来源

| 来源 | 原因 |
|------|------|
| Unsplash / Pexels | 通用素材，非产品实图 |
| AI 生成图片 | 不真实，误导读者 |
| 其他产品的图片 | 张冠李戴 |
| 低分辨率缩略图 (< 5KB) | 无法辨识 |

## 下载与验证流程

### Step 1: 下载

```bash
curl -L -s -o "Attachments/[product]-[desc].jpg" "[URL]"
```

命名规范：`[产品名]-[描述].[扩展名]`
- 示例：`fitbit-air-hero.webp`
- 示例：`whoop-5-sensor.jpg`

### Step 2: 验证文件类型

```bash
file "Attachments/fitbit-air-hero.webp"
```

期望输出：
- `JPEG image data` ✅
- `RIFF...Web/P image` ✅
- `PNG image data` ✅
- `HTML document text` ❌ → 删除
- `ASCII text` ❌ → 删除

### Step 3: 验证文件大小

```bash
ls -lh "Attachments/fitbit-air-hero.webp"
```

判断标准：
- `> 10KB` ✅ 有效图片
- `1-10KB` ⚠️ 可能是缩略图，检查是否可用
- `< 1KB` ❌ 无效，删除
- `0B` ❌ 下载失败，删除

### Step 4: 失败处理

```bash
# 如果验证失败，立即删除
rm "Attachments/invalid-file.jpg"
```

## CDN 防盗链应对策略

| CDN 类型 | 表现 | 应对 |
|----------|------|------|
| Google Storage | 特定 width 参数才有效 | 尝试 width-1600, width-800, width-100 |
| Shopify CDN | 返回 HTML 错误页 | 换 Amazon 或评测文章来源 |
| Amazon CDN | 返回 9B 文本 | 换其他来源 |
| Whoop CDN | 403 Forbidden | 使用 ASCII 图表替代 |

## Obsidian 图片引用格式

### 本地图片（推荐）

```markdown
![[fitbit-air-hero.webp]]
```

### 带尺寸控制

```markdown
![[fitbit-air-hero.webp|500]]
```

### 外部链接（备用）

```markdown
![描述](https://example.com/image.jpg)
```

### 表格内图片（避免使用）

表格内嵌图片容易导致排版问题，建议：
- 图片放在表格外面
- 表格内用文字描述 + 外部链接

## 图片组织结构

```
项目目录/
├── Attachments/
│   ├── fitbit-air-hero.webp        # 产品主图
│   ├── fitbit-air-bands.webp       # 表带展示
│   ├── whoop-5-product.jpg         # 竞品图
│   └── ...
└── 产品调研报告.md
```

## 常见问题

### Q: Google Blog 图片只有 100px 版本怎么办？

Google Blog 的 `_small` 后缀图片只有 width-100 可用。解决方案：
1. 尝试不带 `_small` 的文件名
2. 尝试 `width-1600` 版本（hero 图通常有）
3. 如果都不行，使用 ASCII 图表

### Q: 所有图片来源都失败了怎么办？

1. 在报告中使用 ASCII 图表展示产品外观
2. 添加外部链接引用：`> 📷 产品图片来源：[官网链接](URL)`
3. 不要用无关图片凑数
