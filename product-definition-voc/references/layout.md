# Layout Guide

Keep the output close to the existing HIFI analysis style, but adapt it for product-definition VOC.

## Sheet Order

Use this order:

1. `CleanedComments`
2. `TaggedComments`
3. `NeedClusters`
4. `AhaMoments`
5. `EmotionMap`
6. `SceneCards`
7. `Summary`

## Detail Sheets

Preferred left-to-right order for `TaggedComments`:

- `focus_name`
- `product_name`
- `scene_name`
- `rating`
- `cleaned_comment`
- `观点类型`
- `一级功能`
- `二级功能需求`
- `底层需求`
- `决策信号`
- `情绪极性`
- `情绪强度`
- `场景标签`
- `嘿哈时刻`
- `排序分`

Formatting:

- Freeze header row
- Bold filled headers
- Wrap text on comment, scene-chain, and analysis fields
- Preserve one row per comment for audit

## NeedClusters

Top block:

- `马斯洛层级`
- `数量`
- `占比`
- `层级说明`

Detail block:

- `马斯洛层级`
- `底层需求`
- `一级功能`
- `二级功能需求`
- `主要场景`
- `数量`
- `占比`
- `产品定义启发`

Visualize the top block as a triangular Maslow pyramid image, then sort the detail block by count descending.
For layer assignment, prefer semantic judgment from a large model when credentials exist; otherwise use the built-in semantic fallback.

## AhaMoments

Columns:

- `排名`
- `嘿哈分数`
- `关联度`
- `独到性`
- `商品名`
- `商品链接`
- `场景标签`
- `一级功能`
- `二级功能需求`
- `简明提炼`
- `源评论原文`
- `痛点机会`
- `产品定义启发`
- `图片证据1`
- `图片证据2`

Keep this sheet sorted by aha score descending.
The concise aha line should be a cleaned, product-definition-ready sentence rather than a raw pasted comment.
Prefer the source raw comment text, not the translated text, in `源评论原文`.
When local image files exist, embed the original user images directly into the sheet instead of keeping only path text.
If a high-value aha row has no user image, fall back to the product image.

## EmotionMap

Put the actual heatmap blocks first, then the compact detail table below them.

Recommended columns:

- `一级功能`
- `情绪标签`
- `情绪极性`
- `数量`
- `占比`
- `热度说明`

Also include:

- `功能 x 情绪标签热力图`
- `场景 x 情绪标签热力图`

These should be true Excel heatmaps built with color-scale conditional formatting, not only prose labels.

## SceneCards

Columns:

- `场景标签`
- `商品名`
- `数量`
- `占比`
- `代表人物`
- `代表场地`
- `核心动机`
- `时空体验链路`
- `代表观点`
- `场景图片证据1`
- `场景图片证据2`

Prefer representative rows that also带有用户原图 so the scene card includes visual evidence.
If the representative row has no user image, use the product image as fallback evidence.

## Summary Sheet Structure

Use four visible blocks plus the long-form narrative area.

### Left block

Columns `A:D`

- `一级功能：先按话题看`
- `一级功能`
- `二级功能需求`
- `数量`
- `占比`

### Middle block

Columns `G:J`

- `先看高频：一级功能筛选`
- `一级功能`
- `二级功能需求`
- `数量`
- `占比`

### Right-middle block

Columns `L:O`

- `功能 x 场景排序`
- `一级功能`
- `场景`
- `数量`
- `占比`

### Right block

Merged area `R:Y`

Place the long-form Chinese narrative summary here.

### Lower area

Reserve space below for:

- Top feature chart
- Top scene chart
- Top second-level-need chart
- Aha-scene bar chart
- A compact Summary heatmap

The Summary sheet should visibly cover all major blocks: functions, second-level needs, hidden needs, aha moments, emotion heat, and scenes.
Use light background colors to distinguish different `一级功能` rows in the visible Summary tables.

Keep colors restrained and the tables readable.
