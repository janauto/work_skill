# Summary Patterns

Write like a product-definition analyst, not like a generic BI report.

## Required Structure

Use this structure by default:

```text
总结：
本次聚焦 [focus]，共清洗 [原始条数] 条评论，保留 [有效条数] 条有效观点。高频一级功能集中在 [功能1]xx.xx%、[功能2]yy.yy%、[功能3]zz.zz%，一级功能下最聚焦的二级需求是 [需求1]aa.aa%、[需求2]bb.bb%、[需求3]cc.cc%。

1、用户喜欢什么
……

2、为什么放弃采用
……

3、一级功能与二级需求
……

4、嘿哈时刻
……

5、情绪热区
……

6、场景机会
……

7、底层需求归纳
……

8、合理增配
……
```

## Writing Rules

- Open with a diagnosis, not with raw counts alone
- Prefer one-word top-level function labels in prose
- Distinguish top-level function from second-level functional need
- Distinguish `喜欢什么` from `放弃采用`
- Treat valid suggestions as product-definition opportunity, not as bug by default
- Keep the aha section separate from ordinary positive feedback
- Highlight which feature or scene is emotionally hottest
- Refer to image-backed evidence when the workbook contains useful user images
- Make the opportunity statement concrete enough to influence concept definition

## Visualization Rules

- The Summary sheet should not rely on text alone
- Add charts for top functions, scenes, second-level needs, and aha concentration
- Reuse the same `功能 x 情绪标签` heatmap style in both `EmotionMap` and `Summary`
- When scene or aha sheets have image evidence, the summary narrative should explicitly mention what those image-backed rows reveal

## Rich-Text Rules

- The one-line overview should be bold in Excel
- Key labels and percentages should be red
- Do not color punctuation marks

## Example Phrases

Use phrasing like:

- `用户最稳定的喜欢点集中在……`
- `这类评论更像是在说……`
- `流失信号主要不是价格，而是……`
- `该类嘿哈时刻说明……`
- `可以沉淀为一个更值得定义的能力……`
- `这不是单点功能诉求，而是一条完整场景链路……`

Avoid phrasing like:

- `用户情绪复杂`
- `建议全面优化`
- `数据告诉我们一切`
