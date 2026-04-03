# Summary Patterns

Write like a product manager reviewing user voice, not like a generic BI report.

## Required Structure

Use this structure by default:

```text
总结：
本次聚焦 [产品]，共匹配 [总记录数] 条记录，去重清洗后保留 [清洗后记录数] 条，其中有效反馈 [有效反馈数] 条。综合来看，当前用户对 [产品] 的核心不满主要集中在 [问题1] 和 [问题2]。Top3 为 [分类1]xx.xx%、[分类2]yy.yy%、[分类3]zz.zz%。

说明：[无评论或无效评论数] 条无评论或无效评论未纳入有效反馈统计。

时间趋势上，[峰值月份或日期] 为有效反馈峰值周期，占全部有效反馈 [峰值占比]。

1、[重点问题1]（xx.xx%）
主要问题是……

2、[重点问题2]（yy.yy%）
主要问题是……

3、[重点问题3]（zz.zz%）
主要问题是……

4、[风险/合理需求/优化建议]
……
```

Adapt the number of points to the evidence. Three or four is usually enough.

## Visual Hierarchy Rules

- The single-sentence overview should be bold in Excel rich text.
- Key labels, percentages, and risk words should be red.
- Do not color punctuation marks such as `，` `。` `：` `、` `（` `）`.
- If a highlighted term sits next to punctuation, color only the term itself.
- Keep the text readable even when no color is available. The wording itself should still make the priority clear.
- If the workbook contains representative examples in table cells, rewrite those examples into short analyst summaries instead of quoting raw user language.
- Do not apply that rewriting rule to the final long-form summary. The summary itself should remain complete narrative prose.

## Writing Rules

- Open with a diagnosis, not with raw counts alone.
- For return workbooks, separate “matched rows” from “valid feedback rows” and explicitly say that no-comment reason-code rows were excluded.
- Mention percentages when they help prioritization.
- Distinguish subjective dissatisfaction from hard failure.
- Distinguish user mistake from product issue.
- Highlight risk items even when they are not the largest category.
- Treat repeated requests for missing features as “合理需求”, not as defects.
- Do not intentionally compress the final summary into brief labels or scanning shorthand. Keep the summary wording natural, complete, and suitable for direct presentation.

## Severity Bias

Prefer emphasizing these items in prose:

- 无法开机
- 电源异常
- 保护
- 间歇性静音
- 单声道/声道异常
- HDMI/光纤/接口兼容性失效
- 噪声、爆音、破裂音

These can outrank a mildly larger but low-risk user-reason category in the narrative.

## Example Phrases

Use phrasing like:

- `核心不满主要集中在……`
- `退货前三名为……`
- `可以进一步收敛为……`
- `无评论记录已从有效反馈统计中剔除`
- `时间趋势上，峰值出现在……`
- `代表诉求可概括为……`
- `主要问题是用户觉得……`
- `该需求相对合理，可评估……`
- `存在需要重点排查的风险……`
- `更像是页面表达或用户认知问题，而非硬件故障`

Avoid phrasing like:

- `用户情绪复杂`
- `数据表明大家不开心`
- `建议全面优化所有环节`

## Recommendation Style

When giving follow-up guidance, keep it concrete:

- `优先复核页面表述`
- `排查电源保护和静音问题`
- `评估是否补足耳机口或高通滤波需求`
- `优化增益范围或说明文案`

Do not give generic recommendations without tying them to a category and evidence.
