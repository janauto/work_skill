# Canonical Taxonomy

Use this file as the default tagging system for product-definition VOC.

## Canonical Layers

- `观点类型`
- `一级功能`
- `二级功能需求`
- `底层需求`
- `决策信号`
- `情绪极性`
- `情绪强度`
- `场景标签`
- `嘿哈时刻`

## 1. 观点类型

Use one dominant label:

- `优点`
- `抱怨`
- `建议`
- `对比`
- `期望落差`
- `缺陷复现`
- `疑似灌水`
- `观点其他`

Priority rules:

1. Clear bug or reproducible failure beats everything else
2. Clear suggestion beats generic positive or negative wording
3. Explicit compare beats plain praise or complaint
4. Strong expectation mismatch beats vague dissatisfaction

## 2. 一级功能

Prefer these reusable labels:

- `切换`
- `连接`
- `音量`
- `音质`
- `体积`
- `做工`
- `价格`
- `外观`
- `前级`
- `场景`
- `基础`

Use one-word top-level labels by default. Avoid composite labels that pack multiple concepts into one tag.

## 2.1 二级功能需求

Under each top-level function, add one concrete second-level functional need.

Examples:

- `音量 -> 音量范围更大`
- `音量 -> 电平匹配更准确`
- `连接 -> 接口覆盖更全`
- `连接 -> 连接更稳定`
- `切换 -> 一键切换更直接`
- `切换 -> 自动切换更聪明`
- `音质 -> 底噪更低`
- `前级 -> 补上前级控制`

Rule:

- Prefer a concrete functional ask over a broad hidden-need label
- Use this layer to avoid repeating the same meaning between `一级功能` and `底层需求`

## 3. 底层需求

Prefer these root needs:

- `降摩擦`
- `连接确定性`
- `控制确定性`
- `音质稳定`
- `稳定可靠`
- `空间效率`
- `系统补足`
- `场景适配`
- `性价比安全感`
- `品质认同`
- `灵感启发`

When in doubt, prefer the deepest actionable need, not the surface symptom.

## 4. 决策信号

Use:

- `喜欢什么`
- `放弃采用`
- `合理增配`
- `噪声样本`
- `待归纳`

Interpretation:

- `喜欢什么`: explains why the product earns adoption or delight
- `放弃采用`: explains churn, regret, refusal, or abandonment
- `合理增配`: valid feature or definition opportunity
- `噪声样本`: exaggerated praise, suspicious flooding, or low-trust wording

## 5. 情绪标签

### 情绪极性

- `正向`
- `中性`
- `负向`

### 情绪强度

- `轻微`
- `中等`
- `强烈`

Use stronger intensity when the comment contains:

- Explicit delight or anger
- Repeated emphasis
- Failure consequences
- Strong words such as `完美`, `终于`, `垃圾`, `无法使用`

### 情绪标签

Prefer these reusable labels:

- `惊喜超预期`
- `顺手便利`
- `放心省心`
- `值回票价`
- `灵感打开`
- `满意认可`
- `被噪音打断`
- `失控难调`
- `焦虑不确定`
- `失望落差`
- `愤怒拒绝`
- `烦躁麻烦`
- `理性期待`
- `理性对比`
- `观察确认`

Also tag `情绪触发点`, usually the dominant `一级功能`.

## 6. 场景标签

Prefer these reusable scene clusters when they fit:

- `桌面双机`
- `耳机音箱切换`
- `家庭影音补足`
- `黑胶系统升级`
- `录音创作`
- `游戏主机`
- `DIY项目`
- `小空间部署`
- `通用补洞`

Each scene card should later describe:

- `人物`
- `场地`
- `动机`
- `时空体验链路`

## 7. 嘿哈时刻

Mark a row as `是` only when the comment clearly shows one of these:

- The user unexpectedly discovered a perfect fit
- The product solved a stubborn pain point with surprising simplicity
- The user revealed an unusual use case or product insight
- The comment implies a new product-definition direction

Rank aha rows by:

1. Relevance to the product itself
2. Distinctiveness of the insight
3. Clarity of the pain-point to solution chain
4. Strength of scene detail

## Example Chains

- `优点 -> 切换 -> 一键切换更直接 -> 降摩擦 -> 喜欢什么`
- `抱怨 -> 连接 -> 连接更稳定 -> 连接确定性 -> 放弃采用`
- `建议 -> 音量 -> 音量范围更大 -> 控制确定性 -> 合理增配`
- `缺陷复现 -> 做工 -> 寿命更长 -> 稳定可靠 -> 放弃采用`
- `对比 -> 前级 -> 补上前级控制 -> 系统补足 -> 喜欢什么`
