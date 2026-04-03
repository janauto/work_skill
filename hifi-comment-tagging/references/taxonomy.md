# Canonical Taxonomy

Use this file as the default label system for HIFI comment analysis.

## Canonical Chain

- `一级分类`: broad problem family
- `二级分类`: more specific issue bucket
- `三级问题点`: analyst-readable problem statement
- `四级归因`: likely root-cause layer

If the source workbook only has 2 or 3 levels, map what exists and leave the missing levels empty or `NA`.

## Display Naming Rule

When creating new output labels in generated workbooks, prefer concise slash-free display names when the meaning stays unchanged.

Examples:

- prefer `质量故障` over `质量/故障`
- prefer `音质体验` over `音质/性能体验`
- prefer `兼容连接` over `兼容性/连接问题`
- prefer `价格竞品` over `价格/替代品/竞品`

Keep the underlying category meaning aligned with the canonical taxonomy even when the display form is shortened.

## Priority Rules

Tag by the dominant actionable issue, not by every phrase in the sentence.

Use this order when multiple signals coexist:

1. Safety, fault, silence, protection, or hard failure
2. Compatibility or connection failure
3. Noise, distortion, or obvious performance defect
4. Clear user reason such as wrong order or no longer needed
5. Missing feature or product-definition gap
6. Subjective dissatisfaction or vague expectation gap

## Canonical Level-1 Categories

### 音质 / 音质/性能体验

Use for:

- 未达到预期
- 声音差
- 失真
- 高频刺耳
- 声场、解析、低频、中频等主观表现问题

Common level-2:

- `音质预期不符`
- `整体体验不达预期`
- `失真/破音`

Typical level-4:

- `产品定义问题`
- `硬件设计问题`
- `页面表达问题`

### 噪音

Use for:

- 底噪
- 电流声
- 爆音
- 播放破裂音

Common level-2:

- `噪声/底噪/电流声`

Typical level-4:

- `硬件设计问题`
- `兼容性适配问题`

### 质量/故障 / 品质 / 故障

Use for:

- 无法工作
- 单声道
- 立体声失效
- 无法开机
- 保护
- 静音
- 配件缺失但属于出厂/品控问题

Common level-2:

- `功能故障`
- `损坏`
- `声道`
- `无法正常工作`

Typical level-4:

- `个体质量问题`
- `仓配/版本管理问题`

### 兼容性 / 兼容性/连接问题

Use for:

- 接口不匹配
- HDMI CEC 不支持
- 光纤播放断续
- 与现有系统接法不兼容

Common level-2:

- `与系统不兼容`
- `设备不支持HDMI CEC功能`

Typical level-4:

- `兼容性适配问题`

### 用户原因 / 用户

Use for:

- 不再需要
- 需求变化
- 买错型号
- 下错单
- 找到替代品

Common level-2:

- `需求变化/不再需要`
- `下错单/买错型号`
- `找到替代品/更换其他产品`

Typical level-4:

- `用户认知问题`

### 功能 / 功能需求 / 接口问题

Use for:

- 缺少 AirPlay
- 需要耳机口
- 需要高通滤波
- 需要更合理的音量/增益逻辑

Common level-2:

- `音量/增益控制问题`
- `Subwoofer相关功能问题`
- `需要AirPlay功能`

Typical level-4:

- `产品定义问题`

### 页面/预期不符 / 体验

Use for:

- 页面描述与实际不符
- 宣传误导
- 实际体验与承诺不一致
- 怀疑二手、包装异常、非全新

Common level-2:

- `与描述不一致`
- `疑似二手/包装问题`
- `整体体验不达预期`

Typical level-4:

- `页面表达问题`
- `仓配/版本管理问题`

### 价格 / 价格/替代品/竞品 / 降价

Use for:

- 降价退货
- 竞品更便宜
- 亚马逊价格更低

Common level-2:

- `竞品/替代方案`
- `同款降价`
- `平台价格差异`

Typical level-4:

- `产品定义问题`

### 物流 / 包装 / 配件

Use for:

- 延迟送达
- 包装损坏
- 缺少配件
- 商品疑似二手

Typical level-4:

- `仓配/版本管理问题`

## Common Level-4 Root Causes

Prefer these root-cause labels over inventing near-synonyms:

- `用户认知问题`
- `页面表达问题`
- `兼容性适配问题`
- `个体质量问题`
- `硬件设计问题`
- `产品定义问题`
- `仓配/版本管理问题`
- `暂无法判断`

## Legacy Mapping

Map older sheets like this:

- `问题分类：1级 -> level_1`
- `问题分类：2级 -> level_2`
- `问题分类：3级 -> level_3`

When a legacy level-1 value is coarse:

- `用户问题` usually maps to `用户原因`
- `故障` usually maps to `质量/故障`
- `兼容` usually maps to `兼容性/连接问题`
- `价格问题` usually maps to `价格/替代品/竞品`

## Example Chains

- `用户原因 -> 下错单/买错型号 -> 型号或规格选错 -> 用户认知问题`
- `质量/故障 -> 功能故障 -> 电源异常/无法正常通电 -> 个体质量问题`
- `兼容性/连接问题 -> 与系统不兼容 -> 接口或系统接法不匹配 -> 兼容性适配问题`
- `音质 -> 音质预期不符 -> 声音表现与预期存在落差 -> 产品定义问题`
- `噪音 -> 噪声/底噪/电流声 -> 噪声增加或底噪明显 -> 硬件设计问题`
