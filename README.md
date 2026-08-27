# Work Skill

这是一个面向产品研究、用户评论分析、竞品评论抓取、硬件DVT评审、专利附图CAD处理和结构化洞察输出的 Codex Skill 工作集。技能包覆盖外部评论采集、Excel 清洗、中文标签体系归因、产品定义 VOC、硬件整机透视爆炸图，以及专利框线图的实线化、去附图标记编号与DWG验证。

## 仓库定位

本仓库适合用于以下场景：

- 从 Amazon 商品页搜索竞品、预估评论规模、抓取评论和评论图片，并导出双语 Excel。
- 从 Audio Science Review 论坛收集音频产品相关讨论，提取帖子、图片、翻译和打标表。
- 对 HIFI 产品评论、退货原因、售后反馈进行清洗、归因、分层标签和汇总。
- 对用户评论 VOC 做产品定义分析，提炼喜欢点、流失原因、隐藏需求、Aha moment 和场景卡片。

设计原则：

- 先确认范围，再执行采集或分析。
- 保留来源行号、文件、sheet、图片等证据链。
- 输出以可复核 Excel 为主，而不是只给结论文本。
- 标签体系尽量复用既有分类，避免临时发明近义标签。
- 对抓取任务采用预检和人工确认机制，降低无效执行成本。

## 目录总览

```text
.
├── amazon-review-scraping-skill/
│   ├── SKILL.md
│   ├── README.md
│   ├── package.json
│   ├── references/
│   └── scripts/
├── asr-review-scraping-skill/
│   ├── SKILL.md
│   ├── README.md
│   ├── package.json
│   ├── requirements.txt
│   ├── references/
│   ├── runs/
│   └── scripts/
├── hifi-comment-tagging/
│   ├── SKILL.md
│   ├── README.md
│   ├── references/
│   └── scripts/
├── dvt-exploded-model-visualizer/
│   ├── SKILL.md
│   ├── agents/
│   ├── assets/
│   ├── references/
│   └── scripts/
├── patent-drawing-dwg-cleanup/
│   ├── SKILL.md
│   ├── agents/
│   ├── references/
│   └── scripts/
├── product-definition-voc/
│   ├── SKILL.md
│   ├── references/
│   └── scripts/
└── product-definition-voc.skill
```

## 技能包说明

### 1. Amazon Review Scraping Skill

路径：`amazon-review-scraping-skill/`

用于 Amazon 商品搜索、评论抓取、评论图片下载和多商品 Excel 输出。适合竞品研究、评论 VOC 收集、品类范围预估等任务。

核心能力：

- 支持 `amazon.sg`、`amazon.com` 等 Amazon 站点。
- 根据 2-3 个关键词生成候选商品和编号场景草稿。
- 在正式抓取前输出预检结果，包括候选商品数、评论规模估算和 Top N 预览。
- 支持用户用 `不搜索：1、8、11` 排除不需要的场景。
- 只有用户明确回复 `开始执行` 后才进入真实抓取。
- 使用 Playwright 持久化浏览器会话，可人工登录并复用登录状态。
- 抓取 top reviews、recent reviews、positive reviews、critical reviews 等评论视图。
- 下载评论图片，去重评论，并生成多商品 Excel。
- 内置普通页面抓取和 stealth 反爬页面抓取脚本。

常用命令：

```bash
cd amazon-review-scraping-skill
npm install
npx playwright install chromium
python3 -m pip install openpyxl pillow requests
```

预检并等待确认：

```bash
node scripts/amazon-preflight-workflow.js \
  --marketplace amazon.sg \
  --keywords "rca switcher,3.5mm switcher,audio selector" \
  --category Electronics \
  --price-min 10 \
  --price-max 60 \
  --min-rating 4.0 \
  --top-n 5 \
  --output-dir "./output"
```

排除部分场景：

```bash
node scripts/amazon-preflight-workflow.js \
  --state "./output/preflight_state.json" \
  --reply "不搜索：1、8、11"
```

确认执行：

```bash
node scripts/amazon-preflight-workflow.js \
  --state "./output/preflight_state.json" \
  --reply "开始执行"
```

普通动态网页抓取：

```bash
node scripts/playwright-simple.js "https://example.com"
```

Cloudflare、403 或强反爬页面：

```bash
HEADLESS=false SAVE_HTML=true node scripts/playwright-stealth.js "https://example.com"
```

主要输出：

- `preflight_state.json`
- 抓取运行 manifest
- 单商品评论 JSON
- 评论图片目录
- 多商品评论 Excel，包含场景、候选商品、评论明细和抓取汇总

### 2. ASR Review Scraping Skill

路径：`asr-review-scraping-skill/`

用于采集 Audio Science Review 论坛帖子，整理音频产品相关讨论，并生成可打标 Excel。适合研究 DAC、前级、切换器、AVR、功放等音频产品的真实用户讨论。

核心能力：

- 从 ASR 论坛线程中提取帖子内容、作者、时间、链接等元数据。
- 使用 `r.jina.ai` 文本镜像读取论坛正文。
- 从本地 Chrome 缓存中恢复附件原图，避免只保存低质量截图。
- 可生成中文翻译列和标签列。
- 输出单 sheet Excel，方便后续人工或半自动打标。
- 内置 Playwright simple 和 stealth 抓取脚本，可用于其他网页。

安装依赖：

```bash
cd asr-review-scraping-skill
npm install
npx playwright install chromium
python3 -m pip install -r requirements.txt
```

完整 ASR 流程：

```bash
python3 scripts/run_asr_pipeline.py --dataset-root runs/default
```

只抓线程：

```bash
python3 scripts/fetch_asr_threads.py --dataset-root runs/default
```

只重建 Excel：

```bash
python3 scripts/build_asr_workbook.py --dataset-root runs/default
```

使用自定义 URL 列表：

```bash
python3 scripts/run_asr_pipeline.py \
  --dataset-root runs/project-a \
  --urls-file /abs/path/curated_threads.txt
```

主要输出：

- `raw_threads/`
- `thread_index.json`
- `thread_summary.md`
- `downloaded_images/`
- `preview_images/`
- `translation_cache.json`
- `ASR_切换相关用户内容_打标准备.xlsx`

翻译说明：

- 如果存在 `ZHIPUAI_API_KEY`、`ZHIPU_API_KEY` 或 `BIGMODEL_API_KEY`，脚本会尝试调用智谱 API 生成中文翻译。
- 如果没有 API key，会复用已有缓存，未命中的翻译列保持空白。

### 3. HIFI Comment Tagging

路径：`hifi-comment-tagging/`

用于分析 HIFI 产品评论、退货原因和售后反馈，适合围绕单个产品做结构化问题归因和产品经理式总结。

核心能力：

- 读取 Excel 工作簿并识别候选 sheet、表头和产品信号。
- 聚焦单个目标产品，例如 `P4`、`ZD3`、`ZA3`、`ZP3`、`LC30`、`MC331`。
- 清洗空评论、无效文本、重复评论和无实质反馈的退货原因。
- 保留原始文件、sheet、行号等来源信息。
- 使用 1-4 级中文分类链：
  - `一级分类`
  - `二级分类`
  - `三级问题点`
  - `四级归因`
- 可从历史人工标注表中提取复用标签和示例。
- 输出清洗表、打标表和总结表。

典型流程：

```bash
cd hifi-comment-tagging
python3 scripts/profile_workbook.py /path/to/input.xlsx
python3 scripts/clean_product_comments.py /path/to/input.xlsx --product P4 --output cleaned.xlsx
python3 scripts/extract_taxonomy_examples.py /path/to/labeled.xlsx --output taxonomy_examples.xlsx
python3 scripts/build_summary_scaffold.py /path/to/tagged.xlsx --output summary.xlsx
```

主要输出：

- `CleanedComments`：聚焦目标产品后的标准化有效反馈。
- `TaggedComments`：带完整分类链的评论明细。
- `Summary`：分类统计、收敛路径、关键问题、趋势和产品经理式总结。

适用边界：

- 每次默认只分析一个产品。
- 如果源文件包含多个产品，必须先明确目标产品。
- 对退货分析，空白买家备注默认不计入有效用户反馈，但会保留在审计信息中。

### 4. Product Definition VOC

路径：`product-definition-voc/`

用于从用户评论和竞品评论中提炼产品定义洞察，适合概念定义、竞品学习、预 PRD 洞察和功能机会识别。

核心能力：

- 对一个产品、品类、场景或竞品集合做 VOC 分析。
- 清洗重复、无效、物流、卖家服务、优惠券、跑题等评论。
- 保留原始评论、翻译、评分、图片引用、商品链接和来源行号。
- 使用产品定义向标签体系：
  - `观点类型`
  - `一级功能`
  - `二级功能需求`
  - `底层需求`
  - `决策信号`
  - `情绪极性`
  - `情绪强度`
  - `场景标签`
  - `嘿哈时刻`
- 输出隐藏需求、Aha moment、情绪热区、场景卡片和机会提示。

典型流程：

```bash
cd product-definition-voc
python3 scripts/profile_workbook.py /path/to/input.xlsx
python3 scripts/clean_voc_comments.py /path/to/input.xlsx --focus "RCA 切换器竞品" --output cleaned.xlsx
python3 scripts/tag_voc_comments.py cleaned.xlsx --output tagged.xlsx
python3 scripts/build_voc_summary_workbook.py tagged.xlsx --output voc_summary.xlsx
```

主要输出：

- `CleanedComments`
- `TaggedComments`
- `NeedClusters`
- `AhaMoments`
- `EmotionMap`
- `SceneCards`
- `Summary`

总结重点：

- 用户明确喜欢什么。
- 用户为什么犹豫、流失或停用。
- 隐藏需求及占比。
- Aha moment 的原始证据。
- 不同功能和场景下的情绪热区。
- 可转化为产品定义的机会提示。

### 5. DVT Exploded Model Visualizer

路径：`dvt-exploded-model-visualizer/`

用于从 STEP/STP/GLB 整机模型生成可交互的透视爆炸图，并在不混淆原始CAD和方案概念件的前提下，展示DVT局部结构修改。适合结构评审、装配说明、改模沟通、光学/电子模组堆叠和验证计划。

核心能力：

- 输入文件扫描，识别STEP/STP、GLB/GLTF、2D图、BOM、需求文档和占位文件。
- 生成爆炸距离、透视、标准视图、模块显隐和点选查看等交互。
- 将源CAD、`PROPOSED DVT`和`CONCEPT ONLY`分开标识，避免把概念叠加件写成已冻结CAD。
- 使用实际可制造截面展示平板环、导光件、遮光支架、FPC尾线、安装耳、螺钉柱和紧固方向。
- 在HTML中加入“开始前请放入这些文件”提醒、BOM、装配步骤、工艺步骤和DVT检验关卡。
- 使用真实浏览器回归验证模型加载、交互、桌面/移动布局及控制台错误。

输入文件扫描示例：

```bash
python3 dvt-exploded-model-visualizer/scripts/inspect_model_bundle.py /abs/path/to/project
```

主要输出：

- 可交互HTML透视爆炸图。
- 派生GLB/元数据和组装/爆炸预览图。
- 局部修改方案、制造工艺、装配顺序、验证要求和待补文件清单。

### 6. Patent Drawing DWG Cleanup

路径：`patent-drawing-dwg-cleanup/`

用于把DWG、DXF、Matplotlib、PDF或位图来源的专利框线图整理成连续实线CAD，移除专利附图标记编号，同时保留轴号、端子名、电压、尺寸等功能标注，并使用AutoCAD原生核心引擎生成和审计DWG。

核心能力：

- 在删除前盘点数字文本，区分专利引用编号与`A1`、`24V`、尺寸、公差等有效工程信息。
- 将实体和图层线型统一为`CONTINUOUS`，避免只在预览中看似实线。
- 支持从Matplotlib artist导出LINE、POLYLINE、CIRCLE、箭头和文字。
- 通过ASCII临时保存路径规避AutoCAD macOS命令脚本对中文输出路径的乱码问题。
- 使用AutoCAD Core Console两次`AUDIT`，并检查实体数和非实线对象数。
- 同时交付可编辑文字DXF与可选的字体无关矢量文字DWG。

典型流程：

```bash
python3 patent-drawing-dwg-cleanup/scripts/clean_patent_dxf.py \
  input.dxf cleaned.dxf \
  --reference-number 100 \
  --reference-number 310 \
  --report cleanup-report.json

python3 patent-drawing-dwg-cleanup/scripts/validate_clean_dxf.py cleaned.dxf

python3 patent-drawing-dwg-cleanup/scripts/autocad_core_dxf_to_dwg.py \
  cleaned.dxf cleaned.dwg
```

## 快速选择指南

| 需求 | 推荐技能包 |
| --- | --- |
| 抓 Amazon 商品评论和评论图片 | `amazon-review-scraping-skill` |
| 先按关键词找竞品范围，确认后再抓 | `amazon-review-scraping-skill` |
| 抓 ASR 论坛帖子并做音频用户讨论表 | `asr-review-scraping-skill` |
| 分析单个 HIFI 产品的评论、退货、售后问题 | `hifi-comment-tagging` |
| 从评论中提炼产品定义、隐藏需求和 Aha moment | `product-definition-voc` |
| 从整机CAD生成透视爆炸图，并展示DVT局部改型 | `dvt-exploded-model-visualizer` |
| 把专利框线图转为实线、去附图编号并生成已审计DWG | `patent-drawing-dwg-cleanup` |
| 只抓普通 JS 页面 | `playwright-simple.js` |
| 抓 Cloudflare 或 403 页面 | `playwright-stealth.js` |

## 环境要求

建议环境：

- macOS 或 Linux
- Python 3.10+
- Node.js 18+
- npm
- Chromium Playwright browser

常用 Python 依赖：

```bash
python3 -m pip install openpyxl pillow requests ezdxf
```

ASR 额外依赖：

```bash
python3 -m pip install -r asr-review-scraping-skill/requirements.txt
```

Playwright 依赖：

```bash
cd amazon-review-scraping-skill
npm install
npx playwright install chromium

cd ../asr-review-scraping-skill
npm install
npx playwright install chromium
```

## 输出文件规范

推荐输出方式：

- 抓取任务放到对应 skill 的 `output/`、`runs/` 或自定义项目目录中。
- Excel 输出要保留清洗明细和总结，不只输出最终结论。
- 所有分析型 workbook 建议至少包含：
  - 原始来源信息
  - 清洗后的有效反馈
  - 标签或分类链
  - 汇总统计
  - 可复核的代表性原文

## 使用注意事项

- Amazon 抓取可能需要人工登录，建议设置 `HEADLESS=false`。
- 评论数量预估只代表搜索页或商品页可见信号，最终有效评论数以实际抓取结果为准。
- ASR 图片提取依赖本地 Chrome 缓存；运行前最好先用 Chrome 打开相关图片或线程。
- 对多产品、多品类 Excel，不要默认混合分析，应先明确目标产品或分析范围。
- 自动标签只是初筛，关键结论需要人工复核。
- 清理专利附图编号时应先审核数字候选，不要把电压、尺寸、公差或功能轴号当作引用编号批量删除。
- 评论和图片数据可能涉及平台规则、账号权限和隐私边界，使用时应遵守目标网站条款和内部数据合规要求。

## 推荐工作流

### 竞品评论研究

1. 用 `amazon-review-scraping-skill` 通过关键词生成候选商品。
2. 在预检结果中排除无关场景。
3. 回复 `开始执行` 后抓取评论和图片。
4. 将导出的评论 Excel 输入 `product-definition-voc`。
5. 输出隐藏需求、Aha moment、场景卡片和产品机会。

### HIFI 售后或退货归因

1. 准备目标产品的评论、退货或售后 Excel。
2. 用 `hifi-comment-tagging` profile workbook。
3. 聚焦一个产品，清洗有效反馈。
4. 复用历史人工标签或按标准分类链打标。
5. 生成 summary workbook，用于产品复盘和问题优先级判断。

### 音频论坛洞察

1. 收集 ASR 相关 thread URL。
2. 写入 `asr-review-scraping-skill/runs/<project>/curated_threads.txt`。
3. 运行 ASR pipeline。
4. 得到帖子、图片和打标准备表。
5. 结合 `product-definition-voc` 或人工标签体系继续分析。

## 维护建议

- 每个 skill 的详细规则以对应目录下的 `SKILL.md` 为准。
- `references/` 存放输入契约、标签体系、输出布局和总结模板，不建议随意删除。
- `scripts/` 中的脚本是主要执行入口，优先复用脚本而不是手工改 Excel。
- 新增技能包时建议保持相同结构：

```text
new-skill/
├── SKILL.md
├── README.md
├── references/
└── scripts/
```

## License

各子项目 `package.json` 当前标注为 MIT。若后续加入第三方数据、模型输出或平台采集结果，请根据实际来源补充更细的许可和合规说明。
