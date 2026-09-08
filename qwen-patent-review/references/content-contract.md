# 说明图内容契约 v1

CLI：`scripts/compose_review_sheets.py MANIFEST --font FONT -o EMPTY_DIR [--preview]`。
兼容 Python 3.9 与兄弟 CAD 技能的锁定依赖。只接受其正式 DXF 的
`GEOM / HIDDEN / LEADER / NUM` 图层；其他图层不复制。含未支持实体会失败，不静默漏图。
字体需已经安装并提供实际文件路径；不捆绑客户或商业字体。

`examples/review-sheets.json` 是完整合成内容示例。所有路径相对 manifest 所在目录解析。
生成合成输入的命令见同目录 `examples/make_demo.py --help`。

| 字段 | 含义 |
|---|---|
| schema | 固定为 `qwen-patent-review/1` |
| reference_numerals | 原 CAD CLI 生成的 `patent-numerals/1` 全局标记表 |
| parts[].id | 此内容文件内稳定的条目 id |
| parts[].selector | 精确引用标记表 selector，名称和编号从该表读取 |
| parts[].instance_ids | 原模型实例路径/稳定索引列表，用量由长度计算；不能重复计入不同条目 |
| parts[].function / assembly | 分别解释作用、安装/连接；都不能为空 |
| parts[].unlabelled_reason | 已编号零件在此套源图中未标引线时，必须给出原因 |
| sheets[].id / title | 安全文件名、说明标题；图号按数组顺序自动签发 |
| sheets[].source_dxf | 原 CLI 的已通过 QA 的图纸；不要输入旧版拼接说明图 |
| sheets[].overview | 图下一句概述 |
| sheets[].paragraphs | 右上说明段落列表 |
| sheets[].rows | 本页表格引用的 parts id，顺序即表格顺序 |

功能、装配、概述及段落均用 `{"text":"…","status":"…","evidence":["…"]}`。
状态有三种：`observed`（模型/图纸中观察到）、`confirmed`（有工程资料或明确确认）、
`pending`（具体待核实事项）。前两种需要非空 evidence 数组，例如工作表行号、
几何测量报告位置、用户确认记录。脚本只检查证据字段是否存在，不自动判断证据真实与充分。
原始 evidence 和输入文件随内部交付保存，不直接塞进主图或公开教程。

同名不同实例无法独立标号时，省略 selector，填写 `name`、`unresolved_reason` 和独立实例列表。
脚本生成“—”，而不是编一个号码；最终报告列出所有这类未编号条目。
instances 由输入声明，不能仅凭生成报告中的 `declared_instances` 宣称 STEP 全件已覆盖。

生成报告包含源 DXF / 字体 / 内容 / 标记表 / 输出文件 SHA256 和每页几何变换。
主图及原引线、编号共同做一次等比缩放和平移，几何不重建。
内部说明图编号最小 2.5 mm 是本模板的可读性下限，**不是专利附图法规阈值**；
原 CAD 的 QA 仍单独执行。正文与表格不根据内容长度自动缩字，空间不足时报错要求拆页。

DXF 为毫米模型空间，A3 横版。重开保存后的 DXF 生成 PNG/PDF；
Matplotlib 后端完成排版后重新固定纸张尺寸。PDF 供阅读，DXF/DWG 是可编辑交付源。
DWG 另调用 `../patent-drawing-dwg-cleanup/scripts/autocad_core_dxf_to_dwg.py` 并保存审计输出。

退出 0 = 此次生成成功；1 = 输入/排版/读写失败；2 = 命令参数错误。
只允许空输出目录，避免部分重跑把旧文件误作新交付；失败后用新目录重试并记录失败原因。
`sheet-report.json` 的状态保持 `generated_requires_review`，视觉与工程检查默认 pending；
人工检查结果另存验收记录，不能由生成程序自签“验收通过”。
