---
name: qwen-patent-review
description: Run a Qwen Office hardware drawing workflow from GitHub and STEP/BOM through Web structure review to editable CAD explanation sheets with component functions, assembly notes, Chinese parts tables, previews, and an evidence-based handoff. Use for 千问办公专利图拆解、说明图、教程或流程复用; enriched sheets are internal review material.
---

# 千问办公：从硬件模型到可读说明图

冻结版本：**qwen-review-v1.0.0**。本目录与同仓库的 `patent-drawing-dwg-cleanup`
配合使用；安装时保留这两个相邻目录。读取 [流程与千问提示词](references/workflow.md)。
本次已执行检查及复跑命令见 [冻结验证记录](references/freeze-validation.md)。

交付一页能读懂的结构说明：**左上 CAD 主图、右上功能/装配说明、图下一句概述、
左下“序号 / 名称 / 用量 / 备注”表、图号**。每个表格条目解释作用和连接方式，
证据不足时写清具体待核实内容。整机装配、整机爆炸、局部装配、局部爆炸是不同交付项。
用户要求整机爆炸图时，不能用若干分区爆炸图替代；做不出就明确列为未完成。

## 执行边界

- 千问办公负责拉取指定 GitHub 版本、运行现有 CAD CLI、整理说明和最终交付。
  使用 Computer Use 时先观察真实 UI；中文长 Prompt 用粘贴，发送后核对实际文本与任务状态。
  没有千问访问能力时如实报告，不用本地执行冒充千问执行。
- CAD 生成阶段按兄弟技能 Route A 运行；不修改原 STEP，不改变 CAD 内核、原有 QA 或编号规则。
  本技能是新增的**正式说明图后处理入口**：模型可编写 `review-sheets.json` 中的内容与证据，
  由 `scripts/compose_review_sheets.py` 排版，代替每个项目重写临时绘图脚本。
  原 Route A 的“仅写 plan”约束仍适用于原始 CAD 生成，不限制这个明确选用的后处理阶段。
- Web 是给用户复核结构的步骤。展示实际选中的图、装配态、分解态与细节，并记录用户反馈。
  不代替用户点击确认，不把“页面已打开”记为通过。用户此前已对同一版本确认时复用该记录。
- 编号和名称读取原始 `reference-numerals.json`；同名不同几何不得合并为一种零件。
  当前原始 CLI 不能按实例分别标号，这类条目用独立实例 id、名称、未编号原因记录，
  图表中用“—”，在交付中说明限制，不能宣称逐件标号全部完成。
- 图示位置、候选功能、工程确认分开表达；不能从爆炸排布推断装配工序，不能从轴孔名义尺寸
  宣称过盈/间隙配合，也不能把参考案例的角度、齿数、传动方式复制到新模型。
- 文件名带 `_engineering` 的说明图用于内部评审/教程素材。纯附图仍由原始 CLI 单独生成。
  宣发稿只能描述有记录的执行与交付；不要宣称自动完成工程确认或专利审校。

## 生成说明图

环境沿用兄弟技能锁定依赖；另需一款已安装、支持中文的 OTF/TTF 字体。

```bash
python3 scripts/compose_review_sheets.py /path/to/review-sheets.json \
  --font /path/to/NotoSansCJKsc-Regular.otf -o /path/to/new-output --preview
```

输入见 [内容契约](references/content-contract.md) 与 [合成示例](examples/review-sheets.json)。
脚本计算排版、动态表格行高和用量；不接受项目自填的坐标/字号。
若文字挤占主图、编号过小、图表编号不一致，则报错并要求拆页或修正内容。
生成后的 DXF 是 PNG/PDF 的共同来源，PDF 保持实际 A3 横版尺寸。

输出：每页 `_engineering.dxf/.png/.pdf` 和 `sheet-report.json`。
DWG 使用兄弟技能的 AutoCAD 转换 CLI，并保存 AUDIT 记录。
检查通过退出 0 只代表生成成功，**不代表视觉、工程或用户验收通过**。

## 交付前

按 [验收记录模板](references/acceptance-template.md) 留痕：源文件与计划版本、实例覆盖、
机器 QA、逐页目检、DWG 审计、用户 Web 反馈、未解决事项和最终文件校验。
不能把包围盒占比当作零件可辨识度；必须打开最终 PNG/PDF，而不是仅检查文件存在。
把实际结果交回同一个千问任务，再完成教程/宣发文档。

真实模型、BOM、价格、客户图纸、原始截图、会话令牌与本机路径留在项目工作区。
更新公共 GitHub 时仅同步通用流程、工具和合成测试件。
