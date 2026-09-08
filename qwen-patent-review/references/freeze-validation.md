# 冻结验证记录 · qwen-review-v1.0.0

日期：2026-09-08。基线 CAD 工具 commit：`0602acf`。
验证环境：macOS、Python 3.9.6、OCP 7.7.2、ezdxf 1.4.2、Matplotlib 3.9.4；
说明图使用已安装的 Noto Sans CJK SC Regular。

## 已执行的检查

| 检查 | 实际结果 |
|---|---|
| Skill 格式与引用入口 | quick_validate 通过；README 提供流程、内容契约与验收入口 |
| 新增说明图测试 | 14 项通过，无跳过 |
| 负例拦截 | 空说明、缺来源、实例重复、selector 重复、未知表格行、危险文件名、自填布局/数量被拦截 |
| 几何保留 | 源线条与引线经逆变换后坐标/半径一致，源编号保持不变 |
| 中文与长表 | 按实际字体测量；名称和备注都换行；超长内容挤占主图时报错 |
| PDF | 回读 MediaBox 为 A3 横版，约 1190.55 × 841.89 pt；从 PDF 重新渲染目检正常 |
| 原 CAD 验收套件 | 14 / 14 节点通过，0 跳过，包含 152 项单元测试 |
| 原 CAD 基准 | 两张合成图与仓库金样逐位一致；坏引线的负例仍判失败 |
| STEP → 说明图衔接 | 合成 STEP 经 analyze / validate / render 生成局部分解 DXF，再进入新说明图 CLI，成功 |
| DWG | 新说明图经 AutoCAD Core 转换，AUDIT 错误 0、修复 0；87 个实体，非连续线型 0 |
| 新增文件发布检查 | 通用代码、说明与合成输入；无客户 CAD、BOM、图纸、私有件号或本机绝对路径 |

## 复跑命令

在仓库根目录运行（将字体路径换成当前机器上已安装的中文字体）：

```bash
python3 patent-drawing-dwg-cleanup/acceptance/run_acceptance.py
REVIEW_CJK_FONT=/path/to/NotoSansCJKsc-Regular.otf python3 -m pytest qwen-patent-review/tests -q
python3 qwen-patent-review/examples/make_demo.py -o /tmp/qwen-review-demo
python3 qwen-patent-review/scripts/compose_review_sheets.py \
  /tmp/qwen-review-demo/review-sheets.json --font /path/to/NotoSansCJKsc-Regular.otf \
  -o /tmp/qwen-review-output --preview
python3 patent-drawing-dwg-cleanup/scripts/autocad_core_dxf_to_dwg.py \
  /tmp/qwen-review-output/demo_engineering.dxf /tmp/qwen-review-output/demo_engineering.dwg
```

输出目录需新建；重复验证用新的临时目录。未设置 `REVIEW_CJK_FONT` 时，
字体/渲染相关测试会跳过，不能把跳过写成验证通过。

## 验证范围

千问的实际案例验证支持了流程与版式选择；这次通用代码及回归检查由 Codex 在本地执行。
没有声称千问已重新从此 tag 拉取并运行，也没有声称跨机器、全自动实例识别或工程事实审校已通过。
真实案例的内部材料和过程截图保留在客户工作区，不随公开仓库发布。
该版本固定的是可复用的执行流程、内容格式和工具实现，项目各自仍需 Web 反馈与逐页验收。
