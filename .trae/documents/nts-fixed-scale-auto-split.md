# NTS 固定 1:1 输出 + 溢出自动分页方案

## Context

不按比例(NTS)模式当前在 PDF 输出时无条件缩放:所有图都经 `Projector.fit()`(render.py:244-258, 0.82 边距)缩进图框,链式布局超档位时还整体兜底缩 `schematic_scale`(~80 条 "chain layout scaled" 警告),符号随之变小,风格不一致。用户定稿:

- NTS 模式 PDF **固定 1:1**(图面 mm = 布局 mm),不再 fit、不再兜底缩
- 内容超图框或材料表满 → **自动分页**
- **单 PDF 多页**;每页材料表只列本页物料(项号每页 1..N 重排,球标联动)
- 切点必须落在**直管段内部**,不切穿管件/符号
- GUI 预览**保持现状**(整段内容 auto-fit 矢量回放,不分页)
- 按比例模式**完全不动**

## 方案总览(两遍架构)

```
render_pdf (NTS)
  ├─ 测量遍: absolute 链式布局(scale=1) → 组件页bbox + BOM行数
  ├─ sheetpack.plan(): 连续区间 DP → 页数、每页组件集、异常回退直管切点
  ├─ 切管: splitmerge._split_pipe 生成页子文档 + 合成接续声明
  └─ 渲染遍: for page in pages: add_page() → _render_document(subdoc, sheet_no, sheet_total)
```

## 改动清单

### 1. 新建 `pcf/sheetpack.py`(纯数据分页,不 import render,依赖注入防环)

- `plan(info, usable, capacity_rows, show_cut_table, conn_tol)` 返回页计划、切点、组件/管段归页和警告;`PagePlan` 记录本页组件、被切管段与页间断口
- 测量输入:render.py `_measure_chain()` 以 absolute 链式布局生成 `SheetInfo`,按文档序聚合组件端点/中心点与图面盒
- **组件 bbox 必须加余量**:非 PIPE 件向外扩 `_SCHEMATIC_LEG_MM`(≈4.5mm,手轮/法兰杠/支管台);标注带按现有 fit() margin=0.82 反推 → 可用窗 = draw_region × 0.82
- 分页优先级:①每页主管线为链上连续段;②页联合盒不超可用窗;③最少可行页数下用 DP 最小化各页零件数方差。主链取生成树双 BFS 长链,小支路跟随主链节点,超窗长支路拆成独立连续链;支架等独立小件跟随最近管线节点并计入件数
- BOM 行数超容量仅警告,绘制侧截断兜底,不为凑表容量破坏管线连续性

### 2. 切点规则(sheetpack 内)

- 链式方案按整件连续区间分页,不切管;只有单件超窗等整件划分不可行时才回退窗口网格切片
- 网格回退中,跨格长直管在穿越格线处切开;不足 18mm 的碎段并入邻段,整根短管不切,改走组件接续
- 连接簇整体归页优先;装不下的大簇保持盒中心归格,由切管与页间接续处理
- 切管复用 [splitmerge.py](file:///c:/Users/jiazheng/桌面/PCF_Python/pcf/splitmerge.py) `_split_pipe`/`_build_parts`,公开入口 `split_pages(doc, cuts, page_of)` 按计划归页,不走 connected_sets 重判

### 3. render.py 多页编排

- `render_pdf`(1163)NTS 分支重构:先 `sheetpack.plan()`(内部自建一次性 Projector 跑 absolute 布局,含 pipe_scale 档位迭代;**删除 floored/schematic_scale 兜底分支** L1281-1299),得页计划后循环 `pdf.add_page()` + `_render_document(subdoc, ..., sheet_no=i, sheet_total=m)`
- **每页 Projector:scale=1,不调 fit();offset = draw_region 中心 − 该页 packed bbox 中心**(修正:链式布局居中是整链级的,多页必须逐页平移内容窗)
- 接续合成:切点两侧各注入 Continuation(point=cut, sheets=[兄弟页]) 到 options.continuations;`Continuation` 加 `sheet_tag: str|None` 字段,`_cont_callout_lines`(1878)优先用它打印 "SHT n"(旁路 `_sheet_tag` 文件名正则);continuations.py 声明解析链不动
- 对侧幻影件:`_draw_phantom_mate` 以兄弟页子文档为 neighbour 画半段虚线,沿用现有虚线线宽规则
- 页间切点**不做双写 EXCLUDE**:两半是真材各进本页 BOM/切割清单(长度各算各),视觉延续由幻影承担

### 4. BOM / 球标 / 标题栏

- 页子文档传入 `_render_document` 后,`_build_bom`(5280)天然按页组件子集构建,项号每页 1..N;`_draw_callouts` 经 `_bom_key` 查行自动联动,无需改
- `_draw_title_block`(5733)fields 加 `("SHEET", f"{n}/{m}")`(仅 m>1 时显示)

### 5. 结果合并(pcf_renderer.py)

- RenderResult 按页合并:welds/dimensions/bom/callouts/continuations 拼接,`bom_truncated` 取 OR,warnings 加 `[SHT n]` 前缀
- 侧车 CSV 加 `sheet` 列;焊点号按页前缀 `p{n}-`

### 6. 预览(不改)

`render_preview_scene` / `render_scene` / GUI `_preview_zoom` 全部不动——预览继续整段内容 auto-fit 显示,与 PDF 分页天然解耦。

## 回归与验证

1. 单图验证:选旧日志中 "chain layout scaled" 幅度最大的 3-5 张,核对:每页主管线连续、页联合盒不超窗、零件无遗漏/重复、件数尽量均匀;每页 BOM 项号 1..N 与球标一致、标题栏 SHEET n/m;网格回退时切点必须位于直管内部
2. 不超框的 NTS 图:输出应与现在一致(仍单页,只是不再兜底缩)
3. 按比例模式图:逐字节对比应零变化
4. GUI 预览:超量图仍整段显示、auto-fit 不分页
5. 366 张全量回归:预期 ~80 条 "chain layout scaled" 消失(其中触底兜底图变多页),新增分页信息行;警告基线更新进 skill
6. exe 打包 + git 提交(中文一行式)+ 更新 SKILL.md 基线
