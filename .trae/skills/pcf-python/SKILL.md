---
name: pcf-python
description: PCF_Python 项目速查——模块地图、366 张全量渲染回归、ALIAS.asc 符号库同步、PyInstaller 打包 exe、中文一行式提交规范。Use when 用户要求回归渲染、同步符号库、打包 exe、提交代码或了解本项目结构。Do not use for 与本项目无关的通用 Python 问题。
---

# PCF_Python 项目速查

PCF → 单线轴测图（ISO）渲染器：`pcf/` 渲染引擎 + ttkbootstrap GUI + PyInstaller 单文件 exe。

## 模块地图

- `pcf/render.py` — 渲染引擎（符号调度、接图标注、尺寸、BOM/切割表、预览 SceneModel 双后端）
- `pcf/sheetpack.py` — NTS 1:1 自动分页规划：链式连续区间 DP + 长支路分解 + 均匀分配；网格切片/切管作异常回退（见「NTS 自动分页」节）
- `pcf/splitmerge.py` — 分页执行：按 page_of 拆文档、短管跨格接续、_orig_idx 打标继承
- `pcf/continuations.py` — 接图点判定：**纯声明驱动**，每条 END-CONNECTION-PIPELINE = 一个接图点（CO-ORDS=坐标、CONNECTION-REFERENCE=名字），无几何匹配
- `pcf/asc_symbols.py` + `symbols/ALIAS.asc` — Isogen .ASC 符号库单源（315 符号，先到先得）+ 阀体→阀柄 skey 厂商配对
- `pcf_renderer.py` — CLI 批量渲染与 GUI 预览适配层（render_set 按目录缓存整套文档）
- `pcf_app.py` — GUI（管线清单勾选出图、矢量预览）
- `PCF收集器.spec` + `version_info.py` — PyInstaller 打包配置与版本信息

## 全量渲染回归（改动渲染链后必跑）

```powershell
py -3 -X utf8 pcf_renderer.py "c:\Users\jiazheng\桌面\新建文件夹\HQSH_Iso_Pipeline_PTFE" --out-dir preview
```

- 366 张（当前共 502 页：279 张单页、87 张多页），判定通过 = exit 0 且日志无 Traceback/Error
- 警告基线（全为既有类别，新增类别须解释）：87 continuation（声明指向本交付集外线号）、19 不连通管段组、1 尺寸过小（SHT 4 的 FLG FACE → LINE END 22）；固定 1:1 分页后无切割表截断、无 chain layout scaled
- 测试语料 `c:\Users\jiazheng\桌面\新建文件夹\HQSH_Iso_Pipeline_PTFE` 为**只读**厂商数据，勿改

## 符号库同步（vendor 更新 BSF.asc 后）

1. vendor 源：`d:\Users\jiazheng100002\Desktop\新建文件夹 (2)\BSF.asc`（勿动源文件）
2. 复制为 `symbols/ALIAS.asc`；`_prev_keys.json` 存符号全集基线，同步后对差集即新增清单，更新基线
3. 跑全量回归确认零异常

## 打包 exe

```powershell
py -3 -m PyInstaller "PCF收集器.spec" --noconfirm
```

- onefile → `dist/PCF收集器.exe`；symbols/ 目录、ttkbootstrap 全量已随 spec 打入；pymupdf/fitz 显式排除
- 版本号在 `version_info.py`（当前 1.0.0.0，升级时同步四处 filevers/prodvers/FileVersion/ProductVersion）
- 打包前先跑全量回归

## 提交规范

- 中文一行式 commit message，聚焦"为什么"；只 `git add` 具体文件，禁止 `git add -A`
- `.gitignore` 已覆盖 preview/、dist/、build/、__pycache__/、*.pdf、.screenshots/——构建产物与回归输出不入库
- 深层历史规则（符号语义、渲染细节、用户定稿决策）原在 Trae 记忆 `project_memory.md`；2026-11-14 起代码注释已精简，完整定稿规则改记下方「定稿规则总账」，project_memory 只留当日进展

## 定稿规则总账（2026-11-14 自代码注释迁出）

### 管架与符号库
- 管架选线：优先世界距离≈0（真立管线上）的管段再按页面距离——39021 管架曾对到邻跨接续短管画成斜线；候选集含弯头/三通腿
- 管架 SKEY 语义出自 SystemMaterialData.xml Supports 段：01HG/空=承重、ANCH/GUID/SPRG/HANG；.ASC 方向模式（501 尾参）：0=顺管、2=UP/DOWN 翻转、3=三轴六向
- 符号库优先（START→eps[0]、EXIT→eps[1] 双点仿射），手绘仅回退；SAFETY-DISC 归 GASKET 但必须画官方符号，仅普通垫片不画并父标；WT 支管台=主管中心→支管端对齐
- 法兰常误导出为 COUPLING，派发兼看 description，BOM 只用 classify()
- 阀柄独立 skey（01SP/03SP/04SP/10SP）已归一 START=(0,0)，+X 旋到 SPINDLE-DIRECTION 投影、沿用阀体 sy

### 宽度体系
- 按比例图：孔径正比 2·_bore_half_w，法兰类×B16.5 突面比（1/2"≈1.64→24"≈1.14，C150 全表 21 档最小二乘拟合偏差≤0.05≈0.15mm；锚点 2"→1.53、6"→1.28、12"→1.18）
- 不按比例图：固定尺寸（法兰杠 3.8/阀 1.9/符号 4.5）；高度 4.5 同时校准轴向自然长度（VB** 800→9.0mm、FLWN 500→4.09mm），改高度连带缩长度、链上留死区，宽度走单独通道

### END-POSITION 与接图点
- END-POSITION-OPEN/CLOSED 是文件声明的止点，置信高于连通图行走（模型内 field weld 两端容差内重合，否则 spool 边界无 landmark）；非盲信，经 _is_run_terminus()；END-CONNECTION-* 为跨管线接续跳过；接图点由 continuation 阶段画（图上要的是下一张图名）

### 偏心变径
- 不建弦 leg；轴向取邻接管段（大端优先小端兜底），绝不沿斜弦标注；warp 用区间保护（点对规则在有中间坐标时把偏心撑成折线撕掉小端）
- 符号体两端落自身焊点（旧法沿邻轴伸出致楔尖悬空），run 只定方正（弦定方出菱形）；跨向用投影世界轴（屏幕垂线不在等轴测面）
- 平侧示意画（真实 1" 偏心仅 1/3mm 读作同心），平侧朝向取模型；尺寸以大端为锚、小端投影回原管轴；体深≥页空间中心线偏移否则领结交叉

### 变径就地标签
- 实际造成管径变化的元件就地标 "大X小"（降序，如 200X50），覆盖三类：REDUCER（自身两 END-POINT bore）、变径三通（END-POINT 主管 + BRANCH1/2 支管，逐腿标）、OLET/TEE-STUB（支管 bore + 宿主管口径——宿主管经几何匹配在 _olet_header_points 带回，找不到宿主管跳过）
- 端口 bore 全相等、缺径或不足两口征不标；按比例模式（proj.proportional）与 cont_copy 幻影虚线件不标
- _draw_bore_tag：文字方向顺主管（REDUCER 沿自身轴；TEE/OLET 沿主管向——OLET 用几何匹配 run 的页面向 header，找不到宿主管回退支管向），翻向保持正立、跨向 _iso_cross_dir 抬 3mm 避开本体线、字号 5.0；**旋转角 atan2(-u.y, u.x)**（页面 y 向下，同 F.W. 标签约定，取正 y 会镜像斜率）；锚：REDUCER 取本体中点，TEE 标在中心点（多支腿 CROSS 逐腿抬 2.5mm 防叠字），OLET 在支管端外 1.2mm；三处绘制入口（REDUCER 符号优先路径/TEE 腿循环/OLET 符号前 + REDUCER 手绘回退）各自接入

### 尺寸
- 尺寸线沿投影世界轴；≥4 箭头长才双箭头内置；offset leg 的 spanned 按页空间判定（模型对角线判定使 3" roll 在 40' 线中 offset 盒消失）
- 兜底尺寸按连通组独立跑（旧版整页无尺寸才触发+多组拒绝）；垫片只存于法兰连接侧，此结构证据优先 rank 表判焊端

### 链式布局
- 直管按真实长度单调分档（基准×1.5 定稿）；按边弹性缩放只压管线边（全局乘 s 会压符号端点对、符号画小），rigid_symbols=False 极端兜底全体弹性
- 垫片模型厚度 ~1.5mm 示意不可见，warp 固定 0.6mm 间隙；face-to-face≈0 时轴向取最近管段且方向指"远端"；贴附件吸附必须在打包后（rebuild 弹性拉伸会撕掉）

### NTS 自动分页
- 分页优先级定稿：①每页主管线必须是链上连续的一段；②页联合盒不得超过可用窗口；③在最少可行页数下用 DP 最小化各页零件数方差，尽量均匀。主链取生成树双 BFS 长链，小支路整体跟随主链节点；支路自身超窗时拆成独立连续链，主链两侧分别成段。支架等不与管端共点的独立小件跟随最近管线节点并计入件数，不单独成页
- 分页规划参数真实来源 = render.py：bom_w=0.24·页宽（show_bom）、title_h=0.16·页高，draw_region 再过 _usable_window（0.82 居中收缩）得可用窗——调试脚本用错窗口会导致归格结果与渲染不一致
- 网格切片（仅链式整件划分不可行时回退）：组件按图面盒中心归格（一格一页，行优先）；跨格长直管在穿越格线处切开；切点 <18mm 的碎段并入邻段（删分界），整根短管（全长不足 18mm 装箱）不切、走组件接续（outs pipe_idx=-1）
- 连接簇整体归页（网格回退路径）：实际相接的组件簇（容差桶+3×3×3 邻域聚类，簇=共享一个连接点的组件集合，非传递闭包）联合盒装得进一格时整簇归联合盒中心所在格；装不下的大簇保持盒中心归格由切管+接续处理。一件组件常属多个簇（弯头/直管两端各接一串），多簇争抢按"移动成本"（需改变盒中心归属的成员数）升序归页且归页即锁定，后续簇 target 与锁定成员冲突整簇跳过——无锁定时后处理簇会把已归页成员拉走（BR3R-30001 立管盒中心距行界 0.15mm 落错格；弯头 68 被含长管 70 的簇以 0.13mm 擦边 target 拉回邻页）
- junction 检测：同簇成员 presence 集不相交的对两侧页各挂接续（pipe_idx=-1），dedupe key=(cell_me, cell_ot, round4 断点)——同点接多件时文字相同，重复挂只叠盒子
- 幻影对侧断点（_draw_phantom_mate）用 σ 局部压缩比锚定：σ=|page(inboard)-at|/|raw(inboard)-base_raw|，幻影所有点 at+(raw-base_raw)·σ，绕过 offset_nearest 回退——否则接图线歪斜
- junction tag 无 CONNECTION-REFERENCE 时冠 "SEE ISO"；跨页断开后尺寸链穿过管件延伸到对侧，尺寸值变化属预期（如 1328 变 1433=底到总管中心线，1328 仍在切割表作下料长）

### welds
- 联管节 gap 模型空间测量（页空间被 warp 拉大，1/16" 越过 3mm 阈值致本体线压焊点）；焊耦合经焊点派生但中心线须连续
- socket weld 画方括号朝插入件（旧版统一圆点不可辨）；boss 站开根部 0.9mm 焊点；tie band 用 1e-3 容差（1e-6 时螺栓法兰对两半因投影噪声选不同轴、两面杠 60° 成 X）

### 球标与 GUI
- callout 单一风格，裸标签靠邻近推断在繁忙图纸/NTS 拉近时失效；leader foot 与邻件锚/焊点留白；切割件标注点中点优先、四分点逃生；垫片/螺栓并入任一端法兰，无法兰单独球标；标题栏标签比值低一级（原有层级），仅值用 _SHEET_FONT 8.0
- revision join 容差——未动焊点世界坐标相同（distance 0），容差只消化模型微调；pull-back 容差——地图项自放置，未动 marker 与行差 <1px

### __init__ 本地修改
- BOM 跳无 ITEM-CODE 焊缝；项号/气球固定行序号 1..N（原料号存 code 列）；直管气球标 BOM 项号
- SUPPORT/FLOW-ARROW 优先 .ASC 符号库（SUPPORT-DIRECTION/FLOW 控向，无库回退内置，箭头不进 BOM）；asc_symbols.py 为本地新增解析器（见 symbols/ASC符号格式规则.md）
