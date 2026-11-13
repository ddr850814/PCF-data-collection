# 全矢量预览重构：渲染引擎双后端（fpdf2 出图 + tk.Canvas 预览）

## Context

当前预览链路是"fpdf2 画图 → 临时 PDF → PyMuPDF 栅格化 → PIL 位图 → Canvas 显示"。问题：
1. 用户后续要在图上做元素级功能（改焊点标号、加分图点拆图），位图无法定位元素；
2. 工程图应矢量显示，缩放不失真；
3. PyMuPDF 自带 30MB C 引擎，只为预览一个中转站，拖累 exe 体积（61.9MB）。

目标：渲染引擎抽象出绘图后端接口，出图仍走 fpdf2；预览新增"场景收集后端"产出**矢量场景模型**，GUI 回放到 tk.Canvas（全矢量、元素带语义标签、缩放平移无损）。最终卸掉 pymupdf 依赖。

## 探查结论（设计依据）

- fpdf2 调用面收敛：line×50、text×42、set_font×19、get_string_width×15（**布局决策依赖它**：尺寸文字放线内/外 render.py:3244、标签盒、`_fit_text`:4252）、rect×9、local_context×10、ellipse×2、dash×1、rotation×2、transform(斜切)×1、polygon×1、image×1。
- `_Canvas`（render.py:761-1018）已是薄封装 + 碰撞跟踪，约 40 个 `_draw_*` 函数大多走它；仅 7 处直穿 `canvas.pdf.*`（1394-1444, 1530-1533, 1950-1953, 2541, 2842, 3229, 3943-3946）和 `_draw_title_list`（4324，收裸 pdf）需要收编。
- 最深耦合：`iso_text`（823-862）旋转+斜切 Transform（等轴测文字）。
- fitz/PyMuPDF 全项目仅 pcf_app.py:32 一处 import；`database_helper/pcf_data_collector/pcf_parser` 与渲染无关，不动。
- 预览视图状态机（fit/zoom/pan/防抖，pcf_app.py:618-693）逻辑可整体保留，只换绘制载体。

## 方案

### 1. 绘图后端协议（pcf/render.py 新增，~100 行）

```python
class PlotBackend(Protocol):
    def set_font(self, size, bold=False): ...
    def text_width(self, s, size, bold=False) -> float: ...   # mm
    def set_stroke(self, color=(0,0,0), width=0.4): ...
    def set_fill(self, color): ...
    def line(self, a, b): ...
    def rect(self, x, y, w, h, style="D"): ...               # D/F/DF
    def ellipse(self, x, y, w, h, style="D"): ...
    def polygon(self, pts, style="F"): ...
    def dashed_line(self, a, b, dash, gap): ...
    def save(self) / restore(self): ...                      # 状态栈
    def rotated_text(self, x, y, s, size, angle, centered): ...
    def matrix_text(self, x, y, s, size, matrix, knockout_box): ...  # iso_text 斜切
    def image(self, path, x, y, w, h): ...
```

- `_FpdfBackend`：包住现有 `_Pdf`，方法体就是现在的调用原样搬家。
- `_SceneBackend`：不做绘制，把每次调用收集为 `SceneItem`（见下），颜色/字体等当前状态记进条目。`text_width` 用模块级隐藏 `FPDF(unit="mm")` 单例量宽（`_metrics_fpdf`），保证**场景回放的布局与 PDF 出图逐字节同版式**。

### 2. 场景模型（pcf/render.py 新增）

```python
@dataclass
class SceneItem:
    kind: str            # line|rect|ellipse|polygon|dashed|text|image
    pts: list[float]     # 页 mm 坐标（平铺）
    text: str = ""
    size: float = 7.0    # 字号(pt)
    angle: float = 0.0   # 旋转
    matrix: tuple|None = None   # 斜切矩阵 [a,b,c,d]（iso_text 用）
    color: tuple = (0,0,0)
    width: float = 0.4   # 线宽 mm
    fill: tuple|None = None
    tags: tuple = ()     # 语义标签，如 ("weld","3")、("pipe","2105-...")

@dataclass
class SceneModel:
    page_w_mm: float; page_h_mm: float
    items: list[SceneItem]
    ink_bbox_mm: tuple | None   # 全部墨迹包围盒（预览裁剪用，替代 fitz get_drawings）
```

新增 `render_scene(pcf_path, options) -> SceneModel`：与 `render_pdf` 共用同一套布局/绘制管线（`_Canvas` 逻辑不变），仅 backend 换 `_SceneBackend`。焊点记录逻辑（WeldPosition 坐标）原样保留。

### 3. `_Canvas` 收编直穿调用（render.py 改造点）

- 7 处 `canvas.pdf.*` → 新增封装方法：`white_box(x,y,w,h)`（白底擦除，859/1395/2843/3944）、`push()/pop()`（local_context，786/858/1441/2541）、`measure(text)`、`rotated_text`、`matrix_text`。
- `_draw_title_list` 参数裸 pdf → 改收 canvas。
- `text()`/`iso_text()` 内部改调 backend 对应方法（语义不变）。

### 4. 语义标签埋点（为后续功能铺路）

- 焊点 dot 绘制处 → `("weld", str(number))`
- 管线中心线段 → `("pipe", pipeline)`
- 元件符号 → `("component", uci 或序号)`
- 标注盒/续接块 → `("callout", item)` / `("cont", ref)`
本次只埋这四类，后续功能用到再加。

### 5. GUI 预览重写（pcf_app.py）

- `_start_live_preview` → 后台线程调 `render_preview_scene(...)`（pcf_renderer.py 新增薄封装，替代 render_preview_bytes）拿 SceneModel，token 防过期保留。
- `_show_preview_bytes`/`_show_preview` → `_show_scene(scene)`：缓存 scene，重置视图，调 `_draw_scene()`。
- `_draw_preview_img` → `_draw_scene()`：视图状态机原样（zoom None=fit 居中；Ctrl 滚轮 1.15 锚点缩放；拖动平移；双击复位；120ms 防抖）。绘制改为遍历 `scene.items`：
  - line/dashed → `create_line(..., dash=(6,4))`；rect→`create_rectangle`；ellipse→`create_oval`；polygon→`create_polygon`
  - text：普通→`create_text`（字号 pt→px 随 k 缩放）；rotated→逐字符沿角度排布；matrix（斜切）→逐字符沿矩阵映射的两轴排布（**字形不斜的近似**，与 PDF 出图有细微视觉差，工程可读性等同——计划内明确此折中）
  - mm→px：`px = x * k + ox`，坐标换算在绘制时做
- 每个条目 `tags=("elem", kind, id)` 绑语义，预留点击交互。
- 性能预案：单图 items 数量级几千，Canvas 一次重画预计 <100ms；若卡，按 ink_bbox 做视口裁剪（只画相交条目）。
- 删除 `import fitz`、PIL 预览分支（`HAS_PREVIEW` 逻辑改为总是可用）。

### 6. 依赖与打包

- requirements.txt 去掉 pymupdf；PCF收集器.spec `excludes=['pymupdf','fitz','PIL.ImageTk']`（PIL 主体留给 ttkbootstrap/fpdf2，只排 ImageTk）——exe 预计瘦 ~25-30MB。
- fpdf2 保留（出图+字宽度量）。

## 实施步骤

1. render.py：加 `PlotBackend` 协议 + `_SceneBackend` + `SceneItem/SceneModel` + `_metrics_fpdf`。
2. render.py：`_Canvas` 新增封装方法，收编 7 处直穿调用与 `_draw_title_list`。
3. render.py：`render_pdf` 内 backend 化（`_FpdfBackend` 包 `_Pdf`），新增 `render_scene()`；埋四类语义标签。
4. pcf_renderer.py：`render_preview_scene()` 封装（encoding 参数沿用现有解码体系）。
5. pcf_app.py：预览函数组重写为矢量回放，事件绑定不动。
6. requirements/spec 清理依赖。
7. 验证（下节）+ 打包 exe。

## 验证

1. **出图回归**：改造前后对同一 PCF（`C:\新建文件夹\HQSH_Iso_Pipeline_PTFE\2105-CWS-39011.pcf`、`2105-BR3R-31001.pcf`）各渲染 PDF，改前先留基线文件，改后逐页栅格化对比像素差异应≈0（改造是纯搬家，不动绘制语义）。
2. **场景一致性**：`render_scene()` 输出的 ink_bbox 与 PDF 版式对照（写临时验证脚本，跑完删除）。
3. **GUI 手测**：启动 GUI → 选 PCF 预览矢量图 → Ctrl+滚轮缩放（光标锚点）、拖动平移、双击复位、窗口 resize 防抖——全部行为与现位图版一致；放大后文字/线条边缘清晰无像素化（矢量收益）。
4. **交互冒烟**：点击焊点条目，控制台打印其语义 tag（演示元素可寻址）。
5. 打包 exe 后启动验证预览与出图均正常、体积下降。

## 风险与预案

- **Canvas 大图性能**：几千 items 全量重画若 >200ms，启用视口裁剪绘制。
- **斜切文字近似**：预览字形不斜（沿轴排布），PDF 出图保持真斜切——明确告知，用户可接受。
- **字宽度量漂移**：布局统一用 fpdf 度量，tk 只回放最终坐标，不存在两套度量打架。
- **改造回归**：`_Canvas` 改造是机械替换，靠第 1 步像素级回归兜底。
