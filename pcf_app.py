"""PCF 收集器 - 主 GUI 应用程序（Python 版本）

打开 PCF 文件夹直接列出管线清单，单线轴测图出图/预览、焊点编号、
导出 Excel 全部直接读取 PCF 文件，不依赖数据库。
"""

import json
import math
import os
import queue
import re
import threading
import traceback
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, simpledialog

import ttkbootstrap as ttkb
from ttkbootstrap.constants import *

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# 预览矢量层抗锯齿：有 Pillow 时走超采样位图（方案同管架册 dxf_canvas.py），
# 缺 Pillow 自动回退 Tk 原生画法（边缘硬但功能不受影响）
try:
    from PIL import Image, ImageDraw, ImageTk
    HAS_AA = True
except ImportError:
    HAS_AA = False

_SS = 3   # 超采样倍数：按 3× 分辨率画矢量层后 LANCZOS 降采样，得到抗锯齿边缘


def _read_pipeline_reference(path):
    """快速读取 PCF 的管线号（只扫文件开头的 PIPELINE-REFERENCE 行）"""
    with open(path, "rb") as fh:
        raw = fh.read(65536)
    text = None
    for enc in ("gbk", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("latin-1", errors="replace")
    m = re.search(r"(?m)^PIPELINE-REFERENCE[ \t]+(\S+)", text)
    return m.group(1) if m else ""


class PcfApp:
    """PCF 收集器主应用程序"""

    def __init__(self, root):
        self.root = root
        self.root.title("PCF收集")
        self.root.geometry("1000x650")

        # 状态变量
        self.current_columns = []
        self.current_rows = []
        self._exporting = False

        # 跨线程 UI 汇聚队列：工作线程只 _post(fn)，主线程轮询泵统一执行
        # （Tkinter 非线程安全，工作线程不得直接 root.after / 摸控件）
        self._ui_queue = queue.Queue()

        self._build_ui()
        self.root.after(50, self._drain_ui_queue)

    def _build_ui(self):
        """构建界面"""
        # 配置 Treeview 样式 - 模拟网格线效果
        style = ttkb.Style.get_instance()
        if style:
            style.configure("Treeview.Heading", relief="solid", borderwidth=1, font=("-size", 10, "-weight", "bold"))
            style.configure("Treeview", rowheight=26)
            # 网格线样式
            style.layout("Grid.Treeview", style.layout("Treeview"))
            style.configure("Grid.Treeview", rowheight=26)
            style.configure("Grid.Treeview.Heading", relief="solid", borderwidth=1,
                            font=("-size", 10, "-weight", "bold"))

        # 主容器：左右分割
        main_paned = ttkb.Panedwindow(self.root, orient=tk.HORIZONTAL, bootstyle=SECONDARY)
        main_paned.pack(fill=BOTH, expand=YES, padx=6, pady=6)

        # ========== 左侧：图面预览（不含图框） ==========
        left_frame = ttkb.Frame(main_paned)
        main_paned.add(left_frame, weight=3)

        ttkb.Label(left_frame, text="图面预览（Ctrl+滚轮缩放 / 拖动平移 / 双击复位 / 单击焊点改号）",
                   font=("-size", 11, "-weight", "bold"),
                   bootstyle=INVERSE).pack(fill=X, pady=(0, 4))

        # 焊点编号 JSON 批量导入/导出：编号不落库，{weld_key: 编号} 平铺文件
        weld_btns = ttkb.Frame(left_frame)
        weld_btns.pack(fill=X, pady=(0, 4))
        ttkb.Button(weld_btns, text="导出编号 JSON", width=14,
                    command=self._on_export_weld_numbers,
                    bootstyle=(INFO, OUTLINE)).pack(side=LEFT)
        ttkb.Button(weld_btns, text="导入编号 JSON", width=14,
                    command=self._on_import_weld_numbers,
                    bootstyle=(INFO, OUTLINE)).pack(side=LEFT, padx=(6, 0))

        self.preview_canvas = tk.Canvas(left_frame, bg="white", highlightthickness=0)
        self.preview_canvas.pack(fill=BOTH, expand=YES)
        # 预览状态：矢量场景（后台线程产出）+ 视图状态机，见 _show_scene/_draw_scene
        self._preview_scene = None
        self._preview_result = None
        # 焊点编号覆盖（不落库）：{pcf 路径: {weld_key: 编号}}，允许重复号
        self._weld_overrides = {}
        self._font_cache = {}           # (bold, px) -> (tkfont.Font, ascent)
        self.preview_canvas.bind("<Configure>", lambda e: self._on_preview_resize())
        # Ctrl+滚轮缩放 / 左键拖动平移 / 双击复位 fit / 单击命中元素打印标签
        self.preview_canvas.bind("<MouseWheel>", self._on_preview_wheel)
        self.preview_canvas.bind("<Button-1>", self._on_preview_press)
        self.preview_canvas.bind("<B1-Motion>", self._on_preview_motion)
        self.preview_canvas.bind("<ButtonRelease-1>", self._on_preview_release)
        self.preview_canvas.bind("<Double-Button-1>",
                                 lambda e: self._reset_preview_view())

        # ========== 右侧：控制面板 + 属性面板 ==========
        right_frame = ttkb.Frame(main_paned)
        main_paned.add(right_frame, weight=1)
        self._main_paned = main_paned
        self.root.after(50, lambda: self._main_paned.sashpos(0, 600))

        # 按钮行
        btn_frame = ttkb.Labelframe(right_frame, text="操作", padding=8)
        btn_frame.pack(fill=X, padx=4, pady=4)

        self.btn_new = ttkb.Button(btn_frame, text="打开PCF文件夹", command=self._on_open_pcf_folder,
                                   bootstyle=(SUCCESS, OUTLINE))
        self.btn_new.pack(side=LEFT, padx=4)

        # 编码选择（紧跟在打开 PCF 文件夹按钮右侧）
        ttkb.Label(btn_frame, text="编码:", bootstyle=PRIMARY).pack(side=LEFT, padx=(8, 4))
        self.encoding_options = [
            ("自动（GBK 优先）", "auto"),
            ("GBK", "gbk"),
            ("UTF-8", "utf-8"),
            ("UTF-8-SIG", "utf-8-sig"),
            ("GB18030", "gb18030"),
        ]
        self.enc_var = tk.StringVar(value="自动（GBK 优先）")
        self.combo_encoding = ttkb.Combobox(
            btn_frame, state="readonly",
            values=[label for label, _ in self.encoding_options],
            textvariable=self.enc_var, bootstyle=PRIMARY)
        self.combo_encoding.current(0)
        self.combo_encoding.pack(side=LEFT, fill=X, expand=YES)

        # 导出表格：独立按钮（原数据表右键导出）
        self.btn_export = ttkb.Button(btn_frame, text="导出表格", command=self._on_export,
                                      bootstyle=(DANGER, OUTLINE))
        self.btn_export.pack(side=LEFT, padx=4)

        # 绘图面板：PCF -> 轴测图 PDF（基于 pcf2iso）
        render_frame = ttkb.Labelframe(right_frame, text="绘图", padding=8)
        render_frame.pack(fill=X, padx=4, pady=4)

        self.btn_render = ttkb.Button(render_frame, text="生成轴测图", command=self._on_render_iso,
                                      bootstyle=(WARNING, OUTLINE))
        self.btn_render.pack(side=LEFT, padx=4)

        # 出图设置：全部渲染参数集中到一个弹窗按钮里
        self.render_settings = {
            "north": "Y", "quadrant": "SE", "paper": "tabloid",
            "to_scale": False, "nts_min": 8.0, "nts_max": 70.0,
            "skew": 10.0, "spool": "", "ffw": False, "encoding": "auto",
        }
        self.btn_render_settings = ttkb.Button(render_frame, text="出图设置...",
                                               command=self._open_render_settings,
                                               bootstyle=(PRIMARY, OUTLINE))
        self.btn_render_settings.pack(side=LEFT, padx=4)
        self.render_summary = ttkb.Label(render_frame, text="", bootstyle=SECONDARY)
        self.render_summary.pack(side=LEFT, padx=6)
        self._update_render_summary()

        # 下拉框
        # ========== 管线清单（原类型下拉+属性详情位置） ==========
        list_frame = ttkb.Labelframe(right_frame, text="管线清单", padding=4)
        list_frame.pack(fill=BOTH, expand=YES, padx=4, pady=4)

        list_btns = ttkb.Frame(list_frame)
        list_btns.pack(fill=X, pady=(0, 3))
        ttkb.Button(list_btns, text="全选", command=lambda: self._select_all(True),
                    bootstyle=(INFO, OUTLINE), width=6).pack(side=LEFT, padx=2)
        ttkb.Button(list_btns, text="清空", command=lambda: self._select_all(False),
                    bootstyle=(SECONDARY, OUTLINE), width=6).pack(side=LEFT, padx=2)
        ttkb.Label(list_btns, text="（勾选后生成轴测图）", bootstyle=SECONDARY).pack(side=LEFT, padx=6)

        list_container = ttkb.Frame(list_frame)
        list_container.pack(fill=BOTH, expand=YES)
        self.pipe_tree = ttkb.Treeview(list_container, columns=("chk", "ref", "name", "has"),
                                       show="headings", bootstyle=PRIMARY)
        self.pipe_tree.heading("chk", text="选")
        self.pipe_tree.heading("ref", text="管线号")
        self.pipe_tree.heading("name", text="文件")
        self.pipe_tree.heading("has", text="图")
        self.pipe_tree.column("chk", width=34, anchor=CENTER, stretch=False)
        self.pipe_tree.column("ref", width=130, stretch=True)
        self.pipe_tree.column("name", width=130, stretch=True)
        self.pipe_tree.column("has", width=30, anchor=CENTER, stretch=False)
        list_scroll = ttkb.Scrollbar(list_container, orient=VERTICAL,
                                     command=self.pipe_tree.yview, bootstyle=ROUND)
        self.pipe_tree.configure(yscrollcommand=list_scroll.set)
        self.pipe_tree.grid(row=0, column=0, sticky=NSEW)
        list_scroll.grid(row=0, column=1, sticky=NS)
        list_container.grid_rowconfigure(0, weight=1)
        list_container.grid_columnconfigure(0, weight=1)
        # 单击：切换勾选并预览已有图面
        self.pipe_tree.bind("<ButtonRelease-1>", self._on_pipeline_click)
        # 管线清单数据：[{path, name, ref, checked, pdf}]
        self.pipeline_rows = []

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttkb.Label(self.root, textvariable=self.status_var, bootstyle=(INVERSE, SECONDARY),
                                anchor=W, padding=(8, 4))
        status_bar.pack(fill=X, side=BOTTOM)

    # ==================== 跨线程 UI 泵 ====================

    def _post(self, fn):
        """工作线程投递 UI 操作，由主线程 _drain_ui_queue 执行"""
        self._ui_queue.put(fn)

    def _drain_ui_queue(self):
        """主线程轮询泵：批量执行工作线程投递的 UI 操作"""
        while True:
            try:
                fn = self._ui_queue.get_nowait()
            except queue.Empty:
                break
            try:
                fn()
            except Exception:
                traceback.print_exc()
        self.root.after(50, self._drain_ui_queue)

    # ==================== 事件处理 ====================

    def _on_open_pcf_folder(self):
        """打开 PCF 文件夹：不建数据库，直接列管线清单；出图/预览/导出均读 PCF"""
        folder = filedialog.askdirectory(title="选择包含 PCF 文件的文件夹")
        if not folder:
            return
        self.btn_new.config(state=tk.DISABLED)
        self.status_var.set("正在扫描 PCF 文件...")
        thread = threading.Thread(target=self._scan_pcf_folder, args=(folder,), daemon=True)
        thread.start()

    def _scan_pcf_folder(self, folder):
        """后台扫描：递归收集 .pcf 并读取各文件的管线号（PIPELINE-REFERENCE）"""
        try:
            files = sorted(os.path.join(r, f)
                           for r, _, fs in os.walk(folder)
                           for f in fs if f.lower().endswith(".pcf"))
            if not files:
                raise FileNotFoundError("文件夹中没有 .pcf 文件")
            rows = []
            for i, f in enumerate(files):
                ref = ""
                try:
                    ref = _read_pipeline_reference(f)
                except OSError:
                    pass
                rows.append((f, ref))
                self._post(lambda c=i + 1, t=len(files): self.status_var.set(
                    f"扫描管线 {c}/{t}"))

            def done():
                self.btn_new.config(state=tk.NORMAL)
                self.pipeline_rows.clear()
                for f, ref in rows:
                    stem = os.path.splitext(os.path.basename(f))[0]
                    pdf = os.path.join(folder, "drawings", stem + ".pdf")
                    self.pipeline_rows.append({
                        "path": f, "name": stem, "ref": ref or stem, "folder": folder,
                        "checked": True,
                        "pdf": pdf if os.path.isfile(pdf) else None,
                        "missing": False,
                    })
                self._redraw_list()
                self.status_var.set(f"管线清单：{len(rows)} 条（来源 {folder}）")
            self._post(done)
        except Exception:
            err = traceback.format_exc()

            def fail():
                self.btn_new.config(state=tk.NORMAL)
                self.status_var.set("打开文件夹失败")
                messagebox.showerror("打开PCF文件夹失败", err[-1500:])
            self._post(fail)

    def _redraw_list(self):
        self.pipe_tree.delete(*self.pipe_tree.get_children())
        for i, r in enumerate(self.pipeline_rows):
            self.pipe_tree.insert("", tk.END, iid=str(i),
                                  values=("☑" if r["checked"] else "☐",
                                          r["ref"], r["name"],
                                          "✓" if r["pdf"] else ""))

    def _refresh_pdf_flags(self):
        """出图后刷新"图"列与预览（保留勾选状态）"""
        for r in self.pipeline_rows:
            folder = r.get("folder") or os.path.dirname(r["path"])
            f = os.path.join(folder, "drawings", r["name"] + ".pdf")
            r["pdf"] = f if os.path.isfile(f) else None
        self._redraw_list()

    def _select_all(self, on):
        for r in self.pipeline_rows:
            r["checked"] = on
        self._redraw_list()

    def _on_pipeline_click(self, event):
        iid = self.pipe_tree.identify_row(event.y)
        if not iid:
            return
        r = self.pipeline_rows[int(iid)]
        if self.pipe_tree.identify_column(event.x) == "#1":   # "选"列
            r["checked"] = not r["checked"]
            self._redraw_list()
        self._start_live_preview(r["path"], r["ref"])

    def _start_live_preview(self, pcf_path, ref=""):
        """后台实时渲染该管线的纯图面（不出文件、无图框/材料表）"""
        if not pcf_path or not os.path.isfile(pcf_path):
            self.status_var.set("原始 PCF 文件缺失，无法预览，请重新打开 PCF 文件夹")
            return
        self._preview_token = getattr(self, "_preview_token", 0) + 1
        token = self._preview_token
        self._preview_path = pcf_path
        self._preview_ref = ref
        s = self.render_settings
        self.status_var.set(f"预览渲染中：{ref or pcf_path} ...")

        def work():
            try:
                import pcf_renderer
                scene, result = pcf_renderer.render_preview_scene(
                    pcf_path, north_axis=s["north"], quadrant=s["quadrant"],
                    paper=s["paper"], to_scale=s["to_scale"],
                    nts_min_mm=s["nts_min"], nts_max_mm=s["nts_max"],
                    skew_deg=s["skew"], encoding=s.get("encoding", "auto"),
                    weld_number_overrides=self._weld_overrides.get(pcf_path))
                if token == self._preview_token:
                    self._post(lambda: self._show_scene(scene, result, ref))
            except Exception as exc:
                err = str(exc)

                def fail():
                    self.status_var.set(f"预览失败：{err[:120]}")
                if token == self._preview_token:
                    self._post(fail)

        threading.Thread(target=work, daemon=True).start()

    def _show_scene(self, scene, result=None, label=""):
        """缓存矢量场景并回放到预览画布：按墨迹 bbox 裁剪 + 适配缩放"""
        c = self.preview_canvas
        c.delete("all")
        cw, ch = max(c.winfo_width(), 80), max(c.winfo_height(), 60)
        if not scene.items:
            c.create_text(cw // 2, ch // 2, text="（该管线无图面内容）",
                          fill="#c0c0c0")
            return
        self._preview_scene = scene
        self._preview_result = result
        self._preview_label = label
        # 新图重置视图：回到自动 fit 居中
        self._preview_zoom = None
        self._preview_off = (0.0, 0.0)
        self._draw_scene()
        self.status_var.set(f"预览：{label}" if label else "预览完成")

    def _scene_font(self, size_pt, bold, k):
        """pt 字号 × 缩放系数 → Tk 像素字体（负字号=像素）。返回 (Font, ascent)，
        ascent 用于把 fpdf 的基线 y 换算成 Tk 的文字顶 y。"""
        px = max(4, round(size_pt / 72 * 25.4 * k))
        key = (bool(bold), px)
        cached = self._font_cache.get(key)
        if cached is None:
            f = tkfont.Font(family="Helvetica", size=-px,
                            weight=("bold" if bold else "normal"))
            cached = (f, f.metrics("ascent"))
            self._font_cache[key] = cached
        return cached

    @staticmethod
    def _scene_color(rgb):
        return f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"

    def _draw_scene(self):
        """按当前缩放/平移状态把缓存场景回放到画布（全矢量，缩放无损）"""
        scene = getattr(self, "_preview_scene", None)
        if scene is None:
            return
        c = self.preview_canvas
        # 预览无图框：按全部墨迹 bbox 裁剪 + 6mm 外边距（同旧位图版）
        x0, y0, x1, y1 = scene.ink_bbox_mm or (0.0, 0.0, scene.page_w_mm, scene.page_h_mm)
        pad = 6.0
        bx0, by0 = max(0.0, x0 - pad), max(0.0, y0 - pad)
        bx1, by1 = min(scene.page_w_mm, x1 + pad), min(scene.page_h_mm, y1 + pad)
        bw, bh = max(bx1 - bx0, 1e-6), max(by1 - by0, 1e-6)
        cw, ch = max(c.winfo_width(), 80), max(c.winfo_height(), 60)
        # 视图状态：zoom None = 自动 fit；否则 k 为 px/mm 缩放系数
        if getattr(self, "_preview_zoom", None) is None:
            k = min(cw / bw, ch / bh)
            ox = (cw - bw * k) / 2 - bx0 * k
            oy = (ch - bh * k) / 2 - by0 * k
        else:
            k = self._preview_zoom
            ox, oy = self._preview_off
        self._preview_view = (k, ox, oy)
        c.delete("all")
        PX = lambda v: v * k + ox           # 页 mm → 画布 px
        PY = lambda v: v * k + oy
        if HAS_AA:
            # 抗锯齿路径：矢量层画进 3× 超采样 PIL 图，LANCZOS 降采样贴底
            self._draw_scene_vectors_pil(c, scene.items, PX, PY, k, cw, ch)
        else:
            for i, item in enumerate(scene.items):
                if item.kind == "text":
                    continue
                if item.tags and "furniture" in item.tags:
                    continue   # 家具固定角落画，不随场景变换（见 _draw_preview_furniture）
                tag = ("elem", str(i), item.kind)
                col = self._scene_color(item.color)
                fill = self._scene_color(item.fill) if item.fill else ""
                wpx = max(1, round(item.width * k))
                if item.kind == "line":
                    c.create_line(PX(item.pts[0][0]), PY(item.pts[0][1]),
                                  PX(item.pts[1][0]), PY(item.pts[1][1]),
                                  fill=col, width=wpx, tags=tag,
                                  dash=(max(1, round(item.dash * k)),) * 2
                                  if item.dash else None)
                elif item.kind == "rect":
                    x, y, w, h = item.box
                    c.create_rectangle(PX(x), PY(y), PX(x + w), PY(y + h),
                                       outline=col if item.stroke else "",
                                       width=wpx, fill=fill, tags=tag)
                elif item.kind == "ellipse":
                    x, y, w, h = item.box
                    c.create_oval(PX(x), PY(y), PX(x + w), PY(y + h),
                                  outline=col if item.stroke else "",
                                  width=wpx, fill=fill, tags=tag)
                elif item.kind == "polygon":
                    c.create_polygon(
                        [v for p in item.pts for v in (PX(p[0]), PY(p[1]))],
                        outline=col if item.stroke else "", width=wpx,
                        fill=fill, tags=tag)
        # 文字层两路共用：Tk 系统字体本就抗锯齿，且带矩阵的能整串旋转
        for i, item in enumerate(scene.items):
            if item.kind == "text":
                if item.tags and "furniture" in item.tags:
                    continue
                self._draw_scene_text(c, item, PX, PY, k,
                                      ("elem", str(i), "text"))
        self._draw_preview_furniture(c, scene, cw, ch)

    def _draw_preview_furniture(self, c, scene, cw, ch):
        """指北针/NTS 戳固定在预览画布角落，不随 fit/缩放/平移。

        它们与管线内容无几何关联：若按场景页坐标画，内容居中后小图的
        家具会跑出画布；fit 墨迹范围已在场景端剔除它们（render._ink）。
        """
        north = None
        nts = False
        for item in scene.items:
            if not (item.tags and "furniture" in item.tags):
                continue
            if item.kind == "line" and north is None:
                (x1, y1), (x2, y2) = item.pts
                north = math.atan2(y2 - y1, x2 - x1)   # 页/画布同为 y 向下
            elif item.kind == "text" and "NOT TO SCALE" in item.text:
                nts = True
        if north is not None:
            cx, cy = 34.0, 48.0
            dx, dy = math.cos(north), math.sin(north)
            tipx, tipy = cx + dx * 22, cy + dy * 22
            c.create_line(cx - dx * 6, cy - dy * 6, tipx, tipy,
                          width=2, arrow=tk.LAST, fill="#333333")
            c.create_text(tipx + dx * 11, tipy + dy * 11, text="N",
                          font=("Helvetica", 11, "bold"), fill="#333333")
        if nts:
            c.create_text(12, ch - 10, text="NOT TO SCALE",
                          font=("Helvetica", 9, "bold"), anchor="sw",
                          fill="#666666")

    def _draw_scene_text(self, c, item, PX, PY, k, tag):
        """文字回放。普通文字整串画；带矩阵的（尺寸标注 iso_text、F.W.
        斜置标签等）整串旋转，基线顺着 CTM x 轴方向——与 PDF 出图一致，
        文字方向跟着标注线走。PDF 的 iso_text 还有字形斜切，Tk 画不出
        斜切字形，纯旋转是计划内的近似。"""
        col = self._scene_color(item.color)
        if item.matrix is None:
            font, ascent = self._scene_font(item.size, item.bold, k)
            x, y = item.pts[0]
            # fpdf text() 的 y 是基线；Tk anchor=nw 从顶画，先减 ascent
            c.create_text(PX(x), PY(y) - ascent, text=item.text,
                          fill=col, font=font, anchor="nw", tags=tag)
            return
        m = item.matrix
        dx, dy = m[0], m[1]             # CTM x 轴 = 基线方向（页 mm，y 向下）
        ux, uy = -m[2], -m[3]           # CTM (0,-1) = 字形上方向
        scale = math.hypot(dx, dy) or 1.0
        font, ascent = self._scene_font(item.size * scale, item.bold, k)
        descent = font.metrics("descent")
        x, y = item.pts[0]              # 局部基线锚点，经 CTM 映射到页 mm
        wmm = (sum(item.widths) if item.widths
               else font.measure(item.text) / k)
        sx = m[0] * x + m[2] * y + m[4]
        sy = m[1] * x + m[3] * y + m[5]
        ex = m[0] * (x + wmm) + m[2] * y + m[4]
        ey = m[1] * (x + wmm) + m[3] * y + m[5]
        # Tk angle 逆时针为正、屏幕 y 向下：基线方向 (dx,dy) → 角度 atan2(-dy,dx)
        ang = math.degrees(math.atan2(-dy, dx))
        # anchor="center" 把外接框中心放在锚点上，而基线中心在框中心沿
        # 字形下方向偏 (ascent-descent)/2 处——锚点沿字形上方向抬同量，
        # 基线段中点才落回 mid
        mid = ((PX(sx) + PX(ex)) / 2, (PY(sy) + PY(ey)) / 2)
        nun = math.hypot(ux, uy) or 1.0
        off = (ascent - descent) / 2
        c.create_text(mid[0] + ux / nun * off, mid[1] + uy / nun * off,
                      text=item.text, fill=col, font=font,
                      angle=ang, anchor="center", tags=tag)

    @staticmethod
    def _pil_dash_segment(d, p0, p1, col, w, dash):
        """PIL 线段（支持等长 on/off 虚线，单位为超采样像素）。"""
        if not dash or dash <= 0:
            d.line([p0, p1], fill=col, width=w)
            return
        x0, y0 = p0
        x1, y1 = p1
        length = math.hypot(x1 - x0, y1 - y0)
        if length < 1e-9:
            return
        ux, uy = (x1 - x0) / length, (y1 - y0) / length
        period = 2.0 * dash                     # on + off 等长（同 Tk dash=(n,n)）
        start = 0.0
        while start < length:
            end = min(start + dash, length)
            d.line([(x0 + ux * start, y0 + uy * start),
                    (x0 + ux * end, y0 + uy * end)], fill=col, width=w)
            start += period

    def _draw_scene_vectors_pil(self, c, items, PX, PY, k, cw, ch):
        """矢量层 → 3× 超采样 PIL 图 → LANCZOS 降采样 → 贴画布底层。
        文字不进位图：继续用 Tk 系统字体渲染（本就抗锯齿，且支持旋转）。"""
        ss = _SS
        img = Image.new("RGB", (max(1, cw * ss), max(1, ch * ss)), (255, 255, 255))
        d = ImageDraw.Draw(img)
        for item in items:
            if item.kind == "text":
                continue   # 文字层由 Tk 绘制
            if item.tags and "furniture" in item.tags:
                continue   # 家具固定角落画，不随场景变换
            col = tuple(item.color)
            fill = tuple(item.fill) if item.fill else None
            w = max(1, round(item.width * k * ss))
            dash = item.dash * k * ss if item.dash else None
            if item.kind == "line":
                p0 = (PX(item.pts[0][0]) * ss, PY(item.pts[0][1]) * ss)
                p1 = (PX(item.pts[1][0]) * ss, PY(item.pts[1][1]) * ss)
                self._pil_dash_segment(d, p0, p1, col, w, dash)
            elif item.kind == "rect":
                x, y, bw, bh = item.box
                box = [PX(x) * ss, PY(y) * ss, PX(x + bw) * ss, PY(y + bh) * ss]
                if item.stroke:
                    d.rectangle(box, fill=fill, outline=col, width=w)
                elif fill:
                    d.rectangle(box, fill=fill)
            elif item.kind == "ellipse":
                x, y, bw, bh = item.box
                box = [PX(x) * ss, PY(y) * ss, PX(x + bw) * ss, PY(y + bh) * ss]
                if item.stroke:
                    d.ellipse(box, fill=fill, outline=col, width=w)
                elif fill:
                    d.ellipse(box, fill=fill)
            elif item.kind == "polygon":
                pts = [(PX(p[0]) * ss, PY(p[1]) * ss) for p in item.pts]
                if fill:
                    d.polygon(pts, fill=fill)
                if item.stroke:
                    for a, b in zip(pts, pts[1:] + pts[:1]):
                        self._pil_dash_segment(d, a, b, col, w, dash)
        img = img.resize((cw, ch), Image.LANCZOS)
        self._preview_photo = ImageTk.PhotoImage(img)   # 持引用防垃圾回收
        c.create_image(0, 0, anchor="nw", image=self._preview_photo)

    @staticmethod
    def _pt_seg_dist(px, py, ax, ay, bx, by):
        """点到线段距离（页 mm）"""
        dx, dy = bx - ax, by - ay
        l2 = dx * dx + dy * dy
        if l2 < 1e-12:
            return math.hypot(px - ax, py - ay)
        t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / l2))
        return math.hypot(px - (ax + t * dx), py - (ay + t * dy))

    def _pick_scene_item(self, e):
        """画布像素 → 页 mm，按屏幕 6px 阈值找最近场景图元索引。
        抗锯齿路径下矢量在位图里，无法靠 Tk 条目命中，统一几何拾取。"""
        scene = getattr(self, "_preview_scene", None)
        view = getattr(self, "_preview_view", None)
        if scene is None or view is None:
            return None
        k, ox, oy = view
        x, y = (e.x - ox) / k, (e.y - oy) / k
        best, best_d = None, 6.0 / k
        for idx, item in enumerate(scene.items):
            if item.kind == "line":
                (ax, ay), (bx, by) = item.pts[0], item.pts[1]
                dist = self._pt_seg_dist(x, y, ax, ay, bx, by)
            elif item.kind in ("rect", "ellipse"):
                bx_, by_, bw, bh = item.box
                dx0 = max(bx_ - x, 0.0, x - (bx_ + bw))
                dy0 = max(by_ - y, 0.0, y - (by_ + bh))
                dist = math.hypot(dx0, dy0)
            elif item.kind == "polygon":
                pts = item.pts
                n = len(pts)
                dist = min(self._pt_seg_dist(x, y, pts[i][0], pts[i][1],
                                             pts[(i + 1) % n][0],
                                             pts[(i + 1) % n][1])
                           for i in range(n))
                # 填充多边形内部算命中（射线法判内）
                inside, j = False, n - 1
                for i in range(n):
                    xi, yi = pts[i]
                    xj, yj = pts[j]
                    if (yi > y) != (yj > y) and \
                            x < (xj - xi) * (y - yi) / (yj - yi) + xi:
                        inside = not inside
                    j = i
                if inside:
                    dist = 0.0
            else:  # text：基线锚点（有矩阵先映射），按字高放宽阈值
                if item.matrix:
                    m = item.matrix
                    tx = m[0] * item.pts[0][0] + m[2] * item.pts[0][1] + m[4]
                    ty = m[1] * item.pts[0][0] + m[3] * item.pts[0][1] + m[5]
                else:
                    tx, ty = item.pts[0]
                dist = max(0.0, math.hypot(x - tx, y - ty)
                           - item.size * 0.35)
            if dist < best_d:
                best, best_d = idx, dist
        return best

    def _reset_preview_view(self):
        """双击复位：回到自动 fit 居中"""
        self._preview_zoom = None
        self._preview_off = (0.0, 0.0)
        self._draw_scene()

    def _on_preview_wheel(self, e):
        """Ctrl+滚轮缩放（以光标位置为锚点）；未按 Ctrl 不响应"""
        if getattr(self, "_preview_scene", None) is None or not (e.state & 0x0004):
            return
        if getattr(self, "_preview_zoom", None) is None:
            # 从 fit 状态进入缩放：先把当前 fit 参数固化为视图状态
            self._draw_scene()
            k, ox, oy = self._preview_view
            self._preview_zoom, self._preview_off = k, (ox, oy)
        k = self._preview_zoom
        ox, oy = self._preview_off
        factor = 1.15 if e.delta > 0 else 1 / 1.15
        nk = max(0.05, min(10.0, k * factor))
        if nk == k:
            return
        # 保持光标下的图面点不动：page = (cursor - off) / k
        self._preview_zoom = nk
        self._preview_off = (e.x - (e.x - ox) * nk / k,
                             e.y - (e.y - oy) * nk / k)
        self._draw_scene()

    def _on_preview_press(self, e):
        self._preview_dragged = False
        self._preview_drag = (e.x, e.y) \
            if getattr(self, "_preview_scene", None) is not None \
            and getattr(self, "_preview_zoom", None) is not None else None

    def _on_preview_motion(self, e):
        d = getattr(self, "_preview_drag", None)
        if d is None:
            return
        dx, dy = e.x - d[0], e.y - d[1]
        if dx or dy:
            # 平移走增量 move（不重建条目），比整帧重画轻
            self.preview_canvas.move("all", dx, dy)
            self._preview_drag = (e.x, e.y)
            ox, oy = self._preview_off
            self._preview_off = (ox + dx, oy + dy)
            self._preview_dragged = True

    def _on_preview_release(self, e):
        """单击（未拖动）：几何拾取最近图元，打印其语义标签（焊点/管线/元件/标注）"""
        if getattr(self, "_preview_dragged", False):
            return
        scene = getattr(self, "_preview_scene", None)
        if scene is None:
            return
        idx = self._pick_scene_item(e)
        if idx is None:
            return
        item = scene.items[idx]
        if item.tags and item.tags[0] == "weld":
            self._edit_weld_number(item.tags)      # 单击焊点：改号对话框
            return
        print(f"[预览] kind={item.kind} tags={item.tags} text={item.text!r}")
        if item.tags:
            self.status_var.set(f"选中：{item.tags[0]} {item.tags[1]}")

    def _on_preview_resize(self):
        """窗口/分栏变化：防抖后全量重放矢量（场景已缓存，无需重新渲染）"""
        if getattr(self, "_preview_scene", None) is None:
            return
        if getattr(self, "_preview_resize_job", None):
            self.root.after_cancel(self._preview_resize_job)
        self._preview_resize_job = self.root.after(
            120, lambda: (setattr(self, "_preview_resize_job", None),
                          self._draw_scene()))

    # ==================== 焊点编号覆盖 ====================

    def _edit_weld_number(self, tags):
        """单击焊点改号：允许与别的焊点重复；清空输入=删除覆盖恢复自动编号。
        tags = ("weld", 当前编号, weld_key)，键由渲染端按 UCI/世界坐标生成，
        重复编号也能各自定位。"""
        number, key = str(tags[1]), tags[2]
        path = getattr(self, "_preview_path", None)
        if not path:
            return
        cur = self._weld_overrides.setdefault(path, {})
        value = simpledialog.askstring(
            "焊点编号",
            f"焊点当前编号 {number}\n"
            f"输入新编号（允许重复）；留空恢复自动编号：",
            initialvalue=number, parent=self.root)
        if value is None:                       # 取消：不动
            return
        value = value.strip()
        if not value:
            cur.pop(key, None)
            self.status_var.set(f"焊点 {number} 已恢复自动编号")
        else:
            try:
                cur[key] = int(value)
            except ValueError:
                self.status_var.set(f"编号 {value!r} 不是整数，已忽略")
                return
            self.status_var.set(f"焊点 {number} → {value}")
        self._refresh_weld_tags()

    def _refresh_weld_tags(self):
        """改号/导入后就地更新预览场景的焊点 tags，免整图重渲染。

        编号只是点击拾取的元数据（画布不画号），布局与几何完全不受
        覆盖表影响——渲染端 tags=(编号, weld_key, 自动号)，这里按最新
        覆盖表换算生效号写回，清空的点还原自动号。"""
        scene = getattr(self, "_preview_scene", None)
        path = getattr(self, "_preview_path", None)
        if scene is None or not path:
            return
        ov = self._weld_overrides.get(path, {})
        for item in scene.items:
            if item.tags and item.tags[0] == "weld":
                key = item.tags[2]
                auto = item.tags[3] if len(item.tags) > 3 else item.tags[1]
                # 覆盖表值与自动号均为 int，与渲染端 tags 类型契约一致
                item.tags = ("weld", ov.get(key, auto), key, auto)

    def _on_export_weld_numbers(self):
        """当前预览图的编号覆盖表导出 JSON（只含改过的点，不含自动编号）"""
        path = getattr(self, "_preview_path", None)
        if not path:
            messagebox.showinfo("导出编号", "请先在管线清单点选一条管线出预览",
                                parent=self.root)
            return
        cur = self._weld_overrides.get(path, {})
        stem = os.path.splitext(os.path.basename(path))[0]
        target = filedialog.asksaveasfilename(
            title="导出焊点编号", defaultextension=".json",
            initialfile=f"{stem}-焊点编号.json",
            filetypes=[("JSON", "*.json")], parent=self.root)
        if not target:
            return
        with open(target, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=1, sort_keys=True)
        self.status_var.set(
            f"导出 {len(cur)} 条焊点编号 → {os.path.basename(target)}")

    def _on_import_weld_numbers(self):
        """导入编号覆盖 JSON 作用于当前预览图并即时重渲染"""
        path = getattr(self, "_preview_path", None)
        if not path:
            messagebox.showinfo("导入编号", "请先在管线清单点选一条管线出预览",
                                parent=self.root)
            return
        source = filedialog.askopenfilename(
            title="导入焊点编号", filetypes=[("JSON", "*.json")],
            parent=self.root)
        if not source:
            return
        try:
            with open(source, encoding="utf-8-sig") as f:
                data = json.load(f)
            overrides = {str(k): int(v) for k, v in data.items()}
        except Exception as exc:
            messagebox.showerror("导入编号", f"JSON 无效：{exc}",
                                 parent=self.root)
            return
        self._weld_overrides[path] = overrides
        self.status_var.set(f"导入 {len(overrides)} 条焊点编号")
        self._refresh_weld_tags()

    # ==================== 轴测图渲染 ====================

    def _update_render_summary(self):
        """绘图面板上的当前出图设置摘要"""
        s = self.render_settings
        nts = "按比例" if s["to_scale"] else f"NTS {s['nts_min']:g}-{s['nts_max']:g}mm"
        spool = s["spool"] or "-"
        enc_label = next((lb for lb, v in self.encoding_options
                          if v == s.get("encoding", "auto")), "自动（GBK 优先）")
        self.render_summary.configure(
            text=f"{s['north']}/{s['quadrant']} · {s['paper']} · {nts} · 斜倾{s['skew']:g}° · 分段:{spool} · 编码:{enc_label.split('（')[0]}")

    def _open_render_settings(self):
        """出图设置弹窗：全部渲染参数集中于此"""
        dlg = ttkb.Toplevel(self.root)
        dlg.title("出图设置")
        dlg.transient(self.root)
        dlg.resizable(False, False)
        dlg.grab_set()
        s = dict(self.render_settings)

        def row(parent):
            f = ttkb.Frame(parent)
            f.pack(fill=X, pady=3)
            return f

        # --- 视角与图纸 ---
        g1 = ttkb.Labelframe(dlg, text="视角与图纸", padding=10)
        g1.pack(fill=X, padx=10, pady=6)
        f = row(g1)
        ttkb.Label(f, text="北向轴:", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 2))
        north_var = tk.StringVar(value=s["north"])
        ttkb.Combobox(f, state="readonly", width=12, values=["Y", "X", "Y-decreasing", "X-decreasing"],
                      textvariable=north_var, bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 12))
        ttkb.Label(f, text="象限:", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 2))
        quadrant_var = tk.StringVar(value=s["quadrant"])
        ttkb.Combobox(f, state="readonly", width=5, values=["NE", "NW", "SW", "SE"],
                      textvariable=quadrant_var, bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 12))
        ttkb.Label(f, text="图纸:", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 2))
        paper_var = tk.StringVar(value=s["paper"])
        ttkb.Combobox(f, state="readonly", width=8, values=["tabloid", "a3", "letter", "a4"],
                      textvariable=paper_var, bootstyle=PRIMARY).pack(side=LEFT)
        # 编码：与收集一致的选项；影响 PCF 读入与 CSV 导出（auto → CSV utf-8-sig）
        enc_labels = [label for label, _ in self.encoding_options]
        enc_value = next((v for lb, v in self.encoding_options
                          if v == s.get("encoding", "auto")), "auto")
        ttkb.Label(f, text="编码:", bootstyle=PRIMARY).pack(side=LEFT, padx=(12, 2))
        enc_var = tk.StringVar(value=next(lb for lb, v in self.encoding_options if v == enc_value))
        ttkb.Combobox(f, state="readonly", width=14, values=enc_labels,
                      textvariable=enc_var, bootstyle=PRIMARY).pack(side=LEFT)

        def enc_selected():
            label = enc_var.get()
            return next((v for lb, v in self.encoding_options if lb == label), "auto")

        # --- 非比例伸缩 ---
        g2 = ttkb.Labelframe(dlg, text="非比例伸缩 (NOT TO SCALE)", padding=10)
        g2.pack(fill=X, padx=10, pady=6)
        f = row(g2)
        to_scale_var = tk.BooleanVar(value=s["to_scale"])
        ttkb.Checkbutton(f, text="按真实比例出图（关闭 NTS）", variable=to_scale_var,
                         bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 12))
        f = row(g2)
        ttkb.Label(f, text="最短管段(mm):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 2))
        nts_min_var = tk.StringVar(value=str(s["nts_min"]))
        ttkb.Entry(f, textvariable=nts_min_var, width=6).pack(side=LEFT, padx=(0, 12))
        ttkb.Label(f, text="最长管段(mm):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 2))
        nts_max_var = tk.StringVar(value=str(s["nts_max"]))
        ttkb.Entry(f, textvariable=nts_max_var, width=6).pack(side=LEFT, padx=(0, 12))
        ttkb.Label(f, text="斜管倾斜(°):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 2))
        skew_var = tk.StringVar(value=str(s["skew"]))
        ttkb.Entry(f, textvariable=skew_var, width=6).pack(side=LEFT)

        # --- 分段与焊缝 ---
        g3 = ttkb.Labelframe(dlg, text="分段与焊缝", padding=10)
        g3.pack(fill=X, padx=10, pady=6)
        f = row(g3)
        ttkb.Label(f, text="分段属性(如SPOOL-IDENTIFIER):", bootstyle=PRIMARY).pack(side=LEFT, padx=(0, 2))
        spool_var = tk.StringVar(value=s["spool"])
        ttkb.Entry(f, textvariable=spool_var, width=20).pack(side=LEFT, padx=(0, 12))
        ffw_var = tk.BooleanVar(value=s["ffw"])
        ttkb.Checkbutton(f, text="检测现场调整焊", variable=ffw_var,
                         bootstyle=PRIMARY).pack(side=LEFT)

        def on_ok():
            try:
                s["north"] = north_var.get()
                s["quadrant"] = quadrant_var.get()
                s["paper"] = paper_var.get()
                s["to_scale"] = bool(to_scale_var.get())
                s["nts_min"] = float(nts_min_var.get() or 8.0)
                s["nts_max"] = float(nts_max_var.get() or 70.0)
                s["skew"] = float(skew_var.get() or 10.0)
                s["spool"] = spool_var.get().strip()
                s["ffw"] = bool(ffw_var.get())
                s["encoding"] = enc_selected()
            except ValueError:
                messagebox.showwarning("出图设置", "数值参数必须是数字", parent=dlg)
                return
            self.render_settings.update(s)
            self._update_render_summary()
            dlg.destroy()

        btns = ttkb.Frame(dlg)
        btns.pack(fill=X, padx=10, pady=(2, 10))
        ttkb.Button(btns, text="确定", command=on_ok, bootstyle=SUCCESS).pack(side=RIGHT, padx=4)
        ttkb.Button(btns, text="取消", command=dlg.destroy, bootstyle=(SECONDARY, OUTLINE)).pack(side=RIGHT)
        dlg.wait_window()

    def _on_render_iso(self):
        """生成轴测图：按管线清单勾选项出图；无清单时提示先打开 PCF 文件夹"""
        folder = None
        pcf_files = []
        checked = [r for r in self.pipeline_rows if r["checked"]]
        missing = [r["name"] for r in checked if r.get("missing")]
        if missing:
            messagebox.showwarning(
                "生成轴测图", "以下管线的原始 PCF 文件已缺失，请先重新打开 PCF 文件夹：\n"
                + "\n".join(missing[:10]), parent=self.root)
            return
        if checked:
            folder = checked[0].get("folder") or os.path.dirname(checked[0]["path"])
            pcf_files = [r["path"] for r in checked]
        else:
            messagebox.showinfo("生成轴测图", "请先打开 PCF 文件夹并勾选管线", parent=self.root)
            return

        s = self.render_settings

        self.btn_render.state(["disabled"])
        self.status_var.set(f"正在生成轴测图（来源: {folder}）...")
        thread = threading.Thread(
            target=self._render_iso_worker,
            args=(folder, pcf_files, s["north"], s["quadrant"], s["spool"] or None,
                  s["ffw"], s["paper"], s["to_scale"], s["nts_min"], s["nts_max"],
                  s["skew"], s.get("encoding", "auto")), daemon=True)
        thread.start()

    def _render_iso_worker(self, folder, pcf_files, north, quadrant, spool_key,
                           detect_ffw, paper, to_scale, nts_min, nts_max, skew,
                           encoding="auto"):
        """后台渲染（工作线程）"""
        try:
            import pcf_renderer

            out_dir = os.path.join(folder, "drawings")
            summaries = pcf_renderer.render_pcf_files(
                pcf_files, out_dir, north_axis=north, quadrant=quadrant,
                paper=paper, to_scale=to_scale,
                spool_key=spool_key, detect_ffw=detect_ffw,
                nts_min_mm=nts_min, nts_max_mm=nts_max, skew_deg=skew,
                encoding=encoding,
                weld_number_overrides=self._weld_overrides,
                progress_cb=lambda msg, cur, tot: self._post(
                    lambda m=msg, c=cur, t=tot: self.status_var.set(f"{m} ({c}/{t})")))

            def done():
                self.btn_render.state(["!disabled"])
                n_pdf = len(summaries)
                total_w = sum(s["welds"] for s in summaries)
                warn_lines = [f"{os.path.basename(s['pcf'])}: {w}"
                              for s in summaries for w in s["warnings"]]
                self.status_var.set(f"轴测图完成: {n_pdf} 张, 焊缝 {total_w} 个"
                                    + (f", 警告 {len(warn_lines)} 条" if warn_lines else ""))
                if summaries:
                    msg = (f"已生成 {n_pdf} 张轴测图:\n" +
                           "\n".join(os.path.basename(s['pdf']) for s in summaries[:10]) +
                           ("\n..." if n_pdf > 10 else "") +
                           f"\n\n输出目录:\n{os.path.join(folder, 'drawings')}")
                    if warn_lines:
                        msg += ("\n\n警告:\n" + "\n".join(warn_lines[:8]) +
                                ("\n..." if len(warn_lines) > 8 else ""))
                    messagebox.showinfo("生成轴测图", msg)
                    os.startfile(out_dir)
                # 出图后刷新清单"图"列并实时预览第一条管线
                self._refresh_pdf_flags()
                src = next((s["pcf"] for s in summaries), None)
                if src:
                    self._start_live_preview(src)
            self._post(done)
        except Exception:
            err = traceback.format_exc()

            def fail():
                self.btn_render.state(["!disabled"])
                self.status_var.set("轴测图生成失败")
                messagebox.showerror("生成轴测图失败", err[-1500:])
            self._post(fail)

    # ==================== 导出 Excel ====================

    def _on_export(self):
        """导出数据到 Excel：直接解析 PCF 文件，不依赖数据库"""
        if not HAS_OPENPYXL:
            messagebox.showerror("错误", "缺少 openpyxl 库，请运行: pip install openpyxl")
            return
        rows = [r for r in self.pipeline_rows if not r.get("missing")]
        if not rows:
            messagebox.showwarning("提示", "请先打开 PCF 文件夹")
            return
        if getattr(self, "_exporting", False):
            return
        folder = rows[0].get("folder") or os.path.dirname(rows[0]["path"])
        default_file = os.path.basename(folder.rstrip("\\/")) + ".xlsx"
        output_path = filedialog.asksaveasfilename(
            title="导出到 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=default_file)
        if not output_path:
            return
        encoding = "auto"
        selected_label = self.enc_var.get()
        for label, value in self.encoding_options:
            if label == selected_label:
                encoding = value
                break
        self._exporting = True
        threading.Thread(target=self._export_worker,
                         args=(rows, output_path, encoding), daemon=True).start()

    def _export_worker(self, rows, output_path, encoding):
        """后台导出：PIPELINE 一表 + 每元件类型一表（全字段+点坐标）。

        与原数据库导出同语义：同管线同 UCI 的元件归并为一行（同 spool
        多段直管只保留下料管一行）；无 UCI 的元件不去重。"""
        try:
            from pcf.parser import load
            wb = Workbook()
            wb.remove(wb.active)
            pipe_keys: list[str] = []
            pipe_rows: list[dict] = []
            comp_data: dict[str, dict] = {}
            seen = set()

            def _cell(line):
                parts = line.raw.strip().split(None, 1)
                return parts[1].strip() if len(parts) > 1 else ""

            for i, r in enumerate(rows):
                doc = load(r["path"], encoding=encoding)
                ref = r.get("ref") or r["name"]
                for pl in doc.pipelines:
                    values = {}
                    for a in pl.header_lines:
                        if a.key == "PIPELINE-REFERENCE":
                            continue
                        if a.key not in values and a.key not in pipe_keys:
                            pipe_keys.append(a.key)
                        values[a.key] = _cell(a)
                    pipe_rows.append({"REFERENCE": pl.reference, "FILE": r["name"],
                                      "values": values})
                for c in doc.real_components():
                    uci = c.get("UCI") or ""
                    if uci:
                        key = (ref, c.type, uci)
                        if key in seen:
                            continue
                        seen.add(key)
                    data = comp_data.setdefault(c.type, {"keys": [], "rows": []})
                    fields, points = {}, {}
                    for a in c.attributes:
                        val = _cell(a)
                        if a.key.endswith("-POINT") or a.key == "CO-ORDS":
                            n = sum(1 for k in points if k.startswith(a.key + "#")) + 1
                            k = f"{a.key}#{n}"
                            points[k] = val
                        else:
                            k = a.key
                            fields[k] = val
                        if k not in data["keys"]:
                            data["keys"].append(k)
                    data["rows"].append((ref, fields, points))
                self._post(lambda c=i + 1, t=len(rows): self.status_var.set(
                    f"导出解析 {c}/{t}"))

            if pipe_rows:
                ws = wb.create_sheet("PIPELINE")
                headers = ["REFERENCE", "FILE"] + pipe_keys
                for j, h in enumerate(headers):
                    ws.cell(row=1, column=j + 1, value=h)
                for i, pr in enumerate(pipe_rows):
                    ws.cell(row=i + 2, column=1, value=pr["REFERENCE"])
                    ws.cell(row=i + 2, column=2, value=pr["FILE"])
                    for j, k in enumerate(pipe_keys):
                        ws.cell(row=i + 2, column=3 + j,
                                value=self._safe_str(pr["values"].get(k)))
            for comp_type, data in comp_data.items():
                ws = wb.create_sheet(comp_type[:31])
                headers = ["REFERENCE"] + data["keys"]
                for j, h in enumerate(headers):
                    ws.cell(row=1, column=j + 1, value=h)
                for i, (ref, fields, points) in enumerate(data["rows"]):
                    ws.cell(row=i + 2, column=1, value=ref)
                    for j, k in enumerate(data["keys"]):
                        v = fields.get(k, points.get(k))
                        ws.cell(row=i + 2, column=2 + j, value=self._safe_str(v))

            wb.save(output_path)

            def done():
                self._exporting = False
                self.status_var.set("导出完成")
                messagebox.showinfo("成功", "成功输出文本到Excel！")
            self._post(done)
        except Exception:
            err = traceback.format_exc()

            def fail():
                self._exporting = False
                self.status_var.set("导出失败")
                messagebox.showerror("错误", err[-1500:])
            self._post(fail)

    # ==================== 辅助方法 ====================

    @staticmethod
    def _safe_str(value):
        if value is None:
            return ""
        return str(value)


def main():
    root = ttkb.Window(themename="cosmo")
    app = PcfApp(root)
    root.mainloop()


if __name__ == "__main__":
    import multiprocessing

    # PyInstaller 冻结 exe 的多进程渲染（spawn）必需：子进程引导要经过这里
    multiprocessing.freeze_support()
    main()
