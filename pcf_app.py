"""PCF 收集器 - 主 GUI 应用程序（Python 版本）

功能等同于 VB.NET 版 PCF 收集器：
1. 选择文件夹，批量解析 PCF 文件并存入 SQLite 数据库
2. 打开已有数据库，浏览组件数据
3. 右键导出数据到 Excel
"""

import io
import os
import re
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

import ttkbootstrap as ttkb
from ttkbootstrap.constants import *

from database_helper import DatabaseHelper
from pcf_data_collector import PcfDataCollector
from pcf_parser import PcfParser

try:
    from openpyxl import Workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# 图面预览：PyMuPDF 渲染 PDF + Pillow 显示（ttkbootstrap 自带 Pillow）
try:
    import fitz  # PyMuPDF
    from PIL import Image, ImageTk
    HAS_PREVIEW = True
except ImportError:
    HAS_PREVIEW = False


class PcfApp:
    """PCF 收集器主应用程序"""

    COMPONENT_TYPES = PcfDataCollector.COMPONENT_TYPES

    def __init__(self, root):
        self.root = root
        self.root.title("PCF收集")
        self.root.geometry("1000x650")

        # 状态变量
        self.db_helper = None
        self.data_collector = None
        self.db_path = None
        self.current_columns = []
        self.current_rows = []
        self.success_count = 0
        self.error_count = 0
        self.error_log = []
        self._processing = False

        self._build_ui()

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

        ttkb.Label(left_frame, text="图面预览（Ctrl+滚轮缩放 / 拖动平移 / 双击复位）",
                   font=("-size", 11, "-weight", "bold"),
                   bootstyle=INVERSE).pack(fill=X, pady=(0, 4))

        self.preview_canvas = tk.Canvas(left_frame, bg="#404040", highlightthickness=0)
        self.preview_canvas.pack(fill=BOTH, expand=YES)
        self._preview_photo = None
        self._preview_pdf = None
        if not HAS_PREVIEW:
            self.preview_canvas.create_text(
                200, 60, text="缺少 PyMuPDF/Pillow，无法预览\npip install pymupdf",
                fill="white", anchor="nw")
        self.preview_canvas.bind("<Configure>", lambda e: self._on_preview_resize())
        # Ctrl+滚轮缩放 / 左键拖动平移 / 双击复位 fit
        self.preview_canvas.bind("<MouseWheel>", self._on_preview_wheel)
        self.preview_canvas.bind("<Button-1>", self._on_preview_press)
        self.preview_canvas.bind("<B1-Motion>", self._on_preview_motion)
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

        self.btn_new = ttkb.Button(btn_frame, text="新建数据库", command=self._on_new_database,
                                   bootstyle=(SUCCESS, OUTLINE))
        self.btn_new.pack(side=LEFT, padx=4)

        self.btn_open = ttkb.Button(btn_frame, text="打开数据库", command=self._on_open_database,
                                    bootstyle=(INFO, OUTLINE))
        self.btn_open.pack(side=LEFT, padx=4)

        # 编码选择（紧跟在打开数据库按钮右侧）
        ttkb.Label(btn_frame, text="编码:", bootstyle=PRIMARY).pack(side=LEFT, padx=(8, 4))
        self.encoding_options = [
            ("自动（GBK 优先）", "auto"),
            ("GBK", "gbk"),
            ("UTF-8", "utf-8"),
            ("UTF-8-SIG", "utf-8-sig"),
            ("Latin-1", "latin-1"),
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

    # ==================== 事件处理 ====================

    def _on_new_database(self):
        """新建数据库 - 选择文件夹并处理 PCF 文件"""
        folder = filedialog.askdirectory(title="选择包含 PCF 文件的文件夹")
        if not folder:
            return

        pcf_files = []
        for root_dir, dirs, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".pcf"):
                    pcf_files.append(os.path.join(root_dir, f))

        if not pcf_files:
            messagebox.showwarning("提示", "木有发现pcf")
            return

        self.db_path = os.path.join(folder, "pcf.db")
        self._processing = True
        self.btn_new.config(state=tk.DISABLED)
        self.btn_open.config(state=tk.DISABLED)

        # 关闭当前数据库连接（避免覆写时文件被占用）
        if self.db_helper is not None:
            self.db_helper.close()
            self.db_helper = None

        # 后台线程处理
        thread = threading.Thread(target=self._process_files, args=(folder, pcf_files), daemon=True)
        thread.start()

    def _process_files(self, folder, pcf_files):
        """后台处理 PCF 文件（在工作线程中运行）"""
        total = len(pcf_files)
        self.success_count = 0
        self.error_count = 0
        self.error_log.clear()

        # 初始化数据库
        if os.path.exists(self.db_path):
            os.remove(self.db_path)

        self.db_helper = DatabaseHelper(self.db_path)

        # 创建数据收集器
        self.data_collector = PcfDataCollector()
        self.data_collector.db_helper = self.db_helper

        # 创建所有表
        self.data_collector.create_all_tables()

        # 解析所有 PCF 文件
        encoding = "auto"
        selected_label = self.enc_var.get()
        for label, value in self.encoding_options:
            if label == selected_label:
                encoding = value
                break
        parser = PcfParser(self.data_collector, encoding=encoding)
        fid = 0
        parsed_files = []  # 成功解析的文件（相对路径），供轴测图渲染定位
        for i, pcf_file in enumerate(pcf_files):
            try:
                consumed = parser.process_single_file(pcf_file, fid)
                fid += consumed
                self.success_count += 1
                parsed_files.append(os.path.relpath(pcf_file, folder))
            except Exception as e:
                self.error_count += 1
                error_msg = f"文件 {i + 1}: {os.path.basename(pcf_file)} - {e}"
                self.error_log.append(error_msg)

            # 更新进度
            current = i + 1
            self.root.after(0, lambda c=current, t=total: self.status_var.set(
                f"{c}/{t} (成功:{self.success_count}, 失败:{self.error_count})"))

        # 记录导入来源，轴测图渲染时直接从数据库定位原始 PCF
        if parsed_files:
            self.data_collector.add_import_info("SOURCE-FOLDER", folder)
            self.data_collector.add_import_info("SOURCE-FILES", "\n".join(parsed_files))

        # 插入数据 & 创建视图
        try:
            self.root.after(0, lambda: self.status_var.set("插入数据表..."))
            self.data_collector.insert_all_data()

            self.root.after(0, lambda: self.status_var.set("插入视图..."))
            self.data_collector.create_sqlite_view_all(self.COMPONENT_TYPES)

            # COUNT 视图
            self.db_helper.execute_non_query("""
                DROP VIEW IF EXISTS COUNT;
                CREATE VIEW IF NOT EXISTS COUNT AS
                SELECT PIPELINE.FILENAME, REFERENCE, TOTAL.ITEMCODE, ITEMDESCRIPTION, COUNT(*) As COUNT
                FROM (SELECT * FROM TOTAL GROUP BY UCI,FILENAME) AS TOTAL
                Join PIPELINE ON PIPELINE.FILENAME = TOTAL.FILENAME
                Group BY REFERENCE, TOTAL.ITEMCODE, ITEMDESCRIPTION
                ORDER BY REFERENCE
            """)

            # 各组件类型视图
            self._create_component_views()

            # 带属性的组件视图
            self.data_collector.create_component_with_attributes_view()
        except Exception:
            err_msg = traceback.format_exc()
            self.root.after(0, lambda: self._on_processing_error(err_msg))
            return

        # 完成
        self.root.after(0, self._on_processing_done)

    def _create_component_views(self):
        """为每种组件类型创建视图"""
        for comp_type in self.COMPONENT_TYPES:
            sql = self._get_view_sql(comp_type)
            if sql:
                self.db_helper.execute_non_query(sql)

    def _get_view_sql(self, comp_type):
        """获取组件视图 SQL"""
        xname = f"X{comp_type}"
        if comp_type == "PIPE":
            return f"""DROP VIEW IF EXISTS {xname};
                CREATE VIEW IF NOT EXISTS {xname} AS
                SELECT PIPELINE.FILENAME,UCI, REFERENCE, ITEMCODE, ITEMDESCRIPTION, CUTPIECELENGTH
                FROM (SELECT * FROM {comp_type} GROUP BY UCI,FILENAME) AS {comp_type}
                JOIN PIPELINE ON PIPELINE.FILENAME = {comp_type}.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "BEND":
            return """DROP VIEW IF EXISTS XBEND;
                CREATE VIEW IF NOT EXISTS XBEND AS
                SELECT PIPELINE.FILENAME,UCI, REFERENCE,ITEMCODE, ITEMDESCRIPTION,CUTPIECELENGTH,ANGLE
                FROM (SELECT * FROM BEND GROUP BY UCI,FILENAME) AS BEND
                JOIN PIPELINE ON PIPELINE.FILENAME = BEND.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type in ("INSTRUMENT", "INSTRUMENT_ANGLE"):
            return f"""DROP VIEW IF EXISTS {xname};
                CREATE VIEW IF NOT EXISTS {xname} AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, TAG
                FROM (SELECT * FROM {comp_type} GROUP BY UCI,FILENAME) AS {comp_type}
                JOIN PIPELINE ON PIPELINE.FILENAME = {comp_type}.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "INSTRUMENT_RETURN":
            return """DROP VIEW IF EXISTS XINSTRUMENT_RETURN;
                CREATE VIEW IF NOT EXISTS XINSTRUMENT_RETURN AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, TAG, ITEMDESCRIPTION
                FROM (SELECT * FROM INSTRUMENT_RETURN GROUP BY UCI,FILENAME) AS INSTRUMENT_RETURN
                JOIN PIPELINE ON PIPELINE.FILENAME = INSTRUMENT_RETURN.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "TRAP_OFFSET":
            return """DROP VIEW IF EXISTS XTRAP_OFFSET;
                CREATE VIEW IF NOT EXISTS XTRAP_OFFSET AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, TAG, ITEMDESCRIPTION
                FROM (SELECT * FROM TRAP_OFFSET GROUP BY UCI,FILENAME) AS TRAP_OFFSET
                JOIN PIPELINE ON PIPELINE.FILENAME = TRAP_OFFSET.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "SUPPORT":
            return """DROP VIEW IF EXISTS XSUPPORT;
                CREATE VIEW IF NOT EXISTS XSUPPORT AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,NAME, ITEMDESCRIPTION,SUPPORTDIRECTION,SUPPORTTYPE
                FROM (SELECT * FROM SUPPORT GROUP BY UCI,FILENAME) AS SUPPORT
                JOIN PIPELINE ON PIPELINE.FILENAME = SUPPORT.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "GASKET":
            return """DROP VIEW IF EXISTS XGASKET;
                CREATE VIEW IF NOT EXISTS XGASKET AS
                SELECT PIPELINE.FILENAME,GASKET.UCI,REFERENCE,GASKET.ITEMCODE, GASKET.ITEMDESCRIPTION,
                       TOTAL.ITEMCODE,TOTAL.ITEMDESCRIPTION
                FROM (SELECT * FROM GASKET GROUP BY UCI,FILENAME) AS GASKET
                JOIN PIPELINE ON PIPELINE.FILENAME = GASKET.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = GASKET.FILENAME AND TOTAL.COMPONENTIDENTIFIER=GASKET.MASTERCOMPONENTIDENTIFIER
                ORDER BY REFERENCE"""
        elif comp_type == "BOLT":
            return """DROP VIEW IF EXISTS XBOLT;
                CREATE VIEW IF NOT EXISTS XBOLT AS
                SELECT PIPELINE.FILENAME,BOLT.UCI,REFERENCE,BOLTITEMCODE,BOLTITEMDESCRIPTION,
                       BOLTDIA, BOLTQUANTITY, BOLTLENGTH, ITEMCODE,ITEMDESCRIPTION
                FROM (SELECT * FROM BOLT GROUP BY UCI,FILENAME) AS BOLT
                JOIN PIPELINE ON PIPELINE.FILENAME = BOLT.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = BOLT.FILENAME AND TOTAL.COMPONENTIDENTIFIER=BOLT.MASTERCOMPONENTIDENTIFIER
                ORDER BY REFERENCE"""
        elif comp_type == "WELD":
            return """DROP VIEW IF EXISTS XWELD;
                CREATE VIEW IF NOT EXISTS XWELD AS
                SELECT PIPELINE.FILENAME, WELD.UCI, REFERENCE, Q.NPD,
                       TOTAL.ITEMCODE, TOTAL.ITEMDESCRIPTION
                FROM (SELECT * FROM WELD GROUP BY UCI, FILENAME) AS WELD
                JOIN PIPELINE ON PIPELINE.FILENAME = WELD.FILENAME
                JOIN TOTAL ON TOTAL.FILENAME = WELD.FILENAME AND TOTAL.COMPONENTIDENTIFIER = WELD.MASTERCOMPONENTIDENTIFIER
                JOIN (SELECT * FROM END_POINT GROUP BY FILENAME, UCI) AS Q ON Q.FILENAME = WELD.FILENAME AND Q.UCI = WELD.UCI
                ORDER BY REFERENCE"""
        elif comp_type == "TEE_STUB":
            return """DROP VIEW IF EXISTS XTEE_STUB;
                CREATE VIEW IF NOT EXISTS XTEE_STUB AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE
                FROM (SELECT * FROM TEE_STUB GROUP BY UCI,FILENAME) AS TEE_STUB
                JOIN PIPELINE ON PIPELINE.FILENAME = TEE_STUB.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "_UNION":
            return """DROP VIEW IF EXISTS X_UNION;
                CREATE VIEW IF NOT EXISTS X_UNION AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, ITEMDESCRIPTION
                FROM (SELECT * FROM _UNION GROUP BY UCI,FILENAME) AS _UNION
                JOIN PIPELINE ON PIPELINE.FILENAME = _UNION.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "MESSAGE":
            return """DROP VIEW IF EXISTS XMESSAGE;
                CREATE VIEW IF NOT EXISTS XMESSAGE AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,TEXT
                FROM (SELECT * FROM MESSAGE GROUP BY UCI,FILENAME) AS MESSAGE
                JOIN PIPELINE ON PIPELINE.FILENAME = MESSAGE.FILENAME
                ORDER BY REFERENCE"""
        elif comp_type == "VALVE":
            return """DROP VIEW IF EXISTS XVALVE;
                CREATE VIEW IF NOT EXISTS XVALVE AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, ITEMDESCRIPTION,SPINDLEDIRECTION,DIRECTION
                FROM (SELECT * FROM VALVE GROUP BY UCI,FILENAME) AS VALVE
                JOIN PIPELINE ON PIPELINE.FILENAME = VALVE.FILENAME
                ORDER BY REFERENCE"""
        else:
            return f"""DROP VIEW IF EXISTS {xname};
                CREATE VIEW IF NOT EXISTS {xname} AS
                SELECT PIPELINE.FILENAME,UCI,REFERENCE,ITEMCODE, ITEMDESCRIPTION
                FROM (SELECT * FROM {comp_type} GROUP BY UCI,FILENAME) AS {comp_type}
                JOIN PIPELINE ON PIPELINE.FILENAME = {comp_type}.FILENAME
                ORDER BY REFERENCE"""

    def _on_processing_done(self):
        """文件处理完成回调"""
        self._processing = False
        self.btn_new.config(state=tk.NORMAL)
        self.btn_open.config(state=tk.NORMAL)
        self.status_var.set("写入数据库")

        # 显示处理结果
        if self.error_count > 0:
            error_summary = f"处理完成。成功: {self.success_count}, 失败: {self.error_count}\n\n"
            error_summary += "错误详情 (前10条):\n"
            error_summary += "\n".join(self.error_log[:10])
            if len(self.error_log) > 10:
                error_summary += f"\n... 还有 {len(self.error_log) - 10} 个错误"
            messagebox.showwarning("PCF处理结果", error_summary)
        elif self.success_count > 0:
            messagebox.showinfo("PCF处理完成", f"所有 {self.success_count} 个文件处理成功！")

        self._refresh_pipeline_list()

    def _on_processing_error(self, err_msg):
        """处理过程中出错回调（显示真实异常，避免线程静默退出）"""
        self._processing = False
        self.btn_new.config(state=tk.NORMAL)
        self.btn_open.config(state=tk.NORMAL)
        self.status_var.set("处理出错")
        messagebox.showerror("处理出错", err_msg[-2000:])

    def _on_open_database(self):
        """打开已有数据库"""
        db_file = filedialog.askopenfilename(title="选择数据库文件", filetypes=[("SQLite", "*.db")])
        if not db_file:
            return

        self.db_path = db_file
        self.db_helper = DatabaseHelper(db_file)
        self._refresh_pipeline_list()

    # ==================== 数据加载 ====================

    # ==================== 管线清单与图面预览 ====================

    def _refresh_pipeline_list(self):
        """从数据库导入来源重建管线清单（新建/打开数据库后调用）"""
        self.pipeline_rows = []
        folder, files = None, []
        if self.db_helper is not None and self.db_path:
            try:
                import pcf_renderer
                got = pcf_renderer.sources_from_db(self.db_path)
                if got:
                    folder, files = got
            except Exception:
                pass
        ref_map = {}
        if self.db_helper is not None:
            try:
                _cols, rows = self.db_helper.fill_data_table(
                    "SELECT FILENAME, REFERENCE FROM PIPELINE")
                for r in rows:
                    if r[0] is not None:
                        ref_map[str(r[0])] = "" if r[1] is None else str(r[1])
            except Exception:
                pass
        for f in files:
            base = os.path.basename(f)
            stem = os.path.splitext(base)[0]
            ref = next((v for k, v in ref_map.items()
                        if k == base or k.endswith(base) or base.endswith(k)), "")
            pdf = os.path.join(folder, "drawings", stem + ".pdf") if folder else None
            self.pipeline_rows.append({
                "path": f, "name": stem, "ref": ref or stem, "folder": folder,
                "checked": True,
                "pdf": pdf if (pdf and os.path.isfile(pdf)) else None,
            })
        self._redraw_list()
        self.status_var.set(
            f"管线清单：{len(self.pipeline_rows)} 条（来源 {folder or '—'}）")

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
        self._preview_token = getattr(self, "_preview_token", 0) + 1
        token = self._preview_token
        s = self.render_settings
        self.status_var.set(f"预览渲染中：{ref or pcf_path} ...")

        def work():
            try:
                import pcf_renderer
                data = pcf_renderer.render_preview_bytes(
                    pcf_path, north_axis=s["north"], quadrant=s["quadrant"],
                    paper=s["paper"], to_scale=s["to_scale"],
                    nts_min_mm=s["nts_min"], nts_max_mm=s["nts_max"],
                    skew_deg=s["skew"], encoding=s.get("encoding", "auto"))
                if token == self._preview_token:
                    self.root.after(0, lambda: self._show_preview_bytes(data, ref))
            except Exception as exc:
                err = str(exc)

                def fail():
                    self.status_var.set(f"预览失败：{err[:120]}")
                if token == self._preview_token:
                    self.root.after(0, fail)

        threading.Thread(target=work, daemon=True).start()

    def _show_preview_bytes(self, pdf_bytes, label=""):
        """把内存 PDF 画到预览画布：内容 bbox 裁剪 + 适配缩放"""
        if not HAS_PREVIEW:
            return
        c = self.preview_canvas
        c.delete("all")
        cw, ch = max(c.winfo_width(), 80), max(c.winfo_height(), 60)
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            page = doc[0]
            # 预览模式无图框：直接取全部墨迹 bbox（矢量 + 文字——注释、
            # NOT TO SCALE 戳等是纯文字，漏掉会被裁掉）
            x0, y0, x1, y1 = page.rect.width, page.rect.height, 0.0, 0.0
            rects = [dr["rect"] for dr in page.get_drawings()]
            rects += [fitz.Rect(w[:4]) for w in page.get_text("words")]
            for r in rects:
                x0, y0 = min(x0, r.x0), min(y0, r.y0)
                x1, y1 = max(x1, r.x1), max(y1, r.y1)
            if x1 <= x0 or y1 <= y0:
                c.create_text(cw // 2, ch // 2, text="（该管线无图面内容）",
                              fill="#c0c0c0")
                return
            pad = 6 / 25.4 * 72
            clip = fitz.Rect(max(0, x0 - pad), max(0, y0 - pad),
                             min(page.rect.width, x1 + pad),
                             min(page.rect.height, y1 + pad))
            pix = page.get_pixmap(dpi=150, clip=clip)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            doc.close()
            # 缓存高清位图：窗口尺寸变化时只重缩放，不再重渲染
            self._preview_img = img
            self._preview_label = label
            # 新图重置视图：清空画布条目引用并回到 fit
            self._preview_item = None
            self._preview_photo = None
            self._preview_zoom = None
            self._draw_preview_img()
            self.status_var.set(f"预览：{label}" if label else "预览完成")
        except Exception:
            c.create_text(cw // 2, ch // 2, text="预览失败", fill="white")

    def _draw_preview_img(self):
        """按当前缩放/平移状态把缓存位图画到画布（缩放变化时重缩放，平移时只挪坐标）"""
        img = getattr(self, "_preview_img", None)
        if img is None:
            return
        c = self.preview_canvas
        cw, ch = max(c.winfo_width(), 80), max(c.winfo_height(), 60)
        # 视图状态：zoom None = 自动 fit；否则 k 为对原图的缩放系数
        if getattr(self, "_preview_zoom", None) is None:
            k = min(cw / img.width, ch / img.height)
            ox = (cw - img.width * k) / 2
            oy = (ch - img.height * k) / 2
        else:
            k = self._preview_zoom
            ox, oy = self._preview_off
        self._preview_view = (k, ox, oy)
        w, h = max(1, int(img.width * k)), max(1, int(img.height * k))
        # 仅在目标尺寸变化时才重缩放位图（平移不触发）
        if getattr(self, "_preview_photo", None) is None or \
                (self._preview_photo.width(), self._preview_photo.height()) != (w, h):
            self._preview_photo = ImageTk.PhotoImage(img.resize((w, h)))
        if getattr(self, "_preview_item", None) is None:
            self._preview_item = c.create_image(ox, oy, image=self._preview_photo,
                                                anchor="nw")
        else:
            c.itemconfigure(self._preview_item, image=self._preview_photo)
            c.coords(self._preview_item, ox, oy)

    def _reset_preview_view(self):
        """双击复位：回到自动 fit 居中"""
        self._preview_zoom = None
        self._preview_off = (0.0, 0.0)
        self._draw_preview_img()

    def _on_preview_wheel(self, e):
        """Ctrl+滚轮缩放（以光标位置为锚点）；未按 Ctrl 不响应"""
        if getattr(self, "_preview_img", None) is None or not (e.state & 0x0004):
            return
        if getattr(self, "_preview_zoom", None) is None:
            # 从 fit 状态进入缩放：先把当前 fit 参数固化为视图状态
            self._draw_preview_img()
            k, ox, oy = self._preview_view
            self._preview_zoom, self._preview_off = k, (ox, oy)
        k = self._preview_zoom
        ox, oy = self._preview_off
        factor = 1.15 if e.delta > 0 else 1 / 1.15
        nk = max(0.05, min(10.0, k * factor))
        if nk == k:
            return
        # 保持光标下的图像点不动：img = (cursor - off) / k
        self._preview_zoom = nk
        self._preview_off = (e.x - (e.x - ox) * nk / k,
                             e.y - (e.y - oy) * nk / k)
        self._draw_preview_img()

    def _on_preview_press(self, e):
        self._preview_drag = (e.x, e.y, *self._preview_off) \
            if getattr(self, "_preview_zoom", None) is not None else None

    def _on_preview_motion(self, e):
        if getattr(self, "_preview_drag", None) is None:
            return
        sx, sy, ox, oy = self._preview_drag
        self._preview_off = (ox + e.x - sx, oy + e.y - sy)
        self._draw_preview_img()

    def _on_preview_resize(self):
        """窗口/分栏变化：防抖后用缓存位图重绘（无需重新渲染 PDF）"""
        if getattr(self, "_preview_img", None) is None:
            self._show_preview(self._preview_pdf)
            return
        if getattr(self, "_preview_resize_job", None):
            self.root.after_cancel(self._preview_resize_job)
        self._preview_resize_job = self.root.after(
            120, lambda: (setattr(self, "_preview_resize_job", None),
                          self._draw_preview_img()))

    def _show_preview(self, pdf_path):
        """图面预览：渲染 PDF 首页，裁掉贴边的图框/标题栏，只留图面内容"""
        self._preview_pdf = pdf_path
        if not HAS_PREVIEW:
            return
        c = self.preview_canvas
        c.delete("all")
        cw, ch = max(c.winfo_width(), 80), max(c.winfo_height(), 60)
        if not pdf_path or not os.path.isfile(pdf_path):
            c.create_text(cw // 2, ch // 2, text="（无图面——勾选管线后点“生成轴测图”）",
                          fill="#c0c0c0")
            return
        try:
            doc = fitz.open(pdf_path)
            page = doc[0]
            W, H = page.rect.width, page.rect.height
            margin = 8 / 25.4 * 72          # 距页边 8mm 内的线视为图框/表格骨架
            x0, y0, x1, y1 = W, H, 0.0, 0.0
            for dr in page.get_drawings():
                r = dr["rect"]
                if r.x0 < margin or r.y0 < margin or r.x1 > W - margin or r.y1 > H - margin:
                    continue
                x0, y0 = min(x0, r.x0), min(y0, r.y0)
                x1, y1 = max(x1, r.x1), max(y1, r.y1)
            if x1 <= x0 or y1 <= y0:
                x0, y0, x1, y1 = margin, margin, W - margin, H - margin
            pad = 6 / 25.4 * 72
            clip = fitz.Rect(max(0, x0 - pad), max(0, y0 - pad),
                             min(W, x1 + pad), min(H, y1 + pad))
            pix = page.get_pixmap(dpi=150, clip=clip)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            doc.close()
            # 缓存高清位图：窗口尺寸变化时只重缩放，不再重渲染
            self._preview_img = img
            self._preview_label = ""
            # 新图重置视图：清空画布条目引用并回到 fit
            self._preview_item = None
            self._preview_photo = None
            self._preview_zoom = None
            self._draw_preview_img()
        except Exception:
            c.create_text(cw // 2, ch // 2, text="预览失败", fill="white")

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
        """生成轴测图：按管线清单勾选项出图；无清单时回退数据库来源/选文件夹"""
        import pcf_renderer

        folder = None
        pcf_files = []
        # 1) 管线清单勾选项优先
        checked = [r for r in self.pipeline_rows if r["checked"]]
        if checked:
            folder = checked[0].get("folder") or os.path.dirname(checked[0]["path"])
            pcf_files = [r["path"] for r in checked]
        else:
            sources = None
            # 2) 当前已打开的数据库
            if self.db_helper is not None and self.db_path:
                sources = pcf_renderer.sources_from_db(self.db_path)

            # 3) 让用户选一个数据库文件（读取其导入来源）
            if sources is None:
                db_file = filedialog.askopenfilename(
                    title="选择收集数据库（定位其导入的 PCF 文件）",
                    filetypes=[("SQLite 数据库", "*.db"), ("所有文件", "*.*")])
                if not db_file:
                    return
                sources = pcf_renderer.sources_from_db(db_file)

            # 4) 数据库也定位不到来源（旧库/PCF 已移动）→ 回退选 PCF 文件夹
            if sources is None:
                sel = filedialog.askdirectory(title="数据库中无来源记录，请直接选择 PCF 文件夹")
                if not sel:
                    return
                sources = pcf_renderer.collect_pcf_files([sel])
                sources = (sel, [str(p) for p in sources])
                if not sources[1]:
                    messagebox.showwarning("提示", "所选文件夹中没有 .pcf 文件")
                    return
            folder, pcf_files = sources

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
                progress_cb=lambda msg, cur, tot: self.root.after(
                    0, lambda m=msg, c=cur, t=tot: self.status_var.set(f"{m} ({c}/{t})")))

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
            self.root.after(0, done)
        except Exception:
            err = traceback.format_exc()

            def fail():
                self.btn_render.state(["!disabled"])
                self.status_var.set("轴测图生成失败")
                messagebox.showerror("生成轴测图失败", err[-1500:])
            self.root.after(0, fail)

    # ==================== 导出 Excel ====================

    def _on_export(self):
        """导出数据到 Excel"""
        if not HAS_OPENPYXL:
            messagebox.showerror("错误", "缺少 openpyxl 库，请运行: pip install openpyxl")
            return

        valid_name = re.sub(r'[\x00-\x1F\x7F<>?*:"/\\|]', '', os.path.splitext(os.path.basename(self.db_path))[0])
        valid_name = valid_name.rstrip(' .-').lstrip('.')
        default_file = f"{valid_name}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="导出到 Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=default_file)
        if not output_path:
            return

        try:
            wb = Workbook()
            # 删除默认 sheet
            wb.remove(wb.active)

            # ========== PIPELINE 表 ==========
            # 获取所有属性名
            attr_cols, attr_rows = self.db_helper.fill_data_table(
                "SELECT DISTINCT ATTRIBUTE_NAME FROM COMPONENT_ATTRIBUTES WHERE UCI LIKE 'PIPELINE:%' ORDER BY ATTRIBUTE_NAME")
            attribute_names = [r[0] for r in attr_rows]

            pipe_cols, pipe_rows = self.db_helper.fill_data_table(
                "SELECT p.FILENAME, p.REFERENCE, p.PIPINGSPEC, p.INSULATIONSPEC, p.PAINTINGSPEC, p.TRACINGSPEC FROM PIPELINE p")

            if pipe_rows:
                ws = wb.create_sheet("PIPELINE")
                # 列标题（移除 FILENAME）
                headers = ["REFERENCE", "PIPINGSPEC", "INSULATIONSPEC", "PAINTINGSPEC", "TRACINGSPEC"]
                headers.extend(attribute_names)
                for j, h in enumerate(headers):
                    ws.cell(row=1, column=j + 1, value=h)

                for i, prow in enumerate(pipe_rows):
                    ws.cell(row=i + 2, column=1, value=self._safe_str(prow[1]))
                    ws.cell(row=i + 2, column=2, value=self._safe_str(prow[2]))
                    ws.cell(row=i + 2, column=3, value=self._safe_str(prow[3]))
                    ws.cell(row=i + 2, column=4, value=self._safe_str(prow[4]))
                    ws.cell(row=i + 2, column=5, value=self._safe_str(prow[5]))

                    # 获取属性
                    pipeline_uci = f"PIPELINE:{self._safe_str(prow[1])}"
                    for j, attr_name in enumerate(attribute_names):
                        attr_val_cols, attr_val_rows = self.db_helper.fill_data_table(
                            "SELECT ATTRIBUTE_VALUE FROM COMPONENT_ATTRIBUTES WHERE UCI=? AND ATTRIBUTE_NAME=?",
                            (pipeline_uci, attr_name))
                        val = self._safe_str(attr_val_rows[0][0]) if attr_val_rows else ""
                        ws.cell(row=i + 2, column=6 + j, value=val)

            # ========== 组件表 ==========
            for comp_type in self.COMPONENT_TYPES:
                view_name = "X_UNION" if comp_type == "_UNION" else f"X{comp_type}"
                cols, rows = self.db_helper.fill_data_table(f"SELECT * FROM {view_name}")
                if rows and "FILENAME" in cols:
                    ws = wb.create_sheet(comp_type)
                    # 移除 FILENAME 列
                    file_idx = cols.index("FILENAME")
                    display_cols = [c for i, c in enumerate(cols) if i != file_idx]

                    # 列标题
                    for j, h in enumerate(display_cols):
                        ws.cell(row=1, column=j + 1, value=h)

                    # 数据行
                    for i, drow in enumerate(rows):
                        for j, val in enumerate(drow):
                            if j == file_idx:
                                continue
                            col_idx = j if j < file_idx else j - 1
                            ws.cell(row=i + 2, column=col_idx + 1, value=self._safe_str(val))

            # 保存文件
            wb.save(output_path)
            messagebox.showinfo("成功", "成功输出文本到Excel！")

        except Exception as e:
            messagebox.showerror("错误", str(e))

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
    main()
