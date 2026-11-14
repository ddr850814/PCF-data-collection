"""PCF 轴测图渲染适配层

基于 pcf2iso（github.com/ProcessQualitySolutions/pcf2iso, MIT）实现：
- 输入 .pcf 文件（或目录），输出单线轴测图 PDF + 焊缝/BOM/切割清单 CSV + 标注 sidecar
- 批量渲染时自动解析跨图纸续接（CONT. ON 标注）
- 可作为模块调用（render_pcf_files），也可命令行使用

命令行示例:
    python pcf_renderer.py D:\\data\\pcf --out-dir D:\\data\\iso --north-axis Y --quadrant NE
"""

from __future__ import annotations

import argparse
import csv
import os
import threading
from pathlib import Path

from pcf import (
    ConnectivityGraph,
    RenderOptions,
    build_continuation_map,
    default_split_distance,
    format_feet_inches,
    format_mm,
    is_metric_units,
    is_sheet_pcf,
    load,
    render_pdf,
    sidecar_path,
    write_sidecar,
)


# ========== CSV 导出（与 pcf2iso scripts/_common.py 一致） ==========

def _welds_csv_path(pdf_path: Path) -> Path:
    return pdf_path.with_name(pdf_path.stem + "_welds.csv")


def _bom_csv_path(pdf_path: Path) -> Path:
    return pdf_path.with_name(pdf_path.stem + "_bom.csv")


def _cuts_csv_path(pdf_path: Path) -> Path:
    return pdf_path.with_name(pdf_path.stem + "_cuts.csv")


def csv_output_encoding(encoding: str) -> str:
    """CSV 写出编码：跟随所选编码；auto（读入时自动识别）→ utf-8-sig。

    utf-8-sig 带 BOM，中文 Excel 双击即可正确显示；显式选了 utf-8 时也
    统一带 BOM 写出；GBK/GB18030 等编码则原样写出，与读入编码体系一致。
    """
    if encoding in ("auto", "utf-8"):
        return "utf-8-sig"
    return encoding


def _write_welds_csv(result, path: Path, encoding: str = "utf-8-sig") -> Path:
    fields = [
        "weld_number", "label", "joint", "nps", "x_mm", "y_mm",
        "page_width_mm", "page_height_mm",
        "world_x", "world_y", "world_z", "pipeline", "description", "uci",
        "neighbors", "spool", "notes", "sheet", "map_item_id",
    ]
    with open(path, "w", newline="", encoding=encoding, errors="replace") as handle:
        writer = csv.writer(handle)
        writer.writerow(fields)
        for weld in result.welds:
            writer.writerow([
                weld.number, weld.label, weld.joint or "", weld.nps,
                weld.x_mm, weld.y_mm,
                result.page_width_mm, result.page_height_mm,
                weld.world_x, weld.world_y, weld.world_z,
                weld.pipeline, weld.description or "", weld.uci,
                weld.neighbors, weld.spool, weld.notes, weld.sheet, "",
            ])
    return path


def _write_bom_csv(result, doc, path: Path, encoding: str = "utf-8-sig") -> Path:
    fields = ["item", "code", "description", "nps", "class", "qty", "length_text",
              "length_in", "weight", "ucis", "sheet"]
    metric = is_metric_units(doc.units_coords)
    with open(path, "w", newline="", encoding=encoding, errors="replace") as handle:
        writer = csv.writer(handle)
        writer.writerow(fields)
        for line in result.bom:
            if line.cls == "PIPE":
                length_text = (format_mm(line.length_in * 25.4) if metric
                               else format_feet_inches(line.length_in))
            else:
                length_text = ""
            writer.writerow([
                line.item, getattr(line, "code", ""), line.description, line.nps,
                line.cls, line.qty,
                length_text, round(line.length_in, 3), round(line.weight, 3),
                line.ucis, line.sheet,
            ])
    return path


def _write_cuts_csv(result, doc, path: Path, encoding: str = "utf-8-sig") -> Path:
    units = doc.units_coords
    fields = ["cut", "item_code", "description", "spool", "pipeline", "nps",
              "cut_length", "cut_length_text", "raw_length", "length_source",
              "direction", "end1_connection", "end1_joins", "end1_weld_gap",
              "end2_connection", "end2_joins", "end2_weld_gap", "uci", "sheet"]
    with open(path, "w", newline="", encoding=encoding, errors="replace") as handle:
        writer = csv.writer(handle)
        writer.writerow(fields)
        for piece in sorted(result.cut_pieces, key=lambda p: -p.cut_length):
            e1, e2 = piece.end1_info, piece.end2_info
            writer.writerow([
                piece.label, piece.item_code or "", piece.description or "",
                piece.spool or "", piece.pipeline, piece.nps,
                round(piece.cut_length, 4), piece.cut_length_text(units),
                round(piece.raw_length, 4), piece.length_source, piece.direction,
                e1.connection if e1 else "", e1.adjacent if e1 else "",
                e1.weld_gap if e1 and e1.weld_gap is not None else "",
                e2.connection if e2 else "", e2.adjacent if e2 else "",
                e2.weld_gap if e2 and e2.weld_gap is not None else "",
                piece.uci or "", piece.sheet,
            ])
    return path


# ========== 渲染入口 ==========

def sources_from_db(db_path):
    """从收集数据库读取导入来源，返回 (源文件夹, [PCF 绝对路径...])，无法定位时返回 None。

    优先读 IMPORT_INFO 表（SOURCE-FOLDER / SOURCE-FILES，新版导入时写入）；
    旧库没有该表时回退：数据库所在目录即源文件夹，递归收集其中的 .pcf。
    """
    import sqlite3

    if not db_path or not os.path.isfile(db_path):
        return None

    folder = None
    rel_files = []
    try:
        conn = sqlite3.connect(db_path)
        try:
            rows = dict(conn.execute("SELECT KEY, VALUE FROM IMPORT_INFO").fetchall())
            folder = rows.get("SOURCE-FOLDER")
            rel_files = [n.strip() for n in (rows.get("SOURCE-FILES") or "").splitlines() if n.strip()]
        finally:
            conn.close()
    except sqlite3.Error:
        folder = None

    if not folder or not os.path.isdir(folder):
        # 旧库回退：收集器默认把数据库建在源文件夹里
        folder = os.path.dirname(os.path.abspath(db_path))
        if not any(folder and os.path.isfile(os.path.join(r, f))
                   for r, _, fs in os.walk(folder) for f in fs if f.lower().endswith(".pcf")):
            return None
        rel_files = []

    if rel_files:
        files = [os.path.join(folder, rel) for rel in rel_files]
        files = [f for f in files if os.path.isfile(f) and f.lower().endswith(".pcf")]
    else:
        files = sorted(os.path.join(r, f)
                       for r, _, fs in os.walk(folder)
                       for f in fs if f.lower().endswith(".pcf"))
    if not files:
        return None
    return folder, files


def collect_pcf_files(paths) -> list[Path]:
    """展开文件/目录参数为排序后的 .pcf 文件列表"""
    out = []
    for arg in paths:
        path = Path(arg)
        if path.is_dir():
            out.extend(sorted(path.glob("*.pcf")))
        elif path.is_file():
            out.append(path)
        else:
            raise FileNotFoundError(f"未找到: {arg}")
    if not out:
        raise FileNotFoundError("目录中没有 .pcf 文件")
    return out


# ========== 预览用整套图纸缓存（续接标注/幻影件需要邻图数据） ==========

_SET_CACHE: dict = {}
_SET_LOCK = threading.Lock()
_SET_CACHE_MAX = 4   # 最多同时保留几个目录的整套数据


def _folder_stamp(folder: str) -> tuple:
    """目录内全部 .pcf 的（文件名, mtime）指纹，任一变动即缓存失效。"""
    try:
        return tuple(sorted(
            (e.name, e.stat().st_mtime_ns)
            for e in os.scandir(folder)
            if e.name.lower().endswith(".pcf") and e.is_file()))
    except OSError:
        return ()


def render_set(pcf_path, encoding="auto"):
    """返回 (neighbour_docs, cont_map)：整套图纸的文档表与续接表。

    预览只画单条管线，但 SEE ISO/SHT/坐标续接标注与分图幻影件必须
    拿到邻图文档才能解析绘制（同 render_pcf_files 的批量路径）。按
    目录把整套文档与续接表缓存起来，首次点击加载全套、之后即时复
    用；目录内任一 .pcf 变动自动重建。预览在 GUI 后台线程触发，用
    锁保证只建一份。
    """
    folder = os.path.dirname(os.path.abspath(str(pcf_path)))
    stamp = _folder_stamp(folder)
    key = (folder, encoding)
    with _SET_LOCK:
        hit = _SET_CACHE.get(key)
        if hit is not None and hit[0] == stamp:
            return hit[1], hit[2]
    loaded = [(p, load(p, encoding=encoding))
              for p in collect_pcf_files([folder])]
    try:
        cont_map = build_continuation_map([(p.stem, d) for p, d in loaded])
    except Exception:
        cont_map = {}   # 续接解析失败不阻塞预览
    docs = {p.stem: d for p, d in loaded}
    stamp = _folder_stamp(folder)
    with _SET_LOCK:
        _SET_CACHE[key] = (stamp, docs, cont_map)
        while len(_SET_CACHE) > _SET_CACHE_MAX:
            _SET_CACHE.pop(next(iter(_SET_CACHE)))
    return docs, cont_map


def render_preview_bytes(pcf_path, north_axis=None, quadrant=None, paper=None,
                         to_scale=False, nts_min_mm=None, nts_max_mm=None,
                         skew_deg=None, encoding="auto"):
    """实时预览：渲染单条管线的纯图面（无图框/标题栏/材料表），返回 PDF 字节。

    供 GUI 管线清单点击时后台调用，不出文件、不带图纸家具。
    """
    from pcf import RenderOptions, load, render_pdf as _render_pdf

    options = RenderOptions(preview=True)
    if north_axis:
        options.north_axis = north_axis
    if quadrant:
        options.quadrant = quadrant
    if paper:
        options.paper = paper
    options.not_to_scale = not to_scale
    if nts_min_mm is not None:
        options.nts_min_mm = float(nts_min_mm)
    if nts_max_mm is not None:
        options.nts_max_mm = float(nts_max_mm)
    if skew_deg is not None:
        options.skew_deg = float(skew_deg)

    docs, cont_map = render_set(pcf_path, encoding=encoding)
    stem = Path(pcf_path).stem
    doc = docs.get(stem) or load(pcf_path, encoding=encoding)
    options.continuations = cont_map.get(stem, [])
    options.neighbour_docs = docs
    import tempfile
    fd, tmp = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    try:
        _render_pdf(doc, tmp, options)
        with open(tmp, "rb") as f:
            return f.read()
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def render_preview_scene(pcf_path, north_axis=None, quadrant=None, paper=None,
                         to_scale=False, nts_min_mm=None, nts_max_mm=None,
                         skew_deg=None, encoding="auto"):
    """实时预览（全矢量）：渲染单条管线的纯图面，返回 (SceneModel, RenderResult)。

    与 render_preview_bytes 同参数体系，但走矢量场景后端——不经临时
    PDF/PyMuPDF 栅格化，GUI 主线程直接回放 SceneModel；RenderResult
    携带焊点页坐标等元数据。
    """
    from pcf import RenderOptions, load, render_scene

    options = RenderOptions(preview=True)
    if north_axis:
        options.north_axis = north_axis
    if quadrant:
        options.quadrant = quadrant
    if paper:
        options.paper = paper
    options.not_to_scale = not to_scale
    if nts_min_mm is not None:
        options.nts_min_mm = float(nts_min_mm)
    if nts_max_mm is not None:
        options.nts_max_mm = float(nts_max_mm)
    if skew_deg is not None:
        options.skew_deg = float(skew_deg)

    docs, cont_map = render_set(pcf_path, encoding=encoding)
    stem = Path(pcf_path).stem
    doc = docs.get(stem) or load(pcf_path, encoding=encoding)
    options.continuations = cont_map.get(stem, [])
    options.neighbour_docs = docs
    return render_scene(doc, options)


def render_pcf_files(pcf_paths, out_dir, north_axis=None, quadrant=None,
                     paper=None, project="", job="", revision=None,
                     notes=None, properties=None, to_scale=False,
                     spool_key=None, detect_ffw=False,
                     nts_min_mm=None, nts_max_mm=None, skew_deg=None,
                     encoding="auto", csv_encoding=None,
                     progress_cb=None, failures=None) -> list[dict]:
    """批量渲染 PCF 为轴测图。

    参数:
        pcf_paths:   .pcf 文件或目录列表（一次性传入以解析跨图续接）
        out_dir:     输出目录
        north_axis:  哪根 PCF 轴是 plant north（Y/X/Y-decreasing/X-decreasing）
        quadrant:    轴测视角象限（NE/NW/SW/SE，重叠时换象限重画）
        paper:       图纸尺寸（tabloid/a3/letter/a4）
        project/job/revision/notes/properties: 标题栏内容
        to_scale:    True=按真实比例；False=示意画法（默认，压缩/拉伸管段便于阅读）
        spool_key:   分段属性名（如 SPOOL-IDENTIFIER）。设置后按该属性值标注分段，
                     跨段焊缝自动判定为现场焊（F.W.）。仅影响图面，不改数据
        detect_ffw:  检测现场调整焊（端面刚好贴合、无焊缝记录的接头）
        encoding:    PCF 读入编码（auto/gbk/utf-8/utf-8-sig/gb18030，
                     auto=GBK 优先，同收集器）
        csv_encoding: CSV 导出编码；默认跟随 encoding（auto/utf-8 → utf-8-sig
                     带 BOM，中文 Excel 双击可读；gbk/gb18030 原样写出）
        progress_cb: 可选回调 cb(msg, current, total)
        failures:    可选 list，单文件解析/渲染失败的 (文件名, 错误简述)
                     会追加进来（单文件失败不中止整批，跳过继续）

    返回: 成功文件的摘要列表 [{pcf, pdf, welds, bom, cuts, warnings}, ...]
    """
    if csv_encoding is None:
        csv_encoding = csv_output_encoding(encoding)
    if failures is None:
        failures = []
    files = collect_pcf_files(pcf_paths)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    total = len(files)

    def report(msg, current):
        if progress_cb:
            progress_cb(msg, current, total)

    # 先全部解析（续接标注需要知道整套图纸的信息）；单文件失败记录后继续
    report("解析 PCF 文件...", 0)
    loaded = []
    for i, path in enumerate(files):
        try:
            report(f"解析 {path.name}", i)
            loaded.append((path, load(path, encoding=encoding)))
        except Exception as e:
            failures.append((path.name, f"解析失败: {e}"))

    # 渲染前断连预检（只读检查，不改数据）：文件包含多个不连通管段组时给出警告
    split_warnings: dict[str, str] = {}
    try:
        for path, doc in loaded:
            if is_sheet_pcf(path):
                continue
            comps = doc.real_components()
            if not comps:
                continue
            distance = default_split_distance(doc.units_coords)
            groups = ConnectivityGraph(comps, distance).connected_sets()
            if len(groups) > 1:
                split_warnings[path.name] = (
                    f"包含 {len(groups)} 个不连通管段组（间距阈值 {distance:g} "
                    f"{doc.units_coords}），建议先拆分再出图，否则多段会挤在同一张图上")
    except Exception:
        pass

    cont_map = {}
    report("解析续接关系...", total)
    try:
        cont_map = build_continuation_map([(p.stem, d) for p, d in loaded])
    except Exception:
        cont_map = {}  # 续接解析失败不阻塞出图

    # 整套文档表：分图幻影件需要邻图数据才能把断点对侧的紧邻零件
    # 画成虚线幻影（见 render.RenderOptions.neighbour_docs）
    neighbour_docs = {p.stem: d for p, d in loaded}

    summaries = []
    for i, (path, doc) in enumerate(loaded):
        try:
            report(f"渲染 {path.name}", i + 1)
            options = RenderOptions()
            if north_axis:
                options.north_axis = north_axis
            if quadrant:
                options.quadrant = quadrant
            if paper:
                options.paper = paper
            options.project = project or ""
            options.job = job or ""
            if revision is not None:
                options.revision = revision
            if notes:
                options.notes = list(notes)
            if properties:
                options.properties = dict(properties)
            options.not_to_scale = not to_scale
            # 非比例伸缩与斜管倾斜的可调参数（GUI 出图设置弹窗用）
            if nts_min_mm is not None:
                options.nts_min_mm = float(nts_min_mm)
            if nts_max_mm is not None:
                options.nts_max_mm = float(nts_max_mm)
            if skew_deg is not None:
                options.skew_deg = float(skew_deg)
            if spool_key:
                options.spool_key = spool_key.strip().upper()
            if detect_ffw:
                options.detect_ffw = True
            options.drawing_no = path.stem
            options.continuations = cont_map.get(path.stem, [])
            options.neighbour_docs = neighbour_docs

            pdf_path = out / (path.stem + ".pdf")
            dims_file = sidecar_path(pdf_path)
            if dims_file.is_file():
                options.dims_sidecar = str(dims_file)  # 已有的人工标注覆写继续生效

            result = render_pdf(doc, pdf_path, options)
            _write_welds_csv(result, _welds_csv_path(pdf_path), encoding=csv_encoding)
            _write_bom_csv(result, doc, _bom_csv_path(pdf_path), encoding=csv_encoding)
            _write_cuts_csv(result, doc, _cuts_csv_path(pdf_path), encoding=csv_encoding)
            write_sidecar(dims_file, result.dimensions, doc.units_coords)

            warnings = list(result.warnings)
            if path.name in split_warnings:
                warnings.insert(0, split_warnings[path.name])
            summaries.append({
                "pcf": str(path), "pdf": str(pdf_path),
                "welds": len(result.welds), "bom": len(result.bom),
                "cuts": len(result.cut_pieces),
                "warnings": warnings,
            })
        except Exception as e:
            # 单文件失败不中止整批：记录 (文件名, 错误简述) 后继续
            failures.append((path.name, f"渲染失败: {e}"))

    report("完成", total)
    return summaries


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="PCF 轴测图批量渲染")
    parser.add_argument("paths", nargs="+", help=".pcf 文件或目录")
    parser.add_argument("--out-dir", default="drawings", help="输出目录（默认 ./drawings）")
    parser.add_argument("--north-axis", choices=["Y", "X", "Y-decreasing", "X-decreasing"],
                        help="哪根 PCF 轴是 plant north（默认 Y）")
    parser.add_argument("--quadrant", choices=["NE", "NW", "SW", "SE"],
                        help="轴测视角象限（默认 SE：北向标指向左上，同 ISOGEN 默认西偏北30°）")
    parser.add_argument("--paper", choices=["tabloid", "a3", "letter", "a4"],
                        help="图纸尺寸（默认 tabloid = 11x17）")
    parser.add_argument("--project", default="", help="标题栏 PROJECT")
    parser.add_argument("--job", default="", help="标题栏 JOB")
    parser.add_argument("--to-scale", action="store_true", help="按真实比例（默认示意画法）")
    parser.add_argument("--spool-key", metavar="KEY",
                        help="分段属性名（如 SPOOL-IDENTIFIER），跨段焊缝自动判定为现场焊 F.W.")
    parser.add_argument("--detect-ffw", action="store_true", help="检测现场调整焊（贴合接头）")
    parser.add_argument("--encoding", default="auto",
                        choices=["auto", "gbk", "utf-8", "utf-8-sig", "gb18030"],
                        help="PCF 读入编码（默认 auto=GBK 优先，同收集器）")
    parser.add_argument("--csv-encoding", default=None,
                        choices=["utf-8", "utf-8-sig", "gbk", "gb18030"],
                        help="CSV 导出编码（默认跟随 --encoding，auto/utf-8 → utf-8-sig）")
    args = parser.parse_args(argv)

    failures = []
    summaries = render_pcf_files(
        args.paths, args.out_dir,
        north_axis=args.north_axis, quadrant=args.quadrant, paper=args.paper,
        project=args.project, job=args.job, to_scale=args.to_scale,
        spool_key=args.spool_key, detect_ffw=args.detect_ffw,
        encoding=args.encoding, csv_encoding=args.csv_encoding,
        progress_cb=lambda msg, cur, tot: print(f"[{cur}/{tot}] {msg}"),
        failures=failures)

    for s in summaries:
        print(f"{Path(s['pcf']).name}: 焊缝{s['welds']} BOM{s['bom']} 切割{s['cuts']} -> {s['pdf']}")
        for w in s["warnings"]:
            print(f"  警告: {w}")
    if failures:
        print(f"失败 {len(failures)} 个文件:")
        for name, err in failures:
            print(f"  {name}: {err}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
