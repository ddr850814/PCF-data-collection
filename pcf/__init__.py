"""pcf - parse, edit, split, merge, and render Piping Component Files.

Flat public API; scripts and agents should import from here.
"""
# 本地修改：
# render.py:
# 1. _build_bom 跳过无 ITEM-CODE 的 WELD 记录（焊缝不是材料，避免占 BOM 行）；
# 2. BOM 项号/气球固定用行顺序号 1..N（Isogen 惯例），原料号存 code 列
#    （BomLine 新增 code 字段；图面 BOM 描述前缀带料号）；
# 3. 直管(PIPE)气球改标 BOM 项号（原为切割标签 C1/C2 或不标）；
# 4. 管架(SUPPORT)符号优先取 symbols/ 目录下 Isogen Symbol Editor 导出的
#    .ASC 符号库（SUPPORT-DIRECTION 控制朝向）；无符号库时回退内置画法；
# 5. FLOW-ARROW 流向箭头（SKEY=FLOW）：优先取符号库官方符号（沿最近管段
#    定向，FLOW=2 反向），无符号时回退实心小三角；且不进 BOM。
# asc_symbols.py（本地新增）：.ASC 符号库解析器（规则详见
#    symbols/ASC符号格式规则.md）。

from .continuations import Continuation, build_continuation_map
from .cutlist import BranchInfo, CutPiece, EndInfo, build_cutlist
from .dims import (
    Dimension,
    infer_dimensions,
    load_sidecar,
    merge_with_sidecar,
    sidecar_path,
    write_sidecar,
)
from .geometry import (
    ConnectivityGraph,
    bounding_box,
    flood_values,
    model_tolerance,
    open_ends,
)
from .model import AttrLine, Component, PcfDocument, Pipeline, Point, parse_point
from .options import (
    DEFAULT_TARGET,
    OPTION_NUMBERS,
    aggregate_candidates,
    best_split_key,
    build_split_options,
    cohesive_keys,
    combine_file_options,
    count_field_welds,
    distance_drawing_count,
    options_are_trivial,
    recommend_option,
    survey_report,
)
from .parser import load, parse_document, save, write_document
from .project import (
    ARCHIVE_SUFFIXES,
    DEFAULT_WORKSPACE_NAME,
    DRAWINGS_DIR,
    SOURCE_DIR,
    WORKING_DIR,
    ZIP_RULE,
    ProjectConfig,
    archive_file,
    drawings_dir_for,
    find_project_root,
    find_stray_archives,
    init_project,
    is_archive,
    is_sheet_pcf,
    load_config,
    log_to_wiki,
    resolve_src_dir,
    sanitize_sheet_value,
    save_config,
    sheet_parent_stem,
    stray_archive_problem,
    working_dir_for,
)
from .render import (
    BomLine,
    Callout,
    RenderOptions,
    RenderResult,
    SceneItem,
    SceneModel,
    WeldPosition,
    render_pdf,
    render_scene,
)
from .skeys import classify, connection_from_skey, joint_from_skey, valve_subtype
from .tracking import (
    DEFAULT_RECONCILE_TOLERANCE,
    MAP_ITEM_ID_COLUMN,
    AddedWeld,
    CarriedWeld,
    MapItem,
    RecordMatch,
    RecordResult,
    ReconcileResult,
    RemovedWeld,
    match_map_items,
    reconcile_welds,
)
from .splitmerge import (
    KeySurvey,
    MergeError,
    MergeResult,
    SplitError,
    SplitResult,
    default_split_distance,
    merge_documents,
    split_by_distance,
    split_by_key,
    split_document,
    survey_split_keys,
)
from .units import (
    bore_to_inch,
    bore_to_nps,
    coord_to_inch_factor,
    coord_to_mm_factor,
    format_bore,
    format_feet_inches,
    format_length,
    format_mm,
    format_nps,
    is_metric_units,
)
from .welds import Weld, component_spools, derive_welds

__version__ = "0.1.0"

__all__ = [
    "ARCHIVE_SUFFIXES",
    "DEFAULT_RECONCILE_TOLERANCE", "DEFAULT_TARGET", "DEFAULT_WORKSPACE_NAME",
    "DRAWINGS_DIR", "MAP_ITEM_ID_COLUMN", "OPTION_NUMBERS",
    "SOURCE_DIR", "WORKING_DIR", "ZIP_RULE",
    "AddedWeld", "AttrLine", "BomLine", "BranchInfo", "Callout",
    "CarriedWeld", "Component", "ConnectivityGraph", "Continuation", "CutPiece",
    "Dimension", "EndInfo", "KeySurvey", "MapItem", "MergeError", "MergeResult",
    "PcfDocument", "Pipeline", "Point", "ProjectConfig", "RecordMatch",
    "RecordResult", "ReconcileResult", "RemovedWeld", "RenderOptions",
    "RenderResult", "SceneItem", "SceneModel", "SplitError", "SplitResult",
    "Weld", "WeldPosition",
    "aggregate_candidates", "archive_file", "best_split_key",
    "bore_to_inch", "bore_to_nps", "bounding_box",
    "build_continuation_map",
    "build_cutlist", "build_split_options", "classify", "cohesive_keys",
    "combine_file_options", "component_spools", "connection_from_skey",
    "count_field_welds",
    "coord_to_inch_factor", "coord_to_mm_factor",
    "default_split_distance", "derive_welds", "distance_drawing_count",
    "drawings_dir_for", "find_project_root", "find_stray_archives",
    "flood_values",
    "format_bore", "format_feet_inches", "format_length", "format_mm",
    "format_nps", "infer_dimensions", "init_project", "is_archive",
    "is_metric_units", "is_sheet_pcf",
    "joint_from_skey",
    "load", "load_config", "load_sidecar", "log_to_wiki",
    "match_map_items", "merge_documents", "merge_with_sidecar",
    "model_tolerance", "open_ends", "options_are_trivial",
    "parse_document", "parse_point",
    "recommend_option", "reconcile_welds", "render_pdf", "render_scene",
    "resolve_src_dir",
    "sanitize_sheet_value", "save",
    "save_config", "sheet_parent_stem", "sidecar_path",
    "split_by_distance", "split_by_key", "split_document",
    "stray_archive_problem",
    "survey_report", "survey_split_keys", "valve_subtype",
    "working_dir_for", "write_document", "write_sidecar",
]
