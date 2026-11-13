"""Isogen Symbol Editor .ASC 符号库加载器

格式（由 1.asc 逆向解码，经 01HG/HANG/GUID 等符号几何验证）：
- 501 记录：`501 <SKEY...> <size> <origin_x> <origin_y> ...`
  一个符号可挂多个 SKEY（如 `501 SLVE 01HG ...`）
- 502 记录：`502 <el#> <x> <y> <pen> <x> <y> <pen> ...`
  画笔语义（绘图仪模型）：
    pen 1 = 空移（不画线）   pen 2 = 画线   pen 6 = 画线   pen 0 = 停止/空移
  关键规则：
    - 笔位在记录间连续传递（上一条记录停在哪，下一条从哪开始）
    - 每条 502 首点的 pen 只用于"到达该点"的段（从上一记录末点画过来）；
      从首点出发的段默认空移
    - 符号初始笔位 = 原点 (origin_x, origin_y)，第一条记录首点即从原点到达
- 999：符号结束

符号目录发现顺序：环境变量 PCF_SYMBOL_DIR → 可执行文件/包所在目录的
symbols/ → 当前目录 symbols/。目录下所有 .asc 都会加载（SKEY 先到先得）。
"""

from __future__ import annotations

import json
import os
import re
import sys
from pathlib import Path

_DRAW_PENS = {2, 6}      # 画线的 pen 号
_CACHE: dict | None = None
_JSON_CACHE: dict | None = None


_OP_PAIRS: dict[str, str] = {}   # 阀体→阀柄 skey 配对（501 行第 2 名，如 VB→03SP）


def _parse_asc(text: str) -> dict[str, tuple]:
    """解析一个 .asc 文本，返回 {SKEY: (segments, orient)}

    orient = 501 头部最后一个数字（方向模式）：
      0 = 顺管道方向旋转（如 01HG）；2 = 固定方向（配合 SUPPORT-DIRECTION 上下翻转）；
      3 = 三轴同画（如 ANCH，X/Y/Z 三个方向各画一次）

    画笔语义（经 01HG/ANCH/HANG/SPRG 等符号几何反推验证）：
      - 记录之间笔位连续；第一条 502 记录冷启动于其首点（= 符号连接点，无到达线）
      - 后续记录首点的 pen 作用于"到达"段：pen 1/2/6 均画线
      - 其余点的 pen 作用于"离开"段：pen 2 画线；pen 1 = 空移；pen 6 = 不画（只画到达）；
        pen 0 = 停止
    """
    symbols: dict[str, tuple] = {}
    segs: list[tuple] | None = None
    origin = (0.0, 0.0)
    pos = (0.0, 0.0)
    skeys: list[str] = []
    first_pt = None      # 连接点 = 第一条 502 的首点（画笔起始位置）
    stop_pt = None       # 收笔点 = 最后一条 502 的最后一个真实点（EXIT 接点）
    orient = 0
    rec_idx = 0

    for line in text.splitlines():
        t = line.split()
        if not t:
            continue
        rec = t[0]
        if rec == "501":
            _flush(symbols, skeys, segs, first_pt, orient, stop_pt)
            skeys, rest = [], t[1:]
            while rest and re.search(r"[A-Za-z]", rest[0]):
                skeys.append(rest.pop(0).upper())
            try:
                nums = [float(v) for v in rest]
            except ValueError:
                nums = []
            origin = (nums[1] if len(nums) > 2 else 0.0,
                      nums[2] if len(nums) > 2 else 0.0)
            orient = int(nums[-1]) if nums and nums[-1] in (0.0, 2.0, 3.0) else 0
            pos = origin
            first_pt = None
            stop_pt = None
            segs = []
            rec_idx = 0
        elif rec == "502" and segs is not None:
            vals = t[2:]
            trips = []
            for i in range(0, len(vals) - 2, 3):
                try:
                    trips.append((float(vals[i]), float(vals[i+1]),
                                  int(float(vals[i+2]))))
                except ValueError:
                    break
            # 剔除记录尾部 (0,0) 填充三元组——仅限"跟在 pen0 停笔之后"的：
            # FLWN 记录1 的末点就是真实点 (0,0,pen0)（底边终点），不能删
            while (len(trips) >= 2 and trips[-1][0] == 0.0
                   and trips[-1][1] == 0.0 and trips[-2][2] == 0):
                trips.pop()
            if trips:
                stop_pt = (trips[-1][0], trips[-1][1])
            rec_idx += 1
            for idx, (x, y, pen) in enumerate(trips):
                if idx == 0:
                    if first_pt is None:
                        first_pt = (x, y)   # 符号连接点（画笔起始位置）
                    # 首点到达段：pen 1/2/6 均画线（第一条记录冷启动，无到达段）
                    if rec_idx > 1 and pen in (1, 2, 6):
                        segs.append((pos[0], pos[1], x, y))
                    pos = (x, y)
                    # 首点离开段：pen 2 画向下一点（pen 6 只画到达，pen 1 空移）
                    if pen == 2 and len(trips) > 1:
                        segs.append((x, y, trips[1][0], trips[1][1]))
                    continue
                # 其余点：pen 作用于离开段 —— 只有 pen 2 画向下一点
                if pen == 2 and idx + 1 < len(trips):
                    nx, ny = trips[idx + 1][0], trips[idx + 1][1]
                    segs.append((x, y, nx, ny))
                # pen 1 = 空移；pen 6 = 只画到达；pen 0 = 停止
                pos = (x, y)
        elif rec == "999":
            _flush(symbols, skeys, segs, first_pt, orient, stop_pt)
            segs = None
            skeys = []
    _flush(symbols, skeys, segs, first_pt, orient, stop_pt)  # 文件结束时兜底落盘
    return symbols


def _flush(symbols, skeys, segs, first_pt, orient, stop_pt=None) -> None:
    """把当前累积的符号按 SKEY 存入（先到先得），归一到连接点（首条 502 首点）"""
    if not skeys or not segs:      # 空段符号（如 vendor FLRC）不入库，留给原生兜底
        return
    ox, oy = first_pt if first_pt else (0.0, 0.0)
    norm = [(x1-ox, y1-oy, x2-ox, y2-oy) for x1, y1, x2, y2 in segs]
    stop = ((stop_pt[0]-ox, stop_pt[1]-oy) if stop_pt else (0.0, 0.0))
    for k in skeys:
        symbols.setdefault(k, (norm, orient, stop))
    if len(skeys) > 1 and re.fullmatch(r"\d{1,2}SP", skeys[1]):
        # 厂商阀体→阀柄配对：501 行第 2 个名字是该阀门的操作件（柄）skey，
        # 柄几何另有独立 N**SP 块（如 `501 VB 03SP` → VB** 挂杠杆 03SP）
        _OP_PAIRS.setdefault(skeys[0], skeys[1])


def _candidate_dirs() -> list[Path]:
    dirs = []
    env = os.environ.get("PCF_SYMBOL_DIR")
    if env:
        dirs.append(Path(env))
    # 可执行文件旁（打包 exe：dist\symbols）
    exe_base = Path(sys.executable if getattr(sys, "frozen", False)
                    else sys.argv[0]).resolve().parent
    dirs.append(exe_base / "symbols")
    # 包所在目录旁（开发环境：工作区\symbols；打包后：_MEIPASS\symbols）
    dirs.append(Path(__file__).resolve().parent.parent / "symbols")
    dirs.append(Path.cwd() / "symbols")
    return dirs


def _load_json_symbols() -> dict:
    """skeys.json 官方库（SystemMaterialData.xml 批量解码）作为第二几何源。

    只收录"腿全部干净"的符号（每条腿 4 个 0..10000 整型坐标，≥2 条）；
    混合/空腿符号跳过——领结类组合形状的几何不在腿数据里，只能走 ASC。
    腿与收笔点均按 START 接点归一（START=(0,0)），与 ASC 格式一致。
    """
    global _JSON_CACHE
    if _JSON_CACHE is not None:
        return _JSON_CACHE
    out: dict = {}
    for d in _candidate_dirs():
        f = d / "skeys.json"
        if not f.is_file():
            continue
        try:
            data = json.loads(f.read_text(encoding="utf-8")).get("symbols", {})
        except Exception:
            continue
        for k, v in data.items():
            if not v:
                continue
            # FL** 的腿数据经渲染验证语义错位（画出共起点扇形线而非平行
            # 法兰面线）——腿块语法未完全校准的已知风险，排除并回退手写
            if k == "FL**":
                continue
            legs = v.get("legs") or []
            if len(legs) < 2 or not all(
                    len(l) == 4 and all(isinstance(c, int) and 0 <= c <= 10000
                                        for c in l) for l in legs):
                continue
            pts = v.get("keypoints") or []
            start = next((p for p in pts if p.get("type") == 0), None)
            exitp = next((p for p in pts if p.get("type") == 1), None)
            if (not start or not exitp or start.get("x") is None
                    or exitp.get("x") is None):
                continue
            sx, sy = start["x"], start["y"]
            segs = [(l[0]-sx, l[1]-sy, l[2]-sx, l[3]-sy) for l in legs]
            out.setdefault(k, (segs, int(v.get("orientation") or 0),
                               (exitp["x"]-sx, exitp["y"]-sy)))
        break
    _JSON_CACHE = out
    return out


def load_symbols(force: bool = False) -> dict[str, tuple]:
    """加载并缓存所有符号目录下的 .asc 文件：{SKEY: (segments, orient)}"""
    global _CACHE
    if _CACHE is not None and not force:
        return _CACHE
    _OP_PAIRS.clear()
    merged: dict[str, tuple] = {}
    seen = False
    for d in _candidate_dirs():
        if not d.is_dir():
            continue
        for f in sorted(d.glob("*.asc")):
            try:
                merged.update({k: v for k, v in _parse_asc(
                    f.read_text(errors="replace")).items() if k not in merged})
                seen = True
            except OSError:
                continue
    _CACHE = merged if seen else {}
    return _CACHE


def find_symbol(skey: str):
    """按 SKEY 查符号；返回 (segments, orient)，未加载/未找到返回 None"""
    full = find_symbol_full(skey)
    return full[:2] if full else None


def find_symbol_full(skey: str):
    """按 SKEY 查符号；返回 (segments, orient, stop)，未加载/未找到返回 None。

    stop = 收笔点（EXIT 接点）在归一化符号空间的坐标，START 恒为 (0,0)。
    查找顺序（SKEY 组合规则：前两位基形 + 后两位端部形式，如 VTPL）：
      精确 SKEY → 前两位+"**" 模板 → 前两位基形
    """
    if not skey:
        return None
    key = skey.strip().upper()
    cands = [key, key[:2] + "**"]
    if len(key) > 2:
        cands.append(key[:2])
    lib = load_symbols()
    for k in cands:                      # 第一优先：ASC 符号库（用户可控）
        if k in lib:
            return lib[k]
    js = _load_json_symbols()            # 第二优先：skeys.json 官方库（干净腿）
    for k in cands:
        if k in js:
            return js[k]
    return None


def operator_skey_of(skey: str | None) -> str | None:
    """查厂商 501 行配对的阀柄 skey（`501 VB 03SP ...` → VB** 的柄=03SP）。

    与 find_symbol_full 相同的三级候选：精确 SKEY → 前两位+"**" → 前两位；
    配对信息只在 ASC 符号库（skeys.json 几何源无柄概念）。未命中返回 None。
    """
    if not skey:
        return None
    key = skey.strip().upper()
    cands = [key, key[:2] + "**"]
    if len(key) > 2:
        cands.append(key[:2])
    load_symbols()
    for k in cands:
        if k in _OP_PAIRS:
            return _OP_PAIRS[k]
    return None
