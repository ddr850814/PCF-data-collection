"""NTS 1:1 自动分页装箱（纯数据规划，不依赖渲染器）。

分页优先沿管线邻接链划分连续区间：每页是链上连续的一段，小支路跟随
主链节点、超窗长支路拆成独立连续链，页联合盒不得超过可用窗口；在最少
可行页数下用动态规划使各页零件数尽量均匀。单件自身超窗等不能按整件
划分的情形，回退到窗口网格切片（长直管可在穿越点切开）。BOM 行数超容量
仅警告（绘制侧截断兜底）。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from itertools import product

Vec3 = tuple[float, float, float]

# 打包图面上的最小切分段长：小于它的段不切管（避免 12mm 碎段两头挂满
# 接续标注/尺寸），跨格接头改走组件接续（outs 里 pipe_idx=-1 的条目）。
_MIN_SEG_MM = 18.0


@dataclass
class SheetInfo:
    """测量遍产物（由 render._measure_chain 构建），对齐组件文档序。"""
    boxes: list[tuple[float, float, float, float] | None]  # 打包图面盒(含余量)
    pts: list[list[Vec3]]          # 每件已映射世界点
    pipes: list[bool]              # 直管切点候选
    pipe_axes: list[tuple[Vec3, Vec3] | None]  # 直管轴向两端(世界坐标)
    bom_keys: list[tuple | None]   # BOM 行 key（None=不进 BOM）
    pos: dict[Vec3, tuple[float, float]] = field(default_factory=dict)  # 打包位表


@dataclass
class PagePlan:
    cell: tuple[int, int]         # 唯一页标签：网格坐标或链式页序 (page, 0)
    # 本页组件原索引（文档序）；被切管可出现在多页，未切件只出现一次
    comp_ids: list[int] = field(default_factory=list)
    # 本页内的被切管段：(pipe_idx, 段端 a, 段端 b)（世界坐标）
    pieces: list[tuple[int, Vec3, Vec3]] = field(default_factory=list)
    # 本页边界断口：(切点, 邻格 cell, pipe_idx, inboard 本侧段另一端世界点)
    outs: list[tuple[Vec3, tuple[int, int], int, Vec3]] = field(default_factory=list)


def _union(a, b):
    if a is None:
        return b
    if b is None:
        return a
    return (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))


def _point_clusters(entries: list[tuple[int, Vec3]], tol: float
                    ) -> list[list[int]]:
    """连接点容差聚类：容差桶 + 3×3×3 邻域把彼此距离 ≤ tol 的点归簇，
    返回簇 → entry 索引列表。"""
    buckets: dict[tuple[int, int, int], list[int]] = {}
    cluster_of: dict[int, int] = {}
    clusters: list[list[int]] = []
    for ei, (_i, w) in enumerate(entries):
        bk = tuple(math.floor(v / tol) for v in w)
        cid = None
        for dk in product((-1, 0, 1), repeat=3):
            nb = (bk[0] + dk[0], bk[1] + dk[1], bk[2] + dk[2])
            for ej in buckets.get(nb, []):
                if math.dist(entries[ej][1], w) <= tol:
                    cid = cluster_of[ej]
                    break
            if cid is not None:
                break
        if cid is None:
            cid = len(clusters)
            clusters.append([])
        clusters[cid].append(ei)
        cluster_of[ei] = cid
        buckets.setdefault(bk, []).append(ei)
    return clusters


def _fits(box: tuple[float, float, float, float], uw: float, uh: float) -> bool:
    """页联合盒能否原样装入可用窗口。"""
    return box[2] - box[0] <= uw + 1e-6 and box[3] - box[1] <= uh + 1e-6


def _main_chain_groups(members: set[int], adj: dict[int, set[int]]) -> list[list[int]]:
    """取一棵确定性生成树的长链，支路子树整体并入最近的主链节点。

    即使法兰/垫片/焊点在同一接点形成小环，生成树边也都是真实连接；
    因此主链的任意连续区间连同随附支路，在原件邻接图上仍是一整段。
    """
    def bfs(seed: int):
        dist = {seed: 0}
        parent = {seed: None}
        queue = [seed]
        for cur in queue:
            for nb in sorted(adj[cur]):
                if nb not in members or nb in dist:
                    continue
                dist[nb] = dist[cur] + 1
                parent[nb] = cur
                queue.append(nb)
        return dist, parent

    # 双 BFS 取长链。并列时取文档序较小者，保证每次出图顺序稳定。
    first = min(members)
    dist0, _ = bfs(first)
    a = max(members, key=lambda i: (dist0[i], -i))
    dist_a, parent = bfs(a)
    b = max(members, key=lambda i: (dist_a[i], -i))
    path = []
    cur = b
    while cur is not None:
        path.append(cur)
        cur = parent[cur]
    path.reverse()                       # a → b
    on_path = set(path)

    children = {i: [] for i in members}
    for i, p in parent.items():
        if p is not None:
            children[p].append(i)

    groups: list[list[int]] = []
    for main in path:
        group = [main]
        stack = [i for i in children[main] if i not in on_path]
        while stack:
            cur = stack.pop()
            group.append(cur)
            stack.extend(children[cur])
        groups.append(sorted(group))
    return groups


def _decompose_groups(groups: list[list[int]], adj: dict[int, set[int]],
                      boxes: list, uw: float, uh: float):
    """支路自身超窗时拆成独立连续链，主链两侧也分别成段。

    返回若干条可独立做区间 DP 的链。若只含一个超窗单件，则无法靠
    整件重排解决，返回 None 让网格路径在直管内部落刀。
    """
    def group_box(group):
        box = None
        for i in group:
            box = _union(box, boxes[i])
        return box

    oversized = next((gi for gi, group in enumerate(groups)
                      if not _fits(group_box(group), uw, uh)), None)
    if oversized is None:
        return [groups]

    nodes = set(groups[oversized])
    rebuilt = _main_chain_groups(nodes, adj)
    if len(rebuilt) == 1 and set(rebuilt[0]) == nodes:
        return None             # 单件/不可再分结构，避免无进展递归
    if min(rebuilt[0]) > min(rebuilt[-1]):
        rebuilt.reverse()

    sequences = []
    for part in (groups[:oversized], rebuilt, groups[oversized + 1:]):
        if not part:
            continue
        split = _decompose_groups(part, adj, boxes, uw, uh)
        if split is None:
            return None
        sequences.extend(split)
    return sequences


def _partition_groups(groups: list[list[int]], boxes: list,
                      uw: float, uh: float) -> list[tuple[int, int]] | None:
    """连续区间 DP：先用最少页数装下，再最小化各页零件数方差。"""
    weights = [len(g) for g in groups]
    m = len(groups)
    group_boxes = []
    for group in groups:
        box = None
        for i in group:
            box = _union(box, boxes[i])
        if box is not None and not _fits(box, uw, uh):
            return None             # 单节点连同其支路已超窗，交给切管回退
        group_boxes.append(box)

    # ends[i] = 从第 i 组开始且装得进一页的所有右端点（左闭右开）。
    ends: dict[int, list[int]] = {}
    for i in range(m):
        box = None
        for j in range(i + 1, m + 1):
            box = _union(box, group_boxes[j - 1])
            if box is not None and not _fits(box, uw, uh):
                break               # 区间再向右扩，联合盒不可能重新变小
            ends.setdefault(i, []).append(j)
        if i not in ends:
            return None

    prefix = [0]
    for weight in weights:
        prefix.append(prefix[-1] + weight)
    total = prefix[-1]

    for page_count in range(1, m + 1):
        target = total / page_count
        # end -> (累计方差, 区间路径)；按层迭代可控制恰好 page_count 页。
        previous = {0: (0.0, ())}
        for page_no in range(page_count):
            remaining = page_count - page_no - 1
            current = {}
            for start, (cost, route) in previous.items():
                max_end = m - remaining
                for end in ends.get(start, []):
                    if end > max_end:
                        break
                    count = prefix[end] - prefix[start]
                    new_cost = cost + (count - target) ** 2
                    new_route = route + ((start, end),)
                    old = current.get(end)
                    if old is None or new_cost < old[0] - 1e-9:
                        current[end] = (new_cost, new_route)
            previous = current
            if not previous:
                break
        if m in previous:
            return list(previous[m][1])
    return None


def _chain_plan(info: SheetInfo, usable: tuple[float, float, float, float],
                conn_tol: float):
    """链式连续均匀分页；不能按整件划分时返回 None，由网格切片兜底。"""
    if conn_tol <= 1e-9:
        return None
    boxes = info.boxes
    n = len(boxes)
    present = [i for i in range(n) if boxes[i] is not None]
    if not present:
        return None
    _, _, uw, uh = usable
    whole = None
    for i in present:
        whole = _union(whole, boxes[i])
    if whole is not None and _fits(whole, uw, uh):
        return None                 # 单页路径保持不变

    adj: dict[int, set[int]] = {i: set() for i in present}
    entries = [(i, w) for i in present for w in info.pts[i]]
    for cluster in _point_clusters(entries, conn_tol):
        members = sorted({entries[e][0] for e in cluster})
        for a in members:
            adj[a].update(b for b in members if b != a)

    # 连通分量分别分页，绝不把两条原本不相连的管线拼成一页。
    components: list[set[int]] = []
    rest = set(present)
    while rest:
        seed = min(rest)
        members = {seed}
        queue = [seed]
        rest.remove(seed)
        for cur in queue:
            for nb in adj[cur]:
                if nb not in rest:
                    continue
                rest.remove(nb)
                members.add(nb)
                queue.append(nb)
        components.append(members)

    # 支架等不与管端共点的独立小件不另起一页，而是跟随最近的管线节点；
    # 它们不参与"管线是否连续"的判定，但计入各页零件数。
    main_components = [c for c in components
                       if any(info.pipe_axes[i] is not None for i in c)]
    rider_components = [c for c in components if c not in main_components]
    if not main_components:
        return None
    component_groups: list[list[list[int]]] = []
    node_group: dict[tuple[int, int], int] = {}
    for ci, members in enumerate(main_components):
        groups = _main_chain_groups(members, adj)
        # 主链方向也按两端文档序稳定化，避免每次运行首页对调。
        if min(groups[0]) > min(groups[-1]):
            groups.reverse()
        component_groups.append(groups)
        for gi, group in enumerate(groups):
            for i in group:
                node_group[(ci, i)] = gi

    def members_box(members) -> tuple | None:
        box = None
        for i in members:
            box = _union(box, boxes[i])
        return box

    for riders in sorted(rider_components, key=lambda c: min(c)):
        rider_box = members_box(riders)
        candidates = []
        for ci, members in enumerate(main_components):
            for i in members:
                # 同一组件的所有点都参与距离判断，避免支架对偏心吸附。
                distance = min(math.dist(w, v)
                               for r in riders for w in info.pts[r]
                               for v in info.pts[i])
                candidates.append((distance, ci, i))
        # 管内独立点（流向箭头/支架）优先跟随宿主直管：管中点到两端
        # 等距时，单按最近节点会被页号较小的邻件抢去邻页。
        host_nodes: set[tuple[int, int]] = set()
        if len(riders) == 1 and len(info.pts[next(iter(riders))]) == 1:
            p = info.pts[next(iter(riders))][0]
            for ci, members in enumerate(main_components):
                for i in members:
                    if info.pipe_axes[i] is None or len(info.pts[i]) < 2:
                        continue
                    a, b = info.pts[i][0], info.pts[i][-1]
                    axis = tuple(b[k] - a[k] for k in range(3))
                    lab2 = sum(v * v for v in axis)
                    if lab2 <= 1e-12:
                        continue
                    lab = math.sqrt(lab2)
                    t = sum((p[k] - a[k]) * axis[k]
                            for k in range(3)) / lab2
                    projected = tuple(a[k] + t * axis[k] for k in range(3))
                    if (conn_tol / lab < t < 1 - conn_tol / lab
                            and math.dist(p, projected) <= conn_tol):
                        host_nodes.add((ci, i))
        ordered = sorted(candidates,
                         key=lambda c: ((c[1], c[2]) not in host_nodes,
                                        c[0], c[1], c[2]))
        target = None
        for _distance, ci, i in ordered:
            gi = node_group[(ci, i)]
            group = component_groups[ci][gi]
            candidate = _union(members_box(group), rider_box)
            if candidate is None or _fits(candidate, uw, uh):
                target = (ci, i)
                break
        if target is None:
            # 没有可直接装下的节点时仍跟随最近管线，超窗组随后沿链分解；
            # 不能在此直接退回网格，否则支架会把连续主管线重新拆断。
            _distance, ci, i = ordered[0]
            target = (ci, i)
        ci, i = target
        gi = node_group[(ci, i)]
        component_groups[ci][gi].extend(sorted(riders))
        # 支架不是实际管件，原邻接图中没有边；补一条仅用于分解的挂载边，
        # 保证超窗组重建主链时支架仍随宿主移动，不会成为断开的孤立节点。
        rider_root = min(riders)
        adj[i].add(rider_root)
        adj[rider_root].add(i)

    comp_cell: list[tuple[int, int]] = [(-1, -1)] * n
    pages: list[PagePlan] = []
    for groups in component_groups:
        sequences = _decompose_groups(groups, adj, boxes, uw, uh)
        if sequences is None:
            return None
        for sequence in sequences:
            intervals = _partition_groups(sequence, boxes, uw, uh)
            if intervals is None:
                return None
            for start, end in intervals:
                ids = sorted(i for group in sequence[start:end] for i in group)
                cell = (len(pages), 0) # cell 只需唯一；渲染按本页联合盒落位
                for i in ids:
                    comp_cell[i] = cell
                pages.append(PagePlan(cell, ids))

    # 无图面盒的记录件跟随文档序最近的前置实件，与实际切片归属一致。
    last = (0, 0)
    for i in range(n):
        if boxes[i] is None:
            comp_cell[i] = last
        else:
            last = comp_cell[i]

    # 跨页组件接续。链式方案不切管，所有断点都是组件间的共同连接点。
    by_cell = {page.cell: page for page in pages}
    seen: set[tuple] = set()
    for cluster in _point_clusters(entries, conn_tol):
        members = sorted({entries[e][0] for e in cluster})
        if len(members) < 2:
            continue
        w0 = entries[cluster[0]][1]
        point_key = tuple(round(v, 4) for v in w0)
        for ai in range(len(members)):
            for bi in range(ai + 1, len(members)):
                i, j = members[ai], members[bi]
                ci, cj = comp_cell[i], comp_cell[j]
                if ci == cj:
                    continue
                for me, other, cell_me, cell_ot in (
                        (i, j, ci, cj), (j, i, cj, ci)):
                    key = (cell_me, cell_ot, point_key)
                    if key in seen:
                        continue
                    inb = max(info.pts[me], default=None,
                              key=lambda w: math.dist(w, w0))
                    if inb is None:
                        continue
                    seen.add(key)
                    by_cell[cell_me].outs.append((w0, cell_ot, -1, inb))

    return pages, [], comp_cell, {}, []


def plan(info: SheetInfo,
         usable: tuple[float, float, float, float],
         capacity_rows: int,
         show_cut_table: bool = True,
         conn_tol: float = 0.0,
         ) -> tuple[list[PagePlan], list[tuple[Vec3, int, float]],
                    list[tuple[int, int]],
                    dict[int, list[tuple[float, float, tuple[int, int]]]],
                    list[str]]:
    """连续、不超框优先，零件数均匀次之；异常拓扑回退网格切片。"""
    chained = _chain_plan(info, usable, conn_tol)
    if chained is not None:
        pages, cuts, comp_cell, seg_cells, notes = chained
        for page in pages:
            nkeys = len({info.bom_keys[j] for j in page.comp_ids
                         if info.bom_keys[j] is not None})
            npipe = sum(1 for j in page.comp_ids if info.pipes[j]) \
                if show_cut_table else 0
            if nkeys + npipe > capacity_rows:
                notes.append(
                    f"sheet {page.cell} holds {nkeys + npipe} BOM rows, over the "
                    f"{capacity_rows}-row capacity; table will be truncated")
        return pages, cuts, comp_cell, seg_cells, notes
    return _plan_grid(info, usable, capacity_rows, show_cut_table, conn_tol)


def _plan_grid(info: SheetInfo,
         usable: tuple[float, float, float, float],
         capacity_rows: int,
         show_cut_table: bool = True,
         conn_tol: float = 0.0,
         ) -> tuple[list[PagePlan], list[tuple[Vec3, int, float]],
                    list[tuple[int, int]],
                    dict[int, list[tuple[float, float, tuple[int, int]]]],
                    list[str]]:
    """网格切片 → (页计划(行优先), 切点列表[(世界点, pipe_idx, t)],
    comp_cell(每件归属格), seg_cells(被切管 pipe_idx → [(t0, t1, cell)]),
    notes)。条带装得下单窗口时 cuts/seg_cells 为空。

    conn_tol: 世界单位连接容差——落刀边距取 conn_tol*4（距管端小于它的
    格线穿越不切：端点贴格线切出碎段，且撞 splitmerge 端点容差拒判）；
    同时作为跨格组件接续判定的连接点容差。
    """
    cut_margin = conn_tol * 4
    _, _, uw, uh = usable
    boxes, keys, pipes = info.boxes, info.bom_keys, info.pipes
    pos = info.pos
    notes: list[str] = []
    n = len(boxes)
    present = [i for i in range(n) if boxes[i] is not None]
    if not present:
        return [PagePlan((0, 0), list(range(n)))], [], [(-1, -1)] * n, {}, notes

    x0 = min(boxes[i][0] for i in present)
    y0 = min(boxes[i][1] for i in present)
    x1 = max(boxes[i][2] for i in present)
    y1 = max(boxes[i][3] for i in present)
    ncol = max(1, math.ceil((x1 - x0) / uw - 1e-9))
    nrow = max(1, math.ceil((y1 - y0) / uh - 1e-9))

    def cell_of(px: float, py: float) -> tuple[int, int]:
        c = min(ncol - 1, max(0, int((px - x0) // uw)))
        r = min(nrow - 1, max(0, int((py - y0) // uh)))
        return (r, c)

    # 直管图面/世界两端
    pipe_pp: dict[int, tuple] = {}
    for i in present:
        if pipes[i] and info.pipe_axes[i] is not None:
            pts = info.pts[i]
            if len(pts) >= 2 and pts[0] in pos and pts[1] in pos:
                pipe_pp[i] = (pos[pts[0]], pos[pts[1]], *info.pipe_axes[i])

    # 每根直管按格线切参数区间
    cuts: list[tuple[Vec3, int, float]] = []       # (世界切点, pipe_idx, t)
    # pipe_idx -> [(t0, t1, wa, wb, cell)]
    pipe_segs: dict[int, list[tuple[float, float, Vec3, Vec3, tuple[int, int]]]] = {}
    for i, (pa, pb, a, b) in pipe_pp.items():
        # 边距换算成参数域：距管端 < cut_margin（世界）的穿越不切
        axis_w = math.dist(a, b)
        t_lo = cut_margin / axis_w if axis_w > 1e-12 else 1.0
        ts = {0.0, 1.0}
        for c in range(1, ncol):
            ex = x0 + c * uw
            d = pb[0] - pa[0]
            if abs(d) > 1e-12:
                t = (ex - pa[0]) / d
                if t_lo < t < 1 - t_lo:
                    ts.add(t)
        for r in range(1, nrow):
            ey = y0 + r * uh
            d = pb[1] - pa[1]
            if abs(d) > 1e-12:
                t = (ey - pa[1]) / d
                if t_lo < t < 1 - t_lo:
                    ts.add(t)
        ts_sorted = sorted(ts)
        # 碎段剪枝：打包长 < _MIN_SEG_MM 的段并入邻段（删分界），
        # 整根短管（打包全长不足）自然不切
        pl = math.dist(pa, pb)
        while len(ts_sorted) > 2:
            lens = [(t1 - t0) * pl
                    for t0, t1 in zip(ts_sorted, ts_sorted[1:])]
            k = min(range(len(lens)), key=lambda m: lens[m])
            if lens[k] >= _MIN_SEG_MM:
                break
            if k == 0:
                del ts_sorted[1]
            elif k == len(lens) - 1:
                del ts_sorted[-2]
            elif lens[k - 1] >= lens[k + 1]:
                del ts_sorted[k]       # 并入左邻
            else:
                del ts_sorted[k + 1]   # 并入右邻
        segs = []
        for t0, t1 in zip(ts_sorted, ts_sorted[1:]):
            wa = tuple(a[k] + t0 * (b[k] - a[k]) for k in range(3))
            wb = tuple(a[k] + t1 * (b[k] - a[k]) for k in range(3))
            tm = (t0 + t1) / 2
            cell = cell_of(pa[0] + tm * (pb[0] - pa[0]),
                           pa[1] + tm * (pb[1] - pa[1]))
            segs.append((t0, t1, wa, wb, cell))
        pipe_segs[i] = segs

    # 组件归格
    comp_cell: list[tuple[int, int]] = [(-1, -1)] * n
    for i in present:
        b = boxes[i]
        comp_cell[i] = cell_of((b[0] + b[2]) / 2, (b[1] + b[3]) / 2)

    # 连接簇整体归页：实际相接的组件尽量不跨页——簇的联合盒装得进一格
    # 时整簇归联合盒中心所在格；装不下的大簇保持盒中心归格（由切管 +
    # 接续标注处理跨格）。一件组件常属于多个簇（如弯头/直管两端各接
    # 一串），多簇争抢时按"移动成本"（需改变盒中心归属的成员数）升序
    # 归页且归页即锁定：后续簇的 target 与锁定成员冲突时整簇跳过。
    # 成本低的簇与盒中心默认归属最一致，避免后处理的簇把已归页成员
    # 拉走（BR3R-30001 弯头 68 被含长管 70 的簇以 0.13mm 擦边 target
    # 拉回邻页；西管 52 被西端 2 件小簇抢先锁到邻页）。
    if conn_tol > 1e-9:
        entries = [(i, w) for i in present for w in info.pts[i]]
        origin = comp_cell[:]
        cand: list[tuple[int, int, list[int], tuple[int, int]]] = []
        for cluster in _point_clusters(entries, conn_tol):
            members = sorted({entries[e][0] for e in cluster})
            if len(members) < 2:
                continue
            bx = None
            for m in members:
                bx = _union(bx, boxes[m])
            if bx[2] - bx[0] > uw or bx[3] - bx[1] > uh:
                continue
            target = cell_of((bx[0] + bx[2]) / 2, (bx[1] + bx[3]) / 2)
            cost = sum(1 for m in members if origin[m] != target)
            cand.append((cost, len(members), members, target))
        locked: dict[int, tuple[int, int]] = {}
        for _cost, _sz, members, target in sorted(cand, key=lambda c: (c[0], c[1], c[2])):
            if any(m in locked and locked[m] != target for m in members):
                continue
            for m in members:
                comp_cell[m] = target
                locked[m] = target

    # 无盒件跟随文档序最近的有盒件
    last = (0, 0)
    for i in range(n):
        if boxes[i] is not None:
            last = comp_cell[i]
        else:
            comp_cell[i] = last

    # 空格并段：一格若只装被切管的管段而没有任何完整组件，孤立成页
    # 毫无意义（一段孤管两头挂接续标注）——删分界把空格段并回邻段
    # 所在格。DWC-32165：13.1m 长管中段跨行界，下段 31mm 独占一页。
    if pipe_segs:
        solid = {comp_cell[j] for j in present
                 if len(pipe_segs.get(j, [])) <= 1}
        empty = {s[4] for segs in pipe_segs.values()
                 for s in segs if len(segs) > 1} - solid
        for i in pipe_segs:
            segs = pipe_segs[i]
            if len(segs) < 2:
                continue   # 未跨格管的 comp_cell 已由簇归页定，不碰
            while len(segs) > 1:
                k = next((m for m, s in enumerate(segs)
                          if s[4] in empty), None)
                if k is None:
                    break
                # 并入邻段：端段只有唯一邻；中段优先非空格邻，空格互并
                if k == 0:
                    into = 1
                elif k == len(segs) - 1:
                    into = k - 1
                elif segs[k - 1][4] not in empty:
                    into = k - 1
                elif segs[k + 1][4] not in empty:
                    into = k + 1
                else:
                    into = k - 1
                lo, hi = min(k, into), max(k, into)
                t0, wa = segs[lo][0], segs[lo][2]
                t1, wb = segs[hi][1], segs[hi][3]
                segs[lo:hi + 1] = [(t0, t1, wa, wb, segs[into][4])]
            if len(segs) == 1:
                comp_cell[i] = segs[0][4]   # 整管不再切，归唯一段所在格
        # 段划分已定，统一生成切点
        for i, segs in pipe_segs.items():
            a, b = pipe_pp[i][2], pipe_pp[i][3]
            for si in range(1, len(segs)):
                t = segs[si][0]
                wp = tuple(round(a[k] + t * (b[k] - a[k]), 4)
                           for k in range(3))
                cuts.append((wp, i, t))

    # 每页链连续修复：网格按图面位置分组，U 形回折会把去程+回程留在
    # 一页、连接它们的回折段切到另一页——同页两截互不相连（用户要求
    # 每页只允许连续的一段）。诱导子图恰断成两截、且经另一页完整组件
    # 连接的页，把其中一截并入中介页（联合盒装得进窗口才并）。
    # 跨格被切管视为各段所在页共有，合法切管分页不触发本修复。
    if conn_tol > 1e-9 and (ncol > 1 or nrow > 1):
        adj: dict[int, set[int]] = {i: set() for i in present}
        entries = [(i, w) for i in present for w in info.pts[i]]
        for cluster in _point_clusters(entries, conn_tol):
            mem = sorted({entries[e][0] for e in cluster})
            for a in mem:
                adj[a].update(b for b in mem if b != a)

        def _induced_parts(members: set[int]) -> list[set[int]]:
            parts, rest = [], set(members)
            while rest:
                seed = rest.pop()
                part, stack = {seed}, [seed]
                while stack:
                    for nb in adj[stack.pop()]:
                        if nb in rest:
                            rest.discard(nb)
                            part.add(nb)
                            stack.append(nb)
                parts.append(part)
            return parts

        for _round in range(4):      # 修复可能连锁，限轮次防死循环
            by_cell: dict[tuple[int, int], set[int]] = {}
            for i in present:
                segs = pipe_segs.get(i, [])
                if len(segs) > 1:
                    for s in segs:               # 跨格管各段所在页共有
                        by_cell.setdefault(s[4], set()).add(i)
                else:
                    by_cell.setdefault(comp_cell[i], set()).add(i)
            fixed = False
            for cell, members in sorted(by_cell.items()):
                parts = _induced_parts(members)
                # 只把含直管的分量算作"截"：支架对/孤立符号不接管，
                # 既不计数也不迁移（留在原页）
                main = [p for p in parts
                        if any(info.pipe_axes[j] is not None for j in p)]
                if len(main) != 2:
                    continue
                c1, c2 = main
                # 多源 BFS 找两截间的全图连接路径，取页外中介组件
                prev = {s: None for s in c1}
                queue, qi, target = list(c1), 0, None
                while qi < len(queue) and target is None:
                    cur = queue[qi]
                    qi += 1
                    for nb in adj[cur]:
                        if nb in prev:
                            continue
                        prev[nb] = cur
                        if nb in c2:
                            target = nb
                            break
                        queue.append(nb)
                if target is None:
                    continue        # 两截本就不连通（独立管段组），合法
                path, w = [], target
                while w is not None:
                    path.append(w)
                    w = prev[w]
                mid = [j for j in path if j not in members]
                mid_cells = {comp_cell[j] for j in mid}
                if not mid or len(mid_cells) != 1:
                    continue        # 无中介，或中介跨多页（复杂情形不动）
                mid_cell = next(iter(mid_cells))
                if mid_cell == cell:
                    continue
                # 候选：截 + 中介整页（含支架等叶件），fits 且面积小者优先
                best = None
                for part in (c1, c2):
                    moved = part | by_cell.get(mid_cell, set())
                    bx = None
                    for j in moved:
                        if boxes[j] is not None:
                            bx = _union(bx, boxes[j])
                    if bx is None:
                        continue
                    if bx[2] - bx[0] > uw or bx[3] - bx[1] > uh:
                        continue
                    area = (bx[2] - bx[0]) * (bx[3] - bx[1])
                    if best is None or area < best[0]:
                        best = (area, part)
                if best is None:
                    continue
                for j in best[1]:
                    comp_cell[j] = mid_cell
                fixed = True
                break               # comp_cell 已变，重建分组进入下一轮
            if not fixed:
                break

    if ncol == 1 and nrow == 1:
        page = PagePlan((0, 0), list(range(n)))
        return [page], [], comp_cell, {}, notes

    # 组页
    cells: dict[tuple[int, int], PagePlan] = {}
    for i in range(n):
        if boxes[i] is None:
            continue
        page = cells.setdefault(comp_cell[i], PagePlan(comp_cell[i]))
        page.comp_ids.append(i)
    seg_cells: dict[int, list[tuple[float, float, tuple[int, int]]]] = {}
    for i, segs in pipe_segs.items():
        if len(segs) < 2:
            continue
        seg_cells[i] = [(t0, t1, cell) for t0, t1, _wa, _wb, cell in segs]
        for t0, t1, wa, wb, cell in segs:
            page = cells.setdefault(cell, PagePlan(cell))
            if i not in page.comp_ids:
                page.comp_ids.append(i)
            page.pieces.append((i, wa, wb))
        for si in range(1, len(segs)):
            t = segs[si][0]
            wp = tuple(round(pipe_pp[i][2][k] + t * (pipe_pp[i][3][k] - pipe_pp[i][2][k]), 4)
                       for k in range(3))
            ca, cb = segs[si - 1][4], segs[si][4]
            # inboard = 本侧段的另一端：箭头/幻影以此定向（中段两侧都是切点）
            cells[ca].outs.append((wp, cb, i, segs[si - 1][2]))
            cells[cb].outs.append((wp, ca, i, segs[si][3]))

    pages = [cells[c] for c in sorted(cells)]

    # 跨格组件接续：同一点（连接容差聚类）相接、但 presence 集（归属格
    # + 管段落格）不相交的组件对，两侧页各挂一条 outs（pipe_idx=-1，
    # 无管径/线号引用）——这是短管不切后跨格接头的兜底标注。
    if conn_tol > 1e-9 and len(cells) > 1:
        def presence(i: int) -> set[tuple[int, int]]:
            s = {comp_cell[i]}
            for _t0, _t1, c in seg_cells.get(i, []):
                s.add(c)
            return s

        entries = [(i, w) for i in present for w in info.pts[i]]
        seen: set[tuple] = set()
        for cluster in _point_clusters(entries, conn_tol):
            members = sorted({entries[e][0] for e in cluster})
            if len(members) < 2:
                continue
            w0 = entries[cluster[0]][1]
            pp = pos.get(w0)
            gamma = cell_of(*pp) if pp is not None else None
            for ai in range(len(members)):
                for bi in range(ai + 1, len(members)):
                    i, j = members[ai], members[bi]
                    pi, pj = presence(i), presence(j)
                    if pi & pj:
                        continue
                    for me, other, pme, pother in (
                            (i, j, pi, pj), (j, i, pj, pi)):
                        cell_me = gamma if gamma in pme else comp_cell[me]
                        cell_ot = gamma if gamma in pother else comp_cell[other]
                        if cell_me == cell_ot:
                            continue
                        page = cells.get(cell_me)
                        if page is None:
                            continue
                        # 同页同断点同邻格只挂一条（同一点接多件时
                        # 标注文字相同，重复挂只会叠盒子）
                        key = (cell_me, cell_ot,
                               tuple(round(v, 4) for v in w0))
                        if key in seen:
                            continue
                        seen.add(key)
                        inb = max(info.pts[me], default=None,
                                  key=lambda w: math.dist(w, w0))
                        if inb is None:
                            continue
                        page.outs.append((w0, cell_ot, -1, inb))

    # BOM 行容量检查（仅警告；绘制侧截断兜底）
    for page in pages:
        nkeys = len({keys[j] for j in page.comp_ids if keys[j] is not None})
        npipe = sum(1 for j in page.comp_ids if pipes[j]) if show_cut_table else 0
        if nkeys + npipe > capacity_rows:
            notes.append(
                f"sheet {page.cell} holds {nkeys + npipe} BOM rows, over the "
                f"{capacity_rows}-row capacity; table will be truncated")
    return pages, cuts, comp_cell, seg_cells, notes
