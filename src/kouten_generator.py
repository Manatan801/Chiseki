"""交点計算指示書生成モジュール

交点杭（マイナス接頭辞の杭番号）ごとに、
基準線の2点と方向点（延長線の起点）を特定する。

アルゴリズム:
1. 交点杭を列挙
2. 各交点杭の隣接ノードから同一直線ペアを検出 → 基準線
   （隣接がマイナス付きなら、さらに先の非交点杭まで辿る）
3. 同一直線ペア以外の方向にある非交点杭 → 方向点（指示書の上部）
4. 基準線の左右は方向点から見た位置関係で決定
"""

import math
from dataclasses import dataclass

from src.dxf_parser import ParsedDXF


@dataclass
class IntersectionResult:
    """交点計算指示書の1レコード"""
    intersection_stake: str    # 交点杭番号（例: "-13.298"）
    baseline_point1: str       # 基準線の左（方向点から見て）
    baseline_point2: str       # 基準線の右（方向点から見て）
    extension_point1: str      # 方向点（指示書の上部）
    extension_point2: str      # 方向点2（通常は空）


def generate_kouten(parsed: ParsedDXF,
                    block_number: int | None = None) -> list[IntersectionResult]:
    """交点計算指示書データを生成"""
    if parsed.graph is None:
        return []

    graph = parsed.graph
    results = []

    # 交点杭をリストアップ（ブロック番号でフィルタ）
    intersection_stakes = [s for s in parsed.stakes if s.is_intersection]
    if block_number is not None:
        prefix = f"-{block_number}."
        intersection_stakes = [s for s in intersection_stakes
                               if s.number and s.number.startswith(prefix)]

    for stake in intersection_stakes:
        node = stake.number
        if node not in graph:
            continue

        neighbors = list(graph.neighbors(node))
        if len(neighbors) < 2:
            continue

        result = _process_intersection(node, neighbors, parsed)
        if result is not None:
            results.append(result)

    return results


def _process_intersection(
    node: str,
    neighbors: list[str],
    parsed: ParsedDXF,
) -> IntersectionResult | None:
    """1つの交点杭を処理し、基準線・方向点を決定する。

    1. 隣接ノードの角度から同一直線ペアを検出 → 基準線
    2. ペア以外の隣接 → 方向点の方向
    3. 各点をマイナスなし杭まで解決
    4. 方向点から見た左右で基準線をソート
    """
    stake = parsed.get_stake_by_number(node)
    if stake is None:
        return None

    # 隣接ノードの方向角度を計算
    angles: dict[str, float] = {}
    for n in neighbors:
        ns = parsed.get_stake_by_number(n)
        if ns is None:
            continue
        angles[n] = math.atan2(ns.y - stake.y, ns.x - stake.x)

    # 同一直線ペアを検出
    collinear_pairs = _find_collinear_pairs(angles)

    # 同一直線ペアが1つ + 孤立点が1つの明確なケースのみ処理
    # （2ペア以上やペアなしは人手で対応）
    if len(collinear_pairs) != 1:
        return None

    pair = collinear_pairs[0]
    paired_nodes = set(pair)
    isolated = [n for n in neighbors if n not in paired_nodes]

    if len(isolated) != 1:
        return None

    # 基準線の2点（非交点杭まで解決）
    bp1 = _resolve_non_intersection(node, pair[0], parsed)
    bp2 = _resolve_non_intersection(node, pair[1], parsed)

    # 方向点（孤立点の方向を非交点杭まで解決）
    dp = _resolve_non_intersection(node, isolated[0], parsed)

    # 方向点から見た左右で基準線をソート
    bp1, bp2 = _sort_baseline_by_direction(
        node, bp1, bp2, dp, parsed
    )

    # 重複除去
    if bp1 and bp1 == bp2:
        bp2 = ""

    return IntersectionResult(
        intersection_stake=node,
        baseline_point1=bp1,
        baseline_point2=bp2,
        extension_point1=dp,
        extension_point2="",
    )


def _resolve_non_intersection(
    intersection_node: str,
    target_node: str,
    parsed: ParsedDXF,
) -> str:
    """交点杭をグラフ上で同方向に辿り、最初の非交点杭まで解決する。"""
    if not target_node or not target_node.startswith('-'):
        return target_node

    graph = parsed.graph
    if graph is None:
        return target_node

    current = intersection_node
    next_node = target_node
    visited = {intersection_node}

    while next_node.startswith('-'):
        visited.add(next_node)
        candidates = [n for n in graph.neighbors(next_node) if n not in visited]
        if not candidates:
            break

        stake_curr = parsed.get_stake_by_number(current)
        stake_next = parsed.get_stake_by_number(next_node)
        if stake_curr is None or stake_next is None:
            break

        direction = math.atan2(
            stake_next.y - stake_curr.y, stake_next.x - stake_curr.x
        )

        best = None
        best_diff = float('inf')
        for n in candidates:
            sn = parsed.get_stake_by_number(n)
            if sn is None:
                continue
            angle = math.atan2(
                sn.y - stake_next.y, sn.x - stake_next.x
            )
            diff = abs(angle - direction)
            if diff > math.pi:
                diff = 2 * math.pi - diff
            if diff < best_diff:
                best_diff = diff
                best = n

        if best is None or best_diff > math.radians(45):
            break

        current = next_node
        next_node = best

    return next_node


def _find_collinear_pairs(
    angles: dict[str, float],
    threshold_deg: float = 20.0,
) -> list[tuple[str, str]]:
    """角度差が約180度のペア（同一直線上）を検出"""
    pairs = []
    nodes = list(angles.keys())
    threshold_rad = math.radians(threshold_deg)

    for i, n1 in enumerate(nodes):
        for n2 in nodes[i + 1:]:
            diff = abs(angles[n1] - angles[n2])
            if diff > math.pi:
                diff = 2 * math.pi - diff
            if abs(diff - math.pi) < threshold_rad:
                pairs.append((n1, n2))

    return pairs


def _sort_baseline_by_direction(
    intersection_node: str,
    bp1: str,
    bp2: str,
    direction_point: str,
    parsed: ParsedDXF,
) -> tuple[str, str]:
    """方向点から交点を見たとき、基準線の左右を決定する。

    方向点→交点のベクトルに対する外積の符号で左右判定。
    正=左、負=右。
    """
    if not bp1 or not bp2 or not direction_point:
        return bp1, bp2

    s_int = parsed.get_stake_by_number(intersection_node)
    s_dir = parsed.get_stake_by_number(direction_point)
    s_b1 = parsed.get_stake_by_number(bp1)
    s_b2 = parsed.get_stake_by_number(bp2)

    if any(s is None for s in [s_int, s_dir, s_b1, s_b2]):
        return bp1, bp2

    # 方向点→交点のベクトル
    dx = s_int.x - s_dir.x
    dy = s_int.y - s_dir.y

    # 交点→基準線点1のベクトル
    vx1 = s_b1.x - s_int.x
    vy1 = s_b1.y - s_int.y

    # 外積: dx*vy1 - dy*vx1 > 0 → bp1は方向点から見て左
    cross = dx * vy1 - dy * vx1

    if cross < 0:
        return bp1, bp2   # bp1が左、bp2が右
    else:
        return bp2, bp1   # bp2が左、bp1が右




if __name__ == '__main__':
    from src.dxf_parser import parse_dxf

    filepath = "data/input/三瓶 美紀枝 1070-6,1071-2　１班.DXF"
    parsed = parse_dxf(filepath)
    results = generate_kouten(parsed)

    print(f"=== 交点計算指示書 ({len(results)}件) ===")
    for r in results:
        print(
            f"交点={r.intersection_stake} "
            f"基準線=({r.baseline_point1}, {r.baseline_point2}) "
            f"方向点=({r.extension_point1}, {r.extension_point2})"
        )
