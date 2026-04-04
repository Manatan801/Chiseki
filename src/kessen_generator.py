"""結線指示票生成モジュール

DXF解析結果から各筆の境界杭番号を時計回り閉合順序リストとして生成する。
平面グラフの面（フェイス）探索アルゴリズムを使用。

アルゴリズム概要:
1. 境界線グラフ（平面グラフ）の各ノードの隣接ノードを角度順にソート
2. 半辺構造を使い「常に左に曲がる」探索で全ての最小サイクル（面）を列挙
3. 重複ノードを持つ面を分割（2-connected でない部分への対処）
4. 各地番テキスト座標が内部にある面を Point-in-Polygon で特定
5. 面積の符号で CW/CCW を判定し、時計回りに統一
6. 北西スコア（Y-X）最大の杭を起点にローテーション
"""

import math
from collections import Counter
from dataclasses import dataclass

import networkx as nx

from src.dxf_parser import ParsedDXF


@dataclass
class KessenResult:
    """結線指示票データ"""
    parcel_number: str          # 地番
    stake_sequence: list[str]   # 杭番号の順序リスト（最後は最初と同じ = 閉合）
    land_use: str | None        # 地目
    owner: str | None           # 所有者名


# === 内部ユーティリティ ===

def _compute_angle(cx: float, cy: float, nx_: float, ny: float) -> float:
    """ノード(cx,cy)から(nx_,ny)への角度を返す（ラジアン）"""
    return math.atan2(ny - cy, nx_ - cx)


def _build_half_edge_structure(graph: nx.Graph) -> dict[str, list[str]]:
    """各ノードの隣接ノードを反時計回り角度順にソートした辞書を構築する。"""
    sorted_neighbors: dict[str, list[str]] = {}
    for node in graph.nodes():
        x = graph.nodes[node]['x']
        y = graph.nodes[node]['y']
        neighbors = list(graph.neighbors(node))
        neighbors.sort(key=lambda n: _compute_angle(
            x, y, graph.nodes[n]['x'], graph.nodes[n]['y']
        ))
        sorted_neighbors[node] = neighbors
    return sorted_neighbors


def _next_half_edge(sorted_neighbors: dict[str, list[str]],
                    from_node: str, to_node: str) -> str:
    """半辺 (from_node -> to_node) の「左折」先を返す。

    to_nodeの隣接リスト(CCW角度順)で from_node の次のノード。
    """
    neighbors = sorted_neighbors[to_node]
    idx = neighbors.index(from_node)
    return neighbors[(idx + 1) % len(neighbors)]


def _find_all_faces(graph: nx.Graph) -> list[list[str]]:
    """平面グラフの全ての最小面（フェイス）を列挙する。"""
    sorted_neighbors = _build_half_edge_structure(graph)
    used: set[tuple[str, str]] = set()
    faces: list[list[str]] = []

    for u, v in graph.edges():
        for start_from, start_to in [(u, v), (v, u)]:
            if (start_from, start_to) in used:
                continue
            face: list[str] = []
            cf, ct = start_from, start_to
            limit = graph.number_of_nodes() + 10
            for _ in range(limit):
                if (cf, ct) in used:
                    break
                used.add((cf, ct))
                face.append(cf)
                nt = _next_half_edge(sorted_neighbors, cf, ct)
                cf, ct = ct, nt
                if cf == start_from and ct == start_to:
                    break
            if len(face) >= 3:
                faces.append(face)
    return faces


def _split_face_at_duplicates(face: list[str]) -> list[list[str]]:
    """重複ノードを持つ面を分割する。"""
    counts = Counter(face)
    duplicates = [node for node, cnt in counts.items() if cnt > 1]
    if not duplicates:
        return [face]

    dup = duplicates[0]
    indices = [i for i, n in enumerate(face) if n == dup]
    sub_faces: list[list[str]] = []
    for k in range(len(indices)):
        i = indices[k]
        j = indices[(k + 1) % len(indices)]
        sub = face[i:j] if j > i else face[i:] + face[:j]
        if len(sub) >= 3:
            sub_faces.extend(_split_face_at_duplicates(sub))
    return sub_faces


def _signed_area(face: list[str], graph: nx.Graph) -> float:
    """Shoelace formula で符号付き面積を計算（正=CCW, 負=CW）"""
    area = 0.0
    n = len(face)
    for i in range(n):
        x1 = graph.nodes[face[i]]['x']
        y1 = graph.nodes[face[i]]['y']
        x2 = graph.nodes[face[(i + 1) % n]]['x']
        y2 = graph.nodes[face[(i + 1) % n]]['y']
        area += x1 * y2 - x2 * y1
    return area / 2.0


def _point_in_polygon(px: float, py: float,
                      polygon: list[tuple[float, float]]) -> bool:
    """Ray casting 法による点包含判定。"""
    n = len(polygon)
    inside = False
    j = n - 1
    for i in range(n):
        xi, yi = polygon[i]
        xj, yj = polygon[j]
        if ((yi > py) != (yj > py)) and \
           (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    return inside


def _rotate_to_northwest(face: list[str], graph: nx.Graph) -> list[str]:
    """最も北西の杭を起点にローテーション。

    スコア = Y座標 - X座標（大きいほど北西）で判定。
    """
    best_idx = 0
    best_score = (graph.nodes[face[0]]['y']
                  - graph.nodes[face[0]]['x'])

    for i, node in enumerate(face):
        score = graph.nodes[node]['y'] - graph.nodes[node]['x']
        if score > best_score + 0.001:
            best_score = score
            best_idx = i

    return face[best_idx:] + face[:best_idx]


# === メイン API ===

def generate_kessen(parsed: ParsedDXF,
                    target_parcels: list[str] | None = None) -> list[KessenResult]:
    """結線指示票データを生成する。

    Args:
        parsed: DXF解析結果（ParsedDXF）
        target_parcels: 対象地番リスト（Noneなら全地番）

    Returns:
        各地番の結線指示票データのリスト
    """
    graph = parsed.graph
    if graph is None or graph.number_of_nodes() == 0:
        return []

    parcels = parsed.parcels
    if target_parcels is not None:
        target_set = set(target_parcels)
        parcels = [p for p in parcels if p.number in target_set]

    # 1. 面探索 + 重複ノード分割
    raw_faces = _find_all_faces(graph)
    all_faces: list[list[str]] = []
    for face in raw_faces:
        all_faces.extend(_split_face_at_duplicates(face))

    # 2. CW方向に統一 + 面積でソート（小さい面優先）
    cw_faces: list[tuple[list[str], float]] = []
    for face in all_faces:
        area = _signed_area(face, graph)
        if area > 0:
            face = list(reversed(face))
            area = -area
        if abs(area) > 0:
            cw_faces.append((face, abs(area)))
    cw_faces.sort(key=lambda x: x[1])

    # 3. 各面の重心を計算
    face_centroids: list[tuple[float, float]] = []
    for face, _area in cw_faces:
        cx = sum(graph.nodes[n]['x'] for n in face) / len(face)
        cy = sum(graph.nodes[n]['y'] for n in face) / len(face)
        face_centroids.append((cx, cy))

    # 4. 地番テキスト → 面のマッチング（排他＋重心距離で競合解決）
    #    各地番 → 最小の包含面を見つけ、face_idx → 候補リストを構築
    face_candidates: dict[int, list[tuple[int, float]]] = {}  # fi → [(pi, dist)]

    for pi, parcel in enumerate(parcels):
        for fi, (face, _area) in enumerate(cw_faces):
            polygon = [
                (graph.nodes[n]['x'], graph.nodes[n]['y']) for n in face
            ]
            if _point_in_polygon(parcel.x, parcel.y, polygon):
                cx, cy = face_centroids[fi]
                dist = math.sqrt((parcel.x - cx) ** 2 + (parcel.y - cy) ** 2)
                face_candidates.setdefault(fi, []).append((pi, dist))
                break  # 最小面のみ

    # 各面で重心に最も近い地番を採用（競合解決）
    matched_face_indices: set[int] = set()
    results: list[KessenResult] = []

    for fi, candidates in face_candidates.items():
        # 重心に最も近い地番が勝ち
        candidates.sort(key=lambda x: x[1])
        pi = candidates[0][0]
        parcel = parcels[pi]

        face = cw_faces[fi][0]
        rotated = _rotate_to_northwest(face, graph)
        sequence = rotated + [rotated[0]]
        results.append(KessenResult(
            parcel_number=parcel.number,
            stake_sequence=sequence,
            land_use=parcel.land_use,
            owner=parcel.owner,
        ))
        matched_face_indices.add(fi)

    # 6. どの地番にもマッチしなかった面 → 「地番不明」として生成
    #    （引き出し線で地番テキストが面の外にある狭い区画）
    if cw_faces:
        max_area = max(a for _, a in cw_faces)
        seen_sequences: set[tuple[str, ...]] = set()
        for r in results:
            seen_sequences.add(tuple(r.stake_sequence))

        for fi, (face, area) in enumerate(cw_faces):
            if fi in matched_face_indices:
                continue
            if area >= max_area * 0.9:
                continue  # 外周面をスキップ
            if area < 10.0:
                continue  # 極小面をスキップ

            # 重複面をスキップ
            rotated = _rotate_to_northwest(face, graph)
            sequence = rotated + [rotated[0]]
            seq_key = tuple(sequence)
            if seq_key in seen_sequences:
                continue
            seen_sequences.add(seq_key)

            # 公共用地ラベルが面内にあれば、その名称を採用
            polygon = [
                (graph.nodes[n]['x'], graph.nodes[n]['y']) for n in face
            ]
            parcel_name = '地番不明'
            for pl in parsed.public_lands:
                if _point_in_polygon(pl.x, pl.y, polygon):
                    parcel_name = pl.label
                    break

            results.append(KessenResult(
                parcel_number=parcel_name,
                stake_sequence=sequence,
                land_use=None,
                owner=None,
            ))

    return results


if __name__ == '__main__':
    from src.dxf_parser import parse_dxf

    parsed = parse_dxf('data/input/三瓶 美紀枝 1070-6,1071-2　１班.DXF')
    results = generate_kessen(parsed, ['1070-6', '1069-5', '1069-7', '1070-2'])
    for r in results:
        print(f'{r.parcel_number}: {r.stake_sequence}')
