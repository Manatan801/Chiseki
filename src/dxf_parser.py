"""DXF解析基盤モジュール

JWCADからDXFエクスポートされた地籍調査図面を解析し、
杭座標・杭番号・境界線・地番等の情報を構造化データとして抽出する。
"""

import re
import math
from dataclasses import dataclass, field
from typing import Optional

import ezdxf
import networkx as nx


# === データクラス ===

@dataclass
class Stake:
    """杭（境界点）"""
    x: float
    y: float
    number: Optional[str] = None  # 杭番号 (例: "1.777", "-1.832")
    layer: str = ""

    @property
    def is_intersection(self) -> bool:
        """交点杭（マイナス接頭辞）かどうか"""
        return self.number is not None and self.number.startswith('-')

    def distance_to(self, other_x: float, other_y: float) -> float:
        return math.sqrt((self.x - other_x) ** 2 + (self.y - other_y) ** 2)


@dataclass
class BoundaryLine:
    """境界線（2点の線分）"""
    x1: float
    y1: float
    x2: float
    y2: float
    layer: str = ""
    stake1_number: Optional[str] = None
    stake2_number: Optional[str] = None


@dataclass
class ParcelInfo:
    """地番情報"""
    number: str  # 地番 (例: "1070-6")
    x: float     # テキスト位置X
    y: float     # テキスト位置Y
    layer: str = ""
    land_use: Optional[str] = None  # 地目
    owner: Optional[str] = None     # 所有者名


@dataclass
class PublicLandInfo:
    """公共用地ラベル"""
    label: str   # 例: "県道-5", "河川-2"
    x: float
    y: float
    layer: str = ""


@dataclass
class ParsedDXF:
    """DXF解析結果"""
    stakes: list[Stake] = field(default_factory=list)
    lines: list[BoundaryLine] = field(default_factory=list)
    parcels: list[ParcelInfo] = field(default_factory=list)
    public_lands: list[PublicLandInfo] = field(default_factory=list)
    graph: Optional[nx.Graph] = None
    entity_counts: dict = field(default_factory=dict)

    def get_stake_by_number(self, number: str) -> Optional[Stake]:
        for s in self.stakes:
            if s.number == number:
                return s
        return None

    def get_stakes_with_numbers(self) -> list[Stake]:
        return [s for s in self.stakes if s.number is not None]


# === テキスト処理 ===

def clean_mtext(raw_text: str) -> str:
    """MTEXTの書式コードを除去してプレーンテキストを取得"""
    text = re.sub(r'\\[WwHhAaCcFfPpQqTtLlOo][^;]*;', '', raw_text)
    text = re.sub(r'[{}]', '', text)
    text = re.sub(r'\\P', '\n', text)
    text = re.sub(r'\\[^;]*;', '', text)
    return text.strip()


def classify_text(raw_text: str) -> tuple[str, str]:
    """MTEXTを分類する

    Returns:
        (分類, クリーンテキスト)
        分類: 'stake' | 'intersection_stake' | 'survey_code' |
              'parcel_number' | 'public_land' | 'land_use' | 'other'
    """
    text = clean_mtext(raw_text)

    # 複数行の場合は最初の行で判定
    if '\n' in text:
        text = text.split('\n')[0].strip()

    # 空テキスト
    if not text:
        return 'other', text

    # 交点杭番号（マイナス接頭辞 + 小数）
    if re.match(r'^-\d+\.\d+$', text):
        return 'intersection_stake', text

    # 通常の杭番号（小数）
    if re.match(r'^\d+\.\d+$', text):
        return 'stake', text

    # 筆界点コード（P3P3-Fxxxx-x形式）
    if text.startswith('P3P3-'):
        return 'survey_code', text

    # 地番（整数 or 整数-整数）
    if re.match(r'^\d+(-\d+)?$', text):
        return 'parcel_number', text

    # 公共用地ラベル
    if re.match(r'^(県道|河川|市道|国道|水|道)-', text):
        return 'public_land', text

    # 地目
    land_use_types = {'宅地', '山林', '雑種地', '公衆用道路', '田', '畑',
                      '原野', '墓地', '境内地', '学校用地', '鉄道用地',
                      '池沼', '牧場', '保安林', 'ため池'}
    if text in land_use_types:
        return 'land_use', text

    return 'other', text


# === DXF解析メイン ===

def parse_dxf(filepath: str, stake_match_threshold: float = 5.0,
              line_match_threshold: float = 0.5) -> ParsedDXF:
    """DXFファイルを解析して構造化データを返す

    Args:
        filepath: DXFファイルパス
        stake_match_threshold: 杭番号テキスト↔CIRCLE中心のマッチング閾値
        line_match_threshold: 線分端点↔杭座標のマッチング閾値
    """
    # JW-CAD等で生成されたDXFの破損にも対応するため recover を使う
    try:
        doc = ezdxf.readfile(filepath)
    except ezdxf.DXFStructureError:
        from ezdxf import recover
        doc, _auditor = recover.readfile(filepath)
    except Exception:
        # TypeError等のパースエラーもrecoverで再試行
        from ezdxf import recover
        doc, _auditor = recover.readfile(filepath)
    msp = doc.modelspace()
    result = ParsedDXF()

    # エンティティタイプ別カウント（診断用）
    entity_counts = {}
    for e in msp:
        t = e.dxftype()
        entity_counts[t] = entity_counts.get(t, 0) + 1
    result.entity_counts = entity_counts

    # Step 1: 全CIRCLE(r=0.25)の中心座標を杭として抽出
    stakes_by_pos = {}  # (round_x, round_y) -> Stake
    for entity in msp.query('CIRCLE'):
        if abs(entity.dxf.radius - 0.25) < 0.01:
            c = entity.dxf.center
            stake = Stake(x=c.x, y=c.y, layer=entity.dxf.layer)
            # 重複排除（同一座標は1つにまとめる）
            key = (round(c.x, 3), round(c.y, 3))
            if key not in stakes_by_pos:
                stakes_by_pos[key] = stake
    result.stakes = list(stakes_by_pos.values())

    # Step 2: 全MTEXTを抽出・分類
    stake_texts = []     # (杭番号, x, y, layer)
    parcel_texts = []    # (地番, x, y, layer)
    land_use_texts = []  # (地目, x, y, layer)
    public_land_texts = []
    owner_texts = []

    for entity in msp.query('MTEXT'):
        raw = entity.text
        category, text = classify_text(raw)
        pos = entity.dxf.insert

        if category in ('stake', 'intersection_stake'):
            stake_texts.append((text, pos.x, pos.y, entity.dxf.layer))
        elif category == 'parcel_number':
            parcel_texts.append((text, pos.x, pos.y, entity.dxf.layer))
        elif category == 'land_use':
            land_use_texts.append((text, pos.x, pos.y, entity.dxf.layer))
        elif category == 'public_land':
            public_land_texts.append(PublicLandInfo(
                label=text, x=pos.x, y=pos.y, layer=entity.dxf.layer
            ))

    # Layer 53の所有者名を取得
    for entity in msp.query('MTEXT'):
        if entity.dxf.layer == '53':
            text = clean_mtext(entity.text)
            if text and not re.match(r'^[\d.-]+$', text):
                pos = entity.dxf.insert
                owner_texts.append((text, pos.x, pos.y))

    result.public_lands = public_land_texts

    # Step 3: 杭番号テキスト → 最近傍CIRCLE中心にマッチング
    _match_stake_numbers(result.stakes, stake_texts, stake_match_threshold)

    # Step 4: 地番情報を構築
    for pnum, px, py, player in parcel_texts:
        parcel = ParcelInfo(number=pnum, x=px, y=py, layer=player)
        # 最近傍の地目テキストをマッチング
        best_dist = 15.0
        for ltext, lx, ly, _ in land_use_texts:
            d = math.sqrt((px - lx) ** 2 + (py - ly) ** 2)
            if d < best_dist:
                best_dist = d
                parcel.land_use = ltext
        # 最近傍の所有者名をマッチング
        best_dist = 30.0
        for otext, ox, oy in owner_texts:
            d = math.sqrt((px - ox) ** 2 + (py - oy) ** 2)
            if d < best_dist:
                best_dist = d
                parcel.owner = otext
        result.parcels.append(parcel)

    # Step 5: LWPOLYLINE(2頂点)を境界線として抽出
    for entity in msp.query('LWPOLYLINE'):
        points = list(entity.get_points(format='xy'))
        if len(points) == 2:
            x1, y1 = points[0]
            x2, y2 = points[1]
            result.lines.append(BoundaryLine(
                x1=x1, y1=y1, x2=x2, y2=y2,
                layer=entity.dxf.layer
            ))

    # Step 6: 線分端点 → 杭座標をマッチングし、グラフ構築
    result.graph = _build_boundary_graph(result.stakes, result.lines,
                                         line_match_threshold)

    return result


def _match_stake_numbers(stakes: list[Stake],
                         stake_texts: list[tuple[str, float, float, str]],
                         threshold: float) -> None:
    """杭番号テキストを最近傍のCIRCLE中心にマッチング（ハンガリアン法）

    全(テキスト, CIRCLE)ペアの距離行列を構築し、scipy.optimize.linear_sum_assignment
    で総距離を最小化する最適マッチングを求める。これにより:
    - 複数テキストが同一CIRCLEを争う場合、全体最適な割当が行われる
    - あるテキストが遠いCIRCLEしかなくても、他のテキストに代替があれば譲られる

    同一レイヤーのペアにはボーナス（距離を0.8倍）を与える。
    """
    import numpy as np
    from scipy.optimize import linear_sum_assignment

    n_texts = len(stake_texts)
    n_stakes = len(stakes)
    if n_texts == 0 or n_stakes == 0:
        return

    # コスト行列を構築（閾値超えは大きなコスト）
    big_cost = threshold * 100
    cost_matrix = np.full((n_texts, n_stakes), big_cost)

    for ti, (text, tx, ty, tlayer) in enumerate(stake_texts):
        for si, s in enumerate(stakes):
            d = s.distance_to(tx, ty)
            if d < threshold:
                # 同一レイヤーにボーナス
                cost = d * 0.8 if s.layer == tlayer else d
                cost_matrix[ti, si] = cost

    # ハンガリアン法で最適マッチング
    row_ind, col_ind = linear_sum_assignment(cost_matrix)

    for ti, si in zip(row_ind, col_ind):
        if cost_matrix[ti, si] < big_cost:
            stakes[si].number = stake_texts[ti][0]


def _build_boundary_graph(stakes: list[Stake], lines: list[BoundaryLine],
                          threshold: float) -> nx.Graph:
    """境界線の接続グラフを構築

    杭番号をノード、線分をエッジとするグラフ。
    線分の端点から最近傍の杭を探し、マッチしたら辺を張る。
    """
    graph = nx.Graph()
    named_stakes = [s for s in stakes if s.number is not None]

    # 杭番号をノードとして追加
    for s in named_stakes:
        graph.add_node(s.number, x=s.x, y=s.y)

    # 各線分の端点を杭にマッチング
    for line in lines:
        s1 = _find_nearest_stake(named_stakes, line.x1, line.y1, threshold)
        s2 = _find_nearest_stake(named_stakes, line.x2, line.y2, threshold)
        if s1 and s2 and s1.number != s2.number:
            line.stake1_number = s1.number
            line.stake2_number = s2.number
            graph.add_edge(s1.number, s2.number)

    return graph


def _find_nearest_stake(stakes: list[Stake], x: float, y: float,
                        threshold: float) -> Optional[Stake]:
    """座標(x, y)に最も近い杭を返す。閾値以内になければNone。"""
    best_dist = threshold
    best = None
    for s in stakes:
        d = s.distance_to(x, y)
        if d < best_dist:
            best_dist = d
            best = s
    return best


# === ユーティリティ ===

def print_summary(parsed: ParsedDXF) -> None:
    """解析結果のサマリーを出力"""
    named = parsed.get_stakes_with_numbers()
    unnamed = [s for s in parsed.stakes if s.number is None]
    intersections = [s for s in named if s.is_intersection]

    print(f"=== DXF解析結果 ===")
    print(f"杭(CIRCLE r=0.25): {len(parsed.stakes)}個")
    print(f"  番号あり: {len(named)}個")
    print(f"  番号なし: {len(unnamed)}個")
    print(f"  交点杭(-付き): {len(intersections)}個")
    print(f"境界線(LWPOLYLINE 2点): {len(parsed.lines)}本")
    print(f"  杭マッチ済み: {sum(1 for l in parsed.lines if l.stake1_number and l.stake2_number)}本")
    print(f"地番: {len(parsed.parcels)}件")
    print(f"公共用地: {len(parsed.public_lands)}件")
    if parsed.graph:
        print(f"グラフ: ノード{parsed.graph.number_of_nodes()}, "
              f"エッジ{parsed.graph.number_of_edges()}")


if __name__ == '__main__':
    import sys
    filepath = sys.argv[1] if len(sys.argv) > 1 else \
        "data/input/三瓶 美紀枝 1070-6,1071-2　１班.DXF"
    parsed = parse_dxf(filepath)
    print_summary(parsed)
