"""Excel出力モジュール

結線指示票（.xls）と交点計算指示書（.xlsx）のExcel出力を行う。
テンプレートファイルをコピーし、計算結果データを書き込む。
"""

from __future__ import annotations

import os
from dataclasses import dataclass

import xlrd
import xlwt
from xlutils.copy import copy as xlutils_copy
from xlutils.filter import process, XLRDReader, XLWTWriter

import openpyxl


# ============================================================
# データクラス定義
# ============================================================

@dataclass
class KessenHeaderInfo:
    """結線指示票のヘッダー共通情報"""
    year: str | None = None           # 年度（例: "令和５年度"）
    district: str | None = None       # 地区名（例: "小豆畑・上小津田地区"）
    block_number: float | str | None = None  # ブロック番号
    oaza_name: str | None = None      # 大字名（例: "華川町小豆畑"）
    aza_name: str | None = None       # 字名（例: "内城"）


@dataclass
class KessenResult:
    """結線指示票の1筆分データ"""
    parcel_number: str          # 地番
    stake_sequence: list[str]   # 杭番号の順序リスト（閉合）
    land_use: str | None = None # 地目
    owner: str | None = None    # 所有者名


@dataclass
class KoutenHeaderInfo:
    """交点計算指示書のヘッダー共通情報"""
    district: str | None = None       # 地区名（例: "小豆畑・上小津田"）
    recorder: str | None = None       # 記入者名（例: "前田　高行"）
    block_number: str | None = None   # ブロック番号
    oaza_name: str | None = None      # 大字名（例: "華川町小豆畑"）


@dataclass
class IntersectionResult:
    """交点計算指示書の1交点分データ"""
    intersection_stake: str     # 交点杭番号
    baseline_point1: str        # 基準線の点1
    baseline_point2: str        # 基準線の点2
    extension_point1: str       # 延長線の点1
    extension_point2: str       # 延長線の点2


# ============================================================
# 結線指示票 (.xls) 出力
# ============================================================

def _get_style_list(rb: xlrd.Book) -> list[xlwt.XFStyle]:
    """xlrdブックからxlwt用のスタイルリストを取得する。

    xlutils.filterのprocessを使い、xlrdのXF情報を
    xlwtのXFStyleオブジェクトに変換する。
    """
    reader = XLRDReader(rb, "unknown.xls")
    writer = XLWTWriter()
    process(reader, writer)
    return writer.style_list


def _copy_template_sheet(
    rb: xlrd.Book,
    wb: xlwt.Workbook,
    style_list: list[xlwt.XFStyle],
    sheet_name: str,
) -> xlwt.Worksheet:
    """テンプレートの最初のシートを書式付きで新シートにコピーする。

    Args:
        rb: テンプレートのxlrdブック
        wb: 出力先のxlwtワークブック
        style_list: XFStyleリスト
        sheet_name: 新シートの名前

    Returns:
        書式コピー済みの新しいxlwtワークシート
    """
    rs = rb.sheet_by_index(0)
    ws = wb.add_sheet(sheet_name, cell_overwrite_ok=True)

    # 全セルの値とスタイルをコピー
    for row_idx in range(rs.nrows):
        for col_idx in range(rs.ncols):
            xf_idx = rs.cell_xf_index(row_idx, col_idx)
            val = rs.cell_value(row_idx, col_idx)
            if xf_idx < len(style_list):
                ws.write(row_idx, col_idx, val, style_list[xf_idx])
            else:
                ws.write(row_idx, col_idx, val)

    # セル結合をコピー
    for rlo, rhi, clo, chi in rs.merged_cells:
        ws.merge(rlo, rhi - 1, clo, chi - 1)

    return ws


def _write_kessen_sheet(
    ws: xlwt.Worksheet,
    result: KessenResult,
    style_list: list[xlwt.XFStyle],
    rb: xlrd.Book,
    header: KessenHeaderInfo | None = None,
    page_offset: int = 0,
) -> None:
    """結線指示票の1シートにデータを書き込む。

    杭番号データは行5~29（0-indexed）の4列配置:
      Col A(0): 番号1-25, Col B(1): 杭番号
      Col D(3): 番号26-50, Col E(4): 杭番号
      Col G(6): 番号51-75, Col H(7): 杭番号
      Col J(9): 番号76-100, Col K(10): 杭番号

    page_offset: 杭番号リストの開始位置（100点超の2ページ目は100）
    """
    rs = rb.sheet_by_index(0)

    def _get_style(row: int, col: int) -> xlwt.XFStyle:
        xf_idx = rs.cell_xf_index(row, col)
        if xf_idx < len(style_list):
            return style_list[xf_idx]
        return xlwt.XFStyle()

    # ヘッダー情報の書き込み
    if header is not None:
        if header.year is not None:
            ws.write(1, 0, header.year, _get_style(1, 0))
        if header.district is not None:
            ws.write(1, 3, header.district, _get_style(1, 3))
        if header.block_number is not None:
            ws.write(3, 0, header.block_number, _get_style(3, 0))
        if header.oaza_name is not None:
            ws.write(3, 2, header.oaza_name, _get_style(3, 2))
        if header.aza_name is not None:
            ws.write(3, 4, header.aza_name, _get_style(3, 4))

    # 地番をRow 3, Col 6に書き込み
    ws.write(3, 6, result.parcel_number, _get_style(3, 6))

    # 杭番号データの書き込み
    # 4列配置: 各列25行（行5~29、0-indexed）
    column_configs = [
        (1,),    # 列B: 杭番号 1-25
        (4,),    # 列E: 杭番号 26-50
        (7,),    # 列H: 杭番号 51-75
        (10,),   # 列K: 杭番号 76-100
    ]

    stakes = result.stake_sequence[page_offset:page_offset + 100]
    for i, stake in enumerate(stakes):
        col_group = i // 25           # 0, 1, 2, 3
        row_offset = i % 25           # 0-24
        row = 5 + row_offset          # 行5~29（0-indexed）
        stake_col = column_configs[col_group][0]

        # 杭番号セルのスタイルを取得し、文字列書式で書き込み
        # （"1.830"等の末尾ゼロが消えないよう、書式を"@"=テキストに設定）
        xf_idx = rs.cell_xf_index(row, stake_col)
        style = style_list[xf_idx] if xf_idx < len(style_list) else xlwt.XFStyle()
        text_style = xlwt.XFStyle()
        text_style.font = style.font
        text_style.alignment = style.alignment
        text_style.borders = style.borders
        text_style.pattern = style.pattern
        text_style.protection = style.protection
        text_style.num_format_str = '@'
        ws.write(row, stake_col, stake, text_style)


def write_kessen_excel(
    results: list[KessenResult],
    template_path: str,
    output_path: str,
    header: KessenHeaderInfo | None = None,
) -> None:
    """結線指示票をExcelに出力する。

    テンプレート(.xls)をコピーし、各筆ごとにシートを作成して
    杭番号データを書き込む。

    Args:
        results: 結線指示票データのリスト
        template_path: テンプレートファイルパス
        output_path: 出力ファイルパス
        header: ヘッダー共通情報（年度、地区名等）
    """
    if not results:
        raise ValueError("結果データが空です")

    # 出力ディレクトリを作成
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    # テンプレートを読み込み
    rb = xlrd.open_workbook(template_path, formatting_info=True)
    style_list = _get_style_list(rb)

    # xlutils.copyでワークブックをコピー（テンプレートシートの書式保持）
    wb = xlutils_copy(rb)

    def _unique_sheet_name(name: str) -> str:
        """シート名の重複を連番で回避"""
        if name not in used_names:
            used_names[name] = 1
            return name
        used_names[name] += 1
        return f"{name}({used_names[name]})"

    def _write_result_sheets(result: KessenResult, is_first: bool) -> None:
        """1筆分のシートを作成（100点超なら複数ページ）"""
        total = len(result.stake_sequence)
        pages = (total + 99) // 100  # 切り上げ

        for page in range(pages):
            offset = page * 100
            if page == 0 and is_first:
                ws = wb.get_sheet(0)
                ws.name = _unique_sheet_name(result.parcel_number)
            else:
                suffix = f"_{page + 1}" if pages > 1 and page > 0 else ""
                sheet_name = _unique_sheet_name(
                    f"{result.parcel_number}{suffix}"
                )
                ws = _copy_template_sheet(rb, wb, style_list, sheet_name)
            _write_kessen_sheet(
                ws, result, style_list, rb, header, page_offset=offset
            )

    used_names: dict[str, int] = {}
    _write_result_sheets(results[0], is_first=True)
    for result in results[1:]:
        _write_result_sheets(result, is_first=False)

    # 保存
    wb.save(output_path)


# ============================================================
# 交点計算指示書 (.xlsx) 出力
# ============================================================

def _to_number(value: str) -> float | str:
    """文字列を可能であれば数値に変換する。

    Excelセルに数値として書き込むため、数値に変換可能な
    文字列はfloatに変換する。変換不可の場合は文字列のまま返す。
    """
    try:
        return float(value)
    except (ValueError, TypeError):
        return value


# 交点ブロック配置の定数
# 横3ブロック x 縦5段 = 最大15交点
_KOUTEN_ROWS_PER_BLOCK = 11     # 段の行間隔
_KOUTEN_COLS_PER_BLOCK = 18     # ブロックの列間隔
_KOUTEN_MAX_COLS = 3            # 横ブロック数
_KOUTEN_MAX_ROWS = 5            # 縦段数
_KOUTEN_MAX_POINTS = _KOUTEN_MAX_COLS * _KOUTEN_MAX_ROWS  # 15

# 各段の中央◎行（1-indexed for openpyxl）
_KOUTEN_CENTER_ROWS = [10, 21, 32, 43, 54]


def _get_kouten_cell_positions(
    block_idx: int,
) -> dict[str, tuple[int, int]]:
    """交点ブロックのデータ書き込み位置を取得する。

    Args:
        block_idx: ブロックインデックス（0~14）

    Returns:
        フィールド名 → (row, col) の辞書（1-indexed）
    """
    row_group = block_idx // _KOUTEN_MAX_COLS  # 段 (0-4)
    col_group = block_idx % _KOUTEN_MAX_COLS   # 列 (0-2)

    center_row = _KOUTEN_CENTER_ROWS[row_group]
    col_offset = col_group * _KOUTEN_COLS_PER_BLOCK

    return {
        # 延長線（方向線）の点1: 上○（◎行-5, 列offset+11）
        "extension_point1": (center_row - 5, col_offset + 11),
        # 延長線（方向線）の点2: 上から2番目の○（◎行-2, 列offset+11）
        "extension_point2": (center_row - 2, col_offset + 11),
        # 基準線の左: 左○（◎行+2, 列offset+2）
        "baseline_point1": (center_row + 2, col_offset + 2),
        # 交点杭番号: ◎（◎行+2, 列offset+8）
        "intersection_stake": (center_row + 2, col_offset + 8),
        # 基準線の右: 右○（◎行+2, 列offset+14）
        "baseline_point2": (center_row + 2, col_offset + 14),
    }


def write_kouten_excel(
    results: list[IntersectionResult],
    template_path: str,
    output_path: str,
    header: KoutenHeaderInfo | None = None,
) -> None:
    """交点計算指示書をExcelに出力する。

    テンプレート(.xlsx)をコピーし、各交点のデータを
    横3ブロック x 縦5段の配置で書き込む。

    Args:
        results: 交点計算結果のリスト（最大15件）
        template_path: テンプレートファイルパス
        output_path: 出力ファイルパス
        header: ヘッダー共通情報（地区名、記入者等）

    Raises:
        ValueError: 交点数が最大数(15)を超えた場合
    """
    if not results:
        raise ValueError("結果データが空です")

    if len(results) > _KOUTEN_MAX_POINTS:
        raise ValueError(
            f"交点数が最大数({_KOUTEN_MAX_POINTS})を超えています: {len(results)}"
        )

    # 出力ディレクトリを作成
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

    # テンプレートをコピーして開く
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active

    # ヘッダー情報の書き込み（1-indexed）
    if header is not None:
        if header.district is not None:
            ws.cell(row=2, column=14, value=header.district)
        if header.recorder is not None:
            ws.cell(row=2, column=44, value=header.recorder)
        if header.block_number is not None:
            ws.cell(row=3, column=1, value=_to_number(header.block_number))
        if header.oaza_name is not None:
            ws.cell(row=3, column=22, value=header.oaza_name)

    # 各交点データを書き込み
    # 杭番号は文字列のまま書き込む（"1.830"等の末尾ゼロ保持のため）
    for idx, result in enumerate(results):
        positions = _get_kouten_cell_positions(idx)

        for field, value in [
            ("extension_point1", result.extension_point1),
            ("extension_point2", result.extension_point2),
            ("baseline_point1", result.baseline_point1),
            ("intersection_stake", result.intersection_stake),
            ("baseline_point2", result.baseline_point2),
        ]:
            cell = ws.cell(
                row=positions[field][0],
                column=positions[field][1],
                value=value,
            )
            cell.number_format = '@'

    # 保存
    wb.save(output_path)
