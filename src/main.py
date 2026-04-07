"""地籍調査DXツール - メインエントリポイント

DXFファイルから結線指示票と交点計算指示書を自動生成する。

使い方:
    python -m src.main <DXFファイル> [オプション]

例:
    python -m src.main "data/input/三瓶 美紀枝 1070-6,1071-2　１班.DXF"
    python -m src.main "data/input/三瓶 美紀枝 1070-6,1071-2　１班.DXF" --parcels 1070-6,1069-5
"""

import argparse
import os
import sys

from src.dxf_parser import parse_dxf, print_summary, get_block_numbers
from src.kessen_generator import generate_kessen
from src.kouten_generator import generate_kouten
from src.excel_writer import (
    write_kessen_excel,
    write_kouten_excel,
    KessenResult as ExcelKessenResult,
    IntersectionResult as ExcelIntersectionResult,
)


def main():
    parser = argparse.ArgumentParser(
        description="地籍調査DXツール - DXFから帳票を自動生成"
    )
    parser.add_argument("dxf_file", help="入力DXFファイルパス")
    parser.add_argument(
        "--parcels",
        help="対象地番（カンマ区切り、省略時は全地番）",
        default=None,
    )
    parser.add_argument(
        "--kessen-template",
        help="結線指示票テンプレートパス",
        default="data/templates/結線指示票 - ブランク.xls",
    )
    parser.add_argument(
        "--kouten-template",
        help="交点計算指示書テンプレートパス",
        default="data/templates/交点計算指示書 - ブランク.xlsx",
    )
    parser.add_argument(
        "--output-dir",
        help="出力ディレクトリ",
        default="output",
    )
    parser.add_argument(
        "--block",
        type=int,
        help="交点計算指示書のブロック番号（系統番号）フィルタ",
        default=None,
    )
    parser.add_argument(
        "--no-kessen",
        action="store_true",
        help="結線指示票を生成しない",
    )
    parser.add_argument(
        "--no-kouten",
        action="store_true",
        help="交点計算指示書を生成しない",
    )
    args = parser.parse_args()

    # DXF解析
    print(f"DXFファイルを解析中: {args.dxf_file}")
    parsed = parse_dxf(args.dxf_file)
    print_summary(parsed)

    target_parcels = None
    if args.parcels:
        target_parcels = [p.strip() for p in args.parcels.split(",")]

    os.makedirs(args.output_dir, exist_ok=True)

    # 結線指示票生成
    if not args.no_kessen:
        print("\n結線指示票を生成中...")
        kessen_results = generate_kessen(parsed, target_parcels)
        if kessen_results:
            # KessenResultの型変換（kessen_generator → excel_writer）
            excel_results = [
                ExcelKessenResult(
                    parcel_number=r.parcel_number,
                    stake_sequence=r.stake_sequence,
                    land_use=r.land_use,
                    owner=r.owner,
                )
                for r in kessen_results
            ]
            output_path = os.path.join(args.output_dir, "結線指示票.xls")
            write_kessen_excel(
                excel_results, args.kessen_template, output_path
            )
            print(f"  {len(kessen_results)}筆分の結線指示票を出力: {output_path}")
            for r in kessen_results:
                print(f"    {r.parcel_number}: {len(r.stake_sequence)}杭")
        else:
            print("  対象の地番が見つかりませんでした")

    # 交点計算指示書生成
    if not args.no_kouten:
        print("\n交点計算指示書を生成中...")
        # ブロック番号指定がなければ全交点を取得（ブロック別シート分割はexcel_writerが行う）
        kouten_results = generate_kouten(parsed, block_number=args.block)
        if kouten_results:
            excel_results = [
                ExcelIntersectionResult(
                    intersection_stake=r.intersection_stake,
                    baseline_point1=r.baseline_point1,
                    baseline_point2=r.baseline_point2,
                    extension_point1=r.extension_point1,
                    extension_point2=r.extension_point2,
                )
                for r in kouten_results
            ]
            output_path = os.path.join(args.output_dir, "交点計算指示書.xlsx")
            write_kouten_excel(
                excel_results, args.kouten_template, output_path
            )
            print(f"  {len(kouten_results)}交点分の指示書を出力: {output_path}")
            # ブロック別の内訳を表示
            blocks = get_block_numbers(parsed)
            for b in blocks:
                block_count = sum(1 for r in kouten_results
                                 if r.intersection_stake.lstrip('-交').startswith(f"{b}."))
                if block_count > 0:
                    print(f"    ブロック{b}: {block_count}交点")
            for r in kouten_results:
                print(f"    交点={r.intersection_stake} "
                      f"基準線=({r.baseline_point1}, {r.baseline_point2}) "
                      f"延長線=({r.extension_point1}, {r.extension_point2})")
        else:
            print("  交点杭が見つかりませんでした")

    print("\n完了")


if __name__ == "__main__":
    main()
