"""
登記事項要約書（土地）PDF → CSV変換スクリプト
PaddleOCR v2.9.1 を使用
"""

import os
import sys
import csv
import time
import json
from pathlib import Path

import pypdfium2 as pdfium
from paddleocr import PaddleOCR


# 設定
DPI = 300
OUTPUT_DIR = Path("/home/maaatan/Chiseki/youyakusyo/output")
PDF_FILES = [
    Path("/home/maaatan/Chiseki/youyakusyo/要約書　１／２.pdf"),
    Path("/home/maaatan/Chiseki/youyakusyo/要約書　２／２.pdf"),
]


def pdf_to_images(pdf_path: Path, dpi: int = DPI) -> list:
    """PDFを画像リストに変換"""
    pdf = pdfium.PdfDocument(str(pdf_path))
    images = []
    scale = dpi / 72
    for i in range(len(pdf)):
        page = pdf[i]
        bitmap = page.render(scale=scale)
        img = bitmap.to_pil()
        images.append(img)
    return images


def ocr_page(ocr_engine, img, page_num: int) -> list[dict]:
    """1ページをOCR処理し、Y座標でソートしたテキスト行リストを返す"""
    # 一時ファイルに保存してOCR実行
    tmp_path = OUTPUT_DIR / f"_tmp_page_{page_num}.png"
    img.save(str(tmp_path))

    result = ocr_engine.ocr(str(tmp_path), cls=True)
    tmp_path.unlink(missing_ok=True)

    if not result or not result[0]:
        return []

    lines = []
    for detection in result[0]:
        box = detection[0]
        text = detection[1][0]
        confidence = detection[1][1]

        # バウンディングボックスからY座標（上端）とX座標（左端）を取得
        y_top = min(p[1] for p in box)
        x_left = min(p[0] for p in box)
        y_bottom = max(p[1] for p in box)
        x_right = max(p[0] for p in box)

        lines.append({
            "text": text,
            "confidence": confidence,
            "x_left": x_left,
            "y_top": y_top,
            "x_right": x_right,
            "y_bottom": y_bottom,
        })

    # Y座標 → X座標 の順でソート
    lines.sort(key=lambda l: (l["y_top"], l["x_left"]))
    return lines


def group_lines_by_row(lines: list[dict], y_threshold: float = 15.0) -> list[list[dict]]:
    """Y座標が近いテキストを同一行としてグループ化"""
    if not lines:
        return []

    rows = []
    current_row = [lines[0]]
    current_y = lines[0]["y_top"]

    for line in lines[1:]:
        if abs(line["y_top"] - current_y) <= y_threshold:
            current_row.append(line)
        else:
            current_row.sort(key=lambda l: l["x_left"])
            rows.append(current_row)
            current_row = [line]
            current_y = line["y_top"]

    if current_row:
        current_row.sort(key=lambda l: l["x_left"])
        rows.append(current_row)

    return rows


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("PaddleOCR エンジン初期化中...")
    ocr_engine = PaddleOCR(use_angle_cls=True, lang="japan")

    all_page_data = []
    global_page_num = 0

    for pdf_path in PDF_FILES:
        if not pdf_path.exists():
            print(f"  スキップ: {pdf_path.name} (ファイルなし)")
            continue

        print(f"\n処理中: {pdf_path.name}")
        images = pdf_to_images(pdf_path)
        print(f"  {len(images)} ページ検出")

        for i, img in enumerate(images):
            global_page_num += 1
            start = time.time()

            lines = ocr_page(ocr_engine, img, global_page_num)
            rows = group_lines_by_row(lines)

            elapsed = time.time() - start
            print(f"  ページ {global_page_num:3d}/{111}: {len(lines):3d}テキスト検出 ({elapsed:.1f}秒)")

            # 行ごとにテキストを結合
            for row in rows:
                row_text = " | ".join(item["text"] for item in row)
                avg_conf = sum(item["confidence"] for item in row) / len(row)

                all_page_data.append({
                    "pdf_file": pdf_path.name,
                    "page": global_page_num,
                    "y_position": round(row[0]["y_top"]),
                    "row_text": row_text,
                    "confidence": round(avg_conf, 3),
                    "detail_json": json.dumps(
                        [{"text": item["text"], "x": round(item["x_left"]), "conf": round(item["confidence"], 3)} for item in row],
                        ensure_ascii=False,
                    ),
                })

    # CSV出力
    csv_path = OUTPUT_DIR / "要約書_OCR結果.csv"
    with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["pdf_file", "page", "y_position", "row_text", "confidence", "detail_json"])
        writer.writeheader()
        writer.writerows(all_page_data)

    print(f"\n完了: {csv_path}")
    print(f"  全{global_page_num}ページ → {len(all_page_data)}行のCSVデータ")


if __name__ == "__main__":
    main()
