"""
登記事項要約書 OCR比較テスト
PaddleOCR vs Claude 3.5 Sonnet vs GPT-4o (OpenRouter経由)
対象: 要約書１／２ の先頭10ページ
"""

import os
import sys
import csv
import time
import json
import base64
from pathlib import Path

import pypdfium2 as pdfium
import requests

# 設定
OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
PDF_PATH = Path("/home/maaatan/Chiseki/youyakusyo/要約書　１／２.pdf")
OUTPUT_DIR = Path("/home/maaatan/Chiseki/youyakusyo/output/comparison")
TEST_PAGES = 10
DPI = 300

VISION_PROMPT = """この画像は「登記事項要約書（土地）」です。画像内の全テキストを正確に読み取り、以下のCSV形式で出力してください。

## 出力形式
各土地エントリについて、以下のカラムでCSV形式で出力してください：
番号,所在,地番,地目,地積（㎡）,所有者・共有者住所,持分,所有者・共有者氏名,登記日,受付番号

## ルール
- ヘッダー行は不要です
- 1つの土地に複数の所有者がいる場合、所有者ごとに1行出力してください（番号〜地積は同じ値を繰り返す）
- 地目が変更されている場合（例：田→山林）、現在の地目を記載してください
- 持分が記載されていない場合（単独所有）は空欄にしてください
- 数字は半角で統一してください
- 丸囲み数字（①②③等）がある場合、変更前の情報なので現在の値を優先してください
- CSVデータのみ出力し、説明文は不要です"""


def pdf_to_images(pdf_path: Path, num_pages: int, dpi: int = DPI) -> list:
    """PDFを画像リストに変換"""
    pdf = pdfium.PdfDocument(str(pdf_path))
    images = []
    scale = dpi / 72
    for i in range(min(num_pages, len(pdf))):
        page = pdf[i]
        bitmap = page.render(scale=scale)
        img = bitmap.to_pil()
        images.append(img)
    return images


def image_to_base64(img) -> str:
    """PIL ImageをBase64エンコード"""
    import io
    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    return base64.b64encode(buffer.getvalue()).decode("utf-8")


def call_openrouter(model: str, base64_img: str, prompt: str) -> dict:
    """OpenRouter Vision API呼び出し"""
    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_img}",
                        },
                    },
                ],
            }
        ],
        "max_tokens": 4096,
        "temperature": 0,
    }

    start = time.time()
    resp = requests.post(OPENROUTER_URL, headers=headers, json=payload, timeout=120)
    elapsed = time.time() - start

    if resp.status_code != 200:
        return {"error": f"HTTP {resp.status_code}: {resp.text[:500]}", "elapsed": elapsed}

    data = resp.json()
    if "error" in data:
        return {"error": str(data["error"]), "elapsed": elapsed}

    content = data["choices"][0]["message"]["content"]
    usage = data.get("usage", {})

    return {
        "content": content,
        "elapsed": elapsed,
        "input_tokens": usage.get("prompt_tokens", 0),
        "output_tokens": usage.get("completion_tokens", 0),
    }


def run_paddleocr(img, page_num: int) -> dict:
    """PaddleOCR で処理"""
    from paddleocr import PaddleOCR

    if not hasattr(run_paddleocr, "_engine"):
        run_paddleocr._engine = PaddleOCR(use_angle_cls=True, lang="japan", show_log=False)

    tmp_path = OUTPUT_DIR / f"_tmp_paddle_{page_num}.png"
    img.save(str(tmp_path))

    start = time.time()
    result = run_paddleocr._engine.ocr(str(tmp_path), cls=True)
    elapsed = time.time() - start
    tmp_path.unlink(missing_ok=True)

    if not result or not result[0]:
        return {"content": "", "elapsed": elapsed}

    # Y座標でソートして行グループ化
    lines = []
    for det in result[0]:
        box = det[0]
        text = det[1][0]
        conf = det[1][1]
        y = min(p[1] for p in box)
        x = min(p[0] for p in box)
        lines.append((y, x, text, conf))
    lines.sort()

    rows = []
    cur = [lines[0]]
    for l in lines[1:]:
        if abs(l[0] - cur[0][0]) <= 15:
            cur.append(l)
        else:
            cur.sort(key=lambda a: a[1])
            rows.append(cur)
            cur = [l]
    if cur:
        cur.sort(key=lambda a: a[1])
        rows.append(cur)

    text_output = "\n".join(
        " | ".join(f"{t[2]}" for t in row) for row in rows
    )
    avg_conf = sum(l[3] for l in lines) / len(lines) if lines else 0

    return {
        "content": text_output,
        "elapsed": elapsed,
        "detected_lines": len(lines),
        "avg_confidence": round(avg_conf, 3),
    }


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print(f"PDF読み込み中: {PDF_PATH.name}")
    images = pdf_to_images(PDF_PATH, TEST_PAGES)
    print(f"  {len(images)}ページを画像変換完了\n")

    # モデル定義
    models = {
        "PaddleOCR": None,  # ローカル実行
        "Claude-3.5-Sonnet": "anthropic/claude-3.5-sonnet",
        "GPT-4o": "openai/gpt-4o",
    }

    results = {name: [] for name in models}

    for page_idx in range(len(images)):
        img = images[page_idx]
        page_num = page_idx + 1
        b64 = image_to_base64(img)
        print(f"--- ページ {page_num}/{len(images)} ---")

        # PaddleOCR
        print(f"  PaddleOCR...", end="", flush=True)
        paddle_result = run_paddleocr(img, page_num)
        print(f" {paddle_result['elapsed']:.1f}秒")
        results["PaddleOCR"].append(paddle_result)

        # 保存
        paddle_out = OUTPUT_DIR / f"page{page_num}_paddleocr.txt"
        with open(paddle_out, "w", encoding="utf-8") as f:
            f.write(paddle_result.get("content", ""))

        # Vision API (Claude, GPT-4o)
        for name, model_id in models.items():
            if model_id is None:
                continue

            print(f"  {name}...", end="", flush=True)
            api_result = call_openrouter(model_id, b64, VISION_PROMPT)

            if "error" in api_result:
                print(f" ERROR: {api_result['error'][:100]}")
            else:
                tokens = api_result.get('input_tokens', 0) + api_result.get('output_tokens', 0)
                print(f" {api_result['elapsed']:.1f}秒 ({tokens}トークン)")

            results[name].append(api_result)

            # 保存
            out_file = OUTPUT_DIR / f"page{page_num}_{name.lower().replace('.', '').replace('-', '_')}.txt"
            with open(out_file, "w", encoding="utf-8") as f:
                f.write(api_result.get("content", api_result.get("error", "")))

    # サマリーレポート
    print("\n" + "=" * 70)
    print("比較サマリー")
    print("=" * 70)

    summary_rows = []
    for name in models:
        page_results = results[name]
        total_time = sum(r.get("elapsed", 0) for r in page_results)
        errors = sum(1 for r in page_results if "error" in r)
        avg_time = total_time / len(page_results) if page_results else 0

        if name == "PaddleOCR":
            avg_lines = sum(r.get("detected_lines", 0) for r in page_results) / len(page_results)
            avg_conf = sum(r.get("avg_confidence", 0) for r in page_results) / len(page_results)
            detail = f"平均{avg_lines:.0f}行検出, 信頼度{avg_conf:.3f}"
            cost = "無料"
        else:
            total_in = sum(r.get("input_tokens", 0) for r in page_results)
            total_out = sum(r.get("output_tokens", 0) for r in page_results)
            detail = f"入力{total_in}トークン, 出力{total_out}トークン"
            # OpenRouter pricing estimate
            if "claude" in name.lower():
                cost_est = (total_in * 3 + total_out * 15) / 1_000_000
            else:
                cost_est = (total_in * 2.5 + total_out * 10) / 1_000_000
            cost = f"約${cost_est:.3f}"

        print(f"\n{name}:")
        print(f"  合計処理時間: {total_time:.1f}秒 (平均{avg_time:.1f}秒/ページ)")
        print(f"  エラー: {errors}/{len(page_results)}ページ")
        print(f"  詳細: {detail}")
        print(f"  コスト: {cost}")

        summary_rows.append({
            "model": name,
            "total_time_sec": round(total_time, 1),
            "avg_time_sec": round(avg_time, 1),
            "errors": errors,
            "detail": detail,
            "cost": cost,
        })

    # サマリーCSV
    summary_csv = OUTPUT_DIR / "comparison_summary.csv"
    with open(summary_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["model", "total_time_sec", "avg_time_sec", "errors", "detail", "cost"])
        writer.writeheader()
        writer.writerows(summary_rows)

    print(f"\n結果ファイル: {OUTPUT_DIR}/")
    print(f"サマリー: {summary_csv}")


if __name__ == "__main__":
    main()
