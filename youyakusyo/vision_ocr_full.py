"""
登記事項要約書（土地）全ページ CSV変換
Claude Sonnet 4 (OpenRouter) Vision API使用
"""

import csv
import os
import time
import base64
import io
from pathlib import Path

import pypdfium2 as pdfium
import requests

OPENROUTER_API_KEY = os.environ.get("OPENROUTER_API_KEY", "")
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "anthropic/claude-sonnet-4"
OUTPUT_DIR = Path("/home/maaatan/Chiseki/youyakusyo/output")
DPI = 300

PDF_FILES = [
    Path("/home/maaatan/Chiseki/youyakusyo/要約書　１／２.pdf"),
    Path("/home/maaatan/Chiseki/youyakusyo/要約書　２／２.pdf"),
]

VISION_PROMPT = """この画像は「登記事項要約書（土地）」です。画像内の全テキストを正確に読み取り、以下のCSV形式で出力してください。

## 出力形式
各土地エントリについて、以下のカラムでCSV形式で出力してください：
番号,所在,地番,地目,地積（㎡）,所有者・共有者住所,持分,所有者・共有者氏名,登記日,受付番号

## ページ構造の読み方（重要）
この帳票は一番左の列に「番号」欄があり、各土地エントリを区切っています。

- **番号欄に数字がある行** → 新しい土地エントリの開始です
- **番号欄が空欄の行** → 前ページから続く土地エントリの一部です（所有者情報の続き等）
- ページ先頭で番号欄が空欄の場合、その行は「前ページの最後の土地エントリの続き」です。この場合、番号欄に「続」と記入してください
- ページ末尾で表題部（所在・地番・地目・地積）だけあり、権利部（所有者情報）がない場合でも、その表題部は出力してください

## 表示履歴（地目・地積）の読み方（重要）
1つの土地に対して、地目や地積が縦に複数行並んでいることがあります。これは変更履歴です。

- **下線が引かれた値** → 変更済み（旧値）です。使用しないでください
- **下線のない最後の値** → 現在の値です。この値を使用してください
- 右側に②③などの丸数字や「分筆」「合筆」「錯誤」等の注記がある行は変更履歴です
- 表示履歴が次ページに続く場合もあります。次ページでも下線付きの数値が続く場合は旧値なので、下線のない値まで読み進めてください
- ページ先頭に番号欄が空白で、地積欄の位置に数値だけが縦に並んでいる場合（例: 41475、37388、5075、9857 のように）、それは前ページの表示履歴の続きです。これらの数値のうち下線のない最後の値が現在の地積です。番号「続更新」として、地積にその最終値を設定して1行出力してください。例: 続更新,,,雑種地,9857,,,,,
- この「続更新」行は権利部の「続」行より前に出力してください

## その他ルール
- ヘッダー行は不要です
- 1つの土地に複数の所有者がいる場合、所有者ごとに1行出力してください（番号〜地積は同じ値を繰り返す）
- 持分が記載されていない場合（単独所有）は空欄にしてください
- 数字は半角で統一してください
- CSVデータのみ出力し、説明文やコードブロック記号(```)は不要です"""

MAX_RETRIES = 3
RETRY_DELAY = 5


def call_vision_api(b64_img: str) -> dict:
    """Vision API呼び出し（リトライ付き）"""
    for attempt in range(MAX_RETRIES):
        try:
            start = time.time()
            resp = requests.post(OPENROUTER_URL, headers={
                "Authorization": f"Bearer {OPENROUTER_API_KEY}",
                "Content-Type": "application/json",
            }, json={
                "model": MODEL,
                "messages": [{"role": "user", "content": [
                    {"type": "text", "text": VISION_PROMPT},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64_img}"}},
                ]}],
                "max_tokens": 4096,
                "temperature": 0,
            }, timeout=120)
            elapsed = time.time() - start

            if resp.status_code == 429:
                wait = RETRY_DELAY * (attempt + 1)
                print(f" レート制限、{wait}秒待機...", end="", flush=True)
                time.sleep(wait)
                continue

            if resp.status_code != 200:
                return {"error": f"HTTP {resp.status_code}: {resp.text[:200]}", "elapsed": elapsed}

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
        except requests.exceptions.Timeout:
            if attempt < MAX_RETRIES - 1:
                print(f" タイムアウト、リトライ...", end="", flush=True)
                time.sleep(RETRY_DELAY)
            else:
                return {"error": "Timeout after retries", "elapsed": 0}

    return {"error": "Max retries exceeded", "elapsed": 0}


def clean_csv_content(content: str) -> str:
    """APIレスポンスからCSVデータのみ抽出"""
    lines = content.strip().split("\n")
    csv_lines = []
    for line in lines:
        line = line.strip()
        if line.startswith("```") or not line:
            continue
        csv_lines.append(line)
    return "\n".join(csv_lines)


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    total_tokens_in = 0
    total_tokens_out = 0
    total_time = 0
    errors = 0
    global_page = 0
    all_csv_lines = []

    csv_header = "番号,所在,地番,地目,地積（㎡）,所有者・共有者住所,持分,所有者・共有者氏名,登記日,受付番号"

    for pdf_path in PDF_FILES:
        if not pdf_path.exists():
            print(f"スキップ: {pdf_path.name}")
            continue

        pdf = pdfium.PdfDocument(str(pdf_path))
        num_pages = len(pdf)
        print(f"\n処理中: {pdf_path.name} ({num_pages}ページ)")
        scale = DPI / 72

        for i in range(num_pages):
            global_page += 1
            page = pdf[i]
            bitmap = page.render(scale=scale)
            img = bitmap.to_pil()

            buf = io.BytesIO()
            img.save(buf, format="PNG")
            b64 = base64.b64encode(buf.getvalue()).decode()

            print(f"  [{global_page:3d}/111] ", end="", flush=True)

            result = call_vision_api(b64)

            if "error" in result:
                print(f"ERROR: {result['error'][:80]}")
                errors += 1
                continue

            elapsed = result["elapsed"]
            t_in = result["input_tokens"]
            t_out = result["output_tokens"]
            total_tokens_in += t_in
            total_tokens_out += t_out
            total_time += elapsed

            csv_text = clean_csv_content(result["content"])
            csv_lines = [l for l in csv_text.split("\n") if l.strip()]

            print(f"{elapsed:5.1f}秒 | {len(csv_lines):2d}行 | {t_in+t_out}tok")

            # ページ番号付きで保存
            for line in csv_lines:
                all_csv_lines.append(f"{pdf_path.name},{global_page},{line}")

            # 個別ページ結果も保存
            page_file = OUTPUT_DIR / "pages" / f"page_{global_page:03d}.csv"
            page_file.parent.mkdir(exist_ok=True)
            with open(page_file, "w", encoding="utf-8-sig") as f:
                f.write(csv_header + "\n")
                f.write(csv_text + "\n")

    # 「続」行を前ページの物件に結合する後処理
    merged_lines = []
    last_entry_info = {}  # 最後の番号付きエントリの情報を保持

    for raw_line in all_csv_lines:
        # PDFファイル,ページ,番号,所在,地番,...
        parts = raw_line.split(",", 2)  # PDFファイル,ページ,残り
        if len(parts) < 3:
            merged_lines.append(raw_line)
            continue

        pdf_file, page_num, csv_part = parts
        csv_fields = csv_part.split(",")

        if len(csv_fields) >= 1 and csv_fields[0].strip() == "続更新":
            # 表示履歴の続き → 前ページの物件の地目・地積を更新
            if last_entry_info:
                new_chimoku = csv_fields[3].strip() if len(csv_fields) > 3 and csv_fields[3].strip() else last_entry_info["chimoku"]
                new_chiseki = csv_fields[4].strip() if len(csv_fields) > 4 and csv_fields[4].strip() else last_entry_info["chiseki"]
                last_entry_info["chimoku"] = new_chimoku
                last_entry_info["chiseki"] = new_chiseki
                # 既に出力済みの前ページの行も更新（最後の番号付き行を差し替え）
                for i in range(len(merged_lines) - 1, -1, -1):
                    line_parts = merged_lines[i].split(",", 2)
                    if len(line_parts) >= 3:
                        line_fields = line_parts[2].split(",")
                        if len(line_fields) >= 5 and line_fields[0].strip() == last_entry_info["num"]:
                            line_fields[3] = new_chimoku
                            line_fields[4] = new_chiseki
                            merged_lines[i] = f"{line_parts[0]},{line_parts[1]},{','.join(line_fields)}"
                            break
            # 続更新行自体はCSVに出力しない
            continue

        elif len(csv_fields) >= 1 and csv_fields[0].strip() == "続":
            # 前ページの物件の続き → 番号・所在・地番・地目・地積を引き継ぐ
            if last_entry_info:
                # 「続」行は所有者情報のみ持つ：住所,持分,氏名,登記日,受付番号
                # csv_fields: [続, 住所, (地番位置), (地目位置), (地積位置), ...]
                # Vision APIの出力で「続」行のカラム配置を正しく処理
                owner_addr = csv_fields[1].strip() if len(csv_fields) > 1 else ""
                mochibun = csv_fields[2].strip() if len(csv_fields) > 2 else ""
                # 持分が空で氏名がある場合、カラムがずれている可能性
                owner_name = csv_fields[3].strip() if len(csv_fields) > 3 else ""
                reg_date = csv_fields[4].strip() if len(csv_fields) > 4 else ""
                ref_num = csv_fields[5].strip() if len(csv_fields) > 5 else ""

                new_csv = f"{last_entry_info['num']},{last_entry_info['shozai']},{last_entry_info['chiban']},{last_entry_info['chimoku']},{last_entry_info['chiseki']},{owner_addr},{mochibun},{owner_name},{reg_date},{ref_num}"
                merged_lines.append(f"{pdf_file},{page_num},{new_csv}")
            else:
                merged_lines.append(raw_line)
        else:
            # 通常の番号付き行 → 情報を保持
            if len(csv_fields) >= 5:
                last_entry_info = {
                    "num": csv_fields[0].strip(),
                    "shozai": csv_fields[1].strip(),
                    "chiban": csv_fields[2].strip(),
                    "chimoku": csv_fields[3].strip(),
                    "chiseki": csv_fields[4].strip(),
                }
            merged_lines.append(raw_line)

    all_csv_lines = merged_lines

    # 統合CSV出力
    final_csv = OUTPUT_DIR / "要約書_全データ.csv"
    with open(final_csv, "w", encoding="utf-8-sig") as f:
        f.write(f"PDFファイル,ページ,{csv_header}\n")
        for line in all_csv_lines:
            f.write(line + "\n")

    # サマリー
    cost = (total_tokens_in * 3 + total_tokens_out * 15) / 1_000_000
    print(f"\n{'='*60}")
    print(f"処理完了")
    print(f"{'='*60}")
    print(f"  処理ページ: {global_page}")
    print(f"  CSV行数: {len(all_csv_lines)}")
    print(f"  エラー: {errors}")
    print(f"  合計時間: {total_time:.0f}秒 ({total_time/60:.1f}分)")
    print(f"  トークン: 入力{total_tokens_in:,} / 出力{total_tokens_out:,}")
    print(f"  コスト: ${cost:.3f}")
    print(f"\n出力ファイル: {final_csv}")


if __name__ == "__main__":
    main()
