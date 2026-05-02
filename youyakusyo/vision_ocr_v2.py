"""
登記事項要約書（土地）全ページ CSV変換 v2
高画質PDF対応 + 照合元CSV構造準拠プロンプト
Claude Sonnet 4 (OpenRouter) Vision API使用
"""

import csv
import re
import time
import base64
import io
import os
from pathlib import Path

import pypdfium2 as pdfium
import requests

# === 設定 ===
OPENROUTER_API_KEY = os.environ.get(
    "OPENROUTER_API_KEY",
    "",
)
OPENROUTER_URL = "https://openrouter.ai/api/v1/chat/completions"
MODEL = "anthropic/claude-sonnet-4"
OUTPUT_DIR = Path("/home/maaatan/Chiseki/youyakusyo/output_v2")
DPI = 300

PDF_FILES = sorted(Path("/home/maaatan/Chiseki/youyakusyo/HQ PDF").glob("*.pdf"))

# 照合元CSVのカラム構造に準拠したプロンプト
VISION_PROMPT = """この画像は「登記事項要約書（土地）」です。表形式のデータを正確に読み取り、以下のCSV形式で出力してください。

## CSV出力カラム（9列固定・順序厳守）

番号,所在,地番,地目,地積,住所,持分,氏名,受付情報

各カラムの定義:
1. 番号: 要約書の左端の通し番号（半角数字）
2. 所在: 市区町村＋字名まで（例: 北茨城市関本町富士ケ丘字金畑）
3. 地番: 地番（半角、例: 898番1）
4. 地目: 現在の地目（下線なしの値のみ。下線付き＝旧値は無視）
5. 地積: 現在の地積㎡（半角数字。下線なしの値のみ。小数点は「.」で出力）
6. 住所: 所有者の住所
7. 持分: 持分（例: 持分28分の1）。単独所有で持分なしの場合は「空欄」にする。★絶対にこの列を飛ばさないこと★
8. 氏名: 所有者の氏名（姓と名の間にスペースは入れない）
9. 受付情報: 「○○年○月○日受付 第○○号」の形式

## ★最重要ルール1: 9列固定★

どの行も必ず9列（カンマ8個）で出力してください。

## ★最重要ルール2: 空欄は `(空)` と明示出力★

値がないフィールドは、空文字列ではなく必ず `(空)` と明示的に出力してください。
これにより、列のずれやドロップを防ぎます。

正しい例（単独所有、持分なし）:
1,北茨城市関本町富士ケ丘字金畑,882番8,原野,99,北茨城市関本町関本中2669番地の1,(空),株式会社高砂鐵工所,平成20年2月7日受付 第1500号

正しい例（共有）:
5,北茨城市関本町富士ケ丘字金畑,898番1,原野,555,北茨城市関本町富士ケ丘369番地,持分28分の1,安島嘉弘,平成10年6月19日受付 第3218号

正しい例（所有者型、氏名のみ）:
17,北茨城市関本町富士ケ丘字深沢,1083番,山林,1031,(空),(空),馬上初吉,(空)

正しい例（表題部のみ、所有者情報は次ページ）:
7,北茨城市華川町上小津田字広町,(空),(空),(空),(空),(空),(空),(空)

間違いの例1（列が抜けている - 空文字列のまま）:
1,北茨城市関本町富士ケ丘字金畑,882番8,原野,99,北茨城市関本町関本中2669番地の1,,株式会社高砂鐵工所,平成20年2月7日受付 第1500号
（↑空の持分列が曖昧）

間違いの例2（列数不足）:
1,北茨城市関本町富士ケ丘字金畑,882番8,原野,99,北茨城市関本町関本中2669番地の1,株式会社高砂鐵工所,平成20年2月7日受付 第1500号
（↑8列しかなくNG）

## 地積の小数点について（重要）

地積欄では、整数部と小数部が **縦の点線（破線）** で区切られていることがあります。
この縦の点線は **小数点** を意味します。必ず「.」（ピリオド）として読み取ってください。

例:
- 地積欄に「514」の右に点線があり、その右に「24」→ 地積は「514.24」
- 地積欄に「85」の右に点線があり、その右に「00」→ 地積は「85.00」
- 地積欄に「110」の右に点線があり、その右に「90」→ 地積は「110.90」

宅地・雑種地では小数点以下2桁まで記載されるのが一般的です。
整数部と小数部を連結してしまわないよう注意してください（例: ×「11090」→ ○「110.90」）。

## ページ構造の読み方

- **番号欄に数字がある行** → 新しい土地エントリの開始
- **番号欄が空欄の行（ページ先頭）** → 前ページから続く所有者情報
  - この場合、番号欄に「続」と記入し、所在・地番・地目・地積は `(空)` にする
  - 例: 続,(空),(空),(空),(空),多賀郡関本村山小屋12番地,持分28分の1,金沢幸太郎,明治36年5月27日受付 第1323号
- 1つの土地に複数の所有者がいる場合、所有者ごとに1行出力（番号〜地積は同じ値を繰り返す）

## 所有者記載の2種類（★重要★）

要約書には2種類の所有者記載があります。区別してください。

### 種類A: 「所有者」（旧式、氏名のみ）
ラベルが「所有者」と書かれた行。住所・登記日・受付番号がなく、**氏名のみ**が記載されています。

例（PDF上の表記）:
```
所有者  馬上初吉
```
または
```
所有者  官有地
```

この場合の**出力フォーマット**:
```
17,北茨城市関本町富士ケ丘字深沢,1083番,山林,1031,,,馬上初吉,
```
→ 住所=空、持分=空、氏名=氏名のみ、受付=空

★絶対に氏名を住所列に入れないでください★
★「所有者」記載でも氏名は必ず8列目（氏名列）に配置してください★

### 種類B: 「権利部 所有権」（新式、住所＋氏名＋日付＋番号）
ラベルが「権利部 所有権」と書かれた行。住所・氏名・登記日・受付番号が揃っています。

例（PDF上の表記）:
```
権利部所有権  茨城県北茨城市関本町富士ケ丘488番地  金澤忠一   令和6年11月8日 第13588号
```

この場合の**出力フォーマット**:
```
18,北茨城市関本町富士ケ丘字深沢,1084番1,山林,1041,茨城県北茨城市関本町富士ケ丘488番地,,金澤忠一,令和6年11月8日受付 第13588号
```
→ 住所・氏名・受付情報を正しく配置

### 区別のポイント
- ラベル「所有者」→ 種類A（氏名のみ）
- ラベル「権利部所有権」「権利部 所有権」「所有権」→ 種類B（住所あり）
- どちらの場合も、**氏名は必ず8列目の氏名列**に配置
- 住所がない場合は住所列を空欄にする（氏名を住所列に書かない）

## 重要: 同じ所有者情報の重複出力禁止
1つの物件の1人の所有者情報を、複数回出力しないでください。1物件1所有者につき1行のみです（共有の場合は人数分出力）。

## 表示履歴（地目・地積）の読み方

地目や地積が縦に複数行並んでいる場合は変更履歴です。
- **下線が引かれた値** → 旧値。無視してください
- **下線のない最後の値** → 現在の値。この値を使用してください
- ②③等の丸数字や「分筆」「合筆」の注記がある行は変更履歴の説明です

## 「続更新」の使用条件（厳格ルール）

「続更新」は、**ページの最先頭**で、表題部ヘッダー（番号欄・所在行）がなく、地積の数値だけが縦に並んでいる場合にのみ使用してください。

- ページ途中で別の物件の表示履歴（下線付き旧値→現在値）が見えても、それは「続更新」ではなく、その物件固有の表示履歴です
- 「続更新」は前ページから表示履歴が途切れた場合のみ発生します

例（正しい「続更新」）: ページ先頭に番号なしで数値だけ→ 続更新,(空),(空),雑種地,9857,(空),(空),(空),(空)
例（誤り）: ページ途中で次の物件の下に旧値・現在値がある→ これは「続更新」ではない

## 番号列のルール（★最重要★）

**番号列に数字がある行は、必ずその行を出力してください。**
所在しか情報がない場合でも（ページ末尾で表題部ヘッダーだけが先に書かれ、詳細情報が次ページに続く場合）、番号付きの行として必ず出力してください。

例: ページ末尾に「7 表題部 北茨城市華川町上小津田字広町」とあり、地番・地目・地積が無い場合:
  7,北茨城市華川町上小津田字広町,(空),(空),(空),(空),(空),(空),(空)
  （地番・地目・地積等は `(空)` で次ページで補完する）

## 「続表題」の使用条件

前ページの最後の番号付き物件が**所在だけ**で終わり、その物件の地番・地目・地積・所有者情報が次ページの先頭にある場合、次ページの先頭行は「続表題」として出力してください。

「続表題」行は、所在以外の全情報（地番・地目・地積・住所・持分・氏名・受付情報）を含みます。

例（次ページの冒頭）:
  続表題,,1319番1,宅地,496.83,川崎市多摩区宿河原六丁目9番9号,,関山美津惠,平成28年12月13日受付 第15825号

「続表題」と「続更新」「続」は別の意味です:
- 続更新 → 表示履歴の地目・地積だけの更新（次ページ先頭に数値のみ縦に並ぶケース）
- 続 → 所有者情報だけの継続（次ページ先頭に所有者だけ縦に並ぶケース）
- 続表題 → 前ページの表題部ヘッダー（番号＋所在だけ）の続きで、地番・地目・地積・所有者すべてを含む

## その他ルール
- ヘッダー行は不要
- 数字は半角で統一
- 氏名にスペースは入れない（例: ×「金沢 幸太郎」→ ○「金沢幸太郎」）
- CSVデータのみ出力（説明文やコードブロック記号は不要）
"""

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
                "max_tokens": 8192,
                "temperature": 0,
            }, timeout=180)
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


def normalize_empty_marker(value: str) -> str:
    """空欄マーカー `(空)` を空文字列に変換"""
    v = value.strip()
    if v in ("(空)", "（空）", "(空欄)", "（空欄）", "-", "―"):
        return ""
    return v


def validate_csv_columns(csv_text: str, page_num: int) -> str:
    """CSV行が9列であることを検証し、(空)マーカーを正規化"""
    lines = csv_text.split("\n")
    validated = []
    warnings = 0
    for line in lines:
        if not line.strip():
            continue
        cols = line.split(",")
        if len(cols) < 9:
            cols.extend([""] * (9 - len(cols)))
            warnings += 1
        elif len(cols) > 9:
            cols = cols[:8] + [",".join(cols[8:])]
            warnings += 1
        # 空欄マーカーを空文字列に変換
        cols = [normalize_empty_marker(c) for c in cols]
        validated.append(",".join(cols))
    if warnings > 0:
        print(f" [列数補正: {warnings}行]", end="", flush=True)
    return "\n".join(validated)


def _is_duplicate_zoku(current_fields: list, all_lines: list, current_idx: int) -> bool:
    """対策: 続行の重複検知
    この続行の所有者情報(住所+氏名+受付)が、
    後続の別の番号付きエントリ直後の続行と完全一致するなら、
    現在の続行は誤配置の重複として破棄する。
    """
    cur_addr = current_fields[5].strip() if len(current_fields) > 5 else ""
    cur_name = current_fields[7].strip() if len(current_fields) > 7 else ""
    cur_uketsuke = current_fields[8].strip() if len(current_fields) > 8 else ""

    if not cur_name and not cur_addr:
        return False

    # 後続を走査
    seen_numbered = False
    for idx in range(current_idx + 1, min(current_idx + 10, len(all_lines))):
        parts = all_lines[idx].split(",", 2)
        if len(parts) < 3:
            continue
        fields = parts[2].split(",")
        if not fields:
            continue
        num = fields[0].strip()

        if num.isdigit():
            seen_numbered = True
            continue

        if num == "続" and seen_numbered:
            # 後続の番号エントリの後の続行
            later_addr = fields[5].strip() if len(fields) > 5 else ""
            later_name = fields[7].strip() if len(fields) > 7 else ""
            later_uketsuke = fields[8].strip() if len(fields) > 8 else ""
            if (cur_addr == later_addr and
                cur_name == later_name and
                cur_uketsuke == later_uketsuke):
                return True

    return False


def merge_continuation_lines(all_csv_lines: list) -> list:
    """「続」行を前ページの物件に結合し、「続更新」で地目・地積を更新"""
    merged = []
    last_entry = None  # {num, shozai, chiban, chimoku, chiseki}
    current_page = None
    page_line_index = 0  # 各ページ内での行番号（0始まり）

    # 先読み用にインデックスアクセス
    for line_idx, raw_line in enumerate(all_csv_lines):
        # PDFファイル,ページ,番号,所在,地番,地目,地積,住所,持分,氏名,受付情報
        parts = raw_line.split(",", 2)
        if len(parts) < 3:
            merged.append(raw_line)
            continue

        pdf_file, page_num, csv_part = parts
        fields = csv_part.split(",")

        if len(fields) < 1:
            merged.append(raw_line)
            continue

        # ページ変更の追跡
        if page_num != current_page:
            current_page = page_num
            page_line_index = 0
        else:
            page_line_index += 1

        entry_num = fields[0].strip()

        if entry_num == "続更新":
            # 対策2: 続更新はページ先頭（1行目以内）のみ許可
            if page_line_index > 0:
                # ページ途中の続更新 → 偽検出として破棄
                continue

            # 対策2: 直後のエントリと地目・地積が同一なら偽検出として破棄
            tsuzuki_chimoku = fields[3].strip() if len(fields) > 3 else ""
            tsuzuki_chiseki = fields[4].strip() if len(fields) > 4 else ""
            is_false_detection = False
            for next_idx in range(line_idx + 1, min(line_idx + 3, len(all_csv_lines))):
                next_parts = all_csv_lines[next_idx].split(",", 2)
                if len(next_parts) >= 3:
                    next_fields = next_parts[2].split(",")
                    next_num = next_fields[0].strip() if next_fields else ""
                    if next_num not in ("続", "続更新", ""):
                        # 番号付きエントリを発見
                        next_chimoku = next_fields[3].strip() if len(next_fields) > 3 else ""
                        next_chiseki = next_fields[4].strip() if len(next_fields) > 4 else ""
                        if tsuzuki_chimoku == next_chimoku and tsuzuki_chiseki == next_chiseki:
                            is_false_detection = True
                        break

            if is_false_detection:
                continue

            # 正当な続更新 → 直前の物件の地目・地積を上書き
            if last_entry:
                if tsuzuki_chimoku:
                    last_entry["chimoku"] = tsuzuki_chimoku
                if tsuzuki_chiseki:
                    last_entry["chiseki"] = tsuzuki_chiseki
                # 既出力行の地目・地積も更新（地番で照合、直近1件のみ）
                target_chiban = last_entry["chiban"]
                for i in range(len(merged) - 1, -1, -1):
                    m_parts = merged[i].split(",", 2)
                    if len(m_parts) >= 3:
                        m_fields = m_parts[2].split(",")
                        if len(m_fields) >= 5 and m_fields[2].strip() == target_chiban:
                            if tsuzuki_chimoku:
                                m_fields[3] = tsuzuki_chimoku
                            if tsuzuki_chiseki:
                                m_fields[4] = tsuzuki_chiseki
                            merged[i] = f"{m_parts[0]},{m_parts[1]},{','.join(m_fields)}"
                            break
            continue

        elif entry_num == "続表題":
            # 前ページの表題部（番号＋所在だけ）の続き → 地番・地目・地積・所有者を補完
            if last_entry:
                chiban = fields[2].strip() if len(fields) > 2 else ""
                chimoku = fields[3].strip() if len(fields) > 3 else ""
                chiseki = fields[4].strip() if len(fields) > 4 else ""
                addr = fields[5].strip() if len(fields) > 5 else ""
                mochibun = fields[6].strip() if len(fields) > 6 else ""
                name = fields[7].strip() if len(fields) > 7 else ""
                uketsuke = fields[8].strip() if len(fields) > 8 else ""

                # last_entryに地番等がなければ補完
                if chiban and not last_entry.get("chiban"):
                    last_entry["chiban"] = chiban
                if chimoku and not last_entry.get("chimoku"):
                    last_entry["chimoku"] = chimoku
                if chiseki and not last_entry.get("chiseki"):
                    last_entry["chiseki"] = chiseki

                # 既出力の不完全な行（地番空）を修正
                for i in range(len(merged) - 1, -1, -1):
                    m_parts = merged[i].split(",", 2)
                    if len(m_parts) >= 3:
                        m_fields = m_parts[2].split(",")
                        if (len(m_fields) >= 5 and
                            m_fields[0].strip() == last_entry["num"] and
                            not m_fields[2].strip()):
                            m_fields[2] = last_entry["chiban"]
                            m_fields[3] = last_entry["chimoku"]
                            m_fields[4] = last_entry["chiseki"]
                            merged[i] = f"{m_parts[0]},{m_parts[1]},{','.join(m_fields)}"
                            break

                new_csv = (
                    f"{last_entry['num']},{last_entry['shozai']},"
                    f"{last_entry['chiban']},{last_entry['chimoku']},"
                    f"{last_entry['chiseki']},{addr},{mochibun},{name},{uketsuke}"
                )
                merged.append(f"{pdf_file},{page_num},{new_csv}")
            continue

        elif entry_num == "続":
            # 対策: 続行の重複検知
            # この続行と同じ所有者情報が、後続の別の番号付きエントリ直後の続行と一致するなら、
            # それは誤配置の重複 → 破棄
            if _is_duplicate_zoku(fields, all_csv_lines, line_idx):
                continue

            # 前ページの物件の続き（所有者情報のみ）
            if last_entry:
                addr = fields[5].strip() if len(fields) > 5 else ""
                mochibun = fields[6].strip() if len(fields) > 6 else ""
                name = fields[7].strip() if len(fields) > 7 else ""
                uketsuke = fields[8].strip() if len(fields) > 8 else ""

                new_csv = (
                    f"{last_entry['num']},{last_entry['shozai']},"
                    f"{last_entry['chiban']},{last_entry['chimoku']},"
                    f"{last_entry['chiseki']},{addr},{mochibun},{name},{uketsuke}"
                )
                merged.append(f"{pdf_file},{page_num},{new_csv}")
            else:
                merged.append(raw_line)
        else:
            # 通常の番号付き行
            if len(fields) >= 5:
                last_entry = {
                    "num": fields[0].strip(),
                    "shozai": fields[1].strip(),
                    "chiban": fields[2].strip(),
                    "chimoku": fields[3].strip(),
                    "chiseki": fields[4].strip(),
                }
            merged.append(raw_line)

    return merged


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    (OUTPUT_DIR / "pages").mkdir(exist_ok=True)

    if not PDF_FILES:
        print("エラー: HQ PDFファイルが見つかりません")
        print(f"  検索パス: /home/maaatan/Chiseki/youyakusyo/HQ PDF/")
        return

    print(f"PDFファイル: {len(PDF_FILES)}個")
    for p in PDF_FILES:
        print(f"  {p.name}")

    total_tokens_in = 0
    total_tokens_out = 0
    total_time = 0
    errors = 0
    global_page = 0
    all_csv_lines = []

    csv_header = "番号,所在,地番,地目,地積,住所,持分,氏名,受付情報"

    for pdf_path in PDF_FILES:
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
            csv_text = validate_csv_columns(csv_text, global_page)
            csv_lines = [l for l in csv_text.split("\n") if l.strip()]

            print(f" {elapsed:5.1f}秒 | {len(csv_lines):2d}行 | {t_in+t_out}tok")

            # ページ番号付きで保存
            for line in csv_lines:
                all_csv_lines.append(f"{pdf_path.name},{global_page},{line}")

            # 個別ページ結果
            page_file = OUTPUT_DIR / "pages" / f"page_{global_page:03d}.csv"
            with open(page_file, "w", encoding="utf-8-sig") as f:
                f.write(csv_header + "\n")
                f.write(csv_text + "\n")

    # 生データ保存（後処理前）
    raw_csv = OUTPUT_DIR / "Vision_API生データ.csv"
    with open(raw_csv, "w", encoding="utf-8-sig") as f:
        f.write(f"PDFファイル,ページ,{csv_header}\n")
        for line in all_csv_lines:
            f.write(line + "\n")
    print(f"\n生データ保存: {raw_csv}")

    # 後処理: 続行マージ + 続更新処理
    print("後処理: 続行マージ実行中...")
    merged_lines = merge_continuation_lines(all_csv_lines)

    # 統合CSV出力
    final_csv = OUTPUT_DIR / "要約書_全データ.csv"
    with open(final_csv, "w", encoding="utf-8-sig") as f:
        f.write(f"PDFファイル,ページ,{csv_header}\n")
        for line in merged_lines:
            f.write(line + "\n")

    # サマリー
    cost = (total_tokens_in * 3 + total_tokens_out * 15) / 1_000_000
    print(f"\n{'='*60}")
    print(f"処理完了")
    print(f"{'='*60}")
    print(f"  処理ページ: {global_page}")
    print(f"  生データ行数: {len(all_csv_lines)}")
    print(f"  マージ後行数: {len(merged_lines)}")
    print(f"  エラー: {errors}")
    print(f"  合計時間: {total_time:.0f}秒 ({total_time/60:.1f}分)")
    print(f"  トークン: 入力{total_tokens_in:,} / 出力{total_tokens_out:,}")
    print(f"  推定コスト: ${cost:.3f}")
    print(f"\n出力ファイル:")
    print(f"  生データ: {raw_csv}")
    print(f"  マージ後: {final_csv}")


if __name__ == "__main__":
    main()
