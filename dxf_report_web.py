#!/usr/bin/env python3
"""
DXF帳票自動生成ツールのブラウザUI。

Python標準ライブラリでローカルHTTPサーバーを起動し、ブラウザからDXFを選択して
結線指示票・交点計算指示書をZIPでダウンロードできるようにする。
"""

from __future__ import annotations

import base64
import io
import json
import tempfile
import traceback
import webbrowser
import zipfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

try:
    from src.dxf_parser import get_block_numbers, parse_dxf
    from src.excel_writer import (
        IntersectionResult as ExcelIntersectionResult,
        KessenResult as ExcelKessenResult,
        write_kessen_excel,
        write_kouten_excel,
    )
    from src.kessen_generator import generate_kessen
    from src.kouten_generator import generate_kouten
except ModuleNotFoundError as exc:
    missing = exc.name or "必要ライブラリ"
    print(f"{missing} が見つかりません。")
    print("初回セットアップ.bat を実行して、必要ライブラリをインストールしてください。")
    raise


HOST = "127.0.0.1"
PORT = 8766
PROJECT_ROOT = Path(__file__).resolve().parent
KESSEN_TEMPLATE = PROJECT_ROOT / "data" / "templates" / "結線指示票 - ブランク.xls"
KOUTEN_TEMPLATE = PROJECT_ROOT / "data" / "templates" / "交点計算指示書 - ブランク.xlsx"


HTML = """<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>DXF帳票自動生成</title>
  <style>
    body {
      margin: 0;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #f5f7fa;
      color: #1f2937;
    }
    main {
      max-width: 920px;
      margin: 30px auto;
      padding: 0 18px;
    }
    h1 {
      font-size: 24px;
      margin: 0 0 18px;
      letter-spacing: 0;
    }
    section {
      background: #fff;
      border: 1px solid #d8dde6;
      border-radius: 8px;
      padding: 18px;
      margin-bottom: 14px;
    }
    label {
      display: block;
      font-weight: 650;
      margin-bottom: 8px;
    }
    input[type=file], input[type=text] {
      width: 100%;
      box-sizing: border-box;
      padding: 10px;
      border: 1px solid #c8ced8;
      border-radius: 6px;
      background: #fff;
      font-size: 15px;
    }
    .options {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 12px;
    }
    .check {
      display: flex;
      gap: 8px;
      align-items: center;
      padding: 10px;
      border: 1px solid #d8dde6;
      border-radius: 6px;
    }
    .check label {
      margin: 0;
      font-weight: 600;
    }
    button {
      border: 0;
      border-radius: 6px;
      background: #17613a;
      color: #fff;
      padding: 11px 18px;
      font-size: 15px;
      font-weight: 700;
      cursor: pointer;
    }
    button:disabled {
      background: #9aa8bd;
      cursor: wait;
    }
    pre {
      min-height: 180px;
      overflow: auto;
      white-space: pre-wrap;
      word-break: break-word;
      background: #111827;
      color: #e5e7eb;
      border-radius: 6px;
      padding: 14px;
      font-size: 14px;
      line-height: 1.55;
    }
    .hint {
      margin-top: 8px;
      color: #4b5563;
      font-size: 13px;
    }
    @media (max-width: 720px) {
      main { margin-top: 18px; }
      .options { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main>
    <h1>DXF帳票自動生成</h1>
    <section>
      <label for="dxfFile">DXFファイル</label>
      <input id="dxfFile" type="file" accept=".dxf,.DXF">
      <div class="hint">生成結果はZIPファイルとしてダウンロードされます。</div>
    </section>
    <section>
      <div class="options">
        <div class="check">
          <input id="genKessen" type="checkbox" checked>
          <label for="genKessen">結線指示票を生成</label>
        </div>
        <div class="check">
          <input id="genKouten" type="checkbox" checked>
          <label for="genKouten">交点計算指示書を生成</label>
        </div>
      </div>
    </section>
    <section>
      <label for="parcels">対象地番（任意・カンマ区切り）</label>
      <input id="parcels" type="text" placeholder="例: 1070-6,1069-5">
    </section>
    <section>
      <button id="runButton">帳票を生成</button>
    </section>
    <section>
      <label>結果</label>
      <pre id="log">DXFファイルを選択して「帳票を生成」を押してください。</pre>
    </section>
  </main>
  <script>
    const dxfInput = document.getElementById("dxfFile");
    const genKessen = document.getElementById("genKessen");
    const genKouten = document.getElementById("genKouten");
    const parcels = document.getElementById("parcels");
    const runButton = document.getElementById("runButton");
    const log = document.getElementById("log");

    function readAsBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(String(reader.result).split(",", 2)[1]);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
    }

    function downloadBase64(filename, base64Data) {
      const binary = atob(base64Data);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      const blob = new Blob([bytes], { type: "application/zip" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    }

    runButton.addEventListener("click", async () => {
      const file = dxfInput.files[0];
      if (!file) {
        log.textContent = "DXFファイルを選択してください。";
        return;
      }
      if (!genKessen.checked && !genKouten.checked) {
        log.textContent = "生成する帳票を1つ以上選択してください。";
        return;
      }

      runButton.disabled = true;
      log.textContent = "DXFを解析し、帳票を生成しています。ファイルが大きい場合は数分かかることがあります...";

      try {
        const payload = {
          dxf_name: file.name,
          dxf_data: await readAsBase64(file),
          gen_kessen: genKessen.checked,
          gen_kouten: genKouten.checked,
          parcels: parcels.value
        };
        const response = await fetch("/generate", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        const result = await response.json();
        if (!response.ok) {
          throw new Error(result.error || "帳票生成に失敗しました。");
        }
        log.textContent = result.summary;
        downloadBase64(result.output_name, result.output_data);
      } catch (error) {
        log.textContent = "エラーが発生しました。\\n\\n" + error.message;
      } finally {
        runButton.disabled = false;
      }
    });
  </script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        if self.path != "/":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = HTML.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self) -> None:
        if self.path != "/generate":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            result = generate_uploaded_dxf(payload)
            self.send_json(HTTPStatus.OK, result)
        except Exception as exc:
            detail = traceback.format_exc()
            self.send_json(HTTPStatus.BAD_REQUEST, {"error": f"{type(exc).__name__}: {exc}\n\n{detail}"})

    def send_json(self, status: HTTPStatus, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def safe_dxf_name(raw_name: str) -> str:
    name = Path(raw_name or "input.dxf").name
    return name if name.lower().endswith(".dxf") else f"{name}.dxf"


def parse_parcels(raw: str) -> list[str] | None:
    values = [value.strip() for value in (raw or "").split(",") if value.strip()]
    return values or None


def generate_uploaded_dxf(payload: dict) -> dict:
    if not KESSEN_TEMPLATE.is_file():
        raise FileNotFoundError(f"結線指示票テンプレートが見つかりません: {KESSEN_TEMPLATE}")
    if not KOUTEN_TEMPLATE.is_file():
        raise FileNotFoundError(f"交点計算指示書テンプレートが見つかりません: {KOUTEN_TEMPLATE}")

    dxf_name = safe_dxf_name(payload.get("dxf_name", "input.dxf"))
    dxf_bytes = base64.b64decode(payload["dxf_data"])
    gen_kessen = bool(payload.get("gen_kessen", True))
    gen_kouten = bool(payload.get("gen_kouten", True))
    target_parcels = parse_parcels(payload.get("parcels", ""))

    if not gen_kessen and not gen_kouten:
        raise ValueError("生成する帳票を1つ以上選択してください。")

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        dxf_path = tmp_path / dxf_name
        out_dir = tmp_path / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        dxf_path.write_bytes(dxf_bytes)

        parsed = parse_dxf(str(dxf_path))
        named_stakes = parsed.get_stakes_with_numbers()
        intersection_stakes = [stake for stake in named_stakes if stake.is_intersection]
        blocks = get_block_numbers(parsed)

        generated_files: list[Path] = []
        kessen_count = 0
        kouten_count = 0
        kessen_details: list[str] = []
        kouten_details: list[str] = []

        if gen_kessen:
            kessen_results = generate_kessen(parsed, target_parcels)
            if kessen_results:
                excel_results = [
                    ExcelKessenResult(
                        parcel_number=result.parcel_number,
                        stake_sequence=result.stake_sequence,
                        land_use=result.land_use,
                        owner=result.owner,
                    )
                    for result in kessen_results
                ]
                output_path = out_dir / "結線指示票.xls"
                write_kessen_excel(excel_results, str(KESSEN_TEMPLATE), str(output_path))
                generated_files.append(output_path)
                kessen_count = len(kessen_results)
                kessen_details = [
                    f"{result.parcel_number}: {len(result.stake_sequence)}杭"
                    for result in kessen_results[:80]
                ]

        if gen_kouten:
            kouten_results = generate_kouten(parsed, block_number=None)
            if kouten_results:
                excel_results = [
                    ExcelIntersectionResult(
                        intersection_stake=result.intersection_stake,
                        baseline_point1=result.baseline_point1,
                        baseline_point2=result.baseline_point2,
                        extension_point1=result.extension_point1,
                        extension_point2=result.extension_point2,
                    )
                    for result in kouten_results
                ]
                output_path = out_dir / "交点計算指示書.xlsx"
                write_kouten_excel(excel_results, str(KOUTEN_TEMPLATE), str(output_path))
                generated_files.append(output_path)
                kouten_count = len(kouten_results)
                kouten_details = [
                    (
                        f"交点={result.intersection_stake} "
                        f"基準線=({result.baseline_point1}, {result.baseline_point2}) "
                        f"延長線=({result.extension_point1}, {result.extension_point2})"
                    )
                    for result in kouten_results[:80]
                ]

        if not generated_files:
            raise ValueError("出力できる帳票がありませんでした。DXF内の地番・杭・交点杭を確認してください。")

        zip_bytes = io.BytesIO()
        with zipfile.ZipFile(zip_bytes, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for path in generated_files:
                zf.write(path, arcname=path.name)

        summary_lines = [
            "完了しました。",
            "",
            f"DXFファイル: {dxf_name}",
            f"杭(CIRCLE): {len(parsed.stakes)}",
            f"杭番号あり: {len(named_stakes)}",
            f"交点杭: {len(intersection_stakes)}",
            f"境界線: {len(parsed.lines)}",
            f"地番: {len(parsed.parcels)}",
            f"公共用地: {len(parsed.public_lands)}",
            f"検出ブロック番号: {', '.join(blocks) if blocks else 'なし'}",
            "",
            f"結線指示票: {kessen_count}筆" if gen_kessen else "結線指示票: 生成しない",
            f"交点計算指示書: {kouten_count}交点" if gen_kouten else "交点計算指示書: 生成しない",
            "",
            "ZIPファイルのダウンロードを開始しました。",
        ]
        if kessen_details:
            summary_lines.extend(["", "結線指示票 生成内容（先頭80件）:", *kessen_details])
        if kouten_details:
            summary_lines.extend(["", "交点計算指示書 生成内容（先頭80件）:", *kouten_details])

        return {
            "summary": "\n".join(summary_lines),
            "output_name": f"{Path(dxf_name).stem}_帳票.zip",
            "output_data": base64.b64encode(zip_bytes.getvalue()).decode("ascii"),
        }


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}/"
    print(f"DXF帳票自動生成ツールを起動しました: {url}")
    print("終了するには、この画面で Ctrl+C を押してください。")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n終了しました。")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
