#!/usr/bin/env python3
"""
BM CSV差分検出ツールのブラウザUI。

Python標準ライブラリだけでローカルHTTPサーバーを起動します。
ブラウザで2つのCSVを選択すると、差分CSVをダウンロードできます。
"""

from __future__ import annotations

import base64
import json
import tempfile
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote

from diff_bm_csv import compare_properties, count_by_type, parse_bm_csv, write_diff_csv


HOST = "127.0.0.1"
PORT = 8765


HTML = """<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>BM CSV 差分検出</title>
  <style>
    body {
      margin: 0;
      font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      background: #f6f7f9;
      color: #1f2937;
    }
    main {
      max-width: 860px;
      margin: 32px auto;
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
    input[type=file] {
      display: block;
      width: 100%;
      box-sizing: border-box;
      padding: 10px;
      border: 1px solid #c8ced8;
      border-radius: 6px;
      background: #fff;
      font-size: 15px;
    }
    .row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 14px;
    }
    button {
      border: 0;
      border-radius: 6px;
      background: #1f5fbf;
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
      min-height: 150px;
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
    @media (max-width: 720px) {
      .row { grid-template-columns: 1fr; }
      main { margin-top: 18px; }
    }
  </style>
</head>
<body>
  <main>
    <h1>BM CSV 差分検出</h1>
    <section class="row">
      <div>
        <label for="oldCsv">旧CSV</label>
        <input id="oldCsv" type="file" accept=".csv,text/csv">
      </div>
      <div>
        <label for="newCsv">新CSV</label>
        <input id="newCsv" type="file" accept=".csv,text/csv">
      </div>
    </section>
    <section>
      <button id="runButton">差分CSVを作成</button>
    </section>
    <section>
      <label>結果</label>
      <pre id="log">旧CSVと新CSVを選択して「差分CSVを作成」を押してください。</pre>
    </section>
  </main>
  <script>
    const oldInput = document.getElementById("oldCsv");
    const newInput = document.getElementById("newCsv");
    const runButton = document.getElementById("runButton");
    const log = document.getElementById("log");

    function readAsBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => {
          const dataUrl = reader.result;
          resolve(String(dataUrl).split(",", 2)[1]);
        };
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
    }

    function downloadBase64Csv(filename, base64Data) {
      const binary = atob(base64Data);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
      }
      const blob = new Blob([bytes], { type: "text/csv;charset=utf-8" });
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
      const oldFile = oldInput.files[0];
      const newFile = newInput.files[0];
      if (!oldFile || !newFile) {
        log.textContent = "旧CSVと新CSVを両方選択してください。";
        return;
      }

      runButton.disabled = true;
      log.textContent = "差分を検出しています...";

      try {
        const payload = {
          old_name: oldFile.name,
          new_name: newFile.name,
          old_data: await readAsBase64(oldFile),
          new_data: await readAsBase64(newFile)
        };
        const response = await fetch("/compare", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
        });
        const result = await response.json();
        if (!response.ok) {
          throw new Error(result.error || "差分検出に失敗しました。");
        }
        log.textContent = result.summary;
        downloadBase64Csv(result.output_name, result.output_data);
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
        if self.path != "/compare":
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        try:
            content_length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(content_length).decode("utf-8"))
            result = compare_uploaded_files(payload)
            self.send_json(HTTPStatus.OK, result)
        except Exception as exc:
            self.send_json(HTTPStatus.BAD_REQUEST, {"error": f"{type(exc).__name__}: {exc}"})

    def send_json(self, status: HTTPStatus, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, format: str, *args) -> None:
        return


def safe_name(name: str, fallback: str) -> str:
    name = Path(name or fallback).name
    return name if name.lower().endswith(".csv") else fallback


def compare_uploaded_files(payload: dict) -> dict:
    old_name = safe_name(payload.get("old_name", ""), "old.csv")
    new_name = safe_name(payload.get("new_name", ""), "new.csv")
    old_bytes = base64.b64decode(payload["old_data"])
    new_bytes = base64.b64decode(payload["new_data"])

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        old_path = tmp_path / old_name
        new_path = tmp_path / new_name
        output_path = tmp_path / f"{Path(old_name).stem}__{Path(new_name).stem}__diff.csv"
        old_path.write_bytes(old_bytes)
        new_path.write_bytes(new_bytes)

        old = parse_bm_csv(old_path)
        new = parse_bm_csv(new_path)
        rows = compare_properties(old, new)
        write_diff_csv(output_path, rows)

        summary = "\n".join(
            [
                "完了しました。",
                "",
                f"旧ファイル: {old_name}",
                f"新ファイル: {new_name}",
                f"旧物件数: {len(old)}",
                f"新物件数: {len(new)}",
                f"差分件数: {len(rows)}",
                f"  追加: {count_by_type(rows, '追加')}",
                f"  削除: {count_by_type(rows, '削除')}",
                f"  変更: {count_by_type(rows, '変更')}",
                "",
                "差分CSVのダウンロードを開始しました。",
            ]
        )

        return {
            "summary": summary,
            "output_name": quote(output_path.name, safe="._-"),
            "output_data": base64.b64encode(output_path.read_bytes()).decode("ascii"),
        }


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    url = f"http://{HOST}:{PORT}/"
    print(f"BM CSV差分ツールを起動しました: {url}")
    print("終了するには、このターミナルで Ctrl+C を押してください。")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n終了しました。")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
