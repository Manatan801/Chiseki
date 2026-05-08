#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def extract_docx_text(path: Path) -> str:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    paragraphs: list[str] = []
    for paragraph in root.findall(".//w:p", NS):
        parts = [node.text or "" for node in paragraph.findall(".//w:t", NS)]
        text = "".join(parts).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs)


def normalize_for_tts(text: str) -> str:
    lines = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        line = re.sub(r"^\s*(スライド|P|ページ)\s*\d+[\s:：.-]*", "", line)
        line = re.sub(r"^\s*\d+[\s:：.-]+", "", line)
        if line:
            lines.append(line)
    return "\n\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("docx", type=Path)
    parser.add_argument("--tts", action="store_true")
    args = parser.parse_args()

    text = extract_docx_text(args.docx)
    if args.tts:
        text = normalize_for_tts(text)
    print(text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
