#!/usr/bin/env python3
"""
登記情報提供サービス系BM CSV同士の差分検出ツール。

例:
  python3 youyakusyo/diff_bm_csv.py old.csv new.csv -o diff.csv
  python3 youyakusyo/diff_bm_csv.py --all-pairs youyakusyo/csv -o pair_summary.csv
"""

from __future__ import annotations

import argparse
import csv
import itertools
import re
import sys
import unicodedata
from collections import Counter
from dataclasses import dataclass
from pathlib import Path


GAIJI_MAP = {
    "<00008328>": "茨",
    "<00009AD9>": "髙",
    "<0000FA11>": "﨑",
    "<0000FA19>": "神",
    "<0000F9DC>": "隆",
    "<00005BEC>": "寬",
    "<00005EE3>": "廣",
    "<000080D6>": "胖",
    "<00008FC4>": "迄",
    "<000049FA>": "䧺",
    "<0000E452>": "〓",
    "<00100F20>": "〓",
    "<000F61BE>": "〓",
}

DIFF_HEADERS = [
    "差分種別",
    "物件キー",
    "所在",
    "地番",
    "項目",
    "変更前",
    "変更後",
    "旧ファイル",
    "新ファイル",
    "備考",
]


@dataclass
class Property:
    source_file: str
    number: str
    kind: str = ""
    status: str = ""
    shozai: str = ""
    shozai_full: str = ""
    aza: str = ""
    raw_chiban: str = ""
    chiban_display: str = ""
    fudosan_no: str = ""
    city_code: str = ""
    chimoku: str = ""
    chiseki: str = ""
    display_history: tuple[str, ...] = ()
    owners: tuple[dict[str, str], ...] = ()
    raw_rows: tuple[str, ...] = ()

    @property
    def display_shozai(self) -> str:
        return self.shozai_full or join_nonempty([self.shozai, self.aza], "")

    @property
    def display_chiban(self) -> str:
        return self.chiban_display or normalize_chiban(self.raw_chiban)

    @property
    def key(self) -> str:
        shozai_key = normalize_key_text(self.display_shozai)
        return f"{shozai_key}|{chiban_match_key(self.display_chiban)}"


def zen_to_han(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "")


def replace_gaiji(text: str) -> str:
    text = text or ""
    for code, char in GAIJI_MAP.items():
        text = text.replace(code, char)

    def replace_unknown(match: re.Match[str]) -> str:
        try:
            codepoint = int(match.group(1), 16)
            if 0 < codepoint < 0x110000:
                return chr(codepoint)
        except (OverflowError, ValueError):
            pass
        return "〓"

    return re.sub(r"<([0-9A-Fa-f]{8})>", replace_unknown, text)


def normalize_text(text: str) -> str:
    text = replace_gaiji(zen_to_han(text)).strip()
    text = text.replace("　", " ").replace("\u3000", " ")
    text = re.sub(r"\s+", " ", text)
    return text.replace("北茨城", "北茨城")


def normalize_key_text(text: str) -> str:
    return re.sub(r"\s+", "", normalize_text(text))


def normalize_chiban(raw: str) -> str:
    text = normalize_text(raw)
    text = text.replace("－", "-").replace("ー", "-").replace("—", "-")
    text = re.sub(r"(\d+)-(\d+)", r"\1番\2", text)
    return text


def chiban_match_key(raw: str) -> str:
    text = normalize_chiban(raw)
    text = text.replace("番の", "番")
    return re.sub(r"番$", "", text)


def normalize_chiseki(raw: str) -> str:
    text = normalize_text(raw).replace("・", ".").replace("‧", ".")
    if "." in text:
        text = text.rstrip("0").rstrip(".")
    return text


def normalize_name(raw: str) -> str:
    return normalize_key_text(raw)


def join_nonempty(values: list[str], sep: str) -> str:
    return sep.join(value for value in values if value)


def open_csv_reader(path: Path):
    encodings = ("cp932", "utf-8-sig", "utf-8")
    last_error: UnicodeDecodeError | None = None
    for encoding in encodings:
        try:
            f = path.open("r", encoding=encoding, newline="", errors="strict")
            # Force an early read so encoding errors are caught here.
            f.read(4096)
            f.seek(0)
            return f, csv.reader(f)
        except UnicodeDecodeError as exc:
            last_error = exc
    raise UnicodeDecodeError(
        last_error.encoding if last_error else "unknown",
        last_error.object if last_error else b"",
        last_error.start if last_error else 0,
        last_error.end if last_error else 1,
        f"{path} を cp932/utf-8 として読めません",
    )


def ordinary_csv_files(directory: Path) -> list[Path]:
    return sorted(
        p
        for p in directory.glob("*.csv")
        if p.is_file() and ":Zone.Identifier" not in p.name and ":mshield" not in p.name
    )


def parse_bm_csv(path: Path) -> dict[str, Property]:
    properties: dict[str, Property] = {}
    current: dict | None = None

    f, reader = open_csv_reader(path)
    with f:
        for raw_row in reader:
            if len(raw_row) < 2:
                continue

            row = [normalize_text(cell) for cell in raw_row]
            number = row[0].strip()
            record_type = row[1].strip()

            if record_type == "物件情報":
                if current:
                    save_property(properties, current)
                current = {
                    "source_file": path.name,
                    "number": number,
                    "kind": get(row, 2),
                    "status": get(row, 3),
                    "shozai": get(row, 4),
                    "raw_chiban": get(row, 5),
                    "fudosan_no": get(row, 6),
                    "aza": get(row, 9),
                    "city_code": get(row, 10),
                    "chimoku": "",
                    "chiseki": "",
                    "display_history": [],
                    "owners": [],
                    "raw_rows": [canonical_row(row)],
                }
                continue

            if current is None:
                continue

            current["raw_rows"].append(canonical_row(row))

            if record_type.startswith("所在"):
                current["shozai_full"] = get(row, 2)

            elif record_type.startswith("表示履歴"):
                chiban = get(row, 2)
                chimoku_raw = get(row, 4)
                chiseki_raw = get(row, 5)
                detail = get(row, 6)
                date = get(row, 7)

                if chiban and not current.get("chiban_display"):
                    current["chiban_display"] = normalize_chiban(chiban)

                chimoku = strip_old_value_marks(chimoku_raw)
                chiseki = strip_old_value_marks(chiseki_raw)
                if chimoku and not is_old_value(chimoku_raw):
                    current["chimoku"] = chimoku
                if chiseki and not is_old_value(chiseki_raw):
                    current["chiseki"] = chiseki

                current["display_history"].append(
                    canonical_text("|".join([chiban, chimoku_raw, chiseki_raw, detail, date]))
                )

            elif record_type.startswith("所有権") or record_type.startswith("所有者"):
                if record_type.startswith("所有者"):
                    owner = {
                        "address": "",
                        "mochibun": "",
                        "name": get(row, 2),
                        "reg_info": "",
                        "ref_num": "",
                    }
                else:
                    owner = {
                        "address": get(row, 2),
                        "mochibun": get(row, 3),
                        "name": get(row, 4),
                        "reg_info": get(row, 6),
                        "ref_num": get(row, 7),
                    }
                if owner["name"] or owner["address"]:
                    current["owners"].append(owner)

        if current:
            save_property(properties, current)

    return properties


def get(row: list[str], index: int) -> str:
    return row[index].strip() if len(row) > index else ""


def canonical_text(value: str) -> str:
    return normalize_text(value)


def canonical_row(row: list[str]) -> str:
    return ",".join(canonical_text(cell) for cell in row)


def is_old_value(value: str) -> bool:
    return bool(re.search(r"[【】\[\]]", value or ""))


def strip_old_value_marks(value: str) -> str:
    return re.sub(r"[【】\[\]]", "", value or "").strip()


def save_property(properties: dict[str, Property], raw: dict) -> None:
    prop = Property(
        source_file=raw["source_file"],
        number=raw.get("number", ""),
        kind=raw.get("kind", ""),
        status=raw.get("status", ""),
        shozai=raw.get("shozai", ""),
        shozai_full=raw.get("shozai_full", ""),
        aza=raw.get("aza", ""),
        raw_chiban=raw.get("raw_chiban", ""),
        chiban_display=raw.get("chiban_display", ""),
        fudosan_no=raw.get("fudosan_no", ""),
        city_code=raw.get("city_code", ""),
        chimoku=raw.get("chimoku", ""),
        chiseki=raw.get("chiseki", ""),
        display_history=tuple(raw.get("display_history", [])),
        owners=tuple(raw.get("owners", [])),
        raw_rows=tuple(raw.get("raw_rows", [])),
    )
    if not prop.display_chiban:
        return

    key = prop.key
    if key in properties:
        key = f"{key}|#{prop.number}"
    properties[key] = prop


def compare_properties(old: dict[str, Property], new: dict[str, Property]) -> list[dict[str, str]]:
    diffs: list[dict[str, str]] = []
    for key in sorted(set(old) | set(new)):
        old_prop = old.get(key)
        new_prop = new.get(key)
        prop = new_prop or old_prop
        assert prop is not None

        if old_prop is None:
            diffs.append(diff_row("追加", key, prop, "物件", "", summarize_property(new_prop), "", new_prop.source_file))
            continue
        if new_prop is None:
            diffs.append(diff_row("削除", key, prop, "物件", summarize_property(old_prop), "", old_prop.source_file, ""))
            continue

        diffs.extend(compare_common_property(key, old_prop, new_prop))

    return diffs


def compare_common_property(key: str, old: Property, new: Property) -> list[dict[str, str]]:
    diffs: list[dict[str, str]] = []
    scalar_fields = [
        ("種別", "kind", normalize_text),
        ("状態", "status", normalize_text),
        ("所在", "display_shozai", normalize_key_text),
        ("地番", "display_chiban", chiban_match_key),
        ("不動産番号", "fudosan_no", normalize_text),
        ("市区町村コード等", "city_code", normalize_text),
        ("地目", "chimoku", normalize_text),
        ("地積", "chiseki", normalize_chiseki),
    ]

    for label, attr, normalizer in scalar_fields:
        old_value = getattr(old, attr)
        new_value = getattr(new, attr)
        if normalizer(old_value) != normalizer(new_value):
            diffs.append(
                diff_row("変更", key, new, label, old_value, new_value, old.source_file, new.source_file)
            )

    if Counter(old.display_history) != Counter(new.display_history):
        diffs.append(
            diff_row(
                "変更",
                key,
                new,
                "表示履歴",
                "\n".join(old.display_history),
                "\n".join(new.display_history),
                old.source_file,
                new.source_file,
                "履歴行の追加・削除・内容変更があります",
            )
        )

    diffs.extend(compare_owners(key, old, new))
    return diffs


def owner_identity(owner: dict[str, str]) -> str:
    return normalize_name(owner.get("name", ""))


def owner_signature(owner: dict[str, str]) -> str:
    return "|".join(
        [
            normalize_key_text(owner.get("address", "")),
            normalize_key_text(owner.get("mochibun", "")),
            normalize_name(owner.get("name", "")),
            normalize_key_text(owner.get("reg_info", "")),
            normalize_key_text(owner.get("ref_num", "")),
        ]
    )


def compare_owners(key: str, old: Property, new: Property) -> list[dict[str, str]]:
    diffs: list[dict[str, str]] = []
    old_by_name = group_owners(old.owners)
    new_by_name = group_owners(new.owners)

    for name_key in sorted(set(old_by_name) | set(new_by_name)):
        old_items = old_by_name.get(name_key, [])
        new_items = new_by_name.get(name_key, [])
        old_counter = Counter(owner_signature(item) for item in old_items)
        new_counter = Counter(owner_signature(item) for item in new_items)
        if old_counter == new_counter:
            continue

        if not old_items:
            for owner in new_items:
                diffs.append(
                    diff_row("追加", key, new, "所有者", "", format_owner(owner), old.source_file, new.source_file)
                )
        elif not new_items:
            for owner in old_items:
                diffs.append(
                    diff_row("削除", key, old, "所有者", format_owner(owner), "", old.source_file, new.source_file)
                )
        else:
            diffs.append(
                diff_row(
                    "変更",
                    key,
                    new,
                    "所有者",
                    "\n".join(format_owner(item) for item in old_items),
                    "\n".join(format_owner(item) for item in new_items),
                    old.source_file,
                    new.source_file,
                )
            )

    return diffs


def group_owners(owners: tuple[dict[str, str], ...]) -> dict[str, list[dict[str, str]]]:
    grouped: dict[str, list[dict[str, str]]] = {}
    for index, owner in enumerate(owners):
        key = owner_identity(owner) or f"__blank_{index}"
        grouped.setdefault(key, []).append(owner)
    return grouped


def format_owner(owner: dict[str, str]) -> str:
    parts = [
        owner.get("name", ""),
        owner.get("mochibun", ""),
        owner.get("address", ""),
        owner.get("reg_info", ""),
        owner.get("ref_num", ""),
    ]
    return " / ".join(part for part in parts if part)


def summarize_property(prop: Property | None) -> str:
    if prop is None:
        return ""
    return " / ".join(
        part
        for part in [prop.display_shozai, prop.display_chiban, prop.chimoku, prop.chiseki]
        if part
    )


def diff_row(
    diff_type: str,
    key: str,
    prop: Property,
    item: str,
    old_value: str,
    new_value: str,
    old_file: str,
    new_file: str,
    note: str = "",
) -> dict[str, str]:
    return {
        "差分種別": diff_type,
        "物件キー": key,
        "所在": prop.display_shozai,
        "地番": prop.display_chiban,
        "項目": item,
        "変更前": old_value,
        "変更後": new_value,
        "旧ファイル": old_file,
        "新ファイル": new_file,
        "備考": note,
    }


def write_diff_csv(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=DIFF_HEADERS)
        writer.writeheader()
        writer.writerows(rows)


def write_pair_summary(path: Path, rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = ["旧ファイル", "新ファイル", "旧物件数", "新物件数", "差分件数", "追加", "削除", "変更"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def count_by_type(rows: list[dict[str, str]], diff_type: str) -> int:
    return sum(1 for row in rows if row["差分種別"] == diff_type)


def run_pair(old_csv: Path, new_csv: Path, output: Path) -> int:
    old = parse_bm_csv(old_csv)
    new = parse_bm_csv(new_csv)
    rows = compare_properties(old, new)
    write_diff_csv(output, rows)
    print(f"旧ファイル: {old_csv}")
    print(f"新ファイル: {new_csv}")
    print(f"旧物件数: {len(old)}")
    print(f"新物件数: {len(new)}")
    print(f"差分件数: {len(rows)}")
    print(f"  追加: {count_by_type(rows, '追加')}")
    print(f"  削除: {count_by_type(rows, '削除')}")
    print(f"  変更: {count_by_type(rows, '変更')}")
    print(f"出力: {output}")
    return 0


def run_all_pairs(directory: Path, output: Path) -> int:
    files = ordinary_csv_files(directory)
    if len(files) < 2:
        raise SystemExit(f"比較対象のCSVが2個以上必要です: {directory}")

    summary = []
    parsed_cache: dict[Path, dict[str, Property]] = {}
    for old_csv, new_csv in itertools.combinations(files, 2):
        old = parsed_cache.setdefault(old_csv, parse_bm_csv(old_csv))
        new = parsed_cache.setdefault(new_csv, parse_bm_csv(new_csv))
        rows = compare_properties(old, new)
        summary.append(
            {
                "旧ファイル": old_csv.name,
                "新ファイル": new_csv.name,
                "旧物件数": str(len(old)),
                "新物件数": str(len(new)),
                "差分件数": str(len(rows)),
                "追加": str(count_by_type(rows, "追加")),
                "削除": str(count_by_type(rows, "削除")),
                "変更": str(count_by_type(rows, "変更")),
            }
        )

    write_pair_summary(output, summary)
    print(f"比較ファイル数: {len(files)}")
    print(f"比較ペア数: {len(summary)}")
    print(f"出力: {output}")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="BM形式CSV同士の差分を検出します。")
    parser.add_argument("old_csv", nargs="?", type=Path, help="旧CSV")
    parser.add_argument("new_csv", nargs="?", type=Path, help="新CSV")
    parser.add_argument("-o", "--output", type=Path, default=Path("bm_csv_diff.csv"), help="出力CSV")
    parser.add_argument("--all-pairs", type=Path, help="指定フォルダ内の通常CSV全ペアを比較してサマリー出力")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)

    if args.all_pairs:
        return run_all_pairs(args.all_pairs, args.output)

    if not args.old_csv or not args.new_csv:
        raise SystemExit("旧CSVと新CSVを指定するか、--all-pairs フォルダ を指定してください。")

    return run_pair(args.old_csv, args.new_csv, args.output)


if __name__ == "__main__":
    sys.exit(main())
