"""
基盤地図情報 GML の建築物面 (BldA) を buildings.geojson に変換する。

既定入力: ../GML/FG-GML-*-ALL-*.zip
既定出力: data/buildings.geojson
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_GML_DIR = SCRIPT_DIR.parent / "GML"
DEFAULT_OUTPUT = SCRIPT_DIR / "data" / "buildings.geojson"

FGD_NS = "{http://fgd.gsi.go.jp/spec/2008/FGD_GMLSchema}"
GML_NS = "{http://www.opengis.net/gml/3.2}"


def text_of(parent: ET.Element, tag: str) -> str:
    return (parent.findtext(FGD_NS + tag) or "").strip()


def time_position(parent: ET.Element, tag: str) -> str:
    node = parent.find(FGD_NS + tag)
    if node is None:
        return ""
    return (node.findtext(GML_NS + "timePosition") or "").strip()


def ring_from_pos_list(pos_list: ET.Element) -> list[list[float]]:
    values = [float(value) for value in (pos_list.text or "").split()]
    if len(values) < 8 or len(values) % 2:
        return []

    # JPGIS/GML の緯度経度は lat lon 順。GeoJSON は lon lat 順。
    ring = [[lon, lat] for lat, lon in zip(values[0::2], values[1::2])]
    if ring and ring[0] != ring[-1]:
        ring.append(ring[0])
    if len(ring) < 4:
        return []
    return ring


def first_ring(parent: ET.Element, child_tag: str) -> list[list[float]]:
    child = parent.find(".//" + GML_NS + child_tag)
    if child is None:
        return []
    pos_list = child.find(".//" + GML_NS + "posList")
    if pos_list is None:
        return []
    return ring_from_pos_list(pos_list)


def iter_feature_records(zip_path: Path, xml_name: str):
    with zipfile.ZipFile(zip_path) as archive:
        with archive.open(xml_name) as fp:
            for _, elem in ET.iterparse(fp, events=("end",)):
                if elem.tag != FGD_NS + "BldA":
                    continue

                properties = {
                    "source_zip": zip_path.name,
                    "source_file": xml_name,
                    "source_mesh": zip_path.name.split("-")[2] if "-" in zip_path.name else "",
                    "fid": text_of(elem, "fid"),
                    "orgGILvl": text_of(elem, "orgGILvl"),
                    "lfSpanFr": time_position(elem, "lfSpanFr"),
                    "devDate": time_position(elem, "devDate"),
                }

                polygon_patches = elem.findall(".//" + GML_NS + "PolygonPatch")
                for patch_index, patch in enumerate(polygon_patches):
                    exterior = first_ring(patch, "exterior")
                    if not exterior:
                        continue

                    holes = []
                    for interior in patch.findall(".//" + GML_NS + "interior"):
                        pos_list = interior.find(".//" + GML_NS + "posList")
                        if pos_list is None:
                            continue
                        hole = ring_from_pos_list(pos_list)
                        if hole:
                            holes.append(hole)

                    props = dict(properties)
                    if len(polygon_patches) > 1:
                        props["patch_index"] = patch_index

                    yield {
                        "type": "Feature",
                        "properties": props,
                        "geometry": {
                            "type": "Polygon",
                            "coordinates": [exterior, *holes],
                        },
                    }

                elem.clear()


def find_blda_members(zip_path: Path) -> list[str]:
    with zipfile.ZipFile(zip_path) as archive:
        return [
            name
            for name in archive.namelist()
            if "-BldA-" in name and name.lower().endswith(".xml")
        ]


def update_bbox(bbox: list[float] | None, feature: dict) -> list[float]:
    ring = feature["geometry"]["coordinates"][0]
    lons = [point[0] for point in ring]
    lats = [point[1] for point in ring]
    current = [min(lons), min(lats), max(lons), max(lats)]
    if bbox is None:
        return current
    return [
        min(bbox[0], current[0]),
        min(bbox[1], current[1]),
        max(bbox[2], current[2]),
        max(bbox[3], current[3]),
    ]


def convert(gml_dir: Path, output_path: Path) -> int:
    zip_paths = sorted(gml_dir.glob("FG-GML-*-ALL-*.zip"))
    if not zip_paths:
        raise FileNotFoundError(f"GML ZIP が見つかりません: {gml_dir}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    count = 0
    bbox = None
    level_counts: dict[str, int] = {}

    with output_path.open("w", encoding="utf-8") as out:
        out.write('{"type":"FeatureCollection","features":[')
        first = True

        for zip_path in zip_paths:
            blda_members = find_blda_members(zip_path)
            if not blda_members:
                print(f"[skip] BldA なし: {zip_path.name}", file=sys.stderr)
                continue

            for xml_name in blda_members:
                print(f"変換中: {zip_path.name} / {xml_name}", file=sys.stderr)
                for feature in iter_feature_records(zip_path, xml_name):
                    if not first:
                        out.write(",")
                    first = False
                    json.dump(feature, out, ensure_ascii=False, separators=(",", ":"))

                    count += 1
                    bbox = update_bbox(bbox, feature)
                    level = feature["properties"].get("orgGILvl") or "unknown"
                    level_counts[level] = level_counts.get(level, 0) + 1

        metadata = {
            "source": "国土地理院 基盤地図情報 基本項目 建築物の外周線",
            "source_type": "FG-GML BldA",
            "source_directory": str(gml_dir),
            "feature_count": count,
            "bbox": bbox,
            "orgGILvl_counts": level_counts,
            "coordinate_order": "lon,lat",
        }
        out.write('],"metadata":')
        json.dump(metadata, out, ensure_ascii=False, separators=(",", ":"))
        out.write("}\n")

    return count


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--gml-dir", type=Path, default=DEFAULT_GML_DIR)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    count = convert(args.gml_dir, args.output)
    print(f"完了: {count:,} features -> {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
