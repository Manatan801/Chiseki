"""
傾斜区分分析ツール v2 (北茨城市 オフライン版)

起動: python terrain_web.py
ブラウザが自動的に http://127.0.0.1:8768/ を開きます。

必要ライブラリ: numpy, matplotlib (標準ライブラリのみで動作)
データ: data/kitaibaraki_dem.npz, data/map_tiles/
"""

import base64
import io
import json
import math
import re
import sys
import traceback
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import numpy as np

HOST = "127.0.0.1"
PORT = 8768
SCRIPT_DIR = Path(__file__).parent
DEM_NPZ    = SCRIPT_DIR / "data" / "kitaibaraki_dem.npz"
TILES_DIR  = SCRIPT_DIR / "data" / "map_tiles"
ASSETS_DIR = SCRIPT_DIR / "assets"

# 傾斜区分（国土調査事業事務取扱要領）
SLOPE_CLASSES = [
    (0,   5,  "平坦地",       "#2ecc71", (46, 204, 113)),
    (5,  15,  "緩傾斜地",     "#a8e063", (168, 224, 99)),
    (15, 25,  "中傾斜地",     "#f1c40f", (241, 196, 15)),
    (25, 35,  "急傾斜地(1)",  "#e67e22", (230, 126, 34)),
    (35, 45,  "急傾斜地(2)",  "#e74c3c", (231, 76, 60)),
    (45, 90,  "急峻地",       "#8e0000", (142, 0, 0)),
]

# ===== 起動時にDEMをメモリに読み込む =====
print("DEMデータを読み込み中...", end=" ", flush=True)
try:
    _data = np.load(str(DEM_NPZ))
    DEM = _data["dem"].astype(np.float64) / 10.0      # int16×10 → float メートル
    DEM[_data["dem"] == -32768] = np.nan               # NoData復元
    DEM_BOUNDS = tuple(float(x) for x in _data["bounds"])  # (north, south, west, east)
    print(f"完了 shape={DEM.shape} bounds=N{DEM_BOUNDS[0]:.3f} S{DEM_BOUNDS[1]:.3f}")
except Exception as e:
    print(f"\n[エラー] DEMデータの読み込みに失敗しました: {e}", file=sys.stderr)
    sys.exit(1)


# ===== 傾斜計算ロジック =====

SLOPE_METHOD = "3x3最小二乗平面フィット"


def calculate_slope(dem, pixel_size_x, pixel_size_y):
    """3x3近傍9点に z = ax + by + c を最小二乗フィットして傾斜角を返す。"""
    slope_deg = np.full(dem.shape, np.nan, dtype=np.float64)
    if dem.shape[0] < 3 or dem.shape[1] < 3:
        return slope_deg

    finite = np.isfinite(dem)
    neighborhood_valid = (
        finite[:-2, :-2] & finite[:-2, 1:-1] & finite[:-2, 2:] &
        finite[1:-1, :-2] & finite[1:-1, 1:-1] & finite[1:-1, 2:] &
        finite[2:, :-2] & finite[2:, 1:-1] & finite[2:, 2:]
    )

    left_sum = dem[:-2, :-2] + dem[1:-1, :-2] + dem[2:, :-2]
    right_sum = dem[:-2, 2:] + dem[1:-1, 2:] + dem[2:, 2:]
    top_sum = dem[:-2, :-2] + dem[:-2, 1:-1] + dem[:-2, 2:]
    bottom_sum = dem[2:, :-2] + dem[2:, 1:-1] + dem[2:, 2:]

    dz_dx = (right_sum - left_sum) / (6.0 * pixel_size_x)
    dz_dy = (bottom_sum - top_sum) / (6.0 * pixel_size_y)
    slope_core = np.degrees(np.arctan(np.sqrt(dz_dx**2 + dz_dy**2)))
    slope_inner = slope_deg[1:-1, 1:-1]
    slope_inner[neighborhood_valid] = slope_core[neighborhood_valid]
    return slope_deg


def classify_slope(slope_deg):
    result = np.full(slope_deg.shape, -1, dtype=np.int8)
    for i, (lo, hi, _, _, _) in enumerate(SLOPE_CLASSES):
        mask = (slope_deg >= lo) & (slope_deg < hi)
        result[mask] = i
    return result


def dem_pixel_sizes_at_lat(lat, bounds=DEM_BOUNDS, shape=DEM.shape):
    """DEMの実範囲と配列サイズから、地図ズームに依存しない地上寸法を返す。"""
    north, south, west, east = bounds
    h, w = shape
    earth_radius = 6378137.0
    meters_per_deg_lat = math.pi * earth_radius / 180.0
    meters_per_deg_lon = meters_per_deg_lat * math.cos(math.radians(lat))
    pixel_size_x = ((east - west) / w) * meters_per_deg_lon
    pixel_size_y = ((north - south) / h) * meters_per_deg_lat
    return pixel_size_x, pixel_size_y


def analyze_polygon(polygon_coords):
    """
    polygon_coords: [[lon, lat], ...] GeoJSON形式

    Returns: stats, sub_north, sub_south, sub_west, sub_east, total_area_m2, extra_meta
    """
    from matplotlib.path import Path as MplPath

    north, south, west, east = DEM_BOUNDS
    h, w = DEM.shape

    lons = [c[0] for c in polygon_coords]
    lats = [c[1] for c in polygon_coords]

    # 範囲チェック
    if not (south <= min(lats) and max(lats) <= north and
            west  <= min(lons) and max(lons) <= east):
        raise ValueError(
            f"選択エリアが北茨城市のDEMデータ範囲外です。\n"
            f"データ範囲: N{north:.3f} S{south:.3f} W{west:.3f} E{east:.3f}\n"
            f"選択範囲:   N{max(lats):.3f} S{min(lats):.3f} W{min(lons):.3f} E{max(lons):.3f}"
        )

    # ポリゴンbboxでDEMをサブセット（全体スキャン回避）
    col0 = max(0, int((min(lons) - west)  / (east  - west)  * w) - 4)
    col1 = min(w, int((max(lons) - west)  / (east  - west)  * w) + 4)
    row0 = max(0, int((north - max(lats)) / (north - south) * h) - 4)
    row1 = min(h, int((north - min(lats)) / (north - south) * h) + 4)

    dem_sub = DEM[row0:row1, col0:col1]
    sub_north = north - row0 * (north - south) / h
    sub_south = north - row1 * (north - south) / h
    sub_west  = west  + col0 * (east  - west)  / w
    sub_east  = west  + col1 * (east  - west)  / w

    sh, sw = dem_sub.shape
    mid_lat = (sub_north + sub_south) / 2
    px, py = dem_pixel_sizes_at_lat(mid_lat)

    slope_deg  = calculate_slope(dem_sub, px, py)
    classified = classify_slope(slope_deg)

    # ポリゴンマスク
    sub_lats = np.linspace(sub_north, sub_south, sh)
    sub_lons = np.linspace(sub_west,  sub_east,  sw)
    lon_grid, lat_grid = np.meshgrid(sub_lons, sub_lats)
    points = np.column_stack([lon_grid.ravel(), lat_grid.ravel()])
    path = MplPath([(c[0], c[1]) for c in polygon_coords])
    mask = path.contains_points(points).reshape(sh, sw)

    valid_mask   = mask & (classified >= 0)
    total_pixels = int(valid_mask.sum())
    pixel_area   = px * py  # m²

    stats = []
    for i, (_, _, name, color, _) in enumerate(SLOPE_CLASSES):
        count   = int(((classified == i) & valid_mask).sum())
        area_m2 = round(count * pixel_area, 1)
        percent = round(count / total_pixels * 100, 2) if total_pixels > 0 else 0.0
        stats.append({"name": name, "area_m2": area_m2, "percent": percent, "color": color})

    # 追加統計
    slope_in_mask = slope_deg[valid_mask]
    dem_in_mask   = dem_sub[valid_mask & ~np.isnan(dem_sub)]
    extra_meta = {
        "mean_slope": round(float(np.nanmean(slope_in_mask)), 2) if slope_in_mask.size > 0 else None,
        "max_slope":  round(float(np.nanmax(slope_in_mask)),  2) if slope_in_mask.size > 0 else None,
        "elev_min":   round(float(np.nanmin(dem_in_mask)),    1) if dem_in_mask.size  > 0 else None,
        "elev_max":   round(float(np.nanmax(dem_in_mask)),    1) if dem_in_mask.size  > 0 else None,
        "elev_mean":  round(float(np.nanmean(dem_in_mask)),   1) if dem_in_mask.size  > 0 else None,
        "slope_method": SLOPE_METHOD,
        "pixel_size_x": round(float(px), 3),
        "pixel_size_y": round(float(py), 3),
    }

    return stats, sub_north, sub_south, sub_west, sub_east, total_pixels * pixel_area, extra_meta, \
           classified, mask, sh, sw, sub_north, sub_south, sub_west, sub_east


def make_overlay_png(classified, mask, sh, sw):
    """傾斜区分ごとの色でRGBA画像を生成してbase64 PNG文字列を返す"""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    rgba = np.zeros((sh, sw, 4), dtype=np.uint8)
    for i, (_, _, _, _, rgb) in enumerate(SLOPE_CLASSES):
        cls_mask = (classified == i) & mask
        rgba[cls_mask, 0] = rgb[0]
        rgba[cls_mask, 1] = rgb[1]
        rgba[cls_mask, 2] = rgb[2]
        rgba[cls_mask, 3] = 200  # 約80%不透明

    buf = io.BytesIO()
    plt.imsave(buf, rgba, format="png")
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")


def make_pie_chart(stats):
    """統計データからmatplotlibの円グラフをbase64 PNG文字列で返す"""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.font_manager as fm

    jp_fonts = [f for f in fm.findSystemFonts() if any(
        kw in f.lower() for kw in ("meiryo", "gothic", "mincho", "hiragino", "msgothic")
    )]
    if jp_fonts:
        plt.rcParams["font.family"] = fm.FontProperties(fname=jp_fonts[0]).get_name()
    else:
        plt.rcParams["font.family"] = "sans-serif"

    labels = [s["name"]    for s in stats if s["percent"] > 0]
    sizes  = [s["percent"] for s in stats if s["percent"] > 0]
    colors = [s["color"]   for s in stats if s["percent"] > 0]

    fig, ax = plt.subplots(figsize=(6, 4))
    if sizes:
        ax.pie(sizes, labels=labels, colors=colors, autopct="%1.1f%%",
               startangle=90, pctdistance=0.8)
        ax.axis("equal")
    ax.set_title("傾斜区分割合", fontsize=13)

    buf = io.BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format="png", dpi=110, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("ascii")


# ===== アセット読み込み（起動時に1回） =====

def _read_asset(name):
    path = ASSETS_DIR / name
    if path.exists():
        return path.read_text(encoding="utf-8", errors="replace")
    return ""


def _asset_b64(name, mime):
    path = ASSETS_DIR / name
    if path.exists():
        return f"data:{mime};base64," + base64.b64encode(path.read_bytes()).decode()
    return ""


def _patch_draw_css(css):
    """leaflet.draw.css の url('images/...') をbase64データURIに置換"""
    mime_map = {
        "spritesheet.png":    ("images/spritesheet.png",    "image/png"),
        "spritesheet-2x.png": ("images/spritesheet-2x.png", "image/png"),
        "spritesheet.svg":    ("images/spritesheet.svg",    "image/svg+xml"),
    }
    for fname, (relpath, mime) in mime_map.items():
        data_uri = _asset_b64(relpath, mime)
        if data_uri:
            css = css.replace(f"url('images/{fname}')", f"url('{data_uri}')")
    return css


LEAFLET_CSS     = _read_asset("leaflet.css")
DRAW_CSS        = _patch_draw_css(_read_asset("leaflet.draw.css"))
LEAFLET_JS      = _read_asset("leaflet.js")
LEAFLET_DRAW_JS = _read_asset("leaflet.draw.js")

HTML = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>傾斜区分分析 v2 - 北茨城市</title>
<style>
{LEAFLET_CSS}
{DRAW_CSS}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{ font-family: "Meiryo", "Yu Gothic", sans-serif; background: #f0f2f5; color: #2c3e50; }}
header {{ background: #1a5276; color: #fff; padding: 12px 20px; display: flex; align-items: center; gap: 12px; }}
header h1 {{ font-size: 18px; font-weight: 700; }}
header small {{ font-size: 12px; opacity: 0.75; }}
#map {{ width: 100%; height: 520px; border-bottom: 2px solid #2980b9; }}
#controls {{ padding: 10px 20px; background: #fff; border-bottom: 1px solid #ddd; display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }}
#draw-btns {{ display: flex; gap: 8px; }}
.btn-draw {{ background: #27ae60; color: #fff; border: none; padding: 8px 16px; border-radius: 5px; font-size: 13px; cursor: pointer; font-weight: 600; }}
.btn-draw:hover {{ background: #1e8449; }}
.btn-draw.active {{ background: #154360; box-shadow: inset 0 2px 4px rgba(0,0,0,0.3); }}
#status {{ font-size: 13px; color: #555; flex: 1; min-width: 160px; }}
#btn-analyze {{ background: #2980b9; color: #fff; border: none; padding: 9px 22px; border-radius: 5px; font-size: 14px; cursor: pointer; font-weight: 600; }}
#btn-analyze:hover {{ background: #1a6fa0; }}
#btn-analyze:disabled {{ background: #95a5a6; cursor: not-allowed; }}
#btn-clear {{ background: #e74c3c; color: #fff; border: none; padding: 9px 14px; border-radius: 5px; font-size: 13px; cursor: pointer; }}
#results {{ padding: 20px; display: none; }}
#results h2 {{ font-size: 16px; margin-bottom: 12px; color: #1a5276; }}
#err-msg {{ background: #fdedec; border-left: 4px solid #e74c3c; padding: 12px 16px; border-radius: 4px; font-size: 14px; white-space: pre-wrap; display: none; margin: 12px 20px; }}
table {{ width: 100%; border-collapse: collapse; font-size: 14px; margin-bottom: 12px; }}
th {{ background: #1a5276; color: #fff; padding: 8px 12px; text-align: left; }}
td {{ padding: 8px 12px; border-bottom: 1px solid #eee; }}
tr:nth-child(even) td {{ background: #f8fafc; }}
.swatch {{ display: inline-block; width: 14px; height: 14px; border-radius: 3px; vertical-align: middle; margin-right: 6px; border: 1px solid rgba(0,0,0,0.2); }}
#chart-img {{ max-width: 500px; display: block; }}
#meta {{ font-size: 12px; color: #7f8c8d; margin-top: 8px; line-height: 1.7; }}
.spinner {{ display: inline-block; width: 14px; height: 14px; border: 2px solid #ccc; border-top-color: #2980b9; border-radius: 50%; animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 6px; }}
@keyframes spin {{ to {{ transform: rotate(360deg); }} }}
#overlay-ctrl {{ display: flex; align-items: center; gap: 10px; margin: 8px 0 4px; font-size: 13px; }}
#btn-csv {{ background: #16a085; color: #fff; border: none; padding: 7px 16px; border-radius: 5px; font-size: 13px; cursor: pointer; margin-top: 4px; }}
#btn-csv:hover {{ background: #0e6655; }}
.info.legend {{ background: white; padding: 8px 12px; border-radius: 6px; line-height: 22px; font-size: 12px; box-shadow: 0 1px 5px rgba(0,0,0,0.4); }}
.info.legend i {{ width: 14px; height: 14px; display: inline-block; margin-right: 6px; border-radius: 3px; vertical-align: middle; }}
</style>
</head>
<body>
<header>
  <span style="font-size:24px;">&#128302;</span>
  <div>
    <h1>傾斜区分分析ツール v2</h1>
    <small>北茨城市専用 v2 オフライン版 ／ 3x3最小二乗平面フィット ／ 高詳細地図 z=17 収録済み</small>
  </div>
</header>

<div id="map"></div>

<div id="controls">
  <div id="draw-btns">
    <button class="btn-draw" id="btn-polygon">&#11044; ポリゴンで指定</button>
    <button class="btn-draw" id="btn-rect">&#9645; 長方形で指定</button>
  </div>
  <span id="status">地図上に解析エリアを描いてください</span>
  <button id="btn-analyze" disabled>傾斜を解析する</button>
  <button id="btn-clear">クリア</button>
</div>

<div id="err-msg"></div>
<div id="results">
  <h2>傾斜区分別 面積・割合</h2>
  <div id="overlay-ctrl">
    <label><input type="checkbox" id="chk-overlay" checked> 地図にオーバーレイ表示</label>
    <span>透明度: <input type="range" id="overlay-opacity" min="0" max="100" value="60" style="width:100px;vertical-align:middle;"> <span id="opacity-val">60</span>%</span>
  </div>
  <table id="stats-table">
    <thead><tr><th>傾斜区分</th><th>面積 (ha)</th><th>面積 (m²)</th><th>割合 (%)</th></tr></thead>
    <tbody></tbody>
  </table>
  <button id="btn-csv">CSV ダウンロード</button>
  <h2 style="margin-top:16px;">傾斜区分 割合グラフ</h2>
  <img id="chart-img" src="" alt="円グラフ">
  <p id="meta"></p>
</div>

<script>
{LEAFLET_JS}
</script>
<script>
{LEAFLET_DRAW_JS}
</script>
<script>
// ===== Leaflet 1.9 × Leaflet.draw 1.0.4 互換パッチ =====

// issue #1026: マウスのポインタイベントで _onTouch が誤発火する対策
(function() {{
  var orig = L.Draw.Polyline.prototype._onTouch;
  L.Draw.Polyline.prototype._onTouch = function(e) {{
    var oe = e.originalEvent;
    if (oe && oe.pointerType && oe.pointerType !== 'touch') return;
    return orig.call(this, e);
  }};
}})();

// issue #940: readableArea の type 未定義バグを修正
L.GeometryUtil.readableArea = function(area, isMetric) {{
  if (isMetric) {{
    return area >= 10000 ? (area / 10000).toFixed(2) + ' ha' : Math.round(area) + ' m²';
  }}
  return (area * 10.7639).toFixed(2) + ' sq ft';
}};

// L.drawLocal 日本語化
L.drawLocal.draw.toolbar.actions.title  = 'キャンセル';
L.drawLocal.draw.toolbar.actions.text   = 'キャンセル';
L.drawLocal.draw.toolbar.finish.title   = '描画を完了する';
L.drawLocal.draw.toolbar.finish.text    = '完了';
L.drawLocal.draw.toolbar.undo.title     = '最後の頂点を削除';
L.drawLocal.draw.toolbar.undo.text      = '最後の点を削除';
L.drawLocal.draw.handlers.polygon.tooltip.start   = 'クリックで頂点を追加';
L.drawLocal.draw.handlers.polygon.tooltip.cont    = 'クリックで続ける、ダブルクリックで完了';
L.drawLocal.draw.handlers.polygon.tooltip.end     = '最初の点をクリックして閉じる';
L.drawLocal.draw.handlers.rectangle.tooltip.start = 'クリックしてドラッグで長方形を描く';
L.drawLocal.edit.toolbar.actions.save.title = '変更を保存';
L.drawLocal.edit.toolbar.actions.save.text  = '保存';
L.drawLocal.edit.toolbar.actions.cancel.title = '編集をキャンセル';
L.drawLocal.edit.toolbar.actions.cancel.text  = 'キャンセル';
L.drawLocal.edit.toolbar.actions.clearAll.title = 'すべてクリア';
L.drawLocal.edit.toolbar.actions.clearAll.text  = 'すべてクリア';
L.drawLocal.edit.handlers.edit.tooltip.text    = '頂点をドラッグして編集';
L.drawLocal.edit.handlers.edit.tooltip.subtext = 'クリックで元に戻す';
L.drawLocal.edit.handlers.remove.tooltip.text  = 'クリックで削除';

// ===== 地図初期化 =====
var map = L.map('map').setView([36.85, 140.67], 17);

L.tileLayer('http://127.0.0.1:{PORT}/tiles/{{z}}/{{x}}/{{y}}.png', {{
  minZoom: 10,
  maxNativeZoom: 17,
  maxZoom: 17,
  attribution: '© 国土地理院（淡色地図・標準地図）',
  tileSize: 256
}}).addTo(map);

// ===== 描画ツール =====
var drawnItems = new L.FeatureGroup();
map.addLayer(drawnItems);

var drawControl = new L.Control.Draw({{
  edit: {{ featureGroup: drawnItems }},
  draw: {{
    polygon:      {{ shapeOptions: {{ color: '#2980b9', fillOpacity: 0.2 }} }},
    rectangle:    {{ shapeOptions: {{ color: '#2980b9', fillOpacity: 0.2 }} }},
    polyline: false, circle: false, marker: false, circlemarker: false
  }}
}});
map.addControl(drawControl);

var currentPolygon = null;
var overlayLayer   = null;
var legendControl  = null;
var currentStats   = null;
var currentMeta    = null;
var activeDrawer   = null;

// ===== 日本語ボタンから描画を直接起動 =====
var polygonDrawOpts  = {{ shapeOptions: {{ color: '#2980b9', fillOpacity: 0.2 }} }};
var rectangleDrawOpts= {{ shapeOptions: {{ color: '#2980b9', fillOpacity: 0.2 }} }};

function cancelActiveDrawer() {{
  if (activeDrawer) {{ try {{ activeDrawer.disable(); }} catch(e) {{}} activeDrawer = null; }}
  document.querySelectorAll('.btn-draw').forEach(function(b) {{ b.classList.remove('active'); }});
}}

document.getElementById('btn-polygon').onclick = function() {{
  if (activeDrawer instanceof L.Draw.Polygon) {{ cancelActiveDrawer(); return; }}
  cancelActiveDrawer();
  activeDrawer = new L.Draw.Polygon(map, polygonDrawOpts);
  activeDrawer.enable();
  this.classList.add('active');
  setStatus('クリックで頂点を追加 / ダブルクリックで完了');
}};

document.getElementById('btn-rect').onclick = function() {{
  if (activeDrawer instanceof L.Draw.Rectangle) {{ cancelActiveDrawer(); return; }}
  cancelActiveDrawer();
  activeDrawer = new L.Draw.Rectangle(map, rectangleDrawOpts);
  activeDrawer.enable();
  this.classList.add('active');
  setStatus('クリックしてドラッグで長方形を描く');
}};

function setStatus(msg) {{
  document.getElementById('status').textContent = msg;
}}

map.on(L.Draw.Event.CREATED, function(e) {{
  drawnItems.clearLayers();
  drawnItems.addLayer(e.layer);
  currentPolygon = e.layer;
  activeDrawer   = null;
  document.querySelectorAll('.btn-draw').forEach(function(b) {{ b.classList.remove('active'); }});
  setStatus('エリアを検出しました。「傾斜を解析する」を押してください。');
  document.getElementById('btn-analyze').disabled = false;
  hideResults();
}});

map.on(L.Draw.Event.EDITED, function(e) {{
  e.layers.eachLayer(function(layer) {{ currentPolygon = layer; }});
  hideResults();
}});

map.on(L.Draw.Event.DELETED, function() {{
  currentPolygon = null;
  document.getElementById('btn-analyze').disabled = true;
  setStatus('地図上に解析エリアを描いてください');
  hideResults();
}});

document.getElementById('btn-clear').onclick = function() {{
  cancelActiveDrawer();
  drawnItems.clearLayers();
  currentPolygon = null;
  document.getElementById('btn-analyze').disabled = true;
  setStatus('地図上に解析エリアを描いてください');
  hideResults();
}};

function hideResults() {{
  document.getElementById('results').style.display = 'none';
  document.getElementById('err-msg').style.display = 'none';
  removeOverlay();
  removeLegend();
  currentStats = null;
  currentMeta  = null;
}}

// ===== 解析実行 =====
document.getElementById('btn-analyze').onclick = function() {{
  if (!currentPolygon) return;

  var coords;
  if (currentPolygon instanceof L.Rectangle || currentPolygon instanceof L.Polygon) {{
    coords = currentPolygon.getLatLngs()[0].map(function(ll) {{ return [ll.lng, ll.lat]; }});
    if (coords[0][0] !== coords[coords.length-1][0] || coords[0][1] !== coords[coords.length-1][1]) {{
      coords.push(coords[0]);
    }}
  }} else {{
    return;
  }}

  var btn = document.getElementById('btn-analyze');
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>解析中...';
  setStatus('DEMデータを解析中...');
  hideResults();

  fetch('/analyze', {{
    method: 'POST',
    headers: {{'Content-Type': 'application/json'}},
    body: JSON.stringify({{ polygon: coords }})
  }})
  .then(function(r) {{ return r.json(); }})
  .then(function(data) {{
    btn.disabled = false;
    btn.textContent = '傾斜を解析する';
    if (data.error) {{
      var el = document.getElementById('err-msg');
      el.textContent = '⚠️ ' + data.error;
      el.style.display = 'block';
      setStatus('エラーが発生しました');
      return;
    }}
    currentStats = data.stats;
    currentMeta  = data.meta;
    showResults(data);
    setStatus('解析完了');
  }})
  .catch(function(e) {{
    btn.disabled = false;
    btn.textContent = '傾斜を解析する';
    setStatus('通信エラー: ' + e.message);
  }});
}};

// ===== オーバーレイ =====
function removeOverlay() {{
  if (overlayLayer) {{ map.removeLayer(overlayLayer); overlayLayer = null; }}
}}

function addOverlay(b64, meta) {{
  removeOverlay();
  var bounds = [[meta.south, meta.west], [meta.north, meta.east]];
  var opacity = document.getElementById('overlay-opacity').value / 100;
  overlayLayer = L.imageOverlay('data:image/png;base64,' + b64, bounds, {{
    opacity: opacity, interactive: false
  }}).addTo(map);
}}

function removeLegend() {{
  if (legendControl) {{ map.removeControl(legendControl); legendControl = null; }}
}}

function addLegend(stats) {{
  removeLegend();
  legendControl = L.control({{ position: 'bottomright' }});
  legendControl.onAdd = function() {{
    var div = L.DomUtil.create('div', 'info legend');
    stats.forEach(function(s) {{
      if (s.percent > 0) {{
        div.innerHTML += '<i style="background:' + s.color + '"></i>' +
          s.name + ' (' + s.percent + '%)<br>';
      }}
    }});
    return div;
  }};
  legendControl.addTo(map);
}}

document.getElementById('chk-overlay').onchange = function() {{
  if (!overlayLayer) return;
  if (this.checked) {{
    overlayLayer.setOpacity(document.getElementById('overlay-opacity').value / 100);
  }} else {{
    overlayLayer.setOpacity(0);
  }}
}};

document.getElementById('overlay-opacity').oninput = function() {{
  document.getElementById('opacity-val').textContent = this.value;
  if (overlayLayer && document.getElementById('chk-overlay').checked) {{
    overlayLayer.setOpacity(this.value / 100);
  }}
}};

// ===== 結果表示 =====
function showResults(data) {{
  var tbody = document.querySelector('#stats-table tbody');
  tbody.innerHTML = '';
  data.stats.forEach(function(s) {{
    var ha = (s.area_m2 / 10000).toFixed(3);
    var tr = document.createElement('tr');
    tr.innerHTML =
      '<td><span class="swatch" style="background:' + s.color + '"></span>' + s.name + '</td>' +
      '<td>' + ha + '</td>' +
      '<td>' + s.area_m2.toLocaleString() + '</td>' +
      '<td>' + s.percent + '</td>';
    tbody.appendChild(tr);
  }});

  if (data.chart_b64) {{
    document.getElementById('chart-img').src = 'data:image/png;base64,' + data.chart_b64;
  }}

  if (data.overlay_b64) {{
    addOverlay(data.overlay_b64, data.meta);
    addLegend(data.stats);
  }}

  var m = data.meta;
  var metaParts = [
    '解析範囲: N' + m.north.toFixed(4) + ' S' + m.south.toFixed(4) +
      ' W' + m.west.toFixed(4) + ' E' + m.east.toFixed(4),
    'DEM格子寸法: 約' + m.pixel_size_x.toFixed(3) + 'm × ' + m.pixel_size_y.toFixed(3) + 'm',
    '傾斜計算: ' + m.slope_method,
    '集計面積: ' + m.area_ha.toFixed(2) + ' ha',
  ];
  if (m.mean_slope != null) metaParts.push('平均傾斜: ' + m.mean_slope + '° / 最大傾斜: ' + m.max_slope + '°');
  if (m.elev_min   != null) metaParts.push('標高: 最低 ' + m.elev_min + 'm ～ 最高 ' + m.elev_max + 'm（平均 ' + m.elev_mean + 'm）');
  document.getElementById('meta').textContent = metaParts.join('  ／  ');

  document.getElementById('results').style.display = 'block';
  document.getElementById('err-msg').style.display = 'none';
  document.getElementById('results').scrollIntoView({{ behavior: 'smooth' }});
}}

// ===== CSV ダウンロード =====
document.getElementById('btn-csv').onclick = function() {{
  if (!currentStats || !currentMeta) return;
  var m = currentMeta;
  var now = new Date().toLocaleString('ja-JP');
  var lines = [
    '﻿傾斜区分,面積(ha),面積(m2),割合(%)',  // BOM付き
  ];
  currentStats.forEach(function(s) {{
    lines.push([s.name, (s.area_m2/10000).toFixed(3), s.area_m2, s.percent].join(','));
  }});
  lines.push('');
  lines.push('# 解析情報');
  lines.push('解析日時,' + now);
  lines.push('集計面積(ha),' + m.area_ha.toFixed(2));
  lines.push('解析範囲N,' + m.north.toFixed(6));
  lines.push('解析範囲S,' + m.south.toFixed(6));
  lines.push('解析範囲W,' + m.west.toFixed(6));
  lines.push('解析範囲E,' + m.east.toFixed(6));
  lines.push('傾斜計算方式,' + m.slope_method);
  lines.push('DEM格子寸法X(m),' + m.pixel_size_x.toFixed(3));
  lines.push('DEM格子寸法Y(m),' + m.pixel_size_y.toFixed(3));
  if (m.mean_slope != null) {{
    lines.push('平均傾斜(deg),' + m.mean_slope);
    lines.push('最大傾斜(deg),' + m.max_slope);
  }}
  if (m.elev_min != null) {{
    lines.push('最低標高(m),' + m.elev_min);
    lines.push('最高標高(m),' + m.elev_max);
    lines.push('平均標高(m),' + m.elev_mean);
  }}

  var blob = new Blob([lines.join('\\r\\n')], {{ type: 'text/csv;charset=utf-8;' }});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '傾斜区分分析_' + now.replace(/[:/\\s]/g, '-') + '.csv';
  a.click();
}};
</script>
</body>
</html>
"""


# ===== HTTP ハンドラ =====

class Handler(BaseHTTPRequestHandler):

    def do_GET(self):
        if self.path == "/":
            body = HTML.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif self.path.startswith("/tiles/"):
            self._serve_tile()

        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        if self.path == "/analyze":
            self._handle_analyze()
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def _serve_tile(self):
        parts = self.path[len("/tiles/"):].split("/")
        if len(parts) != 3:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        z, x = parts[0], parts[1]
        y = parts[2].replace(".png", "")
        tile_path = TILES_DIR / z / x / f"{y}.png"
        if not re.fullmatch(r"\d+", z) or not re.fullmatch(r"\d+", x) or not re.fullmatch(r"\d+", y):
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        if tile_path.exists():
            data = tile_path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "image/png")
            self.send_header("Content-Length", str(len(data)))
            self.send_header("Cache-Control", "max-age=86400")
            self.end_headers()
            self.wfile.write(data)
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def _handle_analyze(self):
        try:
            length  = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            polygon = payload.get("polygon", [])
            if len(polygon) < 4:
                raise ValueError("ポリゴンの頂点が不足しています（4点以上必要）")

            result = analyze_polygon(polygon)
            stats, sn, ss, sw, se, area_m2, extra_meta = result[:7]
            classified, mask, sh, sw2, *_ = result[7:]

            chart_b64   = make_pie_chart(stats)
            overlay_b64 = make_overlay_png(classified, mask, sh, sw2)

            self._send_json(HTTPStatus.OK, {
                "stats":       stats,
                "chart_b64":   chart_b64,
                "overlay_b64": overlay_b64,
                "meta": {
                    "north":      round(sn, 6),
                    "south":      round(ss, 6),
                    "west":       round(sw, 6),
                    "east":       round(se, 6),
                    "area_ha":    round(area_m2 / 10000, 3),
                    **extra_meta,
                }
            })

        except Exception as e:
            traceback.print_exc()
            self._send_json(HTTPStatus.OK, {"error": str(e)})

    def _send_json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        pass  # アクセスログを抑制


# ===== 起動 =====

if __name__ == "__main__":
    url = f"http://{HOST}:{PORT}/"
    print(f"\n傾斜区分分析ツール v2 (北茨城市オフライン版)")
    print(f"URL: {url}")
    print(f"終了するにはこのウィンドウを閉じるか Ctrl+C を押してください\n")

    server = ThreadingHTTPServer((HOST, PORT), Handler)
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n終了しました")
        server.server_close()
