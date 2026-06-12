"""
北茨城市 DEMタイル・背景地図タイル ダウンロードスクリプト

- DEM: 国土地理院 dem5a_png z=15 (約3.8m解像度) 1755枚
- 背景地図: 国土地理院 淡色地図 z=10〜14 631枚
- 並列数: 3 (GSIサーバー負荷軽減)
- 出力: /mnt/f/傾斜区分分析_北茨城市/data/
"""

import io
import math
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import numpy as np
import requests
from PIL import Image

# ===== 設定 =====
OUT_DIR = Path(__file__).parent.parent / "傾斜区分分析_北茨城市" / "data"
DEM_NPZ = OUT_DIR / "kitaibaraki_dem.npz"
TILES_DIR = OUT_DIR / "map_tiles"

# 北茨城市バウンディングボックス（余裕を持たせる）
LAT_MIN, LAT_MAX = 36.68, 37.02
LON_MIN, LON_MAX = 140.43, 140.92

DEM_ZOOM = 15
MAP_ZOOMS = list(range(10, 15))
WORKERS = 3
INTERVAL = 0.1  # スレッド間リクエスト間隔(秒)

DEM_URL = "https://cyberjapandata.gsi.go.jp/xyz/dem5a_png/{z}/{x}/{y}.png"
MAP_URL = "https://cyberjapandata.gsi.go.jp/xyz/pale/{z}/{x}/{y}.png"


# ===== ユーティリティ =====

def latlon_to_tile(lat, lon, zoom):
    n = 2 ** zoom
    x = int((lon + 180.0) / 360.0 * n)
    lat_rad = math.radians(lat)
    y = int((1.0 - math.log(math.tan(lat_rad) + 1.0 / math.cos(lat_rad)) / math.pi) / 2.0 * n)
    return x, y


def tile_to_latlon_bounds(tx, ty, zoom):
    n = 2 ** zoom
    west  = tx / n * 360.0 - 180.0
    east  = (tx + 1) / n * 360.0 - 180.0
    north = math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * ty / n))))
    south = math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * (ty + 1) / n))))
    return north, south, west, east


def decode_dem_png(content: bytes) -> np.ndarray:
    """GSI PNG標高タイルを float64 標高配列(m)に変換。NoData=NaN"""
    img = np.array(Image.open(io.BytesIO(content)).convert("RGB"), dtype=np.float64)
    R, G, B = img[:, :, 0], img[:, :, 1], img[:, :, 2]
    u = R * 65536 + G * 256 + B
    nodata = (R == 128) & (G == 0) & (B == 0)
    elev = np.where(u >= 2 ** 23, (u - 2 ** 24) * 0.01, u * 0.01)
    elev[nodata] = np.nan
    return elev


def fetch_with_retry(session, url, retries=3):
    for attempt in range(retries):
        try:
            r = session.get(url, timeout=20)
            if r.status_code == 200:
                return r.content
            time.sleep(1.0 * (attempt + 1))
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(2.0 * (attempt + 1))
    return None


# ===== DEM ダウンロード・統合 =====

def download_dem():
    tx_min, ty_max = latlon_to_tile(LAT_MIN, LON_MIN, DEM_ZOOM)
    tx_max, ty_min = latlon_to_tile(LAT_MAX, LON_MAX, DEM_ZOOM)
    nx = tx_max - tx_min + 1
    ny = ty_max - ty_min + 1
    total = nx * ny

    print(f"\n[DEM] z={DEM_ZOOM}: X={tx_min}〜{tx_max}({nx}枚) Y={ty_min}〜{ty_max}({ny}枚) 合計{total}枚")
    print(f"      並列数={WORKERS}, 推定所要時間={total/WORKERS*0.25/60:.0f}〜{total/WORKERS*0.4/60:.0f}分")

    # タイル座標リスト（行順）
    tile_coords = [(tx, ty) for ty in range(ty_min, ty_max + 1) for tx in range(tx_min, tx_max + 1)]

    results = {}  # (tx,ty) → np.ndarray
    completed = 0
    failed = 0

    session = requests.Session()
    session.headers["User-Agent"] = "ChisekiTerrainTool/1.0"

    def fetch_dem_tile(coord):
        tx, ty = coord
        url = DEM_URL.format(z=DEM_ZOOM, x=tx, y=ty)
        time.sleep(INTERVAL)
        content = fetch_with_retry(session, url)
        if content:
            return coord, decode_dem_png(content)
        return coord, np.full((256, 256), np.nan)

    with ThreadPoolExecutor(max_workers=WORKERS) as exe:
        futures = {exe.submit(fetch_dem_tile, c): c for c in tile_coords}
        for future in as_completed(futures):
            coord, arr = future.result()
            results[coord] = arr
            completed += 1
            if completed % 50 == 0 or completed == total:
                pct = completed / total * 100
                print(f"  DEM進捗: {completed}/{total} ({pct:.0f}%)", flush=True)

    # グリッドに並べる（ty行, tx列）
    print("  タイル結合中...", flush=True)
    rows = []
    for ty in range(ty_min, ty_max + 1):
        row = [results[(tx, ty)] for tx in range(tx_min, tx_max + 1)]
        rows.append(np.hstack(row))
    dem = np.vstack(rows)

    # 境界
    north, _, west, _ = tile_to_latlon_bounds(tx_min, ty_min, DEM_ZOOM)
    _, south, _, east = tile_to_latlon_bounds(tx_max, ty_max, DEM_ZOOM)

    # int16 (×10, 0.1m精度) で保存。NoData=-32768
    int16 = np.where(np.isnan(dem), -32768,
                     np.clip(np.round(dem * 10), -32767, 32767)).astype(np.int16)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    np.savez_compressed(
        str(DEM_NPZ),
        dem=int16,
        bounds=np.array([north, south, west, east])
    )
    size_mb = DEM_NPZ.stat().st_size / 1024 / 1024
    valid_pct = np.sum(int16 != -32768) / int16.size * 100
    print(f"  [完了] DEM保存: {DEM_NPZ.name} ({size_mb:.0f}MB)")
    print(f"         shape={dem.shape}, 有効ピクセル={valid_pct:.1f}%")
    print(f"         bounds: N={north:.4f} S={south:.4f} W={west:.4f} E={east:.4f}")
    session.close()


# ===== 背景地図タイル ダウンロード =====

def download_map_tiles():
    all_coords = []
    for zoom in MAP_ZOOMS:
        tx_min, ty_max = latlon_to_tile(LAT_MIN, LON_MIN, zoom)
        tx_max, ty_min = latlon_to_tile(LAT_MAX, LON_MAX, zoom)
        for ty in range(ty_min, ty_max + 1):
            for tx in range(tx_min, tx_max + 1):
                all_coords.append((zoom, tx, ty))

    total = len(all_coords)
    print(f"\n[背景地図] z={MAP_ZOOMS[0]}〜{MAP_ZOOMS[-1]}: 合計{total}枚")

    session = requests.Session()
    session.headers["User-Agent"] = "ChisekiTerrainTool/1.0"
    completed = 0

    def fetch_map_tile(coord):
        z, tx, ty = coord
        url = MAP_URL.format(z=z, x=tx, y=ty)
        out = TILES_DIR / str(z) / str(tx) / f"{ty}.png"
        if out.exists():
            return True  # キャッシュ済み
        time.sleep(INTERVAL)
        content = fetch_with_retry(session, url)
        if content:
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_bytes(content)
            return True
        return False

    with ThreadPoolExecutor(max_workers=WORKERS) as exe:
        futures = {exe.submit(fetch_map_tile, c): c for c in all_coords}
        for future in as_completed(futures):
            future.result()
            completed += 1
            if completed % 100 == 0 or completed == total:
                print(f"  背景地図進捗: {completed}/{total} ({completed/total*100:.0f}%)", flush=True)

    session.close()
    tile_count = sum(1 for _ in TILES_DIR.rglob("*.png"))
    print(f"  [完了] 背景地図: {tile_count}枚保存")


# ===== Leaflet.js アセットのダウンロード =====

def download_leaflet_assets():
    """Leaflet.js と Leaflet.draw を静的ファイルとしてダウンロード"""
    assets_dir = OUT_DIR.parent / "assets"
    assets_dir.mkdir(exist_ok=True)

    assets = [
        ("https://unpkg.com/leaflet@1.9.4/dist/leaflet.js",
         assets_dir / "leaflet.js"),
        ("https://unpkg.com/leaflet@1.9.4/dist/leaflet.css",
         assets_dir / "leaflet.css"),
        ("https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.js",
         assets_dir / "leaflet.draw.js"),
        ("https://cdnjs.cloudflare.com/ajax/libs/leaflet.draw/1.0.4/leaflet.draw.css",
         assets_dir / "leaflet.draw.css"),
    ]

    print("\n[Leaflet] JSライブラリをダウンロード中...")
    session = requests.Session()
    for url, out_path in assets:
        if out_path.exists():
            print(f"  スキップ(既存): {out_path.name}")
            continue
        content = fetch_with_retry(session, url)
        if content:
            out_path.write_bytes(content)
            print(f"  ダウンロード完了: {out_path.name} ({len(content)//1024}KB)")
        else:
            print(f"  失敗: {url}", file=sys.stderr)
    session.close()
    print("  [完了] Leafletアセット")


# ===== メイン =====

if __name__ == "__main__":
    print("=" * 60)
    print("北茨城市 地形データ ダウンロード開始")
    print("=" * 60)
    t_start = time.time()

    download_leaflet_assets()

    if DEM_NPZ.exists():
        print(f"\n[DEM] 既存ファイルを検出: {DEM_NPZ} → スキップ")
        print("  再ダウンロードする場合は削除してから再実行してください")
    else:
        download_dem()

    download_map_tiles()

    elapsed = time.time() - t_start
    print(f"\n{'='*60}")
    print(f"完了! 総所要時間: {elapsed/60:.1f}分")
    print(f"出力先: {OUT_DIR.parent}")
