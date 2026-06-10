"""地形傾斜区分分析モジュール

国土地理院の標高タイルAPIからDEMデータを取得し、
傾斜角度の計算・区分分類・面積統計を行う。
"""

import math
import numpy as np
import requests

# ===== 定数 =====

TILE_ZOOM = 14
TILE_SIZE = 256

# 5mDEM → 10mDEM フォールバック
DEM_URLS = [
    "https://cyberjapandata.gsi.go.jp/xyz/dem5a/{z}/{x}/{y}.txt",
    "https://cyberjapandata.gsi.go.jp/xyz/dem/{z}/{x}/{y}.txt",
]

# 傾斜区分（角度下限, 角度上限, 名称, 表示色）
SLOPE_CLASSES = [
    (0,  5,  "平坦地",      "#2ecc71"),
    (5,  15, "緩傾斜地",    "#a8e063"),
    (15, 25, "中傾斜地",    "#f1c40f"),
    (25, 35, "急傾斜地(1)", "#e67e22"),
    (35, 45, "急傾斜地(2)", "#e74c3c"),
    (45, 90, "急峻地",      "#8e0000"),
]


def latlon_to_tile(lat: float, lon: float, zoom: int) -> tuple[int, int]:
    """緯度経度をタイル座標 (x, y) に変換する。"""
    n = 2 ** zoom
    x = int((lon + 180.0) / 360.0 * n)
    lat_rad = math.radians(lat)
    y = int((1.0 - math.log(math.tan(lat_rad) + 1.0 / math.cos(lat_rad)) / math.pi) / 2.0 * n)
    return x, y


def tile_to_latlon_bounds(tx: int, ty: int, zoom: int) -> tuple[float, float, float, float]:
    """タイル座標から (north, south, west, east) の緯度経度境界を返す。"""
    n = 2 ** zoom
    west = tx / n * 360.0 - 180.0
    east = (tx + 1) / n * 360.0 - 180.0
    north = math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * ty / n))))
    south = math.degrees(math.atan(math.sinh(math.pi * (1 - 2 * (ty + 1) / n))))
    return north, south, west, east


def fetch_dem_tile(tx: int, ty: int, zoom: int) -> np.ndarray:
    """GSI標高タイルAPIから1タイル(256×256)のDEM配列を取得する。

    5mDEM が存在しない場合は 10mDEM にフォールバックする。
    NoData('e') は NaN に変換する。
    """
    text = None
    for url_template in DEM_URLS:
        url = url_template.format(z=zoom, x=tx, y=ty)
        try:
            resp = requests.get(url, timeout=10)
            if resp.status_code == 200 and resp.text.strip():
                text = resp.text
                break
        except requests.RequestException:
            continue

    if text is None:
        return np.full((TILE_SIZE, TILE_SIZE), np.nan)

    rows = []
    for line in text.strip().split("\n"):
        values = []
        for v in line.split(","):
            v = v.strip()
            if v == "e" or v == "":
                values.append(np.nan)
            else:
                try:
                    values.append(float(v))
                except ValueError:
                    values.append(np.nan)
        while len(values) < TILE_SIZE:
            values.append(np.nan)
        rows.append(values[:TILE_SIZE])

    while len(rows) < TILE_SIZE:
        rows.append([np.nan] * TILE_SIZE)

    return np.array(rows[:TILE_SIZE], dtype=np.float64)


def merge_dem_tiles(
    lat_min: float, lon_min: float, lat_max: float, lon_max: float, zoom: int = TILE_ZOOM
) -> tuple[np.ndarray, tuple[float, float, float, float]]:
    """バウンディングボックスをカバーするDEMタイルを取得・結合する。

    Returns:
        dem: 結合後のDEM配列 (float64, NaN=NoData)
        bounds: (north, south, west, east) の実際のタイル境界
    """
    tx_min, ty_max = latlon_to_tile(lat_min, lon_min, zoom)
    tx_max, ty_min = latlon_to_tile(lat_max, lon_max, zoom)

    rows = []
    for ty in range(ty_min, ty_max + 1):
        row_tiles = []
        for tx in range(tx_min, tx_max + 1):
            tile = fetch_dem_tile(tx, ty, zoom)
            row_tiles.append(tile)
        rows.append(np.hstack(row_tiles))

    dem = np.vstack(rows)

    north, _, west, _ = tile_to_latlon_bounds(tx_min, ty_min, zoom)
    _, south, _, east = tile_to_latlon_bounds(tx_max, ty_max, zoom)

    return dem, (north, south, west, east)


# ===== 傾斜計算 =====

def calculate_slope(
    dem: np.ndarray, pixel_size_x: float, pixel_size_y: float
) -> np.ndarray:
    """DEMから傾斜角度(度)を計算する。

    Args:
        dem: 標高値の2Dアレイ (NaN=NoData)
        pixel_size_x: X方向ピクセルサイズ(メートル)
        pixel_size_y: Y方向ピクセルサイズ(メートル)

    Returns:
        slope: 傾斜角度(度)の2Dアレイ。NoDataセルは NaN。
    """
    dem_filled = _fill_nan(dem)
    dz_dy, dz_dx = np.gradient(dem_filled, pixel_size_y, pixel_size_x)
    slope_rad = np.arctan(np.sqrt(dz_dx**2 + dz_dy**2))
    slope_deg = np.degrees(slope_rad)
    slope_deg[np.isnan(dem)] = np.nan
    return slope_deg


def _fill_nan(arr: np.ndarray) -> np.ndarray:
    """NaN を平均値で置換する（勾配計算用）。"""
    filled = arr.copy()
    nan_mask = np.isnan(arr)
    if not nan_mask.any():
        return filled
    mean_val = np.nanmean(arr)
    if np.isnan(mean_val):
        mean_val = 0.0
    filled[nan_mask] = mean_val
    return filled


def classify_slope(slope_deg: np.ndarray) -> np.ndarray:
    """傾斜角度(度)を傾斜区分インデックス(0-5)に変換する。

    Returns:
        classified: 各ピクセルの区分インデックス (int8, NoData位置は -1)
    """
    result = np.full(slope_deg.shape, -1, dtype=np.int8)
    for i, (lo, hi, _, _) in enumerate(SLOPE_CLASSES):
        mask = (slope_deg >= lo) & (slope_deg < hi)
        result[mask] = i
    return result


def _pixel_sizes_at_lat(lat: float, zoom: int) -> tuple[float, float]:
    """指定緯度・ズームでの1ピクセルあたりのメートルサイズを返す。"""
    earth_circumference = 2 * math.pi * 6378137
    px = earth_circumference * math.cos(math.radians(lat)) / (TILE_SIZE * 2 ** zoom)
    py = earth_circumference / (TILE_SIZE * 2 ** zoom)
    return px, py


def compute_polygon_stats(
    dem: np.ndarray,
    bounds: tuple[float, float, float, float],
    polygon_coords: list,
) -> list:
    """ポリゴン内の傾斜区分別面積・割合を計算する。

    Args:
        dem: DEM配列 (float64)
        bounds: (north, south, west, east) のタイル境界
        polygon_coords: [[lon, lat], ...] のポリゴン座標列 (GeoJSON形式)

    Returns:
        stats: 各区分の {"name", "area_m2", "percent", "color"} リスト (6要素)
    """
    from matplotlib.path import Path

    north, south, west, east = bounds
    h, w = dem.shape

    mid_lat = (north + south) / 2
    px, py = _pixel_sizes_at_lat(mid_lat, TILE_ZOOM)

    slope_deg = calculate_slope(dem, pixel_size_x=px, pixel_size_y=py)
    classified = classify_slope(slope_deg)

    lats = np.linspace(north, south, h)
    lons = np.linspace(west, east, w)
    lon_grid, lat_grid = np.meshgrid(lons, lats)
    points = np.column_stack([lon_grid.ravel(), lat_grid.ravel()])

    path = Path([(c[0], c[1]) for c in polygon_coords])
    mask = path.contains_points(points).reshape(h, w)

    valid_mask = mask & (classified >= 0)
    total_pixels = valid_mask.sum()
    pixel_area = px * py

    stats = []
    for i, (_, _, name, color) in enumerate(SLOPE_CLASSES):
        count = int(((classified == i) & valid_mask).sum())
        area = count * pixel_area
        percent = (count / total_pixels * 100) if total_pixels > 0 else 0.0
        stats.append({
            "name": name,
            "area_m2": round(area, 1),
            "percent": round(float(percent), 2),
            "color": color,
        })

    return stats
