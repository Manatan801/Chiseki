import math
import numpy as np
import pytest
from src.terrain_analysis import (
    latlon_to_tile,
    tile_to_latlon_bounds,
    fetch_dem_tile,
    merge_dem_tiles,
    TILE_ZOOM,
    TILE_SIZE,
    calculate_slope,
    classify_slope,
    compute_polygon_stats,
    SLOPE_CLASSES,
)


# ===== DEM取得関数テスト =====

def test_latlon_to_tile_tokyo():
    """東京駅付近 (35.68, 139.77) のタイル座標確認"""
    x, y = latlon_to_tile(35.68, 139.77, zoom=14)
    assert x == 14553
    assert y == 6451


def test_tile_to_latlon_bounds():
    """タイル境界を lat/lon に変換できること"""
    bounds = tile_to_latlon_bounds(14553, 6451, zoom=14)
    north, south, west, east = bounds
    assert north > south
    assert east > west
    assert 35 < north < 36
    assert 139 < east < 141


def test_fetch_dem_tile_returns_256x256():
    """DEMタイルが 256×256 の numpy 配列を返すこと"""
    arr = fetch_dem_tile(14553, 6451, zoom=14)
    assert arr.shape == (TILE_SIZE, TILE_SIZE)
    assert arr.dtype == np.float64


def test_merge_dem_tiles_single():
    """単一タイルのマージが配列を返すこと"""
    dem, bounds = merge_dem_tiles(35.68, 139.77, 35.69, 139.78)
    assert dem.ndim == 2
    assert dem.shape[0] > 0
    assert dem.shape[1] > 0


def test_merge_dem_tiles_no_data_filled():
    """NoData('e')が NaN に変換されていること"""
    arr = fetch_dem_tile(14553, 6451, zoom=14)
    valid = arr[~np.isnan(arr)]
    assert np.all(np.isfinite(valid))


# ===== 傾斜計算テスト =====

def test_calculate_slope_flat():
    """平坦なDEMから傾斜0度が得られること"""
    dem = np.ones((10, 10)) * 100.0
    slope = calculate_slope(dem, pixel_size_x=5.0, pixel_size_y=5.0)
    assert slope.shape == dem.shape
    assert np.nanmax(slope) < 0.1


def test_calculate_slope_known_angle():
    """1m/5m の傾斜 ≈ 11.3° が得られること"""
    dem = np.zeros((10, 10))
    for i in range(10):
        dem[i, :] = i * 1.0
    slope = calculate_slope(dem, pixel_size_x=5.0, pixel_size_y=5.0)
    center_slope = slope[5, 5]
    expected = math.degrees(math.atan(1.0 / 5.0))
    assert abs(center_slope - expected) < 1.0


def test_classify_slope_all_categories():
    """全傾斜区分への分類が正しいこと"""
    angles = np.array([2.0, 10.0, 20.0, 30.0, 40.0, 50.0])
    classified = classify_slope(angles)
    assert list(classified) == [0, 1, 2, 3, 4, 5]


def test_classify_slope_boundary():
    """境界値 5°, 15° が正しい区分に入ること"""
    angles = np.array([5.0, 15.0, 25.0, 35.0, 45.0])
    classified = classify_slope(angles)
    assert classified[0] == 1
    assert classified[1] == 2
    assert classified[2] == 3
    assert classified[3] == 4
    assert classified[4] == 5


def test_compute_polygon_stats_full_area():
    """ポリゴンが全ピクセルを含む場合、合計が100%になること"""
    dem = np.ones((10, 10)) * 50.0
    bounds = (36.0, 35.9, 139.0, 139.1)
    polygon_coords = [
        [139.0, 36.0], [139.1, 36.0], [139.1, 35.9], [139.0, 35.9], [139.0, 36.0]
    ]
    stats = compute_polygon_stats(dem, bounds, polygon_coords)
    total_pct = sum(s["percent"] for s in stats)
    assert abs(total_pct - 100.0) < 0.1


def test_compute_polygon_stats_returns_all_classes():
    """統計結果が全6区分を返すこと"""
    dem = np.ones((10, 10)) * 50.0
    bounds = (36.0, 35.9, 139.0, 139.1)
    polygon_coords = [
        [139.0, 36.0], [139.1, 36.0], [139.1, 35.9], [139.0, 35.9], [139.0, 36.0]
    ]
    stats = compute_polygon_stats(dem, bounds, polygon_coords)
    assert len(stats) == len(SLOPE_CLASSES)
    for s in stats:
        assert "name" in s
        assert "area_m2" in s
        assert "percent" in s
        assert "color" in s
