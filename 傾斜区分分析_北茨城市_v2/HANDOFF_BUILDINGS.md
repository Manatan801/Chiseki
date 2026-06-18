# 建物外形データ追加 引き継ぎメモ

更新日: 2026-06-18

## 現在の状態

- `terrain_web.py` は `data/buildings.geojson` を任意データとして読み込める。
- 建物外形データがある場合は、地図の「建物レイヤー」で表示できる。
- 解析時は、選択範囲と建物外形ポリゴンの交差面積から以下を出す。
  - 建物面積
  - 建物棟数
  - 選択範囲に占める建物割合
- `data/buildings.geojson` が無い場合も、従来どおりDEM傾斜分析は動く。
- 実データとして `58,034` polygon を収録済み。
- 収録メッシュは `554015`, `554016`, `554025`, `554026`。西部山間部の未取得メッシュは対象外。

## 再変換するとき

```bash
python3 convert_buildings_gml.py
```

- 既定入力: `../GML/FG-GML-*-ALL-*.zip`
- 既定出力: `data/buildings.geojson`
- GML の `lat lon` は GeoJSON の `[lon, lat]` に変換済み。

## 当初の作業メモ

1. 国土地理院の基盤地図情報ダウンロードサービスにログインする。
   - https://service.gsi.go.jp/kiban/
2. 「基本項目」を選ぶ。
3. 北茨城市周辺を対象にする。
   - 市区町村指定または2次メッシュ指定で取得する。
4. 地物等で「建築物の外周線」を含むデータをダウンロードする。
5. ダウンロードしたZIP/GMLをGeoJSONへ変換する。
6. 変換結果を次の場所に置く。
   - `傾斜区分分析_北茨城市_v2/data/buildings.geojson`
7. アプリを起動して、建物レイヤー表示と建物面積集計を確認する。

## 変換先GeoJSONの要件

- `FeatureCollection`
- 座標系はWGS84経度緯度。
- 座標順は `[lon, lat]`。
- ジオメトリは `Polygon` または `MultiPolygon`。

## 参考情報

- 国土地理院の公式ヘルプでは、基盤地図情報「基本項目」に「建築物の外周線」が含まれる。
- ダウンロードファイル形式はJPGIS/GML。
- 2014年7月31日以降の基本項目は2次メッシュ単位で全国提供と説明されている。
- 都市計画区域は縮尺1/2,500相当、都市計画区域外は縮尺1/25,000相当なので、山間部・郊外では精度差が出る可能性がある。

## 実装済みファイル

- `terrain_web.py`
  - `data/buildings.geojson` 読み込み
  - `/buildings?bbox=west,south,east,north`
  - 建物レイヤー表示
  - 建物面積集計
  - CSVへの建物面積出力
- `convert_buildings_gml.py`
  - `../GML` の GML ZIP から `BldA` を GeoJSON へ変換
- `BUILDINGS_DATA.md`
  - 建物GeoJSONの配置方法と形式説明
- `PROGRESS.md`
  - 実装内容と検証内容を追記済み

## 検証済み

- `python3 -m py_compile 傾斜区分分析_北茨城市_v2/terrain_web.py`
- 実DEMの小ポリゴンで傾斜解析が正常に動くこと。
- 建物データ未収録時に解析APIが正常応答すること。
- 仮の建物ポリゴンで建物面積集計が動くこと。
- HTTPで `/` と `/buildings` が200応答すること。
- 実 GML 変換後、`terrain_web.py` が `58,034` polygon を読み込むこと。
- `/buildings` 相当の bbox フィルタで小範囲の建物 GeoJSON が返ること。

## 注意

- 国土地理院サービスはログインが必要なため、完全自動取得は未実施。
- 追加メッシュを取得した場合は、`../GML` に ZIP を追加して再変換する。
