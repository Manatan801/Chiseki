# 建物外形データの追加方法

このツールは、任意データとして `data/buildings.geojson` を読み込みます。
現在は、`../GML` に置いた国土地理院 基盤地図情報 GML から `BldA`
（建築物の外周線）を変換した実データを収録しています。

## 対応形式

- GeoJSON `FeatureCollection`
- 座標系: WGS84 経度緯度 (`[lon, lat]`)
- ジオメトリ: `Polygon` または `MultiPolygon`

## 用途

- 地図上の「建物レイヤー」表示
- 選択範囲内の建物面積、建物棟数、建物割合の集計

## 変換方法

```bash
python3 convert_buildings_gml.py
```

- 既定入力: `../GML/FG-GML-*-ALL-*.zip`
- 既定出力: `data/buildings.geojson`
- 対象地物: `BldA`
- 座標変換: GML の `lat lon` を GeoJSON の `[lon, lat]` に変換

## 収録状況

- 対象メッシュ: `554015`, `554016`, `554025`, `554026`
- 建物ポリゴン数: `58,034`
- bbox: `140.627776056,36.75,140.84683663,36.916666667`
- 西部山間部の未取得メッシュは対象外です。

## 注意

- 現行の傾斜区分と標高統計は `data/kitaibaraki_dem.npz` のDEMから計算します。
- 建物面積はDSM推定ではなく、建物外形ポリゴンとの交差面積で計算します。
- `data/buildings.geojson` がない場合も、従来どおりDEM傾斜分析は利用できます。
