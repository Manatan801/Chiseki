# BM CSV差分ツール

## Windowsへ移すファイル

配布フォルダ内の次の4ファイルを、同じフォルダのままWindowsへコピーしてください。

- `diff_bm_csv.py`
- `diff_bm_csv_web.py`
- `BM_CSV差分ツール起動.bat`
- `BM_CSV差分ツール_README.md`

## 必要なもの

- Python 3.10以上

追加ライブラリは不要です。Python標準ライブラリだけでローカルのブラウザ画面を起動します。

## Windowsでの起動

`BM_CSV差分ツール起動.bat` をダブルクリックしてください。

起動しない場合は、コマンドプロンプトで3ファイルを置いたフォルダに移動し、次のどちらかを実行してください。

```bat
python3 diff_bm_csv_web.py
```

```bat
py -3 diff_bm_csv_web.py
```

```bat
python diff_bm_csv_web.py
```

## Linux / WSLでの起動

ターミナルで次を実行してください。

```bash
cd /home/maaatan/Chiseki/youyakusyo
./BM_CSV差分ツール起動.sh
```

または直接起動します。

```bash
python3 diff_bm_csv_web.py
```

## 使い方

1. `旧CSV` を選択します。
2. `新CSV` を選択します。
3. `差分CSVを作成` を押します。
4. 差分CSVがブラウザからダウンロードされます。

出力CSVはExcelで開きやすいUTF-8 BOM付きCSVです。

## コマンドで使う場合

```bat
py -3 diff_bm_csv.py old.csv new.csv -o diff.csv
```
