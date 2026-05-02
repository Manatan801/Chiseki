# ACCESS差分ツール

2つのAccess MDBファイルを比較し、差分をExcelファイルに出力するローカルツールです。

## Windowsへ移すファイル

`ACCESS差分ツール_Windows配布` フォルダごとWindowsへコピーしてください。

```text
ACCESS差分ツール_Windows配布/
  ACCESS差分ツール起動.bat
  初回セットアップ.bat
  ACCESS差分ツール_README.md
  access_diff_web.py
  requirements.txt
  wheels/
```

## 必要なもの

- Python 3.12 64bit
- Microsoft Access Database Engine 64bit

Access Database Engine が入っていないPCでは、MDBを開けません。Microsoft公式の「Access Database Engine Redistributable」をインストールしてください。

## 初回だけ行うこと

`初回セットアップ.bat` をダブルクリックしてください。

`wheels` フォルダに入っているWindows 64bit + Python 3.12用ライブラリから、オフラインでインストールします。失敗した場合はインターネット経由でインストールを試します。

## 起動方法

`ACCESS差分ツール起動.bat` をダブルクリックしてください。

ブラウザで次の画面が開きます。

```text
http://127.0.0.1:8767/
```

## 使い方

1. 基準MDBに `DT2.MDB` など最新側のファイルを選択します。
2. 比較MDBにチェックリスト・マスター側のMDBを選択します。
3. 比較対象テーブルを確認します。既定は `L31,L33,L34` です。
4. 必要なら無視する列を指定します。既定は `ID` です。
5. `差分Excelを作成` を押します。
6. 差分Excelがダウンロードされます。

## 出力内容

- `概要`: テーブルごとの件数、追加、削除、変更、欠落状況
- `差分`: 列単位で異なる値
- `基準のみ`: 基準MDBにだけある行
- `比較のみ`: 比較MDBにだけある行
- `テーブル一覧`: 両MDBにあるテーブル一覧

行の照合キーは、共通列から自動で判定します。`AzaKY, FDKBN, CHIBN, SEQ, NUMB` などの土地・所有者系キーを優先します。

## 注意

このツールはMDBを修正しません。出力Excelを確認し、人がAccess側を修正してください。
