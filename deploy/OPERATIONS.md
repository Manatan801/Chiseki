# Chiseki Webアプリ 運用ドキュメント

`https://gomi-maru.com/chiseki` として公開している地籍調査DX帳票自動生成ツールの運用マニュアル。

---

## 1. 公開情報

| 項目 | 値 |
|---|---|
| URL | `https://gomi-maru.com/chiseki` |
| ログインID | `kitaibachiseki` |
| ログインPW | `tankachou` |
| サーバー | XserverVPS (Ubuntu 22.04) |
| 配置ディレクトリ | `/opt/chiseki` |
| サービス名 | `chiseki.service` (systemd) |
| 内部ポート | `127.0.0.1:8501` |
| Nginx設定 | `/etc/nginx/sites-enabled/gomi-maru` の `/chiseki/` location |
| SSL証明書 | Let's Encrypt（certbot自動更新） |

---

## 2. 使用上の条件（DXFファイル要件）

このツールは **JW-CADから出力された特定形式のDXF** のみ処理できます。

### 必須要素

| 要素 | DXFエンティティ | 条件 |
|---|---|---|
| 杭（境界点） | **CIRCLE** | 半径 = 0.25 |
| 杭番号 | **MTEXT** | 通常杭: `1.830`形式 / 交点杭: `-1.824`形式 |
| 境界線 | **LWPOLYLINE** | 頂点数 = 2（2点の線分） |
| 地番 | **MTEXT** | `1070-6` 形式（整数 or 整数-整数） |
| 公共用地 | **MTEXT** | `県道-5`, `市道-19`, `水-23` 等 |
| 地目 | **MTEXT** | 宅地、山林、雑種地、公衆用道路、田、畑 等 |

### 処理できないDXFの例

- **LINE / TEXT エンティティを使ったDXF**（生のJW-CADエクスポート等）
- 杭の半径が0.25以外
- 杭番号のフォーマットが上記と異なる

診断画面の「DXF内のエンティティ種別」で、`LWPOLYLINE`と`MTEXT`の数を確認できます。これらが0の場合は対応外の形式です。

### テキスト配置のルール

- 杭番号テキストは対応するCIRCLE（r=0.25）の近くに配置する
- 地番テキストは該当区画の**内部**に配置する
- 狭い区画で引き出し線を使った場合、「地番不明」シートに分類されます

---

## 3. 出力仕様

### 結線指示票（`結線指示票.xls`）

- 各筆ごとに1シート
- 北西の角から**時計回り**に杭番号を並べる
- 100点を超える区画は自動的に2ページに分割
- 「**地番不明**」シートは手動で地番を補記する必要あり
- 公共用地（道路・河川等）もラベルがあればシート生成

### 交点計算指示書（`交点計算指示書.xlsx`）

- 同一直線ペア（基準線）が1つ + 方向点が1つの明確なケースのみ
- 複数の直線が交わる複雑な交点は人手対応が必要
- 3列×5段のブロック配置

---

## 4. 運用コマンド集

すべてVPSにSSH接続してから実行します。

### コード更新

GitHubにpushした変更を本番へ反映：

```bash
cd /opt/chiseki
sudo -u www-data git pull
sudo -u www-data venv/bin/pip install -r requirements.txt  # 依存追加時のみ
sudo systemctl restart chiseki
```

### パスワード変更

```bash
sudo systemctl edit chiseki
```

エディタが開くので以下を追記：

```ini
[Service]
Environment="CHISEKI_USER=新しいID"
Environment="CHISEKI_PASSWORD=新しいパスワード"
```

保存後：
```bash
sudo systemctl restart chiseki
```

### サービス管理

```bash
# 状態確認
sudo systemctl status chiseki

# 再起動
sudo systemctl restart chiseki

# 停止
sudo systemctl stop chiseki

# 起動
sudo systemctl start chiseki

# 自動起動ON/OFF
sudo systemctl enable chiseki
sudo systemctl disable chiseki
```

### ログ確認

```bash
# リアルタイム表示
sudo journalctl -u chiseki -f

# 直近50行
sudo journalctl -u chiseki -n 50

# 今日のログ
sudo journalctl -u chiseki --since today

# エラーのみ
sudo journalctl -u chiseki -p err
```

### Nginx管理

```bash
# 設定ファイル
sudo nano /etc/nginx/sites-enabled/gomi-maru

# 設定テスト
sudo nginx -t

# 設定反映（リロード、ダウンタイムなし）
sudo systemctl reload nginx

# 完全再起動
sudo systemctl restart nginx
```

### SSL証明書

Let's EncryptはCertbotが自動更新します。手動で確認する場合：

```bash
# 有効期限確認
sudo certbot certificates

# 更新テスト（ドライラン）
sudo certbot renew --dry-run

# 強制更新
sudo certbot renew --force-renewal
```

---

## 5. トラブルシューティング

### アプリが起動しない

```bash
sudo systemctl status chiseki
sudo journalctl -u chiseki -n 50
```

よくある原因：
- Python依存ライブラリの不足 → `sudo -u www-data /opt/chiseki/venv/bin/pip install -r /opt/chiseki/requirements.txt`
- ファイル所有権の問題 → `sudo chown -R www-data:www-data /opt/chiseki`
- ポート衝突 → `sudo lsof -i:8501` で8501を使っているプロセスを確認

### ブラウザで画面が真っ白

Streamlitは**WebSocket**を使うため、Nginxの以下の設定が必須：

```nginx
proxy_set_header Upgrade $http_upgrade;
proxy_set_header Connection "upgrade";
```

`/etc/nginx/sites-enabled/gomi-maru` の `/chiseki/` locationに両方が含まれているか確認。

### ファイルアップロードで413エラー

Nginxの`client_max_body_size`を大きくする：

```nginx
client_max_body_size 200M;  # 現在は100M
```

### DXF処理でエラー

診断情報の「DXF内のエンティティ種別」を確認：
- `MTEXT`が0 → LINE/TEXT形式のDXF（本ツール非対応）
- `CIRCLE`が0 → 杭が半径0.25で描かれていない
- `LWPOLYLINE`が0 → 境界線の形式が違う

いずれも**DXFファイル側の修正が必要**です。

### 504 Gateway Timeout

大きなDXFの処理に時間がかかっている可能性。Nginxのタイムアウトを延長：

```nginx
proxy_read_timeout 600s;   # 現在300s
proxy_send_timeout 600s;
```

---

## 6. ファイル配置

```
/opt/chiseki/                      ... アプリルート
├── app.py                         ... Streamlit WebUI
├── requirements.txt
├── src/                           ... DXF解析ロジック
├── data/templates/                ... Excelテンプレート
├── deploy/
│   ├── chiseki.service            ... systemd設定ソース
│   ├── nginx.conf                 ... Nginx設定サンプル
│   ├── README.md                  ... デプロイ手順書
│   └── OPERATIONS.md              ... 本ドキュメント
└── venv/                          ... Python仮想環境

/etc/systemd/system/chiseki.service  ... systemd実稼働ファイル
/etc/nginx/sites-enabled/gomi-maru   ... Nginx実稼働設定
/etc/letsencrypt/live/gomi-maru.com/ ... SSL証明書
```

---

## 7. セキュリティ

### 認証情報

- ログイン情報は**環境変数で管理**（systemdの`Environment=`）
- コードにハードコードされた値はデフォルトで、環境変数があれば上書き
- GitHubリポジトリに認証情報を含めない

### アップロードファイル

- 処理後の一時ファイルは即削除
- DXFファイル本体はサーバーに保存しない
- 生成されたExcelは**セッションメモリ上**のみに保持（再起動で消える）

### 推奨運用

- パスワードは定期変更（3〜6ヶ月ごと）
- OSのセキュリティアップデート: `sudo apt update && sudo apt upgrade`
- バックアップ: `/opt/chiseki` はGitで管理されているのでclone可能

---

## 8. 監視・メトリクス

現在は**監視設定なし**。必要になったら以下を検討：

- `systemd`のWatchdog
- Uptime監視（UptimeRobot等の外部サービス）
- ディスク容量アラート

---

## 9. 問い合わせ・更新

- ソースコード: https://github.com/Manatan801/Chiseki
- 改修はローカルで開発 → GitHub push → VPSで `git pull` & `systemctl restart` の流れ
