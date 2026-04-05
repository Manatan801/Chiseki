# XserverVPS デプロイ手順

`https://gomi-maru.com/chiseki` で公開するまでの手順。

## 前提

- XserverVPS（Ubuntu 22.04 想定）にSSH接続できる
- ドメイン `gomi-maru.com` のDNS Aレコードが VPS のIPを指している
- rootまたはsudo権限がある

---

## 1. 必要パッケージのインストール

```bash
sudo apt update
sudo apt install -y python3 python3-venv python3-pip nginx git certbot python3-certbot-nginx
```

## 2. アプリを配置

```bash
sudo mkdir -p /opt/chiseki
sudo chown $USER:$USER /opt/chiseki
git clone https://github.com/yourname/Chiseki.git /opt/chiseki
cd /opt/chiseki

python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

## 3. 動作確認（手動起動）

```bash
cd /opt/chiseki
source venv/bin/activate
streamlit run app.py --server.port=8501 --server.address=127.0.0.1 --server.baseUrlPath=/chiseki
```

別ターミナルで `curl http://127.0.0.1:8501/chiseki/` が200を返せばOK。
Ctrl+C で停止。

## 4. systemd でサービス化

```bash
sudo chown -R www-data:www-data /opt/chiseki
sudo cp /opt/chiseki/deploy/chiseki.service /etc/systemd/system/chiseki.service
sudo systemctl daemon-reload
sudo systemctl enable chiseki
sudo systemctl start chiseki
sudo systemctl status chiseki
```

**パスワードを変更したい場合:**
```bash
sudo systemctl edit chiseki
# エディタで以下を追記
[Service]
Environment="CHISEKI_USER=新しいID"
Environment="CHISEKI_PASSWORD=新しいパスワード"

sudo systemctl restart chiseki
```

## 5. Nginx でリバースプロキシ設定

### 既存の gomi-maru.com 設定がある場合

`/etc/nginx/sites-available/gomi-maru.com` のHTTPS serverブロック内に、
`deploy/nginx.conf` の `location /chiseki/` と `location /chiseki/_stcore/` のブロック
だけをコピーして追加してください。

### 新規で作る場合

```bash
sudo cp /opt/chiseki/deploy/nginx.conf /etc/nginx/sites-available/gomi-maru.com
sudo ln -s /etc/nginx/sites-available/gomi-maru.com /etc/nginx/sites-enabled/
sudo nginx -t
sudo systemctl reload nginx
```

## 6. SSL証明書（Let's Encrypt）

```bash
sudo certbot --nginx -d gomi-maru.com -d www.gomi-maru.com
```

質問に答えて証明書を取得。完了するとNginx設定が自動更新されHTTPS化される。

## 7. 動作確認

ブラウザで https://gomi-maru.com/chiseki にアクセス。
ログイン画面が出たら成功。
- ID: `kitaibachiseki`
- PW: `tankachou`

---

## 運用コマンド

```bash
# ログ確認
sudo journalctl -u chiseki -f

# 再起動
sudo systemctl restart chiseki

# コード更新後の反映
cd /opt/chiseki
sudo -u www-data git pull
sudo -u www-data venv/bin/pip install -r requirements.txt
sudo systemctl restart chiseki

# Nginx設定変更後
sudo nginx -t && sudo systemctl reload nginx
```

## トラブルシューティング

**404 Not Found**
- Streamlit起動時に `--server.baseUrlPath=/chiseki` がついているか確認
- Nginxの `location /chiseki/` のスラッシュ末尾が一致しているか確認

**WebSocket接続エラー（画面が真っ白）**
- Nginxで `proxy_set_header Upgrade` と `Connection "upgrade"` が設定されているか確認

**ファイルアップロードで413エラー**
- Nginxの `client_max_body_size` を大きくする

**起動失敗**
- `sudo journalctl -u chiseki -n 50` でエラー確認
- `/opt/chiseki` のパーミッションが `www-data` になっているか確認
