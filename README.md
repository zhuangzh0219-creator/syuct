# Nagi POS（iPad向けデモ）

Django製のタブレット向けPOS試作アプリです。商品・顧客管理、カート会計、PayPay / Alipay / 現金 / クレジットカード（デモ）の選択、領収書、売上レポートを提供します。

> **決済に関する重要事項:** これは画面と販売記録の動作確認用です。どの支払方法を選んでも決済事業者には接続せず、実際のお金は動きません。実運用には各社の加盟店契約、公式API/SDK、署名検証、返金・取消・照合処理、セキュリティレビューが必要です。カード番号等の機密情報は扱いません。

## 起動方法（Windows / PowerShell）

Python 3.10以上をインストールしてください。

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python manage.py migrate
python manage.py seed_demo
python manage.py runserver 0.0.0.0:8000
```

PCでは `http://127.0.0.1:8000/` を開きます。iPadを同じWi-Fiに接続し、PCのローカルIPアドレスを使って `http://<PCのIPアドレス>:8000/` を開いてください。Windows Defender ファイアウォールで必要な場合はプライベートネットワーク上のポート8000を許可します。

iPadからアクセスする場合は、起動前にPCのローカルIPアドレスを許可ホストへ追加します（例のIPは実際の値に置き換えてください）。

```powershell
$env:DJANGO_ALLOWED_HOSTS = "127.0.0.1,localhost,192.168.1.20"
```

初期設定では、すぐ試せるようSQLiteを使います。

## MySQLを使う場合

MySQL 8以上でデータベースとユーザーを用意し、同じPowerShellで環境変数を設定してからマイグレーションします。

```powershell
$env:DB_ENGINE = "mysql"
$env:MYSQL_DATABASE = "ipad_pos"
$env:MYSQL_USER = "pos_user"
$env:MYSQL_PASSWORD = "設定したパスワード"
$env:MYSQL_HOST = "127.0.0.1"
$env:MYSQL_PORT = "3306"
python manage.py migrate
python manage.py seed_demo
python manage.py runserver 0.0.0.0:8000
```

MySQL側のデータベースは `utf8mb4` で作成してください。アプリは環境変数でSQLite/MySQLを切り替えます。本番では `DJANGO_DEBUG=0`、固有の `DJANGO_SECRET_KEY`、適切な `DJANGO_ALLOWED_HOSTS`、HTTPS、認証・権限管理を必ず設定してください。この試作は店舗運用向けの認証や複数端末間のユーザー権限をまだ実装していません。

## 主な画面

- `/` ダッシュボード
- `/pos/` 販売・会計（商品検索、数量、顧客、支払方法選択）
- `/products/` 商品と在庫の登録・一覧
- `/customers/` 顧客登録・一覧
- `/receipts/` 領収書一覧。各領収書の詳細・CSVダウンロード
- `/reports/` 期間別売上、支払方法別集計、人気商品、CSV出力
- `/admin/` Django管理画面（必要なら `python manage.py createsuperuser` で管理者を作成）

販売確定時に商品価格をサーバー側で再取得し、在庫を検査・更新して販売、明細、デモ決済記録を同一トランザクションで保存します。領収書データ・レポートのCSVはExcel向けのUTF-8 BOM付きです。
