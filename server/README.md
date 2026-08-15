# server — まゆみ助産院 データベース

スプレッドシートから表を1つずつ引っ越すための、Django + PostgreSQL。

設計の全体像は [../docs/design/移行設計.md](../docs/design/移行設計.md) を参照。
**この段階では読み取りしか受け付けない。**スプレッドシートがまだ正で、
ここは写しにすぎないため、書けてしまうと「どちらが正しいか」が崩れる。

---

## いまできること

| | |
|---|---|
| 計測記録（測定履歴42件）の保管 | ✅ |
| GAS から読み取るAPI | ✅ |
| スプレッドシートからの取り込み（何度でも実行可） | ✅ |
| 書き込みAPI | ❌ まだ作らない |
| 会員・アンケート・回数券 | ❌ 表を1つずつ増やす |

## 手元で試す（PostgreSQL なしで動く）

```bash
cd server
python3 -m venv .venv
./.venv/bin/pip install -e .

USE_SQLITE=true ./.venv/bin/python manage.py migrate
USE_SQLITE=true API_KEY=test-key-1234 ./.venv/bin/python manage.py runserver 127.0.0.1:8123
```

SQLite に切り替わるだけで、**コードは本番と同じ**。

## PCに置く（KEM_DDENKI のPC）

```bash
cp .env.example .env      # 値を入れる
docker compose up -d
docker compose exec web python manage.py migrate
```

ポートは KEM_DDENKI と重ならないようにしてある（DB 5434 / Web 8002）。

Funnel が使えるのは 443 / 8443 / 10000 の3つ。KEM が 443・8443 を使っているので
**10000 を使えば新しいノードは要らない**。

```powershell
tailscale funnel --bg --https=10000 http://127.0.0.1:8002
```

PCでの手順は [../docs/サーバー設置手順.md](../docs/サーバー設置手順.md) に一通り書いてある。

> **前提**：このPCは Docker Desktop が「サインイン時に自動起動」のため、
> 再起動後に誰かがサインインするまでDBが上がらない。
> BIOSの電源復帰・Windows自動サインイン・スリープ無効の3点を先に設定すること。

## API

| | 鍵 | 内容 |
|---|---|---|
| `GET /api/health` | 不要 | 生きているかだけを返す。PCが落ちていないかの確認用 |
| `GET /api/measurements` | **必要** | 計測記録。`?customerName=` `?memberId=` で絞れる |

鍵は `X-Api-Key` ヘッダで送る。`API_KEY` が空のときは**すべて拒否**する
（設定し忘れて素通しになる事故を防ぐため）。

```bash
curl -H "X-Api-Key: $API_KEY" https://<ホスト>/api/measurements
```

## 取り込み

スプレッドシートから書き出した JSON を読み込む。

```bash
# まず下見（書き込まない）
python manage.py 計測記録を取り込む 測定履歴.json --下見

# 実行
python manage.py 計測記録を取り込む 測定履歴.json
```

**「測定ID」で突き合わせるので、何度実行しても二重に増えない。**
中身が変わっている行だけ更新する。

## 構成

```
server/
├── config/              Django 設定
├── apps/
│   └── measurements/    計測記録
│       ├── models.py    表の定義
│       ├── views.py     読み取りAPI
│       └── management/commands/計測記録を取り込む.py
├── docker-compose.yml   PC に置くとき
└── .env.example
```

## 決めごと

- **会員IDは空を許す。** アンケート送信に会員登録が要らないため、
  「記録はあるが会員ではない方」がいつでも発生しうる。必須にすると移せなくなる。
- **金額・数値は Decimal。** 小数の誤差を持ち込まない。
- **日時は必ずタイムゾーン付き**（`USE_TZ=True`）。
- **表を1つ移すたびに、住む場所は1つに決める。**
  スプレッドシートとDBの両方に書ける状態を作らない。
