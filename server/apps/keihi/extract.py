"""レシートの写真を Claude に読ませて、日付・金額・店名・相手科目を取り出す。

鍵はビジリスの回数券分析と同じ `ANTHROPIC_API_KEY`（apps/bijiris/views_ticket.py と同じ呼び方。
SDK は入れず urllib で叩く）。**読み取った結果はそのまま帳簿に入れない。**
金額の桁や日付は読み違えることがあるので、院長が見て確かめてから出納帳の行にする。
"""

import base64
import json
import urllib.error
import urllib.request

from django.conf import settings

from .services import SUGGESTED_ACCOUNTS

API_URL = "https://api.anthropic.com/v1/messages"
API_VERSION = "2023-06-01"
MODEL = "claude-opus-5"
MAX_TOKENS = 4000  # 明細を拾うので長めに取る
通信の待ち秒 = 120

指示 = f"""レシートまたは領収書の写真です。次の項目を読み取って、JSON だけを返してください。

- date: 日付。"YYYY-MM-DD" の形。年が書かれていなければ今年とみなす。読めなければ null
- amount: 支払総額（税込）。整数のみ。読めなければ null
- store_name: 店名・支払先。読めなければ ""
- counter_account: 何に使ったお金かを、次の勘定科目から1つ選ぶ。当てはまるものが無ければ "雑費"
  {"、".join(SUGGESTED_ACCOUNTS)}
- kind: 支払いなら "payment"、入金・売上なら "income"
- note: 品目など、短い覚え書き。無ければ ""
- items: 買ったものの明細。[{{"name": "品名", "amount": 金額(整数)}}, ...] の並び。
  レシートに印字されているとおりに、**値引きや小計・合計の行は入れず**、品物だけを拾う。
  金額は印字されている数字をそのまま（税込・税抜の直しはしない）。読めなければ []

前置きも説明も書かず、JSON だけを出力してください。
読み取れない項目を推測で埋めないでください。読めないものは null または "" にしてください。"""


class 読み取れない(RuntimeError):
    """写真が読めない・API が答えないときに投げる。呼ぶ側は状態を failed にする。"""


def _取り出す(text: str) -> dict:
    """```json ... ``` で包まれて返ることがあるので、最初の { から最後の } までを拾う。"""
    始め, 終わり = text.find("{"), text.rfind("}")
    if 始め < 0 or 終わり <= 始め:
        raise 読み取れない(f"JSON が返りませんでした: {text[:200]}")
    try:
        return json.loads(text[始め:終わり + 1])
    except ValueError as e:
        raise 読み取れない(f"JSON として読めませんでした: {e}")


def 読み取る(画像: bytes) -> dict:
    """写真1枚から項目を取り出す。失敗したら 読み取れない を投げる。"""
    if not settings.ANTHROPIC_API_KEY:
        raise 読み取れない("ANTHROPIC_API_KEY が設定されていません")

    body = json.dumps({
        "model": MODEL,
        "max_tokens": MAX_TOKENS,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "image", "source": {
                    "type": "base64", "media_type": "image/jpeg",
                    "data": base64.b64encode(画像).decode("ascii"),
                }},
                {"type": "text", "text": 指示},
            ],
        }],
    }).encode("utf-8")

    req = urllib.request.Request(API_URL, data=body, method="POST", headers={
        "content-type": "application/json",
        "x-api-key": settings.ANTHROPIC_API_KEY,
        "anthropic-version": API_VERSION,
    })

    try:
        with urllib.request.urlopen(req, timeout=通信の待ち秒) as res:
            status, raw = res.status, res.read()
    except urllib.error.HTTPError as e:
        status, raw = e.code, e.read()
    except (urllib.error.URLError, OSError) as e:  # TimeoutError は OSError
        raise 読み取れない(f"Claude API に届きませんでした: {e}")

    try:
        answer = json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        answer = None
    if status != 200 or not isinstance(answer, dict):
        detail = (answer.get("error") or {}).get("message") if isinstance(answer, dict) else ""
        detail = detail or raw.decode("utf-8", "replace")[:300]
        raise 読み取れない(f"Claude API エラー ({status}): {detail}")

    text = "\n".join(
        b.get("text") or "" for b in answer.get("content") or []
        if isinstance(b, dict) and b.get("type") == "text"
    ).strip()
    if not text:
        raise 読み取れない("答えが空でした")

    中身 = _取り出す(text)
    中身["_raw"] = text
    return 中身
