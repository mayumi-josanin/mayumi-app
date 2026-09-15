"""管理画面のあいだでログインを受け渡す（アプリ管理 ⇄ 予約管理）。

2つの Django は別の箱・別のデータベースだが、使う人は同じ（まゆみ／スタッフ）。
片方で入ったら、もう片方はログインし直さずに開けるようにする。

仕組み: 役割と期限を書いた札に、**両方が同じ合鍵（MANAGE_SSO_SECRET）**で署名を付けて渡す。
受け取った側は署名と期限を確かめ、その役割の人として login() する。
札は60秒で切れ、一度使ったものは受け付けない（キャッシュに控える）。

**合鍵は .env に置く。**両方の .env に同じ値。空なら受け渡しは動かない
（メニューは出るが、押すと相手のログイン画面に着く）。
mayumi-reserve/apps/core/sso.py に同じ中身がある。**変えるときは両方。**
"""

import base64
import hashlib
import hmac
import json
import secrets
import time

TTL_SECONDS = 60


def _b64(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode("ascii").rstrip("=")


def _unb64(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def 札を作る(role: str, secret: str) -> str:
    body = json.dumps({"role": role, "exp": int(time.time()) + TTL_SECONDS, "nonce": secrets.token_hex(8)},
                      ensure_ascii=False).encode("utf-8")
    sig = hmac.new(secret.encode("utf-8"), body, hashlib.sha256).digest()
    return _b64(body) + "." + _b64(sig)


def 札を確かめる(token: str, secret: str) -> dict | None:
    """通れば {"role", "nonce"}。駄目なら None。"""
    if not secret or not token or "." not in token:
        return None
    try:
        body_s, sig_s = token.split(".", 1)
        body = _unb64(body_s)
        sig = _unb64(sig_s)
    except Exception:
        return None
    if not hmac.compare_digest(sig, hmac.new(secret.encode("utf-8"), body, hashlib.sha256).digest()):
        return None
    try:
        中 = json.loads(body.decode("utf-8"))
    except Exception:
        return None
    if int(中.get("exp") or 0) < int(time.time()):
        return None
    if not 中.get("role") or not 中.get("nonce"):
        return None
    return {"role": str(中["role"]), "nonce": str(中["nonce"])}
