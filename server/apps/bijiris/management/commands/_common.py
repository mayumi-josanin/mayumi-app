"""ビジリスの取り込みコマンドで共通に使う小さな道具。

JSON は gas/ビジリスを書き出す.js が作る。中身は GAS が持っている形そのまま（camelCase）。
"""

import json
from datetime import datetime

from django.core.management.base import CommandError
from django.utils import timezone


def 文字(値) -> str:
    return "" if 値 is None else str(値).strip()


def 数(値):
    if 値 is None or 値 == "":
        return None
    try:
        return int(値)
    except (TypeError, ValueError):
        return None


def 日時(値):
    s = 文字(値)
    if not s:
        return None
    try:
        d = datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        return None
    if timezone.is_naive(d):
        d = timezone.make_aware(d, timezone.get_current_timezone())
    return d


def JSONを読む(生, 既定):
    """文字列なら JSON として読む。読めなければ (既定, 元の文字列) を返し、捨てない。"""
    if 生 is None or 生 == "":
        return 既定, ""
    if isinstance(生, (list, dict)):
        return 生, ""
    try:
        return json.loads(str(生)), ""
    except ValueError:
        return 既定, str(生)


def ファイルを開く(path):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError) as e:
        raise CommandError(f"読み込めませんでした: {e}")


def 違い(既存, 値: dict) -> list:
    """既存の行と新しい値で、違う項目の名前。"""
    return [k for k, v in 値.items() if getattr(既存, k) != v]
