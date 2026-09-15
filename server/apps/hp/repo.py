"""公式サイトのリポジトリ（mayumi-site）と、その admin/ の部品を呼ぶ土台。

    場所()      … 設定 SITE_REPO_DIR。無ければ 設定がない を投げる
    部品("core") … mayumi-site/admin/core.py などを読み込んで返す

## なぜ写さずに呼ぶのか

サイトの作り方（目印の差し替え・お知らせやブログのページの生成・公開前の点検）は
mayumi-site/admin/*.py にある。ここへ写すと、向こうを直したときに二重になる。
**リポジトリを1つ置き、その中の部品をそのまま使う。**

## git の認証

GitHub の鍵は設定 SITE_GIT_TOKEN（環境変数）で受け、git の環境変数
（GIT_CONFIG_*）で渡す。**.git/config には書かない**（リポジトリを開けば見えてしまう）。
"""

import base64
import importlib
import os
import sys

from django.conf import settings

部品の名前 = ("core", "news", "blog", "classroom", "layout", "release", "gitops", "calendar_block")


class 設定がない(Exception):
    pass


def 場所() -> str:
    d = (getattr(settings, "SITE_REPO_DIR", "") or "").strip()
    if not d or not os.path.isdir(os.path.join(d, "admin")) or not os.path.isfile(os.path.join(d, "admin", "core.py")):
        raise 設定がない()
    return os.path.abspath(d)


def 設定されているか() -> bool:
    try:
        場所()
        return True
    except 設定がない:
        return False


_読んだ場所 = {"dir": None}


def _gitの環境():
    """git を呼ぶときの設定を環境変数で渡す（.git/config に書かない）。"""
    項目 = [
        ("safe.directory", "*"),  # 箱の中では持ち主が違って見えるため
        ("user.name", getattr(settings, "SITE_GIT_NAME", "") or "まゆみ助産院 管理画面"),
        ("user.email", getattr(settings, "SITE_GIT_EMAIL", "") or "manage@mayumijosanin.com"),
    ]
    鍵 = (getattr(settings, "SITE_GIT_TOKEN", "") or "").strip()
    if 鍵:
        素 = base64.b64encode(f"x-access-token:{鍵}".encode()).decode()
        項目.append(("http.https://github.com/.extraheader", "AUTHORIZATION: basic " + 素))
    os.environ["GIT_CONFIG_COUNT"] = str(len(項目))
    for i, (k, v) in enumerate(項目):
        os.environ[f"GIT_CONFIG_KEY_{i}"] = k
        os.environ[f"GIT_CONFIG_VALUE_{i}"] = v
    os.environ["GIT_TERMINAL_PROMPT"] = "0"


def 部品(名: str):
    """mayumi-site/admin/<名>.py を読み込んで返す。場所が変わっていれば読み直す。"""
    d = 場所()
    admin = os.path.join(d, "admin")
    if _読んだ場所["dir"] != admin:
        # 部品どうしは `import core` のように名前で呼び合うので、sys.path の先頭に置く。
        sys.path[:] = [p for p in sys.path if p != _読んだ場所["dir"]]
        sys.path.insert(0, admin)
        for n in 部品の名前:
            sys.modules.pop(n, None)
        _読んだ場所["dir"] = admin
    _gitの環境()
    return importlib.import_module(名)


def 相対(パス: str) -> str:
    """リポジトリの中の相対パス（表示用）。"""
    try:
        return os.path.relpath(パス, 場所())
    except Exception:
        return パス
