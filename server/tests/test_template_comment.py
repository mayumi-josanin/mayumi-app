"""テンプレートの注記が、画面にそのまま出ないこと（2026-09-21）。

スタンプ・特典の画面が、注記の文字で埋め尽くされた。
`{# … #}` は**1行専用**で、改行をまたぐと Django が注記として扱わず、そのまま出す。
複数行にしたいときは `{% comment %}` を使う。
（同じことで予約システムでもつまずいている。tests/test_screens.py に同じ見張りがある）
"""

import re
import pytest
from pathlib import Path

テンプレートの置き場 = Path(__file__).resolve().parent.parent / "apps"


def test_複数行の注記は使わない():
    悪い = []
    for p in テンプレートの置き場.rglob("*.html"):
        s = p.read_text(encoding="utf-8")
        for m in re.finditer(r"\{#(.*?)#\}", s, flags=re.S):
            if "\n" in m.group(1):
                行 = s[: m.start()].count("\n") + 1
                悪い.append(f"{p.relative_to(テンプレートの置き場)}:{行}")
    assert not 悪い, "複数行の {# #} は画面に出てしまいます。{% comment %} を使ってください: " + " / ".join(悪い)


@pytest.mark.django_db
def test_画面に注記が出ていない(as_owner):
    """実際に開いて、注記の記号が本文に出ていないこと。"""
    for url in ["/manage/rewards/", "/manage/calendar/new/", "/manage/menus/new/"]:
        page = as_owner.get(url).content.decode()
        assert "{#" not in page, url
        assert "{% comment" not in page, url
