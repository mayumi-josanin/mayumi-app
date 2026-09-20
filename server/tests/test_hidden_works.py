"""「隠す」の印が必ず効くこと（2026-09-20）。

メニューのプレビューが、押してもいないのに最初から出ていた。中身は空で、画像は壊れた形。
見た目の指定にある display:flex が、印を打ち消していたため。
ブラウザの既定では印の方が弱いので、見た目の指定で一度だけ強くしてある。
"""

from pathlib import Path

CSS = Path(__file__).resolve().parent.parent / "apps" / "manage" / "static" / "manage" / "style.css"


def test_隠す印を強くしてある():
    s = CSS.read_text(encoding="utf-8")
    assert "[hidden] { display: none !important; }" in s


def test_指定は1か所だけにする():
    """同じことを2か所に書かない。左メニューだけの指定はやめて、全部に効く1つにまとめた。"""
    s = CSS.read_text(encoding="utf-8")
    assert ".preview-backdrop" in s
    assert s.count("[hidden] { display: none !important; }") == 1
    assert ".sidebar-nav a[hidden]" not in s
