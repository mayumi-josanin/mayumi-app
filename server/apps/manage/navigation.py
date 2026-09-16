"""左メニューの組み立て。**画面の中身と URL はそのまま、入口だけをまとめる**（院長の相談 2026-09-16）。

- 左メニューは「まとまり」（会員／アプリの掲載／通知／注文と売上／設定 …）を出す
- まとまりを開くと、上部のタブで中の画面（今までの画面そのもの）を切り替える
- どのタブにいるかは url_name の頭で決める（`一致` の並び）

まとまりを増やす・並びを変えるのはここだけ。テンプレートは触らない。
"""

from django.urls import reverse

# (ラベル, 絵文字, タブの並び)。タブは (ラベル, url_name, [一致する url_name の頭])
アプリ管理 = [
    ("会員", "👥", [
        ("会員一覧", "member_list", ["member_"]),
        ("スタンプ・特典", "reward_list", ["reward_"]),
        ("重複候補", "duplicate_list", ["duplicate_"]),
    ]),
    ("アプリの掲載", "📱", [
        ("NEWS（アプリ）", "news_list", ["news_"]),
        ("カレンダー", "calendar_list", ["calendar_"]),
        ("ホームのメニュー", "menu_list", ["menu_"]),
        ("ショップの商品", "product_list", ["product_"]),
        ("お知らせ一覧の見え方", "notice_list", ["notice_"]),
    ]),
    ("通知", "🔔", [
        ("通知の送信と履歴", "push_list", ["push_"]),
    ]),
    ("注文と売上", "📦", [
        ("注文", "order_list", ["order_"]),
        ("データ分析", "analytics", ["analytics"]),
        ("売上の記録（メニュー）", ("revenue_list", {"kind": "menu"}), []),
        ("売上の記録（商品）", ("revenue_list", {"kind": "product"}), []),
    ]),
    ("設定", "🛠️", [
        ("カテゴリ", "category_list", ["category_"]),
        ("ゴミ箱", "trash_list", ["trash_"]),
        ("システム管理", "system", ["system"]),
    ]),
]

公式サイト = [
    ("記事を書く", "✍️", [
        ("お知らせ（サイト）", "hp_news", ["hp_news"]),
        ("まゆみのつぶやき", "hp_blog", ["hp_blog"]),
        ("お教室", "hp_classes", ["hp_class"]),
        ("診療カレンダー", "hp_closed", ["hp_closed"]),
    ]),
    ("ページを整える", "🧩", [
        ("基本情報", "hp_basic", ["hp_basic"]),
        ("ページの編集", "hp_layout", ["hp_layout"]),
        ("写真", "hp_images", ["hp_images"]),
        ("選択肢", "hp_options", ["hp_options"]),
    ]),
    ("確かめて公開", "🌐", [
        ("プレビュー", "hp_preview", ["hp_preview", "hp_site_file"]),
        ("公開する", "hp_publish", ["hp_publish"]),
        ("保存の記録", "hp_git", ["hp_git"]),
    ]),
]

段 = [("アプリ管理", アプリ管理), ("公式サイト", 公式サイト)]

# 予約管理のまとまり（mayumi-reserve/apps/core/navigation.py と同じ並び）。押すと go_reserve で予約システムへ飛び、
# 向こうの画面に上部タブが出る。このメニューはまゆみだけが見る（スタッフはアプリ管理に入れない）
予約管理 = [("予約", "🗓️", "/manage/"), ("予約枠と休診", "🚫", "/manage/blocks/grid/"), ("予約メニューと質問", "📖", "/manage/menus/"),
        ("売上・産後ケア", "💴", "/manage/sales/"), ("予約の設定", "⚙️", "/manage/staff/")]


def _url(名):
    if isinstance(名, tuple):
        return reverse("manage:" + 名[0], kwargs=名[1])
    return reverse("manage:" + 名)


def _当たる(url_name: str, タブ) -> bool:
    ラベル, 名, 頭たち = タブ
    if isinstance(名, tuple):
        # 売上の記録は kind で分かれる。url_name だけでは区別できないので、ここでは当てない（タブの見た目は両方とも通常）
        return False
    if url_name == 名:
        return True
    return any(url_name.startswith(h) for h in 頭たち)


def 組み立てる(request):
    """テンプレートに渡す形。いまの url_name から、開いているまとまりとタブを決める。"""
    m = getattr(request, "resolver_match", None)
    url_name = (m.url_name if m else "") or ""
    sections = []
    active_tabs = None
    active_group = ""
    for 段名, まとまりたち in 段:
        groups = []
        for ラベル, 絵, タブたち in まとまりたち:
            当たり = any(_当たる(url_name, t) for t in タブたち)
            # 売上の記録（kind 付き）は url_name が revenue_list。まとまりとしては「注文と売上」に当てる
            if not 当たり and url_name == "revenue_list" and ラベル == "注文と売上":
                当たり = True
            tabs = [{"label": t[0], "url": _url(t[1]), "active": _当たる(url_name, t)
                     or (url_name == "revenue_list" and isinstance(t[1], tuple)
                         and request.resolver_match.kwargs.get("kind") == t[1][1]["kind"])}
                    for t in タブたち]
            groups.append({"label": ラベル, "icon": 絵, "url": tabs[0]["url"], "active": 当たり, "tabs": tabs})
            if 当たり:
                active_tabs = tabs
                active_group = ラベル
        sections.append({"label": 段名, "groups": groups})
    go = reverse("manage:go_reserve")
    sections.append({"label": "予約管理", "groups": [{"label": l, "icon": i, "url": f"{go}?to={p}", "active": False, "tabs": []}
                                                 for l, i, p in 予約管理]})
    return {"nav_sections": sections, "nav_tabs": active_tabs, "nav_group": active_group}
