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
        # QRコード案内は旧管理アプリでも最後にあった（お客様アプリの住所はシステム管理で決める）
        ("QRコード案内", "qrcode", ["qrcode"]),
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

# 開発（apps/dev。KEM_DDENKI の「開発管理」の写し）。URL は manage の中に namespace "dev" で載せているので、
# url_name は "dev:project_list" のように書く（アプリ管理側の product_ などと頭がかぶらない）
開発 = [
    ("開発管理", "💻", [
        # タスクの画面（dev:task_*）は「タスク一覧」のタブに当てる。プロジェクトのタブに当てると、
        # タスクを開いた先で押した覚えのないタブが光る
        ("プロジェクト", "dev:project_list", ["dev:project_"]),
        ("タスク一覧", "dev:task_list", ["dev:task_"]),
        ("目安箱", "dev:meyasubako_list", ["dev:meyasubako_"]),
    ]),
]

# ビジリス（apps/bijiris。院長の決定 2026-09-16: 6画面。2026-09-17 に「お客様の画面」を足して7つ）。開発と同じく namespace "bijiris" 付きの url_name で当てる。
# 中身はまだ「準備中」（段取り A）。段取り B で置き換えても、この並びと url_name は変えない
ビジリス = [
    ("ビジリス管理", "💪", [
        ("集計", "bijiris:dashboard", ["bijiris:dashboard"]),
        ("アンケート管理", "bijiris:survey_list", ["bijiris:survey_"]),
        ("回答管理", "bijiris:response_list", ["bijiris:response_"]),
        ("顧客管理", "bijiris:customer_list", ["bijiris:customer_"]),
        ("特典", "bijiris:reward_list", ["bijiris:reward_"]),
        ("回数券分析", "bijiris:ticket_list", ["bijiris:ticket_"]),
        # 「お客様の画面」はここから外した（2026-09-18）。3つのお客様アプリをまとめた下の段へ引っ越し
    ]),
]

# お客様の画面（apps/preview。院長の依頼 2026-09-18）。お客様が見ている3つのアプリを、
# 「いま作っている方（develop）」と「お客様に出ている方（main）」で見比べる。namespace は "preview"
お客様の画面 = [
    ("お客様の画面", "👀", [
        ("まゆみ助産院アプリ", "preview:app", ["preview:app"]),
        ("ビジリス", "preview:bijiris", ["preview:bijiris"]),
        ("予約システム", "preview:reserve", ["preview:reserve"]),
    ]),
]

段 = [("アプリ管理", アプリ管理), ("公式サイト", 公式サイト), ("ビジリス", ビジリス),
     ("お客様の画面", お客様の画面), ("開発", 開発)]

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
    from .permissions import is_owner

    m = getattr(request, "resolver_match", None)
    url_name = (m.url_name if m else "") or ""
    user = getattr(request, "user", None)
    # **スタッフには予約管理だけ**（院長の決定 2026-09-16）。アプリ管理・公式サイト・開発の段は出さない
    まゆみ = bool(user is not None and getattr(user, "is_authenticated", False) and is_owner(user))
    # manage の下にさらに namespace がある画面（開発 = manage:dev:...）は、"dev:project_list" の形で当てる。
    # hp は namespace 無しで include しているので今までどおり素の url_name のまま
    if m and len(m.namespaces) > 1:
        url_name = ":".join(m.namespaces[1:]) + ":" + url_name
    sections = []
    active_tabs = None
    active_group = ""
    for 段名, まとまりたち in (段 if まゆみ else []):
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
