"""公式サイトの「反映・公開・戻す」の実体。mayumi-site/admin/app.py の同名処理と同じ順・同じ文言。

app.py は HTTP サーバーごと1つのファイルなので import せず、必要な3つだけをここに写した
（publish_everything / restore_backup / 公開の2段階）。作り方そのもの（core/news/blog/…）は
向こうの部品を呼ぶ。
"""

import os
import shutil

from . import repo


def 反映する(data: dict) -> list[str]:
    """app.py の publish_everything。目印を差し替え、お知らせ・お教室・診療カレンダー・ブログを作り直す。"""
    core = repo.部品("core")
    news = repo.部品("news")
    classroom = repo.部品("classroom")
    blog = repo.部品("blog")
    log = core.publish(data)
    for name, fn in (
        ("お知らせ", lambda: (news.build(), news.update_sitemap())[0]),
        ("お教室", lambda: (classroom.build(), classroom.update_sitemap())[0]),
    ):
        try:
            log.append("%sを作り直しました（%d件）。" % (name, fn()))
        except Exception as ex:
            log.append("%sの作り直しに失敗: %s" % (name, ex))
    try:
        calendar_block = repo.部品("calendar_block")
        cnt = calendar_block.build_pages()
        calendar_block.update_sitemap()
        log.append("診療カレンダーに %d件の予定を入れました（詳しいページ %d件）。"
                   % (len(calendar_block.collect(data)), cnt))
    except Exception as ex:
        log.append("診療カレンダーの作り直しに失敗: %s" % ex)
    try:
        n, pages, made = blog.build()
        blog.update_sitemap()
        log.append("ブログを作り直しました（記事%d本／HTML%d個）。" % (n, made))
    except Exception as ex:
        log.append("ブログの作り直しに失敗: %s" % ex)
    return log


def 控えの一覧() -> list[str]:
    core = repo.部品("core")
    if not os.path.isdir(core.BACKUP_DIR):
        return []
    return sorted([b for b in os.listdir(core.BACKUP_DIR) if not b.startswith("_")], reverse=True)[:10]


def 控えに戻す(name: str) -> list[str]:
    """app.py の restore_backup。控え（HTML の丸ごと）を書き戻す。"""
    core = repo.部品("core")
    if not name or "/" in name or "\\" in name or name.startswith("."):
        return ["そのバックアップは見つかりませんでした。"]
    src = os.path.join(core.BACKUP_DIR, name)
    if not os.path.isdir(src):
        return ["そのバックアップは見つかりませんでした。"]
    n = 0
    for root, _, files in os.walk(src):
        for fn in files:
            s = os.path.join(root, fn)
            d = os.path.join(core.SITE, os.path.relpath(s, src))
            os.makedirs(os.path.dirname(d), exist_ok=True)
            shutil.copy2(s, d)
            n += 1
    return ["%s の時点に戻しました（%d ファイル）。" % (name, n)]


def 目印の状態() -> dict:
    core = repo.部品("core")
    found, missing, unused = core.check_markers()
    return {"found": sorted((k, ", ".join(v)) for k, v in found.items()), "missing": missing, "unused": unused}


def 公開の下ごしらえ(message: str) -> tuple[bool, list[str]]:
    """app.py の /api/publish-all step=prepare: 反映 → 下書き（draft）に保存 → 点検。**まだ外には出ない。**"""
    core = repo.部品("core")
    gitops = repo.部品("gitops")
    release = repo.部品("release")
    log = []
    try:
        log += 反映する(core.load())
    except Exception as ex:
        return False, log + ["反映に失敗: %s" % ex]
    log.append("")
    log += gitops.commit_push(message or "サイトの内容を更新")
    log.append("")
    ok, clog = release.checks()
    return ok, log + clog


def 公開する() -> tuple[bool, list[str]]:
    """app.py の /api/publish-all step=go: main へ取り込んで送信し、draft に戻る。"""
    release = repo.部品("release")
    return release.release()
