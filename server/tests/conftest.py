import pytest
from django.contrib.auth.models import Group, User

from apps.manage.permissions import OWNER, STAFF, ensure_groups


@pytest.fixture(autouse=True)
def _plain_static(settings, tmp_path):
    settings.STORAGES = {
        "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
        "staticfiles": {"BACKEND": "django.contrib.staticfiles.storage.StaticFilesStorage"},
    }
    settings.MEDIA_ROOT = tmp_path / "media"
    settings.PUBLIC_BASE_URL = "https://api.example.com"


@pytest.fixture
def owner(db):
    ensure_groups()
    user = User.objects.create_user(username=OWNER, password="0714")
    user.groups.add(Group.objects.get(name=OWNER))
    return user


@pytest.fixture
def staff(db):
    ensure_groups()
    user = User.objects.create_user(username=STAFF, password="1234")
    user.groups.add(Group.objects.get(name=STAFF))
    return user


@pytest.fixture
def as_owner(client, owner):
    client.force_login(owner)
    return client


import io  # noqa: E402
import json  # noqa: E402

from apps.content import onesignal  # noqa: E402


@pytest.fixture
def fake_onesignal(monkeypatch, settings):
    settings.ONESIGNAL_APP_ID = "app"
    settings.ONESIGNAL_REST_API_KEY = "key"
    sent = []

    class _Res(io.BytesIO):
        def __enter__(self):
            return self

        def __exit__(self, *a):
            return False

    def fake_urlopen(req, timeout=0):
        sent.append(json.loads(req.data))
        return _Res(json.dumps({"id": "n-1", "recipients": 5}).encode())

    monkeypatch.setattr(onesignal.urllib.request, "urlopen", fake_urlopen)
    return sent


# ---------------------------------------------------------------------------
# 会員の窓口（test_gasapi_member*.py）向け
# ---------------------------------------------------------------------------

@pytest.fixture
def 会員試験の下ごしらえ(settings):
    """合鍵を入れ、回数制限の数え上げ（cache）を空にし、ハッシュを軽くする。

    ・合鍵: 管理者向けの窓口は X-Api-Key が要る。空だと 503 で全部止まる。
    ・cache: ログイン・復元の「10回はずしたらお休み」は cache で数えている。
      試験ごとに空にしないと、前の試験のはずれが次の試験に持ち越される。
    ・ハッシュ: 既定の PBKDF2 は1回 0.1 秒かかる。会員を何十人も作る試験で
      積み上がるので、試験では MD5 にする。照合の作法（check_password）は同じ。
    """
    from django.core.cache import cache

    from tests.gasapi_member_support import 合鍵

    settings.API_KEY = 合鍵
    settings.PASSWORD_HASHERS = ["django.contrib.auth.hashers.MD5PasswordHasher"]
    cache.clear()
    yield
    cache.clear()


# ═════════════════════════════════════════════════════════
# 公式サイト（apps/hp）: mayumi-site の写しを一時フォルダに作って使う
#
# サイトの作り方（admin/*.py）は mayumi-site リポジトリにあり、ここには写していない。
# 手元に mayumi-site が無ければ、公式サイトの試験は飛ばす（他の試験は影響しない）。
# ═════════════════════════════════════════════════════════
import os  # noqa: E402
import shutil  # noqa: E402
import subprocess  # noqa: E402

サイトの元 = os.environ.get("SITE_SOURCE_DIR", os.path.expanduser("~/Desktop/mayumi-site"))


def _git(cwd, *args):
    return subprocess.run(["git", "-c", "user.name=t", "-c", "user.email=t@example.com", *args],
                          cwd=cwd, capture_output=True, text=True, check=True)


@pytest.fixture(scope="session")
def site_repo_source(tmp_path_factory):
    """mayumi-site の写し（assets・控え・プレビュー・.git を除く）を1回だけ作り、
    起点（origin）として bare リポジトリも用意する。各試験はここから複製して使う。"""
    if not os.path.isfile(os.path.join(サイトの元, "admin", "core.py")):
        pytest.skip("mayumi-site が手元に無いので公式サイトの試験は飛ばす")
    base = tmp_path_factory.mktemp("site")
    work = base / "work"

    def 除く(d, names):
        rel = os.path.relpath(d, サイトの元)
        out = set()
        for n in names:
            if n in (".git", "__pycache__", "node_modules") or (rel == "." and n in ("assets", "docs")) \
               or (rel == "admin" and n in ("backups", "_preview")):
                out.add(n)
        return out

    shutil.copytree(サイトの元, work, ignore=除く)
    # プレビューなどが参照する assets は空の器だけ用意する
    (work / "assets" / "img").mkdir(parents=True, exist_ok=True)
    # 本物の下書きには「【下書き】」の印が残っていて、公開前の点検がそこで止まる（正しい動き）。
    # 公開までを試すために、写しからは印だけ外す。
    for root, _dirs, files in os.walk(work):
        if "/admin" in root or "/.git" in root:
            continue
        for fn in files:
            if fn.endswith(".html"):
                fp = os.path.join(root, fn)
                t = open(fp, encoding="utf-8").read()
                t2 = t.replace("【下書き】", "").replace("【要確認】", "")
                if t2 != t:
                    open(fp, "w", encoding="utf-8").write(t2)
    _git(work, "init", "-q", "-b", "draft")
    _git(work, "add", "-A")
    _git(work, "commit", "-q", "-m", "写し")
    _git(work, "branch", "main")
    origin = base / "origin.git"
    _git(base, "init", "-q", "--bare", str(origin))
    _git(work, "remote", "add", "origin", str(origin))
    _git(work, "push", "-q", "-u", "origin", "draft", "main")
    return base


@pytest.fixture
def site_repo(site_repo_source, tmp_path, settings):
    """試験ごとの写し。settings.SITE_REPO_DIR をここへ向ける。"""
    dst = tmp_path / "site"
    shutil.copytree(site_repo_source / "work", dst, symlinks=True)
    origin = tmp_path / "origin.git"
    shutil.copytree(site_repo_source / "origin.git", origin)
    _git(dst, "remote", "set-url", "origin", str(origin))
    settings.SITE_REPO_DIR = str(dst)
    settings.SITE_GIT_TOKEN = ""
    from apps.hp import repo as _repo

    _repo._読んだ場所["dir"] = None
    return dst
