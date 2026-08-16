"""
まゆみ助産院 データベース — Django 設定

置き場所は KEM_DDENKI のPC（Tailscale Funnel で公開）を想定している。
手元で確かめるときは PostgreSQL を立てずに済むよう、SQLite に切り替えられる。

    USE_SQLITE=true python manage.py migrate
"""

import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


def env(key, default=""):
    return os.environ.get(key, default)


def env_bool(key, default=False):
    value = env(key, "").strip().lower()
    if not value:
        return default
    return value in ("1", "true", "yes", "on")


SECRET_KEY = env("DJANGO_SECRET_KEY", "change-me-in-production")
DEBUG = env_bool("DJANGO_DEBUG", False)

# Funnel のホスト名を入れる。カンマ区切り。
ALLOWED_HOSTS = [h.strip() for h in env("DJANGO_ALLOWED_HOSTS", "localhost,127.0.0.1").split(",") if h.strip()]

# Tailscale serve / Funnel は TLS を終端して http で渡してくる。
# これを信頼させないと、POST がすべて 403 になる。
CSRF_TRUSTED_ORIGINS = [o.strip() for o in env("CSRF_TRUSTED_ORIGINS", "").split(",") if o.strip()]

INSTALLED_APPS = [
    "django.contrib.admin",
    "django.contrib.auth",
    "django.contrib.contenttypes",
    "django.contrib.sessions",
    "django.contrib.messages",
    "django.contrib.staticfiles",
    "apps.measurements",
    "apps.members",
]

MIDDLEWARE = [
    # CORS を一番外に置く。中で断ったとき（429 など）にも許可のヘッダが要るため。
    # 付いていないと、ブラウザには理由が伝わらず「通信できませんでした」としか出ない。
    "apps.core.middleware.CORSミドルウェア",
    "apps.core.middleware.呼び出し制限ミドルウェア",
    "apps.core.middleware.管理画面をしまうミドルウェア",
    "django.middleware.security.SecurityMiddleware",
    "whitenoise.middleware.WhiteNoiseMiddleware",
    "django.contrib.sessions.middleware.SessionMiddleware",
    "django.middleware.common.CommonMiddleware",
    "django.middleware.csrf.CsrfViewMiddleware",
    "django.contrib.auth.middleware.AuthenticationMiddleware",
    "django.contrib.messages.middleware.MessageMiddleware",
    "django.middleware.clickjacking.XFrameOptionsMiddleware",
]

ROOT_URLCONF = "config.urls"
WSGI_APPLICATION = "config.wsgi.application"

TEMPLATES = [
    {
        "BACKEND": "django.template.backends.django.DjangoTemplates",
        "DIRS": [BASE_DIR / "templates"],
        "APP_DIRS": True,
        "OPTIONS": {
            "context_processors": [
                "django.template.context_processors.request",
                "django.contrib.auth.context_processors.auth",
                "django.contrib.messages.context_processors.messages",
            ],
        },
    },
]

if env_bool("USE_SQLITE", False):
    DATABASES = {
        "default": {
            "ENGINE": "django.db.backends.sqlite3",
            "NAME": BASE_DIR / "手元確認用.sqlite3",
        }
    }
else:
    DATABASES = {
        "default": {
            "ENGINE": "django.db.backends.postgresql",
            "NAME": env("DB_NAME", "mayumi"),
            "USER": env("DB_USER", "postgres"),
            "PASSWORD": env("DB_PASSWORD", "postgres"),
            "HOST": env("DB_HOST", "db"),
            "PORT": env("DB_PORT", "5432"),
        }
    }

AUTH_PASSWORD_VALIDATORS = [
    {"NAME": "django.contrib.auth.password_validation.UserAttributeSimilarityValidator"},
    {"NAME": "django.contrib.auth.password_validation.MinimumLengthValidator"},
    {"NAME": "django.contrib.auth.password_validation.CommonPasswordValidator"},
    {"NAME": "django.contrib.auth.password_validation.NumericPasswordValidator"},
]

LANGUAGE_CODE = "ja"
TIME_ZONE = "Asia/Tokyo"
USE_I18N = True
# 日時は必ずタイムゾーン付きで扱う。スプレッドシートから移すときのズレを防ぐ。
USE_TZ = True

STATIC_URL = "static/"
STATIC_ROOT = BASE_DIR / "staticfiles"
STORAGES = {
    "default": {"BACKEND": "django.core.files.storage.FileSystemStorage"},
    "staticfiles": {"BACKEND": "whitenoise.storage.CompressedStaticFilesStorage"},
}

DEFAULT_AUTO_FIELD = "django.db.models.BigAutoField"

# ---------------------------------------------------------------
# GAS からの呼び出しを確かめる合鍵。
#
# この API は Funnel で公開されるため、インターネットから誰でも届く。
# GAS だけが知っている文字列を X-Api-Key ヘッダで送ってもらい、突き合わせる。
# 空のままだと API はすべて拒否する（設定し忘れて素通しになるのを防ぐ）。
#
# ※ アプリが直接呼ぶようになったら、この鍵はアプリに持たせられない。
#    アプリのJavaScriptは誰でも読めるため、書いた瞬間に誰でも全会員を読める。
#    その段階では「お客様1人ずつの合鍵」で本人を決める形にする。
# ---------------------------------------------------------------
API_KEY = env("API_KEY", "")

# ---------------------------------------------------------------
# アプリから直接呼ばれるための設定
# ---------------------------------------------------------------

# 住所をまたぐ通信を許す相手。**`*` にはしない。**
# `*` にすると、見知らぬページの JavaScript からも、お客様の合鍵さえあれば
# 読めてしまう。配っている場所だけを並べる。
CORS_ALLOWED_ORIGINS = [
    o.strip()
    for o in env(
        "CORS_ALLOWED_ORIGINS",
        "https://mayumi-josanin.github.io",
    ).split(",")
    if o.strip()
]

# 呼び出しの上限（1分あたり・相手のIPごと）。
# 家庭の回線とPCなので、常識的な使い方より少し上に置く。
# お客様1人がアプリを開いて出す通信は、多くても数十件。
THROTTLE_PER_MINUTE = int(env("THROTTLE_PER_MINUTE", "120") or 120)

# 当てずっぽうを繰り返される道は厳しくする（ログイン・パスコードの復旧など）。
THROTTLE_STRICT_PER_MINUTE = int(env("THROTTLE_STRICT_PER_MINUTE", "10") or 10)
THROTTLE_STRICT_PATHS = ["/api/login", "/api/recover", "/api/passcode"]

# Django の管理画面を開ける住所。**このPCの中からだけ。**
# Funnel は「この口に来たものは全部通す」ので、何もしないと
# 会員情報の管理画面がログイン欄ごとインターネットに出てしまう。
# 手元から見たいときは SSH の転送でこのPCの中を経由する。
ADMIN_ALLOWED_HOSTS = [
    h.strip()
    for h in env("ADMIN_ALLOWED_HOSTS", "localhost,127.0.0.1").split(",")
    if h.strip()
]

# 生死の確認は数えない。見張りが断られると「止まっている」と誤って知らせてしまう。
THROTTLE_EXEMPT_PATHS = ["/api/health"]

# 数を覚える場所。
# Django の既定はプロセスごとに別々に覚えるため、働き手が複数あると
# 実際には人数ぶん通ってしまう。データベース側の入れ物を使って揃える。
#
#   python manage.py createcachetable
#
# 手元確認（SQLite）のときは、その手間を省いてプロセスごとの入れ物にする。
if env_bool("USE_SQLITE", False):
    CACHES = {"default": {"BACKEND": "django.core.cache.backends.locmem.LocMemCache"}}
else:
    CACHES = {
        "default": {
            "BACKEND": "django.core.cache.backends.db.DatabaseCache",
            "LOCATION": "django_cache",
        }
    }
