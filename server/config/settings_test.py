"""テスト用の設定。SQLite・管理画面あり・DEBUG。

環境変数で切り替える作りなので、settings.py を読む前にここで決める。
pytest は pyproject.toml の DJANGO_SETTINGS_MODULE からこれを読む。
"""

import os

os.environ.setdefault("USE_SQLITE", "true")
os.environ.setdefault("MANAGE_ENABLED", "true")
os.environ.setdefault("DJANGO_DEBUG", "true")
# 呼び出し制限（1分120回）はテストの数が増えると同じ相手として引っかかる。テストでは外す。
os.environ.setdefault("THROTTLE_PER_MINUTE", "0")
os.environ.setdefault("THROTTLE_STRICT_PER_MINUTE", "0")
os.environ.setdefault("ADMIN_TOKEN_SECRET", "test-token-secret")

from .settings import *  # noqa: E402,F401,F403
