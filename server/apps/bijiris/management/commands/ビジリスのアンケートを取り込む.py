"""ビジリスのアンケート定義（SURVEYS_JSON）を取り込む。

    python manage.py ビジリスのアンケートを取り込む ビジリスアンケート.json --下見
    python manage.py ビジリスのアンケートを取り込む ビジリスアンケート.json

JSON は gas/ビジリスを書き出す.js の `ビジリスのアンケート定義を書き出す()`（ビジリスの GAS で動かす）。
形は {"surveys": [SURVEYS_JSON の1本, …]}。

**アンケートIDで突き合わせるので、何度実行しても二重に増えない。**
設問は JSON のまま持つ（Survey.questions）。ここで形を直さない。GAS の `validateSurveyPayload_` が
返した形そのままを写す。**直すと default-surveys.js / Code.gs とずれる**（CLAUDE.md 6）。
"""

from django.core.management.base import BaseCommand
from django.db import transaction

from apps.bijiris.models import Survey

from ._common import ファイルを開く, 数, 文字, 日時, 違い

状態 = ("published", "draft", "archived")


def 行にする(s: dict) -> dict:
    status = 文字(s.get("status"))
    q = s.get("questions")
    return {
        "title": 文字(s.get("title"))[:255],
        "description": 文字(s.get("description")),
        "intro_message": 文字(s.get("introMessage")),
        "completion_message": 文字(s.get("completionMessage")),
        # normalizeSurveyStatus_ と同じ: 知らない値は published
        "status": status if status in 状態 else "published",
        "sort_order": 数(s.get("sortOrder")) or 0,
        "accepting_responses": s.get("acceptingResponses") is not False,
        "start_at": 日時(s.get("startAt")),
        "end_at": 日時(s.get("endAt")),
        "questions": q if isinstance(q, list) else [],
        "created_at": 日時(s.get("createdAt")),
        "updated_at": 日時(s.get("updatedAt")),
    }


class Command(BaseCommand):
    help = "ビジリスのアンケート定義（SURVEYS_JSON）の JSON を取り込む"

    def add_arguments(self, parser):
        parser.add_argument("json_path")
        parser.add_argument("--下見", action="store_true", dest="preview", help="何が起きるか見るだけ。書き込まない")

    def handle(self, *args, **options):
        生 = ファイルを開く(options["json_path"])
        一覧 = 生.get("surveys") if isinstance(生, dict) else 生
        if not isinstance(一覧, list):
            一覧 = []
        下見 = options["preview"]

        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        for s in 一覧:
            if not isinstance(s, dict):
                飛ばした += 1
                continue
            sid = 文字(s.get("id"))[:128]
            値 = 行にする(s)
            if not sid or not 値["title"]:
                飛ばした += 1
                continue
            既存 = Survey.objects.filter(survey_id=sid).first()
            if not 既存:
                新規.append((sid, 値))
            elif 違い(既存, 値):
                更新.append((sid, 値))
            else:
                変化なし += 1

        self.stdout.write("")
        self.stdout.write(f"■ ビジリスのアンケート: JSONに {len(一覧)}本")
        self.stdout.write(f"    新しく入る:   {len(新規)}本")
        self.stdout.write(f"    中身が変わる: {len(更新)}本")
        self.stdout.write(f"    変わらない:   {変化なし}本")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（id かタイトルが無い）: {飛ばした}本")
        for sid, 値 in 新規 + 更新:
            写真 = sum(1 for q in 値["questions"] if isinstance(q, dict) and q.get("type") == "photo")
            self.stdout.write(f"    - {値['title']}（{sid}）: {値['status']} / 設問 {len(値['questions'])}問（写真 {写真}問）")

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            for sid, 値 in 新規:
                Survey.objects.create(survey_id=sid, **値)
            for sid, 値 in 更新:
                Survey.objects.filter(survey_id=sid).update(**値)

        self.stdout.write("")
        self.stdout.write(f"    → いま {Survey.objects.count()}本")
