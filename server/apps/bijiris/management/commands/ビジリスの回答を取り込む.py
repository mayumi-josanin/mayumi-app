"""ビジリスの回答（シート「回答一覧」＋アンケートごとのシート）を取り込む。写真は任意で Drive から移す。

    python manage.py ビジリスの回答を取り込む ビジリス回答.json --下見
    python manage.py ビジリスの回答を取り込む ビジリス回答.json
    python manage.py ビジリスの回答を取り込む ビジリス回答.json --写真
    python manage.py ビジリスの回答を取り込む ビジリス回答.json --写真 --写真フォルダ ~/Downloads/bijiris_photos

JSON は gas/ビジリスを書き出す.js の `ビジリスの回答を書き出す()`（まゆみの GAS から openById で読むだけ）。
形は {"responses": [{row, submittedAt, id, surveyId, surveyTitle, customerClientId, customerName, customerEmail,
                     status, adminMemo, answers, files, managedAt, surveySheet: {row, values: {見出し: 値}}}, …]}

**回答IDで突き合わせるので、何度実行しても二重に増えない。**
写真も (回答, 設問ID, Drive のファイルID) で突き合わせ、すでに media に入っているものは取りに行かない。

写真の取り方（--写真）:
  1. --写真フォルダ があれば、そこにある `<DriveのファイルID>.*` を使う（Drive からまとめて落としたもの）
  2. 無ければ `https://drive.google.com/uc?export=download&id=<ファイルID>` から取る。
     **お客様の計測写真は 2026-08-25 からリンク共有をやめている**（Code.gs `savePhotoFiles_`）ので、
     この道は「動かしているアカウントが Drive に入れる」ときだけ通る。取れなかった枚数は数えて出す。
     取れなかったぶんは、院長のアカウントで Drive からフォルダごと落として --写真フォルダ で入れ直す。
写真は apps/manage/images.保存する(…, "bijiris") で縮めて media/bijiris/<乱数>.jpg に置く。**Drive のリンクは持たない。**

**お名前で会員に結びつけない。**member_id は空のまま。一致する件数だけ数えて出す。
"""

import io
import urllib.request
from pathlib import Path

from django.core.management.base import BaseCommand
from django.db import transaction

from apps.bijiris.models import Response, ResponsePhoto, Survey
from apps.manage import images
from apps.members.models import Member

from ._common import JSONを読む, ファイルを開く, 数, 文字, 日時, 違い

対応状況 = ("new", "checked", "done", "trash")

# Code.gs の判定と同じ（TICKET_SURVEY_*_PHOTO_QUESTION_IDS / measureTimingOf_）
ビフォーの設問 = ("q_bijiris_session_monitor_photos_6", "q_bijiris_session_monitor_photos_10", "q_bijiris_session_monitor_photos")
アフターの設問 = ("q_bijiris_session_ticket_end_photos_6", "q_bijiris_session_ticket_end_photos_10",
            "q_bijiris_session_ticket_end_photos", "q_ticket_end_photo_last")
計測の写真設問 = "q_measure_photos"
計測のタイミング設問 = "q_measure_timing"


def 写真の種別(question_id: str, answers: list) -> str:
    if question_id in ビフォーの設問:
        return "before"
    if question_id in アフターの設問:
        return "after"
    if question_id == 計測の写真設問:
        timing = ""
        for a in answers:
            if isinstance(a, dict) and a.get("questionId") == 計測のタイミング設問:
                v = a.get("value")
                timing = "".join(v) if isinstance(v, list) else 文字(v)
        if "初回計測" in timing or "モニター" in timing:
            return "before"
        if "終了" in timing:
            return "after"
    return "other"


def 写真を集める(answers: list, files: list) -> list:
    """回答JSONの answers[].files[] から (設問ID, ファイル) を並び順に。無ければ写真JSON（設問IDは空）。"""
    出 = []
    for a in answers:
        if not isinstance(a, dict) or not isinstance(a.get("files"), list):
            continue
        qid = 文字(a.get("questionId"))
        for i, f in enumerate(a["files"]):
            if isinstance(f, dict) and 文字(f.get("fileId")):
                出.append((qid, i, f))
    if 出:
        return 出
    return [("", i, f) for i, f in enumerate(files) if isinstance(f, dict) and 文字(f.get("fileId"))]


def 行にする(r: dict) -> dict:
    answers, answers_raw = JSONを読む(r.get("answers"), [])
    files, files_raw = JSONを読む(r.get("files"), [])
    status = 文字(r.get("status"))
    sheet = r.get("surveySheet") if isinstance(r.get("surveySheet"), dict) else {}
    values = sheet.get("values") if isinstance(sheet.get("values"), dict) else {}
    return {
        "survey_key": 文字(r.get("surveyId"))[:128],
        "survey_title": 文字(r.get("surveyTitle"))[:255],
        "submitted_at": 日時(r.get("submittedAt")),
        "client_id": 文字(r.get("customerClientId"))[:128],
        # **お名前から会員IDを埋めない。**
        "member_id": "",
        "customer_name": 文字(r.get("customerName"))[:255],
        "customer_email": 文字(r.get("customerEmail"))[:255],
        # normalizeStatus_ と同じ: 知らない値は new
        "status": status if status in 対応状況 else "new",
        "admin_memo": 文字(r.get("adminMemo")),
        "answers": answers if isinstance(answers, list) else [],
        "files": files if isinstance(files, list) else [],
        "answers_raw": answers_raw,
        "files_raw": files_raw,
        "managed_at": 日時(r.get("managedAt")),
        "survey_sheet_row": 数(sheet.get("row")),
        "survey_sheet_values": values,
    }


def _画像を取る(file_id: str, フォルダ):
    """写真の中身（bytes）。フォルダにあればそれ、無ければ Drive。取れなければ None。"""
    if フォルダ:
        for p in sorted(Path(フォルダ).glob(f"{file_id}.*")):
            if p.is_file():
                return p.read_bytes()
    try:
        req = urllib.request.Request(f"https://drive.google.com/uc?export=download&id={file_id}",
                                     headers={"User-Agent": "mayumi-server"})
        with urllib.request.urlopen(req, timeout=30) as res:
            種類 = 文字(res.headers.get("Content-Type"))
            中身 = res.read()
    except Exception:
        return None
    # 共有されていないと HTML のログイン画面が返る。画像でなければ取れなかったとみなす
    if not 種類.startswith("image/"):
        return None
    return 中身


class Command(BaseCommand):
    help = "ビジリスの回答の JSON を取り込む（--写真 で Drive の写真も media へ）"

    def add_arguments(self, parser):
        parser.add_argument("json_path")
        parser.add_argument("--下見", action="store_true", dest="preview", help="何が起きるか見るだけ。書き込まない")
        parser.add_argument("--写真", action="store_true", dest="photos", help="写真を Drive（または --写真フォルダ）から media へ入れる")
        parser.add_argument("--写真フォルダ", dest="photo_dir", default="", help="<DriveのファイルID>.<拡張子> で置いた写真のフォルダ")

    def handle(self, *args, **options):
        生 = ファイルを開く(options["json_path"])
        一覧 = 生.get("responses") if isinstance(生, dict) else 生
        if not isinstance(一覧, list):
            一覧 = []
        下見 = options["preview"]

        調べ = {s.survey_id: s for s in Survey.objects.all()}
        新規, 更新, 変化なし, 飛ばした = [], [], 0, 0
        写真の予定 = []  # (response_id, [(qid, order, file)])
        読めない = []
        for r in 一覧:
            if not isinstance(r, dict):
                飛ばした += 1
                continue
            rid = 文字(r.get("id"))[:128]
            if not rid:
                飛ばした += 1
                continue
            値 = 行にする(r)
            値["sheet_row"] = 数(r.get("row"))
            値["survey"] = 調べ.get(値["survey_key"])
            if 値["answers_raw"] or 値["files_raw"]:
                読めない.append(rid)
            既存 = Response.objects.filter(response_id=rid).first()
            if not 既存:
                新規.append((rid, 値))
            elif 違い(既存, 値):
                更新.append((rid, 値))
            else:
                変化なし += 1
            写真の予定.append((rid, 値, 写真を集める(値["answers"], 値["files"])))

        self.stdout.write("")
        self.stdout.write(f"■ ビジリスの回答: JSONに {len(一覧)}件"
                          + (f"（回答一覧は{生['シートの行数']}行）" if isinstance(生, dict) and 生.get("シートの行数") else ""))
        self.stdout.write(f"    新しく入る:   {len(新規)}件")
        self.stdout.write(f"    中身が変わる: {len(更新)}件")
        self.stdout.write(f"    変わらない:   {変化なし}件")
        if 飛ばした:
            self.stdout.write(f"    飛ばした（回答IDが無い）: {飛ばした}件")
        状態 = {}
        for _, 値, _ in 写真の予定:
            状態[値["status"]] = 状態.get(値["status"], 0) + 1
        self.stdout.write("    対応状況: " + " / ".join(f"{k} {n}件" for k, n in sorted(状態.items(), key=lambda x: -x[1])))
        アンケート無し = sorted({値["survey_key"] for _, 値, _ in 写真の予定 if 値["survey"] is None})
        if アンケート無し:
            self.stdout.write(f"    **アンケートの定義が無い回答: {'・'.join(アンケート無し)}**（先に ビジリスのアンケートを取り込む）")
        枚数 = sum(len(p) for _, _, p in 写真の予定)
        self.stdout.write(f"    写真: のべ {枚数}枚（{sum(1 for _, _, p in 写真の予定 if p)}件の回答）")
        if 読めない:
            self.stdout.write(f"    回答JSON か 写真JSON が読めなかった回答: {len(読めない)}件（元の文字列のまま残します）: {読めない[:5]}")

        # **お名前の照合は数えるだけ。結びつけない。**
        名前ら = {値["customer_name"] for _, 値, _ in 写真の予定 if 値["customer_name"]}
        一致 = sum(1 for n in 名前ら if Member.objects.filter(name=n).count() == 1)
        複数 = [n for n in 名前ら if Member.objects.filter(name=n).count() > 1]
        self.stdout.write(f"  ● 会員のお名前との一致（**数えるだけ。結びつけていません**）: {len(名前ら)}名のうち ちょうど1名と一致 {一致}名"
                          + (f" / 同じお名前が複数 {'・'.join(sorted(複数))}" if 複数 else ""))

        if 下見:
            self.stdout.write("")
            self.stdout.write("■ 下見なので、何も書いていません。よければ --下見 を外して実行してください。")
            return

        with transaction.atomic():
            for rid, 値 in 新規:
                Response.objects.create(response_id=rid, **値)
            for rid, 値 in 更新:
                Response.objects.filter(response_id=rid).update(**値)
            # 写真の記録（ファイルIDなど）。中身はまだ取らない
            写真新規 = 0
            for rid, 値, 写真ら in 写真の予定:
                resp = Response.objects.get(response_id=rid)
                for qid, order, f in 写真ら:
                    fid = 文字(f.get("fileId"))[:128]
                    項目 = {
                        "name": 文字(f.get("name"))[:255],
                        "mime_type": 文字(f.get("type"))[:64],
                        "captured_at": 文字(f.get("capturedAt"))[:64],
                        "customer_folder_name": 文字(f.get("customerFolderName"))[:255],
                        "folder_name": 文字(f.get("folderName"))[:255],
                        "order": order,
                        "kind": 写真の種別(qid, 値["answers"]),
                    }
                    _, できた = ResponsePhoto.objects.update_or_create(
                        response=resp, question_id=qid[:128], drive_file_id=fid, defaults=項目)
                    写真新規 += 1 if できた else 0

        self.stdout.write("")
        self.stdout.write(f"    → 回答 いま {Response.objects.count()}件 / 写真の記録 {ResponsePhoto.objects.count()}枚（新しく {写真新規}枚）")

        if options["photos"]:
            self._写真を入れる(options["photo_dir"])

    def _写真を入れる(self, フォルダ):
        未 = ResponsePhoto.objects.filter(url="")
        self.stdout.write(f"■ 写真を media へ: まだ入っていない {未.count()}枚")
        入れた, 取れない = 0, []
        for p in 未:
            中身 = _画像を取る(p.drive_file_id, フォルダ)
            if not 中身:
                取れない.append(p.drive_file_id)
                continue
            try:
                p.url = images.保存する(io.BytesIO(中身), "bijiris")
            except Exception as e:  # 画像として読めない
                取れない.append(f"{p.drive_file_id}（{e}）")
                continue
            p.save(update_fields=["url"])
            入れた += 1
        self.stdout.write(f"    入れた: {入れた}枚")
        if 取れない:
            self.stdout.write(f"    取れなかった: {len(取れない)}枚 … {取れない[:5]}{'…' if len(取れない) > 5 else ''}")
            self.stdout.write("    → お客様の写真は Drive で共有されていません（2026-08-25〜）。院長のアカウントで")
            self.stdout.write("      Drive の Bijiris/ からファイルを <ファイルID>.jpg の名前で落とし、--写真フォルダ で入れ直してください。")
