"""ビジリス管理の土台（段取り A）: 表・切り替えの印・取り込み・突き合わせ・仮の画面・左メニュー。

画面の中身（段取り B）はここでは試さない。「準備中」で開くことと、まゆみだけが入れることまで。
"""

import io
import json

import pytest
from django.core.management import call_command

from apps.bijiris import gate, preferences
from apps.bijiris.models import CustomerProfile, Response, ResponsePhoto, Survey
from apps.manage import images
from apps.records.models import AppSetting

pytestmark = pytest.mark.django_db


# ---------------------------------------------------------------------------
# 書き出し JSON の見本（gas/ビジリスを書き出す.js が作る形）
# ---------------------------------------------------------------------------

def _設問(qid, label, type_="text", **extra):
    q = {"id": qid, "label": label, "type": type_, "required": True, "options": [], "placeholder": "",
         "visibilityConditions": [], "visibleWhen": None}
    q.update(extra)
    return q


def _アンケート():
    return {"surveys": [
        {"id": "survey_measurement", "title": "計測時アンケート", "description": "計測", "introMessage": "計測の案内",
         "completionMessage": "ありがとうございました", "status": "published", "sortOrder": 1, "acceptingResponses": True,
         "startAt": "", "endAt": "", "createdAt": "2026-04-01T00:00:00.000Z", "updatedAt": "2026-08-01T00:00:00.000Z",
         "questions": [
             _設問("q_measure_timing", "計測のタイミング", "choice", options=["初回計測時", "回数券終了時", "キャンペーン終了時"]),
             _設問("q_measure_photos", "全身写真", "photo"),
             _設問("q_measure_waist", "ウエスト"),
         ]},
        {"id": "survey_bijiris_session", "title": "施術後アンケート", "description": "施術後", "status": "archived",
         "sortOrder": 0, "questions": [_設問("q_bijiris_session_type", "施術内容", "choice", options=["回数券", "都度"])]},
    ]}


def _写真(fid, name="写真01.jpg"):
    return {"name": name, "type": "image/jpeg", "capturedAt": "2026-08-10T10:00:00+09:00", "fileId": fid,
            "customerFolderName": "佐藤花子", "customerFolderUrl": "https://drive.google.com/x", "folderName": "計測時",
            "folderUrl": "https://drive.google.com/y", "url": "https://drive.google.com/file/d/" + fid,
            "previewUrl": "https://drive.google.com/uc?export=view&id=" + fid,
            "downloadUrl": "https://drive.google.com/uc?export=download&id=" + fid,
            "thumbnailUrl": "https://drive.google.com/thumbnail?id=" + fid}


def _回答():
    answers1 = [{"questionId": "q_measure_timing", "value": "初回計測時"},
                {"questionId": "q_measure_photos", "value": "", "files": [_写真("FILE_A"), _写真("FILE_B", "写真02.jpg")]},
                {"questionId": "q_measure_waist", "value": "70"}]
    answers2 = [{"questionId": "q_measure_timing", "value": "回数券終了時"},
                {"questionId": "q_measure_photos", "value": "", "files": [_写真("FILE_C")]}]
    return {"書き出した日時": "2026-09-16T10:00:00+09:00", "シートの行数": 3, "responses": [
        {"row": 2, "submittedAt": "2026-08-10T10:05:00+09:00", "id": "response_1", "surveyId": "survey_measurement",
         "surveyTitle": "計測時アンケート", "customerClientId": "client-1", "customerName": "佐藤花子",
         "customerEmail": "", "status": "new", "adminMemo": "", "answers": json.dumps(answers1, ensure_ascii=False),
         "files": json.dumps([_写真("FILE_A"), _写真("FILE_B")], ensure_ascii=False), "managedAt": None,
         "surveySheet": {"row": 2, "values": {"回答ID": "response_1", "お名前": "佐藤花子", "ウエスト": "70"}}},
        {"row": 3, "submittedAt": "2026-09-01T10:05:00+09:00", "id": "response_2", "surveyId": "survey_measurement",
         "surveyTitle": "計測時アンケート", "customerClientId": "client-1", "customerName": "佐藤花子",
         "customerEmail": "a@example.com", "status": "trash", "adminMemo": "取り消し", "answers": json.dumps(answers2),
         "files": "[]", "managedAt": "2026-09-02T09:00:00+09:00", "surveySheet": None},
        # 回答JSONが壊れている行。捨てずに文字列のまま持つ
        {"row": 4, "submittedAt": "2026-09-03T10:05:00+09:00", "id": "response_3", "surveyId": "survey_unknown",
         "surveyTitle": "昔のアンケート", "customerClientId": "", "customerName": "山田太郎", "customerEmail": "",
         "status": "何か変", "adminMemo": "", "answers": "{壊れている", "files": "", "managedAt": None, "surveySheet": None},
    ]}


def _顧客():
    return {"profiles": {
        "佐藤花子": {"name": "佐藤花子", "memberNumber": "MYM-0012", "nameKana": "サトウ ハナコ", "aliases": ["佐藤 花子", "????"],
                 "clientIds": ["client-1"], "activeTicketCard": {"plan": "10回券", "sheetNumber": 1, "round": 4},
                 "activeTicketCardSource": "response", "lastTicketCardAcquiredAt": "2026-08-10T01:05:00.000Z",
                 "measurementTargets": {"waist": "65", "hip": "", "thighRight": "", "thighLeft": ""},
                 "ticketStampAdjustment": "2", "pushStatus": {"enabled": True, "supported": True, "permission": "granted"},
                 "rewardRedemptions": {"6": {"handed": True, "handedAt": "2026-08-20T00:00:00.000Z"}},
                 "adminManaged": False, "passcodeHash": "h", "passcodeSalt": "s", "passcodeUpdatedAt": "2026-08-01T00:00:00.000Z",
                 "passcodeSetupUntil": "", "updatedAt": "2026-09-01T00:00:00.000Z"},
        "????": {"name": "????"},
        "山田太郎": {"memberNumber": ""},
    }, "memos": {
        "佐藤花子": {"latestMemo": "古い", "entries": [{"at": "2026-08-01", "memo": "古い"}, {"at": "2026-09-01", "memo": "新しい"}]},
        "山田太郎": "一言だけ",
    }}


def _設定():
    return {"preferences": {"notificationEmail": "Mayumi@Example.com", "backupHour": 30, "retentionDays": -1,
                            "milestoneRewardConfig": {"milestones": [{"threshold": "10", "reward": "施術1回"}, {"threshold": 6, "reward": "お茶"},
                                                                     {"threshold": 0, "reward": "無し"}]},
                            "gachaPrizeConfig": {"monthlyPrizes": [{"month": "2026/9", "prizes": {"A": {"content": "A賞", "probability": 101}}}]},
                            "ANTHROPIC_API_KEY": "sk-secret"},
            "ticket_meta": {"lastRunAt": "2026-09-01"}, "ticket_prompt": "", "ANTHROPIC_API_KEY": "sk-secret"}


def _置く(tmp_path, 名, 中身):
    p = tmp_path / 名
    p.write_text(json.dumps(中身, ensure_ascii=False), encoding="utf-8")
    return str(p)


def _流す(名前, *args):
    out = io.StringIO()
    call_command(名前, *args, stdout=out)
    return out.getvalue()


def _jpeg(色=(200, 100, 50)):
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (2000, 1500), 色).save(buf, format="JPEG")
    return buf.getvalue()


# ---------------------------------------------------------------------------
# 表
# ---------------------------------------------------------------------------

def test_表の保存と便利な項目():
    s = Survey.objects.create(survey_id="survey_x", title="X", questions=[_設問("q1", "写真", "photo"), _設問("q2", "文")])
    r = Response.objects.create(response_id="r1", survey=s, survey_key="survey_x", customer_name="佐藤花子", status="trash",
                                answers=[{"questionId": "q2", "value": "はい"}])
    p = ResponsePhoto.objects.create(response=r, question_id="q1", drive_file_id="F1", kind="before")
    c = CustomerProfile.objects.create(name="佐藤花子", passcode_hash="h", passcode_salt="s",
                                       reward_redemptions={"6": {"handed": True}, "10": {"handed": True}, "x": {}})
    assert s.photo_question_ids == ["q1"] and s.question_count == 2
    assert r.trashed and r.answer_of("q2") == "はい" and r.answer_of("無い") is None and r.member_id == ""
    assert not p.stored and r.photos.count() == 1
    assert c.has_passcode and c.redeemed_thresholds == [6, 10] and c.member_id == ""
    # 回答が消えると写真の記録も消える（写真は回答のもの）
    r.delete()
    assert ResponsePhoto.objects.count() == 0


def test_写真は回答と設問とファイルIDで一意():
    from django.db import IntegrityError

    r = Response.objects.create(response_id="r1")
    ResponsePhoto.objects.create(response=r, question_id="q", drive_file_id="F")
    with pytest.raises(IntegrityError):
        ResponsePhoto.objects.create(response=r, question_id="q", drive_file_id="F")


# ---------------------------------------------------------------------------
# 切り替えの印
# ---------------------------------------------------------------------------

def test_印の切り替えと断る文():
    assert not gate.ビジリスはサーバーが正()
    assert "9/24" in gate.断る文() and "旧ビジリス管理アプリ" in gate.断る文()
    gate.切り替える("server")
    assert gate.ビジリスはサーバーが正()
    assert AppSetting.objects.get(pk="bijiris_source").value == {"source": "server"}
    gate.切り替える("sheet")
    assert not gate.ビジリスはサーバーが正()
    with pytest.raises(ValueError):
        gate.切り替える("drive")
    # 会員の印とは別（片方を立てても、もう片方は立たない）
    from apps.manage import member_gate

    gate.切り替える("server")
    assert not member_gate.会員はサーバーが正()


def test_印のコマンド():
    assert "bijiris_source = sheet" in _流す("ビジリスの正を切り替える")
    assert "bijiris_source = server" in _流す("ビジリスの正を切り替える", "server")
    assert gate.ビジリスはサーバーが正()
    assert "bijiris_source = sheet" in _流す("ビジリスの正を切り替える", "sheet")


# ---------------------------------------------------------------------------
# 取り込み
# ---------------------------------------------------------------------------

def test_アンケートの取り込みは下見で書かず本番で入り2回目で増えない(tmp_path):
    p = _置く(tmp_path, "s.json", _アンケート())
    out = _流す("ビジリスのアンケートを取り込む", p, "--下見")
    assert "新しく入る:   2本" in out and "下見なので" in out and Survey.objects.count() == 0
    out = _流す("ビジリスのアンケートを取り込む", p)
    assert Survey.objects.count() == 2
    s = Survey.objects.get(survey_id="survey_measurement")
    assert s.title == "計測時アンケート" and s.status == "published" and s.sort_order == 1 and s.accepting_responses
    assert s.photo_question_ids == ["q_measure_photos"] and s.created_at.year == 2026 and s.start_at is None
    assert Survey.objects.get(survey_id="survey_bijiris_session").status == "archived"
    out = _流す("ビジリスのアンケートを取り込む", p)
    assert "変わらない:   2本" in out and Survey.objects.count() == 2
    # 中身が変わると更新される。知らない status は published
    中身 = _アンケート()
    中身["surveys"][0]["title"] = "計測時アンケート（改）"
    中身["surveys"][0]["status"] = "変な値"
    _流す("ビジリスのアンケートを取り込む", _置く(tmp_path, "s2.json", 中身))
    s.refresh_from_db()
    assert s.title == "計測時アンケート（改）" and s.status == "published" and Survey.objects.count() == 2


def test_回答の取り込み(tmp_path):
    _流す("ビジリスのアンケートを取り込む", _置く(tmp_path, "s.json", _アンケート()))
    p = _置く(tmp_path, "r.json", _回答())
    out = _流す("ビジリスの回答を取り込む", p, "--下見")
    assert "新しく入る:   3件" in out and Response.objects.count() == 0 and ResponsePhoto.objects.count() == 0
    assert "アンケートの定義が無い回答: survey_unknown" in out and "結びつけていません" in out

    out = _流す("ビジリスの回答を取り込む", p)
    assert Response.objects.count() == 3 and ResponsePhoto.objects.count() == 3
    r1 = Response.objects.get(response_id="response_1")
    assert r1.survey.survey_id == "survey_measurement" and r1.sheet_row == 2 and r1.customer_name == "佐藤花子"
    assert r1.member_id == "" and r1.status == "new" and not r1.trashed and r1.client_id == "client-1"
    assert r1.answer_of("q_measure_waist") == "70" and len(r1.files) == 2 and r1.answers_raw == ""
    assert r1.survey_sheet_row == 2 and r1.survey_sheet_values["ウエスト"] == "70"
    assert r1.submitted_at.isoformat().startswith("2026-08-10T01:05:00+00:00")
    r2 = Response.objects.get(response_id="response_2")
    assert r2.trashed and r2.admin_memo == "取り消し" and r2.managed_at is not None and r2.survey_sheet_values == {}
    r3 = Response.objects.get(response_id="response_3")
    assert r3.survey is None and r3.status == "new" and r3.answers == [] and r3.answers_raw == "{壊れている"

    # 写真の記録: 種別はタイミングの回答で決まる。Drive のリンクは持たない
    a = ResponsePhoto.objects.get(drive_file_id="FILE_A")
    assert a.kind == "before" and a.question_id == "q_measure_photos" and a.order == 0 and a.name == "写真01.jpg"
    assert a.folder_name == "計測時" and a.customer_folder_name == "佐藤花子" and a.url == ""
    assert ResponsePhoto.objects.get(drive_file_id="FILE_B").order == 1
    assert ResponsePhoto.objects.get(drive_file_id="FILE_C").kind == "after"
    assert not any(hasattr(a, f) for f in ("preview_url", "download_url", "thumbnail_url", "folder_url"))

    # 2回目で増えない
    out = _流す("ビジリスの回答を取り込む", p)
    assert "変わらない:   3件" in out and Response.objects.count() == 3 and ResponsePhoto.objects.count() == 3
    # 対応状況が変わると更新される
    中身 = _回答()
    中身["responses"][0]["status"] = "done"
    _流す("ビジリスの回答を取り込む", _置く(tmp_path, "r2.json", 中身))
    r1.refresh_from_db()
    assert r1.status == "done" and Response.objects.count() == 3


def test_回答の写真は種別が旧設問でも決まる():
    from apps.bijiris.management.commands.ビジリスの回答を取り込む import 写真の種別

    assert 写真の種別("q_bijiris_session_monitor_photos_6", []) == "before"
    assert 写真の種別("q_ticket_end_photo_last", []) == "after"
    assert 写真の種別("q_measure_photos", [{"questionId": "q_measure_timing", "value": ["キャンペーン終了時"]}]) == "after"
    assert 写真の種別("q_measure_photos", [{"questionId": "q_measure_timing", "value": "モニター時"}]) == "before"
    assert 写真の種別("q_measure_photos", []) == "other"
    assert 写真の種別("q_other", []) == "other"


def test_写真をフォルダから入れる(tmp_path, settings):
    _流す("ビジリスの回答を取り込む", _置く(tmp_path, "r.json", _回答()))
    フォルダ = tmp_path / "photos"
    フォルダ.mkdir()
    (フォルダ / "FILE_A.jpg").write_bytes(_jpeg())
    (フォルダ / "FILE_B.jpg").write_bytes(_jpeg((10, 20, 30)))
    # FILE_C はフォルダに無く、Drive にも取りに行かせない（ネットへ出ない）
    import urllib.request

    def 出ない(*a, **k):
        raise OSError("ネットへ出ない")

    settings_urlopen = urllib.request.urlopen
    urllib.request.urlopen = 出ない
    try:
        out = _流す("ビジリスの回答を取り込む", _置く(tmp_path, "r.json", _回答()), "--写真", "--写真フォルダ", str(フォルダ))
    finally:
        urllib.request.urlopen = settings_urlopen
    assert "入れた: 2枚" in out and "取れなかった: 1枚" in out and "--写真フォルダ" in out
    a = ResponsePhoto.objects.get(drive_file_id="FILE_A")
    assert a.url.startswith(settings.PUBLIC_BASE_URL + "/media/bijiris/") and a.url.endswith(".jpg")
    名前 = a.url.rsplit("/", 1)[1]
    assert (settings.MEDIA_ROOT / "bijiris" / 名前).is_file()
    assert ResponsePhoto.objects.get(drive_file_id="FILE_C").url == ""
    # 縮めて保存されている（MAX_EDGE）
    from PIL import Image

    im = Image.open(settings.MEDIA_ROOT / "bijiris" / 名前)
    assert max(im.size) == images.MAX_EDGE
    # もう一度 --写真 でも、入っているものは取りに行かない
    urllib.request.urlopen = 出ない
    try:
        out = _流す("ビジリスの回答を取り込む", _置く(tmp_path, "r.json", _回答()), "--写真", "--写真フォルダ", str(フォルダ))
    finally:
        urllib.request.urlopen = settings_urlopen
    assert "まだ入っていない 1枚" in out and "入れた: 0枚" in out
    a.refresh_from_db()
    assert a.url.endswith(名前)


def test_写真の配り口(client, tmp_path, settings):
    url = images.保存する(io.BytesIO(_jpeg()), "bijiris")
    名前 = url.rsplit("/", 1)[1]
    r = client.get(f"/media/bijiris/{名前}")
    assert r.status_code == 200 and r["Content-Type"] == "image/jpeg"
    assert client.get("/media/bijiris/nothing.jpg").status_code == 404


def test_顧客の取り込み(tmp_path):
    from apps.members.models import Member

    Member.objects.create(member_id="MYM-0012", name="佐藤花子")
    p = _置く(tmp_path, "c.json", _顧客())
    out = _流す("ビジリスの顧客を取り込む", p, "--下見")
    assert "新しく入る:   2名" in out and "お名前になっていない鍵" in out and CustomerProfile.objects.count() == 0
    assert "ちょうど1名と一致 1名" in out and "結びつけていません" in out
    _流す("ビジリスの顧客を取り込む", p)
    assert CustomerProfile.objects.count() == 2
    c = CustomerProfile.objects.get(name="佐藤花子")
    # **一致していても member_id は空。**GAS の会員番号は写すだけ
    assert c.member_id == "" and c.member_number == "MYM-0012"
    assert c.name_kana == "サトウハナコ" and c.aliases == ["佐藤 花子"] and c.client_ids == ["client-1"]
    assert c.active_ticket_card == {"plan": "10回券", "sheetNumber": 1, "round": 4} and c.active_ticket_card_source == "response"
    assert c.last_ticket_card_acquired_at is not None and c.ticket_stamp_adjustment == 2
    assert c.reward_redemptions == {"6": {"handed": True, "handedAt": "2026-08-20T00:00:00.000Z"}} and c.redeemed_thresholds == [6]
    assert c.has_passcode and c.passcode_updated_at == "2026-08-01T00:00:00.000Z" and c.updated_at is not None
    # メモは新しい順。最新のメモは先頭のもの
    assert c.latest_memo == "新しい" and [e["at"] for e in c.memo_entries] == ["2026-09-01", "2026-08-01"]
    t = CustomerProfile.objects.get(name="山田太郎")
    assert not t.has_passcode and t.latest_memo == "一言だけ" and t.memo_entries == [{"at": "", "memo": "一言だけ"}]
    out = _流す("ビジリスの顧客を取り込む", p)
    assert "変わらない:   2名" in out and CustomerProfile.objects.count() == 2


def test_設定の取り込み(tmp_path):
    p = _置く(tmp_path, "p.json", _設定())
    out = _流す("ビジリスの設定を取り込む", p, "--下見")
    assert "秘密が混ざっていたので捨てました: ANTHROPIC_API_KEY" in out and AppSetting.objects.count() == 0
    _流す("ビジリスの設定を取り込む", p)
    v = AppSetting.objects.get(pk="bijiris_preferences").value
    assert "ANTHROPIC_API_KEY" not in v and json.dumps(v).find("sk-secret") == -1
    # getPreferences_ と同じ既定と正規化
    assert v["notificationEmail"] == "mayumi@example.com" and v["backupHour"] == 3 and v["retentionDays"] == 365
    assert v["notificationSubject"] == preferences.既定の通知件名 and v["consentText"] == preferences.既定の同意文
    assert v["bijirisCategoryConfig"]["concernPaths"] == preferences.既定のお悩みの道 and v["twoFactorEnabled"] is False
    assert v["milestoneRewardConfig"] == {"enabled": True, "milestones": [
        {"threshold": 6, "reward": "お茶", "description": ""}, {"threshold": 10, "reward": "施術1回", "description": ""}]}
    月 = v["gachaPrizeConfig"]["monthlyPrizes"]
    assert len(月) == 1 and 月[0]["month"] == "2026-09" and 月[0]["prizes"]["A"] == {"content": "A賞", "probability": 100}
    assert 月[0]["prizes"]["D"] == {"content": "", "probability": 0}
    assert AppSetting.objects.get(pk="bijiris_ticket_meta").value == {"lastRunAt": "2026-09-01"}
    assert AppSetting.objects.get(pk="bijiris_ticket_prompt").value == {"prompt": ""}
    assert preferences.分析指示文を読む() == preferences.既定の分析指示文
    assert preferences.読む()["notificationEmail"] == "mayumi@example.com"
    out = _流す("ビジリスの設定を取り込む", p)
    assert out.count("変わらない") == 3


def test_設定の既定値はGASと同じ形():
    v = preferences.既定値()
    assert set(v) == {"notificationEnabled", "notificationEmail", "notificationSubject", "notificationBody", "dataPolicyText",
                      "requireConsent", "consentText", "autoBackupEnabled", "backupHour", "retentionDays", "recoveryMemo",
                      "twoFactorEnabled", "bijirisCategoryConfig", "gachaPrizeConfig", "milestoneRewardConfig", "campaignStampEnabled"}
    assert v["milestoneRewardConfig"] == {"enabled": True, "milestones": []}
    assert len(v["gachaPrizeConfig"]["monthlyPrizes"]) == 1 and v["gachaPrizeConfig"]["monthlyPrizes"][0]["prizes"]["A"]["probability"] == 25
    assert v["bijirisCategoryConfig"]["generalCategories"] == preferences.既定の一般カテゴリ
    # 保存先の文言は取り除く（normalizeDataPolicyText_）
    assert preferences.利用目的文を整える("利用します。保存先は Google スプレッドシートおよび Google ドライブです。") == "利用します。"
    assert preferences.月の鍵("2026年9月") == "2026-09" and preferences.月の鍵("9月") == ""


def test_突き合わせは読むだけ(tmp_path):
    p = _置く(tmp_path, "r.json", _回答())
    out = _流す("ビジリスを突き合わせる", p)
    assert "JSON 3件 / サーバー 0件" in out and "JSON にだけある: 3件" in out and Response.objects.count() == 0
    _流す("ビジリスの回答を取り込む", p)
    out = _流す("ビジリスを突き合わせる", p)
    assert "同じ: 3件" in out and "違う: 0件" in out and "サーバーにだけある: 0件" in out
    # サーバー側だけ変えると違いとして出る
    Response.objects.filter(response_id="response_1").update(admin_memo="受付で直した")
    Response.objects.create(response_id="response_9")
    ResponsePhoto.objects.filter(drive_file_id="FILE_B").delete()
    out = _流す("ビジリスを突き合わせる", p)
    assert "違う: 1件" in out and "response_1: admin_memo・写真の枚数" in out and "サーバーにだけある: 1件" in out
    assert Response.objects.get(response_id="response_1").admin_memo == "受付で直した"


# ---------------------------------------------------------------------------
# 仮の画面と左メニュー
# ---------------------------------------------------------------------------

URLS = ["/manage/bijiris/", "/manage/bijiris/surveys/", "/manage/bijiris/responses/", "/manage/bijiris/customers/",
        "/manage/bijiris/rewards/", "/manage/bijiris/tickets/"]


def test_6つの画面が準備中で開く(as_owner):
    Survey.objects.create(survey_id="s", title="X")
    for url, 題 in zip(URLS, ["集計", "アンケート管理", "回答管理", "顧客管理", "特典", "回数券分析"]):
        page = as_owner.get(url)
        assert page.status_code == 200, url
        html = page.content.decode()
        # 中身が入った画面（段取り B）は test_bijiris_<担当>.py で試す
        if 題 in ("顧客管理", "特典"):
            assert "準備中" not in html and gate.断る文() in html
            continue
        assert f"<h1>{題}</h1>" in html and "準備中" in html and "スプレッドシートが正" in html and gate.断る文() in html
        assert "<td>アンケート</td><td>1件</td>" in html and "<td>回数券分析</td><td>0件</td>" in html
    gate.切り替える("server")
    html = as_owner.get("/manage/bijiris/").content.decode()
    assert "サーバーが正" in html and gate.断る文() not in html


def test_スタッフは403で未ログインはログインへ(client, staff):
    r = client.get("/manage/bijiris/")
    assert r.status_code == 302 and r["Location"].startswith("/manage/login/")
    client.force_login(staff)
    for url in URLS:
        assert client.get(url).status_code == 403, url


def test_左メニューにビジリスの段が出る(as_owner):
    from django.urls import reverse

    assert reverse("manage:bijiris:dashboard") == "/manage/bijiris/"
    page = as_owner.get("/manage/bijiris/responses/").content.decode()
    side = page[page.find('<nav class="sidebar-nav">'):page.find("</nav>")]
    assert ">ビジリス</span>" in side
    assert 'class="active"><span class="nav-icon">💪</span> ビジリス管理</a>' in side
    # 段の並び: 公式サイトの次、開発の前
    assert side.find(">公式サイト</span>") < side.find(">ビジリス</span>") < side.find(">開発</span>")
    i = page.find('<nav class="page-tabs"')
    tabs = page[i:page.find("</nav>", i)]
    for t in ["集計", "アンケート管理", "回答管理", "顧客管理", "特典", "回数券分析"]:
        assert t in tabs, t
    assert 'class="active">回答管理</a>' in tabs
    # アプリ管理側の reward_/ticket などと頭がかぶらない
    assert 'class="active"><span class="nav-icon">👥</span> 会員</a>' not in side
    page = as_owner.get("/manage/rewards/").content.decode()
    assert 'class="active"><span class="nav-icon">💪</span> ビジリス管理</a>' not in page


def test_スタッフの左メニューにビジリスは出ない(rf, staff):
    from apps.manage import navigation

    req = rf.get("/manage/")
    req.user = staff
    req.resolver_match = None
    assert [s["label"] for s in navigation.組み立てる(req)["nav_sections"]] == ["予約管理"]
