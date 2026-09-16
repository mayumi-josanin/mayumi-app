"""ビジリス管理「回数券分析」。**中身は旧ビジリス管理アプリの #page-ticket-survey と同じ**
（自動処理・モニター基準画像・分析の設定・お名前と分析状態の絞り込み・お客様ごとの写真と分析文）。

AI（Claude）への問い合わせは urllib を差し替えて試す。**本物の API には出ない。**
"""

import io
import json
import urllib.error
from datetime import date, datetime, timezone

import pytest

from apps.bijiris import gate, preferences, views_ticket
from apps.bijiris.models import Response, ResponsePhoto, Survey
from apps.manage import images
from apps.measurements.models import Measurement
from apps.records.models import AppSetting, TicketAnalysis

pytestmark = pytest.mark.django_db

一覧 = "/manage/bijiris/tickets/"


def _時(y, m, d, h=10):
    return datetime(y, m, d, h, 0, tzinfo=timezone.utc)


def _jpeg(色=(200, 100, 50)):
    from PIL import Image

    buf = io.BytesIO()
    Image.new("RGB", (400, 300), 色).save(buf, format="JPEG")
    return buf.getvalue()


def _写真(response, qid, fid, kind, 取り込む=True, name=""):
    url = images.保存する(io.BytesIO(_jpeg()), "bijiris") if 取り込む else ""
    return ResponsePhoto.objects.create(response=response, question_id=qid, drive_file_id=fid, kind=kind,
                                        url=url, name=name or f"{fid}.jpg")


@pytest.fixture
def 下ごしらえ():
    """佐藤花子: 初回計測（ビフォー2枚・ウエスト70）と回数券終了時（アフター1枚・ウエスト65）。
    山田太郎: 回数券終了時だけ（写真は未取り込み）。鈴木: ゴミ箱。"""
    s = Survey.objects.create(survey_id="survey_measurement", title="計測時アンケート", questions=[
        {"id": "q_measure_timing", "label": "計測のタイミング", "type": "choice"},
        {"id": "q_measure_photos", "label": "全身写真", "type": "photo"},
        {"id": "q_measure_waist", "label": "ウエスト", "type": "text"},
        {"id": "q_measure_improve", "label": "今後もっと改善したい部分はありますか？", "type": "checkbox"},
    ])
    r1 = Response.objects.create(response_id="resp_monitor", survey=s, survey_key=s.survey_id, customer_name="佐藤 花子",
                                 submitted_at=_時(2026, 6, 1), answers=[
                                     {"questionId": "q_measure_timing", "value": "初回計測時"},
                                     {"questionId": "q_measure_photos", "value": "", "files": [{"fileId": "B1"}, {"fileId": "B2"}]},
                                     {"questionId": "q_measure_waist", "value": "70"}])
    _写真(r1, "q_measure_photos", "B1", "before")
    _写真(r1, "q_measure_photos", "B2", "before")
    r2 = Response.objects.create(response_id="resp_after", survey=s, survey_key=s.survey_id, customer_name="佐藤花子",
                                 submitted_at=_時(2026, 9, 1), answers=[
                                     {"questionId": "q_measure_timing", "value": "回数券終了時"},
                                     {"questionId": "q_measure_photos", "value": "", "files": [{"fileId": "A1"}]},
                                     {"questionId": "q_measure_waist", "value": "65"},
                                     {"questionId": "q_measure_improve", "value": ["お腹", "姿勢"]}])
    _写真(r2, "q_measure_photos", "A1", "after")
    r3 = Response.objects.create(response_id="resp_yamada", survey=s, survey_key=s.survey_id, customer_name="山田太郎",
                                 submitted_at=_時(2026, 8, 15), answers=[{"questionId": "q_measure_timing", "value": "キャンペーン終了時"}])
    _写真(r3, "q_measure_photos", "Y1", "after", 取り込む=False)
    r4 = Response.objects.create(response_id="resp_trash", survey=s, survey_key=s.survey_id, customer_name="鈴木",
                                 submitted_at=_時(2026, 8, 20), status="trash")
    _写真(r4, "q_measure_photos", "T1", "after")
    Measurement.objects.create(measurement_id="m1", customer_name="佐藤花子", measured_on=date(2026, 5, 1), waist=72, hip=95)
    Measurement.objects.create(measurement_id="m2", customer_name="佐藤花子", measured_on=date(2026, 8, 1), waist=66.5, hip=92)
    return {"monitor": r1, "after": r2, "yamada": r3}


class _応答(io.BytesIO):
    status = 200

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


@pytest.fixture
def fake_claude(monkeypatch, settings):
    """Claude API の代わり。送った要求を貯め、決まった分析文を返す。"""
    settings.ANTHROPIC_API_KEY = "sk-ant-test"

    class _送った(list):
        答 = None  # 返す中身を試験の途中で差し替えるための入れ物

    sent = _送った()
    答 = {"text": "■ 変化のポイント\n姿勢が整いました。"}

    def fake_urlopen(req, timeout=0):
        sent.append({"url": req.full_url, "headers": dict(req.header_items()), "body": json.loads(req.data)})
        if 答.get("http_error"):
            raise urllib.error.HTTPError(req.full_url, 400, "Bad Request", {}, io.BytesIO(
                json.dumps({"error": {"message": "invalid image"}}).encode()))
        if 答.get("refusal"):
            return _応答(json.dumps({"stop_reason": "refusal", "content": []}).encode())
        return _応答(json.dumps({"stop_reason": "end_turn", "content": [
            {"type": "thinking", "thinking": ""}, {"type": "text", "text": 答["text"]}]}).encode())

    monkeypatch.setattr(views_ticket.urllib.request, "urlopen", fake_urlopen)
    sent.答 = 答
    return sent


# ---------------------------------------------------------------------------
# 一覧
# ---------------------------------------------------------------------------

def test_一覧はアフター写真のある回答だけで新しい順(as_owner, 下ごしらえ):
    page = as_owner.get(一覧)
    assert page.status_code == 200
    html = page.content.decode()
    assert "<h1>回数券分析</h1>" in html
    # 旧アプリのカード: 自動処理・モニター基準画像・分析の設定・絞り込み
    for t in ["自動処理", "自動で取り込み・分析する", "OFF：手動で「写真を取り込む」「分析する」を実行します。",
              "モニター基準画像の一括取り込み（初回のみ）", "既存のモニター画像（計測写真(1回目)）を josanin の Drive に一度だけコピーします。",
              "分析の設定", "分析モデル: claude-opus-5", "分析プロンプト（この指示文をもとに写真を比較します）", "初期値に戻す", "プロンプトを保存",
              "Claude API キー", "未設定です。分析するには Claude API キーが必要です", "自動処理には Claude API キーの設定が必要です。",
              "お名前で絞り込み", "分析状態", "未分析をまとめて分析"]:
        assert t in html, t
    # 鍵は出ない・入力欄も無い
    assert "sk-ant" not in html and 'type="password"' not in html
    # 既定の指示文が入っている
    assert preferences.既定の分析指示文.splitlines()[0] in html
    # 対象: 佐藤花子（アフターあり）と山田太郎（アフターあり・未取り込み）。初回計測・ゴミ箱は出ない
    entries = views_ticket.一覧を作る()
    assert [e["id"] for e in entries] == ["resp_after", "resp_yamada"]
    e = entries[0]
    assert e["submitted_date"] == "2026-09-01" and e["status"] == "none" and e["status_label"] == "分析待ち"
    # ビフォーは同じお名前（空白の違いは無視）の初回計測の写真、アフターは自分の写真
    assert [p["file_id"] for p in e["before_photos"]] == ["B1", "B2"] and [p["file_id"] for p in e["after_photos"]] == ["A1"]
    assert all(p["stored"] for p in e["before_photos"] + e["after_photos"])
    assert e["measurements"]["before"]["waist"] == "70" and e["measurements"]["after"]["waist"] == "65"
    assert e["measurements"]["before_date"] == "2026-06-01" and e["measurements"]["after_date"] == "2026-09-01"
    y = entries[1]
    assert y["before_photos"] == [] and y["after_photos"][0]["stored"] is False
    assert "モニター時（ビフォー）（2枚）" in html and "回数券終了時（アフター）（1枚）" in html
    assert "写真なし" in html and "未取り込み" in html and "まだ分析していません。" in html
    assert "提出日 2026-09-01" in html and "2 / 2 件を表示" in html and "分析する</button>" in html
    assert "ウエスト 70 → 65" in html
    assert "鈴木" not in html


def test_分析結果は回答IDで結び写真はサーバーのものに当てる(as_owner, 下ごしらえ):
    TicketAnalysis.objects.create(sheet_row=2, response_id="resp_after", customer_name="佐藤花子", status="done",
                                  result="■ 変化のポイント\n姿勢がよくなりました。", analyzed_at=_時(2026, 9, 2, 3),
                                  # ビフォーは Drive の「モニター写真」フォルダの写し（サーバーに無い ID）
                                  before_photos=[{"fileId": "COPY_1", "name": "monitor_01.jpg"}],
                                  after_photos=[{"fileId": "A1", "name": "after.jpg"}, {"fileId": "NOPE", "name": "x.jpg"}])
    TicketAnalysis.objects.create(sheet_row=3, response_id="resp_yamada", customer_name="山田太郎", status="error",
                                  error="アフター写真（回数券終了時の写真）がありません。")
    entries = views_ticket.一覧を作る()
    e = entries[0]
    assert e["status"] == "done" and e["analysis_text"].startswith("■") and e["analyzed_at"] is not None
    # 1枚も当たらないビフォーは回答の写真に戻る。アフターは当たった分と「未取り込み」
    assert [p["file_id"] for p in e["before_photos"]] == ["B1", "B2"]
    assert [(p["file_id"], p["stored"]) for p in e["after_photos"]] == [("A1", True), ("NOPE", False)]
    html = as_owner.get(一覧).content.decode()
    assert "分析完了" in html and "姿勢がよくなりました。" in html and "再分析する" in html and "分析結果をコピー" in html
    assert "分析エラー" in html and "アフター写真（回数券終了時の写真）がありません。" in html
    assert "・分析 2026/09/02" in html


def test_お名前と分析状態で絞れる(as_owner, 下ごしらえ):
    TicketAnalysis.objects.create(sheet_row=2, response_id="resp_after", status="done", result="済")
    html = as_owner.get(一覧 + "?q=山田").content.decode()
    assert "山田太郎" in html and "佐藤花子" not in html and "1 / 2 件を表示" in html
    html = as_owner.get(一覧 + "?status=done").content.decode()
    assert "佐藤花子" in html and "山田太郎" not in html
    html = as_owner.get(一覧 + "?status=error").content.decode()
    assert "条件に合うアンケートがありません。" in html
    # 知らない状態は「すべて」
    assert "2 / 2 件を表示" in as_owner.get(一覧 + "?status=zzz").content.decode()


def test_回答が無ければ旧アプリと同じ空の文言(as_owner):
    html = as_owner.get(一覧).content.decode()
    assert "まだ取り込んでいません。「写真を取り込む」を押してください。" in html


def test_進行状況と鍵の有無が出る(as_owner, settings):
    AppSetting.objects.create(key="bijiris_ticket_meta", value={
        "autoEnabled": True, "lastAutoRunAt": "2026-09-10T01:00:00.000Z", "autoError": "APIキー未設定のためスキップ",
        "monitorSeededAt": "2026-05-01T00:00:00.000Z", "monitorSeedSummary": {"folders": 11, "copied": 30, "skipped": 2}})
    settings.ANTHROPIC_API_KEY = "sk-ant-secret"
    html = as_owner.get(一覧).content.decode()
    assert "ON：約30分ごとに新しい回答を自動で取り込み・分析します。" in html and "checked" in html
    assert "最終自動実行 2026-09-10T01:00:00.000Z" in html and "APIキー未設定のためスキップ" in html
    assert "取り込み済みです（最終 2026-05-01T00:00:00.000Z）" in html
    assert "取り込み済み: 11 名分 / 画像 30 枚 / スキップ 2 名（既に取り込み済み）" in html
    assert "設定済みです。" in html and "sk-ant-secret" not in html


def test_保存した指示文が出る(as_owner):
    AppSetting.objects.create(key="bijiris_ticket_prompt", value={"prompt": "{{お名前}}さんの写真を比べてください。"})
    html = as_owner.get(一覧).content.decode()
    assert "{{お名前}}さんの写真を比べてください。" in html


def test_スタッフは403(client, staff):
    client.force_login(staff)
    assert client.get(一覧).status_code == 403
    assert client.post(一覧 + "prompt/", {"prompt": "x"}).status_code == 403


# ---------------------------------------------------------------------------
# 書き込み（印が立つまで断る）
# ---------------------------------------------------------------------------

def test_印が立つまでは書き込みを断る(as_owner, 下ごしらえ, settings):
    settings.ANTHROPIC_API_KEY = "sk-ant-test"
    for path, data in [("prompt/", {"prompt": "x"}), ("auto/", {"enabled": "1"}), ("analyze/", {"entry_id": "resp_after"})]:
        r = as_owner.post(一覧 + path, data, follow=True)
        assert gate.断る文() in r.content.decode(), path
    assert AppSetting.objects.filter(pk="bijiris_ticket_prompt").count() == 0
    assert AppSetting.objects.filter(pk="bijiris_ticket_meta").count() == 0
    assert TicketAnalysis.objects.count() == 0
    assert "見るだけならこの画面でできます" in as_owner.get(一覧).content.decode()


def test_指示文の保存と空なら既定に戻る(as_owner):
    gate.切り替える("server")
    r = as_owner.post(一覧 + "prompt/", {"prompt": "  {{お名前}}さんを比べる  ", "next": 一覧 + "?q=a"}, follow=True)
    assert r.redirect_chain[-1][0] == 一覧 + "?q=a" and "プロンプトを保存しました。" in r.content.decode()
    assert AppSetting.objects.get(pk="bijiris_ticket_prompt").value == {"prompt": "{{お名前}}さんを比べる"}
    assert preferences.分析指示文を読む() == "{{お名前}}さんを比べる"
    as_owner.post(一覧 + "prompt/", {"prompt": ""})
    assert AppSetting.objects.get(pk="bijiris_ticket_prompt").value == {"prompt": ""}
    assert preferences.分析指示文を読む() == preferences.既定の分析指示文
    # この画面の外へは戻さない
    r = as_owner.post(一覧 + "prompt/", {"prompt": "x", "next": "https://evil.example/"})
    assert r["Location"] == 一覧


def test_自動処理の切り替え(as_owner, settings):
    gate.切り替える("server")
    AppSetting.objects.create(key="bijiris_ticket_meta", value={"lastAutoRunAt": "2026-09-10"})
    r = as_owner.post(一覧 + "auto/", {"enabled": "1"}, follow=True)
    assert "先に AI の鍵（ANTHROPIC_API_KEY）をサーバーの設定に入れてください。" in r.content.decode()
    settings.ANTHROPIC_API_KEY = "sk-ant-test"
    r = as_owner.post(一覧 + "auto/", {"enabled": "1"}, follow=True)
    assert "自動処理を有効にしました。" in r.content.decode()
    # 他の項目は残る
    assert AppSetting.objects.get(pk="bijiris_ticket_meta").value == {"lastAutoRunAt": "2026-09-10", "autoEnabled": True}
    r = as_owner.post(一覧 + "auto/", {"enabled": "0"}, follow=True)
    assert "自動処理を停止しました。" in r.content.decode()
    assert AppSetting.objects.get(pk="bijiris_ticket_meta").value["autoEnabled"] is False


# ---------------------------------------------------------------------------
# 分析
# ---------------------------------------------------------------------------

def test_鍵が無ければ分析を断る(as_owner, 下ごしらえ, settings):
    gate.切り替える("server")
    settings.ANTHROPIC_API_KEY = ""
    r = as_owner.post(一覧 + "analyze/", {"entry_id": "resp_after"}, follow=True)
    assert "AI の鍵が設定されていません" in r.content.decode() and TicketAnalysis.objects.count() == 0
    settings.ANTHROPIC_API_KEY = "sk-ant-test"
    r = as_owner.post(一覧 + "analyze/", {}, follow=True)
    assert "分析するアンケートがありません。" in r.content.decode()


def test_分析はCode_gsと同じ要求を送り結果を記録する(as_owner, 下ごしらえ, fake_claude):
    gate.切り替える("server")
    AppSetting.objects.create(key="bijiris_ticket_prompt", value={"prompt": "\n".join([
        "{{お名前}}様（ビフォー {{ビフォー日付}} {{ビフォー枚数}}枚 → アフター {{アフター日付}} {{アフター枚数}}枚）",
        "ウエスト {{初回ウエスト}}→{{今回ウエスト}} ヒップ {{初回ヒップ}}→{{今回ヒップ}} 太もも右 {{初回太もも右}}→{{今回太もも右}}",
        "改善したい: {{今後もっと改善したい部分はありますか？}}"])})
    r = as_owner.post(一覧 + "analyze/", {"entry_id": "resp_after", "next": 一覧 + "#entry-resp_after"}, follow=True)
    assert "分析が完了しました。" in r.content.decode()

    assert len(fake_claude) == 1
    req = fake_claude[0]
    assert req["url"] == "https://api.anthropic.com/v1/messages"
    assert req["headers"]["X-api-key"] == "sk-ant-test" and req["headers"]["Anthropic-version"] == "2023-06-01"
    body = req["body"]
    assert body["model"] == "claude-opus-5" and body["max_tokens"] == 3000
    assert body["thinking"] == {"type": "disabled"} and body["output_config"] == {"effort": "medium"}
    assert body["system"].startswith("あなたは日本語で回答するアシスタントです。")
    content = body["messages"][0]["content"]
    assert body["messages"][0]["role"] == "user"
    種類 = [c["type"] for c in content]
    # 前置き → ビフォーの見出し → ビフォー2枚 → アフターの見出し → アフター1枚 → 指示文
    assert 種類 == ["text", "text", "image", "image", "text", "image", "text"]
    前置き = content[0]["text"]
    assert 前置き.startswith("【お客様】佐藤花子\n【提出日】2026-09-01\n")
    assert "【計測のタイミング】回数券終了時" in 前置き and "【ウエスト】65" in 前置き
    assert "【今後もっと改善したい部分はありますか？】お腹、姿勢" in 前置き and "【全身写真】" not in 前置き
    assert "【計測記録（参考）】\n2026-05-01 / ウエスト 72 / ヒップ 95 / 太もも右 - / 太もも左 -\n2026-08-01 / ウエスト 66.5 / ヒップ 92" in 前置き
    assert content[1]["text"] == "以下は【初回計測時（ビフォー）】の写真です。"
    assert content[4]["text"] == "以下は【回数券終了時（アフター）】の写真です。"
    img = content[2]["source"]
    assert img["type"] == "base64" and img["media_type"] == "image/jpeg" and img["data"].startswith("/9j/")
    assert content[6]["text"] == "\n".join([
        "佐藤花子様（ビフォー 2026-06-01 2枚 → アフター 2026-09-01 1枚）",
        "ウエスト 70→65 ヒップ -→- 太もも右 -→-",
        "改善したい: お腹、姿勢"])

    t = TicketAnalysis.objects.get(response_id="resp_after")
    assert t.status == "done" and t.result == "■ 変化のポイント\n姿勢が整いました。" and t.error == ""
    assert t.analyzed_at is not None and t.customer_name == "佐藤花子" and t.sheet_row == 2 and t.member_id == ""
    assert t.before_photos == [{"fileId": "B1", "name": "B1.jpg"}, {"fileId": "B2", "name": "B2.jpg"}]
    assert t.after_photos == [{"fileId": "A1", "name": "A1.jpg"}]
    assert not json.dumps(t.before_photos).count("drive.google.com")
    html = as_owner.get(一覧).content.decode()
    assert "分析完了" in html and "姿勢が整いました。" in html and "再分析する" in html

    # 再分析は同じ行を書き換える（行は増えない）
    fake_claude.答["text"] = "二度目"
    as_owner.post(一覧 + "analyze/", {"entry_id": "resp_after"})
    assert TicketAnalysis.objects.count() == 1 and TicketAnalysis.objects.get().result == "二度目"


def test_初回の計測値はモニター回答が無ければ計測記録のいちばん古いものから(下ごしらえ):
    Response.objects.filter(response_id="resp_monitor").delete()
    e = [x for x in views_ticket.一覧を作る() if x["id"] == "resp_after"][0]
    assert e["before_photos"] == []
    assert e["measurements"]["before"] == {"waist": "72", "hip": "95", "thigh_right": "", "thigh_left": ""}
    assert e["measurements"]["before_date"] == "2026-05-01"
    text = views_ticket._指示文を埋める("{{初回ウエスト}}/{{ビフォー日付}}/{{初回太もも左}}", e["response"], [], [])
    assert text == "72/2026-05-01/-"
    # 指示文に {{ が無ければそのまま
    assert views_ticket._指示文を埋める("そのまま", e["response"], [], []) == "そのまま"


def test_ビフォー写真が無くても分析できる(as_owner, 下ごしらえ, fake_claude):
    gate.切り替える("server")
    Response.objects.filter(response_id="resp_monitor").delete()
    as_owner.post(一覧 + "analyze/", {"entry_id": "resp_after"})
    content = fake_claude[0]["body"]["messages"][0]["content"]
    assert [c["type"] for c in content] == ["text", "text", "text", "image", "text"]
    assert content[1]["text"] == "【初回計測時（ビフォー）】の写真はありません。"
    assert TicketAnalysis.objects.get(response_id="resp_after").before_photos == []


def test_写真が未取り込みならエラーとして記録し次の手を書く(as_owner, 下ごしらえ, fake_claude):
    gate.切り替える("server")
    r = as_owner.post(一覧 + "analyze/", {"entry_id": "resp_yamada"}, follow=True)
    assert "分析完了：成功 0 件 / 失敗 1 件" in r.content.decode()
    t = TicketAnalysis.objects.get(response_id="resp_yamada")
    assert t.status == "error" and "まだサーバーに取り込まれていません" in t.error and "--写真" in t.error
    assert fake_claude == []  # AI には送っていない
    assert "分析エラー" in as_owner.get(一覧).content.decode()


def test_APIのエラーと拒否はエラーとして記録する(as_owner, 下ごしらえ, fake_claude):
    gate.切り替える("server")
    fake_claude.答["http_error"] = True
    as_owner.post(一覧 + "analyze/", {"entry_id": "resp_after"})
    t = TicketAnalysis.objects.get(response_id="resp_after")
    assert t.status == "error" and t.error == "Claude API エラー (400): invalid image"
    fake_claude.答.pop("http_error")
    fake_claude.答["refusal"] = True
    as_owner.post(一覧 + "analyze/", {"entry_id": "resp_after"})
    t.refresh_from_db()
    assert t.status == "error" and "拒否しました" in t.error
    # 届かないとき
    fake_claude.答.pop("refusal")
    fake_claude.答["text"] = ""
    as_owner.post(一覧 + "analyze/", {"entry_id": "resp_after"})
    t.refresh_from_db()
    assert t.status == "error" and t.error == "分析結果が空でした。もう一度お試しください。"


def test_届かないときはエラー(monkeypatch):
    def 落ちる(req, timeout=0):
        raise urllib.error.URLError("no route")

    monkeypatch.setattr(views_ticket.urllib.request, "urlopen", 落ちる)
    with pytest.raises(RuntimeError, match="届きませんでした"):
        views_ticket.AIに問い合わせる("k", [{"type": "text", "text": "x"}])


def test_まとめて分析は5件までで残りは知らせる(as_owner, 下ごしらえ, fake_claude):
    gate.切り替える("server")
    s = Survey.objects.get()
    for i in range(6):
        r = Response.objects.create(response_id=f"extra_{i}", survey=s, survey_key=s.survey_id, customer_name=f"客{i}",
                                    submitted_at=_時(2026, 7, i + 1),
                                    answers=[{"questionId": "q_measure_timing", "value": "回数券終了時"}])
        _写真(r, "q_measure_photos", f"E{i}", "after")
    html = as_owner.get(一覧).content.decode()
    # 画面の「未分析をまとめて分析」は表示中の先頭 5 件
    assert html.count('name="entry_id"') == 5 + 8  # まとめて 5 ＋ 各お客様の「分析する」8
    ids = [f"extra_{i}" for i in range(6)] + ["resp_after"]
    r = as_owner.post(一覧 + "analyze/", {"entry_id": ids}, follow=True)
    body = r.content.decode()
    assert "分析が完了しました。" in body and "残り 2 件" in body
    assert len(fake_claude) == 5 and TicketAnalysis.objects.filter(status="done").count() == 5
    assert sorted(TicketAnalysis.objects.values_list("sheet_row", flat=True)) == [2, 3, 4, 5, 6]
    # 知らない ID は失敗として数える
    r = as_owner.post(一覧 + "analyze/", {"entry_id": ["nothing"]}, follow=True)
    assert "成功 0 件 / 失敗 1 件" in r.content.decode()
