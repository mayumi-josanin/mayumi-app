"""ビジリス管理「回数券分析」。**中身は旧ビジリス管理アプリの #page-ticket-survey と同じ。**

お手本: bijiris/admin-app/app.js `renderTicketSurvey`（8812行〜）・`renderTicketSurveyEntry`（8778）・
`renderTicketSurveyPhotoGrid`（8754）・`analyzeTicketSurveyEntries`（8980）・`saveTicketSurveyPrompt`（9033）・
`saveTicketSurveyAuto`（9067）。裏側は bijiris/gas/Code.gs 6435行〜（TICKET_SURVEY_*・`getTicketSurveyPayload_`・
`buildTicketSurveyMessageContent_`・`callAnthropicMessages_`・`analyzeTicketSurveyResponses_`）。

読むもの（GAS の置き場 → サーバー）:
  - 回答（アフター写真のあるもの）… bijiris.Response ＋ ResponsePhoto（kind=after/before は取り込み時に Code.gs と同じ判定で付いている）
  - 回数券分析結果 … records.TicketAnalysis（回答IDで結ぶ）
  - TICKET_SURVEY_META_JSON / TICKET_SURVEY_PROMPT … records.AppSetting `bijiris_ticket_meta` / `bijiris_ticket_prompt`
  - 計測記録 … measurements.Measurement
  - ANTHROPIC_API_KEY … settings.ANTHROPIC_API_KEY（.env）。**画面には「設定済み／未設定」だけ出す。**

GAS と違うところ:
  - 写真は Drive ではなくサーバーの media（ResponsePhoto.url）。取り込んでいない写真は「未取り込み」と出し、AI にも渡せない。
  - GAS はモニター写真を Drive の「モニター写真」フォルダへ写してから使う（resolveBeforePhotos_）。サーバーには
    そのフォルダが無いので、回答自身のビフォー写真 → 同じお名前の「初回計測時」の回答の写真、の順で使う。
  - 自動処理は GAS の時間トリガーで動いている。ここでは有効／無効の印（meta.autoEnabled）を書くだけで、
    サーバーは自分では回らない（回す仕組みは切り替え C のあとに考える）。
  - モニター基準画像の一括取り込み（Drive のコピー）と API キーの入力は、サーバーではできないので置かない。
  - GAS は投げっぱなしで画面が取り直しに行くが、ここでは「分析する」を押した通信の中で終わるまで待つ。
    1件 30 秒〜1分かかるので、一度に扱うのは GAS と同じ 5 件まで。

**お名前で会員に結びつけない**（CLAUDE.md）。この画面はビジリスの記録どうし（回答・分析結果・計測記録）を
お名前で寄せるだけで、会員の表は見ない。
"""

import base64
import datetime
import json
import re
import urllib.error
import urllib.request
from pathlib import Path

from django.conf import settings
from django.contrib import messages
from django.db import transaction
from django.shortcuts import redirect, render
from django.urls import reverse
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required
from apps.measurements.models import Measurement
from apps.records.models import AppSetting, TicketAnalysis

from . import gate, preferences
from .models import Response, ResponsePhoto, Survey

# ---- Code.gs の定数（同じ値にしておく。片方だけ変えると切り替えの前後で振る舞いが変わる） ----
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_API_VERSION = "2023-06-01"
ANTHROPIC_MODEL = "claude-opus-5"
ANTHROPIC_MAX_TOKENS = 3000
ANTHROPIC_EFFORT = "medium"
ANTHROPIC_SYSTEM = (
    "あなたは日本語で回答するアシスタントです。指示された出力形式のみを出力し、"
    "考察の途中経過や前置き・後書きは書かないでください。"
)
片側の写真の上限 = 4            # TICKET_SURVEY_MAX_PHOTOS_PER_SIDE
一度に分析する件数 = 5          # TICKET_SURVEY_ANALYZE_BATCH_SIZE
自動処理の間隔分 = 30           # TICKET_SURVEY_AUTO_INTERVAL_MINUTES
画像の大きさの上限 = 4_500_000   # fetchTicketSurveyImageBlob_ の 4500000
通信の待ち秒 = 180
計測タイミングの設問 = "q_measure_timing"

# app.js TICKET_SURVEY_ANALYSIS_LABELS
状態の文言 = {
    "none": "分析待ち",
    "pending": "分析中です",
    "running": "分析中です",
    "done": "分析完了",
    "error": "分析エラー",
}
状態の色 = {"none": "badge-gray", "pending": "badge-orange", "running": "badge-orange", "done": "badge-green", "error": "badge-red"}


# ---------------------------------------------------------------------------
# 小さな道具
# ---------------------------------------------------------------------------

def _文字(v) -> str:
    return "" if v is None else str(v).strip()


def _名を整える(v) -> str:
    """normalizeTicketSurveyName_: 空白（全角も）を全部取る。"""
    return re.sub(r"[\s　]+", "", _文字(v))


def _日付(dt) -> str:
    """formatTicketSurveyDate_: 日本時間の yyyy-MM-dd。無ければ空。"""
    if not dt:
        return ""
    return timezone.localtime(dt).strftime("%Y-%m-%d")


def _数を文字(v) -> str:
    """計測値（Decimal）をシートの文字と同じ見た目に（70.0 → 70、70.5 → 70.5）。"""
    if v is None or v == "":
        return ""
    s = f"{v:f}" if not isinstance(v, str) else v
    return s.rstrip("0").rstrip(".") if "." in s else s


def _回答の文字(response, question_id) -> str:
    """ticketSurveyAnswerText_: 配列なら「、」でつなぐ。"""
    v = response.answer_of(question_id)
    if isinstance(v, list):
        return "、".join(x for x in (_文字(i) for i in v) if x)
    return _文字(v)


def _タイミング(response) -> str:
    """measureTimingOf_: 「初回計測」「モニター」→ monitor、「終了」→ after。"""
    t = _回答の文字(response, 計測タイミングの設問)
    if "初回計測" in t or "モニター" in t:
        return "monitor"
    if "終了" in t:
        return "after"
    return ""


def _メタ() -> dict:
    s = AppSetting.objects.filter(pk=preferences.回数券メタの鍵).first()
    return dict(s.value) if s and isinstance(s.value, dict) else {}


def _メタを更新(patch: dict) -> dict:
    """updateTicketSurveyMeta_: 一部だけ差し替える。"""
    meta = _メタ()
    meta.update(patch)
    AppSetting.objects.update_or_create(
        pk=preferences.回数券メタの鍵,
        defaults={"value": meta, "note": "ビジリス 回数券分析の進行状況（TICKET_SURVEY_META_JSON の写し）"})
    return meta


def _鍵がある() -> bool:
    return bool(_文字(getattr(settings, "ANTHROPIC_API_KEY", "")))


# ---------------------------------------------------------------------------
# 写真
# ---------------------------------------------------------------------------

def _写真の形(p: ResponsePhoto) -> dict:
    """画面と AI に渡す1枚。url が空なら「未取り込み」。"""
    return {"file_id": p.drive_file_id, "name": p.name or "", "url": p.url, "stored": bool(p.url)}


def _記録の写真を当てる(refs, 写真: dict) -> list:
    """分析結果に残っている写真（Drive のファイルID）を、サーバーの写真に当てる。

    GAS はビフォー写真を Drive の「モニター写真」フォルダへ**写して**から記録するので、
    そのファイルIDはサーバーの写真（回答の files）に無いことがある。1枚も当たらなければ空を返し、
    呼ぶ側で回答自身の写真に戻す。
    """
    out, 当たった = [], False
    for ref in refs if isinstance(refs, list) else []:
        fid = _文字(ref.get("fileId")) if isinstance(ref, dict) else ""
        if not fid:
            continue
        p = 写真.get(fid)
        if p:
            当たった = True
            out.append(_写真の形(p))
        else:
            out.append({"file_id": fid, "name": _文字(ref.get("name")), "url": "", "stored": False})
    return out if 当たった else []


def _写真の中身(photo: dict) -> tuple[str, str]:
    """AI に渡す base64。media/bijiris/<名前> を読む（url は PUBLIC_BASE_URL/media/bijiris/<名前>）。"""
    if not photo.get("url"):
        raise RuntimeError(f"{photo.get('name') or '写真'} がまだサーバーに取り込まれていません。"
                           "「ビジリスの回答を取り込む --写真」で入れてください。")
    名前 = photo["url"].rsplit("/", 1)[-1]
    path = Path(settings.MEDIA_ROOT) / "bijiris" / 名前
    if not path.is_file():
        raise RuntimeError(f"{photo.get('name') or '写真'} の画像ファイルがサーバーにありません（{名前}）。")
    中身 = path.read_bytes()
    if len(中身) > 画像の大きさの上限:
        raise RuntimeError(f"{photo.get('name') or '画像'} の画像サイズが大きすぎます。")
    # images.保存する は必ず JPEG で置く
    return "image/jpeg", base64.b64encode(中身).decode("ascii")


# ---------------------------------------------------------------------------
# 一覧（getTicketSurveyPayload_）
# ---------------------------------------------------------------------------

def 一覧を作る() -> list[dict]:
    """アフター写真のある回答（ゴミ箱を除く）を、分析結果と結んで新しい順に。"""
    写真 = {p.drive_file_id: p for p in ResponsePhoto.objects.all()}
    記録 = {t.response_id: t for t in TicketAnalysis.objects.exclude(response_id="")}
    回答ら = list(Response.objects.exclude(status="trash").prefetch_related("photos").order_by("-submitted_at", "-sheet_row"))

    # 同じお名前の「初回計測時」の回答の写真（新しい順で最初に写真があるもの）。findLatestCustomerMonitorPhotos_
    モニター写真 = {}
    for r in 回答ら:
        名 = _名を整える(r.customer_name)
        if not 名 or 名 in モニター写真 or _タイミング(r) != "monitor":
            continue
        own = [p for p in r.photos.all() if p.kind == "before"]
        if own:
            モニター写真[名] = [_写真の形(p) for p in own]

    entries = []
    for r in 回答ら:
        after_own = [p for p in r.photos.all() if p.kind == "after"]
        if not after_own:
            continue
        rec = 記録.get(r.response_id)
        after = _記録の写真を当てる(rec.after_photos, 写真) if rec else []
        if not after:
            after = [_写真の形(p) for p in after_own]
        before = _記録の写真を当てる(rec.before_photos, 写真) if rec else []
        if not before:
            before = [_写真の形(p) for p in r.photos.all() if p.kind == "before"] or モニター写真.get(_名を整える(r.customer_name), [])
        状態 = (_文字(rec.status) if rec else "") or "none"
        entries.append({
            "id": r.response_id,
            "response": r,
            "customer_name": r.customer_name,
            "submitted_at": r.submitted_at,
            "submitted_date": _日付(r.submitted_at),
            "before_photos": before,
            "after_photos": after,
            "status": 状態,
            "status_label": 状態の文言.get(状態, 状態),
            "status_class": 状態の色.get(状態, "badge-gray"),
            "analysis_text": rec.result if rec else "",
            "analyzed_at": rec.analyzed_at if rec else None,
            "error_message": _文字(rec.error) if rec else "",
            "measurements": _計測値の要約(r),
        })
    entries.sort(key=lambda e: e["submitted_at"] or datetime.datetime.min.replace(tzinfo=datetime.timezone.utc), reverse=True)
    return entries


def _絞る(entries, keyword: str, status: str) -> list[dict]:
    """getFilteredTicketSurveyEntries: お名前の部分一致と分析状態。"""
    kw = _文字(keyword).lower()
    return [e for e in entries
            if (not status or e["status"] == status) and (not kw or kw in _文字(e["customer_name"]).lower())]


# ---------------------------------------------------------------------------
# AI への問い合わせの組み立て（buildResponseContextText_ / renderTicketSurveyPrompt_ / buildTicketSurveyMessageContent_）
# ---------------------------------------------------------------------------

def _計測記録(name: str) -> list[Measurement]:
    """同じお名前の計測記録を古い順に。**会員の表は見ない。**"""
    名 = _名を整える(name)
    if not 名:
        return []
    rows = [m for m in Measurement.objects.all() if _名を整える(m.customer_name) == 名]
    rows.sort(key=lambda m: m.measured_on)
    return rows


def _計測記録の参考文(name: str) -> str:
    """buildTicketSurveyMeasurementContext_: 直近 6 回ぶん。"""
    rows = _計測記録(name)[-6:]
    if not rows:
        return ""
    lines = [" / ".join([
        m.measured_on.strftime("%Y-%m-%d"),
        "ウエスト " + (_数を文字(m.waist) or "-"),
        "ヒップ " + (_数を文字(m.hip) or "-"),
        "太もも右 " + (_数を文字(m.thigh_right) or "-"),
        "太もも左 " + (_数を文字(m.thigh_left) or "-"),
    ]) for m in rows]
    return "【計測記録（参考）】\n" + "\n".join(lines)


def _回答の前置き(response: Response) -> str:
    """buildResponseContextText_: お客様・提出日・写真以外の回答・計測記録。"""
    lines = [
        "【お客様】" + (_文字(response.customer_name) or "お名前未設定"),
        "【提出日】" + (_日付(response.submitted_at) or "不明"),
    ]
    label = {}
    survey = response.survey or Survey.objects.filter(survey_id=response.survey_key).first()
    for q in (survey.questions if survey and isinstance(survey.questions, list) else []):
        if isinstance(q, dict) and q.get("id"):
            label[q["id"]] = q.get("label") or q["id"]
    for a in response.answers if isinstance(response.answers, list) else []:
        if not isinstance(a, dict) or isinstance(a.get("files"), list):
            continue  # 写真の回答は除く
        v = a.get("value")
        v = "、".join(_文字(x) for x in v) if isinstance(v, list) else _文字(v)
        if not v:
            continue
        qid = _文字(a.get("questionId"))
        lines.append("【" + (label.get(qid) or qid) + "】" + v)
    参考 = _計測記録の参考文(response.customer_name)
    if 参考:
        lines.append(参考)
    return "\n".join(lines)


def _最新のモニター回答(name: str):
    """findLatestCustomerMonitorResponse_: 同じお名前で「初回計測時」の回答のうち最新。

    GAS は trim だけで比べるが、ここは写真の探し方（_名を整える）とそろえて空白の違いも無視する。
    「佐藤 花子」と「佐藤花子」で初回の計測値が拾えないのを避けるため。
    """
    名 = _名を整える(name)
    if not 名:
        return None
    for r in Response.objects.exclude(status="trash").order_by("-submitted_at", "-sheet_row"):
        if _名を整える(r.customer_name) == 名 and _タイミング(r) == "monitor":
            return r
    return None


def _計測値の要約(response: Response) -> dict:
    """指示文の {{初回ウエスト}} などに入れる値（renderTicketSurveyPrompt_ と同じ探し方）。画面にも出す。"""
    今 = {k: _回答の文字(response, "q_measure_" + k) for k in ("waist", "hip", "thigh_right", "thigh_left")}
    前 = {k: "" for k in 今}
    前日付 = ""
    monitor = _最新のモニター回答(response.customer_name)
    if monitor:
        前 = {k: _回答の文字(monitor, "q_measure_" + k) for k in 前}
        前日付 = _日付(monitor.submitted_at)
    if not any(前.values()):
        rows = _計測記録(response.customer_name)
        if rows:
            m = rows[0]  # getMeasurements_ は新しい順なので末尾＝いちばん古い
            前 = {"waist": _数を文字(m.waist), "hip": _数を文字(m.hip),
                  "thigh_right": _数を文字(m.thigh_right), "thigh_left": _数を文字(m.thigh_left)}
            前日付 = 前日付 or m.measured_on.strftime("%Y-%m-%d")
    return {"before": 前, "after": 今, "before_date": 前日付, "after_date": _日付(response.submitted_at),
            "has_any": any(前.values()) or any(今.values())}


def _指示文を埋める(prompt: str, response: Response, before: list, after: list) -> str:
    """renderTicketSurveyPrompt_: {{お名前}} などを実データに。"""
    if not prompt or "{{" not in prompt:
        return prompt
    値 = _計測値の要約(response)
    improve = _回答の文字(response, "q_measure_improve")
    other = _回答の文字(response, "q_measure_improve_other")
    改善 = (f"{improve}／{other}" if improve else other) if other else improve

    def num(v):
        return v or "-"

    data = {
        "お名前": _文字(response.customer_name) or "お客様",
        "今後もっと改善したい部分はありますか？": 改善,
        "ビフォー日付": 値["before_date"] or "不明",
        "アフター日付": 値["after_date"] or "不明",
        "初回ウエスト": num(値["before"]["waist"]),
        "初回ヒップ": num(値["before"]["hip"]),
        "初回太もも右": num(値["before"]["thigh_right"]),
        "初回太もも左": num(値["before"]["thigh_left"]),
        "今回ウエスト": num(値["after"]["waist"]),
        "今回ヒップ": num(値["after"]["hip"]),
        "今回太もも右": num(値["after"]["thigh_right"]),
        "今回太もも左": num(値["after"]["thigh_left"]),
        "ビフォー枚数": str(len(before)),
        "アフター枚数": str(len(after)),
    }
    for k, v in data.items():
        prompt = prompt.replace("{{" + k + "}}", v)
    return prompt


def 問い合わせの中身(entry: dict, prompt: str) -> tuple[list, list, list]:
    """buildTicketSurveyMessageContent_: 前置き → ビフォー写真 → アフター写真 → 指示文。

    返り値: (content, 使ったビフォー写真, 使ったアフター写真)
    """
    response = entry["response"]
    before = [p for p in entry["before_photos"] if p["stored"]][:片側の写真の上限]
    after_all = entry["after_photos"]
    after = [p for p in after_all if p["stored"]][:片側の写真の上限]
    if not after_all:
        raise RuntimeError("アフター写真（回数券終了時の写真）がありません。")
    if not after:
        raise RuntimeError("アフター写真がまだサーバーに取り込まれていません。"
                           "「ビジリスの回答を取り込む --写真」で入れてください。")

    content = [{"type": "text", "text": _回答の前置き(response)}]
    content.append({"type": "text", "text": "以下は【初回計測時（ビフォー）】の写真です。" if before
                    else "【初回計測時（ビフォー）】の写真はありません。"})
    for p in before:
        mime, data = _写真の中身(p)
        content.append({"type": "image", "source": {"type": "base64", "media_type": mime, "data": data}})
    content.append({"type": "text", "text": "以下は【回数券終了時（アフター）】の写真です。"})
    for p in after:
        mime, data = _写真の中身(p)
        content.append({"type": "image", "source": {"type": "base64", "media_type": mime, "data": data}})
    content.append({"type": "text", "text": _指示文を埋める(prompt, response, before, after)})
    return content, before, after


def AIに問い合わせる(api_key: str, content: list) -> str:
    """callAnthropicMessages_ と同じ要求（モデル・上限・思考なし・effort・system）。urllib で送る。"""
    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": ANTHROPIC_MAX_TOKENS,
        "thinking": {"type": "disabled"},
        "output_config": {"effort": ANTHROPIC_EFFORT},
        "system": ANTHROPIC_SYSTEM,
        "messages": [{"role": "user", "content": content}],
    }
    req = urllib.request.Request(
        ANTHROPIC_API_URL, data=json.dumps(payload).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json", "x-api-key": api_key, "anthropic-version": ANTHROPIC_API_VERSION})
    status, raw = 0, b""
    try:
        with urllib.request.urlopen(req, timeout=通信の待ち秒) as res:
            status, raw = res.status, res.read()
    except urllib.error.HTTPError as e:
        status, raw = e.code, e.read()
    except (urllib.error.URLError, OSError) as e:  # TimeoutError は OSError
        raise RuntimeError(f"Claude API に届きませんでした: {e}")
    try:
        body = json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        body = None
    if status != 200 or not isinstance(body, dict):
        detail = (body.get("error") or {}).get("message") if isinstance(body, dict) else ""
        detail = detail or raw.decode("utf-8", "replace")[:300]
        raise RuntimeError(f"Claude API エラー ({status}): {detail}")
    if body.get("stop_reason") == "refusal":
        raise RuntimeError("Claude API が回答を拒否しました。写真や指示文の内容を確認してください。")
    text = "\n".join(b.get("text") or "" for b in body.get("content") or [] if isinstance(b, dict) and b.get("type") == "text").strip()
    if not text:
        raise RuntimeError("分析結果が空でした。もう一度お試しください。")
    return text


# ---------------------------------------------------------------------------
# 分析（analyzeTicketSurveyResponse_ / analyzeTicketSurveyResponses_）
# ---------------------------------------------------------------------------

def _写真の記録(photos: list) -> list:
    """分析結果に残す写真の手がかり。Drive のリンクは持たない（設計）。"""
    return [{"fileId": p["file_id"], "name": p["name"]} for p in photos]


def _記録を用意(entry: dict) -> TicketAnalysis:
    """回答IDの記録があればそれ、無ければ新しい行（GAS の appendRow と同じく末尾の行番号）。"""
    rec = TicketAnalysis.objects.filter(response_id=entry["id"]).order_by("sheet_row").first()
    if rec:
        return rec
    最後 = TicketAnalysis.objects.order_by("-sheet_row").values_list("sheet_row", flat=True).first() or 1
    now = timezone.now()
    return TicketAnalysis(sheet_row=最後 + 1, response_id=entry["id"], member_id="", created_at=now, updated_at=now)


def 一件を分析する(entry: dict, api_key: str) -> TicketAnalysis:
    """1件ぶん。途中は running、終われば done、失敗は error と理由。失敗は RuntimeError で返す。"""
    rec = _記録を用意(entry)
    rec.customer_name = entry["customer_name"] or ""
    rec.submitted_at = entry["submitted_at"]
    rec.status = "running"
    rec.error = ""
    rec.updated_at = timezone.now()
    rec.save()
    try:
        content, before, after = 問い合わせの中身(entry, preferences.分析指示文を読む())
        rec.before_photos = _写真の記録(before)
        rec.after_photos = _写真の記録(after)
        text = AIに問い合わせる(api_key, content)
        rec.status = "done"
        rec.result = text
        rec.analyzed_at = timezone.now()
        rec.error = ""
        rec.updated_at = rec.analyzed_at
        rec.save()
        return rec
    except RuntimeError as e:
        rec.status = "error"
        rec.error = str(e)
        rec.updated_at = timezone.now()
        rec.save()
        raise


def まとめて分析する(ids: list[str], api_key: str) -> dict:
    """analyzeTicketSurveyResponses_: 先頭 5 件だけ。残りは deferred として数える。"""
    ids = [i for i in (_文字(x) for x in ids) if i]
    if not ids:
        raise RuntimeError("分析する回答を選んでください。")
    batch, deferred = ids[:一度に分析する件数], ids[一度に分析する件数:]
    by_id = {e["id"]: e for e in 一覧を作る()}
    succeeded, failures = 0, []
    for i in batch:
        entry = by_id.get(i)
        if not entry:
            failures.append(f"{i}: 対象の回答が見つかりませんでした。")
            continue
        try:
            一件を分析する(entry, api_key)
            succeeded += 1
        except RuntimeError as e:
            failures.append(f"{i}: {e}")
    return {"succeeded": succeeded, "failures": failures, "deferred": len(deferred)}


# ---------------------------------------------------------------------------
# 画面
# ---------------------------------------------------------------------------

def _戻り先(request):
    """一覧へ戻す。絞り込みを保ったまま戻れるよう next を受けるが、この画面の中だけ。"""
    next_url = request.POST.get("next") or ""
    一覧 = reverse("manage:bijiris:ticket_list")
    return next_url if next_url.startswith(一覧) else 一覧


def _書けるか(request) -> bool:
    if gate.ビジリスはサーバーが正():
        return True
    messages.error(request, gate.断る文())
    return False


@owner_required
def ticket_list(request):
    q = _文字(request.GET.get("q"))
    status = _文字(request.GET.get("status"))
    if status not in 状態の文言:
        status = ""
    entries = 一覧を作る()
    shown = _絞る(entries, q, status)
    # 「未分析をまとめて分析」: 分析完了でなく、アフター写真があるものを先頭 5 件（app.js と同じ）
    pending_ids = [e["id"] for e in shown if e["status"] != "done" and e["after_photos"]][:一度に分析する件数]
    meta = _メタ()
    prompt = preferences.分析指示文を読む()
    seed_summary = meta.get("monitorSeedSummary") if isinstance(meta.get("monitorSeedSummary"), dict) else None
    return render(request, "bijiris/ticket_list.html", {
        "entries": shown,
        "total": len(entries),
        "q": q,
        "status": status,
        "status_choices": [("", "すべて"), ("none", "分析待ち"), ("pending", "分析中です"), ("running", "分析中です"),
                           ("done", "分析完了"), ("error", "分析エラー")],
        "pending_ids": pending_ids,
        "prompt": prompt,
        "default_prompt": preferences.既定の分析指示文,
        "model": ANTHROPIC_MODEL,
        "api_key_configured": _鍵がある(),
        "auto_enabled": meta.get("autoEnabled") is True,
        "auto_interval": 自動処理の間隔分,
        "last_auto_run_at": _文字(meta.get("lastAutoRunAt")),
        "auto_error": _文字(meta.get("autoError")),
        "monitor_seeded_at": _文字(meta.get("monitorSeededAt")),
        "monitor_seed_summary": seed_summary,
        "server_is_source": gate.ビジリスはサーバーが正(),
        "refuse_text": gate.断る文(),
        "empty_text": ("条件に合うアンケートがありません。" if entries
                       else "まだ取り込んでいません。「写真を取り込む」を押してください。"),
    })


@owner_required
@require_POST
def ticket_analyze(request):
    """「分析する」「再分析する」「未分析をまとめて分析」。"""
    if not _書けるか(request):
        return redirect(_戻り先(request))
    ids = request.POST.getlist("entry_id")
    if not ids:
        messages.error(request, "分析するアンケートがありません。")
        return redirect(_戻り先(request))
    if not _鍵がある():
        # 鍵は .env（ANTHROPIC_API_KEY）。画面からは入れられない
        messages.error(request, "AI の鍵が設定されていません。サーバーの設定（ANTHROPIC_API_KEY）に入れてから、もう一度お試しください。")
        return redirect(_戻り先(request))
    try:
        r = まとめて分析する(ids, settings.ANTHROPIC_API_KEY)
    except RuntimeError as e:
        messages.error(request, str(e))
        return redirect(_戻り先(request))
    総数 = r["succeeded"] + len(r["failures"])
    if r["failures"]:
        messages.error(request, f"分析完了：成功 {r['succeeded']} 件 / 失敗 {len(r['failures'])} 件（理由は各お客様の欄に出ています）")
    else:
        messages.success(request, "分析が完了しました。" if 総数 else "分析するアンケートがありません。")
    if r["deferred"]:
        messages.info(request, f"一度に分析できるのは {一度に分析する件数} 件までです。残り {r['deferred']} 件はもう一度「未分析をまとめて分析」を押してください。")
    return redirect(_戻り先(request))


@owner_required
@require_POST
def ticket_prompt_save(request):
    """「プロンプトを保存」。空なら既定の指示文に戻る（saveTicketSurveyPrompt_ と同じ）。"""
    if not _書けるか(request):
        return redirect(_戻り先(request))
    text = _文字(request.POST.get("prompt"))
    with transaction.atomic():
        AppSetting.objects.update_or_create(
            pk=preferences.回数券指示文の鍵,
            defaults={"value": {"prompt": text}, "note": "ビジリス 回数券分析の AI への指示文（TICKET_SURVEY_PROMPT の写し）"})
    messages.success(request, "プロンプトを保存しました。")
    return redirect(_戻り先(request))


@owner_required
@require_POST
def ticket_auto(request):
    """「自動で取り込み・分析する」の有効／無効（setTicketSurveyAuto_）。"""
    if not _書けるか(request):
        return redirect(_戻り先(request))
    on = request.POST.get("enabled") == "1"
    if on and not _鍵がある():
        messages.error(request, "先に AI の鍵（ANTHROPIC_API_KEY）をサーバーの設定に入れてください。")
        return redirect(_戻り先(request))
    _メタを更新({"autoEnabled": on})
    messages.success(request, "自動処理を有効にしました。" if on else "自動処理を停止しました。")
    return redirect(_戻り先(request))
