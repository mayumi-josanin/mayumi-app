"""ビジリス管理「🎁 特典」。**中身は旧ビジリス管理アプリ（#page-gacha）と同じ。**

    マイルストーン特典   回数券を使い切った「完了枚数」がしきい値に達したお客様に渡す特典
                        （お客様アプリに表示する／しきい値(枚)・特典内容・特典内容説明／行を追加・削除・保存）
    月別ガチャ特典       行ごとに月、列ごとに A〜D賞の特典内容と出現確率（月を追加・削除・保存）
    キャンペーンスタンプ お客様アプリにキャンペーンスタンプを表示する（旧アプリでは 🔐 設定にあった）
    受け取り状況         特典に届いた方の一覧（スタンプ・未受取・受取済み。渡すのは顧客管理の画面で）

旧アプリの特典ページ（index.html #page-gacha）に置かれているのはマイルストーン特典だけで、
月別ガチャ特典（app.js renderGachaPrizeManager 545行〜）は置き場（#gachaPrizeManagerStage）が無く描かれていなかった。
お客様アプリが読む設定（gachaPrizeConfig）なので、ここでは直せるようにする。

設定の置き場は records.AppSetting の鍵 `bijiris_preferences`（旧アプリの ADMIN_PREFERENCES_JSON）。
保存は preferences.正規化する を通す（Code.gs の normalizeGachaPrizeConfig_ / normalizeMilestoneRewardConfig_ と同じ）。
**書き込みは gate.ビジリスはサーバーが正() が立つまで断る**（旧アプリの GAS と正が2つになるため）。
"""

from django.contrib import messages
from django.shortcuts import redirect, render
from django.utils import timezone
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required
from apps.records.models import AppSetting

from . import gate, preferences
from .models import CustomerProfile, Response
from .views_customers import マイルストーンの状況

賞 = preferences.ガチャの景品の鍵


def _月の見出し(月) -> str:
    """旧アプリの formatMonthLabel（2026年9月。読めなければ 未設定）。"""
    月 = preferences.月の鍵(月)
    return f"{月[:4]}年{int(月[5:7])}月" if 月 else "未設定"


def _次の月(月々) -> str:
    """旧アプリの getNextMonthKey: いちばん遅い月の翌月。無ければ今月の翌月。"""
    最後 = max((e["month"] for e in 月々), default=preferences.月の鍵(timezone.localdate().isoformat()))
    年, 月 = int(最後[:4]), int(最後[5:7])
    return f"{年 + (月 // 12):04d}-{(月 % 12) + 1:02d}"


def _ガチャの行(設定) -> list:
    出 = []
    for i, e in enumerate(設定["monthlyPrizes"]):
        合計 = round(sum(float(e["prizes"][k]["probability"]) for k in 賞), 1)
        出.append({"index": i, "month": e["month"], "label": _月の見出し(e["month"]),
                   "prizes": [{"key": k, **e["prizes"][k]} for k in 賞],
                   "total": int(合計) if 合計 == int(合計) else 合計, "total_ok": abs(合計 - 100) < 0.001})
    return 出


def _受け取り状況(設定) -> list:
    """特典に届いた方（スタンプが最初の節目以上）の一覧。顧客ごとの数え方は顧客管理と同じ。"""
    if not 設定["milestones"]:
        return []
    回答 = {}
    for r in Response.objects.exclude(status="trash").order_by("-submitted_at", "-sheet_row"):
        回答.setdefault(r.customer_name, []).append(r)
    名前 = set(回答) | set(CustomerProfile.objects.values_list("name", flat=True))
    プロフィール = {p.name: p for p in CustomerProfile.objects.filter(name__in=名前)}
    出 = []
    for name in 名前:
        状況 = マイルストーンの状況(プロフィール.get(name), 回答.get(name, []), 設定)
        達成 = [m for m in 状況["rows"] if m["achieved"]]
        if not 達成:
            continue
        出.append({"name": name, "member_number": プロフィール[name].member_number if name in プロフィール else "",
                   "total": 状況["total"], "pending": 状況["pending"], "handed": sum(1 for m in 達成 if m["handed"]),
                   "achieved": 達成})
    # 未受取が多い方から上に。同じなら スタンプが多い方・お名前順
    return sorted(出, key=lambda x: (-x["pending"], -x["total"], x["member_number"], x["name"]))


@owner_required
def reward_list(request):
    設定 = preferences.読む()
    節目 = 設定["milestoneRewardConfig"]
    return render(request, "bijiris/reward_list.html", {
        "milestone_enabled": 節目["enabled"],
        "milestones": 節目["milestones"] or [{"threshold": "", "reward": "", "description": ""}],
        "gacha_rows": _ガチャの行(設定["gachaPrizeConfig"]),
        "gacha_next_month": _次の月(設定["gachaPrizeConfig"]["monthlyPrizes"]),
        "campaign_stamp_enabled": 設定["campaignStampEnabled"],
        "redemptions": _受け取り状況(節目),
        "server_is_source": gate.ビジリスはサーバーが正(), "refuse_text": gate.断る文(),
    })


def _保存する(更新: dict):
    """いまの設定に重ねて正規化し、AppSetting へ。updatePreferences_ と同じく足りない項目は既定で埋まる。"""
    次 = preferences.正規化する({**preferences.読む(), **更新})
    AppSetting.objects.update_or_create(pk=preferences.鍵, defaults={"value": 次, "note": "ビジリスの設定（ADMIN_PREFERENCES_JSON）"})
    return 次


@owner_required
@require_POST
def reward_milestone_save(request):
    """旧アプリの saveMilestoneRewardConfig。しきい値が無い・特典内容が空の行は捨てる（collectMilestoneRewardConfigFromTable）。"""
    if not gate.ビジリスはサーバーが正():
        messages.error(request, gate.断る文())
        return redirect("manage:bijiris:reward_list")
    行 = []
    for i in request.POST.getlist("row"):
        行.append({"threshold": request.POST.get(f"threshold_{i}", ""), "reward": request.POST.get(f"reward_{i}", ""),
                  "description": request.POST.get(f"description_{i}", "")})
    _保存する({"milestoneRewardConfig": {"enabled": request.POST.get("enabled") == "on", "milestones": 行}})
    messages.success(request, "マイルストーン特典設定を保存しました。")
    return redirect("manage:bijiris:reward_list")


@owner_required
@require_POST
def reward_gacha_save(request):
    """旧アプリの saveGachaPrizeConfig。月が無い・重なる行は旧アプリと同じ言い分で断る。"""
    if not gate.ビジリスはサーバーが正():
        messages.error(request, gate.断る文())
        return redirect("manage:bijiris:reward_list")
    月々 = []
    見た = set()
    for i in request.POST.getlist("row"):
        月 = preferences.月の鍵(request.POST.get(f"month_{i}", ""))
        if not 月:
            messages.error(request, "月を選択してください。")
            return redirect("manage:bijiris:reward_list")
        if 月 in 見た:
            messages.error(request, "同じ月が重複しています。月ごとに1行だけ設定してください。")
            return redirect("manage:bijiris:reward_list")
        見た.add(月)
        月々.append({"month": 月, "prizes": {k: {"content": request.POST.get(f"content_{i}_{k}", ""),
                                             "probability": request.POST.get(f"probability_{i}_{k}", "")} for k in 賞}})
    if not 月々:
        messages.error(request, "設定する月は1件以上必要です。")
        return redirect("manage:bijiris:reward_list")
    _保存する({"gachaPrizeConfig": {"monthlyPrizes": 月々}})
    messages.success(request, "ガチャ特典設定を保存しました。")
    return redirect("manage:bijiris:reward_list")


@owner_required
@require_POST
def reward_campaign_save(request):
    """キャンペーンスタンプの表示（旧アプリの 🔐 設定にあったチェック1つ）。"""
    if not gate.ビジリスはサーバーが正():
        messages.error(request, gate.断る文())
        return redirect("manage:bijiris:reward_list")
    _保存する({"campaignStampEnabled": request.POST.get("enabled") == "on"})
    messages.success(request, "キャンペーンスタンプの設定を保存しました。")
    return redirect("manage:bijiris:reward_list")
