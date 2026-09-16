"""公式サイトの「🧩 ページの編集」。中身は mayumi-site/admin/app.py の sec_layout と /api/layout-* と同じ。

ページの中身を「かたまり（ブロック）」の並びとして取り出し、
  文章を直す・写真を足す・幅と寄せを変える・並べ替える → 見え方を確かめる → 「この配置で確定」
という流れ。かたまりの取り出しと書き戻しは mayumi-site/admin/layout.py（repo.部品("layout")）が行う。

app.py にしか無いもの（is_sizable / layout_payload / price_key_of / refresh_price_block /
PRICE_LABELS / 見え方の中で直すための EDIT_JS）はここに写した。
"""

import json
import os
import re

from django.http import Http404, HttpResponse, JsonResponse
from django.views.decorators.clickjacking import xframe_options_sameorigin
from django.shortcuts import render
from django.urls import reverse
from django.views.decorators.http import require_POST

from apps.manage.permissions import owner_required

from . import repo, site_images, views

# 料金表の名前（content.json の prices のキー）→ 画面に出す名前。app.py の PRICE_LABELS と同じ
PRICE_LABELS = {
    "bonyu-massage": "母乳外来：乳房マッサージ",
    "bonyu-oushin": "母乳外来：往診",
    "bonyu-soudan": "母乳外来：ご相談",
    "bonyu-all": "母乳外来：まとめた表（料金ページ用）",
    "femcare": "産後ケア：フェムケア",
    "itothermy": "イトオテルミー温熱療法",
    "acupuncture-day": "鍼灸：診療日",
    "acupuncture": "鍼灸：コース",
}

# 見え方の枠。スマホ/タブレット/PC の幅（app.py の sec_layout と同じ）
DEVICES = [("スマホ", 390), ("タブレット", 820), ("PC", 1280)]


def _json(obj, status=200):
    return JsonResponse(obj, status=status, json_dumps_params={"ensure_ascii": False})


def is_sizable(b):
    """幅・寄せを変えられるブロックか。写真・表・地図など、幅が意味を持つものだけ。"""
    h = b.get("html", "")
    cls = re.search(r'class="([^"]*)"', h)
    cls = cls.group(1) if cls else ""
    if b.get("tag") == "img" or h.lstrip().startswith("<img"):
        return True
    return any(k in cls for k in ("lead-photo", "map-embed", "price-table", "staff__photo"))


def price_key_of(b):
    """このブロックが料金表なら、その表の名前（content.json の prices のキー）を返す。"""
    core = repo.部品("core")
    m = core.BLOCK_RE.search(b.get("html", ""))
    name = m.group(2) if m else ""
    return name[len("price-"):] if name.startswith("price-") else ""


def _新しい(b):
    """layout.new_block が作るブロックは HTML の先頭に改行と字下げが入っている。
    そのままだと「まるごと文章として直せるか」の判定（先頭が < か）に外れて、
    確定するまで編集欄が開かない（旧アプリの癖）。前の空白を pre に移す。書き戻す HTML は同じ。"""
    h = b.get("html", "")
    b["pre"] = (b.get("pre") or "") + h[:len(h) - len(h.lstrip())]
    b["html"] = h.lstrip()
    return b


def _整える(b):
    """画面に出すための札（種類・要約・幅と寄せ・直し方）を、ブロックに付け直す。"""
    layout = repo.部品("layout")
    b["label"] = layout.label(b)
    b["summary"] = layout.summary(b)
    b["opts"] = layout.get_opts(b)
    b["sizable"] = is_sizable(b)
    try:
        kind = layout.editable_kind(b)
    except Exception:
        kind = ""
    b["editable"] = kind or ""
    b["text"] = layout.to_text(b) if kind else ""
    # まるごと直せないものは、中の文章をひとつずつ直せるようにする
    b["parts"] = [] if kind else layout.text_parts(b)
    # 箇条書きは、1つの枠に1行ずつ書く形にして、項目を足せるようにする
    b["lists"] = [] if kind else layout.list_parts(b)
    # 料金表は「料金表」の欄と同じ形で直せるようにする
    b["price_key"] = price_key_of(b)
    return b


def _runs(layout, blocks):
    """ひと続きの文章（見出し・段落・箇条書きが続くところ）は1つの枠で書けるようにする。"""
    return [{"a": a, "z": z, "text": layout.run_to_text(blocks[a:z])} for a, z in layout.runs_of(blocks)]


def layout_payload(name):
    """配置編集の画面に渡すデータを作る（app.py の layout_payload と同じ）。"""
    layout = repo.部品("layout")
    core = repo.部品("core")
    d = layout.load_page(name)
    for b in d["blocks"]:
        _整える(b)
    d["runs"] = _runs(layout, d["blocks"])
    # 幅の選択肢は「選択肢の設定」で変えられる
    d["widths"] = [[o.get("label", ""), o.get("value", "")] for o in core.options(None, "widths")]
    d["aligns"] = layout.ALIGNS
    # 合わせた表は、行も入れて渡す（画面では読むだけ）
    data = core.load()
    d["prices"] = {k: (core.merged_price(data, k) if v.get("merge") else v) for k, v in data["prices"].items()}
    d["price_labels"] = PRICE_LABELS
    return d


def refresh_price_block(b, data):
    """料金表ブロックのHTMLを、いまのデータで作り直す（目印の中だけ差し替える）。"""
    core = repo.部品("core")
    blocks = core.build_blocks(data)

    def sub(m):
        name = m.group(2)
        if name not in blocks:
            return m.group(0)
        return "%s\n%s\n%s" % (m.group(1), blocks[name], " " * 8 + m.group(4))

    nb = dict(b)
    nb["html"] = core.BLOCK_RE.sub(sub, b["html"])
    return _整える(nb)


def _preview_dir():
    core = repo.部品("core")
    return os.path.join(core.BASE, "_preview")


# ─────────────────────────────────────────────── 画面

@owner_required
def hp_layout(request):
    if not repo.設定されているか():
        return views._設定なし(request)
    layout = repo.部品("layout")
    page = request.GET.get("page") or layout.PAGES[0]
    if page not in layout.PAGES:
        page = layout.PAGES[0]
    return render(request, "manage/hp_layout.html", {
        "pages": [(p, layout.PAGE_LABEL[p]) for p in layout.PAGES],
        "page": page,
        "devices": DEVICES,
    })


@owner_required
def hp_layout_data(request):
    """選んだページのかたまり一覧（/api/layout と同じ）。"""
    if not repo.設定されているか():
        return _json({"ok": False, "log": "公式サイトの置き場所がまだ設定されていません"})
    layout = repo.部品("layout")
    name = request.GET.get("page") or ""
    if name not in layout.PAGES:
        return _json({"ok": False, "log": "そのページはありません"})
    try:
        return _json({"ok": True, "data": layout_payload(name)})
    except Exception as ex:
        return _json({"ok": False, "log": str(ex)})


@owner_required
@require_POST
def hp_layout_api(request, what):
    """かたまりの操作（/api/layout-opt … /api/layout-save と同じ）。JSON で受けて JSON で返す。"""
    if not repo.設定されているか():
        return _json({"ok": False, "log": "公式サイトの置き場所がまだ設定されていません"})
    layout = repo.部品("layout")
    core = repo.部品("core")
    try:
        body = json.loads(request.body.decode("utf-8") or "{}")
    except ValueError:
        body = {}
    try:
        if what == "opt":
            # 幅・寄せ
            b = layout.set_opts(body["block"], **{body["key"]: body["value"]})
            return _json({"ok": True, "block": _整える(b)})

        if what == "new":
            # 「＋」から文章・見出しを足す
            kind = body.get("kind", "p")
            if kind not in ("p", "h2", "h3"):
                return _json({"ok": False, "log": "不明な種類です"})
            b = _新しい(layout.new_block(kind, value=body.get("value", "")))
            return _json({"ok": True, "block": _整える(b)})

        if what == "text":
            # まるごと文章として直せるかたまり
            b = layout.from_text(body["block"], body["text"])
            return _json({"ok": True, "block": _整える(b)})

        if what == "run":
            # ひと続きの文章を、1つの枠の内容で作り直す
            a = int(body["a"])
            z = int(body["z"])
            blocks = body["blocks"]
            made = [_整える(nb) for nb in layout.run_from_text(body["text"], blocks[a:z])]
            out = blocks[:a] + made + blocks[z:]
            return _json({"ok": True, "blocks": out, "runs": _runs(layout, out)})

        if what == "runs":
            # 並べ替え・追加・削除のあとに、「1つの枠で書ける範囲」を計算し直す
            return _json({"ok": True, "runs": _runs(layout, body["blocks"])})

        if what == "parts":
            # かたまりの中の文章だけを直す。直していないところは1文字も変えない。
            vals = {int(k): v for k, v in (body.get("values") or {}).items()}
            b = layout.apply_text_parts(body["block"], vals)
            # 箇条書きは項目ごと作り直す（行を足す・減らすため）。
            # 文章の差し替えより後に行う。先にやると位置がずれる。
            lists = {int(k): v for k, v in (body.get("lists") or {}).items()}
            if lists:
                b = layout.apply_list_parts(b, lists)
            return _json({"ok": True, "block": _整える(b)})

        if what == "preview":
            name = body["page"]
            if name not in layout.PAGES:
                return _json({"ok": False, "log": "そのページはありません"})
            html = open(layout.page_path(name), encoding="utf-8").read()
            new = layout.apply_to_html(html, body["data"])
            # 画面の文字と編集欄を結びつける印を付ける。
            # プレビュー用のフォルダにしか書かないので、サイトには混ざらない。
            # 料金表はページではなくデータ側を直すので、料金データも渡す。
            new = layout.annotate(new, (core.load() or {}).get("prices"))
            dst = os.path.join(_preview_dir(), name, "index.html")
            os.makedirs(os.path.dirname(dst), exist_ok=True)
            with open(dst, "w", encoding="utf-8") as f:
                f.write(new)
            return _json({"ok": True})

        if what == "save":
            name = body["page"]
            if name not in layout.PAGES:
                return _json({"ok": False, "log": ["そのページはありません。"]})
            log = ["バックアップ: " + os.path.basename(core.backup())]
            try:
                changed = layout.save_page(name, body["data"])
                log.append("配置を書き込みました。" if changed else "変更はありませんでした。")
            except Exception as ex:
                log.append("失敗しました: " + str(ex))
            return _json({"ok": True, "log": log})

        if what == "price-save":
            # 「ページの編集」から料金表を直す。中身は content.json に入れて、
            # そのブロックのHTMLを作り直して返す（サイトのファイルはまだ書き換えない）。
            d = core.load()
            key = body["key"]
            if key not in d["prices"]:
                return _json({"ok": False, "log": "その料金表はありません"})
            p = d["prices"][key]
            p["caption"] = (body.get("caption") or "").strip()
            p["note"] = (body.get("note") or "").strip()
            if not p.get("merge"):
                rows = [[str(c) for c in row] for row in (body.get("rows") or [])]
                p["rows"] = [r for r in rows if any(x.strip() for x in r)]
            core.save(d)
            return _json({"ok": True, "block": refresh_price_block(body["block"], d), "prices": _prices(core, d)})

        if what == "price-preview":
            # 入力した内容で見え方だけを作る。content.json は書き換えない。
            d = core.load()
            key = body["key"]
            p = dict(d["prices"][key])
            p["caption"] = (body.get("caption") or "").strip()
            p["note"] = (body.get("note") or "").strip()
            if not p.get("merge"):
                p["rows"] = [[str(c) for c in row] for row in (body.get("rows") or [])]
            d["prices"] = dict(d["prices"])
            d["prices"][key] = p
            return _json({"ok": True, "block": refresh_price_block(body["block"], d)})

        if what == "price-cell":
            # 見え方の画面で料金表のマス目を押して直す。
            # 旧アプリは「料金表」の入力欄ごと保存していたが、ここは料金の欄が別画面なので
            # そのマスだけ content.json に入れる。
            d = core.load()
            key = body["key"]
            if key not in d["prices"] or d["prices"][key].get("merge"):
                return _json({"ok": False, "log": "その料金表はここでは直せません"})
            r, c = int(body["r"]), int(body["c"])
            rows = d["prices"][key].get("rows") or []
            if not (0 <= r < len(rows) and 0 <= c < len(rows[r])):
                return _json({"ok": False, "log": "そのマスは見つかりませんでした"})
            rows[r][c] = str(body.get("text") or "")
            core.save(d)
            return _json({"ok": True, "block": refresh_price_block(body["block"], d), "prices": _prices(core, d)})
    except Exception as ex:
        return _json({"ok": False, "log": str(ex)})
    raise Http404()


def _prices(core, d):
    return {k: (core.merged_price(d, k) if v.get("merge") else v) for k, v in d["prices"].items()}


@owner_required
@require_POST
def hp_layout_image(request):
    """新しい写真をサイトに追加する（既存の差し替えではなく、増やす）。app.py の handle_new_image と同じ。"""
    if not repo.設定されているか():
        return _json({"ok": False, "log": "公式サイトの置き場所がまだ設定されていません"})
    f = request.FILES.get("file")
    if not f:
        return _json({"ok": False, "log": "ファイルが受け取れませんでした。"})
    try:
        saved = site_images.新しい写真(f.read(), f.name)
    except Exception as ex:
        return _json({"ok": False, "log": str(ex)})
    layout = repo.部品("layout")
    b = _整える(_新しい(layout.new_block("img", filename=saved, alt=(request.POST.get("alt") or "").strip())))
    b["sizable"] = True
    return _json({"ok": True, "block": b, "filename": saved})


# 見え方の枠の中で「文字を押したら、その場で書き直せる」ための仕掛け。app.py の EDIT_JS の写し。
# 配るときだけ差し込むので、_preview のファイルにも、サイトのファイルにも残らない。
EDIT_JS = r"""
<style>
  [data-ed]:hover, [data-part]:hover, [data-list]:hover, [data-price]:hover{
    cursor:text; border-radius:5px;
    box-shadow:0 0 0 2px #f38e0a, 0 0 0 6px rgba(243,142,10,.20);
  }
  [data-ed][data-inner]:hover{ box-shadow:none; cursor:default }
  [data-ed]:has([data-price]):hover{ box-shadow:none; cursor:default }
  [data-ed].ed-on, [data-part].ed-on, [data-list].ed-on, [data-price].ed-on{
    border-radius:5px;
    box-shadow:0 0 0 2px #e36565, 0 0 0 7px rgba(227,101,101,.25);
  }
  #ed-box{
    position:absolute; z-index:99998; box-sizing:border-box;
    background:#fff; border:2px solid #e36565; border-radius:10px;
    box-shadow:0 12px 40px rgba(87,66,55,.28); padding:10px 12px 11px;
    font:14px/1.7 "Noto Sans JP","Hiragino Kaku Gothic ProN","Yu Gothic",sans-serif;
    color:#574237;
  }
  #ed-box .ed-ttl{ font-weight:700; font-size:12px; color:#84694e; margin-bottom:6px }
  #ed-box textarea{
    width:100%; box-sizing:border-box; min-height:84px; resize:vertical;
    border:1px solid #e6d7c8; border-radius:6px; padding:9px 10px;
    font:15px/1.9 inherit; color:#574237; background:#fffdf8;
  }
  #ed-box textarea:focus{ outline:2px solid #f38e0a; outline-offset:1px }
  #ed-box .ed-hint{ font-size:11.5px; color:#84694e; margin:6px 2px 8px }
  #ed-box .ed-foot{ display:flex; gap:8px; align-items:center }
  #ed-box button{
    font:700 13px/1 inherit; padding:9px 15px; border-radius:999px;
    border:2px solid #e36565; cursor:pointer;
  }
  #ed-box .ed-ok{ background:#e36565; color:#fff }
  #ed-box .ed-ng{ background:#fff; color:#574237; border-color:#e6d7c8 }
  #ed-box .ed-keys{ margin-left:auto; font-size:11px; color:#a08a75 }
  html.ed-touch{ scrollbar-width:none; -ms-overflow-style:none; }
  html.ed-touch::-webkit-scrollbar{ width:0; height:0; }
  #ed-tip{
    position:fixed; left:50%; bottom:14px; transform:translateX(-50%);
    background:rgba(87,66,55,.92); color:#fff; padding:7px 16px;
    border-radius:999px; font:500 13px/1.6 sans-serif; z-index:99999;
    pointer-events:none; transition:opacity .5s;
  }
</style>
<script>
(function(){
  var cur = null, box = null, curBlock = -1;

  function mark(el){
    if(cur) cur.classList.remove('ed-on');
    cur = el;
    if(el) el.classList.add('ed-on');
  }
  function send(o){ o.from = 'ed-preview'; parent.postMessage(o, '*'); }
  function tip(msg, ms){
    var t = document.getElementById('ed-tip');
    if(!t){ t = document.createElement('div'); t.id = 'ed-tip'; document.body.appendChild(t) }
    t.textContent = msg; t.style.opacity = 1;
    clearTimeout(t._h);
    t._h = setTimeout(function(){ t.style.opacity = 0 }, ms || 2600);
  }
  function close(){
    if(box){ box.remove(); box = null }
    curBlock = -1;
  }

  /* 押した文字のすぐ下に入力欄を開く。下に入らなければ上に出す。 */
  function openEditor(i, text, kind, label, part, list){
    close();
    var el;
    if(part != null || list != null){
      var blk = document.querySelector('[data-ed="' + i + '"]');
      var sel = part != null ? '[data-part="' + part + '"]' : '[data-list="' + list + '"]';
      el = blk ? blk.querySelector(sel) : null;
    }else{
      el = document.querySelector('[data-ed="' + i + '"]');
    }
    if(!el) return;
    mark(el); curBlock = i;

    var hint =
      list != null   ? '1行が1項目になります。' :
      kind === 'list' ? '1行が1項目になります。' :
      kind === 'note' ? '1行が1つの文になります。' :
      kind === 'dl'   ? '「見出し ｜ 中身」の形で1行ずつ書きます。' :
                        '改行するとサイトでも改行されます。';
    hint += ' 太字は **はさむ**、リンクは [文字](URL)。';

    box = document.createElement('div');
    box.id = 'ed-box';
    box.innerHTML =
      '<div class="ed-ttl">' + (label || '文章') + 'を直す</div>' +
      '<textarea></textarea>' +
      '<p class="ed-hint"></p>' +
      '<div class="ed-foot">' +
        '<button class="ed-ok">この内容にする</button>' +
        '<button class="ed-ng">取消</button>' +
        '<span class="ed-keys">Ctrl+Enter で確定／Esc で取消</span>' +
      '</div>';
    box.querySelector('.ed-hint').textContent = hint;
    var ta = box.querySelector('textarea');
    ta.value = text || '';
    document.body.appendChild(box);

    var r = el.getBoundingClientRect();
    var sx = window.scrollX, sy = window.scrollY;
    var w = Math.min(Math.max(r.width, 320), document.documentElement.clientWidth - 24);
    box.style.width = w + 'px';
    var left = Math.min(Math.max(r.left + sx, 12),
                        document.documentElement.clientWidth - w - 12 + sx);
    box.style.left = left + 'px';
    box.style.top = (r.bottom + sy + 8) + 'px';
    var bh = box.getBoundingClientRect().height;
    if(r.bottom + 8 + bh > document.documentElement.clientHeight && r.top - bh - 8 > 0){
      box.style.top = (r.top + sy - bh - 8) + 'px';
    }

    ta.style.height = 'auto';
    ta.style.height = Math.min(ta.scrollHeight + 6, 320) + 'px';
    ta.focus();
    ta.setSelectionRange(ta.value.length, ta.value.length);

    box.querySelector('.ed-ng').onclick = function(){ mark(null); close() };
    box.querySelector('.ed-ok').onclick = apply;
    ta.addEventListener('keydown', function(e){
      if(e.key === 'Escape'){ e.preventDefault(); mark(null); close() }
      if(e.key === 'Enter' && (e.ctrlKey || e.metaKey)){ e.preventDefault(); apply() }
    });
    function apply(){
      var o = {act:'apply', block:i, text:ta.value, y:window.scrollY};
      if(part != null) o.part = part;
      if(list != null) o.list = list;
      send(o);
      tip('直しています…', 1500);
      close();
    }
  }

  /* 料金表のマス目を押したとき。中身が短いので入力欄も1行にする。 */
  function openPrice(el){
    close();
    mark(el);
    var key = el.getAttribute('data-price');
    var rc  = el.getAttribute('data-cell').split('|');
    box = document.createElement('div');
    box.id = 'ed-box';
    box.innerHTML =
      '<div class="ed-ttl">料金表のこのマスを直す</div>' +
      '<textarea rows="1"></textarea>' +
      '<p class="ed-hint">料金は「3,500円（税別）」のように書きます。' +
        '同じ表がほかのページにも出ていれば、そちらも一緒に直ります。</p>' +
      '<div class="ed-foot">' +
        '<button class="ed-ok">この内容にする</button>' +
        '<button class="ed-ng">取消</button>' +
        '<span class="ed-keys">Enter で確定／Esc で取消</span>' +
      '</div>';
    var ta = box.querySelector('textarea');
    ta.value = (el.textContent || '').trim();
    document.body.appendChild(box);

    var r = el.getBoundingClientRect();
    var w = Math.min(Math.max(r.width, 300), document.documentElement.clientWidth - 24);
    box.style.width = w + 'px';
    box.style.left = Math.min(Math.max(r.left + window.scrollX, 12),
                     document.documentElement.clientWidth - w - 12 + window.scrollX) + 'px';
    box.style.top = (r.bottom + window.scrollY + 8) + 'px';
    var bh = box.getBoundingClientRect().height;
    if(r.bottom + 8 + bh > document.documentElement.clientHeight && r.top - bh - 8 > 0){
      box.style.top = (r.top + window.scrollY - bh - 8) + 'px';
    }
    ta.style.minHeight = '0';
    ta.style.height = 'auto';
    ta.style.height = (ta.scrollHeight + 4) + 'px';
    ta.focus(); ta.setSelectionRange(ta.value.length, ta.value.length);

    function apply(){
      send({act:'price', key:key, r:+rc[0], c:+rc[1], text:ta.value.trim(), y:window.scrollY});
      tip('直しています…', 1500);
      close();
    }
    box.querySelector('.ed-ng').onclick = function(){ mark(null); close() };
    box.querySelector('.ed-ok').onclick = apply;
    ta.addEventListener('keydown', function(e){
      if(e.key === 'Escape'){ e.preventDefault(); mark(null); close() }
      if(e.key === 'Enter'){ e.preventDefault(); apply() }
    });
  }

  /* 文字を押したとき */
  document.addEventListener('click', function(e){
    if(box && box.contains(e.target)) return;
    var t = e.target;
    var el = (t && t.closest) ? t.closest('[data-price],[data-ed],[data-part],[data-list]') : null;
    if(!el){ if(box){ mark(null); close() } return }
    e.preventDefault(); e.stopPropagation();
    if(el.hasAttribute('data-price')) return openPrice(el);
    if(el.hasAttribute('data-inner')){
      var ins = el.querySelectorAll('[data-part],[data-list]');
      for(var k = 0; k < ins.length; k++){
        (function(x){ x.classList.add('ed-on'); setTimeout(function(){ x.classList.remove('ed-on') }, 900) })(ins[k]);
      }
      tip('直したい行を押してください', 2600);
      return;
    }
    mark(el);
    var blk = el.closest('[data-ed]');
    var o = {block: blk ? +blk.getAttribute('data-ed') : -1};
    if(el.hasAttribute('data-part')) o.part = +el.getAttribute('data-part');
    else if(el.hasAttribute('data-list')) o.list = +el.getAttribute('data-list');
    send(o);
  }, true);

  /* 管理画面からの合図 */
  window.addEventListener('message', function(e){
    var d = e.data || {};
    if(d.from !== 'ed-admin') return;
    if(d.act === 'edit') return openEditor(d.block, d.text, d.kind, d.label, d.part, d.list);
    if(d.act === 'info'){
      tip((d.label || 'この部分') + 'は' + (d.msg || '左の欄で直します'), 3400);
      return;
    }
    if(d.act === 'error'){ tip('直せませんでした: ' + (d.msg || ''), 4000); return }
    var el = document.querySelector('[data-ed="' + d.block + '"]');
    if(!el) return;
    mark(el);
    if(typeof d.y === 'number'){ window.scrollTo(0, d.y); return }
    var r = el.getBoundingClientRect();
    var want = r.top + window.scrollY - (window.innerHeight - r.height) / 2;
    var max = Math.max(0, document.documentElement.scrollHeight - window.innerHeight);
    window.scrollTo({top: Math.min(Math.max(0, want), max), behavior:'smooth'});
  });

  /* ページ全体の高さを管理画面に知らせる（枠の縦を合わせるため） */
  function deviceLook(){
    document.documentElement.classList.toggle('ed-touch', window.innerWidth <= 900);
  }
  function reportSize(){
    deviceLook();
    var b = document.body;
    send({act:'size', h: Math.max(document.documentElement.scrollHeight,
                                  b ? b.scrollHeight : 0)});
  }
  deviceLook();
  window.addEventListener('load', reportSize);
  window.addEventListener('resize', reportSize);
  document.addEventListener('DOMContentLoaded', function(){
    tip('直したいところを押すと、その場で書き直せます', 4200);
    reportSize();
    setTimeout(reportSize, 700);
    setTimeout(reportSize, 2000);
  });
})();
</script>
"""

# ページの中の「/care02/」のようなリンクを、管理画面の中のプレビューの住所に付け替える
_ABS_LINK = re.compile(r'\b(href|src|action)="(/(?!/)[^"]*)"')


@owner_required
@xframe_options_sameorigin  # 見え方の枠（iframe）に出すため
def hp_layout_view(request, page):
    """見え方の枠に出すページ（編集中の配置で作った _preview/<page>/index.html）。

    旧アプリの /draft/<page>/ と同じく、配るときだけ
      ・「/care02/」のような絶対リンクをプレビューの住所に付け替え
      ・文字を押して直すための script を差し込む
    ので、ファイルには何も残らない。写真やCSSは <base> で手元のサイトから読む。
    """
    if not repo.設定されているか():
        raise Http404()
    layout = repo.部品("layout")
    if page not in layout.PAGES:
        raise Http404()
    src = os.path.join(_preview_dir(), page, "index.html")
    if not os.path.isfile(src):
        # まだ見え方を作っていなければ、手元のサイトのページをそのまま出す
        src = layout.page_path(page)
        if not os.path.isfile(src):
            raise Http404()
    html = open(src, encoding="utf-8").read()
    files = reverse("manage:hp_preview_file", kwargs={"path": ""})
    html = _ABS_LINK.sub(lambda m: '%s="%s%s"' % (m.group(1), files.rstrip("/"), m.group(2)), html)
    base = '<base href="%s%s/">' % (files, page)
    html = re.sub(r"<head([^>]*)>", lambda m: "<head%s>%s" % (m.group(1), base), html, count=1)
    if 'data-ed="' in html:
        html = html.replace("</body>", EDIT_JS + "</body>", 1)
    res = HttpResponse(html, content_type="text/html; charset=utf-8")
    res["Cache-Control"] = "no-store"
    return res
