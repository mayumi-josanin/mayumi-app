// 計測写真を「アプリ経由でだけ」表示するための共通の読み込み。
//
// ## なぜ要るのか
//
// これまで写真は Drive に「リンクを知っている全員が閲覧可」で保存し、
// 画面は `drive.google.com/thumbnail?id=…` を直接 `<img src>` に入れていた。
//
// **URLが一度でも漏れれば、誰でもお客様のお体の写真を見られる。**
//
// これからは `?action=photoData` を通す。GAS 側で
//
//     管理者の札 → すべての写真
//     お客様の札 → **自分の回答に含まれる写真だけ**
//
// と確かめてから中身を返す（`Code.gs` の `getPhotoData_`）。
//
// ## 呼び出し側を書き換えないための作り
//
// 写真は19か所で `<img src="…">` として組み立てられている。
// **そこを全部書き換えるのは危ない**（見落とすと写真が出なくなる）。
//
// そこで、
//
//   1. 差し込む側は「仮の絵」を返す。URLの末尾に `#fid=<ファイルID>` を付ける
//   2. ここが画面を見張り、`#fid=` の付いた `<img>` を見つけたら中身を取りに行く
//   3. 取れたら `src` を実物に差し替える
//
// **呼び出し側は1か所も変えなくてよい。**
//
// ## 気をつけていること
//
//   ・同じ写真を何度も取りに行かない（ファイルIDごとに覚える）
//   ・取れなかった写真は**印を出す**。黙って空白にしない
//   ・`<a href>`（開く・保存）も同じ仕組みで差し替える

window.BijirisPhotoLoader = (() => {
  "use strict";

  // SVG を data URI にする。
  //
  // **btoa は使わない。**Latin1 しか扱えず、日本語を入れると
  // InvalidCharacterError で落ちる（2026-08-25 に実際に落とした）。
  // percent-encoding なら日本語のまま入れられる。
  function 絵にする(svg) {
    return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
  }

  // 読み込み中に見せる絵。薄いグレーの四角。
  const 仮の絵 = 絵にする(
    '<svg xmlns="http://www.w3.org/2000/svg" width="240" height="240">' +
      '<rect width="240" height="240" fill="#efeae2"/>' +
      '<circle cx="120" cy="120" r="26" fill="none" stroke="#c9bfb2" stroke-width="6" ' +
      'stroke-dasharray="120" stroke-dashoffset="40"/></svg>',
  );

  // 取れなかったときの絵。**黙って空白にしない。**
  const だめだった絵 = 絵にする(
    '<svg xmlns="http://www.w3.org/2000/svg" width="240" height="240">' +
      '<rect width="240" height="240" fill="#f6efe9"/>' +
      '<text x="120" y="116" text-anchor="middle" font-size="15" fill="#a4795f">' +
      "写真を読み込めません</text>" +
      '<text x="120" y="140" text-anchor="middle" font-size="12" fill="#b9a897">' +
      "時間をおいてお試しください</text></svg>",
  );

  const 覚えたもの = new Map();   // ファイルID → data URI
  const 取りに行き中 = new Map(); // ファイルID → Promise
  let 札を出す = null;            // 管理者のとき、札を返す関数

  function 印を付ける(ファイルID) {
    const id = String(ファイルID || "").trim();
    return id ? 仮の絵 + "#fid=" + encodeURIComponent(id) : "";
  }

  function 印を読む(値) {
    const m = String(値 || "").match(/#fid=([^&#]+)$/);
    if (!m) return "";
    try {
      return decodeURIComponent(m[1]);
    } catch {
      return m[1];
    }
  }

  async function 取ってくる(ファイルID) {
    if (覚えたもの.has(ファイルID)) return 覚えたもの.get(ファイルID);
    if (取りに行き中.has(ファイルID)) return 取りに行き中.get(ファイルID);

    const api = window.MayumiSurveyApi;
    if (!api) return "";

    const 問い合わせ = { fileId: ファイルID };
    // 管理者のときは自分の札を渡す。お客様のときは通信層が自動で付ける。
    const 札 = 札を出す ? String(札を出す() || "") : "";
    if (札) 問い合わせ.token = 札;

    const 約束 = (async () => {
      try {
        const r = await api.request("/api/photo-data", { query: 問い合わせ });
        const 種類 = String((r && r.mimeType) || "image/jpeg");
        const 中身 = String((r && r.base64) || "");
        if (!中身) return "";
        const url = "data:" + 種類 + ";base64," + 中身;
        覚えたもの.set(ファイルID, url);
        return url;
      } catch (e) {
        // 権限が無い・消された・通信が切れた、のいずれか。
        // **覚えない。**次に開いたときに、もう一度試せるようにする。
        console.log("[写真] 読み込めませんでした", ファイルID, e && e.message);
        return "";
      } finally {
        取りに行き中.delete(ファイルID);
      }
    })();

    取りに行き中.set(ファイルID, 約束);
    return 約束;
  }

  async function 差し替える(要素) {
    const 属性 = 要素.tagName === "A" ? "href" : "src";
    const ファイルID = 印を読む(要素.getAttribute(属性));
    if (!ファイルID) return;

    // 二度手間を防ぐ。
    if (要素.dataset.写真読み込み === "済" || 要素.dataset.写真読み込み === "中") return;
    要素.dataset.写真読み込み = "中";

    const url = await 取ってくる(ファイルID);
    if (url) {
      要素.setAttribute(属性, url);
      要素.dataset.写真読み込み = "済";
    } else {
      if (属性 === "src") 要素.setAttribute("src", だめだった絵);
      要素.dataset.写真読み込み = "失敗";
    }
  }

  function 見回る(根 = document) {
    const 対象 = 根.querySelectorAll
      ? 根.querySelectorAll('img[src*="#fid="], a[href*="#fid="]')
      : [];
    対象.forEach(差し替える);
  }

  let 待ち = null;
  function あとで見回る() {
    if (待ち) return;
    待ち = setTimeout(() => {
      待ち = null;
      見回る();
    }, 30);
  }

  function 見張りを始める(options = {}) {
    札を出す = typeof options.札 === "function" ? options.札 : null;

    見回る();
    // 画面が描き直されるたびに拾う。**呼び出し側に手を入れないため。**
    const 見張り = new MutationObserver(あとで見回る);
    見張り.observe(document.documentElement, { childList: true, subtree: true });
  }

  return {
    印を付ける,
    見回る,
    見張りを始める,
    仮の絵,
    // ログアウトのときに呼ぶ。**他の方の写真を残さない。**
    忘れる() {
      覚えたもの.clear();
      取りに行き中.clear();
    },
  };
})();
