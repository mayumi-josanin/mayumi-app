// GAS からサーバー（自宅PCの PostgreSQL）へ、表ごとに渡すための土台。
//
//   サーバーへの道を確かめる()   … いま届くか。**何も書きません**
//   渡している表を見る()         … いまどの表を転送しているか
//
// ─────────────────────────────────────────────────────────
// なぜこの形にしたのか
//
// 当初は「お客様アプリの通信先をサーバーに向ける」つもりだった。
// **それでは管理アプリが取り残される**（2026-09-05 に判明）。
//
//   お客様アプリ → サーバー        会員・お知らせを読み書き
//   管理アプリ   → GAS → シート    同じものを読み書き
//
// 同じ表に「正」が2つできる。院長がお知らせを投稿してもお客様に出ず、
// 受付が押したスタンプもアプリでは増えない。
// 実際、9/5 の時点で**お知らせが 91件（本番）と 88件（サーバー）**にずれていた。
//
// だから**窓口はGASのまま**にして、GASの中身だけを表ごとに移す。
// 両方のアプリがGASを見ている限り、正はいつも1つ。
//
// ─────────────────────────────────────────────────────────
// 使うスクリプトプロパティ
//
//   SERVER_TABLES    渡す表をカンマ区切りで。**空なら一切渡さない**
//                    例  news        … お知らせだけ
//                        news,items  … お知らせと商品
//   SERVER_BASE_URL  省略時は下の既定値
//   SERVER_API_KEY   サーバーの API_KEY と同じ値。
//                    **お客様向けの窓口には要らない**（サーバー側が
//                    公開アクションとして通すため）。管理者向けを
//                    渡すようになったときに要る。
//
// **戻すときはデプロイが要らない。**SERVER_TABLES を空にするだけで、
// その場でシートに戻る。切り替えの日に、いちばん速く戻せる道を残しておく。

var 渡す_既定のURL = 'https://mayumi-api.tail8efe0d.ts.net/api';
var 渡す_待つミリ秒 = 20000;

function 渡す_設定_(鍵, 既定) {
  var v = PropertiesService.getScriptProperties().getProperty(鍵);
  return v === null || v === undefined || v === '' ? (既定 || '') : String(v);
}

function 渡す_URL_() {
  return 渡す_設定_('SERVER_BASE_URL', 渡す_既定のURL).replace(/\/+$/, '');
}

/**
 * その表を、いまサーバーへ渡しているか。
 *
 * **既定は「渡さない」。**プロパティを設定して初めて渡り始める。
 * 逆にしてしまうと、デプロイした瞬間に全部の表が移ってしまう。
 */
function サーバーへ渡すか_(表名) {
  var 一覧 = 渡す_設定_('SERVER_TABLES', '');
  if (!一覧) return false;
  var 並び = 一覧.split(',').map(function (x) { return String(x || '').trim(); });
  return 並び.indexOf(String(表名 || '').trim()) !== -1;
}

function 渡す_ヘッダ_() {
  var h = { 'Content-Type': 'application/json' };
  var 鍵 = 渡す_設定_('SERVER_API_KEY', '');
  if (鍵) h['X-Api-Key'] = 鍵;
  return h;
}

/**
 * サーバーから読む。**失敗したら null を返す。**
 *
 * 呼び出し側は null のときシートへ落ちること。
 * 真っ白な画面を出すより、少し古い中身を出すほうがお客様の害が小さい。
 */
function サーバーから読む_(action, 引数) {
  try {
    var url = 渡す_URL_() + '?action=' + encodeURIComponent(action) +
      (引数 ? '&data=' + encodeURIComponent(JSON.stringify(引数)) : '');
    var res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: 渡す_ヘッダ_(),
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
      escaping: false
    });
    if (res.getResponseCode() !== 200) {
      渡す_書き留める_('読み', action, 'HTTP ' + res.getResponseCode());
      return null;
    }
    var 中 = JSON.parse(res.getContentText());
    if (!中 || 中.notImplemented) {
      渡す_書き留める_('読み', action, 'サーバーにまだ無い');
      return null;
    }
    return 中;
  } catch (e) {
    渡す_書き留める_('読み', action, String(e));
    return null;
  }
}

/**
 * サーバーへ書く。**失敗しても、シートには書かない。**
 *
 * 読みと違って、落ちる先を作ってはいけない。
 * サーバーに書けなかったぶんをシートに書くと、**どちらが正か分からなくなる。**
 * 分からなくなるくらいなら、その場で失敗として返し、もう一度やっていただく。
 *
 * 返り値: サーバーの答え。届かなければ null（呼び出し側でエラーにする）
 */
function サーバーへ書く_(中身) {
  try {
    var res = UrlFetchApp.fetch(渡す_URL_(), {
      method: 'post',
      headers: 渡す_ヘッダ_(),
      payload: JSON.stringify(中身 || {}),
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      escaping: false
    });
    if (res.getResponseCode() !== 200) {
      渡す_書き留める_('書き', 中身 && 中身.type, 'HTTP ' + res.getResponseCode());
      return null;
    }
    var 答 = JSON.parse(res.getContentText());
    // **まだ無い窓口は「書けた」と見なさない。**
    // notImplemented を成功として扱うと、書けていないのに書けたことになる。
    if (!答 || 答.notImplemented) {
      渡す_書き留める_('書き', 中身 && 中身.type, 'サーバーにまだ無い');
      return null;
    }
    return 答;
  } catch (e) {
    渡す_書き留める_('書き', 中身 && 中身.type, String(e));
    return null;
  }
}

/** うまくいかなかったことだけを残す。うまくいった分まで残すと埋もれる。 */
function 渡す_書き留める_(向き, 名, 理由) {
  try {
    Logger.log('[サーバーへ渡す] ' + 向き + ' ' + (名 || '?') + ' → ' + 理由);
  } catch (e) { /* ログが取れなくても本筋を止めない */ }
}

// ═════════════════════════════════════════════════════════
// エディタから実行して確かめるもの（読むだけ）
// ═════════════════════════════════════════════════════════

/** いま届くか、どの表を渡しているかを見る。**何も書きません。** */
function サーバーへの道を確かめる() {
  Logger.log('■ サーバーへの道');
  Logger.log('');
  Logger.log('  行き先        : ' + 渡す_URL_());
  Logger.log('  合鍵          : ' + (渡す_設定_('SERVER_API_KEY', '') ? '設定あり' : '未設定（お客様向けの窓口には不要）'));
  var 一覧 = 渡す_設定_('SERVER_TABLES', '');
  Logger.log('  渡している表  : ' + (一覧 || '**なし**（すべてシートのまま）'));
  Logger.log('');

  var t = new Date().getTime();
  var 中 = サーバーから読む_('getNews');
  var かかった = new Date().getTime() - t;

  if (!中) {
    Logger.log('  **届きません。**上の行き先を確かめてください。');
    return;
  }
  var 件数 = (中.news || 中.blogs || []).length;
  Logger.log('  届きました（' + かかった + 'ミリ秒）');
  Logger.log('  サーバーのお知らせ: ' + 件数 + '件');
  Logger.log('');
  Logger.log('  ※ これは見るだけです。渡す設定は変わっていません。');
}

/** どの表を渡しているかだけを見る。 */
function 渡している表を見る() {
  var 一覧 = 渡す_設定_('SERVER_TABLES', '');
  Logger.log('■ いまサーバーへ渡している表');
  Logger.log('');
  if (!一覧) {
    Logger.log('  **ありません。**すべてスプレッドシートのままです。');
  } else {
    一覧.split(',').forEach(function (x) {
      if (String(x || '').trim()) Logger.log('  ・' + String(x).trim());
    });
  }
  Logger.log('');
  Logger.log('  変えるには スクリプトプロパティ SERVER_TABLES を書き換えます。');
  Logger.log('  **空にすれば、その場でシートに戻ります（デプロイ不要）。**');
}
