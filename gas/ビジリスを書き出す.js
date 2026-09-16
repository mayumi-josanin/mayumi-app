// ビジリスの記録を、データベース（server/apps/bijiris）へ取り込める形（JSON）で書き出す。
//
//   ビジリスの回答を書き出す()           … シート「回答一覧」＋アンケートごとのシート → ビジリス回答_*.json   【まゆみの GAS で動く】
//   ビジリスのアンケート定義を書き出す() … プロパティ SURVEYS_JSON → ビジリスアンケート_*.json                 【ビジリスの GAS で動かす】
//   ビジリスの顧客プロフィールを書き出す() … CUSTOMER_PROFILES_JSON ＋ CUSTOMER_MEMOS_JSON → ビジリス顧客_*.json  【ビジリスの GAS で動かす】
//   ビジリスの設定を書き出す()           … ADMIN_PREFERENCES_JSON ＋ TICKET_SURVEY_META_JSON ＋ TICKET_SURVEY_PROMPT → ビジリス設定_*.json 【ビジリスの GAS で動かす】
//
// **読むだけ。**スプレッドシートもプロパティも一切変えない。
//
// なぜ「まゆみ側」に置くのか（gas/回数券分析を書き出す.js と同じ理由）:
//   bijiris/gas は `.claspignore` が許可制で **Code.gs も push 対象**。push すると確認していない変更が
//   本番の動きを変えてしまう恐れがある。回答はビジリスのスプレッドシートにあるので、ここから
//   **openById で読むだけ**で済む。
//
// ただし、アンケート定義・顧客プロフィール・設定は**スプレッドシートには無く、ビジリス GAS の
// スクリプトプロパティにある**（docs/design/データ棚卸し.md「見落としやすい保管場所」）。
// まゆみの GAS からは読めない。bijiris/gas の読むだけの道具（設定の下見.js・データベースへ書き出す.js）にも
// プロパティを生のまま書き出す関数は無い（2026-09-16 に確かめた。設定の下見() は getPreferences_() を
// 通して Logger に出すだけ）。
//
// そこで、この3つの関数は**ビジリスの GAS で動かす前提**で書いてある。手順:
//   1. 院長の確認を取る（bijiris/gas を push すると Code.gs の未確認の変更も一緒に本番へ出るため）
//   2. このファイルを bijiris/gas/データベースへ書き出す.js の末尾へ写す（またはファイルごと bijiris/gas/ に置く）
//   3. `cd bijiris && clasp push`（**deploy はしない。**エディタから実行するだけなので deploy は要らない）
//   4. Apps Script のエディタで関数を選んで実行 → ログの URL から JSON を落とす
//   5. サーバーで python manage.py ビジリスのアンケートを取り込む … --下見 ／ ビジリスの顧客を取り込む … ／ ビジリスの設定を取り込む …
//   6. 取り込みが済んだら JSON を片付ける（ビジリスのアカウントの Drive に出るので、そちらでゴミ箱へ）
// まゆみの GAS で間違って実行しても、SURVEYS_JSON が無いのでその旨を出して止まる（何も書かない）。
//
// **Code.gs の「読む関数」を呼ばない。**loadSurveys_ / getPreferences_ / getCustomerProfiles_ / getCustomerMemos_ は
// 正規化した結果が元と違うとプロパティへ書き戻す（getCustomerProfiles_ は会員番号の採番までする）。
// ここは PropertiesService.getScriptProperties().getProperty(鍵) で生のまま読む。正規化はサーバー側の取り込みで行う。
//
// **お名前で会員に結びつけない。**ビジリスの記録はどれもお名前が鍵で会員番号を持たない。
// ここではお名前をそのまま持ち、何件が会員のお名前と一致するかだけ数えて出す。
//
// **秘密は書き出さない。**ANTHROPIC_API_KEY / ONESIGNAL_REST_API_KEY / TOKEN_SECRET / ADMIN_PASSWORD など。

var ビジリス書出_ビジリスID = '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E';
var ビジリス書出_回答一覧 = '回答一覧';

// MASTER_HEADERS（bijiris/gas/Code.gs 237行〜）と同じ並び。見出しの名前で拾う
var ビジリス書出_回答の列 = {
  submittedAt: ['送信日時'],
  id: ['回答ID'],
  surveyId: ['アンケートID'],
  surveyTitle: ['アンケート名'],
  customerClientId: ['端末ID'],
  customerName: ['お名前'],
  customerEmail: ['メールアドレス'],
  status: ['対応状況'],
  adminMemo: ['管理メモ'],
  answers: ['回答JSON'],
  files: ['写真JSON'],
  managedAt: ['管理更新日時']
};

// ---------------------------------------------------------------------------
// 回答（まゆみの GAS から openById で読む）
// ---------------------------------------------------------------------------
function ビジリスの回答を書き出す() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(ビジリス書出_ビジリスID);
  } catch (e) {
    Logger.log('■ ビジリスのスプレッドシートを開けませんでした。');
    Logger.log('  ' + e);
    Logger.log('  このGASを動かしているアカウントに、閲覧の権限が要ります。');
    return;
  }

  var sheet = ss.getSheetByName(ビジリス書出_回答一覧);
  if (!sheet) {
    Logger.log('■ 「' + ビジリス書出_回答一覧 + '」シートがありません。');
    return;
  }
  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) {
    Logger.log('■ 回答一覧が空です。書き出すものがありません。');
    return;
  }

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  var 位 = 掲載書出_列を引く_(v[0], ビジリス書出_回答の列);
  var 見つからない列 = Object.keys(位).filter(function (k) { return 位[k] < 0; });

  var 行 = [];
  var 空行 = 0;
  for (var r = 1; r < v.length; r += 1) {
    var id = 位.id >= 0 ? 掲載書出_文字_(v[r][位.id]) : '';
    // readMasterRows_ と同じ: 回答IDが無い行は行として数えない
    if (!id) { 空行 += 1; continue; }
    var o = { row: r + 1 };
    Object.keys(位).forEach(function (鍵) {
      var i = 位[鍵];
      if (i < 0) { o[鍵] = null; return; }
      var x = v[r][i];
      if (鍵 === 'submittedAt' || 鍵 === 'managedAt') o[鍵] = 掲載書出_日時_(x);
      else o[鍵] = 掲載書出_文字_(x);   // 回答JSON・写真JSON は文字列のまま（サーバー側で読む。読めなければ残す）
    });
    行.push(o);
  }

  // アンケートごとのシート（シート名 = アンケートのタイトル。ensureSurveySheet_）。回答IDで結ぶ
  var 題ごと = {};
  行.forEach(function (o) { if (o.surveyTitle) 題ごと[o.surveyTitle] = true; });
  var 別シートの索引 = {};
  var 無いシート = [];
  Object.keys(題ごと).forEach(function (題) {
    var sh = ss.getSheetByName(題);
    if (!sh || sh.getLastRow() < 2) { 無いシート.push(題); return; }
    var sv = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    var 見出し = sv[0].map(function (h) { return String(h == null ? '' : h).trim(); });
    var id列 = 見出し.indexOf('回答ID');
    if (id列 < 0) { 無いシート.push(題 + '（回答ID の列が無い）'); return; }
    for (var k = 1; k < sv.length; k += 1) {
      var rid = 掲載書出_文字_(sv[k][id列]);
      if (!rid) continue;
      var values = {};
      見出し.forEach(function (h, j) {
        if (!h) return;
        var x = sv[k][j];
        values[h] = (x instanceof Date) ? 掲載書出_日時_(x) : 掲載書出_文字_(x);
      });
      別シートの索引[rid] = { row: k + 1, values: values };
    }
  });
  var 別シートに無い = 0;
  行.forEach(function (o) {
    o.surveySheet = 別シートの索引[o.id] || null;
    if (!o.surveySheet) 別シートに無い += 1;
  });

  // 写真の枚数（Drive のファイルID）。回答JSONの answers[].files[] を数える。無ければ写真JSON
  var 写真 = 0, 写真ID = {}, 読めない = 0;
  行.forEach(function (o) {
    var answers = ビジリス書出_JSON_(o.answers, null);
    var files = ビジリス書出_JSON_(o.files, null);
    if ((o.answers && answers === null) || (o.files && files === null)) 読めない += 1;
    var 集め = [];
    (Array.isArray(answers) ? answers : []).forEach(function (a) {
      if (a && Array.isArray(a.files)) 集め = 集め.concat(a.files);
    });
    if (!集め.length && Array.isArray(files)) 集め = files;
    集め.forEach(function (f) {
      if (f && f.fileId) { 写真 += 1; 写真ID[String(f.fileId)] = true; }
    });
  });

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    シートの行数: v.length - 1,
    responses: 行
  }, null, 2);
  var 名前 = 'ビジリス回答_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024 * 10) / 10 + 'KB');
  Logger.log('');
  Logger.log('■ ビジリスの回答: ' + 行.length + '件（回答一覧は' + (v.length - 1) + '行）');
  if (空行) Logger.log('    回答IDが無い行: ' + 空行 + '件');
  var 状態 = {};
  行.forEach(function (x) { var s = x.status || '（空）'; 状態[s] = (状態[s] || 0) + 1; });
  Logger.log('    対応状況: ' + Object.keys(状態).map(function (k) { return k + ' ' + 状態[k] + '件'; }).join(' / '));
  var 題数 = {};
  行.forEach(function (x) { var t = x.surveyTitle || '（空）'; 題数[t] = (題数[t] || 0) + 1; });
  Logger.log('    アンケート: ' + Object.keys(題数).map(function (k) { return k + ' ' + 題数[k] + '件'; }).join(' / '));
  Logger.log('    アンケートごとのシートに同じ回答IDがある: ' + (行.length - 別シートに無い) + '件 / 無い: ' + 別シートに無い + '件');
  if (無いシート.length) Logger.log('    **シートが見つからないアンケート: ' + 無いシート.join('・') + '**');
  Logger.log('    写真: のべ ' + 写真 + '枚（ファイルIDの種類 ' + Object.keys(写真ID).length + '）');
  if (読めない) Logger.log('    **回答JSON か 写真JSON が読めない回答: ' + 読めない + '件**（文字列のまま入れてあります）');
  ビジリス書出_お名前の一致を数える_(行.map(function (x) { return x.customerName; }));
  if (見つからない列.length) Logger.log('  **見つからない列: ' + 見つからない列.join('・') + '**');
  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ **お客様のお名前・回答・写真の場所が入っています。**取り込みが済んだら 書き出しJSONを片付ける() で消してください。');
}

// ---------------------------------------------------------------------------
// プロパティにあるもの（ビジリスの GAS で動かす。まゆみの GAS では止まる）
// ---------------------------------------------------------------------------
function ビジリス書出_プロパティを読む_(鍵) {
  // **getProperty だけ。**Code.gs の load*/get* は書き戻すので呼ばない
  return PropertiesService.getScriptProperties().getProperty(鍵);
}

function ビジリス書出_ここはビジリスか_() {
  if (ビジリス書出_プロパティを読む_('SURVEYS_JSON')) return true;
  Logger.log('■ この GAS には SURVEYS_JSON がありません。ここはビジリスの GAS ではないようです。');
  Logger.log('  アンケート定義・顧客プロフィール・設定はビジリス GAS のスクリプトプロパティにあります。');
  Logger.log('  このファイルを bijiris/gas/ に写して（院長の確認後に clasp push。deploy は不要）、そちらのエディタから実行してください。');
  Logger.log('  何も書いていません。');
  return false;
}

function ビジリスのアンケート定義を書き出す() {
  if (!ビジリス書出_ここはビジリスか_()) return;
  var 生 = ビジリス書出_プロパティを読む_('SURVEYS_JSON');
  var surveys = ビジリス書出_JSON_(生, null);
  if (!Array.isArray(surveys)) {
    Logger.log('■ SURVEYS_JSON が読めません（配列ではない）。先頭: ' + String(生).slice(0, 80));
    return;
  }
  var file = ビジリス書出_置く_('ビジリスアンケート', { surveys: surveys });
  Logger.log('■ アンケート定義: ' + surveys.length + '本');
  surveys.forEach(function (s) {
    var q = Array.isArray(s.questions) ? s.questions : [];
    var 写真 = q.filter(function (x) { return x && x.type === 'photo'; }).length;
    Logger.log('    - ' + (s.title || '（無題）') + '（' + (s.id || '?') + '）: ' + (s.status || '?') +
      ' / 設問 ' + q.length + '問（写真 ' + 写真 + '問）');
  });
  Logger.log('  ※ 読むだけです。プロパティは変えていません。');
  Logger.log('  → python manage.py ビジリスのアンケートを取り込む ' + file.getName() + ' --下見');
}

function ビジリスの顧客プロフィールを書き出す() {
  if (!ビジリス書出_ここはビジリスか_()) return;
  var profiles = ビジリス書出_JSON_(ビジリス書出_プロパティを読む_('CUSTOMER_PROFILES_JSON'), {}) || {};
  var memos = ビジリス書出_JSON_(ビジリス書出_プロパティを読む_('CUSTOMER_MEMOS_JSON'), {}) || {};
  var file = ビジリス書出_置く_('ビジリス顧客', { profiles: profiles, memos: memos });
  var 名前 = Object.keys(profiles);
  Logger.log('■ 顧客プロフィール: ' + 名前.length + '名 / メモ: ' + Object.keys(memos).length + '名');
  Logger.log('    パスコード設定済み: ' + 名前.filter(function (n) { var p = profiles[n] || {}; return p.passcodeHash && p.passcodeSalt; }).length + '名');
  Logger.log('    回数券カードあり: ' + 名前.filter(function (n) { return profiles[n] && profiles[n].activeTicketCard; }).length + '名');
  Logger.log('    特典の受け取りあり: ' + 名前.filter(function (n) { return profiles[n] && profiles[n].rewardRedemptions; }).length + '名');
  Logger.log('    GAS の会員番号あり: ' + 名前.filter(function (n) { return profiles[n] && profiles[n].memberNumber; }).length + '名');
  Logger.log('  ※ 読むだけです。プロパティは変えていません（getCustomerProfiles_ は呼んでいません）。');
  Logger.log('  ※ **お客様のお名前・パスコードのハッシュ・メモが入っています。**取り込んだら必ず消してください。');
  Logger.log('  → python manage.py ビジリスの顧客を取り込む ' + file.getName() + ' --下見');
}

// 書き出す鍵はこの3つだけ。ANTHROPIC_API_KEY / ONESIGNAL_* / TOKEN_SECRET / ADMIN_* は**出さない**
function ビジリスの設定を書き出す() {
  if (!ビジリス書出_ここはビジリスか_()) return;
  var preferences = ビジリス書出_JSON_(ビジリス書出_プロパティを読む_('ADMIN_PREFERENCES_JSON'), {}) || {};
  var meta = ビジリス書出_JSON_(ビジリス書出_プロパティを読む_('TICKET_SURVEY_META_JSON'), {}) || {};
  var prompt = String(ビジリス書出_プロパティを読む_('TICKET_SURVEY_PROMPT') || '');
  var file = ビジリス書出_置く_('ビジリス設定', { preferences: preferences, ticket_meta: meta, ticket_prompt: prompt });
  Logger.log('■ 設定: ADMIN_PREFERENCES_JSON ' + Object.keys(preferences).length + '項目 / TICKET_SURVEY_META_JSON ' +
    Object.keys(meta).length + '項目 / TICKET_SURVEY_PROMPT ' + (prompt ? prompt.length + '文字' : '空（既定を使う）'));
  Logger.log('    通知メール: ' + (preferences.notificationEmail || '（空）') +
    ' / 特典（節目）: ' + JSON.stringify(preferences.milestoneRewardConfig || null));
  Logger.log('  ※ 読むだけです。秘密（API キーなど）は出していません。');
  Logger.log('  → python manage.py ビジリスの設定を取り込む ' + file.getName() + ' --下見');
}

// ---------------------------------------------------------------------------
// 小道具
// ---------------------------------------------------------------------------
function ビジリス書出_JSON_(値, 既定) {
  if (値 === '' || 値 === null || 値 === undefined) return 既定;
  if (typeof 値 !== 'string') return 値;
  try { return JSON.parse(値); } catch (e) { return null; }
}

function ビジリス書出_置く_(頭, 中身) {
  var 名前 = 頭 + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, JSON.stringify(Object.assign({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX")
  }, 中身), null, 2), MimeType.PLAIN_TEXT);
  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024 * 10) / 10 + 'KB');
  Logger.log('');
  return file;
}

// 会員のお名前と一致するかだけを数える。**結びつけはしない。**まゆみの会員データが無ければ飛ばす
function ビジリス書出_お名前の一致を数える_(名前ら) {
  var 会員名 = {};
  try {
    var us = getOrCreateSpreadsheet().getSheetByName(SHEETS.USERS);
    if (us && us.getLastRow() > 1) {
      var uv = us.getRange(1, 1, us.getLastRow(), us.getLastColumn()).getValues();
      var 名列 = -1;
      for (var j = 0; j < uv[0].length; j += 1) if (String(uv[0][j]).trim() === '氏名') { 名列 = j; break; }
      if (名列 >= 0) for (var k = 1; k < uv.length; k += 1) {
        var n = String(uv[k][名列] || '').trim();
        if (n) 会員名[n] = (会員名[n] || 0) + 1;
      }
    }
  } catch (e) {
    Logger.log('  （会員のお名前を読めませんでした: ' + e + '）');
    return;
  }
  var 種類 = {};
  名前ら.forEach(function (n) { if (n) 種類[n] = true; });
  var 一致 = 0, 複数 = [], 無し = 0;
  Object.keys(種類).forEach(function (n) {
    var c = 会員名[n] || 0;
    if (c === 1) 一致 += 1; else if (c > 1) 複数.push(n + '（会員に' + c + '名）'); else 無し += 1;
  });
  Logger.log('  ● 会員のお名前との一致（**数えるだけ。結びつけていません**）: ' + Object.keys(種類).length + '名のうち ちょうど1名 ' + 一致 +
    ' / 複数 ' + 複数.length + (複数.length ? '（' + 複数.join('・') + '）' : '') + ' / 無し ' + 無し);
}
