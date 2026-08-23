// 「知ったきっかけアンケート」まわりの、**会員ごとの印**を書き出す道具。
//
//   会員ごとの印の下見()    … 何件あるかを見るだけ
//   会員ごとの印を書き出す() … JSONにしてドライブへ置く
//
// **どちらも読むだけ。**スクリプトプロパティは一切変えない。
//
// ## なぜ、これが要るのか
//
// この4つは**表ではなく、会員ごとのスクリプトプロパティ**に入っている。
//
//   SURVEY_ANSWERED:<会員ID>       … 回答した事実。アプリでボタンを隠す判定
//   SURVEY_STAMP:<会員ID>          … お礼スタンプを付けた記録（二重付与の防止）
//   SURVEY_STAMP_PENDING:<会員ID>  … カードが満杯で付けられず、空き待ちにした記録
//   REWARD_ADMIN_SET:<会員ID>      … 管理者がスタンプを直接編集した時刻
//
// 設定の書き出し（設定を書き出す.js）は「移すと決めた2つだけを名指しで」
// 書き出す作りなので、**これらは残っていた**（2026-08-23に判明）。
//
// **移し忘れると、お礼スタンプが二重に付く。**
// SURVEY_STAMP が空に見えるので、「まだ付けていない」と判断してしまう。
//
// ## 書き出さないもの
//
// **この4つの頭文字で始まる鍵だけを書き出す。**
// 「全部書き出して後で選ぶ」はしない。ADMIN_TOKEN_SECRET のような秘密が
// 同じ場所に入っているので、書き出した時点で漏れる。

var 印書出_移すもの = [
  { 頭: 'SURVEY_ANSWERED:',     名: 'surveyAnsweredAt' },
  { 頭: 'SURVEY_STAMP:',        名: 'surveyStampGrantedAt' },
  { 頭: 'SURVEY_STAMP_PENDING:', 名: 'surveyStampPendingAt' },
  { 頭: 'REWARD_ADMIN_SET:',    名: 'rewardAdminSetAt' }
];

function 会員ごとの印を書き出す() {
  var r = 印書出_集める_();
  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    件数: r.行.length,
    members: r.行
  }, null, 2);

  var 名前 = '会員ごとの印_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('');
  印書出_内訳を出す_(r);
  Logger.log('');
  Logger.log('  ※ 秘密は入っていません（この4種の頭文字の鍵だけを書き出しています）。');
  Logger.log('     それでも、取り込みが済んだら消してください。');
}

function 会員ごとの印の下見() {
  var r = 印書出_集める_();
  Logger.log('■ 会員ごとの印');
  Logger.log('');
  印書出_内訳を出す_(r);
  Logger.log('');
  Logger.log('  ※ 「お礼スタンプ付与」が移っていないと、**二重に付きます。**');
  Logger.log('     いま残っている全部の鍵の数: ' + r.鍵の総数 + '個');
  Logger.log('     （そのうち、上の4種にあてはまるものだけを書き出します）');
}

function 印書出_内訳を出す_(r) {
  Logger.log('  会員: ' + r.行.length + '名ぶんの印があります');
  印書出_移すもの.forEach(function (m) {
    Logger.log('    ' + m.頭 + ' … ' + r.数[m.名] + '件');
  });
}

function 印書出_集める_() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var 鍵たち = Object.keys(props);

  var 会員ごと = {};
  var 数 = {};
  印書出_移すもの.forEach(function (m) { 数[m.名] = 0; });

  鍵たち.forEach(function (鍵) {
    for (var i = 0; i < 印書出_移すもの.length; i++) {
      var m = 印書出_移すもの[i];
      if (鍵.indexOf(m.頭) !== 0) continue;

      var 会員ID = 鍵.slice(m.頭.length).trim();
      if (!会員ID) return;

      if (!会員ごと[会員ID]) 会員ごと[会員ID] = { memberId: 会員ID };
      会員ごと[会員ID][m.名] = String(props[鍵] || '');
      数[m.名]++;
      return;
    }
  });

  var 行 = Object.keys(会員ごと).map(function (k) { return 会員ごと[k]; });
  return { 行: 行, 数: 数, 鍵の総数: 鍵たち.length };
}
