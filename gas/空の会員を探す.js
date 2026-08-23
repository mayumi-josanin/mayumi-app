// お名前が空の会員の行を探す道具。**読むだけ。**
//
//   空の会員を探す() … お名前が空の行と、指定した会員IDの行を出す
//
// ## なぜ要るのか
//
// `updateUser` は、`name` を**そもそも送らなければ**検査を通り抜けて、
// お名前も電話番号も生年月日も空の行を作っていた。
// この窓口は PUBLIC_ACTIONS に入っていて**合鍵なしで誰でも呼べる。**
//
// 空の行は、復元の条件（生年月日の完全一致）を満たしようがないので、
// **その会員IDは永久にアプリへ入れない。**
// 2026年8月に「入れないお客様が21名」いた原因も、おそらくこれ。
//
// 2026-08-24 に @240 で塞いだ（handleUpdateUser に1行）。
// **塞ぐ前に作られた行が残っていないか**を、これで見る。
//
// お名前だけで判断しない。**その行の全項目を見てから**片付けを決めること
// （テスト会員を片付ける.js の教訓）。

// ついでに探したい会員ID。空の配列でよい。
var 空探し_この会員IDも = ['MYM-GAS-CHECK-DELETEME'];

function 空の会員を探す() {
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet());
  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) {
    Logger.log('■ 会員データに行がありません。');
    return;
  }

  var 値 = sheet.getRange(2, 1, 最終行 - 1, USER_HEADERS.length).getValues();
  var 空の行 = [];
  var 指定の行 = [];

  値.forEach(function (row, i) {
    var 行番号 = i + 2;
    var 会員ID = String(row[USER_COL.MEMBER_ID - 1] || '').trim();
    var 氏名 = String(row[USER_COL.NAME - 1] || '').trim();

    if (会員ID && !氏名) 空の行.push({ 行番号: 行番号, row: row });
    if (会員ID && 空探し_この会員IDも.indexOf(会員ID) !== -1) {
      指定の行.push({ 行番号: 行番号, row: row });
    }
  });

  Logger.log('■ 会員データ: ' + (最終行 - 1) + '行');
  Logger.log('');
  Logger.log('■ **お名前が空の行: ' + 空の行.length + '件**');
  空の行.forEach(function (x) { 空探し_中身を出す_(x); });
  if (!空の行.length) Logger.log('    ありません。');
  Logger.log('');

  if (空探し_この会員IDも.length) {
    Logger.log('■ 指定した会員IDの行: ' + 指定の行.length + '件');
    Logger.log('    （探した会員ID: ' + 空探し_この会員IDも.join(' / ') + '）');
    指定の行.forEach(function (x) { 空探し_中身を出す_(x); });
    if (!指定の行.length) Logger.log('    ありません。**作られていません。**');
  }
}

function 空探し_中身を出す_(x) {
  Logger.log('');
  Logger.log('  ' + x.行番号 + '行目');
  USER_HEADERS.forEach(function (見出し, i) {
    var v = x.row[i];
    var s = (v === '' || v === null || v === undefined) ? '（空）' : String(v);
    // 端末や履歴のJSONは長いので、頭だけ出す。
    if (s.length > 60) s = s.slice(0, 60) + ' …（以下略）';
    Logger.log('      ' + 見出し + ': ' + s);
  });
}
