// ポチコのお客様と突き合わせるためだけに、会員データの**3列だけ**を書き出す道具。
//
//   突き合わせ用に会員を書き出す()   … ドライブへ置く
//
// **読むだけ。**スプレッドシートは一切変えない。
//
// ─────────────────────────────────────────────────────────
// なぜ『会員データを書き出す』を使わないのか
//
// あちらは移行のための道具で、**パスコードを平文で書き出す。**
// 「ポチコから何名のLINEIDが取れるか」を数えるだけなら、
// 氏名・フリガナ・電話番号があれば足りる。
// **要らないものを外に出さない。**
//
// 出す列（会員データシート）
//   1  会員ID     MYM-####
//   3  氏名        ポチコの本名と突き合わせる（**弱い鍵**。同姓同名がある）
//   4  フリガナ
//   5  電話番号    **これが確実な鍵。**ポチコのフォーム回答に184名分ある
//  34  LINEユーザーID   すでに入っている方は数から外すため
//
// 消した会員（21列目に印がある行）は出さない。
//
// 済んだらドライブからも手元からも消すこと。

// **getOrCreateSpreadsheet() を呼ばない。**
// あれは5番目で「見つからなければ新しいスプレッドシートを作る」。
// 読むだけの道具から呼ぶと、本番とは別の空の表ができる。
// 写真フォルダで同じ罠を踏んでいる（CLAUDE.md 2026-08-25）。
function 突き合わせ_表を開く_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
    || FALLBACK_SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);   // 無ければ例外。**作らない**
}

function 突き合わせ用に会員を書き出す() {
  var sheet = 突き合わせ_表を開く_().getSheetByName('会員データ');
  if (!sheet) { Logger.log('会員データのシートが見つかりません'); return; }

  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) { Logger.log('会員がいません'); return; }
  var 値 = sheet.getRange(2, 1, 最終行 - 1, 34).getValues();

  var 会員 = [];
  var 消した = 0;
  for (var i = 0; i < 値.length; i++) {
    var row = 値[i];
    var id = String(row[0] || '').trim();
    if (!id) continue;
    // 21列目に印がある行は、消された会員
    if (String(row[20] || '').trim()) { 消した++; continue; }
    会員.push({
      id: id,
      名前: String(row[2] || '').trim(),
      カナ: String(row[3] || '').trim(),
      電話: String(row[4] || '').trim(),
      LINEID: String(row[33] || '').trim()
    });
  }

  var 名前 = '突き合わせ用会員_' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, JSON.stringify({ 件数: 会員.length, 会員: 会員 }, null, 1),
    MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('');
  Logger.log('  会員 ' + 会員.length + '名（消した会員 ' + 消した + '件は除いた）');
  Logger.log('  電話あり  ' + 会員.filter(function (x) { return x.電話; }).length + '名');
  Logger.log('  LINEID済 ' + 会員.filter(function (x) { return x.LINEID; }).length + '名');
  Logger.log('');
  Logger.log('  ※ 数え終わったら、ドライブからも手元からも消してください。');
}
