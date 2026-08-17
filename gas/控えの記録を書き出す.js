// 控えの記録（BACKUP_LOG）を、データベースへ取り込める形（JSON）で書き出す。
//
//   控えの記録を書き出す()   … JSONにしてドライブへ置く
//
// **読むだけ。**元のスプレッドシートは一切変えない。
//
// これは控えそのものではなく、「いつ・何という名前で控えを作ったか」の一覧。
// 実体はドライブにあり、ファイルIDとURLで辿れる。
//
// **ドライブ上のファイルが消えても、この記録は残る。**
// 「あの日の控えはもう無い」と分かること自体に意味がある。
//
// 設計書には130件とあったが、実物は139行（2026-08-17 の下見）。

var 控え記録書出_列 = {
  created_at: ['作成日時'],
  kind: ['種別'],
  file_name: ['ファイル名'],
  file_id: ['ファイルID'],
  url: ['URL']
};

function 控えの記録を書き出す() {
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.BACKUP_LOG);
  if (!sheet) {
    Logger.log('■ 「' + SHEETS.BACKUP_LOG + '」シートがありません。');
    return;
  }

  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) {
    Logger.log('■ 空です。書き出すものがありません。');
    return;
  }

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  var 位 = 掲載書出_列を引く_(v[0], 控え記録書出_列);
  var 見つからない列 = Object.keys(位).filter(function (k) { return 位[k] < 0; });

  var 行 = [];
  var 空行 = 0;
  for (var r = 1; r < v.length; r += 1) {
    var 日 = 位.created_at >= 0 ? 掲載書出_日時_(v[r][位.created_at]) : null;
    var 名 = 位.file_name >= 0 ? 掲載書出_文字_(v[r][位.file_name]) : '';
    // 日時もファイル名も無い行は、書きかけの空行。
    if (!日 && !名) { 空行 += 1; continue; }

    行.push({
      row: r + 1,
      created_at: 日,
      kind: 位.kind >= 0 ? 掲載書出_文字_(v[r][位.kind]) : '',
      file_name: 名,
      file_id: 位.file_id >= 0 ? 掲載書出_文字_(v[r][位.file_id]) : '',
      url: 位.url >= 0 ? 掲載書出_文字_(v[r][位.url]) : ''
    });
  }

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    シートの行数: v.length - 1,
    backup_records: 行
  }, null, 2);

  var 名前 = '控えの記録_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024 * 10) / 10 + 'KB');
  Logger.log('');
  Logger.log('■ 控えの記録: ' + 行.length + '件（シートは' + (v.length - 1) + '行）');
  if (空行) Logger.log('    空行: ' + 空行 + '件');

  var 種 = {};
  行.forEach(function (x) { 種[x.kind || '（空）'] = (種[x.kind || '（空）'] || 0) + 1; });
  Logger.log('    種別: ' + Object.keys(種).sort(function (a, b) { return 種[b] - 種[a]; })
    .map(function (k) { return k + ' ' + 種[k] + '件'; }).join(' / '));

  var ID無し = 行.filter(function (x) { return !x.file_id; }).length;
  var URL無し = 行.filter(function (x) { return !x.url; }).length;
  Logger.log('    ファイルIDが無い: ' + ID無し + '件 / URLが無い: ' + URL無し + '件');

  var 日ら = 行.map(function (x) { return x.created_at; }).filter(Boolean).sort();
  if (日ら.length) {
    Logger.log('    期間: ' + 日ら[0].slice(0, 10) + ' 〜 ' + 日ら[日ら.length - 1].slice(0, 10));
  }
  var 日時無し = 行.filter(function (x) { return !x.created_at; }).length;
  if (日時無し) Logger.log('    **日時が無い: ' + 日時無し + '件**');

  // 同じファイルIDが2行あると、同じ控えを二重に記録している。
  var ID = {};
  行.forEach(function (x) { if (x.file_id) ID[x.file_id] = (ID[x.file_id] || 0) + 1; });
  var 重なり = Object.keys(ID).filter(function (k) { return ID[k] > 1; });
  if (重なり.length) {
    Logger.log('    **同じファイルIDが ' + 重なり.length + '種あります（二重に記録された控え）**');
  }
  Logger.log('');

  if (見つからない列.length) {
    Logger.log('  **見つからない列: ' + 見つからない列.join('・') + '**');
    Logger.log('');
  }
  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ 個人情報は含みません（控えのファイル名とURLだけ）。');
  Logger.log('     取り込みが済んだら 書き出しJSONを片付ける() で消してください。');
}
