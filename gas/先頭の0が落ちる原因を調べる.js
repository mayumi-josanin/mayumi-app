// 電話番号とパスコードの先頭の0が落ちる原因を調べる道具。**読むだけ。**
//
//   先頭の0が落ちる原因を調べる()
//
// 列の書式が「文字列（@）」になっているかを、行ごとに見る。
// 書式が数値のままの行に入力すると、080 が 80 になる。

function 先頭の0が落ちる原因を調べる() {
  var ss = getOrCreateSpreadsheet();
  var sheet = getOrCreateUsersSheet_(ss);
  var maxRows = sheet.getMaxRows();
  var 最終 = sheet.getLastRow();

  Logger.log('■ 会員シートの書式');
  Logger.log('');
  Logger.log('  行数: 使っている ' + 最終 + ' 行 / 用意されている ' + maxRows + ' 行');
  Logger.log('');

  [['電話番号', USER_COL.PHONE], ['パスコード', USER_COL.PASSCODE]].forEach(function (組) {
    var 名 = 組[0], 列 = 組[1];
    var 書式 = sheet.getRange(2, 列, Math.max(maxRows - 1, 1), 1).getNumberFormats();
    var 文字 = 0, 数値 = 0, 最初の数値行 = 0;
    書式.forEach(function (r, i) {
      if (String(r[0]) === '@') 文字++;
      else { 数値++; if (!最初の数値行) 最初の数値行 = i + 2; }
    });
    Logger.log('  ' + 名 + '（' + 列 + '列目）');
    Logger.log('    文字列の行: ' + 文字 + ' / 数値のままの行: ' + 数値);
    if (数値) Logger.log('    **数値のまま残っている最初の行: ' + 最初の数値行 + '行目**');
    Logger.log('    いちばん下の行の書式: ' + sheet.getRange(maxRows, 列).getNumberFormat());
    Logger.log('');
  });

  // 実際に落ちている行が、どのあたりにあるか（中身は出さない）
  var 件 = Math.max(0, 最終 - 1);
  if (件) {
    var 値 = sheet.getRange(2, 1, 件, USER_HEADERS.length).getDisplayValues();
    var 電話の行 = [], パスの行 = [];
    値.forEach(function (row, i) {
      var t = String(row[USER_COL.PHONE - 1] || '').replace(/\D/g, '');
      if (t && t.charAt(0) !== '0') 電話の行.push(i + 2);
      var p = String(row[USER_COL.PASSCODE - 1] || '').trim();
      if (p && /^\d+$/.test(p) && p.length !== 4 && p.length !== 6) パスの行.push(i + 2);
    });
    Logger.log('  落ちている行（行番号だけ・中身は出しません）');
    Logger.log('    電話番号  : ' + (電話の行.join(', ') || 'なし'));
    Logger.log('    パスコード: ' + (パスの行.join(', ') || 'なし'));
    Logger.log('');
    Logger.log('    登録日時（その行が作られた時期）');
    電話の行.concat(パスの行).forEach(function (r) {
      Logger.log('      ' + r + '行目: ' + String(値[r - 2][USER_COL.TIMESTAMP - 1] || '(空)'));
    });
  }
  Logger.log('');
  Logger.log('  ※ 読むだけです。');
}
