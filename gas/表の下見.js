// これから移す表の、実際の列と中身を見る道具。**読むだけ。**
//
//   表の下見()   … 対象のシートの見出し・件数・埋まり方を出す
//
// なぜ要るのか:
//   コードから列を推測して表を作ると、実際とずれる。
//   会員データのときは USER_HEADERS があったので確かだったが、
//   お知らせ・カレンダーは列が後から足されている（更新日時・公開・
//   画像URL・削除状態など、ensure〜 で必要に応じて増える作り）。
//   **実物の1行目を見てから表を作る。**

var 表下見_対象 = ['ブログ・お知らせ', 'カレンダー'];

function 表の下見() {
  var ss = getOrCreateSpreadsheet();

  表下見_対象.forEach(function (名) {
    var sh = ss.getSheetByName(名);
    Logger.log('==============================');
    Logger.log('■ ' + 名);
    if (!sh) {
      Logger.log('  **シートがありません**');
      Logger.log('');
      return;
    }

    var 最終行 = sh.getLastRow();
    var 最終列 = sh.getLastColumn();
    Logger.log('  ' + (最終行 - 1) + '行 × ' + 最終列 + '列');
    Logger.log('');

    if (最終行 < 1) { Logger.log('  空です'); Logger.log(''); return; }

    var v = sh.getRange(1, 1, Math.min(最終行, 200), 最終列).getValues();
    var 見出し = v[0];

    Logger.log('  列と、埋まっている件数:');
    見出し.forEach(function (h, i) {
      var 埋 = 0, 例 = '';
      for (var r = 1; r < v.length; r += 1) {
        var x = v[r][i];
        if (x !== '' && x !== null && x !== undefined) {
          埋 += 1;
          if (!例) {
            例 = x instanceof Date
              ? Utilities.formatDate(x, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')
              : String(x);
            if (例.length > 40) 例 = 例.slice(0, 40) + '…';
          }
        }
      }
      Logger.log('    ' + (i + 1) + '. ' + (String(h) || '（見出しなし）') +
        '  … ' + 埋 + '件' + (例 ? '  例: ' + 例 : ''));
    });
    Logger.log('');
  });

  Logger.log('==============================');
  Logger.log('  ※ 読むだけです。シートは変えていません。');
}
