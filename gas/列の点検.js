// 移した表の列が、そのあと増えていないかを点検する道具。
//
//   列を点検する()   … 各シートの見出しを全部出す
//
// **読むだけ。**
//
// なぜ要るのか:
//   このスプレッドシートは、機能を足すたびに **ensure〜 が列を自動で増やす。**
//   移した時点の列数は「その時点のもの」でしかない。
//
//   実際、商品マスタは移行時（2026-08-17）に16列だったが、
//   5日後には**22列**になっていた。増えた「売切状態」を見落としたまま
//   getProducts を作りかけ、**売切の商品が「在庫あり」として出る**ところだった。
//
//   **作る直前に、必ずこれで見出しを確かめる。**
//   サーバー側のモデルと突き合わせて、足りない列を見つける。

var 点検_対象 = [
  'ブログ・お知らせ', 'カレンダー', 'MENUS', '商品マスタ',
  'カテゴリマスタ', 'APP_SUPPORT_FAQ', 'PUSH_NOTICES'
];

function 列を点検する() {
  var ss = getOrCreateSpreadsheet();

  Logger.log('■ 各シートの見出し（サーバー側と突き合わせてください）');
  Logger.log('');

  点検_対象.forEach(function (名) {
    var sh = ss.getSheetByName(名);
    if (!sh) {
      Logger.log('■ ' + 名 + '  **シートがありません**');
      Logger.log('');
      return;
    }
    var 最終行 = sh.getLastRow();
    var 最終列 = sh.getLastColumn();
    var 見出し = sh.getRange(1, 1, 1, 最終列).getValues()[0];

    Logger.log('■ ' + 名 + '  ' + (最終行 - 1) + '行 × ' + 最終列 + '列');

    // 埋まっている件数も一緒に出す。空の列は移さなくても実害が無いため。
    var v = 最終行 > 1
      ? sh.getRange(2, 1, 最終行 - 1, 最終列).getValues()
      : [];
    var 行 = [];
    見出し.forEach(function (h, i) {
      var 埋 = 0;
      for (var r = 0; r < v.length; r += 1) {
        var x = v[r][i];
        if (x !== '' && x !== null && x !== undefined) 埋 += 1;
      }
      var 名前 = String(h).trim() || '（見出しなし）';
      行.push((i + 1) + '.' + 名前 + '(' + 埋 + ')');
    });
    // 1行にまとめて出す。1列1行だと画面が流れて読めない。
    Logger.log('  ' + 行.join('  '));
    Logger.log('');
  });

  Logger.log('  ※ 読むだけです。');
  Logger.log('  ※ ( ) の中は埋まっている件数。0件の列は、移さなくても実害はありません。');
  Logger.log('  ※ **サーバーのモデルに無い列があれば、それが見落としです。**');
}
