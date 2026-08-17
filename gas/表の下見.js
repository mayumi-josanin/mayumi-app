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

// 次に移す表。移す前に必ずここに足して、実物を見てから作る。
// 設計書の件数が実物と違っていたことがある（お知らせ94→95、
// カレンダー116→143）。**推測で表を作らない。**
var 表下見_対象 = ['APP_SUPPORT_FAQ', 'PUSH_NOTICES'];

// 列の埋まり方に偏りがあったとき、その正体を数えるための道具。
// 「ここが空なのは古い行だからだろう」と推測すると外す。実際に数える。
//
//   { シート: 名前, 軸: '見出し', 見る: ['見出し', ...] }
//   … 軸の値ごとに、見る列が埋まっている件数を出す
//   値を見る: ['見出し', ...]
//   … 埋まっている件数が少ない列の、実際の値を並べる
//     （書き出すと個人情報がドライブに出る列かどうかを、見てから決める）
var 表下見_内訳 = [
  {
    シート: 'PUSH_NOTICES',
    軸: 'ステータス',
    見る: ['日時', '通知ID', '送信結果', '送信件数'],
    値を見る: ['送信対象詳細']
  }
];

function 表の内訳() {
  var ss = getOrCreateSpreadsheet();

  表下見_内訳.forEach(function (指定) {
    Logger.log('==============================');
    Logger.log('■ ' + 指定.シート + ' … ' + 指定.軸 + 'ごとの内訳');
    var sh = ss.getSheetByName(指定.シート);
    if (!sh) { Logger.log('  **シートがありません**'); return; }

    var 最終行 = sh.getLastRow();
    if (最終行 < 2) { Logger.log('  空です'); return; }

    var v = sh.getRange(1, 1, 最終行, sh.getLastColumn()).getValues();
    var 見出し = v[0].map(function (h) { return String(h).trim(); });
    var 軸i = 見出し.indexOf(指定.軸);
    if (軸i < 0) { Logger.log('  **' + 指定.軸 + ' 列がありません**'); return; }

    var 見るi = 指定.見る.map(function (名) { return { 名: 名, i: 見出し.indexOf(名) }; });

    var 集計 = {};
    var 並び = [];
    for (var r = 1; r < v.length; r += 1) {
      var 値 = String(v[r][軸i] || '（空）').trim() || '（空）';
      if (!集計[値]) { 集計[値] = { 件数: 0, 埋: {} }; 並び.push(値); }
      集計[値].件数 += 1;
      見るi.forEach(function (c) {
        if (c.i < 0) return;
        var x = v[r][c.i];
        if (x !== '' && x !== null && x !== undefined) {
          集計[値].埋[c.名] = (集計[値].埋[c.名] || 0) + 1;
        }
      });
    }

    並び.forEach(function (値) {
      var s = 集計[値];
      var 内 = 見るi.map(function (c) {
        return c.名 + ' ' + (s.埋[c.名] || 0) + '/' + s.件数;
      }).join('　');
      Logger.log('  ' + 値 + ' … ' + s.件数 + '件　（' + 内 + '）');
    });
    Logger.log('');

    (指定.値を見る || []).forEach(function (名) {
      var i = 見出し.indexOf(名);
      if (i < 0) { Logger.log('  **' + 名 + ' 列がありません**'); return; }
      Logger.log('  ' + 名 + ' の中身:');
      var 数えた = 0;
      for (var r = 1; r < v.length; r += 1) {
        var x = v[r][i];
        if (x === '' || x === null || x === undefined) continue;
        数えた += 1;
        var t = String(x);
        // 長い値は頭だけ。会員の一覧が入っていれば、頭を見れば分かる。
        if (t.length > 120) t = t.slice(0, 120) + '…（全' + String(x).length + '文字）';
        Logger.log('    ' + (r + 1) + '行目: ' + t);
        if (数えた >= 20) { Logger.log('    …（ここまで20件）'); break; }
      }
      if (!数えた) Logger.log('    （すべて空）');
      Logger.log('');
    });
  });

  Logger.log('==============================');
  Logger.log('  ※ 読むだけです。シートは変えていません。');
}

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
