// 売上（MENU_REVENUE ＋ PRODUCT_REVENUE）を、
// データベースへ取り込める形（JSON）で書き出す。
//
//   売上を書き出す()   … JSONにしてドライブへ置く
//
// **読むだけ。**元のスプレッドシートは一切変えない。
//
// 掲載物とは別のファイルにしている。掲載物のJSONは486KBあり、
// 毎回それを落とすのは手間なので、用のあるぶんだけ小さく出す。
//
// 2枚のシートは同じ意味の列を違う名前で持っている（2026-08-17 の下見）。
//
//   ここでの名前   MENU_REVENUE   PRODUCT_REVENUE
//   name          メニュー種別     商品名
//   quantity      件数            個数
//   unit_price    単価            単価
//   unit_cost     原価単価         原価
//
// **「原価単価」と「原価」は、どちらも1つあたりの値段。**
// メニュー側の見出しだけが「単価」まで書いてある。合計ではない。

var 売上書出_メニューの列 = {
  recorded_on: ['記録日'],
  name: ['メニュー種別'],
  quantity: ['件数'],
  unit_price: ['単価'],
  unit_cost: ['原価単価'],
  memo: ['メモ'],
  deleted: ['削除状態'],
  deleted_at: ['削除日時']
};

var 売上書出_商品の列 = {
  // 1列目の見出しが読み取れなかったときのために位置も持たせる。
  // カレンダーの画像URLで、見出しが空のまま30件落ちかけた前例がある。
  recorded_on: ['記録日', '__1列目'],
  name: ['商品名'],
  quantity: ['個数'],
  unit_price: ['単価'],
  unit_cost: ['原価'],
  memo: ['メモ'],
  deleted: ['削除状態'],
  deleted_at: ['削除日時']
};

function 売上書出_1枚_(sheet, 定義, 種別) {
  if (!sheet) return { 行: [], 見つからない列: [], 全体: 0 };
  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) return { 行: [], 見つからない列: [], 全体: 0 };

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  // 列を引く仕組みは掲載物の書き出しと同じものを使う。
  // 見出しの名前で拾い、見つからなければ '__N列目' で位置から拾う。
  var 位 = 掲載書出_列を引く_(v[0], 定義);
  var 見つからない列 = Object.keys(位).filter(function (k) { return 位[k] < 0; });

  var 行 = [];
  for (var r = 1; r < v.length; r += 1) {
    var row = v[r];
    var 名 = 位.name >= 0 ? 掲載書出_文字_(row[位.name]) : '';
    var 日 = 位.recorded_on >= 0 ? 掲載書出_日付_(row[位.recorded_on]) : null;
    // 名前も日付も無い行は、書きかけの空行。
    if (!名 && !日) continue;

    var o = { kind: 種別, row: r + 1 };
    Object.keys(位).forEach(function (鍵) {
      var i = 位[鍵];
      if (i < 0) { o[鍵] = null; return; }
      var x = row[i];
      if (鍵 === 'recorded_on') o[鍵] = 掲載書出_日付_(x);
      else if (鍵.indexOf('_at') >= 0) o[鍵] = 掲載書出_日時_(x);
      else if (鍵 === 'quantity' || 鍵 === 'unit_price' || 鍵 === 'unit_cost') {
        // **空欄は null のまま。0 と空欄は意味が違う。**
        // 原価0（仕入れの無い教室など）と、原価を記録していない、を混ぜない。
        // 混ぜると粗利が変わる。
        var t = 掲載書出_文字_(x);
        if (!t) { o[鍵] = null; }
        else {
          var n = Number(t);
          o[鍵] = isFinite(n) ? n : null;
        }
      }
      else o[鍵] = 掲載書出_文字_(x);
    });
    行.push(o);
  }
  return { 行: 行, 見つからない列: 見つからない列, 全体: v.length - 1 };
}

function 売上書出_まとめ_(名, r) {
  Logger.log('■ ' + 名 + ': ' + r.行.length + '件（シートは' + r.全体 + '行）');
  var 削除 = r.行.filter(function (y) { return y.deleted; }).length;

  // 0 と空欄がきちんと分かれているかを、書き出した時点で数えておく。
  // 取り込んだあとに「なぜか粗利が合わない」となってから探すより早い。
  ['quantity', 'unit_price', 'unit_cost'].forEach(function (鍵) {
    var 空 = r.行.filter(function (y) { return y[鍵] === null; }).length;
    var ゼロ = r.行.filter(function (y) { return y[鍵] === 0; }).length;
    Logger.log('    ' + 鍵 + ': 空欄 ' + 空 + '件 / 0 ' + ゼロ + '件');
  });

  var 金額 = r.行.reduce(function (a, y) {
    return (y.quantity === null || y.unit_price === null)
      ? a : a + y.quantity * y.unit_price;
  }, 0);
  var 出せない = r.行.filter(function (y) {
    return y.quantity === null || y.unit_price === null;
  }).length;
  Logger.log('    売上の合計: ' + 金額.toLocaleString() + '円'
    + (出せない ? '（数か単価が空で出せない行が ' + 出せない + '件）' : ''));
  Logger.log('    削除済み: ' + 削除 + '件');
  if (r.見つからない列.length) {
    Logger.log('    **見つからない列: ' + r.見つからない列.join('・') + '**');
  }
  Logger.log('');
}

function 売上を書き出す() {
  var ss = getOrCreateSpreadsheet();

  var メニュー = 売上書出_1枚_(ss.getSheetByName(SHEETS.MENU_REVENUE), 売上書出_メニューの列, 'メニュー');
  var 商品 = 売上書出_1枚_(ss.getSheetByName(SHEETS.PRODUCT_REVENUE), 売上書出_商品の列, '商品');

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    revenues: メニュー.行.concat(商品.行)
  }, null, 2);

  var 名前 = '売上_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024) + 'KB');
  Logger.log('');

  売上書出_まとめ_('メニューの売上', メニュー);
  売上書出_まとめ_('商品の売上', 商品);
  Logger.log('■ 合わせて ' + (メニュー.行.length + 商品.行.length) + '件');
  Logger.log('');
  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ 個人情報は含みません（お名前も会員番号も入っていません）。');
  Logger.log('     取り込みが済んだら 書き出しJSONを片付ける() で消してください。');
}
