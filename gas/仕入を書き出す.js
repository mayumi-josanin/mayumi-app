// 仕入値（管理マスタ）を、データベースへ取り込める形（JSON）で書き出す。
//
//   仕入を書き出す()   … JSONにしてドライブへ置く
//
// **読むだけ。**元のスプレッドシートは一切変えない。
//
// このシートは5列あるが、**使うのは3列だけ**（2026-08-17 の下見）。
//
//   1. 商品名（完全一致）  22件
//   2. 仕入値（円）        22件
//   3. 備考               14件
//   4. （見出しなし）        0件  ← 空の列
//   5. 【仕入値の自動入力について】 3件  ← 人向けの説明文
//
// 5列目には「注文管理シートでは、D列の『商品名』とここの『商品名』を
// 照合して…」という運用メモが3行だけ入っている。
// **列だと思って移すと、仕入値の表に意味のない文字列が混ざる。**
// 見出しの名前で拾うので、この2列は自然と落ちる。
// ただし「落ちたことが分かる」ように、下の 仕入書出_使わない列 に
// 名前を書いて、書き出しのたびに件数を出す。**黙って捨てない。**

var 仕入書出_列 = {
  product_name: ['商品名（完全一致）', '商品名'],
  price: ['仕入値（円）', '仕入値'],
  memo: ['備考']
};

// 移さないと決めた列。書き出しのたびに、何を置いていったかを出す。
var 仕入書出_使わない列 = ['【仕入値の自動入力について】'];

function 仕入を書き出す() {
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.MASTER);
  if (!sheet) {
    Logger.log('■ 「' + SHEETS.MASTER + '」シートがありません。');
    return;
  }

  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) {
    Logger.log('■ 空です。書き出すものがありません。');
    return;
  }

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  var 位 = 掲載書出_列を引く_(v[0], 仕入書出_列);
  var 見つからない列 = Object.keys(位).filter(function (k) { return 位[k] < 0; });

  var 行 = [];
  for (var r = 1; r < v.length; r += 1) {
    var 名 = 位.product_name >= 0 ? 掲載書出_文字_(v[r][位.product_name]) : '';
    // 商品名が無ければ、どの商品の仕入値か分からない。移す意味がない。
    if (!名) continue;

    var 値 = 掲載書出_文字_(位.price >= 0 ? v[r][位.price] : '');
    var n = 値 === '' ? null : Number(値);
    行.push({
      row: r + 1,
      product_name: 名,
      // **空欄は null。0（もらいもの・原価なし）と混ぜない。**
      price: (値 !== '' && isFinite(n)) ? n : null,
      memo: 位.memo >= 0 ? 掲載書出_文字_(v[r][位.memo]) : ''
    });
  }

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    supplier_prices: 行
  }, null, 2);

  var 名前 = '仕入_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024 * 10) / 10 + 'KB');
  Logger.log('');
  Logger.log('■ 仕入値: ' + 行.length + '件（シートは' + (v.length - 1) + '行）');

  var 値なし = 行.filter(function (x) { return x.price === null; }).length;
  var ゼロ = 行.filter(function (x) { return x.price === 0; }).length;
  var 備考あり = 行.filter(function (x) { return x.memo; }).length;
  Logger.log('    仕入値: 空欄 ' + 値なし + '件 / 0 ' + ゼロ + '件');
  Logger.log('    備考がある: ' + 備考あり + '件');

  var 名前の重なり = 行.length - Object.keys(行.reduce(function (a, x) {
    a[x.product_name] = 1; return a;
  }, {})).length;
  if (名前の重なり) {
    // 商品名で商品マスタと結びつけるので、同じ名前が2行あると
    // どちらの仕入値か決まらない。**気づかず通さない。**
    Logger.log('    **同じ商品名が ' + 名前の重なり + '件あります。**');
    Logger.log('      商品名で商品マスタと結びつけるため、どちらの値か決まりません。');
  }
  Logger.log('');

  // 移さないと決めた列に何が入っていたかを出す。あとから
  // 「あの列はどうしたのか」と思ったときに、ここを見れば分かる。
  Logger.log('■ 移さない列:');
  仕入書出_使わない列.forEach(function (名) {
    var i = -1;
    for (var j = 0; j < v[0].length; j += 1) {
      if (String(v[0][j]).trim() === 名) { i = j; break; }
    }
    if (i < 0) { Logger.log('  ' + 名 + ' … 見つかりません（すでに無い列）'); return; }
    var 中 = [];
    for (var r2 = 1; r2 < v.length; r2 += 1) {
      var x = 掲載書出_文字_(v[r2][i]);
      if (x) 中.push((r2 + 1) + '行目: ' + (x.length > 60 ? x.slice(0, 60) + '…' : x));
    }
    Logger.log('  ' + 名 + ' … ' + 中.length + '件（人向けの説明文なので移しません）');
    中.forEach(function (t) { Logger.log('      ' + t); });
  });
  Logger.log('');

  if (見つからない列.length) {
    Logger.log('  **見つからない列: ' + 見つからない列.join('・') + '**');
    Logger.log('');
  }
  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ 個人情報は含みません。');
  Logger.log('     取り込みが済んだら 書き出しJSONを片付ける() で消してください。');
}
