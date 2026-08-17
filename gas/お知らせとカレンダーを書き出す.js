// お知らせ・カレンダー・メニュー・商品・カテゴリを、
// データベースへ取り込める形（JSON）で書き出す。
//
//   掲載物を書き出す()   … JSONにしてドライブへ置く
//
// **読むだけ。**元のスプレッドシートは一切変えない。
//
// 列は**見出しの名前で拾う。**位置で拾うと、列が足されたときに黙ってずれる。
// この2つのシートは、公開設定・画像URL・削除状態などが後から
// ensure〜 で必要に応じて足される作りなので、位置は当てにできない。
//
// 削除済みの行も、印を付けたまま書き出す。消してしまうと
// 「あの日は何を出していたか」を辿れなくなる。

var 掲載書出_お知らせの列 = {
  posted_on: ['投稿日（例：2025-06-15）', '投稿日'],
  title: ['タイトル'],
  category: ['カテゴリ'],
  icon: ['アイコン絵文字'],
  body: ['本文'],
  published: ['公開設定'],
  updated_at: ['更新日時'],
  sort_order: ['表示順'],
  notice_listed: ['お知らせ一覧公開'],
  deleted: ['削除状態'],
  deleted_at: ['削除日時'],
  delete_reason: ['削除理由'],
  publish_at: ['公開開始日時'],
  notice_listed_at: ['お知らせ一覧掲載日時'],
  notice_delisted_at: ['お知らせ一覧削除日時'],
  image_url: ['画像URL'],
  link_url: ['リンクURL'],
  link_label: ['リンクボタン名'],
  button_text: ['ボタンテキスト']
};

var 掲載書出_カレンダーの列 = {
  event_on: ['日付（例：2025-06-15）', '日付'],
  title: ['イベント名'],
  detail: ['詳細'],
  color: ['カラー（色名またはコード）', 'カラー'],
  published: ['公開設定'],
  // カレンダーの画像URLは、**見出しが空のまま**6列目に入っている。
  // シートを作ったときの7列（日付・イベント名・詳細・カラー・公開設定・
  // 画像URL・更新日時）の6番目がそれ。見出しだけが入っていない。
  // 名前で探すと見つからず、**画像30件が黙って落ちる。**
  // ここだけは位置（6列目）も手がかりにする。
  image_url: ['画像URL', '__6列目'],
  updated_at: ['更新日時'],
  sort_order: ['表示順'],
  notice_listed: ['お知らせ一覧公開'],
  deleted: ['削除状態'],
  deleted_at: ['削除日時'],
  delete_reason: ['削除理由'],
  publish_at: ['公開開始日時'],
  notice_listed_at: ['お知らせ一覧掲載日時'],
  notice_delisted_at: ['お知らせ一覧削除日時'],
  link_url: ['リンクURL'],
  button_text: ['ボタンテキスト'],
  menu_row: ['対象メニュー行']
};

// ---- メニュー・商品・カテゴリ ----

var 掲載書出_メニューの列 = {
  registered_on: ['登録日'],
  name: ['メニュー名'],
  image_urls: ['画像URL'],
  summary: ['概要説明'],
  booking_status: ['予約状況'],
  published: ['公開設定'],
  updated_at: ['更新日時'],
  category: ['カテゴリ'],
  sort_key: ['表示順'],
  notice_listed: ['お知らせ一覧公開'],
  publish_at: ['公開開始日時'],
  deleted: ['削除状態'],
  deleted_at: ['削除日時'],
  notice_delisted_at: ['お知らせ一覧削除日時'],
  notice_listed_at: ['お知らせ一覧掲載日時'],
  delete_reason: ['削除理由']
};

var 掲載書出_商品の列 = {
  category: ['カテゴリ'],
  name: ['商品名'],
  price: ['価格（円）', '価格'],
  icon_url: ['アイコン'],
  background_color: ['背景色コード'],
  published: ['公開設定'],
  description: ['商品説明'],
  description_image_url: ['商品説明画像'],
  updated_at: ['更新日時'],
  stock: ['在庫数'],
  stock_warning: ['在庫警告閾値'],
  deleted_at: ['削除日時'],
  publish_at: ['公開開始日時'],
  notice_listed: ['お知らせ一覧公開'],
  deleted: ['削除状態'],
  special_price: ['特別価格（円）', '特別価格']
};

var 掲載書出_カテゴリの列 = {
  name: ['カテゴリ名'],
  // 2列目に見出しが入っていない。値は「お知らせ」「ブログ」で、
  // カテゴリの種別。名前で拾えないので位置で拾う。
  kind: ['種別', '__2列目']
};

function 掲載書出_空か_(x) { return x === '' || x === null || x === undefined; }

function 掲載書出_文字_(値) {
  return 掲載書出_空か_(値) ? '' : String(値).trim();
}

function 掲載書出_日付_(値) {
  if (掲載書出_空か_(値)) return null;
  var d = 値 instanceof Date ? 値 : new Date(値);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function 掲載書出_日時_(値) {
  if (掲載書出_空か_(値)) return null;
  var d = 値 instanceof Date ? 値 : new Date(値);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

// 見出しの名前から列の位置を作る。見つからない列は -1 にして、
// 後で「この列は無かった」と分かるようにする。
function 掲載書出_列を引く_(見出し, 定義) {
  var 位 = {};
  Object.keys(定義).forEach(function (鍵) {
    位[鍵] = -1;
    定義[鍵].forEach(function (名) {
      if (位[鍵] >= 0) return;

      // '__N列目' と書かれていたら、見出しではなく位置で拾う。
      // 見出しが入っていない列があるため、最後の手段として用意している。
      // **見出しが付いたら、そちらが先に見つかるので自然に切り替わる。**
      var m = /^__(\d+)列目$/.exec(名);
      if (m) {
        var i = parseInt(m[1], 10) - 1;
        if (i >= 0 && i < 見出し.length) 位[鍵] = i;
        return;
      }

      for (var j = 0; j < 見出し.length; j += 1) {
        if (String(見出し[j]).trim() === 名) { 位[鍵] = j; return; }
      }
    });
  });
  return 位;
}

// 名前の鍵は表によって違う（お知らせ・カレンダーは title、
// メニュー・商品・カテゴリは name）。**ここを間違えると、
// 空行の判定にかからず全行が飛ばされる。**
function 掲載書出_1枚_(sheet, 定義, 日付の鍵, 名前の鍵) {
  名前の鍵 = 名前の鍵 || 'title';
  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) return { 行: [], 見つからない列: [], 全体: 0 };

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  var 位 = 掲載書出_列を引く_(v[0], 定義);
  var 見つからない列 = Object.keys(位).filter(function (k) { return 位[k] < 0; });

  var 行 = [];
  for (var r = 1; r < v.length; r += 1) {
    var row = v[r];
    var 題 = 位[名前の鍵] >= 0 ? 掲載書出_文字_(row[位[名前の鍵]]) : '';
    var 日 = (日付の鍵 && 位[日付の鍵] >= 0) ? 掲載書出_日付_(row[位[日付の鍵]]) : null;
    // 題も日付も無い行は、書きかけの空行。移す意味がない。
    if (!題 && !日) continue;

    var o = { row: r + 1 };
    Object.keys(位).forEach(function (鍵) {
      var i = 位[鍵];
      if (i < 0) { o[鍵] = null; return; }
      var x = row[i];
      if (鍵 === 'posted_on' || 鍵 === 'event_on') o[鍵] = 掲載書出_日付_(x);
      else if (鍵.indexOf('_at') >= 0) o[鍵] = 掲載書出_日時_(x);
      else if (鍵 === 'sort_order' || 鍵 === 'menu_row' || 鍵 === 'sort_key' ||
               鍵 === 'price' || 鍵 === 'special_price' ||
               鍵 === 'stock' || 鍵 === 'stock_warning') {
        // 空欄は null のまま。**0 と空欄は意味が違う。**
        // 在庫0（売り切れ）と、在庫を管理していない、を混ぜない。
        var t = 掲載書出_文字_(x);
        if (!t) { o[鍵] = null; }
        else {
          var n = Number(t);
          o[鍵] = isFinite(n) ? n : null;
        }
      }
      else if (鍵 === 'image_urls') {
        // メニューの画像は配列で入っている（1つのメニューに複数枚）。
        // 文字列のまま渡すと、あとで1枚ずつ扱えない。読める形にして渡す。
        var v = 掲載書出_文字_(x);
        if (!v) { o[鍵] = []; }
        else if (v.charAt(0) === '[') {
          try { o[鍵] = JSON.parse(v); } catch (e) { o[鍵] = [v]; }
        } else { o[鍵] = [v]; }
      }
      else o[鍵] = 掲載書出_文字_(x);
    });
    行.push(o);
  }
  return { 行: 行, 見つからない列: 見つからない列, 全体: v.length - 1 };
}

function 掲載物を書き出す() {
  var ss = getOrCreateSpreadsheet();

  var お知らせ = 掲載書出_1枚_(ss.getSheetByName(SHEETS.BLOG), 掲載書出_お知らせの列, 'posted_on');
  var カレンダー = 掲載書出_1枚_(ss.getSheetByName(SHEETS.CALENDAR), 掲載書出_カレンダーの列, 'event_on');

  var メニュー = 掲載書出_1枚_(ss.getSheetByName(SHEETS.MENUS), 掲載書出_メニューの列, 'registered_on', 'name');
  var 商品 = 掲載書出_1枚_(ss.getSheetByName(SHEETS.PRODUCTS), 掲載書出_商品の列, '', 'name');
  var カテゴリ = 掲載書出_1枚_(ss.getSheetByName(SHEETS.CATEGORIES), 掲載書出_カテゴリの列, '', 'name');

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    news: お知らせ.行,
    calendar: カレンダー.行,
    menus: メニュー.行,
    products: 商品.行,
    categories: カテゴリ.行
  }, null, 2);

  var 名前 = '掲載物_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('');
  [
    { 名: 'お知らせ', r: お知らせ },
    { 名: 'カレンダー', r: カレンダー },
    { 名: 'メニュー', r: メニュー },
    { 名: '商品', r: 商品 },
    { 名: 'カテゴリ', r: カテゴリ }
  ].forEach(function (x) {
    Logger.log('■ ' + x.名 + ': ' + x.r.行.length + '件（シートは' + x.r.全体 + '行）');
    var 公開 = x.r.行.filter(function (y) { return y.published === '公開'; }).length;
    var 削除 = x.r.行.filter(function (y) { return y.deleted; }).length;
    if (x.名 !== 'カテゴリ') {
      Logger.log('    公開: ' + 公開 + '件 / 削除済み: ' + 削除 + '件');
    }
    if (x.名 === 'メニュー') {
      var 画像 = x.r.行.filter(function (y) { return y.image_urls && y.image_urls.length; }).length;
      var 枚数 = x.r.行.reduce(function (a, y) { return a + ((y.image_urls || []).length); }, 0);
      Logger.log('    画像がある: ' + 画像 + '件（のべ' + 枚数 + '枚）');
    }
    if (x.r.見つからない列.length) {
      Logger.log('    **見つからない列: ' + x.r.見つからない列.join('・') + '**');
    }
    Logger.log('');
  });

  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ 個人情報は含みませんが、取り込みが済んだら消してください。');
}
