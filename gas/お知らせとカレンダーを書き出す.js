// お知らせとカレンダーを、データベースへ取り込める形（JSON）で書き出す。
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
  image_url: ['画像URL'],
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
      for (var i = 0; i < 見出し.length; i += 1) {
        if (String(見出し[i]).trim() === 名) { 位[鍵] = i; return; }
      }
    });
  });
  return 位;
}

function 掲載書出_1枚_(sheet, 定義, 日付の鍵) {
  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) return { 行: [], 見つからない列: [], 全体: 0 };

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  var 位 = 掲載書出_列を引く_(v[0], 定義);
  var 見つからない列 = Object.keys(位).filter(function (k) { return 位[k] < 0; });

  var 行 = [];
  for (var r = 1; r < v.length; r += 1) {
    var row = v[r];
    var 題 = 位.title >= 0 ? 掲載書出_文字_(row[位.title]) : '';
    var 日 = 位[日付の鍵] >= 0 ? 掲載書出_日付_(row[位[日付の鍵]]) : null;
    // 題も日付も無い行は、書きかけの空行。移す意味がない。
    if (!題 && !日) continue;

    var o = { row: r + 1 };
    Object.keys(位).forEach(function (鍵) {
      var i = 位[鍵];
      if (i < 0) { o[鍵] = null; return; }
      var x = row[i];
      if (鍵 === 'posted_on' || 鍵 === 'event_on') o[鍵] = 掲載書出_日付_(x);
      else if (鍵.indexOf('_at') >= 0) o[鍵] = 掲載書出_日時_(x);
      else if (鍵 === 'sort_order' || 鍵 === 'menu_row') {
        var n = parseInt(掲載書出_文字_(x), 10);
        o[鍵] = isNaN(n) ? null : n;
      } else o[鍵] = 掲載書出_文字_(x);
    });
    行.push(o);
  }
  return { 行: 行, 見つからない列: 見つからない列, 全体: v.length - 1 };
}

function 掲載物を書き出す() {
  var ss = getOrCreateSpreadsheet();

  var お知らせ = 掲載書出_1枚_(ss.getSheetByName(SHEETS.BLOG), 掲載書出_お知らせの列, 'posted_on');
  var カレンダー = 掲載書出_1枚_(ss.getSheetByName(SHEETS.CALENDAR), 掲載書出_カレンダーの列, 'event_on');

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    news: お知らせ.行,
    calendar: カレンダー.行
  }, null, 2);

  var 名前 = '掲載物_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('');
  Logger.log('■ お知らせ: ' + お知らせ.行.length + '件（シートは' + お知らせ.全体 + '行）');
  Logger.log('    公開:       ' + お知らせ.行.filter(function (x) { return x.published === '公開'; }).length + '件');
  Logger.log('    削除済み:   ' + お知らせ.行.filter(function (x) { return x.deleted; }).length + '件');
  if (お知らせ.見つからない列.length) {
    Logger.log('    **見つからない列: ' + お知らせ.見つからない列.join('・') + '**');
  }
  Logger.log('');
  Logger.log('■ カレンダー: ' + カレンダー.行.length + '件（シートは' + カレンダー.全体 + '行）');
  Logger.log('    公開:       ' + カレンダー.行.filter(function (x) { return x.published === '公開'; }).length + '件');
  Logger.log('    削除済み:   ' + カレンダー.行.filter(function (x) { return x.deleted; }).length + '件');
  if (カレンダー.見つからない列.length) {
    Logger.log('    **見つからない列: ' + カレンダー.見つからない列.join('・') + '**');
  }
  Logger.log('');
  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ 個人情報は含みませんが、取り込みが済んだら消してください。');
}
