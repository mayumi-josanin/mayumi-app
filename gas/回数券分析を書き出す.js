// 回数券分析結果（ビジリスのスプレッドシート）を、
// データベースへ取り込める形（JSON）で書き出す。
//
//   回数券分析を書き出す()   … JSONにしてドライブへ置く
//
// **読むだけ。**元のスプレッドシートは一切変えない。
//
// なぜ「まゆみ側」に置くのか:
//   この表はビジリスのスプレッドシートにあるので、本来はビジリスのGASに
//   置くのが素直。しかし bijiris/gas は `.claspignore` が許可制で、
//   **Code.gs も push 対象**になっている。そして bijiris/WORK_NOTES.md に
//   「未反映の変更がある」と書かれた状態が残っている。
//   push すると、確認していない変更が本番の動きを変えてしまう恐れがある。
//   ここからは **openById で読むだけ**なので、その心配がない。
//
// **お名前で会員に結びつけることはしない。**
//   この表には会員番号が無く、あるのはお名前だけ。CLAUDE.md にあるとおり、
//   お名前で探すと同姓同名・改名・表記ゆれで別の方に行き着く
//   （実際に記録が混ざったことがある）。
//   ここでは**お名前をそのまま持ち、結びつけは行わない。**
//   代わりに「何件が会員のお名前と一致するか」だけを数えて出す。
//   結びつけ方は、それを見てから決める。

var 回数券分析書出_ビジリスID = '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E';
var 回数券分析書出_シート名 = '回数券分析結果';

var 回数券分析書出_列 = {
  created_at: ['作成日時'],
  updated_at: ['更新日時'],
  response_id: ['回答ID'],
  customer_name: ['お名前'],
  submitted_at: ['提出日時'],
  before_photos: ['ビフォー写真JSON'],
  after_photos: ['アフター写真JSON'],
  status: ['分析状態'],
  result: ['分析結果'],
  analyzed_at: ['分析日時'],
  error: ['エラー']
};

function 回数券分析を書き出す() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(回数券分析書出_ビジリスID);
  } catch (e) {
    Logger.log('■ ビジリスのスプレッドシートを開けませんでした。');
    Logger.log('  ' + e);
    Logger.log('  このGASを動かしているアカウントに、閲覧の権限が要ります。');
    return;
  }

  var sheet = ss.getSheetByName(回数券分析書出_シート名);
  if (!sheet) {
    Logger.log('■ 「' + 回数券分析書出_シート名 + '」シートがありません。');
    return;
  }

  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) {
    Logger.log('■ 空です。書き出すものがありません。');
    return;
  }

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  var 位 = 掲載書出_列を引く_(v[0], 回数券分析書出_列);
  var 見つからない列 = Object.keys(位).filter(function (k) { return 位[k] < 0; });

  var 行 = [];
  var 空行 = 0;
  for (var r = 1; r < v.length; r += 1) {
    var 回答 = 位.response_id >= 0 ? 掲載書出_文字_(v[r][位.response_id]) : '';
    var 名 = 位.customer_name >= 0 ? 掲載書出_文字_(v[r][位.customer_name]) : '';
    if (!回答 && !名) { 空行 += 1; continue; }

    var o = { row: r + 1 };
    Object.keys(位).forEach(function (鍵) {
      var i = 位[鍵];
      if (i < 0) { o[鍵] = null; return; }
      var x = v[r][i];
      if (鍵.indexOf('_at') >= 0) o[鍵] = 掲載書出_日時_(x);
      else o[鍵] = 掲載書出_文字_(x);
    });
    行.push(o);
  }

  // 会員のお名前と一致するかだけを数える。**結びつけはしない。**
  var 会員名 = {};
  try {
    var まゆみ = getOrCreateSpreadsheet();
    var us = まゆみ.getSheetByName(SHEETS.USERS);
    if (us && us.getLastRow() > 1) {
      var uv = us.getRange(1, 1, us.getLastRow(), us.getLastColumn()).getValues();
      var 名列 = -1;
      for (var j = 0; j < uv[0].length; j += 1) {
        if (String(uv[0][j]).trim() === '氏名') { 名列 = j; break; }
      }
      if (名列 >= 0) {
        for (var k = 1; k < uv.length; k += 1) {
          var n = String(uv[k][名列] || '').trim();
          if (n) 会員名[n] = (会員名[n] || 0) + 1;
        }
      }
    }
  } catch (e2) {
    Logger.log('  （会員のお名前を読めませんでした: ' + e2 + '）');
  }

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    シートの行数: v.length - 1,
    ticket_analyses: 行
  }, null, 2);

  var 名前 = '回数券分析_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024 * 10) / 10 + 'KB');
  Logger.log('');
  Logger.log('■ 回数券分析: ' + 行.length + '件（シートは' + (v.length - 1) + '行）');
  if (空行) Logger.log('    空行: ' + 空行 + '件');

  var 状態 = {};
  行.forEach(function (x) { var s = x.status || '（空）'; 状態[s] = (状態[s] || 0) + 1; });
  Logger.log('    分析状態: ' + Object.keys(状態).sort(function (a, b) { return 状態[b] - 状態[a]; })
    .map(function (k) { return k + ' ' + 状態[k] + '件'; }).join(' / '));

  Logger.log('    分析結果がある: ' + 行.filter(function (x) { return x.result; }).length + '件');
  Logger.log('    エラーがある: ' + 行.filter(function (x) { return x.error; }).length + '件');
  Logger.log('    ビフォー写真がある: ' + 行.filter(function (x) { return x.before_photos; }).length + '件');
  Logger.log('    アフター写真がある: ' + 行.filter(function (x) { return x.after_photos; }).length + '件');
  Logger.log('    回答IDが空: ' + 行.filter(function (x) { return !x.response_id; }).length + '件');

  // お名前の照合は「数えるだけ」。結びつけない。
  var 一致 = 0, 複数 = [], 無し = [];
  行.forEach(function (x) {
    var n = x.customer_name;
    if (!n) return;
    var c = 会員名[n] || 0;
    if (c === 1) 一致 += 1;
    else if (c > 1) 複数.push(n + '（会員に' + c + '名）');
    else 無し.push(n);
  });
  Logger.log('');
  Logger.log('  ● 会員のお名前との一致（**数えるだけ。結びつけていません**）:');
  Logger.log('      ちょうど1名と一致: ' + 一致 + '件');
  if (複数.length) Logger.log('      **同じお名前が複数: ' + 複数.join('・') + '**');
  if (無し.length) Logger.log('      一致する会員が無い: ' + 無し.length + '件（' + 無し.slice(0, 5).join('・') + (無し.length > 5 ? '…' : '') + '）');
  Logger.log('      → お名前で結びつけると別の方に行き着くことがあります。');
  Logger.log('        結びつけ方は、この数を見てから決めてください。');
  Logger.log('');

  if (見つからない列.length) {
    Logger.log('  **見つからない列: ' + 見つからない列.join('・') + '**');
    Logger.log('');
  }
  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ **お客様のお名前・お体の分析結果・写真の場所が入っています。**');
  Logger.log('     取り込みが済んだら必ず 書き出しJSONを片付ける() で消してください。');
}
