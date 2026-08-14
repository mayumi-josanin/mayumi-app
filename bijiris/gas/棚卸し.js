// データ棚卸し — 3つのスプレッドシートの構造を調べる道具。
//
//   棚卸しをする()  … 3ファイルを調べ、結果をドライブのテキストに書き出す
//
// 中身（お名前・電話番号など）は一切書き出さない。
// 出すのは「列の名前・何件埋まっているか・どんな型か」だけ。
// 移行先のデータベース設計に使うための下調べ。

var 棚卸し_対象 = [
  { 名前: 'まゆみ助産院_管理', id: '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w' },
  { 名前: 'ビジリス アンケート回答', id: '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E' },
  { 名前: 'ビジリス 会員別まとめ', id: '1KXs8e5W_iGtj8c4g2v7M4mhcWCkUUA6fxfs267o4lNw' },
];

// 値を見て型を言い当てる。中身そのものは返さない。
function 棚卸し_型を見る_(値) {
  if (値 instanceof Date) return '日付';
  if (typeof 値 === 'number') return '数値';
  if (typeof 値 === 'boolean') return '真偽';
  var s = String(値);
  if (!s) return '';
  if (/^\{[\s\S]*\}$|^\[[\s\S]*\]$/.test(s)) return 'JSON';
  if (/^https?:\/\//.test(s)) return 'URL';
  if (/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/.test(s)) return '日付文字';
  if (/^-?\d+(\.\d+)?$/.test(s)) return '数字文字';
  return '文字';
}

function 棚卸し_列を調べる_(sh) {
  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  if (行数 < 1 || 列数 < 1) return { 見出し: [], 明細: [], データ行: 0 };

  var 見出し = sh.getRange(1, 1, 1, 列数).getValues()[0];
  var データ行 = Math.max(0, 行数 - 1);
  if (データ行 === 0) {
    return {
      見出し: 見出し,
      データ行: 0,
      明細: 見出し.map(function (h, i) {
        return { 番号: i + 1, 名前: String(h || '(無題)'), 埋: 0, 型: '', 最大長: 0 };
      }),
    };
  }

  // 大きいシートでも読み切れるよう、上限を決めて読む
  var 読む行 = Math.min(データ行, 2000);
  var v = sh.getRange(2, 1, 読む行, 列数).getValues();

  var 明細 = [];
  for (var c = 0; c < 列数; c += 1) {
    var 埋 = 0;
    var 型集 = {};
    var 最大長 = 0;
    for (var r = 0; r < 読む行; r += 1) {
      var 値 = v[r][c];
      if (値 === '' || 値 === null || 値 === undefined) continue;
      埋 += 1;
      var t = 棚卸し_型を見る_(値);
      if (t) 型集[t] = (型集[t] || 0) + 1;
      var len = String(値).length;
      if (len > 最大長) 最大長 = len;
    }
    var 型 = Object.keys(型集)
      .sort(function (a, b) { return 型集[b] - 型集[a]; })
      .slice(0, 2)
      .join('/');
    明細.push({
      番号: c + 1,
      名前: String(見出し[c] || '(無題)'),
      埋: 埋,
      型: 型,
      最大長: 最大長,
    });
  }
  return { 見出し: 見出し, データ行: データ行, 明細: 明細, 読んだ行: 読む行 };
}

function 棚卸しをする() {
  var 出力 = [];
  var 書く = function (s) { 出力.push(s === undefined ? '' : String(s)); };

  書く('# データ棚卸し');
  書く('作成: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
  書く('');
  書く('中身（お名前・電話番号など）は含めていません。');
  書く('列の名前・埋まっている件数・型・最大文字数だけを出しています。');
  書く('');

  棚卸し_対象.forEach(function (t) {
    var ss = null;
    try {
      ss = SpreadsheetApp.openById(t.id);
    } catch (error) {
      書く('## ' + t.名前);
      書く('開けませんでした（権限なし）: ' + error.message);
      書く('');
      return;
    }

    書く('## ' + ss.getName());
    書く('ID: ' + t.id);
    var sheets = ss.getSheets();
    書く('シート数: ' + sheets.length);
    書く('');

    sheets.forEach(function (sh) {
      var 結果 = 棚卸し_列を調べる_(sh);
      書く('### ' + sh.getName());
      書く('データ行: ' + 結果.データ行 + ' / 列: ' + 結果.明細.length +
        (結果.読んだ行 && 結果.読んだ行 < 結果.データ行 ? '（先頭' + 結果.読んだ行 + '行を調査）' : ''));
      if (!結果.明細.length) { 書く('（空）'); 書く(''); return; }
      書く('');
      書く('| # | 列名 | 埋 | 型 | 最大長 |');
      書く('|---|------|----|----|--------|');
      結果.明細.forEach(function (d) {
        書く('| ' + d.番号 + ' | ' + d.名前 + ' | ' + d.埋 + ' | ' + d.型 + ' | ' + d.最大長 + ' |');
      });
      書く('');
    });
    書く('');
  });

  // スクリプトプロパティは「鍵の名前」だけ。値は書かない。
  書く('## スクリプトプロパティ（このビジリスGAS）');
  書く('値は出しません。鍵の名前だけです。');
  書く('');
  var keys = PropertiesService.getScriptProperties().getKeys().sort();
  keys.forEach(function (k) { 書く('- ' + k); });
  書く('');

  var 名前 = 'データ棚卸し_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm');
  // 新しいスプレッドシートに1行ずつ書く。あとから素の文章として読み出せるため。
  var ss = SpreadsheetApp.create(名前);
  var sh = ss.getSheets()[0];
  sh.setName('棚卸し');
  if (出力.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 出力.length - sh.getMaxRows());
  sh.getRange(1, 1, 出力.length, 1).setValues(出力.map(function (s) { return ["'" + s]; }));

  Logger.log('書き出しました: ' + 名前);
  Logger.log('ファイルID: ' + ss.getId());
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('行数: ' + 出力.length);
  return ss.getId();
}

// 調べ終わったら消す
function 棚卸しのファイルを消す() {
  var it = DriveApp.searchFiles('title contains "データ棚卸し_" and trashed = false');
  var n = 0;
  while (it.hasNext()) { it.next().setTrashed(true); n += 1; }
  Logger.log(n + '本をゴミ箱へ入れました');
}
