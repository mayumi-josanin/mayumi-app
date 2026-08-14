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

// ---------- 会員の入り口の点検 ----------
//
// 「いま入れない方が何人いるか」「自力で復旧できない方が何人いるか」を数える。
// お名前などの中身は出さない。出すのは件数と、直すべき行の番号だけ。

var 点検_まゆみファイルID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';

function 会員の入り口を点検する() {
  var ss = SpreadsheetApp.openById(点検_まゆみファイルID);
  var sh = ss.getSheetByName('会員データ');
  if (!sh) { Logger.log('会員データのシートが見つかりません'); return; }

  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  var v = sh.getRange(1, 1, 行数, 列数).getValues();
  var 見出し = v[0];

  var 位置 = function (名) {
    var i = 見出し.indexOf(名);
    return i;
  };
  var C = {
    id: 位置('ID'), 名: 位置('氏名'), 電話: 位置('電話番号'), 誕生: 位置('生年月日'),
    パス: 位置('パスコード'), ハッシュ: 位置('パスワードハッシュ'), 権限: 位置('権限'),
    削除: 位置('削除状態'), 最終: 位置('最終オンライン日時'), 登録: 位置('登録日時'),
  };

  var 空 = function (x) { return x === '' || x === null || x === undefined; };

  var 合計 = 0;
  var 名前なし = [];        // 行番号
  var 入れない = 0;          // パスコードもハッシュも無い
  var 復旧できない = 0;      // 入れない かつ 電話も生年月日も無い
  var 復旧できる = 0;        // 入れない が 電話か生年月日がある
  var 一度も開いていない = 0;
  var 管理者 = 0;
  var 中身が空の行 = [];

  for (var r = 1; r < v.length; r += 1) {
    var row = v[r];
    var 何か入っている = row.some(function (x) { return !空(x); });
    if (!何か入っている) continue;

    合計 += 1;
    if (C.権限 >= 0 && String(row[C.権限] || '').trim() === '管理者') 管理者 += 1;

    var 名 = C.名 >= 0 ? String(row[C.名] || '').trim() : '';
    if (!名) {
      名前なし.push(r + 1);
      // お名前が無い行は、ほかに何が入っているかも見る
      var 埋まっている = row.filter(function (x) { return !空(x); }).length;
      if (埋まっている <= 2) 中身が空の行.push(r + 1);
    }

    var パス = C.パス >= 0 ? String(row[C.パス] || '').trim() : '';
    var ハッシュ = C.ハッシュ >= 0 ? String(row[C.ハッシュ] || '').trim() : '';
    if (!パス && !ハッシュ) {
      入れない += 1;
      var 電話 = C.電話 >= 0 ? String(row[C.電話] || '').trim() : '';
      var 誕生 = C.誕生 >= 0 ? row[C.誕生] : '';
      if (!電話 && 空(誕生)) 復旧できない += 1; else 復旧できる += 1;
    }

    if (C.最終 >= 0 && 空(row[C.最終])) 一度も開いていない += 1;
  }

  Logger.log('■ 会員データの点検');
  Logger.log('  会員の行: ' + 合計 + '（うち管理者 ' + 管理者 + '）');
  Logger.log('');
  Logger.log('▼ 入り口に入れない方');
  Logger.log('  パスコードもパスワードも無い: ' + 入れない + '名');
  Logger.log('    └ 電話か生年月日があり、自力で復旧できる: ' + 復旧できる + '名');
  Logger.log('    └ どちらも無く、受付対応が要る: ' + 復旧できない + '名');
  Logger.log('');
  Logger.log('▼ お名前が空の行');
  Logger.log('  ' + 名前なし.length + '行  行番号: ' + (名前なし.join(', ') || 'なし'));
  Logger.log('  うち、ほぼ空っぽの行: ' + 中身が空の行.length + '  行番号: ' + (中身が空の行.join(', ') || 'なし'));
  Logger.log('');
  Logger.log('▼ 参考');
  Logger.log('  最終オンライン日時が空（この記録が始まって以降に開いていない）: ' + 一度も開いていない + '名');
}

// ---------- 操作履歴の中身を調べる ----------
//
// 何がどれだけ記録されているかを数える。お名前などの中身は出さない。

function 操作履歴を調べる() {
  var sh = SpreadsheetApp.openById(点検_まゆみファイルID).getSheetByName('ADMIN_AUDIT_LOG');
  if (!sh) { Logger.log('ADMIN_AUDIT_LOG が見つかりません'); return; }

  var 行数 = sh.getLastRow();
  if (行数 < 2) { Logger.log('記録はありません'); return; }
  var v = sh.getRange(2, 1, 行数 - 1, 6).getValues();   // 詳細JSONは読まない（重いので）

  var 種別 = {};
  var 結果 = {};
  var 操作者 = {};
  var 日ごと = {};
  var 最古 = '';
  var 最新 = '';

  v.forEach(function (r) {
    var 日時 = String(r[0] || '');
    if (日時) {
      if (!最古 || 日時 < 最古) 最古 = 日時;
      if (!最新 || 日時 > 最新) 最新 = 日時;
      var 日 = 日時.slice(0, 10);
      日ごと[日] = (日ごと[日] || 0) + 1;
    }
    var t = String(r[1] || '(空)');
    種別[t] = (種別[t] || 0) + 1;
    var k = String(r[2] || '(空)');
    結果[k] = (結果[k] || 0) + 1;
    var o = String(r[5] || '(空)');
    操作者[o] = (操作者[o] || 0) + 1;
  });

  Logger.log('■ 操作履歴（ADMIN_AUDIT_LOG）');
  Logger.log('  記録数: ' + v.length + '行');
  Logger.log('  いちばん古い: ' + 最古);
  Logger.log('  いちばん新しい: ' + 最新);
  Logger.log('  記録のある日数: ' + Object.keys(日ごと).length + '日');
  Logger.log('');

  var 並べる = function (見出し, 集計, 上限) {
    Logger.log('▼ ' + 見出し);
    Object.keys(集計)
      .sort(function (a, b) { return 集計[b] - 集計[a]; })
      .slice(0, 上限 || 100)
      .forEach(function (k) {
        Logger.log('  ' + k + ' … ' + 集計[k] + '件');
      });
    Logger.log('');
  };

  並べる('結果', 結果);
  並べる('操作者', 操作者);
  並べる('種別（多い順・上位20）', 種別, 20);
}

// 調べ終わったら消す
function 棚卸しのファイルを消す() {
  var it = DriveApp.searchFiles('title contains "データ棚卸し_" and trashed = false');
  var n = 0;
  while (it.hasNext()) { it.next().setTrashed(true); n += 1; }
  Logger.log(n + '本をゴミ箱へ入れました');
}
