// ビジリスが実際に使われているかを、日ごとに数える道具。**読むだけ。**
//
//   ビジリスが使われているか() … 回答一覧の提出日を日ごとに数える
//
// ## なぜ要るのか
//
// 「画面が出る」「?action=surveys が返る」は、**壊れていない証拠**でしかない。
// **お客様が実際に回答できているか**は、別に確かめる必要がある。
//
// まゆみ側は端末の記録（lastSeenAt）で数えられるが、ビジリスは
// 回答そのものが残るので、**提出日を数えるのがいちばん確か。**
//
// 中身（お名前・回答の本文）は一切出さない。日付と件数だけ。

var ビ使用_ファイルID = '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E';
var ビ使用_シート名 = '回答一覧';

// **並びはわざとこの順です。**GASのエディタは開いたファイルの先頭の関数を
// 既定で選びます。選び間違いを防ぐため、いま確かめたい方を先頭に置いています。

// 回答が本当に「回答一覧」だけに入っているのかを確かめる道具。**読むだけ。**
//
// 回答一覧が1件しかなかったので（2026-08-25）、
// **別のシートに入っているのではないか**を疑って作った。
// 「1件しかない」と言う前に、**同じファイルの全シートを数える。**
//
// 中身は出さない。シート名と行数だけ。
function ビジリスの入れ物を数える() {
  var book = SpreadsheetApp.openById(ビ使用_ファイルID);
  var シートたち = book.getSheets();

  Logger.log('■ ファイル: ' + book.getName());
  Logger.log('   シート数: ' + シートたち.length);
  Logger.log('');
  Logger.log('   シート名' + Array(30).join(' ') + '行数（見出しを除く）');

  var 合計 = 0;
  シートたち.forEach(function (sh) {
    var 行 = Math.max(0, sh.getLastRow() - 1);
    合計 += 行;
    var 名 = sh.getName();
    var 詰め = Array(Math.max(1, 34 - 名.length)).join(' ');
    Logger.log('     ' + 名 + 詰め + 行);
  });
  Logger.log('');
  Logger.log('   見出しを除いた行の合計: ' + 合計);
}

function ビジリスが使われているか() {
  var sh = SpreadsheetApp.openById(ビ使用_ファイルID).getSheetByName(ビ使用_シート名);
  if (!sh) { Logger.log('■ 「' + ビ使用_シート名 + '」が見つかりません。'); return; }

  var 最終行 = sh.getLastRow();
  if (最終行 < 2) { Logger.log('■ 回答が1件もありません。'); return; }

  var 見出し = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  // 日付らしい列を、見出しの名前から探す。位置で決め打ちにしない
  // （列が足されたときに黙ってずれるため）。
  var 日付列 = -1;
  for (var i = 0; i < 見出し.length; i++) {
    var h = String(見出し[i] || '');
    if (h.indexOf('日時') >= 0 || h.indexOf('提出') >= 0 || h.indexOf('タイムスタンプ') >= 0) {
      日付列 = i; break;
    }
  }
  if (日付列 < 0) 日付列 = 0;   // 見つからなければ1列目

  Logger.log('■ 回答一覧: ' + (最終行 - 1) + '件');
  Logger.log('   日付として見る列: 「' + 見出し[日付列] + '」（' + (日付列 + 1) + '列目）');
  Logger.log('');

  var 値 = sh.getRange(2, 日付列 + 1, 最終行 - 1, 1).getValues();
  var 日ごと = {};
  var 読めない = 0;

  値.forEach(function (r) {
    var v = r[0];
    var t = '';
    if (v instanceof Date) {
      t = Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
    } else {
      var s = String(v || '');
      var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
      if (m) {
        t = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
      }
    }
    if (!t) { 読めない++; return; }
    日ごと[t] = (日ごと[t] || 0) + 1;
  });

  var 日 = Object.keys(日ごと).sort().reverse();
  Logger.log('■ 日ごとの回答数（新しい順に21日分）');
  日.slice(0, 21).forEach(function (t) {
    Logger.log('     ' + t + '  ' + Array(日ごと[t] + 1).join('■') + ' ' + 日ごと[t] + '件');
  });
  Logger.log('');
  Logger.log('   いちばん新しい回答: ' + (日[0] || '（なし）'));
  Logger.log('   記録のある日数: ' + 日.length + '日');
  if (読めない) Logger.log('   日付として読めなかった行: ' + 読めない + '件');
}


