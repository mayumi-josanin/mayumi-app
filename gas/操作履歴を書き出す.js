// 操作履歴（ADMIN_AUDIT_LOG）を、データベースへ取り込める形（JSON）で書き出す。
//
//   操作履歴を書き出す()   … JSONにしてドライブへ置く
//
// **読むだけ。**元のスプレッドシートは一切変えない。
//
// **9,373行のうち「人の操作」794件だけを書き出す**（院長の判断）。
// 残り8,579件（91.5%）はアプリが勝手に行う自動処理の記録で、誰も見ない。
// 全部移すとデータ量が11倍になり、控えも大きくなる。
//
//   syncUserDeviceSession  7,788件（83.1%）  端末の記録合わせ
//   syncUserRewardStatus     791件（ 8.4%）  特典状態の同期
//
// **除いた件数は必ずログに出す。**黙って減らすと、あとから
// 「なぜ件数が合わないのか」を探すことになる。

var 履歴書出_除く種別 = ['syncUserDeviceSession', 'syncUserRewardStatus'];

// 1回に書き出す上限。9,373行を全部読むと6分の制限に近づくため。
// いまは794件しか出ないので余裕があるが、増えたときに黙って
// 途中で切れないよう、**切れたことが分かる作りにする。**
var 履歴書出_上限 = 5000;

function 操作履歴を書き出す() {
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.ADMIN_AUDIT_LOG);
  if (!sheet) {
    Logger.log('■ 「' + SHEETS.ADMIN_AUDIT_LOG + '」シートがありません。');
    return;
  }

  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) {
    Logger.log('■ 空です。書き出すものがありません。');
    return;
  }

  var v = sheet.getRange(1, 1, 最終行, sheet.getLastColumn()).getValues();
  var 見出し = v[0].map(function (h) { return String(h).trim(); });
  var i = {
    日時: 見出し.indexOf('日時'),
    種別: 見出し.indexOf('種別'),
    結果: 見出し.indexOf('結果'),
    対象: 見出し.indexOf('対象'),
    概要: 見出し.indexOf('概要'),
    操作者: 見出し.indexOf('操作者'),
    詳細JSON: 見出し.indexOf('詳細JSON')
  };
  var 無い列 = Object.keys(i).filter(function (k) { return i[k] < 0; });

  var 行 = [];
  var 除いた = {};
  var 種別ごと = {};
  var 空行 = 0;
  var 打ち切った = false;

  for (var r = 1; r < v.length; r += 1) {
    var 種 = i.種別 >= 0 ? 掲載書出_文字_(v[r][i.種別]) : '';
    var 日 = i.日時 >= 0 ? 掲載書出_日時_(v[r][i.日時]) : null;

    // 種別も日時も無い行は、書きかけの空行。
    if (!種 && !日) { 空行 += 1; continue; }

    if (履歴書出_除く種別.indexOf(種) >= 0) {
      除いた[種] = (除いた[種] || 0) + 1;
      continue;
    }

    if (行.length >= 履歴書出_上限) { 打ち切った = true; break; }

    種別ごと[種] = (種別ごと[種] || 0) + 1;

    var 詳細 = i.詳細JSON >= 0 ? 掲載書出_文字_(v[r][i.詳細JSON]) : '';
    行.push({
      row: r + 1,
      happened_at: 日,
      kind: 種,
      result: i.結果 >= 0 ? 掲載書出_文字_(v[r][i.結果]) : '',
      target: i.対象 >= 0 ? 掲載書出_文字_(v[r][i.対象]) : '',
      // **「概要」は種別によって埋まり方が違う。**先頭200行では空に
      // 見えるが、deleteOrders は30/30すべて埋まっている。空欄のまま渡す。
      summary: i.概要 >= 0 ? 掲載書出_文字_(v[r][i.概要]) : '',
      operator: i.操作者 >= 0 ? 掲載書出_文字_(v[r][i.操作者]) : '',
      // 読める形にするのは取り込み側でやる。ここでは文字のまま渡す。
      // ここで JSON.parse して失敗すると、その行ごと落ちてしまう。
      detail_raw: 詳細
    });
  }

  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    シートの行数: v.length - 1,
    除いた種別: 履歴書出_除く種別,
    audit_logs: 行
  }, null, 2);

  var 名前 = '操作履歴_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024 * 10) / 10 + 'KB');
  Logger.log('');

  Logger.log('■ シートは ' + (v.length - 1) + '行');
  var 除いた合計 = 0;
  履歴書出_除く種別.forEach(function (k) {
    var n = 除いた[k] || 0;
    除いた合計 += n;
    Logger.log('    除いた: ' + k + ' … ' + n + '件');
  });
  Logger.log('    除いた合計: ' + 除いた合計 + '件（自動処理。誰も見ない記録）');
  if (空行) Logger.log('    空行: ' + 空行 + '件');
  Logger.log('    **書き出した（人の操作）: ' + 行.length + '件**');

  // 足し算が合うかを、その場で確かめる。合わなければ数え漏れがある。
  var 計 = 行.length + 除いた合計 + 空行;
  if (計 !== (v.length - 1) && !打ち切った) {
    Logger.log('    **足すと ' + 計 + 'で、シートの ' + (v.length - 1) + '行と合いません。**');
  } else if (!打ち切った) {
    Logger.log('    （' + 行.length + ' + ' + 除いた合計 +
      (空行 ? ' + ' + 空行 : '') + ' = ' + 計 + '行 で合っています）');
  }
  Logger.log('');

  if (打ち切った) {
    Logger.log('  **上限 ' + 履歴書出_上限 + '件で打ち切りました。全部は出ていません。**');
    Logger.log('    履歴書出_上限 を増やして、もう一度実行してください。');
    Logger.log('');
  }

  Logger.log('■ 書き出した種別（多い順）:');
  Object.keys(種別ごと).sort(function (a, b) { return 種別ごと[b] - 種別ごと[a]; })
    .forEach(function (k) { Logger.log('    ' + k + ' … ' + 種別ごと[k] + '件'); });
  Logger.log('');

  var 失敗 = 行.filter(function (x) { return x.result === '失敗'; }).length;
  var 概要あり = 行.filter(function (x) { return x.summary; }).length;
  var 対象あり = 行.filter(function (x) { return x.target; }).length;
  Logger.log('    結果が「失敗」: ' + 失敗 + '件');
  Logger.log('    概要がある: ' + 概要あり + '件 / 対象がある: ' + 対象あり + '件');
  Logger.log('');

  if (無い列.length) {
    Logger.log('  **見つからない列: ' + 無い列.join('・') + '**');
    Logger.log('');
  }
  Logger.log('  ※ 読むだけです。シートは変えていません。');
  Logger.log('  ※ **詳細JSONに会員のお名前や電話番号が入っている場合があります。**');
  Logger.log('     取り込みが済んだら必ず 書き出しJSONを片付ける() で消してください。');
}
