// もう使っていないシートに、触ったときの警告を付ける。
//
//   使っていませんの印をつける()   … 警告を付ける（何度実行してもよい）
//   使っていませんの印を外す()     … 元に戻す
//
// なぜ要るか。
// 2026-09-19、会員台帳をサーバーへ移した。9/21 に院長の判断で確定し、戻す道は閉じた。
// **それ以降、このシートを直しても、アプリにも管理画面にも届かない。**
// 見た目は今までどおりなので、うっかり直してしまう。直した本人は直ったつもりでいる。
// 気づくのは、お客様が「変わっていません」とおっしゃったとき。
//
// なぜ「行を足して大きく書く」ではないのか。
//   ・**行を足すと、全部の行が1つずれる。**行番号を鍵に使っている所があるので触れない
//   ・**シート名も変えられない。**仕掛けが名前でシートを探しており、
//     変えると「無いから作る」が動いて、空のシートができてしまう
// そこで、中身には一切触らずに「保護（警告だけ）」を掛ける。
// 触ろうとすると、下の文が出て、それでも直すかを聞かれる。読んで手を止められる。

var 印の合言葉 = '【使っていません】サーバーへ移しました';

var 印をつけるシート = [
  {
    名前: '会員データ',
    文: '【このシートは使っていません】\n\n' +
        '2026年9月19日に、会員の記録はサーバーへ移りました。\n' +
        'ここを直しても、お客様のアプリにも管理画面にも届きません。\n\n' +
        '会員の変更は、管理画面から行ってください。\n' +
        '  https://desktop-rmsk0vg.tail8efe0d.ts.net:10002/manage/members/\n\n' +
        'このシートは、移す前の姿を確かめるために残してあります。',
  },
];


function 使っていませんの印をつける() {
  // **getActive() は使えない。**この仕掛けはスプレッドシートに紐づいていない形なので null になる。
  // 他の所と同じ getOrCreateSpreadsheet()（スクリプトプロパティの SPREADSHEET_ID で開く）を使う。
  var ss = getOrCreateSpreadsheet();
  Logger.log('■ 使っていないシートに警告を付けます');
  Logger.log('  中身には一切触りません。行も足しません。名前も変えません。');
  Logger.log('');

  印をつけるシート.forEach(function (もの) {
    var sheet = ss.getSheetByName(もの.名前);
    if (!sheet) {
      Logger.log('  「' + もの.名前 + '」が見つかりません。飛ばします');
      return;
    }

    // 何度実行してもよいように、前に付けた分を外してから付け直す
    var 前 = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    前.forEach(function (p) {
      if (String(p.getDescription() || '').indexOf(印の合言葉) === 0) p.remove();
    });

    var 保護 = sheet.protect().setDescription(印の合言葉 + '\n\n' + もの.文);
    保護.setWarningOnly(true);   // **止めない。**警告を出して、それでも直せる
    sheet.setTabColor('#c53030');

    Logger.log('  「' + もの.名前 + '」に付けました（タブも赤くしました）');
  });

  Logger.log('');
  Logger.log('  触ろうとすると警告が出ます。**止めはしません。**読んで手を止めていただくためのものです。');
  Logger.log('  外すときは 使っていませんの印を外す() を実行してください。');
}


function 使っていませんの印を外す() {
  var ss = getOrCreateSpreadsheet();
  印をつけるシート.forEach(function (もの) {
    var sheet = ss.getSheetByName(もの.名前);
    if (!sheet) return;
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) {
      if (String(p.getDescription() || '').indexOf(印の合言葉) === 0) p.remove();
    });
    sheet.setTabColor(null);
    Logger.log('  「' + もの.名前 + '」の印を外しました');
  });
}
