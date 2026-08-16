// 動作確認で作られた会員の行を片付ける道具。
//
//   テスト会員の下見()   … 対象の行の中身を全部見るだけ（何も消さない）
//   テスト会員を片付ける() … 控えを取ってから消す
//
// なぜ要るのか:
//   まゆみ会員データに「テスト」という行が残っていた（2026-08-16 に3件消した）。
//
//   最初「ビジリスに記録が無いから空の行だろう」と判断しかけたが、下見を通したら
//   実在の電話番号・住所・生年月日が入り、管理者権限と登録済みのiPhoneまであった。
//   **他の表に記録が無いことは、会員データが空である根拠にならない。**
//   お名前や別の表だけで判断せず、必ずその行の全項目を見ること。
//
// **消す前に必ず下見を通すこと。**
// お名前だけで判断すると、本当に「テスト」というお名前の方がいらしたときに
// 消してしまう。登録日時・連絡先・スタンプ数・最終オンライン日時まで見て、
// 実在の方でないことを確かめてから消す。

var 片付け_消す氏名 = ['テスト'];

function テスト会員の下見() {
  var r = 片付け_探す_();
  if (!r.行.length) {
    Logger.log('■ 対象の行はありません。');
    return;
  }

  Logger.log('■ 対象: ' + r.行.length + '件');
  Logger.log('');
  r.行.forEach(function (x) {
    Logger.log('  ' + x.行番号 + '行目');
    x.中身.forEach(function (c) {
      Logger.log('      ' + c.見出し + ': ' + (c.値 === '' ? '（空）' : c.値));
    });
    Logger.log('');
  });

  Logger.log('■ 実在の方でないことを、次で確かめてください。');
  Logger.log('    ・登録日時が動作確認をした日か');
  Logger.log('    ・電話番号・生年月日・住所が空か');
  Logger.log('    ・スタンプ数が0か');
  Logger.log('    ・最終オンライン日時が無いか');
  Logger.log('');
  Logger.log('  よければ テスト会員を片付ける() を実行してください（先に控えを取ります）。');
}

function テスト会員を片付ける() {
  var r = 片付け_探す_();
  if (!r.行.length) {
    Logger.log('■ 対象の行はありません。何もしませんでした。');
    return;
  }

  // 控え。スプレッドシートごと複製する。
  var book = r.sheet.getParent();
  var 控えの名 = '【控え】' + book.getName() + ' ' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + ' テスト会員を消す前';
  var 控え = DriveApp.getFileById(book.getId()).makeCopy(控えの名);
  Logger.log('■ 控えを取りました: ' + 控えの名);
  Logger.log('  ' + 控え.getUrl());
  Logger.log('');

  // 下の行から消す。上から消すと、消したぶん行番号がずれる。
  var 行番号 = r.行.map(function (x) { return x.行番号; }).sort(function (a, b) { return b - a; });
  行番号.forEach(function (n) {
    r.sheet.deleteRow(n);
    Logger.log('  ' + n + '行目を消しました');
  });

  Logger.log('');
  Logger.log('■ ' + 行番号.length + '件を消しました。会員データはいま ' +
    (r.sheet.getLastRow() - 1) + '行です。');
}

function 片付け_探す_() {
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet());
  var 最終 = sheet.getLastRow();
  if (最終 < 2) return { sheet: sheet, 行: [] };

  var 値 = sheet.getRange(2, 1, 最終 - 1, USER_HEADERS.length).getValues();
  var 行 = [];

  値.forEach(function (row, i) {
    var 氏名 = String(row[USER_COL.NAME - 1] || '').trim();
    if (片付け_消す氏名.indexOf(氏名) < 0) return;

    // 中身を全部そのまま見せる。一部だけ見せると、見えていない列に
    // 実在の方の情報が入っていても気づけない。
    var 中身 = USER_HEADERS.map(function (見出し, 列) {
      var v = row[列];
      var s = v instanceof Date
        ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm')
        : String(v === null || v === undefined ? '' : v);
      // パスコードやハッシュは、値そのものではなく有無だけ出す。
      if (['パスコード', 'パスワードハッシュ', 'パスワードソルト'].indexOf(見出し) >= 0) {
        s = s ? '（設定あり）' : '';
      }
      if (s.length > 120) s = s.slice(0, 120) + '…';
      return { 見出し: 見出し, 値: s };
    });

    行.push({ 行番号: i + 2, 氏名: 氏名, 中身: 中身 });
  });

  return { sheet: sheet, 行: 行 };
}
