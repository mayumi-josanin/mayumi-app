// 会員データの **メモ（7列目）** と **通知の届け先（8列目）** だけを書き出す道具。
//
//   落ちた2列の下見()    … 何件埋まっているかを見るだけ
//   落ちた2列を書き出す() … JSONにしてドライブへ置く
//
// **どちらも読むだけ。**スプレッドシートは一切変えない。
//
// ## なぜ、この2列だけを別に書き出すのか
//
// 最初の「会員データを書き出す()」で、この2列が落ちていた（2026-08-23に判明）。
//
//   ・メモ … そもそも書き出していなかった
//   ・通知の届け先 … 書き出してはいたが、取り込み側が `真偽()` に通していた。
//     購読IDは "true" でも "1" でもないので **偽に落ちていた。**
//     153名のうち、サーバーで通知オンになっていたのは37名だけだった。
//
// **通知の届け先は、配信の宛先そのもの。**
// GAS の getPushUsers() が、この値を `subscription:` として渡している。
// 失ったままサーバーへ切り替えると、**お知らせが誰にも届かなくなる。**
//
// ## 扱いの注意
//
// メモには、受付が残した覚え書きが入る。**お客様ご本人には見えない内容。**
// 取り込みが済んだら、ドライブからも手元からも消すこと。
//
// 使いかた:
//   1. 落ちた2列の下見() で件数を確かめる
//   2. 落ちた2列を書き出す() を実行し、ログのURLからダウンロードする
//   3. サーバー側で
//        python manage.py 落ちた2列を取り込む 落ちた2列.json --下見
//   4. 件数が合っていれば --下見 を外して実行

var 落ち2列_シート名 = '会員データ';

function 落ちた2列を書き出す() {
  var r = 落ち2列_集める_();
  var 中身 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    件数: r.行.length,
    members: r.行
  }, null, 2);

  var 名前 = '落ちた2列_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 中身, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('');
  Logger.log('  会員: ' + r.行.length + '名');
  Logger.log('  メモあり: ' + r.メモ数 + '名 / 通知の届け先あり: ' + r.届け先数 + '名');
  Logger.log('');
  Logger.log('  ※ メモには受付の覚え書きが入っています。');
  Logger.log('     取り込みが済んだら、ドライブからも手元からも消してください。');
}

function 落ちた2列の下見() {
  var r = 落ち2列_集める_();
  Logger.log('■ 会員データ: ' + r.行.length + '名');
  Logger.log('');
  Logger.log('  メモあり:       ' + r.メモ数 + '名');
  Logger.log('  通知の届け先あり: ' + r.届け先数 + '名');
  Logger.log('');
  Logger.log('■ 通知の届け先は、どんな中身か');
  Logger.log('  （**中身そのものは出しません。**形だけ数えます）');
  Logger.log('    文字列の "true": ' + r.内訳.trueだけ + '名');
  Logger.log('    購読IDらしいもの: ' + r.内訳.購読ID + '名');
  Logger.log('    その他:          ' + r.内訳.その他 + '名');
  Logger.log('');
  Logger.log('  ※ サーバーには現在 37名しか通知オンで入っていません。');
  Logger.log('     ここの「通知の届け先あり」と食い違うなら、それが取りこぼしです。');
}

function 落ち2列_集める_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(落ち2列_シート名);
  if (!sheet) throw new Error('「' + 落ち2列_シート名 + '」が見つかりません。');

  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) return { 行: [], メモ数: 0, 届け先数: 0, 内訳: { trueだけ: 0, 購読ID: 0, その他: 0 } };

  // **必要な3列だけを読む。**ほかの列は読まない（読めば控えに残る）。
  var 値 = sheet.getRange(2, 1, 最終行 - 1, 8).getValues();

  var 行 = [];
  var メモ数 = 0;
  var 届け先数 = 0;
  var 内訳 = { trueだけ: 0, 購読ID: 0, その他: 0 };

  値.forEach(function (row) {
    var 会員ID = String(row[0] || '').trim();
    if (!会員ID) return;

    var メモ = row[6];
    var 届け先 = row[7];

    // 真偽値の false はセルにそのまま入っている。文字にすると "false" になるので、
    // **ここで空文字に寄せる。**「届け先なし」と同じ扱いにする。
    var 届け先文字 = (届け先 === false || 届け先 === null || 届け先 === undefined) ? '' : String(届け先).trim();
    if (届け先文字.toLowerCase() === 'false') 届け先文字 = '';

    var メモ文字 = (メモ === null || メモ === undefined) ? '' : String(メモ);

    if (メモ文字.trim()) メモ数++;
    if (届け先文字) {
      届け先数++;
      if (届け先文字.toLowerCase() === 'true') 内訳.trueだけ++;
      else if (/^[A-Za-z0-9_-]{16,}$/.test(届け先文字)) 内訳.購読ID++;
      else 内訳.その他++;
    }

    行.push({ memberId: 会員ID, memo: メモ文字, pushSubscription: 届け先文字 });
  });

  return { 行: 行, メモ数: メモ数, 届け先数: 届け先数, 内訳: 内訳 };
}
