// お客様が実際にアプリを開けているかを、日ごとに数える道具。**読むだけ。**
//
//   使われているか確かめる() … 端末の「最後に開いた日時」を日ごとに数える
//
// ## なぜ要るのか
//
// 「画面が出る」「?action=… が status:ok を返す」は、**壊れていない証拠**でしかない。
// **お客様が実際に入れているか**は、別に確かめる必要がある。
//
// アプリは開くたびに端末の記録を合わせに来る（syncUserDeviceSession）。
// その `lastSeenAt` を日ごとに数えれば、**何人が実際に開いたか**が分かる。
//
// デプロイや切り替えの前後で数を見比べれば、**壊したかどうかが分かる。**
// （2026-08-24 に @240 を出したので、その前後を見る）
//
// 操作履歴には残らない種類の通信なので、ここでしか数えられない。

function 使われているか確かめる() {
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet());
  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) { Logger.log('■ 会員データに行がありません。'); return; }

  var 値 = sheet.getRange(2, 1, 最終行 - 1, USER_HEADERS.length).getValues();

  var 日ごと = {};      // '2026-08-25' → 会員IDの集合
  var 端末の総数 = 0;
  var 記録のある会員 = 0;

  値.forEach(function (row) {
    var 会員ID = String(row[USER_COL.MEMBER_ID - 1] || '').trim();
    if (!会員ID) return;

    var 端末たち;
    try {
      端末たち = JSON.parse(String(row[USER_COL.DEVICE_SESSIONS - 1] || '[]'));
    } catch (e) { return; }
    if (!Array.isArray(端末たち) || !端末たち.length) return;

    記録のある会員++;
    端末の総数 += 端末たち.length;

    端末たち.forEach(function (d) {
      var t = String((d && d.lastSeenAt) || '').slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) return;
      if (!日ごと[t]) 日ごと[t] = {};
      日ごと[t][会員ID] = true;
    });
  });

  var 日 = Object.keys(日ごと).sort().reverse();

  Logger.log('■ 会員 ' + 値.length + '名 / 端末の記録がある方 ' + 記録のある会員 + '名 / 端末 ' + 端末の総数 + '台');
  Logger.log('');
  Logger.log('■ 日ごとに「アプリを開いた方」の人数（新しい順に21日分）');
  Logger.log('   ※ 端末の記録は1台につき最後の1回だけ残るので、');
  Logger.log('      **古い日は少なく見えます。**直近の数日を見てください。');
  Logger.log('');

  日.slice(0, 21).forEach(function (t) {
    var n = Object.keys(日ごと[t]).length;
    Logger.log('     ' + t + '  ' + Array(n + 1).join('■') + ' ' + n + '名');
  });

  if (!日.length) {
    Logger.log('     **1件もありません。**これは異常です。');
    return;
  }

  Logger.log('');
  Logger.log('■ いちばん新しい記録: ' + 日[0]);
  Logger.log('');
  Logger.log('   きょう・きのうに記録があれば、**お客様は問題なく入れています。**');
  Logger.log('   2026-08-24 に @240 を出したので、その前後で途切れていないかを見てください。');
}
