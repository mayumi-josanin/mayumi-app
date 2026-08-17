// 移行のために書き出したJSONを、ドライブから片付ける道具。
//
//   書き出しJSONの下見()   … 何が残っているかを見るだけ
//   書き出しJSONを片付ける() … ゴミ箱へ入れる（30日は戻せる）
//
// なぜ要るのか:
//   移行のたびに、シートの中身をJSONにしてドライブへ置いている。
//   取り込みが済めば用済みだが、**置きっぱなしにすると個人情報が
//   ドライブに残り続ける。**とくに会員データのJSONにはパスコードが
//   平文で入っている。
//
// 消すのではなくゴミ箱へ入れる。30日は戻せるので、
// 「まだ取り込んでいなかった」と後から気づいても間に合う。

// **表を増やしたら、ここにも名前を足すこと。**
// 足し忘れると、書き出したJSONがドライブに残り続ける。
var 片付けJSON_対象の形 = /^(会員データ|測定履歴|掲載物|売上)_\d{8}-\d{4}\.json$/;

function 片付けJSON_集める_() {
  var 一覧 = [];
  var it = DriveApp.getFilesByType(MimeType.PLAIN_TEXT);
  while (it.hasNext()) {
    var f = it.next();
    if (f.isTrashed()) continue;
    if (!片付けJSON_対象の形.test(f.getName())) continue;
    一覧.push({
      file: f,
      名: f.getName(),
      大きさ: f.getSize(),
      日: f.getDateCreated(),
      // 会員データにはパスコードが平文で入っている。ひと目で分かるようにする。
      要注意: f.getName().indexOf('会員データ') === 0
    });
  }
  一覧.sort(function (a, b) { return a.名 < b.名 ? -1 : 1; });
  return 一覧;
}

function 書き出しJSONを片付ける() {
  var 一覧 = 片付けJSON_集める_();
  if (!一覧.length) {
    Logger.log('■ 残っていません。何もしませんでした。');
    return;
  }
  一覧.forEach(function (x) {
    x.file.setTrashed(true);
    Logger.log('  ゴミ箱へ: ' + x.名);
  });
  Logger.log('');
  Logger.log('■ ' + 一覧.length + '件をゴミ箱へ入れました。');
  Logger.log('  30日は戻せます。完全に消すときはドライブのゴミ箱を空にしてください。');
}

function 書き出しJSONの下見() {
  var 一覧 = 片付けJSON_集める_();
  Logger.log('■ 書き出したJSON: ' + 一覧.length + '件');
  Logger.log('');
  if (!一覧.length) {
    Logger.log('  残っていません。片付け済みです。');
    return;
  }
  一覧.forEach(function (x) {
    Logger.log('  ' + (x.要注意 ? '**' : '  ') + x.名 +
      '  ' + Math.round(x.大きさ / 1024) + 'KB  ' +
      Utilities.formatDate(x.日, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') +
      (x.要注意 ? '  ← パスコードが平文で入っています**' : ''));
  });
  Logger.log('');
  Logger.log('  よければ 書き出しJSONを片付ける() を実行してください。');
  Logger.log('  ゴミ箱へ入れるだけなので、30日は戻せます。');
}
