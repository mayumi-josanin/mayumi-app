// サーバーへの道が、いまどのくらいの割合で通るかを測る。**何も書きません。**
//
//   道の確からしさを測る()   … 20回つないで、通った回数と落ちた回数を出す
//
// なぜ要るか。
// 2026-09-15 に GAS からだけ名前引き（DNS）が90分ぜんぶ失敗した。自然に戻ったので
// 一度きりと見ていたが、9/18 に失敗の記録を見ると **17日の昼から18日の朝まで
// 毎回のように落ちていた**。読みは「前回の答え」でしのげていたので、誰も気づかなかった。
//
// 9/19 に会員をサーバーへ移すと、お客様のログインと保存がこの道を通る。
// **保存は前回の答えでしのげない。**その場でお客様に失敗が見える。
// だから「たまに落ちる」のか「半分落ちる」のかで、移すかどうかの判断が変わる。
// 推し量らずに、数える。
//
// やり直しの仕掛け（渡す_取りに行く_）は通さずに、**素のまま**測る。
// 素の通り具合が分からないと、やり直しで足りているのかが判断できない。

function 道の確からしさを測る() {
  var 回数 = 20;
  var url = 渡す_URL_() + '?action=getNews';

  Logger.log('■ サーバーへの道の確からしさ');
  Logger.log('  行き先: ' + url);
  Logger.log('  ' + 回数 + '回つなぎます（やり直しはしません。素の通り具合を見ます）');
  Logger.log('');

  var 通った = 0;
  var 名前引きで落ちた = 0;
  var その他で落ちた = 0;
  var かかった = [];
  var 落ちた訳 = {};

  for (var i = 0; i < 回数; i++) {
    var 始め = new Date().getTime();
    try {
      var res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, escaping: false });
      var ミリ秒 = new Date().getTime() - 始め;
      if (res.getResponseCode() === 200) {
        通った++;
        かかった.push(ミリ秒);
      } else {
        その他で落ちた++;
        落ちた訳['HTTP ' + res.getResponseCode()] = (落ちた訳['HTTP ' + res.getResponseCode()] || 0) + 1;
      }
    } catch (e) {
      var 文 = String(e);
      if (/DNS|アドレス/i.test(文)) {
        名前引きで落ちた++;
      } else {
        その他で落ちた++;
      }
      var 鍵 = 文.slice(0, 60);
      落ちた訳[鍵] = (落ちた訳[鍵] || 0) + 1;
    }
    Utilities.sleep(1500);   // 続けて叩きすぎない
  }

  var 割合 = Math.round((通った / 回数) * 1000) / 10;
  Logger.log('  通った        : ' + 通った + ' / ' + 回数 + '（' + 割合 + '%）');
  Logger.log('  名前引きで落ちた: ' + 名前引きで落ちた);
  Logger.log('  その他で落ちた  : ' + その他で落ちた);
  if (かかった.length) {
    かかった.sort(function (a, b) { return a - b; });
    Logger.log('  かかった時間    : 早い ' + かかった[0] + 'ms ／ 真ん中 ' +
      かかった[Math.floor(かかった.length / 2)] + 'ms ／ 遅い ' + かかった[かかった.length - 1] + 'ms');
  }
  if (Object.keys(落ちた訳).length) {
    Logger.log('');
    Logger.log('  落ちた訳:');
    Object.keys(落ちた訳).forEach(function (k) { Logger.log('    ' + 落ちた訳[k] + '回  ' + k); });
  }
  Logger.log('');
  if (通った === 回数) {
    Logger.log('  → いまは全部通っています。');
  } else if (割合 >= 90) {
    Logger.log('  → たまに落ちます。やり直し（4回）でほぼ拾えます。');
  } else {
    Logger.log('  → **よく落ちます。**会員を移すと、お客様の保存が失敗します。移すのを見送る判断を。');
  }
  Logger.log('  ※ これは見るだけです。何も書いていません。');
}
