// GAS からサーバーの名前が引けないとき（2026-09-15 16:24〜「DNS error」が続いた）に、
// どこで止まっているかを切り分ける道具。
//
//   名前引きを確かめる()   … 3つの相手に GAS から届くかを見る（**読むだけ**）
//
//   ① Google の DNS（dns.google）に「この名前は引けるか」を聞く … GAS の外向き通信そのものの確認
//   ② サーバーの名前で取りに行く                                … いつもの道
//   ③ サーバーの別名（Funnel の共通名）で取りに行く                … 名前だけの問題かの確認

function 名前引きを確かめる() {
  var 名 = 'mayumi-api.tail8efe0d.ts.net';
  var 相手 = [
    ['① dns.google で名前を引く', 'https://dns.google/resolve?name=' + 名 + '&type=A'],
    ['② サーバー（いつもの名前）', 'https://' + 名 + '/api?action=getNews'],
    ['③ 外の一般サイト（比較用）', 'https://www.example.com/'],
  ];
  Logger.log('■ GAS からの名前引き（' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'HH:mm:ss') + '）');
  相手.forEach(function (x) {
    var t = new Date().getTime();
    try {
      var res = UrlFetchApp.fetch(x[1], { muteHttpExceptions: true, followRedirects: true });
      var 本文 = String(res.getContentText() || '');
      Logger.log('  ' + x[0] + ' → HTTP ' + res.getResponseCode() + '（' + (new Date().getTime() - t) + 'ミリ秒）' +
        (x[1].indexOf('dns.google') !== -1 ? ' 答え: ' + 本文.slice(0, 160).replace(/\s+/g, ' ') : ' 先頭: ' + 本文.slice(0, 40).replace(/\s+/g, ' ')));
    } catch (e) {
      Logger.log('  ' + x[0] + ' → **しくじり: ' + String(e).slice(0, 120) + '**');
    }
  });
  Logger.log('');
  Logger.log('  ①が届いて②だけ止まるなら、Google の側でこの名前だけが引けていない（時間で戻ることが多い）。');
  Logger.log('  ①も止まるなら、GAS の外向き通信そのものが止まっている。');
  Logger.log('  ※ これは見るだけです。設定は変わっていません。');
}
