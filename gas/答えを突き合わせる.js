// GAS の答えと、サーバーの答えが同じかを確かめる道具。
//
//   GASの答えを数える()   … GAS が実際に返す件数と項目を出す
//
// **読むだけ。**
//
// なぜ要るのか:
//   段階Bでサーバーへ切り替えるとき、**返す形が1項目でも違うと画面が崩れる。**
//   件数が違えば、お客様には「記事が消えた」ように見える。
//   作ったら必ず、**GASの実物と数を突き合わせる。**

function GASの答えを数える() {
  var 対象 = [
    { 名: 'getNews',        呼ぶ: function () { return getBlogNews(); },      中身: 'news' },
    { 名: 'getMenus',       呼ぶ: function () { return getMenus(); },         中身: 'menus' },
    { 名: 'getCalendar',    呼ぶ: function () { return getCalendarEvents(); }, 中身: 'events' },
    { 名: 'getSupportFaq',  呼ぶ: function () { return getSupportFaq(); },    中身: 'faqs' },
    { 名: 'getPushNotices', 呼ぶ: function () { return getPushNotices(); },   中身: 'notices' },
    { 名: 'getProducts',    呼ぶ: function () { return getProducts(); },      中身: 'products' }
  ];

  対象.forEach(function (x) {
    Logger.log('■ ' + x.名);
    var r;
    try {
      r = x.呼ぶ();
    } catch (e) {
      Logger.log('  **しくじりました: ' + e + '**');
      Logger.log('');
      return;
    }
    Logger.log('  status : ' + r.status);
    var 一覧 = r[x.中身] || [];
    Logger.log('  ' + x.中身 + ' : ' + 一覧.length + '件');

    // 一緒に返している他のもの
    Object.keys(r).forEach(function (k) {
      if (k === 'status' || k === x.中身) return;
      var v = r[k];
      Logger.log('  ' + k + ' : ' + (v && v.length !== undefined ? v.length + '件' : String(v)));
    });

    if (一覧.length) {
      var 鍵 = Object.keys(一覧[0]);
      Logger.log('  項目 ' + 鍵.length + '個: ' + 鍵.join(' / '));
      // 1件目の中身までは出さない。5種ぶんだと画面が流れて読めない。
      // **件数と項目名が合っていれば、まず十分。**
    }
    Logger.log('');
  });

  Logger.log('  ※ 読むだけです。');
  Logger.log('  ※ この件数・項目が、サーバー側と一致していること。');
}
