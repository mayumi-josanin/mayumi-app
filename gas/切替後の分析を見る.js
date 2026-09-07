// 切り替えたあとの分析が、正しく出ているかを見る道具。**読むだけ。**
//
//   切替後の分析を見る()
//
// どこから来ているか（サーバー／シート）も見分ける。

function 切替後の分析を見る() {
  var 渡している = PropertiesService.getScriptProperties().getProperty('SERVER_TABLES') || '';
  Logger.log('■ いま渡している表: ' + (渡している || 'なし'));
  Logger.log('');

  var t = new Date().getTime();
  var r = getAnalyticsData();
  var かかった = new Date().getTime() - t;

  if (!r || r.status !== 'ok') {
    Logger.log('★ 分析を取れません: ' + JSON.stringify(r).slice(0, 200));
    return;
  }

  Logger.log('  かかった時間: ' + かかった + 'ミリ秒');
  Logger.log('  月: ' + (r.months || []).length + '件');
  Logger.log('  商品: ' + (r.products || []).length + '種類');
  Logger.log('');
  Logger.log('  **登録経路**: ' + Object.keys(r.registrationRoutes || {}).length + '種類'
    + (Object.keys(r.registrationRoutes || {}).length ? '' : '   ← **空です。会員から作る分が落ちています**'));
  Logger.log('  **カテゴリ利用**: ' + Object.keys(r.categoryUsage || {}).length + '種類'
    + (Object.keys(r.categoryUsage || {}).length ? '' : '   ← **空です**'));
  Logger.log('');

  (r.months || []).slice(0, 3).forEach(function (m) {
    var x = (r.matrix || {})[m];
    if (!x) return;
    Logger.log('  ' + m + '  合計 ' + Math.round(x.combinedRevenue || 0)
      + ' / 粗利 ' + Math.round(x.combinedProfit || 0));
  });

  Logger.log('');
  Logger.log('  ※ 読むだけです。');
}
