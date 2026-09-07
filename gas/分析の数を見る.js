// 分析画面の数字を、サーバーと突き合わせるために書き出す道具。**読むだけ。**
//
//   分析の数を見る()
//
// GAS の getAnalyticsData() をそのまま呼び、月ごとの合計だけを出す。
// **お客様の情報は含まない。**金額と件数だけ。

function 分析の数を見る() {
  var r = getAnalyticsData();
  if (!r || r.status !== 'ok') {
    Logger.log('分析を取れませんでした: ' + JSON.stringify(r).slice(0, 200));
    return;
  }
  Logger.log('■ 本番の分析（GAS）');
  Logger.log('');
  Logger.log('  月: ' + (r.months || []).length + '件  ' + (r.months || []).slice(0, 6).join(', '));
  Logger.log('  商品: ' + (r.products || []).length + '種類');
  Logger.log('  メニュー種別: ' + (r.menuTypes || []).join(', '));
  Logger.log('');

  (r.months || []).forEach(function (m) {
    var x = (r.matrix || {})[m];
    if (!x) return;
    Logger.log('  ' + m);
    Logger.log('    メニュー 売上' + Math.round(x.menuRevenueTotal || 0)
      + ' 原価' + Math.round(x.menuCostTotal || 0) + ' 粗利' + Math.round(x.menuProfitTotal || 0));
    Logger.log('    商品     売上' + Math.round(x.sales || 0)
      + ' 原価' + Math.round(x.cost || 0) + ' 粗利' + Math.round(x.profit || 0));
    Logger.log('    合計     売上' + Math.round(x.combinedRevenue || 0)
      + ' 粗利' + Math.round(x.combinedProfit || 0));
    var 種 = [];
    Object.keys(x.menuRevenue || {}).forEach(function (k) {
      if (x.menuRevenue[k]) 種.push(k + ':' + Math.round(x.menuRevenue[k]));
    });
    Logger.log('    種別     ' + (種.join(' / ') || 'なし'));
  });
  Logger.log('');
  Logger.log('  ※ 読むだけです。');
}
