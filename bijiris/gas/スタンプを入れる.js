// 回数券スタンプの個数を、記録に合わせて入れる道具。
//
//   個数の下見()   … 誰がいくつになるか見るだけ（何も書かない）
//   個数を入れる() … 控えを取ってから実行する
//
// 数え方の基本は施術後アンケートの提出だが、その回答は1件も残っていない
// （アプリ導入前の記録のため）。実際に使い切った証拠は回数券分析シートの
// 15件しかないので、その枚数を手当てとして入れる。
//
// アンケートが出されるようになれば、そちらが自動で加算される。二重に数え
// ないよう、手当ては「アンケートから数えた分」との合計で目標枚数になるよう
// 調整する（同じ関数を何度実行しても同じ結果になる）。

var 入れる_目標 = {
  '前多洋子': 3,
  '岩谷梨奈': 3,
  '廣田沙織': 2,
  '木原真理': 2,
  '宮村綾子': 2,
  '小瀬村真理子': 1,
  '小澤美奈子': 1,
  '藤田茉衣': 1
};

function 個数を入れる() {
  手直し_控えを取る_();

  var s = 入れる_今の状態_();
  var 直した = 0;
  s.行.forEach(function (x) {
    var r = s.profiles[x.key];
    if (Math.floor(Number(r.ticketStampAdjustment || 0)) === x.次の手当て) {
      Logger.log('  変わりません: ' + x.名 + '（' + x.目標 + '個）');
      return;
    }
    r.ticketStampAdjustment = normalizeTicketStampAdjustment_(x.次の手当て);
    r.updatedAt = new Date().toISOString();
    s.profiles[x.key] = r;
    直した += 1;
    Logger.log('  入れました: ' + x.名 + ' → ' + x.目標 + '個（手当て ' + r.ticketStampAdjustment + '）');
  });
  saveCustomerProfiles_(s.profiles);
  appendAuditLog_('customer.ticketStamp.bulkSet', { updated: 直した });

  Logger.log('');
  Logger.log('■ ' + 直した + '名のスタンプを入れました');
  Logger.log('  管理アプリの顧客管理でご確認ください。');
}

function 入れる_名前をそろえる_(値) {
  return String(値 == null ? '' : 値).replace(/[\s　]+/g, '').trim();
}

function 入れる_今の状態_() {
  var counts = getCompletedTicketCardCountsByCustomer_(getResponses_({}));
  var profiles = getCustomerProfiles_() || {};
  var 行 = [], 見つからない = [];

  Object.keys(入れる_目標).forEach(function (名) {
    var key = null;
    Object.keys(profiles).forEach(function (k) {
      if (入れる_名前をそろえる_(profiles[k] && profiles[k].name) === 入れる_名前をそろえる_(名)) key = k;
    });
    if (!key) { 見つからない.push(名); return; }

    var r = profiles[key];
    var アンケートから = Math.floor(Number(counts[r.name] || 0));
    var 目標 = 入れる_目標[名];
    var 今の手当て = Math.floor(Number(r.ticketStampAdjustment || 0));
    var 次の手当て = 目標 - アンケートから;
    行.push({
      key: key, 名: r.name, アンケートから: アンケートから, 目標: 目標,
      今の手当て: 今の手当て, 次の手当て: 次の手当て,
      今の合計: Math.max(0, アンケートから + 今の手当て)
    });
  });
  return { profiles: profiles, 行: 行, 見つからない: 見つからない };
}

function 個数の下見() {
  var s =入れる_今の状態_();
  Logger.log('■ スタンプの個数（まだ書いていません）');
  Logger.log('');
  s.行.forEach(function (x) {
    Logger.log('  ' + x.名 + ': いま ' + x.今の合計 + '個 → ' + x.目標 + '個');
    Logger.log('      アンケートから ' + x.アンケートから + ' ／ 手当て ' + x.今の手当て + ' → ' + x.次の手当て);
  });
  Logger.log('');
  Logger.log('  合計 ' + s.行.reduce(function (a, x) { return a + x.目標; }, 0) + '個 / ' + s.行.length + '名');
  if (s.見つからない.length) {
    Logger.log('  ※ 顧客管理に見つからない方: ' + s.見つからない.join('・'));
  }
  Logger.log('');
  Logger.log('  よければ 個数を入れる() を実行してください（先に控えを取ります）。');
}

