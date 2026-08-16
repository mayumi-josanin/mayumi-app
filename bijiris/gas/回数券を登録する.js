// 顧客管理の回数券カードを、まとめて登録する道具。
//
//   回数券登録の下見()   … 誰に何を入れるか見るだけ（何も書かない）
//   回数券を登録する()   … 控えを取ってから実行する
//
// すでに登録されている方は触らない。受付で個別に直した内容を
// 上書きしてしまわないため。

var 回数券登録_内容 = { plan: '10回券', sheetNumber: 2, round: 1 };

var 回数券登録_対象 = [
  '國分彩子', '宮村綾子', '小澤美奈子', '小瀬村真理子', '山谷未央',
  '岩谷梨奈', '廣田沙織', '木原真理', '藤田茉衣'
];

function 回数券を登録する() {
  手直し_控えを取る_();

  var r = 回数券登録_調べる_();
  var c = 回数券登録_内容;
  var 入れた = 0;

  r.入れる.forEach(function (x) {
    var rec = r.profiles[x.key];
    rec.activeTicketCard = normalizeActiveTicketCard_(c);
    // 受付で設定したものと同じ扱いにする。回答が届けばそちらで上書きされる。
    rec.activeTicketCardSource = 'admin';
    rec.updatedAt = new Date().toISOString();
    r.profiles[x.key] = rec;
    入れた += 1;
    Logger.log('  入れました: ' + x.名 + ' → ' + c.plan + '・' + c.sheetNumber + '枚目・' + c.round + '回目');
  });

  saveCustomerProfiles_(r.profiles);
  appendAuditLog_('customer.ticketCard.bulkSet', { updated: 入れた, card: c });

  Logger.log('');
  Logger.log('■ ' + 入れた + '名に登録しました');
}

function 回数券登録_そろえる_(値) {
  return String(値 == null ? '' : 値).replace(/[\s　]+/g, '').trim();
}

function 回数券登録_調べる_() {
  var profiles = getCustomerProfiles_() || {};
  var 入れる = [], すでに = [], 見つからない = [];

  回数券登録_対象.forEach(function (名) {
    var key = null;
    Object.keys(profiles).forEach(function (k) {
      if (回数券登録_そろえる_(profiles[k] && profiles[k].name) === 回数券登録_そろえる_(名)) key = k;
    });
    if (!key) { 見つからない.push(名); return; }

    var 今 = normalizeActiveTicketCard_(profiles[key].activeTicketCard);
    if (今) {
      すでに.push(profiles[key].name + '（' + 今.plan + '・' + 今.sheetNumber + '枚目・' + 今.round + '回目）');
      return;
    }
    入れる.push({ key: key, 名: profiles[key].name });
  });

  return { profiles: profiles, 入れる: 入れる, すでに: すでに, 見つからない: 見つからない };
}

function 回数券登録の下見() {
  var r = 回数券登録_調べる_();
  var c = 回数券登録_内容;
  Logger.log('■ 入れる内容: ' + c.plan + '・' + c.sheetNumber + '枚目・' + c.round + '回目');
  Logger.log('');
  Logger.log('  入れる: ' + r.入れる.length + '名');
  r.入れる.forEach(function (x) { Logger.log('    ' + x.名); });
  Logger.log('');
  if (r.すでに.length) {
    Logger.log('  すでに登録あり（触りません）: ' + r.すでに.length + '名');
    r.すでに.forEach(function (x) { Logger.log('    ' + x); });
    Logger.log('');
  }
  if (r.見つからない.length) {
    Logger.log('  顧客管理に見つからない: ' + r.見つからない.join('・'));
    Logger.log('');
  }
  Logger.log('  よければ 回数券を登録する() を実行してください（先に控えを取ります）。');
}

