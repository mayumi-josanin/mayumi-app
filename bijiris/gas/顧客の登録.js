// まゆみ側の会員データから、ビジリスをお使いの方を顧客管理へ登録する道具。
//
//   顧客登録の下見()   … 誰を登録するか見るだけ（何も書かない）
//   顧客を登録する()   … 実際に登録する
//
// ビジリスの顧客管理は、これまで回答が届いたときに自動でできていた。
// その回答が1件も無いため、顧客管理が空のままだった。
// まゆみ側には「ビジリス」の印があるので、そちらを正として作る。
//
// 会員IDではなく**お名前**で結ぶ。ビジリス側の記録（測定履歴・回数券分析）が
// お名前で入っているため、ここだけ会員IDにすると突き合わなくなる。
//
// すでにある方は触らない（お名前の付け替えや管理者の設定を上書きしないため）。

function 顧客登録の下見() {
  var r = 顧客登録_対象を集める_();
  Logger.log('■ ビジリスの印がある方: ' + r.対象.length + '名');
  Logger.log('');
  Logger.log('  新しく登録する: ' + r.新規.length + '名');
  r.新規.forEach(function (x) {
    Logger.log('    ' + x.name + (x.nameKana ? '（' + x.nameKana + '）' : '') +
      (x.memberNumber ? ' / ' + x.memberNumber : ''));
  });
  Logger.log('');
  Logger.log('  すでに顧客管理にある: ' + r.既存.length + '名');
  r.既存.forEach(function (n) { Logger.log('    ' + n); });
  Logger.log('');
  if (r.印は無いが記録あり.length) {
    Logger.log('  ※ ビジリスの印は無いが、測定や回数券の記録があるお名前: ' + r.印は無いが記録あり.length + '名');
    r.印は無いが記録あり.forEach(function (n) { Logger.log('    ' + n); });
    Logger.log('    （まゆみ側の会員に見つからない方です。受付でご確認ください）');
    Logger.log('');
  }
  Logger.log('  よければ 顧客を登録する() を実行してください。');
}

function 顧客を登録する() {
  var r = 顧客登録_対象を集める_();
  if (!r.新規.length) {
    Logger.log('新しく登録する方はいません。');
    return;
  }

  var profiles = getCustomerProfiles_();
  var 作った = 0;
  r.新規.forEach(function (x) {
    // 既存の作りに合わせて記録を作る。ここで直接書かず、正規化を通す。
    var record = normalizeCustomerProfileRecord_({
      name: x.name,
      nameKana: x.nameKana,
      memberNumber: x.memberNumber,
      aliases: [],
      clientIds: [],
      adminManaged: true,
    }, x.name);
    if (!record) return;
    record.updatedAt = new Date().toISOString();
    saveCustomerProfileRecord_(profiles, null, record, {});
    作った += 1;
    Logger.log('  登録: ' + x.name);
  });
  saveCustomerProfiles_(profiles);

  appendAuditLog_('customer.profile.bulkCreate', { created: 作った, source: 'まゆみ会員データのビジリス印' });

  Logger.log('');
  Logger.log('■ ' + 作った + '名を顧客管理に登録しました');
  Logger.log('  管理アプリの顧客管理でご確認ください。');
  Logger.log('  スタンプの個数は、顧客ごとの ＋1 で足してください。');
}



var 顧客登録_まゆみID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';

function 顧客登録_名前をそろえる_(値) {
  return String(値 == null ? '' : 値).replace(/[\s　]+/g, '').trim();
}

function 顧客登録_対象を集める_() {
  var sh = SpreadsheetApp.openById(顧客登録_まゆみID).getSheetByName('会員データ');
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var 位 = {};
  v[0].forEach(function (h, i) { 位[String(h)] = i; });

  var profiles = getCustomerProfiles_();
  var ある = {};
  Object.keys(profiles || {}).forEach(function (key) {
    var n = 顧客登録_名前をそろえる_(profiles[key] && profiles[key].name);
    if (n) ある[n] = true;
  });

  var 対象 = [], 新規 = [], 既存 = [];
  var 会員名 = {};
  for (var r = 1; r < v.length; r += 1) {
    var 名 = 顧客登録_名前をそろえる_(v[r][位['氏名']]);
    if (!名) continue;
    会員名[名] = true;
    // 管理者はお客様ではないので入れない。
    if (String(v[r][位['権限']] || '').trim() === '管理者') continue;
    if (!String(v[r][位['ビジリス']] || '').trim()) continue;

    var x = {
      name: String(v[r][位['氏名']] || '').trim(),
      nameKana: String(v[r][位['フリガナ']] || v[r][位['ふりがな']] || '').trim(),
      memberNumber: String(v[r][位['会員番号']] || v[r][位['ID']] || '').trim(),
    };
    対象.push(x);
    if (ある[名]) 既存.push(x.name);
    else 新規.push(x);
  }

  // ビジリスの記録があるのに、まゆみ側の会員に見つからないお名前。
  var 記録 = ビジリス_記録のあるお名前_();
  var 印は無いが記録あり = Object.keys(記録).filter(function (n) { return !会員名[n]; });

  return { 対象: 対象, 新規: 新規, 既存: 既存, 印は無いが記録あり: 印は無いが記録あり };
}
