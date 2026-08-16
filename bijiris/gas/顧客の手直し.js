// ビジリスの顧客管理を手直しする道具。
//
//   手直しの下見()   … 何を消すか・何を外すか見るだけ（何も書かない）
//   手直しを行う()   … 控えを取ってから実行する
//
// なぜ要るのか:
//  ・テスト用の記録が残っていた（「テスト確認」）
//  ・会員データから消した方が、ビジリス側にだけ残っていた（「釼持裕基」）
//  ・前多洋子さんの「別名」に、別の方のお名前（千葉萌恵）が混ざっていた
//
// 別名はお名前の照合に使われる。別の方のお名前が入っていると、
// その方がログインしたときに前多さんの記録が出てしまう。放置できない。

var 手直し_消す名前 = ['テスト確認', '釼持裕基'];

// お名前として意味をなさない顧客も消す（「????」など）。
// 作られないようにサーバー側でも弾いたが、すでにできている分はここで片付ける。

// 前多洋子さんの別名から外すもの。
// お名前として意味をなさないもの（記号だけ・文字化け）も一緒に外す。
var 手直し_別名を外す = {
  '前多洋子': ['千葉萌恵', 'テスト']
};


function 手直しを行う() {
  手直し_控えを取る_();

  var profiles = getCustomerProfiles_() || {};
  var 消した = 0, 直した = 0;

  Object.keys(profiles).forEach(function (key) {
    var r = profiles[key] || {};
    var 名 = String(r.name || '');

    if (手直し_消す名前.indexOf(名) >= 0 || 手直し_記号だけか_(名)) {
      delete profiles[key];
      消した += 1;
      Logger.log('  消しました: ' + 名);
      return;
    }

    var 別名 = Array.isArray(r.aliases) ? r.aliases : [];
    var 残す = 別名.filter(function (a) {
      if (手直し_記号だけか_(a)) return false;
      var 一覧 = 手直し_別名を外す[名] || [];
      return 一覧.indexOf(String(a).trim()) < 0;
    });
    if (残す.length !== 別名.length) {
      r.aliases = 残す;
      r.updatedAt = new Date().toISOString();
      profiles[key] = r;
      直した += 1;
      Logger.log('  別名を直しました: ' + 名 + ' → ' + (残す.length ? 残す.join('・') : '別名なし'));
    }
  });

  saveCustomerProfiles_(profiles);
  appendAuditLog_('customer.profile.cleanup', { removed: 消した, aliasFixed: 直した });

  Logger.log('');
  Logger.log('■ 消した顧客: ' + 消した + '件 / 別名を直した方: ' + 直した + '名');
  Logger.log('  管理アプリの顧客管理でご確認ください。');
}

function 手直し_記号だけか_(値) {
  var s = String(値 || '').trim();
  if (!s) return true;
  // ひらがな・カタカナ・漢字・英数字を1文字も含まないものは、お名前ではない。
  return !/[ぁ-んァ-ヶ一-龥a-zA-Z0-9]/.test(s);
}

function 手直しの下見() {
  var profiles = getCustomerProfiles_() || {};
  var 消す = [], 外す = [];

  Object.keys(profiles).forEach(function (key) {
    var r = profiles[key] || {};
    var 名 = String(r.name || '');

    if (手直し_消す名前.indexOf(名) >= 0 || 手直し_記号だけか_(名)) {
      消す.push({ key: key, 名: 名 || '（記号だけ）', 会員番号: String(r.memberNumber || '') });
      return;
    }

    var 別名 = Array.isArray(r.aliases) ? r.aliases : [];
    var 外す対象 = 別名.filter(function (a) {
      if (手直し_記号だけか_(a)) return true;
      var 一覧 = 手直し_別名を外す[名] || [];
      return 一覧.indexOf(String(a).trim()) >= 0;
    });
    if (外す対象.length) {
      外す.push({ 名: 名, 外す: 外す対象, 残る: 別名.filter(function (a) { return 外す対象.indexOf(a) < 0; }) });
    }
  });

  Logger.log('■ 消す顧客: ' + 消す.length + '件');
  消す.forEach(function (x) { Logger.log('    ' + x.名 + '（' + x.会員番号 + '）'); });
  Logger.log('');
  Logger.log('■ 別名から外す: ' + 外す.length + '名');
  外す.forEach(function (x) {
    Logger.log('    ' + x.名);
    x.外す.forEach(function (a) {
      // 文字化けの正体が分かるよう、文字コードも出す。
      var 符号 = String(a).split('').map(function (c) { return c.charCodeAt(0); }).join(',');
      Logger.log('      外す: 「' + a + '」 (文字コード: ' + 符号 + ')');
    });
    Logger.log('      残る別名: ' + (x.残る.length ? x.残る.join('・') : 'なし'));
  });
  Logger.log('');
  Logger.log('  よければ 手直しを行う() を実行してください（先に控えを取ります）。');
}

function 手直し_控えを取る_() {
  var 生 = PropertiesService.getScriptProperties().getProperty(CUSTOMER_PROFILES_PROPERTY_KEY) || '{}';
  var 名前 = 'ビジリス顧客管理_控え_' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 生, MimeType.PLAIN_TEXT);
  Logger.log('■ 控えを取りました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('');
  return file;
}

