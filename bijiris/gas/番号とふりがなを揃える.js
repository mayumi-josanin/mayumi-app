// ビジリスの顧客管理を、まゆみ側の会員データに合わせる道具。
//
//   揃えるの下見()   … 何がどう変わるか見るだけ（何も書かない）
//   揃える()         … 控えを取ってから実行する
//
//  ・会員番号 … まゆみ側の会員番号に合わせる。照合しやすくするため。
//               ビジリス側は M+4桁 に整形される作りなので、
//               MYM-3928 → M3928 のように数字を合わせる形になる。
//  ・ふりがな … 空いているところだけ入れる。入っているものは触らない。
//
// まゆみ側のふりがなも空のときは、こちらで決めた読みを両方に入れる。
// 片方だけ入れると、次に突き合わせたときにまた空に見えてしまう。

var 揃える_まゆみID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';

// 受付でご確認いただいた読み。まゆみ側にも入っていない方のぶん。
var 揃える_手で決めたふりがな = {
  '小瀬村真理子': 'こせむらまりこ',
  '廣田沙織': 'ひろたさおり',
  '藤田茉衣': 'ふじたまい'
};

function 書式統一の下見() {
  var profiles = getCustomerProfiles_() || {};
  var 顧客 = [];
  Object.keys(profiles).forEach(function (k) {
    var r = profiles[k] || {};
    var 今 = String(r.memberNumber || '');
    var 次 = normalizeMemberNumber_(今);
    if (次 !== 今) 顧客.push(String(r.name || '') + '： ' + (今 || '空') + ' → ' + (次 || '空'));
  });

  var m = 書式統一_対象を集める_();

  Logger.log('■ 顧客管理で書式が変わる方: ' + 顧客.length + '名');
  顧客.forEach(function (x) { Logger.log('    ' + x); });
  Logger.log('');
  Logger.log('■ 測定履歴シートで書き換える行: ' + m.変える.length + '行');
  m.変える.slice(0, 20).forEach(function (x) {
    Logger.log('    ' + x.行 + '行目 ' + x.名 + '： ' + x.今 + ' → ' + x.次);
  });
  if (m.変える.length > 20) Logger.log('    …ほか ' + (m.変える.length - 20) + '行');
  Logger.log('');
  Logger.log('  よければ 書式を統一する() を実行してください（先に控えを取ります）。');
}

function 揃える() {
  手直し_控えを取る_();   // ビジリス顧客管理の控え（顧客の手直し.gs にある）

  var まゆみ = 揃える_まゆみを読む_();
  var profiles = getCustomerProfiles_() || {};
  var 直した = 0;
  var まゆみに書く = [];

  Object.keys(profiles).forEach(function (key) {
    var r = profiles[key] || {};
    var 名 = String(r.name || '');
    var m = まゆみ.表[揃える_名前をそろえる_(名)];
    if (!m) return;

    var 次の番号 = normalizeMemberNumber_(m.会員番号);
    var 次のかな = String(r.nameKana || '') || m.ふりがな || 揃える_手で決めたふりがな[名] || '';
    var 変えた = false;

    if (次の番号 && String(r.memberNumber || '') !== 次の番号) { r.memberNumber = 次の番号; 変えた = true; }
    if (次のかな && String(r.nameKana || '') !== 次のかな) { r.nameKana = 次のかな; 変えた = true; }

    // まゆみ側のふりがなが空なら、そちらにも入れる。片方だけだと次に見たとき空に見える。
    if (!m.ふりがな && 揃える_手で決めたふりがな[名]) {
      まゆみに書く.push({ 行: m.行, かな: 揃える_手で決めたふりがな[名], 名: 名 });
    }

    if (変えた) {
      r.updatedAt = new Date().toISOString();
      profiles[key] = r;
      直した += 1;
      Logger.log('  直しました: ' + 名 + ' / ' + r.memberNumber + ' / ' + (r.nameKana || 'ふりがななし'));
    }
  });

  saveCustomerProfiles_(profiles);

  var 列 = まゆみ.位['フリガナ'] !== undefined ? まゆみ.位['フリガナ'] : まゆみ.位['ふりがな'];
  まゆみに書く.forEach(function (x) {
    まゆみ.sheet.getRange(x.行, 列 + 1).setValue(x.かな);
    Logger.log('  まゆみ側にも入れました: ' + x.名 + ' → ' + x.かな);
  });

  appendAuditLog_('customer.profile.alignWithMayumi', { updated: 直した, kanaWrittenToMayumi: まゆみに書く.length });

  Logger.log('');
  Logger.log('■ ビジリス側を直した方: ' + 直した + '名 / まゆみ側にふりがなを入れた方: ' + まゆみに書く.length + '名');
}

function 揃える_名前をそろえる_(値) {
  return String(値 == null ? '' : 値).replace(/[\s　]+/g, '').trim();
}

// まゆみ側の会員データを、お名前で引ける形にする。
function 揃える_まゆみを読む_() {
  var sh = SpreadsheetApp.openById(揃える_まゆみID).getSheetByName('会員データ');
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var 位 = {};
  v[0].forEach(function (h, i) { 位[String(h)] = i; });
  var 表 = {};
  for (var r = 1; r < v.length; r += 1) {
    var 名 = 揃える_名前をそろえる_(v[r][位['氏名']]);
    if (!名) continue;
    表[名] = {
      行: r + 1,
      会員番号: String(v[r][位['会員番号']] || v[r][位['ID']] || '').trim(),
      ふりがな: String(v[r][位['フリガナ']] !== undefined ? v[r][位['フリガナ']] : (v[r][位['ふりがな']] || '')).trim()
    };
  }
  return { sheet: sh, 位: 位, 表: 表 };
}

function 揃えるの下見() {
  var まゆみ = 揃える_まゆみを読む_();
  var profiles = getCustomerProfiles_() || {};
  var 変わる = [], 変わらない = [], 見つからない = [];

  Object.keys(profiles).forEach(function (key) {
    var r = profiles[key] || {};
    var 名 = String(r.name || '');
    var m = まゆみ.表[揃える_名前をそろえる_(名)];
    if (!m) { 見つからない.push(名); return; }

    var 今の番号 = String(r.memberNumber || '');
    var 次の番号 = normalizeMemberNumber_(m.会員番号);
    var 今のかな = String(r.nameKana || '');
    var 次のかな = 今のかな || m.ふりがな || 揃える_手で決めたふりがな[名] || '';

    if (今の番号 === 次の番号 && 今のかな === 次のかな) { 変わらない.push(名); return; }
    変わる.push({
      名: 名,
      番号: 今の番号 + ' → ' + 次の番号 + '（まゆみ: ' + m.会員番号 + '）',
      かな: (今のかな || '空') + ' → ' + (次のかな || '空'),
      まゆみのかなが空: !m.ふりがな && Boolean(揃える_手で決めたふりがな[名]),
      行: m.行
    });
  });

  Logger.log('■ 変わる方: ' + 変わる.length + '名');
  変わる.forEach(function (x) {
    Logger.log('    ' + x.名);
    Logger.log('      会員番号: ' + x.番号);
    Logger.log('      ふりがな: ' + x.かな + (x.まゆみのかなが空 ? '（まゆみ側の' + x.行 + '行目にも入れます）' : ''));
  });
  Logger.log('');
  Logger.log('■ 変わらない方: ' + 変わらない.length + '名 … ' + 変わらない.join('・'));
  if (見つからない.length) {
    Logger.log('');
    Logger.log('■ まゆみの会員データに見つからない方: ' + 見つからない.join('・'));
  }
  Logger.log('');
  Logger.log('  よければ 揃える() を実行してください（先に控えを取ります）。');
}


// ===== 会員番号の書式を MYM-0000 に統一する =====
//
//   書式統一の下見()   … どこが変わるか見るだけ（何も書かない）
//   書式を統一する()   … 控えを取ってから実行する
//
// ビジリスは M0000、まゆみは MYM-0000 と書式が違っていた。数字は同じでも、
// 受付で並べたときに「同じ方か」を毎回考えることになるため、文字列ごと揃える。
//
// 顧客管理の側は normalizeMemberNumber_ を通せば自動で MYM- になる。
// 測定履歴シートに直接書かれている会員番号は、そのままでは古い書式が残るので
// ここで書き換える。

function 書式統一_対象を集める_() {
  var sh = getSpreadsheet_().getSheetByName(MEASUREMENTS_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return { sheet: null, 列: 0, 変える: [] };
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var 列 = v[0].indexOf('会員番号');
  if (列 < 0) return { sheet: null, 列: 0, 変える: [] };

  var 変える = [];
  for (var r = 1; r < v.length; r += 1) {
    var 今 = String(v[r][列] || '').trim();
    if (!今) continue;
    var 次 = normalizeMemberNumber_(今);
    if (次 && 次 !== 今) 変える.push({ 行: r + 1, 今: 今, 次: 次, 名: String(v[r][v[0].indexOf('顧客名')] || '') });
  }
  return { sheet: sh, 列: 列, 変える: 変える };
}


function 書式を統一する() {
  手直し_控えを取る_();

  // 顧客管理は、読み直して書き戻すだけで normalizeMemberNumber_ を通る。
  var profiles = getCustomerProfiles_() || {};
  var 直した = 0;
  Object.keys(profiles).forEach(function (k) {
    var r = profiles[k] || {};
    var 次 = normalizeMemberNumber_(r.memberNumber);
    if (次 !== String(r.memberNumber || '')) {
      r.memberNumber = 次;
      r.updatedAt = new Date().toISOString();
      profiles[k] = r;
      直した += 1;
      Logger.log('  顧客管理: ' + r.name + ' → ' + (次 || '番号なし'));
    }
  });
  saveCustomerProfiles_(profiles);

  // 測定履歴は控えを取ってから書き換える。
  var m = 書式統一_対象を集める_();
  if (m.sheet && m.変える.length) {
    片付け_控えを取る_相当_(m.sheet);
    m.変える.forEach(function (x) {
      m.sheet.getRange(x.行, m.列 + 1).setValue(x.次);
    });
    Logger.log('  測定履歴: ' + m.変える.length + '行を書き換えました');
  }

  appendAuditLog_('customer.memberNumber.format', { profiles: 直した, measurements: m.変える.length });
  Logger.log('');
  Logger.log('■ 顧客管理 ' + 直した + '名 / 測定履歴 ' + m.変える.length + '行 を MYM- 書式に揃えました');
}

// 測定履歴を触る前の控え。スプレッドシートごと複製する。
function 片付け_控えを取る_相当_(sheet) {
  var book = sheet.getParent();
  var 名前 = '【控え】' + book.getName() + ' ' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + ' 会員番号の書式統一前';
  var 控え = DriveApp.getFileById(book.getId()).makeCopy(名前);
  Logger.log('■ 測定履歴の控え: ' + 名前);
  Logger.log('  ' + 控え.getUrl());
}
