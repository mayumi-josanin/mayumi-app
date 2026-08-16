// まゆみ会員データの「ビジリス」列と、実態を突き合わせる道具。
//
//   ビジリス印の下見()   … 読むだけ。何も書かない。
//
// なぜ要るのか:
//   入口のアプリ一覧は、まゆみ会員データの「ビジリス」列で出し分けている。
//   この列が実態とずれていると、
//     ・記録があるのにビジリスアプリが出てこない方
//     ・使っていないのにビジリスアプリが出ている方
//   が生まれる。どちらもお客様には理由が分からない。
//
// 3つを突き合わせる。
//   ① まゆみ会員データの「ビジリス」列（＝入口の出し分けの根拠）
//   ② ビジリスの顧客管理（＝受付が見ている一覧）
//   ③ ビジリスの測定履歴（＝実際に施術を受けられた証拠）

var 印下見_まゆみID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';

function 印下見_そろえる_(値) {
  var s = String(値 == null ? '' : 値).replace(/[\s　]+/g, '');
  try { s = s.normalize('NFKC'); } catch (e) { }
  return s;
}

function ビジリス印の下見() {
  // ① まゆみ会員データ
  var sh = SpreadsheetApp.openById(印下見_まゆみID).getSheetByName('会員データ');
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var 位 = {};
  v[0].forEach(function (h, i) { 位[String(h)] = i; });

  var 印あり = {}, 会員の名 = {};
  for (var r = 1; r < v.length; r += 1) {
    var 名 = String(v[r][位['氏名']] || '').trim();
    if (!名) continue;
    var 鍵 = 印下見_そろえる_(名);
    会員の名[鍵] = 名;
    if (String(v[r][位['ビジリス']] || '').trim()) 印あり[鍵] = 名;
  }

  // ② ビジリスの顧客管理
  var profiles = getCustomerProfiles_() || {};
  var 顧客 = {};
  Object.keys(profiles).forEach(function (k) {
    var 名 = String((profiles[k] || {}).name || '').trim();
    if (名) 顧客[印下見_そろえる_(名)] = 名;
  });

  // ③ ビジリスの測定履歴（実際に施術を受けられた証拠）
  var 記録 = {};
  var ms = getSpreadsheet_().getSheetByName(MEASUREMENTS_SHEET_NAME);
  if (ms && ms.getLastRow() > 1) {
    var mv = ms.getRange(1, 1, ms.getLastRow(), ms.getLastColumn()).getValues();
    var 名列 = mv[0].indexOf('顧客名');
    for (var i = 1; i < mv.length; i += 1) {
      var mn = String(mv[i][名列] || '').trim();
      if (mn) 記録[印下見_そろえる_(mn)] = mn;
    }
  }

  var 印 = Object.keys(印あり), 客 = Object.keys(顧客), 録 = Object.keys(記録);

  Logger.log('■ いまの数');
  Logger.log('    まゆみ「ビジリス」列に印: ' + 印.length + '名');
  Logger.log('    ビジリス顧客管理:         ' + 客.length + '名');
  Logger.log('    ビジリス測定履歴に記録:   ' + 録.length + '名');
  Logger.log('');

  var 印だけ = 印.filter(function (k) { return 客.indexOf(k) < 0; });
  var 客だけ = 客.filter(function (k) { return 印.indexOf(k) < 0; });
  var 記録あるが印なし = 録.filter(function (k) { return 印.indexOf(k) < 0; });
  var 記録あるが顧客なし = 録.filter(function (k) { return 客.indexOf(k) < 0; });

  Logger.log('■ 印はあるが、顧客管理にいない: ' + 印だけ.length + '名');
  Logger.log('    （アプリ一覧にビジリスが出るが、記録が何も無い方）');
  印だけ.forEach(function (k) {
    Logger.log('    ' + 印あり[k] + (記録[k] ? '（測定履歴には記録あり）' : '（測定履歴にも無し）'));
  });
  Logger.log('');

  Logger.log('■ 顧客管理にいるが、印が無い: ' + 客だけ.length + '名');
  Logger.log('    （受付では扱っているのに、アプリ一覧にビジリスが出ない方）');
  客だけ.forEach(function (k) {
    Logger.log('    ' + 顧客[k] + (会員の名[k] ? '' : ' ← **まゆみの会員データに見つかりません**'));
  });
  Logger.log('');

  Logger.log('■ 測定の記録があるのに、印が無い: ' + 記録あるが印なし.length + '名');
  記録あるが印なし.forEach(function (k) { Logger.log('    ' + 記録[k]); });
  Logger.log('');

  Logger.log('■ 測定の記録があるのに、顧客管理にいない: ' + 記録あるが顧客なし.length + '名');
  記録あるが顧客なし.forEach(function (k) { Logger.log('    ' + 記録[k]); });
  Logger.log('');

  if (!印だけ.length && !客だけ.length && !記録あるが印なし.length && !記録あるが顧客なし.length) {
    Logger.log('■ 3つとも一致しています。直すところはありません。');
  } else {
    Logger.log('■ 直し方の目安');
    Logger.log('    「顧客管理にいるが印が無い」… まゆみ会員データの「ビジリス」列に');
    Logger.log('      「登録済み」を入れる。入口にビジリスが出るようになります。');
    Logger.log('    「印はあるが顧客管理にいない」… 本当にご利用かを受付でご確認ください。');
    Logger.log('      ご利用でなければ印を外す。出しっぱなしだと、開いても何も無い画面になります。');
  }
}
