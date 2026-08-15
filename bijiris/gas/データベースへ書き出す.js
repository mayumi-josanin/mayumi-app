// スプレッドシートの中身を、データベースへ取り込める形（JSON）で書き出す道具。
//
//   測定履歴を書き出す()  … ビジリスの測定履歴を JSON にしてドライブへ置く
//
// **読むだけ。**元のスプレッドシートは一切変えない。
//
// 書き出すときに、お名前から会員IDを引いて一緒に入れる。
// データベース側は会員IDで結ぶ設計なので、ここで済ませておくと
// あとから名寄せをやり直さずに済む（「山谷→山谷未央」のような改名に強くなる）。
//
// 使いかた:
//   1. この関数を実行する
//   2. ログに出たURLからJSONをダウンロードする
//   3. サーバー側で  python manage.py 計測記録を取り込む 測定履歴.json --下見
//   4. 件数が合っていれば --下見 を外して実行

var 書出_まゆみID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';
var 書出_ビジリスID = '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E';

function 書出_空か_(x) { return x === '' || x === null || x === undefined; }

// 突き合わせ用にお名前を整える。空白の有無や全角半角のゆれを吸収する。
function 書出_名前をそろえる_(値) {
  var s = String(値 == null ? '' : 値).replace(/[\s　]+/g, '');
  try { s = s.normalize('NFKC'); } catch (e) { }
  return s;
}

function 書出_日付_(値) {
  if (書出_空か_(値)) return null;
  var d = 値 instanceof Date ? 値 : new Date(値);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function 書出_日時_(値) {
  if (書出_空か_(値)) return null;
  var d = 値 instanceof Date ? 値 : new Date(値);
  if (isNaN(d.getTime())) return null;
  return Utilities.formatDate(d, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function 書出_数_(値) {
  if (書出_空か_(値)) return null;
  var n = Number(値);
  return isNaN(n) ? null : n;
}

// お名前 → 会員ID の対応表を作る。同姓同名がいれば、その名前は使わない。
function 書出_会員の対応表_() {
  var sh = SpreadsheetApp.openById(書出_まゆみID).getSheetByName('会員データ');
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var 位 = {};
  v[0].forEach(function (h, i) { 位[String(h)] = i; });

  var 表 = {};
  var 重複 = {};
  for (var r = 1; r < v.length; r += 1) {
    var 名 = 書出_名前をそろえる_(v[r][位['氏名']]);
    var id = String(v[r][位['ID']] || '').trim();
    if (!名 || !id) continue;
    if (表[名]) { 重複[名] = true; continue; }
    表[名] = id;
  }
  // 同姓同名は、どちらか分からないので結ばない（間違った人に紐づけない）
  Object.keys(重複).forEach(function (k) { delete 表[k]; });
  return { 表: 表, 重複: Object.keys(重複) };
}

function 測定履歴を書き出す() {
  var 会員 = 書出_会員の対応表_();
  var sh = SpreadsheetApp.openById(書出_ビジリスID).getSheetByName('測定履歴');
  if (!sh) { Logger.log('測定履歴のシートが見つかりません'); return; }

  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var 位 = {};
  v[0].forEach(function (h, i) { 位[String(h)] = i; });

  var 記録 = [];
  var 結べた = 0;
  var 結べない = {};

  for (var r = 1; r < v.length; r += 1) {
    var row = v[r];
    if (!row.some(function (x) { return !書出_空か_(x); })) continue;

    var 顧客名 = String(row[位['顧客名']] || '').trim();
    var かぎ = 書出_名前をそろえる_(顧客名);
    var memberId = 会員.表[かぎ] || '';
    if (memberId) 結べた += 1;
    else if (顧客名) 結べない[顧客名] = (結べない[顧客名] || 0) + 1;

    記録.push({
      measurementId: String(row[位['測定ID']] || '').trim(),
      memberId: memberId,
      customerName: 顧客名,
      memberNumber: String(row[位['会員番号']] || '').trim(),
      measuredOn: 書出_日付_(row[位['測定日']]),
      waist: 書出_数_(row[位['ウエスト(cm)']]),
      hip: 書出_数_(row[位['ヒップ(cm)']]),
      thighRight: 書出_数_(row[位['太もも右(cm)']]),
      thighLeft: 書出_数_(row[位['太もも左(cm)']]),
      whr: 書出_数_(row[位['WHR']]),
      staffMemo: String(row[位['スタッフメモ']] || ''),
      createdAt: 書出_日時_(row[位['作成日時']]),
      updatedAt: 書出_日時_(row[位['更新日時']]),
    });
  }

  var 本文 = JSON.stringify({
    exportedAt: 書出_日時_(new Date()),
    source: '測定履歴',
    measurements: 記録,
  }, null, 2);

  var 名前 = '測定履歴_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 本文, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  ファイルID: ' + file.getId());
  Logger.log('');
  Logger.log('  件数: ' + 記録.length);
  Logger.log('  会員IDに結べた: ' + 結べた);
  var 鍵 = Object.keys(結べない);
  Logger.log('  結べなかった: ' + (記録.length - 結べた) +
    (鍵.length ? '（' + 鍵.join('・') + '）' : ''));
  if (会員.重複.length) {
    Logger.log('  ※ 同姓同名のため結ばなかったお名前: ' + 会員.重複.join('・'));
  }
  Logger.log('');
  Logger.log('  中身の見本（1件目）:');
  Logger.log('  ' + JSON.stringify(記録[0]));
  Logger.log('');
  Logger.log('  次の手順:');
  Logger.log('   1. 上のURLからJSONをダウンロード');
  Logger.log('   2. python manage.py 計測記録を取り込む 測定履歴.json --下見');
  Logger.log('   3. 件数が合えば --下見 を外して実行');
  return file.getId();
}
