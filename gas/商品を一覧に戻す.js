// 「一覧から削除」してしまった商品を、お客様のショップに戻す道具。
//
//   商品の一覧状態を見る()   … いまどれが外れているかを見るだけ
//   商品を一覧に戻す()       … **書き込みます。**控えを取ってから実行
//
// ─────────────────────────────────────────────────────────
// なぜ要るのか（2026-09-05）
//
// 管理画面の「お知らせ管理」で商品の「一覧から削除」を押すと、
// **お客様のショップからもその商品が消える。**
//
//   getProducts（3420行）
//     if (isNoticeListingDeleted_(row, deletedAtCol)) continue;   ← ここで外れる
//
// 画面には「お知らせ一覧から完全に消えます」としか書かれていないので、
// ショップから消えるとは思わずに押せてしまう。実際、4件が消えていた。
//
// **一度外すと、管理画面の一覧からもその商品が消えるので、画面からは戻せない。**
// だからこの道具が要る。
//
// 戻し方は「お知らせ一覧削除日時」の列を空にするだけ。

var 戻す_対象 = [
  'よもぎ茶（30パック）',
  'よもぎ入浴剤（10パック）',
  'よもぎ入浴剤（2パック）',
  '石鹸（あこがれのきよら）'
];

function 商品を一覧に戻す_下ごしらえ_() {
  var ss = getOrCreateSpreadsheet();
  var sheet = ss.getSheetByName(SHEETS.PRODUCTS);
  if (!sheet) return null;
  var 列 = ensureNoticeDeletedAtColumn_(sheet);
  var 最終 = sheet.getLastRow();
  if (最終 < 2 || !列) return null;
  return {
    ss: ss, sheet: sheet, 列: 列,
    値: sheet.getRange(2, 1, 最終 - 1, sheet.getLastColumn()).getDisplayValues()
  };
}

/** いまどの商品が一覧から外れているかを見る。**読むだけ。** */
function 商品の一覧状態を見る() {
  var 下 = 商品を一覧に戻す_下ごしらえ_();
  if (!下) { Logger.log('商品の表が読めません'); return; }

  Logger.log('■ 商品の「お知らせ一覧削除日時」');
  Logger.log('');
  下.値.forEach(function (row, i) {
    var 名 = String(row[1] || '').trim();
    if (!名) return;
    var 外 = String(row[下.列 - 1] || '').trim();
    var 公開 = String(row[5] || '').trim() || '公開';
    Logger.log('  行' + (i + 2) + '  ' + 名);
    Logger.log('       公開設定=' + 公開 + '  一覧から外した=' + (外 || '—')
      + (外 ? '   ← **ショップに出ていません**' : ''));
  });
  Logger.log('');
  Logger.log('  ※ 読むだけです。何も変えていません。');
}

/** 上の 戻す_対象 に挙げた商品を、一覧に戻す。**書き込みます。** */
function 商品を一覧に戻す() {
  var 下 = 商品を一覧に戻す_下ごしらえ_();
  if (!下) { Logger.log('商品の表が読めません'); return; }

  // **触る前に控えを取る。**スプレッドシートごと複製する。
  var 控え = DriveApp.getFileById(下.ss.getId())
    .makeCopy('控え_商品を一覧に戻す前_' +
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm'));
  Logger.log('■ 控えを取りました: ' + 控え.getName());
  Logger.log('   ' + 控え.getUrl());
  Logger.log('');

  var 戻した = [];
  下.値.forEach(function (row, i) {
    var 名 = String(row[1] || '').trim();
    if (戻す_対象.indexOf(名) === -1) return;
    var 外 = String(row[下.列 - 1] || '').trim();
    if (!外) return;   // すでに一覧に出ている
    下.sheet.getRange(i + 2, 下.列).setValue('');
    戻した.push(名);
  });

  Logger.log('■ 一覧に戻した商品: ' + 戻した.length + '件');
  戻した.forEach(function (n) { Logger.log('  ・' + n); });
  if (!戻した.length) {
    Logger.log('  （戻す対象がありませんでした。すでに出ているか、名前が違います）');
  }
  Logger.log('');
  Logger.log('  お客様のアプリは、開き直すと反映されます。');
}
