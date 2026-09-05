// 「お困りのとき」とチャットボットが、実際に使われているかを数える道具。
//
//   使い方案内の使われ方を数える()   … 件数と時期だけを出す
//
// **読むだけ。**シートは一切変えない。
//
// なぜ要るのか:
//   FAQ を消してよいかを決めるため。**使われていないと決めつけない。**
//   「別の表に記録が無いことを、その行が空である根拠にする」で
//   一度しくじっている（2026-08-16）。数えてから判断する。
//
// お客様が入力した文言は**出さない。**件数と日付だけを見る。
// 相談の中身は、その方の困りごとそのもの。数えるのに中身は要らない。

function 使い方案内の使われ方を数える() {
  var ss = getOrCreateSpreadsheet();

  Logger.log('■ 使い方案内（FAQ・チャットボット）の使われ方');
  Logger.log('');

  // ── FAQ そのもの ──
  var faq = ss.getSheetByName(SHEETS.APP_SUPPORT_FAQ);
  if (!faq) {
    Logger.log('  FAQ の表: **ありません**');
  } else {
    var 最終 = faq.getLastRow();
    var 件 = Math.max(0, 最終 - 1);
    var 公開 = 0;
    if (件 > 0) {
      faq.getRange(2, 1, 件, 1).getDisplayValues().forEach(function (r) {
        if (String(r[0] || '公開').trim() !== '非公開') 公開++;
      });
    }
    Logger.log('  FAQ の表      : ' + 件 + '件（公開 ' + 公開 + '件）');
  }

  // ── お客様が実際に聞いた記録 ──
  var log = ss.getSheetByName(SHEETS.SUPPORT_CHAT_LOG);
  if (!log) {
    Logger.log('  相談の記録    : **表がありません**（一度も使われていない）');
  } else {
    var 行 = Math.max(0, log.getLastRow() - 1);
    Logger.log('  相談の記録    : ' + 行 + '件');
    if (行 > 0) {
      var 日 = log.getRange(2, 1, 行, 1).getDisplayValues()
        .map(function (r) { return String(r[0] || '').slice(0, 10); })
        .filter(Boolean).sort();
      Logger.log('    はじめ      : ' + 日[0]);
      Logger.log('    いちばん新しい: ' + 日[日.length - 1]);

      // 直近90日に何件あったか
      var 境 = new Date();
      境.setDate(境.getDate() - 90);
      var 境の字 = Utilities.formatDate(境, 'Asia/Tokyo', 'yyyy-MM-dd');
      var 最近 = 日.filter(function (d) { return d >= 境の字; }).length;
      Logger.log('    直近90日     : ' + 最近 + '件');
    }
  }

  Logger.log('');
  Logger.log('  ※ 読むだけです。お客様が入力された文言は出していません。');
}
