// まだサーバーへ移していない表が、本番でどうなっているかを数える道具。
//
//   移していない表の下見()   … 件数と埋まりぐあいだけを出す
//
// **読むだけ。**シートは一切変えない。お名前などの中身は出さない。
//
// なぜ要るのか:
//   会員と注文はまだシートが正。サーバー側は8月の写しのまま止まっている。
//   **どれだけ離れているかを知らないと、切り替えの当日に驚くことになる。**

function 移していない表の下見() {
  var ss = getOrCreateSpreadsheet();

  Logger.log('■ まだ移していない表（本番のシート）');
  Logger.log('');

  // ── 会員 ──
  var u = getOrCreateUsersSheet_(ss);
  var 最終 = u.getLastRow();
  var 件 = Math.max(0, 最終 - 1);
  var 値 = 件 ? u.getRange(2, 1, 件, USER_HEADERS.length).getDisplayValues() : [];
  var 届け先 = 0, パス = 0, 削除 = 0, ID無し = 0, 名無し = 0, ビジリス = 0;
  var 使ったID = {}, 重複 = 0;
  値.forEach(function (row) {
    var id = String(row[USER_COL.MEMBER_ID - 1] || '').trim();
    if (!id) ID無し++;
    else { if (使ったID[id]) 重複++; 使ったID[id] = true; }
    if (!String(row[USER_COL.NAME - 1] || '').trim()) 名無し++;
    if (String(row[USER_COL.PUSH - 1] || '').trim()) 届け先++;
    if (String(row[USER_COL.PASSCODE - 1] || '').trim()) パス++;
    if (String(row[USER_COL.DELETE_STATUS - 1] || '').trim()) 削除++;
    if (String(row[USER_COL.BIJIRIS - 1] || '').trim()) ビジリス++;
  });
  Logger.log('  会員: ' + 件 + '名');
  Logger.log('    通知の届け先: ' + 届け先 + '名');
  Logger.log('    パスコード  : ' + パス + '名');
  Logger.log('    削除の印    : ' + 削除 + '名');
  Logger.log('    ビジリス登録: ' + ビジリス + '名');
  Logger.log('    **会員IDが空: ' + ID無し + '件 / 重複: ' + 重複 + '件 / お名前が空: ' + 名無し + '件**');

  // ── パスコードの桁（0が落ちていないか）──
  var 短い = 0, 変 = 0;
  値.forEach(function (row) {
    var p = String(row[USER_COL.PASSCODE - 1] || '').trim();
    if (!p) return;
    if (!/^\d+$/.test(p)) { 変++; return; }
    if (p.length !== 4 && p.length !== 6) 短い++;
  });
  Logger.log('    **桁がおかしいパスコード: ' + 短い + '名**（0が落ちている疑い）');
  Logger.log('    数字でないパスコード: ' + 変 + '名');

  // ── 電話番号（0が落ちていないか）──
  var 電話短い = 0;
  値.forEach(function (row) {
    var t = String(row[USER_COL.PHONE - 1] || '').replace(/\D/g, '');
    if (t && t.charAt(0) !== '0') 電話短い++;
  });
  Logger.log('    **先頭が0でない電話番号: ' + 電話短い + '件**');

  // ── 注文 ──
  var o = ss.getSheetByName(SHEETS.ORDERS);
  var 注文行 = o ? Math.max(0, o.getLastRow() - 1) : 0;
  Logger.log('');
  Logger.log('  注文: ' + 注文行 + '行');

  Logger.log('');
  Logger.log('  ※ 読むだけです。中身（お名前・電話番号）は出していません。');
}
