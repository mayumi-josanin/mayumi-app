// 9/19 の切り替え前に、会員の窓口がサーバーに届くかを見る道具。
//
//   会員の道を確かめる()   … シートの会員数と、サーバーの答えの件数を並べる
//
// **読むだけ。**SERVER_TABLES は変えない。お名前などの中身は出さない。
//
// なぜ要るのか:
//   `サーバーへの道を確かめる()` はお知らせの窓口しか叩かない。会員の窓口
//   （getAdminUsers / getPushUsers / checkMemberOnServer）は合鍵が要り、
//   届かなければ切り替えた瞬間に管理アプリが古いシートを見続ける。
//   **切り替える前に、会員の窓口そのもので確かめる。**
//   サーバーは 2026-09-15 の稽古で取り込んだ会員を持っているので、
//   シートと件数が合っていれば「道が通っている」と「中身が新しい」の両方が分かる。

function 会員の道を確かめる() {
  Logger.log('■ 会員の道');
  Logger.log('  行き先        : ' + 渡す_URL_());
  Logger.log('  渡している表  : ' + (渡す_設定_('SERVER_TABLES', '') || '**なし**'));
  Logger.log('');

  // ── シート側（本番）──
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet());
  var 最終 = sheet.getLastRow();
  var 値 = 最終 >= 2 ? sheet.getRange(2, 1, 最終 - 1, USER_HEADERS.length).getDisplayValues() : [];
  var シート会員 = 0, シート届け先 = 0;
  値.forEach(function (row) {
    if (!String(row[USER_COL.MEMBER_ID - 1] || '').trim()) return;
    if (String(row[USER_COL.DELETE_STATUS - 1] || '').trim() === SOFT_DELETE_STATUS) return;
    シート会員 += 1;
    var 届 = String(row[USER_COL.PUSH - 1] || '').trim();
    // "true"/"false" は届け先ではない（取り込みも空にする）
    if (届 && 届 !== 'true' && 届 !== 'false') シート届け先 += 1;
  });
  Logger.log('  シート: 会員 ' + シート会員 + '名 ／ 届け先あり ' + シート届け先 + '名');

  // ── サーバー側 ──
  var t = new Date().getTime();
  var 一覧 = サーバーから読む_('getAdminUsers');
  var かかった = new Date().getTime() - t;
  if (!一覧 || !一覧.users) {
    Logger.log('  サーバー getAdminUsers: **届きません**（合鍵 SERVER_API_KEY か行き先を確かめてください）');
    Logger.log('  **まだ切り替えないでください。**');
    return;
  }
  Logger.log('  サーバー: getAdminUsers ' + 一覧.users.length + '名（' + かかった + 'ミリ秒）');

  var 届け先 = サーバーから読む_('getPushUsers');
  var 届け先数 = (届け先 && 届け先.users) ? 届け先.users.filter(function (u) { return String(u.subscription || '').trim(); }).length : -1;
  Logger.log('  サーバー: getPushUsers 届け先あり ' + (届け先数 < 0 ? '**届きません**' : 届け先数 + '名'));

  // 存在しない会員IDで、会員の札の窓口が「無効」と答えるか（中身は何も返らない）
  var 札 = サーバーから読む_('checkMemberOnServer', { memberId: 'MYM-0000' });
  Logger.log('  サーバー: checkMemberOnServer（いない会員）→ ' + (札 ? ('valid=' + 札.valid) : '**届きません**'));

  Logger.log('');
  var 会員一致 = (一覧.users.length === シート会員);
  var 届け先一致 = (届け先数 === シート届け先);
  Logger.log('  会員数    : ' + (会員一致 ? '一致' : '**違います**（シート ' + シート会員 + ' / サーバー ' + 一覧.users.length + '）→ 書き出し→取り込みをやり直す'));
  Logger.log('  届け先    : ' + (届け先一致 ? '一致' : '**違います**（シート ' + シート届け先 + ' / サーバー ' + 届け先数 + '）'));
  if (会員一致 && 届け先一致 && 札 && 札.valid === false) {
    Logger.log('');
    Logger.log('  **会員の道は通っています。**当日は 書き出し→取り込み→突き合わせ のあと SERVER_TABLES に member を足してください。');
  }
  Logger.log('');
  Logger.log('  ※ これは見るだけです。渡す設定は変わっていません。');
}
