// 段階Bの前に、**GASが作った札をサーバー側が同じと判定できるか**を確かめる道具。
//
//   札を確かめる()   … 見本の札を作って、その中身を出す
//
// **読むだけ。**実データは変えない。存在しない会員IDで作るので、
// この札で誰かがログインできるわけでもない。
//
// なぜ要るのか:
//   [移行設計] の絶対ルール「お客様の入り直しを発生させない」。
//   いま発行済みの札は、お客様の端末に入っている。切り替えたあとも
//   **同じ札が通らないと、全員がログイン画面に戻される。**
//   ADMIN_TOKEN_SECRET は変えない前提だが、**実物で確かめるまで
//   「たぶん合う」で進めてはいけない。**
//
// 出すもの:
//   ・札そのもの（base64url）
//   ・中身（t / u / exp / sig）
//   ・署名のもとになった文字列
//   これをサーバー側の Python に渡して、同じ sig が出るかを比べる。

function 札を確かめる() {
  // 実在しない会員IDを使う。**本物の札を画面に出さないため。**
  var 見本のID = 'MYM-0000-TEST';
  var 期限 = 1893456000000;   // 2030-01-01 固定。毎回同じ札が出るように

  var 札 = makeMemberToken_(見本のID, 期限);
  var 中身 = JSON.parse(
    Utilities.newBlob(Utilities.base64DecodeWebSafe(札)).getDataAsString()
  );

  Logger.log('■ 見本の札（存在しない会員IDで作ったもの）');
  Logger.log('');
  Logger.log('  会員ID : ' + 見本のID);
  Logger.log('  期限   : ' + 期限);
  Logger.log('');
  Logger.log('  署名のもと : member|' + 見本のID + '|' + 期限);
  Logger.log('  署名      : ' + 中身.sig);
  Logger.log('');
  Logger.log('  札        : ' + 札);
  Logger.log('');

  // 自分で検証できることも確かめる（作り方と読み方が食い違っていないか）
  var 戻り = verifyMemberToken_(札);
  Logger.log('  自分で検証 : ' + (戻り === 見本のID ? '通った' : '**通らない：' + 戻り + '**'));
  Logger.log('');

  // 管理者の札も同じ仕組みなので、あわせて確かめる
  var 管理札 = makeAdminToken_(期限);
  var 管理中身 = JSON.parse(
    Utilities.newBlob(Utilities.base64DecodeWebSafe(管理札)).getDataAsString()
  );
  Logger.log('■ 管理者の札');
  Logger.log('  署名のもと : admin|' + 期限);
  Logger.log('  署名      : ' + 管理中身.sig);
  Logger.log('');

  Logger.log('■ 確かめ方');
  Logger.log('  サーバー側の Python で、同じ「署名のもと」と');
  Logger.log('  ADMIN_TOKEN_SECRET から HMAC-SHA256 を計算し、');
  Logger.log('  **上の署名と1文字も違わないこと**を確かめる。');
  Logger.log('');
  Logger.log('  合わなければ、切り替えた瞬間にお客様全員が');
  Logger.log('  ログイン画面に戻される。**合うまで切り替えない。**');
  Logger.log('');
  Logger.log('  ※ 読むだけです。実データは変えていません。');
  Logger.log('  ※ 見本は存在しない会員IDなので、この札では誰も入れません。');
}


// サーバーが計算した署名が正しいかを、GAS自身に判定させる。
//
//   サーバーの署名を判定する()
//
// 画面の文字を人が写すと、l と I、0 と O、c と g のような読み違いが起きる。
// **写さずに、GASに直接比べさせる。**
var 判定したい署名 = {
  'member|MYM-0000-TEST|1893456000000': 'e_LOXY7kgSYmgxnL7CqQv15fzbpijFvTw95O6gMRrvc=',
  'admin|1893456000000': '0bWBVhv3WvGquKq34othsepm5Rdn_nP-eSUifEDFUp4='
};

function サーバーの署名を判定する() {
  Logger.log('■ サーバーが計算した署名を、GASの署名と比べます');
  Logger.log('');
  var 全部合った = true;
  Object.keys(判定したい署名).forEach(function (もと) {
    var 正しい = adminSign_(もと);
    var 渡された = 判定したい署名[もと];
    var 合う = (正しい === 渡された);
    if (!合う) 全部合った = false;
    Logger.log('  ' + もと);
    Logger.log('    ' + (合う ? '**一致しました**' : '**違います**'));
    if (!合う) {
      Logger.log('      GASの署名     : ' + 正しい);
      Logger.log('      渡された署名   : ' + 渡された);
    }
    Logger.log('');
  });
  Logger.log(全部合った
    ? '■ すべて一致。サーバーはGASと同じ札を作れます。'
    : '■ **一致しないものがあります。切り替えてはいけません。**');
  Logger.log('  ※ 読むだけです。');
}
