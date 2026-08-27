// パスコードの先頭の「0」が消えているのを戻す道具。
//
//   パスコードの下見()   … 何件が短くなっているかを数えるだけ（**何も書かない**）
//   パスコードの0を戻す() … 控えを取ってから直す
//
// **並びはわざとこの順です。**エディタは開いたファイルの先頭の関数を既定で選びます。
//
// ## なぜ要るのか
//
// 電話番号と同じで、スプレッドシートが数値として受け取ると先頭の0が落ちる。
//
//     0123 → 123      001234 → 1234
//
// ## いままで気づかれなかった理由
//
// GAS は照合のときに**0を補って**比べている（`matchesRowPasscode_`）。
//
//     stored "123" を input "0123" の長さに合わせて "0123" にしてから比べる
//
// だから**ログインは通っていた。**記録が壊れていても表に出なかった。
//
// ## なぜ、いま直すのか
//
// **サーバーはパスコードをハッシュで持つ。**元の文字が違えば一致しない。
// 補いようがないので、**切り替えた瞬間にその方は入れなくなる。**
//
// 切り替える前に、記録そのものを正しくしておく必要がある。
//
// ## 直し方の考え方
//
// パスコードは**4桁か6桁**しか作れない（`isValidPasscodeValue_`）。
// だから桁が足りないものは、頭に0を足せば元に戻る。
//
//     1〜3桁 → 4桁に揃える     5桁 → 6桁に揃える
//     4桁・6桁 → **触らない**（正しい可能性が高い）
//
// 4桁のものが本当は6桁だった場合（001234 → 1234）は判別できないが、
// GAS 側は補って比べるので**いまも通っている**。サーバー側で通らなくなるが、
// **推測で6桁に伸ばすと、本当に4桁の方が入れなくなる。**触らない。

var パス戻し_列 = 17;  // USER_COL.PASSCODE

function パスコードの0を戻す() {
  var r = パス戻し_調べる_();
  パス戻し_出す_(r, true);

  if (!r.直す.length) {
    Logger.log('');
    Logger.log('■ 直すものがありません。');
    return;
  }

  // **控えを取る。**実データを書き換える前に必ず。
  var ss = getOrCreateSpreadsheet();
  var 控え名 = '会員データ_パスコードを直す前_'
    + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm');
  ss.getSheetByName('会員データ').copyTo(ss).setName(控え名);
  Logger.log('');
  Logger.log('■ 控えを作りました: ' + 控え名);

  // getOrCreateUsersSheet_ を通す。中で列を「書式なしテキスト」にしてくれる。
  var sheet = getOrCreateUsersSheet_(ss);
  r.直す.forEach(function (x) {
    sheet.getRange(x.行番号, パス戻し_列).setValue(x.あと);
  });

  Logger.log('■ **' + r.直す.length + '件を直しました。**');
  Logger.log('   ログインの通り方は変わりません（GASは元から0を補って比べていました）。');
  Logger.log('   **サーバーへ切り替えたときに効いてきます。**');
}

function パスコードの下見() {
  var r = パス戻し_調べる_();
  パス戻し_出す_(r, false);
}

function パス戻し_調べる_() {
  var sheet = getOrCreateUsersSheet_(getOrCreateSpreadsheet());
  var 最終行 = sheet.getLastRow();
  var 値 = sheet.getRange(2, 1, 最終行 - 1, USER_HEADERS.length).getValues();

  var 直す = [];
  var そのまま = 0;
  var 空 = 0;
  var 変な桁 = [];
  var 桁ごと = {};

  値.forEach(function (row, i) {
    var 会員ID = String(row[USER_COL.MEMBER_ID - 1] || '').trim();
    if (!会員ID) return;

    var 生 = row[パス戻し_列 - 1];
    var 数字 = String(生 == null ? '' : 生).replace(/[^0-9]/g, '');
    if (!数字) { 空++; return; }

    桁ごと[数字.length] = (桁ごと[数字.length] || 0) + 1;

    if (数字.length === 4 || 数字.length === 6) { そのまま++; return; }

    var あと = '';
    if (数字.length < 4) あと = ('0000' + 数字).slice(-4);
    else if (数字.length === 5) あと = ('000000' + 数字).slice(-6);
    else { 変な桁.push(会員ID + '（' + 数字.length + '桁）'); return; }

    直す.push({ 行番号: i + 2, 会員ID: 会員ID, まえ桁: 数字.length, あと: あと });
  });

  return { 直す: 直す, そのまま: そのまま, 空: 空, 変な桁: 変な桁, 桁ごと: 桁ごと, 全体: 値.length };
}

function パス戻し_出す_(r, 直したか) {
  Logger.log('■ 会員データ: ' + r.全体 + '名');
  Logger.log('');
  Logger.log('   パスコードが空:   ' + r.空 + '名');
  Logger.log('   4桁か6桁（正常）: ' + r.そのまま + '名');
  Logger.log('   **短くなっている: ' + r.直す.length + '名**');
  if (r.変な桁.length) {
    Logger.log('   **説明のつかない桁: ' + r.変な桁.length + '名**');
    r.変な桁.forEach(function (x) { Logger.log('       ' + x); });
  }
  Logger.log('');
  Logger.log('■ 桁ごとの数（**中身は出しません**）');
  Object.keys(r.桁ごと).sort(function (a, b) { return a - b; }).forEach(function (k) {
    var 印 = (k === '4' || k === '6') ? '' : '  ← **足りない**';
    Logger.log('     ' + k + '桁: ' + r.桁ごと[k] + '名' + 印);
  });

  if (r.直す.length) {
    Logger.log('');
    Logger.log('■ 直すもの（**パスコードそのものは出しません**）');
    r.直す.slice(0, 10).forEach(function (x) {
      Logger.log('     ' + x.行番号 + '行目 ' + x.会員ID + '  ' + x.まえ桁 + '桁 → ' + x.あと.length + '桁');
    });
    if (r.直す.length > 10) Logger.log('     …ほか ' + (r.直す.length - 10) + '名');
  }

  if (!直したか) {
    Logger.log('');
    Logger.log('■ **下見なので、何も書いていません。**');
    Logger.log('   直すときは パスコードの0を戻す() を実行してください。');
  }
}
