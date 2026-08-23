// 電話番号の先頭の「0」が消えているのを戻す道具。
//
//   電話番号の下見()   … 何がどう変わるかを見るだけ（**何も書かない**）
//   電話番号の0を戻す() … 控えを取ってから直す
//
// ## なぜ消えるのか
//
// スプレッドシートが電話番号を**数値**として受け取ると、先頭の0が落ちる。
//
//     08012345678  →  8012345678
//
// 153名のうち **121名**がこの形になっていた（2026-08-24）。
// 0が残っている3名は、セルに空白やハイフンが入っていて
// 文字として扱われた方（「080 6758 7760」など）。
//
// ## 直したあと、また消えないようにする
//
// 直すだけでは、次に手で入れたときにまた消える。
// **列の表示形式を「書式なしテキスト」にしてから書き戻す。**
//
// ## 直さないもの
//
//   ・すでに0で始まっているもの
//   ・0を足しても日本の電話番号にならないもの → **画面に出して、人が決める**
//
// 桁数の見かた（0を足したあと）:
//   11桁で 070/080/090 … 携帯
//   10桁で 0で始まる    … 固定電話
//   それ以外            … 判断が要る

var 電話戻し_シート名 = '会員データ';
var 電話戻し_列 = 5;  // USER_COL.PHONE

// **この2つの並びは、わざとこの順にしています。**
// GASのエディタは「開いたファイルの先頭の関数」を既定で選びます。
// 直す側を先頭に置いてあるので、このファイルを開いて実行を押せば
// 電話番号の0を戻す() が動きます。**関数の選び間違いを防ぐため。**
// 2回動かしても害はありません（2回目は直すものが無いので何もしません）。

function 電話番号の0を戻す() {
  var r = 電話戻し_調べる_();
  電話戻し_出す_(r);

  if (!r.直す.length) {
    Logger.log('');
    Logger.log('■ 直すものがありません。');
    return;
  }

  // **控えを取る。**実データを書き換える前に必ず。
  var ss = getOrCreateSpreadsheet();
  var 控え名 = '会員データ_電話番号を直す前_'
    + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm');
  ss.getSheetByName(電話戻し_シート名).copyTo(ss).setName(控え名);
  Logger.log('');
  Logger.log('■ 控えを作りました: ' + 控え名);

  var sheet = ss.getSheetByName(電話戻し_シート名);

  // **先に列を「書式なしテキスト」にする。**
  // これをしないと、書き戻した先頭の0がまた落ちる。
  sheet.getRange(2, 電話戻し_列, sheet.getLastRow() - 1, 1).setNumberFormat('@');

  r.直す.forEach(function (x) {
    sheet.getRange(x.行番号, 電話戻し_列).setValue(x.あと);
  });

  Logger.log('■ **' + r.直す.length + '件を直しました。**');
  Logger.log('   列の表示形式も「書式なしテキスト」にしたので、次からは消えません。');
  if (r.要判断.length) {
    Logger.log('');
    Logger.log('■ **' + r.要判断.length + '件は直していません。**上の「要判断」をご確認ください。');
  }
}

function 電話番号の下見() {
  var r = 電話戻し_調べる_();
  電話戻し_出す_(r);
  Logger.log('');
  Logger.log('■ **下見なので、何も書いていません。**');
  Logger.log('   直すときは 電話番号の0を戻す() を実行してください。');
}

function 電話戻し_調べる_() {
  var sheet = getOrCreateSpreadsheet().getSheetByName(電話戻し_シート名);
  if (!sheet) throw new Error('「' + 電話戻し_シート名 + '」が見つかりません。');

  var 最終行 = sheet.getLastRow();
  var 値 = sheet.getRange(2, 1, 最終行 - 1, Math.max(電話戻し_列, USER_COL.NAME)).getValues();

  var 直す = [];
  var 要判断 = [];
  var そのまま = 0;
  var 空 = 0;

  値.forEach(function (row, i) {
    var 行番号 = i + 2;
    var 会員ID = String(row[USER_COL.MEMBER_ID - 1] || '').trim();
    if (!会員ID) return;

    var 生 = row[電話戻し_列 - 1];
    var 文字 = (生 === '' || 生 === null || 生 === undefined) ? '' : String(生).trim();
    if (!文字) { 空++; return; }

    var 数字 = 文字.replace(/\D/g, '');
    if (!数字) { 要判断.push({ 行番号: 行番号, 会員ID: 会員ID, まえ: 文字, 理由: '数字がありません' }); return; }
    if (数字.charAt(0) === '0') { そのまま++; return; }

    // 国番号81で始まっている（12桁）→ 81を0に置き換える
    if (数字.length === 12 && 数字.slice(0, 2) === '81') {
      直す.push({ 行番号: 行番号, 会員ID: 会員ID, まえ: 文字, あと: '0' + 数字.slice(2), 種: '国番号81' });
      return;
    }

    var 候補 = '0' + 数字;

    // 11桁の携帯（070/080/090）
    if (候補.length === 11 && /^0(70|80|90)/.test(候補)) {
      直す.push({ 行番号: 行番号, 会員ID: 会員ID, まえ: 文字, あと: 候補, 種: '携帯' });
      return;
    }
    // 10桁の固定電話。070〜079 は固定電話の市外局番として存在しないので除く。
    if (候補.length === 10 && /^0[1-9]/.test(候補) && !/^07/.test(候補)) {
      直す.push({ 行番号: 行番号, 会員ID: 会員ID, まえ: 文字, あと: 候補, 種: '固定' });
      return;
    }

    要判断.push({
      行番号: 行番号, 会員ID: 会員ID, まえ: 文字,
      理由: '0を足すと ' + 候補 + '（' + 候補.length + '桁）になり、日本の電話番号の形になりません'
    });
  });

  return { 直す: 直す, 要判断: 要判断, そのまま: そのまま, 空: 空, 全体: 値.length };
}

function 電話戻し_伏せる_(s) {
  var t = String(s || '');
  if (t.length <= 5) return t;
  return t.slice(0, 3) + Array(t.length - 4).join('*') + t.slice(-2);
}

function 電話戻し_出す_(r) {
  Logger.log('■ 会員データ: ' + r.全体 + '行');
  Logger.log('');
  Logger.log('   電話番号が空:        ' + r.空 + '名');
  Logger.log('   すでに0で始まる:      ' + r.そのまま + '名（触りません）');
  Logger.log('   **0を足すもの:        ' + r.直す.length + '名**');
  Logger.log('   **判断が要るもの:      ' + r.要判断.length + '名**');
  Logger.log('');

  var 種ごと = {};
  r.直す.forEach(function (x) { 種ごと[x.種] = (種ごと[x.種] || 0) + 1; });
  Object.keys(種ごと).forEach(function (k) {
    Logger.log('     ' + k + ': ' + 種ごと[k] + '名');
  });

  if (r.直す.length) {
    Logger.log('');
    Logger.log('■ 直すものの例（**番号は伏せています**）');
    r.直す.slice(0, 5).forEach(function (x) {
      Logger.log('     ' + x.行番号 + '行目 ' + x.会員ID + '  '
        + 電話戻し_伏せる_(x.まえ) + ' → ' + 電話戻し_伏せる_(x.あと));
    });
    if (r.直す.length > 5) Logger.log('     …ほか ' + (r.直す.length - 5) + '名');
  }

  if (r.要判断.length) {
    Logger.log('');
    Logger.log('■ **判断が要るもの（直しません）**');
    Logger.log('   受付でご本人に確かめるか、手で直してください。');
    r.要判断.forEach(function (x) {
      Logger.log('     ' + x.行番号 + '行目 ' + x.会員ID + '  '
        + 電話戻し_伏せる_(x.まえ) + '  … ' + x.理由);
    });
  }
}
