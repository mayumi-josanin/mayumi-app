// 回数券スタンプ（1枚使い切るごとに1個）の下見。
//
// **読むだけ。**何も書かない。
//
// 数え方は Code.gs の getCompletedTicketCardCountsByCustomer_ をそのまま使う。
// ここで別に数えると、管理アプリの表示とズレて「どちらが正しいか」が分からなくなる。
//
// 過去にお使いいただいた分も遡って数えるので、始めた瞬間に
// 受け取り資格をお持ちの方が何人か出る。その人数と個数を先に把握するための道具。

// 回答が0件と出たので、シートそのものを数える。読むだけ。
function シートの中身を数える() {
  var book = getSpreadsheet_();
  Logger.log('■ ' + book.getName());
  Logger.log('  ' + book.getUrl());
  Logger.log('');
  book.getSheets().forEach(function (sh) {
    Logger.log('  ' + sh.getName() + ' … ' + Math.max(0, sh.getLastRow() - 1) + '行（見出しを除く）');
  });
  Logger.log('');
  var master = book.getSheetByName(MASTER_SHEET_NAME);
  Logger.log('  回答一覧の見出し行: ' + (master ? JSON.stringify(master.getRange(1, 1, 1, MASTER_HEADERS.length).getValues()[0]) : 'シートが無い'));
  if (master && master.getLastRow() >= 2) {
    Logger.log('  1行目のデータ: ' + JSON.stringify(master.getRange(2, 1, 1, MASTER_HEADERS.length).getValues()[0]).slice(0, 300));
  }
}

// なぜ0件なのかを調べる。数え方は「最終回の回答が送信されていること」なので、
// 最終回だけアンケートを出し忘れると、使い切っていても1個も付かない。
function なぜ数えられないのかを調べる() {
  var responses = getResponses_({});
  Logger.log('■ 回答の総数: ' + responses.length);

  var 回数券の回答 = 0;
  var 揃っている = 0;
  var 最高回 = {};      // お名前 → { カード: 最大round }
  var 種類の内訳 = {};
  var 欠け = { plan: 0, sheet: 0, round: 0 };

  responses.forEach(function (response) {
    var name = normalizeText_(response && response.customerName);
    var answers = response && response.answers;
    var plan = getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.plan);
    var sheetRaw = getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.sheet);
    var roundRaw = getAnswerValueByQuestionIds_(answers, CUSTOMER_TICKET_INFO_QUESTION_IDS.round);
    if (!plan && !sheetRaw && !roundRaw) return;
    回数券の回答 += 1;
    種類の内訳[String(plan || '(空)')] = (種類の内訳[String(plan || '(空)')] || 0) + 1;

    var sheet = parseTicketLabelNumber_(sheetRaw);
    var round = parseTicketLabelNumber_(roundRaw);
    if (!plan) 欠け.plan += 1;
    if (!(sheet > 0)) 欠け.sheet += 1;
    if (!(round > 0)) 欠け.round += 1;
    if (!name || !plan || sheet <= 0 || round <= 0) return;
    揃っている += 1;

    var key = name + ' / ' + plan + ' ' + sheet + '枚目';
    if (!最高回[key] || round > 最高回[key]) 最高回[key] = round;
  });

  Logger.log('  うち回数券に触れている回答: ' + 回数券の回答);
  Logger.log('  3つとも揃っている回答: ' + 揃っている);
  Logger.log('  欠け → 種類:' + 欠け.plan + ' 何枚目:' + 欠け.sheet + ' 何回目:' + 欠け.round);
  Logger.log('');
  Logger.log('  種類の内訳: ' + JSON.stringify(種類の内訳));
  Logger.log('');
  Logger.log('■ カードごとの「いちばん進んだ回」');
  Logger.log('  （6回券なら6、10回券なら10 に届いて初めて1個になる）');
  Logger.log('');
  var keys = Object.keys(最高回);
  keys.sort(function (a, b) { return 最高回[b] - 最高回[a]; });
  keys.slice(0, 30).forEach(function (k) {
    Logger.log('    ' + 最高回[k] + '回目まで … ' + k);
  });
  Logger.log('');
  Logger.log('  カードの数: ' + keys.length);
}

function 回数券スタンプの下見() {
  var responses = getResponses_({});
  var counts = getCompletedTicketCardCountsByCustomer_(responses);

  var 名前 = Object.keys(counts).filter(function (n) { return counts[n] > 0; });
  名前.sort(function (a, b) { return counts[b] - counts[a]; });

  var 分布 = {};
  var 合計 = 0;
  名前.forEach(function (n) {
    分布[counts[n]] = (分布[counts[n]] || 0) + 1;
    合計 += counts[n];
  });

  Logger.log('■ 回数券を使い切った枚数（＝スタンプの数）');
  Logger.log('');
  Logger.log('  1枚以上の方: ' + 名前.length + '名 ／ スタンプ総数: ' + 合計 + '個');
  Logger.log('');

  Logger.log('  個数ごとの人数:');
  Object.keys(分布)
    .map(Number)
    .sort(function (a, b) { return a - b; })
    .forEach(function (k) {
      Logger.log('    ' + k + '個 … ' + 分布[k] + '名');
    });
  Logger.log('');

  // 節目にいくつ置くかを決める材料。ここを超えている方には、
  // 始めた時点で受け取っていただくものが要る。
  [1, 3, 5, 10].forEach(function (節目) {
    var 該当 = 名前.filter(function (n) { return counts[n] >= 節目; });
    Logger.log('  ' + 節目 + '個以上: ' + 該当.length + '名'
      + (該当.length && 該当.length <= 12 ? '（' + 該当.join('・') + '）' : ''));
  });
  Logger.log('');

  Logger.log('  上位の方:');
  名前.slice(0, 15).forEach(function (n) {
    Logger.log('    ' + counts[n] + '個 … ' + n);
  });
  Logger.log('');
  Logger.log('  ※ 回答履歴から算出。管理アプリの手動カード設定は含めない。');
}
