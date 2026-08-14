// データの片付け。消す前に必ず「下見」で中身を確かめること。
//
//   お名前が空の行の下見()   … まゆみ会員データの、氏名が空の行に何が入っているかを見る
//   お名前が空の行を消す()   … 控えを取ってから、その行を消す
//   テストデータの下見()     … 「テスト」を含むお名前の記録を数える
//   テストデータを消す()     … 控えを取ってから、その記録を消す
//
// どちらも実行前にスプレッドシートの控え（コピー）を取る。30日間はゴミ箱から戻せる。

// エディタは「ファイルの先頭の関数」を最初に選ぶ。
// 関数の選び直しが効かないことがあるため、実行したいものをここから呼ぶ。
// 既定は「見るだけ」にしておく。うっかり実行しても何も壊れないように。
function いま実行する() {
  会員の一覧を出す();
}

// 自動処理が「素通り」できているかを、2回続けて動かして時間で確かめる。
function 回数券の自動処理を計る() {
  for (var i = 1; i <= 2; i += 1) {
    var t = Date.now();
    runTicketSurveyAutoProcess();
    Logger.log(i + '回目 … ' + ((Date.now() - t) / 1000).toFixed(2) + '秒');
  }
  var meta = getTicketSurveyMeta_();
  Logger.log('');
  Logger.log('覚えている行数: ' + meta.lastSeenResponseRows);
  Logger.log('前回きちんと通した時刻: ' + meta.lastFullRunAt);
}

// 日次バックアップが本当に効いているかを確かめ、いまの内容で1本作る。
function バックアップを確かめる() {
  var p = getPreferences_();
  Logger.log('自動バックアップ: ' + (p.autoBackupEnabled ? 'オン' : 'オフ'));

  // getProjectTriggers は「いま実行している人が作ったもの」しか返さない。
  // 別のアカウントが作ったトリガーは 0 件に見えるので、無いと決めつけないこと。
  // 全部を見るには、エディタ左の時計アイコン（トリガー画面）を開く。
  var 予定 = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  Logger.log('この実行者が作った自動実行: ' + (予定.join('・') || 'なし'));
  Logger.log('  ※ 他の人が作ったものは見えません。トリガー画面で確認してください。');
  Logger.log('');

  var meta = writeBackupFile_();
  Logger.log('いまの内容で1本作りました: ' + meta.fileName);
  Logger.log('  ' + meta.fileUrl);

  var 中身 = JSON.parse(DriveApp.getFileById(meta.fileId).getBlob().getDataAsString());
  Logger.log('');
  Logger.log('入っている項目:');
  Object.keys(中身).sort().forEach(function (k) {
    var v = 中身[k];
    var 量 = Array.isArray(v) ? v.length + '件'
      : (v && typeof v === 'object') ? Object.keys(v).length + '項目'
      : String(v).length + '文字';
    Logger.log('  ' + k + ' … ' + 量);
  });
}

var 片付け_まゆみID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';
var 片付け_ビジリスID = '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E';

// 「テスト」として扱うお名前。含まれていれば対象。
var 片付け_テストの名前 = ['テスト'];

function 片付け_空か_(x) { return x === '' || x === null || x === undefined; }

function 片付け_控えを取る_(id, ラベル) {
  var f = DriveApp.getFileById(id);
  var 控え = f.makeCopy(
    '【控え】' + f.getName() + ' ' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + ' ' + ラベル
  );
  Logger.log('控えを作りました: ' + 控え.getName());
  Logger.log('  ' + 控え.getUrl());
  Logger.log('');
  return 控え;
}

// ---------- ① お名前が空の行 ----------

function 片付け_空の行を探す_() {
  var sh = SpreadsheetApp.openById(片付け_まゆみID).getSheetByName('会員データ');
  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  var v = sh.getRange(1, 1, 行数, 列数).getValues();
  var 見出し = v[0];
  var 名列 = 見出し.indexOf('氏名');
  var 対象 = [];
  for (var r = 1; r < v.length; r += 1) {
    var row = v[r];
    if (!row.some(function (x) { return !片付け_空か_(x); })) continue;   // 完全な空行は無視
    if (String(row[名列] || '').trim()) continue;
    対象.push({ 行: r + 1, 値: row });
  }
  return { sheet: sh, 見出し: 見出し, 対象: 対象 };
}

function お名前が空の行の下見() {
  var 結果 = 片付け_空の行を探す_();
  Logger.log('■ お名前が空の行: ' + 結果.対象.length + '件');
  Logger.log('  （値そのものは出さず、どの項目が入っているかだけを出します）');
  Logger.log('');
  結果.対象.forEach(function (t) {
    var 入っている = [];
    t.値.forEach(function (x, i) {
      if (!片付け_空か_(x)) 入っている.push(結果.見出し[i]);
    });
    Logger.log('  ' + t.行 + '行目 … ' + 入っている.join('・'));
  });
  Logger.log('');
  Logger.log('消してよければ「お名前が空の行を消す」を実行してください。');
}

function お名前が空の行を消す() {
  var 結果 = 片付け_空の行を探す_();
  if (!結果.対象.length) { Logger.log('お名前が空の行はありません'); return; }

  片付け_控えを取る_(片付け_まゆみID, '会員データ整理前');

  var 行 = 結果.対象.map(function (t) { return t.行; });
  行.sort(function (a, b) { return b - a; });   // 下から消す
  行.forEach(function (r) { 結果.sheet.deleteRow(r); });

  Logger.log('■ 消しました: ' + 行.length + '行（' + 行.slice().reverse().join(', ') + '行目）');
  Logger.log('  控えから戻せます。');
}

// ---------- ② テストデータ ----------

function 片付け_テストか_(名前) {
  var s = String(名前 || '').trim();
  if (!s) return false;
  return 片付け_テストの名前.some(function (t) { return s.indexOf(t) >= 0; });
}

// 名前が入っていそうな列を探す
function 片付け_名前列_(見出し) {
  var 候補 = ['お名前', '顧客名', '氏名', '注文者名'];
  for (var i = 0; i < 候補.length; i += 1) {
    var n = 見出し.indexOf(候補[i]);
    if (n >= 0) return n;
  }
  return -1;
}

function 片付け_テストを集める_(id) {
  var ss = SpreadsheetApp.openById(id);
  var 一覧 = [];
  ss.getSheets().forEach(function (sh) {
    var 行数 = sh.getLastRow();
    var 列数 = sh.getLastColumn();
    if (行数 < 2 || 列数 < 1) return;
    var 見出し = sh.getRange(1, 1, 1, 列数).getValues()[0];
    var 名列 = 片付け_名前列_(見出し);
    if (名列 < 0) return;
    var 値 = sh.getRange(2, 名列 + 1, 行数 - 1, 1).getValues();
    var 行 = [];
    var 名前ごと = {};
    値.forEach(function (r, i) {
      if (!片付け_テストか_(r[0])) return;
      行.push(i + 2);
      var n = String(r[0]).trim();
      名前ごと[n] = (名前ごと[n] || 0) + 1;
    });
    if (行.length) 一覧.push({ ss: ss, sheet: sh, 行: 行, 名前ごと: 名前ごと });
  });
  return 一覧;
}

function テストデータの下見() {
  Logger.log('■ 「' + 片付け_テストの名前.join('」「') + '」を含むお名前の記録');
  Logger.log('');
  var 合計 = 0;
  [['ビジリス', 片付け_ビジリスID], ['まゆみ', 片付け_まゆみID]].forEach(function (pair) {
    Logger.log('▼ ' + pair[0]);
    var 一覧 = 片付け_テストを集める_(pair[1]);
    if (!一覧.length) { Logger.log('  該当なし'); Logger.log(''); return; }
    一覧.forEach(function (t) {
      合計 += t.行.length;
      Logger.log('  ' + t.sheet.getName() + ' … ' + t.行.length + '行');
      Object.keys(t.名前ごと).sort().forEach(function (n) {
        Logger.log('     ' + n + ' … ' + t.名前ごと[n] + '行');
      });
    });
    Logger.log('');
  });
  Logger.log('合計 ' + 合計 + '行');
  Logger.log('');
  Logger.log('消してよければ「テストデータを消す」を実行してください。');
  Logger.log('会員別まとめのシートは、そのあと作り直せば自動で消えます。');
}

function テストデータを消す() {
  var 予定 = [];
  [['ビジリス', 片付け_ビジリスID], ['まゆみ', 片付け_まゆみID]].forEach(function (pair) {
    var 一覧 = 片付け_テストを集める_(pair[1]);
    if (一覧.length) 予定.push({ ラベル: pair[0], id: pair[1], 一覧: 一覧 });
  });
  if (!予定.length) { Logger.log('消すものはありません'); return; }

  予定.forEach(function (p) {
    片付け_控えを取る_(p.id, 'テスト削除前');
    var 合計 = 0;
    p.一覧.forEach(function (t) {
      var 行 = t.行.slice().sort(function (a, b) { return b - a; });
      行.forEach(function (r) { t.sheet.deleteRow(r); });
      合計 += 行.length;
      Logger.log('  ' + t.sheet.getName() + ' … ' + 行.length + '行を削除');
    });
    Logger.log('▼ ' + p.ラベル + ' 合計 ' + 合計 + '行');
    Logger.log('');
  });

  // 会員別まとめを作り直すと、テストのシートも自動で外れる
  try {
    測定記録の会員別一覧を作る();
    会員ごとのファイルを作る();
    Logger.log('会員別まとめを作り直しました');
  } catch (error) {
    Logger.log('会員別まとめの作り直しに失敗: ' + error.message);
  }
}
