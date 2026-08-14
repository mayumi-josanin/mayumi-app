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
  移行の下見をする();
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

// ---------- 会員のお名前を付け替える ----------
//
// ビジリスの記録は「お名前の文字列」で会員と結びつくため、表記がそろっていないと
// 別人として扱われる。同じ方が漢字とカタカナで登録されている場合にそろえる。
// 新しく登録し直すと二重登録になるので、必ず付け替えで対応すること。

var 付替え_旧 = 'コセムラマリコ';
var 付替え_新 = '小瀬村真理子';

function お名前の付け替えの下見() {
  var sh = SpreadsheetApp.openById(片付け_まゆみID).getSheetByName('会員データ');
  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  var v = sh.getRange(1, 1, 行数, 列数).getValues();
  var 名列 = v[0].indexOf('氏名');
  var 旧 = [];
  var 新 = [];
  for (var r = 1; r < v.length; r += 1) {
    var n = String(v[r][名列] || '').trim();
    if (n === 付替え_旧) 旧.push(r + 1);
    if (n === 付替え_新) 新.push(r + 1);
  }
  Logger.log('「' + 付替え_旧 + '」… ' + 旧.length + '件（行 ' + (旧.join(', ') || 'なし') + '）');
  Logger.log('「' + 付替え_新 + '」… ' + 新.length + '件（行 ' + (新.join(', ') || 'なし') + '）');
  Logger.log('');
  if (旧.length === 1 && 新.length === 0) {
    Logger.log('付け替えできます。「会員のお名前を付け替える」を実行してください。');
  } else if (新.length) {
    Logger.log('※ すでに新しいお名前の登録があります。統合が必要なため、付け替えでは対応できません。');
  } else {
    Logger.log('※ 対象が1件ではありません。確認してください。');
  }
}

function 会員のお名前を付け替える() {
  var sh = SpreadsheetApp.openById(片付け_まゆみID).getSheetByName('会員データ');
  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  var v = sh.getRange(1, 1, 行数, 列数).getValues();
  var 名列 = v[0].indexOf('氏名');

  var 対象 = [];
  for (var r = 1; r < v.length; r += 1) {
    var n = String(v[r][名列] || '').trim();
    if (n === 付替え_新) { Logger.log('すでに「' + 付替え_新 + '」の登録があります。中止します。'); return; }
    if (n === 付替え_旧) 対象.push(r + 1);
  }
  if (対象.length !== 1) { Logger.log('対象が1件ではありません（' + 対象.length + '件）。中止します。'); return; }

  片付け_控えを取る_(片付け_まゆみID, 'お名前付け替え前');
  sh.getRange(対象[0], 名列 + 1).setValue(付替え_新);
  Logger.log('■ ' + 対象[0] + '行目のお名前を「' + 付替え_旧 + '」→「' + 付替え_新 + '」に変えました。');
  Logger.log('  会員IDもスタンプもそのままです。ビジリスの記録と結びつくようになります。');
}

// ---------- ② テストデータ ----------

// 「テスト」を含んでいても、絶対に消してはいけないお名前。
// 管理者アカウントの氏名を「テスト」にしたため、この守りが無いと
// テストデータの掃除で管理者ごと消えて、誰も管理アプリに入れなくなる。
var 片付け_守る名前 = ['テスト'];

function 片付け_テストか_(名前) {
  var s = String(名前 || '').trim();
  if (!s) return false;
  if (片付け_守る名前.indexOf(s) >= 0) return false;   // 完全一致で守る
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
    var 権限列 = 見出し.indexOf('権限');   // 会員データにだけある
    var 値 = sh.getRange(2, 1, 行数 - 1, 列数).getValues();
    var 行 = [];
    var 名前ごと = {};
    値.forEach(function (r, i) {
      // 管理者の行は、名前が何であっても消さない。
      if (権限列 >= 0 && String(r[権限列] || '').trim() === '管理者') return;
      if (!片付け_テストか_(r[名列])) return;
      行.push(i + 2);
      var n = String(r[名列]).trim();
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

// ---------- 指定した会員を消す ----------
//
// お名前の「完全一致」だけを対象にする。部分一致にすると
// 「あゆみ」で「尾形あゆみ」、「伊藤」で「伊藤由貴」まで巻き込む。
//
// 管理者の行は、名前が一致しても絶対に消さない。
// 消すと管理アプリとビジリス管理に誰も入れなくなるため。

var 消す会員の名前 = ['ま', 'ひろき', 'あゆみ', 'デモ2', '伊藤', 'YUKOOCHI', '釼持裕基'];

function 指定した会員の下見() {
  var 結果 = 指定会員_探す_();
  Logger.log('■ 消す対象（お名前が完全に一致するもの）');
  Logger.log('');
  結果.見つかった.forEach(function (t) {
    Logger.log('  ' + t.行 + '行 ' + t.id + ' ' + t.名 +
      ' … 項目' + t.埋 + '個 / スタンプ' + t.スタンプ + ' / 登録' + t.登録);
  });
  Logger.log('');
  if (結果.守った.length) {
    Logger.log('▼ 一致したが管理者なので守ったもの');
    結果.守った.forEach(function (t) { Logger.log('  ' + t.行 + '行 ' + t.名); });
    Logger.log('');
  }
  if (結果.見つからず.length) {
    Logger.log('▼ 見つからなかったお名前: ' + 結果.見つからず.join('・'));
    Logger.log('');
  }
  Logger.log('合計 ' + 結果.見つかった.length + '名。消してよければ「指定した会員を消す」を実行してください。');
}

function 指定会員_探す_() {
  var sh = SpreadsheetApp.openById(片付け_まゆみID).getSheetByName('会員データ');
  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  var v = sh.getRange(1, 1, 行数, 列数).getValues();
  var 位 = {};
  v[0].forEach(function (h, i) { 位[String(h)] = i; });

  var 見つかった = [];
  var 守った = [];
  var 出た = {};
  for (var r = 1; r < v.length; r += 1) {
    var row = v[r];
    var 名 = String(row[位['氏名']] || '').trim();
    if (消す会員の名前.indexOf(名) < 0) continue;
    出た[名] = true;
    var t = {
      行: r + 1,
      id: String(row[位['ID']] || ''),
      名: 名,
      スタンプ: row[位['現在スタンプ数']],
      登録: row[位['登録日時']],
      埋: row.filter(function (x) { return !片付け_空か_(x); }).length,
    };
    if (String(row[位['権限']] || '').trim() === '管理者') 守った.push(t);
    else 見つかった.push(t);
  }
  var 見つからず = 消す会員の名前.filter(function (n) { return !出た[n]; });
  return { sheet: sh, 見つかった: 見つかった, 守った: 守った, 見つからず: 見つからず };
}

function 指定した会員を消す() {
  var 結果 = 指定会員_探す_();
  if (!結果.見つかった.length) { Logger.log('消す対象がありません'); return; }

  片付け_控えを取る_(片付け_まゆみID, '指定会員の削除前');

  var 行 = 結果.見つかった.map(function (t) { return t.行; });
  行.sort(function (a, b) { return b - a; });     // 下から消す
  行.forEach(function (r) { 結果.sheet.deleteRow(r); });

  Logger.log('■ ' + 結果.見つかった.length + '名を消しました');
  結果.見つかった.forEach(function (t) { Logger.log('  ' + t.id + ' ' + t.名); });
  if (結果.守った.length) {
    Logger.log('');
    Logger.log('管理者なので守ったもの: ' + 結果.守った.map(function (t) { return t.名; }).join('・'));
  }
  Logger.log('');
  Logger.log('控えから戻せます。');
}

// ---------- 今回の整理をまとめて行う ----------
//
// 控えは最初に1本だけ取る。中の処理は控えを取らない。
//   1. 指定した7名を消す（管理者は対象外）
//   2. 管理者の氏名を「まゆみ」→「テスト」に変える
//   3. 「コセムラマリコ」→「小瀬村真理子」に変える（ビジリスの記録と結びつける）

function 整理_名前を変える_(sh, 旧, 新) {
  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  var v = sh.getRange(1, 1, 行数, 列数).getValues();
  var 名列 = v[0].indexOf('氏名');
  var 対象 = [];
  var すでに = [];
  for (var r = 1; r < v.length; r += 1) {
    var n = String(v[r][名列] || '').trim();
    if (n === 旧) 対象.push(r + 1);
    if (n === 新) すでに.push(r + 1);
  }
  if (すでに.length) {
    Logger.log('  × 「' + 新 + '」がすでに ' + すでに.length + '件あります。変更しません。');
    return false;
  }
  if (対象.length !== 1) {
    Logger.log('  × 「' + 旧 + '」が ' + 対象.length + '件です（1件のときだけ変えます）。変更しません。');
    return false;
  }
  sh.getRange(対象[0], 名列 + 1).setValue(新);
  Logger.log('  ○ ' + 対象[0] + '行目: 「' + 旧 + '」→「' + 新 + '」');
  return true;
}

function 今回の整理をまとめて行う() {
  片付け_控えを取る_(片付け_まゆみID, '整理前');
  var sh = SpreadsheetApp.openById(片付け_まゆみID).getSheetByName('会員データ');

  // 1. 指定した会員を消す
  var 結果 = 指定会員_探す_();
  Logger.log('■ 1. 指定した会員を消す');
  if (!結果.見つかった.length) {
    Logger.log('  対象なし');
  } else {
    var 行 = 結果.見つかった.map(function (t) { return t.行; });
    行.sort(function (a, b) { return b - a; });
    行.forEach(function (r) { 結果.sheet.deleteRow(r); });
    結果.見つかった.forEach(function (t) { Logger.log('  ○ ' + t.id + ' ' + t.名); });
    Logger.log('  ' + 結果.見つかった.length + '名を消しました');
  }
  if (結果.守った.length) {
    Logger.log('  ※ 管理者なので守った: ' + 結果.守った.map(function (t) { return t.名; }).join('・'));
  }
  Logger.log('');

  // 2. 管理者の氏名を変える
  Logger.log('■ 2. 管理者の氏名を「まゆみ」→「テスト」');
  整理_名前を変える_(sh, 'まゆみ', 'テスト');
  Logger.log('');

  // 3. ビジリスの記録と結びつける
  Logger.log('■ 3. 「コセムラマリコ」→「小瀬村真理子」');
  整理_名前を変える_(sh, 'コセムラマリコ', '小瀬村真理子');
  Logger.log('');

  var 残り = sh.getLastRow() - 1;
  Logger.log('会員は ' + 残り + '名になりました。控えから戻せます。');
}
