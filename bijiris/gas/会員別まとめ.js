// 会員ごとに、測定記録とアンケート回答をまとめるための道具。
// アプリからは呼ばれない。必要なときにエディタから手で実行する。
//
//   測定記録の会員別一覧を作る()  … このファイル内に「測定記録_会員別」を作り直す
//   会員ごとのファイルを作る()    … 別ファイルに、会員1人につき1シートで作り直す
//   競合シートを比べる()          … 「_conflict」シートと本体の差を数える（消す前の確認用）
//
// どれも何度でも実行できる。実行するたびに、そのときの最新の内容で作り直す。

var 会員別_一覧シート名 = '測定記録_会員別';
var 会員別_ファイル名 = 'まゆみ助産院 ビジリス 会員別まとめ';
var 会員別_ファイルID保存キー = 'MEMBER_DIGEST_SPREADSHEET_ID';
// いま使っている「会員別まとめ」のファイル。ここを直せば行き先を変えられる。
var 会員別_ファイルID = '1KXs8e5W_iGtj8c4g2v7M4mhcWCkUUA6fxfs267o4lNw';

// 会員別まとめのファイルを開く（無ければ作る）
function 会員別_ファイルを開く_() {
  var props = PropertiesService.getScriptProperties();
  var 候補 = [会員別_ファイルID, props.getProperty(会員別_ファイルID保存キー)];
  for (var i = 0; i < 候補.length; i += 1) {
    if (!候補[i]) continue;
    try {
      var book = SpreadsheetApp.openById(候補[i]);
      props.setProperty(会員別_ファイルID保存キー, book.getId());
      return book;
    } catch (error) { /* 次の候補へ */ }
  }
  var 新規 = SpreadsheetApp.create(会員別_ファイル名);
  props.setProperty(会員別_ファイルID保存キー, 新規.getId());
  Logger.log('新しいファイルを作りました: ' + 新規.getUrl());
  return 新規;
}

// ---------- 共通 ----------

function 会員別_名前をそろえる_(value) {
  return String(value == null ? '' : value).trim().replace(/[\s　]+/g, '');
}

function 会員別_日付にする_(value) {
  if (value instanceof Date) return value;
  var t = new Date(value);
  return isNaN(t.getTime()) ? null : t;
}

function 会員別_日付の文字_(value) {
  var d = 会員別_日付にする_(value);
  if (!d) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

// 測定履歴を読み、お名前ごとにまとめて古い順に並べる
function 会員別_測定記録を集める_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(MEASUREMENTS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var 見出し = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var 列 = {};
  見出し.forEach(function (name, i) { 列[String(name).trim()] = i; });

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var まとめ = {};
  values.forEach(function (row) {
    var 名前 = 会員別_名前をそろえる_(row[列['顧客名']]);
    if (!名前) return;
    if (!まとめ[名前]) まとめ[名前] = [];
    まとめ[名前].push({
      名前: String(row[列['顧客名']] || '').trim(),
      会員番号: String(row[列['会員番号']] || '').trim(),
      測定日: row[列['測定日']],
      ウエスト: row[列['ウエスト(cm)']],
      ヒップ: row[列['ヒップ(cm)']],
      太もも右: row[列['太もも右(cm)']],
      太もも左: row[列['太もも左(cm)']],
      メモ: String(row[列['スタッフメモ']] || '').trim(),
    });
  });

  Object.keys(まとめ).forEach(function (名前) {
    まとめ[名前].sort(function (a, b) {
      var x = 会員別_日付にする_(a.測定日), y = 会員別_日付にする_(b.測定日);
      return (x ? x.getTime() : 0) - (y ? y.getTime() : 0);
    });
  });
  return まとめ;
}




// 2つのファイルの中身を並べて確かめる（読むだけ・何も変えない）
var 会員別_ファイルA = '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E';  // 設定が指しているファイル
var 会員別_ファイルB = '1oDNTqlvKv1rGOGXIpnzPlegpFDeQ0WHGRLuY3ZAnZYc';  // コードに直接書かれたファイル

function ファイルを2つ調べる() {
  var props = PropertiesService.getScriptProperties();
  Logger.log('ACTIVE_SPREADSHEET_ID = ' + props.getProperty('ACTIVE_SPREADSHEET_ID'));
  Logger.log('SPREADSHEET_ID(プロパティ) = ' + props.getProperty('SPREADSHEET_ID'));
  Logger.log('SPREADSHEET_ID(コード) = ' + SPREADSHEET_ID);
  Logger.log('いま使われるファイル = ' + getSpreadsheet_().getId());
  Logger.log('');
  [['A（設定の指す先）', 会員別_ファイルA], ['B（コードの既定）', 会員別_ファイルB]].forEach(function (pair) {
    調べる_(pair[0], pair[1]);
  });
}

function 調べる_(ラベル, id) {
  var ss = null;
  try { ss = SpreadsheetApp.openById(id); }
  catch (error) { Logger.log('■ ' + ラベル + ' … 開けません（権限なし）'); return; }
  Logger.log('■ ' + ラベル + '  ' + ss.getName());
  Logger.log('   ' + ss.getUrl());
  ss.getSheets().forEach(function (sh) {
    var 行 = Math.max(0, sh.getLastRow() - 1);
    var 見出し = sh.getLastRow() > 0 && sh.getLastColumn() > 0
      ? sh.getRange(1, 1, 1, Math.min(sh.getLastColumn(), 8)).getValues()[0].join(',')
      : '';
    Logger.log('   ' + sh.getName() + ' … ' + 行 + '行  [' + 見出し.slice(0, 90) + ']');
  });
  Logger.log('');
}

// 回答一覧を読み、お名前ごと・アンケートごとにまとめる
function 会員別_回答を集める_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var 見出し = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var 列 = {};
  見出し.forEach(function (name, i) { 列[String(name).trim()] = i; });

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var まとめ = {};
  values.forEach(function (row) {
    var 名前 = 会員別_名前をそろえる_(row[列['お名前']]);
    if (!名前) return;
    var アンケート名 = String(row[列['アンケート名']] || '').trim() || 'アンケート';
    if (!まとめ[名前]) まとめ[名前] = {};
    if (!まとめ[名前][アンケート名]) まとめ[名前][アンケート名] = [];
    まとめ[名前][アンケート名].push({
      送信日時: row[列['送信日時']],
      対応状況: String(row[列['対応状況']] || '').trim(),
      管理メモ: String(row[列['管理メモ']] || '').trim(),
      答え: 会員別_答えを開く_(row[列['回答JSON']]),
    });
  });

  Object.keys(まとめ).forEach(function (名前) {
    Object.keys(まとめ[名前]).forEach(function (ア) {
      まとめ[名前][ア].sort(function (a, b) {
        var x = 会員別_日付にする_(a.送信日時), y = 会員別_日付にする_(b.送信日時);
        return (x ? x.getTime() : 0) - (y ? y.getTime() : 0);
      });
    });
  });
  return まとめ;
}

// 回答JSONを { 質問の見出し: 答え } の形にする（写真の回答は除く）
function 会員別_答えを開く_(回答JSON) {
  var out = {};
  var answers = null;
  try { answers = JSON.parse(String(回答JSON || '')); } catch (error) { return out; }
  if (!Array.isArray(answers)) return out;
  answers.forEach(function (a) {
    if (!a) return;
    var 見出し = String(a.label || a.questionId || '').trim();
    if (!見出し) return;
    // 計測写真などは、DriveのURLをそのまま欄に入れる（1枚ずつ改行で並べる）。
    if (Array.isArray(a.files)) {
      var urls = a.files.map(function (f) {
        return String((f && (f.url || f.previewUrl)) || '').trim();
      }).filter(String);
      if (urls.length) out[見出し] = urls.join('\n');
      return;
    }
    var v = a.value;
    if (Array.isArray(v)) v = v.join('、');
    v = String(v == null ? '' : v).trim();
    if (!v) return;
    out[見出し] = v;
  });
  return out;
}

// アンケートごとの質問の並びを、アンケート定義から取る
function 会員別_質問の並び_() {
  var 並び = {};
  try {
    getSurveys_().forEach(function (s) {
      var 名 = String(s.title || '').trim();
      if (!名) return;
      並び[名] = (s.questions || []).map(function (q) {
        return String(q.label || q.id || '').trim();
      }).filter(String);
    });
  } catch (error) { /* 定義が取れなければ、回答に出てきた順で作る */ }
  return 並び;
}


// ---------- どなたのシートを作るかの決め方 ----------
//
// ・回答や測定が届いた方は、どなたでも自動でシートを作る
// ・一度シートを削除した方が、また送信してきた場合も作り直す
//   ただし過去の記録は持ち込まず、戻ってこられた時点からの分だけを載せる
// ・除外リストに書いた方は、最初から作らない
//
// そのために「一度でも作った方」と「作り直した日時」を覚えている。

var 会員別_除外する方 = [
  // 例: 'テスト',
];
var 会員別_作った方の記録キー = 'MEMBER_DIGEST_KNOWN_NAMES';
var 会員別_再開日時の記録キー = 'MEMBER_DIGEST_RESTART_AT';

function 会員別_除外か_(名前) {
  var n = 会員別_名前をそろえる_(名前);
  return 会員別_除外する方.some(function (x) {
    return 会員別_名前をそろえる_(x) === n;
  });
}

function 会員別_記録を読む_(キー) {
  var raw = PropertiesService.getScriptProperties().getProperty(キー);
  try {
    var v = JSON.parse(raw || '{}');
    return (v && typeof v === 'object' && !Array.isArray(v)) ? v : {};
  } catch (error) { return {}; }
}

function 会員別_記録を書く_(キー, 値) {
  PropertiesService.getScriptProperties().setProperty(キー, JSON.stringify(値));
}

// 再開日時より前の記録は載せない（削除後に戻ってこられた方のため）
function 会員別_再開日時より後か_(日時, 再開) {
  if (!再開) return true;
  var d = 会員別_日付にする_(日時);
  if (!d) return true;
  return d.getTime() >= new Date(再開).getTime();
}

// ---------- 見た目の調整 ----------

// 使っていない行と列を削って、シートを中身の大きさに合わせる
function 会員別_切り詰める_(sheet, 行数, 列数) {
  var 余分な行 = sheet.getMaxRows() - 行数;
  if (余分な行 > 0) sheet.deleteRows(行数 + 1, 余分な行);
  var 余分な列 = sheet.getMaxColumns() - 列数;
  if (余分な列 > 0) sheet.deleteColumns(列数 + 1, 余分な列);
}

// 書き込む前に、必要な行数・列数を確保する。
// 前回の切り詰めでシートが小さくなっていると、そのままでは書き込めないため。
function 会員別_大きさを確保_(sheet, 行数, 列数) {
  var 足りない行 = 行数 - sheet.getMaxRows();
  if (足りない行 > 0) sheet.insertRowsAfter(sheet.getMaxRows(), 足りない行);
  var 足りない列 = 列数 - sheet.getMaxColumns();
  if (足りない列 > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), 足りない列);
}

// 中身がすべて見える幅にする。
// 自動調整（autoResizeColumns）は画面の描画に左右されて、
// シートによって効いたり効かなかったりする。そこで文字数から自分で計算する。
var 会員別_列幅の上限 = 420;

// 全角は2文字ぶんとして数える
function 会員別_文字幅_(value) {
  var s = String(value == null ? '' : value);
  var 幅 = 0;
  for (var i = 0; i < s.length; i += 1) {
    var c = s.charCodeAt(i);
    // 半角英数記号はおよそ1、それ以外（日本語など）は2として数える
    幅 += (c < 0x0250 || (c >= 0xFF61 && c <= 0xFF9F)) ? 1 : 2;
  }
  return 幅;
}

// rows は書き込んだ内容そのもの。改行があれば一番長い行で測る。
function 会員別_幅を整える_(sheet, rows, 列数) {
  for (var c = 0; c < 列数; c += 1) {
    var 最大 = 0;
    for (var r = 0; r < rows.length; r += 1) {
      var v = rows[r][c];
      if (v === '' || v == null) continue;
      String(v).split('\n').forEach(function (line) {
        var w = 会員別_文字幅_(line);
        if (w > 最大) 最大 = w;
      });
    }
    var 幅 = Math.min(会員別_列幅の上限, Math.max(80, 最大 * 8 + 24));
    sheet.setColumnWidth(c + 1, 幅);
    // 上限まで届いた列だけ折り返す
    sheet.getRange(1, c + 1, sheet.getMaxRows(), 1).setWrap(幅 >= 会員別_列幅の上限);
  }
}


// URLが入っているセルを、押せるリンクにする。
// 1つのセルに複数のURLを改行で並べると、そのままでは自動リンクにならない。
function 会員別_リンクにする_(sheet, rows, 幅) {
  for (var r = 0; r < rows.length; r += 1) {
    for (var c = 0; c < 幅; c += 1) {
      var v = String(rows[r][c] == null ? '' : rows[r][c]);
      if (v.indexOf('http') < 0) continue;
      var 行 = v.split('\n');
      var 作る = SpreadsheetApp.newRichTextValue().setText(v);
      var 位置 = 0;
      var 付けた = false;
      行.forEach(function (line, i) {
        var 開始 = 位置;
        var 終了 = 位置 + line.length;
        if (/^https?:\/\//.test(line.trim())) {
          作る.setLinkUrl(開始, 終了, line.trim());
          付けた = true;
        }
        位置 = 終了 + 1;   // 改行のぶん
      });
      if (付けた) sheet.getRange(r + 1, c + 1).setRichTextValue(作る.build());
    }
  }
}

// 55 を 55.0 のように、小数点第一位まで表示する
var 会員別_小数の書式 = '0.0';
var 会員別_増減の書式 = '+0.0;-0.0;0.0';

// ---------- ① 測定記録の会員別一覧 ----------

function 測定記録の会員別一覧を作る() {
  var ss = getSpreadsheet_();
  var 測定 = 会員別_測定記録を集める_();
  var 名前一覧 = Object.keys(測定).filter(function (n) { return !会員別_除外か_(n); }).sort();

  var 見出し = ['お名前', '会員番号', '測定日', '回数',
    'ウエスト', 'ヒップ', '太もも右', '太もも左',
    'ウエスト増減', 'ヒップ増減', '初回からのウエスト増減', 'スタッフメモ'];

  var rows = [];
  名前一覧.forEach(function (名前) {
    var 記録 = 測定[名前];
    var 初回 = 記録[0];
    記録.forEach(function (r, i) {
      var 前 = i > 0 ? 記録[i - 1] : null;
      var 差 = function (今, 前の) {
        if (今 === '' || 今 == null || 前の === '' || 前の == null) return '';
        var a = Number(今), b = Number(前の);
        if (isNaN(a) || isNaN(b)) return '';
        var d = Math.round((a - b) * 10) / 10;
        return d > 0 ? '+' + d : String(d);
      };
      rows.push([
        r.名前, r.会員番号, 会員別_日付の文字_(r.測定日), i + 1,
        r.ウエスト, r.ヒップ, r.太もも右, r.太もも左,
        前 ? 差(r.ウエスト, 前.ウエスト) : '',
        前 ? 差(r.ヒップ, 前.ヒップ) : '',
        i > 0 ? 差(r.ウエスト, 初回.ウエスト) : '',
        r.メモ,
      ]);
    });
  });

  var sheet = ss.getSheetByName(会員別_一覧シート名);
  if (!sheet) sheet = ss.insertSheet(会員別_一覧シート名);
  sheet.clear();
  会員別_大きさを確保_(sheet, rows.length + 1, 見出し.length);
  sheet.getRange(1, 1, 1, 見出し.length).setValues([見出し])
    .setFontWeight('bold').setBackground('#e8efe4');
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 見出し.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
  if (rows.length) {
    // ウエスト〜太もも左 は 55.0 の形、増減は +1.5 / -2.0 の形にする
    sheet.getRange(2, 5, rows.length, 4).setNumberFormat(会員別_小数の書式);
    sheet.getRange(2, 9, rows.length, 3).setNumberFormat(会員別_増減の書式);
  }
  会員別_切り詰める_(sheet, Math.max(rows.length + 1, 2), 見出し.length);
  会員別_幅を整える_(sheet, [見出し].concat(rows), 見出し.length);

  Logger.log('「' + 会員別_一覧シート名 + '」を作り直しました');
  Logger.log('  会員 ' + 名前一覧.length + '名 / 測定 ' + rows.length + '件');
  return { 会員数: 名前一覧.length, 件数: rows.length };
}

// ---------- ② 会員ごとのシートを持つ別ファイル ----------

function 会員ごとのファイルを作る() {
  var 測定 = 会員別_測定記録を集める_();
  var 回答 = 会員別_回答を集める_();
  var 質問の並び = 会員別_質問の並び_();

  var 名前一覧 = {};
  Object.keys(測定).forEach(function (n) { 名前一覧[n] = true; });
  Object.keys(回答).forEach(function (n) { 名前一覧[n] = true; });
  var 候補 = Object.keys(名前一覧).sort();
  if (!候補.length) { Logger.log('対象の会員がいません'); return; }

  var book = 会員別_ファイルを開く_();
  var 作った方 = 会員別_記録を読む_(会員別_作った方の記録キー);
  var 再開日時 = 会員別_記録を読む_(会員別_再開日時の記録キー);
  var いま = new Date().toISOString();
  var 新しい方 = [];
  var 戻られた方 = [];

  var 名前 = 候補.filter(function (n) { return !会員別_除外か_(n); });
  if (!名前.length) { Logger.log('作る対象がありません'); return; }

  // シートが無い方は、これから作る。
  // 以前に作って削除された方なら、ここを起点にして過去の記録は載せない。
  名前.forEach(function (n) {
    var かぎ = 会員別_名前をそろえる_(n);
    if (book.getSheetByName(n)) return;
    if (作った方[かぎ]) {
      再開日時[かぎ] = いま;
      戻られた方.push(n);
    } else {
      新しい方.push(n);
    }
  });

  // 出したいアンケートの順番。ここに無いものは、そのあとに続けて出す。
  var 先に出す = ['施術後アンケート', '計測時アンケート'];

  var 目次 = book.getSheetByName('目次') || book.insertSheet('目次', 0);
  目次.clear();
  var 目次行 = [['お名前', '測定回数', '初回', '最新', 'ウエストの変化', '施術後', '計測時']];
  var 使っているシート名 = {};   // 今回作った（＝いま在籍している）方のシート名

  名前.forEach(function (キー) {
    var 起点 = 再開日時[会員別_名前をそろえる_(キー)] || '';
    var m = (測定[キー] || []).filter(function (r) {
      return 会員別_再開日時より後か_(r.測定日, 起点);
    });
    var ア = {};
    Object.keys(回答[キー] || {}).forEach(function (アンケート名) {
      var 一覧 = (回答[キー][アンケート名] || []).filter(function (r) {
        return 会員別_再開日時より後か_(r.送信日時, 起点);
      });
      if (一覧.length) ア[アンケート名] = 一覧;
    });
    var 表示名 = (m[0] && m[0].名前) || キー;
    var 初回 = m[0];
    var 最新 = m[m.length - 1];
    目次行.push([
      表示名, m.length,
      初回 ? 会員別_日付の文字_(初回.測定日) : '',
      最新 ? 会員別_日付の文字_(最新.測定日) : '',
      初回 && 最新 ? 会員別_増減_(最新.ウエスト, 初回.ウエスト) : '',
      (ア['施術後アンケート'] || []).length,
      (ア['計測時アンケート'] || []).length,
    ]);

    var シート名 = 会員別_シート名にする_(表示名, book);
    使っているシート名[シート名] = true;
    var sheet = book.getSheetByName(シート名) || book.insertSheet(シート名);
    sheet.clear();

    var rows = [];
    var 見出し行 = [];   // 太字にする行番号

    // ■ 測定記録
    見出し行.push(rows.length + 1);
    rows.push(['■ 測定記録']);
    rows.push(['測定日', '回数', 'ウエスト', 'ヒップ', '太もも右', '太もも左',
      'ウエスト増減', 'ヒップ増減', '初回からのウエスト増減', 'スタッフメモ']);
    var 測定の開始行 = rows.length + 1;
    if (m.length) {
      m.forEach(function (r, i) {
        var 前 = i > 0 ? m[i - 1] : null;
        rows.push([
          会員別_日付の文字_(r.測定日), i + 1,
          r.ウエスト, r.ヒップ, r.太もも右, r.太もも左,
          前 ? 会員別_増減_(r.ウエスト, 前.ウエスト) : '',
          前 ? 会員別_増減_(r.ヒップ, 前.ヒップ) : '',
          i > 0 ? 会員別_増減_(r.ウエスト, 初回.ウエスト) : '',
          r.メモ,
        ]);
      });
    } else {
      rows.push(['（測定記録はありません）']);
    }

    // ■ アンケートごと（質問を列にする）
    var 出す = 先に出す.slice();
    Object.keys(ア).forEach(function (n) { if (出す.indexOf(n) < 0) 出す.push(n); });

    出す.forEach(function (アンケート名) {
      var 一覧 = ア[アンケート名] || [];
      rows.push([]);
      見出し行.push(rows.length + 1);
      rows.push(['■ ' + アンケート名]);

      // 質問の並びは、定義があればそれに従う。無ければ回答に出てきた順。
      var 質問 = (質問の並び[アンケート名] || []).slice();
      一覧.forEach(function (r) {
        Object.keys(r.答え).forEach(function (q) {
          if (質問.indexOf(q) < 0) 質問.push(q);
        });
      });

      rows.push(['提出日'].concat(質問).concat(['対応状況', '管理メモ']));
      if (一覧.length) {
        一覧.forEach(function (r) {
          var line = [会員別_日付の文字_(r.送信日時)];
          質問.forEach(function (q) { line.push(r.答え[q] || ''); });
          line.push(r.対応状況, r.管理メモ);
          rows.push(line);
        });
      } else {
        rows.push(['（まだ回答はありません）']);
      }
    });

    var 幅 = rows.reduce(function (max, r) { return Math.max(max, r.length); }, 1);
    rows = rows.map(function (r) {
      var c = r.slice();
      while (c.length < 幅) c.push('');
      return c;
    });
    会員別_大きさを確保_(sheet, rows.length, 幅);
    sheet.getRange(1, 1, rows.length, 幅).setValues(rows);
    sheet.getRange(1, 1, rows.length, 幅).setVerticalAlignment('top');
    見出し行.forEach(function (n) {
      sheet.getRange(n, 1, 1, 幅).setFontWeight('bold').setBackground('#e8efe4');
      sheet.getRange(n + 1, 1, 1, 幅).setFontWeight('bold');
    });
    // 測定記録の数値だけ、小数点第一位まで表示する
    if (m.length) {
      sheet.getRange(測定の開始行, 3, m.length, 4).setNumberFormat(会員別_小数の書式);
      sheet.getRange(測定の開始行, 7, m.length, 3).setNumberFormat(会員別_増減の書式);
    }
    会員別_切り詰める_(sheet, rows.length, 幅);
    会員別_幅を整える_(sheet, rows, 幅);
    会員別_リンクにする_(sheet, rows, 幅);
  });

  会員別_大きさを確保_(目次, 目次行.length, 7);
  目次.getRange(1, 1, 目次行.length, 7).setValues(目次行);
  目次.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#e8efe4');
  目次.setFrozenRows(1);
  if (目次行.length > 1) {
    目次.getRange(2, 5, 目次行.length - 1, 1).setNumberFormat(会員別_増減の書式);
  }
  会員別_切り詰める_(目次, 目次行.length, 7);
  会員別_幅を整える_(目次, 目次行, 7);

  // 管理アプリで削除された方（＝もうどこにも記録が無い方）のシートを外す。
  // 「作った方」の記録は残すので、また登録されたときは再開日時からの記録だけを載せる。
  var 外したシート = [];
  book.getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (n === '目次') return;
    if (使っているシート名[n]) return;
    if (book.getSheets().length <= 1) return;
    book.deleteSheet(sh);
    外したシート.push(n);
  });

  名前.forEach(function (n) { 作った方[会員別_名前をそろえる_(n)] = true; });
  会員別_記録を書く_(会員別_作った方の記録キー, 作った方);
  会員別_記録を書く_(会員別_再開日時の記録キー, 再開日時);

  Logger.log('会員ごとのファイルを作り直しました（測定記録＋アンケートごと）');
  Logger.log('  会員 ' + 名前.length + '名');
  if (外したシート.length) Logger.log('  外したシート（削除された方）: ' + 外したシート.join('・'));
  if (新しい方.length) Logger.log('  新しく作った方: ' + 新しい方.join('・'));
  if (戻られた方.length) {
    Logger.log('  作り直した方（ここからの記録のみ）: ' + 戻られた方.join('・'));
  }
  Logger.log('  ' + book.getUrl());
  return book.getUrl();
}

// 増減を「+1.5」「-2」の形にする
function 会員別_増減_(今, 前) {
  if (今 === '' || 今 == null || 前 === '' || 前 == null) return '';
  var a = Number(今), b = Number(前);
  if (isNaN(a) || isNaN(b)) return '';
  // 文字ではなく数値で返す。「+」「-」は表示の書式で付ける。
  return Math.round((a - b) * 10) / 10;
}

// シート名に使えない文字を避けつつ、重複しない名前にする
function 会員別_シート名にする_(表示名, book) {
  var base = String(表示名 || '会員').replace(/[\[\]\*\/\\\?:]/g, '');
  base = base.slice(0, 90) || '会員';
  if (base === '目次') base = '目次(会員)';
  return base;
}

// ---------- ③ 競合シートの確認 ----------

function 競合シートを比べる() {
  var ss = getSpreadsheet_();
  ss.getSheets().forEach(function (sheet) {
    var name = sheet.getName();
    var m = name.match(/^(.+)_conflict\d+$/);
    if (!m) return;
    会員別_比較_(ss, m[1], name);
  });
}

function 会員別_行のかぎ_(row) {
  return row.map(function (v) {
    return v instanceof Date ? v.toISOString() : String(v == null ? '' : v);
  }).join('');
}

function 会員別_比較_(ss, 本体名, 競合名) {
  var a = ss.getSheetByName(本体名);
  var b = ss.getSheetByName(競合名);
  if (!a || !b) { Logger.log('― ' + 本体名 + ': 片方が見つかりません'); return; }

  var 読む = function (sh) {
    if (sh.getLastRow() < 2) return { 見出し: [], 行: [] };
    var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    return {
      見出し: v[0],
      行: v.slice(1).filter(function (r) { return r.join('').trim() !== ''; }),
    };
  };
  var A = 読む(a), B = 読む(b);

  var setA = {}, setB = {};
  A.行.forEach(function (r) { setA[会員別_行のかぎ_(r)] = true; });
  B.行.forEach(function (r) { setB[会員別_行のかぎ_(r)] = true; });
  var 競合だけ = B.行.filter(function (r) { return !setA[会員別_行のかぎ_(r)]; });
  var 本体だけ = A.行.filter(function (r) { return !setB[会員別_行のかぎ_(r)]; });

  Logger.log('― ' + 本体名 + ' ↔ ' + 競合名);
  Logger.log('   行数: 本体 ' + A.行.length + ' / 競合 ' + B.行.length);
  Logger.log('   見出しが同じか: ' + (A.見出し.join('|') === B.見出し.join('|') ? 'はい' : 'いいえ'));
  Logger.log('   競合にしか無い行: ' + 競合だけ.length + '  ← これが0なら競合側は消して差し支えない');
  Logger.log('   本体にしか無い行: ' + 本体だけ.length);
  競合だけ.slice(0, 5).forEach(function (r) {
    Logger.log('     競合のみ例: ' + String(r[0]).slice(0, 22) + ' | ' + String(r[1]).slice(0, 22));
  });
}

// ---------- 毎日ひとりでに作り直す ----------
// 手で実行しなくても、顧客ごとのシートが最新に保たれるようにする。

var 会員別_自動実行キー = 'MEMBER_DIGEST_TRIGGER_ID';

function 毎日の自動更新を入れる() {
  毎日の自動更新をやめる();
  var t = ScriptApp.newTrigger('会員別まとめを毎日作り直す')
    .timeBased().everyDays(1).atHour(4).create();
  PropertiesService.getScriptProperties().setProperty(会員別_自動実行キー, t.getUniqueId());
  Logger.log('毎日 午前4時ごろに、自動で作り直すようにしました');
}

function 毎日の自動更新をやめる() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(会員別_自動実行キー);
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getUniqueId() === id || t.getHandlerFunction() === '会員別まとめを毎日作り直す') {
      ScriptApp.deleteTrigger(t);
    }
  });
  props.deleteProperty(会員別_自動実行キー);
  Logger.log('自動更新の設定を外しました');
}

function 会員別まとめを毎日作り直す() {
  try {
    測定記録の会員別一覧を作る();
    会員ごとのファイルを作る();
  } catch (error) {
    Logger.log('会員別まとめの自動更新に失敗: ' + error.message);
  }
}

// ---------- 継続していない方の記録を消す ----------

// 名簿（会員別_対象の方）に無い方の記録を、まとめて削除する。
// 控えを取ってから消すので、間違えても戻せる。
function 継続していない方の記録を消す() {
  var live = getSpreadsheet_();

  var 控え = DriveApp.getFileById(live.getId()).makeCopy(
    '【控え】' + live.getName() + ' ' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm')
  );
  Logger.log('控えを作りました: ' + 控え.getUrl());
  Logger.log('除外する方: ' + (会員別_除外する方.length ? 会員別_除外する方.join('・') : 'なし'));
  Logger.log('');

  var 合計 = 0;
  live.getSheets().forEach(function (sh) {
    var 行数 = sh.getLastRow();
    var 列数 = sh.getLastColumn();
    if (行数 < 2 || 列数 < 1) return;
    var 見出し = sh.getRange(1, 1, 1, 列数).getValues()[0];
    var 名前列 = -1;
    ['顧客名', 'お名前', '氏名'].forEach(function (l) {
      if (名前列 < 0) { var i = 見出し.indexOf(l); if (i >= 0) 名前列 = i; }
    });
    if (名前列 < 0) return;

    var values = sh.getRange(2, 名前列 + 1, 行数 - 1, 1).getValues();
    var 消す = [];
    var 消した名前 = {};
    values.forEach(function (r, i) {
      var 名前 = String(r[0] || '').trim();
      if (!名前) return;
      if (!会員別_除外か_(名前)) return;
      消す.push(i + 2);
      消した名前[名前] = (消した名前[名前] || 0) + 1;
    });
    if (!消す.length) return;

    消す.sort(function (a, b) { return b - a; });   // 下から消す
    消す.forEach(function (r) { sh.deleteRow(r); });
    合計 += 消す.length;
    Logger.log(sh.getName() + ' … ' + 消す.length + '行を削除');
    Object.keys(消した名前).sort().forEach(function (n) {
      Logger.log('   ' + n + ' … ' + 消した名前[n] + '行');
    });
  });

  Logger.log('');
  Logger.log('合計 ' + 合計 + '行を削除しました');

  // 会員ごとのファイルからも、対象外の方のシートを外す
  try {
    var book = 会員別_ファイルを開く_();
    var 外した = [];
    book.getSheets().forEach(function (sh) {
      var n = sh.getName();
      if (n === '目次') return;
      if (!会員別_除外か_(n)) return;
      if (book.getSheets().length <= 1) return;
      book.deleteSheet(sh);
      外した.push(n);
    });
    Logger.log('会員ごとのファイルから外したシート: ' + (外した.length ? 外した.join(' / ') : 'なし'));
  } catch (error) {
    Logger.log('会員ごとのファイルを開けませんでした: ' + error.message);
  }

  // 一覧を作り直す
  var 結果 = 測定記録の会員別一覧を作る();
  Logger.log('測定記録_会員別: ' + 結果.会員数 + '名 / ' + 結果.件数 + '件');
}

// ---------- お名前の付け替え ----------
//
// ご結婚などでお名前が変わったとき、記録を残したまま名前だけ差し替える。
// 実行前に控えを取る。同じ日の記録がすでに新しい名前で入っていれば、
// 二重にならないよう古いほうを消す。

var 付け替え_古い名前 = '山谷';
var 付け替え_新しい名前 = '山谷未央';

function お名前を付け替える() {
  var live = getSpreadsheet_();
  var 控え = DriveApp.getFileById(live.getId()).makeCopy(
    '【控え】' + live.getName() + ' ' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm')
  );
  Logger.log('控えを作りました: ' + 控え.getUrl());
  Logger.log('「' + 付け替え_古い名前 + '」→「' + 付け替え_新しい名前 + '」');
  Logger.log('');

  var 古 = 会員別_名前をそろえる_(付け替え_古い名前);
  var 新 = 会員別_名前をそろえる_(付け替え_新しい名前);
  var 付け替えた = 0;
  var 消した = 0;

  live.getSheets().forEach(function (sh) {
    var 行数 = sh.getLastRow();
    var 列数 = sh.getLastColumn();
    if (行数 < 2 || 列数 < 1) return;
    var 見出し = sh.getRange(1, 1, 1, 列数).getValues()[0];
    var 名前列 = -1;
    ['顧客名', 'お名前', '氏名'].forEach(function (l) {
      if (名前列 < 0) { var i = 見出し.indexOf(l); if (i >= 0) 名前列 = i; }
    });
    if (名前列 < 0) return;
    var 日付列 = 見出し.indexOf('測定日');
    var ID列 = 見出し.indexOf('測定ID');

    var values = sh.getRange(2, 1, 行数 - 1, 列数).getValues();

    // すでに新しい名前で入っている「日付」を控える（測定履歴の二重防止）
    var すでにある = {};
    if (日付列 >= 0) {
      values.forEach(function (r) {
        if (会員別_名前をそろえる_(r[名前列]) !== 新) return;
        var d = 会員別_日付にする_(r[日付列]);
        if (d) すでにある[会員別_日付の文字_(d)] = true;
      });
    }

    var 消す = [];
    values.forEach(function (r, i) {
      if (会員別_名前をそろえる_(r[名前列]) !== 古) return;
      var 行番号 = i + 2;
      if (日付列 >= 0) {
        var d = 会員別_日付にする_(r[日付列]);
        var かぎ = d ? 会員別_日付の文字_(d) : '';
        if (かぎ && すでにある[かぎ]) {
          消す.push(行番号);   // 新しい名前で同じ日の記録がある＝重複
          return;
        }
        if (かぎ) すでにある[かぎ] = true;
      }
      sh.getRange(行番号, 名前列 + 1).setValue(付け替え_新しい名前);
      // 取り込みのIDは「お名前＋測定日」から作っているので、付け替えに合わせて振り直す
      if (ID列 >= 0 && 日付列 >= 0) {
        var 元ID = String(r[ID列] || '');
        var d2 = 会員別_日付にする_(r[日付列]);
        if (d2 && 元ID.indexOf('imp-') === 0) {
          sh.getRange(行番号, ID列 + 1).setValue(取込_測定IDを作る_(付け替え_新しい名前, d2));
        }
      }
      付け替えた += 1;
    });

    消す.sort(function (a, b) { return b - a; });
    消す.forEach(function (r) { sh.deleteRow(r); });
    消した += 消す.length;

    if (付け替えた || 消す.length) {
      Logger.log(sh.getName() + ' … 付け替え ' + 付け替えた + '行 / 重複削除 ' + 消す.length + '行');
    }
    付け替えた = 0;
  });

  // 会員別まとめのシート名も付け替える
  try {
    var book = 会員別_ファイルを開く_();
    var 旧シート = book.getSheetByName(付け替え_古い名前);
    if (旧シート) {
      if (book.getSheetByName(付け替え_新しい名前)) {
        book.deleteSheet(旧シート);
        Logger.log('会員別まとめ: 新しい名前のシートがあったので、古いほうを削除しました');
      } else {
        旧シート.setName(付け替え_新しい名前);
        Logger.log('会員別まとめ: シート名を付け替えました');
      }
    }
  } catch (error) {
    Logger.log('会員別まとめを開けませんでした: ' + error.message);
  }

  Logger.log('');
  Logger.log('重複として削除した行の合計: ' + 消した);
  var 結果 = 測定記録の会員別一覧を作る();
  Logger.log('測定記録_会員別: ' + 結果.会員数 + '名 / ' + 結果.件数 + '件');
}
