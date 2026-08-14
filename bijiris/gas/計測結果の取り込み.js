// 「SNS投稿」ファイルの『計測結果』シートを、ビジリスの『測定履歴』へ取り込む。
// 元の表は 1人につき4行（計測日／ウエスト／ヒップ／太もも）で、日付が横に並ぶ形。
// 測定履歴は 1行1計測なので、縦に開いて入れ直す。
//
//   計測結果を取り込む()  … 控えを取る → 釼持裕基の記録を消す → 取り込む → 一覧を作り直す
//   取り込む前に確認する()  … 何が入って何が消えるかを、書き込まずに表示するだけ
//
// 同じ人・同じ計測日は同じ測定IDにしてあるので、何度実行しても二重に増えない。

var 取込元ファイルID = '1-dBYtQmUuKlTWV-Db10Bv7---KfTTsz7mPObf2dst84';
var 取込元シート名 = '計測結果';
var 取込_除外する名前 = ['釼持裕基'];

// ---------- 元の表を読む ----------

function 取込_元データを読む_() {
  var ss = SpreadsheetApp.openById(取込元ファイルID);
  var sheet = ss.getSheetByName(取込元シート名);
  if (!sheet) throw new Error('『' + 取込元シート名 + '』シートが見つかりません。');

  var 行数 = sheet.getLastRow();
  var 列数 = sheet.getLastColumn();
  if (行数 < 2 || 列数 < 3) return [];
  var v = sheet.getRange(1, 1, 行数, 列数).getValues();

  var 結果 = [];
  var 今の名前 = '';
  var 日付 = {};   // 列番号 → 計測日
  var 値 = {};     // 項目 → { 列番号: 値 }

  var 確定 = function () {
    if (!今の名前) return;
    Object.keys(日付).forEach(function (c) {
      var col = Number(c);
      var 測定日 = 日付[col];
      if (!測定日) return;
      var 取る = function (項目, 何番目) {
        var m = 値[項目];
        var 数 = m ? m[col] : null;
        if (!数 || !数.length) return '';
        var x = 数[何番目 || 0];
        return (x === '' || x == null) ? '' : x;
      };
      var ウエスト = 取る('ウエスト');
      var ヒップ = 取る('ヒップ');
      // 「太もも」に2つ書かれていれば、1つ目を右・2つ目を左とする。
      var 太もも右 = 取る('太もも右') !== '' ? 取る('太もも右') : 取る('太もも', 0);
      var 太もも左 = 取る('太もも左') !== '' ? 取る('太もも左') : 取る('太もも', 1);
      if (ウエスト === '' && ヒップ === '' && 太もも右 === '' && 太もも左 === '') return;
      結果.push({
        名前: 今の名前,
        測定日: 測定日,
        ウエスト: ウエスト,
        ヒップ: ヒップ,
        太もも右: 太もも右,
        太もも左: 太もも左,
      });
    });
    日付 = {};
    値 = {};
  };

  for (var r = 1; r < 行数; r += 1) {
    var 名前欄 = String(v[r][0] == null ? '' : v[r][0]).trim();
    if (名前欄) {
      確定();
      今の名前 = 名前欄;
    }
    var 項目 = String(v[r][1] == null ? '' : v[r][1]).trim();
    if (!項目) continue;

    if (項目 === '計測日') {
      for (var c = 2; c < 列数; c += 1) {
        var d = 取込_日付にする_(v[r][c]);
        if (d) 日付[c] = d;
      }
    } else {
      var キー = 項目.replace(/[\s　()（）cm]/g, '');
      if (!値[キー]) 値[キー] = {};
      for (var c2 = 2; c2 < 列数; c2 += 1) {
        var x = v[r][c2];
        if (x === '' || x == null) continue;
        // 「53/55」「54 55」のように左右2つ書かれているセルがある。
        // 数字をつなげてしまうと 5355 のような値になるので、1つずつ取り出す。
        var 数 = 取込_数字を取り出す_(x);
        if (!数.length) continue;
        値[キー][c2] = 数;
      }
    }
  }
  確定();
  return 結果;
}


// セルの中の数字をすべて取り出す。「53/55」→ [53, 55]、「76」→ [76]
function 取込_数字を取り出す_(value) {
  if (typeof value === 'number') return isNaN(value) ? [] : [value];
  var s = String(value == null ? '' : value);
  var m = s.match(/\d+(?:\.\d+)?/g);
  if (!m) return [];
  return m.map(Number).filter(function (n) { return !isNaN(n); });
}

function 取込_日付にする_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var s = String(value == null ? '' : value).trim();
  if (!s) return null;
  var m = s.match(/(\d{4})[\/\-年.](\d{1,2})[\/\-月.](\d{1,2})/);
  if (!m) return null;
  var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return isNaN(d.getTime()) ? null : d;
}

function 取込_日付の文字_(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
}

// 同じ人・同じ日なら必ず同じIDになるようにする（何度実行しても増えない）
function 取込_測定IDを作る_(名前, 測定日) {
  var seed = 'bijiris-import|' + 名前 + '|' + 取込_日付の文字_(測定日);
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8);
  return 'imp-' + Utilities.base64EncodeWebSafe(digest).replace(/[^A-Za-z0-9]/g, '').slice(0, 20);
}

function 取込_除外か_(名前) {
  var n = String(名前 || '').replace(/[\s　]+/g, '');
  return 取込_除外する名前.some(function (x) {
    return n === String(x).replace(/[\s　]+/g, '');
  });
}

// ---------- 確認だけ（書き込まない） ----------

function 取り込む前に確認する() {
  var 元 = 取込_元データを読む_();
  var 対象 = 元.filter(function (r) { return !取込_除外か_(r.名前); });
  var 除外 = 元.length - 対象.length;

  var 人 = {};
  対象.forEach(function (r) { 人[r.名前] = (人[r.名前] || 0) + 1; });

  Logger.log('■ 取り込む内容（まだ何も書き込んでいません）');
  Logger.log('   元の計測: ' + 元.length + '件');
  Logger.log('   取り込む: ' + 対象.length + '件 / ' + Object.keys(人).length + '名');
  Logger.log('   除外(' + 取込_除外する名前.join('・') + '): ' + 除外 + '件');
  Object.keys(人).sort().forEach(function (名前) {
    Logger.log('     ' + 名前 + ' … ' + 人[名前] + '回');
  });

  Logger.log('');
  Logger.log('■ 消える予定の記録（' + 取込_除外する名前.join('・') + '）');
  取込_消す対象を数える_().forEach(function (line) { Logger.log('   ' + line); });
}

// 各シートで「その名前」の行が何行あるかを数える
function 取込_消す対象を数える_() {
  var out = [];
  [getSpreadsheet_(), 取込_もう一方のファイル_()].forEach(function (ss) {
    if (!ss) return;
    out.push(ss.getName() + '  ' + ss.getId());
    ss.getSheets().forEach(function (sh) {
      var n = 取込_名前の行を集める_(sh).length;
      if (n) out.push('   ' + sh.getName() + ' … ' + n + '行');
    });
  });
  return out;
}

function 取込_もう一方のファイル_() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return ss.getId() === getSpreadsheet_().getId() ? null : ss;
  } catch (error) {
    return null;
  }
}

// お名前の列を見つけて、除外対象の行番号を返す
function 取込_名前の行を集める_(sheet) {
  var 行数 = sheet.getLastRow();
  var 列数 = sheet.getLastColumn();
  if (行数 < 2 || 列数 < 1) return [];
  var 見出し = sheet.getRange(1, 1, 1, 列数).getValues()[0];
  var 名前列 = -1;
  ['顧客名', 'お名前', '氏名'].forEach(function (label) {
    if (名前列 >= 0) return;
    var i = 見出し.indexOf(label);
    if (i >= 0) 名前列 = i;
  });
  if (名前列 < 0) return [];
  var values = sheet.getRange(2, 名前列 + 1, 行数 - 1, 1).getValues();
  var rows = [];
  values.forEach(function (r, i) {
    if (取込_除外か_(r[0])) rows.push(i + 2);
  });
  return rows;
}

// ---------- 本番（控えを取ってから書き込む） ----------

function 計測結果を取り込む() {
  var live = getSpreadsheet_();

  // 1. 控えを取る
  var 控え = DriveApp.getFileById(live.getId()).makeCopy(
    '【控え】' + live.getName() + ' ' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm')
  );
  Logger.log('控えを作りました: ' + 控え.getUrl());

  // 2. 除外対象の記録を消す
  var 消した = 0;
  [live, 取込_もう一方のファイル_()].forEach(function (ss) {
    if (!ss) return;
    ss.getSheets().forEach(function (sh) {
      var rows = 取込_名前の行を集める_(sh);
      rows.sort(function (a, b) { return b - a; });   // 下から消さないと行番号がずれる
      rows.forEach(function (r) { sh.deleteRow(r); });
      if (rows.length) {
        Logger.log('  ' + ss.getName() + ' / ' + sh.getName() + ' … ' + rows.length + '行を削除');
        消した += rows.length;
      }
    });
  });
  Logger.log('合計 ' + 消した + '行を削除しました');

  // 3. 計測結果を取り込む
  var 元 = 取込_元データを読む_().filter(function (r) { return !取込_除外か_(r.名前); });
  var sheet = live.getSheetByName(MEASUREMENTS_SHEET_NAME);
  if (!sheet) {
    sheet = live.insertSheet(MEASUREMENTS_SHEET_NAME);
    sheet.getRange(1, 1, 1, MEASUREMENT_HEADERS.length).setValues([MEASUREMENT_HEADERS]);
  }

  // すでに入っている測定IDを控えておき、同じものは上書きする
  var 既存 = {};
  if (sheet.getLastRow() >= 2) {
    var ids = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
    ids.forEach(function (r, i) {
      var id = String(r[0] || '').trim();
      if (id) 既存[id] = i + 2;
    });
  }

  var いま = new Date().toISOString();
  var 追加 = 0, 更新 = 0;
  元.forEach(function (r) {
    var id = 取込_測定IDを作る_(r.名前, r.測定日);
    var whr = (r.ウエスト !== '' && r.ヒップ !== '' && Number(r.ヒップ))
      ? Math.round((Number(r.ウエスト) / Number(r.ヒップ)) * 100) / 100
      : '';
    var row = [
      いま, いま, id, r.名前, '',
      取込_日付の文字_(r.測定日),
      r.ウエスト, r.ヒップ, r.太もも右, r.太もも左,
      whr, '',
    ];
    if (既存[id]) {
      sheet.getRange(既存[id], 1, 1, MEASUREMENT_HEADERS.length).setValues([row]);
      更新 += 1;
    } else {
      sheet.appendRow(row);
      追加 += 1;
    }
  });
  Logger.log('測定履歴に ' + 追加 + '件を追加、' + 更新 + '件を更新しました');

  // 4. 会員別の一覧を作り直す
  var 結果 = 測定記録の会員別一覧を作る();
  Logger.log('会員別の一覧: ' + 結果.会員数 + '名 / ' + 結果.件数 + '件');
  Logger.log('');
  Logger.log('※ 元の表は「太もも」が1つだけなので、太もも右に入れ、太もも左は空にしています。');
  return { 控え: 控え.getUrl(), 削除: 消した, 追加: 追加, 更新: 更新 };
}

// ---------- 整理の下見（読むだけ） ----------

function 整理の下見() {
  var live = getSpreadsheet_();

  Logger.log('■ 分析シート2つの列を並べる（移し替えの対応を決めるため）');
  ['回数券分析結果', '回数券終了分析'].forEach(function (名前) {
    var sh = live.getSheetByName(名前);
    if (!sh) { Logger.log('   ' + 名前 + ' … ありません'); return; }
    var 列 = sh.getLastColumn();
    var 見出し = 列 ? sh.getRange(1, 1, 1, 列).getValues()[0] : [];
    Logger.log('   ' + 名前 + ' … ' + Math.max(0, sh.getLastRow() - 1) + '行 / ' + 列 + '列');
    見出し.forEach(function (h, i) {
      Logger.log('      ' + (i + 1) + '. ' + h);
    });
  });

  Logger.log('');
  Logger.log('■ いまアプリが使っているアンケート（この名前のシートは消しても復活します）');
  try {
    getSurveys_().forEach(function (s) {
      Logger.log('   ' + s.title + '  (' + s.status + ')');
    });
  } catch (error) {
    Logger.log('   取得できませんでした: ' + error.message);
  }

  Logger.log('');
  Logger.log('■ 空のシート（消してよい候補）');
  [live, 取込_もう一方のファイル_()].forEach(function (ss) {
    if (!ss) return;
    Logger.log('   ' + ss.getName() + '  ' + ss.getId());
    ss.getSheets().forEach(function (sh) {
      var 行 = Math.max(0, sh.getLastRow() - 1);
      if (行 === 0) Logger.log('      ' + sh.getName() + ' … 空');
    });
  });
}

// ---------- 分析シートの移し替えと、不要シートの片付け ----------

var 移し替え_元 = '回数券終了分析';
var 移し替え_先 = '回数券分析結果';
// 名前が変わった列の対応（元 → 先）
var 移し替え_列の対応 = {
  'ID': '回答ID',
  '1回目写真JSON': 'ビフォー写真JSON',
  '6回目写真JSON': 'アフター写真JSON',
};

// いま入っているお名前を並べる（表記ゆれを見つけるため）
function お名前の一覧を見る() {
  var live = getSpreadsheet_();
  ['測定履歴', '回答一覧', '回数券終了分析', '回数券分析結果'].forEach(function (名前) {
    var sh = live.getSheetByName(名前);
    if (!sh || sh.getLastRow() < 2) { Logger.log('― ' + 名前 + ' … 空'); return; }
    var 見出し = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var 列 = -1;
    ['顧客名', 'お名前', '氏名'].forEach(function (l) {
      if (列 < 0) { var i = 見出し.indexOf(l); if (i >= 0) 列 = i; }
    });
    if (列 < 0) { Logger.log('― ' + 名前 + ' … お名前の列なし'); return; }
    var v = sh.getRange(2, 列 + 1, sh.getLastRow() - 1, 1).getValues();
    var 数え = {};
    v.forEach(function (r) {
      var n = String(r[0] || '').trim();
      if (n) 数え[n] = (数え[n] || 0) + 1;
    });
    Logger.log('― ' + 名前);
    Object.keys(数え).sort().forEach(function (n) {
      Logger.log('   ' + n + ' … ' + 数え[n] + '行');
    });
  });
}

function 分析シートを移して片付ける() {
  var live = getSpreadsheet_();

  // 1. 控えを取る
  var 控え = DriveApp.getFileById(live.getId()).makeCopy(
    '【控え】' + live.getName() + ' ' +
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm')
  );
  Logger.log('控えを作りました: ' + 控え.getUrl());

  var 元 = live.getSheetByName(移し替え_元);
  var 先 = live.getSheetByName(移し替え_先);
  if (!先) { Logger.log('『' + 移し替え_先 + '』が見つかりません。中止します。'); return; }

  if (元 && 元.getLastRow() >= 2) {
    var 元見出し = 元.getRange(1, 1, 1, 元.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h).trim(); });
    var 先見出し = 先.getRange(1, 1, 1, 先.getLastColumn()).getValues()[0]
      .map(function (h) { return String(h).trim(); });

    // 先の「回答ID」の位置。重複を見分けるのに使う。
    var ID列 = 先見出し.indexOf('回答ID');
    var 既存 = {};
    if (ID列 >= 0 && 先.getLastRow() >= 2) {
      先.getRange(2, ID列 + 1, 先.getLastRow() - 1, 1).getValues().forEach(function (r, i) {
        var id = String(r[0] || '').trim();
        if (id) 既存[id] = i + 2;
      });
    }

    // 行き先の無い列は、文章として残す先を探す
    var メモ列 = -1;
    ['分析結果', '分析テキスト', '備考', 'メモ'].forEach(function (l) {
      if (メモ列 < 0) { var i = 先見出し.indexOf(l); if (i >= 0) メモ列 = i; }
    });

    var 元行 = 元.getRange(2, 1, 元.getLastRow() - 1, 元.getLastColumn()).getValues();
    var 追加 = 0, 更新 = 0, 余り = [];
    元行.forEach(function (row) {
      if (row.join('').trim() === '') return;
      var 新 = new Array(先見出し.length).fill('');
      var のこり = [];
      元見出し.forEach(function (h, i) {
        var 行き先 = 移し替え_列の対応[h] || h;
        var j = 先見出し.indexOf(行き先);
        if (j >= 0) {
          // お名前は空白を取り除いて揃える。「廣田　沙織」と「廣田沙織」が
          // 別人として分かれてしまうため。測定履歴側の表記に合わせる。
          新[j] = (行き先 === 'お名前' || 行き先 === '顧客名')
            ? String(row[i] == null ? '' : row[i]).trim().replace(/[\s　]+/g, '')
            : row[i];
        } else if (String(row[i] || '').trim() !== '') {
          のこり.push(h + '：' + row[i]);
        }
      });
      if (のこり.length) {
        if (メモ列 >= 0) {
          新[メモ列] = [新[メモ列], のこり.join(' / ')].filter(String).join('\n');
        } else {
          余り = 余り.concat(のこり.map(function (x) { return x.split('：')[0]; }));
        }
      }
      var id = ID列 >= 0 ? String(新[ID列] || '').trim() : '';
      if (id && 既存[id]) {
        先.getRange(既存[id], 1, 1, 先見出し.length).setValues([新]);
        更新 += 1;
      } else {
        先.appendRow(新);
        追加 += 1;
      }
    });
    Logger.log('『' + 移し替え_元 + '』→『' + 移し替え_先 + '』 追加 ' + 追加 + '件 / 更新 ' + 更新 + '件');
    if (余り.length) {
      var 一意 =余り.filter(function (v, i, a) { return a.indexOf(v) === i; });
      Logger.log('  ※ 行き先の無かった列: ' + 一意.join(', ') + '（値は移せていません）');
    }

    live.deleteSheet(元);
    Logger.log('『' + 移し替え_元 + '』を削除しました');
  } else {
    Logger.log('『' + 移し替え_元 + '』は空、または見つかりません（移し替えなし）');
    if (元) { live.deleteSheet(元); Logger.log('  空だったので削除しました'); }
  }

  // 2. 空のシートを片付ける（いま有効なアンケートの名前は残す）
  var 残す = {};
  try {
    getSurveys_().forEach(function (s) { 残す[String(s.title).trim()] = true; });
  } catch (error) { /* 取れなければ、アンケート名らしきものは触らない */ }
  [MASTER_SHEET_NAME, MEASUREMENTS_SHEET_NAME, BIJIRIS_POSTS_SHEET_NAME,
   BIJIRIS_POST_ATTACHMENTS_SHEET_NAME, MEMBER_SHEET_NAME,
   TICKET_SURVEY_STORAGE_SHEET_NAME, 会員別_一覧シート名].forEach(function (n) {
    残す[n] = true;
  });

  var 消した = [];
  live.getSheets().forEach(function (sh) {
    var 名前 = sh.getName();
    if (残す[名前]) return;
    if (sh.getLastRow() >= 2) return;              // 中身があるものは触らない
    if (live.getSheets().length - 消した.length <= 1) return;  // 最後の1枚は消せない
    live.deleteSheet(sh);
    消した.push(名前);
  });
  Logger.log('空のシートを削除: ' + (消した.length ? 消した.join(' / ') : 'なし'));
  Logger.log('');
  Logger.log('残ったシート: ' + live.getSheets().map(function (s) {
    return s.getName() + '(' + Math.max(0, s.getLastRow() - 1) + ')';
  }).join(' / '));
}

// ---------- 古いファイルの測定履歴を取りこぼさない ----------

// もう一方のファイルの測定履歴を見て、いまのファイルに無い記録だけを足す。
// アンケート回答と豆知識は、方針により移さない。
function 古い測定履歴を取り込む() {
  var live = getSpreadsheet_();
  var old = 取込_もう一方のファイル_();
  if (!old) { Logger.log('もう一方のファイルが見つかりません'); return; }

  var oldSheet = old.getSheetByName(MEASUREMENTS_SHEET_NAME);
  if (!oldSheet || oldSheet.getLastRow() < 2) { Logger.log('古い測定履歴は空です'); return; }

  var sheet = live.getSheetByName(MEASUREMENTS_SHEET_NAME);
  if (!sheet) { Logger.log('いまの測定履歴が見つかりません'); return; }

  // いまある「お名前＋測定日」を控える
  var ある = {};
  if (sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, MEASUREMENT_HEADERS.length).getValues()
      .forEach(function (r) {
        var 名前 = String(r[3] || '').trim().replace(/[\s　]+/g, '');
        var 日 = 取込_日付にする_(r[5]);
        if (名前 && 日) ある[名前 + '|' + 取込_日付の文字_(日)] = true;
      });
  }

  var 元 = oldSheet.getRange(2, 1, oldSheet.getLastRow() - 1, MEASUREMENT_HEADERS.length).getValues();
  var 足す = [], 重複 = 0;
  元.forEach(function (r) {
    var 名前 = String(r[3] || '').trim().replace(/[\s　]+/g, '');
    var 日 = 取込_日付にする_(r[5]);
    if (!名前 || !日) return;
    var かぎ = 名前 + '|' + 取込_日付の文字_(日);
    if (ある[かぎ]) { 重複 += 1; Logger.log('  すでにある: ' + かぎ); return; }
    var row = r.slice();
    row[3] = 名前;
    row[5] = 取込_日付の文字_(日);
    row[2] = 取込_測定IDを作る_(名前, 日);
    足す.push(row);
    ある[かぎ] = true;
    Logger.log('  足す: ' + かぎ);
  });

  足す.forEach(function (row) { sheet.appendRow(row); });
  Logger.log('古い測定履歴 ' + 元.length + '件のうち、' + 足す.length + '件を追加（重複 ' + 重複 + '件）');

  if (足す.length) {
    var 結果 = 測定記録の会員別一覧を作る();
    Logger.log('会員別の一覧を作り直しました: ' + 結果.会員数 + '名 / ' + 結果.件数 + '件');
  }
}

// ---------- 送信がうまくいかないときの調べ物 ----------

// アプリが記録したエラーの新しいものから並べる（読むだけ）
function 最近のエラーを見る() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(ERROR_LOGS_PROPERTY_KEY);
  var 一覧 = [];
  try { 一覧 = JSON.parse(raw || '[]'); } catch (error) { 一覧 = []; }
  if (!一覧.length) { Logger.log('エラーの記録はありません'); return; }

  Logger.log('新しいものから最大15件');
  一覧.slice(-15).reverse().forEach(function (e) {
    Logger.log('― ' + (e.at || '') + '  [' + (e.source || '') + ']');
    Logger.log('   ' + String(e.message || '').slice(0, 300));
    var d = e.detail ? JSON.stringify(e.detail) : '';
    if (d && d !== '{}') Logger.log('   ' + d.slice(0, 300));
  });
}

// 直近に届いた回答を、新しいものから並べる（どこまで届いているかの確認）
function 最近の回答を見る() {
  var 一覧 = getResponses_({ includeTrashed: true });
  if (!一覧.length) { Logger.log('回答はまだありません'); return; }
  一覧.sort(function (a, b) {
    return new Date(b.submittedAt || 0) - new Date(a.submittedAt || 0);
  });
  Logger.log('新しいものから最大10件（全' + 一覧.length + '件）');
  一覧.slice(0, 10).forEach(function (r) {
    var 写真 = 0;
    (r.answers || []).forEach(function (a) {
      if (a && Array.isArray(a.files)) 写真 += a.files.length;
    });
    Logger.log('― ' + (r.submittedAt || '') + '  ' + (r.surveyTitle || '') +
      '  ' + (r.customerName || '') + '  写真' + 写真 + '枚');
  });
}
