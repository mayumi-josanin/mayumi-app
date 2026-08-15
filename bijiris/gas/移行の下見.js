// 移行の下見 — データベースへ移すとき、何件移せて何件が突き合わないかを数える。
//
//   移行の下見をする()  … すべて調べて結果を新しいスプレッドシートに書き出す
//
// **実データは1件も変えない。読むだけ。**
// 移行しない判断をしても無駄にならないよう、下調べだけを切り出してある。
//
// いちばん確かめたいこと:
//   ビジリスの測定・分析は「お名前の文字列」で会員と結びついている。
//   会員IDに置き換えられるか（＝一意に決まるか）をここで確認する。

var 下見_まゆみID = '1gIcUGxg2PEuFoU5a_IgQ6lDWgghceJ7v2dgqo9iPe4w';
var 下見_ビジリスID = '1pONQ8MfFSllKNOeQlcp56IRon3ZRWfFkbnjEDPchq8E';

// 記録に残さないと決めた種別（移行対象から外す）
var 下見_履歴から外す種別 = ['syncUserDeviceSession', 'syncUserRewardStatus'];

function 下見_空か_(x) { return x === '' || x === null || x === undefined; }

// お名前を突き合わせるための正規化。空白・全角半角のゆれを吸収する。
function 下見_名前をそろえる_(値) {
  return String(値 == null ? '' : 値)
    .replace(/[\s　]+/g, '')
    .normalize('NFKC')
    .trim();
}

function 下見_シートを読む_(id, 名前) {
  var sh = SpreadsheetApp.openById(id).getSheetByName(名前);
  if (!sh) return null;
  var 行数 = sh.getLastRow();
  var 列数 = sh.getLastColumn();
  if (行数 < 1 || 列数 < 1) return { 見出し: [], 行: [] };
  var v = sh.getRange(1, 1, 行数, 列数).getValues();
  // 中身が0件でも見出しは返す。「列が無い」と誤って報告しないため。
  return { 見出し: v[0], 行: v.slice(1) };
}

// ---------- 会員 ----------

function 下見_会員を読む_() {
  var d = 下見_シートを読む_(下見_まゆみID, '会員データ');
  var 位置 = {};
  d.見出し.forEach(function (h, i) { 位置[String(h)] = i; });

  var 会員 = [];
  var 名前ごと = {};      // そろえた名前 → 会員の配列
  var ID重複 = {};

  d.行.forEach(function (row, i) {
    if (!row.some(function (x) { return !下見_空か_(x); })) return;
    var id = String(row[位置['ID']] || '').trim();
    var 名 = 下見_名前をそろえる_(row[位置['氏名']]);
    var m = {
      行: i + 2,
      id: id,
      名: 名,
      電話: String(row[位置['電話番号']] || '').trim(),
      生年月日: row[位置['生年月日']],
      パスコード: String(row[位置['パスコード']] || '').trim(),
      ハッシュ: String(row[位置['パスワードハッシュ']] || '').trim(),
      権限: String(row[位置['権限']] || '').trim(),
    };
    会員.push(m);
    if (id) ID重複[id] = (ID重複[id] || 0) + 1;
    if (名) {
      if (!名前ごと[名]) 名前ごと[名] = [];
      名前ごと[名].push(m);
    }
  });
  return { 会員: 会員, 名前ごと: 名前ごと, ID重複: ID重複 };
}

// ---------- 名前で会員を探す ----------

function 下見_会員を探す_(名前ごと, 名) {
  var かぎ = 下見_名前をそろえる_(名);
  if (!かぎ) return { 状態: '名前が空' };
  var 候補 = 名前ごと[かぎ];
  if (!候補 || !候補.length) return { 状態: '見つからない' };
  if (候補.length > 1) return { 状態: '複数該当', 件数: 候補.length };
  return { 状態: 'ok', 会員: 候補[0] };
}

// ---------- 本体 ----------

function 移行の下見をする() {
  var 出 = [];
  var 書く = function (s) { 出.push(s === undefined ? '' : String(s)); };

  書く('# 移行の下見');
  書く('作成: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
  書く('');
  書く('実データは変えていません。読んで数えただけです。');
  書く('お名前・電話番号などの中身は出しません。件数と行番号だけです。');
  書く('');

  // ===== 1. 会員 =====
  var M = 下見_会員を読む_();
  書く('## 1. 会員（members）');
  書く('');
  書く('- 会員の行: ' + M.会員.length);

  var ID空 = M.会員.filter(function (m) { return !m.id; });
  var ID重複 = Object.keys(M.ID重複).filter(function (k) { return M.ID重複[k] > 1; });
  var 名前空 = M.会員.filter(function (m) { return !m.名; });
  var 同名 = Object.keys(M.名前ごと).filter(function (k) { return M.名前ごと[k].length > 1; });

  書く('- IDが空: ' + ID空.length + (ID空.length ? '（行 ' + ID空.map(function (m) { return m.行; }).join(', ') + '）' : ''));
  書く('- IDが重複: ' + ID重複.length + '種類');
  書く('- お名前が空: ' + 名前空.length);
  書く('- **同じお名前が複数いる: ' + 同名.length + '組**');
  同名.forEach(function (k) {
    書く('    - ' + M.名前ごと[k].length + '名（行 ' + M.名前ごと[k].map(function (m) { return m.行; }).join(', ') + '）');
  });
  書く('');

  var 平文 = M.会員.filter(function (m) { return m.パスコード; });
  var ハッシュ済 = M.会員.filter(function (m) { return m.ハッシュ; });
  書く('- パスコードが平文で入っている: **' + 平文.length + '名** → 移行時にハッシュへ変換する');
  書く('- すでにハッシュ方式: ' + ハッシュ済.length + '名');
  書く('- どちらも無い（入り口に入れない）: ' +
    M.会員.filter(function (m) { return !m.パスコード && !m.ハッシュ; }).length + '名');
  書く('');

  // ===== 2. ビジリスの記録を会員に結び直せるか =====
  書く('## 2. ビジリスの記録を会員IDに結び直せるか');
  書く('');
  書く('いまは「お名前の文字列」でしか結びついていない。ここが移行の山場。');
  書く('');

  var 調べる = [
    { シート: '測定履歴', 名前列: '顧客名' },
    { シート: '回数券分析結果', 名前列: 'お名前' },
    { シート: '回答一覧', 名前列: 'お名前' },
  ];

  var 要確認 = [];

  調べる.forEach(function (t) {
    var d = 下見_シートを読む_(下見_ビジリスID, t.シート);
    if (!d) { 書く('### ' + t.シート + ' … シートが見つかりません'); 書く(''); return; }
    var 名列 = d.見出し.indexOf(t.名前列);
    if (名列 < 0) { 書く('### ' + t.シート + ' … 「' + t.名前列 + '」列がありません'); 書く(''); return; }

    var 集計 = { ok: 0, 見つからない: 0, 複数該当: 0, 名前が空: 0 };
    var 困った名前 = {};

    d.行.forEach(function (row, i) {
      if (!row.some(function (x) { return !下見_空か_(x); })) return;
      var 結果 = 下見_会員を探す_(M.名前ごと, row[名列]);
      集計[結果.状態] = (集計[結果.状態] || 0) + 1;
      if (結果.状態 !== 'ok') {
        var かぎ = 下見_名前をそろえる_(row[名列]) || '(空)';
        if (!困った名前[かぎ]) 困った名前[かぎ] = { 状態: 結果.状態, 行: [] };
        困った名前[かぎ].行.push(i + 2);
      }
    });

    var 合計 = 集計.ok + 集計.見つからない + 集計.複数該当 + 集計.名前が空;
    書く('### ' + t.シート + '（' + 合計 + '件）');
    書く('');
    書く('| 結果 | 件数 |');
    書く('|------|------|');
    書く('| 会員IDに結べる | ' + 集計.ok + ' |');
    書く('| 会員が見つからない | ' + 集計.見つからない + ' |');
    書く('| 同名が複数いて決められない | ' + 集計.複数該当 + ' |');
    書く('| お名前が空 | ' + 集計.名前が空 + ' |');
    書く('');

    var 鍵 = Object.keys(困った名前);
    if (鍵.length) {
      書く('突き合わないお名前（' + 鍵.length + '通り）:');
      鍵.sort().forEach(function (k) {
        var x = 困った名前[k];
        書く('  - 「' + k + '」… ' + x.状態 + '（' + t.シート + ' の ' + x.行.join(', ') + '行目）');
        要確認.push(t.シート + ' / ' + k + ' … ' + x.状態);
      });
      書く('');
    }
  });

  // ===== 3. 操作履歴 =====
  書く('## 3. 操作履歴（audit_logs）');
  書く('');
  var A = 下見_シートを読む_(下見_まゆみID, 'ADMIN_AUDIT_LOG');
  if (A) {
    var 種別列 = A.見出し.indexOf('種別');
    var 全部 = 0;
    var 移す = 0;
    A.行.forEach(function (row) {
      if (!row.some(function (x) { return !下見_空か_(x); })) return;
      全部 += 1;
      if (下見_履歴から外す種別.indexOf(String(row[種別列] || '')) < 0) 移す += 1;
    });
    書く('- いまの記録: ' + 全部 + '行');
    書く('- 自動同期ぶんを除いて移す: **' + 移す + '行**');
    書く('- 移さない: ' + (全部 - 移す) + '行');
  } else {
    書く('- シートが見つかりません');
  }
  書く('');

  // ===== 4. その他の表 =====
  書く('## 4. そのまま移せる表');
  書く('');
  書く('| 表 | 件数 |');
  書く('|----|------|');
  [
    ['ブログ・お知らせ', 'news'], ['カレンダー', 'calendar_events'], ['MENUS', 'menus'],
    ['商品マスタ', 'products'], ['カテゴリマスタ', 'categories'], ['APP_SUPPORT_FAQ', 'faqs'],
    ['PUSH_NOTICES', 'push_notices'], ['MENU_REVENUE', 'revenue_records（メニュー）'],
    ['PRODUCT_REVENUE', 'revenue_records（商品）'], ['管理マスタ', 'supplier_prices'],
    ['BACKUP_LOG', 'backups'],
  ].forEach(function (pair) {
    var d = 下見_シートを読む_(下見_まゆみID, pair[0]);
    var n = d ? d.行.filter(function (r) { return r.some(function (x) { return !下見_空か_(x); }); }).length : 0;
    書く('| ' + pair[0] + ' → ' + pair[1] + ' | ' + n + ' |');
  });
  書く('');

  // ===== まとめ =====
  書く('## まとめ');
  書く('');
  if (要確認.length) {
    書く('**人の目で確認が要るもの: ' + 要確認.length + '件**');
    書く('');
    要確認.forEach(function (s) { 書く('- ' + s); });
  } else {
    書く('**突き合わないものはありませんでした。すべて会員IDに結べます。**');
  }
  書く('');

  var 名前 = '移行の下見_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm');
  var ss = SpreadsheetApp.create(名前);
  var sh = ss.getSheets()[0];
  sh.setName('下見');
  if (出.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 出.length - sh.getMaxRows());
  sh.getRange(1, 1, 出.length, 1).setValues(出.map(function (s) { return ["'" + s]; }));

  Logger.log('書き出しました: ' + 名前);
  Logger.log('  ' + ss.getUrl());
  Logger.log('');
  出.forEach(function (s) { Logger.log(s); });
  return ss.getId();
}
