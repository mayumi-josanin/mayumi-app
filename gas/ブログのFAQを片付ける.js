// 「まゆみのブログ」を、お困りのときの案内（FAQ）から外す道具。
//
//   ブログのFAQを下見()     … 何が変わるかを見るだけ。**書きません**
//   ブログのFAQを片付ける()  … 実際に直す
//
// ─────────────────────────────────────────────────────────
// なぜ消さずに「非公開」にするのか
//
// FAQシートは1列目の「状態」が '公開' の行だけをお客様に出す
// （getSupportFaqEntries_）。**非公開にすれば画面から消える。**
//
// 行を削ると番号が繰り上がり、戻すときに元の場所が分からなくなる。
// CLAUDE.md の「消す前に控えを取る」も、消さずに済むならそれが一番よい。
//
// ─────────────────────────────────────────────────────────
// 何をするか（2行）
//
//   「まゆみのブログとは何ですか？」        → 状態を 非公開 に
//   「公式LINEやSNSの開き方を知りたい」      → 回答から「まゆみのブログ」を削る
//                                          （この質問自体は残す）
//
// **行番号ではなく質問の文で探す。**番号は動く。
//
// シートの列: 1状態 2カテゴリ 3質問 4キーワード 5回答 6優先度 7更新日時

var ブログFAQ_消す質問 = 'まゆみのブログとは何ですか？';
var ブログFAQ_直す質問 = '公式LINEやSNSの開き方を知りたい';

// **getOrCreateSpreadsheet() を呼ばない。**
// あれは見つからなければ**新しいスプレッドシートを作る**（5番目の分岐）。
// 読むだけの道具から呼んで、本番とは別の空の表を作ってしまう事故がある。
function ブログFAQ_表を開く_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
    || FALLBACK_SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);   // 無ければ例外。作らない
}

function ブログFAQ_回答を直す_(文) {
  return String(文 || '')
    .replace(/、まゆみのブログ/g, '')
    .replace(/まゆみのブログ、/g, '')
    .replace(/まゆみのブログ/g, '');
}

function ブログFAQ_調べる_() {
  var sheet = ブログFAQ_表を開く_().getSheetByName('APP_SUPPORT_FAQ');
  if (!sheet) throw new Error('APP_SUPPORT_FAQ のシートが見つかりません');
  var 最終行 = sheet.getLastRow();
  if (最終行 < 2) return { sheet: sheet, 見つけた: [] };

  var 値 = sheet.getRange(2, 1, 最終行 - 1, 7).getDisplayValues();
  var 見つけた = [];
  for (var i = 0; i < 値.length; i++) {
    var 行 = 値[i];
    var 質問 = String(行[2] || '').trim();
    if (質問 === ブログFAQ_消す質問) {
      見つけた.push({ 行番号: i + 2, やること: '非公開にする', いま: 行[0], あと: '非公開', 質問: 質問 });
    } else if (質問 === ブログFAQ_直す質問) {
      var 新 = ブログFAQ_回答を直す_(行[4]);
      if (新 !== String(行[4] || '')) {
        見つけた.push({ 行番号: i + 2, やること: '回答を直す', いま: 行[4], あと: 新, 質問: 質問 });
      }
    } else if (String(行[4] || '').indexOf('まゆみのブログ') !== -1) {
      // 想定外。**黙って直さない。**人が見て決める
      見つけた.push({ 行番号: i + 2, やること: '**要確認**（回答にブログの記述あり）', いま: 行[4], あと: '', 質問: 質問 });
    }
  }
  return { sheet: sheet, 見つけた: 見つけた };
}

function ブログのFAQを下見() {
  var r = ブログFAQ_調べる_();
  Logger.log('■ 見つけた行: ' + r.見つけた.length + '件（**まだ何も書いていません**）');
  r.見つけた.forEach(function (x) {
    Logger.log('');
    Logger.log('  ' + x.行番号 + '行目  ' + x.やること);
    Logger.log('    質問: ' + x.質問);
    Logger.log('    いま: ' + x.いま);
    if (x.あと) Logger.log('    あと: ' + x.あと);
  });
  if (!r.見つけた.length) Logger.log('  直すところはありません。');
}

function ブログのFAQを片付ける() {
  var r = ブログFAQ_調べる_();
  var 直した = 0;
  r.見つけた.forEach(function (x) {
    if (x.やること === '非公開にする') {
      r.sheet.getRange(x.行番号, 1).setValue('非公開');
      直した++;
    } else if (x.やること === '回答を直す') {
      r.sheet.getRange(x.行番号, 5).setValue(x.あと);
      直した++;
    } else {
      Logger.log('  ' + x.行番号 + '行目は**触っていません**。人が見て決めてください: ' + x.質問);
    }
  });
  Logger.log('■ ' + 直した + '件を直しました。');
  Logger.log('  戻すときは、状態を「公開」に戻すだけです（行は消していません）。');
  Logger.log('  お客様の画面に出るまで、キャッシュの都合で最大10分かかります。');
}
