// 写真の置き場が、どこを指しているかを確かめる道具。**読むだけ。**
//
//   写真フォルダを確かめる() … 設定されているフォルダと、Drive の実物を照らす
//
// ## なぜ要るのか
//
// `getRootPhotoFolder_()` は、`PHOTO_ROOT_FOLDER_ID` が使えないと
// **新しい「Bijiris」フォルダを作って、そのIDを保存し直す。**
//
// つまり呼んだだけで置き場が変わることがある。
// 2026-08-25、共有の下見が「0件」を返したので、それを疑って作った。
//
// **この道具は getRootPhotoFolder_() を呼ばない。**呼ぶと作ってしまう。
// 保存されているIDを直接見て、Drive も直接探す。

// **並びはわざとこの順です。**エディタは先頭の関数を既定で選びます。

// **本物の置き場を、中身から突き止める。**読むだけ。
//
// 2026-08-25、私の実行で `getRootPhotoFolder_()` が空の「Bijiris」を作り、
// `PHOTO_ROOT_FOLDER_ID` をそちらへ書き換えてしまった。
// 元のIDが分からないので、**中身から本物を探す。**
//
// 「計測時」フォルダの親が、本物の置き場。
function 本物の置き場を探す() {
  Logger.log('■ 「' + MEASUREMENT_TIME_ROOT_NAME + '」フォルダを探す');
  var it = DriveApp.getFoldersByName(MEASUREMENT_TIME_ROOT_NAME);
  var n = 0;
  while (it.hasNext() && n < 10) {
    var f = it.next();
    n++;
    Logger.log('');
    Logger.log('   ' + n + ') ' + f.getUrl());
    Logger.log('       作成: ' + f.getDateCreated());
    Logger.log('       中の子フォルダ: ' + 写真確認_数える_(f.getFolders()) + '個'
      + ' / ファイル ' + 写真確認_数える_(f.getFiles()) + '件');
    var 親 = f.getParents();
    while (親.hasNext()) {
      var p = 親.next();
      Logger.log('       **親（＝本物の置き場の候補）**: ' + p.getName());
      Logger.log('           ' + p.getUrl());
      Logger.log('           作成: ' + p.getDateCreated()
        + ' / 子フォルダ ' + 写真確認_数える_(p.getFolders()) + '個');
    }
  }
  if (!n) Logger.log('   見つかりません。');

  Logger.log('');
  Logger.log('■ 「' + BIJIRIS_POSTS_FOLDER_NAME + '」フォルダの親も見る（同じ置き場のはず）');
  var it2 = DriveApp.getFoldersByName(BIJIRIS_POSTS_FOLDER_NAME);
  var m = 0;
  while (it2.hasNext() && m < 5) {
    var g = it2.next();
    m++;
    var 親2 = g.getParents();
    while (親2.hasNext()) {
      var q = 親2.next();
      Logger.log('   親: ' + q.getName() + ' … ' + q.getUrl());
    }
  }
  if (!m) Logger.log('   見つかりません。');
}

function 写真フォルダを確かめる() {
  var props = PropertiesService.getScriptProperties();
  var id = String(props.getProperty('PHOTO_ROOT_FOLDER_ID') || '').trim();

  Logger.log('■ 設定されている置き場（PHOTO_ROOT_FOLDER_ID）');
  if (!id) {
    Logger.log('   **設定されていません。**');
  } else {
    Logger.log('   ID の文字数: ' + id.length);
    try {
      var f = DriveApp.getFolderById(id);
      Logger.log('   名前: ' + f.getName());
      Logger.log('   URL:  ' + f.getUrl());
      Logger.log('   作成: ' + f.getDateCreated());
      Logger.log('   子フォルダ: ' + 写真確認_数える_(f.getFolders()) + '個');
      Logger.log('   直下のファイル: ' + 写真確認_数える_(f.getFiles()) + '件');
    } catch (e) {
      Logger.log('   **このIDのフォルダを開けません。** ' + e.message);
      Logger.log('   → 次に getRootPhotoFolder_() が呼ばれると、新しく作られます。');
    }
  }

  Logger.log('');
  Logger.log('■ マイドライブにある「' + ROOT_DRIVE_FOLDER_NAME + '」フォルダ');
  var it = DriveApp.getRootFolder().getFoldersByName(ROOT_DRIVE_FOLDER_NAME);
  var n = 0;
  while (it.hasNext()) {
    var g = it.next();
    n++;
    Logger.log('   ' + n + ') ' + g.getUrl());
    Logger.log('       作成: ' + g.getDateCreated()
      + ' / 子フォルダ ' + 写真確認_数える_(g.getFolders()) + '個'
      + ' / ファイル ' + 写真確認_数える_(g.getFiles()) + '件');
    Logger.log('       設定と同じ: ' + (g.getId() === id ? '**はい**' : 'いいえ'));
  }
  if (!n) Logger.log('   ありません。');

  Logger.log('');
  Logger.log('■ 検索でも探す（マイドライブの外にあるかもしれない）');
  var it2 = DriveApp.getFoldersByName(ROOT_DRIVE_FOLDER_NAME);
  var m = 0;
  while (it2.hasNext() && m < 10) {
    var h = it2.next();
    m++;
    Logger.log('   ' + m + ') 子フォルダ ' + 写真確認_数える_(h.getFolders()) + '個'
      + ' / ファイル ' + 写真確認_数える_(h.getFiles()) + '件'
      + ' / 設定と同じ: ' + (h.getId() === id ? '**はい**' : 'いいえ'));
    Logger.log('       ' + h.getUrl());
  }
  if (!m) Logger.log('   ありません。');

  Logger.log('');
  Logger.log('■ 記録にある写真が、実際に開けるか（3件だけ試す）');
  写真確認_記録から試す_();
}

function 写真確認_数える_(it) {
  var n = 0;
  while (it.hasNext()) { it.next(); n++; if (n > 500) return '500超'; }
  return n;
}

// 回答一覧に残っているファイルIDを拾って、実際に開けるか試す。
// **置き場がどこであれ、これが開ければアプリは写真を出せる。**
function 写真確認_記録から試す_() {
  var sh = getSpreadsheet_().getSheetByName(MASTER_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) { Logger.log('   回答がありません。'); return; }

  var 値 = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var ids = [];
  値.forEach(function (row) {
    row.forEach(function (cell) {
      var s = String(cell || '');
      if (s.indexOf('fileId') < 0) return;
      var m = s.match(/"fileId"\s*:\s*"([^"]+)"/g) || [];
      m.forEach(function (x) {
        var id = (x.match(/"fileId"\s*:\s*"([^"]+)"/) || [])[1];
        if (id && ids.indexOf(id) < 0 && ids.length < 3) ids.push(id);
      });
    });
  });

  if (!ids.length) { Logger.log('   記録の中にファイルIDが見つかりません。'); return; }

  ids.forEach(function (id) {
    try {
      var f = DriveApp.getFileById(id);
      var 大きさ = f.getBlob().getBytes().length;
      var 共有 = String(f.getSharingAccess());
      Logger.log('   開けました（' + 大きさ + 'バイト） 共有: ' + 共有);
      Logger.log('       置き場: ' + (f.getParents().hasNext() ? f.getParents().next().getName() : '（不明）'));
    } catch (e) {
      Logger.log('   **開けません**: ' + e.message);
    }
  });
}


