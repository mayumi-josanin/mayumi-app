// 管理アプリの設定が、いまサーバーでどうなっているかを見る道具。**読むだけ。**
//
// 特典の保存ができない原因を切り分けるために作った。
// 保存は「送って、読み直して一致を確かめる」作りだが、送信は応答を読まないため、
// サーバー側で例外が出ても画面には「確認できませんでした」としか出ない。

// 空のときだけ片付ける。中身があれば触らない。
function 豆知識を片付ける() {
  var book = getSpreadsheet_();
  var 消したシート = [];
  [BIJIRIS_POSTS_SHEET_NAME, BIJIRIS_POST_ATTACHMENTS_SHEET_NAME].forEach(function (n) {
    var sh = book.getSheetByName(n);
    if (!sh) return;
    if (sh.getLastRow() > 1) { Logger.log('  中身があるので残します: ' + n); return; }
    if (book.getSheets().length <= 1) return;
    book.deleteSheet(sh);
    消したシート.push(n);
  });

  var 消したフォルダ = [];
  var it = DriveApp.getFoldersByName(BIJIRIS_POSTS_FOLDER_NAME);
  while (it.hasNext()) {
    var f = it.next();
    var 中身 = 0;
    var fi = f.getFiles(); while (fi.hasNext()) { fi.next(); 中身 += 1; }
    var fo = f.getFolders(); while (fo.hasNext()) { fo.next(); 中身 += 1; }
    if (中身 > 0) { Logger.log('  中身があるので残します: ' + f.getName()); continue; }
    f.setTrashed(true);   // ゴミ箱なので30日は戻せる
    消したフォルダ.push(f.getName());
  }

  豆知識の控えを捨てる_();
  appendAuditLog_('bijirisPosts.cleanup', { sheets: 消したシート, folders: 消したフォルダ });
  Logger.log('');
  Logger.log('■ 片付けたシート: ' + (消したシート.join('・') || 'なし'));
  Logger.log('■ ゴミ箱へ入れたフォルダ: ' + (消したフォルダ.join('・') || 'なし'));
}

// 豆知識まわりに何が残っているかを見る。**読むだけ。**
function 豆知識の残りを見る() {
  var book = getSpreadsheet_();
  Logger.log('■ スプレッドシート: ' + book.getName());
  [BIJIRIS_POSTS_SHEET_NAME, BIJIRIS_POST_ATTACHMENTS_SHEET_NAME].forEach(function (n) {
    var sh = book.getSheetByName(n);
    Logger.log('  シート「' + n + '」: ' +
      (sh ? Math.max(0, sh.getLastRow() - 1) + '行' : '無し'));
  });
  Logger.log('');

  Logger.log('■ ドライブのフォルダ「' + BIJIRIS_POSTS_FOLDER_NAME + '」');
  try {
    var it = DriveApp.getFoldersByName(BIJIRIS_POSTS_FOLDER_NAME);
    var 見つけた = 0;
    while (it.hasNext()) {
      var f = it.next();
      var 中身 = 0;
      var fi = f.getFiles(); while (fi.hasNext()) { fi.next(); 中身 += 1; }
      var fo = f.getFolders(); while (fo.hasNext()) { fo.next(); 中身 += 1; }
      Logger.log('  ' + f.getName() + ' … 中身 ' + 中身 + '件 / ' + f.getUrl());
      見つけた += 1;
    }
    if (!見つけた) Logger.log('  ありません');
  } catch (e) {
    Logger.log('  読めませんでした: ' + e.message);
  }
  Logger.log('');
  Logger.log('  空であれば 豆知識を片付ける() でシートとフォルダを整理できます。');
}

function 設定の下見() {
  var p = getPreferences_();
  Logger.log('■ いまの設定');
  Logger.log('');
  Logger.log('  通知を出す: ' + p.notificationEnabled);
  Logger.log('  通知メール: 「' + (p.notificationEmail || '（空）') + '」');
  Logger.log('    ※ 通知を出す=true なのにメールが空だと、設定の保存が必ず失敗する');
  Logger.log('  オーナーのメール: 「' + (getOwnerEmail_() || '（取得できない）') + '」');
  Logger.log('');
  Logger.log('  特典の設定: ' + JSON.stringify(p.milestoneRewardConfig));
  Logger.log('  キャンペーンスタンプ: ' + p.campaignStampEnabled);
  Logger.log('  自動バックアップ: ' + p.autoBackupEnabled + ' / ' + p.backupHour + '時 / ' + p.retentionDays + '日');
  Logger.log('');
  Logger.log('  豆知識の件数: ' + getBijirisPosts_({ includeDrafts: true }).length +
    '（公開ぶん ' + getBijirisPosts_({ publishedOnly: true }).length + '）');
}


