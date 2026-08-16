// 管理アプリの設定が、いまサーバーでどうなっているかを見る道具。**読むだけ。**
//
// 特典の保存ができない原因を切り分けるために作った。
// 保存は「送って、読み直して一致を確かめる」作りだが、送信は応答を読まないため、
// サーバー側で例外が出ても画面には「確認できませんでした」としか出ない。

// 記録を読む通信で、本当に何も書かないかを確かめる。**読むだけ。**
//
// 顧客管理まるごとを読む前後で見比べる。1文字でも変われば、どこかで書いている。
// 「読むだけのはずが書いていた」のが、スタンプが消えた原因だったため、
// 直したあとも、ここで見張れるようにしておく。
function 読むだけか確かめる() {
  var 鍵 = CUSTOMER_PROFILES_PROPERTY_KEY;
  var 前 = PropertiesService.getScriptProperties().getProperty(鍵) || '{}';

  var profiles = getCustomerProfiles_() || {};
  var 名前一覧 = Object.keys(profiles).map(function (k) {
    return String((profiles[k] || {}).name || '');
  }).filter(Boolean);

  名前一覧.forEach(function (名) {
    var r = profiles[名] || {};
    getCustomerHistoryPayload_({
      memberNumber: r.memberNumber,
      customerName: 名,
      matchByNameOnly: true,
      includeTrashed: false
    });
  });

  var 後 = PropertiesService.getScriptProperties().getProperty(鍵) || '{}';

  Logger.log('■ ' + 名前一覧.length + '名ぶん、お客様アプリと同じ読み方をしました');
  Logger.log('');
  if (前 === 後) {
    Logger.log('■ 顧客管理は1文字も変わっていません。読むだけになっています。');
  } else {
    Logger.log('■ **顧客管理が変わりました。まだどこかで書いています。**');
    Logger.log('    読む前: ' + 前.length + '文字 ／ 読んだ後: ' + 後.length + '文字');
    var 前の表 = JSON.parse(前), 後の表 = JSON.parse(後);
    Object.keys(後の表).forEach(function (k) {
      if (JSON.stringify(前の表[k]) !== JSON.stringify(後の表[k])) {
        Logger.log('    変わった方: ' + k);
        Logger.log('      前: ' + JSON.stringify(前の表[k]));
        Logger.log('      後: ' + JSON.stringify(後の表[k]));
      }
    });
  }
}

// お客様アプリが受け取る内容を、そのまま見る。**読むだけ。**
// 「管理側で登録したのにカードが出ない」の切り分け用。
function お客様に渡る内容を見る() {
  var 名前 = '前多洋子';   // 確かめたい方のお名前
  var payload = getCustomerHistoryPayload_({
    customerName: 名前,
    matchByNameOnly: true,
    includeTrashed: false
  });
  Logger.log('■ ' + 名前 + ' さんに渡る内容');
  Logger.log('');
  Logger.log('  回答: ' + (payload.responses || []).length + '件');
  Logger.log('  計測: ' + (payload.measurements || []).length + '件');
  var p = payload.customerProfile;
  Logger.log('  顧客情報: ' + (p ? 'あり' : '**無し**'));
  if (p) {
    Logger.log('    お名前: ' + p.name);
    Logger.log('    会員番号: ' + p.memberNumber);
    Logger.log('    回数券カード: ' + JSON.stringify(p.activeTicketCard));
    Logger.log('    スタンプ手当て: ' + p.ticketStampAdjustment);
  }
  Logger.log('');
  Logger.log('  ※ 回数券カードが null だと、アプリは「回数券を追加」を出す。');

  // 会員番号だけを手がかりにしても、同じ内容が返るかを見る。
  // お名前を一切渡さずに引けるかどうかが、切り替えの成否になる。
  if (p && p.memberNumber) {
    var 番号で = getCustomerHistoryPayload_({
      memberNumber: p.memberNumber,
      customerName: '',
      matchByNameOnly: true,
      includeTrashed: false
    });
    var q = 番号で.customerProfile;
    Logger.log('');
    Logger.log('■ 会員番号（' + p.memberNumber + '）だけで引いた場合');
    Logger.log('    お名前: ' + (q ? q.name : '**引けない**'));
    Logger.log('    回答: ' + (番号で.responses || []).length + '件 ／ 計測: ' + (番号で.measurements || []).length + '件');
    Logger.log('    回数券カード: ' + JSON.stringify(q && q.activeTicketCard));
    Logger.log('    スタンプ手当て: ' + (q ? q.ticketStampAdjustment : '-'));
    Logger.log('    お名前で引いた場合と同じか: ' + (q && q.name === p.name ? 'はい' : '**いいえ**'));
  }
}

// 顧客管理に登録されている回数券カードが、お客様アプリへどう渡るかを見る。
// **読むだけ。**何も書かない。
function 回数券の登録を見る() {
  var profiles = getCustomerProfiles_() || {};
  var 鍵 = Object.keys(profiles);
  var 登録あり = 0;

  Logger.log('■ 顧客管理の回数券カード（' + 鍵.length + '名）');
  Logger.log('');
  鍵.sort(function (a, b) { return a < b ? -1 : 1; }).forEach(function (k) {
    var r = profiles[k] || {};
    var 公開 = publicCustomerProfile_(r);
    var c = 公開 && 公開.activeTicketCard;
    if (c) {
      登録あり += 1;
      Logger.log('  ' + r.name + ': ' + c.plan + '・' + c.sheetLabel + '・' + c.roundLabel +
        '（出どころ: ' + (r.activeTicketCardSource || '不明') + '）');
    } else {
      Logger.log('  ' + r.name + ': 未登録');
    }
  });
  Logger.log('');
  Logger.log('■ 登録あり: ' + 登録あり + '名 / 未登録: ' + (鍵.length - 登録あり) + '名');
  Logger.log('');
  Logger.log('  未登録の方は、お客様アプリで「回数券を追加」のボタンが出ます。');
  Logger.log('  管理アプリの顧客編集で 種類・何枚目・何回目 を入れると、カードが出ます。');
}

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




