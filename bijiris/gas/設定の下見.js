// 管理アプリの設定が、いまサーバーでどうなっているかを見る道具。**読むだけ。**
//
// 特典の保存ができない原因を切り分けるために作った。
// 保存は「送って、読み直して一致を確かめる」作りだが、送信は応答を読まないため、
// サーバー側で例外が出ても画面には「確認できませんでした」としか出ない。

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
