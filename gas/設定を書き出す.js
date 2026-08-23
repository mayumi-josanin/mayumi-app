// アプリの設定（スクリプトプロパティ）を、サーバーへ移せる形で書き出す。
//
//   設定を書き出す()   … いまの設定を JSON にしてドライブへ置く
//
// **読むだけ。**設定は変えない。
//
// なぜ要るのか:
//   `getAppRuntimeConfig` と `getRewardGachaConfig` は、
//   **表ではなくスクリプトプロパティ**から読んでいる。
//   段階Bでこの2つの口をサーバーに作るには、まず中身を移す必要がある。
//
// **秘密は書き出さない。**
//   スクリプトプロパティには ADMIN_TOKEN_SECRET など、
//   外に出してはいけないものも入っている。
//   **移すと決めた2つだけを、名指しで書き出す。**
//   「全部書き出して後で選ぶ」はしない。書き出した時点で漏れる。

var 設定書出_移すもの = [
  { 鍵: APP_RUNTIME_CONFIG_PROPERTY,  名: 'appRuntime',  既定: DEFAULT_APP_RUNTIME_CONFIG },
  { 鍵: REWARD_GACHA_CONFIG_PROPERTY, 名: 'rewardGacha', 既定: null }
];

function 設定を書き出す() {
  var props = PropertiesService.getScriptProperties();

  var 中 = {};
  設定書出_移すもの.forEach(function (x) {
    var 生 = props.getProperty(x.鍵);
    var 値 = null;
    var 備考 = '';
    if (!生) {
      // **保存されていないときは、GASも既定値を返している。**
      // 空のまま移すと、サーバーだけ設定が消えたことになる。
      値 = x.既定;
      備考 = '保存されていないので既定値';
    } else {
      try {
        値 = JSON.parse(生);
      } catch (e) {
        値 = null;
        備考 = '**読めませんでした: ' + e + '**';
      }
    }
    中[x.名] = { key: x.鍵, value: 値, note: 備考 };
  });

  // GAS が実際に返している形も一緒に入れる。
  // **保存された値と、返る値は違う。**sanitize が既定値で補うため。
  中.actualGetAppRuntimeConfig = getAppRuntimeConfig();
  中.actualGetRewardGachaConfig = getRewardGachaConfig();

  var 本文 = JSON.stringify({
    書き出した日時: Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ssXXX"),
    settings: 中
  }, null, 2);

  var 名前 = '設定_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmm') + '.json';
  var file = DriveApp.createFile(名前, 本文, MimeType.PLAIN_TEXT);

  Logger.log('■ 書き出しました: ' + 名前);
  Logger.log('  ' + file.getUrl());
  Logger.log('  大きさ: ' + Math.round(file.getSize() / 1024 * 10) / 10 + 'KB');
  Logger.log('');

  設定書出_移すもの.forEach(function (x) {
    var e = 中[x.名];
    Logger.log('■ ' + x.名 + '（' + x.鍵 + '）');
    if (e.note) Logger.log('    ' + e.note);
    if (e.value && typeof e.value === 'object') {
      Object.keys(e.value).forEach(function (k) {
        var v = e.value[k];
        var t = (v && typeof v === 'object') ? JSON.stringify(v) : String(v);
        if (t.length > 60) t = t.slice(0, 60) + '…';
        Logger.log('      ' + k + ' = ' + t);
      });
    } else {
      Logger.log('      （中身がありません）');
    }
    Logger.log('');
  });

  Logger.log('■ GAS が実際に返している形');
  ['actualGetAppRuntimeConfig', 'actualGetRewardGachaConfig'].forEach(function (k) {
    var r = 中[k];
    var c = r && r.config;
    Logger.log('  ' + k + ': status=' + (r && r.status) +
      ' 項目' + (c ? Object.keys(c).length : 0) + '個' +
      (c ? '  ' + Object.keys(c).join(' / ') : ''));
  });
  Logger.log('');

  Logger.log('  ※ 読むだけです。設定は変えていません。');
  Logger.log('  ※ **移すと決めた2つだけを書き出しています。**');
  Logger.log('     ADMIN_TOKEN_SECRET などの秘密は含まれません。');
  Logger.log('     取り込みが済んだら 書き出しJSONを片付ける() で消してください。');
}
