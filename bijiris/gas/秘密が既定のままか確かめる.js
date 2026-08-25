// 秘密が「コードに書かれた既定値のまま」になっていないかを確かめる道具。**読むだけ。**
//
//   秘密が既定のままか確かめる() … 設定の有無と、既定と同じかどうかだけを出す
//
// ## なぜ要るのか
//
// `Code.gs` の冒頭に、管理者のユーザー名・パスワードと札の秘密が
// **既定値として直接書かれている。**
//
//   var DEFAULT_ADMIN_USERNAME = "…";
//   var DEFAULT_ADMIN_PASSWORD = "…";
//   var DEFAULT_TOKEN_SECRET   = "…";
//
// スクリプトプロパティ（ADMIN_USERNAME / ADMIN_PASSWORD / TOKEN_SECRET）が
// 設定されていれば、そちらが使われる。**設定されていなければ、
// コードに書かれた値がそのまま生きている。**
//
// そして **このリポジトリは GitHub で公開されている**（2026-08-25 に確認）。
// 既定のままなら、**誰でも管理アプリに入れるし、札も偽造できる。**
//
// ## この道具は値を出さない
//
// 2026-08-16 に、確認のつもりで秘密を画面に出してしまったことがある。
// **出すのは「設定されているか」「既定と同じか」「文字数」だけ。**
// 中身は絶対にログへ書かない。

function 秘密が既定のままか確かめる() {
  var props = PropertiesService.getScriptProperties();

  var 見るもの = [
    { 鍵: 'ADMIN_USERNAME', 既定: DEFAULT_ADMIN_USERNAME, 名: '管理者のユーザー名' },
    { 鍵: 'ADMIN_PASSWORD', 既定: DEFAULT_ADMIN_PASSWORD, 名: '管理者のパスワード' },
    { 鍵: 'TOKEN_SECRET',   既定: DEFAULT_TOKEN_SECRET,   名: '札の秘密' }
  ];

  Logger.log('■ 秘密の設定ぐあい（**値は出しません**）');
  Logger.log('');

  var 危ない = 0;

  見るもの.forEach(function (m) {
    var v = props.getProperty(m.鍵);
    var ある = !!(v && String(v).length);
    var 既定と同じ = ある ? (String(v) === String(m.既定)) : true;

    Logger.log('   ' + m.名 + '（' + m.鍵 + '）');
    Logger.log('       設定されている: ' + (ある ? 'はい（' + String(v).length + '文字）' : '**いいえ**'));
    Logger.log('       コードの既定値と同じ: ' + (既定と同じ ? '**はい**' : 'いいえ'));

    if (既定と同じ) {
      危ない++;
      Logger.log('       → **公開されているコードを読めば分かる状態です。**');
    }
    Logger.log('');
  });

  // 管理者の一覧が別に保存されている場合も見る（移行で作られる）。
  var 一覧 = props.getProperty('ADMIN_ACCOUNTS_JSON');
  if (一覧) {
    var n = 0;
    var 既定パスワードの人 = 0;
    try {
      var a = JSON.parse(一覧);
      if (Array.isArray(a)) {
        n = a.length;
        a.forEach(function (x) {
          if (x && String(x.password || '') === String(DEFAULT_ADMIN_PASSWORD)) 既定パスワードの人++;
        });
      }
    } catch (e) { }
    Logger.log('   管理者の一覧（ADMIN_ACCOUNTS_JSON）: ' + n + '人');
    Logger.log('       そのうちパスワードが既定のまま: ' + 既定パスワードの人 + '人');
    if (既定パスワードの人) 危ない++;
    Logger.log('');
  }

  Logger.log('■ まとめ');
  if (危ない) {
    Logger.log('   **' + 危ない + '件が、公開コードから分かる状態です。**');
    Logger.log('   スクリプトプロパティに別の値を設定してください。');
    Logger.log('   （プロジェクトの設定 → スクリプト プロパティ）');
    Logger.log('');
    Logger.log('   ※ TOKEN_SECRET を変えると、**いま入っている方が入口に戻されます。**');
    Logger.log('      まゆみ側の ADMIN_TOKEN_SECRET とは別物です。混同しないこと。');
  } else {
    Logger.log('   すべて既定値から変えられています。');
  }
}
