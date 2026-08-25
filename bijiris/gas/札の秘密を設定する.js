// ビジリスの札の秘密（TOKEN_SECRET）を、その場で作って設定する道具。
//
//   札の秘密を設定する() … 64文字をランダムに作って保存する（**値は表示しない**）
//
// ## なぜ要るのか
//
// `TOKEN_SECRET` が**未設定**だと、`Code.gs` の
// `DEFAULT_TOKEN_SECRET`（公開リポジトリに書かれている文字列）が使われる。
//
// 管理者の札は `sign_("<ユーザー名>|<期限>")` の署名だけで通る。
// `verifyToken_` は署名と期限しか見ておらず、**ユーザー名が実在するかは見ない。**
//
// つまり秘密が既定のままだと、**公開コードを読んだ人が誰でも
// 管理者の札を自分で作れる。**パスワードは要らない。
// その札で 29 か所の管理操作（お客様の氏名・計測値・写真など）が通る。
//
// 2026-08-25 に、未設定であることを確認した。
// （2026-08-11 の記録では「設定済み」となっていたが、本番はそうなっていなかった）
//
// ## 実行するとどうなるか
//
//   ・**いま管理アプリに入っている方は、入口に戻されます。**
//     もう一度ログインしていただければ入れます。お客様側には影響しません。
//   ・秘密は画面にもログにも出しません。作って、そのまま保存します。
//   ・すでに設定されている場合は、**何もしません**（上書きしません）。
//
// ## まゆみ側の ADMIN_TOKEN_SECRET とは別物です
//
// **まゆみ側は絶対に変えないでください**（お客様全員が入口に戻されます）。
// ここで触るのはビジリスの TOKEN_SECRET だけです。

function 札の秘密を設定する() {
  var props = PropertiesService.getScriptProperties();
  var いま = props.getProperty('TOKEN_SECRET');

  if (いま && String(いま).length) {
    Logger.log('■ すでに設定されています（' + String(いま).length + '文字）。');
    Logger.log('   **何もしませんでした。**上書きはしません。');
    if (String(いま) === String(DEFAULT_TOKEN_SECRET)) {
      Logger.log('');
      Logger.log('   ただし **コードの既定値と同じ**です。');
      Logger.log('   作り直す場合は、プロジェクトの設定から TOKEN_SECRET を一度削除してから、');
      Logger.log('   もう一度この関数を実行してください。');
    }
    return;
  }

  // 英数字だけにする。記号を混ぜると、設定画面での取り回しで事故が起きうる。
  var 文字 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var 秘密 = '';
  for (var i = 0; i < 64; i++) {
    秘密 += 文字.charAt(Math.floor(Math.random() * 文字.length));
  }

  props.setProperty('TOKEN_SECRET', 秘密);

  // **値は絶対に出さない。**出した時点で漏れる。
  var 確認 = props.getProperty('TOKEN_SECRET');
  Logger.log('■ 設定しました。');
  Logger.log('   文字数: ' + (確認 ? String(確認).length : 0));
  Logger.log('   コードの既定値と同じ: ' + (String(確認) === String(DEFAULT_TOKEN_SECRET) ? 'はい（**失敗**）' : 'いいえ'));
  Logger.log('');
  Logger.log('   **いま管理アプリに入っている方は、入口に戻されます。**');
  Logger.log('   もう一度ログインしていただければ入れます。お客様側には影響しません。');
  Logger.log('');
  Logger.log('   確かめるときは 秘密が既定のままか確かめる() を実行してください。');
}
