// お客様アプリと管理アプリで、お知らせの件数が食い違う原因を突き止める道具。
//
//   お知らせの食い違いを調べる()   … 読むだけ。**何も書きません**
//
// 2026-09-13、「9/5以降のNEWSがお客様アプリには出るのに管理アプリに無い」
// と院長からご報告。9/5 はお知らせの表をサーバーへ切り替えた日。
//
// 疑っているのはこれ:
//   お客様向け getNews        → 合鍵なしの窓口 → サーバー   ← 見えている
//   管理向け   getAdminBlogs  → **合鍵が要る窓口** → サーバー ← ここが弾かれると
//                                                     GAS はシート（9/5で止まっている）へ落ちる
//
// 落ちてもエラーにならず古い一覧が出る作りなので、外からは見分けがつかない。
// 転送の失敗は Logger.log にしか残らない。だからエディタから直接叩いて見る。
//
// **管理アプリのキャッシュ（15分）の可能性も、この結果で切り分けられる。**
// ここで getAdminBlogs が最新を返すなら、原因は管理アプリ側のキャッシュ。

function お知らせの食い違いを調べる() {
  Logger.log('■ 設定');
  Logger.log('  渡している表 : ' + (渡す_設定_('SERVER_TABLES', '') || '**なし**'));
  Logger.log('  合鍵         : ' + (渡す_設定_('SERVER_API_KEY', '') ? '設定あり' : '**未設定**'));
  Logger.log('  news を渡すか: ' + サーバーへ渡すか_('news'));
  Logger.log('');

  // ① 管理向けの窓口を、GAS からサーバーへ直接（管理者の札は要らない。GASの中なので）
  Logger.log('■ サーバーへ直接: getAdminBlogs（管理向け・合鍵あり）');
  調べ_窓口_('getAdminBlogs', 'blogs');
  Logger.log('');
  Logger.log('■ サーバーへ直接: getNews（お客様向け・合鍵なし）');
  調べ_窓口_('getNews', 'news');
  Logger.log('');

  // ② GAS の関数を通した結果（管理アプリが実際に受け取るもの）
  Logger.log('■ GAS の getAdminBlogs() を通した結果（管理アプリが受け取るもの）');
  try {
    var r = getAdminBlogs();
    var b = (r && r.blogs) || [];
    Logger.log('  status=' + (r && r.status) + ' / ' + b.length + '件');
    調べ_新しい順_(b, 'date');
  } catch (e) {
    Logger.log('  **例外**: ' + e);
  }
  Logger.log('');
  Logger.log('  → 直接叩いた件数と、ここの件数が違えば、GAS がシートへ落ちています。');
  Logger.log('  → 同じなら、原因は管理アプリ側のキャッシュ（15分）です。');
}

function 調べ_窓口_(action, 鍵) {
  try {
    var url = 渡す_URL_() + '?action=' + encodeURIComponent(action);
    var res = UrlFetchApp.fetch(url, {
      method: 'get', headers: 渡す_ヘッダ_(), muteHttpExceptions: true,
      followRedirects: true, validateHttpsCertificates: true
    });
    var code = res.getResponseCode();
    var 文 = res.getContentText();
    Logger.log('  HTTP ' + code);
    if (code !== 200) { Logger.log('  中身: ' + 文.slice(0, 200)); return; }
    var 中 = JSON.parse(文);
    var 一覧 = (中 && 中[鍵]) || [];
    Logger.log('  status=' + (中 && 中.status) + ' / ' + 一覧.length + '件' +
      (中 && 中.message ? ' / ' + 中.message : ''));
    調べ_新しい順_(一覧, 'date');
  } catch (e) {
    Logger.log('  **失敗**: ' + e);
  }
}

function 調べ_新しい順_(一覧, 日付の鍵) {
  var 並び = 一覧.slice().sort(function (a, b) {
    return String(b[日付の鍵] || '').localeCompare(String(a[日付の鍵] || ''));
  });
  並び.slice(0, 5).forEach(function (x) {
    Logger.log('    ' + String(x[日付の鍵] || '').slice(0, 10) + ' | ' + String(x.title || '').slice(0, 28));
  });
}
