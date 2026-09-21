// 旧管理アプリの「覚えさせる仕掛け」は役目を終えました（2026-09-21）。
//
// 管理は新しい管理画面（サーバー側）へ移りました。
// この仕掛けが残っていると、**差し替えたあとも古い管理アプリの画面が出続けます。**
// ブラウザは、開くたびにこのファイルを見に来ます。そのときに自分を外し、
// 覚えさせていたものを消して、開いているページを読み直させます。
//
// **お客様アプリのものには触りません。**同じ住所の下に同居しているので、
// 消すのは名前が mayumi-admin- で始まるものだけにします。

self.addEventListener('install', function (event) {
  self.skipWaiting();
});

self.addEventListener('activate', function (event) {
  event.waitUntil((async function () {
    try {
      const 名前たち = await caches.keys();
      await Promise.all(
        名前たち
          .filter(function (k) { return k.indexOf('mayumi-admin-') === 0; })
          .map(function (k) { return caches.delete(k); })
      );
    } catch (e) { /* 消せなくても、下の外す処理は進める */ }

    try {
      await self.registration.unregister();
    } catch (e) { /* 外せなくても、下の読み直しは進める */ }

    // 開いているページを読み直させる。新しい「移転しました」の画面が出る。
    try {
      const ページたち = await self.clients.matchAll({ type: 'window' });
      ページたち.forEach(function (c) {
        if (String(c.url || '').indexOf('/admin') !== -1) c.navigate(c.url);
      });
    } catch (e) { /* 読み直せなくても、次に開いたときには新しい画面になる */ }
  })());
});

// 何も横取りしない。いつもインターネットから取りに行く。
