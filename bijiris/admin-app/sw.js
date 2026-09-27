// 旧ビジリス管理アプリを閉じるための Service Worker（2026-09-27）。
//
// 前の版は画面を端末に控えていたので、ページを差し替えただけでは古い画面が出続ける。
// この版は、入ったらすぐ前の控えを消し、開いている画面を読み込み直して「移りました」の案内を出す。
// **消すのは管理アプリの控え（mayumi-admin-survey-…）だけ。**同じ住所にあるお客様のアプリの控えは消さない。
const 消す頭 = "mayumi-admin-survey";

self.addEventListener("install", () => self.skipWaiting());

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((k) => k.startsWith(消す頭)).map((k) => caches.delete(k)));
    await self.clients.claim();
    const clients = await self.clients.matchAll({ type: "window" });
    clients.forEach((c) => { try { c.navigate(c.url); } catch (e) { /* 読み込み直せなくても次に開いたとき案内が出る */ } });
  })());
});

// 控えは使わず、いつも新しいものを取りに行く
self.addEventListener("fetch", () => {});
