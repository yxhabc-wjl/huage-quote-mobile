/* 华哥报价手机版 Service Worker：App Shell 离线缓存（V1.1）
 * 发新版时必须把 CACHE 版本号 +1（如 wujin-shell-v1.1），否则旧缓存不更新 */
const CACHE = "wujin-shell-v1.1";
const ASSETS = [
  "./", "./index.html", "./engine.js", "./manifest.webmanifest",
  "./icon-192.png", "./icon-512.png", "./apple-touch-icon.png"
];
self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(CACHE).then((c) => c.addAll(ASSETS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys()
      .then((ks) => Promise.all(ks.filter((k) => k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
self.addEventListener("fetch", (e) => {
  if (e.request.method !== "GET") return;
  e.respondWith(
    caches.match(e.request).then((hit) => {
      if (hit) return hit;
      return fetch(e.request).then((res) => {
        const cp = res.clone();
        caches.open(CACHE).then((c) => c.put(e.request, cp));
        return res;
      }).catch(() => caches.match("./index.html"));
    })
  );
});
