// sw.js
const CACHE = 'app-v1.4';
const ASSETS = [
  './mus/Prontera.mp3',
  './mus/levelup.wav',
  './mus/login1.mp3',
  './mus/login2.wav',
  './mus/login3.mp3'
];

// 安裝階段：預先緩存指定檔案
self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

// 啟用階段：清理舊版快取
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// 取用階段：僅攔截 ASSETS 內的檔案
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;

  const url = new URL(req.url);
  const path = '.' + url.pathname.replace(self.location.origin, '');

  if (!ASSETS.includes(path)) return; // 其它請求放行

  e.respondWith(
    caches.match(req).then(hit =>
      hit ||
      fetch(req).then(res => {
        if (res.ok) {
          const copy = res.clone();
          caches.open(CACHE).then(c => c.put(req, copy));
        }
        return res;
      })
    )
  );
});