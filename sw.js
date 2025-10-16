// sw.js
// --- 分模組定義 ---
const MODULES = {
  mus1: {
    cache: 'mus1-v1.1',
    assets: [
      './mus/Detektiv.mp3',
      './mus/Prontera.mp3',
      './mus/login1.mp3',
      './mus/login2.mp3'
    ]
  },
  mus2: {
    cache: 'mus2-v1.23',
    assets: [
      './mus/levelup.mp3',
      './mus/clock01.wav',
      './mus/clock02.mp3',
      './mus/clock03.mp3'
    ]
  }
};

// 建立路徑→模組與快取對照
const LOOKUP = (() => {
  const map = new Map();
  for (const m of Object.values(MODULES)) {
    m.assets.forEach(p => map.set(p, m.cache));
  }
  return map;
})();

const ALL_CURRENT_CACHES = new Set(Object.values(MODULES).map(m => m.cache));

// --- install：各模組各自快取 ---
self.addEventListener('install', e => {
  e.waitUntil(
    Promise.all(
      Object.values(MODULES).map(m =>
        caches.open(m.cache).then(c => c.addAll(m.assets))
      )
    )
  );
  self.skipWaiting();
});

// --- activate：只清理「不在清單」的 cache ---
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys
          .filter(k => !ALL_CURRENT_CACHES.has(k))
          .map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// --- fetch：只攔截在 LOOKUP 的資產，並導到對應模組的 cache ---
self.addEventListener('fetch', e => {
  if (e.request.method !== 'GET') return;

  const url = new URL(e.request.url);
  let path = url.pathname;
  // 若你的站點不是根目錄，視情況處理 basePath
  if (path.startsWith(self.registration.scope.replace(self.location.origin, ''))) {
    path = '.' + path.slice(self.registration.scope.replace(self.location.origin, '').length - 1);
  } else {
    path = '.' + path;
  }

  const targetCache = LOOKUP.get(path);
  if (!targetCache) return; // 非受管，放行

  e.respondWith(
    caches.open(targetCache).then(cache =>
      cache.match(e.request).then(hit =>
        hit || fetch(e.request).then(res => {
          if (res.ok) cache.put(e.request, res.clone());
          return res;
        })
      )
    )
  );
});