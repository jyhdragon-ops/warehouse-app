const CACHE_NAME = 'warehouse-v1';
const ASSETS = [
  '/warehouse-app/',
  '/warehouse-app/index.html',
  '/warehouse-app/manifest.json',
  '/warehouse-app/icon-192.svg',
  '/warehouse-app/icon-512.svg'
];

// 설치 — 앱 셸 캐시
self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE_NAME).then(cache => cache.addAll(ASSETS))
  );
  self.skipWaiting();
});

// 활성화 — 오래된 캐시 제거
self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// 요청 처리 — 네트워크 우선, 실패 시 캐시
self.addEventListener('fetch', e => {
  // Google Apps Script 요청은 캐시하지 않음
  if (e.request.url.includes('script.google.com')) return;

  e.respondWith(
    fetch(e.request)
      .then(res => {
        // 성공하면 캐시 업데이트
        const clone = res.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(e.request, clone));
        return res;
      })
      .catch(() => caches.match(e.request))
  );
});
