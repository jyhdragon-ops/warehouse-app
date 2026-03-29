// 입출고대장 Service Worker
const CACHE_NAME = 'ipchulgo-v2';
const ASSETS = [
  './',
  './입출고대장_앱_v2.html',
  './manifest.json',
  './icon-192.svg',
  './icon-512.svg',
  'https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;600;700&display=swap'
];

// 설치: 핵심 파일 캐시
self.addEventListener('install', event => {
  event.waitUntil(
    caches.open(CACHE_NAME).then(cache => {
      console.log('[SW] 캐시 저장 중...');
      return cache.addAll(ASSETS.filter(url => !url.startsWith('https://fonts')));
    })
  );
  self.skipWaiting();
});

// 활성화: 이전 캐시 정리
self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys().then(keys =>
      Promise.all(
        keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
      )
    )
  );
  self.clients.claim();
});

// 요청 가로채기: 캐시 우선, 없으면 네트워크
self.addEventListener('fetch', event => {
  // 외부 API(Google Sheets 등)는 항상 네트워크 사용
  if (event.request.url.includes('googleapis.com') ||
      event.request.url.includes('script.google.com') ||
      event.request.method !== 'GET') {
    return;
  }

  event.respondWith(
    caches.match(event.request).then(cached => {
      if (cached) return cached;

      return fetch(event.request).then(response => {
        // 유효한 응답만 캐시
        if (!response || response.status !== 200 || response.type !== 'basic') {
          return response;
        }
        const responseClone = response.clone();
        caches.open(CACHE_NAME).then(cache => cache.put(event.request, responseClone));
        return response;
      }).catch(() => {
        // 오프라인 시 HTML 요청은 메인 페이지 반환
        if (event.request.destination === 'document') {
          return caches.match('./입출고대장_앱_v2.html');
        }
      });
    })
  );
});