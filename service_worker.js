const CACHE = 'siksucasa-v1';
const ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png',
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.19.0/dist/tabler-icons.min.css',
  'https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.19.0/fonts/tabler-icons.woff2'
];

/* 설치: 핵심 파일 캐시 */
self.addEventListener('install', function(e){
  e.waitUntil(
    caches.open(CACHE).then(function(cache){
      return cache.addAll(ASSETS);
    })
  );
  self.skipWaiting();
});

/* 활성화: 이전 캐시 정리 */
self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(keys){
      return Promise.all(
        keys.filter(function(k){ return k !== CACHE; })
            .map(function(k){ return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

/* 요청 가로채기: 캐시 우선, 없으면 네트워크 */
self.addEventListener('fetch', function(e){
  e.respondWith(
    caches.match(e.request).then(function(cached){
      return cached || fetch(e.request).then(function(res){
        /* CDN 리소스는 캐시에 추가 저장 */
        if(res && res.status === 200 && e.request.method === 'GET'){
          var resClone = res.clone();
          caches.open(CACHE).then(function(cache){
            cache.put(e.request, resClone);
          });
        }
        return res;
      });
    }).catch(function(){
      /* 오프라인 fallback */
      return caches.match('./index.html');
    })
  );
});
