// Service Worker для кэширования и работы офлайн
const CACHE_NAME = 'smeta-pwa-v1';
const ASSETS_TO_CACHE = [
    '/',
    '/index.html',
    '/static/manifest.json'
];

// Установка SW и кэширование
self.addEventListener('install', (event) => {
    event.waitUntil(
        caches.open(CACHE_NAME).then((cache) => {
            console.log('Кэширование ассетов');
            return cache.addAll(ASSETS_TO_CACHE);
        })
    );
    self.skipWaiting();
});

// Активация и очистка старого кэша
self.addEventListener('activate', (event) => {
    event.waitUntil(
        caches.keys().then((cacheNames) => {
            return Promise.all(
                cacheNames.map((cacheName) => {
                    if (cacheName !== CACHE_NAME) {
                        console.log('Удаление старого кэша:', cacheName);
                        return caches.delete(cacheName);
                    }
                })
            );
        })
    );
    self.clients.claim();
});

// Перехват запросов (Network First, fallback to Cache)
self.addEventListener('fetch', (event) => {
    // Не обрабатываем POST запросы через кэш (для загрузки файлов)
    if (event.request.method === 'POST') {
        return;
    }

    event.respondWith(
        fetch(event.request)
            .then((response) => {
                // Клонируем ответ, чтобы положить в кэш и вернуть пользователю
                const responseClone = response.clone();
                caches.open(CACHE_NAME).then((cache) => {
                    cache.put(event.request, responseClone);
                });
                return response;
            })
            .catch(() => {
                // Если сети нет, пытаемся взять из кэша
                return caches.match(event.request);
            })
    );
});
