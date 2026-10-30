// Service worker mínimo: permite instalar como aplicativo. Não guarda dados (sempre busca da internet).
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () {});
