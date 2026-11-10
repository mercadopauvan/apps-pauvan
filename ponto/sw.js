// Ponto Pauvan: service worker simples (deixa instalar o app na tela inicial)
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () { });
