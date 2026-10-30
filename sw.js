// Service worker do Apps Pauvan: permite instalar como aplicativo e recebe as notificações (Firebase). Não guarda dados.
importScripts('https://www.gstatic.com/firebasejs/10.12.2/firebase-app-compat.js', 'https://www.gstatic.com/firebasejs/10.12.2/firebase-messaging-compat.js');
firebase.initializeApp({"apiKey": "AIzaSyC4lagk9rwIo-_G5sZxV0GPyhcaTvJz0ng", "authDomain": "apps-pauvan.firebaseapp.com", "projectId": "apps-pauvan", "storageBucket": "apps-pauvan.firebasestorage.app", "messagingSenderId": "442345355000", "appId": "1:442345355000:web:b311386f57ad07527dad0d"});
var messaging = firebase.messaging();
self.addEventListener('install', function () { self.skipWaiting(); });
self.addEventListener('activate', function (e) { e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function () {});
// clique nas notificações mostradas com o app aberto (as do Firebase o próprio SDK abre)
self.addEventListener('notificationclick', function (e) {
  var d = e.notification.data || {};
  if (d.FCM_MSG) return;
  e.notification.close();
  var url = self.registration.scope + (d.abrir ? '?abrir=' + d.abrir : '');
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(function (cs) {
    for (var i = 0; i < cs.length; i++) { if (cs[i].url.indexOf(self.registration.scope) === 0 && 'focus' in cs[i]) { if (d.abrir && 'navigate' in cs[i]) cs[i].navigate(url); return cs[i].focus(); } }
    return self.clients.openWindow(url);
  }));
});
