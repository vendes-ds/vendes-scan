var CACHE = 'vendes-scan-live-v1';
var CACHE_FAMILY = /^vendes-scan-live-|^vendes-scan-v\d+$/;
var SHELL = [
  './',
  './index.html',
  './styles.css',
  './app.js',
  './logic.js',
  './pwa-keys.js',
  './config.js',
  './manifest.webmanifest',
  './icon-192.png',
  './icon-512.png',
  './apple-touch-icon.png',
  './brand/vendes-logo.png',
  './vendor/html5-qrcode.min.js',
  './vendor/xlsx.mini.min.js'
];

self.addEventListener('install', function (event) {
  event.waitUntil(caches.open(CACHE).then(function (cache) {
    return cache.addAll(SHELL);
  }).then(function () {
    return self.skipWaiting();
  }));
});

self.addEventListener('activate', function (event) {
  event.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (key) {
      return key !== CACHE && CACHE_FAMILY.test(key);
    }).map(function (key) {
      return caches.delete(key);
    }));
  }).then(function () {
    return self.clients.claim();
  }));
});

self.addEventListener('fetch', function (event) {
  var url = new URL(event.request.url);
  if (event.request.method !== 'GET' || url.origin !== self.location.origin) return;
  event.respondWith(fetch(event.request).then(function (response) {
    var copy = response.clone();
    caches.open(CACHE).then(function (cache) { cache.put(event.request, copy); });
    return response;
  }).catch(function () {
    return caches.match(event.request).then(function (hit) {
      return hit || caches.match('./index.html');
    });
  }));
});
