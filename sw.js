'use strict';
// オフラインで使うためのサービスワーカー（アプリとしてインストールしたときに働く）。
// アプリのファイルをすべて保存しておき、ネットにつながっていなくても起動できるようにする。
// 更新：このファイルの VERSION が変わると、ブラウザが新しい版をダウンロードし、次に開いたときから新しい版になる。
const VERSION = '2026-10-16.1';
const CACHE = 'bansho-' + VERSION;
const FILES = [
  './',
  './index.html',
  './manifest.webmanifest',
  './css/page.css',
  './css/app.css',
  './js/model.js',
  './js/ink.js',
  './js/rich.js',
  './js/geom.js',
  './js/view.js',
  './js/presenter.js',
  './js/editor.js',
  './js/fileio.js',
  './js/pptx.js',
  './js/backup.js',
  './js/app.js',
  './icons/icon-192.png',
  './icons/icon-512.png',
  './icons/icon-maskable-512.png',
  './icons/favicon-32.png',
];

self.addEventListener('install', (e) => {
  e.waitUntil(
    caches.open(CACHE)
      .then((c) => c.addAll(FILES.map((f) => new Request(f, { cache: 'reload' }))))
      .then(() => self.skipWaiting()),
  );
});

self.addEventListener('activate', (e) => {
  e.waitUntil(
    caches.keys()
      .then((keys) => Promise.all(keys.filter((k) => k.startsWith('bansho-') && k !== CACHE).map((k) => caches.delete(k))))
      .then(() => self.clients.claim()),
  );
});

// 保存してあるファイルを返す（なければネットから取る）。開いている版の途中で新旧が混ざらないよう、
// 新しい版は次に開いたときから使う。
self.addEventListener('fetch', (e) => {
  const req = e.request;
  if (req.method !== 'GET' || new URL(req.url).origin !== self.location.origin) return;
  e.respondWith(
    caches.open(CACHE).then((c) => c.match(req, { ignoreSearch: true }).then((hit) => hit || fetch(req))),
  );
});
