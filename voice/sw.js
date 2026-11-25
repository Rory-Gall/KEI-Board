/* build 20261125-1037 */
/* KEI Voice - service worker.
 *   - keeps the app itself on the phone, so it opens with no signal
 *   - receives audio shared to it from other apps (a call recording from the
 *     Phone app: Share -> KEI Voice) and queues it
 *   - sends waiting recordings when signal comes back (Background Sync), even
 *     with the app closed
 */
importScripts('config.js', 'send.js');
var CACHE = 'kei-voice-' + BUILD;
var SHELL = ['./', './index.html', './app.js', './send.js', './config.js', './manifest.webmanifest',
             './icon-192.png', './icon-512.png'];
var SCOPE_PATH = new URL(self.registration.scope).pathname;

self.addEventListener('install', function (e) {
  e.waitUntil(caches.open(CACHE).then(function (c) { return c.addAll(SHELL); }).then(function () { return self.skipWaiting(); }));
});

self.addEventListener('activate', function (e) {
  e.waitUntil(caches.keys().then(function (keys) {
    return Promise.all(keys.filter(function (k) { return k.indexOf('kei-voice-') === 0 && k !== CACHE; })
                           .map(function (k) { return caches.delete(k); }));
  }).then(function () { return self.clients.claim(); }));
});

function takeShare(request) {
  return request.formData().then(function (form) {
    var files = form.getAll('audio').filter(function (f) { return f && typeof f === 'object' && f.size; });
    return Promise.all(files.map(function (f) { return VOICE.enqueue(VOICE.sharedItem(f)); }))
      .then(function () { return files.length; });
  }).catch(function () { return 0; }).then(function (n) {
    if (n && self.registration.sync) self.registration.sync.register('kei-voice-send').catch(function () {});
    return Response.redirect(new URL('./?shared=' + n, self.registration.scope).href, 303);
  });
}

self.addEventListener('fetch', function (e) {
  var url = new URL(e.request.url);
  if (e.request.method === 'POST' && url.origin === location.origin && url.pathname === SCOPE_PATH + 'share-target') {
    e.respondWith(takeShare(e.request));
    return;
  }
  if (e.request.method !== 'GET' || url.origin !== location.origin || url.pathname.indexOf(SCOPE_PATH) !== 0) return;
  /* The app's own files: answer from the phone at once (it must open with no
     signal), and refresh the stored copy in the background for next time. */
  e.respondWith(caches.open(CACHE).then(function (c) {
    return c.match(e.request, { ignoreSearch: true }).then(function (hit) {
      var fresh = fetch(e.request).then(function (r) {
        if (r && r.ok && r.type === 'basic') c.put(url.origin + url.pathname, r.clone());
        return r;
      }).catch(function () { return null; });
      e.waitUntil(fresh);
      return hit || fresh.then(function (r) { return r || c.match('./index.html'); });
    });
  }));
});

self.addEventListener('sync', function (e) {
  if (e.tag !== 'kei-voice-send') return;
  e.waitUntil(self.clients.matchAll({ type: 'window' }).then(function (cs) {
    var shown = cs.filter(function (c) { return c.visibilityState === 'visible'; });
    if (shown.length) { shown.forEach(function (c) { c.postMessage({ type: 'send' }); }); return; }
    return VOICE.sendAll().then(function (r) {
      cs.forEach(function (c) { c.postMessage({ type: 'changed' }); });
      /* Still no signal: reject, and Chrome tries again later. A sign-in that has run
         out needs the app opened, so retrying in the background would not help. */
      if (r.left && (r.error === 'offline' || r.error === 'retry')) throw new Error('not sent yet: ' + r.error);
    });
  }));
});
