/* KEI Voice - storage, sign-in tokens and sending. Shared by the page (app.js)
 * and the service worker (sw.js), so it uses only what both have: IndexedDB,
 * fetch and crypto. No libraries.
 *
 * IndexedDB "kei-voice":
 *   chunks  {rec, seq, blob}  - a recording as it is made, 2 s at a time, so a
 *                               crash or a killed app loses at most 2 seconds
 *   queue   {id, name, kind, blob, mime, created, ms, state, tries, lastError, sentAt}
 *                             - finished recordings and shared files; state is
 *                               "waiting" until OneDrive has the whole file, then "sent"
 *   kv      {k, v}            - "auth" (tokens) and "active" (the recording in progress)
 *
 * Sending = a Microsoft Graph upload session into the person's own OneDrive,
 * CONFIG.folder/<name>. The PC watcher picks it up from there.
 */
var BUILD = '20261125-1037';
var VOICE = (function () {
  'use strict';

  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var AUTHORITY = 'https://login.microsoftonline.com/' + CONFIG.tenantId + '/oauth2/v2.0';
  /* The same delegated scopes the KEI board already has consented. Sites.ReadWrite.All
     covers the person's own OneDrive (tested 2026-11-25), so no new consent is needed. */
  var SCOPES = 'https://graph.microsoft.com/Sites.ReadWrite.All https://graph.microsoft.com/User.Read offline_access openid profile';
  var CHUNK = 10 * 327680;          /* 3.2 MB - Graph wants multiples of 320 KiB */
  var AUDIO_EXT = ['.webm', '.m4a', '.mp3', '.wav', '.ogg', '.opus', '.amr', '.3gp', '.aac', '.mp4', '.flac'];

  function err(code, message) { var e = new Error(message); e.code = code; return e; }

  /* ---------------------------------------------------------------- storage */
  var dbp = null;
  function db() {
    if (dbp) return dbp;
    dbp = new Promise(function (res, rej) {
      var r = indexedDB.open('kei-voice', 1);
      r.onupgradeneeded = function () {
        var d = r.result;
        if (!d.objectStoreNames.contains('chunks')) d.createObjectStore('chunks', { keyPath: ['rec', 'seq'] });
        if (!d.objectStoreNames.contains('queue')) d.createObjectStore('queue', { keyPath: 'id' });
        if (!d.objectStoreNames.contains('kv')) d.createObjectStore('kv', { keyPath: 'k' });
      };
      r.onsuccess = function () { res(r.result); };
      r.onerror = function () { dbp = null; rej(r.error || err('storage', 'could not open storage')); };
    });
    return dbp;
  }
  /* One transaction; resolves with whatever fn hands to set() once it has COMMITTED. */
  function run(name, mode, fn) {
    return db().then(function (d) {
      return new Promise(function (res, rej) {
        var t = d.transaction(name, mode), out;
        fn(t.objectStore(name), function (v) { out = v; });
        t.oncomplete = function () { res(out); };
        t.onerror = function () { rej(t.error || err('storage', 'storage write failed')); };
        t.onabort = function () { rej(t.error || err('storage', 'storage write aborted')); };
      });
    });
  }
  function range(rec) { return IDBKeyRange.bound([rec, 0], [rec, Infinity]); }

  function kvGet(k) {
    return run('kv', 'readonly', function (s, set) {
      var r = s.get(k); r.onsuccess = function () { set(r.result ? r.result.v : null); };
    });
  }
  function kvSet(k, v) { return run('kv', 'readwrite', function (s) { s.put({ k: k, v: v }); }); }
  function kvDel(k) { return run('kv', 'readwrite', function (s) { s.delete(k); }); }

  function putChunk(rec, seq, blob) {
    return run('chunks', 'readwrite', function (s) { s.put({ rec: rec, seq: seq, blob: blob }); });
  }
  function chunksOf(rec) {
    return run('chunks', 'readonly', function (s, set) {
      var r = s.getAll(range(rec));
      r.onsuccess = function () {
        set(r.result.sort(function (a, b) { return a.seq - b.seq; }).map(function (x) { return x.blob; }));
      };
    });
  }
  function dropChunks(rec) { return run('chunks', 'readwrite', function (s) { s.delete(range(rec)); }); }
  function chunkRecs() {
    return run('chunks', 'readonly', function (s, set) {
      var r = s.getAllKeys();
      r.onsuccess = function () {
        var seen = {};
        r.result.forEach(function (k) { seen[k[0]] = true; });
        set(Object.keys(seen));
      };
    });
  }

  function enqueue(item) { return run('queue', 'readwrite', function (s) { s.put(item); }); }
  function items() {
    return run('queue', 'readonly', function (s, set) {
      var r = s.getAll();
      r.onsuccess = function () { set(r.result.sort(function (a, b) { return a.created - b.created; })); };
    });
  }
  function update(id, patch) {
    return run('queue', 'readwrite', function (s) {
      var r = s.get(id);
      r.onsuccess = function () { if (r.result) s.put(Object.assign(r.result, patch)); };
    });
  }
  /* Keep a week of sent records for the Recent list; their audio is already gone. */
  function prune() {
    var cutoff = Date.now() - 7 * 864e5;
    return run('queue', 'readwrite', function (s) {
      var r = s.openCursor();
      r.onsuccess = function () {
        var c = r.result;
        if (!c) return;
        if (c.value.state === 'sent' && (c.value.sentAt || 0) < cutoff) c.delete();
        c.continue();
      };
    });
  }

  /* ------------------------------------------------------------------ names */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  /* Local time, as the PC side expects: YYYY-MM-DD_HHMMSS */
  function stamp(ms) {
    var d = new Date(ms || Date.now());
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + '_' +
      pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  }
  function rid() {
    var a = new Uint8Array(3);
    crypto.getRandomValues(a);
    return Array.prototype.map.call(a, function (b) { return pad(b.toString(16)).slice(-2); }).join('').slice(0, 5);
  }
  function extFor(mime, filename) {
    var m = /\.[A-Za-z0-9]{2,5}$/.exec(filename || '');
    if (m && AUDIO_EXT.indexOf(m[0].toLowerCase()) >= 0) return m[0].toLowerCase();
    mime = (mime || '').toLowerCase();
    if (mime.indexOf('webm') >= 0) return '.webm';
    if (mime.indexOf('ogg') >= 0) return '.ogg';
    if (mime.indexOf('mpeg') >= 0 || mime.indexOf('mp3') >= 0) return '.mp3';
    if (mime.indexOf('wav') >= 0) return '.wav';
    if (mime.indexOf('amr') >= 0) return '.amr';
    if (mime.indexOf('3gp') >= 0) return '.3gp';
    if (mime.indexOf('aac') >= 0) return '.aac';
    return '.m4a';      /* what Android call recorders usually make; the PC decodes by content anyway */
  }
  function memoItem(id, blob, mime, started, ms) {
    return { id: id, name: id + extFor(mime), kind: 'memo', blob: blob, mime: mime, created: started,
             ms: ms, state: 'waiting', tries: 0, lastError: '' };
  }
  function newMemoId(started) { return stamp(started) + '_memo_' + rid(); }
  function sharedItem(file) {
    var when = file.lastModified || Date.now();
    var stem = (file.name || 'shared audio').replace(/\.[^.]*$/, '').replace(/[^A-Za-z0-9 ._-]+/g, ' ')
      .replace(/\s+/g, ' ').trim().slice(0, 60);
    var id = stamp(when) + '_shared_' + rid() + (stem ? '_' + stem : '');
    return { id: id, name: id + extFor(file.type, file.name), kind: 'shared', blob: file, mime: file.type || '',
             created: when, ms: 0, state: 'waiting', tries: 0, lastError: '', original: file.name || '' };
  }

  /* ------------------------------------------------------------------ tokens */
  function form(url, data) {
    var body = Object.keys(data).map(function (k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(data[k]);
    }).join('&');
    return fetch(url, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: body })
      .then(function (r) { return r.json(); });
  }
  function identity(tok) {
    var out = { name: '', email: '' };
    if (!tok.id_token) return out;
    try {
      var p = tok.id_token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      var c = JSON.parse(decodeURIComponent(escape(atob(p + '==='.slice((p.length + 3) % 4)))));
      out.name = c.name || '';
      out.email = (c.preferred_username || c.upn || '').toLowerCase();
    } catch (e) { /* keep the previous identity */ }
    return out;
  }
  function saveAuth(tok, prev) {
    prev = prev || {};
    var who = identity(tok);
    var s = {
      access_token: tok.access_token,
      refresh_token: tok.refresh_token || prev.refresh_token || '',
      expires_at: Date.now() + Number(tok.expires_in || 3600) * 1000,
      name: who.name || prev.name || '',
      email: who.email || prev.email || ''
    };
    return kvSet('auth', s).then(function () { return s; });
  }
  /* A valid access token, refreshed silently. Rejects with code "signin" when the
     person has to sign in again (Microsoft limits a web app's refresh token to 24 h),
     or "offline" when there is no connection - which is NOT a sign-out. */
  var refreshing = null;
  function token() {
    if (refreshing) return refreshing;
    refreshing = kvGet('auth').then(function (s) {
      if (!s || !s.email && !s.refresh_token && !s.access_token) throw err('signin', 'not signed in yet');
      if (s.access_token && s.expires_at - 120000 > Date.now()) return s.access_token;
      if (!s.refresh_token) throw err('signin', 'sign-in has run out');
      return form(AUTHORITY + '/token', {
        client_id: CONFIG.clientId, grant_type: 'refresh_token', refresh_token: s.refresh_token, scope: SCOPES
      }).then(function (tok) {
        if (tok.access_token) return saveAuth(tok, s).then(function (n) { return n.access_token; });
        /* keep who it was, so the next sign-in can be silent */
        return kvSet('auth', { email: s.email, name: s.name, access_token: '', refresh_token: '', expires_at: 0 })
          .then(function () { throw err('signin', 'sign-in has run out (' + (tok.error || 'refused') + ')'); });
      }, function () { throw err('offline', 'no connection'); });
    });
    var done = function () { refreshing = null; };
    refreshing.then(done, done);
    return refreshing;
  }
  function who() { return kvGet('auth').then(function (s) { return s || {}; }); }

  /* ----------------------------------------------------------------- sending */
  function pathUrl(name) {
    var segs = (CONFIG.folder + '/' + name).split('/').map(encodeURIComponent).join('/');
    return GRAPH + '/me/drive/root:/' + segs + ':/createUploadSession';
  }
  function net(p) { return p.catch(function () { throw err('offline', 'no connection'); }); }
  function upload(item, tok) {
    var blob = item.blob;
    if (!blob || !blob.size) return Promise.reject(err('empty', 'the recording is empty'));
    return net(fetch(pathUrl(item.name), {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + tok, 'Content-Type': 'application/json' },
      body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'fail' } })
    })).then(function (r) {
      if (r.status === 409) return null;                       /* already there from an earlier try */
      if (r.status === 401) {
        return kvGet('auth').then(function (s) {
          if (s) { s.expires_at = 0; return kvSet('auth', s); }
        }).then(function () { throw err('retry', 'Microsoft wanted a fresh sign-in token'); });
      }
      if (!r.ok) return r.text().then(function (t) { throw err('http', 'OneDrive said ' + r.status + ': ' + t.slice(0, 160)); });
      return r.json().then(function (j) { return j.uploadUrl; });
    }).then(function (uploadUrl) {
      if (!uploadUrl) return;
      var size = blob.size, pos = 0;
      function next() {
        if (pos >= size) return Promise.resolve();
        var part = blob.slice(pos, Math.min(pos + CHUNK, size));
        return net(fetch(uploadUrl, {
          method: 'PUT',
          headers: { 'Content-Range': 'bytes ' + pos + '-' + (pos + part.size - 1) + '/' + size },
          body: part
        })).then(function (r) {
          if (r.status === 409) { pos = size; return; }        /* finished by an earlier try */
          if (r.status !== 200 && r.status !== 201 && r.status !== 202) {
            return r.text().then(function (t) { throw err('http', 'upload stopped at ' + pos + ' bytes: ' + r.status + ' ' + t.slice(0, 120)); });
          }
          pos += part.size;
          return next();
        });
      }
      return next();
    });
  }

  /* Send everything waiting, oldest first. Never runs twice at once (the page and
     the service worker can both ask). Resolves {sent, left, error, message}. */
  var sending = null;
  function sendAll() {
    if (sending) return sending;
    sending = items().then(function (all) {
      var todo = all.filter(function (x) { return x.state === 'waiting'; });
      if (!todo.length) return { sent: 0, left: 0, error: '' };
      return token().then(function (tok) {
        var sent = 0, stop = null, chain = Promise.resolve();
        todo.forEach(function (it) {
          chain = chain.then(function () {
            if (stop) return;
            return upload(it, tok).then(function () {
              sent++;
              return update(it.id, { state: 'sent', sentAt: Date.now(), blob: null, lastError: '' });
            }, function (e) {
              if (e.code === 'offline' || e.code === 'retry' || e.code === 'signin') stop = e;
              if (e.code === 'empty') return update(it.id, { state: 'sent', sentAt: Date.now(), blob: null, lastError: 'empty - nothing to send' });
              return update(it.id, { tries: (it.tries || 0) + 1, lastError: e.message });
            });
          });
        });
        return chain.then(function () {
          return { sent: sent, left: todo.length - sent, error: stop ? stop.code : (sent < todo.length ? 'http' : ''),
                   message: stop ? stop.message : '' };
        });
      }, function (e) {
        return { sent: 0, left: todo.length, error: e.code || 'error', message: e.message };
      });
    });
    var done = function () { sending = null; };
    sending.then(done, done);
    return sending;
  }

  return {
    kvGet: kvGet, kvSet: kvSet, kvDel: kvDel,
    putChunk: putChunk, chunksOf: chunksOf, dropChunks: dropChunks, chunkRecs: chunkRecs,
    enqueue: enqueue, items: items, update: update, prune: prune,
    memoItem: memoItem, newMemoId: newMemoId, sharedItem: sharedItem, extFor: extFor,
    saveAuth: saveAuth, token: token, who: who, form: form,
    sendAll: sendAll, upload: upload,
    AUTHORITY: AUTHORITY, SCOPES: SCOPES
  };
})();
