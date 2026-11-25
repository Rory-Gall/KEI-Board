/* KEI Voice - the page. One button: tap, talk, tap again, and it goes to Claude.
 *
 * Recording never needs a connection or a sign-in: each 2 s of audio is saved on
 * the phone as it is made, and a finished recording waits in a queue until it
 * can be sent. Sending (send.js) puts it in OneDrive; the PC transcribes it and
 * hands it to the Master session (KEI Recordings/_tools/phone_inbox.py).
 *
 * Sign-in is Microsoft's, straight against its endpoints (code + PKCE), like the
 * KEI board. Microsoft lets a web app keep a sign-in for only 24 hours, so when
 * that runs out the app re-connects silently with one quick redirect - but never
 * while a recording is running.
 */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  var btn = $('rec'), label = $('recLabel'), timer = $('timer'), meter = $('meter'), level = $('level');
  var statusEl = $('status'), alertEl = $('alert'), signinBtn = $('signin'), recentEl = $('recent');

  var REC = null;            /* the recording in progress */
  var busy = false, lastTap = 0, navigating = false, sendingNow = false;
  var wake = null, lastResult = null, note = '', authNote = '';
  var params = new URLSearchParams(location.search);
  var sharedN = params.get('shared');

  function show(msg) { alertEl.textContent = msg; alertEl.hidden = !msg; }
  function buzz(p) { try { if (navigator.vibrate) navigator.vibrate(p); } catch (e) { /* no vibration */ } }
  function two(n) { return (n < 10 ? '0' : '') + n; }
  function clock(ms) { var s = Math.floor(ms / 1000); return Math.floor(s / 60) + ':' + two(s % 60); }
  function dur(ms) { var s = Math.round(ms / 1000); return s < 60 ? s + ' s' : Math.floor(s / 60) + ' min ' + two(s % 60) + ' s'; }

  /* ---------------------------------------------------------------- sign-in */
  function redirectUri() { return location.origin + location.pathname.replace(/[^/]*$/, ''); }
  function b64url(bytes) {
    var s = '';
    new Uint8Array(bytes).forEach(function (b) { s += String.fromCharCode(b); });
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function randomVerifier() { var a = new Uint8Array(32); crypto.getRandomValues(a); return b64url(a); }
  function challenge(v) { return crypto.subtle.digest('SHA-256', new TextEncoder().encode(v)).then(b64url); }

  function signIn(silent) {
    if (REC || navigating) return Promise.resolve();
    return VOICE.who().then(function (w) {
      var verifier = randomVerifier();
      try { localStorage.setItem('kei_voice_pkce', verifier); } catch (e) { /* private mode */ }
      return challenge(verifier).then(function (ch) {
        if (REC) return;                         /* a recording started meanwhile: stay put */
        var q = {
          client_id: CONFIG.clientId, response_type: 'code', redirect_uri: redirectUri(),
          response_mode: 'query', scope: VOICE.SCOPES, code_challenge: ch, code_challenge_method: 'S256',
          prompt: silent ? 'none' : 'select_account'
        };
        if (w.email) q.login_hint = w.email;
        navigating = true;
        label.textContent = silent ? 'Reconnecting…' : 'Opening Microsoft sign-in…';
        location.href = VOICE.AUTHORITY + '/authorize?' + Object.keys(q).map(function (k) {
          return k + '=' + encodeURIComponent(q[k]);
        }).join('&');
      });
    });
  }

  /* Coming back from Microsoft: trade the one-time code for tokens. */
  function finishRedirect() {
    var code = params.get('code'), e = params.get('error');
    if (!code && !e) return Promise.resolve(null);
    history.replaceState({}, '', redirectUri());
    if (e) {
      try { sessionStorage.setItem('kei_voice_silent', 'failed'); } catch (x) { /* ignore */ }
      return Promise.resolve({ error: params.get('error_description') || e,
                               interactive: /interaction_required|login_required|consent_required/.test(e) });
    }
    var verifier = '';
    try { verifier = localStorage.getItem('kei_voice_pkce') || ''; localStorage.removeItem('kei_voice_pkce'); } catch (x) { /* ignore */ }
    return VOICE.form(VOICE.AUTHORITY + '/token', {
      client_id: CONFIG.clientId, grant_type: 'authorization_code', code: code,
      redirect_uri: redirectUri(), code_verifier: verifier, scope: VOICE.SCOPES
    }).then(function (tok) {
      if (!tok.access_token) return { error: tok.error_description || tok.error || 'sign-in failed' };
      return VOICE.who().then(function (prev) { return VOICE.saveAuth(tok, prev); }).then(function () { return { ok: true }; });
    }, function () { return { error: 'no connection while finishing sign-in' }; });
  }

  /* Microsoft ended the sign-in (24 h limit). Re-connect without asking, once per
     app session, only when nothing is being recorded and the person is not tapping. */
  var silentTimer = 0;
  function silentSoon(delay) {
    var tried = '';
    try { tried = sessionStorage.getItem('kei_voice_silent') || ''; } catch (e) { /* ignore */ }
    if (tried || !navigator.onLine) return;
    clearTimeout(silentTimer);
    silentTimer = setTimeout(function () {
      if (REC || busy || Date.now() - lastTap < 3000) return silentSoon(3000);
      VOICE.who().then(function (w) {
        if (!w.email) return;                    /* never signed in: that needs the button */
        try { sessionStorage.setItem('kei_voice_silent', 'tried'); } catch (e) { /* ignore */ }
        signIn(true);
      });
    }, delay);
  }

  /* -------------------------------------------------------------- recording */
  function pickMime() {
    var c = ['audio/webm;codecs=opus', 'audio/webm', 'audio/mp4;codecs=mp4a.40.2', 'audio/mp4'];
    for (var i = 0; i < c.length; i++) {
      if (window.MediaRecorder && MediaRecorder.isTypeSupported && MediaRecorder.isTypeSupported(c[i])) return c[i];
    }
    return '';
  }
  function lockScreen() {
    if (!('wakeLock' in navigator) || wake) return;
    navigator.wakeLock.request('screen').then(function (w) { wake = w; w.addEventListener('release', function () { wake = null; }); },
                                              function () { /* not allowed here: recording still works */ });
  }
  function unlockScreen() { if (wake) { wake.release().catch(function () {}); wake = null; } }

  function meterOn(r) {
    r.tick = setInterval(function () { timer.textContent = clock(Date.now() - r.started); }, 250);
    try {
      var Ctx = window.AudioContext || window.webkitAudioContext;
      if (!Ctx) return;
      r.ctx = new Ctx();
      var an = r.ctx.createAnalyser();
      an.fftSize = 512;
      r.ctx.createMediaStreamSource(r.stream).connect(an);
      var data = new Uint8Array(an.fftSize);
      (function frame() {
        an.getByteTimeDomainData(data);
        var peak = 0;
        for (var i = 0; i < data.length; i++) { var v = Math.abs(data[i] - 128); if (v > peak) peak = v; }
        level.style.width = Math.min(100, Math.round(peak / 128 * 160)) + '%';
        r.raf = requestAnimationFrame(frame);
      })();
    } catch (e) { /* the level bar is a nicety */ }
  }
  function meterOff(r) {
    clearInterval(r.tick);
    if (r.raf) cancelAnimationFrame(r.raf);
    if (r.ctx) r.ctx.close().catch(function () {});
    level.style.width = '0';
  }

  function start() {
    if (!navigator.mediaDevices || !window.MediaRecorder) {
      show('This browser cannot record audio. Open the app from its home-screen icon in Chrome.');
      return Promise.resolve();
    }
    show('');
    return navigator.mediaDevices.getUserMedia({
      audio: { channelCount: 1, echoCancellation: true, noiseSuppression: true, autoGainControl: true }
    }).then(function (stream) {
      var mime = pickMime(), opts = { audioBitsPerSecond: 32000 };
      if (mime) opts.mimeType = mime;
      var mr = new MediaRecorder(stream, opts);
      var started = Date.now();
      var r = { mr: mr, stream: stream, id: VOICE.newMemoId(started), seq: 0, started: started,
                mime: mr.mimeType || mime || 'audio/webm', writes: [] };
      REC = r;
      mr.ondataavailable = function (e) {
        if (!e.data || !e.data.size) return;
        r.writes.push(VOICE.putChunk(r.id, r.seq++, e.data)
          .then(function () { return VOICE.kvSet('active', { id: r.id, mime: r.mime, started: r.started, beat: Date.now() }); })
          .catch(function (x) { r.writeError = x; }));
      };
      mr.onstop = function () { finish(r); };
      mr.onerror = function (e) {
        show('The recorder hit a problem (' + ((e.error && e.error.message) || 'unknown') + '). What was recorded is kept.');
        try { mr.stop(); } catch (x) { finish(r); }
      };
      return VOICE.kvSet('active', { id: r.id, mime: r.mime, started: started, beat: Date.now() }).then(function () {
        mr.start(2000);           /* a piece every 2 s: a killed app loses at most 2 s */
        meterOn(r);
        lockScreen();
        buzz(40);
        paint();
      });
    }, function (e) {
      var denied = e && (e.name === 'NotAllowedError' || e.name === 'SecurityError');
      show(denied ? 'The microphone is blocked. Allow it for this app in Chrome\'s site settings, then tap again.'
                  : 'Could not start the microphone: ' + ((e && e.message) || e));
    });
  }

  function stop() {
    if (!REC) return Promise.resolve();
    var r = REC;
    r.stopping = true;
    paint();
    buzz([30, 60, 30]);
    try { r.mr.stop(); } catch (x) { finish(r); }
    return Promise.resolve();
  }

  /* The final dataavailable fires before stop, so every chunk write is queued by now. */
  function finish(r) {
    if (r.finished) return;
    r.finished = true;
    r.stream.getTracks().forEach(function (t) { t.stop(); });
    meterOff(r);
    unlockScreen();
    var ms = Date.now() - r.started;
    Promise.all(r.writes).then(function () {
      if (r.writeError) throw r.writeError;
      return VOICE.chunksOf(r.id);
    }).then(function (blobs) {
      var blob = new Blob(blobs, { type: r.mime.split(';')[0] });
      return VOICE.enqueue(VOICE.memoItem(r.id, blob, r.mime, r.started, ms)).then(function () { return blob; });
    }).then(function (blob) {
      return VOICE.dropChunks(r.id).then(function () { return VOICE.kvDel('active'); }).then(function () {
        REC = null;
        note = blob.size ? 'Saved (' + dur(ms) + ').' : 'That recording came out empty - nothing to send.';
        kick(true);
      });
    }).catch(function (e) {
      REC = null;
      show('Could not finish saving that recording (' + ((e && e.message) || e) + '). Its pieces are kept and ' +
           'will be put back together the next time the app opens.');
      paint();
    });
  }

  /* Pieces of a recording that never finished (the app was closed or killed). A
     recording still beating in another window of the app is left alone. */
  function recover() {
    return Promise.all([VOICE.kvGet('active'), VOICE.chunkRecs()]).then(function (x) {
      var active = x[0], recs = x[1], done = 0;
      var live = active && Date.now() - (active.beat || active.started || 0) < 30000;
      return recs.reduce(function (p, id) {
        return p.then(function () {
          if (live && active.id === id) return;
          return VOICE.chunksOf(id).then(function (blobs) {
            var mine = active && active.id === id;
            var mime = (mine && active.mime) || 'audio/webm';
            var started = (mine && active.started) || Date.now();
            var item = VOICE.memoItem(id, new Blob(blobs, { type: mime.split(';')[0] }), mime, started,
                                      mine ? Math.max(0, (active.beat || started) - started) : 0);
            item.recovered = true;
            done++;
            return VOICE.enqueue(item).then(function () { return VOICE.dropChunks(id); });
          });
        });
      }, Promise.resolve()).then(function () {
        if (done) note = 'Recovered ' + done + ' recording' + (done === 1 ? '' : 's') + ' that had been cut off.';
        if (active && !live) return VOICE.kvDel('active');
      });
    });
  }

  /* ---------------------------------------------------------------- sending */
  function syncLater() {
    if (!navigator.serviceWorker) return;
    navigator.serviceWorker.ready.then(function (reg) {
      if (reg.sync) return reg.sync.register('kei-voice-send');
    }).catch(function () { /* the next open or the 60 s timer will send */ });
  }
  function kick(justStopped) {
    sendingNow = true;
    paint();
    /* pick up a recording another window left cut off (it stops beating after 30 s) */
    return (REC ? Promise.resolve() : recover().catch(function () {})).then(function () {
      return VOICE.sendAll();
    }).then(function (r) {
      sendingNow = false;
      lastResult = r;
      if (r.left) syncLater();
      if (r.error === 'signin') silentSoon(justStopped ? 400 : 3000);
      return paint();
    }, function (e) {
      sendingNow = false;
      lastResult = { error: 'error', message: e.message, left: 1 };
      return paint();
    });
  }

  /* ------------------------------------------------------------------ paint */
  function paint() {
    btn.classList.toggle('on', !!REC);
    btn.setAttribute('aria-pressed', REC ? 'true' : 'false');
    if (!navigating) label.textContent = REC ? (REC.stopping ? 'Saving…' : 'Tap to stop') : 'Tap to talk';
    meter.hidden = !REC;
    if (!REC) timer.textContent = '';
    return Promise.all([VOICE.items(), VOICE.who()]).then(function (x) {
      var all = x[0], w = x[1];
      var waiting = all.filter(function (i) { return i.state === 'waiting'; });
      var signedIn = !!(w.refresh_token || w.access_token);
      var r = lastResult || {};
      var msg;
      if (sendingNow && waiting.length) msg = 'Sending to Claude…';
      else if (!waiting.length) msg = all.length ? 'Everything has gone to Claude.' : 'Nothing recorded yet.';
      else if (!signedIn || r.error === 'signin') msg = waiting.length + ' waiting. Sign in once and they go to Claude.';
      else if (r.error === 'offline' || !navigator.onLine) msg = waiting.length + ' waiting. They go to Claude when you have signal.';
      else if (r.error && r.message) msg = waiting.length + ' waiting. The last try failed: ' + r.message + '. Trying again.';
      else msg = waiting.length + ' waiting to go to Claude.';
      statusEl.textContent = (note ? note + ' ' : '') + msg + (authNote ? ' ' + authNote : '');
      signinBtn.hidden = !(!signedIn || r.error === 'signin') || navigating;
      signinBtn.textContent = w.email ? 'Sign in again (' + w.email + ')' : 'Sign in with your KEI Microsoft account';
      renderRecent(all);
    }).catch(function (e) { statusEl.textContent = 'Storage problem: ' + e.message; });
  }

  function renderRecent(all) {
    var list = all.slice(-6).reverse();
    recentEl.textContent = '';
    list.forEach(function (it) {
      var li = document.createElement('li');
      var t = new Date(it.created);
      var when = t.toLocaleDateString([], { month: 'short', day: 'numeric' }) + ', ' +
                 t.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
      var what = it.kind === 'shared' ? 'Shared audio' : 'Memo' + (it.ms ? ', ' + dur(it.ms) : '');
      if (it.recovered) what += ' (recovered)';
      var st, cls;
      if (it.state === 'sent') { st = it.lastError ? it.lastError : 'Sent'; cls = it.lastError ? 'warn' : 'ok'; }
      else if ((it.tries || 0) >= 3) { st = 'Not sent yet: ' + it.lastError; cls = 'bad'; }
      else { st = 'Waiting'; cls = 'wait'; }
      li.className = cls;
      var a = document.createElement('span'); a.className = 'what'; a.textContent = when + ' · ' + what;
      var b = document.createElement('span'); b.className = 'state'; b.textContent = st;
      li.appendChild(a); li.appendChild(b);
      recentEl.appendChild(li);
    });
    $('recentWrap').hidden = !list.length;
  }

  /* ------------------------------------------------------------------ wiring */
  btn.addEventListener('click', function () {
    var now = Date.now();
    if (now - lastTap < 700 || busy || navigating) return;
    lastTap = now;
    note = '';
    busy = true;
    (REC ? stop() : start()).then(function () { busy = false; }, function () { busy = false; });
  });
  signinBtn.addEventListener('click', function () { signIn(false); });
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState !== 'visible') return;
    if (REC) lockScreen();
    kick();
  });
  window.addEventListener('online', function () { kick(); });
  setInterval(function () {
    if (document.visibilityState !== 'visible' || sendingNow || navigating) return;
    VOICE.items().then(function (a) {
      if (a.some(function (i) { return i.state === 'waiting'; })) kick();
    });
  }, 60000);

  if (navigator.serviceWorker) {
    navigator.serviceWorker.register('sw.js').catch(function (e) {
      show('Offline mode is not set up (' + e.message + '). Recording works; open the app with signal once.');
    });
    navigator.serviceWorker.addEventListener('message', function (e) {
      if (e.data && e.data.type === 'send') kick(); else paint();
    });
  }

  $('build').textContent = 'build ' + BUILD;
  finishRedirect().then(function (res) {
    if (res && res.error) {
      authNote = res.interactive ? 'Microsoft wants you to sign in again - tap the button.' : 'Sign-in did not finish: ' + res.error;
    }
    if (sharedN !== null) {
      var n = +sharedN || 0;
      note = n ? 'Got ' + n + ' file' + (n === 1 ? '' : 's') + ' to send to Claude.' : '';
      if (!n) show('Nothing came through from that share. Try sharing it again.');
      if (!res) history.replaceState({}, '', redirectUri());
    }
    return recover();
  }).then(function () { return VOICE.prune(); }).then(function () { return paint(); })
    .then(function () { return kick(); })
    .catch(function (e) { show('Start-up problem: ' + ((e && e.message) || e)); paint(); });

  /* For the test harness (tests/harness.html). Harmless in normal use. */
  window.KEIVoice = { start: start, stop: stop, kick: kick, recover: recover, paint: paint,
                      recording: function () { return !!REC; } };
})();
