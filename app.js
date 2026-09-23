/* KEI Board — the team's shared task board.
 * Tasks live in two SharePoint lists (see store.js); this file is the board
 * everyone looks at. Nothing is stored in the page itself, so what a person
 * sees is whatever their own Microsoft account is allowed to read. */
'use strict';

/* ---------- state + constants ---------- */
/* Empty until the lists answer. boot() shows a "loading" line rather than a
   blank board, and connect() fills this in. */
var STATE = { v: 1, board: 'team', labels: [], lanes: ['Rory'], tasks: [], activity: [], personalLabels: null };
var TITLE = 'KEI Team Board';
var FONT_LINKS = '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:wght@400;600;700;800&family=Archivo:wght@400;500;600;700&family=Spline+Sans+Mono:wght@400;500;600&display=swap">';
var BUCKETS = ['tosort', 'now', 'progress', 'today', 'week', 'later', 'someday'];
var BUCKET_NAME = { tosort: 'To Sort', now: 'Now', progress: 'In Progress / Waiting On',
                    today: 'Today', week: 'This Week', later: 'Later', someday: 'Some Day / Maybe' };
/* short form for the picker on a task row, where space is tight */
var BUCKET_SHORT = { tosort: 'To Sort', now: 'Now', progress: 'In Prog / Waiting',
                     today: 'Today', week: 'Week', later: 'Later', someday: 'Someday' };
/* buckets that existed before, and where their tasks go now */
var BUCKET_WAS = { next: 'week', inbox: 'tosort' };

/* Work / Personal is a second axis, independent of the project filter: you can
   be in Personal AND filtered to House. A project is personal if it is named
   here - add one and it moves sides. */
var PERSONAL_LABELS = (STATE.personalLabels && STATE.personalLabels.length)
  ? STATE.personalLabels : ['Personal', 'House'];
function isPersonal(t) { return PERSONAL_LABELS.indexOf(t.label) >= 0; }
function inScope(t) {
  var s = PREFS.scope || 'all';
  if (s === 'personal') return isPersonal(t);
  if (s === 'work') return !isPersonal(t);
  return true;
}

var VIEW = { q: '', filter: 'all', open: null, activity: false, projects: false, projDel: null };
try { var _v = sessionStorage.getItem('kei_view'); if (_v) VIEW = Object.assign(VIEW, JSON.parse(_v)); } catch (e) {}
var WHO = null;
try { WHO = localStorage.getItem('kei_who') || null; } catch (e) {}

/* Per-person view preferences: which lanes are shown and how wide each is.
   Browser-local only - never part of the shared board, never published, so
   one person's layout can't touch anyone else's. */
var LANE_MIN = 160, LANE_MAX = 1200;
var PREFS = { hide: {}, w: {}, scope: 'all' };
var PREFS_KEY = 'kei_prefs_' + (STATE.board || 'team');
try {
  var _p = JSON.parse(localStorage.getItem(PREFS_KEY) || 'null');
  if (_p && typeof _p === 'object') { PREFS.hide = _p.hide || {}; PREFS.w = _p.w || {}; PREFS.scope = _p.scope || 'all'; }
} catch (e) {}
function savePrefs() { try { localStorage.setItem(PREFS_KEY, JSON.stringify(PREFS)); } catch (e) {} }
function laneHidden(l) { return !!PREFS.hide[l]; }
function laneWidth(l) { var w = +PREFS.w[l]; return (w >= LANE_MIN && w <= LANE_MAX) ? w : 0; }

var READONLY = false;
/* No store at all (e.g. the file opened straight off disk): draw, never save. */
var LOCALMODE = !(window.STORE && window.AUTH);
var SAVESTATE = 'idle'; /* idle | dirty | saving | saved | viewonly | offline */
var whoCallback = null;
var whoPrev = null;
var FAILMSG = '';
var DIRTYSEQ = 0;

/* ---------- utilities ---------- */
function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function nowISO() { return new Date().toISOString(); }
function shortDate(iso) {
  if (!iso) return '';
  var d = new Date(iso);
  if (isNaN(d)) return '';
  var mo = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()];
  return mo + ' ' + d.getDate();
}
function ago(iso) {
  var ms = Date.now() - new Date(iso).getTime();
  var m = Math.floor(ms / 60000);
  if (m < 1) return 'just now';
  if (m < 60) return m + 'm ago';
  var h = Math.floor(m / 60);
  if (h < 24) return h + 'h ago';
  var d = Math.floor(h / 24);
  if (d < 30) return d + 'd ago';
  return shortDate(iso);
}
function byId(id) { return STATE.tasks.find(function (t) { return t.id === id; }); }
function labelOf(name) { return STATE.labels.find(function (l) { return l.name === name; }); }
function openCount(lane) {
  return STATE.tasks.filter(function (t) {
    return t.lane === lane && !t.done && inScope(t);
  }).length;
}
function act(text) {
  var e = { id: String(Date.now()) + '_' + Math.floor(Math.random() * 100000),
            who: WHO || 'Someone', at: nowISO(), text: text };
  STATE.activity.unshift(e);
  if (STATE.activity.length > 60) STATE.activity.length = 60;
  NEWACT.push(e);   /* picked up by the next pushChanges */
}

/* ---------- shared store ----------
   Each task is one row in a SharePoint list (store.js). A change writes only
   the rows that actually changed, and every open board picks up other people's
   changes on a short poll - so two people working at the same time no longer
   collide unless they are typing into the very same task, and even then only
   that one task is at stake, never the board.

   Which list a task is written to is decided by its tag: personal tags and the
   Contractors lane go to the private list that only Rory and Anne can open.

   STATE stays the single in-memory model the whole UI renders from. It is fed
   from the lists on the way in, and diffed against them on the way out. */

var CONNECTED = false;         /* the lists answered at least once */
var SERVER = {};               /* last body the store gave us, by task id */
var NEWACT = [];               /* activity entries this mutation created */
var RETRY = [];                /* writes to try once more */
var retryTimer = null;
var renderWanted = false;
var STORE_READY = false;
var POLL_MS = 15000;
var pollTimer = null;
var SIGNEDIN = null;           /* {name, email} once signed in */

/* Ids must be unique without a shared counter: the store is last-writer-wins
   and has no transactions, so ++seq across five people is a lost-update bug
   waiting to happen. Milliseconds plus a random tail stays inside the safe
   integer range and keeps ids numeric, which the rest of the app assumes. */
function newId() { return Date.now() * 1000 + Math.floor(Math.random() * 1000); }

/* ---------- shapes stored ---------- */
var TASK_FIELDS = ['id', 't', 'lane', 'bucket', 'label', 'notes', 'done',
                   'doneAt', 'created', 'by', 'src', 'comments'];
function taskBody(t) {
  var b = {};
  TASK_FIELDS.forEach(function (k) { b[k] = (k === 'comments') ? (t.comments || []) : t[k]; });
  return b;
}
/* The board-wide settings, as one object. Written to the Board Settings list,
   one row per key, whenever any of them changes. */
function metaBody() {
  return { labels: STATE.labels, lanes: STATE.lanes,
           personalLabels: STATE.personalLabels || null, schemaVersion: STATE.v || 1 };
}
/* A comment carries an id so two people adding one to the same task at the
   same moment merge instead of overwriting. Older comments predate the id. */
function cmtKey(c) { return c.id || (String(c.who) + '|' + String(c.at) + '|' + String(c.text)); }
function mergeComments(mine, theirs) {
  if (!theirs || !theirs.length) return mine || [];
  var out = [], seen = {};
  (mine || []).concat(theirs).forEach(function (c) {
    var k = cmtKey(c);
    if (seen[k]) return;
    seen[k] = 1; out.push(c);
  });
  out.sort(function (a, b) { return String(a.at) < String(b.at) ? -1 : 1; });
  return out;
}

/* ---------- reading the store ---------- */
/* The fixups boot used to do once now run on everything arriving from the
   store, because another view may be one version behind on either. */
function normalizeTask(b, docId) {
  var t = {};
  TASK_FIELDS.forEach(function (k) { t[k] = b[k]; });
  var n = Number(docId);
  t.id = isFinite(n) ? n : Number(b.id);
  t.t = String(t.t == null ? '' : t.t);
  t.comments = t.comments || [];
  t.done = !!t.done;
  if (BUCKET_WAS[t.bucket]) t.bucket = BUCKET_WAS[t.bucket];
  if (BUCKETS.indexOf(t.bucket) < 0) t.bucket = 'tosort';
  if (!t.lane || STATE.lanes.indexOf(t.lane) < 0) t.lane = STATE.lanes[0];
  return t;
}
/* The old board kept new tasks at the front of the array and rendered in array
   order. Sorting newest-first reproduces exactly that, from any delivery. */
function byNewest(a, b) {
  var x = String(a.created || ''), y = String(b.created || '');
  if (x !== y) return x < y ? 1 : -1;
  return (b.id || 0) - (a.id || 0);
}
/* A whole load: everything the signed-in person can read. */
function applyTasks(tasks) {
  var list = [];
  SERVER = {};
  tasks.forEach(function (t) {
    SERVER[t.id] = taskBody(t);
    list.push(normalizeTask(t, t.id));
  });
  list.sort(byNewest);
  STATE.tasks = list;
  wantRender();
}
/* A poll: only what changed since last time, so merge rather than replace.
   Anything the person is editing right now is left alone - see safeToRender. */
function applyChanges(changes) {
  if (!changes.length) return;
  changes.forEach(function (c) {
    if (c.deletedId != null) {
      delete SERVER[c.deletedId];
      STATE.tasks = STATE.tasks.filter(function (t) { return t.id !== c.deletedId; });
      return;
    }
    var t = normalizeTask(c.task, c.task.id);
    SERVER[t.id] = taskBody(t);
    var at = STATE.tasks.findIndex(function (x) { return x.id === t.id; });
    if (at < 0) STATE.tasks.push(t);
    else STATE.tasks[at] = t;
  });
  STATE.tasks.sort(byNewest);
  wantRender();
}
function applyMeta(s) {
  if (!s) return;
  if (s.labels && s.labels.length) STATE.labels = s.labels;
  if (s.lanes && s.lanes.length) STATE.lanes = s.lanes;
  STATE.personalLabels = s.personalLabels || null;
  PERSONAL_LABELS = (STATE.personalLabels && STATE.personalLabels.length)
    ? STATE.personalLabels : ['Personal', 'House'];
  STORE.setPersonalLabels(PERSONAL_LABELS);
  wantRender();
}
function applyActivity(list) {
  list = (list || []).slice();
  list.sort(function (a, b) { return String(a.at) < String(b.at) ? 1 : -1; });
  STATE.activity = list;
  wantRender();
}

/* Re-render as soon as it cannot cost anyone their typing. Someone else's edit
   arriving must never move the caret or empty a half-written task, so while a
   field has focus (or a drag or an inline edit is in flight) the data is taken
   in and the redraw waits for the next quiet moment. */
function safeToRender() {
  if (document.body.classList.contains('dragging')) return false;
  if (document.querySelector('.tinline')) return false;
  if (VIEW.projects) return false;
  var ae = document.activeElement;
  if (ae && /^(INPUT|TEXTAREA|SELECT)$/.test(ae.tagName)) return false;
  return true;
}
function wantRender() {
  if (safeToRender()) { renderWanted = false; render(); }
  else renderWanted = true;
}
setInterval(function () { if (renderWanted) wantRender(); }, 1200);

/* ---------- writing to the store ---------- */
function goReadOnly() {
  READONLY = true;
  SAVESTATE = 'viewonly';
  render();
}
function saveFailedHard(msg) {
  RETRY = [];
  if (retryTimer) { clearTimeout(retryTimer); retryTimer = null; }
  SAVESTATE = 'failed';
  FAILMSG = msg || '';
  renderChip();
}
/* One row, one write. `job` is {task}, {del:id}, {meta:true} or {act:entry}. */
function runWrite(job) {
  if (job.del != null) return STORE.remove({ id: job.del });
  if (job.task) return STORE.put(job.task);
  if (job.meta) {
    var m = metaBody();
    return Promise.all(Object.keys(m).map(function (k) { return STORE.putSetting(k, m[k]); }));
  }
  if (job.act) return STORE.addActivity(job.act);
  return Promise.resolve();
}
function flush(jobs, isRetry) {
  if (!jobs.length) return;
  SAVESTATE = 'saving';
  renderChip();
  var failed = [], fatal = null;
  Promise.all(jobs.map(function (job) {
    return runWrite(job).catch(function (err) {
      var msg = (err && err.message) || '';
      if (msg === AUTH.EXPIRED) { fatal = 'signedout'; return; }
      if (err && err.status === 403) { fatal = 'readonly'; return; }
      if (err && err.status === 400) { fatal = fatal || 'bad'; return; }
      failed.push(job);   /* offline, throttled, server hiccup - worth retrying */
    });
  })).then(function () {
    if (fatal === 'signedout') { signedOut(); return; }
    if (fatal === 'readonly') { goReadOnly(); return; }
    if (fatal === 'bad') { saveFailedHard(''); return; }
    if (!failed.length) {
      SAVESTATE = 'saved';
      renderChip();
      return;
    }
    /* The store asked us to slow down or was briefly unreachable. Try the
       failed documents once more; give up honestly rather than spin. */
    if (isRetry) { saveFailedHard('some changes did not save - reload the page'); return; }
    RETRY = RETRY.concat(failed);
    SAVESTATE = 'offline';
    renderChip();
    if (retryTimer) clearTimeout(retryTimer);
    retryTimer = setTimeout(function () {
      retryTimer = null;
      var again = RETRY; RETRY = [];
      flush(again, true);
    }, 1200 + Math.random() * 1800);
  });
}

/* Work out what this mutation actually changed and write only that. Comparing
   the before and after of STATE is what lets every existing call site stay a
   plain `mutate(function () { ... })` with no bookkeeping of its own. */
function snapTasks() {
  var m = {};
  STATE.tasks.forEach(function (t) { m[t.id] = JSON.stringify(taskBody(t)); });
  return m;
}
function pushChanges(before, metaBefore) {
  var acts = NEWACT.splice(0);
  if (LOCALMODE || !CONNECTED) return;
  var jobs = [], after = snapTasks();
  Object.keys(after).forEach(function (id) {
    if (before[id] === after[id]) return;
    var t = byId(Number(id));
    if (!t) return;
    var srv = SERVER[id];
    if (srv) t.comments = mergeComments(t.comments, srv.comments);
    jobs.push({ task: t });
  });
  Object.keys(before).forEach(function (id) {
    if (!(id in after)) jobs.push({ del: Number(id) });
  });
  if (JSON.stringify(metaBody()) !== metaBefore) jobs.push({ meta: true });
  acts.forEach(function (e) { jobs.push({ act: e }); });
  flush(jobs, false);
}
function mutate(fn) {
  if (READONLY) return;
  var before = snapTasks();
  var metaBefore = JSON.stringify(metaBody());
  fn();
  DIRTYSEQ++;
  if (!LOCALMODE && CONNECTED) { SAVESTATE = 'dirty'; }
  render();
  pushChanges(before, metaBefore);
}

/* ---------- connecting ---------- */
/* Sign in, read both lists once, then keep up with other people by asking
   SharePoint what changed every POLL_MS. There is nothing to seed and no lock
   to take: the lists already exist, and each row is written on its own. */
function signedOut() {
  CONNECTED = false;
  stopPolling();
  SAVESTATE = 'signedout';
  FAILMSG = 'signed out - sign in again to save';
  render();
}
function stopPolling() {
  if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
}
function startPolling() {
  stopPolling();
  pollTimer = setInterval(function () {
    /* Nothing to do while the tab is hidden; the next look catches up anyway. */
    if (document.hidden) return;
    STORE.poll().then(applyChanges).catch(function (err) {
      if ((err && err.message) === AUTH.EXPIRED) { signedOut(); return; }
      /* A poll that fails is not a save that failed - stay quiet and try again. */
    });
  }, POLL_MS);
}
function loadAll() {
  return STORE.load().then(function (data) {
    applyMeta(data.settings);
    applyTasks(data.tasks);
    applyActivity(data.activity);
    CONNECTED = true;
    STORE_READY = true;
    SAVESTATE = 'idle';
    if (data.badComments && data.badComments.length) {
      FAILMSG = 'some comments could not be read (' + data.badComments.length + ') - tell Claude';
    }
    render();
    startPolling();
  });
}
function connect() {
  if (LOCALMODE) { render(); return; }
  AUTH.start().then(function (r) {
    SIGNEDIN = r.signedIn ? { name: r.name, email: r.email } : null;
    if (!r.signedIn) { SAVESTATE = 'signedout'; render(); return; }
    if (!WHO) {
      /* The signed-in name is the person, so nobody has to pick from a list. */
      WHO = laneForEmail(r.email) || r.name;
      try { localStorage.setItem('kei_who', WHO); } catch (e) {}
    }
    return loadAll();
  }).catch(function (err) {
    if ((err && err.message) === AUTH.EXPIRED) { signedOut(); return; }
    SAVESTATE = 'offline';
    FAILMSG = 'could not reach the board - reload the page';
    render();
  });
}
/* Refresh when the tab comes back, so a board left open overnight is current
   the moment it is looked at rather than up to POLL_MS later. */
document.addEventListener('visibilitychange', function () {
  if (!document.hidden && CONNECTED) {
    STORE.poll().then(applyChanges).catch(function () {});
  }
});

/* Read whatever is currently typed in the open task and commit it. Every close
   path calls this, so an edit can never be lost to a missing blur/change. */
function commitOpenEdits() {
  if (READONLY || VIEW.open == null) return;
  var t = byId(VIEW.open);
  if (!t) return;
  var before = snapTasks(), metaBefore = JSON.stringify(metaBody());
  var changed = false;
  var ti = document.querySelector('[data-edit-title]');
  if (ti) { var v = ti.value.trim(); if (v && v !== t.t) { t.t = v; act('edited a task title'); changed = true; } }
  var no = document.querySelector('[data-edit-notes]');
  if (no) { var nv = no.value.trim(); if (nv !== (t.notes || '')) { t.notes = nv; changed = true; } }
  if (changed) { DIRTYSEQ++; pushChanges(before, metaBefore); }
}

/* Commit a text edit WITHOUT rebuilding the DOM: re-rendering on blur swallowed
   the click that caused the blur, so Close / Mark done / Post needed two taps. */
function mutateQuiet(fn) {
  if (READONLY) return;
  var before = snapTasks(), metaBefore = JSON.stringify(metaBody());
  fn();
  DIRTYSEQ++;
  pushChanges(before, metaBefore);
}

/* ---------- who-are-you ---------- */
/* Signing in already says who you are, so there is nothing to ask. */
function needWho(cb) { cb(); }
function setWho(name) {
  WHO = name;
  whoPrev = null;
  try { localStorage.setItem('kei_who', name); } catch (e) {}
  var cb = whoCallback; whoCallback = null;
  render();
  if (cb) cb();
}
function cancelWho() {
  whoCallback = null;
  if (!WHO && whoPrev) {
    WHO = whoPrev;
    try { localStorage.setItem('kei_who', WHO); } catch (e) {}
  }
  whoPrev = null;
  render();
}

/* ---------- CSS ---------- */
var CSS = [
'*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}',
':root{--ground:#EEF2EF;--surface:#FFFFFF;--surface2:#E3EAE5;--ink:#1B2A24;--muted:#5B6E64;--line:#D5DED8;',
' --accent:#0F766E;--accent-ink:#FFFFFF;--accent-soft:#DCEAE7;--now:#A34114;--now-ink:#FFFFFF;--now-soft:#F7E4D8;',
' --danger:#B3372F;--chipbg:#E7EEE9;--shadow:0 1px 2px rgba(27,42,36,.08),0 2px 8px rgba(27,42,36,.05)}',
'@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--ground:#0F1613;--surface:#18211C;--surface2:#1F2A24;',
' --ink:#E4ECE7;--muted:#93A69B;--line:#2A3830;--accent:#34A79B;--accent-ink:#07211E;--accent-soft:#14332F;',
' --now:#E67E3F;--now-ink:#1A1207;--now-soft:#3A2417;--danger:#E06258;--chipbg:#232F29;--shadow:0 1px 2px rgba(0,0,0,.4)}}',
':root[data-theme="dark"]{--ground:#0F1613;--surface:#18211C;--surface2:#1F2A24;',
' --ink:#E4ECE7;--muted:#93A69B;--line:#2A3830;--accent:#34A79B;--accent-ink:#07211E;--accent-soft:#14332F;',
' --now:#E67E3F;--now-ink:#1A1207;--now-soft:#3A2417;--danger:#E06258;--chipbg:#232F29;--shadow:0 1px 2px rgba(0,0,0,.4)}',
'html{height:100%}',
'body{background:var(--ground);color:var(--ink);font:13.5px/1.35 Archivo,system-ui,sans-serif;min-height:100%;',
' -webkit-font-smoothing:antialiased}',
'button{font:inherit;color:inherit;background:none;border:none;cursor:pointer}',
'button:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible{outline:2px solid var(--accent);outline-offset:2px}',
'input,select,textarea{font:inherit;color:var(--ink)}',
'@media (prefers-reduced-motion:reduce){*{transition:none!important;animation:none!important}}',
/* header */
'.top{position:sticky;top:0;z-index:20;background:var(--ground);border-bottom:1px solid var(--line);padding:5px 12px;',
' display:flex;align-items:center;gap:8px;flex-wrap:wrap}',
'.brand{font-family:"Bricolage Grotesque",Archivo,sans-serif;font-weight:800;font-size:15px;letter-spacing:-.02em;white-space:nowrap}',
'.brand .dot{color:var(--accent)}',
'.datechip{font-family:"Spline Sans Mono",monospace;font-size:10.5px;color:var(--muted);white-space:nowrap}',
'.stat{font-family:"Spline Sans Mono",monospace;font-size:10.5px;padding:2px 7px;border-radius:99px;background:var(--chipbg);color:var(--muted);white-space:nowrap}',
'.stat.hot{background:var(--now-soft);color:var(--now);font-weight:600}',
'.savechip{font-size:11.5px;font-family:"Spline Sans Mono",monospace;color:var(--muted);margin-left:auto;white-space:nowrap}',
'.stat.exp{background:var(--accent-soft);color:var(--accent);font-weight:600}',
'.stat.exp:hover{background:var(--accent);color:var(--accent-ink)}',
'.stat.exp[disabled]{opacity:.6;cursor:default}',
'.savechip.err{color:var(--danger)}',
'.whobtn{font-size:11px;padding:2px 8px;border:1px solid var(--line);border-radius:99px;background:var(--surface);color:var(--muted)}',
'.whobtn:hover{border-color:var(--accent);color:var(--accent)}',
/* banner */
'.banner{margin:10px 16px 0;padding:9px 14px;border-radius:10px;font-size:13px;background:var(--accent-soft);color:var(--ink);border:1px solid var(--line)}',
'.banner.warn{background:var(--now-soft)}',
/* quick add */
'.qa{display:flex;gap:6px;align-items:center;padding:6px 12px 2px;flex-wrap:wrap}',
'.qa input{flex:1;min-width:180px;padding:6px 10px;border:1px solid var(--line);border-radius:8px;background:var(--surface);font-size:13px}',
'.qa input::placeholder{color:var(--muted)}',
'.qa .to{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.07em}',
'.pbtn{padding:8px 13px;border-radius:9px;background:var(--surface);border:1px solid var(--line);font-weight:600;font-size:13px}',
'.pbtn:hover{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}',
/* filters */
'.filters{display:flex;gap:4px;align-items:center;padding:3px 12px 2px;flex-wrap:wrap}',
'.fchip{padding:3px 9px;border-radius:99px;font-size:11.5px;background:var(--chipbg);color:var(--muted);border:1px solid transparent}',
'.fchip:hover{color:var(--ink)}',
'.fchip.on{background:var(--accent);color:var(--accent-ink);font-weight:600}',
'.fchip.review{color:var(--now)}',
'.fchip.review.on{background:var(--now);color:var(--now-ink)}',
'.fcap{font-size:9.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-right:2px}',
'.seg{display:inline-flex;border:1px solid var(--line);border-radius:99px;overflow:hidden;background:var(--surface)}',
'.segb{padding:3px 11px;font-size:11.5px;font-weight:600;color:var(--muted);border:0;background:none;cursor:pointer}',
'.segb + .segb{border-left:1px solid var(--line)}',
'.segb:hover{color:var(--ink)}',
'.segb.on{background:var(--accent);color:var(--accent-ink)}',
'.lchip{padding:3px 9px;border-radius:99px;font-size:11.5px;border:1px solid var(--line);background:var(--surface);color:var(--ink);font-weight:600}',
'.lchip.off{background:none;color:var(--muted);font-weight:400;text-decoration:line-through;border-style:dashed}',
'.lchip:hover{border-color:var(--accent)}',
'.fsep{width:1px;height:16px;background:var(--line);margin:0 4px}',
'.fdot{display:inline-block;width:6px;height:6px;border-radius:50%;margin-right:4px;vertical-align:0}',
'.search{margin-left:auto;padding:4px 10px;border:1px solid var(--line);border-radius:99px;background:var(--surface);font-size:12px;width:160px}',
/* inbox */
'.inbox{margin:6px 12px 0;border:1.5px dashed var(--accent);border-radius:10px;padding:6px 10px;background:var(--surface)}',
'.inbox h2{font-family:"Bricolage Grotesque",sans-serif;font-size:11.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--accent);margin-bottom:1px}',
'.inbox .sub{font-size:11px;color:var(--muted);margin-bottom:4px}',
'.stat.sort{background:var(--accent-soft);color:var(--accent);font-weight:600}',
'.pbtn.primary{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}',
'.pbtn.primary:hover{filter:brightness(1.08)}',
'.fchip.edit{color:var(--muted);border:1px dashed var(--line)}',
'.fchip.edit:hover{color:var(--accent);border-color:var(--accent)}',
'.modal.proj{max-width:420px}',
'.msub{font-size:11.5px;color:var(--muted);margin-bottom:8px}',
'.projrow{display:flex;gap:6px;align-items:center;padding:3px 0;border-top:1px solid var(--line)}',
'.projrow.add{border-top:1px solid var(--line);margin-top:4px;padding-top:6px}',
'.hueswatch{flex:0 0 auto;width:14px;height:14px;border-radius:50%;border:1px solid var(--line);cursor:pointer}',
'.projname{flex:1 1 auto;min-width:0;font:inherit;font-size:12.5px;padding:3px 6px;border:1px solid var(--line);border-radius:6px;background:var(--surface);color:var(--ink)}',
'.projname:focus{outline:none;border-color:var(--accent)}',
'.projused{flex:0 0 auto;font-size:10.5px;color:var(--muted);font-family:"Spline Sans Mono",monospace}',
'.abtn.danger{background:var(--now-soft);color:var(--now);border-color:var(--now);font-weight:600}',
'.inrow{display:flex;gap:6px;align-items:center;padding:4px 0;border-top:1px solid var(--line);flex-wrap:wrap}',
'.inrow .txt{flex:1;min-width:180px;font-weight:600;font-size:13px}',
'.inrow .src{font-size:11px;color:var(--muted)}',
'.inrow .abtn{padding:3px 8px;border-radius:6px;border:1px solid var(--line);font-size:11.5px;background:var(--surface2)}',
'.inrow .abtn:hover{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}',
/* board */
'.board{display:flex;gap:6px;padding:6px 10px 30px;overflow-x:auto;align-items:flex-start}',
'.lane{flex:0 0 340px;min-width:0;position:relative;background:var(--surface2);border-radius:6px;padding:4px 4px 5px;min-height:40px}',
'.lane h2{font-family:"Bricolage Grotesque",sans-serif;font-size:13px;font-weight:700;padding:1px 4px 4px;display:flex;align-items:baseline;gap:6px}',
'.lane h2 .n{font-family:"Spline Sans Mono",monospace;font-size:10px;color:var(--muted);font-weight:400}',
'.lgrip{position:absolute;top:0;right:-4px;width:8px;height:100%;cursor:col-resize;border-radius:4px;z-index:2;touch-action:none}',
'.lgrip:hover,.lgrip.on{background:var(--accent);opacity:.45}',
'body.resizing{cursor:col-resize;user-select:none}',
'.bsec{margin-bottom:5px}',
'.bsec.empty{display:none}',
'body.dragging .bsec.empty{display:block;border:1px dashed var(--line);border-radius:4px;min-height:22px}',
'.bsec.over{outline:2px dashed var(--accent);outline-offset:1px;background:var(--accent-soft);border-radius:4px}',
'.grab{flex:0 0 auto;font-size:11px;line-height:1;color:var(--muted);opacity:.45;cursor:grab;touch-action:none;-webkit-user-select:none;user-select:none;margin-right:-2px}',
'.card:hover .grab{opacity:1}',
'.card.lifted{opacity:.35}',
'.card.ghost{position:fixed;z-index:60;pointer-events:none;opacity:.94;box-shadow:0 6px 18px rgba(0,0,0,.28);margin:0;cursor:grabbing}',
'body.dragging,body.dragging .card,body.dragging .card .t{cursor:grabbing!important;-webkit-user-select:none;user-select:none}',
'.bhead{font-size:9.5px;text-transform:uppercase;letter-spacing:.1em;color:var(--muted);padding:1px 4px 2px;font-weight:600}',
'.bhead.now{color:var(--now)}',
'.bhead.tosort{color:var(--accent)}',
'.card{background:var(--surface);border-radius:4px;padding:2px 30px 2px 6px;margin-bottom:2px;box-shadow:var(--shadow);display:flex;flex-wrap:nowrap;align-items:center;column-gap:6px;',
' border-left:2px solid transparent;cursor:pointer;position:relative}',
'.card:hover{border-left-color:var(--accent)}',
'.card.now{border-left-color:var(--now)}',
'.card .t{flex:1 1 auto;min-width:0;font-size:12.5px;font-weight:500;line-height:1.2;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:text;-webkit-user-select:none;user-select:none}',
'.card .t:hover{text-decoration:underline dotted var(--muted)}',
'.tinline{flex:1 1 auto;min-width:0;font:inherit;font-size:12.5px;font-weight:500;line-height:1.2;padding:0 3px;border:1px solid var(--accent);border-radius:3px;background:var(--surface);color:var(--ink);outline:none}',
'.lblsel,.bktsel{font-size:9.5px;font-family:inherit;padding:0 2px;border-radius:99px;border:1px solid var(--line);background:var(--chipbg);color:var(--muted);cursor:pointer;max-width:78px}',
'.lblsel:hover,.bktsel:hover{border-color:var(--accent);color:var(--ink)}',
'.bktsel{font-family:"Spline Sans Mono",monospace;background:none}',
'.bktsel.now{color:var(--now);border-color:var(--now)}',
'.more{font-size:12px;line-height:1;color:var(--muted);padding:0 4px;border-radius:4px}',
'.more:hover{background:var(--chipbg);color:var(--ink)}',
'.card.done .t{text-decoration:line-through;color:var(--muted)}',
'.card .meta{flex:0 0 auto;display:flex;gap:4px;align-items:center;margin-top:0;flex-wrap:nowrap}',
'.lbl{font-size:9.5px;padding:0 5px;border-radius:99px;background:var(--chipbg);color:var(--muted);white-space:nowrap}',
'.bktbtn{font-size:9.5px;font-family:"Spline Sans Mono",monospace;padding:0 5px;border-radius:99px;border:1px solid var(--line);color:var(--muted)}',
'.bktbtn:hover{border-color:var(--accent);color:var(--accent)}',
'.bktbtn.now{color:var(--now);border-color:var(--now)}',
'.cmt{font-size:9.5px;font-family:"Spline Sans Mono",monospace;color:var(--muted)}',
'.prunetag{font-size:9px;color:var(--now);font-weight:600;text-transform:uppercase;letter-spacing:.05em}',
'.ckwrap{position:absolute;top:0;right:0;width:30px;height:100%;min-height:26px;display:flex;align-items:center;justify-content:center;cursor:pointer}',
'.ck{width:15px;height:15px;accent-color:var(--accent);cursor:pointer}',
'.whocancel{margin-top:12px}',
'.empty{font-size:11px;color:var(--muted);padding:2px 4px}',
/* overlay */
'.ovl{position:fixed;inset:0;background:rgba(15,22,19,.5);z-index:50;display:flex;align-items:flex-start;justify-content:center;padding:5vh 14px}',
'.panel{background:var(--surface);border-radius:16px;width:100%;max-width:560px;max-height:88vh;overflow:auto;padding:20px;box-shadow:0 16px 50px rgba(0,0,0,.35)}',
'.panel h3{font-family:"Bricolage Grotesque",sans-serif;font-size:13px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:10px}',
'.panel .title{width:100%;font-size:16.5px;font-weight:600;border:none;border-bottom:2px solid var(--line);padding:4px 2px 8px;background:none}',
'.panel .title:focus{border-bottom-color:var(--accent);outline:none}',
'.frow{display:flex;gap:10px;margin-top:14px;flex-wrap:wrap}',
'.frow>div{flex:1;min-width:120px}',
'.frow label{display:block;font-size:10.5px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:4px}',
'.frow select{width:100%;padding:8px;border:1px solid var(--line);border-radius:8px;background:var(--surface)}',
'.noteseA{width:100%;min-height:64px;margin-top:4px;padding:9px;border:1px solid var(--line);border-radius:8px;background:var(--surface);resize:vertical;font-size:13.5px}',
'.metaline{font-size:11.5px;font-family:"Spline Sans Mono",monospace;color:var(--muted);margin-top:12px}',
'.metaline a{color:var(--accent)}',
'.prunebox{margin-top:14px;padding:10px 12px;border-radius:10px;background:var(--now-soft);font-size:13px}',
'.prunebox b{color:var(--now)}',
'.pact{display:flex;gap:8px;margin-top:8px}',
'.btn{padding:8px 16px;border-radius:9px;font-weight:600;font-size:13.5px;border:1px solid var(--line);background:var(--surface2)}',
'.btn.pri{background:var(--accent);color:var(--accent-ink);border-color:var(--accent)}',
'.btn.warn{background:var(--now);color:var(--now-ink);border-color:var(--now)}',
'.btn.danger{color:var(--danger);border-color:var(--danger);background:none}',
'.comments{margin-top:18px;border-top:1px solid var(--line);padding-top:12px}',
'.c-item{margin-bottom:10px}',
'.c-head{font-size:11px;font-family:"Spline Sans Mono",monospace;color:var(--muted);margin-bottom:2px}',
'.c-head b{color:var(--accent);font-weight:600}',
'.c-text{font-size:13.5px;overflow-wrap:anywhere}',
'.c-new{display:flex;gap:8px;margin-top:10px}',
'.c-new input{flex:1;padding:9px 12px;border:1px solid var(--line);border-radius:9px;background:var(--surface)}',
'.pfoot{display:flex;gap:8px;margin-top:18px;align-items:center}',
'.pfoot .sp{flex:1}',
/* who modal + activity */
'.whobox{background:var(--surface);border-radius:16px;padding:26px;max-width:400px;width:100%;text-align:center}',
'.whobox h3{font-family:"Bricolage Grotesque",sans-serif;font-size:18px;margin-bottom:4px;text-transform:none;letter-spacing:0;color:var(--ink)}',
'.whobox p{font-size:13px;color:var(--muted);margin-bottom:14px}',
'.whogrid{display:flex;flex-wrap:wrap;gap:8px;justify-content:center}',
'.whogrid .pbtn{min-width:84px}',
'.alist{list-style:none}',
'.alist li{padding:7px 0;border-bottom:1px solid var(--line);font-size:13px}',
'.alist .aw{color:var(--accent);font-weight:600}',
'.alist .at{font-family:"Spline Sans Mono",monospace;font-size:10.5px;color:var(--muted);margin-left:6px}',
'@media(max-width:760px){.board{flex-direction:column}.lane{flex:1 1 auto;width:100%}.search{width:100%;margin-left:0}}',
].join('\n');

/* ---------- rendering ---------- */
function laneNames() { return STATE.lanes.slice(); }

function visibleTasks() {
  var q = VIEW.q.trim().toLowerCase();
  return STATE.tasks.filter(function (t) {
    if (VIEW.filter === 'done') { if (!t.done) return false; }
    else {
      if (t.done) return false;
      if (VIEW.filter !== 'all' && t.label !== VIEW.filter) return false;
    }
    if (!inScope(t)) return false;
    if (q) {
      var hay = (t.t + ' ' + (t.notes || '') + ' ' + t.label + ' ' + t.lane).toLowerCase();
      if (hay.indexOf(q) < 0) return false;
    }
    return true;
  });
}

/* Rows are draggable only in the bucket views - Done and Review render flat
   lists with no sections to drop into. */
function dragEnabled() {
  return !READONLY && VIEW.filter !== 'done';
}

function cardHTML(t) {
  var lbl = labelOf(t.label);
  var hue = lbl ? lbl.hue : 160;
  var cls = 'card' + (t.bucket === 'now' && !t.done ? ' now' : '') + (t.done ? ' done' : '');
  var canDrag = !t.done && dragEnabled();
  var h = '<div class="' + cls + '" data-tid="' + t.id + '"' + (canDrag ? ' data-drag="' + t.id + '"' : '') + '>';
  h += '<label class="ckwrap" data-ckwrap><input type="checkbox" class="ck" data-done="' + t.id + '"' + (t.done ? ' checked' : '') + ' aria-label="Mark done"></label>';
  if (canDrag) h += '<span class="grab" title="Drag to move (Now / Next / Later)">&#10303;</span>';
  h += '<div class="t" data-tedit="' + t.id + '" title="' + esc(t.t) + ' (click to edit)">' + esc(t.t) + '</div>';
  h += '<div class="meta">';
  /* label: pick it right on the row */
  h += '<span class="fdot" style="background:hsl(' + hue + ' 55% 45%)' + (t.label ? '' : ';opacity:.25') + '"></span>';
  h += '<select class="lblsel" data-edit-label-inline="' + t.id + '" title="Label"><option value=""' + (t.label ? '' : ' selected') + '>&middot;</option>';
  STATE.labels.forEach(function (l) { h += '<option' + (t.label === l.name ? ' selected' : '') + '>' + esc(l.name) + '</option>'; });
  h += '</select>';
  /* Now / Next / Later: pick it right on the row */
  if (!t.done) {
    h += '<select class="bktsel' + (t.bucket === 'now' ? ' now' : '') + '" data-edit-bucket-inline="' + t.id + '" title="Now / Next / Later">';
    BUCKETS.forEach(function (b) { h += '<option value="' + b + '"' + (t.bucket === b ? ' selected' : '') + '>' + (BUCKET_SHORT[b] || BUCKET_NAME[b]) + '</option>'; });
    h += '</select>';
  }
  if (t.comments && t.comments.length) h += '<span class="cmt">&#9993; ' + t.comments.length + '</span>';
  h += '<button class="more" data-open="' + t.id + '" title="Notes, updates, delete">&#8943;</button>';
  h += '</div></div>';
  return h;
}

function boardHTML() {
  var vis = visibleTasks();
  var h = '<div class="board">';
  var shown = 0;
  laneNames().forEach(function (lane) {
    if (laneHidden(lane)) return;                    /* hidden via the Show row (per person) */
    shown++;
    var mine = vis.filter(function (t) { return t.lane === lane; });
    var w = laneWidth(lane);
    h += '<section class="lane" data-lane="' + esc(lane) + '"' + (w ? ' style="flex-basis:' + w + 'px"' : '') + '>';
    h += '<h2>' + esc(lane) + ' <span class="n">' + openCount(lane) + ' open</span></h2>';
    if (VIEW.filter === 'done') {
      if (!mine.length) { h += '<div class="empty">none</div>'; }
      mine.forEach(function (t) { h += cardHTML(t); });
    } else {
      BUCKETS.forEach(function (b) {
        var bt = mine.filter(function (t) { return t.bucket === b; });
        /* empty sections stay in the DOM (hidden) so they can be drop targets while dragging */
        h += '<div class="bsec' + (bt.length ? '' : ' empty') + '" data-bucket="' + b + '"><div class="bhead' + (b === 'now' ? ' now' : b === 'tosort' ? ' tosort' : '') + '">' + BUCKET_NAME[b] + '</div>';
        bt.forEach(function (t) { h += cardHTML(t); });
        h += '</div>';
      });
      if (!mine.length) h += '<div class="empty">clear</div>';
    }
    h += '<div class="lgrip" data-grip="' + esc(lane) + '" title="Drag to resize &middot; double-click to reset"></div>';
    h += '</section>';
  });
  if (!shown) h += '<div class="empty">All lanes hidden &mdash; click a name in the Show row to bring it back.</div>';
  h += '</div>';
  return h;
}

function filtersHTML() {
  var h = '<div class="filters">';
  var sc = PREFS.scope || 'all';
  h += '<span class="seg">';
  [['all', 'All'], ['work', 'Work'], ['personal', 'Personal']].forEach(function (o) {
    h += '<button class="segb' + (sc === o[0] ? ' on' : '') + '" data-scope="' + o[0] + '">' + o[1] + '</button>';
  });
  h += '</span><span class="fsep"></span>';
  h += '<span class="fcap">Show</span>';
  laneNames().forEach(function (l) {
    var off = laneHidden(l);
    h += '<button class="lchip' + (off ? ' off' : '') + '" data-lane-toggle="' + esc(l) + '" title="' + (off ? 'Show ' : 'Hide ') + esc(l) + '">' + esc(l) + '</button>';
  });
  h += '<span class="fsep"></span>';
  var chips = [['all', 'All']];
  STATE.labels.forEach(function (l) {
    var p = PERSONAL_LABELS.indexOf(l.name) >= 0;
    if (sc === 'work' && p) return;
    if (sc === 'personal' && !p) return;
    chips.push([l.name, l.name]);
  });
  chips.push(['done', 'Done']);
  chips.forEach(function (c) {
    var on = VIEW.filter === c[0] ? ' on' : '';
    var dot = '';
    var lbl = labelOf(c[0]);
    if (lbl) dot = '<span class="fdot" style="background:hsl(' + lbl.hue + ' 55% 45%)"></span>';
    h += '<button class="fchip' + on + '" data-filter="' + esc(c[0]) + '">' + dot + esc(c[1]) + '</button>';
  });
  h += '<button class="fchip edit" data-projects title="Add, rename, recolour or remove projects">Projects&hellip;</button>';
  h += '<input class="search" data-search placeholder="Search&hellip;" value="' + esc(VIEW.q) + '">';
  h += '</div>';
  return h;
}

function quickAddHTML() {
  var h = '<div class="qa"><input data-qa placeholder="Type a task and press Enter&hellip;" maxlength="300">';
  h += '<button class="pbtn primary" data-qadd-sort title="Add to your own To Sort">Add</button>';
  h += '<span class="to">or to</span>';
  laneNames().forEach(function (l) {
    h += '<button class="pbtn" data-qadd="' + esc(l) + '">' + esc(l) + '</button>';
  });
  h += '</div>';
  return h;
}

function headerHTML() {
  var nowN = STATE.tasks.filter(function (t) { return t.bucket === 'now' && !t.done && inScope(t); }).length;
  var inboxN = STATE.tasks.filter(function (t) { return t.bucket === 'tosort' && !t.done && inScope(t); }).length;
  var openN = STATE.tasks.filter(function (t) { return !t.done && inScope(t); }).length;
  var d = new Date();
  var days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
  var h = '<div class="top">';
  h += '<span class="brand">KEI<span class="dot">&#9679;</span> ' + (STATE.board === 'personal' ? 'Personal' : 'Team') + ' Board</span>';
  h += '<span class="datechip">' + days[d.getDay()] + ' ' + shortDate(d.toISOString()) + '</span>';
  if (nowN) h += '<span class="stat hot">' + nowN + ' now</span>';
  h += '<span class="stat">' + openN + ' open</span>';
  if (inboxN) h += '<span class="stat sort">' + inboxN + ' to sort</span>';
  h += '<button class="stat" data-showact title="Recent activity">activity</button>';
  h += '<button class="stat" data-refresh title="Redraw the board now\u2019">&#8635; refresh</button>';
  if (SIGNEDIN) h += '<button class="stat" data-signout title="Sign out of this board on this device">' + esc(SIGNEDIN.name.split(' ')[0]) + ' &middot; sign out</button>';
  h += '<span class="savechip' + (SAVESTATE === 'offline' || SAVESTATE === 'failed' ? ' err' : '') + '" data-chip>' + chipText() + '</span>';
  h += '<button class="whobtn" data-who>' + (WHO ? esc(WHO) : 'Who are you?') + '</button>';
  h += '</div>';
  return h;
}
function chipText() {
  if (LOCALMODE) return 'preview';
  if (SAVESTATE === 'signedout') return 'signed out';
  if (READONLY) return 'view-only';
  if (SAVESTATE === 'dirty') return '&#9679; saving';
  if (SAVESTATE === 'saving') return 'saving&hellip;';
  if (SAVESTATE === 'saved') return 'saved &#10003;';
  if (SAVESTATE === 'offline') return esc(FAILMSG || 'save failed - retrying');
  if (SAVESTATE === 'failed') return esc(FAILMSG || 'not saved - reload the page');
  if (STORE_READY) return 'live &#9679;';
  return '';
}
function renderChip() {
  var el = document.querySelector('[data-chip]');
  if (el) { el.innerHTML = chipText(); el.className = 'savechip' + (SAVESTATE === 'offline' || SAVESTATE === 'failed' ? ' err' : ''); }
}

function overlayHTML() {
  var t = VIEW.open != null ? byId(VIEW.open) : null;
  if (!t) return '';
  var h = '<div class="ovl" data-ovl><div class="panel">';
  h += '<h3>Task</h3>';
  h += '<input class="title" data-edit-title value="' + esc(t.t) + '" maxlength="300">';
  h += '<div class="frow">';
  h += '<div><label>Who</label><select data-edit-lane>';
  laneNames().forEach(function (l) { h += '<option' + (t.lane === l ? ' selected' : '') + '>' + esc(l) + '</option>'; });
  h += '</select></div>';
  h += '<div><label>When</label><select data-edit-bucket>';
  BUCKETS.forEach(function (b) { h += '<option value="' + b + '"' + (t.bucket === b ? ' selected' : '') + '>' + BUCKET_NAME[b] + '</option>'; });
  h += '</select></div>';
  h += '<div><label>Label</label><select data-edit-label><option value="">none</option>';
  STATE.labels.forEach(function (l) { h += '<option' + (t.label === l.name ? ' selected' : '') + '>' + esc(l.name) + '</option>'; });
  h += '</select></div>';
  h += '</div>';
  h += '<div class="frow"><div style="flex:100%"><label>Notes</label>';
  h += '<textarea class="noteseA" data-edit-notes>' + esc(t.notes || '') + '</textarea></div></div>';
  h += '<div class="metaline">added ' + shortDate(t.created) + (t.by ? ' by ' + esc(t.by) : '') +
       (t.src ? ' &middot; <a href="' + esc(t.src) + '" target="_blank" rel="noopener">voice memo</a>' : '') +
       (t.done ? ' &middot; done ' + shortDate(t.doneAt) : '') + '</div>';
  h += '<div class="comments"><h3>Updates</h3>';
  (t.comments || []).forEach(function (c) {
    h += '<div class="c-item"><div class="c-head"><b>' + esc(c.who) + '</b> ' + ago(c.at) + '</div>' +
         '<div class="c-text">' + esc(c.text) + '</div></div>';
  });
  if (!(t.comments || []).length) h += '<div class="empty">No updates yet &mdash; add one below so the team can follow along.</div>';
  h += '<div class="c-new"><input data-cmt placeholder="e.g. called them, waiting on reply&hellip;" maxlength="500">' +
       '<button class="btn pri" data-cmtgo="' + t.id + '">Post</button></div>';
  h += '</div>';
  h += '<div class="pfoot">' +
       '<button class="btn ' + (t.done ? '' : 'pri') + '" data-toggledone="' + t.id + '">' + (t.done ? 'Reopen' : 'Mark done') + '</button>' +
       '<span class="sp"></span>' +
       '<button class="btn danger" data-del="' + t.id + '">Delete</button>' +
       '<button class="btn" data-close>Close</button></div>';
  h += '</div></div>';
  return h;
}

/* Projects (labels) editor. Renaming carries every task that uses the project
   with it; removing one clears it off those tasks and needs a second click. */
function projectsHTML() {
  if (!VIEW.projects) return '';
  var h = '<div class="ovl" data-projovl><div class="modal proj">';
  h += '<h3>Projects</h3>';
  h += '<div class="msub">These are the tags you can put on a task. Renaming one updates every task using it.</div>';
  STATE.labels.forEach(function (l) {
    var used = STATE.tasks.filter(function (t) { return t.label === l.name; }).length;
    h += '<div class="projrow">';
    h += '<button class="hueswatch" data-projhue="' + esc(l.name) + '" title="Click for the next colour" style="background:hsl(' + l.hue + ' 55% 45%)"></button>';
    h += '<input class="projname" data-projname="' + esc(l.name) + '" value="' + esc(l.name) + '" maxlength="24">';
    h += '<span class="projused">' + used + ' task' + (used === 1 ? '' : 's') + '</span>';
    if (VIEW.projDel === l.name) {
      h += '<button class="abtn danger" data-projdel-yes="' + esc(l.name) + '">Remove' + (used ? ' from ' + used : '') + '?</button>';
      h += '<button class="abtn" data-projdel-no>Cancel</button>';
    } else {
      h += '<button class="abtn" data-projdel="' + esc(l.name) + '" title="Remove this project">&times;</button>';
    }
    h += '</div>';
  });
  h += '<div class="projrow add"><input class="projname" data-newproj placeholder="New project name&hellip;" maxlength="24">';
  h += '<button class="abtn" data-projadd>Add</button></div>';
  h += '<div class="mrow"><button class="pbtn" data-projclose>Done</button></div>';
  h += '</div></div>';
  return h;
}

function whoHTML() {
  if (WHO || whoCallback == null) return '';
  var h = '<div class="ovl" data-whoovl><div class="whobox"><h3>Who are you?</h3><p>So your updates are signed. Saved on this device.</p><div class="whogrid">';
  laneNames().forEach(function (l) {
    if (l === 'Unassigned') return;
    h += '<button class="pbtn" data-setwho="' + esc(l) + '">' + esc(l) + '</button>';
  });
  h += '</div><button class="btn whocancel" data-whocancel>Cancel</button></div></div>';
  return h;
}

function activityHTML() {
  if (!VIEW.activity) return '';
  var h = '<div class="ovl" data-actovl><div class="panel"><h3>Recent activity</h3><ul class="alist">';
  if (!STATE.activity.length) h += '<li>Nothing yet.</li>';
  STATE.activity.forEach(function (a) {
    h += '<li><span class="aw">' + esc(a.who) + '</span> ' + esc(a.text) + '<span class="at">' + ago(a.at) + '</span></li>';
  });
  h += '</ul><div class="pfoot"><span class="sp"></span><button class="btn" data-closeact>Close</button></div></div></div>';
  return h;
}

var DRAFT_ATTRS = ['data-search', 'data-qa', 'data-cmt', 'data-edit-title', 'data-edit-notes', 'data-newproj'];
var LAST_OPEN = null;
function render() {
  var app = document.getElementById('app');
  /* Snapshot what is being typed (and the caret) so rebuilding the DOM never
     eats a half-written task, comment or edit. */
  var drafts = {}, focusOn = null, selS = null, selE = null;
  var ae = document.activeElement;
  DRAFT_ATTRS.forEach(function (a) {
    var n = document.querySelector('[' + a + ']');
    if (!n) return;
    if (n.value) drafts[a] = n.value;
    if (n === ae) {
      focusOn = a;
      try { selS = n.selectionStart; selE = n.selectionEnd; } catch (x) { selS = selE = null; }
    }
  });
  var sameTask = (LAST_OPEN === VIEW.open);

  var h = headerHTML();
  if (LOCALMODE) h += '<div class="banner warn">Preview mode &mdash; this copy can&#39;t save. Open the shared board link to make changes.</div>';
  if (SAVESTATE === 'signedout') h += '<div class="banner warn"><b>Signed out.</b> Your board is safe; nothing is being saved until you sign in again. <button class="pbtn primary" data-signin>Sign in with Microsoft</button></div>';
  if (READONLY) h += '<div class="banner warn"><b>View only.</b> Your account can read this board but not change it.</div>';
  else if (READONLY) h += '<div class="banner warn">View-only &mdash; nothing you change here is saved for anyone else. To make changes stick, ask Rory for edit access.</div>';
  h += quickAddHTML() + filtersHTML() + boardHTML() + overlayHTML() + projectsHTML() + whoHTML() + activityHTML();
  app.innerHTML = h;
  LAST_OPEN = VIEW.open;

  if (VIEW.d) {                                      /* first render after a reload */
    var stashed = VIEW.d; VIEW.d = null;
    Object.keys(stashed).forEach(function (a) {
      var n = document.querySelector('[' + a + ']');
      if (n && !n.value) n.value = stashed[a];
    });
  }
  Object.keys(drafts).forEach(function (a) {
    if (a === 'data-search') return;                 /* driven by VIEW.q */
    var isEdit = (a === 'data-edit-title' || a === 'data-edit-notes');
    if (isEdit && !sameTask) return;                 /* never bleed one task's text into another */
    var n = document.querySelector('[' + a + ']');
    if (!n) return;
    if (isEdit || !n.value) n.value = drafts[a];
  });
  /* Never leave the caret in the search box that sits BEHIND an open modal -
     keystrokes meant for the task would silently re-filter the board instead. */
  var modalOpen = (VIEW.open != null) || VIEW.activity || (!WHO && whoCallback != null);
  if (focusOn === 'data-search' && modalOpen) focusOn = null;
  if (focusOn) {
    var f = document.querySelector('[' + focusOn + ']');
    if (f) { f.focus(); if (selS != null) { try { f.setSelectionRange(selS, selE); } catch (x) {} } }
  }
}

/* ---------- mutations ---------- */
/* Every task has an owner from the moment it exists: no lane given means the
   person adding it. New tasks start in To Sort until someone promotes them. */
/* Which lane belongs to the signed-in person. Anyone not on the list (a guest,
   or a new account) simply gets the first lane, never someone else's. */
var LANE_BY_EMAIL = {
  'rory@kootenayenvironmental.ca': 'Rory',
  'anne@kootenayenvironmental.ca': 'Anne',
  'cam@kootenayenvironmental.ca': 'Cam',
  'allie@kootenayenvironmental.ca': 'Allie'
};
function laneForEmail(email) { return LANE_BY_EMAIL[String(email || '').toLowerCase()] || ''; }
function myLane() { return (WHO && STATE.lanes.indexOf(WHO) >= 0) ? WHO : STATE.lanes[0]; }
function addTask(text, lane, bucket) {
  lane = lane || myLane();
  bucket = bucket || 'tosort';
  var t = {
    id: newId(), t: text, lane: lane, bucket: bucket,
    label: '', notes: '', done: false, doneAt: null,
    created: nowISO(), by: WHO || '', src: '', comments: []
  };
  STATE.tasks.unshift(t);
  act('added “' + text.slice(0, 60) + '” for ' + lane + (bucket === 'tosort' ? ' (to sort)' : ''));
}

/* Click-to-edit a title in place: swap the text for an input, commit on
   Enter/blur, cancel on Escape. No overlay. */
function startTitleEdit(id) {
  var el = document.querySelector('[data-tedit="' + id + '"]');
  var t = byId(id);
  if (!el || !t || READONLY) return;
  var inp = document.createElement('input');
  inp.className = 'tinline';
  inp.value = t.t;
  inp.maxLength = 300;
  inp.setAttribute('data-tinline', id);
  el.replaceWith(inp);
  inp.focus();
  inp.select();
  var settled = false;
  function commit() {
    if (settled) return;
    settled = true;
    var v = inp.value.trim();
    if (v && v !== t.t) {
      needWho(function () { mutate(function () { t.t = v; act('edited a task title'); }); });
    } else {
      render();
    }
  }
  inp.addEventListener('keydown', function (ev) {
    ev.stopPropagation();
    if (ev.key === 'Enter') { ev.preventDefault(); commit(); }
    else if (ev.key === 'Escape') { settled = true; render(); }
  });
  inp.addEventListener('blur', commit);
}

/* Pick a hue for a new project that sits as far as possible from the ones in use. */
function nextHue() {
  var used = STATE.labels.map(function (l) { return l.hue; });
  if (!used.length) return 200;
  var best = 0, bestGap = -1;
  for (var h = 0; h < 360; h += 10) {
    var gap = 360;
    used.forEach(function (u) {
      var d = Math.abs(((h - u + 540) % 360) - 180);
      gap = Math.min(gap, 180 - d);
    });
    if (gap > bestGap) { bestGap = gap; best = h; }
  }
  return best;
}

/* Apply any project renames typed in the editor, carrying their tasks along. */
function commitProjectNames() {
  var inputs = document.querySelectorAll('[data-projname]');
  var renames = [];
  Array.prototype.forEach.call(inputs, function (inp) {
    var was = inp.getAttribute('data-projname'), now = inp.value.trim();
    if (!now || now === was) return;
    if (labelOf(now)) return;                       /* would collide with an existing project */
    renames.push([was, now]);
  });
  if (!renames.length) return;
  needWho(function () {
    mutate(function () {
      renames.forEach(function (r) {
        var l = labelOf(r[0]); if (!l) return;
        l.name = r[1];
        STATE.tasks.forEach(function (t) { if (t.label === r[0]) t.label = r[1]; });
        if (VIEW.filter === r[0]) VIEW.filter = r[1];
        act('renamed project “' + r[0] + '” to “' + r[1] + '”');
      });
    });
  });
}

/* Add whatever is typed in the quick-add box. With no lane it goes to To Sort,
   so capture never has to stop to decide who it is for. */
function quickAdd(lane, bucket) {
  var inp = document.querySelector('[data-qa]');
  var txt = inp ? inp.value.trim() : '';
  if (!txt) { if (inp) inp.focus(); return; }
  needWho(function () {
    mutate(function () { addTask(txt, lane, bucket); });
    var ni = document.querySelector('[data-qa]');
    if (ni) { ni.value = ''; ni.focus(); }
  });
}

/* ---------- read-only export for the team ----------
   A plain HTML page with no scripts, no controls and no save engine: whoever
   opens it sees the board as it stood, and cannot change anything. Grouped by
   person then priority, which is how someone reads their own list. */
/* Everything the team is allowed to see: open, on the board, and not tagged
   with one of the personal projects. Never conditional on the view. */


/* Same content as a plain list, for when HTML downloads aren't permitted. */

/* ---------- event wiring (delegated) ---------- */
document.addEventListener('click', function (e) {
  var el = e.target.closest('[data-done],[data-ckwrap],[data-scope],[data-signin],[data-signout],[data-lane-toggle],[data-refresh],[data-qadd-sort],[data-projects],[data-projclose],[data-projovl],[data-projdel],[data-projdel-yes],[data-projdel-no],[data-projhue],[data-projadd],[data-tedit],[data-tinline],[data-edit-label-inline],[data-edit-bucket-inline],[data-qadd],[data-filter],[data-assign],[data-open],[data-cycle],[data-del],[data-close],[data-cmtgo],[data-toggledone],[data-setwho],[data-who],[data-whocancel],[data-whoovl],[data-showact],[data-closeact],[data-ovl],[data-actovl],.card');
  if (!el) return;

  /* inline row controls manage themselves - never let them open the task */
  if (el.hasAttribute('data-tinline') || el.hasAttribute('data-edit-label-inline') || el.hasAttribute('data-edit-bucket-inline')) return;
  if (el.hasAttribute('data-tedit')) { startTitleEdit(+el.getAttribute('data-tedit')); return; }

  /* The done checkbox (and its padded hit-target) must NOT fall through to the
     .card branch - that opened the task and re-rendered before 'change' fired,
     silently dropping the toggle. Let the native change handler commit it. */
  if (el.hasAttribute('data-done') || el.hasAttribute('data-ckwrap')) return;

  if (el.hasAttribute('data-setwho')) { setWho(el.getAttribute('data-setwho')); return; }
  if (el.hasAttribute('data-who')) { whoPrev = WHO; WHO = null; try { localStorage.removeItem('kei_who'); } catch (x) {} needWho(function () {}); return; }
  if (el.hasAttribute('data-whocancel')) { cancelWho(); return; }
  if (el.hasAttribute('data-whoovl')) { if (e.target === el) cancelWho(); return; }
  if (el.hasAttribute('data-showact')) { VIEW.activity = true; render(); return; }
  if (el.hasAttribute('data-closeact')) { VIEW.activity = false; render(); return; }
  if (el.hasAttribute('data-actovl')) { if (e.target === el) { VIEW.activity = false; render(); } return; }

  if (el.hasAttribute('data-qadd-sort')) { quickAdd(null, 'tosort'); return; }
  if (el.hasAttribute('data-qadd')) {
    quickAdd(el.getAttribute('data-qadd'), 'tosort');
    return;
  }
  if (el.hasAttribute('data-refresh')) { render(); return; }
  if (el.hasAttribute('data-signin')) { AUTH.signIn(); return; }
  if (el.hasAttribute('data-signout')) { AUTH.signOut(); location.reload(); return; }
  if (el.hasAttribute('data-projects')) { VIEW.projects = true; VIEW.projDel = null; render(); return; }
  if (el.hasAttribute('data-projclose')) { commitProjectNames(); VIEW.projects = false; VIEW.projDel = null; render(); return; }
  if (el.hasAttribute('data-projovl')) {
    if (e.target === el) { commitProjectNames(); VIEW.projects = false; VIEW.projDel = null; render(); }
    return;
  }
  if (el.hasAttribute('data-projdel')) { VIEW.projDel = el.getAttribute('data-projdel'); render(); return; }
  if (el.hasAttribute('data-projdel-no')) { VIEW.projDel = null; render(); return; }
  if (el.hasAttribute('data-projdel-yes')) {
    var dn = el.getAttribute('data-projdel-yes');
    needWho(function () {
      mutate(function () {
        STATE.labels = STATE.labels.filter(function (l) { return l.name !== dn; });
        STATE.tasks.forEach(function (t) { if (t.label === dn) t.label = ''; });
        if (VIEW.filter === dn) VIEW.filter = 'all';
        VIEW.projDel = null;
        act('removed the project “' + dn + '”');
      });
    });
    return;
  }
  if (el.hasAttribute('data-projhue')) {
    var hn = el.getAttribute('data-projhue');
    needWho(function () {
      mutate(function () {
        var l = labelOf(hn); if (!l) return;
        l.hue = (l.hue + 40) % 360;
      });
    });
    return;
  }
  if (el.hasAttribute('data-projadd')) {
    var ni = document.querySelector('[data-newproj]');
    var nm = ni ? ni.value.trim() : '';
    if (!nm) { if (ni) ni.focus(); return; }
    if (labelOf(nm)) { if (ni) { ni.value = ''; ni.focus(); } return; }   /* already exists */
    needWho(function () {
      mutate(function () {
        STATE.labels.push({ name: nm, hue: nextHue() });
        act('added the project “' + nm + '”');
      });
      var n2 = document.querySelector('[data-newproj]');
      if (n2) { n2.value = ''; n2.focus(); }
    });
    return;
  }
  if (el.hasAttribute('data-scope')) {
    PREFS.scope = el.getAttribute('data-scope');
    /* a project filter from the other side would show nothing - drop it */
    if (VIEW.filter !== 'all' && VIEW.filter !== 'done') {
      var p = PERSONAL_LABELS.indexOf(VIEW.filter) >= 0;
      if ((PREFS.scope === 'work' && p) || (PREFS.scope === 'personal' && !p)) VIEW.filter = 'all';
    }
    savePrefs(); render(); return;
  }
  if (el.hasAttribute('data-lane-toggle')) {
    var ln = el.getAttribute('data-lane-toggle');
    if (PREFS.hide[ln]) delete PREFS.hide[ln]; else PREFS.hide[ln] = true;
    savePrefs(); render(); return;
  }
  if (el.hasAttribute('data-filter')) { VIEW.filter = el.getAttribute('data-filter'); render(); return; }

  if (el.hasAttribute('data-assign')) {
    var id = +el.getAttribute('data-assign'), lane2 = el.getAttribute('data-lane');
    needWho(function () {
      mutate(function () {
        var t = byId(id); if (!t) return;
        t.lane = lane2; t.bucket = 'next';
        act('filed “' + t.t.slice(0, 50) + '” to ' + lane2);
      });
    });
    return;
  }
  if (el.hasAttribute('data-open')) { VIEW.open = +el.getAttribute('data-open'); render(); return; }
  if (el.hasAttribute('data-cycle')) {
    e.stopPropagation();
    var id2 = +el.getAttribute('data-cycle');
    needWho(function () {
      mutate(function () {
        var t = byId(id2); if (!t) return;
        var i = BUCKETS.indexOf(t.bucket);
        t.bucket = BUCKETS[(i + 1) % 3];
        act('moved “' + t.t.slice(0, 40) + '” to ' + BUCKET_NAME[t.bucket]);
      });
    });
    return;
  }
  if (el.hasAttribute('data-del')) {
    var id5 = +el.getAttribute('data-del');
    if (!window.confirm('Delete this task for everyone?')) return;
    needWho(function () {
      mutate(function () {
        var t = byId(id5); if (!t) return;
        STATE.tasks = STATE.tasks.filter(function (x) { return x.id !== id5; });
        act('deleted “' + t.t.slice(0, 40) + '”');
        VIEW.open = null;
      });
    });
    return;
  }
  if (el.hasAttribute('data-toggledone')) {
    var id6 = +el.getAttribute('data-toggledone');
    needWho(function () {
      mutate(function () {
        var t = byId(id6); if (!t) return;
        t.done = !t.done;
        t.doneAt = t.done ? nowISO() : null;
        if (t.done) VIEW.open = null;
        act((t.done ? 'completed' : 'reopened') + ' “' + t.t.slice(0, 40) + '”');
      });
    });
    return;
  }
  if (el.hasAttribute('data-cmtgo')) {
    var id7 = +el.getAttribute('data-cmtgo');
    var ci = document.querySelector('[data-cmt]');
    var text = ci ? ci.value.trim() : '';
    if (!text) { if (ci) ci.focus(); return; }
    needWho(function () {
      mutate(function () {
        var t = byId(id7); if (!t) return;
        t.comments = t.comments || [];
        t.comments.push({ who: WHO, at: nowISO(), text: text });
        act('updated “' + t.t.slice(0, 40) + '”: ' + text.slice(0, 50));
      });
    });
    return;
  }
  if (el.hasAttribute('data-close')) { commitOpenEdits(); VIEW.open = null; render(); return; }
  if (el.hasAttribute('data-ovl')) { if (e.target === el) { commitOpenEdits(); VIEW.open = null; render(); } return; }

  if (el.classList.contains('card')) {
    VIEW.open = +el.getAttribute('data-tid');
    render();
    return;
  }
});

/* Drag a task row up/down into Now / Next / Later (or into another lane's
   section). Press anywhere on the row and move >= 6px to lift it - a plain
   click still edits/opens as before. Controls on the row (checkbox, title
   editor, dropdowns, ... button) never start a drag. Touch users drag by the
   grab handle (touch-action:none there only, so the page still scrolls). */
(function () {
  var THRESH = 6;
  var d = null;            /* the drag in progress */
  var swallow = false;     /* eat the click that follows a real drag */

  function cardFrom(target) {
    if (!target || !target.closest) return null;
    if (target.closest('input,select,button,textarea,label,.tinline,[data-grip],.ghost')) return null;
    return target.closest('.card[data-drag]');
  }
  function sectionAt(x, y) {
    d.ghost.style.display = 'none';
    var el = document.elementFromPoint(x, y);
    d.ghost.style.display = '';
    if (!el || !el.closest) return null;
    var sec = el.closest('.bsec');
    if (sec) return sec;
    var lane = el.closest('.lane');          /* between sections: snap to the nearest one */
    if (!lane) return null;
    var best = null, bd = 1e9;
    lane.querySelectorAll('.bsec').forEach(function (s) {
      var r = s.getBoundingClientRect();
      var dy = y < r.top ? r.top - y : (y > r.bottom ? y - r.bottom : 0);
      if (dy < bd) { bd = dy; best = s; }
    });
    return best;
  }
  var EDGE = 64, SPEED = 18;
  function autoScroll() {
    if (!d || !d.live) return;
    var moved = false, before;
    var board = document.querySelector('.board');
    if (board) {
      var br = board.getBoundingClientRect();
      if (d.lx < br.left + EDGE) {
        before = board.scrollLeft; board.scrollLeft -= SPEED;
        if (board.scrollLeft !== before) moved = true;
      } else if (d.lx > br.right - EDGE) {
        before = board.scrollLeft; board.scrollLeft += SPEED;
        if (board.scrollLeft !== before) moved = true;
      }
    }
    if (d.ly < EDGE) {
      before = window.pageYOffset; window.scrollBy(0, -SPEED);
      if (window.pageYOffset !== before) moved = true;
    } else if (d.ly > window.innerHeight - EDGE) {
      before = window.pageYOffset; window.scrollBy(0, SPEED);
      if (window.pageYOffset !== before) moved = true;
    }
    if (moved) setOver(sectionAt(d.lx, d.ly));
  }

  function begin(e) {
    d.live = true;
    try { d.card.setPointerCapture(d.pid); } catch (x) {}
    try { getSelection().removeAllRanges(); } catch (x) {}
    var r = d.card.getBoundingClientRect();                    /* measure BEFORE the empty sections appear */
    d.dx = d.x0 - r.left; d.dy = d.y0 - r.top;                 /* offset from where the user PRESSED, not from the move that tripped the threshold */
    document.body.classList.add('dragging');
    var g = document.createElement('div');
    g.className = 'card ghost' + (d.card.classList.contains('now') ? ' now' : '');
    g.style.width = r.width + 'px';
    g.innerHTML = '<span class="grab">&#10303;</span><div class="t">' + esc(d.task.t) + '</div>';
    document.body.appendChild(g);
    d.ghost = g;
    d.card.classList.add('lifted');
    moveGhost(e);
    d.raf = setInterval(autoScroll, 16);            /* let the pointer drag the board to off-screen lanes */
  }
  function moveGhost(e) {
    d.ghost.style.left = (e.clientX - d.dx) + 'px';
    d.ghost.style.top = (e.clientY - d.dy) + 'px';
  }
  function setOver(sec) {
    if (sec === d.over) return;
    if (d.over) d.over.classList.remove('over');
    d.over = sec;
    if (sec) sec.classList.add('over');
  }
  function finish(cancel) {
    var was = d; d = null;
    if (was && was.raf) clearInterval(was.raf);
    if (!was || !was.live) return;                    /* never lifted: the click goes through */
    if (was.ghost && was.ghost.parentNode) was.ghost.parentNode.removeChild(was.ghost);
    was.card.classList.remove('lifted');
    if (was.over) was.over.classList.remove('over');
    document.body.classList.remove('dragging');
    try { was.card.releasePointerCapture(was.pid); } catch (x) {}
    swallow = true;         /* cleared by the click it eats, or by the next pointerdown */
    if (cancel || !was.over) return;
    var bucket = was.over.getAttribute('data-bucket');
    var laneEl = was.over.closest('.lane');
    var lane = laneEl ? laneEl.getAttribute('data-lane') : null;
    if (BUCKETS.indexOf(bucket) < 0 || !lane || STATE.lanes.indexOf(lane) < 0) return;
    var id = was.id;
    var t0 = byId(id);
    if (!t0 || t0.done) return;                                  /* vanished or completed mid-drag */
    if (t0.bucket === bucket && t0.lane === lane) return;        /* dropped where it already was: no prompt, no publish */
    var bucketChanged = t0.bucket !== bucket, laneChanged = t0.lane !== lane;
    var msg = 'moved \u201c' + t0.t.slice(0, 40) + '\u201d to ' +
      (bucketChanged && laneChanged ? BUCKET_NAME[bucket] + ' (' + lane + ')'
        : bucketChanged ? BUCKET_NAME[bucket] : lane);
    needWho(function () {
      mutate(function () {
        var t = byId(id); if (!t || t.done) return;
        if (t.bucket === bucket && t.lane === lane) return;
        t.bucket = bucket; t.lane = lane;
        act(msg);
      });
    });
  }

  document.addEventListener('pointerdown', function (e) {
    swallow = false;
    if (d || e.button !== 0 || READONLY) return;
    if (!dragEnabled()) return;
    var c = cardFrom(e.target); if (!c) return;
    var t = byId(+c.getAttribute('data-drag')); if (!t || t.done) return;
    d = { id: t.id, task: t, card: c, pid: e.pointerId, x0: e.clientX, y0: e.clientY, lx: e.clientX, ly: e.clientY, live: false, ghost: null, over: null, dx: 0, dy: 0, raf: 0 };
  });
  document.addEventListener('pointermove', function (e) {
    if (!d || e.pointerId !== d.pid) return;
    d.lx = e.clientX; d.ly = e.clientY;
    if (!d.live) {
      if (e.buttons === 0) { d = null; return; }                 /* the button came up unseen */
      if (Math.abs(e.clientX - d.x0) < THRESH && Math.abs(e.clientY - d.y0) < THRESH) return;
      if (!document.contains(d.card)) { d = null; return; }   /* re-rendered under us */
      begin(e);
    }
    e.preventDefault();
    moveGhost(e);
    setOver(sectionAt(e.clientX, e.clientY));
  });
  document.addEventListener('scroll', function () {
    if (d && d.live) setOver(sectionAt(d.lx, d.ly));             /* the board moved under the pointer */
  }, true);
  document.addEventListener('pointerup', function (e) { if (d && e.pointerId === d.pid) finish(false); });
  document.addEventListener('pointercancel', function (e) { if (d && e.pointerId === d.pid) finish(true); });
  document.addEventListener('keydown', function (e) {
    if (d && d.live && e.key === 'Escape') { e.stopPropagation(); e.preventDefault(); finish(true); }
  }, true);
  document.addEventListener('click', function (e) {
    if (swallow) { swallow = false; e.stopPropagation(); e.preventDefault(); }
  }, true);
  window.addEventListener('blur', function () { if (d) finish(true); });
})();

/* Drag a lane's right edge to resize it; double-click the edge to reset.
   The width is a per-person preference (PREFS), not board data. */
(function () {
  var drag = null;
  document.addEventListener('pointerdown', function (e) {
    var g = e.target.closest && e.target.closest('[data-grip]');
    if (!g || e.button !== 0) return;
    var lane = g.parentNode;
    drag = { name: g.getAttribute('data-grip'), el: lane, grip: g, x0: e.clientX, w0: lane.getBoundingClientRect().width, w: 0 };
    g.classList.add('on');
    document.body.classList.add('resizing');
    try { g.setPointerCapture(e.pointerId); } catch (x) {}
    e.preventDefault();
  });
  document.addEventListener('pointermove', function (e) {
    if (!drag) return;
    var w = Math.round(Math.min(LANE_MAX, Math.max(LANE_MIN, drag.w0 + (e.clientX - drag.x0))));
    drag.el.style.flexBasis = w + 'px';
    drag.w = w;
  });
  function endDrag() {
    if (!drag) return;
    if (drag.w) { PREFS.w[drag.name] = drag.w; savePrefs(); }
    drag.grip.classList.remove('on');
    document.body.classList.remove('resizing');
    drag = null;
  }
  document.addEventListener('pointerup', endDrag);
  document.addEventListener('pointercancel', endDrag);
  document.addEventListener('dblclick', function (e) {
    var g = e.target.closest && e.target.closest('[data-grip]');
    if (!g) return;
    delete PREFS.w[g.getAttribute('data-grip')];
    savePrefs();
    g.parentNode.style.flexBasis = '';
  });
})();

document.addEventListener('change', function (e) {
  var el = e.target;
  if (el.hasAttribute('data-done')) {
    var id = +el.getAttribute('data-done');
    needWho(function () {
      mutate(function () {
        var t = byId(id); if (!t) return;
        t.done = el.checked;
        t.doneAt = t.done ? nowISO() : null;
        act((t.done ? 'completed' : 'reopened') + ' “' + t.t.slice(0, 40) + '”');
      });
    });
    return;
  }
  if (el.hasAttribute('data-edit-label-inline')) {
    var lid = +el.getAttribute('data-edit-label-inline'), lv = el.value;
    needWho(function () { mutate(function () { var lt = byId(lid); if (lt) lt.label = lv; }); });
    return;
  }
  if (el.hasAttribute('data-edit-bucket-inline')) {
    var bid = +el.getAttribute('data-edit-bucket-inline'), bv = el.value;
    needWho(function () { mutate(function () {
      var bt = byId(bid); if (!bt) return;
      bt.bucket = bv;
      act('moved “' + bt.t.slice(0, 40) + '” to ' + (BUCKET_NAME[bv] || bv));
    }); });
    return;
  }
  if (VIEW.open != null) {
    var t = byId(VIEW.open);
    if (!t) return;
    if (el.hasAttribute('data-edit-lane')) { needWho(function () { mutate(function () { t.lane = el.value; act('reassigned “' + t.t.slice(0, 40) + '” to ' + el.value); }); }); }
    else if (el.hasAttribute('data-edit-bucket')) { needWho(function () { mutate(function () { t.bucket = el.value; act('moved “' + t.t.slice(0, 40) + '” to ' + (BUCKET_NAME[el.value] || 'Inbox')); }); }); }
    else if (el.hasAttribute('data-edit-label')) { needWho(function () { mutate(function () { t.label = el.value; }); }); }
    else if (el.hasAttribute('data-edit-title')) { needWho(function () { mutateQuiet(function () { var v = el.value.trim(); if (v) { t.t = v; act('edited a task title'); } }); }); }
    else if (el.hasAttribute('data-edit-notes')) { needWho(function () { mutateQuiet(function () { t.notes = el.value.trim(); }); }); }
  }
});

document.addEventListener('input', function (e) {
  if (e.target.hasAttribute && e.target.hasAttribute('data-search')) {
    VIEW.q = e.target.value;
    render();   /* render() restores focus + caret from the live DOM */
  }
});

document.addEventListener('keydown', function (e) {
  if (e.key === 'Escape') {
    if (!WHO && whoCallback != null) { cancelWho(); return; }
    if (VIEW.projects) { commitProjectNames(); VIEW.projects = false; VIEW.projDel = null; render(); return; }
    if (VIEW.open != null || VIEW.activity) {
      commitOpenEdits();   /* never throw away an in-progress edit */
      VIEW.open = null; VIEW.activity = false; render();
    }
  }
  if (e.key === 'Enter' && e.target.hasAttribute && e.target.hasAttribute('data-qa')) {
    e.preventDefault(); quickAdd(null, 'tosort'); return;
  }
  if (e.key === 'Enter' && e.target.hasAttribute && e.target.hasAttribute('data-newproj')) {
    e.preventDefault();
    var ab = document.querySelector('[data-projadd]'); if (ab) ab.click();
    return;
  }
  if (e.key === 'Enter' && e.target.hasAttribute && e.target.hasAttribute('data-cmt')) {
    var b = document.querySelector('[data-cmtgo]');
    if (b) b.click();
  }
});

/* ---------- keeping the view in step ----------
   Other people's changes arrive on the store subscriptions and land in STATE
   the moment they happen, so there is nothing to poll for and no reload: the
   only judgement left is WHEN to redraw, which safeToRender() answers. */
function stashView() {
  try {
    var d = {};
    ['data-qa', 'data-cmt'].forEach(function (a) {
      var n = document.querySelector('[' + a + ']');
      if (n && n.value) d[a] = n.value;
    });
    VIEW.d = Object.keys(d).length ? d : null;
    sessionStorage.setItem('kei_view', JSON.stringify(VIEW));
  } catch (e) {}
}
window.addEventListener('beforeunload', stashView);

/* ---------- boot ---------- */
(function boot() {
  var style = document.createElement('style');
  style.textContent = CSS;
  document.head.appendChild(style);
  if (!document.querySelector('link[href*="fonts.googleapis"]')) {
    var tmp = document.createElement('div');
    tmp.innerHTML = FONT_LINKS;
    while (tmp.firstChild) document.head.appendChild(tmp.firstChild);
  }
  document.title = TITLE;
  var app = document.getElementById('app') || document.createElement('div');
  app.id = 'app';
  if (!app.parentNode) document.body.appendChild(app);
  /* Nothing is baked into the page: sign in, then fill the board from the
     lists. Until then the chip says so rather than showing a blank board. */
  render();
  connect();
})();
