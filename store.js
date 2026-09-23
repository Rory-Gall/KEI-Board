/* The board's storage, on SharePoint Lists.
 *
 * The app already saves one task at a time (see pushChanges in app.js). This
 * file is the transport underneath that: read everything once, write single
 * tasks, delete single tasks, and poll for what other people changed.
 *
 *   STORE.load()            -> {tasks, settings, activity, personalReadable}
 *   STORE.put(task)         -> writes one task, resolves to the stored task
 *   STORE.remove(task)      -> deletes one task
 *   STORE.putSetting(k, v)  -> board settings (lanes, labels, personal labels)
 *   STORE.addActivity(e)    -> one activity line
 *   STORE.poll()            -> [{task}|{deletedId}] since the last call
 *
 * Which list a task lives in is decided by its label, not by a flag on the row:
 * anything tagged personal (or in the Contractors lane) belongs to the private
 * Personal Tasks list, everything else to KEI Tasks. Staff simply cannot read
 * the personal list, so for them it does not exist.
 */
var STORE = (function () {
  'use strict';

  var GRAPH = 'https://graph.microsoft.com/v1.0';
  var SITE = '/sites/' + CONFIG.siteId;
  var TEAM = 'KEI Tasks', PERSONAL = 'Personal Tasks';

  /* The board's buckets are keys; the list shows readable names. */
  var BUCKET_OUT = { tosort: 'To Sort', now: 'Now', progress: 'In Progress / Waiting On',
                     today: 'Today', week: 'This Week', later: 'Later', someday: 'Some Day / Maybe' };
  var BUCKET_IN = {};
  Object.keys(BUCKET_OUT).forEach(function (k) { BUCKET_IN[BUCKET_OUT[k]] = k; });

  var personalLabels = ['Personal', 'House', 'Motorcycle', 'Yard'];
  var personalLanes = ['Contractors'];
  var deltaLink = {};       /* list name -> Microsoft's "what changed since" marker */
  var rowOf = {};           /* task id -> {list, itemId} so an edit knows where to go */
  var personalReadable = true;

  function listId(name) { return CONFIG.lists[name]; }
  function listFor(t) {
    return (personalLabels.indexOf(t.label) >= 0 || personalLanes.indexOf(t.lane) >= 0) ? PERSONAL : TEAM;
  }

  function call(method, url, body) {
    return AUTH.token().then(function (tok) {
      var opt = { method: method, headers: { Authorization: 'Bearer ' + tok, Accept: 'application/json' } };
      if (body !== undefined) {
        opt.headers['Content-Type'] = 'application/json';
        opt.body = JSON.stringify(body);
      }
      return fetch(url.indexOf('https://') === 0 ? url : GRAPH + url, opt);
    }).then(function (r) {
      if (r.status === 204) return {};
      if (r.status === 429 || r.status >= 500) {
        var wait = (Number(r.headers.get('Retry-After')) || 4) * 1000;
        return new Promise(function (res) { setTimeout(res, wait); })
          .then(function () { return call(method, url, body); });
      }
      return r.json().then(function (j) {
        if (!r.ok) {
          var e = new Error((j.error && j.error.message) || ('HTTP ' + r.status));
          e.status = r.status;
          e.code = j.error && j.error.code;
          throw e;
        }
        return j;
      });
    });
  }

  function pages(url) {
    var out = [];
    function next(u) {
      return call('GET', u).then(function (page) {
        out = out.concat(page.value || []);
        if (page['@odata.nextLink']) return next(page['@odata.nextLink']);
        if (page['@odata.deltaLink']) out.deltaLink = page['@odata.deltaLink'];
        return out;
      });
    }
    return next(url);
  }

  /* ---------- shape conversion ---------- */
  function toFields(t) {
    return {
      Title: String(t.t || '').slice(0, 255),       /* SharePoint's own cap */
      Lane: t.lane || '',
      Bucket: BUCKET_OUT[t.bucket] || 'To Sort',
      Label: t.label || '',
      Notes: t.notes || '',
      Done: !!t.done,
      DoneAt: t.doneAt || null,
      CreatedAt: t.created || null,
      AddedBy: t.by || '',
      Source: t.src || '',
      Comments: JSON.stringify(t.comments || []),
      LegacyId: String(t.id)      /* ALWAYS a string: a number here fails with HTTP 500 */
    };
  }
  function fromFields(f, listName, itemId) {
    var comments = [];
    if (f.Comments) {
      try { comments = JSON.parse(f.Comments) || []; }
      catch (e) { comments = []; BAD_COMMENTS.push(itemId); }
    }
    var id = Number(f.LegacyId);
    var t = {
      id: isFinite(id) && id ? id : Number(itemId),
      t: f.Title || '',
      lane: f.Lane || '',
      bucket: BUCKET_IN[f.Bucket] || 'tosort',
      label: f.Label || '',
      notes: f.Notes || '',
      done: !!f.Done,
      doneAt: f.DoneAt || null,
      created: f.CreatedAt || f.Created || null,
      by: f.AddedBy || '',
      src: f.Source || '',
      comments: comments
    };
    rowOf[t.id] = { list: listName, itemId: itemId };
    return t;
  }
  var BAD_COMMENTS = [];   /* surfaced by load(), never swallowed */

  /* ---------- reading ---------- */
  function readList(name) {
    return pages(SITE + '/lists/' + listId(name) + '/items?expand=fields&$top=200')
      .then(function (items) {
        return items.map(function (i) { return fromFields(i.fields || {}, name, i.id); });
      })
      .catch(function (e) {
        /* Staff get 403 on the personal list. That is the design, not a fault. */
        if (name === PERSONAL && (e.status === 403 || e.status === 404)) {
          personalReadable = false;
          return [];
        }
        throw e;
      });
  }

  function markDelta(name) {
    return call('GET', SITE + '/lists/' + listId(name) + '/items/delta?token=latest')
      .then(function (d) { deltaLink[name] = d['@odata.deltaLink'] || null; })
      .catch(function () { deltaLink[name] = null; });
  }

  function load() {
    BAD_COMMENTS = [];
    return Promise.all([
      readList(TEAM),
      readList(PERSONAL),
      pages(SITE + '/lists/' + listId('Board Settings') + '/items?expand=fields&$top=100'),
      pages(SITE + '/lists/' + listId('Activity') + '/items?expand=fields&$top=60&$orderby=fields/At desc')
        .catch(function () { return []; })
    ]).then(function (r) {
      var settings = {};
      r[2].forEach(function (i) {
        var f = i.fields || {};
        try { settings[f.Title] = JSON.parse(f.Value); } catch (e) { settings[f.Title] = f.Value; }
      });
      var activity = r[3].map(function (i) {
        var f = i.fields || {};
        return { id: i.id, who: f.Who || '', at: f.At || '', text: f.What || '' };
      });
      return Promise.all([markDelta(TEAM), personalReadable ? markDelta(PERSONAL) : null])
        .then(function () {
          return {
            tasks: r[0].concat(r[1]),
            settings: settings,
            activity: activity,
            personalReadable: personalReadable,
            badComments: BAD_COMMENTS.slice()
          };
        });
    });
  }

  /* ---------- writing ---------- */
  function put(t) {
    var want = listFor(t), at = rowOf[t.id];
    var fields = toFields(t);
    /* A task that changed sides (e.g. tagged House) moves list: write it to the
       new one first, and only then remove the old row, so a failure halfway
       leaves a duplicate rather than a hole. */
    if (at && at.list !== want) {
      return create(want, fields, t.id).then(function () {
        return call('DELETE', SITE + '/lists/' + listId(at.list) + '/items/' + at.itemId);
      }).then(function () { return t; });
    }
    if (!at) return create(want, fields, t.id).then(function () { return t; });
    return call('PATCH', SITE + '/lists/' + listId(at.list) + '/items/' + at.itemId + '/fields', fields)
      .then(function () { return t; });
  }
  function create(listName, fields, id) {
    return call('POST', SITE + '/lists/' + listId(listName) + '/items', { fields: fields })
      .then(function (made) {
        rowOf[id] = { list: listName, itemId: made.id };
        return made;
      });
  }
  function remove(t) {
    var at = rowOf[t.id];
    if (!at) return Promise.resolve();
    return call('DELETE', SITE + '/lists/' + listId(at.list) + '/items/' + at.itemId)
      .then(function () { delete rowOf[t.id]; });
  }
  /* Read the whole settings list and match here. Graph refuses to filter on a
     column that is not indexed, and this list holds a handful of rows. */
  function putSetting(key, value) {
    var fields = { Title: key, Value: JSON.stringify(value) };
    return pages(SITE + '/lists/' + listId('Board Settings') + '/items?expand=fields&$top=100')
      .then(function (items) {
        var found = items.filter(function (i) { return (i.fields || {}).Title === key; })[0];
        if (found) {
          return call('PATCH', SITE + '/lists/' + listId('Board Settings') + '/items/' + found.id + '/fields', fields);
        }
        return call('POST', SITE + '/lists/' + listId('Board Settings') + '/items', { fields: fields });
      });
  }
  function addActivity(e) {
    return call('POST', SITE + '/lists/' + listId('Activity') + '/items',
                { fields: { Title: (e.text || '').slice(0, 255), Who: e.who || '', At: e.at, What: e.text || '' } });
  }

  /* ---------- what other people changed ---------- */
  function pollList(name) {
    if (!deltaLink[name]) return Promise.resolve([]);
    return pages(deltaLink[name]).then(function (items) {
      if (items.deltaLink) deltaLink[name] = items.deltaLink;
      var out = [];
      items.forEach(function (i) {
        if (i.deleted) {
          var gone = Object.keys(rowOf).filter(function (id) {
            return rowOf[id].list === name && String(rowOf[id].itemId) === String(i.id);
          });
          gone.forEach(function (id) { delete rowOf[id]; out.push({ deletedId: Number(id) }); });
          return;
        }
        if (!i.fields) return;
        out.push({ task: fromFields(i.fields, name, i.id) });
      });
      return out;
    }).catch(function (e) {
      /* A delta marker can expire; start a fresh one and let the next poll work. */
      if (e.status === 410 || e.code === 'resyncRequired') return markDelta(name).then(function () { return []; });
      throw e;
    });
  }
  function poll() {
    return Promise.all([pollList(TEAM), personalReadable ? pollList(PERSONAL) : Promise.resolve([])])
      .then(function (r) { return r[0].concat(r[1]); });
  }

  return {
    load: load, put: put, remove: remove, putSetting: putSetting, addActivity: addActivity, poll: poll,
    setPersonalLabels: function (list) { if (list && list.length) personalLabels = list.slice(); },
    listFor: listFor,
    canReadPersonal: function () { return personalReadable; }
  };
})();
