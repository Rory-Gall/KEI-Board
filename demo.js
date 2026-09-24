/* Demo mode: the board with stand-in data and no Microsoft at all.
 *
 *   index.html?demo=1
 *
 * It replaces AUTH and STORE with in-memory versions, so the whole board can be
 * exercised - drag, edit, tick, filter, comment - without a sign-in and without
 * touching a real task. Used to check the interface; never part of the real
 * board's behaviour, and never loaded unless ?demo=1 is in the address.
 */
(function () {
  'use strict';

  var now = new Date();
  function iso(daysAgo) { return new Date(now.getTime() - daysAgo * 86400000).toISOString(); }
  var n = 0;
  function task(t, lane, bucket, label, done, extra) {
    n++;
    return {
      id: 1700000000000000 + n, t: t, lane: lane, bucket: bucket, label: label,
      notes: (extra && extra.notes) || '', done: !!done, doneAt: done ? iso(1) : null,
      created: iso(n), by: (extra && extra.by) || 'Rory', src: (extra && extra.src) || '',
      comments: (extra && extra.comments) || []
    };
  }

  var tasks = [
    task('Renew the commercial general liability insurance', 'Rory', 'now', 'Admin', false,
         { notes: 'Said top-top-top priority.' }),
    task('Review and finish sending the invoices', 'Rory', 'now', 'Finances'),
    task('Get the survey rod to Rory today', 'Cam', 'today', 'Sensors', false,
         { by: 'Rory (voice)', src: 'transcripts/2026-09-21_Cam Survey Rod Today.md' }),
    task('Install a game camera at Apex Upper', 'Cam', 'today', 'LLC'),
    task('Shift Anderson/Fell Discharge SDR', 'Allie', 'week', 'CoN', false,
         { comments: [{ id: 'c1', who: 'Allie', at: iso(1), text: 'Waiting on the new tool.' }] }),
    task('Update the KEI website to the new standard', 'Allie', 'week', 'BizDev', true),
    task('Buy Plane Tickets to PG', 'Anne', 'tosort', ''),
    task('Make a block for the drawer', 'Anne', 'tosort', 'House'),
    task('Put a gutter on the shed', 'Contractors', 'later', 'House'),
    task('Wash and clean the motorcycle gear', 'Rory', 'tosort', 'Motorcycle'),
    task('Organize the flooring contract', 'Rory', 'now', 'House'),
    task('Check out the new path in the campground', 'Rory', 'someday', 'Personal')
  ];
  var settings = {
    lanes: ['Rory', 'Anne', 'Cam', 'Allie', 'Claude', 'Contractors'],
    labels: [{ name: 'CoN', hue: 200 }, { name: 'LLC', hue: 150 }, { name: 'HPWPS', hue: 260 },
             { name: 'BizDev', hue: 320 }, { name: 'Finances', hue: 45 }, { name: 'HR', hue: 15 },
             { name: 'Sensors', hue: 180 }, { name: 'Admin', hue: 280 }, { name: 'Personal', hue: 95 },
             { name: 'House', hue: 25 }, { name: 'Motorcycle', hue: 50 }, { name: 'Yard', hue: 110 }],
    personalLabels: ['Personal', 'House', 'Motorcycle', 'Yard'],
    schemaVersion: 1
  };
  var activity = [{ id: 'a1', who: 'Allie', at: iso(0), text: 'completed "Update the KEI website"' }];
  var pending = [];   /* changes a fake "other person" has made */

  window.AUTH = {
    EXPIRED: 'signed-out',
    start: function () { return Promise.resolve({ signedIn: true, name: 'Demo User', email: 'rory@kootenayenvironmental.ca' }); },
    signIn: function () {},
    signOut: function () {},
    token: function () { return Promise.resolve('demo'); },
    who: function () { return { name: 'Demo User', email: 'rory@kootenayenvironmental.ca' }; }
  };

  window.STORE = {
    load: function () {
      return Promise.resolve({
        tasks: tasks.map(function (t) { return JSON.parse(JSON.stringify(t)); }),
        settings: settings, activity: activity, personalReadable: true, badComments: []
      });
    },
    put: function (t) {
      var at = tasks.findIndex(function (x) { return x.id === t.id; });
      if (at < 0) tasks.push(JSON.parse(JSON.stringify(t)));
      else tasks[at] = JSON.parse(JSON.stringify(t));
      return Promise.resolve(t);
    },
    remove: function (t) {
      tasks = tasks.filter(function (x) { return x.id !== t.id; });
      return Promise.resolve();
    },
    putSetting: function (k, v) { settings[k] = v; return Promise.resolve(); },
    addActivity: function (e) { activity.unshift(e); return Promise.resolve(); },
    poll: function () { var out = pending; pending = []; return Promise.resolve(out); },
    setPersonalLabels: function () {},
    listFor: function () { return 'KEI Tasks'; },
    canReadPersonal: function () { return true; }
  };

  /* Pretend someone else changed something, to check the board takes it in
     without disturbing what you are doing:  DEMO.otherPersonEdits()  */
  window.DEMO = {
    tasks: function () { return tasks; },
    otherPersonEdits: function (title) {
      var t = JSON.parse(JSON.stringify(tasks[0]));
      t.t = title || 'EDITED BY SOMEONE ELSE';
      tasks[0] = t;
      pending.push({ task: t });
      return t;
    },
    otherPersonDeletes: function () {
      var gone = tasks.pop();
      pending.push({ deletedId: gone.id });
      return gone;
    }
  };
})();
