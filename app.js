(function () {
  'use strict';

  var L = window.VendesLogic;
  var CFG = window.VENDES_WEB || {};
  var params = new URLSearchParams(location.search);
  var QUEUE_KEY = 'vendes.pwa.queue';
  var MODEL_KEY = 'vendes.pwa.model';
  var STAFF_KEY = 'vendes.pwa.staff';
  var TOKEN_KEY = 'vendes.pwa.token';

  var S = {
    demo: params.get('demo') === '1',
    screen: 'loading',
    stage: '',
    staff: null,
    model: null,
    grids: {},
    payload: null,
    dash: null,
    catalog: null,
    waybillNo: '',
    batch: [],
    banner: null,
    queue: [],
    resolves: [],
    online: params.get('offline') === '1' ? false : navigator.onLine,
    syncing: false,
    token: '',
    scanner: null,
    lastCode: '',
    lastAt: 0,
    labelMode: 'a4',
    labelOn: {},
    labelQty: {},
    tokenClient: null
  };

  function $(id) { return document.getElementById(id); }

  function esc(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function configured() {
    return CFG.clientId && CFG.clientId.indexOf('PASTE_') !== 0
      && CFG.spreadsheetId && CFG.spreadsheetId.indexOf('PASTE_') !== 0;
  }

  function pendingCount() {
    var scans = S.queue.filter(function (item) { return item.phase !== 'done'; }).length;
    return scans + S.resolves.length;
  }

  function loadQueue() {
    try {
      var saved = JSON.parse(localStorage.getItem(QUEUE_KEY) || '{}');
      S.queue = saved.queue || [];
      S.resolves = saved.resolves || [];
    } catch (err) {
      S.queue = [];
      S.resolves = [];
    }
  }

  function saveQueue() {
    try {
      localStorage.setItem(QUEUE_KEY, JSON.stringify({ queue: S.queue, resolves: S.resolves }));
    } catch (err) {}
  }

  function loadCache() {
    try {
      var model = JSON.parse(localStorage.getItem(MODEL_KEY) || 'null');
      var staff = JSON.parse(localStorage.getItem(STAFF_KEY) || 'null');
      if (model && staff) {
        S.model = model;
        S.staff = staff;
      }
    } catch (err) {}
  }

  function saveCache() {
    try {
      localStorage.setItem(MODEL_KEY, JSON.stringify(S.model));
      localStorage.setItem(STAFF_KEY, JSON.stringify(S.staff));
    } catch (err) {}
  }

  function readToken() {
    try {
      var saved = JSON.parse(sessionStorage.getItem(TOKEN_KEY) || 'null');
      if (!saved || !saved.token || saved.expires < Date.now() + 60000) return '';
      return saved.token;
    } catch (err) {
      return '';
    }
  }

  function storeToken(token, expiresIn) {
    S.token = token;
    sessionStorage.setItem(TOKEN_KEY, JSON.stringify({
      token: token,
      expires: Date.now() + (Number(expiresIn) || 3600) * 1000
    }));
  }

  function sheetHasLocalId(localId) {
    var needle = 'ID:' + localId;
    return (S.model && S.model.scans || []).some(function (scan) {
      var message = String(scan.message || '');
      return message === needle || message.indexOf(needle + ' ') === 0;
    });
  }

  function refreshViews() {
    if (!S.model) return;
    var stage = S.stage || 'PICK';
    S.payload = L.buildStagePayload(S.model, stage, null, new Date());
    S.queue.forEach(function (item) {
      if (!item || item.phase === 'done') return;
      var event = Object.assign({ localId: item.localId }, item.event || {});
      var no = L.normWaybill(event.waybill || event.code);
      var view = S.payload.waybills[no] || null;
      var decision = L.decide(view, event, S.payload.codes);
      if (decision.result === 'OK' && view) {
        S.payload.waybills[no] = L.applyDecisionToView(view, decision);
      }
    });
  }

  function enqueue(event) {
    event.localId = 'L' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    S.queue.push({ localId: event.localId, event: event, phase: 'new' });
    saveQueue();
    refreshViews();
    paint();
    syncNow();
  }

  function api(path, options) {
    var opts = options || {};
    opts.headers = opts.headers || {};
    opts.headers.Authorization = 'Bearer ' + S.token;
    return fetch('https://sheets.googleapis.com/v4/spreadsheets/' + CFG.spreadsheetId + path, opts).then(function (res) {
      if (res.status === 401) {
        S.token = '';
        sessionStorage.removeItem(TOKEN_KEY);
        throw new Error('Sign in again to sync.');
      }
      if (!res.ok) throw new Error('Sheet request failed (' + res.status + ')');
      return res.json();
    });
  }

  function loadSheet() {
    var tabs = ['STAFF', 'ITEMS', 'SKU_MAP', 'WAYBILLS', 'SCANS', 'FLAGS', 'LEDGER'];
    var ranges = tabs.map(function (key) {
      var spec = L.SHEETS[key];
      return spec.name + '!A:' + L.columnName(spec.headers.length);
    });
    var query = ranges.map(function (range) { return 'ranges=' + encodeURIComponent(range); }).join('&');
    return api('/values:batchGet?' + query).then(function (body) {
      var tables = {};
      tabs.forEach(function (key, index) {
        var spec = L.SHEETS[key];
        var values = (body.valueRanges[index] && body.valueRanges[index].values) || [spec.headers];
        S.grids[key] = values;
        tables[key] = L.rowsToRecords(values, spec.fields);
      });
      S.model = L.hydrateModel({
        staff: tables.STAFF,
        items: tables.ITEMS,
        skuMap: tables.SKU_MAP,
        waybills: tables.WAYBILLS,
        scans: tables.SCANS,
        flags: tables.FLAGS,
        ledger: tables.LEDGER
      });
      saveCache();
    });
  }

  function writeValues(range, rows) {
    return api('/values/' + encodeURIComponent(range) + '?valueInputOption=USER_ENTERED', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ values: rows })
    });
  }

  function appendRows(tab, records) {
    if (!records.length) return Promise.resolve();
    var spec = L.SHEETS[tab];
    var matrix = L.recordsToMatrix(records, spec.fields);
    var grid = S.grids[tab] || [spec.headers];
    var range = L.nextSheetRange(spec.name, spec.headers.length, grid, matrix.length);
    return writeValues(range, matrix).then(function () {
      matrix.forEach(function (row) { grid.push(row); });
      S.grids[tab] = grid;
    });
  }

  function syncNow() {
    if (S.demo || S.syncing || !S.online || !S.token || !S.model) return;
    if (!pendingCount()) return;
    S.syncing = true;
    paint();
    loadSheet().then(function () {
      S.queue = L.attachSyncRows(S.model, S.staff, S.queue, new Date(), null);
      var chain = Promise.resolve();
      S.queue.forEach(function (item) {
        chain = chain.then(function () {
          if (item.phase === 'done') return;
          var step = L.syncWrites(item, sheetHasLocalId(item.localId), S.model.flags);
          return appendRows('SCANS', step.scans).then(function () {
            if (step.scans.length) {
              item.phase = step.phase;
              S.model.scans = S.model.scans.concat(step.scans);
              saveQueue();
            }
            var flags = L.syncWrites(item, true, S.model.flags);
            return appendRows('FLAGS', flags.flags).then(function () {
              if (flags.flags.length) S.model.flags = S.model.flags.concat(flags.flags);
              item.phase = 'done';
              saveQueue();
            });
          });
        });
      });
      return chain.then(function () {
        var resolves = S.resolves.slice();
        var next = Promise.resolve();
        resolves.forEach(function (patch) {
          next = next.then(function () { return writeValues(patch.range, [[true]]); });
        });
        return next.then(function () { S.resolves = []; });
      });
    }).then(function () {
      S.queue = S.queue.filter(function (item) { return item.phase !== 'done'; });
      saveQueue();
      saveCache();
      refreshViews();
      if (S.screen === 'admin') S.dash = L.buildAdminDashboard(S.model, new Date(), null);
    }).catch(function () {
      S.banner = { kind: 'warn', text: 'Sync paused. Scans stay on this phone.' };
    }).then(function () {
      S.syncing = false;
      paint();
    });
  }

  function afterToken() {
    return fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: 'Bearer ' + S.token }
    }).then(function (res) {
      if (!res.ok) throw new Error('profile');
      return res.json();
    }).then(function (profile) {
      return loadSheet().then(function () {
        var staff = L.findStaff(S.model.staff, profile.email);
        if (!staff || !staff.active) {
          S.staff = null;
          S.screen = 'denied';
          S.banner = null;
          render();
          return;
        }
        S.staff = staff;
        saveCache();
        if (S.screen === 'loading' || S.screen === 'signin' || S.screen === 'setup') S.screen = 'home';
        refreshViews();
        render();
        syncNow();
      });
    }).catch(function () {
      if (!S.model) {
        S.screen = 'signin';
        S.banner = { kind: 'bad', text: 'Could not reach Google. Try again when you are online.' };
        render();
      }
    });
  }

  function signIn() {
    if (!S.tokenClient) {
      S.banner = { kind: 'bad', text: 'Google sign-in is still loading. Check the connection and try again.' };
      paint();
      return;
    }
    var prompt = readToken() ? '' : 'consent';
    S.tokenClient.requestAccessToken({ prompt: prompt });
  }

  function initGis() {
    var redirect = /iPad|iPhone|iPod/.test(navigator.userAgent) && (navigator.standalone || window.matchMedia('(display-mode: standalone)').matches);
    S.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CFG.clientId,
      scope: 'openid email https://www.googleapis.com/auth/spreadsheets',
      ux_mode: redirect ? 'redirect' : 'popup',
      callback: function (resp) {
        if (!resp || resp.error) {
          S.banner = { kind: 'bad', text: 'Google sign-in was cancelled.' };
          S.screen = 'signin';
          render();
          return;
        }
        storeToken(resp.access_token, resp.expires_in);
        afterToken();
      }
    });
  }

  function headerHtml(tag) {
    var back = '<span class="back-spacer"></span>';
    if (S.screen !== 'home' && S.screen !== 'denied' && S.screen !== 'signin' && S.screen !== 'setup' && S.screen !== 'loading') {
      back = '<button type="button" class="back" data-go="home">Back</button>';
    }
    var label = L.pendingLabel(pendingCount(), S.online, S.syncing);
    var state = S.syncing || pendingCount() ? 'saving' : (S.online ? 'saved' : 'saving');
    return '<header>' + back +
      '<div><div class="brand">VENDES</div><div class="stage-tag">' + esc(tag || 'Warehouse') + '</div></div>' +
      '<div id="save-pill" class="save ' + state + '">' + esc(label) + '</div></header>' +
      (S.online ? '' : '<div class="offline-bar">Offline. Scans stay on this phone and sync when you reconnect.</div>');
  }

  function bannerHtml() {
    if (!S.banner) return '';
    var bang = S.banner.kind === 'bad' ? '<span class="bang">!</span>' : '';
    return '<div class="banner ' + esc(S.banner.kind) + '" role="alert">' + bang + '<span>' + esc(S.banner.text) + '</span></div>';
  }

  function installHtml() {
    var standalone = window.matchMedia('(display-mode: standalone)').matches || navigator.standalone;
    if (standalone || S.demo) return '';
    var ios = /iPad|iPhone|iPod/.test(navigator.userAgent);
    var text = ios
      ? 'In Safari, tap Share, then Add to Home Screen.'
      : 'Install Vendes so the icon stays on the home screen.';
    var button = ios ? '' : '<button type="button" class="wide" id="install-btn" hidden>Install</button>';
    return '<section class="install"><b>Add Vendes to the home screen</b><p class="hint">' + esc(text) + '</p>' + button + '</section>';
  }

  function homeHtml() {
    var admin = S.staff && L.roleCanUseAdmin(S.staff)
      ? '<button type="button" class="ghost" data-go="admin">Admin</button>'
      : '';
    return headerHtml('Warehouse') + bannerHtml() + '<main>' + installHtml() +
      '<p class="hi">Hi, ' + esc(S.staff ? S.staff.name : '') +
      (S.staff && S.staff.role ? ' · ' + esc(S.staff.role) : '') + '</p>' +
      '<p class="sub">Choose a step</p>' +
      '<button type="button" class="big pick" data-go="pick"><b>PICK</b><span>Collect the items</span></button>' +
      '<button type="button" class="big pack" data-go="pack"><b>PACK</b><span>Put the items in the parcel</span></button>' +
      '<button type="button" class="big dispatch" data-go="dispatch"><b>DISPATCH</b><span>Hand the parcel to the rider</span></button>' +
      admin +
      '<p class="note">Scan only. Stock drops when the J&amp;T rider picks the parcel up.</p></main>';
  }

  function currentView() {
    if (!S.payload || !S.waybillNo) return null;
    return S.payload.waybills[S.waybillNo] || null;
  }

  function haveOf(line) {
    return S.stage === 'PACK' ? line.packed : line.picked;
  }

  function stageHtml() {
    var waiting = !S.payload;
    var hint = '';
    if (waiting) hint = '<p class="hint">' + esc(L.MSG.LOADING) + '</p>';
    else if (S.stage === 'DISPATCH') hint = '<p class="hint">Scan the waybill. Rider handover. It must be fully packed.</p>';
    else if (!currentView()) hint = '<p class="hint">Scan the waybill. The camera and a Bluetooth scanner both work.</p>';
    else {
      var view = currentView();
      hint = '<div class="card"><div class="wb">' + esc(view.waybillNo) + '</div>' +
        '<span class="pill" data-client="' + esc(view.client) + '">' + esc(view.clientLabel) + '</span></div>';
    }
    var lines = '';
    var viewNow = currentView();
    if (viewNow && S.stage !== 'DISPATCH') {
      lines = viewNow.lines.map(function (line) {
        var have = haveOf(line);
        var pct = line.expected ? Math.max(0, Math.min(100, Math.round(have / line.expected * 100))) : 0;
        return '<div class="' + (have >= line.expected ? 'line full' : 'line') + '">' +
          '<div class="line-top"><b>' + esc(line.sku) + '</b><span class="count">' + have + ' / ' + line.expected + '</span></div>' +
          '<div class="pname">' + esc(line.name) + '</div>' +
          '<div class="bar"><i style="width:' + pct + '%"></i></div></div>';
      }).join('');
    }
    var batch = '';
    if (S.stage === 'DISPATCH') {
      batch = S.batch.length ? S.batch.map(function (row) {
        return '<div class="batch-row"><div><b class="wb">' + esc(row.waybillNo) + '</b><div class="chips">' +
          esc(row.count) + (row.count === 1 ? ' item' : ' items') + '</div></div>' +
          '<span class="pill" data-client="' + esc(row.client) + '">' + esc(row.clientLabel) + '</span></div>';
      }).join('') : '<p class="hint">Nothing in the batch yet.</p>';
    }
    var doneLabel = S.stage === 'DISPATCH' ? 'Handover done' : 'Done — next waybill';
    var doneOff = S.stage === 'DISPATCH' ? !S.batch.length : !(viewNow && (S.stage === 'PACK' ? viewNow.packComplete : viewNow.pickComplete));
    var short = S.stage === 'DISPATCH' || !viewNow ? '' : '<button type="button" class="wide warn" data-action="short">Submit with shortage</button>';
    return headerHtml(S.stage) + '<div id="banner-slot">' + bannerHtml() + '</div><main>' +
      '<div id="live">' + hint + lines + batch + '</div>' +
      '<button type="button" class="wide cam" id="cam-btn" data-action="camera">' + (S.scanner ? 'Close camera' : 'Open camera') + '</button>' +
      '<div id="reader" hidden></div>' +
      '<form id="scan-form"><input id="code-input" maxlength="64" autocomplete="off" autocapitalize="characters" enterkeyhint="go"' +
      (waiting ? ' disabled' : '') + ' placeholder="' + esc(waiting ? L.MSG.LOADING : (S.stage === 'DISPATCH' || !S.waybillNo ? 'Type the waybill' : 'Type the SKU or barcode')) + '">' +
      '<button type="submit" id="add-btn"' + (waiting ? ' disabled' : '') + '>Add</button></form>' +
      short +
      '<button type="button" class="wide good" data-action="done"' + (doneOff ? ' disabled' : '') + '>' + doneLabel + '</button></main>';
  }

  function adminHtml() {
    if (!S.dash) return headerHtml('Admin') + '<main><p class="hint">Loading dashboard…</p></main>';
    var counts = S.dash.counts || {};
    var tiles = [['PICK', 'Pick'], ['PACK', 'Pack'], ['DISPATCH', 'Dispatch'], ['PICKUP', 'Pickup']].map(function (pair) {
      return '<div class="tile"><b>' + (counts[pair[0]] || 0) + '</b><span>' + pair[1] + '</span></div>';
    }).join('');
    var buckets = [
      ['printedNotPicked', 'Printed, not picked yet'],
      ['pickedNotPacked', 'Picked, not packed yet'],
      ['packedNotDispatched', 'Packed, not dispatched yet'],
      ['dispatchedNotPickedUp', 'Dispatched, no J&T pickup']
    ];
    var backlog = buckets.map(function (pair) {
      var list = (S.dash.backlog && S.dash.backlog[pair[0]]) || [];
      return '<div class="back-row"><div><div>' + pair[1] + '</div><div class="chips">' +
        esc(list.slice(0, 3).map(function (item) { return L.clipDisplay(item); }).join('  ')) +
        '</div></div><b>' + list.length + '</b></div>';
    }).join('');
    var flags = (S.dash.flags || []).map(function (flag) {
      var text = L.adminCardText(flag);
      return '<article class="flag"><div class="flag-top"><span class="flag-type">' + esc(text.type) + '</span>' +
        '<button type="button" data-resolve="1" data-row="' + esc(flag._row || '') + '" data-ts="' + esc(flag.timestamp) +
        '" data-wb="' + esc(flag.waybillKey || flag.waybill) + '" data-type="' + esc(flag.type) + '">Resolved</button></div>' +
        '<b>' + esc(text.waybill) + '</b> · ' + esc(text.client) +
        '<p>' + esc(text.details) + '</p></article>';
    }).join('') || '<p class="hint">No open flags.</p>';
    return headerHtml('Admin') + '<main>' +
      '<h2>Today · ' + esc(L.clipDisplay(S.dash.today)) + '</h2><div class="tiles">' + tiles + '</div>' +
      '<h2>Open flags</h2>' + flags +
      '<h2>Backlog</h2>' + backlog +
      '<p class="note">J&amp;T import still runs from the Google Sheet menu. Stock changes only there.</p>' +
      '<button type="button" class="wide" data-go="labels">Print SKU labels</button></main>';
  }

  function labelsHtml() {
    var rows = (S.catalog || []).map(function (item) {
      var qty = S.labelQty[item.sku] || 1;
      var checked = S.labelOn[item.sku] ? ' checked' : '';
      return '<div class="sku-row"><label><input type="checkbox" data-sku="' + esc(item.sku) + '"' + checked + '> <b>' +
        esc(item.sku) + '</b></label><span class="chips">' + esc(item.name) +
        '</span><input type="number" min="1" max="100" value="' + qty + '" data-qty="' + esc(item.sku) + '"></div>';
    }).join('') || '<p class="hint">No SKUs.</p>';
    return headerHtml('Labels') + '<main><div class="modes">' +
      '<button type="button" class="' + (S.labelMode === 'a4' ? 'on' : '') + '" data-action="mode-a4">A4 sheet</button>' +
      '<button type="button" class="' + (S.labelMode === 'thermal' ? 'on' : '') + '" data-action="mode-thermal">50×30 printer</button></div>' +
      rows +
      '<button type="button" class="wide good" data-action="print-labels">Print</button></main>';
  }

  function setupHtml() {
    return '<div class="denied"><div class="brand">VENDES</div><h1>Setup needed</h1>' +
      '<p>This phone app is ready, but it still needs a Google sign-in client and the sheet id. The steps are in SETUP.md under “Installable phone app”.</p>' +
      installHtml() + '</div>';
  }

  function signinHtml() {
    return headerHtml('Sign in') + bannerHtml() + '<main><section class="signin-card"><h1>Vendes Scan</h1>' +
      '<p>Sign in with the Google account on the STAFF list.</p>' +
      '<button type="button" class="wide good" data-action="signin">Sign in with Google</button></section>' +
      installHtml() + '</main>';
  }

  function deniedHtml() {
    return '<div class="denied"><div class="brand">VENDES</div><h1>Not authorized, ask admin</h1>' +
      '<p>This Google account is not an active row on the STAFF tab.</p></div>';
  }

  function render() {
    var app = $('app');
    if (!app) return;
    if (S.screen === 'loading') app.innerHTML = '<div class="loading"><div class="brand">VENDES</div><p>One moment…</p></div>';
    else if (S.screen === 'setup') app.innerHTML = setupHtml();
    else if (S.screen === 'signin') app.innerHTML = signinHtml();
    else if (S.screen === 'denied') app.innerHTML = deniedHtml();
    else if (S.screen === 'home') app.innerHTML = homeHtml();
    else if (S.screen === 'admin') app.innerHTML = adminHtml();
    else if (S.screen === 'labels') app.innerHTML = labelsHtml();
    else app.innerHTML = stageHtml();
    var input = $('code-input');
    if (input && !S.scanner) input.focus();
    wireInstall();
  }

  function paint() {
    if (S.screen === 'stage' && $('live')) {
      var banner = $('banner-slot');
      if (banner) banner.innerHTML = bannerHtml();
      var live = $('live');
      var view = currentView();
      var hint = '';
      if (!S.payload) hint = '<p class="hint">' + esc(L.MSG.LOADING) + '</p>';
      else if (S.stage === 'DISPATCH') hint = '<p class="hint">Scan the waybill. Rider handover. It must be fully packed.</p>';
      else if (!view) hint = '<p class="hint">Scan the waybill. The camera and a Bluetooth scanner both work.</p>';
      else hint = '<div class="card"><div class="wb">' + esc(view.waybillNo) + '</div><span class="pill" data-client="' + esc(view.client) + '">' + esc(view.clientLabel) + '</span></div>';
      var lines = '';
      if (view && S.stage !== 'DISPATCH') {
        lines = view.lines.map(function (line) {
          var have = haveOf(line);
          var pct = line.expected ? Math.max(0, Math.min(100, Math.round(have / line.expected * 100))) : 0;
          return '<div class="' + (have >= line.expected ? 'line full' : 'line') + '"><div class="line-top"><b>' + esc(line.sku) + '</b><span class="count">' + have + ' / ' + line.expected + '</span></div><div class="pname">' + esc(line.name) + '</div><div class="bar"><i style="width:' + pct + '%"></i></div></div>';
        }).join('');
      }
      var batch = '';
      if (S.stage === 'DISPATCH') {
        batch = S.batch.length ? S.batch.map(function (row) {
          return '<div class="batch-row"><div><b class="wb">' + esc(row.waybillNo) + '</b></div><span class="pill">' + esc(row.clientLabel) + '</span></div>';
        }).join('') : '<p class="hint">Nothing in the batch yet.</p>';
      }
      live.innerHTML = hint + lines + batch;
      var pill = $('save-pill');
      if (pill) pill.textContent = L.pendingLabel(pendingCount(), S.online, S.syncing);
      var input = $('code-input');
      if (input && !S.scanner) {
        input.placeholder = !S.payload ? L.MSG.LOADING : (S.stage === 'DISPATCH' || !S.waybillNo ? 'Type the waybill' : 'Type the SKU or barcode');
      }
      var done = document.querySelector('[data-action="done"]');
      if (done) {
        var doneOff = S.stage === 'DISPATCH' ? !S.batch.length : !(view && (S.stage === 'PACK' ? view.packComplete : view.pickComplete));
        done.disabled = doneOff;
      }
      return;
    }
    render();
  }

  var deferredInstall = null;
  function wireInstall() {
    var button = $('install-btn');
    if (!button) return;
    button.hidden = !deferredInstall;
    button.onclick = function () {
      if (!deferredInstall) return;
      deferredInstall.prompt();
      deferredInstall = null;
      button.hidden = true;
    };
  }

  function go(target) {
    stopCamera();
    S.banner = null;
    if (target === 'home') {
      S.screen = 'home';
      S.waybillNo = '';
      render();
      return;
    }
    if (target === 'admin') {
      if (!L.roleCanUseAdmin(S.staff)) {
        S.screen = 'denied';
        render();
        return;
      }
      S.dash = S.model ? L.buildAdminDashboard(S.model, new Date(), null) : null;
      S.screen = 'admin';
      render();
      return;
    }
    if (target === 'labels') {
      if (!L.roleCanUseAdmin(S.staff)) {
        S.screen = 'denied';
        render();
        return;
      }
      S.catalog = L.labelCatalog(S.model || { items: [] });
      S.screen = 'labels';
      render();
      return;
    }
    S.stage = String(target || '').toUpperCase();
    S.screen = 'stage';
    S.waybillNo = '';
    S.batch = [];
    refreshViews();
    render();
  }

  function onCode(raw) {
    var code = String(raw || '').trim();
    if (!code) {
      S.banner = { kind: 'bad', text: L.MSG.EMPTY };
      paint();
      return;
    }
    var gate = L.stageEntryWhileLoading(!!S.payload);
    if (gate.blocked) {
      S.banner = { kind: 'warn', text: gate.message };
      paint();
      return;
    }
    if (code.length > 64) {
      S.banner = { kind: 'bad', text: L.MSG.TOO_LONG };
      paint();
      return;
    }
    if (S.stage === 'DISPATCH' || !S.waybillNo) openWaybill(code);
    else scanItem(code);
  }

  function openWaybill(code) {
    var no = L.normWaybill(code);
    var view = S.payload.waybills[no];
    if (S.stage === 'DISPATCH') {
      dispatchOne(no, view);
      return;
    }
    if (!view) {
      S.banner = { kind: 'bad', text: L.MSG.UNKNOWN_WB };
      beep(false);
      enqueue({ action: 'ITEM', stage: S.stage, waybill: no, code: no });
      return;
    }
    if (view.pickedUp) {
      S.banner = { kind: 'bad', text: L.MSG.PICKED_UP };
      beep(false);
      paint();
      return;
    }
    S.waybillNo = no;
    if (S.stage === 'PACK' && !view.pickComplete) {
      S.banner = { kind: 'bad', text: L.MSG.NOT_PICKED };
      enqueue({ action: 'OPEN', stage: 'PACK', waybill: no, code: '' });
      return;
    }
    S.banner = null;
    refreshViews();
    paint();
  }

  function dispatchOne(no, view) {
    if (S.batch.some(function (row) { return row.waybillNo === no; })) {
      S.banner = { kind: 'warn', text: 'Already in the batch' };
      beep(false);
      paint();
      return;
    }
    var decision = L.decide(view || null, { action: 'DISPATCH', stage: 'DISPATCH', waybill: no, code: no }, S.payload.codes);
    if (decision.result !== 'OK') {
      S.banner = { kind: 'bad', text: decision.message };
      beep(false);
      enqueue({ action: 'DISPATCH', stage: 'DISPATCH', waybill: no, code: no });
      return;
    }
    S.batch.push({
      waybillNo: no,
      client: view.client,
      clientLabel: view.clientLabel,
      count: (view.lines || []).length
    });
    S.banner = { kind: 'good', text: decision.message };
    beep(true);
    enqueue({ action: 'DISPATCH', stage: 'DISPATCH', waybill: no, code: '' });
  }

  function scanItem(code) {
    var view = currentView();
    var event = { action: 'ITEM', stage: S.stage, waybill: S.waybillNo, code: code };
    var decision = L.decide(view, event, S.payload.codes);
    if (decision.result === 'OK') {
      S.banner = { kind: 'good', text: decision.sku + ' added' };
      beep(true);
    } else {
      S.banner = { kind: 'bad', text: decision.message };
      beep(false);
    }
    enqueue(event);
  }

  function submitShort() {
    var view = currentView();
    if (!view) return;
    var event = { action: 'SHORTAGE', stage: S.stage, waybill: S.waybillNo, code: '' };
    var decision = L.decide(view, event, S.payload.codes);
    if (decision.result !== 'OK') {
      S.banner = { kind: 'warn', text: decision.message };
      paint();
      return;
    }
    if (!window.confirm('Some pieces are missing. Submit with a shortage and save a flag?')) return;
    S.banner = { kind: 'warn', text: decision.message };
    beep(false);
    enqueue(event);
  }

  function nextStep() {
    if (S.stage === 'DISPATCH') {
      var count = S.batch.length;
      S.batch = [];
      S.screen = 'home';
      S.banner = { kind: 'good', text: 'Handover done: ' + count + ' waybills' };
      render();
      return;
    }
    var view = currentView();
    if (!view) return;
    var complete = S.stage === 'PACK' ? view.packComplete : view.pickComplete;
    if (!complete) return;
    S.waybillNo = '';
    S.banner = { kind: 'good', text: 'Done. Scan the next waybill.' };
    refreshViews();
    paint();
  }

  function resolveFlag(button) {
    var row = Number(button.getAttribute('data-row'));
    var key = {
      timestamp: button.getAttribute('data-ts'),
      waybill: button.getAttribute('data-wb'),
      type: button.getAttribute('data-type')
    };
    S.dash.flags = (S.dash.flags || []).filter(function (flag) {
      return !(flag.timestamp === key.timestamp && (flag.waybillKey || flag.waybill) === key.waybill && flag.type === key.type);
    });
    if (row) S.resolves.push({ range: 'FLAGS!G' + row, value: true });
    saveQueue();
    render();
    syncNow();
  }

  function printLabels() {
    var items = [];
    (S.catalog || []).forEach(function (item) {
      if (!S.labelOn[item.sku]) return;
      items.push({
        sku: item.sku,
        name: item.name,
        clientLabel: item.clientLabel,
        copies: S.labelQty[item.sku] || 1
      });
    });
    if (!items.length) {
      S.banner = { kind: 'warn', text: 'Choose a SKU first.' };
      paint();
      return;
    }
    var html = L.labelPrintDocument(items, S.labelMode);
    var popup = window.open('', '_blank');
    if (!popup) {
      S.banner = { kind: 'bad', text: 'Allow pop-ups to print labels.' };
      paint();
      return;
    }
    popup.document.open();
    popup.document.write(html);
    popup.document.close();
  }

  function beep(ok) {
    try {
      if (navigator.vibrate) navigator.vibrate(ok ? 25 : [70, 40, 70]);
    } catch (err) {}
  }

  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      if (window.Html5Qrcode) return resolve();
      var script = document.createElement('script');
      script.src = src;
      script.onload = function () { resolve(); };
      script.onerror = function () { reject(new Error('script')); };
      document.head.appendChild(script);
    });
  }

  function stopCamera() {
    var scanner = S.scanner;
    S.scanner = null;
    if (!scanner) return;
    scanner.stop().then(function () { scanner.clear(); }).catch(function () {});
  }

  function startCamera() {
    if (S.scanner) {
      stopCamera();
      paint();
      return;
    }
    var reader = $('reader');
    if (!reader) return;
    reader.hidden = false;
    loadScript('vendor/html5-qrcode.min.js').then(function () {
      var formats = [
        Html5QrcodeSupportedFormats.QR_CODE,
        Html5QrcodeSupportedFormats.CODE_128,
        Html5QrcodeSupportedFormats.CODE_39,
        Html5QrcodeSupportedFormats.EAN_13,
        Html5QrcodeSupportedFormats.EAN_8
      ];
      S.scanner = new Html5Qrcode('reader');
      return S.scanner.start(
        { facingMode: 'environment' },
        { fps: 8, qrbox: { width: 240, height: 140 }, formatsToSupport: formats },
        function (text) {
          var now = Date.now();
          if (text === S.lastCode && now - S.lastAt < 1600) return;
          S.lastCode = text;
          S.lastAt = now;
          onCode(text);
        },
        function () {}
      );
    }).then(function () {
      var button = $('cam-btn');
      if (button) button.textContent = 'Close camera';
    }).catch(function () {
      S.banner = { kind: 'bad', text: 'Camera is blocked. Type the code instead.' };
      if (reader) reader.hidden = true;
      S.scanner = null;
      paint();
    });
  }

  function startDemo() {
    var model = L.buildSeedModel(new Date());
    S.model = model;
    S.staff = L.findStaff(model.staff, 'admin@vendes.example');
    var count = parseInt(params.get('queue') || '0', 10) || 0;
    if (params.has('queue')) S.queue = [];
    for (var i = 0; i < count; i++) {
      S.queue.push({
        localId: 'demo-' + i,
        phase: 'new',
        event: { action: 'ITEM', stage: 'PICK', waybill: 'WB-SIR-1001', code: 'WVT-LW' }
      });
    }
    S.screen = 'home';
    refreshViews();
    render();
  }

  document.body.addEventListener('click', function (event) {
    var el = event.target.closest('[data-go], [data-action], [data-resolve]');
    if (!el) return;
    if (el.getAttribute('data-go')) go(el.getAttribute('data-go'));
    else if (el.getAttribute('data-action') === 'signin') signIn();
    else if (el.getAttribute('data-action') === 'camera') startCamera();
    else if (el.getAttribute('data-action') === 'short') submitShort();
    else if (el.getAttribute('data-action') === 'done') nextStep();
    else if (el.getAttribute('data-action') === 'print-labels') printLabels();
    else if (el.getAttribute('data-action') === 'mode-a4' || el.getAttribute('data-action') === 'mode-thermal') {
      S.labelMode = el.getAttribute('data-action') === 'mode-a4' ? 'a4' : 'thermal';
      render();
    } else if (el.hasAttribute('data-resolve')) resolveFlag(el);
  });

  document.body.addEventListener('submit', function (event) {
    if (!event.target || event.target.id !== 'scan-form') return;
    event.preventDefault();
    var input = $('code-input');
    var gate = L.stageEntryWhileLoading(!!S.payload);
    if (gate.blocked) {
      S.banner = { kind: 'warn', text: gate.message };
      paint();
      return;
    }
    var value = input.value;
    input.value = '';
    onCode(value);
  });

  document.body.addEventListener('change', function (event) {
    var sku = event.target.getAttribute && event.target.getAttribute('data-sku');
    var qty = event.target.getAttribute && event.target.getAttribute('data-qty');
    if (sku) S.labelOn[sku] = !!event.target.checked;
    if (qty) S.labelQty[qty] = Math.max(1, Math.min(100, parseInt(event.target.value, 10) || 1));
  });

  window.addEventListener('beforeinstallprompt', function (event) {
    event.preventDefault();
    deferredInstall = event;
    wireInstall();
  });

  window.addEventListener('online', function () {
    if (params.get('offline') === '1') return;
    S.online = true;
    paint();
    syncNow();
  });

  window.addEventListener('offline', function () {
    S.online = false;
    paint();
  });

  function boot() {
    if (!L) {
      document.body.textContent = 'Scan logic did not load.';
      return;
    }
    loadQueue();
    if (S.demo) {
      startDemo();
      return;
    }
    loadCache();
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch(function () {});
    }
    if (!configured()) {
      S.screen = 'setup';
      render();
      return;
    }
    if (S.staff && S.model) {
      S.screen = 'home';
      refreshViews();
      render();
    } else {
      S.screen = 'signin';
      render();
    }
    if (!S.online) return;
    var script = document.createElement('script');
    script.src = 'https://accounts.google.com/gsi/client';
    script.onload = function () {
      initGis();
      var token = readToken();
      if (token) {
        S.token = token;
        afterToken();
      }
    };
    script.onerror = function () {};
    document.head.appendChild(script);
  }

  boot();
})();
