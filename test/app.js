(function () {
  'use strict';

  var L = window.VendesLogic;
  var CFG = window.VENDES_WEB || {};
  var Keys = window.VendesPwaKeys;
  var params = new URLSearchParams(location.search);
  var QUEUE_KEY = Keys.storageKey(CFG.spreadsheetId, 'queue');
  var MODEL_KEY = Keys.storageKey(CFG.spreadsheetId, 'model');
  var STAFF_KEY = Keys.storageKey(CFG.spreadsheetId, 'staff');
  var TOKEN_KEY = Keys.storageKey(CFG.spreadsheetId, 'token');

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
      var decision = L.decide(view, event, S.payload.codes, decisionContext());
      if (decision.result === 'OK' && view) {
        S.payload.waybills[no] = L.applyDecisionToView(view, decision);
      }
    });
  }

  function acceptedScans() {
    var extra = [];
    S.queue.forEach(function (item) {
      if (!item.accepted || !item.event) return;
      extra.push({
        waybill: item.event.waybill,
        stage: item.event.stage,
        result: 'OK',
        staffEmail: S.staff && S.staff.email,
        staffName: S.staff && S.staff.name
      });
    });
    return ((S.model && S.model.scans) || []).concat(extra);
  }

  function decisionContext() {
    var settings = S.model && S.model.settings;
    var waybillNos = {};
    ((S.model && S.model.waybills) || []).forEach(function (row) {
      var no = L.normWaybill(row.waybillNo);
      if (no) waybillNos[no] = true;
    });
    if (S.payload && S.payload.waybills) {
      Object.keys(S.payload.waybills).forEach(function (no) { waybillNos[no] = true; });
    }
    return {
      staff: S.staff,
      scans: acceptedScans(),
      waybillNos: waybillNos,
      samePersonStages: !settings || settings.samePersonStages !== false
    };
  }

  function enqueue(event, accepted) {
    event.localId = 'L' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
    event.scannedAt = new Date().toISOString();
    S.queue.push({ localId: event.localId, event: event, phase: 'new', accepted: !!accepted });
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
        ledger: tables.LEDGER,
        settings: []
      });
      return api('/values/' + encodeURIComponent('SETTINGS!A:B')).then(function (sheet) {
        tables.SETTINGS = L.rowsToRecords(sheet.values || [], L.SHEETS.SETTINGS.fields);
      }).catch(function () {}).then(function () {
        return api('/values/' + encodeURIComponent('TO_MAP!A:J')).then(function (sheet) {
          tables.TO_MAP = L.rowsToRecords(sheet.values || [], L.SHEETS.TO_MAP.fields, L.SHEETS.TO_MAP.headers);
        }).catch(function () { tables.TO_MAP = []; });
      }).then(function () {
        S.model = L.hydrateModel({
          staff: tables.STAFF,
          items: tables.ITEMS,
          skuMap: tables.SKU_MAP,
          waybills: tables.WAYBILLS,
          scans: tables.SCANS,
          flags: tables.FLAGS,
          ledger: tables.LEDGER,
          settings: tables.SETTINGS || [],
          toMap: tables.TO_MAP || []
        });
      }).then(function () {
        saveCache();
      });
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
      setBanner({ kind: 'warn', text: 'Sync paused. Scans stay on this phone.' }, false);
    }).then(function () {
      S.syncing = false;
      finishCheck();
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

  function sheetBadgeHtml() {
    var label = String(CFG.sheetLabel || '').trim().toUpperCase();
    if (!label) return '';
    return '<span class="sheet-badge' + (label === 'TEST' ? ' test' : '') + '">' + esc(label) + '</span>';
  }

  function headerHtml(tag) {
    var back = '<span class="back-spacer"></span>';
    if (S.screen !== 'home' && S.screen !== 'denied' && S.screen !== 'signin' && S.screen !== 'setup' && S.screen !== 'loading') {
      back = '<button type="button" class="back" data-go="home">Back</button>';
    }
    var label = L.pendingLabel(pendingCount(), S.online, S.syncing);
    var state = S.syncing || pendingCount() ? 'saving' : (S.online ? 'saved' : 'saving');
    return '<header>' + back +
      '<div class="brand-lock"><img class="brand-logo" src="brand/vendes-logo.png" alt="VENDES"><div class="stage-tag">' + esc(tag || 'Warehouse') + sheetBadgeHtml() + '</div></div>' +
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

  function stageButton(goName, stage, text) {
    if (!L.stageAssigned(S.staff, stage)) return '';
    var badge = stage === 'PACK' ? 'K' : stage.slice(0, 1);
    return '<button type="button" class="big ' + goName + '" data-go="' + goName + '"><span class="mark">' +
      badge + '</span><span class="stage-copy"><b>' + stage + '</b><span>' + text + '</span></span></button>';
  }

  function homeHtml() {
    var admin = S.staff && L.roleCanUseAdmin(S.staff)
      ? '<button type="button" class="ghost" data-go="admin">Admin</button>'
      : '';
    return headerHtml('Warehouse') + bannerHtml() + '<main>' + installHtml() +
      '<p class="hi">Hi, ' + esc(S.staff ? S.staff.name : '') +
      (S.staff && S.staff.role ? ' · ' + esc(S.staff.role) : '') + '</p>' +
      '<p class="sub">Choose a step</p>' +
      stageButton('pick', 'PICK', 'Collect the items') +
      stageButton('pack', 'PACK', 'Put the items in the parcel') +
      stageButton('dispatch', 'DISPATCH', 'Hand the parcel to the rider') +
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
    var hint = '';
    if (!S.payload) hint = '<p class="hint">' + esc(L.MSG.LOADING) + '</p>';
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
    var short = L.shortageAllowed(viewNow, S.stage) ? '<button type="button" class="wide warn" id="short-btn" data-action="short">Submit with shortage</button>' : '';
    return headerHtml(S.stage) + '<div id="banner-slot">' + bannerHtml() + '</div><main>' +
      '<div id="live">' + hint + lines + batch + '</div>' +
      '<button type="button" class="wide cam" id="cam-btn" data-action="camera">' + (S.scanner ? 'Close camera' : 'Open camera') + '</button>' +
      '<div id="reader" hidden></div>' +
      '<form id="scan-form"><input id="code-input" maxlength="64" autocomplete="off" autocapitalize="characters" enterkeyhint="go"' +
      ' placeholder="' + esc(S.checking ? 'Checking…' : (S.stage === 'DISPATCH' || !S.waybillNo ? 'Type the waybill' : 'Type the SKU or barcode')) + '">' +
      '<button type="submit" id="add-btn"' + (S.checking ? ' disabled' : '') + '>Add</button></form>' +
      short +
      '<button type="button" class="wide good" data-action="done"' + (doneOff ? ' disabled' : '') + '>' + doneLabel + '</button></main>';
  }

  function adminHtml() {
    if (!S.dash) return headerHtml('Admin') + '<main><p class="hint">Loading dashboard…</p></main>';
    var counts = S.dash.counts || {};
    var tiles = [['PICK', 'Pick'], ['PACK', 'Pack'], ['DISPATCH', 'Dispatch'], ['PICKUP', 'Pickup']].map(function (pair) {
      var short = counts[pair[0] + '_SHORT'] || 0;
      var extra = short ? '<small>+ ' + short + ' short</small>' : '';
      return '<div class="tile" title="Waybills that finished ' + pair[1] + ' today"><b>' + (counts[pair[0]] || 0) + '</b><span>' + pair[1] + '</span>' + extra + '</div>';
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
    var flagRows = (S.dash.flags || []).slice();
    if (S.showResolved) flagRows = flagRows.concat(S.dash.resolvedFlags || []);
    var flags = flagRows.map(function (flag) {
      var text = L.adminCardText(flag);
      return '<article class="flag"><div class="flag-top"><span class="flag-type">' + esc(text.type) + '</span>' +
        '<button type="button" data-resolve="1" data-row="' + esc(flag._row || '') + '" data-ts="' + esc(flag.timestamp) +
        '" data-wb="' + esc(flag.waybillKey || flag.waybill) + '" data-type="' + esc(flag.type) + '">' + (flag.resolved ? 'Resolved' : 'Mark resolved') + '</button></div>' +
        '<b>' + esc(text.waybill) + '</b> · ' + esc(text.client) +
        '<p>' + esc(text.details) + '</p></article>';
    }).join('') || '<p class="hint">No open flags.</p>';
    return headerHtml('Admin') + '<main>' +
      '<h2>Today · ' + esc(L.clipDisplay(S.dash.today)) + '</h2>' +
      '<p class="hint" title="A count is waybills that finished that stage today. A shortage is shown beside the count, not inside it.">Waybills that finished the stage today. Shortages are counted apart.</p>' +
      '<div class="tiles">' + tiles + '</div>' +
      mapQueueHtml() +
      '<h2>Open flags</h2>' +
      '<button type="button" class="ghost" data-action="toggle-resolved">' + (S.showResolved ? 'Hide resolved' : 'Show resolved') + '</button>' +
      flags +
      '<h2>Backlog</h2>' + backlog +
      '<h2>Who did each stage</h2>' + ownersHtml() +
      '<p class="note">J&amp;T import still runs from the Google Sheet menu. Stock changes only there.</p>' +
      '<button type="button" class="wide" data-go="jt">Import J&T export</button>' +
      '<button type="button" class="wide" data-go="labels">Print SKU labels</button></main>';
  }

  function pendingMaps() {
    return ((S.model && S.model.toMap) || []).filter(function (row) {
      return String(row.status || '').toUpperCase() === 'PENDING';
    });
  }

  function mapQueueHtml() {
    var rows = pendingMaps();
    if (!rows.length || !S.staff || !L.roleCanUseAdmin(S.staff)) return '';
    return '<div class="banner bad" role="alert"><span>' + rows.length + ' items need a SKU mapping</span></div>' +
      rows.map(function (row, index) {
        var options = (S.model.items || []).filter(function (item) {
          return item.active !== false && String(item.client || '').toUpperCase() === String(row.client || '').toUpperCase();
        }).map(function (item) {
          var selected = String(item.sku || '').toUpperCase() === String(row.suggestedSku || '').toUpperCase();
          return '<option' + (selected ? ' selected' : '') + '>' + esc(item.sku) + '</option>';
        }).join('');
        return '<div class="card"><b>' + esc(row.token) + '</b><p class="hint">' + esc(row.client) +
          ' · ' + (row.waybillCount || 1) + ' waybill(s)</p>' +
          '<select data-map="' + index + '" style="width:100%;height:44px;margin:8px 0">' + options + '</select>' +
          '<button type="button" class="wide" data-action="approve-map" data-map="' + index + '">Approve SKU</button></div>';
      }).join('');
  }

  function approveMap(el) {
    var index = Number(el.getAttribute('data-map'));
    var row = pendingMaps()[index];
    var select = document.querySelector('select[data-map="' + index + '"]');
    var sku = select ? select.value : '';
    if (!row || !sku) return;
    if (!S.token || !S.online) {
      setBanner({ kind: 'bad', text: 'Stay online to approve a mapping.' }, true);
      render();
      return;
    }
    row.approvedSku = sku;
    var plan = L.applyApprovedMappings(S.model, new Date(), S.staff && S.staff.email);
    var chain = writeValues('SKU_MAP!D1', [['Qty']]).then(function () {
      return appendRows('SKU_MAP', plan.skuMapAdded || []);
    });
    (plan.toMap || []).forEach(function (item) {
      if (!item._row) return;
      chain = chain.then(function () {
        var values = L.recordsToMatrix([item], L.SHEETS.TO_MAP.fields);
        return writeValues('TO_MAP!A' + item._row, values);
      });
    });
    (plan.waybills || []).forEach(function (item) {
      if (!item._row) return;
      chain = chain.then(function () {
        return writeValues('WAYBILLS!C' + item._row, [[item.itemString || '']]);
      }).then(function () {
        return writeValues('WAYBILLS!K' + item._row, [[item.unmapped || '']]);
      });
    });
    (plan.flags || []).forEach(function (item) {
      if (!item._row) return;
      chain = chain.then(function () {
        return writeValues('FLAGS!G' + item._row, [[true]]);
      });
    });
    chain.then(function () { return loadSheet(); }).then(function () {
      S.banner = { kind: 'good', text: 'Mapping saved. Pick is open once the item is on the waybill.' };
      if (S.screen === 'admin') render();
    }).catch(function () {
      setBanner({ kind: 'bad', text: 'Could not save the mapping.' }, true);
      render();
    });
  }

  function ownersHtml() {
    var rows = (S.dash && S.dash.stageOwners) || [];
    if (!rows.length) return '<p class="hint">No completed stages yet.</p>';
    return rows.map(function (row) {
      return '<div class="staff-row"><b class="wb">' + esc(row.waybill) + '</b><span>PICK ' + esc(row.PICK || '—') +
        ' · PACK ' + esc(row.PACK || '—') + ' · DISPATCH ' + esc(row.DISPATCH || '—') + '</span></div>';
    }).join('');
  }

  function jtHtml() {
    var preview = S.jtPreview;
    var body = '';
    if (preview) {
      var clients = preview.byClient || {};
      var unknown = (preview.unknownItems || []).slice(0, 12).map(function (row) {
        return '<li>' + esc(row.client) + ' · ' + esc(row.name) + ' · ' + row.count + '</li>';
      }).join('') || '<li>None</li>';
      body = '<div class="card"><b>Preview</b>' +
        '<p>ECOM ' + (clients.ECOM || 0) + ' · SIRWEN ' + (clients.SIRWEN || 0) + ' · TREXO ' + (clients.TREXO || 0) + '</p>' +
        '<p>New waybills: ' + preview.newWaybills + '</p>' +
        '<p>Status updates: ' + preview.updated + '</p>' +
        '<p>Dropped rows: ' + preview.dropped + '</p>' +
        '<p>Unknown item names</p><ul>' + unknown + '</ul>' +
        '<p class="hint">Stock is not changed. Only status is updated on waybills that are already in the sheet.</p>' +
        '<button type="button" class="wide good" data-action="jt-write">Write to the sheet</button></div>';
    }
    var note = S.jtMessage ? '<p class="banner warn">' + esc(S.jtMessage) + '</p>' : '';
    return headerHtml('Import') + '<main>' + note +
      '<p class="hint">Choose the J&T My Waybill xlsx. Other creator codes are dropped. Personal columns are not stored.</p>' +
      '<input id="jt-file" type="file" accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">' +
      '<p id="jt-status" class="hint">' + (S.jtBusy ? 'Reading the file…' : '') + '</p>' +
      body + '</main>';
  }

  function loadXlsx() {
    return new Promise(function (resolve, reject) {
      if (window.XLSX) return resolve();
      var script = document.createElement('script');
      script.src = 'vendor/xlsx.mini.min.js';
      script.onload = function () { resolve(); };
      script.onerror = function () { reject(new Error('xlsx')); };
      document.head.appendChild(script);
    });
  }

  function readJtFile(file) {
    S.jtBusy = true;
    S.jtMessage = 'Reading the file…';
    render();
    var reader = new FileReader();
    reader.onload = function () {
      setTimeout(function () {
        try {
          var wb = XLSX.read(new Uint8Array(reader.result), { type: 'array' });
          var name = wb.SheetNames.indexOf('Sheet0') >= 0 ? 'Sheet0' : wb.SheetNames[0];
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '', raw: false });
          var kept = [];
          var index = 0;
          function step() {
            var end = Math.min(index + 4000, rows.length);
            for (; index < end; index++) {
              var view = L.jtExportView(rows[index]);
              if (!view.waybillNo && !view.creatorCode) continue;
              kept.push({
                'Creator Code': view.creatorCode,
                'Waybill Number': view.waybillNo,
                'Order Status': view.orderStatus,
                'SigningTime': view.signingTime,
                'Submission Time': view.submissionTime,
                'RTS Reason': view.rtsReason,
                'Item Name': view.itemName
              });
            }
            if (index < rows.length) {
              S.jtMessage = 'Filtering ' + index + ' / ' + rows.length;
              var status = $('jt-status');
              if (status) status.textContent = S.jtMessage;
              setTimeout(step, 0);
              return;
            }
            S.jtPlan = L.planJtExport(kept, S.model, new Date());
            S.jtPreview = S.jtPlan.preview;
            S.jtBusy = false;
            S.jtMessage = '';
            render();
          }
          step();
        } catch (err) {
          S.jtBusy = false;
          S.jtMessage = 'Could not read that file.';
          render();
        }
      }, 30);
    };
    reader.readAsArrayBuffer(file);
  }

  function writeJtPlan() {
    var plan = S.jtPlan;
    if (!plan || S.demo) {
      S.jtMessage = S.demo ? 'Demo preview only. Nothing was written.' : 'Choose a file first.';
      render();
      return;
    }
    if (!S.token || !S.online) {
      S.jtMessage = 'Sign in and stay online to write the sheet.';
      render();
      return;
    }
    S.jtBusy = true;
    S.jtMessage = 'Writing…';
    render();
    var inserts = plan.inserts || [];
    var chain = appendRows('WAYBILLS', inserts.map(function (row) {
      return {
        waybillNo: row.waybillNo,
        client: row.client,
        itemString: row.itemString,
        orderDate: row.orderDate,
        importedAt: row.importedAt,
        jtStatus: row.jtStatus,
        pickupTime: '',
        rawItems: row.rawItems,
        signingTime: row.signingTime,
        rtsReason: row.rtsReason,
        unmapped: row.unmapped
      };
    }));
    (plan.updates || []).forEach(function (update) {
      if (!update._row) return;
      var fields = L.SHEETS.WAYBILLS.fields;
      function put(field) {
        var col = fields.indexOf(field) + 1;
        if (col < 1) return;
        chain = chain.then(function () {
          return writeValues('WAYBILLS!' + L.columnName(col) + update._row, [[update[field] || '']]);
        });
      }
      put('jtStatus');
      put('signingTime');
      put('rtsReason');
      if (update.remap) {
        put('itemString');
        put('rawItems');
        put('unmapped');
      }
    });
    chain = chain.then(function () { return appendRows('FLAGS', plan.flags || []); });
    chain = chain.then(function () { return appendRows('TO_MAP', plan.toMap || []); });
    chain.then(function () {
      S.jtBusy = false;
      S.jtMessage = 'Wrote ' + inserts.length + ' new waybills and ' + (plan.updates || []).length + ' status updates.';
      S.jtPlan = null;
      return loadSheet();
    }).then(function () {
      render();
    }).catch(function () {
      S.jtBusy = false;
      S.jtMessage = 'The write stopped. Run it again. Existing rows are not duplicated.';
      render();
    });
  }

  function labelsHtml() {
    var rows = (S.catalog || []).map(function (item) {
      var key = L.labelKey(item);
      var qty = S.labelQty[key] || 1;
      var checked = S.labelOn[key] ? ' checked' : '';
      return '<div class="sku-row"><label><input type="checkbox" data-key="' + esc(key) + '"' + checked + '> <b>' +
        esc(item.sku) + '</b></label><span class="chips">' + esc(item.clientLabel) + ' · ' + esc(item.name) +
        (item.barcode ? ' · ' + esc(item.barcode) : '') +
        '</span><input type="number" min="1" max="100" value="' + qty + '" data-qty="' + esc(key) + '"></div>';
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
    else if (S.screen === 'jt') app.innerHTML = jtHtml();
    else app.innerHTML = stageHtml();
    var input = $('code-input');
    armBanner();
    focusScan();
    wireInstall();
  }

  function focusScan() {
    if (S.checking || S.scanner || (S.screen !== 'stage' && S.screen !== 'pick' && S.screen !== 'pack' && S.screen !== 'dispatch')) return;
    var input = $('code-input');
    if (!input || input.disabled) return;
    try { input.focus({ preventScroll: true }); } catch (err) { input.focus(); }
  }

  function setBanner(next, force) {
    var decision = L.nextBanner(S.banner, next || null, Date.now(), S.bannerHold || 0, !!force);
    S.banner = decision.banner;
    S.bannerHold = decision.holdUntil || 0;
    if (!decision.kept) S.bannerHeld = null;
  }

  function armBanner() {
    if (S.banner && (S.banner.kind === 'bad' || S.banner.kind === 'warn') && S.bannerStamp !== S.banner) {
      S.bannerStamp = S.banner;
      S.bannerHold = Date.now() + 4000;
    }
    if (S.bannerHeld === S.banner) return;
    S.bannerHeld = S.banner;
    if (S.bannerTimer) clearTimeout(S.bannerTimer);
    S.bannerTimer = null;
    if (!S.banner || (S.banner.kind !== 'bad' && S.banner.kind !== 'warn')) return;
    var held = S.banner;
    S.bannerTimer = setTimeout(function () {
      if (S.banner !== held) return;
      if (S.bannerHold && Date.now() < S.bannerHold) return;
      S.banner = null;
      S.bannerHeld = null;
      S.bannerHold = 0;
      paint();
    }, 4000);
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
        input.disabled = false;
        input.readOnly = !!S.checking;
        input.placeholder = S.checking ? 'Checking…' : (S.stage === 'DISPATCH' || !S.waybillNo ? 'Type the waybill' : 'Type the SKU or barcode');
      }
      var add = $('add-btn');
      if (add) add.disabled = !!S.checking;
      var shortBtn = $('short-btn');
      if (shortBtn) shortBtn.hidden = !L.shortageAllowed(view, S.stage);
      armBanner();
      focusScan();
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
    if (target === 'pick' || target === 'pack' || target === 'dispatch') {
      if (!L.stageAssigned(S.staff, target)) {
        S.banner = { kind: 'bad', text: L.MSG.STAGE_NOT_ASSIGNED };
        S.screen = 'home';
        render();
        return;
      }
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
    if (target === 'jt') {
      if (!L.roleCanUseAdmin(S.staff)) {
        S.screen = 'denied';
        render();
        return;
      }
      S.screen = 'jt';
      S.jtMessage = '';
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
    var pending = S.pendingScan || '';
    if (S.payload && pending) {
      S.pendingScan = '';
      render();
      onCode(pending);
      return;
    }
    render();
  }

  function rememberedBlock(stage, waybill) {
    var row = S.recentBlock && S.recentBlock[stage + '|' + waybill];
    if (!row) return null;
    if (Date.now() - row.at >= 60000) return null;
    return row;
  }

  function finishCheck() {
    if (!S.checking && !S.checkMessage) return;
    S.checking = false;
    if (S.checkMessage) {
      setBanner({ kind: 'bad', text: S.checkMessage }, true);
      S.checkMessage = '';
    }
  }

  function onCode(raw) {
    if (S.checking) return;
    var code = String(raw || '').trim();
    if (!code) {
      if (S.bannerHold && Date.now() < S.bannerHold) return;
      setBanner({ kind: 'bad', text: L.MSG.EMPTY }, true);
      paint();
      return;
    }
    var gate = L.stageEntryWhileLoading(!!S.payload);
    if (gate.queue) {
      S.pendingScan = code;
      setBanner({ kind: 'warn', text: gate.message }, true);
      paint();
      return;
    }
    if (gate.blocked) {
      setBanner({ kind: 'warn', text: gate.message }, true);
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
    var block = L.stageGate(decisionContext(), S.stage, no);
    if (block) {
      S.waybillNo = '';
      beep(false);
      if (block.flag === 'SAME_PERSON_STAGE') {
        var again = rememberedBlock(S.stage, no);
        if (again) {
          S.banner = { kind: 'bad', text: again.message };
          paint();
          return;
        }
        S.recentBlock = S.recentBlock || {};
        S.recentBlock[S.stage + '|' + no] = { at: Date.now(), message: block.message };
        S.checking = true;
        S.checkMessage = block.message;
        S.banner = { kind: 'warn', text: 'Checking…' };
        paint();
        enqueue({ action: 'OPEN', stage: S.stage, waybill: no, code: no }, false);
        if (!S.syncing) finishCheck();
      } else {
        S.banner = { kind: 'bad', text: block.message };
        paint();
      }
      return;
    }
    if (S.stage === 'PICK' && view) {
      var unknown = L.unknownItemMessage(view);
      if (unknown) {
        S.waybillNo = '';
        S.banner = { kind: 'bad', text: unknown };
        beep(false);
        paint();
        return;
      }
    }
    var gate = L.decide(view || null, {
      action: S.stage === 'DISPATCH' ? 'DISPATCH' : 'OPEN',
      stage: S.stage,
      waybill: no,
      code: no
    }, S.payload.codes, decisionContext());
    if (gate.flag === 'SAME_PERSON_STAGE' || gate.flag === 'STAGE_NOT_ASSIGNED') {
      S.banner = { kind: 'bad', text: gate.message };
      beep(false);
      enqueue({ action: 'OPEN', stage: S.stage, waybill: no, code: no }, false);
      return;
    }
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
    var decision = L.decide(view || null, { action: 'DISPATCH', stage: 'DISPATCH', waybill: no, code: no }, S.payload.codes, decisionContext());
    if (decision.result !== 'OK') {
      S.banner = { kind: 'bad', text: decision.message };
      beep(false);
      enqueue({ action: 'DISPATCH', stage: 'DISPATCH', waybill: no, code: no }, false);
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
    enqueue({ action: 'DISPATCH', stage: 'DISPATCH', waybill: no, code: '' }, true);
  }

  function scanItem(code) {
    var view = currentView();
    var event = { action: 'ITEM', stage: S.stage, waybill: S.waybillNo, code: code };
    var decision = L.decide(view, event, S.payload.codes, decisionContext());
    if (decision.result === 'OK') {
      S.banner = { kind: 'good', text: decision.sku + ' added' };
      beep(true);
    } else {
      S.banner = { kind: 'bad', text: decision.message };
      beep(false);
    }
    enqueue(event, decision.result === 'OK');
  }

  function submitShort() {
    var view = currentView();
    if (!view) return;
    var event = { action: 'SHORTAGE', stage: S.stage, waybill: S.waybillNo, code: '' };
    var decision = L.decide(view, event, S.payload.codes, decisionContext());
    if (decision.result !== 'OK') {
      S.banner = { kind: 'warn', text: decision.message };
      paint();
      return;
    }
    if (!window.confirm('Some pieces are missing. Submit with a shortage and save a flag?')) return;
    S.waybillNo = '';
    setBanner({ kind: 'good', text: decision.message }, true);
    beep(false);
    enqueue(event, true);
    refreshViews();
    paint();
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
    if (S.stage === 'PICK' && L.unknownItemMessage(view)) return;
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
      var key = L.labelKey(item);
      if (!S.labelOn[key]) return;
      items.push({
        sku: item.sku,
        barcode: item.barcode,
        name: item.name,
        clientLabel: item.clientLabel,
        copies: S.labelQty[key] || 1
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
    else if (el.getAttribute('data-action') === 'approve-map') approveMap(el);
    else if (el.getAttribute('data-action') === 'toggle-resolved') {
      S.showResolved = !S.showResolved;
      render();
    }
    else if (el.getAttribute('data-action') === 'signin') signIn();
    else if (el.getAttribute('data-action') === 'camera') startCamera();
    else if (el.getAttribute('data-action') === 'short') submitShort();
    else if (el.getAttribute('data-action') === 'done') nextStep();
    else if (el.getAttribute('data-action') === 'print-labels') printLabels();
    else if (el.getAttribute('data-action') === 'jt-write') writeJtPlan();
    else if (el.getAttribute('data-action') === 'mode-a4' || el.getAttribute('data-action') === 'mode-thermal') {
      S.labelMode = el.getAttribute('data-action') === 'mode-a4' ? 'a4' : 'thermal';
      render();
    } else if (el.hasAttribute('data-resolve')) resolveFlag(el);
  });

  document.body.addEventListener('keydown', function (event) {
    if (S.screen !== 'stage' || S.checking) return;
    if (event.metaKey || event.ctrlKey || event.altKey) return;
    var input = $('code-input');
    if (event.key === 'Enter') {
      if (event.target && event.target.id === 'code-input') return;
      event.preventDefault();
      var code = ((input && input.value) || S.draft || '').trim();
      if (input) input.value = '';
      S.draft = '';
      onCode(code);
      return;
    }
    if (!event.key || event.key.length !== 1) return;
    if (event.target && event.target.id === 'code-input') return;
    S.draft = (S.draft || '') + event.key;
    if (input) {
      input.disabled = false;
      input.value = (input.value || '') + event.key;
      try { input.focus({ preventScroll: true }); } catch (err) { input.focus(); }
    }
  });

  document.body.addEventListener('submit', function (event) {
    if (!event.target || event.target.id !== 'scan-form') return;
    event.preventDefault();
    var input = $('code-input');
    var value = input.value;
    input.value = '';
    onCode(value);
  });

  document.body.addEventListener('change', function (event) {
    var sku = event.target.getAttribute && (event.target.getAttribute('data-key') || event.target.getAttribute('data-sku'));
    var qty = event.target.getAttribute && event.target.getAttribute('data-qty');
    if (event.target && event.target.id === 'jt-file' && event.target.files && event.target.files[0]) {
      var file = event.target.files[0];
      loadXlsx().then(function () { readJtFile(file); }).catch(function () {
        S.jtMessage = 'Could not load the spreadsheet reader.';
        render();
      });
      return;
    }
    if (sku) S.labelOn[sku] = !!event.target.checked;
    if (qty) S.labelQty[qty] = Math.max(1, Math.min(100, parseInt(event.target.value, 10) || 1));
  });

  window.addEventListener('beforeinstallprompt', function (event) {
    event.preventDefault();
    deferredInstall = event;
    wireInstall();
  });

  function loadGis() {
    if (params.get('offline') === '1' || S.demo) return;
    if (window.google && google.accounts) {
      if (!S.tokenClient) initGis();
      var ready = readToken();
      if (ready && !S.token) S.token = ready;
      if (S.token && !S.model) afterToken();
      else syncNow();
      return;
    }
    if (S.gisLoading) return;
    S.gisLoading = true;
    var script = document.createElement('script');
    script.src = 'https://accounts.google.com/gsi/client';
    script.onload = function () {
      S.gisLoading = false;
      initGis();
      var token = readToken();
      if (token) {
        S.token = token;
        if (!S.model) afterToken();
        else syncNow();
      }
    };
    script.onerror = function () { S.gisLoading = false; };
    document.head.appendChild(script);
  }

  function noteConnection() {
    if (params.get('offline') === '1') return;
    var online = navigator.onLine !== false;
    var changed = online !== S.online;
    S.online = online;
    if (changed) paint();
    if (online && pendingCount()) loadGis();
  }

  window.addEventListener('online', noteConnection);
  window.addEventListener('offline', noteConnection);
  setInterval(noteConnection, 15000);

  function watchServiceWorker() {
    if (!('serviceWorker' in navigator)) return;
    var hadController = !!navigator.serviceWorker.controller;
    var reloading = false;
    navigator.serviceWorker.addEventListener('controllerchange', function () {
      if (!hadController || reloading) return;
      reloading = true;
      window.location.reload();
    });
    navigator.serviceWorker.register('./sw.js', { scope: './' }).catch(function () {});
  }

  function boot() {
    if (!L) {
      document.body.textContent = 'Scan logic did not load.';
      return;
    }
    var legacyNotice = '';
    try { legacyNotice = Keys.retireLegacyQueue(window.localStorage); } catch (err) { legacyNotice = ''; }
    loadQueue();
    if (legacyNotice) S.banner = { kind: 'warn', text: legacyNotice };
    watchServiceWorker();
    if (S.demo) {
      startDemo();
      return;
    }
    loadCache();
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
    if (S.online) loadGis();
  }

  boot();
})();
