/**
 * Pure warehouse rules for Vendes Scan.
 * No SpreadsheetApp calls in the decision functions. Sheet reads and writes
 * go through a tiny sheet object so Node tests can pass a mock.
 *
 * STOCK RULE: scans never change stock. Stock drops only when a J&T import
 * says the rider picked the waybill up. That posts SHIP_OUT ledger rows
 * once per waybill. Importing the same file again does not deduct twice.
 *
 * Edit J&T column names in Config.js (VENDES_CONFIG), not here.
 */
(function (root, factory) {
  'use strict';
  var api = factory();
  if (typeof module !== 'undefined' && module && module.exports) {
    try {
      module.exports = api;
    } catch (ignore) {}
  }
  // HtmlService runs each script in the sandboxed iframe. Assigning only to
  // globalThis can miss the window the next script reads. window is shared.
  var host = (typeof window !== 'undefined' && window) ? window : root;
  if (host) host.VendesLogic = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  var SHEET_ORDER = ['STAFF', 'ITEMS', 'SKU_MAP', 'WAYBILLS', 'SCANS', 'FLAGS', 'LEDGER', 'STOCK', 'IMPORT', 'SETTINGS'];

  var SHEETS = {
    STAFF: {
      name: 'STAFF',
      headers: ['Email', 'Name', 'Role', 'Active', 'Pick', 'Pack', 'Dispatch'],
      fields: ['email', 'name', 'role', 'active', 'pick', 'pack', 'dispatch']
    },
    ITEMS: {
      name: 'ITEMS',
      headers: ['SKU', 'Client', 'Product Name', 'Barcode', 'Beginning Qty', 'Active', 'J&T Alias'],
      fields: ['sku', 'client', 'productName', 'barcode', 'beginningQty', 'active', 'jtAlias']
    },
    SKU_MAP: {
      name: 'SKU_MAP',
      headers: ['Raw Text', 'Client', 'SKU'],
      fields: ['raw', 'client', 'sku']
    },
    WAYBILLS: {
      name: 'WAYBILLS',
      headers: ['Waybill No', 'Client', 'Item String', 'Order Date', 'Imported At', 'J&T Status', 'Pickup Time', 'Raw Items', 'Signing Time', 'RTS Reason', 'Unmapped'],
      fields: ['waybillNo', 'client', 'itemString', 'orderDate', 'importedAt', 'jtStatus', 'pickupTime', 'rawItems', 'signingTime', 'rtsReason', 'unmapped']
    },
    SCANS: {
      name: 'SCANS',
      headers: ['Timestamp', 'Staff Email', 'Staff Name', 'Stage', 'Waybill', 'SKU', 'Qty', 'Result', 'Message'],
      fields: ['timestamp', 'staffEmail', 'staffName', 'stage', 'waybill', 'sku', 'qty', 'result', 'message']
    },
    FLAGS: {
      name: 'FLAGS',
      headers: ['Timestamp', 'Waybill', 'Client', 'Type', 'Details', 'Staff', 'Resolved'],
      fields: ['timestamp', 'waybill', 'client', 'type', 'details', 'staff', 'resolved']
    },
    LEDGER: {
      name: 'LEDGER',
      headers: ['Date', 'Client', 'SKU', 'Type', 'Qty', 'Ref', 'Source'],
      fields: ['date', 'client', 'sku', 'type', 'qty', 'ref', 'source']
    },
    STOCK: {
      name: 'STOCK',
      headers: ['SKU', 'Client', 'Product Name', 'Beginning', 'Ledger Qty', 'On Hand'],
      fields: ['sku', 'client', 'productName', 'beginning', 'ledgerQty', 'onHand']
    },
    IMPORT: {
      name: 'IMPORT',
      headers: ['Waybill No', 'Item String', 'J&T Status', 'Pickup Time', 'Order Date', 'Client'],
      fields: ['waybillNo', 'itemText', 'jtStatus', 'pickupTime', 'orderDate', 'client']
    },
    SETTINGS: {
      name: 'SETTINGS',
      headers: ['Setting', 'Value'],
      fields: ['setting', 'value']
    }
  };

  var MSG = {
    WRONG_SKU: 'Wrong SKU — not on this waybill',
    UNKNOWN_CODE: 'Unknown code',
    QTY_OVER: 'Too many — over the expected quantity',
    SKIP_PACK: 'Stage skipped — finish PICK first',
    SKIP_DISPATCH: 'Stage skipped — pack before dispatch',
    NOT_PICKED: 'Not picked yet',
    NOT_PACKED_OR_PICKED: 'Not packed yet / not picked yet',
    NOT_PACKED: 'Not packed yet',
    ITEM_NOT_IN_LIST: 'Item not in list: ',
    UNKNOWN_WB: 'Unknown waybill — not on the list',
    ALREADY_DISPATCHED: 'Already dispatched',
    PICKED_UP: 'Already picked up by the J&T rider',
    SHORTAGE_OK: 'Shortage submitted. Flag saved.',
    NO_SHORTAGE: 'Already complete — no shortage',
    NO_ITEMS: 'No mapped items',
    ALREADY_SHORT: 'Shortage was already submitted',
    EMPTY: 'Nothing entered',
    WAYBILL_AS_ITEM: "That's a waybill, scan the item barcode",
    TOO_LONG: 'Too long (max 64 characters)',
    LOADING: 'Still loading',
    STAGE_NOT_ASSIGNED: 'You are not assigned to this stage.',
    UNMAPPED: 'Unmapped items. Map the J&T name before picking.',
    OK: 'OK'
  };

  var SCAN_TEXT_MAX = 64;

  var FALLBACK_CONFIG = {
    pickedUpStatuses: [
      'picked up',
      'picked up by rider',
      'pickup success',
      'successful pickup',
      'already picked up',
      'in transit',
      'on the way',
      'out for delivery',
      'delivered',
      'departed',
      'departure'
    ],
    pickupTimeMeansPickedUp: true,
    staleDispatchHours: 24,
    recentPickupHours: 48,
    headerRow: 1,
    columns: {}
  };

  function cfg(config) {
    if (config) return config;
    if (typeof VENDES_CONFIG !== 'undefined') return VENDES_CONFIG;
    return FALLBACK_CONFIG;
  }

  function normalizeSku(value) {
    return String(value == null ? '' : value).replace(/\s+/g, ' ').trim().toUpperCase();
  }

  function normHeader(value) {
    return String(value == null ? '' : value).toLowerCase().replace(/&/g, '').replace(/[^a-z0-9]+/g, '');
  }

  function normalizeClient(value) {
    return String(value == null ? '' : value).trim().toUpperCase();
  }

  function normalizeEmail(value) {
    return String(value == null ? '' : value).trim().toLowerCase();
  }

  function normWaybill(value) {
    if (typeof value === 'number' && isFinite(value)) {
      return String(Math.round(value));
    }
    return String(value == null ? '' : value).trim().toUpperCase().replace(/\s+/g, '');
  }

  function clientLabel(code) {
    var client = normalizeClient(code);
    if (client === 'TREXO') return 'TREXO (Josie)';
    return client || '—';
  }

  function isTruthy(value) {
    if (value === true || value === 1) return true;
    if (value === false || value === 0 || value == null) return false;
    var text = String(value).trim().toLowerCase();
    return text === 'true' || text === 'yes' || text === 'y' || text === '1';
  }

  function isResolved(flag) {
    return !!(flag && isTruthy(flag.resolved));
  }

  function isActiveItem(item) {
    if (!item || !normalizeSku(item.sku)) return false;
    if (item.active === false || item.active === 0) return false;
    if (item.active === '' || item.active == null) return true;
    return isTruthy(item.active);
  }

  function num(value) {
    if (typeof value === 'number' && isFinite(value)) return value;
    if (value == null || value === '') return 0;
    var parsed = parseFloat(String(value).replace(/,/g, ''));
    return isFinite(parsed) ? parsed : 0;
  }

  function clampQty(value) {
    var qty = parseInt(value, 10);
    if (!qty || qty < 1) return 1;
    if (qty > 999) return 999;
    return qty;
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function formatManila(date) {
    var instant = date instanceof Date ? date : new Date(date);
    if (isNaN(instant.getTime())) return '';
    var parts = new Intl.DateTimeFormat('en-GB', {
      timeZone: 'Asia/Manila',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit',
      hourCycle: 'h23'
    }).formatToParts(instant);
    function get(type) {
      for (var i = 0; i < parts.length; i++) {
        if (parts[i].type === type) return parts[i].value;
      }
      return '00';
    }
    var hour = get('hour');
    if (hour === '24') hour = '00';
    return get('year') + '-' + get('month') + '-' + get('day') + ' ' + hour + ':' + get('minute') + ':' + get('second');
  }

  function manilaDate(date) {
    var full = formatManila(date);
    return full ? full.slice(0, 10) : '';
  }

  function parseManila(value) {
    if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
    var text = String(value == null ? '' : value).trim();
    if (!text || text === '-' || text === '—' || text.toLowerCase() === 'n/a' || text.toLowerCase() === 'null') {
      return null;
    }
    var match = text.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if (match) {
      return new Date(Date.UTC(
        Number(match[1]),
        Number(match[2]) - 1,
        Number(match[3]),
        Number(match[4] || 0) - 8,
        Number(match[5] || 0),
        Number(match[6] || 0)
      ));
    }
    var parsed = new Date(text);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  function asManilaString(value) {
    if (value == null || value === '') return '';
    if (value instanceof Date) return formatManila(value);
    var text = String(value).trim();
    if (!text) return '';
    if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(text)) return text;
    var parsed = parseManila(text);
    return parsed ? formatManila(parsed) : text;
  }

  function asDateOnly(value) {
    if (value == null || value === '') return '';
    if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) return value.trim();
    var full = asManilaString(value);
    return full ? full.slice(0, 10) : '';
  }

  function hasPickupTime(value) {
    if (value instanceof Date) return !isNaN(value.getTime());
    var text = String(value == null ? '' : value).trim();
    if (!text || text === '-' || text === '—' || text.toLowerCase() === 'n/a' || text.toLowerCase() === 'null') {
      return false;
    }
    return true;
  }

  function statusMeansPickedUp(status, config) {
    var text = String(status || '').trim().toLowerCase();
    if (!text) return false;
    if (/fail|cancel|unsuccessful|void|reject|pending|for pickup|wait/.test(text)) return false;
    var list = (config && config.pickedUpStatuses) || FALLBACK_CONFIG.pickedUpStatuses;
    for (var i = 0; i < list.length; i++) {
      var alias = String(list[i] || '').trim().toLowerCase();
      if (!alias) continue;
      if (text === alias || text.indexOf(alias) === 0) return true;
    }
    return false;
  }

  function isPickedUp(status, pickupTime, config) {
    var conf = cfg(config);
    var failed = /fail|cancel|unsuccessful|void|reject/.test(String(status || '').toLowerCase());
    if (!failed && conf.pickupTimeMeansPickedUp !== false && hasPickupTime(pickupTime)) return true;
    return statusMeansPickedUp(status, conf);
  }

  function mergeQty(items) {
    var out = [];
    var index = {};
    (items || []).forEach(function (item) {
      var sku = normalizeSku(item.sku);
      if (!sku) return;
      var qty = clampQty(item.qty);
      if (index[sku] == null) {
        index[sku] = out.length;
        out.push({
          sku: sku,
          qty: qty,
          client: item.client ? normalizeClient(item.client) : '',
          name: item.name || item.productName || ''
        });
      } else {
        out[index[sku]].qty = Math.min(999, out[index[sku]].qty + qty);
      }
    });
    return out;
  }

  function parseItemSegments(text) {
    var raw = String(text || '').trim();
    if (!raw) return [];
    var parts = raw.split(/\s*(?:\+|\;|\||\n|\r)+\s*/);
    var out = [];
    parts.forEach(function (part) {
      var token = String(part || '').replace(/\s+/g, ' ').trim();
      if (!token) return;
      var lead = token.match(/^(\d+)\s*x\s*(.+)$/i);
      if (lead) {
        out.push({ name: lead[2].trim(), qty: clampQty(lead[1]) });
        return;
      }
      var pcs = token.match(/^(\d+)\s*(?:pcs|pc|pieces|piece|sets|set)\s+(.+)$/i);
      if (pcs) {
        out.push({ name: pcs[2].trim(), qty: clampQty(pcs[1]) });
        return;
      }
      var trail = token.match(/^(.+?)\s*x\s*(\d+)$/i);
      if (trail) {
        out.push({ name: trail[1].trim(), qty: clampQty(trail[2]) });
        return;
      }
      var pcsAfter = token.match(/^(.+?)\s+(\d+)\s*(?:pcs|pc|pieces|piece|sets|set)$/i);
      if (pcsAfter) {
        out.push({ name: pcsAfter[1].trim(), qty: clampQty(pcsAfter[2]) });
        return;
      }
      out.push({ name: token, qty: 1 });
    });
    return out;
  }

  function tryParseCompact(text) {
    var segments = parseItemSegments(text);
    if (!segments.length) return null;
    return mergeQty(segments.map(function (seg) {
      return { sku: normalizeSku(seg.name), qty: seg.qty, client: '', name: '' };
    }));
  }

  function parseItemString(text) {
    var segments = parseItemSegments(text);
    return segments.map(function (seg) {
      return { sku: normalizeSku(seg.name), qty: seg.qty, client: '', name: '' };
    });
  }

  function formatItemString(items) {
    return mergeQty(items).map(function (item) {
      return item.qty + 'x' + item.sku;
    }).join('+');
  }

  function findItem(items, sku, client) {
    var want = normalizeSku(sku);
    var hint = client ? normalizeClient(client) : '';
    var matches = (items || []).filter(function (item) {
      return isActiveItem(item) && normalizeSku(item.sku) === want;
    });
    if (!matches.length) return null;
    if (hint) {
      var byClient = matches.filter(function (item) {
        return normalizeClient(item.client) === hint;
      });
      return byClient[0] || null;
    }
    var clients = {};
    matches.forEach(function (item) { clients[normalizeClient(item.client)] = true; });
    if (Object.keys(clients).length > 1) return { ambiguous: true, sku: want };
    return matches[0];
  }

  function qtyNear(source, needle) {
    var hay = String(source || '').toLowerCase();
    var token = String(needle || '').toLowerCase().trim();
    var idx = token ? hay.indexOf(token) : -1;
    if (idx < 0) return 1;
    var before = hay.slice(Math.max(0, idx - 20), idx);
    var after = hay.slice(idx + token.length, idx + token.length + 12);
    var match = before.match(/(\d+)\s*x\s*$/);
    if (match) return clampQty(match[1]);
    match = after.match(/^\s*x\s*(\d+)/);
    if (match) return clampQty(match[1]);
    match = before.match(/(\d+)\s*(?:pcs|pc|pieces|piece|sets|set)\s*$/);
    if (match) return clampQty(match[1]);
    match = after.match(/^\s*(\d+)\s*(?:pcs|pc|pieces|piece|sets|set)\b/);
    if (match) return clampQty(match[1]);
    return 1;
  }

  function blankRange(text, start, length) {
    return text.slice(0, start) + new Array(length + 1).join(' ') + text.slice(start + length);
  }

  function findToken(hay, token) {
    var re = new RegExp('(^|[^a-z0-9])' + escapeRegExp(String(token).toLowerCase()) + '([^a-z0-9]|$)', 'i');
    var match = re.exec(String(hay).toLowerCase());
    if (!match) return -1;
    return match.index + match[1].length;
  }

  function findSkuMention(hay, sku) {
    var at = findToken(hay, sku);
    if (at >= 0) return at;
    var needle = String(sku).toLowerCase();
    var re = new RegExp('(?:^|[^a-z0-9])\\d+\\s*x\\s*' + escapeRegExp(needle) + '(?=[^a-z0-9]|$)', 'i');
    var match = re.exec(String(hay).toLowerCase());
    if (!match) return -1;
    return match.index + match[0].toLowerCase().lastIndexOf(needle);
  }

  function leftoverPhrase(working) {
    var cleaned = String(working || '')
      .replace(/\d+/g, ' ')
      .replace(/\bx\b/g, ' ')
      .replace(/[^a-z0-9]+/g, ' ');
    var stop = {
      pcs: 1, pc: 1, piece: 1, pieces: 1, set: 1, sets: 1,
      the: 1, and: 1, with: 1, for: 1, item: 1, items: 1, order: 1, qty: 1
    };
    var words = cleaned.split(/\s+/).filter(function (word) {
      return word && word.length >= 3 && !stop[word];
    });
    return words.length ? [words.join(' ')] : [];
  }

  function onlyClient(items) {
    var found = {};
    (items || []).forEach(function (item) {
      if (item.client) found[normalizeClient(item.client)] = true;
    });
    var keys = Object.keys(found);
    return keys.length === 1 ? keys[0] : '';
  }

  function formatSegments(segments) {
    var out = [];
    var index = {};
    (segments || []).forEach(function (seg) {
      var label = seg.unknown ? String(seg.name || '').replace(/\s+/g, ' ').trim() : seg.sku;
      if (!label) return;
      var key = (seg.unknown ? '?' : '=') + normalizeSku(label);
      if (index[key] == null) {
        index[key] = out.length;
        out.push({ label: label, qty: clampQty(seg.qty) });
      } else {
        out[index[key]].qty = Math.min(999, out[index[key]].qty + clampQty(seg.qty));
      }
    });
    return out.map(function (row) { return row.qty + 'x' + row.label; }).join('+');
  }

  function unknownSegment(name, qty) {
    return {
      sku: '',
      qty: clampQty(qty),
      client: '',
      name: String(name || '').replace(/\s+/g, ' ').trim(),
      unknown: true
    };
  }

  /**
   * One whole segment becomes one SKU, or stays unknown with its own qty.
   * Order: exact SKU, exact J&T Alias, exact Product Name. Same client only.
   * A shorter word inside the name is not a match.
   */
  function resolveSegment(name, qty, client, items) {
    var label = normalizeSku(name);
    var hint = normalizeClient(client);
    if (!label) return unknownSegment(name, qty);
    var pool = (items || []).filter(function (item) {
      if (!isActiveItem(item)) return false;
      if (hint && normalizeClient(item.client) !== hint) return false;
      return true;
    });
    function hit(row) {
      return {
        sku: normalizeSku(row.sku),
        qty: clampQty(qty),
        client: normalizeClient(row.client),
        name: row.productName || '',
        unknown: false
      };
    }
    var bySku = pool.filter(function (item) { return normalizeSku(item.sku) === label; });
    if (bySku.length === 1) return hit(bySku[0]);
    var byAlias = pool.filter(function (item) {
      return aliasList(item.jtAlias).some(function (alias) { return normalizeSku(alias) === label; });
    });
    if (byAlias.length === 1) return hit(byAlias[0]);
    var byName = pool.filter(function (item) {
      return normalizeSku(item.productName) === label;
    });
    if (byName.length === 1) return hit(byName[0]);
    return unknownSegment(name, qty);
  }

  function resolveLineItems(rawText, clientHint, skuMap, items) {
    var hint = normalizeClient(clientHint);
    var segments = parseItemSegments(rawText);
    if (!segments.length) {
      return { items: [], unknown: [], segments: [], client: hint, itemString: '' };
    }
    var resolved = segments.map(function (seg) {
      return resolveSegment(seg.name, seg.qty, hint, items);
    });
    var good = resolved.filter(function (row) { return !row.unknown; });
    var unknown = [];
    var seen = {};
    resolved.forEach(function (row) {
      if (!row.unknown) return;
      var key = normalizeSku(row.name);
      if (!key || seen[key]) return;
      seen[key] = true;
      unknown.push(row.name);
    });
    return {
      items: mergeQty(good),
      unknown: unknown,
      segments: resolved,
      client: hint || onlyClient(good) || '',
      itemString: formatSegments(resolved)
    };
  }

  function unknownItemMessage(view) {
    var names = [];
    var seen = {};
    ((view && view.lines) || []).forEach(function (line) {
      if (!line || !line.unknown) return;
      var name = String(line.name || line.sku || '').replace(/\s+/g, ' ').trim();
      var key = normalizeSku(name);
      if (!key || seen[key]) return;
      seen[key] = true;
      names.push(name);
    });
    if (!names.length && view && String(view.unmapped || '').trim()) {
      String(view.unmapped).split(',').forEach(function (part) {
        var name = part.replace(/\s+/g, ' ').trim();
        var key = normalizeSku(name);
        if (!key || seen[key]) return;
        seen[key] = true;
        names.push(name);
      });
    }
    if (!names.length) return '';
    return 'Item not in list: ' + names.join(', ') + ' — ask admin';
  }

  function missingStageMessage(view) {
    if (!view || !view.pickComplete) return 'Not picked yet';
    return 'Not packed yet';
  }

  function messageParts(message) {
    var text = String(message || '');
    var match = text.match(/^ID:(\S+)\s*([\s\S]*)$/);
    if (!match) return { id: '', body: text };
    return { id: match[1], body: match[2] || '' };
  }

  function withId(localId, body) {
    var id = String(localId || '').trim();
    var rest = body ? String(body) : '';
    if (!id) return rest;
    return 'ID:' + id + (rest ? ' ' + rest : '');
  }

  function lineTotals(scans, waybillNo) {
    var picked = {};
    var packed = {};
    var seen = {};
    (scans || []).forEach(function (scan) {
      if (normWaybill(scan.waybill) !== normWaybill(waybillNo)) return;
      if (String(scan.result || '').toUpperCase() !== 'OK') return;
      var parts = messageParts(scan.message);
      if (parts.id) {
        if (seen[parts.id]) return;
        seen[parts.id] = true;
      }
      if (String(parts.body).toUpperCase() === 'SHORTAGE_SUBMIT') return;
      var sku = normalizeSku(scan.sku);
      if (!sku) return;
      var qty = num(scan.qty);
      var stage = String(scan.stage || '').toUpperCase();
      if (stage === 'PICK') picked[sku] = (picked[sku] || 0) + qty;
      if (stage === 'PACK') packed[sku] = (packed[sku] || 0) + qty;
    });
    return { picked: picked, packed: packed };
  }

  function hasMarker(scans, waybillNo, stage, marker) {
    var want = String(marker).toUpperCase();
    return (scans || []).some(function (scan) {
      if (normWaybill(scan.waybill) !== normWaybill(waybillNo)) return false;
      if (String(scan.stage || '').toUpperCase() !== stage) return false;
      if (String(scan.result || '').toUpperCase() !== 'OK') return false;
      return String(messageParts(scan.message).body).toUpperCase() === want;
    });
  }

  function isDispatched(scans, waybillNo) {
    return (scans || []).some(function (scan) {
      return normWaybill(scan.waybill) === normWaybill(waybillNo)
        && String(scan.stage || '').toUpperCase() === 'DISPATCH'
        && String(scan.result || '').toUpperCase() === 'OK';
    });
  }

  function findWaybill(model, waybillNo) {
    var want = normWaybill(waybillNo);
    var rows = (model && model.waybills) || [];
    for (var i = 0; i < rows.length; i++) {
      if (normWaybill(rows[i].waybillNo) === want) return rows[i];
    }
    return null;
  }

  function buildWaybillView(waybill, items, scans, config) {
    var resolved = resolveLineItems(waybill.itemString, waybill.client, [], items);
    var totals = lineTotals(scans, waybill.waybillNo);
    var lines = [];
    var lineAt = {};
    (resolved.segments || []).forEach(function (part) {
      var key = part.unknown ? ('?' + normalizeSku(part.name)) : normalizeSku(part.sku);
      if (!key) return;
      if (lineAt[key] != null) {
        lines[lineAt[key]].expected += part.qty;
        return;
      }
      lineAt[key] = lines.length;
      if (part.unknown) {
        lines.push({
          sku: part.name,
          name: part.name,
          barcode: '',
          unknown: true,
          expected: part.qty,
          picked: 0,
          packed: 0
        });
        return;
      }
      var sku = normalizeSku(part.sku);
      var item = findItem(items, sku, waybill.client);
      lines.push({
        sku: sku,
        name: part.name || sku,
        barcode: item && !item.ambiguous ? String(item.barcode || '') : '',
        unknown: false,
        expected: part.qty,
        picked: totals.picked[sku] || 0,
        packed: totals.packed[sku] || 0
      });
    });
    var known = lines.filter(function (line) { return !line.unknown; });
    var pickComplete = known.length > 0 && lines.every(function (line) { return !line.unknown && line.picked >= line.expected; });
    var packComplete = known.length > 0 && lines.every(function (line) { return !line.unknown && line.packed >= line.expected; });
    return {
      waybillNo: normWaybill(waybill.waybillNo),
      client: normalizeClient(waybill.client),
      clientLabel: clientLabel(waybill.client),
      itemString: waybill.itemString || '',
      unmapped: String(waybill.unmapped || '').trim(),
      orderDate: waybill.orderDate || '',
      jtStatus: waybill.jtStatus || '',
      pickupTime: waybill.pickupTime || '',
      signingTime: waybill.signingTime || '',
      rtsReason: waybill.rtsReason || '',
      pickedUp: isPickedUp(waybill.jtStatus, waybill.pickupTime, config),
      lines: lines,
      pickComplete: pickComplete,
      packComplete: packComplete,
      pickClosed: pickComplete || hasMarker(scans, waybill.waybillNo, 'PICK', 'SHORTAGE_SUBMIT'),
      packClosed: packComplete || hasMarker(scans, waybill.waybillNo, 'PACK', 'SHORTAGE_SUBMIT'),
      dispatched: isDispatched(scans, waybill.waybillNo)
    };
  }

  function buildCodeIndex(items) {
    var codes = {};
    (items || []).forEach(function (item) {
      if (!isActiveItem(item)) return;
      var info = {
        sku: normalizeSku(item.sku),
        name: item.productName || '',
        client: normalizeClient(item.client)
      };
      codes[info.sku] = info;
      var barcode = String(item.barcode || '').trim().toUpperCase();
      if (barcode) codes[barcode] = info;
    });
    return codes;
  }

  function groupScans(scans) {
    var byWaybill = {};
    (scans || []).forEach(function (scan) {
      var no = normWaybill(scan.waybill);
      if (!no) return;
      if (!byWaybill[no]) byWaybill[no] = [];
      byWaybill[no].push(scan);
    });
    return byWaybill;
  }

  function scansFor(grouped, waybillNo) {
    if (!grouped) return [];
    return grouped[normWaybill(waybillNo)] || [];
  }

  function buildStagePayload(model, stage, config, now) {
    var conf = cfg(config);
    var moment = now || new Date();
    var hours = Number(conf.recentPickupHours);
    if (!hours && hours !== 0) hours = 48;
    var grouped = groupScans(model.scans);
    var waybills = {};
    (model.waybills || []).forEach(function (waybill) {
      var picked = isPickedUp(waybill.jtStatus, waybill.pickupTime, conf);
      if (picked) {
        var when = parseManila(waybill.pickupTime) || parseManila(waybill.importedAt);
        if (!when) return;
        var age = (moment.getTime() - when.getTime()) / 36e5;
        if (age > hours) return;
      }
      var view = buildWaybillView(waybill, model.items, scansFor(grouped, waybill.waybillNo), conf);
      waybills[view.waybillNo] = view;
    });
    return {
      stage: String(stage || '').toUpperCase(),
      waybills: waybills,
      codes: buildCodeIndex(model.items),
      actors: stageActorMap(model.scans),
      samePersonStages: samePersonEnabled(model.settings)
    };
  }

  function lineOnView(view, code) {
    var key = normalizeSku(code);
    if (!key || !view) return null;
    var lines = view.lines || [];
    for (var i = 0; i < lines.length; i++) {
      if (lines[i].unknown) continue;
      if (normalizeSku(lines[i].sku) === key) return lines[i];
      if (lines[i].barcode && normalizeSku(lines[i].barcode) === key) return lines[i];
    }
    return null;
  }

  function stageQty(line, stage) {
    return stage === 'PACK' ? line.packed : line.picked;
  }

  function mismatchText(view) {
    var bits = [];
    (view.lines || []).forEach(function (line) {
      if (line.picked !== line.packed) {
        bits.push(line.sku + ' pick ' + line.picked + ' / pack ' + line.packed);
      }
    });
    return bits.join(', ');
  }

  function shortText(view, stage) {
    var bits = [];
    (view.lines || []).forEach(function (line) {
      var have = stageQty(line, stage);
      if (have < line.expected) bits.push(line.sku + ' ' + have + '/' + line.expected);
    });
    return bits.join(', ');
  }

  /**
   * Decide one phone action. Does not write anything.
   * event.action is ITEM, SHORTAGE, or DISPATCH.
   */
  function clipDisplay(text, limit) {
    var value = String(text == null ? '' : text);
    var max = limit || 80;
    if (value.length <= max) return value;
    return value.slice(0, max - 1) + '…';
  }

  function saveBadge(event) {
    return event === 'failed' ? 'saving' : 'saved';
  }

  function clipScanText(value) {
    var text = String(value == null ? '' : value).trim();
    if (text.length <= SCAN_TEXT_MAX) return text;
    return text.slice(0, SCAN_TEXT_MAX);
  }

  function limitScanText(value) {
    var text = String(value == null ? '' : value).trim();
    if (text.length > SCAN_TEXT_MAX) {
      return { ok: false, text: clipScanText(text), message: MSG.TOO_LONG };
    }
    return { ok: true, text: text, message: '' };
  }

  function labelBarcodeValue(itemOrSku, barcode) {
    var sku = '';
    var code = barcode;
    if (itemOrSku && typeof itemOrSku === 'object') {
      sku = itemOrSku.sku;
      if (code == null) code = itemOrSku.barcode;
    } else {
      sku = itemOrSku;
    }
    var filled = String(code == null ? '' : code).trim();
    if (filled) return filled;
    return String(sku == null ? '' : sku).trim();
  }

  function labelKey(item) {
    var sku = String(item && item.sku || '').trim();
    var client = String(item && item.client || '').trim();
    return sku + '|' + client;
  }

  function labelCaption(item) {
    var sku = String(item && item.sku || '').trim();
    var code = labelBarcodeValue(item);
    return code + ' SKU: ' + sku;
  }

  // Code 128 set B patterns. Index 103–105 are start codes. 106 is stop.
  var CODE128_PATTERNS = [
    '212222', '222122', '222221', '121223', '121322', '131222', '122213', '122312', '132212', '221213',
    '221312', '231212', '112232', '122132', '122231', '113222', '123122', '123221', '223211', '221132',
    '221231', '213212', '223112', '312131', '311222', '321122', '321221', '312212', '322112', '322211',
    '212123', '212321', '232121', '111323', '131123', '131321', '112313', '132113', '132311', '211313',
    '231113', '231311', '112133', '112331', '132131', '113123', '113321', '133121', '313121', '211331',
    '231131', '213113', '213311', '213131', '311123', '311321', '331121', '312113', '312311', '332111',
    '314111', '221411', '431111', '111224', '111422', '121124', '121421', '141122', '141221', '112214',
    '112412', '122114', '122411', '142112', '142211', '241211', '221114', '413111', '241112', '134111',
    '111242', '121142', '121241', '114212', '124112', '124211', '411212', '421112', '421211', '212141',
    '214121', '412121', '111143', '111341', '131141', '114113', '114311', '411113', '411311', '113141',
    '114131', '311141', '411131', '211412', '211214', '211232', '2331112'
  ];

  function code128Symbols(text) {
    var value = labelBarcodeValue(text);
    var symbols = [104];
    for (var i = 0; i < value.length; i++) {
      var code = value.charCodeAt(i);
      if (code < 32 || code > 126) return null;
      symbols.push(code - 32);
    }
    var sum = symbols[0];
    for (var j = 1; j < symbols.length; j++) sum += j * symbols[j];
    symbols.push(sum % 103);
    symbols.push(106);
    return symbols;
  }

  function code128Widths(text) {
    var symbols = code128Symbols(text);
    if (!symbols) return '';
    var widths = '';
    for (var i = 0; i < symbols.length; i++) widths += CODE128_PATTERNS[symbols[i]];
    return widths;
  }

  function escHtml(value) {
    return String(value == null ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function code128Svg(text) {
    var sku = labelBarcodeValue(text);
    var widths = code128Widths(sku);
    if (!widths) return '';
    var rects = '';
    var x = 0;
    var ink = true;
    for (var i = 0; i < widths.length; i++) {
      var w = Number(widths.charAt(i));
      if (ink) rects += '<rect x="' + x + '" y="0" width="' + w + '" height="40"/>';
      x += w;
      ink = !ink;
    }
    return '<svg xmlns="http://www.w3.org/2000/svg" data-sku="' + escHtml(sku) + '" viewBox="0 0 ' + x + ' 40" preserveAspectRatio="none">' + rects + '</svg>';
  }

  function labelPrintMode(mode) {
    var name = String(mode || 'a4').toLowerCase();
    if (name === 'thermal' || name === '50x30') {
      return { name: 'thermal', pageSize: '50mm 30mm', margin: '0', columns: 1, rows: 1, perPage: 1, breakAfter: 'page' };
    }
    return { name: 'a4', pageSize: 'A4', margin: '4mm', columns: 4, rows: 9, perPage: 36, breakAfter: 'auto' };
  }

  function expandLabelCopies(items) {
    var out = [];
    (items || []).forEach(function (item) {
      var sku = String(item && item.sku || '').trim();
      if (!sku) return;
      var code = labelBarcodeValue(item);
      var copies = parseInt(item.copies, 10);
      if (!copies || copies < 1) copies = 1;
      if (copies > 200) copies = 200;
      for (var i = 0; i < copies; i++) {
        out.push({
          sku: sku,
          code: code,
          barcode: code,
          name: item.name || '',
          clientLabel: item.clientLabel || ''
        });
      }
    });
    return out;
  }

  function labelPrintDocument(items, mode) {
    var spec = labelPrintMode(mode);
    var labels = expandLabelCopies(items);
    var pages = [];
    if (spec.name === 'thermal') {
      labels.forEach(function (label) { pages.push([label]); });
    } else {
      for (var i = 0; i < labels.length; i += spec.perPage) pages.push(labels.slice(i, i + spec.perPage));
    }
    var body = pages.map(function (page) {
      return '<section class="page">' + page.map(function (label) {
        return '<article class="label"><div class="who">' + escHtml(label.clientLabel) + '</div>' +
          '<div class="name">' + escHtml(label.name) + '</div>' +
          code128Svg(label.code || label.sku) +
          '<div class="sku">' + escHtml(label.code || label.sku) + '</div>' +
          '<div class="sku-line">SKU: ' + escHtml(label.sku) + '</div></article>';
      }).join('') + '</section>';
    }).join('');
    var pageRule = spec.name === 'thermal'
      ? '@page { size: 50mm 30mm; margin: 0; }'
      : '@page { size: A4; margin: 4mm; }';
    var layout = spec.name === 'thermal'
      ? '.page { width: 50mm; height: 30mm; page-break-after: always; break-after: page; } .label { width: 50mm; height: 30mm; }'
      : '.page { display: grid; grid-template-columns: repeat(4, 50mm); grid-template-rows: repeat(9, 30mm); gap: 1.4mm 0.6mm; justify-content: center; page-break-after: always; break-after: page; } .label { width: 50mm; height: 30mm; overflow: hidden; }';
    return '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Vendes labels</title><style>' +
      pageRule +
      'html,body{margin:0;padding:0;background:#fff;}' +
      layout +
      '.label{box-sizing:border-box;padding:1mm 1.5mm;font-family:sans-serif;}' +
      '.who{font-size:7pt;font-weight:700;}' +
      '.name{font-size:8pt;font-weight:700;white-space:nowrap;overflow:hidden;}' +
      '.sku{font-size:8pt;font-weight:800;font-family:monospace;line-height:1.1;}' +
      '.sku-line{font-size:7pt;font-weight:700;font-family:monospace;line-height:1.1;}' +
      'svg{display:block;width:100%;height:12mm;}' +
      'svg rect{fill:#000;}' +
      '</style></head><body>' + body + '</body></html>';
  }

  function roleCanUseAdmin(staff) {
    if (!staff || staff.active === false || staff.active === 0) return false;
    if (staff.active != null && staff.active !== '' && !isTruthy(staff.active)) return false;
    return String(staff.role || '').trim().toUpperCase() === 'ADMIN';
  }

  function stageFlagSet(staff) {
    return !!(staff && (isTruthy(staff.pick) || isTruthy(staff.pack) || isTruthy(staff.dispatch)));
  }

  function stageAssigned(staff, stage) {
    var want = String(stage || '').toUpperCase();
    if (!staff) return false;
    if (staff.active === false) return false;
    if (String(staff.role || '').toUpperCase() === 'ADMIN') return true;
    if (!stageFlagSet(staff)) return true;
    if (want === 'PICK') return isTruthy(staff.pick);
    if (want === 'PACK') return isTruthy(staff.pack);
    if (want === 'DISPATCH') return isTruthy(staff.dispatch);
    return false;
  }

  function samePersonEnabled(source) {
    if (!source || source.samePersonStages == null || source.samePersonStages === '') return true;
    return isTruthy(source.samePersonStages);
  }

  function samePersonConflict(scans, staff, waybill, stage) {
    var email = normalizeEmail(staff && staff.email);
    if (!email) return null;
    var want = String(stage || '').toUpperCase();
    var prefer = {
      DISPATCH: ['PACK', 'PICK'],
      PACK: ['PICK', 'DISPATCH'],
      PICK: ['PACK', 'DISPATCH']
    };
    var order = prefer[want] || ['PICK', 'PACK', 'DISPATCH'];
    var found = {};
    (scans || []).forEach(function (scan) {
      if (normWaybill(scan.waybill) !== normWaybill(waybill)) return;
      if (String(scan.result || '').toUpperCase() !== 'OK') return;
      var done = String(scan.stage || '').toUpperCase();
      if (done === want) return;
      if (done !== 'PICK' && done !== 'PACK' && done !== 'DISPATCH') return;
      if (normalizeEmail(scan.staffEmail) !== email) return;
      if (!found[done]) found[done] = { stage: done, name: scan.staffName || email, email: email };
    });
    for (var i = 0; i < order.length; i++) {
      if (found[order[i]]) return found[order[i]];
    }
    return null;
  }

  function samePersonMessage(actor, nextStage) {
    var did = actor.stage === 'PICK' ? 'picked' : actor.stage === 'PACK' ? 'packed' : 'dispatched';
    var next = String(nextStage || '').toUpperCase() === 'PICK' ? 'pick'
      : String(nextStage || '').toUpperCase() === 'PACK' ? 'pack' : 'dispatch';
    return 'You already ' + did + ' this waybill (' + (actor.name || actor.email) + '). Someone else must ' + next + ' it.';
  }

  function looksLikeWaybill(code, context) {
    var text = normWaybill(code);
    if (!text) return false;
    if (/^JT\d+$/.test(text)) return true;
    var known = context && context.waybillNos;
    if (!known) return false;
    if (Array.isArray(known)) {
      for (var i = 0; i < known.length; i++) {
        if (normWaybill(known[i]) === text) return true;
      }
      return false;
    }
    return !!known[text];
  }

  function stageGate(context, stage, waybill) {
    if (!context || !context.staff) return null;
    var want = String(stage || '').toUpperCase();
    if (want !== 'PICK' && want !== 'PACK' && want !== 'DISPATCH') return null;
    if (!stageAssigned(context.staff, want)) {
      return errorDecision(null, want, 'STAGE_NOT_ASSIGNED', MSG.STAGE_NOT_ASSIGNED, '', 0, MSG.STAGE_NOT_ASSIGNED);
    }
    if (!samePersonEnabled(context)) return null;
    var actor = samePersonConflict(context.scans, context.staff, waybill, want);
    if (!actor) return null;
    var text = samePersonMessage(actor, want);
    var blocked = errorDecision(null, want, 'SAME_PERSON_STAGE', text, '', 0, text);
    blocked.result = 'BLOCKED';
    return blocked;
  }

  function decide(view, event, codes, context) {
    var stage = String((event && event.stage) || '').toUpperCase();
    var action = String((event && event.action) || 'ITEM').toUpperCase();
    var waybillCheck = limitScanText(event && event.waybill);
    var codeCheck = limitScanText(event && event.code);
    if (!waybillCheck.ok || !codeCheck.ok) {
      return errorDecision(null, stage, '', MSG.TOO_LONG, '', 0, MSG.TOO_LONG);
    }
    var waybillNo = normWaybill(waybillCheck.text || codeCheck.text);
    var code = codeCheck.text;
    var gated = stageGate(context, stage, waybillNo);
    if (gated) {
      gated.waybill = waybillNo;
      return gated;
    }

    if (!view) {
      return {
        result: 'ERROR',
        flag: 'UNKNOWN_WAYBILL',
        message: MSG.UNKNOWN_WB,
        logMessage: MSG.UNKNOWN_WB,
        sku: code.toUpperCase(),
        qty: 0,
        stage: stage,
        waybill: waybillNo,
        details: code ? code + ' @ ' + stage : stage
      };
    }

    var unknownText = unknownItemMessage(view);
    if (unknownText && stage === 'PICK') {
      return errorDecision(view, stage, 'UNKNOWN_ITEM', unknownText, '', 0, unknownText);
    }

    if (view.pickedUp && action !== 'SHORTAGE') {
      return {
        result: 'ERROR',
        flag: '',
        message: MSG.PICKED_UP,
        logMessage: MSG.PICKED_UP,
        sku: '',
        qty: 0,
        stage: stage,
        waybill: view.waybillNo
      };
    }

    if (action === 'DISPATCH' || stage === 'DISPATCH') {
      if (view.dispatched) {
        return errorDecision(view, 'DISPATCH', '', MSG.ALREADY_DISPATCHED, '', 0);
      }
      // Block unless both stages are fully scanned. pickComplete and
      // packComplete come from OK rows in SCANS, not from the phone's memory.
      if (!view.pickComplete || !view.packComplete) {
        if (!view.pickComplete && !view.pickClosed) {
          return errorDecision(
            view,
            'DISPATCH',
            'NOT_PICKED',
            'Not picked yet. Finish pick before dispatch.',
            '',
            0,
            'No one has picked this waybill'
          );
        }
        if (view.pickComplete && !view.packComplete) {
          return errorDecision(
            view,
            'DISPATCH',
            'NOT_PACKED',
            'Not packed yet',
            '',
            0,
            'Picked, not packed yet'
          );
        }
        return errorDecision(
          view,
          'DISPATCH',
          'DISPATCH_WITHOUT_PACK',
          missingStageMessage(view),
          '',
          0,
          'Pick complete: ' + (view.pickComplete ? 'yes' : 'no') + ', pack complete: ' + (view.packComplete ? 'yes' : 'no')
        );
      }
      var mismatch = mismatchText(view);
      return {
        result: 'OK',
        flag: mismatch ? 'STAGE_MISMATCH' : '',
        message: view.waybillNo + ' added to batch',
        logMessage: 'BATCH',
        sku: '',
        qty: 1,
        stage: 'DISPATCH',
        waybill: view.waybillNo,
        details: mismatch
      };
    }

    if (action === 'OPEN' && stage === 'PACK') {
      if (!view.pickComplete) {
        return {
          result: 'ERROR',
          flag: 'PACK_WITHOUT_PICK',
          message: MSG.NOT_PICKED,
          logMessage: 'PACK_WITHOUT_PICK',
          sku: '',
          qty: 0,
          stage: 'PACK',
          waybill: view.waybillNo,
          details: 'Opened in PACK with no completed PICK'
        };
      }
      return {
        result: 'OK',
        flag: '',
        message: MSG.OK,
        logMessage: 'OPEN',
        sku: '',
        qty: 0,
        stage: 'PACK',
        waybill: view.waybillNo,
        details: ''
      };
    }

    if (!view.lines || !view.lines.length) {
      return errorDecision(view, stage, 'WRONG_SKU', MSG.NO_ITEMS, code.toUpperCase(), 0, 'No mapped items');
    }

    if (action === 'SHORTAGE') {
      var closed = stage === 'PACK' ? view.packClosed : view.pickClosed;
      if (closed && !(stage === 'PACK' ? view.packComplete : view.pickComplete)) {
        return errorDecision(view, stage, '', MSG.ALREADY_SHORT, '', 0);
      }
      var missing = shortText(view, stage);
      if (!missing) return errorDecision(view, stage, '', MSG.NO_SHORTAGE, '', 0);
      return {
        result: 'OK',
        flag: 'QTY_SHORT',
        message: MSG.SHORTAGE_OK,
        logMessage: 'SHORTAGE_SUBMIT',
        sku: '',
        qty: 0,
        stage: stage,
        waybill: view.waybillNo,
        details: stage + ' short: ' + missing
      };
    }

    var line = lineOnView(view, code);
    if (!line) {
      if (looksLikeWaybill(code, context)) {
        return errorDecision(view, stage, '', MSG.WAYBILL_AS_ITEM, '', 0, '');
      }
      var known = codes && codes[String(code).trim().toUpperCase()];
      if (known) {
        return errorDecision(view, stage, 'WRONG_SKU', MSG.WRONG_SKU, known.sku, 0, known.sku + ' @ ' + stage);
      }
      return errorDecision(view, stage, 'UNKNOWN_CODE', MSG.UNKNOWN_CODE, code.toUpperCase(), 0, 'Unknown code ' + code.toUpperCase() + ' @ ' + stage);
    }

    var have = stageQty(line, stage);
    if (have + 1 > line.expected) {
      return errorDecision(view, stage, 'QTY_OVER', MSG.QTY_OVER, line.sku, 0, line.sku + ' @ ' + stage);
    }

    return {
      result: 'OK',
      flag: '',
      message: MSG.OK,
      logMessage: '',
      sku: line.sku,
      qty: 1,
      stage: stage,
      waybill: view.waybillNo,
      details: ''
    };
  }

  function errorDecision(view, stage, flag, message, sku, qty, details) {
    return {
      result: 'ERROR',
      flag: flag || '',
      message: message,
      logMessage: message,
      sku: sku || '',
      qty: qty || 0,
      stage: stage,
      waybill: view ? view.waybillNo : '',
      details: details || message
    };
  }

  function applyDecisionToView(view, decision) {
    var next = {
      waybillNo: view.waybillNo,
      client: view.client,
      clientLabel: view.clientLabel,
      itemString: view.itemString,
      orderDate: view.orderDate,
      jtStatus: view.jtStatus,
      pickupTime: view.pickupTime,
      pickedUp: view.pickedUp,
      lines: (view.lines || []).map(function (line) {
        return {
          sku: line.sku,
          name: line.name,
          barcode: line.barcode,
          unknown: !!line.unknown,
          expected: line.expected,
          picked: line.picked,
          packed: line.packed
        };
      }),
      pickComplete: view.pickComplete,
      packComplete: view.packComplete,
      pickClosed: view.pickClosed,
      packClosed: view.packClosed,
      dispatched: view.dispatched
    };
    if (!decision || decision.result !== 'OK') return next;
    if (decision.logMessage === 'SHORTAGE_SUBMIT') {
      if (decision.stage === 'PACK') next.packClosed = true;
      else next.pickClosed = true;
      return next;
    }
    if (decision.stage === 'DISPATCH') {
      next.dispatched = true;
      return next;
    }
    next.lines.forEach(function (line) {
      if (line.unknown) return;
      if (normalizeSku(line.sku) !== normalizeSku(decision.sku)) return;
      if (decision.stage === 'PACK') line.packed += decision.qty;
      else line.picked += decision.qty;
    });
    var known = next.lines.filter(function (line) { return !line.unknown; });
    next.pickComplete = known.length > 0 && next.lines.every(function (line) { return !line.unknown && line.picked >= line.expected; });
    next.packComplete = known.length > 0 && next.lines.every(function (line) { return !line.unknown && line.packed >= line.expected; });
    if (next.pickComplete) next.pickClosed = true;
    if (next.packComplete) next.packClosed = true;
    return next;
  }

  function scanIdSet(scans) {
    var ids = {};
    (scans || []).forEach(function (scan) {
      var id = messageParts(scan.message).id;
      if (id) ids[id] = scan;
    });
    return ids;
  }

  function normalizeItemKey(text) {
    return String(text || '').split(',').map(function (part) {
      return part.trim().toLowerCase();
    }).filter(Boolean).sort().join('|');
  }

  function flagItemKey(flag) {
    if (!flag) return '';
    if (flag.itemKey) return normalizeItemKey(String(flag.itemKey).split('|').join(','));
    var details = String(flag.details || '');
    var match = details.match(/unmapped:\s*([^;]+)/i);
    if (!match) return '';
    return normalizeItemKey(match[1]);
  }

  function pushFlag(bucket, existing, flag, mode) {
    var pool = (existing || []).concat(bucket);
    var wantType = String(flag.type || '').toUpperCase();
    var duplicate = pool.some(function (row) {
      if (normWaybill(row.waybill) !== normWaybill(flag.waybill)) return false;
      if (String(row.type || '').toUpperCase() !== wantType) return false;
      if (mode === 'type') return true;
      if (mode === 'item') {
        if (isResolved(row)) return false;
        return flagItemKey(row) === flagItemKey(flag);
      }
      if (mode === 'person') {
        if (isResolved(row)) return false;
        if (staffEmailOfFlag(row) !== staffEmailOfFlag(flag)) return false;
        return samePersonStageOf(row) === samePersonStageOf(flag);
      }
      if (mode === 'staff') {
        if (isResolved(row)) return false;
        return staffEmailOfFlag(row) === staffEmailOfFlag(flag);
      }
      if (String(row.details || '') !== String(flag.details || '')) return false;
      if (mode === 'ever') return true;
      return !isResolved(row);
    });
    if (!duplicate) bucket.push(flag);
  }

  function formatStaff(staff) {
    if (!staff) return '';
    if (typeof staff === 'string') return staff;
    var name = staff.name || '';
    var email = normalizeEmail(staff.email);
    if (name && email) return name + ' (' + email + ')';
    return name || email;
  }

  function makeFlag(now, waybill, client, type, details, staff) {
    return {
      timestamp: formatManila(now),
      waybill: normWaybill(waybill),
      client: normalizeClient(client),
      type: type,
      details: details || '',
      staff: formatStaff(staff),
      resolved: false
    };
  }

  /**
   * Apply a batch of phone events onto a model. Append-only.
   * The same localId is ignored on retry so a flaky phone cannot double-count.
   * Ledger is never touched here.
   */
  function applyScans(model, staff, events, now, config) {
    var moment = now || new Date();
    var conf = cfg(config);
    var applied = [];
    var flags = [];
    var knownIds = scanIdSet(model.scans);
    var touched = {};

    (events || []).forEach(function (event) {
      var localId = String((event && event.localId) || '').trim();
      if (localId && knownIds[localId]) {
        var prior = knownIds[localId];
        applied.push({
          localId: localId,
          result: String(prior.result || '').toUpperCase() === 'OK' ? 'OK' : 'ERROR',
          message: messageParts(prior.message).body || prior.result,
          sku: prior.sku || '',
          duplicate: true,
          flag: ''
        });
        touched[normWaybill(prior.waybill)] = true;
        return;
      }

      var waybillNo = normWaybill(event.waybill || event.code);
      var record = findWaybill(model, waybillNo);
      var view = record ? buildWaybillView(record, model.items, model.scans, conf) : null;
      var ruleSource = (conf && conf.samePersonStages != null && conf.samePersonStages !== '')
        ? conf
        : ((model.settings && model.settings.samePersonStages != null) ? model.settings : conf);
      var decision = decide(view, event, buildCodeIndex(model.items), {
        staff: staff,
        scans: model.scans,
        samePersonStages: samePersonEnabled(ruleSource),
        waybillNos: waybillIndex(model.waybills)
      });
      if (decision.result === 'BLOCKED' && blockedWithin(model.scans, staff, decision.waybill || waybillNo, decision.stage, moment)) {
        applied.push({
          localId: localId,
          result: 'BLOCKED',
          message: decision.message,
          sku: '',
          stage: decision.stage,
          waybill: decision.waybill || waybillNo,
          flag: '',
          duplicate: true
        });
        return;
      }
      var scan = {
        timestamp: formatManila(moment),
        staffEmail: normalizeEmail(staff && staff.email),
        staffName: (staff && staff.name) || '',
        stage: decision.stage || String(event.stage || '').toUpperCase(),
        waybill: normWaybill(clipScanText(decision.waybill || event.waybill || event.code)),
        sku: clipScanText(decision.sku || ''),
        qty: decision.qty || 0,
        result: decision.result,
        message: withId(localId, decision.logMessage || '')
      };
      model.scans.push(scan);
      if (localId) knownIds[localId] = scan;
      touched[scan.waybill] = true;

      if (decision.flag) {
        var flagRow = makeFlag(
          moment,
          scan.waybill,
          view ? view.client : '',
          decision.flag,
          decision.details || decision.message,
          staff
        );
        // Every failed scan gets its own FLAGS row. An older open flag for
        // the same waybill must not swallow the next person's scan.
        // A same-person block keeps one open flag per waybill, stage, and staff.
        var onceWhileOpen = decision.flag === 'PACK_WITHOUT_PICK' || decision.flag === 'DISPATCH_WITHOUT_PACK';
        if (decision.flag === 'SAME_PERSON_STAGE') pushFlag(flags, model.flags, flagRow, 'person');
        else if (decision.flag === 'NOT_PACKED' || decision.flag === 'NOT_PICKED') pushFlag(flags, model.flags, flagRow, 'staff');
        else if ((decision.result === 'ERROR' || decision.result === 'BLOCKED') && !onceWhileOpen) flags.push(flagRow);
        else pushFlag(flags, model.flags, flagRow, decision.flag === 'STAGE_MISMATCH' ? 'type' : 'open');
      }

      if (decision.result === 'OK' && decision.stage === 'PACK') {
        var after = buildWaybillView(findWaybill(model, scan.waybill), model.items, model.scans, conf);
        var neverPicked = after && !after.pickClosed && (after.lines || []).every(function (line) {
          return !line.picked;
        });
        if (after && after.packClosed && !neverPicked) {
          var diff = mismatchText(after);
          if (diff) {
            pushFlag(flags, model.flags, makeFlag(
              moment,
              after.waybillNo,
              after.client,
              'STAGE_MISMATCH',
              diff,
              staff
            ), 'type');
          }
        }
      }

      applied.push({
        localId: localId,
        result: decision.result,
        message: decision.message,
        sku: decision.sku,
        stage: decision.stage || scan.stage,
        waybill: decision.waybill || scan.waybill,
        flag: decision.flag || '',
        duplicate: false
      });
    });

    flags.forEach(function (flag) { model.flags.push(flag); });

    var views = [];
    Object.keys(touched).forEach(function (waybillNo) {
      var record = findWaybill(model, waybillNo);
      if (!record) return;
      views.push(buildWaybillView(record, model.items, model.scans, conf));
    });

    return { applied: applied, flags: flags, views: views };
  }

  function mapImportHeaders(headerRow, config) {
    var conf = cfg(config);
    var normalized = (headerRow || []).map(function (cell) {
      return String(cell == null ? '' : cell).trim().toLowerCase();
    });
    var mapping = {};
    Object.keys(conf.columns || {}).forEach(function (field) {
      var aliases = (conf.columns[field] || []).map(function (alias) {
        return String(alias).trim().toLowerCase();
      });
      var found = -1;
      for (var i = 0; i < normalized.length; i++) {
        if (aliases.indexOf(normalized[i]) !== -1) {
          found = i;
          break;
        }
      }
      mapping[field] = found;
    });
    var missing = [];
    if (mapping.waybillNo == null || mapping.waybillNo < 0) missing.push('Waybill No');
    if (mapping.itemText == null || mapping.itemText < 0) missing.push('Item / Remarks');
    var warnings = [];
    if ((mapping.jtStatus == null || mapping.jtStatus < 0) && (mapping.pickupTime == null || mapping.pickupTime < 0)) {
      warnings.push('No status or pickup column found. Nothing will be marked picked up.');
    }
    return { mapping: mapping, missing: missing, warnings: warnings, headers: headerRow || [] };
  }

  function rowToImport(row, mapping) {
    var obj = {};
    Object.keys(mapping || {}).forEach(function (field) {
      var index = mapping[field];
      obj[field] = index >= 0 && row ? row[index] : '';
    });
    return obj;
  }

  function rowsFromImportSheet(values, config) {
    var conf = cfg(config);
    var headerIndex = (conf.headerRow || 1) - 1;
    if (!values || values.length <= headerIndex) {
      return { rows: [], missing: ['Waybill No', 'Item / Remarks'], warnings: ['IMPORT tab is empty.'] };
    }
    var mapped = mapImportHeaders(values[headerIndex], conf);
    if (mapped.missing.length) return { rows: [], missing: mapped.missing, warnings: mapped.warnings, headers: mapped.headers };
    var rows = [];
    for (var i = headerIndex + 1; i < values.length; i++) {
      var row = values[i] || [];
      var hasAny = row.some(function (cell) { return String(cell == null ? '' : cell).trim() !== ''; });
      if (!hasAny) continue;
      rows.push(rowToImport(row, mapped.mapping));
    }
    return { rows: rows, missing: [], warnings: mapped.warnings, headers: values[headerIndex] };
  }

  function ledgerHasShipOut(ledger, waybillNo) {
    var want = normWaybill(waybillNo);
    return (ledger || []).some(function (row) {
      return String(row.type || '').toUpperCase() === 'SHIP_OUT' && normWaybill(row.ref) === want;
    });
  }

  function waybillIndex(rows) {
    var set = {};
    (rows || []).forEach(function (row) {
      var no = normWaybill(row && row.waybillNo);
      if (no) set[no] = true;
    });
    return set;
  }

  function planAutoFlags(model, config, now, options) {
    var conf = cfg(config);
    var moment = now || new Date();
    var hours = Number(conf.staleDispatchHours);
    if (!hours && hours !== 0) hours = 24;
    var fresh = (options && options.notInWaybills) || {};
    var grouped = groupScans(model.scans);
    var flags = [];
    (model.waybills || []).forEach(function (waybill) {
      var subset = scansFor(grouped, waybill.waybillNo);
      var view = buildWaybillView(waybill, model.items, subset, conf);
      if (view.pickedUp && !view.dispatched) {
        var summaries = (options && options.summaries) || {};
        var reason = 'J&T picked up but no dispatch scan';
        var stockNote = summaries[view.waybillNo]
          ? summaries[view.waybillNo]
          : (fresh[view.waybillNo] ? 'Waybill was not in WAYBILLS.' : '');
        var details = stockNote ? reason + '. ' + stockNote : reason;
        pushFlag(flags, model.flags, makeFlag(
          moment,
          view.waybillNo,
          view.client,
          'PICKED_UP_NOT_DISPATCHED',
          details,
          'SYSTEM'
        ), 'type');
      }
      if (!view.pickedUp && view.dispatched) {
        var sent = dispatchInstant(subset, view.waybillNo);
        if (sent && (moment.getTime() - sent.getTime()) / 36e5 >= hours) {
          pushFlag(flags, model.flags, makeFlag(
            moment,
            view.waybillNo,
            view.client,
            'DISPATCHED_NOT_PICKED_UP',
            'Dispatched, no J&T pickup after ' + hours + 'h',
            'SYSTEM'
          ), 'type');
        }
      }
    });
    return flags;
  }

  function dispatchInstant(scans, waybillNo) {
    var earliest = null;
    (scans || []).forEach(function (scan) {
      if (normWaybill(scan.waybill) !== normWaybill(waybillNo)) return;
      if (String(scan.stage || '').toUpperCase() !== 'DISPATCH') return;
      if (String(scan.result || '').toUpperCase() !== 'OK') return;
      var when = parseManila(scan.timestamp);
      if (!when) return;
      if (!earliest || when < earliest) earliest = when;
    });
    return earliest;
  }

  function describeStockPost(items, unknown, notInSheet) {
    var stock = (items && items.length)
      ? 'Stock deducted: ' + items.map(function (item) {
        return item.sku + ' x' + Math.abs(item.qty);
      }).join(', ')
      : 'Stock deducted: none';
    var bits = [stock];
    if (unknown && unknown.length) bits.push('unmapped: ' + unknown.join(', '));
    if (notInSheet) bits.push('not in WAYBILLS');
    return bits.join('; ');
  }

  var JT_CREATORS = {
    'NEL-F2003-0048': 'TREXO',
    'NEL-F2003-0031': 'SIRWEN',
    'NEL-F2003-0014': 'ECOM'
  };

  function jtCreatorClient(code) {
    return JT_CREATORS[String(code || '').trim().toUpperCase()] || '';
  }

  function jtField(row, header, fallback) {
    var src = row || {};
    if (src[header] != null && String(src[header]).trim() !== '') return String(src[header]).trim();
    if (fallback && src[fallback] != null && String(src[fallback]).trim() !== '') return String(src[fallback]).trim();
    return '';
  }

  function jtExportView(row) {
    return {
      creatorCode: jtField(row, 'Creator Code', 'creatorCode'),
      waybillNo: normWaybill(jtField(row, 'Waybill Number', 'waybillNo')),
      orderStatus: jtField(row, 'Order Status', 'orderStatus'),
      signingTime: jtField(row, 'SigningTime', 'signingTime'),
      submissionTime: jtField(row, 'Submission Time', 'submissionTime'),
      rtsReason: jtField(row, 'RTS Reason', 'rtsReason'),
      itemName: jtField(row, 'Item Name', 'itemName')
    };
  }

  function parseJtItemName(text) {
    var raw = String(text || '').trim();
    if (!raw) return [];
    var parts = raw.split('+');
    var out = [];
    parts.forEach(function (part) {
      var token = String(part || '').trim();
      if (!token) return;
      var match = token.match(/^(\d+)\s*x\s*(.+)$/i);
      if (!match) {
        out.push({ name: token, qty: 1 });
        return;
      }
      out.push({ name: String(match[2]).trim(), qty: clampQty(match[1]) });
    });
    return out.filter(function (part) { return part.name; });
  }

  function aliasList(value) {
    return String(value || '').split(',').map(function (part) {
      return part.trim().toLowerCase();
    }).filter(Boolean);
  }

  function findJtItem(items, name, client) {
    var want = String(name || '').trim().toLowerCase();
    var hint = normalizeClient(client);
    if (!want) return null;
    var rows = (items || []).filter(function (item) {
      return isActiveItem(item) && normalizeClient(item.client) === hint;
    });
    for (var i = 0; i < rows.length; i++) {
      if (aliasList(rows[i].jtAlias).indexOf(want) !== -1) return rows[i];
    }
    var sku = normalizeSku(name);
    for (var j = 0; j < rows.length; j++) {
      if (normalizeSku(rows[j].sku) === sku) return rows[j];
    }
    return null;
  }

  function mapJtNames(parts, client, items) {
    var resolved = (parts || []).map(function (part) {
      return resolveSegment(part.name, part.qty, client, items);
    });
    var unknown = [];
    var seen = {};
    resolved.forEach(function (row) {
      if (!row.unknown) return;
      var key = normalizeSku(row.name);
      if (!key || seen[key]) return;
      seen[key] = true;
      unknown.push(row.name);
    });
    return {
      items: mergeQty(resolved.filter(function (row) { return !row.unknown; })),
      unknown: unknown,
      segments: resolved,
      itemString: formatSegments(resolved)
    };
  }

  function waybillHasScan(scans, waybillNo) {
    var want = normWaybill(waybillNo);
    return (scans || []).some(function (scan) {
      return normWaybill(scan.waybill) === want && String(scan.result || '').toUpperCase() === 'OK';
    });
  }

  /**
   * J&T My Waybill export. Keeps three creator codes, never stores personal
   * columns, and does not deduct stock. Re-import updates status fields only,
   * unless the waybill is still unmapped and nobody has scanned it.
   */
  function planJtExport(rows, model, now) {
    var moment = now || new Date();
    var inserts = [];
    var updates = [];
    var flags = [];
    var byClient = { TREXO: 0, SIRWEN: 0, ECOM: 0 };
    var unknownCounts = {};
    var dropped = 0;
    var seen = {};
    var fresh = {};
    var noted = {};
    (model.waybills || []).forEach(function (row) {
      seen[normWaybill(row.waybillNo)] = row;
    });

    function noteUnknown(client, waybillNo, names) {
      var text = (names || []).filter(Boolean).join(', ');
      if (!text || noted[waybillNo]) return;
      noted[waybillNo] = true;
      text.split(',').forEach(function (name) {
        var clean = name.trim();
        if (!clean) return;
        var key = client + '|' + clean.toLowerCase();
        unknownCounts[key] = unknownCounts[key] || { client: client, name: clean, count: 0 };
        unknownCounts[key].count += 1;
      });
      var unknownFlag = makeFlag(moment, waybillNo, client, 'UNKNOWN_ITEM', 'Unmapped: ' + text, 'IMPORT');
      unknownFlag.itemKey = normalizeItemKey(text);
      pushFlag(flags, model.flags, unknownFlag, 'item');
    }

    (rows || []).forEach(function (source) {
      var row = jtExportView(source);
      var client = jtCreatorClient(row.creatorCode);
      if (!client || !row.waybillNo) {
        dropped += 1;
        return;
      }
      byClient[client] += 1;
      var parsed = mapJtNames(parseJtItemName(row.itemName), client, model.items);
      var existing = seen[row.waybillNo];
      if (!existing) {
        var created = {
          waybillNo: row.waybillNo,
          client: client,
          itemString: parsed.itemString || '',
          rawItems: row.itemName,
          orderDate: asDateOnly(row.submissionTime),
          importedAt: formatManila(moment),
          jtStatus: row.orderStatus,
          pickupTime: '',
          signingTime: row.signingTime,
          rtsReason: row.rtsReason,
          unmapped: parsed.unknown.join(', ')
        };
        inserts.push(created);
        seen[row.waybillNo] = created;
        fresh[row.waybillNo] = created;
        noteUnknown(client, row.waybillNo, parsed.unknown);
        return;
      }
      if (fresh[row.waybillNo]) {
        existing.jtStatus = row.orderStatus || existing.jtStatus;
        existing.signingTime = row.signingTime || existing.signingTime;
        existing.rtsReason = row.rtsReason || existing.rtsReason;
        return;
      }
      var prior = null;
      for (var u = 0; u < updates.length; u++) {
        if (updates[u].waybillNo === row.waybillNo) prior = updates[u];
      }
      if (prior) {
        prior.jtStatus = row.orderStatus || prior.jtStatus;
        prior.signingTime = row.signingTime || prior.signingTime;
        prior.rtsReason = row.rtsReason || prior.rtsReason;
        return;
      }
      var scanned = waybillHasScan(model.scans, row.waybillNo);
      var remap = !scanned && String(existing.unmapped || '').trim();
      var update = {
        _row: existing._row,
        waybillNo: row.waybillNo,
        client: existing.client,
        jtStatus: row.orderStatus,
        signingTime: row.signingTime,
        rtsReason: row.rtsReason,
        pickupTime: existing.pickupTime || '',
        itemString: existing.itemString || '',
        rawItems: existing.rawItems || '',
        unmapped: existing.unmapped || '',
        remap: !!remap
      };
      if (remap) {
        update.itemString = parsed.itemString || existing.itemString || '';
        update.rawItems = row.itemName || existing.rawItems || '';
        update.unmapped = parsed.unknown.join(', ');
      }
      updates.push(update);
      if (String(update.unmapped || '').trim()) noteUnknown(client, row.waybillNo, String(update.unmapped).split(','));
    });

    var unknownItems = Object.keys(unknownCounts).map(function (key) { return unknownCounts[key]; });
    unknownItems.sort(function (a, b) { return b.count - a.count || (a.name < b.name ? -1 : 1); });
    return {
      dropped: dropped,
      byClient: byClient,
      inserts: inserts,
      updates: updates,
      flags: flags,
      ledger: [],
      unknownItems: unknownItems,
      preview: {
        dropped: dropped,
        byClient: byClient,
        newWaybills: inserts.length,
        updated: updates.length,
        unknownItems: unknownItems,
        flags: flags.length
      }
    };
  }

  /**
   * Plan sheet changes for one J&T paste. Does not deduct twice for the same waybill.
   * Mapped SKUs are deducted even when part of the text is unknown.
   */
  function planJtImport(importRows, model, config, now) {
    var conf = cfg(config);
    var moment = now || new Date();
    var inserts = [];
    var updates = [];
    var ledger = [];
    var flags = [];
    var warnings = [];
    var posted = {};
    (model.ledger || []).forEach(function (row) {
      if (String(row.type || '').toUpperCase() === 'SHIP_OUT') posted[normWaybill(row.ref)] = true;
    });

    (importRows || []).forEach(function (row, index) {
      var waybillNo = normWaybill(row.waybillNo);
      var itemText = String(row.itemText == null ? '' : row.itemText).trim();
      if (!waybillNo) {
        warnings.push('Row ' + (index + 2) + ' has no waybill number.');
        return;
      }
      var resolved = resolveLineItems(itemText, row.client, model.skuMap, model.items);
      var pending = null;
      for (var p = 0; p < inserts.length; p++) {
        if (inserts[p].waybillNo === waybillNo) pending = inserts[p];
      }
      var existing = pending || findWaybill(model, waybillNo);
      var itemString = resolved.itemString
        || (existing && existing.itemString ? existing.itemString : itemText);

      var record = {
        waybillNo: waybillNo,
        client: resolved.client || (existing && existing.client) || '',
        itemString: itemString,
        rawItems: itemText || (existing && existing.rawItems) || '',
        orderDate: asDateOnly(row.orderDate) || (existing && existing.orderDate) || '',
        importedAt: existing && existing.importedAt ? existing.importedAt : formatManila(moment),
        jtStatus: String(row.jtStatus == null ? '' : row.jtStatus).trim() || (existing && existing.jtStatus) || '',
        pickupTime: hasPickupTime(row.pickupTime) ? asManilaString(row.pickupTime) : ((existing && existing.pickupTime) || ''),
        unmapped: resolved.unknown.join(', ')
      };

      var notInSheet = !findWaybill(model, waybillNo);
      if (pending) {
        record.importedAt = pending.importedAt;
        Object.assign(pending, record);
      } else if (existing) {
        updates.push(copyWaybill(record, existing._row));
      } else {
        inserts.push(record);
        record._notInWaybills = notInSheet;
      }

      var picked = isPickedUp(record.jtStatus, record.pickupTime, conf);
      var postedThis = [];
      if (picked && !posted[waybillNo] && resolved.items.length) {
        resolved.items.forEach(function (item) {
          var line = {
            date: formatManila(moment),
            client: item.client || record.client,
            sku: item.sku,
            type: 'SHIP_OUT',
            qty: -Math.abs(item.qty),
            ref: waybillNo,
            source: 'JT_IMPORT'
          };
          ledger.push(line);
          postedThis.push({ sku: item.sku, qty: Math.abs(item.qty) });
        });
        posted[waybillNo] = true;
      }
      var summary = describeStockPost(postedThis, resolved.unknown, notInSheet && picked);
      record._importSummary = summary;
      if (pending) pending._importSummary = summary;
      if (existing && existing._row) {
        var updated = updates[updates.length - 1];
        if (updated && updated.waybillNo === waybillNo) updated._importSummary = summary;
      }

      if (resolved.unknown.length) {
        var unknownFlag = makeFlag(
          moment,
          waybillNo,
          record.client,
          'UNKNOWN_ITEM',
          summary,
          'IMPORT'
        );
        unknownFlag.itemKey = normalizeItemKey(resolved.unknown.join(','));
        pushFlag(flags, model.flags, unknownFlag, 'item');
      }
      if (resolved.unknown.length || postedThis.length) {
        warnings.push(waybillNo + ': ' + summary);
      }
    });

    var overlay = {
      items: model.items,
      skuMap: model.skuMap,
      waybills: (model.waybills || []).map(function (row) { return Object.assign({}, row); }).concat(inserts.map(function (row) {
        return Object.assign({}, row);
      })),
      scans: model.scans,
      flags: (model.flags || []).concat(flags),
      ledger: (model.ledger || []).concat(ledger)
    };
    updates.forEach(function (update) {
      overlay.waybills.forEach(function (row) {
        if (normWaybill(row.waybillNo) === update.waybillNo) Object.assign(row, update);
      });
    });
    var notInWaybills = {};
    var summaries = {};
    inserts.concat(updates).forEach(function (row) {
      if (row._notInWaybills) notInWaybills[row.waybillNo] = true;
      if (row._importSummary) summaries[row.waybillNo] = row._importSummary;
    });
    planAutoFlags(overlay, conf, moment, { notInWaybills: notInWaybills, summaries: summaries }).forEach(function (flag) {
      pushFlag(flags, model.flags, flag, 'type');
    });

    return {
      inserts: inserts,
      updates: updates,
      ledger: ledger,
      flags: flags,
      warnings: warnings
    };
  }

  function copyWaybill(record, rowNumber) {
    return {
      _row: rowNumber,
      waybillNo: record.waybillNo,
      client: record.client,
      itemString: record.itemString,
      rawItems: record.rawItems || '',
      orderDate: record.orderDate,
      importedAt: record.importedAt,
      jtStatus: record.jtStatus,
      pickupTime: record.pickupTime,
      signingTime: record.signingTime || '',
      rtsReason: record.rtsReason || '',
      unmapped: record.unmapped || ''
    };
  }

  function stageActorMap(scans) {
    var map = {};
    (scans || []).forEach(function (scan) {
      if (String(scan.result || '').toUpperCase() !== 'OK') return;
      var stage = String(scan.stage || '').toUpperCase();
      if (stage !== 'PICK' && stage !== 'PACK' && stage !== 'DISPATCH') return;
      var no = normWaybill(scan.waybill);
      if (!no) return;
      if (!map[no]) map[no] = {};
      if (!map[no][stage]) {
        map[no][stage] = {
          name: scan.staffName || scan.staffEmail || '',
          email: normalizeEmail(scan.staffEmail)
        };
      }
    });
    return map;
  }

  function stageIsComplete(view, stage) {
    if (!view) return false;
    if (stage === 'PICK') return !!view.pickComplete;
    if (stage === 'PACK') return !!view.packComplete;
    if (stage === 'DISPATCH') return !!view.dispatched;
    return false;
  }

  function countStageToday(views, scans, stage, isToday) {
    var name = String(stage || '').toUpperCase();
    var last = {};
    var short = {};
    (scans || []).forEach(function (scan) {
      if (String(scan.stage || '').toUpperCase() !== name) return;
      if (String(scan.result || '').toUpperCase() !== 'OK') return;
      var no = normWaybill(scan.waybill);
      var ts = asManilaString(scan.timestamp);
      if (String(messageParts(scan.message).body).toUpperCase() === 'SHORTAGE_SUBMIT') {
        if (isToday(ts)) short[no] = true;
        return;
      }
      if (!last[no] || ts > last[no]) last[no] = ts;
    });
    var done = 0;
    var shortCount = 0;
    (views || []).forEach(function (view) {
      if (stageIsComplete(view, name)) {
        var ts = last[view.waybillNo];
        if (ts && isToday(ts)) done += 1;
        return;
      }
      if (short[view.waybillNo]) shortCount += 1;
    });
    return { done: done, short: shortCount };
  }

  function blockedWithin(scans, staff, waybill, stage, now) {
    var email = normalizeEmail(staff && staff.email);
    var want = normWaybill(waybill);
    var stageName = String(stage || '').toUpperCase();
    var moment = (now || new Date()).getTime();
    return (scans || []).some(function (scan) {
      if (String(scan.result || '').toUpperCase() !== 'BLOCKED') return false;
      if (normalizeEmail(scan.staffEmail) !== email) return false;
      if (normWaybill(scan.waybill) !== want) return false;
      if (String(scan.stage || '').toUpperCase() !== stageName) return false;
      var when = parseManila(scan.timestamp);
      if (!when) return false;
      var age = moment - when.getTime();
      return age >= 0 && age < 60000;
    });
  }

  function staffEmailOfFlag(flag) {
    var text = String(flag && flag.staff || '');
    var match = text.match(/\(([^)]+)\)\s*$/);
    return normalizeEmail(match ? match[1] : text);
  }

  function samePersonStageOf(flag) {
    var text = String(flag && (flag.details || '') || '');
    if (/must pack/i.test(text)) return 'PACK';
    if (/must dispatch/i.test(text)) return 'DISPATCH';
    if (/must pick/i.test(text)) return 'PICK';
    return '';
  }

  function nextBanner(current, next, now, holdUntil, force) {
    var kind = current && current.kind;
    var held = kind === 'bad' || kind === 'warn';
    if (!force && held && holdUntil && now < holdUntil) {
      return { banner: current, holdUntil: holdUntil, kept: true };
    }
    var banner = next || null;
    var hold = 0;
    if (banner && (banner.kind === 'bad' || banner.kind === 'warn')) hold = now + 4000;
    return { banner: banner, holdUntil: hold, kept: false };
  }

  function shortageAllowed(view, stage) {
    var name = String(stage || '').toUpperCase();
    if (!view || name === 'DISPATCH') return false;
    if (name === 'PACK') return !view.packComplete;
    if (name === 'PICK') return !view.pickComplete;
    return false;
  }

  function stageOwnerRows(scans, pickComplete) {
    var done = pickComplete || null;
    var map = {};
    (scans || []).forEach(function (scan) {
      if (String(scan.result || '').toUpperCase() !== 'OK') return;
      var stage = String(scan.stage || '').toUpperCase();
      if (stage !== 'PICK' && stage !== 'PACK' && stage !== 'DISPATCH') return;
      var no = normWaybill(scan.waybill);
      if (!no) return;
      if (stage === 'PICK' && done && !done[no]) return;
      if (!map[no]) map[no] = { waybill: no, PICK: '', PACK: '', DISPATCH: '' };
      var name = scan.staffName || scan.staffEmail || '';
      if (!name) return;
      if (!map[no][stage]) map[no][stage] = name;
      else if (map[no][stage].indexOf(name) === -1) map[no][stage] += ', ' + name;
    });
    return Object.keys(map).sort().map(function (no) {
      var row = map[no];
      return {
        waybill: clipDisplay(row.waybill),
        PICK: clipDisplay(row.PICK),
        PACK: clipDisplay(row.PACK),
        DISPATCH: clipDisplay(row.DISPATCH)
      };
    });
  }

  function buildAdminSummary(model, now) {
    var moment = now || new Date();
    var today = manilaDate(moment);
    function isToday(value) {
      return asManilaString(value).slice(0, 10) === today;
    }
    var summaryViews = (model.waybills || []).map(function (waybill) {
      return buildWaybillView(waybill, model.items, model.scans, null);
    });
    var counts = { PICK: 0, PACK: 0, DISPATCH: 0, PICKUP: 0, PICK_SHORT: 0, PACK_SHORT: 0 };
    ['PICK', 'PACK', 'DISPATCH'].forEach(function (stage) {
      var tally = countStageToday(summaryViews, model.scans, stage, isToday);
      counts[stage] = tally.done;
      if (stage === 'PICK') counts.PICK_SHORT = tally.short;
      if (stage === 'PACK') counts.PACK_SHORT = tally.short;
    });
    var flags = (model.flags || []).filter(function (flag) {
      return !isResolved(flag);
    }).map(function (flag) {
      return {
        timestamp: clipDisplay(asManilaString(flag.timestamp)),
        waybill: clipDisplay(normWaybill(flag.waybill)),
        waybillKey: normWaybill(flag.waybill),
        _row: flag._row || 0,
        client: clipDisplay(normalizeClient(flag.client)),
        clientLabel: clipDisplay(clientLabel(flag.client)),
        type: clipDisplay(String(flag.type || '').toUpperCase()),
        details: clipDisplay(flag.details || ''),
        staff: clipDisplay(flag.staff || '')
      };
    }).sort(function (a, b) {
      return a.timestamp < b.timestamp ? 1 : -1;
    });
    return {
      today: today,
      counts: counts,
      flags: flags,
      backlog: null,
      stageOwners: null,
      staffOutput: null,
      partial: true
    };
  }

  function buildAdminDashboard(model, now, config) {
    var conf = cfg(config);
    var moment = now || new Date();
    var today = manilaDate(moment);
    var grouped = groupScans(model.scans);
    var views = (model.waybills || []).map(function (waybill) {
      return buildWaybillView(waybill, model.items, scansFor(grouped, waybill.waybillNo), conf);
    });

    function isToday(value) {
      var text = asManilaString(value);
      return text.slice(0, 10) === today;
    }

    var counts = { PICK: 0, PACK: 0, DISPATCH: 0, PICKUP: 0, PICK_SHORT: 0, PACK_SHORT: 0 };
    ['PICK', 'PACK', 'DISPATCH'].forEach(function (stage) {
      var tally = countStageToday(views, model.scans, stage, isToday);
      counts[stage] = tally.done;
      if (stage === 'PICK') counts.PICK_SHORT = tally.short;
      if (stage === 'PACK') counts.PACK_SHORT = tally.short;
    });
    views.forEach(function (view) {
      if (view.pickedUp && isToday(view.pickupTime)) counts.PICKUP += 1;
    });

    var buckets = {
      printedNotPicked: [],
      pickedNotPacked: [],
      packedNotDispatched: [],
      dispatchedNotPickedUp: []
    };
    views.forEach(function (view) {
      var shown = clipDisplay(view.waybillNo);
      if (view.pickedUp) return;
      if (view.dispatched) buckets.dispatchedNotPickedUp.push(shown);
      else if (view.packComplete) buckets.packedNotDispatched.push(shown);
      else if (view.pickClosed) buckets.pickedNotPacked.push(shown);
      else buckets.printedNotPicked.push(shown);
    });

    var byStaff = {};
    (model.scans || []).forEach(function (scan) {
      if (!isToday(scan.timestamp)) return;
      if (String(scan.result || '').toUpperCase() !== 'OK') return;
      if (String(messageParts(scan.message).body).toUpperCase() === 'SHORTAGE_SUBMIT') return;
      var name = scan.staffName || scan.staffEmail || 'Unknown';
      if (!byStaff[name]) byStaff[name] = { name: name, PICK: 0, PACK: 0, DISPATCH: 0, units: 0 };
      var stage = String(scan.stage || '').toUpperCase();
      if (byStaff[name][stage] != null) byStaff[name][stage] += 1;
      byStaff[name].units += num(scan.qty);
    });

    var flags = (model.flags || []).filter(function (flag) {
      return !isResolved(flag);
    }).map(function (flag) {
      return {
        timestamp: clipDisplay(asManilaString(flag.timestamp)),
        waybill: clipDisplay(normWaybill(flag.waybill)),
        waybillKey: normWaybill(flag.waybill),
        _row: flag._row || 0,
        client: clipDisplay(normalizeClient(flag.client)),
        clientLabel: clipDisplay(clientLabel(flag.client)),
        type: clipDisplay(String(flag.type || '').toUpperCase()),
        details: clipDisplay(flag.details || ''),
        staff: clipDisplay(flag.staff || '')
      };
    }).sort(function (a, b) {
      return a.timestamp < b.timestamp ? 1 : -1;
    });

    return {
      today: today,
      counts: counts,
      backlog: buckets,
      flags: flags,
      partial: false,
      stageOwners: stageOwnerRows(model.scans, (function () {
        var done = {};
        views.forEach(function (view) {
          if (view.pickComplete) done[view.waybillNo] = true;
        });
        return done;
      })()),
      staffOutput: Object.keys(byStaff).map(function (name) {
        var row = byStaff[name];
        row.name = clipDisplay(row.name);
        return row;
      }).sort(function (a, b) {
        return a.name < b.name ? -1 : 1;
      })
    };
  }

  function onHand(beginningQty, ledgerRows, sku, client) {
    var wantSku = normalizeSku(sku);
    var wantClient = client ? normalizeClient(client) : '';
    var sum = (ledgerRows || []).reduce(function (total, row) {
      if (normalizeSku(row.sku) !== wantSku) return total;
      if (wantClient && normalizeClient(row.client) !== wantClient) return total;
      return total + num(row.qty);
    }, 0);
    return num(beginningQty) + sum;
  }

  function stockFormulasForRow(rowIndex, occurrence) {
    var row = rowIndex;
    var nth = occurrence || 1;
    if (nth > 1) {
      var pick = 'SMALL(INDEX(IF(UPPER(TRIM(ITEMS!A:A))=UPPER(TRIM(A' + row + ')),ROW(ITEMS!A:A)),0),' + nth + ')';
      return {
        client: '=IFERROR(INDEX(ITEMS!B:B,' + pick + '),"")',
        name: '=IFERROR(INDEX(ITEMS!C:C,' + pick + '),"")',
        beginning: '=IFERROR(INDEX(ITEMS!E:E,' + pick + '),0)',
        ledger: '=SUMIFS(LEDGER!E:E,LEDGER!C:C,A' + row + ',LEDGER!B:B,B' + row + ')',
        onHand: '=D' + row + '+E' + row
      };
    }
    return {
      client: '=IFERROR(INDEX(ITEMS!B:B,MATCH(A' + row + ',ITEMS!A:A,0)),"")',
      name: '=IFERROR(INDEX(ITEMS!C:C,MATCH(A' + row + ',ITEMS!A:A,0)),"")',
      beginning: '=IFERROR(INDEX(ITEMS!E:E,MATCH(A' + row + ',ITEMS!A:A,0)),0)',
      ledger: '=SUMIFS(LEDGER!E:E,LEDGER!C:C,A' + row + ',LEDGER!B:B,B' + row + ')',
      onHand: '=D' + row + '+E' + row
    };
  }

  function cellHasContent(value) {
    if (value == null || value === false || value === true) return false;
    var text = String(value).trim();
    if (!text) return false;
    var upper = text.toUpperCase();
    if (upper === 'FALSE' || upper === 'TRUE') return false;
    return true;
  }

  /**
   * Checkbox validation can leave the word FALSE under the real rows.
   * Clear those offsets when the key cell (Email, SKU, or timestamp) is empty.
   */
  function falseOnlySpans(keys, values) {
    var spans = [];
    var start = -1;
    var rows = values || [];
    for (var i = 0; i < rows.length; i++) {
      var key = keys && keys[i] ? keys[i][0] : '';
      var cell = rows[i] ? rows[i][0] : '';
      var emptyKey = !String(key == null ? '' : key).trim();
      var falseCell = cell === false || String(cell == null ? '' : cell).trim().toUpperCase() === 'FALSE';
      if (emptyKey && falseCell) {
        if (start < 0) start = i;
      } else if (start >= 0) {
        spans.push({ start: start, count: i - start });
        start = -1;
      }
    }
    if (start >= 0) spans.push({ start: start, count: rows.length - start });
    return spans;
  }

  function isBlankSheetRow(row) {
    var cells = row || [];
    for (var i = 0; i < cells.length; i++) {
      if (cellHasContent(cells[i])) return false;
    }
    return true;
  }

  /**
   * Last row whose key column has real text.
   * Checkbox FALSE values and data-validation padding do not count,
   * so an append does not land below a block of blank checkbox rows.
   */
  function lastContentRow(sheet, keyColumn) {
    if (!sheet) return 0;
    var col = keyColumn || 1;
    var last = sheet.getLastRow();
    if (last < 1) return 0;
    var values = sheet.getRange(1, col, last, 1).getValues();
    for (var i = values.length - 1; i >= 0; i--) {
      if (cellHasContent(values[i][0])) return i + 1;
    }
    return 0;
  }

  function appendMatrix(sheet, matrix, keyColumn) {
    if (!sheet || !matrix || !matrix.length) return 0;
    var start = lastContentRow(sheet, keyColumn || 1) + 1;
    if (start < 2) start = 2;
    sheet.getRange(start, 1, matrix.length, matrix[0].length).setValues(matrix);
    return start;
  }

  function readTableValues(sheet, columnCount, keyColumn) {
    if (!sheet) return [];
    var cols = columnCount || 1;
    var last = lastContentRow(sheet, keyColumn || 1);
    if (last < 1) return [];
    return sheet.getRange(1, 1, last, cols).getValues();
  }

  /**
   * Pack real rows under the header. Rows that are only checkbox TRUE/FALSE
   * are dropped. Every row with text is kept, in order.
   */
  function compactSheetBody(sheet, columnCount, keyColumn) {
    var cols = columnCount || 1;
    var last = sheet.getLastRow();
    if (last < 2) return { kept: 0, removed: 0 };
    var values = sheet.getRange(1, 1, last, cols).getValues();
    var kept = [];
    for (var i = 1; i < values.length; i++) {
      if (!isBlankSheetRow(values[i])) kept.push(values[i].slice(0, cols));
    }
    replaceSheetBody(sheet, kept, cols);
    return { kept: kept.length, removed: (values.length - 1) - kept.length };
  }

  function columnName(index) {
    var n = index;
    var name = '';
    while (n > 0) {
      var rem = (n - 1) % 26;
      name = String.fromCharCode(65 + rem) + name;
      n = Math.floor((n - 1) / 26);
    }
    return name;
  }

  /**
   * Where to write the next real rows. Checkbox TRUE/FALSE padding is ignored,
   * so a blank block near row 1000 does not push new scans below it.
   */
  function nextSheetRange(sheetName, columnCount, values, addCount) {
    var count = addCount || 0;
    if (!count || count < 1) return '';
    var cols = columnCount || 1;
    var last = 1;
    var rows = values || [];
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i] || [];
      var any = false;
      for (var c = 0; c < cols; c++) {
        if (cellHasContent(row[c])) { any = true; break; }
      }
      if (any) last = i + 1;
    }
    var start = last + 1;
    var end = start + count - 1;
    return sheetName + '!A' + start + ':' + columnName(cols) + end;
  }

  function scratchModel(model) {
    return {
      items: model.items,
      skuMap: model.skuMap,
      waybills: model.waybills,
      scans: (model.scans || []).slice(),
      flags: (model.flags || []).slice(),
      ledger: model.ledger || [],
      staff: model.staff
    };
  }

  /**
   * Turn queued phone events into sheet rows. Recomputes against the latest
   * model so two phones cannot double-count a localId, and a later sync sees
   * scans that already landed.
   */
  function buildSyncBatch(model, staff, events, now, config) {
    var scratch = scratchModel(model || {});
    var items = [];
    (events || []).forEach(function (event) {
      var source = Object.assign({}, event);
      if (!source.localId && event.localId) source.localId = event.localId;
      var scansBefore = scratch.scans.length;
      var flagsBefore = scratch.flags.length;
      var result = applyScans(scratch, staff, [source], now || new Date(), config);
      var applied = (result.applied && result.applied[0]) || {};
      var duplicate = !!applied.duplicate;
      items.push({
        localId: String(source.localId || ''),
        duplicate: duplicate,
        phase: duplicate ? 'done' : 'new',
        scanRow: !duplicate && scratch.scans.length > scansBefore ? scratch.scans[scratch.scans.length - 1] : null,
        flagRow: scratch.flags.length > flagsBefore ? scratch.flags[scratch.flags.length - 1] : null,
        result: applied.result || '',
        message: applied.message || ''
      });
    });
    return items;
  }

  function sameFlagRow(row, flagRow) {
    if (!row || !flagRow) return false;
    return normWaybill(row.waybill) === normWaybill(flagRow.waybill)
      && String(row.type || '').toUpperCase() === String(flagRow.type || '').toUpperCase()
      && String(row.details || '') === String(flagRow.details || '')
      && String(row.staff || '') === String(flagRow.staff || '');
  }

  /**
   * Next writes for one prepared queue item.
   * phase is what to store after those writes succeed.
   * sheetHasScan covers a crash after the scan append was accepted.
   */
  function syncWrites(item, sheetHasScan, existingFlags) {
    if (!item || item.phase === 'done' || item.duplicate) {
      return { scans: [], flags: [], phase: 'done' };
    }
    var flags = [];
    if (item.flagRow && !(existingFlags || []).some(function (row) { return sameFlagRow(row, item.flagRow); })) {
      flags = [item.flagRow];
    }
    if (item.phase === 'scans-sent' || sheetHasScan) {
      return { scans: [], flags: flags, phase: 'done' };
    }
    return { scans: item.scanRow ? [item.scanRow] : [], flags: [], phase: 'scans-sent' };
  }

  function attachSyncRows(model, staff, queue, now, config) {
    var open = [];
    (queue || []).forEach(function (item) {
      if (item && item.phase === 'new' && !item.scanRow && !item.duplicate) open.push(item);
    });
    var events = open.map(function (item) {
      var event = Object.assign({}, item.event);
      event.localId = item.localId;
      return event;
    });
    var built = buildSyncBatch(model, staff, events, now, config);
    var byId = {};
    built.forEach(function (row) { byId[row.localId] = row; });
    return (queue || []).map(function (item) {
      var row = byId[item.localId];
      if (!row) return item;
      return {
        localId: item.localId,
        event: item.event,
        phase: row.phase,
        scanRow: row.scanRow,
        flagRow: row.flagRow,
        duplicate: row.duplicate
      };
    });
  }

  function pendingLabel(count, online, syncing) {
    if (syncing) return 'Syncing…';
    var n = count || 0;
    if (n === 1) return '1 to sync';
    if (n > 1) return n + ' to sync';
    if (!online) return 'Offline';
    return 'Saved';
  }

  function waybillToRow(record) {
    return [
      record.waybillNo,
      record.client,
      record.itemString,
      record.orderDate || '',
      record.importedAt || '',
      record.jtStatus || '',
      record.pickupTime || '',
      record.rawItems || '',
      record.signingTime || '',
      record.rtsReason || '',
      record.unmapped || ''
    ];
  }

  function ledgerToRow(record) {
    return [
      record.date || '',
      record.client || '',
      record.sku || '',
      record.type || '',
      num(record.qty),
      record.ref || '',
      record.source || ''
    ];
  }

  function flagToRow(record) {
    return [
      record.timestamp || '',
      record.waybill || '',
      record.client || '',
      record.type || '',
      record.details || '',
      record.staff || '',
      record.resolved === true
    ];
  }

  function commitJtExport(ss, plan) {
    var sheet = ss.getSheetByName(SHEETS.WAYBILLS.name);
    var fields = SHEETS.WAYBILLS.fields;
    function writeCell(rowNumber, field, value) {
      var col = fields.indexOf(field) + 1;
      if (rowNumber && col > 0) sheet.getRange(rowNumber, col).setValue(value || '');
    }
    (plan.updates || []).forEach(function (update) {
      if (!update._row) return;
      writeCell(update._row, 'jtStatus', update.jtStatus);
      writeCell(update._row, 'signingTime', update.signingTime);
      writeCell(update._row, 'rtsReason', update.rtsReason);
      if (update.remap) {
        writeCell(update._row, 'itemString', update.itemString);
        writeCell(update._row, 'rawItems', update.rawItems);
        writeCell(update._row, 'unmapped', update.unmapped);
      }
    });
    var waybillStart = appendMatrix(sheet, (plan.inserts || []).map(waybillToRow));
    var flagStart = appendMatrix(ss.getSheetByName(SHEETS.FLAGS.name), (plan.flags || []).map(flagToRow));
    return {
      inserted: (plan.inserts || []).length,
      updated: (plan.updates || []).length,
      flags: (plan.flags || []).length,
      waybillStart: waybillStart,
      flagStart: flagStart
    };
  }

  function commitImport(ss, plan) {
    var waybillSheet = ss.getSheetByName(SHEETS.WAYBILLS.name);
    (plan.updates || []).forEach(function (update) {
      if (!update._row) return;
      waybillSheet.getRange(update._row, 1, 1, SHEETS.WAYBILLS.headers.length).setValues([waybillToRow(update)]);
    });
    var waybillStart = appendMatrix(waybillSheet, (plan.inserts || []).map(waybillToRow));
    var ledgerStart = appendMatrix(ss.getSheetByName(SHEETS.LEDGER.name), (plan.ledger || []).map(ledgerToRow));
    var flagStart = appendMatrix(ss.getSheetByName(SHEETS.FLAGS.name), (plan.flags || []).map(flagToRow));
    return {
      inserted: (plan.inserts || []).length,
      updated: (plan.updates || []).length,
      ledgerRows: (plan.ledger || []).length,
      flags: (plan.flags || []).length,
      warnings: plan.warnings || [],
      waybillStart: waybillStart,
      ledgerStart: ledgerStart,
      flagStart: flagStart
    };
  }

  function refreshStockFormulas(ss, items) {
    var sheet = ss.getSheetByName(SHEETS.STOCK.name);
    var values = readTableValues(sheet, SHEETS.STOCK.headers.length, 1);
    var rowsBySku = {};
    for (var i = 1; i < values.length; i++) {
      var existing = normalizeSku(values[i][0]);
      if (!existing) continue;
      if (!rowsBySku[existing]) rowsBySku[existing] = [];
      rowsBySku[existing].push(i + 1);
    }
    var needed = {};
    (items || []).forEach(function (item) {
      if (!isActiveItem(item)) return;
      var sku = normalizeSku(item.sku);
      if (!sku) return;
      needed[sku] = (needed[sku] || 0) + 1;
    });
    var rewritten = 0;
    var start = lastContentRow(sheet, 1) + 1;
    if (start < 2) start = 2;
    var matrix = [];
    Object.keys(needed).forEach(function (sku) {
      var have = rowsBySku[sku] || [];
      var n = 0;
      for (; n < have.length && n < needed[sku]; n++) {
        var formulas = stockFormulasForRow(have[n], n + 1);
        sheet.getRange(have[n], 2, 1, 5).setValues([[
          formulas.client,
          formulas.name,
          formulas.beginning,
          formulas.ledger,
          formulas.onHand
        ]]);
        rewritten += 1;
      }
      for (; n < needed[sku]; n++) {
        var row = start + matrix.length;
        var added = stockFormulasForRow(row, n + 1);
        matrix.push([sku, added.client, added.name, added.beginning, added.ledger, added.onHand]);
      }
    });
    appendMatrix(sheet, matrix);
    return { formulasRewritten: rewritten, added: matrix.length, active: Object.keys(needed).reduce(function (sum, sku) { return sum + needed[sku]; }, 0) };
  }

  function clearSheetBody(sheet, columnCount) {
    if (!sheet) return 0;
    var last = sheet.getLastRow();
    if (last < 2) return 0;
    var cols = columnCount || 1;
    sheet.getRange(2, 1, last - 1, cols).clearContent();
    return last - 1;
  }

  function recordsToMatrix(records, fields) {
    return (records || []).map(function (record) {
      return fields.map(function (field) {
        var value = record[field];
        if (value == null) return '';
        return value;
      });
    });
  }

  /**
   * Replace rows under the header without deleteRows.
   * Sheets throws "it is not possible to delete all non-frozen rows"
   * when row 1 is frozen and the delete would remove every row below it.
   * clearContent leaves that row in place and drops the values.
   */
  function replaceSheetBody(sheet, matrix, columnCount) {
    var cols = columnCount || (matrix && matrix[0] ? matrix[0].length : 1);
    var last = sheet.getLastRow();
    if (last > 1) {
      sheet.getRange(2, 1, last - 1, cols).clearContent();
    }
    if (matrix && matrix.length) {
      sheet.getRange(2, 1, matrix.length, matrix[0].length).setValues(matrix);
    }
  }

  function writeTable(ss, sheetKey, records) {
    var spec = SHEETS[sheetKey];
    var sheet = ss.getSheetByName(spec.name);
    replaceSheetBody(sheet, recordsToMatrix(records, spec.fields), spec.headers.length);
  }

  /**
   * Demo load keeps people already on STAFF (the real admin row) and only
   * adds sample emails that are not already there.
   */
  function staffRowsForDemo(existing, demoStaff, currentEmail) {
    var rows = [];
    var have = {};
    (existing || []).forEach(function (row) {
      var email = normalizeEmail(row.email);
      if (!email || have[email]) return;
      have[email] = true;
      rows.push({
        email: row.email,
        name: row.name,
        role: row.role,
        active: row.active,
        pick: row.pick,
        pack: row.pack,
        dispatch: row.dispatch
      });
    });
    (demoStaff || []).forEach(function (row) {
      var email = normalizeEmail(row.email);
      if (!email || have[email]) return;
      have[email] = true;
      rows.push({
        email: row.email,
        name: row.name,
        role: row.role,
        active: row.active,
        pick: row.pick,
        pack: row.pack,
        dispatch: row.dispatch
      });
    });
    var me = normalizeEmail(currentEmail);
    if (me && !have[me]) {
      rows.unshift({
        email: currentEmail,
        name: 'Admin',
        role: 'ADMIN',
        active: true,
        pick: true,
        pack: true,
        dispatch: true
      });
    }
    return rows;
  }

  function headerColumnIndex(headerRow, headers) {
    if (!headers || !headers.length) return null;
    var normalized = (headerRow || []).map(normHeader);
    var index = [];
    var found = 0;
    for (var i = 0; i < headers.length; i++) {
      var at = normalized.indexOf(normHeader(headers[i]));
      if (at >= 0) found += 1;
      index.push(at);
    }
    if (!found) return null;
    return index;
  }

  function rowsToRecords(values, fields, headers) {
    var out = [];
    if (!values || values.length < 2) return out;
    var mapped = headerColumnIndex(values[0], headers);
    for (var i = 1; i < values.length; i++) {
      var row = values[i] || [];
      function cell(fieldIndex) {
        var at = mapped ? mapped[fieldIndex] : fieldIndex;
        if (at == null || at < 0) return '';
        return row[at] == null ? '' : row[at];
      }
      var blank = true;
      for (var c = 0; c < fields.length; c++) {
        if (cell(c) !== '' && cell(c) != null) blank = false;
      }
      if (blank) continue;
      if (!cellHasContent(cell(0))) continue;
      var record = { _row: i + 1 };
      for (var f = 0; f < fields.length; f++) record[fields[f]] = cell(f);
      out.push(record);
    }
    return out;
  }

  function settingsFromRows(rows) {
    var out = { samePersonStages: true };
    (rows || []).forEach(function (row) {
      if (String(row.setting || '').trim().toUpperCase() === 'SAME_PERSON_STAGES') {
        out.samePersonStages = isTruthy(row.value);
      }
    });
    return out;
  }

  function hydrateModel(raw) {
    raw = raw || {};
    return {
      staff: (raw.staff || []).map(function (row) {
        return {
          _row: row._row,
          email: normalizeEmail(row.email),
          name: String(row.name || '').trim(),
          role: String(row.role || '').trim().toUpperCase() === 'ADMIN' ? 'ADMIN' : 'STAFF',
          active: isTruthy(row.active),
          pick: row.pick,
          pack: row.pack,
          dispatch: row.dispatch
        };
      }).filter(function (row) { return row.email; }),
      settings: settingsFromRows(raw.settings),
      items: (raw.items || []).map(function (row) {
        return {
          _row: row._row,
          sku: normalizeSku(row.sku),
          client: normalizeClient(row.client),
          productName: String(row.productName || '').trim(),
          barcode: row.barcode == null ? '' : String(row.barcode).trim(),
          beginningQty: num(row.beginningQty),
          active: isTruthy(row.active),
          jtAlias: String(row.jtAlias || '').trim()
        };
      }).filter(function (row) { return row.sku; }),
      skuMap: (raw.skuMap || []).map(function (row) {
        return {
          _row: row._row,
          raw: String(row.raw || '').trim(),
          client: normalizeClient(row.client),
          sku: normalizeSku(row.sku)
        };
      }),
      waybills: (raw.waybills || []).map(function (row) {
        return {
          _row: row._row,
          waybillNo: normWaybill(row.waybillNo),
          client: normalizeClient(row.client),
          itemString: String(row.itemString || '').trim(),
          rawItems: String(row.rawItems || '').trim(),
          orderDate: asDateOnly(row.orderDate),
          importedAt: asManilaString(row.importedAt),
          jtStatus: String(row.jtStatus || '').trim(),
          pickupTime: hasPickupTime(row.pickupTime) ? asManilaString(row.pickupTime) : '',
          signingTime: String(row.signingTime || '').trim(),
          rtsReason: String(row.rtsReason || '').trim(),
          unmapped: String(row.unmapped || '').trim()
        };
      }),
      scans: (raw.scans || []).map(function (row) {
        return {
          _row: row._row,
          timestamp: asManilaString(row.timestamp),
          staffEmail: normalizeEmail(row.staffEmail),
          staffName: String(row.staffName || '').trim(),
          stage: String(row.stage || '').trim().toUpperCase(),
          waybill: normWaybill(row.waybill),
          sku: normalizeSku(row.sku),
          qty: num(row.qty),
          result: String(row.result || '').trim().toUpperCase(),
          message: row.message == null ? '' : String(row.message)
        };
      }),
      flags: (raw.flags || []).map(function (row) {
        return {
          _row: row._row,
          timestamp: asManilaString(row.timestamp),
          waybill: normWaybill(row.waybill),
          client: normalizeClient(row.client),
          type: String(row.type || '').trim().toUpperCase(),
          details: row.details == null ? '' : String(row.details),
          staff: row.staff == null ? '' : String(row.staff),
          resolved: isTruthy(row.resolved)
        };
      }),
      ledger: (raw.ledger || []).map(function (row) {
        return {
          _row: row._row,
          date: asManilaString(row.date),
          client: normalizeClient(row.client),
          sku: normalizeSku(row.sku),
          type: String(row.type || '').trim().toUpperCase(),
          qty: num(row.qty),
          ref: normWaybill(row.ref),
          source: row.source == null ? '' : String(row.source)
        };
      })
    };
  }

  function findStaff(staffRows, email) {
    var want = normalizeEmail(email);
    if (!want) return null;
    var rows = staffRows || [];
    for (var i = 0; i < rows.length; i++) {
      if (normalizeEmail(rows[i].email) === want) {
        var role = String(rows[i].role || '').trim().toUpperCase() === 'ADMIN' ? 'ADMIN' : 'STAFF';
        return {
          email: want,
          name: String(rows[i].name || '').trim() || want,
          role: role,
          active: isTruthy(rows[i].active),
          pick: rows[i].pick,
          pack: rows[i].pack,
          dispatch: rows[i].dispatch
        };
      }
    }
    return null;
  }

  function stageEntryWhileLoading(ready) {
    if (ready) return { blocked: false, message: '', queue: false };
    return { blocked: true, message: MSG.LOADING, queue: false };
  }

  function adminCardText(flag) {
    var row = flag || {};
    return {
      timestamp: clipDisplay(row.timestamp),
      waybill: clipDisplay(row.waybill),
      client: clipDisplay(row.clientLabel || row.client),
      type: clipDisplay(row.type),
      details: clipDisplay(row.details),
      staff: clipDisplay(row.staff)
    };
  }

  function markFlagResolved(sheet, key) {
    var values = readTableValues(sheet, SHEETS.FLAGS.headers.length, 1);
    for (var i = 1; i < values.length; i++) {
      var sameTime = String(values[i][0]) === String(key.timestamp);
      var sameWaybill = normWaybill(values[i][1]) === normWaybill(key.waybill);
      var sameType = String(values[i][3] || '').toUpperCase() === String(key.type || '').toUpperCase();
      if (sameTime && sameWaybill && sameType) {
        sheet.getRange(i + 1, 7).setValue(true);
        return true;
      }
    }
    return false;
  }

  function labelCatalog(model) {
    return (model.items || []).filter(isActiveItem).map(function (item) {
      return {
        sku: normalizeSku(item.sku),
        client: normalizeClient(item.client),
        clientLabel: clientLabel(item.client),
        name: item.productName || '',
        barcode: item.barcode || ''
      };
    });
  }

  function hoursAgo(now, hours) {
    return new Date(now.getTime() - hours * 3600000);
  }

  function scan(when, name, email, stage, waybill, sku, qty, result, message) {
    return {
      timestamp: formatManila(when),
      staffEmail: email,
      staffName: name,
      stage: stage,
      waybill: waybill,
      sku: sku,
      qty: qty,
      result: result,
      message: message || ''
    };
  }

  /**
   * Demo rows for SIRWEN, TREXO (Josie), and ECOM.
   * Timestamps are relative to `now` so "today" on the admin screen stays true.
   */
  function buildSeedModel(now) {
    var moment = now || new Date();
    var ago = function (hours) { return hoursAgo(moment, hours); };
    var maria = 'Maria Santos';
    var jun = 'Jun Reyes';
    var mariaEmail = 'admin@vendes.example';
    var junEmail = 'staff@vendes.example';

    var items = [
      { sku: 'WVT-LW', client: 'SIRWEN', productName: 'Women Vest Light Wash', barcode: '4801001000011', beginningQty: 50, active: true, jtAlias: 'women vest light wash, vest lw' },
      { sku: 'FTB-QUEEN', client: 'SIRWEN', productName: 'Fitted Top Queen', barcode: '', beginningQty: 30, active: true, jtAlias: 'fitted top queen' },
      { sku: 'JOS-TEE', client: 'TREXO', productName: 'Josie Basic Tee', barcode: '4802002000022', beginningQty: 40, active: true, jtAlias: 'josie tee, josie basic tee' },
      { sku: 'JOS-CAP', client: 'TREXO', productName: 'Josie Cap', barcode: '', beginningQty: 20, active: true, jtAlias: 'josie cap' },
      { sku: 'ECM-MUG', client: 'ECOM', productName: 'Ecom Ceramic Mug', barcode: '4803003000033', beginningQty: 100, active: true, jtAlias: 'ecom mug, ceramic mug' },
      { sku: 'ECM-BAG', client: 'ECOM', productName: 'Ecom Tote Bag', barcode: '', beginningQty: 60, active: true, jtAlias: 'tote bag, ecom tote' },
      { sku: 'OLD-SKU', client: 'ECOM', productName: 'Discontinued Sample', barcode: '0000000000000', beginningQty: 0, active: false }
    ];

    var skuMap = [
      { raw: 'women vest light wash', client: 'SIRWEN', sku: 'WVT-LW' },
      { raw: 'vest lw', client: 'SIRWEN', sku: 'WVT-LW' },
      { raw: 'fitted top queen', client: 'SIRWEN', sku: 'FTB-QUEEN' },
      { raw: 'josie basic tee', client: 'TREXO', sku: 'JOS-TEE' },
      { raw: 'josie tee', client: 'TREXO', sku: 'JOS-TEE' },
      { raw: 'josie cap', client: 'TREXO', sku: 'JOS-CAP' },
      { raw: 'ecom mug', client: 'ECOM', sku: 'ECM-MUG' },
      { raw: 'ceramic mug', client: 'ECOM', sku: 'ECM-MUG' },
      { raw: 'tote bag', client: 'ECOM', sku: 'ECM-BAG' },
      { raw: 'ecom tote', client: 'ECOM', sku: 'ECM-BAG' }
    ];

    var staff = [
      { email: mariaEmail, name: maria, role: 'ADMIN', active: true, pick: true, pack: true, dispatch: true },
      { email: junEmail, name: jun, role: 'STAFF', active: true, pick: true, pack: true, dispatch: false },
      { email: 'old@vendes.example', name: 'Inactive Sample', role: 'STAFF', active: false, pick: false, pack: false, dispatch: false }
    ];

    function wb(no, client, itemString, status, pickupHours, orderHours) {
      return {
        waybillNo: no,
        client: client,
        itemString: itemString,
        orderDate: manilaDate(ago(orderHours == null ? 20 : orderHours)),
        importedAt: formatManila(ago(orderHours == null ? 18 : orderHours)),
        jtStatus: status,
        pickupTime: pickupHours == null ? '' : formatManila(ago(pickupHours))
      };
    }

    var waybills = [
      wb('WB-SIR-1001', 'SIRWEN', '2xWVT-LW+1xFTB-QUEEN', 'Printed', null, 16),
      wb('WB-SIR-1003', 'SIRWEN', '1xFTB-QUEEN', 'Printed', null, 15),
      wb('WB-SIR-1002', 'SIRWEN', '1xWVT-LW', 'Picked Up', 5, 30),
      wb('WB-TRX-2001', 'TREXO', '3xJOS-TEE+1xJOS-CAP', 'Printed', null, 12),
      wb('WB-TRX-2002', 'TREXO', '1xJOS-TEE', 'Printed', null, 11),
      wb('WB-TRX-OLD', 'TREXO', '1xJOS-CAP', 'Printed', null, 40),
      wb('WB-ECM-3001', 'ECOM', '1xECM-MUG+2xECM-BAG', 'Printed', null, 10),
      wb('WB-ECM-3002', 'ECOM', '1xECM-MUG', 'Printed', null, 9),
      wb('WB-ECM-DONE', 'ECOM', '1xECM-BAG', 'Printed', null, 8),
      wb('WB-ECM-UP', 'ECOM', '2xECM-MUG', 'Picked Up', 4, 14)
    ];

    var scans = [];
    function add(hours, name, email, stage, waybill, sku, qty, result, message) {
      scans.push(scan(ago(hours), name, email, stage, waybill, sku, qty, result, message));
    }

    add(0.4, maria, mariaEmail, 'PICK', 'WB-SIR-1001', 'JOS-TEE', 0, 'ERROR', MSG.WRONG_SKU);

    add(5.2, jun, junEmail, 'PICK', 'WB-SIR-1002', 'WVT-LW', 1, 'OK', '');
    add(5.1, jun, junEmail, 'PACK', 'WB-SIR-1002', 'WVT-LW', 1, 'OK', '');
    add(5.05, maria, mariaEmail, 'DISPATCH', 'WB-SIR-1002', '', 1, 'OK', 'BATCH');

    add(2, maria, mariaEmail, 'PICK', 'WB-TRX-2001', 'JOS-TEE', 1, 'OK', '');
    add(1.9, maria, mariaEmail, 'PICK', 'WB-TRX-2001', 'JOS-TEE', 1, 'OK', '');
    add(1.8, maria, mariaEmail, 'PICK', 'WB-TRX-2001', 'JOS-TEE', 1, 'OK', '');
    add(1.7, maria, mariaEmail, 'PICK', 'WB-TRX-2001', 'JOS-CAP', 1, 'OK', '');
    add(1.2, jun, junEmail, 'PACK', 'WB-TRX-2001', 'JOS-TEE', 1, 'OK', '');
    add(1.1, jun, junEmail, 'PACK', 'WB-TRX-2001', 'JOS-TEE', 1, 'OK', '');
    add(1.0, jun, junEmail, 'PACK', 'WB-TRX-2001', 'JOS-CAP', 1, 'OK', '');

    add(3, maria, mariaEmail, 'PICK', 'WB-TRX-2002', 'JOS-TEE', 1, 'OK', '');

    add(36, jun, junEmail, 'PICK', 'WB-TRX-OLD', 'JOS-CAP', 1, 'OK', '');
    add(35, jun, junEmail, 'PACK', 'WB-TRX-OLD', 'JOS-CAP', 1, 'OK', '');
    add(30, maria, mariaEmail, 'DISPATCH', 'WB-TRX-OLD', '', 1, 'OK', 'BATCH');

    add(3, maria, mariaEmail, 'PICK', 'WB-ECM-3001', 'ECM-MUG', 1, 'OK', '');
    add(2.9, jun, junEmail, 'PICK', 'WB-ECM-3001', 'ECM-BAG', 1, 'OK', '');
    add(2.8, jun, junEmail, 'PICK', 'WB-ECM-3001', 'ECM-BAG', 1, 'OK', '');
    add(2.4, jun, junEmail, 'PACK', 'WB-ECM-3001', 'ECM-MUG', 1, 'OK', '');
    add(2.3, jun, junEmail, 'PACK', 'WB-ECM-3001', 'ECM-BAG', 1, 'OK', '');
    add(2.2, jun, junEmail, 'PACK', 'WB-ECM-3001', 'ECM-BAG', 1, 'OK', '');

    add(2.6, maria, mariaEmail, 'PICK', 'WB-ECM-3002', 'ECM-MUG', 1, 'OK', '');
    add(2.1, jun, junEmail, 'PACK', 'WB-ECM-3002', 'ECM-MUG', 1, 'OK', '');

    add(3.4, maria, mariaEmail, 'PICK', 'WB-ECM-DONE', 'ECM-BAG', 1, 'OK', '');
    add(3.2, jun, junEmail, 'PACK', 'WB-ECM-DONE', 'ECM-BAG', 1, 'OK', '');
    add(2, maria, mariaEmail, 'DISPATCH', 'WB-ECM-DONE', '', 1, 'OK', 'BATCH');

    var flags = [
      makeFlag(ago(0.4), 'WB-SIR-1001', 'SIRWEN', 'WRONG_SKU', 'JOS-TEE @ PICK', maria),
      makeFlag(ago(1), 'WB-TRX-OLD', 'TREXO', 'DISPATCHED_NOT_PICKED_UP', 'Dispatched, no J&T pickup after 24h', 'SYSTEM'),
      makeFlag(ago(4), 'WB-ECM-UP', 'ECOM', 'PICKED_UP_NOT_DISPATCHED', 'J&T rider already picked up but there is no DISPATCH scan', 'SYSTEM'),
      {
        timestamp: formatManila(ago(26)),
        waybill: 'WB-SIR-1003',
        client: 'SIRWEN',
        type: 'QTY_SHORT',
        details: 'Old short that was already fixed',
        staff: jun,
        resolved: true
      }
    ];

    var ledger = [
      { date: formatManila(ago(5)), client: 'SIRWEN', sku: 'WVT-LW', type: 'SHIP_OUT', qty: -1, ref: 'WB-SIR-1002', source: 'JT_IMPORT' },
      { date: formatManila(ago(4)), client: 'ECOM', sku: 'ECM-MUG', type: 'SHIP_OUT', qty: -2, ref: 'WB-ECM-UP', source: 'JT_IMPORT' },
      { date: formatManila(ago(48)), client: 'ECOM', sku: 'ECM-BAG', type: 'PURCHASE', qty: 10, ref: 'PO-100', source: 'MANUAL' }
    ];

    var importSample = [
      {
        waybillNo: 'WB-SIR-1099',
        itemText: '2x women vest light wash + 1x fitted top queen',
        jtStatus: 'Picked Up',
        pickupTime: formatManila(ago(1)),
        orderDate: manilaDate(ago(6)),
        client: 'SIRWEN'
      },
      {
        waybillNo: 'WB-SIR-1002',
        itemText: '1xWVT-LW',
        jtStatus: 'Picked Up',
        pickupTime: formatManila(ago(5)),
        orderDate: manilaDate(ago(30)),
        client: 'SIRWEN'
      },
      {
        waybillNo: 'WB-TRX-2099',
        itemText: '1x josie tee + mystery gadget',
        jtStatus: 'Printed',
        pickupTime: '',
        orderDate: manilaDate(ago(2)),
        client: 'TREXO'
      }
    ];

    return {
      staff: staff,
      items: items,
      skuMap: skuMap,
      waybills: waybills,
      scans: scans,
      flags: flags,
      ledger: ledger,
      importSample: importSample,
      settings: { samePersonStages: true }
    };
  }

  return {
    SHEET_ORDER: SHEET_ORDER,
    SHEETS: SHEETS,
    MSG: MSG,
    normalizeSku: normalizeSku,
    normalizeClient: normalizeClient,
    normalizeEmail: normalizeEmail,
    normWaybill: normWaybill,
    limitScanText: limitScanText,
    clipDisplay: clipDisplay,
    saveBadge: saveBadge,
    adminCardText: adminCardText,
    stageEntryWhileLoading: stageEntryWhileLoading,
    stageAssigned: stageAssigned,
    samePersonConflict: samePersonConflict,
    samePersonMessage: samePersonMessage,
    samePersonEnabled: samePersonEnabled,
    stageGate: stageGate,
    looksLikeWaybill: looksLikeWaybill,
    cellHasContent: cellHasContent,
    stageOwnerRows: stageOwnerRows,
    flagItemKey: flagItemKey,
    labelBarcodeValue: labelBarcodeValue,
    labelKey: labelKey,
    labelCaption: labelCaption,
    labelPrintMode: labelPrintMode,
    code128Symbols: code128Symbols,
    code128Widths: code128Widths,
    code128Svg: code128Svg,
    expandLabelCopies: expandLabelCopies,
    labelPrintDocument: labelPrintDocument,
    roleCanUseAdmin: roleCanUseAdmin,
    clientLabel: clientLabel,
    isTruthy: isTruthy,
    isResolved: isResolved,
    isActiveItem: isActiveItem,
    formatManila: formatManila,
    manilaDate: manilaDate,
    parseManila: parseManila,
    isPickedUp: isPickedUp,
    statusMeansPickedUp: statusMeansPickedUp,
    tryParseCompact: tryParseCompact,
    parseItemString: parseItemString,
    formatItemString: formatItemString,
    resolveLineItems: resolveLineItems,
    qtyNear: qtyNear,
    buildWaybillView: buildWaybillView,
    buildStagePayload: buildStagePayload,
    decide: decide,
    applyDecisionToView: applyDecisionToView,
    applyScans: applyScans,
    mapImportHeaders: mapImportHeaders,
    rowToImport: rowToImport,
    rowsFromImportSheet: rowsFromImportSheet,
    planJtImport: planJtImport,
    planJtExport: planJtExport,
    parseJtItemName: parseJtItemName,
    jtCreatorClient: jtCreatorClient,
    jtExportView: jtExportView,
    falseOnlySpans: falseOnlySpans,
    planAutoFlags: planAutoFlags,
    ledgerHasShipOut: ledgerHasShipOut,
    buildAdminDashboard: buildAdminDashboard,
    buildAdminSummary: buildAdminSummary,
    unknownItemMessage: unknownItemMessage,
    shortageAllowed: shortageAllowed,
    nextBanner: nextBanner,
    countStageToday: countStageToday,
    blockedWithin: blockedWithin,
    missingStageMessage: missingStageMessage,
    clearSheetBody: clearSheetBody,
    onHand: onHand,
    stockFormulasForRow: stockFormulasForRow,
    commitImport: commitImport,
    commitJtExport: commitJtExport,
    refreshStockFormulas: refreshStockFormulas,
    recordsToMatrix: recordsToMatrix,
    writeTable: writeTable,
    replaceSheetBody: replaceSheetBody,
    staffRowsForDemo: staffRowsForDemo,
    rowsToRecords: rowsToRecords,
    hydrateModel: hydrateModel,
    findStaff: findStaff,
    markFlagResolved: markFlagResolved,
    labelCatalog: labelCatalog,
    buildSeedModel: buildSeedModel,
    appendMatrix: appendMatrix,
    lastContentRow: lastContentRow,
    readTableValues: readTableValues,
    compactSheetBody: compactSheetBody,
    columnName: columnName,
    nextSheetRange: nextSheetRange,
    buildSyncBatch: buildSyncBatch,
    syncWrites: syncWrites,
    attachSyncRows: attachSyncRows,
    pendingLabel: pendingLabel
  };
});
