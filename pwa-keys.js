(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root) root.VendesPwaKeys = api;
})(typeof window !== 'undefined' ? window : null, function () {
  var LEGACY_QUEUE_KEY = 'vendes.pwa.queue';
  var NOTICE = 'Old test scans on this phone were not sent';

  function storageKey(spreadsheetId, name) {
    return 'vendes.pwa.' + name + '.' + String(spreadsheetId || '');
  }

  function legacyQueueHasItems(raw) {
    if (!raw) return false;
    var saved;
    try { saved = JSON.parse(raw); } catch (err) { return false; }
    if (!saved || typeof saved !== 'object' || Array.isArray(saved)) return false;
    var queue = saved.queue || [];
    var resolves = saved.resolves || [];
    return queue.length > 0 || resolves.length > 0;
  }

  // Drop the old un-namespaced queue. Do not copy it onto a sheet key.
  function retireLegacyQueue(storage) {
    if (!storage) return '';
    var raw = '';
    try { raw = storage.getItem(LEGACY_QUEUE_KEY) || ''; } catch (err) { return ''; }
    if (!legacyQueueHasItems(raw)) return '';
    try { storage.removeItem(LEGACY_QUEUE_KEY); } catch (err) {}
    return NOTICE;
  }

  return {
    LEGACY_QUEUE_KEY: LEGACY_QUEUE_KEY,
    NOTICE: NOTICE,
    storageKey: storageKey,
    legacyQueueHasItems: legacyQueueHasItems,
    retireLegacyQueue: retireLegacyQueue
  };
});
