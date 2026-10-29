/* OpenCode for Excel - persistent workbook change history.
 *
 * Stores bounded change checkpoints in IndexedDB, scoped to a workbook key.
 * Snapshots can contain thousands of cells, so localStorage is intentionally
 * not used. Hosts without IndexedDB (and the test harness) fall back to an
 * in-memory store; restoration still works for the current pane session.
 *
 * Exposes window.OpenCodeExcelHistory.
 */
(function (root) {
  "use strict";

  var DB_NAME = "opencode-excel-history";
  var DB_VERSION = 1;
  var STORE = "checkpoints";
  var memory = {}; // workbookKey -> [entry]
  var dbPromise = null;

  function indexedDBAvailable() {
    try {
      return typeof indexedDB !== "undefined" && indexedDB && typeof indexedDB.open === "function";
    } catch (err) {
      return false;
    }
  }

  function clone(value) {
    if (value == null) return value;
    try { return JSON.parse(JSON.stringify(value)); }
    catch (err) { return value; }
  }

  function openDB() {
    if (!indexedDBAvailable()) return Promise.resolve(null);
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (resolve) {
      var req;
      try { req = indexedDB.open(DB_NAME, DB_VERSION); }
      catch (err) { resolve(null); return; }
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(STORE)) {
          var store = db.createObjectStore(STORE, { keyPath: "id" });
          store.createIndex("workbookKey", "workbookKey", { unique: false });
        }
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { resolve(null); };
      req.onblocked = function () { resolve(null); };
    });
    return dbPromise;
  }

  function withStore(mode, fn) {
    return openDB().then(function (db) {
      if (!db) return null;
      return new Promise(function (resolve, reject) {
        var tx;
        try { tx = db.transaction(STORE, mode); }
        catch (err) { resolve(null); return; }
        var store = tx.objectStore(STORE);
        var result;
        try { result = fn(store); }
        catch (err) { reject(err); return; }
        tx.oncomplete = function () { resolve(result && result.result !== undefined ? result.result : result); };
        tx.onerror = function () { reject(tx.error || new Error("History transaction failed")); };
        tx.onabort = function () { reject(tx.error || new Error("History transaction aborted")); };
      });
    }).catch(function () { return null; });
  }

  function memoryList(key) {
    return (memory[key] || []).slice().sort(function (a, b) { return (b.at || 0) - (a.at || 0); });
  }

  async function put(workbookKey, entry) {
    var record = clone(entry);
    record.workbookKey = workbookKey;
    var saved = await withStore("readwrite", function (store) { store.put(record); return record; });
    if (!saved) {
      if (!memory[workbookKey]) memory[workbookKey] = [];
      var list = memory[workbookKey];
      var index = -1;
      for (var i = 0; i < list.length; i++) if (list[i].id === record.id) { index = i; break; }
      if (index >= 0) list[index] = record; else list.push(record);
      return clone(record);
    }
    return clone(saved);
  }

  async function list(workbookKey, limit) {
    var entries = await withStore("readonly", function (store) {
      return store.index("workbookKey").getAll(workbookKey);
    });
    if (!entries) entries = memoryList(workbookKey);
    entries.sort(function (a, b) { return (b.at || 0) - (a.at || 0); });
    if (limit && entries.length > limit) entries = entries.slice(0, limit);
    // Metadata is enough for rendering; snapshots stay in the store.
    return entries.map(function (entry) {
      return {
        id: entry.id,
        at: entry.at,
        summary: entry.summary,
        ops: entry.ops || [],
        okCount: entry.okCount,
        total: entry.total,
        restorable: !!entry.restorable,
        hasSnapshot: !!entry.snapshot,
        workbookKey: entry.workbookKey,
      };
    });
  }

  async function get(id) {
    var entry = await withStore("readonly", function (store) { return store.get(id); });
    if (entry) return clone(entry);
    for (var key in memory) {
      for (var i = 0; i < memory[key].length; i++) {
        if (memory[key][i].id === id) return clone(memory[key][i]);
      }
    }
    return null;
  }

  async function remove(id) {
    var removed = await withStore("readwrite", function (store) { store.delete(id); return true; });
    for (var key in memory) {
      memory[key] = memory[key].filter(function (entry) { return entry.id !== id; });
    }
    return !!removed || true;
  }

  async function clear(workbookKey) {
    var removed = await withStore("readwrite", function (store) {
      var index = store.index("workbookKey");
      var req = index.openCursor(IDBKeyRange.only(workbookKey));
      var count = 0;
      req.onsuccess = function () {
        var cursor = req.result;
        if (!cursor) return;
        cursor.delete();
        count++;
        cursor.continue();
      };
      return { get result() { return count; } };
    });
    var local = (memory[workbookKey] || []).length;
    memory[workbookKey] = [];
    return Math.max(Number(removed) || 0, local);
  }

  root.OpenCodeExcelHistory = {
    put: put,
    list: list,
    get: get,
    remove: remove,
    clear: clear,
    available: indexedDBAvailable,
  };
})(typeof globalThis !== "undefined" ? globalThis : this);
