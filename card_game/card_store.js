// =============================================================
//  共有カードストレージ (window.CardStore)
//  - 数百〜数千枚のカードを写真込みで保存するため IndexedDB を使用する。
//  - 旧来の localStorage('mlb_card_extras_v1') は容量が約5MB
//    (= 1枚 約50〜60KB のカードで 45〜50枚程度) ですぐ上限に達し、
//    それ以上「追加がきかない」状態になっていた。
//  - 初回起動時に旧 localStorage のカードを IndexedDB へ自動移行し、
//    localStorage 側は解放する。
//  - IndexedDB が使えない環境では localStorage にフォールバックする。
//  - game.js は同期的に選手を読み込むため、起動時に init() で
//    全件をメモリキャッシュ(_cache)へ載せ、getCachedSync() で同期参照する。
// =============================================================
(function(){
  'use strict';
  const DB_NAME    = 'mlb_card_db';
  const STORE      = 'cards';
  const DB_VER     = 1;
  const LEGACY_KEY = 'mlb_card_extras_v1';

  let _cache = null;    // 同期参照用メモリキャッシュ (card オブジェクト配列)
  let _mode  = 'idb';   // 'idb' | 'ls'

  // カードの同一性キー: 名前 + 年 + チーム + 種別(投手/打者)。
  // これにより、大谷翔平のように同名・同年・同チームでも「投手版」と「打者版」を
  // 別カードとして両方登録できる。種別はカードの type、無ければ球種の有無から判定。
  function cardType(p) {
    if (p && p.type) return p.type;
    return (p && p.pitches && p.pitches.length > 0) ? 'pitcher' : 'batter';
  }
  function cardKey(p) {
    if (!p) return '';
    return (p.fullNameTop || '') + '_' + (p.year || '') + '_' + (p.team || '') + '_' + cardType(p);
  }
  function hasIDB() { try { return !!window.indexedDB; } catch (e) { return false; } }

  function openDB() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(DB_NAME, DB_VER);
      req.onupgradeneeded = () => {
        const db = req.result;
        if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, { keyPath: 'key' });
      };
      req.onsuccess = () => resolve(req.result);
      req.onerror   = () => reject(req.error);
    });
  }

  // ---- localStorage フォールバック ----
  function lsLoad() { try { return JSON.parse(localStorage.getItem(LEGACY_KEY)) || []; } catch (e) { return []; } }
  function lsSave(arr) { localStorage.setItem(LEGACY_KEY, JSON.stringify(arr)); }

  // ---- IndexedDB 操作 ----
  async function idbGetAll() {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).getAll();
      req.onsuccess = () => resolve((req.result || []).map(r => r.card));
      req.onerror   = () => reject(req.error);
    });
  }
  async function idbPut(cards) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      const store = tx.objectStore(STORE);
      (Array.isArray(cards) ? cards : [cards]).forEach(c => store.put({ key: cardKey(c), card: c }));
      tx.oncomplete = () => resolve();
      tx.onerror    = () => reject(tx.error);
    });
  }
  async function idbDelete(key) {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).delete(key);
      tx.oncomplete = () => resolve();
      tx.onerror    = () => reject(tx.error);
    });
  }
  async function idbClear() {
    const db = await openDB();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      tx.objectStore(STORE).clear();
      tx.oncomplete = () => resolve();
      tx.onerror    = () => reject(tx.error);
    });
  }

  // 旧 localStorage → IndexedDB への一度きり移行
  async function migrate() {
    const legacy = lsLoad();
    if (!legacy.length) return 0;
    await idbPut(legacy);
    try { localStorage.removeItem(LEGACY_KEY); } catch (e) {}
    return legacy.length;
  }

  // キー形式の変更 (名前+年 → 名前+年+チーム+種別) に伴う再キー移行。
  // 旧キーで保存された既存レコードを新キーへ振り直し、重複が増えないようにする。
  async function rekeyMigrationIDB() {
    const db = await openDB();
    const records = await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readonly');
      const req = tx.objectStore(STORE).getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror   = () => reject(req.error);
    });
    const toFix = records.filter(r => r && r.card && r.key !== cardKey(r.card));
    if (!toFix.length) return 0;
    await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, 'readwrite');
      const store = tx.objectStore(STORE);
      for (const r of toFix) {
        store.delete(r.key);
        store.put({ key: cardKey(r.card), card: r.card });
      }
      tx.oncomplete = () => resolve();
      tx.onerror    = () => reject(tx.error);
    });
    return toFix.length;
  }

  // 起動時: モード判定 → 移行 → 全件をキャッシュ
  async function init() {
    if (hasIDB()) {
      try {
        await migrate();
        await rekeyMigrationIDB();   // 旧キー形式のレコードを新キーへ振り直し
        _cache = await idbGetAll();
        _mode = 'idb';
        return _cache;
      } catch (e) {
        console.error('CardStore: IndexedDB 初期化に失敗。localStorage を使用します。', e);
      }
    }
    _mode = 'ls';
    _cache = lsLoad();
    return _cache;
  }

  // ---- 公開 API (すべて Promise を返す) ----
  async function getAll() {
    _cache = (_mode === 'idb') ? await idbGetAll() : lsLoad();
    return _cache;
  }
  async function addOrUpdate(card) {
    if (_mode === 'idb') {
      await idbPut(card);
    } else {
      const arr = lsLoad();
      const k = cardKey(card);
      const i = arr.findIndex(x => cardKey(x) === k);
      if (i >= 0) arr[i] = card; else arr.push(card);
      lsSave(arr);   // localStorage モードでは容量超過で例外が出る可能性あり (呼び出し側で捕捉)
    }
    return getAll();
  }
  async function removeByKey(key) {
    if (_mode === 'idb') await idbDelete(key);
    else lsSave(lsLoad().filter(x => cardKey(x) !== key));
    return getAll();
  }
  async function clearAll() {
    if (_mode === 'idb') await idbClear();
    else { try { localStorage.removeItem(LEGACY_KEY); } catch (e) {} }
    _cache = [];
    return _cache;
  }

  function getCachedSync() { return _cache || []; }
  function mode() { return _mode; }

  // ブラウザのストレージ見積もり (使用量/上限) — 取得できなければ null
  async function estimate() {
    if (navigator.storage && navigator.storage.estimate) {
      try { return await navigator.storage.estimate(); } catch (e) {}
    }
    return null;
  }

  window.CardStore = {
    init, getAll, addOrUpdate, removeByKey, clearAll,
    getCachedSync, cardKey, mode, estimate,
  };
})();
