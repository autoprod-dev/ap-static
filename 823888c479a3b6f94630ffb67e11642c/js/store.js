/* Local-first storage in IndexedDB. Files: {id,name,kind,content,blob?,createdAt,updatedAt,rev,baseRev,dirty,deleted,deviceId} */
(function(){
  'use strict';
  const DB_NAME = 'my-office', DB_VER = 1;
  let dbp = null;
  function open(){
    if (dbp) return dbp;
    dbp = new Promise((res, rej) => {
      const r = indexedDB.open(DB_NAME, DB_VER);
      r.onupgradeneeded = () => {
        const db = r.result;
        if (!db.objectStoreNames.contains('files')) db.createObjectStore('files', {keyPath:'id'}).createIndex('updatedAt','updatedAt');
        if (!db.objectStoreNames.contains('meta')) db.createObjectStore('meta');
        if (!db.objectStoreNames.contains('queue')) db.createObjectStore('queue', {keyPath:'fileId'});
        if (!db.objectStoreNames.contains('handles')) db.createObjectStore('handles');
      };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
    return dbp;
  }
  function tx(store, mode, fn){
    return open().then(db => new Promise((res, rej) => {
      const t = db.transaction(store, mode); const s = t.objectStore(store);
      let out; const r = fn(s);
      if (r && 'onsuccess' in r) r.onsuccess = () => { out = r.result; };
      t.oncomplete = () => res(out); t.onerror = () => rej(t.error); t.onabort = () => rej(t.error);
    }));
  }
  const memHandles = new Map();
  const uid = () => (crypto.randomUUID ? crypto.randomUUID() : 'id-' + Date.now().toString(36) + Math.random().toString(36).slice(2));
  const Store = {
    uid,
    get: id => tx('files','readonly', s => s.get(id)),
    all: () => tx('files','readonly', s => s.getAll()),
    rawPut: f => tx('files','readwrite', s => s.put(f)),
    async deviceId(){
      let d = await Store.getMeta('deviceId');
      if (!d) { d = 'dev-' + uid().slice(0,8); await Store.setMeta('deviceId', d); }
      return d;
    },
    async recent(){
      const all = await Store.all();
      return all.filter(f => !f.deleted).sort((a,b) => b.updatedAt - a.updatedAt);
    },
    async create(kind, name, content, extra){
      const now = Date.now();
      const f = Object.assign({id: uid(), kind, name, content, createdAt: now, updatedAt: now, rev: null, baseRev: null, dirty: true, deleted: false, deviceId: await Store.deviceId()}, extra || {});
      await Store.rawPut(f); await Store.enqueue(f.id);
      return f;
    },
    /* Local edit: bumps updatedAt, marks dirty, queues for sync. */
    async save(id, patch){
      const f = await Store.get(id); if (!f) throw new Error('File not found');
      Object.assign(f, patch, {updatedAt: Date.now(), dirty: true, deviceId: await Store.deviceId()});
      await Store.rawPut(f); await Store.enqueue(id);
      if (window.MOSync) MOSync.poke();
      return f;
    },
    async remove(id){
      const f = await Store.get(id); if (!f) return;
      f.deleted = true; f.content = null; f.blob = null; f.updatedAt = Date.now(); f.dirty = true;
      await Store.rawPut(f); await Store.enqueue(id);
      memHandles.delete(id); await tx('handles','readwrite', s => s.delete(id)).catch(() => {});
      if (window.MOSync) MOSync.poke();
    },
    enqueue: fileId => tx('queue','readwrite', s => s.put({fileId, at: Date.now()})),
    dequeue: fileId => tx('queue','readwrite', s => s.delete(fileId)),
    queue: () => tx('queue','readonly', s => s.getAll()),
    getMeta: k => tx('meta','readonly', s => s.get(k)),
    setMeta: (k, v) => tx('meta','readwrite', s => s.put(v, k)),
    async getHandle(id){ if (memHandles.has(id)) return memHandles.get(id); try { return await tx('handles','readonly', s => s.get(id)); } catch (e) { return null; } },
    async setHandle(id, h){ memHandles.set(id, h); try { await tx('handles','readwrite', s => s.put(h, id)); } catch (e) { /* not storable (older browser): keep for this session */ } },
  };
  window.MOStore = Store;
})();
