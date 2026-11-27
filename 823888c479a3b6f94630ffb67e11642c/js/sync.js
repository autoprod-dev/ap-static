/* Pluggable sync: local-first IndexedDB + adapter. Last-write-wins with a conflict copy, offline queue.
   Adapter interface:
     name: string
     ready(): Promise<boolean>          // signed in / configured
     signIn(): Promise<void>            // interactive, user-initiated only
     list(): Promise<[{id, rev, updatedAt, deleted}]>
     get(id): Promise<payload>          // payload = {meta, content, blobB64?}
     put(meta, payload): Promise<void>  // meta = {id, rev, updatedAt, deleted, kind, deviceId}
*/
(function(){
  'use strict';
  const S = () => window.MOStore;
  const listeners = new Set();
  let adapter = null, running = null, pokeTimer = null, intervalTimer = null;
  let status = {state:'off', text:'Sync off', pending:0, last:null};

  function setStatus(p){ status = Object.assign({}, status, p); listeners.forEach(fn => { try{fn(status);}catch(e){} }); }

  function blobToB64(blob){ return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.onerror = () => rej(r.error); r.readAsDataURL(blob); }); }
  function b64ToBlob(b64, type){ const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i=0;i<bin.length;i++) u[i]=bin.charCodeAt(i); return new Blob([u], {type: type || 'application/octet-stream'}); }

  async function toPayload(f){
    const meta = {id:f.id, name:f.name, kind:f.kind, updatedAt:f.updatedAt, createdAt:f.createdAt, deleted:!!f.deleted, deviceId:f.deviceId, mime:f.mime||null};
    const p = {meta, content: f.deleted ? null : f.content};
    if (f.blob && !f.deleted) p.blobB64 = await blobToB64(f.blob);
    return p;
  }
  function fromPayload(p, rev){
    const m = p.meta || {};
    return {id:m.id, name:m.name, kind:m.kind, updatedAt:m.updatedAt, createdAt:m.createdAt||m.updatedAt, deleted:!!m.deleted, deviceId:m.deviceId, mime:m.mime||null,
      content: p.content, blob: p.blobB64 ? b64ToBlob(p.blobB64, m.mime) : null, rev, baseRev: rev, dirty:false};
  }
  function stamp(d){ try { return new Date(d).toLocaleString('en-AU', {day:'numeric', month:'short', hour:'numeric', minute:'2-digit'}); } catch(e){ return ''; } }

  async function conflictCopy(src, fromWhere){
    const c = Object.assign({}, src, {id: S().uid(), name: `${src.name} (conflict copy, edited ${stamp(src.updatedAt)})`, rev:null, baseRev:null, dirty:true, deleted:false});
    await S().rawPut(c); await S().enqueue(c.id);
    return c;
  }

  async function pushOne(f){
    const rev = S().uid();
    const before = f.updatedAt;
    const payload = await toPayload(f);
    await adapter.put({id:f.id, rev, updatedAt:f.updatedAt, deleted:!!f.deleted, kind:f.kind, deviceId:f.deviceId}, payload);
    const now = await S().get(f.id);
    now.rev = rev; now.baseRev = rev;
    if (now.updatedAt === before) { now.dirty = false; await S().rawPut(now); await S().dequeue(f.id); }
    else await S().rawPut(now); // edited during upload: stays queued
  }

  async function syncNow(){
    if (!adapter) { setStatus({state:'off', text:'Sync off'}); return status; }
    if (running) return running;
    running = (async () => {
      const q = await S().queue();
      if (!navigator.onLine) { setStatus({state:'offline', text: q.length ? `Offline · ${q.length} waiting` : 'Offline', pending:q.length}); return status; }
      if (!(await adapter.ready())) { setStatus({state:'signin', text:'Sync: sign in', pending:q.length}); return status; }
      setStatus({state:'syncing', text:'Syncing…'});
      try {
        const remote = await adapter.list();
        const rmap = new Map(remote.map(r => [r.id, r]));
        const pushed = new Set();
        let conflicts = 0;
        let failed = 0;
        for (const item of q) {
          try {
          const f = await S().get(item.fileId);
          if (!f) { await S().dequeue(item.fileId); continue; }
          const r = rmap.get(f.id);
          const remoteChanged = r && r.rev && r.rev !== f.baseRev;
          if (remoteChanged) {
            if (f.deleted && r.deleted) { f.dirty=false; f.baseRev=f.rev=r.rev; await S().rawPut(f); await S().dequeue(f.id); pushed.add(f.id); continue; }
            const rp = await adapter.get(f.id);
            const rf = fromPayload(rp, r.rev);
            conflicts++;
            if ((f.updatedAt || 0) >= (rf.updatedAt || 0)) {
              // Local is newer: keep local as the main file, save the other device's version as a copy.
              if (!rf.deleted) await conflictCopy(rf, 'other device');
              await pushOne(f);
            } else {
              // Remote is newer: it wins, local edits are kept as a copy.
              if (!f.deleted) await conflictCopy(f, 'this device');
              await S().rawPut(rf); await S().dequeue(f.id);
            }
          } else {
            await pushOne(f);
          }
          pushed.add(f.id);
          } catch (e) { if (/Not signed in|Sign-in expired|Failed to fetch|NetworkError/i.test(e.message)) throw e; failed++; console.warn('Could not sync one file:', e.message); }
        }
        // Conflict copies created above are queued; push them now.
        for (const item of await S().queue()) {
          if (pushed.has(item.fileId)) continue;
          const f = await S().get(item.fileId); if (f && !rmap.has(f.id)) { await pushOne(f); pushed.add(f.id); }
        }
        let pulled = 0;
        for (const r of remote) {
          if (pushed.has(r.id)) continue;
          const f = await S().get(r.id);
          if (f && f.dirty) continue;
          if (f && f.baseRev === r.rev) continue;
          if (!f && r.deleted) continue;
          const rp = await adapter.get(r.id);
          await S().rawPut(fromPayload(rp, r.rev)); pulled++;
        }
        const left = (await S().queue()).length;
        setStatus({state:'on', text: failed ? `Synced · ${failed} couldn't sync` : left ? `Synced · ${left} waiting` : 'Synced', pending:left, last:Date.now(), conflicts, pulled, error: failed ? 'A file over 5 MB cannot sync yet' : null});
        if (pulled || conflicts) window.dispatchEvent(new CustomEvent('mo-files-changed'));
      } catch (e) {
        console.warn('Sync failed:', e && e.message);
        const left = (await S().queue()).length;
        setStatus({state:'error', text: left ? `Sync paused · ${left} waiting` : 'Sync paused', pending:left, error: String(e && e.message || e)});
      }
      return status;
    })();
    try { return await running; } finally { running = null; }
  }

  /* -------- Google Drive appDataFolder adapter (hidden app-only folder; scope drive.appdata) -------- */
  function DriveAppDataAdapter(opts){
    const api = opts.apiBase || 'https://www.googleapis.com/drive/v3';
    const up = opts.uploadBase || 'https://www.googleapis.com/upload/drive/v3';
    // Google's access pass lasts about an hour; keep it across app restarts until it expires (it only covers the hidden app folder).
    let token = null, tokenExp = 0, idCache = new Map();
    try { const t = JSON.parse(localStorage.getItem('mo-gtoken') || 'null'); if (t && t.c === opts.clientId && t.e > Date.now() + 60000) { token = t.t; tokenExp = t.e; } } catch (e) {}
    async function getToken(interactive){
      if (opts.getToken) return opts.getToken(interactive);
      if (token && Date.now() < tokenExp - 60000) return token;
      if (!interactive) return null;
      await loadGis();
      return new Promise((res, rej) => {
        const tc = google.accounts.oauth2.initTokenClient({client_id: opts.clientId, scope:'https://www.googleapis.com/auth/drive.appdata', callback: r => {
          if (r.error) return rej(new Error(r.error));
          token = r.access_token; tokenExp = Date.now() + (r.expires_in||3600)*1000;
          try { localStorage.setItem('mo-gtoken', JSON.stringify({t: token, e: tokenExp, c: opts.clientId})); } catch (e) {}
          res(token);
        }});
        tc.requestAccessToken({prompt: ''});
      });
    }
    function loadGis(){
      if (window.google && google.accounts && google.accounts.oauth2) return Promise.resolve();
      return new Promise((res, rej) => { const s = document.createElement('script'); s.src = 'https://accounts.google.com/gsi/client'; s.onload = res; s.onerror = () => rej(new Error('Could not load Google sign-in')); document.head.appendChild(s); });
    }
    async function call(url, init){
      const t = await getToken(false); if (!t) throw new Error('Not signed in');
      init = init || {}; init.headers = Object.assign({Authorization: 'Bearer ' + t}, init.headers || {});
      const r = await fetch(url, init);
      if (r.status === 401) { token = null; try { localStorage.removeItem('mo-gtoken'); } catch (e) {} throw new Error('Sign-in expired'); }
      if (!r.ok) throw new Error('Drive error ' + r.status);
      return r;
    }
    function multipart(metadata, body){
      const b = 'mo' + Math.random().toString(36).slice(2);
      const data = `--${b}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n--${b}\r\nContent-Type: application/json\r\n\r\n${body}\r\n--${b}--`;
      return {body:data, type:`multipart/related; boundary=${b}`};
    }
    return {
      name: opts.label || 'Google Drive (app-only folder)',
      async ready(){ return !!(await getToken(false)); },
      async signIn(){ await getToken(true); },
      async list(){
        const out = []; let pageToken = '';
        do {
          const u = `${api}/files?spaces=appDataFolder&pageSize=1000&fields=${encodeURIComponent('nextPageToken,files(id,name,appProperties,modifiedTime)')}` + (pageToken ? `&pageToken=${encodeURIComponent(pageToken)}` : '');
          const j = await (await call(u)).json();
          for (const f of j.files || []) {
            const a = f.appProperties || {}; if (!a.moId) continue;
            idCache.set(a.moId, f.id);
            out.push({id:a.moId, rev:a.rev, updatedAt:Number(a.updatedAt)||0, deleted:a.deleted === '1', kind:a.kind});
          }
          pageToken = j.nextPageToken || '';
        } while (pageToken);
        return out;
      },
      async get(id){
        const did = idCache.get(id); if (!did) throw new Error('Unknown remote file');
        return (await call(`${api}/files/${did}?alt=media`)).json();
      },
      async put(meta, payload){
        const appProperties = {moId:meta.id, rev:meta.rev, updatedAt:String(meta.updatedAt), deleted: meta.deleted ? '1':'0', kind:meta.kind||'', device:String(meta.deviceId||'').slice(0,40)};
        const did = idCache.get(meta.id);
        const m = did ? {appProperties} : {name:`mo-${meta.id}.json`, parents:['appDataFolder'], mimeType:'application/json', appProperties};
        const body = JSON.stringify(payload);
        if (body.length > 5 * 1024 * 1024) throw new Error('File too large to sync yet (over 5 MB)');
        const mp = multipart(m, body);
        const r = await call(did ? `${up}/files/${did}?uploadType=multipart&fields=id` : `${up}/files?uploadType=multipart&fields=id`, {method: did ? 'PATCH' : 'POST', headers:{'Content-Type': mp.type}, body: mp.body});
        const j = await r.json(); idCache.set(meta.id, j.id);
      }
    };
  }

  async function configure(cfg){
    await S().setMeta('sync', cfg);
    return applyConfig(cfg);
  }
  async function applyConfig(cfg){
    adapter = null;
    if (!cfg || cfg.mode !== 'gdrive') { try { localStorage.removeItem('mo-gtoken'); } catch (e) {} }
    if (cfg && cfg.mode === 'gdrive' && cfg.clientId) adapter = DriveAppDataAdapter({clientId: cfg.clientId});
    if (cfg && cfg.mode === 'mock' && cfg.mockBase) adapter = DriveAppDataAdapter({label:'Test server (mock Drive)', apiBase: cfg.mockBase + '/drive/v3', uploadBase: cfg.mockBase + '/upload/drive/v3', getToken: async () => 'mock-token'});
    if (!adapter) setStatus({state:'off', text:'Sync off'});
    clearInterval(intervalTimer);
    if (adapter) { intervalTimer = setInterval(() => { if (document.visibilityState === 'visible') syncNow(); }, 60000); syncNow(); }
    return adapter;
  }
  async function init(){
    const cfg = await S().getMeta('sync');
    window.addEventListener('online', () => syncNow());
    window.addEventListener('offline', () => syncNow());
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && adapter) syncNow(); });
    return applyConfig(cfg);
  }

  window.MOSync = {
    init, configure, syncNow, DriveAppDataAdapter,
    signIn: async () => { if (adapter) { await adapter.signIn(); return syncNow(); } },
    poke(){ if (!adapter) return; clearTimeout(pokeTimer); pokeTimer = setTimeout(syncNow, 4000); },
    onStatus(fn){ listeners.add(fn); fn(status); return () => listeners.delete(fn); },
    status: () => status,
    adapterName: () => adapter ? adapter.name : null,
    _setAdapter(a){ adapter = a; },
  };
})();
