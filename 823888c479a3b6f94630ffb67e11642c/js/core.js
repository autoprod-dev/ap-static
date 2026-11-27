/* App shell: home, routing, lazy loading, open/save, PWA file handling, settings. */
(function(){
  'use strict';
  const VERSION = '0.1.1';
  const $ = s => document.querySelector(s);
  const KIND = {
    doc:   {label:'Document', badge:'Doc', ext:'.docx', mime:'application/vnd.openxmlformats-officedocument.wordprocessingml.document', script:'js/docs.js'},
    sheet: {label:'Spreadsheet', badge:'Sheet', ext:'.xlsx', mime:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', script:'js/sheets.js'},
    slides:{label:'Presentation', badge:'Slides', ext:'.pptx', mime:'application/vnd.openxmlformats-officedocument.presentationml.presentation', script:'js/slides.js'},
    pdf:   {label:'PDF', badge:'PDF', ext:'.pdf', mime:'application/pdf', script:'js/pdfview.js'},
  };
  const EXT_KIND = {docx:'doc', xlsx:'sheet', csv:'sheet', pptx:'slides', pdf:'pdf'};
  window.MOApps = window.MOApps || {};

  /* ---------- utilities ---------- */
  const loaded = new Map();
  function loadScript(src){
    if (!loaded.has(src)) loaded.set(src, new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = src; s.async = true;
      s.onload = () => res(); s.onerror = () => { loaded.delete(src); rej(new Error('Could not load ' + src)); };
      document.head.appendChild(s);
    }));
    return loaded.get(src);
  }
  async function loadScripts(list){ for (const s of list) await loadScript(s); }
  let toastTimer;
  function toast(msg, ms){ const t = $('#toast'); t.textContent = msg; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(() => t.hidden = true, ms || 2600); }
  function busy(on, text){ $('#loading').hidden = !on; $('#loading-text').textContent = text || 'Loading…'; }
  function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function when(ts){
    const d = new Date(ts), now = new Date();
    if (d.toDateString() === now.toDateString()) return 'Today ' + d.toLocaleTimeString('en-AU', {hour:'numeric', minute:'2-digit'});
    return d.toLocaleDateString('en-AU', {day:'numeric', month:'short', year: d.getFullYear() === now.getFullYear() ? undefined : 'numeric'});
  }
  function baseName(n){ return String(n || 'Untitled').replace(/\.(docx|xlsx|pptx|pdf|csv)$/i, ''); }
  function download(blob, name){
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  async function appFor(kind){
    await loadScript(KIND[kind].script);
    const app = MOApps[kind];
    if (app.libs) await loadScripts(app.libs);
    return app;
  }

  /* ---------- home ---------- */
  async function renderHome(){
    const q = ($('#recent-search').value || '').toLowerCase();
    const files = (await MOStore.recent()).filter(f => !q || f.name.toLowerCase().includes(q));
    const ul = $('#recent-list');
    ul.innerHTML = files.map(f => `<li data-id="${f.id}">
      <button class="open-area" data-open="${f.id}"><span class="ri ${f.kind}">${KIND[f.kind].badge}</span>
      <span class="rtext"><div class="rname">${esc(f.name)}</div><div class="rmeta">${KIND[f.kind].label} · ${when(f.updatedAt)}${f.dirty && MOSync.adapterName() ? ' · not synced yet' : ''}</div></span></button>
      <button class="icon-btn" data-more="${f.id}" aria-label="More for ${esc(f.name)}">⋮</button></li>`).join('');
    ul.hidden = !files.length; $('#recent-empty').hidden = !!files.length;
  }
  function showView(id){ for (const v of document.querySelectorAll('.view')) v.hidden = v.id !== id; }

  /* ---------- editor shell ---------- */
  let current = null; // {file, app, inst, saveTimer}
  async function openFile(id){
    const file = await MOStore.get(id);
    if (!file || file.deleted) { toast('That file is no longer here.'); location.hash = ''; return; }
    await closeEditor();
    busy(true, 'Opening ' + file.name + '…');
    try {
      const app = await appFor(file.kind);
      showView('editor');
      $('#kind-badge').textContent = KIND[file.kind].badge; $('#kind-badge').className = 'kind-badge ' + file.kind;
      $('#title-input').value = file.name; $('#save-state').textContent = 'Saved on this device';
      $('#save-btn').textContent = file.kind === 'pdf' ? 'Download' : 'Save';
      const host = $('#editor-host'); host.innerHTML = '';
      current = {file, app, inst:null, saveTimer:null};
      current.inst = await app.mount(host, file, {changed: scheduleSave, toast, busy, loadScripts, esc, download});
      document.title = file.name + ' – My Office';
    } catch (e) {
      console.error(e); toast('Could not open: ' + e.message, 5000); location.hash = '';
    } finally { busy(false); }
  }
  function scheduleSave(){
    if (!current) return;
    $('#save-state').textContent = 'Editing…';
    clearTimeout(current.saveTimer);
    current.saveTimer = setTimeout(saveNow, 700);
  }
  async function saveNow(){
    if (!current || !current.inst || !current.inst.getContent) return;
    clearTimeout(current.saveTimer); current.saveTimer = null;
    const content = current.inst.getContent();
    current.file = await MOStore.save(current.file.id, {content});
    $('#save-state').textContent = 'Saved on this device';
  }
  async function closeEditor(){
    if (!current) return;
    if (current.saveTimer) await saveNow();
    try { current.inst && current.inst.destroy && current.inst.destroy(); } catch(e){}
    $('#editor-host').innerHTML = ''; current = null; document.title = 'My Office';
  }
  async function exportBlob(){
    if (current.file.kind === 'pdf') return current.file.blob;
    await saveNow();
    return current.inst.exportBlob();
  }
  async function saveToDevice(asNew){
    if (!current) return;
    const k = KIND[current.file.kind];
    busy(true, 'Saving ' + k.ext + '…');
    try {
      const blob = await exportBlob();
      const name = baseName(current.file.name) + k.ext;
      let handle = asNew ? null : await MOStore.getHandle(current.file.id);
      if (handle && handle.createWritable) {
        if (handle.queryPermission && (await handle.queryPermission({mode:'readwrite'})) !== 'granted') {
          if ((await handle.requestPermission({mode:'readwrite'})) !== 'granted') handle = null;
        }
        if (handle) { const w = await handle.createWritable(); await w.write(blob); await w.close(); toast('Saved to ' + handle.name); return; }
      }
      if (window.showSaveFilePicker && window.isSecureContext && window.top === window) {
        try {
          const h = await showSaveFilePicker({suggestedName: name, types:[{description:k.label, accept:{[k.mime]:[k.ext]}}]});
          const w = await h.createWritable(); await w.write(blob); await w.close();
          await MOStore.setHandle(current.file.id, h); toast('Saved to ' + h.name); return;
        } catch (e) { if (e && e.name === 'AbortError') return; }
      }
      download(blob, name); toast('Downloaded ' + name);
    } catch (e) { console.error(e); toast('Save failed: ' + e.message, 5000); }
    finally { busy(false); }
  }
  async function preparePrint(){
    if (!current || !current.inst.printView) return false;
    await saveNow();
    const pv = await current.inst.printView();
    let root = $('#print-root'); if (!root) { root = document.createElement('div'); root.id = 'print-root'; document.body.appendChild(root); }
    root.innerHTML = pv.html;
    let st = $('#print-page-style'); if (!st) { st = document.createElement('style'); st.id = 'print-page-style'; document.head.appendChild(st); }
    st.textContent = `@page{${pv.page || 'size:A4;margin:18mm'}}` + (pv.css || '');
    return true;
  }
  async function exportPdf(){
    if (current.file.kind === 'pdf') return saveToDevice();
    if (!(await preparePrint())) return;
    toast('Choose "Save as PDF" as the printer.', 4000);
    setTimeout(() => window.print(), 50);
  }
  function fileMenu(){
    const m = $('#file-menu');
    const k = current.file.kind;
    const items = [];
    if (k !== 'pdf') items.push(['Save ' + KIND[k].ext + ' to device', () => saveToDevice()], ['Save a copy as…', () => saveToDevice(true)], ['Export PDF', exportPdf]);
    else items.push(['Download PDF', () => saveToDevice()]);
    for (const x of (current.inst.extraMenu || [])) items.push([x.label, x.fn]);
    items.push(null, ['Duplicate', () => duplicate(current.file.id, true)], ['Delete file', () => del(current.file.id)]);
    m.innerHTML = items.map((it, i) => it ? `<button data-i="${i}">${esc(it[0])}</button>` : '<hr>').join('');
    m.onclick = e => { const b = e.target.closest('button'); if (!b) return; m.hidden = true; items[+b.dataset.i][1](); };
    m.hidden = !m.hidden;
  }

  /* ---------- create / import ---------- */
  async function createNew(kind){
    busy(true);
    try {
      const app = await appFor(kind);
      const n = {doc:'Untitled document', sheet:'Untitled spreadsheet', slides:'Untitled presentation'}[kind];
      const f = await MOStore.create(kind, n, app.blank());
      location.hash = '#/f/' + f.id;
    } finally { busy(false); }
  }
  async function importFile(file, handle){
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    const kind = EXT_KIND[ext];
    if (!kind) { toast('This file type is not supported yet: .' + ext); return null; }
    busy(true, 'Opening ' + file.name + '…');
    try {
      const buf = await file.arrayBuffer();
      let existing = null;
      if (handle && handle.isSameEntry) {
        for (const f of await MOStore.recent()) {
          const h = await MOStore.getHandle(f.id);
          if (h && await h.isSameEntry(handle).catch(() => false)) { existing = f; break; }
        }
      }
      let rec;
      if (kind === 'pdf') {
        const blob = new Blob([buf], {type:'application/pdf'});
        rec = existing ? await MOStore.save(existing.id, {blob}) : await MOStore.create('pdf', baseName(file.name), null, {blob, mime:'application/pdf'});
      } else {
        const app = await appFor(kind);
        const content = await app.importFile(buf, file.name);
        rec = existing ? await MOStore.save(existing.id, {content}) : await MOStore.create(kind, baseName(file.name), content);
      }
      if (handle) await MOStore.setHandle(rec.id, handle);
      return rec;
    } catch (e) { console.error(e); toast('Could not open ' + file.name + ': ' + e.message, 6000); return null; }
    finally { busy(false); }
  }
  async function openFromDevice(){
    if (window.showOpenFilePicker && window.isSecureContext && window.top === window) {
      try {
        const hs = await showOpenFilePicker({multiple:true, types:[{description:'Office files and PDFs', accept:{
          [KIND.doc.mime]:['.docx'], [KIND.sheet.mime]:['.xlsx'], [KIND.slides.mime]:['.pptx'], 'application/pdf':['.pdf'], 'text/csv':['.csv']}}]});
        let last = null;
        for (const h of hs) last = await importFile(await h.getFile(), h) || last;
        if (last) location.hash = '#/f/' + last.id; else renderHome();
        return;
      } catch (e) { if (e && e.name === 'AbortError') return; }
    }
    $('#file-input').click();
  }
  async function duplicate(id, open){
    const f = await MOStore.get(id);
    const c = await MOStore.create(f.kind, f.name + ' (copy)', f.content ? JSON.parse(JSON.stringify(f.content)) : null, f.blob ? {blob:f.blob, mime:f.mime} : {});
    if (open) location.hash = '#/f/' + c.id; else renderHome();
  }
  async function del(id){
    const f = await MOStore.get(id);
    if (!confirm(`Delete "${f.name}"? This removes it from My Office${MOSync.adapterName() ? ' on all synced devices' : ''}.`)) return;
    if (current && current.file.id === id) { clearTimeout(current.saveTimer); current.saveTimer = null; }
    await MOStore.remove(id); toast('Deleted'); location.hash = ''; renderHome();
  }
  async function rename(id){
    const f = await MOStore.get(id);
    const n = prompt('New name', f.name); if (!n || !n.trim()) return;
    await MOStore.save(id, {name:n.trim()}); renderHome();
  }
  async function downloadFile(id){
    const f = await MOStore.get(id);
    if (f.kind === 'pdf') return download(f.blob, baseName(f.name) + '.pdf');
    location.hash = '#/f/' + id;
    setTimeout(() => saveToDevice(), 600);
  }

  /* ---------- settings: sync + backup ---------- */
  async function renderSettings(){
    const cfg = (await MOStore.getMeta('sync')) || {mode:'off'};
    const showMock = cfg.mode === 'mock' || new URLSearchParams(location.search).has('mocksync');
    const st = MOSync.status();
    $('#settings-main').innerHTML = `
      <section><h2>Sync between your phone and laptop</h2>
        <p class="muted small">Your files are always kept on this device first, so everything works offline. Sync copies them to private storage and back. Changes made while offline wait in a queue and upload when you are back online. If the same file was changed on two devices, the newest edit wins and the other is kept as a "conflict copy", so nothing is lost.</p>
        <label for="sync-mode">Sync method</label>
        <select id="sync-mode">
          <option value="off">Off (this device only)</option>
          <option value="gdrive">Google Drive hidden app folder (needs one-time setup)</option>
          ${showMock ? '<option value="mock">Test server (for testing only)</option>' : ''}
        </select>
        <div id="gdrive-fields"><label for="client-id">Google sign-in client ID</label>
          <input type="text" id="client-id" placeholder="Not set up yet" value="${esc(cfg.clientId || '')}">
          <p class="muted small">This uses Google's "app data" permission only (drive.appdata): a hidden folder that just this app can see. It can't see, open or change any of your other Drive files. No password or secret is stored here.</p>
          <p class="muted small">When you make the client ID, the "Authorised JavaScript origin" to enter is: <b>${esc(location.origin)}</b></p></div>
        <div id="mock-fields"><label for="mock-base">Test server address</label><input type="text" id="mock-base" value="${esc(cfg.mockBase || new URLSearchParams(location.search).get('mocksync') || '')}"></div>
        <div class="row"><button class="btn primary" id="sync-save">Save sync settings</button><button class="btn" id="sync-signin">Sign in</button><button class="btn" id="sync-now">Sync now</button></div>
        <p class="small" id="sync-status">Status: ${esc(st.text)}${st.error ? ' (' + esc(st.error) + ')' : ''}</p>
      </section>
      <section><h2>Backup (works without any account)</h2>
        <p class="muted small">Download all your files as one backup file, then open it on your other device with Restore. Restore keeps the newest version of each file.</p>
        <div class="row"><button class="btn" id="backup-dl">Download backup</button><button class="btn" id="backup-restore">Restore from backup</button></div>
        <input type="file" id="backup-input" accept=".zip" hidden>
      </section>
      <section><h2>About</h2><p class="small">My Office ${VERSION}. A private office suite for personal use, built from open-source libraries (see LICENCES.md). Not for sale.</p>
      <p class="small muted">Install: in Chrome or Edge choose "Install My Office" from the address bar or menu. After installing, it can be set as the app that opens .docx, .xlsx, .pptx and .pdf files.</p></section>`;
    const mode = $('#sync-mode'); mode.value = cfg.mode || 'off';
    const vis = () => { $('#gdrive-fields').hidden = mode.value !== 'gdrive'; $('#mock-fields').hidden = mode.value !== 'mock'; };
    mode.onchange = vis; vis();
    $('#sync-save').onclick = async () => { await MOSync.configure({mode: mode.value, clientId: $('#client-id').value.trim(), mockBase: $('#mock-base').value.trim().replace(/\/$/, '')}); toast('Sync settings saved'); renderSettings(); };
    $('#sync-signin').onclick = async () => { try { await MOSync.signIn(); toast('Signed in'); } catch (e) { toast('Sign-in failed: ' + e.message, 5000); } renderSettings(); };
    $('#sync-now').onclick = async () => { await MOSync.syncNow(); renderSettings(); };
    $('#backup-dl').onclick = backupDownload;
    $('#backup-restore').onclick = () => $('#backup-input').click();
    $('#backup-input').onchange = e => { const f = e.target.files[0]; e.target.value = ''; if (f) backupRestore(f); };
  }
  function blobToB64(blob){ return new Promise(res => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1]); r.readAsDataURL(blob); }); }
  async function backupDownload(){
    busy(true, 'Making backup…');
    try {
      await loadScript('vendor/jszip.min.js');
      const zip = new JSZip(); const files = await MOStore.recent();
      for (const f of files) {
        const o = Object.assign({}, f); delete o.blob;
        if (f.blob) o.blobB64 = await blobToB64(f.blob);
        zip.file(`files/${f.id}.json`, JSON.stringify(o));
      }
      zip.file('backup.json', JSON.stringify({app:'my-office', version:VERSION, made:new Date().toISOString(), count:files.length}));
      const blob = await zip.generateAsync({type:'blob', compression:'DEFLATE'});
      download(blob, `my-office-backup-${new Date().toISOString().slice(0,10)}.zip`);
    } finally { busy(false); }
  }
  async function backupRestore(file){
    busy(true, 'Restoring…');
    try {
      await loadScript('vendor/jszip.min.js');
      const zip = await JSZip.loadAsync(file); let n = 0;
      for (const name of Object.keys(zip.files).filter(n => n.startsWith('files/'))) {
        const o = JSON.parse(await zip.file(name).async('string'));
        const mine = await MOStore.get(o.id);
        if (mine && mine.updatedAt >= o.updatedAt) continue;
        if (o.blobB64) { const bin = atob(o.blobB64); const u = new Uint8Array(bin.length); for (let i=0;i<bin.length;i++) u[i] = bin.charCodeAt(i); o.blob = new Blob([u], {type:o.mime || 'application/pdf'}); delete o.blobB64; }
        o.dirty = true; await MOStore.rawPut(o); await MOStore.enqueue(o.id); n++;
      }
      toast(`Restored ${n} file${n === 1 ? '' : 's'}`); MOSync.poke();
    } catch (e) { toast('Restore failed: ' + e.message, 5000); }
    finally { busy(false); }
  }

  /* ---------- routing ---------- */
  async function route(){
    const h = location.hash;
    $('#file-menu').hidden = true;
    if (h.startsWith('#/f/')) return openFile(h.slice(4));
    await closeEditor();
    if (h === '#/settings') { showView('settings'); return renderSettings(); }
    showView('home'); renderHome();
  }

  /* ---------- wiring ---------- */
  function wire(){
    document.querySelectorAll('[data-new]').forEach(b => b.onclick = () => createNew(b.dataset.new));
    $('#open-btn').onclick = openFromDevice;
    $('#file-input').onchange = async e => {
      const files = [...e.target.files]; e.target.value = ''; let last = null;
      for (const f of files) last = await importFile(f) || last;
      if (last) location.hash = '#/f/' + last.id; else renderHome();
    };
    $('#recent-search').oninput = renderHome;
    $('#recent-list').onclick = e => {
      const o = e.target.closest('[data-open]'); if (o) { location.hash = '#/f/' + o.dataset.open; return; }
      const m = e.target.closest('[data-more]'); if (m) itemMenu(m);
    };
    $('#back-btn').onclick = () => { location.hash = ''; };
    $('#settings-btn').onclick = () => { location.hash = '#/settings'; };
    $('#sync-chip').onclick = () => { location.hash = '#/settings'; };
    $('#settings-back').onclick = () => { location.hash = ''; };
    $('#save-btn').onclick = () => saveToDevice();
    $('#menu-btn').onclick = e => { e.stopPropagation(); fileMenu(); };
    document.addEventListener('click', e => { if (!e.target.closest('.menu-wrap')) $('#file-menu').hidden = true; const im = document.getElementById('item-menu'); if (im && !e.target.closest('#item-menu') && !e.target.closest('[data-more]')) im.remove(); });
    $('#title-input').onchange = async () => { if (!current) return; const n = $('#title-input').value.trim() || 'Untitled'; current.file = await MOStore.save(current.file.id, {name:n}); };
    document.addEventListener('keydown', e => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && current) { e.preventDefault(); saveToDevice(); }
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'p' && current) { e.preventDefault(); exportPdf(); }
    });
    window.addEventListener('hashchange', route);
    window.addEventListener('mo-files-changed', () => { if (!current) renderHome(); });
    window.addEventListener('pagehide', () => { if (current && current.saveTimer) saveNow(); });
    document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && current && current.saveTimer) saveNow(); });
    MOSync.onStatus(s => { const c = $('#sync-chip'); c.textContent = s.text; c.className = 'chip' + (s.state === 'on' ? ' on' : '') + (s.state === 'error' ? ' err' : ''); const ss = document.getElementById('sync-status'); if (ss) ss.textContent = 'Status: ' + s.text + (s.error ? ' (' + s.error + ')' : ''); });
  }
  function itemMenu(btn){
    const old = document.getElementById('item-menu'); if (old) old.remove();
    const id = btn.dataset.more; const r = btn.getBoundingClientRect();
    const m = document.createElement('div'); m.className = 'menu'; m.id = 'item-menu'; m.style.position = 'fixed';
    m.style.top = Math.min(r.bottom + 4, innerHeight - 240) + 'px'; m.style.right = Math.max(8, innerWidth - r.right) + 'px';
    const items = [['Open', () => location.hash = '#/f/' + id], ['Rename', () => rename(id)], ['Duplicate', () => duplicate(id)], ['Save to device', () => downloadFile(id)], ['Delete', () => del(id)]];
    m.innerHTML = items.map((it, i) => `<button data-i="${i}">${it[0]}</button>`).join('');
    m.onclick = e => { const b = e.target.closest('button'); if (!b) return; m.remove(); items[+b.dataset.i][1](); };
    document.body.appendChild(m);
  }

  /* ---------- PWA: service worker + file handling ---------- */
  function pwa(){
    if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
      navigator.serviceWorker.register('sw.js').then(reg => {
        setTimeout(() => { const sw = reg.active || navigator.serviceWorker.controller; if (sw) sw.postMessage({type:'warm'}); }, 15000);
      }).catch(e => console.warn('SW not registered:', e.message));
    }
    if ('launchQueue' in window) {
      launchQueue.setConsumer(async params => {
        if (!params.files || !params.files.length) return;
        let last = null;
        for (const h of params.files) { try { last = await importFile(await h.getFile(), h) || last; } catch (e) { console.error(e); } }
        if (last) { if (location.hash === '#/f/' + last.id) openFile(last.id); else location.hash = '#/f/' + last.id; }
      });
    }
  }

  window.MO = {VERSION, KIND, toast, busy, loadScript, loadScripts, esc, download, importFile, createNew, openFile, saveNow, exportBlob, preparePrint, current: () => current, renderHome};
  async function boot(){
    wire(); pwa();
    try { await MOSync.init(); } catch (e) { console.warn(e); }
    route();
  }
  boot();
})();
