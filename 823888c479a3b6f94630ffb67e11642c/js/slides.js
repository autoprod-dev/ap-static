/* Slides: own light editor (960x540 stage). Export .pptx with PptxGenJS, import .pptx text/images/shapes with JSZip. */
(function(){
  'use strict';
  const W = 960, H = 540;
  const THEMES = {
    light: {name:'Light', bg:'#ffffff', text:'#1f2937', title:'#111827', accent:'#0f766e'},
    dark:  {name:'Dark', bg:'#111827', text:'#e5e7eb', title:'#ffffff', accent:'#14b8a6'},
    ocean: {name:'Ocean', bg:'#0c4a6e', text:'#e0f2fe', title:'#ffffff', accent:'#38bdf8'},
    warm:  {name:'Warm', bg:'#fff7ed', text:'#431407', title:'#7c2d12', accent:'#ea580c'},
    mint:  {name:'Mint', bg:'#ecfdf5', text:'#064e3b', title:'#065f46', accent:'#10b981'},
  };
  const uid = () => Math.random().toString(36).slice(2, 10);
  function layout(kind){
    if (kind === 'title') return {els: [
      {id: uid(), type: 'text', role: 'title', x: 80, y: 170, w: 800, h: 110, text: 'Click to add title', fontSize: 54, bold: true, align: 'center'},
      {id: uid(), type: 'text', role: 'sub', x: 160, y: 300, w: 640, h: 70, text: 'Subtitle', fontSize: 26, align: 'center'}]};
    if (kind === 'content') return {els: [
      {id: uid(), type: 'text', role: 'title', x: 60, y: 30, w: 840, h: 90, text: 'Slide title', fontSize: 40, bold: true},
      {id: uid(), type: 'text', role: 'body', x: 60, y: 135, w: 840, h: 360, text: 'First point\nSecond point\nThird point', fontSize: 26, bullets: true}]};
    return {els: []};
  }
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

  function elHtml(el, theme, editing){
    const t = THEMES[theme] || THEMES.light;
    const base = `left:${el.x}px;top:${el.y}px;width:${el.w}px;height:${el.h}px;`;
    const font = `font-size:${el.fontSize || 24}px;${el.bold ? 'font-weight:700;' : ''}${el.italic ? 'font-style:italic;' : ''}text-align:${el.align || 'left'};color:${el.color || (el.role === 'title' ? t.title : t.text)};font-family:${el.font || 'Arial, Helvetica, sans-serif'};`;
    const textBody = () => el.bullets && !editing ? (el.text || '').split('\n').map(l => `<div style="padding-left:1.1em;text-indent:-1.1em">${l.trim() ? '•&nbsp;&nbsp;' + esc(l) : '&nbsp;'}</div>`).join('') : esc(el.text || '');
    if (el.type === 'image') return `<div class="sl-el image" data-id="${el.id}" style="${base}"><img src="${el.src}" alt=""></div>`;
    if (el.type === 'shape') {
      const r = el.shape === 'ellipse' ? 'border-radius:50%;' : el.shape === 'roundRect' ? 'border-radius:18px;' : '';
      return `<div class="sl-el shape" data-id="${el.id}" style="${base}"><div class="shp" style="background:${el.fill || t.accent};${r}${font}${el.color ? '' : 'color:#fff;'}">${textBody()}</div></div>`;
    }
    return `<div class="sl-el text" data-id="${el.id}" style="${base}${font}">${textBody()}</div>`;
  }
  function slideHtml(slide, theme, editingId, extra){
    const t = THEMES[theme] || THEMES.light;
    return `<div class="slide" style="background:${slide.bg || t.bg};${extra || ''}">${slide.els.map(e => elHtml(e, theme, e.id === editingId)).join('')}</div>`;
  }

  /* ---------- PPTX export ---------- */
  async function exportPptx(model){
    await MO.loadScript('vendor/pptxgen.bundle.js');
    const pptx = new PptxGenJS(); pptx.layout = 'LAYOUT_16x9'; pptx.author = 'My Office'; pptx.title = model.title || 'Presentation';
    const t = THEMES[model.theme] || THEMES.light;
    const inch = px => +(px / 96).toFixed(3), hex = c => (c || '').replace('#', '').slice(0, 6) || '000000';
    for (const s of model.slides) {
      const sl = pptx.addSlide(); sl.background = {color: hex(s.bg || t.bg)};
      for (const el of s.els) {
        const pos = {x: inch(el.x), y: inch(el.y), w: inch(el.w), h: inch(el.h)};
        const txt = {fontSize: Math.round((el.fontSize || 24) * 0.75), bold: !!el.bold, italic: !!el.italic, align: el.align || 'left', color: hex(el.color || (el.type === 'shape' ? '#ffffff' : el.role === 'title' ? t.title : t.text)), fontFace: 'Arial', valign: el.type === 'shape' ? 'middle' : 'top'};
        const runs = () => el.bullets ? (el.text || '').split('\n').map(l => ({text: l, options: {bullet: true, breakLine: true}})) : (el.text || '');
        if (el.type === 'image') {
          let p = pos;
          if (el.nw && el.nh) { const sc = Math.min(el.w / el.nw, el.h / el.nh); const w = el.nw * sc, h = el.nh * sc; p = {x: inch(el.x + (el.w - w) / 2), y: inch(el.y + (el.h - h) / 2), w: inch(w), h: inch(h)}; }
          sl.addImage(Object.assign({data: el.src.replace(/^data:/, '')}, p));
        } else if (el.type === 'shape') {
          const shp = el.shape === 'ellipse' ? pptx.ShapeType.ellipse : el.shape === 'roundRect' ? pptx.ShapeType.roundRect : pptx.ShapeType.rect;
          sl.addText(runs(), Object.assign({shape: shp, fill: {color: hex(el.fill || t.accent)}, line: {color: hex(el.fill || t.accent)}}, pos, txt, el.shape === 'roundRect' ? {rectRadius: 0.15} : {}));
        } else {
          sl.addText(runs(), Object.assign({}, pos, txt, {margin: 5}));
        }
      }
    }
    return pptx.write({outputType: 'blob'});
  }

  /* ---------- PPTX import ---------- */
  async function importPptx(buf){
    await MO.loadScript('vendor/jszip.min.js');
    const zip = await JSZip.loadAsync(buf);
    const xml = async p => { const f = zip.file(p); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
    const q = (n, tag) => n ? [...n.getElementsByTagNameNS('*', tag)] : [];
    const q1 = (n, tag) => q(n, tag)[0] || null;
    const kids = (n, tag) => n ? [...n.children].filter(c => c.localName === tag) : [];
    const kid = (n, tag) => kids(n, tag)[0] || null;
    const resolve = (base, target) => { const parts = base.split('/'); parts.pop(); for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg !== '.') parts.push(seg); } return parts.join('/'); };
    async function rels(path){
      const p = path.replace(/([^/]+)$/, '_rels/$1.rels'); const d = await xml(p); const m = {};
      for (const r of q(d, 'Relationship')) m[r.getAttribute('Id')] = {target: r.getAttribute('TargetMode') === 'External' ? null : resolve(path, r.getAttribute('Target')), type: r.getAttribute('Type')};
      return m;
    }
    const pres = await xml('ppt/presentation.xml'); if (!pres) throw new Error('Not a PowerPoint file');
    const sz = q1(pres, 'sldSz'); const cx = +(sz && sz.getAttribute('cx')) || 12192000, cy = +(sz && sz.getAttribute('cy')) || 6858000;
    const sx = W / cx, sy = H / cy, fpt = 12700 * W / cx; // EMU->px, pt->px
    const prel = await rels('ppt/presentation.xml');
    const slidePaths = q(pres, 'sldId').map(s => prel[s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id') || s.getAttribute('r:id')]).filter(Boolean).map(r => r.target);
    const mimeOf = p => ({png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif', bmp:'image/bmp', svg:'image/svg+xml', webp:'image/webp'})[(p.split('.').pop() || '').toLowerCase()];
    const colour = n => { const c = n && q1(n, 'srgbClr'); return c ? '#' + c.getAttribute('val').toLowerCase() : null; };
    function xfrmOf(sp){ const spPr = kid(sp, 'spPr') || kid(sp, 'grpSpPr'); const x = spPr && kid(spPr, 'xfrm'); if (!x) return null; const off = kid(x, 'off'), ext = kid(x, 'ext'); if (!off || !ext) return null;
      const o = {x: +off.getAttribute('x'), y: +off.getAttribute('y'), w: +ext.getAttribute('cx'), h: +ext.getAttribute('cy')};
      const choff = kid(x, 'chOff'), chext = kid(x, 'chExt'); if (choff && chext) { o.chx = +choff.getAttribute('x'); o.chy = +choff.getAttribute('y'); o.chw = +chext.getAttribute('cx') || 1; o.chh = +chext.getAttribute('cy') || 1; }
      return o; }
    function phOf(sp){ const ph = q1(kid(sp, 'nvSpPr') || kid(sp, 'nvPicPr'), 'ph'); return ph ? {type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx')} : null; }
    function findPh(doc, ph){ if (!doc || !ph) return null; const sps = q(doc, 'sp');
      return sps.find(s => { const p = phOf(s); return p && ph.idx != null && p.idx === ph.idx; }) || sps.find(s => { const p = phOf(s); return p && p.type === ph.type; }) || (ph.type === 'ctrTitle' ? sps.find(s => { const p = phOf(s); return p && p.type === 'title'; }) : null) || (ph.type === 'subTitle' || ph.type === 'obj' ? sps.find(s => { const p = phOf(s); return p && p.type === 'body'; }) : null); }
    const slides = [];
    for (const sp of slidePaths) {
      const doc = await xml(sp); if (!doc) continue;
      const srel = await rels(sp);
      const layRel = Object.values(srel).find(r => /slideLayout$/.test(r.type));
      const lay = layRel ? await xml(layRel.target) : null;
      let master = null; if (layRel) { const lr = await rels(layRel.target); const mr = Object.values(lr).find(r => /slideMaster$/.test(r.type)); if (mr) master = await xml(mr.target); }
      const slide = {els: []};
      const bg = q1(q1(doc, 'bg'), 'solidFill'); if (bg && colour(bg)) slide.bg = colour(bg);
      async function walk(container, tf){
        for (const n of container.children) {
          const ln = n.localName;
          if (ln === 'grpSp') { const g = xfrmOf(n); const map = g && g.chw ? (p => ({x: g.x + (p.x - g.chx) * g.w / g.chw, y: g.y + (p.y - g.chy) * g.h / g.chh, w: p.w * g.w / g.chw, h: p.h * g.h / g.chh})) : (p => p); await walk(n, p => tf(map(p))); continue; }
          if (ln !== 'sp' && ln !== 'pic') continue;
          const ph = phOf(n);
          let x = xfrmOf(n);
          if (!x && ph) { const l = findPh(lay, ph); x = l && xfrmOf(l); if (!x) { const m = findPh(master, ph); x = m && xfrmOf(m); } }
          if (!x) x = ph && /title/i.test(ph.type) ? {x: cx * 0.07, y: cy * 0.05, w: cx * 0.86, h: cy * 0.18} : {x: cx * 0.07, y: cy * 0.27, w: cx * 0.86, h: cy * 0.62};
          x = tf(x);
          const pos = {x: Math.round(x.x * sx), y: Math.round(x.y * sy), w: Math.max(10, Math.round(x.w * sx)), h: Math.max(10, Math.round(x.h * sy))};
          if (ln === 'pic') {
            const blip = q1(n, 'blip'); const rid = blip && (blip.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'embed') || blip.getAttribute('r:embed'));
            const r = rid && srel[rid]; const mime = r && r.target && mimeOf(r.target); const f = r && r.target && zip.file(r.target);
            if (f && mime) slide.els.push(Object.assign({id: uid(), type: 'image', src: `data:${mime};base64,${await f.async('base64')}`}, pos));
            continue;
          }
          const body = kid(n, 'txBody');
          const paras = body ? kids(body, 'p') : [];
          const lines = paras.map(p => [...p.children].map(c => c.localName === 'r' || c.localName === 'fld' ? (q1(c, 't') || {}).textContent || '' : c.localName === 'br' ? '\n' : '').join(''));
          const text = lines.join('\n').replace(/\s+$/, '');
          const rPr = body && (q1(body, 'rPr') || q1(body, 'defRPr'));
          let pt = rPr && rPr.getAttribute('sz') ? +rPr.getAttribute('sz') / 100 : null;
          const isTitle = ph && /title/i.test(ph.type) && ph.type !== 'subTitle';
          if (!pt) pt = isTitle ? 40 : ph ? 24 : 18;
          const algn = (q1(body, 'pPr') || {getAttribute: () => null}).getAttribute('algn');
          const el = Object.assign({id: uid(), text, fontSize: Math.max(8, Math.round(pt * fpt)), bold: !!(rPr && rPr.getAttribute('b') === '1') || (isTitle && !(rPr && rPr.getAttribute('b') === '0')), italic: !!(rPr && rPr.getAttribute('i') === '1'), align: algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' ? 'left' : (ph && /ctrTitle|subTitle/.test(ph.type) ? 'center' : 'left')}, pos);
          const fc = rPr && colour(kid(rPr, 'solidFill')); if (fc) el.color = fc;
          if (isTitle) el.role = 'title';
          if (ph && (ph.type === 'body' || ph.type === 'obj') && !q1(body, 'buNone') && lines.length > 0) el.bullets = true;
          const spPr = kid(n, 'spPr'); const geom = q1(spPr, 'prstGeom'); const fill = spPr && kid(spPr, 'solidFill');
          if (fill && colour(fill)) { el.type = 'shape'; el.fill = colour(fill); const g = geom && geom.getAttribute('prst'); el.shape = g === 'ellipse' ? 'ellipse' : g === 'roundRect' ? 'roundRect' : 'rect'; }
          else { if (!text) continue; el.type = 'text'; }
          slide.els.push(el);
        }
      }
      const tree = q1(doc, 'spTree'); if (tree) await walk(tree, p => p);
      slides.push(slide);
    }
    if (!slides.length) slides.push(layout('title'));
    return {theme: 'light', slides};
  }

  /* ---------- editor ---------- */
  const app = {
    libs: [],
    blank: () => ({theme: 'light', slides: [layout('title')]}),
    importFile: buf => importPptx(buf),
    async mount(host, file, ctx){
      const model = JSON.parse(JSON.stringify(file.content || app.blank()));
      if (!model.slides || !model.slides.length) model.slides = [layout('title')];
      let cur = 0, selId = null, editId = null, scale = 1;
      const undo = [], redo = [];
      host.innerHTML = `
        <div class="toolbar" role="toolbar" aria-label="Slide tools">
          <button data-a="undo" title="Undo">↶</button><button data-a="redo" title="Redo">↷</button><span class="sep"></span>
          <select data-new title="New slide"><option value="">＋ Slide</option><option value="content">Title and content</option><option value="title">Title slide</option><option value="blank">Blank</option></select>
          <button data-a="dup" title="Duplicate slide"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="2" width="10" height="10" rx="2" fill="none" stroke="currentColor" stroke-width="1.6"/><rect x="6" y="6" width="10" height="10" rx="2" fill="#fff" stroke="currentColor" stroke-width="1.6"/></svg></button><button data-a="up" title="Move slide up">▲</button><button data-a="down" title="Move slide down">▼</button><button data-a="delslide" title="Delete slide">🗑</button><span class="sep"></span>
          <button data-a="text" title="Add text box">T Text</button>
          <label class="tbb" title="Add picture">🖼 Picture<input type="file" accept="image/*" data-img hidden></label>
          <select data-shape title="Add shape"><option value="">◼ Shape</option><option value="rect">Rectangle</option><option value="roundRect">Rounded rectangle</option><option value="ellipse">Oval</option></select>
          <select data-theme title="Theme">${Object.entries(THEMES).map(([k, t]) => `<option value="${k}">Theme: ${t.name}</option>`).join('')}</select><span class="sep"></span>
          <select data-size title="Font size">${[12,14,16,18,20,24,28,32,40,48,54,64,80].map(s => `<option>${s}</option>`).join('')}</select>
          <button data-a="bold" title="Bold"><b>B</b></button><button data-a="italic" title="Italic"><i>I</i></button><button data-a="bullets" title="Bullets">•≡</button>
          <button data-a="left" title="Align left"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button><button data-a="center" title="Centre"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="4.5" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="4.5" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button><button data-a="right" title="Align right"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="7" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="7" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button>
          <label class="tbb" title="Text colour">A<input type="color" data-color="color" value="#1f2937"></label>
          <label class="tbb" title="Shape fill">▉<input type="color" data-color="fill" value="#0f766e"></label>
          <button data-a="front" title="Bring to front">⬆︎</button><button data-a="back" title="Send to back">⬇︎</button><button data-a="del" title="Delete selected">✕</button><span class="sep"></span>
          <button data-a="present" title="Present">▶ Present</button>
        </div>
        <div class="sl-body"><div class="sl-thumbs"></div><div class="sl-stage-wrap"><div class="sl-stage-box"></div></div></div>`;
      const tb = host.querySelector('.toolbar'), thumbs = host.querySelector('.sl-thumbs'), stageWrap = host.querySelector('.sl-stage-wrap'), box = host.querySelector('.sl-stage-box');
      tb.querySelector('[data-theme]').value = model.theme || 'light';
      const slide = () => model.slides[cur];
      const sel = () => selId && slide().els.find(e => e.id === selId);
      const snapshot = () => { undo.push(JSON.stringify(model)); if (undo.length > 40) undo.shift(); redo.length = 0; };
      let thumbTimer;
      const changed = (thumbOnly) => { ctx.changed(); clearTimeout(thumbTimer); thumbTimer = setTimeout(() => renderThumb(cur), 300); };

      function fit(){
        const r = stageWrap.getBoundingClientRect(); const pad = innerWidth < 700 ? 12 : 32;
        scale = Math.max(0.1, Math.min((r.width - pad) / W, (r.height - pad) / H));
        box.style.width = W * scale + 'px'; box.style.height = H * scale + 'px';
        const s = box.querySelector('.slide'); if (s) s.style.transform = `scale(${scale})`;
      }
      function renderStage(){
        box.innerHTML = slideHtml(slide(), model.theme, editId);
        const s = box.querySelector('.slide'); s.style.transform = `scale(${scale})`;
        if (selId) { const n = s.querySelector(`[data-id="${selId}"]`); if (n) { n.classList.add('sel'); n.insertAdjacentHTML('beforeend', '<div class="handle" data-handle></div>'); } else selId = null; }
        if (editId) { const n = s.querySelector(`[data-id="${editId}"]`); const tgt = n && (n.querySelector('.shp') || n);
          if (tgt) { try { tgt.contentEditable = 'plaintext-only'; } catch (e) { tgt.contentEditable = 'true'; } if (tgt.contentEditable !== 'plaintext-only') tgt.contentEditable = 'true'; tgt.focus(); const rg = document.createRange(); rg.selectNodeContents(tgt); rg.collapse(false); const ss = getSelection(); ss.removeAllRanges(); ss.addRange(rg);
            tgt.addEventListener('input', () => { const el = slide().els.find(e => e.id === editId); if (el) { el.text = tgt.innerText.replace(/\n$/, ''); changed(); } }); } }
        syncProps();
      }
      function renderThumb(i){
        const b = thumbs.querySelector(`[data-i="${i}"] .mini`); if (!b) return renderThumbs();
        b.innerHTML = slideHtml(model.slides[i], model.theme, null, 'transform:none;box-shadow:none');
      }
      function renderThumbs(){
        thumbs.innerHTML = model.slides.map((s, i) => `<button class="sl-thumb ${i === cur ? 'active' : ''}" data-i="${i}" aria-label="Slide ${i + 1}"><span class="num">${i + 1}</span><div class="mini"></div></button>`).join('');
        const tw = thumbs.querySelector('.sl-thumb'); const sc = tw ? tw.clientWidth / W : 0.15;
        thumbs.querySelectorAll('.mini').forEach((m, i) => { m.style.transform = `scale(${sc})`; renderThumb(i); });
      }
      function go(i){ editId = null; selId = null; cur = Math.max(0, Math.min(i, model.slides.length - 1)); renderThumbs(); renderStage(); const a = thumbs.querySelector('.active'); a && a.scrollIntoView({block: 'nearest', inline: 'nearest'}); }
      function syncProps(){
        const el = sel();
        tb.querySelector('[data-size]').value = String(el && el.fontSize || 24);
        ['bold', 'italic', 'bullets'].forEach(k => tb.querySelector(`[data-a="${k}"]`).classList.toggle('active', !!(el && el[k])));
      }
      function addEl(el){ snapshot(); slide().els.push(el); selId = el.id; editId = null; renderStage(); changed(); }
      function stopEdit(){ if (editId) { editId = null; renderStage(); } }

      // pointer: select, move, resize, tap-again-to-edit
      stageWrap.addEventListener('pointerdown', e => {
        const handle = e.target.closest('[data-handle]'); const node = e.target.closest('.sl-el');
        if (editId && node && node.dataset.id === editId) return; // typing inside the box
        if (!node) { if (selId || editId) { selId = null; editId = null; renderStage(); } return; }
        e.preventDefault();
        const id = node.dataset.id, el = slide().els.find(x => x.id === id); if (!el) return;
        const wasSel = selId === id; if (editId) editId = null;
        selId = id; if (!wasSel || editId === null) renderStage();
        const sx = e.clientX, sy = e.clientY, ox = el.x, oy = el.y, ow = el.w, oh = el.h; let moved = false, snap = false;
        const dom = () => box.querySelector(`[data-id="${id}"]`);
        const mv = ev => {
          const dx = (ev.clientX - sx) / scale, dy = (ev.clientY - sy) / scale;
          if (!moved && Math.abs(dx) + Math.abs(dy) < 4 / scale) return;
          if (!snap) { snapshot(); snap = true; } moved = true;
          if (handle) { el.w = Math.max(20, Math.round(ow + dx)); el.h = Math.max(20, Math.round(oh + dy)); if (el.type === 'image' && ev.shiftKey && el.nw) el.h = Math.round(el.w * el.nh / el.nw); }
          else { el.x = Math.round(Math.max(-el.w + 20, Math.min(W - 20, ox + dx))); el.y = Math.round(Math.max(-el.h + 20, Math.min(H - 20, oy + dy))); }
          const d = dom(); if (d) { d.style.left = el.x + 'px'; d.style.top = el.y + 'px'; d.style.width = el.w + 'px'; d.style.height = el.h + 'px'; }
        };
        const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up);
          if (moved) changed(); else if (wasSel && el.type !== 'image' && !handle) { editId = id; renderStage(); } };
        window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
      });
      stageWrap.addEventListener('dblclick', e => { const node = e.target.closest('.sl-el'); if (node && !editId) { const el = slide().els.find(x => x.id === node.dataset.id); if (el && el.type !== 'image') { selId = editId = el.id; renderStage(); } } });
      thumbs.addEventListener('click', e => { const b = e.target.closest('.sl-thumb'); if (b) go(+b.dataset.i); });
      const onKey = e => {
        if (!host.isConnected || document.querySelector('.present')) return;
        if (editId) { if (e.key === 'Escape') stopEdit(); return; }
        if (/INPUT|SELECT|TEXTAREA/.test(document.activeElement.tagName)) return;
        const el = sel();
        if ((e.key === 'Delete' || e.key === 'Backspace') && el) { e.preventDefault(); snapshot(); slide().els = slide().els.filter(x => x !== el); selId = null; renderStage(); changed(); }
        else if (el && /^Arrow/.test(e.key)) { e.preventDefault(); const d = e.shiftKey ? 10 : 2; snapshot(); if (e.key === 'ArrowLeft') el.x -= d; if (e.key === 'ArrowRight') el.x += d; if (e.key === 'ArrowUp') el.y -= d; if (e.key === 'ArrowDown') el.y += d; renderStage(); changed(); }
        else if (!el && (e.key === 'PageDown' || e.key === 'ArrowDown')) go(cur + 1);
        else if (!el && (e.key === 'PageUp' || e.key === 'ArrowUp')) go(cur - 1);
        else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); doUndo(e.shiftKey); }
        else if (e.key === 'F5') { e.preventDefault(); present(); }
      };
      document.addEventListener('keydown', onKey);
      function doUndo(isRedo){ const from = isRedo ? redo : undo, to = isRedo ? undo : redo; if (!from.length) return; to.push(JSON.stringify(model)); const p = JSON.parse(from.pop()); model.slides = p.slides; model.theme = p.theme; tb.querySelector('[data-theme]').value = model.theme; editId = selId = null; cur = Math.min(cur, model.slides.length - 1); renderThumbs(); renderStage(); ctx.changed(); }
      function setProp(fn){ const el = sel(); if (!el) return ctx.toast('Select a text box or shape first.'); snapshot(); fn(el); renderStage(); changed(); }
      tb.addEventListener('mousedown', e => { if (e.target.closest('button') && editId) e.preventDefault(); });
      tb.addEventListener('click', e => {
        const b = e.target.closest('button'); if (!b) return; const a = b.dataset.a;
        if (a === 'undo') return doUndo(false); if (a === 'redo') return doUndo(true);
        if (a === 'dup') { snapshot(); const c = JSON.parse(JSON.stringify(slide())); c.els.forEach(x => x.id = uid()); model.slides.splice(cur + 1, 0, c); go(cur + 1); changed(); }
        else if (a === 'up' && cur > 0) { snapshot(); const [s] = model.slides.splice(cur, 1); model.slides.splice(cur - 1, 0, s); go(cur - 1); changed(); }
        else if (a === 'down' && cur < model.slides.length - 1) { snapshot(); const [s] = model.slides.splice(cur, 1); model.slides.splice(cur + 1, 0, s); go(cur + 1); changed(); }
        else if (a === 'delslide') { if (model.slides.length < 2) return ctx.toast('A presentation needs at least one slide.'); snapshot(); model.slides.splice(cur, 1); go(Math.min(cur, model.slides.length - 1)); changed(); }
        else if (a === 'text') addEl({id: uid(), type: 'text', x: 280, y: 220, w: 400, h: 80, text: 'New text', fontSize: 28});
        else if (a === 'bold' || a === 'italic' || a === 'bullets') setProp(el => el[a] = !el[a]);
        else if (a === 'left' || a === 'center' || a === 'right') setProp(el => el.align = a);
        else if (a === 'front') setProp(el => { const s = slide(); s.els = s.els.filter(x => x !== el).concat(el); });
        else if (a === 'back') setProp(el => { const s = slide(); s.els = [el].concat(s.els.filter(x => x !== el)); });
        else if (a === 'del') setProp(el => { slide().els = slide().els.filter(x => x !== el); selId = null; });
        else if (a === 'present') present();
      });
      tb.querySelector('[data-new]').onchange = e => { const v = e.target.value; e.target.value = ''; if (!v) return; snapshot(); model.slides.splice(cur + 1, 0, layout(v)); go(cur + 1); changed(); };
      tb.querySelector('[data-shape]').onchange = e => { const v = e.target.value; e.target.value = ''; if (!v) return; addEl({id: uid(), type: 'shape', shape: v, x: 360, y: 180, w: 240, h: 160, text: '', fontSize: 24, fill: THEMES[model.theme || 'light'].accent}); };
      tb.querySelector('[data-theme]').onchange = e => { snapshot(); model.theme = e.target.value; renderStage(); renderThumbs(); ctx.changed(); };
      tb.querySelector('[data-size]').onchange = e => setProp(el => el.fontSize = +e.target.value);
      tb.querySelectorAll('[data-color]').forEach(inp => inp.onchange = e => setProp(el => el[inp.dataset.color] = e.target.value));
      tb.querySelector('[data-img]').onchange = e => {
        const f = e.target.files[0]; e.target.value = ''; if (!f) return;
        const img = new Image(), url = URL.createObjectURL(f);
        img.onload = () => { const max = 1600, sc = Math.min(1, max / Math.max(img.width, img.height)); const c = document.createElement('canvas'); c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc); c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
          const src = c.toDataURL(f.type === 'image/png' ? 'image/png' : 'image/jpeg', 0.85); const fitSc = Math.min(600 / c.width, 400 / c.height, 1); const w = Math.round(c.width * fitSc), h = Math.round(c.height * fitSc);
          addEl({id: uid(), type: 'image', src, nw: c.width, nh: c.height, x: Math.round((W - w) / 2), y: Math.round((H - h) / 2), w, h}); };
        img.src = url;
      };

      function present(){
        stopEdit(); let i = cur;
        const ov = document.createElement('div'); ov.className = 'present'; ov.tabIndex = 0;
        ov.innerHTML = '<div class="pbox"></div><span class="pcount"></span><button class="pclose" aria-label="Exit">✕</button>';
        document.body.appendChild(ov); const pb = ov.querySelector('.pbox');
        const draw = () => { const sc = Math.min(innerWidth / W, innerHeight / H); pb.style.width = W * sc + 'px'; pb.style.height = H * sc + 'px'; pb.style.position = 'relative'; pb.innerHTML = slideHtml(model.slides[i], model.theme); pb.querySelector('.slide').style.transform = `scale(${sc})`; ov.querySelector('.pcount').textContent = `${i + 1} / ${model.slides.length}`; };
        const next = () => { if (i < model.slides.length - 1) { i++; draw(); } else close(); };
        const prev = () => { if (i > 0) { i--; draw(); } };
        const key = e => { if (['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'].includes(e.key)) { e.preventDefault(); next(); } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) { e.preventDefault(); prev(); } else if (e.key === 'Escape') close(); };
        let tx = null;
        ov.addEventListener('pointerdown', e => { tx = e.clientX; });
        ov.addEventListener('pointerup', e => { if (e.target.closest('.pclose')) return close(); if (tx == null) return; const dx = e.clientX - tx; tx = null; if (Math.abs(dx) > 50) { dx < 0 ? next() : prev(); } else if (e.clientX > innerWidth / 3) next(); else prev(); });
        function close(){ document.removeEventListener('keydown', key); removeEventListener('resize', draw); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); ov.remove(); go(i); }
        document.addEventListener('keydown', key); addEventListener('resize', draw);
        if (ov.requestFullscreen) ov.requestFullscreen().catch(() => {});
        draw(); ov.focus();
      }

      const ro = new ResizeObserver(() => { fit(); });
      ro.observe(stageWrap);
      renderThumbs(); renderStage(); fit();
      return {
        getContent: () => JSON.parse(JSON.stringify(model)),
        exportBlob: () => exportPptx(Object.assign({title: file.name}, model)),
        printView: () => ({html: model.slides.map(s => `<div class="print-slide">${slideHtml(s, model.theme, null, 'transform:none;box-shadow:none')}</div>`).join(''), page: 'size:254mm 142.875mm;margin:0', css: '@media print{.print-slide *{-webkit-print-color-adjust:exact;print-color-adjust:exact}}'}),
        present,
        destroy(){ ro.disconnect(); document.removeEventListener('keydown', onKey); },
      };
    }
  };
  MOApps.slides = app;
})();
