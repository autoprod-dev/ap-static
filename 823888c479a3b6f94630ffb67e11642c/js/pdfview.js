/* PDF viewer with pdf.js. Pages render lazily and far-away pages are released to keep memory low. */
(function(){
  'use strict';
  async function loadPdfjs(){
    await MO.loadScript('vendor/pdf.js');
    const lib = window.pdfjsLib;
    if (!lib.__moReady) {
      if (/^https?:$/.test(location.protocol) && typeof Worker !== 'undefined') {
        lib.GlobalWorkerOptions.workerSrc = new URL('vendor/pdf.worker.min.mjs', location.href).href;
      } else {
        await MO.loadScript('vendor/pdf.worker.iife.js'); // file:// fallback: runs on the main thread
        lib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.iife.js';
      }
      lib.__moReady = true;
    }
    return lib;
  }
  const app = {
    libs: [],
    blank: () => null,
    async mount(host, file, ctx){
      const lib = await loadPdfjs();
      host.innerHTML = `<div class="toolbar"><button data-a="out" title="Zoom out">－</button><button data-a="fit" title="Fit width">Fit</button><button data-a="in" title="Zoom in">＋</button><span class="sep"></span><span class="pdf-info muted small" style="color:inherit"></span></div><div class="pdf-scroll"></div>`;
      const scroller = host.querySelector('.pdf-scroll'), info = host.querySelector('.pdf-info');
      const data = new Uint8Array(await file.blob.arrayBuffer());
      const task = lib.getDocument({data, isEvalSupported: false, disableFontFace: false});
      const pdf = await task.promise;
      const first = await pdf.getPage(1); const vp1 = first.getViewport({scale: 1});
      let zoom = null; const pages = [];
      const fitZoom = () => Math.min(2.5, Math.max(0.3, (scroller.clientWidth - 16) / vp1.width));
      const dpr = Math.min(window.devicePixelRatio || 1, 1.5);
      info.textContent = `${pdf.numPages} page${pdf.numPages === 1 ? '' : 's'}`;
      function layout(){
        const z = zoom || fitZoom();
        scroller.innerHTML = '';
        pages.length = 0;
        for (let i = 1; i <= pdf.numPages; i++) {
          const d = document.createElement('div'); d.className = 'pdf-page'; d.dataset.p = i;
          d.style.width = Math.floor(vp1.width * z) + 'px'; d.style.height = Math.floor(vp1.height * z) + 'px';
          scroller.appendChild(d); pages.push({div: d, rendered: false, task: null});
          io.observe(d);
        }
      }
      async function render(i){
        const p = pages[i - 1]; if (!p || p.rendered) return; p.rendered = true;
        const page = await pdf.getPage(i); const z = zoom || fitZoom();
        const vp = page.getViewport({scale: z * dpr});
        p.div.style.height = Math.floor(vp.height / dpr) + 'px';
        const c = document.createElement('canvas'); c.width = Math.floor(vp.width); c.height = Math.floor(vp.height);
        p.div.innerHTML = ''; p.div.appendChild(c);
        p.task = page.render({canvasContext: c.getContext('2d'), viewport: vp});
        try { await p.task.promise; } catch (e) { if (e && e.name !== 'RenderingCancelledException') console.warn(e); }
        page.cleanup();
      }
      function release(i){ const p = pages[i - 1]; if (!p || !p.rendered) return; if (p.task) p.task.cancel(); const c = p.div.querySelector('canvas'); if (c) { c.width = c.height = 0; } p.div.innerHTML = ''; p.rendered = false; }
      const io = new IntersectionObserver(entries => {
        for (const en of entries) { const i = +en.target.dataset.p; if (en.isIntersecting) render(i); }
        // release pages far from view
        const vis = entries.filter(e => e.isIntersecting).map(e => +e.target.dataset.p);
        if (vis.length) { const mid = vis[0]; pages.forEach((p, k) => { if (Math.abs(k + 1 - mid) > 6) release(k + 1); }); }
      }, {root: scroller, rootMargin: '600px 0px'});
      host.querySelector('.toolbar').onclick = e => { const b = e.target.closest('button'); if (!b) return; const z = zoom || fitZoom();
        if (b.dataset.a === 'in') zoom = Math.min(4, z * 1.25); else if (b.dataset.a === 'out') zoom = Math.max(0.25, z / 1.25); else zoom = null;
        io.disconnect(); layout(); };
      layout();
      return {
        destroy(){ io.disconnect(); pages.forEach((p, k) => release(k + 1)); pdf.destroy(); },
        _pages: () => pdf.numPages,
      };
    }
  };
  MOApps.pdf = app;
})();
