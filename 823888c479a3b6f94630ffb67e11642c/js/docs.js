/* Documents: rich text with contenteditable. Import .docx with mammoth, export .docx with docx (+docx/html). */
(function(){
  'use strict';
  const DOC_CSS = 'body{font-family:Arial,sans-serif;font-size:11pt;line-height:1.45} h1{font-size:22pt} h2{font-size:16pt} h3{font-size:13pt} p{margin:0 0 .6em} table{border-collapse:collapse;width:100%} td,th{border:1px solid #9ca3af;padding:4px 6px}';
  const FONTS = ['Arial','Calibri','Cambria','Georgia','Times New Roman','Verdana','Tahoma','Trebuchet MS','Courier New'];
  const SIZES = [8,9,10,11,12,14,16,18,20,24,28,36,48,72];

  function sanitize(html){
    const t = document.createElement('template'); t.innerHTML = html;
    t.content.querySelectorAll('script,style,iframe,object,embed,link,meta').forEach(n => n.remove());
    t.content.querySelectorAll('*').forEach(el => { for (const a of [...el.attributes]) { if (/^on/i.test(a.name) || (/^(href|src)$/i.test(a.name) && /^\s*javascript:/i.test(a.value))) el.removeAttribute(a.name); } });
    return t.innerHTML;
  }

  /* Repair the odd DOM that browser editing can leave behind (blocks inside <p>, stray table parts). */
  function normalize(root){
    const BLOCK = /^(UL|OL|TABLE|P|DIV|H[1-6]|BLOCKQUOTE|HR|PRE)$/;
    root.querySelectorAll('tbody, thead, tr').forEach(n => { if (!n.closest('table')) { const t = document.createElement('table'); n.before(t); t.appendChild(n.nodeName === 'TR' ? (() => { const tb = document.createElement('tbody'); tb.appendChild(n); return tb; })() : n); } });
    root.querySelectorAll('table').forEach(t => { if (!t.querySelector('td,th')) t.remove(); });
    let changed = true;
    while (changed) {
      changed = false;
      for (const pEl of root.querySelectorAll('p, h1, h2, h3')) {
        if (![...pEl.children].some(c => BLOCK.test(c.nodeName))) continue;
        const frag = document.createDocumentFragment(); let run = null; const runs = [];
        for (const ch of [...pEl.childNodes]) {
          if (ch.nodeType === 1 && BLOCK.test(ch.nodeName)) { run = null; frag.appendChild(ch); }
          else { if (!run) { run = document.createElement(pEl.nodeName); if (pEl.getAttribute('style')) run.setAttribute('style', pEl.getAttribute('style')); frag.appendChild(run); runs.push(run); } run.appendChild(ch); }
        }
        for (const r of runs) if (!r.textContent.trim() && !r.querySelector('img,br')) r.remove();
        pEl.replaceWith(frag); changed = true; break;
      }
    }
    return root;
  }
  const app = {
    libs: [],
    blank: () => ({html: '<h1>Untitled document</h1><p><br></p>'}),
    async importFile(buf){
      await MO.loadScript('vendor/mammoth.min.js');
      const r = await mammoth.convertToHtml({arrayBuffer: buf}, {styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Subtitle'] => h2:fresh"]});
      let html = sanitize(r.value || '<p><br></p>');
      // Give imported tables visible borders like Word's default grid.
      return {html: html || '<p><br></p>'};
    },
    async mount(host, file, ctx){
      host.innerHTML = `
        <div class="toolbar" role="toolbar" aria-label="Formatting">
          <button data-cmd="undo" title="Undo">↶</button><button data-cmd="redo" title="Redo">↷</button><span class="sep"></span>
          <select data-block title="Style"><option value="P">Normal</option><option value="H1">Heading 1</option><option value="H2">Heading 2</option><option value="H3">Heading 3</option><option value="BLOCKQUOTE">Quote</option></select>
          <select data-font title="Font">${FONTS.map(f => `<option>${f}</option>`).join('')}</select>
          <select data-size title="Size">${SIZES.map(s => `<option ${s===11?'selected':''}>${s}</option>`).join('')}</select><span class="sep"></span>
          <button data-cmd="bold" title="Bold (Ctrl+B)"><b>B</b></button><button data-cmd="italic" title="Italic (Ctrl+I)"><i>I</i></button><button data-cmd="underline" title="Underline (Ctrl+U)"><u>U</u></button><button data-cmd="strikeThrough" title="Strikethrough"><s>S</s></button>
          <label class="tbb" title="Text colour">A<input type="color" data-color="foreColor" value="#c00000"></label>
          <label class="tbb" title="Highlight">▌<input type="color" data-color="hiliteColor" value="#ffff00"></label><span class="sep"></span>
          <button data-cmd="justifyLeft" title="Align left"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button><button data-cmd="justifyCenter" title="Centre"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="4.5" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="4.5" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button><button data-cmd="justifyRight" title="Align right"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="7" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="7" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button><button data-cmd="justifyFull" title="Justify"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="7" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="15" width="14" height="2" rx="1" fill="currentColor"/></svg></button><span class="sep"></span>
          <button data-cmd="insertUnorderedList" title="Bullets">• ≡</button><button data-cmd="insertOrderedList" title="Numbering">1.≡</button><button data-cmd="outdent" title="Decrease indent">⇤</button><button data-cmd="indent" title="Increase indent">⇥</button><span class="sep"></span>
          <button data-act="table" title="Insert table">▦ Table</button><button data-act="row" title="Add row below">+Row</button><button data-act="col" title="Add column right">+Col</button><button data-act="delrow" title="Delete row">−Row</button>
          <label class="tbb" title="Insert picture">🖼 Picture<input type="file" accept="image/*" data-img hidden></label>
          <button data-act="link" title="Link">🔗</button><button data-act="hr" title="Page break">⤓ Break</button><button data-cmd="removeFormat" title="Clear formatting">⌫ Clear</button>
        </div>
        <div class="doc-scroll"><div class="doc-page" contenteditable="true" spellcheck="true"></div></div>
        <div class="doc-stats"></div>`;
      const ed = host.querySelector('.doc-page'), stats = host.querySelector('.doc-stats'), tb = host.querySelector('.toolbar');
      ed.innerHTML = (file.content && file.content.html) || '<p><br></p>';
      normalize(ed);
      try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch(e){}
      let savedRange = null;
      const saveSel = () => { const s = getSelection(); if (s.rangeCount && ed.contains(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange(); };
      const restoreSel = () => { ed.focus(); if (savedRange) { const s = getSelection(); s.removeAllRanges(); s.addRange(savedRange); } };
      const exec = (c, v) => { restoreSel(); document.execCommand('styleWithCSS', false, c === 'hiliteColor' || c === 'foreColor'); document.execCommand(c, false, v); changed(); refresh(); };
      const updateStats = () => { const t = ed.innerText.trim(); stats.textContent = `${t ? t.split(/\s+/).length : 0} words · ${t.length} characters`; };
      const changed = () => { ctx.changed(); updateStats(); };
      function topBlock(){ const sl = getSelection(); let n = sl.rangeCount ? sl.anchorNode : null; if (!n || !ed.contains(n)) return null; while (n && n.parentNode !== ed) n = n.parentNode; return n; }
      function insertBlock(node){
        restoreSel(); const blk = topBlock(); const after = document.createElement('p'); after.innerHTML = '<br>';
        if (blk) { blk.after(node); } else ed.appendChild(node);
        node.after(after);
        const target = node.querySelector ? (node.querySelector('td') || after) : after; const rg = document.createRange(); rg.selectNodeContents(target); rg.collapse(true); const sl = getSelection(); sl.removeAllRanges(); sl.addRange(rg);
        changed();
      }
      function cellAtCaret(){ const s = getSelection(); let n = s.anchorNode; while (n && n !== ed) { if (n.nodeName === 'TD' || n.nodeName === 'TH') return n; n = n.parentNode; } return null; }
      function refresh(){
        for (const b of tb.querySelectorAll('[data-cmd]')) { try { b.classList.toggle('active', document.queryCommandState(b.dataset.cmd)); } catch(e){} }
        try { const blk = (document.queryCommandValue('formatBlock') || 'p').toUpperCase(); const sel = tb.querySelector('[data-block]'); if ([...sel.options].some(o => o.value === blk)) sel.value = blk; else sel.value = 'P'; } catch(e){}
      }
      tb.addEventListener('mousedown', e => { if (e.target.closest('button')) e.preventDefault(); });
      tb.addEventListener('click', e => {
        const b = e.target.closest('button'); if (!b) return;
        if (b.dataset.cmd) return exec(b.dataset.cmd);
        const a = b.dataset.act;
        if (a === 'table') {
          const spec = prompt('Table size (columns x rows)', '3x3'); if (!spec) return;
          const [c, r] = spec.toLowerCase().split('x').map(n => Math.max(1, Math.min(20, parseInt(n, 10) || 3)));
          const t = document.createElement('table'); const tb2 = document.createElement('tbody'); t.appendChild(tb2);
          for (let i = 0; i < r; i++) { const tr = document.createElement('tr'); tr.innerHTML = '<td><br></td>'.repeat(c); tb2.appendChild(tr); }
          insertBlock(t);
        } else if (a === 'row' || a === 'col' || a === 'delrow') {
          restoreSel(); const td = cellAtCaret(); if (!td) return ctx.toast('Put the cursor in a table cell first.');
          const tr = td.parentNode;
          if (a === 'row') { const nr = tr.cloneNode(true); nr.querySelectorAll('td,th').forEach(x => x.innerHTML = '<br>'); tr.after(nr); }
          if (a === 'delrow') { if (tr.parentNode.rows.length > 1) tr.remove(); else tr.closest('table').remove(); }
          if (a === 'col') { const i = td.cellIndex; for (const r of td.closest('table').rows) { const c = r.cells[Math.min(i, r.cells.length-1)]; const n = document.createElement(c.nodeName); n.innerHTML = '<br>'; c.after(n); } }
          changed();
        } else if (a === 'link') {
          saveSel(); const u = prompt('Link address', 'https://'); if (u && /^(https?:|mailto:)/i.test(u)) exec('createLink', u);
        } else if (a === 'hr') {
          const hr = document.createElement('hr'); hr.className = 'page-break'; insertBlock(hr);
        }
      });
      tb.querySelector('[data-block]').onchange = e => exec('formatBlock', e.target.value);
      tb.querySelector('[data-font]').onchange = e => exec('fontName', e.target.value);
      tb.querySelector('[data-size]').onchange = e => {
        restoreSel(); document.execCommand('styleWithCSS', false, false); document.execCommand('fontSize', false, '7');
        ed.querySelectorAll('font[size="7"]').forEach(f => { const s = document.createElement('span'); s.style.fontSize = e.target.value + 'pt'; s.innerHTML = f.innerHTML; f.replaceWith(s); });
        changed();
      };
      tb.querySelectorAll('[data-color]').forEach(inp => { inp.addEventListener('pointerdown', saveSel); inp.onchange = e => exec(inp.dataset.color, e.target.value); });
      tb.querySelector('[data-img]').onchange = e => {
        const f = e.target.files[0]; e.target.value = ''; if (!f) return;
        const img = new Image(); const url = URL.createObjectURL(f);
        img.onload = () => { // downscale big photos to save memory and file size
          const max = 1400, sc = Math.min(1, max / Math.max(img.width, img.height));
          const c = document.createElement('canvas'); c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc);
          c.getContext('2d').drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
          const data = c.toDataURL(f.type === 'image/png' ? 'image/png' : 'image/jpeg', 0.85);
          exec('insertHTML', `<img src="${data}" width="${Math.min(c.width, 600)}">`);
        };
        img.src = url;
      };
      ed.addEventListener('input', changed);
      ed.addEventListener('keyup', () => { saveSel(); refresh(); });
      ed.addEventListener('mouseup', () => { saveSel(); refresh(); });
      ed.addEventListener('touchend', () => setTimeout(() => { saveSel(); refresh(); }, 50));
      document.addEventListener('selectionchange', saveSel);
      ed.addEventListener('paste', e => {
        const html = e.clipboardData && e.clipboardData.getData('text/html');
        if (html) { e.preventDefault(); document.execCommand('insertHTML', false, sanitize(html.replace(/<!--[\s\S]*?-->/g, ''))); changed(); }
      });
      updateStats();
      return {
        getContent: () => ({html: ed.innerHTML}),
        normalize: () => normalize(ed),
        async exportBlob(){
          await MO.loadScript('vendor/docx-kit.js');
          const K = window.DocxKit;
          const clean = normalize(ed.cloneNode(true));
          const children = K.htmlToContent(clean.innerHTML, {stylesheet: DOC_CSS, maxImageWidth: 600});
          const doc = new K.Document({creator: 'My Office', styles: {default: {document: {run: {font: 'Arial', size: 22}}}}, sections: [{children: children.length ? children : [new K.Paragraph('')]}]});
          return K.Packer.toBlob(doc);
        },
        printView: () => ({html: `<div class="print-doc">${normalize(ed.cloneNode(true)).innerHTML}</div>`, page: 'size:A4;margin:20mm'}),
        destroy(){ document.removeEventListener('selectionchange', saveSel); },
      };
    }
  };
  MOApps.doc = app;
})();
