/* Spreadsheets: own light grid + fast-formula-parser/formulajs engine. .xlsx via ExcelJS, CSV import/export. */
(function(){
  'use strict';
  const DEF_ROWS = 100, DEF_COLS = 26, MAX_RENDER_ROWS = 1000, DEF_W = 96;
  const colName = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const colNum = s => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const addr = (r, c) => colName(c) + r;
  const parseAddr = a => { const m = /^([A-Z]+)(\d+)$/.exec(a); return m ? {c: colNum(m[1]), r: +m[2]} : null; };
  const isErr = v => v && typeof v === 'object' && (v._error || v.error);
  const errText = v => (v && (v._error || v.error)) || '#ERROR!';

  function blankSheet(name){ return {name, cells:{}, colW:{}, rows:DEF_ROWS, cols:DEF_COLS}; }

  /* ---------- reference shifting (fill, copy/paste, insert/delete rows and columns) ---------- */
  const REF_RE = /((?:'[^']+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?)([A-Z]{1,3})(\$?)([0-9]{1,7})(?![A-Za-z0-9_(])/g;
  function mapRefs(formula, fn){
    // only outside "string literals"
    return formula.split(/("(?:[^"]|"")*")/).map((seg, i) => i % 2 ? seg : seg.replace(REF_RE, (m, sh, ac, col, ar, row, off, str) => {
      const prev = off > 0 ? str[off - 1] : '';
      if (/[A-Za-z0-9_.]/.test(prev) && !sh) return m;
      const sheet = sh ? sh.slice(0, -1).replace(/^'|'$/g, '') : null;
      const out = fn({sheet, c: colNum(col), r: +row, absC: !!ac, absR: !!ar});
      if (!out) return (sh || '') + '#REF!';
      return (sh || '') + (ac ? '$' : '') + colName(out.c) + (ar ? '$' : '') + out.r;
    })).join('');
  }
  const shiftRel = (f, dr, dc) => mapRefs(f, ref => { const r = ref.absR ? ref.r : ref.r + dr, c = ref.absC ? ref.c : ref.c + dc; return r < 1 || c < 1 ? null : {r, c}; });

  /* ---------- formula engine ---------- */
  let FP_FUNCS = null;
  function buildFunctions(){
    if (FP_FUNCS) return FP_FUNCS;
    const FJ = window.formulajs, F = {};
    const probe = new FormulaParser({onCell: () => null, onRange: () => [[null]]});
    const have = new Set(Object.keys(probe.functions || {}));
    const wrap = k => (...args) => { const r = FJ[k](...args.map(a => a.value)); if (r instanceof Error) return new FormulaParser.FormulaError(/^#/.test(r.message) ? r.message : '#VALUE!'); return r; };
    for (const k of Object.keys(FJ)) {
      if (typeof FJ[k] !== 'function' || !/^[A-Z][A-Z0-9.]*$/.test(k)) continue;
      let need = !have.has(k);
      if (!need) { try { probe.parse(k + '(1)', {sheet:'S', row:1, col:1}); } catch (e) { if (/not implemented/i.test(String(e.details && (e.details.message || e.details)))) need = true; } }
      if (need) F[k] = wrap(k);
    }
    // Excel's dotted names (STDEV.S, RANK.EQ, PERCENTILE.INC...) are undotted in formula.js
    for (const k of have) { if (k.includes('.') && !F[k] && FJ[k.replace(/\./g, '')]) { try { probe.parse(k + '(1)', {sheet:'S', row:1, col:1}); } catch (e) { if (/not implemented/i.test(String(e.details && (e.details.message || e.details)))) F[k] = wrap(k.replace(/\./g, '')); } } }
    F.COUNTIF = wrap('COUNTIF');
    const flat = v => Array.isArray(v) ? v.flat(Infinity) : [v];
    F.XLOOKUP = (look, arr, ret, nf) => {
      const a = flat(arr.value), r = flat(ret.value); const L = look.value;
      const i = a.findIndex(x => x === L || (typeof x === 'string' && typeof L === 'string' && x.toLowerCase() === L.toLowerCase()) || (x != null && L != null && Number(x) === Number(L) && x !== '' && L !== ''));
      if (i >= 0) return r[i] == null ? 0 : r[i];
      return nf ? nf.value : new FormulaParser.FormulaError('#N/A');
    };
    return (FP_FUNCS = F);
  }
  function Engine(model){
    const cache = new Map(), busy = new Set();
    const sheetIdx = name => { if (name == null) return -1; const n = String(name).toLowerCase(); return model.sheets.findIndex(s => s.name.toLowerCase() === n); };
    function literal(v){
      if (v == null || v === '') return null;
      if (typeof v !== 'string') return v;
      const t = v.trim();
      if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return Number(t);
      if (/^[-+]?\d+(\.\d+)?%$/.test(t)) return Number(t.slice(0, -1)) / 100;
      if (/^true$/i.test(t)) return true; if (/^false$/i.test(t)) return false;
      return v;
    }
    function value(si, r, c){
      const sh = model.sheets[si]; if (!sh) return new FormulaParser.FormulaError('#REF!');
      const key = si + ':' + r + ':' + c;
      if (cache.has(key)) return cache.get(key);
      const cell = sh.cells[addr(r, c)];
      let out;
      if (!cell || cell.v == null || cell.v === '') out = null;
      else if (typeof cell.v === 'string' && cell.v[0] === '=' && cell.v.length > 1) {
        if (busy.has(key)) return new FormulaParser.FormulaError('#REF!', 'Circular reference');
        busy.add(key);
        try {
          depth++;
          // fast-formula-parser is not re-entrant, so nested evaluations each get their own parser
          const P = parsers[depth] || (parsers[depth] = makeParser());
          try { out = P.parse(cell.v.slice(1), {sheet: sh.name, row: r, col: c}); } finally { depth--; }
          if (Array.isArray(out)) out = Array.isArray(out[0]) ? out[0][0] : out[0];
          if (out === undefined) out = null;
        } catch (e) { out = e && (e._error || e.error) ? e : new FormulaParser.FormulaError('#ERROR!'); }
        busy.delete(key);
      } else out = literal(cell.v);
      cache.set(key, out);
      return out;
    }
    const resolve = name => { const i = name == null ? cur : sheetIdx(name); return i; };
    let cur = 0, depth = -1; const parsers = [];
    const makeParser = () => new FormulaParser({
      functions: buildFunctions(),
      onCell: ref => { const si = resolve(ref.sheet); if (si < 0) return new FormulaParser.FormulaError('#REF!'); return value(si, ref.row, ref.col); },
      onRange: ref => {
        const si = resolve(ref.sheet); if (si < 0) return [[new FormulaParser.FormulaError('#REF!')]];
        const sh = model.sheets[si]; const used = usedRange(sh);
        const r2 = Math.min(ref.to.row, Math.max(used.r, ref.from.row)), c2 = Math.min(ref.to.col, Math.max(used.c, ref.from.col));
        const out = [];
        for (let r = ref.from.row; r <= r2; r++) { const row = []; for (let c = ref.from.col; c <= c2; c++) row.push(value(si, r, c)); out.push(row); }
        return out;
      },
    });
    return {
      value: (si, r, c) => { cur = si; return value(si, r, c); },
      // evaluate formulas top-to-bottom so long dependency chains stay shallow
      warm(){ model.sheets.forEach((sh, i) => { Object.keys(sh.cells).map(parseAddr).filter(Boolean).sort((a, b) => a.r - b.r || a.c - b.c).forEach(p => { cur = i; value(i, p.r, p.c); }); }); },
      clear: () => cache.clear(),
    };
  }
  function usedRange(sh){
    let r = 0, c = 0;
    for (const a in sh.cells) { const cell = sh.cells[a]; if (cell.v == null || cell.v === '') continue; const p = parseAddr(a); if (!p) continue; if (p.r > r) r = p.r; if (p.c > c) c = p.c; }
    return {r, c};
  }

  /* ---------- number formats ---------- */
  function serialToDate(n){ return new Date(Math.round((n - 25569) * 86400000)); }
  function dateToSerial(d){ return d.getTime() / 86400000 + 25569; }
  function fmt(v, nf){
    if (v == null) return '';
    if (isErr(v)) return errText(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v !== 'number') return String(v);
    if (!nf || nf === 'General') { if (Math.abs(v) >= 1e11 || (Math.abs(v) < 1e-9 && v !== 0)) return v.toExponential(4); return String(Math.round(v * 1e10) / 1e10); }
    if (/[dmy]/i.test(nf) && !/[0#]/.test(nf)) { const d = serialToDate(v); const p = n => String(n).padStart(2, '0');
      if (/h/i.test(nf)) return `${d.getUTCDate()}/${p(d.getUTCMonth()+1)}/${d.getUTCFullYear()} ${p(d.getUTCHours())}:${p(d.getUTCMinutes())}`;
      return `${d.getUTCDate()}/${p(d.getUTCMonth()+1)}/${d.getUTCFullYear()}`; }
    const pct = /%/.test(nf); let x = pct ? v * 100 : v;
    const dm = /0\.(0+)/.exec(nf); const dp = dm ? dm[1].length : 0;
    const grp = /#,##/.test(nf);
    let s = Math.abs(x).toLocaleString('en-AU', {minimumFractionDigits: dp, maximumFractionDigits: dp, useGrouping: grp});
    const cur = /\$/.test(nf) ? '$' : '';
    return (x < 0 ? '-' : '') + cur + s + (pct ? '%' : '');
  }
  const NUMFMTS = [['General','General'],['0','Number (0)'],['0.00','Number (0.00)'],['#,##0.00','Thousands (1,234.00)'],['$#,##0.00','Currency ($)'],['0%','Percent (0%)'],['0.00%','Percent (0.00%)'],['d/mm/yyyy','Date']];

  /* ---------- CSV ---------- */
  function parseCSV(text){
    const rows = []; let row = [], f = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i+1] === '"') { f += '"'; i++; } else q = false; } else f += ch; }
      else if (ch === '"') q = true;
      else if (ch === ',') { row.push(f); f = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && text[i+1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
      else f += ch;
    }
    if (f !== '' || row.length) { row.push(f); rows.push(row); }
    return rows;
  }
  const csvCell = s => /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;

  /* ---------- xlsx import/export (ExcelJS) ---------- */
  // Newer Excel functions must be stored with the _xlfn. prefix or Excel shows #NAME?
  const XLFN_RE = /(^|[^A-Za-z0-9_.])(XLOOKUP|XMATCH|TEXTJOIN|CONCAT|IFS|SWITCH|MAXIFS|MINIFS|UNIQUE|SORT|SORTBY|FILTER|SEQUENCE|IFNA|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P|DAYS|ISOWEEKNUM)\(/g;
  const argb = c => c && c.argb && /^[0-9A-F]{8}$/i.test(c.argb) ? '#' + c.argb.slice(2).toLowerCase() : null;
  async function importXlsx(buf){
    await MO.loadScript('vendor/exceljs.min.js');
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
    const sheets = [];
    wb.eachSheet(ws => {
      const sh = blankSheet(ws.name || ('Sheet' + (sheets.length + 1)));
      let maxR = 0, maxC = 0;
      ws.eachRow({includeEmpty:false}, (row, rn) => row.eachCell({includeEmpty:false}, (cell, cn) => {
        let v = cell.value, raw = null, nf = cell.numFmt || null;
        if (cell.formula) raw = '=' + String(cell.formula).replace(/_xlfn\.|_xlws\./g, '');
        else if (v instanceof Date) { raw = String(dateToSerial(v)); if (!nf || nf === 'General') nf = 'd/mm/yyyy'; }
        else if (v && typeof v === 'object') {
          if (v.richText) raw = v.richText.map(t => t.text).join('');
          else if (v.text != null) raw = String(v.text);
          else if (v.error) raw = v.error;
          else if (v.result != null) raw = String(v.result);
        } else if (v != null) raw = typeof v === 'string' ? v : String(v);
        if (raw == null || raw === '') return;
        const s = {};
        const font = cell.font || {}; if (font.bold) s.b = 1; if (font.italic) s.i = 1; if (font.underline) s.u = 1; const fc = argb(font.color); if (fc && fc !== '#000000') s.fc = fc;
        const al = cell.alignment && cell.alignment.horizontal; if (al === 'center' || al === 'right' || al === 'left') s.al = al;
        const fill = cell.fill; if (fill && fill.type === 'pattern' && fill.pattern === 'solid') { const bg = argb(fill.fgColor); if (bg && bg !== '#ffffff') s.bg = bg; }
        if (nf && nf !== 'General') s.nf = nf;
        sh.cells[addr(rn, cn)] = Object.keys(s).length ? {v: raw, s} : {v: raw};
        if (rn > maxR) maxR = rn; if (cn > maxC) maxC = cn;
      }));
      for (let c = 1; c <= Math.max(maxC, 1); c++) { const w = ws.getColumn(c).width; if (w) sh.colW[colName(c)] = Math.round(Math.min(600, w * 7 + 8)); }
      sh.rows = Math.max(DEF_ROWS, maxR + 20); sh.cols = Math.max(DEF_COLS, maxC + 3);
      sheets.push(sh);
    });
    if (!sheets.length) sheets.push(blankSheet('Sheet1'));
    return {sheets, active: 0};
  }
  async function exportXlsx(model){
    await MO.loadScript('vendor/exceljs.min.js');
    const eng = Engine(model); eng.warm();
    const wb = new ExcelJS.Workbook(); wb.creator = 'My Office';
    model.sheets.forEach((sh, si) => {
      const ws = wb.addWorksheet(sh.name.slice(0, 31) || ('Sheet' + (si + 1)));
      for (const a in sh.cells) {
        const cell = sh.cells[a]; if (cell.v == null || cell.v === '') continue;
        const p = parseAddr(a); if (!p) continue;
        const xc = ws.getCell(a);
        if (typeof cell.v === 'string' && cell.v[0] === '=') {
          const res = eng.value(si, p.r, p.c);
          const fx = cell.v.slice(1).replace(XLFN_RE, (m, pre, name) => pre + '_xlfn.' + name + '(');
          xc.value = isErr(res) || res == null ? {formula: fx} : {formula: fx, result: res};
        } else {
          const lit = (() => { const t = String(cell.v).trim(); if (/^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return Number(t); if (/^[-+]?\d+(\.\d+)?%$/.test(t)) return Number(t.slice(0,-1))/100; if (/^true$/i.test(t)) return true; if (/^false$/i.test(t)) return false; return cell.v; })();
          xc.value = lit;
        }
        const s = cell.s || {};
        if (s.b || s.i || s.u || s.fc) xc.font = {bold: !!s.b, italic: !!s.i, underline: !!s.u, color: s.fc ? {argb: 'FF' + s.fc.slice(1).toUpperCase()} : undefined};
        if (s.al) xc.alignment = {horizontal: s.al};
        if (s.bg) xc.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF' + s.bg.slice(1).toUpperCase()}};
        if (s.nf) xc.numFmt = s.nf;
      }
      for (const k in sh.colW) ws.getColumn(colNum(k)).width = Math.max(4, (sh.colW[k] - 8) / 7);
    });
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], {type: MO.KIND.sheet.mime});
  }

  /* ---------- the editor ---------- */
  const app = {
    libs: ['vendor/formula-engine.js'],
    blank: () => ({sheets: [blankSheet('Sheet1')], active: 0}),
    async importFile(buf, name){
      if (/\.csv$/i.test(name)) {
        const rows = parseCSV(new TextDecoder().decode(buf)); const sh = blankSheet('Sheet1');
        rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') sh.cells[addr(r + 1, c + 1)] = {v}; }));
        sh.rows = Math.max(DEF_ROWS, rows.length + 20); sh.cols = Math.max(DEF_COLS, Math.max(0, ...rows.map(r => r.length)) + 3);
        return {sheets: [sh], active: 0};
      }
      return importXlsx(buf);
    },
    async mount(host, file, ctx){
      const model = JSON.parse(JSON.stringify(file.content || app.blank()));
      if (!model.sheets || !model.sheets.length) model.sheets = [blankSheet('Sheet1')];
      let si = Math.min(model.active || 0, model.sheets.length - 1);
      let eng = Engine(model);
      let cur = {r: 1, c: 1}, anchor = {r: 1, c: 1}, extendMode = false, editing = false;
      const undo = [], redo = [];
      let tds = []; // tds[r][c]
      let renderRows = 0;
      host.innerHTML = `
        <div class="toolbar" role="toolbar" aria-label="Spreadsheet tools">
          <button data-a="undo" title="Undo">↶</button><button data-a="redo" title="Redo">↷</button><span class="sep"></span>
          <button data-a="b" title="Bold"><b>B</b></button><button data-a="i" title="Italic"><i>I</i></button><button data-a="u" title="Underline"><u>U</u></button>
          <button data-a="left" title="Align left"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="2" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button><button data-a="center" title="Centre"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="4.5" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="4.5" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button><button data-a="right" title="Align right"><svg width="18" height="18" viewBox="0 0 18 18" aria-hidden="true"><rect x="2" y="3" width="14" height="2" rx="1" fill="currentColor"/><rect x="7" y="7" width="9" height="2" rx="1" fill="currentColor"/><rect x="2" y="11" width="14" height="2" rx="1" fill="currentColor"/><rect x="7" y="15" width="9" height="2" rx="1" fill="currentColor"/></svg></button>
          <label class="tbb" title="Fill colour">▉<input type="color" data-color="bg" value="#fff2cc"></label>
          <label class="tbb" title="Text colour">A<input type="color" data-color="fc" value="#c00000"></label>
          <select data-nf title="Number format">${NUMFMTS.map(n => `<option value="${n[0]}">${n[1]}</option>`).join('')}</select><span class="sep"></span>
          <button data-a="sum" title="AutoSum">Σ Sum</button>
          <select data-fn title="Insert function"><option value="">ƒx Function</option>${['SUM','AVERAGE','COUNT','MIN','MAX','IF','VLOOKUP','XLOOKUP','COUNTIF','SUMIF','ROUND','CONCAT','TODAY'].map(f => `<option>${f}</option>`).join('')}</select>
          <button data-a="filldown" title="Fill down">⇩ Fill</button><button data-a="extend" title="Select a range by tapping the other corner">⇲ Range</button><span class="sep"></span>
          <button data-a="insrow" title="Insert row above">+Row</button><button data-a="delrow" title="Delete row">−Row</button><button data-a="inscol" title="Insert column left">+Col</button><button data-a="delcol" title="Delete column">−Col</button>
          <button data-a="wider" title="Wider column">↔+</button><button data-a="narrower" title="Narrower column">↔−</button>
          <button data-a="sortasc" title="Sort selection A to Z by its first column">A→Z</button><button data-a="sortdesc" title="Sort Z to A">Z→A</button><button data-a="clear" title="Clear cells">⌫</button>
        </div>
        <div class="fx-bar"><span class="addr" aria-live="polite">A1</span><span class="fx">fx</span><input type="text" class="fx-input" aria-label="Cell contents" autocomplete="off" autocapitalize="off" spellcheck="false" enterkeyhint="done"></div>
        <div class="grid-wrap" tabindex="0"></div>
        <div class="sheet-tabs"></div>`;
      const wrap = host.querySelector('.grid-wrap'), fx = host.querySelector('.fx-input'), addrEl = host.querySelector('.addr'), tb = host.querySelector('.toolbar'), tabs = host.querySelector('.sheet-tabs');
      const S = () => model.sheets[si];
      const snapshot = () => { undo.push(JSON.stringify(model)); if (undo.length > 50) undo.shift(); redo.length = 0; };
      const changed = () => { eng.clear(); refreshValues(); ctx.changed(); };
      const selRect = () => ({r1: Math.min(cur.r, anchor.r), r2: Math.max(cur.r, anchor.r), c1: Math.min(cur.c, anchor.c), c2: Math.max(cur.c, anchor.c)});
      function forSel(fn){ const s = selRect(); for (let r = s.r1; r <= s.r2; r++) for (let c = s.c1; c <= s.c2; c++) fn(r, c); }

      function render(){
        const sh = S();
        renderRows = Math.min(sh.rows, MAX_RENDER_ROWS);
        let tw = 44; for (let c = 1; c <= sh.cols; c++) tw += sh.colW[colName(c)] || DEF_W;
        let h = `<table class="grid" style="width:${tw}px"><colgroup><col style="width:44px">`;
        for (let c = 1; c <= sh.cols; c++) h += `<col style="width:${sh.colW[colName(c)] || DEF_W}px">`;
        h += '</colgroup><thead><tr><th class="rh corner"></th>';
        for (let c = 1; c <= sh.cols; c++) h += `<th data-c="${c}">${colName(c)}</th>`;
        h += '</tr></thead><tbody>';
        for (let r = 1; r <= renderRows; r++) { h += `<tr><th class="rh" data-r="${r}">${r}</th>`; for (let c = 1; c <= sh.cols; c++) h += `<td data-r="${r}" data-c="${c}"></td>`; h += '</tr>'; }
        h += '</tbody></table>';
        if (sh.rows > renderRows) h += `<button class="btn more-rows">Show more rows (${sh.rows - renderRows} hidden)</button>`;
        wrap.innerHTML = h;
        tds = [];
        for (const td of wrap.querySelectorAll('td')) { const r = +td.dataset.r, c = +td.dataset.c; (tds[r] || (tds[r] = []))[c] = td; }
        refreshValues(); renderTabs(); select(cur.r, cur.c, false, true);
      }
      function refreshValues(){
        const sh = S();
        for (let r = 1; r <= renderRows; r++) { const row = tds[r]; if (!row) continue; for (let c = 1; c <= sh.cols; c++) paint(row[c], r, c); }
      }
      function paint(td, r, c){
        const cell = S().cells[addr(r, c)];
        if (!cell) { if (td.textContent || td.className || td.getAttribute('style')) { td.textContent = ''; td.removeAttribute('style'); td.className = td.className.replace(/\b(num|err)\b/g, '').trim(); } return; }
        const v = eng.value(si, r, c); const s = cell.s || {};
        td.textContent = fmt(v, s.nf);
        let st = '';
        if (s.b) st += 'font-weight:700;'; if (s.i) st += 'font-style:italic;'; if (s.u) st += 'text-decoration:underline;';
        if (s.bg) st += `background:${s.bg};`; if (s.fc) st += `color:${s.fc};`; if (s.al) st += `text-align:${s.al};`;
        td.setAttribute('style', st);
        td.classList.toggle('num', typeof v === 'number' && !s.al); td.classList.toggle('err', !!isErr(v));
      }
      function renderTabs(){
        tabs.innerHTML = model.sheets.map((s, i) => `<button data-s="${i}" class="${i === si ? 'active' : ''}">${ctx.esc(s.name)}</button>`).join('') + '<button data-s="add" title="Add sheet">＋</button><button data-s="ren" title="Rename sheet">✎</button><button data-s="del" title="Delete sheet">🗑</button>';
      }
      function select(r, c, extend, noScroll){
        const sh = S(); r = Math.max(1, Math.min(r, renderRows)); c = Math.max(1, Math.min(c, sh.cols));
        wrap.querySelectorAll('td.cur,td.insel,th.hsel').forEach(e => e.classList.remove('cur', 'insel', 'hsel'));
        cur = {r, c}; if (!extend) anchor = {r, c};
        const s = selRect();
        if (s.r2 - s.r1 + s.c2 - s.c1 > 0) forSel((rr, cc) => tds[rr] && tds[rr][cc] && tds[rr][cc].classList.add('insel'));
        const td = tds[r] && tds[r][c]; if (td) { td.classList.add('cur'); if (!noScroll) td.scrollIntoView({block: 'nearest', inline: 'nearest'}); }
        const rh = wrap.querySelector(`th.rh[data-r="${r}"]`), ch = wrap.querySelector(`thead th[data-c="${c}"]`); rh && rh.classList.add('hsel'); ch && ch.classList.add('hsel');
        addrEl.textContent = s.r1 === s.r2 && s.c1 === s.c2 ? addr(r, c) : `${addr(s.r1, s.c1)}:${addr(s.r2, s.c2)}`;
        const cell = sh.cells[addr(r, c)]; if (!editing) fx.value = cell ? cell.v : '';
        const st = (cell && cell.s) || {}; tb.querySelector('[data-nf]').value = st.nf || 'General';
        ['b', 'i', 'u'].forEach(k => tb.querySelector(`[data-a="${k}"]`).classList.toggle('active', !!st[k]));
      }
      function setCell(r, c, v){
        const sh = S(), a = addr(r, c);
        if (v === '' || v == null) { if (sh.cells[a]) { if (sh.cells[a].s) sh.cells[a].v = ''; else delete sh.cells[a]; } }
        else { sh.cells[a] = Object.assign(sh.cells[a] || {}, {v}); }
        if (r > sh.rows - 5) sh.rows = r + 20;
      }
      function commit(move){
        if (!editing) return;
        editing = false;
        const cell = S().cells[addr(cur.r, cur.c)];
        const nv = fx.value;
        if ((cell ? cell.v : '') !== nv) { snapshot(); setCell(cur.r, cur.c, nv.startsWith('=') ? nv.replace(/^=\s*/, '=') : nv); changed(); }
        if (move === 'down') select(cur.r + 1, cur.c); else if (move === 'right') select(cur.r, cur.c + 1); else select(cur.r, cur.c);
      }
      function startEdit(initial){ editing = true; if (initial != null) fx.value = initial; fx.focus(); if (initial != null) fx.setSelectionRange(fx.value.length, fx.value.length); }
      function style(fn){ snapshot(); forSel((r, c) => { const sh = S(), a = addr(r, c); const cell = sh.cells[a] || (sh.cells[a] = {v: ''}); cell.s = cell.s || {}; fn(cell.s); if (!Object.keys(cell.s).length) delete cell.s; }); changed(); select(cur.r, cur.c, true, true); }

      function shiftAll(kind, at, delta){
        // kind 'r' or 'c'; insert (delta=1) or delete (delta=-1) at index `at` on sheet si
        snapshot();
        const target = S(), tname = target.name.toLowerCase();
        const moved = {};
        for (const a in target.cells) {
          const p = parseAddr(a); const k = kind === 'r' ? p.r : p.c;
          if (delta < 0 && k === at) continue;
          const nk = k >= at ? k + delta : k;
          moved[kind === 'r' ? addr(nk, p.c) : addr(p.r, nk)] = target.cells[a];
        }
        target.cells = moved;
        model.sheets.forEach((sh, i) => { for (const a in sh.cells) { const cell = sh.cells[a]; if (typeof cell.v !== 'string' || cell.v[0] !== '=') continue;
          cell.v = '=' + mapRefs(cell.v.slice(1), ref => {
            const onTarget = ref.sheet ? ref.sheet.toLowerCase() === tname : i === si; if (!onTarget) return {r: ref.r, c: ref.c};
            const k = kind === 'r' ? ref.r : ref.c;
            if (delta < 0 && k === at) return null;
            const nk = k >= at ? k + delta : k;
            return kind === 'r' ? {r: nk, c: ref.c} : {r: ref.r, c: nk};
          }); } });
        if (kind === 'r') target.rows = Math.max(DEF_ROWS, target.rows + delta); else target.cols = Math.max(DEF_COLS, target.cols + delta);
        eng.clear(); render(); ctx.changed();
      }
      function sortSel(desc){
        let s = selRect();
        if (s.r1 === s.r2) { const u = usedRange(S()); s = {r1: 1, r2: u.r, c1: 1, c2: Math.max(u.c, 1)}; if (!confirm('Sort all data (rows 1 to ' + u.r + ') by column ' + colName(cur.c) + '? Tip: select a range first to sort only that.')) return; }
        snapshot();
        const sh = S(), rows = [];
        for (let r = s.r1; r <= s.r2; r++) { const row = []; for (let c = s.c1; c <= s.c2; c++) row.push(sh.cells[addr(r, c)] || null); rows.push({row, key: eng.value(si, r, s.r1 === s.r2 ? cur.c : Math.max(s.c1, Math.min(cur.c, s.c2)))}); }
        rows.sort((a, b) => { const x = a.key, y = b.key; if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1;
          const cmp = typeof x === 'number' && typeof y === 'number' ? x - y : String(x).localeCompare(String(y), 'en-AU', {numeric: true, sensitivity: 'base'}); return desc ? -cmp : cmp; });
        rows.forEach((o, i) => o.row.forEach((cell, j) => { const a = addr(s.r1 + i, s.c1 + j); if (cell) sh.cells[a] = cell; else delete sh.cells[a]; }));
        changed();
      }
      function copySel(cut){
        const s = selRect(), sh = S(); const block = [];
        for (let r = s.r1; r <= s.r2; r++) { const row = []; for (let c = s.c1; c <= s.c2; c++) row.push(sh.cells[addr(r, c)] ? JSON.parse(JSON.stringify(sh.cells[addr(r, c)])) : null); block.push(row); }
        clip = {block, r: s.r1, c: s.c1};
        const tsv = block.map((row, i) => row.map((cell, j) => fmt(eng.value(si, s.r1 + i, s.c1 + j), cell && cell.s && cell.s.nf)).join('\t')).join('\n');
        clip.text = tsv;
        if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(tsv).catch(() => {});
        if (cut) { snapshot(); forSel((r, c) => delete sh.cells[addr(r, c)]); changed(); }
      }
      let clip = null;
      function pasteBlock(){
        if (!clip) return; snapshot(); const sh = S();
        clip.block.forEach((row, i) => row.forEach((cell, j) => { const r = cur.r + i, c = cur.c + j, a = addr(r, c);
          if (!cell) { delete sh.cells[a]; return; }
          const n = JSON.parse(JSON.stringify(cell)); if (typeof n.v === 'string' && n.v[0] === '=') n.v = '=' + shiftRel(n.v.slice(1), r - (clip.r + i), c - (clip.c + j)); sh.cells[a] = n; }));
        changed();
      }
      function pasteText(text){
        const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
        snapshot(); rows.forEach((row, i) => row.forEach((v, j) => setCell(cur.r + i, cur.c + j, v))); changed();
      }

      // ---- events ----
      wrap.addEventListener('pointerdown', e => {
        const td = e.target.closest('td'); if (!td) { const th = e.target.closest('th'); if (th && th.dataset.c) { if (editing) commit(); anchor = {r: 1, c: +th.dataset.c}; select(renderRows, +th.dataset.c, true, true); cur = {r: 1, c: +th.dataset.c}; } else if (th && th.dataset.r) { if (editing) commit(); anchor = {r: +th.dataset.r, c: 1}; select(+th.dataset.r, S().cols, true, true); } return; }
        const r = +td.dataset.r, c = +td.dataset.c;
        if (editing && fx.value.startsWith('=') && e.pointerType === 'mouse') { // click a cell to insert its reference while typing a formula
          e.preventDefault(); const ins = addr(r, c); const p = fx.selectionStart; fx.value = fx.value.slice(0, p) + ins + fx.value.slice(fx.selectionEnd); fx.focus(); fx.setSelectionRange(p + ins.length, p + ins.length); return;
        }
        if (editing) commit();
        const ext = e.shiftKey || extendMode;
        select(r, c, ext, true);
        if (extendMode) { extendMode = false; tb.querySelector('[data-a="extend"]').classList.remove('active'); }
        if (e.pointerType === 'mouse') {
          wrap.focus({preventScroll: true});
          const move = ev => { const t = document.elementFromPoint(ev.clientX, ev.clientY); const td2 = t && t.closest && t.closest('td'); if (td2 && wrap.contains(td2)) select(+td2.dataset.r, +td2.dataset.c, true, true); };
          const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); };
          window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
        }
      });
      wrap.addEventListener('dblclick', e => { const td = e.target.closest('td'); if (td) startEdit(); });
      wrap.addEventListener('click', e => { if (e.target.classList.contains('more-rows')) renderMore(); });
      function renderMore(){ const sh = S(); const prev = renderRows; const tbody = wrap.querySelector('tbody'); const upto = Math.min(sh.rows, prev + 1000); let h = '';
        for (let r = prev + 1; r <= upto; r++) { h += `<tr><th class="rh" data-r="${r}">${r}</th>`; for (let c = 1; c <= sh.cols; c++) h += `<td data-r="${r}" data-c="${c}"></td>`; h += '</tr>'; }
        tbody.insertAdjacentHTML('beforeend', h);
        for (const td of tbody.querySelectorAll('td')) { const r = +td.dataset.r, c = +td.dataset.c; if (r > prev) (tds[r] || (tds[r] = []))[c] = td; }
        renderRows = upto; for (let r = prev + 1; r <= upto; r++) for (let c = 1; c <= sh.cols; c++) paint(tds[r][c], r, c);
        const mb = wrap.querySelector('.more-rows'); if (mb) { if (sh.rows > renderRows) mb.textContent = `Show more rows (${sh.rows - renderRows} hidden)`; else mb.remove(); } }
      wrap.addEventListener('keydown', e => {
        if (editing) return;
        const k = e.key, mod = e.ctrlKey || e.metaKey;
        const nav = {ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1]}[k];
        if (nav) { e.preventDefault(); select(cur.r + nav[0], cur.c + nav[1], e.shiftKey); return; }
        if (k === 'Enter') { e.preventDefault(); select(cur.r + (e.shiftKey ? -1 : 1), cur.c); return; }
        if (k === 'Tab') { e.preventDefault(); select(cur.r, cur.c + (e.shiftKey ? -1 : 1)); return; }
        if (k === 'F2') { e.preventDefault(); startEdit(); return; }
        if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); snapshot(); forSel((r, c) => setCell(r, c, '')); changed(); select(cur.r, cur.c, true, true); return; }
        if (mod && k.toLowerCase() === 'z') { e.preventDefault(); doUndo(e.shiftKey); return; }
        if (mod && k.toLowerCase() === 'y') { e.preventDefault(); doUndo(true); return; }
        if (mod && k.toLowerCase() === 'b') { e.preventDefault(); style(s => s.b ? delete s.b : s.b = 1); return; }
        if (mod && k.toLowerCase() === 'i') { e.preventDefault(); style(s => s.i ? delete s.i : s.i = 1); return; }
        if (mod && k.toLowerCase() === 'u') { e.preventDefault(); style(s => s.u ? delete s.u : s.u = 1); return; }
        if (mod && k.toLowerCase() === 'c') { copySel(false); return; }
        if (mod && k.toLowerCase() === 'x') { copySel(true); return; }
        if (!mod && !e.altKey && k.length === 1) { e.preventDefault(); startEdit(k); }
      });
      wrap.addEventListener('paste', e => { const t = e.clipboardData.getData('text/plain'); e.preventDefault(); if (clip && t === clip.text) pasteBlock(); else if (t) pasteText(t); });
      fx.addEventListener('focus', () => { editing = true; });
      fx.addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); commit('down'); wrap.focus({preventScroll: true}); }
        else if (e.key === 'Tab') { e.preventDefault(); commit('right'); wrap.focus({preventScroll: true}); }
        else if (e.key === 'Escape') { editing = false; select(cur.r, cur.c); wrap.focus({preventScroll: true}); }
      });
      fx.addEventListener('blur', () => setTimeout(() => { if (editing && document.activeElement !== fx) commit(); }, 150));
      function doUndo(isRedo){
        const from = isRedo ? redo : undo, to = isRedo ? undo : redo; if (!from.length) return;
        to.push(JSON.stringify(model)); const prev = JSON.parse(from.pop()); model.sheets = prev.sheets; si = Math.min(si, model.sheets.length - 1);
        eng = Engine(model); render(); ctx.changed();
      }
      tb.addEventListener('mousedown', e => { if (e.target.closest('button') && editing) e.preventDefault(); });
      tb.addEventListener('click', e => {
        const b = e.target.closest('button'); if (!b) return; const a = b.dataset.a;
        if (editing && a !== 'sum') commit();
        if (a === 'undo') doUndo(false); else if (a === 'redo') doUndo(true);
        else if (a === 'b' || a === 'i' || a === 'u') { const on = !((S().cells[addr(cur.r, cur.c)] || {}).s || {})[a]; style(s => on ? s[a] = 1 : delete s[a]); }
        else if (a === 'left' || a === 'center' || a === 'right') style(s => s.al === a ? delete s.al : s.al = a);
        else if (a === 'sum') {
          const s = selRect(); let r0 = cur.r - 1; const sh = S();
          if (s.r1 !== s.r2 || s.c1 !== s.c2) { snapshot(); for (let c = s.c1; c <= s.c2; c++) setCell(s.r2 + 1, c, `=SUM(${addr(s.r1, c)}:${addr(s.r2, c)})`); changed(); select(s.r2 + 1, s.c1); return; }
          let top = r0; while (top >= 1 && sh.cells[addr(top, cur.c)] && sh.cells[addr(top, cur.c)].v !== '') top--;
          const f = r0 >= 1 && top < r0 ? `=SUM(${addr(top + 1, cur.c)}:${addr(r0, cur.c)})` : '=SUM()';
          startEdit(f);
        }
        else if (a === 'filldown') {
          const s = selRect(); if (s.r1 === s.r2) return ctx.toast('Select a range first: the top row is copied down.');
          snapshot(); const sh = S();
          for (let c = s.c1; c <= s.c2; c++) { const src = sh.cells[addr(s.r1, c)]; for (let r = s.r1 + 1; r <= s.r2; r++) { const a2 = addr(r, c); if (!src) { delete sh.cells[a2]; continue; } const n = JSON.parse(JSON.stringify(src)); if (typeof n.v === 'string' && n.v[0] === '=') n.v = '=' + shiftRel(n.v.slice(1), r - s.r1, 0); sh.cells[a2] = n; } }
          changed();
        }
        else if (a === 'extend') { extendMode = !extendMode; b.classList.toggle('active', extendMode); if (extendMode) ctx.toast('Now tap the other corner of the range.'); }
        else if (a === 'insrow') shiftAll('r', selRect().r1, 1);
        else if (a === 'delrow') shiftAll('r', selRect().r1, -1);
        else if (a === 'inscol') shiftAll('c', selRect().c1, 1);
        else if (a === 'delcol') shiftAll('c', selRect().c1, -1);
        else if (a === 'wider' || a === 'narrower') { snapshot(); const s = selRect(); for (let c = s.c1; c <= s.c2; c++) { const k = colName(c); S().colW[k] = Math.max(30, Math.min(600, (S().colW[k] || DEF_W) + (a === 'wider' ? 30 : -30))); } render(); ctx.changed(); }
        else if (a === 'sortasc' || a === 'sortdesc') sortSel(a === 'sortdesc');
        else if (a === 'clear') { snapshot(); forSel((r, c) => delete S().cells[addr(r, c)]); changed(); }
      });
      tb.querySelector('[data-nf]').onchange = e => { const v = e.target.value; style(s => v === 'General' ? delete s.nf : s.nf = v); };
      tb.querySelector('[data-fn]').onchange = e => { const f = e.target.value; e.target.value = ''; if (!f) return; const hint = {VLOOKUP: '(value, range, column, FALSE)', XLOOKUP: '(value, lookup_range, return_range)', IF: '(test, if_true, if_false)', COUNTIF: '(range, ">0")', SUMIF: '(range, ">0")', ROUND: '(number, 2)', TODAY: '()'}[f]; startEdit('=' + f + (hint && f === 'TODAY' ? '()' : '(')); if (hint && f !== 'TODAY') ctx.toast(f + hint); };
      tb.querySelectorAll('[data-color]').forEach(inp => inp.onchange = e => { const k = inp.dataset.color, v = e.target.value; style(s => (k === 'bg' && v === '#ffffff') ? delete s[k] : s[k] = v); });
      tabs.addEventListener('click', e => {
        const b = e.target.closest('button'); if (!b) return; if (editing) commit(); const s = b.dataset.s;
        if (s === 'add') { snapshot(); let n = model.sheets.length + 1; while (model.sheets.some(x => x.name === 'Sheet' + n)) n++; model.sheets.push(blankSheet('Sheet' + n)); si = model.sheets.length - 1; }
        else if (s === 'ren') { const n = prompt('Sheet name', S().name); if (!n || !n.trim()) return; const old = S().name; snapshot(); S().name = n.trim().slice(0, 31);
          const q = x => /^[A-Za-z_][A-Za-z0-9_.]*$/.test(x) ? x : `'${x}'`;
          model.sheets.forEach(sh => { for (const a in sh.cells) { const c = sh.cells[a]; if (typeof c.v === 'string' && c.v[0] === '=') c.v = '=' + c.v.slice(1).split(/("(?:[^"]|"")*")/).map((seg, i) => i % 2 ? seg : seg.split(q(old) + '!').join(q(S().name) + '!')).join(''); } }); }
        else if (s === 'del') { if (model.sheets.length < 2) return ctx.toast('A workbook needs at least one sheet.'); if (!confirm(`Delete sheet "${S().name}"?`)) return; snapshot(); model.sheets.splice(si, 1); si = Math.max(0, si - 1); }
        else si = +s;
        model.active = si; eng = Engine(model); cur = {r: 1, c: 1}; anchor = {r: 1, c: 1}; render(); ctx.changed();
      });
      render();
      wrap.focus({preventScroll: true});

      return {
        getContent: () => { model.active = si; return JSON.parse(JSON.stringify(model)); },
        exportBlob: () => exportXlsx(model),
        extraMenu: [{label: 'Export this sheet as CSV', fn: () => {
          const sh = S(), u = usedRange(sh); const lines = [];
          for (let r = 1; r <= u.r; r++) { const row = []; for (let c = 1; c <= u.c; c++) { const v = eng.value(si, r, c); row.push(csvCell(v == null ? '' : isErr(v) ? errText(v) : String(v))); } lines.push(row.join(',')); }
          ctx.download(new Blob(['\ufeff' + lines.join('\r\n')], {type: 'text/csv'}), `${file.name} - ${sh.name}.csv`);
        }}],
        printView(){
          const sh = S(), u = usedRange(sh); let h = '<div class="print-sheet"><table>';
          for (let r = 1; r <= Math.max(u.r, 1); r++) { h += '<tr>'; for (let c = 1; c <= Math.max(u.c, 1); c++) { const cell = sh.cells[addr(r, c)]; const s = (cell && cell.s) || {}; const v = eng.value(si, r, c);
            let st = (s.b ? 'font-weight:700;' : '') + (s.i ? 'font-style:italic;' : '') + (s.bg ? `background:${s.bg};` : '') + (s.fc ? `color:${s.fc};` : '') + `text-align:${s.al || (typeof v === 'number' ? 'right' : 'left')};`;
            h += `<td style="${st}">${ctx.esc(fmt(v, s.nf))}</td>`; } h += '</tr>'; }
          return {html: h + '</table></div>', page: 'size:A4 landscape;margin:12mm', css: '@media print{.print-sheet td{-webkit-print-color-adjust:exact;print-color-adjust:exact}}'};
        },
        _engineValue: (sheetIndex, a) => { const p = parseAddr(a); return eng.value(sheetIndex, p.r, p.c); },
        destroy(){},
      };
    },
    _fmt: fmt,
  };
  MOApps.sheet = app;
})();
