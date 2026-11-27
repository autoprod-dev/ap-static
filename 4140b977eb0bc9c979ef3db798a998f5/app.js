/* NFC Tags v0.1.0 — read and write NDEF tags only. Cannot clone or emulate cards. */
(() => {
  const VERSION = '0.1.0';
  const LIB_KEY = 'nfc-tags-library-v1';
  const SCAN_WAIT_MS = 15000;

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

  /* ---------- capability ---------- */
  function isAppleDevice() {
    const ua = navigator.userAgent || '';
    if (/iPhone|iPad|iPod/i.test(ua)) return true;
    // iPadOS 13+ can report as Macintosh
    if (/Macintosh/i.test(ua) && navigator.maxTouchPoints > 1) return true;
    return false;
  }
  function hasNdef() {
    return typeof window.NDEFReader === 'function';
  }
  function isSecureContextOk() {
    return window.isSecureContext === true;
  }
  function capability() {
    if (isAppleDevice()) {
      return {
        level: 'bad',
        title: 'This will not work on iPhone',
        body: 'Apple does not let websites use NFC. There is no way for this app to read a tag, card or fob in Safari or Chrome on iPhone or iPad.',
        canScan: false,
        reason: 'ios'
      };
    }
    if (!isSecureContextOk()) {
      return {
        level: 'warn',
        title: 'This page is not on a secure link',
        body: 'Chrome only allows NFC on a secure (https) page. Open the private link you were given, then try again.',
        canScan: false,
        reason: 'insecure'
      };
    }
    if (!hasNdef()) {
      return {
        level: 'warn',
        title: 'NFC is not available in this browser',
        body: 'Reading and writing tags only works in Chrome on an Android phone. On a computer, or in another Android browser, you can still look through the app and use your saved library, but this device cannot talk to a tag.',
        canScan: false,
        reason: 'no-ndef'
      };
    }
    return {
      level: 'ok',
      title: 'This phone can use ordinary NFC tags',
      body: 'Chrome on this Android phone can read and write NDEF tags — the kind used for links, notes, contact cards and Wi-Fi stickers. It cannot copy keycards or most door fobs, and it cannot make your phone act as a card.',
      canScan: true,
      reason: 'ready'
    };
  }

  /* ---------- records ---------- */
  function bytesOf(record) {
    if (!record || !record.data) return new Uint8Array();
    const d = record.data;
    if (d instanceof DataView) return new Uint8Array(d.buffer, d.byteOffset, d.byteLength);
    if (d instanceof ArrayBuffer) return new Uint8Array(d);
    if (d instanceof Uint8Array) return d;
    return new Uint8Array();
  }
  function toHex(bytes) {
    return [...bytes].map((b) => b.toString(16).padStart(2, '0')).join(' ');
  }
  function decodeTextBytes(bytes, encoding) {
    try {
      return new TextDecoder(encoding || 'utf-8').decode(bytes);
    } catch (err) {
      try { return new TextDecoder('utf-8').decode(bytes); } catch (e2) { return ''; }
    }
  }
  function decodeRecord(record) {
    const out = {
      recordType: record.recordType || 'unknown',
      mediaType: record.mediaType || '',
      id: record.id || '',
      encoding: record.encoding || '',
      lang: record.lang || '',
      text: '',
      url: '',
      hex: '',
      byteLength: 0,
      label: '',
      summary: ''
    };
    const bytes = bytesOf(record);
    out.byteLength = bytes.length;

    if (out.recordType === 'empty') {
      out.label = 'Empty';
      out.summary = 'An empty record.';
      return out;
    }
    if (out.recordType === 'text') {
      out.text = decodeTextBytes(bytes, out.encoding);
      out.label = 'Text';
      out.summary = out.text;
      return out;
    }
    if (out.recordType === 'url' || out.recordType === 'absolute-url') {
      out.url = decodeTextBytes(bytes, 'utf-8');
      out.label = 'Link';
      out.summary = out.url;
      return out;
    }
    if (out.recordType === 'smart-poster') {
      out.hex = toHex(bytes);
      out.label = 'Smart poster';
      out.summary = 'A smart poster, usually a link with a title.';
      return out;
    }
    if (out.recordType === 'mime') {
      const mime = (out.mediaType || '').toLowerCase();
      if (mime === 'text/vcard' || mime === 'text/x-vcard' || mime === 'text/directory') {
        out.text = decodeTextBytes(bytes, 'utf-8');
        out.label = 'Contact card';
        out.summary = vcardSummary(out.text) || out.text;
        return out;
      }
      if (mime === 'application/vnd.wfa.wsc') {
        out.hex = toHex(bytes);
        const wifi = decodeWifi(bytes);
        out.text = wifi.ssid ? `Wi-Fi network: ${wifi.ssid}` : 'Wi-Fi details';
        out.label = 'Wi-Fi';
        out.summary = out.text + (wifi.security ? ` (${wifi.security})` : '');
        out.wifi = wifi;
        return out;
      }
      if (mime.startsWith('text/') || mime === 'application/json') {
        out.text = decodeTextBytes(bytes, 'utf-8');
        out.label = mime;
        out.summary = out.text;
        return out;
      }
      out.hex = toHex(bytes);
      out.label = mime || 'File data';
      out.summary = `${out.byteLength} bytes of ${out.label}.`;
      return out;
    }
    // External types such as android.com:pkg, or unknown
    const maybeText = decodeTextBytes(bytes, 'utf-8');
    const printable = maybeText && /^[\x09\x0a\x0d\x20-\x7e\u00a0-\uFFFF]*$/.test(maybeText);
    if (printable && maybeText.trim()) out.text = maybeText;
    else out.hex = toHex(bytes);
    out.label = out.recordType === 'unknown' ? 'Unknown record' : out.recordType;
    out.summary = out.text || (out.byteLength ? `${out.byteLength} bytes.` : 'No data.');
    return out;
  }

  function vcardSummary(text) {
    const pick = (key) => {
      const re = new RegExp('^' + key + '(?:;[^:]*)?:\\s*(.*)$', 'im');
      const m = text.match(re);
      return m ? m[1].trim() : '';
    };
    const parts = [pick('FN') || pick('N'), pick('TEL'), pick('EMAIL')].filter(Boolean);
    return parts.join(' · ');
  }

  function decodeWifi(bytes) {
    const info = { ssid: '', security: '' };
    const view = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    const tlvs = readTlvs(view);
    const walk = (list) => {
      for (const t of list) {
        if (t.id === 0x1045) info.ssid = decodeTextBytes(t.value, 'utf-8');
        if (t.id === 0x1003) {
          const v = t.value.length >= 2 ? (t.value[0] << 8) | t.value[1] : 0;
          info.security = v === 0x0001 ? 'open' : 'password-protected';
        }
        if (t.id === 0x100e) walk(readTlvs(t.value));
      }
    };
    walk(tlvs);
    return info;
  }
  function readTlvs(bytes) {
    const out = [];
    let i = 0;
    const b = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
    while (i + 4 <= b.length) {
      const id = (b[i] << 8) | b[i + 1];
      const len = (b[i + 2] << 8) | b[i + 3];
      i += 4;
      if (i + len > b.length) break;
      out.push({ id, value: b.slice(i, i + len) });
      i += len;
    }
    return out;
  }

  function describeRecords(records) {
    if (!records.length) return 'an NFC tag with no readable records';
    const labels = records.map((r) => r.label.toLowerCase());
    const uniq = [...new Set(labels)];
    if (uniq.length === 1) return uniq[0] === 'link' ? 'a link tag' : `a ${uniq[0]} tag`;
    return 'an NFC tag with ' + uniq.join(', ');
  }

  /* ---------- wifi / vcard encode ---------- */
  function u16(n) { return [(n >> 8) & 0xff, n & 0xff]; }
  function tlv(id, payload) {
    const p = payload instanceof Uint8Array ? [...payload] : payload;
    return [...u16(id), ...u16(p.length), ...p];
  }
  function strBytes(s) { return [...new TextEncoder().encode(s)]; }
  function encodeWifi(ssid, password, security) {
    const open = security === 'open';
    const auth = open ? 0x0001 : 0x0020; // Open or WPA2-Personal
    const enc = open ? 0x0001 : 0x0008;  // None or AES
    const cred = [
      ...tlv(0x1026, [0x01]),
      ...tlv(0x1045, strBytes(ssid)),
      ...tlv(0x1003, u16(auth)),
      ...tlv(0x100f, u16(enc)),
      ...tlv(0x1027, strBytes(open ? '' : password)),
      ...tlv(0x1020, [0xff, 0xff, 0xff, 0xff, 0xff, 0xff])
    ];
    return new Uint8Array([...tlv(0x104a, [0x10]), ...tlv(0x100e, cred)]);
  }
  function escapeVcard(s) {
    return String(s || '').replace(/\\/g, '\\\\').replace(/,/g, '\\,').replace(/;/g, '\\;').replace(/\n/g, '\\n');
  }
  function encodeVcard(name, tel, email) {
    const lines = ['BEGIN:VCARD', 'VERSION:3.0'];
    if (name) lines.push('FN:' + escapeVcard(name));
    if (tel) lines.push('TEL:' + escapeVcard(tel));
    if (email) lines.push('EMAIL:' + escapeVcard(email));
    lines.push('END:VCARD');
    return lines.join('\r\n');
  }

  function normaliseUrl(raw) {
    const s = String(raw || '').trim();
    if (!s) return '';
    if (/^(https?:|mailto:|tel:|sms:)/i.test(s)) return s;
    if (/^\/\//.test(s)) return 'https:' + s;
    return 'https://' + s;
  }

  function buildWrite(kind, fields) {
    if (kind === 'text') {
      const text = String(fields.text || '').trim();
      if (!text) return { error: 'Enter some text to write.' };
      return {
        records: [{ recordType: 'text', data: text }],
        stored: [{ recordType: 'text', text, label: 'Text', summary: text }],
        title: text.slice(0, 48)
      };
    }
    if (kind === 'url') {
      const url = normaliseUrl(fields.url);
      if (!url) return { error: 'Enter a link to write.' };
      try { new URL(url); } catch (e) { return { error: 'That link does not look quite right.' }; }
      return {
        records: [{ recordType: 'url', data: url }],
        stored: [{ recordType: 'url', url, label: 'Link', summary: url }],
        title: url
      };
    }
    if (kind === 'contact') {
      const name = String(fields.name || '').trim();
      const tel = String(fields.tel || '').trim();
      const email = String(fields.email || '').trim();
      if (!name && !tel && !email) return { error: 'Enter a name, phone or email.' };
      const data = encodeVcard(name, tel, email);
      const summary = [name, tel, email].filter(Boolean).join(' · ');
      return {
        records: [{ recordType: 'mime', mediaType: 'text/vcard', data }],
        stored: [{ recordType: 'mime', mediaType: 'text/vcard', text: data, label: 'Contact card', summary }],
        title: name || summary
      };
    }
    if (kind === 'wifi') {
      const ssid = String(fields.ssid || '').trim();
      const security = fields.security === 'open' ? 'open' : 'wpa2';
      const password = String(fields.password || '');
      if (!ssid) return { error: 'Enter the network name (SSID).' };
      if (security !== 'open' && !password) return { error: 'Enter the Wi-Fi password, or choose Open.' };
      const bytes = encodeWifi(ssid, password, security);
      const summary = `Wi-Fi network: ${ssid}` + (security === 'open' ? ' (open)' : ' (password-protected)');
      return {
        records: [{ recordType: 'mime', mediaType: 'application/vnd.wfa.wsc', data: bytes }],
        stored: [{ recordType: 'mime', mediaType: 'application/vnd.wfa.wsc', hex: toHex(bytes), byteLength: bytes.length, label: 'Wi-Fi', summary, wifi: { ssid, security: security === 'open' ? 'open' : 'password-protected' } }],
        title: 'Wi-Fi: ' + ssid
      };
    }
    return { error: 'Choose what to write.' };
  }

  function storedToWriteRecords(stored) {
    return stored.map((r) => {
      if (r.recordType === 'text') return { recordType: 'text', data: r.text || '' };
      if (r.recordType === 'url' || r.recordType === 'absolute-url') return { recordType: 'url', data: r.url || r.text || '' };
      if (r.recordType === 'mime') {
        const rec = { recordType: 'mime', mediaType: r.mediaType || 'application/octet-stream' };
        if (r.text) rec.data = r.text;
        else if (r.hex) rec.data = hexToBytes(r.hex);
        else rec.data = '';
        return rec;
      }
      if (r.text) return { recordType: r.recordType || 'text', data: r.text };
      if (r.url) return { recordType: 'url', data: r.url };
      return { recordType: 'empty' };
    });
  }
  function hexToBytes(hex) {
    const parts = String(hex).trim().split(/\s+/).filter(Boolean);
    return new Uint8Array(parts.map((p) => parseInt(p, 16)));
  }

  /* ---------- library ---------- */
  function loadLibrary() {
    try {
      const raw = localStorage.getItem(LIB_KEY);
      const data = raw ? JSON.parse(raw) : [];
      return Array.isArray(data) ? data : [];
    } catch (err) {
      return [];
    }
  }
  function saveLibrary(items) {
    localStorage.setItem(LIB_KEY, JSON.stringify(items));
  }
  function uid() {
    return 't_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8);
  }
  function addToLibrary(entry) {
    const items = loadLibrary();
    const item = {
      id: entry.id || uid(),
      name: entry.name || 'Saved tag',
      savedAt: entry.savedAt || new Date().toISOString(),
      serialNumber: entry.serialNumber || '',
      source: entry.source || 'read',
      records: entry.records || []
    };
    items.unshift(item);
    saveLibrary(items);
    return item;
  }
  function removeFromLibrary(id) {
    saveLibrary(loadLibrary().filter((x) => x.id !== id));
  }
  function renameInLibrary(id, name) {
    const items = loadLibrary();
    const it = items.find((x) => x.id === id);
    if (it) { it.name = name; saveLibrary(items); }
  }
  function exportLibrary() {
    const payload = {
      app: 'NFC Tags',
      version: VERSION,
      exportedAt: new Date().toISOString(),
      items: loadLibrary()
    };
    return JSON.stringify(payload, null, 2);
  }
  function importLibrary(text, { merge = true } = {}) {
    let data;
    try { data = JSON.parse(text); } catch (e) { return { error: 'That file is not valid JSON.' }; }
    let items = [];
    if (Array.isArray(data)) items = data;
    else if (data && Array.isArray(data.items)) items = data.items;
    else return { error: 'That file has no saved tags in it.' };
    const clean = [];
    for (const it of items) {
      if (!it || typeof it !== 'object') continue;
      clean.push({
        id: typeof it.id === 'string' && it.id ? it.id : uid(),
        name: String(it.name || 'Saved tag'),
        savedAt: it.savedAt || new Date().toISOString(),
        serialNumber: String(it.serialNumber || ''),
        source: it.source === 'write' ? 'write' : 'read',
        records: Array.isArray(it.records) ? it.records : []
      });
    }
    if (!merge) { saveLibrary(clean); return { added: clean.length, total: clean.length }; }
    const existing = loadLibrary();
    const ids = new Set(existing.map((x) => x.id));
    let added = 0;
    for (const it of clean) {
      if (ids.has(it.id)) it.id = uid();
      existing.unshift(it);
      ids.add(it.id);
      added += 1;
    }
    saveLibrary(existing);
    return { added, total: existing.length };
  }

  /* ---------- NFC engine ---------- */
  let reader = null;
  let scanAbort = null;
  let scanBusy = false;
  let writeBusy = false;
  let scanHandlers = { onReading: null, onReadingError: null };

  function getReader() {
    if (!hasNdef()) throw new Error('NFC is not available in this browser.');
    if (!reader) reader = new NDEFReader();
    return reader;
  }

  function onReaderReading(event) {
    const records = [];
    try {
      const list = (event.message && event.message.records) || [];
      for (const rec of list) records.push(decodeRecord(rec));
    } catch (err) {
      records.push({ recordType: 'unknown', label: 'Unread record', summary: 'The tag was found, but a record could not be decoded.', text: '', url: '', hex: '' });
    }
    scanHandlers.onReading && scanHandlers.onReading({
      serialNumber: event.serialNumber || '',
      records,
      rawEvent: event
    });
  }
  function onReaderError() {
    scanHandlers.onReadingError && scanHandlers.onReadingError();
  }

  async function startScan({ onReading, onReadingError, onStarted }) {
    const cap = capability();
    if (!cap.canScan) return { error: cap.body, reason: cap.reason };
    scanHandlers.onReading = onReading;
    scanHandlers.onReadingError = onReadingError;
    if (scanBusy) {
      onStarted && onStarted();
      return { ok: true, already: true };
    }
    const r = getReader();
    if (scanAbort) {
      try { scanAbort.abort(); } catch (e) {}
    }
    scanAbort = new AbortController();
    r.addEventListener('reading', onReaderReading);
    r.addEventListener('readingerror', onReaderError);
    try {
      await r.scan({ signal: scanAbort.signal });
      scanBusy = true;
      onStarted && onStarted();
      const stop = () => {
        r.removeEventListener('reading', onReaderReading);
        r.removeEventListener('readingerror', onReaderError);
        try { scanAbort && scanAbort.abort(); } catch (e) {}
        scanBusy = false;
        scanHandlers = { onReading: null, onReadingError: null };
      };
      return { ok: true, stop };
    } catch (err) {
      r.removeEventListener('reading', onReaderReading);
      r.removeEventListener('readingerror', onReaderError);
      scanBusy = false;
      return { error: explainNfcError(err, 'scan'), name: err && err.name };
    }
  }

  async function writeRecords(records) {
    const cap = capability();
    if (!cap.canScan) return { error: cap.body, reason: cap.reason };
    if (writeBusy) return { error: 'A write is already in progress. Hold the tag still, or wait a moment.' };
    writeBusy = true;
    try {
      const r = getReader();
      await r.write({ records });
      return { ok: true };
    } catch (err) {
      return { error: explainNfcError(err, 'write'), name: err && err.name };
    } finally {
      writeBusy = false;
    }
  }

  function explainNfcError(err, action) {
    const name = err && err.name;
    const msg = (err && err.message) || '';
    if (name === 'NotAllowedError' || name === 'SecurityError') {
      return 'Chrome did not get permission to use NFC. When the prompt appears, choose Allow.';
    }
    if (name === 'NotSupportedError') {
      return 'This browser cannot use NFC. Open this page in Chrome on an Android phone.';
    }
    if (name === 'NotReadableError') {
      return action === 'write'
        ? 'The tag could not be written. It may be locked, damaged, or a type this app cannot use.'
        : 'The tag was found, but could not be read as an ordinary NDEF tag.';
    }
    if (name === 'NetworkError' || /tag/i.test(msg) && /not supported|incompatible/i.test(msg)) {
      return 'This tag is not one Chrome can use. Many keycards and some older tags fall into this group.';
    }
    if (name === 'AbortError') {
      return action === 'write' ? 'The write was cancelled.' : 'The scan was cancelled.';
    }
    if (name === 'InvalidStateError') {
      return 'NFC is already in use. Close any other NFC prompt, then try again.';
    }
    if (action === 'write') {
      return 'The tag could not be written. Use a blank writable NFC sticker, and hold it still against the back of your phone, near the top.';
    }
    return 'NFC could not be started. Check that NFC is turned on in this phone’s settings, then try again.';
  }

  /* ---------- verdicts ---------- */
  function verdictForReading({ records, serialNumber }) {
    const kind = describeRecords(records);
    return {
      level: 'ok',
      title: 'This is an ordinary NFC tag',
      body: `This app can read it. It looks like ${kind}.` +
        (serialNumber ? ' The tag’s serial number is shown below. Chrome only shares that number when it can already read the tag as NDEF.' : ' Chrome did not share a serial number for this tag.') +
        ' You can save the records to your library, or write the same kind of thing onto a blank sticker. This app still cannot make your phone act as this tag, and it cannot copy a keycard.'
    };
  }
  function verdictForReadingError() {
    return {
      level: 'bad',
      title: 'A tag was found, but this app cannot read it',
      body: 'Your phone noticed something, but Chrome could not open it as an ordinary NDEF tag. This is what usually happens with office or hotel keycards, and with many encrypted access cards (for example MIFARE Classic or DESFire). This app cannot copy, clone or emulate that card.'
    };
  }
  function verdictForTimeout() {
    return {
      level: 'warn',
      title: 'Nothing was detected',
      body: 'No tag answered in time. Move it around the back of the phone, near the top, and hold it still. If it still does nothing, it may be a low-frequency (125 kHz) fob or a car remote, which phones cannot read at all. It may also be a card this app cannot see.'
    };
  }

  /* ---------- UI ---------- */
  function toast(msg) {
    const el = $('#toast');
    el.textContent = msg;
    el.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { el.hidden = true; }, 2800);
  }
  function setCapCard() {
    const cap = capability();
    const el = $('#cap-card');
    el.className = 'card cap ' + cap.level;
    el.innerHTML = `<h2>${escapeHtml(cap.title)}</h2><p>${escapeHtml(cap.body)}</p>`;
    $('#check-btn').disabled = !cap.canScan;
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function showTab(name) {
    $$('.tab').forEach((b) => {
      const on = b.dataset.tab === name;
      b.classList.toggle('on', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
    });
    $$('.view').forEach((v) => { v.hidden = v.dataset.view !== name; });
  }
  function renderRecords(records, serialNumber) {
    const parts = [];
    if (serialNumber) parts.push(`<p class="meta">Tag serial number: <strong>${escapeHtml(serialNumber)}</strong></p>`);
    else parts.push(`<p class="meta">No serial number was shared by the browser.</p>`);
    if (!records.length) {
      parts.push('<p class="muted">The tag opened, but it has no records.</p>');
      return parts.join('');
    }
    parts.push('<ul class="records">');
    for (const r of records) {
      parts.push('<li class="rec">');
      parts.push(`<p class="kind">${escapeHtml(r.label || r.recordType)}</p>`);
      if (r.url) parts.push(`<p class="body"><a href="${escapeHtml(r.url)}" rel="noreferrer noopener">${escapeHtml(r.url)}</a></p>`);
      else if (r.text) parts.push(`<pre class="body">${escapeHtml(r.text)}</pre>`);
      else if (r.hex) parts.push(`<pre class="body">${escapeHtml(r.hex)}</pre>`);
      else parts.push(`<p class="body">${escapeHtml(r.summary || '')}</p>`);
      if (r.lang) parts.push(`<p class="meta">Language: ${escapeHtml(r.lang)}</p>`);
      parts.push('</li>');
    }
    parts.push('</ul>');
    return parts.join('');
  }
  function renderVerdict(v) {
    return `<article class="card verdict ${v.level}"><h2>${escapeHtml(v.title)}</h2><p>${escapeHtml(v.body)}</p></article>`;
  }

  let lastReading = null;
  let waitTimer = null;
  let activeStop = null;

  function setScanStatus(el, text, on) {
    if (!on) { el.hidden = true; el.innerHTML = ''; return; }
    el.hidden = false;
    el.innerHTML = `<span class="pulse" aria-hidden="true"></span><span>${escapeHtml(text)}</span>`;
  }

  function showCheckResult(html) {
    const box = $('#check-result');
    box.hidden = false;
    box.innerHTML = html;
  }

  function handleReading(reading) {
    lastReading = reading;
    if (waitTimer) { clearTimeout(waitTimer); waitTimer = null; }
    const v = verdictForReading(reading);
    const actions = `<div class="result-actions">
      <button type="button" class="btn primary" id="save-reading">Save to library</button>
    </div>`;
    showCheckResult(renderVerdict(v) + `<div class="card">${renderRecords(reading.records, reading.serialNumber)}${actions}</div>`);
    const btn = $('#save-reading');
    if (btn) btn.addEventListener('click', () => {
      const name = defaultName(reading);
      addToLibrary({ name, serialNumber: reading.serialNumber, source: 'read', records: reading.records });
      renderLibrary();
      toast('Saved to your library on this device.');
    });
    $('#limits-box').open = false;
  }
  function defaultName(reading) {
    if (reading.records && reading.records[0] && reading.records[0].summary) {
      return String(reading.records[0].summary).slice(0, 60);
    }
    return 'Saved tag';
  }

  async function beginCheck() {
    const cap = capability();
    if (!cap.canScan) {
      showCheckResult(renderVerdict({ level: cap.level, title: cap.title, body: cap.body }));
      return;
    }
    $('#check-result').hidden = true;
    lastReading = null;
    setScanStatus($('#scan-status'), 'Listening… hold the card or tag still against the back, near the top.', true);
    $('#check-btn').disabled = true;
    let gotResult = false;
    const result = await startScan({
      onReading: (reading) => { gotResult = true; handleReading(reading); },
      onReadingError: () => {
        gotResult = true;
        if (waitTimer) { clearTimeout(waitTimer); waitTimer = null; }
        showCheckResult(renderVerdict(verdictForReadingError()));
      },
      onStarted: () => {}
    });
    if (result.error) {
      setScanStatus($('#scan-status'), '', false);
      $('#check-btn').disabled = false;
      showCheckResult(renderVerdict({ level: 'bad', title: 'NFC could not be started', body: result.error }));
      return;
    }
    if (waitTimer) clearTimeout(waitTimer);
    waitTimer = setTimeout(() => {
      if (!gotResult) showCheckResult(renderVerdict(verdictForTimeout()));
    }, SCAN_WAIT_MS);
    $('#check-btn').disabled = false;
    $('#check-btn').textContent = 'Check again';
    if (result.stop) activeStop = result.stop;
  }

  function renderLibrary() {
    const items = loadLibrary();
    const empty = $('#lib-empty');
    const list = $('#lib-list');
    if (!items.length) {
      empty.hidden = false;
      list.hidden = true;
      list.innerHTML = '';
      return;
    }
    empty.hidden = true;
    list.hidden = false;
    list.innerHTML = items.map((it) => {
      const when = formatWhen(it.savedAt);
      const summary = (it.records && it.records[0] && it.records[0].summary) || (it.serialNumber ? 'Serial ' + it.serialNumber : 'Saved tag');
      return `<li class="lib-item" data-id="${escapeHtml(it.id)}">
        <h3>${escapeHtml(it.name)}</h3>
        <p class="meta">${escapeHtml(when)} · ${escapeHtml(it.source === 'write' ? 'Written / composed' : 'Read')} · ${escapeHtml(summary)}</p>
        ${it.serialNumber ? `<p class="meta">Serial: ${escapeHtml(it.serialNumber)}</p>` : ''}
        <div class="actions">
          <button type="button" class="btn" data-act="view">View</button>
          <button type="button" class="btn" data-act="write">Write to a tag</button>
          <button type="button" class="btn" data-act="rename">Rename</button>
          <button type="button" class="btn danger" data-act="delete">Delete</button>
        </div>
        <div class="lib-detail" hidden></div>
      </li>`;
    }).join('');
  }
  function formatWhen(iso) {
    try {
      const d = new Date(iso);
      return d.toLocaleString('en-AU', { dateStyle: 'medium', timeStyle: 'short' });
    } catch (e) { return iso || ''; }
  }

  function bindLibraryClicks() {
    $('#lib-list').addEventListener('click', async (ev) => {
      const btn = ev.target.closest('button[data-act]');
      if (!btn) return;
      const itemEl = ev.target.closest('.lib-item');
      const id = itemEl && itemEl.dataset.id;
      const item = loadLibrary().find((x) => x.id === id);
      if (!item) return;
      const act = btn.dataset.act;
      if (act === 'view') {
        const detail = itemEl.querySelector('.lib-detail');
        const open = detail.hidden;
        detail.hidden = !open;
        if (open) detail.innerHTML = renderRecords(item.records || [], item.serialNumber || '');
        return;
      }
      if (act === 'rename') {
        const name = prompt('Name for this saved tag', item.name);
        if (name && name.trim()) { renameInLibrary(id, name.trim()); renderLibrary(); }
        return;
      }
      if (act === 'delete') {
        if (confirm('Delete this saved tag from this device?')) { removeFromLibrary(id); renderLibrary(); toast('Deleted.'); }
        return;
      }
      if (act === 'write') {
        const recs = storedToWriteRecords(item.records || []);
        if (!recs.length) { toast('This saved item has nothing that can be written.'); return; }
        btn.disabled = true;
        toast('Hold a writable tag against the back of your phone, near the top.');
        const res = await writeRecords(recs);
        btn.disabled = false;
        toast(res.ok ? 'Written to the tag.' : res.error);
      }
    });
  }

  function wireWriteForm() {
    const kind = $('#write-kind');
    const show = () => {
      $('#fields-text').hidden = kind.value !== 'text';
      $('#fields-url').hidden = kind.value !== 'url';
      $('#fields-contact').hidden = kind.value !== 'contact';
      $('#fields-wifi').hidden = kind.value !== 'wifi';
    };
    kind.addEventListener('change', show);
    show();
    $('#wifi-sec').addEventListener('change', () => {
      const open = $('#wifi-sec').value === 'open';
      $('#wifi-pass').disabled = open;
      if (open) $('#wifi-pass').value = '';
    });
    $('#write-btn').addEventListener('click', onWrite);
  }

  async function onWrite() {
    const kind = $('#write-kind').value;
    const built = buildWrite(kind, {
      text: $('#text-value').value,
      url: $('#url-value').value,
      name: $('#contact-name').value,
      tel: $('#contact-tel').value,
      email: $('#contact-email').value,
      ssid: $('#wifi-ssid').value,
      security: $('#wifi-sec').value,
      password: $('#wifi-pass').value
    });
    const out = $('#write-result');
    if (built.error) {
      out.hidden = false;
      out.innerHTML = renderVerdict({ level: 'warn', title: 'Nothing was written', body: built.error });
      return;
    }
    const cap = capability();
    if ($('#write-save').checked) {
      addToLibrary({ name: built.title, source: 'write', records: built.stored });
      renderLibrary();
    }
    if (!cap.canScan) {
      out.hidden = false;
      out.innerHTML = renderVerdict({
        level: 'warn',
        title: $('#write-save').checked ? 'Saved to your library, but not written' : 'This device cannot write a tag',
        body: cap.body
      });
      return;
    }
    setScanStatus($('#write-status'), 'Hold a blank writable tag against the back of your phone, near the top.', true);
    $('#write-btn').disabled = true;
    const res = await writeRecords(built.records);
    $('#write-btn').disabled = false;
    setScanStatus($('#write-status'), '', false);
    out.hidden = false;
    if (res.ok) {
      out.innerHTML = renderVerdict({
        level: 'ok',
        title: 'Written to the tag',
        body: 'The tag now holds your ' + (built.stored[0] && built.stored[0].label ? built.stored[0].label.toLowerCase() : 'record') + '. You can check it on the Check tab. A locked keycard would not have accepted this.'
      });
      if ($('#write-save').checked) toast('Saved to your library as well.');
    } else {
      out.innerHTML = renderVerdict({ level: 'bad', title: 'The tag could not be written', body: res.error });
    }
  }

  function wireImportExport() {
    $('#export-btn').addEventListener('click', () => {
      const text = exportLibrary();
      const blob = new Blob([text], { type: 'application/json' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'nfc-tags-library.json';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 1000);
      toast('Library exported.');
    });
    $('#import-btn').addEventListener('click', () => $('#import-file').click());
    $('#import-file').addEventListener('change', async () => {
      const file = $('#import-file').files && $('#import-file').files[0];
      $('#import-file').value = '';
      if (!file) return;
      const text = await file.text();
      const res = importLibrary(text);
      if (res.error) { toast(res.error); return; }
      renderLibrary();
      toast(`Imported ${res.added} tag${res.added === 1 ? '' : 's'}.`);
    });
  }

  function wireTabs() {
    $$('.tab').forEach((b) => b.addEventListener('click', () => showTab(b.dataset.tab)));
  }

  function boot() {
    if ('serviceWorker' in navigator) {
      navigator.serviceWorker.register('./sw.js').catch(() => {});
    }
    setCapCard();
    renderLibrary();
    wireTabs();
    wireWriteForm();
    wireImportExport();
    bindLibraryClicks();
    $('#check-btn').addEventListener('click', beginCheck);
  }

  // Test hook: used by the headless tests. Not a public feature.
  window.NFCTags = {
    version: VERSION,
    capability,
    decodeRecord,
    buildWrite,
    encodeWifi,
    encodeVcard,
    decodeWifi,
    normaliseUrl,
    loadLibrary,
    addToLibrary,
    removeFromLibrary,
    exportLibrary,
    importLibrary,
    storedToWriteRecords,
    verdictForReading,
    verdictForReadingError,
    verdictForTimeout,
    describeRecords,
    showTab,
    renderLibrary,
    handleReading,
    LIB_KEY,
    simulateReading(reading) { handleReading(reading); },
    simulateError() { showCheckResult(renderVerdict(verdictForReadingError())); },
    simulateTimeout() { showCheckResult(renderVerdict(verdictForTimeout())); }
  };

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();
