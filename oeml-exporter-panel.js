// ==UserScript==
// @name         Outlook EML Exporter (Local Archive)
// @namespace    https://github.com/DlfnAntx/Outlook-EML-Exporter
// @version      1.0.0
// @description  Export Outlook Web mail to local .eml files with attachments using Graph or Outlook APIs.
// @author       DlfnAntx
// @license      MIT
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.live.com/*
// @grant        none
// @run-at       document-idle
// @noframes
// ==/UserScript==

(() => {
  'use strict';

  const APP_ID = 'oeml-exporter-panel';
  const CLASS_COLLAPSED = 'oeml-collapsed';
  const DRAG_THRESHOLD_PX = 4;
  const ABORT_MSG = 'Aborted by user.';

  const existing = document.getElementById(APP_ID);
    if (existing) existing.remove();


  // ====== CONFIG ======
  // Example: ['Inbox', 'Sent Items'] to limit scope.
  // Leave as null to export all top-level folders recursively.
  const ONLY_TOP_LEVEL_FOLDERS = null;
  const REQUEST_DELAY_MS = 120;
  const MAX_FILENAME_BYTES = 180;
  const MAX_DIRNAME_BYTES = 80;

  // ====== UTILITIES ======
  const utf8 = new TextEncoder();

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function utf8Len(s) {
    return utf8.encode(String(s ?? '')).length;
  }

  function truncateUtf8(s, maxBytes) {
    s = String(s ?? '');
    if (utf8Len(s) <= maxBytes) return s;
    let out = '';
    let used = 0;
    for (const ch of s) {
      const n = utf8Len(ch);
      if (used + n > maxBytes) break;
      out += ch;
      used += n;
    }
    return out;
  }

  const reservedWin = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;

  function sanitizePartRaw(s) {
    return String(s ?? '')
      .normalize('NFKD')
      .replace(/\p{M}+/gu, '') // strip combining marks
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
      .replace(/\s+/g, ' ')
      .trim()
      .replace(/[. ]+$/g, '');
  }

  function safePart(s, fallback = 'untitled', maxBytes = 100) {
    s = sanitizePartRaw(s);
    if (!s) s = fallback;
    s = truncateUtf8(s, maxBytes);
    s = s.replace(/[. ]+$/g, '');
    if (!s) s = fallback;
    if (reservedWin.test(s)) s = '_' + s;
    return s;
  }

  function sanitizeWholeFilename(name) {
    let s = String(name ?? '')
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
      .replace(/[ ]+$/g, '');

    if (s.toLowerCase().endsWith('.eml')) {
      const base = s.slice(0, -4).replace(/[. ]+$/g, '');
      s = (base || 'message') + '.eml';
    } else {
      s = s.replace(/[. ]+$/g, '');
    }

    if (!s) s = 'message.eml';
    if (reservedWin.test(s)) s = '_' + s;
    return s;
  }

  function truncateFilenamePreserveExt(name, maxBytes = MAX_FILENAME_BYTES) {
    name = sanitizeWholeFilename(name);
    const ext = name.toLowerCase().endsWith('.eml') ? '.eml' : '';
    let base = ext ? name.slice(0, -4) : name;

    if (utf8Len(name) <= maxBytes) return name;

    const budget = Math.max(1, maxBytes - utf8Len(ext));
    base = truncateUtf8(base, budget).replace(/[. ]+$/g, '');
    if (!base) base = 'message';
    return base + ext;
  }

  function shortId(id) {
    return String(id ?? '')
      .replace(/[^A-Za-z0-9_-]/g, '_')
      .slice(-24) || 'id';
  }

  function fnv1a64(str) {
    let h = 0xcbf29ce484222325n;
    for (const b of utf8.encode(String(str ?? ''))) {
      h ^= BigInt(b);
      h = (h * 0x100000001b3n) & 0xffffffffffffffffn;
    }
    return h.toString(16).padStart(16, '0');
  }

  function messageKey(id) {
    return `m-${fnv1a64(id)}`;
  }

  function stamp(dt) {
    if (!dt) return 'unknown-date';
    return String(dt)
      .replace(/:/g, '-')
      .replace(/\.\d+Z$/, 'Z')
      .replace('T', '_');
  }

  function b64urlDecode(seg) {
    const base64 = seg.replace(/-/g, '+').replace(/_/g, '/');
    const padded = base64 + '='.repeat((4 - (base64.length % 4 || 4)) % 4);
    return decodeURIComponent(Array.from(atob(padded), (c) =>
      '%' + c.charCodeAt(0).toString(16).padStart(2, '0')
    ).join(''));
  }

  function inspectJwt(token) {
    const parts = String(token).split('.');
    if (parts.length !== 3) return null;
    try {
      return {
        header: JSON.parse(b64urlDecode(parts[0])),
        payload: JSON.parse(b64urlDecode(parts[1]))
      };
    } catch {
      return null;
    }
  }

  function tokenSummary(token) {
    token = String(token || '');
    if (token.length <= 24) return token;
    return `${token.slice(0, 12)}…${token.slice(-12)} (len ${token.length})`;
  }

  // ====== TOKEN PARSING ======
  function extractBearerToken(input) {
    let s = String(input ?? '').trim();
    if (!s) throw new Error('Empty token input.');

    s = s
      .replace(/[\u2018\u2019]/g, "'")
      .replace(/[\u201C\u201D]/g, '"');

    // Copy as cURL / fetch / raw header / raw token
    const patterns = [
      /authorization["']?\s*:\s*["']\s*bearer\s+([^"'\\\r\n]+)["']/i,
      /authorization\s*:\s*bearer\s+([^\s"'\\\r\n]+)/i,
      /-H\s+['"]authorization:\s*Bearer\s+([^'"]+)['"]/i,
      /--header\s+['"]authorization:\s*Bearer\s+([^'"]+)['"]/i,
      /bearer\s+([^\s"'`,\\\r\n]+)/i
    ];

    for (const re of patterns) {
      const m = s.match(re);
      if (m?.[1]) return m[1].trim();
    }

    s = s.replace(/^["'`\s]+|["'`,;\s]+$/g, '');
    if (!s) throw new Error('Could not extract a bearer token.');
    if (/\s/.test(s)) {
      throw new Error('Could not extract a bearer token. Paste the raw token or full "Copy as fetch".');
    }
    return s;
  }

  // ====== API / BACKEND ======
  function authHeaders(token, backend, kind = 'json') {
    const headers = {
      Authorization: `Bearer ${token}`,
      Accept: kind === 'bytes' ? '*/*' : 'application/json'
    };

    if (backend?.kind === 'outlook-rest') {
      // Anchor mailbox improves affinity for Outlook REST
      const p = inspectJwt(token)?.payload || {};
      const anchor =
        p.preferred_username ||
        p.upn ||
        p.unique_name ||
        p.email;

      if (anchor) headers['X-AnchorMailbox'] = anchor;
      headers['X-PreferServerAffinity'] = 'true';
    }

    return headers;
  }

  function buildApiError(res, bodyText, token, backend, url) {
    let detail = bodyText.trim();
    try {
      const j = JSON.parse(bodyText);
      if (j?.error?.message) {
        detail = `${j.error.code || 'Error'}: ${j.error.message}`;
      }
    } catch {}

    const jwt = inspectJwt(token);
    const aud = String(jwt?.payload?.aud || '');
    const scp = String(jwt?.payload?.scp || '');
    const hints = [];

    if (backend?.kind === 'graph' && /^https:\/\/outlook\.office\.com/i.test(aud)) {
      hints.push('wrong audience: token is for outlook.office.com, not Graph');
    }

    if (backend?.kind === 'outlook-rest' &&
        /(graph\.microsoft\.com|00000003-0000-0000-c000-000000000000)/i.test(aud)) {
      hints.push('wrong audience: token is for Graph, not outlook.office.com');
    }

    if (backend?.kind === 'graph' &&
        /\/me\/mailFolders|\/me\/messages/i.test(url) &&
        !/\bMail\./i.test(scp)) {
      hints.push('this Graph token lacks Mail.* scopes');
    }

    if (/invalid audience/i.test(detail)) {
      hints.push('copy a token from a request to the same host as the API being called');
    }

    if (/CompactToken|malformed|parse|IDX14100|IDX12729|JWT|invalid token/i.test(detail)) {
      hints.push('token was likely truncated; paste full "Copy as fetch"');
    }

    if (/expired/i.test(detail)) {
      hints.push('token may be expired; reload Outlook and copy a fresh request');
    }

    let msg = `${res.status} ${res.statusText}`;
    if (detail) msg += ` — ${detail}`;
    if (hints.length) msg += `\nHints: ${hints.join(' | ')}`;
    return new Error(msg);
  }

  async function apiRequest(token, backend, url, kind = 'json', attempt = 0) {
    const res = await fetch(url, {
      headers: authHeaders(token, backend, kind)
    });

    // Handle rate limiting with Retry-After
    if (res.status === 429 && attempt < 8) {
      const retryAfter = Number(res.headers.get('Retry-After') || '5');
      log(`Rate-limited. Waiting ${retryAfter}s...`);
      await sleep(retryAfter * 1000);
      return apiRequest(token, backend, url, kind, attempt + 1);
    }

    if (!res.ok) {
      const bodyText = await res.text().catch(() => '');
      throw buildApiError(res, bodyText, token, backend, url);
    }

    if (kind === 'bytes') return new Uint8Array(await res.arrayBuffer());
    return res.json();
  }

  async function* paged(token, backend, url) {
    while (url) {
      const page = await apiRequest(token, backend, url, 'json');
      for (const item of page.value || []) yield item;
      url = page['@odata.nextLink'] || page['@odata.nextlink'] || null;
    }
  }

  async function resolveBackend(token) {
    const aud = String(inspectJwt(token)?.payload?.aud || '');

    // Use Graph if token audience indicates Graph
    if (/(graph\.microsoft\.com|00000003-0000-0000-c000-000000000000)/i.test(aud)) {
      return { kind: 'graph', base: 'https://graph.microsoft.com/v1.0', aud };
    }

    // Use Outlook REST if audience is outlook.office.com
    if (/^https:\/\/outlook\.office\.com\/?$/i.test(aud)) {
      const candidates = [
        'https://outlook.office.com/api/v2.0',
        'https://outlook.office.com/api/beta'
      ];

      const failures = [];
      for (const base of candidates) {
        const res = await fetch(`${base}/me/mailfolders?$top=1`, {
          headers: authHeaders(token, { kind: 'outlook-rest', base }, 'json')
        });
        const bodyText = await res.text().catch(() => '');

        if (res.ok) {
          return { kind: 'outlook-rest', base, aud };
        }

        failures.push(`${base} -> ${res.status} ${res.statusText}${bodyText ? ` — ${bodyText.slice(0, 220)}` : ''}`);
      }

      throw new Error(
        `Token audience is outlook.office.com, but no supported Outlook mail endpoint accepted it.\n` +
        failures.join('\n')
      );
    }

    throw new Error(`Unsupported token audience: ${aud || '(none)'}`);
  }

  function topFoldersUrl(backend) {
    return backend.kind === 'graph'
      ? `${backend.base}/me/mailFolders?$top=100`
      : `${backend.base}/me/mailfolders?$top=100`;
  }

  function childFoldersUrl(backend, folderId) {
    return backend.kind === 'graph'
      ? `${backend.base}/me/mailFolders/${encodeURIComponent(folderId)}/childFolders?$top=100`
      : `${backend.base}/me/mailfolders/${encodeURIComponent(folderId)}/childfolders?$top=100`;
  }

  function messagesUrl(backend, folderId) {
    return backend.kind === 'graph'
      ? `${backend.base}/me/mailFolders/${encodeURIComponent(folderId)}/messages?$top=100&$select=id,subject,receivedDateTime,sentDateTime,createdDateTime,from,hasAttachments`
      : `${backend.base}/me/mailfolders/${encodeURIComponent(folderId)}/messages?$top=100&$select=Id,Subject,ReceivedDateTime,SentDateTime,LastModifiedDateTime,From,HasAttachments`;
  }

  function mimeUrl(backend, messageId) {
    return `${backend.base}/me/messages/${encodeURIComponent(messageId)}/$value`;
  }

  function normalizeFolder(raw, backend) {
    return backend.kind === 'graph'
      ? { id: raw.id, displayName: raw.displayName }
      : { id: raw.Id, displayName: raw.DisplayName };
  }

  function normalizeMessage(raw, backend) {
    return backend.kind === 'graph'
      ? {
          id: raw.id,
          subject: raw.subject,
          receivedDateTime: raw.receivedDateTime,
          sentDateTime: raw.sentDateTime,
          createdDateTime: raw.createdDateTime,
          fromAddress: raw.from?.emailAddress?.address || 'unknown'
        }
      : {
          id: raw.Id,
          subject: raw.Subject,
          receivedDateTime: raw.ReceivedDateTime,
          sentDateTime: raw.SentDateTime,
          createdDateTime: raw.LastModifiedDateTime,
          fromAddress: raw.From?.EmailAddress?.Address || 'unknown'
        };
  }

  function summarizeMe(me, backend) {
    return backend.kind === 'graph'
      ? (me.userPrincipalName || me.mail || me.displayName || me.id || '(unknown)')
      : (me.EmailAddress || me.Name || me.DisplayName || me.Alias || me.Id || '(unknown)');
  }

  // ====== FILE SYSTEM ======
  async function getOrCreateDir(parent, name) {
    const dirName = safePart(name, 'folder', MAX_DIRNAME_BYTES);
    try {
      return await parent.getDirectoryHandle(dirName, { create: true });
    } catch {
      // Fallback to a hashed folder name if the OS rejects this name
      const fallback = `folder-${fnv1a64(dirName).slice(0, 8)}`;
      return parent.getDirectoryHandle(fallback, { create: true });
    }
  }

  async function scanExisting(dir) {
    const names = new Set();
    const keys = new Set();
    const keyRegex = /(?:__)?(m-[0-9a-f]{16}|[A-Za-z0-9_-]{1,24})\.eml$/i;

    try {
      for await (const [name, handle] of dir.entries()) {
        if (handle.kind !== 'file') continue;
        names.add(name);

        const m = name.match(keyRegex);
        if (m?.[1]) keys.add(m[1]);
      }
    } catch (err) {
      // If directory listing fails, resume logic will be less effective but still safe
      console.warn('[oeml] scanExisting failed:', err);
    }

    return { names, keys };
  }

  function buildMessageFilename(msg) {
    const date = safePart(
      stamp(msg.receivedDateTime || msg.sentDateTime || msg.createdDateTime),
      'unknown-date',
      32
    );
    const sender = safePart(msg.fromAddress || 'unknown', 'unknown', 48);
    const subject = safePart(msg.subject || '(no subject)', 'no-subject', 96);

    // Use a stable per-message suffix to avoid collisions and enable resume
    const stableKey = messageKey(msg.id);
    const legacyKey = shortId(msg.id);

    const suffix = `__${stableKey}.eml`;
    let prefix = `${date}__${sender}__${subject}`;
    const maxPrefixBytes = Math.max(1, MAX_FILENAME_BYTES - utf8Len(suffix));
    prefix = truncateUtf8(prefix, maxPrefixBytes).replace(/[. ]+$/g, '');
    if (!prefix) prefix = date || 'message';

    let filename = sanitizeWholeFilename(`${prefix}${suffix}`);
    filename = truncateFilenamePreserveExt(filename, MAX_FILENAME_BYTES);

    return { filename, stableKey, legacyKey, date };
  }

  async function writeBytesSafe(dir, preferredName, bytes, stableKey, date) {
    const candidates = [
      preferredName,
      `${date}__${stableKey}.eml`,
      `${stableKey}.eml`
    ]
      .map((n) => truncateFilenamePreserveExt(n, MAX_FILENAME_BYTES))
      .filter((v, i, a) => v && a.indexOf(v) === i);

    let lastErr = null;

    for (const name of candidates) {
      try {
        const fh = await dir.getFileHandle(name, { create: true });
        const w = await fh.createWritable();
        await w.write(bytes);
        await w.close();
        return name;
      } catch (e) {
        lastErr = e;
        const msg = String(e?.message || e);
        if (!/Name is not allowed/i.test(msg)) throw e;
      }
    }

    throw lastErr || new Error('Failed to write file.');
  }

  // ====== UI ======
  injectStyle();
  const ui = buildUI();
  const { panel } = ui;

  let isCollapsed = true;
  setCollapsed(true);
  setupDrag(panel, ui.header, toggleCollapsed);

  // Ensure panel stays visible if the window is resized while the panel is on screen
  window.addEventListener('resize', () => {
    const rect = panel.getBoundingClientRect();
    const maxLeft = Math.max(0, window.innerWidth - rect.width);
    const maxTop  = Math.max(0, window.innerHeight - rect.height);
    let nx = rect.left;
    let ny = rect.top;

    if (nx > maxLeft) nx = maxLeft;
    if (ny > maxTop)  ny = maxTop;

    panel.style.left = `${nx}px`;
    panel.style.top  = `${ny}px`;
  });

  // ====== STATE ======
  const state = {
    rootDir: null,
    exporting: false,
    abort: false
  };

  // ====== ACTIONS ======
  ui.tokenInput.addEventListener('input', updateButtons);
  ui.pickBtn.addEventListener('click', pickFolder);
  ui.testBtn.addEventListener('click', testToken);
  ui.startBtn.addEventListener('click', startExport);
  ui.stopBtn.addEventListener('click', stopExport);

  log('Ready. Click the header to expand.');
  updateButtons();

  // ====== UI HELPERS ======
  function log(...args) {
    const line = args.map((x) => {
      if (typeof x === 'string') return x;
      try { return JSON.stringify(x); } catch { return String(x); }
    }).join(' ');

    console.log('[oeml]', ...args);
    ui.logBox.textContent += line + '\n';
    ui.logBox.scrollTop = ui.logBox.scrollHeight;
  }

  function setStatus(text) {
    ui.status.textContent = text;
  }

  function setRunState(stateText) {
    ui.stateBadge.textContent = stateText;
    ui.stateBadge.dataset.state = stateText;
  }

  function setCollapsed(value) {
    isCollapsed = value;
    ui.panel.classList.toggle(CLASS_COLLAPSED, isCollapsed);
    ui.toggle.textContent = isCollapsed ? '▸' : '▾';
  }

  function toggleCollapsed() {
    setCollapsed(!isCollapsed);
  }

  function updateButtons() {
    const hasToken = !!ui.tokenInput.value.trim();
    const hasFolder = !!state.rootDir;

    ui.pickBtn.disabled = state.exporting;
    ui.testBtn.disabled = state.exporting || !hasToken;
    ui.startBtn.disabled = state.exporting || !hasToken || !hasFolder;
    ui.stopBtn.disabled = !state.exporting;
  }

  // ====== DRAG ======
  function setupDrag(panel, header, onToggle) {
    let active = false;
    let moved = false;
    let startX = 0;
    let startY = 0;
    let startLeft = 0;
    let startTop = 0;

    header.addEventListener('pointerdown', (e) => {
      if (e.button !== 0) return;
      active = true;
      moved = false;

      const rect = panel.getBoundingClientRect();
      startX = e.clientX;
      startY = e.clientY;
      startLeft = rect.left;
      startTop = rect.top;

      header.setPointerCapture(e.pointerId);
      e.preventDefault();
    });

    header.addEventListener('pointermove', (e) => {
      if (!active) return;

      const dx = e.clientX - startX;
      const dy = e.clientY - startY;

      if (!moved && Math.hypot(dx, dy) < DRAG_THRESHOLD_PX) return;
      moved = true;

      // Clamp to viewport
      const rect = panel.getBoundingClientRect();
      const panelW = rect.width;
      const panelH = rect.height;

      let newLeft = startLeft + dx;
      let newTop  = startTop + dy;

      const maxLeft = Math.max(0, window.innerWidth - panelW);
      const maxTop  = Math.max(0, window.innerHeight - panelH);

      newLeft = Math.max(0, Math.min(newLeft, maxLeft));
      newTop  = Math.max(0, Math.min(newTop, maxTop));
    
      panel.style.left = `${newLeft}px`;
      panel.style.top  = `${newTop}px`;
      panel.style.right = 'auto';
      panel.style.bottom = 'auto';
    });

    function endDrag(e) {
      if (!active) return;
      active = false;
      header.releasePointerCapture(e.pointerId);
      if (!moved) onToggle();
    }

    header.addEventListener('pointerup', endDrag);
    header.addEventListener('pointercancel', endDrag);
  }

  // ====== ACTIONS IMPLEMENTATION ======
  async function pickFolder() {
    if (!window.showDirectoryPicker) {
      log('ERROR: showDirectoryPicker not available. Use a Chromium-based browser.');
      return;
    }

    try {
      state.rootDir = await window.showDirectoryPicker({ mode: 'readwrite' });
      setStatus(`Folder selected: ${state.rootDir.name || '(unnamed)'}`);
    } catch (e) {
      if (e?.name !== 'AbortError') log('ERROR:', e?.message || String(e));
    }

    updateButtons();
  }

  async function testToken() {
    ui.logBox.textContent = '';

    try {
      const token = extractBearerToken(ui.tokenInput.value);
      ui.tokenInput.value = token;
      updateButtons();

      log(`Token extracted: ${tokenSummary(token)}`);

      const jwt = inspectJwt(token);
      if (jwt) {
        const exp = jwt.payload.exp ? new Date(jwt.payload.exp * 1000).toISOString() : '(none)';
        log(`JWT aud: ${jwt.payload.aud ?? '(none)'}`);
        log(`JWT scp: ${jwt.payload.scp ?? '(none)'}`);
        log(`JWT exp: ${exp}`);
        if (jwt.payload.preferred_username || jwt.payload.upn) {
          log(`JWT user: ${jwt.payload.preferred_username || jwt.payload.upn}`);
        }
      } else {
        log('Token is not a decodable JWT in-browser. That is not necessarily a problem.');
      }

      const backend = await resolveBackend(token);
      log(`Backend: ${backend.kind} @ ${backend.base}`);

      const me = await apiRequest(
        token,
        backend,
        backend.kind === 'graph'
          ? `${backend.base}/me?$select=id,displayName,userPrincipalName,mail`
          : `${backend.base}/me`,
        'json'
      );
      log(`Identity probe OK: ${summarizeMe(me, backend)}`);

      const folderProbe = await apiRequest(
        token,
        backend,
        backend.kind === 'graph'
          ? `${backend.base}/me/mailFolders?$top=1&$select=id,displayName`
          : `${backend.base}/me/mailfolders?$top=1&$select=Id,DisplayName`,
        'json'
      );
      log(`Mail probe OK: ${folderProbe.value?.length ?? 0} item(s) returned in probe`);
      log('Token test PASSED.');
    } catch (err) {
      log('ERROR:', err?.message || String(err));
    }
  }

  async function startExport() {
    if (!state.rootDir) {
      log('ERROR: pick a folder first.');
      return;
    }

    let token;
    try {
      token = extractBearerToken(ui.tokenInput.value);
      ui.tokenInput.value = token;
      updateButtons();
    } catch (err) {
      log('ERROR:', err?.message || String(err));
      return;
    }

    state.exporting = true;
    state.abort = false;
    updateButtons();
    ui.logBox.textContent = '';

    setRunState('running');
    setStatus('Export running...');

    try {
      const backend = await resolveBackend(token);
      log(`Starting export with backend ${backend.kind} @ ${backend.base}`);
      log(`Token: ${tokenSummary(token)}`);
      await runExport(token, backend, state.rootDir);
      log('DONE');
      setRunState('done');
      setStatus('Export complete.');
    } catch (err) {
      if (err?.message === ABORT_MSG) {
        log('Stopped by user.');
        setRunState('stopped');
        setStatus('Stopped.');
      } else {
        log('ERROR:', err?.message || String(err));
        setRunState('error');
        setStatus('Error. See log for details.');
      }
    } finally {
      state.exporting = false;
      updateButtons();
    }
  }

  function stopExport() {
    if (state.exporting) {
      state.abort = true;
      log('Stop requested (will finish current request).');
    }
  }

  // ====== EXPORT ======
  async function runExport(token, backend, rootDir) {
    const assertNotAborted = () => {
      if (state.abort) throw new Error(ABORT_MSG);
    };

    async function exportFolder(rawFolder, parentDir, depth = 0) {
      assertNotAborted();

      const folder = normalizeFolder(rawFolder, backend);
      const folderName = safePart(folder.displayName || 'folder', 'folder', MAX_DIRNAME_BYTES);
      const dir = await getOrCreateDir(parentDir, folderName);
      const existing = await scanExisting(dir);

      const pad = '  '.repeat(depth);
      log(`${pad}Folder: ${folderName}`);
      setStatus(`Exporting: ${folderName}`);

      let seen = 0;
      let saved = 0;
      let skipped = 0;

      for await (const rawMsg of paged(token, backend, messagesUrl(backend, folder.id))) {
        assertNotAborted();
        seen += 1;

        const msg = normalizeMessage(rawMsg, backend);
        const meta = buildMessageFilename(msg);

        if (
          existing.keys.has(meta.stableKey) ||
          existing.keys.has(meta.legacyKey) ||
          existing.names.has(meta.filename)
        ) {
          skipped += 1;
          if (skipped % 100 === 0) log(`${pad}  skipped existing: ${skipped}`);
          continue;
        }

        const bytes = await apiRequest(token, backend, mimeUrl(backend, msg.id), 'bytes');
        const actualName = await writeBytesSafe(dir, meta.filename, bytes, meta.stableKey, meta.date);

        existing.names.add(actualName);
        existing.keys.add(meta.stableKey);
        existing.keys.add(meta.legacyKey);

        saved += 1;
        if (saved % 20 === 0) log(`${pad}  saved: ${saved}`);
        await sleep(REQUEST_DELAY_MS);
      }

      log(`${pad}  messages complete: seen=${seen}, saved=${saved}, skipped=${skipped}`);

      for await (const childRaw of paged(token, backend, childFoldersUrl(backend, folder.id))) {
        await exportFolder(childRaw, dir, depth + 1);
      }
    }

    for await (const rawFolder of paged(token, backend, topFoldersUrl(backend))) {
      const folder = normalizeFolder(rawFolder, backend);
      if (ONLY_TOP_LEVEL_FOLDERS && !ONLY_TOP_LEVEL_FOLDERS.includes(folder.displayName)) {
        continue;
      }
      await exportFolder(rawFolder, rootDir, 0);
    }
  }

  // ====== UI BUILD ======
  function injectStyle() {
    const css = `
#${APP_ID} {
  position: fixed;
  top: 12px;
  right: 12px;
  width: 620px;
  max-width: calc(100vw - 24px);
  color: #eaeaea;
  background: rgba(0,0,0,0.90);
  font: 12px/1.4 system-ui, -apple-system, Segoe UI, Roboto, sans-serif;
  border-radius: 10px;
  box-shadow: 0 6px 24px rgba(0,0,0,.35);
  z-index: 2147483647;
}
#${APP_ID}.${CLASS_COLLAPSED} {
  width: 280px;
}
#${APP_ID} .oeml-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  padding: 8px 10px;
  cursor: move;
  user-select: none;
  touch-action: none;
  background: rgba(255,255,255,0.06);
  border-top-left-radius: 10px;
  border-top-right-radius: 10px;
}
#${APP_ID} .oeml-title {
  font-weight: 700;
  font-size: 13px;
}
#${APP_ID} .oeml-header-right {
  display: flex;
  align-items: center;
  gap: 8px;
}
#${APP_ID} .oeml-badge {
  padding: 2px 6px;
  border-radius: 6px;
  background: #333;
  font-size: 10px;
  text-transform: uppercase;
  letter-spacing: .06em;
}
#${APP_ID} .oeml-badge[data-state="running"] { background: #107c10; }
#${APP_ID} .oeml-badge[data-state="done"] { background: #0f6cbd; }
#${APP_ID} .oeml-badge[data-state="stopped"] { background: #a4262c; }
#${APP_ID} .oeml-badge[data-state="error"] { background: #a4262c; }
#${APP_ID} .oeml-toggle { font-size: 14px; }
#${APP_ID} .oeml-body {
  display: flex;
  flex-direction: column;
  gap: 8px;
  padding: 10px;
}
#${APP_ID}.${CLASS_COLLAPSED} .oeml-body {
  display: none;
}
#${APP_ID} .oeml-help {
  font-size: 11px;
  color: #c9c9c9;
}
#${APP_ID} .oeml-help ol {
  margin: 6px 0 8px 18px;
  padding: 0;
}
#${APP_ID} .oeml-help li {
  margin: 2px 0;
}
#${APP_ID} .oeml-help code {
  background: rgba(255,255,255,.08);
  padding: 1px 4px;
  border-radius: 4px;
}
#${APP_ID} .oeml-note {
  color: #a7c7ff;
}
#${APP_ID} .oeml-input {
  width: 100%;
  box-sizing: border-box;
  background: #111;
  color: #fff;
  border: 1px solid #444;
  border-radius: 6px;
  padding: 6px;
  font: 11px/1.35 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  user-select: text;
}
#${APP_ID} .oeml-buttons {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}
#${APP_ID} .oeml-btn {
  padding: 8px 10px;
  border: 0;
  border-radius: 6px;
  color: #fff;
  cursor: pointer;
  font-weight: 600;
}
#${APP_ID} .oeml-btn:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
#${APP_ID} .oeml-status {
  color: #d0d0ff;
}
#${APP_ID} .oeml-log {
  padding: 8px;
  max-height: 45vh;
  overflow: auto;
  background: rgba(0,0,0,0.55);
  border-radius: 6px;
  white-space: pre-wrap;
  font: 11px/1.35 ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  user-select: text;
  cursor: text;
}
`;
    const style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);
  }

  function buildUI() {
    const panel = document.createElement('div');
    panel.id = APP_ID;
    panel.className = CLASS_COLLAPSED;

    const header = document.createElement('div');
    header.className = 'oeml-header';

    const title = document.createElement('div');
    title.className = 'oeml-title';
    title.textContent = 'Outlook → EML Export';

    const headerRight = document.createElement('div');
    headerRight.className = 'oeml-header-right';

    const stateBadge = document.createElement('span');
    stateBadge.className = 'oeml-badge';
    stateBadge.textContent = 'idle';
    stateBadge.dataset.state = 'idle';

    const toggle = document.createElement('span');
    toggle.className = 'oeml-toggle';
    toggle.textContent = '▸';

    headerRight.append(stateBadge, toggle);
    header.append(title, headerRight);

    const body = document.createElement('div');
    body.className = 'oeml-body';

    const help = document.createElement('div');
    help.className = 'oeml-help';
    help.innerHTML = `
      <div><strong>Usage</strong></div>
      <ol>
        <li>Click <b>Pick export folder</b> (Chromium required).</li>
        <li>In Outlook Web: <b>F12 → Network</b>, filter <code>mailfolders</code> or <code>messages</code>, pick a <b>200 OK</b> request to <code>outlook.office.com</code> or <code>graph.microsoft.com</code>, then <b>Copy as fetch</b>.</li>
        <li>Paste into the token box, click <b>Test token</b>.</li>
        <li>Click <b>Start export</b> and keep this tab open. Re‑runs skip existing files.</li>
      </ol>
      <div class="oeml-note">Exports .eml with attachments embedded. Drag the header to move; click it to collapse/expand. Tokens are sensitive—don’t share them.</div>
    `;

    const tokenInput = document.createElement('textarea');
    tokenInput.className = 'oeml-input';
    tokenInput.rows = 5;
    tokenInput.placeholder = 'Paste token / Authorization header / Copy as fetch / Copy as cURL here';
    tokenInput.spellcheck = false;

    const buttons = document.createElement('div');
    buttons.className = 'oeml-buttons';

    const pickBtn = makeButton('Pick export folder', '#0f6cbd');
    const testBtn = makeButton('Test token', '#5c2d91');
    const startBtn = makeButton('Start export', '#107c10');
    const stopBtn = makeButton('Stop', '#a4262c');

    buttons.append(pickBtn, testBtn, startBtn, stopBtn);

    const status = document.createElement('div');
    status.className = 'oeml-status';
    status.textContent = 'No folder selected.';

    const logBox = document.createElement('pre');
    logBox.className = 'oeml-log';
    logBox.tabIndex = 0;

    body.append(help, tokenInput, buttons, status, logBox);
    panel.append(header, body);
    document.body.appendChild(panel);

    return {
      panel,
      header,
      toggle,
      stateBadge,
      tokenInput,
      pickBtn,
      testBtn,
      startBtn,
      stopBtn,
      status,
      logBox
    };
  }

  function makeButton(label, bg) {
    const btn = document.createElement('button');
    btn.className = 'oeml-btn';
    btn.textContent = label;
    btn.style.background = bg;
    return btn;
  }

})();
