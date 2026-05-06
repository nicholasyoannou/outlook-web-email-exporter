// ==UserScript==
// @name         Outlook Email Exporter (REST, complete)
// @namespace    https://outlook.cloud.microsoft/
// @version      3.6
// @description  Exports every email in a folder as real .eml using the Outlook REST API: /api/v2.0/me/mailFolders/{folder}/messages for listing (paged via $top/$skip, with accurate $count) and /api/v2.0/me/messages/{id}/$value for download. Single host, single auth, includes attachments. Bundled into ZIPs. Retries on 429/5xx with shared cooldown so the server's mailbox-concurrency cap doesn't drop messages. Atomic zip-flush so concurrent workers can't trigger duplicate downloads with the same filename.
// @match        https://outlook.cloud.microsoft/*
// @match        https://outlook.office.com/*
// @match        https://outlook.office365.com/*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
// @run-at       document-start
// ==/UserScript==

(function () {
    'use strict';

    const ORIGIN = location.origin;

    // -----------------------------------------------------------------
    // Auth capture (bearer + anchor mailbox) from real Outlook traffic
    // -----------------------------------------------------------------
    let bearerAuth = null;
    let anchorMailbox = null;
    let clientVersion = null;

    function readHeader(headers, name) {
        if (!headers) return null;
        const lower = name.toLowerCase();
        if (headers instanceof Headers) return headers.get(name) || headers.get(lower);
        if (Array.isArray(headers)) {
            const e = headers.find(([k]) => k.toLowerCase() === lower);
            return e ? e[1] : null;
        }
        for (const [k, v] of Object.entries(headers)) {
            if (k.toLowerCase() === lower) return v;
        }
        return null;
    }

    function captureFrom(url, headers) {
        if (!url || !headers) return;
        if (url.includes('/owa/service.svc') ||
            url.includes('/api/beta/me/') ||
            url.includes('/api/v2.0/me/') ||
            url.includes('/api/v1.0/me/')) {
            const auth = readHeader(headers, 'authorization');
            if (auth && /^Bearer\s+/i.test(auth)) bearerAuth = auth;
            const a = readHeader(headers, 'x-anchormailbox'); if (a) anchorMailbox = a;
            const v = readHeader(headers, 'x-client-version'); if (v) clientVersion = v;
        }
    }

    const origFetch = window.fetch.bind(window);
    window.fetch = function (input, init) {
        try {
            const url = typeof input === 'string' ? input : (input && input.url);
            const headers = (init && init.headers) || (input && input.headers);
            captureFrom(url, headers);
        } catch (_) {}
        return origFetch.apply(this, arguments);
    };

    const xhrOpen = XMLHttpRequest.prototype.open;
    const xhrSet = XMLHttpRequest.prototype.setRequestHeader;
    const xhrSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function (m, u) { this.__url = u; this.__headers = {}; return xhrOpen.apply(this, arguments); };
    XMLHttpRequest.prototype.setRequestHeader = function (n, v) {
        try { if (this.__headers) this.__headers[n] = v; } catch (_) {}
        return xhrSet.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function () {
        try { if (this.__url && this.__headers) captureFrom(this.__url, this.__headers); } catch (_) {}
        return xhrSend.apply(this, arguments);
    };

    function ready() { return !!(bearerAuth && anchorMailbox); }

    function restHeaders(extra) {
        const h = {
            'Accept': 'application/json',
            'Authorization': bearerAuth,
            'x-anchormailbox': anchorMailbox,
            'prefer': 'IdType="ImmutableId"',
        };
        if (clientVersion) h['x-client-version'] = clientVersion;
        if (extra) Object.assign(h, extra);
        return h;
    }

    // -----------------------------------------------------------------
    // Throttling: shared cooldown so when one worker hits 429,
    // all workers pause until the throttle clears.
    // -----------------------------------------------------------------
    let cooldownUntil = 0;
    let throttleHits = 0;

    async function waitForCooldown() {
        const now = Date.now();
        if (cooldownUntil > now) {
            await sleep(cooldownUntil - now);
        }
    }

    function bumpCooldown(retryAfterHeader, attempt) {
        let ms;
        if (retryAfterHeader) {
            const seconds = parseFloat(retryAfterHeader);
            if (!isNaN(seconds)) ms = Math.max(1000, Math.ceil(seconds * 1000));
        }
        if (!ms) ms = Math.min(30000, 1000 * Math.pow(2, attempt)); // 1s,2s,4s,8s,16s,30s
        const target = Date.now() + ms;
        if (target > cooldownUntil) cooldownUntil = target;
        throttleHits++;
        return ms;
    }

    // -----------------------------------------------------------------
    // Robust fetch with retry on 429 / 5xx / network errors.
    // Returns the Response on success; throws on permanent failure.
    // -----------------------------------------------------------------
    async function fetchWithRetry(url, init, opts = {}) {
        const maxAttempts = opts.maxAttempts ?? 6;
        let lastErr = null;
        for (let attempt = 0; attempt < maxAttempts; attempt++) {
            await waitForCooldown();
            if (cancelled) throw new Error('cancelled');
            try {
                const r = await origFetch(url, init);
                if (r.ok) return r;
                if (r.status === 429) {
                    bumpCooldown(r.headers.get('retry-after') || r.headers.get('Retry-After'), attempt);
                    continue;
                }
                if (r.status >= 500 && r.status < 600) {
                    bumpCooldown(null, attempt);
                    continue;
                }
                if (r.status === 401) {
                    // bearer expired mid-run; we can't refresh from here
                    const t = await r.text().catch(() => '');
                    throw new Error(`HTTP 401 (auth expired — refresh Outlook tab): ${t.slice(0, 120)}`);
                }
                // 4xx other than 429/401 → permanent for this id
                const t = await r.text().catch(() => '');
                throw new Error(`HTTP ${r.status}: ${t.slice(0, 200)}`);
            } catch (e) {
                lastErr = e;
                // Network error / abort: backoff and retry
                if (e.message === 'cancelled' || /HTTP 4\d\d/.test(e.message)) throw e;
                bumpCooldown(null, attempt);
            }
        }
        throw lastErr || new Error('exhausted retries');
    }

    // -----------------------------------------------------------------
    // List messages in a folder via REST (every message, paginated)
    // -----------------------------------------------------------------
    async function listMessageIds(folder, statusEl) {
        const first = `${ORIGIN}/api/v2.0/me/mailFolders/${encodeURIComponent(folder)}/messages?$top=200&$count=true&$select=Id&$orderby=ReceivedDateTime%20desc`;
        const ids = [];
        let total = null;
        let url = first;
        let pageNum = 0;
        while (url) {
            pageNum++;
            const r = await fetchWithRetry(url, { credentials: 'include', headers: restHeaders() });
            const j = await r.json();
            if (total == null && typeof j['@odata.count'] === 'number') total = j['@odata.count'];
            for (const m of (j.value || [])) {
                const id = m.Id || m.id;
                if (id) ids.push(id);
            }
            statusEl.textContent = `Listing ${folder}: ${ids.length}${total ? ' / ' + total : ''} message ids...`;
            url = j['@odata.nextLink'] || null;
        }
        return { ids, total: total ?? ids.length };
    }

    // -----------------------------------------------------------------
    // Fetch a single .eml (with retry)
    // -----------------------------------------------------------------
    async function fetchEml(id) {
        const url = `${ORIGIN}/api/v2.0/me/messages/${encodeURIComponent(id)}/$value`;
        const r = await fetchWithRetry(url, {
            credentials: 'include',
            headers: { 'Authorization': bearerAuth, 'x-anchormailbox': anchorMailbox }
        });
        const buf = await r.arrayBuffer();
        return new Uint8Array(buf);
    }

    // -----------------------------------------------------------------
    // MIME header parsing for filenames
    // -----------------------------------------------------------------
    function parseMimeHeaders(bytes) {
        const cap = Math.min(bytes.length, 64 * 1024);
        let end = -1;
        for (let i = 0; i < cap - 3; i++) {
            if (bytes[i] === 0x0d && bytes[i + 1] === 0x0a && bytes[i + 2] === 0x0d && bytes[i + 3] === 0x0a) { end = i; break; }
            if (bytes[i] === 0x0a && bytes[i + 1] === 0x0a) { end = i; break; }
        }
        if (end < 0) end = cap;
        let s = '';
        for (let i = 0; i < end; i++) s += String.fromCharCode(bytes[i]);
        s = s.replace(/\r?\n[ \t]+/g, ' ');
        const headers = {};
        for (const line of s.split(/\r?\n/)) {
            const m = line.match(/^([^:]+):\s*(.*)$/);
            if (m) {
                const name = m[1].trim().toLowerCase();
                if (!(name in headers)) headers[name] = m[2].trim();
            }
        }
        return headers;
    }

    function decodeRfc2047(s) {
        if (!s) return '';
        return s.replace(/=\?([^?]+)\?([BbQq])\?([^?]*)\?=/g, (_, charset, enc, payload) => {
            try {
                let bytes;
                if (enc.toUpperCase() === 'B') {
                    const bin = atob(payload);
                    bytes = new Uint8Array(bin.length);
                    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
                } else {
                    const arr = [];
                    for (let i = 0; i < payload.length; i++) {
                        const ch = payload[i];
                        if (ch === '_') arr.push(0x20);
                        else if (ch === '=' && /[0-9A-Fa-f]{2}/.test(payload.slice(i + 1, i + 3))) {
                            arr.push(parseInt(payload.slice(i + 1, i + 3), 16));
                            i += 2;
                        } else arr.push(ch.charCodeAt(0));
                    }
                    bytes = new Uint8Array(arr);
                }
                try { return new TextDecoder(charset).decode(bytes); }
                catch { return new TextDecoder('utf-8').decode(bytes); }
            } catch { return _; }
        });
    }

    function extractEmail(addrLine) {
        if (!addrLine) return '';
        const m = addrLine.match(/<([^>]+)>/);
        if (m) return m[1].trim();
        return addrLine.replace(/[",]/g, '').trim();
    }

    function safe(s, max = 60) {
        return (s || '').replace(/[\\\/:*?"<>|\x00-\x1f]/g, '_').replace(/\s+/g, ' ').trim().slice(0, max) || 'untitled';
    }

    function buildFilenameFromMime(bytes, itemId) {
        const h = parseMimeHeaders(bytes);
        const subject = decodeRfc2047(h['subject'] || '') || 'No subject';
        const from = extractEmail(decodeRfc2047(h['from'] || h['sender'] || '')) || 'unknown';
        let date = 'nodate';
        if (h['date']) {
            const d = new Date(h['date']);
            if (!isNaN(d.getTime())) date = d.toISOString().slice(0, 10);
        }
        const tail = String(itemId).replace(/[^a-zA-Z0-9]/g, '').slice(-10);
        return `${date}__${safe(from, 40)}__${safe(subject, 60)}__${tail}.eml`;
    }

    // -----------------------------------------------------------------
    // Export with bounded concurrency
    // -----------------------------------------------------------------
    let cancelled = false, paused = false;
    const sleep = ms => new Promise(r => setTimeout(r, ms));
    async function gate() { while (paused && !cancelled) await sleep(200); }

    async function runWithConcurrency(items, concurrency, worker) {
        let i = 0;
        const runners = Array.from({ length: concurrency }, async () => {
            while (true) {
                if (cancelled) return;
                await gate();
                const idx = i++;
                if (idx >= items.length) return;
                await worker(items[idx], idx);
            }
        });
        await Promise.all(runners);
    }

    async function exportAll(folderName, batchSize, concurrency, statusEl) {
        const startedAt = Date.now();

        // 1. List every message id in the folder.
        let ids = [], total = 0;
        try {
            const r = await listMessageIds(folderName, statusEl);
            ids = r.ids; total = r.total;
        } catch (e) {
            statusEl.textContent = 'Listing failed: ' + e.message;
            return;
        }
        if (ids.length === 0) {
            statusEl.textContent = `No messages found in ${folderName}.`;
            return;
        }
        statusEl.textContent = `Found ${ids.length} of ${total} messages. Downloading...`;

        // 2. Fetch each message and roll into ZIPs of `batchSize`.
        let zip = new JSZip();
        let zipIdx = 1, inZip = 0, done = 0, errors = 0;
        // Atomic flush: snapshot zip + idx and rotate state synchronously
        // (no await before the rotation) so that concurrent workers can't all
        // see "inZip >= batchSize" and trigger duplicate flushes of the same
        // zipIdx. Subsequent emails go straight into the new zip.
        const flushZipIfFull = async () => {
            if (inZip < batchSize) return;
            const myZip = zip;
            const myIdx = zipIdx;
            zip = new JSZip();
            inZip = 0;
            zipIdx++;
            statusEl.textContent = `Building zip #${myIdx}... (${batchSize} emails)`;
            await flushZip(myZip, myIdx);
        };

        await runWithConcurrency(ids, concurrency, async (id) => {
            try {
                const bytes = await fetchEml(id);
                if (!bytes || bytes.length === 0) throw new Error('empty');
                const filename = buildFilenameFromMime(bytes, id);
                zip.file(filename, bytes);
                done++; inZip++;
            } catch (e) {
                if (e.message === 'cancelled') return;
                errors++;
                if (errors === 1) statusEl.textContent = 'First failure: ' + e.message;
                console.warn('[exporter] failed', id, e.message);
            }
            if ((done + errors) % 25 === 0 || done === ids.length) {
                const elapsed = (Date.now() - startedAt) / 1000;
                const rate = done / Math.max(elapsed, 0.01);
                const eta = rate > 0 ? Math.round((ids.length - done - errors) / rate) : 0;
                const inCooldown = cooldownUntil > Date.now();
                const throttleNote = throttleHits > 0
                    ? ` • throttled ${throttleHits}x${inCooldown ? ' (cooling down)' : ''}`
                    : '';
                statusEl.textContent =
                    `${done}/${ids.length} (errors ${errors}) • zip #${zipIdx} (${inZip}/${batchSize}) • ${rate.toFixed(1)}/s • eta ${eta}s${throttleNote}`;
            }
            await flushZipIfFull();
        });

        if (inZip > 0) {
            statusEl.textContent = `Building final zip #${zipIdx}... (${inZip} emails)`;
            await flushZip(zip, zipIdx);
        }

        const secs = Math.round((Date.now() - startedAt) / 1000);
        statusEl.textContent = `${cancelled ? 'Stopped' : 'Done'} — ${done} of ${ids.length} emails, ${errors} errors, ${secs}s.`;
    }

    async function flushZip(zip, idx) {
        const blob = await zip.generateAsync({
            type: 'blob',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 }
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `outlook-export-${String(idx).padStart(3, '0')}.zip`;
        document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 60000);
    }

    // -----------------------------------------------------------------
    // UI
    // -----------------------------------------------------------------
    function el(tag, props, ...children) {
        const node = document.createElement(tag);
        if (props) for (const [k, v] of Object.entries(props)) {
            if (k === 'style' && typeof v === 'object') Object.assign(node.style, v);
            else if (k.startsWith('on') && typeof v === 'function') node.addEventListener(k.slice(2).toLowerCase(), v);
            else if (k in node) { try { node[k] = v; } catch { node.setAttribute(k, v); } }
            else node.setAttribute(k, v);
        }
        for (const c of children) {
            if (c == null || c === false) continue;
            node.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
        }
        return node;
    }

    function buildUI() {
        if (document.getElementById('oee-panel')) return;

        const styleEl = document.createElement('style');
        styleEl.textContent = `
#oee-panel { position: fixed; top: 70px; right: 20px; z-index: 2147483647; width: 340px;
  background: #fff; color: #222; font: 13px/1.4 system-ui, sans-serif;
  border: 1px solid #c7c7c7; border-radius: 10px; box-shadow: 0 6px 24px rgba(0,0,0,.18);
  padding: 14px; box-sizing: border-box; }
#oee-panel * { box-sizing: border-box; }
#oee-panel h3 { margin: 0 0 10px; font-size: 15px; padding-right: 22px; }
#oee-panel label { display: block; margin: 8px 0 3px; font-size: 12px; color: #555; }
#oee-panel select, #oee-panel input { width: 100%; padding: 5px;
  border: 1px solid #bbb; border-radius: 4px; font: inherit; }
#oee-panel .row { display: flex; gap: 8px; }
#oee-panel .row > * { flex: 1; }
#oee-panel button { padding: 7px 10px; cursor: pointer; border: 1px solid #bbb;
  border-radius: 5px; background: #f6f6f6; font: inherit; }
#oee-panel button:hover { background: #ececec; }
#oee-panel .btns { display: flex; gap: 6px; margin-top: 10px; }
#oee-status { margin-top: 10px; padding: 8px; background: #f3f6fa; border: 1px solid #e0e6ee;
  border-radius: 5px; min-height: 36px; font-size: 12px; word-break: break-word; }
#oee-ready { font-size: 11px; color: #888; margin-top: 6px; }
#oee-ready .ok { color: #2a7; font-weight: 600; }
#oee-ready .miss { color: #b00; }
#oee-close { position: absolute; top: 6px; right: 8px; background: transparent; border: none;
  font-size: 16px; cursor: pointer; color: #999; }
        `;
        document.head.appendChild(styleEl);

        const folderSel = el('select', { id: 'oee-folder' },
            el('option', { value: 'inbox' }, 'Inbox'),
            el('option', { value: 'sentitems' }, 'Sent Items'),
            el('option', { value: 'archive' }, 'Archive'),
            el('option', { value: 'drafts' }, 'Drafts'),
            el('option', { value: 'deleteditems' }, 'Deleted Items'),
            el('option', { value: 'junkemail' }, 'Junk')
        );
        const batchInput = el('input', { type: 'number', id: 'oee-batch', value: '500', min: '50', max: '5000' });
        const concInput = el('input', { type: 'number', id: 'oee-conc', value: '4', min: '1', max: '8',
            title: 'Concurrent message fetches. Higher than 4 can trip server-side throttling on this mailbox.' });
        const startBtn = el('button', { id: 'oee-start' }, '▶ Start');
        const pauseBtn = el('button', { id: 'oee-pause' }, '⏸ Pause');
        const stopBtn = el('button', { id: 'oee-stop' }, '⏹ Stop');
        const closeBtn = el('button', { id: 'oee-close', title: 'Close' }, '✕');
        const status = el('div', { id: 'oee-status' },
            'Waiting... click any email so the script captures bearer + anchor.');
        const readyState = el('span', null, 'waiting');
        const readyLine = el('div', { id: 'oee-ready' }, 'Status: ', readyState);

        const panel = el('div', { id: 'oee-panel' },
            closeBtn,
            el('h3', null, '📥 Outlook Email Exporter v3.6'),
            el('label', null, 'Folder'),
            folderSel,
            el('div', { className: 'row' },
                el('div', null, el('label', null, 'Emails per ZIP'), batchInput),
                el('div', null, el('label', null, 'Concurrency'), concInput)
            ),
            el('div', { className: 'btns' }, startBtn, pauseBtn, stopBtn),
            status,
            readyLine
        );
        document.body.appendChild(panel);

        setInterval(() => {
            readyState.textContent = '';
            const missing = [];
            if (!bearerAuth) missing.push('bearer');
            if (!anchorMailbox) missing.push('mailbox');
            if (missing.length === 0) {
                readyState.appendChild(el('span', { className: 'ok' }, '✓ Ready'));
                if (status.textContent.startsWith('Waiting')) {
                    status.textContent = 'Ready. Pick a folder and click ▶ Start.';
                }
            } else {
                readyState.appendChild(el('span', { className: 'miss' }, `missing: ${missing.join(', ')}`));
            }
        }, 700);

        startBtn.addEventListener('click', async () => {
            if (!ready()) {
                status.textContent = 'Not ready — click any email first to capture session.';
                return;
            }
            cancelled = false; paused = false;
            const folder = folderSel.value;
            const batch = Math.max(50, parseInt(batchInput.value, 10) || 500);
            const conc = Math.max(1, Math.min(8, parseInt(concInput.value, 10) || 4));
            status.textContent = 'Starting...';
            try { await exportAll(folder, batch, conc, status); }
            catch (e) { status.textContent = 'Error: ' + e.message; console.error(e); }
        });
        pauseBtn.addEventListener('click', () => {
            paused = !paused;
            pauseBtn.textContent = paused ? '▶ Resume' : '⏸ Pause';
        });
        stopBtn.addEventListener('click', () => { cancelled = true; status.textContent = 'Cancelling...'; });
        closeBtn.addEventListener('click', () => panel.remove());
    }

    if (document.body) buildUI();
    else document.addEventListener('DOMContentLoaded', buildUI);
})();