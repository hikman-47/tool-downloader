// ==UserScript==
// @name         CEISA 4.0 Tool Downloader by Hikman
// @namespace    http://tampermonkey.net/
// @version      2.0
// @description  Unduh Batch PDF (SPPB & Detail) dan Ekspor Excel dari portal CEISA 4.0 dalam satu panel. Gratis 3x pakai, selanjutnya perlu license key.
// @author       Hikman
// @updateURL    https://raw.githubusercontent.com/hikman-47/tool-downloader/main/ceisa-downloader.user.js
// @downloadURL  https://raw.githubusercontent.com/hikman-47/tool-downloader/main/ceisa-downloader.user.js
// @match        https://portal.beacukai.go.id/*
// @run-at       document-start
// @grant        GM_xmlhttpRequest
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// @connect      apis-gw.beacukai.go.id
// @connect      ceisa-tool.hikman47.workers.dev
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

(function () {
  'use strict';
  const SERVER = 'https://ceisa-tool.hikman47.workers.dev';
  const VER = '2.0';

  const W = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
  const API = 'https://apis-gw.beacukai.go.id';

  /* =========================================================================
     1. FUNGSI BERSAMA & PENANGKAP TOKEN (INTERCEPTOR)
     ========================================================================= */
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  class AuthError extends Error {}

  const auth = { token: '', key: '', idUser: '', nomorIdentitas: '', bumn: 'false', browseTpl: '' };

  const take = (k, v) => {
    const n = String(k || '').toLowerCase();
    if (!v) { return; }
    if (n === 'authorization') { auth.token = String(v).replace(/^Bearer\s+/i, '').trim(); }
    else if (n === 'beacukai-api-key') { auth.key = String(v).trim(); }
  };

  const grabHeaders = (h) => {
    try {
      if (!h) { return; }
      if (typeof W.Headers !== 'undefined' && h instanceof W.Headers) { h.forEach((v, k) => take(k, v)); }
      else if (Array.isArray(h)) { h.forEach((p) => take(p[0], p[1])); }
      else { Object.keys(h).forEach((k) => take(k, h[k])); }
    } catch (e) { /* abaikan */ }
  };

  const grabUrl = (u) => {
    try {
      const x = new URL(String(u), location.origin);
      const p = x.searchParams;
      if (p.get('idUser')) { auth.idUser = p.get('idUser'); }
      if (p.get('nomorIdentitas')) { auth.nomorIdentitas = p.get('nomorIdentitas'); }
      if (p.get('isBumnEkspor')) { auth.bumn = p.get('isBumnEkspor'); }
      if (/dokumen-pabean-portal/.test(x.pathname)) { auth.browseTpl = x.search.replace(/^\?/, ''); }
    } catch (e) { /* abaikan */ }
  };

  const origFetch = W.fetch;
  W.fetch = function (...args) {
    try {
      const a0 = args[0];
      grabUrl(a0 && a0.url ? a0.url : a0);
      if (a0 && typeof a0 === 'object' && a0.headers) { grabHeaders(a0.headers); }
      if (args[1] && args[1].headers) { grabHeaders(args[1].headers); }
    } catch (e) { /* abaikan */ }
    return origFetch.apply(this, args);
  };

  const xOpen = W.XMLHttpRequest.prototype.open;
  W.XMLHttpRequest.prototype.open = function (m, u, ...r) {
    grabUrl(u);
    return xOpen.call(this, m, u, ...r);
  };
  const xSet = W.XMLHttpRequest.prototype.setRequestHeader;
  W.XMLHttpRequest.prototype.setRequestHeader = function (k, v) {
    take(k, v);
    return xSet.call(this, k, v);
  };

  const claims = () => {
    try {
      if (!auth.token) { return {}; }
      const p = auth.token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      return JSON.parse(decodeURIComponent(escape(atob(p))));
    } catch (e) { return {}; }
  };
  const getIdUser = () => auth.idUser || claims().id_user || '';
  const getIdent = () => auth.nomorIdentitas || claims().nitku || '';
  const tokenLeftMs = () => {
    const e = claims().exp;
    return e ? e * 1000 - Date.now() : null;
  };



  /* =========================================================================
     LISENSI: Device ID, panggilan server, panel kunci
     ========================================================================= */
  const devId = (() => {
    let d = '';
    try { d = String(GM_getValue('ck_dev', '') || ''); } catch (e) {}
    if (!/^[A-Za-z0-9-]{16,64}$/.test(d)) {
      d = (W.crypto && W.crypto.randomUUID) ? W.crypto.randomUUID() : (Date.now().toString(36) + Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2));
      try { GM_setValue('ck_dev', d); } catch (e) {}
    }
    return d;
  })();

  const call = (path, params) => new Promise((resolve) => {
    const q = Object.keys(params || {}).map((k) => encodeURIComponent(k) + '=' + encodeURIComponent(params[k])).join('&');
    GM_xmlhttpRequest({
      method: 'GET', url: SERVER + path + '?d=' + encodeURIComponent(devId) + (q ? '&' + q : ''),
      timeout: 20000,
      onload: (r) => { try { resolve(JSON.parse(r.responseText)); } catch (e) { resolve({ ok: false, reason: 'bad_response' }); } },
      onerror: () => resolve({ ok: false, reason: 'network' }),
      ontimeout: () => resolve({ ok: false, reason: 'network' })
    });
  });

  const st = { mode: 'free', left: 0, pay: {} };
  const C = {
    W, API, sleep, AuthError, auth, claims, getIdUser, getIdent, tokenLeftMs, origFetch,
    GM_xmlhttpRequest, XLSX: (typeof XLSX !== 'undefined' ? XLSX : W.XLSX), ver: VER,
    onStatus: null,
    statusText: () => (st.mode === 'full' ? 'Lisensi aktif ✔' : 'Gratis · sisa ' + st.left + 'x'),
    showLock: (why) => showLock(why),
    gate: async (kind) => {
      const r = await call('/use', { a: kind });
      if (r && r.mode) { st.mode = r.mode; if (typeof r.left === 'number') { st.left = r.left; } if (C.onStatus) { C.onStatus(); } }
      if (r && r.ok) { return true; }
      if (r && r.reason === 'quota') { showLock('quota'); }
      else { alert('Tidak bisa menghubungi server lisensi. Cek koneksi internet lalu coba lagi.'); }
      return false;
    }
  };

  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  function showLock(why) {
    const old = document.getElementById('ckLockBox'); if (old) { old.remove(); }
    const p = st.pay || {};
    const full = st.mode === 'full';
    const box = document.createElement('div');
    box.id = 'ckLockBox';
    box.style.cssText = 'position:fixed;inset:0;z-index:2147483647;background:rgba(0,0,0,.55);display:flex;align-items:center;justify-content:center;font:13px/1.45 system-ui,Segoe UI,Arial,sans-serif;';
    const head = full ? 'Lisensi aktif ✔' : (why === 'quota' ? 'Jatah gratis habis' : 'Lisensi');
    const sub = full ? 'Panel berjalan dalam versi penuh.' : (why === 'quota' ? 'Kamu sudah memakai 3x jatah gratis. Aktifkan license key untuk lanjut.' : 'Gratis ' + st.left + 'x lagi. Masukkan license key untuk versi penuh.');
    box.innerHTML = `
      <div style="background:#fff;color:#222;border-radius:10px;width:min(440px,92vw);padding:18px 20px;box-shadow:0 10px 40px rgba(0,0,0,.4);">
        <div style="font-size:16px;font-weight:700;margin-bottom:2px;">${esc(head)}</div>
        <div style="color:#555;margin-bottom:12px;">${esc(sub)}</div>
        ${full ? '' : `
        <div style="background:#f6f8fa;border:1px solid #e1e4e8;border-radius:8px;padding:10px 12px;margin-bottom:12px;">
          <b>Cara beli</b><br>
          Harga: <b>${esc(p.harga || '-')}</b><br>
          Transfer ke <b>${esc(p.bank || '')} ${esc(p.rek || '')}</b><br>a.n. ${esc(p.an || '')}<br>
          Kirim bukti transfer + <b>Device ID</b> di bawah ke CP: <b>${esc(p.cp || '')}</b>
        </div>`}
        <div style="margin-bottom:10px;">Device ID:<br>
          <div style="display:flex;gap:6px;"><input id="ckDevId" readonly value="${esc(devId)}" style="flex:1;padding:5px;font:12px monospace;user-select:all"><button id="ckDevCopy" style="padding:5px 10px;cursor:pointer">Salin</button></div>
        </div>
        ${full ? '' : `
        <div style="margin-bottom:6px;">License key:<br><input id="ckKeyIn" placeholder="CEISA-XXXX-XXXX-XXXX-XXXX" style="width:100%;box-sizing:border-box;padding:6px;font:13px monospace;text-transform:uppercase"></div>
        <div id="ckKeyMsg" style="min-height:18px;color:#b00020;margin-bottom:8px;"></div>`}
        <div style="display:flex;gap:8px;justify-content:flex-end;">
          ${full ? '' : '<button id="ckKeyGo" style="background:#1fae4b;color:#fff;font-weight:700;border:0;padding:7px 14px;border-radius:6px;cursor:pointer">Aktifkan</button>'}
          <button id="ckLockClose" style="padding:7px 14px;border-radius:6px;cursor:pointer;border:1px solid #bbb;background:#f3f3f3">Tutup</button>
        </div>
      </div>`;
    document.body.appendChild(box);
    const $$ = (s) => box.querySelector(s);
    $$('#ckLockClose').onclick = () => box.remove();
    $$('#ckDevCopy').onclick = async () => { try { await navigator.clipboard.writeText(devId); $$('#ckDevCopy').textContent = 'Tersalin'; } catch (e) { $$('#ckDevId').select(); } };
    if (!full) {
      $$('#ckKeyGo').onclick = async () => {
        const k = $$('#ckKeyIn').value.trim();
        if (!k) { $$('#ckKeyMsg').textContent = 'Isi license key dulu.'; return; }
        $$('#ckKeyMsg').style.color = '#555'; $$('#ckKeyMsg').textContent = 'Memeriksa…';
        const r = await call('/activate', { k });
        if (r && r.ok) {
          st.mode = 'full'; if (C.onStatus) { C.onStatus(); }
          $$('#ckKeyMsg').style.color = '#1a7f37'; $$('#ckKeyMsg').textContent = 'Aktif! Memuat ulang halaman…';
          setTimeout(() => location.reload(), 900);
        } else {
          const m = { invalid_key: 'Key tidak valid.', key_other_device: 'Key ini sudah dipakai di perangkat lain.', network: 'Tidak bisa terhubung ke server.' };
          $$('#ckKeyMsg').style.color = '#b00020'; $$('#ckKeyMsg').textContent = m[r && r.reason] || 'Gagal mengaktifkan.';
        }
      };
    }
  }

  /* =========================================================================
     BOOT: minta kode inti ke server, lalu jalankan
     ========================================================================= */
  const onReady = (fn) => { if (document.body) { fn(); } else { document.addEventListener('DOMContentLoaded', fn); } };
  (async () => {
    const r = await call('/boot', { v: VER });
    if (!r || !r.ok) { console.warn('[CEISA] boot gagal', r && r.reason); return; }
    st.mode = r.mode; st.left = r.left || 0; st.pay = r.pay || {};
    if (!r.code) { onReady(() => showLock('quota')); return; }   // jatah habis & belum berlisensi
    try { new Function('C', r.code)(C); }
    catch (e) { console.error('[CEISA] gagal menjalankan kode inti', e); }
  })();
})();
