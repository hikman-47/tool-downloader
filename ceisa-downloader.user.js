// ==UserScript==
// @name         CEISA 4.0 Tool Downloader by Hikman
// @namespace    http://tampermonkey.net/
// @version      3.3
// @description  Unduh Batch PDF (SPPB & Detail) dan Ekspor Excel dari portal CEISA 4.0 dalam satu panel. Gratis 3x pakai per akun, selanjutnya perlu license key.
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
  const VER = '3.3';

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
      if (p.get('nomorIdentitas')) {
        const baru = !auth.nomorIdentitas;
        auth.nomorIdentitas = p.get('nomorIdentitas');
        if (baru) { setTimeout(() => { refreshStat(); }, 800); } // identitas baru tertangkap -> betulkan footer
      }
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
     2. IDENTITAS EXIM (v3): e = SHA-256(NITKU|username), n = SHA-256(NITKU)
     Server hanya menerima HASH — NITKU & username mentah tidak pernah dikirim.
     ========================================================================= */
  const utf8Bytes = (s) => {
    const out = [];
    for (let i = 0; i < s.length; i++) {
      let c = s.charCodeAt(i);
      if (c < 0x80) { out.push(c); }
      else if (c < 0x800) { out.push(0xc0 | (c >> 6), 0x80 | (c & 0x3f)); }
      else if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length) {
        const cp = 0x10000 + ((c - 0xd800) * 0x400) + (s.charCodeAt(++i) - 0xdc00);
        out.push(0xf0 | (cp >> 18), 0x80 | ((cp >> 12) & 0x3f), 0x80 | ((cp >> 6) & 0x3f), 0x80 | (cp & 0x3f));
      } else { out.push(0xe0 | (c >> 12), 0x80 | ((c >> 6) & 0x3f), 0x80 | (c & 0x3f)); }
    }
    return new Uint8Array(out);
  };
  const sha256hex = async (s) => {
    try {
      const buf = await W.crypto.subtle.digest('SHA-256', utf8Bytes(s));
      return Array.from(new Uint8Array(buf)).map((b) => b.toString(16).padStart(2, '0')).join('');
    } catch (e) { return ''; }
  };

  let identCache = null; // { nitku, user, e, n }
  async function computeIdent() {
    const nitku = String(getIdent() || '').trim();
    if (!nitku) { return null; }
    const cl = claims();
    const user = String(cl.preferred_username || cl.username || '').toLowerCase().trim();
    if (identCache && identCache.nitku === nitku && identCache.user === user) { return identCache; }
    const n = await sha256hex(nitku);
    if (!n) { return null; }
    const e = user ? await sha256hex(nitku + '|' + user) : '';
    identCache = { nitku, user, e, n };
    try { GM_setValue('ck_ident', JSON.stringify({ e, n, u: user })); } catch (err) { /* abaikan */ }
    return identCache;
  }

  // Parameter identitas untuk panggilan server. Menunggu sampai waitMs agar
  // tangkapan portal (tabel dokumen) sempat tiba; hasil terbaik yang ada:
  // { e, n } lengkap, { n } bila username belum terbaca, {} bila belum ada.
  async function identParams(waitMs) {
    const deadline = Date.now() + (waitMs || 0);
    for (;;) {
      const id = await computeIdent();
      if (id && id.e) { const o = { e: id.e, n: id.n }; if (id.user) { o.u = id.user; } return o; }
      if (Date.now() >= deadline) {
        if (id && id.n) { return { n: id.n }; }
        return {};
      }
      await sleep(200);
    }
  }

  const cachedIdent = () => {
    try {
      const j = JSON.parse(String(GM_getValue('ck_ident', '') || ''));
      if (j && (j.e || j.n)) { return j; }
    } catch (e) { /* abaikan */ }
    return null;
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
  // Server mengirim pay sebagai string JSON; panel membacanya sebagai objek.
  const payObj = (v) => { if (v && typeof v === 'object') { return v; } try { return JSON.parse(String(v || '')) || {}; } catch (e) { return {}; } };

  // Footer otoritatif per exim: sesudah identitas tertangkap, angka /boot
  // (basis perangkat) dikoreksi lewat /stat. Worker lama tanpa /stat -> diabaikan.
  async function refreshStat() {
    const id = await identParams(0);
    if (!id.e && !id.n) { return; }
    const r = await call('/stat', id);
    if (r && r.ok && r.mode) {
      st.mode = r.mode;
      if (typeof r.left === 'number') { st.left = r.left; }
      if (C.onStatus) { C.onStatus(); }
    }
  }

  const C = {
    W, API, sleep, AuthError, auth, claims, getIdUser, getIdent, tokenLeftMs, origFetch,
    GM_xmlhttpRequest, XLSX: (typeof XLSX !== 'undefined' ? XLSX : W.XLSX), ver: VER,
    onStatus: null,
    statusText: () => (st.mode === 'full' ? 'Lisensi aktif ✔' : 'Gratis · sisa ' + st.left + 'x'),
    showLock: (why) => showLock(why),
    gate: async (kind) => {
      const id = await identParams(2500);
      const r = await call('/use', Object.assign({ a: kind }, id));
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
    let storedKey = '';
    try { storedKey = String(GM_getValue('ck_key', '') || ''); } catch (e) {}
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
          Kirim bukti transfer ke CP: <b>${esc(p.cp || '')}</b>
        </div>`}
        <div style="margin-bottom:10px;">Device ID (referensi dukungan):<br>
          <div style="display:flex;gap:6px;"><input id="ckDevId" readonly value="${esc(devId)}" style="flex:1;padding:5px;font:12px monospace;user-select:all"><button id="ckDevCopy" style="padding:5px 10px;cursor:pointer">Salin</button></div>
        </div>
        ${full ? `
        <div style="border-top:1px solid #eee;padding-top:10px;margin-bottom:6px;">Pindah perusahaan? Lepaskan lisensi dari akun ini dulu:<br>
          <div style="display:flex;gap:6px;margin-top:4px;"><input id="ckRelKey" value="${esc(storedKey)}" placeholder="CEISA-XXXX-XXXX-XXXX-XXXX" style="flex:1;padding:6px;font:13px monospace;text-transform:uppercase"><button id="ckRelGo" style="padding:6px 12px;border-radius:6px;cursor:pointer;border:1px solid #bbb;background:#f3f3f3">Lepaskan</button></div>
        </div>
        <div id="ckRelMsg" style="min-height:18px;color:#b00020;margin-bottom:8px;"></div>` : `
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
        const id = await identParams(3000);
        if (!id.e) {
          $$('#ckKeyMsg').style.color = '#b00020';
          $$('#ckKeyMsg').textContent = 'Akun CEISA-mu belum terbaca. Buka dulu tabel Dokumen Pabean sampai daftarnya tampil, lalu klik Aktifkan lagi — agar lisensi menempel ke akunmu, bukan ke perangkat ini.';
          return;
        }
        const ap = { k, e: id.e }; if (id.n) { ap.n = id.n; } if (id.u) { ap.u = id.u; }
        const r = await call('/activate', ap);
        if (r && r.ok) {
          try { GM_setValue('ck_key', k.toUpperCase()); } catch (e) {}
          st.mode = 'full'; if (C.onStatus) { C.onStatus(); }
          $$('#ckKeyMsg').style.color = '#1a7f37'; $$('#ckKeyMsg').textContent = 'Aktif! Memuat ulang halaman…';
          setTimeout(() => location.reload(), 900);
        } else {
          const m = {
            invalid_key: 'Key tidak valid.',
            key_other_user: 'Key ini sudah terikat ke akun CEISA lain.',
            key_other_device: 'Key ini sudah dipakai di perangkat lain.',
            transfer_cooldown: 'Key ini baru dilepas dari akun lain. Bisa dipindah lagi dalam ' + ((r && r.waitDays) || '?') + ' hari — atau hubungi penjual untuk dibantu pindah sekarang.',
            network: 'Tidak bisa terhubung ke server.'
          };
          $$('#ckKeyMsg').style.color = '#b00020'; $$('#ckKeyMsg').textContent = m[r && r.reason] || 'Gagal mengaktifkan.';
        }
      };
    } else {
      $$('#ckRelGo').onclick = async () => {
        const k = $$('#ckRelKey').value.trim();
        if (!k) { $$('#ckRelMsg').textContent = 'Isi license key yang mau dilepas dulu.'; return; }
        $$('#ckRelMsg').style.color = '#555'; $$('#ckRelMsg').textContent = 'Memeriksa…';
        const id = await identParams(3000);
        if (!id.e) {
          $$('#ckRelMsg').style.color = '#b00020';
          $$('#ckRelMsg').textContent = 'Akun CEISA-mu belum terbaca. Buka dulu tabel Dokumen Pabean sampai daftarnya tampil, lalu coba lagi.';
          return;
        }
        const rp = { k, e: id.e }; if (id.n) { rp.n = id.n; } if (id.u) { rp.u = id.u; }
        const r = await call('/release', rp);
        if (r && r.ok) {
          try { GM_setValue('ck_key', ''); } catch (e) {}
          $$('#ckRelMsg').style.color = '#1a7f37'; $$('#ckRelMsg').textContent = 'Lisensi dilepas dari akun ini. Memuat ulang halaman…';
          setTimeout(() => location.reload(), 900);
        } else {
          const m = { invalid_key: 'Key tidak valid.', not_bound: 'Key itu tidak terikat ke akun ini.', network: 'Tidak bisa terhubung ke server.' };
          $$('#ckRelMsg').style.color = '#b00020'; $$('#ckRelMsg').textContent = m[r && r.reason] || 'Gagal melepas lisensi.';
        }
      };
    }
  }

  /* =========================================================================
     BOOT: minta kode inti ke server, lalu jalankan
     ========================================================================= */
  const onReady = (fn) => { if (document.body) { fn(); } else { document.addEventListener('DOMContentLoaded', fn); } };

  function runCore(code) {
    try { new Function('C', code)(C); }
    catch (e) { console.error('[CEISA] gagal menjalankan kode inti', e); }
  }

  // /boot menolak memberi core (jatah perangkat habis), tetapi exim-nya bisa
  // saja masih punya jatah / berlisensi. Begitu identitas tertangkap dari
  // lalu lintas portal, periksa /stat; bila berhak, minta core sekali lagi
  // dengan identitas itu dan jalankan.
  async function rescueWatch() {
    for (let i = 0; i < 24; i++) {
      await sleep(1500);
      const id = await identParams(0);
      if (!id.e && !id.n) { continue; }
      const s = await call('/stat', id);
      if (!s || !s.ok || !s.mode) { continue; }
      if (s.mode !== 'full' && !(s.left > 0)) { return; }
      const bp = { v: VER }; if (id.e) { bp.e = id.e; } if (id.n) { bp.n = id.n; } if (id.u) { bp.u = id.u; }
      const r = await call('/boot', bp);
      if (r && r.ok && r.code) {
        st.mode = r.mode; st.left = r.left || 0; st.pay = payObj(r.pay);
        const old = document.getElementById('ckLockBox'); if (old) { old.remove(); }
        runCore(r.code);
        coreOn = true; syncUi();
      }
      return;
    }
  }

  /* =========================================================================
     GERBANG ALAMAT (v3.1): core hanya dinyalakan di kawasan /dokumen-pabean/.
     Portal berpindah halaman tanpa memuat ulang, jadi loader memantau
     perpindahan alamat: masuk kawasan -> boot (atau tampilkan lagi),
     keluar kawasan -> tampilan tool disembunyikan (proses unduhan yang
     sedang berjalan TIDAK dihentikan, hanya tampilannya yang disembunyikan).
     Di luar kawasan loader diam: tidak ada panggilan /boot, tidak ada ikon.
     ========================================================================= */
  const IN_AREA = () => /^\/dokumen-pabean(\/|$)/.test(location.pathname);
  let coreOn = false;    // core sudah dijalankan di dokumen ini
  let bootBusy = false;  // boot sedang berjalan
  let uiHidden = false;  // tampilan sedang disembunyikan oleh gerbang

  function hideEl(el) {
    if (!el) { return; }
    if (el.dataset.ckDisp === undefined) { el.dataset.ckDisp = el.style.display || ''; }
    el.style.display = 'none';
  }
  function showEl(el) {
    if (!el) { return; }
    el.style.display = (el.dataset.ckDisp !== undefined ? el.dataset.ckDisp : '');
  }
  function syncUi() {
    const p = document.getElementById('ckPanelMaster');
    const l = document.getElementById('ckLockBox');
    if (!IN_AREA()) {
      if (p || l) { hideEl(p); hideEl(l); uiHidden = true; }
    } else if (uiHidden) {
      showEl(p); showEl(l); uiHidden = false;
    }
  }

  async function startBoot() {
    if (coreOn || bootBusy) { return; }
    bootBusy = true;
    try {
      const ci = cachedIdent();
      const bp = { v: VER }; if (ci) { if (ci.e) { bp.e = ci.e; } if (ci.n) { bp.n = ci.n; } if (ci.u) { bp.u = ci.u; } }
      const r = await call('/boot', bp);
      if (!r || !r.ok) { console.warn('[CEISA] boot gagal', r && r.reason); return; }
      st.mode = r.mode; st.left = r.left || 0; st.pay = payObj(r.pay);
      coreOn = true;
      if (!r.code) { onReady(() => { if (IN_AREA()) { showLock('quota'); } rescueWatch(); }); return; }   // jatah habis & belum berlisensi
      runCore(r.code);
      syncUi();
      setTimeout(syncUi, 400);
      setTimeout(() => { refreshStat(); }, 1200);
    } finally { bootBusy = false; }
  }

  let lastPath = location.pathname;
  function onRoute() {
    lastPath = location.pathname;
    if (IN_AREA() && !coreOn) { startBoot(); return; }
    syncUi();
  }
  try {
    const h = W.history;
    if (h) {
      const ps = h.pushState, rs = h.replaceState;
      h.pushState = function (...a) { const r = ps.apply(this, a); setTimeout(onRoute, 0); return r; };
      h.replaceState = function (...a) { const r = rs.apply(this, a); setTimeout(onRoute, 0); return r; };
    }
  } catch (e) { /* abaikan: pemantau berkala di bawah tetap bekerja */ }
  W.addEventListener('popstate', onRoute);
  W.addEventListener('hashchange', onRoute);
  setInterval(() => { if (location.pathname !== lastPath) { onRoute(); } else { syncUi(); } }, 700);

  if (IN_AREA()) { startBoot(); }
})();
