// ==UserScript==
// @name         CEISA 4.0 Tool Downloader by Hikman
// @namespace    http://tampermonkey.net/
// @version      3.6
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
  const VER = '3.6';

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

  // Username CEISA yang sedang login (untuk ditampilkan di kotak lisensi).
  const myUser = () => {
    if (identCache && identCache.user) { return identCache.user; }
    const c = cachedIdent();
    return (c && c.u) || '';
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
    statusText: () => {
      const u = myUser();
      const dasar = st.mode === 'full' ? 'Lisensi aktif ✔' : 'Gratis · sisa ' + st.left + 'x';
      return u ? dasar + ' · ' + u : dasar;
    },
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

  // Nomor CP dari server (pay.cp) dijadikan tautan WhatsApp: nomor tampil
  // tetap seperti aslinya (format 08...), href memakai format internasional
  // 62... lewat wa.me. Teks di luar nomor tetap di-escape seperti biasa.
  const cpHtml = (raw) => {
    const s = String(raw == null ? '' : raw);
    const m = s.match(/(\+?62|0)8\d{7,12}/);
    if (!m) { return esc(s); }
    const num = m[0];
    const intl = num.charAt(0) === '+' ? num.slice(1) : (num.charAt(0) === '0' ? '62' + num.slice(1) : num);
    return esc(s.slice(0, m.index)) +
      '<a href="https://wa.me/' + intl + '" target="_blank" rel="noopener" style="color:#1fae4b;text-decoration:underline;">' + esc(num) + '</a>' +
      esc(s.slice(m.index + num.length));
  };

  /* =========================================================================
     GACHA HARGA (v3.5): harga lisensi ditentukan undian — hasil TERAKHIR yang
     berlaku, harga sebelumnya hangus. Kesempatan 3x, tersimpan (GM) agar
     tidak bisa diakali dengan menutup-membuka panel.
     Drop rate final: <10rb 20% · 10–49rb 40% · 50–99rb 20% ·
     100–499rb 12% · 500–999rb 6% · pas Rp1.000.000 2%.
     Kalimat pop-up: pool kalimat-gacha.md (21 Nov 2026), dipakai persis.
     ========================================================================= */
  const QRIS_B64 = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAssAAALLCAAAAAAqY5AHAAAUDklEQVR42u3d0Zaktg5A0dKs/v9f1n1IcjOTDB0b2cbAPk+ZVAHGPhhRqOXID/AIfugCcBngMsBlgMvgMsBlgMsAl8FlgMsAlwEuA1wGlwEuA1wGuAwuA1wGuAxwGeAyuAxwGeAywGVwGeAywGWAywCXwWWAywCXAS6DywCXAS4DXAa4DC4DXAa4DHAZXAa4DHAZ4DJezVf/JjGpKdl0wGz54HDz7D6RrHVQ1nouJ3V8rRePN1+hg3n5bMfHjdsuxgC4DHAZ4DLAZXAZ4DLAZYDL4DKG4GXdwr7O6vDkqIHOho9OHTtyull50KdH7Y3PsE684grLtkEsNWt9PkbmKAO+Oas82CJ7D7NGmfz8Z3szr1R53vFWn4gYA+LlDUhDBvMyuOx3AXAZ4LJ4GWIMcNm8DC6bl8Fl4Pku50NijOMMg9ymWY/ha3kP5SSZo3TFDMyQaUyRaaqX0lbrpalyS3/iTuyjgxhjg8k/V08VYgyAywCXAS4DXAaXAS4DXAa4DC7fmU3XYvCybh1fN1Az5wj8a6pPtOz2xFI4sfgKbTuRtu5t6QUujxnDWclMUTt4cZWl/kyhkUe/m8Di5Xew4s7BZexDchmemrkMMy6X8c4Zl8vmZS7DvMxlgMsAlyFevo/L8nOu6ZK7XQqb5mP8mlQQld6O/q1ztWh52N5sOZH+IjCPlPkG83LW+vrWN1plTcXL4DLAZYDLAJcBLoPLAJcBLmMu/t4PT+Ex6S4j8zEyBs4CMX9M4rBVh5/EYp3iQplzJx1WuzywYTFwk5aVp+IwDSeWL9J06eJPn6M+uUHwop4cxMueYMBlgMsAl8FlgMsAlwEug8sAl4FdqOZjrC6gE91b1NZM+ozsn6ZKH3HU3sMPPt1bRH/3LljY6w3zci7Y4m5nWNttsWLTprkvP9416o/MQIotLyQuYx/JuWx4trlX5bu61Lz84OnXvAxwGXd8HuYySM5leFDmMvqUNS8D97ymc+A8MDC6i0Fb/LvgSP7xzX99kJ/ImHoix8397pMY1Flt7W3qxVxgw4mbStzgPhQ1+2vFWqJv1JtHoS3pqFTAptreqKklxgC4DC4DXAa4DHAZ4DK4DHAZ4DLA5RH0J4Ngh2HLOSPd6EBM2lXOOY+c0yUnjnFiHZ7sbntTEZimqjMnPrjbvHxpjlUuPl6xT7LYrCztd8mfdMedXcYtp4pZJJcBLj+cTR9Pg8twiXFZvCzGALgM8TKX3U/xdpe9UuayuRi7jNTt3vvFgt1G3Picml5iR3+flO5oTYvyxLUzzbTcov5iJG3HqBVoaWpJNY9m4LWQk66eWjmaJdll/V3qd4yzvSjAES8DXAa4DC4DXAa4DHAZXAa4DHD5DyS33ZKBb+xHvj0trr0TtZ4oLhRz4l1+fw5G93kMdCAm+dS/qk40nnoOc2Z5PkYuv6jHTQkDWzLrnHLSEYvLzx1/kFcKIF4WMYiX8RCCyzAvcxnmZS4DXAaXAS7Dsx+X4dmPyxBjbDENxLipInY5p02PGLUtdq1Rnlf2adtCUE2JRsWMlf5VmkamyDSVjZlV/2ZkxZzuwT2xOtamMcbAAjgj54pLV5h6fg2Zw2Gz9g7AZXAZ4DLAZYDL4DLAZYDLeDPB5Q7yZsODnr7OHccqmw6S3U3JS0XLtoN3r3KT407qcFf9xzjeIk8M7qUu11YqutucObDm0on0nuIZ1sop9a9UlJN0EC/fP44Blz1LcRngMsBlgMvgMsBlgMsYSnIZmEns6bLUmTdKs+nt4mt5D60oRlLLQziuP1JbbaoxnShaPhh48LbBPSxqEZf6K8YQZoqX4brgMsBlgMvgMsBlgMu4DGvv4ClYewfg8hNvb3E0z0lYWceJfIyMYZZG9wFzkoA5cF/9pxiXThUrmjhpCKsuF4uRnMjVyTkj+stkmnO64cRaTDm/S4qd1b95savXxxhyD3ClGiNdFhvivc9+AJcBLoPLAJcBLgNcBpeBl7p859cjseCkigU1nsksaWLkgi6lcivFmiFtrTrMJ4qmZuW4/in2dVM/NHZWKU0kqsMz7jreJ8bISYO7vEv3mVhHdlbufxsRL+Nm0xmX4dnv7Q8U4DJecmfnshFxR+MywGVwGeAy8BaXH/nA5tfBi3ruTLGCtiVkGm3ubv+fW2T840qIpuslBl578bvGZ5Sv6GjYR/dZVbuhb/vm0Rk5zRULbzQuudS0YlMx1ae/ClFxQGt5P/0dd+LUo9SqT/cBo2ZjcfNijJGzrrGjzW+QuPP8lsSkgX7Zs59g9JaScxngMrgMcBngMopIEeQyuAw83eXimyWcRPeOd/nSn+YN6G6B9Hodclbbe3Mwon/rtkHsrr1yfCIDy7jEan+7kx0ac22KTRz4PvzHrF7sXg1p/SSSc07ErP7KZz+BAZ7iMsBlgMvgMsBlgMsAl8FlgMtAP7Gny20v/BtXb4oL3wmeOHLE0zX76QxjI39/2e+sdaSaLC8uGVastpKlk6oOTzEbKRfYVFoX7MSkd22tF1zEU5a6Ei8DXAaXAS4DXAa4DC4DXEYnCmdt7bLh6cCf7W7tsuFZd+GHkfr3WVXzMfL3/7+x2krDkH5TQmZoFsTve+KbJIiWFI4YKHBj9knTQjzfNPH//dDW9jMdNKd/yvNy9uv789Y55jvzJso8Mbcd9cmsICyamhWNTcy+tmfp2hs6uGq94Clxl98x4NkP4DLAZYDL4DImPXSDy9vg90QuA1zeKzAQYwy/o8XjXXYzf4vX1XyrvLF0OatTxs3Fs8rcjDxGbb85rkvVenn9LV+049mP5FzGO56luAxwGeAywGVwGTcgB36Ly7gSv2PgXfMyl3ebd2RtnO2HWHCNLBmdr5HTQNxq4HKgA021QbLYWbG6TxZompu63JakcrhFf2ZJbYsYuI7UicHaZtGtz1GfDNzvknBHvDz3niDvR7wMcBlcBrgMcBngMsBlcBngMvACl+O3/9m/cev2J9ZU6W5KlI5x4qRm0dRZcea0JsmUa429ekgOW5UNHxzmSsSk9kbTiVQHIRcPYTYd5MW1XgZ2aY7b1awm7tNZYgyAywCXwWWAywCXAS6jF7Ve8BSCywCXgSk3mFnL1GwaDsac/caWZ9h/6rlioJualctdHrgaUuN+m8qM9GevFNdMGlin5kStl2J7V59hcb+zYgylTDBrwhYvw7MfwGWAywCXAS6DywCXAS4DXG4iuj/Yp4mrdxs3ONuL196Jhq2PP8liDx0cPRrXqek+eg5sYktfRSxwq20ID7912N671XrZc/2a/MxReXkTl+dAXntD+7G8g3H3COUN8bKlT+HZDxdGIlyGGIPLAJchxuAyuAzxMpeBG7vsXYl4+UpG5mP039IyKkfJ5bfX2hnmrKPHQy6eYmd9DWx9rDj5n1dRqxUQiTsM76TRyXH9kNt01t3i5XiIZfDsB3AZXAa4DHAZ4DK4DHAZ4DLA5T/xx7KG7Se+prX4qDJJzOmVwwIi5X7M3j2tOEZt+ZslRWCO2httS11tVOslt7zYY88T6S/MlKVjrE9FyYODj5zmxMtu5uJlgMsAlyUd4zku+xkOYgxwGeCyeBlcFi+78MUYeO+F/+MG/WguRpMnuUCtE4vZ5AKX+yuhNH3QdlKH3xq4q2KQEcUB7R7cYky0JMYwseqsBae+xGU/PqBFhzvMy8Cdn/3cNvFEl8UYeIrLAJcBLoPLAJeBGv4Oe7s+xTU9Grmg7dl/Wtm14/h2Qar87Uk2d282nda3ffxfvVzrk3/u6/+HO5FQ0d+S6B7oWbkkk+blXPKTcv79H/ndt7LSxPob11z6A3v+fbjlv+tfe0N7TIzhfczrYyXxMrgMcBngMsBlcBngMsBlgMvg8mTi4L+P2ect3kbvE2PHQ8SCXbVun+NO/kTCSpbOMbtbsqlZKyrbNKZrdW9RLIwzsDaNGOMt5P4HVx8DGO6yFHY8xWVplzAvA+Zl4HHzspsCl59yIm4KXNYF4DLAZTHyG2jKwhk5hCPzMU688W8LeEsnnE27zcdfJFk6wVzRPUUXv/a5jLP7HJtSWdqu1qil9xxvnaUZIcbt6kzjx815K7KnxMsLZ0NNFC/T11MHl/EiuAwuA1wGuAxwGVwG3uiytEvc2OU4/Mf+rMl3udeMcOv3Jifes7clDxyNT2O2QnfllsaKJdHU9uz9oLiAUq3fq99qGs9axsmJXrxBvJyT5oeR+805Mj6TXLDFRi5TgPAL1PA7Bjz7AVwGuDz0aRFcBt7r8oo5028lXP7wDFwWL2OnMbw2xhg4YZv717FpX4+sj5Hd2TqxoOeyv+1RO9uma7pWXCb7TyRr+43+k1p9jYzMLerfIlZf+gMzb6q3p/7qJznu4imWm6qVoxEvuzV7huAyuAw3Ai5jC8QYAJcBLgNcBriMP1BLHE+xVC1xc0U3cWmfxAJ9rx22r5ENHpjKcujAyMIvk9Qcuc5MaWmuX1sSLR+0Hbw2IjluoKfNy+/6aX7J2T5l6Z4lnSVehngZ4DLAZdyU5LLRBZfBZbPhIl5cA8Ha7gwyoXDZHUKMAcy4H866OX4NbeXP///voiN59MHYSSv6vpUrRq7xIN90ybCB/+cx4j8P3XYiZ8ZzziDEwOyrkUk1/dVEYptdDax+0l91Zv06UuPGc6PcIuAp8bLHGX3i2Q/Yy2U/3sK8DIiXATEGuExfHfRIl1OMIe7y7Adw+U4383hiqx66jlT2fhBNX4soDkn2jsJBLkDU+iT6d5CHu9rnquqW+bA6TAxMXlk+L+eKKzyzNNI5rIE5zrnnRN456azEy9jmETi5DM88XMYNZ28u4/FwGVwGuAxwGeAyuAxw+YD45l+d20/LMIjuLYrfatkiRvb76u1jRRtP1HqJcT5UR7TpW7POcPmVXzuRrPXi4RYDK7e8udaLRPcRHWfdVezPkjv7M132t223vG+Zl93yzcvmZXieEC/D5MJliJfd0p5vqXgZ4DLA5UV3x8fcdF/27BdDKjjk5xO/i87i88nv1xnKT/y1fVPYV1z2pTjuTRknOXa/+Ym/umlc8Hw6pybjN23J4y93DUg1+yQvnR+yt39nuRzdu4qB196s1ZuKLnc7V8wzso4U9mfJT1xchmc/gMsAlwEug8sAlwEuA/dzORdsgQ2Ix7h8/NKyqXJLrO6UcUNy6YI3a3oxhn1QLkczsIdyxegOLO9yuLbRwJYsKXJSvKwG1nqZtMWmMUbuv9+Ra0eJie4ZL/s7bLz32Q/gspsCzMtY9ATDZbijvd1lPxiIMXSBGIPLEGM81mW3eTzF5Usv8H1ml3jkpHeHG0wuUCjHWde2q7yzTi3lI050Q1M/NG6R87s3N3V5ebpM1kZkRemV/s460ZKBiU21Mi4rkqcuzvnES7pXrReAy+AywGWAywCXAS6Dy9gQ78kXuqyzsb/Ll77CbmvWiSbG/hdl7HK8aPrakuaeyMcolkIZeCkcVlWJBa1az8AsiBM5GHPO4xYxRl450jdr+qZdcrfQ0LPf6x8K/Y0UwGXccfbmMsBlgMvg8htIZ8tlzznOlssAl7EkxnhMJHKmPkaM66Aodfa2iRO739zzeMmkWH2JxbCLKva8Kou5LyPrFk1aCGrksk5Nu1pR0qU4OkUXxRgQLwNcBrgMcBlcBrgMcBngMrgMcBmjeX7S8WElnmK2wVetKesHMUr7jZYjtpXiOdHc7D71ScvqRIy7ktoSKn4ZhKjt6ynzciy4ePLSS3rW7J0L5v5r+0SM8ZKbuXgZ4DLAZYDL4DLAZYDLAJfRihpceAove1fytWezitVSTpSQ6T5gdp9IzjIwfnfc/ES15ky0DEfbGUbOv7KqLuck52bdHtsKqeT8A7b1SZ51PD7xXSGfpv1G26H701eycuGKMZ4cZ8aNuze5jPc+4nH5wcMe2vtGl63M/eAJRYwBMQbAZYDL5Vjrkc9+fse4p8tx6eYeaLkMcPmSGMNEd8/7U624x4lVbrJ7V8W1bLr9j22u0Dj6WjRtP7JSzKyzHVh1ZuS8HFvuqr9T9nmWyqN/ZVOD5S97mPdQyGWAywCXAS6DywCXAS4DXAaXAS4P5715noerAI1cHugxfH1uLHPbukwRpdGtFWs5rmvSlE/0ayJWjOqg6F+wqX+pq1w9n/g77OHk0T9ymzkzL+0TLgNcBpdx82hno/s/l/Gqq4rL98TfDnAZ4DK4LOoDlwEuz5qL88omhvvI1Nvp14N64u8VjY7+//HSUdl0jKxt/kuTonjptSwSFb29mEdH6F90Kw+b+7vRyaiuHXYPl9uSVPKnL/4jP6dhNo5cMKNkk4nZuKuc0MbIMy1p+tZB0aW/PsjyQnjiZTfwxzxncxme/QAuA1wGuAwuA1wGuAxwGb9gjYe3dF13tYlvzjaaNo9xvRjDvjStG5e36uvW88DPhVT6R/eXPI+jT9ryJqKl4kn0r471GXddTFtUKsYNW61Pbh1j5KTrKBc0cdfFqq4ctmJLxMuvx99hg0FcpuzLYwwu37Pj+8tjukO8xWW3ZogxwGWAywCX4XH6ZS63ZUF4Vt2sibPq33zdoCP+cYyMzyc/kfHv1Wxi2gF/M4M1HKyxpEtrS7oKSLRXT4mxe/zvYWjJ6jpxt+ivcTJN3/5EmJj0QdMtuC0pZmRGT145UsU0pRXl0sTLEC8DXAa4DHAZXAa4DHD5FeRDjtFGcPnBxFMMMi+bl181L7/Z5Zw0PLngg+ebuWlnhXWZYV4GuAxwGeAyuAxwGeAywGVwGeAywGWAywCXwWWAywCXAS6DywCXAS4DXAa4DC4DXAa4DHAZXAa4DHAZ4DLAZXAZ4DLAZYDL4DLAZYDLAJcBLoPLAJcBLgNcBpcBLgNcBrgMLgMP4H9ePtje8VekfAAAAABJRU5ErkJggg==';
  const GACHA_MAX = 3;
  const GACHA_POOL = {
    ribuan: {
      more: [
        'Rp segini? Penjualnya nangis sesenggukan. Saran: berhenti, sebelum dia berubah pikiran.',
        'Ini bukan harga, ini sedekah. Penjualnya lagi menghapus air mata pakai struk transfer.',
        'Kucingnya nangis, penjualnya ikut nangis. Ambil sekarang, jangan digacha lagi — kasihan.',
        'Harga segini cuma cukup buat beli bensin penjualnya pulang. Berhenti di sini, kamu sudah menang banyak.'
      ],
      last: [
        'Harga final: recehan. Penjualnya resmi menangis di pojokan. Selamat, kamu kejam tapi beruntung.',
        'Segini harga finalmu. Penjualnya lagi mempertimbangkan pensiun dini.',
        'Final di ribuan. Kucingnya sudah habis air mata, penjualnya habis kata-kata.'
      ]
    },
    puluhan: {
      more: [
        'Harga bersahabat, penjualnya berbunga-bunga. Tapi ingat: gacha lagi, harga ini hangus.',
        'Segini pas di kantong, pas di hati penjualnya. Berhenti boleh, lanjut gacha juga boleh — tanggung sendiri.',
        'Kucingnya jatuh cinta, harganya juga ramah. Keputusan di tanganmu, jangan salahkan kucingnya.',
        'Puluhan ribu: tidak bikin penjual nangis, tidak bikin kamu miskin. Titik tengah yang damai.'
      ],
      last: [
        'Harga final yang bersahabat. Penjualnya senang, kamu senang, kucingnya apalagi.',
        'Final di puluhan ribu. Adil buat dua pihak — jarang-jarang gacha berakhir damai.',
        'Segini harga finalmu. Penjualnya tersenyum, dompetmu juga masih tersenyum.'
      ]
    },
    seratus: {
      more: [
        'Pas seratus ribu. Penjualnya masih berbunga-bunga, tapi dompetmu mulai melirik. Gacha lagi = harga ini hangus.',
        'Seratus ribu pas, angka cantik. Berhenti di angka cantik, atau berjudi lagi?'
      ],
      last: [
        'Final: seratus ribu pas. Angka bulat, hati penjual bulat senang.',
        'Harga finalmu seratus ribu. Tidak murah, tidak mahal — pas kayak baju lebaran.'
      ]
    },
    ratusan: {
      more: [
        'Penjualnya tertawa lepas dan akan berangkat umroh. Gacha lagi bisa menyelamatkan — atau memperparah.',
        'Ratusan ribu! Penjualnya ketawa sambil packing koper umroh. Masih ada kesempatan kabur dari harga ini.',
        'Kucingnya tertawa, penjualnya tertawa, yang menangis tinggal dompetmu. Gacha lagi kalau berani.',
        'Harga segini penjualnya bisa traktir satu RT. Selamatkan dirimu selagi kesempatan masih ada.'
      ],
      last: [
        'Penjualnya tertawa lepas dan akan berangkat umroh — dompetmu yang menangis. Ini harga finalmu.',
        'Final di ratusan ribu. Penjualnya sudah booking tiket umroh, kamu booking jadwal transfer.',
        'Harga finalmu bikin penjualnya ketawa sampai lupa cara nangis. Terima kasih, donatur.'
      ]
    },
    sejuta: {
      more: [
        'SEJUTA. Harga normalnya cuma Rp 50 ribu — kamu baru membayar 20 kali lipat. Penjualnya lagi bikin syukuran kecil-kecilan. Masih bisa diselamatkan, kalau berani.',
        'Penjualnya resmi naik kelas jadi sultan berkat kamu. Satu gacha lagi, atau ikhlas jadi donatur?',
        'Sejuta pas. Denda di bea cukai saja jarang semahal ini. Gacha lagi — selamatkan dompet dan harga dirimu.',
        'Kucingnya ketawa sampai keselek. Penjualnya sudah cek tiket umroh sekalian oleh-olehnya. Kesempatanmu masih ada — pakai.'
      ],
      last: [
        'Final: sejuta. Harga aslinya 50 ribu. Kamu bukan pembeli, kamu investor tanpa saham. Penjualnya berangkat umroh sambil mendoakanmu.',
        'Sejuta, final, sah. Penjualnya akan menceritakan hari ini ke cucu-cucunya. Terima kasih, orang baik.',
        'Harga finalmu cukup buat penjualnya umroh plus beli koper baru. Kucingnya ikut — kamu di rumah saja, jaga lilin keberuntungan.'
      ]
    }
  };
  const rpFmt = (n) => 'Rp ' + String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const gachaLoad = () => {
    try {
      const a = JSON.parse(String(GM_getValue('ck_gacha', '[]') || '[]'));
      return Array.isArray(a) ? a.filter((v) => typeof v === 'number' && v > 0).slice(0, GACHA_MAX) : [];
    } catch (e) { return []; }
  };
  const gachaSave = (a) => { try { GM_setValue('ck_gacha', JSON.stringify(a)); } catch (e) { /* abaikan */ } };
  const gachaRoll = () => {
    const r = Math.random() * 100;
    let lo, hi;
    if (r < 20) { lo = 1; hi = 9; }
    else if (r < 60) { lo = 10; hi = 49; }
    else if (r < 80) { lo = 50; hi = 99; }
    else if (r < 92) { lo = 100; hi = 499; }
    else if (r < 98) { lo = 500; hi = 999; }
    else { lo = 1000; hi = 1000; }
    return (lo + Math.floor(Math.random() * (hi - lo + 1))) * 1000;
  };
  const gachaBand = (v) => {
    if (v <= 9000) { return { key: 'ribuan', name: 'RIBUAN!', emoji: '😿' }; }
    if (v === 100000) { return { key: 'seratus', name: 'SERATUS RIBU!', emoji: '😻' }; }
    if (v <= 99000) { return { key: 'puluhan', name: 'PULUHAN RIBU!', emoji: '😻' }; }
    if (v < 1000000) { return { key: 'ratusan', name: 'RATUSAN RIBU!', emoji: '😹' }; }
    return { key: 'sejuta', name: 'SATU JUTA!', emoji: '😹' };
  };
  const gachaLine = (v, last) => {
    const pool = GACHA_POOL[gachaBand(v).key][last ? 'last' : 'more'];
    return pool[Math.floor(Math.random() * pool.length)];
  };

  function wireGacha(box) {
    const $$ = (s) => box.querySelector(s);
    const bigPrice = $$('#ckBigPrice'), priceNote = $$('#ckPriceNote'), btn = $$('#ckGachaBtn');
    const dots = $$('#ckDots'), chanceTxt = $$('#ckChanceTxt'), history = $$('#ckHistory');
    const pop = $$('#ckPop'), popCard = $$('#ckPopCard'), popEmoji = $$('#ckPopEmoji');
    const popBand = $$('#ckPopBand'), popVal = $$('#ckPopVal'), popSub = $$('#ckPopSub');
    if (!bigPrice || !btn || !pop) { return; }
    if (!document.getElementById('ckGachaStyle')) {
      const stl = document.createElement('style');
      stl.id = 'ckGachaStyle';
      stl.textContent = '@keyframes ckPopIn{0%{transform:scale(.5) rotate(-4deg);opacity:0}60%{transform:scale(1.08) rotate(1deg);opacity:1}100%{transform:scale(1) rotate(0deg);opacity:1}}.ckPopAnim{animation:ckPopIn .45s cubic-bezier(.2,.9,.3,1.3) both}';
      (document.head || document.documentElement).appendChild(stl);
    }
    $$('#ckGachaHelp').onclick = () => { const rb = $$('#ckRateBox'); rb.style.display = rb.style.display === 'none' ? 'block' : 'none'; };
    let draws = gachaLoad();
    let rolling = false;
    const current = () => (draws.length ? draws[draws.length - 1] : null);
    function render() {
      const left = GACHA_MAX - draws.length;
      dots.textContent = '●'.repeat(left) + '○'.repeat(GACHA_MAX - left);
      chanceTxt.textContent = left + 'x';
      history.innerHTML = '';
      draws.forEach((v, i) => {
        const isLast = i === draws.length - 1;
        const s = document.createElement('span');
        s.textContent = 'Gacha ' + (i + 1) + ': ' + rpFmt(v) + (isLast ? ' (berlaku)' : '');
        s.style.cssText = 'font-size:11px;padding:3px 8px;border-radius:999px;border:1px solid #ccc;background:#fff;color:#555;';
        if (!isLast) { s.style.opacity = '.55'; s.style.textDecoration = 'line-through'; }
        if (isLast) { s.style.borderColor = '#1fae4b'; s.style.color = '#1fae4b'; s.style.fontWeight = '700'; }
        history.appendChild(s);
      });
      if (rolling) { btn.textContent = 'MENGACAK…'; btn.style.opacity = '.6'; btn.style.pointerEvents = 'none'; return; }
      btn.style.opacity = '1'; btn.style.pointerEvents = 'auto';
      if (!draws.length) {
        bigPrice.textContent = 'Rp ••••••';
        priceNote.textContent = 'Klik tombol dibawah untuk generate harga keberuntunganmu.';
        btn.textContent = 'GACHA HARGAMU'; btn.style.display = 'block';
      } else if (left > 0) {
        bigPrice.textContent = rpFmt(current());
        priceNote.textContent = 'Ini hargamu sekarang. Berhenti = transfer sebesar ini. Gacha lagi = harga ini hangus, diganti hasil baru.';
        btn.textContent = 'GACHA LAGI'; btn.style.display = 'block';
      } else {
        bigPrice.textContent = rpFmt(current());
        priceNote.textContent = 'Kesempatan habis. Harga finalmu terkunci — transfer sebesar itu, key langsung jadi milikmu.';
        btn.style.display = 'none';
      }
    }
    // Animasi slot: semua digit berputar, lalu berhenti satu-satu mulai dari
    // digit PALING BELAKANG (kanan) maju ke kiri — digit terbesar terbuka
    // terakhir (perintah user 23 Nov 2026, menggantikan arah demo).
    function spinTo(target, done) {
      const finalStr = rpFmt(target);
      const chars = finalStr.split('');
      const digitPos = [];
      chars.forEach((c, i) => { if (/\d/.test(c)) { digitPos.push(i); } });
      const locked = new Set();
      const SPIN = 450, STEP = 120;
      const lockAt = (k) => SPIN + (digitPos.length - 1 - k) * STEP;
      const t0 = Date.now();
      const iv = setInterval(() => {
        const el = Date.now() - t0;
        for (let k = 0; k < digitPos.length; k++) { if (el >= lockAt(k)) { locked.add(digitPos[k]); } }
        bigPrice.textContent = chars.map((c, i) => (/\d/.test(c) && !locked.has(i) ? String(Math.floor(Math.random() * 10)) : c)).join('');
        if (locked.size >= digitPos.length) { clearInterval(iv); bigPrice.textContent = finalStr; done(); }
      }, 50);
    }
    function showPop(v, leftAfter) {
      const b = gachaBand(v);
      popEmoji.textContent = b.emoji;
      popBand.textContent = b.name;
      popVal.textContent = rpFmt(v);
      popSub.textContent = gachaLine(v, leftAfter <= 0);
      pop.style.display = 'flex';
      popCard.classList.remove('ckPopAnim'); void popCard.offsetWidth; popCard.classList.add('ckPopAnim');
    }
    pop.onclick = () => { pop.style.display = 'none'; }; // pop-up TIDAK hilang otomatis; user menutup sendiri
    btn.onclick = () => {
      if (rolling || draws.length >= GACHA_MAX) { return; }
      rolling = true; render();
      const target = gachaRoll();
      spinTo(target, () => {
        rolling = false;
        draws = draws.concat([target]);
        gachaSave(draws);
        render();
        showPop(target, GACHA_MAX - draws.length);
      });
    };
    render();
  }

  function showLock(why) {
    const old = document.getElementById('ckLockBox'); if (old) { old.remove(); }
    const p = st.pay || {};
    const full = st.mode === 'full';
    const uname = myUser();
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
        <div style="background:#f6f8fa;border:1px solid #e1e4e8;border-radius:8px;padding:10px 12px;margin-bottom:12px;position:relative;">
          <div style="display:flex;align-items:center;justify-content:space-between;">
            <b>Cara beli</b>
            <button id="ckGachaHelp" title="Lihat drop rate" style="width:22px;height:22px;border-radius:50%;border:1px solid #ccc;background:#fff;color:#1fae4b;font-weight:800;font-size:13px;cursor:pointer;line-height:1;padding:0;">?</button>
          </div>
          <div id="ckRateBox" style="display:none;margin-top:8px;font-size:12px;line-height:1.65;border-top:1px dashed #ccc;padding-top:8px;">
            <div style="font-weight:700;margin-bottom:2px;">Drop rate gacha</div>
            <div>Ribuan (Rp 1–9 rb) — <b>20%</b></div>
            <div>Puluhan ribu (Rp 10–49 rb) — <b>40%</b></div>
            <div>Puluhan ribu atas (Rp 50–99 rb) — <b>20%</b></div>
            <div>Ratusan ribu (Rp 100–499 rb) — <b>12%</b></div>
            <div>Ratusan ribu atas (Rp 500–999 rb) — <b>6%</b></div>
            <div>Pas Rp 1 juta — <b>2%</b></div>
            <div style="color:#555;margin-top:4px;">Yang berlaku: hasil gacha terakhirmu. Gacha lagi = harga lama hangus.</div>
          </div>
          <div style="font-size:13px;color:#555;margin-top:8px;">Harga:</div>
          <div id="ckBigPrice" style="font-size:30px;font-weight:800;letter-spacing:1px;margin:2px 0 4px;">Rp ••••••</div>
          <div id="ckPriceNote" style="font-size:12px;color:#555;margin-bottom:10px;">Klik tombol dibawah untuk generate harga keberuntunganmu.</div>
          <button id="ckGachaBtn" style="background:#1fae4b;color:#fff;font-weight:700;border:0;padding:9px 16px;border-radius:6px;cursor:pointer;font-size:14px;width:100%;">GACHA HARGAMU</button>
          <div style="display:flex;align-items:center;gap:6px;margin-top:10px;font-size:12px;color:#555;">
            <span>Kesempatan gacha:</span><span id="ckDots" style="letter-spacing:3px;font-size:14px;color:#1fae4b;">●●●</span><span id="ckChanceTxt">3x</span>
          </div>
          <div id="ckHistory" style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;"></div>
          <div style="border-top:1px dashed #ccc;margin-top:12px;padding-top:10px;text-align:center;">
            <div style="font-size:13px;">Bayar pakai <b>QRIS</b> — semua e-wallet &amp; m-banking:</div>
            <img src="${QRIS_B64}" alt="QRIS Hikman, Service" style="width:190px;height:190px;margin:8px auto;display:block;image-rendering:pixelated;background:#fff;">
            <div style="font-size:13px;">a.n. <b>Hikman, Service</b></div>
            <div style="font-size:12px;color:#555;margin-top:4px;">QRIS ini tidak berisi nominal — masukkan nominal <b>sesuai hasil gacha terakhir</b> di aplikasi pembayarmu.</div>
            <div style="font-size:13px;margin-top:8px;">Kirim bukti pembayaran ke CP: <b>${cpHtml(p.cp || '')}</b></div>
          </div>
          <div id="ckPop" style="display:none;position:absolute;inset:0;background:rgba(0,0,0,.45);align-items:center;justify-content:center;z-index:5;border-radius:8px;">
            <div id="ckPopCard" style="background:#fff;border:2px solid #1fae4b;border-radius:14px;padding:14px 22px 18px;text-align:center;max-width:84%;box-shadow:0 10px 40px rgba(0,0,0,.4);">
              <div id="ckPopEmoji" style="font-size:64px;line-height:1.1;">😿</div>
              <div id="ckPopBand" style="font-size:24px;font-weight:900;letter-spacing:1px;color:#1fae4b;">RIBUAN!</div>
              <div id="ckPopVal" style="font-size:19px;font-weight:800;margin-top:2px;">Rp 5.000</div>
              <div id="ckPopSub" style="font-size:12px;color:#555;margin-top:5px;"></div>
              <div style="font-size:11px;color:#888;margin-top:9px;">Klik untuk menutup</div>
            </div>
          </div>
        </div>`}
        <div style="margin-bottom:10px;">Akun CEISA:<br>
          <div style="display:flex;gap:6px;"><input id="ckUname" readonly value="${esc(uname)}" placeholder="belum terbaca — buka tabel Dokumen Pabean dulu" style="flex:1;padding:5px;font:12px monospace;user-select:all"></div>
        </div>
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
    if (!full) { wireGacha(box); }
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
