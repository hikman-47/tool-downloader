# CEISA 4.0 Tool Downloader

Tool untuk mengunduh dokumen pabean di portal **CEISA 4.0** secara massal:

- **PDF** : respon SPPB dan Detail Dokumen, per rentang Nomor Pengajuan (termasuk BC 2.7 In)
- **Excel** : ekspor dokumen BC per jenis dokumen dan rentang tanggal

Panel tool muncul otomatis di halaman **Dokumen Pabean** CEISA.

---

## Cara Install (2 langkah)

### Langkah 1. Pasang Tampermonkey (sekali saja)

Tampermonkey adalah ekstensi browser yang menjalankan tool ini. Klik link sesuai browser, lalu klik **Add / Tambahkan** (atau **Add to Chrome**):

- Chrome / Brave: [Install Tampermonkey untuk Chrome](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
- Edge: [Install Tampermonkey untuk Edge](https://microsoftedge.microsoft.com/addons/detail/tampermonkey/iikmkjmpaadaobahmlepeloendndfphd)
- Firefox: [Install Tampermonkey untuk Firefox](https://addons.mozilla.org/firefox/addon/tampermonkey/)
- Browser lain: [tampermonkey.net](https://www.tampermonkey.net/)

**Penting untuk Chrome, Edge, Opera (dan browser Chromium lain): aktifkan izin script**

Setelah Tampermonkey terpasang, browser belum mengizinkannya menjalankan script. Tanpa langkah ini tool terinstal tapi **tidak berjalan** (panel tidak muncul).

1. Klik kanan ikon Tampermonkey, lalu pilih **Manage Extension / Kelola ekstensi** (atau buka `chrome://extensions` / `edge://extensions`, klik **Details** pada Tampermonkey)
2. Nyalakan toggle **Allow User Scripts**
3. Jika toggle itu tidak ada (browser versi lama), nyalakan **Developer mode** di pojok kanan atas halaman ekstensi

Firefox tidak perlu langkah ini.

> Sudah punya Tampermonkey? Lewati ke langkah 2.

### Langkah 2. Pasang Tool

Buka link di bawah ini di browser yang sama:

**[>> INSTALL TOOL <<](https://raw.githubusercontent.com/hikman-47/tool-downloader/main/ceisa-downloader.user.js)**

Halaman Tampermonkey akan terbuka. Klik **Install**. Selesai.

Tool akan **update otomatis** lewat Tampermonkey, tidak perlu install ulang.

---

## Cara Pakai

1. Buka **portal.beacukai.go.id** dan login ke CEISA 4.0
2. Masuk ke menu **Dokumen Pabean**
3. Panel tool muncul di halaman tersebut (tidak muncul di halaman login)
4. Pilih fitur **PDF** atau **Excel**
5. Isi pengaturan:
   - PDF: jenis dokumen (mis. BC 4.1) dan rentang Nomor Pengajuan
   - Excel: jenis dokumen dan rentang tanggal
6. Klik **Mulai**, biarkan proses berjalan sampai selesai
7. File otomatis terunduh ke folder Download

### Format nama file

- PDF: `[Jenis BC] [NOMOR REGISTER][S/i] [tanggal].pdf`
  Contoh: `BC41 018746S 28-10-2026.pdf` (S = SPPB, i = Detail Dokumen)
- Excel: `[Jenis BC]_[tanggal awal]_sd_[tanggal akhir].xlsx`

---

## Jatah Gratis dan License Key

- Gratis **3x pakai per akun CEISA**, gabungan PDF dan Excel (1x proses unduhan yang jalan = 1x pakai). Hitungannya nempel ke akun CEISA-mu, bukan ke perangkat.
- Sesudah jatah habis, panel akan meminta **license key**.
- Satu key berlaku untuk **satu akun CEISA** — jadi per orang (eximer), bukan per perusahaan dan bukan per perangkat.
- Pindah perusahaan? Key bisa dilepas dari akun lama lewat tombol **Lepaskan** di panel, lalu diaktifkan lagi di akun barumu.

### Harga key: gacha

Harga key tidak paten. Di panel ada tombol **GACHA HARGAMU**:

1. Klik tombolnya, angka harga berputar seperti slot lalu berhenti di satu nominal.
2. Kesempatanmu **3x**. Yang berlaku adalah **hasil gacha terakhir** — kalau kamu gacha lagi, harga sebelumnya hangus dan diganti hasil baru.
3. Mau tahu peluang tiap kisaran harga? Klik ikon **?** di panel, di situ ada drop rate-nya.

### Bayar: QRIS

Pembayaran sekarang lewat **QRIS**, tidak lagi transfer ke rekening:

1. Scan QR QRIS yang tampil di panel (a.n. **Hikman, Service**) pakai e-wallet atau m-banking apa saja.
2. QR ini tidak berisi nominal. Ketik sendiri nominalnya di aplikasimu, persis sebesar **hasil gacha terakhirmu**.
3. Kirim bukti pembayaran ke nomor CP yang tampil di panel.
4. Kamu akan dikirimi license key. Tempel key-nya di panel, lalu klik **Aktifkan**.
5. Saat aktivasi, pastikan baris **Akun CEISA** di panel sudah menampilkan username-mu (buka dulu tabel Dokumen Pabean sampai daftarnya tampil kalau masih kosong), supaya key nempel ke akun yang benar.

---

## Masalah Umum

**Panel tidak muncul**
- Pastikan toggle **Allow User Scripts** (atau **Developer mode**) sudah menyala di pengaturan ekstensi Tampermonkey
- Pastikan Tampermonkey aktif dan tool dalam keadaan ON di dashboard Tampermonkey
- Pastikan berada di halaman Dokumen Pabean, bukan halaman login
- Refresh halaman (Ctrl + F5)

**Halaman Install tidak muncul saat klik link**
- Pastikan Tampermonkey sudah terpasang di browser yang sama
- Alternatif: buka `ceisa-downloader.user.js` di repo ini, klik **Raw**, lalu klik **Install** di Tampermonkey

**Unduhan diblokir browser**
- Izinkan **multiple downloads** untuk situs portal.beacukai.go.id di pengaturan browser

---

## Catatan

- Tool ini tidak resmi dan tidak berafiliasi dengan DJBC
- Pemakaian menjadi tanggung jawab pengguna
- Pembuat: Hikman
