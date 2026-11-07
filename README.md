# CEISA 4.0 Tool Downloader

Tool untuk mengunduh dokumen pabean di portal **CEISA 4.0** secara massal:

- **PDF** : respon SPPB dan Detail Dokumen, per rentang Nomor Pengajuan
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

- **Gratis 3x klik Mulai** per perangkat (gabungan PDF dan Excel)
- Setelah jatah habis, tool meminta **license key**
- Cara mendapatkan key: ikuti petunjuk pembayaran yang tampil di panel tool, lalu hubungi admin dengan menyertakan **Device ID** yang tampil di panel
- Satu key berlaku untuk **satu perangkat**
- Masukkan key di panel, klik aktivasi

---

## Masalah Umum

**Panel tidak muncul**
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
