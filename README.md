# CEISA 4.0 Tool Downloader

Tool untuk mengunduh dokumen pabean di portal **CEISA 4.0** secara massal:

- **PDF** : respon SPPB dan Detail Dokumen, per rentang Nomor Pengajuan (BC 2.7 In dibaca dari daftar dashboard)
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

Panel bisa digeser ke posisi lain, dan akan mengecil sendiri beberapa detik setelah kursor meninggalkan area panel.
4. Pilih fitur **PDF** atau **Excel**
5. Isi pengaturan:
   - PDF: jenis dokumen (mis. BC 4.1) dan rentang Nomor Pengajuan
   - Excel: jenis dokumen dan rentang tanggal
6. Klik **Mulai**, biarkan proses berjalan sampai selesai
7. File otomatis terunduh ke folder Download

### Selama unduhan berjalan

- Biarkan tab tempat tool berjalan apa adanya: jangan ditutup, jangan dimuat ulang, dan jangan dipakai membuka halaman lain. Halaman berganti = unduhan berhenti.
- Kalau mau tetap kerja di CEISA (misalnya membuat dokumen), buka **tab baru** dengan login yang sama. Unduhan di tab lama tetap jalan.
- Jangan menjalankan tool di dua tab sekaligus.
- Tarikan yang berjalan di tab belakang bisa sedikit lebih lambat, itu normal.
- Bila yang akan diunduh lebih dari 300 dokumen, muncul pertanyaan "Lanjut unduh?". Klik **OK** dulu sebelum pindah tab — pertanyaan yang tidak dijawab OK dianggap batal dan unduhan berhenti.

### Khusus BC 2.7 In

Daftar dokumen masuk dibaca dari tabel dashboard CEISA, jadi sebelum memulai, buka dashboard dan saring **Jenis Dokumen** ke **BC 2.7** agar scan berjalan cepat. Untuk perusahaan grup, isi tabel izin TPB tujuan (satu baris per unit: `NAMA = nomor izin`) agar hanya dokumen yang bertujuan ke unit Anda yang diunduh. Tabel boleh dikosongkan bila tidak memakai penyaringan unit.

### Format nama file

- PDF: `[Jenis BC] [NOMOR REGISTER][S/i] [tanggal].pdf`
  Contoh: `BC41 018746S 28-10-2026.pdf` (S = SPPB, i = Detail Dokumen)
- Excel: `[Jenis BC]_[tanggal awal]_sd_[tanggal akhir].xlsx`

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
