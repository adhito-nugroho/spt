# Spesifikasi Desain: Pengaturan Tata Letak Per Blok Preview Surat Tugas

**Tanggal:** 26 Oktober 2026  
**Status:** Disetujui  
**Target Modul:** `modules/generate-surat.php`, `modules/save-template-layout.php`

---

## 1. Latar Belakang & Tujuan
Pada aplikasi Surat Perintah Tugas (SPT), pengguna membutuhkan fleksibilitas penuh untuk mengatur jarak dan kerapatan setiap blok dokumen secara mandiri di halaman *Preview Surat Tugas*. Tujuannya adalah agar dokumen dapat disesuaikan dengan presisi tinggi, muat pas dalam 1 halaman (atau jumlah halaman yang dikehendaki), serta tata letak tetap rapi dan proporsional saat dicetak maupun diunduh dalam format PDF dan Word (.docx).

Pengaturan yang sebelumnya hanya berupa spasi global (`section_spacing`) kini diperinci menjadi pengaturan per blok spesifik:
1. Blok Kop Surat
2. Judul Surat (Tulisan "SURAT TUGAS" dan "Nomor : ...")
3. Blok Dasar
4. Blok Tulisan "MEMERINTAHKAN :"
5. Blok Pegawai (termasuk pengaturan jarak antar pegawai)
6. Blok Untuk
7. Blok Tanda Tangan (Tempat/Tanggal, Jabatan, Ruang Tanda Tangan & Stempel, serta Nama & NIP Penandatangan)

---

## 2. Skema Parameter Layout & Batasan (Ranges)

Semua nilai spasi menggunakan satuan piksel (`px`), margin halaman menggunakan milimeter (`mm`), ukuran huruf menggunakan point (`pt`), dan jarak antar baris menggunakan faktor pengali (`line_height`).

| Kategori Blok | Nama Kunci Parameter | Rentang Nilai | Step | Nilai Default | Keterangan |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Halaman & Huruf** | `margin_top` | 5 – 40 mm | 0.5 | 15 mm | Margin atas halaman |
| | `margin_bottom` | 5 – 40 mm | 0.5 | 15 mm | Margin bawah halaman |
| | `margin_left` | 10 – 40 mm | 0.5 | 20 mm | Margin kiri halaman |
| | `margin_right` | 10 – 40 mm | 0.5 | 20 mm | Margin kanan halaman |
| | `font_size` | 9 – 13 pt | 0.5 | 11 pt | Ukuran huruf global |
| | `line_height` | 1.0 – 2.0 | 0.05 | 1.0 | Jarak antar baris teks |
| **Kop Surat** | `kop_bottom_spacing` | 0 – 50 px | 1 | 10 px | Jarak bawah kop sebelum judul |
| **Judul Surat** | `judul_top_spacing` | -15 – 40 px | 1 | 0 px | Jarak sebelum "SURAT TUGAS" |
| | `nomor_surat_spacing` | 0 – 30 px | 1 | 4 px | Jarak antara "SURAT TUGAS" & "Nomor" |
| | `judul_bottom_spacing` | 0 – 40 px | 1 | 10 px | Jarak setelah nomor sebelum DASAR |
| **Blok Dasar** | `dasar_top_spacing` | 0 – 30 px | 1 | 6 px | Jarak sebelum tulisan "DASAR :" |
| | `dasar_item_spacing` | 0 – 20 px | 1 | 3 px | Kerapatan antar butir dasar (1, 2, 3...) |
| | `dasar_bottom_spacing`| 0 – 30 px | 1 | 8 px | Jarak setelah butir dasar |
| **Blok Memerintahkan**| `memerintahkan_spacing`| 0 – 30 px | 1 | 8 px | Jarak sebelum & sesudah "MEMERINTAHKAN :" |
| **Blok Pegawai** | `pegawai_top_spacing` | 0 – 30 px | 1 | 6 px | Jarak sebelum teks "KEPADA :" |
| | `pegawai_gap_spacing` | 0 – 35 px | 1 | 8 px | Jarak antar pegawai (gap pegawai 1 ke 2) |
| | `pegawai_line_spacing`| 0 – 15 px | 1 | 2 px | Kerapatan baris data dalam 1 pegawai |
| **Blok Untuk** | `untuk_top_spacing` | 0 – 30 px | 1 | 6 px | Jarak sebelum tulisan "UNTUK :" |
| | `untuk_item_spacing` | 0 – 20 px | 1 | 3 px | Kerapatan antar butir tugas (1, 2, 3...) |
| | `untuk_bottom_spacing`| 0 – 30 px | 1 | 10 px | Jarak sesudah tugas sebelum blok TTD |
| **Blok Tanda Tangan** | `ttd_offset` | -20 – 80 px | 1 | 10 px | Jarak atas keseluruhan blok TTD |
| | `ttd_x` | -50 – 50 px | 1 | 0 px | Pergeseran horizontal posisi TTD |
| | `ttd_date_to_jabatan` | 0 – 25 px | 1 | 4 px | Jarak tempat & tanggal ke jabatan |
| | `ttd_space_height` | 20 – 120 px | 2 | 60 px | Tinggi ruang tanda tangan & stempel |
| | `ttd_name_nip_spacing`| 0 – 25 px | 1 | 4 px | Jarak dari TTD ke nama & NIP |
| | `lock_ttd` | 0 atau 1 | 1 | 0 | Cegah blok TTD terpotong antar halaman |

---

## 3. Desain Antarmuka Pengguna (UI)

Di panel kontrol samping kiri (`lg:col-span-4` pada `generate-surat.php`):
1. **Bagian Atas**:
   - Preset tombol **Paksa Muat 1 Halaman**: Menerapkan skala kompresi proporsional pada margin, font, dan semua parameter spasi blok.
   - Tombol **Reset ke Default**: Mengembalikan semua nilai ke default template yang tersimpan atau standar sistem.
   - Status indikator konfigurasi (Default tersimpan / Kustom aktif).

2. **Pengelompokan Akordion (Collapsible Panels)**:
   - **Akordion 1: Margin & Huruf Global**: Margin 4 sisi, Ukuran Font, Line Height.
   - **Akordion 2: Kop & Judul Surat**: Spasi bawah kop, spasi atas judul, jarak ke nomor, spasi bawah nomor.
   - **Akordion 3: Dasar & Memerintahkan**: Spasi atas dasar, spasi butir dasar, spasi bawah dasar, spasi tulisan memerintahkan.
   - **Akordion 4: Blok Pegawai**: Spasi atas "KEPADA :", jarak antar pegawai, spasi baris identitas pegawai.
   - **Akordion 5: Blok Untuk**: Spasi atas "UNTUK :", kerapatan butir tugas, spasi bawah tugas.
   - **Akordion 6: Blok Tanda Tangan**: Jarak atas TTD, pergeseran horizontal TTD, jarak tanggal ke jabatan, tinggi ruang TTD & stempel, jarak ke nama/NIP, checkbox kunci blok TTD.

3. **Interaktivitas**:
   - Setiap kontrol memiliki label, slider (`input[type=range]`), nilai numerik satuan, dan tombol presisi `−` / `+`.
   - Perubahan slider memicu debounce rendering PDF di panel kanan dengan indikator animasi memuat (*spinner*).
   - Tombol **Terapkan Perubahan** dan **Simpan sebagai Default Template** di bagian footer panel.

---

## 4. Logika Pemrosesan Dokumen XML & Konversi PDF

1. **Manipulasi DOCX (`spt_layout_apply_to_docx`)**:
   - Membaca dan mem-parsing `word/document.xml` dengan `DOMDocument` dan `DOMXPath`.
   - Mengatur margin seksi pada node `w:pgMar`.
   - Menemukan paragraf target berdasarkan pola teks dan penandaan semantis:
     - Kop Surat: paragraf sebelum `SURAT TUGAS`.
     - Judul: paragraf berteks `SURAT TUGAS`.
     - Nomor: paragraf berteks `Nomor :`.
     - Dasar: paragraf `DASAR :` dan butir nomor `1.`, `2.`, dst.
     - Memerintahkan: paragraf `MEMERINTAHKAN :`.
     - Pegawai: paragraf `KEPADA :`, penanda nama pegawai berikutnya (Pegawai 2, 3...), dan baris NIP/Pangkat/Jabatan. Untuk tabel lampiran, diaplikasikan pada baris tabel (row height / cell padding).
     - Untuk: paragraf `UNTUK :` dan butir tugas `1.`, `2.`, dst.
     - Tanda Tangan: baris `Bojonegoro [tanggal]`, baris jabatan, baris spasi/gambar TTD (`w:drawing`), baris nama penandatangan, dan baris NIP.
   - Menerapkan spasi pada properti paragraf `w:pPr -> w:spacing` menggunakan fungsi helper konversi piksel ke twips (`spt_layout_px_to_twips`).

2. **Dukungan Dompdf Fallback (`spt_layout_build_css` & `spt_layout_inject_html`)**:
   - Menyuntikkan class khusus per blok ke elemen HTML hasil konversi PHPWord:
     - `.spt-kop`, `.spt-title`, `.spt-nomor`, `.spt-dasar`, `.spt-memerintahkan`, `.spt-pegawai`, `.spt-untuk`, `.spt-ttd`, dll.
   - Menghasilkan aturan CSS yang sesuai dengan nilai parameter slider.

3. **Penyimpanan Template (`save-template-layout.php`)**:
   - Memvalidasi seluruh kunci baru sesuai rentang yang ditentukan di tabel skema.
   - Menyimpan seluruh payload JSON ke kolom `layout_settings` di tabel `templates`.

---

## 5. Rencana Pengujian & Verifikasi
1. **Pengujian UI**:
   - Memastikan semua akordion dapat dibuka dan ditutup dengan mulus.
   - Memastikan tombol `−` dan `+` menambah/mengurangi nilai sesuai step.
   - Memastikan tombol *Paksa Muat 1 Halaman* dan *Reset ke Default* berfungsi tepat.
2. **Pengujian Render PDF**:
   - Menguji surat tugas 1 pegawai.
   - Menguji surat tugas multi-pegawai (badan surat dan lampiran).
   - Memastikan spasi per blok berubah secara visual sesuai nilai slider.
   - Memastikan unduh Word (.docx) dan unduh PDF menghasilkan dokumen dengan tata letak yang identik.
3. **Pengujian Penyimpanan**:
   - Memastikan tombol *Simpan sebagai Default Template* berhasil menyimpan konfigurasi ke database dan otomatis teraplikasi saat membuka surat lain dengan template yang sama.
