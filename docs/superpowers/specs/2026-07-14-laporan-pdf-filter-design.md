# Design Spec: Generate PDF Laporan Semua Data Berdasarkan Filter

## Pendahuluan
Fitur "Cetak Laporan" saat ini menggunakan fungsi JavaScript `window.print()` langsung pada halaman aktif. Karena halaman laporan umum (`laporan.php`) dan laporan per pegawai (`laporan-pegawai.php`) menggunakan paginasi (membatasi data per halaman), maka hasil cetak hanya mencakup data terbatas yang ada di layar saat itu. 

Dokumen ini mendefinisikan desain untuk mengubah perilaku tombol "Cetak Laporan" menjadi ekspor PDF secara dinamis dari server menggunakan library Dompdf, agar seluruh data hasil pencarian/filter dapat tercetak secara utuh.

## Alur Kerja Terpilih
1. Pengguna memfilter laporan berdasarkan kriteria tanggal atau pegawai (pada modul Laporan Per Pegawai).
2. Pengguna mengeklik tombol **Cetak Laporan**.
3. Sistem membuka tab baru dan mengirim permintaan ke server dengan parameter query `export=pdf` beserta parameter filter yang aktif.
4. Server mendeteksi parameter `export=pdf`:
   * Mengambil semua data dari database yang sesuai dengan filter (mengabaikan paginasi).
   * Me-render data tersebut ke dalam struktur HTML yang bersih dan terformat khusus untuk konsumsi cetak.
   * Menginisialisasi library Dompdf, mengubah HTML menjadi PDF, dan mengalirkannya (stream) langsung ke browser dengan header PDF yang tepat agar siap diunduh atau dicetak.

---

## Spesifikasi Perubahan Modul

### 1. Laporan Rekapitulasi Surat Tugas (`laporan.php`)
* **Pemicu**: Tombol "Cetak Laporan" diubah dari `<button onclick="window.print()">` menjadi `<a href="?start_date=...&end_date=...&export=pdf" target="_blank">`.
* **Kueri Data**: Jika `export=pdf`, jalankan kueri SQL untuk mengambil seluruh data surat tugas pada periode tersebut tanpa klausa `LIMIT`.
* **Renderer**: Menggunakan Dompdf dengan orientasi **Landscape** A4.

### 2. Laporan Surat Tugas Per Pegawai (`laporan-pegawai.php`)
* **Pemicu**: Tombol "Cetak Laporan" diubah dari `<button onclick="window.print()">` menjadi `<a href="?start_date=...&end_date=...&nip=...&export=pdf" target="_blank">`.
* **Kueri Data**: Jika `export=pdf`, ambil seluruh pegawai beserta detail surat tugasnya yang masuk filter periode tanggal dan NIP (tanpa paginasi).
* **Renderer**: Menggunakan Dompdf dengan orientasi **Portrait** A4.

---

## Rencana Verifikasi

### Pengujian Manual
1. Membuka menu Laporan.
2. Memilih range tanggal filter yang menghasilkan lebih dari 10 data (sehingga terjadi paginasi).
3. Mengeklik tombol **Cetak Laporan**.
4. Memverifikasi bahwa dokumen PDF yang terunduh/tampil di tab baru memiliki daftar seluruh data surat tugas (bukan hanya 10 baris pertama), lengkap dengan header laporan dan tanda tangan di bagian bawah.
5. Melakukan hal serupa pada Laporan Per Pegawai dan memverifikasi hasilnya.
