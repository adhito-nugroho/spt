# Laporan PDF Filter Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Mengubah fungsi tombol "Cetak Laporan" pada menu Laporan dan Laporan Per Pegawai untuk menghasilkan berkas PDF utuh berisi semua data sesuai filter yang dipilih, menggunakan Dompdf di sisi server.

**Architecture:** Mengirim request cetak ke halaman yang sama dengan parameter query `export=pdf`. Server memproses data utuh (tanpa batasan paginasi), me-render halaman HTML minimalis khusus print, dan menyalurkannya (stream) sebagai dokumen PDF langsung di browser melalui Dompdf.

**Tech Stack:** PHP, PDO (MySQL), Dompdf.

## Global Constraints
- Tetap menggunakan struktur database PDO yang telah dikonfigurasi.
- Menggunakan autoloader `../vendor/autoload.php` untuk memuat dependensi Dompdf.

---

### Task 1: Modifikasi Laporan Rekapitulasi Umum (`laporan.php`)

**Files:**
- Modify: `d:\laragon\www\spt-php\modules\laporan.php`

**Interfaces:**
- Consumes: Database `surat_tugas`, `pegawai`, `pegawai_tugas`
- Produces: PDF stream output untuk laporan rekapitulasi surat tugas

- [ ] **Step 1: Tambahkan autoload vendor di bagian atas**
  Letakkan autoload di baris paling atas berkas setelah tag pembuka php:
  ```php
  require_once '../vendor/autoload.php';
  use Dompdf\Dompdf;
  use Dompdf\Options;
  ```

- [ ] **Step 2: Ubah kondisi kueri data untuk mendukung export=pdf**
  Modifikasi bagian penarikan data (baris 18-38) agar kondisi `export === 'excel'` juga menyertakan `export === 'pdf'`:
  ```php
  // Ambil data jika ada filter
  if ($start_date && $end_date) {
      // Untuk export dan PDF, tetap ambil semua data sesuai filter periode
      if ($export === 'excel' || $export === 'pdf') {
          $query = "SELECT 
              st.*,
              GROUP_CONCAT(DISTINCT p.nama SEPARATOR ', ') as pegawai_names
              FROM surat_tugas st 
              LEFT JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas 
              LEFT JOIN pegawai p ON pt.nip = p.nip 
              WHERE st.tanggal_surat BETWEEN :start_date AND :end_date
              GROUP BY st.id 
              ORDER BY st.tanggal_surat ASC, st.id ASC";
          
          $stmt = $pdo->prepare($query);
          $stmt->execute([
              ':start_date' => $start_date,
              ':end_date' => $end_date
          ]);
          $laporan = $stmt->fetchAll();
      } else {
          // Kueri dengan paginasi
  ```

- [ ] **Step 3: Tambahkan penanganan ekspor PDF (Dompdf)**
  Tambahkan blok ekspor PDF di bawah penanganan ekspor Excel (sebelum `include '../includes/header.php'`):
  ```php
  // Export PDF berdasarkan data/filter yang sedang ditampilkan
  if ($export === 'pdf') {
      ob_start();
      ?>
      <!DOCTYPE html>
      <html lang="id">
      <head>
          <meta charset="UTF-8">
          <title>Laporan Rekapitulasi Surat Tugas</title>
          <style>
              body {
                  font-family: 'Helvetica', Arial, sans-serif;
                  font-size: 11px;
                  line-height: 1.4;
                  color: #333;
                  margin: 0;
                  padding: 0;
              }
              h1 {
                  text-align: center;
                  font-size: 16px;
                  margin-bottom: 2px;
                  text-transform: uppercase;
              }
              .subtitle {
                  text-align: center;
                  font-size: 11px;
                  margin-bottom: 20px;
                  color: #555;
              }
              table {
                  width: 100%;
                  border-collapse: collapse;
                  margin-top: 10px;
              }
              th {
                  background-color: #f3f4f6;
                  border: 1px solid #000;
                  padding: 6px 8px;
                  font-weight: bold;
                  text-align: left;
                  text-transform: uppercase;
                  font-size: 9px;
              }
              td {
                  border: 1px solid #000;
                  padding: 6px 8px;
                  vertical-align: top;
              }
              .text-center { text-align: center; }
              .text-nowrap { white-space: nowrap; }
              .signature-container {
                  margin-top: 40px;
                  float: right;
                  width: 250px;
                  text-align: center;
              }
              .signature-space {
                  height: 60px;
              }
          </style>
      </head>
      <body>
          <h1>Laporan Rekapitulasi Surat Tugas</h1>
          <p class="subtitle">Periode: <?php echo date('d/m/Y', strtotime($start_date)); ?> s/d <?php echo date('d/m/Y', strtotime($end_date)); ?></p>
          
          <table>
              <thead>
                  <tr>
                      <th style="width: 5%;" class="text-center">No</th>
                      <th style="width: 20%;">Nomor Surat</th>
                      <th style="width: 12%;">Tanggal</th>
                      <th style="width: 30%;">Tujuan / Maksud</th>
                      <th style="width: 20%;">Pegawai</th>
                      <th style="width: 13%;">Pelaksanaan</th>
                  </tr>
              </thead>
              <tbody>
                  <?php if (count($laporan) > 0): ?>
                      <?php $no = 1; foreach ($laporan as $row): ?>
                          <tr>
                              <td class="text-center"><?php echo $no++; ?></td>
                              <td><?php echo htmlspecialchars($row['nomor_surat'], ENT_QUOTES, 'UTF-8'); ?></td>
                              <td class="text-nowrap"><?php echo date('d/m/Y', strtotime($row['tanggal_surat'])); ?></td>
                              <td><?php echo htmlspecialchars($row['untuk'], ENT_QUOTES, 'UTF-8'); ?></td>
                              <td><?php echo htmlspecialchars($row['pegawai_names'], ENT_QUOTES, 'UTF-8'); ?></td>
                              <td class="text-nowrap">
                                  <?php 
                                  $pelaksanaan = date('d/m/Y', strtotime($row['tanggal_mulai']));
                                  if ($row['tanggal_mulai'] != $row['tanggal_selesai']) {
                                      $pelaksanaan .= ' - ' . date('d/m/Y', strtotime($row['tanggal_selesai']));
                                  }
                                  echo $pelaksanaan;
                                  ?>
                              </td>
                          </tr>
                      <?php endforeach; ?>
                  <?php else: ?>
                      <tr>
                          <td colspan="6" class="text-center" style="padding: 20px;">Tidak ada data surat tugas pada periode ini.</td>
                      </tr>
                  <?php endif; ?>
              </tbody>
          </table>

          <div class="signature-container">
              <p>Mengetahui,</p>
              <div class="signature-space"></div>
              <p style="font-weight: bold; text-decoration: underline;">Nama Pejabat</p>
              <p>NIP. ...........................</p>
          </div>
      </body>
      </html>
      <?php
      $html = ob_get_clean();

      $options = new Options();
      $options->set('isHtml5ParserEnabled', true);
      $options->set('isRemoteEnabled', true);
      
      $dompdf = new Dompdf($options);
      $dompdf->loadHtml($html);
      $dompdf->setPaper('A4', 'landscape');
      $dompdf->render();
      
      $filename = 'laporan_surat_tugas_' . $start_date . '_sd_' . $end_date . '.pdf';
      $filename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', $filename);
      $dompdf->stream($filename, ['Attachment' => false]);
      exit;
  }
  ```

- [ ] **Step 4: Ubah tombol "Cetak Laporan" menjadi link**
  Modifikasi tombol cetak lama di HTML (sekitar baris 135) menjadi tag `<a>` yang mengarah ke ekspor PDF:
  ```html
  <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&export=pdf" target="_blank" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
      <i class='bx bx-printer'></i> Cetak Laporan
  </a>
  ```

---

### Task 2: Modifikasi Laporan Per Pegawai (`laporan-pegawai.php`)

**Files:**
- Modify: `d:\laragon\www\spt-php\modules\laporan-pegawai.php`

**Interfaces:**
- Consumes: Database `pegawai`, `pegawai_tugas`, `surat_tugas`
- Produces: PDF stream output untuk laporan surat tugas per pegawai

- [ ] **Step 1: Tambahkan autoload vendor di bagian atas**
  Tambahkan baris autoload di baris paling atas setelah tag pembuka php:
  ```php
  require_once '../vendor/autoload.php';
  use Dompdf\Dompdf;
  use Dompdf\Options;
  ```

- [ ] **Step 2: Ubah kondisi kueri data untuk mendukung export=pdf**
  Modifikasi bagian penarikan data (baris 45-47) agar kondisi `export === 'excel'` juga menyertakan `export === 'pdf'`:
  ```php
  if ($start_date && $end_date) {
      if ($export === 'excel' || $export === 'pdf') {
          // Untuk export dan PDF: ambil semua data sesuai filter periode (tanpa pagination)
  ```

- [ ] **Step 3: Tambahkan penanganan ekspor PDF (Dompdf)**
  Tambahkan penanganan cetak PDF di bawah penanganan ekspor Excel (sebelum `include '../includes/header.php'`):
  ```php
  // Export PDF berdasarkan data/filter yang sedang ditampilkan
  if ($export === 'pdf') {
      ob_start();
      ?>
      <!DOCTYPE html>
      <html lang="id">
      <head>
          <meta charset="UTF-8">
          <title>Laporan Surat Tugas Per Pegawai</title>
          <style>
              body {
                  font-family: 'Helvetica', Arial, sans-serif;
                  font-size: 11px;
                  line-height: 1.4;
                  color: #333;
                  margin: 0;
                  padding: 0;
              }
              h1 {
                  text-align: center;
                  font-size: 16px;
                  margin-bottom: 2px;
                  text-transform: uppercase;
              }
              .subtitle {
                  text-align: center;
                  font-size: 11px;
                  margin-bottom: 20px;
                  color: #555;
              }
              .employee-card {
                  border: 1px solid #94a3b8;
                  margin-bottom: 20px;
                  page-break-inside: avoid;
              }
              .employee-header {
                  background-color: #4f46e5;
                  color: #ffffff;
                  padding: 10px;
              }
              .employee-name {
                  font-size: 14px;
                  font-weight: bold;
                  margin: 0 0 5px 0;
              }
              .employee-info {
                  font-size: 10px;
                  margin: 0;
              }
              table {
                  width: 100%;
                  border-collapse: collapse;
              }
              th {
                  background-color: #f3f4f6;
                  border: 1px solid #000;
                  padding: 6px;
                  font-weight: bold;
                  text-align: left;
                  text-transform: uppercase;
                  font-size: 9px;
              }
              td {
                  border: 1px solid #000;
                  padding: 6px;
                  vertical-align: top;
              }
              .text-center { text-align: center; }
              .text-nowrap { white-space: nowrap; }
              .badge {
                  background-color: #e0e7ff;
                  color: #3730a3;
                  padding: 2px 6px;
                  border-radius: 4px;
                  font-size: 9px;
              }
              .signature-container {
                  margin-top: 30px;
                  float: right;
                  width: 200px;
                  text-align: center;
              }
              .signature-space {
                  height: 50px;
              }
          </style>
      </head>
      <body>
          <h1>Laporan Surat Tugas Per Pegawai</h1>
          <p class="subtitle">
              Periode: <?php echo date('d/m/Y', strtotime($start_date)); ?> s/d <?php echo date('d/m/Y', strtotime($end_date)); ?>
              <?php if (!empty($filter_nip) && isset($laporan[0]['nama'])): ?>
                  | Pegawai: <?php echo htmlspecialchars($laporan[0]['nama'], ENT_QUOTES, 'UTF-8'); ?>
              <?php endif; ?>
          </p>

          <?php if (count($laporan) > 0): ?>
              <?php foreach ($laporan as $pegawai): ?>
                  <div class="employee-card">
                      <div class="employee-header">
                          <table style="width: 100%; border: none;">
                              <tr style="border: none;">
                                  <td style="width: 70%; border: none; padding: 0; color: white; vertical-align: middle;">
                                      <div class="employee-name"><?php echo htmlspecialchars($pegawai['nama'], ENT_QUOTES, 'UTF-8'); ?></div>
                                      <p class="employee-info">NIP: <?php echo htmlspecialchars($pegawai['nip'], ENT_QUOTES, 'UTF-8'); ?> | Pangkat: <?php echo htmlspecialchars($pegawai['pangkat'], ENT_QUOTES, 'UTF-8'); ?> | Jabatan: <?php echo htmlspecialchars($pegawai['jabatan'], ENT_QUOTES, 'UTF-8'); ?></p>
                                  </td>
                                  <td style="width: 30%; border: none; padding: 0; text-align: right; color: white; vertical-align: middle;">
                                      <strong>Total Tugas: <?php echo $pegawai['total_tugas']; ?></strong> (<?php echo $pegawai['total_hari']; ?> Hari)
                                  </td>
                              </tr>
                          </table>
                      </div>
                      <table style="width: 100%;">
                          <thead>
                              <tr>
                                  <th style="width: 5%;" class="text-center">No</th>
                                  <th style="width: 25%;">Nomor Surat</th>
                                  <th style="width: 15%;">Tanggal Surat</th>
                                  <th style="width: 35%;">Tujuan / Maksud</th>
                                  <th style="width: 15%;">Pelaksanaan</th>
                                  <th style="width: 8%;" class="text-center">Durasi</th>
                              </tr>
                          </thead>
                          <tbody>
                              <?php $no = 1; foreach ($pegawai['surat_tugas'] as $surat): ?>
                                  <tr>
                                      <td class="text-center"><?php echo $no++; ?></td>
                                      <td><?php echo htmlspecialchars($surat['nomor_surat'], ENT_QUOTES, 'UTF-8'); ?></td>
                                      <td class="text-nowrap"><?php echo date('d/m/Y', strtotime($surat['tanggal_surat'])); ?></td>
                                      <td><?php echo htmlspecialchars($surat['untuk'], ENT_QUOTES, 'UTF-8'); ?></td>
                                      <td class="text-nowrap">
                                          <?php 
                                          $pelaksanaan = date('d/m/Y', strtotime($surat['tanggal_mulai']));
                                          if ($surat['tanggal_mulai'] != $surat['tanggal_selesai']) {
                                              $pelaksanaan .= ' - ' . date('d/m/Y', strtotime($surat['tanggal_selesai']));
                                          }
                                          echo $pelaksanaan;
                                          ?>
                                      </td>
                                      <td class="text-center">
                                          <span class="badge"><?php echo $surat['durasi_hari']; ?> hari</span>
                                      </td>
                                  </tr>
                              <?php endforeach; ?>
                          </tbody>
                      </table>
                  </div>
              <?php endforeach; ?>
          <?php else: ?>
              <div style="text-align: center; padding: 30px; border: 1px solid #94a3b8;">
                  <p>Tidak ada data surat tugas dengan filter yang dipilih.</p>
              </div>
          <?php endif; ?>

          <div class="signature-container">
              <p>Mengetahui,</p>
              <div class="signature-space"></div>
              <p style="font-weight: bold; text-decoration: underline;">Nama Pejabat</p>
              <p>NIP. ...........................</p>
          </div>
      </body>
      </html>
      <?php
      $html = ob_get_clean();

      $options = new Options();
      $options->set('isHtml5ParserEnabled', true);
      $options->set('isRemoteEnabled', true);
      
      $dompdf = new Dompdf($options);
      $dompdf->loadHtml($html);
      $dompdf->setPaper('A4', 'portrait');
      $dompdf->render();
      
      $suffixPegawai = !empty($filter_nip) ? '_' . $filter_nip : '_semua_pegawai';
      $filename = 'laporan_per_pegawai_' . $start_date . '_sd_' . $end_date . $suffixPegawai . '.pdf';
      $filename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', $filename);
      $dompdf->stream($filename, ['Attachment' => false]);
      exit;
  }
  ```

- [ ] **Step 4: Ubah tombol "Cetak Laporan" menjadi link**
  Modifikasi tombol cetak lama di HTML (sekitar baris 303) menjadi tag `<a>` yang mengarah to ekspor PDF:
  ```html
  <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&nip=<?php echo urlencode($filter_nip); ?>&export=pdf" target="_blank" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
      <i class='bx bx-printer'></i> Cetak Laporan
  </a>
  ```

---

## Rencana Verifikasi

### Manual Verification
1. Jalankan server lokal.
2. Buka menu **Laporan**. Terapkan filter tanggal yang memiliki lebih dari 10 data.
3. Klik tombol **Cetak Laporan** (ekspor PDF). Pastikan tab baru terbuka dan menghasilkan PDF landscape berisi seluruh baris data (lebih dari 10 baris).
4. Buka menu **Laporan Per Pegawai**. Terapkan filter tanggal/pegawai.
5. Klik tombol **Cetak Laporan**. Pastikan tab baru terbuka dan menghasilkan PDF portrait berisi seluruh baris data per pegawai yang difilter.
