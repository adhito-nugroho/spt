<?php
$pdo = require_once '../config/database.php';

function ensureSuratKeluarTable(PDO $pdo): void
{
    $pdo->exec("CREATE TABLE IF NOT EXISTS surat_keluar (
        id INT AUTO_INCREMENT PRIMARY KEY,
        tahun INT NOT NULL,
        nomor_urut INT NOT NULL,
        nomor_surat VARCHAR(120) NOT NULL,
        tanggal_surat DATE NOT NULL,
        tujuan_surat TEXT NOT NULL,
        perihal TEXT NOT NULL,
        keterangan TEXT DEFAULT NULL,
        file_pdf VARCHAR(255) DEFAULT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        UNIQUE KEY uniq_tahun_nomor_urut (tahun, nomor_urut),
        UNIQUE KEY uniq_tahun_nomor_surat (tahun, nomor_surat),
        INDEX idx_tanggal_surat (tanggal_surat),
        INDEX idx_tahun_tanggal (tahun, tanggal_surat)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci");
}

function ensureSuratKeluarPdfColumn(PDO $pdo): void
{
    try {
        $pdo->query("SELECT file_pdf FROM surat_keluar LIMIT 1");
    } catch (PDOException $e) {
        $pdo->exec("ALTER TABLE surat_keluar ADD COLUMN file_pdf VARCHAR(255) DEFAULT NULL AFTER keterangan");
    }
}

function ensureBukuNomorSuratKeluarTable(PDO $pdo): void
{
    $pdo->exec("CREATE TABLE IF NOT EXISTS buku_nomor_surat_keluar (
        id INT AUTO_INCREMENT PRIMARY KEY,
        tahun INT NOT NULL,
        nomor_urut INT NOT NULL,
        tanggal_surat DATE DEFAULT NULL,
        status ENUM('kosong','terisi') NOT NULL DEFAULT 'kosong',
        id_surat_keluar INT DEFAULT NULL,
        keterangan VARCHAR(255) DEFAULT NULL,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        UNIQUE KEY uniq_tahun_nomor_urut (tahun, nomor_urut),
        UNIQUE KEY uniq_id_surat_keluar (id_surat_keluar),
        INDEX idx_status_tanggal (status, tanggal_surat)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci");
}

function normalizeHeaderName(string $value): string
{
    $value = preg_replace('/^\xEF\xBB\xBF/', '', $value);
    $value = strtolower(trim($value));
    $value = str_replace(['.', '_', '-'], ' ', $value);
    return preg_replace('/\s+/', ' ', $value);
}

function parseTanggalSurat($value): ?string
{
    if ($value === null) {
        return null;
    }

    $value = trim((string)$value);
    if ($value === '') {
        return null;
    }

    if (is_numeric($value)) {
        $serial = (int)$value;
        if ($serial > 20000) {
            $timestamp = ($serial - 25569) * 86400;
            return gmdate('Y-m-d', $timestamp);
        }
    }

    $formats = ['Y-m-d', 'd/m/Y', 'd-m-Y', 'm/d/Y', 'd.m.Y'];
    foreach ($formats as $format) {
        $date = DateTime::createFromFormat($format, $value);
        if ($date instanceof DateTime) {
            return $date->format('Y-m-d');
        }
    }

    $timestamp = strtotime($value);
    return $timestamp === false ? null : date('Y-m-d', $timestamp);
}

function handleOptionalSuratKeluarPdfUpload(?string $existingFile = null): ?string
{
    if (!isset($_FILES['file_pdf']) || $_FILES['file_pdf']['error'] === UPLOAD_ERR_NO_FILE) {
        return $existingFile;
    }

    $file = $_FILES['file_pdf'];
    if ($file['error'] !== UPLOAD_ERR_OK) {
        throw new Exception('Upload file PDF gagal.');
    }

    $extension = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
    if ($extension !== 'pdf') {
        throw new Exception('File lampiran harus berformat PDF.');
    }

    $maxSize = 20 * 1024 * 1024;
    if ((int)$file['size'] > $maxSize) {
        throw new Exception('Ukuran file PDF maksimal 20 MB.');
    }

    $uploadDir = __DIR__ . '/../assets/surat_keluar_pdf/';
    if (!is_dir($uploadDir) && !mkdir($uploadDir, 0775, true)) {
        throw new Exception('Folder upload PDF tidak bisa dibuat.');
    }

    $safeBaseName = preg_replace('/[^a-zA-Z0-9_-]+/', '_', pathinfo($file['name'], PATHINFO_FILENAME));
    $safeBaseName = trim($safeBaseName, '_') ?: 'surat_keluar';
    $newFileName = date('Ymd_His') . '_' . bin2hex(random_bytes(4)) . '_' . $safeBaseName . '.pdf';
    $targetPath = $uploadDir . $newFileName;

    if (!move_uploaded_file($file['tmp_name'], $targetPath)) {
        throw new Exception('File PDF gagal disimpan.');
    }

    return $newFileName;
}

function getNextSuratKeluarNomor(PDO $pdo, int $tahun): int
{
    $stmt = $pdo->prepare("SELECT COALESCE(MAX(nomor_urut), 0) FROM buku_nomor_surat_keluar WHERE tahun = ?");
    $stmt->execute([$tahun]);
    return ((int)$stmt->fetchColumn()) + 1;
}

function getLastTanggalSuratKeluar(PDO $pdo, int $tahun): ?string
{
    $stmt = $pdo->prepare("SELECT MAX(tanggal_surat) FROM buku_nomor_surat_keluar WHERE tahun = ? AND tanggal_surat IS NOT NULL");
    $stmt->execute([$tahun]);
    $value = $stmt->fetchColumn();
    return $value ?: null;
}

function syncSuratKeluarTerpakaiKeBukuNomor(PDO $pdo): void
{
    $pdo->exec("INSERT INTO buku_nomor_surat_keluar (
            tahun, nomor_urut, tanggal_surat, status, id_surat_keluar, keterangan
        )
        SELECT tahun, nomor_urut, tanggal_surat, 'terisi', id, 'Sinkron dari data surat keluar'
        FROM surat_keluar
        WHERE nomor_urut > 0
        ON DUPLICATE KEY UPDATE
            tanggal_surat = VALUES(tanggal_surat),
            status = 'terisi',
            id_surat_keluar = VALUES(id_surat_keluar),
            keterangan = VALUES(keterangan)");
}

function createSuratKeluarSpaceByTanggal(PDO $pdo, int $tahun, string $tanggal, int $jumlah): array
{
    $lastTanggal = getLastTanggalSuratKeluar($pdo, $tahun);
    if ($lastTanggal !== null && $tanggal < $lastTanggal) {
        throw new Exception('Buat Space tidak bisa memakai tanggal sebelum tanggal terakhir buku (' . date('d/m/Y', strtotime($lastTanggal)) . '). Gunakan Tambah Slot Kosong untuk tanggal mundur.');
    }

    $jumlah = max(1, min(100, $jumlah));
    $nomor = getNextSuratKeluarNomor($pdo, $tahun);
    $created = 0;
    $first = null;
    $last = null;

    while ($created < $jumlah) {
        $stmt = $pdo->prepare("INSERT INTO buku_nomor_surat_keluar (tahun, nomor_urut, tanggal_surat, status, keterangan)
            VALUES (:tahun, :nomor_urut, :tanggal_surat, 'kosong', 'Space otomatis per tanggal')");
        $stmt->execute([
            ':tahun' => $tahun,
            ':nomor_urut' => $nomor,
            ':tanggal_surat' => $tanggal,
        ]);
        if ($first === null) {
            $first = $nomor;
        }
        $last = $nomor;
        $created++;
        $nomor++;
    }

    return ['created' => $created, 'first' => $first, 'last' => $last];
}

function bindNomorKeSuratKeluar(PDO $pdo, int $idSuratKeluar, int $nomorUrut, string $tanggalSurat): void
{
    $tahun = (int)date('Y', strtotime($tanggalSurat));
    $tanggalSurat = date('Y-m-d', strtotime($tanggalSurat));

    $stmtRelease = $pdo->prepare("UPDATE buku_nomor_surat_keluar
        SET id_surat_keluar = NULL, status = 'kosong', keterangan = 'Nomor dikosongkan setelah update'
        WHERE id_surat_keluar = :id_surat_keluar
          AND NOT (tahun = :tahun AND nomor_urut = :nomor_urut)");
    $stmtRelease->execute([
        ':id_surat_keluar' => $idSuratKeluar,
        ':tahun' => $tahun,
        ':nomor_urut' => $nomorUrut,
    ]);

    $stmt = $pdo->prepare("SELECT id_surat_keluar, status, tanggal_surat
        FROM buku_nomor_surat_keluar
        WHERE tahun = :tahun AND nomor_urut = :nomor_urut
        LIMIT 1");
    $stmt->execute([
        ':tahun' => $tahun,
        ':nomor_urut' => $nomorUrut,
    ]);
    $row = $stmt->fetch(PDO::FETCH_ASSOC);

    if (!$row) {
        throw new Exception('Nomor belum dibuat di Buku Nomor Surat Keluar.');
    }
    if (!empty($row['tanggal_surat']) && $row['tanggal_surat'] !== $tanggalSurat) {
        throw new Exception('Nomor ini diperuntukkan untuk tanggal surat lain.');
    }
    if (!empty($row['id_surat_keluar']) && (int)$row['id_surat_keluar'] !== $idSuratKeluar) {
        throw new Exception('Nomor surat keluar sudah terpakai.');
    }
    if ($row['status'] !== 'kosong' && (int)($row['id_surat_keluar'] ?? 0) !== $idSuratKeluar) {
        throw new Exception('Status nomor bukan kosong, tidak dapat dipakai.');
    }

    $stmt = $pdo->prepare("UPDATE buku_nomor_surat_keluar
        SET status = 'terisi',
            id_surat_keluar = :id_surat_keluar,
            tanggal_surat = :tanggal_surat,
            keterangan = 'Terbit dari modul surat keluar'
        WHERE tahun = :tahun AND nomor_urut = :nomor_urut");
    $stmt->execute([
        ':id_surat_keluar' => $idSuratKeluar,
        ':tanggal_surat' => $tanggalSurat,
        ':tahun' => $tahun,
        ':nomor_urut' => $nomorUrut,
    ]);
}

function saveSuratKeluar(PDO $pdo, array $data, ?int $id = null): int
{
    $tanggal = parseTanggalSurat($data['tanggal_surat'] ?? '');
    if ($tanggal === null) {
        throw new Exception('Tanggal surat wajib diisi dan harus valid.');
    }

    $tahun = (int)date('Y', strtotime($tanggal));
    $nomorUrut = (int)($data['nomor_urut'] ?? 0);
    $nomorSurat = trim((string)($data['nomor_surat'] ?? ''));
    $tujuan = trim((string)($data['tujuan_surat'] ?? ''));
    $perihal = trim((string)($data['perihal'] ?? ''));
    $keterangan = trim((string)($data['keterangan'] ?? ''));
    $filePdf = trim((string)($data['file_pdf'] ?? ''));

    if ($nomorUrut <= 0) {
        throw new Exception('Nomor urut wajib diisi.');
    }
    if ($nomorSurat === '') {
        throw new Exception('Nomor surat wajib diisi.');
    }
    if ($tujuan === '') {
        throw new Exception('Tujuan surat wajib diisi.');
    }
    if ($perihal === '') {
        throw new Exception('Perihal wajib diisi.');
    }

    if ($id === null) {
        $stmt = $pdo->prepare("INSERT INTO surat_keluar
            (tahun, nomor_urut, nomor_surat, tanggal_surat, tujuan_surat, perihal, keterangan, file_pdf)
            VALUES (:tahun, :nomor_urut, :nomor_surat, :tanggal_surat, :tujuan_surat, :perihal, :keterangan, :file_pdf)");
    } else {
        $stmt = $pdo->prepare("UPDATE surat_keluar SET
            tahun = :tahun,
            nomor_urut = :nomor_urut,
            nomor_surat = :nomor_surat,
            tanggal_surat = :tanggal_surat,
            tujuan_surat = :tujuan_surat,
            perihal = :perihal,
            keterangan = :keterangan,
            file_pdf = :file_pdf
            WHERE id = :id");
        $stmt->bindValue(':id', $id, PDO::PARAM_INT);
    }

    $stmt->bindValue(':tahun', $tahun, PDO::PARAM_INT);
    $stmt->bindValue(':nomor_urut', $nomorUrut, PDO::PARAM_INT);
    $stmt->bindValue(':nomor_surat', $nomorSurat);
    $stmt->bindValue(':tanggal_surat', $tanggal);
    $stmt->bindValue(':tujuan_surat', $tujuan);
    $stmt->bindValue(':perihal', $perihal);
    $stmt->bindValue(':keterangan', $keterangan !== '' ? $keterangan : null);
    $stmt->bindValue(':file_pdf', $filePdf !== '' ? $filePdf : null);
    $stmt->execute();

    return $id ?? (int)$pdo->lastInsertId();
}

function saveImportedSuratKeluar(PDO $pdo, array $data, ?int $id = null): int
{
    $tanggal = parseTanggalSurat($data['tanggal_surat'] ?? '');
    if ($tanggal === null) {
        throw new Exception('Tanggal surat import tidak valid.');
    }

    $tahun = (int)date('Y', strtotime($tanggal));
    $nomorUrut = (int)($data['nomor_urut'] ?? 0);
    $nomorSurat = trim((string)($data['nomor_surat'] ?? ''));
    if ($nomorUrut <= 0 || $nomorSurat === '') {
        throw new Exception('Nomor urut dan nomor surat import wajib ada untuk data surat terisi.');
    }

    if ($id === null) {
        $stmt = $pdo->prepare("INSERT INTO surat_keluar
            (tahun, nomor_urut, nomor_surat, tanggal_surat, tujuan_surat, perihal, keterangan, file_pdf)
            VALUES (:tahun, :nomor_urut, :nomor_surat, :tanggal_surat, :tujuan_surat, :perihal, :keterangan, :file_pdf)");
    } else {
        $stmt = $pdo->prepare("UPDATE surat_keluar SET
            tahun = :tahun,
            nomor_urut = :nomor_urut,
            nomor_surat = :nomor_surat,
            tanggal_surat = :tanggal_surat,
            tujuan_surat = :tujuan_surat,
            perihal = :perihal,
            keterangan = :keterangan,
            file_pdf = COALESCE(file_pdf, :file_pdf)
            WHERE id = :id");
        $stmt->bindValue(':id', $id, PDO::PARAM_INT);
    }

    $stmt->bindValue(':tahun', $tahun, PDO::PARAM_INT);
    $stmt->bindValue(':nomor_urut', $nomorUrut, PDO::PARAM_INT);
    $stmt->bindValue(':nomor_surat', $nomorSurat);
    $stmt->bindValue(':tanggal_surat', $tanggal);
    $stmt->bindValue(':tujuan_surat', trim((string)($data['tujuan_surat'] ?? '')));
    $stmt->bindValue(':perihal', trim((string)($data['perihal'] ?? '')));
    $keterangan = trim((string)($data['keterangan'] ?? ''));
    $stmt->bindValue(':keterangan', $keterangan !== '' ? $keterangan : null);
    $stmt->bindValue(':file_pdf', null);
    $stmt->execute();

    return $id ?? (int)$pdo->lastInsertId();
}

function csvRows(string $path): array
{
    $rows = [];
    $handle = fopen($path, 'r');
    if ($handle === false) {
        throw new Exception('File CSV tidak bisa dibaca.');
    }
    while (($row = fgetcsv($handle, 0, ',')) !== false) {
        if (count($row) === 1 && strpos((string)$row[0], ';') !== false) {
            $row = str_getcsv((string)$row[0], ';');
        }
        $rows[] = array_map(static fn($value) => trim((string)$value), $row);
    }
    fclose($handle);
    return $rows;
}

function xlsxRows(string $path): array
{
    if (!class_exists('ZipArchive')) {
        throw new Exception('Import XLSX membutuhkan ekstensi PHP ZipArchive. Gunakan CSV jika ekstensi belum aktif.');
    }
    if (!function_exists('simplexml_load_string')) {
        throw new Exception('Import XLSX membutuhkan ekstensi PHP SimpleXML. Gunakan CSV jika ekstensi belum aktif.');
    }

    $zip = new ZipArchive();
    if ($zip->open($path) !== true) {
        throw new Exception('File XLSX tidak bisa dibuka.');
    }

    $sharedStrings = [];
    $sharedXml = $zip->getFromName('xl/sharedStrings.xml');
    if ($sharedXml !== false) {
        $xml = simplexml_load_string($sharedXml);
        if ($xml !== false) {
            foreach ($xml->si as $si) {
                $parts = [];
                if (isset($si->t)) {
                    $parts[] = (string)$si->t;
                }
                if (isset($si->r)) {
                    foreach ($si->r as $run) {
                        $parts[] = (string)$run->t;
                    }
                }
                $sharedStrings[] = implode('', $parts);
            }
        }
    }

    $sheetXml = $zip->getFromName('xl/worksheets/sheet1.xml');
    $zip->close();
    if ($sheetXml === false) {
        throw new Exception('Sheet pertama pada file XLSX tidak ditemukan.');
    }

    $xml = simplexml_load_string($sheetXml);
    if ($xml === false) {
        throw new Exception('Isi XLSX tidak valid.');
    }

    $rows = [];
    foreach ($xml->sheetData->row as $rowXml) {
        $row = [];
        foreach ($rowXml->c as $cell) {
            $ref = (string)$cell['r'];
            preg_match('/[A-Z]+/', $ref, $match);
            $column = $match[0] ?? 'A';
            $index = 0;
            for ($i = 0; $i < strlen($column); $i++) {
                $index = ($index * 26) + (ord($column[$i]) - 64);
            }
            $index--;

            $type = (string)$cell['t'];
            $value = isset($cell->v) ? (string)$cell->v : '';
            if ($type === 's') {
                $value = $sharedStrings[(int)$value] ?? '';
            } elseif ($type === 'inlineStr' && isset($cell->is->t)) {
                $value = (string)$cell->is->t;
            }
            $row[$index] = trim($value);
        }
        if (!empty($row)) {
            ksort($row);
            $maxIndex = max(array_keys($row));
            $normalizedRow = [];
            for ($i = 0; $i <= $maxIndex; $i++) {
                $normalizedRow[] = $row[$i] ?? '';
            }
            $rows[] = $normalizedRow;
        }
    }

    return $rows;
}

function rowsToSuratKeluarData(array $rows): array
{
    $rows = array_values(array_filter($rows, static function ($row) {
        return count(array_filter($row, static fn($value) => trim((string)$value) !== '')) > 0;
    }));
    if (count($rows) < 2) {
        throw new Exception('File import tidak memiliki data.');
    }

    $headerIndex = null;
    $header = [];
    foreach ($rows as $i => $row) {
        $normalized = array_map('normalizeHeaderName', $row);
        if (in_array('no', $normalized, true) || in_array('nomor', $normalized, true) || in_array('no surat', $normalized, true) || in_array('nomor surat', $normalized, true)) {
            $headerIndex = $i;
            $header = $normalized;
            break;
        }
    }
    if ($headerIndex === null) {
        throw new Exception('Header kolom tidak ditemukan. Gunakan header: No, No. Surat, Tanggal Surat, Tujuan Surat, Perihal, Keterangan.');
    }

    $map = [];
    foreach ($header as $i => $name) {
        if (in_array($name, ['no', 'nomor'], true)) $map['nomor_urut'] = $i;
        if (in_array($name, ['no surat', 'nomor surat'], true)) $map['nomor_surat'] = $i;
        if ($name === 'tanggal surat') $map['tanggal_surat'] = $i;
        if ($name === 'tujuan surat') $map['tujuan_surat'] = $i;
        if ($name === 'perihal') $map['perihal'] = $i;
        if ($name === 'keterangan') $map['keterangan'] = $i;
    }

    foreach (['nomor_urut'] as $required) {
        if (!array_key_exists($required, $map)) {
            throw new Exception('Kolom wajib belum lengkap. Pastikan ada kolom No untuk nomor slot.');
        }
    }

    $data = [];
    $lastNomorUrut = 0;
    for ($i = $headerIndex + 1; $i < count($rows); $i++) {
        $row = $rows[$i];
        $hasRowData = count(array_filter($row, static fn($value) => trim((string)$value) !== '')) > 0;
        if (!$hasRowData) {
            continue;
        }

        $nomorRaw = (string)($row[$map['nomor_urut']] ?? '');
        $nomorUrut = (int)preg_replace('/\D+/', '', $nomorRaw);
        if ($nomorUrut <= 0) {
            $nomorUrut = $lastNomorUrut + 1;
        }
        $lastNomorUrut = $nomorUrut;

        $data[] = [
            'nomor_urut' => $nomorUrut,
            'nomor_surat' => isset($map['nomor_surat']) ? trim((string)($row[$map['nomor_surat']] ?? '')) : '',
            'tanggal_surat' => isset($map['tanggal_surat']) ? ($row[$map['tanggal_surat']] ?? '') : '',
            'tujuan_surat' => isset($map['tujuan_surat']) ? ($row[$map['tujuan_surat']] ?? '') : '',
            'perihal' => isset($map['perihal']) ? ($row[$map['perihal']] ?? '') : '',
            'keterangan' => isset($map['keterangan']) ? ($row[$map['keterangan']] ?? '') : '',
            'source_row' => $i + 1,
        ];
    }

    return $data;
}

ensureSuratKeluarTable($pdo);
ensureSuratKeluarPdfColumn($pdo);
ensureBukuNomorSuratKeluarTable($pdo);
syncSuratKeluarTerpakaiKeBukuNomor($pdo);

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_POST['action'])) {
    try {
        if ($_POST['action'] === 'create') {
            $pdo->beginTransaction();
            $newFile = null;
            try {
                $data = $_POST;
                $newFile = handleOptionalSuratKeluarPdfUpload(null);
                $data['file_pdf'] = $newFile;
                $id = saveSuratKeluar($pdo, $data);
                bindNomorKeSuratKeluar($pdo, $id, (int)$_POST['nomor_urut'], (string)$_POST['tanggal_surat']);
                $pdo->commit();
            } catch (Throwable $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                if ($newFile) {
                    $newPath = __DIR__ . '/../assets/surat_keluar_pdf/' . basename((string)$newFile);
                    if (is_file($newPath)) {
                        @unlink($newPath);
                    }
                }
                throw $e;
            }
            header('Location: surat-keluar.php?ok=' . urlencode('Data surat keluar berhasil ditambahkan.'));
            exit;
        }

        if ($_POST['action'] === 'update') {
            $id = (int)($_POST['id'] ?? 0);
            if ($id <= 0) {
                throw new Exception('ID surat keluar tidak valid.');
            }
            $pdo->beginTransaction();
            $oldFile = null;
            $newFile = null;
            try {
                $stmtOld = $pdo->prepare("SELECT file_pdf FROM surat_keluar WHERE id = ? LIMIT 1");
                $stmtOld->execute([$id]);
                $oldFile = $stmtOld->fetchColumn() ?: null;

                $data = $_POST;
                $newFile = handleOptionalSuratKeluarPdfUpload($oldFile ? (string)$oldFile : null);
                $data['file_pdf'] = $newFile;
                saveSuratKeluar($pdo, $data, $id);
                bindNomorKeSuratKeluar($pdo, $id, (int)$_POST['nomor_urut'], (string)$_POST['tanggal_surat']);
                $pdo->commit();

                if ($oldFile && $newFile && $oldFile !== $newFile) {
                    $oldPath = __DIR__ . '/../assets/surat_keluar_pdf/' . basename((string)$oldFile);
                    if (is_file($oldPath)) {
                        @unlink($oldPath);
                    }
                }
            } catch (Throwable $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                if ($newFile && $newFile !== $oldFile) {
                    $newPath = __DIR__ . '/../assets/surat_keluar_pdf/' . basename((string)$newFile);
                    if (is_file($newPath)) {
                        @unlink($newPath);
                    }
                }
                throw $e;
            }
            header('Location: surat-keluar.php?ok=' . urlencode('Data surat keluar berhasil diperbarui.'));
            exit;
        }

        if ($_POST['action'] === 'delete') {
            $id = (int)($_POST['id'] ?? 0);
            if ($id <= 0) {
                throw new Exception('ID surat keluar tidak valid.');
            }
            $pdo->beginTransaction();
            try {
                $stmtFile = $pdo->prepare("SELECT file_pdf FROM surat_keluar WHERE id = ? LIMIT 1");
                $stmtFile->execute([$id]);
                $deletedFile = $stmtFile->fetchColumn() ?: null;

                $stmt = $pdo->prepare("UPDATE buku_nomor_surat_keluar
                    SET id_surat_keluar = NULL, status = 'kosong', keterangan = 'Nomor dikosongkan setelah data surat keluar dihapus'
                    WHERE id_surat_keluar = ?");
                $stmt->execute([$id]);

                $stmt = $pdo->prepare("DELETE FROM surat_keluar WHERE id = ?");
                $stmt->execute([$id]);
                $pdo->commit();

                if ($deletedFile) {
                    $deletedPath = __DIR__ . '/../assets/surat_keluar_pdf/' . basename((string)$deletedFile);
                    if (is_file($deletedPath)) {
                        @unlink($deletedPath);
                    }
                }
            } catch (Throwable $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                throw $e;
            }
            header('Location: surat-keluar.php?ok=' . urlencode('Data surat keluar berhasil dihapus.'));
            exit;
        }

        if ($_POST['action'] === 'delete_slot') {
            $id = (int)($_POST['id'] ?? 0);
            if ($id <= 0) {
                throw new Exception('ID slot tidak valid.');
            }
            $stmt = $pdo->prepare("SELECT id_surat_keluar FROM buku_nomor_surat_keluar WHERE id = ? LIMIT 1");
            $stmt->execute([$id]);
            $slot = $stmt->fetch(PDO::FETCH_ASSOC);
            if (!$slot) {
                throw new Exception('Slot nomor tidak ditemukan.');
            }
            if (!empty($slot['id_surat_keluar'])) {
                throw new Exception('Slot sudah terisi surat keluar dan tidak bisa dihapus langsung.');
            }
            $stmt = $pdo->prepare("DELETE FROM buku_nomor_surat_keluar WHERE id = ?");
            $stmt->execute([$id]);
            header('Location: surat-keluar.php?ok=' . urlencode('Slot kosong berhasil dihapus.'));
            exit;
        }

        if ($_POST['action'] === 'bulk_delete') {
            $bukuIds = $_POST['buku_ids'] ?? [];
            if (!is_array($bukuIds) || empty($bukuIds)) {
                throw new Exception('Pilih setidaknya satu nomor yang ingin dihapus.');
            }
            $bukuIds = array_values(array_filter(array_map('intval', $bukuIds), function($v) { return $v > 0; }));
            if (empty($bukuIds)) {
                throw new Exception('Daftar ID nomor tidak valid.');
            }

            $placeholders = implode(',', array_fill(0, count($bukuIds), '?'));
            $pdo->beginTransaction();
            try {
                // Cari relasi id_surat_keluar yang terisi
                $stmt = $pdo->prepare("SELECT id, id_surat_keluar FROM buku_nomor_surat_keluar WHERE id IN ($placeholders)");
                $stmt->execute($bukuIds);
                $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

                $suratIds = [];
                foreach ($rows as $r) {
                    if (!empty($r['id_surat_keluar'])) {
                        $suratIds[] = (int)$r['id_surat_keluar'];
                    }
                }

                $filesToDelete = [];
                if (!empty($suratIds)) {
                    $suratPlaceholders = implode(',', array_fill(0, count($suratIds), '?'));
                    $stmtFiles = $pdo->prepare("SELECT file_pdf FROM surat_keluar WHERE id IN ($suratPlaceholders)");
                    $stmtFiles->execute($suratIds);
                    $filesToDelete = $stmtFiles->fetchAll(PDO::FETCH_COLUMN);

                    $stmtDelSurat = $pdo->prepare("DELETE FROM surat_keluar WHERE id IN ($suratPlaceholders)");
                    $stmtDelSurat->execute($suratIds);
                }

                $stmtDelBuku = $pdo->prepare("DELETE FROM buku_nomor_surat_keluar WHERE id IN ($placeholders)");
                $stmtDelBuku->execute($bukuIds);
                $deletedCount = $stmtDelBuku->rowCount();

                $pdo->commit();

                foreach ($filesToDelete as $file) {
                    if ($file) {
                        $deletedPath = __DIR__ . '/../assets/surat_keluar_pdf/' . basename((string)$file);
                        if (is_file($deletedPath)) {
                            @unlink($deletedPath);
                        }
                    }
                }
            } catch (Throwable $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                throw $e;
            }
            $tahunRedirect = (int)($_POST['current_tahun'] ?? date('Y'));
            header('Location: surat-keluar.php?tahun=' . $tahunRedirect . '&ok=' . urlencode("Berhasil menghapus {$deletedCount} nomor sekaligus."));
            exit;
        }

        if ($_POST['action'] === 'delete_range') {
            $tahunRange = (int)($_POST['tahun'] ?? date('Y'));
            $nomorDari = (int)($_POST['nomor_dari'] ?? 0);
            $nomorSampai = (int)($_POST['nomor_sampai'] ?? 0);
            $hanyaKosong = !isset($_POST['include_terisi']) || $_POST['include_terisi'] !== '1';

            if ($tahunRange < 1900) {
                throw new Exception('Tahun tidak valid.');
            }
            if ($nomorDari <= 0 || $nomorSampai <= 0) {
                throw new Exception('Nomor awal dan nomor akhir harus angka lebih besar dari 0.');
            }
            if ($nomorDari > $nomorSampai) {
                throw new Exception('Nomor awal tidak boleh lebih besar dari nomor akhir.');
            }

            $pdo->beginTransaction();
            try {
                if ($hanyaKosong) {
                    $stmt = $pdo->prepare("DELETE FROM buku_nomor_surat_keluar
                        WHERE tahun = ? AND nomor_urut BETWEEN ? AND ? AND id_surat_keluar IS NULL");
                    $stmt->execute([$tahunRange, $nomorDari, $nomorSampai]);
                    $deletedCount = $stmt->rowCount();
                } else {
                    $stmt = $pdo->prepare("SELECT id, id_surat_keluar FROM buku_nomor_surat_keluar
                        WHERE tahun = ? AND nomor_urut BETWEEN ? AND ?");
                    $stmt->execute([$tahunRange, $nomorDari, $nomorSampai]);
                    $rows = $stmt->fetchAll(PDO::FETCH_ASSOC);

                    $suratIds = [];
                    $bukuIds = [];
                    foreach ($rows as $r) {
                        $bukuIds[] = (int)$r['id'];
                        if (!empty($r['id_surat_keluar'])) {
                            $suratIds[] = (int)$r['id_surat_keluar'];
                        }
                    }

                    $filesToDelete = [];
                    if (!empty($suratIds)) {
                        $suratPlaceholders = implode(',', array_fill(0, count($suratIds), '?'));
                        $stmtFiles = $pdo->prepare("SELECT file_pdf FROM surat_keluar WHERE id IN ($suratPlaceholders)");
                        $stmtFiles->execute($suratIds);
                        $filesToDelete = $stmtFiles->fetchAll(PDO::FETCH_COLUMN);

                        $stmtDelSurat = $pdo->prepare("DELETE FROM surat_keluar WHERE id IN ($suratPlaceholders)");
                        $stmtDelSurat->execute($suratIds);
                    }

                    if (!empty($bukuIds)) {
                        $bukuPlaceholders = implode(',', array_fill(0, count($bukuIds), '?'));
                        $stmtDelBuku = $pdo->prepare("DELETE FROM buku_nomor_surat_keluar WHERE id IN ($bukuPlaceholders)");
                        $stmtDelBuku->execute($bukuIds);
                        $deletedCount = $stmtDelBuku->rowCount();
                    } else {
                        $deletedCount = 0;
                    }

                    foreach ($filesToDelete as $file) {
                        if ($file) {
                            $deletedPath = __DIR__ . '/../assets/surat_keluar_pdf/' . basename((string)$file);
                            if (is_file($deletedPath)) {
                                @unlink($deletedPath);
                            }
                        }
                    }
                }
                $pdo->commit();
            } catch (Throwable $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                throw $e;
            }
            header('Location: surat-keluar.php?tahun=' . $tahunRange . '&ok=' . urlencode("Berhasil menghapus {$deletedCount} nomor dalam rentang {$nomorDari} s/d {$nomorSampai}."));
            exit;
        }

        if ($_POST['action'] === 'create_space_tanggal') {
            $tanggal = parseTanggalSurat($_POST['tanggal_surat'] ?? '');
            $jumlah = (int)($_POST['jumlah_space'] ?? 10);
            if ($tanggal === null) {
                throw new Exception('Tanggal surat wajib diisi dan harus valid.');
            }
            $tahunSpace = (int)date('Y', strtotime($tanggal));
            $result = createSuratKeluarSpaceByTanggal($pdo, $tahunSpace, $tanggal, $jumlah);
            header('Location: surat-keluar.php?tahun=' . $tahunSpace . '&ok=' . urlencode("Space berhasil dibuat: {$result['created']} nomor ({$result['first']} s/d {$result['last']})."));
            exit;
        }

        if ($_POST['action'] === 'reserve_slot') {
            $tanggal = parseTanggalSurat($_POST['tanggal_surat'] ?? '');
            $nomorUrut = (int)($_POST['nomor_urut'] ?? 0);
            $keterangan = trim((string)($_POST['keterangan'] ?? 'Slot kosong manual'));
            if ($tanggal === null) {
                throw new Exception('Tanggal surat wajib diisi dan harus valid.');
            }
            if ($nomorUrut <= 0) {
                throw new Exception('Nomor urut wajib diisi.');
            }

            $tahunSlot = (int)date('Y', strtotime($tanggal));
            $stmt = $pdo->prepare("SELECT id_surat_keluar FROM buku_nomor_surat_keluar WHERE tahun = ? AND nomor_urut = ? LIMIT 1");
            $stmt->execute([$tahunSlot, $nomorUrut]);
            $existing = $stmt->fetch(PDO::FETCH_ASSOC);
            if ($existing && !empty($existing['id_surat_keluar'])) {
                throw new Exception('Nomor sudah terpakai surat keluar.');
            }

            $stmt = $pdo->prepare("INSERT INTO buku_nomor_surat_keluar (tahun, nomor_urut, tanggal_surat, status, id_surat_keluar, keterangan)
                VALUES (:tahun, :nomor_urut, :tanggal_surat, 'kosong', NULL, :keterangan)
                ON DUPLICATE KEY UPDATE
                    tanggal_surat = VALUES(tanggal_surat),
                    status = 'kosong',
                    id_surat_keluar = NULL,
                    keterangan = VALUES(keterangan)");
            $stmt->execute([
                ':tahun' => $tahunSlot,
                ':nomor_urut' => $nomorUrut,
                ':tanggal_surat' => $tanggal,
                ':keterangan' => $keterangan !== '' ? $keterangan : 'Slot kosong manual',
            ]);
            header('Location: surat-keluar.php?tahun=' . $tahunSlot . '&ok=' . urlencode('Slot kosong berhasil disimpan.'));
            exit;
        }

        if ($_POST['action'] === 'import') {
            if (!isset($_FILES['file_import']) || $_FILES['file_import']['error'] !== UPLOAD_ERR_OK) {
                throw new Exception('File import wajib dipilih.');
            }
            $ext = strtolower(pathinfo($_FILES['file_import']['name'], PATHINFO_EXTENSION));
            if ($ext === 'csv') {
                $rows = csvRows($_FILES['file_import']['tmp_name']);
            } elseif ($ext === 'xlsx') {
                $rows = xlsxRows($_FILES['file_import']['tmp_name']);
            } else {
                throw new Exception('Format file belum didukung. Gunakan CSV atau XLSX.');
            }

            $data = rowsToSuratKeluarData($rows);
            $mode = (string)($_POST['mode_import'] ?? 'skip');
            $tahunImport = (int)($_POST['tahun_import'] ?? date('Y'));
            if ($tahunImport < 1900) {
                $tahunImport = (int)date('Y');
            }
            $imported = 0;
            $slotImported = 0;
            $skipped = 0;
            $pdo->beginTransaction();
            try {
                foreach ($data as $item) {
                    $tanggal = parseTanggalSurat($item['tanggal_surat']);
                    if ($item['nomor_urut'] <= 0) {
                        $skipped++;
                        continue;
                    }

                    $tahun = $tanggal !== null ? (int)date('Y', strtotime($tanggal)) : $tahunImport;
                    $slotKeterangan = trim((string)($item['keterangan'] ?? ''));
                    if ($slotKeterangan === '') {
                        $slotKeterangan = 'Import baris ' . (int)($item['source_row'] ?? 0);
                    }

                    $stmt = $pdo->prepare("SELECT id_surat_keluar FROM buku_nomor_surat_keluar WHERE tahun = ? AND nomor_urut = ? LIMIT 1");
                    $stmt->execute([$tahun, $item['nomor_urut']]);
                    $slot = $stmt->fetch(PDO::FETCH_ASSOC);
                    if ($slot && !empty($slot['id_surat_keluar']) && $mode !== 'replace') {
                        $skipped++;
                        continue;
                    }

                    $stmt = $pdo->prepare("INSERT INTO buku_nomor_surat_keluar (tahun, nomor_urut, tanggal_surat, status, id_surat_keluar, keterangan)
                        VALUES (:tahun, :nomor_urut, :tanggal_surat, 'kosong', NULL, :keterangan)
                        ON DUPLICATE KEY UPDATE
                            tanggal_surat = COALESCE(VALUES(tanggal_surat), tanggal_surat),
                            keterangan = VALUES(keterangan)");
                    $stmt->execute([
                        ':tahun' => $tahun,
                        ':nomor_urut' => (int)$item['nomor_urut'],
                        ':tanggal_surat' => $tanggal,
                        ':keterangan' => $slotKeterangan,
                    ]);
                    $slotImported++;

                    $nomorSurat = trim((string)($item['nomor_surat'] ?? ''));
                    if ($nomorSurat !== '' && $tanggal !== null) {
                        $stmt = $pdo->prepare("SELECT id FROM surat_keluar WHERE tahun = ? AND (nomor_urut = ? OR nomor_surat = ?) LIMIT 1");
                        $stmt->execute([$tahun, $item['nomor_urut'], $nomorSurat]);
                        $existingId = $stmt->fetchColumn();
                        if ($existingId && $mode !== 'replace') {
                            continue;
                        }

                        $id = saveImportedSuratKeluar($pdo, $item, $existingId ? (int)$existingId : null);
                        bindNomorKeSuratKeluar($pdo, $id, (int)$item['nomor_urut'], $tanggal);
                        $imported++;
                    }
                }
                syncSuratKeluarTerpakaiKeBukuNomor($pdo);
                $pdo->commit();
            } catch (Throwable $e) {
                if ($pdo->inTransaction()) {
                    $pdo->rollBack();
                }
                throw $e;
            }
            header('Location: surat-keluar.php?ok=' . urlencode("Import selesai. {$slotImported} slot diproses, {$imported} surat terisi, {$skipped} dilewati."));
            exit;
        }
    } catch (Throwable $e) {
        header('Location: surat-keluar.php?error=' . urlencode($e->getMessage()));
        exit;
    }
}

$tahun = isset($_GET['tahun']) ? (int)$_GET['tahun'] : (int)date('Y');
$search = trim((string)($_GET['search'] ?? ''));
$tanggalDari = parseTanggalSurat($_GET['tanggal_dari'] ?? '') ?? '';
$tanggalSampai = parseTanggalSurat($_GET['tanggal_sampai'] ?? '') ?? '';
$page = max(1, (int)($_GET['page'] ?? 1));
$limit = 25;

$where = ['b.tahun = :tahun'];
$params = [':tahun' => $tahun];
if ($search !== '') {
    $where[] = "(sk.nomor_surat LIKE :search OR sk.tujuan_surat LIKE :search OR sk.perihal LIKE :search OR COALESCE(sk.keterangan, '') LIKE :search OR COALESCE(b.keterangan, '') LIKE :search)";
    $params[':search'] = '%' . $search . '%';
}
if ($tanggalDari !== '') {
    $where[] = "b.tanggal_surat >= :tanggal_dari";
    $params[':tanggal_dari'] = $tanggalDari;
}
if ($tanggalSampai !== '') {
    $where[] = "b.tanggal_surat <= :tanggal_sampai";
    $params[':tanggal_sampai'] = $tanggalSampai;
}
$whereSql = implode(' AND ', $where);

$stmt = $pdo->prepare("SELECT COUNT(*)
    FROM buku_nomor_surat_keluar b
    LEFT JOIN surat_keluar sk ON sk.id = b.id_surat_keluar
    WHERE {$whereSql}");
$stmt->execute($params);
$totalRecords = (int)$stmt->fetchColumn();
$totalPages = max(1, (int)ceil($totalRecords / $limit));
$page = min($page, $totalPages);
$offset = ($page - 1) * $limit;

$stmt = $pdo->prepare("SELECT
        sk.*,
        b.id AS buku_id,
        b.tahun,
        b.nomor_urut,
        b.tanggal_surat AS tanggal_slot,
        b.status,
        b.keterangan AS keterangan_slot
    FROM buku_nomor_surat_keluar b
    LEFT JOIN surat_keluar sk ON sk.id = b.id_surat_keluar
    WHERE {$whereSql}
    ORDER BY b.nomor_urut ASC
    LIMIT {$offset}, {$limit}");
$stmt->execute($params);
$suratKeluar = $stmt->fetchAll();
$nextNomor = getNextSuratKeluarNomor($pdo, $tahun);
$lastTanggalTahun = getLastTanggalSuratKeluar($pdo, $tahun);

include '../includes/header.php';
?>

<style>
    .line-clamp-2 {
        display: -webkit-box;
        -webkit-line-clamp: 2;
        -webkit-box-orient: vertical;
        overflow: hidden;
    }
    .sticky-shadow {
        box-shadow: -4px 0 6px -2px rgba(0, 0, 0, 0.08);
    }
</style>

<div class="space-y-6">
    <div class="flex flex-col lg:flex-row lg:items-center lg:justify-between gap-4">
        <div>
            <h1 class="page-header-title">Penomoran Surat Keluar</h1>
            <p class="page-header-sub">Kelola nomor, tujuan, perihal, dan keterangan surat keluar per tahun</p>
        </div>
        <div class="flex items-center gap-2 relative">
            <!-- Dropdown trigger -->
            <div class="relative inline-block text-left" id="headerDropdownContainer">
                <button type="button" id="btnHeaderDropdown" class="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg shadow-sm hover:bg-gray-50 transition-colors">
                    Lainnya ▾
                </button>
                <!-- Dropdown Menu -->
                <div id="headerDropdownMenu" class="absolute right-0 mt-2 w-52 rounded-lg shadow-lg bg-white border border-gray-100 divide-y divide-gray-100 z-50 hidden">
                    <div class="py-1">
                        <button onclick="openModal('spaceModal'); document.getElementById('headerDropdownMenu').classList.add('hidden')" class="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center gap-2 transition-colors">
                            <i class='bx bx-grid-alt text-gray-400'></i> Buat Space
                        </button>
                        <button onclick="openReserveSlotModal(); document.getElementById('headerDropdownMenu').classList.add('hidden')" class="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center gap-2 transition-colors">
                            <i class='bx bx-plus-circle text-gray-400'></i> Tambah Slot
                        </button>
                        <button onclick="openModal('importModal'); document.getElementById('headerDropdownMenu').classList.add('hidden')" class="w-full text-left px-4 py-2 text-sm text-gray-700 hover:bg-gray-50 flex items-center gap-2 transition-colors">
                            <i class='bx bx-upload text-gray-400'></i> Import Data
                        </button>
                    </div>
                    <div class="py-1">
                        <button onclick="openDeleteRangeModal(); document.getElementById('headerDropdownMenu').classList.add('hidden')" class="w-full text-left px-4 py-2 text-sm text-red-600 hover:bg-red-50 flex items-center gap-2 transition-colors">
                            <i class='bx bx-trash text-red-500'></i> Hapus Rentang Nomor
                        </button>
                    </div>
                </div>
            </div>
            
            <!-- Primary Button -->
            <button onclick="openAddModal()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg inline-flex items-center gap-2 transition-colors font-medium text-sm">
                ＋ Tambah Surat
            </button>
        </div>
    </div>

    <?php if (isset($_GET['ok'])): ?>
        <div class="bg-emerald-50 border border-emerald-200 text-emerald-700 px-4 py-3 rounded-lg text-sm"><?php echo htmlspecialchars((string)$_GET['ok']); ?></div>
    <?php endif; ?>
    <?php if (isset($_GET['error'])): ?>
        <div class="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-lg text-sm"><?php echo htmlspecialchars((string)$_GET['error']); ?></div>
    <?php endif; ?>

    <div class="grid grid-cols-1 sm:grid-cols-4 gap-4">
        <!-- Card 1: Tahun -->
        <div class="stat-card blue">
            <div class="flex justify-between items-center mb-1">
                <span class="text-xs font-semibold text-gray-500 uppercase tracking-wide">Tahun</span>
                <div class="stat-icon blue">
                    <i class='bx bx-calendar'></i>
                </div>
            </div>
            <p class="text-3xl font-bold text-gray-900 mt-2"><?php echo $tahun; ?></p>
        </div>

        <!-- Card 2: Total Data -->
        <div class="stat-card indigo">
            <div class="flex justify-between items-center mb-1">
                <span class="text-xs font-semibold text-gray-500 uppercase tracking-wide">Total Data</span>
                <div class="stat-icon indigo">
                    <i class='bx bx-file'></i>
                </div>
            </div>
            <p class="text-3xl font-bold text-gray-900 mt-2"><?php echo $totalRecords; ?></p>
        </div>

        <!-- Card 3: Nomor Urut Berikutnya -->
        <div class="stat-card violet">
            <div class="flex justify-between items-center mb-1">
                <span class="text-xs font-semibold text-gray-500 uppercase tracking-wide">Nomor Urut Berikutnya</span>
                <div class="stat-icon violet">
                    <i class='bx bx-hash'></i>
                </div>
            </div>
            <p class="text-3xl font-bold text-gray-900 mt-2"><?php echo $nextNomor; ?></p>
        </div>

        <!-- Card 4: Tanggal Terakhir Buku -->
        <div class="stat-card teal">
            <div class="flex justify-between items-center mb-1">
                <span class="text-xs font-semibold text-gray-500 uppercase tracking-wide">Tanggal Terakhir Buku</span>
                <div class="stat-icon teal">
                    <i class='bx bx-time-five'></i>
                </div>
            </div>
            <p class="text-3xl font-bold text-gray-900 mt-2"><?php echo $lastTanggalTahun ? date('d/m/Y', strtotime($lastTanggalTahun)) : '-'; ?></p>
        </div>
    </div>

    <div class="card p-4">
        <form method="GET" class="grid grid-cols-1 md:grid-cols-6 gap-3">
            <div>
                <label class="block text-xs font-medium text-gray-500 mb-1">Tahun</label>
                <input type="number" name="tahun" value="<?php echo $tahun; ?>" class="input-premium py-1.5">
            </div>
            <div>
                <label class="block text-xs font-medium text-gray-500 mb-1">Tanggal Dari</label>
                <input type="date" name="tanggal_dari" value="<?php echo htmlspecialchars($tanggalDari); ?>" class="input-premium py-1.5">
            </div>
            <div>
                <label class="block text-xs font-medium text-gray-500 mb-1">Tanggal Sampai</label>
                <input type="date" name="tanggal_sampai" value="<?php echo htmlspecialchars($tanggalSampai); ?>" class="input-premium py-1.5">
            </div>
            <div>
                <label class="block text-xs font-medium text-gray-500 mb-1">Cari</label>
                <input type="text" name="search" value="<?php echo htmlspecialchars($search); ?>" placeholder="Nomor, tujuan, perihal..." class="input-premium py-1.5">
            </div>
            <div>
                <label class="block text-xs font-medium text-gray-500 mb-1">Status</label>
                <select name="status" class="input-premium py-1.5">
                    <option value="" <?php echo !isset($_GET['status']) || $_GET['status'] === '' ? 'selected' : ''; ?>>Semua Status</option>
                    <option value="terisi" <?php echo isset($_GET['status']) && $_GET['status'] === 'terisi' ? 'selected' : ''; ?>>Lengkap</option>
                    <option value="kosong" <?php echo isset($_GET['status']) && $_GET['status'] === 'kosong' ? 'selected' : ''; ?>>Belum Lengkap</option>
                </select>
            </div>
            <div class="flex items-end gap-2">
                <button type="submit" class="w-full px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg transition-colors font-medium text-sm">Tampilkan</button>
                <a href="surat-keluar.php?tahun=<?php echo date('Y'); ?>" class="w-full px-4 py-2 border border-gray-200 text-gray-700 bg-white hover:bg-gray-50 rounded-lg inline-flex items-center justify-center transition-colors text-sm font-medium">Reset</a>
            </div>
        </form>
    </div>

    <div class="card overflow-hidden">
        <div class="overflow-x-auto">
            <table class="w-full text-left table-premium">
                <thead>
                    <tr>
                        <th class="w-10 px-4 py-3 text-center">
                            <input type="checkbox" id="selectAll" class="w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer" title="Pilih Semua di Halaman Ini">
                        </th>
                        <th class="w-12">No</th>
                        <th class="min-w-48">No. Surat</th>
                        <th class="min-w-36">Tanggal Surat</th>
                        <th class="min-w-64">Tujuan Surat</th>
                        <th class="min-w-64">Perihal</th>
                        <th class="text-right">Aksi</th>
                    </tr>
                </thead>
                <tbody class="bg-white divide-y divide-gray-200">
                    <?php if (count($suratKeluar) > 0): ?>
                        <?php foreach ($suratKeluar as $row): ?>
                            <?php
                                $statusLabel = $row['status'] === 'terisi' ? 'Lengkap' : 'Kosong';
                                $statusClass = 'bg-blue-100 text-blue-700';
                                if ($statusLabel === 'Lengkap') {
                                    $statusClass = 'bg-green-100 text-green-700';
                                } elseif ($statusLabel === 'Pending') {
                                    $statusClass = 'bg-yellow-100 text-yellow-700';
                                } elseif ($statusLabel === 'Error') {
                                    $statusClass = 'bg-red-100 text-red-700';
                                }

                                // PDF detection logic
                                $isPdf = false;
                                $pdfFile = '';
                                if (!empty($row['file_pdf'])) {
                                    $isPdf = true;
                                    $pdfFile = $row['file_pdf'];
                                } elseif (!empty($row['keterangan']) && preg_match('/\.pdf$/i', trim((string)$row['keterangan']))) {
                                    $isPdf = true;
                                    $pdfFile = trim((string)$row['keterangan']);
                                }
                            ?>
                            <tr class="border-b border-gray-200 hover:bg-gray-50 group transition-colors">
                                <td class="px-4 py-3 align-middle text-center w-10">
                                    <input type="checkbox"
                                           class="row-checkbox w-4 h-4 rounded border-gray-300 text-indigo-600 focus:ring-indigo-500 cursor-pointer"
                                           value="<?php echo (int)$row['buku_id']; ?>"
                                           data-nomor="<?php echo (int)$row['nomor_urut']; ?>"
                                           data-has-surat="<?php echo !empty($row['id']) ? '1' : '0'; ?>">
                                </td>
                                <td class="px-4 py-3 align-middle text-sm font-semibold text-gray-800 w-12"><?php echo (int)$row['nomor_urut']; ?></td>
                                <td class="px-4 py-3 align-middle text-sm">
                                    <div class="flex flex-col gap-1">
                                        <?php if (!empty($row['nomor_surat'])): ?>
                                            <span class="text-gray-800 font-medium"><?php echo htmlspecialchars($row['nomor_surat']); ?></span>
                                        <?php else: ?>
                                            <span class="text-gray-400 font-medium">-</span>
                                        <?php endif; ?>
                                        <div>
                                            <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium <?php echo $statusClass; ?>">
                                                <?php echo $statusLabel; ?>
                                            </span>
                                        </div>
                                    </div>
                                </td>
                                <td class="px-4 py-3 align-middle text-sm text-gray-600"><?php echo !empty($row['tanggal_slot']) ? date('d/m/Y', strtotime($row['tanggal_slot'])) : '-'; ?></td>
                                <td class="px-4 py-3 align-middle text-sm <?php echo !empty($row['tujuan_surat']) ? 'text-gray-700' : 'text-gray-500'; ?> whitespace-pre-line"><?php echo !empty($row['tujuan_surat']) ? htmlspecialchars($row['tujuan_surat']) : '-'; ?></td>
                                <td class="px-4 py-3 align-middle text-sm <?php echo !empty($row['perihal']) ? 'text-gray-700' : 'text-gray-500'; ?>">
                                    <?php if (!empty($row['perihal'])): ?>
                                        <div class="line-clamp-2 text-sm" title="<?php echo htmlspecialchars($row['perihal']); ?>">
                                            <?php echo htmlspecialchars($row['perihal']); ?>
                                        </div>
                                    <?php else: ?>
                                        <span class="text-gray-500">-</span>
                                    <?php endif; ?>
                                </td>
                                   <td class="px-4 py-3 align-middle text-right sticky right-0 bg-white group-hover:bg-gray-50 transition-colors sticky-shadow">
                                       <div class="inline-flex items-center justify-end gap-3">
                                           <?php if (!empty($row['id'])): ?>
                                               <?php if ($isPdf): ?>
                                               <a href="../assets/surat_keluar_pdf/<?php echo htmlspecialchars(basename((string)$pdfFile), ENT_QUOTES); ?>"
                                                  target="_blank"
                                                  class="text-blue-500 hover:text-blue-600 transition-colors"
                                                  title="Lihat PDF">
                                                   <i class='bx bx-file text-xl'></i>
                                               </a>
                                               <?php endif; ?>
                                               
                                               <button type="button"
                                                   class="edit-btn text-amber-500 hover:text-amber-600 transition-colors"
                                                   title="Edit"
                                                   data-id="<?php echo (int)$row['id']; ?>"
                                                   data-nomor-urut="<?php echo (int)$row['nomor_urut']; ?>"
                                                   data-nomor-surat="<?php echo htmlspecialchars($row['nomor_surat'], ENT_QUOTES); ?>"
                                                   data-tanggal="<?php echo htmlspecialchars($row['tanggal_slot'], ENT_QUOTES); ?>"
                                                   data-tujuan="<?php echo htmlspecialchars($row['tujuan_surat'], ENT_QUOTES); ?>"
                                                   data-perihal="<?php echo htmlspecialchars($row['perihal'], ENT_QUOTES); ?>"
                                                   data-keterangan="<?php echo htmlspecialchars((string)$row['keterangan'], ENT_QUOTES); ?>"
                                                   data-file-pdf="<?php echo htmlspecialchars((string)($row['file_pdf'] ?? ''), ENT_QUOTES); ?>">
                                                   <i class='bx bx-edit text-xl'></i>
                                               </button>
                                               
                                               <form method="POST" class="inline" onsubmit="return confirm('Hapus data surat keluar ini?');">
                                                   <input type="hidden" name="action" value="delete">
                                                   <input type="hidden" name="id" value="<?php echo (int)$row['id']; ?>">
                                                   <button type="submit" class="text-red-500 hover:text-red-600 transition-colors" title="Hapus">
                                                       <i class='bx bx-trash text-xl'></i>
                                                   </button>
                                               </form>
                                           <?php else: ?>
                                               <button type="button"
                                                   class="fill-slot-btn text-amber-500 hover:text-amber-600 transition-colors"
                                                   title="Isi/Edit Slot"
                                                   data-buku-id="<?php echo (int)$row['buku_id']; ?>"
                                                   data-nomor-urut="<?php echo (int)$row['nomor_urut']; ?>"
                                                   data-tanggal="<?php echo htmlspecialchars((string)$row['tanggal_slot'], ENT_QUOTES); ?>"
                                                   data-keterangan="<?php echo htmlspecialchars((string)$row['keterangan_slot'], ENT_QUOTES); ?>">
                                                   <i class='bx bx-edit text-xl'></i>
                                               </button>
                                               
                                               <form method="POST" class="inline" onsubmit="return confirm('Hapus slot kosong ini?');">
                                                   <input type="hidden" name="action" value="delete_slot">
                                                   <input type="hidden" name="id" value="<?php echo (int)$row['buku_id']; ?>">
                                                   <button type="submit" class="text-red-500 hover:text-red-600 transition-colors" title="Hapus Slot">
                                                       <i class='bx bx-trash text-xl'></i>
                                                   </button>
                                               </form>
                                           <?php endif; ?>
                                       </div>
                                    </td>
                            </tr>
                        <?php endforeach; ?>
                    <?php else: ?>
                        <tr>
                            <td colspan="7" class="px-4 py-10 align-middle text-center text-gray-500">Belum ada data surat keluar.</td>
                        </tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>

        <?php
        $query = ['tahun' => $tahun, 'search' => $search];
        if ($tanggalDari !== '') $query['tanggal_dari'] = $tanggalDari;
        if ($tanggalSampai !== '') $query['tanggal_sampai'] = $tanggalSampai;
        $basePageUrl = 'surat-keluar.php?' . http_build_query($query, '', '&', PHP_QUERY_RFC3986);
        ?>
        <div class="px-4 py-3 border-t border-slate-100 flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 text-sm text-slate-600">
            <span>
              Menampilkan <?php echo $totalRecords > 0 ? $offset + 1 : 0; ?> – <?php echo min($offset + $limit, $totalRecords); ?> dari <?php echo $totalRecords; ?> data
            </span>
            <?php if ($totalPages > 1): ?>
                <div class="flex flex-wrap items-center gap-1">
                    <!-- Tombol Pertama -->
                    <?php if ($page > 1): ?>
                        <a class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm inline-flex items-center" 
                           href="<?php echo htmlspecialchars($basePageUrl . '&page=1'); ?>" 
                           title="Halaman Pertama">
                           <i class='bx bx-chevrons-left text-base'></i><span class="hidden sm:inline ml-1 font-medium">Pertama</span>
                        </a>
                        <a class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm inline-flex items-center" 
                           href="<?php echo htmlspecialchars($basePageUrl . '&page=' . ($page - 1)); ?>" 
                           title="Sebelumnya">
                           <i class='bx bx-chevron-left text-base'></i><span class="hidden sm:inline ml-1 font-medium">Sebelumnya</span>
                        </a>
                    <?php else: ?>
                        <span class="px-3 py-1 rounded border border-slate-100 text-slate-300 text-sm inline-flex items-center cursor-not-allowed">
                           <i class='bx bx-chevrons-left text-base'></i><span class="hidden sm:inline ml-1 font-medium">Pertama</span>
                        </span>
                        <span class="px-3 py-1 rounded border border-slate-100 text-slate-300 text-sm inline-flex items-center cursor-not-allowed">
                           <i class='bx bx-chevron-left text-base'></i><span class="hidden sm:inline ml-1 font-medium">Sebelumnya</span>
                        </span>
                    <?php endif; ?>

                    <!-- Nomor Halaman -->
                    <?php
                    $maxLinks = 5;
                    $half = (int)floor($maxLinks / 2);
                    $from = max(1, $page - $half);
                    $to = min($totalPages, $from + $maxLinks - 1);
                    $from = max(1, $to - $maxLinks + 1);
                    
                    if ($from > 1):
                    ?>
                        <a class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm" href="<?php echo htmlspecialchars($basePageUrl . '&page=1'); ?>">1</a>
                        <?php if ($from > 2): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                    <?php endif; ?>

                    <?php for ($i = $from; $i <= $to; $i++): ?>
                        <a class="px-3 py-1 rounded border <?php echo $i === $page ? 'bg-indigo-600 text-white border-indigo-600 font-semibold' : 'border-slate-200 text-slate-600 hover:bg-slate-50'; ?>" 
                           href="<?php echo htmlspecialchars($basePageUrl . '&page=' . $i); ?>">
                           <?php echo $i; ?>
                        </a>
                    <?php endfor; ?>

                    <?php if ($to < $totalPages): ?>
                        <?php if ($to < $totalPages - 1): ?>
                            <span class="px-2 py-1 text-slate-400 text-sm">...</span>
                        <?php endif; ?>
                        <a class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm" href="<?php echo htmlspecialchars($basePageUrl . '&page=' . $totalPages); ?>"><?php echo $totalPages; ?></a>
                    <?php endif; ?>

                    <!-- Tombol Berikutnya / Terakhir -->
                    <?php if ($page < $totalPages): ?>
                        <a class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm inline-flex items-center" 
                           href="<?php echo htmlspecialchars($basePageUrl . '&page=' . ($page + 1)); ?>" 
                           title="Berikutnya">
                           <span class="hidden sm:inline mr-1 font-medium">Berikutnya</span><i class='bx bx-chevron-right text-base'></i>
                        </a>
                        <a class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm inline-flex items-center" 
                           href="<?php echo htmlspecialchars($basePageUrl . '&page=' . $totalPages); ?>" 
                           title="Halaman Terakhir">
                           <span class="hidden sm:inline mr-1 font-medium">Terakhir</span><i class='bx bx-chevrons-right text-base'></i>
                        </a>
                    <?php else: ?>
                        <span class="px-3 py-1 rounded border border-slate-100 text-slate-300 text-sm inline-flex items-center cursor-not-allowed">
                           <span class="hidden sm:inline mr-1 font-medium">Berikutnya</span><i class='bx bx-chevron-right text-base'></i>
                        </span>
                        <span class="px-3 py-1 rounded border border-slate-100 text-slate-300 text-sm inline-flex items-center cursor-not-allowed">
                           <span class="hidden sm:inline mr-1 font-medium">Terakhir</span><i class='bx bx-chevrons-right text-base'></i>
                        </span>
                    <?php endif; ?>
                </div>
            <?php endif; ?>
        </div>
    </div>
</div>

<!-- Floating Bulk Action Bar -->
<div id="bulkActionBar" class="fixed bottom-6 left-1/2 transform -translate-x-1/2 bg-slate-900 text-white px-5 py-3 rounded-2xl shadow-2xl flex items-center gap-4 z-40 hidden border border-slate-700">
    <div class="flex items-center gap-2">
        <span class="w-2.5 h-2.5 rounded-full bg-indigo-400 animate-pulse"></span>
        <span id="selectedCount" class="font-medium text-sm">0 nomor dipilih</span>
    </div>
    <div class="h-4 w-px bg-slate-700"></div>
    <div class="flex items-center gap-2">
        <button type="button" onclick="confirmBulkDelete()" class="bg-red-600 hover:bg-red-700 text-white px-3.5 py-1.5 rounded-lg text-sm font-medium transition-colors inline-flex items-center gap-1.5 shadow-sm">
            <i class='bx bx-trash text-base'></i> Hapus Terpilih
        </button>
        <button type="button" onclick="clearSelection()" class="text-slate-400 hover:text-slate-200 text-sm font-medium transition-colors px-2 py-1">
            Batal
        </button>
    </div>
</div>

<div id="formModal" class="fixed inset-0 z-50 hidden overflow-y-auto">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 z-0" onclick="closeModal('formModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="relative z-10 inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl sm:my-8 sm:align-middle sm:max-w-2xl sm:w-full max-h-[90vh] overflow-y-auto">
            <form method="POST" enctype="multipart/form-data" class="p-6 space-y-4">
                <div class="flex items-center justify-between">
                    <h3 id="formTitle" class="text-lg font-semibold text-slate-900">Tambah Surat Keluar</h3>
                    <button type="button" onclick="closeModal('formModal')" class="text-slate-400 hover:text-slate-600"><i class='bx bx-x text-2xl'></i></button>
                </div>
                <input type="hidden" name="action" id="formAction" value="create">
                <input type="hidden" name="id" id="formId">
                <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">No</label>
                        <select name="nomor_urut" id="formNomorUrut" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                            <option value="">Pilih tanggal terlebih dahulu</option>
                        </select>
                        <p id="nomorSlotInfo" class="text-xs text-slate-500 mt-1">Nomor diambil dari Buku Nomor Surat Keluar.</p>
                    </div>
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Tanggal Surat</label>
                        <input type="date" name="tanggal_surat" id="formTanggal" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                    </div>
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">No. Surat</label>
                    <input type="text" name="nomor_surat" id="formNomorSurat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg" placeholder="Contoh: 500.4/053/123.6.6/2026">
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Tujuan Surat</label>
                    <textarea name="tujuan_surat" id="formTujuan" rows="3" required class="w-full px-3 py-2 border border-slate-300 rounded-lg"></textarea>
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Perihal</label>
                    <textarea name="perihal" id="formPerihal" rows="3" required class="w-full px-3 py-2 border border-slate-300 rounded-lg"></textarea>
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Keterangan</label>
                    <textarea name="keterangan" id="formKeterangan" rows="2" class="w-full px-3 py-2 border border-slate-300 rounded-lg"></textarea>
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">File PDF <span class="text-slate-400">(opsional)</span></label>
                    <input type="file" name="file_pdf" id="formFilePdf" accept="application/pdf,.pdf" class="w-full px-3 py-2 border border-slate-300 rounded-lg bg-white">
                    <p id="currentPdfInfo" class="text-xs text-slate-500 mt-1">Kosongkan jika tidak ada lampiran PDF.</p>
                </div>
                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeModal('formModal')" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg">Simpan</button>
                </div>
            </form>
        </div>
    </div>
</div>

<div id="spaceModal" class="fixed inset-0 z-50 hidden overflow-y-auto">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 z-0" onclick="closeModal('spaceModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="relative z-10 inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <form method="POST" class="p-6 space-y-4">
                <div class="flex items-center justify-between">
                    <h3 class="text-lg font-semibold text-slate-900">Buat Space Nomor Surat Keluar</h3>
                    <button type="button" onclick="closeModal('spaceModal')" class="text-slate-400 hover:text-slate-600"><i class='bx bx-x text-2xl'></i></button>
                </div>
                <input type="hidden" name="action" value="create_space_tanggal">
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Tanggal Surat</label>
                    <input type="date" name="tanggal_surat" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Jumlah Space</label>
                    <input type="number" name="jumlah_space" value="10" min="1" max="100" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                </div>
                <p class="text-xs text-slate-500">Nomor dibuat berurutan setelah nomor terakhir tahun tersebut.</p>
                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeModal('spaceModal')" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg">Buat Space</button>
                </div>
            </form>
        </div>
    </div>
</div>

<div id="reserveSlotModal" class="fixed inset-0 z-50 hidden overflow-y-auto">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 z-0" onclick="closeModal('reserveSlotModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="relative z-10 inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <form method="POST" class="p-6 space-y-4">
                <div class="flex items-center justify-between">
                    <h3 class="text-lg font-semibold text-slate-900">Tambah Slot Kosong</h3>
                    <button type="button" onclick="closeModal('reserveSlotModal')" class="text-slate-400 hover:text-slate-600"><i class='bx bx-x text-2xl'></i></button>
                </div>
                <input type="hidden" name="action" value="reserve_slot">
                <div>
                    <label class="block text-sm text-slate-600 mb-1">No</label>
                    <input type="number" name="nomor_urut" id="reserveNomorUrut" min="1" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Tanggal Surat</label>
                    <input type="date" name="tanggal_surat" id="reserveTanggal" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Keterangan</label>
                    <input type="text" name="keterangan" value="Slot kosong manual" class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                </div>
                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeModal('reserveSlotModal')" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-slate-700 hover:bg-slate-800 text-white rounded-lg">Simpan Slot</button>
                </div>
            </form>
        </div>
    </div>
</div>

<div id="importModal" class="fixed inset-0 z-50 hidden overflow-y-auto">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 z-0" onclick="closeModal('importModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="relative z-10 inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl sm:my-8 sm:align-middle sm:max-w-xl sm:w-full">
            <form method="POST" enctype="multipart/form-data" class="p-6 space-y-4">
                <div class="flex items-center justify-between">
                    <h3 class="text-lg font-semibold text-slate-900">Import Data Surat Keluar</h3>
                    <button type="button" onclick="closeModal('importModal')" class="text-slate-400 hover:text-slate-600"><i class='bx bx-x text-2xl'></i></button>
                </div>
                <input type="hidden" name="action" value="import">
                <div class="bg-slate-50 border border-slate-200 rounded-lg p-3 text-sm text-slate-600">
                    Header yang dibaca: <strong>No</strong>, <strong>No. Surat</strong>, <strong>Tanggal Surat</strong>, <strong>Tujuan Surat</strong>, <strong>Perihal</strong>, <strong>Keterangan</strong>. Hanya kolom <strong>No</strong> yang wajib; baris dengan kolom lain kosong tetap dibuat sebagai slot.
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">File CSV/XLSX</label>
                    <input type="file" name="file_import" accept=".csv,.xlsx" required class="w-full px-3 py-2 border border-slate-300 rounded-lg bg-white">
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Tahun untuk baris tanpa tanggal</label>
                    <input type="number" name="tahun_import" value="<?php echo $tahun; ?>" min="1900" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                </div>
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Jika nomor sudah ada</label>
                    <select name="mode_import" class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                        <option value="skip">Lewati data tersebut</option>
                        <option value="replace">Perbarui data lama</option>
                    </select>
                </div>
                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeModal('importModal')" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg">Import</button>
                </div>
            </form>
        </div>
    </div>
</div>

<!-- Modal Hapus Terpilih (Bulk Delete) -->
<div id="bulkDeleteModal" class="fixed inset-0 z-50 hidden overflow-y-auto">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 z-0" onclick="closeModal('bulkDeleteModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="relative z-10 inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl sm:my-8 sm:align-middle sm:max-w-md sm:w-full">
            <form method="POST" class="p-6 space-y-4" id="formBulkDelete">
                <input type="hidden" name="action" value="bulk_delete">
                <input type="hidden" name="current_tahun" value="<?php echo $tahun; ?>">
                <div id="bulkDeleteHiddenInputs"></div>

                <div class="flex items-center gap-3 text-red-600">
                    <div class="w-10 h-10 rounded-full bg-red-100 flex items-center justify-center flex-shrink-0">
                        <i class='bx bx-trash text-2xl'></i>
                    </div>
                    <div>
                        <h3 class="text-lg font-semibold text-slate-900">Hapus Banyak Nomor</h3>
                        <p class="text-xs text-slate-500">Konfirmasi penghapusan data terpilih</p>
                    </div>
                </div>

                <div class="bg-red-50 border border-red-200 text-red-800 rounded-lg p-3 text-sm space-y-2">
                    <p id="bulkDeleteSummary">Anda akan menghapus <strong>0 nomor</strong> yang dipilih.</p>
                    <p class="text-xs text-red-600" id="bulkDeleteWarning">Perhatian: Slot kosong akan dihapus permanen. Data surat keluar beserta file lampiran PDF juga akan dihapus jika termasuk dalam pilihan.</p>
                </div>

                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeModal('bulkDeleteModal')" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-sm font-medium">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-sm font-medium transition-colors inline-flex items-center gap-1.5">
                        <i class='bx bx-trash'></i> Ya, Hapus Sekarang
                    </button>
                </div>
            </form>
        </div>
    </div>
</div>

<!-- Modal Hapus Rentang Nomor (Delete Range) -->
<div id="deleteRangeModal" class="fixed inset-0 z-50 hidden overflow-y-auto">
    <div class="flex items-end justify-center min-h-screen pt-4 px-4 pb-20 text-center sm:block sm:p-0">
        <div class="fixed inset-0 bg-slate-900 bg-opacity-75 z-0" onclick="closeModal('deleteRangeModal')"></div>
        <span class="hidden sm:inline-block sm:align-middle sm:h-screen">&#8203;</span>
        <div class="relative z-10 inline-block align-bottom bg-white rounded-xl text-left overflow-hidden shadow-xl sm:my-8 sm:align-middle sm:max-w-lg sm:w-full">
            <form method="POST" class="p-6 space-y-4">
                <div class="flex items-center justify-between">
                    <h3 class="text-lg font-semibold text-slate-900">Hapus Rentang Nomor Surat Keluar</h3>
                    <button type="button" onclick="closeModal('deleteRangeModal')" class="text-slate-400 hover:text-slate-600"><i class='bx bx-x text-2xl'></i></button>
                </div>
                <input type="hidden" name="action" value="delete_range">
                
                <div>
                    <label class="block text-sm text-slate-600 mb-1">Tahun</label>
                    <input type="number" name="tahun" id="rangeTahun" value="<?php echo $tahun; ?>" min="1900" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                </div>

                <div class="grid grid-cols-2 gap-4">
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Dari No. Urut</label>
                        <input type="number" name="nomor_dari" min="1" placeholder="Contoh: 10" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                    </div>
                    <div>
                        <label class="block text-sm text-slate-600 mb-1">Sampai No. Urut</label>
                        <input type="number" name="nomor_sampai" min="1" placeholder="Contoh: 50" required class="w-full px-3 py-2 border border-slate-300 rounded-lg">
                    </div>
                </div>

                <div class="bg-amber-50 border border-amber-200 rounded-lg p-3 text-sm text-amber-800 space-y-2">
                    <label class="flex items-start gap-2 cursor-pointer">
                        <input type="checkbox" name="include_terisi" value="1" class="mt-1 rounded border-amber-300 text-red-600 focus:ring-red-500">
                        <span class="text-xs text-amber-900">
                            <strong>Sertakan juga nomor yang sudah terisi surat</strong> (Hati-hati: data surat dan file PDF dalam rentang tersebut akan ikut terhapus).
                            <br><span class="text-amber-700">Jika tidak dicentang, hanya slot yang masih kosong yang akan dihapus.</span>
                        </span>
                    </label>
                </div>

                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeModal('deleteRangeModal')" class="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-sm font-medium">Batal</button>
                    <button type="submit" class="px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-sm font-medium">Hapus Rentang</button>
                </div>
            </form>
        </div>
    </div>
</div>

<script>
    function openModal(id) {
        document.getElementById(id).classList.remove('hidden');
    }

    function closeModal(id) {
        document.getElementById(id).classList.add('hidden');
        if (id === 'formModal') {
            const tanggalInput = document.getElementById('formTanggal');
            const nomorSelect = document.getElementById('formNomorUrut');
            const fileInput = document.getElementById('formFilePdf');
            if (tanggalInput) tanggalInput.readOnly = false;
            if (nomorSelect) nomorSelect.disabled = false;
            if (fileInput) fileInput.value = '';
        }
    }

    function openDeleteRangeModal() {
        openModal('deleteRangeModal');
    }

    function todayYmd() {
        return new Date().toISOString().slice(0, 10);
    }

    async function loadNomorSuratKeluarKosong(selectedNomor = '') {
        const tanggalInput = document.getElementById('formTanggal');
        const nomorSelect = document.getElementById('formNomorUrut');
        const info = document.getElementById('nomorSlotInfo');
        if (!tanggalInput || !nomorSelect) return;

        const tanggal = tanggalInput.value;
        nomorSelect.innerHTML = '<option value="">Memuat nomor kosong...</option>';
        if (!tanggal) {
            nomorSelect.innerHTML = '<option value="">Pilih tanggal terlebih dahulu</option>';
            if (info) info.textContent = 'Nomor diambil dari Buku Nomor Surat Keluar.';
            return;
        }

        try {
            const response = await fetch(`get_nomor_surat_keluar_kosong.php?tanggal_surat=${encodeURIComponent(tanggal)}`);
            const data = await response.json();
            nomorSelect.innerHTML = '';

            if (selectedNomor) {
                nomorSelect.append(new Option(selectedNomor, selectedNomor, true, true));
            }

            if (data.success && Array.isArray(data.nomor) && data.nomor.length > 0) {
                if (!selectedNomor) {
                    nomorSelect.append(new Option('Pilih nomor kosong', ''));
                }
                data.nomor.forEach((item) => {
                    if (String(item.nomor_urut) !== String(selectedNomor)) {
                        nomorSelect.append(new Option(`${item.nomor_urut} - kosong`, item.nomor_urut));
                    }
                });
                if (info) {
                    const last = Number(data.last_nomor_tahun) || 0;
                    info.textContent = last > 0 ? `Nomor terakhir tahun ${data.tahun}: ${last}.` : `Belum ada nomor tahun ${data.tahun}.`;
                }
            } else if (!selectedNomor) {
                nomorSelect.append(new Option('Tidak ada nomor kosong untuk tanggal ini', ''));
                if (info) info.textContent = 'Buat Space atau Tambah Slot untuk tanggal ini terlebih dahulu.';
            }
        } catch (error) {
            nomorSelect.innerHTML = '<option value="">Gagal memuat nomor</option>';
            if (info) info.textContent = 'Gagal memuat nomor kosong.';
        }
    }

    function openReserveSlotModal() {
        document.getElementById('reserveNomorUrut').value = '<?php echo $nextNomor; ?>';
        document.getElementById('reserveTanggal').value = todayYmd();
        openModal('reserveSlotModal');
    }

    function openAddModal() {
        document.getElementById('formTitle').textContent = 'Tambah Surat Keluar';
        document.getElementById('formAction').value = 'create';
        document.getElementById('formId').value = '';
        document.getElementById('formNomorUrut').disabled = false;
        document.getElementById('formTanggal').value = todayYmd();
        document.getElementById('formTanggal').readOnly = false;
        document.getElementById('formNomorSurat').value = '';
        document.getElementById('formTujuan').value = '';
        document.getElementById('formPerihal').value = '';
        document.getElementById('formKeterangan').value = '';
        document.getElementById('formFilePdf').value = '';
        document.getElementById('currentPdfInfo').textContent = 'Kosongkan jika tidak ada lampiran PDF.';
        openModal('formModal');
        loadNomorSuratKeluarKosong();
    }

    document.querySelectorAll('.edit-btn').forEach((button) => {
        button.addEventListener('click', () => {
            document.getElementById('formTitle').textContent = 'Edit Surat Keluar';
            document.getElementById('formAction').value = 'update';
            document.getElementById('formId').value = button.dataset.id;
            document.getElementById('formTanggal').value = button.dataset.tanggal;
            document.getElementById('formTanggal').readOnly = false;
            document.getElementById('formNomorUrut').disabled = false;
            loadNomorSuratKeluarKosong(button.dataset.nomorUrut);
            document.getElementById('formNomorSurat').value = button.dataset.nomorSurat;
            document.getElementById('formTujuan').value = button.dataset.tujuan;
            document.getElementById('formPerihal').value = button.dataset.perihal;
            document.getElementById('formKeterangan').value = button.dataset.keterangan;
            document.getElementById('formFilePdf').value = '';
            document.getElementById('currentPdfInfo').innerHTML = button.dataset.filePdf
                ? `PDF saat ini: <a class="text-indigo-600 hover:text-indigo-700" target="_blank" href="../assets/surat_keluar_pdf/${encodeURIComponent(button.dataset.filePdf)}">lihat file</a>. Upload file baru untuk mengganti.`
                : 'Belum ada lampiran PDF. Kosongkan jika tidak ingin menambahkan.';
            openModal('formModal');
        });
    });

    document.querySelectorAll('.fill-slot-btn').forEach((button) => {
        button.addEventListener('click', () => {
            const nomorSelect = document.getElementById('formNomorUrut');
            const tanggalInput = document.getElementById('formTanggal');
            const tanggal = button.dataset.tanggal || todayYmd();

            document.getElementById('formTitle').textContent = 'Isi Slot Surat Keluar';
            document.getElementById('formAction').value = 'create';
            document.getElementById('formId').value = '';
            tanggalInput.value = tanggal;
            tanggalInput.readOnly = true;
            nomorSelect.disabled = false;
            nomorSelect.innerHTML = '';
            nomorSelect.append(new Option(button.dataset.nomorUrut, button.dataset.nomorUrut, true, true));
            document.getElementById('nomorSlotInfo').textContent = 'Mengisi slot nomor yang sudah dipilih dari tabel.';
            document.getElementById('formNomorSurat').value = '';
            document.getElementById('formTujuan').value = '';
            document.getElementById('formPerihal').value = '';
            document.getElementById('formKeterangan').value = button.dataset.keterangan || '';
            document.getElementById('formFilePdf').value = '';
            document.getElementById('currentPdfInfo').textContent = 'Kosongkan jika tidak ada lampiran PDF.';
            openModal('formModal');
        });
    });

    document.getElementById('formTanggal')?.addEventListener('change', () => {
        const isUpdate = document.getElementById('formAction')?.value === 'update';
        loadNomorSuratKeluarKosong(isUpdate ? document.getElementById('formNomorUrut')?.value : '');
    });

    // Header Dropdown Toggle
    document.getElementById('btnHeaderDropdown')?.addEventListener('click', function(e) {
        e.stopPropagation();
        document.getElementById('headerDropdownMenu')?.classList.toggle('hidden');
    });
    document.addEventListener('click', function() {
        document.getElementById('headerDropdownMenu')?.classList.add('hidden');
    });

    // Checkbox and Bulk Actions Logic
    const selectAllCheckbox = document.getElementById('selectAll');
    const rowCheckboxes = document.querySelectorAll('.row-checkbox');
    const bulkActionBar = document.getElementById('bulkActionBar');
    const selectedCountSpan = document.getElementById('selectedCount');

    function updateBulkActionBar() {
        const checkedBoxes = document.querySelectorAll('.row-checkbox:checked');
        const count = checkedBoxes.length;
        if (count > 0) {
            bulkActionBar.classList.remove('hidden');
            selectedCountSpan.textContent = `${count} nomor dipilih`;
        } else {
            bulkActionBar.classList.add('hidden');
        }

        if (selectAllCheckbox) {
            selectAllCheckbox.checked = rowCheckboxes.length > 0 && count === rowCheckboxes.length;
            selectAllCheckbox.indeterminate = count > 0 && count < rowCheckboxes.length;
        }
    }

    if (selectAllCheckbox) {
        selectAllCheckbox.addEventListener('change', function() {
            rowCheckboxes.forEach(cb => {
                cb.checked = selectAllCheckbox.checked;
            });
            updateBulkActionBar();
        });
    }

    rowCheckboxes.forEach(cb => {
        cb.addEventListener('change', updateBulkActionBar);
    });

    function clearSelection() {
        rowCheckboxes.forEach(cb => cb.checked = false);
        if (selectAllCheckbox) {
            selectAllCheckbox.checked = false;
            selectAllCheckbox.indeterminate = false;
        }
        updateBulkActionBar();
    }

    function confirmBulkDelete() {
        const checkedBoxes = Array.from(document.querySelectorAll('.row-checkbox:checked'));
        if (checkedBoxes.length === 0) return;

        let totalSelected = checkedBoxes.length;
        let suratCount = 0;
        let slotCount = 0;

        const hiddenContainer = document.getElementById('bulkDeleteHiddenInputs');
        hiddenContainer.innerHTML = '';

        checkedBoxes.forEach(cb => {
            if (cb.dataset.hasSurat === '1') {
                suratCount++;
            } else {
                slotCount++;
            }
            const input = document.createElement('input');
            input.type = 'hidden';
            input.name = 'buku_ids[]';
            input.value = cb.value;
            hiddenContainer.appendChild(input);
        });

        let summaryText = `Anda akan menghapus <strong>${totalSelected} nomor</strong>`;
        if (suratCount > 0 && slotCount > 0) {
            summaryText += ` (terdiri dari <strong>${slotCount} slot kosong</strong> dan <strong>${suratCount} surat keluar</strong>).`;
        } else if (suratCount > 0) {
            summaryText += ` (semuanya berisi <strong>surat keluar</strong>).`;
        } else {
            summaryText += ` (semuanya berupa <strong>slot kosong</strong>).`;
        }

        document.getElementById('bulkDeleteSummary').innerHTML = summaryText;
        openModal('bulkDeleteModal');
    }
</script>

<?php include '../includes/footer.php'; ?>

