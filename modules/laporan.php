<?php
require_once '../vendor/autoload.php';
use Dompdf\Dompdf;
use Dompdf\Options;

$pdo = require_once '../config/database.php';

// Inisialisasi variabel
$start_date = isset($_GET['start_date']) ? $_GET['start_date'] : date('Y-m-01');
$end_date = isset($_GET['end_date']) ? $_GET['end_date'] : date('Y-m-d');
$export = isset($_GET['export']) ? strtolower(trim($_GET['export'])) : '';
$laporan = [];

// Pagination (khusus tampilan, bukan untuk export)
$limit = 10;
$page = isset($_GET['page']) ? (int)$_GET['page'] : 1;
if ($page < 1) $page = 1;
$start = ($page - 1) * $limit;
$total_records = 0;
$total_pages = 0;

// Ambil data jika ada filter
if ($start_date && $end_date) {
    // Untuk export, tetap ambil semua data sesuai filter periode
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
        // Untuk tampilan, hitung total & ambil data per halaman
        $query_count = "SELECT COUNT(DISTINCT st.id) AS total
            FROM surat_tugas st
            LEFT JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas
            LEFT JOIN pegawai p ON pt.nip = p.nip
            WHERE st.tanggal_surat BETWEEN :start_date AND :end_date";
        $stmt_count = $pdo->prepare($query_count);
        $stmt_count->execute([
            ':start_date' => $start_date,
            ':end_date' => $end_date
        ]);
        $total_records = (int)($stmt_count->fetch()['total'] ?? 0);
        $total_pages = (int)ceil($total_records / $limit);

        $query = "SELECT 
            st.*,
            GROUP_CONCAT(DISTINCT p.nama SEPARATOR ', ') as pegawai_names
            FROM surat_tugas st 
            LEFT JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas 
            LEFT JOIN pegawai p ON pt.nip = p.nip 
            WHERE st.tanggal_surat BETWEEN :start_date AND :end_date
            GROUP BY st.id 
            ORDER BY st.tanggal_surat ASC, st.id ASC
            LIMIT :start, :limit";
        
        $stmt = $pdo->prepare($query);
        $stmt->bindValue(':start', $start, PDO::PARAM_INT);
        $stmt->bindValue(':limit', $limit, PDO::PARAM_INT);
        $stmt->bindValue(':start_date', $start_date);
        $stmt->bindValue(':end_date', $end_date);
        $stmt->execute();
        $laporan = $stmt->fetchAll();
    }
}

// Export Excel berdasarkan data/filter yang sedang ditampilkan
if ($export === 'excel') {
    $filename = 'laporan_surat_tugas_' . $start_date . '_sd_' . $end_date . '.xls';
    $filename = preg_replace('/[^a-zA-Z0-9_\-\.]/', '_', $filename);

    header('Content-Type: application/vnd.ms-excel; charset=utf-8');
    header('Content-Disposition: attachment; filename="' . $filename . '"');
    header('Pragma: no-cache');
    header('Expires: 0');

    echo '<table border="1">';
    echo '<tr><th colspan="6">Laporan Rekapitulasi Surat Tugas</th></tr>';
    echo '<tr><th colspan="6">Periode: ' .
        htmlspecialchars(date('d/m/Y', strtotime($start_date)), ENT_QUOTES, 'UTF-8') .
        ' s/d ' .
        htmlspecialchars(date('d/m/Y', strtotime($end_date)), ENT_QUOTES, 'UTF-8') .
        '</th></tr>';
    echo '<tr>';
    echo '<th>No</th>';
    echo '<th>Nomor Surat</th>';
    echo '<th>Tanggal</th>';
    echo '<th>Tujuan / Maksud</th>';
    echo '<th>Pegawai</th>';
    echo '<th>Pelaksanaan</th>';
    echo '</tr>';

    if (count($laporan) > 0) {
        $no = 1;
        foreach ($laporan as $row) {
            $pelaksanaan = date('d/m/Y', strtotime($row['tanggal_mulai']));
            if ($row['tanggal_mulai'] != $row['tanggal_selesai']) {
                $pelaksanaan .= ' - ' . date('d/m/Y', strtotime($row['tanggal_selesai']));
            }

            echo '<tr>';
            echo '<td>' . $no++ . '</td>';
            echo '<td>' . htmlspecialchars($row['nomor_surat'], ENT_QUOTES, 'UTF-8') . '</td>';
            echo '<td>' . htmlspecialchars(date('d/m/Y', strtotime($row['tanggal_surat'])), ENT_QUOTES, 'UTF-8') . '</td>';
            echo '<td>' . htmlspecialchars($row['untuk'], ENT_QUOTES, 'UTF-8') . '</td>';
            echo '<td>' . htmlspecialchars($row['pegawai_names'], ENT_QUOTES, 'UTF-8') . '</td>';
            echo '<td>' . htmlspecialchars($pelaksanaan, ENT_QUOTES, 'UTF-8') . '</td>';
            echo '</tr>';
        }
    } else {
        echo '<tr><td colspan="6">Tidak ada data surat tugas pada periode ini.</td></tr>';
    }

    echo '</table>';
    exit;
}

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

include '../includes/header.php';
?>

<div class="space-y-6">
    <div class="flex flex-col md:flex-row justify-between items-center gap-4 print:hidden">
        <h2 class="text-2xl font-bold text-slate-800">Laporan Surat Tugas</h2>
        <div class="flex gap-2">
            <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&export=excel" class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
                <i class='bx bx-spreadsheet'></i> Export Excel
            </a>
            <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&export=pdf" target="_blank" class="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg flex items-center gap-2 transition-colors shadow-sm">
                <i class='bx bx-printer'></i> Cetak Laporan
            </a>
        </div>
    </div>

    <!-- Filter Section (Hidden on Print) -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4 print:hidden">
        <form action="" method="GET" class="flex flex-col md:flex-row gap-4 items-end">
            <div class="w-full md:w-auto">
                <label class="block text-sm font-medium text-slate-700 mb-1">Dari Tanggal</label>
                <input type="date" name="start_date" value="<?php echo $start_date; ?>" 
                    class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
            </div>
            <div class="w-full md:w-auto">
                <label class="block text-sm font-medium text-slate-700 mb-1">Sampai Tanggal</label>
                <input type="date" name="end_date" value="<?php echo $end_date; ?>" 
                    class="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none">
            </div>
            <button type="submit" class="bg-slate-800 hover:bg-slate-900 text-white px-6 py-2 rounded-lg transition-colors w-full md:w-auto">
                Tampilkan
            </button>
        </form>
    </div>

    <!-- Report Header (Visible only on Print) -->
    <div class="hidden print:block text-center mb-8">
        <h1 class="text-2xl font-bold uppercase mb-2">Laporan Rekapitulasi Surat Tugas</h1>
        <p class="text-slate-600">Periode: <?php echo date('d/m/Y', strtotime($start_date)); ?> s/d <?php echo date('d/m/Y', strtotime($end_date)); ?></p>
    </div>

    <!-- Table Result -->
    <div class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden print:shadow-none print:border-0">
        <div class="overflow-x-auto">
            <table class="w-full text-left border-collapse">
                <thead>
                    <tr class="bg-slate-50 text-slate-600 text-sm uppercase tracking-wider print:bg-white print:border-b-2 print:border-slate-800">
                        <th class="px-6 py-4 font-semibold border-b border-slate-100 print:px-2 print:py-2">No</th>
                        <th class="px-6 py-4 font-semibold border-b border-slate-100 print:px-2 print:py-2">Nomor Surat</th>
                        <th class="px-6 py-4 font-semibold border-b border-slate-100 print:px-2 print:py-2">Tanggal</th>
                        <th class="px-6 py-4 font-semibold border-b border-slate-100 print:px-2 print:py-2">Tujuan / Maksud</th>
                        <th class="px-6 py-4 font-semibold border-b border-slate-100 print:px-2 print:py-2">Pegawai</th>
                        <th class="px-6 py-4 font-semibold border-b border-slate-100 print:px-2 print:py-2">Pelaksanaan</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-slate-100 print:divide-slate-200">
                    <?php if (count($laporan) > 0): ?>
                        <?php $no = $start + 1; foreach ($laporan as $row): ?>
                            <tr class="hover:bg-slate-50 transition-colors print:hover:bg-white">
                                <td class="px-6 py-4 text-sm text-slate-600 print:px-2 print:py-2 align-top"><?php echo $no++; ?></td>
                                <td class="px-6 py-4 text-sm font-medium text-slate-900 print:px-2 print:py-2 align-top whitespace-nowrap">
                                    <?php echo $row['nomor_surat']; ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600 print:px-2 print:py-2 align-top whitespace-nowrap">
                                    <?php echo date('d/m/Y', strtotime($row['tanggal_surat'])); ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600 print:px-2 print:py-2 align-top">
                                    <?php echo $row['untuk']; ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600 print:px-2 print:py-2 align-top">
                                    <?php echo $row['pegawai_names']; ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600 print:px-2 print:py-2 align-top whitespace-nowrap">
                                    <?php 
                                    echo date('d/m/Y', strtotime($row['tanggal_mulai']));
                                    if ($row['tanggal_mulai'] != $row['tanggal_selesai']) {
                                        echo ' - ' . date('d/m/Y', strtotime($row['tanggal_selesai']));
                                    }
                                    ?>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    <?php else: ?>
                        <tr>
                            <td colspan="6" class="px-6 py-8 text-center text-slate-500">
                                Tidak ada data surat tugas pada periode ini.
                            </td>
                        </tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>

        <!-- Pagination -->
        <?php if ($total_pages > 1): ?>
            <div class="px-6 py-4 border-t border-slate-100 flex justify-center print:hidden">
                <nav class="flex items-center gap-1">
                    <?php if ($page > 1): ?>
                        <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&page=<?php echo ($page - 1); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Previous</a>
                    <?php endif; ?>

                    <?php
                        $maxPageLinks = 7; // batasi tampilan nomor halaman
                        $half = (int)floor($maxPageLinks / 2);
                        $from = max(1, $page - $half);
                        $to = min($total_pages, $from + $maxPageLinks - 1);
                        $from = max(1, $to - $maxPageLinks + 1);
                    ?>
                    <?php for ($i = $from; $i <= $to; $i++): ?>
                        <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&page=<?php echo $i; ?>"
                           class="px-3 py-1 rounded border <?php echo($page == $i) ? 'bg-indigo-600 text-white border-indigo-600' : 'border-slate-200 text-slate-600 hover:bg-slate-50'; ?> text-sm">
                            <?php echo $i; ?>
                        </a>
                    <?php endfor; ?>

                    <?php if ($page < $total_pages): ?>
                        <a href="?start_date=<?php echo urlencode($start_date); ?>&end_date=<?php echo urlencode($end_date); ?>&page=<?php echo ($page + 1); ?>"
                           class="px-3 py-1 rounded border border-slate-200 text-slate-600 hover:bg-slate-50 text-sm">Next</a>
                    <?php endif; ?>
                </nav>
            </div>
        <?php endif; ?>
    </div>
    
    <!-- Print Signature Section -->
    <div class="hidden print:block mt-12">
        <div class="flex justify-end">
            <div class="text-center w-64">
                <p class="mb-20">Mengetahui,</p>
                <p class="font-bold underline">Nama Pejabat</p>
                <p>NIP. ...........................</p>
            </div>
        </div>
    </div>
</div>

<style>
    @media print {
        @page {
            size: landscape;
            margin: 1cm;
        }
        body {
            background: white;
        }
        /* Hide sidebar and other non-printable elements */
        aside, header, .print\:hidden {
            display: none !important;
        }
        /* Ensure main content takes full width */
        .flex-1 {
            margin: 0 !important;
            padding: 0 !important;
            width: 100% !important;
        }
        /* Reset shadows and borders for clean print */
        .shadow-sm, .shadow-lg {
            box-shadow: none !important;
        }
        .border {
            border: none !important;
        }
        /* Show hidden print elements */
        .print\:block {
            display: block !important;
        }
        /* Table styling for print */
        table {
            width: 100% !important;
            border: 1px solid #000 !important;
        }
        th, td {
            border: 1px solid #000 !important;
            padding: 8px !important;
        }
        thead tr {
            background-color: #f0f0f0 !important;
            -webkit-print-color-adjust: exact;
        }
    }
</style>

<?php include '../includes/footer.php'; ?>
