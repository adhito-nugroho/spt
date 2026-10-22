<?php
// Set timezone ke Indonesia (WIB)
date_default_timezone_set('Asia/Jakarta');

require_once '../config/database.php';
include '../includes/header.php';

// Mengambil total data
$stmt_pegawai = $pdo->query("SELECT COUNT(*) as total FROM pegawai");
$total_pegawai = $stmt_pegawai->fetch()['total'];

$stmt_surat = $pdo->query("SELECT COUNT(*) as total FROM surat_tugas");
$total_surat = $stmt_surat->fetch()['total'];

// Mengambil total surat bulan ini dan hari ini
$stmt_bulan = $pdo->query("SELECT COUNT(*) as total FROM surat_tugas WHERE MONTH(tanggal_surat) = MONTH(NOW()) AND YEAR(tanggal_surat) = YEAR(NOW())");
$surat_bulan_ini = $stmt_bulan->fetch()['total'];

$stmt_hari = $pdo->query("SELECT COUNT(*) as total FROM surat_tugas WHERE DATE(tanggal_surat) = CURDATE()");
$surat_hari_ini = $stmt_hari->fetch()['total'];

// Mengambil surat tugas terbaru
$stmt_recent = $pdo->query("SELECT 
    st.id,
    st.nomor_surat,
    st.tanggal_surat,
    COALESCE(bn.status, 'terisi') as status_buku,
    GROUP_CONCAT(p.nama SEPARATOR ', ') as pegawai_names,
    COUNT(p.nip) as jumlah_pegawai
FROM surat_tugas st 
LEFT JOIN pegawai_tugas pt ON st.id = pt.id_surat_tugas 
LEFT JOIN pegawai p ON pt.nip = p.nip 
LEFT JOIN buku_nomor_surat_tugas bn ON bn.id_surat_tugas = st.id
GROUP BY st.id, st.nomor_surat, st.tanggal_surat, bn.status
ORDER BY st.tanggal_surat DESC 
LIMIT 5");
$recent_surat = $stmt_recent->fetchAll();

// Mengambil data untuk grafik
$stmt_chart = $pdo->query("SELECT 
    DATE_FORMAT(tanggal_surat, '%M') as bulan,
    MONTH(tanggal_surat) as bulan_angka,
    YEAR(tanggal_surat) as tahun,
    COUNT(*) as total 
FROM surat_tugas 
WHERE tanggal_surat >= DATE_SUB(NOW(), INTERVAL 6 MONTH) 
GROUP BY YEAR(tanggal_surat), MONTH(tanggal_surat), DATE_FORMAT(tanggal_surat, '%M')
ORDER BY tahun, bulan_angka");
$chart_data = $stmt_chart->fetchAll();

// Siapkan label grafik dengan format "Jan 2026"
$chart_labels = [];
$chart_tahun = [];
foreach ($chart_data as $cd) {
    $bulan_short = substr($cd['bulan'], 0, 3);
    $chart_labels[] = $bulan_short . ' ' . $cd['tahun'];
    $chart_tahun[] = $cd['tahun'];
}
?>

<div class="space-y-6">
    <!-- Page Header -->
    <div class="flex items-center justify-between">
        <div>
            <h1 class="page-header-title">Dashboard</h1>
            <p class="page-header-sub">Ringkasan data aplikasi SIPENSURAT</p>
        </div>
        <div class="text-xs text-gray-400">
            Diperbarui: <?php echo date('d M Y H:i'); ?> WIB
        </div>
    </div>

    <!-- Stat Cards -->
    <div class="grid grid-cols-1 min-[480px]:grid-cols-2 md:grid-cols-4 gap-6">
        <!-- Card 1: Total Pegawai -->
        <div class="stat-card blue">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Total Pegawai</p>
                    <h3 class="text-3xl font-bold text-gray-900"><?php echo $total_pegawai; ?></h3>
                </div>
                <div class="stat-icon blue">
                    <i class='bx bxs-user'></i>
                </div>
            </div>
            <div class="mt-4 flex items-center gap-2">
                <span class="badge badge-gray">Aktif</span>
                <span class="text-xs text-gray-500">Pegawai Terdaftar</span>
            </div>
        </div>

        <!-- Card 2: Total Surat Tugas -->
        <div class="stat-card indigo">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Total Surat Tugas</p>
                    <h3 class="text-3xl font-bold text-gray-900"><?php echo $total_surat; ?></h3>
                </div>
                <div class="stat-icon indigo">
                    <i class='bx bxs-envelope'></i>
                </div>
            </div>
            <div class="mt-4 flex items-center gap-2">
                <span class="badge badge-indigo">Total</span>
                <span class="text-xs text-gray-500">Surat Dibuat</span>
            </div>
        </div>

        <!-- Card 3: Surat Bulan Ini -->
        <div class="stat-card violet">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Surat Bulan Ini</p>
                    <h3 class="text-3xl font-bold text-gray-900"><?php echo $surat_bulan_ini; ?></h3>
                </div>
                <div class="stat-icon violet">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor" class="w-5 h-5">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M6.75 3v2.25M17.25 3v2.25M3 18.75V7.5a2.25 2.25 0 012.25-2.25h13.5A2.25 2.25 0 0121 7.5v11.25m-18 0A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75m-18 0v-7.5A2.25 2.25 0 015.25 9h13.5A2.25 2.25 0 0121 11.25v7.5m-9-6h.008v.008H12v-.008zM12 15h.008v.008H12V15zm0 2.25h.008v.008H12v-.008zM9.75 15h.008v.008H9.75V15zm0 2.25h.008v.008H9.75v-.008zM7.5 15h.008v.008H7.5V15zm0 2.25h.008v.008H7.5v-.008zm6.75-4.5h.008v.008h-.008v-.008zm0 2.25h.008v.008h-.008V15zm0 2.25h.008v.008h-.008v-.008zm2.25-4.5h.008v.008H16.5v-.008zm0 2.25h.008v.008H16.5V15z" />
                    </svg>
                </div>
            </div>
            <div class="mt-4 flex items-center gap-2">
                <span class="badge badge-violet">Bulan Ini</span>
                <span class="text-xs text-gray-500">Surat Diterbitkan</span>
            </div>
        </div>

        <!-- Card 4: Surat Hari Ini -->
        <div class="stat-card teal">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-xs font-semibold text-gray-500 uppercase tracking-wide mb-2">Surat Hari Ini</p>
                    <h3 class="text-3xl font-bold text-gray-900"><?php echo $surat_hari_ini; ?></h3>
                </div>
                <div class="stat-icon teal">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor" class="w-5 h-5">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M9 12.75L11.25 15 15 9.75M21 12a9 9 0 11-18 0 9 9 0 0118 0z" />
                    </svg>
                </div>
            </div>
            <div class="mt-4 flex items-center gap-2">
                <span class="badge badge-emerald">Hari Ini</span>
                <span class="text-xs text-gray-500">Surat Diterbitkan</span>
            </div>
        </div>
    </div>

    <!-- Recent Surat Tugas & Chart -->
    <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <!-- Table -->
        <div class="lg:col-span-7 card overflow-hidden">
            <div class="px-6 py-4 border-b border-gray-100 flex justify-between items-center">
                <div>
                    <h2 class="text-sm font-semibold text-gray-900">Surat Tugas Terbaru</h2>
                    <p class="text-xs text-gray-400 mt-0.5">5 surat terakhir diterbitkan</p>
                </div>
                <a href="surat-tugas.php" class="text-xs font-medium text-indigo-600 hover:text-indigo-700">Lihat Semua →</a>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full text-left table-premium">
                    <thead>
                        <tr>
                            <th>Nomor Surat</th>
                            <th>Tanggal</th>
                            <th>Pegawai</th>
                            <th>Status</th>
                            <th class="text-right">Aksi</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php foreach ($recent_surat as $surat):
                            $nama_list = explode(', ', $surat['pegawai_names']);
                            $nama_pertama = $nama_list[0] ?? '-';
                            $sisa_pegawai = count($nama_list) - 1;
                            
                            $status = strtolower($surat['status_buku'] ?? 'terisi');
                            if ($status == 'terisi') {
                                $badge_class = 'badge-green';
                                $badge_label = 'Diterbitkan';
                            } elseif ($status == 'kosong') {
                                $badge_class = 'badge-gray';
                                $badge_label = 'Draft';
                            } else {
                                $badge_class = 'badge-blue';
                                $badge_label = 'Diterbitkan';
                            }
                        ?>
                            <tr>
                                <td class="font-medium text-gray-900">
                                    <?php echo htmlspecialchars($surat['nomor_surat']); ?>
                                </td>
                                <td class="text-gray-500">
                                    <?php echo date('d/m/Y', strtotime($surat['tanggal_surat'])); ?>
                                </td>
                                <td>
                                    <div class="flex items-center gap-2">
                                        <div class="avatar avatar-0">
                                            <?php echo strtoupper(substr($nama_pertama, 0, 1)); ?>
                                        </div>
                                        <span class="truncate max-w-[130px] text-gray-700" title="<?php echo htmlspecialchars($surat['pegawai_names']); ?>">
                                            <?php echo htmlspecialchars($nama_pertama); ?>
                                        </span>
                                        <?php if ($sisa_pegawai > 0): ?>
                                            <span class="badge badge-blue cursor-help" title="<?php echo $sisa_pegawai; ?> pegawai lainnya">
                                                +<?php echo $sisa_pegawai; ?>
                                            </span>
                                        <?php endif; ?>
                                    </div>
                                </td>
                                <td>
                                    <span class="badge <?php echo $badge_class; ?>">
                                        <?php echo $badge_label; ?>
                                    </span>
                                </td>
                                <td class="text-right">
                                    <a href="generate-surat.php?id=<?php echo $surat['id']; ?>" class="btn-action primary">
                                        <i class='bx bx-show'></i> Lihat
                                    </a>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
            <div class="px-6 py-3 border-t border-gray-50 text-center">
                <a href="surat-tugas.php" class="text-xs font-medium text-indigo-600 hover:text-indigo-700 inline-flex items-center gap-1">
                    Lihat semua surat <i class='bx bx-right-arrow-alt'></i>
                </a>
            </div>
        </div>

        <!-- Chart -->
        <div class="lg:col-span-5 card">
            <div class="px-6 py-4 border-b border-gray-100">
                <h2 class="text-sm font-semibold text-gray-900">Surat Tugas per Bulan</h2>
                <p class="text-xs text-gray-400 mt-0.5">Data 6 bulan terakhir</p>
            </div>
            <div class="p-6">
                <div style="height: 280px;">
                    <canvas id="suratChart"></canvas>
                </div>
            </div>
        </div>
    </div>
</div>

<script>
    document.addEventListener('DOMContentLoaded', function () {
        const ctx = document.getElementById('suratChart').getContext('2d');
        
        // Gradient fill
        const gradient = ctx.createLinearGradient(0, 0, 0, 300);
        gradient.addColorStop(0, 'rgba(99, 102, 241, 0.18)');
        gradient.addColorStop(1, 'rgba(99, 102, 241, 0.0)');

        const chartData = <?php echo json_encode(array_column($chart_data, 'total')); ?>;
        const chartLabels = <?php echo json_encode($chart_labels); ?>;
        const chartTahun = <?php echo json_encode($chart_tahun); ?>;
        const chartBulan = <?php echo json_encode(array_column($chart_data, 'bulan')); ?>;

        new Chart(ctx, {
            type: 'line',
            data: {
                labels: chartLabels,
                datasets: [{
                    label: 'Jumlah Surat',
                    data: chartData,
                    borderColor: '#6366f1',
                    backgroundColor: gradient,
                    borderWidth: 2,
                    pointBackgroundColor: '#6366f1',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 2,
                    pointRadius: 3,
                    pointHoverRadius: 5,
                    pointHoverBackgroundColor: '#4f46e5',
                    pointHoverBorderColor: '#ffffff',
                    pointHoverBorderWidth: 2,
                    fill: true,
                    tension: 0.4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: '#111827',
                        padding: 10,
                        titleFont: { size: 12, weight: '600', family: 'Inter' },
                        bodyFont: { size: 12, family: 'Inter' },
                        cornerRadius: 6,
                        displayColors: false,
                        callbacks: {
                            title: function(items) {
                                const idx = items[0].dataIndex;
                                return chartBulan[idx] + ' ' + chartTahun[idx];
                            },
                            label: function(ctx) {
                                return 'Jumlah Surat: ' + ctx.parsed.y;
                            }
                        }
                    }
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        grid: {
                            color: 'rgba(0,0,0,0.04)',
                            drawBorder: false
                        },
                        ticks: {
                            color: '#9ca3af',
                            font: { size: 11 },
                            stepSize: 5,
                            precision: 0
                        }
                    },
                    x: {
                        grid: { display: false },
                        ticks: {
                            color: '#9ca3af',
                            font: { size: 11 }
                        }
                    }
                }
            }
        });
    });
</script>

<?php include '../includes/footer.php'; ?>
