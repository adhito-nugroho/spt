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
    <div class="flex items-center justify-between">
        <h2 class="text-2xl font-bold text-slate-800">Dashboard</h2>
        <div class="text-sm text-slate-500">
            Terakhir diperbarui: <?php echo date('d M Y H:i:s'); ?> WIB
        </div>
    </div>

    <!-- Info Cards -->
    <div class="grid grid-cols-1 min-[480px]:grid-cols-2 md:grid-cols-4 gap-6">
        <!-- Card 1 -->
        <div class="bg-gradient-to-br from-blue-700 to-blue-800 rounded-xl shadow-lg p-6 text-white transform transition-all hover:scale-[1.02]">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-blue-100 font-medium mb-1">Total Pegawai</p>
                    <h3 class="text-4xl font-bold"><?php echo $total_pegawai; ?></h3>
                </div>
                <div class="p-3 bg-white/20 rounded-lg backdrop-blur-sm">
                    <i class='bx bxs-user text-2xl'></i>
                </div>
            </div>
            <div class="mt-4 flex items-center text-sm text-blue-100">
                <span class="bg-white/20 px-2 py-1 rounded text-xs mr-2">Aktif</span>
                Pegawai Terdaftar
            </div>
        </div>

        <!-- Card 2 -->
        <div class="bg-gradient-to-br from-blue-500 to-blue-600 rounded-xl shadow-lg p-6 text-white transform transition-all hover:scale-[1.02]">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-blue-100 font-medium mb-1">Total Surat Tugas</p>
                    <h3 class="text-4xl font-bold"><?php echo $total_surat; ?></h3>
                </div>
                <div class="p-3 bg-white/20 rounded-lg backdrop-blur-sm">
                    <i class='bx bxs-envelope text-2xl'></i>
                </div>
            </div>
            <div class="mt-4 flex items-center text-sm text-blue-100">
                <span class="bg-white/20 px-2 py-1 rounded text-xs mr-2">Total</span>
                Surat Dibuat
            </div>
        </div>

        <!-- Card 3: Surat Bulan Ini -->
        <div class="bg-gradient-to-br from-[#4f46e5] to-[#6366f1] rounded-xl shadow-lg p-6 text-white transform transition-all hover:scale-[1.02]">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-indigo-100 font-medium mb-1">Surat Bulan Ini</p>
                    <h3 class="text-4xl font-bold"><?php echo $surat_bulan_ini; ?></h3>
                </div>
                <div class="p-3 bg-white/20 rounded-lg backdrop-blur-sm">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor" class="w-6 h-6">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M6.75 3v2.25M17.25 3v2.25M3 18.75V7.5a2.25 2.25 0 012.25-2.25h13.5A2.25 2.25 0 0121 7.5v11.25m-18 0A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75m-18 0v-7.5A2.25 2.25 0 015.25 9h13.5A2.25 2.25 0 0121 11.25v7.5m-9-6h.008v.008H12v-.008zM12 15h.008v.008H12V15zm0 2.25h.008v.008H12v-.008zM9.75 15h.008v.008H9.75V15zm0 2.25h.008v.008H9.75v-.008zM7.5 15h.008v.008H7.5V15zm0 2.25h.008v.008H7.5v-.008zm6.75-4.5h.008v.008h-.008v-.008zm0 2.25h.008v.008h-.008V15zm0 2.25h.008v.008h-.008v-.008zm2.25-4.5h.008v.008H16.5v-.008zm0 2.25h.008v.008H16.5V15z" />
                    </svg>
                </div>
            </div>
            <div class="mt-4 flex items-center text-sm text-indigo-100">
                <span class="bg-white/20 px-2 py-1 rounded text-xs mr-2">Bulan Ini</span>
                Surat Diterbitkan
            </div>
        </div>

        <!-- Card 4: Surat Hari Ini -->
        <div class="bg-gradient-to-br from-[#0d9488] to-[#14b8a6] rounded-xl shadow-lg p-6 text-white transform transition-all hover:scale-[1.02]">
            <div class="flex justify-between items-start">
                <div>
                    <p class="text-teal-100 font-medium mb-1">Surat Hari Ini</p>
                    <h3 class="text-4xl font-bold"><?php echo $surat_hari_ini; ?></h3>
                </div>
                <div class="p-3 bg-white/20 rounded-lg backdrop-blur-sm">
                    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor" class="w-6 h-6">
                        <path stroke-linecap="round" stroke-linejoin="round" d="M10.125 2.25h3.75a2.25 2.25 0 012.25 2.25v15a2.25 2.25 0 01-2.25 2.25h-3.75a2.25 2.25 0 01-2.25-2.25v-15a2.25 2.25 0 012.25-2.25z" />
                        <path stroke-linecap="round" stroke-linejoin="round" d="M9 11.25l2.25 2.25L15 9" />
                    </svg>
                </div>
            </div>
            <div class="mt-4 flex items-center text-sm text-teal-100">
                <span class="bg-white/20 px-2 py-1 rounded text-xs mr-2">Hari Ini</span>
                Surat Diterbitkan
            </div>
        </div>
    </div>

    <!-- Recent Surat Tugas & Chart -->
    <div class="grid grid-cols-1 lg:grid-cols-12 gap-6">
        <!-- Table -->
        <div class="lg:col-span-7 bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
            <div class="p-6 border-b border-slate-100 flex justify-between items-center">
                <h5 class="font-bold text-slate-800">Surat Tugas Terbaru</h5>
                <a href="surat-tugas.php" class="text-sm text-blue-600 hover:text-blue-700 font-medium">Lihat Semua</a>
            </div>
            <div class="overflow-x-auto">
                <table class="w-full text-left border-collapse">
                    <thead>
                        <tr class="bg-slate-50 text-slate-600 text-xs uppercase tracking-wider">
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Nomor Surat</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Tanggal</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Pegawai</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100">Status</th>
                            <th class="px-6 py-4 font-semibold border-b border-slate-100 text-right">Aksi</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-100">
                        <?php foreach ($recent_surat as $surat): 
                            // Proses nama pegawai: tampilkan nama pertama + badge jika > 1
                            $nama_list = explode(', ', $surat['pegawai_names']);
                            $nama_pertama = $nama_list[0] ?? '-';
                            $sisa_pegawai = count($nama_list) - 1;
                            
                            // Tentukan badge status berdasarkan status buku nomor
                            $status = strtolower($surat['status_buku'] ?? 'terisi');
                            if ($status == 'terisi') {
                                $badge_class = 'bg-green-100 text-green-700';
                                $badge_label = 'Diterbitkan';
                            } elseif ($status == 'kosong') {
                                $badge_class = 'bg-slate-100 text-slate-600';
                                $badge_label = 'Draft';
                            } else {
                                $badge_class = 'bg-blue-100 text-blue-700';
                                $badge_label = 'Diterbitkan';
                            }
                        ?>
                            <tr class="hover:bg-slate-50 transition-colors">
                                <td class="px-6 py-4 text-sm font-medium text-slate-900">
                                    <?php echo htmlspecialchars($surat['nomor_surat']); ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600">
                                    <?php echo date('d/m/Y', strtotime($surat['tanggal_surat'])); ?>
                                </td>
                                <td class="px-6 py-4 text-sm text-slate-600">
                                    <div class="flex items-center gap-2">
                                        <div class="w-6 h-6 rounded-full bg-blue-100 flex items-center justify-center text-blue-700 text-xs font-bold flex-shrink-0">
                                            <?php echo strtoupper(substr($nama_pertama, 0, 1)); ?>
                                        </div>
                                        <span class="truncate max-w-[140px]" title="<?php echo htmlspecialchars($surat['pegawai_names']); ?>">
                                            <?php echo htmlspecialchars($nama_pertama); ?>
                                        </span>
                                        <?php if ($sisa_pegawai > 0): ?>
                                            <span class="inline-flex items-center px-1.5 py-0.5 rounded text-xs font-medium bg-blue-100 text-blue-700 flex-shrink-0 cursor-help" title="<?php echo $sisa_pegawai; ?> pegawai lainnya dalam surat ini">
                                                +<?php echo $sisa_pegawai; ?>
                                            </span>
                                        <?php endif; ?>
                                    </div>
                                </td>
                                <td class="px-6 py-4 text-sm">
                                    <span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium <?php echo $badge_class; ?>">
                                        <?php echo $badge_label; ?>
                                    </span>
                                </td>
                                <td class="px-6 py-4 text-sm text-right whitespace-nowrap">
                                    <a href="generate-surat.php?id=<?php echo $surat['id']; ?>" class="inline-flex items-center gap-1 px-2 py-1 text-xs font-medium border border-blue-200 rounded text-blue-600 hover:bg-blue-50 transition-colors" title="Lihat Detail">
                                        <i class='bx bx-show text-sm'></i>
                                        <span>Lihat</span>
                                    </a>
                                </td>
                            </tr>
                        <?php endforeach; ?>
                    </tbody>
                </table>
            </div>
            <div class="p-4 border-t border-slate-100 text-center">
                <a href="surat-tugas.php" class="text-sm text-blue-600 hover:text-blue-700 font-medium inline-flex items-center gap-1">
                    Lihat semua surat <i class='bx bx-right-arrow-alt'></i>
                </a>
            </div>
        </div>

        <!-- Chart -->
        <div class="lg:col-span-5 bg-white rounded-xl shadow-sm border border-slate-200">
            <div class="p-6 border-b border-slate-100">
                <h5 class="font-bold text-slate-800">Jumlah Surat Tugas per Bulan</h5>
                <p class="text-xs text-slate-500 mt-1">Data 6 bulan terakhir</p>
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
    // Inisialisasi Chart
    document.addEventListener('DOMContentLoaded', function () {
        const ctx = document.getElementById('suratChart').getContext('2d');
        
        // Gradient for chart
        const gradient = ctx.createLinearGradient(0, 0, 0, 300);
        gradient.addColorStop(0, 'rgba(37, 99, 235, 0.3)');
        gradient.addColorStop(1, 'rgba(37, 99, 235, 0.0)');

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
                    borderColor: '#2563eb',
                    backgroundColor: gradient,
                    borderWidth: 2.5,
                    pointBackgroundColor: '#2563eb',
                    pointBorderColor: '#ffffff',
                    pointBorderWidth: 2,
                    pointRadius: 5,
                    pointHoverRadius: 7,
                    pointHoverBackgroundColor: '#2563eb',
                    pointHoverBorderColor: '#ffffff',
                    pointHoverBorderWidth: 3,
                    fill: true,
                    tension: 0.4
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: {
                        display: false
                    },
                    tooltip: {
                        backgroundColor: '#1e293b',
                        padding: 12,
                        titleFont: {
                            size: 13,
                            weight: 'bold'
                        },
                        bodyFont: {
                            size: 13
                        },
                        cornerRadius: 8,
                        displayColors: false,
                        callbacks: {
                            title: function(tooltipItems) {
                                const idx = tooltipItems[0].dataIndex;
                                return chartBulan[idx] + ' ' + chartTahun[idx];
                            },
                            label: function(context) {
                                return 'Jumlah Surat: ' + context.parsed.y;
                            }
                        }
                    },
                    // Plugin untuk menampilkan nilai di atas titik data
                    datalabels: undefined
                },
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Jumlah Surat',
                            font: {
                                size: 12,
                                weight: '500'
                            },
                            color: '#64748b'
                        },
                        grid: {
                            borderDash: [2, 4],
                            color: '#e2e8f0',
                            drawBorder: false
                        },
                        ticks: {
                            font: {
                                size: 11
                            },
                            color: '#64748b',
                            stepSize: 1
                        }
                    },
                    x: {
                        grid: {
                            display: false
                        },
                        ticks: {
                            font: {
                                size: 11
                            },
                            color: '#64748b'
                        }
                    }
                }
            },
            plugins: [{
                // Custom plugin: tampilkan nilai angka di atas setiap titik
                afterDatasetsDraw: function(chart) {
                    const ctx = chart.ctx;
                    chart.data.datasets.forEach(function(dataset, i) {
                        const meta = chart.getDatasetMeta(i);
                        meta.data.forEach(function(element, index) {
                            const data = dataset.data[index];
                            ctx.fillStyle = '#1e40af';
                            ctx.font = 'bold 11px sans-serif';
                            ctx.textAlign = 'center';
                            ctx.textBaseline = 'bottom';
                            ctx.fillText(data, element.x, element.y - 10);
                        });
                    });
                }
            }]
        });
    });
</script>

<?php include '../includes/footer.php'; ?>
