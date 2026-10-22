<?php
if (session_status() === PHP_SESSION_NONE && !headers_sent()) {
    session_start();
}

$headerUser = isset($_SESSION['user']) && is_array($_SESSION['user']) ? $_SESSION['user'] : [];
$userName = trim((string)($headerUser['nama'] ?? $_SESSION['nama'] ?? ''));
$userJabatanRaw = trim((string)($headerUser['jabatan'] ?? $_SESSION['jabatan'] ?? ''));
$userGolongan = trim((string)($headerUser['golongan'] ?? $_SESSION['golongan'] ?? ''));
$userNip = trim((string)($headerUser['nip'] ?? $_SESSION['nip'] ?? ''));

if ($userName === '' && $userNip !== '' && isset($pdo) && $pdo instanceof PDO) {
    try {
        $stmtHeaderUser = $pdo->prepare("SELECT nama, pangkat, jabatan FROM pegawai WHERE nip = ? LIMIT 1");
        $stmtHeaderUser->execute([$userNip]);
        $pegawaiHeader = $stmtHeaderUser->fetch();

        if ($pegawaiHeader) {
            $userName = trim((string)($pegawaiHeader['nama'] ?? ''));
            $userJabatanRaw = trim((string)($pegawaiHeader['jabatan'] ?? ''));
            if ($userGolongan === '') {
                $userGolongan = trim((string)($pegawaiHeader['pangkat'] ?? ''));
            }
        }
    } catch (PDOException $e) {
        error_log('Gagal memuat data user header: ' . $e->getMessage());
    }
}

if ($userName === '') {
    $userName = "MUHIMATUN NI'MAH";
    $userJabatanRaw = 'Penyuluh Kehutanan Pemula';
    $userGolongan = 'V';
}

$userJabatan = trim((string)preg_replace('/\s+pada\s+.*$/i', '', $userJabatanRaw));
if ($userJabatan === '') {
    $userJabatan = 'Pegawai';
}

$userMeta = trim(
    ($userGolongan !== '' ? 'Golongan (' . $userGolongan . ')' : '')
    . ($userGolongan !== '' && $userJabatan !== '' ? ' · ' : '')
    . $userJabatan
);

$userUnit = 'CDK Wilayah Bojonegoro';
$userInitial = function_exists('mb_substr')
    ? mb_strtoupper(mb_substr($userName, 0, 1, 'UTF-8'), 'UTF-8')
    : strtoupper(substr($userName, 0, 1));
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Pengelolaan Surat Tugas</title>
    <link rel="icon" type="image/svg+xml" href="../assets/img/logo-jawa-timur.svg">
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <?php include __DIR__ . '/tailwind-theme.php'; ?>
    
    <!-- Boxicons -->
    <link href="https://unpkg.com/boxicons@2.1.4/css/boxicons.min.css" rel="stylesheet">
    
    <!-- Select2 -->
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    
    <!-- jQuery -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    
    <!-- Select2 JS -->
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    
    <!-- SweetAlert2 -->
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
    
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>

    <!-- SIPENSURAT Design System -->
    <link rel="stylesheet" href="../assets/css/app.css">
</head>
<body class="bg-gray-50 font-sans text-gray-900 antialiased">

    <div class="flex h-screen overflow-hidden">
        <!-- Sidebar -->
        <aside class="w-64 bg-white border-r border-gray-200 hidden md:flex flex-col z-10">
            <div class="h-16 flex items-center px-6 border-b border-gray-100">
                <a href="../" class="flex items-center gap-2 text-indigo-600 font-bold text-xl tracking-tight">
                    <img src="../assets/img/logo-jawa-timur.svg" alt="Lambang Provinsi Jawa Timur" class="h-9 w-auto shrink-0 object-contain" decoding="async">
                    <span>SIPENSURAT</span>
                </a>
            </div>
            
            <nav class="flex-1 overflow-y-auto py-5">
                <ul class="space-y-0.5 px-3 pl-4">
                    <li>
                        <a href="../modules/dashboard.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'dashboard.php' ? 'active' : ''; ?>">
                            <i class='bx bxs-dashboard' style="font-size:1.1rem;flex-shrink:0;"></i>
                            <span>Dashboard</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/pegawai.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'pegawai.php' ? 'active' : ''; ?>">
                            <i class='bx bxs-user-detail' style="font-size:1.1rem;flex-shrink:0;"></i>
                            <span>Data Pegawai</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/surat-tugas.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'surat-tugas.php' ? 'active' : ''; ?>">
                            <i class='bx bxs-envelope' style="font-size:1.1rem;flex-shrink:0;"></i>
                            <span>Surat Tugas</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/buku-nomor.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'buku-nomor.php' ? 'active' : ''; ?>">
                            <i class='bx bxs-book-content' style="font-size:1.1rem;flex-shrink:0;"></i>
                            <span>Buku Nomor</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/surat-keluar.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'surat-keluar.php' ? 'active' : ''; ?>">
                            <i class='bx bx-send' style="font-size:1.1rem;flex-shrink:0;"></i>
                            <span>Surat Keluar</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/template.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'template.php' ? 'active' : ''; ?>">
                            <i class='bx bxs-file-blank' style="font-size:1.1rem;flex-shrink:0;"></i>
                            <span>Template Word</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/penandatangan.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'penandatangan.php' ? 'active' : ''; ?>">
                            <i class='bx bxs-pen' style="font-size:1.1rem;flex-shrink:0;"></i>
                            <span>Penanda Tangan</span>
                        </a>
                    </li>
                </ul>
                
                <div class="mt-6 px-4">
                    <p class="text-xs font-medium text-gray-400 uppercase tracking-wider mb-2 px-3">Lainnya</p>
                    <ul class="space-y-0.5">
                        <li>
                            <a href="../modules/laporan.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'laporan.php' ? 'active' : ''; ?>">
                                <i class='bx bxs-report' style="font-size:1.1rem;flex-shrink:0;"></i>
                                <span>Laporan</span>
                            </a>
                        </li>
                        <li>
                            <a href="../modules/laporan-pegawai.php" class="sidebar-link <?php echo basename($_SERVER['PHP_SELF']) == 'laporan-pegawai.php' ? 'active' : ''; ?>">
                                <i class='bx bxs-user-check' style="font-size:1.1rem;flex-shrink:0;"></i>
                                <span>Laporan Per Pegawai</span>
                            </a>
                        </li>
                        <li>
                            <a href="#" class="sidebar-link">
                                <i class='bx bxs-help-circle' style="font-size:1.1rem;flex-shrink:0;"></i>
                                <span>Bantuan</span>
                            </a>
                        </li>
                    </ul>
                </div>
            </nav>
            
            <div class="p-4 border-t border-gray-100">
                <a href="#" class="sidebar-link text-red-600 hover:text-red-700" style="color:#dc2626;">
                    <i class='bx bx-log-out' style="font-size:1.1rem;flex-shrink:0;"></i>
                    <span>Keluar</span>
                </a>
            </div>
        </aside>

        <!-- Main Content -->
        <div class="flex-1 flex flex-col min-w-0 overflow-hidden">
            <!-- Top Header -->
            <header class="h-16 bg-white border-b border-gray-200 flex items-center justify-between px-4 sm:px-6 lg:px-8">
                <button class="md:hidden p-2 rounded-lg text-gray-500 hover:bg-gray-100">
                    <i class='bx bx-menu text-2xl'></i>
                </button>
                
                <div class="flex items-center gap-3 ml-auto">
                    <!-- Notification Bell Button -->
                    <button class="relative p-1.5 text-gray-400 hover:text-primary hover:bg-gray-50 rounded-lg transition-colors" aria-label="Notifikasi">
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor" class="w-5 h-5">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M14.857 17.082a23.848 23.848 0 005.454-1.31A8.967 8.967 0 0118 9.75v-.7V9A6 6 0 006 9v.75a8.967 8.967 0 01-2.312 6.022c1.733.64 3.56 1.085 5.455 1.31m5.714 0a24.255 24.255 0 01-5.714 0m5.714 0a3 3 0 11-5.714 0" />
                        </svg>
                        <span class="absolute top-0.5 right-0.5 w-1.5 h-1.5 bg-red-500 rounded-full"></span>
                    </button>

                    <!-- Separator -->
                    <div class="h-8 w-px bg-gray-200"></div>

                    <div class="flex min-w-0 items-center gap-2.5">
                        <div class="hidden min-w-0 max-w-[15rem] text-right sm:block">
                            <p class="truncate text-sm font-semibold leading-tight text-gray-800"><?= htmlspecialchars($userName, ENT_QUOTES, 'UTF-8') ?></p>
                            <p class="truncate text-[11px] leading-tight text-gray-500" title="<?= htmlspecialchars($userMeta, ENT_QUOTES, 'UTF-8') ?>"><?= htmlspecialchars($userMeta, ENT_QUOTES, 'UTF-8') ?></p>
                            <p class="truncate text-[10px] leading-tight text-gray-400"><?= htmlspecialchars($userUnit, ENT_QUOTES, 'UTF-8') ?></p>
                        </div>
                        <div class="flex h-9 w-9 shrink-0 items-center justify-center rounded-full bg-primary-light text-sm font-bold text-primary ring-1 ring-primary-border">
                            <?= htmlspecialchars($userInitial, ENT_QUOTES, 'UTF-8') ?>
                        </div>
                    </div>
                </div>
            </header>

            <!-- Page Content -->
            <main class="flex-1 overflow-y-auto p-6 lg:p-8">
                <div class="max-w-7xl mx-auto">
