<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Pengelolaan Surat Tugas</title>
    
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    
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

    <script>
        tailwind.config = {
            theme: {
                extend: {
                    colors: {
                        primary: '#2c3e50',
                        secondary: '#34495e',
                    }
                }
            }
        }
    </script>
</head>
<body class="bg-slate-50 font-sans text-slate-800 antialiased">

    <div class="flex h-screen overflow-hidden">
        <!-- Sidebar -->
        <aside class="w-64 bg-white border-r border-slate-200 hidden md:flex flex-col z-10">
            <div class="h-16 flex items-center px-6 border-b border-slate-200">
                <a href="../" class="flex items-center gap-2 text-indigo-600 font-bold text-xl">
                    <i class='bx bxs-file-doc text-2xl'></i>
                    <span>SIPENSURAT</span>
                </a>
            </div>
            
            <nav class="flex-1 overflow-y-auto py-4">
                <ul class="space-y-1 px-3">
                    <li>
                        <a href="../modules/dashboard.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'dashboard.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                            <i class='bx bxs-dashboard text-xl'></i>
                            <span class="font-medium">Dashboard</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/pegawai.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'pegawai.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                            <i class='bx bxs-user-detail text-xl'></i>
                            <span class="font-medium">Data Pegawai</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/surat-tugas.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'surat-tugas.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                            <i class='bx bxs-envelope text-xl'></i>
                            <span class="font-medium">Surat Tugas</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/buku-nomor.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'buku-nomor.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                            <i class='bx bxs-book-content text-xl'></i>
                            <span class="font-medium">Buku Nomor</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/surat-keluar.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'surat-keluar.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                            <i class='bx bx-send text-xl'></i>
                            <span class="font-medium">Surat Keluar</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/template.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'template.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                            <i class='bx bxs-file-blank text-xl'></i>
                            <span class="font-medium">Template Word</span>
                        </a>
                    </li>
                    <li>
                        <a href="../modules/penandatangan.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'penandatangan.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                            <i class='bx bxs-pen text-xl'></i>
                            <span class="font-medium">Penanda Tangan</span>
                        </a>
                    </li>
                </ul>
                
                <div class="mt-8 px-6">
                    <p class="text-xs font-semibold text-slate-400 tracking-wider mb-2">Lainnya</p>
                    <ul class="space-y-1">
                        <li>
                            <a href="../modules/laporan.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'laporan.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                                <i class='bx bxs-report text-xl'></i>
                                <span class="font-medium">Laporan</span>
                            </a>
                        </li>
                        <li>
                            <a href="../modules/laporan-pegawai.php" class="flex items-center gap-3 px-3 py-2 rounded-lg transition-colors <?php echo basename($_SERVER['PHP_SELF']) == 'laporan-pegawai.php' ? 'bg-indigo-50 text-indigo-600' : 'text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150'; ?>">
                                <i class='bx bxs-user-check text-xl'></i>
                                <span class="font-medium">Laporan Per Pegawai</span>
                            </a>
                        </li>
                        <li>
                            <a href="#" class="flex items-center gap-3 px-3 py-2 rounded-lg text-slate-600 hover:bg-gray-100 hover:text-indigo-600 transition-colors duration-150">
                                <i class='bx bxs-help-circle text-xl'></i>
                                <span class="font-medium">Bantuan</span>
                            </a>
                        </li>
                    </ul>
                </div>
            </nav>
            
            <div class="p-4 border-t border-slate-200">
                <a href="#" class="flex items-center gap-3 px-3 py-2 rounded-lg text-red-600 hover:bg-red-50 transition-colors">
                    <i class='bx bx-log-out text-xl'></i>
                    <span class="font-medium">Keluar</span>
                </a>
            </div>
        </aside>

        <!-- Main Content -->
        <div class="flex-1 flex flex-col min-w-0 overflow-hidden">
            <!-- Top Header -->
            <header class="h-16 bg-white border-b border-slate-200 flex items-center justify-between px-4 sm:px-6 lg:px-8">
                <button class="md:hidden p-2 rounded-lg text-slate-500 hover:bg-slate-100">
                    <i class='bx bx-menu text-2xl'></i>
                </button>
                
                <div class="flex items-center gap-4 ml-auto">
                    <!-- Notification Bell Button -->
                    <button class="relative p-1.5 text-gray-500 hover:text-indigo-600 hover:bg-slate-50 rounded-lg transition-colors" aria-label="Notifikasi">
                        <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor" class="w-5 h-5">
                            <path stroke-linecap="round" stroke-linejoin="round" d="M14.857 17.082a23.848 23.848 0 005.454-1.31A8.967 8.967 0 0118 9.75v-.7V9A6 6 0 006 9v.75a8.967 8.967 0 01-2.312 6.022c1.733.64 3.56 1.085 5.455 1.31m5.714 0a24.255 24.255 0 01-5.714 0m5.714 0a3 3 0 11-5.714 0" />
                        </svg>
                        <span class="absolute top-0 right-0 w-2 h-2 bg-red-500 rounded-full"></span>
                    </button>

                    <div class="flex items-center gap-3">
                        <div class="text-right hidden sm:block">
                            <p class="text-sm font-medium text-slate-900">Admin User</p>
                            <p class="text-xs text-slate-500">Administrator</p>
                        </div>
                        <div class="h-10 w-10 rounded-full bg-indigo-100 flex items-center justify-center text-indigo-600 font-bold">
                            A
                        </div>
                    </div>
                </div>
            </header>

            <!-- Page Content -->
            <main class="flex-1 overflow-y-auto p-4 sm:p-6 lg:p-8">
                <div class="max-w-7xl mx-auto">
