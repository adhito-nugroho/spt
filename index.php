<?php
session_start();
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Pengelolaan Surat Tugas</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://unpkg.com/boxicons@2.1.4/css/boxicons.min.css" rel="stylesheet">
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
<body class="bg-slate-50 font-sans text-slate-800">
    
    <div class="min-h-screen flex items-center justify-center relative overflow-hidden">
        <!-- Background Decoration -->
        <div class="absolute top-0 left-0 w-full h-full overflow-hidden -z-10">
            <div class="absolute -top-40 -right-40 w-96 h-96 rounded-full bg-indigo-100 blur-3xl opacity-50"></div>
            <div class="absolute top-40 -left-20 w-72 h-72 rounded-full bg-blue-100 blur-3xl opacity-50"></div>
        </div>

        <div class="container mx-auto px-6 py-12">
            <div class="flex flex-col md:flex-row items-center gap-12">
                <div class="md:w-1/2 space-y-8">
                    <div class="inline-flex items-center gap-2 px-4 py-2 rounded-full bg-indigo-50 text-indigo-700 text-sm font-medium">
                        <span class="relative flex h-3 w-3">
                          <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-indigo-400 opacity-75"></span>
                          <span class="relative inline-flex rounded-full h-3 w-3 bg-indigo-500"></span>
                        </span>
                        Sistem Persuratan CDK Bojonegoro
                    </div>
                    
                    <h1 class="text-5xl md:text-6xl font-bold text-slate-900 leading-tight">
                        Sistem Pengelolaan <br>
                        <span class="text-indigo-600">Surat Tugas</span>
                    </h1>
                    
                    <p class="text-lg text-slate-600 leading-relaxed max-w-lg">
                        Kelola surat tugas dengan mudah, cepat, dan efisien. Generate surat tugas otomatis dengan template yang sudah ditentukan, hemat waktu dan tenaga.
                    </p>
                    
                    <div class="flex flex-col sm:flex-row items-stretch sm:items-center gap-4">
                        <a href="modules/dashboard.php" class="inline-flex items-center justify-center gap-2 px-8 py-4 bg-indigo-600 hover:bg-indigo-700 text-white text-lg font-semibold rounded-xl transition-all shadow-lg hover:shadow-indigo-200 transform hover:-translate-y-1">
                            Masuk ke Dashboard
                            <i class='bx bx-right-arrow-alt text-2xl'></i>
                        </a>
                        <a href="modules/surat-tugas.php" class="inline-flex items-center justify-center gap-2 px-8 py-4 border-2 border-indigo-600 text-indigo-600 hover:bg-indigo-50 text-lg font-semibold rounded-xl transition-all">
                            <i class='bx bx-file text-xl'></i>
                            Lihat Contoh Surat
                        </a>
                    </div>
                    
                    <div class="pt-8 border-t border-slate-200 flex items-center gap-8 text-slate-500">
                        <div class="flex items-center gap-2">
                            <i class='bx bx-check-circle text-indigo-500 text-xl'></i>
                            <span class="text-sm">500+ Surat Diproses</span>
                        </div>
                        <div class="flex items-center gap-2">
                            <i class='bx bx-check-circle text-indigo-500 text-xl'></i>
                            <span class="text-sm">Format Sesuai Ketentuan Dishut Jatim</span>
                        </div>
                    </div>
                </div>
                
                <div class="md:w-1/2 flex justify-center relative">
                    <!-- Decorative watermark dokumen di belakang card -->
                    <div class="absolute top-6 right-4 w-full h-full -z-10 flex items-center justify-center opacity-[0.07] text-indigo-600 pointer-events-none" style="transform:rotate(6deg)">
                        <i class='bx bxs-file-doc' style="font-size:280px;"></i>
                    </div>
                    <div class="absolute -bottom-4 -left-4 w-full h-full -z-10 flex items-end justify-start opacity-[0.05] text-indigo-600 pointer-events-none" style="transform:rotate(-4deg)">
                        <i class='bx bxs-folder-open' style="font-size:200px;"></i>
                    </div>

                    <!-- Card surat tugas -->
                    <div class="relative z-10 bg-white p-8 rounded-2xl shadow-2xl transform rotate-2 hover:rotate-0 transition-all duration-500 border border-slate-100 w-full max-w-sm">
                        <!-- Header card — TIDAK DIUBAH -->
                        <div class="flex items-center gap-4 mb-6 border-b border-slate-100 pb-4">
                            <div class="w-12 h-12 bg-indigo-100 rounded-full flex items-center justify-center text-indigo-600">
                                <i class='bx bxs-file-doc text-2xl'></i>
                            </div>
                            <div>
                                <h3 class="font-bold text-slate-800">Surat Tugas Baru</h3>
                                <p class="text-xs text-slate-500">Dibuat pada <?php echo date('d M Y'); ?></p>
                            </div>
                        </div>

                        <!-- Mock field data surat tugas -->
                        <div class="space-y-4">
                            <div>
                                <p class="text-xs text-slate-400 uppercase tracking-wide mb-0.5">Nomor Surat</p>
                                <p class="text-sm font-medium text-slate-700">094/ST/CDK-BJN/VIII/2026</p>
                            </div>
                            <div>
                                <p class="text-xs text-slate-400 uppercase tracking-wide mb-0.5">Nama Pegawai</p>
                                <p class="text-sm font-medium text-slate-700">Purwo Yulianto, S.P.</p>
                            </div>
                            <div>
                                <p class="text-xs text-slate-400 uppercase tracking-wide mb-0.5">Tujuan</p>
                                <p class="text-sm font-medium text-slate-700">Penyuluhan kehutanan di Kec. Ngraho</p>
                            </div>
                            <div>
                                <p class="text-xs text-slate-400 uppercase tracking-wide mb-0.5">Tanggal Pelaksanaan</p>
                                <p class="text-sm font-medium text-slate-700">11 – 13 Agustus 2026</p>
                            </div>
                        </div>

                        <!-- Tombol aksi -->
                        <div class="mt-6 flex justify-end">
                            <a href="modules/dashboard.php" class="inline-flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white text-sm font-medium rounded-lg transition-colors">
                                <i class='bx bx-printer'></i>
                                Cetak Surat
                            </a>
                        </div>
                    </div>

                    <div class="absolute -bottom-10 -left-10 w-24 h-24 bg-yellow-400 rounded-full blur-2xl opacity-20"></div>
                </div>
            </div>
        </div>
    </div>
</body>
</html>