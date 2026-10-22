<?php
session_start();
?>
<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Pengelolaan Surat Tugas CDK Bojonegoro</title>
    <link rel="icon" type="image/svg+xml" href="assets/img/logo-jawa-timur.svg">
    <script src="https://cdn.tailwindcss.com"></script>
    <?php include __DIR__ . '/includes/tailwind-theme.php'; ?>
    <link href="https://unpkg.com/boxicons@2.1.4/css/boxicons.min.css" rel="stylesheet">
    <style>
        @keyframes landing-modal-in {
            from {
                opacity: 0;
                transform: translateY(14px) scale(0.985);
            }
            to {
                opacity: 1;
                transform: translateY(0) scale(1);
            }
        }

        #contohSuratModal:not(.hidden) .modal-panel {
            animation: landing-modal-in 0.22s ease-out both;
        }

        @media (prefers-reduced-motion: reduce) {
            #contohSuratModal:not(.hidden) .modal-panel {
                animation: none;
            }
        }
    </style>
</head>
<body class="bg-slate-50 font-sans text-slate-800 antialiased">
    <div class="relative flex min-h-screen flex-col overflow-hidden">
        <!-- Background decoration -->
        <div class="pointer-events-none absolute inset-0 -z-10 overflow-hidden">
            <div class="absolute -top-40 -right-40 h-96 w-96 rounded-full bg-primary-light opacity-60 blur-3xl"></div>
            <div class="absolute -left-24 top-40 h-72 w-72 rounded-full bg-primary-light opacity-50 blur-3xl"></div>
            <div class="absolute bottom-0 left-1/2 h-64 w-64 -translate-x-1/2 rounded-full bg-primary-light opacity-40 blur-3xl"></div>
        </div>

        <!-- Top navbar -->
        <header class="relative z-20">
            <div class="container mx-auto px-5 pt-5 sm:px-6">
                <nav class="flex items-center justify-between gap-4 rounded-2xl border border-white/70 bg-white/70 px-4 py-3 shadow-sm backdrop-blur-md sm:px-5">
                    <a href="index.php" class="group flex min-w-0 items-center gap-3">
                        <img src="assets/img/logo-jawa-timur.svg" alt="Lambang Provinsi Jawa Timur" class="h-11 w-auto shrink-0 object-contain drop-shadow-sm" decoding="async">

                        <span class="min-w-0">
                            <span class="block truncate text-sm font-bold text-slate-900 sm:text-base">CDK Wilayah Bojonegoro</span>
                            <span class="block truncate text-[10px] leading-tight text-slate-500 sm:text-xs">Dinas Kehutanan Provinsi Jawa Timur</span>
                        </span>
                    </a>

                    <div class="flex shrink-0 items-center gap-2 sm:gap-4">
                        <div class="hidden items-center gap-5 md:flex">
                            <a href="modules/surat-tugas.php" class="text-sm font-medium text-slate-600 transition-colors hover:text-primary">Panduan Penggunaan</a>
                            <a href="modules/template.php" class="text-sm font-medium text-slate-600 transition-colors hover:text-primary">Format Dishut</a>
                        </div>
                        <a href="modules/dashboard.php" class="inline-flex items-center gap-1.5 rounded-lg border border-primary-border bg-white/80 px-3 py-2 text-xs font-semibold text-primary transition-colors hover:border-primary hover:bg-primary-light sm:px-4 sm:text-sm">
                            <i class='bx bx-log-in-circle text-base'></i>
                            <span class="hidden sm:inline">Login Pegawai</span>
                            <span class="sm:hidden">Login</span>
                        </a>
                    </div>
                </nav>
            </div>
        </header>

        <!-- Hero -->
        <main class="relative z-10 flex flex-1 items-center">
            <div class="container mx-auto px-5 pb-14 pt-12 sm:px-6 lg:pb-20 lg:pt-16">
                <div class="grid items-center gap-12 lg:grid-cols-[minmax(0,1.15fr)_minmax(260px,0.85fr)] lg:gap-16">
                    <div class="max-w-2xl">
                        <h1 class="text-4xl font-bold leading-tight tracking-tight text-slate-900 sm:text-5xl xl:text-6xl">
                            <span class="block">Sistem Pengelolaan</span>
                            <span class="block">Surat Tugas</span>
                        </h1>

                        <p class="mt-6 max-w-2xl text-base leading-relaxed text-slate-600 sm:text-lg">
                            Aplikasi internal untuk penerbitan surat tugas dan pengelolaan buku nomor pegawai CDK Wilayah Bojonegoro, Dinas Kehutanan Provinsi Jawa Timur.
                        </p>

                        <div class="mt-8 flex max-w-xl flex-col items-stretch gap-3 sm:flex-row sm:items-center sm:gap-4">
                            <a href="modules/dashboard.php" class="inline-flex items-center justify-center gap-2 rounded-lg bg-primary px-6 py-3.5 text-base font-semibold text-white shadow-sm transition-colors hover:bg-primary-dark sm:px-8 sm:py-4 sm:text-lg">
                                Masuk ke Dashboard
                                <i class='bx bx-right-arrow-alt text-2xl'></i>
                            </a>
                            <button type="button" id="lihatContohBtn" aria-haspopup="dialog" aria-controls="contohSuratModal" class="inline-flex items-center justify-center gap-2 rounded-lg border border-primary px-6 py-3.5 text-base font-semibold text-primary transition-colors hover:bg-primary-light sm:px-8 sm:py-4 sm:text-lg">
                                <i class='bx bx-file text-xl'></i>
                                Lihat Contoh Surat
                            </button>
                        </div>
                    </div>

                    <div class="relative hidden items-center justify-center lg:flex" aria-hidden="true">
                        <div class="absolute h-72 w-72 rounded-full bg-primary-light opacity-80 blur-3xl"></div>
                        <img src="assets/img/logo-jawa-timur.svg" alt="" class="relative h-64 w-auto object-contain drop-shadow-lg xl:h-72" loading="eager" decoding="async">
                    </div>
                </div>
            </div>
        </main>
    </div>

    <!-- Modal contoh surat -->
    <div id="contohSuratModal" class="fixed inset-0 z-50 hidden items-center justify-center p-3 sm:p-6" role="dialog" aria-modal="true" aria-labelledby="contohSuratTitle">
        <button type="button" tabindex="-1" class="absolute inset-0 bg-slate-900/60 backdrop-blur-sm" data-modal-close aria-label="Tutup contoh surat"></button>

        <div class="modal-panel relative z-10 max-h-[92vh] w-full max-w-4xl overflow-y-auto rounded-2xl bg-slate-100 shadow-2xl ring-1 ring-white/20">
            <div class="sticky top-0 z-20 flex items-center justify-between gap-3 border-b border-slate-200 bg-white/95 px-4 py-3 backdrop-blur sm:px-6">
                <div>
                    <h2 id="contohSuratTitle" class="text-base font-bold text-slate-900 sm:text-lg">Contoh Surat Tugas</h2>
                    <p class="text-xs text-slate-500">Pratinjau format resmi Surat Tugas CDK Wilayah Bojonegoro.</p>
                </div>
                <button type="button" class="flex h-9 w-9 shrink-0 items-center justify-center rounded-lg text-slate-500 transition-colors hover:bg-slate-100 hover:text-slate-800" data-modal-close aria-label="Tutup">
                    <i class='bx bx-x text-2xl'></i>
                </button>
            </div>

            <div class="p-3 sm:p-6">
                <article class="mx-auto max-w-3xl bg-white px-5 py-6 shadow-sm ring-1 ring-slate-200 sm:px-10 sm:py-10">
                    <!-- Kop surat -->
                    <div class="border-b-[3px] border-double border-slate-800 pb-3">
                        <div class="flex flex-col items-center justify-center gap-2 text-center sm:flex-row sm:gap-4">
                            <img src="assets/img/logo-jawa-timur.svg" alt="Lambang Provinsi Jawa Timur" class="h-16 w-auto shrink-0 object-contain" decoding="async">

                            <div class="leading-tight">
                                <p class="text-[11px] font-semibold tracking-wide text-slate-700 sm:text-xs">PEMERINTAH PROVINSI JAWA TIMUR</p>
                                <p class="text-base font-bold text-slate-900 sm:text-lg">DINAS KEHUTANAN</p>
                                <p class="text-sm font-bold text-slate-800 sm:text-base">UPTD CDK WILAYAH BOJONEGORO</p>
                                <p class="text-[10px] text-slate-500 sm:text-xs">Bojonegoro, Jawa Timur</p>
                            </div>
                        </div>
                    </div>

                    <!-- Judul surat -->
                    <div class="mt-5 text-center sm:mt-6">
                        <p class="text-base font-bold tracking-[0.2em] text-slate-900 underline underline-offset-4 sm:text-lg">SURAT TUGAS</p>
                        <p class="mt-1 text-xs text-slate-600 sm:text-sm">Nomor: 094/ST/CDK-BJN/X/2026</p>
                    </div>

                    <!-- Isi surat -->
                    <div class="mt-5 space-y-3 text-sm leading-relaxed text-slate-700">
                        <p>
                            Kepala UPTD CDK Wilayah Bojonegoro Dinas Kehutanan Provinsi Jawa Timur, berdasarkan kebutuhan pelaksanaan tugas kedinasan, dengan ini menugaskan kepada:
                        </p>

                        <div class="overflow-x-auto">
                            <table class="w-full min-w-[420px] text-left text-xs sm:text-sm">
                                <tbody class="divide-y divide-slate-100">
                                    <tr>
                                        <th class="w-36 py-2 pr-3 align-top font-semibold text-slate-600">Nama</th>
                                        <td class="py-2 font-medium text-slate-800">Purwo Yulianto, S.P.</td>
                                    </tr>
                                    <tr>
                                        <th class="py-2 pr-3 align-top font-semibold text-slate-600">NIP</th>
                                        <td class="py-2 text-slate-700">19870512 201001 1 002</td>
                                    </tr>
                                    <tr>
                                        <th class="py-2 pr-3 align-top font-semibold text-slate-600">Jabatan</th>
                                        <td class="py-2 text-slate-700">Penyuluh Kehutanan Ahli Muda</td>
                                    </tr>
                                    <tr>
                                        <th class="py-2 pr-3 align-top font-semibold text-slate-600">Untuk</th>
                                        <td class="py-2 text-slate-700">Melaksanakan penyuluhan kehutanan dan pendampingan kelompok tani hutan.</td>
                                    </tr>
                                    <tr>
                                        <th class="py-2 pr-3 align-top font-semibold text-slate-600">Lokasi</th>
                                        <td class="py-2 text-slate-700">Kecamatan Ngraho, Kabupaten Bojonegoro</td>
                                    </tr>
                                    <tr>
                                        <th class="py-2 pr-3 align-top font-semibold text-slate-600">Tanggal Pelaksanaan</th>
                                        <td class="py-2 font-medium text-slate-800">25 – 27 Oktober 2026</td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>

                        <p>
                            Demikian Surat Tugas ini dibuat untuk dilaksanakan dengan penuh tanggung jawab dan melaporkan hasilnya kepada Kepala UPTD CDK Wilayah Bojonegoro.
                        </p>
                    </div>

                    <!-- Tanda tangan -->
                    <div class="mt-6 flex justify-end sm:mt-8">
                        <div class="w-64 text-center text-xs text-slate-700 sm:text-sm">
                            <p>Bojonegoro, 22 Oktober 2026</p>
                            <p class="mt-1 font-medium">Kepala UPTD CDK Wilayah Bojonegoro</p>
                            <div class="h-16 sm:h-20"></div>
                            <p class="font-semibold text-slate-900 underline underline-offset-2">...............................</p>
                            <p class="mt-1 text-[11px] text-slate-500 sm:text-xs">NIP. ........................</p>
                        </div>
                    </div>
                </article>
            </div>

            <div class="sticky bottom-0 z-20 flex flex-col-reverse gap-2 border-t border-slate-200 bg-white/95 px-4 py-3 backdrop-blur sm:flex-row sm:justify-end sm:px-6">
                <button type="button" class="inline-flex items-center justify-center rounded-lg border border-slate-300 bg-white px-5 py-2.5 text-sm font-semibold text-slate-700 transition-colors hover:bg-slate-50" data-modal-close>
                    Tutup
                </button>
                <a href="modules/dashboard.php" class="inline-flex items-center justify-center gap-2 rounded-lg bg-primary px-5 py-2.5 text-sm font-semibold text-white transition-colors hover:bg-primary-dark">
                    Masuk untuk Buat Surat
                    <i class='bx bx-right-arrow-alt text-lg'></i>
                </a>
            </div>
        </div>
    </div>

    <script>
        (function () {
            const modal = document.getElementById('contohSuratModal');
            const openBtn = document.getElementById('lihatContohBtn');
            let lastFocusedElement = null;

            if (!modal || !openBtn) return;

            function openModal() {
                lastFocusedElement = document.activeElement;
                modal.classList.remove('hidden');
                modal.classList.add('flex');
                document.body.classList.add('overflow-hidden');
                const closeBtn = modal.querySelector('button[aria-label="Tutup"]');
                if (closeBtn) closeBtn.focus();
            }

            function closeModal() {
                modal.classList.add('hidden');
                modal.classList.remove('flex');
                document.body.classList.remove('overflow-hidden');
                if (lastFocusedElement && typeof lastFocusedElement.focus === 'function') {
                    lastFocusedElement.focus();
                }
            }

            openBtn.addEventListener('click', openModal);

            modal.querySelectorAll('[data-modal-close]').forEach(function (element) {
                element.addEventListener('click', closeModal);
            });

            document.addEventListener('keydown', function (event) {
                if (event.key === 'Escape' && !modal.classList.contains('hidden')) {
                    closeModal();
                }
            });

            modal.addEventListener('keydown', function (event) {
                if (event.key !== 'Tab' || modal.classList.contains('hidden')) return;

                const focusable = Array.from(modal.querySelectorAll('a[href], button:not([disabled])'))
                    .filter(function (element) {
                        return element.offsetParent !== null;
                    });
                if (!focusable.length) return;

                const first = focusable[0];
                const last = focusable[focusable.length - 1];

                if (event.shiftKey && document.activeElement === first) {
                    event.preventDefault();
                    last.focus();
                } else if (!event.shiftKey && document.activeElement === last) {
                    event.preventDefault();
                    first.focus();
                }
            });
        })();
    </script>
</body>
</html>
