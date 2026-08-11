# Task 1: Design System CSS Global

## Context
Proyek: SIPENSURAT — Sistem Pengelolaan Surat Tugas (PHP + Tailwind CDN)
Workspace: `d:\laragon\www\spt-php`
Task ini adalah fondasi dari redesign — semua task berikutnya bergantung pada file CSS yang dibuat di sini.

## Global Constraints (WAJIB DIPATUHI)
- JANGAN ubah logika PHP backend (query, PDO, CRUD, validasi, session)
- JANGAN ubah struktur routing/URL dan nama file
- Perubahan HANYA pada: HTML markup class Tailwind, CSS, dan config visual Chart.js

## Yang Harus Dikerjakan

### File 1: `assets/css/app.css` (BUAT BARU / TIMPA)

Buat file CSS design system dengan konten TEPAT seperti di bawah ini:

```css
/* =============================================
   SIPENSURAT — Design System CSS
   ============================================= */

/* --- Stat Card (putih, border, aksen garis kiri) --- */
.stat-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 0.75rem;
    box-shadow: 0 1px 3px 0 rgba(0,0,0,.05);
    padding: 1.5rem;
    transition: box-shadow .15s ease, transform .15s ease;
    position: relative;
    overflow: hidden;
}
.stat-card:hover {
    box-shadow: 0 4px 12px 0 rgba(0,0,0,.08);
    transform: translateY(-1px);
}
.stat-card::before {
    content: '';
    position: absolute;
    left: 0; top: 0; bottom: 0;
    width: 3px;
    border-radius: 3px 0 0 3px;
}
.stat-card.blue::before   { background: #4f46e5; }
.stat-card.indigo::before { background: #6366f1; }
.stat-card.violet::before { background: #8b5cf6; }
.stat-card.teal::before   { background: #0d9488; }
.stat-card.emerald::before{ background: #059669; }

.stat-icon {
    width: 2.5rem; height: 2.5rem;
    border-radius: 0.625rem;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.2rem;
    flex-shrink: 0;
}
.stat-icon.blue   { background:#eef2ff; color:#4f46e5; }
.stat-icon.indigo { background:#ede9fe; color:#6366f1; }
.stat-icon.violet { background:#f5f3ff; color:#8b5cf6; }
.stat-icon.teal   { background:#f0fdfa; color:#0d9488; }
.stat-icon.emerald{ background:#ecfdf5; color:#059669; }

/* --- Sidebar active state dengan garis kiri --- */
.sidebar-link {
    display: flex; align-items: center; gap: 0.75rem;
    padding: 0.625rem 0.75rem;
    border-radius: 0.5rem;
    color: #6b7280;
    font-size: 0.875rem;
    font-weight: 500;
    transition: background .15s ease, color .15s ease;
    position: relative;
    text-decoration: none;
}
.sidebar-link:hover {
    background: #f9fafb;
    color: #4f46e5;
}
.sidebar-link.active {
    background: #eef2ff;
    color: #4f46e5;
    font-weight: 600;
}
.sidebar-link.active::before {
    content: '';
    position: absolute;
    left: -12px; top: 6px; bottom: 6px;
    width: 3px;
    background: #4f46e5;
    border-radius: 0 3px 3px 0;
}

/* --- Tabel premium --- */
.table-premium thead tr {
    background: #f9fafb;
}
.table-premium thead th {
    padding: 0.75rem 1.25rem;
    font-size: 0.7rem;
    font-weight: 600;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    border-bottom: 1px solid #e5e7eb;
}
.table-premium tbody tr {
    transition: background .1s ease;
}
.table-premium tbody tr:hover {
    background: #f9fafb;
}
.table-premium tbody td {
    padding: 0.75rem 1.25rem;
    font-size: 0.875rem;
    color: #374151;
    border-bottom: 1px solid #f3f4f6;
}

/* --- Badge pill --- */
.badge {
    display: inline-flex; align-items: center;
    padding: 0.2rem 0.6rem;
    border-radius: 9999px;
    font-size: 0.7rem;
    font-weight: 600;
    letter-spacing: 0.02em;
    white-space: nowrap;
}
.badge-green   { background:#dcfce7; color:#166534; }
.badge-blue    { background:#dbeafe; color:#1d4ed8; }
.badge-indigo  { background:#e0e7ff; color:#3730a3; }
.badge-gray    { background:#f3f4f6; color:#4b5563; }
.badge-emerald { background:#d1fae5; color:#065f46; }
.badge-amber   { background:#fef3c7; color:#92400e; }
.badge-red     { background:#fee2e2; color:#991b1b; }
.badge-violet  { background:#f5f3ff; color:#5b21b6; }

/* --- Tombol aksi outline --- */
.btn-action {
    display: inline-flex; align-items: center; gap: 0.25rem;
    padding: 0.25rem 0.625rem;
    font-size: 0.75rem;
    font-weight: 500;
    border: 1px solid #d1d5db;
    border-radius: 0.375rem;
    color: #374151;
    background: transparent;
    cursor: pointer;
    transition: background .1s ease, border-color .1s ease;
    text-decoration: none;
}
.btn-action:hover { background: #f9fafb; border-color: #9ca3af; }
.btn-action.primary { border-color: #c7d2fe; color: #4f46e5; }
.btn-action.primary:hover { background: #eef2ff; }
.btn-action.warning { border-color: #fcd34d; color: #d97706; }
.btn-action.warning:hover { background: #fffbeb; }
.btn-action.danger  { border-color: #fca5a5; color: #dc2626; }
.btn-action.danger:hover  { background: #fef2f2; }

/* --- Card container --- */
.card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 0.75rem;
    box-shadow: 0 1px 2px 0 rgba(0,0,0,.04);
}

/* --- Page header --- */
.page-header-title {
    font-size: 1.5rem;
    font-weight: 700;
    color: #111827;
    letter-spacing: -0.025em;
    line-height: 1.2;
}
.page-header-sub {
    font-size: 0.875rem;
    color: #6b7280;
    margin-top: 0.25rem;
}

/* --- Avatar inisial rotasi warna --- */
.avatar {
    width:2rem; height:2rem; border-radius:9999px;
    display:flex; align-items:center; justify-content:center;
    font-size:0.7rem; font-weight:700; flex-shrink:0;
}
.avatar-0 { background:#dbeafe; color:#1d4ed8; }
.avatar-1 { background:#ede9fe; color:#6d28d9; }
.avatar-2 { background:#dcfce7; color:#166534; }
.avatar-3 { background:#fef3c7; color:#92400e; }
.avatar-4 { background:#fce7f3; color:#9d174d; }

/* --- Input focus states --- */
.input-premium {
    width: 100%;
    border: 1px solid #d1d5db;
    border-radius: 0.5rem;
    padding: 0.5rem 0.75rem;
    font-size: 0.875rem;
    color: #111827;
    background: #ffffff;
    outline: none;
    transition: border-color .15s, box-shadow .15s;
}
.input-premium:focus {
    border-color: #6366f1;
    box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.12);
}

/* --- Pagination --- */
.pagination-btn {
    display: inline-flex; align-items: center; justify-content: center;
    min-width: 2rem; height: 2rem;
    padding: 0 0.5rem;
    border: 1px solid #e5e7eb;
    border-radius: 0.375rem;
    font-size: 0.8rem;
    color: #6b7280;
    background: #ffffff;
    cursor: pointer;
    text-decoration: none;
    transition: background .1s, border-color .1s;
}
.pagination-btn:hover { background: #f9fafb; border-color: #d1d5db; }
.pagination-btn.active { background: #4f46e5; color: #ffffff; border-color: #4f46e5; }
.pagination-btn:disabled, .pagination-btn.disabled { opacity: 0.4; cursor: not-allowed; }
```

### File 2: `includes/header.php` (MODIFIKASI)

Baca file ini terlebih dahulu lalu lakukan perubahan berikut:

1. **Di dalam `<head>`, setelah baris `<script src="...chart.js">`, tambahkan:**
```html
    <!-- Google Fonts: Inter -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
    <!-- Design System -->
    <link rel="stylesheet" href="../assets/css/app.css">
```

2. **Ganti blok `<script>tailwind.config...` dengan:**
```html
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'ui-sans-serif', 'system-ui', 'sans-serif'],
                    },
                    colors: {
                        primary: '#4f46e5',
                        'primary-light': '#eef2ff',
                    }
                }
            }
        }
    </script>
```

3. **Ganti class `<body>` dari `bg-slate-50 font-sans text-slate-800 antialiased` menjadi:**
```html
<body class="bg-gray-50 font-sans text-gray-900 antialiased">
```

## Cara Kerja

1. Baca file `includes/header.php` untuk memahami strukturnya
2. Buat file `assets/css/app.css` dengan konten persis seperti di atas
3. Edit `includes/header.php` untuk menambahkan Google Fonts, link CSS, update Tailwind config, dan update body class
4. Verifikasi kedua file tersimpan dengan benar

## Report

Tulis laporan ke file: `d:\laragon\www\spt-php\.superpowers\sdd\task-1-report.md`

Laporan harus mencakup:
- Status: DONE / DONE_WITH_CONCERNS / BLOCKED
- File yang dibuat/dimodifikasi
- Verifikasi singkat bahwa file tersimpan benar

Kembalikan hanya: status, file yang diubah, dan apakah ada concerns.
