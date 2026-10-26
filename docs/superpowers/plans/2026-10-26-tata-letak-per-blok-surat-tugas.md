# Tata Letak Per Blok Preview Surat Tugas Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Menambahkan kontrol penyesuaian tata letak per blok (Kop, Judul & Nomor, Dasar, Memerintahkan, Pegawai & Jarak per pegawai, Untuk, dan Tanda Tangan) pada halaman Preview Surat Tugas dengan UI akordion dan manipulasi XML DOCX / PDF yang presisi.

**Architecture:** Memperluas skema parameter layout `spt_layout_*` di `modules/generate-surat.php` dan `modules/save-template-layout.php`, menyempurnakan parsing AST `word/document.xml` untuk mengenali dan mengontrol spasi setiap blok secara semantis, dan merestrukturisasi UI panel kontrol di `generate-surat.php` ke dalam kelompok akordion (collapsible) yang rapi.

**Tech Stack:** PHP 8.x, PHPWord (TemplateProcessor & IOFactory), DOMDocument / DOMXPath (XML manipulation), Tailwind CSS, Vanilla JavaScript, MySQL (PDO).

## Global Constraints
- Satuan spasi blok menggunakan piksel (`px`), margin halaman menggunakan milimeter (`mm`), font menggunakan point (`pt`), dan line height menggunakan multiplier.
- Semua nilai parameter harus divalidasi dan di-clamp sesuai rentang yang ditentukan sebelum ditulis ke DOCX atau disimpan ke database.
- File DOCX hasil manipulasi harus tetap valid dan dapat dibuka tanpa corrupt oleh Microsoft Word, LibreOffice, dan PHPWord.
- Nilai layout harus tersimpan dalam `templates.layout_settings` (JSON) dan otomatis dimuat saat surat dibuka.

---

### Task 1: Update Parameter Layout Config, Helper Functions & Storage Endpoint

**Files:**
- Modify: `modules/generate-surat.php:237-360`
- Modify: `modules/save-template-layout.php:33-65`
- Test: `tests/test_layout_config.php`

**Interfaces:**
- Produces: `spt_layout_default_values()`, `spt_layout_ranges()`, `spt_layout_request_settings()`, `spt_layout_query_params()`, `spt_layout_normalize_stored_settings()` containing all 21 parameters.

- [ ] **Step 1: Write the failing test**

Create `tests/test_layout_config.php`:
```php
<?php
require_once __DIR__ . '/../modules/generate-surat.php';

$defaults = spt_layout_default_values();
$ranges = spt_layout_ranges();

$expectedKeys = [
    'margin_top', 'margin_bottom', 'margin_left', 'margin_right',
    'font_size', 'line_height',
    'kop_bottom_spacing',
    'judul_top_spacing', 'nomor_surat_spacing', 'judul_bottom_spacing',
    'dasar_top_spacing', 'dasar_item_spacing', 'dasar_bottom_spacing',
    'memerintahkan_spacing',
    'pegawai_top_spacing', 'pegawai_gap_spacing', 'pegawai_line_spacing',
    'untuk_top_spacing', 'untuk_item_spacing', 'untuk_bottom_spacing',
    'ttd_offset', 'ttd_x', 'ttd_date_to_jabatan', 'ttd_space_height', 'ttd_name_nip_spacing',
    'lock_ttd'
];

foreach ($expectedKeys as $key) {
    if (!array_key_exists($key, $defaults)) {
        fwrite(STDERR, "FAIL: Key '$key' not found in defaults\n");
        exit(1);
    }
    if ($key !== 'lock_ttd' && !array_key_exists($key, $ranges)) {
        fwrite(STDERR, "FAIL: Key '$key' not found in ranges\n");
        exit(1);
    }
}

// Test clamping and normalization
$mockRaw = [
    'kop_bottom_spacing' => 999, // Should clamp to 50
    'judul_top_spacing' => -999, // Should clamp to -15
    'pegawai_gap_spacing' => 20,
    'lock_ttd' => 1
];
$norm = spt_layout_normalize_stored_settings($mockRaw);
if ($norm['kop_bottom_spacing'] != 50 || $norm['judul_top_spacing'] != -15 || $norm['pegawai_gap_spacing'] != 20 || $norm['lock_ttd'] !== 1) {
    fwrite(STDERR, "FAIL: Normalization / clamping unexpected result\n");
    exit(1);
}

echo "PASS: Layout config test passed.\n";
```

- [ ] **Step 2: Run test to verify it fails**

Run: `php tests/test_layout_config.php`  
Expected: FAIL with missing key errors.

- [ ] **Step 3: Update `generate-surat.php` and `save-template-layout.php`**

Update `spt_layout_default_values()` and `spt_layout_ranges()` in `modules/generate-surat.php`:
```php
function spt_layout_default_values()
{
    return [
        'margin_top' => 15,
        'margin_bottom' => 15,
        'margin_left' => 20,
        'margin_right' => 20,
        'font_size' => 11,
        'line_height' => 1.0,
        'kop_bottom_spacing' => 10,
        'judul_top_spacing' => 0,
        'nomor_surat_spacing' => 4,
        'judul_bottom_spacing' => 10,
        'dasar_top_spacing' => 6,
        'dasar_item_spacing' => 3,
        'dasar_bottom_spacing' => 8,
        'memerintahkan_spacing' => 8,
        'pegawai_top_spacing' => 6,
        'pegawai_gap_spacing' => 8,
        'pegawai_line_spacing' => 2,
        'untuk_top_spacing' => 6,
        'untuk_item_spacing' => 3,
        'untuk_bottom_spacing' => 10,
        'ttd_offset' => 10,
        'ttd_x' => 0,
        'ttd_date_to_jabatan' => 4,
        'ttd_space_height' => 60,
        'ttd_name_nip_spacing' => 4,
        'lock_ttd' => 0,
    ];
}

function spt_layout_ranges()
{
    return [
        'margin_top' => [5, 40, 0.5],
        'margin_bottom' => [5, 40, 0.5],
        'margin_left' => [10, 40, 0.5],
        'margin_right' => [10, 40, 0.5],
        'font_size' => [9, 13, 0.5],
        'line_height' => [1.0, 2.0, 0.05],
        'kop_bottom_spacing' => [0, 50, 1],
        'judul_top_spacing' => [-15, 40, 1],
        'nomor_surat_spacing' => [0, 30, 1],
        'judul_bottom_spacing' => [0, 40, 1],
        'dasar_top_spacing' => [0, 30, 1],
        'dasar_item_spacing' => [0, 20, 1],
        'dasar_bottom_spacing' => [0, 30, 1],
        'memerintahkan_spacing' => [0, 30, 1],
        'pegawai_top_spacing' => [0, 30, 1],
        'pegawai_gap_spacing' => [0, 35, 1],
        'pegawai_line_spacing' => [0, 15, 1],
        'untuk_top_spacing' => [0, 30, 1],
        'untuk_item_spacing' => [0, 20, 1],
        'untuk_bottom_spacing' => [0, 30, 1],
        'ttd_offset' => [-20, 80, 1],
        'ttd_x' => [-50, 50, 1],
        'ttd_date_to_jabatan' => [0, 25, 1],
        'ttd_space_height' => [20, 120, 2],
        'ttd_name_nip_spacing' => [0, 25, 1],
    ];
}
```

Update `modules/save-template-layout.php`:
Update the `$ranges` array with the exact same keys and ranges as `spt_layout_ranges()`.

- [ ] **Step 4: Run test to verify it passes**

Run: `php tests/test_layout_config.php`  
Expected: `PASS: Layout config test passed.`

- [ ] **Step 5: Commit**

```bash
git add modules/generate-surat.php modules/save-template-layout.php tests/test_layout_config.php
git commit -m "feat(layout): expand layout settings schema with per-block parameters"
```

---

### Task 2: Implement Per-Block Word Document XML Manipulation & Dompdf Fallback

**Files:**
- Modify: `modules/generate-surat.php:650-1000`
- Test: `tests/test_layout_docx_processing.php`

**Interfaces:**
- Consumes: `$settings` array containing all block spacing parameters.
- Produces: `spt_layout_apply_to_docx($docxPath, array $settings)`, `spt_layout_build_css(array $settings)`, `spt_layout_inject_html($html, array $settings)`.

- [ ] **Step 1: Write the failing test**

Create `tests/test_layout_docx_processing.php`:
```php
<?php
require_once __DIR__ . '/../modules/generate-surat.php';

$sourceDocx = __DIR__ . '/../assets/templates/1766636680_template_satu_orang.docx';
if (!file_exists($sourceDocx)) {
    echo "SKIP: Template not found\n";
    exit(0);
}

$tmpDocx = sys_get_temp_dir() . '/test_layout_' . uniqid() . '.docx';
copy($sourceDocx, $tmpDocx);

$settings = spt_layout_default_values();
$settings['kop_bottom_spacing'] = 25;
$settings['judul_top_spacing'] = 15;
$settings['pegawai_gap_spacing'] = 18;
$settings['ttd_space_height'] = 75;

$result = spt_layout_apply_to_docx($tmpDocx, $settings);
if (!$result) {
    fwrite(STDERR, "FAIL: spt_layout_apply_to_docx returned false\n");
    @unlink($tmpDocx);
    exit(1);
}

// Verify that tmpDocx can still be opened and parsed as XML
$zip = new ZipArchive();
if ($zip->open($tmpDocx) !== true) {
    fwrite(STDERR, "FAIL: Resulting DOCX is not a valid zip\n");
    @unlink($tmpDocx);
    exit(1);
}
$xml = $zip->getFromName('word/document.xml');
$zip->close();
@unlink($tmpDocx);

$dom = new DOMDocument();
if (!$dom->loadXML($xml)) {
    fwrite(STDERR, "FAIL: Resulting document.xml is invalid XML\n");
    exit(1);
}

echo "PASS: DOCX per-block layout applied successfully.\n";
```

- [ ] **Step 2: Run test to verify it fails or tests existing behavior**

Run: `php tests/test_layout_docx_processing.php`

- [ ] **Step 3: Implement per-block XML manipulation in `spt_layout_apply_to_docx`**

In `modules/generate-surat.php`:
1. Traverse paragraphs and categorize each paragraph into its respective block:
   - **Kop block**: Paragraf sebelum "SURAT TUGAS". Paragraf terakhir kop diberikan `w:after` sebesar `kop_bottom_spacing`.
   - **Judul block**:
     - Paragraf "SURAT TUGAS": `w:before` = `judul_top_spacing`, `w:after` = `nomor_surat_spacing`.
     - Paragraf "Nomor :": `w:after` = `judul_bottom_spacing`.
   - **Dasar block**:
     - Paragraf pembuka "DASAR :": `w:before` = `dasar_top_spacing`.
     - Paragraf butir dasar (berisi penomoran `1.`, `2.`, dll. sebelum MEMERINTAHKAN): `w:after` = `dasar_item_spacing`.
     - Paragraf terakhir dasar: `w:after` = `dasar_bottom_spacing`.
   - **Memerintahkan block**:
     - Paragraf "MEMERINTAHKAN :": `w:before` = `memerintahkan_spacing`, `w:after` = `memerintahkan_spacing`.
   - **Pegawai block**:
     - Paragraf "KEPADA :": `w:before` = `pegawai_top_spacing`.
     - Untuk setiap pegawai berikutnya (misal teks diawali angka urut atau nama/nip pegawai berikutnya): diberikan `w:before` = `pegawai_gap_spacing`.
     - Baris-baris identitas di dalam 1 pegawai (Nama, NIP, Pangkat, Jabatan): diberikan spasi `w:after` = `pegawai_line_spacing`.
     - Untuk tabel lampiran jika ada tabel pegawai: atur tinggi baris `w:trHeight` atau padding cell sesuai `pegawai_gap_spacing`.
   - **Untuk block**:
     - Paragraf pembuka "UNTUK :": `w:before` = `untuk_top_spacing`.
     - Paragraf butir tugas (`1.`, `2.`, dst): `w:after` = `untuk_item_spacing`.
     - Paragraf penutup/terakhir tugas: `w:after` = `untuk_bottom_spacing`.
   - **Tanda Tangan block**:
     - Baris tempat & tanggal (`Bojonegoro, ...`): `w:before` = `ttd_offset`, `w:after` = `ttd_date_to_jabatan`.
     - Paragraf jabatan: `w:after` = `ttd_space_height`.
     - Paragraf nama penandatangan & NIP: `w:before` = `ttd_name_nip_spacing`.
     - Posisi horizontal `ttd_x`: tambahkan ke indentasi `w:left`.
     - Opsi `lock_ttd`: sematkan flag `keepLines` dan `keepNext`.
2. Update `spt_layout_build_css()` dan `spt_layout_inject_html()` agar mendukung class-class styling Dompdf per blok yang bersesuaian.

- [ ] **Step 4: Run test to verify it passes**

Run: `php tests/test_layout_docx_processing.php`  
Expected: `PASS: DOCX per-block layout applied successfully.`

- [ ] **Step 5: Commit**

```bash
git add modules/generate-surat.php tests/test_layout_docx_processing.php
git commit -m "feat(layout): implement per-block XML and Dompdf layout adjustments"
```

---

### Task 3: Build Modern UI Accordions and Controls in Preview Page

**Files:**
- Modify: `modules/generate-surat.php:1880-2460`

**Interfaces:**
- Produces: Collapsible Accordion UI panel in `generate-surat.php` with all 21 controls, step buttons, live displays, URL synchronization, and auto-fit/reset functionality.

- [ ] **Step 1: Replace layout panel markup with Accordion structure**

In `modules/generate-surat.php`:
Organize `#layoutPanelBody` with modern styled collapsible accordions:
- **Akordion 1: Margin & Huruf Global** (Margin 4 sisi, Ukuran Font, Line Height)
- **Akordion 2: Kop & Judul Surat** (Jarak Bawah Kop, Jarak Atas Judul, Jarak ke Nomor, Jarak Bawah Nomor)
- **Akordion 3: Dasar & Memerintahkan** (Jarak Atas Dasar, Kerapatan Butir Dasar, Jarak Bawah Dasar, Spasi Memerintahkan)
- **Akordion 4: Blok Pegawai** (Jarak Sebelum KEPADA, Jarak Antar Pegawai, Kerapatan Baris Identitas)
- **Akordion 5: Blok Untuk** (Jarak Sebelum UNTUK, Kerapatan Butir Tugas, Jarak Bawah Tugas)
- **Akordion 6: Blok Tanda Tangan** (Jarak Atas TTD, Geser Horizontal TTD, Jarak Tanggal ke Jabatan, Tinggi Ruang TTD/Stempel, Jarak ke Nama/NIP, Kunci Blok TTD)

Each control follows the established UI pattern:
```html
<div class="space-y-1.5">
    <div class="flex items-center justify-between gap-2">
        <label class="text-xs font-medium text-slate-700">{Label}</label>
        <div class="flex items-center gap-1">
            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="{key}" data-dir="-1">−</button>
            <span class="text-xs font-semibold text-slate-700 tabular-nums w-12 text-center"><span data-value="{key}">{val}</span> {unit}</span>
            <button type="button" class="layout-step w-6 h-6 inline-flex items-center justify-center rounded-md border border-slate-200 text-slate-600 hover:bg-slate-100" data-target="{key}" data-dir="1">+</button>
        </div>
    </div>
    <input type="range" data-layout-key="{key}" min="{min}" max="{max}" step="{step}" value="{val}" class="w-full accent-indigo-600">
</div>
```

- [ ] **Step 2: Update JavaScript state, event listeners, and presets**

In the `<script>` block in `generate-surat.php`:
1. Update `KEYS` array to include all 21 keys.
2. Update `STEP_CONFIG` with appropriate step values and unit formatters (`px`, `mm`, `pt`, `×`).
3. Add accordion toggle click listeners (`.layout-accordion-toggle`).
4. Update `layoutAutoFit` preset click handler:
   - Margins: 10 mm
   - Font size: 10 pt
   - Line height: 1.0
   - Spasi antar-blok: dikurangi rata-rata 30% - 50% dari nilai default agar pas 1 halaman.
   - `lock_ttd`: 1
5. Update `layoutReset` to restore defaults for all 21 parameters.
6. Ensure `buildParams`, `saveDefaultBtn`, and `localStorage` seamlessly support all keys.

- [ ] **Step 3: Test syntax and load in browser**

Run: `php -l modules/generate-surat.php`  
Expected: `No syntax errors detected in modules/generate-surat.php`

- [ ] **Step 4: Commit**

```bash
git add modules/generate-surat.php
git commit -m "feat(ui): implement collapsible accordion layout panel for preview surat tugas"
```

---

### Task 4: Integration Verification with Live Preview & Cleanup

**Files:**
- Test: Verification on live surat tugas ID via browser / cURL / CLI
- Cleanup: Remove temporary test files (`tests/test_layout_config.php`, `tests/test_layout_docx_processing.php`)

- [ ] **Step 1: Run comprehensive PHP syntax check across all touched files**

Run: `php -l modules/generate-surat.php; php -l modules/save-template-layout.php`  
Expected: No syntax errors detected.

- [ ] **Step 2: Verify with real database surat ID**

Run a verification script that fetches the latest `id` from `surat_tugas`, generates the URL with new layout parameters, and tests generating DOCX and PDF without errors.

- [ ] **Step 3: Remove test files if no longer needed or keep in `tests/`**

```bash
git add -A
git commit -m "test: verify layout per block integration"
```
