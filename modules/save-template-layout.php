<?php
header('Content-Type: application/json; charset=utf-8');

require_once __DIR__ . '/../config/database.php';

function spt_json_response($statusCode, array $payload)
{
    http_response_code($statusCode);
    echo json_encode($payload, JSON_UNESCAPED_UNICODE);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    spt_json_response(405, ['success' => false, 'message' => 'Metode tidak diizinkan.']);
}

$rawBody = file_get_contents('php://input');
$input = json_decode($rawBody, true);
if (!is_array($input)) {
    spt_json_response(400, ['success' => false, 'message' => 'Payload JSON tidak valid.']);
}

$templateId = isset($input['template_id']) ? (int)$input['template_id'] : 0;
$rawSettings = isset($input['settings']) && is_array($input['settings']) ? $input['settings'] : null;

if ($templateId <= 0) {
    spt_json_response(422, ['success' => false, 'message' => 'Template tidak valid.']);
}
if ($rawSettings === null) {
    spt_json_response(422, ['success' => false, 'message' => 'Pengaturan layout tidak ditemukan.']);
}

$ranges = [
    'margin_top' => [5, 40, 0.5],
    'margin_bottom' => [5, 40, 0.5],
    'margin_left' => [10, 40, 0.5],
    'margin_right' => [10, 40, 0.5],
    'font_size' => [9, 13, 0.5],
    'line_height' => [1.0, 2.0, 0.05],
    'kop_offset' => [-20, 60, 1],
    'section_spacing' => [0, 30, 1],
    'ttd_offset' => [-20, 100, 1],
    'ttd_x' => [-50, 50, 1],
];

$settings = [];
foreach ($ranges as $key => $range) {
    if (!isset($rawSettings[$key]) || !is_numeric($rawSettings[$key])) {
        continue;
    }

    $value = (float)$rawSettings[$key];
    $value = max($range[0], min($range[1], $value));
    $step = $range[2];
    $settings[$key] = round(round($value / $step) * $step, 2);
}

if (array_key_exists('lock_ttd', $rawSettings)) {
    $settings['lock_ttd'] = !empty($rawSettings['lock_ttd']) ? 1 : 0;
}

if (empty($settings)) {
    spt_json_response(422, ['success' => false, 'message' => 'Tidak ada pengaturan layout yang valid.']);
}

try {
    $check = $pdo->prepare('SELECT id FROM templates WHERE id = ? LIMIT 1');
    $check->execute([$templateId]);
    if (!$check->fetch()) {
        spt_json_response(404, ['success' => false, 'message' => 'Template tidak ditemukan.']);
    }

    try {
        $pdo->exec('ALTER TABLE templates ADD COLUMN layout_settings TEXT DEFAULT NULL');
    } catch (PDOException $e) {
        // Kolom sudah ada.
    }

    $stmt = $pdo->prepare('UPDATE templates SET layout_settings = ? WHERE id = ?');
    $stmt->execute([json_encode($settings, JSON_UNESCAPED_UNICODE), $templateId]);
} catch (PDOException $e) {
    spt_json_response(500, ['success' => false, 'message' => 'Gagal menyimpan pengaturan layout.']);
}

spt_json_response(200, [
    'success' => true,
    'message' => 'Pengaturan layout berhasil disimpan sebagai default template.',
    'template_id' => $templateId,
    'settings' => $settings,
]);
