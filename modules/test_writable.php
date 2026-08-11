<?php
$dirs = [
    'C:\\laragon\\tmp',
    'C:\\laragon\\www\\tmp',
    sys_get_temp_dir(),
    'C:\\Windows\\Temp',
    __DIR__,
];

foreach ($dirs as $dir) {
    $exists = is_dir($dir) ? 'ADA' : 'TIDAK ADA';
    $writable = (is_dir($dir) && is_writable($dir)) ? 'WRITABLE ✓' : 'TIDAK WRITABLE ✗';
    echo "$dir → $exists, $writable<br>";
}

echo "<br>sys_get_temp_dir(): " . sys_get_temp_dir();
echo "<br>PHP running as user: " . get_current_user();
?>