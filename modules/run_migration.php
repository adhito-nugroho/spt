<?php
/**
 * Migration Runner: Add urutan column to pegawai_tugas
 * Access this file via browser to run the migration
 */

require_once __DIR__ . '/../config/database.php';

// Set header for plain text output
header('Content-Type: text/plain; charset=utf-8');

echo "=== Migration: Add urutan column to pegawai_tugas ===\n\n";

try {
    // Check if column already exists
    $stmt = $pdo->query("SHOW COLUMNS FROM pegawai_tugas LIKE 'urutan'");
    $columnExists = $stmt->fetch();
    
    if (!$columnExists) {
        echo "Step 1: Adding 'urutan' column to pegawai_tugas table...\n";
        $pdo->exec("ALTER TABLE pegawai_tugas ADD COLUMN urutan INT NOT NULL DEFAULT 0 AFTER nip");
        echo "✓ Column 'urutan' successfully added\n\n";
        
        echo "Step 2: Updating existing records with sequential order...\n";
        // For each surat_tugas, set sequential order for its pegawai
        $stmt = $pdo->query("SELECT DISTINCT id_surat_tugas FROM pegawai_tugas ORDER BY id_surat_tugas");
        $suratTugasIds = $stmt->fetchAll(PDO::FETCH_COLUMN);
        
        foreach ($suratTugasIds as $idSurat) {
            $stmt = $pdo->prepare("SELECT id FROM pegawai_tugas WHERE id_surat_tugas = ? ORDER BY id");
            $stmt->execute([$idSurat]);
            $pegawaiIds = $stmt->fetchAll(PDO::FETCH_COLUMN);
            
            $urutan = 1;
            $updateStmt = $pdo->prepare("UPDATE pegawai_tugas SET urutan = ? WHERE id = ?");
            foreach ($pegawaiIds as $pegawaiId) {
                $updateStmt->execute([$urutan, $pegawaiId]);
                $urutan++;
            }
        }
        echo "✓ Updated " . count($suratTugasIds) . " surat tugas records\n\n";
        
        echo "=== Migration completed successfully! ===\n";
        echo "\nYou can now close this page.\n";
        echo "The employee order will be preserved in generated documents.\n";
    } else {
        echo "ℹ Column 'urutan' already exists in pegawai_tugas table\n";
        echo "Migration has already been run.\n";
    }
    
} catch (PDOException $e) {
    echo "❌ Error: " . $e->getMessage() . "\n";
    echo "\nPlease contact your system administrator.\n";
    exit(1);
}
