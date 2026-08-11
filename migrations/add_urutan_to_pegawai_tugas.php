<?php
/**
 * Migration: Add urutan column to pegawai_tugas table
 * Purpose: To maintain the order of employees as they were input
 */

require_once __DIR__ . '/../config/database.php';

try {
    // Check if column already exists
    $stmt = $pdo->query("SHOW COLUMNS FROM pegawai_tugas LIKE 'urutan'");
    $columnExists = $stmt->fetch();
    
    if (!$columnExists) {
        // Add urutan column
        $pdo->exec("ALTER TABLE pegawai_tugas ADD COLUMN urutan INT NOT NULL DEFAULT 0 AFTER nip");
        echo "✓ Column 'urutan' successfully added to pegawai_tugas table\n";
        
        // Update existing records to have sequential order based on id
        $pdo->exec("
            SET @row_number = 0;
            UPDATE pegawai_tugas 
            SET urutan = (@row_number:=@row_number + 1)
            ORDER BY id_surat_tugas, id
        ");
        echo "✓ Existing records updated with sequential order\n";
    } else {
        echo "ℹ Column 'urutan' already exists in pegawai_tugas table\n";
    }
    
    echo "\nMigration completed successfully!\n";
    
} catch (PDOException $e) {
    echo "Error: " . $e->getMessage() . "\n";
    exit(1);
}
