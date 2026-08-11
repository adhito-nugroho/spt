-- Add urutan column to pegawai_tugas table
-- This will preserve the order of employees as they were input

-- Add the urutan column
ALTER TABLE pegawai_tugas ADD COLUMN IF NOT EXISTS urutan INT NOT NULL DEFAULT 0 AFTER nip;

-- Update existing records with sequential order
SET @row_num = 0;
UPDATE pegawai_tugas 
SET urutan = (@row_num := @row_num + 1)
ORDER BY id_surat_tugas, id;
