USE excel_webapp;

ALTER TABLE excel_cells CONVERT TO CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
ALTER TABLE excel_sheets CONVERT TO CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
ALTER TABLE excel_documents CONVERT TO CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;



ALTER TABLE excel_documents ADD COLUMN version INT DEFAULT 1;
ALTER TABLE excel_documents ADD COLUMN is_current BOOLEAN DEFAULT TRUE;
