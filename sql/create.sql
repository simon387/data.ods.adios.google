-- setup_database.sql

-- Crea il database
CREATE DATABASE IF NOT EXISTS excel_webapp
CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE excel_webapp;

-- Tabella principale dei documenti
CREATE TABLE IF NOT EXISTS excel_documents (
                                               id INT AUTO_INCREMENT PRIMARY KEY,
                                               title VARCHAR(255) NOT NULL DEFAULT 'Documento Excel',
                                               created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                                               updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                                               INDEX idx_updated (updated_at)
) ENGINE=InnoDB COMMENT='Documenti Excel salvati';

-- Tabella dei fogli di lavoro
CREATE TABLE IF NOT EXISTS excel_sheets (
                                            id INT AUTO_INCREMENT PRIMARY KEY,
                                            document_id INT NOT NULL,
                                            sheet_name VARCHAR(100) NOT NULL,
                                            sheet_order INT NOT NULL DEFAULT 0,
                                            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                                            FOREIGN KEY (document_id) REFERENCES excel_documents(id) ON DELETE CASCADE,
                                            UNIQUE KEY unique_sheet (document_id, sheet_name),
                                            INDEX idx_document (document_id),
                                            INDEX idx_order (document_id, sheet_order)
) ENGINE=InnoDB COMMENT='Fogli di lavoro dei documenti Excel';

-- Tabella delle celle
CREATE TABLE IF NOT EXISTS excel_cells (
                                           id INT AUTO_INCREMENT PRIMARY KEY,
                                           sheet_id INT NOT NULL,
                                           row_index INT NOT NULL,
                                           col_index INT NOT NULL,
                                           cell_value TEXT,
                                           cell_type ENUM('string', 'number', 'formula', 'boolean', 'date', 'empty') DEFAULT 'string',
                                           cell_format VARCHAR(50) DEFAULT NULL COMMENT 'Formato della cella (opzionale)',
                                           created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                                           updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT