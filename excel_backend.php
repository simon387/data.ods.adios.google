<?php
// excel_backend.php

require_once __DIR__ . '/Config.php';
use App\Config\Config;

// Configurazione database
define('DB_HOST', Config::$db_host);
define('DB_USER', Config::$db_username);
define('DB_PASS', Config::$db_password);
define('DB_NAME', Config::$db_name);

// Headers CORS e JSON
header('Content-Type: application/json');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, PUT, DELETE');
header('Access-Control-Allow-Headers: Content-Type');

class ExcelDatabase {
	private $pdo;

	public function __construct() {
		try {
			$this->pdo = new PDO(
				"mysql:host=" . DB_HOST . ";dbname=" . DB_NAME . ";charset=utf8mb4",
				DB_USER,
				DB_PASS,
				[
					PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
					PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
					PDO::ATTR_EMULATE_PREPARES => false
				]
			);
		} catch (PDOException $e) {
			die(json_encode(['success' => false, 'message' => 'Errore connessione DB: ' . $e->getMessage()]));
		}
	}

	public function initDatabase() {
		$sql = "
        CREATE TABLE IF NOT EXISTS excel_documents (
            id INT AUTO_INCREMENT PRIMARY KEY,
            title VARCHAR(255) NOT NULL DEFAULT 'Documento Excel',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
        ) ENGINE=InnoDB;

        CREATE TABLE IF NOT EXISTS excel_sheets (
            id INT AUTO_INCREMENT PRIMARY KEY,
            document_id INT NOT NULL,
            sheet_name VARCHAR(100) NOT NULL,
            sheet_order INT NOT NULL DEFAULT 0,
            FOREIGN KEY (document_id) REFERENCES excel_documents(id) ON DELETE CASCADE,
            UNIQUE KEY unique_sheet (document_id, sheet_name)
        ) ENGINE=InnoDB;

        CREATE TABLE IF NOT EXISTS excel_cells (
            id INT AUTO_INCREMENT PRIMARY KEY,
            sheet_id INT NOT NULL,
            row_index INT NOT NULL,
            col_index INT NOT NULL,
            cell_value TEXT,
            cell_type ENUM('string', 'number', 'formula', 'boolean', 'date') DEFAULT 'string',
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            FOREIGN KEY (sheet_id) REFERENCES excel_sheets(id) ON DELETE CASCADE,
            UNIQUE KEY unique_cell (sheet_id, row_index, col_index)
        ) ENGINE=InnoDB;

        CREATE INDEX idx_cells_position ON excel_cells(sheet_id, row_index, col_index);
        ";

		$this->pdo->exec($sql);
	}

	public function saveDocument($documentId, $data, $sheetNames) {
		try {
			$this->pdo->beginTransaction();

			// Crea o aggiorna documento
			if (!$documentId) {
				$stmt = $this->pdo->prepare("INSERT INTO excel_documents (title) VALUES (?)");
				$stmt->execute(['Documento Excel ' . date('Y-m-d H:i:s')]);
				$documentId = $this->pdo->lastInsertId();
			} else {
				$stmt = $this->pdo->prepare("UPDATE excel_documents SET updated_at = NOW() WHERE id = ?");
				$stmt->execute([$documentId]);
			}

			// Elimina fogli esistenti per aggiornare
			$stmt = $this->pdo->prepare("DELETE FROM excel_sheets WHERE document_id = ?");
			$stmt->execute([$documentId]);

			// Salva i fogli
			foreach ($sheetNames as $order => $sheetName) {
				$stmt = $this->pdo->prepare("INSERT INTO excel_sheets (document_id, sheet_name, sheet_order) VALUES (?, ?, ?)");
				$stmt->execute([$documentId, $sheetName, $order]);
				$sheetId = $this->pdo->lastInsertId();

				// Salva le celle del foglio
				if (isset($data[$sheetName]) && is_array($data[$sheetName])) {
					$this->saveCells($sheetId, $data[$sheetName]);
				}
			}

			$this->pdo->commit();
			return ['success' => true, 'documentId' => $documentId];

		} catch (Exception $e) {
			$this->pdo->rollBack();
			return ['success' => false, 'message' => $e->getMessage()];
		}
	}

	private function saveCells($sheetId, $sheetData) {
		$stmt = $this->pdo->prepare("
            INSERT INTO excel_cells (sheet_id, row_index, col_index, cell_value, cell_type) 
            VALUES (?, ?, ?, ?, ?)
            ON DUPLICATE KEY UPDATE 
                cell_value = VALUES(cell_value),
                cell_type = VALUES(cell_type),
                updated_at = NOW()
        ");

		foreach ($sheetData as $rowIndex => $row) {
			if (is_array($row)) {
				foreach ($row as $colIndex => $cellValue) {
					if ($cellValue !== null && $cellValue !== '') {
						$cellType = $this->detectCellType($cellValue);
						$stmt->execute([$sheetId, $rowIndex, $colIndex, $cellValue, $cellType]);
					}
				}
			}
		}
	}

	private function detectCellType($value) {
		if (is_numeric($value)) {
			return 'number';
		} elseif (is_bool($value)) {
			return 'boolean';
		} elseif (preg_match('/^\d{4}-\d{2}-\d{2}/', $value)) {
			return 'date';
		} elseif (str_starts_with($value, '=')) {
			return 'formula';
		}
		return 'string';
	}

	public function loadDocument($documentId = null) {
		try {
			// Se non è specificato un ID, carica l'ultimo documento
			if (!$documentId) {
				$stmt = $this->pdo->query("SELECT id FROM excel_documents ORDER BY updated_at DESC LIMIT 1");
				$doc = $stmt->fetch();
				if (!$doc) {
					return ['success' => false, 'message' => 'Nessun documento trovato'];
				}
				$documentId = $doc['id'];
			}

			// Carica i fogli
			$stmt = $this->pdo->prepare("SELECT id, sheet_name FROM excel_sheets WHERE document_id = ? ORDER BY sheet_order");
			$stmt->execute([$documentId]);
			$sheets = $stmt->fetchAll();

			if (empty($sheets)) {
				return ['success' => false, 'message' => 'Documento non trovato'];
			}

			$data = [];
			$sheetNames = [];

			foreach ($sheets as $sheet) {
				$sheetNames[] = $sheet['sheet_name'];
				$data[$sheet['sheet_name']] = $this->loadSheetData($sheet['id']);
			}

			return [
				'success' => true,
				'documentId' => $documentId,
				'data' => $data,
				'sheetNames' => $sheetNames
			];

		} catch (Exception $e) {
			return ['success' => false, 'message' => $e->getMessage()];
		}
	}

	private function loadSheetData($sheetId) {
		$stmt = $this->pdo->prepare("
            SELECT row_index, col_index, cell_value 
            FROM excel_cells 
            WHERE sheet_id = ? AND cell_value IS NOT NULL AND cell_value != ''
            ORDER BY row_index, col_index
        ");
		$stmt->execute([$sheetId]);
		$cells = $stmt->fetchAll();

		$sheetData = [];
		foreach ($cells as $cell) {
			$sheetData[$cell['row_index']][$cell['col_index']] = $cell['cell_value'];
		}

		// Converti in array indicizzato per compatibilità con SheetJS
		$maxRow = empty($sheetData) ? 0 : max(array_keys($sheetData));
		$result = [];

		for ($row = 0; $row <= $maxRow; $row++) {
			$rowData = [];
			if (isset($sheetData[$row])) {
				$maxCol = max(array_keys($sheetData[$row]));
				for ($col = 0; $col <= $maxCol; $col++) {
					$rowData[$col] = isset($sheetData[$row][$col]) ? $sheetData[$row][$col] : '';
				}
			}
			$result[$row] = $rowData;
		}

		return $result;
	}

	public function getDocumentsList() {
		$stmt = $this->pdo->query("
            SELECT id, title, created_at, updated_at,
                   (SELECT COUNT(*) FROM excel_sheets WHERE document_id = excel_documents.id) as sheet_count
            FROM excel_documents 
            ORDER BY updated_at DESC
        ");

		return [
			'success' => true,
			'documents' => $stmt->fetchAll()
		];
	}

	public function deleteDocument($documentId) {
		try {
			$stmt = $this->pdo->prepare("DELETE FROM excel_documents WHERE id = ?");
			$stmt->execute([$documentId]);

			return [
				'success' => true,
				'message' => 'Documento eliminato con successo'
			];
		} catch (Exception $e) {
			return [
				'success' => false,
				'message' => $e->getMessage()
			];
		}
	}
}

// Main execution
try {
	$db = new ExcelDatabase();
	$db->initDatabase();

	$method = $_SERVER['REQUEST_METHOD'];
	$input = json_decode(file_get_contents('php://input'), true);

	if ($method === 'GET') {
		$action = $_GET['action'] ?? '';

		switch ($action) {
			case 'load':
				$documentId = $_GET['documentId'] ?? null;
				echo json_encode($db->loadDocument($documentId));
				break;

			case 'list':
				echo json_encode($db->getDocumentsList());
				break;

			default:
				echo json_encode(['success' => false, 'message' => 'Azione non riconosciuta']);
		}

	} elseif ($method === 'POST') {
		$action = $input['action'] ?? '';

		switch ($action) {
			case 'save':
				$documentId = $input['documentId'] ?? null;
				$data = $input['data'] ?? [];
				$sheetNames = $input['sheetNames'] ?? [];

				echo json_encode($db->saveDocument($documentId, $data, $sheetNames));
				break;

			case 'delete':
				$documentId = $input['documentId'] ?? null;
				if ($documentId) {
					echo json_encode($db->deleteDocument($documentId));
				} else {
					echo json_encode(['success' => false, 'message' => 'ID documento richiesto']);
				}
				break;

			default:
				echo json_encode(['success' => false, 'message' => 'Azione non riconosciuta']);
		}
	}

} catch (Exception $e) {
	echo json_encode([
		'success' => false,
		'message' => 'Errore server: ' . $e->getMessage()
	]);
}
