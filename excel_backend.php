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
		try {
			// Crea le tabelle una per volta per gestire meglio gli errori
			$tables = [
				"CREATE TABLE IF NOT EXISTS excel_documents (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    title VARCHAR(255) NOT NULL DEFAULT 'Documento Excel',
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_updated (updated_at)
                ) ENGINE=InnoDB",

				"CREATE TABLE IF NOT EXISTS excel_sheets (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    document_id INT NOT NULL,
                    sheet_name VARCHAR(100) NOT NULL,
                    sheet_order INT NOT NULL DEFAULT 0,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    INDEX idx_document (document_id),
                    INDEX idx_order (document_id, sheet_order),
                    UNIQUE KEY unique_sheet (document_id, sheet_name)
                ) ENGINE=InnoDB",

				"CREATE TABLE IF NOT EXISTS excel_cells (
                    id INT AUTO_INCREMENT PRIMARY KEY,
                    sheet_id INT NOT NULL,
                    row_index INT NOT NULL,
                    col_index INT NOT NULL,
                    cell_value TEXT,
                    cell_type ENUM('string', 'number', 'formula', 'boolean', 'date', 'empty') DEFAULT 'string',
                    cell_format VARCHAR(50) DEFAULT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                    INDEX idx_cells_position (sheet_id, row_index, col_index),
                    UNIQUE KEY unique_cell (sheet_id, row_index, col_index)
                ) ENGINE=InnoDB"
			];

			foreach ($tables as $sql) {
				$this->pdo->exec($sql);
			}

			// Migra la struttura per il versioning
			$this->migrateToVersioning();

			// Aggiungi foreign keys dopo aver creato le tabelle
			$this->addForeignKeys();

		} catch (PDOException $e) {
			// Se l'errore è per chiave duplicata, ignora (tabelle già esistenti)
			if (strpos($e->getMessage(), 'Duplicate key name') === false) {
				throw $e;
			}
		}
	}

	private function migrateToVersioning() {
		try {
			// Controlla se le colonne di versioning esistono già
			$stmt = $this->pdo->query("DESCRIBE excel_documents");
			$columns = $stmt->fetchAll(PDO::FETCH_COLUMN);

			$hasVersion = in_array('version', $columns);
			$hasCurrent = in_array('is_current', $columns);

			// Aggiungi le colonne se non esistono
			if (!$hasVersion) {
				$this->pdo->exec("ALTER TABLE excel_documents ADD COLUMN version INT NOT NULL DEFAULT 1");
			}

			if (!$hasCurrent) {
				$this->pdo->exec("ALTER TABLE excel_documents ADD COLUMN is_current BOOLEAN DEFAULT TRUE");
			}

			// Crea gli indici se non esistono
			if ($hasVersion && $hasCurrent) {
				try {
					$this->pdo->exec("CREATE INDEX idx_current ON excel_documents (is_current, updated_at)");
				} catch (PDOException $e) {
					// Indice già esistente, ignora
				}

				try {
					$this->pdo->exec("CREATE INDEX idx_version ON excel_documents (title, version)");
				} catch (PDOException $e) {
					// Indice già esistente, ignora
				}
			}

			// Se abbiamo appena aggiunto le colonne, inizializza i dati esistenti
			if (!$hasVersion || !$hasCurrent) {
				// Marca tutti i documenti esistenti come versione 1 e correnti
				$this->pdo->exec("UPDATE excel_documents SET version = 1, is_current = TRUE WHERE version IS NULL OR is_current IS NULL");
			}

		} catch (PDOException $e) {
			// Log dell'errore ma continua
			error_log("Migration warning: " . $e->getMessage());
		}
	}

	private function addForeignKeys() {
		$foreignKeys = [
			"ALTER TABLE excel_sheets ADD CONSTRAINT fk_sheets_document 
             FOREIGN KEY (document_id) REFERENCES excel_documents(id) ON DELETE CASCADE",

			"ALTER TABLE excel_cells ADD CONSTRAINT fk_cells_sheet 
             FOREIGN KEY (sheet_id) REFERENCES excel_sheets(id) ON DELETE CASCADE"
		];

		foreach ($foreignKeys as $sql) {
			try {
				$this->pdo->exec($sql);
			} catch (PDOException $e) {
				// Ignora se la foreign key esiste già
				if (strpos($e->getMessage(), 'Duplicate foreign key constraint name') === false &&
					strpos($e->getMessage(), 'already exists') === false) {
					// Log l'errore ma continua
					error_log("Warning: " . $e->getMessage());
				}
			}
		}
	}

	public function saveDocument($documentId, $data, $sheetNames, $title = null) {
		try {
			$this->pdo->beginTransaction();

			if ($documentId) {
				// ⭐ MODIFICA: Aggiorna il documento esistente invece di creare una nuova versione
				$stmt = $this->pdo->prepare("SELECT title FROM excel_documents WHERE id = ?");
				$stmt->execute([$documentId]);
				$existingDoc = $stmt->fetch();

				if (!$existingDoc) {
					throw new Exception("Documento non trovato");
				}

				// Aggiorna solo il timestamp
				$stmt = $this->pdo->prepare("UPDATE excel_documents SET updated_at = CURRENT_TIMESTAMP WHERE id = ?");
				$stmt->execute([$documentId]);

				// Elimina i vecchi fogli e celle per questo documento
				$stmt = $this->pdo->prepare("DELETE FROM excel_sheets WHERE document_id = ?");
				$stmt->execute([$documentId]);
				// Le celle vengono eliminate automaticamente tramite CASCADE

				$newDocumentId = $documentId; // ⭐ USA LO STESSO ID

			} else {
				// Nuovo documento
				$baseTitle = $title ?: 'Documento Excel ' . date('Y-m-d H:i:s');
				$stmt = $this->pdo->prepare("INSERT INTO excel_documents (title, version, is_current) VALUES (?, 1, TRUE)");
				$stmt->execute([$baseTitle]);
				$newDocumentId = $this->pdo->lastInsertId();
			}

			// Salva i fogli nella nuova/esistente versione
			foreach ($sheetNames as $order => $sheetName) {
				$stmt = $this->pdo->prepare("INSERT INTO excel_sheets (document_id, sheet_name, sheet_order) VALUES (?, ?, ?)");
				$stmt->execute([$newDocumentId, $sheetName, $order]);
				$sheetId = $this->pdo->lastInsertId();

				// Salva le celle del foglio
				if (isset($data[$sheetName]) && is_array($data[$sheetName])) {
					$this->saveCells($sheetId, $data[$sheetName]);
				}
			}

			$this->pdo->commit();
			return ['success' => true, 'documentId' => $newDocumentId];

		} catch (Exception $e) {
			$this->pdo->rollBack();
			return ['success' => false, 'message' => $e->getMessage()];
		}
	}

	private function saveCells($sheetId, $sheetData) {
		$stmt = $this->pdo->prepare("
        INSERT INTO excel_cells (sheet_id, row_index, col_index, cell_value, cell_type) 
        VALUES (?, ?, ?, ?, ?)
    ");

		foreach ($sheetData as $rowIndex => $row) {
			if (is_array($row)) {
				// ⭐ Trova l'ultima colonna con dati in questa riga
				$lastColWithData = -1;
				foreach ($row as $colIndex => $cellValue) {
					if ($cellValue !== null && $cellValue !== '') {
						$lastColWithData = max($lastColWithData, $colIndex);
					}
				}

				// ⭐ Salva TUTTE le celle fino all'ultima con dati (incluse quelle vuote in mezzo)
				if ($lastColWithData >= 0) {
					for ($colIndex = 0; $colIndex <= $lastColWithData; $colIndex++) {
						$cellValue = isset($row[$colIndex]) ? $row[$colIndex] : '';
						$cellType = $this->detectCellType($cellValue);

						// Salva anche le celle vuote se sono prima di celle con dati
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

	public function loadDocument($documentId = null, $version = null) {
		try {
			// Se non è specificato un ID, carica l'ultimo documento corrente
			if (!$documentId) {
				// Controlla prima se la colonna is_current esiste
				$stmt = $this->pdo->query("DESCRIBE excel_documents");
				$columns = $stmt->fetchAll(PDO::FETCH_COLUMN);
				$hasCurrent = in_array('is_current', $columns);

				if ($hasCurrent) {
					$stmt = $this->pdo->query("SELECT id FROM excel_documents WHERE is_current = TRUE ORDER BY updated_at DESC LIMIT 1");
				} else {
					$stmt = $this->pdo->query("SELECT id FROM excel_documents ORDER BY updated_at DESC LIMIT 1");
				}

				$doc = $stmt->fetch();
				if (!$doc) {
					return ['success' => false, 'message' => 'Nessun documento trovato'];
				}
				$documentId = $doc['id'];
			} elseif ($version) {
				// Cerca una versione specifica per titolo
				$stmt = $this->pdo->prepare("SELECT id FROM excel_documents WHERE title = (SELECT title FROM excel_documents WHERE id = ?) AND version = ?");
				$stmt->execute([$documentId, $version]);
				$doc = $stmt->fetch();
				if ($doc) {
					$documentId = $doc['id'];
				}
			}

			// Carica informazioni documento - gestisci colonne che potrebbero non esistere
			$stmt = $this->pdo->query("DESCRIBE excel_documents");
			$columns = $stmt->fetchAll(PDO::FETCH_COLUMN);
			$hasVersion = in_array('version', $columns);
			$hasCurrent = in_array('is_current', $columns);

			$selectFields = "title";
			if ($hasVersion) $selectFields .= ", version";
			if ($hasCurrent) $selectFields .= ", is_current";

			$stmt = $this->pdo->prepare("SELECT {$selectFields} FROM excel_documents WHERE id = ?");
			$stmt->execute([$documentId]);
			$docInfo = $stmt->fetch();

			if (!$docInfo) {
				return ['success' => false, 'message' => 'Documento non trovato'];
			}

			// Carica i fogli
			$stmt = $this->pdo->prepare("SELECT id, sheet_name FROM excel_sheets WHERE document_id = ? ORDER BY sheet_order");
			$stmt->execute([$documentId]);
			$sheets = $stmt->fetchAll();

			if (empty($sheets)) {
				return ['success' => false, 'message' => 'Documento vuoto'];
			}

			$data = [];
			$sheetNames = [];
			$debugInfo = []; // ⭐ AGGIUNGI QUESTO

			foreach ($sheets as $sheet) {
				$sheetNames[] = $sheet['sheet_name'];
				$data[$sheet['sheet_name']] = $this->loadSheetData($sheet['id']);
			}

			$result = [
				'success' => true,
				'documentId' => $documentId,
				'title' => $docInfo['title'],
				'data' => $data,
				'sheetNames' => $sheetNames,
				'debug' => $debugInfo // ⭐ AGGIUNGI QUESTO
			];

			// Aggiungi campi versioning solo se esistono
			if ($hasVersion) {
				$result['version'] = $docInfo['version'] ?? 1;
			}
			if ($hasCurrent) {
				$result['isCurrent'] = $docInfo['is_current'] ?? true;
			}

			return $result;

		} catch (Exception $e) {
			return ['success' => false, 'message' => $e->getMessage()];
		}
	}

	private function loadSheetData($sheetId) {
		$stmt = $this->pdo->prepare("
        SELECT row_index, col_index, cell_value 
        FROM excel_cells 
        WHERE sheet_id = ?
        ORDER BY row_index, col_index
    ");
		$stmt->execute([$sheetId]);
		$cells = $stmt->fetchAll();

		if (empty($cells)) {
			return [];
		}

		// Trova il massimo row_index e col_index
		$maxRow = 0;
		$maxColPerRow = [];

		foreach ($cells as $cell) {
			$maxRow = max($maxRow, $cell['row_index']);
			if (!isset($maxColPerRow[$cell['row_index']])) {
				$maxColPerRow[$cell['row_index']] = 0;
			}
			$maxColPerRow[$cell['row_index']] = max(
				$maxColPerRow[$cell['row_index']],
				$cell['col_index']
			);
		}

		// ⭐ MODIFICA: Riempi TUTTE le celle, anche quelle vuote
		$sheetData = [];

		for ($row = 0; $row <= $maxRow; $row++) {
			$maxCol = isset($maxColPerRow[$row]) ? $maxColPerRow[$row] : 0;
			$rowData = [];

			for ($col = 0; $col <= $maxCol; $col++) {
				// Trova la cella nel database
				$found = false;
				foreach ($cells as $cell) {
					if ($cell['row_index'] == $row && $cell['col_index'] == $col) {
						$rowData[$col] = $cell['cell_value'];
						$found = true;
						break;
					}
				}

				// ⭐ Se la cella non esiste, metti stringa vuota
				if (!$found) {
					$rowData[$col] = '';
				}
			}

			$sheetData[$row] = $rowData;
		}

		return $sheetData;
	}

	public function getDocumentsList($includeVersions = false) {
		// Controlla se le colonne di versioning esistono
		$stmt = $this->pdo->query("DESCRIBE excel_documents");
		$columns = $stmt->fetchAll(PDO::FETCH_COLUMN);
		$hasVersion = in_array('version', $columns);
		$hasCurrent = in_array('is_current', $columns);

		$selectFields = "id, title, created_at, updated_at, (SELECT COUNT(*) FROM excel_sheets WHERE document_id = excel_documents.id) as sheet_count";
		if ($hasVersion) $selectFields .= ", version";
		if ($hasCurrent) $selectFields .= ", is_current";

		if ($includeVersions || !$hasCurrent) {
			// Restituisce tutte le versioni
			$orderBy = $hasVersion ? "title, version DESC" : "updated_at DESC";
			$stmt = $this->pdo->query("
                SELECT {$selectFields}
                FROM excel_documents 
                ORDER BY {$orderBy}
            ");
		} else {
			// Restituisce solo le versioni correnti
			$stmt = $this->pdo->query("
                SELECT {$selectFields}
                FROM excel_documents 
                WHERE is_current = TRUE
                ORDER BY updated_at DESC
            ");
		}

		return [
			'success' => true,
			'documents' => $stmt->fetchAll()
		];
	}

	public function getDocumentVersions($title) {
		$stmt = $this->pdo->prepare("
            SELECT id, version, is_current, created_at, updated_at,
                   (SELECT COUNT(*) FROM excel_sheets WHERE document_id = excel_documents.id) as sheet_count
            FROM excel_documents 
            WHERE title = ?
            ORDER BY version DESC
        ");
		$stmt->execute([$title]);

		return [
			'success' => true,
			'versions' => $stmt->fetchAll()
		];
	}

	public function deleteDocument($documentId, $deleteAllVersions = false) {
		try {
			$this->pdo->beginTransaction();

			if ($deleteAllVersions) {
				// Elimina tutte le versioni del documento
				$stmt = $this->pdo->prepare("SELECT title FROM excel_documents WHERE id = ?");
				$stmt->execute([$documentId]);
				$doc = $stmt->fetch();

				if ($doc) {
					$stmt = $this->pdo->prepare("DELETE FROM excel_documents WHERE title = ?");
					$stmt->execute([$doc['title']]);
				}
			} else {
				// Elimina solo questa versione
				$stmt = $this->pdo->prepare("DELETE FROM excel_documents WHERE id = ?");
				$stmt->execute([$documentId]);

				// Se abbiamo eliminato la versione corrente, marca la versione precedente come corrente
				$stmt = $this->pdo->prepare("
                    SELECT title FROM excel_documents 
                    WHERE id = ? AND is_current = TRUE
                ");
				$stmt->execute([$documentId]);
				$wasCurrentDoc = $stmt->fetch();

				if ($wasCurrentDoc) {
					$stmt = $this->pdo->prepare("
                        UPDATE excel_documents 
                        SET is_current = TRUE 
                        WHERE title = ? AND id = (
                            SELECT id FROM (
                                SELECT id FROM excel_documents 
                                WHERE title = ? 
                                ORDER BY version DESC 
                                LIMIT 1
                            ) as subq
                        )
                    ");
					$stmt->execute([$wasCurrentDoc['title'], $wasCurrentDoc['title']]);
				}
			}

			$this->pdo->commit();
			return [
				'success' => true,
				'message' => 'Documento eliminato con successo'
			];
		} catch (Exception $e) {
			$this->pdo->rollBack();
			return [
				'success' => false,
				'message' => $e->getMessage()
			];
		}
	}

	public function cleanOldVersions($title, $keepVersions = 10) {
		try {
			// Mantiene solo le ultime N versioni di un documento
			$stmt = $this->pdo->prepare("
                DELETE FROM excel_documents 
                WHERE title = ? 
                AND id NOT IN (
                    SELECT id FROM (
                        SELECT id FROM excel_documents 
                        WHERE title = ? 
                        ORDER BY version DESC 
                        LIMIT ?
                    ) as keep_versions
                )
            ");
			$stmt->execute([$title, $title, $keepVersions]);

			return [
				'success' => true,
				'message' => "Mantenute le ultime {$keepVersions} versioni"
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
				$version = $_GET['version'] ?? null;
				echo json_encode($db->loadDocument($documentId, $version));
				break;

			case 'list':
				$includeVersions = isset($_GET['includeVersions']) && $_GET['includeVersions'] === 'true';
				echo json_encode($db->getDocumentsList($includeVersions));
				break;

			case 'versions':
				$title = $_GET['title'] ?? '';
				if ($title) {
					echo json_encode($db->getDocumentVersions($title));
				} else {
					echo json_encode(['success' => false, 'message' => 'Titolo richiesto']);
				}
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
				$title = $input['title'] ?? null;

				echo json_encode($db->saveDocument($documentId, $data, $sheetNames, $title));
				break;

			case 'delete':
				$documentId = $input['documentId'] ?? null;
				$deleteAllVersions = $input['deleteAllVersions'] ?? false;
				if ($documentId) {
					echo json_encode($db->deleteDocument($documentId, $deleteAllVersions));
				} else {
					echo json_encode(['success' => false, 'message' => 'ID documento richiesto']);
				}
				break;

			case 'cleanup':
				$title = $input['title'] ?? '';
				$keepVersions = $input['keepVersions'] ?? 10;
				if ($title) {
					echo json_encode($db->cleanOldVersions($title, $keepVersions));
				} else {
					echo json_encode(['success' => false, 'message' => 'Titolo richiesto']);
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