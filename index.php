<?php
// Funzione per leggere la versione dal changelog
function getCurrentVersion(): string
{
	$changelogFile = 'changelog.txt';

	// Controlla se il file esiste
	if (!file_exists($changelogFile)) {
		return '1.0.0'; // Versione di fallback
	}

	// Legge il file
	$content = file_get_contents($changelogFile);
	if ($content === false) {
		return '1.0.0'; // Versione di fallback
	}

	// Cerca la prima occorrenza di "Version X.X.X"
	if (preg_match('/Version\s+([\d\.]+(?:\s+\w+)?)/i', $content, $matches)) {
		return trim($matches[1]);
	}

	return '1.0.0'; // Versione di fallback
}

$currentVersion = getCurrentVersion();
?>
<!DOCTYPE html>
<html lang="it">
<head>
	<meta charset="UTF-8">
	<meta name="viewport" content="width=device-width, initial-scale=1.0">
	<title>Money v<?php echo htmlspecialchars($currentVersion); ?></title>
	<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
	<link rel="stylesheet" href="styles.css">
</head>
<body>
<div class="header">
	<h1>Money v<?php echo htmlspecialchars($currentVersion); ?></h1>
	<div class="controls">
		<button style="display: none" onclick="document.getElementById('file-upload').click()">
			📁 Carica Excel
		</button>
		<input style="display: none" type="file" id="file-upload" class="file-input" accept=".xlsx,.xls" onchange="loadFile(event)">
		<button onclick="saveData()" id="save-btn">💾</button>
		<button onclick="exportExcel()">📤</button>
		<button onclick="toggleBypassMode()" id="bypass-btn" class="secondary">🔒</button>
		<button onclick="deleteSelectedRow()" id="delete-row-btn" class="danger" disabled>🗑️ Elimina Riga</button>
		<button style="display: none" onclick="createNewSheet()">➕ Nuovo Foglio</button>
		<button style="display: none" onclick="deleteCurrentSheet()" id="delete-sheet-btn" class="danger">🗑️ Elimina Foglio</button>
		<button style="display: none" onclick="deleteDocument()" id="delete-doc-btn" class="danger">❌ Elimina Documento</button>
	</div>
</div>

<div class="container">
	<div class="spreadsheet">
		<div class="sheet-tabs" id="sheet-tabs" style="display: none">
			<!-- I tab dei fogli verranno inseriti qui -->
		</div>
		<div class="grid-container">
			<div id="loading" class="loading">
				Carica un file Excel per iniziare o crea un nuovo foglio
			</div>
			<table id="spreadsheet-table" style="display: none;">
				<!-- La tabella verrà generata dinamicamente -->
			</table>
		</div>
	</div>
</div>

<div class="status" id="status"></div>

<!-- Modal di conferma -->
<div id="confirm-modal" class="modal">
	<div class="modal-content">
		<h3 id="modal-title">Conferma azione</h3>
		<p id="modal-message">Sei sicuro di voler procedere?</p>
		<div class="modal-actions">
			<button onclick="closeModal()" class="secondary">Annulla</button>
			<button onclick="confirmAction()" id="confirm-btn" class="danger">Conferma</button>
		</div>
	</div>
</div>

<!-- Modal descrizione  -->
<div id="description-modal" class="modal">
	<div class="modal-content">
		<h3 id="description-modal-title">💰 Aggiungi Descrizione</h3>
		<p>Hai inserito un nuovo importo. Vuoi aggiungere una descrizione?</p>
		<div class="description-input-container">
			<input
					type="text"
					id="description-input"
					placeholder="Es: Spesa supermercato, Stipendio, Bolletta luce..."
					maxlength="100"
			>
		</div>
		<div class="modal-actions">
			<button id="description-cancel-btn" class="secondary">Salta</button>
			<button id="description-confirm-btn" class="primary">Aggiungi</button>
		</div>
	</div>
</div>

<script src="script.js"></script>
</body>
</html>