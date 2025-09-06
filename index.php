<?php
// Funzione per leggere la versione dal changelog
function getCurrentVersion(): string
{
	$changelogFile = 'changelog.txt';
	$noVersion = 'No Version Found';

	// Controlla se il file esiste
	if (!file_exists($changelogFile)) {
		return $noVersion;
	}

	// Legge il file
	$content = file_get_contents($changelogFile);
	if ($content === false) {
		return $noVersion;
	}

	// Cerca la prima occorrenza di "Version X.X.X"
	if (preg_match('/Version\s+([\d.]+(?:\s+\w+)?)/i', $content, $matches)) {
		return trim($matches[1]);
	}

	return $noVersion;
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
	<h1 class="header-title">Money v<?php echo htmlspecialchars($currentVersion); ?></h1>
	<div class="controls">
		<button class="displaynone" id="load-xls-btn" onclick="document.getElementById('file-upload').click()" title="Carica Excel">
			📁 Carica Excel
		</button>
		<input type="file" id="file-upload" class="file-input displaynone" accept=".xlsx,.xls" onchange="loadFile(event)">
		<button onclick="saveData()" id="save-btn" title="Salva">💾</button>
		<button onclick="exportExcel()" title="Esporta">📤</button>
		<button onclick="toggleBypassMode()" id="bypass-btn" class="secondary" title="BypassMode">🔒</button>
		<button onclick="createNewSheet()" class="displaynone" id="create-sheet-btn" title="Nuovo Foglio">➕ Nuovo Foglio</button>
		<button onclick="deleteSelectedRow()" id="delete-row-btn" class="danger displaynone" disabled title="Elimina Riga">🗑️ Elimina Riga</button>
		<button onclick="deleteCurrentSheet()" id="delete-sheet-btn" class="danger displaynone" title="Elimina Foglio">🗑️ Elimina Foglio</button>
		<button onclick="deleteDocument()" id="delete-doc-btn" class="danger displaynone" title="Elimina Documento">❌ Elimina Documento</button>
	</div>
</div>

<div class="container">
	<div class="spreadsheet">
		<div class="sheet-tabs displaynone" id="sheet-tabs">
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
		<p id="modal-message">Sicuro di voler procedere?</p>
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
		<p>Nuovo importo inserito, vuoi aggiungere una descrizione?</p>
		<div class="description-input-container">
			<label for="description-input"></label>
			<input id="description-input" list="autocomplete-list" autocomplete="off" />
			<datalist id="autocomplete-list">
				<option value="Sandra">
				<option value="Mercadona">
				<option value="Gasolina">
			</datalist>
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