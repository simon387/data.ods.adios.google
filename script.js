// Variabili globali
let workbook = null;
let currentSheet = 0;
let data = {};
let selectedCell = null;
let documentId = null;
let pendingAction = null;

// Rileva se siamo su un dispositivo mobile
function isMobile() {
	return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent) ||
		(window.innerWidth <= 768 && 'ontouchstart' in window);
}

// Carica file Excel
function loadFile(event) {
	const file = event.target.files[0];
	if (!file) return;

	const reader = new FileReader();
	reader.onload = function (e) {
		try {
			workbook = XLSX.read(e.target.result, {type: 'binary'});
			processWorkbook();
			showStatus('File caricato con successo!', 'success');
			updateButtonStates();
		} catch (error) {
			showStatus('Errore nel caricamento del file: ' + error.message, 'error');
		}
	};
	reader.readAsBinaryString(file);
}

// Processa il workbook Excel
function processWorkbook() {
	data = {};
	const tabsContainer = document.getElementById('sheet-tabs');
	tabsContainer.innerHTML = '';

	workbook.SheetNames.forEach((sheetName, index) => {
		const worksheet = workbook.Sheets[sheetName];

		// Controlla se è una tabella pivot
		const isPivot = detectPivotTable(worksheet);

		try {
			data[sheetName] = XLSX.utils.sheet_to_json(worksheet, {header: 1, raw: false});

			// Se la conversione restituisce dati vuoti ma il foglio non è vuoto
			if (data[sheetName].length === 0 && worksheet['!ref']) {
				// Prova a estrarre i valori delle celle direttamente
				data[sheetName] = extractCellValues(worksheet);
			}
		} catch (error) {
			console.warn(`Errore nel processare il foglio ${sheetName}:`, error);
			data[sheetName] = [['Errore nel caricamento del foglio']];
		}

		// Crea tab per il foglio
		const tab = document.createElement('div');
		tab.className = `sheet-tab ${index === 0 ? 'active' : ''}`;
		tab.textContent = sheetName + (isPivot ? ' (Pivot)' : '');
		tab.onclick = () => switchSheet(index, sheetName);

		// Aggiungi indicatore visivo per le pivot
		if (isPivot) {
			tab.style.fontStyle = 'italic';
			tab.style.color = '#666';
			tab.title = 'Questo foglio contiene una tabella pivot';
		}

		tabsContainer.appendChild(tab);
	});

	currentSheet = 0;
	displaySheet(workbook.SheetNames[0]);
}

// Rileva se un foglio contiene una tabella pivot
function detectPivotTable(worksheet) {
	// Le pivot table hanno spesso proprietà specifiche
	if (worksheet['!pivots'] || worksheet['!tables']) {
		return true;
	}

	// Controlla le celle per indicatori di pivot table
	const range = worksheet['!ref'];
	if (!range) return false;

	const decoded = XLSX.utils.decode_range(range);
	for (let row = decoded.s.r; row <= decoded.e.r; row++) {
		for (let col = decoded.s.c; col <= decoded.e.c; col++) {
			const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
			const cell = worksheet[cellAddress];
			if (cell && cell.f && cell.f.includes('PIVOT')) {
				return true;
			}
		}
	}

	return false;
}

// Estrae i valori delle celle direttamente
function extractCellValues(worksheet) {
	const result = [];
	const range = worksheet['!ref'];
	if (!range) return result;

	const decoded = XLSX.utils.decode_range(range);

	for (let row = decoded.s.r; row <= decoded.e.r; row++) {
		const rowData = [];
		for (let col = decoded.s.c; col <= decoded.e.c; col++) {
			const cellAddress = XLSX.utils.encode_cell({r: row, c: col});
			const cell = worksheet[cellAddress];

			let cellValue = '';
			if (cell) {
				// Prova prima il valore formattato, poi quello raw
				cellValue = cell.w || cell.v || '';
			}

			rowData[col] = cellValue;
		}
		result[row] = rowData;
	}

	return result;
}

// Cambia foglio
function switchSheet(index, sheetName) {
	currentSheet = index;

	// Aggiorna tab attivi
	document.querySelectorAll('.sheet-tab').forEach((tab, i) => {
		tab.classList.toggle('active', i === index);
	});

	displaySheet(sheetName);
	updateButtonStates();
}

// Visualizza il foglio corrente
function displaySheet(sheetName) {
	const table = document.getElementById('spreadsheet-table');
	const loading = document.getElementById('loading');

	if (!data[sheetName]) {
		data[sheetName] = [[]];
	}

	const sheetData = data[sheetName];
	const maxCols = Math.max(26, Math.max(...sheetData.map(row => row.length)));
	const maxRows = Math.max(20, sheetData.length);

	let html = '<thead><tr><th></th>';

	// Header colonne (A, B, C, ...)
	for (let col = 0; col < maxCols; col++) {
		html += `<th>${String.fromCharCode(65 + col)}</th>`;
	}
	html += '</tr></thead><tbody>';

	// Righe dati
	for (let row = 0; row < maxRows; row++) {
		html += `<tr><th class="row-header">${row + 1}</th>`;

		for (let col = 0; col < maxCols; col++) {
			const cellValue = sheetData[row] && sheetData[row][col] ? sheetData[row][col] : '';
			// Su mobile usa solo onclick, su desktop mantieni ondblclick
			const clickHandler = isMobile() ?
				`onclick="selectCell(this)" ontouchstart="this.click()"` :
				`onclick="selectCell(this)" ondblclick="editCell(this)"`;
			html += `<td class="cell" data-row="${row}" data-col="${col}" ${clickHandler}>${cellValue}</td>`;
		}
		html += '</tr>';
	}
	html += '</tbody>';

	table.innerHTML = html;
	table.style.display = 'table';
	loading.style.display = 'none';
}

// Seleziona cella
function selectCell(cell) {
	if (selectedCell) {
		selectedCell.classList.remove('selected');
	}
	selectedCell = cell;
	cell.classList.add('selected');

	// Su mobile, entra automaticamente in modalità modifica
	if (isMobile()) {
		setTimeout(() => {
			editCell(cell);
		}, 100);
	}
}

// Modifica cella
function editCell(cell) {
	const currentValue = cell.textContent;
	const input = document.createElement('input');
	input.type = 'text';
	input.value = currentValue;

	// Aggiungi attributi specifici per mobile
	if (isMobile()) {
		input.setAttribute('inputmode', 'text');
		input.setAttribute('enterkeyhint', 'done');
		input.style.fontSize = '16px'; // Previene lo zoom su iOS
		input.style.padding = '8px';
		input.style.border = '2px solid #007bff';
		input.style.borderRadius = '4px';
		input.style.width = '100%';
		input.style.minHeight = '40px';
	}

	input.onblur = function () {
		finishEdit(cell, input.value);
	};

	input.onkeydown = function (e) {
		if (e.key === 'Enter') {
			e.preventDefault();
			input.blur(); // Su mobile, chiudi la tastiera
			finishEdit(cell, input.value);
		} else if (e.key === 'Escape') {
			finishEdit(cell, currentValue);
		}
	};

	// Aggiungi evento per il tasto "Done" su mobile
	input.addEventListener('input', function(e) {
		if (e.inputType === 'insertCompositionText' || e.inputType === 'insertText') {
			// Aggiorna in tempo reale se necessario
		}
	});

	cell.innerHTML = '';
	cell.appendChild(input);

	// Focus e selezione con un piccolo delay per mobile
	if (isMobile()) {
		setTimeout(() => {
			input.focus();
			input.setSelectionRange(0, input.value.length);
			// Forza l'apertura della tastiera su Android
			input.click();
		}, 150);
	} else {
		input.focus();
		input.select();
	}
}

// Completa modifica
function finishEdit(cell, newValue) {
	const row = parseInt(cell.dataset.row);
	const col = parseInt(cell.dataset.col);
	const sheetName = workbook.SheetNames[currentSheet];

	// Aggiorna i dati
	if (!data[sheetName][row]) {
		data[sheetName][row] = [];
	}
	data[sheetName][row][col] = newValue;

	cell.innerHTML = newValue;

	// Auto-save dopo 1 secondo
	clearTimeout(window.saveTimeout);
	window.saveTimeout = setTimeout(() => {
		saveData();
	}, 1000);
}

// Crea nuovo foglio
function createNewSheet() {
	if (!workbook) {
		workbook = XLSX.utils.book_new();
		data = {};
	}

	const sheetName = `Foglio${Object.keys(data).length + 1}`;
	data[sheetName] = [[]];
	workbook.SheetNames.push(sheetName);
	workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([[]]);

	// Aggiorna UI
	const tabsContainer = document.getElementById('sheet-tabs');
	const tab = document.createElement('div');
	tab.className = 'sheet-tab';
	tab.textContent = sheetName;
	tab.onclick = () => switchSheet(workbook.SheetNames.length - 1, sheetName);
	tabsContainer.appendChild(tab);

	switchSheet(workbook.SheetNames.length - 1, sheetName);
	updateButtonStates();
}

// Elimina foglio corrente
function deleteCurrentSheet() {
	if (!workbook || workbook.SheetNames.length <= 1) {
		showStatus('Non puoi eliminare l\'ultimo foglio', 'error');
		return;
	}

	const sheetName = workbook.SheetNames[currentSheet];
	showConfirmModal(
		'Elimina Foglio',
		`Sei sicuro di voler eliminare il foglio "${sheetName}"? Questa azione non può essere annullata.`,
		() => executeDeleteSheet()
	);
}

// Esegue l'eliminazione del foglio
function executeDeleteSheet() {
	const sheetName = workbook.SheetNames[currentSheet];

	// Rimuovi dal workbook
	workbook.SheetNames.splice(currentSheet, 1);
	delete workbook.Sheets[sheetName];
	delete data[sheetName];

	// Aggiusta l'indice del foglio corrente
	if (currentSheet >= workbook.SheetNames.length) {
		currentSheet = workbook.SheetNames.length - 1;
	}

	// Ricostruisci i tab
	processWorkbook();
	updateButtonStates();
	showStatus(`Foglio "${sheetName}" eliminato`, 'success');

	// Auto-save
	setTimeout(() => saveData(), 500);
}

// Elimina intero documento
function deleteDocument() {
	showConfirmModal(
		'Elimina Documento',
		'Sei sicuro di voler eliminare completamente questo documento? Tutti i dati andranno persi e questa azione non può essere annullata.',
		() => executeDeleteDocument()
	);
}

// Esegue l'eliminazione del documento
async function executeDeleteDocument() {
	try {
		const response = await fetch('excel_backend.php', {
			method: 'POST',
			headers: {
				'Content-Type': 'application/json'
			},
			body: JSON.stringify({
				action: 'delete',
				documentId: documentId
			})
		});

		const result = await response.json();

		if (result.success) {
			// Reset dell'applicazione
			workbook = null;
			data = {};
			documentId = null;
			selectedCell = null;
			currentSheet = 0;

			// Reset UI
			document.getElementById('sheet-tabs').innerHTML = '';
			document.getElementById('spreadsheet-table').style.display = 'none';
			document.getElementById('loading').style.display = 'block';
			document.getElementById('loading').textContent = 'Carica un file Excel per iniziare o crea un nuovo foglio';

			updateButtonStates();
			showStatus('Documento eliminato completamente', 'success');
		} else {
			showStatus('Errore nell\'eliminazione: ' + result.message, 'error');
		}
	} catch (error) {
		showStatus('Errore di connessione: ' + error.message, 'error');
	}
}

// Aggiorna stato dei pulsanti
function updateButtonStates() {
	const deleteSheetBtn = document.getElementById('delete-sheet-btn');
	const deleteDocBtn = document.getElementById('delete-doc-btn');

	// Disabilita eliminazione foglio se ce n'è solo uno o nessuno
	if (workbook && workbook.SheetNames.length > 1) {
		deleteSheetBtn.disabled = false;
		deleteSheetBtn.title = '';
	} else {
		deleteSheetBtn.disabled = true;
		deleteSheetBtn.title = 'Non puoi eliminare l\'ultimo foglio';
	}

	// Disabilita eliminazione documento se non c'è niente da eliminare
	if (workbook && Object.keys(data).length > 0) {
		deleteDocBtn.disabled = false;
		deleteDocBtn.title = '';
	} else {
		deleteDocBtn.disabled = true;
		deleteDocBtn.title = 'Nessun documento da eliminare';
	}
}

// Mostra modal di conferma
function showConfirmModal(title, message, onConfirm) {
	document.getElementById('modal-title').textContent = title;
	document.getElementById('modal-message').textContent = message;
	document.getElementById('confirm-modal').style.display = 'block';
	pendingAction = onConfirm;
}

// Chiude modal
function closeModal() {
	document.getElementById('confirm-modal').style.display = 'none';
	pendingAction = null;
}

// Conferma azione
function confirmAction() {
	if (pendingAction) {
		pendingAction();
		pendingAction = null;
	}
	closeModal();
}

// Salva dati su server
async function saveData() {
	if (!workbook || Object.keys(data).length === 0) {
		showStatus('Nessun dato da salvare', 'error');
		return;
	}

	try {
		const saveBtn = document.getElementById('save-btn');
		saveBtn.textContent = '💾 Salvando...';
		saveBtn.disabled = true;

		const response = await fetch('excel_backend.php', {
			method: 'POST',
			headers: {
				'Content-Type': 'application/json'
			},
			body: JSON.stringify({
				action: 'save',
				documentId: documentId,
				data: data,
				sheetNames: workbook.SheetNames
			})
		});

		const result = await response.json();

		if (result.success) {
			if (!documentId) {
				documentId = result.documentId;
			}
			showStatus('Dati salvati con successo!', 'success');
			updateButtonStates();
		} else {
			showStatus('Errore nel salvataggio: ' + result.message, 'error');
		}
	} catch (error) {
		showStatus('Errore di connessione: ' + error.message, 'error');
	} finally {
		const saveBtn = document.getElementById('save-btn');
		saveBtn.textContent = '💾 Salva';
		saveBtn.disabled = false;
	}
}

// Esporta in Excel
function exportExcel() {
	if (!workbook) {
		showStatus('Nessun dato da esportare', 'error');
		return;
	}

	// Aggiorna il workbook con i dati correnti
	Object.keys(data).forEach(sheetName => {
		workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(data[sheetName]);
	});

	const wbout = XLSX.write(workbook, {bookType: 'xlsx', type: 'binary'});

	function s2ab(s) {
		const buf = new ArrayBuffer(s.length);
		const view = new Uint8Array(buf);
		for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	}

	const blob = new Blob([s2ab(wbout)], {type: "application/octet-stream"});
	const url = URL.createObjectURL(blob);

	const a = document.createElement('a');
	a.href = url;
	a.download = `documento_${new Date().toISOString().split('T')[0]}.xlsx`;
	document.body.appendChild(a);
	a.click();
	document.body.removeChild(a);
	URL.revokeObjectURL(url);

	showStatus('File esportato con successo!', 'success');
}

// Mostra messaggio di stato
function showStatus(message, type) {
	const status = document.getElementById('status');
	status.textContent = message;
	status.className = `status ${type} show`;

	setTimeout(() => {
		status.classList.remove('show');
	}, 3000);
}

// Event listeners per il modal (click fuori per chiudere)
window.onclick = function (event) {
	const modal = document.getElementById('confirm-modal');
	if (event.target === modal) {
		closeModal();
	}
};

// Event listener per ESC per chiudere il modal
document.addEventListener('keydown', function (event) {
	if (event.key === 'Escape') {
		closeModal();
	}
});

// Carica dati esistenti all'avvio
window.onload = async function () {
	try {
		const response = await fetch('excel_backend.php?action=load');
		const result = await response.json();

		if (result.success && result.data) {
			documentId = result.documentId;
			data = result.data;
			workbook = XLSX.utils.book_new();

			result.sheetNames.forEach(sheetName => {
				workbook.SheetNames.push(sheetName);
				workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(data[sheetName] || []);
			});

			processWorkbook();
			showStatus('Dati caricati dal server', 'success');
		}
	} catch (error) {
		console.log('Nessun dato esistente da caricare');
	}

	updateButtonStates();
};