/* =========================
   script.js - versione pulita
   ========================= */

/* ---------- Cattura errori visibile (utile su mobile senza console) ---------- */
window.addEventListener('error', function (e) {
	try {
		const div = document.createElement('div');
		div.style.position = 'fixed';
		div.style.top = '0';
		div.style.left = '0';
		div.style.right = '0';
		div.style.background = 'red';
		div.style.color = 'white';
		div.style.padding = '6px';
		div.style.fontSize = '14px';
		div.style.zIndex = '99999';
		div.textContent = "JS Error: " + e.message + " @ " + e.filename + ":" + e.lineno;
		document.body.appendChild(div);
	} catch (_) {
	}
});

window.addEventListener('unhandledrejection', function (e) {
	try {
		const div = document.createElement('div');
		div.style.position = 'fixed';
		div.style.top = '30px';
		div.style.left = '0';
		div.style.right = '0';
		div.style.background = 'darkred';
		div.style.color = 'white';
		div.style.padding = '6px';
		div.style.fontSize = '14px';
		div.style.zIndex = '99999';
		div.textContent = "Promise Error: " + (e.reason && e.reason.message ? e.reason.message : e.reason);
		document.body.appendChild(div);
	} catch (_) {
	}
});

/* --------------------- Variabili globali --------------------- */
let workbook = null;
let currentSheet = 0;
let data = {};              // { sheetName: string[][] }
let selectedCell = null;
let documentId = null;
let pendingAction = null;
const STARTING_EDITABLE_ROW = 1; // Cambia questo numero per impostare da quale riga iniziare (0-based)
let bypassMode = false;
let descriptionModalData = null; // Variabile per tenere traccia dei dati della modal descrizione
const AUTO_SCROLL_THRESHOLD = 50; // Configurazione: da quante righe iniziare l'auto-scroll

/* --------------------- Utility --------------------- */

// Funzione per togglere il bypass mode
function toggleBypassMode() {
	bypassMode = !bypassMode;
	const bypassBtn = document.getElementById('bypass-btn');
	const loadXLSBtn = document.getElementById('load-xls-btn');
	const fileUploadBtn = document.getElementById('file-upload');
	const createSheetBtn = document.getElementById('create-sheet-btn');
	const deleteSheetBtn = document.getElementById('delete-sheet-btn');
	const deleteDocBtn = document.getElementById('delete-doc-btn');
	const deleteRowBtn = document.getElementById('delete-row-btn');
	const sheetTab = document.getElementById('sheet-tabs');
	const cornici = document.querySelectorAll('.A1');

	if (bypassMode) {
		cornici.forEach(el => el.classList.remove('displaynone'));
		loadXLSBtn.classList.remove('displaynone');
		fileUploadBtn.classList.remove('displaynone');
		createSheetBtn.classList.remove('displaynone');
		createSheetBtn.classList.remove('displaynone');
		deleteSheetBtn.classList.remove('displaynone');
		deleteDocBtn.classList.remove('displaynone');
		deleteRowBtn.classList.remove('displaynone');
		sheetTab.classList.remove('displaynone');
		bypassBtn.textContent = '🔓';
		bypassBtn.classList.add('danger');
		bypassBtn.classList.remove('secondary');
		showStatus('⚠️ Modalità bypass attivata - Tutte le regole disabilitate', 'error');
	} else {
		cornici.forEach(el => el.classList.add('displaynone'));
		loadXLSBtn.classList.add('displaynone');
		fileUploadBtn.classList.add('displaynone');
		createSheetBtn.classList.add('displaynone');
		createSheetBtn.classList.add('displaynone');
		deleteSheetBtn.classList.add('displaynone');
		deleteDocBtn.classList.add('displaynone');
		deleteRowBtn.classList.add('displaynone');
		sheetTab.classList.add('displaynone');
		bypassBtn.textContent = '🔒';
		bypassBtn.classList.remove('danger');
		bypassBtn.classList.add('secondary');
		showStatus('✅ Modalità normale ripristinata - Regole riattivate', 'success');
	}

	// Aggiorna lo stato del pulsante elimina riga quando cambia la modalità
	updateDeleteRowButtonState();
}


function isMobile() {
	return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent)
		|| (('ontouchstart' in window) && window.innerWidth <= 1024);
}

function colName(n) {
	// 0 -> A, 25 -> Z, 26 -> AA
	let s = "";
	while (n >= 0) {
		s = String.fromCharCode((n % 26) + 65) + s;
		n = Math.floor(n / 26) - 1;
	}
	return s;
}

function showStatus(message, type) {
	const status = document.getElementById('status');
	if (!status) return;
	status.textContent = message;
	status.className = `status ${type} show`;
	setTimeout(() => status.classList.remove('show'), 3000);
}

// Funzione helper per trovare la prima riga vuota
function findFirstEmptyRow(sheetName, startFromRow) {
	const sheetData = data[sheetName] || [];

	for (let row = startFromRow; row < Math.max(sheetData.length, 100); row++) {
		// Controlla se la riga è vuota o contiene solo celle vuote/whitespace
		const rowData = sheetData[row] || [];
		const isEmpty = rowData.every(cell => !cell || String(cell).trim() === '');

		if (isEmpty) {
			return row;
		}
	}

	// Se non trova righe vuote, restituisce la prima riga dopo i dati esistenti
	return Math.max(sheetData.length, startFromRow);
}

/* --------------------- Caricamento file Excel --------------------- */
function loadFile(event) {
	const file = event.target.files[0];
	if (!file) return;

	const reader = new FileReader();
	reader.onload = function (e) {
		try {
			workbook = XLSX.read(e.target.result, {type: 'binary'});
			// Costruisci "data" da workbook
			data = {};
			const tabsContainer = document.getElementById('sheet-tabs');
			tabsContainer.innerHTML = '';

			workbook.SheetNames.forEach((sheetName, index) => {
				const ws = workbook.Sheets[sheetName];
				let arr = [];
				try {
					arr = XLSX.utils.sheet_to_json(ws, {header: 1, raw: false});
					if ((!arr || arr.length === 0) && ws['!ref']) {
						arr = extractCellValues(ws);
					}
				} catch (_) {
					arr = [['Errore nel caricamento del foglio']];
				}
				data[sheetName] = arr;

				// Tab
				const tab = document.createElement('div');
				tab.className = `sheet-tab ${index === 0 ? 'active' : ''}`;
				tab.textContent = sheetName;
				tab.addEventListener('click', () => switchSheet(index, sheetName));
				tabsContainer.appendChild(tab);
			});

			currentSheet = 0;
			displaySheet(workbook.SheetNames[0]);
			showStatus('File caricato con successo!', 'success');
			updateButtonStates();
		} catch (error) {
			showStatus('Errore nel caricamento del file: ' + error.message, 'error');
		}
	};
	reader.readAsBinaryString(file);
}

function extractCellValues(worksheet) {
	const result = [];
	const range = worksheet['!ref'];
	if (!range) return result;

	const decoded = XLSX.utils.decode_range(range);
	for (let r = decoded.s.r; r <= decoded.e.r; r++) {
		const rowData = [];
		for (let c = decoded.s.c; c <= decoded.e.c; c++) {
			const cellAddress = XLSX.utils.encode_cell({r, c});
			const cell = worksheet[cellAddress];
			let cellValue = '';
			if (cell) cellValue = cell.w || cell.v || '';
			rowData[c - decoded.s.c] = (cellValue === undefined || cellValue === null) ? '' : String(cellValue);
		}
		result[r - decoded.s.r] = rowData;
	}
	return result;
}

/* --------------------- Visualizzazione e interazione griglia --------------------- */
function switchSheet(index, sheetName) {
	currentSheet = index;
	document.querySelectorAll('.sheet-tab').forEach((tab, i) => {
		tab.classList.toggle('active', i === index);
	});
	displaySheet(sheetName);
	updateButtonStates();
}

function displaySheet(sheetName) {
	const table = document.getElementById('spreadsheet-table');
	const loading = document.getElementById('loading');
	if (!table || !loading) return;

	if (!data[sheetName]) data[sheetName] = [[]];

	const sheetData = data[sheetName];
	const maxCols = Math.max(6, ...sheetData.map(row => row.length || 0)); // MOSTRA SOLO 6 COLONNE, QUINDI FINO ALLA LETTERA F
	const maxRows = Math.max(64, sheetData.length + 10); // +10 righe vuote extra per comodità

	// Crea header a due righe: lettere delle colonne + intestazioni dei dati
	let html = '<thead>';

	// Prima riga dell'header: lettere delle colonne (A, B, C, ...)
	html += '<tr class="displaynone A1"><th></th>';
	for (let col = 0; col < maxCols; col++) {
		html += `<th>${colName(col)}</th>`;
	}
	html += '</tr>';

	// Seconda riga dell'header: intestazioni dei dati (usando th invece di td)
	html += '<tr><th class="row-header displaynone A1">1</th>';
	for (let col = 0; col < maxCols; col++) {
		const cellValue = (sheetData[0] && sheetData[0][col]) ? sheetData[0][col] : '';
		html += `<th class="header-data-cell" data-row="0" data-col="${col}">${escapeHtml(cellValue)}</th>`;
	}
	html += '</tr>';

	html += '</thead><tbody>';

	// Righe dati (a partire dalla riga 1, dato che la riga 0 è ora nell'header)
	for (let row = 1; row < maxRows; row++) {
		html += `<tr><th class="row-header displaynone A1">${row + 1}</th>`;
		for (let col = 0; col < maxCols; col++) {
			const cellValue = (sheetData[row] && sheetData[row][col]) ? sheetData[row][col] : '';
			html += `<td class="cell" data-row="${row}" data-col="${col}">${escapeHtml(cellValue)}</td>`;
		}
		html += '</tr>';
	}
	html += '</tbody>';

	table.innerHTML = html;
	table.style.display = 'table';
	loading.style.display = 'none';

	// Attacca listener alle celle (desktop + mobile) - incluse quelle nell'header
	attachCellListeners(table);

	// Auto-scroll al fondo se ci sono molte righe
	autoScrollToBottom(sheetName);
}

function escapeHtml(s) {
	return String(s).replace(/[&<>"']/g, m =>
		({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'}[m])
	);
}

function attachCellListeners(table) {
	// Seleziona sia le celle normali che quelle nell'header
	const cells = table.querySelectorAll('td.cell, th.header-data-cell');

	cells.forEach((cell) => {
		// Desktop: click selezione, doppio click edit
		cell.addEventListener('click', () => selectCell(cell), {passive: true});
		cell.addEventListener('dblclick', () => editCell(cell), {passive: true});

		// Mobile: touchend -> seleziona + entra in edit con piccolo delay
		cell.addEventListener('touchend', () => {
			selectCell(cell);
			setTimeout(() => editCell(cell), 0);
		}, {passive: true});
	});
}

// Modifica la funzione selectCell per aggiornare lo stato del pulsante
function selectCell(cell) {
	if (selectedCell) selectedCell.classList.remove('selected');
	selectedCell = cell;
	cell.classList.add('selected');

	// Aggiorna lo stato del pulsante elimina riga
	updateDeleteRowButtonState();
}

/* --------------------- Editing cella (mobile vs desktop) --------------------- */

// Modifica la funzione isCellEditable per rispettare il bypass
function isCellEditable(row, col, sheetName) {
	// Se è attiva la modalità bypass, TUTTE le celle sono modificabili (incluse le intestazioni)
	if (bypassMode) {
		return true; // Rimuovo la restrizione row !== 0
	}

	// Prima riga sempre non modificabile (intestazioni) SOLO in modalità normale
	if (row === 0) {
		return false;
	}

	if (col === 0) {
		return false; // Colonna A (date) non modificabile in modalità normale
	}

	// Controlla se è la prima riga vuota}
	const firstEmptyRow = findFirstEmptyRow(sheetName, STARTING_EDITABLE_ROW);

	// Se è la prima riga vuota, tutte le colonne sono modificabili
	if (row === firstEmptyRow) {
		return true;
	}

	// Se non è la prima riga vuota, controlla se è colonna C (index 2) e se è l'ultima colonna con dati
	if (col === 2) {
		const rowData = data[sheetName][row] || [];

		// Trova l'ultima colonna con dati in questa riga
		let lastColWithData = -1;
		for (let c = 0; c < rowData.length; c++) {
			if (rowData[c] && String(rowData[c]).trim() !== '') {
				lastColWithData = c;
			}
		}

		// La colonna C è modificabile solo se è l'ultima colonna con dati o se non ci sono ancora dati in questa riga
		return lastColWithData === -1 || col === lastColWithData + 1;
	}

	return false;
}

function editCell(cell) {
	const row = parseInt(cell.dataset.row, 10);
	const col = parseInt(cell.dataset.col, 10);
	const sheetName = workbook ? workbook.SheetNames[currentSheet] : Object.keys(data)[currentSheet];
	const oldValue = cell.textContent;

	// Controllo se la cella è modificabile
	if (!isCellEditable(row, col, sheetName)) {
		if (row === 0 && !bypassMode) {
			console.log("Can't edit header row in normal mode!");
			showStatus('Non puoi modificare le intestazioni. Usa la Modalità Libera per modificarle', 'error');
		} else if (!bypassMode && row < STARTING_EDITABLE_ROW) {
			console.log(`Can't edit row ${row + 1}. Editing starts from row ${STARTING_EDITABLE_ROW + 1}`);
			showStatus(`Modifica consentita solo dalla riga ${STARTING_EDITABLE_ROW + 1} in poi. Usa la Modalità Libera per bypassare`, 'error');
		} else if (!bypassMode && col === 2) {
			console.log(`Can only edit column C if it's the last column or if adding new data`);
			showStatus('Puoi modificare la colonna C solo se è l\'ultima colonna. Usa la Modalità Libera per bypassare', 'error');
		} else if (!bypassMode) {
			console.log(`Can only edit the first empty row. Row ${row + 1} is not the first empty row.`);
			showStatus('Puoi modificare solo la prima riga vuota. Usa la Modalità Libera per bypassare', 'error');
		}
		return;
	}

	console.log(`Editing cell ${colName(col)}${row + 1} in sheet "${sheetName}" with old value: "${oldValue}"`);

	// Evita duplicare editor
	if (cell.dataset.editing === 'true') {
		return;
	}
	cell.dataset.editing = 'true';

	// Per la colonna B (importi), converti il valore formattato in valore raw per l'editing
	let editValue = oldValue;
	if (col === 1 && oldValue.includes('€')) {
		editValue = parseEuroAmount(oldValue).toString().replace('.', ',');
	}

	if (isMobile()) {
		// MOBILE: contenteditable
		cell.setAttribute('contenteditable', 'plaintext-only');
		cell.textContent = editValue; // Mostra valore non formattato per l'editing
		cell.focus({preventScroll: true});

		requestAnimationFrame(() => {
			try {
				const range = document.createRange();
				range.selectNodeContents(cell);
				const sel = window.getSelection();
				sel.removeAllRanges();
				sel.addRange(range);
			} catch (_) {}
		});

		const onBlur = () => {
			cell.removeAttribute('contenteditable');
			cell.dataset.editing = 'false';
			const newValue = cell.textContent;
			applyEdit(sheetName, row, col, newValue, cell);
			cell.removeEventListener('blur', onBlur);
			cell.removeEventListener('keydown', onKey);
		};

		const onKey = (e) => {
			if (e.key === 'Enter') {
				e.preventDefault();
				cell.blur();
			} else if (e.key === 'Escape') {
				cell.textContent = oldValue;
				cell.blur();
			}
		};

		cell.addEventListener('blur', onBlur);
		cell.addEventListener('keydown', onKey);
	} else {
		// DESKTOP: input
		const input = document.createElement('input');
		input.type = 'text';
		input.value = editValue; // Valore non formattato per l'editing
		input.className = 'cell-input';
		input.style.width = '100%';
		input.style.boxSizing = 'border-box';
		input.style.fontSize = '14px';

		// Se è colonna importi, mantieni l'allineamento a destra
		if (col === 1) {
			input.style.textAlign = 'right';
			input.style.fontFamily = "'Courier New', monospace";
		}

		cell.innerHTML = '';
		cell.appendChild(input);

		function save() {
			const newValue = input.value;
			cell.dataset.editing = 'false';
			applyEdit(sheetName, row, col, newValue, cell);
		}

		function cancel() {
			cell.dataset.editing = 'false';
			cell.innerHTML = escapeHtml(oldValue);
		}

		input.addEventListener('keydown', (e) => {
			if (e.key === 'Enter') {
				e.preventDefault();
				input.blur();
			} else if (e.key === 'Escape') {
				cancel();
			}
		});

		input.addEventListener('blur', save);

		requestAnimationFrame(() => {
			requestAnimationFrame(() => {
				input.focus({preventScroll: true});
				try {
					input.setSelectionRange(0, input.value.length);
				} catch (_) {}
			});
		});
	}
}

// Funzione helper per validare se un valore è un importo valido
function isValidAmount(value) {
	if (!value || value.trim() === '') return true; // Celle vuote sono valide

	// Se contiene già il simbolo euro, prova a convertirlo
	if (value.includes('€')) {
		const parsed = parseEuroAmount(value);
		return !isNaN(parsed) && parsed !== 0;
	}

	// Rimuovi spazi e converti virgole in punti per la validazione
	const cleanValue = value.trim().replace(',', '.');

	// Regex per importi validi
	const amountRegex = /^[+-]?(\d+([.,]\d+)?|[.,]\d+)$/;

	return amountRegex.test(cleanValue);
}

// funzione applyEdit per rispettare il bypass nella validazione
function applyEdit(sheetName, row, col, newValue, cell) {
	// Validazione specifica per colonna B (importi) - SOLO se non in modalità bypass
	if (!bypassMode && col === 1 && row > 0 && newValue.trim() !== '') {
		if (!isValidAmount(newValue)) {
			// Mostra errore e ripristina il valore precedente
			showStatus('Errore: Inserire un importo valido (es: 123.45, 123,45, -50)', 'error');

			// Ripristina il valore precedente nella cella
			const previousValue = data[sheetName][row] ? (data[sheetName][row][col] || '') : '';
			cell.innerHTML = escapeHtml(previousValue);

			// Aggiungi una classe CSS per evidenziare l'errore
			cell.classList.add('error');
			setTimeout(() => cell.classList.remove('error'), 2000);

			return; // Non salvare il valore non valido
		}

		// Formatta automaticamente l'importo in euro
		const formattedAmount = formatAsEuro(newValue);
		if (formattedAmount) {
			newValue = formattedAmount;
		}
	}

	// Salva il valore precedente per sapere se dobbiamo ricalcolare le medie
	const isAmountChange = (col === 1 && row > 0); // Cambio nella colonna importi
	const isDateChange = (col === 0 && row > 0); // Cambio nella colonna date

	if (!data[sheetName][row]) data[sheetName][row] = [];
	data[sheetName][row][col] = newValue;

	// Regola speciale: se sto inserendo nella colonna B (index 1) e non è la prima riga
	// SOLO se non in modalità bypass
	if (!bypassMode && col === 1 && row > 0 && newValue.trim() !== '') {
		// Aggiungi la data attuale nella colonna A (index 0) della stessa riga
		const today = new Date();
		const formattedDate = today.toLocaleDateString('it-IT', {
			day: '2-digit',
			month: '2-digit',
			year: 'numeric'
		});

		data[sheetName][row][0] = formattedDate;

		// Aggiorna anche la cella visibile della colonna A se esiste
		const dateCell = document.querySelector(`td.cell[data-row="${row}"][data-col="0"], th.header-data-cell[data-row="${row}"][data-col="0"]`);
		if (dateCell) {
			dateCell.innerHTML = escapeHtml(formattedDate);
		}

		console.log(`Data automatica inserita: ${formattedDate} nella cella A${row + 1}`);

		// Aggiorna le medie mensili dopo aver inserito un importo
		setTimeout(() => {
			updateAllMonthlyAverages(sheetName);
		}, 100);

		// Mostra modal per la descrizione
		setTimeout(() => {
			showDescriptionModal(row, sheetName);
		}, 200);
	}

	// Se cambiamo una data o un importo manualmente (anche in modalità bypass), aggiorna le medie
	if (bypassMode && (isAmountChange || isDateChange)) {
		setTimeout(() => {
			updateAllMonthlyAverages(sheetName);
		}, 100);
	}

	// Aggiorna UI (solo se non abbiamo già un input dentro)
	if (!cell.querySelector('input')) {
		cell.innerHTML = escapeHtml(newValue);

		// Se abbiamo appena formattato un importo, assicuriamoci che la UI si aggiorni
		if (!bypassMode && col === 1 && row > 0 && newValue.includes('€')) {
			// Forza un refresh della cella per mostrare la formattazione
			setTimeout(() => {
				cell.innerHTML = escapeHtml(newValue);
			}, 50);
		}
	}

	// Auto-save debounce (solo se non stiamo per mostrare la modal descrizione)
	if (bypassMode || col !== 1 || row === 0 || newValue.trim() === '') {
		clearTimeout(window.__saveTimeout);
		window.__saveTimeout = setTimeout(() => saveData(), 800);
	}
}

// Funzione da chiamare quando carichi i dati esistenti per ricalcolare tutte le medie
function recalculateAllAverages() {
	if (!workbook) return;

	workbook.SheetNames.forEach(sheetName => {
		updateAllMonthlyAverages(sheetName);
	});

	console.log('Tutte le medie mensili ricalcolate (ultimo giorno OR ultima riga)');
}

/* --------------------- Creazione / Eliminazione fogli --------------------- */
function createNewSheet() {
	if (!workbook) {
		workbook = XLSX.utils.book_new();
		data = {};
	}
	const base = 'Foglio';
	let idx = 1;
	let sheetName = `${base}${idx}`;
	while (data[sheetName]) {
		idx++;
		sheetName = `${base}${idx}`;
	}

	data[sheetName] = [[]];
	workbook.SheetNames.push(sheetName);
	workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet([[]]);

	const tabsContainer = document.getElementById('sheet-tabs');
	const tab = document.createElement('div');
	tab.className = 'sheet-tab';
	tab.textContent = sheetName;
	const newIndex = workbook.SheetNames.length - 1;
	tab.addEventListener('click', () => switchSheet(newIndex, sheetName));
	tabsContainer.appendChild(tab);

	switchSheet(newIndex, sheetName);
	updateButtonStates();
}

function deleteCurrentSheet() {
	if (!workbook || workbook.SheetNames.length <= 1) {
		showStatus('Non puoi eliminare l\'ultimo foglio', 'error');
		return;
	}
	const sheetName = workbook.SheetNames[currentSheet];
	showConfirmModal(
		'Elimina Foglio',
		`Sei sicuro di voler eliminare il foglio "${sheetName}"? Questa azione non può essere annullata.`,
		executeDeleteSheet
	);
}

function executeDeleteSheet() {
	const sheetName = workbook.SheetNames[currentSheet];
	workbook.SheetNames.splice(currentSheet, 1);
	delete workbook.Sheets[sheetName];
	delete data[sheetName];

	if (currentSheet >= workbook.SheetNames.length) currentSheet = workbook.SheetNames.length - 1;

	// Ricostruisci tabs e vista
	const tabsContainer = document.getElementById('sheet-tabs');
	tabsContainer.innerHTML = '';
	workbook.SheetNames.forEach((name, idx) => {
		const tab = document.createElement('div');
		tab.className = `sheet-tab ${idx === currentSheet ? 'active' : ''}`;
		tab.textContent = name;
		tab.addEventListener('click', () => switchSheet(idx, name));
		tabsContainer.appendChild(tab);
	});

	displaySheet(workbook.SheetNames[currentSheet]);
	updateButtonStates();
	showStatus(`Foglio "${sheetName}" eliminato`, 'success');

	setTimeout(saveData, 500);
}

function deleteDocument() {
	showConfirmModal(
		'Elimina Documento',
		'Sei sicuro di voler eliminare completamente questo documento? Tutti i dati andranno persi e questa azione non può essere annullata.',
		executeDeleteDocument
	);
}

async function executeDeleteDocument() {
	try {
		const response = await fetch('excel_backend.php', {
			method: 'POST',
			headers: {'Content-Type': 'application/json'},
			body: JSON.stringify({action: 'delete', documentId})
		});
		const result = await response.json();
		if (result.success) {
			workbook = null;
			data = {};
			documentId = null;
			selectedCell = null;
			currentSheet = 0;

			document.getElementById('sheet-tabs').innerHTML = '';
			document.getElementById('spreadsheet-table').style.display = 'none';
			const loading = document.getElementById('loading');
			loading.style.display = 'block';
			loading.textContent = 'Carica un file Excel per iniziare o crea un nuovo foglio';

			updateButtonStates();
			showStatus('Documento eliminato completamente', 'success');
		} else {
			showStatus('Errore nell\'eliminazione: ' + result.message, 'error');
		}
	} catch (error) {
		showStatus('Errore di connessione: ' + error.message, 'error');
	}
}

/* --------------------- Salvataggio / Esportazione --------------------- */
async function saveData() {
	if (!workbook || Object.keys(data).length === 0) {
		showStatus('Nessun dato da salvare', 'error');
		return;
	}

	try {
		const saveBtn = document.getElementById('save-btn');
		if (saveBtn) {
			saveBtn.textContent = '💾 Salvando...';
			saveBtn.disabled = true;
		}

		const response = await fetch('excel_backend.php', {
			method: 'POST',
			headers: {'Content-Type': 'application/json'},
			body: JSON.stringify({
				action: 'save',
				documentId,
				data,
				sheetNames: workbook.SheetNames
			})
		});

		const result = await response.json();
		if (result.success) {
			if (!documentId) documentId = result.documentId;
			showStatus('Dati salvati con successo!', 'success');
			updateButtonStates();
		} else {
			showStatus('Errore nel salvataggio: ' + result.message, 'error');
		}
	} catch (error) {
		showStatus('Errore di connessione: ' + error.message, 'error');
	} finally {
		const saveBtn = document.getElementById('save-btn');
		if (saveBtn) {
			saveBtn.textContent = '💾 Salva';
			saveBtn.disabled = false;
		}
	}
}

function exportExcel() {
	if (!workbook) {
		if (Object.keys(data).length === 0) {
			showStatus('Nessun dato da esportare', 'error');
			return;
		}
		workbook = XLSX.utils.book_new();
		Object.keys(data).forEach((sheetName) => {
			XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(data[sheetName] || [[]]), sheetName);
		});
	} else {
		// Aggiorna fogli dal "data"
		Object.keys(data).forEach(sheetName => {
			workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(data[sheetName]);
		});
	}

	const wbout = XLSX.write(workbook, {bookType: 'xlsx', type: 'binary'});
	const s2ab = (s) => {
		const buf = new ArrayBuffer(s.length);
		const view = new Uint8Array(buf);
		for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
		return buf;
	};

	const blob = new Blob([s2ab(wbout)], {type: "application/octet-stream"});
	const url = URL.createObjectURL(blob);

	const a = document.createElement('a');
	a.href = url;
	a.download = `backup_money_${new Date().toISOString().split('T')[0]}.xlsx`;
	document.body.appendChild(a);
	a.click();
	document.body.removeChild(a);
	URL.revokeObjectURL(url);

	showStatus('File esportato con successo!', 'success');
}

/* --------------------- Modal conferma --------------------- */
function showConfirmModal(title, message, onConfirm) {
	document.getElementById('modal-title').textContent = title;
	document.getElementById('modal-message').textContent = message;
	document.getElementById('confirm-modal').style.display = 'block';
	pendingAction = onConfirm;
}

function closeModal() {
	document.getElementById('confirm-modal').style.display = 'none';
	pendingAction = null;
}

function confirmAction() {
	if (pendingAction) {
		pendingAction();
		pendingAction = null;
	}
	closeModal();
}

/* --------------------- Stato pulsanti --------------------- */
function updateButtonStates() {
	const deleteSheetBtn = document.getElementById('delete-sheet-btn');
	const deleteDocBtn = document.getElementById('delete-doc-btn');

	if (!deleteSheetBtn || !deleteDocBtn) return;

	if (workbook && workbook.SheetNames.length > 1) {
		deleteSheetBtn.disabled = false;
		deleteSheetBtn.title = '';
	} else {
		deleteSheetBtn.disabled = true;
		deleteSheetBtn.title = 'Non puoi eliminare l\'ultimo foglio';
	}

	if (workbook && Object.keys(data).length > 0) {
		deleteDocBtn.disabled = false;
		deleteDocBtn.title = '';
	} else {
		deleteDocBtn.disabled = true;
		deleteDocBtn.title = 'Nessun documento da eliminare';
	}
}

/* --------------------- Eventi globali pagina --------------------- */
window.onclick = function (event) {
	const modal = document.getElementById('confirm-modal');
	if (event.target === modal) closeModal();
};

document.addEventListener('keydown', function (event) {
	if (event.key === 'Escape') closeModal();
}, {passive: true});

/* --------------------- Bootstrap all’avvio --------------------- */
window.onload = async function () {
	// Collega i pulsanti (se non già con onclick inline)
	const fileInput = document.getElementById('file-upload');
	if (fileInput) fileInput.addEventListener('change', loadFile);

	const saveBtn = document.getElementById('save-btn');
	if (saveBtn) saveBtn.addEventListener('click', saveData);

	const exportBtn = Array.from(document.querySelectorAll('button')).find(b => b.textContent && b.textContent.includes('Esporta'));
	if (exportBtn) exportBtn.addEventListener('click', exportExcel);

	const newSheetBtn = Array.from(document.querySelectorAll('button')).find(b => b.textContent && b.textContent.includes('Nuovo Foglio'));
	if (newSheetBtn) newSheetBtn.addEventListener('click', createNewSheet);

	const delSheetBtn = document.getElementById('delete-sheet-btn');
	if (delSheetBtn) delSheetBtn.addEventListener('click', deleteCurrentSheet);

	const delDocBtn = document.getElementById('delete-doc-btn');
	if (delDocBtn) delDocBtn.addEventListener('click', deleteDocument);

	const confirmBtn = document.getElementById('confirm-btn');
	if (confirmBtn) confirmBtn.addEventListener('click', confirmAction);

	// Carica dati esistenti dal server (se presenti)
	try {
		const response = await fetch('excel_backend.php?action=load');
		const result = await response.json();

		if (result.success && result.data) {
			documentId = result.documentId || null;
			data = result.data || {};

			// Ricostruisci un workbook "vuoto" ma con i sheet
			workbook = XLSX.utils.book_new();
			result.sheetNames.forEach(sheetName => {
				workbook.SheetNames.push(sheetName);
				workbook.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(data[sheetName] || [[]]);
			});

			// Tabs
			const tabsContainer = document.getElementById('sheet-tabs');
			tabsContainer.innerHTML = '';
			workbook.SheetNames.forEach((name, idx) => {
				const tab = document.createElement('div');
				tab.className = `sheet-tab ${idx === 0 ? 'active' : ''}`;
				tab.textContent = name;
				tab.addEventListener('click', () => switchSheet(idx, name));
				tabsContainer.appendChild(tab);
			});

			currentSheet = 0;
			displaySheet(workbook.SheetNames[0]);
			showStatus('Dati caricati dal server', 'success');
		} else {
			// Nessun dato: prepara workbook nuovo con un foglio
			workbook = XLSX.utils.book_new();
			const defaultName = 'Foglio1';
			data[defaultName] = [[]];
			workbook.SheetNames.push(defaultName);
			workbook.Sheets[defaultName] = XLSX.utils.aoa_to_sheet([[]]);

			const tabsContainer = document.getElementById('sheet-tabs');
			tabsContainer.innerHTML = '';
			const tab = document.createElement('div');
			tab.className = 'sheet-tab active';
			tab.textContent = defaultName;
			tab.addEventListener('click', () => switchSheet(0, defaultName));
			tabsContainer.appendChild(tab);

			displaySheet(defaultName);
		}
	} catch (error) {
		// In caso di errore di rete, crea un workbook nuovo
		workbook = XLSX.utils.book_new();
		const defaultName = 'Foglio1';
		data[defaultName] = [[]];
		workbook.SheetNames.push(defaultName);
		workbook.Sheets[defaultName] = XLSX.utils.aoa_to_sheet([[]]);

		const tabsContainer = document.getElementById('sheet-tabs');
		tabsContainer.innerHTML = '';
		const tab = document.createElement('div');
		tab.className = 'sheet-tab active';
		tab.textContent = defaultName;
		tab.addEventListener('click', () => switchSheet(0, defaultName));
		tabsContainer.appendChild(tab);

		displaySheet(defaultName);
	}

	updateButtonStates();
	initializeDescriptionModal();
	initializeDeleteRowButton();
	formatAllSheetsAmounts();

	// Ricalcola tutte le medie dopo aver caricato i dati
	setTimeout(() => {
		recalculateAllAverages();
	}, 500);
};

// Funzione per eliminare la riga della cella selezionata
function deleteSelectedRow() {
	if (!selectedCell) {
		showStatus('Nessuna cella selezionata', 'error');
		return;
	}

	if (!bypassMode) {
		showStatus('Eliminazione riga disponibile solo in Modalità Libera', 'error');
		return;
	}

	const row = parseInt(selectedCell.dataset.row, 10);
	const sheetName = workbook ? workbook.SheetNames[currentSheet] : Object.keys(data)[currentSheet];

	// Non permettere eliminazione della prima riga (intestazioni)
	if (row === 0) {
		showStatus('Non puoi eliminare la riga delle intestazioni', 'error');
		return;
	}

	// Mostra conferma prima di eliminare
	showConfirmModal(
		'Elimina Riga',
		`Sei sicuro di voler eliminare la riga ${row + 1}? Questa azione non può essere annullata.`,
		() => executeDeleteRow(row, sheetName)
	);
}

// Esegue l'eliminazione della riga
function executeDeleteRow(rowIndex, sheetName) {
	try {
		// Rimuovi la riga dai dati
		if (data[sheetName] && data[sheetName][rowIndex]) {
			data[sheetName].splice(rowIndex, 1);
		}

		// Aggiorna la visualizzazione
		displaySheet(sheetName);

		// Deseleziona la cella eliminata
		selectedCell = null;

		// IMPORTANTE: Ricalcola tutte le medie/stime/totali dopo l'eliminazione
		setTimeout(() => {
			updateAllMonthlyAverages(sheetName);
			console.log('Calcoli mensili aggiornati dopo eliminazione riga');
		}, 100);

		// Aggiorna lo stato del pulsante
		updateDeleteRowButtonState();

		showStatus(`Riga ${rowIndex + 1} eliminata con successo`, 'success');

		// Auto-save
		setTimeout(() => saveData(), 500);

	} catch (error) {
		showStatus('Errore nell\'eliminazione della riga: ' + error.message, 'error');
	}
}

// Aggiorna lo stato del pulsante elimina riga
function updateDeleteRowButtonState() {
	const deleteRowBtn = document.getElementById('delete-row-btn');
	if (!deleteRowBtn) return;

	const hasSelection = selectedCell !== null;
	const canDelete = bypassMode && hasSelection && selectedCell && parseInt(selectedCell.dataset.row, 10) > 0;

	deleteRowBtn.disabled = !canDelete;

	if (!bypassMode) {
		deleteRowBtn.title = 'Disponibile solo in Modalità Libera';
		deleteRowBtn.classList.add('disabled-bypass');
	} else if (!hasSelection) {
		deleteRowBtn.title = 'Seleziona una cella per eliminare la sua riga';
		deleteRowBtn.classList.remove('disabled-bypass');
	} else if (selectedCell && parseInt(selectedCell.dataset.row, 10) === 0) {
		deleteRowBtn.title = 'Non puoi eliminare la riga delle intestazioni';
		deleteRowBtn.classList.remove('disabled-bypass');
	} else {
		deleteRowBtn.title = 'Elimina la riga selezionata';
		deleteRowBtn.classList.remove('disabled-bypass');
	}
}

// (aggiungi alla fine della funzione window.onload esistente)
function initializeDeleteRowButton() {
	const deleteRowBtn = document.getElementById('delete-row-btn');
	if (deleteRowBtn) {
		deleteRowBtn.addEventListener('click', deleteSelectedRow);
		updateDeleteRowButtonState(); // Stato iniziale
	}
}

// Funzione per mostrare la modal descrizione
function showDescriptionModal(row, sheetName) {
	descriptionModalData = { row, sheetName };

	const modal = document.getElementById('description-modal');
	const input = document.getElementById('description-input');

	// Pulisci e focalizza l'input
	input.value = '';
	modal.style.display = 'block';

	// Focus sull'input dopo l'animazione
	setTimeout(() => {
		input.focus();
	}, 100);
}

// Funzione per chiudere la modal descrizione
function closeDescriptionModal() {
	document.getElementById('description-modal').style.display = 'none';
	descriptionModalData = null;
}

// Funzione per confermare e inserire la descrizione
function confirmDescription() {
	if (!descriptionModalData) return;

	const input = document.getElementById('description-input');
	const description = input.value.trim();

	// Se non c'è descrizione, chiedi conferma per procedere senza
	if (!description) {
		if (!confirm('Vuoi procedere senza inserire una descrizione?')) {
			return; // Rimani nella modal
		}
	}

	const { row, sheetName } = descriptionModalData;

	// Inserisci la descrizione nella colonna C (index 2)
	if (!data[sheetName][row]) data[sheetName][row] = [];
	data[sheetName][row][2] = description;

	// Aggiorna la cella visibile della colonna C se esiste
	const descriptionCell = document.querySelector(`td.cell[data-row="${row}"][data-col="2"]`);
	if (descriptionCell) {
		descriptionCell.innerHTML = escapeHtml(description);
	}

	// Mostra messaggio di conferma
	if (description) {
		showStatus(`Descrizione "${description}" aggiunta alla riga ${row + 1}`, 'success');
	} else {
		showStatus(`Importo salvato senza descrizione alla riga ${row + 1}`, 'success');
	}

	// Chiudi la modal
	closeDescriptionModal();

	// Auto-save
	setTimeout(() => saveData(), 500);
}

function setupDescriptionModalEvents() {
	const modal = document.getElementById('description-modal');
	const input = document.getElementById('description-input');
	const confirmBtn = document.getElementById('description-confirm-btn');
	const cancelBtn = document.getElementById('description-cancel-btn');

	// Chiudi modal cliccando fuori
	window.addEventListener('click', function(event) {
		if (event.target === modal) {
			closeDescriptionModal();
		}
	});

	// Gestione tasti nella modal
	document.addEventListener('keydown', function(event) {
		if (modal.style.display === 'block') {
			if (event.key === 'Enter') {
				event.preventDefault();
				confirmDescription();
			} else if (event.key === 'Escape') {
				event.preventDefault();
				closeDescriptionModal();
			}
		}
	});

	// Event listeners per i pulsanti
	if (confirmBtn) confirmBtn.addEventListener('click', confirmDescription);
	if (cancelBtn) cancelBtn.addEventListener('click', closeDescriptionModal);

	// Event listener per l'input (Enter per confermare)
	if (input) {
		input.addEventListener('keydown', function(event) {
			if (event.key === 'Enter') {
				event.preventDefault();
				confirmDescription();
			}
		});
	}
}

// Aggiungi questa chiamata nella funzione window.onload
function initializeDescriptionModal() {
	setupDescriptionModalEvents();
}

// Funzione per estrarre mese e anno da una data italiana (gg/mm/aaaa)
function getMonthYearFromDate(dateString) {
	if (!dateString || typeof dateString !== 'string') return null;

	// Gestisce formato italiano gg/mm/aaaa
	const parts = dateString.trim().split('/');
	if (parts.length !== 3) return null;

	const day = parseInt(parts[0], 10);
	const month = parseInt(parts[1], 10);
	const year = parseInt(parts[2], 10);

	// Validazione base
	if (isNaN(day) || isNaN(month) || isNaN(year) || month < 1 || month > 12) {
		return null;
	}

	// Restituisce oggetto con informazioni complete
	return {
		day: day,
		month: month,
		year: year,
		monthYear: `${month.toString().padStart(2, '0')}/${year}`
	};
}

// Funzione per ottenere l'ultimo giorno di un mese
function getLastDayOfMonth(month, year) {
	// Crea una data per il primo giorno del mese successivo, poi sottrae un giorno
	return new Date(year, month, 0).getDate();
}

// Funzione per controllare se una data è l'ultimo giorno del mese
function isLastDayOfMonth(dateString) {
	const dateInfo = getMonthYearFromDate(dateString);
	if (!dateInfo) return false;

	const lastDay = getLastDayOfMonth(dateInfo.month, dateInfo.year);
	return dateInfo.day === lastDay;
}

// Funzione per trovare l'ultima riga con dati nell'intero foglio
function findLastRowOfSheet(sheetName) {
	if (!data[sheetName]) return -1;

	const sheetData = data[sheetName];

	// Scorre dal fondo verso l'alto per trovare l'ultima riga con dati
	for (let row = sheetData.length - 1; row >= 1; row--) {
		const rowData = sheetData[row] || [];
		const dateCell = rowData[0];
		const amountCell = rowData[1];

		// Se la riga ha sia data che importo, è l'ultima riga con dati
		if (dateCell && amountCell) {
			return row;
		}
	}

	return -1;
}

// Funzione per calcolare il totale effettivo degli importi di un mese specifico
function calculateMonthlyTotal(sheetName, targetMonthYear) {
	if (!data[sheetName]) return 0;

	const sheetData = data[sheetName];
	let totalAmount = 0;

	// Scorre tutte le righe (escludendo la prima riga di intestazioni)
	for (let row = 1; row < sheetData.length; row++) {
		const rowData = sheetData[row] || [];
		const dateCell = rowData[0]; // Colonna A (data)
		const amountCell = rowData[1]; // Colonna B (importo)

		// Controlla se abbiamo sia data che importo
		if (!dateCell || !amountCell) continue;

		// Estrae mese/anno dalla data
		const dateInfo = getMonthYearFromDate(dateCell);
		if (!dateInfo || dateInfo.monthYear !== targetMonthYear) continue;

		// Converte l'importo in numero (gestisce sia formati grezzi che formattati)
		const amount = parseEuroAmount(amountCell);
		if (!isNaN(amount)) {
			totalAmount += amount;
		}
	}

	return totalAmount;
}


// Funzione helper per aggiornare una cella calcolata
function updateCalculatedCell(row, col, value, tooltip) {
	const cell = document.querySelector(`td.cell[data-row="${row}"][data-col="${col}"]`);
	if (cell) {
		cell.innerHTML = escapeHtml(value);
		cell.classList.add('calculated-cell');
		cell.title = tooltip;

		// Aggiungi classe specifica per tipo di calcolo
		if (col === 3) cell.classList.add('calculated-average');
		if (col === 4) cell.classList.add('calculated-estimated');
		if (col === 5) cell.classList.add('calculated-total');
	}
}

// Funzione per calcolare la media degli importi di un mese specifico
function calculateMonthlyAverage(sheetName, targetMonthYear) {
	if (!data[sheetName]) return 0;

	const sheetData = data[sheetName];
	let totalAmount = 0;
	let count = 0;

	// Scorre tutte le righe (escludendo la prima riga di intestazioni)
	for (let row = 1; row < sheetData.length; row++) {
		const rowData = sheetData[row] || [];
		const dateCell = rowData[0]; // Colonna A (data)
		const amountCell = rowData[1]; // Colonna B (importo)

		// Controlla se abbiamo sia data che importo
		if (!dateCell || !amountCell) continue;

		// Estrae mese/anno dalla data
		const dateInfo = getMonthYearFromDate(dateCell);
		if (!dateInfo || dateInfo.monthYear !== targetMonthYear) continue;

		// Converte l'importo in numero
		const amount = parseFloat(String(amountCell).replace(',', '.'));
		if (!isNaN(amount)) {
			totalAmount += amount;
			count++;
		}
	}

	return count > 0 ? totalAmount / count : 0;
}

// Funzione per aggiornare le medie mensili secondo le due regole
function updateAllMonthlyAverages(sheetName) {
	if (!data[sheetName]) return;

	const sheetData = data[sheetName];

	// Prima cancella tutte le celle calcolate esistenti (colonne D, E, F)
	for (let row = 1; row < sheetData.length; row++) {
		if (data[sheetName][row]) {
			data[sheetName][row][3] = ''; // Pulisce colonna D (media)
			data[sheetName][row][4] = ''; // Pulisce colonna E (stima)
			data[sheetName][row][5] = ''; // Pulisce colonna F (totale)
		}

		// Rimuovi anche dalla UI
		for (let col = 3; col <= 5; col++) {
			const cell = document.querySelector(`td.cell[data-row="${row}"][data-col="${col}"]`);
			if (cell) {
				cell.innerHTML = '';
				cell.classList.remove('calculated-cell', 'calculated-average', 'calculated-estimated', 'calculated-total');
			}
		}
	}

	// Trova l'ultima riga dell'intero foglio
	const lastRowOfSheet = findLastRowOfSheet(sheetName);

	// Raccoglie tutti i mesi e trova la riga migliore per ogni mese
	const monthCalculations = new Map(); // monthYear -> {row, reason}

	// Prima passata: identifica tutte le righe candidate per ogni mese
	for (let row = 1; row < sheetData.length; row++) {
		const rowData = sheetData[row] || [];
		const dateCell = rowData[0];
		const amountCell = rowData[1];

		if (!dateCell || !amountCell) continue;

		const dateInfo = getMonthYearFromDate(dateCell);
		if (!dateInfo) continue;

		let shouldConsider = false;
		let reason = '';
		let priority = 0;

		// PRIMO CASO: È l'ultimo giorno del mese (priorità alta)
		if (isLastDayOfMonth(dateCell)) {
			shouldConsider = true;
			reason = `Ultimo giorno del mese (${dateInfo.day})`;
			priority = 2; // Priorità alta
		}

		// SECONDO CASO: È l'ultima riga dell'intero foglio (priorità media)
		if (row === lastRowOfSheet) {
			shouldConsider = true;
			if (reason) {
				reason += ' + Ultima riga del foglio';
				priority = 3; // Priorità massima (entrambi i casi)
			} else {
				reason = 'Ultima riga del foglio';
				priority = 1; // Priorità media
			}
		}

		if (shouldConsider) {
			const existing = monthCalculations.get(dateInfo.monthYear);

			// Se non esiste ancora questo mese, o se questa riga ha priorità maggiore, o se ha stessa priorità ma riga più alta
			if (!existing || priority > existing.priority || (priority === existing.priority && row > existing.row)) {
				monthCalculations.set(dateInfo.monthYear, {
					row: row,
					reason: reason,
					priority: priority,
					dateInfo: dateInfo
				});
			}
		}
	}

	// Seconda passata: applica i calcoli solo alle righe selezionate
	monthCalculations.forEach((calc, monthYear) => {
		const { row, reason, dateInfo } = calc;

		const average = calculateMonthlyAverage(sheetName, monthYear);
		const actualTotal = calculateMonthlyTotal(sheetName, monthYear);

		if (average > 0) {
			// Calcola la stima del totale mensile (media × giorni del mese)
			const daysInMonth = getLastDayOfMonth(dateInfo.month, dateInfo.year);
			const estimatedTotal = average * daysInMonth;

			if (!data[sheetName][row]) data[sheetName][row] = [];

			// Colonna D (index 3): Media mensile
			data[sheetName][row][3] = formatAsEuro(average);

			// Colonna E (index 4): Stima totale mensile
			data[sheetName][row][4] = formatAsEuro(estimatedTotal);

			// Colonna F (index 5): Totale effettivo mensile
			data[sheetName][row][5] = formatAsEuro(actualTotal);

			// Aggiorna le celle visibili
			updateCalculatedCell(row, 3, data[sheetName][row][3], `Media mensile (${monthYear}) - ${reason}`);
			updateCalculatedCell(row, 4, data[sheetName][row][4], `Stima totale mensile (${monthYear}) - Media × ${daysInMonth} giorni`);
			updateCalculatedCell(row, 5, data[sheetName][row][5], `Totale effettivo mensile (${monthYear}) - Somma di tutti gli importi`);

			console.log(`Calcoli mensili ${monthYear} nella riga ${row + 1} (UNICA per il mese):`);
			console.log(`- Motivo: ${reason}`);
			console.log(`- Media: ${data[sheetName][row][3]}`);
			console.log(`- Stima totale: ${data[sheetName][row][4]} (${daysInMonth} giorni)`);
			console.log(`- Totale effettivo: ${data[sheetName][row][5]}`);
		}
	});
}


// Funzione per formattare un numero come importo in euro
function formatAsEuro(value) {
	if (!value || value === '' || isNaN(parseFloat(String(value).replace(',', '.')))) {
		return '';
	}

	// Converte in numero (gestisce sia punto che virgola come decimali)
	const number = parseFloat(String(value).replace(',', '.'));

	// Formatta con 2 decimali, virgola come separatore decimale e simbolo euro
	return number.toLocaleString('it-IT', {
		minimumFractionDigits: 2,
		maximumFractionDigits: 2
	}) + ' €';
}

// Funzione per convertire un importo formattato in numero per i calcoli
function parseEuroAmount(formattedValue) {
	if (!formattedValue || formattedValue === '') return 0;

	// Rimuove il simbolo euro e spazi, poi converte virgola in punto
	return parseFloat(String(formattedValue)
		.replace(' €', '')
		.replace('€', '')
		.replace(/\s/g, '')
		.replace(',', '.')
	) || 0;
}

function formatAllExistingAmounts(sheetName) {
	if (!data[sheetName]) return;

	const sheetData = data[sheetName];
	let formattedCount = 0;

	// Scorre tutte le righe (escludendo la prima riga di intestazioni)
	for (let row = 1; row < sheetData.length; row++) {
		const rowData = sheetData[row] || [];
		const amountCell = rowData[1]; // Colonna B (importo)

		// Se c'è un importo e non è già formattato
		if (amountCell && !String(amountCell).includes('€')) {
			const formatted = formatAsEuro(amountCell);
			if (formatted && formatted !== amountCell) {
				data[sheetName][row][1] = formatted;

				// Aggiorna anche la cella visibile
				const cell = document.querySelector(`td.cell[data-row="${row}"][data-col="1"], th.header-data-cell[data-row="${row}"][data-col="1"]`);
				if (cell) {
					cell.innerHTML = escapeHtml(formatted);
				}

				formattedCount++;
			}
		}
	}

	// Ricalcola le medie dopo aver formattato
	if (formattedCount > 0) {
		updateAllMonthlyAverages(sheetName);
		console.log(`Formattati ${formattedCount} importi in formato Euro`);
		showStatus(`${formattedCount} importi formattati in Euro`, 'success');
	}
}

function formatAllSheetsAmounts() {
	if (!workbook) return;

	workbook.SheetNames.forEach(sheetName => {
		formatAllExistingAmounts(sheetName);
	});

	console.log('Tutti gli importi esistenti formattati in Euro');
}

// Funzione per contare le righe con dati reali (non vuote)
function countDataRows(sheetName) {
	if (!data[sheetName]) return 0;

	const sheetData = data[sheetName];
	let dataRows = 0;

	// Conta le righe che hanno almeno una cella con dati (escludendo intestazioni)
	for (let row = 1; row < sheetData.length; row++) {
		const rowData = sheetData[row] || [];
		const hasData = rowData.some(cell => cell && String(cell).trim() !== '');

		if (hasData) {
			dataRows = row; // Salva l'indice dell'ultima riga con dati
		}
	}

	return dataRows;
}

// Funzione per scrollare automaticamente al fondo
function autoScrollToBottom(sheetName) {
	const dataRowsCount = countDataRows(sheetName);

	// Se ci sono abbastanza righe, scrolla al fondo
	if (dataRowsCount >= AUTO_SCROLL_THRESHOLD) {
		const gridContainer = document.getElementById('spreadsheet-table').closest('.grid-container');
		if (!gridContainer) return;

		// Aspetta che la tabella sia completamente renderizzata
		setTimeout(() => {
			// Calcola la posizione di scroll per essere vicini al fondo ma con un po' di spazio
			const containerHeight = gridContainer.clientHeight;
			const scrollHeight = gridContainer.scrollHeight;
			const headerHeight = 80; // Altezza approssimativa dell'header fisso

			// Scrolla verso il fondo lasciando un po' di spazio per vedere alcune righe vuote
			const targetScroll = scrollHeight - containerHeight + headerHeight + 200; // +200px per vedere righe vuote

			gridContainer.scrollTo({
				top: Math.max(0, targetScroll),
				behavior: 'smooth' // Scroll animato
			});

			console.log(`Auto-scroll attivato: ${dataRowsCount} righe rilevate (soglia: ${AUTO_SCROLL_THRESHOLD})`);
			showStatus(`📍 Auto-scroll: ${dataRowsCount} righe caricate, portato al fondo`, 'success');
		}, 300); // Aspetta che la tabella sia completamente renderizzata
	} else {
		console.log(`Auto-scroll non necessario: solo ${dataRowsCount} righe (soglia: ${AUTO_SCROLL_THRESHOLD})`);
	}
}
