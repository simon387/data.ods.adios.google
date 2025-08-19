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

/* --------------------- Utility --------------------- */
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
	const maxCols = Math.max(8, ...sheetData.map(row => row.length || 0)); // 26
	const maxRows = Math.max(64, sheetData.length);

	let html = '<thead><tr><th></th>';
	for (let col = 0; col < maxCols; col++) {
		html += `<th>${colName(col)}</th>`;
	}
	html += '</tr></thead><tbody>';

	for (let row = 0; row < maxRows; row++) {
		html += `<tr><th class="row-header">${row + 1}</th>`;
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

	// Attacca listener alle celle (desktop + mobile)
	attachCellListeners(table);
}

function escapeHtml(s) {
	return String(s).replace(/[&<>"']/g, m =>
		({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'}[m])
	);
}

function attachCellListeners(table) {
	const cells = table.querySelectorAll('td.cell');

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

function selectCell(cell) {
	if (selectedCell) selectedCell.classList.remove('selected');
	selectedCell = cell;
	cell.classList.add('selected');
}

/* --------------------- Editing cella (mobile vs desktop) --------------------- */
function editCell(cell) {
	const row = parseInt(cell.dataset.row, 10);
	const col = parseInt(cell.dataset.col, 10);
	const sheetName = workbook ? workbook.SheetNames[currentSheet] : Object.keys(data)[currentSheet];
	const oldValue = cell.textContent;

	// Evita duplicare editor
	if (cell.dataset.editing === 'true') return;
	cell.dataset.editing = 'true';

	if (isMobile()) {
		// MOBILE: contenteditable = apertura tastiera più affidabile
		cell.setAttribute('contenteditable', 'plaintext-only'); // fallback: 'true' se vecchio Safari
		cell.focus({preventScroll: true});

		// Seleziona tutto il testo al focus
		requestAnimationFrame(() => {
			try {
				const range = document.createRange();
				range.selectNodeContents(cell);
				const sel = window.getSelection();
				sel.removeAllRanges();
				sel.addRange(range);
			} catch (_) {
			}
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
		// DESKTOP: input type="text"
		const input = document.createElement('input');
		input.type = 'text';
		input.value = oldValue;
		input.className = 'cell-input';
		input.style.width = '100%';
		input.style.boxSizing = 'border-box';
		input.style.fontSize = '14px';

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

		// Focus dopo il reflow per compat mobile/desktop
		requestAnimationFrame(() => {
			requestAnimationFrame(() => {
				input.focus({preventScroll: true});
				try {
					input.setSelectionRange(0, input.value.length);
				} catch (_) {
				}
			});
		});
	}
}

function applyEdit(sheetName, row, col, newValue, cell) {
	if (!data[sheetName][row]) data[sheetName][row] = [];
	data[sheetName][row][col] = newValue;

	// Aggiorna UI (solo se non abbiamo già un input dentro)
	if (!cell.querySelector('input')) {
		cell.innerHTML = escapeHtml(newValue);
	}

	// Auto-save debounce
	clearTimeout(window.__saveTimeout);
	window.__saveTimeout = setTimeout(() => saveData(), 800);
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
	a.download = `documento_${new Date().toISOString().split('T')[0]}.xlsx`;
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
};
