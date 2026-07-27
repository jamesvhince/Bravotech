// ===== Global State =====
let workbook = null;
let currentSheetData = [];
let allColumns = [];
let selectedColumns = [];
let fileName = '';
let customColumnNames = {};
let cellHighlights = {};
let currentHighlightColor = 'none';
let fontStyle = { bold: false, italic: false };
let totalColumns = new Set();
let columnAlignments = {};
let conditionalRules = []; // NEW: [{type, value1, value2, bg, fg, columns: [colIdx...]}]
let editingRuleIndex = -1; // -1 for new, otherwise index of rule being edited
let tempPresetBg = '#ffc7ce';
let tempPresetFg = '#9c0006';
let tempCfColumns = new Set();

// ===== DOM =====
const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const fileInfo = document.getElementById('file-info');
const fileNameEl = document.getElementById('file-name');
const fileMeta = document.getElementById('file-meta');
const btnRemove = document.getElementById('btn-remove');
const sheetSelector = document.getElementById('sheet-selector');
const sheetSelect = document.getElementById('sheet-select');
const btnNext1 = document.getElementById('btn-next-1');
const btnBack2 = document.getElementById('btn-back-2');
const btnNext2 = document.getElementById('btn-next-2');
const btnBack3 = document.getElementById('btn-back-3');
const btnSelectAll = document.getElementById('btn-select-all');
const btnDeselectAll = document.getElementById('btn-deselect-all');
const columnSearch = document.getElementById('column-search');
const columnsGrid = document.getElementById('columns-grid');
const selectedCountText = document.getElementById('selected-count-text');
const columnOrderSection = document.getElementById('column-order-section');
const sortableList = document.getElementById('sortable-list');
const previewThead = document.getElementById('preview-thead');
const previewTbody = document.getElementById('preview-tbody');
const previewTable = document.getElementById('preview-table');
const previewInfoText = document.getElementById('preview-info-text');
const btnExportPdf = document.getElementById('btn-export-pdf');
const btnAllRows = document.getElementById('btn-all-rows');
const rowStart = document.getElementById('row-start');
const rowEnd = document.getElementById('row-end');
const btnClearHighlights = document.getElementById('btn-clear-highlights');
const columnFit = document.getElementById('column-fit');
const btnBold = document.getElementById('btn-bold');
const btnItalic = document.getElementById('btn-italic');
const btnTotalAll = document.getElementById('btn-total-all');
const btnTotalNone = document.getElementById('btn-total-none');
const totalCheckboxes = document.getElementById('total-checkboxes');
const fontFamily = document.getElementById('font-family');
const loadingOverlay = document.getElementById('loading-overlay');
const loadingText = document.getElementById('loading-text');
const toast = document.getElementById('toast');

// CF Modal
const cfModal = document.getElementById('cf-modal');
const cfModalTitle = document.getElementById('cf-modal-title');
const cfModalClose = document.getElementById('cf-modal-close');
const btnAddRule = document.getElementById('btn-add-rule');
const cfRulesList = document.getElementById('cf-rules-list');
const cfRuleType = document.getElementById('cf-rule-type');
const cfValueLabel = document.getElementById('cf-value-label');
const cfValue1 = document.getElementById('cf-value1');
const cfValue2 = document.getElementById('cf-value2');
const cfValueBetween = document.getElementById('cf-value-between');
const cfCustomBg = document.getElementById('cf-custom-bg');
const cfCustomFg = document.getElementById('cf-custom-fg');
const cfColumnList = document.getElementById('cf-column-list');
const cfSelectAllCols = document.getElementById('cf-select-all-cols');
const cfDeselectAllCols = document.getElementById('cf-deselect-all-cols');
const cfCancel = document.getElementById('cf-cancel');
const cfSave = document.getElementById('cf-save');

// ===== Helpers =====
function formatNumber(value) {
    if (value === null || value === undefined || value === '') return '';
    const num = Number(value);
    if (isNaN(num)) return String(value);
    if (Number.isInteger(num)) return num.toLocaleString('en-US');
    return num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function isNumeric(value) {
    if (value === null || value === undefined || value === '') return false;
    return !isNaN(Number(value)) && isFinite(Number(value));
}

function isRowBlank(row, colIndices) {
    let count = 0;
    for (let i = 0; i < colIndices.length; i++) {
        if (i === 0) continue;
        const value = row[colIndices[i]];
        if (value !== null && value !== undefined && value !== '') {
            const s = String(value).trim();
            if (s !== '' && s !== '0') count++;
        }
    }
    return count === 0;
}

function hexToRgb(hex) {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return [r, g, b];
}

function getColumnAlignment(colIdx, position) {
    if (columnAlignments[colIdx]) return columnAlignments[colIdx];
    if (position === 0) return 'center';
    const dataRows = currentSheetData.slice(1);
    let numericCount = 0, totalCount = 0;
    dataRows.forEach(row => {
        const val = row[colIdx];
        if (val !== '' && val !== null && val !== undefined) {
            totalCount++;
            if (isNumeric(val)) numericCount++;
        }
    });
    if (totalCount > 0 && (numericCount / totalCount) >= 0.7) return 'right';
    return 'center';
}

// Map "bookman" to a serif font that jsPDF can render
function getPdfFontFamily(family) {
    if (family === 'bookman') return 'times'; // fallback for PDF
    return family;
}

// Check if a cell value matches a conditional rule
function cellMatchesRule(cellValue, rule) {
    const num = Number(cellValue);
    const numRule1 = Number(rule.value1);
    const numRule2 = Number(rule.value2);
    const strCell = String(cellValue).toLowerCase();
    const strRule = String(rule.value1).toLowerCase();

    switch (rule.type) {
        case 'less':
            return !isNaN(num) && !isNaN(numRule1) && num < numRule1;
        case 'greater':
            return !isNaN(num) && !isNaN(numRule1) && num > numRule1;
        case 'equal':
            if (!isNaN(num) && !isNaN(numRule1)) return num === numRule1;
            return strCell === strRule;
        case 'notequal':
            if (!isNaN(num) && !isNaN(numRule1)) return num !== numRule1;
            return strCell !== strRule;
        case 'between':
            return !isNaN(num) && !isNaN(numRule1) && !isNaN(numRule2)
                && num >= Math.min(numRule1, numRule2) && num <= Math.max(numRule1, numRule2);
        case 'contains':
            return strCell.includes(strRule);
        default:
            return false;
    }
}

// Get conditional format for a cell (returns {bg, fg} or null)
function getConditionalFormat(cellValue, colIdx) {
    for (let i = conditionalRules.length - 1; i >= 0; i--) {
        const rule = conditionalRules[i];
        if (rule.columns.includes(colIdx) && cellMatchesRule(cellValue, rule)) {
            return { bg: rule.bg, fg: rule.fg };
        }
    }
    return null;
}

function getRuleDescription(rule) {
    const typeText = {
        'less': `< ${rule.value1}`,
        'greater': `> ${rule.value1}`,
        'equal': `= ${rule.value1}`,
        'notequal': `≠ ${rule.value1}`,
        'between': `${rule.value1}..${rule.value2}`,
        'contains': `~"${rule.value1}"`
    };
    return typeText[rule.type] || '';
}

// ===== Navigation =====
function goToStep(step) {
    document.querySelectorAll('.step-section').forEach(s => s.classList.remove('active'));
    document.getElementById(`step-${step}`).classList.add('active');
    for (let i = 1; i <= 3; i++) {
        const ind = document.getElementById(`step-ind-${i}`);
        ind.classList.remove('active', 'completed');
        if (i < step) ind.classList.add('completed');
        if (i === step) ind.classList.add('active');
    }
    document.querySelectorAll('.step-line').forEach((line, idx) => {
        line.classList.toggle('active', idx < step - 1);
    });
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function showToast(message, type = 'info') {
    toast.textContent = message;
    toast.className = `toast ${type} show`;
    setTimeout(() => toast.classList.remove('show'), 3500);
}

function showLoading(text = 'Processing...') {
    loadingText.textContent = text;
    loadingOverlay.style.display = 'flex';
}
function hideLoading() { loadingOverlay.style.display = 'none'; }

// ===== File Upload =====
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
});
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) handleFile(e.target.files[0]);
});

function handleFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls', 'csv'].includes(ext)) {
        showToast('Please upload a valid Excel or CSV file.', 'error');
        return;
    }
    showLoading('Reading file...');
    fileName = file.name;
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array', cellDates: true });
            const sheetNames = workbook.SheetNames;
            sheetSelect.innerHTML = '';
            sheetNames.forEach((name) => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });
            sheetSelector.style.display = sheetNames.length > 1 ? 'flex' : 'none';
            loadSheet(sheetNames[0]);
            const fileSizeKB = (file.size / 1024).toFixed(1);
            fileNameEl.textContent = file.name;
            fileMeta.textContent = `Size: ${fileSizeKB} KB | Rows: ${currentSheetData.length - 1} | Columns: ${allColumns.length}`;
            dropZone.style.display = 'none';
            fileInfo.style.display = 'block';
            hideLoading();
            showToast('File loaded successfully!', 'success');
        } catch (err) {
            hideLoading();
            showToast('Error reading file.', 'error');
            console.error(err);
        }
    };
    reader.readAsArrayBuffer(file);
}

function loadSheet(sheetName) {
    const sheet = workbook.Sheets[sheetName];
    currentSheetData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    if (currentSheetData.length > 0) {
        allColumns = currentSheetData[0].map((col, idx) => ({
            name: col ? String(col).trim() : `Column ${idx + 1}`,
            index: idx
        }));
    } else allColumns = [];
    selectedColumns = [];
    customColumnNames = {};
    cellHighlights = {};
    totalColumns = new Set();
    columnAlignments = {};
    conditionalRules = [];
}

sheetSelect.addEventListener('change', (e) => {
    loadSheet(e.target.value);
    fileMeta.textContent = `Rows: ${currentSheetData.length - 1} | Columns: ${allColumns.length}`;
});

btnRemove.addEventListener('click', () => {
    workbook = null;
    currentSheetData = [];
    allColumns = [];
    selectedColumns = [];
    customColumnNames = {};
    cellHighlights = {};
    totalColumns = new Set();
    columnAlignments = {};
    conditionalRules = [];
    fileName = '';
    fileInput.value = '';
    fileInfo.style.display = 'none';
    dropZone.style.display = 'block';
});

btnNext1.addEventListener('click', () => {
    if (!workbook) { showToast('Please upload a file first.', 'error'); return; }
    renderColumns();
    goToStep(2);
});

// ===== Step 2 =====
function renderColumns(filter = '') {
    columnsGrid.innerHTML = '';
    const filterLower = filter.toLowerCase();
    allColumns.forEach((col) => {
        if (filter && !col.name.toLowerCase().includes(filterLower)) return;
        const card = document.createElement('div');
        card.className = 'column-card';
        if (selectedColumns.includes(col.index)) card.classList.add('selected');
        card.innerHTML = `
            <div class="checkbox"></div>
            <span class="col-name" title="${col.name}">${col.name}</span>
            <span class="col-index">${getColLetter(col.index)}</span>
        `;
        card.addEventListener('click', () => {
            toggleColumn(col.index);
            card.classList.toggle('selected');
            updateSelectedCount();
            updateColumnOrder();
        });
        columnsGrid.appendChild(card);
    });
    updateSelectedCount();
}

function getColLetter(idx) {
    let letter = '', n = idx;
    while (n >= 0) {
        letter = String.fromCharCode((n % 26) + 65) + letter;
        n = Math.floor(n / 26) - 1;
    }
    return letter;
}

function toggleColumn(index) {
    const pos = selectedColumns.indexOf(index);
    if (pos > -1) selectedColumns.splice(pos, 1);
    else selectedColumns.push(index);
}

function updateSelectedCount() {
    selectedCountText.textContent = `${selectedColumns.length} column${selectedColumns.length !== 1 ? 's' : ''} selected`;
}

function updateColumnOrder() {
    if (selectedColumns.length === 0) { columnOrderSection.style.display = 'none'; return; }
    columnOrderSection.style.display = 'block';
    sortableList.innerHTML = '';
    selectedColumns.forEach((colIdx) => {
        const col = allColumns.find(c => c.index === colIdx);
        if (!col) return;
        const li = document.createElement('li');
        li.setAttribute('draggable', 'true');
        li.dataset.index = colIdx;
        li.innerHTML = `<i class="fas fa-grip-vertical"></i> ${col.name}`;
        li.addEventListener('dragstart', () => li.classList.add('dragging'));
        li.addEventListener('dragend', () => {
            li.classList.remove('dragging');
            const items = sortableList.querySelectorAll('li');
            selectedColumns = Array.from(items).map(item => parseInt(item.dataset.index));
        });
        li.addEventListener('dragover', (e) => {
            e.preventDefault();
            const dragging = sortableList.querySelector('.dragging');
            if (dragging && dragging !== li) {
                const rect = li.getBoundingClientRect();
                const midX = rect.left + rect.width / 2;
                if (e.clientX < midX) sortableList.insertBefore(dragging, li);
                else sortableList.insertBefore(dragging, li.nextSibling);
            }
        });
        sortableList.appendChild(li);
    });
}

columnSearch.addEventListener('input', (e) => renderColumns(e.target.value));
btnSelectAll.addEventListener('click', () => {
    selectedColumns = allColumns.map(c => c.index);
    renderColumns(columnSearch.value);
    updateColumnOrder();
});
btnDeselectAll.addEventListener('click', () => {
    selectedColumns = [];
    renderColumns(columnSearch.value);
    updateColumnOrder();
});
btnNext2.addEventListener('click', () => {
    if (selectedColumns.length === 0) { showToast('Please select at least one column.', 'error'); return; }
    renderPreview();
    goToStep(3);
});
btnBack2.addEventListener('click', () => goToStep(1));

// ===== Step 3 =====
function renderPreview() {
    const totalRows = currentSheetData.length - 1;
    rowEnd.max = totalRows;
    rowEnd.value = totalRows;
    rowStart.max = totalRows;
    rowStart.value = 1;
    autoDetectTotals();
    renderTotalCheckboxes();
    renderConditionalRules();
    buildTable();
}

function autoDetectTotals() {
    const dataRows = currentSheetData.slice(1);
    totalColumns = new Set();
    selectedColumns.forEach((colIdx, i) => {
        if (i === 0) return;
        let numericCount = 0, totalCount = 0;
        dataRows.forEach(row => {
            const val = row[colIdx];
            if (val !== '' && val !== null && val !== undefined) {
                totalCount++;
                if (isNumeric(val)) numericCount++;
            }
        });
        if (totalCount > 0 && (numericCount / totalCount) >= 0.7) totalColumns.add(colIdx);
    });
}

function renderTotalCheckboxes() {
    totalCheckboxes.innerHTML = '';
    selectedColumns.forEach((colIdx, i) => {
        const col = allColumns.find(c => c.index === colIdx);
        const displayName = customColumnNames[colIdx] !== undefined ? customColumnNames[colIdx] : col.name;
        const item = document.createElement('div');
        item.className = 'total-check-item';
        if (i === 0) {
            item.classList.add('disabled');
            item.innerHTML = `<div class="check-icon"></div><span>${displayName}</span>`;
        } else {
            if (totalColumns.has(colIdx)) item.classList.add('checked');
            item.innerHTML = `<div class="check-icon"></div><span>${displayName}</span>`;
            item.addEventListener('click', () => {
                if (totalColumns.has(colIdx)) {
                    totalColumns.delete(colIdx);
                    item.classList.remove('checked');
                } else {
                    totalColumns.add(colIdx);
                    item.classList.add('checked');
                }
                buildTable();
            });
        }
        totalCheckboxes.appendChild(item);
    });
}

// ===== Conditional Formatting UI =====
function renderConditionalRules() {
    if (conditionalRules.length === 0) {
        cfRulesList.innerHTML = '<p class="cf-empty">No rules yet. Click "Add Rule" to create conditional formatting.</p>';
        return;
    }
    cfRulesList.innerHTML = '';
    conditionalRules.forEach((rule, idx) => {
        const chip = document.createElement('div');
        chip.className = 'cf-rule-chip';
        const colNames = rule.columns.map(ci => {
            const c = allColumns.find(x => x.index === ci);
            return c ? (customColumnNames[ci] || c.name) : '?';
        }).join(', ');
        chip.innerHTML = `
            <div class="cf-swatch" style="background:${rule.bg};border:1px solid ${rule.fg}"></div>
            <span class="cf-rule-desc">${getRuleDescription(rule)}</span>
            <span>→ ${colNames.substring(0, 30)}${colNames.length > 30 ? '...' : ''}</span>
            <button title="Edit" onclick="editRule(${idx})"><i class="fas fa-edit"></i></button>
            <button title="Delete" onclick="deleteRule(${idx})"><i class="fas fa-trash"></i></button>
        `;
        cfRulesList.appendChild(chip);
    });
}

window.editRule = function(idx) {
    editingRuleIndex = idx;
    const rule = conditionalRules[idx];
    openCfModal(rule);
};

window.deleteRule = function(idx) {
    if (confirm('Delete this conditional formatting rule?')) {
        conditionalRules.splice(idx, 1);
        renderConditionalRules();
        buildTable();
        showToast('Rule deleted', 'info');
    }
};

btnAddRule.addEventListener('click', () => {
    editingRuleIndex = -1;
    openCfModal(null);
});

function openCfModal(rule) {
    if (rule) {
        cfModalTitle.textContent = 'Edit Rule';
        cfRuleType.value = rule.type;
        cfValue1.value = rule.value1;
        cfValue2.value = rule.value2 || '';
        tempPresetBg = rule.bg;
        tempPresetFg = rule.fg;
        cfCustomBg.value = rule.bg;
        cfCustomFg.value = rule.fg;
        tempCfColumns = new Set(rule.columns);
    } else {
        cfModalTitle.textContent = 'New Conditional Formatting Rule';
        cfRuleType.value = 'less';
        cfValue1.value = '';
        cfValue2.value = '';
        tempPresetBg = '#ffc7ce';
        tempPresetFg = '#9c0006';
        cfCustomBg.value = '#fef08a';
        cfCustomFg.value = '#000000';
        tempCfColumns = new Set();
    }
    updateValueLabel();
    renderCfPresets();
    renderCfColumnList();
    cfModal.style.display = 'flex';
}

function closeCfModal() {
    cfModal.style.display = 'none';
    editingRuleIndex = -1;
}

cfModalClose.addEventListener('click', closeCfModal);
cfCancel.addEventListener('click', closeCfModal);
cfModal.addEventListener('click', (e) => {
    if (e.target === cfModal) closeCfModal();
});

cfRuleType.addEventListener('change', updateValueLabel);

function updateValueLabel() {
    const type = cfRuleType.value;
    const labels = {
        'less': 'Format cells that are LESS THAN:',
        'greater': 'Format cells that are GREATER THAN:',
        'equal': 'Format cells that are EQUAL TO:',
        'notequal': 'Format cells that are NOT EQUAL TO:',
        'between': 'Format cells BETWEEN:',
        'contains': 'Format cells that CONTAIN text:'
    };
    cfValueLabel.textContent = labels[type];
    if (type === 'between') {
        cfValueBetween.style.display = 'inline';
        cfValue2.style.display = 'inline-block';
    } else {
        cfValueBetween.style.display = 'none';
        cfValue2.style.display = 'none';
    }
}

function renderCfPresets() {
    document.querySelectorAll('.cf-preset').forEach(el => {
        el.classList.remove('selected');
        const bg = el.dataset.bg;
        const fg = el.dataset.fg;
        if (bg === tempPresetBg && fg === tempPresetFg) {
            el.classList.add('selected');
        }
        if (!el.classList.contains('custom-preset')) {
            el.addEventListener('click', function handler() {
                tempPresetBg = el.dataset.bg;
                tempPresetFg = el.dataset.fg;
                renderCfPresets();
            }, { once: true });
        }
    });
}

cfCustomBg.addEventListener('input', () => {
    tempPresetBg = cfCustomBg.value;
    tempPresetFg = cfCustomFg.value;
    document.querySelectorAll('.cf-preset').forEach(el => el.classList.remove('selected'));
    document.querySelector('.cf-preset.custom-preset').classList.add('selected');
});
cfCustomFg.addEventListener('input', () => {
    tempPresetBg = cfCustomBg.value;
    tempPresetFg = cfCustomFg.value;
    document.querySelectorAll('.cf-preset').forEach(el => el.classList.remove('selected'));
    document.querySelector('.cf-preset.custom-preset').classList.add('selected');
});

function renderCfColumnList() {
    cfColumnList.innerHTML = '';
    selectedColumns.forEach(colIdx => {
        const col = allColumns.find(c => c.index === colIdx);
        const displayName = customColumnNames[colIdx] !== undefined ? customColumnNames[colIdx] : col.name;
        const item = document.createElement('div');
        item.className = 'cf-col-item';
        if (tempCfColumns.has(colIdx)) item.classList.add('selected');
        item.innerHTML = `<div class="check-mini"></div><span>${displayName}</span>`;
        item.addEventListener('click', () => {
            if (tempCfColumns.has(colIdx)) {
                tempCfColumns.delete(colIdx);
                item.classList.remove('selected');
            } else {
                tempCfColumns.add(colIdx);
                item.classList.add('selected');
            }
        });
        cfColumnList.appendChild(item);
    });
}

cfSelectAllCols.addEventListener('click', () => {
    selectedColumns.forEach(c => tempCfColumns.add(c));
    renderCfColumnList();
});
cfDeselectAllCols.addEventListener('click', () => {
    tempCfColumns.clear();
    renderCfColumnList();
});

cfSave.addEventListener('click', () => {
    const type = cfRuleType.value;
    const v1 = cfValue1.value.trim();
    const v2 = cfValue2.value.trim();

    if (!v1) { showToast('Please enter a value', 'error'); return; }
    if (type === 'between' && !v2) { showToast('Please enter both values', 'error'); return; }
    if (tempCfColumns.size === 0) { showToast('Please select at least one column', 'error'); return; }

    const newRule = {
        type,
        value1: v1,
        value2: v2,
        bg: tempPresetBg,
        fg: tempPresetFg,
        columns: Array.from(tempCfColumns)
    };

    if (editingRuleIndex >= 0) {
        conditionalRules[editingRuleIndex] = newRule;
        showToast('Rule updated!', 'success');
    } else {
        conditionalRules.push(newRule);
        showToast('Rule added!', 'success');
    }

    closeCfModal();
    renderConditionalRules();
    buildTable();
});

// ===== Build Table =====
function buildTable() {
    const start = Math.max(0, parseInt(rowStart.value) - 1) || 0;
    const end = Math.min(currentSheetData.length - 1, parseInt(rowEnd.value)) || currentSheetData.length - 1;
    const dataRows = currentSheetData.slice(1);
    let visibleRows = dataRows.slice(start, end).map((row, idx) => ({ row, origIdx: start + idx }));
    visibleRows = visibleRows.filter(item => !isRowBlank(item.row, selectedColumns));

    previewTable.className = `fit-${columnFit.value}`;
    if (fontFamily.value === 'bookman') {
        previewTable.classList.add('font-bookman');
    }

    // Header
    previewThead.innerHTML = '';
    const headerRow = document.createElement('tr');
    selectedColumns.forEach((colIdx, position) => {
        const th = document.createElement('th');
        const col = allColumns.find(c => c.index === colIdx);
        const displayName = customColumnNames[colIdx] !== undefined ? customColumnNames[colIdx] : col.name;
        const originalName = col ? col.name : `Col ${colIdx}`;
        const currentAlign = getColumnAlignment(colIdx, position);
        th.innerHTML = `
            <div class="th-content">
                <span class="th-text" contenteditable="true" data-col-idx="${colIdx}">${displayName}</span>
                <button class="th-reset" data-col-idx="${colIdx}" title="Reset to: ${originalName}">
                    <i class="fas fa-undo"></i>
                </button>
            </div>
            <div class="align-controls">
                <button class="align-btn ${currentAlign === 'left' ? 'active' : ''}" data-col-idx="${colIdx}" data-align="left">
                    <i class="fas fa-align-left"></i>
                </button>
                <button class="align-btn ${currentAlign === 'center' ? 'active' : ''}" data-col-idx="${colIdx}" data-align="center">
                    <i class="fas fa-align-center"></i>
                </button>
                <button class="align-btn ${currentAlign === 'right' ? 'active' : ''}" data-col-idx="${colIdx}" data-align="right">
                    <i class="fas fa-align-right"></i>
                </button>
            </div>
        `;
        headerRow.appendChild(th);
    });
    previewThead.appendChild(headerRow);

    document.querySelectorAll('.th-text').forEach(el => {
        el.addEventListener('blur', (e) => {
            const idx = parseInt(e.target.dataset.colIdx);
            const newName = e.target.textContent.trim();
            const originalCol = allColumns.find(c => c.index === idx);
            if (newName === '' || newName === originalCol.name) {
                delete customColumnNames[idx];
                e.target.textContent = originalCol.name;
            } else {
                customColumnNames[idx] = newName;
            }
            renderTotalCheckboxes();
        });
        el.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); e.target.blur(); }
        });
        el.addEventListener('focus', (e) => {
            const range = document.createRange();
            range.selectNodeContents(e.target);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        });
    });

    document.querySelectorAll('.th-reset').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const idx = parseInt(btn.dataset.colIdx);
            delete customColumnNames[idx];
            buildTable();
            renderTotalCheckboxes();
        });
    });

    document.querySelectorAll('.align-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const idx = parseInt(btn.dataset.colIdx);
            const align = btn.dataset.align;
            columnAlignments[idx] = align;
            buildTable();
        });
    });

    // Body
    previewTbody.innerHTML = '';
    const columnTotals = {};
    selectedColumns.forEach(colIdx => { columnTotals[colIdx] = 0; });

    visibleRows.forEach(({ row, origIdx }) => {
        const tr = document.createElement('tr');
        selectedColumns.forEach((colIdx, position) => {
            const td = document.createElement('td');
            let cellValue = row[colIdx] !== undefined ? row[colIdx] : '';
            if (cellValue instanceof Date) cellValue = cellValue.toLocaleDateString();

            if (position === 0) {
                td.textContent = cellValue;
            } else if (isNumeric(cellValue)) {
                if (totalColumns.has(colIdx)) columnTotals[colIdx] += Number(cellValue);
                td.textContent = formatNumber(cellValue);
            } else {
                td.textContent = cellValue;
            }

            const align = getColumnAlignment(colIdx, position);
            td.classList.add(`align-${align}`);

            td.dataset.rowIdx = origIdx;
            td.dataset.colIdx = colIdx;
            const key = `${origIdx}-${colIdx}`;

            // Apply conditional formatting FIRST
            const cf = getConditionalFormat(cellValue, colIdx);
            if (cf) {
                td.classList.add('cf-formatted');
                td.style.backgroundColor = cf.bg;
                td.style.color = cf.fg;
                td.style.fontWeight = '600';
            }

            // Manual highlight overrides CF
            if (cellHighlights[key]) {
                td.classList.add('highlighted');
                td.style.backgroundColor = cellHighlights[key];
                td.style.color = '';
            }

            td.addEventListener('click', () => {
                if (currentHighlightColor === 'none') {
                    delete cellHighlights[key];
                    td.classList.remove('highlighted');
                    // Restore CF if applicable
                    const cf2 = getConditionalFormat(cellValue, colIdx);
                    if (cf2) {
                        td.style.backgroundColor = cf2.bg;
                        td.style.color = cf2.fg;
                    } else {
                        td.style.backgroundColor = '';
                        td.style.color = '';
                    }
                } else {
                    cellHighlights[key] = currentHighlightColor;
                    td.classList.add('highlighted');
                    td.style.backgroundColor = currentHighlightColor;
                    td.style.color = '';
                }
            });

            td.title = td.textContent;
            tr.appendChild(td);
        });
        previewTbody.appendChild(tr);
    });

    // Totals row
    if (visibleRows.length > 0 && totalColumns.size > 0) {
        const totalRow = document.createElement('tr');
        totalRow.className = 'total-row';
        selectedColumns.forEach((colIdx, position) => {
            const td = document.createElement('td');
            if (position === 0) td.textContent = 'TOTAL';
            else if (totalColumns.has(colIdx) && columnTotals[colIdx] !== 0) {
                td.textContent = formatNumber(columnTotals[colIdx]);
            } else td.textContent = '';
            const align = getColumnAlignment(colIdx, position);
            td.classList.add(`align-${align}`);
            totalRow.appendChild(td);
        });
        previewTbody.appendChild(totalRow);
    }

    previewInfoText.textContent = `Showing ${visibleRows.length} rows × ${selectedColumns.length} columns • ${totalColumns.size} totals • ${conditionalRules.length} rules • ${Object.keys(cellHighlights).length} highlights`;
}

// Events
rowStart.addEventListener('change', buildTable);
rowEnd.addEventListener('change', buildTable);
columnFit.addEventListener('change', buildTable);
fontFamily.addEventListener('change', buildTable);
btnAllRows.addEventListener('click', () => {
    rowStart.value = 1;
    rowEnd.value = currentSheetData.length - 1;
    buildTable();
});
btnBack3.addEventListener('click', () => goToStep(2));

document.querySelectorAll('.color-btn').forEach(btn => {
    btn.addEventListener('click', () => {
        document.querySelectorAll('.color-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        currentHighlightColor = btn.dataset.color;
    });
});

btnClearHighlights.addEventListener('click', () => {
    if (Object.keys(cellHighlights).length === 0) { showToast('No highlights to clear', 'info'); return; }
    cellHighlights = {};
    buildTable();
    showToast('All highlights cleared', 'success');
});

btnBold.addEventListener('click', () => {
    fontStyle.bold = !fontStyle.bold;
    btnBold.dataset.active = fontStyle.bold;
});
btnItalic.addEventListener('click', () => {
    fontStyle.italic = !fontStyle.italic;
    btnItalic.dataset.active = fontStyle.italic;
});
btnTotalAll.addEventListener('click', () => {
    selectedColumns.forEach((c, i) => { if (i > 0) totalColumns.add(c); });
    renderTotalCheckboxes();
    buildTable();
});
btnTotalNone.addEventListener('click', () => {
    totalColumns.clear();
    renderTotalCheckboxes();
    buildTable();
});

// ===== Export PDF =====
btnExportPdf.addEventListener('click', () => {
    if (selectedColumns.length === 0) { showToast('No columns selected.', 'error'); return; }
    showLoading('Generating PDF...');

    setTimeout(() => {
        try {
            const { jsPDF } = window.jspdf;
            const pageSize = document.getElementById('pdf-size').value;
            const fontSize = parseInt(document.getElementById('font-size').value);
            const rawFontFamily = fontFamily.value;
            const pdfFont = getPdfFontFamily(rawFontFamily); // maps bookman -> times
            const title = document.getElementById('pdf-title').value;
            const leftTitle = document.getElementById('pdf-left-title').value;
            const dateInput = document.getElementById('pdf-date').value;

            const dateText = dateInput.trim() || new Date().toLocaleDateString('en-US', {
                year: 'numeric', month: 'long', day: 'numeric'
            });

            const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: pageSize });
            doc.setFont(pdfFont);
            const pageWidth = doc.internal.pageSize.getWidth();
            const pageHeight = doc.internal.pageSize.getHeight();

            const marginLR = 12.7;
            const marginTB = 12.7;
            const usableWidth = pageWidth - (marginLR * 2);
            const HEADER_HEIGHT = 14;

            function drawHeader() {
                const headerY = marginTB + 6;
                doc.setFont(pdfFont, 'bold');
                doc.setFontSize(14);
                doc.setTextColor(0, 0, 0);
                if (leftTitle) doc.text(leftTitle.toUpperCase(), marginLR, headerY, { align: 'left' });
                if (title) doc.text(title.toUpperCase(), pageWidth / 2, headerY, { align: 'center' });
                doc.text(dateText.toUpperCase(), pageWidth - marginLR, headerY, { align: 'right' });
                doc.setTextColor(0, 0, 0);
                const lineY = headerY + 3;
                doc.setDrawColor(37, 99, 235);
                doc.setLineWidth(0.5);
                doc.line(marginLR, lineY, pageWidth - marginLR, lineY);
            }

            drawHeader();
            let startY = marginTB + HEADER_HEIGHT;

            const headers = selectedColumns.map(colIdx => {
                const col = allColumns.find(c => c.index === colIdx);
                return customColumnNames[colIdx] !== undefined ? customColumnNames[colIdx] : (col ? col.name : `Col ${colIdx}`);
            });

            const start = Math.max(0, parseInt(rowStart.value) - 1) || 0;
            const end = Math.min(currentSheetData.length - 1, parseInt(rowEnd.value)) || currentSheetData.length - 1;
            let dataRows = currentSheetData.slice(1).slice(start, end).map((row, idx) => ({ row, origIdx: start + idx }));
            dataRows = dataRows.filter(item => !isRowBlank(item.row, selectedColumns));

            const columnTotals = {};
            selectedColumns.forEach(colIdx => { columnTotals[colIdx] = 0; });

            const highlightMap = {};
            const cfMap = {}; // {bodyRowIdx-colPos: {bg, fg}}
            const rawValues = {}; // Store raw values for CF checking

            const body = dataRows.map((item, bodyRowIdx) => {
                return selectedColumns.map((colIdx, i) => {
                    let val = item.row[colIdx] !== undefined ? item.row[colIdx] : '';
                    if (val instanceof Date) val = val.toLocaleDateString();

                    const key = `${item.origIdx}-${colIdx}`;
                    rawValues[`${bodyRowIdx}-${i}`] = val;

                    // Check CF
                    const cf = getConditionalFormat(val, colIdx);
                    if (cf) cfMap[`${bodyRowIdx}-${i}`] = cf;

                    // Manual highlight overrides
                    if (cellHighlights[key]) highlightMap[`${bodyRowIdx}-${i}`] = cellHighlights[key];

                    if (i === 0) return String(val);
                    if (isNumeric(val)) {
                        if (totalColumns.has(colIdx)) columnTotals[colIdx] += Number(val);
                        return formatNumber(val);
                    }
                    return String(val);
                });
            });

            let totalRowIdx = -1;
            if (totalColumns.size > 0) {
                const totalRow = selectedColumns.map((colIdx, i) => {
                    if (i === 0) return 'TOTAL';
                    if (totalColumns.has(colIdx) && columnTotals[colIdx] !== 0) return formatNumber(columnTotals[colIdx]);
                    return '';
                });
                body.push(totalRow);
                totalRowIdx = body.length - 1;
            }

            const columnStyles = {};
            selectedColumns.forEach((colIdx, i) => {
                const align = getColumnAlignment(colIdx, i);
                columnStyles[i] = { halign: align };
            });

            let bodyFontStyle = 'normal';
            if (fontStyle.bold && fontStyle.italic) bodyFontStyle = 'bolditalic';
            else if (fontStyle.bold) bodyFontStyle = 'bold';
            else if (fontStyle.italic) bodyFontStyle = 'italic';

            doc.autoTable({
                head: [headers],
                body: body,
                startY: startY,
                theme: 'grid',
                tableWidth: usableWidth,
                margin: { top: marginTB + HEADER_HEIGHT, right: marginLR, bottom: marginTB + 5, left: marginLR },
                styles: {
                    font: pdfFont, fontStyle: bodyFontStyle, fontSize: fontSize,
                    cellPadding: 1.8, overflow: 'linebreak',
                    lineColor: [180, 180, 180], lineWidth: 0.2,
                    halign: 'center', valign: 'middle',
                },
                headStyles: {
                    font: pdfFont, fillColor: [37, 99, 235], textColor: [255, 255, 255],
                    fontStyle: 'bold', fontSize: fontSize + 0.5,
                    halign: 'center', valign: 'middle', cellPadding: 2.5,
                },
                alternateRowStyles: { fillColor: [245, 247, 250] },
                columnStyles: columnStyles,
                didParseCell: function (data) {
                    // Totals
                    if (totalRowIdx > -1 && data.row.index === totalRowIdx && data.row.section === 'body') {
                        data.cell.styles.fillColor = [254, 243, 199];
                        data.cell.styles.textColor = [146, 64, 14];
                        data.cell.styles.fontStyle = 'bold';
                        data.cell.styles.lineWidth = 0.4;
                        data.cell.styles.lineColor = [245, 158, 11];
                        return;
                    }

                    const key = `${data.row.index}-${data.column.index}`;

                    // Manual highlight (highest priority)
                    if (highlightMap[key] && data.row.section === 'body') {
                        data.cell.styles.fillColor = hexToRgb(highlightMap[key]);
                        data.cell.styles.fontStyle = 'bold';
                        return;
                    }

                    // Conditional formatting
                    if (cfMap[key] && data.row.section === 'body') {
                        data.cell.styles.fillColor = hexToRgb(cfMap[key].bg);
                        data.cell.styles.textColor = hexToRgb(cfMap[key].fg);
                        data.cell.styles.fontStyle = 'bold';
                    }
                },
                didDrawPage: function (data) {
                    if (data.pageNumber > 1) drawHeader();
                    const pageCount = doc.internal.getNumberOfPages();
                    const currentPage = doc.internal.getCurrentPageInfo().pageNumber;
                    doc.setFontSize(7);
                    doc.setFont(pdfFont, 'normal');
                    doc.setTextColor(150);
                    doc.text(`Page ${currentPage} of ${pageCount}`, pageWidth / 2, pageHeight - 5, { align: 'center' });
                },
            });

            const pdfName = title
                ? title.replace(/[^a-zA-Z0-9]/g, '_') + '.pdf'
                : fileName.replace(/\.[^/.]+$/, '') + '_export.pdf';

            doc.save(pdfName);
            hideLoading();
            showToast('PDF exported successfully!', 'success');
        } catch (err) {
            hideLoading();
            showToast('Error generating PDF.', 'error');
            console.error(err);
        }
    }, 300);
});