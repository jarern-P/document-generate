// ============================================
// Excel Report Code Generator - Main App
// ============================================

let currentState = {
    workbook: null,
    worksheet: null,
    originalWorksheet: null,
    meta: [],
    originalImages: [],
    fileName: '',
    fileSize: 0,
    sheetNames: []
};

document.addEventListener('DOMContentLoaded', () => {
    initApp();
});

function initApp() {
    const elements = {
        fileInput: document.getElementById('fileInput'),
        dropZone: document.getElementById('dropZone'),
        browseLink: document.getElementById('browseLink'),
        fileInfo: document.getElementById('fileInfo'),
        fileName: document.getElementById('fileName'),
        fileSize: document.getElementById('fileSize'),
        sheetCount: document.getElementById('sheetCount'),
        sheetSelector: document.getElementById('sheetSelector'),
        toolbar: document.getElementById('toolbar'),
        sectionsPanel: document.getElementById('sectionsPanel'),
        sectionsList: document.getElementById('sectionsList'),
        previewContainer: document.getElementById('previewContainer'),
        tableContainer: document.getElementById('tableContainer'),
        loadingOverlay: document.getElementById('loadingOverlay'),
        loadingText: document.getElementById('loadingText'),
        exportNpoiBtn: document.getElementById('exportNpoiBtn'),
        exportJrxmlBtn: document.getElementById('exportJrxmlBtn'),
        exportRdlcBtn: document.getElementById('exportRdlcBtn'),
        exportNgprimeBtn: document.getElementById('exportNgprimeBtn'),
        exportXlsxStyleBtn: document.getElementById('exportXlsxStyleBtn'),
        exportEfCoreBtn: document.getElementById('exportEfCoreBtn'),
        exportTypeScriptBtn: document.getElementById('exportTypeScriptBtn'),
        exportItextsharpBtn: document.getElementById('exportItextsharpBtn'),
        clearBtn: document.getElementById('clearBtn')
    };

    // ============================
    // DRAG & DROP
    // ============================
    elements.dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        e.stopPropagation();
        elements.dropZone.classList.add('drag-over');
    });

    elements.dropZone.addEventListener('dragleave', () => {
        elements.dropZone.classList.remove('drag-over');
    });

    elements.dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        e.stopPropagation();
        elements.dropZone.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file) handleFile(file, elements);
    });

    elements.browseLink.addEventListener('click', (e) => {
        e.preventDefault();
        elements.fileInput.click();
    });

    elements.dropZone.addEventListener('click', () => {
        if (!elements.dropZone.classList.contains('has-file')) {
            elements.fileInput.click();
        }
    });

    elements.fileInput.addEventListener('click', () => {
        elements.fileInput.value = '';
    });

    elements.fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleFile(file, elements);
    });

    // ============================
    // SHEET SELECTOR
    // ============================
    elements.sheetSelector.addEventListener('change', () => {
        const idx = parseInt(elements.sheetSelector.value);
        if (!isNaN(idx) && currentState.workbook) {
            loadSheet(idx, elements);
        }
    });

    // ============================
    // EXPORT BUTTONS
    // ============================
    elements.exportNpoiBtn.addEventListener('click', () => doExport('npoi', elements));
    elements.exportJrxmlBtn.addEventListener('click', () => doExport('jrxml', elements));
    elements.exportRdlcBtn.addEventListener('click', () => doExport('rdlc', elements));
    elements.exportNgprimeBtn.addEventListener('click', () => doExport('ngprime', elements));
    elements.exportXlsxStyleBtn.addEventListener('click', () => doExport('xlsxstyle', elements));
    elements.exportEfCoreBtn.addEventListener('click', () => doExport('efcore', elements));
    elements.exportTypeScriptBtn.addEventListener('click', () => doExport('typescript', elements));
    elements.exportItextsharpBtn.addEventListener('click', () => doExport('itextsharp', elements));
    elements.clearBtn.addEventListener('click', () => clearFile(elements));
}

// ============================================
// FILE HANDLING
// ============================================
async function handleFile(file, elements) {
    if (!file.name.endsWith('.xlsx')) {
        showToast('กรุณาเลือกไฟล์ .xlsx เท่านั้น', 'error');
        return;
    }

    showLoading('กำลังอ่านไฟล์...', elements);

    try {
        const arrayBuffer = await file.arrayBuffer();

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer, {
            comments: 'emit',
            ignoreNodes: ['tableParts', 'autoFilter', 'dataValidations']
        });

        // Reset state
        currentState.workbook = workbook;
        currentState.sheetNames = workbook.worksheets.map(ws => ws.name);
        currentState.fileName = file.name;
        currentState.fileSize = file.size;

        // Populate sheet selector
        elements.sheetSelector.innerHTML = '';
        workbook.worksheets.forEach((ws, i) => {
            const opt = document.createElement('option');
            opt.value = i;
            opt.textContent = `${ws.name} (Sheet ${i + 1})`;
            if (i === 0) opt.selected = true;
            elements.sheetSelector.appendChild(opt);
        });

        // Load first sheet
        await loadSheet(0, elements);

        // Update UI
        elements.dropZone.style.display = 'none';
        elements.toolbar.style.display = 'flex';
        elements.fileName.textContent = file.name;
        elements.fileSize.textContent = formatFileSize(file.size);
        elements.sheetCount.textContent = workbook.worksheets.length;

        hideLoading(elements);
        showToast('โหลดไฟล์สำเร็จ ✓', 'success');

    } catch (err) {
        console.error('File load error:', err);
        hideLoading(elements);
        showToast('เกิดข้อผิดพลาดในการอ่านไฟล์: ' + (err.message || 'ไม่ทราบสาเหตุ'), 'error');
    }
}

// ============================================
// SHEET LOADING
// ============================================
async function loadSheet(index, elements) {
    showLoading('กำลังโหลดชีต...', elements);

    try {
        const ws = currentState.workbook.worksheets[index];
        currentState.originalWorksheet = ws;
        currentState.originalImages = ws.getImages();

        // Clone and remove Column A
        const result = cloneWithoutColumnA(currentState.workbook, ws);
        currentState.worksheet = result.worksheet;
        currentState.meta = result.meta;

        // Render preview
        elements.tableContainer.innerHTML = '';
        const renderWs = cloneWorksheet(ws);
        renderTable(renderWs, currentState.workbook, ws.getImages());

        // Show sections
        renderSections(currentState.meta, elements);

        // Show panels
        elements.sectionsPanel.style.display = 'block';
        elements.previewContainer.style.display = 'block';

        hideLoading(elements);

    } catch (err) {
        console.error('Sheet load error:', err);
        hideLoading(elements);
        showToast('เกิดข้อผิดพลาดในการโหลดชีต', 'error');
    }
}

// ============================================
// SECTION VISUALIZATION
// ============================================
function renderSections(meta, elements) {
    const list = elements.sectionsList;
    list.innerHTML = '';

    if (!meta || meta.length === 0) {
        list.innerHTML = '<span class="section-tag section-default">ไม่มีส่วนกำหนดใน Column A (แสดงทั้งหมด)</span>';
        return;
    }

    const totalRows = currentState.originalWorksheet.rowCount || 
                      currentState.originalWorksheet.lastRow?.number || 0;

    for (let i = 0; i < meta.length; i++) {
        const current = meta[i];
        const next = meta[i + 1];
        const endRow = next ? next.row - 1 : totalRows;

        const canonicalKey = normalizeSectionKey(current.key);
        const displayName = getSectionDisplayName(current.key);

        const tag = document.createElement('span');
        tag.className = `section-tag section-${canonicalKey || current.key.toLowerCase()}`;

        // Build export compatibility badges
        let badges = '';
        const exportTypes = ['npoi', 'jrxml', 'rdlc', 'ngprime', 'xlsxstyle', 'efcore', 'typescript', 'itextsharp'];
        exportTypes.forEach(type => {
            if (isSectionExportable(current.key, type)) {
                const config = EXPORT_SECTION_CONFIG[type];
                badges += `<span class="export-badge export-${type}">${config ? config.label : type.toUpperCase()}</span>`;
            }
        });

        // If this is a group:Name definition, show it with special styling
        if (/^group[:_]/i.test(String(current.key).trim())) {
            const groupTag = document.createElement('span');
            groupTag.className = 'section-tag section-group';
            groupTag.innerHTML = `${escapeHtml(current.key)} <span class="section-range">แถว ${current.row}</span>`;
            list.appendChild(groupTag);
        } else {
            tag.innerHTML = `${escapeHtml(displayName)} <span class="section-range">แถว ${current.row} - ${endRow}</span> ${badges ? `<span class="section-badges">${badges}</span>` : ''}`;
            list.appendChild(tag);
        }
    }

    // Add export compatibility legend
    const existingLegend = list.querySelector('.sections-legend');
    if (existingLegend) existingLegend.remove();

    const legend = document.createElement('div');
    legend.className = 'sections-legend';
    legend.innerHTML = `
        <span class="legend-item"><span class="legend-badge" style="background:#1e293b"></span> NPOI</span>
        <span class="legend-item"><span class="legend-badge" style="background:#7c3aed"></span> JRXML</span>
        <span class="legend-item"><span class="legend-badge" style="background:#0891b2"></span> RDLC</span>
        <span class="legend-item"><span class="legend-badge" style="background:#c5280d"></span> PrimeNG</span>
        <span class="legend-item"><span class="legend-badge" style="background:#059669"></span> xlsx-style</span>
        <span class="legend-item"><span class="legend-badge" style="background:#2563eb"></span> EF Core</span>
        <span class="legend-item"><span class="legend-badge" style="background:#d97706"></span> TypeScript</span>
        <span class="legend-item"><span class="legend-badge" style="background:#c026d3"></span> iTextSharp</span>
        <span class="legend-item">✓ = ส่วนนี้สามารถ Export ได้</span>
    `;
    list.appendChild(legend);
}

// ============================================
// CLEAR / RESET
// ============================================
function clearFile(elements) {
    currentState = {
        workbook: null,
        worksheet: null,
        originalWorksheet: null,
        meta: [],
        originalImages: [],
        fileName: '',
        fileSize: 0,
        sheetNames: []
    };

    elements.toolbar.style.display = 'none';
    elements.dropZone.style.display = 'block';
    elements.dropZone.classList.remove('has-file');
    elements.dropZone.querySelector('.drop-text').textContent = 'ลากไฟล์ Excel (.xlsx) มาไว้ที่นี่';
    elements.sectionsPanel.style.display = 'none';
    elements.previewContainer.style.display = 'none';
    elements.tableContainer.innerHTML = '';
    elements.sectionsList.innerHTML = '';
    elements.sheetSelector.innerHTML = '';
    elements.fileInput.value = '';

    showToast('ยกเลิกไฟล์เรียบร้อย', 'info');
}

// ============================================
// EXPORT
// ============================================
function doExport(type, elements) {
    if (!currentState.worksheet) {
        showToast('กรุณา upload Excel ก่อน', 'warning');
        return;
    }

    try {
        const exportWs = cloneWorksheet(currentState.worksheet);
        const exportImages = shiftImagesLeft(currentState.originalImages, 1);

        // Build a clean JSON input object — no raw ExcelJS internals leaked
        const input = {
            ws: worksheetToJson(exportWs),
            workbook: workbookToJson(currentState.workbook),
            meta: currentState.meta,
            images: exportImages,
            fileName: currentState.fileName
        };

        // Safe JSON log for debugging (handles circular references gracefully)
        console.log('=== EXPORT INPUT JSON ===');
        console.log(safeStringify(input));
        // const json = safeStringify(input);
        // let testInput = JSON.parse(json);
        // exportJson(input);

        let code = '';
        let filename = '';

        switch (type) {
            case 'npoi':
                code = exportToNPOI(input);
                filename = 'npoi-code.txt';
                break;
            case 'jrxml':
                code = exportToJrxmlV6(input);
                filename = 'jrxml-code.txt';
                break;
            case 'rdlc':
                code = exportToRDLC(input);
                filename = 'rdlc-code.txt';
                break;
            case 'ngprime':
                code = exportToNgPrime(input);
                filename = 'ngprime-code.txt';
                break;
            case 'xlsxstyle':
                code = exportToXlsxStyle(input);
                filename = 'xlsx-style-export.txt';
                break;
            case 'efcore':
                code = exportToEfCore(input);
                filename = 'efcore-code.txt';
                break;
            case 'typescript':
                code = exportToTypeScript(input);
                filename = 'typescript-code.txt';
                break;
            case 'itextsharp':
                code = exportToItextSharp(input);
                filename = 'itextsharp-code.txt';
                break;
            // case 'json':
            //     code = exportJson(input);
            //     filename = 'export.json';
            //     break;
        }

        downloadFile(code, filename);
        showToast(`ส่งออก ${filename} สำเร็จ ✓`, 'success');

    } catch (err) {
        console.error('Export error:', err);
        showToast('เกิดข้อผิดพลาดในการส่งออก', 'error');
    }
}

function downloadFile(content, filename) {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ============================================
// LOADING OVERLAY
// ============================================
function showLoading(text, elements) {
    elements.loadingText.textContent = text || 'กำลังประมวลผล...';
    elements.loadingOverlay.style.display = 'flex';
}

function hideLoading(elements) {
    elements.loadingOverlay.style.display = 'none';
}

// ============================================
// TOAST NOTIFICATIONS
// ============================================
function showToast(message, type = 'info') {
    const container = document.getElementById('toastContainer');
    if (!container) return;

    const icons = {
        success: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>',
        error: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>',
        warning: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>',
        info: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#0891b2" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>'
    };

    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
        <div class="toast-icon">${icons[type] || icons.info}</div>
        <div class="toast-message">${escapeHtml(message)}</div>
    `;
    container.appendChild(toast);

    setTimeout(() => {
        if (toast.parentNode) {
            toast.remove();
        }
    }, 4000);
}

// ============================================
// UTILITY FUNCTIONS
// ============================================
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
}

// ============================================
// SAFE JSON CLONE & STRINGIFY (handles circular references)
// ============================================

/**
 * Deep clone any value safely — handles circular references,
 * Dates, RegExps, Maps, Sets, and plain objects/arrays.
 * Falls back to a simple toString() for unserializable values.
 */
function safeClone(obj) {
    // Primitives & null
    if (obj === null || obj === undefined) return obj;
    if (typeof obj !== 'object') return obj;

    // Date
    if (obj instanceof Date) return new Date(obj.getTime());

    // Use try-catch around JSON approach first (fastest for plain objects)
    try {
        return JSON.parse(JSON.stringify(obj));
    } catch (_) {
        // Fallback: manual copy with circular reference tracking
        const seen = new WeakSet();
        function deepCopy(val) {
            if (val === null || val === undefined) return val;
            if (typeof val !== 'object') return val;
            if (val instanceof Date) return new Date(val.getTime());
            if (seen.has(val)) return '[Circular]';
            seen.add(val);

            if (Array.isArray(val)) {
                return val.map(deepCopy);
            }

            const result = {};
            for (const key of Object.keys(val)) {
                try {
                    result[key] = deepCopy(val[key]);
                } catch (_) {
                    result[key] = String(val[key]);
                }
            }
            return result;
        }
        return deepCopy(obj);
    }
}

/**
 * Safe JSON.stringify that handles circular references and
 * other unserializable values gracefully.
 * Uses a WeakSet to detect and replace circular references.
 */
function safeStringify(obj, space = 2) {
    const seen = new WeakSet();
    try {
        return JSON.stringify(obj, (key, value) => {
            if (typeof value === 'object' && value !== null) {
                if (seen.has(value)) return '[Circular]';
                seen.add(value);
            }
            // Handle non-enumerable special types
            if (typeof value === 'function') return '[Function]';
            if (typeof value === 'symbol') return value.toString();
            if (value instanceof Error) return value.message;
            return value;
        }, space);
    } catch (err) {
        return `[safeStringify error: ${err.message}]`;
    }
}

// ============================================
// EXCEL UTILITY FUNCTIONS
// ============================================

function cloneWithoutColumnA(workbook, ws, sheetName = 'FilteredSheet') {
    // Remove previous temp sheet if exists (fixes error on sheet switch)
    const existing = workbook.getWorksheet(sheetName);
    if (existing) {
        workbook.removeWorksheet(existing.id);
    }
    const newWs = workbook.addWorksheet(sheetName);
    const meta = [];

    ws.eachRow({ includeEmpty: true }, (row, rIdx) => {
        const cellA = row.getCell(1);
        // Skip empty-string markers (continuation rows of the previous section),
        // otherwise each continuation row would become its own empty section and
        // multi-line sections (e.g. detail) would collapse to a single row.
        if (cellA && cellA.value != null && String(cellA.value).trim() !== '') {
            meta.push({
                row: rIdx,
                key: cellA.value
            });
        }
    });

    const maxCol = ws.lastColumn?.number || 0;

    // COLUMN WIDTH
    for (let i = 2; i <= maxCol; i++) {
        const col = ws.getColumn(i);
        const newCol = newWs.getColumn(i - 1);
        if (col?.width) newCol.width = col.width;
    }

    // COPY ROWS + CELLS
    ws.eachRow({ includeEmpty: true }, (row, rIdx) => {
        const newRow = newWs.getRow(rIdx);
        if (row.height) newRow.height = row.height;

        row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
            if (cIdx === 1) return;

            const newCell = newRow.getCell(cIdx - 1);

            if (cell.value === null || cell.value === undefined) {
                newCell.value = null;
            } else if (typeof cell.value === 'object' && cell.value.formula) {
                newCell.value = { formula: cell.value.formula, result: safeClone(cell.value.result) };
            } else if (typeof cell.value === 'object' && cell.value.richText) {
                newCell.value = safeClone(cell.value);
            } else {
                newCell.value = cell.value;
            }

            if (cell.style) newCell.style = safeClone(cell.style);
            if (cell.numFmt) newCell.numFmt = cell.numFmt;

            if (cell.hyperlink) {
                newCell.value = { text: cell.text, hyperlink: cell.hyperlink };
            }

            if (cell.note) {
                newCell.note = safeClone(cell.note);
            }
        });

        newRow.commit();
    });

    // MERGE CELLS
    if (ws.model?.merges) {
        ws.model.merges.forEach(m => {
            const [start, end] = m.split(':');
            const s = decodeAddr(start);
            const e = decodeAddr(end);
            if (e.c <= 1) return;
            const newStartCol = Math.max(1, s.c - 1);
            const newEndCol = Math.max(1, e.c - 1);
            newWs.mergeCells(s.r, newStartCol, e.r, newEndCol);
        });
    }

    return { worksheet: newWs, meta };
}

function cloneWorksheet(ws) {
    const tempWb = new ExcelJS.Workbook();
    const newWs = tempWb.addWorksheet('clone');

    const maxCol = ws.lastColumn?.number || 0;
    for (let i = 1; i <= maxCol; i++) {
        const col = ws.getColumn(i);
        const newCol = newWs.getColumn(i);
        if (col?.width) newCol.width = col.width;
    }

    ws.eachRow({ includeEmpty: true }, (row, r) => {
        const newRow = newWs.getRow(r);
        if (row.height) newRow.height = row.height;

        row.eachCell({ includeEmpty: true }, (cell, c) => {
            const newCell = newRow.getCell(c);

            if (cell.value === null || cell.value === undefined) {
                newCell.value = null;
            } else if (typeof cell.value === 'object') {
                newCell.value = safeClone(cell.value);
            } else {
                newCell.value = cell.value;
            }

            if (cell.style) newCell.style = safeClone(cell.style);
            if (cell.numFmt) newCell.numFmt = cell.numFmt;
            if (cell.border) newCell.border = safeClone(cell.border);

            const note = cell.note || cell.comment || cell.model?.note;
            if (note) {
                newCell.note = safeClone(note);
            }
        });

        newRow.commit();
    });

    if (ws.model?.merges) {
        ws.model.merges.forEach(m => {
            const [start, end] = m.split(':');
            const s = decodeAddr(start);
            const e = decodeAddr(end);
            newWs.mergeCells(s.r, s.c, e.r, e.c);
        });
    }

    return newWs;
}

function shiftImagesLeft(images, shiftCols = 1) {
    return images
        .map(img => {
            if (!img || !img.range?.tl) return null;

            // Manual safe clone — avoid structuredClone which may preserve ExcelJS internals
            const tl = img.range.tl;
            const br = img.range.br;
            const ext = img.range.ext;

            const clonedTl = {
                nativeCol: tl.nativeCol,
                nativeRow: tl.nativeRow,
                nativeColOff: tl.nativeColOff,
                nativeRowOff: tl.nativeRowOff
            };

            const clonedBr = br ? {
                nativeCol: br.nativeCol,
                nativeRow: br.nativeRow,
                nativeColOff: br.nativeColOff,
                nativeRowOff: br.nativeRowOff
            } : null;

            const cloned = {
                imageId: img.imageId,
                name: img.name,
                mimeType: img.mimeType,
                range: {
                    tl: clonedTl,
                    br: clonedBr,
                    ext: ext ? { ...ext } : undefined
                }
            };

            // Shift columns left
            if (clonedTl.nativeCol != null) {
                clonedTl.nativeCol = Math.max(0, clonedTl.nativeCol - shiftCols);
                if (clonedBr?.nativeCol != null) {
                    clonedBr.nativeCol = Math.max(0, clonedBr.nativeCol - shiftCols);
                }
                if (clonedBr?.nativeCol < 0) return null;
            }

            if (ext && clonedTl.nativeCol != null) {
                clonedTl.nativeCol = Math.max(0, clonedTl.nativeCol - shiftCols);
            }

            return cloned;
        })
        .filter(Boolean);
}

// ============================================
// JSON SERIALIZATION — Extract ExcelJS → JSON
// ============================================

function worksheetToJson(ws) {
    const data = {
        columns: [],
        rows: [],
        merges: [],
        properties: {}
    };

    // Properties
    if (ws.properties) {
        if (ws.properties.defaultRowHeight != null)
            data.properties.defaultRowHeight = ws.properties.defaultRowHeight;
        if (ws.properties.defaultColWidth != null)
            data.properties.defaultColWidth = ws.properties.defaultColWidth;
    }

    // Columns
    const maxCol = ws.lastColumn?.number || 0;
    for (let i = 1; i <= maxCol; i++) {
        const col = ws.getColumn(i);
        data.columns.push({ width: col.width });
    }

    // Rows & Cells
    ws.eachRow({ includeEmpty: true }, (row, rIdx) => {
        const rowData = { rowNumber: rIdx, cells: [] };
        if (row.height) rowData.height = row.height;

        row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
            const cellData = { colNumber: cIdx };

            // Value
            if (cell.value === null || cell.value === undefined) {
                cellData.value = null;
            } else if (typeof cell.value === 'object' && cell.value.formula) {
                cellData.value = { formula: cell.value.formula, result: safeClone(cell.value.result) };
            } else if (typeof cell.value === 'object' && cell.value.richText) {
                cellData.value = safeClone(cell.value);
            } else if (typeof cell.value === 'object' && cell.value.text) {
                cellData.value = { text: cell.value.text, hyperlink: cell.value.hyperlink };
            } else if (typeof cell.value === 'object') {
                cellData.value = safeClone(cell.value);
            } else {
                cellData.value = cell.value;
            }

            // Style properties — use safeClone to handle any circular references
            if (cell.style) cellData.style = safeClone(cell.style);
            if (cell.numFmt) cellData.numFmt = cell.numFmt;
            if (cell.font) cellData.font = safeClone(cell.font);
            if (cell.fill) cellData.fill = safeClone(cell.fill);
            if (cell.border) cellData.border = safeClone(cell.border);
            if (cell.alignment) cellData.alignment = safeClone(cell.alignment);

            // Note
            if (cell.note) cellData.note = safeClone(cell.note);

            rowData.cells.push(cellData);
        });

        data.rows.push(rowData);
    });

    // Merges
    if (ws.model?.merges) {
        data.merges = safeClone(ws.model.merges);
    }

    return data;
}

function workbookToJson(workbook) {
    const data = {
        themeXml: null,
        media: []
    };

    // Theme XML (used by getThemeColors)
    if (workbook._themes?.theme1) {
        data.themeXml = workbook._themes.theme1;
    }

    // Media (image buffers) — convert to base64 for JSON safety
    if (workbook.model?.media) {
        workbook.model.media.forEach(m => {
            if (!m) return;
            const item = {
                extension: m.extension,
                mimeType: m.mimeType,
                name: m.name
            };
            if (m.buffer) {
                item.buffer = toBase64(m.buffer);
            }
            data.media.push(item);
        });
    }

    return data;
}


// ============================================
// JSON DESERIALIZATION — Rebuild ExcelJS from JSON
// ============================================

function buildWorksheetFromJson(data) {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('rebuilt');

    // Properties
    if (data.properties) {
        if (data.properties.defaultRowHeight != null)
            ws.properties.defaultRowHeight = data.properties.defaultRowHeight;
        if (data.properties.defaultColWidth != null)
            ws.properties.defaultColWidth = data.properties.defaultColWidth;
    }

    // Columns
    data.columns.forEach((col, i) => {
        if (col.width != null) ws.getColumn(i + 1).width = col.width;
    });

    // Rows
    data.rows.forEach(rowData => {
        const row = ws.getRow(rowData.rowNumber);
        if (rowData.height) row.height = rowData.height;

        rowData.cells.forEach(cellData => {
            const cell = row.getCell(cellData.colNumber);

            // Value
            if (cellData.value === null || cellData.value === undefined) {
                cell.value = null;
            } else if (typeof cellData.value === 'object' && cellData.value.formula) {
                cell.value = { formula: cellData.value.formula, result: cellData.value.result };
            } else if (typeof cellData.value === 'object' && cellData.value.richText) {
                cell.value = JSON.parse(JSON.stringify(cellData.value));
            } else if (typeof cellData.value === 'object' && cellData.value.text) {
                cell.value = { text: cellData.value.text, hyperlink: cellData.value.hyperlink };
            } else if (typeof cellData.value === 'object') {
                cell.value = JSON.parse(JSON.stringify(cellData.value));
            } else {
                cell.value = cellData.value;
            }

            // Styles
            if (cellData.style) cell.style = JSON.parse(JSON.stringify(cellData.style));
            if (cellData.numFmt) cell.numFmt = cellData.numFmt;

            // Note
            if (cellData.note) cell.note = JSON.parse(JSON.stringify(cellData.note));
        });

        row.commit();
    });

    // Merges
    if (data.merges) {
        data.merges.forEach(m => {
            const [start, end] = m.split(':');
            const s = decodeAddr(start);
            const e = decodeAddr(end);
            ws.mergeCells(s.r, s.c, e.r, e.c);
        });
    }

    return ws;
}

function buildWorkbookFromJson(data) {
    const wb = new ExcelJS.Workbook();

    // Theme XML (for getThemeColors)
    if (data.themeXml) {
        wb._themes = { theme1: data.themeXml };
    }

    // Media (images) — convert base64 back to ArrayBuffer
    if (data.media && data.media.length > 0) {
        const mediaArray = data.media.map(m => {
            let buffer = null;
            if (m.buffer) {
                const binary = atob(m.buffer);
                const bytes = new Uint8Array(binary.length);
                for (let i = 0; i < binary.length; i++) {
                    bytes[i] = binary.charCodeAt(i);
                }
                buffer = bytes.buffer;
            }
            return {
                buffer: buffer,
                extension: m.extension,
                mimeType: m.mimeType,
                name: m.name
            };
        });
        if (!wb.model) wb.model = {};
        wb.model.media = mediaArray;
        // ⭐ MUST also set wb.media directly — ExcelJS's getImage(id) reads from this.media[],
        // NOT from this.model.media[]. Without this, getImage() returns undefined.
        wb.media = mediaArray;
    }

    return wb;
}

function exportJson(obj) {
    const json = safeStringify(obj);

    const blob = new Blob([json], {
        type: 'application/json'
    });

    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'test.json';
    a.click();

    URL.revokeObjectURL(url);
}