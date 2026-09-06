// ============================================
// Excel -> JRXML pipeline (shared browser helpers)
// Extracted from app.js — the Excel→JRXML conversion behavior is unchanged:
//   load workbook -> cloneWithoutColumnA (col A holds section markers)
//   -> serialize to plain JSON -> exportToJrxmlV6() -> JRXML text
// ============================================

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
 */
function safeStringify(obj, space = 2) {
    const seen = new WeakSet();
    try {
        return JSON.stringify(obj, (key, value) => {
            if (typeof value === 'object' && value !== null) {
                if (seen.has(value)) return '[Circular]';
                seen.add(value);
            }
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
        // Skip empty-string markers (continuation rows of the previous section)
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

    if (ws.properties) {
        if (ws.properties.defaultRowHeight != null)
            data.properties.defaultRowHeight = ws.properties.defaultRowHeight;
        if (ws.properties.defaultColWidth != null)
            data.properties.defaultColWidth = ws.properties.defaultColWidth;
    }

    const maxCol = ws.lastColumn?.number || 0;
    for (let i = 1; i <= maxCol; i++) {
        const col = ws.getColumn(i);
        data.columns.push({ width: col.width });
    }

    ws.eachRow({ includeEmpty: true }, (row, rIdx) => {
        const rowData = { rowNumber: rIdx, cells: [] };
        if (row.height) rowData.height = row.height;

        row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
            const cellData = { colNumber: cIdx };

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

            if (cell.style) cellData.style = safeClone(cell.style);
            if (cell.numFmt) cellData.numFmt = cell.numFmt;
            if (cell.font) cellData.font = safeClone(cell.font);
            if (cell.fill) cellData.fill = safeClone(cell.fill);
            if (cell.border) cellData.border = safeClone(cell.border);
            if (cell.alignment) cellData.alignment = safeClone(cell.alignment);

            if (cell.note) cellData.note = safeClone(cell.note);

            rowData.cells.push(cellData);
        });

        data.rows.push(rowData);
    });

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

    if (workbook._themes?.theme1) {
        data.themeXml = workbook._themes.theme1;
    }

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

    if (data.properties) {
        if (data.properties.defaultRowHeight != null)
            ws.properties.defaultRowHeight = data.properties.defaultRowHeight;
        if (data.properties.defaultColWidth != null)
            ws.properties.defaultColWidth = data.properties.defaultColWidth;
    }

    data.columns.forEach((col, i) => {
        if (col.width != null) ws.getColumn(i + 1).width = col.width;
    });

    data.rows.forEach(rowData => {
        const row = ws.getRow(rowData.rowNumber);
        if (rowData.height) row.height = rowData.height;

        rowData.cells.forEach(cellData => {
            const cell = row.getCell(cellData.colNumber);

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

            if (cellData.style) cell.style = JSON.parse(JSON.stringify(cellData.style));
            if (cellData.numFmt) cell.numFmt = cellData.numFmt;

            if (cellData.note) cell.note = JSON.parse(JSON.stringify(cellData.note));
        });

        row.commit();
    });

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

    if (data.themeXml) {
        wb._themes = { theme1: data.themeXml };
    }

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
        // ⭐ MUST also set wb.media directly — ExcelJS's getImage(id) reads this.media[]
        wb.media = mediaArray;
    }

    return wb;
}

/**
 * Convert an uploaded Excel workbook into the exact input shape that
 * exportToJrxmlV6() consumes. Returns the JRXML string.
 */
function buildJrxml(workbook, wsIndex, fileName) {
    const ws = workbook.worksheets[wsIndex];
    if (!ws) throw new Error('ไม่พบชีตที่เลือก');

    const originalImages = ws.getImages();
    const { worksheet, meta } = cloneWithoutColumnA(workbook, ws);

    const exportWs = cloneWorksheet(worksheet);
    const exportImages = shiftImagesLeft(originalImages, 1);

    const input = {
        ws: worksheetToJson(exportWs),
        workbook: workbookToJson(workbook),
        meta: meta,
        images: exportImages,
        fileName: fileName
    };

    return exportToJrxmlV6(input);
}
