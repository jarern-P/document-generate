
function renderTable(ws, workbook, images = []) {;
    normalizeBorders(ws);
    const container = document.getElementById('tableContainer');
    container.innerHTML = '';
    container.style.position = 'relative';

    const table = document.createElement('table');
    table.style.borderCollapse = 'collapse';
    table.style.borderSpacing = '0';
    table.style.tableLayout = 'fixed';
    const themeColors = getThemeColors(workbook);

    const colWidths = [];
    const rowHeights = [];
    const colOffsets = [];
    const rowOffsets = [];

    // =========================
    // ✅ STYLE CACHE (เหมือน NPOI)
    // =========================
    const styleCache = new Map();
    let styleIndex = 0;
    const styleSheet = document.createElement('style');
    document.head.appendChild(styleSheet);

    // =========================
    // ✅ COLUMN WIDTH (สำคัญมาก)
    // =========================
    // function getColWidthPx(width) {
    //     return Math.floor((width || 8.43) * 7 + 5);
    // }

    const DEFAULT_COL_WIDTH = ws.properties?.defaultColWidth ?? 8.43;

    ws.columns.forEach((col, i) => {
        colWidths[i] = getColWidthPx(col.width || DEFAULT_COL_WIDTH);
    });

    // =========================
    // ✅ ROW HEIGHT
    // =========================
    const DEFAULT_ROW_HEIGHT = ws.properties?.defaultRowHeight ?? 15;

    ws.eachRow({ includeEmpty: true }, (row, r) => {
        rowHeights[r - 1] = (row.height || DEFAULT_ROW_HEIGHT) * 96 / 72;
    });

    // =========================
    // ✅ OFFSETS (สำหรับ image)
    // =========================
    let x = 0;
    colWidths.forEach((w, i) => {
        colOffsets[i] = x;
        x += w;
    });

    let y = 0;
    rowHeights.forEach((h, i) => {
        rowOffsets[i] = y;
        y += h;
    });

    // =========================
    // ✅ MERGE MAP
    // =========================
    const mergeMap = {};
    ws.model.merges?.forEach(m => {
        const [start, end] = m.split(':');
        const s = decodeAddr(start);
        const e = decodeAddr(end);

        mergeMap[start] = {
            colSpan: e.c - s.c + 1,
            rowSpan: e.r - s.r + 1
        };
    });

    // =========================
    // ✅ RENDER TABLE
    // =========================
    ws.eachRow({ includeEmpty: true }, (row, r) => {

        const tr = document.createElement('tr');
        tr.style.height = rowHeights[r - 1] + 'px';

        row.eachCell({ includeEmpty: true }, (cell, c) => {

            if (cell.isMerged && cell.address !== cell.master.address) return;

            const td = document.createElement('td');

            td.style.width = colWidths[c - 1] + 'px';
            td.style.padding = '0';
            td.style.margin = '0';
            td.style.boxSizing = 'border-box';
            td.style.overflow = 'hidden';
            td.style.lineHeight = '1';
            td.style.height = rowHeights[r - 1] + 'px';
            td.textContent = cell.text || '';

            const merge = mergeMap[cell.address];
            if (merge) {
                td.colSpan = merge.colSpan;
                td.rowSpan = merge.rowSpan;
            }

            //   td.className = getClass(cell); // 🔥 ใช้ cache style
              applyCellStyle(td, cell, themeColors);

            tr.appendChild(td);
        });

        table.appendChild(tr);
    });

    container.appendChild(table);

    // =========================
    // 🔥 IMAGE LAYER (เตรียมไว้)
    // =========================
    const imageLayer = document.createElement('div');
    imageLayer.style.position = 'absolute';
    imageLayer.style.left = '0';
    imageLayer.style.top = '0';
    imageLayer.style.pointerEvents = 'none';
    imageLayer.style.width = x + 'px';
    imageLayer.style.height = y + 'px';

    container.appendChild(imageLayer);
    renderImages(images, workbook, imageLayer, colOffsets, rowOffsets);

    // 👉 return สำหรับใช้ต่อ render image
    return {
        colWidths,
        rowHeights,
        colOffsets,
        rowOffsets,
        imageLayer
    };
}

function applyCellStyle(td, cell, themeColors) {

    const argbFont = toARGB(cell.font?.color, themeColors);
    const argbBg = toARGB(cell.fill?.fgColor, themeColors);

    // =========================
    // FONT
    // =========================
    if (cell.font) {
        if (cell.font.bold) td.style.fontWeight = 'bold';
        if (cell.font.italic) td.style.fontStyle = 'italic';

        if (cell.font.size) {
            td.style.fontSize = (cell.font.size * 96 / 72) + 'px'; // 🔥 ใช้ px เหมือนกัน
        }

        if (cell.font.name) {
            td.style.fontFamily = cell.font.name;
        }

        if (cell.font.color) {
            td.style.color = argbToHex(argbFont);
        }
    }

    // =========================
    // BACKGROUND
    // =========================
    if (cell.fill && cell.fill.fgColor) {
        td.style.background = argbToHex(argbBg); // 🔥 ใช้ background เหมือนกัน
    }

    // =========================
    // ALIGNMENT
    // =========================
    td.style.textAlign = 'right';
    td.style.verticalAlign = 'bottom';

    if (cell.alignment) {
        if (cell.alignment.horizontal) {
            td.style.textAlign = cell.alignment.horizontal;
        }

        if (cell.alignment.vertical) {
            td.style.verticalAlign = cell.alignment.vertical; // 🔥 ตรงกัน
        }

        if (cell.alignment?.indent) {
            // td.style.textIndent = cell.alignment.indent + 'px';
            td.style.paddingLeft = (cell.alignment.indent * 10) + 'px';
        }
    }

    // =========================
    // BORDER
    // =========================
    if (cell.border) {
        const b = cell.border;

        if (b.top) td.style.borderTop = getBorderStyle(b.top, themeColors);
        if (b.bottom) td.style.borderBottom = getBorderStyle(b.bottom, themeColors);
        if (b.left) td.style.borderLeft = getBorderStyle(b.left, themeColors);
        if (b.right) td.style.borderRight = getBorderStyle(b.right, themeColors);

        // =========================
        // DIAGONAL BORDER (เส้นทแยง)
        // =========================
        if (b.diagonal && b.diagonal.style) {
            td.style.position = 'relative';
            const diagColor = argbToHex(toARGB(b.diagonal.color, themeColors)) || '#000000';

            // Map border style to width (same as getBorderStyle)
            const widthMap = { thin: '1px', medium: '2px', thick: '3px', hair: '1px', double: '3px', dashed: '1px', dotted: '1px' };
            const lineWidth = parseFloat(widthMap[b.diagonal.style] || '1px');
            const halfWidth = lineWidth / 2;

            if (b.diagonalDown) {
                const line = document.createElement('div');
                line.style.cssText = `position:absolute;top:0;left:0;right:0;bottom:0;pointer-events:none;z-index:-1;background:linear-gradient(to bottom right, transparent calc(50% - ${halfWidth}px), ${diagColor} calc(50% - ${halfWidth}px), ${diagColor} calc(50% + ${halfWidth}px), transparent calc(50% + ${halfWidth}px))`;
                td.appendChild(line);
            }

            if (b.diagonalUp) {
                const line = document.createElement('div');
                line.style.cssText = `position:absolute;top:0;left:0;right:0;bottom:0;pointer-events:none;z-index:-1;background:linear-gradient(to top right, transparent calc(50% - ${halfWidth}px), ${diagColor} calc(50% - ${halfWidth}px), ${diagColor} calc(50% + ${halfWidth}px), transparent calc(50% + ${halfWidth}px))`;
                td.appendChild(line);
            }
        }
    }
}

function getBorderStyle(borderSide, themeColors = []) {
    if (!borderSide || !borderSide.style) return '1px solid #d1d1d1';

    const widthMap = {
        thin: '1px', medium: '2px', thick: '3px', hair: '1px',
        double: '3px', dashed: '1px', dotted: '1px'
    };

    const width = widthMap[borderSide.style] || '1px';
    const type = (borderSide.style === 'dashed' || borderSide.style === 'dotted')
        ? borderSide.style
        : 'solid';

    const argb = toARGB(borderSide.color, themeColors);
    const color = argb ? argbToHex(argb) : '#000000';

    return `${width} ${type} ${color}`;
}

function normalizeBorders(ws) {
    ws.eachRow({ includeEmpty: true }, (row, r) => {
        row.eachCell({ includeEmpty: true }, (cell, c) => {

            if (!cell.border) return;

            // LEFT → ไป set RIGHT ของ cell ซ้าย
            if (cell.border.left && c > 1) {
                const leftCell = row.getCell(c - 1);
                leftCell.border = leftCell.border || {};
                leftCell.border.right = cell.border.left;
            }

            // TOP → ไป set BOTTOM ของ cell บน
            if (cell.border.top && r > 1) {
                const topCell = ws.getRow(r - 1).getCell(c);
                topCell.border = topCell.border || {};
                topCell.border.bottom = cell.border.top;
            }
        });
    });
}


function anchorToPx(anchor, colOffsets, rowOffsets, colWidths, rowHeights) {
    const EMU_PER_PIXEL = 9525;
    const col = anchor.nativeCol;
    const row = anchor.nativeRow;

    const colOffsetPx = anchor.nativeColOff / EMU_PER_PIXEL;
    const rowOffsetPx = anchor.nativeRowOff / EMU_PER_PIXEL;

    const x = colOffsets[col] + colOffsetPx;
    const y = rowOffsets[row] + rowOffsetPx;

    return { x, y };
}

// Track created object URLs to revoke them later
let _imageUrls = [];

function renderImages(
    images,
    workbook,
    imageLayer,
    colOffsets,
    rowOffsets,
    colWidths,
    rowHeights
) {
    // Revoke previous image URLs to free memory
    _imageUrls.forEach(url => URL.revokeObjectURL(url));
    _imageUrls = [];

    images.forEach(img => {
        const image = workbook.getImage(img.imageId);
        if (!image?.buffer) return;

        const tl = anchorToPx(
            img.range.tl,
            colOffsets,
            rowOffsets,
            colWidths,
            rowHeights
        );

        const br = anchorToPx(
            img.range.br,
            colOffsets,
            rowOffsets,
            colWidths,
            rowHeights
        );

        const width = br.x - tl.x;
        const height = br.y - tl.y;

        const imgEl = document.createElement('img');

        const blob = new Blob(
            [image.buffer],
            { type: `image/${image.extension}` }
        );

        const url = URL.createObjectURL(blob);
        _imageUrls.push(url);
        imgEl.src = url;

        imgEl.style.position = 'absolute';
        imgEl.style.left = tl.x + 'px';
        imgEl.style.top = tl.y + 'px';
        imgEl.style.width = width + 'px';
        imgEl.style.height = height + 'px';
        imgEl.style.imageRendering = 'auto';

        imageLayer.appendChild(imgEl);
    });
}