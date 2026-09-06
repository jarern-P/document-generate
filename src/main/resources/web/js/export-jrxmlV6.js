function exportToJrxmlV6({ ws: wsJson, workbook: wbJson, meta, images = [] }) {
    const ws = buildWorksheetFromJson(wsJson);
    const workbook = buildWorkbookFromJson(wbJson);
    // const ws = wsJson;
    // const workbook = wbJson;
let xml = [];
const themeColors = getThemeColors(workbook);
const A4_PORTRAIT_WIDTH = 595;
const A4_PORTRAIT_HEIGHT = 842;
const A4_LANDSCAPE_WIDTH = 842;
const A4_LANDSCAPE_HEIGHT = 595;
const isSnapToGrid = true;
let totalWidthPt = 0;
for (let i = 1; i <= ws.columnCount; i++) {

    const col = ws.getColumn(i);

    totalWidthPt += getColWidthPt(col.width);
}

const leftMargin = 20;
const rightMargin = 20;
const topMargin = 20;
const bottomMargin = 20;

const portraitUsableWidth = 520;

const isLandscape =
    totalWidthPt > portraitUsableWidth;

const pageWidth =
    isLandscape
        ? A4_LANDSCAPE_WIDTH
        : A4_PORTRAIT_WIDTH;

const pageHeight =
    isLandscape
        ? A4_LANDSCAPE_HEIGHT
        : A4_PORTRAIT_HEIGHT;

const orientation =
    isLandscape ? 'Landscape' : 'Portrait';

const columnWidth =
    pageWidth - leftMargin - rightMargin;

// const scaleX =
//     columnWidth / totalWidthPt;
const scaleX =
    totalWidthPt > 0
        ? columnWidth / totalWidthPt
        : 1;
     
const scaleY = 1;

const parameters =
    collectParameters(ws);
    let xmlParametersArray = [];

parameters.forEach(p => {

    xmlParametersArray.push(`
<parameter
    name="${p.name}"
    class="${p.type}"
/>
`);
});

let xmlParameters =
    xmlParametersArray.join('\n');

const fields = collectFields(ws);
let xmlFieldsArray = [];
fields.forEach(f => {

    xmlFieldsArray.push(`
<field
    name="${f.name}"
    class="${f.type}"
/>
`);
});
let xmlFields = xmlFieldsArray.join('\n');

  const sections = buildSections(meta, ws.rowCount || ws.lastRow?.number || 0);

  let titleHeight = 0;
  let pageHeaderHeight = 0;
  let columnHeaderHeight = 0;
  let detailHeight = 0;
  let columnFooterHeight = 0;
  let pageFooterHeight = 0;
  let summaryHeight = 0;

  sections.forEach(section => {

    let bandHeight = 0;

    for (
        let r = section.start;
        r <= section.end;
        r++
    ) {

        const row = ws.getRow(r);
        bandHeight +=
            getRowHeightPt(row.height, ws.properties.defaultRowHeight);
        
    }

     section.height = bandHeight;
});

 // ===== Use canonical section names (normalized by buildSections) =====
 titleHeight = getSectionHeight(sections, ws, 'title');
 pageHeaderHeight = getSectionHeight(sections, ws, 'pageheader');
 columnHeaderHeight = getSectionHeight(sections, ws, 'columnheader');
 detailHeight = getSectionHeight(sections, ws, 'detail');
 columnFooterHeight = getSectionHeight(sections, ws, 'columnfooter');
 pageFooterHeight = getSectionHeight(sections, ws, 'pagefooter');
 summaryHeight = getSectionHeight(sections, ws, 'summary');

 let xmlTitle = buildBand(ws, sections.find(x => x.key === 'title'), scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '';
 let xmlColumnHeader = buildBand(ws, sections.find(x => x.key === 'columnheader'), scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '';
 let xmlPageHeader = buildBand(ws, sections.find(x => x.key === 'pageheader'), scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '';
 let xmlDetail = buildBand(ws, sections.find(x => x.key === 'detail'), scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '';
 let xmlColumnFooter = buildBand(ws, sections.find(x => x.key === 'columnfooter'), scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '';
 let xmlPageFooter = buildBand(ws, sections.find(x => x.key === 'pagefooter'), scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '';
 let xmlSummary = buildBand(ws, sections.find(x => x.key === 'summary'), scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '';

 // ============================================================
 // GROUPS — Build <group> XML elements from meta definitions
 // ============================================================
 const groups = buildGroups(meta);
 let xmlGroups = '';

 if (groups && groups.length > 0) {
   const groupXmlParts = [];

   groups.forEach(group => {
     // Find the section for group header (matches by key + start row)
     const headerSection = sections.find(
       s => s.key === 'groupheader' && s.start === group.headerRow
     );
     const footerSection = sections.find(
       s => s.key === 'groupfooter' && s.start === group.footerRow
     );

     // Build group header band content
     // Use headerSection.height directly (already computed in the height loop above)
     // Do NOT use getSectionHeight() — it uses sections.find() which only returns the FIRST match
     const headerHeight = headerSection?.height ?? 0;
     const xmlGroupHeader = headerSection
       ? (buildBand(ws, headerSection, scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '')
       : '';

     // Build group footer band content
     const footerHeight = footerSection?.height ?? 0;
     const xmlGroupFooter = footerSection
       ? (buildBand(ws, footerSection, scaleX, scaleY, themeColors, images, workbook, fields, isSnapToGrid) || '')
       : '';

     const groupName = group.name;

     groupXmlParts.push(`
<group name="${escapeXml(groupName)}">
    <groupExpression><![CDATA[$F{${escapeXml(groupName)}}]]></groupExpression>

    <groupHeader>
        <band height="${headerHeight}">
        ${xmlGroupHeader}
        </band>
    </groupHeader>

    <groupFooter>
        <band height="${footerHeight}">
        ${xmlGroupFooter}
        </band>
    </groupFooter>

</group>
`);
   });

   xmlGroups = groupXmlParts.join('\n');
 }

    xml.push(`<?xml version="1.0" encoding="UTF-8"?>`);

    xml.push(`
<jasperReport
    xmlns="http://jasperreports.sourceforge.net/jasperreports"
    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
    xsi:schemaLocation="
        http://jasperreports.sourceforge.net/jasperreports
        http://jasperreports.sourceforge.net/xsd/jasperreport.xsd"
    name="report"
    orientation="${orientation}"
    pageWidth="${pageWidth}"
    pageHeight="${pageHeight}"
    columnWidth="${columnWidth}"
    leftMargin="20"
    rightMargin="20"
    topMargin="20"
    bottomMargin="20"
>
<property
    name="net.sf.jasperreports.export.xls.remove.empty.space.between.columns"
    value="true"/>

<property
    name="net.sf.jasperreports.export.xls.white.page.background"
    value="false"/>

<property
    name="net.sf.jasperreports.export.xls.detect.cell.type"
    value="true"/>

<property
    name="net.sf.jasperreports.export.xls.wrap.text"
    value="false"/>

        ${xmlParameters}
        ${xmlFields}
        ${xmlGroups}
    <title>
        <band height="${titleHeight}">
        ${xmlTitle}
        </band>
    </title>

    <pageHeader>
        <band height="${pageHeaderHeight}">
        </band>
    </pageHeader>

    <columnHeader>
        <band height="${columnHeaderHeight}">
        ${xmlColumnHeader}
        </band>
    </columnHeader>

    <detail>
        <band height="${detailHeight}">
        ${xmlDetail}
        </band>
    </detail>

    <columnFooter>
        <band height="${columnFooterHeight}">
        ${xmlColumnFooter}
        </band>
    </columnFooter>

    <pageFooter>
        <band height="${pageFooterHeight}">
        ${xmlPageFooter}
        </band>
    </pageFooter>

    <summary>
        <band height="${summaryHeight}">
        ${xmlSummary}
        </band>
    </summary>
`);

    xml.push(`</jasperReport>`);

    return xml.join('\n');
}


function getRowHeightPt(height, defaultRowHeight) {
    return height || defaultRowHeight || 15;
}

function getSectionHeight(sections, ws, ...keys) {

    return Math.round(
        sections.find(x =>
            keys.includes(
                x.key?.toLowerCase()
            )
        )?.height || 0
    );
}

if (typeof escapeXml !== 'function') {
  function escapeXml(value) {
    if (value === null || value === undefined) return '';
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
  }
}


function buildBand(ws, section, scaleX = 1, scaleY = 1, themeColors, images = [], workbook, fields  = [], isSnapToGrid = false) {
   let xml = [];
    if (!section || section.key.toLowerCase() === 'end') {
        return '';
    }

        const gridCols =
        isSnapToGrid
            ? buildGridColumns(ws, scaleX, 5)
            : null;

    const gridRows =
        isSnapToGrid
            ? buildGridRows(ws, scaleY, 1)
            : null;

for (
    let r = section.start;
    r <= section.end;
    r++
) {

    const row = ws.getRow(r);

    for (
        let colNumber = 1;
        colNumber <= ws.columnCount;
        colNumber++
    ) {

        const cell =
            row.getCell(colNumber);

        if (
            cell.isMerged &&
            cell.address !== cell.master.address
        ) {
            continue;
        }
        
        // const value =
        //     cell.value != null
        //         ? String(cell.value)
        //         : '';

        const value = getCellDisplayValue(cell);
        // const fieldName = parseField(value);
        const expression =
            parseExpression(value);


        const isField =
            expression?.type === 'field';

        const isParameter =
            expression?.type === 'parameter';
        
        const isMixed =
            expression?.type === 'mixed';

    // const isField =
    //     !!fieldName;
        const fieldName =
            expression?.name;

        let x =
            getColumnX(ws, colNumber) * scaleX;

        let width =
            getMergedWidth(ws, cell) * scaleX;

        let y =
            getRowY(ws, section.start, r) * scaleY;

        let height =
            getMergedHeight(ws, cell) * scaleY;

            if (isSnapToGrid) {

                let startCol = colNumber;
                let endCol = colNumber;

                let startRow = r;
                let endRow = r;

                if (cell.isMerged) {

                    const range =
                        ws._merges[cell.master.address];

                    if (range) {

                        startCol = range.left;
                        endCol = range.right;

                        startRow = range.top;
                        endRow = range.bottom;
                    }
                }

                x =
                    gridCols[startCol - 1];

                width =
                    gridCols[endCol] -
                    gridCols[startCol - 1];


                const baseY =
                    gridRows[section.start - 1];

                y =
                    gridRows[startRow - 1] -
                    baseY;

                height =
                    gridRows[endRow] -
                    gridRows[startRow - 1];
            }

        const horizontal =
            convertHorizontalAlign(
                cell.alignment?.horizontal
            );

        const vertical =
            convertVerticalAlign(
                cell.alignment?.vertical
            );

        const fontName =
            cell.font?.name || 'SansSerif';

        const fontSize =
            adjustFontSize(
                cell.font?.size || 10,
                fontName
            );

        const isBold =
            cell.font?.bold ? 'true' : 'false';

        const isItalic =
            cell.font?.italic ? 'true' : 'false';

        const isUnderline =
            cell.font?.underline ? 'true' : 'false';

     const indent =
    (cell.alignment?.indent || 0) * 10; 

        const paragraphIndent =
            indent > 0
                ? `<paragraph leftIndent="${indent}"/>`
        : '';

        const argbFont =
            toARGB(
                cell.font?.color,
                themeColors
            );
        

        const fontColorHex =
            argbToHex(argbFont);
           
        const fontColor =
            fontColorHex
                ? `forecolor="${fontColorHex}"`
                : '';

        const argbBg =
            toARGB(
                cell.fill?.fgColor,
                themeColors
            );

        const fillColor =
            argbToHex(argbBg) === ''
                ? ''
                : `backcolor="${argbToHex(argbBg)}"`;

        const mode =
            fillColor
                ? 'Opaque'
                : 'Transparent';

        const borderXml =
            buildBorderXml(cell.border);

        // ===== IMPORTANT =====
        // ไม่มี text ไม่มี border → ข้าม
        if (
            value === '' &&
            borderXml === '' &&
            fillColor === ''
        ) {
            continue;
        }

const field =
    fields?.find(
        x => x.name === fieldName
    );

const patternAttr = 
    field?.pattern
        ? `pattern="${field.pattern}"`
        : '';

const noteText =
    getCellNoteText(cell);

const rpFormula =
    parseRPPRINTIF(noteText);

const printWhenXml =
    rpFormula?.type === 'printWhen'
        ? `>
        <printWhenExpression>${convertToJasperExpression(rpFormula.expression)}</printWhenExpression> 
        </reportElement>`
        : '/>';
let expressionText = '';

if (rpFormula?.type === 'textExpression') {
    // RP() overrides the cell expression entirely
    expressionText = convertToJasperExpression(rpFormula.expression);
}
else if (isParameter) {
    expressionText = `$P{${fieldName}}`;
}
else if (isField) {
    expressionText = `$F{${fieldName}}`;
}
else if (isMixed) {
    expressionText = expression.expression;
}

if (isField || isParameter || isMixed || rpFormula?.type === 'textExpression') {

    xml.push(`
<textField ${patternAttr}>
    <reportElement
        stretchType="RelativeToTallestObject"
        x="${Math.round(x)}"
        y="${Math.round(y)}"
        width="${Math.round(width)}"
        height="${Math.round(height)}"
        mode="${mode}"
        ${fillColor}
        ${fontColor} 
        ${printWhenXml}
    
    ${borderXml}

    <textElement
        textAlignment="${horizontal}"
        verticalAlignment="${vertical}"
    >
        <font
            fontName="${fontName}"
            size="${Math.round(fontSize)}"
            isBold="${isBold}"
            isItalic="${isItalic}"
            isUnderline="${isUnderline}"
        />
        ${paragraphIndent}
    </textElement>

<textFieldExpression><![CDATA[
    ${expressionText}
]]></textFieldExpression>

</textField>
`);
}
else {

    xml.push(`
<staticText>
    <reportElement
        stretchType="RelativeToTallestObject"
        x="${Math.round(x)}"
        y="${Math.round(y)}"
        width="${Math.round(width)}"
        height="${Math.round(height)}"
        mode="${mode}"
        ${fillColor}
        ${fontColor} 
        ${printWhenXml}
    
    ${borderXml}

    <textElement
        textAlignment="${horizontal}"
        verticalAlignment="${vertical}"
    >
        <font
            fontName="${fontName}"
            size="${Math.round(fontSize)}"
            isBold="${isBold}"
            isItalic="${isItalic}"
            isUnderline="${isUnderline}"
        />
        ${paragraphIndent}
    </textElement>

    <text><![CDATA[${value}]]></text>

</staticText>
`);
}
    
}


    }
    // return xml.join('\n');
    // ===== IMAGES =====
xml.push(
    buildImages(
        ws,
        section,
        images,
        workbook,
        scaleX,
        scaleY,
        isSnapToGrid
    )
);

return xml.join('\n');
}

function getColumnX(ws, colNumber) {

    let x = 0;

    for (let c = 1; c < colNumber; c++) {

        x += getColWidthPt(
            ws.getColumn(c).width
        );
    }

    return x;
    // return Math.round(x);
}

function getRowY(ws, startRow, currentRow) {

    let y = 0;

    for (
        let r = startRow;
        r < currentRow;
        r++
    ) {

        y += getRowHeightPt(
            ws.getRow(r).height, ws.properties.defaultRowHeight
        );
    }

    return y;
}

function getMergedWidth(ws, cell) {

    if (!cell.isMerged) {

        return Math.round(getColWidthPt(
            ws.getColumn(cell.col).width
        ));
    }

    const master = cell.master;

    const range =
        ws._merges[master.address];

    if (!range) {

        return Math.round(getColWidthPt(
            ws.getColumn(cell.col).width
        ));
    }

    let width = 0;

    for (
        let c = range.left;
        c <= range.right;
        c++
    ) {

        width += getColWidthPt(
            ws.getColumn(c).width
        );
    }

    return width;
    // return Math.round(width);
}

function getMergedHeight(ws, cell) {

    if (!cell.isMerged) {

        return getRowHeightPt(
            ws.getRow(cell.row).height, ws.properties.defaultRowHeight
        );
    }

    const master = cell.master;

    const range =
        ws._merges[master.address];

    if (!range) {

        return getRowHeightPt(
            ws.getRow(cell.row).height, ws.properties.defaultRowHeight
        );
    }

    let height = 0;

    for (
        let r = range.top;
        r <= range.bottom;
        r++
    ) {

        height += getRowHeightPt(
            ws.getRow(r).height, ws.properties.defaultRowHeight
        );
    }

    return height;
}

function convertHorizontalAlign(align) {

    switch ((align || '').toLowerCase()) {

        case 'center':
            return 'Center';

        case 'right':
            return 'Right';

        case 'justify':
            return 'Justified';

        default:
            return 'Left';
    }
}

function convertVerticalAlign(align) {

    switch ((align || '').toLowerCase()) {

        case 'middle':
            return 'Middle';

        case 'top':
            return 'Top';

        default:
            return 'Bottom';
    }
}

function getFontScale(fontName) {

    const name =
        (fontName || '').toLowerCase();

    if (
        name.includes('sarabun') ||
        name.includes('angsana') ||
        name.includes('cordia')
    ) {
        return 0.62;
    }

    if (
        name.includes('calibri')
    ) {
        return 0.75;
    }

    // return 0.70;
}

function adjustFontSize(size, fontName) {
         return size; 
        //  ที่ใช้แบบนี้เพราะว่า เิม scale เพื่อให้ อักษรไม่เกินขอบ แต่ตอนนี้ใช้ไปคลุมแถวแล้วกดเลือกขนาด font ระบบจะขยายแถวให้เอง 

    const name =
        (fontName || '').toLowerCase();

    // ===== THAI =====
    if (
        name.includes('sarabun') ||
        name.includes('angsana') ||
        name.includes('cordia')
    ) {

        if (size >= 18) {
            return size - 5;
        }

        if (size >= 14) {
            return size - 4;
        }

        return size - 3;
    }

    // ===== CALIBRI =====
    if (name.includes('calibri')) {

        if (size >= 16) {
            return size - 2;
        }

        return size - 1;
    }

    return size;
}

function buildBorderXml(border) {

    if (!border || Object.keys(border).length === 0) {
        return '';
    }

    function buildPenXml(sideName) {

        const side = border?.[sideName];

        if (!side?.style) {
            return '';
        }

        // ===== COLOR =====
        let color = '#000000';

        if (side.color?.argb) {

            // remove alpha
            color = '#' + side.color.argb.substring(2);
        }

        // ===== STYLE =====

        let lineStyle = 'Solid';

        switch (side.style) {

            case 'dashed':
                lineStyle = 'Dashed';
                break;

            case 'dotted':
                lineStyle = 'Dotted';
                break;

            case 'double':
                lineStyle = 'Double';
                break;

            default:
                lineStyle = 'Solid';
                break;
        }

        // ===== WIDTH =====

        let lineWidth = 0.5;

        switch (side.style) {

            case 'medium':
                lineWidth = 1;
                break;

            case 'thick':
                lineWidth = 2;
                break;

            default:
                lineWidth = 0.5;
                break;
        }

        return `
        <${sideName}Pen
            lineWidth="${lineWidth}"
            lineStyle="${lineStyle}"
            lineColor="${color}"
        />
        `;
    }

    return `
    <box>
        ${buildPenXml('top')}
        ${buildPenXml('left')}
        ${buildPenXml('bottom')}
        ${buildPenXml('right')}
    </box>
    `;
}


function buildImages(
    ws,
    section,
    images,
    workbook,
    scaleX = 1,
    scaleY = 1,
    isSnapToGrid = false
) {

    if (!images || images.length === 0) {
        return '';
    }

    const gridCols =
    isSnapToGrid
        ? buildGridColumns(ws, scaleX, 5)
        : null;

const gridRows =
    isSnapToGrid
        ? buildGridRows(ws, scaleY, 1)
        : null;

    let xml = [];

    images.forEach((img, idx) => {

        const tl = img.range?.tl;
        const br = img.range?.br;

        if (!tl || !br) {
            return;
        }

        // ===== CHECK IMAGE IN SECTION =====

        const row1 = tl.nativeRow + 1;
        const row2 = br.nativeRow + 1;

        const isInSection =
            row1 >= section.start &&
            row2 <= section.end;

        if (!isInSection) {
            return;
        }

        // ===== POSITION =====

        // const bounds = getImageBoundsPt(ws, img, scaleX, scaleY, section.start);

        // const x = bounds.x;
        // const y = bounds.y;
        // const width = bounds.width;
        // const height = bounds.height;
        let { x, y, width, height } =
    getImageBoundsPt(ws, img, scaleX, scaleY, section.start);

        // ===== SNAP TO GRID =====
        // if (isSnapToGrid) {

        //     const col = img.range.tl.nativeCol + 1;
        //     const row = img.range.tl.nativeRow + 1;

        //     const baseY = gridRows[section.start - 1];

        //     const snappedX = gridCols[col - 1];
        //     const snappedY = gridRows[row - 1] - baseY;

        //     // 👉 ใช้ offset เดิม (สำคัญมาก)
        //     x = snappedX + (x - Math.floor(x));
        //     y = snappedY + (y - Math.floor(y));
        // }
        if (isSnapToGrid) {

            let startCol = tl.nativeCol + 1;
            let endCol = br.nativeCol + 1;

            let startRow = tl.nativeRow + 1;
            let endRow = br.nativeRow + 1;

            // ===== SNAP แบบเต็ม cell =====
            x = gridCols[startCol - 1];

            width =
                gridCols[endCol] -
                gridCols[startCol - 1];

            const baseY = gridRows[section.start - 1];

            y =
                gridRows[startRow - 1] -
                baseY;

            height =
                gridRows[endRow] -
                gridRows[startRow - 1];
        }

        // ===== IMAGE DATA =====

        const image =
            workbook.getImage(img.imageId);

        if (!image?.buffer) {
            return;
        }

        const base64 =
            toBase64(image.buffer);

        // const imageType =
        //     image.extension?.toLowerCase() === 'jpg'
        //         ? 'jpeg'
        //         : image.extension;
    // scaleImage:   FillFrame, RetainShape
        xml.push(`
        <image scaleImage="FillFrame" onErrorType="Blank">
            <reportElement
                x="${Math.round(x)}"
                y="${Math.round(y)}"
                width="${Math.round(width)}"
                height="${Math.round(height)}"
                uuid="${uuidv4()}"
            />

        <imageExpression class="java.io.InputStream"><![CDATA[
        new java.io.ByteArrayInputStream(
            java.util.Base64.getDecoder().decode(
                        "${base64}"
                    )
                )
            ]]></imageExpression>

        </image>
        `);
    });

    return xml.join('\n');
}

function uuidv4() {

    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'
        .replace(/[xy]/g, c => {

            const r = Math.random() * 16 | 0;
            const v = c === 'x'
                ? r
                : (r & 0x3 | 0x8);

            return v.toString(16);
        });
}

function getImageBoundsPt(
    ws,
    img,
    scaleX = 1,
    scaleY = 1,
    sectionStart = 1
) {

    const EMU_PER_PIXEL = 9525;

    const tl = img.range.tl;
    const br = img.range.br;

    // ===== START POSITION =====

    let x =
        getColumnX(ws, tl.nativeCol + 1);

    let y =
        getRowY(ws, sectionStart, tl.nativeRow + 1);

    // ===== COLUMN OFFSET =====

    x += (
        (tl.nativeColOff || 0)
        / EMU_PER_PIXEL
    ) * 0.75; // px -> pt

    // ===== ROW OFFSET =====

    y += (
        (tl.nativeRowOff || 0)
        / EMU_PER_PIXEL
    ) * 0.75;

    // ===== END POSITION =====

    let x2 =
        getColumnX(ws, br.nativeCol + 1);

    let y2 =
        getRowY(ws, sectionStart, br.nativeRow + 1);

    x2 += (
        (br.nativeColOff || 0)
        / EMU_PER_PIXEL
    ) * 0.75;

    y2 += (
        (br.nativeRowOff || 0)
        / EMU_PER_PIXEL
    ) * 0.75;

    // ===== SCALE =====

    x *= scaleX;
    y *= scaleY;
    x2 *= scaleX;
    y2 *= scaleY;

    return {
        x,
        y,
        width: x2 - x,
        height: y2 - y
    };
}

function parseExpression(value) {

    if (!value) {
        return null;
    }

    const str =
        String(value).trim();

    // ===== PARAMETER ONLY =====

    let match =
        str.match(/^[pP]\{\{(.*?)\}\}$/);

    if (match) {

        return {
            type: 'parameter',
            name: match[1].trim()
        };
    }

    // ===== FIELD ONLY =====

    match =
        str.match(/^\{\{(.*?)\}\}$/);

    if (match) {

        return {
            type: 'field',
            name: match[1].trim()
        };
    }

    // ===== MIXED EXPRESSION =====

    match =
        str.match(/([pP]?\{\{.*?\}\})/);

    if (match) {
        return {
            type: 'mixed',
            expression: convertToJasperExpression(str)
        };
    }

    return null;
}
// ===== Shared RPPRINTIF utilities moved to utils-cal.js =====
// getCellNoteText(), parseRPPRINTIF(), convertToJasperExpression() are now in utils-cal.js



function inferFieldType(cell) {

    const numFmt =
        (
            cell.numFmt ||
            cell.style?.numFmt ||
            'General'
        )
        .trim()
        .toLowerCase();

    // ===== TEXT =====

    if (
        numFmt === '@'
    ) {
        return 'java.lang.String';
    }

    if (
            typeof cell.value === 'number'
        ) {

            if (Number.isInteger(cell.value)) {
                return 'java.lang.Integer';
            }

            return 'java.math.BigDecimal';
        }

    // ===== DATE =====
    // if (
    //     numFmt.includes('yy') ||
    //     numFmt.includes('dd') ||
    //     numFmt.includes('mm')
    // ) {
    //     return 'java.util.Date';
    // }
const isDateFormat =
    /(yy|yyyy|dd|mm|mmm)/i
        .test(numFmt);

    if (
        isDateFormat
    ) {
        return 'java.util.Date';
    }

    // ===== TIME =====
    if (
        numFmt.includes('hh') ||
        numFmt.includes('ss')
    ) {
        return 'java.sql.Time';
    }

    // ===== PERCENT =====
    if (
        numFmt.includes('%')
    ) {
        return 'java.math.BigDecimal';
    }

    // ===== DECIMAL =====
    if (
        numFmt.includes('0.00') ||
        numFmt.includes('#,##0.00') ||
        numFmt.includes('#.##')
    ) {
        return 'java.math.BigDecimal';
    }

    // ===== INTEGER =====
    if (
        numFmt === '0' ||
        numFmt === '#,##0'
    ) {
        return 'java.lang.Integer';
    }

    // ===== CURRENCY =====
    if (
        numFmt.includes('$') ||
        numFmt.includes('฿')
    ) {
        return 'java.math.BigDecimal';
    }

    // ===== DEFAULT =====
    return 'java.lang.String';
}

function mapExcelFormatToJasper(numFmt) {

    if (!numFmt) {
        return '';
    }

    const fmt =
        numFmt.toLowerCase();

    if (
        fmt.includes('yy') ||
        fmt.includes('dd')
    ) {
        return 'dd/MM/yyyy';
    }

    if (
        fmt.includes('#,##0.00')
    ) {
        return '#,##0.00';
    }

    if (
        fmt.includes('#,##0')
    ) {
        return '#,##0';
    }

    if (
        fmt.includes('0%')
    ) {
        return '#0%';
    }

    return '';
}
function collectFields(ws) {

    const fields = [];

    ws.eachRow(row => {

        row.eachCell(cell => {

            // const fieldName =
            //     parseField(cell.value);
            const value =
            getCellDisplayValue(cell);

        // const fieldName =
        //     parseField(value);

            const expression =
                parseExpression(value);

            if (
                !expression ||
                expression.type !== 'field'
            ) {
                return;
            }

            const exists =
                fields.find(
                    x => x.name === expression.name
                );

            if (exists) {
                return;
            }

            fields.push({
                name: expression.name,
                type: inferFieldType(cell),
                pattern: mapExcelFormatToJasper(
                    cell.numFmt
                )
            });
        });
    });

    return fields;
}

function collectParameters(ws) {

    const params = [];

    ws.eachRow(row => {

        row.eachCell(cell => {

            const value =
                getCellDisplayValue(cell);

            const expression =
                parseExpression(value);

            if (
                !expression ||
                expression.type !== 'parameter'
            ) {
                return;
            }

            const exists =
                params.find(
                    x => x.name === expression.name
                );

            if (exists) {
                return;
            }

            params.push({
                name: expression.name,
                type: inferFieldType(cell)
            });
        });
    });

    return params;
}

function getCellDisplayValue(cell) {
    if (
        cell.value == null
    ) {
        return '';
    }

    // ===== STRING =====

    if (
        typeof cell.value === 'string'
    ) {
        return cell.value;
    }

    // ===== NUMBER =====

    if (
        typeof cell.value === 'number'
    ) {
        return String(cell.value);
    }

    // ===== BOOLEAN =====

    if (
        typeof cell.value === 'boolean'
    ) {
        return String(cell.value);
    }

    // ===== DATE =====

    if (
        cell.value instanceof Date
    ) {
        // return cell.value.toISOString();
        return cell.value;
    }

    // ===== FORMULA =====

    if (
        typeof cell.value === 'object' &&
        (
            cell.value?.formula ||
            cell.value?.sharedFormula
        )
    ) {
        return String(
            cell.value.result ?? ''
            // cell.value.model?.formula ?? ''
        );
    }

    // ===== RICHTEXT =====

    if (
        typeof cell.value === 'object' &&
        Array.isArray(cell.value.richText)
    ) {

        return cell.value.richText
            .map(x => x.text || '')
            .join('');
    }

    // ===== HYPERLINK =====

    if (
        typeof cell.value === 'object' &&
        cell.value?.text
    ) {
        return cell.value.text;
    }

    return String(cell.value);
}


function buildGridColumns(ws, scaleX = 1, grid = 5) {
    // เดิม grid เป็น 5 แต่ลอง hardcode ให้เป็น 2 เพื่อให้ snap grid ละเอียดขึ้น
    grid = 2;
    const positions = [0];

    let currentX = 0;

    for (let c = 1; c <= ws.columnCount; c++) {

        const width =
            getColWidthPt(
                ws.getColumn(c).width
            ) * scaleX;

        // currentX += width;

        // positions[c] =
        //     Math.round(currentX / grid) * grid;
        const snappedWidth =
            Math.round(width / grid) * grid;

        currentX += snappedWidth;

        positions[c] = currentX;
    }

    return positions;
}


function buildGridRows(
    ws,
    scaleY = 1,
    grid = 1
) {

    const positions = [0];

    let currentY = 0;

    const rowCount =
        ws.rowCount ||
        ws.lastRow?.number ||
        0;

    for (let r = 1; r <= rowCount; r++) {

        const row =
            ws.getRow(r);

        const height =
            getRowHeightPt(
                row.height,
                ws.properties.defaultRowHeight
            ) * scaleY;

        currentY += height;

        positions[r] =
            Math.round(currentY / grid) * grid;
    }

    return positions;
}
