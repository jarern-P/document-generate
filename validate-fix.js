const ExcelJS = require('exceljs');
const { exportToRDLC } = require('./export-rdlc');
const fs = require('fs');

// Polyfill for browser-only global functions used in export-rdlc.js
global.DOMParser = class { parseFromString() { return { getElementsByTagName: () => [] }; } };
global.getThemeColors = () => [];

// Override the internal calls that go to utils modules
global.getColWidthPt = (w) => Math.floor((w || 8.43) * 7 + 5) * 0.75;
global.getRowHeightPt = (h, def) => {
  if (h) return h * 0.75;
  return (def || 15) * 0.75;
};
global.buildSections = (meta, totalRows) => {
  const sections = [];
  if (!meta || !meta.sections) return [{ key: 'DEFAULT', start: 1, end: totalRows }];
  for (let i = 0; i < meta.sections.length; i++) {
    const current = meta.sections[i];
    const next = meta.sections[i + 1];
    sections.push({
      key: current.key,
      start: current.start,
      end: next ? next.start - 1 : totalRows
    });
  }
  return sections;
};
global.toARGB = () => null;
global.ptToCm = (pt) => {
  pt = Number(pt);
  if (isNaN(pt) || !isFinite(pt)) return '0';
  return (pt * 0.0352778).toFixed(3);
};
global.escapeXml = (value) => {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
};
global.parseRdlcExpression = () => null;
global.escapeRdlcString = (s) => String(s).replace(/"/g, '""');
global.formatRdlcExpression = (expr, value) => {
  return global.escapeXml(value === null || value === undefined || value === '' ? ' ' : value);
};
global.getCellDisplayValue = (cell) => {
  if (cell.value == null) return '';
  return String(cell.value);
};
global.getCellNoteText = () => '';
global.parseRPPRINTIF = () => null;
global.convertToRdlcConditionExpression = (e) => e || '';
global.getMergeRange = (ws, row, col) => null;
global.inferFieldType = () => 'java.lang.String';
global.buildTextboxStyle = () => '';
global.buildTextRunStyle = () => '<Style></Style>';
global.argbToHex = (a) => a ? '#' + a.substring(2) : '';
global.ColSpan = null;

async function main() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');

  // Setup columns
  for (let c = 1; c <= 4; c++) ws.getColumn(c).width = 15;

  // Setup rows with merged cells (to test ColSpan)
  ws.getRow(1).height = 20;
  ws.getCell('A1').value = 'TITLE';
  ws.getCell('B1').value = 'Report Title';
  ws.mergeCells('B1:C1');  // ColSpan=2

  ws.getRow(2).height = 20;
  ws.getCell('A2').value = 'PAGEHEADER';
  ws.getCell('B2').value = 'Header Data';

  ws.getRow(3).height = 20;
  ws.getCell('A3').value = 'DETAIL';
  ws.getCell('B3').value = '{{field1}}';
  ws.getCell('C3').value = '{{field2}}';
  ws.getCell('D3').value = '{{field3}}';

  ws.getRow(4).height = 20;
  ws.getCell('B4').value = '{{field1}}';
  ws.getCell('C4').value = '{{field2}}';

  const meta = {
    sections: [
      { key: 'title', start: 1 },
      { key: 'pageheader', start: 2 },
      { key: 'detail', start: 3 },
    ]
  };

  const xml = exportToRDLC(ws, wb, meta);
  fs.writeFileSync('test-validate.xml', xml);

  // Check ColSpan/RowSpan position
  const colSpanOutsideCellContents = xml.match(/<\/CellContents>\s*<ColSpan>/g);
  const colSpanInsideCellContents = xml.match(/<CellContents>[\s\S]*?<ColSpan>/g);
  const rowSpanOutsideCellContents = xml.match(/<\/CellContents>\s*<RowSpan>/g);
  const rowSpanInsideCellContents = xml.match(/<CellContents>[\s\S]*?<RowSpan>/g);

  console.log('=== Validation Results ===\n');

  // Check ColSpan is NOT outside CellContents
  if (colSpanOutsideCellContents) {
    console.log('❌ FAIL: ColSpan found OUTSIDE CellContents (' + colSpanOutsideCellContents.length + ' occurrences)');
  } else {
    console.log('✅ PASS: No ColSpan outside CellContents');
  }

  // Count ColSpan
  const colSpanTotal = (xml.match(/<ColSpan>\d+<\/ColSpan>/g) || []).length;
  console.log('✅ Total ColSpan elements: ' + colSpanTotal);

  // Check RowSpan is NOT outside CellContents
  if (rowSpanOutsideCellContents) {
    console.log('❌ FAIL: RowSpan found OUTSIDE CellContents (' + rowSpanOutsideCellContents.length + ' occurrences)');
  } else {
    console.log('✅ PASS: No RowSpan outside CellContents');
  }

  // Show sample
  const colSpanSample = xml.match(/<CellContents>[\s\S]*?<ColSpan>\d+<\/ColSpan>[\s\S]*?<\/CellContents>/);
  if (colSpanSample) {
    console.log('\n--- ColSpan Inside CellContents (correct) ---');
    console.log(colSpanSample[0]);
  }

  // Check that ColSpan/RowSpan is inside CellContents (NOT directly under TablixCell)
  // Look for ColSpan/RowSpan NOT followed by </CellContents> before </TablixCell>
  const badSpanRegex = /<ColSpan>\d+<\/ColSpan>(?![\s\S]*?<\/CellContents>)[\s\S]*?<\/TablixCell>/g;
  const badRowSpanRegex = /<RowSpan>\d+<\/RowSpan>(?![\s\S]*?<\/CellContents>)[\s\S]*?<\/TablixCell>/g;
  const badColSpans = xml.match(badSpanRegex) || [];
  const badRowSpans = xml.match(badRowSpanRegex) || [];
  const totalBadSpans = badColSpans.length + badRowSpans.length;
  if (totalBadSpans > 0) {
    console.log('\n❌ FAIL: ColSpan/RowSpan found OUTSIDE CellContents (' + totalBadSpans + ' occurrences, ColSpan: ' + badColSpans.length + ', RowSpan: ' + badRowSpans.length + ')');
  } else {
    console.log('\n✅ PASS: No ColSpan/RowSpan outside CellContents');
  }

  // Validate that the XML is well-formed
  console.log('\n--- RDLC Structure Checks ---');
  
  // Check required elements exist
  const hasReport = xml.includes('<Report xmlns=');
  const hasReportSections = xml.includes('<ReportSections>');
  const hasReportSection = xml.includes('<ReportSection>');
  const hasBody = xml.includes('<Body>');
  const hasTablix = xml.includes('<Tablix');
  const hasTablixBody = xml.includes('<TablixBody>');
  const hasTablixColumns = xml.includes('<TablixColumns>');
  const hasTablixRows = xml.includes('<TablixRows>');
  const hasPage = xml.includes('<Page>');
  const hasLanguage = xml.includes('<Language>');
  
  console.log('✅ Report element: ' + (hasReport ? 'YES' : 'MISSING'));
  console.log('✅ ReportSections: ' + (hasReportSections ? 'YES' : 'MISSING'));
  console.log('✅ ReportSection: ' + (hasReportSection ? 'YES' : 'MISSING'));
  console.log('✅ Body: ' + (hasBody ? 'YES' : 'MISSING'));
  console.log('✅ Tablix: ' + (hasTablix ? 'YES' : 'MISSING'));
  console.log('✅ TablixBody: ' + (hasTablixBody ? 'YES' : 'MISSING'));
  console.log('✅ TablixColumns: ' + (hasTablixColumns ? 'YES' : 'MISSING'));
  console.log('✅ TablixRows: ' + (hasTablixRows ? 'YES' : 'MISSING'));
  console.log('✅ Page: ' + (hasPage ? 'YES' : 'MISSING'));
  console.log('✅ Language: ' + (hasLanguage ? 'YES' : 'MISSING'));

  // Check TablixRows count
  const tablixRowCount = (xml.match(/<TablixRow>/g) || []).length;
  const hierarchyMemberCount = (xml.match(/<TablixMember(?:\s|>)/g) || []).length;
  const detailMemberCount = (xml.match(/<TablixMember \/>/g) || []).length;
  console.log('\n📊 TablixRows: ' + tablixRowCount);
  console.log('📊 TablixMembers (total): ' + hierarchyMemberCount);
  console.log('📊 TablixMembers (detail children): ' + detailMemberCount);

  // Check no empty style tags with trailing whitespace
  const emptyStyleWithNewline = xml.match(/<Style\s*\/>/g);
  console.log('\n📊 Empty Style tags: ' + (emptyStyleWithNewline ? emptyStyleWithNewline.length : 0));

  console.log('\n=== Overall ===');
  const allPass = !colSpanOutsideCellContents && !rowSpanOutsideCellContents && totalBadSpans === 0 && hasReport && hasLanguage;
  console.log(allPass ? '✅ ALL CHECKS PASSED' : '❌ SOME CHECKS FAILED');
  console.log('Output saved to: test-validate.xml');
}

main().catch(err => {
  console.error('Validation failed:', err.message);
  process.exit(1);
});
