    // function getColWidthPx(width) {
    //     return Math.floor((width || 8.43) * 7 + 5);
    // }
    function getColWidthPx(width) {

    if (!width) {
        width = 8.43;
    }

    return Math.floor(
        ((256 * width + 128) / 256) * 7
    );
}

    function getColWidthPt(width) {

    const px = Math.floor((width || 8.43) * 7 + 5);

    return px * 0.75;
}


// ============================================
// SECTION MAPPING SYSTEM
// ============================================

// Canonical section names (ordered by typical report layout)
const CANONICAL_SECTIONS = [
  'title',
  'pageheader',
  'columnheader',
  'groupheader',
  'detail',
  'groupfooter',
  'columnfooter',
  'pagefooter',
  'summary'
];

// Section key aliases → canonical name
const SECTION_ALIASES = {
  'header': 'columnheader',
  'colheader': 'columnheader',
  'col_header': 'columnheader',
  'col-header': 'columnheader',
  'column header': 'columnheader',
  'column_header': 'columnheader',
  'column-header': 'columnheader',
  'page header': 'pageheader',
  'page_header': 'pageheader',
  'page-header': 'pageheader',
  'page footer': 'pagefooter',
  'page_footer': 'pagefooter',
  'page-footer': 'pagefooter',
  'column footer': 'columnfooter',
  'column_footer': 'columnfooter',
  'column-footer': 'columnfooter',
  'colfooter': 'columnfooter',
  'col_footer': 'columnfooter',
  'col-footer': 'columnfooter',
  'header footer': 'pagefooter',
  'foot': 'pagefooter',
  'end': 'end',
  // Group-related aliases
  'groupheader': 'groupheader',
  'grpheader': 'groupheader',
  'grp_header': 'groupheader',
  'grp-header': 'groupheader',
  'group header': 'groupheader',
  'group_header': 'groupheader',
  'group-header': 'groupheader',
  'groupfooter': 'groupfooter',
  'grpfooter': 'groupfooter',
  'grp_footer': 'groupfooter',
  'grp-footer': 'groupfooter',
  'group footer': 'groupfooter',
  'group_footer': 'groupfooter',
  'group-footer': 'groupfooter',
};

// Export type section configuration
// Each export type defines which canonical sections it supports and how they render
const EXPORT_SECTION_CONFIG = {
  npoi: {
    label: 'NPOI',
    supported: ['title', 'pageheader', 'columnheader', 'detail', 'columnfooter', 'pagefooter', 'summary'],
    // static: rendered as fixed CreateRow rows
    static: ['title', 'pageheader', 'columnheader', 'columnfooter', 'pagefooter', 'summary'],
    // dynamic: rendered as DataRow loop
    dynamic: ['detail'],
    // fallback for supported sections not in static/dynamic
    fallback: 'static'
  },
  jrxml: {
    label: 'JRXML',
    supported: ['title', 'pageheader', 'columnheader', 'groupheader', 'detail', 'groupfooter', 'columnfooter', 'pagefooter', 'summary'],
    // Jasper band name mapping
    bandMap: {
      'title': 'title',
      'pageheader': 'pageHeader',
      'columnheader': 'columnHeader',
      'groupheader': 'groupHeader',
      'detail': 'detail',
      'groupfooter': 'groupFooter',
      'columnfooter': 'columnFooter',
      'pagefooter': 'pageFooter',
      'summary': 'summary'
    }
  },
  rdlc: {
    label: 'RDLC',
    supported: ['title', 'pageheader', 'columnheader', 'detail', 'pagefooter'],
    // RDLC element placement
    pageHeader: ['title', 'pageheader', 'columnheader'],
    body: ['detail'],
    pageFooter: ['pagefooter']
  },
  ngprime: {
    label: 'PrimeNG',
    supported: ['title', 'pageheader', 'columnheader', 'detail', 'pagefooter', 'summary'],
    // PrimeNG uses sections as:
    // title → table caption
    // pageheader → above table
    // columnheader → table headers
    // detail → table body
    // pagefooter → below table
    // summary → table footer
    sectionMap: {
      'title': 'caption',
      'pageheader': 'above_table',
      'columnheader': 'header',
      'detail': 'body',
      'pagefooter': 'below_table',
      'summary': 'footer'
    }
  },
  xlsxstyle: {
    label: 'xlsx-style',
    supported: ['title', 'pageheader', 'columnheader', 'detail', 'columnfooter', 'pagefooter', 'summary'],
    detail: ['detail'],
    static: ['title', 'pageheader', 'columnheader', 'columnfooter', 'pagefooter', 'summary']
  },
  efcore: {
    label: 'EF Core',
    supported: ['title', 'pageheader', 'columnheader', 'detail', 'pagefooter', 'summary'],
    // Sections rendered as C# entity/DbContext/Excel export code
    // title → entity class name and Excel title row
    // columnheader → entity property names
    // detail → data loop
    // summary → summary rows in Excel
    // pagefooter → footer section
    sectionMap: {
      'title': 'entity_header',
      'pageheader': 'page_header',
      'columnheader': 'entity_columns',
      'detail': 'data_loop',
      'pagefooter': 'page_footer',
      'summary': 'summary_rows'
    }
  },
  typescript: {
    label: 'TypeScript',
    supported: ['title', 'pageheader', 'columnheader', 'detail', 'pagefooter', 'summary'],
    // TypeScript uses sections for:
    // title → component title header
    // columnheader → table column headers and search fields
    // detail → data table display
    // pageheader → additional header info
    // pagefooter → footer
    // summary → aggregate info
    sectionMap: {
      'title': 'component_title',
      'pageheader': 'page_header',
      'columnheader': 'entity_columns',
      'detail': 'data_loop',
      'pagefooter': 'page_footer',
      'summary': 'summary_rows'
    }
  },
  itextsharp: {
    label: 'iTextSharp',
    supported: ['title', 'pageheader', 'columnheader', 'detail', 'columnfooter', 'pagefooter', 'summary'],
    // iTextSharp renders all sections as PdfPTable rows in a single table
    // static: rendered as fixed PdfPCell rows
    static: ['title', 'pageheader', 'columnheader', 'columnfooter', 'pagefooter', 'summary'],
    // dynamic: rendered as data-driven loop
    dynamic: ['detail'],
    // fallback for supported sections not in static/dynamic
    fallback: 'static'
  }
};

// Normalize section key to canonical form (lowercase + alias resolution)
function normalizeSectionKey(key) {
  if (!key) return null;
  const lower = String(key).toLowerCase().trim();
  return SECTION_ALIASES[lower] || lower;
}

// Get display name for a section (e.g., 'pageheader' → 'PAGEHEADER')
function getSectionDisplayName(key) {
  const canonical = normalizeSectionKey(key);
  if (!canonical) return String(key).toUpperCase();
  return canonical.toUpperCase();
}

// Check if a section is supported by a specific export type
function isSectionExportable(sectionKey, exportType) {
  const config = EXPORT_SECTION_CONFIG[exportType];
  if (!config) return false;
  const normalizedKey = normalizeSectionKey(sectionKey);
  if (!normalizedKey) return false;
  return config.supported.includes(normalizedKey);
}

// Build sections from meta, normalizing all keys
// Filters out 'group:Name' definitions (they are processed by buildGroups separately)
function buildSections(meta, totalRows) {
    if (!meta || meta.length === 0) {
        return [{
            key: 'detail',
            start: 1,
            end: totalRows
        }];
    }

    const sections = [];

    for (let i = 0; i < meta.length; i++) {
        const current = meta[i];
        const next = meta[i + 1];

        // Skip group:Name definitions — they are not section bands
        if (/^group[:_]\s*\S+/i.test(String(current.key).trim())) {
            continue;
        }

        sections.push({
            key: current.key,
            start: current.row,
            end: next ? next.row - 1 : totalRows
        });
    }

    // Normalize section keys to canonical form
    sections.forEach(s => {
      const canonical = normalizeSectionKey(s.key);
      if (canonical && canonical !== 'end') {
        s.key = canonical;
      }
    });

    // Filter out 'end' sections
    return sections.filter(s => s.key.toLowerCase() !== 'end');
}

/**
 * Extract group definitions from meta.
 * Groups are defined in Column A with the pattern "group:Name" or "group_Name".
 * The function pairs each group definition with its groupheader/groupfooter
 * section markers that appear in subsequent rows.
 *
 * @param {Array} meta - Array of { row, key } from Column A
 * @returns {Array} Array of { name, headerRow, footerRow, detailRow }
 */
function buildGroups(meta) {
    if (!meta || meta.length === 0) return [];

    const groups = [];
    const stack = [];  // Stack for nested group tracking (LIFO)
    let lastGroup = null;

    for (const item of meta) {
        const key = String(item.key || '').trim();

        // Detect group:Name or group_Name pattern
        const groupMatch = key.match(/^group[:_]\s*(.+)$/i);
        if (groupMatch) {
            // Push current group to stack before starting a nested group
            if (lastGroup) {
                stack.push(lastGroup);
            }
            lastGroup = {
                name: groupMatch[1].trim(),
                headerRow: null,
                footerRow: null
            };
            groups.push(lastGroup);
            continue;
        }

        if (!lastGroup) continue;

        const canonical = normalizeSectionKey(key);
        if (canonical === 'groupheader' && lastGroup.headerRow === null) {
            lastGroup.headerRow = item.row;
        }
        if (canonical === 'groupfooter' && lastGroup.footerRow === null) {
            lastGroup.footerRow = item.row;
            // After footer is assigned, pop the parent group from stack (if nested)
            if (stack.length > 0) {
                lastGroup = stack.pop();
            } else {
                lastGroup = null;
            }
        }
    }

    return groups;
}



function parseDetailTemplate(row) {
    const config = [];

    row.eachCell({ includeEmpty: true }, (cell, cIdx) => {
        const text = (cell.text || '').trim();

        // Check for P{{param}} first (parameter reference)
        let param = null;
        let field = null;

        const paramMatch = text.match(/^[pP]\{\{(.+?)\}\}$/);
        if (paramMatch) {
            param = paramMatch[1];
        } else {
            const fieldMatch = text.match(/\{\{(.+?)\}\}/);
            if (fieldMatch) {
                field = fieldMatch[1];
            }
        }

        config.push({
            colIndex: cIdx - 1,
            field: field,
            param: param,
            // Plain text with no {{field}} / P{{param}} renders as a literal
            // that repeats on every record (e.g. a label row in multi-line detail).
            staticText: (!field && !param && text) ? text : null,
            isMerge: cell.isMerged,
        });
    });

    return config;
}
function toBase64(buffer) {
    // Node.js
    if (typeof Buffer !== 'undefined') {
        const buf = Buffer.isBuffer(buffer)
            ? buffer
            : Buffer.from(buffer);
        return buf.toString('base64');
    }

    // Browser (ArrayBuffer / Uint8Array)
    let bytes;

    if (buffer instanceof ArrayBuffer) {
        bytes = new Uint8Array(buffer);
    } else if (buffer instanceof Uint8Array) {
        bytes = buffer;
    } else {
        return '';
    }

    let binary = '';
    const len = bytes.length;

    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }

    return btoa(binary);
}





function pxToEMU(px) {
    const emu = px * 9525;
    return Math.max(0, Math.min(1023 * 9525, Math.round(emu)));
}

// ============================================
// RPPRINTIF / RP — Shared Utilities
// ============================================

/**
 * Extract note/comment text from a cell reliably.
 * Handles various ExcelJS note structures.
 */
function getCellNoteText(cell) {
    if (!cell) return '';

    // Internal _comment (ExcelJS v6+)
    if (cell._comment?.note?.texts) {
        return cell._comment.note.texts.map(x => x.text || '').join('').trim();
    }

    // Standard note (cloned worksheet structure)
    if (cell.note) {
        if (cell.note.texts) {
            return cell.note.texts.map(x => x.text || '').join('').trim();
        }
        if (cell.note.note?.texts) {
            return cell.note.note.texts.map(x => x.text || '').join('').trim();
        }
        if (typeof cell.note === 'string') {
            return cell.note.trim();
        }
    }

    // comment property (ExcelJS fallback)
    if (cell.comment?.texts) {
        return cell.comment.texts.map(x => x.text || '').join('').trim();
    }
    if (cell.comment?.note?.texts) {
        return cell.comment.note.texts.map(x => x.text || '').join('').trim();
    }

    // model?.note (internal ExcelJS)
    if (cell.model?.note?.texts) {
        return cell.model.note.texts.map(x => x.text || '').join('').trim();
    }

    return '';
}

/**
 * Parse a cell note for RPPRINTIF or RP formulas.
 * Supports multi-line expressions.
 *
 * @param {string} noteText - Text extracted from cell note/comment
 * @returns {object|null} { type: 'printWhen'|'textExpression', expression: string } or null
 */
function parseRPPRINTIF(noteText) {
    if (!noteText) return null;

    const str = String(noteText).trim();
    if (!str) return null;

    // Use [\s\S]*? to support multi-line (dotAll workaround)
    // =RPPRINTIF(condition) → print when expression
    let match = str.match(/^=RPPRINTIF\(([\s\S]*)\)$/i);
    if (match) {
        return {
            type: 'printWhen',
            expression: match[1].trim()
        };
    }

    // =RP(expression) → text expression (replaces cell value)
    match = str.match(/^=RP\(([\s\S]*)\)$/i);
    if (match) {
        return {
            type: 'textExpression',
            expression: match[1].trim()
        };
    }

    return null;
}

/**
 * Convert {{field}}/P{{param}} syntax to C# expressions (for NPOI).
 * Expects note text using {{fieldName}} for DataTable fields.
 */
function convertToCSharpExpression(expr) {
    if (!expr) return '';

    // P{{param}} → parameter reference (Params dictionary lookup)
    expr = expr.replace(/[pP]\{\{(.*?)\}\}/g, (_, name) => `Params["${name.trim()}"]`);

    // {{field}} → dtRow field access
    expr = expr.replace(/\{\{(.*?)\}\}/g, (_, name) => {
        const field = name.trim();
        return `Convert.ToDouble(dtRow["${field}"])`;
    });

    return expr;
}

/**
 * Convert {{field}}/P{{param}} syntax to RDLC VB-like expressions.
 */
function convertToRdlcConditionExpression(expr) {
    if (!expr) return '';

    // P{{param}} → Parameters!param.Value
    expr = expr.replace(/[pP]\{\{(.*?)\}\}/g, (_, name) => `Parameters!${name.trim()}.Value`);

    // {{field}} → Fields!field.Value
    expr = expr.replace(/\{\{(.*?)\}\}/g, (_, name) => `Fields!${name.trim()}.Value`);

    return expr;
}

/**
 * Convert {{field}}/P{{param}} syntax to Jasper expression (same as original).
 */
function convertToJasperExpression(expr) {
    if (!expr) return '';

    // P{{param}} → $P{param}
    expr = expr.replace(/[pP]\{\{(.*?)\}\}/g, (_, name) => `$P{${name.trim()}}`);

    // {{field}} → $F{field}
    expr = expr.replace(/\{\{(.*?)\}\}/g, (_, name) => `$F{${name.trim()}}`);

    return expr;
}