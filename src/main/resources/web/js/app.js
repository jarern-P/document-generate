// ============================================
// Project Document Generator — Front-end
// ============================================

(function () {
    'use strict';

    // ================= State =================
    const state = {
        workbook: null,        // ExcelJS workbook (template)
        meta: [],
        originalImages: [],
        jrxml: '',
        generatedJrxml: '',    // last full regeneration from the Excel file
        fileName: '',
        dataBase64: null,
        dataFileName: '',
        currentTemplateId: null
    };

    const $ = id => document.getElementById(id);

    // ================= Toast / Loading =================
    function showToast(message, type = 'info') {
        const container = $('toastContainer');
        if (!container) return;
        const icons = {
            success: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2"><polyline points="20 6 9 17 4 12"></polyline></svg>',
            error: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#dc2626" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>',
            warning: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>',
            info: '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#0891b2" stroke-width="2"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>'
        };
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.innerHTML = `<div class="toast-icon">${icons[type] || icons.info}</div><div class="toast-message">${escapeHtml(message)}</div>`;
        container.appendChild(toast);
        setTimeout(() => toast.remove(), 4500);
    }

    function showLoading(text) {
        $('loadingText').textContent = text || 'กำลังประมวลผล...';
        $('loadingOverlay').style.display = 'flex';
    }
    function hideLoading() {
        $('loadingOverlay').style.display = 'none';
    }

    function escapeHtml(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    function formatFileSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }

    function fileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result).split(',')[1] || '');
            reader.onerror = () => reject(reader.error);
            reader.readAsDataURL(file);
        });
    }

    async function api(url, opts = {}) {
        const res = await fetch(url, {
            headers: opts.body ? { 'Content-Type': 'application/json' } : {},
            ...opts
        });
        if (!res.ok) {
            let msg = 'HTTP ' + res.status;
            try {
                const j = await res.json();
                if (j && j.error) msg = j.error;
            } catch (_) { /* not json */ }
            throw new Error(msg);
        }
        return res;
    }

    // ================= Menu switching =================
    function switchMenu(name) {
        document.querySelectorAll('.menu-panel').forEach(p => p.style.display = 'none');
        document.querySelectorAll('.menu-btn').forEach(b => b.classList.toggle('active', b.dataset.menu === name));
        $('menu-' + name).style.display = 'block';
        if (name === 'generate') refreshTemplates(true);
    }

    // ================= Drop zone helpers =================
    function bindDropZone(zoneId, inputId, onFile) {
        const zone = $(zoneId);
        const input = $(inputId);
        zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
        zone.addEventListener('drop', e => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            const f = e.dataTransfer.files[0];
            if (f) onFile(f);
        });
        zone.addEventListener('click', e => {
            if (e.target.tagName === 'A' || e.target.closest('a')) return;
            input.click();
        });
        input.addEventListener('change', () => {
            const f = input.files[0];
            if (f) onFile(f);
        });
    }

    // ================= MENU 1 : Upload Template =================
    function parseTplSections() {
        // Reuse the shared section utilities (utils-cal.js) to build chips.
        const chips = $('tplSectionsChips');
        chips.innerHTML = '';
        const totalRows = state.meta.length
            ? (state.originalWorksheet ? state.originalWorksheet.rowCount || 0 : 0)
            : 0;
        if (!state.meta || state.meta.length === 0) {
            chips.innerHTML = '<span class="section-tag section-default">ไม่มีส่วนกำหนดใน Column A (แสดงเป็น detail ทั้งหมด)</span>';
            return;
        }
        let rowsToEnd = totalRows;
        // estimate: use original worksheet rowCount
        for (let i = 0; i < state.meta.length; i++) {
            const cur = state.meta[i];
            const next = state.meta[i + 1];
            const endRow = next ? next.row - 1 : rowsToEnd;
            const key = String(cur.key).trim();
            let cls = 'section-tag section-' + String(key).toLowerCase();
            if (/^group[:_]/i.test(key)) cls = 'section-tag section-group';
            const span = document.createElement('span');
            span.className = cls;
            span.textContent = (key.toUpperCase()) + (endRow ? ` · แถว ${cur.row}-${endRow}` : ` · แถว ${cur.row}`);
            chips.appendChild(span);
        }
        $('tplSectionsBox').style.display = 'block';
    }

    function renderDetected(jrxml) {
        // Parse generated JRXML with DOMParser to list fields & parameters.
        const fieldsRow = $('tplFieldsRow');
        const paramsRow = $('tplParamsRow');
        fieldsRow.style.display = 'none';
        paramsRow.style.display = 'none';
        if (!jrxml) return;
        try {
            const doc = new DOMParser().parseFromString(jrxml, 'application/xml');
            const fieldEls = doc.getElementsByTagName('field');
            const paramEls = doc.getElementsByTagName('parameter');
            const fc = $('tplFieldsChips');
            fc.innerHTML = '';
            for (let i = 0; i < fieldEls.length; i++) {
                const name = fieldEls[i].getAttribute('name') || '';
                const cls = fieldEls[i].getAttribute('class') || '';
                const chip = document.createElement('span');
                chip.className = 'chip chip-field';
                chip.innerHTML = `<b>${escapeHtml(name)}</b><i>${escapeHtml(shortType(cls))}</i>`;
                fc.appendChild(chip);
            }
            if (fieldEls.length) fieldsRow.style.display = 'block';

            const pc = $('tplParamsChips');
            pc.innerHTML = '';
            for (let i = 0; i < paramEls.length; i++) {
                const name = paramEls[i].getAttribute('name') || '';
                const cls = paramEls[i].getAttribute('class') || '';
                const chip = document.createElement('span');
                chip.className = 'chip chip-param';
                chip.innerHTML = `<b>${escapeHtml(name)}</b><i>${escapeHtml(shortType(cls))}</i>`;
                pc.appendChild(chip);
            }
            if (paramEls.length) paramsRow.style.display = 'block';
        } catch (e) {
            console.warn('parse jrxml for display failed', e);
        }
    }

    function shortType(cls) {
        if (!cls) return '';
        const s = cls.split('.').pop();
        return s === 'String' ? 'text' : s === 'Integer' ? 'int' : s;
    }

    async function handleTplFile(file) {
        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            showToast('กรุณาเลือกไฟล์ .xlsx เท่านั้น', 'error');
            return;
        }
        showLoading('กำลังอ่านไฟล์ Excel template...');
        try {
            const buf = await file.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(buf, {
                comments: 'emit',
                ignoreNodes: ['tableParts', 'autoFilter', 'dataValidations']
            });
            state.workbook = workbook;
            state.fileName = file.name;

            // Sheet selector
            const sel = $('tplSheetSelector');
            sel.innerHTML = '';
            if (workbook.worksheets.length > 1) {
                workbook.worksheets.forEach((ws, i) => {
                    const opt = document.createElement('option');
                    opt.value = i;
                    opt.textContent = ws.name;
                    sel.appendChild(opt);
                });
                $('tplSheetField').style.display = 'block';
            } else {
                $('tplSheetField').style.display = 'none';
            }

            $('tplFileName').textContent = file.name;
            $('tplFileSize').textContent = formatFileSize(file.size);
            $('tplFileInfo').style.display = 'flex';
            $('tplDetailCard').style.display = 'block';

            await convertSelectedSheet(0);
            hideLoading();
            showToast('โหลดไฟล์สำเร็จ ✓ แปลงเป็น JRXML เรียบร้อย', 'success');
        } catch (err) {
            console.error(err);
            hideLoading();
            showToast('เกิดข้อผิดพลาด: ' + (err.message || 'ไม่ทราบสาเหตุ'), 'error');
        }
    }

    async function convertSelectedSheet(idx) {
        if (!state.workbook) return;
        const ws = state.workbook.worksheets[idx];
        state.originalWorksheet = ws;
        const jrxml = buildJrxml(state.workbook, idx, state.fileName);
        state.generatedJrxml = jrxml;
        state.jrxml = jrxml;
        $('tplJrxml').value = jrxml;

        // Recompute meta for section chips — read Column A markers directly.
        const m = [];
        ws.eachRow({ includeEmpty: true }, (row, rIdx) => {
            const cellA = row.getCell(1);
            if (cellA && cellA.value != null && String(cellA.value).trim() !== '') {
                m.push({ row: rIdx, key: cellA.value });
            }
        });
        state.meta = m;
        parseTplSections();
        renderDetected(jrxml);
    }

    async function saveTemplate() {
        const name = $('tplName').value.trim();
        const desc = $('tplDesc').value.trim();
        const jrxml = $('tplJrxml').value.trim();
        if (!name) { showToast('กรุณากรอกชื่อเทมเพลต', 'warning'); return; }
        if (!jrxml) { showToast('JRXML ว่างเปล่า', 'warning'); return; }
        showLoading('กำลังบันทึกเทมเพลต...');
        try {
            const res = await api('/api/templates', {
                method: 'POST',
                body: JSON.stringify({ name, description: desc, jrxml })
            });
            const data = await res.json();
            hideLoading();
            showToast(`บันทึกเทมเพลต "${name}" สำเร็จ ✓ (id=${data.id})`, 'success');
            refreshTemplates(false);
        } catch (err) {
            hideLoading();
            showToast('บันทึกไม่สำเร็จ: ' + err.message, 'error');
        }
    }

    async function refreshTemplates(selectOnly) {
        try {
            const res = await api('/api/templates');
            const list = await res.json();

            // Fill the generate select
            const genSel = $('genTemplateSelect');
            const prevVal = genSel.value;
            genSel.innerHTML = '<option value="">-- กรุณาเลือกเทมเพลต --</option>';
            list.forEach(t => {
                const opt = document.createElement('option');
                opt.value = t.id;
                opt.textContent = t.name;
                if (t.description) opt.title = t.description;
                genSel.appendChild(opt);
            });
            if (list.some(t => String(t.id) === prevVal)) genSel.value = prevVal;
            else onGenTemplateChange();

            // List on upload page
            const listEl = $('tplList');
            if (selectOnly) return;
            if (list.length === 0) {
                listEl.innerHTML = '<div class="empty">ยังไม่มีเทมเพลต — สร้างจากฝั่งซ้ายก่อน</div>';
                return;
            }
            listEl.innerHTML = '';
            list.forEach(t => {
                const item = document.createElement('div');
                item.className = 'template-item';
                const fcount = (t.fields && t.fields.length) || 0;
                item.innerHTML = `
                    <div class="tpl-item-main">
                        <div class="tpl-item-name">${escapeHtml(t.name)}</div>
                        ${t.description ? `<div class="tpl-item-desc">${escapeHtml(t.description)}</div>` : ''}
                        <div class="tpl-item-meta">ฟิลด์ ${fcount} · แก้ไขล่าสุด ${escapeHtml(t.updatedAt || '')}</div>
                    </div>
                    <div class="tpl-item-actions">
                        <button class="btn btn-sm btn-use" data-id="${t.id}">ใช้สร้างเอกสาร</button>
                        <button class="btn btn-sm btn-danger" data-del="${t.id}">ลบ</button>
                    </div>`;
                listEl.appendChild(item);
            });
            listEl.querySelectorAll('[data-del]').forEach(b => {
                b.addEventListener('click', () => deleteTemplate(Number(b.dataset.del)));
            });
            listEl.querySelectorAll('[data-id]').forEach(b => {
                b.addEventListener('click', () => {
                    switchMenu('generate');
                    const sel = $('genTemplateSelect');
                    sel.value = b.dataset.id;
                    onGenTemplateChange();
                });
            });
        } catch (err) {
            if (!selectOnly) showToast('โหลดรายการเทมเพลตไม่สำเร็จ: ' + err.message, 'error');
        }
    }

    async function deleteTemplate(id) {
        if (!confirm('ต้องการลบเทมเพลตนี้หรือไม่?')) return;
        try {
            await api('/api/templates/' + id, { method: 'DELETE' });
            showToast('ลบเทมเพลตสำเร็จ', 'success');
            refreshTemplates(false);
        } catch (err) {
            showToast('ลบไม่สำเร็จ: ' + err.message, 'error');
        }
    }

    // ================= MENU 2 : Generate =================
    function onGenTemplateChange() {
        const id = Number($('genTemplateSelect').value);
        state.currentTemplateId = id;
        const info = $('genTemplateInfo');
        const dataCard = $('genDataCard');
        const outputCard = $('genOutputCard');
        if (!id) {
            info.style.display = 'none';
            dataCard.style.display = 'none';
            outputCard.style.display = 'none';
            setGenButtons(false);
            return;
        }
        // Fetch detail (includes fields/params)
        fetch('/api/templates/' + id)
            .then(r => r.json())
            .then(t => {
                info.style.display = 'block';
                $('genTemplateDesc').textContent = t.description || '(ไม่มีคำอธิบาย)';

                const fc = $('genFieldsChips');
                fc.innerHTML = '';
                (t.fields || []).forEach(f => {
                    const chip = document.createElement('span');
                    chip.className = 'chip chip-field';
                    chip.innerHTML = `<b>${escapeHtml(f.name)}</b><i>${escapeHtml(shortType(f.type))}</i>`;
                    fc.appendChild(chip);
                });

                // Params inputs
                const pi = $('genParamsInputs');
                pi.innerHTML = '';
                const params = (t.params || []).filter(p => !p.name.startsWith('REPORT_'));
                if (params.length > 0) {
                    $('genParamsRow').style.display = 'block';
                    params.forEach(p => {
                        const type = p.type || 'java.lang.String';
                        const wrap = document.createElement('div');
                        wrap.className = 'field field-inline';
                        const inputType = type.includes('Date') ? 'date' : (isNumericType(type) ? 'number' : 'text');
                        wrap.innerHTML = `
                            <label>${escapeHtml(p.name)} <i class="type-hint">(${escapeHtml(shortType(type))})</i></label>
                            <input type="${inputType}" class="input param-input" data-param="${escapeHtml(p.name)}" step="any">`;
                        pi.appendChild(wrap);
                    });
                } else {
                    $('genParamsRow').style.display = 'none';
                }

                dataCard.style.display = 'block';
                outputCard.style.display = 'block';
                updateGenButtons();
            })
            .catch(err => showToast('โหลดข้อมูลเทมเพลตไม่สำเร็จ: ' + err.message, 'error'));
    }

    function isNumericType(cls) {
        return /Integer|Long|Double|Float|BigDecimal|Short|Byte/.test(cls || '');
    }

    function updateGenButtons() {
        const ready = !!state.currentTemplateId && !!state.dataBase64;
        setGenButtons(ready);
    }

    function setGenButtons(enabled) {
        $('previewPdfBtn').disabled = !enabled;
        $('exportExcelBtn').disabled = !enabled;
    }

    async function handleDataFile(file) {
        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            showToast('กรุณาเลือกไฟล์ .xlsx เท่านั้น', 'error');
            return;
        }
        showLoading('กำลังอ่านไฟล์ข้อมูล...');
        try {
            state.dataBase64 = await fileToBase64(file);
            state.dataFileName = file.name;
            $('dataFileName').textContent = file.name;
            $('dataFileSize').textContent = formatFileSize(file.size);
            $('dataFileInfo').style.display = 'flex';
            hideLoading();
            updateGenButtons();
            previewDataSheet(file);
            showToast('โหลดข้อมูลสำเร็จ ✓', 'success');
        } catch (err) {
            hideLoading();
            showToast('อ่านไฟล์ข้อมูลไม่สำเร็จ: ' + err.message, 'error');
        }
    }

    async function previewDataSheet(file) {
        try {
            const buf = await file.arrayBuffer();
            const wb = new ExcelJS.Workbook();
            await wb.xlsx.load(buf, { ignoreNodes: ['tableParts', 'autoFilter', 'dataValidations'] });
            const ws = wb.worksheets[0];
            const rows = [];
            ws.eachRow({ includeEmpty: false }, (row, r) => {
                if (r > 11) return;
                const cells = [];
                row.eachCell({ includeEmpty: false }, cell => cells.push(cell.text == null ? '' : cell.text));
                if (cells.length) rows.push(cells);
            });
            if (!rows.length) return;
            const maxCols = Math.max(...rows.map(r => r.length));
            let html = '<thead><tr>';
            for (let c = 0; c < maxCols; c++) {
                html += `<th>${escapeHtml(rows[0][c] || '')}</th>`;
            }
            html += '</tr></thead><tbody>';
            for (let r = 1; r < rows.length; r++) {
                html += '<tr>';
                for (let c = 0; c < maxCols; c++) html += `<td>${escapeHtml(rows[r][c] || '')}</td>`;
                html += '</tr>';
            }
            html += '</tbody>';
            $('dataPreviewTable').innerHTML = html;
            $('dataPreviewBox').style.display = 'block';
        } catch (e) {
            console.warn('preview failed', e);
        }
    }

    function collectParams() {
        const out = {};
        document.querySelectorAll('.param-input').forEach(inp => {
            const name = inp.dataset.param;
            if (inp.value !== '') out[name] = inp.value;
        });
        return out;
    }

    async function generate(format) {
        if (!state.currentTemplateId) { showToast('กรุณาเลือกเทมเพลตก่อน', 'warning'); return; }
        if (!state.dataBase64) { showToast('กรุณาอัปโหลดไฟล์ข้อมูลก่อน', 'warning'); return; }
        showLoading(format === 'pdf' ? 'กำลังสร้างเอกสาร PDF...' : 'กำลังส่งออก Excel...');
        try {
            const res = await api('/api/render', {
                method: 'POST',
                body: JSON.stringify({
                    templateId: state.currentTemplateId,
                    format,
                    data: state.dataBase64,
                    params: collectParams()
                })
            });
            const blob = await res.blob();
            hideLoading();
            if (format === 'pdf') {
                const url = URL.createObjectURL(blob);
                $('pdfFrame').src = url;
                const dl = $('pdfDownloadLink');
                dl.href = url;
                dl.download = 'document.pdf';
                $('pdfModal').style.display = 'flex';
            } else {
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                const name = ($('genTemplateSelect').selectedOptions[0] || {}).textContent || 'document';
                a.download = name + '.xlsx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                setTimeout(() => URL.revokeObjectURL(url), 2000);
                showToast('ส่งออก Excel สำเร็จ ✓', 'success');
            }
        } catch (err) {
            hideLoading();
            showToast('สร้างเอกสารไม่สำเร็จ: ' + err.message, 'error');
        }
    }

    // ================= Boot =================
    document.addEventListener('DOMContentLoaded', () => {
        // Menus
        $('menuUploadBtn').addEventListener('click', () => switchMenu('upload'));
        $('menuGenerateBtn').addEventListener('click', () => switchMenu('generate'));

        // Upload page
        bindDropZone('tplDropZone', 'tplFileInput', handleTplFile);
        $('tplBrowseLink').addEventListener('click', e => { e.preventDefault(); $('tplFileInput').click(); });
        $('tplSheetSelector').addEventListener('change', e => convertSelectedSheet(Number(e.target.value)));
        $('tplJrxml').addEventListener('input', e => { state.jrxml = e.target.value; });
        $('tplRevertBtn').addEventListener('click', () => {
            $('tplJrxml').value = state.generatedJrxml || '';
            state.jrxml = $('tplJrxml').value;
            renderDetected(state.jrxml);
        });
        $('tplSaveBtn').addEventListener('click', saveTemplate);

        // Generate page
        $('genTemplateSelect').addEventListener('change', onGenTemplateChange);
        bindDropZone('dataDropZone', 'dataFileInput', handleDataFile);
        $('previewPdfBtn').addEventListener('click', () => generate('pdf'));
        $('exportExcelBtn').addEventListener('click', () => generate('xlsx'));

        // Modal
        $('pdfCloseBtn').addEventListener('click', () => {
            $('pdfModal').style.display = 'none';
            $('pdfFrame').src = '';
        });
        $('pdfModal').addEventListener('click', e => {
            if (e.target === $('pdfModal')) { $('pdfModal').style.display = 'none'; $('pdfFrame').src = ''; }
        });

        // Server health
        fetch('/api/health').then(() => {
            $('serverBadge').textContent = 'เซิร์ฟเวอร์พร้อมใช้งาน';
            $('serverBadge').classList.add('badge-ok');
        }).catch(() => {
            $('serverBadge').textContent = 'ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์';
            $('serverBadge').classList.add('badge-error');
        });

        refreshTemplates(false);
    });
})();
