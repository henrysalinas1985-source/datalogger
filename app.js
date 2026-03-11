document.addEventListener('DOMContentLoaded', () => {
    // === ESTADO ===
    let db = null;
    let allSheetsData = {};
    let currentClinic = '';
    let calibrationDates = {};
    let instrumentsBank = [];
    let savedTemplates = [];
    let selectedSerieForEdit = null;

    // === SCHEMA EXACTO DEL EXCEL DATALOGGER.xlsx ===
    // 8.1 Test de Inspección y Funcionalidad — resultado va en col C, filas 24-29
    const SCHEMA_81 = [
        { code: '8.1.1', label: 'Sensores', row: 24 },
        { code: '8.1.2', label: 'Ubicación de los sensores', row: 25 },
        { code: '8.1.3', label: 'Vainas de sensores', row: 26 },
        { code: '8.1.4', label: 'Cables UTP', row: 27 },
        { code: '8.1.5', label: 'Interfaces LAN', row: 28 },
        { code: '8.1.6', label: 'Fuentes de alimentación', row: 29 },
    ];

    // 8.2.1 Características del termómetro patrón — valores fijos en col I, filas 31-34
    const SCHEMA_821 = [
        { label: 'Incertidumbre de calibración (Up) (Hasta -70°C)', row: 31, value: 0.18 },
        { label: 'Incertidumbre de calibración (Up) (-20.10°C)', row: 32, value: 0.14 },
        { label: 'Incertidumbre de calibración (Up) (29°C)', row: 33, value: 0.11 },
        { label: 'Resolución (rp)', row: 34, value: 0.01 },
    ];

    // 8.2.2 Características del termómetro a calibrar — valores fijos en col I, filas 38/40
    const SCHEMA_822 = [
        { label: 'Error permitido', row: 38, value: 3 },
        { label: 'Resolución (rx)', row: 40, value: 2 },
    ];

    // Estado de Valoración — col B, filas 15-18
    const EVALUATION_SCHEMA = [
        { label: 'Inspección superada, el equipo es apto para el uso', row: 15 },
        { label: 'El equipo ha necesitado reparación', row: 16 },
        { label: 'El equipo no está reparado. No se puede usar', row: 17 },
    ];

    const DB_NAME = 'CalibracionesDB_v3';
    const DB_VERSION = 1;

    // === UTILIDADES ===
    function escapeHtml(str) {
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(str));
        return div.innerHTML;
    }

    function safeParseFloat(val) {
        if (val === '' || val === null || val === undefined) return null;
        const n = parseFloat(val);
        return isNaN(n) ? null : n;
    }

    // DOM
    const fileInput = document.getElementById('fileInput');
    const fileLabel = document.getElementById('fileLabel');
    const mainContent = document.getElementById('mainContent');
    const sheetSelector = document.getElementById('sheetSelector');
    const serieFilter = document.getElementById('serieFilter');
    const equiposTableBody = document.getElementById('equiposTableBody');
    const editModal = document.getElementById('editModal');
    const calibDateInput = document.getElementById('calibDateInput');
    const ordenMInput = document.getElementById('ordenMInput');
    const technicianInput = document.getElementById('technicianInput');
    const buildingInput = document.getElementById('buildingInput');
    const sectorInput = document.getElementById('sectorInput');
    const locationInput = document.getElementById('locationInput');
    const commentsInput = document.getElementById('commentsInput');
    const equipmentNameInput = document.getElementById('equipmentNameInput');
    const modalSerieInput = document.getElementById('modalSerieInput');
    const modelInput = document.getElementById('modelInput');
    const brandInput = document.getElementById('brandInput');
    const addInstrumentBtn = document.getElementById('addInstrumentBtn');
    const instrumentsContainer = document.getElementById('instrumentsContainer');
    const certFileInput = document.getElementById('certFileInput');
    const certStatus = document.getElementById('certStatus');
    const templateSelector = document.getElementById('templateSelector');
    const templateNameInput = document.getElementById('templateNameInput');
    const saveNewTemplateBtn = document.getElementById('saveNewTemplateBtn');
    const saveTemplateRow = document.getElementById('saveTemplateRow');
    const saveCalibBtn = document.getElementById('saveCalibBtn');
    const totalEquiposEl = document.getElementById('totalEquipos').querySelector('.val');
    const cercaVencerEl = document.getElementById('cercaVencer').querySelector('.val');
    const vencidosEl = document.getElementById('vencidos').querySelector('.val');

    // === INIT ===
    async function init() {
        try {
            await initDB();
            await loadSavedData();
            setupEventListeners();
            loadTemplates();
        } catch (err) {
            console.error('Init error:', err);
            alert('Error al iniciar: ' + err.message);
        }
    }

    // === INDEXEDDB ===
    function initDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(DB_NAME, DB_VERSION);
            req.onupgradeneeded = e => {
                const d = e.target.result;
                if (!d.objectStoreNames.contains('calibrations')) d.createObjectStore('calibrations', { keyPath: 'serie' });
                if (!d.objectStoreNames.contains('appData')) d.createObjectStore('appData', { keyPath: 'id' });
                if (!d.objectStoreNames.contains('templates')) d.createObjectStore('templates', { keyPath: 'id', autoIncrement: true });
            };
            req.onsuccess = e => { db = e.target.result; resolve(); };
            req.onerror = e => reject(e.target.error);
        });
    }

    async function storeCalibration(data) {
        const tx = db.transaction('calibrations', 'readwrite');
        const store = tx.objectStore('calibrations');
        // Preserve existing certificate if none provided
        if (!data.certificate) {
            const existing = await new Promise(r => { const q = store.get(data.serie); q.onsuccess = () => r(q.result); });
            if (existing && existing.certificate) {
                data.certificate = existing.certificate;
                data.certName = existing.certName;
            }
        }
        store.put(data);
    }

    function getAllCalibrations() {
        return new Promise(resolve => {
            if (!db) { resolve({}); return; }
            const map = {}, tx = db.transaction('calibrations', 'readonly');
            tx.objectStore('calibrations').openCursor().onsuccess = e => {
                const cur = e.target.result;
                if (cur) { map[cur.key] = cur.value; cur.continue(); }
                else { calibrationDates = map; updateInstrumentsBank(); resolve(map); }
            };
        });
    }

    function updateInstrumentsBank() {
        const uniq = new Map();
        Object.values(calibrationDates).forEach(c => {
            (c.instruments || []).forEach(i => {
                if (i.name && !uniq.has(i.name.toUpperCase())) uniq.set(i.name.toUpperCase(), i);
            });
        });
        instrumentsBank = Array.from(uniq.values());
        const dl = document.getElementById('instrumentsHistory');
        if (dl) { dl.innerHTML = ''; instrumentsBank.forEach(i => { const o = document.createElement('option'); o.value = i.name; dl.appendChild(o); }); }
    }

    // === EXCEL LOADING ===
    fileInput.addEventListener('change', e => {
        const file = e.target.files[0]; if (!file) return;
        const reader = new FileReader();
        reader.onload = ev => processWorkbook(XLSX.read(new Uint8Array(ev.target.result), { type: 'array' }), file.name);
        reader.readAsArrayBuffer(file);
    });

    async function processWorkbook(wb, filename) {
        allSheetsData = {};
        sheetSelector.innerHTML = '';
        wb.SheetNames.forEach(name => {
            allSheetsData[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: '' }).filter(r => Object.values(r).some(v => v !== ''));
            const opt = document.createElement('option'); opt.value = opt.textContent = name;
            sheetSelector.appendChild(opt);
        });
        currentClinic = wb.SheetNames[0];
        fileLabel.textContent = `✅ ${filename}`;
        db.transaction('appData', 'readwrite').objectStore('appData').put({ id: 'lastExcel', filename, allSheetsData, sheetNames: wb.SheetNames, currentClinic });
        mainContent.classList.remove('hidden');
        document.getElementById('configActions').classList.remove('hidden');
        renderTable();
    }

    async function loadSavedData() {
        const tx = db.transaction('appData', 'readonly');
        const last = await new Promise(r => { const q = tx.objectStore('appData').get('lastExcel'); q.onsuccess = () => r(q.result); });
        if (!last) return;
        allSheetsData = last.allSheetsData; currentClinic = last.currentClinic;
        sheetSelector.innerHTML = '';
        last.sheetNames.forEach(n => { const o = document.createElement('option'); o.value = o.textContent = n; if (n === currentClinic) o.selected = true; sheetSelector.appendChild(o); });
        fileLabel.textContent = `✅ ${last.filename} (Recuperado)`;
        mainContent.classList.remove('hidden');
        document.getElementById('configActions').classList.remove('hidden');
        renderTable();
    }

    document.getElementById('clearDataBtn').addEventListener('click', () => {
        if (confirm('¿Borrar datos cargados?')) { db.transaction('appData', 'readwrite').objectStore('appData').delete('lastExcel'); location.reload(); }
    });

    // === BACKUP: EXPORTAR / IMPORTAR ===
    function blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    }

    function base64ToBlob(dataUrl) {
        const [header, data] = dataUrl.split(',');
        const mime = header.match(/:(.*?);/)[1];
        const bytes = atob(data);
        const arr = new Uint8Array(bytes.length);
        for (let i = 0; i < bytes.length; i++) arr[i] = bytes.charCodeAt(i);
        return new Blob([arr], { type: mime });
    }

    document.getElementById('exportBackupBtn').addEventListener('click', async () => {
        try {
            const backup = { version: 1, exportDate: new Date().toISOString(), calibrations: {}, templates: [] };

            // Exportar calibraciones
            const calTx = db.transaction('calibrations', 'readonly');
            const allCals = await new Promise(r => { const q = calTx.objectStore('calibrations').getAll(); q.onsuccess = () => r(q.result); });
            for (const cal of allCals) {
                const entry = { ...cal };
                if (entry.certificate instanceof Blob) {
                    entry._certBase64 = await blobToBase64(entry.certificate);
                    delete entry.certificate;
                }
                backup.calibrations[cal.serie] = entry;
            }

            // Exportar plantillas
            const tmplTx = db.transaction('templates', 'readonly');
            const allTmpls = await new Promise(r => { const q = tmplTx.objectStore('templates').getAll(); q.onsuccess = () => r(q.result); });
            for (const tmpl of allTmpls) {
                const entry = { ...tmpl };
                if (entry.blob instanceof Blob) {
                    entry._blobBase64 = await blobToBase64(entry.blob);
                    delete entry.blob;
                }
                backup.templates.push(entry);
            }

            const json = JSON.stringify(backup, null, 2);
            const blob = new Blob([json], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `backup_calibraciones_${new Date().toISOString().slice(0, 10)}.json`;
            a.click();
            URL.revokeObjectURL(url);
            alert('✅ Backup exportado correctamente.');
        } catch (err) {
            console.error('Error al exportar:', err);
            alert('Error al exportar: ' + err.message);
        }
    });

    document.getElementById('importBackupBtn').addEventListener('click', () => {
        document.getElementById('importBackupFile').click();
    });

    document.getElementById('importBackupFile').addEventListener('change', async e => {
        const file = e.target.files[0];
        if (!file) return;
        try {
            const text = await file.text();
            const backup = JSON.parse(text);
            if (!backup.version || !backup.calibrations) throw new Error('Formato de backup inválido');

            const count = { cals: 0, tmpls: 0 };

            // Importar calibraciones
            const calTx = db.transaction('calibrations', 'readwrite');
            const calStore = calTx.objectStore('calibrations');
            for (const [serie, cal] of Object.entries(backup.calibrations)) {
                const entry = { ...cal, serie };
                if (entry._certBase64) {
                    entry.certificate = base64ToBlob(entry._certBase64);
                    delete entry._certBase64;
                }
                calStore.put(entry);
                count.cals++;
            }

            // Importar plantillas
            if (backup.templates && backup.templates.length > 0) {
                const tmplTx = db.transaction('templates', 'readwrite');
                const tmplStore = tmplTx.objectStore('templates');
                for (const tmpl of backup.templates) {
                    const entry = { ...tmpl };
                    delete entry.id; // Dejar que autoIncrement asigne nuevo ID
                    if (entry._blobBase64) {
                        entry.blob = base64ToBlob(entry._blobBase64);
                        delete entry._blobBase64;
                    }
                    tmplStore.add(entry);
                    count.tmpls++;
                }
            }

            alert(`✅ Backup importado: ${count.cals} calibraciones, ${count.tmpls} plantillas.`);
            location.reload();
        } catch (err) {
            console.error('Error al importar:', err);
            alert('Error al importar: ' + err.message);
        }
        e.target.value = ''; // Reset para permitir reimportar el mismo archivo
    });

    // === TABLA ===
    async function renderTable() {
        if (!currentClinic || !allSheetsData[currentClinic]) return;
        await getAllCalibrations();
        const rows = allSheetsData[currentClinic];
        const search = serieFilter.value.trim().toUpperCase();
        equiposTableBody.innerHTML = '';
        let stats = { total: 0, warn: 0, danger: 0 };

        rows.forEach(row => {
            const keys = Object.keys(row);
            // Datalogger Excel: columna "Sensor" tiene el número de serie
            const serieKey = keys.find(k => k.toLowerCase().includes('sensor') || k.toLowerCase().includes('serie') || k.toLowerCase().includes('n°'));
            // Ubicación del sensor como nombre
            const nombreKey = keys.find(k => k.toLowerCase().includes('ubicacion') || k.toLowerCase().includes('ubicación') || k.toLowerCase().includes('lugar'));

            const serie = String(row[serieKey] || '').toUpperCase().trim();
            if (!serie || serie === '') return;
            if (search && !serie.includes(search)) return;

            stats.total++;
            const cal = calibrationDates[serie] || null;
            const status = getStatus(cal?.date);
            if (status.class === 'status-warning') stats.warn++;
            if (status.class === 'status-danger') stats.danger++;

            const displayName = cal?.editedName || (nombreKey ? row[nombreKey] : 'N/A');
            const displaySerie = cal?.editedSerie || serie;
            const safeSerie = escapeHtml(serie);
            const safeDisplayName = escapeHtml(String(displayName));
            const safeDisplaySerie = escapeHtml(String(displaySerie));

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${safeDisplayName}</td>
                <td>${safeDisplaySerie}</td>
                <td>${cal ? formatDate(cal.date) : '<span style="color:#aaa">Sin registrar</span>'}</td>
                <td>${escapeHtml(cal?.technician || '-')}</td>
                <td>${cal?.certificate ? `<button class="btn btn-small" data-action="viewCert" data-serie="${safeSerie}">📄</button>` : '-'}</td>
                <td><span class="status-badge ${status.class}">${escapeHtml(status.text)}</span></td>
                <td><button class="btn btn-secondary btn-small" data-action="openEdit" data-serie="${safeSerie}">📅 Registrar</button></td>
            `;
            equiposTableBody.appendChild(tr);
        });

        totalEquiposEl.textContent = stats.total;
        cercaVencerEl.textContent = stats.warn;
        vencidosEl.textContent = stats.danger;
    }

    function getStatus(dateStr) {
        if (!dateStr) return { text: 'Pendiente', class: '' };
        const next = new Date(dateStr); next.setFullYear(next.getFullYear() + 1);
        const diff = Math.ceil((next - new Date()) / 86400000);
        if (diff < 0) return { text: 'Vencido', class: 'status-danger' };
        if (diff <= 30) return { text: `Vence ${diff}d`, class: 'status-warning' };
        return { text: 'Vigente', class: 'status-ok' };
    }

    // === INSTRUMENTAL ===
    function createInstrumentRow(data = {}) {
        const div = document.createElement('div'); div.className = 'instrument-item';
        div.innerHTML = `
            <button type="button" class="remove-instrument">×</button>
            <div class="field-group full-width"><label>Nombre del Instrumental</label><input type="text" class="inst-name" list="instrumentsHistory" value="${escapeHtml(data.name || '')}"></div>
            <div class="field-group"><label>Marca</label><input type="text" class="inst-brand" value="${escapeHtml(data.brand || '')}"></div>
            <div class="field-group"><label>Modelo</label><input type="text" class="inst-model" value="${escapeHtml(data.model || '')}"></div>
            <div class="field-group"><label>N° Serie</label><input type="text" class="inst-serie" value="${escapeHtml(data.serie || '')}"></div>
            <div class="field-group"><label>Últ. Calibración</label><input type="text" class="inst-date" placeholder="DD/MM/YYYY" value="${escapeHtml(data.date || '')}"></div>
        `;
        div.querySelector('.remove-instrument').onclick = () => div.remove();
        instrumentsContainer.appendChild(div);
    }
    addInstrumentBtn.onclick = () => createInstrumentRow();

    function getInstrumentsData() {
        return Array.from(instrumentsContainer.querySelectorAll('.instrument-item')).map(div => ({
            name: div.querySelector('.inst-name').value,
            brand: div.querySelector('.inst-brand').value,
            model: div.querySelector('.inst-model').value,
            serie: div.querySelector('.inst-serie').value,
            date: div.querySelector('.inst-date').value,
        }));
    }

    // === INSPECCIÓN UI ===
    function renderInspectionPoints(saved = {}) {
        const container = document.getElementById('inspectionPointsContainer');
        container.innerHTML = '';

        // ── Sección 8.1 ──────────────────────────────────────────────────
        addSectionHeader(container, '8.1 Test de Inspección y Funcionalidad');
        addSubLabel(container, 'Marcar P (Pasó) o NP (No Pasó) según corresponda');
        SCHEMA_81.forEach(item => {
            const saved_val = saved[item.label] || 'na';
            const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
            rowEl.innerHTML = `
                <div class="inspection-label"><strong>${item.code}</strong> ${item.label}</div>
                <div class="inspection-options" data-label="${item.label}" data-type="choice" data-row="${item.row}" data-col="C">
                    <div class="inspection-opt ${saved_val === 'P' ? 'selected' : ''}" data-val="P">P</div>
                    <div class="inspection-opt ${saved_val === 'NP' ? 'selected' : ''}" data-val="NP" style="background:${saved_val === 'NP' ? '#e53e3e' : ''}">NP</div>
                    <div class="inspection-opt ${saved_val === 'na' ? 'selected' : ''}" data-val="na">NA</div>
                </div>`;
            wireChoiceOpts(rowEl);
            container.appendChild(rowEl);
        });

        // ── Sección 8.2.1 ────────────────────────────────────────────────
        addSectionHeader(container, '8.2.1 Características del Termómetro Patrón');
        SCHEMA_821.forEach(item => {
            const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
            rowEl.innerHTML = `
                <div class="inspection-label">${item.label}</div>
                <div class="inspection-options" style="justify-content:flex-end;padding:6px 12px;">
                    <span style="color:var(--accent);font-weight:700;">${item.value} °C</span>
                </div>`;
            container.appendChild(rowEl);
        });

        // ── Sección 8.2.2 ────────────────────────────────────────────────
        addSectionHeader(container, '8.2.2 Características del Termómetro a Calibrar');
        SCHEMA_822.forEach(item => {
            const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
            rowEl.innerHTML = `
                <div class="inspection-label">${item.label}</div>
                <div class="inspection-options" style="justify-content:flex-end;padding:6px 12px;">
                    <span style="color:var(--accent);font-weight:700;">${item.value} °C</span>
                </div>`;
            container.appendChild(rowEl);
        });

        // ── Sección 8.2.3 Mediciones por Sensor ─────────────────────────
        addSectionHeader(container, '8.2.3 Mediciones por Sensor (Columnas C=Datalogger | I=Patrón)');
        addSubLabel(container, 'Cada sensor tiene 3 lecturas. Las columnas J-N se calculan automáticamente al generar el Excel.');

        // Mostrar sensores ya guardados + opción de agregar más
        const savedSensors = saved['_sensors'] || [];
        const sensorsDiv = document.createElement('div');
        sensorsDiv.id = 'sensorsContainer';
        container.appendChild(sensorsDiv);

        if (savedSensors.length > 0) {
            savedSensors.forEach((s, i) => addSensorBlock(sensorsDiv, i, s));
        } else {
            addSensorBlock(sensorsDiv, 0); // primer sensor vacío
        }

        const addSensorBtn = document.createElement('button');
        addSensorBtn.type = 'button';
        addSensorBtn.className = 'btn btn-secondary btn-small';
        addSensorBtn.textContent = '+ Agregar Sensor';
        addSensorBtn.onclick = () => {
            const count = sensorsDiv.querySelectorAll('.sensor-block').length;
            addSensorBlock(sensorsDiv, count);
        };
        container.appendChild(addSensorBtn);
    }

    function addSectionHeader(container, text) {
        const h = document.createElement('div'); h.className = 'inspection-category'; h.textContent = text;
        container.appendChild(h);
    }
    function addSubLabel(container, text) {
        const p = document.createElement('p'); p.style.cssText = 'font-size:0.78em;color:#aaa;margin:2px 0 8px 0;'; p.textContent = text;
        container.appendChild(p);
    }

    function addSensorBlock(container, index, saved = {}) {
        const startRow = 43 + index * 3; // Bloques: 43-45, 46-48, 49-51 ...
        const block = document.createElement('div');
        block.className = 'sensor-block';
        block.dataset.index = index;
        block.dataset.startRow = startRow;
        block.style.cssText = 'border:1px solid #3a3a5c;border-radius:8px;padding:12px;margin-bottom:12px;background:#1a1a2e;';
        const eSensor = escapeHtml(saved.sensor || '');
        const eUb1 = escapeHtml(saved.ub1 || '');
        const eUb2 = escapeHtml(saved.ub2 || '');
        const eUb3 = escapeHtml(saved.ub3 || '');
        const eDl1 = escapeHtml(saved.dl1 || '');
        const eDl2 = escapeHtml(saved.dl2 || '');
        const eDl3 = escapeHtml(saved.dl3 || '');
        const ePt1 = escapeHtml(saved.pt1 || '');
        const ePt2 = escapeHtml(saved.pt2 || '');
        const ePt3 = escapeHtml(saved.pt3 || '');
        block.innerHTML = `
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
                <strong style="color:#7c83fd;">Sensor ${index + 1} (filas ${startRow}-${startRow + 2})</strong>
                ${index > 0 ? `<button type="button" class="remove-sensor btn btn-small" style="background:#e53e3e">× Quitar</button>` : ''}
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px;">
                <div class="field-group"><label>N° Sensor / Serie (Col B)</label><input type="text" class="s-sensor" value="${eSensor}"></div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:4px;background:#2a2a4a;padding:6px;border-radius:4px;">
                <div class="field-group"><label>Ubicación 1 (A${startRow})</label><input type="text" class="s-ub1" value="${eUb1}"></div>
                <div class="field-group"><label>DL L1 (H${startRow})</label><input type="number" step="any" class="s-dl1" value="${eDl1}"></div>
                <div class="field-group"><label>Patrón L1 (I${startRow})</label><input type="number" step="any" class="s-pt1" value="${ePt1}"></div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:4px;background:#2a2a4a;padding:6px;border-radius:4px;">
                <div class="field-group"><label>Ubicación 2 (A${startRow + 1})</label><input type="text" class="s-ub2" value="${eUb2}"></div>
                <div class="field-group"><label>DL L2 (H${startRow + 1})</label><input type="number" step="any" class="s-dl2" value="${eDl2}"></div>
                <div class="field-group"><label>Patrón L2 (I${startRow + 1})</label><input type="number" step="any" class="s-pt2" value="${ePt2}"></div>
            </div>
            <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;background:#2a2a4a;padding:6px;border-radius:4px;">
                <div class="field-group"><label>Ubicación 3 (A${startRow + 2})</label><input type="text" class="s-ub3" value="${eUb3}"></div>
                <div class="field-group"><label>DL L3 (H${startRow + 2})</label><input type="number" step="any" class="s-dl3" value="${eDl3}"></div>
                <div class="field-group"><label>Patrón L3 (I${startRow + 2})</label><input type="number" step="any" class="s-pt3" value="${ePt3}"></div>
            </div>
        `;
        if (index > 0) block.querySelector('.remove-sensor').onclick = () => block.remove();
        container.appendChild(block);
    }

    function wireChoiceOpts(rowEl) {
        rowEl.querySelectorAll('.inspection-opt').forEach(opt => {
            opt.onclick = () => {
                rowEl.querySelectorAll('.inspection-opt').forEach(o => {
                    o.classList.remove('selected');
                    if (o.dataset.val === 'NP') o.style.background = '';
                });
                opt.classList.add('selected');
                if (opt.dataset.val === 'NP') opt.style.background = '#e53e3e';
            };
        });
    }

    function getInspectionsData() {
        const data = {};
        document.querySelectorAll('.inspection-options').forEach(g => {
            const label = g.dataset.label;
            if (g.dataset.type === 'numeric') {
                data[label] = g.querySelector('input')?.value || '';
            } else if (g.dataset.type === 'choice') {
                const sel = g.querySelector('.inspection-opt.selected');
                data[label] = sel ? sel.dataset.val : 'na';
            }
        });
        const sensors = [];
        document.querySelectorAll('.sensor-block').forEach(block => {
            sensors.push({
                startRow: parseInt(block.dataset.startRow),
                sensor: block.querySelector('.s-sensor')?.value || '',
                ub1: block.querySelector('.s-ub1')?.value || '',
                ub2: block.querySelector('.s-ub2')?.value || '',
                ub3: block.querySelector('.s-ub3')?.value || '',
                dl1: block.querySelector('.s-dl1')?.value || '',
                dl2: block.querySelector('.s-dl2')?.value || '',
                dl3: block.querySelector('.s-dl3')?.value || '',
                pt1: block.querySelector('.s-pt1')?.value || '',
                pt2: block.querySelector('.s-pt2')?.value || '',
                pt3: block.querySelector('.s-pt3')?.value || '',
            });
        });
        data['_sensors'] = sensors;
        return data;
    }

    function getEvaluationsData() {
        const data = {};
        document.querySelectorAll('#evaluationStatusContainer .inspection-options').forEach(g => {
            const sel = g.querySelector('.inspection-opt.selected');
            data[g.dataset.label] = sel ? sel.dataset.val : '';
        });
        return data;
    }

    // === TEMPLATES ===
    async function loadTemplates() {
        if (!db) return;
        const tx = db.transaction('templates', 'readonly');
        savedTemplates = await new Promise(r => { const q = tx.objectStore('templates').getAll(); q.onsuccess = () => r(q.result); });
        templateSelector.innerHTML = '<option value="">-- Seleccionar Plantilla --</option>';
        savedTemplates.forEach(t => { const o = document.createElement('option'); o.value = t.id; o.textContent = t.name; templateSelector.appendChild(o); });
    }

    // === EVENT LISTENERS ===
    function setupEventListeners() {
        document.getElementById('dropZone').addEventListener('click', () => fileInput.click());
        sheetSelector.addEventListener('change', e => { currentClinic = e.target.value; renderTable(); });
        serieFilter.addEventListener('input', renderTable);

        certFileInput.addEventListener('change', async e => {
            const file = e.target.files[0]; if (!file) return;
            const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
            saveTemplateRow.classList.toggle('hidden', !isExcel);

            if (isExcel) {
                try {
                    const extracted = await extractInstrumentsFromExcel(file);
                    if (extracted && extracted.length > 0) {
                        const currentInstruments = getInstrumentsData();
                        if (currentInstruments.length > 0) {
                            if (confirm(`Se detectaron ${extracted.length} instrumentos en el Excel. ¿Deseas añadirlos a la lista actual? (Cancelar para limpiar primero)`)) {
                                extracted.forEach(inst => createInstrumentRow(inst));
                            } else {
                                instrumentsContainer.innerHTML = '';
                                extracted.forEach(inst => createInstrumentRow(inst));
                            }
                        } else {
                            extracted.forEach(inst => createInstrumentRow(inst));
                        }
                    }
                } catch (err) {
                    console.error("Error al extraer instrumentos del archivo subido:", err);
                }
            }
        });

        saveNewTemplateBtn.addEventListener('click', async () => {
            const file = certFileInput.files[0], name = templateNameInput.value.trim();
            if (!file || !name) { alert('Falta archivo o nombre'); return; }
            const tx = db.transaction('templates', 'readwrite');
            tx.objectStore('templates').add({ name, blob: file });
            tx.oncomplete = () => { alert('Plantilla guardada'); templateNameInput.value = ''; saveTemplateRow.classList.add('hidden'); loadTemplates(); };
        });

        document.getElementById('closeModalBtn').onclick = () => editModal.classList.add('hidden');

        // Event delegation para botones de la tabla (reemplaza window.openEdit / window.viewCert)
        equiposTableBody.addEventListener('click', e => {
            const btn = e.target.closest('[data-action]');
            if (!btn) return;
            const serie = btn.dataset.serie;
            if (btn.dataset.action === 'viewCert') {
                const c = calibrationDates[serie];
                if (c?.certificate) window.open(URL.createObjectURL(c.certificate), '_blank');
            } else if (btn.dataset.action === 'openEdit') {
                openEditModal(serie);
            }
        });

        function openEditModal(serie) {
            selectedSerieForEdit = serie;
            const existing = calibrationDates[serie] || {};
            const eqRow = (allSheetsData[currentClinic] || [])
                .find(r => String(r[Object.keys(r).find(k => k.toLowerCase().includes('sensor') || k.toLowerCase().includes('serie'))] || '').toUpperCase() === serie) || {};

            calibDateInput.value = existing.date || '';
            ordenMInput.value = existing.ordenM || '';
            technicianInput.value = existing.technician || '';
            buildingInput.value = existing.building || eqRow.edificio || '';
            sectorInput.value = existing.sector || eqRow.sector || '';
            locationInput.value = existing.location || eqRow.ubicacion || eqRow.ubicación || '';
            equipmentNameInput.value = existing.editedName || eqRow.equipo || '';
            modalSerieInput.value = existing.editedSerie || serie;
            modelInput.value = existing.model || eqRow.modelo || '';
            brandInput.value = existing.brand || eqRow.marca || '';
            commentsInput.value = existing.comments || '';

            instrumentsContainer.innerHTML = '';
            // Si hay instrumentos guardados, los renderizamos
            if (existing.instruments && existing.instruments.length > 0) {
                existing.instruments.forEach(i => createInstrumentRow(i));
            }
            // Si no hay instrumentos pero hay un certificado en la BD, lo escaneamos en busca de instrumentos
            else if (existing.certificate && existing.certName && (existing.certName.toLowerCase().endsWith('.xlsx') || existing.certName.toLowerCase().endsWith('.xls'))) {
                extractInstrumentsFromExcel(existing.certificate).then(extracted => {
                    if (extracted && extracted.length > 0) {
                        instrumentsContainer.innerHTML = ''; // Limpiar de nuevo por si acaso antes de rellenar asincrónicamente
                        extracted.forEach(inst => createInstrumentRow(inst));
                    }
                }).catch(err => console.error("Error al extraer instrumentos iniciales:", err));
            }

            // Estado de Valoración
            const evalContainer = document.getElementById('evaluationStatusContainer');
            evalContainer.innerHTML = '';
            EVALUATION_SCHEMA.forEach(item => {
                const cur = (existing.evaluations || {})[item.label] || '';
                const rowEl = document.createElement('div'); rowEl.className = 'inspection-row';
                rowEl.innerHTML = `
                    <div class="inspection-label">${item.label}</div>
                    <div class="inspection-options" data-label="${item.label}" data-type="evaluation">
                        <div class="inspection-opt ${cur === 'P' ? 'selected' : ''}" data-val="P">P</div>
                        <div class="inspection-opt ${cur === 'NA' ? 'selected' : ''}" data-val="NA">NA</div>
                    </div>`;
                rowEl.querySelectorAll('.inspection-opt').forEach(o => o.onclick = () => {
                    rowEl.querySelectorAll('.inspection-opt').forEach(x => x.classList.remove('selected'));
                    o.classList.add('selected');
                });
                evalContainer.appendChild(rowEl);
            });

            certStatus.textContent = existing.certName ? `Certificado: ${existing.certName}` : 'Sin certificado';

            renderInspectionPoints(existing.inspections || {});
            editModal.classList.remove('hidden');
        }

        // Resetear equipo: borra la calibración de IndexedDB
        document.getElementById('resetCalibBtn').onclick = async () => {
            if (!selectedSerieForEdit) return;
            if (!confirm(`¿Estás seguro de resetear todos los datos de calibración para "${selectedSerieForEdit}"? Esta acción no se puede deshacer.`)) return;
            try {
                const tx = db.transaction('calibrations', 'readwrite');
                tx.objectStore('calibrations').delete(selectedSerieForEdit);
                tx.oncomplete = () => {
                    editModal.classList.add('hidden');
                    renderTable();
                    alert('Equipo reseteado correctamente.');
                };
            } catch (err) {
                console.error(err);
                alert('Error al resetear: ' + err.message);
            }
        };

        saveCalibBtn.onclick = async () => {
            if (!selectedSerieForEdit) return;
            if (!calibDateInput.value) { alert('Fecha requerida'); return; }
            const inspections = getInspectionsData();
            const evaluations = getEvaluationsData();
            const instruments = getInstrumentsData();
            const selectedTmplId = templateSelector.value;
            const tmpl = savedTemplates.find(t => String(t.id) === String(selectedTmplId));
            let blob = certFileInput.files[0] || (tmpl ? tmpl.blob : null);

            try {
                let finalCert = blob;
                if (blob && (blob.name?.endsWith('.xlsx') || blob.name?.endsWith('.xls'))) {
                    finalCert = await updateExcelCertificate(blob, {
                        editedName: equipmentNameInput.value,
                        editedSerie: modalSerieInput.value,
                        model: modelInput.value,
                        brand: brandInput.value,
                        building: buildingInput.value,
                        sector: sectorInput.value,
                        location: locationInput.value,
                        date: calibDateInput.value,
                        ordenM: ordenMInput.value,
                        technician: technicianInput.value,
                        instruments,
                        inspections,
                        evaluations,
                    });
                }

                await storeCalibration({
                    serie: selectedSerieForEdit,
                    date: calibDateInput.value,
                    technician: technicianInput.value,
                    ordenM: ordenMInput.value,
                    building: buildingInput.value,
                    sector: sectorInput.value,
                    location: locationInput.value,
                    brand: brandInput.value,
                    model: modelInput.value,
                    comments: commentsInput.value,
                    editedName: equipmentNameInput.value,
                    editedSerie: modalSerieInput.value,
                    instruments,
                    inspections,
                    evaluations,
                    certificate: finalCert,
                    certName: finalCert?.name,
                });
                editModal.classList.add('hidden');
                renderTable();
                alert('✅ Calibración guardada exitosamente.');
            } catch (err) {
                console.error(err);
                alert('Error al guardar: ' + err.message);
            }
        };
    }

    // === EXCEL CERTIFICATE UPDATE ===
    async function updateExcelCertificate(blob, d) {
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(await blob.arrayBuffer());
        const ws = wb.getWorksheet('Certificado') || wb.worksheets[0];
        if (!ws) throw new Error('Hoja "Certificado" no encontrada');

        // ── Cabecera (según DOCUMENTACION_CAMBIOS y lectura real del Excel) ──
        // A5: Equipo, E5: Modelo, H5: Edificio, H6: Sector, H7: Ubicación
        // A8: N° serie, H8: Fecha, A9: Orden M / H9, H10: Técnico
        ws.getCell('A5').value = `Equipo: ${d.editedName}`;
        ws.getCell('E5').value = `Modelo: ${d.model}`;
        ws.getCell('A8').value = `N° serie: ${d.editedSerie}`;
        ws.getCell('E8').value = `Marca: ${d.brand}`;
        ws.getCell('H5').value = d.building;
        ws.getCell('H6').value = d.sector;
        ws.getCell('H7').value = d.location;
        ws.getCell('H8').value = d.date ? formatDate(d.date) : '';
        ws.getCell('H9').value = d.ordenM;
        ws.getCell('H10').value = d.technician;

        // ── Instrumental (fila 12, col A=nombre B=marca D=modelo E=serie F=fecha) ──
        (d.instruments || []).forEach((inst, i) => {
            if (i >= 4) return; // máx 4 instrumentos
            const r = 12 + i;
            ws.getCell(`A${r}`).value = inst.name || '';
            ws.getCell(`B${r}`).value = inst.brand || '';
            ws.getCell(`D${r}`).value = inst.model || '';
            ws.getCell(`E${r}`).value = inst.serie || '';
            ws.getCell(`F${r}`).value = inst.date || '';
        });

        // ── Estado de Valoración (col B, filas 15-17) ──
        const evSchema = [
            'Inspección superada, el equipo es apto para el uso',
            'El equipo ha necesitado reparación',
            'El equipo no está reparado. No se puede usar',
        ];
        evSchema.forEach((label, i) => {
            ws.getCell(`B${15 + i}`).value = (d.evaluations || {})[label] || '';
        });

        // ── 8.1 Test de Inspección — resultado en col J, filas 24-29 ──
        SCHEMA_81.forEach(item => {
            const val = (d.inspections || {})[item.label] || '';
            const cell = ws.getCell(`J${item.row}`);
            cell.value = val;
            if (val === 'NP' || val === 'F') { // Handle F or NP
                cell.font = { color: { argb: 'FFFF0000' }, bold: true };
            }
        });

        // ── 8.2.1 Características Patrón — col I, filas 31-34 (valores fijos) ──
        SCHEMA_821.forEach(item => {
            ws.getCell(`I${item.row}`).value = item.value;
        });

        // ── 8.2.2 Características Calibrar — col I, filas 38/40 (valores fijos) ──
        SCHEMA_822.forEach(item => {
            ws.getCell(`I${item.row}`).value = item.value;
        });

        // ── 8.2.3 Mediciones Multi-Sensor ──────────────────────────────────
        const sensors = (d.inspections || {})['_sensors'] || [];
        sensors.forEach(s => {
            const r = s.startRow; // startRow = 43, 46, 49 ...
            // Ubicación (col A, 3 filas)
            ws.getCell(`A${r}`).value = s.ub1 || '';
            ws.getCell(`A${r + 1}`).value = s.ub2 || '';
            ws.getCell(`A${r + 2}`).value = s.ub3 || '';
            // Sensor/Serie (col B, solo primera fila del bloque)
            ws.getCell(`B${r}`).value = s.sensor || '';

            // Datalogger lecturas → col H (¡Era H, no C!)
            const dl1 = safeParseFloat(s.dl1);
            const dl2 = safeParseFloat(s.dl2);
            const dl3 = safeParseFloat(s.dl3);
            if (dl1 !== null) ws.getCell(`H${r}`).value = dl1;
            if (dl2 !== null) ws.getCell(`H${r + 1}`).value = dl2;
            if (dl3 !== null) ws.getCell(`H${r + 2}`).value = dl3;
            // Patrón lecturas → col I
            const pt1 = safeParseFloat(s.pt1);
            const pt2 = safeParseFloat(s.pt2);
            const pt3 = safeParseFloat(s.pt3);
            if (pt1 !== null) ws.getCell(`I${r}`).value = pt1;
            if (pt2 !== null) ws.getCell(`I${r + 1}`).value = pt2;
            if (pt3 !== null) ws.getCell(`I${r + 2}`).value = pt3;

            // Fórmulas en columnas J-N usando H e I para lecturas, y C/K para parámetros
            ws.getCell(`J${r}`).value = { formula: `IF(ABS(AVERAGE(H${r}:H${r + 2})-AVERAGE(I${r}:I${r + 2}))<=$K$32,"Paso","Fallo")` };
            // Fórmula Uc (restaurada a $I$34, $I$40, $I$32 para coincidir exactamente con el cálculo matemático original de 2.2608)
            ws.getCell(`K${r}`).value = { formula: `2*SQRT(_xlfn.STDEV.S(H${r}:H${r + 2})^2+_xlfn.STDEV.S(I${r}:I${r + 2})^2+0.084*$I$34^2+0.084*$I$40^2+0.25*$I$32^2)` };
            ws.getCell(`L${r}`).value = { formula: `ABS(AVERAGE(H${r}:H${r + 2})-AVERAGE(I${r}:I${r + 2}))` };
            ws.getCell(`M${r}`).value = { formula: `ABS(K${r}+L${r})` };
            // Se calcula como PROMEDIO(I)-PROMEDIO(H) porque AVERAGE(I-H) inyectado por script calcula solo la primera fila (-1.91)
            ws.getCell(`N${r}`).value = { formula: `AVERAGE(I${r}:I${r + 2})-AVERAGE(H${r}:H${r + 2})` };
        });

        const out = await wb.xlsx.writeBuffer();
        return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }

    function formatDate(s) {
        if (!s) return '';
        return new Date(s + 'T00:00:00').toLocaleDateString('es-ES');
    }

    async function extractInstrumentsFromExcel(blob) {
        try {
            const arrayBuffer = await blob.arrayBuffer();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            const worksheet = workbook.getWorksheet('Certificado') || workbook.worksheets[0];

            if (!worksheet) return [];

            const instruments = [];
            const startRow = 12; // Datalogger usa fila 12 en vez de 17
            const maxRows = 4;

            for (let i = 0; i < maxRows; i++) {
                const rowIdx = startRow + i;
                const name = worksheet.getCell(`A${rowIdx}`).value;

                // Si la celda A está vacía, es cabecera o "Estado...", asumimos que no hay más instrumentos
                if (!name || (typeof name === 'string' && name.trim() === '')) break;
                const lowerName = String(name).toLowerCase().trim();

                // Si la fila en sí parece estar "vacía" de datos útiles o tiene cabeceras como 'Estado', la ignoramos.
                if (lowerName.includes('estado de') || lowerName.includes('comentarios')) break;
                if (lowerName.includes('instrumental') || lowerName.includes('patrón')) continue; // Ignorar si es el título cabecera "Instrumental patrón"

                // Extrayendo atributos reales
                const instName = String(name).trim();
                const brand = String(worksheet.getCell(`B${rowIdx}`).value || '').trim();
                const model = String(worksheet.getCell(`D${rowIdx}`).value || '').trim();
                const serie = String(worksheet.getCell(`E${rowIdx}`).value || '').trim();
                const date = String(worksheet.getCell(`F${rowIdx}`).value || '').trim();

                // Solo si tiene algo de validez (excluir "N/A" solitarios sin nombre real)
                if (instName.length > 0 && !instName.startsWith('N/A')) {
                    instruments.push({ name: instName, brand, model, serie, date });
                }
            }
            return instruments;
        } catch (err) {
            console.error("Error en extractInstrumentsFromExcel:", err);
            return [];
        }
    }

    init();
});
