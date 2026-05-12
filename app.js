// ============================================
// app.js - UC Case Registration App Main Logic
// ============================================

const App = {
  currentStep: 1,
  totalSteps: 6,
  editingId: null,
  pendingAction: null,
  cases: [],
  STORAGE_KEY: 'uc_case_registry',

  // ---- Initialization ----
  async init() {
    this.loadTheme();
    await this.initStorage();
    this.renderStepIndicator();
    this.renderAllSteps();
    this.goToStep(1);
    this.renderCaseList();
  },

  // ---- Theme Toggle ----
  loadTheme() {
    const saved = localStorage.getItem('uc_theme') || 'dark';
    this.applyTheme(saved);
  },

  toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'dark';
    const next = current === 'dark' ? 'light' : 'dark';
    this.applyTheme(next);
    localStorage.setItem('uc_theme', next);
  },

  applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    const btn = document.getElementById('btnTheme');
    if (btn) btn.textContent = theme === 'dark' ? '☀️' : '🌙';
  },

  // ---- Step Indicator ----
  renderStepIndicator() {
    const el = document.getElementById('stepIndicator');
    let html = '';
    STEPS.forEach((s, i) => {
      if (i > 0) html += `<div class="step-connector" id="conn-${i}"></div>`;
      html += `<div class="step-item" id="stepItem-${s.id}" onclick="App.goToStep(${s.id})">
        <div class="step-item__circle">${s.id}</div>
        <div class="step-item__label">${s.label}</div>
      </div>`;
    });
    el.innerHTML = html;
  },

  // ---- Render All Form Steps ----
  renderAllSteps() {
    for (let step = 1; step <= this.totalSteps; step++) {
      const container = document.getElementById(`step-${step}`);
      const stepInfo = STEPS[step - 1];
      const groups = FIELD_CONFIG[step];
      let html = `<div class="section-header">
        <h1 class="section-header__title">${stepInfo.title}</h1>
        <p class="section-header__desc">${stepInfo.desc}</p>
      </div>`;
      groups.forEach(g => {
        const langLabel = this.voiceLang === 'ja-JP' ? '🇯🇵 JA' : '🇺🇸 EN';
        const micBtn = g.voicePrefix
          ? `<button type="button" class="btn-voice-lang" onclick="App.toggleVoiceLang()" title="Switch voice language"><span class="voice-lang-label">${langLabel}</span></button><button type="button" class="btn-mic" id="mic-${g.voicePrefix}" onclick="App.startVoiceInput('${g.voicePrefix}')" title="Voice input for lab data">🎤</button>`
          : '';
        html += `<div class="field-group"><div class="field-group__title">${g.group}${micBtn}</div><div class="field-group__grid">`;
        g.fields.forEach(fd => { html += this.renderField(fd); });
        html += '</div></div>';
      });
      container.innerHTML = html;
    }
    // Attach listeners for auto-calc
    this.attachCalcListeners();
  },

  // ---- Render Single Field ----
  renderField(fd) {
    const cls = fd.type === 'textarea' ? 'form-field form-field--full' : 'form-field';
    const reqMark = fd.required ? '<span class="required">*</span>' : '';
    const autoMark = fd.auto ? '<span class="auto-calc">Auto</span>' : '';
    const unitMark = fd.unit ? `<span class="unit">(${fd.unit})</span>` : '';
    const hintMark = fd.hint ? `<span class="hint">${fd.hint}</span>` : '';
    let input = '';

    switch (fd.type) {
      case 'text':
        input = `<input type="text" id="${fd.key}" name="${fd.key}" ${fd.auto?'readonly class="auto-calculated"':''}>`;
        break;
      case 'number':
        input = fd.auto
          ? `<input type="text" id="${fd.key}" name="${fd.key}" readonly class="auto-calculated" tabindex="-1">`
          : `<input type="number" id="${fd.key}" name="${fd.key}" step="any">`;
        break;
      case 'date':
        input = fd.auto
          ? `<input type="date" id="${fd.key}" name="${fd.key}" readonly class="auto-calculated" tabindex="-1">`
          : `<input type="date" id="${fd.key}" name="${fd.key}">`;
        break;
      case 'select':
        input = `<select id="${fd.key}" name="${fd.key}"><option value="">— Select —</option>`;
        (fd.opts || []).forEach(o => { input += `<option value="${o[0]}">${o[1]}</option>`; });
        input += '</select>';
        break;
      case 'radio':
        input = `<div class="radio-group">`;
        (fd.opts || []).forEach(o => {
          input += `<label class="radio-option"><input type="radio" name="${fd.key}" value="${o[0]}"> ${o[1]}</label>`;
        });
        input += '</div>';
        break;
      case 'textarea':
        input = `<textarea id="${fd.key}" name="${fd.key}" rows="3"></textarea>`;
        break;
    }
    return `<div class="${cls}"><label for="${fd.key}">${fd.label}${reqMark}${autoMark}${unitMark}${hintMark}</label>${input}</div>`;
  },

  // ---- Auto-calculation Listeners ----
  attachCalcListeners() {
    const listen = (ids, fn) => {
      ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', fn);
      });
    };
    // Age
    listen(['date_of_birth', 'ADT_start'], () => this.runCalc('calcAge'));
    // BMI
    listen(['height', 'weight'], () => this.runCalc('calcBMI'));
    // Obs End / Baseline / Week dates
    const startEl = document.getElementById('ADT_start');
    if (startEl) startEl.addEventListener('input', () => {
      this.runCalc('calcObsEnd');
      this.runCalc('calcBlDate');
      this.runCalc('calcW12Date');
      this.runCalc('calcW26Date');
      this.runCalc('calcW52Date');
    });
    // Disease Duration
    listen(['disease_diagnosed', 'ADT_start'], () => this.runCalc('calcDiseaseDuration'));
    // Bio count
    ['history_ADT_1st','history_ADT_2nd','history_ADT_3rd','history_ADT_4th','history_ADT_5th'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.addEventListener('change', () => this.runCalc('calcBioCount'));
    });
    // PRO-2 for each timepoint
    ['bl','w12','w26','w52'].forEach(p => {
      listen([`${p}_SFS`, `${p}_RBS`], () => this.calcPRO2(p));
    });
    // ANC / ALC for each timepoint
    ['bl','w12','w26','w52'].forEach(p => {
      listen([`${p}_WBC`, `${p}_neutro_pct`], () => this.calcANC(p));
      listen([`${p}_WBC`, `${p}_lymph_pct`], () => this.calcALC(p));
      // Also recalc ALC when WBC changes
      const wbcEl = document.getElementById(`${p}_WBC`);
      if (wbcEl) wbcEl.addEventListener('input', () => this.calcALC(p));
    });
  },

  runCalc(fn) {
    const calcs = {
      calcAge: () => {
        const dob = document.getElementById('date_of_birth')?.value;
        const start = document.getElementById('ADT_start')?.value;
        if (dob && start) {
          const d1 = new Date(dob), d2 = new Date(start);
          let age = d2.getFullYear() - d1.getFullYear();
          const m = d2.getMonth() - d1.getMonth();
          if (m < 0 || (m === 0 && d2.getDate() < d1.getDate())) age--;
          this.setVal('age', age);
        }
      },
      calcBMI: () => {
        const h = parseFloat(this.getVal('height'));
        const w = parseFloat(this.getVal('weight'));
        if (h > 0 && w > 0) this.setVal('bmi', (w / ((h/100)**2)).toFixed(1));
      },
      calcObsEnd: () => {
        const d = this.getVal('ADT_start');
        if (d) { const dt = new Date(d); dt.setDate(dt.getDate() + 52*7); this.setVal('observation_end', this.fmtDate(dt)); }
      },
      calcBlDate: () => { const d = this.getVal('ADT_start'); if (d) this.setVal('bl_date', d); },
      calcW12Date: () => {
        const d = this.getVal('ADT_start');
        if (d) { const dt = new Date(d); dt.setDate(dt.getDate() + 12*7); this.setVal('w12_date', this.fmtDate(dt)); }
      },
      calcW26Date: () => {
        const d = this.getVal('ADT_start');
        if (d) { const dt = new Date(d); dt.setDate(dt.getDate() + 26*7); this.setVal('w26_date', this.fmtDate(dt)); }
      },
      calcW52Date: () => {
        const d = this.getVal('ADT_start');
        if (d) { const dt = new Date(d); dt.setDate(dt.getDate() + 52*7); this.setVal('w52_date', this.fmtDate(dt)); }
      },
      calcDiseaseDuration: () => {
        const diag = this.getVal('disease_diagnosed');
        const start = this.getVal('ADT_start');
        if (diag && start) {
          const d1 = new Date(diag), d2 = new Date(start);
          const months = (d2.getFullYear() - d1.getFullYear()) * 12 + d2.getMonth() - d1.getMonth();
          this.setVal('disease_duration', (months / 12).toFixed(1));
        }
      },
      calcBioCount: () => {
        let count = 0;
        ['history_ADT_1st','history_ADT_2nd','history_ADT_3rd','history_ADT_4th','history_ADT_5th'].forEach(id => {
          const v = this.getVal(id);
          if (v && v !== '' && v !== 'N/A') count++;
        });
        this.setVal('num_biologic_history', count);
      },
    };
    if (calcs[fn]) calcs[fn]();
  },

  calcPRO2(prefix) {
    const sfs = this.getSelectedVal(`${prefix}_SFS`);
    const rbs = this.getSelectedVal(`${prefix}_RBS`);
    if (sfs !== '' && rbs !== '') {
      this.setVal(`${prefix}_PRO2`, parseInt(sfs) + parseInt(rbs));
    }
  },

  calcANC(prefix) {
    const wbc = parseFloat(this.getVal(`${prefix}_WBC`));
    const pct = parseFloat(this.getVal(`${prefix}_neutro_pct`));
    if (!isNaN(wbc) && !isNaN(pct)) this.setVal(`${prefix}_ANC`, Math.round(wbc * pct / 100));
  },

  calcALC(prefix) {
    const wbc = parseFloat(this.getVal(`${prefix}_WBC`));
    const pct = parseFloat(this.getVal(`${prefix}_lymph_pct`));
    if (!isNaN(wbc) && !isNaN(pct)) this.setVal(`${prefix}_ALC`, Math.round(wbc * pct / 100));
  },

  // ---- Value Helpers ----
  getVal(id) { const el = document.getElementById(id); return el ? el.value : ''; },
  setVal(id, val) { const el = document.getElementById(id); if (el) el.value = val; },
  getSelectedVal(id) {
    const el = document.getElementById(id);
    if (!el) return '';
    if (el.tagName === 'SELECT') return el.value;
    // for radio in select
    return el.value;
  },
  getRadioVal(name) {
    const checked = document.querySelector(`input[name="${name}"]:checked`);
    return checked ? checked.value : '';
  },
  setRadioVal(name, val) {
    const radios = document.querySelectorAll(`input[name="${name}"]`);
    radios.forEach(r => { r.checked = (r.value === String(val)); });
  },
  fmtDate(d) {
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  },

  // ---- Step Navigation ----
  goToStep(n) {
    if (n < 1 || n > this.totalSteps) return;
    this.currentStep = n;
    // Update sections
    document.querySelectorAll('.form-section').forEach(s => s.classList.remove('active'));
    document.getElementById(`step-${n}`).classList.add('active');
    // Update indicator
    for (let i = 1; i <= this.totalSteps; i++) {
      const item = document.getElementById(`stepItem-${i}`);
      item.classList.remove('active', 'completed');
      if (i === n) item.classList.add('active');
      else if (i < n) item.classList.add('completed');
    }
    for (let i = 1; i < this.totalSteps; i++) {
      const conn = document.getElementById(`conn-${i}`);
      if (conn) { conn.classList.toggle('completed', i < n); }
    }
    // Update nav buttons
    document.getElementById('btnPrev').style.display = n === 1 ? 'none' : '';
    document.getElementById('btnNext').classList.toggle('hidden', n === this.totalSteps);
    document.getElementById('btnSubmit').classList.toggle('hidden', n !== this.totalSteps);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  },

  nextStep() { this.goToStep(this.currentStep + 1); },
  prevStep() { this.goToStep(this.currentStep - 1); },

  // ---- View Switching ----
  switchView(view) {
    document.getElementById('tabForm').classList.toggle('active', view === 'form');
    document.getElementById('tabList').classList.toggle('active', view === 'list');
    document.getElementById('formView').style.display = view === 'form' ? '' : 'none';
    document.getElementById('listView').classList.toggle('active', view === 'list');
    if (view === 'list') this.renderCaseList();
  },

  // ---- Collect Form Data ----
  collectFormData() {
    const data = {};
    // Iterate all fields
    for (let step = 1; step <= this.totalSteps; step++) {
      const groups = FIELD_CONFIG[step];
      groups.forEach(g => {
        g.fields.forEach(fd => {
          if (fd.type === 'radio') {
            data[fd.key] = this.getRadioVal(fd.key);
          } else {
            data[fd.key] = this.getVal(fd.key);
          }
        });
      });
    }
    return data;
  },

  // ---- Load Data into Form ----
  loadFormData(data) {
    for (let step = 1; step <= this.totalSteps; step++) {
      const groups = FIELD_CONFIG[step];
      groups.forEach(g => {
        g.fields.forEach(fd => {
          const val = data[fd.key];
          if (val === undefined || val === null) return;
          if (fd.type === 'radio') {
            this.setRadioVal(fd.key, val);
          } else {
            this.setVal(fd.key, val);
          }
        });
      });
    }
    // Trigger auto-calculations
    ['calcAge','calcBMI','calcObsEnd','calcBlDate','calcW12Date','calcW26Date','calcW52Date','calcDiseaseDuration','calcBioCount'].forEach(fn => this.runCalc(fn));
    ['bl','w12','w26','w52'].forEach(p => { this.calcPRO2(p); this.calcANC(p); this.calcALC(p); });
  },

  // ---- Clear Form ----
  clearForm() {
    document.getElementById('caseForm').reset();
    // Clear auto-calc fields
    document.querySelectorAll('.auto-calculated').forEach(el => { el.value = ''; });
    this.editingId = null;
    this.goToStep(1);
  },

  // ---- Storage (File System Access API + localStorage fallback) ----
  fileHandle: null,

  async initStorage() {
    // Load from localStorage first (immediate)
    try { this.cases = JSON.parse(localStorage.getItem(this.STORAGE_KEY)) || []; }
    catch { this.cases = []; }
    this.updateFileStatus();
  },

  // Link to a local JSON file (new or existing)
  async linkFile() {
    if (!('showSaveFilePicker' in window)) {
      this.toast('File save is not supported in this browser. Please use Chrome or Edge.', 'error');
      return;
    }
    try {
      const handle = await window.showSaveFilePicker({
        suggestedName: 'cases.json',
        types: [{
          description: 'JSON Data',
          accept: { 'application/json': ['.json'] },
        }],
      });
      this.fileHandle = handle;

      // Try to read existing data from the file
      try {
        const file = await handle.getFile();
        const text = await file.text();
        if (text.trim()) {
          const existing = JSON.parse(text);
          if (Array.isArray(existing) && existing.length > 0) {
            // Merge: existing file data takes priority
            this.cases = existing;
            localStorage.setItem(this.STORAGE_KEY, JSON.stringify(this.cases));
            this.renderCaseList();
            this.toast(`File linked (loaded ${existing.length} case(s))`, 'success');
            this.updateFileStatus();
            return;
          }
        }
      } catch (e) {
        // New file or empty - that's fine
      }

      // Write current data to the new file
      await this.writeToFile();
      this.toast(`File linked: ${handle.name}`, 'success');
      this.updateFileStatus();
    } catch (e) {
      if (e.name !== 'AbortError') {
        this.toast('Failed to link file: ' + e.message, 'error');
      }
    }
  },

  // Load from a JSON file
  async loadFromFile() {
    if (!('showOpenFilePicker' in window)) {
      this.toast('File loading is not supported in this browser. Please use Chrome or Edge.', 'error');
      return;
    }
    try {
      const [handle] = await window.showOpenFilePicker({
        types: [{
          description: 'JSON Data',
          accept: { 'application/json': ['.json'] },
        }],
      });
      const file = await handle.getFile();
      const text = await file.text();
      const data = JSON.parse(text);
      if (!Array.isArray(data)) throw new Error('Invalid format');

      this.fileHandle = handle;
      this.cases = data;
      localStorage.setItem(this.STORAGE_KEY, JSON.stringify(this.cases));
      this.renderCaseList();
      this.toast(`Loaded ${data.length} case(s) from: ${handle.name}`, 'success');
      this.updateFileStatus();
    } catch (e) {
      if (e.name !== 'AbortError') {
        this.toast('Failed to load file: ' + e.message, 'error');
      }
    }
  },

  // Write data to the linked file
  async writeToFile() {
    if (!this.fileHandle) return;
    try {
      const writable = await this.fileHandle.createWritable();
      await writable.write(JSON.stringify(this.cases, null, 2));
      await writable.close();
    } catch (e) {
      console.warn('File write failed:', e);
      // Permission might have been revoked
      if (e.name === 'NotAllowedError') {
        this.fileHandle = null;
        this.updateFileStatus();
        this.toast('File write permission lost. Please re-link the file.', 'error');
      }
    }
  },

  updateFileStatus() {
    const el = document.getElementById('fileStatus');
    if (!el) return;
    if (this.fileHandle) {
      el.innerHTML = `<span class="file-status__icon">📁</span><span class="file-status__name">${this.fileHandle.name}</span>`;
      el.classList.add('linked');
    } else {
      el.innerHTML = `<span class="file-status__icon">⚠️</span><span class="file-status__name">Not linked (localStorage)</span>`;
      el.classList.remove('linked');
    }
  },

  getCases() {
    return this.cases;
  },

  saveCases(cases) {
    this.cases = cases;
    localStorage.setItem(this.STORAGE_KEY, JSON.stringify(cases));
    // Auto-save to linked file
    this.writeToFile();
  },

  // ---- Save Draft ----
  saveDraft() {
    const data = this.collectFormData();
    if (!data.patient_id && !data.study_id) {
      this.toast('Please enter at least a Patient ID or Study ID.', 'error');
      return;
    }
    const cases = this.getCases();
    if (this.editingId !== null) {
      const idx = cases.findIndex(c => c._id === this.editingId);
      if (idx >= 0) { data._id = this.editingId; data._updated = new Date().toISOString(); cases[idx] = data; }
    } else {
      data._id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
      data._created = new Date().toISOString();
      cases.push(data);
      this.editingId = data._id;
    }
    this.saveCases(cases);
    this.toast('Draft saved successfully!', 'success');
  },

  // ---- Submit Case ----
  submitCase() {
    const data = this.collectFormData();
    // Basic validation
    if (!data.ADT_name) { this.toast('ADT Name is required.', 'error'); this.goToStep(1); return; }
    if (!data.ADT_start) { this.toast('ADT Start date is required.', 'error'); this.goToStep(1); return; }

    const cases = this.getCases();
    if (this.editingId !== null) {
      const idx = cases.findIndex(c => c._id === this.editingId);
      if (idx >= 0) { data._id = this.editingId; data._updated = new Date().toISOString(); data._created = cases[idx]._created; cases[idx] = data; }
    } else {
      data._id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
      data._created = new Date().toISOString();
      cases.push(data);
    }
    this.saveCases(cases);
    this.toast('Case registered successfully! ✓', 'success');
    this.clearForm();
    this.switchView('list');
  },

  // ---- Case List ----
  renderCaseList() {
    const cases = this.getCases();
    const tbody = document.getElementById('caseTableBody');
    const empty = document.getElementById('emptyState');
    const table = document.querySelector('.data-table');
    const countEl = document.getElementById('caseCount');
    const query = (document.getElementById('searchInput')?.value || '').toLowerCase();

    const filtered = query
      ? cases.filter(c => (c.patient_id || '').toLowerCase().includes(query) || (c.ADT_name || '').toLowerCase().includes(query) || (c.study_id || '').toString().includes(query))
      : cases;

    countEl.textContent = filtered.length;

    if (filtered.length === 0) {
      table.style.display = 'none';
      empty.classList.remove('hidden');
      tbody.innerHTML = '';
      return;
    }
    table.style.display = '';
    empty.classList.add('hidden');

    tbody.innerHTML = filtered.map((c, i) => {
      const adtClass = (c.ADT_name || '').toLowerCase();
      const sexLabel = c.sex === '0' ? 'M' : c.sex === '1' ? 'F' : '–';
      const extentMap = {'1':'E1','2':'E2','3':'E3'};
      return `<tr onclick="App.editCase('${c._id}')">
        <td>${i + 1}</td>
        <td>${c.patient_id || c.study_id || '–'}</td>
        <td><span class="badge badge--${adtClass}">${c.ADT_name || '–'}</span></td>
        <td>${c.ADT_start || '–'}</td>
        <td>${c.age || '–'}</td>
        <td>${sexLabel}</td>
        <td>${extentMap[c.disease_extent] || '–'}</td>
        <td>${c.bl_PRO2 || '–'}</td>
        <td class="action-btns" onclick="event.stopPropagation()">
          <button class="btn btn--secondary btn--sm" onclick="App.editCase('${c._id}')">Edit</button>
          <button class="btn btn--danger btn--sm" onclick="App.confirmDelete('${c._id}')">Delete</button>
        </td>
      </tr>`;
    }).join('');
  },

  filterCases() { this.renderCaseList(); },

  // ---- Edit Case ----
  editCase(id) {
    const cases = this.getCases();
    const c = cases.find(x => x._id === id);
    if (!c) return;
    this.clearForm();
    this.editingId = id;
    this.loadFormData(c);
    this.switchView('form');
    this.goToStep(1);
    this.toast('Editing case: ' + (c.patient_id || c.study_id || id), 'info');
  },

  // ---- Delete Case ----
  confirmDelete(id) {
    this.pendingAction = () => {
      let cases = this.getCases();
      cases = cases.filter(c => c._id !== id);
      this.saveCases(cases);
      this.renderCaseList();
      this.toast('Case deleted.', 'info');
    };
    document.getElementById('modalTitle').textContent = 'Delete Case';
    document.getElementById('modalDesc').textContent = 'Are you sure you want to delete this case? This action cannot be undone.';
    document.getElementById('confirmModal').classList.add('visible');
  },
  confirmAction() {
    if (this.pendingAction) this.pendingAction();
    this.closeModal();
    this.pendingAction = null;
  },
  closeModal() { document.getElementById('confirmModal').classList.remove('visible'); },

  // ---- Helper: Build row for one case using EXCEL_COLUMN_MAP ----
  buildExcelRow(c) {
    return EXCEL_COLUMN_MAP.map(col => {
      if (col === null) return '';
      let v = c[col] || '';
      if (typeof v === 'string') {
        // Convert date format yyyy-mm-dd → yyyy/mm/dd for Excel
        if (/^\d{4}-\d{2}-\d{2}$/.test(v)) {
          v = v.replace(/-/g, '/');
        }
        // Sanitize: remove tabs and newlines that break TSV alignment
        v = v.replace(/\t/g, ' ').replace(/[\r\n]+/g, '; ');
      }
      return v;
    });
  },

  // ---- CSV Export (download file) ----
  exportCSV() {
    const cases = this.getCases();
    if (cases.length === 0) { this.toast('No cases to export.', 'error'); return; }
    const header = EXCEL_COLUMN_MAP.map(col => col || '').join(',');
    const rows = cases.map(c => {
      return this.buildExcelRow(c).map(v => {
        if (typeof v === 'string' && (v.includes(',') || v.includes('"') || v.includes('\n'))) {
          v = '"' + v.replace(/"/g, '""') + '"';
        }
        return v;
      }).join(',');
    });
    const csv = [header, ...rows].join('\n');
    const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `uc_cases_${new Date().toISOString().slice(0,10)}.csv`;
    a.click(); URL.revokeObjectURL(url);
    this.toast(`Exported ${cases.length} cases as CSV.`, 'success');
  },

  // ---- Copy for Excel (TSV to clipboard, paste directly into Excel) ----
  async copyForExcel() {
    const cases = this.getCases();
    if (cases.length === 0) { this.toast('No cases to copy.', 'error'); return; }
    const rows = cases.map(c => this.buildExcelRow(c).join('\t'));
    const tsv = rows.join('\n');
    try {
      await navigator.clipboard.writeText(tsv);
      this.toast(`${cases.length} cases copied! Paste into Excel Row 4.`, 'success');
    } catch (e) {
      // Fallback
      const ta = document.createElement('textarea');
      ta.value = tsv; ta.style.position = 'fixed'; ta.style.left = '-9999px';
      document.body.appendChild(ta); ta.select(); document.execCommand('copy');
      document.body.removeChild(ta);
      this.toast(`${cases.length} cases copied! Paste into Excel Row 4.`, 'success');
    }
  },

  // ---- Toast ----
  toast(msg, type = 'info') {
    const container = document.getElementById('toastContainer');
    const el = document.createElement('div');
    el.className = `toast toast--${type}`;
    const icons = { success: '✓', error: '✕', info: 'ℹ' };
    el.innerHTML = `<span>${icons[type] || ''}</span> ${msg}`;
    container.appendChild(el);
    setTimeout(() => el.remove(), 3000);
  },

  // ---- Voice Input for Lab Data ----
  activeRecognition: null,
  voiceLang: (navigator.language || 'en').startsWith('ja') ? 'ja-JP' : 'en-US',

  toggleVoiceLang() {
    this.voiceLang = this.voiceLang === 'ja-JP' ? 'en-US' : 'ja-JP';
    const label = this.voiceLang === 'ja-JP' ? '🇯🇵 JA' : '🇺🇸 EN';
    document.querySelectorAll('.voice-lang-label').forEach(el => { el.textContent = label; });
    this.toast(`Voice language: ${this.voiceLang === 'ja-JP' ? 'Japanese' : 'English'}`, 'info');
  },

  startVoiceInput(prefix) {
    const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!SpeechRecognition) {
      this.toast('Speech Recognition is not supported in this browser. Please use Chrome.', 'error');
      return;
    }

    // If already recording, stop
    const micBtn = document.getElementById(`mic-${prefix}`);
    if (this.activeRecognition) {
      this.activeRecognition.stop();
      this.activeRecognition = null;
      document.querySelectorAll('.btn-mic').forEach(b => b.classList.remove('recording'));
      return;
    }

    const recognition = new SpeechRecognition();
    recognition.lang = this.voiceLang;
    recognition.continuous = true;
    recognition.interimResults = true;
    this.activeRecognition = recognition;

    // UI feedback
    micBtn.classList.add('recording');
    const example = this.voiceLang === 'ja-JP'
      ? '🎤 Listening... (e.g. "CRP 5.2 アルブミン 3.8 白血球 6500")'
      : '🎤 Listening... (e.g. "CRP 5.2 albumin 3.8 white blood cell 6500")';
    this.toast(example, 'info');

    let finalTranscript = '';

    recognition.onresult = (event) => {
      let interim = '';
      for (let i = event.resultIndex; i < event.results.length; i++) {
        const t = event.results[i][0].transcript;
        if (event.results[i].isFinal) {
          finalTranscript += ' ' + t;
          this.parseLabValues(finalTranscript, prefix);
        } else {
          interim += t;
        }
      }
    };

    recognition.onerror = (event) => {
      console.error('Speech error:', event.error);
      if (event.error !== 'no-speech') {
        this.toast('Voice input error: ' + event.error, 'error');
      }
      micBtn.classList.remove('recording');
      this.activeRecognition = null;
    };

    recognition.onend = () => {
      micBtn.classList.remove('recording');
      if (finalTranscript.trim()) {
        this.parseLabValues(finalTranscript, prefix);
        this.toast('Voice input completed ✓', 'success');
      }
      this.activeRecognition = null;
    };

    recognition.start();

    // Auto-stop after 30 seconds
    setTimeout(() => {
      if (this.activeRecognition === recognition) {
        recognition.stop();
      }
    }, 30000);
  },

  parseLabValues(text, prefix) {
    // Normalize text
    const t = text
      .replace(/\s+/g, ' ')
      .replace(/。/g, ' ')
      .replace(/、/g, ' ')
      .replace(/,/g, ' ')
      .replace(/:/g, ' ')
      .toLowerCase();

    // Keyword → field key mapping
    // Patterns include both English and Japanese terms for bilingual support
    const mappings = [
      {
        patterns: ['crp', 'c-reactive', 'c reactive', 'シーアールピー', 'しーあーるぴー', 'c反応性', 'シー・アール・ピー'],
        field: `${prefix}_CRP`,
        label: 'CRP',
      },
      {
        patterns: ['fcal', 'f-cal', 'fecal calprotectin', 'calprotectin', 'エフキャル', 'えふきゃる', 'カルプロテクチン', 'かるぷろてくちん', 'fキャル'],
        field: `${prefix}_fCAL`,
        label: 'fCAL',
      },
      {
        patterns: ['albumin', 'alb', 'アルブミン', 'あるぶみん'],
        field: `${prefix}_Alb`,
        label: 'Alb',
      },
      {
        patterns: ['white blood cell', 'white cell', 'wbc', 'leukocyte', '白血球', 'はっけっきゅう', 'ダブリュービーシー'],
        field: `${prefix}_WBC`,
        label: 'WBC',
      },
      {
        patterns: ['hemoglobin', 'haemoglobin', 'hgb', 'hb', 'ヘモグロビン', 'へもぐろびん'],
        field: `${prefix}_Hb`,
        label: 'Hb',
      },
      {
        patterns: ['platelet', 'plt', 'thrombocyte', '血小板', 'けっしょうばん'],
        field: `${prefix}_Plt`,
        label: 'Plt',
      },
      {
        patterns: ['neutrophil', 'neutro', 'neut', '好中球', 'こうちゅうきゅう', 'ニュートロ', 'ニュートロフィル'],
        field: `${prefix}_neutro_pct`,
        label: 'Neutro%',
      },
      {
        patterns: ['lymphocyte', 'lymph', 'リンパ球', 'リンパ', 'りんぱ'],
        field: `${prefix}_lymph_pct`,
        label: 'Lymph%',
      },
    ];

    let filled = [];

    mappings.forEach(m => {
      for (const pattern of m.patterns) {
        // Find pattern followed by a number (with optional filler words like "of", "is", "at")
        const regex = new RegExp(pattern.toLowerCase() + '[\\s]*(of|at|is|was)?[\\s]*([\\d]+\\.?[\\d]*)', 'i');
        const match = t.match(regex);
        if (match) {
          const val = parseFloat(match[2]);
          if (!isNaN(val)) {
            this.setVal(m.field, val);
            filled.push(m.label);
            // Trigger auto-calc for ANC/ALC
            if (m.field.includes('neutro_pct') || m.field.includes('WBC')) {
              this.calcANC(prefix);
            }
            if (m.field.includes('lymph_pct') || m.field.includes('WBC')) {
              this.calcALC(prefix);
            }
          }
          break;
        }
      }
    });

    if (filled.length > 0) {
      this.toast(`Filled: ${filled.join(', ')}`, 'success');
    }
  },
};

// ---- Boot ----
document.addEventListener('DOMContentLoaded', () => App.init());
