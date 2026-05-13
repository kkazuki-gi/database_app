// ============================================
// app.js - UC Case Registration App Main Logic
// ============================================

const I18N = {
  en: {
    langCode: 'en',
    langButton: '日本語',
    subtitle: 'Completeness-controlled data collection',
    newCase: '＋ New Case',
    caseList: '📋 Case List',
    linkFile: '📁 Link File',
    load: '📂 Load',
    copyExcel: '📋 Copy for Excel',
    exportCsv: '⬇ Export CSV',
    previous: '← Previous',
    saveDraft: '💾 Save Draft',
    next: 'Next →',
    register: '✓ Register Case',
    searchPlaceholder: 'Search by Patient ID, ADT...',
    total: 'Total:',
    cases: 'cases',
    noCases: 'No cases registered yet',
    noCasesDesc: 'Click "+ New Case" to register your first case.',
    notLinked: 'Not linked (localStorage)',
    select: '— Select —',
    auto: 'Auto',
    confirmMissing: 'Confirm missing',
    qualityOk: 'Completeness check passed. All tracked fields are either entered or confirmed as intentionally missing.',
    qualityWarn: count => `${count} unresolved missing-data check(s). Enter a value or check "Confirm missing" before registration.`,
    more: count => `+${count} more on this step`,
    checks: count => `${count} checks`,
    edit: 'Edit',
    delete: 'Delete',
    cancel: 'Cancel',
    deleteCase: 'Delete Case',
    deleteConfirm: 'Are you sure you want to delete this case? This action cannot be undone.',
    tableHeaders: ['No.', 'Patient ID', 'ADT', 'ADT Start', 'Age', 'Sex', 'Disease Extent', 'Baseline PRO-2', 'Quality', 'Actions'],
  },
  ja: {
    langCode: 'ja',
    langButton: 'English',
    subtitle: '入力漏れ確認つきデータ収集',
    newCase: '＋ 新規症例',
    caseList: '📋 症例一覧',
    linkFile: '📁 保存先',
    load: '📂 読み込み',
    copyExcel: '📋 Excel用コピー',
    exportCsv: '⬇ CSV出力',
    previous: '← 前へ',
    saveDraft: '💾 下書き保存',
    next: '次へ →',
    register: '✓ 症例登録',
    searchPlaceholder: 'Patient ID、ADTで検索...',
    total: '合計:',
    cases: '例',
    noCases: '登録症例はまだありません',
    noCasesDesc: '「＋ 新規症例」から最初の症例を登録してください。',
    notLinked: '未連携 (localStorage)',
    select: '— 選択 —',
    auto: '自動',
    confirmMissing: '未入力確認',
    qualityOk: '完全性チェック完了。確認対象の項目はすべて入力済み、または意図的な未入力として確認済みです。',
    qualityWarn: count => `未解決の未入力確認が${count}件あります。登録前に値を入力するか、「未入力確認」をチェックしてください。`,
    more: count => `このステップにさらに${count}件`,
    checks: count => `${count}件`,
    edit: '編集',
    delete: '削除',
    cancel: 'キャンセル',
    deleteCase: '症例削除',
    deleteConfirm: 'この症例を削除しますか？この操作は元に戻せません。',
    tableHeaders: ['No.', 'Patient ID', 'ADT', 'ADT開始', '年齢', '性別', '病変範囲', 'Baseline PRO-2', '品質', '操作'],
  },
};

const STEP_JA = {
  1: { label: '基本情報', title: '患者基本情報', desc: '患者背景、治療歴、観察情報' },
  2: { label: 'Baseline', title: 'Baseline評価 (Week 0)', desc: '登録時の併用薬、PRO-2、検査、内視鏡、IUS' },
  3: { label: 'Week 12', title: 'Week 12評価', desc: '許容範囲: Week 10-14' },
  4: { label: 'Week 26', title: 'Week 26評価', desc: '許容範囲: Week 24-28' },
  5: { label: 'Week 52', title: 'Week 52評価', desc: 'Week 52に最も近い評価' },
  6: { label: 'TEAE', title: '治療下発現有害事象', desc: '安全性フォローアップ 30日以内' },
};

const GROUP_JA = {
  'Identification': '識別情報',
  'Patient Characteristics': '患者背景',
  'Disease Information': '疾患情報',
  'Treatment History': '治療歴',
  'Other': 'その他',
  'Concomitant Medications at Enrollment': '登録時併用薬',
  'PRO-2 (Patient Reported Outcomes)': 'PRO-2 (患者報告アウトカム)',
  'Laboratory (within 3 months before index date)': '血液/便検査 (index date前3か月以内)',
  'Endoscopy (closest to index date)': '内視鏡 (index dateに最も近い評価)',
  'IUS (Intestinal Ultrasound)': 'IUS (腸管超音波)',
  'Timing': '評価時期',
  'PRO-2': 'PRO-2',
  'Laboratory': '血液/便検査',
  'Endoscopy (Exploratory)': '内視鏡 (探索的)',
  'IUS (Exploratory)': 'IUS (探索的)',
  'Corticosteroid-free Status': 'ステロイドフリー状態',
  'Endoscopy': '内視鏡',
  'Major Adverse Events': '主要有害事象',
  'Hematologic AEs': '血液学的有害事象',
  'Drug-related Liver Dysfunction': '薬剤関連肝機能障害',
  'Cardiac AEs': '心臓関連有害事象',
  'Infectious AEs': '感染症有害事象',
};

const FIELD_JA = {
  study_id: 'Study ID',
  patient_id: 'Patient ID',
  ADT_name: 'ADT名',
  ADT_start: 'ADT開始日 (Week 0)',
  observation_end: '観察終了日 (Week 52)',
  discontinuation_date: '中止日',
  discontinuation_reason: '中止理由',
  date_of_birth: '生年月日',
  age: '登録時年齢',
  sex: '性別',
  height: '身長',
  weight: '体重',
  bmi: 'BMI',
  disease_extent: '病変範囲 (Montreal分類)',
  disease_diagnosed: '診断日',
  disease_duration: '罹病期間',
  smoking_history: '喫煙歴',
  history_UC_hospitalization: 'UC関連入院歴',
  history_immunomodulators: '免疫調節薬使用歴 (AZA, 6-MP)',
  history_corticosteroid: 'ステロイド使用歴',
  history_CNIs: 'CNI使用歴 (Tacrolimus, Cyclosporine)',
  num_biologic_history: '既使用生物学的製剤数',
  history_ADT_1st: '既使用ADT 1st',
  history_ADT_2nd: '既使用ADT 2nd',
  history_ADT_3rd: '既使用ADT 3rd',
  history_ADT_4th: '既使用ADT 4th',
  history_ADT_5th: '既使用ADT 5th',
  comorbidities: '併存疾患',
  comment: 'コメント',
  bl_date: 'Baseline日',
  bl_5ASA: '5-ASA使用 (SASP含む)',
  bl_corticosteroids: 'ステロイド使用',
  bl_immunomodulators: '免疫調節薬使用 (AZA, 6-MP)',
  bl_SFS: '排便回数サブスコア (SFS)',
  bl_RBS: '血便サブスコア (RBS)',
  bl_PRO2: 'PRO-2スコア',
  bl_lab_date: '評価日',
  bl_CRP: 'CRP',
  bl_fCAL: 'fCAL',
  bl_Alb: 'アルブミン',
  bl_WBC: '白血球数',
  bl_neutro_pct: '好中球',
  bl_ANC: '好中球実数',
  bl_lymph_pct: 'リンパ球',
  bl_ALC: 'リンパ球実数',
  bl_endo_date: '評価日',
  bl_MES: 'MES',
  bl_UCEIS: 'UCEIS (探索的)',
  bl_ius_date: '評価日',
  bl_IUS_BWT: 'IUS BWT (mm)',
  bl_IUS_doppler: 'IUS Color Doppler',
  w12_date: 'Week 12日',
  w26_date: 'Week 26日',
  w52_date: 'Week 52日',
  w12_pro2_date: '評価日 (PRO-2)',
  w26_pro2_date: '評価日 (PRO-2)',
  w52_pro2_date: '評価日 (PRO-2)',
  w12_SFS: 'SFS',
  w26_SFS: 'SFS',
  w52_SFS: 'SFS',
  w12_RBS: 'RBS',
  w26_RBS: 'RBS',
  w52_RBS: 'RBS',
  w12_PRO2: 'PRO-2スコア',
  w26_PRO2: 'PRO-2スコア',
  w52_PRO2: 'PRO-2スコア',
  w12_lab_date: '評価日',
  w26_lab_date: '評価日',
  w52_lab_date: '評価日',
  w12_CRP: 'CRP',
  w26_CRP: 'CRP',
  w52_CRP: 'CRP',
  w12_fCAL: 'fCAL',
  w26_fCAL: 'fCAL',
  w52_fCAL: 'fCAL',
  w12_Alb: 'アルブミン',
  w26_Alb: 'アルブミン',
  w52_Alb: 'アルブミン',
  w12_WBC: '白血球数',
  w26_WBC: '白血球数',
  w52_WBC: '白血球数',
  w12_neutro_pct: '好中球',
  w26_neutro_pct: '好中球',
  w52_neutro_pct: '好中球',
  w12_ANC: '好中球実数',
  w26_ANC: '好中球実数',
  w52_ANC: '好中球実数',
  w12_lymph_pct: 'リンパ球',
  w26_lymph_pct: 'リンパ球',
  w52_lymph_pct: 'リンパ球',
  w12_ALC: 'リンパ球実数',
  w26_ALC: 'リンパ球実数',
  w52_ALC: 'リンパ球実数',
  w12_endo_date: '評価日',
  w26_endo_date: '評価日',
  w52_endo_date: '評価日',
  w12_MES: 'MES',
  w26_MES: 'MES',
  w52_MES: 'MES',
  w12_UCEIS: 'UCEIS',
  w26_UCEIS: 'UCEIS',
  w52_UCEIS: 'UCEIS (探索的)',
  w12_ius_date: '評価日',
  w52_ius_date: '評価日',
  w12_IUS_BWT: 'IUS BWT (mm)',
  w52_IUS_BWT: 'IUS BWT (mm)',
  w12_IUS_doppler: 'IUS Color Doppler',
  w52_IUS_doppler: 'IUS Color Doppler',
  w26_CS_free: 'ステロイドフリー',
  w52_CS_free: 'ステロイドフリー',
  w52_histologic_remission: '組織学的寛解 (探索的)',
  w52_biopsy_site: '生検部位 (探索的)',
  ae_worsening_UC: 'UC悪化',
  ae_herpes_zoster: '帯状疱疹',
  ae_VTE: 'VTE',
  ae_MACE: 'MACE',
  ae_macular_edema: '黄斑浮腫',
  ae_malignancies: '悪性腫瘍 (種類)',
  ae_hematologic: '血液学的有害事象',
  ae_liver: '肝機能障害',
  ae_cardiac: '徐脈/心伝導障害',
  ae_infectious: '感染症有害事象',
  ae_other_details: 'その他の有害事象/詳細',
  other_info: 'その他重要情報',
};

const HINT_JA = {
  'If drop out before W52': 'Week 52前に脱落した場合',
  'Year only: yyyy/01/01': '年のみの場合: yyyy/01/01',
  'FIL only': 'FILのみ',
  'Greatest bowel wall thickening, excl. rectum': '直腸を除く最大腸管壁肥厚',
  'Without systemic CS use for ≥12 weeks before Week 26': 'Week 26前12週間以上、全身性ステロイドなし',
  'Without systemic CS use for ≥12 weeks before Week 52': 'Week 52前12週間以上、全身性ステロイドなし',
  'Free write: type of cancer': '自由記載: がん種',
  'e.g. For OZA: dates of interruption, re-initiation, re-titration': '例 OZA: 中断、再開、再漸増の日付',
};

const OPTION_JA = {
  'Male': '男性',
  'Female': '女性',
  'No': 'なし',
  'Yes': 'あり',
  'Never': 'なし',
  'Former': '過去',
  'Current': '現在',
  'Unknown': '不明',
  '0 – Normal': '0 - 正常',
  '1 – 1-2 more': '1 - 1-2回増加',
  '2 – 3-4 more': '2 - 3-4回増加',
  '3 – 5+ more': '3 - 5回以上増加',
  '0 – No blood': '0 - 血便なし',
  '1 – Streaks <50%': '1 - 血液付着 <50%',
  '2 – Obvious >50%': '2 - 明らかな血便 >50%',
  '3 – Blood alone': '3 - ほぼ血液のみ',
  'Signal (-)': 'シグナル (-)',
  'Signal (+)': 'シグナル (+)',
  'Lack of effectiveness': '効果不十分',
  'Disease relapse/worsening': '再燃/悪化',
  'Adverse event': '有害事象',
  'Patient preference': '患者希望',
  'Loss to follow-up': '追跡不能',
  'Transfer': '転院',
  'Pregnancy': '妊娠/妊娠希望',
  'Surgery (CRC)': 'UC関連大腸癌手術',
  'Death': '死亡',
  'Other': 'その他',
  'Rectum': '直腸',
  'None': 'なし',
  'Anemia: Hb<8.0': '貧血: Hb<8.0',
  'Platelet < 25,000': '血小板 < 25,000',
  'ALT ≧3×ULN': 'ALT 3×ULN以上',
  'ALT ≧5×ULN': 'ALT 5×ULN以上',
  'γGTP increased': 'γGTP上昇',
  'Bradycardia': '徐脈',
  'ECG abnormality': '心電図異常',
  'AV block 1st': '1度房室ブロック',
  'AV block 2nd (Mobitz I)': '2度房室ブロック (Mobitz I)',
  'Upper respiratory tract infection': '上気道感染',
  'Respiratory tract infection viral': 'ウイルス性呼吸器感染',
  'Serious infection': '重篤感染症',
  'Opportunistic infection': '日和見感染',
};

const COMPLETENESS_EXCLUDED_KEYS = new Set([
  'comment',
  'comorbidities',
  'ae_other_details',
  'other_info',
]);

const App = {
  currentStep: 1,
  totalSteps: 6,
  editingId: null,
  pendingAction: null,
  cases: [],
  lang: 'en',
  STORAGE_KEY: 'uc_case_registry_v2',

  // ---- Initialization ----
  async init() {
    this.loadTheme();
    this.loadLanguage();
    await this.initStorage();
    this.renderStepIndicator();
    this.renderAllSteps();
    this.applyLanguage();
    this.goToStep(1);
    this.renderCaseList();
    this.updateCompletenessStatus();
  },

  // ---- Language Toggle ----
  loadLanguage() {
    const saved = localStorage.getItem('uc_lang');
    this.lang = saved || ((navigator.language || 'en').startsWith('ja') ? 'ja' : 'en');
    document.documentElement.lang = this.t('langCode');
  },

  toggleLanguage() {
    const data = document.getElementById('caseForm') ? this.collectFormData() : null;
    const step = this.currentStep;
    this.lang = this.lang === 'ja' ? 'en' : 'ja';
    localStorage.setItem('uc_lang', this.lang);
    document.documentElement.lang = this.t('langCode');
    this.renderStepIndicator();
    this.renderAllSteps();
    if (data) this.loadFormData(data);
    this.applyLanguage();
    this.goToStep(step);
    this.renderCaseList();
    this.toast(this.lang === 'ja' ? '表示言語: 日本語' : 'Display language: English', 'info');
  },

  t(key) {
    return (I18N[this.lang] && I18N[this.lang][key]) || I18N.en[key] || key;
  },

  stepText(stepInfo) {
    return this.lang === 'ja' ? { ...stepInfo, ...(STEP_JA[stepInfo.id] || {}) } : stepInfo;
  },

  groupLabel(label) {
    return this.lang === 'ja' ? (GROUP_JA[label] || label) : label;
  },

  fieldLabel(fd) {
    return this.lang === 'ja' ? (FIELD_JA[fd.key] || fd.label) : fd.label;
  },

  hintLabel(hint) {
    return this.lang === 'ja' ? (HINT_JA[hint] || hint) : hint;
  },

  optionLabel(label) {
    return this.lang === 'ja' ? (OPTION_JA[label] || label) : label;
  },

  applyLanguage() {
    const setText = (id, text) => { const el = document.getElementById(id); if (el) el.textContent = text; };
    const setHtml = (id, text) => { const el = document.getElementById(id); if (el) el.innerHTML = text; };
    setText('appSubtitle', this.t('subtitle'));
    setText('tabForm', this.t('newCase'));
    setText('tabList', this.t('caseList'));
    setText('btnCopyExcel', this.t('copyExcel'));
    setText('btnExport', this.t('exportCsv'));
    setText('btnLang', this.t('langButton'));
    setText('btnPrev', this.t('previous'));
    setText('btnSaveDraft', this.t('saveDraft'));
    setText('btnNext', this.t('next'));
    setText('btnSubmit', this.t('register'));
    setText('caseCountLabel', this.t('total'));
    setText('caseCountUnit', this.t('cases'));
    setText('emptyTitle', this.t('noCases'));
    setText('emptyDesc', this.t('noCasesDesc'));
    setText('modalCancelBtn', this.t('cancel'));
    setHtml('thNo', this.t('tableHeaders')[0]);
    setHtml('thPatient', this.t('tableHeaders')[1]);
    setHtml('thAdt', this.t('tableHeaders')[2]);
    setHtml('thAdtStart', this.t('tableHeaders')[3]);
    setHtml('thAge', this.t('tableHeaders')[4]);
    setHtml('thSex', this.t('tableHeaders')[5]);
    setHtml('thExtent', this.t('tableHeaders')[6]);
    setHtml('thBlPro2', this.t('tableHeaders')[7]);
    setHtml('thQuality', this.t('tableHeaders')[8]);
    setHtml('thActions', this.t('tableHeaders')[9]);
    const search = document.getElementById('searchInput');
    if (search) search.placeholder = this.t('searchPlaceholder');
    document.querySelectorAll('.btn--file').forEach((btn, index) => {
      btn.textContent = index === 0 ? this.t('linkFile') : this.t('load');
    });
    this.updateFileStatus();
    this.updateCompletenessStatus();
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
      const stepText = this.stepText(s);
      if (i > 0) html += `<div class="step-connector" id="conn-${i}"></div>`;
      html += `<div class="step-item" id="stepItem-${s.id}" onclick="App.goToStep(${s.id})">
        <div class="step-item__circle">${s.id}</div>
        <div class="step-item__quality" id="stepQuality-${s.id}">OK</div>
        <div class="step-item__label">${stepText.label}</div>
      </div>`;
    });
    el.innerHTML = html;
  },

  // ---- Render All Form Steps ----
  renderAllSteps() {
    for (let step = 1; step <= this.totalSteps; step++) {
      const container = document.getElementById(`step-${step}`);
      const stepInfo = this.stepText(STEPS[step - 1]);
      const groups = FIELD_CONFIG[step];
      let html = `<div class="section-header">
        <h1 class="section-header__title">${stepInfo.title}</h1>
        <p class="section-header__desc">${stepInfo.desc}</p>
      </div>`;
      groups.forEach((g, groupIndex) => {
        const langLabel = this.voiceLang === 'ja-JP' ? '🇯🇵 JA' : '🇺🇸 EN';
        const micBtn = g.voicePrefix
          ? `<button type="button" class="btn-voice-lang" onclick="App.toggleVoiceLang()" title="Switch voice language"><span class="voice-lang-label">${langLabel}</span></button><button type="button" class="btn-mic" id="mic-${g.voicePrefix}" onclick="App.startVoiceInput('${g.voicePrefix}')" title="Voice input for lab data">🎤</button>`
          : '';
        html += `<div class="field-group" data-step="${step}" data-group-index="${groupIndex}"><div class="field-group__title">${this.groupLabel(g.group)}${micBtn}</div><div class="field-group__grid">`;
        g.fields.forEach(fd => { html += this.renderField(fd); });
        html += '</div></div>';
      });
      container.innerHTML = html;
    }
    // Attach listeners for auto-calc
    this.attachCalcListeners();
    this.attachQualityListeners();
  },

  // ---- Render Single Field ----
  renderField(fd) {
    const cls = fd.type === 'textarea' ? 'form-field form-field--full' : 'form-field';
    const reqMark = fd.required ? '<span class="required">*</span>' : '';
    const autoMark = fd.auto ? `<span class="auto-calc">${this.t('auto')}</span>` : '';
    const unitMark = fd.unit ? `<span class="unit">(${fd.unit})</span>` : '';
    const hintMark = fd.hint ? `<span class="hint">${this.hintLabel(fd.hint)}</span>` : '';
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
        if (fd.auto) {
          input = `<input type="date" id="${fd.key}" name="${fd.key}" readonly class="auto-calculated" tabindex="-1">`;
        } else if (fd.copyDateFrom) {
          input = `<div class="date-with-copy"><input type="date" id="${fd.key}" name="${fd.key}"><button type="button" class="btn-copy-date" id="copyBtn-${fd.key}" data-source="${fd.copyDateFrom}" data-target="${fd.key}" onclick="App.copyDateToField('${fd.copyDateFrom}','${fd.key}')" title="Copy week date"><span class="copy-date-icon">📅</span><span class="copy-date-val" id="copyVal-${fd.key}">—</span></button></div>`;
        } else {
          input = `<input type="date" id="${fd.key}" name="${fd.key}">`;
        }
        break;
      case 'select':
        input = `<select id="${fd.key}" name="${fd.key}"><option value="">${this.t('select')}</option>`;
        (fd.opts || []).forEach(o => { input += `<option value="${o[0]}">${this.optionLabel(o[1])}</option>`; });
        input += '</select>';
        break;
      case 'radio':
        input = `<div class="radio-group">`;
        (fd.opts || []).forEach(o => {
          input += `<label class="radio-option"><input type="radio" name="${fd.key}" value="${o[0]}"> ${this.optionLabel(o[1])}</label>`;
        });
        input += '</div>';
        break;
      case 'textarea':
        input = `<textarea id="${fd.key}" name="${fd.key}" rows="3"></textarea>`;
        break;
    }
    return `<div class="${cls}" data-field-key="${fd.key}"><label for="${fd.key}">${this.fieldLabel(fd)}${reqMark}${autoMark}${unitMark}${hintMark}</label>${input}${this.renderMissingControl(fd)}</div>`;
  },

  renderMissingControl(fd) {
    if (!this.shouldTrackCompleteness(fd)) return '';
    return `<div class="missing-control" id="missingWrap-${fd.key}">
      <label class="missing-control__check">
        <input type="checkbox" id="missingChk-${fd.key}" data-completeness-control="1" onchange="App.toggleMissingReason('${fd.key}')">
        ${this.t('confirmMissing')}
      </label>
    </div>`;
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
      // Update all copy-date buttons with new calculated dates
      setTimeout(() => this.updateCopyDateButtons(), 50);
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

  attachQualityListeners() {
    const form = document.getElementById('caseForm');
    if (!form) return;
    form.querySelectorAll('input, select, textarea').forEach(el => {
      if (el.dataset.completenessControl) return;
      const eventName = el.type === 'radio' || el.tagName === 'SELECT' ? 'change' : 'input';
      el.addEventListener(eventName, () => {
        this.handleFieldInput(el.name || el.id);
        this.updateCompletenessStatus();
      });
    });
  },

  getAllFieldDefs() {
    const fields = [];
    for (let step = 1; step <= this.totalSteps; step++) {
      FIELD_CONFIG[step].forEach(group => {
        group.fields.forEach(fd => fields.push({ ...fd, step, group: group.group }));
      });
    }
    return fields;
  },

  getFieldDef(key) {
    return this.getAllFieldDefs().find(fd => fd.key === key);
  },

  shouldTrackCompleteness(fd) {
    return Boolean(fd && !fd.auto && !COMPLETENESS_EXCLUDED_KEYS.has(fd.key));
  },

  getFieldValue(fd) {
    return fd.type === 'radio' ? this.getRadioVal(fd.key) : this.getVal(fd.key);
  },

  isValuePresent(value) {
    return value !== undefined && value !== null && String(value).trim() !== '';
  },

  getMissingMeta(key) {
    return {
      confirmed: document.getElementById(`missingChk-${key}`)?.checked || false,
    };
  },

  isFieldResolved(fd) {
    if (!this.shouldTrackCompleteness(fd)) return true;
    if (this.isValuePresent(this.getFieldValue(fd))) return true;
    const missing = this.getMissingMeta(fd.key);
    return missing.confirmed;
  },

  handleFieldInput(key) {
    const fd = this.getFieldDef(key);
    if (!fd || !this.shouldTrackCompleteness(fd)) return;
    if (this.isValuePresent(this.getFieldValue(fd))) this.clearMissingForField(key);
  },

  clearMissingForField(key) {
    const chk = document.getElementById(`missingChk-${key}`);
    if (chk) chk.checked = false;
  },

  toggleMissingReason(key) {
    const fd = this.getFieldDef(key);
    const chk = document.getElementById(`missingChk-${key}`);
    if (!chk) return;
    if (chk.checked && fd) {
      if (fd.type === 'radio') this.setRadioVal(fd.key, '');
      else this.setVal(fd.key, '');
    }
    this.updateCompletenessStatus();
  },

  getCompletenessIssues() {
    const issues = [];
    this.getAllFieldDefs().forEach(fd => {
      if (!this.shouldTrackCompleteness(fd)) return;
      if (!this.isFieldResolved(fd)) {
        issues.push({ step: fd.step, group: fd.group, key: fd.key, label: fd.label, status: 'unconfirmed_missing' });
      }
    });
    return issues;
  },

  getCaseCompletenessIssues(data) {
    const issues = [];
    this.getAllFieldDefs().forEach(fd => {
      if (!this.shouldTrackCompleteness(fd)) return;
      const value = data[fd.key];
      const confirmed = data[`_missing_${fd.key}`] === '1';
      if (!this.isValuePresent(value) && !confirmed) {
        issues.push({ step: fd.step, group: fd.group, key: fd.key, label: fd.label });
      }
    });
    return issues;
  },

  updateCompletenessStatus() {
    const issues = this.getCompletenessIssues();
    const counts = {};
    for (let i = 1; i <= this.totalSteps; i++) counts[i] = 0;
    issues.forEach(issue => { counts[issue.step] += 1; });

    Object.entries(counts).forEach(([step, count]) => {
      const el = document.getElementById(`stepQuality-${step}`);
      const item = document.getElementById(`stepItem-${step}`);
      if (!el || !item) return;
      el.textContent = count > 0 ? String(count) : 'OK';
      el.classList.toggle('has-issues', count > 0);
      item.classList.toggle('has-issues', count > 0);
    });

    document.querySelectorAll('[data-field-key]').forEach(el => el.classList.remove('field-issue'));
    issues.forEach(issue => {
      document.querySelector(`[data-field-key="${issue.key}"]`)?.classList.add('field-issue');
    });

    this.renderQualityPanel(issues);
    return issues;
  },

  renderQualityPanel(issues) {
    const panel = document.getElementById('qualityPanel');
    if (!panel) return;
    const currentIssues = issues.filter(issue => issue.step === this.currentStep);
    if (issues.length === 0) {
      panel.innerHTML = `<div class="quality-panel__ok">${this.t('qualityOk')}</div>`;
      panel.className = 'quality-panel quality-panel--ok';
      return;
    }
    const visible = currentIssues.slice(0, 6).map(issue => {
      const fd = this.getFieldDef(issue.key);
      return `<button type="button" class="quality-chip" onclick="App.focusField('${issue.key}')">${this.groupLabel(issue.group)}: ${fd ? this.fieldLabel(fd) : issue.label}</button>`;
    }).join('');
    const more = currentIssues.length > 6 ? `<span class="quality-panel__more">${this.t('more')(currentIssues.length - 6)}</span>` : '';
    panel.className = 'quality-panel quality-panel--warn';
    panel.innerHTML = `<div class="quality-panel__summary">${this.t('qualityWarn')(issues.length)}</div><div class="quality-panel__chips">${visible}${more}</div>`;
  },

  focusField(key) {
    const fd = this.getFieldDef(key);
    if (!fd) return;
    this.goToStep(fd.step);
    setTimeout(() => {
      const field = document.querySelector(`[data-field-key="${key}"]`);
      const input = document.getElementById(key) || document.querySelector(`input[name="${key}"]`) || document.getElementById(`missingChk-${key}`);
      if (field) field.scrollIntoView({ behavior: 'smooth', block: 'center' });
      if (input) input.focus({ preventScroll: true });
    }, 150);
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
  fmtDateShort(dateStr) {
    if (!dateStr) return '—';
    const d = new Date(dateStr);
    return `${d.getMonth()+1}/${d.getDate()}`;
  },

  // ---- Copy Date Helpers ----
  copyDateToField(sourceId, targetId) {
    const val = this.getVal(sourceId);
    if (!val) {
      this.toast(this.lang === 'ja' ? 'Week日がまだ計算されていません。先にADT開始日を入力してください。' : 'Week date not yet calculated. Enter ADT Start first.', 'error');
      return;
    }
    this.setVal(targetId, val);
    this.clearMissingForField(targetId);
    this.updateCompletenessStatus();
    // Flash animation on the target input
    const el = document.getElementById(targetId);
    if (el) {
      el.classList.add('date-copied');
      setTimeout(() => el.classList.remove('date-copied'), 600);
    }
    this.toast(this.lang === 'ja' ? `日付をコピーしました: ${val}` : `Date copied: ${val}`, 'success');
  },

  updateCopyDateButtons() {
    document.querySelectorAll('.btn-copy-date').forEach(btn => {
      const sourceId = btn.dataset.source;
      const targetId = btn.dataset.target;
      const valEl = document.getElementById(`copyVal-${targetId}`);
      const sourceVal = this.getVal(sourceId);
      if (valEl) {
        valEl.textContent = sourceVal ? this.fmtDateShort(sourceVal) : '—';
      }
    });
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
    this.updateCompletenessStatus();
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
          if (this.shouldTrackCompleteness(fd)) {
            const missing = this.getMissingMeta(fd.key);
            data[`_missing_${fd.key}`] = missing.confirmed ? '1' : '';
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
          if (val !== undefined && val !== null) {
            if (fd.type === 'radio') {
              this.setRadioVal(fd.key, val);
            } else {
              this.setVal(fd.key, val);
            }
          }
          if (this.shouldTrackCompleteness(fd)) {
            const chk = document.getElementById(`missingChk-${fd.key}`);
            if (chk) chk.checked = data[`_missing_${fd.key}`] === '1';
          }
        });
      });
    }
    // Trigger auto-calculations
    ['calcAge','calcBMI','calcObsEnd','calcBlDate','calcW12Date','calcW26Date','calcW52Date','calcDiseaseDuration','calcBioCount'].forEach(fn => this.runCalc(fn));
    ['bl','w12','w26','w52'].forEach(p => { this.calcPRO2(p); this.calcANC(p); this.calcALC(p); });
    this.updateCompletenessStatus();
    // Update copy-date buttons
    setTimeout(() => this.updateCopyDateButtons(), 50);
  },

  // ---- Clear Form ----
  clearForm() {
    document.getElementById('caseForm').reset();
    // Clear auto-calc fields
    document.querySelectorAll('.auto-calculated').forEach(el => { el.value = ''; });
    this.editingId = null;
    this.goToStep(1);
    this.updateCompletenessStatus();
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
      el.innerHTML = `<span class="file-status__icon">⚠️</span><span class="file-status__name">${this.t('notLinked')}</span>`;
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
      this.toast(this.lang === 'ja' ? 'Patient ID または Study ID を入力してください。' : 'Please enter at least a Patient ID or Study ID.', 'error');
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
    this.toast(this.lang === 'ja' ? '下書きを保存しました。' : 'Draft saved successfully!', 'success');
  },

  // ---- Submit Case ----
  submitCase() {
    const data = this.collectFormData();
    // Basic validation
    if (!data.ADT_name) { this.toast(this.lang === 'ja' ? 'ADT名は必須です。' : 'ADT Name is required.', 'error'); this.goToStep(1); return; }
    if (!data.ADT_start) { this.toast(this.lang === 'ja' ? 'ADT開始日は必須です。' : 'ADT Start date is required.', 'error'); this.goToStep(1); return; }
    const issues = this.updateCompletenessStatus();
    if (issues.length > 0) {
      this.toast(this.lang === 'ja' ? `登録前に${issues.length}件の未入力確認が必要です。` : `${issues.length} missing-data check(s) need confirmation before registration.`, 'error');
      this.focusField(issues[0].key);
      return;
    }

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
    this.toast(this.lang === 'ja' ? '症例を登録しました。' : 'Case registered successfully! ✓', 'success');
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
      const qualityIssues = this.getCaseCompletenessIssues(c).length;
      const quality = qualityIssues === 0
        ? '<span class="quality-status quality-status--ok">OK</span>'
        : `<span class="quality-status quality-status--warn">${this.t('checks')(qualityIssues)}</span>`;
      return `<tr onclick="App.editCase('${c._id}')">
        <td>${i + 1}</td>
        <td>${c.patient_id || c.study_id || '–'}</td>
        <td><span class="badge badge--${adtClass}">${c.ADT_name || '–'}</span></td>
        <td>${c.ADT_start || '–'}</td>
        <td>${c.age || '–'}</td>
        <td>${sexLabel}</td>
        <td>${extentMap[c.disease_extent] || '–'}</td>
        <td>${c.bl_PRO2 || '–'}</td>
        <td>${quality}</td>
        <td class="action-btns" onclick="event.stopPropagation()">
          <button class="btn btn--secondary btn--sm" onclick="App.editCase('${c._id}')">${this.t('edit')}</button>
          <button class="btn btn--danger btn--sm" onclick="App.confirmDelete('${c._id}')">${this.t('delete')}</button>
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
    this.toast((this.lang === 'ja' ? '編集中: ' : 'Editing case: ') + (c.patient_id || c.study_id || id), 'info');
  },

  // ---- Delete Case ----
  confirmDelete(id) {
    this.pendingAction = () => {
      let cases = this.getCases();
      cases = cases.filter(c => c._id !== id);
      this.saveCases(cases);
      this.renderCaseList();
      this.toast(this.lang === 'ja' ? '症例を削除しました。' : 'Case deleted.', 'info');
    };
    document.getElementById('modalTitle').textContent = this.t('deleteCase');
    document.getElementById('modalDesc').textContent = this.t('deleteConfirm');
    document.getElementById('modalConfirmBtn').textContent = this.t('delete');
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
