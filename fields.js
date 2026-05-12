// ============================================
// fields.js - Field Configuration for all 127 columns
// ============================================

const STEPS = [
  { id: 1, label: 'Demographics', title: 'Patient Demographics', desc: 'Basic patient information and treatment history' },
  { id: 2, label: 'Baseline', title: 'Baseline Assessment (Week 0)', desc: 'Concomitant medications, PRO-2, labs, endoscopy, and IUS at enrollment' },
  { id: 3, label: 'Week 12', title: 'Week 12 Assessment', desc: 'Allowance: Week 10–14' },
  { id: 4, label: 'Week 26', title: 'Week 26 Assessment', desc: 'Allowance: Week 24–28' },
  { id: 5, label: 'Week 52', title: 'Week 52 Assessment', desc: 'Assess at the closest Week 52' },
  { id: 6, label: 'TEAEs', title: 'Treatment-Emergent Adverse Events', desc: 'Safety follow-up within 30 days' }
];

const SEX_OPTS = [['0','Male'],['1','Female']];
const YN_OPTS = [['0','No'],['1','Yes']];
const EXTENT_OPTS = [['1','E1 (Proctitis)'],['2','E2 (Left-sided)'],['3','E3 (Pancolitis)']];
const SMOKE_OPTS = [['0','Never'],['1','Former'],['2','Current'],['3','Unknown']];
const ADT_OPTS = [['FIL','FIL'],['OZA','OZA'],['ETR','ETR']];
const ADT_HIST = [['N/A','N/A (Bio-naïve)'],['TNF','TNF'],['VDZ','VDZ'],['UST','UST'],['JAK(TOF or UPA)','JAK (TOF/UPA)'],['IL23p19','IL23p19']];
const SFS_OPTS = [['0','0 – Normal'],['1','1 – 1-2 more'],['2','2 – 3-4 more'],['3','3 – 5+ more']];
const RBS_OPTS = [['0','0 – No blood'],['1','1 – Streaks <50%'],['2','2 – Obvious >50%'],['3','3 – Blood alone']];
const MES_OPTS = [['0','0'],['1','1'],['2','2'],['3','3']];
const UCEIS_OPTS = Array.from({length:9},(_,i)=>[String(i),String(i)]);
const DOPPLER_OPTS = [['0','Signal (-)'],['1','Signal (+)']];
const DISC_OPTS = [['Lack of effectiveness','Lack of effectiveness'],['Disease relapse/worsening','Disease relapse/worsening'],['Adverse event','Adverse event'],['Patient preference','Patient preference'],['Loss to follow-up','Loss to follow-up'],['Transfer to another hospital','Transfer'],['Pregnancy/desire for pregnancy','Pregnancy'],['surgery for UC-related colorectal cancer','Surgery (CRC)'],['death','Death'],['Other','Other']];
const BIOPSY_OPTS = [['0','Rectum'],['1','S/C'],['2','D/C'],['3','T/C'],['4','A/C'],['5','Cecum']];
const HEMA_AE_OPTS = [['0','None'],['ANC<500','ANC < 500'],['ALC<500','ALC < 500'],['ALC<200','ALC < 200'],['Hb<8','Anemia: Hb<8.0'],['PLT<25000','Platelet < 25,000']];
const LIVER_AE_OPTS = [['0','None'],['ALT3x','ALT ≧3×ULN'],['ALT5x','ALT ≧5×ULN'],['gGTP','γGTP increased']];
const CARDIAC_AE_OPTS = [['0','None'],['brady','Bradycardia'],['ECG','ECG abnormality'],['AVB1','AV block 1st'],['AVB2','AV block 2nd (Mobitz I)']];
const INFECT_AE_OPTS = [['0','None'],['Nasopharyngitis','Nasopharyngitis'],['Sinusitis','Sinusitis'],['COVID-19','COVID-19'],['Serious infection','Serious infection'],['URTI','Upper respiratory tract infection'],['Viral RTI','Respiratory tract infection viral'],['Bronchitis','Bronchitis'],['Influenza','Influenza'],['Gastroenteritis','Gastroenteritis'],['Opportunistic','Opportunistic infection'],['Pneumonia','Pneumonia'],['HSV','Herpes simplex'],['CMV','CMV infection/colitis'],['CDI','C. difficile infection'],['TB','Tuberculosis'],['HBV','HBV reactivation'],['Fungal','Fungal infection'],['PJP','Pneumocystis jirovecii']];

// f(key, label, type, opts, extra)
function f(key,label,type,opts,extra){return{key,label,type,opts:opts||null,...(extra||{})};}

const FIELD_CONFIG = {
  1: [ // Demographics
    { group: 'Identification', fields: [
      f('study_id','Study ID','number'),
      f('patient_id','Patient ID','text'),
      f('ADT_name','ADT Name','select',ADT_OPTS,{required:true}),
      f('ADT_start','ADT Start (Week 0)','date',null,{required:true}),
      f('observation_end','Observation End (Week 52)','date',null,{auto:true,calcFn:'calcObsEnd'}),
      f('discontinuation_date','Date of Discontinuation','date',null,{hint:'If drop out before W52'}),
      f('discontinuation_reason','Discontinuation Reason','select',DISC_OPTS),
    ]},
    { group: 'Patient Characteristics', fields: [
      f('date_of_birth','Date of Birth','date'),
      f('age','Age at Enrollment','number',null,{auto:true,calcFn:'calcAge',unit:'years'}),
      f('sex','Sex','radio',SEX_OPTS),
      f('height','Height','number',null,{unit:'cm'}),
      f('weight','Weight','number',null,{unit:'kg'}),
      f('bmi','BMI','number',null,{auto:true,calcFn:'calcBMI',unit:'kg/m²'}),
    ]},
    { group: 'Disease Information', fields: [
      f('disease_extent','Disease Extent (Montreal)','radio',EXTENT_OPTS),
      f('disease_diagnosed','Disease Diagnosed','date',null,{hint:'Year only: yyyy/01/01'}),
      f('disease_duration','Disease Duration','number',null,{auto:true,calcFn:'calcDiseaseDuration',unit:'years'}),
      f('smoking_history','Smoking History','radio',SMOKE_OPTS),
    ]},
    { group: 'Treatment History', fields: [
      f('history_UC_hospitalization','UC-related Hospitalization','radio',YN_OPTS),
      f('history_immunomodulators','Immunomodulators (AZA, 6-MP)','radio',YN_OPTS),
      f('history_corticosteroid','Corticosteroid Use','radio',YN_OPTS),
      f('history_CNIs','CNIs (Tacrolimus, Cyclosporine)','radio',YN_OPTS),
      f('num_biologic_history','Number of Prior Biologic Therapy','number',null,{auto:true,calcFn:'calcBioCount'}),
      f('history_ADT_1st','Prior ADT 1st','select',ADT_HIST),
      f('history_ADT_2nd','Prior ADT 2nd','select',ADT_HIST),
      f('history_ADT_3rd','Prior ADT 3rd','select',ADT_HIST),
      f('history_ADT_4th','Prior ADT 4th','select',ADT_HIST),
      f('history_ADT_5th','Prior ADT 5th','select',ADT_HIST),
    ]},
    { group: 'Other', fields: [
      f('comorbidities','Comorbidities','textarea'),
      f('comment','Comment','textarea'),
    ]},
  ],
  2: [ // Baseline
    { group: 'Concomitant Medications at Enrollment', fields: [
      f('bl_date','Baseline Date','date',null,{auto:true,calcFn:'calcBlDate'}),
      f('bl_5ASA','5-ASA Use (incl. SASP)','radio',YN_OPTS),
      f('bl_corticosteroids','Corticosteroid Use','radio',YN_OPTS),
      f('bl_immunomodulators','Immunomodulators (AZA, 6-MP)','radio',YN_OPTS,{hint:'FIL only'}),
    ]},
    { group: 'PRO-2 (Patient Reported Outcomes)', fields: [
      f('bl_SFS','Stool Frequency Sub-score (SFS)','select',SFS_OPTS),
      f('bl_RBS','Rectal Bleeding Sub-score (RBS)','select',RBS_OPTS),
      f('bl_PRO2','PRO-2 Score','number',null,{auto:true,calcFn:'calcPRO2',prefix:'bl'}),
    ]},
    { group: 'Laboratory (within 3 months before index date)', voicePrefix: 'bl', fields: [
      f('bl_lab_date','Date of Assessment','date'),
      f('bl_CRP','CRP','number',null,{unit:'mg/L'}),
      f('bl_fCAL','fCAL','number',null,{unit:'μg/g'}),
      f('bl_Alb','Albumin','number',null,{unit:'g/dL'}),
      f('bl_WBC','WBC','number',null,{unit:'/mm³'}),
      f('bl_neutro_pct','Neutrophil','number',null,{unit:'%'}),
      f('bl_ANC','Absolute Neutrophil Count','number',null,{auto:true,calcFn:'calcANC',prefix:'bl'}),
      f('bl_lymph_pct','Lymphocyte','number',null,{unit:'%'}),
      f('bl_ALC','Absolute Lymphocyte Count','number',null,{auto:true,calcFn:'calcALC',prefix:'bl'}),
    ]},
    { group: 'Endoscopy (closest to index date)', fields: [
      f('bl_endo_date','Date of Assessment','date'),
      f('bl_MES','MES','select',MES_OPTS),
      f('bl_UCEIS','UCEIS (Exploratory)','select',UCEIS_OPTS),
    ]},
    { group: 'IUS (Intestinal Ultrasound)', fields: [
      f('bl_ius_date','Date of Assessment','date'),
      f('bl_IUS_BWT','IUS BWT (mm)','number',null,{hint:'Greatest bowel wall thickening, excl. rectum'}),
      f('bl_IUS_doppler','IUS Color Doppler','radio',DOPPLER_OPTS),
    ]},
  ],
  3: [ // Week 12
    { group: 'Timing', fields: [
      f('w12_date','Week 12 Date','date',null,{auto:true,calcFn:'calcW12Date'}),
    ]},
    { group: 'PRO-2', fields: [
      f('w12_pro2_date','Date of Assessment (PRO-2)','date'),
      f('w12_SFS','SFS','select',SFS_OPTS),
      f('w12_RBS','RBS','select',RBS_OPTS),
      f('w12_PRO2','PRO-2 Score','number',null,{auto:true,calcFn:'calcPRO2',prefix:'w12'}),
    ]},
    { group: 'Laboratory', voicePrefix: 'w12', fields: [
      f('w12_lab_date','Date of Assessment','date'),
      f('w12_CRP','CRP','number',null,{unit:'mg/L'}),
      f('w12_fCAL','fCAL','number',null,{unit:'μg/g'}),
      f('w12_Alb','Albumin','number',null,{unit:'g/dL'}),
      f('w12_WBC','WBC','number',null,{unit:'/mm³'}),
      f('w12_neutro_pct','Neutrophil','number',null,{unit:'%'}),
      f('w12_ANC','Absolute Neutrophil Count','number',null,{auto:true,calcFn:'calcANC',prefix:'w12'}),
      f('w12_lymph_pct','Lymphocyte','number',null,{unit:'%'}),
      f('w12_ALC','Absolute Lymphocyte Count','number',null,{auto:true,calcFn:'calcALC',prefix:'w12'}),
    ]},
    { group: 'Endoscopy (Exploratory)', fields: [
      f('w12_endo_date','Date of Assessment','date'),
      f('w12_MES','MES','select',MES_OPTS),
      f('w12_UCEIS','UCEIS','select',UCEIS_OPTS),
    ]},
    { group: 'IUS (Exploratory)', fields: [
      f('w12_ius_date','Date of Assessment','date'),
      f('w12_IUS_BWT','IUS BWT (mm)','number'),
      f('w12_IUS_doppler','IUS Color Doppler','radio',DOPPLER_OPTS),
    ]},
  ],
  4: [ // Week 26
    { group: 'Timing', fields: [
      f('w26_date','Week 26 Date','date',null,{auto:true,calcFn:'calcW26Date'}),
    ]},
    { group: 'PRO-2', fields: [
      f('w26_pro2_date','Date of Assessment (PRO-2)','date'),
      f('w26_SFS','SFS','select',SFS_OPTS),
      f('w26_RBS','RBS','select',RBS_OPTS),
      f('w26_PRO2','PRO-2 Score','number',null,{auto:true,calcFn:'calcPRO2',prefix:'w26'}),
    ]},
    { group: 'Laboratory', voicePrefix: 'w26', fields: [
      f('w26_lab_date','Date of Assessment','date'),
      f('w26_CRP','CRP','number',null,{unit:'mg/L'}),
      f('w26_fCAL','fCAL','number',null,{unit:'μg/g'}),
      f('w26_Alb','Albumin','number',null,{unit:'g/dL'}),
      f('w26_WBC','WBC','number',null,{unit:'/mm³'}),
      f('w26_neutro_pct','Neutrophil','number',null,{unit:'%'}),
      f('w26_ANC','Absolute Neutrophil Count','number',null,{auto:true,calcFn:'calcANC',prefix:'w26'}),
      f('w26_lymph_pct','Lymphocyte','number',null,{unit:'%'}),
      f('w26_ALC','Absolute Lymphocyte Count','number',null,{auto:true,calcFn:'calcALC',prefix:'w26'}),
    ]},
    { group: 'Corticosteroid-free Status', fields: [
      f('w26_CS_free','CS-free','radio',YN_OPTS,{hint:'Without systemic CS use for ≥12 weeks before Week 26'}),
    ]},
    { group: 'Endoscopy (Exploratory)', fields: [
      f('w26_endo_date','Date of Assessment','date'),
      f('w26_MES','MES','select',MES_OPTS),
      f('w26_UCEIS','UCEIS','select',UCEIS_OPTS),
    ]},
  ],
  5: [ // Week 52
    { group: 'Timing', fields: [
      f('w52_date','Week 52 Date','date',null,{auto:true,calcFn:'calcW52Date'}),
    ]},
    { group: 'PRO-2', fields: [
      f('w52_pro2_date','Date of Assessment (PRO-2)','date'),
      f('w52_SFS','SFS','select',SFS_OPTS),
      f('w52_RBS','RBS','select',RBS_OPTS),
      f('w52_PRO2','PRO-2 Score','number',null,{auto:true,calcFn:'calcPRO2',prefix:'w52'}),
    ]},
    { group: 'Laboratory', voicePrefix: 'w52', fields: [
      f('w52_lab_date','Date of Assessment','date'),
      f('w52_CRP','CRP','number',null,{unit:'mg/L'}),
      f('w52_fCAL','fCAL','number',null,{unit:'μg/g'}),
      f('w52_Alb','Albumin','number',null,{unit:'g/dL'}),
      f('w52_WBC','WBC','number',null,{unit:'/mm³'}),
      f('w52_neutro_pct','Neutrophil','number',null,{unit:'%'}),
      f('w52_ANC','Absolute Neutrophil Count','number',null,{auto:true,calcFn:'calcANC',prefix:'w52'}),
      f('w52_lymph_pct','Lymphocyte','number',null,{unit:'%'}),
      f('w52_ALC','Absolute Lymphocyte Count','number',null,{auto:true,calcFn:'calcALC',prefix:'w52'}),
    ]},
    { group: 'Corticosteroid-free Status', fields: [
      f('w52_CS_free','CS-free','radio',YN_OPTS,{hint:'Without systemic CS use for ≥12 weeks before Week 52'}),
    ]},
    { group: 'Endoscopy', fields: [
      f('w52_endo_date','Date of Assessment','date'),
      f('w52_MES','MES','select',MES_OPTS),
      f('w52_UCEIS','UCEIS (Exploratory)','select',UCEIS_OPTS),
      f('w52_histologic_remission','Histologic Remission (Exploratory)','radio',YN_OPTS),
      f('w52_biopsy_site','Biopsy Site (Exploratory)','select',BIOPSY_OPTS),
    ]},
    { group: 'IUS (Exploratory)', fields: [
      f('w52_ius_date','Date of Assessment','date'),
      f('w52_IUS_BWT','IUS BWT (mm)','number'),
      f('w52_IUS_doppler','IUS Color Doppler','radio',DOPPLER_OPTS),
    ]},
  ],
  6: [ // TEAEs
    { group: 'Major Adverse Events', fields: [
      f('ae_worsening_UC','Worsening UC','radio',YN_OPTS),
      f('ae_herpes_zoster','Herpes Zoster','radio',YN_OPTS),
      f('ae_VTE','VTE','radio',YN_OPTS),
      f('ae_MACE','MACE','radio',YN_OPTS),
      f('ae_macular_edema','Macular Edema','radio',YN_OPTS),
      f('ae_malignancies','Malignancies (type)','textarea',null,{hint:'Free write: type of cancer'}),
    ]},
    { group: 'Hematologic AEs', fields: [
      f('ae_hematologic','Hematologic AEs','select',HEMA_AE_OPTS),
    ]},
    { group: 'Drug-related Liver Dysfunction', fields: [
      f('ae_liver','Liver Dysfunction','select',LIVER_AE_OPTS),
    ]},
    { group: 'Cardiac AEs', fields: [
      f('ae_cardiac','Bradycardia / Heart Conduction','select',CARDIAC_AE_OPTS),
    ]},
    { group: 'Infectious AEs', fields: [
      f('ae_infectious','Infectious AEs','select',INFECT_AE_OPTS),
    ]},
    { group: 'Other', fields: [
      f('ae_other_details','Other AEs / Details','textarea'),
      f('other_info','Other Important Information','textarea',null,{hint:'e.g. For OZA: dates of interruption, re-initiation, re-titration'}),
    ]},
  ],
};

// Exact Excel column mapping: Col A (1) through Col DW (127)
// null = empty column (section header or spacer in Excel)
// string = field key from our form data
const EXCEL_COLUMN_MAP = [
  // A-AC: Demographics (Cols 1-29)
  'study_id',               // A  (1)
  'patient_id',             // B  (2)
  'ADT_name',               // C  (3)
  'ADT_start',              // D  (4)
  'observation_end',        // E  (5)
  'discontinuation_date',   // F  (6)
  'discontinuation_reason', // G  (7)
  'date_of_birth',          // H  (8)
  'age',                    // I  (9)
  'sex',                    // J  (10)
  'height',                 // K  (11)
  'weight',                 // L  (12)
  'bmi',                    // M  (13)
  'disease_extent',         // N  (14)
  'disease_diagnosed',      // O  (15)
  'disease_duration',       // P  (16)
  'smoking_history',        // Q  (17)
  'history_UC_hospitalization', // R  (18)
  'history_immunomodulators',   // S  (19)
  'history_corticosteroid',     // T  (20)
  'history_CNIs',               // U  (21)
  'num_biologic_history',       // V  (22)
  'history_ADT_1st',        // W  (23)
  'history_ADT_2nd',        // X  (24)
  'history_ADT_3rd',        // Y  (25)
  'history_ADT_4th',        // Z  (26)
  'history_ADT_5th',        // AA (27)
  'comorbidities',          // AB (28)
  'comment',                // AC (29)
  // AD-AY: Baseline (Cols 30-51)
  'bl_date',                // AD (30)
  'bl_5ASA',                // AE (31)
  'bl_corticosteroids',     // AF (32)
  'bl_immunomodulators',    // AG (33)
  'bl_SFS',                 // AH (34)
  'bl_RBS',                 // AI (35)
  'bl_PRO2',                // AJ (36)
  'bl_lab_date',            // AK (37)
  'bl_CRP',                 // AL (38)
  'bl_fCAL',                // AM (39)
  'bl_Alb',                 // AN (40)
  'bl_WBC',                 // AO (41)
  'bl_neutro_pct',          // AP (42)
  'bl_ANC',                 // AQ (43)
  'bl_lymph_pct',           // AR (44)
  'bl_ALC',                 // AS (45)
  'bl_endo_date',           // AT (46)
  'bl_MES',                 // AU (47)
  'bl_UCEIS',               // AV (48)
  'bl_ius_date',            // AW (49)
  'bl_IUS_BWT',             // AX (50)
  'bl_IUS_doppler',         // AY (51)
  // AZ-BS: Week 12 (Cols 52-71)
  'w12_date',               // AZ (52)
  'w12_pro2_date',          // BA (53)
  'w12_SFS',                // BB (54)
  'w12_RBS',                // BC (55)
  'w12_PRO2',               // BD (56)
  'w12_lab_date',           // BE (57)
  'w12_CRP',                // BF (58)
  'w12_fCAL',               // BG (59)
  'w12_Alb',                // BH (60)
  'w12_WBC',                // BI (61)
  'w12_neutro_pct',         // BJ (62)
  'w12_ANC',                // BK (63)
  'w12_lymph_pct',          // BL (64)
  'w12_ALC',                // BM (65)
  'w12_endo_date',          // BN (66)
  'w12_MES',                // BO (67)
  'w12_UCEIS',              // BP (68)
  'w12_ius_date',           // BQ (69)
  'w12_IUS_BWT',            // BR (70)
  'w12_IUS_doppler',        // BS (71)
  // BT-CK: Week 26 (Cols 72-89)
  'w26_date',               // BT (72)
  'w26_pro2_date',          // BU (73)
  'w26_SFS',                // BV (74)
  'w26_RBS',                // BW (75)
  'w26_PRO2',               // BX (76)
  'w26_lab_date',           // BY (77)
  'w26_CRP',                // BZ (78)
  'w26_fCAL',               // CA (79)
  'w26_Alb',                // CB (80)
  'w26_WBC',                // CC (81)
  'w26_neutro_pct',         // CD (82)
  'w26_ANC',                // CE (83)
  'w26_lymph_pct',          // CF (84)
  'w26_ALC',                // CG (85)
  'w26_CS_free',            // CH (86)
  'w26_endo_date',          // CI (87)
  'w26_MES',                // CJ (88)
  'w26_UCEIS',              // CK (89)
  // CL-DH: Week 52 (Cols 90-112)
  'w52_date',               // CL (90)
  'w52_pro2_date',          // CM (91)
  'w52_SFS',                // CN (92)
  'w52_RBS',                // CO (93)
  'w52_PRO2',               // CP (94)
  'w52_lab_date',           // CQ (95)
  'w52_CRP',                // CR (96)
  'w52_fCAL',               // CS (97)
  'w52_Alb',                // CT (98)
  'w52_WBC',                // CU (99)
  'w52_neutro_pct',         // CV (100)
  'w52_ANC',                // CW (101)
  'w52_lymph_pct',          // CX (102)
  'w52_ALC',                // CY (103)
  'w52_CS_free',            // CZ (104)
  'w52_endo_date',          // DA (105)
  'w52_MES',                // DB (106)
  'w52_UCEIS',              // DC (107)
  'w52_histologic_remission', // DD (108)
  'w52_biopsy_site',        // DE (109)
  'w52_ius_date',           // DF (110)
  'w52_IUS_BWT',            // DG (111)
  'w52_IUS_doppler',        // DH (112)
  // DI-DW: TEAEs (Cols 113-127)
  null,                     // DI (113) – section header
  'ae_worsening_UC',        // DJ (114)
  'ae_herpes_zoster',       // DK (115)
  'ae_VTE',                 // DL (116)
  'ae_MACE',                // DM (117)
  'ae_macular_edema',       // DN (118)
  'ae_malignancies',        // DO (119)
  'ae_hematologic',         // DP (120)
  'ae_liver',               // DQ (121)
  'ae_cardiac',             // DR (122)
  'ae_infectious',          // DS (123)
  'ae_other_details',       // DT (124)
  null,                     // DU (125) – empty
  null,                     // DV (126) – empty
  'other_info',             // DW (127)
];
