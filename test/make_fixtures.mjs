// يولّد ملفات عملاء "مخربطة" بأشكال مختلفة لنفس 40 موظفًا وهميًا + الحقيقة المتوقعة (truth.json)
// كل ملف يختبر نوعًا من الفوضى: ترويسة بصفين، بدون ترويسة، شيتات متعددة، CSV بترميز ويندوز، أسماء مقسّمة...
// البيانات وهمية بالكامل (هويات وآيبانات بخانات تحقق صحيحة لكنها غير حقيقية).
import ExcelJS from 'exceljs';
import fs from 'node:fs';

const OUT = new URL('./fixtures/', import.meta.url);
fs.mkdirSync(OUT, { recursive: true });

let seed = 20261022;
const rnd = () => (seed = (seed * 1103515245 + 12345) % 2147483648) / 2147483648;
const pick = (a) => a[Math.floor(rnd() * a.length)];
const ri = (a, b) => a + Math.floor(rnd() * (b - a + 1));
const pad = (n, w = 2) => String(n).padStart(w, '0');
const AR = '٠١٢٣٤٥٦٧٨٩';
const arDigits = (s) => String(s).replace(/\d/g, (d) => AR[d]);

/* ---------------- الجنسيات ---------------- */
const NAT = {
  SA: { canon: 'السعودية', m: 'سعودي', f: 'سعودية', en: 'Saudi Arabia', dem: 'Saudi', code: 'KSA' },
  EG: { canon: 'مصر', m: 'مصري', f: 'مصرية', en: 'Egypt', dem: 'Egyptian', code: 'EG' },
  IN: { canon: 'الهند', m: 'هندي', f: 'هندية', en: 'India', dem: 'Indian', code: 'IN' },
  PK: { canon: 'باكستان', m: 'باكستاني', f: 'باكستانية', en: 'Pakistan', dem: 'Pakistani', code: 'PK' },
  PH: { canon: 'الفلبّين', m: 'فلبيني', f: 'فلبينية', en: 'Philippines', dem: 'Filipino', code: 'PH' },
  JO: { canon: 'الأردن', m: 'أردني', f: 'أردنية', en: 'Jordan', dem: 'Jordanian', code: 'JO' },
  SY: { canon: 'سوريا', m: 'سوري', f: 'سورية', en: 'Syria', dem: 'Syrian', code: 'SY' },
  SD: { canon: 'السودان', m: 'سوداني', f: 'سودانية', en: 'Sudan', dem: 'Sudanese', code: 'SD' },
  YE: { canon: 'اليمن', m: 'يمني', f: 'يمنية', en: 'Yemen', dem: 'Yemeni', code: 'YE' },
  LB: { canon: 'لبنان', m: 'لبناني', f: 'لبنانية', en: 'Lebanon', dem: 'Lebanese', code: 'LB' },
  MA: { canon: 'المغرب', m: 'مغربي', f: 'مغربية', en: 'Morocco', dem: 'Moroccan', code: 'MA' },
  BD: { canon: 'بنغلاديش', m: 'بنغلاديشي', f: 'بنغلاديشية', en: 'Bangladesh', dem: 'Bangladeshi', code: 'BD' },
  NP: { canon: 'نيبال', m: 'نيبالي', f: 'نيبالية', en: 'Nepal', dem: 'Nepali', code: 'NP' },
  LK: { canon: 'سريلانكا', m: 'سريلانكي', f: 'سريلانكية', en: 'Sri Lanka', dem: 'Sri Lankan', code: 'LK' },
  GB: { canon: 'المملكة المتحدة', m: 'بريطاني', f: 'بريطانية', en: 'United Kingdom', dem: 'British', code: 'UK' },
  US: { canon: 'الولايات المتّحدة', m: 'أمريكي', f: 'أمريكية', en: 'United States', dem: 'American', code: 'USA' },
  KE: { canon: 'كينيا', m: 'كيني', f: 'كينية', en: 'Kenya', dem: 'Kenyan', code: 'KE' },
  TR: { canon: 'تركيّا', m: 'تركي', f: 'تركية', en: 'Turkey', dem: 'Turkish', code: 'TR' },
};

// [جنسية, جنس, [أجزاء الاسم العربي], [أجزاء الاسم الإنجليزي]]
const PEOPLE = [
  ['SA', 'M', ['محمد', 'عبدالله', 'العتيبي'], ['Mohammed', 'Abdullah', 'Alotaibi']],
  ['SA', 'F', ['سارة', 'خالد', 'القحطاني'], ['Sara', 'Khalid', 'Alqahtani']],
  ['SA', 'M', ['عبد الرحمن', 'سعد', 'الدوسري'], ['Abdul Rahman', 'Saad', 'Aldosari']],
  ['SA', 'M', ['فهد', 'ناصر', 'الشمري'], ['Fahad', 'Nasser', 'Alshammari']],
  ['SA', 'F', ['نورة', 'صالح', 'الغامدي'], ['Noura', 'Saleh', 'Alghamdi']],
  ['SA', 'F', ['ليلى', 'إبراهيم', 'الزهراني'], ['Layla', 'Ibrahim', 'Alzahrani']],
  ['SA', 'M', ['بدر', 'عبدالعزيز', 'المطيري'], ['Badr', 'Abdulaziz', 'Almutairi']],
  ['SA', 'F', ['هند', 'محمد', 'الحربي'], ['Hind', 'Mohammed', 'Alharbi']],
  ['SA', 'M', ['تركي', 'سلطان', 'السبيعي'], ['Turki', 'Sultan', 'Alsubaie']],
  ['SA', 'M', ['نايف', 'ماجد', 'العنزي'], ['Naif', 'Majed', 'Alanazi']],
  ['SA', 'F', ['ريم', 'فيصل', 'الشهري'], ['Reem', 'Faisal', 'Alshehri']],
  ['SA', 'M', ['ياسر', 'بندر', 'المالكي'], ['Yasser', 'Bandar', 'Almalki']],
  ['EG', 'M', ['أحمد', 'محمود', 'السيد'], ['Ahmed', 'Mahmoud', 'Elsayed']],
  ['EG', 'F', ['ياسمين', 'مصطفى', 'عبدالفتاح'], ['Yasmin', 'Mostafa', 'Abdelfattah']],
  ['EG', 'M', ['مصطفى', 'حسن', 'الشريف'], ['Mostafa', 'Hassan', 'Elsherif']],
  ['IN', 'M', ['راجيش', 'كومار'], ['Rajesh', 'Kumar']],
  ['IN', 'M', ['سوريش', 'ناير'], ['Suresh', 'Nair']],
  ['IN', 'F', ['بريا', 'شارما'], ['Priya', 'Sharma']],
  ['IN', 'M', ['محمد', 'عرفان'], ['Mohammed', 'Irfan']],
  ['PK', 'M', ['عمران', 'خان'], ['Imran', 'Khan']],
  ['PK', 'M', ['بلال', 'أحمد'], ['Bilal', 'Ahmed']],
  ['PK', 'F', ['عائشة', 'صديقي'], ['Ayesha', 'Siddiqui']],
  ['PH', 'F', ['ماريا', 'سانتوس'], ['Maria', 'Santos']],
  ['PH', 'M', ['خوسيه', 'رييس'], ['Jose', 'Reyes']],
  ['PH', 'F', ['أنجليكا', 'كروز'], ['Angelica', 'Cruz']],
  ['JO', 'M', ['يزن', 'رامي', 'الخطيب'], ['Yazan', 'Rami', 'Alkhatib']],
  ['JO', 'F', ['دانة', 'عمر', 'العبادي'], ['Dana', 'Omar', 'Alabbadi']],
  ['SY', 'M', ['عمر', 'وليد', 'الحلبي'], ['Omar', 'Waleed', 'Alhalabi']],
  ['SD', 'M', ['عثمان', 'أحمد', 'عبدالرحيم'], ['Osman', 'Ahmed', 'Abdelrahim']],
  ['YE', 'M', ['صالح', 'علي', 'اليافعي'], ['Saleh', 'Ali', 'Alyafei']],
  ['LB', 'F', ['لينا', 'جورج', 'حداد'], ['Lina', 'George', 'Haddad']],
  ['MA', 'M', ['يوسف', 'العلوي'], ['Youssef', 'Alaoui']],
  ['BD', 'M', ['رحيم', 'الدين'], ['Rahim', 'Uddin']],
  ['BD', 'M', ['كريم', 'حسين'], ['Karim', 'Hossain']],
  ['NP', 'M', ['رام', 'ثابا'], ['Ram', 'Thapa']],
  ['LK', 'M', ['نوان', 'بيريرا'], ['Nuwan', 'Perera']],
  ['GB', 'M', ['جيمس', 'ويلسون'], ['James', 'Wilson']],
  ['US', 'F', ['إيميلي', 'جونسون'], ['Emily', 'Johnson']],
  ['KE', 'M', ['جون', 'موانجي'], ['John', 'Mwangi']],
  ['TR', 'M', ['محمد', 'يلماز'], ['Mehmet', 'Yilmaz']],
];

const DEPTS = [
  { canon: 'المالية', ar: ['المالية', 'الحسابات', 'المالية والمحاسبة'], en: ['Finance', 'Accounting'] },
  { canon: 'الموارد البشرية', ar: ['الموارد البشرية', 'شؤون الموظفين', 'الموارد البشريه'], en: ['HR', 'Human Resources'] },
  { canon: 'المبيعات', ar: ['المبيعات', 'مبيعات'], en: ['Sales'] },
  { canon: 'الهندسة', ar: ['الهندسة', 'القسم الهندسي'], en: ['Engineering'] },
  { canon: 'العمليات', ar: ['العمليات', 'التشغيل'], en: ['Operations', 'Ops'] },
  { canon: 'التسويق', ar: ['التسويق'], en: ['Marketing'] },
  { canon: 'تقنية المعلومات', ar: ['تقنية المعلومات', 'تقنيه المعلومات'], en: ['IT', 'Information Technology'] },
  { canon: 'خدمة العملاء', ar: ['خدمة العملاء'], en: ['Customer Service'] },
];
const JOBS = {
  'المالية': [['محاسب', 'Accountant'], ['محلل مالي', 'Financial Analyst']],
  'الموارد البشرية': [['أخصائي موارد بشرية', 'HR Specialist'], ['مسؤول رواتب', 'Payroll Officer']],
  'المبيعات': [['مندوب مبيعات', 'Sales Representative'], ['مدير مبيعات', 'Sales Manager']],
  'الهندسة': [['مهندس مدني', 'Civil Engineer'], ['مهندس كهرباء', 'Electrical Engineer']],
  'العمليات': [['فني صيانة', 'Maintenance Technician'], ['سائق', 'Driver'], ['مشرف عمليات', 'Operations Supervisor']],
  'التسويق': [['مصمم جرافيك', 'Graphic Designer'], ['أخصائي تسويق', 'Marketing Specialist']],
  'تقنية المعلومات': [['مطور برمجيات', 'Software Developer'], ['مهندس شبكات', 'Network Engineer']],
  'خدمة العملاء': [['موظف خدمة عملاء', 'Customer Service Agent']],
};
const LOCS = [
  { canon: 'المقر الرئيسي - الرياض', ar: 'الرياض', en: 'Riyadh', sheet: 'الرياض' },
  { canon: 'فرع جدة', ar: 'جدة', en: 'Jeddah', sheet: 'جدة' },
  { canon: 'فرع الدمام', ar: 'الدمام', en: 'Dammam', sheet: 'الدمام' },
];
const MARITAL = [
  { canon: 'أعزب', m: 'أعزب', f: 'عزباء', en: 'Single' },
  { canon: 'متزوج', m: 'متزوج', f: 'متزوجة', en: 'Married' },
  { canon: 'مطلق', m: 'مطلق', f: 'مطلقة', en: 'Divorced' },
  { canon: 'أرمل', m: 'أرمل', f: 'أرملة', en: 'Widowed' },
];
const EDU = [
  { canon: 'شهادة ثانوية', ar: 'ثانوية عامة', en: 'High School' },
  { canon: 'درجة البكالوريوس', ar: 'بكالوريوس', en: "Bachelor's" },
  { canon: 'درجة ماجستير', ar: 'ماجستير', en: 'Master' },
  { canon: 'درجة البكالوريوس', ar: 'بكالوريوس', en: 'BSc' },
];
const STATUS = [
  { canon: 'مفعّل', ar: ['نشط', 'على رأس العمل', 'فعال'], en: 'Active' },
  { canon: 'فترة الإشعار', ar: ['فترة إشعار'], en: 'Notice Period' },
  { canon: 'غير مفعّل', ar: ['منتهي'], en: 'Inactive' },
];
const BANKS = [['80', 'مصرف الراجحي', 'Al Rajhi Bank'], ['10', 'البنك الأهلي السعودي', 'SNB'], ['45', 'ساب', 'SAB'], ['20', 'بنك الرياض', 'Riyad Bank']];

function luhnId(first) {
  const p = first + Array.from({ length: 8 }, () => ri(0, 9)).join('');
  for (let d = 0; d <= 9; d++) {
    const s = p + d; let t = 0;
    for (let i = 0; i < 10; i++) { let n = +s[i]; if (i % 2 === 0) { n *= 2; if (n > 9) n -= 9; } t += n; }
    if (t % 10 === 0) return s;
  }
}
function iban(bank) {
  const bban = bank + Array.from({ length: 18 }, () => ri(0, 9)).join('');
  const num = (bban + 'SA00').replace(/[A-Z]/g, (c) => c.charCodeAt(0) - 55);
  let rem = 0; for (const ch of num) rem = (rem * 10 + +ch) % 97;
  return `SA${pad(98 - rem)}${bban}`;
}
const iso = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;
const randDate = (y1, y2) => iso(ri(y1, y2), ri(1, 12), ri(1, 28));
const addYears = (s, n) => { const [y, m, d] = s.split('-').map(Number); return iso(y + n, m, d); };
const COMPOUND = new Set(['عبد', 'abdul', 'abd', 'ابو', 'أبو', 'abu', 'بن', 'bin']);
const splitRest = (parts) => {
  const t = parts.join(' ').split(' ');
  const n = COMPOUND.has(t[0].toLowerCase()) && t.length > 2 ? 2 : 1;
  return [t.slice(0, n).join(' '), t.slice(n).join(' ')];
};

/* ---------------- بناء الموظفين ---------------- */
const EMPS = PEOPLE.map(([nat, g, ar, en], i) => {
  const N = NAT[nat];
  const isSA = nat === 'SA';
  const dept = DEPTS[i % DEPTS.length];
  const [jobAr, jobEn] = pick(JOBS[dept.canon]);
  const loc = LOCS[i % 3];
  const dob = randDate(1970, 2001);
  const issue = randDate(2019, 2025);
  const expiry = isSA ? addYears(issue, 10) : addYears(issue, ri(2, 3));
  const join = randDate(2012, 2025);
  const marital = MARITAL[i % 7 === 3 ? 2 : i % 11 === 5 ? 3 : i % 3 === 0 ? 0 : 1];
  const status = STATUS[i === 17 || i === 30 ? 1 : i === 36 ? 2 : 0];
  const bank = BANKS[i % BANKS.length];
  const mob = `05${pick(['0', '3', '4', '5', '6', '9'])}${Array.from({ length: 7 }, () => ri(0, 9)).join('')}`;
  const [fa, la] = splitRest(ar);
  const [fe, le] = splitRest(en);
  return {
    i, nat, N, g, isSA, ar, en,
    first_name_ar: fa, last_name_ar: la, first_name: fe, last_name: le,
    gender: g === 'M' ? 'ذكر' : 'أنثى',
    nationality: N.canon,
    national_id_number: luhnId(isSA ? '1' : '2'),
    date_of_birth: dob, national_id_issued_date: issue, national_id_expiry_date: expiry,
    email: `${en[0]}.${en[en.length - 1]}`.toLowerCase().replace(/\s+/g, '') + '@alufuq-demo.sa',
    work_phone_number: mob,
    job_title_ar: jobAr, job_title_en: jobEn,
    department: dept.canon, deptRep: dept,
    location: loc.canon, locRep: loc,
    employee_no: `E-${1001 + i}`,
    join_date: join,
    employment_type: i % 13 === 4 ? 'دوام جزئي' : 'دوام كامل',
    contract_type: isSA ? 'غير محدد المدة' : 'محدد المدة',
    contract_end_date: isSA ? null : iso(2027 + (i % 2), ri(1, 12), ri(1, 28)),
    marital_status: marital.canon, maritalRep: marital,
    employment_status: status.canon, statusRep: status,
    education: EDU[i % EDU.length],
    salary: ri(60, 600) * 50,
    bank, iban: iban(bank[0]),
  };
});
// المدير المباشر: أول موظف بنفس القسم
for (const e of EMPS) { const m = EMPS.find((x) => x.department === e.department && x !== e); e.managerAr = m && m.i < e.i ? m.ar.join(' ') : ''; }

/* ---------------- صيغ العرض ---------------- */
const dParts = (s) => s.split('-').map(Number);
const fmt = {
  dmy: (s) => { const [y, m, d] = dParts(s); return `${pad(d)}/${pad(m)}/${y}`; },
  mdy: (s) => { const [y, m, d] = dParts(s); return `${m}/${d}/${y}`; },
  ymdSlash: (s) => s.replace(/-/g, '/'),
  dmyDash: (s) => { const [y, m, d] = dParts(s); return `${pad(d)}-${pad(m)}-${y}`; },
  withTime: (s) => `${s} 00:00:00`,
  dMonY: (s) => { const [y, m, d] = dParts(s); return `${d}-${['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'][m - 1]}-${y}`; },
  arabic: (s) => arDigits(fmt.dmy(s)),
  date: (s) => { const [y, m, d] = dParts(s); return new Date(Date.UTC(y, m - 1, d)); },
};
const hf = new Intl.DateTimeFormat('en-u-ca-islamic-umalqura-nu-latn', { timeZone: 'UTC', year: 'numeric', month: 'numeric', day: 'numeric' });
fmt.hijri = (s) => { const p = Object.fromEntries(hf.formatToParts(fmt.date(s)).map((x) => [x.type, x.value])); return `${p.year}/${pad(p.month)}/${pad(p.day)}هـ`; };
const natAr = (e) => (e.g === 'F' ? e.N.f : e.N.m);
const marAr = (e) => (e.g === 'F' ? e.maritalRep.f : e.maritalRep.m);
const spaced = (p) => `${p.slice(0, 3)} ${p.slice(3, 6)} ${p.slice(6)}`;
const ibanSpaced = (s) => s.replace(/(.{4})/g, '$1 ').trim();
const money = (n) => n.toLocaleString('en-US');

/* ---------------- الكاتب العام ---------------- */
// spec: { file, sheets: [{name, title:[...], groups:[{text,span}], cols:[{h, key, v}], emps, sections, repeatHeaderAt, totals, startCol, mergeKey, rtl}] }
const truth = {};
async function writeXlsx(name, spec) {
  const wb = new ExcelJS.Workbook();
  const expectMap = {};
  for (const sh of spec.sheets) {
    const ws = wb.addWorksheet(sh.name, { views: [{ rightToLeft: sh.rtl !== false }] });
    const c0 = sh.startCol || 1;
    let r = 1;
    const ncol = sh.cols.length + (sh.gapAt ? 1 : 0);
    const colOf = (j) => c0 + j + (sh.gapAt && j >= sh.gapAt ? 1 : 0);
    for (const t of sh.title || []) {
      ws.getCell(r, c0).value = t; ws.getCell(r, c0).font = { bold: true, size: 14 };
      ws.mergeCells(r, c0, r, c0 + ncol - 1); r++;
    }
    if (sh.title?.length) r++;
    if (sh.groups) {
      let j = 0;
      for (const g of sh.groups) {
        ws.getCell(r, colOf(j)).value = g.text;
        if (g.span > 1) ws.mergeCells(r, colOf(j), r, colOf(j + g.span - 1));
        j += g.span;
      }
      ws.getRow(r).font = { bold: true }; r++;
    }
    const writeHeader = () => {
      if (sh.noHeader) return;
      sh.cols.forEach((c, j) => { ws.getCell(r, colOf(j)).value = c.h; });
      ws.getRow(r).font = { bold: true }; r++;
    };
    writeHeader();
    let lastSection = null; let k = 0; let mergeStart = null; let mergeVal = null;
    const flushMerge = (endRow) => {
      if (sh.mergeKey === undefined || mergeStart === null) return;
      const j = sh.cols.findIndex((c) => c.key === sh.mergeKey);
      if (endRow > mergeStart) ws.mergeCells(mergeStart, colOf(j), endRow, colOf(j));
    };
    for (const e of sh.emps) {
      if (sh.sections) {
        const s = sh.sections(e);
        if (s !== lastSection) { flushMerge(r - 1); mergeStart = null; ws.getCell(r, colOf(1)).value = s; ws.getCell(r, colOf(1)).font = { bold: true }; r++; lastSection = s; }
      }
      if (sh.repeatHeaderAt && k === sh.repeatHeaderAt) { flushMerge(r - 1); mergeStart = null; writeHeader(); }
      sh.cols.forEach((c, j) => {
        const v = c.v(e, k);
        if (v === null || v === undefined || v === '') return;
        if (sh.mergeKey !== undefined && c.key === sh.mergeKey) {
          if (mergeStart !== null && v === mergeVal) return;
          flushMerge(r - 1); mergeStart = r; mergeVal = v;
        }
        const cell = ws.getCell(r, colOf(j));
        cell.value = v;
        if (v instanceof Date) cell.numFmt = c.numFmt || 'yyyy-mm-dd';
      });
      r++; k++;
    }
    flushMerge(r - 1);
    if (sh.totals) { ws.getCell(r, colOf(1)).value = 'الإجمالي'; const j = sh.cols.findIndex((c) => c.key === 'salary_basic_salary_amount'); if (j >= 0) ws.getCell(r, colOf(j)).value = sh.emps.reduce((a, e) => a + e.salary, 0); r += 2; }
    for (const f of sh.footer || []) { ws.getCell(r, colOf(0)).value = f; r++; }
    sh.cols.forEach((c, j) => { ws.getColumn(colOf(j)).width = 18; if (c.key && !expectMap[c.key] && !sh.skipExpect) expectMap[c.key] = colOf(j) - 1; });
  }
  if (spec.extraSheets) for (const x of spec.extraSheets) { const ws = wb.addWorksheet(x.name); x.rows.forEach((row) => ws.addRow(row)); }
  await wb.xlsx.writeFile(new URL(name, OUT));
  return expectMap;
}

function record(e, over = {}) {
  const base = {
    first_name_ar: e.first_name_ar, last_name_ar: e.last_name_ar, first_name: e.first_name, last_name: e.last_name,
    gender: e.gender, nationality: e.nationality, national_id_number: e.national_id_number,
    date_of_birth: e.date_of_birth, email: e.email, work_phone_number: e.work_phone_number.slice(-9),
    department: e.department, join_date: e.join_date, salary_basic_salary_amount: e.salary, iban: e.iban,
    employment_status: e.employment_status, marital_status: e.marital_status, employee_no: e.employee_no,
    national_id_issued_date: e.national_id_issued_date, national_id_expiry_date: e.national_id_expiry_date,
    location: e.location, contract_end_date: e.contract_end_date,
  };
  return { ...base, ...over };
}
function addTruth(name, map, fields, opts = {}) {
  truth[name] = {
    expectMap: map,
    fields,
    synthetic: opts.synthetic || {},
    records: EMPS.map((e) => {
      const r = record(e, opts.over ? opts.over(e) : {});
      return Object.fromEntries(fields.map((f) => [f, r[f]]));
    }),
  };
}

/* ======================= 1) عربي كلاسيكي بعناوين وإجماليات ======================= */
{
  const cols = [
    { h: 'م', v: (e, k) => k + 1 },
    { h: 'الاسم الكامل', key: 'full_name_ar', v: (e) => e.ar.join(' ') },
    { h: 'Full Name', key: 'full_name_en', v: (e) => e.en.join(' ') },
    { h: 'الجنسية', key: 'nationality', v: natAr },
    { h: 'الجنس', key: 'gender', v: (e) => e.gender },
    { h: 'الحالة الاجتماعية', key: 'marital_status', v: marAr },
    { h: 'تاريخ الميلاد', key: 'date_of_birth', v: (e, k) => [fmt.date, fmt.dmy, fmt.arabic, fmt.ymdSlash][k % 4](e.date_of_birth) },
    { h: 'رقم الهوية / الإقامة', key: 'national_id_number', v: (e, k) => (k % 2 ? +e.national_id_number : e.national_id_number) },
    { h: 'تاريخ إصدار الهوية', key: 'national_id_issued_date', v: (e) => fmt.date(e.national_id_issued_date) },
    { h: 'تاريخ انتهاء الهوية', key: 'national_id_expiry_date', v: (e) => fmt.date(e.national_id_expiry_date) },
    { h: 'البريد الإلكتروني', key: 'email', v: (e, k) => (k % 5 === 1 ? e.email.toUpperCase() : e.email) },
    { h: 'الجوال', key: 'work_phone_number', v: (e, k) => (k % 3 === 0 ? `+966 ${spaced(e.work_phone_number.slice(1))}` : k % 3 === 1 ? arDigits(e.work_phone_number) : e.work_phone_number) },
    { h: 'المسمى الوظيفي', key: 'job_title', v: (e) => e.job_title_ar },
    { h: 'القسم', key: 'department', v: (e, k) => e.deptRep.ar[k % e.deptRep.ar.length] },
    { h: 'الفرع', key: 'location', v: (e) => e.locRep.ar },
    { h: 'تاريخ التعيين', key: 'join_date', v: (e) => fmt.date(e.join_date) },
    { h: 'نوع الدوام', key: 'employment_type', v: (e) => (e.employment_type === 'دوام جزئي' ? 'جزئي' : 'دوام كامل') },
    { h: 'نهاية العقد', key: 'contract_end_date', v: (e) => (e.contract_end_date ? fmt.date(e.contract_end_date) : null) },
    { h: 'الراتب الأساسي', key: 'salary_basic_salary_amount', v: (e, k) => (k % 3 === 0 ? `${money(e.salary)} ريال` : e.salary) },
    { h: 'البنك', key: 'bank_name', v: (e) => e.bank[1] },
    { h: 'الآيبان', key: 'iban', v: (e, k) => (k % 2 ? ibanSpaced(e.iban) : e.iban) },
    { h: 'المؤهل', key: 'education_level', v: (e) => e.education.ar },
    { h: 'الحالة', key: 'employment_status', v: (e) => e.statusRep.ar[0] },
    { h: 'رقم الموظف', key: 'employee_no', v: (e) => e.employee_no },
    { h: 'المدير المباشر', key: 'manager', v: (e) => e.managerAr },
  ];
  const map = await writeXlsx('01_classic_ar.xlsx', { sheets: [{ name: 'الموظفين', title: ['شركة الأفق للتجارة المحدودة', 'كشف بيانات الموظفين — أكتوبر 2026'], cols, emps: EMPS, totals: true, footer: ['المدير العام: ______________', 'التوقيع والختم'] }] });
  addTruth('01_classic_ar.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'national_id_issued_date', 'national_id_expiry_date', 'email', 'work_phone_number', 'department', 'location', 'join_date', 'salary_basic_salary_amount', 'iban', 'employment_status', 'marital_status', 'employee_no', 'contract_end_date']);
}

/* ======================= 2) تصدير إنجليزي من نظام قديم (تواريخ أمريكية، جوال رقمي) ======================= */
{
  const cols = [
    { h: 'Emp #', key: 'employee_no', v: (e) => 1001 + e.i },
    { h: 'First Name', key: 'first_name', v: (e) => e.first_name },
    { h: 'Last Name', key: 'last_name', v: (e) => e.last_name },
    { h: 'Name (Arabic)', key: 'full_name_ar', v: (e) => e.ar.join(' ') },
    { h: 'Sex', key: 'gender', v: (e) => e.g },
    { h: 'DOB', key: 'date_of_birth', v: (e) => fmt.mdy(e.date_of_birth) },
    { h: 'Marital', key: 'marital_status', v: (e) => e.maritalRep.en },
    { h: 'Nationality', key: 'nationality', v: (e) => e.N.dem },
    { h: 'Iqama / ID No.', key: 'national_id_number', v: (e) => +e.national_id_number },
    { h: 'ID Issue', key: 'national_id_issued_date', v: (e) => fmt.mdy(e.national_id_issued_date) },
    { h: 'ID Expiry', key: 'national_id_expiry_date', v: (e) => fmt.mdy(e.national_id_expiry_date) },
    { h: 'Personal Email', key: 'email', v: (e) => e.email },
    { h: 'Mobile No.', key: 'work_phone_number', v: (e) => +e.work_phone_number },
    { h: 'Position', key: 'job_title', v: (e) => e.job_title_en },
    { h: 'Dept.', key: 'department', v: (e, k) => e.deptRep.en[k % e.deptRep.en.length] },
    { h: 'Branch', key: 'location', v: (e) => e.locRep.en },
    { h: 'Hire Date', key: 'join_date', v: (e) => fmt.mdy(e.join_date) },
    { h: 'Employment Type', key: 'employment_type', v: (e) => (e.employment_type === 'دوام جزئي' ? 'Part-time' : 'Full-time') },
    { h: 'Basic Salary', key: 'salary_basic_salary_amount', v: (e) => e.salary },
    { h: 'Housing Allowance', v: (e) => Math.round(e.salary * 0.25) },
    { h: 'Transport Allowance', v: () => 500 },
    { h: 'Gross Salary', v: (e) => e.salary + Math.round(e.salary * 0.25) + 500 },
    { h: 'Bank', key: 'bank_name', v: (e) => e.bank[2] },
    { h: 'IBAN', key: 'iban', v: (e) => e.iban },
    { h: 'Education', key: 'education_level', v: (e) => e.education.en },
    { h: 'Status', key: 'employment_status', v: (e) => e.statusRep.en },
  ];
  const map = await writeXlsx('02_english_export.xlsx', { sheets: [{ name: 'Sheet1', rtl: false, cols, emps: EMPS }] });
  addTruth('02_english_export.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'national_id_issued_date', 'national_id_expiry_date', 'email', 'work_phone_number', 'department', 'location', 'join_date', 'salary_basic_salary_amount', 'iban', 'employment_status', 'marital_status'], { over: (e) => ({ employee_no: String(1001 + e.i) }) });
}

/* ======================= 3) ترويسة بمستويين + عناوين غامضة (الرقم، الحالة مرتين) + هجري ======================= */
{
  const cols = [
    { h: 'الاسم بالعربي', key: 'full_name_ar', v: (e) => e.ar.join(' ') },
    { h: 'الاسم بالإنجليزي', key: 'full_name_en', v: (e) => e.en.join(' ') },
    { h: 'الجنس', key: 'gender', v: (e) => (e.g === 'M' ? 'ذكر' : 'انثى') },
    { h: 'الجنسية', key: 'nationality', v: (e, k) => (k % 2 ? e.N.dem : natAr(e)) },
    { h: 'تاريخ الميلاد', key: 'date_of_birth', v: (e) => fmt.dmyDash(e.date_of_birth) },
    { h: 'الحالة', key: 'marital_status', v: marAr },
    { h: 'الرقم', key: 'national_id_number', v: (e) => e.national_id_number },
    { h: 'تاريخ الإصدار', key: 'national_id_issued_date', v: (e) => (e.isSA ? fmt.hijri(e.national_id_issued_date) : fmt.dmyDash(e.national_id_issued_date)) },
    { h: 'تاريخ الانتهاء', key: 'national_id_expiry_date', v: (e) => (e.isSA ? fmt.hijri(e.national_id_expiry_date) : fmt.dmyDash(e.national_id_expiry_date)) },
    { h: 'الرقم الوظيفي', key: 'employee_no', v: (e) => e.employee_no },
    { h: 'المسمى', key: 'job_title', v: (e) => e.job_title_ar },
    { h: 'الإدارة', key: 'department', v: (e) => e.deptRep.ar[0] },
    { h: 'تاريخ المباشرة', key: 'join_date', v: (e) => fmt.dmyDash(e.join_date) },
    { h: 'الحالة', key: 'employment_status', v: (e) => e.statusRep.ar[e.i % e.statusRep.ar.length] },
    { h: 'نوع العقد', key: 'contract_type', v: (e) => e.contract_type },
    { h: 'الجوال', key: 'work_phone_number', v: (e) => e.work_phone_number },
    { h: 'الإيميل', key: 'email', v: (e) => e.email },
    { h: 'الأساسي', key: 'salary_basic_salary_amount', v: (e) => e.salary },
    { h: 'البنك', key: 'bank_name', v: (e) => e.bank[1] },
    { h: 'رقم الآيبان', key: 'iban', v: (e) => e.iban },
  ];
  const groups = [{ text: 'البيانات الشخصية', span: 6 }, { text: 'بيانات الهوية / الإقامة', span: 3 }, { text: 'بيانات العمل', span: 6 }, { text: 'التواصل', span: 2 }, { text: 'الراتب والبنك', span: 3 }];
  const map = await writeXlsx('03_two_level_header.xlsx', { sheets: [{ name: 'بيانات الموظفين', title: ['بيان الموظفين'], groups, cols, emps: EMPS }] });
  addTruth('03_two_level_header.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'national_id_issued_date', 'national_id_expiry_date', 'email', 'work_phone_number', 'department', 'join_date', 'salary_basic_salary_amount', 'iban', 'employment_status', 'marital_status', 'employee_no']);
}

/* ======================= 4) بدون ترويسة نهائيًا — المطابقة من المحتوى فقط ======================= */
{
  const cols = [
    { h: '', key: 'employee_no', v: (e) => e.employee_no },
    { h: '', key: 'full_name_ar', v: (e) => e.ar.join(' ') },
    { h: '', key: 'full_name_en', v: (e) => e.en.join(' ') },
    { h: '', key: 'national_id_number', v: (e) => e.national_id_number },
    { h: '', key: 'nationality', v: natAr },
    { h: '', key: 'gender', v: (e) => e.gender },
    { h: '', key: 'date_of_birth', v: (e) => fmt.date(e.date_of_birth) },
    { h: '', key: 'work_phone_number', v: (e) => e.work_phone_number },
    { h: '', key: 'email', v: (e) => e.email },
    { h: '', key: 'department', v: (e) => e.deptRep.ar[0] },
    { h: '', key: 'join_date', v: (e) => fmt.date(e.join_date) },
    { h: '', key: 'salary_basic_salary_amount', v: (e) => e.salary },
    { h: '', key: 'iban', v: (e) => e.iban },
    { h: '', key: 'employment_status', v: (e) => e.statusRep.ar[0] },
  ];
  const map = await writeXlsx('04_no_headers.xlsx', { sheets: [{ name: 'Sheet1', noHeader: true, cols, emps: EMPS }] });
  addTruth('04_no_headers.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'email', 'work_phone_number', 'department', 'join_date', 'salary_basic_salary_amount', 'iban', 'employment_status', 'employee_no']);
}

/* ======================= 5) اسم رباعي بأعمدة منفصلة + عناوين أقسام بين الصفوف + ترويسة مكررة ======================= */
{
  const byDept = [...EMPS].sort((a, b) => a.department.localeCompare(b.department, 'ar'));
  const cols = [
    { h: 'م', v: (e, k) => k + 1 },
    { h: 'الاسم الأول', key: 'first_name_ar', v: (e) => e.ar[0] },
    { h: 'اسم الأب', v: (e) => (e.ar.length > 2 ? e.ar[1] : '') },
    { h: 'اسم الجد', v: () => '' },
    { h: 'اسم العائلة', key: 'last_name_ar', v: (e) => e.ar[e.ar.length - 1] },
    { h: 'First Name', key: 'first_name', v: (e) => e.en[0] },
    { h: 'Family Name', key: 'last_name', v: (e) => e.en[e.en.length - 1] },
    { h: 'الجنسية', key: 'nationality', v: (e) => e.N.code },
    { h: 'الجنس', key: 'gender', v: (e) => (e.g === 'M' ? 'Male' : 'Female') },
    { h: 'تاريخ الميلاد', key: 'date_of_birth', v: (e) => (e.isSA ? fmt.hijri(e.date_of_birth) : fmt.dmyDash(e.date_of_birth)) },
    { h: 'رقم السجل المدني / الإقامة', key: 'national_id_number', v: (e) => e.national_id_number },
    { h: 'الحالة', key: 'marital_status', v: (e) => e.maritalRep.en },
    { h: 'الوضع الوظيفي', key: 'employment_status', v: (e) => e.statusRep.ar[0] },
    { h: 'المسمى الوظيفي', key: 'job_title', v: (e) => e.job_title_ar },
    { h: 'تاريخ الالتحاق', key: 'join_date', v: (e) => fmt.dMonY(e.join_date) },
    { h: 'الجوال', key: 'work_phone_number', v: (e) => `${e.work_phone_number.slice(0, 3)}-${e.work_phone_number.slice(3, 6)}-${e.work_phone_number.slice(6)}` },
    { h: 'البريد', key: 'email', v: (e) => e.email },
    { h: 'الراتب', key: 'salary_basic_salary_amount', v: (e) => `${money(e.salary)}.00` },
    { h: 'IBAN', key: 'iban', v: (e) => ibanSpaced(e.iban) },
  ];
  const map = await writeXlsx('05_split_names_sections.xlsx', { sheets: [{ name: 'الكشف', title: ['كشف الموظفين حسب الإدارات'], cols, emps: byDept, sections: (e) => `إدارة ${e.department}`, repeatHeaderAt: 20 }] });
  addTruth('05_split_names_sections.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'email', 'work_phone_number', 'department', 'join_date', 'salary_basic_salary_amount', 'iban', 'employment_status', 'marital_status'], {
    over: (e) => ({ last_name_ar: e.ar[e.ar.length - 1], last_name: e.en[e.en.length - 1], first_name_ar: e.ar[0], first_name: e.en[0] }),
    synthetic: { department: '@group' },
  });
}

/* ======================= 6) شيت لكل فرع + شيت ملخص + شيت ملاحظات ======================= */
{
  const mk = (order) => {
    const all = [
      { h: 'اسم الموظف', key: 'full_name_ar', v: (e) => e.ar.join(' ') },
      { h: 'Employee Name', key: 'full_name_en', v: (e) => e.en.join(' ') },
      { h: 'الجنسية', key: 'nationality', v: (e) => e.N.en },
      { h: 'الجنس', key: 'gender', v: (e) => e.gender },
      { h: 'تاريخ الميلاد', key: 'date_of_birth', v: (e) => fmt.date(e.date_of_birth) },
      { h: 'رقم الهوية', key: 'national_id_number', v: (e) => (e.i % 4 === 0 ? `'${e.national_id_number}` : e.national_id_number) },
      { h: 'الايميل', key: 'email', v: (e) => e.email },
      { h: 'رقم الجوال', key: 'work_phone_number', v: (e) => `00966${e.work_phone_number.slice(1)}` },
      { h: 'الوظيفة', key: 'job_title', v: (e) => e.job_title_ar },
      { h: 'القسم', key: 'department', v: (e) => e.deptRep.ar[0] },
      { h: 'تاريخ التعيين', key: 'join_date', v: (e) => fmt.date(e.join_date) },
      { h: 'الراتب الأساسي', key: 'salary_basic_salary_amount', v: (e) => e.salary },
      { h: 'الآيبان', key: 'iban', v: (e) => e.iban },
      { h: 'حالة الموظف', key: 'employment_status', v: (e) => e.statusRep.ar[0] },
    ];
    return order ? order.map((i) => all[i]) : all;
  };
  const sheets = LOCS.map((L, li) => ({
    name: L.sheet, cols: mk(li === 2 ? [1, 0, 2, 3, 5, 4, 7, 6, 8, 9, 10, 12, 11, 13] : null),
    emps: EMPS.filter((e) => e.location === L.canon), skipExpect: li > 0, title: [`موظفو فرع ${L.ar}`],
  }));
  const extraSheets = [
    { name: 'ملاحظات', rows: [['ملاحظات الموارد البشرية'], ['تم تحديث البيانات بتاريخ 2026-10-01'], ['يرجى مراجعة الإقامات المنتهية']] },
  ];
  const wbSpec = { sheets: [{ name: 'ملخص', cols: [{ h: 'الفرع', v: (e) => e.l }, { h: 'عدد الموظفين', v: (e) => e.n }], emps: LOCS.map((L) => ({ l: L.ar, n: EMPS.filter((e) => e.location === L.canon).length })), skipExpect: true }, ...sheets], extraSheets };
  const map = await writeXlsx('06_multi_sheet_branches.xlsx', wbSpec);
  addTruth('06_multi_sheet_branches.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'email', 'work_phone_number', 'department', 'join_date', 'salary_basic_salary_amount', 'iban', 'employment_status', 'location'], {
    over: (e) => ({ work_phone_number: e.work_phone_number.slice(-9) }), synthetic: { location: '@sheet' },
  });
}

/* ======================= 7) CSV بترميز Windows-1256 وفاصلة منقوطة ======================= */
{
  const cols = [
    ['الرقم الوظيفي', 'employee_no', (e) => e.employee_no],
    ['الاسم', 'full_name_ar', (e) => e.ar.join(' ')],
    ['الاسم باللغة الانجليزية', 'full_name_en', (e) => e.en.join(' ')],
    ['الجنس', 'gender', (e) => e.gender],
    ['الجنسية', 'nationality', natAr],
    ['رقم الاقامة', 'national_id_number', (e) => e.national_id_number],
    ['تاريخ الميلاد', 'date_of_birth', (e) => fmt.dmy(e.date_of_birth)],
    ['تاريخ الانضمام', 'join_date', (e) => fmt.dmy(e.join_date)],
    ['الجوال', 'work_phone_number', (e) => e.work_phone_number],
    ['الايميل', 'email', (e) => e.email],
    ['القسم', 'department', (e) => e.deptRep.ar[0]],
    ['الراتب', 'salary_basic_salary_amount', (e) => money(e.salary)],
    ['الايبان', 'iban', (e) => e.iban],
  ];
  const q = (s) => (/[;"\n,]/.test(String(s)) ? `"${String(s).replace(/"/g, '""')}"` : String(s));
  const text = [cols.map((c) => c[0]).join(';'), ...EMPS.map((e) => cols.map((c) => q(c[2](e))).join(';'))].join('\r\n');
  // ترميز cp1256 يدويًا (بقلب جدول فك الترميز)
  const dec = new TextDecoder('windows-1256');
  const enc = new Map();
  for (let b = 0; b < 256; b++) enc.set(dec.decode(new Uint8Array([b])), b);
  const bytes = Uint8Array.from([...text].map((ch) => { const b = enc.get(ch); if (b === undefined) throw new Error('cp1256 لا يدعم ' + ch); return b; }));
  fs.writeFileSync(new URL('07_windows1256.csv', OUT), bytes);
  const map = Object.fromEntries(cols.map((c, j) => [c[1], j]));
  addTruth('07_windows1256.csv', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'email', 'work_phone_number', 'department', 'join_date', 'salary_basic_salary_amount', 'iban', 'employee_no']);
}

/* ======================= 8) كل الفوضى: اسم عربي+إنجليزي بخلية، رقمين بخلية، جواز سفر، أعمدة فاضية، دمج عمودي ======================= */
{
  const byDept = [...EMPS].sort((a, b) => a.department.localeCompare(b.department, 'ar'));
  const cols = [
    { h: 'القسم', key: 'department', v: (e) => e.deptRep.ar[0] },
    { h: 'الاسم / Name', key: 'full_name_ar', v: (e) => `${e.ar.join(' ')} - ${e.en.join(' ')}` },
    { h: 'رقم الجواز', v: (e) => `P${e.national_id_number.slice(2, 9)}` },
    { h: 'انتهاء الجواز', v: (e) => fmt.date(addYears(e.national_id_expiry_date, 1)) },
    { h: 'رقم الهوية\n(10 أرقام) *', key: 'national_id_number', v: (e) => (e.i % 3 === 0 ? `'${e.national_id_number}` : `${e.national_id_number.slice(0, 4)} ${e.national_id_number.slice(4)}`) },
    { h: 'تاريخ\nالميلاد *', key: 'date_of_birth', v: (e, k) => [fmt.withTime, fmt.dMonY, fmt.dmy][k % 3](e.date_of_birth) },
    { h: 'Nationality / الجنسية', key: 'nationality', v: (e, k) => (k % 2 ? e.N.en : natAr(e)) },
    { h: 'Gender', key: 'gender', v: (e, k) => (k % 2 ? (e.g === 'M' ? 'Male' : 'Female') : e.gender) },
    { h: 'أرقام التواصل', key: 'work_phone_number', v: (e) => `${e.work_phone_number} / 0114567890` },
    { h: 'E-mail', key: 'email', v: (e, k) => (k % 4 === 0 ? `mailto:${e.email}` : ` ${e.email} `) },
    { h: 'تاريخ مباشرة العمل', key: 'join_date', v: (e) => fmt.ymdSlash(e.join_date) },
    { h: 'الراتب الأساسي (ريال)', key: 'salary_basic_salary_amount', v: (e) => `SAR ${money(e.salary)}.00` },
    { h: 'بدل السكن', v: (e) => Math.round(e.salary * 0.25) },
    { h: 'إجمالي الراتب', v: (e) => Math.round(e.salary * 1.25) },
    { h: 'رقم الحساب الدولي (IBAN)', key: 'iban', v: (e, k) => (k % 2 ? `IBAN: ${ibanSpaced(e.iban)}` : e.iban.toLowerCase()) },
    { h: 'ملاحظات', v: (e, k) => (k % 7 === 0 ? 'تم التجديد' : '') },
  ];
  const map = await writeXlsx('08_kitchen_sink.xlsx', { sheets: [{ name: 'data', startCol: 2, gapAt: 9, cols, emps: byDept, mergeKey: 'department', title: ['بيانات الموظفين (نسخة غير مدققة)'] }] });
  addTruth('08_kitchen_sink.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'email', 'work_phone_number', 'department', 'join_date', 'salary_basic_salary_amount', 'iban']);
}

/* ======================= 9) قالب قيود نفسه لكن بترتيب مختلف وأعمدة ناقصة ======================= */
{
  const cols = [
    { h: 'national_id_number *', key: 'national_id_number', v: (e) => e.national_id_number },
    { h: 'first_name_ar *', key: 'first_name_ar', v: (e) => e.first_name_ar },
    { h: 'last_name_ar *', key: 'last_name_ar', v: (e) => e.last_name_ar },
    { h: 'first_name *', key: 'first_name', v: (e) => e.first_name },
    { h: 'last_name *', key: 'last_name', v: (e) => e.last_name },
    { h: 'gender *', key: 'gender', v: (e) => e.gender },
    { h: 'nationality *', key: 'nationality', v: (e) => e.nationality },
    { h: 'date_of_birth *', key: 'date_of_birth', v: (e) => fmt.date(e.date_of_birth) },
    { h: 'email *', key: 'email', v: (e) => e.email },
    { h: 'work_phone_number *', key: 'work_phone_number', v: (e) => `+966${e.work_phone_number.slice(1)}` },
    { h: 'department *', key: 'department', v: (e) => e.department },
    { h: 'join_date *', key: 'join_date', v: (e) => fmt.date(e.join_date) },
    { h: 'salary_basic_salary_amount', key: 'salary_basic_salary_amount', v: (e) => e.salary },
    { h: 'iban', key: 'iban', v: (e) => e.iban },
    { h: 'employment_status *', key: 'employment_status', v: (e) => e.employment_status },
    { h: 'marital_status *', key: 'marital_status', v: (e) => e.marital_status },
  ];
  const map = await writeXlsx('09_qoyod_keys_shuffled.xlsx', { sheets: [{ name: 'Employees', rtl: false, cols, emps: EMPS }] });
  addTruth('09_qoyod_keys_shuffled.xlsx', map, ['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'gender', 'nationality', 'national_id_number', 'date_of_birth', 'email', 'work_phone_number', 'department', 'join_date', 'salary_basic_salary_amount', 'iban', 'employment_status', 'marital_status']);
}

/* ======================= 10) ملف واقعي فيه أخطاء بيانات متعمدة (لتجربة خطوة المراجعة — ليس ضمن اختبار الدقة) ======================= */
{
  const E = EMPS.map((e) => ({ ...e }));
  const errs = [];
  const bad = (i, what, f) => { f(E[i]); errs.push([E[i].ar.join(' '), what]); };
  bad(2, 'رقم هوية مكرر مع موظف آخر', (e) => { e.national_id_number = E[0].national_id_number; });
  bad(5, 'بريد بدون نطاق', (e) => { e.email = e.email.replace('.sa', ''); });
  bad(8, 'الاسم الإنجليزي فاضي', (e) => { e.en = ['']; });
  bad(13, 'الإقامة منتهية', (e) => { e.national_id_expiry_date = '2025-03-01'; });
  bad(16, 'قسم غير موجود بقائمة القالب ("الأمن والسلامة")', (e) => { e.deptRep = { ar: ['الأمن والسلامة'] }; });
  bad(19, 'هوية 9 خانات', (e) => { e.national_id_number = e.national_id_number.slice(0, 9); });
  bad(22, 'راتب صفر', (e) => { e.salary = 0; });
  bad(25, 'جوال ناقص رقم', (e) => { e.work_phone_number = e.work_phone_number.slice(0, 9); });
  bad(27, 'غير سعودي بدون تاريخ نهاية عقد', (e) => { e.contract_end_date = null; });
  bad(30, 'تاريخ ميلاد مستقبلي (خطأ كتابة 2099)', (e) => { e.date_of_birth = '2099-05-10'; });
  bad(33, 'الاسم العربي مكتوب بالإنجليزي', (e) => { e.ar = e.en; });
  bad(35, 'رقم وظيفي مكرر', (e) => { e.employee_no = E[34].employee_no; });
  bad(38, 'بداية العقد بعد نهايته', (e) => { e.contract_end_date = '2015-01-01'; e.join_date = '2020-01-01'; });
  const cols = [
    { h: 'م', v: (e, k) => k + 1 },
    { h: 'اسم الموظف', v: (e) => e.ar.join(' ') },
    { h: 'Employee Name', v: (e) => e.en.join(' ') },
    { h: 'الجنسية', v: natAr },
    { h: 'الجنس', v: (e, k) => (k % 3 ? e.gender : e.g) },
    { h: 'الحالة الاجتماعية', v: marAr },
    { h: 'تاريخ الميلاد', v: (e, k) => (e.isSA && k % 2 ? fmt.hijri(e.date_of_birth) : fmt.dmy(e.date_of_birth)) },
    { h: 'رقم الهوية/الإقامة', v: (e) => e.national_id_number },
    { h: 'تاريخ انتهاء الهوية', v: (e) => fmt.date(e.national_id_expiry_date) },
    { h: 'الايميل', v: (e) => e.email },
    { h: 'الجوال', v: (e) => e.work_phone_number },
    { h: 'الوظيفة', v: (e) => e.job_title_ar },
    { h: 'القسم', v: (e) => e.deptRep.ar[0] },
    { h: 'الفرع', v: (e) => e.locRep.ar },
    { h: 'تاريخ المباشرة', v: (e) => fmt.date(e.join_date) },
    { h: 'نوع العقد', v: (e) => (e.isSA ? 'غير محدد' : 'محدد') },
    { h: 'نهاية العقد', v: (e) => (e.contract_end_date ? fmt.date(e.contract_end_date) : null) },
    { h: 'الراتب الأساسي', v: (e) => e.salary },
    { h: 'بدل السكن', v: (e) => Math.round(e.salary * 0.25) },
    { h: 'الآيبان', v: (e) => ibanSpaced(e.iban) },
    { h: 'رقم الموظف', v: (e) => e.employee_no },
    { h: 'المؤهل', v: (e) => e.education.ar },
    { h: 'الحالة', v: (e) => e.statusRep.ar[0] },
  ];
  await writeXlsx('10_realistic_with_errors.xlsx', { sheets: [{ name: 'الموظفين', title: ['مؤسسة الأفق — كشف الموظفين'], cols, emps: E, totals: true }] });
  fs.writeFileSync(new URL('10_errors_list.txt', OUT), errs.map(([n, w]) => `${n}: ${w}`).join('\n'));
}

/* ======================= 11) أسماء بلغة وحدة: نص الموظفين اسمهم إنجليزي فقط والباقي عربي فقط (لتجربة الترجمة الحرفية) ======================= */
{
  const cols = [
    { h: 'الرقم الوظيفي', v: (e) => e.employee_no },
    { h: 'الاسم العربي', v: (e) => (e.isSA || e.i % 2 ? e.ar.join(' ') : '') },
    { h: 'English Name', v: (e) => (e.isSA || e.i % 2 ? '' : e.en.join(' ')) },
    { h: 'الجنسية', v: natAr },
    { h: 'الجنس', v: (e) => e.gender },
    { h: 'الحالة الاجتماعية', v: marAr },
    { h: 'تاريخ الميلاد', v: (e) => fmt.date(e.date_of_birth) },
    { h: 'رقم الهوية/الإقامة', v: (e) => e.national_id_number },
    { h: 'تاريخ إصدار الهوية', v: (e) => fmt.date(e.national_id_issued_date) },
    { h: 'تاريخ انتهاء الهوية', v: (e) => fmt.date(e.national_id_expiry_date) },
    { h: 'الايميل', v: (e) => e.email },
    { h: 'الجوال', v: (e) => e.work_phone_number },
    { h: 'المسمى الوظيفي', v: (e) => e.job_title_ar },
    { h: 'القسم', v: (e) => e.deptRep.ar[0] },
    { h: 'تاريخ التعيين', v: (e) => fmt.date(e.join_date) },
    { h: 'نوع التوظيف', v: (e) => e.employment_type },
    { h: 'نهاية العقد', v: (e) => (e.contract_end_date ? fmt.date(e.contract_end_date) : null) },
    { h: 'الحالة', v: (e) => e.statusRep.ar[0] },
  ];
  await writeXlsx('11_names_one_language.xlsx', { sheets: [{ name: 'الموظفين', cols, emps: EMPS }] });
}

fs.writeFileSync(new URL('truth.json', OUT), JSON.stringify(truth, null, 1));
console.log('fixtures:', Object.keys(truth).join(', '));
