// Blind-test fixture generator: realistic messy Saudi/GCC client HR files + ground truth.
// Run from the test folder:  node make_blind.mjs   -> writes ./fixtures_blind/*.xlsx|csv + truth.json
// All people, IDs, IBANs, phones and emails are fake (deterministic PRNG).
import ExcelJS from 'exceljs';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUT = path.join(__dirname, 'fixtures_blind');
fs.mkdirSync(OUT, { recursive: true });

// ---------------------------------------------------------------- PRNG / helpers
function mulberry32(a) { return function () { a |= 0; a = a + 0x6D2B79F5 | 0; let t = Math.imul(a ^ a >>> 15, 1 | a); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
const rng = mulberry32(20261022);
const rnd = () => rng();
const ri = (a, b) => a + Math.floor(rnd() * (b - a + 1));
const pick = (arr) => arr[Math.floor(rnd() * arr.length)];
const chance = (p) => rnd() < p;
function weighted(obj) { const ks = Object.keys(obj); const tot = ks.reduce((s, k) => s + obj[k], 0); let r = rnd() * tot; for (const k of ks) { r -= obj[k]; if (r < 0) return k; } return ks[ks.length - 1]; }
const pad2 = (n) => String(n).padStart(2, '0');
const AR_DIGITS = '٠١٢٣٤٥٦٧٨٩';
const toArDigits = (s) => String(s).replace(/\d/g, (d) => AR_DIGITS[d]);
const thousands = (n) => n.toLocaleString('en-US');

// ---------------------------------------------------------------- IDs / IBAN
const usedIds = new Set();
function luhnOk(id) { let s = 0; for (let i = 0; i < 10; i++) { let v = +id[i]; if (i % 2 === 0) { v *= 2; if (v > 9) v -= 9; } s += v; } return s % 10 === 0; }
function makeId(first) {
  for (;;) {
    const d = [first]; for (let i = 0; i < 8; i++) d.push(ri(0, 9));
    let s = 0; for (let i = 0; i < 9; i++) { let v = d[i]; if (i % 2 === 0) { v *= 2; if (v > 9) v -= 9; } s += v; }
    d.push((10 - s % 10) % 10);
    const id = d.join('');
    if (!usedIds.has(id) && luhnOk(id)) { usedIds.add(id); return id; }
  }
}
function makeIban(bankCode) {
  const acct = Array.from({ length: 18 }, () => ri(0, 9)).join('');
  const bban = bankCode + acct;
  const chk = 98n - (BigInt(bban + '281000') % 97n); // 'SA00' -> S=28 A=10 00
  return 'SA' + String(chk).padStart(2, '0') + bban;
}
function ibanOk(iban) { const r = iban.slice(4) + iban.slice(0, 4); return BigInt(r.replace(/[A-Z]/g, (c) => c.charCodeAt(0) - 55)) % 97n === 1n; }
const ibanSpaced = (i) => i.replace(/(.{4})/g, '$1 ').trim();

// ---------------------------------------------------------------- dates
function randDate(y1, y2) { const a = Date.UTC(y1, 0, 1), b = Date.UTC(y2, 11, 31); const days = Math.round((b - a) / 86400000); return new Date(a + ri(0, days) * 86400000).toISOString().slice(0, 10); }
const ymd = (s) => s.split('-').map(Number);
const jsDate = (s) => { const [y, m, d] = ymd(s); return new Date(Date.UTC(y, m - 1, d)); };
const serial = (s) => jsDate(s).getTime() / 86400000 + 25569;
const dmy = (s, sep = '/') => { const [y, m, d] = ymd(s); return `${pad2(d)}${sep}${pad2(m)}${sep}${y}`; };
const mdy = (s) => { const [y, m, d] = ymd(s); return `${pad2(m)}/${pad2(d)}/${y}`; };
const compact = (s) => s.replace(/-/g, '');
const MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const dMonYY = (s) => { const [y, m, d] = ymd(s); return `${pad2(d)}-${MON[m - 1]}-${String(y).slice(2)}`; };
const dMonYYYY = (s) => { const [y, m, d] = ymd(s); return `${d} ${MON[m - 1]} ${y}`; };
const hijriFmt = new Intl.DateTimeFormat('en-u-ca-islamic-umalqura-nu-latn', { timeZone: 'UTC', year: 'numeric', month: 'numeric', day: 'numeric' });
function hijri(s) { const p = hijriFmt.formatToParts(jsDate(s)); const g = (t) => Number(p.find((x) => x.type === t).value); return { y: g('year'), m: g('month'), d: g('day') }; }
const firstOfNextMonth = (s) => { const [y, m] = ymd(s); const d = new Date(Date.UTC(y, m, 1)); return d.toISOString().slice(0, 10); };
const addYears = (s, n) => { const [y, m, d] = ymd(s); return `${y + n}-${pad2(m)}-${pad2(Math.min(d, 28))}`; };

// ---------------------------------------------------------------- phones
function phone(m, style) {
  const a = m.slice(0, 2), b = m.slice(2, 5), c = m.slice(5);
  switch (style) {
    case '05sp': return `0${a} ${b} ${c}`;
    case '05dash': return `0${a}-${b}-${c}`;
    case '05': return '0' + m;
    case 'plus': return `+966 ${a} ${b} ${c}`;
    case 'plusnosp': return '+966' + m;
    case '00966': return `00966 ${a}-${b}-${c}`;
    case '00966nosp': return '00966' + m;
    case '966': return '966' + m;
    case 'num': return Number(m);
    case 'paren': return `+966 (${a}) ${b}-${c}`;
    case 'plusdash': return `+966-${a}-${b}-${c}`;
    case 'sp9': return `${a} ${b} ${c}`;
    case '966dash': return `966-${a}-${b}-${c}`;
    default: throw new Error(style);
  }
}

// ---------------------------------------------------------------- reference data
const NAT = {
  SA: { ar: 'السعودية', m: 'سعودي', f: 'سعودية', en: 'Saudi', iso: 'SA', country: 'المملكة العربية السعودية', kind: 'saudi' },
  EG: { ar: 'مصر', m: 'مصري', f: 'مصرية', en: 'Egyptian', iso: 'EG', country: 'مصر', kind: 'arab', fam: 'EG' },
  JO: { ar: 'الأردن', m: 'أردني', f: 'أردنية', en: 'Jordanian', iso: 'JO', country: 'الأردن', kind: 'arab', fam: 'LEV' },
  SY: { ar: 'سوريا', m: 'سوري', f: 'سورية', en: 'Syrian', iso: 'SY', country: 'سوريا', kind: 'arab', fam: 'LEV' },
  LB: { ar: 'لبنان', m: 'لبناني', f: 'لبنانية', en: 'Lebanese', iso: 'LB', country: 'لبنان', kind: 'arab', fam: 'LEV' },
  SD: { ar: 'السودان', m: 'سوداني', f: 'سودانية', en: 'Sudanese', iso: 'SD', country: 'السودان', kind: 'arab', fam: 'SD' },
  YE: { ar: 'اليمن', m: 'يمني', f: 'يمنية', en: 'Yemeni', iso: 'YE', country: 'اليمن', kind: 'arab', fam: 'YE' },
  IN: { ar: 'الهند', m: 'هندي', f: 'هندية', en: 'Indian', iso: 'IN', country: 'الهند', kind: 'other' },
  PK: { ar: 'باكستان', m: 'باكستاني', f: 'باكستانية', en: 'Pakistani', iso: 'PK', country: 'باكستان', kind: 'other' },
  BD: { ar: 'بنغلاديش', m: 'بنغلاديشي', f: 'بنغلاديشية', en: 'Bangladeshi', iso: 'BD', country: 'بنغلاديش', kind: 'other' },
  PH: { ar: 'الفلبّين', m: 'فلبيني', f: 'فلبينية', en: 'Filipino', enF: 'Filipina', iso: 'PH', country: 'الفلبين', kind: 'other' },
  NP: { ar: 'نيبال', m: 'نيبالي', f: 'نيبالية', en: 'Nepali', iso: 'NP', country: 'نيبال', kind: 'other' },
  LK: { ar: 'سريلانكا', m: 'سريلانكي', f: 'سريلانكية', en: 'Sri Lankan', iso: 'LK', country: 'سريلانكا', kind: 'other' },
  GB: { ar: 'المملكة المتحدة', m: 'بريطاني', f: 'بريطانية', en: 'British', iso: 'GB', country: 'بريطانيا', kind: 'other' },
  US: { ar: 'الولايات المتّحدة', m: 'أمريكي', f: 'أمريكية', en: 'American', iso: 'US', country: 'الولايات المتحدة الأمريكية', kind: 'other' },
};

const SA_M = [['محمد', 'Mohammed'], ['أحمد', 'Ahmed'], ['عبدالله', 'Abdullah'], ['خالد', 'Khalid'], ['فهد', 'Fahad'], ['سعد', 'Saad'], ['عبدالرحمن', 'Abdulrahman'], ['فيصل', 'Faisal'], ['سلطان', 'Sultan'], ['ناصر', 'Nasser'], ['تركي', 'Turki'], ['عمر', 'Omar'], ['يوسف', 'Yousef'], ['ماجد', 'Majed'], ['بندر', 'Bandar'], ['علي', 'Ali'], ['إبراهيم', 'Ibrahim'], ['عبد العزيز', 'Abdul Aziz'], ['عبد الرحمن', 'Abdul Rahman'], ['عبد الإله', 'Abdul Elah'], ['نايف', 'Naif'], ['مشعل', 'Mishal'], ['راكان', 'Rakan']];
const SA_F = [['نورة', 'Noura'], ['سارة', 'Sarah'], ['ريم', 'Reem'], ['هند', 'Hind'], ['لمى', 'Lama'], ['أمل', 'Amal'], ['منى', 'Mona'], ['فاطمة', 'Fatimah'], ['العنود', 'Alanoud'], ['دانة', 'Dana'], ['شهد', 'Shahad'], ['مها', 'Maha'], ['أريج', 'Areej'], ['جواهر', 'Jawaher']];
const SA_FATHER = SA_M.filter(([a]) => !a.includes(' '));
const SA_FAM = [['القحطاني', 'Al-Qahtani'], ['الشهري', 'Al-Shehri'], ['العتيبي', 'Al-Otaibi'], ['الدوسري', 'Al-Dosari'], ['الغامدي', 'Al-Ghamdi'], ['الزهراني', 'Al-Zahrani'], ['الحربي', 'Al-Harbi'], ['المطيري', 'Al-Mutairi'], ['السبيعي', 'Al-Subaie'], ['العنزي', 'Al-Anazi'], ['الشمري', 'Al-Shammari'], ['البقمي', 'Al-Buqami'], ['الرشيد', 'Al-Rasheed'], ['التميمي', 'Al-Tamimi'], ['الجهني', 'Al-Juhani'], ['السهلي', 'Al-Sahli']];
const ARAB_M = [['محمد', 'Mohamed'], ['أحمد', 'Ahmed'], ['محمود', 'Mahmoud'], ['مصطفى', 'Mostafa'], ['هاني', 'Hany'], ['كريم', 'Karim'], ['عمرو', 'Amr'], ['طارق', 'Tarek'], ['سامر', 'Samer'], ['وليد', 'Waleed'], ['حسام', 'Hossam'], ['ياسر', 'Yasser']];
const ARAB_F = [['رنا', 'Rana'], ['دينا', 'Dina'], ['ياسمين', 'Yasmin'], ['لينا', 'Lina'], ['هبة', 'Heba'], ['نور', 'Nour'], ['رهف', 'Rahaf']];
const ARAB_FAM = {
  EG: [['عبدالفتاح', 'Abdelfattah'], ['السيد', 'El-Sayed'], ['منصور', 'Mansour'], ['فهمي', 'Fahmy'], ['الشافعي', 'El-Shafei'], ['رمضان', 'Ramadan']],
  LEV: [['الخطيب', 'Al-Khatib'], ['حداد', 'Haddad'], ['النجار', 'Al-Najjar'], ['العلي', 'Al-Ali'], ['خوري', 'Khoury'], ['الحسيني', 'Al-Husseini']],
  SD: [['عثمان', 'Osman'], ['الأمين', 'Elamin'], ['إدريس', 'Idris'], ['الطيب', 'Eltayeb']],
  YE: [['العمودي', 'Al-Amoudi'], ['باوزير', 'Bawazir'], ['الحضرمي', 'Al-Hadrami'], ['باعشن', 'Baeshen']],
};
const OTHER_NAMES = {
  IN: { m: [['راجيش', 'Rajesh'], ['سوريش', 'Suresh'], ['أرجون', 'Arjun'], ['فينود', 'Vinod'], ['أنيل', 'Anil'], ['ساجد', 'Sajid']], f: [['بريا', 'Priya'], ['أنجالي', 'Anjali'], ['ديفيا', 'Divya']], mid: [], last: [['كومار', 'Kumar'], ['ناير', 'Nair'], ['مينون', 'Menon'], ['شارما', 'Sharma'], ['بيلاي', 'Pillai']] },
  PK: { m: [['محمد', 'Muhammad'], ['عمران', 'Imran'], ['عبد الرحمن', 'Abdul Rahman'], ['عبد القادر', 'Abdul Qadir'], ['طارق', 'Tariq'], ['بلال', 'Bilal'], ['عبد الغفور', 'Abdul Ghafoor']], f: [['عائشة', 'Ayesha'], ['صبا', 'Saba']], mid: [['آصف', 'Asif'], ['نواز', 'Nawaz'], ['شهزاد', 'Shahzad']], last: [['خان', 'Khan'], ['إقبال', 'Iqbal'], ['بلوش', 'Baloch'], ['حسين', 'Hussain'], ['شيخ', 'Sheikh']] },
  BD: { m: [['رفيق', 'Rafiq'], ['كريم', 'Karim'], ['شاهد', 'Shahid'], ['جمال', 'Jamal']], f: [['نسرين', 'Nasrin']], mid: [['الدين', 'Uddin']], last: [['إسلام', 'Islam'], ['رحمن', 'Rahman'], ['حسين', 'Hossain'], ['ميا', 'Mia']] },
  PH: { m: [['مارك', 'Mark'], ['جوزيه', 'Jose'], ['رامون', 'Ramon'], ['جون', 'John']], f: [['ماريا', 'Maria'], ['جينيفر', 'Jennifer'], ['أنجيليكا', 'Angelica']], mid: [['أنتوني', 'Anthony'], ['كلارا', 'Clara']], last: [['رييس', 'Reyes'], ['سانتوس', 'Santos'], ['غارسيا', 'Garcia'], ['باوتيستا', 'Bautista'], ['ديلا كروز', 'Dela Cruz']] },
  NP: { m: [['رام', 'Ram'], ['بيشال', 'Bishal'], ['سانجاي', 'Sanjay']], f: [['سيتا', 'Sita']], mid: [], last: [['تامانغ', 'Tamang'], ['غورونغ', 'Gurung'], ['ثابا', 'Thapa']] },
  LK: { m: [['نيروشان', 'Niroshan'], ['كاسون', 'Kasun']], f: [['دلشاني', 'Dilshani']], mid: [], last: [['بيريرا', 'Perera'], ['فرناندو', 'Fernando']] },
  GB: { m: [['جيمس', 'James'], ['ديفيد', 'David']], f: [['إميلي', 'Emily']], mid: [], last: [['سميث', 'Smith'], ['براون', 'Brown'], ['ويلسون', 'Wilson']] },
  US: { m: [['مايكل', 'Michael'], ['روبرت', 'Robert']], f: [['جيسيكا', 'Jessica']], mid: [], last: [['جونسون', 'Johnson'], ['ميلر', 'Miller']] },
};

const DEPT = {
  HR: { ar: 'الموارد البشرية', code: 'HR', en: 'Human Resources', arv: 'الموارد البشريه', titles: [['أخصائي موارد بشرية', 'HR Specialist'], ['مدير الموارد البشرية', 'HR Manager'], ['أخصائي توظيف', 'Recruiter']] },
  FIN: { ar: 'المالية', code: 'FIN', en: 'Finance', arv: 'إدارة المالية', titles: [['محاسب', 'Accountant'], ['محاسب أول', 'Senior Accountant'], ['مدير مالي', 'Finance Manager']] },
  ENG: { ar: 'الهندسة', code: 'ENG', en: 'Engineering', arv: 'الهندسه', titles: [['مهندس مدني', 'Civil Engineer'], ['مهندس موقع', 'Site Engineer'], ['مهندس كهرباء', 'Electrical Engineer']] },
  SAL: { ar: 'المبيعات', code: 'SLS', en: 'Sales', arv: 'مبيعات', titles: [['مندوب مبيعات', 'Sales Representative'], ['مدير مبيعات', 'Sales Manager'], ['مسؤول حسابات', 'Account Executive']] },
  MKT: { ar: 'التسويق', code: 'MKT', en: 'Marketing', arv: 'تسويق', titles: [['أخصائي تسويق', 'Marketing Specialist'], ['مصمم جرافيك', 'Graphic Designer'], ['مدير تسويق', 'Marketing Manager']] },
  OPS: { ar: 'العمليات', code: 'OPS', en: 'Operations', arv: 'إدارة العمليات', titles: [['مشرف عمليات', 'Operations Supervisor'], ['منسق عمليات', 'Operations Coordinator'], ['مراقب', 'Foreman']] },
  IT: { ar: 'تقنية المعلومات', code: 'IT', en: 'Information Technology', arv: 'تقنية معلومات', titles: [['مطور برمجيات', 'Software Developer'], ['مهندس شبكات', 'Network Engineer'], ['دعم فني', 'IT Support']] },
  CS: { ar: 'خدمة العملاء', code: 'CS', en: 'Customer Service', arv: 'خدمة عملاء', titles: [['موظف خدمة عملاء', 'Customer Service Agent'], ['مشرف خدمة عملاء', 'CS Supervisor']] },
  PRC: { ar: 'المشتريات', code: 'PROC', en: 'Procurement', arv: 'مشتريات', titles: [['أخصائي مشتريات', 'Procurement Specialist'], ['مسؤول مشتريات', 'Buyer']] },
  ADM: { ar: 'الشؤون الإدارية', code: 'ADMIN', en: 'Administration', arv: 'الشئون الإدارية', titles: [['سكرتير تنفيذي', 'Executive Secretary'], ['مسؤول شؤون إدارية', 'Admin Officer'], ['مندوب حكومي', 'Government Relations Officer']] },
};
const LOC = { RUH: 'المقر الرئيسي - الرياض', JED: 'فرع جدة', DMM: 'فرع الدمام' };
const MAR = { S: 'أعزب', M: 'متزوج', D: 'مطلق', W: 'أرمل' };
const MAR_AR = { M: { S: 'أعزب', M: 'متزوج', D: 'مطلق', W: 'أرمل' }, F: { S: 'عزباء', M: 'متزوجة', D: 'مطلقة', W: 'أرملة' } };
const MAR_EN = { S: 'Single', M: 'Married', D: 'Divorced', W: 'Widowed' };
const STAT = { A: 'مفعّل', I: 'غير مفعّل', N: 'فترة الإشعار' };
const EMPT = { FT: 'دوام كامل', PT: 'دوام جزئي' };
const CONT = { F: 'محدد المدة', U: 'غير محدد المدة' };
const ATT = { D: 'افتراضي', E: 'دوام مسائي' };
const EDU = { HS: { ar: 'شهادة ثانوية', raw: 'ثانوية عامة', en: 'High School' }, BA: { ar: 'درجة البكالوريوس', raw: 'بكالوريوس', en: "Bachelor's" }, MA: { ar: 'درجة ماجستير', raw: 'ماجستير', en: "Master's" }, PHD: { ar: 'درجة دكتوراه', raw: 'دكتوراه', en: 'PhD' } };
const MAJORS = ['محاسبة', 'إدارة أعمال', 'علوم الحاسب', 'نظم المعلومات', 'تسويق', 'هندسة مدنية', 'هندسة كهربائية', 'لغة إنجليزية', 'اقتصاد'];
const BANKS = {
  '80': { en: 'Al Rajhi Bank', ar: 'مصرف الراجحي', bic: 'RJHISARI' },
  '10': { en: 'Saudi National Bank', ar: 'البنك الأهلي السعودي', bic: 'NCBKSAJE' },
  '20': { en: 'Riyad Bank', ar: 'بنك الرياض', bic: 'RIBLSARI' },
  '45': { en: 'SAB', ar: 'البنك السعودي الأول', bic: 'SABBSARI' },
  '05': { en: 'Alinma Bank', ar: 'مصرف الإنماء', bic: 'INMASARI' },
  '15': { en: 'Bank AlBilad', ar: 'بنك البلاد', bic: 'ALBISARI' },
  '55': { en: 'Banque Saudi Fransi', ar: 'البنك السعودي الفرنسي', bic: 'BSFRSARI' },
  '30': { en: 'Arab National Bank', ar: 'البنك العربي الوطني', bic: 'ARNBSARI' },
};
const RELS_M = ['أخ', 'أب', 'أم', 'زوجة', 'أخت', 'ابن'];
const RELS_F = ['أخ', 'أب', 'أم', 'زوج', 'أخت'];
const DISTRICTS = { RUH: ['الرياض، حي النرجس', 'الرياض، حي الملقا', 'الرياض، حي السليمانية', 'الرياض، حي الروضة'], JED: ['جدة، حي الصفا', 'جدة، حي الروضة', 'جدة، حي النعيم'], DMM: ['الدمام، حي الفيصلية', 'الدمام، حي الشاطئ', 'الخبر، حي العقربية'] };
const STREETS = ['شارع الأمير سلطان', 'شارع التحلية', 'شارع أنس بن مالك', 'شارع الملك فهد', 'شارع عثمان بن عفان', 'شارع الأمير محمد بن سلمان'];

// ---------------------------------------------------------------- employee factory
const usedEmails = new Set();
const ascii = (s) => s.toLowerCase().replace(/[^a-z]/g, '');
function uniqEmail(local, domain) { let l = local, n = 1; while (usedEmails.has(l + '@' + domain)) l = local + (++n); usedEmails.add(l + '@' + domain); return l + '@' + domain; }

function makeEmp(nat, o = {}) {
  const N = NAT[nat];
  const g = o.gender ?? (chance(o.femaleRate ?? 0.25) ? 'F' : 'M');
  let ar, en;
  if (N.kind === 'saudi') {
    const f = pick(g === 'M' ? SA_M : SA_F), fa = pick(SA_FATHER), gf = pick(SA_FATHER), fam = pick(SA_FAM);
    ar = [f[0], fa[0], gf[0], fam[0]]; en = [f[1], fa[1], gf[1], fam[1]];
  } else if (N.kind === 'arab') {
    const f = pick(g === 'M' ? ARAB_M : ARAB_F), fa = pick(ARAB_M), fam = pick(ARAB_FAM[N.fam]);
    ar = [f[0], fa[0], fam[0]]; en = [f[1], fa[1], fam[1]];
  } else {
    const S = OTHER_NAMES[nat];
    const f = pick(g === 'M' ? S.m : S.f), last = pick(S.last);
    const parts = [f]; if (S.mid.length && chance(0.4)) parts.push(pick(S.mid)); parts.push(last);
    ar = parts.map((p) => p[0]); en = parts.map((p) => p[1]);
  }
  const deptKey = o.dept ?? pick(o.depts ?? Object.keys(DEPT));
  const D = DEPT[deptKey];
  const title = pick(D.titles);
  const dob = randDate(o.dobFrom ?? 1970, o.dobTo ?? 2001);
  const joinTo = o.joinTo ?? 2025;
  const join = randDate(Math.min(joinTo, Math.max(o.joinFrom ?? 2009, ymd(dob)[0] + 20)), joinTo); // hired at 19+
  const mob = '5' + pick(['0', '3', '4', '5', '6', '8', '9']) + Array.from({ length: 7 }, () => ri(0, 9)).join('');
  const bank = weighted({ '80': 5, '10': 3, '20': 2, '45': 1, '05': 2, '15': 1, '55': 1, '30': 1 });
  const salMin = o.salMin ?? 4000, salMax = o.salMax ?? 28000;
  const e = {
    nat, g, ar, en,
    id: makeId(nat === 'SA' ? 1 : 2),
    dob,
    join,
    salStart: firstOfNextMonth(join),
    idIss: randDate(2017, 2024),
    idExp: nat === 'SA' ? randDate(2027, 2034) : randDate(2027, 2028), // all still valid on 2026-10-22
    mar: weighted({ M: 55, S: 35, D: 7, W: 3 }),
    status: o.status ?? weighted({ A: 80, I: 12, N: 8 }),
    dept: deptKey, titleAr: title[0], titleEn: title[1],
    salary: o.salaryStep ? ri(salMin / o.salaryStep, salMax / o.salaryStep) * o.salaryStep : ri(salMin / 250, salMax / 250) * 250,
    mob, bank, iban: makeIban(bank),
    loc: o.loc ?? weighted({ RUH: 6, JED: 3, DMM: 2 }),
    edu: weighted({ HS: 2, BA: 6, MA: 2, PHD: 0.4 }),
    empType: chance(0.15) ? 'PT' : 'FT',
    contract: nat === 'SA' ? (chance(0.75) ? 'U' : 'F') : (chance(0.85) ? 'F' : 'U'),
    att: chance(0.2) ? 'E' : 'D',
    major: pick(MAJORS),
  };
  if (e.mar === 'S' && e.dob < '1980-01-01') e.mar = 'M';
  const famAr = ar[ar.length - 1];
  const relPool = g === 'M' ? RELS_M : RELS_F;
  e.emRel = e.mar === 'M' ? (g === 'M' ? 'زوجة' : 'زوج') : pick(relPool.filter((r) => !r.startsWith('زو')));
  const relFirst = ['أم', 'أخت', 'زوجة'].includes(e.emRel) ? pick(N.kind === 'saudi' ? SA_F : ARAB_F)[0] : pick(N.kind === 'saudi' ? SA_FATHER : ARAB_M)[0];
  e.emName = N.kind === 'other' ? (e.emRel === 'زوجة' ? pick(OTHER_NAMES[nat].f)[0] : pick(OTHER_NAMES[nat].m)[0]) + ' ' + famAr : relFirst + ' ' + famAr;
  e.emMob = '5' + pick(['0', '4', '5', '6']) + Array.from({ length: 7 }, () => ri(0, 9)).join('');
  e.homePhone = '01' + ri(1, 4) + ri(1000000, 9999999);
  e.address = `${pick(DISTRICTS[e.loc])}، ${pick(STREETS)}، مبنى ${ri(1000, 9999)}`;
  const fn = ascii(en[0].split(' ').join('')), ln = ascii(en[en.length - 1]);
  e.wemail = uniqEmail(`${fn}.${ln}`, o.domain ?? 'company.com.sa');
  e.pemail = uniqEmail(`${fn}${ln}${ri(1, 99)}`, pick(['gmail.com', 'hotmail.com', 'outlook.com', 'yahoo.com']));
  e.fullAr = ar.join(' ');
  e.fullEn = en.join(' ');
  return e;
}
function makeEmps(n, natWeights, o = {}) { const out = []; for (let i = 0; i < n; i++) out.push(makeEmp(weighted(natWeights), o)); return out; }

// ---------------------------------------------------------------- truth
function splitFull(full, lang) {
  const t = full.trim().split(/\s+/);
  const compound = t.length > 2 && (lang === 'ar' ? t[0] === 'عبد' : t[0].toLowerCase() === 'abdul');
  const k = compound ? 2 : 1;
  return [t.slice(0, k).join(' '), t.slice(k).join(' ')];
}
function truthVal(e, f) {
  switch (f) {
    case 'national_id_number': return e.id;
    case 'first_name_ar': return e.ar[0];
    case 'last_name_ar': return e.ar[e.ar.length - 1];
    case 'first_name': return e.en[0];
    case 'last_name': return e.en[e.en.length - 1];
    case 'gender': return e.g === 'M' ? 'ذكر' : 'أنثى';
    case 'date_of_birth': return e.dob;
    case 'join_date': return e.join;
    case 'national_id_expiry_date': return e.idExp;
    case 'national_id_issued_date': return e.idIss;
    case 'salary_start_date': return e.salStart;
    case 'nationality': return NAT[e.nat].ar;
    case 'residence_country': return 'السعودية';
    case 'marital_status': return MAR[e.mar];
    case 'employment_status': return STAT[e.status];
    case 'department': return DEPT[e.dept].ar;
    case 'location': return LOC[e.loc];
    case 'work_phone_number': return e.mob;
    case 'salary_basic_salary_amount': return e.salary;
    case 'iban': return e.iban;
    case 'swift_code': return BANKS[e.bank].bic;
    case 'email': return e.pemail;
    case 'work_email': return e.wemail;
    case 'employee_no': return e.empNo;
    case 'employment_type': return EMPT[e.empType];
    case 'contract_type': return CONT[e.contract];
    case 'attendance_setting': return ATT[e.att];
    case 'education_level': return EDU[e.edu].ar;
    case 'education_description': return e.major;
    case 'bank_account_currency': return 'ريال سعودي';
    case 'salary_pay_frequency': return 'Monthly';
    case 'salary_basic_salary_kind': return 'Fixed';
    case 'emergency_contact_name': return e.emName;
    case 'emergency_contact_relationship': return e.emRel;
    case 'address': return e.address;
    default: throw new Error('no truth for ' + f);
  }
}

// cols: [{h, f?, v(e,i), t?(e), check?:false, numFmt?, width?}]  ;  h===null -> unnamed spacer column
function writeTable(ws, cols, emps, { headerRow, colOffset = 0, blankAfter = [], headerStyle = true }) {
  const expectMap = {}, fields = [];
  cols.forEach((c, j) => {
    const cell = ws.getCell(headerRow, colOffset + j + 1);
    if (c.h != null) cell.value = c.h;
    if (headerStyle && c.h != null) { cell.font = { bold: true }; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } }; cell.alignment = { wrapText: true, vertical: 'middle', horizontal: 'center' }; }
    if (c.width) ws.getColumn(colOffset + j + 1).width = c.width;
    if (!c.f) return;
    expectMap[c.f] = colOffset + j;
    if (c.check === false) return;
    if (c.f === 'full_name_ar') fields.push('first_name_ar', 'last_name_ar');
    else if (c.f === 'full_name_en') fields.push('first_name', 'last_name');
    else fields.push(c.f);
  });
  let r = headerRow + 1;
  const records = [];
  emps.forEach((e, i) => {
    cols.forEach((c, j) => {
      if (!c.v) return;
      const v = c.v(e, i);
      if (v === undefined || v === null || v === '') return;
      const cell = ws.getCell(r, colOffset + j + 1);
      cell.value = v;
      const nf = typeof c.numFmt === 'function' ? c.numFmt(v) : c.numFmt;
      if (nf) cell.numFmt = nf;
    });
    const rec = { national_id_number: e.id };
    for (const c of cols) {
      if (!c.f || c.check === false) continue;
      if (c.f === 'full_name_ar') { const [a, b] = splitFull(e.fullAr, 'ar'); rec.first_name_ar = a; rec.last_name_ar = b; }
      else if (c.f === 'full_name_en') { const [a, b] = splitFull(e.fullEn, 'en'); rec.first_name = a; rec.last_name = b; }
      else rec[c.f] = c.t ? c.t(e) : truthVal(e, c.f);
    }
    records.push(rec);
    r++;
    if (blankAfter.includes(i)) r++;
  });
  return { expectMap, fields: [...new Set(['national_id_number', ...fields])], records, lastRow: r - 1 };
}
const truth = {};
function register(name, t) { truth[name] = { expectMap: t.expectMap, fields: t.fields, synthetic: {}, records: t.records }; }

// =================================================================== FILE 1
// Arabic-only contractor sheet: title block, header on row 6, Hijri DOB, adjectives, ذ/أ, totals + signature footer.
async function file1() {
  const name = 'b1_riyadh_contracting_ar.xlsx';
  const emps = makeEmps(32, { SA: 20, EG: 3, JO: 1, SD: 2, YE: 2, SY: 1, PK: 2, IN: 1 }, { femaleRate: 0.15, depts: ['HR', 'FIN', 'ENG', 'OPS', 'PRC', 'ADM', 'IT'] });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('كشف الموظفين', { views: [{ rightToLeft: true }] });
  ws.mergeCells('A1:Q1'); ws.getCell('A1').value = 'مؤسسة الرياض للمقاولات العامة'; ws.getCell('A1').font = { bold: true, size: 16 }; ws.getCell('A1').alignment = { horizontal: 'center' };
  ws.mergeCells('A2:Q2'); ws.getCell('A2').value = 'سجل تجاري رقم 1010458821 — الرياض';
  ws.mergeCells('A3:Q3'); ws.getCell('A3').value = 'كشف بيانات الموظفين عن شهر سبتمبر 2026م';
  ws.getCell('A4').value = 'ملاحظة: تاريخ الميلاد بالتقويم الهجري، وتاريخ الدخول بالميلادي (يوم/شهر/سنة)';
  const arVariant = (k) => pick([DEPT[k].ar, DEPT[k].arv, DEPT[k].arv]);
  const cols = [
    { h: 'م', v: (e, i) => i + 1, width: 5 },
    { h: 'اسم الموظف رباعي', f: 'full_name_ar', v: (e) => e.fullAr, width: 32 },
    { h: 'رقم الإثبات', f: 'national_id_number', v: (e) => Number(e.id), width: 14 },
    { h: 'جنسية الموظف', f: 'nationality', v: (e) => (e.g === 'F' ? NAT[e.nat].f : NAT[e.nat].m) },
    { h: 'الجنس', f: 'gender', v: (e) => (e.g === 'M' ? 'ذ' : 'أ') },
    { h: 'تاريخ الميلاد (هـ)', f: 'date_of_birth', v: (e) => { const h = hijri(e.dob); return chance(0.75) ? `${pad2(h.d)}/${pad2(h.m)}/${h.y}هـ` : `${h.y}/${pad2(h.m)}/${pad2(h.d)} هـ`; }, width: 16 },
    { h: 'الحالة الاجتماعية', f: 'marital_status', v: (e) => MAR_AR[e.g][e.mar] },
    { h: 'تاريخ الدخول للشركة', f: 'join_date', v: (e) => dmy(e.join), width: 14 },
    { h: 'المهنة', f: 'job_title', v: (e) => e.titleAr, t: (e) => e.titleAr, width: 20 },
    { h: 'الإدارة', f: 'department', v: (e) => arVariant(e.dept), width: 18 },
    { h: 'الأجر الأساسي', f: 'salary_basic_salary_amount', v: (e) => (chance(0.5) ? `${thousands(e.salary)} ريال` : e.salary), numFmt: (v) => (typeof v === 'number' ? '#,##0' : undefined) },
    { h: 'بدل السكن', v: (e) => Math.round(e.salary * 0.25), numFmt: '#,##0' },
    { h: 'بدل النقل', v: (e) => Math.round(e.salary * 0.1), numFmt: '#,##0' },
    { h: 'إجمالي الراتب', v: (e) => e.salary + Math.round(e.salary * 0.25) + Math.round(e.salary * 0.1), numFmt: '#,##0' },
    { h: 'رقم الجوال', f: 'work_phone_number', v: (e) => phone(e.mob, weighted({ '05sp': 5, '05dash': 2, '05': 3, plus: 1 })), width: 15 },
    { h: 'الحالة', f: 'employment_status', v: (e) => ({ A: 'على رأس العمل', I: pick(['مستقيل', 'منتهي الخدمة']), N: 'تحت الإشعار' })[e.status] },
    { h: 'ملاحظات', v: (e) => (e.status === 'I' ? 'تم إخلاء الطرف' : chance(0.15) ? pick(['نقل كفالة', 'إجازة سنوية حتى نهاية الشهر', 'يحتاج تجديد إقامة', 'تحت التجربة']) : undefined), width: 22 },
  ];
  const t = writeTable(ws, cols, emps, { headerRow: 6 });
  const tr = t.lastRow + 2;
  ws.getCell(tr, 2).value = 'الإجمالي'; ws.getCell(tr, 2).font = { bold: true };
  for (const [j, fn] of [[12, (e) => Math.round(e.salary * 0.25)], [13, (e) => Math.round(e.salary * 0.1)], [14, (e) => e.salary + Math.round(e.salary * 0.25) + Math.round(e.salary * 0.1)], [11, (e) => e.salary]]) {
    ws.getCell(tr, j).value = emps.reduce((s, e) => s + fn(e), 0); ws.getCell(tr, j).numFmt = '#,##0'; ws.getCell(tr, j).font = { bold: true };
  }
  ws.getCell(tr + 3, 2).value = 'إعداد: شؤون الموظفين'; ws.getCell(tr + 3, 9).value = 'اعتماد: المدير العام';
  await wb.xlsx.writeFile(path.join(OUT, name));
  register(name, t);
  return [name, 'Arabic-only; title block + note, header on row 6; اسم رباعي; ID as number; nationality adjectives (m/f); ذ/أ; Hijri DOB text with هـ in two orders; dd/mm/yyyy join text; dept spelling variants (ه/ة, شئون); salary mixed number / "5,500 ريال"; allowances+gross; status words; notes; totals row + signature footer.'];
}

// =================================================================== FILE 2
// English HR master, pivot "Summary" sheet first, data on 3rd-from-left "Staff List"; US dates; serial expiry; dept codes.
async function file2() {
  const name = 'b2_gulftech_hr_master.xlsx';
  const emps = makeEmps(35, { SA: 14, EG: 4, IN: 5, PH: 3, JO: 2, LB: 2, PK: 2, GB: 1, US: 1, SY: 1 }, { femaleRate: 0.35, domain: 'gulftech.sa' });
  emps.forEach((e, i) => (e.empNo = 'GT-' + String(101 + i * 3).padStart(4, '0')));
  const wb = new ExcelJS.Workbook();
  const sum = wb.addWorksheet('Summary');
  const ws = wb.addWorksheet('Staff List', { views: [{ state: 'frozen', ySplit: 1 }] });
  const lists = wb.addWorksheet('Lists');
  const cols = [
    { h: 'Emp Code', f: 'employee_no', v: (e) => e.empNo },
    { h: 'First Name', f: 'first_name', v: (e) => e.en[0] },
    { h: 'Last Name', f: 'last_name', v: (e) => e.en[e.en.length - 1] },
    { h: 'Gender', f: 'gender', v: (e) => e.g },
    { h: 'D.O.B', f: 'date_of_birth', v: (e) => mdy(e.dob) },
    { h: 'Nationality', f: 'nationality', v: (e) => (e.g === 'F' && NAT[e.nat].enF) || NAT[e.nat].en },
    { h: 'Marital Status', f: 'marital_status', v: (e) => MAR_EN[e.mar] },
    { h: 'ID / Iqama #', f: 'national_id_number', v: (e) => Number(e.id), width: 14 },
    { h: 'ID Expiry', f: 'national_id_expiry_date', v: (e) => serial(e.idExp) },
    { h: 'Passport No', v: (e) => (e.nat === 'SA' ? undefined : String.fromCharCode(65 + ri(0, 25)) + ri(1000000, 9999999)) },
    { h: 'GOSI No', v: () => String(ri(100000000, 999999999)) },
    { h: 'Mob', f: 'work_phone_number', v: (e) => phone(e.mob, chance(0.8) ? 'plus' : 'plusnosp'), width: 16 },
    { h: 'Work E-mail', f: 'work_email', v: (e) => e.wemail, width: 28 },
    { h: 'Personal E-mail', f: 'email', v: (e) => e.pemail, width: 28 },
    { h: 'Position', f: 'job_title', v: (e) => e.titleEn, t: (e) => e.titleEn, width: 22 },
    { h: 'Dept', f: 'department', v: (e) => DEPT[e.dept].code },
    { h: 'Location', f: 'location', v: (e) => ({ RUH: 'Riyadh HQ', JED: 'Jeddah', DMM: 'Dammam' })[e.loc] },
    { h: 'Hire Date', f: 'join_date', v: (e) => jsDate(e.join), numFmt: 'dd-mmm-yyyy', width: 13 },
    { h: 'Contract', f: 'contract_type', v: (e) => (e.contract === 'F' ? 'Fixed Term' : 'Indefinite') },
    { h: 'Basic Salary', f: 'salary_basic_salary_amount', v: (e) => e.salary, numFmt: '#,##0.00' },
    { h: 'Housing', v: (e) => Math.round(e.salary * 0.25), numFmt: '#,##0.00' },
    { h: 'Transport', v: () => 500, numFmt: '#,##0.00' },
    { h: 'Net Salary', v: (e) => Math.round(e.salary * 1.25 + 500 - (e.nat === 'SA' ? e.salary * 0.0975 : 0)), numFmt: '#,##0.00' },
    { h: 'Status', f: 'employment_status', v: (e) => ({ A: 'Active', I: pick(['Terminated', 'Resigned']), N: 'Notice Period' })[e.status] },
    { h: 'Education', f: 'education_level', v: (e) => EDU[e.edu].en },
  ];
  const t = writeTable(ws, cols, emps, { headerRow: 1 });
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: cols.length } };
  // Summary (pivot-like)
  sum.getCell('A1').value = 'GulfTech Solutions — Headcount Summary (Sep 2026)'; sum.getCell('A1').font = { bold: true, size: 14 };
  ['Dept', 'Headcount', 'Total Basic', 'Avg Basic'].forEach((h, j) => { const c = sum.getCell(3, j + 1); c.value = h; c.font = { bold: true }; });
  let r = 4;
  for (const k of Object.keys(DEPT)) {
    const g = emps.filter((e) => e.dept === k); if (!g.length) continue;
    const tot = g.reduce((s, e) => s + e.salary, 0);
    sum.getCell(r, 1).value = DEPT[k].code;
    sum.getCell(r, 2).value = { formula: `COUNTIF('Staff List'!P:P,A${r})`, result: g.length };
    sum.getCell(r, 3).value = { formula: `SUMIF('Staff List'!P:P,A${r},'Staff List'!T:T)`, result: tot };
    sum.getCell(r, 4).value = Math.round(tot / g.length); r++;
  }
  sum.getCell(r, 1).value = 'Grand Total'; sum.getCell(r, 2).value = emps.length; sum.getCell(r, 3).value = emps.reduce((s, e) => s + e.salary, 0);
  // Lists sheet: code legend
  lists.getCell('A1').value = 'Code'; lists.getCell('B1').value = 'Department'; lists.getCell('D1').value = 'Locations';
  Object.values(DEPT).forEach((d, i) => { lists.getCell(i + 2, 1).value = d.code; lists.getCell(i + 2, 2).value = d.en; });
  ['Riyadh HQ', 'Jeddah', 'Dammam'].forEach((l, i) => (lists.getCell(i + 2, 4).value = l));
  await wb.xlsx.writeFile(path.join(OUT, name));
  register(name, t);
  return [name, "English-only; 'Summary' pivot sheet first, employees on 2nd sheet 'Staff List' (+ 'Lists' legend sheet); US MM/DD/YYYY DOB text; ID expiry as raw Excel serial numbers (no date format); M/F; nationality adjectives incl. Filipina/British/American; dept codes (SLS, PROC, ADMIN...); +966 phones; separate Work/Personal e-mail; Fixed Term/Indefinite; passport/GOSI/housing/net-salary noise."];
}

// =================================================================== FILE 3
// Payroll-system export: UTF-8 BOM + TAB delimited CSV, shuffled columns, yyyymmdd dates, ISO country codes, lowercase IBAN.
async function file3() {
  const name = 'b3_payroll_export.csv';
  const emps = makeEmps(30, { SA: 12, IN: 4, PK: 4, EG: 3, PH: 3, BD: 2, JO: 2 }, { femaleRate: 0.2 });
  emps.forEach((e, i) => (e.empNo = String(10230 + i * 7)));
  const cols = [
    { h: 'EmpNo', f: 'employee_no', v: (e) => e.empNo },
    { h: 'Basic', f: 'salary_basic_salary_amount', v: (e) => e.salary.toFixed(2) },
    { h: 'IBAN', f: 'iban', v: (e) => e.iban.toLowerCase() },
    { h: 'Name_EN', f: 'full_name_en', v: (e) => e.fullEn },
    { h: 'GOSI_Ded', v: (e) => (e.nat === 'SA' ? (e.salary * 0.0975).toFixed(2) : '0.00') },
    { h: 'Gender', f: 'gender', v: (e) => (e.g === 'M' ? 'Male' : 'Female') },
    { h: 'Nat', f: 'nationality', v: (e) => NAT[e.nat].iso },
    { h: 'DOB', f: 'date_of_birth', v: (e) => compact(e.dob) },
    { h: 'Name_AR', f: 'full_name_ar', v: (e) => e.fullAr },
    { h: 'NID', f: 'national_id_number', v: (e) => e.id },
    { h: 'MobileNo', f: 'work_phone_number', v: (e) => phone(e.mob, '966') },
    { h: 'DateOfJoining', f: 'join_date', v: (e) => compact(e.join) },
    { h: 'Dept', f: 'department', v: (e) => DEPT[e.dept].en },
    { h: 'NetPay', v: (e) => (e.salary * 1.35 - (e.nat === 'SA' ? e.salary * 0.0975 : 0)).toFixed(2) },
    { h: 'EmpStatus', f: 'employment_status', v: (e) => ({ A: 'ACTIVE', I: 'INACTIVE', N: 'NOTICE' })[e.status] },
    { h: 'PayGroup', v: () => 'MONTHLY-SAR' },
  ];
  // build truth via a throwaway worksheet so logic stays identical
  const tmp = new ExcelJS.Workbook().addWorksheet('x');
  const t = writeTable(tmp, cols, emps, { headerRow: 1, headerStyle: false });
  const lines = [cols.map((c) => c.h).join('\t')];
  for (let r = 2; r <= t.lastRow; r++) lines.push(cols.map((c, j) => { const v = tmp.getCell(r, j + 1).value; return v == null ? '' : String(v); }).join('\t'));
  fs.writeFileSync(path.join(OUT, name), '﻿' + lines.join('\r\n') + '\r\n', 'utf8');
  register(name, t);
  return [name, 'Payroll export: UTF-8 with BOM, TAB-delimited despite .csv; system column names (EmpNo, NID, Nat, DateOfJoining); shuffled order; both full Arabic and full English name columns; yyyymmdd dates; ISO-2 nationality codes; 9665xxxxxxxx mobiles; lowercase IBAN; Male/Female; ACTIVE/INACTIVE/NOTICE; GOSI/NetPay/PayGroup noise.'];
}

// =================================================================== FILE 4
// Bilingual Jeddah branch: two-level merged header (rows 2-3) with line breaks, Arabic-Indic digits, Hijri expiry, emergency contact.
async function file4() {
  const name = 'b4_jeddah_branch_bilingual.xlsx';
  const emps = makeEmps(28, { SA: 13, EG: 3, PH: 3, IN: 2, JO: 2, SY: 2, LB: 1, YE: 1, SD: 1 }, { femaleRate: 0.4, loc: 'JED', depts: ['SAL', 'CS', 'OPS', 'FIN', 'HR', 'MKT'] });
  emps.forEach((e, i) => (e.empNo = 'JED-' + String(12 + i).padStart(3, '0')));
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('موظفي فرع جدة', { views: [{ rightToLeft: true, state: 'frozen', ySplit: 3 }] });
  ws.mergeCells('A1:Y1'); ws.getCell('A1').value = 'فرع جدة - بيانات الموظفين  |  Jeddah Branch - Employee Data'; ws.getCell('A1').font = { bold: true, size: 14 };
  const groups = [['A2:E2', 'الاسم / Name'], ['F2:I2', 'البيانات الشخصية / Personal'], ['J2:L2', 'الهوية / ID'], ['M2:O2', 'التواصل / Contact'], ['P2:V2', 'بيانات العمل / Job'], ['W2:Y2', 'الطوارئ / Emergency']];
  for (const [rg, txt] of groups) { ws.mergeCells(rg); const c = ws.getCell(rg.split(':')[0]); c.value = txt; c.font = { bold: true }; c.alignment = { horizontal: 'center' }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } }; }
  const natCountry = (e) => (e.nat === 'SA' ? pick(['المملكة العربية السعودية', 'السعودية']) : NAT[e.nat].country);
  const sal = (n) => (chance(0.5) ? toArDigits(n.toLocaleString('en-US')).replace(/,/g, '٬') : toArDigits(n));
  const cols = [
    { h: 'الرقم الوظيفي\nStaff No', f: 'employee_no', v: (e) => e.empNo },
    { h: 'الاسم الأول\nFirst Name (AR)', f: 'first_name_ar', v: (e) => e.ar[0] },
    { h: 'اسم العائلة\nFamily Name (AR)', f: 'last_name_ar', v: (e) => e.ar[e.ar.length - 1] },
    { h: 'First Name\n(English)', f: 'first_name', v: (e) => e.en[0] },
    { h: 'Family Name\n(English)', f: 'last_name', v: (e) => e.en[e.en.length - 1] },
    { h: 'الجنس\nGender', f: 'gender', v: (e) => (e.g === 'M' ? 'ذكر' : 'انثى') },
    { h: 'الجنسية\nNationality', f: 'nationality', v: natCountry },
    { h: 'الحالة الاجتماعية\nMarital', f: 'marital_status', v: (e) => MAR_AR[e.g][e.mar] },
    { h: 'تاريخ الميلاد\nDOB', f: 'date_of_birth', v: (e) => toArDigits(dmy(e.dob)) },
    { h: 'رقم الهوية/الإقامة\nID / Iqama', f: 'national_id_number', v: (e) => toArDigits(e.id), width: 14 },
    { h: 'تاريخ الإصدار\nIssue Date', f: 'national_id_issued_date', v: (e) => jsDate(e.idIss), numFmt: 'dd/mm/yyyy' },
    { h: 'تاريخ الانتهاء (هجري)\nExpiry (Hijri)', f: 'national_id_expiry_date', v: (e) => { const h = hijri(e.idExp); return toArDigits(`${h.y}/${pad2(h.m)}/${pad2(h.d)}`) + 'هـ'; } },
    { h: 'الجوال\nMobile', f: 'work_phone_number', v: (e) => toArDigits(phone(e.mob, weighted({ '05': 3, '00966nosp': 1, '05sp': 1 }))) },
    { h: 'البريد الإلكتروني (العمل)\nWork Email', f: 'work_email', v: (e) => e.wemail },
    { h: 'البريد الشخصي\nPersonal Email', f: 'email', v: (e) => e.pemail },
    { h: 'المسمى الوظيفي\nJob Title', f: 'job_title', v: (e) => e.titleAr, t: (e) => e.titleAr },
    { h: 'القسم\nDept.', f: 'department', v: (e) => pick([DEPT[e.dept].ar, DEPT[e.dept].arv]) },
    { h: 'تاريخ المباشرة\nStart Date', f: 'join_date', v: (e) => jsDate(e.join), numFmt: 'dd/mm/yyyy' },
    { h: 'نوع الدوام\nWork Type', f: 'employment_type', v: (e) => (e.empType === 'FT' ? 'كامل' : 'جزئي') },
    { h: 'المؤهل\nQualification', f: 'education_level', v: (e) => EDU[e.edu].raw },
    { h: 'التخصص\nMajor', f: 'education_description', v: (e) => e.major },
    { h: 'الراتب الأساسي\nBasic', f: 'salary_basic_salary_amount', v: (e) => sal(e.salary) },
    { h: 'اسم جهة الطوارئ\nEmergency Contact', f: 'emergency_contact_name', v: (e) => e.emName },
    { h: 'صلة القرابة\nRelation', f: 'emergency_contact_relationship', v: (e) => e.emRel },
    { h: 'جوال الطوارئ\nEmergency Mobile', f: 'emergency_contact_mobile_phone_number', check: false, v: (e) => toArDigits(phone(e.emMob, '05')) },
  ];
  const t = writeTable(ws, cols, emps, { headerRow: 3 });
  ws.getRow(3).height = 32;
  await wb.xlsx.writeFile(path.join(OUT, name));
  register(name, t);
  return [name, 'Bilingual, RTL sheet; title row + merged group-header row 2, real headers on row 3 with "Arabic\\nEnglish" line breaks; separate AR/EN first+family names; Arabic-Indic digits for DOB text, ID, mobile, salary (with ٬ separator); Hijri expiry with Arabic digits + هـ; country names as nationality (المملكة العربية السعودية, الفلبين without shadda); انثى without hamza; feminine marital forms; real Date cells for issue/join; كامل/جزئي; qualification words; emergency contact block.'];
}

// =================================================================== FILE 5
// Dammam labour manpower list: data starts in column B, header row 4, unnamed spacer column, Border No 10-digit decoy,
// Hijri Iqama expiry w/ and w/o هـ, dd.mm.yyyy DOB, serial joining dates, mobile numbers without leading zero.
async function file5() {
  const name = 'b5_dammam_manpower_iqama.xlsx';
  const emps = makeEmps(38, { PK: 9, IN: 9, BD: 6, PH: 5, NP: 4, EG: 2, SD: 1, LK: 1, YE: 1 }, { femaleRate: 0.05, loc: 'DMM', depts: ['ENG', 'OPS'], salMin: 1500, salMax: 6000, salaryStep: 50, dobFrom: 1972, dobTo: 2001, joinFrom: 2015 });
  const jobs = ['Electrician', 'Steel Fixer', 'Driver', 'Site Engineer', 'Foreman', 'Safety Officer', 'Mason', 'Plumber', 'Welder', 'Storekeeper'];
  emps.forEach((e, i) => { e.empNo = 'DMM-' + (3301 + i); e.job = pick(jobs); });
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1');
  ws.getCell('B1').value = 'DAMMAM SITE - MANPOWER LIST'; ws.getCell('B1').font = { bold: true, size: 14 };
  ws.getCell('B2').value = 'Updated: 01-Oct-2026   (Iqama expiry in Hijri)';
  const cols = [
    { h: 'S.No', v: (e, i) => i + 1 },
    { h: 'Badge No', f: 'employee_no', v: (e) => e.empNo },
    { h: 'Employee Name (as per Iqama)', f: 'full_name_en', v: (e) => e.fullEn, width: 30 },
    { h: 'Iqama#', f: 'national_id_number', v: (e) => Number(e.id), width: 13 },
    { h: 'Iqama Expiry (H)', f: 'national_id_expiry_date', v: (e) => { const h = hijri(e.idExp); return chance(0.5) ? `${h.y}/${pad2(h.m)}/${pad2(h.d)}` : `${pad2(h.d)}/${pad2(h.m)}/${h.y} هـ`; } },
    { h: 'Nationality', f: 'nationality', v: (e) => (e.g === 'F' && NAT[e.nat].enF) || NAT[e.nat].en },
    { h: 'Profession (Iqama)', f: 'job_title', v: (e) => e.job, t: (e) => e.job },
    { h: 'Passport No', v: () => String.fromCharCode(65 + ri(0, 25)) + String.fromCharCode(65 + ri(0, 25)) + ri(1000000, 9999999) },
    { h: 'Passport Expiry', v: () => dmy(randDate(2027, 2033), '.') },
    { h: 'Border No.', v: () => Number('3' + ri(100000000, 999999999)), width: 13 },
    { h: null },
    { h: 'D.O.B', f: 'date_of_birth', v: (e) => dmy(e.dob, '.') },
    { h: 'Gender', f: 'gender', v: (e) => e.g },
    { h: 'Joining Date', f: 'join_date', v: (e) => serial(e.join) },
    { h: 'Mobile', f: 'work_phone_number', v: (e) => phone(e.mob, 'num') },
    { h: 'Basic', f: 'salary_basic_salary_amount', v: (e) => e.salary },
    { h: 'Food Allow.', v: () => 300 },
    { h: 'OT Rate/hr', v: (e) => Math.round((e.salary / 240) * 1.5 * 100) / 100 },
    { h: 'Status', f: 'employment_status', v: (e) => ({ A: 'على رأس العمل', I: 'خروج نهائي', N: 'تحت الإشعار' })[e.status] },
    { h: 'Remarks', v: (e) => (e.status === 'I' ? 'Final exit issued' : chance(0.2) ? pick(['Vacation', 'Iqama renewal pending', 'Transferred from Jubail']) : undefined) },
  ];
  const t = writeTable(ws, cols, emps, { headerRow: 4, colOffset: 1 });
  await wb.xlsx.writeFile(path.join(OUT, name));
  register(name, t);
  return [name, "Expat labour list: column A empty (table starts at B), header row 4; full English names incl. 'Abdul X' compounds and 2-token 'Dela Cruz'; Iqama# as number; Hijri expiry mixing yyyy/mm/dd (no هـ) and dd/mm/yyyy هـ; 10-digit 'Border No.' decoy starting with 3; unnamed blank spacer column; DOB dd.mm.yyyy; joining dates as raw Excel serials; mobile as number w/o leading 0; Arabic status words in English file."];
}

// =================================================================== FILE 6
// Hand-made agency sheet: mixed languages/formats within each column, blank rows, header on row 2 with odd spacing.
async function file6() {
  const name = 'b6_agency_handmade_mixed.xlsx';
  const emps = makeEmps(26, { SA: 12, EG: 4, LB: 3, JO: 3, SY: 2, IN: 1, GB: 1 }, { femaleRate: 0.5, depts: ['MKT', 'SAL', 'IT', 'HR', 'FIN', 'ADM', 'CS'] });
  emps.forEach((e, i) => (e.manager = i === 0 ? 'Khalid Al-Rasheed' : emps[i < 6 ? 0 : ri(1, 5)].fullEn));
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('staff 2026');
  ws.getCell('A1').value = 'Employees list (draft) - please do not share / قائمة الموظفين - مسودة';
  const mixDate = (s) => {
    const [, , d] = ymd(s);
    const style = weighted({ date: 4, iso: 2, dmy: 3, mon: 1 });
    if (style === 'iso') return s;
    if (style === 'mon') return dMonYYYY(s);
    if (style === 'dmy' && d > 12) return dmy(s);
    return jsDate(s);
  };
  const gen = (e) => (e.g === 'M' ? pick(['Male', 'M', 'ذكر', 'male']) : pick(['Female', 'female', 'F', 'أنثى']));
  const natMix = (e) => { const N = NAT[e.nat]; if (e.nat === 'SA') return pick(['Saudi', 'سعودي', 'KSA', e.g === 'F' ? 'سعودية' : 'Saudi Arabian']); return pick([N.en, e.g === 'F' ? N.f : N.m]); };
  const idMix = (e) => { const s = weighted({ num: 3, str: 2, ar: 1, sp: 1 }); return s === 'num' ? Number(e.id) : s === 'ar' ? toArDigits(e.id) : s === 'sp' ? ` ${e.id} ` : e.id; };
  const team = (e) => ({ MKT: ['Marketing', 'Mktg', 'التسويق'], SAL: ['Sales', 'المبيعات'], IT: ['IT', 'I.T.'], HR: ['HR', 'Human Resources', 'الموارد البشرية'], FIN: ['Finance', 'Accounts & Finance'], ADM: ['Admin', 'Administration'], CS: ['Customer Care', 'Customer Service'] })[e.dept];
  const salMix = (e) => { const s = weighted({ n: 3, c: 2, sar: 2, rs: 2 }); return s === 'n' ? e.salary : s === 'c' ? thousands(e.salary) : s === 'sar' ? `SAR ${thousands(e.salary)}` : `${e.salary} ر.س`; };
  const statMix = (e) => ({ A: ['Active', 'active ', 'نشط', 'Active'], I: ['Left', 'Resigned'], N: ['On notice', 'Serving notice'] })[e.status];
  const cols = [
    { h: '#', v: (e, i) => i + 1 },
    { h: 'Name', f: 'full_name_en', v: (e) => e.fullEn, width: 30 },
    { h: 'الاسم', f: 'full_name_ar', v: (e) => e.fullAr, width: 30 },
    { h: 'Gender / الجنس', f: 'gender', v: gen },
    { h: 'Nationality ', f: 'nationality', v: natMix },
    { h: 'ID No.', f: 'national_id_number', v: idMix, width: 14 },
    { h: 'Birth date', f: 'date_of_birth', v: (e) => mixDate(e.dob), numFmt: (v) => (v instanceof Date ? 'd/m/yyyy' : undefined), width: 13 },
    { h: 'Joined', f: 'join_date', v: (e) => mixDate(e.join), numFmt: (v) => (v instanceof Date ? 'd-mmm-yy' : undefined), width: 13 },
    { h: 'Title', f: 'job_title', v: (e) => e.titleEn, t: (e) => e.titleEn, width: 22 },
    { h: 'Team', f: 'department', v: (e) => pick(team(e)) },
    { h: 'FT/PT', f: 'employment_type', v: (e) => (e.empType === 'FT' ? pick(['Full Time', 'Full-time', 'FT']) : pick(['Part Time', 'Part-time'])) },
    { h: 'Salary (SAR)', f: 'salary_basic_salary_amount', v: salMix },
    { h: 'Phone', f: 'work_phone_number', v: (e) => phone(e.mob, pick(['05', '05sp', 'plusnosp', '00966', '966dash', 'sp9', 'num', 'paren'])), width: 17 },
    { h: 'Email', f: 'email', v: (e) => e.pemail, width: 28 },
    { h: 'Reports to', f: 'manager', v: (e) => e.manager, t: (e) => e.manager, width: 28 },
    { h: 'Status', f: 'employment_status', v: (e) => pick(statMix(e)) },
    { h: 'Notes', v: (e) => (chance(0.25) ? pick(['remote 2 days/week', 'probation ends Dec', 'laptop not returned', 'promoted Jan 2026', 'تم التجديد']) : undefined), width: 22 },
  ];
  const t = writeTable(ws, cols, emps, { headerRow: 2, blankAfter: [7, 16] });
  await wb.xlsx.writeFile(path.join(OUT, name));
  register(name, t);
  return [name, "Hand-typed sheet: note in row 1, header row 2 (trailing space 'Nationality '), 2 blank rows inside the data; every column mixes formats/languages: gender Male/M/ذكر/female/F/أنثى; nationality Saudi/سعودي/KSA/Saudi Arabian/مصرية…; ID as number, text, Arabic-Indic or space-padded; dates mix real dates, ISO, dd/mm/yyyy (day>12 only) and '7 Mar 1991'; Team names/abbrevs AR+EN; salary 12,000 / SAR 9,500 / 8500 ر.س; every phone format; manager names; Active/active /نشط/Left/On notice."];
}

// =================================================================== FILE 7
// WPS bank salary-transfer file: metadata block with company IBAN decoy, header row 8, spaced IBANs, BIC, totals footer.
async function file7() {
  const name = 'b7_wps_salary_transfer.xlsx';
  const emps = makeEmps(30, { SA: 15, EG: 3, IN: 4, PK: 3, PH: 2, SD: 1, JO: 2 }, { femaleRate: 0.25 });
  emps.forEach((e, i) => (e.empNo = 'E-' + (5521 + i)));
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Payroll Sep-2026');
  const companyIban = makeIban('80');
  [['Payroll Transfer File (WPS)'], ['Company Name:', 'Al Noor Trading Est.'], ['CR No:', '2050112233'], ['MOL Establishment ID:', '7-1234567'], ['Salary Month:', '09/2026'], ['Debit Account IBAN:', ibanSpaced(companyIban)], ['Value Date:', '2026-09-27']]
    .forEach((row, i) => row.forEach((v, j) => (ws.getCell(i + 1, j + 1).value = v)));
  ws.getCell('A1').font = { bold: true, size: 14 };
  const cols = [
    { h: 'Sr', v: (e, i) => i + 1 },
    { h: 'Employee Code', f: 'employee_no', v: (e) => e.empNo },
    { h: 'Beneficiary Name', f: 'full_name_en', v: (e) => e.fullEn, width: 32 },
    { h: 'Beneficiary ID Type', v: (e) => (e.nat === 'SA' ? 'NID' : 'IQAMA') },
    { h: 'Beneficiary ID', f: 'national_id_number', v: (e) => e.id, width: 13 },
    { h: 'Bank', f: 'bank_name', v: (e) => BANKS[e.bank].en, t: (e) => BANKS[e.bank].en, width: 20 },
    { h: 'Bank BIC', f: 'swift_code', v: (e) => BANKS[e.bank].bic },
    { h: 'Account IBAN', f: 'iban', v: (e) => ibanSpaced(e.iban), width: 32 },
    { h: 'Basic Salary', f: 'salary_basic_salary_amount', v: (e) => e.salary, numFmt: '#,##0.00' },
    { h: 'Housing Allowance', v: (e) => Math.round(e.salary * 0.25), numFmt: '#,##0.00' },
    { h: 'Other Earnings', v: () => ri(0, 6) * 100, numFmt: '#,##0.00' },
    { h: 'Deductions', v: (e) => (e.nat === 'SA' ? Math.round(e.salary * 0.0975 * 100) / 100 : 0), numFmt: '#,##0.00' },
    { h: 'Net Amount', v: (e, i) => null },
    { h: 'Mobile', f: 'work_phone_number', v: (e) => phone(e.mob, 'plusdash'), width: 17 },
    { h: 'Payment Description', v: () => 'Salary Sep 2026' },
  ];
  const t = writeTable(ws, cols, emps, { headerRow: 8 });
  for (let r = 9; r <= t.lastRow; r++) { ws.getCell(r, 13).value = { formula: `I${r}+J${r}+K${r}-L${r}` }; ws.getCell(r, 13).numFmt = '#,##0.00'; }
  const tr = t.lastRow + 1;
  ws.getCell(tr, 3).value = 'TOTAL';
  for (const col of ['I', 'J', 'K', 'L', 'M']) ws.getCell(`${col}${tr}`).value = { formula: `SUM(${col}9:${col}${t.lastRow})` };
  ws.getCell(tr + 2, 1).value = `Number of records: ${emps.length}`;
  await wb.xlsx.writeFile(path.join(OUT, name));
  register(name, t);
  return [name, 'WPS bank-transfer file: 7-row metadata block (incl. a company Debit Account IBAN decoy), header row 8; Beneficiary Name/ID wording; ID Type NID/IQAMA column; IBAN in 4-char groups; bank name + BIC; earnings/deductions/net formulas; +966-5x-xxx-xxxx; TOTAL row with SUM formulas and record-count footer.'];
}

// =================================================================== FILE 8
// Arabic HR-system export with rare fields: 2-digit-year dd-Mon-yy dates, hidden column, attendance/pay settings, address, bank.
async function file8() {
  const name = 'b8_hr_system_export_ar.xlsx';
  const emps = makeEmps(25, { SA: 15, EG: 2, JO: 2, IN: 2, PH: 2, SY: 1, LB: 1 }, { femaleRate: 0.35, joinFrom: 2012 });
  emps.forEach((e, i) => { e.empNo = String(2045 + i); });
  emps.forEach((e, i) => (e.manager = i === 0 ? 'سعود بن عبدالله الفيصل' : emps[ri(0, Math.min(i - 1, 4))].fullAr));
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Export', { views: [{ rightToLeft: true }] });
  const natAr = (e) => (e.nat === 'SA' ? 'السعودية' : NAT[e.nat].country);
  const cols = [
    { h: 'الرقم الوظيفي', f: 'employee_no', v: (e) => Number(e.empNo), t: (e) => e.empNo },
    { h: 'الاسم الكامل', f: 'full_name_ar', v: (e) => e.fullAr, width: 30 },
    { h: 'الهوية الوطنية / الإقامة', f: 'national_id_number', v: (e) => e.id, width: 14 },
    { h: 'الجنس', f: 'gender', v: (e) => (e.g === 'M' ? 'ذكر' : 'أنثى') },
    { h: 'تاريخ الميلاد', f: 'date_of_birth', v: (e) => dMonYY(e.dob) },
    { h: 'الجنسية', f: 'nationality', v: natAr },
    { h: 'بلد الإقامة', f: 'residence_country', v: () => pick(['المملكة العربية السعودية', 'السعودية']) },
    { h: 'الحالة الاجتماعية', f: 'marital_status', v: (e) => MAR[e.mar] },
    { h: 'العنوان الوطني', f: 'address', v: (e) => e.address, t: (e) => e.address, width: 40 },
    { h: 'هاتف المنزل', f: 'home_phone_number', check: false, v: (e) => e.homePhone },
    { h: 'الجوال', f: 'work_phone_number', v: (e) => phone(e.mob, 'paren'), width: 17 },
    { h: 'البريد', f: 'work_email', v: (e) => e.wemail, width: 28 },
    { h: 'الفرع', f: 'location', v: (e) => ({ RUH: 'الرياض', JED: 'جدة', DMM: 'الدمام' })[e.loc] },
    { h: 'الإدارة', f: 'department', v: (e) => 'إدارة ' + DEPT[e.dept].ar },
    { h: 'المدير المباشر', f: 'manager', v: (e) => e.manager, t: (e) => e.manager, width: 30 },
    { h: 'الوظيفة', f: 'job_title', v: (e) => e.titleAr, t: (e) => e.titleAr },
    { h: 'تاريخ التعيين', f: 'join_date', v: (e) => dMonYY(e.join) },
    { h: 'نوع العقد', f: 'contract_type', v: (e) => (e.contract === 'F' ? 'محدد' : 'غير محدد') },
    { h: 'نظام الدوام', f: 'attendance_setting', v: (e) => (e.att === 'D' ? 'عادي' : 'مسائي') },
    { h: 'العملة', f: 'bank_account_currency', v: () => pick(['SAR', 'ر.س', 'ريال']) },
    { h: 'تاريخ بداية الراتب', f: 'salary_start_date', v: (e) => dMonYY(e.salStart) },
    { h: 'دورية الصرف', f: 'salary_pay_frequency', v: () => 'شهري' },
    { h: 'نوع الراتب', f: 'salary_basic_salary_kind', v: () => 'ثابت' },
    { h: 'الراتب الأساسي', f: 'salary_basic_salary_amount', v: (e) => e.salary, numFmt: '0.00' },
    { h: 'رقم الحساب القديم', v: () => String(ri(100000000, 999999999)) + String(ri(100, 999)) },
    { h: 'الآيبان', f: 'iban', v: (e) => (chance(0.5) ? ibanSpaced(e.iban) : e.iban), width: 32 },
    { h: 'اسم البنك', f: 'bank_name', v: (e) => BANKS[e.bank].ar, t: (e) => BANKS[e.bank].ar },
    { h: 'اسم قريب للطوارئ', f: 'emergency_contact_name', v: (e) => e.emName },
    { h: 'صلة القرابة', f: 'emergency_contact_relationship', v: (e) => e.emRel },
    { h: 'جوال الطوارئ', f: 'emergency_contact_mobile_phone_number', check: false, v: (e) => phone(e.emMob, '05') },
    { h: 'حالة الموظف', f: 'employment_status', v: (e) => ({ A: 'مفعل', I: 'غير مفعل', N: 'فترة اشعار' })[e.status] },
  ];
  const t = writeTable(ws, cols, emps, { headerRow: 1 });
  ws.getColumn(25).hidden = true; // رقم الحساب القديم
  await wb.xlsx.writeFile(path.join(OUT, name));
  register(name, t);
  return [name, "Arabic HR-system export (RTL) with rare fields: dd-Mon-yy dates with 2-digit years (DOB 1970-2001) for DOB/join/salary start; employee no as number; 'إدارة X' dept prefix; branch city names → locations; residence country; national address; home phone; +966 (5x) xxx-xxxx; manager full names; محدد/غير محدد; عادي/مسائي; SAR/ر.س/ريال; شهري; ثابت; hidden old-account column; IBAN spaced/unspaced; Arabic bank names; emergency contact; مفعل without shadda, فترة اشعار."];
}

// =================================================================== run
const notes = [];
for (const f of [file1, file2, file3, file4, file5, file6, file7, file8]) notes.push(await f());

// self-checks
for (const [file, t] of Object.entries(truth)) {
  for (const r of t.records) {
    if (!/^[12]\d{9}$/.test(r.national_id_number) || !luhnOk(r.national_id_number)) throw new Error(`bad id ${file} ${r.national_id_number}`);
    if (r.iban && !ibanOk(r.iban)) throw new Error(`bad iban ${file} ${r.iban}`);
    for (const f of t.fields) if (r[f] === undefined) throw new Error(`missing ${f} in ${file}`);
  }
}
fs.writeFileSync(path.join(OUT, 'truth.json'), JSON.stringify(truth, null, 2), 'utf8');
for (const [n, d] of notes) console.log(`${n.padEnd(34)} ${String(truth[n].records.length).padStart(3)} rows, ${truth[n].fields.length} fields`);
console.log('truth.json written to', OUT);
