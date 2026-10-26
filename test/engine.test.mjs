// اختبار شامل للمسار كامل على ملف العميل التجريبي: قراءة ← مطابقة ← تنظيف ← تعبئة القالب ← التقرير
import fs from 'node:fs';
import os from 'node:os';
import assert from 'node:assert/strict';
import ExcelJS from 'exceljs';
import { parseFollowup } from '../src/engine.js';
import { buildFollowup, followupRows } from '../src/template.js';
import { readClientFile, analyzeWorkbook, buildRecords, processRecords, cleanPhone, cleanEmail, cleanIban, cleanId, cleanAmount, parseDate, makeResolver, splitScripts } from '../src/engine.js';
import { readTemplate, fillTemplate, buildReport } from '../src/template.js';
import { FIELDS } from '../src/fields.js';

const tplBuf = fs.readFileSync(new URL('./template_test.xlsx', import.meta.url));
const tpl = await readTemplate(tplBuf);
console.log('template headers', tpl.headers.length, 'required', tpl.required.size, 'lists', Object.keys(tpl.lists).length);
assert.equal(tpl.headers.length, 46);

const cbuf = fs.readFileSync(new URL('./client_messy_sample.xlsx', import.meta.url));
const ab = cbuf.buffer.slice(cbuf.byteOffset, cbuf.byteOffset + cbuf.byteLength);
const sheets = readClientFile(ab, 'client.xlsx');
const a = analyzeWorkbook(sheets, tpl.lists);
const P = a.prepared;
console.log('sheet', P.name, 'header row', a.hdr.idx + 1, 'data rows', P.data.length, 'skipped', P.skipped.map((s) => `${s.src}:${s.reason}`).join(', '));
assert.equal(a.hdr.idx, 3);
const { mapping } = a;
for (const [k, v] of Object.entries(mapping)) console.log('  ', k.padEnd(30), '←', P.cols.find((c) => c.col === v.col).text, v.score.toFixed(2), v.by);
assert.ok(!Object.values(mapping).some((m) => P.cols.find((c) => c.col === m.col).text === 'م'), 'عمود م التسلسلي لا يُربط');
assert.equal(Object.keys(mapping).length, 24);

const recs = buildRecords(P, mapping, { residence_country: 'السعودية', attendance_setting: 'افتراضي' }, { nameSplit: 'rest' });
assert.equal(recs.length, 13, 'الصف الفاضي وصف الإجمالي يُتجاهلان');
const rows = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings: { dateOrder: 'dmy', addCC: false } });
const bySrc = Object.fromEntries(rows.map((r) => [r.src, r]));
// توقعات محددة على الأخطاء المتعمدة في الملف التجريبي
assert.equal(bySrc[6].cells.email.value, 'sara.q@horizon-demo.sa');
assert.equal(bySrc[6].cells.salary_basic_salary_amount.value, 9800);
assert.equal(bySrc[7].cells.nationality.value, 'السعودية');
assert.equal(bySrc[7].cells.work_phone_number.value, '0505050505');
assert.equal(bySrc[11].cells.date_of_birth.status, 'fixed', 'تاريخ هجري يتحول بدقة (مُصحَّح)');
assert.equal(bySrc[12].cells.national_id_number.status, 'review', 'هوية 9 خانات');
assert.equal(bySrc[13].cells.email.status, 'error', 'بريد بدون نطاق');
assert.equal(bySrc[14].cells.employee_no.status, 'error', 'رقم وظيفي مكرر');
const dupNote = bySrc[14].cells.employee_no.notes.find((n) => n.ref !== undefined);
assert.equal(dupNote.ref, 6, 'ملاحظة التكرار تشير للصف الآخر');
assert.ok(dupNote.msg.includes('سارة'), 'وتذكر اسم الموظف الآخر');
// "تمت المراجعة": الخانة الصفراء تصير سليمة، وترجع صفراء إذا تغيّرت القيمة
const revCell = bySrc[12].cells.national_id_number;
assert.equal(revCell.status, 'review');
const ack = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings: { dateOrder: 'dmy', addCC: false }, reviewed: { 12: { national_id_number: String(revCell.value) } } });
assert.equal(ack.find((r) => r.src === 12).cells.national_id_number.status, 'ok');
assert.equal(ack.find((r) => r.src === 12).cells.national_id_number.acked, true);
const ack2 = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings: {}, reviewed: { 12: { national_id_number: 'قيمة قديمة' } } });
assert.equal(ack2.find((r) => r.src === 12).cells.national_id_number.status, 'review', 'المراجعة تسقط لو القيمة تغيّرت');
assert.equal(bySrc[16].cells.first_name.value, 'Tariq');
// الترجمة الحرفية المتبادلة
assert.equal(bySrc[16].cells.first_name_ar.value, 'طارق');
assert.equal(bySrc[16].cells.last_name_ar.value, 'محمود');
assert.equal(bySrc[16].cells.first_name_ar.status, 'review');
assert.equal(bySrc[13].cells.first_name.value, 'Yousef');
assert.equal(bySrc[13].cells.last_name.value, 'Hassan Ali');
const noTl = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings: { translit: false } });
assert.equal(noTl.find((r) => r.src === 16).cells.first_name_ar.status, 'error', 'بدون ترجمة يبقى خطأ');
const withOv = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings: {}, overrides: { 16: { first_name_ar: 'طارق' } } });
assert.equal(withOv.find((r) => r.src === 16).cells.first_name_ar.status, 'ok', 'التعديل اليدوي يتقدّم على الترجمة');
assert.equal(bySrc[17].cells.salary_basic_salary_amount.status, 'error', 'راتب صفر');
for (const r of rows) {
  const bad = FIELDS.filter((f) => ['error', 'review'].includes(r.cells[f.key].status)).map((f) => `${f.key}[${r.cells[f.key].status[0]}]: ${r.cells[f.key].msg}`);
  console.log(`row ${r.src} ${r.status.toUpperCase()}`, bad.length ? '\n    ' + bad.join('\n    ') : '');
}

// وحدات التنظيف
assert.deepEqual(cleanPhone('0501234567 / 0114567890', { addCC: true }).value, '+966501234567');
assert.equal(cleanPhone('00966 50 123 4567', { addCC: false }).value, '+966501234567');
assert.equal(cleanPhone(501234567, { addCC: false }).status, 'review');
assert.equal(cleanEmail('mailto:Ali@X.com').value, 'ali@x.com');
assert.equal(cleanEmail(' a@b.com; c@d.com ').value, 'a@b.com');
assert.equal(cleanIban('IBAN: sa03 8000 0000 6080 1016 7519').value, 'SA0380000000608010167519');
assert.equal(cleanId("'1081234567").value, '1081234567');
assert.equal(cleanId('1081 234567').value, '1081234567');
assert.equal(cleanAmount('SAR 12,500.00').value, 12500);
assert.equal(cleanAmount('12500').status, 'ok');
assert.equal(parseDate('1990-03-15 00:00:00').iso, '1990-03-15');
assert.equal(parseDate('15-Mar-1990').iso, '1990-03-15');
assert.equal(parseDate('٠٣/٠٥/١٩٨٥', 'dmy').iso, '1985-05-03');
assert.equal(parseDate('1410/02/01هـ').iso, '1989-09-02');
assert.equal(parseDate('20230115').iso, '2023-01-15');
assert.equal(parseDate('3/25/1990', 'unknown').iso, '1990-03-25');
const nat = makeResolver('nationality', tpl.lists.nationality);
for (const [raw, want] of [['أمريكية', 'الولايات المتّحدة'], ['EG', 'مصر'], ['Filipino', 'الفلبّين'], ['سعوديه', 'السعودية'], ['KSA', 'السعودية'], ['Indian', 'الهند']]) assert.equal(nat(raw).value, want, raw);
assert.deepEqual(splitScripts('محمد العتيبي - Mohammed Alotaibi'), { ar: 'محمد العتيبي', en: 'Mohammed Alotaibi' });

// تعبئة القالب: نفس الترويسة، القيم في أماكنها، التواريخ أرقام تسلسلية، الهويات نص
const okRows = rows.filter((r) => r.status !== 'error');
const out = await fillTemplate(tplBuf, tpl, okRows);
const tmp = os.tmpdir();
fs.writeFileSync(`${tmp}/out_upload.xlsx`, out);
fs.writeFileSync(`${tmp}/out_report.xlsx`, await buildReport(rows, { clientFile: 'client.xlsx', templateFile: 'tpl.xlsx', exported: okRows.length }));
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(out);
const ws = wb.getWorksheet('Employees');
assert.equal(ws.getCell(1, 1).value, 'first_name *');
assert.ok(wb.getWorksheet('Reference'), 'شيت القوائم باقٍ');
const col = (key) => tpl.headers.find((h) => h.key === key).col;
assert.equal(ws.getCell(2, col('first_name_ar')).value, okRows[0].cells.first_name_ar.value);
assert.equal(typeof ws.getCell(2, col('national_id_number')).value, 'string', 'الهوية نص');
assert.ok(ws.getCell(2, col('date_of_birth')).value instanceof Date, 'التاريخ تاريخ');
assert.equal(ws.actualRowCount, okRows.length + 1);
assert.ok(ws.dataValidations && Object.keys(ws.dataValidations.model).length > 0, 'قيود القوائم باقية');
// ملف الاستكمال للعميل: يُبنى ← "العميل" يعبّيه ← يُقرأ ويُطبَّق كتعديلات
const fu = await buildFollowup(rows, tpl, { includeReview: true });
fs.writeFileSync(`${tmp}/out_followup.xlsx`, fu);
const fwb = new ExcelJS.Workbook();
await fwb.xlsx.load(fu);
const fws = fwb.getWorksheet('استكمال البيانات');
assert.ok(fws && fwb.getWorksheet('التعليمات'));
assert.equal(fws.getRow(2).hidden, true);
const keyRow = fws.getRow(2).values;
const colOfKey = (k) => keyRow.indexOf(k);
const rowOfSrc = (src) => { for (let i = 3; i <= fws.rowCount; i++) if (fws.getCell(i, colOfKey('__src')).value === String(src)) return i; return -1; };
assert.equal(fws.rowCount - 2, followupRows(rows).length, 'صف لكل موظف يحتاج متابعة');
assert.ok(rowOfSrc(13) > 0 && rowOfSrc(17) > 0);
assert.ok(colOfKey('email') > 0 && colOfKey('salary_basic_salary_amount') > 0 && colOfKey('department') > 0);
const yellow = (r, k) => fws.getCell(r, colOfKey(k)).fill?.fgColor?.argb === 'FFFFF3C4';
assert.ok(yellow(rowOfSrc(13), 'email'), 'البريد الخاطئ أصفر');
assert.ok(fws.getCell(rowOfSrc(11), colOfKey('department')).dataValidation?.type === 'list', 'قائمة منسدلة للقسم');
// خانات التاريخ: تاريخ حقيقي بتنسيق إنجليزي + تحقق "تاريخ"
const dc = fws.getCell(rowOfSrc(13), colOfKey('contract_end_date'));
assert.equal(dc.numFmt, '[$-409]dd-mmm-yyyy;@');
assert.equal(dc.dataValidation?.type, 'date', 'تحقق تاريخ على الخانة المطلوبة');
const expCol = colOfKey('national_id_expiry_date');
if (expCol > 0) for (let i = 3; i <= fws.rowCount; i++) { const c = fws.getCell(i, expCol); if (c.value) assert.ok(c.value instanceof Date && c.numFmt.includes('mmm')); }
// العميل يعبّي (تاريخ مختار كتاريخ حقيقي، وتاريخ مكتوب كنص لم يحوّله إكسل)
fws.getCell(rowOfSrc(13), colOfKey('email')).value = 'yousef.hassan@horizon-demo.sa';
fws.getCell(rowOfSrc(13), colOfKey('contract_end_date')).value = new Date(Date.UTC(2027, 11, 31));
const r16 = rowOfSrc(16);
if (r16 > 0 && expCol > 0) fws.getCell(r16, expCol).value = '15-Mar-2031';
fws.getCell(rowOfSrc(17), colOfKey('salary_basic_salary_amount')).value = 7200;
fws.getCell(rowOfSrc(11), colOfKey('department')).value = 'المالية';
const filled = await fwb.xlsx.writeBuffer();
const back = parseFollowup(readClientFile(filled, 'followup.xlsx'), rows);
const typed = r16 > 0 && expCol > 0 ? 1 : 0;
assert.ok(back && back.changed === 4 + typed, `تعديلات مقروءة: ${back?.changed}`);
const fixed = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings: { dateOrder: 'dmy' }, overrides: back.overrides });
const fx = Object.fromEntries(fixed.map((r) => [r.src, r]));
assert.equal(fx[13].cells.email.status, 'ok');
assert.equal(fx[13].cells.contract_end_date.value, '2027-12-31');
assert.equal(fx[17].cells.salary_basic_salary_amount.value, 7200);
assert.equal(fx[11].cells.department.value, 'المالية');
if (typed) assert.equal(fx[16].cells.national_id_expiry_date.value, '2031-03-15', 'تاريخ مكتوب نصًا يُفهم');
assert.equal(parseFollowup(readClientFile(cbuf.buffer.slice(cbuf.byteOffset, cbuf.byteOffset + cbuf.byteLength), 'x.xlsx'), rows), null, 'ملف عادي ليس ملف استكمال');
console.log('followup round-trip OK:', back.changed, 'تعديلات على', back.rowsMatched, 'موظف');

console.log('\nOK — كل الفحوص نجحت. صفوف الرفع:', okRows.length);
