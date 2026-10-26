// شروط رفع قيود HR الفعلية (من رسائل رفض حقيقية) على كل صف يدخل ملف الرفع:
//  - الراتب يحتاج تاريخ بدء ("Start date can't be blank")
//  - العقد غير محدد المدة (أو الفاضي للسعودي) لا يقبل تاريخ نهاية
//  - رقم الجوال لازم يكون صالحًا ("Work phone number is invalid")
// كما يحاكي "طبّقها على باقي الفاضي" لتاريخ نهاية العقد ويتأكد أنها لا تصل للسعوديين.
import fs from 'node:fs';
import assert from 'node:assert/strict';
import { readClientFile, analyzeWorkbook, buildRecords, processRecords, normKey } from '../src/engine.js';
import { readTemplate } from '../src/template.js';

const tpl = await readTemplate(fs.readFileSync(new URL('./template_rich.xlsx', import.meta.url)));
const FX = new URL('./fixtures/', import.meta.url);
const files = fs.readdirSync(FX).filter((f) => /\.(xlsx|csv)$/.test(f));
const settings = { dateOrder: 'dmy', addCC: true, nameSplit: 'rest', translit: true };
const SA = normKey('السعودية');
let checked = 0;

for (const file of files) {
  const buf = fs.readFileSync(new URL(file, FX));
  const a = analyzeWorkbook(readClientFile(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), file), tpl.lists);
  const recs = buildRecords(a.prepared, a.mapping, { residence_country: 'السعودية' }, settings);
  let rows = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings });

  // محاكاة "طبّقها على باقي الفاضي المطلوب" لنهاية العقد: فقط الخانات الحمراء الفاضية
  const targets = rows.filter((r) => r.cells.contract_end_date.status === 'error' && !r.cells.contract_end_date.raw).map((r) => r.src);
  const overrides = Object.fromEntries(targets.map((s) => [s, { contract_end_date: '2028-10-12' }]));
  rows = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings, overrides });
  for (const s of targets) {
    const r = rows.find((x) => x.src === s);
    assert.notEqual(normKey(r.cells.nationality.value), SA, `${file} صف ${s}: تعبئة نهاية العقد وصلت لسعودي`);
  }

  for (const r of rows.filter((x) => x.status !== 'error')) {
    const c = r.cells;
    const where = `${file} صف ${r.src}`;
    if (c.salary_basic_salary_amount.value) assert.ok(c.salary_start_date.value, `${where}: راتب بدون تاريخ بدء`);
    if (c.contract_end_date.value) {
      assert.equal(normKey(c.contract_type.value), normKey('محدد المدة'), `${where}: تاريخ نهاية مع نوع عقد "${c.contract_type.value}"`);
    }
    if (c.work_phone_number.value) assert.match(String(c.work_phone_number.value), /^\+\d{9,15}$/, `${where}: جوال غير صالح لقيود "${c.work_phone_number.value}"`);
    checked++;
  }
}
// حالات محددة
const tplLists = tpl.lists;
const one = (vals, ov = {}) => processRecords([{ src: 1, vals }], { lists: tplLists, required: tpl.required, settings, overrides: { 1: ov } })[0].cells;
let c = one({ nationality: 'سعودي', contract_type: 'غير محدد المدة', contract_end_date: '2028-01-01' });
assert.equal(c.contract_end_date.status, 'error', 'غير محدد المدة + تاريخ نهاية = خطأ');
c = one({ nationality: 'مصري', contract_end_date: '2028-01-01' });
assert.equal(c.contract_type.value, 'محدد المدة'); assert.equal(c.contract_type.status, 'fixed');
c = one({ nationality: 'سعودي', contract_end_date: '2028-01-01' });
assert.equal(c.contract_type.value, 'محدد المدة'); assert.equal(c.contract_type.status, 'review', 'سعودي بعقد محدد: للتأكيد');
c = one({ salary_basic_salary_amount: 5000, join_date: '2024-03-01' });
assert.equal(c.salary_start_date.value, '2024-03-01'); assert.equal(c.salary_pay_frequency.value, 'Monthly'); assert.equal(c.salary_basic_salary_kind.value, 'Fixed');
c = one({ salary_basic_salary_amount: 5000, join_date: '2024-03-01', salary_start_date: '2024-04-01' });
assert.equal(c.salary_start_date.value, '2024-04-01', 'تاريخ بدء موجود لا يُستبدل');
c = one({ work_phone_number: '050123456' });
assert.equal(c.work_phone_number.status, 'error', 'جوال ناقص رقم = خطأ');
c = one({ work_phone_number: '0114567890' });
assert.equal(c.work_phone_number.value, '+966114567890', 'هاتف ثابت يتحول دولي');
console.log(`OK — شروط قيود صحيحة على ${checked} صف جاهز للرفع من ${files.length} ملف`);
