// اختبار دقة القراءة: لكل ملف عميل "مخربط" في fixtures/ نشغّل الأداة تلقائيًا بالكامل (بدون أي تدخل يدوي)
// ونقارن (1) مطابقة الأعمدة و(2) القيم النهائية لكل موظف مع الحقيقة المعروفة في truth.json.
import fs from 'node:fs';
import * as engine from '../src/engine.js';
import { readTemplate } from '../src/template.js';

const { readClientFile, buildRecords, processRecords } = engine;
const dirArg = process.argv.indexOf('--dir');
const FX = new URL(`./${dirArg > 0 ? process.argv[dirArg + 1] : 'fixtures'}/`, import.meta.url);
const truth = JSON.parse(fs.readFileSync(new URL('truth.json', FX), 'utf8'));
const tpl = await readTemplate(fs.readFileSync(new URL('./template_rich.xlsx', import.meta.url)));
const verbose = process.argv.includes('-v');
const only = process.argv.find((a) => a.endsWith('.xlsx') || a.endsWith('.csv'));

const norm = (k, v) => {
  if (v === null || v === undefined || v === '') return '';
  if (k === 'work_phone_number') return String(v).replace(/\D/g, '').slice(-9);
  if (k === 'salary_basic_salary_amount') return String(Number(v));
  return String(v);
};

let totalMap = 0, okMap = 0, totalVal = 0, okVal = 0, failedFiles = 0;
for (const [file, T] of Object.entries(truth)) {
  if (only && file !== only) continue;
  const buf = fs.readFileSync(new URL(file, FX));
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const sheets = readClientFile(ab, file).filter((s) => s.rows.length);
  const a = engine.analyzeWorkbook(sheets, tpl.lists);
  const { prepared: sheet, mapping, headers } = a;
  const headerIdx = a.hdr.idx;
  const problems = [];

  // (1) المطابقة
  const colText = (c) => headers.find((h) => h.col === c)?.text ?? `#${c}`;
  for (const [key, col] of Object.entries(T.expectMap)) {
    if (!engine.MAPPABLE_KEYS.has(key)) continue;
    totalMap++;
    let want = col;
    if (T.synthetic[key]) want = headers.find((h) => h.synthetic === T.synthetic[key])?.col;
    const got = mapping[key]?.col;
    if (got === want) okMap++;
    else problems.push(`  map ${key}: توقعت [${colText(want)}] ووجدت [${got === undefined ? '—' : colText(got)}]`);
  }
  for (const [key, synth] of Object.entries(T.synthetic)) {
    if (T.expectMap[key] !== undefined) continue;
    totalMap++;
    const want = headers.find((h) => h.synthetic === synth)?.col;
    if (want !== undefined && mapping[key]?.col === want) okMap++;
    else problems.push(`  map ${key}: توقعت العمود المشتق ${synth} ووجدت [${mapping[key] ? colText(mapping[key].col) : '—'}]`);
  }
  const expectedCols = new Map(Object.entries(T.expectMap).map(([k, c]) => [c, k]));
  for (const [key, m] of Object.entries(mapping)) {
    if (T.expectMap[key] !== undefined || T.synthetic[key]) continue;
    const owner = expectedCols.get(m.col);
    problems.push(`  map زائد ${key} ← [${colText(m.col)}]${owner ? ` (العمود يخص ${owner})` : ''}`);
  }

  // (2) القيم
  const recs = buildRecords(sheet, mapping, {}, { nameSplit: 'rest' });
  const rows = processRecords(recs, { lists: tpl.lists, required: tpl.required, settings: { dateOrder: 'dmy', addCC: true } });
  const byId = new Map(rows.map((r) => [String(r.cells.national_id_number.value), r]));
  if (rows.length !== T.records.length) problems.push(`  عدد الصفوف ${rows.length} والمتوقع ${T.records.length}`);
  for (const exp of T.records) {
    const r = byId.get(exp.national_id_number);
    for (const k of T.fields) {
      if (exp[k] === undefined) continue;
      totalVal++;
      const got = r ? norm(k, r.cells[k].value) : '∅';
      const want = norm(k, exp[k]);
      if (got === want) okVal++;
      else if (problems.length < 60) problems.push(`  قيمة ${exp.national_id_number} ${k}: توقعت "${want}" ووجدت "${got}"${r ? ` (${r.cells[k].status}: ${r.cells[k].msg || ''}, الأصل "${r.cells[k].raw}")` : ''}`);
    }
  }
  const status = problems.length ? 'FAIL' : 'PASS';
  if (problems.length) failedFiles++;
  console.log(`${status} ${file}  (شيت "${sheet.name}", ترويسة ${headerIdx + 1}, ${rows.length} صف)`);
  if (problems.length) console.log((verbose ? problems : problems.slice(0, 12)).join('\n') + (problems.length > 12 && !verbose ? `\n  … +${problems.length - 12}` : ''));
}
console.log(`\nمطابقة الأعمدة: ${okMap}/${totalMap}   القيم: ${okVal}/${totalVal}   ملفات فاشلة: ${failedFiles}`);
process.exitCode = failedFiles ? 1 : 0;
