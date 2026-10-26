// يجهّز ملف استكمال "معبّأ من العميل" لملف 11 لتجربة الرفع بالواجهة: node make_filled_followup.mjs <out.xlsx> [--blank <fixture>]
// مع --blank: ملف استكمال كما يستلمه العميل (بدون تعبئة) للملف المحدد
import fs from 'node:fs';
import ExcelJS from 'exceljs';
import { readClientFile, analyzeWorkbook, buildRecords, processRecords } from '../src/engine.js';
import { readTemplate, buildFollowup } from '../src/template.js';
const tpl = await readTemplate(fs.readFileSync(new URL('./template_rich.xlsx', import.meta.url)));
const bi = process.argv.indexOf('--blank');
const src = bi > 0 ? process.argv[bi + 1] : '11_names_one_language.xlsx';
const buf = fs.readFileSync(new URL(`./fixtures/${src}`, import.meta.url));
const a = analyzeWorkbook(readClientFile(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), 'x.xlsx'), tpl.lists);
const rows = processRecords(buildRecords(a.prepared, a.mapping, {}, { nameSplit: 'rest' }), { lists: tpl.lists, required: tpl.required, settings: { translit: true } });
if (bi > 0) { fs.writeFileSync(process.argv[2], await buildFollowup(rows, tpl, { includeReview: true })); console.log('blank followup written'); process.exit(0); }
const wb = new ExcelJS.Workbook();
await wb.xlsx.load(await buildFollowup(rows, tpl, { includeReview: true }));
const ws = wb.getWorksheet('استكمال البيانات');
const keys = ws.getRow(2).values;
let n = 0;
for (let r = 3; r <= ws.rowCount; r++) {
  const c = ws.getCell(r, keys.indexOf('residence_country'));
  if (c.fill?.fgColor?.argb === 'FFFFF3C4') { c.value = 'السعودية'; n++; }
  const t = ws.getCell(r, keys.indexOf('first_name_ar'));
  if (t.value === 'خوسيه') { t.value = 'جوزيه'; n++; }
}
await wb.xlsx.writeFile(process.argv[2]);
console.log('filled cells', n, 'rows', ws.rowCount - 2);
