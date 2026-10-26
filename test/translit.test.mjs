// قياس جودة الترجمة الحرفية بالاتجاهين على أزواج الأسماء (عربي/إنجليزي) الموجودة في ملفات الاختبار
import fs from 'node:fs';
import assert from 'node:assert/strict';
import { transliterate, transliterateAr } from '../src/translit.js';

const normA = (s) => String(s).replace(/[ً-ٰٟـ]/g, '').replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/ة/g, 'ه').replace(/\s+/g, ' ').trim();
const pairs = new Map();
for (const dir of ['fixtures', 'fixtures_blind']) {
  const p = new URL(`./${dir}/truth.json`, import.meta.url);
  if (!fs.existsSync(p)) continue;
  for (const T of Object.values(JSON.parse(fs.readFileSync(p, 'utf8')))) {
    for (const r of T.records) {
      for (const [a, e] of [['first_name_ar', 'first_name'], ['last_name_ar', 'last_name']]) {
        if (!r[a] || !r[e]) continue;
        // نقارن كلمة بكلمة فقط إذا عدد الكلمات متساوٍ
        const A = r[a].split(' '), E = r[e].replace(/-/g, ' ').replace(/\bAl /gi, 'Al').split(' ');
        if (A.length === E.length) A.forEach((w, i) => pairs.set(`${w}|${E[i]}`, [w, E[i]]));
      }
    }
  }
}
let okAr = 0, okEn = 0;
const badAr = [], badEn = [];
for (const [ar, en] of pairs.values()) {
  const gotAr = transliterate(en).ar;
  if (normA(gotAr) === normA(ar)) okAr++; else badAr.push(`${en} → ${gotAr} (الصحيح ${ar})`);
  const gotEn = transliterateAr(ar).en;
  if (gotEn.toLowerCase().replace(/[^a-z]/g, '') === en.toLowerCase().replace(/[^a-z]/g, '')) okEn++; else badEn.push(`${ar} → ${gotEn} (بالملف ${en})`);
}
console.log(`إنجليزي ← عربي: ${okAr}/${pairs.size} مطابق حرفيًا`);
console.log(`عربي ← إنجليزي: ${okEn}/${pairs.size} مطابق لتهجئة الملف`);
if (process.argv.includes('-v')) { console.log('\n--- en→ar:\n' + badAr.join('\n')); console.log('\n--- ar→en:\n' + badEn.join('\n')); }

// حالات أساسية يجب أن تنجح دائمًا
for (const [en, ar] of [['Mohammed Abdullah Alotaibi', 'محمد عبدالله العتيبي'], ['Abdul Rahman', 'عبدالرحمن'], ['Al-Qahtani', 'القحطاني'], ['Maria Santos', 'ماريا سانتوس'], ['Sarah', 'سارة'], ['Rajesh Kumar', 'راجيش كومار']]) assert.equal(transliterate(en).ar, ar, en);
for (const [ar, en] of [['محمد عبدالله العتيبي', 'Mohammed Abdullah Alotaibi'], ['عبد الرحمن', 'Abdulrahman'], ['فاطمة', 'Fatima'], ['نورة الشهري', 'Noura Alshehri']]) assert.equal(transliterateAr(ar).en, en, ar);
assert.ok(transliterate('Zzyxq').ar.length > 0, 'اسم غير معروف يُترجم صوتيًا');
assert.ok(transliterateAr('شمعون').en.length > 0);
console.log('OK');
