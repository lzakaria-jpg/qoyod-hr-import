// يرفع رقم الإصدار في src/version.js: 1.0.1 ← 1.0.2 … 1.0.9 ← 1.1.0 … 1.9.9 ← 2.0.0
import fs from 'node:fs';
const file = new URL('../src/version.js', import.meta.url);
const src = fs.readFileSync(file, 'utf8');
const m = src.match(/'QHR\.(\d+)\.(\d)\.(\d)'/);
if (!m) throw new Error('VERSION not found in src/version.js');
let [a, b, c] = m.slice(1).map(Number);
if (++c > 9) { c = 0; if (++b > 9) { b = 0; a++; } }
const next = `QHR.${a}.${b}.${c}`;
fs.writeFileSync(file, src.replace(m[0], `'${next}'`));
console.log(`${m[0].slice(1, -1)} -> ${next}`);
