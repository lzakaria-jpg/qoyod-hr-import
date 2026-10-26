import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { FIELDS, FIELD_BY_KEY, GROUPS, VIRTUAL_FIELDS } from './fields.js';
import {
  readClientFile, detectHeader, prepareSheet, mergePrepared, autoMap, analyzeWorkbook, buildRecords, processRecords,
  unmatchedValues, approxValues, cellText, columnSample, normKey, parseFollowup,
} from './engine.js';
import { readTemplate, fillTemplate, buildReport, buildFollowup, followupRows } from './template.js';
import DateField, { fmtDate } from './DateField.jsx';

const STEPS = ['الملفات', 'مطابقة الأعمدة', 'المراجعة والتصحيح', 'التصدير'];
const LS_KEY = 'qhr-import-template-v1';
const LS_MAPS = 'qhr-import-maps-v1';
const PAGE = 40;
const STATUS_LABEL = { ok: 'سليم', fixed: 'مُصحَّح', review: 'للمراجعة', error: 'خطأ', empty: 'فاضي' };

/* ---------- أدوات صغيرة ---------- */
const b64ToBuf = (b64) => { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u.buffer; };
const bufToB64 = (buf) => { let s = ''; const u = new Uint8Array(buf); for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode(...u.subarray(i, i + 0x8000)); return btoa(s); };
const lsGet = (k) => { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch { return null; } };
const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch { return false; } };
const download = async (bytes, name) => {
  // داخل صفحة منشورة: قدرة التنزيل الرسمية (تطلب موافقة المستخدم). خارجها: رابط تنزيل عادي.
  try {
    const d = await window.claude?.use?.('downloads');
    if (d) { await d.save({ filename: name, data: bytes }); return; }
  } catch (e) { if (e && e.code === 'declined') return; }
  const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
};
const Icon = ({ d, size = 22 }) => (
  <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d={d} /></svg>
);
const ICONS = {
  upload: 'M12 16V4m0 0L7 9m5-5l5 5M4 16v3a1 1 0 001 1h14a1 1 0 001-1v-3',
  file: 'M14 3H7a2 2 0 00-2 2v14a2 2 0 002 2h10a2 2 0 002-2V8l-5-5zM14 3v5h5',
  check: 'M5 12l5 5L20 7',
  users: 'M16 11a4 4 0 10-8 0 4 4 0 008 0zM4 20a8 8 0 0116 0',
  down: 'M12 4v12m0 0l-5-5m5 5l5-5M4 20h16',
};
// بصمة شكل الملف (عناوين الأعمدة) لحفظ المطابقة واسترجاعها لملفات نفس العميل
const sigOf = (P) => P.cols.filter((c) => !c.synthetic && c.own).map((c) => normKey(c.own)).join('|');
const colKey = (c) => (c.synthetic ? c.synthetic : normKey(c.own || c.text));

function DropZone({ title, hint, accept, onFile, icon = 'upload' }) {
  const ref = useRef(null);
  const [over, setOver] = useState(false);
  return (
    <div
      className={`drop ${over ? 'over' : ''}`}
      onClick={() => ref.current?.click()}
      onDragOver={(e) => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); const f = e.dataTransfer.files?.[0]; if (f) onFile(f); }}
    >
      <Icon d={ICONS[icon]} size={30} />
      <div className="big">{title}</div>
      <div className="hint">{hint}</div>
      <input ref={ref} type="file" accept={accept} hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ''; }} />
    </div>
  );
}

const isIsoDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
// محرر حقل واحد حسب نوعه: تقويم للتواريخ، قائمة للقوائم، نص للباقي
function FieldEditor({ f, cell, list, onCommit, onCancel, autoFocus = false, inline = false }) {
  const start = cell.value !== '' && cell.value !== undefined ? cell.value : cell.raw;
  if (f.type === 'date') {
    return <DateField value={isIsoDate(cell.value) ? cell.value : (cell.raw || '')} autoFocus={autoFocus} autoOpen={autoFocus} onChange={onCommit} onCancel={onCancel} />;
  }
  if (list?.length) {
    return (
      <select autoFocus={autoFocus} value={list.includes(cell.value) ? cell.value : ''} onChange={(e) => onCommit(e.target.value)} onBlur={() => !inline && onCancel?.()}>
        <option value="">{cell.value && !list.includes(cell.value) ? `— ${cell.value} (غير موجود بالقائمة) —` : '— فاضي —'}</option>
        {list.map((v) => <option key={v} value={v}>{v}</option>)}
      </select>
    );
  }
  const ltr = ['email', 'phone', 'iban', 'id', 'code', 'amount', 'name_en'].includes(f.type);
  return (
    <input type="text" dir={ltr ? 'ltr' : undefined} autoFocus={autoFocus} defaultValue={start ?? ''} key={String(start)}
      onKeyDown={(e) => { if (e.key === 'Enter') { e.currentTarget.blur(); } if (e.key === 'Escape') { e.currentTarget.value = start ?? ''; onCancel?.(); } }}
      onBlur={(e) => { const v = e.currentTarget.value; if (v !== String(start ?? '')) onCommit(v); else onCancel?.(); }} />
  );
}

function FileButton({ label, accept, onFile, className = 'btn' }) {
  const ref = useRef(null);
  return (
    <>
      <button className={className} onClick={() => ref.current?.click()}>{label}</button>
      <input ref={ref} type="file" accept={accept} hidden onChange={(e) => { const f = e.target.files?.[0]; if (f) onFile(f); e.target.value = ''; }} />
    </>
  );
}

function makePrepared(client, view) {
  if (!client) return null;
  if (view.merged) {
    const list = view.mergeNames.map((n) => client.sheets.find((s) => s.name === n)).filter(Boolean);
    return mergePrepared(list.map((s) => prepareSheet(s, detectHeader(s.rows))));
  }
  const s = client.sheets[view.sheetIdx];
  return s ? prepareSheet(s, view.hdr) : null;
}

/* ============================== التطبيق ============================== */
export default function App() {
  const [step, setStep] = useState(0);
  const [toast, setToast] = useState('');
  const [err, setErr] = useState('');

  // القالب الرسمي
  const [tplFile, setTplFile] = useState(null); // {name, buffer, tpl, saved}
  // ملف العميل
  const [client, setClient] = useState(null); // {name, sheets, mergeNames}
  const [view, setView] = useState({ merged: false, mergeNames: [], sheetIdx: 0, hdr: { top: -1, idx: -1 } });
  const [mapping, setMapping] = useState({}); // key -> {col, score, by}
  const [mapTouched, setMapTouched] = useState(false);
  const [constants, setConstants] = useState({});
  const [settings, setSettings] = useState({ dateOrder: 'dmy', addCC: true, nameSplit: 'rest', translit: true }); // قالب قيود يكتب الجوال +966…
  const [fuReview, setFuReview] = useState(true);
  const [restored, setRestored] = useState(false);
  const [showRaw, setShowRaw] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [mapView, setMapView] = useState('simple'); // simple = كروت أعمدة العميل ، detailed = كل حقول قيود
  // المراجعة
  const [userMaps, setUserMaps] = useState({});
  const [overrides, setOverrides] = useState({});
  const [reviewed, setReviewed] = useState({}); // {src: {field: القيمة وقت المراجعة}}
  const markReviewed = (row, keys) => setReviewed((m) => {
    const n = { ...m, [row.src]: { ...(m[row.src] || {}) } };
    for (const k of keys) n[row.src][k] = String(row.cells[k].value ?? '');
    return n;
  });
  const unmarkReviewed = (src, key) => setReviewed((m) => { const n = { ...m, [src]: { ...(m[src] || {}) } }; delete n[src][key]; return n; });
  const [excluded, setExcluded] = useState(() => new Set());
  const [filter, setFilter] = useState('all');
  const [query, setQuery] = useState('');
  const [page, setPage] = useState(0);
  const [showAllCols, setShowAllCols] = useState(false);
  const [drawerRow, setDrawerRow] = useState(null);
  const [drawerAll, setDrawerAll] = useState(false);
  const [busy, setBusy] = useState(false);
  const [askSample, setAskSample] = useState(false);

  const flash = useCallback((m) => { setToast(m); setTimeout(() => setToast(''), 3200); }, []);
  const lists = tplFile?.tpl.lists;

  // استرجاع القالب المحفوظ
  useEffect(() => {
    (async () => {
      try {
        const saved = lsGet(LS_KEY);
        if (!saved) return;
        const buffer = b64ToBuf(saved.b64);
        const tpl = await readTemplate(buffer);
        setTplFile({ name: saved.name, buffer, tpl, saved: true });
      } catch { /* لا مشكلة */ }
    })();
  }, []);

  const loadTemplate = async (file) => {
    setErr('');
    try {
      const buffer = await file.arrayBuffer();
      const tpl = await readTemplate(buffer);
      setTplFile({ name: file.name, buffer, tpl, saved: false });
      if (lsSet(LS_KEY, { name: file.name, b64: bufToB64(buffer) })) setTplFile((t) => t && { ...t, saved: true });
    } catch (e) { setErr(`تعذّر قراءة القالب: ${e.message}`); }
  };
  const forgetTemplate = () => { try { localStorage.removeItem(LS_KEY); } catch { /* */ } setTplFile(null); };

  // تطبيق مطابقة محفوظة لنفس شكل الملف إن وجدت
  const applySaved = (P, m) => {
    const saved = lsGet(LS_MAPS)?.[sigOf(P)];
    if (!saved) return { mapping: m, restored: false };
    const out = {};
    for (const [k, ck] of Object.entries(saved.map)) {
      const c = P.cols.find((x) => colKey(x) === ck);
      if (c) out[k] = { col: c.col, score: 1, by: 'saved' };
    }
    if (saved.constants) setConstants(saved.constants);
    if (saved.settings) setSettings((s) => ({ ...s, ...saved.settings }));
    return { mapping: out, restored: true };
  };
  const saveProfile = () => {
    if (!prepared) return;
    const all = lsGet(LS_MAPS) || {};
    const map = {};
    for (const [k, m] of Object.entries(mapping)) { const c = prepared.cols.find((x) => x.col === m.col); if (c) map[k] = colKey(c); }
    all[sigOf(prepared)] = { map, constants, settings, at: Date.now(), file: client?.name };
    lsSet(LS_MAPS, all);
  };

  const remap = (P, l = lists) => {
    if (!P) return;
    setMapping(autoMap(P, l).mapping);
    setMapTouched(false);
  };

  const loadClient = async (file) => {
    setErr('');
    try {
      const ab = await file.arrayBuffer();
      const sheets = readClientFile(ab, file.name);
      if (!sheets.some((s) => s.rows.length)) throw new Error('الملف فاضي');
      const a = analyzeWorkbook(sheets, lists);
      const v = { merged: a.merged, mergeNames: a.mergeNames, sheetIdx: a.sheetIdx, hdr: a.hdr };
      setClient({ name: file.name, sheets, mergeNames: a.mergeNames });
      setView(v);
      setConstants({}); setUserMaps({}); setOverrides({}); setExcluded(new Set());
      const sv = applySaved(a.prepared, a.mapping);
      setMapping(sv.mapping); setRestored(sv.restored); setMapTouched(sv.restored);
      if (sv.restored) flash('طُبّقت مطابقة محفوظة سابقًا لملف بنفس الشكل');
      else if (a.merged) flash(`دُمجت ${a.mergeNames.length} شيتات بنفس البنية: ${a.mergeNames.join('، ')}`);
    } catch (e) { setErr(`تعذّر قراءة ملف العميل: ${e.message}`); }
  };

  // إعادة المطابقة عند تحميل القالب بعد ملف العميل (القوائم تحسّن الاكتشاف بالمحتوى)
  useEffect(() => {
    if (client && !mapTouched) remap(makePrepared(client, view), lists);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [tplFile]);

  const prepared = useMemo(() => makePrepared(client, view), [client, view]);
  const headers = prepared?.cols || [];
  const rawSheet = client && !view.merged ? client.sheets[view.sheetIdx] : null;

  const changeView = (patch) => {
    const v = { ...view, ...patch };
    setView(v);
    remap(makePrepared(client, v));
  };

  const records = useMemo(() => {
    if (!prepared || !tplFile) return [];
    const cleanMap = Object.fromEntries(Object.entries(mapping).filter(([, v]) => v && v.col !== undefined && v.col !== null && v.col !== ''));
    return buildRecords(prepared, cleanMap, constants, settings);
  }, [prepared, mapping, constants, settings, tplFile]);

  const rows = useMemo(() => {
    if (!tplFile || !records.length) return [];
    const r = processRecords(records, { lists: tplFile.tpl.lists, required: tplFile.tpl.required, settings, userMaps, overrides, reviewed });
    r.forEach((x) => { x.excluded = excluded.has(x.src); });
    return r;
  }, [records, tplFile, settings, userMaps, overrides, excluded, reviewed]);

  const live = rows.filter((r) => !r.excluded);
  const counts = {
    all: rows.length,
    ok: live.filter((r) => r.status === 'ok').length,
    review: live.filter((r) => r.status === 'review').length,
    error: live.filter((r) => r.status === 'error').length,
    excluded: rows.length - live.length,
  };
  const reviewCells = live.reduce((s, r) => s + FIELDS.filter((f) => r.cells[f.key].status === 'review').length, 0);
  // "للمراجعة" لا يمنع التصدير أبدًا: فقط الأخطاء الحمراء
  const exportable = live.filter((r) => r.status !== 'error');

  const requiredKeys = useMemo(() => [...(tplFile?.tpl.required || [])], [tplFile]);
  const covered = (k) => (mapping[k] && mapping[k].col !== undefined && mapping[k].col !== '' && mapping[k].col !== null)
    || constants[k] || (k === 'first_name' && (mapping.full_name_en)) || (k === 'last_name' && (mapping.full_name_en))
    || (k === 'first_name_ar' && mapping.full_name_ar) || (k === 'last_name_ar' && mapping.full_name_ar);
  const missingRequired = requiredKeys.filter((k) => !covered(k));
  const usedCols = new Set(Object.values(mapping).map((m) => m.col));
  const unusedCols = headers.filter((h) => !usedCols.has(h.col));

  const SampleWarn = tplFile?.tpl.isSample && (
    <div className="banner err">
      <b>القالب المحمّل ({tplFile?.name}) تجريبي للاختبار فقط.</b> قوائم الفروع والأقسام وملفات الدوام فيه مخترعة، فقيود يرفض الملف المصدّر منه
      (مثل: "location ليس ضمن الخيارات المسموح بها"). للرفع الحقيقي: من قيود ← الموظفين ← إضافة جماعية ← تنزيل القالب، لنفس شركة العميل، ثم اضغط "استبدال" وارفعه هنا.
    </div>
  );

  /* --------------------------- الخطوة 1 --------------------------- */
  const StepFiles = (
    <>
      <div className="grid2">
        <div className="card">
          <h2>١. قالب قيود الرسمي</h2>
          <p className="sub">نزّله من الموظفين ← إضافة جماعية ← تنزيل القالب. القوائم (الأقسام، المواقع، ملفات الدوام) خاصة بكل عميل، فاستخدم قالب نفس شركة العميل.</p>
          {!tplFile ? (
            <DropZone title="اسحب القالب هنا أو اضغط للاختيار" hint="ملف xlsx الذي نزّلته من قيود" accept=".xlsx" onFile={loadTemplate} icon="file" />
          ) : (
            <>
              <div className="filebox">
                <div className="ico"><Icon d={ICONS.check} /></div>
                <div className="meta">
                  <div className="name">{tplFile.name}</div>
                  <div className="info">{tplFile.tpl.headers.length} عمودًا · {tplFile.tpl.required.size} إلزاميًا · {Object.keys(tplFile.tpl.lists).length} قائمة{tplFile.saved ? ' · محفوظ على هذا المتصفح' : ''}</div>
                </div>
                <button className="btn small" onClick={forgetTemplate}>استبدال</button>
              </div>
              <div className="chips">
                {['department', 'location', 'attendance_setting'].map((k) => (
                  <span key={k} className="chip">{FIELD_BY_KEY[k].label}: {(tplFile.tpl.lists[k] || []).length}</span>
                ))}
                <span className="chip gray">الجنسيات: {(tplFile.tpl.lists.nationality || []).length}</span>
              </div>
              {SampleWarn}
              {tplFile.tpl.missingHeaders.length > 0 && <div className="banner warn">أعمدة ناقصة بالقالب: {tplFile.tpl.missingHeaders.join('، ')}</div>}
              {tplFile.tpl.unknownHeaders.length > 0 && <div className="banner warn">أعمدة جديدة بالقالب لا تعرفها الأداة: {tplFile.tpl.unknownHeaders.join('، ')}</div>}
            </>
          )}
        </div>
        <div className="card">
          <h2>٢. ملف العميل</h2>
          <p className="sub">أي شكل: عناوين عربي أو إنجليزي أو بدون عناوين، ترويسة بصفين، شيت لكل فرع، أسماء كاملة بعمود واحد، تواريخ هجري وميلادي، CSV بأي ترميز. الأداة تكتشف وتطابق وتنظّف.</p>
          {!client ? (
            <DropZone title="اسحب ملف العميل هنا أو اضغط للاختيار" hint="xlsx أو xls أو csv" accept=".xlsx,.xls,.xlsm,.csv,.txt,.tsv" onFile={loadClient} icon="users" />
          ) : (
            <>
              <div className="filebox">
                <div className="ico"><Icon d={ICONS.check} /></div>
                <div className="meta">
                  <div className="name">{client.name}</div>
                  <div className="info">
                    {client.sheets.length} شيت · {view.merged ? `دمج: ${view.mergeNames.join(' + ')}` : `${client.sheets[view.sheetIdx]?.name}`}
                    {' · '}{view.merged ? 'الترويسة لكل شيت تلقائيًا' : view.hdr.idx < 0 ? 'بدون ترويسة' : `الترويسة بالصف ${view.hdr.idx + 1}`}
                  </div>
                </div>
                <button className="btn small" onClick={() => setClient(null)}>استبدال</button>
              </div>
              <div className="chips">
                <span className="chip ok">{Object.keys(mapping).length} عمودًا طابقتهم تلقائيًا</span>
                {Object.values(mapping).some((m) => m.by === 'content') && <span className="chip fixed">{Object.values(mapping).filter((m) => m.by === 'content').length} منها من المحتوى</span>}
                <span className="chip gray">{records.length || prepared?.data.length || 0} موظف</span>
                {restored && <span className="chip review">مطابقة محفوظة</span>}
              </div>
            </>
          )}
        </div>
      </div>
      {err && <div className="banner err">{err}</div>}
      <div className="actions">
        <span />
        <div className="right"><button className="btn primary" disabled={!tplFile || !client} onClick={() => setStep(1)}>التالي: مطابقة الأعمدة</button></div>
      </div>
    </>
  );

  /* --------------------------- الخطوة 2 --------------------------- */
  const setMap = (key, col) => {
    setMapTouched(true);
    setMapping((m) => {
      const n = { ...m };
      if (col === '' || col === null) delete n[key]; else n[key] = { col: Number(col), score: 1, by: 'manual' };
      return n;
    });
  };
  const setConst = (key, v) => setConstants((c) => { const n = { ...c }; if (v === '') delete n[key]; else n[key] = v; return n; });
  const sampleOf = (key) => {
    const m = mapping[key];
    if (!m || !prepared) return '';
    return columnSample(prepared, m.col, 3).join('  ·  ');
  };
  const confBadge = (key) => {
    const m = mapping[key];
    if (!m) return null;
    if (m.by === 'content') return <span className="chip fixed conf" title="العنوان غير واضح، الأداة استنتجت الحقل من شكل القيم">من المحتوى</span>;
    if (m.by === 'header' && m.score < 0.9) return <span className="chip review conf">تقريبي — تأكد</span>;
    return null;
  };
  const derivedFrom = (k) => {
    if (['first_name', 'last_name'].includes(k) && mapping.full_name_en) return 'الاسم الكامل (إنجليزي)';
    if (['first_name_ar', 'last_name_ar'].includes(k) && mapping.full_name_ar) return 'الاسم الكامل (عربي)';
    return '';
  };
  const colLabel = (h) => `${h.text}${h.letter && h.letter !== '—' ? `  (${h.letter})` : ''}`;
  const rowFor = (f) => {
    const hasConstOption = !f.virtual;
    const listVals = tplFile?.tpl.lists[f.key];
    const missing = f.req && !covered(f.key);
    const dup = mapping[f.key] && Object.entries(mapping).some(([k, m]) => k !== f.key && m.col === mapping[f.key].col);
    return (
      <div key={f.key} className={`maprow ${missing ? 'missing' : ''}`}>
        <div className="lab">{f.label}{f.req && <span className="req">*</span>}<small>{f.key}</small></div>
        <select value={mapping[f.key]?.col ?? ''} onChange={(e) => setMap(f.key, e.target.value)}>
          <option value="">— غير مرتبط —</option>
          {headers.map((h) => <option key={h.col} value={h.col}>{usedCols.has(h.col) && mapping[f.key]?.col !== h.col ? '• ' : ''}{colLabel(h)}</option>)}
        </select>
        <div className="sample" title={sampleOf(f.key)}>
          {dup && <span className="chip error conf">نفس العمود مربوط بحقل آخر</span>}
          {mapping[f.key] ? <>{confBadge(f.key)} {sampleOf(f.key)}</> : derivedFrom(f.key) ? <span className="chip fixed">يُشتق من {derivedFrom(f.key)}</span> : ''}
        </div>
        {hasConstOption ? (
          listVals?.length ? (
            <select value={constants[f.key] ?? ''} onChange={(e) => setConst(f.key, e.target.value)} disabled={!!mapping[f.key]}>
              <option value="">قيمة ثابتة للكل…</option>
              {listVals.map((v) => <option key={v} value={v}>{v}</option>)}
            </select>
          ) : f.type === 'date' ? (
            mapping[f.key] ? <input type="text" disabled placeholder="قيمة ثابتة للكل…" /> : <DateField value={constants[f.key] ?? ''} placeholder="تاريخ ثابت للكل…" onChange={(iso) => setConst(f.key, iso)} />
          ) : (
            <input type="text" placeholder="قيمة ثابتة للكل…" value={constants[f.key] ?? ''} onChange={(e) => setConst(f.key, e.target.value)} disabled={!!mapping[f.key]} />
          )
        ) : <span />}
      </div>
    );
  };
  const salaryDefaults = () => {
    const L = tplFile.tpl.lists;
    const pick = (k, want) => (L[k] || []).find((v) => normKey(v) === normKey(want)) || '';
    setConstants((c) => ({
      ...c,
      ...(pick('salary_pay_frequency', 'Monthly') && { salary_pay_frequency: pick('salary_pay_frequency', 'Monthly') }),
      ...(pick('salary_basic_salary_kind', 'Fixed') && { salary_basic_salary_kind: pick('salary_basic_salary_kind', 'Fixed') }),
      ...(L.salary_input_time_type?.[0] && { salary_input_time_type: L.salary_input_time_type[0] }),
    }));
    if (mapping.join_date) setMapping((m) => ({ ...m, salary_start_date: m.join_date }));
    flash('طُبّقت: شهري، ثابت، أول خيار بأساس الوقت، وبدء الراتب = تاريخ الانضمام');
  };

  const RawPreview = rawSheet && (
    <div className="rawwrap">
      <table className="raw">
        <tbody>
          {rawSheet.rows.slice(0, 18).map((r, i) => {
            const inHdr = i >= view.hdr.top && i <= view.hdr.idx && view.hdr.idx >= 0;
            return (
              <tr key={i} className={inHdr ? 'hdr' : ''} onClick={() => changeView({ hdr: { top: i, idx: i } })} title="اضغط لجعل هذا الصف هو الترويسة">
                <th>{i + 1}</th>
                {r.slice(0, 14).map((c, j) => <td key={j}>{cellText(c)}</td>)}
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );

  /* ---- العرض المبسّط: كرت لكل عمود بملف العميل ---- */
  const keysOfCol = (col) => Object.entries(mapping).filter(([, m]) => m.col === col).map(([k]) => k);
  const setColField = (col, key) => {
    setMapTouched(true);
    setMapping((m) => {
      const n = {};
      for (const [k, v] of Object.entries(m)) if (v.col !== col) n[k] = v;
      if (key) n[key] = { col, score: 1, by: 'manual' };
      return n;
    });
  };
  const labelOf = (k) => (FIELD_BY_KEY[k] || VIRTUAL_FIELDS.find((v) => v.key === k))?.label || k;
  const cardState = (col) => {
    const ks = keysOfCol(col);
    if (!ks.length) return { cls: 'ignored', tag: 'متجاهَل' };
    const m = mapping[ks[0]];
    if (m.by === 'content') return { cls: 'check', tag: 'من المحتوى — تأكد' };
    if (m.by === 'header' && m.score < 0.9) return { cls: 'check', tag: 'تقريبي — تأكد' };
    if (m.by === 'manual') return { cls: 'mapped', tag: 'اخترته أنت' };
    if (m.by === 'saved') return { cls: 'mapped', tag: 'مطابقة محفوظة' };
    return { cls: 'mapped', tag: 'تلقائي' };
  };
  const ConstEditor = ({ f }) => {
    const list = tplFile.tpl.lists[f.key];
    if (list?.length) return (
      <select value={constants[f.key] ?? ''} onChange={(e) => setConst(f.key, e.target.value)}>
        <option value="">اختر قيمة للكل…</option>
        {list.map((v) => <option key={v} value={v}>{v}</option>)}
      </select>
    );
    if (f.type === 'date') return <DateField value={constants[f.key] ?? ''} placeholder="تاريخ للكل…" onChange={(iso) => setConst(f.key, iso)} />;
    return <input type="text" placeholder="قيمة للكل…" defaultValue={constants[f.key] ?? ''} key={constants[f.key] ?? ''} onBlur={(e) => setConst(f.key, e.target.value.trim())} />;
  };
  const mappedCount = headers.filter((h) => keysOfCol(h.col).length).length;
  const ColumnCards = (
    <div className="colgrid">
      {headers.map((h) => {
        const ks = keysOfCol(h.col);
        const st = cardState(h.col);
        const samples = columnSample(prepared, h.col, 3);
        return (
          <div key={h.col} className={`colcard ${st.cls}`}>
            <div className="colcard-head">
              <b title={h.text}>{h.text}</b>
              <span className="tag">{st.tag}</span>
            </div>
            <div className="samples">
              {samples.length ? samples.map((s, i) => <div key={i} dir="auto" title={s}>{s}</div>) : <div className="muted">— فاضي —</div>}
            </div>
            <div className="arrow">↓ يروح إلى</div>
            <select value={ks[0] || ''} onChange={(e) => setColField(h.col, e.target.value)}>
              <option value="">— تجاهل هذا العمود —</option>
              <optgroup label="الاسم الكامل (يُقسَّم تلقائيًا)">
                {VIRTUAL_FIELDS.map((v) => <option key={v.key} value={v.key}>{v.label}</option>)}
              </optgroup>
              {Object.entries(GROUPS).map(([g, gl]) => (
                <optgroup key={g} label={gl}>
                  {FIELDS.filter((f) => f.group === g).map((f) => {
                    const other = mapping[f.key] && mapping[f.key].col !== h.col;
                    return <option key={f.key} value={f.key}>{f.label}{tplFile.tpl.required.has(f.key) ? ' *' : ''}{other ? ' (مربوط بعمود آخر)' : ''}</option>;
                  })}
                </optgroup>
              ))}
            </select>
            {ks.length > 1 && <div className="extra">وكمان: {ks.slice(1).map((k) => <span key={k} className="chip gray">{labelOf(k)} <button className="x" onClick={() => setMapping((m) => { const n = { ...m }; delete n[k]; return n; })} aria-label="إزالة">×</button></span>)}</div>}
          </div>
        );
      })}
    </div>
  );
  const MissingRequired = missingRequired.length > 0 ? (
    <div className="card">
      <h2>حقول قيود الإلزامية الناقصة ({missingRequired.length})</h2>
      <p className="sub">ملف العميل ما فيه عمود لها. اختر لها عمود من الكروت فوق، أو ثبّت قيمة وحدة لكل الموظفين هنا (وتقدر تعدّلها لكل موظف بالمراجعة).</p>
      <div className="missgrid">
        {missingRequired.map((k) => (
          <div key={k} className="missrow">
            <b>{FIELD_BY_KEY[k].label}</b>
            <ConstEditor f={FIELD_BY_KEY[k]} />
          </div>
        ))}
      </div>
    </div>
  ) : (
    <div className="banner ok-banner">✓ كل حقول قيود الإلزامية لها مصدر</div>
  );

  const StepMap = tplFile && client && prepared && (
    <>
      <div className="card">
        <div className="maphead">
          <div>
            <h2>مطابقة ملف العميل مع قالب قيود</h2>
            <p className="sub" style={{ margin: 0 }}>
              قرأنا <b>{prepared.data.length}</b> موظف من {view.merged ? `الشيتات (${view.mergeNames.join(' + ')})` : `شيت "${client.sheets[view.sheetIdx]?.name}"`}
              {view.merged ? '' : view.hdr.idx < 0 ? ' — بدون ترويسة' : ` — الترويسة بالصف ${view.hdr.idx + 1}`}.
              {' '}راجع وين يروح كل عمود، وعدّل إذا لزم.
            </p>
          </div>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            <button className={`btn small ${showSettings ? 'primary' : ''}`} onClick={() => setShowSettings((x) => !x)}>⚙ إعدادات القراءة</button>
            <button className="btn small" onClick={() => setMapView((v) => (v === 'simple' ? 'detailed' : 'simple'))}>{mapView === 'simple' ? 'عرض تفصيلي (كل حقول قيود)' : 'العرض المبسّط'}</button>
          </div>
        </div>
        <div className="mapsummary">
          <span className="chip ok">{mappedCount} عمود مربوط</span>
          <span className="chip gray">{headers.length - mappedCount} متجاهَل</span>
          <span className={`chip ${missingRequired.length ? 'error' : 'ok'}`}>الإلزامي: {requiredKeys.length - missingRequired.length} / {requiredKeys.length}</span>
          {Object.values(mapping).some((m) => m.by === 'content' || (m.by === 'header' && m.score < 0.9)) && <span className="chip review">الأصفر: تأكد منه</span>}
        </div>
        {showSettings && (
        <div className="settings" style={{ marginTop: 12 }}>
          <div className="field"><label>الشيت</label>
            <select value={view.merged ? 'merged' : view.sheetIdx} onChange={(e) => {
              if (e.target.value === 'merged') changeView({ merged: true });
              else { const i = +e.target.value; changeView({ merged: false, sheetIdx: i, hdr: detectHeader(client.sheets[i].rows) }); }
            }}>
              {client.mergeNames.length >= 2 && <option value="merged">كل الشيتات المتشابهة ({client.mergeNames.join(' + ')})</option>}
              {client.sheets.map((s, i) => <option key={s.name} value={i} disabled={!s.rows.length}>{s.name}{s.hidden ? ' (مخفي)' : ''}{!s.rows.length ? ' (فاضي)' : ''}</option>)}
            </select>
          </div>
          {!view.merged && (
            <div className="field"><label>رقم صف الترويسة (0 = بدون ترويسة)</label>
              <input type="number" min="0" value={view.hdr.idx + 1} onChange={(e) => { const h = Math.max(-1, (+e.target.value || 0) - 1); changeView({ hdr: { top: h, idx: h } }); }} />
            </div>
          )}
          <div className="field"><label>التواريخ الملتبسة (مثل 04/05/1991)</label>
            <div className="seg">
              <button className={settings.dateOrder === 'dmy' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, dateOrder: 'dmy' }))}>يوم/شهر</button>
              <button className={settings.dateOrder === 'mdy' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, dateOrder: 'mdy' }))}>شهر/يوم</button>
            </div>
          </div>
          <div className="field"><label>تقسيم الاسم الكامل</label>
            <div className="seg">
              <button className={settings.nameSplit === 'rest' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, nameSplit: 'rest' }))}>الأول + الباقي</button>
              <button className={settings.nameSplit === 'last' ? 'on' : ''} onClick={() => setSettings((s) => ({ ...s, nameSplit: 'last' }))}>الأول + الأخير</button>
            </div>
          </div>
          <label className="switch"><input type="checkbox" checked={settings.addCC} onChange={(e) => setSettings((s) => ({ ...s, addCC: e.target.checked }))} />أضف +966 للجوالات السعودية المحلية (05…)</label>
          <label className="switch"><input type="checkbox" checked={settings.translit !== false} onChange={(e) => setSettings((s) => ({ ...s, translit: e.target.checked }))} />ترجمة حرفية للاسم الناقص (عربي ⇄ إنجليزي)</label>
          <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
            {!view.merged && <button className="btn small" onClick={() => setShowRaw((x) => !x)}>{showRaw ? 'إخفاء' : 'عرض'} الملف الخام</button>}
            <button className="btn small" onClick={() => remap(prepared)}>إعادة المطابقة التلقائية</button>
          </div>
        </div>
        )}
        {showSettings && showRaw && !view.merged && <><p className="sub" style={{ margin: '12px 0 6px' }}>أول 18 صفًا كما هي بالملف — اضغط أي صف لتجعله الترويسة.</p>{RawPreview}</>}
        {showSettings && prepared.skipped.length > 0 && (
          <details className="skipped">
            <summary>تجاهلت الأداة {prepared.skipped.length} صفًا غير بيانات (عناوين أقسام، ترويسة مكررة، إجماليات، ملاحظات)</summary>
            <ul>{prepared.skipped.slice(0, 40).map((s, i) => <li key={i}><b>صف {s.src}</b> — {s.reason}{s.text ? `: ${s.text}` : ''}</li>)}</ul>
          </details>
        )}
      </div>

      {mapView === 'simple' && (
        <>
          {mapping.salary_basic_salary_amount && (!constants.salary_pay_frequency || !constants.salary_basic_salary_kind) && !mapping.salary_pay_frequency && (
            <div className="banner info" style={{ display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
              <span>الراتب موجود لكن دورية الدفع ونوعه وتاريخ بدئه غير محددة.</span>
              <button className="btn small" onClick={salaryDefaults}>تطبيق الافتراضيات الشائعة</button>
            </div>
          )}
          <div className="card">
            <h2>أعمدة ملف العميل ({headers.length})</h2>
            <p className="sub">كل كرت = عمود من ملف العميل مع عيّنة من المكتوب فيه، وتحته الحقل اللي يروح له بقيود.</p>
            {ColumnCards}
          </div>
          {MissingRequired}
        </>
      )}
      {mapView === 'detailed' && (
      <div className="card">
        <div className="maphead">
          <div><h2>مطابقة الأعمدة — عرض تفصيلي</h2><p className="sub" style={{ margin: 0 }}>لكل حقل بالقالب اختر عمود العميل، أو ثبّت قيمة واحدة لكل الموظفين.</p></div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <span className="chip" style={{ background: missingRequired.length ? 'var(--error-bg)' : 'var(--ok-bg)', color: missingRequired.length ? 'var(--error)' : 'var(--ok)' }}>
              الإلزامي: {requiredKeys.length - missingRequired.length} / {requiredKeys.length}
            </span>
            <div className="meter"><i style={{ width: `${((requiredKeys.length - missingRequired.length) / Math.max(1, requiredKeys.length)) * 100}%` }} /></div>
          </div>
        </div>
        {missingRequired.length > 0 && <div className="banner warn">حقول إلزامية بدون مصدر: {missingRequired.map((k) => FIELD_BY_KEY[k].label).join('، ')}. تقدر تكمل وتعبّيها بالمراجعة، أو تثبّت لها قيمة.</div>}
        {mapping.salary_basic_salary_amount && (!constants.salary_pay_frequency || !constants.salary_basic_salary_kind) && !mapping.salary_pay_frequency && (
          <div className="banner info" style={{ display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
            <span>الراتب موجود لكن دورية الدفع ونوعه وأساس الوقت وتاريخ البدء غير محددة.</span>
            <button className="btn small" onClick={salaryDefaults}>تطبيق الافتراضيات الشائعة</button>
          </div>
        )}
        {unusedCols.length > 0 && (
          <div className="unused">
            <span className="sub" style={{ margin: 0 }}>أعمدة العميل غير المستخدمة ({unusedCols.length}):</span>
            {unusedCols.map((h) => <span key={h.col} className="chip gray" title={columnSample(prepared, h.col, 3).join(' · ')}>{h.text}</span>)}
          </div>
        )}
        <div className="maprow head"><div>حقل القالب</div><div>عمود العميل</div><div>عيّنة من البيانات</div><div>قيمة ثابتة</div></div>
        <div className="grp">{GROUPS.personal} — الاسم الكامل (يُقسَّم تلقائيًا)</div>
        {VIRTUAL_FIELDS.map(rowFor)}
        {Object.entries(GROUPS).map(([g, label]) => (
          <React.Fragment key={g}>
            <div className="grp">{label}</div>
            {FIELDS.filter((f) => f.group === g).map(rowFor)}
          </React.Fragment>
        ))}
      </div>
      )}
      <div className="actions">
        <button className="btn" onClick={() => setStep(0)}>رجوع</button>
        <div className="right"><button className="btn primary" disabled={!records.length} onClick={() => { saveProfile(); setPage(0); setStep(2); }}>التالي: المراجعة ({records.length} موظف)</button></div>
      </div>
    </>
  );

  /* --------------------------- الخطوة 3 --------------------------- */
  const um = useMemo(() => unmatchedValues(rows.filter((r) => !r.excluded)), [rows]);
  const umCount = Object.values(um).reduce((a, b) => a + b.length, 0);
  const ax = useMemo(() => approxValues(rows.filter((r) => !r.excluded)), [rows]);
  const axCount = Object.values(ax).reduce((a, b) => a + b.length, 0);
  const emptyReq = useMemo(() => {
    const m = {};
    for (const r of rows) {
      if (r.excluded) continue;
      for (const f of FIELDS) {
        const c = r.cells[f.key];
        if (c.status === 'error' && c.msg === 'حقل إلزامي فاضي') m[f.key] = (m[f.key] || 0) + 1;
      }
    }
    return Object.entries(m).filter(([, n]) => n >= 2);
  }, [rows]);
  const [fillVals, setFillVals] = useState({});
  const visibleFields = useMemo(() => {
    if (showAllCols) return FIELDS;
    return FIELDS.filter((f) => f.req || rows.some((r) => r.cells[f.key].status !== 'empty'));
  }, [rows, showAllCols]);
  const filtered = useMemo(() => {
    const q = normKey(query);
    return rows.filter((r) => {
      if (filter === 'excluded' && !r.excluded) return false;
      if (filter === 'ready' && (r.excluded || r.status === 'error')) return false;
      if (!['all', 'excluded', 'ready'].includes(filter) && (r.excluded || r.status !== filter)) return false;
      if (!q) return true;
      const hay = normKey(['first_name_ar', 'last_name_ar', 'first_name', 'last_name', 'national_id_number', 'employee_no', 'email', 'work_phone_number'].map((k) => r.cells[k].value).join(' ') + ' ' + r.src);
      return hay.includes(q);
    });
  }, [rows, filter, query]);
  const pages = Math.max(1, Math.ceil(filtered.length / PAGE));
  const curPage = Math.min(page, pages - 1);
  const pageRows = filtered.slice(curPage * PAGE, curPage * PAGE + PAGE);

  const [editing, setEditing] = useState(null); // {src,key}
  const commitEdit = (src, key, val) => {
    setEditing(null);
    if (val === undefined) return;
    setOverrides((o) => ({ ...o, [src]: { ...(o[src] || {}), [key]: val } }));
  };
  const toggleExclude = (src) => setExcluded((s) => { const n = new Set(s); n.has(src) ? n.delete(src) : n.add(src); return n; });
  const applyMap = (field, nk, val) => { if (!val) return; setUserMaps((m) => ({ ...m, [field]: { ...(m[field] || {}), [nk]: val } })); };
  const acceptAllApprox = () => {
    setUserMaps((m) => {
      const n = { ...m };
      for (const [field, list] of Object.entries(ax)) { n[field] = { ...(n[field] || {}) }; for (const e of list) n[field][e.nk] = e.value; }
      return n;
    });
    flash('اعتُمدت كل المطابقات التقريبية');
  };

  /* --------------------------- المتابعة مع العميل --------------------------- */
  const fuCount = useMemo(() => followupRows(rows, fuReview).length, [rows, fuReview]);
  const tlCount = useMemo(() => rows.filter((r) => !r.excluded && ['first_name_ar', 'last_name_ar', 'first_name', 'last_name'].some((k) => r.cells[k].translit)).length, [rows]);
  const doFollowup = async () => {
    setBusy(true);
    try {
      const bytes = await buildFollowup(rows, tplFile.tpl, { includeReview: fuReview });
      await download(bytes, `qhr_followup_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (e) { setErr(e.message); } finally { setBusy(false); }
  };
  const loadFollowup = async (file) => {
    setErr('');
    try {
      const res = parseFollowup(readClientFile(await file.arrayBuffer(), file.name), rows);
      if (!res) { setErr('هذا الملف ليس ملف استكمال من الأداة (ما فيه عمود مرجع الصف). ارفع الملف اللي نزّلته من زر "ملف استكمال للعميل" بعد ما يعبيه العميل.'); return; }
      setOverrides((o) => {
        const n = { ...o };
        for (const [src, vals] of Object.entries(res.overrides)) n[src] = { ...(n[src] || {}), ...vals };
        return n;
      });
      flash(`طُبّق ${res.changed} تعديل على ${Object.keys(res.overrides).length} موظف${res.unknownRows.length ? ` · ${res.unknownRows.length} صف غير معروف تُجوهل` : ''}`);
    } catch (e) { setErr(`تعذّر قراءة ملف الاستكمال: ${e.message}`); }
  };
  const FollowupCard = tplFile && rows.length > 0 && (
    <div className="card">
      <div className="maphead">
        <div>
          <h2>متابعة مع العميل</h2>
          <p className="sub" style={{ margin: 0 }}>
            نزّل ملف فيه فقط الموظفين والخانات الناقصة أو الخاطئة (أصفر، مع سبب كل خانة وقوائم منسدلة)، أرسله للعميل، وبعد ما يرجعه ارفعه هنا وتنطبق التعبئة تلقائيًا.
            {tlCount > 0 && ` · ${tlCount} موظف اسمه مترجم تلقائيًا ويحتاج تأكيد.`}
          </p>
        </div>
      </div>
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center', marginTop: 10 }}>
        <button className="btn primary" disabled={busy || !fuCount} onClick={doFollowup}>تنزيل ملف استكمال للعميل ({fuCount} موظف)</button>
        <FileButton label="رفع الملف بعد تعبئة العميل" accept=".xlsx" onFile={loadFollowup} />
        <label className="switch"><input type="checkbox" checked={fuReview} onChange={(e) => setFuReview(e.target.checked)} />ضمّن الخانات اللي تحتاج تأكيد (أسماء مترجمة، هوية منتهية…)</label>
      </div>
      {err && step === 2 && <div className="banner err">{err}</div>}
    </div>
  );

  const renderCell = (r, f) => {
    const c = r.cells[f.key];
    const isEdit = editing && editing.src === r.src && editing.key === f.key;
    const shown = c.value !== '' && c.value !== undefined ? String(c.value) : (c.status === 'error' ? c.raw : '');
    const display = f.type === 'date' && isIsoDate(shown) ? fmtDate(shown) : shown;
    const tip = [c.msg, c.raw && String(c.raw) !== String(c.value) ? `الأصل: ${c.raw}` : ''].filter(Boolean).join('\n');
    if (isEdit) {
      return (
        <td key={f.key}><div className="cell-edit">
          <FieldEditor f={f} cell={c} list={tplFile.tpl.lists[f.key]} autoFocus onCommit={(v) => commitEdit(r.src, f.key, v)} onCancel={() => setEditing(null)} />
        </div></td>
      );
    }
    return (
      <td key={f.key}>
        <div className={`cell ${c.status} ${c.manual ? 'manual' : ''} ${c.acked ? 'acked' : ''}`} title={c.acked ? `تمت المراجعة\n${tip}` : tip} onClick={() => setEditing({ src: r.src, key: f.key })}>
          {['fixed', 'review', 'error'].includes(c.status) && <i className="dot" />}
          <span>{display || (c.status === 'empty' ? '—' : '')}</span>
        </div>
      </td>
    );
  };

  const StepReview = tplFile && (
    <>
      <div className="stats">
        <button className={`stat ${filter === 'all' ? 'on' : ''}`} onClick={() => { setFilter('all'); setPage(0); }}><b>{counts.all}</b><span>إجمالي الصفوف</span></button>
        <button className={`stat ok ${filter === 'ready' ? 'on' : ''}`} onClick={() => { setFilter('ready'); setPage(0); }}><b>{counts.ok + counts.review}</b><span>جاهز للرفع</span></button>
        <button className={`stat review ${filter === 'review' ? 'on' : ''}`} onClick={() => { setFilter('review'); setPage(0); }}><b>{counts.review}</b><span>منها عليها ملاحظات (تنرفع عادي)</span></button>
        <button className={`stat error ${filter === 'error' ? 'on' : ''}`} onClick={() => { setFilter('error'); setPage(0); }}><b>{counts.error}</b><span>فيه أخطاء</span></button>
        <button className={`stat ${filter === 'excluded' ? 'on' : ''}`} onClick={() => { setFilter('excluded'); setPage(0); }}><b>{counts.excluded}</b><span>مستبعد</span></button>
      </div>

      {FollowupCard}

      {emptyReq.length > 0 && (
        <div className="card">
          <h2>حقول إلزامية فاضية عند أكثر من موظف</h2>
          <p className="sub">لو القيمة نفسها لكل الفاضيين (مثل بلد الإقامة) عبّها مرة وحدة هنا. الخلايا المعبأة أصلًا ما تتغير.</p>
          <div className="mapper">
            {emptyReq.map(([key, n]) => {
              const f = FIELD_BY_KEY[key];
              const list = tplFile.tpl.lists[key];
              return (
                <div className="mrow" key={key}>
                  <div><span className="chip gray">{f.label}</span></div>
                  <div className="raw"><span className="chip error">{n} صف فاضي</span></div>
                  {list?.length ? (
                    <select value={fillVals[key] ?? ''} onChange={(e) => setFillVals((v) => ({ ...v, [key]: e.target.value }))}>
                      <option value="">اختر القيمة…</option>
                      {list.map((v) => <option key={v} value={v}>{v}</option>)}
                    </select>
                  ) : f.type === 'date' ? (
                    <DateField value={fillVals[key] ?? ''} onChange={(iso) => setFillVals((v) => ({ ...v, [key]: iso }))} />
                  ) : (
                    <input type="text" placeholder={f.type === 'date' ? 'YYYY-MM-DD' : 'القيمة…'} value={fillVals[key] ?? ''} onChange={(e) => setFillVals((v) => ({ ...v, [key]: e.target.value }))} />
                  )}
                  <button className="btn small" disabled={!fillVals[key]} onClick={() => { setConst(key, fillVals[key]); flash(`عُبّئ ${f.label}`); }}>تعبئة الفاضي</button>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {umCount > 0 && (
        <div className="card">
          <h2>قيم غير موجودة بقائمة القالب</h2>
          <p className="sub">اختر القيمة الصحيحة مرة وحدة وتنطبق على كل الصفوف اللي فيها نفس القيمة.</p>
          <div className="mapper">
            {Object.entries(um).flatMap(([field, list]) => list.map((e) => (
              <div className="mrow" key={field + e.nk}>
                <div><span className="chip gray">{FIELD_BY_KEY[field].label}</span></div>
                <div className="raw">"{e.raw}" <span className="chip error">{e.count} صف</span></div>
                <select defaultValue="" onChange={(ev) => applyMap(field, e.nk, ev.target.value)}>
                  <option value="">اختر القيمة الصحيحة…</option>
                  {[...(e.suggestions || []), ...(tplFile.tpl.lists[field] || []).filter((v) => !(e.suggestions || []).includes(v))].map((v) => <option key={v} value={v}>{v}</option>)}
                </select>
                <span className="chip gray">أو أضفه كخيار في قيود ثم نزّل قالبًا جديدًا</span>
              </div>
            )))}
          </div>
        </div>
      )}

      {axCount > 0 && (
        <div className="card">
          <div className="maphead">
            <div><h2>مطابقات تقريبية تحتاج تأكيد</h2><p className="sub" style={{ margin: 0 }}>الأداة خمّنت القيمة الأقرب من قائمة القالب. اعتمدها كلها مرة وحدة، أو غيّر أي وحدة.</p></div>
            <button className="btn small primary" onClick={acceptAllApprox}>اعتماد الكل ({axCount})</button>
          </div>
          <div className="mapper">
            {Object.entries(ax).flatMap(([field, list]) => list.map((e) => (
              <div className="mrow" key={field + e.nk}>
                <div><span className="chip gray">{FIELD_BY_KEY[field].label}</span></div>
                <div className="raw">"{e.raw}" ← {e.value} <span className="chip review">{e.count} صف</span></div>
                <select value={e.value} onChange={(ev) => applyMap(field, e.nk, ev.target.value)}>
                  {(tplFile.tpl.lists[field] || []).map((v) => <option key={v} value={v}>{v}</option>)}
                </select>
                <button className="btn small" onClick={() => applyMap(field, e.nk, e.value)}>اعتماد</button>
              </div>
            )))}
          </div>
        </div>
      )}

      <div className="toolbar">
        <div style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap' }}>
          <input type="text" className="search" placeholder="بحث بالاسم أو الهوية أو الرقم الوظيفي…" value={query} onChange={(e) => { setQuery(e.target.value); setPage(0); }} />
          <span className="chip fixed">أزرق = صُحّح تلقائيًا</span>
          <span className="chip review">أصفر = راجعه</span>
          <span className="chip error">أحمر = لازم يتصلح</span>
          <span className="chip gray">اضغط أي خلية للتعديل · رقم الصف يفتح نموذج الموظف</span>
          {reviewCells > 0 && (
            <button className="btn small" title="يعتمد كل الخانات الصفراء الحالية كمراجَعة (تقدر تتراجع من نموذج الموظف)" onClick={() => {
              setReviewed((m) => {
                const n = { ...m };
                for (const r of rows) {
                  if (r.excluded) continue;
                  for (const f of FIELDS) if (r.cells[f.key].status === 'review') n[r.src] = { ...(n[r.src] || {}), [f.key]: String(r.cells[f.key].value ?? '') };
                }
                return n;
              });
              flash(`اعتُمدت ${reviewCells} خانة صفراء كمراجَعة`);
            }}>✓ تمت مراجعة كل الأصفر ({reviewCells})</button>
          )}
        </div>
        <label className="switch"><input type="checkbox" checked={showAllCols} onChange={(e) => setShowAllCols(e.target.checked)} />عرض كل الأعمدة</label>
      </div>

      <div className="tablewrap">
        <table className="dt">
          <thead>
            <tr>
              <th className="sticky c0">الصف</th>
              <th className="sticky c1">حالة</th>
              {visibleFields.map((f) => <th key={f.key}>{f.label}{f.req ? ' *' : ''}</th>)}
            </tr>
          </thead>
          <tbody>
            {pageRows.map((r) => (
              <tr key={r.src} className={r.excluded ? 'excluded' : ''}>
                <td className="sticky c0"><button className="rownum" onClick={() => setDrawerRow(r.src)} title="تفاصيل الصف">{r.src}</button></td>
                <td className="sticky c1"><div className="rowstate"><span className={`chip ${r.status}`}>{r.status === 'ok' ? 'جاهز' : r.status === 'review' ? `مراجعة ${r.reviewCount}` : `خطأ ${r.errorCount}`}</span></div></td>
                {visibleFields.map((f) => renderCell(r, f))}
              </tr>
            ))}
            {!pageRows.length && <tr><td colSpan={visibleFields.length + 2}><div className="empty">لا توجد صفوف بهذا التصفية</div></td></tr>}
          </tbody>
        </table>
      </div>
      <div className="pager">
        <button className="btn small" disabled={curPage === 0} onClick={() => setPage(curPage - 1)}>السابق</button>
        <span>صفحة {curPage + 1} من {pages} · {filtered.length} صف</span>
        <button className="btn small" disabled={curPage >= pages - 1} onClick={() => setPage(curPage + 1)}>التالي</button>
      </div>
      <div className="actions">
        <button className="btn" onClick={() => setStep(1)}>رجوع للمطابقة</button>
        <div className="right"><button className="btn primary" onClick={() => setStep(3)}>التالي: التصدير</button></div>
      </div>
    </>
  );

  /* --------------------------- الخطوة 4 --------------------------- */
  const doExport = async (force = false) => {
    // تأكيد داخل الصفحة (نوافذ confirm المنبثقة محجوبة في بعض المتصفحات والصفحات المضمّنة)
    if (tplFile.tpl.isSample && !force) { setAskSample(true); return; }
    setAskSample(false);
    setBusy(true);
    try {
      const max = tplFile.tpl.maxRows;
      const chunks = [];
      for (let i = 0; i < exportable.length; i += max) chunks.push(exportable.slice(i, i + max));
      for (let i = 0; i < chunks.length; i++) {
        const bytes = await fillTemplate(tplFile.buffer, tplFile.tpl, chunks[i]);
        await download(bytes, chunks.length > 1 ? `qhr_employees_upload_part${i + 1}.xlsx` : 'qhr_employees_upload.xlsx');
        await new Promise((r) => setTimeout(r, 400));
      }
      flash(`نُزّل ${chunks.length} ملف (${exportable.length} موظف)`);
    } catch (e) { setErr(e.message); } finally { setBusy(false); }
  };
  const doReport = async () => {
    setBusy(true);
    try {
      const bytes = await buildReport(rows, { clientFile: client.name, templateFile: tplFile.name, exported: exportable.length, skipped: prepared?.skipped || [] });
      await download(bytes, 'qhr_import_report.xlsx');
    } catch (e) { setErr(e.message); } finally { setBusy(false); }
  };
  // ملخص أسباب الأخطاء: كم صف يتأثر بكل سبب (عشان يكون واضح ليش الصفوف ما تدخل ملف الرفع)
  const errorSummary = useMemo(() => {
    const m = new Map();
    for (const r of rows) {
      if (r.excluded || r.status !== 'error') continue;
      const seen = new Set();
      for (const f of FIELDS) {
        const c = r.cells[f.key];
        if (c.status !== 'error') continue;
        const why = /غير موجود بقائمة/.test(c.msg || '') ? 'قيمة غير موجودة بقائمة القالب' : (c.msg || 'خطأ').split(' • ').pop().replace(/"[^"]*"/g, '').replace(/:\s*$/, '').trim();
        const k = `${f.label}|${why}`;
        if (seen.has(k)) continue;
        seen.add(k);
        m.set(k, { field: f.label, why, n: (m.get(k)?.n || 0) + 1 });
      }
    }
    return [...m.values()].sort((a, b) => b.n - a.n);
  }, [rows]);
  const goFixErrors = () => { setFilter('error'); setPage(0); setStep(2); };

  const StepExport = tplFile && client && (
    <>
      <div className="card">
        <h2>جاهز للتصدير</h2>
        <p className="sub">يُعبَّأ قالب قيود الرسمي نفسه بدون أي تغيير بالترويسة أو القوائم، فيصير جاهز لرفعه كما هو.</p>
        {SampleWarn}
        <div className="stats">
          <div className="stat ok"><b>{counts.ok}</b><span>جاهز</span></div>
          <div className="stat review"><b>{counts.review}</b><span>للمراجعة (تدخل الملف عادي)</span></div>
          <div className="stat error"><b>{counts.error}</b><span>أخطاء (لن تُضمَّن)</span></div>
          <div className="stat"><b>{exportable.length}</b><span>صفوف في ملف الرفع</span></div>
        </div>
        {counts.error > 0 && (
          <div className={`banner ${exportable.length ? 'warn' : 'err'}`}>
            <div style={{ display: 'flex', justifyContent: 'space-between', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
              <b>{exportable.length
                ? `${counts.error} صف فيه أخطاء ما راح يدخل بملف الرفع.`
                : `زر التنزيل معطّل لأنه ما في ولا صف جاهز: كل الصفوف (${counts.error}) فيها خطأ واحد على الأقل، وقيود يرفضها.`}</b>
              <button className="btn small primary" onClick={goFixErrors}>صلّح الأخطاء الآن</button>
            </div>
            <div style={{ marginTop: 6 }}>أكثر الأسباب:</div>
            <ul className="errsum">
              {errorSummary.slice(0, 8).map((e) => <li key={e.field + e.why}><b>{e.field}</b>: {e.why} — <span className="chip error">{e.n} صف</span></li>)}
            </ul>
            <div style={{ fontSize: 13 }}>الحقول الفاضية عند كل الموظفين تتعبّى مرة وحدة من لوحة "حقول إلزامية فاضية" بخطوة المراجعة، والقيم غير الموجودة بالقائمة من لوحة "قيم غير موجودة بقائمة القالب".</div>
          </div>
        )}
        {exportable.length > tplFile.tpl.maxRows && <div className="banner info">القالب يدعم {tplFile.tpl.maxRows} صفًا بالملف، فبينقسم لعدة ملفات.</div>}
        {err && <div className="banner err">{err}</div>}
        <div className="actions">
          <button className="btn" onClick={() => setStep(2)}>رجوع للمراجعة</button>
          <div className="right">
            <button className="btn" disabled={busy || !rows.length} onClick={doReport}>تنزيل تقرير الملاحظات</button>
            <button className="btn primary" disabled={busy || !exportable.length} onClick={() => doExport(false)}>{busy ? 'جارٍ التجهيز…' : `تنزيل ملف الرفع (${exportable.length})`}</button>
          </div>
        </div>
        {askSample && (
          <div className="banner err" style={{ display: 'flex', gap: 10, alignItems: 'center', flexWrap: 'wrap', justifyContent: 'space-between' }}>
            <span>القالب الحالي تجريبي، وقيود سيرفض الملف. تبي تنزّله للتجربة فقط؟</span>
            <span style={{ display: 'flex', gap: 8 }}>
              <button className="btn small" onClick={() => doExport(true)}>نعم، نزّله للتجربة</button>
              <button className="btn small" onClick={() => setAskSample(false)}>إلغاء</button>
            </span>
          </div>
        )}
      </div>
      {FollowupCard}
      <div className="card">
        <h2>بعد التنزيل</h2>
        <ol className="steps-howto">
          <li>افتح قيود HR ← الموظفين ← إضافة جماعية.</li>
          <li>ارفع الملف الذي نزّلته من هنا (نفس القالب، فقط معبّأ).</li>
          <li>الموظفون يُنشأون بالخلفية، راجع قائمة الموظفين وتنبيه "بيانات ناقصة" بعدها.</li>
          <li>تقرير الملاحظات يوضح كل ما صُحّح وكل ما استُبعد وسببه، سلّمه للعميل للمتابعة.</li>
        </ol>
      </div>
    </>
  );

  /* --------------------------- درج تفاصيل الصف --------------------------- */
  const dRow = rows.find((r) => r.src === drawerRow);
  const Drawer = dRow && (
    <>
      <div className="drawer-bg" onClick={() => setDrawerRow(null)} />
      <aside className="drawer">
        <h3>الصف {dRow.src} في ملف العميل</h3>
        <p className="sub">{dRow.cells.first_name_ar.value} {dRow.cells.last_name_ar.value} · {dRow.cells.first_name.value} {dRow.cells.last_name.value}</p>
        <div style={{ display: 'flex', gap: 14, flexWrap: 'wrap', alignItems: 'center' }}>
          <label className="switch"><input type="checkbox" checked={!!dRow.excluded} onChange={() => toggleExclude(dRow.src)} />استبعد من التصدير</label>
          <div className="seg">
            <button className={!drawerAll ? 'on' : ''} onClick={() => setDrawerAll(false)}>الملاحظات ({FIELDS.filter((f) => ['error', 'review'].includes(dRow.cells[f.key].status)).length})</button>
            <button className={drawerAll ? 'on' : ''} onClick={() => setDrawerAll(true)}>كل الحقول</button>
          </div>
        </div>
        <p className="sub" style={{ margin: '10px 0 0' }}>عدّل مباشرة: التواريخ بالتقويم، القوائم بالاختيار، والباقي بالكتابة. التعديل ينحفظ فورًا. الأصفر تقدر تعتمده بـ "تمت المراجعة".</p>
        {(() => {
          const pending = FIELDS.filter((f) => dRow.cells[f.key].status === 'review');
          return pending.length > 0 && (
            <button className="btn small primary" style={{ marginTop: 10 }} onClick={() => markReviewed(dRow, pending.map((f) => f.key))}>✓ تمت مراجعة كل ملاحظات هذا الموظف ({pending.length})</button>
          );
        })()}
        {FIELDS.filter((f) => drawerAll || ['error', 'review', 'fixed'].includes(dRow.cells[f.key].status) || dRow.cells[f.key].manual || dRow.cells[f.key].acked)
          .sort((a, b) => (drawerAll ? 0 : ({ error: 0, review: 1, fixed: 2, ok: 3, empty: 3 }[dRow.cells[a.key].status] - { error: 0, review: 1, fixed: 2, ok: 3, empty: 3 }[dRow.cells[b.key].status])))
          .map((f) => {
            const c = dRow.cells[f.key];
            const NOTE_TAG = { fixed: 'صُحّح تلقائيًا', review: 'للمراجعة', error: 'خطأ' };
            return (
              <div key={f.key} className={`issue edit ${c.status}`}>
                <div className="issue-top">
                  <b>{f.label}{tplFile.tpl.required.has(f.key) ? ' *' : ''}</b>
                  {c.status !== 'ok' && c.status !== 'empty' && <span className={`chip ${c.status}`}>{STATUS_LABEL[c.status]}</span>}
                  {c.acked && <span className="chip ok">✓ تمت المراجعة</span>}
                  {c.manual && <span className="chip gray">معدّل يدويًا</span>}
                </div>
                <FieldEditor inline f={f} cell={c} list={tplFile.tpl.lists[f.key]} onCommit={(v) => commitEdit(dRow.src, f.key, v)} />
                {(c.notes || []).length > 0 && (
                  <ul className="notes">
                    {c.notes.map((n, i) => (
                      <li key={i} className={c.acked && n.status === 'review' ? 'ok' : n.status}>
                        <span className="tag">{c.acked && n.status === 'review' ? 'روجعت' : NOTE_TAG[n.status] || ''}</span>
                        <span>{n.msg}</span>
                        {n.ref !== undefined && <button className="btn small" onClick={() => setDrawerRow(n.ref)}>افتح الصف {n.ref}</button>}
                      </li>
                    ))}
                  </ul>
                )}
                {c.status === 'review' && <div><button className="btn small" onClick={() => markReviewed(dRow, [f.key])}>✓ تمت المراجعة</button></div>}
                {c.acked && <div><button className="btn small ghost" onClick={() => unmarkReviewed(dRow.src, f.key)}>تراجع عن المراجعة</button></div>}
                {c.raw && String(c.raw) !== String(c.value) && <div className="diff">الأصل بملف العميل: {c.raw}</div>}
              </div>
            );
          })}
        <div className="actions"><span /><button className="btn" onClick={() => setDrawerRow(null)}>إغلاق</button></div>
      </aside>
    </>
  );

  const stepOK = [true, !!tplFile && !!client, !!records.length, !!records.length];
  const STEP_SUB = [
    'ارفع قالب قيود الرسمي وملف العميل كما هو',
    'وين يروح كل عمود من ملف العميل بقالب قيود',
    'صحّح الأخطاء واعتمد الملاحظات قبل الرفع',
    'نزّل ملف الرفع الجاهز لقيود والتقارير',
  ];
  const STEP_ICON = [ICONS.upload, 'M4 6h7M4 12h7M4 18h7M15 6l5 6-5 6', ICONS.check, ICONS.down];
  return (
    <div className="shell">
      <aside className="side">
        <div className="side-logo"><span className="mark" dir="ltr"><span className="q">Q</span><span className="hr">HR</span></span><small>مُجهّز بيانات الموظفين</small></div>
        <nav className="side-nav">
          {STEPS.map((s, i) => (
            <button key={s} className={`side-item ${i === step ? 'on' : ''} ${i < step ? 'done' : ''}`} disabled={i > step && !stepOK[i - 1]} onClick={() => setStep(i)}>
              <Icon d={i < step ? ICONS.check : STEP_ICON[i]} size={18} />
              <span>{s}</span>
              {i === 2 && rows.length > 0 && counts.error > 0 && <em className="badge">{counts.error}</em>}
            </button>
          ))}
        </nav>
        <div className="side-note">
          <b>🔒 بياناتك تبقى عندك</b>
          كل المعالجة داخل متصفحك، ولا يُرفع أي ملف لأي مكان.
        </div>
      </aside>
      <main className="main">
        <header className="topbar">
          <div>
            <h1>{STEPS[step]}</h1>
            <p>{STEP_SUB[step]}</p>
          </div>
          <div className="topbar-meta">
            {tplFile && <span className="pill" title={tplFile.name}>القالب: {tplFile.name}</span>}
            {client && <span className="pill" title={client.name}>العميل: {client.name}</span>}
            {rows.length > 0 && <span className={`pill ${counts.error ? 'warn' : 'ok'}`}>{exportable.length} / {rows.length} جاهز للرفع</span>}
          </div>
        </header>
        <div className="content">
          {step === 0 && StepFiles}
          {step === 1 && StepMap}
          {step === 2 && StepReview}
          {step === 3 && StepExport}
        </div>
      </main>
      {Drawer}
      {toast && <div className="toast">{toast}</div>}
    </div>
  );
}
