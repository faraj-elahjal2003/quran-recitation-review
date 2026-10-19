'use strict';
/* ===== ثوابت مأخوذة من قالب Recitation Annotation.xlsx (ورقة Annotation) ===== */
const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const AHKAM = ['مد طبيعي', 'مد عارض للسكون', 'إدغام المثلين', 'إدغام النون الساكنة والتنوين بغير غنة', 'التفخيم'];
const FIXED = ['Record_ID', 'File_Name', 'Extension', 'Reciter_ID', 'Name_Check'];
const EDIT = [...AHKAM, 'Audio_Status', 'Error_Type', 'Review_Status', 'Notes'];
const DONE = 'تمت المراجعة', UNSEEN = 'لم يُراجع', WIP = 'قيد المراجعة', SECOND = 'يحتاج إلى مراجعة ثانية', VALID = 'صالح للتحكيم';
// قيم احتياطية إن لم تُقرأ القوائم من الملف نفسه
const DEFOPTS = {
  ahkam: ['نعم', 'لا', 'غير واضح', 'لم يُقيّم'],
  Audio_Status: ['صالح للتحكيم', 'ملغي', 'ملف تالف', 'تسجيل غير واضح', 'تسجيل مكرر', 'ملف فارغ', 'يحتاج إلى مراجعة'],
  Error_Type: ['المدود', 'الإدغام', 'التفخيم والترقيق', 'مخارج الحروف', 'صفات الحروف', 'الوقف والابتداء', 'أحكام أخرى', 'لا يوجد خطأ محدد'],
  Review_Status: [UNSEEN, WIP, DONE, SECOND]
};
const S = { recs: [], idx: 0, opts: {}, audio: new Map(), buf: null, srcName: '', reviewer: '', dirty: 0, lastExport: null, idbOK: true, restored: 0 };
const $ = id => document.getElementById(id);
const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* ===== أدوات الأعمدة ===== */
const colNum = L => [...L].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const colOf = ref => ref.replace(/\d+/g, '');
const txt = el => [...el.getElementsByTagName('t')].map(t => t.textContent).join('');

/* ===== قراءة ملف xlsx (بدون أي خادم) ===== */
async function parseWorkbook(buf) {
  const zip = await JSZip.loadAsync(buf);
  const P = new DOMParser();
  const readXml = async p => { const f = zip.file(p); if (!f) return null; return P.parseFromString(await f.async('string'), 'application/xml'); };
  const wb = await readXml('xl/workbook.xml');
  if (!wb) throw new Error('هذا ليس ملف Excel (.xlsx) صالحًا.');
  const sheets = [...wb.getElementsByTagName('sheet')];
  const sh = sheets.find(s => s.getAttribute('name') === 'Annotation') || sheets[0];
  const rid = sh.getAttributeNS(RNS, 'id') || sh.getAttribute('r:id');
  const rels = await readXml('xl/_rels/workbook.xml.rels');
  const rel = [...rels.getElementsByTagName('Relationship')].find(r => r.getAttribute('Id') === rid);
  const tg = rel.getAttribute('Target');
  const sheetPath = tg.startsWith('/') ? tg.slice(1) : 'xl/' + tg;
  const sst = await readXml('xl/sharedStrings.xml');
  const ss = sst ? [...sst.getElementsByTagName('si')].map(si => {
    // تجاهل النص الصوتي rPh
    const c = si.cloneNode(true); [...c.getElementsByTagName('rPh')].forEach(n => n.remove()); return txt(c);
  }) : [];
  const doc = await readXml(sheetPath);
  const val = c => {
    const t = c.getAttribute('t'), v = c.getElementsByTagName('v')[0];
    if (t === 'inlineStr') return txt(c);
    if (!v) return '';
    return t === 's' ? (ss[+v.textContent] || '') : v.textContent;
  };
  const rows = [...doc.getElementsByTagName('row')];
  const hdrRow = rows.find(r => r.getAttribute('r') === '1');
  if (!hdrRow) throw new Error('لم أجد صف العناوين في الورقة.');
  const cols = {};
  [...hdrRow.getElementsByTagName('c')].forEach(c => { const h = val(c).trim(); if (h) cols[h] = colOf(c.getAttribute('r')); });
  const missing = ['Record_ID', 'File_Name', ...EDIT].filter(h => !(h in cols));
  if (missing.length) throw new Error('هذا الملف لا يطابق قالب Recitation Annotation. أعمدة ناقصة: ' + missing.join('، '));
  // قوائم التحقق (Data Validation) من القالب نفسه
  const byCol = {};
  [...doc.getElementsByTagName('dataValidation')].forEach(d => {
    const f = d.getElementsByTagName('formula1')[0];
    if (d.getAttribute('type') === 'list' && f && /^"/.test(f.textContent.trim()))
      byCol[colOf((d.getAttribute('sqref') || '').split(/[ :]/)[0])] = f.textContent.trim().replace(/^"|"$/g, '').split(',');
  });
  const opts = {};
  EDIT.forEach(h => { const o = byCol[cols[h]]; if (o && o.length) opts[h] = o; });
  AHKAM.forEach(h => { if (!opts[h]) opts[h] = DEFOPTS.ahkam; });
  ['Audio_Status', 'Error_Type', 'Review_Status'].forEach(h => { if (!opts[h]) opts[h] = DEFOPTS[h]; });
  const recs = [];
  rows.forEach(r => {
    const n = +r.getAttribute('r'); if (n < 2) return;
    const m = {}; [...r.getElementsByTagName('c')].forEach(c => { m[colOf(c.getAttribute('r'))] = val(c); });
    const get = h => (m[cols[h]] || '').trim();
    if (!get('Record_ID') && !get('File_Name')) return;
    const f = {}; FIXED.forEach(h => { f[h] = h in cols ? get(h) : ''; });
    const v = {}; EDIT.forEach(h => { v[h] = h === 'Notes' ? (m[cols[h]] || '') : get(h); });
    recs.push({ row: n, id: f.Record_ID || ('ROW' + n), f, v });
  });
  if (!recs.length) throw new Error('لا توجد تسجيلات في الملف.');
  return { cols, opts, recs, sheetPath };
}

/* ===== تصدير: تعديل الملف الأصلي مباشرة فتبقى القوائم والتنسيق والمعادلات والورقات الأخرى ===== */
async function buildXlsx(buf, sheetPath, cols, recs, onlyRows) {
  const zip = await JSZip.loadAsync(buf);
  const P = new DOMParser(), X = new XMLSerializer();
  const ser = d => { let s = X.serializeToString(d); if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s; return s; };
  const doc = P.parseFromString(await zip.file(sheetPath).async('string'), 'application/xml');
  const rowEl = {}; [...doc.getElementsByTagName('row')].forEach(r => { rowEl[r.getAttribute('r')] = r; });
  const setCell = (row, ref, value) => {
    let c = [...row.children].find(x => x.getAttribute('r') === ref);
    if (!c) {
      c = doc.createElementNS(NS, 'c'); c.setAttribute('r', ref);
      const after = [...row.children].find(x => colNum(colOf(x.getAttribute('r'))) > colNum(colOf(ref)));
      row.insertBefore(c, after || null);
    }
    while (c.firstChild) c.removeChild(c.firstChild);
    c.removeAttribute('t');
    if (value !== '' && value != null) {
      c.setAttribute('t', 'inlineStr');
      const is = doc.createElementNS(NS, 'is'), t = doc.createElementNS(NS, 't');
      t.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
      t.textContent = String(value); is.appendChild(t); c.appendChild(is);
    }
  };
  if (onlyRows) {
    Object.keys(rowEl).forEach(k => { if (+k > 1 && !onlyRows.has(+k)) rowEl[k].remove(); });
    const dm = doc.getElementsByTagName('dimension')[0]; if (dm) dm.setAttribute('ref', 'A1:N' + Math.max(...onlyRows));
  }
  recs.forEach(r => { const row = rowEl[r.row]; if (row) EDIT.forEach(h => setCell(row, cols[h] + r.row, r.v[h])); });
  zip.file(sheetPath, ser(doc));
  // إعادة حساب ورقة Summary عند فتح الملف في Excel
  const wbf = zip.file('xl/workbook.xml'), wdoc = P.parseFromString(await wbf.async('string'), 'application/xml');
  let cp = wdoc.getElementsByTagName('calcPr')[0];
  if (!cp) { cp = wdoc.createElementNS(NS, 'calcPr'); wdoc.documentElement.appendChild(cp); }
  cp.setAttribute('fullCalcOnLoad', '1');
  zip.file('xl/workbook.xml', ser(wdoc));
  return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', compression: 'DEFLATE' });
}
if (typeof window !== 'undefined') window.TQ = { parseWorkbook, buildXlsx };

/* ===== التخزين المحلي IndexedDB ===== */
let db = null;
const openDB = () => new Promise((res, rej) => {
  const r = indexedDB.open('tilawa-annotation', 1);
  r.onupgradeneeded = () => { r.result.createObjectStore('records', { keyPath: 'id' }); r.result.createObjectStore('meta'); };
  r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
});
const tx = (st, mode, fn) => new Promise((res, rej) => {
  const t = db.transaction(st, mode), q = fn(t.objectStore(st));
  t.oncomplete = () => res(q && q.result); t.onerror = t.onabort = () => rej(t.error);
});
async function persist(r) {
  if (!db) return;
  try { await tx('records', 'readwrite', s => s.put({ id: r.id, v: r.v, at: Date.now() })); }
  catch (e) { S.idbOK = false; renderWarn(); }
}
const meta = (k, v) => db ? tx('meta', 'readwrite', s => s.put(v, k)).catch(() => {}) : null;
const getMeta = k => db ? tx('meta', 'readonly', s => s.get(k)).catch(() => null) : null;

/* ===== واجهة المستخدم ===== */
const cur = () => S.recs[S.idx];
const isDone = r => r.v.Review_Status === DONE;
const isFlag = r => r.v.Audio_Status && r.v.Audio_Status !== VALID;
function toast(m) {
  const t = document.createElement('div'); t.className = 'toast'; t.textContent = m; document.body.appendChild(t); setTimeout(() => t.remove(), 3200);
}

async function init() {
  try {
    const bin = atob(EMBED), buf = new Uint8Array(bin.length);
    for (let k = 0; k < bin.length; k++) buf[k] = bin.charCodeAt(k);
    const p = await parseWorkbook(buf.buffer);
    const keep = new Set(BATCH.ids);
    S.buf = buf.buffer; S.reviewer = String(BATCH.n); S.opts = p.opts; S.cols = p.cols; S.sheetPath = p.sheetPath;
    S.recs = p.recs.filter(r => keep.has(r.id));
    if (db) {
      const saved = await tx('records', 'readonly', s => s.getAll()).catch(() => []);
      const m = new Map(saved.map(x => [x.id, x]));
      S.recs.forEach(r => { const x = m.get(r.id); if (x) { EDIT.forEach(h => { if (x.v[h] != null) r.v[h] = x.v[h]; }); S.restored++; } });
    }
    S.lastExport = (await getMeta('lastExport')) || null;
    S.idx = Math.min((await getMeta('lastIdx')) || 0, S.recs.length - 1);
    try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch (e) { }
    $('foot').classList.remove('hidden');
    renderAll();
  } catch (e) { $('head').innerHTML = '<div class="banner e">تعذّر تحميل البيانات: ' + esc(e.message || e) + '</div>'; }
}

function renderWarn() {
  const w = [];
  if (!S.idbOK || !db) w.push('<div class="banner e"><b>التخزين المحلي غير متاح</b> في هذا المتصفح (ربما وضع التصفح الخاص). لن تُحفظ تقييماتك عند الإغلاق؛ نزّل Excel بعد كل جلسة.</div>');
  const when = S.lastExport ? new Date(S.lastExport).toLocaleString('ar') : 'لم يُنزَّل بعد';
  w.push(`<div class="banner ${S.dirty > 0 ? 'w' : 'i'}">التقييمات محفوظة على هذا الجهاز فقط وليست داخل ملف Excel تلقائيًا. آخر تنزيل لـ Excel: <b>${esc(when)}</b>${S.dirty > 0 ? ` — ${S.dirty} تعديل منذ ذلك الحين، نزّل نسخة احتياطية.` : ''}</div>`);
  $('warn').innerHTML = w.join('');
}
function renderProg() {
  const n = S.recs.length, d = S.recs.filter(isDone).length;
  $('progTxt').textContent = `مكتمل ${d} • متبقٍ ${n - d} • من ${n}`;
  $('posTxt').textContent = `${S.idx + 1} / ${n}`;
  $('progBar').style.width = (d / n * 100) + '%';
}
function renderAll() { S.ddOpen = false; renderProg(); renderWarn(); renderHead(); renderAudio(); renderForm(); window.scrollTo(0, 0); meta('lastIdx', S.idx); }
function renderHead() {
  const r = cur(), f = r.f;
  $('head').innerHTML = `<div class="meta"><span>رقم التسجيل: <b>${esc(f.Record_ID)}</b></span><span>القارئ: <b>${esc(f.Reciter_ID || '—')}</b></span><span>الامتداد: <b>${esc(f.Extension)}</b></span></div>
  <div class="fname">${esc(f.File_Name)}</div>
  ${f.Name_Check && f.Name_Check !== 'مطابق للنمط' ? `<div class="banner w">تنبيه في اسم الملف: ${esc(f.Name_Check)}</div>` : ''}`;
}

let objUrl = null;
function renderAudio() {
  const r = cur(), box = $('audioBox');
  if (objUrl) { URL.revokeObjectURL(objUrl); objUrl = null; }
  const file = S.audio.get(r.f.File_Name.toLowerCase());
  const pick = '<button class="btn full" id="aPick" style="margin-top:8px">تحميل ملف الصوت (اختياري)</button>';
  if (!file) {
    box.innerHTML = `<p class="hint" style="margin:0">يمكنك الاستماع من واتساب أو من جهازك مباشرة وإكمال التقييم هنا، أو تحميل الملف ليُشغَّل داخل التطبيق مع الموجة الصوتية. الملف المطلوب: <b dir="ltr">${esc(r.f.File_Name)}</b></p>${pick}`;
    $('aPick').onclick = () => $('pickA').click(); return;
  }
  objUrl = URL.createObjectURL(file);
  box.innerHTML = `<canvas id="wave" style="width:100%;height:80px;border-radius:10px;background:var(--bg);touch-action:none;cursor:pointer"></canvas>
  <audio id="au" controls preload="metadata" src="${objUrl}"></audio>
  <div class="row"><button class="btn" id="aBack">⏪ ١٠ ث</button><button class="btn" id="aFwd">١٠ ث ⏩</button>
  <select id="aRate" aria-label="السرعة"><option value="0.75">0.75×</option><option value="1" selected>1×</option><option value="1.25">1.25×</option></select></div>
  <div id="aErr"></div>${pick}`;
  $('aPick').onclick = () => $('pickA').click();
  const au = $('au');
  $('aBack').onclick = () => { au.currentTime = Math.max(0, au.currentTime - 10); };
  $('aFwd').onclick = () => { au.currentTime = Math.min(au.duration || 1e9, au.currentTime + 10); };
  $('aRate').onchange = e => { au.playbackRate = +e.target.value; };
  au.onerror = () => { $('wave') && $('wave').remove(); $('aErr').innerHTML = `<div class="banner w">يتعذّر تشغيل صيغة <b>${esc(r.f.Extension)}</b> في هذا المتصفح. افتح الملف من واتساب أو بمشغّل الجهاز ثم ارجع هنا لإكمال التقييم.</div>`; };
  drawWave(file, au);
}
async function drawWave(file, au) {
  const cv = $('wave'); if (!cv) return;
  const N = 160; let peaks;
  try {
    const AC = window.AudioContext || window.webkitAudioContext, ctx = new AC();
    const ab = await ctx.decodeAudioData(await file.arrayBuffer()); if (ctx.close) ctx.close();
    const d = ab.getChannelData(0), step = Math.max(1, Math.floor(d.length / N)); peaks = [];
    for (let i = 0; i < N; i++) { let m = 0; for (let k = i * step; k < (i + 1) * step && k < d.length; k += Math.max(1, Math.floor(step / 60))) m = Math.max(m, Math.abs(d[k])); peaks.push(m); }
    const mx = Math.max(...peaks) || 1; peaks = peaks.map(p => p / mx);
  } catch (e) { cv.remove(); return; }   // صيغة غير مدعومة: نخفي الموجة فقط
  if (!document.body.contains(cv)) return;
  const draw = () => {
    const dpr = window.devicePixelRatio || 1, w = cv.clientWidth, h = 80;
    cv.width = w * dpr; cv.height = h * dpr; const g = cv.getContext('2d'); g.scale(dpr, dpr);
    const p = au.duration ? au.currentTime / au.duration : 0, bw = w / N;
    peaks.forEach((v, i) => { g.fillStyle = i / N < p ? '#2f6b4a' : '#9bb8a8'; const bh = Math.max(2, v * h * .9); g.fillRect(i * bw + 1, (h - bh) / 2, Math.max(1, bw - 2), bh); });
  };
  const seek = e => { const t = e.touches ? e.touches[0] : e, b = cv.getBoundingClientRect(); if (au.duration) { au.currentTime = Math.min(1, Math.max(0, (t.clientX - b.left) / b.width)) * au.duration; draw(); } };
  cv.onclick = seek;
  ['timeupdate', 'loadedmetadata', 'seeked'].forEach(ev => au.addEventListener(ev, draw));
  window.addEventListener('resize', draw); draw();
}

function group(f, label, hint, multi) {
  const v = cur().v[f], sel = multi ? v.split(/[;؛]\s*/).filter(Boolean) : [v];
  const opts = [...S.opts[f]]; sel.forEach(x => { if (x && !opts.includes(x)) opts.push(x); });
  return `<section class="card"><h3>${esc(label)}</h3>${hint ? `<p class="hint">${hint}</p>` : ''}
  <div class="choices" data-f="${esc(f)}" ${multi ? 'data-multi="1"' : ''}>${opts.map(o => `<button type="button" class="ch${sel.includes(o) ? ' on' : ''}${f === 'Audio_Status' && o !== VALID ? ' neg' : ''}" data-v="${esc(o)}">${esc(o)}</button>`).join('')}</div></section>`;
}
function selectGroup(f, label, hint) {
  const v = cur().v[f], opts = [...S.opts[f]]; if (v && !opts.includes(v)) opts.push(v);
  return `<section class="card"><h3>${esc(label)}</h3>${hint ? `<p class="hint">${hint}</p>` : ''}<select data-f="${esc(f)}"><option value="">— اختر —</option>${opts.map(o => `<option value="${esc(o)}"${o === v ? ' selected' : ''}>${esc(o)}</option>`).join('')}</select></section>`;
}
function multiGroup(f, label, hint) {
  const sel = cur().v[f].split(/[;؛]\s*/).filter(Boolean), opts = [...S.opts[f]]; sel.forEach(x => { if (!opts.includes(x)) opts.push(x); });
  return `<section class="card"><h3>${esc(label)}</h3>${hint ? `<p class="hint">${hint}</p>` : ''}<details class="dd"${S.ddOpen ? ' open' : ''}><summary>${sel.length ? esc(sel.join('، ')) : '— اختر —'}</summary>${opts.map(o => `<label class="ck"><input type="checkbox" data-f="${esc(f)}" data-v="${esc(o)}"${sel.includes(o) ? ' checked' : ''}> ${esc(o)}</label>`).join('')}</details></section>`;
}
function renderForm() {
  const r = cur();
  $('form').innerHTML =
    `<div class="banner i">لكل حكم: <b>هل ظهر في التسجيل؟</b> (وجود الحكم ليس خطأً). أما صحة التطبيق فتُسجَّل في «نوع الخطأ».</div>` +
    selectGroup('Audio_Status', 'حالة الملف الصوتي', 'مطلوبة قبل الحفظ. لا تُلغِ تسجيلًا قبل الاستماع إليه.') +
    AHKAM.map(a => group(a, a, '')).join('') +
    multiGroup('Error_Type', 'نوع الخطأ (إن وُجد)', 'يمكن اختيار أكثر من نوع.') +
    group('Review_Status', 'حالة المراجعة', 'تتحول إلى «تمت المراجعة» عند الضغط على «حفظ التقييم».') +
    `<section class="card"><h3>ملاحظات</h3><p class="hint">سبب الإلغاء أو وصف الخطأ (الحكم والموضع) أو أي ملاحظة.</p><textarea id="notes">${esc(r.v.Notes)}</textarea></section>`;
  $('notes').oninput = e => setField('Notes', e.target.value, true);
}
function setField(f, val, quiet) {
  const r = cur();
  r.v[f] = val;
  if (f !== 'Review_Status' && r.v.Review_Status === UNSEEN) r.v.Review_Status = WIP;
  S.dirty++; persist(r); renderProg(); renderWarn();
  if (!quiet) renderForm();
}
$('form').addEventListener('change', e => {
  const t = e.target, f = t.dataset.f; if (!f) return;
  if (t.tagName === 'SELECT') return setField(f, t.value);
  if (t.type === 'checkbox') {
    const vals = [...document.querySelectorAll('#form input[type=checkbox][data-f="' + f + '"]:checked')].map(x => x.dataset.v);
    S.ddOpen = true; setField(f, vals.join('؛'));
  }
});
$('form').addEventListener('click', e => {
  const b = e.target.closest('.ch'); if (!b) return;
  const g = b.parentElement, f = g.dataset.f, v = b.dataset.v, cv = cur().v[f];
  if (g.dataset.multi) {
    const a = cv.split(/[;؛]\s*/).filter(Boolean), i = a.indexOf(v);
    i < 0 ? a.push(v) : a.splice(i, 1);
    setField(f, a.join('؛'));
  } else setField(f, cv === v ? '' : v);   // الضغط مرة ثانية يمسح الاختيار (فارغ = لم يُدخَل)
});

function go(i) { S.idx = Math.max(0, Math.min(S.recs.length - 1, i)); renderAll(); }
function save() {
  const r = cur(), v = r.v;
  if (!v.Audio_Status) { toast('اختر «حالة الملف الصوتي» أولًا'); return; }
  if (v.Audio_Status === VALID) {
    const miss = AHKAM.filter(a => !v[a]);
    if (miss.length) { toast('أكمل الأحكام (أو اختر «لم يُقيّم»): ' + miss.join('، ')); return; }
  }
  if (v.Review_Status !== SECOND) v.Review_Status = DONE;
  S.dirty++; persist(r); toast('تم حفظ التقييم على الجهاز');
  if (S.idx < S.recs.length - 1) go(S.idx + 1); else { renderProg(); renderWarn(); renderForm(); }
}

async function doExport() {
  try {
    const blob = await buildXlsx(S.buf, S.sheetPath, S.cols, S.recs, new Set(S.recs.map(r => r.row)));
    const nm = 'Recitation_Annotation_Reviewer_' + S.reviewer + '.xlsx';
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nm; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
    S.dirty = 0; S.lastExport = Date.now(); await meta('lastExport', S.lastExport); renderWarn();
    toast('تم تنزيل ' + nm + ' — أرسله عبر واتساب');
  } catch (e) { toast('تعذّر إنشاء الملف: ' + (e.message || e)); }
}

function addAudio(files) {
  let ok = 0; const names = new Set(S.recs.map(r => r.f.File_Name.toLowerCase())), un = [];
  [...files].forEach(f => { const k = f.name.toLowerCase(); S.audio.set(k, f); names.has(k) ? ok++ : un.push(f.name); });
  toast(`رُبط ${ok} ملف صوتي` + (un.length ? ` — ${un.length} ملف لا يطابق أي تسجيل` : ''));
  renderAudio();
}

function openList() {
  const q = $('q').value.trim().toLowerCase(), fl = $('flt').value;
  const items = S.recs.map((r, i) => ({ r, i })).filter(({ r }) => {
    if (fl === 'todo' && isDone(r)) return false; if (fl === 'done' && !isDone(r)) return false; if (fl === 'flag' && !isFlag(r)) return false;
    return !q || (r.id + ' ' + r.f.File_Name + ' ' + r.f.Reciter_ID).toLowerCase().includes(q);
  });
  $('items').innerHTML = items.map(({ r, i }) => `<button class="li${i === S.idx ? ' cur' : ''}" data-i="${i}"><span class="dot ${isDone(r) ? 'd' : r.v.Review_Status === WIP ? 'p' : ''}${isFlag(r) ? ' x' : ''}"></span><span dir="ltr">${esc(r.f.File_Name)}</span><span style="margin-inline-start:auto;color:var(--mute)">${esc(r.id)}</span></button>`).join('') || '<p class="hint">لا نتائج.</p>';
}

/* ===== ربط الأحداث ===== */
if (typeof document !== 'undefined' && $('app')) {
  $('pickA').onchange = e => { addAudio(e.target.files); e.target.value = ''; };
  $('bAudio').onclick = () => $('pickA').click();
  $('bPrev').onclick = () => go(S.idx - 1); $('bNext').onclick = () => go(S.idx + 1); $('bSave').onclick = save;
  $('bExport').onclick = doExport;
  $('bList').onclick = () => { openList(); $('dlg').showModal(); };
  $('bClose').onclick = () => $('dlg').close();
  $('q').oninput = $('flt').onchange = openList;
  $('items').onclick = e => { const b = e.target.closest('.li'); if (b) { $('dlg').close(); go(+b.dataset.i); } };
  (async () => {
    try { db = await openDB(); }
    catch (e) { S.idbOK = false; db = null; }
    init();
  })();
}
