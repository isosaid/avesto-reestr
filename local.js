// Локальный режим (без базы): разбор Excel в браузере + хранение в IndexedDB этого браузера.
// Ничего не отправляется на сервер. Алгоритм разбора = scripts/parse_excel.py.

/* ================= Разбор Excel ================= */
const W = '[\\p{L}\\p{N}_]';
const clean = v => {
  if (v == null) return null;
  let s = String(v).replace(/[ \t ]+/g, ' ').trim().replace(/\s*\n\s*/g, '\n');
  return ['', '-', '—'].includes(s) ? null : s;
};
const oneLine = v => { const s = clean(v); return s ? s.replace(/\s+/g, ' ') : null; };
const num = v => { if (v == null) return null; if (typeof v === 'number') return v; const d = String(v).replace(/\D/g, ''); return d ? Number(d) : null; };
const dmy = s => { const [d, m, y] = s.split('.'); return `${y}-${String(+m).padStart(2, '0')}-${String(+d).padStart(2, '0')}`; };
const pct = s => Number(String(s).replace(',', '.'));
export function qn(s) {
  if (!s) return s;
  s = s.replace(/[“”«»„]/g, '"');
  s = s.replace(/"\s+([^"]*?)\s*"/g, '"$1"');
  s = s.replace(/"([^"]*?)\s+"/g, '"$1"');
  return s.replace(/\s+/g, ' ').trim();
}
function dates(v, XLSX) {
  if (v == null) return [null, null];
  if (typeof v === 'number') { const p = XLSX.SSF.parse_date_code(v); return [`${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`, null]; }
  const f = String(v).match(/\d{1,2}\.\d{1,2}\.\d{4}/g) || [];
  return [f[0] ? dmy(f[0]) : null, f[1] ? dmy(f[1]) : null];
}
function splitItems(text) {
  if (!text) return [];
  const t = text.replace(/\s+/g, ' ').trim();
  const parts = /^\s*1\./.test(t) ? t.split(/(?:^|[;.]?\s)(?=\d{1,2}\.\s?\D)/u) : [t];
  const out = [];
  for (const p of parts) for (let q of p.split(/;\s*(?=\S)/)) {
    q = q.replace(/^\d{1,2}\.\s*/, '').replace(/^[ ;.]+|[ ;.]+$/g, '');
    if (q) out.push(q);
  }
  return out;
}
function parseFounders(text) {
  const res = [];
  for (const it of splitItems(text)) {
    let name, share;
    const m = it.match(/^(.*?)\s*[-–]\s*(\d+(?:[,.]\d+)?)\s*%/);
    if (m) { name = m[1].trim(); share = pct(m[2]); }
    else { const m2 = it.match(/\(дорандаи\s*(\d+(?:[,.]\d+)?)%/); if (!m2) continue; share = pct(m2[1]); name = it.slice(0, m2.index).trim(); }
    const nom = it.match(new RegExp(`но${W}*налии\\s*([\\d\\s]+?)\\s*(?:\\(|сомонӣ|$)`, 'u'));
    name = qn(name.replace(/^шаҳрванди\s+(ҶТ|Покистон)\s+/, ''));
    res.push({ name, share, nominal: nom ? num(nom[1]) : null });
  }
  return res;
}
function parseBenef(text) {
  const res = [];
  for (const it of splitItems(text)) {
    const m = it.match(/^(?:шаҳрванди\s+(\S+)\s+)?(.*?)\s*\(дорандаи\s*(\d+(?:[,.]\d+)?)%/);
    if (m) res.push({ name: m[2].trim(), citizenship: m[1] ?? null, share: pct(m[3]) });
  }
  return res;
}
function parseExec(text) {
  const t = oneLine(text);
  if (!t) return {};
  const m = t.match(/^(.*?)\s*[-–]\s*(.+?)(?=\s*(почтаи|номери|муовини|директори тиҷоратӣ|иҷрокунандаи|$))/i);
  const email = t.match(new RegExp(`[${W.slice(1, -1)}.\\-]+@[${W.slice(1, -1)}\\-]+(?:\\.[${W.slice(1, -1)}\\-]+)+`, 'u'));
  const phone = t.match(/номери телефон:\s*([+\d][\d\s]{5,}\d)/);
  return { position: m ? m[1].trim().toLowerCase() : null, director: m ? m[2].trim() : null,
    email: email ? email[0].replace(/\.$/, '') : null, phone: phone ? phone[1].replace(/\s+/g, ' ').trim() : null };
}

/** ArrayBuffer .xlsx → { sections, companies } (формат data/companies.json) */
export function parseWorkbook(XLSX, buf) {
  const wb = XLSX.read(buf, { type: 'array', cellDates: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rng = XLSX.utils.decode_range(ws['!ref'] || 'A1:M1'); rng.s.r = 0; rng.s.c = 0;   // лист может начинаться не с A1 (у реестра — A3)
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true, range: rng });
  const companies = [], sections = [];
  let topSection = 'Ҷамъияти асосӣ'; sections.push(topSection);
  const sub = {}; let pending = null;
  for (let r = 4; r < rows.length; r++) {
    const row = Array.from({ length: 13 }, (_, i) => rows[r]?.[i] ?? null);
    const empty = v => v == null || v === '';
    if (!empty(row[0]) && empty(row[2]) && row.slice(1).every(empty)) { pending = clean(row[0]); sections.push(pending); continue; }
    if (empty(row[2])) continue;
    const code = (clean(row[0]) || '0').replace(/\.+$/, '').trim();
    const top = code.split('.')[0];
    if (pending) { if (code.includes('.')) sub[top] = pending; else topSection = pending; pending = null; }
    const section = code.includes('.') ? (sub[top] ?? topSection) : topSection;
    const parent = code.includes('.') ? code.split('.').slice(0, -1).join('.') : (code !== '0' ? '0' : null);
    const [founded, rereg] = dates(row[3], XLSX);
    const [extract] = dates(row[12], XLSX);
    companies.push({
      code, parent_code: parent, section, excel_row: r + 1,
      name: qn(oneLine(row[1])), rma: String(row[2]).trim(),
      founded_on: founded, reregistered_on: rereg,
      legal_address: oneLine(row[4]), activity: clean(row[5]),
      capital_declared: num(row[6]), capital_formed: num(row[7]),
      founders: clean(row[8]), founders_parsed: parseFounders(clean(row[8])),
      executive: clean(row[9]), ...parseExec(row[9]),
      supervisory_board: clean(row[10]),
      beneficiary: clean(row[11]), beneficiaries_parsed: parseBenef(clean(row[11])),
      extract_date: extract,
    });
  }
  const byName = new Map(companies.map(c => [c.name.toLowerCase(), c.code]));
  for (const c of companies) {
    for (const f of c.founders_parsed) f.owner_code = byName.get(qn(f.name).toLowerCase()) ?? null;
    c.sort_key = c.code.split('.').map(Number);
  }
  if (!companies.length) throw new Error('В файле не найдено ни одной компании (ожидается формат реестра: РМА в колонке C, данные с 5-й строки)');
  return { sections, companies };
}

/* ================= Таблицы (как в Supabase) ================= */
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2) + Date.now().toString(36));
/** seed → строки таблиц; prev — прежние данные (сохраняем id, выписки и правки id) */
export function toTables(seed, prev) {
  const secId = new Map(seed.sections.map((t, i) => [t, i + 1]));
  const sections = seed.sections.map((t, i) => ({ id: i + 1, title: t, sort_order: i }));
  const oldByCode = new Map((prev?.companies || []).map(c => [c.code, c]));
  const idOf = code => oldByCode.get(code)?.id || 'c-' + code;
  const companies = seed.companies.map(c => ({ id: idOf(c.code), code: c.code, parent_code: c.parent_code, sort_key: c.sort_key, section_id: secId.get(c.section),
    name: c.name, rma: c.rma, founded_on: c.founded_on, reregistered_on: c.reregistered_on, legal_address: c.legal_address, activity: c.activity,
    capital_declared: c.capital_declared, capital_formed: c.capital_formed, founders: c.founders, executive: c.executive, position: c.position ?? null,
    director: c.director ?? null, email: c.email ?? null, phone: c.phone ?? null, supervisory_board: c.supervisory_board, beneficiary: c.beneficiary,
    extract_date: c.extract_date, notes: oldByCode.get(c.code)?.notes ?? null }));
  const shareholders = [], beneficiaries = [];
  for (const c of seed.companies) {
    c.founders_parsed.forEach((f, i) => shareholders.push({ id: uid(), company_id: idOf(c.code), name: f.name, owner_company_id: f.owner_code ? idOf(f.owner_code) : null, share_pct: f.share, nominal: f.nominal, sort_order: i }));
    c.beneficiaries_parsed.forEach((b, i) => beneficiaries.push({ id: uid(), company_id: idOf(c.code), name: b.name, citizenship: b.citizenship, share_pct: b.share, sort_order: i }));
  }
  const ids = new Set(companies.map(c => c.id));
  const extracts = (prev?.extracts || []).filter(e => ids.has(e.company_id));
  // выписка свежее даты в Excel → дата в карточке = дата выписки (как триггер в Supabase)
  for (const c of companies) for (const e of extracts) if (e.company_id === c.id && (!c.extract_date || c.extract_date <= e.extract_date)) c.extract_date = e.extract_date;
  return { sections, companies, shareholders, beneficiaries, extracts, audit: prev?.audit || [] };
}
export function viewCompanies(db) {
  return db.companies.map(c => {
    const ex = db.extracts.filter(x => x.company_id === c.id).sort((a, b) => ((b.extract_date === c.extract_date) - (a.extract_date === c.extract_date)) || b.extract_date.localeCompare(a.extract_date));
    const sec = db.sections.find(s => s.id === c.section_id);
    return { ...c, section_title: sec?.title, extract_id: ex[0]?.id ?? null, extract_path: ex[0]?.storage_path ?? null, extract_file_name: ex[0]?.file_name ?? null, extracts_count: ex.length };
  });
}

/* ================= IndexedDB ================= */
const DB_NAME = 'avesto-reestr', DB_VER = 1;
let _db = null, memory = { kv: new Map(), files: new Map() };
function idb() {
  if (_db) return _db;
  _db = new Promise((res, rej) => {
    try {
      const r = indexedDB.open(DB_NAME, DB_VER);
      r.onupgradeneeded = () => { r.result.createObjectStore('kv'); r.result.createObjectStore('files'); };
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    } catch (e) { rej(e); }
  }).catch(() => null);
  return _db;
}
async function op(store, mode, fn) {
  const db = await idb();
  if (!db) return fn(null, memory[store]);
  return new Promise((res, rej) => { const tx = db.transaction(store, mode); const q = fn(tx.objectStore(store)); tx.oncomplete = () => res(q?.result); tx.onerror = () => rej(tx.error); });
}
const kvGet = k => op('kv', 'readonly', (s, m) => s ? s.get(k) : { result: m.get(k) });
const kvSet = (k, v) => op('kv', 'readwrite', (s, m) => s ? s.put(v, k) : void m.set(k, v));
const fileGet = k => op('files', 'readonly', (s, m) => s ? s.get(k) : { result: m.get(k) });
const fileSet = (k, v) => op('files', 'readwrite', (s, m) => s ? s.put(v, k) : void m.set(k, v));
const fileDel = k => op('files', 'readwrite', (s, m) => s ? s.delete(k) : void m.delete(k));
const clearAll = async () => { await op('kv', 'readwrite', (s, m) => s ? s.clear() : m.clear()); await op('files', 'readwrite', (s, m) => s ? s.clear() : m.clear()); };

/* ================= Local API (тот же интерфейс, что Supabase API) ================= */
const idbStorage = { loadDb: () => kvGet('db'), saveDb: db => kvSet('db', db), getFile: fileGet, putFile: fileSet, delFile: fileDel, clear: clearAll };
/** storage — где лежат данные: IndexedDB (по умолчанию) или зашифрованное облако (cloud.js). opts: { cloud, who(), role() } */
export async function makeLocalApi(loadXLSX, storage = idbStorage, opts = {}) {
  let db = await storage.loadDb() || null;
  const who = () => (opts.who ? opts.who() : 'локально');
  const user = { id: 'local', email: 'локально' };
  // не удалось сохранить (нет прав / кто-то изменил раньше) → возвращаем состояние из хранилища, чтобы экран не врал
  const save = async () => { try { await storage.saveDb(db); } catch (e) { db = await storage.loadDb().catch(() => db); throw e; } };
  const clone = x => JSON.parse(JSON.stringify(x));
  const log = (table_name, action, old_data, new_data) => { db.audit.unshift({ id: Date.now() + Math.random(), table_name, action, old_data, new_data, row_id: (new_data || old_data).id, user_email: who(), created_at: new Date().toISOString() }); db.audit = db.audit.slice(0, 1000); };
  const urls = new Map();
  return {
    local: !opts.cloud,
    async refresh() { db = await storage.loadDb() || db; },
    hasData: () => !!db,
    async importExcel(file) {
      const XLSX = await loadXLSX();
      const seed = parseWorkbook(XLSX, await file.arrayBuffer());
      const prevCount = db?.companies.length || 0;
      db = toTables(seed, db);
      db.source = { file_name: file.name, imported_at: new Date().toISOString() };
      log('reg_companies', 'IMPORT', null, { id: 'import', name: file.name, count: db.companies.length, prev: prevCount });
      await save();
      return { count: db.companies.length, sections: db.sections.length };
    },
    async clear() { await storage.clear(); db = null; },
    source: () => db?.source,
    async getUser() { return user; },
    onAuth() {},
    async signIn() {}, async signOut() {}, async resetPassword() {}, async updatePassword() {},
    async profile() { return { user_id: 'local', email: who(), role: opts.role ? opts.role() : 'editor', full_name: 'Режими локалӣ' }; },
    async loadAll() { if (opts.cloud) await this.refresh(); return clone({ sections: db.sections, companies: viewCompanies(db), shareholders: db.shareholders, beneficiaries: db.beneficiaries, extracts: db.extracts }); },
    async signedUrl(path) {
      if (urls.has(path)) return urls.get(path);
      const blob = await storage.getFile(path); if (!blob) throw new Error('Файл не найден в браузере');
      const u = URL.createObjectURL(blob instanceof Blob ? blob : new Blob([blob], { type: 'application/pdf' })); urls.set(path, u); return u;
    },
    async uploadExtract(company, file, date) {
      const path = `${company.code}/${date}_${Date.now().toString(36)}.pdf`;
      await storage.putFile(path, new Blob([await file.arrayBuffer()], { type: 'application/pdf' }));
      const row = { id: uid(), company_id: company.id, extract_date: date, storage_path: path, file_name: file.name, file_size: file.size, created_at: new Date().toISOString() };
      db.extracts.push(row); log('reg_extracts', 'INSERT', null, row);
      const c = db.companies.find(x => x.id === company.id); if (c && (!c.extract_date || c.extract_date <= date)) c.extract_date = date;
      await save(); return clone(row);
    },
    async deleteExtract(ex) { db.extracts = db.extracts.filter(x => x.id !== ex.id); log('reg_extracts', 'DELETE', ex, null); await save(); await storage.delFile(ex.storage_path).catch(() => {}); },
    async saveCompany(id, patch) {
      if (id) { const c = db.companies.find(x => x.id === id); const old = clone(c); Object.assign(c, patch); log('reg_companies', 'UPDATE', old, clone(c)); await save(); return clone(c); }
      if (db.companies.some(x => x.code === patch.code)) throw new Error(`Компания с № ${patch.code} уже есть`);
      const c = { id: 'c-' + patch.code, ...patch }; db.companies.push(c); log('reg_companies', 'INSERT', null, clone(c)); await save(); return clone(c);
    },
    // карточка + учредители + бенефициары одним сохранением (в облаке = один коммит)
    async saveCompanyFull(id, patch, sh, bn) {
      let c;
      if (id) { c = db.companies.find(x => x.id === id); const old = clone(c); Object.assign(c, patch); log('reg_companies', 'UPDATE', old, clone(c)); }
      else { if (db.companies.some(x => x.code === patch.code)) throw new Error(`Компания с № ${patch.code} уже есть`); c = { id: 'c-' + patch.code, ...patch }; db.companies.push(c); log('reg_companies', 'INSERT', null, clone(c)); }
      const rows = (list) => list.map((r, i) => ({ id: uid(), ...r, company_id: c.id, sort_order: i }));
      db.shareholders = db.shareholders.filter(x => x.company_id !== c.id).concat(rows(sh));
      db.beneficiaries = db.beneficiaries.filter(x => x.company_id !== c.id).concat(rows(bn));
      await save(); return clone(c);
    },
    async replaceRows(table, companyId, rows) {
      const nr = rows.map((r, i) => ({ id: uid(), ...r, company_id: companyId, sort_order: i }));
      if (table === 'reg_shareholders') db.shareholders = db.shareholders.filter(x => x.company_id !== companyId).concat(nr);
      else db.beneficiaries = db.beneficiaries.filter(x => x.company_id !== companyId).concat(nr);
      await save();
    },
    async audit(ids, limit = 300) { return clone(db.audit.filter(a => !ids || ids.includes(a.row_id)).slice(0, limit)); },
    async profiles() { return []; }, async setRole() {}, async addMember() {}, async removeMember() {},
  };
}
