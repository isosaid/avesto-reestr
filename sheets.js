// Режим Google Sheets: данные в Google Таблице, PDF в папке Drive, доступ по ключу через Apps Script.
import { viewCompanies } from './local.js';

const KEY_STORE = 'reg.sheetsKey';
const store = {
  get() { try { return localStorage.getItem(KEY_STORE); } catch { return null; } },
  set(v) { try { v ? localStorage.setItem(KEY_STORE, v) : localStorage.removeItem(KEY_STORE); } catch {} },
};
const toB64 = async file => {
  const buf = new Uint8Array(await file.arrayBuffer()); let s = '';
  for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
  return btoa(s);
};
const fromB64 = b64 => { const bin = atob(b64); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
const n = v => (v === null || v === undefined || v === '' ? null : Number(String(v).replace(/[\s ]/g, '').replace(',', '.')));
const s = v => (v === null || v === undefined || v === '' ? null : String(v));
const d = v => (v ? String(v).slice(0, 10) : null);
const uid = () => Math.random().toString(36).slice(2) + Date.now().toString(36);

export function makeSheetsApi(url) {
  let key = store.get(), role = null, last = null;
  const pdfCache = new Map();
  async function call(action, payload = {}) {
    if (!key) throw Object.assign(new Error('no key'), { code: 401 });
    let r;
    try {
      r = await fetch(url, { method: 'POST', redirect: 'follow', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ action, key, ...payload }) });
    } catch (e) { throw new Error('Нет связи с Google Apps Script: ' + e.message); }
    const j = await r.json().catch(() => ({ ok: false, error: 'Ответ сервера не JSON (проверьте развертывание: доступ «Все»)' }));
    if (!j.ok) { if (j.code === 401) { key = null; store.set(null); } throw Object.assign(new Error(j.error || 'Ошибка'), { code: j.code }); }
    role = j.role; return j.data;
  }
  function tables(raw) {
    const titles = [...new Set([...(raw.sections || []).map(x => s(x.title)).filter(Boolean), ...raw.companies.map(c => s(c.section)).filter(Boolean)])];
    const sections = titles.map((t, i) => ({ id: i + 1, title: t, sort_order: i }));
    const secId = new Map(titles.map((t, i) => [t, i + 1]));
    const codes = new Set(raw.companies.map(c => s(c.code)));
    const companies = raw.companies.map(c => {
      const code = s(c.code).replace(/\.$/, '');
      return { id: 'c-' + code, code, parent_code: s(c.parent_code), sort_key: code.split('.').map(Number), section_id: secId.get(s(c.section)) ?? null,
        name: s(c.name), rma: s(c.rma)?.padStart(9, '0'), founded_on: d(c.founded_on), reregistered_on: d(c.reregistered_on), legal_address: s(c.legal_address),
        activity: s(c.activity), capital_declared: n(c.capital_declared), capital_formed: n(c.capital_formed), founders: s(c.founders), executive: s(c.executive),
        position: s(c.position), director: s(c.director), email: s(c.email), phone: s(c.phone), supervisory_board: s(c.supervisory_board),
        beneficiary: s(c.beneficiary), extract_date: d(c.extract_date), notes: s(c.notes) };
    });
    const by = (rows, map) => rows.filter(r => codes.has(s(r.company_code))).map((r, i) => ({ id: uid(), company_id: 'c-' + s(r.company_code), sort_order: i, ...map(r) }));
    const shareholders = by(raw.shareholders || [], r => ({ name: s(r.name), owner_company_id: r.owner_code && codes.has(s(r.owner_code)) ? 'c-' + s(r.owner_code) : null, share_pct: n(r.share_pct), nominal: n(r.nominal) }));
    const beneficiaries = by(raw.beneficiaries || [], r => ({ name: s(r.name), citizenship: s(r.citizenship), share_pct: n(r.share_pct) }));
    const extracts = (raw.extracts || []).filter(e => codes.has(s(e.company_code)) && e.file_id).map(e => ({ id: s(e.file_id), company_id: 'c-' + s(e.company_code),
      extract_date: d(e.extract_date), storage_path: s(e.file_id), file_name: s(e.file_name), created_at: s(e.uploaded_at) }));
    return { sections, companies, shareholders, beneficiaries, extracts };
  }
  let sectionsCache = [];
  return {
    sheets: true, keyLogin: true,
    async getUser() { if (!key) return null; try { last = await call('load'); return { id: 'key', email: role }; } catch (e) { if (e.code === 401) return null; throw e; } },
    onAuth() {},
    async signIn(_email, password) { key = String(password || '').trim(); last = await call('load'); store.set(key); },
    async signOut() { key = null; role = null; store.set(null); },
    async resetPassword() {}, async updatePassword() {},
    async profile() { return { user_id: 'key', email: role === 'editor' ? 'муҳаррир' : 'тамошобин', role, full_name: 'Google Sheets' }; },
    folderUrl: () => last?.folderUrl,
    async loadAll() {
      const raw = last || await call('load'); last = null;
      const db = tables(raw); sectionsCache = db.sections; this._folder = raw.folderUrl;
      return { sections: db.sections, companies: viewCompanies(db), shareholders: db.shareholders, beneficiaries: db.beneficiaries, extracts: db.extracts };
    },
    async signedUrl(fileId) {
      if (pdfCache.has(fileId)) return pdfCache.get(fileId);
      const r = await call('getPdf', { file_id: fileId });
      const u = URL.createObjectURL(new Blob([fromB64(r.base64)], { type: 'application/pdf' })); pdfCache.set(fileId, u); return u;
    },
    async uploadExtract(company, file, date) {
      const r = await call('uploadPdf', { code: company.code, date, file_name: file.name, base64: await toB64(file) });
      return { id: r.file_id, company_id: company.id, extract_date: r.extract_date, storage_path: r.file_id, file_name: r.file_name };
    },
    async deleteExtract(ex) { await call('deleteExtract', { file_id: ex.storage_path }); },
    async saveCompanyFull(id, patch, sh, bn, byId) {
      const sec = sectionsCache.find(x => x.id === patch.section_id);
      const p = { ...patch, section: sec?.title ?? null }; delete p.section_id; delete p.sort_key;
      const codeOf = cid => byId.get(cid)?.code ?? null;
      const row = await call('saveCompany', { old_code: id ? id.replace(/^c-/, '') : null, patch: p,
        shareholders: sh.map(x => ({ name: x.name, owner_code: x.owner_company_id ? codeOf(x.owner_company_id) : null, share_pct: x.share_pct, nominal: x.nominal })),
        beneficiaries: bn.map(x => ({ name: x.name, citizenship: x.citizenship, share_pct: x.share_pct })) });
      return { ...row, id: 'c-' + row.code };
    },
    async audit(ids, limit = 300) {
      const rows = await call('audit', { limit: 1000 });
      return rows.map((a, i) => ({ id: i, created_at: a.time, user_email: a.role, table_name: 'reg_companies', action: a.action, row_id: 'c-' + s(a.company_code), fields: s(a.details) }))
        .filter(a => !ids || ids.includes(a.row_id)).slice(0, limit);
    },
    async profiles() { return []; }, async setRole() {}, async addMember() {}, async removeMember() {},
  };
}
