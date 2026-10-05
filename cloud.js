// Облачный режим без сервера: данные и PDF лежат в репозитории GitHub ЗАШИФРОВАННЫМИ (AES-256-GCM).
// Ключ выводится из пароля (PBKDF2-SHA256); без пароля файлы store.* — нечитаемый набор байт.
// Чтение — любому, кто знает пароль. Запись — тому, у кого ещё и GitHub-токен с правом записи в репозиторий.
import { makeLocalApi } from './local.js';

export const STORE = { meta: 'store.meta.json', db: 'store.db.enc', file: p => 'store.f.' + String(p).replace(/[^\w.-]+/g, '_') + '.enc' };
const CHECK = 'avesto-reestr';
const b64 = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
const unb64 = s => { const bin = atob(s.replace(/\s/g, '')); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
const enc = new TextEncoder(), dec = new TextDecoder();

/* ---------- крипто (одинаково работает в браузере и в Node 20+) ---------- */
export async function deriveKey(password, saltB64, iter) {
  const base = await crypto.subtle.importKey('raw', enc.encode(password.normalize('NFKC')), 'PBKDF2', false, ['deriveKey']);
  return crypto.subtle.deriveKey({ name: 'PBKDF2', salt: unb64(saltB64), iterations: iter, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, true, ['encrypt', 'decrypt']);
}
export async function encrypt(key, bytes) {
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const ct = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, bytes));
  const out = new Uint8Array(12 + ct.length); out.set(iv); out.set(ct, 12); return out;
}
export async function decrypt(key, bytes) {
  return new Uint8Array(await crypto.subtle.decrypt({ name: 'AES-GCM', iv: bytes.subarray(0, 12) }, key, bytes.subarray(12)));
}
export async function makeMeta(password, iter = 600000) {
  const salt = b64(crypto.getRandomValues(new Uint8Array(16)));
  const key = await deriveKey(password, salt, iter);
  return { meta: { v: 1, kdf: 'PBKDF2-SHA256', iter, salt, check: b64(await encrypt(key, enc.encode(CHECK))) }, key };
}
export const encryptJson = async (key, obj) => encrypt(key, enc.encode(JSON.stringify(obj)));
export const decryptJson = async (key, bytes) => JSON.parse(dec.decode(await decrypt(key, bytes)));

/* ---------- хранилище ---------- */
const ls = {
  get(k) { try { return localStorage.getItem(k); } catch { return null; } },
  set(k, v) { try { v == null ? localStorage.removeItem(k) : localStorage.setItem(k, v); } catch {} },
};

export async function makeCloudApi(cfg, loadXLSX) {
  const API = `${cfg.api || 'https://api.github.com'}/repos/${cfg.repo}/contents/`;
  const branch = cfg.branch || 'main';
  let key = null, token = ls.get('reg.ghToken'), inner = null, dbSha = null, login = ls.get('reg.ghLogin') || null;
  const savedKey = ls.get('reg.cloudKey');
  if (savedKey) { try { key = await crypto.subtle.importKey('raw', unb64(savedKey), 'AES-GCM', true, ['encrypt', 'decrypt']); } catch { key = null; } }

  const headers = (raw) => ({ Accept: raw ? 'application/vnd.github.raw+json' : 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28', ...(token ? { Authorization: `Bearer ${token}` } : {}) });
  const gh = async (path, opt = {}, raw = false) => {
    const r = await fetch(API + encodeURIComponent(path) + (opt.method ? '' : `?ref=${branch}&t=${Date.now()}`), { ...opt, headers: { ...headers(raw), ...(opt.body ? { 'Content-Type': 'application/json' } : {}) }, cache: 'no-store' });
    return r;
  };
  // статическая копия на GitHub Pages — запасной путь, когда у анонимного API кончился лимит (60 запросов/час с одного IP)
  const pages = path => fetch(new URL(path, location.href.split(/[?#]/)[0]).href + `?t=${Date.now()}`, { cache: 'no-store' });

  async function readBytes(path, wantSha) {
    if (wantSha || token) {
      const r = await gh(path);
      if (r.ok) { const j = await r.json(); if (j.content) return { bytes: unb64(j.content), sha: j.sha };
        const r2 = await gh(path, {}, true); if (r2.ok) return { bytes: new Uint8Array(await r2.arrayBuffer()), sha: j.sha }; }
      else if (r.status === 404) return null;
      else if (r.status === 401) { token = null; ls.set('reg.ghToken', null); }
    }
    const p = await pages(path);
    if (p.ok) return { bytes: new Uint8Array(await p.arrayBuffer()), sha: null };
    if (p.status === 404) return null;
    throw new Error(`Не удалось загрузить ${path} (${p.status})`);
  }
  async function writeBytes(path, bytes, sha, message) {
    if (!token) throw new Error('Нужен режим правки: введите GitHub-токен (кнопка с ключом вверху)');
    const r = await gh(path, { method: 'PUT', body: JSON.stringify({ message, content: b64(bytes), branch, ...(sha ? { sha } : {}) }) });
    if (r.status === 409 || r.status === 422) throw new Error('Данные изменил кто-то другой — они обновлены на экране, повторите правку');
    if (r.status === 401 || r.status === 403 || r.status === 404) throw new Error('GitHub не принял токен: нужен доступ Contents «Read and write» к репозиторию ' + cfg.repo);
    if (!r.ok) throw new Error('GitHub: ошибка сохранения ' + r.status);
    return (await r.json()).content.sha;
  }
  const storage = {
    async loadDb() { const f = await readBytes(STORE.db, true); if (!f) return null; dbSha = f.sha; return decryptJson(key, f.bytes); },
    async saveDb(db) { dbSha = await writeBytes(STORE.db, await encryptJson(key, db), dbSha, 'Реестр: обновление данных'); },
    async getFile(p) { const f = await readBytes(STORE.file(p)); return f ? new Blob([await decrypt(key, f.bytes)], { type: 'application/pdf' }) : null; },
    async putFile(p, blob) { await writeBytes(STORE.file(p), await encrypt(key, new Uint8Array(await blob.arrayBuffer())), null, 'Реестр: выписка'); },
    async delFile(p) { if (!token) return; const r = await gh(STORE.file(p)); if (!r.ok) return; const j = await r.json();
      await gh(STORE.file(p), { method: 'DELETE', body: JSON.stringify({ message: 'Реестр: удаление выписки', sha: j.sha, branch }) }); },
    async clear() { throw new Error('В облачном режиме очистка отключена'); },
  };
  const role = () => (token ? 'editor' : 'viewer');
  const build = async () => { inner = await makeLocalApi(loadXLSX, storage, { cloud: true, role, who: () => login || (token ? 'редактор' : 'просмотр') }); };

  const api = {
    cloud: true, keyLogin: true,
    hasToken: () => !!token,
    async getUser() {
      if (!key) return null;
      try { await build(); } catch (e) { if (e.name === 'OperationError') { key = null; ls.set('reg.cloudKey', null); return null; } throw e; }
      return { id: 'cloud', email: role() };
    },
    onAuth() {},
    async signIn(_e, password) {
      const m = await readBytes(STORE.meta);
      if (!m) throw new Error('Данные ещё не опубликованы (нет store.meta.json)');
      const meta = JSON.parse(dec.decode(m.bytes));
      const k = await deriveKey(String(password || ''), meta.salt, meta.iter);
      try { if (dec.decode(await decrypt(k, unb64(meta.check))) !== CHECK) throw 0; } catch { throw new Error('Рамз нодуруст аст / Неверный пароль'); }
      key = k; ls.set('reg.cloudKey', b64(new Uint8Array(await crypto.subtle.exportKey('raw', k))));
      await build();
    },
    async signOut() { key = null; inner = null; token = null; login = null; ['reg.cloudKey', 'reg.ghToken', 'reg.ghLogin'].forEach(k => ls.set(k, null)); },
    async resetPassword() {}, async updatePassword() {},
    /** включить режим правки: проверяем, что токен реально может писать в репозиторий */
    async setToken(t) {
      t = String(t || '').trim();
      if (!t) { token = null; login = null; ls.set('reg.ghToken', null); ls.set('reg.ghLogin', null); return; }
      const r = await fetch(`${cfg.api || 'https://api.github.com'}/repos/${cfg.repo}`, { headers: { Accept: 'application/vnd.github+json', Authorization: `Bearer ${t}` }, cache: 'no-store' });
      if (!r.ok) throw new Error('GitHub не принял токен (' + r.status + ')');
      const j = await r.json();
      if (!j.permissions?.push) throw new Error('У токена нет права записи: нужен доступ Contents «Read and write» к ' + cfg.repo);
      token = t; ls.set('reg.ghToken', t);
      const u = await fetch(`${cfg.api || 'https://api.github.com'}/user`, { headers: { Authorization: `Bearer ${t}` } }).then(x => x.ok ? x.json() : null).catch(() => null);
      login = u?.login || j.owner?.login || null; ls.set('reg.ghLogin', login);
    },
  };
  for (const m of ['hasData', 'importExcel', 'source', 'profile', 'loadAll', 'signedUrl', 'uploadExtract', 'deleteExtract', 'saveCompany', 'saveCompanyFull', 'replaceRows', 'audit', 'profiles', 'setRole', 'addMember', 'removeMember', 'refresh'])
    api[m] = (...a) => inner[m](...a);
  return api;
}
