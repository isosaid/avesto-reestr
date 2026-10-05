// Феҳристи ҷамъиятҳои «Авесто Гуруҳ» — веб-клиент Supabase
import { CONFIG } from './config.js';
import { makeLocalApi } from './local.js';
import { makeSheetsApi } from './sheets.js';
import { makeCloudApi } from './cloud.js';

const LIB = {
  supabase: 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2.45.4/+esm',
  xlsx: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
  pdfjs: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.mjs',
  pdfjsWorker: 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs',
};

/* ================= i18n ================= */
const I18N = {
  tg: {
    appTitle: 'Феҳристи ҷамъиятҳо', appSub: 'ҶДММ «Авесто Гуруҳ» · ҷамъиятҳои фаръӣ ва вобаста',
    login: 'Ворид шудан', loginSub: 'Барои дастрасӣ ба феҳрист ворид шавед', email: 'Почтаи электронӣ', password: 'Рамз',
    forgot: 'Рамзро фаромӯш кардед?', resetSent: 'Истинод барои барқарорсозӣ ба почта фиристода шуд', newPassword: 'Рамзи нав', savePassword: 'Рамзро нигоҳ доштан',
    logout: 'Баромадан', noAccess: 'Ба ҳисоби шумо ҳанӯз нақш дода нашудааст. Ба администратор муроҷиат кунед.',
    tabRegistry: 'Феҳрист', tabTree: 'Сохтор', tabChecks: 'Санҷиш', tabAudit: 'Журнал', tabUsers: 'Корбарон',
    kCompanies: 'Ҷамъиятҳо', kFull: 'Фаръии 100%', kPartial: 'Муштарак / вобаста', kCapital: 'Сармояи эълоншуда', kFormed: 'ташаккул ёфт', kUnformed: 'Сармоя пурра нест',
    kExtracts: 'Иқтибосҳо (PDF)', kOld: 'Иқтибоси кӯҳна', kOldSub: 'зиёда аз {0} рӯз', kIssues: 'Хатоҳо дар маълумот', kOf: 'аз {0}',
    search: 'Ҷустуҷӯ: ном, РМА, роҳбар, почта, телефон…', found: '{0} аз {1}', compact: 'Мухтасар', full: 'Пурра', all: 'Ҳама',
    exportXlsx: 'Excel', importPdf: 'Бор кардани иқтибосҳо', newCompany: 'Ҷамъияти нав',
    cCode: '№', cName: 'Номи ҷамъият', cRma: 'РМА', cFounded: 'Санаи таъсис', cAddress: 'Суроғаи ҳуқуқӣ', cActivity: 'Намуди фаъолият',
    cDeclared: 'Сармояи эълоншуда', cFormed: 'Сармояи ташаккулёфта', cFounders: 'Муассисон', cExecutive: 'Мақоми иҷроия', cBoard: 'Шӯрои нозирон',
    cBenef: 'Молик-бенефитсиар', cExtract: 'Иқтибос', noFile: 'файл нест', openPdf: 'Кушодани иқтибос', noRows: 'Ҳеҷ чиз ёфт нашуд',
    rereg: 'бақайдгирии такрорӣ', mainInfo: 'Маълумоти асосӣ', capital: 'Сармояи оинномавӣ', declared: 'Эълоншуда', formed: 'Ташаккулёфта',
    founders: 'Муассисон', share: 'Ҳисса', nominal: 'Арзиши номиналӣ', expected: 'бояд', officialText: 'Матни расмӣ',
    beneficiaries: 'Молик-бенефитсиарон', citizenship: 'Шаҳрвандӣ', management: 'Мақоми иҷроия', board: 'Шӯрои нозирон', subsidiaries: 'Ҷамъиятҳои фаръӣ',
    extracts: 'Иқтибос аз ФЯД', current: 'ҷорӣ', uploadNew: 'Бор кардани иқтибоси нав', extractDate: 'Санаи иқтибос', upload: 'Бор кардан', download: 'Боргирӣ', openNew: 'Дар варақаи нав',
    del: 'Нест кардан', confirmDel: 'Иқтибосро нест кунем?', noExtracts: 'Иқтибос бор нашудааст', notes: 'Эзоҳ', history: 'Таърихи тағйирот', showHistory: 'Нишон додан',
    edit: 'Таҳрир', save: 'Нигоҳ доштан', cancel: 'Бекор', close: 'Пӯшидан', saved: 'Нигоҳ дошта шуд', uploaded: 'Иқтибос бор шуд', deleted: 'Нест карда шуд',
    position: 'Вазифа', director: 'Роҳбар', phone: 'Телефон', section: 'Бахш', parent: 'Ҷамъияти асосӣ', addRow: '+ Сатр', name: 'Ном', code: 'Рақам (№)',
    effShare: 'Ҳиссаи самаранок', groupShare: 'ҳиссаи гурӯҳ', typeRoot: 'асосӣ', typeFull: 'фаръии 100%', typePartial: 'вобаста',
    impTitle: 'Бор кардани иқтибосҳо (PDF)', impDrop: '<b>Файлҳои PDF-ро ба ин ҷо кашед</b> ё пахш кунед', impHint: 'Ҷамъият аз рӯи рақам дар номи файл (масалан «1.2. … 13.04.2026.pdf») ва РМА дар дохили PDF муайян мешавад',
    impFile: 'Файл', impCompany: 'Ҷамъият', impDate: 'Сана', impStatus: 'Ҳолат', impGo: 'Бор кардан ({0})', impOkRma: 'РМА мувофиқ', impByCode: 'аз рӯи рақам', impRmaDiff: 'РМА фарқ мекунад',
    impNoMatch: 'ёфт нашуд', impExists: 'аллакай ҳаст', impDone: 'бор шуд', impErr: 'хато', impReading: 'хондан…', impUploading: 'бор шуда истодааст…', choose: '— интихоб —',
    chkNone: 'Ҳама санҷишҳо гузаштанд', aTime: 'Вақт', aUser: 'Корбар', aAction: 'Амал', aObject: 'Объект', aFields: 'Майдонҳо', role: 'Нақш',
    usersHint: 'Танҳо корбарони ин рӯйхат феҳристро мебинанд. Корбари навро аввал дар Supabase → Authentication → Users даъват кунед, баъд почтаашро ин ҷо илова кунед.', addMember: 'Дастрасӣ додан',
    rViewer: 'тамошобин', rEditor: 'муҳаррир', rAdmin: 'админ', copied: 'Нусха гирифта шуд', loadErr: 'Хатои боркунӣ', required: 'Майдонҳои ҳатмӣ: №, ном, РМА (9 рақам)',
    welcomeTitle: 'Феҳристи ҷамъиятҳо', welcomeSub: 'Файли Excel-и феҳристро кашед — ҳамроҳ бо PDF-и иқтибосҳо (ихтиёрӣ).',
    welcomeDrop: '<b>Excel (.xlsx) ва PDF-ҳоро ба ин ҷо кашед</b> ё пахш кунед', welcomePrivacy: 'Маълумот танҳо дар ҳамин браузер нигоҳ дошта мешавад ва ба ҳеҷ сервер фиристода намешавад.',
    welcomeNeedXlsx: 'Файли Excel (.xlsx) лозим аст', localMode: 'Режими локалӣ', localHint: 'маълумот танҳо дар ин браузер', reimport: 'Навсозӣ аз Excel',
    clearData: 'Тоза кардани маълумот', confirmClear: 'Ҳама маълумот ва PDF-ҳо аз ин браузер нест мешаванд. Боварӣ доред? Бори дигар пахш кунед.', imported: 'Бор шуд: {0} ҷамъият', reading: 'Хондан…',
    accessKey: 'Калиди дастрасӣ', keyHint: 'Калидро аз администратор гиред', pdfFolder: 'Папкаи PDF', openSheet: 'Google Sheets',
    editMode: 'Режими таҳрир', tokenTitle: 'Режими таҳрир (GitHub-токен)', tokenOn: 'Режими таҳрир фаъол аст', tokenOff: 'Хомӯш кардан', tokenSave: 'Фаъол кардан',
    tokenHelp: 'Барои таҳрир ва бор кардани PDF токени GitHub лозим аст: github.com → Settings → Developer settings → Fine-grained tokens → Generate new token → Repository access: танҳо avesto-reestr → Permissions → Contents: Read and write. Токен танҳо дар ҳамин браузер нигоҳ дошта мешавад.',
    themeToggle: 'Мавзӯъ', legendFull: '100% дар гурӯҳ', legendPartial: '50–99%', legendMinor: 'то 50%', groupOwned: 'ҳиссаи Авесто Гуруҳ',
  },
  ru: {
    appTitle: 'Реестр юрлиц', appSub: 'ООО «Авесто Гуруҳ» · дочерние и зависимые общества',
    login: 'Войти', loginSub: 'Войдите, чтобы открыть реестр', email: 'Эл. почта', password: 'Пароль',
    forgot: 'Забыли пароль?', resetSent: 'Ссылка для восстановления отправлена на почту', newPassword: 'Новый пароль', savePassword: 'Сохранить пароль',
    logout: 'Выйти', noAccess: 'Вашей учётной записи ещё не назначена роль. Обратитесь к администратору.',
    tabRegistry: 'Реестр', tabTree: 'Структура', tabChecks: 'Проверки', tabAudit: 'Журнал', tabUsers: 'Пользователи',
    kCompanies: 'Компаний', kFull: 'Дочерние 100%', kPartial: 'Совместные / зависимые', kCapital: 'Объявленный капитал', kFormed: 'сформировано', kUnformed: 'Капитал не сформирован',
    kExtracts: 'Выписки (PDF)', kOld: 'Старые выписки', kOldSub: 'старше {0} дней', kIssues: 'Ошибки в данных', kOf: 'из {0}',
    search: 'Поиск: название, РМА, руководитель, почта, телефон…', found: '{0} из {1}', compact: 'Кратко', full: 'Полностью', all: 'Все',
    exportXlsx: 'Excel', importPdf: 'Загрузить выписки', newCompany: 'Новая компания',
    cCode: '№', cName: 'Название', cRma: 'РМА', cFounded: 'Дата создания', cAddress: 'Юр. адрес', cActivity: 'Вид деятельности',
    cDeclared: 'Капитал объявленный', cFormed: 'Капитал сформированный', cFounders: 'Учредители', cExecutive: 'Исп. орган', cBoard: 'Набл. совет',
    cBenef: 'Бенефициар', cExtract: 'Выписка', noFile: 'нет файла', openPdf: 'Открыть выписку', noRows: 'Ничего не найдено',
    rereg: 'перерегистрация', mainInfo: 'Основное', capital: 'Уставный капитал', declared: 'Объявленный', formed: 'Сформированный',
    founders: 'Учредители', share: 'Доля', nominal: 'Номинал', expected: 'должно', officialText: 'Официальный текст',
    beneficiaries: 'Бенефициары', citizenship: 'Гражданство', management: 'Исполнительный орган', board: 'Наблюдательный совет', subsidiaries: 'Дочерние общества',
    extracts: 'Выписки из ЕГР', current: 'текущая', uploadNew: 'Загрузить новую выписку', extractDate: 'Дата выписки', upload: 'Загрузить', download: 'Скачать', openNew: 'В новой вкладке',
    del: 'Удалить', confirmDel: 'Удалить выписку?', noExtracts: 'Выписка не загружена', notes: 'Примечание', history: 'История изменений', showHistory: 'Показать',
    edit: 'Редактировать', save: 'Сохранить', cancel: 'Отмена', close: 'Закрыть', saved: 'Сохранено', uploaded: 'Выписка загружена', deleted: 'Удалено',
    position: 'Должность', director: 'Руководитель', phone: 'Телефон', section: 'Раздел', parent: 'Материнская', addRow: '+ Строка', name: 'Имя', code: 'Номер (№)',
    effShare: 'Эффективная доля', groupShare: 'доля группы', typeRoot: 'головная', typeFull: 'дочерняя 100%', typePartial: 'зависимая',
    impTitle: 'Загрузка выписок (PDF)', impDrop: '<b>Перетащите PDF-файлы сюда</b> или нажмите', impHint: 'Компания определяется по номеру в имени файла («1.2. … 13.04.2026.pdf») и по РМА внутри PDF',
    impFile: 'Файл', impCompany: 'Компания', impDate: 'Дата', impStatus: 'Статус', impGo: 'Загрузить ({0})', impOkRma: 'РМА совпал', impByCode: 'по номеру', impRmaDiff: 'РМА не совпал',
    impNoMatch: 'не найдено', impExists: 'уже есть', impDone: 'загружено', impErr: 'ошибка', impReading: 'чтение…', impUploading: 'загрузка…', choose: '— выбрать —',
    chkNone: 'Все проверки пройдены', aTime: 'Время', aUser: 'Пользователь', aAction: 'Действие', aObject: 'Объект', aFields: 'Поля', role: 'Роль',
    usersHint: 'Реестр видят только пользователи из этого списка. Нового сначала пригласите в Supabase → Authentication → Users, затем добавьте его почту здесь.', addMember: 'Дать доступ',
    rViewer: 'просмотр', rEditor: 'редактор', rAdmin: 'админ', copied: 'Скопировано', loadErr: 'Ошибка загрузки', required: 'Обязательно: №, название, РМА (9 цифр)',
    welcomeTitle: 'Реестр юрлиц', welcomeSub: 'Перетащите Excel-файл реестра — вместе с PDF выписок (необязательно).',
    welcomeDrop: '<b>Перетащите сюда Excel (.xlsx) и PDF</b> или нажмите', welcomePrivacy: 'Данные хранятся только в этом браузере и никуда не отправляются.',
    welcomeNeedXlsx: 'Нужен файл Excel (.xlsx)', localMode: 'Локальный режим', localHint: 'данные только в этом браузере', reimport: 'Обновить из Excel',
    clearData: 'Очистить данные', confirmClear: 'Все данные и PDF будут удалены из этого браузера. Точно? Нажмите ещё раз.', imported: 'Загружено: {0} компаний', reading: 'Чтение…',
    accessKey: 'Ключ доступа', keyHint: 'Ключ выдаёт администратор', pdfFolder: 'Папка PDF', openSheet: 'Google Sheets',
    editMode: 'Режим правки', tokenTitle: 'Режим правки (GitHub-токен)', tokenOn: 'Режим правки включён', tokenOff: 'Выключить', tokenSave: 'Включить',
    tokenHelp: 'Для правки и загрузки PDF нужен токен GitHub: github.com → Settings → Developer settings → Fine-grained tokens → Generate new token → Repository access: только avesto-reestr → Permissions → Contents: Read and write. Токен хранится только в этом браузере.',
    themeToggle: 'Тема', legendFull: '100% в группе', legendPartial: '50–99%', legendMinor: 'до 50%', groupOwned: 'доля Авесто Гуруҳ',
  },
};
const store = {
  get(k, d) { try { return localStorage.getItem(k) ?? d; } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, v); } catch {} },
};
let lang = store.get('reg.lang', 'tg');
const t = (k, ...a) => { let s = I18N[lang][k] ?? I18N.tg[k] ?? k; a.forEach((v, i) => { s = s.replace(`{${i}}`, v); }); return s; };

/* ================= Checks (правила проверки данных) ================= */
const CHECKS = [
  { id: 'unformed', sev: 'warn', tg: 'Сармоя пурра ташаккул наёфтааст', ru: 'Капитал сформирован не полностью',
    dtg: 'Сармояи ташаккулёфта аз эълоншуда кам аст', dru: 'Сформированный капитал меньше объявленного',
    test: c => c.capital_declared > 0 && (c.capital_formed ?? 0) < c.capital_declared ? `${fmtNum(c.capital_formed ?? 0)} / ${fmtNum(c.capital_declared)}` : null },
  { id: 'nominal', sev: 'error', tg: 'Арзиши номиналӣ ба ҳисса мувофиқ нест', ru: 'Номинал не соответствует доле',
    dtg: 'Номинал ≠ сармояи эълоншуда × ҳисса', dru: 'Номинал ≠ объявленный капитал × доля',
    test: c => { const bad = c.sh.filter(s => s.nominal != null && c.capital_declared && Math.abs(s.nominal - c.capital_declared * s.share_pct / 100) > 1);
      return bad.length ? bad.map(s => `${s.share_pct}% → ${fmtNum(s.nominal)} (${t('expected')} ${fmtNum(Math.round(c.capital_declared * s.share_pct / 100))})`).join('; ') : null; } },
  { id: 'sum', sev: 'error', tg: 'Ҷамъи ҳиссаҳо ≠ 100%', ru: 'Сумма долей ≠ 100%',
    dtg: 'Ҳиссаҳои муассисон якҷоя 100% намешаванд', dru: 'Доли учредителей в сумме не дают 100%',
    test: c => { if (!c.sh.length) return null; const s = c.sh.reduce((a, x) => a + Number(x.share_pct || 0), 0); return Math.abs(s - 100) > .01 ? `Σ ${round2(s)}%` : null; } },
  { id: 'benef_diff', sev: 'error', tg: 'Ҳиссаи бенефитсиар ба занҷир мувофиқ нест', ru: 'Доля бенефициара не сходится с цепочкой',
    dtg: 'Ҳиссаи самаранок аз рӯи муассисон ≠ ҳиссаи навишташуда', dru: 'Расчётная доля по учредителям ≠ указанной',
    test: c => c.ownerStated != null && c.eff != null && Math.abs(c.ownerStated - c.eff) > .05 ? `${round2(c.ownerStated)}% ≠ ${round2(c.eff)}%` : null },
  { id: 'no_benef', sev: 'warn', tg: 'Бенефитсиар нишон дода нашудааст', ru: 'Бенефициар не указан',
    dtg: 'Сутуни «Молик-бенефитсиар» холӣ', dru: 'Колонка «Бенефициар» пустая',
    test: c => !c.bn.length ? (c.eff ? `${t('effShare')}: ${round2(c.eff)}%` : '—') : null },
  { id: 'no_file', sev: 'warn', tg: 'Файли иқтибос бор нашудааст', ru: 'Файл выписки не загружен',
    dtg: 'Сана ҳаст, вале PDF нест', dru: 'Дата есть, PDF нет',
    test: c => !c.extract_path ? fmtDate(c.extract_date) || '—' : null },
  { id: 'old', sev: 'info', tg: 'Иқтибос кӯҳна аст', ru: 'Выписка устарела',
    dtg: `Иқтибос зиёда аз ${CONFIG.EXTRACT_MAX_AGE_DAYS} рӯз пеш гирифта шудааст`, dru: `Выписке больше ${CONFIG.EXTRACT_MAX_AGE_DAYS} дней`,
    test: c => c.extract_date && daysSince(c.extract_date) > CONFIG.EXTRACT_MAX_AGE_DAYS ? `${fmtDate(c.extract_date)} · ${daysSince(c.extract_date)} д.` : null },
  { id: 'no_date', sev: 'warn', tg: 'Санаи иқтибос нест', ru: 'Нет даты выписки', dtg: 'Сутуни «Иқтибос» холӣ', dru: 'Колонка «Выписка» пустая',
    test: c => !c.extract_date ? '—' : null },
];
const chkTitle = ch => lang === 'ru' ? ch.ru : ch.tg;
const chkDesc = ch => lang === 'ru' ? ch.dru : ch.dtg;
const SEV_ORDER = { error: 0, warn: 1, info: 2 };

/* ================= Utils ================= */
const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
const NF = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 2 });
const fmtNum = n => (n == null || n === '' ? '—' : NF.format(Number(n)));
const fmtShort = n => { n = Number(n || 0); if (n >= 1e9) return NF.format(round2(n / 1e9)) + (lang === 'ru' ? ' млрд' : ' млрд'); if (n >= 1e6) return NF.format(round2(n / 1e6)) + ' млн'; return NF.format(n); };
const round2 = n => Math.round(Number(n) * 100) / 100;
const fmtDate = d => { if (!d) return ''; const [y, m, dd] = String(d).slice(0, 10).split('-'); return `${dd}.${m}.${y}`; };
const today = () => new Date().toISOString().slice(0, 10);
const daysSince = d => Math.floor((Date.now() - new Date(d + 'T00:00:00').getTime()) / 864e5);
const cmpKey = (a, b) => { for (let i = 0; i < Math.max(a.length, b.length); i++) { const x = a[i] ?? -1, y = b[i] ?? -1; if (x !== y) return x - y; } return 0; };
const codeKey = code => String(code).split('.').filter(Boolean).map(Number);
const TJ = { 'ӣ': 'и', 'ҳ': 'х', 'ҷ': 'ч', 'ғ': 'г', 'қ': 'к', 'ӯ': 'у', 'ё': 'е' };
// нормализация с сохранением длины (для подсветки)
const norm = s => String(s ?? '').toLowerCase().replace(/[ӣҳҷғқӯё]/g, ch => TJ[ch]).replace(/[“”«»"„]/g, ' ');
const normName = s => norm(s).replace(/\s+/g, ' ').trim();
function hl(text, terms) {
  const src = String(text ?? '');
  if (!terms.length || !src) return esc(src);
  const n = norm(src); const marks = [];
  for (const term of terms) { let i = 0; while ((i = n.indexOf(term, i)) !== -1) { marks.push([i, i + term.length]); i += term.length; } }
  if (!marks.length) return esc(src);
  marks.sort((a, b) => a[0] - b[0]); const merged = [];
  for (const m of marks) { const last = merged[merged.length - 1]; if (last && m[0] <= last[1]) last[1] = Math.max(last[1], m[1]); else merged.push([...m]); }
  let out = '', p = 0; for (const [a, b] of merged) { out += esc(src.slice(p, a)) + '<mark>' + esc(src.slice(a, b)) + '</mark>'; p = b; }
  return out + esc(src.slice(p));
}
function toast(msg, err) { const el = document.createElement('div'); el.className = 'toast' + (err ? ' err' : ''); el.textContent = msg; $('#toast-root').append(el); setTimeout(() => el.remove(), err ? 6000 : 2600); }
const loadScript = src => new Promise((res, rej) => { const s = document.createElement('script'); s.src = src; s.onload = res; s.onerror = () => rej(new Error('load ' + src)); document.head.append(s); });
const initials = s => (s || '?').split(/[\s@.]/).filter(Boolean).slice(0, 2).map(x => x[0].toUpperCase()).join('');

const IC = {
  search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
  pdf: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M9 13h6M9 17h4"/></svg>',
  x: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M18 6 6 18M6 6l12 12"/></svg>',
  up: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 16V4m0 0-5 5m5-5 5 5M4 20h16"/></svg>',
  down: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 4v12m0 0-5-5m5 5 5-5M4 20h16"/></svg>',
  xls: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M3 15h18M9 3v18"/></svg>',
  plus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>',
  ext: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 4h6v6M20 4l-9 9M19 14v5a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V6a1 1 0 0 1 1-1h5"/></svg>',
  mail: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3 7 9 6 9-6"/></svg>',
  phone: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2"/></svg>',
  copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="9" y="9" width="11" height="11" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>',
  moon: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8"/></svg>',
  trash: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 7h16M10 11v6M14 11v6M5 7l1 12a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2l1-12M9 7V4h6v3"/></svg>',
  key: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="15" r="4"/><path d="m11 12 9-9m-3 3 2.5 2.5M14 9l2 2"/></svg>',
  out: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M15 12H3m0 0 4-4m-4 4 4 4M13 4h6a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-6"/></svg>',
  mark: '<svg viewBox="0 0 32 32" width="22" height="22"><path d="M9 23V9h14v14M9 14h14M9 18.5h14M14 9v14" stroke="#34d399" stroke-width="2" fill="none"/></svg>',
  alert: '<svg viewBox="0 0 24 24" width="15" height="15" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 9v4m0 4h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0"/></svg>',
};

/* ================= API (Supabase) ================= */
// Без ключей Supabase (или с ?local) сайт работает локально: Excel + PDF хранятся в IndexedDB браузера
const isLocalMode = () => CONFIG.MODE === 'local' || !CONFIG.SUPABASE_URL || /YOUR-PROJECT/.test(CONFIG.SUPABASE_URL) || new URLSearchParams(location.search).has('local');
async function makeApi() {
  if (window.__MOCK_API__) return window.__MOCK_API__;
  if (CONFIG.SHEETS_URL && !new URLSearchParams(location.search).has('local')) return makeSheetsApi(CONFIG.SHEETS_URL);
  if (CONFIG.STORE?.repo && !new URLSearchParams(location.search).has('local')) return makeCloudApi(CONFIG.STORE, async () => { if (!window.XLSX) await loadScript(LIB.xlsx); return window.XLSX; });
  if (isLocalMode()) return makeLocalApi(async () => { if (!window.XLSX) await loadScript(LIB.xlsx); return window.XLSX; });
  const { createClient } = await import(LIB.supabase);
  const sb = createClient(CONFIG.SUPABASE_URL, CONFIG.SUPABASE_ANON_KEY, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true } });
  const ok = r => { if (r.error) throw r.error; return r.data; };
  return {
    async getUser() { const { data } = await sb.auth.getSession(); return data.session?.user ?? null; },
    onAuth(cb) { sb.auth.onAuthStateChange((ev, s) => cb(ev, s?.user ?? null)); },
    async signIn(email, password) { ok(await sb.auth.signInWithPassword({ email, password })); },
    async signOut() { await sb.auth.signOut(); },
    async resetPassword(email) { ok(await sb.auth.resetPasswordForEmail(email, { redirectTo: location.origin + location.pathname })); },
    async updatePassword(password) { ok(await sb.auth.updateUser({ password })); },
    async profile(uid) { return ok(await sb.from('reg_members').select('*').eq('user_id', uid).maybeSingle()); },
    async loadAll() {
      const [sections, companies, shareholders, beneficiaries, extracts] = await Promise.all([
        sb.from('reg_sections').select('*').order('sort_order'),
        sb.from('reg_companies_view').select('*'),
        sb.from('reg_shareholders').select('*').order('sort_order'),
        sb.from('reg_beneficiaries').select('*').order('sort_order'),
        sb.from('reg_extracts').select('*').order('extract_date', { ascending: false }),
      ]).then(rs => rs.map(ok));
      return { sections, companies, shareholders, beneficiaries, extracts };
    },
    async signedUrl(path, downloadName) {
      return ok(await sb.storage.from('reg-extracts').createSignedUrl(path, 600, downloadName ? { download: downloadName } : undefined)).signedUrl;
    },
    async uploadExtract(company, file, date) {
      const path = `${company.code}/${date}_${Date.now().toString(36)}.pdf`;
      ok(await sb.storage.from('reg-extracts').upload(path, file, { contentType: 'application/pdf', upsert: false }));
      const { data: { user } } = await sb.auth.getUser();
      const r = await sb.from('reg_extracts').insert({ company_id: company.id, extract_date: date, storage_path: path, file_name: file.name, file_size: file.size, uploaded_by: user?.id }).select().single();
      if (r.error) { await sb.storage.from('reg-extracts').remove([path]); throw r.error; }
      return r.data;
    },
    async deleteExtract(ex) { ok(await sb.from('reg_extracts').delete().eq('id', ex.id)); await sb.storage.from('reg-extracts').remove([ex.storage_path]); },
    async saveCompany(id, patch) {
      return id ? ok(await sb.from('reg_companies').update(patch).eq('id', id).select().single())
                : ok(await sb.from('reg_companies').insert(patch).select().single());
    },
    async replaceRows(table, companyId, rows) {
      ok(await sb.from(table).delete().eq('company_id', companyId));
      if (rows.length) ok(await sb.from(table).insert(rows.map((r, i) => ({ ...r, company_id: companyId, sort_order: i }))));
    },
    async audit(rowIds, limit = 300) {
      let q = sb.from('reg_audit_log').select('*').order('created_at', { ascending: false }).limit(limit);
      if (rowIds) q = q.in('row_id', rowIds);
      return ok(await q);
    },
    async profiles() { return ok(await sb.from('reg_members').select('*').order('created_at')); },
    async setRole(userId, role) { ok(await sb.from('reg_members').update({ role }).eq('user_id', userId)); },
    async addMember(email, role) { return ok(await sb.rpc('reg_add_member', { p_email: email, p_role: role })); },
    async removeMember(userId) { ok(await sb.from('reg_members').delete().eq('user_id', userId)); },
  };
}

/* ================= State ================= */
const S = {
  api: null, user: null, profile: null, tab: store.get('reg.tab', 'registry'),
  sections: [], companies: [], byId: new Map(), byCode: new Map(), extracts: [],
  q: '', terms: [], section: null, flag: null, view: store.get('reg.view', 'compact'), sort: { key: 'code', dir: 1 },
  openId: null, editing: false, collapsed: new Set(),
};
const canEdit = () => ['editor', 'admin'].includes(S.profile?.role);
const isAdmin = () => S.profile?.role === 'admin';

function prepare(d) {
  S.sections = d.sections;
  S.extracts = d.extracts;
  const sh = groupBy(d.shareholders, 'company_id'), bn = groupBy(d.beneficiaries, 'company_id'), ex = groupBy(d.extracts, 'company_id');
  const list = d.companies.map(c => ({ ...c, sort_key: c.sort_key || codeKey(c.code), sh: sh.get(c.id) || [], bn: bn.get(c.id) || [], ex: ex.get(c.id) || [] }));
  list.sort((a, b) => cmpKey(a.sort_key, b.sort_key));
  S.companies = list; S.byId = new Map(list.map(c => [c.id, c])); S.byCode = new Map(list.map(c => [c.code, c]));
  const owner = normName(CONFIG.ULTIMATE_OWNER);
  const memo = new Map();
  const eff = (c, stack = new Set()) => {
    if (memo.has(c.id)) return memo.get(c.id);
    if (stack.has(c.id)) return 0; stack.add(c.id);
    let v = 0;
    for (const s of c.sh) {
      if (normName(s.name).includes(owner)) v += Number(s.share_pct || 0);
      else if (s.owner_company_id && S.byId.has(s.owner_company_id)) v += Number(s.share_pct || 0) * eff(S.byId.get(s.owner_company_id), stack) / 100;
    }
    memo.set(c.id, v); return v;
  };
  for (const c of list) {
    c.depth = c.code === '0' ? 0 : c.sort_key.length;
    c.groupDirect = c.sh.filter(s => s.owner_company_id).reduce((a, s) => a + Number(s.share_pct || 0), 0);
    c.eff = round2(eff(c));
    const st = c.bn.find(b => normName(b.name).includes(owner));
    c.ownerStated = st ? Number(st.share_pct) : null;
    c.type = c.code === '0' ? 'root' : c.groupDirect >= 99.99 ? 'full' : 'partial';
    c.children = list.filter(x => x.parent_code === c.code);
    c.issues = CHECKS.map(ch => { const det = ch.test(c); return det ? { ch, det } : null; }).filter(Boolean).sort((a, b) => SEV_ORDER[a.ch.sev] - SEV_ORDER[b.ch.sev]);
    c._search = norm([c.code, c.name, c.rma, c.director, c.position, c.email, c.phone, c.activity, c.legal_address, c.founders, c.beneficiary, c.supervisory_board, c.executive, c.notes].join(' | '));
  }
}
function groupBy(arr, k) { const m = new Map(); for (const x of arr) { if (!m.has(x[k])) m.set(x[k], []); m.get(x[k]).push(x); } return m; }

const FLAGS = {
  full: c => c.type === 'full', partial: c => c.type === 'partial', unformed: c => c.issues.some(i => i.ch.id === 'unformed'),
  nofile: c => !c.extract_path, old: c => c.issues.some(i => i.ch.id === 'old'), issues: c => c.issues.some(i => i.ch.sev === 'error'),
};
function filtered() {
  return S.companies.filter(c =>
    (!S.section || c.section_id === S.section) &&
    (!S.flag || FLAGS[S.flag](c)) &&
    S.terms.every(term => c._search.includes(term)));
}

/* ================= Boot / Auth ================= */
async function boot() {
  applyTheme();
  try { S.api = await makeApi(); } catch (e) { $('#app').innerHTML = `<div class="boot"><div class="form-error">${esc(t('loadErr'))}: ${esc(e.message)}</div></div>`; return; }
  if (S.api.local) { if (S.api.hasData()) { S.user = await S.api.getUser(); await enter(); } else renderWelcome(); return; }
  let recovering = false;
  S.api.onAuth(async (ev, user) => {
    if (ev === 'PASSWORD_RECOVERY') { recovering = true; renderRecovery(); return; }
    if (ev === 'SIGNED_OUT') { S.user = null; S.profile = null; renderLogin(); }
  });
  S.user = await S.api.getUser();
  if (recovering) return;
  if (S.user) await enter(); else renderLogin();
}

function langSwitch(cls = '') {
  return `<div class="lang ${cls}"><button data-lang="tg" class="${lang === 'tg' ? 'on' : ''}">ТҶ</button><button data-lang="ru" class="${lang === 'ru' ? 'on' : ''}">РУ</button></div>`;
}
function bindLang(root, rerender) { $$('[data-lang]', root).forEach(b => b.onclick = () => { lang = b.dataset.lang; store.set('reg.lang', lang); document.documentElement.lang = lang === 'ru' ? 'ru' : 'tg'; rerender(); }); }

function renderWelcome(msg = '') {
  $('#app').innerHTML = `
  <div class="login"><div class="login-card" style="max-width:520px">
    <div style="display:flex;justify-content:space-between;align-items:center"><div class="brand-mark">${IC.mark}</div>${langSwitch()}</div>
    <h1>${esc(t('welcomeTitle'))}</h1><p>${esc(t('welcomeSub'))}</p>
    <label class="drop" id="wdrop">${t('welcomeDrop')}<div class="muted" style="font-size:12px;margin-top:6px">.xlsx + .pdf</div><input type="file" id="wfile" accept=".xlsx,.xls,application/pdf,.pdf" multiple hidden></label>
    <div class="form-error" id="werr">${esc(msg)}</div>
    <div class="login-foot"><span>🔒 ${esc(t('welcomePrivacy'))}</span></div>
  </div></div>`;
  $$('.login-card .lang button').forEach(b => b.style.color = b.classList.contains('on') ? '#fff' : 'var(--text-2)');
  $('.login-card .lang').style.borderColor = 'var(--line-strong)';
  bindLang($('#app'), () => renderWelcome());
  const drop = $('#wdrop'), inp = $('#wfile');
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); localImport([...e.dataTransfer.files]); };
  inp.onchange = () => localImport([...inp.files]);
}
async function localImport(files, fromShell) {
  const xlsx = files.find(f => /\.xlsx?$/i.test(f.name));
  const pdfs = files.filter(f => /\.pdf$/i.test(f.name) || f.type === 'application/pdf');
  if (!xlsx && !S.api.hasData()) { const e = $('#werr'); if (e) e.textContent = t('welcomeNeedXlsx'); return; }
  try {
    if (xlsx) { const r = await S.api.importExcel(xlsx); toast(t('imported', r.count)); }
    if (fromShell) await reload(); else { S.user = await S.api.getUser(); await enter(); }
    if (pdfs.length) openImport(pdfs, true);
  } catch (e) { const el = $('#werr'); if (el) el.textContent = e.message; else toast(e.message, true); }
}

function renderLogin(msg = '') {
  $('#app').innerHTML = `
  <div class="login"><form class="login-card" id="login-form" autocomplete="on">
    <div style="display:flex;justify-content:space-between;align-items:center"><div class="brand-mark">${IC.mark}</div>${langSwitch()}</div>
    <h1>${esc(t('appTitle'))}</h1><p>${esc(t('loginSub'))}</p>
    ${S.api.keyLogin ? '' : `<div class="field"><label for="em">${esc(t('email'))}</label><input class="input" id="em" type="email" autocomplete="username" required></div>`}
    <div class="field"><label for="pw">${esc(S.api.cloud ? t('password') : S.api.keyLogin ? t('accessKey') : t('password'))}</label><input class="input" id="pw" type="password" autocomplete="current-password" required></div>
    <div class="form-error" id="lerr">${esc(msg)}</div>
    <button class="btn primary block" type="submit" id="lbtn">${esc(t('login'))}</button>
    <div class="login-foot">${S.api.keyLogin ? `<span>${esc(t('keyHint'))}</span>` : `<a href="#" id="forgot">${esc(t('forgot'))}</a>`}<span>Авесто Гуруҳ</span></div>
  </form></div>`;
  const lang_s = $('.login-card .lang'); lang_s.style.cssText = 'border-color:var(--line-strong)';
  $$('.login-card .lang button').forEach(b => b.style.color = b.classList.contains('on') ? '#fff' : 'var(--text-2)');
  bindLang($('#app'), () => renderLogin());
  $('#login-form').onsubmit = async e => {
    e.preventDefault(); $('#lbtn').disabled = true; $('#lerr').textContent = '';
    try { await S.api.signIn($('#em')?.value.trim(), $('#pw').value); S.user = await S.api.getUser(); await enter(); }
    catch (err) { $('#lerr').textContent = err.message || String(err); $('#lbtn').disabled = false; }
  };
  if ($('#forgot')) $('#forgot').onclick = async e => {
    e.preventDefault(); const em = $('#em').value.trim(); if (!em) { $('#em').focus(); return; }
    try { await S.api.resetPassword(em); $('#lerr').style.color = 'var(--accent)'; $('#lerr').textContent = t('resetSent'); } catch (err) { $('#lerr').textContent = err.message; }
  };
}
function renderRecovery() {
  $('#app').innerHTML = `<div class="login"><form class="login-card" id="rf"><div class="brand-mark">${IC.mark}</div><h1>${esc(t('newPassword'))}</h1><p></p>
    <div class="field"><input class="input" id="npw" type="password" minlength="8" autocomplete="new-password" required></div>
    <div class="form-error" id="rerr"></div><button class="btn primary block">${esc(t('savePassword'))}</button></form></div>`;
  $('#rf').onsubmit = async e => { e.preventDefault(); try { await S.api.updatePassword($('#npw').value); history.replaceState(null, '', location.pathname); S.user = await S.api.getUser(); await enter(); } catch (err) { $('#rerr').textContent = err.message; } };
}

async function enter() {
  $('#app').innerHTML = '<div class="boot"><div class="spinner"></div></div>';
  try {
    S.profile = await S.api.profile(S.user.id);
    if (!S.profile || !['viewer', 'editor', 'admin'].includes(S.profile.role)) { await S.api.signOut(); renderLogin(t('noAccess')); return; }
    prepare(await S.api.loadAll());
  } catch (e) { renderLogin(`${t('loadErr')}: ${e.message}`); return; }
  renderShell();
  const m = location.hash.match(/c=([\d.]+)/); if (m && S.byCode.has(m[1])) openDrawer(S.byCode.get(m[1]).id);
}
async function reload() { prepare(await S.api.loadAll()); update(); if (S.openId) renderDrawer(); }

/* ================= Shell ================= */
function renderShell() {
  document.documentElement.lang = lang === 'ru' ? 'ru' : 'tg';
  document.title = `${t('appTitle')} · Авесто Гуруҳ`;
  const who = S.api.cloud ? 'Авесто Гуруҳ' : S.api.local ? t('localMode') : (S.profile.full_name || S.profile.email || S.user.email);
  const tabs = [['registry', t('tabRegistry')], ['tree', t('tabTree')], ['checks', t('tabChecks')]];
  if (canEdit()) tabs.push(['audit', t('tabAudit')]);
  if (isAdmin()) tabs.push(['users', t('tabUsers')]);
  if (!tabs.some(x => x[0] === S.tab)) S.tab = 'registry';
  $('#app').innerHTML = `
  <header class="topbar"><div class="topbar-in">
    <div class="brand"><div class="brand-mark">${IC.mark}</div><div><div class="brand-title">${esc(t('appTitle'))}</div><div class="brand-sub">${esc(t('appSub'))}</div></div></div>
    <nav class="tabs">${tabs.map(([k, l]) => `<button class="tab ${S.tab === k ? 'active' : ''}" data-tab="${k}">${esc(l)}${k === 'checks' ? `<span class="count" id="chk-count"></span>` : ''}</button>`).join('')}</nav>
    <div class="spacer"></div>
    <div class="top-actions">${langSwitch()}
      <button class="icon-btn" id="theme" title="${esc(t('themeToggle'))}">${IC.moon}</button>
      <div class="user-chip"><div class="avatar">${esc(initials(who))}</div><div class="nm"><div style="font-weight:600;line-height:1.1">${esc(who)}</div><div class="role-badge">${S.api.cloud ? esc(canEdit() ? t('rEditor') : t('rViewer')) : S.api.local ? esc(t('localHint')) : esc(t('r' + S.profile.role[0].toUpperCase() + S.profile.role.slice(1)))}</div></div>
      ${S.api.cloud ? `<button class="icon-btn" id="editkey" title="${esc(t('editMode'))}" style="${S.api.hasToken() ? 'color:#6ee7b7' : ''}">${IC.key}</button>` : ''}
      ${S.api.local ? `<button class="icon-btn" id="clear" title="${esc(t('clearData'))}">${IC.trash}</button>` : `<button class="icon-btn" id="logout" title="${esc(t('logout'))}">${IC.out}</button>`}</div>
    </div></div></header>
  <main class="page">
    <section class="kpis" id="kpis"></section>
    <div id="registry-tools">
      <div class="toolbar">
        <div class="search">${IC.search}<input class="input" id="q" placeholder="${esc(t('search'))}" value="${esc(S.q)}" autocomplete="off"><kbd id="kbd">/</kbd><button class="icon-btn clear hidden" id="qclear">${IC.x}</button></div>
        <div class="seg" id="viewseg"><button data-view="compact" class="${S.view === 'compact' ? 'on' : ''}">${esc(t('compact'))}</button><button data-view="full" class="${S.view === 'full' ? 'on' : ''}">${esc(t('full'))}</button></div>
        <span class="result-count" id="rc"></span>
        <div class="spacer"></div>
        <button class="btn" id="exp">${IC.xls}${esc(t('exportXlsx'))}</button>
        ${S.api.local || (S.api.cloud && canEdit()) ? `<button class="btn" id="reimp">${IC.xls}${esc(t('reimport'))}</button><input type="file" id="reimpf" accept=".xlsx,.xls,.pdf" multiple hidden>` : ''}
        ${S.api.sheets && canEdit() && S.api._folder ? `<a class="btn" href="${esc(S.api._folder)}" target="_blank" rel="noopener">${IC.ext}${esc(t('pdfFolder'))}</a>` : ''}
        ${canEdit() ? `<button class="btn" id="imp">${IC.up}${esc(t('importPdf'))}</button><button class="btn primary" id="newc">${IC.plus}${esc(t('newCompany'))}</button>` : ''}
      </div>
      <div class="chips" id="chips"></div>
    </div>
    <section id="view"></section>
    <div class="footer">Разработано <b>Отделом автоматизации</b> холдинга ООО «Авесто Групп» · <span class="isj">ISJ</span></div>
  </main>`;
  bindLang($('.topbar'), () => { renderShell(); if (S.openId) renderDrawer(); });
  $$('[data-tab]').forEach(b => b.onclick = () => { S.tab = b.dataset.tab; store.set('reg.tab', S.tab); $$('[data-tab]').forEach(x => x.classList.toggle('active', x === b)); update(); });
  if (S.api.local) $('#clear').onclick = async e => { const b = e.currentTarget; if (!b.dataset.armed) { b.dataset.armed = 1; toast(t('confirmClear'), true); setTimeout(() => delete b.dataset.armed, 5000); return; } await S.api.clear(); closeDrawer(); renderWelcome(); };
  else $('#logout').onclick = async () => { await S.api.signOut(); S.user = null; renderLogin(); };
  if (S.api.cloud) $('#editkey').onclick = openTokenDialog;
  if ($('#reimp')) { $('#reimp').onclick = () => $('#reimpf').click(); $('#reimpf').onchange = e => { localImport([...e.target.files], true); e.target.value = ''; }; }
  $('#theme').onclick = () => { const cur = store.get('reg.theme', ''); const next = cur === 'dark' ? 'light' : cur === 'light' ? '' : 'dark'; store.set('reg.theme', next); applyTheme(); };
  const q = $('#q');
  q.oninput = () => { S.q = q.value; S.terms = normName(S.q).split(' ').filter(Boolean); update(); };
  $('#qclear').onclick = () => { q.value = ''; q.oninput(); q.focus(); };
  $$('[data-view]').forEach(b => b.onclick = () => { S.view = b.dataset.view; store.set('reg.view', S.view); $$('[data-view]').forEach(x => x.classList.toggle('on', x === b)); update(); });
  $('#exp').onclick = exportXlsx;
  if (canEdit()) { $('#imp').onclick = openImport; $('#newc').onclick = () => { S.openId = 'new'; S.editing = true; renderDrawer(); }; }
  update();
}
function applyTheme() { const th = store.get('reg.theme', ''); if (th) document.documentElement.dataset.theme = th; else delete document.documentElement.dataset.theme; }

function update() {
  const rows = filtered();
  $('#registry-tools').classList.toggle('hidden', S.tab !== 'registry');
  $('#kpis').classList.toggle('hidden', !['registry', 'tree'].includes(S.tab));
  renderKpis(); renderChips();
  $('#rc').textContent = t('found', rows.length, S.companies.length);
  $('#qclear').classList.toggle('hidden', !S.q); $('#kbd').classList.toggle('hidden', !!S.q);
  const errN = S.companies.reduce((a, c) => a + c.issues.filter(i => i.ch.sev === 'error').length, 0);
  $('#chk-count').textContent = errN || ''; $('#chk-count').classList.toggle('hidden', !errN);
  const v = $('#view');
  if (S.tab === 'registry') v.innerHTML = renderTable(rows);
  else if (S.tab === 'tree') v.innerHTML = renderTree();
  else if (S.tab === 'checks') v.innerHTML = renderChecks();
  else if (S.tab === 'audit') renderAudit(v);
  else if (S.tab === 'users') renderUsers(v);
  bindView();
}

/* ================= KPIs & chips ================= */
function renderKpis() {
  const cs = S.companies, n = cs.length;
  const decl = cs.reduce((a, c) => a + Number(c.capital_declared || 0), 0), formed = cs.reduce((a, c) => a + Number(c.capital_formed || 0), 0);
  const files = cs.filter(c => c.extract_path).length;
  const cnt = f => cs.filter(FLAGS[f]).length;
  const pct = decl ? Math.round(formed / decl * 1000) / 10 : 0;
  const k = (flag, label, value, sub, cls = '', extra = '') => `<button class="kpi ${cls} ${S.flag === flag && flag ? 'on' : ''}" data-flag="${flag || ''}"><div class="kpi-label">${label}</div><div class="kpi-value">${value}</div><div class="kpi-sub">${sub}</div>${extra}</button>`;
  $('#kpis').innerHTML = [
    k('', esc(t('kCompanies')), n, esc(t('all'))),
    k('full', `<span class="dot" style="background:var(--accent)"></span>${esc(t('kFull'))}`, cnt('full'), esc(t('kOf', n))),
    k('partial', `<span class="dot" style="background:var(--warn)"></span>${esc(t('kPartial'))}`, cnt('partial'), esc(t('kOf', n))),
    k('unformed', esc(t('kCapital')), fmtShort(decl), `${pct}% ${esc(t('kFormed'))} · ${cnt('unformed')} ${esc(t('kUnformed')).toLowerCase()}`, '', `<div class="bar"><i style="width:${pct}%"></i></div>`),
    k('nofile', esc(t('kExtracts')), `${files}<span style="font-size:14px;color:var(--text-3)">/${n}</span>`, esc(`${n - files} ${t('noFile')}`), '', `<div class="bar"><i style="width:${n ? files / n * 100 : 0}%"></i></div>`),
    k('old', esc(t('kOld')), cnt('old'), esc(t('kOldSub', CONFIG.EXTRACT_MAX_AGE_DAYS)), cnt('old') ? 'warnk' : ''),
    k('issues', esc(t('kIssues')), cnt('issues'), esc(t('tabChecks')), cnt('issues') ? 'alert' : ''),
  ].join('');
  $$('#kpis [data-flag]').forEach(b => b.onclick = () => { const f = b.dataset.flag || null; S.flag = S.flag === f ? null : f; if (f) S.section = null; if (!f) { S.section = null; S.flag = null; } S.tab = S.tab === 'tree' ? 'registry' : S.tab; $$('[data-tab]').forEach(x => x.classList.toggle('active', x.dataset.tab === S.tab)); update(); });
}
function renderChips() {
  const c = S.companies;
  const chip = (id, label, n) => `<button class="chip ${S.section === id ? 'on' : ''}" data-sec="${id ?? ''}">${esc(label)} <span class="n">${n}</span></button>`;
  const secLabel = s => s.title.replace(/^ҶАМЪИЯТҲОИ\s+/i, '').replace(/^ФАРЪИИ\s+/i, '').replace(/^ФАРЪӢ ВА ВОБАСТАИ\s+/i, '').replace(/ҶДММ\s+|ҶСП\s+/g, '');
  $('#chips').innerHTML = chip(null, t('all'), c.length) + S.sections.map(s => chip(s.id, secLabel(s), c.filter(x => x.section_id === s.id).length)).join('')
    + (S.flag ? `<button class="chip active-filter" data-clearflag>${esc(kpiName(S.flag))} ✕</button>` : '');
  $$('#chips [data-sec]').forEach(b => b.onclick = () => { S.section = b.dataset.sec ? Number(b.dataset.sec) : null; S.flag = null; update(); });
  const cf = $('#chips [data-clearflag]'); if (cf) cf.onclick = () => { S.flag = null; update(); };
}
const kpiName = f => ({ full: t('kFull'), partial: t('kPartial'), unformed: t('kUnformed'), nofile: t('noFile'), old: t('kOld'), issues: t('kIssues') }[f]);

/* ================= Registry table ================= */
const shareSummary = (rows, terms) => rows.map(s => `<div>${hl(s.name, terms)} <b class="mono">${round2(s.share_pct)}%</b></div>`).join('');
function extractCell(c) {
  if (c.extract_path) {
    const old = daysSince(c.extract_date) > CONFIG.EXTRACT_MAX_AGE_DAYS;
    return `<button class="extract-link ${old ? 'age-old' : ''}" data-pdf="${c.id}" title="${esc(t('openPdf'))}${c.extract_file_name ? ': ' + esc(c.extract_file_name) : ''}">${IC.pdf}${esc(fmtDate(c.extract_date))}</button>`;
  }
  return `<span class="extract-none">${esc(fmtDate(c.extract_date) || '—')}<span class="hint">${esc(t('noFile'))}</span></span>`;
}
const COLS = [
  { k: 'code', l: 'cCode', cls: 'c-code', sort: c => c.sort_key, r: (c, T) => hl(c.code === '0' ? '—' : c.code, T) },
  { k: 'name', l: 'cName', cls: 'c-name', sort: c => normName(c.name), r: (c, T) => `<div class="name">${'<span class="depth" style="width:' + Math.max(0, (c.depth - 1) * 10) + 'px"></span>'}${hl(c.name, T)}${c.issues.length ? `<span class="flags">${[...new Set(c.issues.map(i => i.ch.sev))].map(s => `<i class="flag ${s}"></i>`).join('')}</span>` : ''}</div><div class="sub">${esc(c.type === 'root' ? t('typeRoot') : c.type === 'full' ? t('typeFull') : `${t('groupShare')} ${round2(c.groupDirect)}%`)}</div>` },
  { k: 'rma', l: 'cRma', cls: 'mono nowrap', sort: c => c.rma, r: (c, T) => hl(c.rma, T) },
  { k: 'founded', l: 'cFounded', cls: 'mono nowrap', sort: c => c.founded_on || '', r: c => `${esc(fmtDate(c.founded_on))}${c.reregistered_on ? `<div class="sub">(${esc(fmtDate(c.reregistered_on))})</div>` : ''}` },
  { k: 'address', l: 'cAddress', full: true, r: (c, T) => `<div class="clamp">${hl(c.legal_address, T)}</div>` },
  { k: 'activity', l: 'cActivity', full: true, r: (c, T) => `<div class="clamp">${hl(c.activity, T)}</div>` },
  { k: 'declared', l: 'cDeclared', cls: 'num', sort: c => Number(c.capital_declared || 0), r: c => { const bad = c.capital_declared > 0 && (c.capital_formed ?? 0) < c.capital_declared; return `<span class="cap">${fmtNum(c.capital_declared)}</span>`; }, td: c => (c.capital_declared > 0 && (c.capital_formed ?? 0) < c.capital_declared ? 'cap bad' : 'cap') },
  { k: 'formed', l: 'cFormed', cls: 'num', sort: c => Number(c.capital_formed || 0), r: c => { const bad = c.capital_declared > 0 && (c.capital_formed ?? 0) < c.capital_declared; return `${fmtNum(c.capital_formed)}${bad ? `<div class="cap-pct">${Math.round((c.capital_formed || 0) / c.capital_declared * 100)}%</div>` : ''}`; }, td: c => (c.capital_declared > 0 && (c.capital_formed ?? 0) < c.capital_declared ? 'cap bad' : 'cap') },
  { k: 'founders', l: 'cFounders', r: (c, T) => S.view === 'full' ? `<div class="clamp w">${hl(c.founders, T)}</div>` : `<div class="clamp">${shareSummary(c.sh, T)}</div>` },
  { k: 'exec', l: 'cExecutive', r: (c, T) => S.view === 'full' ? `<div class="clamp w">${hl(c.executive, T)}</div>` : `<div class="clamp"><b>${hl(c.director, T)}</b><div class="sub">${hl(c.position, T)}</div><div class="sub">${hl([c.email, c.phone].filter(Boolean).join(' · '), T)}</div></div>` },
  { k: 'board', l: 'cBoard', full: true, r: (c, T) => `<div class="clamp">${hl(c.supervisory_board || '—', T)}</div>` },
  { k: 'benef', l: 'cBenef', r: (c, T) => S.view === 'full' ? `<div class="clamp w">${hl(c.beneficiary || '—', T)}</div>` : (c.bn.length ? `<div class="clamp">${shareSummary(c.bn, T)}</div>` : '<span class="muted">—</span>') },
  { k: 'extract', l: 'cExtract', sort: c => c.extract_date || '', r: c => extractCell(c) },
];
function renderTable(rows) {
  const cols = COLS.filter(c => !c.full || S.view === 'full');
  const T = S.terms;
  const sortCol = COLS.find(c => c.k === S.sort.key);
  const byCode = S.sort.key === 'code';
  if (!byCode && sortCol?.sort) rows = [...rows].sort((a, b) => { const x = sortCol.sort(a), y = sortCol.sort(b); return (Array.isArray(x) ? cmpKey(x, y) : x < y ? -1 : x > y ? 1 : 0) * S.sort.dir; });
  else if (S.sort.dir < 0) rows = [...rows].reverse();
  const head = cols.map(c => `<th class="${c.cls || ''} ${c.sort ? '' : 'nosort'} ${S.sort.key === c.k ? 'sorted' : ''}" data-sort="${c.sort ? c.k : ''}">${esc(t(c.l))}${c.sort ? `<span class="sort">${S.sort.key === c.k ? (S.sort.dir > 0 ? '▲' : '▼') : '↕'}</span>` : ''}</th>`).join('');
  let body = '', lastSec; const shown = new Set();
  for (const c of rows) {
    // как в Excel: заголовок раздела — один раз, при первом появлении; после блока дочерних не повторяется
    if (byCode && S.sort.dir > 0 && c.section_id !== lastSec) {
      lastSec = c.section_id; const sec = S.sections.find(s => s.id === c.section_id);
      if (sec && c.code !== '0' && !shown.has(sec.id)) { shown.add(sec.id); body += `<tr class="section ${c.depth > 1 ? 'subsec' : ''}"><td colspan="${cols.length}">${esc(sec.title)}<span class="n">${rows.filter(x => x.section_id === sec.id).length}</span></td></tr>`; }
    }
    body += `<tr class="row ${S.openId === c.id ? 'sel' : ''}" data-open="${c.id}">${cols.map(col => `<td class="${col.cls || ''} ${col.td ? col.td(c) : ''}">${col.r(c, T)}</td>`).join('')}</tr>`;
  }
  return `<div class="card"><div class="table-wrap"><table class="reg"><thead><tr>${head}</tr></thead><tbody>${body || `<tr><td colspan="${cols.length}" class="empty">${esc(t('noRows'))}</td></tr>`}</tbody></table></div></div>`;
}

/* ================= Structure tree ================= */
function renderTree() {
  const root = S.byCode.get('0');
  if (!root) return `<div class="card empty">—</div>`;
  const shareCls = v => v >= 99.99 ? '' : v >= 50 ? 'partial' : 'minor';
  const node = c => {
    const kids = c.children; const col = S.collapsed.has(c.id);
    const direct = c.code === '0' ? null : c.groupDirect;
    return `<li><div class="node ${c.code === '0' ? 'root' : ''}" data-open="${c.id}">
      ${kids.length ? `<button class="toggle" data-toggle="${c.id}">${col ? '▶' : '▼'}</button>` : ''}
      <span class="code">${esc(c.code === '0' ? '' : c.code)}</span>
      <span><div class="nm">${esc(c.name)}</div><div class="meta">${esc(fmtShort(c.capital_declared))} · РМА ${esc(c.rma)}${kids.length ? ` · ${kids.length} ${esc(t('subsidiaries')).toLowerCase()}` : ''}</div></span>
      ${direct != null ? `<span class="share ${shareCls(direct)}" title="${esc(t('groupOwned'))}">${round2(direct)}%</span>` : ''}
      ${c.code !== '0' && Math.abs(c.eff - direct) > .01 ? `<span class="share minor" title="${esc(t('effShare'))}">eff ${round2(c.eff)}%</span>` : ''}
      ${c.extract_path ? `<button class="icon-btn" data-pdf="${c.id}" title="${esc(t('openPdf'))} ${esc(fmtDate(c.extract_date))}">${IC.pdf}</button>` : ''}
      ${c.issues.some(i => i.ch.sev === 'error') ? `<i class="flag error" title="${esc(t('kIssues'))}"></i>` : ''}
    </div>${kids.length && !col ? `<ul>${kids.map(node).join('')}</ul>` : ''}</li>`;
  };
  const orphans = S.companies.filter(c => c.code !== '0' && !S.byCode.has(c.parent_code));
  return `<div class="card"><div class="legend"><span><span class="share">100%</span> ${esc(t('legendFull'))}</span><span><span class="share partial">50–99%</span> ${esc(t('legendPartial'))}</span><span><span class="share minor">&lt;50%</span> ${esc(t('legendMinor'))}</span><span>eff — ${esc(t('effShare'))} (${esc(CONFIG.ULTIMATE_OWNER)})</span></div>
    <div class="tree"><ul>${node(root)}${orphans.map(node).join('')}</ul></div></div>`;
}

/* ================= Checks ================= */
function renderChecks() {
  const blocks = CHECKS.map(ch => {
    const hits = S.companies.map(c => ({ c, i: c.issues.find(x => x.ch.id === ch.id) })).filter(x => x.i);
    if (!hits.length) return '';
    return `<div class="card check"><h3><span class="sev ${ch.sev}">${ch.sev}</span>${esc(chkTitle(ch))}<span class="muted mono" style="margin-left:auto">${hits.length}</span></h3><p>${esc(chkDesc(ch))}</p>
      <ul>${hits.map(({ c, i }) => `<li data-open="${c.id}"><span class="code">${esc(c.code === '0' ? '—' : c.code)}</span><span class="nm">${esc(c.name)}</span><span class="det">${esc(i.det)}</span></li>`).join('')}</ul></div>`;
  }).join('');
  return blocks ? `<div class="checks">${blocks}</div>` : `<div class="card ok-all">✓ ${esc(t('chkNone'))}</div>`;
}

/* ================= Audit & users ================= */
async function renderAudit(v) {
  v.innerHTML = '<div class="card empty"><div class="spinner" style="margin:auto"></div></div>';
  try {
    const rows = await S.api.audit(null, 300);
    v.innerHTML = `<div class="card"><div class="table-wrap"><table class="reg"><thead><tr><th>${esc(t('aTime'))}</th><th>${esc(t('aUser'))}</th><th>${esc(t('aAction'))}</th><th>${esc(t('aObject'))}</th><th>${esc(t('aFields'))}</th></tr></thead><tbody>
      ${rows.map(a => { const cid = a.table_name === 'reg_companies' ? a.row_id : (a.new_data || a.old_data)?.company_id; const c = S.byId.get(cid);
        return `<tr class="${c ? 'row' : ''}" ${c ? `data-open="${c.id}"` : ''}><td class="mono nowrap">${esc(new Date(a.created_at).toLocaleString('ru-RU'))}</td><td>${esc(a.user_email || '—')}</td><td><span class="pill mono">${esc(a.table_name)} · ${esc(a.action)}</span></td><td>${esc(c ? c.name : (a.new_data || a.old_data)?.name || '—')}</td><td class="muted">${esc(diffKeys(a).join(', '))}</td></tr>`; }).join('') || `<tr><td colspan="5" class="empty">—</td></tr>`}
    </tbody></table></div></div>`;
    bindView();
  } catch (e) { v.innerHTML = `<div class="card empty">${esc(e.message)}</div>`; }
}
function diffKeys(a) { if (a.fields != null) return [a.fields]; if (!a.old_data || !a.new_data) return a.table_name === 'reg_extracts' ? [fmtDate((a.new_data || a.old_data).extract_date)] : []; return Object.keys(a.new_data).filter(k => !['updated_at', 'updated_by'].includes(k) && JSON.stringify(a.new_data[k]) !== JSON.stringify(a.old_data[k])); }
const roleOpts = sel => ['viewer', 'editor', 'admin'].map(r => `<option value="${r}" ${sel === r ? 'selected' : ''}>${esc(t('r' + r[0].toUpperCase() + r.slice(1)))}</option>`).join('');
async function renderUsers(v) {
  v.innerHTML = '<div class="card empty"><div class="spinner" style="margin:auto"></div></div>';
  try {
    const ps = await S.api.profiles();
    v.innerHTML = `<div class="card"><div style="padding:12px 16px;color:var(--text-2);font-size:13px;border-bottom:1px solid var(--line)">${esc(t('usersHint'))}</div><table class="mini" style="margin:6px 0"><thead><tr><th style="padding-left:16px">${esc(t('email'))}</th><th>${esc(t('role'))}</th></tr></thead><tbody>
      ${ps.map(p => `<tr><td style="padding-left:16px">${esc(p.email)}${p.user_id === S.user.id ? ' <span class="badge-cur">you</span>' : ''}</td><td><select class="input" style="width:auto" data-role="${p.user_id}" ${p.user_id === S.user.id ? 'disabled' : ''}>${roleOpts(p.role)}</select> ${p.user_id === S.user.id ? '' : `<button class="btn danger" data-rm="${p.user_id}">${esc(t('del'))}</button>`}</td></tr>`).join('')}
    </tbody></table>
    <div style="display:flex;gap:8px;flex-wrap:wrap;padding:12px 16px;border-top:1px solid var(--line)"><input class="input" id="newem" type="email" placeholder="${esc(t('email'))}" style="flex:1 1 240px"><select class="input" id="newrole" style="width:auto">${roleOpts('viewer')}</select><button class="btn primary" id="addm">${IC.plus}${esc(t('addMember'))}</button></div></div>`;
    $$('[data-role]', v).forEach(s => s.onchange = async () => { try { await S.api.setRole(s.dataset.role, s.value); toast(t('saved')); } catch (e) { toast(e.message, true); } });
    $$('[data-rm]', v).forEach(b => b.onclick = async () => { if (!confirmInline(b)) return; try { await S.api.removeMember(b.dataset.rm); toast(t('deleted')); renderUsers(v); } catch (e) { toast(e.message, true); } });
    $('#addm', v).onclick = async () => { const em = $('#newem', v).value.trim(); if (!em) return; try { await S.api.addMember(em, $('#newrole', v).value); toast(t('saved')); renderUsers(v); } catch (e) { toast(e.message, true); } };
  } catch (e) { v.innerHTML = `<div class="card empty">${esc(e.message)}</div>`; }
}

/* ================= View bindings ================= */
function bindView() {
  const v = $('#view');
  $$('[data-pdf]', v).forEach(b => b.onclick = e => { e.stopPropagation(); const c = S.byId.get(b.dataset.pdf); openPdf(c, c.ex.find(x => x.storage_path === c.extract_path) || { storage_path: c.extract_path, extract_date: c.extract_date, file_name: c.extract_file_name }); });
  $$('[data-toggle]', v).forEach(b => b.onclick = e => { e.stopPropagation(); const id = b.dataset.toggle; S.collapsed.has(id) ? S.collapsed.delete(id) : S.collapsed.add(id); update(); });
  $$('[data-open]', v).forEach(r => r.onclick = () => openDrawer(r.dataset.open));
  $$('th[data-sort]', v).forEach(th => th.onclick = () => { const k = th.dataset.sort; if (!k) return; if (S.sort.key === k) S.sort.dir *= -1; else S.sort = { key: k, dir: 1 }; update(); });
}

/* ================= Режим правки (облако) ================= */
function openTokenDialog() {
  const on = S.api.hasToken();
  const m = modal(`<div class="modal-card" style="width:min(560px,100%)"><div class="modal-head"><span style="color:var(--accent)">${IC.key}</span><h3>${esc(t('tokenTitle'))}</h3><button class="icon-btn" data-close>${IC.x}</button></div>
    <div class="modal-body">${on ? `<div class="issue info"><b>✓ ${esc(t('tokenOn'))}</b></div>` : ''}<p class="muted" style="margin-top:0">${esc(t('tokenHelp'))}</p>
      <input class="input mono" id="tok" type="password" placeholder="github_pat_…" autocomplete="off"><div class="form-error" id="tokerr"></div></div>
    <div class="modal-foot">${on ? `<button class="btn danger" id="tokoff">${esc(t('tokenOff'))}</button>` : ''}<button class="btn" data-close>${esc(t('close'))}</button><button class="btn primary" id="tokok">${esc(t('tokenSave'))}</button></div></div>`);
  const apply = async v => { try { $('#tokok', m).disabled = true; await S.api.setToken(v); S.profile = await S.api.profile(); m._close(); renderShell(); toast(v ? t('tokenOn') : t('saved')); } catch (e) { $('#tokerr', m).textContent = e.message; $('#tokok', m).disabled = false; } };
  $('#tokok', m).onclick = () => apply($('#tok', m).value);
  if (on) $('#tokoff', m).onclick = () => apply('');
}

/* ================= PDF viewer ================= */
async function openPdf(c, ex) {
  const m = modal(`<div class="modal-card wide"><div class="modal-head"><span style="color:var(--link)">${IC.pdf}</span><h3>${esc(c.name)} · ${esc(t('cExtract'))} ${esc(fmtDate(ex.extract_date))}</h3>
     <a class="btn" id="pnew" target="_blank" rel="noopener">${IC.ext}${esc(t('openNew'))}</a><a class="btn" id="pdl">${IC.down}${esc(t('download'))}</a><button class="icon-btn" data-close>${IC.x}</button></div>
     <div style="flex:1;min-height:0"><div class="boot" id="pload" style="min-height:100%"><div class="spinner"></div></div></div></div>`);
  try {
    const nice = `${c.code === '0' ? '0' : c.code}. ${c.name.replace(/["“”«»]/g, '')} ${fmtDate(ex.extract_date)}.pdf`;
    const [url, dl] = await Promise.all([S.api.signedUrl(ex.storage_path), S.api.signedUrl(ex.storage_path, nice)]);
    $('#pnew', m).href = url; $('#pdl', m).href = dl; $('#pdl', m).download = nice;
    $('#pload', m).outerHTML = `<iframe class="pdf-frame" src="${esc(url)}" title="PDF"></iframe>`;
  } catch (e) { $('#pload', m).innerHTML = `<div class="form-error">${esc(e.message)}</div>`; }
}
function modal(html, onClose) {
  const wrap = document.createElement('div'); wrap.className = 'modal';
  wrap.innerHTML = `<div class="scrim" data-close></div>${html}`;
  $('#modal-root').append(wrap);
  const close = () => { wrap.remove(); onClose?.(); };
  wrap._close = close;
  $$('[data-close]', wrap).forEach(b => b.onclick = close);
  return wrap;
}

/* ================= Drawer (карточка компании) ================= */
function openDrawer(id) { S.openId = id; S.editing = false; renderDrawer(); const c = S.byId.get(id); if (c) history.replaceState(null, '', `#c=${c.code}`); $$('tr.row').forEach(r => r.classList.toggle('sel', r.dataset.open === id)); }
function closeDrawer() { S.openId = null; S.editing = false; $('#drawer')?.remove(); $('#drawer-scrim')?.remove(); history.replaceState(null, '', location.pathname + location.search); $$('tr.row.sel').forEach(r => r.classList.remove('sel')); }
function renderDrawer() {
  $('#drawer')?.remove(); $('#drawer-scrim')?.remove();
  const isNew = S.openId === 'new';
  const c = isNew ? { code: '', name: '', rma: '', sh: [], bn: [], ex: [], children: [], issues: [], section_id: S.section } : S.byId.get(S.openId);
  if (!c) return;
  const scrim = document.createElement('div'); scrim.className = 'scrim'; scrim.id = 'drawer-scrim'; scrim.onclick = closeDrawer;
  const d = document.createElement('aside'); d.className = 'drawer'; d.id = 'drawer';
  const chain = []; let p = S.byCode.get(c.parent_code); while (p && chain.length < 8) { chain.unshift(p); p = S.byCode.get(p.parent_code); }
  const sec = S.sections.find(s => s.id === c.section_id);
  d.innerHTML = `
    <div class="drawer-head"><div style="flex:1;min-width:0">
      <div class="crumbs">${chain.map(x => `<a data-open="${x.id}">${esc(x.name)}</a>`).join(' › ') || '&nbsp;'}</div>
      <h2>${esc(isNew ? t('newCompany') : c.name)}</h2>
      ${isNew ? '' : `<div><span class="pill mono">№ ${esc(c.code === '0' ? '—' : c.code)}</span><span class="pill mono">РМА ${esc(c.rma)} <button data-copy="${esc(c.rma)}" title="copy">${IC.copy.replace('<svg', '<svg width="12" height="12"')}</button></span>
        <span class="pill">${esc(c.type === 'root' ? t('typeRoot') : c.type === 'full' ? t('typeFull') : `${t('groupShare')} ${round2(c.groupDirect)}%`)}</span>${sec ? `<span class="pill">${esc(sec.title)}</span>` : ''}</div>`}
    </div><button class="icon-btn" id="dclose">${IC.x}</button></div>
    <div class="drawer-body">${S.editing ? editForm(c, isNew) : viewCard(c)}</div>
    <div class="drawer-foot">${S.editing ? `<button class="btn" id="dcancel">${esc(t('cancel'))}</button><button class="btn primary" id="dsave">${esc(t('save'))}</button>`
      : `${canEdit() ? `<button class="btn" id="dedit">${esc(t('edit'))}</button>` : ''}<button class="btn" id="dclose2">${esc(t('close'))}</button>`}</div>`;
  document.body.append(scrim, d);
  $('#dclose', d).onclick = closeDrawer; $('#dclose2', d) && ($('#dclose2', d).onclick = closeDrawer);
  $$('[data-open]', d).forEach(a => a.onclick = () => openDrawer(a.dataset.open));
  $$('[data-copy]', d).forEach(b => b.onclick = () => { navigator.clipboard?.writeText(b.dataset.copy); toast(t('copied')); });
  if (S.editing) bindEdit(d, c, isNew);
  else {
    $('#dedit', d) && ($('#dedit', d).onclick = () => { S.editing = true; renderDrawer(); });
    $$('[data-exopen]', d).forEach(b => b.onclick = () => openPdf(c, c.ex.find(x => x.id === b.dataset.exopen)));
    $$('[data-exdel]', d).forEach(b => b.onclick = async () => { if (!confirmInline(b)) return; try { await S.api.deleteExtract(c.ex.find(x => x.id === b.dataset.exdel)); toast(t('deleted')); await reload(); } catch (e) { toast(e.message, true); } });
    const up = $('#exup', d);
    if (up) up.onclick = async () => {
      const f = $('#exfile', d).files[0], date = $('#exdate', d).value;
      if (!f || !date) { toast(t('extractDate') + ' / PDF', true); return; }
      up.disabled = true;
      try { await S.api.uploadExtract(c, f, date); toast(t('uploaded')); await reload(); } catch (e) { toast(e.message, true); up.disabled = false; }
    };
    const hb = $('#hist', d);
    if (hb) hb.onclick = async () => {
      hb.disabled = true;
      try {
        const rows = await S.api.audit([c.id, ...c.sh.map(x => x.id), ...c.bn.map(x => x.id), ...c.ex.map(x => x.id)], 100);
        $('#histbox', d).innerHTML = rows.length ? `<table class="mini">${rows.map(a => `<tr><td class="mono nowrap">${esc(new Date(a.created_at).toLocaleString('ru-RU'))}</td><td>${esc(a.user_email || '—')}</td><td class="mono">${esc(a.table_name)} ${esc(a.action)}</td><td class="muted">${esc(diffKeys(a).join(', '))}</td></tr>`).join('')}</table>` : '<span class="muted">—</span>';
      } catch (e) { toast(e.message, true); }
    };
  }
}
function confirmInline(b) { if (b.dataset.armed) return true; b.dataset.armed = '1'; b.textContent = t('confirmDel'); setTimeout(() => { delete b.dataset.armed; b.textContent = t('del'); }, 3000); return false; }

function viewCard(c) {
  const kv = rows => `<dl class="kv">${rows.filter(r => r[1]).map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join('')}</dl>`;
  const bad = c.capital_declared > 0 && (c.capital_formed ?? 0) < c.capital_declared;
  const pct = c.capital_declared ? Math.min(100, (c.capital_formed || 0) / c.capital_declared * 100) : 0;
  const cur = c.extract_path;
  return `
    ${c.issues.length ? `<div class="blk"><h4>${IC.alert} ${esc(t('tabChecks'))}</h4>${c.issues.map(i => `<div class="issue ${i.ch.sev}"><b>${esc(chkTitle(i.ch))}</b><span style="margin-left:auto" class="mono">${esc(i.det)}</span></div>`).join('')}</div>` : ''}
    <div class="blk"><h4>${IC.pdf.replace('<svg', '<svg width="15" height="15"')} ${esc(t('extracts'))}<span class="spacer"></span>${cur ? `<button class="extract-link" data-exopen="${esc(c.ex.find(x => x.storage_path === cur)?.id || '')}">${IC.pdf}${esc(fmtDate(c.extract_date))}</button>` : `<span class="extract-none">${esc(fmtDate(c.extract_date) || '—')}</span>`}</h4>
      ${c.ex.length ? `<ul class="ext-list">${c.ex.map(x => `<li><span class="date">${esc(fmtDate(x.extract_date))}</span>${x.storage_path === cur ? `<span class="badge-cur">${esc(t('current'))}</span>` : ''}<span class="fn" title="${esc(x.file_name)}">${esc(x.file_name || '')}</span>
        <button class="btn" data-exopen="${x.id}">${IC.pdf}${esc(t('openPdf'))}</button>${canEdit() ? `<button class="btn danger" data-exdel="${x.id}">${esc(t('del'))}</button>` : ''}</li>`).join('')}</ul>` : `<div class="muted">${esc(t('noExtracts'))}</div>`}
      ${canEdit() ? `<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:12px;align-items:center"><input type="file" id="exfile" accept="application/pdf,.pdf" class="input" style="flex:1 1 220px"><input type="date" id="exdate" class="input" style="width:auto" value="${today()}"><button class="btn primary" id="exup">${IC.up}${esc(t('upload'))}</button></div>` : ''}
    </div>
    <div class="blk"><h4>${esc(t('capital'))}</h4><div class="cap-box"><div><div class="l">${esc(t('declared'))}</div><div class="v">${fmtNum(c.capital_declared)}</div></div><div><div class="l">${esc(t('formed'))}</div><div class="v" style="${bad ? 'color:var(--danger)' : ''}">${fmtNum(c.capital_formed)} <span style="font-size:12px">(${Math.round(pct)}%)</span></div></div><div class="bar ${bad ? 'bad' : ''}"><i style="width:${pct}%"></i></div></div></div>
    <div class="blk"><h4>${esc(t('mainInfo'))}</h4>${kv([[t('cFounded'), esc(fmtDate(c.founded_on)) + (c.reregistered_on ? ` <span class="muted">(${esc(t('rereg'))}: ${esc(fmtDate(c.reregistered_on))})</span>` : '')], [t('cAddress'), esc(c.legal_address)], [t('cActivity'), esc(c.activity)], [t('notes'), esc(c.notes)]])}</div>
    <div class="blk"><h4>${esc(t('founders'))}</h4>
      <table class="mini"><thead><tr><th>${esc(t('name'))}</th><th class="num">${esc(t('share'))}</th><th class="num">${esc(t('nominal'))}</th></tr></thead><tbody>
      ${c.sh.map(s => { const exp = c.capital_declared ? c.capital_declared * s.share_pct / 100 : null; const mism = s.nominal != null && exp != null && Math.abs(s.nominal - exp) > 1;
        return `<tr><td>${s.owner_company_id && S.byId.has(s.owner_company_id) ? `<span class="int-link" data-open="${s.owner_company_id}">${esc(s.name)}</span>` : esc(s.name)}</td><td class="num"><span class="sharebar"><i style="width:${s.share_pct}%"></i></span>${round2(s.share_pct)}%</td><td class="num" style="${mism ? 'color:var(--danger)' : ''}">${s.nominal != null ? fmtNum(s.nominal) : '<span class="muted">—</span>'}${mism ? `<div style="font-size:11px">${esc(t('expected'))} ${fmtNum(Math.round(exp))}</div>` : ''}</td></tr>`; }).join('') || '<tr><td colspan="3" class="muted">—</td></tr>'}
      </tbody></table>
      ${c.founders ? `<details style="margin-top:10px"><summary class="muted" style="cursor:pointer;font-size:12px">${esc(t('officialText'))}</summary><div style="white-space:pre-line;font-size:12.5px;margin-top:6px">${esc(c.founders)}</div></details>` : ''}
    </div>
    <div class="blk"><h4>${esc(t('beneficiaries'))}<span class="spacer"></span>${c.code !== '0' ? `<span class="share minor" title="${esc(CONFIG.ULTIMATE_OWNER)}">${esc(t('effShare'))} ${round2(c.eff)}%</span>` : ''}</h4>
      <table class="mini"><thead><tr><th>${esc(t('name'))}</th><th>${esc(t('citizenship'))}</th><th class="num">${esc(t('share'))}</th></tr></thead><tbody>
      ${c.bn.map(b => `<tr><td>${esc(b.name)}</td><td>${esc(b.citizenship || '—')}</td><td class="num"><span class="sharebar"><i style="width:${b.share_pct}%"></i></span>${round2(b.share_pct)}%</td></tr>`).join('') || '<tr><td colspan="3" class="muted">—</td></tr>'}</tbody></table>
    </div>
    <div class="blk"><h4>${esc(t('management'))}</h4>${kv([[t('position'), esc(c.position)], [t('director'), `<b>${esc(c.director)}</b>`]])}
      <div class="contact">${c.email ? `<a href="mailto:${esc(c.email)}">${IC.mail}${esc(c.email)}</a>` : ''}${c.phone ? `<a href="tel:${esc(c.phone.replace(/[^\d+]/g, ''))}">${IC.phone}${esc(c.phone)}</a>` : ''}</div>
      ${c.executive ? `<details style="margin-top:10px"><summary class="muted" style="cursor:pointer;font-size:12px">${esc(t('officialText'))}</summary><div style="white-space:pre-line;font-size:12.5px;margin-top:6px">${esc(c.executive)}</div></details>` : ''}</div>
    <div class="blk"><h4>${esc(t('board'))}</h4><div style="white-space:pre-line">${esc(c.supervisory_board || '—')}</div></div>
    ${c.children.length ? `<div class="blk"><h4>${esc(t('subsidiaries'))} · ${c.children.length}</h4><ul class="ext-list">${c.children.map(k => `<li data-open="${k.id}" style="cursor:pointer"><span class="mono muted" style="min-width:52px">${esc(k.code)}</span><span class="int-link" style="flex:1">${esc(k.name)}</span><span class="share ${k.groupDirect >= 99.99 ? '' : k.groupDirect >= 50 ? 'partial' : 'minor'}">${round2(k.groupDirect)}%</span></li>`).join('')}</ul></div>` : ''}
    ${canEdit() ? `<div class="blk"><h4>${esc(t('history'))}<span class="spacer"></span><button class="btn" id="hist">${esc(t('showHistory'))}</button></h4><div id="histbox"></div></div>` : ''}`;
}

const EDIT_FIELDS = [
  ['code', 'code', 'text'], ['rma', 'cRma', 'text'], ['name', 'cName', 'text', 'full'], ['section_id', 'section', 'section'], ['parent_code', 'parent', 'parent'],
  ['founded_on', 'cFounded', 'date'], ['reregistered_on', 'rereg', 'date'], ['legal_address', 'cAddress', 'text', 'full'], ['activity', 'cActivity', 'area', 'full'],
  ['capital_declared', 'cDeclared', 'number'], ['capital_formed', 'cFormed', 'number'], ['founders', 'cFounders', 'area', 'full'],
  ['position', 'position', 'text'], ['director', 'director', 'text'], ['email', 'email', 'text'], ['phone', 'phone', 'text'],
  ['executive', 'cExecutive', 'area', 'full'], ['supervisory_board', 'cBoard', 'area', 'full'], ['beneficiary', 'cBenef', 'area', 'full'],
  ['extract_date', 'extractDate', 'date'], ['notes', 'notes', 'area', 'full'],
];
function editForm(c, isNew) {
  const f = ([k, l, type, full]) => {
    const v = c[k] ?? '';
    let inp;
    if (type === 'area') inp = `<textarea class="input" name="${k}" rows="3">${esc(v)}</textarea>`;
    else if (type === 'section') inp = `<select class="input" name="${k}"><option value="">${esc(t('choose'))}</option>${S.sections.map(s => `<option value="${s.id}" ${s.id === c.section_id ? 'selected' : ''}>${esc(s.title)}</option>`).join('')}</select>`;
    else if (type === 'parent') inp = `<select class="input" name="${k}"><option value="">—</option>${S.companies.filter(x => x.id !== c.id).map(x => `<option value="${esc(x.code)}" ${x.code === c.parent_code ? 'selected' : ''}>${esc(x.code)} · ${esc(x.name)}</option>`).join('')}</select>`;
    else inp = `<input class="input ${type === 'number' || k === 'rma' || k === 'code' ? 'mono' : ''}" name="${k}" type="${type}" value="${esc(type === 'date' ? String(v).slice(0, 10) : v)}" ${type === 'number' ? 'step="0.01" min="0"' : ''}>`;
    return `<div class="field ${full ? 'full' : ''}"><label>${esc(t(l))}</label>${inp}</div>`;
  };
  const rowsTable = (id, rows, cols) => `<table class="mini rows-edit" id="${id}"><thead><tr>${cols.map(x => `<th>${esc(t(x[1]))}</th>`).join('')}<th></th></tr></thead><tbody>
    ${rows.map(r => rowTpl(cols, r)).join('')}</tbody></table><button class="btn" style="margin-top:6px" data-addrow="${id}">${esc(t('addRow'))}</button>`;
  return `<form id="eform"><div class="blk"><div class="form-grid">${EDIT_FIELDS.map(f).join('')}</div></div>
    <div class="blk"><h4>${esc(t('founders'))}</h4>${rowsTable('sh-edit', c.sh, SH_COLS)}</div>
    <div class="blk"><h4>${esc(t('beneficiaries'))}</h4>${rowsTable('bn-edit', c.bn, BN_COLS)}</div>
    <div class="form-error" id="eerr"></div></form>`;
}
const SH_COLS = [['name', 'name', 'text'], ['share_pct', 'share', 'number'], ['nominal', 'nominal', 'number']];
const BN_COLS = [['name', 'name', 'text'], ['citizenship', 'citizenship', 'text'], ['share_pct', 'share', 'number']];
const rowTpl = (cols, r = {}) => `<tr>${cols.map(([k, , type]) => `<td><input class="input ${type === 'number' ? 'mono' : ''}" data-k="${k}" type="${type}" ${type === 'number' ? 'step="0.0001" min="0"' : ''} value="${esc(r[k] ?? '')}"></td>`).join('')}<td><button type="button" class="icon-btn" data-delrow>${IC.x}</button></td></tr>`;
function bindEdit(d, c, isNew) {
  const form = $('#eform', d);
  const bindDel = () => $$('[data-delrow]', d).forEach(b => b.onclick = () => b.closest('tr').remove());
  bindDel();
  $$('[data-addrow]', d).forEach(b => b.onclick = e => { e.preventDefault(); const id = b.dataset.addrow; $(`#${id} tbody`, d).insertAdjacentHTML('beforeend', rowTpl(id === 'sh-edit' ? SH_COLS : BN_COLS)); bindDel(); });
  $('#dcancel', d).onclick = () => { if (isNew) closeDrawer(); else { S.editing = false; renderDrawer(); } };
  $('#dsave', d).onclick = async () => {
    const fd = Object.fromEntries(new FormData(form));
    const patch = {};
    for (const [k, , type] of EDIT_FIELDS) {
      let v = (fd[k] ?? '').toString().trim();
      patch[k] = v === '' ? null : type === 'number' ? Number(v) : type === 'section' ? Number(v) : v;
    }
    if (!patch.code || !patch.name || !/^\d{9}$/.test(patch.rma || '')) { $('#eerr', d).textContent = t('required'); return; }
    patch.code = patch.code.replace(/\.$/, '');
    patch.sort_key = codeKey(patch.code);
    if (!patch.parent_code && patch.code !== '0') patch.parent_code = patch.code.includes('.') ? patch.code.split('.').slice(0, -1).join('.') : '0';
    const readRows = (id, cols) => $$(`#${id} tbody tr`, d).map(tr => Object.fromEntries(cols.map(([k, , type]) => { const v = $(`[data-k="${k}"]`, tr).value.trim(); return [k, v === '' ? null : type === 'number' ? Number(v) : v]; }))).filter(r => r.name);
    const byName = new Map(S.companies.map(x => [normName(x.name), x.id]));
    const sh = readRows('sh-edit', SH_COLS).map(r => ({ ...r, owner_company_id: byName.get(normName(r.name)) ?? null }));
    const bn = readRows('bn-edit', BN_COLS);
    $('#dsave', d).disabled = true;
    try {
      let saved;
      if (S.api.saveCompanyFull) saved = await S.api.saveCompanyFull(isNew ? null : c.id, patch, sh, bn, S.byId);
      else { saved = await S.api.saveCompany(isNew ? null : c.id, patch); await S.api.replaceRows('reg_shareholders', saved.id, sh); await S.api.replaceRows('reg_beneficiaries', saved.id, bn); }
      S.openId = saved.id; S.editing = false; toast(t('saved')); await reload();
    } catch (e) { $('#eerr', d).textContent = e.message; $('#dsave', d).disabled = false; }
  };
}

/* ================= Bulk import PDF ================= */
let pdfjsLib = null;
async function pdfText(file) {
  if (!pdfjsLib) { pdfjsLib = await import(LIB.pdfjs); pdfjsLib.GlobalWorkerOptions.workerSrc = LIB.pdfjsWorker; }
  const doc = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
  let s = ''; for (let i = 1; i <= Math.min(doc.numPages, 2); i++) { const tc = await (await doc.getPage(i)).getTextContent(); s += tc.items.map(x => x.str).join(' ') + '\n'; }
  return s;
}
export function parseExtractFileName(name) {
  const base = name.replace(/^[0-9a-f]{8}-(?=\d)/i, '').trim();
  const code = base.match(/^(\d+(?:\.\d+)*)\.?(?=[\s_\-]|\D|$)/)?.[1] ?? null;
  const ds = [...base.matchAll(/(\d{1,2})\.(\d{1,2})\.(\d{4})/g)].pop();
  const date = ds ? `${ds[3]}-${ds[2].padStart(2, '0')}-${ds[1].padStart(2, '0')}` : null;
  return { code, date };
}
function openImport(initial = [], autoUpload = false) {
  const items = [];
  const m = modal(`<div class="modal-card"><div class="modal-head"><span style="color:var(--accent)">${IC.up}</span><h3>${esc(t('impTitle'))}</h3><button class="icon-btn" data-close>${IC.x}</button></div>
    <div class="modal-body"><label class="drop" id="drop">${t('impDrop')}<div class="muted" style="font-size:12px;margin-top:6px">${esc(t('impHint'))}</div><input type="file" id="impf" accept="application/pdf,.pdf" multiple hidden></label>
      <div id="implist" style="margin-top:14px"></div></div>
    <div class="modal-foot"><span class="muted" id="impsum" style="margin-right:auto;font-size:12.5px"></span><button class="btn" data-close>${esc(t('close'))}</button><button class="btn primary" id="impgo" disabled>${esc(t('impGo', 0))}</button></div></div>`, () => { if (items.some(i => i.st === 'done')) reload(); });
  const drop = $('#drop', m), inp = $('#impf', m);
  drop.ondragover = e => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = e => { e.preventDefault(); drop.classList.remove('over'); add([...e.dataTransfer.files]); };
  inp.onchange = () => { add([...inp.files]); inp.value = ''; };
  const byRma = new Map(S.companies.map(c => [c.rma, c]));
  if (initial.length) setTimeout(async () => { await add(initial); if (autoUpload && items.some(ready)) $('#impgo', m).click(); }, 0);
  async function add(files) {
    for (const f of files.filter(f => /\.pdf$/i.test(f.name) || f.type === 'application/pdf')) {
      const it = { file: f, st: 'reading', ...parseExtractFileName(f.name) };
      it.companyId = S.byCode.get(it.code)?.id ?? null; items.push(it); draw();
      try {
        const text = await pdfText(f);
        const rmas = [...new Set(text.match(/\b\d{9}\b/g) || [])];
        const hit = rmas.map(r => byRma.get(r)).find(Boolean);
        it.pdfRma = hit?.rma ?? null;
        if (hit && !it.companyId) it.companyId = hit.id;
        it.how = hit ? (hit.id === it.companyId ? 'rma' : 'rmadiff') : 'code';
        if (!it.date) { const ds = [...text.matchAll(/(\d{2})\.(\d{2})\.(\d{4})/g)].pop(); if (ds) it.date = `${ds[3]}-${ds[2]}-${ds[1]}`; }
      } catch { it.how = 'code'; }
      it.st = 'ready'; draw();
    }
  }
  function status(it) {
    if (it.st === 'reading') return `<span class="st wait">${esc(t('impReading'))}</span>`;
    if (it.st === 'uploading') return `<span class="st wait">${esc(t('impUploading'))}</span>`;
    if (it.st === 'done') return `<span class="st ok">✓ ${esc(t('impDone'))}</span>`;
    if (it.st === 'error') return `<span class="st err" title="${esc(it.err)}">${esc(t('impErr'))}: ${esc(it.err).slice(0, 60)}</span>`;
    if (!it.companyId) return `<span class="st err">${esc(t('impNoMatch'))}</span>`;
    const c = S.byId.get(it.companyId);
    if (c.ex.some(x => x.extract_date === it.date)) return `<span class="st warn">${esc(t('impExists'))}</span>`;
    if (it.how === 'rmadiff') return `<span class="st err">${esc(t('impRmaDiff'))}</span>`;
    return it.how === 'rma' ? `<span class="st ok">✓ ${esc(t('impOkRma'))}</span>` : `<span class="st warn">${esc(t('impByCode'))}</span>`;
  }
  const ready = it => it.st === 'ready' && it.companyId && it.date && it.how !== 'rmadiff' && !S.byId.get(it.companyId).ex.some(x => x.extract_date === it.date);
  function draw() {
    $('#implist', m).innerHTML = items.length ? `<table class="mini imp"><thead><tr><th>${esc(t('impFile'))}</th><th>${esc(t('impCompany'))}</th><th>${esc(t('impDate'))}</th><th>${esc(t('impStatus'))}</th></tr></thead><tbody>
      ${items.map((it, i) => `<tr><td style="max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${esc(it.file.name)}">${esc(it.file.name)}</td>
        <td><select class="input" data-ic="${i}" ${it.st === 'done' ? 'disabled' : ''}><option value="">${esc(t('choose'))}</option>${S.companies.map(c => `<option value="${c.id}" ${c.id === it.companyId ? 'selected' : ''}>${esc(c.code)} · ${esc(c.name)}</option>`).join('')}</select></td>
        <td><input type="date" class="input" data-id="${i}" value="${esc(it.date || '')}" ${it.st === 'done' ? 'disabled' : ''}></td><td>${status(it)}</td></tr>`).join('')}</tbody></table>` : '';
    $$('[data-ic]', m).forEach(s => s.onchange = () => { const it = items[s.dataset.ic]; it.companyId = s.value || null; it.how = it.pdfRma ? (S.byId.get(it.companyId)?.rma === it.pdfRma ? 'rma' : 'rmadiff') : 'code'; draw(); });
    $$('[data-id]', m).forEach(s => s.onchange = () => { items[s.dataset.id].date = s.value || null; draw(); });
    const n = items.filter(ready).length;
    $('#impgo', m).disabled = !n; $('#impgo', m).textContent = t('impGo', n);
    $('#impsum', m).textContent = items.length ? `${items.filter(i => i.st === 'done').length} / ${items.length}` : '';
  }
  $('#impgo', m).onclick = async () => {
    $('#impgo', m).disabled = true;
    for (const it of items.filter(ready)) {
      it.st = 'uploading'; draw();
      try { const ex = await S.api.uploadExtract(S.byId.get(it.companyId), it.file, it.date); S.byId.get(it.companyId).ex.push(ex); it.st = 'done'; }
      catch (e) { it.st = 'error'; it.err = e.message; }
      draw();
    }
    await reload(); draw();
  };
}

/* ================= Excel export ================= */
async function exportXlsx() {
  if (!window.XLSX) await loadScript(LIB.xlsx);
  const rows = filtered();
  const H = ['№', 'Номи ҷамъият', 'РМА', 'Санаи таъсис ё бадастовардани ҷамъият', 'Суроғаи ҳуқуқӣ', 'Намуди фаъолият', 'Сармояи оинномавии эълоншуда', 'Сармояи оинномавии ташаккулёфта', 'Муассисон', 'Мақоми иҷроия', 'Шӯрои нозирон', 'Маълумот дар бораи молик-бенефитсиар', 'Иқтибос бо истинод', 'Файли иқтибос'];
  const aoa = [H]; const merges = []; let last; const shown = new Set();
  for (const c of rows) {
    if (c.section_id !== last && c.code !== '0' && !shown.has(c.section_id)) { last = c.section_id; shown.add(c.section_id); const s = S.sections.find(x => x.id === c.section_id); if (s) { merges.push({ s: { r: aoa.length, c: 0 }, e: { r: aoa.length, c: H.length - 1 } }); aoa.push([s.title]); } }
    aoa.push([c.code === '0' ? '' : c.code + '.', c.name, c.rma, fmtDate(c.founded_on) + (c.reregistered_on ? ` (${fmtDate(c.reregistered_on)})` : ''), c.legal_address, c.activity,
      Number(c.capital_declared ?? 0), Number(c.capital_formed ?? 0), c.founders, c.executive, c.supervisory_board || '-', c.beneficiary || '-', fmtDate(c.extract_date) || '-', c.extract_file_name || '']);
  }
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  ws['!merges'] = merges; ws['!cols'] = [6, 30, 12, 14, 30, 30, 16, 16, 40, 40, 30, 40, 12, 24].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Феҳрист');
  XLSX.writeFile(wb, `Fehrist_jamiyatho_${today()}.xlsx`);
}

/* ================= Keyboard ================= */
document.addEventListener('keydown', e => {
  if (e.key === '/' && !/INPUT|TEXTAREA|SELECT/.test(document.activeElement.tagName)) { const q = $('#q'); if (q && !q.closest('.hidden')) { e.preventDefault(); q.focus(); } }
  if (e.key === 'Escape') {
    const mods = $$('#modal-root .modal'); if (mods.length) { mods.pop()._close(); return; }
    if (S.openId && !S.editing) { closeDrawer(); return; }
    const q = $('#q'); if (q && document.activeElement === q && q.value) { q.value = ''; q.oninput(); }
  }
});

boot();
