// Настройки подключения к Supabase.
// Project Settings → API → Project URL и anon public key.
// anon-ключ публичный по замыслу Supabase: данные защищены RLS и входом по логину.
export const CONFIG = {
  // Google Sheets: URL веб-приложения Apps Script (…/exec). Если задан — данные берутся из Google Таблицы.
  SHEETS_URL: '',

  // Облако без сервера: зашифрованные данные и PDF в этом же репозитории GitHub (файлы store.*).
  // Чтение — по паролю; правка — по паролю + GitHub-токену с правом записи.
  STORE: { repo: 'isosaid/avesto-reestr', branch: 'main' },

  SUPABASE_URL: 'https://YOUR-PROJECT-REF.supabase.co',
  SUPABASE_ANON_KEY: 'YOUR-ANON-PUBLIC-KEY',

  // Конечный бенефициар холдинга — для расчёта эффективной доли по цепочке владения
  ULTIMATE_OWNER: 'Мирзо Муҳамад Файзуллозода',

  // Выписка старше этого числа дней помечается как устаревшая
  EXTRACT_MAX_AGE_DAYS: 365,
};
