// Настройки подключения к Supabase.
// Project Settings → API → Project URL и anon public key.
// anon-ключ публичный по замыслу Supabase: данные защищены RLS и входом по логину.
export const CONFIG = {
  SUPABASE_URL: 'https://YOUR-PROJECT-REF.supabase.co',
  SUPABASE_ANON_KEY: 'YOUR-ANON-PUBLIC-KEY',

  // Конечный бенефициар холдинга — для расчёта эффективной доли по цепочке владения
  ULTIMATE_OWNER: 'Мирзо Муҳамад Файзуллозода',

  // Выписка старше этого числа дней помечается как устаревшая
  EXTRACT_MAX_AGE_DAYS: 365,
};
