import { createI18n } from 'vue-i18n'
import en from './en'
import th from './th'

export type Locale = 'th' | 'en'

export const LOCALES: { code: Locale; short: string; label: string }[] = [
  { code: 'th', short: 'TH', label: 'ไทย' },
  { code: 'en', short: 'EN', label: 'English' },
]

const STORAGE_KEY = 'excel-editor.locale'

/** Saved choice first, then the browser's language. */
function initialLocale(): Locale {
  const saved = localStorage.getItem(STORAGE_KEY)
  if (saved === 'th' || saved === 'en') return saved
  return navigator.languages.some((l) => l.toLowerCase().startsWith('th')) ? 'th' : 'en'
}

export const i18n = createI18n<[typeof en], Locale, false>({
  legacy: false,
  locale: initialLocale(),
  fallbackLocale: 'en',
  messages: { en, th },
})

function syncDocument(locale: Locale) {
  document.documentElement.lang = locale
  document.title = i18n.global.t('app.title')
}

export function setLocale(locale: Locale) {
  i18n.global.locale.value = locale
  localStorage.setItem(STORAGE_KEY, locale)
  syncDocument(locale)
}

syncDocument(i18n.global.locale.value)
