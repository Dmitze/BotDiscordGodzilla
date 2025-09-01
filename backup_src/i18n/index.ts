/* Simple i18n utility with default locale = 'uk'; 'uk-UA' maps to 'uk' dictionary.
 * Additionally exposes Intl helpers that use a region-aware locale.
 */
import uk from './uk.json';
import en from './en.json';
// Types are imported as type-only to avoid runtime dependency on discord.js
import type { Interaction, User, Guild, GuildMember } from 'discord.js';

// Narrow dictionary types to avoid any/unsafe
type DictLeaf = string | string[];
export interface Dictionary {
  [key: string]: DictLeaf | Dictionary;
}

export type LocaleKey = 'uk' | 'en';
export type LocaleInput = 'uk' | 'uk-UA' | 'en' | 'en-US';

const locales: Record<LocaleKey, Dictionary> = {
  uk: uk as unknown as Dictionary,
  en: en as unknown as Dictionary,
};

// The dictionary locale (for resource lookups)
let dictLocale: LocaleKey = 'uk';
// The UI/Intl locale (can include region subtags, e.g., 'uk-UA')
let uiLocale: LocaleInput = 'uk';

function resolveDictLocale(input: LocaleInput): LocaleKey {
  // Map region to base: 'uk-UA' -> 'uk', 'en-US' -> 'en'
  if (input === 'uk' || input === 'uk-UA') return 'uk';
  if (input === 'en' || input === 'en-US') return 'en';
  return 'uk';
}

export function setLocale(locale: LocaleInput): void {
  uiLocale = locale;
  dictLocale = resolveDictLocale(locale);
}

export function getLocale(): LocaleInput {
  return uiLocale;
}

export function getDictLocale(): LocaleKey {
  return dictLocale;
}

// For Intl, prefer a region-aware locale. For Ukrainian we standardize to 'uk-UA'.
export function getIntlLocale(): string {
  if (uiLocale === 'uk' || uiLocale === 'uk-UA') return 'uk-UA';
  if (uiLocale === 'en' || uiLocale === 'en-US') return 'en-US';
  return String(uiLocale);
}

// --- Discord-aware locale resolution helpers ---
function mapToLocaleInput(input?: string | null): LocaleInput | undefined {
  if (!input) return undefined;
  const lower = input.toLowerCase();
  if (lower.startsWith('uk')) return 'uk';
  if (lower.startsWith('en')) return 'en';
  return undefined;
}

// Accepts Interaction | User | Guild | GuildMember | arbitrary object with locale fields
function resolveDiscordLocale(source?: unknown): LocaleInput {
  // 1) Try interaction.locale / interaction.guildLocale (discord.js v14)
  const asAny = source as any;
  const fromInteraction: string | undefined = asAny?.locale || asAny?.guildLocale;
  const fromInteractionUser: string | undefined = asAny?.user?.locale;
  const fromMember: string | undefined = asAny?.member?.user?.locale;
  const fromGuildPref: string | undefined = asAny?.guild?.preferredLocale || asAny?.preferredLocale;

  const loc = mapToLocaleInput(fromInteraction)
    || mapToLocaleInput(fromInteractionUser)
    || mapToLocaleInput(fromMember)
    || mapToLocaleInput(fromGuildPref);

  // Fallback strictly to Ukrainian per product preference
  return loc ?? 'uk';
}

function getDeep(obj: Dictionary, path: string): DictLeaf | Dictionary | undefined {
  return path
    .split('.')
    .reduce<DictLeaf | Dictionary | undefined>((acc, key) => {
      if (!acc || typeof acc !== 'object') return undefined;
      const next = (acc as Dictionary)[key];
      return next;
    }, obj);
}

function interpolate(template: string, vars?: Record<string, string | number>): string {
  if (!template || !vars) return template;
  return template.replace(/\{\{(.*?)\}\}/g, (_: string, k: string) => {
    const key = k.trim();
    const v = vars[key];
    return String(v ?? '');
  });
}

// Core translate with explicit locale support
function translate(key: string, vars?: Record<string, string | number>, locale?: LocaleInput): string {
  const locKey: LocaleKey = locale ? resolveDictLocale(locale) : dictLocale;
  const dict = locales[locKey];
  let raw = getDeep(dict, key);

  // Fallback to Ukrainian if missing in selected locale
  if (raw === undefined) {
    raw = getDeep(locales.uk, key);
  }

  if (typeof raw === 'string') return interpolate(raw, vars);
  if (Array.isArray(raw)) {
    const pick = raw[Math.floor(Math.random() * raw.length)] || '';
    return interpolate(pick, vars);
  }
  return key; // fallback to key when missing
}

// Overloads for t()
export function t(key: string, vars?: Record<string, string | number>): string;
export function t(key: string, vars: Record<string, string | number> | undefined, locale: LocaleInput): string;
export function t(key: string, vars?: Record<string, string | number>, locale?: LocaleInput): string {
  return translate(key, vars, locale);
}

// Translate using locale derived from a Discord Interaction/User/Guild
export function tUser(
  key: string,
  source: Interaction | User | Guild | GuildMember | unknown,
  vars?: Record<string, string | number>
): string {
  const loc = resolveDiscordLocale(source);
  return translate(key, vars, loc);
}

// Intl helpers
export function formatDate(value: Date | number, options?: Intl.DateTimeFormatOptions): string {
  const locale = getIntlLocale();
  return new Intl.DateTimeFormat(locale, options).format(value);
}

export function formatNumber(value: number, options?: Intl.NumberFormatOptions): string {
  const locale = getIntlLocale();
  return new Intl.NumberFormat(locale, options).format(value);
}

export function formatCurrency(value: number, currency: string = 'UAH', options?: Intl.NumberFormatOptions): string {
  const locale = getIntlLocale();
  return new Intl.NumberFormat(locale, { style: 'currency', currency, ...options }).format(value);
}
