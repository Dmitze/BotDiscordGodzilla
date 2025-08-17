/**
 * UserPreferencesService
 * - Stores per-user locale preferences with default 'uk'
 * - Resolves locale from Discord interaction on first use
 * - Applies i18n locale via src/i18n
 */

import type { ChatInputCommandInteraction, LocaleString } from 'discord.js';
import { setLocale, type LocaleInput } from '@/i18n';

export type SupportedLocale = 'uk' | 'uk-UA' | 'en' | 'en-US';

export interface UserPreferences {
  userId: string;
  locale: SupportedLocale; // we keep UI/Intl-aware locale; i18n maps to dict locale internally
  // extendable extras
  extras?: Record<string, unknown>;
}

// In-memory store for now; can be swapped for Redis/File later
const prefs = new Map<string, UserPreferences>();

function normalizeLocale(input?: string | null): SupportedLocale {
  // default to 'uk'
  if (!input) return 'uk';
  const lc = input.toLowerCase();
  if (lc === 'uk' || lc === 'uk-ua') return 'uk-UA';
  if (lc === 'en' || lc === 'en-us') return 'en-US';
  return 'uk';
}

export const UserPreferencesService = {
  get(userId: string): UserPreferences {
    const existing = prefs.get(userId);
    if (existing) return existing;
    const def: UserPreferences = { userId, locale: 'uk' };
    prefs.set(userId, def);
    return def;
  },

  setLocale(userId: string, locale: SupportedLocale): UserPreferences {
    const current = this.get(userId);
    const next: UserPreferences = { ...current, locale };
    prefs.set(userId, next);
    // Apply to i18n immediately
    setLocale(locale as LocaleInput);
    return next;
  },

  getLocale(userId: string): SupportedLocale {
    return this.get(userId).locale;
  },

  /**
   * Resolve from interaction.locale/guildLocale on first run and apply to i18n.
   * Always ensures i18n has a locale set for this interaction.
   */
  async resolveAndApplyLocale(interaction: ChatInputCommandInteraction): Promise<SupportedLocale> {
    const userId = interaction.user?.id ?? 'unknown';
    const existing = prefs.get(userId);

    if (existing) {
      setLocale(existing.locale as LocaleInput);
      return existing.locale;
    }

    // Prefer user locale, then guild locale
    const candidate: LocaleString | string | undefined =
      (interaction.locale as LocaleString | undefined) ??
      (interaction.guildLocale as LocaleString | undefined);

    const normalized = normalizeLocale(candidate ?? null);
    const set = this.setLocale(userId, normalized);
    return set.locale;
  },

  /** For tests: reset in-memory store */
  __reset(): void {
    prefs.clear();
    // keep last i18n locale as is; tests can call setLocale separately if needed
  },
};
