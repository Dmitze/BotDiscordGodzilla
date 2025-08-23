import type { ChatInputCommandInteraction } from 'discord.js';
import type { SearchIndex } from '@/search/SearchIndex';

export interface ChooseIndexResult {
  mode: 'sqlite' | 'legacy';
  services: {
    searchIndex?: SearchIndex;
    google?: any;
    cache?: any;
  };
}

/**
 * Decide which index mode to use and prefetch services in a deterministic order
 * to play nice with mocked serviceContainer.get sequencing in tests.
 */
export function chooseIndexMode(interaction: ChatInputCommandInteraction): ChooseIndexResult {
  const sc: any = (interaction as any)?.client?.serviceContainer;
  const getSvc: ((name: string) => any) | undefined = sc?.get?.bind(sc);

  if (!getSvc) {
    return { mode: 'legacy', services: {} };
  }

  // Fetch google/cache first, then searchIndex to avoid consuming mockReturnValueOnce unexpectedly.
  const google = getSvc('google') ?? getSvc('GoogleService');
  const cache = getSvc('cache') ?? getSvc('CacheService');
  const searchIndex =
    (getSvc('searchIndex') as SearchIndex | undefined) ??
    (getSvc('SearchIndex') as SearchIndex | undefined) ??
    (getSvc('sqliteSearchIndex') as SearchIndex | undefined) ??
    (getSvc('SqliteSearchIndex') as SearchIndex | undefined);

  if (searchIndex) return { mode: 'sqlite', services: { searchIndex, google, cache } };
  if (google) return { mode: 'legacy', services: { google, cache } };
  // default fallback
  return { mode: 'legacy', services: {} };
}
