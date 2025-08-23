import type { ChatInputCommandInteraction } from 'discord.js';
import type { SearchParams } from '@/types';

export type ExtractFn = (interaction: ChatInputCommandInteraction) => Promise<SearchParams>;

/**
 * Lightweight wrapper to parse and validate options using a provided extractor.
 * Keeps SearchCommand internal validation encapsulated while reducing onExecute complexity.
 */
export async function parseOptions(
  interaction: ChatInputCommandInteraction,
  extract: ExtractFn,
): Promise<SearchParams> {
  return extract(interaction);
}
