import type { ChatInputCommandInteraction } from 'discord.js';
import { SearchCommand as SearchCommandClass } from './SearchCommand';
import type { GoogleService } from '@/services/GoogleService';

// Lightweight adapter expected by integration tests
// It exposes execute(interaction, bot) and internally uses our class-based command

const getInstance = (bot?: any) => {
  const google: GoogleService | undefined = bot?.getService?.('google');
  // Use minimal config to satisfy constructor; command reads options at runtime
  const config: any = { locale: 'uk', features: {} };
  return new SearchCommandClass(config, google);
};

export async function execute(interaction: ChatInputCommandInteraction, bot?: any): Promise<void> {
  const cmd = getInstance(bot);
  await cmd.execute({ interaction });
}
