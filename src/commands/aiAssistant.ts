import type { ChatInputCommandInteraction } from 'discord.js';
import { AIAssistantCommand as AIAssistantCommandClass } from './AIAssistantCommand';
import type { GoogleService } from '@/services/GoogleService';

const getInstance = (bot?: any) => {
  const google: GoogleService | undefined = bot?.getService?.('google');
  const config: any = { locale: 'uk', features: {} };
  return new AIAssistantCommandClass(config, google);
};

export async function execute(interaction: ChatInputCommandInteraction, bot?: any): Promise<void> {
  const cmd = getInstance(bot);
  await cmd.execute({ interaction });
}

export = { execute };
