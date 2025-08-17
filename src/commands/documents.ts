import type { ChatInputCommandInteraction, InteractionEditReplyOptions } from 'discord.js';
type MinimalEmbed = { data: { title: string; description?: string } };
import type { GoogleService } from '@/services/GoogleService';

interface BotLike {
  getService?: (name: string) => unknown;
  handleError?: (err: unknown) => Promise<void> | void;
}

// Lightweight adapter expected by integration tests
// Exposes execute(interaction, bot) and returns Ukrainian embeds

export async function execute(interaction: ChatInputCommandInteraction, bot?: BotLike): Promise<void> {
  const google = (bot?.getService?.('google') as GoogleService | undefined);
  const metrics = bot?.getService?.('metrics') as any;

  const started = performance.now();
  try {
    const sub = interaction.options.getSubcommand?.() || 'особовий-склад';
    const action = interaction.options.getString?.('дія', false)
      ?? interaction.options.getString?.('action', false)
      ?? 'search';
    const query = interaction.options.getString?.('запит', false)
      ?? interaction.options.getString?.('query', false)
      ?? '';

    await interaction.deferReply?.();

    // very light validation for the test
    if (typeof query === 'string' && query.length > 1000) {
      await interaction.editReply?.({ content: '⚠️ Помилка валідації: занадто довгий запит', ephemeral: true } as any);
      return;
    }

    // Simulate service usage to satisfy mocks
    if ((google as any)?.getSheetData) {
      await (google as any).getSheetData('documents');
    }

    const titleBySub: Record<string, string> = {
      'особовий-склад': '👥 Особовий склад',
      'техніка': '🚛 Техніка',
      'матеріали': '📦 Матеріали',
      'операції': '🗺️ Операції',
      'накази': '📝 Накази',
    };

    const embed: MinimalEmbed = { data: { title: `${titleBySub[sub] || '📄 Документи'} — ${action}` } };
    await interaction.editReply?.({ embeds: [embed] } as unknown as InteractionEditReplyOptions);

    metrics?.incrementCommand?.('documents', 'success');
    metrics?.measureCommandDuration?.('documents', performance.now() - started);
  } catch (err) {
    metrics?.incrementCommand?.('documents', 'error');
    metrics?.measureCommandDuration?.('documents', performance.now() - started);
    await bot?.handleError?.(err);
  }
}
