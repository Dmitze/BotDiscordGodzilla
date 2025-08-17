import type { ChatInputCommandInteraction, InteractionEditReplyOptions } from 'discord.js';

type MinimalEmbed = { data: { title: string; description?: string } };

interface BotLike {
  getService?: (name: string) => unknown;
  handleError?: (err: unknown) => Promise<void> | void;
}

export async function execute(interaction: ChatInputCommandInteraction, bot?: BotLike): Promise<void> {
  const ai = bot?.getService?.('ai') as any;
  const cache = bot?.getService?.('cache') as any;
  const metrics = bot?.getService?.('metrics') as any;

  const started = performance.now();
  try {
    const query = interaction.options.getString?.('запит', false)
      ?? interaction.options.getString?.('query', false)
      ?? '';
    const context = interaction.options.getString?.('контекст', false)
      ?? interaction.options.getString?.('context', false)
      ?? null;
    const mode = interaction.options.getString?.('режим', false)
      ?? interaction.options.getString?.('mode', false)
      ?? 'general';

    await interaction.deferReply?.();

    const cacheKey = `ai:${mode}:${query}:${context ?? ''}`;
    let answer: string | null = null;
    if (cache?.get) {
      answer = await cache.get(cacheKey);
    }

    if (!answer) {
      if (!ai?.generateResponse) {
        throw new Error('AI service unavailable');
      }
      answer = await ai.generateResponse({ query, context, mode });
      await cache?.set?.(cacheKey, answer);
    }

    const embed: MinimalEmbed = { data: { title: '🤖 AI Відповідь', description: String(answer) } };
    await interaction.editReply?.({ embeds: [embed] } as unknown as InteractionEditReplyOptions);
    metrics?.incrementCommand?.('ai', 'success');
    metrics?.measureCommandDuration?.('ai', performance.now() - started);
  } catch (err) {
    metrics?.incrementCommand?.('ai', 'error');
    metrics?.measureCommandDuration?.('ai', performance.now() - started);
    await bot?.handleError?.(err);
  }
}
