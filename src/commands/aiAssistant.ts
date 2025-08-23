import type { ChatInputCommandInteraction, InteractionEditReplyOptions } from 'discord.js';

type MinimalEmbed = { data: { title: string; description?: string } };

interface BotLike {
  getService?: (name: string) => unknown;
  handleError?: (err: unknown) => Promise<void> | void;
}

export async function execute(interaction: ChatInputCommandInteraction, bot?: BotLike): Promise<void> {
  const ai = bot?.getService?.('ai') as any;
  const rag = bot?.getService?.('rag') as any;
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
      // Try RAG first if available
      if (rag?.answer && typeof rag.answer === 'function') {
        const res = await rag.answer(query, {
          k: Number(process.env['RETRIEVER_K'] ?? 6),
          alpha: Number(process.env['RETRIEVER_ALPHA'] ?? 0.5),
        }, {
          maskPII: true,
          maxTokens: Number(process.env['RAG_MAX_CONTEXT_TOKENS'] ?? 1200),
        }, {
          maxTokens: Number(process.env['AI_MAX_TOKENS'] ?? 512),
        });
        const sources = res.chunks?.map((c: any, i: number) => `[${i + 1}] ${c.name}`).join(', ');
        answer = `${res.answer}\n\nДжерела: ${sources || '—'}`;
      } else {
        if (!ai?.generateResponse) {
          throw new Error('AI service unavailable');
        }
        const plain = await ai.generateResponse(String(query || 'Запит від користувача'), { maxTokens: 512 });
        answer = typeof plain === 'string' ? plain : String(plain?.content ?? '');
      }
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
