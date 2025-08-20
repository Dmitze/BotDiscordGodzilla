import type {
  ChatInputCommandInteraction,
  MessageComponentInteraction,
  InteractionReplyOptions,
  BaseMessageOptions,
} from 'discord.js';

/**
 * Ephemeral-first reply helper with optional public share follow-up.
 *
 * Usage:
 *  await replyWithPrivacy(interaction, { content: 'Done' }, { share: true, public: { content: '✅ Done' } });
 */
export type ShareReplyOptions = {
  // If true, post a second, public message with sanitized content
  share?: boolean;
  // Public message content/options (sanitized, PII-free)
  public?: string | BaseMessageOptions | InteractionReplyOptions;
  // Override default ephemeral flag for the first reply
  ephemeral?: boolean;
};

export async function replyWithPrivacy(
  interaction: ChatInputCommandInteraction | MessageComponentInteraction,
  content: string | InteractionReplyOptions,
  opts: ShareReplyOptions = {}
): Promise<void> {
  const ephemeral = opts.ephemeral ?? true; // ephemeral by default
  const firstReply: InteractionReplyOptions = typeof content === 'string' ? { content, ephemeral } : { ...content, ephemeral };

  if (!interaction.deferred && !interaction.replied) {
    await interaction.reply(firstReply);
  } else {
    await interaction.followUp(firstReply);
  }

  if (opts.share) {
    const publicPayload: InteractionReplyOptions = typeof opts.public === 'string' ? { content: opts.public } : { ...(opts.public ?? {}) };
    // Enforce public visibility for shared message
    delete (publicPayload as any).ephemeral;

    await interaction.followUp(publicPayload);
  }
}
