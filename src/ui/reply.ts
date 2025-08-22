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
  // If provided, defines what the default ephemeral behavior should be when 'ephemeral' is not set
  ephemeralByDefault?: boolean;
  // If true, allows inferring share/public from the payload itself (e.g. payload.share === true)
  shareFlagSupport?: boolean;
};

export async function replyWithPrivacy(
  interaction: ChatInputCommandInteraction | MessageComponentInteraction,
  content: string | InteractionReplyOptions,
  opts: ShareReplyOptions = {}
): Promise<void> {
  // Determine default ephemeral behavior
  const ephemeralDefault = opts.ephemeralByDefault ?? true;
  const ephemeral = opts.ephemeral ?? ephemeralDefault;

  // Infer share/public from payload when enabled and not explicitly set
  let share = opts.share;
  let publicFromPayload: string | BaseMessageOptions | InteractionReplyOptions | undefined = opts.public;
  if (share === undefined && opts.shareFlagSupport && typeof content !== 'string') {
    const anyContent = content as any;
    if (anyContent && (anyContent.share === true || anyContent?.flags?.share === true)) {
      share = true;
      publicFromPayload = publicFromPayload ?? anyContent.public;
    }
  }

  const firstReply: InteractionReplyOptions =
    typeof content === 'string' ? { content, ephemeral } : { ...content, ephemeral };

  if (!interaction.deferred && !interaction.replied) {
    await interaction.reply(firstReply);
  } else {
    await interaction.followUp(firstReply);
  }

  if (share) {
    const base = publicFromPayload ?? opts.public;
    const publicPayload: InteractionReplyOptions =
      typeof base === 'string' ? { content: base } : { ...(base ?? {}) };
    // Enforce public visibility for shared message (strip ephemeral if present)
    const { ephemeral: _ignored, ...rest } = publicPayload as InteractionReplyOptions & {
      ephemeral?: boolean;
    };

    await interaction.followUp(rest);
  }
}
