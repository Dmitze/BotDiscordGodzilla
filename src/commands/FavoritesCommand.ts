import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import { t } from '@/i18n';
import logger from '@/utils/logger';
import type { GoogleService } from '@/services/GoogleService';
import { defaultWorkspaceService, WorkspaceService } from '@/services/WorkspaceService';
import { replyWithPrivacy } from '@/ui/reply';

function getWorkspace(interaction: any): WorkspaceService {
  const svc = (interaction.client as any)?.serviceContainer?.get?.('workspace');
  return (svc as WorkspaceService) || defaultWorkspaceService;
}

function getGoogle(interaction: any): GoogleService | undefined {
  return ((interaction.client as any)?.serviceContainer?.get?.('google') as GoogleService) || undefined;
}

export class FavoritesCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('fav', t('workspace.fav.command.description'), config, {}, (builder: any) => {
      return builder
        .addSubcommand((sub: any) =>
          sub
            .setName('add')
            .setDescription(t('workspace.fav.sub.add'))
            .addStringOption((opt: any) =>
              opt.setName('fileid').setDescription('Google Drive file ID').setRequired(true)
            )
        )
        .addSubcommand((sub: any) =>
          sub.setName('list').setDescription(t('workspace.fav.sub.list'))
        )
        .addSubcommand((sub: any) =>
          sub
            .setName('remove')
            .setDescription(t('workspace.fav.sub.remove'))
            .addStringOption((opt: any) =>
              opt.setName('fileid').setDescription('Google Drive file ID').setRequired(true)
            )
        );
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options as any;
    try {
      const sub = interaction.options.getSubcommand();
      if (sub === 'add') return this.handleAdd(interaction);
      if (sub === 'list') return this.handleList(interaction);
      if (sub === 'remove') return this.handleRemove(interaction);
      await replyWithPrivacy(interaction, t('workspace.common.unknownSub'));
    } catch (error) {
      logger.error('FavoritesCommand failed', { error: String(error) });
      if (interaction.deferred || interaction.replied) {
        await interaction.editReply({ content: t('workspace.common.execError') });
      } else {
        await replyWithPrivacy(interaction, t('workspace.common.execError'));
      }
    }
  }

  private async handleAdd(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const fileId = interaction.options.getString('fileid', true);
    const ws = getWorkspace(interaction);

    const { added } = await ws.addFavorite(userId, fileId);
    await replyWithPrivacy(interaction, added ? t('workspace.fav.added') : t('workspace.fav.exists'));
  }

  private async handleList(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const ws = getWorkspace(interaction);
    const list = await ws.listFavorites(userId);
    if (!list.length) {
      await replyWithPrivacy(interaction, t('workspace.fav.empty'));
      return;
    }

    const google = getGoogle(interaction);
    const lines: string[] = [];
    for (const f of list.slice(0, 25)) {
      if (google) {
        try {
          const meta = await google.getDriveFile(f.fileId);
          lines.push(`• ${meta.name} (${meta.id})`);
        } catch {
          lines.push(`• ${f.fileId}`);
        }
      } else {
        lines.push(`• ${f.fileId}`);
      }
    }

    await replyWithPrivacy(interaction, `${t('workspace.fav.listTitle')}\n${lines.join('\n')}`);
  }

  private async handleRemove(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const fileId = interaction.options.getString('fileid', true);
    const ws = getWorkspace(interaction);
    const ok = await ws.removeFavorite(userId, fileId);
    await replyWithPrivacy(interaction, ok ? t('workspace.fav.removed') : t('workspace.common.notFound'));
  }
}
