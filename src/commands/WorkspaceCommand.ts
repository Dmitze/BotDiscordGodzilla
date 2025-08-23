/**
 * WorkspaceCommand — персональні закладки користувача
 * /ws add|list|remove|run
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import { t } from '@/i18n';
import logger from '@/utils/logger';
import { UserWorkspaceService, type WorkspaceItemType } from '@/services/UserWorkspaceService';
import { replyWithPrivacy } from '@/ui/reply';

export class WorkspaceCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super('ws', t('ws.command.description'), config, { i18n: { nameKey: 'commands.ws.name', descriptionKey: 'ws.command.description' } }, (builder: any) => {
      return builder
        .addSubcommand((sub: any) =>
          sub
            .setName('add')
            .setDescription(t('ws.sub.add.description'))
            .addStringOption((opt: any) =>
              opt
                .setName('type')
                .setDescription(t('ws.opt.type.description'))
                .setRequired(true)
                .addChoices({ name: 'file', value: 'file' }, { name: 'query', value: 'query' })
            )
            .addStringOption((opt: any) =>
              opt.setName('title').setDescription(t('ws.opt.title.description')).setRequired(false)
            )
            .addStringOption((opt: any) =>
              opt.setName('fileid').setDescription(t('ws.opt.fileid.description')).setRequired(false)
            )
            .addStringOption((opt: any) =>
              opt.setName('query').setDescription(t('ws.opt.query.description')).setRequired(false)
            )
        )
        .addSubcommand((sub: any) =>
          sub
            .setName('list')
            .setDescription(t('ws.sub.list.description'))
            .addStringOption((opt: any) =>
              opt
                .setName('type')
                .setDescription(t('ws.opt.type.description'))
                .setRequired(false)
                .addChoices({ name: 'file', value: 'file' }, { name: 'query', value: 'query' })
            )
        )
        .addSubcommand((sub: any) =>
          sub
            .setName('remove')
            .setDescription(t('ws.sub.remove.description'))
            .addStringOption((opt: any) =>
              opt.setName('id').setDescription('ID').setRequired(true)
            )
        )
        .addSubcommand((sub: any) =>
          sub
            .setName('run')
            .setDescription(t('ws.sub.run.description'))
            .addStringOption((opt: any) =>
              opt.setName('id').setDescription('ID').setRequired(true)
            )
        );
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      // Feature flag: workspace can be disabled via config
      if (!this.config.features?.enableUserWorkspace) {
        await replyWithPrivacy(interaction, { content: t('ws.error.disabled') });
        return;
      }
      const sub = interaction.options.getSubcommand();
      switch (sub) {
        case 'add':
          await this.handleAdd(interaction);
          break;
        case 'list':
          await this.handleList(interaction);
          break;
        case 'remove':
          await this.handleRemove(interaction);
          break;
        case 'run':
          await this.handleRun(interaction);
          break;
        default:
          await replyWithPrivacy(interaction, { content: t('ws.reply.unknownSub') });
      }
    } catch (error) {
      logger.error('Помилка WorkspaceCommand', {
        type: 'command',
        component: 'WorkspaceCommand.onExecute',
        error: error instanceof Error ? error.message : String(error),
      });
      if (interaction.deferred) {
        await interaction.editReply({ content: t('ws.error.exec') });
      } else {
        await replyWithPrivacy(interaction, { content: t('ws.error.exec') });
      }
    }
  }

  private async handleAdd(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const type = interaction.options.getString('type', true) as WorkspaceItemType;
    const title = interaction.options.getString('title') || '';
    const fileId = interaction.options.getString('fileid');
    const query = interaction.options.getString('query');

    if (type === 'file' && !fileId) {
      await replyWithPrivacy(interaction, { content: t('ws.error.missingFileId') });
      return;
    }
    if (type === 'query' && !query) {
      await replyWithPrivacy(interaction, { content: t('ws.error.missingQuery') });
      return;
    }

    const item = await UserWorkspaceService.addItem(userId, {
      type,
      title:
        title ||
        (type === 'file'
          ? t('ws.format.defaultTitleFile', { fileId: String(fileId) })
          : t('ws.format.defaultTitleQuery', { query: String(query) })),
      payload: { fileId: fileId || undefined, query: query || undefined },
    });

    await replyWithPrivacy(interaction, {
      content: t('ws.reply.addSuccess', { title: item.title, id: item.id }),
    });
  }

  private async handleList(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const type = interaction.options.getString('type') as WorkspaceItemType | null;

    const items = await UserWorkspaceService.list(
      userId,
      type ? { type } : undefined
    );
    if (!items.length) {
      await replyWithPrivacy(interaction, { content: t('ws.reply.listEmpty') });
      return;
    }

    const lines = items.map(i =>
      t('ws.format.listLine', { type: i.type, title: i.title, id: i.id })
    );
    await replyWithPrivacy(interaction, {
      content: `${t('ws.reply.listTitle')}\n${lines.join('\n')}`,
    });
  }

  private async handleRemove(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const id = interaction.options.getString('id', true);

    const ok = await UserWorkspaceService.remove(userId, id);
    await replyWithPrivacy(interaction, {
      content: ok ? t('ws.reply.removed') : t('ws.reply.notFound'),
    });
  }

  private async handleRun(interaction: any): Promise<void> {
    const userId = interaction.user.id;
    const id = interaction.options.getString('id', true);

    const item = await UserWorkspaceService.get(userId, id);
    if (!item) {
      await replyWithPrivacy(interaction, { content: t('ws.reply.notFound') });
      return;
    }

    if (item.type === 'query') {
      await replyWithPrivacy(interaction, {
        content: t('ws.reply.runQuery', { query: String(item.payload.query) }),
      });
      return;
    }

    if (item.type === 'file') {
      await replyWithPrivacy(interaction, {
        content: t('ws.reply.runFile', { fileId: String(item.payload.fileId) }),
      });
      return;
    }

    await replyWithPrivacy(interaction, { content: t('ws.error.unsupportedType') });
  }
}
