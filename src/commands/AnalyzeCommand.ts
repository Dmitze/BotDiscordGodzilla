import { EmbedBuilder } from 'discord.js';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';
import { DuplicateResolver } from '@/components/DuplicateResolver';
import { uiState } from '@/services/UIStateService';

export class AnalyzeCommand extends BaseCommand {
  constructor(config: BotConfig) {
    super(
      'analyze',
      '🧠 Аналіз документів/запитів',
      config,
      {},
      (builder: any) =>
        builder
          .addStringOption((o: any) =>
            o
              .setName('docid')
              .setDescription('ID документа Drive')
              .setRequired(false)
          )
          .addStringOption((o: any) =>
            o
              .setName('query')
              .setDescription('Вільний запит для пошуку/аналізу')
              .setRequired(false)
          )
          .addStringOption((o: any) =>
            o
              .setName('mode')
              .setDescription('Режим аналізу')
              .setRequired(false)
              .addChoices(
                { name: 'Авто', value: 'auto' },
                { name: 'Текст', value: 'text' },
                { name: 'Таблиця', value: 'sheets' }
              )
          )
    );
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    try {
      const docId = interaction.options.getString('docid');
      const query = interaction.options.getString('query');

      const indexer = (interaction.client as any).serviceContainer?.get?.('driveIndexer') as
        | import('@/services/DriveIndexerService').DriveIndexerService
        | undefined;

      if (!indexer) {
        await interaction.reply({ content: '🔎 Індексатор недоступний', ephemeral: true });
        return;
      }

      if (docId) {
        await this.analyzeById(interaction as any, indexer, docId);
        return;
      }

      if (query) {
        const results = await indexer.search(query, 10);
        if (!results.length) {
          await interaction.reply({ content: 'Нічого не знайдено', ephemeral: true });
          return;
        }
        if (results.length === 1) {
          const first = results.at(0);
          if (first?.file?.id) {
            await this.analyzeById(interaction as any, indexer, first.file.id);
            return;
          }
          return;
        }

        // дубли: сохраняем список и отрисовываем DuplicateResolver
        const files = results.map(r => {
          return {
            id: r.file.id,
            name: r.file.name,
            mimeType: r.file.mimeType,
            ...(r.file.owners ? { owners: r.file.owners } : {}),
          } as Pick<import('@/types/drive').DriveFile, 'id' | 'name' | 'mimeType' | 'webViewLink' | 'owners'>;
        });
        const userId = interaction.user?.id ?? '0';
        const nonce = Date.now().toString(36);
        const scope = this.getName();
        const key = uiState.makeKey({ scope, userId, nonce });
        uiState.set(key, files, 300);
        const { embed, rows } = DuplicateResolver.buildPage({ scope, userId, nonce, files, title: 'Знайдено кілька збігів' });
        await interaction.reply({ embeds: [embed], components: rows, ephemeral: true });
        return;
      }

      await interaction.reply({ content: 'Вкажіть docid або query', ephemeral: true });
    } catch (error) {
      logger.error('analyze_execute_error', { error: error instanceof Error ? error.message : String(error) });
      await interaction.reply({ content: '❌ Помилка аналізу', ephemeral: true });
    }
  }

  private async analyzeById(
    interaction: any,
    indexer: import('@/services/DriveIndexerService').DriveIndexerService,
    fileId: string
  ): Promise<void> {
    const entry = await indexer.getEntry(fileId);
    if (!entry) {
      await interaction.reply({ content: 'Документ не індексовано або відсутній', ephemeral: true });
      return;
    }

    const embed = new EmbedBuilder()
      .setTitle(`🧠 Аналіз: ${entry.name}`)
      .setColor(0x2f3136)
      .addFields(
        { name: 'MIME', value: entry.mimeType, inline: true },
        { name: 'Довжина тексту', value: String(entry.textLength), inline: true },
        ...(entry.modifiedTime ? [{ name: 'Змінено', value: entry.modifiedTime, inline: true }] : []),
      )
      .setDescription('Попередній перегляд тексту недоступний у цьому режимі. Використайте пошук для перегляду фрагментів.')
      .setFooter({ text: 'Бета-аналітика — базова метаінформація' });

    await interaction.reply({ embeds: [embed] });
  }

  protected override async onComponent(options: { interaction: any }): Promise<void> {
    const { interaction } = options;
    const customId = (interaction as any).customId as string | undefined;
    if (!customId || !customId.startsWith(DuplicateResolver.PREFIX)) return;

    await DuplicateResolver.handleComponent(interaction as any, {
      fetchFiles: async ({ scope, userId, nonce }) => {
        const key = uiState.makeKey({ scope, userId, nonce });
        return uiState.get<any[]>(key) ?? [];
      },
      onSelect: async ({ fileId }) => {
        try {
          const indexer = (interaction.client as any).serviceContainer?.get?.('driveIndexer') as
            | import('@/services/DriveIndexerService').DriveIndexerService
            | undefined;
          if (!indexer) return;
          await this.analyzeById(interaction, indexer, fileId);
        } catch (e) {
          logger.error('analyze_select_error', { error: e instanceof Error ? e.message : String(e) });
        }
      },
      title: 'Знайдено кілька збігів',
      perPage: 5,
    });
  }
}

export default AnalyzeCommand;
