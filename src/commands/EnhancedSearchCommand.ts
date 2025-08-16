/**
 * 🔍 Покращений пошук з діапазонами та сортуванням
 * Розширені можливості пошуку та фільтрації даних
 */

// No runtime imports needed from discord.js here
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { GoogleService } from '@/services/GoogleService';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

export class EnhancedSearchCommand extends BaseCommand {
  private readonly googleService: GoogleService | undefined;
  constructor(config: BotConfig, googleService?: GoogleService) {
    super(
      'розширений_пошук',
      '🔍 Покращений пошук з діапазонами та сортуванням',
      config,
      {},
      (builder: any) => {
        return builder
          .addStringOption((option: any) =>
            option
              .setName('запит')
              .setDescription('Текст запиту пошуку')
              .setRequired(false)
              .setMaxLength(200)
          )
          .addIntegerOption((option: any) =>
            option
              .setName('ціна_від')
              .setDescription('Мінімальна ціна')
              .setRequired(false)
              .setMinValue(0)
          )
          .addIntegerOption((option: any) =>
            option
              .setName('ціна_до')
              .setDescription('Максимальна ціна')
              .setRequired(false)
              .setMinValue(0)
          )
          .addStringOption((option: any) =>
            option
              .setName('дата_від')
              .setDescription('Початкова дата (YYYY-MM-DD)')
              .setRequired(false)
              .setMaxLength(20)
          )
          .addStringOption((option: any) =>
            option
              .setName('дата_до')
              .setDescription('Кінцева дата (YYYY-MM-DD)')
              .setRequired(false)
              .setMaxLength(20)
          )
          .addIntegerOption((option: any) =>
            option
              .setName('ліміт')
              .setDescription('Максимальна кількість результатів')
              .setRequired(false)
              .setMinValue(1)
          )
          .addIntegerOption((option: any) =>
            option
              .setName('сторінка')
              .setDescription('Номер сторінки результатів')
              .setRequired(false)
              .setMinValue(1)
          )
          .addStringOption((option: any) =>
            option
              .setName('сортування')
              .setDescription('Поле для сортування')
              .setRequired(false)
          )
          .addStringOption((option: any) =>
            option
              .setName('порядок')
              .setDescription('Порядок сортування (asc/desc)')
              .setRequired(false)
          );
      }
    );
    this.googleService = googleService;
  }

  /**
   * Виконання команди
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    try {
      // Тести очікують як 'запит', так і спадний варіант 'номенклатура'
      const query = interaction.options.getString('запит')
        ?? interaction.options.getString('номенклатура');
      if (!query || query.trim().length === 0) {
        await interaction.reply({
          content: 'Будь ласка, вкажіть запит для пошуку',
          ephemeral: true,
        });
        return;
      }

      const priceFrom = interaction.options.getInteger('ціна_від') ?? undefined;
      const priceTo = interaction.options.getInteger('ціна_до') ?? undefined;
      const dateFrom = interaction.options.getString('дата_від') ?? undefined;
      const dateTo = interaction.options.getString('дата_до') ?? undefined;
      const limit = interaction.options.getInteger('ліміт') ?? undefined;
      const page = interaction.options.getInteger('сторінка') ?? undefined;
      const sortBy = interaction.options.getString('сортування') ?? undefined;
      const order = interaction.options.getString('порядок') ?? undefined;

      const service: any = interaction.client?.serviceContainer?.get?.('google');
      const google = service || this.googleService;

      if (!google || typeof (google as any).enhancedSearch !== 'function') {
        await interaction.reply({ content: 'Помилка: сервіс пошуку недоступний', ephemeral: true });
        return;
      }

      const params: Record<string, any> = { query };
      if (priceFrom !== undefined) params['priceFrom'] = priceFrom;
      if (priceTo !== undefined) params['priceTo'] = priceTo;
      if (dateFrom !== undefined) params['dateFrom'] = dateFrom;
      if (dateTo !== undefined) params['dateTo'] = dateTo;
      if (limit !== undefined) params['limit'] = limit;
      if (page !== undefined) params['page'] = page;
      if (sortBy !== undefined) params['sortBy'] = sortBy;
      if (order !== undefined) params['order'] = order;

      let result: any;
      try {
        result = await (google as any).enhancedSearch(params);
      } catch (e) {
        logger.error('EnhancedSearchCommand: service error', { error: String(e) });
        await interaction.reply({ content: 'Помилка при пошуку', ephemeral: true });
        return;
      }

      // Підтримка двох форматів відповіді: масив або обʼєкт з пагінацією
      let items: any[] = Array.isArray(result) ? result : result?.data || [];
      if (!items || items.length === 0) {
        await interaction.reply({ content: 'Результатів не знайдено', ephemeral: true });
        return;
      }

      // Формуємо коротку відповідь
      const lines = items.slice(0, limit ?? 10).map((it: any) => `• ${it.name ?? it.id ?? 'запис'}`);
      let content = lines.join('\n');
      if (!Array.isArray(result) && result?.totalPages && result?.page) {
        content = `Сторінка ${result.page} з ${result.totalPages}\n` + content;
      }

      await interaction.reply({ content });
    } catch (error) {
      logger.error('Помилка покращеного пошуку', {
        error: error instanceof Error ? error.message : String(error),
        userId: options.interaction.user?.id,
      });
      await options.interaction.reply({ content: '❌ Помилка при виконанні пошуку', ephemeral: true });
    }
  }
}
