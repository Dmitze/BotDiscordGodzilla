import { SlashCommandBuilder, TextChannel, ThreadChannel } from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { AIService } from '@/services/AIService';
import logger from '@/utils/logger';

export default class SummarizeCommand extends BaseCommand {
  private aiService: AIService | undefined;

  constructor(config: BotConfig) {
    super('summarize', 'Генерує AI-самарі переписки та список задач (до 100 повідомлень).', config, {
      category: 'business'
    });
  }

  protected buildCommand(): Omit<SlashCommandBuilder, 'addSubcommand' | 'addSubcommandGroup'> {
    return new SlashCommandBuilder()
      .setName(this.name)
      .setDescription(this.description)
      .addIntegerOption(option => 
        option.setName('count')
          .setDescription('Кількість останніх повідомлень для аналізу (за замовчуванням 50, макс 100)')
          .setRequired(false)
          .setMinValue(5)
          .setMaxValue(100)
      ) as SlashCommandBuilder;
  }

  protected override async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    const count = interaction.options.getInteger('count') || 50;

    if (!this.aiService) {
      const anyClient = interaction.client as any;
      // Depending on where ServiceManager is available on the client
      this.aiService = anyClient.serviceManager?.getService('ai') || anyClient.services?.get('ai') || anyClient.getService?.('ai');
    }

    if (!this.aiService) {
      // Якщо aiService все ще не знайдено, намагаємося дістати його іншим шляхом, але якщо ні — помилка
      await interaction.reply({ content: '❌ AI Сервіс недоступний. Спробуйте пізніше.', ephemeral: true });
      return;
    }

    await interaction.deferReply();

    try {
      const channel = interaction.channel as TextChannel | ThreadChannel;
      if (!channel || typeof channel.messages?.fetch !== 'function') {
         await interaction.editReply('❌ Неможливо отримати повідомлення в цьому каналі.');
         return;
      }

      const messages = await channel.messages.fetch({ limit: count });
      
      const conversation = messages
        .filter(m => !m.author.bot && m.content.trim().length > 0)
        .reverse()
        .map(m => `${m.author.username}: ${m.content}`)
        .join('\n');

      if (!conversation) {
        await interaction.editReply('❌ Не знайдено корисних текстових повідомлень для аналізу (повідомлення ботів ігноруються).');
        return;
      }

      const prompt = `Ти — корпоративний AI-асистент. Проаналізуй наступну переписку з чату і надай:
1. 📝 Коротке самарі (про що йшла мова, головні обговорення та рішення).
2. ✅ Action Items (список задач, хто що має зробити, якщо це обговорювалося).

Переписка:
${conversation}`;

      const response = await this.aiService.generateResponse(prompt, { 
        temperature: 0.3
      } as any);
      
      let replyText = response.content || 'Не вдалося згенерувати самарі.';
      if (replyText.length > 2000) {
        replyText = replyText.substring(0, 1997) + '...';
      }

      await interaction.editReply(replyText);
    } catch (error) {
      logger.error('Помилка генерації самарі', { errorMessage: String(error) } as any);
      await interaction.editReply('❌ Виникла помилка під час аналізу переписки.');
    }
  }
}
