/**
 * Менеджер команд Discord бота
 * Централізоване управління всіма командами
 */

import fs from 'fs/promises';
import path from 'path';
import { Collection, ChatInputCommandInteraction } from 'discord.js';
import type { BotConfig } from '@/types';
import { BaseCommand } from '@/commands/BaseCommand';

// Імпорт всіх команд
import { SearchCommand } from '@/commands/SearchCommand';
import { PerformanceCommand } from '@/commands/PerformanceCommand';
import { AIAssistantCommand } from '@/commands/AIAssistantCommand';
import { DocumentsCommand } from '@/commands/DocumentsCommand';
import { FileManagerCommand } from '@/commands/FileManagerCommand';
import { OperationsCommand } from '@/commands/OperationsCommand';
import { AnalyticsCommand } from '@/commands/AnalyticsCommand';
import { EnhancedSearchCommand } from '@/commands/EnhancedSearchCommand';

interface CommandStats {
  totalCommands: number;
  categories: number;
  commandsByCategory: Record<string, number>;
  lastUsed: Date;
}

export class CommandManager {
  private bot: any;
  private config: BotConfig;
  private commands: Collection<string, BaseCommand>;
  private commandCategories: Map<string, string[]>;
  private stats: CommandStats;

  constructor(bot: any, config: BotConfig) {
    this.bot = bot;
    this.config = config;
    this.commands = new Collection();
    this.commandCategories = new Map();
    this.stats = {
      totalCommands: 0,
      categories: 0,
      commandsByCategory: {},
      lastUsed: new Date()
    };
  }

  /**
   * Ініціалізація менеджера команд
   */
  async initialize(): Promise<void> {
    try {
      console.log('📋 Ініціалізація менеджера команд...');

      // Завантаження команд
      await this.loadCommands();

      // Реєстрація обробників подій
      this.registerEventHandlers();

      console.log(`✅ Завантажено ${this.commands.size} команд`);
    } catch (error) {
      console.error('❌ Помилка ініціалізації менеджера команд:', error);
      throw error;
    }
  }

  /**
   * Завантаження всіх команд
   */
  private async loadCommands(): Promise<void> {
    try {
      // Створюємо екземпляри всіх команд
      const commandInstances = [
        new SearchCommand(this.config),
        new PerformanceCommand(this.config),
        new AIAssistantCommand(this.config),
        new DocumentsCommand(this.config),
        new FileManagerCommand(this.config),
        new OperationsCommand(this.config),
        new AnalyticsCommand(this.config),
        new EnhancedSearchCommand(this.config)
      ];

      // Реєструємо команди
      for (const command of commandInstances) {
        if (this.validateCommand(command)) {
          const commandName = command.getName();
          this.commands.set(commandName, command);

          // Категоризація команд
          const category = this.getCommandCategory(command);
          if (!this.commandCategories.has(category)) {
            this.commandCategories.set(category, []);
          }
          this.commandCategories.get(category)!.push(commandName);

          console.log(`📝 Завантажено команду: ${commandName} (${category})`);
        }
      }

      // Оновлюємо статистику
      this.updateStats();
    } catch (error) {
      console.error('❌ Помилка завантаження команд:', error);
      throw error;
    }
  }

  /**
   * Валідація команди
   */
  private validateCommand(command: BaseCommand): boolean {
    if (!command.getName()) {
      console.warn('Команда не має назви');
      return false;
    }

    if (!command.getDescription()) {
      console.warn(`Команда ${command.getName()} не має опису`);
      return false;
    }

    return true;
  }

  /**
   * Визначення категорії команди
   */
  private getCommandCategory(command: BaseCommand): string {
    const name = command.getName();
    
    if (name.includes('пошук') || name.includes('search')) {
      return 'Пошук';
    }
    if (name.includes('продуктивність') || name.includes('performance')) {
      return 'Моніторинг';
    }
    if (name.includes('ai') || name.includes('асистент')) {
      return 'AI';
    }
    if (name.includes('документи') || name.includes('documents')) {
      return 'Документи';
    }
    if (name.includes('файли') || name.includes('file')) {
      return 'Файли';
    }
    if (name.includes('операції') || name.includes('operations')) {
      return 'Операції';
    }
    if (name.includes('аналітика') || name.includes('analytics')) {
      return 'Аналітика';
    }
    
    return 'Інші';
  }

  /**
   * Реєстрація обробників подій
   */
  private registerEventHandlers(): void {
    this.bot.on('interactionCreate', async (interaction: any) => {
      if (interaction.isChatInputCommand()) {
        await this.handleCommand(interaction);
      }
    });
  }

  /**
   * Обробка команди
   */
  private async handleCommand(interaction: ChatInputCommandInteraction): Promise<void> {
    try {
      const commandName = interaction.commandName;
      const command = this.commands.get(commandName);

      if (!command) {
        await interaction.reply({
          content: '❌ Команда не знайдена',
          ephemeral: true
        });
        return;
      }

      // Оновлюємо статистику
      this.stats.lastUsed = new Date();

      // Перевірка прав доступу
      const hasPermission = await this.checkPermissions(interaction, command);
      if (!hasPermission) {
        await interaction.reply({
          content: '❌ Недостатньо прав для виконання цієї команди',
          ephemeral: true
        });
        return;
      }

      // Виконання команди
      await command.execute({
        interaction
      });

      console.log(`✅ Команда ${commandName} виконана користувачем ${interaction.user.tag}`);

    } catch (error) {
      console.error(`❌ Помилка виконання команди ${interaction.commandName}:`, error);
      
      const errorMessage = '❌ Помилка при виконанні команди. Спробуйте ще раз або зверніться до адміністратора.';
      
      if (interaction.replied || interaction.deferred) {
        await interaction.editReply({ content: errorMessage });
      } else {
        await interaction.reply({ content: errorMessage, ephemeral: true });
      }
    }
  }

  /**
   * Перевірка прав доступу
   */
  private async checkPermissions(interaction: ChatInputCommandInteraction, command: BaseCommand): Promise<boolean> {
    // TODO: Реалізувати перевірку прав доступу
    // Тимчасова реалізація - дозволяємо всім
    return true;
  }

  /**
   * Отримання команди за назвою
   */
  getCommand(name: string): BaseCommand | undefined {
    return this.commands.get(name);
  }

  /**
   * Отримання всіх команд
   */
  getAllCommands(): Collection<string, BaseCommand> {
    return this.commands;
  }

  /**
   * Отримання команд за категорією
   */
  getCommandsByCategory(category: string): string[] {
    return this.commandCategories.get(category) || [];
  }

  /**
   * Отримання всіх категорій
   */
  getCategories(): string[] {
    return Array.from(this.commandCategories.keys());
  }

  /**
   * Отримання статистики
   */
  getStats(): CommandStats {
    return { ...this.stats };
  }

  /**
   * Оновлення статистики
   */
  private updateStats(): void {
    this.stats.totalCommands = this.commands.size;
    this.stats.categories = this.commandCategories.size;
    this.stats.commandsByCategory = {};

    for (const [category, commands] of this.commandCategories.entries()) {
      this.stats.commandsByCategory[category] = commands.length;
    }
  }

  /**
   * Отримання даних для реєстрації команд в Discord
   */
  getCommandsData(): any[] {
    return Array.from(this.commands.values()).map(command => command.getData());
  }

  /**
   * Перезавантаження команд
   */
  async reloadCommands(): Promise<void> {
    console.log('🔄 Перезавантаження команд...');
    
    this.commands.clear();
    this.commandCategories.clear();
    
    await this.loadCommands();
    
    console.log(`✅ Перезавантажено ${this.commands.size} команд`);
  }
} 