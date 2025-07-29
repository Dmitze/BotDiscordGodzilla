/**
 * Основний клас Discord бота
 * Управляє всіма компонентами та сервісами
 */

import { Client, GatewayIntentBits, Collection } from 'discord.js';
import type { BotConfig, BaseService, CommandInteraction, BaseCommand } from '@/types';
import { ServiceContainer } from './ServiceContainer';
import { BaseService as BaseServiceClass } from './BaseService';

export class Bot extends BaseServiceClass {
  public readonly client: Client;
  public readonly serviceContainer: ServiceContainer;
  private commands = new Collection<string, BaseCommand>();
  private isReady = false;

  constructor(config: BotConfig) {
    super('DiscordBot', config);
    
    this.client = new Client({
      intents: [
        GatewayIntentBits.Guilds,
        GatewayIntentBits.GuildMessages,
        GatewayIntentBits.MessageContent,
        GatewayIntentBits.GuildMembers,
      ],
    });

    this.serviceContainer = new ServiceContainer(config);
    this.setupEventHandlers();
  }

  /**
   * Ініціалізація бота
   */
  protected async onInitialize(): Promise<void> {
    try {
      // Ініціалізація сервісів
      await this.serviceContainer.initialize();
      
      // Підключення до Discord
      await this.client.login(this.config.discord.token);
      
      // Очікування готовності клієнта
      await this.waitForReady();
      
      console.log('🤖 Discord бот успішно ініціалізовано');
    } catch (error) {
      throw new Error(`Помилка ініціалізації бота: ${error}`);
    }
  }

  /**
   * Завершення роботи бота
   */
  protected async onShutdown(): Promise<void> {
    try {
      // Завершення сервісів
      await this.serviceContainer.shutdown();
      
      // Відключення від Discord
      this.client.destroy();
      
      console.log('🤖 Discord бот успішно завершено');
    } catch (error) {
      throw new Error(`Помилка завершення бота: ${error}`);
    }
  }

  /**
   * Health check бота
   */
  protected async onHealthCheck(): Promise<{ healthy: boolean; error?: string; details?: Record<string, unknown> }> {
    try {
      const isConnected = this.client.isReady();
      const servicesHealth = await this.serviceContainer.getHealthStatus();
      
      const allHealthy = Object.values(servicesHealth).every(health => health.healthy);
      
      return {
        healthy: isConnected && allHealthy,
        details: {
          connected: isConnected,
          services: servicesHealth,
          uptime: this.getStats().uptime,
        },
      };
    } catch (error) {
      return {
        healthy: false,
        error: `Health check failed: ${error}`,
      };
    }
  }

  /**
   * Отримання статистики бота
   */
  protected onGetStats(): Partial<{ uptime: number; requests: number; errors: number }> {
    return {
      requests: this.commands.size,
      errors: 0, // Буде оновлюватися в реальному часі
    };
  }

  /**
   * Налаштування обробників подій
   */
  private setupEventHandlers(): void {
    this.client.on('ready', () => {
      this.isReady = true;
      console.log(`🤖 Бот ${this.client.user?.tag} готовий до роботи`);
    });

    this.client.on('interactionCreate', async (interaction) => {
      if (!interaction.isCommand()) return;
      
      try {
        await this.handleCommand(interaction);
      } catch (error) {
        console.error('Помилка обробки команди:', error);
        await this.handleCommandError(interaction, error);
      }
    });

    this.client.on('error', (error) => {
      console.error('Помилка Discord клієнта:', error);
    });

    this.client.on('disconnect', () => {
      this.isReady = false;
      console.log('🔌 Discord клієнт відключено');
    });
  }

  /**
   * Обробка команд
   */
  private async handleCommand(interaction: any): Promise<void> {
    const command = this.commands.get(interaction.commandName);
    if (!command) {
      await interaction.reply({ content: '❌ Команда не знайдена', ephemeral: true });
      return;
    }

    await command.execute(interaction);
  }

  /**
   * Обробка помилок команд
   */
  private async handleCommandError(interaction: any, error: unknown): Promise<void> {
    const errorMessage = error instanceof Error ? error.message : 'Невідома помилка';
    
    try {
      if (interaction.replied || interaction.deferred) {
        await interaction.editReply({ content: `❌ Помилка: ${errorMessage}` });
      } else {
        await interaction.reply({ content: `❌ Помилка: ${errorMessage}`, ephemeral: true });
      }
    } catch (replyError) {
      console.error('Помилка відповіді на помилку:', replyError);
    }
  }

  /**
   * Очікування готовності клієнта
   */
  private waitForReady(): Promise<void> {
    return new Promise((resolve) => {
      if (this.isReady) {
        resolve();
        return;
      }

      this.client.once('ready', () => {
        resolve();
      });
    });
  }

  /**
   * Реєстрація команди
   */
  public registerCommand(command: BaseCommand): void {
    this.commands.set(command.data.setName('').toJSON().name, command);
  }

  /**
   * Отримання всіх команд
   */
  public getCommands(): Collection<string, BaseCommand> {
    return this.commands;
  }

  /**
   * Перевірка чи бот готовий
   */
  public isBotReady(): boolean {
    return this.isReady && this.client.isReady();
  }
} 