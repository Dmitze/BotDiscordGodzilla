/**
 * Скрипт для реєстрації команд в Discord
 * Використовується для розгортання slash-команд
 */

import { REST, Routes } from 'discord.js';
import { config } from 'dotenv';
import { Config } from '@/config/Config';

// Завантаження змінних середовища
config();

async function deployCommands() {
  try {
    console.log('🚀 Початок реєстрації команд в Discord...');

    // Завантаження конфігурації
    const botConfig = Config.load();

    // Створення REST клієнта
    const rest = new REST({ version: '10' }).setToken(botConfig.discord.token);

    // Створення екземплярів команд
    const { SearchCommand } = await import('@/commands/SearchCommand');
    const { PerformanceCommand } = await import('@/commands/PerformanceCommand');
    const { AIAssistantCommand } = await import('@/commands/AIAssistantCommand');
    const { DocumentsCommand } = await import('@/commands/DocumentsCommand');
    const { FileManagerCommand } = await import('@/commands/FileManagerCommand');
    const { OperationsCommand } = await import('@/commands/OperationsCommand');
    const { AnalyticsCommand } = await import('@/commands/AnalyticsCommand');
    const { EnhancedSearchCommand } = await import('@/commands/EnhancedSearchCommand');

    const commands = [
      new SearchCommand(botConfig),
      new PerformanceCommand(botConfig),
      new AIAssistantCommand(botConfig),
      new DocumentsCommand(botConfig),
      new FileManagerCommand(botConfig),
      new OperationsCommand(botConfig),
      new AnalyticsCommand(botConfig),
      new EnhancedSearchCommand(botConfig)
    ];

    // Підготовка даних команд
    const commandsData = commands.map(command => command.getData().toJSON());

    console.log(`📋 Підготовлено ${commandsData.length} команд для реєстрації`);

    // Реєстрація команд глобально
    console.log('🌍 Реєстрація команд глобально...');
    
    const globalData = await rest.put(
      Routes.applicationCommands(botConfig.discord.clientId),
      { body: commandsData }
    ) as any[];

    console.log(`✅ Успішно зареєстровано ${globalData.length} глобальних команд`);

    // Реєстрація команд для конкретного сервера (якщо вказано guildId)
    if (botConfig.discord.guildId) {
      console.log(`🏠 Реєстрація команд для сервера ${botConfig.discord.guildId}...`);
      
      const guildData = await rest.put(
        Routes.applicationGuildCommands(botConfig.discord.clientId, botConfig.discord.guildId),
        { body: commandsData }
      ) as any[];

      console.log(`✅ Успішно зареєстровано ${guildData.length} команд для сервера`);
    }

    console.log('🎉 Реєстрація команд завершена успішно!');
    console.log('\n📊 Статистика команд:');
    
    commands.forEach(command => {
      console.log(`  - ${command.getName()}: ${command.getDescription()}`);
    });

  } catch (error) {
    console.error('❌ Помилка реєстрації команд:', error);
    process.exit(1);
  }
}

// Запуск скрипта
if (require.main === module) {
  deployCommands();
}

export { deployCommands }; 