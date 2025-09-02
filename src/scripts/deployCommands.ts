/**
 * Скрипт для реєстрації команд в Discord
 * Використовується для розгортання slash-команд
 */

import 'tsconfig-paths/register';
import { config } from 'dotenv';
import { REST, Routes } from 'discord.js';
import { Config } from '@/config/Config';

// Завантаження змінних середовища
config();

type Mode = 'global' | 'guild' | 'both';
interface DeployOptions {
  dry?: boolean;
  mode?: Mode;
  guildId?: string;
}

function parseArgs(argv: string[]): DeployOptions {
  const opts: DeployOptions = { dry: false, mode: 'both' };
  for (const arg of argv) {
    if (arg === '--dry') opts.dry = true;
    else if (arg === '--execute' || arg === '--no-dry') opts.dry = false;
    else if (arg.startsWith('--mode=')) {
      const m = arg.split('=')[1] as Mode;
      if (m === 'global' || m === 'guild' || m === 'both') opts.mode = m;
    } else if (arg.startsWith('--guild=')) {
      opts.guildId = arg.split('=')[1];
    }
  }
  return opts;
}

function maskId(id?: string): string {
  if (!id) return '';
  const s = String(id);
  if (s.length <= 6) return s.replace(/.(?=.{2})/g, '*');
  return s.slice(0, 2) + '***' + s.slice(-4);
}

async function deployCommands(options: DeployOptions = parseArgs(process.argv.slice(2))) {
  try {
    const { dry = false } = options;
    console.log(`🚀 Початок реєстрації команд в Discord... (dry=${dry}, mode=${options.mode || 'both'})`);

    // Завантаження конфігурації
    const botConfig = Config.load();

    // Створення екземплярів команд
    const { SearchCommand } = await import('@/commands/SearchCommand');
    const { PerformanceCommand } = await import('@/commands/PerformanceCommand');
    const { AIAssistantCommand } = await import('@/commands/AIAssistantCommand');
    const { DocumentsCommand } = await import('@/commands/DocumentsCommand');
    const { FileManagerCommand } = await import('@/commands/FileManagerCommand');
    const { OperationsCommand } = await import('@/commands/OperationsCommand');
    const { AnalyticsCommand } = await import('@/commands/AnalyticsCommand');
    const { EnhancedSearchCommand } = await import('@/commands/EnhancedSearchCommand');
    const { WeatherCommand } = await import('@/commands/WeatherCommand');

    const commands = [
      new SearchCommand(botConfig),
      new PerformanceCommand(botConfig),
      new AIAssistantCommand(botConfig),
      new DocumentsCommand(botConfig),
      new FileManagerCommand(botConfig),
      new OperationsCommand(botConfig),
      new AnalyticsCommand(botConfig),
      new EnhancedSearchCommand(botConfig),
      new WeatherCommand(botConfig)
    ];

    // Підготовка даних команд
    const commandsData = commands.map(command => command.getData().toJSON());

    console.log(`📋 Підготовлено ${commandsData.length} команд для реєстрації`);

    const mode: Mode = options.mode || 'both';
    const guildId = options.guildId || botConfig.discord.guildId;

    if (dry) {
      console.log('🧪 Режим dry-run: реєстрація НЕ буде виконана');
      const targets: string[] = [];
      if (mode === 'global' || mode === 'both') targets.push('global');
      if ((mode === 'guild' || mode === 'both') && guildId) targets.push(`guild:${maskId(guildId)}`);
      console.log(`🎯 Цілі: ${targets.join(', ') || '—'}`);
      console.log('📦 Команди:');
      commands.forEach(c => console.log(`  - ${c.getName()}`));
      return;
    }

    // Створення REST клієнта
    const rest = new REST({ version: '10' }).setToken(botConfig.discord.token);

    // Реєстрація команд глобально (за режимом)
    if (mode === 'global' || mode === 'both') {
      console.log('🌍 Реєстрація команд глобально...');
      const globalData = await rest.put(
        Routes.applicationCommands(botConfig.discord.clientId),
        { body: commandsData }
      ) as any[];
      console.log(`✅ Успішно зареєстровано ${globalData.length} глобальних команд`);
    }

    // Реєстрація команд для конкретного сервера (за режимом)
    if ((mode === 'guild' || mode === 'both') && guildId) {
      console.log(`🏠 Реєстрація команд для сервера ${maskId(guildId)}...`);
      const guildData = await rest.put(
        Routes.applicationGuildCommands(botConfig.discord.clientId, guildId),
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
