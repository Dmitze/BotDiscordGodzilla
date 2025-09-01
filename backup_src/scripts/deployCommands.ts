/**
 * Скрипт для реєстрації команд в Discord
 * Використовується для розгортання slash-команд
 */

import 'tsconfig-paths/register';
import { config } from 'dotenv';

// Завантаження змінних середовища
config();

type Mode = 'global' | 'guild' | 'both';
interface DeployOptions {
  dry?: boolean;
  mode?: Mode;
  guildId?: string;
}

function parseArgs(argv: string[]): DeployOptions {
  const opts: DeployOptions = { dry: true, mode: 'both' };
  for (const arg of argv) {
    if (arg === '--dry') opts.dry = true;
    else if (arg === '--execute' || arg === '--no-dry') opts.dry = false;
    else if (arg.startsWith('--mode=')) {
      const m = arg.split('=')[1] as Mode;
      if (m === 'global' || m === 'guild' || m === 'both') opts.mode = m;
    } else if (arg.startsWith('--guild=')) {
      const parts = arg.split('=');
      const v = parts.length > 1 ? parts[1] : undefined;
      if (v && v.length > 0) opts.guildId = v;
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
    const { dry = true } = options;
    console.log(
      `🚀 Початок реєстрації команд в Discord... (dry=${dry}, mode=${options.mode || 'both'})`
    );

    // Легкий dry-run: уникаємо важких імпортів/ініціалізацій
    if (dry) {
      const mode: Mode = options.mode || 'both';
      const envGuild = process.env['DISCORD_GUILD_ID'] || '';
      const targets: string[] = [];
      if (mode === 'global' || mode === 'both') targets.push('global');
      if ((mode === 'guild' || mode === 'both') && envGuild)
        targets.push(`guild:${maskId(envGuild)}`);
      console.log('🧪 Режим dry-run (light): пропущено Config.load() та створення команд');
      console.log(`🎯 Цілі: ${targets.join(', ') || '—'}`);
      console.log('📦 Команди: пропущено побудову (light mode)');
      return;
    }

    // Динамічні імпорти тільки для виконання
    const [{ REST, Routes }, { Config }] = await Promise.all([
      import('discord.js'),
      import('@/config/Config'),
    ]);

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
    const { OCRCommand } = await import('@/commands/OCRCommand');
    const { DriveExtractCommand } = await import('@/commands/DriveExtractCommand');
    const { DocCommand } = await import('@/commands/DocCommand');

    const commands = [
      new SearchCommand(botConfig),
      new PerformanceCommand(botConfig),
      new AIAssistantCommand(botConfig),
      new DocumentsCommand(botConfig),
      new FileManagerCommand(botConfig),
      new OperationsCommand(botConfig),
      new AnalyticsCommand(botConfig),
      new EnhancedSearchCommand(botConfig),
      new OCRCommand(botConfig),
      new DriveExtractCommand(botConfig),
      new DocCommand(botConfig),
    ];

    // Підготовка даних команд
    const commandsData = commands.map(command => command.getData().toJSON());

    console.log(`📋 Підготовлено ${commandsData.length} команд для реєстрації`);

    const mode: Mode = options.mode || 'both';
    const guildId = options.guildId || botConfig.discord.guildId;

    // Валідація режиму/цілей
    if (mode === 'guild' && !guildId) {
      console.error(
        '❌ Помилка: для режиму "guild" необхідно вказати --guild=<ID> або налаштувати discord.guildId у конфігурації'
      );
      process.exit(2);
    }

    // Створення REST клієнта
    const rest = new REST({ version: '10' }).setToken(botConfig.discord.token);

    // Реєстрація команд глобально (за режимом)
    if (mode === 'global' || mode === 'both') {
      console.log('🌍 Реєстрація команд глобально...');
      const globalData = (await rest.put(Routes.applicationCommands(botConfig.discord.clientId), {
        body: commandsData,
      })) as any[];
      console.log(`✅ Успішно зареєстровано ${globalData.length} глобальних команд`);
    }

    // Реєстрація команд для конкретного сервера (за режимом)
    if ((mode === 'guild' || mode === 'both') && guildId) {
      console.log(`🏠 Реєстрація команд для сервера ${maskId(guildId)}...`);
      const guildData = (await rest.put(
        Routes.applicationGuildCommands(botConfig.discord.clientId, guildId),
        { body: commandsData }
      )) as any[];
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
