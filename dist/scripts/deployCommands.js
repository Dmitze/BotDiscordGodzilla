"use strict";
/**
 * Скрипт для реєстрації команд в Discord
 * Використовується для розгортання slash-команд
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.deployCommands = deployCommands;
const discord_js_1 = require("discord.js");
const dotenv_1 = require("dotenv");
const Config_1 = require("@/config/Config");
// Завантаження змінних середовища
(0, dotenv_1.config)();
function parseArgs(argv) {
    const opts = { dry: true, mode: 'both' };
    for (const arg of argv) {
        if (arg === '--dry')
            opts.dry = true;
        else if (arg === '--execute' || arg === '--no-dry')
            opts.dry = false;
        else if (arg.startsWith('--mode=')) {
            const m = arg.split('=')[1];
            if (m === 'global' || m === 'guild' || m === 'both')
                opts.mode = m;
        }
        else if (arg.startsWith('--guild=')) {
            const parts = arg.split('=');
            const v = parts.length > 1 ? parts[1] : undefined;
            if (v && v.length > 0)
                opts.guildId = v;
        }
    }
    return opts;
}
function maskId(id) {
    if (!id)
        return '';
    const s = String(id);
    if (s.length <= 6)
        return s.replace(/.(?=.{2})/g, '*');
    return s.slice(0, 2) + '***' + s.slice(-4);
}
async function deployCommands(options = parseArgs(process.argv.slice(2))) {
    try {
        const { dry = true } = options;
        console.log(`🚀 Початок реєстрації команд в Discord... (dry=${dry}, mode=${options.mode || 'both'})`);
        // Завантаження конфігурації
        const botConfig = Config_1.Config.load();
        // Створення екземплярів команд
        const { SearchCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/SearchCommand')));
        const { PerformanceCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/PerformanceCommand')));
        const { AIAssistantCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/AIAssistantCommand')));
        const { DocumentsCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/DocumentsCommand')));
        const { FileManagerCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/FileManagerCommand')));
        const { OperationsCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/OperationsCommand')));
        const { AnalyticsCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/AnalyticsCommand')));
        const { EnhancedSearchCommand } = await Promise.resolve().then(() => __importStar(require('@/commands/EnhancedSearchCommand')));
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
        const mode = options.mode || 'both';
        const guildId = options.guildId || botConfig.discord.guildId;
        if (dry) {
            console.log('🧪 Режим dry-run: реєстрація НЕ буде виконана');
            const targets = [];
            if (mode === 'global' || mode === 'both')
                targets.push('global');
            if ((mode === 'guild' || mode === 'both') && guildId)
                targets.push(`guild:${maskId(guildId)}`);
            console.log(`🎯 Цілі: ${targets.join(', ') || '—'}`);
            console.log('📦 Команди:');
            commands.forEach(c => console.log(`  - ${c.getName()}`));
            return;
        }
        // Валідація режиму/цілей
        if (mode === 'guild' && !guildId) {
            console.error('❌ Помилка: для режиму "guild" необхідно вказати --guild=<ID> або налаштувати discord.guildId у конфігурації');
            process.exit(2);
        }
        // Створення REST клієнта
        const rest = new discord_js_1.REST({ version: '10' }).setToken(botConfig.discord.token);
        // Реєстрація команд глобально (за режимом)
        if (mode === 'global' || mode === 'both') {
            console.log('🌍 Реєстрація команд глобально...');
            const globalData = await rest.put(discord_js_1.Routes.applicationCommands(botConfig.discord.clientId), { body: commandsData });
            console.log(`✅ Успішно зареєстровано ${globalData.length} глобальних команд`);
        }
        // Реєстрація команд для конкретного сервера (за режимом)
        if ((mode === 'guild' || mode === 'both') && guildId) {
            console.log(`🏠 Реєстрація команд для сервера ${maskId(guildId)}...`);
            const guildData = await rest.put(discord_js_1.Routes.applicationGuildCommands(botConfig.discord.clientId, guildId), { body: commandsData });
            console.log(`✅ Успішно зареєстровано ${guildData.length} команд для сервера`);
        }
        console.log('🎉 Реєстрація команд завершена успішно!');
        console.log('\n📊 Статистика команд:');
        commands.forEach(command => {
            console.log(`  - ${command.getName()}: ${command.getDescription()}`);
        });
    }
    catch (error) {
        console.error('❌ Помилка реєстрації команд:', error);
        process.exit(1);
    }
}
// Запуск скрипта
if (require.main === module) {
    deployCommands();
}
//# sourceMappingURL=deployCommands.js.map