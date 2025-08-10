"use strict";
/**
 * Команда для моніторингу продуктивності
 * Відстеження метрик та оптимізація системи
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.PerformanceCommand = void 0;
const discord_js_1 = require("discord.js");
const BaseCommand_1 = require("./BaseCommand");
class PerformanceCommand extends BaseCommand_1.BaseCommand {
    constructor(config) {
        super('продуктивність', '📊 Моніторинг продуктивності системи', config, (builder) => {
            return builder
                .addSubcommand((subcommand) => subcommand
                .setName('статус')
                .setDescription('Загальний статус продуктивності'))
                .addSubcommand((subcommand) => subcommand
                .setName('кеш')
                .setDescription('Статистика кешування'))
                .addSubcommand((subcommand) => subcommand
                .setName('черги')
                .setDescription('Статистика черг завдань'))
                .addSubcommand((subcommand) => subcommand
                .setName('api')
                .setDescription('Статистика API запитів'))
                .addSubcommand((subcommand) => subcommand
                .setName('оптимізація')
                .setDescription('Рекомендації по оптимізації'));
        });
    }
    /**
     * Виконання команди
     */
    async onExecute(options) {
        const { interaction } = options;
        try {
            const subcommand = interaction.options.getSubcommand();
            switch (subcommand) {
                case 'статус':
                    await this.showGeneralStatus(interaction);
                    break;
                case 'кеш':
                    await this.showCacheStats(interaction);
                    break;
                case 'черги':
                    await this.showQueueStats(interaction);
                    break;
                case 'api':
                    await this.showApiStats(interaction);
                    break;
                case 'оптимізація':
                    await this.showOptimizationRecommendations(interaction);
                    break;
                default:
                    await interaction.reply('❌ Невідома підкоманда');
            }
        }
        catch (error) {
            console.error('❌ Помилка команди продуктивності:', error);
            await interaction.reply('❌ Помилка отримання статистики продуктивності');
        }
    }
    /**
     * Показ загального статусу
     */
    async showGeneralStatus(interaction) {
        const bot = interaction.client.bot;
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('📊 Статус продуктивності системи')
            .setColor(0x00ff00)
            .setTimestamp();
        // Основні метрики
        const memoryUsage = process.memoryUsage();
        const uptime = process.uptime();
        embed.addFields({
            name: '💾 Пам\'ять',
            value: `Використано: ${Math.round(memoryUsage.heapUsed / 1024 / 1024)}MB\nВсього: ${Math.round(memoryUsage.heapTotal / 1024 / 1024)}MB`,
            inline: true
        }, {
            name: '⏱️ Час роботи',
            value: `${Math.floor(uptime / 3600)}г ${Math.floor((uptime % 3600) / 60)}хв`,
            inline: true
        }, {
            name: '🔄 CPU',
            value: `${Math.round(process.cpuUsage().user / 1000000)}ms`,
            inline: true
        });
        // Статистика сервісів
        if (bot?.serviceContainer) {
            const services = bot.serviceContainer.getHealthStatus();
            const healthyServices = Object.values(services).filter((s) => s.healthy).length;
            const totalServices = Object.keys(services).length;
            embed.addFields({
                name: '🔧 Сервіси',
                value: `${healthyServices}/${totalServices} працюють`,
                inline: true
            });
        }
        // Статистика Discord
        if (bot?.client) {
            embed.addFields({
                name: '👥 Користувачі',
                value: bot.client.users.cache.size.toString(),
                inline: true
            }, {
                name: '🏠 Сервери',
                value: bot.client.guilds.cache.size.toString(),
                inline: true
            });
        }
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Показ статистики кешу
     */
    async showCacheStats(interaction) {
        const bot = interaction.client.bot;
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('📋 Статистика кешування')
            .setColor(0x0099ff)
            .setTimestamp();
        if (bot?.serviceContainer) {
            const cacheService = bot.serviceContainer.get('cache');
            if (cacheService) {
                const stats = cacheService.getCacheStats();
                const hitRate = stats.hits / (stats.hits + stats.misses) * 100;
                embed.addFields({
                    name: '🎯 Попадання',
                    value: `${stats.hits} (${hitRate.toFixed(1)}%)`,
                    inline: true
                }, {
                    name: '❌ Промахи',
                    value: stats.misses.toString(),
                    inline: true
                }, {
                    name: '💾 Записи',
                    value: stats.sets.toString(),
                    inline: true
                }, {
                    name: '🗑️ Видалення',
                    value: stats.deletes.toString(),
                    inline: true
                }, {
                    name: '⚠️ Помилки',
                    value: stats.errors.toString(),
                    inline: true
                });
                // Рекомендації
                if (hitRate < 50) {
                    embed.addFields({
                        name: '💡 Рекомендація',
                        value: 'Низький відсоток попадань. Розгляньте збільшення TTL або оптимізацію ключів кешу.',
                        inline: false
                    });
                }
            }
            else {
                embed.setDescription('❌ Сервіс кешування недоступний');
            }
        }
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Показ статистики черг
     */
    async showQueueStats(interaction) {
        const bot = interaction.client.bot;
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('📋 Статистика черг завдань')
            .setColor(0xff9900)
            .setTimestamp();
        if (bot?.queueManager) {
            const stats = bot.queueManager.getQueueStats();
            embed.addFields({
                name: '🔴 Високий пріоритет',
                value: `Завдань: ${stats.high?.length || 0}\nОбробляється: ${stats.high?.processing || 0}`,
                inline: true
            }, {
                name: '🟡 Звичайний пріоритет',
                value: `Завдань: ${stats.normal?.length || 0}\nОбробляється: ${stats.normal?.processing || 0}`,
                inline: true
            }, {
                name: '🟢 Низький пріоритет',
                value: `Завдань: ${stats.low?.length || 0}\nОбробляється: ${stats.low?.processing || 0}`,
                inline: true
            }, {
                name: '📊 Загальна статистика',
                value: `Оброблено: ${stats.processed || 0}\nНевдало: ${stats.failed || 0}\nСередній час: ${Math.round(stats.averageProcessingTime || 0)}ms`,
                inline: false
            });
            // Попередження про довгу чергу
            const totalPending = (stats.high?.length || 0) + (stats.normal?.length || 0) + (stats.low?.length || 0);
            if (totalPending > 50) {
                embed.addFields({
                    name: '⚠️ Попередження',
                    value: 'Довга черга завдань. Розгляньте збільшення кількості workers.',
                    inline: false
                });
            }
        }
        else {
            embed.setDescription('❌ Queue Manager недоступний');
        }
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Показ статистики API
     */
    async showApiStats(interaction) {
        const bot = interaction.client.bot;
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('🌐 Статистика API запитів')
            .setColor(0x9932cc)
            .setTimestamp();
        if (bot?.serviceContainer) {
            // Google API статистика
            const googleService = bot.serviceContainer.get('google');
            if (googleService) {
                const googleStats = googleService.getStats();
                embed.addFields({
                    name: '📊 Google API',
                    value: `Успішні запити: ${googleStats.requests?.success || 0}\nПомилки: ${googleStats.errors?.count || 0}\nСередній час: ${Math.round(googleStats.requests?.averageDuration || 0)}ms`,
                    inline: true
                });
            }
            // AI статистика
            const aiService = bot.serviceContainer.get('ai');
            if (aiService) {
                const aiStats = aiService.getStats();
                embed.addFields({
                    name: '🤖 AI API',
                    value: `Запити: ${aiStats.totalRequests || 0}\nУспішні: ${aiStats.successfulRequests || 0}\nСередній час: ${Math.round(aiStats.averageResponseTime || 0)}ms`,
                    inline: true
                });
            }
        }
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Показ рекомендацій по оптимізації
     */
    async showOptimizationRecommendations(interaction) {
        const bot = interaction.client.bot;
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('💡 Рекомендації по оптимізації')
            .setColor(0x00ff88)
            .setTimestamp();
        const recommendations = [];
        // Перевірка пам'яті
        const memoryUsage = process.memoryUsage();
        const memoryUsageMB = memoryUsage.heapUsed / 1024 / 1024;
        if (memoryUsageMB > 500) {
            recommendations.push('💾 Високе використання пам\'яті. Розгляньте очищення кешу або оптимізацію алгоритмів.');
        }
        // Перевірка кешу
        if (bot?.serviceContainer) {
            const cacheService = bot.serviceContainer.get('cache');
            if (cacheService) {
                const cacheStats = cacheService.getCacheStats();
                const hitRate = cacheStats.hits / (cacheStats.hits + cacheStats.misses) * 100;
                if (hitRate < 60) {
                    recommendations.push('📋 Низький відсоток попадань в кеш. Оптимізуйте стратегію кешування.');
                }
            }
        }
        // Перевірка черг
        if (bot?.queueManager) {
            const queueStats = bot.queueManager.getQueueStats();
            const totalPending = (queueStats.high?.length || 0) + (queueStats.normal?.length || 0) + (queueStats.low?.length || 0);
            if (totalPending > 30) {
                recommendations.push('📋 Довга черга завдань. Збільшіть кількість workers або оптимізуйте завдання.');
            }
        }
        // Перевірка часу відповіді
        const uptime = process.uptime();
        if (uptime > 86400) { // Більше 24 годин
            recommendations.push('⏰ Система працює довго. Розгляньте перезапуск для очищення пам\'яті.');
        }
        if (recommendations.length === 0) {
            recommendations.push('✅ Система працює оптимально!');
        }
        embed.setDescription(recommendations.join('\n\n'));
        await interaction.reply({ embeds: [embed] });
    }
}
exports.PerformanceCommand = PerformanceCommand;
//# sourceMappingURL=PerformanceCommand.js.map