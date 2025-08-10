"use strict";
/**
 * ⚔️ Команди оперативного управління ЗСУ
 * Спеціалізовані функції для оперативної роботи
 */
Object.defineProperty(exports, "__esModule", { value: true });
exports.OperationsCommand = void 0;
const discord_js_1 = require("discord.js");
const BaseCommand_1 = require("./BaseCommand");
class OperationsCommand extends BaseCommand_1.BaseCommand {
    constructor(config) {
        super('операції', '⚔️ Оперативне управління ЗСУ', config, (builder) => {
            return builder
                .addSubcommand((subcommand) => subcommand
                .setName('ситуація')
                .setDescription('📊 Поточна оперативна ситуація')
                .addStringOption((option) => option
                .setName('сектор')
                .setDescription('Оперативний сектор')
                .setRequired(false)
                .addChoices({ name: 'Всі сектори', value: 'all' }, { name: 'Сектор А', value: 'A' }, { name: 'Сектор Б', value: 'B' }, { name: 'Сектор В', value: 'C' }, { name: 'Сектор Г', value: 'D' })))
                .addSubcommand((subcommand) => subcommand
                .setName('завдання')
                .setDescription('🎯 Управління завданнями')
                .addStringOption((option) => option
                .setName('дія')
                .setDescription('Дія з завданнями')
                .setRequired(true)
                .addChoices({ name: 'Поточні завдання', value: 'current' }, { name: 'Нове завдання', value: 'new' }, { name: 'Оновити статус', value: 'update' }, { name: 'Завершити завдання', value: 'complete' }, { name: 'Архів завдань', value: 'archive' }))
                .addStringOption((option) => option.setName('запит').setDescription('Пошуковий запит або дані').setRequired(false)))
                .addSubcommand((subcommand) => subcommand
                .setName('координація')
                .setDescription('🔄 Координація між підрозділами')
                .addStringOption((option) => option
                .setName('тип')
                .setDescription('Тип координації')
                .setRequired(true)
                .addChoices({ name: 'Вогнева підтримка', value: 'fire_support' }, { name: 'Логістика', value: 'logistics' }, { name: 'Розвідка', value: 'intelligence' }, { name: 'Медична допомога', value: 'medical' }, { name: "Зв'язок", value: 'communications' }))
                .addStringOption((option) => option.setName('підрозділ').setDescription('Підрозділ для координації').setRequired(false)))
                .addSubcommand((subcommand) => subcommand
                .setName('розвідка')
                .setDescription('🔍 Розвідувальні дані')
                .addStringOption((option) => option
                .setName('тип')
                .setDescription('Тип розвідки')
                .setRequired(true)
                .addChoices({ name: 'Повітряна розвідка', value: 'air' }, { name: 'Наземна розвідка', value: 'ground' }, { name: 'Технічна розвідка', value: 'technical' }, { name: 'Агентурна розвідка', value: 'agent' }, { name: 'Зведена розвідка', value: 'summary' }))
                .addStringOption((option) => option.setName('район').setDescription('Район розвідки').setRequired(false)))
                .addSubcommand((subcommand) => subcommand
                .setName('зв\'язок')
                .setDescription('📡 Управління зв\'язком')
                .addStringOption((option) => option
                .setName('дія')
                .setDescription('Дія зі зв\'язком')
                .setRequired(true)
                .addChoices({ name: 'Статус зв\'язку', value: 'status' }, { name: 'Налаштування каналів', value: 'channels' }, { name: 'Передача повідомлення', value: 'message' }, { name: 'Перевірка якості', value: 'quality' }, { name: 'Резервні канали', value: 'backup' }))
                .addStringOption((option) => option.setName('канал').setDescription('Канал зв\'язку').setRequired(false))
                .addStringOption((option) => option.setName('повідомлення').setDescription('Текст повідомлення').setRequired(false)));
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
                case 'ситуація':
                    await this.handleSituation(interaction);
                    break;
                case 'завдання':
                    await this.handleTasks(interaction);
                    break;
                case 'координація':
                    await this.handleCoordination(interaction);
                    break;
                case 'розвідка':
                    await this.handleIntelligence(interaction);
                    break;
                case 'зв\'язок':
                    await this.handleCommunications(interaction);
                    break;
                default:
                    await interaction.reply('❌ Невідома підкоманда');
            }
        }
        catch (error) {
            console.error('❌ Помилка команди операцій:', error);
            await interaction.reply('❌ Помилка оперативного управління');
        }
    }
    /**
     * Обробка оперативної ситуації
     */
    async handleSituation(interaction) {
        const sector = interaction.options.getString('сектор') || 'all';
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('📊 Оперативна ситуація')
            .setColor(0xff6b6b)
            .setTimestamp();
        if (sector === 'all') {
            embed.setDescription('**Загальна оперативна ситуація**');
            embed.addFields({ name: 'Сектор А', value: '✅ Стабільна ситуація', inline: true }, { name: 'Сектор Б', value: '⚠️ Активні дії', inline: true }, { name: 'Сектор В', value: '✅ Контрольована ситуація', inline: true }, { name: 'Сектор Г', value: '🟡 Потребує уваги', inline: true });
        }
        else {
            embed.setDescription(`**Оперативна ситуація в секторі ${sector}**`);
            embed.addFields({ name: 'Статус', value: '✅ Стабільна ситуація', inline: true }, { name: 'Активні завдання', value: '3', inline: true }, { name: 'Підрозділи', value: '5', inline: true });
        }
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Обробка завдань
     */
    async handleTasks(interaction) {
        const action = interaction.options.getString('дія', true);
        const query = interaction.options.getString('запит');
        const taskOptions = {
            action,
            query: query || undefined,
        };
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('🎯 Управління завданнями')
            .setColor(0x0099ff)
            .setTimestamp();
        const actionName = this.getTaskActionName(action);
        switch (action) {
            case 'current':
                embed.setDescription('**Поточні завдання**');
                embed.addFields({ name: 'Активні завдання', value: '5', inline: true }, { name: 'В процесі', value: '3', inline: true }, { name: 'Очікують', value: '2', inline: true });
                break;
            case 'new':
                embed.setDescription(`**Нове завдання**\n\nДані: ${query || 'Не вказано'}`);
                embed.addFields({ name: 'Статус', value: '✅ Завдання створено', inline: false });
                break;
            case 'update':
                embed.setDescription(`**Оновлення статусу**\n\nДані: ${query || 'Не вказано'}`);
                embed.addFields({ name: 'Статус', value: '✅ Статус оновлено', inline: false });
                break;
            case 'complete':
                embed.setDescription(`**Завершення завдання**\n\nДані: ${query || 'Не вказано'}`);
                embed.addFields({ name: 'Статус', value: '✅ Завдання завершено', inline: false });
                break;
            case 'archive':
                embed.setDescription('**Архів завдань**');
                embed.addFields({ name: 'Завершені', value: '15', inline: true }, { name: 'Архівовані', value: '8', inline: true });
                break;
            default:
                embed.setDescription('❌ Невідома дія');
        }
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Обробка координації
     */
    async handleCoordination(interaction) {
        const type = interaction.options.getString('тип', true);
        const unit = interaction.options.getString('підрозділ');
        const coordinationOptions = {
            type,
            unit: unit || undefined,
        };
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('🔄 Координація між підрозділами')
            .setColor(0xff9900)
            .setTimestamp();
        const typeName = this.getCoordinationTypeName(type);
        embed.setDescription(`**${typeName}**\n\nПідрозділ: ${unit || 'Всі підрозділи'}`);
        embed.addFields({ name: 'Статус координації', value: '✅ Активна', inline: true }, { name: 'Учасники', value: '3 підрозділи', inline: true }, { name: 'Канал зв\'язку', value: 'Основний', inline: true });
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Обробка розвідки
     */
    async handleIntelligence(interaction) {
        const type = interaction.options.getString('тип', true);
        const area = interaction.options.getString('район');
        const intelligenceOptions = {
            type,
            area: area || undefined,
        };
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('🔍 Розвідувальні дані')
            .setColor(0x00ff88)
            .setTimestamp();
        const typeName = this.getIntelligenceTypeName(type);
        embed.setDescription(`**${typeName}**\n\nРайон: ${area || 'Всі райони'}`);
        embed.addFields({ name: 'Останні дані', value: '2 години тому', inline: true }, { name: 'Достовірність', value: 'Висока', inline: true }, { name: 'Джерело', value: 'Підтверджено', inline: true });
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Обробка зв'язку
     */
    async handleCommunications(interaction) {
        const action = interaction.options.getString('дія', true);
        const channel = interaction.options.getString('канал');
        const message = interaction.options.getString('повідомлення');
        const communicationOptions = {
            action,
            channel: channel || undefined,
            message: message || undefined,
        };
        const embed = new discord_js_1.EmbedBuilder()
            .setTitle('📡 Управління зв\'язком')
            .setColor(0x9932cc)
            .setTimestamp();
        const actionName = this.getCommunicationActionName(action);
        switch (action) {
            case 'status':
                embed.setDescription('**Статус зв\'язку**');
                embed.addFields({ name: 'Основний канал', value: '✅ Працює', inline: true }, { name: 'Резервний канал', value: '✅ Готовий', inline: true }, { name: 'Якість сигналу', value: 'Висока', inline: true });
                break;
            case 'channels':
                embed.setDescription('**Налаштування каналів**');
                embed.addFields({ name: 'Активні канали', value: '3', inline: true }, { name: 'Резервні канали', value: '2', inline: true });
                break;
            case 'message':
                embed.setDescription(`**Передача повідомлення**\n\nКанал: ${channel || 'Основний'}\nПовідомлення: ${message || 'Не вказано'}`);
                embed.addFields({ name: 'Статус', value: '✅ Повідомлення передано', inline: false });
                break;
            case 'quality':
                embed.setDescription('**Перевірка якості зв\'язку**');
                embed.addFields({ name: 'Якість сигналу', value: '95%', inline: true }, { name: 'Затримка', value: '50ms', inline: true }, { name: 'Стабільність', value: 'Висока', inline: true });
                break;
            case 'backup':
                embed.setDescription('**Резервні канали**');
                embed.addFields({ name: 'Канал 1', value: '✅ Активний', inline: true }, { name: 'Канал 2', value: '✅ Готовий', inline: true });
                break;
            default:
                embed.setDescription('❌ Невідома дія');
        }
        await interaction.reply({ embeds: [embed] });
    }
    /**
     * Отримання назви дії завдання
     */
    getTaskActionName(action) {
        const actionNames = {
            current: 'Поточні завдання',
            new: 'Нове завдання',
            update: 'Оновити статус',
            complete: 'Завершити завдання',
            archive: 'Архів завдань',
        };
        return actionNames[action] || action;
    }
    /**
     * Отримання назви типу координації
     */
    getCoordinationTypeName(type) {
        const typeNames = {
            fire_support: 'Вогнева підтримка',
            logistics: 'Логістика',
            intelligence: 'Розвідка',
            medical: 'Медична допомога',
            communications: "Зв'язок",
        };
        return typeNames[type] || type;
    }
    /**
     * Отримання назви типу розвідки
     */
    getIntelligenceTypeName(type) {
        const typeNames = {
            air: 'Повітряна розвідка',
            ground: 'Наземна розвідка',
            technical: 'Технічна розвідка',
            agent: 'Агентурна розвідка',
            summary: 'Зведена розвідка',
        };
        return typeNames[type] || type;
    }
    /**
     * Отримання назви дії зв'язку
     */
    getCommunicationActionName(action) {
        const actionNames = {
            status: 'Статус зв\'язку',
            channels: 'Налаштування каналів',
            message: 'Передача повідомлення',
            quality: 'Перевірка якості',
            backup: 'Резервні канали',
        };
        return actionNames[action] || action;
    }
}
exports.OperationsCommand = OperationsCommand;
//# sourceMappingURL=OperationsCommand.js.map