"use strict";
/**
 * Валідатор команд Discord бота
 * Централізована логіка валідації та санітизації
 * Версія 1.0.0 - Виокремлено з BaseCommand
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.CommandValidator = void 0;
const security_1 = require("@/utils/security");
const logger_1 = __importDefault(require("@/utils/logger"));
class CommandValidator {
    constructor() {
        if (CommandValidator.instance) {
            return CommandValidator.instance;
        }
        CommandValidator.instance = this;
    }
    /**
     * Головна функція валідації команди
     */
    async validateCommand(interaction, rules) {
        try {
            const errors = [];
            const warnings = [];
            const sanitizedValues = {};
            // Валідація базових параметрів
            if (!interaction.commandName) {
                errors.push('Назва команди відсутня');
            }
            if (!interaction.user) {
                errors.push('Користувач не ідентифікований');
            }
            // Валідація опцій команди
            const optionsResult = this.validateOptions(interaction, rules);
            errors.push(...optionsResult.errors);
            warnings.push(...optionsResult.warnings);
            Object.assign(sanitizedValues, optionsResult.sanitizedValues);
            // Валідація дозволів користувача
            const permissionsResult = this.validateUserPermissions(interaction);
            errors.push(...permissionsResult.errors);
            warnings.push(...permissionsResult.warnings);
            // Валідація контексту виконання
            const contextResult = this.validateExecutionContext(interaction);
            errors.push(...contextResult.errors);
            warnings.push(...contextResult.warnings);
            const isValid = errors.length === 0;
            // Логування результату валідації
            if (!isValid) {
                logger_1.default.warn('⚠️ Валідація команди невдала', {
                    command: interaction.commandName,
                    userId: interaction.user?.id,
                    errors,
                    warnings
                });
            }
            else if (warnings.length > 0) {
                logger_1.default.debug('ℹ️ Валідація команди з попередженнями', {
                    command: interaction.commandName,
                    userId: interaction.user?.id,
                    warnings
                });
            }
            return {
                isValid,
                errors,
                warnings,
                sanitizedValues
            };
        }
        catch (error) {
            logger_1.default.error('❌ Помилка валідації команди:', error);
            return {
                isValid: false,
                errors: ['Внутрішня помилка валідації'],
                warnings: []
            };
        }
    }
    /**
     * Валідація опцій команди
     */
    validateOptions(interaction, rules) {
        const errors = [];
        const warnings = [];
        const sanitizedValues = {};
        try {
            // Отримання всіх опцій
            const options = interaction.options.data;
            for (const option of options) {
                const { name, value, type } = option;
                // Перевірка обов'язкових полів
                if (rules?.requiredFields?.includes(name) && (!value || value === '')) {
                    errors.push(`Поле '${name}' є обов'язковим`);
                    continue;
                }
                // Санітизація значень
                if (typeof value === 'string') {
                    const sanitized = this.sanitizeStringValue(value, rules);
                    sanitizedValues[name] = sanitized.value;
                    errors.push(...sanitized.errors);
                    warnings.push(...sanitized.warnings);
                }
                else if (typeof value === 'number') {
                    const validated = this.validateNumberValue(value, name, rules);
                    sanitizedValues[name] = validated.value;
                    errors.push(...validated.errors);
                    warnings.push(...validated.warnings);
                }
                else {
                    sanitizedValues[name] = value;
                }
                // Перевірка дозволених значень
                if (rules?.allowedValues?.[name]) {
                    if (!rules.allowedValues[name].includes(value)) {
                        errors.push(`Недозволене значення для поля '${name}': ${value}`);
                    }
                }
                // Користувацькі валідатори
                if (rules?.customValidators) {
                    for (const validator of rules.customValidators) {
                        const result = validator(value, name);
                        errors.push(...result.errors);
                        warnings.push(...result.warnings);
                    }
                }
            }
            return {
                isValid: errors.length === 0,
                errors,
                warnings,
                sanitizedValues
            };
        }
        catch (error) {
            return {
                isValid: false,
                errors: ['Помилка валідації опцій'],
                warnings: []
            };
        }
    }
    /**
     * Санітизація строкових значень
     */
    sanitizeStringValue(value, rules) {
        const errors = [];
        const warnings = [];
        // Санітизація через security utils
        const sanitized = (0, security_1.sanitizeInput)(value);
        // Перевірка довжини
        if (rules?.maxStringLength && sanitized.length > rules.maxStringLength) {
            errors.push(`Текст занадто довгий (макс. ${rules.maxStringLength} символів)`);
            const truncated = sanitized.substring(0, rules.maxStringLength);
            warnings.push('Текст було обрізано');
            return { value: truncated, errors, warnings };
        }
        // Перевірка на підозрілий контент
        if (this.containsSuspiciousContent(sanitized)) {
            warnings.push('Виявлено потенційно небезпечний контент');
        }
        return { value: sanitized, errors, warnings };
    }
    /**
     * Валідація числових значень
     */
    validateNumberValue(value, fieldName, rules) {
        const errors = [];
        const warnings = [];
        // Перевірка діапазону
        if (rules?.minNumberValue !== undefined && value < rules.minNumberValue) {
            errors.push(`Значення '${fieldName}' занадто мале (мін. ${rules.minNumberValue})`);
        }
        if (rules?.maxNumberValue !== undefined && value > rules.maxNumberValue) {
            errors.push(`Значення '${fieldName}' занадто велике (макс. ${rules.maxNumberValue})`);
        }
        // Перевірка на розумність значення
        if (value < 0 && fieldName.includes('count')) {
            warnings.push(`Від'ємне значення для лічильника: ${fieldName}`);
        }
        return { value, errors, warnings };
    }
    /**
     * Валідація дозволів користувача
     */
    validateUserPermissions(interaction) {
        const errors = [];
        const warnings = [];
        try {
            // Перевірка що користувач існує
            if (!interaction.user) {
                errors.push('Користувач не ідентифікований');
                return { isValid: false, errors, warnings };
            }
            // Перевірка що команда виконується на сервері (якщо потрібно)
            if (!interaction.guild) {
                warnings.push('Команда виконується поза сервером');
            }
            // Перевірка member об'єкта
            if (interaction.guild && !interaction.member) {
                errors.push('Не вдалося отримати інформацію про учасника сервера');
            }
            return {
                isValid: errors.length === 0,
                errors,
                warnings
            };
        }
        catch (error) {
            return {
                isValid: false,
                errors: ['Помилка валідації дозволів'],
                warnings: []
            };
        }
    }
    /**
     * Валідація контексту виконання
     */
    validateExecutionContext(interaction) {
        const errors = [];
        const warnings = [];
        try {
            // Перевірка каналу
            if (!interaction.channel) {
                errors.push('Канал недоступний');
            }
            // Перевірка що interaction не застарілий
            const interactionAge = Date.now() - interaction.createdTimestamp;
            if (interactionAge > 15 * 60 * 1000) { // 15 хвилин
                warnings.push('Interaction застарілий');
            }
            // Перевірка що bot має дозволи у каналі
            if (interaction.guild && interaction.channel) {
                const botMember = interaction.guild.members.me;
                if (botMember && 'permissionsFor' in interaction.channel) {
                    const permissions = interaction.channel.permissionsFor(botMember);
                    if (!permissions?.has(['SendMessages', 'ViewChannel'])) {
                        errors.push('Бот не має необхідних дозволів у цьому каналі');
                    }
                }
            }
            return {
                isValid: errors.length === 0,
                errors,
                warnings
            };
        }
        catch (error) {
            return {
                isValid: false,
                errors: ['Помилка валідації контексту'],
                warnings: []
            };
        }
    }
    /**
     * Перевірка на підозрілий контент
     */
    containsSuspiciousContent(text) {
        const suspiciousPatterns = [
            /discord\.gg\/[a-zA-Z0-9]+/gi, // Discord invite links
            /https?:\/\/[^\s]+/gi, // URLs
            /@everyone|@here/gi, // Mass mentions
            /\b(free|nitro|giveaway)\b/gi, // Suspicious keywords
            /<[@#!&][0-9]+>/gi, // Discord mentions
        ];
        return suspiciousPatterns.some(pattern => pattern.test(text));
    }
    /**
     * Валідація з кастомними правилами
     */
    async validateWithRules(interaction, rules) {
        return this.validateCommand(interaction, rules);
    }
    /**
     * Швидка валідація без складних перевірок
     */
    quickValidate(interaction) {
        const errors = [];
        if (!interaction.commandName)
            errors.push('Назва команди відсутня');
        if (!interaction.user)
            errors.push('Користувач не ідентифікований');
        if (!interaction.channel)
            errors.push('Канал недоступний');
        return {
            isValid: errors.length === 0,
            errors,
            warnings: []
        };
    }
    /**
     * Отримання статистики валідації
     */
    getValidationStats() {
        // TODO: Реалізувати збір статистики
        return {
            totalValidations: 0,
            successfulValidations: 0,
            failedValidations: 0
        };
    }
}
exports.CommandValidator = CommandValidator;
CommandValidator.instance = null;
exports.default = CommandValidator;
//# sourceMappingURL=CommandValidator.js.map