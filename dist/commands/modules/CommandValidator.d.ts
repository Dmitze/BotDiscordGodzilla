/**
 * Валідатор команд Discord бота
 * Централізована логіка валідації та санітизації
 * Версія 1.0.0 - Виокремлено з BaseCommand
 */
import type { ChatInputCommandInteraction } from 'discord.js';
export interface ValidationResult {
    isValid: boolean;
    errors: string[];
    warnings: string[];
    sanitizedOptions?: any;
    sanitizedValues?: Record<string, unknown>;
}
export interface ValidationRules {
    maxStringLength?: number;
    maxNumberValue?: number;
    minNumberValue?: number;
    requiredFields?: string[];
    allowedValues?: Record<string, unknown[]>;
    customValidators?: Array<(value: unknown, field: string) => ValidationResult>;
}
export declare class CommandValidator {
    private static instance;
    constructor();
    /**
     * Головна функція валідації команди
     */
    validateCommand(interaction: ChatInputCommandInteraction, rules?: ValidationRules): Promise<ValidationResult>;
    /**
     * Валідація опцій команди
     */
    private validateOptions;
    /**
     * Санітизація строкових значень
     */
    private sanitizeStringValue;
    /**
     * Валідація числових значень
     */
    private validateNumberValue;
    /**
     * Валідація дозволів користувача
     */
    private validateUserPermissions;
    /**
     * Валідація контексту виконання
     */
    private validateExecutionContext;
    /**
     * Перевірка на підозрілий контент
     */
    private containsSuspiciousContent;
    /**
     * Валідація з кастомними правилами
     */
    validateWithRules(interaction: ChatInputCommandInteraction, rules: ValidationRules): Promise<ValidationResult>;
    /**
     * Швидка валідація без складних перевірок
     */
    quickValidate(interaction: ChatInputCommandInteraction): ValidationResult;
    /**
     * Отримання статистики валідації
     */
    getValidationStats(): {
        totalValidations: number;
        successfulValidations: number;
        failedValidations: number;
    };
}
export default CommandValidator;
//# sourceMappingURL=CommandValidator.d.ts.map