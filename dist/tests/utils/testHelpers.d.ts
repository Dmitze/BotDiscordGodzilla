/**
 * Утилиты для тестирования
 */
import { jest } from '@jest/globals';
/**
 * Создание мок конфигурации для тестов
 */
export declare function createMockConfig(): any;
/**
 * Создание мок Discord взаимодействия
 */
export declare function createMockInteraction(): {
    commandName: string;
    user: {
        id: string;
        username: string;
        tag: string;
    };
    guild: {
        id: string;
        name: string;
    };
    channel: {
        id: string;
        name: string;
    };
    options: {
        getString: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
        getInteger: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
        getBoolean: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
        getSubcommand: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
    };
    reply: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
    editReply: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
    followUp: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
    deferReply: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
    replied: boolean;
    deferred: boolean;
    isCommand: () => boolean;
    client: {
        serviceContainer: {
            get: import("jest-mock").Mock<import("jest-mock").UnknownFunction>;
        };
    };
};
/**
 * Создание мок Google Sheets данных
 */
export declare function createMockSheetData(): string[][];
/**
 * Ожидание асинхронной операции
 */
export declare function wait(ms: number): Promise<void>;
/**
 * Очистка моков
 */
export declare function clearMocks(): void;
/**
 * Проверка вызова функции
 */
export declare function expectFunctionCalled(fn: jest.Mock, times?: number): void;
/**
 * Проверка вызова функции с параметрами
 */
export declare function expectFunctionCalledWith(fn: jest.Mock, ...args: any[]): void;
//# sourceMappingURL=testHelpers.d.ts.map