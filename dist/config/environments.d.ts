/**
 * 🌍 Конфігурація середовищ розгортання
 * Discord AI Assistant Bot v2.3.0
 * TypeScript версія
 */
interface LoggingConfig {
    level: string;
    maxFiles: number;
    maxSize: string;
    directory: string;
}
interface MetricsConfig {
    enabled: boolean;
    port: number;
    path: string;
}
interface SecurityConfig {
    rateLimitWindow: number;
    rateLimitMax: number;
    adminRole: string;
    botUserRole: string;
}
interface PerformanceConfig {
    cacheTTL: number;
    maxSearchResults: number;
    maxAnalysisRows: number;
    requestTimeout: number;
    maxRetries: number;
}
interface DiscordConfig {
    token: string;
    clientId: string;
    guildId: string;
    intents: string[];
}
interface GoogleConfig {
    apiKey: string;
    appScriptUrl: string;
    sheetName: string;
}
interface OpenAIConfig {
    apiKey: string;
    model: string;
    maxTokens: number;
    temperature: number;
}
interface OllamaConfig {
    enabled: boolean;
    url: string;
    model: string;
}
interface AIConfig {
    openai: OpenAIConfig;
    ollama: OllamaConfig;
}
interface RedisConfig {
    enabled: boolean;
    host: string;
    port: number;
    password: string | null;
    db: number;
}
interface EnvironmentSpecificConfig {
    debug: boolean;
    verbose: boolean;
    hotReload: boolean;
    testMode: boolean;
    monitoring?: boolean;
    clustering?: boolean;
    loadBalancing?: boolean;
    mockExternalServices?: boolean;
}
interface BaseConfig {
    logging: LoggingConfig;
    metrics: MetricsConfig;
    security: SecurityConfig;
    performance: PerformanceConfig;
}
interface EnvironmentConfig extends BaseConfig {
    name: string;
    nodeEnv: string;
    discord: DiscordConfig;
    google: GoogleConfig;
    ai: AIConfig;
    redis: RedisConfig;
    development?: EnvironmentSpecificConfig;
    testing?: EnvironmentSpecificConfig;
    staging?: EnvironmentSpecificConfig;
    production?: EnvironmentSpecificConfig;
}
declare const development: EnvironmentConfig;
declare const testing: EnvironmentConfig;
declare const staging: EnvironmentConfig;
declare const production: EnvironmentConfig;
declare function getConfig(environment?: string | null): EnvironmentConfig;
declare function validateConfig(config: EnvironmentConfig): boolean;
declare function getValidatedConfig(environment?: string | null): EnvironmentConfig;
export { development, testing, staging, production, getConfig, validateConfig, getValidatedConfig, type EnvironmentConfig, type BaseConfig, type DiscordConfig, type GoogleConfig, type AIConfig, type RedisConfig, };
//# sourceMappingURL=environments.d.ts.map