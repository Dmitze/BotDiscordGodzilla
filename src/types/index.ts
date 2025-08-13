// Основні типи для Discord AI Assistant Bot
// Версія 3.0.0 - Повністю рефакторовано з детальним логуванням

// Конфігурація
export interface BotConfig {
  discord: DiscordConfig;
  google: GoogleConfig;
  ai: AIConfig;
  redis: RedisConfig;
  metrics: MetricsConfig;
  security: SecurityConfig;
  performance: PerformanceConfig;
  logging: LoggingConfig;
}

export interface DiscordConfig {
  token: string;
  clientId: string;
  guildId: string;
  prefix: string;
  intents: string[];
}

export interface GoogleConfig {
  spreadsheetId: string;
  driveFolderId: string;
  apiKey: string;
  applicationCredentials: string;
  appScriptUrl: string;
  sheetName: string;
  credentials?: GoogleCredentials;
  // OCR settings
  ocrProvider?: 'vision' | 'tesseract' | 'off';
  ocrCacheTTL?: number;
  tesseractLangs?: string; // e.g. "eng+ukr+rus"
  tesseractLangPath?: string; // local path to traineddata directory
  // Analytics cache
  analyticsCacheTTL?: number;
}

export interface GoogleCredentials {
  client_email: string;
  private_key: string;
  project_id: string;
}

export interface AIConfig {
  provider: 'openai' | 'ollama';
  openai: OpenAIConfig;
  ollama: OllamaConfig;
}

export interface OpenAIConfig {
  apiKey: string;
  model: string;
  maxTokens: number;
  temperature: number;
}

export interface OllamaConfig {
  host: string;
  model: string;
}

export interface RedisConfig {
  host: string;
  port: number;
  password?: string;
  database: number;
  enabled: boolean;
  url?: string;
}

export interface MetricsConfig {
  enabled: boolean;
  port: number;
  path: string;
}

export interface SecurityConfig {
  rateLimitWindow: number;
  rateLimitMax: number;
  adminRole: string;
  botUserRole: string;
}

export interface PerformanceConfig {
  cacheTTL: number;
  maxSearchResults: number;
  maxAnalysisRows: number;
  requestTimeout: number;
  maxRetries: number;
}

export interface LoggingConfig {
  level: string;
  maxFiles: number;
  maxSize: string;
  directory: string;
}

// Сервіси
export interface BaseService {
  name: string;
  config: BotConfig;
  initialize(): Promise<void>;
  shutdown(): Promise<void>;
  healthCheck(): Promise<HealthStatus>;
  getStats(): ServiceStats;
}

export interface HealthStatus {
  healthy: boolean;
  service: string;
  error?: string;
  details?: Record<string, unknown>;
}

export interface ServiceStats {
  service: string;
  uptime: number;
  requests: number;
  errors: number;
  isInitialized?: boolean;
  isShuttingDown?: boolean;
  retryCount?: number;
  [key: string]: unknown;
}

// Google API типи
export interface SheetData {
  range: string;
  majorDimension: string;
  values: string[][];
}

export interface BatchSheetData {
  valueRanges: SheetData[];
  spreadsheetId: string;
}

export interface GoogleApiResponse<T = unknown> {
  data: T;
  status: number;
  statusText: string;
}

// AI типи
export interface AIResponse {
  content: string;
  provider: string;
  model: string;
  tokens: number;
  duration: number;
}

export interface AIRequest {
  prompt: string;
  options?: AIRequestOptions;
}

export interface AIRequestOptions {
  provider?: string;
  model?: string;
  maxTokens?: number;
  temperature?: number;
  useCache?: boolean;
  cacheTTL?: number;
  retryAttempts?: number;
  forceRefresh?: boolean;
  timeout?: number;
}

// Кеш типи
export interface CacheStats {
  hits: number;
  misses: number;
  sets: number;
  deletes: number;
  errors: number;
}

export interface CacheOptions {
  ttl?: number;
  compress?: boolean;
}

// Черги
export interface QueueJob {
  id: string;
  priority: 'high' | 'normal' | 'low';
  type: string;
  data: unknown;
  timestamp: number;
  retries: number;
  maxRetries: number;
}

export interface QueueStats {
  processed: number;
  failed: number;
  pending: number;
  averageProcessingTime: number;
  high: QueuePriorityStats;
  normal: QueuePriorityStats;
  low: QueuePriorityStats;
}

export interface QueuePriorityStats {
  length: number;
  processing: number;
  maxConcurrent: number;
}

// Discord типи
export interface CommandInteraction {
  commandName: string;
  options: Map<string, unknown>;
  user: DiscordUser;
  channel: DiscordChannel;
  guild?: DiscordGuild;
  reply(content: string | DiscordEmbed): Promise<void>;
  deferReply(): Promise<void>;
  editReply(content: string | DiscordEmbed): Promise<void>;
}

export interface DiscordUser {
  id: string;
  username: string;
  discriminator: string;
  avatar?: string;
}

export interface DiscordChannel {
  id: string;
  name: string;
  type: string;
}

export interface DiscordGuild {
  id: string;
  name: string;
  memberCount: number;
}

export interface DiscordEmbed {
  title?: string;
  description?: string;
  color?: number;
  fields?: DiscordEmbedField[];
  footer?: DiscordEmbedFooter;
  timestamp?: Date;
}

export interface DiscordEmbedField {
  name: string;
  value: string;
  inline?: boolean;
}

export interface DiscordEmbedFooter {
  text: string;
  iconURL?: string;
}

// Команди
export interface BaseCommand {
  data: SlashCommandBuilder;
  execute(interaction: CommandInteraction): Promise<void>;
}

export interface SlashCommandBuilder {
  setName(name: string): SlashCommandBuilder;
  setDescription(description: string): SlashCommandBuilder;
  addSubcommand(subcommand: SubcommandBuilder): SlashCommandBuilder;
  addStringOption(option: StringOptionBuilder): SlashCommandBuilder;
  addIntegerOption(option: IntegerOptionBuilder): SlashCommandBuilder;
  addBooleanOption(option: BooleanOptionBuilder): SlashCommandBuilder;
  toJSON(): unknown;
}

export interface SubcommandBuilder {
  setName(name: string): SubcommandBuilder;
  setDescription(description: string): SubcommandBuilder;
  addStringOption(option: StringOptionBuilder): SubcommandBuilder;
  addIntegerOption(option: IntegerOptionBuilder): SubcommandBuilder;
  addBooleanOption(option: BooleanOptionBuilder): SubcommandBuilder;
}

export interface StringOptionBuilder {
  setName(name: string): StringOptionBuilder;
  setDescription(description: string): StringOptionBuilder;
  setRequired(required: boolean): StringOptionBuilder;
  addChoices(...choices: Array<{ name: string; value: string }>): StringOptionBuilder;
}

export interface IntegerOptionBuilder {
  setName(name: string): IntegerOptionBuilder;
  setDescription(description: string): IntegerOptionBuilder;
  setRequired(required: boolean): IntegerOptionBuilder;
  setMinValue(min: number): IntegerOptionBuilder;
  setMaxValue(max: number): IntegerOptionBuilder;
}

export interface BooleanOptionBuilder {
  setName(name: string): BooleanOptionBuilder;
  setDescription(description: string): BooleanOptionBuilder;
  setRequired(required: boolean): BooleanOptionBuilder;
}

// Кластеризація
export interface ClusterConfig {
  workers: number;
  restartDelay: number;
  maxRestarts: number;
}

export interface WorkerInfo {
  id: number;
  pid: number;
  status: 'starting' | 'online' | 'disconnected' | 'exited';
  startTime: number;
  restarts: number;
  stats?: WorkerStats;
  lastUpdate?: number;
}

export interface WorkerStats {
  load: number;
  memory: number;
  uptime: number;
}

export interface ClusterStats {
  totalWorkers: number;
  activeWorkers: number;
  restarts: number;
  startTime: number;
  uptime: number;
  isActive: boolean;
  workers: WorkerInfo[];
  restartCounts: Record<string, number>;
}

// Логування типи
export interface LogMeta {
  [key: string]: any;
  timestamp?: string;
  level?: string;
  service?: string;
  userId?: string;
  guildId?: string;
  channelId?: string;
  requestId?: string;
  correlationId?: string;
  type?: 'command' | 'api_request' | 'performance' | 'security' | 'system';
  duration?: string;
  performance?: 'fast' | 'medium' | 'slow';
  severity?: 'low' | 'medium' | 'high' | 'critical';
  component?: string;
  category?: string;
}

export interface LoggerStats {
  totalLogs: number;
  errors: number;
  commands: number;
  apiRequests: number;
  performance: number;
  security: number;
  system: number;
  debug: number;
  warnings: number;
  lastLogTime: Date;
  averageLogSize: number;
  logBufferSize: number;
}

export interface LogEntry {
  timestamp: Date;
  level: string;
  message: string;
  meta: LogMeta;
  size: number;
}

export interface Logger {
  info(message: string, meta?: LogMeta): void;
  warn(message: string, meta?: LogMeta): void;
  error(message: string, meta?: LogMeta): void;
  debug(message: string, meta?: LogMeta): void;
  command(command: string, user: string, duration: number, success?: boolean, meta?: LogMeta): void;
  commandError(command: string, user: string, error: Error, duration: number, meta?: LogMeta): void;
  apiRequest(
    service: string,
    endpoint: string,
    duration: number,
    success?: boolean,
    meta?: LogMeta
  ): void;
  apiError(service: string, endpoint: string, error: Error, duration: number, meta?: LogMeta): void;
  security(event: string, user: string, details?: LogMeta): void;
  performance(operation: string, duration: number, details?: LogMeta): void;
  system(event: string, details?: LogMeta): void;
  getStats(): LoggerStats;
  getLogBuffer(): LogEntry[];
  cleanup(): Promise<void>;
  isHealthy(): boolean;
}

// Безпека типи
export interface SecurityValidationResult {
  isValid: boolean;
  sanitizedValue: string;
  errors: string[];
  warnings: string[];
}

export interface RateLimitInfo {
  count: number;
  resetTime: number;
}

export interface SecurityEvent {
  type: 'rate_limit' | 'invalid_input' | 'unauthorized_access' | 'suspicious_activity';
  userId: string;
  details: Record<string, unknown>;
  timestamp: Date;
  severity: 'low' | 'medium' | 'high' | 'critical';
}

// Утиліти
export interface PaginationOptions {
  itemsPerPage?: number;
  maxPages?: number;
  embedColor?: number;
  title?: string;
  description?: string;
  fields?: string[];
  footer?: string;
  timestamp?: Date;
}

// Command types
export interface CommandOptions {
  useCache?: boolean;
  cacheTTL?: number;
  timeout?: number;
  retryAttempts?: number;
  priority?: 'high' | 'normal' | 'low';
}

export interface CommandStats extends ServiceStats {
  totalExecutions: number;
  successfulExecutions: number;
  failedExecutions: number;
  averageExecutionTime: number;
  totalExecutionTime: number;
  cacheHits: number;
  cacheMisses: number;
  retries: number;
}

export interface CommandContext {
  userId: string;
  guildId?: string;
  channelId: string;
  timestamp: number;
  metadata?: Record<string, unknown>;
}

export interface CommandExecuteOptions {
  interaction: any;
  context?: CommandContext;
  options?: CommandOptions;
  startTime?: number;
  retryCount?: number;
}

export interface CommandAutocompleteOptions {
  interaction: any;
  context?: CommandContext;
  query?: string;
}

export interface CommandComponentOptions {
  interaction: any;
  context?: CommandContext;
  componentType?: 'button' | 'select' | 'modal';
}

export interface CommandValidationResult {
  isValid: boolean;
  errors: string[];
  warnings: string[];
  sanitizedOptions?: any;
}

export interface SearchParams {
  query: string;
  documentType: string;
  dateFrom?: string;
  dateTo?: string;
  unit?: string;
  priority: string;
  limit: number;
}

// Моніторинг та метрики
export interface MonitoringConfig {
  healthCheckInterval: number;
  memoryThreshold: number;
  cpuThreshold: number;
  maxRestarts: number;
  restartDelay: number;
}

export interface SystemMetrics {
  memory: {
    rss: number;
    heapUsed: number;
    heapTotal: number;
    external: number;
  };
  cpu: {
    usage: number;
    load: number;
  };
  uptime: number;
  processId: number;
}

export interface PerformanceMetrics {
  operation: string;
  duration: number;
  category: string;
  metadata?: Record<string, unknown>;
}

// Всі типи експортовані безпосередньо з цього файлу
