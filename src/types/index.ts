// Основні типи для Discord AI Assistant Bot

// Конфігурація
export interface BotConfig {
  discord: DiscordConfig;
  google: GoogleConfig;
  ai: AIConfig;
  redis: RedisConfig;
  metrics: MetricsConfig;
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

// Утиліти
export interface Logger {
  info(message: string, ...args: unknown[]): void;
  warn(message: string, ...args: unknown[]): void;
  error(message: string, ...args: unknown[]): void;
  debug(message: string, ...args: unknown[]): void;
}

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

// Всі типи експортовані безпосередньо з цього файлу 