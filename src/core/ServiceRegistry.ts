import type { AIService } from '@/services/AIService';
import type { GoogleService } from '@/services/GoogleService';
import type { CacheService } from '@/services/CacheService';
import type { MemoryCacheService } from '@/services/MemoryCacheService';
import type { MetricsService } from '@/services/MetricsService';
import type SchedulerService from '@/services/SchedulerService';
import type { SheetsContextService } from '@/services/SheetsContextService';
import type { DriveIndexerService } from '@/services/DriveIndexerService';
import type { DriveChangesService } from '@/services/DriveChangesService';
import type { WorkspaceDbService } from '@/services/WorkspaceDbService';
import type { RagService } from '@/services/RagService';
import type { EmbeddingsService } from '@/services/EmbeddingsService';
import type { SearchIndex } from '@/search/SearchIndex';

// Union of all service keys managed by ServiceManager
export type ServiceKey =
  | 'ai'
  | 'google'
  | 'cache'
  | 'metrics'
  | 'scheduler'
  | 'sheetsContext'
  | 'driveChanges'
  | 'driveIndexer'
  | 'workspace'
  | 'searchIndex'
  | 'embeddings'
  | 'rag';

// Registry types mapping keys to concrete instances
export interface ServiceRegistry {
  ai: AIService;
  google: GoogleService;
  cache: CacheService | MemoryCacheService;
  metrics?: MetricsService; // optional based on config
  scheduler: SchedulerService;
  sheetsContext: SheetsContextService;
  driveChanges: DriveChangesService;
  driveIndexer: DriveIndexerService;
  workspace?: WorkspaceDbService; // optional if initialization failed
  searchIndex: SearchIndex;
  embeddings?: EmbeddingsService; // optional
  rag?: RagService; // optional if AI missing
}

// Helper for DI resolve generics
export type ResolveService = <K extends ServiceKey>(name: K) => ServiceRegistry[K] | undefined;
