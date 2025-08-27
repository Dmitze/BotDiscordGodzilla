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
import type { AdvancedDocumentAnalyzer } from '@/services/AdvancedDocumentAnalyzer';
import type { IntelligentWorkflowOrchestrator } from '@/services/IntelligentWorkflowOrchestrator';
import type { SmartSearchEngine } from '@/services/SmartSearchEngine';
import type { WorkflowAutomationEngine } from '@/services/WorkflowAutomationEngine';
import type { EnhancedDocumentService } from '@/services/EnhancedDocumentService';
import type { ContextMemoryService } from '@/services/ContextMemoryService';
import type { ResponseCacheService } from '@/services/ResponseCacheService';
import type { KnowledgeBaseService } from '@/services/KnowledgeBaseService';
import type { EnhancedRagService } from '@/services/EnhancedRagService';
import type { DocumentAnalysisService } from '@/services/DocumentAnalysisService';

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
  | 'rag'
  | 'documentAnalyzer'
  | 'workflowOrchestrator'
  | 'smartSearch'
  | 'workflowEngine'
  | 'enhancedDocumentService'
  | 'contextMemory'
  | 'responseCache'
  | 'knowledgeBase'
  | 'enhancedRag'
  | 'documentAnalysis';

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
  documentAnalyzer?: AdvancedDocumentAnalyzer; // optional if AI or Google missing
  workflowOrchestrator?: IntelligentWorkflowOrchestrator; // optional if dependencies missing
  smartSearch?: SmartSearchEngine; // optional if dependencies missing
  workflowEngine?: WorkflowAutomationEngine; // optional if dependencies missing
  enhancedDocumentService?: EnhancedDocumentService; // optional if dependencies missing
  contextMemory?: ContextMemoryService; // optional service for user context
  responseCache?: ResponseCacheService; // optional caching service
  knowledgeBase?: KnowledgeBaseService; // optional if dependencies missing
  enhancedRag?: EnhancedRagService; // enhanced RAG with auto-indexing
  documentAnalysis?: DocumentAnalysisService; // document analysis service
}

// Helper for DI resolve generics
export type ResolveService = <K extends ServiceKey>(name: K) => ServiceRegistry[K] | undefined;