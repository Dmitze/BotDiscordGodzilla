import type { AIService } from '@/services/AIService';
import type { GoogleService } from '@/services/GoogleService';
import type { CacheService } from '@/services/CacheService';
import type { MetricsService } from '@/services/MetricsService';
import type SchedulerService from '@/services/SchedulerService';
import type { SheetsContextService } from '@/services/SheetsContextService';
import type { DriveChangesService } from '@/services/DriveChangesService';
import type { WorkspaceDbService } from '@/services/WorkspaceDbService';
import type { SearchIndex } from '@/search/SearchIndex';
import type { RagService } from '@/services/RagService';
import type { DriveIndexerService } from '@/services/DriveIndexerService';
import type { EmbeddingsService } from '@/services/EmbeddingsService';
import type { ResponseCacheService } from '@/services/ResponseCacheService';
import type { DocumentAnalysisService } from '@/services/DocumentAnalysisService';
import type { AdvancedDocumentAnalyzer } from '@/services/AdvancedDocumentAnalyzer';
import type { IntelligentWorkflowOrchestrator } from '@/services/IntelligentWorkflowOrchestrator';
import type { SmartSearchEngine } from '@/services/SmartSearchEngine';
import type { EnhancedDocumentService } from '@/services/EnhancedDocumentService';
import type { WorkflowAutomationEngine } from '@/services/WorkflowAutomationEngine';
import type { ContextMemoryService } from '@/services/ContextMemoryService';
import type { KnowledgeBaseService } from '@/services/KnowledgeBaseService';
import type { EnhancedRagService } from '@/services/EnhancedRagService';
import type { MultimodalRagService } from '@/services/MultimodalRagService';
import type { HybridSearchService } from '@/services/HybridSearchService';

/**
 * Union type of all valid service keys
 */
export type ServiceKey =
  | 'ai'
  | 'google'
  | 'cache'
  | 'metrics'
  | 'scheduler'
  | 'sheetsContext'
  | 'driveChanges'
  | 'workspace'
  | 'searchIndex'
  | 'rag'
  | 'driveIndexer'
  | 'embeddings'
  | 'responseCache'
  | 'documentAnalysis'
  | 'documentAnalyzer'
  | 'workflowOrchestrator'
  | 'smartSearch'
  | 'enhancedDocumentService'
  | 'workflowEngine'
  | 'contextMemory'
  | 'knowledgeBase'
  | 'enhancedRag'
  | 'multimodalRag'
  | 'hybridSearch';

/**
 * Registry mapping service keys to their respective types
 */
export interface ServiceRegistry {
  ai: AIService;
  google: GoogleService;
  cache: CacheService;
  metrics: MetricsService;
  scheduler: SchedulerService;
  sheetsContext: SheetsContextService;
  driveChanges: DriveChangesService;
  workspace: WorkspaceDbService;
  searchIndex: SearchIndex;
  rag: RagService;
  driveIndexer: DriveIndexerService;
  embeddings: EmbeddingsService;
  responseCache: ResponseCacheService;
  documentAnalysis: DocumentAnalysisService;
  documentAnalyzer: AdvancedDocumentAnalyzer;
  workflowOrchestrator: IntelligentWorkflowOrchestrator;
  smartSearch: SmartSearchEngine;
  enhancedDocumentService: EnhancedDocumentService;
  workflowEngine: WorkflowAutomationEngine;
  contextMemory: ContextMemoryService;
  knowledgeBase: KnowledgeBaseService;
  enhancedRag: EnhancedRagService;
  multimodalRag: MultimodalRagService;
  hybridSearch: HybridSearchService;
}