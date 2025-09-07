// Existing exports
export * from './AIService';
export * from './GoogleService';
export * from './CacheService';
export * from './MetricsService';
export * from './SchedulerService';
export * from './EnhancedRagService';
export * from './SmartSearchEngine';
export * from './DriveIndexerService';
export * from './ContextMemoryService';
export * from './ResponseCacheService';
export * from './UserPreferencesService';
export * from './WorkspaceService';
export * from './AnalyticsService';
export * from './AIPromptTemplateService';
export * from './AdvancedDocumentAnalyzer';
export * from './KnowledgeBaseService';
export * from './EmbeddingsService';
export * from './RagService';
export * from './SheetsContextService';
export * from './UIStateService';
export * from './WorkflowAutomationEngine';
export * from './IntelligentWorkflowOrchestrator';
export * from './SmartDocumentClassifier';
export * from './DriveChangesService';
// Fix: Remove conflicting export
// export * from './MultilingualDocumentProcessor';
export * from './DocumentAnalyticsService';
export * from './DocumentMentionHandler';
export * from './AutomatedDocumentProcessor';
export * from './DocumentExportImportService';
export * from './DocumentAnalysisService';

// Add missing exports (avoiding duplicates and conflicts)
export { EnhancedCacheService } from './EnhancedCacheService';
export { HybridSearchService } from './HybridSearchService';
export { MultimodalRagService } from './MultimodalRagService';
export { DocumentAccessAuditService } from './DocumentAccessAuditService';
export { DocumentEncryptionService } from './DocumentEncryptionService';
export { DocumentSummarizationService } from './DocumentSummarizationService';
export { DocumentVersionComparisonService } from './DocumentVersionComparisonService';
export { GoogleApiRateLimitService } from './GoogleApiRateLimitService';
export { LoadBalancingService } from './LoadBalancingService';
export { MemoryOptimizationService } from './MemoryOptimizationService';
export { SlackIntegrationService } from './SlackIntegrationService';