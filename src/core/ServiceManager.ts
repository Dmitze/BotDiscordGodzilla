/**
 * Менеджер сервісів Discord бота
 * Централізоване управління всіма сервісами
 * TypeScript версія
 */

import logger from '@/utils/logger';
import type { ServiceKey, ServiceRegistry } from '@/core/ServiceRegistry';

import { AIService } from '../services/AIService';
import { GoogleService } from '../services/GoogleService';
import { GoogleSheetsService } from '../services/GoogleSheetsService';
import { CacheService } from '../services/CacheService';
import { MemoryCacheService } from '@/services/MemoryCacheService';
import { MetricsService } from '../services/MetricsService';
import { SheetsContextService } from '../services/SheetsContextService';
import SchedulerService from '../services/SchedulerService';
import { DriveIndexerService } from '../services/DriveIndexerService';
import { SqliteSearchIndex } from '@/search/sqlite/SqliteSearchIndex';
import type { SearchIndex } from '@/search/SearchIndex';
import { DriveChangesService } from '../services/DriveChangesService';
import type { BotConfig } from '@/types';
import { WorkspaceDbService } from '@/services/WorkspaceDbService';
import { RagService } from '@/services/RagService';
import { EmbeddingsService } from '@/services/EmbeddingsService';
import { AdvancedDocumentAnalyzer } from '@/services/AdvancedDocumentAnalyzer';
import { IntelligentWorkflowOrchestrator } from '@/services/IntelligentWorkflowOrchestrator';
import { SmartSearchEngine } from '@/services/SmartSearchEngine';
import { WorkflowAutomationEngine } from '@/services/WorkflowAutomationEngine';
import { EnhancedDocumentService } from '@/services/EnhancedDocumentService';
import { ContextMemoryService } from '@/services/ContextMemoryService';
import { ResponseCacheService } from '@/services/ResponseCacheService';
import { KnowledgeBaseService } from '@/services/KnowledgeBaseService';
import { EnhancedRagService } from '@/services/EnhancedRagService';

interface Bot {
  config: BotConfig;
  getService: (name: string) => any;
  serviceManager?: any;
  client?: any;
}

// Legacy dynamic service shape kept implicitly via `any` where needed

interface ServiceStatus {
  isActive: boolean;
  hasMethod: (method: string) => boolean;
  stats: any;
}

interface ServiceManagerStats {
  total: number;
  active: number;
  services: string[];
  status: Record<string, ServiceStatus>;
}

class ServiceManager {
  private bot: Bot;
  private services: Map<ServiceKey, NonNullable<ServiceRegistry[ServiceKey]>>;
  //

  constructor(bot: Bot) {
    this.bot = bot;
    this.services = new Map();
    //
  }

  /**
   * Ініціалізація менеджера сервісів
   */
  async initialize(): Promise<void> {
    try {
      logger.info('🔧 Ініціалізація менеджера сервісів...', {
        type: 'service_manager',
        event: 'init',
        component: 'ServiceManager',
      });

      // Створення сервісів
      await this.createServices();

      // Ініціалізація сервісів
      await this.initializeServices();

      logger.info('✅ Менеджер сервісів ініціалізовано', {
        type: 'service_manager',
        event: 'init_success',
        component: 'ServiceManager',
      });
    } catch (error) {
      logger.error('❌ Помилка ініціалізації менеджера сервісів', {
        type: 'service_manager',
        event: 'init_failed',
        component: 'ServiceManager',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });

      throw error;
    }
  }

  /**
   * Створення сервісів
   */
  private async createServices(): Promise<void> {
    // AI Service
    this.services.set('ai', new AIService(this.bot.config));

    // Google Service - using the new GoogleService
    this.services.set('google', new GoogleService(this.bot.config));

    // Cache Service: завжди доступний у контейнері
    // - Якщо Redis увімкнено — використовуємо Redis CacheService
    // - Інакше або у тестах — легкий MemoryCacheService
    const useRedis = Boolean((this.bot.config as any).redis?.enabled);
    if (useRedis) {
      this.services.set('cache', new CacheService(this.bot.config));
    } else {
      this.services.set('cache', new MemoryCacheService(this.bot.config));
    }

    // Metrics Service (якщо метрики увімкнені)
    const metricsEnabled = Boolean((this.bot.config as any).metrics?.enabled);
    if (metricsEnabled) {
      this.services.set('metrics', new MetricsService(this.bot.config));
    }

    // Scheduler Service
    this.services.set('scheduler', new SchedulerService(this.bot));

    // Sheets Context Service
    this.services.set('sheetsContext', new SheetsContextService(this.bot.config));

    // Drive Changes Service (polling changes)
    this.services.set('driveChanges', new DriveChangesService(this.bot.config));

    // Drive Indexer Service (потребує доступу до інших сервісів через getService)
    this.services.set(
      'driveIndexer',
      new DriveIndexerService({
        config: this.bot.config,
        getService: (name: string) => this.getService(name as ServiceKey),
      } as any)
    );

    // Workspace (персональний простір користувача) на SQLite
    try {
      const workspace = new WorkspaceDbService(this.bot.config);
      this.services.set('workspace', workspace as unknown as NonNullable<ServiceRegistry['workspace']>);
      logger.info('🗂️ WorkspaceDbService зареєстровано', {
        type: 'service_manager',
        event: 'workspace_registered',
        component: 'ServiceManager',
      });
    } catch (e) {
      logger.error('❌ Не вдалося створити WorkspaceDbService', {
        type: 'service_manager',
        event: 'workspace_register_failed',
        component: 'ServiceManager',
        errorMessage: e instanceof Error ? e.message : String(e),
      });
    }

    // Persistent Search Index (SQLite FTS5)
    try {
      const dbPath =
        process.env['SEARCH_INDEX_PATH'] ||
        process.env['BOT_INDEX_DB_PATH'] ||
        './data/search-index.db';
      const searchIndex: SearchIndex = new SqliteSearchIndex({ dbPath });
      this.services.set('searchIndex', searchIndex as unknown as ServiceRegistry['searchIndex']);
      logger.info('🗂️ SqliteSearchIndex зареєстровано', {
        type: 'service_manager',
        event: 'search_index_registered',
        component: 'ServiceManager',
        dbPath,
      });

      // Embeddings Service (optional, can fallback to mock)
      try {
        const embeddings = new EmbeddingsService(this.bot.config);
        this.services.set('embeddings', embeddings as unknown as NonNullable<ServiceRegistry['embeddings']>);
        logger.info('🧮 EmbeddingsService зареєстровано', {
          type: 'service_manager',
          event: 'embeddings_registered',
          component: 'ServiceManager',
        });
      } catch (er) {
        logger.error('❌ Не вдалося створити EmbeddingsService', {
          type: 'service_manager',
          event: 'embeddings_register_failed',
          component: 'ServiceManager',
          errorMessage: er instanceof Error ? er.message : String(er),
        });
      }

      // RAG Service (depends on AI + SearchIndex + (optional) Embeddings)
      try {
        const aiSvc = this.services.get('ai');
        const emb = this.services.get('embeddings');
        const cacheSvc = this.services.get('cache') as
          | { get<T = unknown>(key: string): Promise<T | null>; set<T = unknown>(key: string, value: T, ttlSec?: number): Promise<unknown> }
          | undefined;
        if (aiSvc) {
          const rag = new RagService(
            searchIndex as any,
            aiSvc as any,
            (emb as unknown as { embed: (t: string) => Promise<number[]> } | undefined),
            cacheSvc ? { cache: cacheSvc } : undefined
          );
          this.services.set('rag', rag as unknown as NonNullable<ServiceRegistry['rag']>);
          logger.info('🧩 RagService зареєстровано', {
            type: 'service_manager',
            event: 'rag_registered',
            component: 'ServiceManager',
          });
        } else {
          logger.warn('RagService не зареєстровано: AI service недоступний');
        }
      } catch (er) {
        logger.error('❌ Не вдалося створити RagService', {
          type: 'service_manager',
          event: 'rag_register_failed',
          component: 'ServiceManager',
          errorMessage: er instanceof Error ? er.message : String(er),
        });
      }
    } catch (e) {
      logger.error('❌ Не вдалося створити SqliteSearchIndex', {
        type: 'service_manager',
        event: 'search_index_register_failed',
        component: 'ServiceManager',
        errorMessage: e instanceof Error ? e.message : String(e),
      });
    }

    // Advanced Document Analyzer (depends on AI + Google services)
    try {
      const aiSvc = this.services.get('ai');
      const googleSvc = this.services.get('google');
      if (aiSvc && googleSvc) {
        const documentAnalyzer = new AdvancedDocumentAnalyzer(aiSvc as any, googleSvc as any);
        this.services.set('documentAnalyzer', documentAnalyzer as unknown as NonNullable<ServiceRegistry['documentAnalyzer']>);
        logger.info('🧠 AdvancedDocumentAnalyzer зареєстровано', {
          type: 'service_manager',
          event: 'document_analyzer_registered',
          component: 'ServiceManager',
        });
      } else {
        logger.warn('AdvancedDocumentAnalyzer не зареєстровано: AI або Google service недоступний');
      }
    } catch (er) {
      logger.error('❌ Не вдалося створити AdvancedDocumentAnalyzer', {
        type: 'service_manager',
        event: 'document_analyzer_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Intelligent Workflow Orchestrator (depends on AI + DocumentAnalyzer)
    try {
      const aiSvc = this.services.get('ai');
      const documentAnalyzer = this.services.get('documentAnalyzer');
      if (aiSvc && documentAnalyzer) {
        const workflowOrchestrator = new IntelligentWorkflowOrchestrator(aiSvc as any, documentAnalyzer as any);
        this.services.set('workflowOrchestrator', workflowOrchestrator as unknown as NonNullable<ServiceRegistry['workflowOrchestrator']>);
        logger.info('🔄 IntelligentWorkflowOrchestrator зареєстровано', {
          type: 'service_manager',
          event: 'workflow_orchestrator_registered',
          component: 'ServiceManager',
        });
      } else {
        logger.warn('IntelligentWorkflowOrchestrator не зареєстровано: AI або DocumentAnalyzer service недоступний');
      }
    } catch (er) {
      logger.error('❌ Не вдалося створити IntelligentWorkflowOrchestrator', {
        type: 'service_manager',
        event: 'workflow_orchestrator_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Smart Search Engine (depends on AI + Google + RAG services)
    try {
      const aiSvc = this.services.get('ai');
      const googleSvc = this.services.get('google');
      const ragSvc = this.services.get('rag');
      if (aiSvc && googleSvc && ragSvc) {
        const smartSearch = new SmartSearchEngine(aiSvc as any, googleSvc as any, ragSvc as any);
        this.services.set('smartSearch', smartSearch as unknown as NonNullable<ServiceRegistry['smartSearch']>);
        logger.info('🔍 SmartSearchEngine зареєстровано', {
          type: 'service_manager',
          event: 'smart_search_registered',
          component: 'ServiceManager',
        });
      } else {
        logger.warn('SmartSearchEngine не зареєстровано: AI, Google або RAG service недоступний');
      }
    } catch (er) {
      logger.error('❌ Не вдалося створити SmartSearchEngine', {
        type: 'service_manager',
        event: 'smart_search_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Enhanced Document Service (depends on AI + Google services)
    try {
      const aiSvc = this.services.get('ai');
      const googleSvc = this.services.get('google');
      if (aiSvc && googleSvc) {
        const enhancedDocumentService = new EnhancedDocumentService(googleSvc as any, aiSvc as any);
        this.services.set('enhancedDocumentService', enhancedDocumentService as unknown as NonNullable<ServiceRegistry['enhancedDocumentService']>);
        logger.info('📋 EnhancedDocumentService зареєстровано', {
          type: 'service_manager',
          event: 'enhanced_document_service_registered',
          component: 'ServiceManager',
        });
      } else {
        logger.warn('EnhancedDocumentService не зареєстровано: AI або Google service недоступний');
      }
    } catch (er) {
      logger.error('❌ Не вдалося створити EnhancedDocumentService', {
        type: 'service_manager',
        event: 'enhanced_document_service_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Workflow Automation Engine (depends on AI + Google + EnhancedDocument services)
    try {
      const aiSvc = this.services.get('ai');
      const googleSvc = this.services.get('google');
      const enhancedDocumentService = this.services.get('enhancedDocumentService');
      if (aiSvc && googleSvc && enhancedDocumentService) {
        const workflowEngine = new WorkflowAutomationEngine(aiSvc as any, googleSvc as any, enhancedDocumentService as any);
        this.services.set('workflowEngine', workflowEngine as unknown as NonNullable<ServiceRegistry['workflowEngine']>);
        logger.info('⚙️ WorkflowAutomationEngine зареєстровано', {
          type: 'service_manager',
          event: 'workflow_engine_registered',
          component: 'ServiceManager',
        });
      } else {
        logger.warn('WorkflowAutomationEngine не зареєстровано: AI, Google або EnhancedDocument service недоступний');
      }
    } catch (er) {
      logger.error('❌ Не вдалося створити WorkflowAutomationEngine', {
        type: 'service_manager',
        event: 'workflow_engine_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Context Memory Service (standalone service)
    try {
      const contextMemory = new ContextMemoryService();
      this.services.set('contextMemory', contextMemory as unknown as NonNullable<ServiceRegistry['contextMemory']>);
      logger.info('🧠 ContextMemoryService зареєстровано', {
        type: 'service_manager',
        event: 'context_memory_registered',
        component: 'ServiceManager',
      });
    } catch (er) {
      logger.error('❌ Не вдалося створити ContextMemoryService', {
        type: 'service_manager',
        event: 'context_memory_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Response Cache Service (standalone service)
    try {
      const responseCache = new ResponseCacheService(30, 1000); // 30 min TTL, 1000 max entries
      this.services.set('responseCache', responseCache as unknown as NonNullable<ServiceRegistry['responseCache']>);
      logger.info('💾 ResponseCacheService зареєстровано', {
        type: 'service_manager',
        event: 'response_cache_registered',
        component: 'ServiceManager',
      });
    } catch (er) {
      logger.error('❌ Не вдалося створити ResponseCacheService', {
        type: 'service_manager',
        event: 'response_cache_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Knowledge Base Service (depends on Google + AI + RAG + ResponseCache services)
    try {
      const googleSvc = this.services.get('google');
      const aiSvc = this.services.get('ai');
      const ragSvc = this.services.get('rag');
      const responseCacheSvc = this.services.get('responseCache');
      if (googleSvc && aiSvc && ragSvc && responseCacheSvc) {
        const knowledgeBase = new KnowledgeBaseService(
          aiSvc as any,
          ragSvc as any,
          responseCacheSvc as any
        );
        this.services.set('knowledgeBase', knowledgeBase as unknown as NonNullable<ServiceRegistry['knowledgeBase']>);
        logger.info('📚 KnowledgeBaseService зареєстровано', {
          type: 'service_manager',
          event: 'knowledge_base_registered',
          component: 'ServiceManager',
        });
      } else {
        logger.warn('KnowledgeBaseService не зареєстровано: Google, AI, RAG або ResponseCache service недоступний');
      }
    } catch (er) {
      logger.error('❌ Не вдалося створити KnowledgeBaseService', {
        type: 'service_manager',
        event: 'knowledge_base_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Enhanced RAG Service (replaces standard RAG with auto-indexing)
    try {
      const searchIndexSvc = this.services.get('searchIndex');
      const aiSvc = this.services.get('ai');
      const googleSvc = this.services.get('google');
      const responseCacheSvc = this.services.get('responseCache');
      const embSvc = this.services.get('embeddings');
      
      if (searchIndexSvc && aiSvc && googleSvc && responseCacheSvc) {
        const enhancedRag = new EnhancedRagService(
          searchIndexSvc as any,
          aiSvc as any,
          googleSvc as any,
          responseCacheSvc as any,
          embSvc as unknown as { embed: (t: string) => Promise<number[]> } | undefined,
          { enabled: true, interval: '0 */2 * * *' } // Every 2 hours auto-indexing
        );
        this.services.set('enhancedRag', enhancedRag as unknown as NonNullable<ServiceRegistry['enhancedRag']>);
        logger.info('🚀 EnhancedRagService зареєстровано', {
          type: 'service_manager',
          event: 'enhanced_rag_registered',
          component: 'ServiceManager',
        });
      } else {
        logger.warn('EnhancedRagService не зареєстровано: недостатньо залежностей');
      }
    } catch (er) {
      logger.error('❌ Не вдалося створити EnhancedRagService', {
        type: 'service_manager',
        event: 'enhanced_rag_register_failed',
        component: 'ServiceManager',
        errorMessage: er instanceof Error ? er.message : String(er),
      });
    }

    // Зв'язуємо MetricsService з GoogleService (якщо обидва доступні)
    try {
      const google = this.services.get('google');
      const metrics = this.services.get('metrics');
      const searchIndex = this.services.get('searchIndex');
      const embeddings = this.services.get('embeddings');
      
      (google as any)?.setMetricsService?.(metrics);
      
      // Встановлюємо індекс пошуку та сервіс ембеддінгів для GoogleDocsService
      if (google && searchIndex) {
        (google as any)?.setSearchIndex?.(searchIndex);
      }
      
      if (google && embeddings) {
        (google as any)?.setEmbeddingsService?.(embeddings);
      }
      
      if (google && metrics) {
        logger.debug('🔗 Підключено MetricsService до GoogleService');
      }
    } catch (e) {
      logger.warn('Не вдалося підключити MetricsService до GoogleService', { error: (e as Error).message });
    }
  }

  /**
   * Ініціалізація сервісів
   */
  private async initializeServices(): Promise<void> {
    const initPromises = Array.from(this.services.entries()).map(async ([name, service]) => {
      try {
        if ((service as any)?.initialize) {
          await (service as any).initialize();
          logger.debug('✅ Сервіс ініціалізовано', {
            type: 'service_manager',
            event: 'service_initialized',
            component: 'ServiceManager',
            service: name,
          });
        }
      } catch (error) {
        logger.error('❌ Помилка ініціалізації сервісу', {
          type: 'service_manager',
          event: 'service_init_failed',
          component: 'ServiceManager',
          service: name,
          errorName: error instanceof Error ? error.name : undefined,
          errorMessage: error instanceof Error ? error.message : String(error),
          stack: error instanceof Error ? error.stack : undefined,
        });

        // Видаляємо сервіс, який не вдалося ініціалізувати
        this.services.delete(name);
      }
    });

    await Promise.allSettled(initPromises);
  }

  /** Повертає сервіс за назвою */
  public getService<K extends ServiceKey>(name: K): ServiceRegistry[K] | undefined {
    return this.services.get(name) as ServiceRegistry[K] | undefined;
  }

  /** Повертає сервіс або кидає помилку, якщо його немає */
  public getRequiredService<K extends ServiceKey>(name: K): ServiceRegistry[K] {
    const svc = this.getService(name);
    if (!svc) {
      throw new Error(`Service '${name}' is not available`);
    }
    return svc;
  }

  /** Список зареєстрованих сервісів */
  public getServiceNames(): ServiceKey[] {
    return Array.from(this.services.keys());
  }

  /**
   * Запуск метрик
   */
  async startMetrics(): Promise<void> {
    const metricsService = this.services.get('metrics');
    if ((metricsService as any)?.start) {
      await (metricsService as any).start();
      logger.info('📊 Метрики запущено', {
        type: 'service_manager',
        event: 'metrics_started',
        component: 'ServiceManager',
      });
    }
  }

  /**
   * Запуск кешування
   */
  async startCache(): Promise<void> {
    const cacheService = this.services.get('cache');
    if ((cacheService as any)?.start) {
      await (cacheService as any).start();
      logger.info('💾 Кеш запущено', {
        type: 'service_manager',
        event: 'cache_started',
        component: 'ServiceManager',
      });
    }
  }

  /**
   * Запуск планувальника
   */
  async startScheduler(): Promise<void> {
    const schedulerService = this.services.get('scheduler');
    if ((schedulerService as any)?.start) {
      await (schedulerService as any).start();
      logger.info('⏰ Планувальник запущено', {
        type: 'service_manager',
        event: 'scheduler_started',
        component: 'ServiceManager',
      });
    }
  }

  /**
   * Отримання сервісу за назвою
   */
  // getService(name: string): Service | undefined { // removed duplicate; prefer generic variant above
  //   return this.services.get(name);
  // }

  /**
   * Перевірка наявності сервісу
   */
  hasService(name: ServiceKey): boolean {
    return this.services.has(name);
  }

  /**
   * Отримання всіх сервісів
   */
  getAllServices(): Array<NonNullable<ServiceRegistry[ServiceKey]>> {
    return Array.from(this.services.values());
  }

  /**
   * Отримання назв всіх сервісів
   */
  // getServiceNames(): string[] { // removed duplicate; prefer public generic variant above
  //   return Array.from(this.services.keys());
  // }

  /**
   * Виконання методу на всіх сервісах
   */
  async executeOnAllServices(
    methodName: string,
    ...args: any[]
  ): Promise<PromiseSettledResult<any>[]> {
    const promises = Array.from(this.services.values()).map(async service => {
      const fn = (service as any)[methodName];
      if (typeof fn === 'function') {
        try {
          return await fn.apply(service, args);
        } catch (error) {
          logger.error('Помилка виконання методу на сервісі', {
            type: 'service_manager',
            event: 'method_execution_failed',
            component: 'ServiceManager',
            methodName: methodName,
            errorName: error instanceof Error ? error.name : undefined,
            errorMessage: error instanceof Error ? error.message : String(error),
            stack: error instanceof Error ? error.stack : undefined,
          });
          return null;
        }
      }
      return null;
    });

    return Promise.allSettled(promises);
  }

  /**
   * Отримання статусу сервісів
   */
  getServicesStatus(): Record<string, ServiceStatus> {
    const status: Record<string, ServiceStatus> = {};

    for (const [name, service] of this.services.entries()) {
      status[name] = {
        isActive: (service as any)?.isActive ? (service as any).isActive() : true,
        hasMethod: (method: string) => typeof (service as any)[method] === 'function',
        stats: (service as any)?.getStats ? (service as any).getStats() : null,
      };
    }

    return status;
  }

  /**
   * Graceful shutdown всіх сервісів
   */
  async shutdown(): Promise<void> {
    logger.info('🛑 Завершення роботи сервісів...', {
      type: 'service_manager',
      event: 'shutdown',
      component: 'ServiceManager',
    });

    try {
      await this.executeOnAllServices('shutdown');
      logger.info('✅ Сервіси успішно завершено', {
        type: 'service_manager',
        event: 'shutdown_success',
        component: 'ServiceManager',
      });
    } catch (error) {
      logger.error('❌ Помилка при завершенні сервісів', {
        type: 'service_manager',
        event: 'shutdown_failed',
        component: 'ServiceManager',
        errorName: error instanceof Error ? error.name : undefined,
        errorMessage: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
      });
    }
  }

  /**
   * Статистика сервісів
   */
  getStats(): ServiceManagerStats {
    return {
      total: this.services.size,
      active: Array.from(this.services.values()).filter(service =>
        (service as any)?.isActive ? (service as any).isActive() : true
      ).length,
      services: this.getServiceNames(),
      status: this.getServicesStatus(),
    };
  }
}

export default ServiceManager;
