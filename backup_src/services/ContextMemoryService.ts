/**
 * 🧠 Context Memory Service
 * Manages user query context and conversation history for enhanced AI responses
 */

import logger from '@/utils/logger';

export interface QueryContext {
  id: string;
  userId: string;
  channelId: string;
  query: string;
  response?: string;
  timestamp: Date;
  command?: string;
  parameters?: Record<string, any>;
  metadata?: {
    responseTime?: number;
    tokensUsed?: number;
    confidence?: number;
    sources?: string[];
  };
}

export interface UserContext {
  userId: string;
  queries: QueryContext[];
  preferences?: {
    language: string;
    domain: string;
    responseStyle: string;
  };
  lastActivity: Date;
}

export class ContextMemoryService {
  private static readonly MAX_QUERIES_PER_USER = 5;
  private static readonly CONTEXT_TTL_HOURS = 24;
  
  private userContexts: Map<string, UserContext> = new Map();
  private cleanupInterval?: NodeJS.Timeout | undefined;

  constructor() {
    this.startCleanupTask();
  }

  /**
   * 📝 Add user query to context memory
   */
  addQuery(
    userId: string,
    channelId: string,
    query: string,
    command?: string,
    parameters?: Record<string, any>
  ): string {
    const queryId = this.generateQueryId();
    const queryContext: QueryContext = {
      id: queryId,
      userId,
      channelId,
      query,
      timestamp: new Date()
    };
    
    if (command) queryContext.command = command;
    if (parameters) queryContext.parameters = parameters;

    let userContext = this.userContexts.get(userId);
    if (!userContext) {
      userContext = {
        userId,
        queries: [],
        lastActivity: new Date()
      };
      this.userContexts.set(userId, userContext);
    }

    // Add new query and maintain max limit
    userContext.queries.unshift(queryContext);
    if (userContext.queries.length > ContextMemoryService.MAX_QUERIES_PER_USER) {
      userContext.queries = userContext.queries.slice(0, ContextMemoryService.MAX_QUERIES_PER_USER);
    }
    
    userContext.lastActivity = new Date();

    logger.debug('Query added to context memory', {
      component: 'ContextMemoryService',
      userId,
      queryId,
      queriesCount: userContext.queries.length
    });

    return queryId;
  }

  /**
   * ✅ Update query with response information
   */
  updateQueryResponse(
    queryId: string,
    response: string,
    metadata?: QueryContext['metadata']
  ): void {
    for (const userContext of this.userContexts.values()) {
      const query = userContext.queries.find(q => q.id === queryId);
      if (query) {
        query.response = response;
        if (metadata) {
          query.metadata = metadata;
        }
        
        logger.debug('Query response updated', {
          component: 'ContextMemoryService',
          queryId,
          responseLength: response.length
        });
        return;
      }
    }

    logger.warn('Query not found for response update', {
      component: 'ContextMemoryService',
      queryId
    });
  }

  /**
   * 📚 Get user context history
   */
  getUserContext(userId: string): UserContext | undefined {
    return this.userContexts.get(userId);
  }

  /**
   * 🔍 Get recent queries for user
   */
  getRecentQueries(userId: string, limit: number = 5): QueryContext[] {
    const userContext = this.userContexts.get(userId);
    if (!userContext) {
      return [];
    }

    return userContext.queries.slice(0, Math.min(limit, userContext.queries.length));
  }

  /**
   * 🎯 Build contextual prompt from user history
   */
  buildContextualPrompt(userId: string, currentQuery: string): string {
    const recentQueries = this.getRecentQueries(userId, 3);
    
    if (recentQueries.length === 0) {
      return currentQuery;
    }

    let contextPrompt = `Контекст попередніх запитів користувача:\n`;
    
    recentQueries.reverse().forEach((query, index) => {
      contextPrompt += `${index + 1}. Запит: "${query.query}"`;
      if (query.command) {
        contextPrompt += ` (команда: ${query.command})`;
      }
      if (query.response) {
        const shortResponse = query.response.length > 100 
          ? query.response.substring(0, 100) + '...'
          : query.response;
        contextPrompt += `\n   Відповідь: "${shortResponse}"`;
      }
      contextPrompt += `\n`;
    });

    contextPrompt += `\nПоточний запит: "${currentQuery}"\n`;
    contextPrompt += `Враховуючи контекст попередніх запитів, надайте релевантну відповідь:`;

    return contextPrompt;
  }

  /**
   * ⚙️ Set user preferences
   */
  setUserPreferences(
    userId: string,
    preferences: UserContext['preferences']
  ): void {
    let userContext = this.userContexts.get(userId);
    if (!userContext) {
      userContext = {
        userId,
        queries: [],
        lastActivity: new Date()
      };
      this.userContexts.set(userId, userContext);
    }

    if (preferences) {
      userContext.preferences = preferences;
    }
    userContext.lastActivity = new Date();

    logger.debug('User preferences updated', {
      component: 'ContextMemoryService',
      userId,
      preferences
    });
  }

  /**
   * 🗑️ Clear user context
   */
  clearUserContext(userId: string): void {
    this.userContexts.delete(userId);
    
    logger.debug('User context cleared', {
      component: 'ContextMemoryService',
      userId
    });
  }

  /**
   * 📊 Get context statistics
   */
  getStats(): {
    totalUsers: number;
    totalQueries: number;
    avgQueriesPerUser: number;
    oldestContext: Date | null;
  } {
    const totalUsers = this.userContexts.size;
    let totalQueries = 0;
    let oldestContext: Date | null = null;

    for (const userContext of this.userContexts.values()) {
      totalQueries += userContext.queries.length;
      
      if (!oldestContext || userContext.lastActivity < oldestContext) {
        oldestContext = userContext.lastActivity;
      }
    }

    return {
      totalUsers,
      totalQueries,
      avgQueriesPerUser: totalUsers > 0 ? Math.round(totalQueries / totalUsers * 100) / 100 : 0,
      oldestContext
    };
  }

  /**
   * 🔄 Start periodic cleanup task
   */
  private startCleanupTask(): void {
    // Run cleanup every hour
    this.cleanupInterval = setInterval(() => {
      this.cleanupExpiredContexts();
    }, 60 * 60 * 1000);

    logger.debug('Context memory cleanup task started', {
      component: 'ContextMemoryService',
      intervalHours: 1
    });
  }

  /**
   * 🧹 Clean up expired contexts
   */
  private cleanupExpiredContexts(): void {
    const now = new Date();
    const expiredUsers: string[] = [];

    for (const [userId, userContext] of this.userContexts.entries()) {
      const hoursSinceLastActivity = (now.getTime() - userContext.lastActivity.getTime()) / (1000 * 60 * 60);
      
      if (hoursSinceLastActivity > ContextMemoryService.CONTEXT_TTL_HOURS) {
        expiredUsers.push(userId);
      }
    }

    for (const userId of expiredUsers) {
      this.userContexts.delete(userId);
    }

    if (expiredUsers.length > 0) {
      logger.debug('Expired contexts cleaned up', {
        component: 'ContextMemoryService',
        expiredCount: expiredUsers.length,
        remainingUsers: this.userContexts.size
      });
    }
  }

  /**
   * 🆔 Generate unique query ID
   */
  private generateQueryId(): string {
    return `query_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * 🛑 Shutdown service
   */
  shutdown(): void {
    if (this.cleanupInterval) {
      clearInterval(this.cleanupInterval);
      this.cleanupInterval = undefined;
    }

    logger.debug('ContextMemoryService shutdown completed', {
      component: 'ContextMemoryService'
    });
  }
}