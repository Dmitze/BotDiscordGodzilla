import { BaseService } from '@/core/BaseService';
import type { BotConfig, ServiceStats } from '@/types';
import logger from '@/utils/logger';
import { CacheService } from './CacheService';

interface OllamaMessage {
  role: 'user' | 'assistant' | 'system';
  content: string;
}

interface OllamaConfig {
  host: string;
  model: string;
  ctx: number;
  chatMaxLength: number;
}

interface OllamaServiceStats extends ServiceStats {
  requests: number;
  errors: number;
  avgResponseTime: number;
}

export class OllamaService extends BaseService {
  private ollamaConfig: OllamaConfig;
  private cacheService: CacheService | null = null;
  private stats: OllamaServiceStats;

  constructor(config: BotConfig, cacheService?: CacheService) {
    super('OllamaService', config);
    this.cacheService = cacheService || null;
    
    this.ollamaConfig = {
      host: (config as any).ai?.ollama?.host || 'http://localhost:11434',
      model: (config as any).ai?.ollama?.model || 'llama3',
      ctx: (config as any).ai?.ollama?.ctx || 2048,
      chatMaxLength: (config as any).ai?.ollama?.chatMaxLength || 500,
    };
    
    this.stats = {
      service: 'OllamaService',
      uptime: 0,
      requests: 0,
      errors: 0,
      avgResponseTime: 0,
    };
  }

  protected override async onInitialize(): Promise<void> {
    try {
      // Test connection to Ollama
      const response = await fetch(`${this.ollamaConfig.host}/api/tags`);
      if (!response.ok) {
        throw new Error(`Failed to connect to Ollama: ${response.status} ${response.statusText}`);
      }
      
      logger.info('✅ Ollama Service initialized', { 
        component: 'OllamaService',
        host: this.ollamaConfig.host,
        model: this.ollamaConfig.model
      });
    } catch (error) {
      logger.error('❌ Failed to initialize Ollama Service:', { 
        component: 'OllamaService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Generate a response from Ollama
   */
  public async generate(prompt: string, options: { 
    model?: string; 
    temperature?: number; 
    maxTokens?: number;
    channelId?: string;
  } = {}): Promise<string> {
    const startTime = Date.now();
    this.stats.requests++;
    
    try {
      // Get conversation history if channelId is provided
      let messages: OllamaMessage[] = [];
      if (options.channelId) {
        messages = await this.getChannelHistory(options.channelId);
      }
      
      // Add user message to history
      messages.push({ role: 'user', content: prompt });
      
      // Limit history length
      if (messages.length > this.ollamaConfig.chatMaxLength) {
        messages = messages.slice(-this.ollamaConfig.chatMaxLength);
      }
      
      const response = await fetch(`${this.ollamaConfig.host}/api/chat`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: options.model || this.ollamaConfig.model,
          messages,
          stream: false,
          options: {
            temperature: options.temperature !== undefined ? options.temperature : 0.7,
            num_predict: options.maxTokens || 1000,
            num_ctx: this.ollamaConfig.ctx,
          },
        }),
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Ollama API error: ${response.status} ${response.statusText} - ${errorText}`);
      }

      const data: any = await response.json();
      const responseTime = Date.now() - startTime;
      
      // Update average response time
      this.stats.avgResponseTime = 
        (this.stats.avgResponseTime * (this.stats.requests - 1) + responseTime) / this.stats.requests;
      
      const responseText = data.message?.content || '';
      
      // Add assistant response to history
      if (options.channelId) {
        messages.push({ role: 'assistant', content: responseText });
        await this.saveChannelHistory(options.channelId, messages);
      }
      
      logger.debug('✅ Ollama response generated', {
        component: 'OllamaService',
        model: data.model || this.ollamaConfig.model,
        responseTime: `${responseTime}ms`,
        responseLength: responseText.length
      });
      
      return responseText;
    } catch (error) {
      this.stats.errors++;
      const responseTime = Date.now() - startTime;
      
      logger.error('❌ Ollama request failed:', {
        component: 'OllamaService',
        error: error instanceof Error ? error.message : String(error),
        responseTime: `${responseTime}ms`
      });
      
      throw error;
    }
  }

  /**
   * Get conversation history for a channel
   */
  private async getChannelHistory(channelId: string): Promise<OllamaMessage[]> {
    if (!this.cacheService) {
      return [];
    }
    
    try {
      const key = `ollama:channel:${channelId}`;
      const cached = await this.cacheService.get<OllamaMessage[]>(key);
      return cached || [];
    } catch (error) {
      logger.warn('⚠️ Failed to get channel history from cache:', {
        component: 'OllamaService',
        error: error instanceof Error ? error.message : String(error)
      });
      return [];
    }
  }

  /**
   * Save conversation history for a channel
   */
  private async saveChannelHistory(channelId: string, messages: OllamaMessage[]): Promise<void> {
    if (!this.cacheService) {
      return;
    }
    
    try {
      const key = `ollama:channel:${channelId}`;
      await this.cacheService.set(key, messages, 60 * 60 * 24 * 7); // 7 days
    } catch (error) {
      logger.warn('⚠️ Failed to save channel history to cache:', {
        component: 'OllamaService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Reset conversation history for a channel
   */
  public async resetChannelHistory(channelId: string): Promise<void> {
    if (!this.cacheService) {
      return;
    }
    
    try {
      const key = `ollama:channel:${channelId}`;
      await this.cacheService.delete(key);
      logger.info('🧹 Channel history reset', { 
        component: 'OllamaService',
        channelId
      });
    } catch (error) {
      logger.error('❌ Failed to reset channel history:', {
        component: 'OllamaService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Pull a model from Ollama
   */
  public async pullModel(modelName: string): Promise<void> {
    try {
      const response = await fetch(`${this.ollamaConfig.host}/api/pull`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ name: modelName }),
      });
      
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to pull model: ${response.status} ${response.statusText} - ${errorText}`);
      }
      
      logger.info('📥 Model pulled successfully', { 
        component: 'OllamaService',
        modelName
      });
    } catch (error) {
      logger.error('❌ Failed to pull model:', {
        component: 'OllamaService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * List available models
   */
  public async listModels(): Promise<any[]> {
    try {
      const response = await fetch(`${this.ollamaConfig.host}/api/tags`);
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to list models: ${response.status} ${response.statusText} - ${errorText}`);
      }
      
      const data: any = await response.json();
      return data.models || [];
    } catch (error) {
      logger.error('❌ Failed to list models:', {
        component: 'OllamaService',
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * Get service statistics
   */
  public override getStats(): OllamaServiceStats {
    return {
      ...this.stats,
      uptime: Date.now() - this.startTime,
    };
  }

  protected override async onShutdown(): Promise<void> {
    logger.info('🛑 Ollama Service shutdown', { component: 'OllamaService' });
  }

  protected override async onHealthCheck(): Promise<any> {
    try {
      const response = await fetch(`${this.ollamaConfig.host}/api/tags`);
      return {
        healthy: response.ok,
        service: 'OllamaService',
        message: response.ok ? 'Ollama is available' : 'Ollama is not available'
      };
    } catch (error) {
      return {
        healthy: false,
        service: 'OllamaService',
        error: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  }

  protected override onGetStats(): Partial<ServiceStats> {
    return {
      requests: this.stats.requests,
      errors: this.stats.errors,
      avgResponseTime: this.stats.avgResponseTime,
    };
  }

  public override async healthCheck(): Promise<any> {
    try {
      const response = await fetch(`${this.ollamaConfig.host}/api/tags`);
      return {
        healthy: response.ok,
        service: 'OllamaService',
        message: response.ok ? 'Ollama is available' : 'Ollama is not available'
      };
    } catch (error) {
      return {
        healthy: false,
        service: 'OllamaService',
        error: error instanceof Error ? error.message : 'Unknown error'
      };
    }
  }
}
