/**
 * Slack Integration Service for Discord AI Assistant Bot
 * Provides integration with Slack for document notifications and collaboration
 * Version 1.0.0
 */

import type { BotConfig } from '@/types';
import { BaseService } from '@/core/BaseService';
import logger from '@/utils/logger';
import { WebClient } from '@slack/web-api';

// Types for Slack integration
export interface SlackNotificationConfig {
  channelId: string;
  botToken: string;
  enabled: boolean;
}

export interface DocumentEvent {
  fileId: string;
  fileName: string;
  userId: string;
  userName: string;
  action: 'created' | 'updated' | 'deleted' | 'shared' | 'downloaded' | 'analyzed';
  timestamp: Date;
  metadata?: Record<string, any>;
}

export interface SlackMessageOptions {
  text?: string;
  blocks?: any[];
  attachments?: any[];
}

export interface SlackIntegrationStats {
  totalNotificationsSent: number;
  totalNotificationsFailed: number;
  averageResponseTime: number;
  totalResponseTime: number;
  lastNotification?: {
    timestamp: Date;
    success: boolean;
    channelId: string;
  };
}

export class SlackIntegrationService extends BaseService {
  private slackClient: WebClient | null = null;
  private stats: SlackIntegrationStats;
  private notificationQueue: DocumentEvent[] = [];
  private isProcessingQueue = false;

  constructor(config: BotConfig) {
    super('SlackIntegrationService', config);
    
    this.stats = {
      totalNotificationsSent: 0,
      totalNotificationsFailed: 0,
      averageResponseTime: 0,
      totalResponseTime: 0,
    };

    // Initialize Slack client if configured
    this.initializeSlackClient();
  }

  /**
   * Initialize Slack client with bot token from config
   */
  private initializeSlackClient(): void {
    try {
      const slackConfig = (this.config as any).integrations?.slack;
      
      if (slackConfig?.enabled && slackConfig.botToken) {
        this.slackClient = new WebClient(slackConfig.botToken);
        logger.info('🔗 Slack client initialized successfully', {
          component: 'SlackIntegrationService'
        });
      } else {
        logger.info('⏭️ Slack integration not configured or disabled', {
          component: 'SlackIntegrationService'
        });
      }
    } catch (error) {
      logger.error('❌ Error initializing Slack client', {
        component: 'SlackIntegrationService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Send document event notification to Slack
   */
  public async sendDocumentNotification(event: DocumentEvent, options?: SlackMessageOptions): Promise<boolean> {
    // If Slack is not configured, return early
    const integrations = (this.config as any).integrations;
    if (!this.slackClient || !integrations?.slack?.enabled) {
      logger.debug('⏭️ Slack notification skipped - integration not enabled', {
        component: 'SlackIntegrationService',
        fileId: event.fileId,
        action: event.action
      });
      return false;
    }

    const startTime = Date.now();
    
    try {
      const channelId = integrations.slack.channelId;
      
      if (!channelId) {
        throw new Error('Slack channel ID not configured');
      }

      // Create default message if none provided
      const message = options || this.createDefaultMessage(event);
      
      // Send message to Slack
      const response = await this.slackClient.chat.postMessage({
        channel: channelId,
        ...message,
        attachments: message.attachments || [] // Ensure attachments is always provided
      });

      // Update stats
      const duration = Date.now() - startTime;
      this.updateStats(true, duration);

      logger.info('✅ Document notification sent to Slack', {
        component: 'SlackIntegrationService',
        fileId: event.fileId,
        fileName: event.fileName,
        action: event.action,
        channelId,
        duration: `${duration}ms`
      });

      return response.ok as boolean;
    } catch (error) {
      const duration = Date.now() - startTime;
      this.updateStats(false, duration);
      
      logger.error('❌ Error sending document notification to Slack', {
        component: 'SlackIntegrationService',
        fileId: event.fileId,
        fileName: event.fileName,
        action: event.action,
        error: error instanceof Error ? error.message : String(error)
      });
      
      // Add to queue for retry
      this.addToQueue(event);
      
      return false;
    }
  }

  /**
   * Send batch notifications to Slack
   */
  public async sendBatchNotifications(events: DocumentEvent[]): Promise<boolean[]> {
    const results: boolean[] = [];
    
    for (const event of events) {
      try {
        const result = await this.sendDocumentNotification(event);
        results.push(result);
      } catch (error) {
        logger.error('Error sending batch notification', {
          component: 'SlackIntegrationService',
          fileId: event.fileId,
          error: error instanceof Error ? error.message : String(error)
        });
        results.push(false);
      }
    }
    
    return results;
  }

  /**
   * Create default Slack message for document event
   */
  private createDefaultMessage(event: DocumentEvent): SlackMessageOptions {
    const actionEmoji = {
      created: '🆕',
      updated: '✏️',
      deleted: '🗑️',
      shared: '🔗',
      downloaded: '📥',
      analyzed: '🔍'
    }[event.action] || '📄';

    const actionText = {
      created: 'created',
      updated: 'updated',
      deleted: 'deleted',
      shared: 'shared',
      downloaded: 'downloaded',
      analyzed: 'analyzed'
    }[event.action] || 'modified';

    const isSensitive = this.isSensitiveDocument(event.fileName);
    const sensitivityWarning = isSensitive ? '\n⚠️ *Sensitive Document*' : '';

    return {
      blocks: [
        {
          type: 'section',
          text: {
            type: 'mrkdwn',
            text: `${actionEmoji} *Document ${actionText}*\n` +
                  `*File:* ${event.fileName}\n` +
                  `*User:* ${event.userName}\n` +
                  `*Action:* ${event.action}\n` +
                  `*Time:* ${event.timestamp.toLocaleString()}` +
                  sensitivityWarning
          }
        },
        {
          type: 'context',
          elements: [
            {
              type: 'mrkdwn',
              text: `ID: ${event.fileId} | User ID: ${event.userId}`
            }
          ]
        }
      ]
    };
  }

  /**
   * Check if document is sensitive based on filename
   */
  private isSensitiveDocument(fileName: string): boolean {
    const sensitiveKeywords = [
      'confidential', 'secret', 'private', 'internal', 'restricted', 
      'classified', 'proprietary', 'sensitive', 'password', 'credential'
    ];
    
    const lowerFileName = fileName.toLowerCase();
    
    return sensitiveKeywords.some(keyword => lowerFileName.includes(keyword));
  }

  /**
   * Add event to notification queue for retry
   */
  private addToQueue(event: DocumentEvent): void {
    this.notificationQueue.push(event);
    
    // Process queue if not already processing
    if (!this.isProcessingQueue) {
      this.processQueue();
    }
  }

  /**
   * Process notification queue with retry logic
   */
  private async processQueue(): Promise<void> {
    if (this.notificationQueue.length === 0) {
      this.isProcessingQueue = false;
      return;
    }

    this.isProcessingQueue = true;
    
    // Process one item at a time
    const event = this.notificationQueue.shift();
    
    if (event) {
      try {
        // Wait a bit before retry
        await new Promise(resolve => setTimeout(resolve, 5000));
        
        // Try to send notification again
        await this.sendDocumentNotification(event);
      } catch (error) {
        logger.warn('⚠️ Failed to process queued notification', {
          component: 'SlackIntegrationService',
          fileId: event.fileId,
          error: error instanceof Error ? error.message : String(error)
        });
        
        // Add back to queue for another retry
        this.notificationQueue.push(event);
      }
    }
    
    // Continue processing queue
    setTimeout(() => this.processQueue(), 1000);
  }

  /**
   * Update service statistics
   */
  private updateStats(success: boolean, duration: number): void {
    try {
      if (success) {
        this.stats.totalNotificationsSent++;
        this.stats.totalResponseTime += duration;
        if (this.stats.totalNotificationsSent > 0) {
          this.stats.averageResponseTime = this.stats.totalResponseTime / this.stats.totalNotificationsSent;
        }
      } else {
        this.stats.totalNotificationsFailed++;
      }
      
      const integrations = (this.config as any).integrations;
      this.stats.lastNotification = {
        timestamp: new Date(),
        success,
        channelId: integrations?.slack?.channelId || 'unknown'
      };
    } catch (error) {
      logger.warn('⚠️ Error updating Slack integration stats', {
        component: 'SlackIntegrationService',
        error: error instanceof Error ? error.message : String(error)
      });
    }
  }

  /**
   * Get service statistics
   */
  override getStats(): any {
    return { ...this.stats };
  }

  // === BaseService required methods ===
  
  protected async onInitialize(): Promise<void> {
    logger.info('🔗 Slack Integration Service initialized', {
      component: 'SlackIntegrationService',
      configured: this.isConfigured()
    });
  }

  protected async onShutdown(): Promise<void> {
    logger.info('🧹 Slack Integration Service shutdown', {
      component: 'SlackIntegrationService'
    });
  }

  protected async onHealthCheck(): Promise<any> {
    return {
      healthy: true,
      service: this.name,
      configured: this.isConfigured()
    };
  }

  protected onGetStats(): any {
    return this.getStats();
  }
  
  // Fixing the methods that access config to use proper casting
  
  public async sendCustomMessage(message: SlackMessageOptions, channelId?: string): Promise<boolean> {
    const integrations = (this.config as any).integrations;
    if (!this.slackClient || !integrations?.slack?.enabled) {
      return false;
    }

    const targetChannel = channelId || integrations.slack.channelId;
    
    if (!targetChannel) {
      throw new Error('Slack channel ID not configured');
    }

    try {
      const response = await this.slackClient.chat.postMessage({
        channel: targetChannel,
        ...message,
        attachments: message.attachments || [] // Ensure attachments is always provided
      });

      logger.info('✅ Custom message sent to Slack', {
        component: 'SlackIntegrationService',
        channelId: targetChannel
      });

      return response.ok as boolean;
    } catch (error) {
      logger.error('❌ Error sending custom message to Slack', {
        component: 'SlackIntegrationService',
        channelId: targetChannel,
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }
  
  public isConfigured(): boolean {
    const integrations = (this.config as any).integrations;
    return !!(
      this.slackClient && 
      integrations?.slack?.enabled && 
      integrations?.slack?.channelId
    );
  }
  
  public async testConnection(): Promise<boolean> {
    const integrations = (this.config as any).integrations;
    if (!this.slackClient || !integrations?.slack?.enabled) {
      return false;
    }

    try {
      const channelId = integrations.slack.channelId;
      
      if (!channelId) {
        return false;
      }

      // Test by sending a simple message
      const response = await this.slackClient.chat.postMessage({
        channel: channelId,
        text: '✅ Slack integration test successful!',
        attachments: [] // Ensure attachments is always provided
      });

      return response.ok as boolean;
    } catch (error) {
      logger.error('❌ Slack connection test failed', {
        component: 'SlackIntegrationService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      return false;
    }
  }
}

export default SlackIntegrationService;