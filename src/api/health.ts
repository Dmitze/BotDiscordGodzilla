/**
 * Health Check Endpoints
 * Provides monitoring endpoints for bot services
 */

import type { Bot } from '@/core/Bot';
import type { HealthStatus } from '@/types/index';
import logger from '@/utils/logger';

interface HealthCheckResponse {
  status: 'ok' | 'error';
  timestamp: string;
  services: Record<string, HealthStatus>;
  uptime: number;
}

/**
 * Create health check endpoints
 */
export function createHealthEndpoints(bot: Bot) {
  return {
    /**
     * Overall health check
     */
    async health(): Promise<HealthCheckResponse> {
      const startTime = Date.now();
      const services: Record<string, HealthStatus> = {};
      
      try {
        // Check AI Service
        try {
          const aiService = bot.getService('ai') as { onHealthCheck?: () => Promise<HealthStatus> } | undefined;
          if (aiService && typeof aiService.onHealthCheck === 'function') {
            services['ai'] = await aiService.onHealthCheck();
          } else {
            services['ai'] = {
              healthy: false,
              service: 'ai',
              error: 'AI service not available',
            };
          }
        } catch (error) {
          services['ai'] = {
            healthy: false,
            service: 'ai',
            error: error instanceof Error ? error.message : String(error),
          };
        }

        // Check Google Service
        try {
          const googleService = bot.getService('google') as { onHealthCheck?: () => Promise<HealthStatus> } | undefined;
          if (googleService && typeof googleService.onHealthCheck === 'function') {
            services['google'] = await googleService.onHealthCheck();
          } else {
            services['google'] = {
              healthy: false,
              service: 'google',
              error: 'Google service not available',
            };
          }
        } catch (error) {
          services['google'] = {
            healthy: false,
            service: 'google',
            error: error instanceof Error ? error.message : String(error),
          };
        }

        // Check Cache Service
        try {
          const cacheService = bot.getService('cache') as { onHealthCheck?: () => Promise<HealthStatus> } | undefined;
          if (cacheService && typeof cacheService.onHealthCheck === 'function') {
            services['cache'] = await cacheService.onHealthCheck();
          } else {
            services['cache'] = {
              healthy: true,
              service: 'cache',
              details: { message: 'Cache service initialized or not health-checkable' }, // Using proper Record<string, unknown> type
            };
          }
        } catch (error) {
          services['cache'] = {
            healthy: false,
            service: 'cache',
            error: error instanceof Error ? error.message : String(error),
          };
        }

        const uptime = Date.now() - startTime;
        
        // Determine overall status
        const allHealthy = Object.values(services).every(service => service.healthy);
        
        return {
          status: allHealthy ? 'ok' : 'error',
          timestamp: new Date().toISOString(),
          services,
          uptime,
        };
      } catch (error) {
        logger.error('Health check failed:', { error: error instanceof Error ? error.message : String(error) });
        return {
          status: 'error',
          timestamp: new Date().toISOString(),
          services: {
            healthEndpoint: {
              healthy: false,
              service: 'health',
              error: error instanceof Error ? error.message : String(error),
            },
          },
          uptime: Date.now() - startTime,
        };
      }
    },

    /**
     * AI Service health check
     */
    async aiHealth(): Promise<HealthStatus> {
      try {
        const aiService = bot.getService('ai') as { onHealthCheck?: () => Promise<HealthStatus> } | undefined;
        if (aiService && typeof aiService.onHealthCheck === 'function') {
          return await aiService.onHealthCheck();
        }
        return {
          healthy: false,
          service: 'ai',
          error: 'AI service not available',
        };
      } catch (error) {
        return {
          healthy: false,
          service: 'ai',
          error: error instanceof Error ? error.message : String(error),
        };
      }
    },

    /**
     * Google Service health check
     */
    async googleHealth(): Promise<HealthStatus> {
      try {
        const googleService = bot.getService('google') as { onHealthCheck?: () => Promise<HealthStatus> } | undefined;
        if (googleService && typeof googleService.onHealthCheck === 'function') {
          return await googleService.onHealthCheck();
        }
        return {
          healthy: false,
          service: 'google',
          error: 'Google service not available',
        };
      } catch (error) {
        return {
          healthy: false,
          service: 'google',
          error: error instanceof Error ? error.message : String(error),
        };
      }
    },

    /**
     * Detailed service status
     */
    async detailedStatus(): Promise<any> {
      const status: any = {};
      
      try {
        // AI Service detailed status
        try {
          const aiService = bot.getService('ai') as { getDetailedStatus?: () => Promise<any> } | undefined;
          if (aiService && typeof aiService.getDetailedStatus === 'function') {
            status.ai = await aiService.getDetailedStatus();
          }
        } catch (error) {
          status.ai = { error: error instanceof Error ? error.message : String(error) };
        }

        // Google Service detailed status
        try {
          const googleService = bot.getService('google') as { getDetailedStatus?: () => Promise<any> } | undefined;
          if (googleService && typeof googleService.getDetailedStatus === 'function') {
            status.google = await googleService.getDetailedStatus();
          }
        } catch (error) {
          status.google = { error: error instanceof Error ? error.message : String(error) };
        }

        return status;
      } catch (error) {
        logger.error('Detailed status check failed:', { error: error instanceof Error ? error.message : String(error) });
        return {
          error: error instanceof Error ? error.message : String(error),
        };
      }
    },
  };
}