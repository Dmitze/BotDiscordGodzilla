/**
 * Unit тесты для MetricsService
 */

import { jest, describe, it, expect, beforeEach, afterEach } from '@jest/globals';
import { MetricsService } from '../../../services/MetricsService';
import { createMockConfig } from '../../utils/testHelpers';

// Моки для prom-client
jest.mock('prom-client', () => ({
  Registry: jest.fn(() => ({
    registerMetric: jest.fn(),
    metrics: (jest.fn() as any).mockResolvedValue('test_metrics' as any),
    clear: jest.fn(),
  })),
  Counter: jest.fn(() => ({
    inc: jest.fn(),
    get: jest.fn(() => ({ values: [{ value: 10 }] })),
  })),
  Histogram: jest.fn(() => ({
    observe: jest.fn(),
    get: jest.fn(() => ({ values: [{ value: 0.5 }] })),
  })),
  Gauge: jest.fn(() => ({
    set: jest.fn(),
    inc: jest.fn(),
    dec: jest.fn(),
    get: jest.fn(() => ({ values: [{ value: 100 }] })),
  })),
}));

describe('MetricsService', () => {
  let metricsService: MetricsService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    metricsService = new MetricsService(mockConfig);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('constructor', () => {
    it('should create MetricsService instance', () => {
      expect(metricsService).toBeInstanceOf(MetricsService);
    });

    it('should have correct service name', () => {
      expect(metricsService.getName()).toBe('MetricsService');
    });
  });

  describe('initialization', () => {
    it('should initialize successfully when metrics are enabled', async () => {
      mockConfig.metrics.enabled = true;

      await expect(metricsService.initialize()).resolves.not.toThrow();
    });

    it('should skip initialization when metrics are disabled', async () => {
      mockConfig.metrics.enabled = false;

      await expect(metricsService.initialize()).resolves.not.toThrow();
    });

    it('should create metrics when enabled', async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();

      expect(metricsService.isInitialized()).toBe(true);
    });
  });

  describe('counter metrics', () => {
    beforeEach(async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();
    });

    it('should increment command counter', () => {
      metricsService.incrementCommand('пошук');

      const counter = (metricsService as any).commandCounter;
      expect(counter.inc).toHaveBeenCalledWith({ command: 'пошук' });
    });

    it('should increment error counter', () => {
      metricsService.incrementError('api_error');

      const counter = (metricsService as any).errorCounter;
      expect(counter.inc).toHaveBeenCalledWith({ type: 'api_error' });
    });

    it('should increment user counter', () => {
      metricsService.incrementUser('user123');

      const counter = (metricsService as any).userCounter;
      expect(counter.inc).toHaveBeenCalledWith({ user: 'user123' });
    });
  });

  describe('histogram metrics', () => {
    beforeEach(async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();
    });

    it('should observe command duration', () => {
      metricsService.observeCommandDuration('пошук', 150);

      const histogram = (metricsService as any).commandDuration;
      expect(histogram.observe).toHaveBeenCalledWith({ command: 'пошук' }, 150);
    });

    it('should observe response time', () => {
      metricsService.observeResponseTime('google_api', 200);

      const histogram = (metricsService as any).responseTime;
      expect(histogram.observe).toHaveBeenCalledWith({ service: 'google_api' }, 200);
    });
  });

  describe('gauge metrics', () => {
    beforeEach(async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();
    });

    it('should set active users gauge', () => {
      metricsService.setActiveUsers(50);

      const gauge = (metricsService as any).activeUsers;
      expect(gauge.set).toHaveBeenCalledWith(50);
    });

    it('should increment cache hits', () => {
      metricsService.incrementCacheHits();

      const gauge = (metricsService as any).cacheHits;
      expect(gauge.inc).toHaveBeenCalled();
    });

    it('should increment cache misses', () => {
      metricsService.incrementCacheMisses();

      const gauge = (metricsService as any).cacheMisses;
      expect(gauge.inc).toHaveBeenCalled();
    });

    it('should set memory usage', () => {
      metricsService.setMemoryUsage(1024);

      const gauge = (metricsService as any).memoryUsage;
      expect(gauge.set).toHaveBeenCalledWith(1024);
    });
  });

  describe('metrics collection', () => {
    beforeEach(async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();
    });

    it('should get metrics string', async () => {
      const metrics = await metricsService.getMetrics();

      expect(metrics).toBe('test_metrics');
    });

    it('should get metrics registry', () => {
      const registry = metricsService.getRegistry();

      expect(registry).toBeDefined();
    });
  });

  describe('custom metrics', () => {
    beforeEach(async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();
    });

    it('should create custom counter', () => {
      const counter = metricsService.createCounter('custom_counter', 'Custom counter');

      expect(counter).toBeDefined();
    });

    it('should create custom histogram', () => {
      const histogram = metricsService.createHistogram('custom_histogram', 'Custom histogram');

      expect(histogram).toBeDefined();
    });

    it('should create custom gauge', () => {
      const gauge = metricsService.createGauge('custom_gauge', 'Custom gauge');

      expect(gauge).toBeDefined();
    });
  });

  describe('metrics reporting', () => {
    beforeEach(async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();
    });

    it('should get metrics summary', () => {
      const summary = metricsService.getMetricsSummary();

      expect(summary).toHaveProperty('totalCommands');
      expect(summary).toHaveProperty('totalErrors');
      expect(summary).toHaveProperty('activeUsers');
      expect(summary).toHaveProperty('cacheHitRate');
    });

    it('should calculate cache hit rate', () => {
      // Симулируем hits и misses
      const hitsGauge = (metricsService as any).cacheHits;
      const missesGauge = (metricsService as any).cacheMisses;
      
      hitsGauge.get.mockReturnValue({ values: [{ value: 80 }] });
      missesGauge.get.mockReturnValue({ values: [{ value: 20 }] });

      const summary = metricsService.getMetricsSummary();

      expect(summary.cacheHitRate).toBe(0.8); // 80 / (80 + 20)
    });

    it('should handle zero cache requests', () => {
      const hitsGauge = (metricsService as any).cacheHits;
      const missesGauge = (metricsService as any).cacheMisses;
      
      hitsGauge.get.mockReturnValue({ values: [{ value: 0 }] });
      missesGauge.get.mockReturnValue({ values: [{ value: 0 }] });

      const summary = metricsService.getMetricsSummary();

      expect(summary.cacheHitRate).toBe(0);
    });
  });

  describe('health check', () => {
    it('should return healthy status when metrics are enabled', async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();

      const health = await metricsService.getHealthStatus();

      expect(health.healthy).toBe(true);
      expect(health.service).toBe('MetricsService');
    });

    it('should return healthy status when metrics are disabled', async () => {
      mockConfig.metrics.enabled = false;

      const health = await metricsService.getHealthStatus();

      expect(health.healthy).toBe(true);
      expect(health.service).toBe('MetricsService');
    });
  });

  describe('shutdown', () => {
    it('should clear metrics on shutdown', async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();

      await metricsService.shutdown();

      const registry = (metricsService as any).registry;
      expect(registry.clear).toHaveBeenCalled();
    });

    it('should handle shutdown when metrics are disabled', async () => {
      mockConfig.metrics.enabled = false;

      await expect(metricsService.shutdown()).resolves.not.toThrow();
    });
  });

  describe('error handling', () => {
    it('should handle metrics collection error', async () => {
      mockConfig.metrics.enabled = true;
      await metricsService.initialize();

      const registry = (metricsService as any).registry;
      registry.metrics.mockRejectedValue(new Error('Metrics error'));

      await expect(metricsService.getMetrics()).rejects.toThrow('Metrics error');
    });

    it('should handle disabled metrics gracefully', () => {
      mockConfig.metrics.enabled = false;

      // Методы должны работать без ошибок
      expect(() => metricsService.incrementCommand('test')).not.toThrow();
      expect(() => metricsService.incrementError('test')).not.toThrow();
      expect(() => metricsService.observeCommandDuration('test', 100)).not.toThrow();
    });
  });
}); 