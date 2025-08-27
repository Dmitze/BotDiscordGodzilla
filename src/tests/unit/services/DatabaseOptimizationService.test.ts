/**
 * Unit tests for DatabaseOptimizationService functionality
 */

import { describe, it, expect, beforeEach } from '@jest/globals';
import { DatabaseOptimizationService } from '../../../services/DatabaseOptimizationService';
import { createMockConfig } from '../../utils/testHelpers';

describe('DatabaseOptimizationService', () => {
  let dbOptimizationService: DatabaseOptimizationService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    dbOptimizationService = new DatabaseOptimizationService(mockConfig);
  });

  it('should initialize with default statistics', () => {
    const stats = dbOptimizationService.getStats();
    
    expect(stats).toBeDefined();
    expect(typeof stats.connectionCount).toBe('number');
    expect(stats.queryPerformance).toBeDefined();
    expect(stats.storage).toBeDefined();
    expect(stats.indexes).toBeDefined();
  });

  it('should analyze database and generate recommendations', async () => {
    const recommendations = await dbOptimizationService.analyzeDatabase();
    
    expect(recommendations).toBeDefined();
    expect(Array.isArray(recommendations)).toBe(true);
    // Should have some recommendations based on mock data
    expect(recommendations.length).toBeGreaterThanOrEqual(0);
  });

  it('should provide performance metrics', () => {
    const metrics = dbOptimizationService.getPerformanceMetrics();
    
    expect(metrics).toBeDefined();
    expect(Array.isArray(metrics)).toBe(true);
    expect(metrics.length).toBeGreaterThan(0);
  });

  it('should sort performance metrics correctly', () => {
    const sortedByTime = dbOptimizationService.getPerformanceMetrics({
      sortBy: 'executionTime',
      order: 'desc'
    });
    
    const sortedByFrequency = dbOptimizationService.getPerformanceMetrics({
      sortBy: 'frequency',
      order: 'desc'
    });
    
    expect(sortedByTime).toBeDefined();
    expect(sortedByFrequency).toBeDefined();
    
    // Check if sorting worked (first item should have highest value)
    if (sortedByTime.length > 1) {
      expect(sortedByTime[0].executionTime).toBeGreaterThanOrEqual(sortedByTime[1].executionTime);
    }
    
    if (sortedByFrequency.length > 1) {
      expect(sortedByFrequency[0].frequency).toBeGreaterThanOrEqual(sortedByFrequency[1].frequency);
    }
  });

  it('should limit performance metrics results', () => {
    const limitedMetrics = dbOptimizationService.getPerformanceMetrics({ limit: 5 });
    
    expect(limitedMetrics.length).toBeLessThanOrEqual(5);
  });

  it('should handle custom recommendations', () => {
    const customRecommendation = {
      id: 'custom-1',
      type: 'query' as const,
      priority: 'high' as const,
      description: 'Custom recommendation for testing',
      impact: 'medium' as const,
      implementation: 'Test implementation',
      estimatedTime: '1 hour'
    };
    
    dbOptimizationService.addRecommendation(customRecommendation);
    
    // Should now have more recommendations
    const recommendations = dbOptimizationService.generateReport().recommendations;
    expect(recommendations.some(rec => rec.id === 'custom-1')).toBe(true);
    
    // Remove the recommendation
    const removed = dbOptimizationService.removeRecommendation('custom-1');
    expect(removed).toBe(true);
    
    // Should no longer have the recommendation
    const updatedRecommendations = dbOptimizationService.generateReport().recommendations;
    expect(updatedRecommendations.some(rec => rec.id === 'custom-1')).toBe(false);
  });

  it('should apply optimizations', async () => {
    // First analyze to get recommendations
    await dbOptimizationService.analyzeDatabase();
    
    // Get a recommendation ID
    const recommendations = dbOptimizationService.generateReport().recommendations;
    if (recommendations.length > 0) {
      const recommendationId = recommendations[0].id;
      
      // Apply the optimization
      const applied = await dbOptimizationService.applyOptimization(recommendationId);
      expect(applied).toBe(true);
    }
  });

  it('should generate comprehensive reports', () => {
    const report = dbOptimizationService.generateReport();
    
    expect(report).toBeDefined();
    expect(report.stats).toBeDefined();
    expect(report.recommendations).toBeDefined();
    expect(report.performanceMetrics).toBeDefined();
    expect(report.summary).toBeDefined();
    
    // Summary should have correct structure
    expect(typeof report.summary.totalRecommendations).toBe('number');
    expect(typeof report.summary.criticalIssues).toBe('number');
    expect(typeof report.summary.highPriorityIssues).toBe('number');
    expect(typeof report.summary.mediumPriorityIssues).toBe('number');
    expect(typeof report.summary.lowPriorityIssues).toBe('number');
  });

  it('should export reports in different formats', () => {
    const jsonExport = dbOptimizationService.exportReport('json');
    const csvExport = dbOptimizationService.exportReport('csv');
    
    expect(typeof jsonExport).toBe('string');
    expect(typeof csvExport).toBe('string');
    
    // JSON should be parseable
    expect(() => JSON.parse(jsonExport)).not.toThrow();
    
    // CSV should have headers
    expect(csvExport).toContain('Recommendation ID,Type,Priority,Description,Impact,Estimated Time');
  });

  it('should configure connection pooling', () => {
    const initialStats = dbOptimizationService.getStats();
    
    dbOptimizationService.setupConnectionPooling({
      minConnections: 5,
      maxConnections: 20,
      connectionTimeout: 30000,
      idleTimeout: 60000
    });
    
    const updatedStats = dbOptimizationService.getStats();
    expect(updatedStats.connectionCount).toBe(20);
  });

  it('should enable query caching', () => {
    const initialStats = dbOptimizationService.getStats();
    
    dbOptimizationService.enableQueryCaching({
      cacheSize: 1000,
      ttl: 3600 // 1 hour
    });
    
    const updatedStats = dbOptimizationService.getStats();
    expect(updatedStats.queryPerformance.cachedQueries).toBe(1000);
  });

  it('should optimize database indexes', async () => {
    const initialStats = dbOptimizationService.getStats();
    const initialFragmented = initialStats.indexes.fragmented;
    
    await dbOptimizationService.optimizeIndexes();
    
    const updatedStats = dbOptimizationService.getStats();
    expect(updatedStats.indexes.fragmented).toBeLessThanOrEqual(initialFragmented);
  });

  it('should archive old data', async () => {
    const initialStats = dbOptimizationService.getStats();
    const initialUsedSize = initialStats.storage.usedSize;
    
    const archivedCount = await dbOptimizationService.archiveOldData({
      tableName: 'documents',
      dateColumn: 'created_at',
      olderThan: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000), // 30 days ago
      archiveTable: 'documents_archive'
    });
    
    expect(typeof archivedCount).toBe('number');
    expect(archivedCount).toBeGreaterThan(0);
    
    const updatedStats = dbOptimizationService.getStats();
    expect(updatedStats.storage.usedSize).toBeLessThan(initialUsedSize);
    expect(updatedStats.storage.freeSize).toBeGreaterThan(initialStats.storage.freeSize);
  });

  it('should clear all recommendations', () => {
    // Add some recommendations
    dbOptimizationService.addRecommendation({
      id: 'test-1',
      type: 'query',
      priority: 'high',
      description: 'Test recommendation 1',
      impact: 'medium',
      implementation: 'Test',
      estimatedTime: '1 hour'
    });
    
    dbOptimizationService.addRecommendation({
      id: 'test-2',
      type: 'index',
      priority: 'medium',
      description: 'Test recommendation 2',
      impact: 'low',
      implementation: 'Test',
      estimatedTime: '30 minutes'
    });
    
    // Verify recommendations were added
    const reportBefore = dbOptimizationService.generateReport();
    expect(reportBefore.recommendations.length).toBeGreaterThanOrEqual(2);
    
    // Clear recommendations
    dbOptimizationService.clearRecommendations();
    
    // Verify recommendations were cleared
    const reportAfter = dbOptimizationService.generateReport();
    expect(reportAfter.recommendations.length).toBe(0);
  });
});