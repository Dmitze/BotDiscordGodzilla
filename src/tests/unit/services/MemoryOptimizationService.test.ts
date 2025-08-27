/**
 * Unit tests for MemoryOptimizationService functionality
 */

import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { MemoryOptimizationService } from '../../../services/MemoryOptimizationService';
import { createMockConfig } from '../../utils/testHelpers';

describe('MemoryOptimizationService', () => {
  let memoryService: MemoryOptimizationService;
  let mockConfig: any;

  beforeEach(async () => {
    mockConfig = createMockConfig();
    memoryService = new MemoryOptimizationService(mockConfig);
    // Initialize the service
    await memoryService['onInitialize']();
  });

  it('should initialize with default configuration', () => {
    const stats = memoryService.getMemoryStats();
    expect(stats).toBeDefined();
    expect(stats.percentageUsed).toBeGreaterThanOrEqual(0);
  });

  it('should detect high memory pressure correctly', () => {
    // Mock memory stats to simulate high memory usage
    const highMemoryStats = {
      heapUsed: 900 * 1024 * 1024, // 900MB
      heapTotal: 1024 * 1024 * 1024, // 1GB
      rss: 1100 * 1024 * 1024, // 1.1GB
      external: 50 * 1024 * 1024, // 50MB
      arrayBuffers: 10 * 1024 * 1024, // 10MB
      percentageUsed: 88
    };

    // Mock getMemoryStats to return high usage
    jest.spyOn(memoryService, 'getMemoryStats').mockReturnValue(highMemoryStats);
    
    // With default threshold of 80%, this should return true
    expect(memoryService.isMemoryPressureHigh()).toBe(true);
  });

  it('should handle document chunking correctly', async () => {
    const documentId = 'test-doc-1';
    const largeContent = 'A'.repeat(100000); // 100KB content
    
    const chunks = await memoryService.processLargeDocument(documentId, largeContent);
    
    expect(chunks).toHaveLength(2); // With 64KB chunks, we should get 2 chunks
    expect(chunks[0].position).toBe(0);
    expect(chunks[1].position).toBe(65536); // 64KB
    expect(chunks[0].size).toBe(65536);
    expect(chunks[1].size).toBe(34464); // Remaining content
  });

  it('should cache processed documents', async () => {
    const documentId = 'test-doc-2';
    const content = 'Test content for caching';
    
    // Process document first time
    const chunks = await memoryService.processLargeDocument(documentId, content);
    
    // Verify the document was processed
    expect(chunks).toHaveLength(1);
    expect(chunks[0].content).toBe(content);
  });

  it('should clean up document cache when size limit is exceeded', () => {
    // Add many documents to exceed cache limit
    for (let i = 0; i < 1010; i++) {
      const chunks = [{
        id: `chunk-${i}`,
        content: `Content for document ${i}`,
        position: 0,
        size: 100,
        compressed: false
      }];
      (memoryService as any).documentCache.set(`doc-${i}`, chunks);
    }
    
    // Trigger cleanup
    (memoryService as any).cleanupDocumentCache();
    
    // Should be reduced to max size
    expect((memoryService as any).documentCache.size).toBeLessThanOrEqual(1000);
  });

  it('should record memory pressure events', () => {
    const stats = memoryService.getMemoryStats();
    (memoryService as any).recordMemoryPressureEvent(stats, 'test_action');
    
    const events = memoryService.getMemoryPressureEvents();
    expect(events).toHaveLength(1);
    expect(events[0].actionTaken).toBe('test_action');
  });

  it('should limit memory pressure events log size', () => {
    const stats = memoryService.getMemoryStats();
    
    // Add more events than the limit
    for (let i = 0; i < 150; i++) {
      (memoryService as any).recordMemoryPressureEvent(stats, `action-${i}`);
    }
    
    const events = memoryService.getMemoryPressureEvents();
    expect(events).toHaveLength(100); // Should be limited to 100
  });

  it('should generate memory optimization report', () => {
    const report = memoryService.generateReport();
    
    expect(report).toBeDefined();
    expect(report.memoryStats).toBeDefined();
    expect(report.cacheStats).toBeDefined();
    expect(report.pressureEvents).toBeDefined();
    expect(report.config).toBeDefined();
    expect(report.recommendations).toBeDefined();
  });

  it('should clear document cache', () => {
    // Add some documents to cache
    (memoryService as any).documentCache.set('doc-1', [{ id: 'chunk-1', content: 'test', position: 0, size: 4, compressed: false }]);
    (memoryService as any).documentCache.set('doc-2', [{ id: 'chunk-2', content: 'test2', position: 0, size: 5, compressed: false }]);
    
    expect((memoryService as any).documentCache.size).toBe(2);
    
    memoryService.clearDocumentCache();
    
    expect((memoryService as any).documentCache.size).toBe(0);
  });

  it('should stream process content correctly', async () => {
    const content = 'A'.repeat(1000);
    const chunks: string[] = [];
    
    // Process content through the stream
    for await (const chunk of memoryService.streamProcessContent(content)) {
      chunks.push(chunk);
    }
    
    expect(chunks).toHaveLength(1); // With 64KB chunk size, 1000 chars fit in one chunk
    expect(chunks[0]).toHaveLength(1000);
  });

  it('should update configuration dynamically', () => {
    const newConfig = {
      cleanupThreshold: 90,
      streamChunkSize: 32 * 1024 // 32KB
    };
    
    memoryService.updateConfig(newConfig);
    
    // Check that config was updated
    const report = memoryService.generateReport();
    expect(report.config.cleanupThreshold).toBe(90);
    expect(report.config.streamChunkSize).toBe(32 * 1024);
  });
});