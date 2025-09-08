/**
 * Unit tests for LoadBalancingService functionality
 */

import { describe, it, expect, beforeEach, afterEach, jest } from '@jest/globals';
import { LoadBalancingService } from '../../../services/LoadBalancingService';
import { createMockConfig } from '../../utils/testHelpers';

describe('LoadBalancingService', () => {
  let loadBalancingService: LoadBalancingService;
  let mockConfig: any;

  beforeEach(() => {
    mockConfig = createMockConfig();
    loadBalancingService = new LoadBalancingService(mockConfig);
  });

  afterEach(async () => {
    // Clean up any intervals
    await loadBalancingService.shutdown();
  });

  it('should initialize with default configuration', () => {
    const stats = loadBalancingService.getStats();
    
    expect(stats).toBeDefined();
    expect(stats.totalNodes).toBe(0);
    expect(stats.activeNodes).toBe(0);
    expect(stats.strategy).toBe('round-robin');
  });

  it('should initialize with server nodes', () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 },
      { id: 'node-2', host: '192.168.1.11', port: 3000 },
      { id: 'node-3', host: '192.168.1.12', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    const stats = loadBalancingService.getStats();
    expect(stats.totalNodes).toBe(3);
    expect(stats.activeNodes).toBe(3);
  });

  it('should select nodes using round-robin strategy', () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 },
      { id: 'node-2', host: '192.168.1.11', port: 3000 },
      { id: 'node-3', host: '192.168.1.12', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    // Select nodes multiple times
    const selectedNodes = [];
    for (let i = 0; i < 6; i++) {
      const node = loadBalancingService.getNextNode();
      if (node) {
        selectedNodes.push(node.id);
      }
    }
    
    // Should cycle through nodes in order
    expect(selectedNodes).toEqual([
      'node-1', 'node-2', 'node-3',
      'node-1', 'node-2', 'node-3'
    ]);
  });

  it('should select nodes using least-connections strategy', () => {
    // Reinitialize with least-connections strategy
    const config = createMockConfig();
    config.loadBalancer = { strategy: 'least-connections' };
    const service = new LoadBalancingService(config);
    
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 },
      { id: 'node-2', host: '192.168.1.11', port: 3000 },
      { id: 'node-3', host: '192.168.1.12', port: 3000 }
    ];
    
    service.initializeNodes(nodes);
    
    // Manually set connection counts
    const node1 = service.getNode('node-1');
    const node2 = service.getNode('node-2');
    const node3 = service.getNode('node-3');
    
    if (node1) node1.currentConnections = 5;
    if (node2) node2.currentConnections = 2;
    if (node3) node3.currentConnections = 8;
    
    // Should select node with least connections (node-2)
    const selectedNode = service.getNextNode();
    expect(selectedNode?.id).toBe('node-2');
  });

  it('should select nodes using weighted-round-robin strategy', () => {
    // Reinitialize with weighted-round-robin strategy
    const config = createMockConfig();
    config.loadBalancer = { strategy: 'weighted-round-robin' };
    const service = new LoadBalancingService(config);
    
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000, weight: 3 },
      { id: 'node-2', host: '192.168.1.11', port: 3000, weight: 1 },
      { id: 'node-3', host: '192.168.1.12', port: 3000, weight: 2 }
    ];
    
    service.initializeNodes(nodes);
    
    // Should select node with highest weight
    const selectedNode = service.getNextNode();
    expect(selectedNode?.id).toBe('node-1');
  });

  it('should select nodes using IP hash strategy', () => {
    // Reinitialize with IP hash strategy
    const config = createMockConfig();
    config.loadBalancer = { strategy: 'ip-hash' };
    const service = new LoadBalancingService(config);
    
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 },
      { id: 'node-2', host: '192.168.1.11', port: 3000 },
      { id: 'node-3', host: '192.168.1.12', port: 3000 }
    ];
    
    service.initializeNodes(nodes);
    
    // Should select the same node for the same IP
    const node1 = service.getNextNode('192.168.1.100');
    const node2 = service.getNextNode('192.168.1.100');
    
    expect(node1?.id).toBe(node2?.id);
  });

  it('should handle node connection management', () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    // Select a node
    const selectedNode = loadBalancingService.getNextNode();
    expect(selectedNode?.id).toBe('node-1');
    expect(selectedNode?.currentConnections).toBe(1);
    
    // Release the connection
    loadBalancingService.releaseNode('node-1');
    
    const updatedNode = loadBalancingService.getNode('node-1');
    expect(updatedNode?.currentConnections).toBe(0);
  });

  it('should manage node active status', () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    // Deactivate node
    const deactivated = loadBalancingService.setNodeActive('node-1', false);
    expect(deactivated).toBe(true);
    
    // Try to select node (should fail since it's inactive)
    const selectedNode = loadBalancingService.getNextNode();
    expect(selectedNode).toBeNull();
    
    // Reactivate node
    const activated = loadBalancingService.setNodeActive('node-1', true);
    expect(activated).toBe(true);
    
    // Try to select node again (should succeed)
    const selectedNode2 = loadBalancingService.getNextNode();
    expect(selectedNode2?.id).toBe('node-1');
  });

  it('should add and remove nodes', () => {
    // Initially no nodes
    expect(loadBalancingService.getStats().totalNodes).toBe(0);
    
    // Add nodes
    loadBalancingService.addNode({ id: 'node-1', host: '192.168.1.10', port: 3000 });
    loadBalancingService.addNode({ id: 'node-2', host: '192.168.1.11', port: 3000 });
    
    expect(loadBalancingService.getStats().totalNodes).toBe(2);
    
    // Remove a node
    const removed = loadBalancingService.removeNode('node-1');
    expect(removed).toBe(true);
    expect(loadBalancingService.getStats().totalNodes).toBe(1);
    
    // Try to remove non-existent node
    const notRemoved = loadBalancingService.removeNode('node-3');
    expect(notRemoved).toBe(false);
  });

  it('should perform health checks and update node status', async () => {
    // Mock the health check to always succeed
    const checkNodeHealthSpy = jest.spyOn(loadBalancingService as any, 'checkNodeHealth')
      .mockResolvedValue(true);
    
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    // Perform health checks
    await (loadBalancingService as any).performHealthChecks();
    
    const node = loadBalancingService.getNode('node-1');
    expect(node?.health).toBe('healthy');
    expect(node?.successCount).toBeGreaterThan(0);
    
    checkNodeHealthSpy.mockRestore();
  });

  it('should handle unhealthy nodes', async () => {
    // Mock the health check to always fail
    const checkNodeHealthSpy = jest.spyOn(loadBalancingService as any, 'checkNodeHealth')
      .mockResolvedValue(false);
    
    // Set failure threshold to 1 for testing
    (loadBalancingService as any).loadBalancerConfig.failureThreshold = 1;
    
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    // Perform health checks
    await (loadBalancingService as any).performHealthChecks();
    
    const node = loadBalancingService.getNode('node-1');
    expect(node?.health).toBe('unhealthy');
    expect(node?.failureCount).toBe(1);
    
    checkNodeHealthSpy.mockRestore();
  });

  it('should generate load balancing reports', () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    const report = loadBalancingService.generateReport();
    
    expect(report).toBeDefined();
    expect(report.nodes).toHaveLength(1);
    expect(report.stats).toBeDefined();
    expect(report.config).toBeDefined();
    expect(report.recommendations).toBeDefined();
  });

  it('should export and import configuration', () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000, weight: 2 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    const exportedConfig = loadBalancingService.exportConfig();
    expect(typeof exportedConfig).toBe('string');
    expect(exportedConfig).toContain('node-1');
    
    // Create a new service and import the configuration
    const newService = new LoadBalancingService(createMockConfig());
    newService.importConfig(exportedConfig);
    
    const importedNodes = newService.getNodes();
    expect(importedNodes).toHaveLength(1);
    expect(importedNodes[0].id).toBe('node-1');
    expect(importedNodes[0].weight).toBe(2);
  });

  it('should rebalance nodes', () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    const beforeRebalance = loadBalancingService.getStats().lastRebalanced;
    loadBalancingService.rebalance();
    const afterRebalance = loadBalancingService.getStats().lastRebalanced;
    
    expect(afterRebalance).not.toEqual(beforeRebalance);
  });

  it('should handle shutdown gracefully', async () => {
    const nodes = [
      { id: 'node-1', host: '192.168.1.10', port: 3000 }
    ];
    
    loadBalancingService.initializeNodes(nodes);
    
    // Should not throw an error
    await expect(loadBalancingService.shutdown()).resolves.toBeUndefined();
  });

  it('should return null when no healthy nodes are available', () => {
    // Don't initialize any nodes
    const node = loadBalancingService.getNextNode();
    expect(node).toBeNull();
  });
});