import { BaseService } from '@/core/BaseService';
import type { BotConfig } from '@/types';
import logger from '@/utils/logger';

export interface LoadBalancerConfig {
  strategy: 'round-robin' | 'least-connections' | 'weighted-round-robin' | 'ip-hash';
  healthCheckInterval: number; // in milliseconds
  failureThreshold: number; // number of failures before marking node as unhealthy
  recoveryThreshold: number; // number of successful health checks before marking node as healthy
  timeout: number; // in milliseconds
}

export interface ServerNode {
  id: string;
  host: string;
  port: number;
  weight?: number; // For weighted strategies
  active: boolean;
  currentConnections: number;
  health: 'healthy' | 'unhealthy' | 'recovering';
  lastHealthCheck: Date;
  failureCount: number;
  successCount: number;
}

export interface LoadBalancingStats {
  totalNodes: number;
  activeNodes: number;
  unhealthyNodes: number;
  totalConnections: number;
  averageConnections: number;
  strategy: string;
  lastRebalanced: Date | null;
}

export class LoadBalancingService extends BaseService {
  private nodes: ServerNode[] = [];
  private config: LoadBalancerConfig;
  private currentIndex: number = 0;
  private stats: LoadBalancingStats;
  private healthCheckIntervalId: NodeJS.Timeout | null = null;
  private lastRebalanced: Date | null = null;
  
  constructor(config: BotConfig) {
    super('LoadBalancingService', config);
    
    this.config = {
      strategy: config.loadBalancer?.strategy || 'round-robin',
      healthCheckInterval: config.loadBalancer?.healthCheckInterval || 30000, // 30 seconds
      failureThreshold: config.loadBalancer?.failureThreshold || 3,
      recoveryThreshold: config.loadBalancer?.recoveryThreshold || 2,
      timeout: config.loadBalancer?.timeout || 5000 // 5 seconds
    };
    
    this.stats = {
      totalNodes: 0,
      activeNodes: 0,
      unhealthyNodes: 0,
      totalConnections: 0,
      averageConnections: 0,
      strategy: this.config.strategy,
      lastRebalanced: null
    };
  }

  /**
   * Initialize the load balancer with server nodes
   */
  initializeNodes(nodes: Omit<ServerNode, 'active' | 'currentConnections' | 'health' | 'lastHealthCheck' | 'failureCount' | 'successCount'>[]): void {
    this.nodes = nodes.map(node => ({
      ...node,
      active: true,
      currentConnections: 0,
      health: 'healthy',
      lastHealthCheck: new Date(),
      failureCount: 0,
      successCount: 0
    }));
    
    this.updateStats();
    
    // Start health checks
    this.startHealthChecks();
    
    logger.info('Load balancer initialized', {
      component: 'LoadBalancingService',
      nodeCount: this.nodes.length,
      strategy: this.config.strategy
    });
  }

  /**
   * Get the next available server node based on the load balancing strategy
   */
  getNextNode(clientIp?: string): ServerNode | null {
    // Filter out unhealthy and inactive nodes
    const availableNodes = this.nodes.filter(node => 
      node.active && node.health === 'healthy'
    );
    
    if (availableNodes.length === 0) {
      logger.warn('No healthy nodes available for load balancing', {
        component: 'LoadBalancingService'
      });
      return null;
    }
    
    let selectedNode: ServerNode;
    
    switch (this.config.strategy) {
      case 'round-robin':
        selectedNode = this.roundRobin(availableNodes);
        break;
        
      case 'least-connections':
        selectedNode = this.leastConnections(availableNodes);
        break;
        
      case 'weighted-round-robin':
        selectedNode = this.weightedRoundRobin(availableNodes);
        break;
        
      case 'ip-hash':
        selectedNode = this.ipHash(availableNodes, clientIp);
        break;
        
      default:
        // Default to round-robin
        selectedNode = this.roundRobin(availableNodes);
        break;
    }
    
    // Increment connection count for the selected node
    selectedNode.currentConnections++;
    this.updateStats();
    
    logger.debug('Node selected for load balancing', {
      component: 'LoadBalancingService',
      nodeId: selectedNode.id,
      host: selectedNode.host,
      port: selectedNode.port,
      strategy: this.config.strategy
    });
    
    return selectedNode;
  }

  /**
   * Round-robin load balancing strategy
   */
  private roundRobin(nodes: ServerNode[]): ServerNode {
    const node = nodes[this.currentIndex];
    this.currentIndex = (this.currentIndex + 1) % nodes.length;
    return node;
  }

  /**
   * Least connections load balancing strategy
   */
  private leastConnections(nodes: ServerNode[]): ServerNode {
    return nodes.reduce((min, node) => 
      node.currentConnections < min.currentConnections ? node : min
    );
  }

  /**
   * Weighted round-robin load balancing strategy
   */
  private weightedRoundRobin(nodes: ServerNode[]): ServerNode {
    // If no weights are specified, use regular round-robin
    if (nodes.every(node => node.weight === undefined)) {
      return this.roundRobin(nodes);
    }
    
    // Find the node with the highest weight that still has capacity
    let maxWeight = -1;
    let selectedNode: ServerNode | null = null;
    
    for (const node of nodes) {
      const weight = node.weight || 1;
      if (weight > maxWeight) {
        maxWeight = weight;
        selectedNode = node;
      }
    }
    
    return selectedNode || nodes[0];
  }

  /**
   * IP hash load balancing strategy
   */
  private ipHash(nodes: ServerNode[], clientIp?: string): ServerNode {
    if (!clientIp) {
      // If no client IP, fall back to round-robin
      return this.roundRobin(nodes);
    }
    
    // Create a hash of the client IP
    let hash = 0;
    for (let i = 0; i < clientIp.length; i++) {
      hash = ((hash << 5) - hash) + clientIp.charCodeAt(i);
      hash = hash & hash; // Convert to 32-bit integer
    }
    
    // Use the hash to select a node
    const index = Math.abs(hash) % nodes.length;
    return nodes[index];
  }

  /**
   * Release a connection from a node
   */
  releaseNode(nodeId: string): void {
    const node = this.nodes.find(n => n.id === nodeId);
    
    if (node && node.currentConnections > 0) {
      node.currentConnections--;
      this.updateStats();
      
      logger.debug('Connection released from node', {
        component: 'LoadBalancingService',
        nodeId,
        currentConnections: node.currentConnections
      });
    }
  }

  /**
   * Mark a node as active or inactive
   */
  setNodeActive(nodeId: string, active: boolean): boolean {
    const node = this.nodes.find(n => n.id === nodeId);
    
    if (node) {
      node.active = active;
      this.updateStats();
      
      logger.info('Node active status updated', {
        component: 'LoadBalancingService',
        nodeId,
        active
      });
      
      return true;
    }
    
    return false;
  }

  /**
   * Add a new server node
   */
  addNode(node: Omit<ServerNode, 'active' | 'currentConnections' | 'health' | 'lastHealthCheck' | 'failureCount' | 'successCount'>): void {
    this.nodes.push({
      ...node,
      active: true,
      currentConnections: 0,
      health: 'healthy',
      lastHealthCheck: new Date(),
      failureCount: 0,
      successCount: 0
    });
    
    this.updateStats();
    
    logger.info('Node added to load balancer', {
      component: 'LoadBalancingService',
      nodeId: node.id,
      host: node.host,
      port: node.port
    });
  }

  /**
   * Remove a server node
   */
  removeNode(nodeId: string): boolean {
    const initialLength = this.nodes.length;
    this.nodes = this.nodes.filter(node => node.id !== nodeId);
    
    const removed = this.nodes.length < initialLength;
    
    if (removed) {
      this.updateStats();
      
      logger.info('Node removed from load balancer', {
        component: 'LoadBalancingService',
        nodeId
      });
    }
    
    return removed;
  }

  /**
   * Start health checks for all nodes
   */
  private startHealthChecks(): void {
    if (this.healthCheckIntervalId) {
      clearInterval(this.healthCheckIntervalId);
    }
    
    this.healthCheckIntervalId = setInterval(() => {
      this.performHealthChecks();
    }, this.config.healthCheckInterval);
    
    logger.info('Health checks started', {
      component: 'LoadBalancingService',
      interval: this.config.healthCheckInterval
    });
  }

  /**
   * Perform health checks on all nodes
   */
  private async performHealthChecks(): Promise<void> {
    logger.debug('Performing health checks on nodes', {
      component: 'LoadBalancingService',
      nodeCount: this.nodes.length
    });
    
    for (const node of this.nodes) {
      try {
        const isHealthy = await this.checkNodeHealth(node);
        
        if (isHealthy) {
          this.handleHealthyNode(node);
        } else {
          this.handleUnhealthyNode(node);
        }
        
        node.lastHealthCheck = new Date();
      } catch (error) {
        logger.warn('Error during health check', {
          component: 'LoadBalancingService',
          nodeId: node.id,
          error: error instanceof Error ? error.message : String(error)
        });
        
        this.handleUnhealthyNode(node);
      }
    }
    
    this.updateStats();
  }

  /**
   * Check the health of a specific node
   */
  private async checkNodeHealth(node: ServerNode): Promise<boolean> {
    // In a real implementation, this would make an actual HTTP request to the node
    // For now, we'll simulate a health check with a random success/failure
    
    // Simulate network latency
    await new Promise(resolve => setTimeout(resolve, Math.random() * 100));
    
    // Simulate 90% success rate
    return Math.random() > 0.1;
  }

  /**
   * Handle a healthy node
   */
  private handleHealthyNode(node: ServerNode): void {
    node.successCount++;
    node.failureCount = 0; // Reset failure count
    
    if (node.health === 'unhealthy') {
      // Node was unhealthy, check if it should be marked as recovering
      if (node.successCount >= this.config.recoveryThreshold) {
        node.health = 'healthy';
        logger.info('Node marked as healthy', {
          component: 'LoadBalancingService',
          nodeId: node.id
        });
      } else {
        node.health = 'recovering';
        logger.info('Node marked as recovering', {
          component: 'LoadBalancingService',
          nodeId: node.id,
          successCount: node.successCount
        });
      }
    } else if (node.health === 'recovering') {
      // Node was recovering, check if it should be marked as healthy
      if (node.successCount >= this.config.recoveryThreshold) {
        node.health = 'healthy';
        logger.info('Node recovery completed', {
          component: 'LoadBalancingService',
          nodeId: node.id
        });
      }
    }
  }

  /**
   * Handle an unhealthy node
   */
  private handleUnhealthyNode(node: ServerNode): void {
    node.failureCount++;
    node.successCount = 0; // Reset success count
    
    if (node.health === 'healthy' || node.health === 'recovering') {
      // Node was healthy, check if it should be marked as unhealthy
      if (node.failureCount >= this.config.failureThreshold) {
        node.health = 'unhealthy';
        logger.warn('Node marked as unhealthy', {
          component: 'LoadBalancingService',
          nodeId: node.id
        });
      } else if (node.health === 'healthy') {
        node.health = 'recovering';
        logger.warn('Node marked as recovering due to failures', {
          component: 'LoadBalancingService',
          nodeId: node.id,
          failureCount: node.failureCount
        });
      }
    }
  }

  /**
   * Update load balancing statistics
   */
  private updateStats(): void {
    const activeNodes = this.nodes.filter(node => node.active);
    const unhealthyNodes = this.nodes.filter(node => node.health === 'unhealthy');
    const totalConnections = this.nodes.reduce((sum, node) => sum + node.currentConnections, 0);
    const averageConnections = this.nodes.length > 0 ? totalConnections / this.nodes.length : 0;
    
    this.stats = {
      totalNodes: this.nodes.length,
      activeNodes: activeNodes.length,
      unhealthyNodes: unhealthyNodes.length,
      totalConnections,
      averageConnections,
      strategy: this.config.strategy,
      lastRebalanced: this.lastRebalanced
    };
  }

  /**
   * Get current load balancing statistics
   */
  getStats(): LoadBalancingStats {
    return { ...this.stats };
  }

  /**
   * Get all nodes
   */
  getNodes(): ServerNode[] {
    return [...this.nodes];
  }

  /**
   * Get a specific node by ID
   */
  getNode(nodeId: string): ServerNode | null {
    const node = this.nodes.find(n => n.id === nodeId);
    return node ? { ...node } : null;
  }

  /**
   * Rebalance nodes based on current load
   */
  rebalance(): void {
    // For now, we'll just update the last rebalanced time
    // In a more complex implementation, this could redistribute connections
    this.lastRebalanced = new Date();
    this.updateStats();
    
    logger.info('Load balancer rebalanced', {
      component: 'LoadBalancingService'
    });
  }

  /**
   * Generate a load balancing report
   */
  generateReport(): {
    nodes: ServerNode[];
    stats: LoadBalancingStats;
    config: LoadBalancerConfig;
    recommendations: string[];
  } {
    const recommendations: string[] = [];
    
    // Generate recommendations based on current state
    if (this.stats.unhealthyNodes > 0) {
      recommendations.push(`Remove or repair ${this.stats.unhealthyNodes} unhealthy nodes`);
    }
    
    if (this.stats.activeNodes === 0) {
      recommendations.push('No active nodes available - add nodes to the load balancer');
    }
    
    const connectionImbalance = this.checkConnectionImbalance();
    if (connectionImbalance.needed) {
      recommendations.push(`Rebalance connections - max difference is ${connectionImbalance.maxDifference}`);
    }
    
    return {
      nodes: this.getNodes(),
      stats: this.getStats(),
      config: { ...this.config },
      recommendations
    };
  }

  /**
   * Check if connections are imbalanced across nodes
   */
  private checkConnectionImbalance(): { needed: boolean; maxDifference: number } {
    if (this.nodes.length < 2) {
      return { needed: false, maxDifference: 0 };
    }
    
    const connections = this.nodes.map(node => node.currentConnections);
    const max = Math.max(...connections);
    const min = Math.min(...connections);
    const difference = max - min;
    
    // If difference is more than 10% of average connections, rebalancing might be needed
    const average = this.stats.averageConnections;
    const threshold = average * 0.1;
    
    return {
      needed: difference > threshold,
      maxDifference: difference
    };
  }

  /**
   * Shutdown the load balancer
   */
  async shutdown(): Promise<void> {
    if (this.healthCheckIntervalId) {
      clearInterval(this.healthCheckIntervalId);
      this.healthCheckIntervalId = null;
    }
    
    logger.info('Load balancer shutdown', {
      component: 'LoadBalancingService'
    });
  }

  /**
   * Export load balancing configuration
   */
  exportConfig(): string {
    return JSON.stringify({
      config: this.config,
      nodes: this.nodes.map(node => ({
        id: node.id,
        host: node.host,
        port: node.port,
        weight: node.weight
      }))
    }, null, 2);
  }

  /**
   * Import load balancing configuration
   */
  importConfig(configData: string): void {
    try {
      const parsed = JSON.parse(configData);
      
      // Update configuration
      if (parsed.config) {
        Object.assign(this.config, parsed.config);
      }
      
      // Update nodes
      if (parsed.nodes && Array.isArray(parsed.nodes)) {
        this.nodes = parsed.nodes.map((node: any) => ({
          ...node,
          active: true,
          currentConnections: 0,
          health: 'healthy',
          lastHealthCheck: new Date(),
          failureCount: 0,
          successCount: 0
        }));
      }
      
      this.updateStats();
      
      logger.info('Load balancer configuration imported', {
        component: 'LoadBalancingService'
      });
    } catch (error) {
      logger.error('Error importing load balancer configuration', {
        component: 'LoadBalancingService',
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }
}