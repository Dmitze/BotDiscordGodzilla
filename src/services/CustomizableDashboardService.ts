import { BaseService } from '@/core/BaseService';
import type { BotConfig, HealthStatus, ServiceStats } from '@/types';
import logger from '@/utils/logger';

export interface DashboardWidget {
  id: string;
  type: 'document-list' | 'analytics-chart' | 'recent-activity' | 'quick-actions' | 'search' | 'custom';
  title: string;
  position: { x: number; y: number; width: number; height: number };
  config: Record<string, any>;
  visible: boolean;
}

export interface UserDashboard {
  userId: string;
  widgets: DashboardWidget[];
  layout: 'grid' | 'flex' | 'custom';
  theme: 'light' | 'dark' | 'auto';
  updatedAt: Date;
}

export interface DashboardPreferences {
  defaultLayout?: 'grid' | 'flex' | 'custom';
  defaultTheme?: 'light' | 'dark' | 'auto';
  showWelcomeMessage?: boolean;
  enableAutoRefresh?: boolean;
  refreshInterval?: number; // in seconds
}

export class CustomizableDashboardService extends BaseService {
  private userDashboards: Map<string, UserDashboard> = new Map();
  private defaultWidgets: DashboardWidget[] = [
    {
      id: 'recent-docs',
      type: 'document-list',
      title: 'Останні документи',
      position: { x: 0, y: 0, width: 6, height: 4 },
      config: { limit: 10, showPreview: true },
      visible: true,
    },
    {
      id: 'quick-search',
      type: 'search',
      title: 'Швидкий пошук',
      position: { x: 6, y: 0, width: 6, height: 2 },
      config: {},
      visible: true,
    },
    {
      id: 'analytics',
      type: 'analytics-chart',
      title: 'Аналітика',
      position: { x: 6, y: 2, width: 6, height: 4 },
      config: { chartType: 'bar', period: 'week' },
      visible: true,
    },
    {
      id: 'quick-actions',
      type: 'quick-actions',
      title: 'Швидкі дії',
      position: { x: 0, y: 4, width: 12, height: 2 },
      config: {},
      visible: true,
    },
  ];
  private readonly MAX_WIDGETS_PER_DASHBOARD = 20;
  private readonly MAX_DASHBOARDS = 1000;

  constructor(config: BotConfig) {
    super('CustomizableDashboardService', config);
  }

  /**
   * Initialize service
   */
  protected async onInitialize(): Promise<void> {
    // Implementation for initialization if needed
    logger.info('CustomizableDashboardService initialized', {
      component: 'CustomizableDashboardService'
    });
  }

  /**
   * Shutdown service
   */
  protected async onShutdown(): Promise<void> {
    // Implementation for shutdown if needed
    logger.info('CustomizableDashboardService shutdown', {
      component: 'CustomizableDashboardService'
    });
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
    return {
      healthy: true,
      service: 'CustomizableDashboardService'
    };
  }

  /**
   * Get service stats
   */
  protected onGetStats(): Partial<ServiceStats> {
    return {
      userDashboards: this.userDashboards.size
    };
  }

  /**
   * Get service statistics
   */
  public getStats(): ServiceStats {
    // Get base stats from parent class
    const baseStats = super.getStats();
    
    return {
      ...baseStats,
      userDashboards: this.userDashboards.size
    };
  }

  /**
   * Get user dashboard
   */
  getUserDashboard(userId: string): UserDashboard {
    // Check if user already has a dashboard
    let dashboard = this.userDashboards.get(userId);
    
    if (!dashboard) {
      // Create default dashboard for user
      dashboard = this.createDefaultDashboard(userId);
      this.userDashboards.set(userId, dashboard);
      
      // Limit the number of stored dashboards
      if (this.userDashboards.size > this.MAX_DASHBOARDS) {
        const firstKey = this.userDashboards.keys().next().value;
        if (firstKey) {
          this.userDashboards.delete(firstKey);
        }
      }
    }
    
    return dashboard;
  }

  /**
   * Create default dashboard for a user
   */
  private createDefaultDashboard(userId: string): UserDashboard {
    return {
      userId,
      widgets: [...this.defaultWidgets],
      layout: 'grid',
      theme: 'auto',
      updatedAt: new Date(),
    };
  }

  /**
   * Update user dashboard
   */
  updateUserDashboard(userId: string, dashboard: Partial<UserDashboard>): UserDashboard {
    const existingDashboard = this.getUserDashboard(userId);
    
    // Merge updates
    const updatedDashboard: UserDashboard = {
      ...existingDashboard,
      ...dashboard,
      updatedAt: new Date(),
    };
    
    // Validate widget count
    if (updatedDashboard.widgets.length > this.MAX_WIDGETS_PER_DASHBOARD) {
      throw new Error(`Maximum ${this.MAX_WIDGETS_PER_DASHBOARD} widgets allowed per dashboard`);
    }
    
    this.userDashboards.set(userId, updatedDashboard);
    
    logger.debug('User dashboard updated', {
      component: 'CustomizableDashboardService',
      userId,
    });
    
    return updatedDashboard;
  }

  /**
   * Add widget to user dashboard
   */
  addWidget(userId: string, widget: Omit<DashboardWidget, 'id'>): DashboardWidget {
    const dashboard = this.getUserDashboard(userId);
    
    // Generate unique ID for widget
    const widgetId = `widget-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    
    const newWidget: DashboardWidget = {
      id: widgetId,
      ...widget,
    };
    
    // Check widget limit
    if (dashboard.widgets.length >= this.MAX_WIDGETS_PER_DASHBOARD) {
      throw new Error(`Maximum ${this.MAX_WIDGETS_PER_DASHBOARD} widgets allowed per dashboard`);
    }
    
    // Add widget to dashboard
    dashboard.widgets.push(newWidget);
    dashboard.updatedAt = new Date();
    
    this.userDashboards.set(userId, dashboard);
    
    logger.debug('Widget added to user dashboard', {
      component: 'CustomizableDashboardService',
      userId,
      widgetId,
    });
    
    return newWidget;
  }

  /**
   * Update widget configuration
   */
  updateWidget(userId: string, widgetId: string, updates: Partial<DashboardWidget>): DashboardWidget | null {
    const dashboard = this.getUserDashboard(userId);
    
    const widgetIndex = dashboard.widgets.findIndex(w => w.id === widgetId);
    if (widgetIndex === -1) {
      return null;
    }
    
    // Get the existing widget with non-null assertion since we know it exists
    const existingWidget = dashboard.widgets[widgetIndex]!;
    
    // Create updated widget with proper handling of optional properties
    const updatedWidget: DashboardWidget = {
      id: existingWidget.id,
      type: updates.type !== undefined ? updates.type : existingWidget.type,
      title: updates.title !== undefined ? updates.title : existingWidget.title,
      position: updates.position !== undefined ? updates.position : existingWidget.position,
      config: updates.config !== undefined ? updates.config : existingWidget.config,
      visible: updates.visible !== undefined ? updates.visible : existingWidget.visible
    };
    
    // Update widget
    dashboard.widgets[widgetIndex] = updatedWidget;
    
    dashboard.updatedAt = new Date();
    this.userDashboards.set(userId, dashboard);
    
    logger.debug('Widget updated', {
      component: 'CustomizableDashboardService',
      userId,
      widgetId,
    });
    
    return updatedWidget;
  }

  /**
   * Remove widget from user dashboard
   */
  removeWidget(userId: string, widgetId: string): boolean {
    const dashboard = this.getUserDashboard(userId);
    
    const initialLength = dashboard.widgets.length;
    dashboard.widgets = dashboard.widgets.filter(w => w.id !== widgetId);
    
    const removed = dashboard.widgets.length < initialLength;
    
    if (removed) {
      dashboard.updatedAt = new Date();
      this.userDashboards.set(userId, dashboard);
      
      logger.debug('Widget removed from user dashboard', {
        component: 'CustomizableDashboardService',
        userId,
        widgetId,
      });
    }
    
    return removed;
  }

  /**
   * Reset user dashboard to default
   */
  resetToDefault(userId: string): UserDashboard {
    const defaultDashboard = this.createDefaultDashboard(userId);
    this.userDashboards.set(userId, defaultDashboard);
    
    logger.debug('User dashboard reset to default', {
      component: 'CustomizableDashboardService',
      userId,
    });
    
    return defaultDashboard;
  }

  /**
   * Get dashboard statistics
   */
  getDashboardStats(): {
    totalDashboards: number;
    averageWidgetsPerDashboard: number;
    mostPopularWidgetTypes: { type: string; count: number }[];
  } {
    const dashboards = Array.from(this.userDashboards.values());
    const totalWidgets = dashboards.reduce((sum, dashboard) => sum + dashboard.widgets.length, 0);
    const averageWidgets = dashboards.length > 0 ? totalWidgets / dashboards.length : 0;
    
    // Calculate most popular widget types
    const widgetTypeCount: Record<string, number> = {};
    for (const dashboard of dashboards) {
      for (const widget of dashboard.widgets) {
        widgetTypeCount[widget.type] = (widgetTypeCount[widget.type] || 0) + 1;
      }
    }
    
    const mostPopularWidgetTypes = Object.entries(widgetTypeCount)
      .map(([type, count]) => ({ type, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);
    
    return {
      totalDashboards: dashboards.length,
      averageWidgetsPerDashboard: parseFloat(averageWidgets.toFixed(1)),
      mostPopularWidgetTypes,
    };
  }

  /**
   * Export user dashboard configuration
   */
  exportDashboard(userId: string): string {
    const dashboard = this.getUserDashboard(userId);
    return JSON.stringify(dashboard, null, 2);
  }

  /**
   * Import user dashboard configuration
   */
  importDashboard(userId: string, dashboardData: string): UserDashboard {
    try {
      const parsedDashboard = JSON.parse(dashboardData) as UserDashboard;
      
      // Validate structure
      if (!parsedDashboard.widgets || !Array.isArray(parsedDashboard.widgets)) {
        throw new Error('Invalid dashboard data structure');
      }
      
      // Validate widget count
      if (parsedDashboard.widgets.length > this.MAX_WIDGETS_PER_DASHBOARD) {
        throw new Error(`Dashboard contains too many widgets (max: ${this.MAX_WIDGETS_PER_DASHBOARD})`);
      }
      
      // Set user ID and update timestamp
      parsedDashboard.userId = userId;
      parsedDashboard.updatedAt = new Date();
      
      this.userDashboards.set(userId, parsedDashboard);
      
      logger.debug('Dashboard imported for user', {
        component: 'CustomizableDashboardService',
        userId,
      });
      
      return parsedDashboard;
    } catch (error) {
      logger.error('Failed to import dashboard', {
        component: 'CustomizableDashboardService',
        userId,
        error: error instanceof Error ? error.message : String(error),
      });
      
      throw new Error(`Failed to import dashboard: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
}