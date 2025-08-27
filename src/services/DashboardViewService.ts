/**
 * DashboardViewService
 * - Manages customizable dashboard views for users
 * - Stores user preferences for dashboard layout and file display options
 */

import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';

export interface DashboardViewConfig {
  userId: string;
  viewName: string;
  layout: 'grid' | 'list' | 'compact';
  sortBy: 'name' | 'modifiedTime' | 'size' | 'mimeType';
  sortOrder: 'asc' | 'desc';
  showPreview: boolean;
  showTags: boolean;
  showOwner: boolean;
  showDates: boolean;
  itemsPerPage: number;
  fileFilters: {
    mimeTypes?: string[];
    owners?: string[];
    dateFrom?: string;
    dateTo?: string;
    sizeMin?: number;
    sizeMax?: number;
  };
  customColumns?: string[]; // For additional metadata columns
}

export interface UserDashboardPreferences {
  userId: string;
  defaultView: string;
  views: DashboardViewConfig[];
  lastUsedView?: string;
}

// In-memory store for now; can be swapped for Redis/File later
const userDashboardPrefs = new Map<string, UserDashboardPreferences>();

export class DashboardViewService {
  /**
   * Get user dashboard preferences
   */
  getUserPreferences(userId: string): UserDashboardPreferences {
    const existing = userDashboardPrefs.get(userId);
    if (existing) return existing;
    
    // Create default preferences
    const defaultView: DashboardViewConfig = {
      userId,
      viewName: 'default',
      layout: 'list',
      sortBy: 'modifiedTime',
      sortOrder: 'desc',
      showPreview: true,
      showTags: true,
      showOwner: true,
      showDates: true,
      itemsPerPage: 25,
      fileFilters: {}
    };
    
    const prefs: UserDashboardPreferences = {
      userId,
      defaultView: 'default',
      views: [defaultView]
    };
    
    userDashboardPrefs.set(userId, prefs);
    return prefs;
  }

  /**
   * Set user dashboard preferences
   */
  setUserPreferences(userId: string, preferences: UserDashboardPreferences): void {
    userDashboardPrefs.set(userId, preferences);
    logger.debug('User dashboard preferences updated', {
      component: 'DashboardViewService',
      userId,
      viewCount: preferences.views.length
    });
  }

  /**
   * Get a specific view configuration for a user
   */
  getViewConfig(userId: string, viewName: string): DashboardViewConfig | undefined {
    const prefs = this.getUserPreferences(userId);
    return prefs.views.find(view => view.viewName === viewName);
  }

  /**
   * Create or update a view configuration
   */
  saveViewConfig(userId: string, viewConfig: DashboardViewConfig): void {
    const prefs = this.getUserPreferences(userId);
    
    // Check if view already exists
    const existingIndex = prefs.views.findIndex(view => view.viewName === viewConfig.viewName);
    
    if (existingIndex >= 0) {
      // Update existing view
      prefs.views[existingIndex] = viewConfig;
    } else {
      // Add new view
      prefs.views.push(viewConfig);
    }
    
    // Update default view if this is the default
    if (viewConfig.viewName === prefs.defaultView) {
      prefs.defaultView = viewConfig.viewName;
    }
    
    this.setUserPreferences(userId, prefs);
    logger.debug('Dashboard view config saved', {
      component: 'DashboardViewService',
      userId,
      viewName: viewConfig.viewName
    });
  }

  /**
   * Delete a view configuration
   */
  deleteViewConfig(userId: string, viewName: string): boolean {
    const prefs = this.getUserPreferences(userId);
    
    // Cannot delete the default view if it's the only one
    if (prefs.views.length <= 1 && prefs.defaultView === viewName) {
      logger.warn('Cannot delete the only view', {
        component: 'DashboardViewService',
        userId,
        viewName
      });
      return false;
    }
    
    const initialLength = prefs.views.length;
    prefs.views = prefs.views.filter(view => view.viewName !== viewName);
    
    // If we deleted the default view, set a new default
    if (prefs.defaultView === viewName && prefs.views.length > 0) {
      prefs.defaultView = prefs.views[0].viewName;
    }
    
    this.setUserPreferences(userId, prefs);
    
    const deleted = prefs.views.length < initialLength;
    if (deleted) {
      logger.debug('Dashboard view config deleted', {
        component: 'DashboardViewService',
        userId,
        viewName
      });
    }
    
    return deleted;
  }

  /**
   * Apply view configuration to a list of files
   */
  applyViewConfig(files: DriveFile[], config: DashboardViewConfig): DriveFile[] {
    // Apply sorting
    const sortedFiles = [...files].sort((a, b) => {
      let comparison = 0;
      
      switch (config.sortBy) {
        case 'name':
          comparison = (a.name || '').localeCompare(b.name || '');
          break;
        case 'modifiedTime':
          const dateA = a.modifiedTime ? new Date(a.modifiedTime).getTime() : 0;
          const dateB = b.modifiedTime ? new Date(b.modifiedTime).getTime() : 0;
          comparison = dateA - dateB;
          break;
        case 'size':
          const sizeA = typeof a.size === 'number' ? a.size : 0;
          const sizeB = typeof b.size === 'number' ? b.size : 0;
          comparison = sizeA - sizeB;
          break;
        case 'mimeType':
          comparison = (a.mimeType || '').localeCompare(b.mimeType || '');
          break;
        default:
          comparison = 0;
      }
      
      return config.sortOrder === 'asc' ? comparison : -comparison;
    });
    
    // Apply filters
    let filteredFiles = sortedFiles;
    
    // MIME type filter
    if (config.fileFilters.mimeTypes && config.fileFilters.mimeTypes.length > 0) {
      filteredFiles = filteredFiles.filter(file => 
        config.fileFilters.mimeTypes?.includes(file.mimeType || '')
      );
    }
    
    // Owner filter
    if (config.fileFilters.owners && config.fileFilters.owners.length > 0) {
      filteredFiles = filteredFiles.filter(file => 
        file.owners?.some(owner => config.fileFilters.owners?.includes(owner))
      );
    }
    
    // Date range filter
    if (config.fileFilters.dateFrom || config.fileFilters.dateTo) {
      filteredFiles = filteredFiles.filter(file => {
        if (!file.modifiedTime) return true;
        
        const fileDate = new Date(file.modifiedTime);
        
        if (config.fileFilters.dateFrom) {
          const fromDate = new Date(config.fileFilters.dateFrom);
          if (fileDate < fromDate) return false;
        }
        
        if (config.fileFilters.dateTo) {
          const toDate = new Date(config.fileFilters.dateTo);
          if (fileDate > toDate) return false;
        }
        
        return true;
      });
    }
    
    // Size filter
    if (typeof config.fileFilters.sizeMin === 'number' || typeof config.fileFilters.sizeMax === 'number') {
      filteredFiles = filteredFiles.filter(file => {
        const fileSize = typeof file.size === 'number' ? file.size : 0;
        
        if (typeof config.fileFilters.sizeMin === 'number' && fileSize < config.fileFilters.sizeMin) {
          return false;
        }
        
        if (typeof config.fileFilters.sizeMax === 'number' && fileSize > config.fileFilters.sizeMax) {
          return false;
        }
        
        return true;
      });
    }
    
    return filteredFiles;
  }

  /**
   * Get available view templates
   */
  getViewTemplates(): DashboardViewConfig[] {
    return [
      {
        userId: 'template',
        viewName: 'detailed',
        layout: 'list',
        sortBy: 'modifiedTime',
        sortOrder: 'desc',
        showPreview: true,
        showTags: true,
        showOwner: true,
        showDates: true,
        itemsPerPage: 25,
        fileFilters: {},
        customColumns: ['mimeType', 'size']
      },
      {
        userId: 'template',
        viewName: 'compact',
        layout: 'compact',
        sortBy: 'name',
        sortOrder: 'asc',
        showPreview: false,
        showTags: false,
        showOwner: false,
        showDates: false,
        itemsPerPage: 50,
        fileFilters: {}
      },
      {
        userId: 'template',
        viewName: 'grid',
        layout: 'grid',
        sortBy: 'modifiedTime',
        sortOrder: 'desc',
        showPreview: true,
        showTags: true,
        showOwner: true,
        showDates: true,
        itemsPerPage: 12,
        fileFilters: {}
      }
    ];
  }

  /**
   * For tests: reset in-memory store
   */
  __reset(): void {
    userDashboardPrefs.clear();
  }
}

export const dashboardViewService = new DashboardViewService();