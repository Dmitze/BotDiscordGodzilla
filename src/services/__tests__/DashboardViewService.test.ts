import { DashboardViewService, dashboardViewService } from '../DashboardViewService';
import type { DriveFile } from '@/types/drive';

describe('DashboardViewService', () => {
  const userId = 'test-user-id';
  const testFile: DriveFile = {
    id: 'file1',
    name: 'Test Document',
    mimeType: 'application/pdf',
    modifiedTime: '2023-01-01T10:00:00Z',
    size: 1024,
    owners: ['owner1@example.com']
  } as any;

  beforeEach(() => {
    // Reset the service before each test
    dashboardViewService.__reset();
  });

  test('should create default preferences for new user', () => {
    const prefs = dashboardViewService.getUserPreferences(userId);
    
    expect(prefs.userId).toBe(userId);
    expect(prefs.defaultView).toBe('default');
    expect(prefs.views).toHaveLength(1);
    expect(prefs.views[0].viewName).toBe('default');
  });

  test('should save and retrieve user preferences', () => {
    const prefs = dashboardViewService.getUserPreferences(userId);
    prefs.defaultView = 'detailed';
    
    dashboardViewService.setUserPreferences(userId, prefs);
    
    const updatedPrefs = dashboardViewService.getUserPreferences(userId);
    expect(updatedPrefs.defaultView).toBe('detailed');
  });

  test('should create new view configuration', () => {
    const newView = {
      userId,
      viewName: 'compact-view',
      layout: 'compact' as const,
      sortBy: 'name' as const,
      sortOrder: 'asc' as const,
      showPreview: false,
      showTags: false,
      showOwner: false,
      showDates: false,
      itemsPerPage: 50,
      fileFilters: {}
    };
    
    dashboardViewService.saveViewConfig(userId, newView);
    
    const viewConfig = dashboardViewService.getViewConfig(userId, 'compact-view');
    expect(viewConfig).toEqual(newView);
  });

  test('should update existing view configuration', () => {
    // First create a view
    const initialView = {
      userId,
      viewName: 'test-view',
      layout: 'list' as const,
      sortBy: 'name' as const,
      sortOrder: 'asc' as const,
      showPreview: true,
      showTags: true,
      showOwner: true,
      showDates: true,
      itemsPerPage: 25,
      fileFilters: {}
    };
    
    dashboardViewService.saveViewConfig(userId, initialView);
    
    // Update the view
    const updatedView = {
      ...initialView,
      layout: 'compact' as const,
      itemsPerPage: 100
    };
    
    dashboardViewService.saveViewConfig(userId, updatedView);
    
    const viewConfig = dashboardViewService.getViewConfig(userId, 'test-view');
    expect(viewConfig?.layout).toBe('compact');
    expect(viewConfig?.itemsPerPage).toBe(100);
  });

  test('should delete view configuration', () => {
    // Create multiple views
    const view1 = {
      userId,
      viewName: 'view1',
      layout: 'list' as const,
      sortBy: 'name' as const,
      sortOrder: 'asc' as const,
      showPreview: true,
      showTags: true,
      showOwner: true,
      showDates: true,
      itemsPerPage: 25,
      fileFilters: {}
    };
    
    const view2 = {
      ...view1,
      viewName: 'view2'
    };
    
    dashboardViewService.saveViewConfig(userId, view1);
    dashboardViewService.saveViewConfig(userId, view2);
    
    // Delete one view
    const result = dashboardViewService.deleteViewConfig(userId, 'view1');
    expect(result).toBe(true);
    
    // Check that the view was deleted
    const viewConfig = dashboardViewService.getViewConfig(userId, 'view1');
    expect(viewConfig).toBeUndefined();
    
    // Check that the other view still exists
    const remainingView = dashboardViewService.getViewConfig(userId, 'view2');
    expect(remainingView).toBeDefined();
  });

  test('should not delete the only view', () => {
    // Try to delete the default view when it's the only one
    const result = dashboardViewService.deleteViewConfig(userId, 'default');
    expect(result).toBe(false);
    
    // Check that the view still exists
    const viewConfig = dashboardViewService.getViewConfig(userId, 'default');
    expect(viewConfig).toBeDefined();
  });

  test('should apply sorting to files', () => {
    const files: DriveFile[] = [
      { ...testFile, id: 'file1', name: 'Alpha', modifiedTime: '2023-01-01T10:00:00Z' },
      { ...testFile, id: 'file2', name: 'Beta', modifiedTime: '2023-01-02T10:00:00Z' },
      { ...testFile, id: 'file3', name: 'Gamma', modifiedTime: '2023-01-03T10:00:00Z' }
    ] as any;
    
    const viewConfig = dashboardViewService.getViewConfig(userId, 'default')!;
    viewConfig.sortBy = 'name';
    viewConfig.sortOrder = 'asc';
    
    const sortedFiles = dashboardViewService.applyViewConfig(files, viewConfig);
    
    expect(sortedFiles[0].name).toBe('Alpha');
    expect(sortedFiles[1].name).toBe('Beta');
    expect(sortedFiles[2].name).toBe('Gamma');
  });

  test('should apply date filtering to files', () => {
    const files: DriveFile[] = [
      { ...testFile, id: 'file1', name: 'Old File', modifiedTime: '2022-01-01T10:00:00Z' },
      { ...testFile, id: 'file2', name: 'Recent File', modifiedTime: '2023-06-01T10:00:00Z' },
      { ...testFile, id: 'file3', name: 'New File', modifiedTime: '2023-12-01T10:00:00Z' }
    ] as any;
    
    const viewConfig = dashboardViewService.getViewConfig(userId, 'default')!;
    viewConfig.fileFilters = {
      dateFrom: '2023-01-01T00:00:00Z',
      dateTo: '2023-10-01T00:00:00Z'
    };
    
    const filteredFiles = dashboardViewService.applyViewConfig(files, viewConfig);
    
    expect(filteredFiles).toHaveLength(1);
    expect(filteredFiles[0].name).toBe('Recent File');
  });

  test('should provide view templates', () => {
    const templates = dashboardViewService.getViewTemplates();
    
    expect(templates).toHaveLength(3);
    expect(templates.some(t => t.viewName === 'detailed')).toBe(true);
    expect(templates.some(t => t.viewName === 'compact')).toBe(true);
    expect(templates.some(t => t.viewName === 'grid')).toBe(true);
  });
});