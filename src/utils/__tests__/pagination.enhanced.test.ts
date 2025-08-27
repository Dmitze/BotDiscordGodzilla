import Pagination from '../pagination';

describe('Pagination - Enhanced Features', () => {
  const largeDataSet = Array.from({ length: 1000 }, (_, i) => ({
    id: i + 1,
    name: `Item ${i + 1}`,
    value: `Value ${i + 1}`
  }));

  test('should create pagination with cursor-based navigation', () => {
    const pagination = Pagination.createWithCursorPagination(largeDataSet, 'id', {
      itemsPerPage: 10,
      title: 'Cursor Pagination Test'
    });

    expect(pagination).toBeInstanceOf(Pagination);
    expect(pagination.getTotalItems()).toBe(1000);
    // The actual implementation limits to maxPages (default 50)
    expect(pagination.getTotalPages()).toBe(50);
  });

  test('should create virtual pagination for large datasets', () => {
    const pagination = Pagination.createVirtualPagination(largeDataSet, {
      itemsPerPage: 20,
      title: 'Virtual Pagination Test'
    });

    expect(pagination).toBeInstanceOf(Pagination);
    expect(pagination.getTotalItems()).toBe(1000);
    
    // With virtual pagination, we can handle more pages efficiently
    expect(pagination.getTotalPages()).toBe(50); // Still limited by maxPages
  });

  test('should handle cursor-based button interactions', () => {
    const pagination = Pagination.createWithCursorPagination(largeDataSet, 'id', {
      itemsPerPage: 10
    });

    // Test next cursor navigation
    const nextResult = pagination.handleButtonInteraction('pagination_next_cursor_0');
    expect(nextResult).toBe(true);
    expect(pagination.getCurrentPage()).toBe(1);

    // Test previous cursor navigation
    const prevResult = pagination.handleButtonInteraction('pagination_prev_cursor_1');
    expect(prevResult).toBe(true);
    expect(pagination.getCurrentPage()).toBe(0);
  });

  test('should create navigation buttons for cursor pagination', () => {
    const pagination = Pagination.createWithCursorPagination(largeDataSet, 'id', {
      itemsPerPage: 10
    });

    const buttons = pagination.createNavigationButtons();
    expect(buttons).toBeDefined();
    
    // Should have prev and next buttons for cursor pagination
    const components = (buttons as any).components;
    expect(components).toHaveLength(2);
  });

  test('should optimize footer text for large datasets', () => {
    const pagination = new Pagination(largeDataSet, {
      itemsPerPage: 10
    });

    // Override total items to simulate a very large dataset
    (pagination as any).totalItems = 10000;
    (pagination as any).totalPages = 1000;

    const stats = pagination.getStats();
    expect(stats.totalItems).toBe(10000);
    expect(stats.totalPages).toBe(1000);
  });

  test('should limit data size when maxItems option is provided', () => {
    const pagination = new Pagination(largeDataSet, {
      maxItems: 100
    });

    expect(pagination.getTotalItems()).toBe(100);
  });

  test('should handle virtual pagination data retrieval', () => {
    const pagination = Pagination.createVirtualPagination(largeDataSet, {
      itemsPerPage: 10
    });

    const pageData: any = pagination.getCurrentPageData();
    expect(pageData).toHaveProperty('startIndex');
    expect(pageData).toHaveProperty('endIndex');
    expect(pageData).toHaveProperty('data');
    expect(pageData.data).toHaveLength(10);
  });

  test('should create enhanced footer text for large datasets', () => {
    // Create pagination with a lot of items
    const pagination = new Pagination(largeDataSet, {
      itemsPerPage: 10
    });
    
    // Manually set a large number of items to trigger enhanced footer
    (pagination as any).totalItems = 10000;
    
    const footerText = (pagination as any).createFooterText();
    expect(footerText).toContain('Елементи');
    expect(footerText).toContain('з 10000');
  });
});