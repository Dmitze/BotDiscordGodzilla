/**
 * Unit тесты для утилиты pagination
 */

import { describe, it, expect } from '@jest/globals';

import { createPaginationEmbed, createPaginationRow } from '../../../utils/pagination';

describe('Pagination Utils', () => {
  describe('createPaginationEmbed', () => {
    it('should create pagination embed with data', () => {
      const data = [
        { id: '1', name: 'Item 1' },
        { id: '2', name: 'Item 2' },
        { id: '3', name: 'Item 3' },
      ];

      const embed = createPaginationEmbed(data, 0, 2, 'Test Title');

      expect(embed.title).toBe('Test Title');
      expect(embed.fields).toHaveLength(2);
      expect(embed.footer?.text).toContain('1-2 з 3');
    });

    it('should handle empty data', () => {
      const embed = createPaginationEmbed([], 0, 10, 'Empty Title');

      expect(embed.title).toBe('Empty Title');
      expect(embed.fields).toHaveLength(0);
      expect(embed.footer?.text).toContain('0 з 0');
    });

    it('should handle single page', () => {
      const data = [{ id: '1', name: 'Item 1' }];
      const embed = createPaginationEmbed(data, 0, 10, 'Single Page');

      expect(embed.footer?.text).toContain('1 з 1');
    });

    it('should handle last page', () => {
      const data = [
        { id: '1', name: 'Item 1' },
        { id: '2', name: 'Item 2' },
        { id: '3', name: 'Item 3' },
        { id: '4', name: 'Item 4' },
        { id: '5', name: 'Item 5' },
      ];

      const embed = createPaginationEmbed(data, 4, 2, 'Last Page');

      expect(embed.footer?.text).toContain('5 з 5');
    });

    it('should format field values correctly', () => {
      const data = [
        { id: '1', name: 'Item 1', value: 'Value 1' },
        { id: '2', name: 'Item 2', value: 'Value 2' },
      ];

      const embed = createPaginationEmbed(data, 0, 2, 'Formatted Title');

      expect(embed.fields?.[0]?.name).toContain('Item 1');
      expect(embed.fields?.[0]?.value).toContain('Value 1');
    });
  });

  describe('createPaginationRow', () => {
    it('should create pagination row with all buttons', () => {
      const row = createPaginationRow(1, 5, 'test');

      expect(row.components).toHaveLength(4); // Previous, Next, First, Last
    });

    it('should disable previous button on first page', () => {
      const row = createPaginationRow(0, 5, 'test');

      const previousButton = row.components.find(comp => 
        comp.data?.custom_id === 'test_prev'
      );
      expect(previousButton?.disabled).toBe(true);
    });

    it('should disable next button on last page', () => {
      const row = createPaginationRow(4, 5, 'test');

      const nextButton = row.components.find(comp => 
        comp.data?.custom_id === 'test_next'
      );
      expect(nextButton?.disabled).toBe(true);
    });

    it('should disable first button on first page', () => {
      const row = createPaginationRow(0, 5, 'test');

      const firstButton = row.components.find(comp => 
        comp.data?.custom_id === 'test_first'
      );
      expect(firstButton?.disabled).toBe(true);
    });

    it('should disable last button on last page', () => {
      const row = createPaginationRow(4, 5, 'test');

      const lastButton = row.components.find(comp => 
        comp.data?.custom_id === 'test_last'
      );
      expect(lastButton?.disabled).toBe(true);
    });

    it('should enable all buttons on middle page', () => {
      const row = createPaginationRow(2, 5, 'test');

      row.components.forEach(component => {
        expect(component.disabled).toBe(false);
      });
    });

    it('should handle single page', () => {
      const row = createPaginationRow(0, 1, 'test');

      row.components.forEach(component => {
        expect(component.disabled).toBe(true);
      });
    });

    it('should use correct custom IDs', () => {
      const row = createPaginationRow(1, 3, 'custom_prefix');

      const expectedIds = [
        'custom_prefix_prev',
        'custom_prefix_next',
        'custom_prefix_first',
        'custom_prefix_last',
      ];

      row.components.forEach((component, index) => {
        expect(component.data?.custom_id).toBe(expectedIds[index]);
      });
    });
  });

  describe('edge cases', () => {
    it('should handle negative page index', () => {
      const embed = createPaginationEmbed([{ id: '1', name: 'Item' }], -1, 10, 'Test');
      expect(embed.footer?.text).toContain('0 з 1');
    });

    it('should handle page index beyond data length', () => {
      const data = [{ id: '1', name: 'Item' }];
      const embed = createPaginationEmbed(data, 5, 10, 'Test');
      expect(embed.fields).toHaveLength(0);
    });

    it('should handle zero items per page', () => {
      const data = [{ id: '1', name: 'Item' }];
      const embed = createPaginationEmbed(data, 0, 0, 'Test');
      expect(embed.fields).toHaveLength(0);
    });

    it('should handle null or undefined data', () => {
      const embed = createPaginationEmbed(null as any, 0, 10, 'Test');
      expect(embed.fields).toHaveLength(0);
    });
  });
}); 