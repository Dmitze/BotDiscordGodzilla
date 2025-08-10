import { describe, it, expect } from '@jest/globals';
import { sortResults } from '../searchHelpers';

describe('searchHelpers hardening', () => {
  it('parseLocaleNumber used in sortResults handles 1 234,56', () => {
    const headers = ['ціна'];
    const data: any[][] = [["1 234,56"],["10,5"],["100.25"]];
    const sorted = sortResults(data, headers, 'ціна', 'asc');
    expect(sorted?.[0]?.[0]).toBe("10,5");
    expect(sorted?.[2]?.[0]).toBe("1 234,56");
  });
});

