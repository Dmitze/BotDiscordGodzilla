import { SheetsService } from '@/services/google/SheetsService';

describe('SheetsService', () => {
  const svc = new SheetsService();

  test('toSheetDataFromGet normalizes range and values', () => {
    const raw: any = { range: '  Sheet1!A1:B2  ', values: [[1, null], ['x']] };
    const out = svc.toSheetDataFromGet(raw, 'Fallback!A1');
    expect(out.range).toBe('Sheet1!A1:B2');
    expect(out.majorDimension).toBe('ROWS');
    expect(out.values).toEqual([[ '1', '' ], [ 'x' ]]);
  });

  test('parseBatchGet maps valueRanges safely', () => {
    const raw: any = { valueRanges: [ { range: 'S!A1', values: [[1,'2'],[null]] } ] };
    const out = svc.parseBatchGet(raw);
    expect(out.valueRanges.length).toBe(1);
    const vr0 = out.valueRanges[0];
    expect(vr0 && vr0.range).toBe('S!A1');
    expect(vr0 && vr0.values).toEqual([[1,'2'],[null]]);
  });

  test('buildBatchUpdate normalizes ranges and validates size', () => {
    const req: any = {
      valueInputOption: 'RAW',
      data: [ { range: '  S!A1 ', values: [[true, { a: 1 }]] } ]
    };
    const out = svc.buildBatchUpdate(req);
    expect(out.valueInputOption).toBe('RAW');
    expect(out.data.length).toBe(1);
    const first = out.data[0];
    expect(first?.range).toBe('S!A1');
    expect(first?.values).toEqual([['TRUE', '{"a":1}']]);
  });

  test('validateBatchWrite throws on too many cells', () => {
    expect(() => svc.validateBatchWrite([ { range: 'S!A1', values: Array.from({length: 300}, () => Array(300).fill(0)) } ], { maxCells: 1000 } as any))
      .toThrow(/Превышен лимит/);
  });

  test('normalizeWriteValues converts types as expected', () => {
    const input: any = [[undefined, null, true, false, 42, 'str', { x: 1 }]];
    const norm = svc.normalizeWriteValues(input);
    expect(norm).toEqual([[
      null,
      null,
      'TRUE',
      'FALSE',
      42,
      'str',
      '{"x":1}',
    ]]);
  });

  test('validateBatchWrite does not throw under limit', () => {
    const small = [ { range: 'S!A1', values: Array.from({length: 10}, () => Array(10).fill(1)) } ];
    expect(() => svc.validateBatchWrite(small, { maxCells: 200 } as any)).not.toThrow();
  });
});
