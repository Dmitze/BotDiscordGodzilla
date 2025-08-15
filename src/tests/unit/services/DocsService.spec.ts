import { DocsService } from '@/services/google/DocsService';

describe('DocsService', () => {
  const svc = new DocsService();

  test('extractTextFromDoc returns plain text from paragraphs', () => {
    const doc = {
      body: {
        content: [
          { paragraph: { elements: [ { textRun: { content: 'Hello ' } }, { textRun: { content: 'World\n' } } ] } },
          { paragraph: { elements: [ { textRun: { content: 'Next line' } } ] } },
        ],
      },
    };
    const text = svc.extractTextFromDoc(doc);
    expect(text).toBe('Hello World\n\nNext line');
  });

  test('extractBlocksFromDoc parses headings, list items, table, footnotes', () => {
    const doc = {
      body: {
        content: [
          { paragraph: { paragraphStyle: { namedStyleType: 'HEADING_2' }, elements: [ { textRun: { content: 'Title\n' } } ] } },
          { paragraph: { bullet: { listId: 'k1' }, elements: [ { textRun: { content: 'Item 1\n' } } ] } },
          { table: { tableRows: [ { tableCells: [ { content: [ { paragraph: { elements: [ { textRun: { content: 'A' } } ] } } ] }, { content: [ { paragraph: { elements: [ { textRun: { content: 'B' } } ] } } ] } ] } ] } },
          { paragraph: { elements: [ { textRun: { content: 'Plain paragraph' } } ] } },
        ],
      },
      footnotes: {
        f1: { content: [ { paragraph: { elements: [ { textRun: { content: 'Foot note' } } ] } } ] },
      },
    };

    const blocks = svc.extractBlocksFromDoc(doc);
    expect(blocks.find(b => b.kind === 'heading' && (b as any).level === 2)).toBeTruthy();
    expect(blocks.find(b => b.kind === 'listItem' && (b as any).text.includes('Item 1'))).toBeTruthy();
    const table = blocks.find(b => b.kind === 'table') as any;
    expect(table).toBeTruthy();
    expect(table.rows[0].cells.map((c: any) => c.text)).toEqual(['A', 'B']);
    expect(blocks.find(b => b.kind === 'paragraph' && (b as any).text === 'Plain paragraph')).toBeTruthy();
    expect(blocks.find(b => b.kind === 'footnote' && (b as any).id === 'f1')).toBeTruthy();
  });
});
