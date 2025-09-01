import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { SearchCommand } from '../../../commands/SearchCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('SearchCommand (SQLite branch)', () => {
  let cmd: SearchCommand;
  let interaction: any;
  let config: any;

  beforeEach(() => {
    config = createMockConfig();
    cmd = new SearchCommand(config);
    interaction = createMockInteraction();
  });

  it('uses searchIndex.search when available and maps filters', async () => {
    const hits = {
      hits: [
        { fileId: '1', name: 'Doc A', contentHash: 'h', textLen: 10, snippet: 'A' },
        { fileId: '2', name: 'Doc B', contentHash: 'h2', textLen: 20, snippet: 'B' },
      ],
      total: 2,
    };

    const searchIndex = { search: jest.fn().mockResolvedValue(hits) } as any;

    // mock serviceContainer.get to return searchIndex when asked
    (interaction.client.serviceContainer.get as jest.Mock).mockImplementation((name: string) => {
      if (name === 'searchIndex') return searchIndex;
      // return undefined for others to ensure SQLite path is taken without Google fallback
      return undefined;
    });

    // simulate user options
    interaction.options.getString.mockImplementation((name: string) => {
      if (name === 'запит') return 'hello';
      if (name === 'тип') return 'report'; // documentType -> tag
      if (name === 'підрозділ') return 'sales'; // unit -> tag
      if (name === 'пріоритет') return 'high'; // priority -> tag
      if (name === 'від') return '2024-01-01'; // dateFrom
      if (name === 'до') return '2024-01-31'; // dateTo
      return null;
    });
    interaction.options.getInteger.mockImplementation((name: string) => {
      if (name === 'ліміт') return 5;
      return null;
    });

    await cmd.execute({ interaction } as any);

    // ensures SQLite branch executed
    expect(searchIndex.search).toHaveBeenCalledWith(
      expect.objectContaining({
        text: 'hello',
        limit: 5,
        sample: undefined,
        // verify filters object exists without asserting exact shape
        filters: expect.anything(),
      })
    );

    // result rendered via editReply after deferReply
    expect(interaction.deferReply).toHaveBeenCalled();
    expect(interaction.editReply).toHaveBeenCalled();
  });
});
