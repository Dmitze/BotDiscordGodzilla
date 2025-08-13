import { describe, it, expect, beforeEach, jest } from '@jest/globals';
import { SavedSearchCommand } from '../../../commands/SavedSearchCommand';
import { createMockConfig, createMockInteraction } from '../../utils/testHelpers';

describe('SavedSearchCommand (SQLite branch)', () => {
  let cmd: SavedSearchCommand;
  let interaction: any;
  let config: any;

  beforeEach(() => {
    config = createMockConfig();
    cmd = new SavedSearchCommand(config);
    interaction = createMockInteraction();
  });

  it('uses searchIndex.search with saved filters and renders list', async () => {
    const hits = { hits: [ { fileId: 'a', name: 'Alpha', contentHash: 'h', textLen: 5 } ], total: 1 };
    const searchIndex = { search: jest.fn().mockResolvedValue(hits) } as any;

    const saved = {
      name: 'my',
      filters: {
        query: 'budget',
        mimeIncludes: ['application/pdf'],
        ownerAllowlist: ['u@example.com'],
        dateFrom: '2024-01-01',
        dateTo: '2024-02-01',
        sizeMin: 1,
        sizeMax: 10,
        pageSize: 5,
        tags: ['finance']
      }
    };

    const ws = {
      getSavedSearch: jest.fn().mockReturnValue(saved),
      runSearch: jest.fn(),
      listSearches: jest.fn(),
      saveSearch: jest.fn(),
      removeSearch: jest.fn(),
    } as any;

    (interaction.client.serviceContainer.get as jest.Mock).mockImplementation((name: string) => {
      if (name === 'workspace') return ws;
      if (name === 'searchIndex') return searchIndex;
      if (name === 'google') return undefined; // чтобы не уходить во fallback
      return undefined;
    });

    interaction.options.getSubcommand.mockReturnValue('run');
    interaction.options.getString.mockImplementation((name: string, req?: boolean) => {
      if (name === 'name') return 'my';
      return null;
    });

    await cmd.execute({ interaction } as any);

    expect(ws.getSavedSearch).toHaveBeenCalledWith(interaction.user.id, 'my');
    expect(searchIndex.search).toHaveBeenCalledWith(
      expect.objectContaining({
        text: 'budget',
        limit: 5,
        filters: expect.objectContaining({
          mime: ['application/pdf'],
          owner: ['u@example.com'],
          tags: ['finance'],
        })
      })
    );
    expect(interaction.reply).toHaveBeenCalled();
  });
});
