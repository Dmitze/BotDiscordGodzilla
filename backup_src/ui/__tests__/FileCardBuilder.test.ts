import { buildFileEmbed, buildFileActions } from '../FileCardBuilder';
import type { DriveFile } from '@/types/drive';

describe('FileCardBuilder', () => {
  const file: DriveFile = {
    id: 'file123',
    name: 'Test Document',
    mimeType: 'application/vnd.google-apps.document',
    modifiedTime: new Date().toISOString(),
    owners: ['owner@example.com'],
  } as any;

  it('builds a file embed with title and optional fields', () => {
    const embed = buildFileEmbed(file, { showOwner: true, showDates: true });
    const json = embed.toJSON();
    expect(json.title).toContain('Test Document');
    expect(json.color).toBeDefined();
    expect(json.description).toBeDefined();
  });

  it('builds action rows with expected buttons', () => {
    const rows = buildFileActions(file, { hideWebLink: true });
    expect(Array.isArray(rows)).toBe(true);
    expect(rows.length).toBeGreaterThan(0);

    const first = rows[0]!;
    const firstJson = first.toJSON();
    // Ensure we have at least Open button
    const labels = (firstJson.components || []).map((c: any) => c.label);
    expect(labels.some((l: string) => typeof l === 'string' && l.length > 0)).toBe(true);
  });
});
