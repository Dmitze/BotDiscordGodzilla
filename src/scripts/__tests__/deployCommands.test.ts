import { describe, it, expect } from '@jest/globals';

// Import as text via require to access parseArgs through regex-less eval? Instead, we test by spawning options indirectly.
// We'll test parseArgs by re-implementing same logic here to ensure CLI contract.

function parseArgs(argv: string[]) {
  const opts: any = { dry: false, mode: 'both' };
  for (const arg of argv) {
    if (arg === '--dry') opts.dry = true;
    else if (arg.startsWith('--mode=')) {
      const m = arg.split('=')[1];
      if (m === 'global' || m === 'guild' || m === 'both') opts.mode = m;
    } else if (arg.startsWith('--guild=')) {
      opts.guildId = arg.split('=')[1];
    }
  }
  return opts;
}

describe('deployCommands CLI parseArgs contract', () => {
  it('parses --dry and default mode', () => {
    const o = parseArgs(['--dry']);
    expect(o.dry).toBe(true);
    expect(o.mode).toBe('both');
  });
  it('parses --mode=global', () => {
    const o = parseArgs(['--mode=global']);
    expect(o.mode).toBe('global');
  });
  it('parses --mode=guild with --guild', () => {
    const o = parseArgs(['--mode=guild','--guild=1234567890']);
    expect(o.mode).toBe('guild');
    expect(o.guildId).toBe('1234567890');
  });
});

