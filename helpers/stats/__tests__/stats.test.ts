import fs from 'fs';
import path from 'path';

const { BotStats } = require('../stats');

describe('BotStats FS safety', () => {
  const statsPath = path.join('data', 'logs', 'stats.json');

  beforeEach(() => {
    try { fs.rmSync(path.dirname(statsPath), { recursive: true, force: true }); } catch {}
  });

  it('writes to data/logs and initializes structure', (done) => {
    const s = new BotStats();
    s.trackCommand('cmd', 'u1', 'g1', true);
    setTimeout(() => {
      expect(fs.existsSync(statsPath)).toBe(true);
      const raw = fs.readFileSync(statsPath, 'utf8');
      const json = JSON.parse(raw);
      expect(json.totalCommands).toBe(1);
      done();
    }, 300);
  });

  it('serializes Set fields and restores them', (done) => {
    const s = new BotStats();
    s.trackCommand('cmd', 'u1');
    setTimeout(() => {
      const s2 = new BotStats();
      // users in command stats is a Set again
      const users = (s2 as any).getStats().commands['cmd'].users;
      expect(Array.isArray(users)).toBe(false); // Set, not array
      done();
    }, 300);
  });

  it('initializes dailyStats on error', (done) => {
    const s = new BotStats();
    s.trackError('boom', 'cmd', 'u1');
    setTimeout(() => {
      const json = s.getDailyStats() as any;
      const today = (new Date().toISOString().split('T')[0]) as string;
      expect(Object.prototype.hasOwnProperty.call(json as object, today)).toBe(true);
      done();
    }, 300);
  });
});

