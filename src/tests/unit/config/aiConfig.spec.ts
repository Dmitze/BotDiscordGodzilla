import { Config } from '@/config/Config';

describe('Config AI and Metrics defaults', () => {
  const OLD = process.env;

  beforeEach(() => {
    jest.resetModules();
    process.env = { ...OLD };
    Config.clearCache();
  });

  afterEach(() => {
    process.env = OLD;
  });

  const baseEnv = () => {
    process.env['DISCORD_TOKEN'] = 'x';
    process.env['DISCORD_CLIENT_ID'] = 'y';
    process.env['DISCORD_GUILD_ID'] = 'z';
    process.env['GOOGLE_SPREADSHEET_ID'] = 'sheet';
    process.env['GOOGLE_DRIVE_FOLDER_ID'] = 'folder';
    process.env['GOOGLE_APPLICATION_CREDENTIALS'] = '{}';
    process.env['GOOGLE_API_KEY'] = 'AIza_test';
    process.env['GOOGLE_APP_SCRIPT_URL'] = 'http://local';
  };

  it('does not require OPENAI_API_KEY when AI_PROVIDER is ollama', () => {
    baseEnv();
    process.env['GOOGLE_DRIVE_FOLDER_ID'] = 'folder';
    process.env['AI_PROVIDER'] = 'ollama';
    process.env['OLLAMA_HOST'] = 'http://localhost:11434';
    process.env['OLLAMA_MODEL'] = 'llama2';

    const cfg = Config.load();
    // fill minimal google vars required by Config.load to avoid throw in tests

    expect(cfg.ai.provider).toBe('ollama');
    // no throw when OPENAI_API_KEY is not set
    expect(typeof cfg.ai.openai.apiKey).toBe('string');
  });

  it('uses 9091 as default METRICS_PORT', () => {
    baseEnv();
    process.env['GOOGLE_DRIVE_FOLDER_ID'] = 'folder';
    delete process.env['METRICS_PORT'];

    const cfg = Config.load();
    expect(cfg.metrics.port).toBe(9091);
  });
});
<<<<<<< HEAD
=======

>>>>>>> e563d3ca (test(unit): add aiConfig.spec for AI and Metrics defaults)
