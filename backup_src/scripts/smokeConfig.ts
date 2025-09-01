/*
 * Local smoke test for configuration
 * Loads env, builds Config and prints a safe summary without any network calls.
 */
import 'dotenv/config';
import { Config } from '@/config/Config';

function mask(val?: string): string {
  if (!val) return '';
  if (val.length <= 8) return '******';
  return `${val.slice(0, 4)}***${val.slice(-4)}`;
}

function main(): never {
  try {
    const cfg = Config.get();

    const summary = {
      discord: {
        clientId: mask(cfg.discord.clientId),
        guildId: cfg.discord.guildId ? mask(cfg.discord.guildId) : '(none)',
        prefix: cfg.discord.prefix,
        intents: cfg.discord.intents,
        enableChat: cfg.discord.enableChat ?? false,
        enableSlash: cfg.discord.enableSlash ?? false,
        enableMessageContentIntent: cfg.discord.enableMessageContentIntent ?? false,
      },
      google: {
        spreadsheetId: mask(cfg.google.spreadsheetId),
        driveFolderId: mask(cfg.google.driveFolderId),
        apiKey: mask(cfg.google.apiKey),
        appScriptUrl: cfg.google.appScriptUrl ? '(set)' : '(empty)',
        ocrProvider: cfg.google.ocrProvider ?? 'off',
      },
      drive: {
        pageSize: cfg.drive.pageSize,
        allowedMime: cfg.drive.allowedMime,
        fileMaxSizeMb: cfg.drive.fileMaxSizeMb,
        enableTextIndex: cfg.drive.enableTextIndex,
        indexCron: cfg.drive.indexCron,
      },
      ai: {
        provider: cfg.ai.provider,
        openaiModel: cfg.ai.openai.model,
        ollamaModel: cfg.ai.ollama.model,
      },
      metrics: cfg.metrics,
      security: cfg.security,
      performance: cfg.performance,
      logging: cfg.logging,
    };

    // eslint-disable-next-line no-console
    console.log('\n=== Config Smoke Summary ===');
    // eslint-disable-next-line no-console
    console.dir(summary, { depth: 3, colors: true });

    // eslint-disable-next-line no-console
    console.log('\nOK: configuration loaded successfully.');
    process.exit(0);
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error('Failed to load configuration:', err);
    process.exit(1);
  }
}

main();
