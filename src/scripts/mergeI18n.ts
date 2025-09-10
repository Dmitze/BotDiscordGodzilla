/*
 * i18n Merge Script
 * Merges command-specific translations into the main i18n files
 */

import fs from 'fs';
import path from 'path';

type Json = string | number | boolean | null | JsonObject | Json[];
type JsonObject = { [key: string]: Json };

function isObject(v: unknown): v is JsonObject {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

const ROOT = path.resolve(__dirname, '..', '..');
const I18N_DIR = path.join(ROOT, 'src', 'i18n');

function readJson(file: string): JsonObject {
  const raw = fs.readFileSync(file, 'utf8');
  const parsed: unknown = JSON.parse(raw);
  if (!isObject(parsed)) return {};
  return parsed;
}

function writeJson(file: string, data: JsonObject): void {
  const json = JSON.stringify(data, null, 2) + '\n';
  fs.writeFileSync(file, json, 'utf8');
}

function mergeCommandsIntoMain(mainFile: string, commandsFile: string): void {
  const main = readJson(mainFile);
  const commands = readJson(commandsFile);
  
  // Merge commands into main under the "commands" key
  if (!main['commands']) {
    main['commands'] = {};
  }
  
  // Merge each command
  for (const [key, value] of Object.entries(commands)) {
    (main['commands'] as JsonObject)[key] = value;
  }
  
  writeJson(mainFile, main);
  console.log(`Merged ${commandsFile} into ${mainFile}`);
}

function main(): void {
  const ukMainFile = path.join(I18N_DIR, 'uk.json');
  const enMainFile = path.join(I18N_DIR, 'en.json');
  const ukCommandsFile = path.join(I18N_DIR, 'uk', 'commands.json');
  const enCommandsFile = path.join(I18N_DIR, 'en', 'commands.json');
  
  if (!fs.existsSync(ukMainFile) || !fs.existsSync(enMainFile)) {
    console.error('Main i18n files not found');
    process.exit(1);
  }
  
  if (!fs.existsSync(ukCommandsFile) || !fs.existsSync(enCommandsFile)) {
    console.error('Command i18n files not found');
    process.exit(1);
  }
  
  mergeCommandsIntoMain(ukMainFile, ukCommandsFile);
  mergeCommandsIntoMain(enMainFile, enCommandsFile);
  
  console.log('i18n merge completed successfully');
}

main();