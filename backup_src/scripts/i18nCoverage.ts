/*
 * i18n Coverage Report Generator
 * Scans src/ for i18n key usage (t()/tUser()) and compares with src/i18n/uk.json, en.json.
 * Outputs docs/i18n-coverage.md
 */

import fs from 'fs';
import path from 'path';

type Json = string | number | boolean | null | JsonObject | Json[];
type JsonObject = { [key: string]: Json };

function isObject(v: unknown): v is JsonObject {
  return typeof v === 'object' && v !== null && !Array.isArray(v);
}

const ROOT = path.resolve(__dirname, '..', '..');
const SRC_DIR = path.join(ROOT, 'src');
const I18N_DIR = path.join(SRC_DIR, 'i18n');
const REPORT_PATH = path.join(ROOT, 'docs', 'i18n-coverage.md');

function readJson(file: string): JsonObject {
  const raw = fs.readFileSync(file, 'utf8');
  const parsed: unknown = JSON.parse(raw);
  if (!isObject(parsed)) return {};
  return parsed;
}

function flattenKeys(obj: Json, prefix = ''): string[] {
  if (!isObject(obj)) return [];
  const keys: string[] = [];
  for (const k of Object.keys(obj)) {
    const v = obj[k];
    const full = prefix ? `${prefix}.${k}` : k;
    if (isObject(v)) {
      keys.push(...flattenKeys(v, full));
    } else {
      keys.push(full);
    }
  }
  return keys;
}

function walk(dir: string, out: string[] = []): string[] {
  const items = fs.readdirSync(dir, { withFileTypes: true });
  for (const it of items) {
    if (it.name === 'node_modules' || it.name === 'dist') continue;
    const p = path.join(dir, it.name);
    if (it.isDirectory()) {
      walk(p, out);
    } else if (it.isFile() && it.name.endsWith('.ts')) {
      out.push(p);
    }
  }
  return out;
}

function extractKeysFromFile(filePath: string): string[] {
  const content = fs.readFileSync(filePath, 'utf8');
  const keys = new Set<string>();
  // t('key') | t("key") | t(`key`)
  const re1 = /\bt\(\s*["'`]([\w.-]+)["'`]/g;
  const re2 = /\btUser\(\s*["'`]([\w.-]+)["'`]/g;
  let m: RegExpExecArray | null;
  while ((m = re1.exec(content))) {
    const k = m[1];
    if (typeof k === 'string' && k.length) keys.add(k);
  }
  while ((m = re2.exec(content))) {
    const k = m[1];
    if (typeof k === 'string' && k.length) keys.add(k);
  }
  return Array.from(keys);
}

function main(): void {
  const ukJsonPath = path.join(I18N_DIR, 'uk.json');
  const enJsonPath = path.join(I18N_DIR, 'en.json');
  if (!fs.existsSync(ukJsonPath) || !fs.existsSync(enJsonPath)) {
    // eslint-disable-next-line no-console
    console.error('i18n files not found at src/i18n/uk.json or src/i18n/en.json');
    process.exit(1);
  }

  const uk = readJson(ukJsonPath);
  const en = readJson(enJsonPath);
  const ukKeys = new Set(flattenKeys(uk));
  const enKeys = new Set(flattenKeys(en));

  const tsFiles = walk(SRC_DIR);
  const usedKeys = new Set<string>();
  for (const f of tsFiles) {
    // skip tests if they reside under src/tests
    if (f.includes(`${path.sep}tests${path.sep}`)) continue;
    extractKeysFromFile(f).forEach(k => usedKeys.add(k));
  }

  const used = Array.from(usedKeys).sort();
  const onlyInUk = used.filter(k => !enKeys.has(k) && ukKeys.has(k));
  const onlyInEn = used.filter(k => !ukKeys.has(k) && enKeys.has(k));
  const missingInUk = used.filter(k => !ukKeys.has(k));
  const missingInEn = used.filter(k => !enKeys.has(k));

  const summary = [
    `# i18n Coverage Report`,
    '',
    `Generated: ${new Date().toISOString()}`,
    '',
    `- Total used keys in code: ${used.length}`,
    `- Present in uk.json: ${used.filter(k => ukKeys.has(k)).length}`,
    `- Present in en.json: ${used.filter(k => enKeys.has(k)).length}`,
    `- Missing in uk.json: ${missingInUk.length}`,
    `- Missing in en.json: ${missingInEn.length}`,
    '',
    `## Missing in uk.json (${missingInUk.length})`,
    '',
    missingInUk.length ? missingInUk.map(k => `- ${k}`).join('\n') : '_None_',
    '',
    `## Missing in en.json (${missingInEn.length})`,
    '',
    missingInEn.length ? missingInEn.map(k => `- ${k}`).join('\n') : '_None_',
    '',
    `## Used keys only found in uk.json (${onlyInUk.length})`,
    '',
    onlyInUk.length ? onlyInUk.map(k => `- ${k}`).join('\n') : '_None_',
    '',
    `## Used keys only found in en.json (${onlyInEn.length})`,
    '',
    onlyInEn.length ? onlyInEn.map(k => `- ${k}`).join('\n') : '_None_',
    '',
  ].join('\n');

  // ensure docs directory exists
  const docsDir = path.join(ROOT, 'docs');
  if (!fs.existsSync(docsDir)) fs.mkdirSync(docsDir, { recursive: true });
  fs.writeFileSync(REPORT_PATH, summary, 'utf8');

  // eslint-disable-next-line no-console
  console.log(`i18n coverage report written to ${REPORT_PATH}`);
}

main();
