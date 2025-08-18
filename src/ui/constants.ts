// UI constants and defaults
import { t } from '@/i18n';

export const UI_COLORS = {
  success: 0x22c55e, // green-500
  error: 0xef4444,   // red-500
  info: 0x3b82f6,    // blue-500
  warn: 0xf59e0b,    // amber-500
} as const;

export const UI_EMOJI = {
  file: '📄',
  folder: '📁',
  sheet: '📊',
  package: '📦',
  new: '🆕',
  edited: '✏️',
  link: '🔗',
} as const;

export function i18nTitleFor(kind: 'search' | 'read' | 'file' | 'folder'): string {
  switch (kind) {
    case 'search': return t('files.sub.search.title') || 'Пошук';
    case 'read': return t('files.sub.read.title') || 'Читання';
    case 'folder': return t('files.common.folder') || 'Папка';
    case 'file':
    default:
      return t('files.common.file') || 'Файл';
  }
}
