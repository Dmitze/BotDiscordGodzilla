import type { Interaction } from 'discord.js';
import { tUser } from '@/i18n';

export function label(
  interaction: Interaction | unknown,
  key: string,
  params?: Record<string, string | number>
): string {
  return tUser(key, interaction as any, params);
}

export function desc(
  interaction: Interaction | unknown,
  key: string,
  params?: Record<string, string | number>
): string {
  return tUser(key, interaction as any, params);
}
