import { EmbedBuilder, Colors } from 'discord.js';
import type { Interaction , APIEmbedField } from 'discord.js';
import { tUser } from '@/i18n';

export interface LocalizedField {
  name: string; // i18n key
  value: string; // i18n key
  inline?: boolean;
  paramsName?: Record<string, string | number>;
  paramsValue?: Record<string, string | number>;
}

// Новий об'єктний формат опцій для побудови Embed
export interface EmbedOptions {
  interaction: Interaction | unknown;
  titleKey: string;
  descKey?: string;
  fields?: LocalizedField[];
  params?: {
    title?: Record<string, string | number>;
    desc?: Record<string, string | number>;
  };
  color?: number;
}

function mapFields(
  interaction: Interaction | unknown,
  fields?: LocalizedField[]
): APIEmbedField[] | undefined {
  if (!fields || fields.length === 0) return undefined;
  return fields.map((f) => ({
    name: tUser(f.name, interaction as any, f.paramsName),
    value: tUser(f.value, interaction as any, f.paramsValue),
    inline: f.inline ?? false,
  }));
}

// Нові функції V2 з об'єктними опціями
export function successEmbedV2(opts: EmbedOptions): EmbedBuilder {
  const builder = new EmbedBuilder()
    .setColor(opts.color ?? Colors.Green)
    .setTitle(tUser(opts.titleKey, opts.interaction as any, opts.params?.title));

  if (opts.descKey) {
    builder.setDescription(tUser(opts.descKey, opts.interaction as any, opts.params?.desc));
  }

  const mapped = mapFields(opts.interaction, opts.fields);
  if (mapped && mapped.length > 0) builder.setFields(mapped);

  return builder;
}

export function errorEmbedV2(opts: EmbedOptions): EmbedBuilder {
  const builder = new EmbedBuilder()
    .setColor(opts.color ?? Colors.Red)
    .setTitle(tUser(opts.titleKey, opts.interaction as any, opts.params?.title));

  if (opts.descKey) {
    builder.setDescription(tUser(opts.descKey, opts.interaction as any, opts.params?.desc));
  }

  const mapped = mapFields(opts.interaction, opts.fields);
  if (mapped && mapped.length > 0) builder.setFields(mapped);

  return builder;
}

// Залишаємо старі API як обгортки для зворотної сумісності
export function successEmbed(
  interaction: Interaction | unknown,
  keyTitle: string,
  keyDesc?: string,
  fields?: LocalizedField[],
  paramsTitle?: Record<string, string | number>,
  paramsDesc?: Record<string, string | number>
): EmbedBuilder {
  const params: { title?: Record<string, string | number>; desc?: Record<string, string | number> } = {};
  if (paramsTitle !== undefined) params.title = paramsTitle;
  if (paramsDesc !== undefined) params.desc = paramsDesc;
  const opts: EmbedOptions = { interaction, titleKey: keyTitle, params, color: Colors.Green };
  if (fields !== undefined) (opts as any).fields = fields;
  if (keyDesc !== undefined) (opts as any).descKey = keyDesc;
  return successEmbedV2(opts);
}

export function errorEmbed(
  interaction: Interaction | unknown,
  keyTitle: string,
  keyDesc?: string,
  fields?: LocalizedField[],
  paramsTitle?: Record<string, string | number>,
  paramsDesc?: Record<string, string | number>
): EmbedBuilder {
  const params: { title?: Record<string, string | number>; desc?: Record<string, string | number> } = {};
  if (paramsTitle !== undefined) params.title = paramsTitle;
  if (paramsDesc !== undefined) params.desc = paramsDesc;
  const opts: EmbedOptions = { interaction, titleKey: keyTitle, params, color: Colors.Red };
  if (fields !== undefined) (opts as any).fields = fields;
  if (keyDesc !== undefined) (opts as any).descKey = keyDesc;
  return errorEmbedV2(opts);
}
