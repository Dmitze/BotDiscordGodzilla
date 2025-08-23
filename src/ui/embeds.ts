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

export function successEmbed(
  interaction: Interaction | unknown,
  keyTitle: string,
  keyDesc?: string,
  fields?: LocalizedField[],
  paramsTitle?: Record<string, string | number>,
  paramsDesc?: Record<string, string | number>
): EmbedBuilder {
  const builder = new EmbedBuilder()
    .setColor(Colors.Green)
    .setTitle(tUser(keyTitle, interaction as any, paramsTitle));

  if (keyDesc) {
    builder.setDescription(tUser(keyDesc, interaction as any, paramsDesc));
  }

  const mapped = mapFields(interaction, fields);
  if (mapped && mapped.length > 0) builder.setFields(mapped);

  return builder;
}

export function errorEmbed(
  interaction: Interaction | unknown,
  keyTitle: string,
  keyDesc?: string,
  fields?: LocalizedField[],
  paramsTitle?: Record<string, string | number>,
  paramsDesc?: Record<string, string | number>
): EmbedBuilder {
  const builder = new EmbedBuilder()
    .setColor(Colors.Red)
    .setTitle(tUser(keyTitle, interaction as any, paramsTitle));

  if (keyDesc) {
    builder.setDescription(tUser(keyDesc, interaction as any, paramsDesc));
  }

  const mapped = mapFields(interaction, fields);
  if (mapped && mapped.length > 0) builder.setFields(mapped);

  return builder;
}
