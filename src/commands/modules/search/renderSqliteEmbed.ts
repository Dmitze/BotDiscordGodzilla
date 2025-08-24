import { EmbedBuilder } from 'discord.js';

export async function renderSqliteEmbed(
  interaction: { editReply: (arg: any) => Promise<any> },
  query: string,
  hits: Array<{ name?: string; fileId?: string; id?: string; snippet?: string }>,
  total: number,
  limit: number
): Promise<void> {
  const lines = hits.map(h => {
    const title = h.name || h.fileId || h.id;
    const snip = h.snippet ? ` — ${String(h.snippet).replace(/\n/g, ' ').slice(0, 120)}${String(h.snippet).length > 120 ? '…' : ''}` : '';
    return `• ${title}${snip}`;
  });
  const embed = new EmbedBuilder()
    .setColor('#4CAF50')
    .setTitle('🔍 Результати пошуку (SQLite)')
    .setDescription(`**Запит:** ${query}`)
    .addFields(
      { name: '📊 Знайдено (оцінено)', value: String(total ?? hits.length), inline: true },
      { name: '⚡ Джерело', value: 'SQLite FTS', inline: true },
    )
    .setTimestamp();
  const body = lines.slice(0, limit || 10).join('\n');
  if (body.length > 0) {
    embed.addFields({ name: `📋 Результати (${Math.min(lines.length, limit || 10)})`, value: body.length > 1024 ? body.slice(0, 1021) + '...' : body });
  }
  await interaction.editReply({ embeds: [embed], components: [] });
}
