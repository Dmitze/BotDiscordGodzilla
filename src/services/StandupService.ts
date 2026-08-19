import { Client, TextChannel, EmbedBuilder, ActionRowBuilder, ModalBuilder, TextInputBuilder, TextInputStyle, Interaction, ButtonBuilder, ButtonStyle, ModalSubmitInteraction, ButtonInteraction } from 'discord.js';
import logger from '@/utils/logger';

export class StandupService {
  private client: Client;
  private isRegistered = false;
  
  constructor(client: Client) {
    this.client = client;
    this.registerInteractionHandlers();
  }
  
  private registerInteractionHandlers() {
    if (this.isRegistered) return;
    
    this.client.on('interactionCreate', async (interaction: Interaction) => {
      try {
        if (interaction.isButton() && interaction.customId === 'standup_start') {
          await this.handleStandupStart(interaction as ButtonInteraction);
        } else if (interaction.isModalSubmit() && interaction.customId === 'standup_modal') {
          await this.handleStandupSubmit(interaction as ModalSubmitInteraction);
        }
      } catch (error) {
        logger.error('Error handling standup interaction', { errorMessage: String(error) } as any);
      }
    });
    this.isRegistered = true;
  }

  public async triggerStandup(channelId: string) {
    try {
      const channel = await this.client.channels.fetch(channelId) as TextChannel;
      if (!channel) return;

      const embed = new EmbedBuilder()
        .setTitle('🗓️ Час для Daily Standup!')
        .setDescription('Натисніть кнопку нижче, щоб заповнити свій щоденний звіт. Це допоможе команді залишатися синхронізованою!')
        .setColor('#3498db')
        .setFooter({ text: 'Godzilla Business Assistant' });

      const button = new ButtonBuilder()
        .setCustomId('standup_start')
        .setLabel('📝 Написати звіт')
        .setStyle(ButtonStyle.Primary)
        .setEmoji('📋');

      const row = new ActionRowBuilder<ButtonBuilder>().addComponents(button);

      await channel.send({ embeds: [embed], components: [row] });
    } catch (err) {
      logger.error('Failed to trigger standup', { errorMessage: String(err) } as any);
    }
  }

  private async handleStandupStart(interaction: ButtonInteraction) {
    const modal = new ModalBuilder()
      .setCustomId('standup_modal')
      .setTitle('Твій Daily Standup');

    const yesterdayInput = new TextInputBuilder()
      .setCustomId('standup_yesterday')
      .setLabel('Що ти зробив(ла) вчора?')
      .setStyle(TextInputStyle.Paragraph)
      .setRequired(true);

    const todayInput = new TextInputBuilder()
      .setCustomId('standup_today')
      .setLabel('Що плануєш робити сьогодні?')
      .setStyle(TextInputStyle.Paragraph)
      .setRequired(true);

    const blockersInput = new TextInputBuilder()
      .setCustomId('standup_blockers')
      .setLabel('Чи є якісь блокери?')
      .setStyle(TextInputStyle.Short)
      .setRequired(false)
      .setValue('Немає');

    const actionRow1 = new ActionRowBuilder<TextInputBuilder>().addComponents(yesterdayInput);
    const actionRow2 = new ActionRowBuilder<TextInputBuilder>().addComponents(todayInput);
    const actionRow3 = new ActionRowBuilder<TextInputBuilder>().addComponents(blockersInput);

    modal.addComponents(actionRow1, actionRow2, actionRow3);
    await interaction.showModal(modal);
  }

  private async handleStandupSubmit(interaction: ModalSubmitInteraction) {
    const yesterday = interaction.fields.getTextInputValue('standup_yesterday');
    const today = interaction.fields.getTextInputValue('standup_today');
    const blockers = interaction.fields.getTextInputValue('standup_blockers');

    const embed = new EmbedBuilder()
      .setAuthor({ name: interaction.user.username, iconURL: interaction.user.displayAvatarURL() })
      .setTitle('📝 Standup Report')
      .setColor('#2ecc71')
      .addFields(
        { name: '🔄 Вчора', value: yesterday },
        { name: '🎯 Сьогодні', value: today },
        { name: '🛑 Блокери', value: blockers || 'Немає' }
      )
      .setTimestamp();

    await interaction.reply({ embeds: [embed] });
  }
}
