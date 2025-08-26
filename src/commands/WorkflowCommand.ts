/**
 * Команда управління робочими процесами
 * Enhanced Workflow Management Command
 */

import {
  SlashCommandBuilder,
  SlashCommandStringOption,
  SlashCommandSubcommandBuilder,
  ChatInputCommandInteraction,
  EmbedBuilder,
  ButtonBuilder,
  ActionRowBuilder,
  ButtonStyle,
} from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { WorkflowAutomationEngine } from '@/services/WorkflowAutomationEngine';

import logger from '@/utils/logger';

export class WorkflowCommand extends BaseCommand {
  private workflowEngine: WorkflowAutomationEngine;

  constructor(config: BotConfig, workflowEngine: WorkflowAutomationEngine) {
    super('workflow', 'Управління автоматизованими робочими процесами', config, {
      i18n: { nameKey: 'commands.workflow.name', descriptionKey: 'commands.workflow.description' }
    }, (builder: SlashCommandBuilder): SlashCommandBuilder => {
      
      // Підкоманда запуску робочого процесу
      builder.addSubcommand((subcommand: SlashCommandSubcommandBuilder) =>
        subcommand
          .setName('start')
          .setDescription('Запустити робочий процес')
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('workflow_id')
              .setDescription('ID робочого процесу')
              .setRequired(true)
              .addChoices(
                { name: 'Обробка документів', value: 'document_intake' },
                { name: 'Процес затвердження', value: 'approval_process' },
                { name: 'Плановий аналіз', value: 'scheduled_analysis' }
              )
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('file_id')
              .setDescription('ID файлу Google Drive (для документів)')
              .setRequired(false)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('parameters')
              .setDescription('Додаткові параметри у форматі JSON')
              .setRequired(false)
          )
      );

      // Підкоманда статусу
      builder.addSubcommand((subcommand: SlashCommandSubcommandBuilder) =>
        subcommand
          .setName('status')
          .setDescription('Переглянути статус робочого процесу')
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('instance_id')
              .setDescription('ID інстансу робочого процесу')
              .setRequired(true)
          )
      );

      // Підкоманда списку активних процесів
      builder.addSubcommand((subcommand: SlashCommandSubcommandBuilder) =>
        subcommand
          .setName('list')
          .setDescription('Показати активні робочі процеси')
      );

      // Підкоманда створення кастомного процесу
      builder.addSubcommand((subcommand: SlashCommandSubcommandBuilder) =>
        subcommand
          .setName('create')
          .setDescription('Створити кастомний робочий процес')
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('name')
              .setDescription('Назва робочого процесу')
              .setRequired(true)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('description')
              .setDescription('Опис робочого процесу')
              .setRequired(true)
          )
          .addStringOption((option: SlashCommandStringOption) =>
            option
              .setName('steps')
              .setDescription('Кроки робочого процесу у форматі JSON')
              .setRequired(true)
          )
      );

      return builder;
    });

    this.workflowEngine = workflowEngine;
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    if (!this.workflowEngine) {
      await interaction.reply({
        content: '❌ Сервіс робочих процесів недоступний',
        ephemeral: true
      });
      return;
    }

    try {
      await interaction.deferReply({ ephemeral: true });

      const subcommand = interaction.options.getSubcommand();

      switch (subcommand) {
        case 'start':
          await this.handleStartWorkflow(interaction);
          break;
        case 'status':
          await this.handleWorkflowStatus(interaction);
          break;
        case 'list':
          await this.handleListWorkflows(interaction);
          break;
        case 'create':
          await this.handleCreateWorkflow(interaction);
          break;
        default:
          await interaction.editReply({
            content: '❌ Невідома підкоманда'
          });
      }
    } catch (error) {
      logger.error('Помилка команди робочих процесів', {
        component: 'WorkflowCommand',
        subcommand: interaction.options.getSubcommand(),
        error: error instanceof Error ? error.message : String(error)
      });

      await interaction.editReply({
        content: '❌ Помилка виконання команди робочих процесів'
      });
    }
  }

  /**
   * Запуск робочого процесу
   */
  private async handleStartWorkflow(interaction: ChatInputCommandInteraction): Promise<void> {
    const workflowId = interaction.options.getString('workflow_id', true);
    const fileId = interaction.options.getString('file_id');
    const parametersStr = interaction.options.getString('parameters');

    let parameters: Record<string, any> = {};

    if (parametersStr) {
      try {
        parameters = JSON.parse(parametersStr);
      } catch (error) {
        await interaction.editReply({
          content: '❌ Невірний формат JSON для параметрів'
        });
        return;
      }
    }

    if (fileId) {
      parameters['fileId'] = fileId;
    }

    parameters['triggeredBy'] = interaction.user.id;
    parameters['channelId'] = interaction.channelId;

    try {
      const instanceId = await this.workflowEngine!.startWorkflow(
        workflowId,
        parameters,
        interaction.user.tag
      );

      const embed = new EmbedBuilder()
        .setTitle('🚀 Робочий процес запущено')
        .setDescription(`Робочий процес **${this.getWorkflowName(workflowId)}** успішно запущено`)
        .addFields(
          { name: 'ID інстансу', value: `\`${instanceId}\``, inline: true },
          { name: 'ID процесу', value: `\`${workflowId}\``, inline: true },
          { name: 'Запущено', value: `<@${interaction.user.id}>`, inline: true }
        )
        .setColor(0x00ff00)
        .setTimestamp();

      // Додаємо кнопку для перегляду статусу
      const statusButton = new ButtonBuilder()
        .setCustomId(`workflow_status_${instanceId}`)
        .setLabel('Переглянути статус')
        .setStyle(ButtonStyle.Secondary)
        .setEmoji('📊');

      const row = new ActionRowBuilder<ButtonBuilder>()
        .addComponents(statusButton);

      await interaction.editReply({
        embeds: [embed],
        components: [row]
      });

    } catch (error) {
      await interaction.editReply({
        content: `❌ Помилка запуску робочого процесу: ${error instanceof Error ? error.message : String(error)}`
      });
    }
  }

  /**
   * Перегляд статусу робочого процесу
   */
  private async handleWorkflowStatus(interaction: ChatInputCommandInteraction): Promise<void> {
    const instanceId = interaction.options.getString('instance_id', true);

    const instance = this.workflowEngine!.getWorkflowStatus(instanceId);
    if (!instance) {
      await interaction.editReply({
        content: '❌ Робочий процес не знайдено'
      });
      return;
    }

    const embed = new EmbedBuilder()
      .setTitle('📊 Статус робочого процесу')
      .setDescription(`Інстанс: \`${instanceId}\``)
      .addFields(
        { name: 'Статус', value: this.getStatusEmoji(instance.status) + ' ' + this.translateStatus(instance.status), inline: true },
        { name: 'Поточний крок', value: `\`${instance.currentStep}\``, inline: true },
        { name: 'Створено', value: `<t:${Math.floor(instance.createdAt.getTime() / 1000)}:R>`, inline: true },
        { name: 'Оновлено', value: `<t:${Math.floor(instance.updatedAt.getTime() / 1000)}:R>`, inline: true }
      )
      .setColor(this.getStatusColor(instance.status))
      .setTimestamp();

    // Додаємо історію виконання
    if (instance.history.length > 0) {
      const historyText = instance.history
        .slice(-5) // Останні 5 записів
        .map(entry => {
          const statusEmoji = entry.status === 'completed' ? '✅' : 
                             entry.status === 'failed' ? '❌' : 
                             entry.status === 'skipped' ? '⏭️' : '🔄';
          return `${statusEmoji} \`${entry.stepId}\` - ${this.translateStatus(entry.status)} <t:${Math.floor(entry.timestamp.getTime() / 1000)}:R>`;
        })
        .join('\n');

      embed.addFields({ name: 'Історія виконання', value: historyText });
    }

    await interaction.editReply({ embeds: [embed] });
  }

  /**
   * Список активних робочих процесів
   */
  private async handleListWorkflows(interaction: ChatInputCommandInteraction): Promise<void> {
    const activeWorkflows = this.workflowEngine!.getActiveWorkflows();

    if (activeWorkflows.length === 0) {
      await interaction.editReply({
        content: '📋 Наразі немає активних робочих процесів'
      });
      return;
    }

    const embed = new EmbedBuilder()
      .setTitle('📋 Активні робочі процеси')
      .setDescription(`Знайдено ${activeWorkflows.length} активних процесів`)
      .setColor(0x3498db)
      .setTimestamp();

    const workflowsList = activeWorkflows
      .slice(0, 10) // Обмежуємо до 10
      .map(workflow => {
        const duration = Date.now() - workflow.createdAt.getTime();
        const hours = Math.floor(duration / (1000 * 60 * 60));
        const minutes = Math.floor((duration % (1000 * 60 * 60)) / (1000 * 60));
        
        return `🔄 \`${workflow.id}\`\n   ${this.getWorkflowName(workflow.workflowId)} • ${hours}г ${minutes}хв\n   Крок: \`${workflow.currentStep}\``;
      })
      .join('\n\n');

    embed.addFields({ name: 'Процеси', value: workflowsList });

    await interaction.editReply({ embeds: [embed] });
  }

  /**
   * Створення кастомного робочого процесу
   */
  private async handleCreateWorkflow(interaction: ChatInputCommandInteraction): Promise<void> {
    const name = interaction.options.getString('name', true);
    const description = interaction.options.getString('description', true);
    const stepsStr = interaction.options.getString('steps', true);

    try {
      const steps = JSON.parse(stepsStr);
      
      const workflowId = `custom_${Date.now()}`;
      const workflow = {
        id: workflowId,
        name,
        description,
        trigger: 'manual' as const,
        steps
      };

      this.workflowEngine!.registerWorkflow(workflow);

      const embed = new EmbedBuilder()
        .setTitle('✅ Робочий процес створено')
        .setDescription(`Кастомний робочий процес **${name}** успішно створено`)
        .addFields(
          { name: 'ID процесу', value: `\`${workflowId}\``, inline: true },
          { name: 'Кроків', value: `${steps.length}`, inline: true },
          { name: 'Автор', value: `<@${interaction.user.id}>`, inline: true }
        )
        .setColor(0x00ff00)
        .setTimestamp();

      await interaction.editReply({ embeds: [embed] });

    } catch (error) {
      await interaction.editReply({
        content: `❌ Помилка створення робочого процесу: ${error instanceof Error ? error.message : String(error)}`
      });
    }
  }

  /**
   * Допоміжні методи
   */
  private getWorkflowName(workflowId: string): string {
    const names = {
      document_intake: 'Обробка документів',
      approval_process: 'Процес затвердження',
      scheduled_analysis: 'Плановий аналіз'
    };
    return names[workflowId as keyof typeof names] || workflowId;
  }

  private translateStatus(status: string): string {
    const translations = {
      running: 'Виконується',
      completed: 'Завершено',
      failed: 'Провалено',
      paused: 'Призупинено',
      started: 'Розпочато',
      skipped: 'Пропущено'
    };
    return translations[status as keyof typeof translations] || status;
  }

  private getStatusEmoji(status: string): string {
    const emojis = {
      running: '🔄',
      completed: '✅',
      failed: '❌',
      paused: '⏸️'
    };
    return emojis[status as keyof typeof emojis] || '❓';
  }

  private getStatusColor(status: string): number {
    const colors = {
      running: 0x3498db,
      completed: 0x00ff00,
      failed: 0xff0000,
      paused: 0xffff00
    };
    return colors[status as keyof typeof colors] || 0x808080;
  }
}