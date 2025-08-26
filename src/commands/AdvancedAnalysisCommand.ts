/**
 * 🧠 Команда розширеного аналізу документів
 * Advanced Document Analysis Command
 */

import {
  SlashCommandBuilder,
  SlashCommandStringOption,
  SlashCommandBooleanOption,
  ChatInputCommandInteraction,
  EmbedBuilder,
  AttachmentBuilder
} from 'discord.js';
import { BaseCommand } from './BaseCommand';
import type { BotConfig, CommandExecuteOptions } from '@/types';
import type { AdvancedDocumentAnalyzer } from '@/services/AdvancedDocumentAnalyzer';
import type { IntelligentWorkflowOrchestrator } from '@/services/IntelligentWorkflowOrchestrator';
import type { GoogleService } from '@/services/GoogleService';
import { t } from '@/i18n';
import logger from '@/utils/logger';

export class AdvancedAnalysisCommand extends BaseCommand {
  constructor(
    config: BotConfig,
    private documentAnalyzer?: AdvancedDocumentAnalyzer,
    private workflowOrchestrator?: IntelligentWorkflowOrchestrator,
    private googleService?: GoogleService
  ) {
    super('advanced-analysis', 'Розширений аналіз документів з AI', config, {
      i18n: { nameKey: 'commands.advanced_analysis.name', descriptionKey: 'commands.advanced_analysis.description' }
    }, (builder: SlashCommandBuilder): SlashCommandBuilder => {
      
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('file_id')
          .setDescription('ID файлу Google Drive для аналізу')
          .setRequired(true)
      );

      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('analysis_type')
          .setDescription('Тип аналізу')
          .setRequired(false)
          .addChoices(
            { name: 'Повний аналіз', value: 'full' },
            { name: 'Тільки сутності', value: 'entities' },
            { name: 'Відповідність вимогам', value: 'compliance' },
            { name: 'Оцінка ризиків', value: 'risk' },
            { name: 'Аналіз настрою', value: 'sentiment' }
          )
      );

      builder.addBooleanOption((option: SlashCommandBooleanOption) =>
        option
          .setName('start_workflow')
          .setDescription('Запустити автоматичний робочий процес')
          .setRequired(false)
      );

      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('language')
          .setDescription('Мова аналізу')
          .setRequired(false)
          .addChoices(
            { name: 'Українська', value: 'uk' },
            { name: 'English', value: 'en' }
          )
      );

      return builder;
    });
  }

  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;

    if (!this.documentAnalyzer || !this.googleService) {
      await interaction.reply({
        content: '❌ Сервіс аналізу документів недоступний',
        ephemeral: true
      });
      return;
    }

    try {
      await interaction.deferReply();

      const fileId = interaction.options.getString('file_id', true);
      const analysisType = interaction.options.getString('analysis_type') || 'full';
      const startWorkflow = interaction.options.getBoolean('start_workflow') || false;
      const language = interaction.options.getString('language') as 'uk' | 'en' || 'uk';

      // Перевірка доступу до файлу
      try {
        await this.googleService.getFileInfo(fileId);
      } catch (error) {
        await interaction.editReply({
          content: '❌ Файл не знайдено або немає доступу'
        });
        return;
      }

      // Визначення опцій аналізу
      const analysisOptions = this.getAnalysisOptions(analysisType, language);

      // Виконання аналізу
      const analysis = await this.documentAnalyzer.analyzeDocument(fileId, analysisOptions);

      // Створення звіту
      const report = await this.documentAnalyzer.generateAnalysisReport(analysis, language);

      // Створення embed відповіді
      const embed = this.createAnalysisEmbed(analysis, fileId, language);

      // Запуск робочого процесу якщо потрібно
      let workflowId: string | undefined;
      if (startWorkflow && this.workflowOrchestrator) {
        try {
          workflowId = await this.workflowOrchestrator.processDocument({
            fileId,
            userId: interaction.user.id,
            channelId: interaction.channelId!,
            documentType: analysis.documentType,
            urgency: analysis.urgencyLevel,
            metadata: { analysis }
          });
        } catch (workflowError) {
          logger.warn('Помилка запуску робочого процесу', {
            component: 'AdvancedAnalysisCommand',
            fileId,
            error: workflowError
          });
        }
      }

      // Додавання інформації про робочий процес
      if (workflowId) {
        embed.addFields({
          name: '🔄 Робочий процес',
          value: `Запущено автоматичний процес: \`${workflowId}\``,
          inline: false
        });
      }

      // Створення файлу з детальним звітом
      const reportAttachment = new AttachmentBuilder(
        Buffer.from(report, 'utf-8'),
        { name: `analysis_report_${fileId}.txt` }
      );

      await interaction.editReply({
        embeds: [embed],
        files: [reportAttachment]
      });

      logger.info('Розширений аналіз завершено', {
        component: 'AdvancedAnalysisCommand',
        fileId,
        analysisType,
        userId: interaction.user.id,
        workflowStarted: !!workflowId
      });

    } catch (error) {
      logger.error('Помилка розширеного аналізу', {
        component: 'AdvancedAnalysisCommand',
        userId: interaction.user.id,
        error: error instanceof Error ? error.message : String(error)
      });

      await interaction.editReply({
        content: '❌ Помилка під час аналізу документа'
      });
    }
  }

  /**
   * Отримання опцій аналізу на основі типу
   */
  private getAnalysisOptions(analysisType: string, language: 'uk' | 'en'): any {
    const baseOptions = { language };

    switch (analysisType) {
      case 'entities':
        return {
          ...baseOptions,
          includeEntities: true,
          includeRelationships: false,
          includeCompliance: false,
          includeSentiment: false,
          includeRiskAssessment: false
        };
      
      case 'compliance':
        return {
          ...baseOptions,
          includeEntities: false,
          includeRelationships: false,
          includeCompliance: true,
          includeSentiment: false,
          includeRiskAssessment: false
        };
      
      case 'risk':
        return {
          ...baseOptions,
          includeEntities: false,
          includeRelationships: false,
          includeCompliance: false,
          includeSentiment: false,
          includeRiskAssessment: true
        };
      
      case 'sentiment':
        return {
          ...baseOptions,
          includeEntities: false,
          includeRelationships: false,
          includeCompliance: false,
          includeSentiment: true,
          includeRiskAssessment: false
        };
      
      case 'full':
      default:
        return {
          ...baseOptions,
          includeEntities: true,
          includeRelationships: true,
          includeCompliance: true,
          includeSentiment: true,
          includeRiskAssessment: true
        };
    }
  }

  /**
   * Створення embed з результатами аналізу
   */
  private createAnalysisEmbed(analysis: any, fileId: string, language: 'uk' | 'en'): EmbedBuilder {
    const embed = new EmbedBuilder()
      .setTitle(language === 'uk' ? '🧠 Розширений аналіз документа' : '🧠 Advanced Document Analysis')
      .setDescription(analysis.summary)
      .setColor(this.getUrgencyColor(analysis.urgencyLevel))
      .setTimestamp()
      .addFields(
        {
          name: language === 'uk' ? '📄 Тип документа' : '📄 Document Type',
          value: this.translateDocumentType(analysis.documentType, language),
          inline: true
        },
        {
          name: language === 'uk' ? '⚡ Терміновість' : '⚡ Urgency',
          value: this.translateUrgency(analysis.urgencyLevel, language),
          inline: true
        },
        {
          name: language === 'uk' ? '🔗 ID файлу' : '🔗 File ID',
          value: `\`${fileId}\``,
          inline: true
        }
      );

    // Ключові теми
    if (analysis.keyTopics && analysis.keyTopics.length > 0) {
      embed.addFields({
        name: language === 'uk' ? '🔑 Ключові теми' : '🔑 Key Topics',
        value: analysis.keyTopics.slice(0, 5).map((topic: string, i: number) => `${i + 1}. ${topic}`).join('\n'),
        inline: false
      });
    }

    // Дії до виконання
    if (analysis.actionItems && analysis.actionItems.length > 0) {
      const actions = analysis.actionItems.slice(0, 3).map((item: any, i: number) => 
        `${i + 1}. ${item.action} (${this.translatePriority(item.priority, language)})`
      ).join('\n');
      
      embed.addFields({
        name: language === 'uk' ? '✅ Дії до виконання' : '✅ Action Items',
        value: actions,
        inline: false
      });
    }

    // Сутності
    if (analysis.entities && analysis.entities.length > 0) {
      const entities = analysis.entities.slice(0, 5).map((entity: any) => 
        `• ${entity.type}: ${entity.value} (${Math.round(entity.confidence * 100)}%)`
      ).join('\n');
      
      embed.addFields({
        name: language === 'uk' ? '👥 Виявлені сутності' : '👥 Detected Entities',
        value: entities,
        inline: false
      });
    }

    // Відповідність вимогам
    if (analysis.compliance) {
      embed.addFields({
        name: language === 'uk' ? '⚖️ Відповідність вимогам' : '⚖️ Compliance',
        value: `${analysis.compliance.score}%`,
        inline: true
      });
    }

    // Оцінка ризиків
    if (analysis.riskAssessment) {
      embed.addFields({
        name: language === 'uk' ? '⚠️ Рівень ризику' : '⚠️ Risk Level',
        value: this.translateRiskLevel(analysis.riskAssessment.level, language),
        inline: true
      });
    }

    // Читабельність
    if (analysis.readability) {
      embed.addFields({
        name: language === 'uk' ? '📈 Читабельність' : '📈 Readability',
        value: `${analysis.readability.score}%`,
        inline: true
      });
    }

    return embed;
  }

  /**
   * Допоміжні методи перекладу
   */
  private getUrgencyColor(urgency: string): number {
    const colors = {
      critical: 0xff0000, // червоний
      high: 0xff8800,     // помаранчевий
      medium: 0xffff00,   // жовтий
      low: 0x00ff00       // зелений
    };
    return colors[urgency as keyof typeof colors] || 0x808080;
  }

  private translateDocumentType(type: string, language: 'uk' | 'en'): string {
    if (language === 'en') {
      return type.replace(/_/g, ' ').toUpperCase();
    }
    
    const translations: Record<string, string> = {
      'military_order': 'Військовий наказ',
      'administrative_doc': 'Адміністративний документ',
      'legal_contract': 'Юридичний договір',
      'financial_report': 'Фінансовий звіт',
      'technical_spec': 'Технічна специфікація',
      'communication': 'Листування',
      'other': 'Інший'
    };
    return translations[type] || type;
  }

  private translateUrgency(urgency: string, language: 'uk' | 'en'): string {
    if (language === 'en') {
      return urgency.toUpperCase();
    }
    
    const translations: Record<string, string> = {
      'critical': 'Критична',
      'high': 'Висока',
      'medium': 'Середня',
      'low': 'Низька'
    };
    return translations[urgency] || urgency;
  }

  private translatePriority(priority: string, language: 'uk' | 'en'): string {
    if (language === 'en') {
      return priority.toUpperCase();
    }
    
    const translations: Record<string, string> = {
      'high': 'Висока',
      'medium': 'Середня',
      'low': 'Низька'
    };
    return translations[priority] || priority;
  }

  private translateRiskLevel(level: string, language: 'uk' | 'en'): string {
    if (language === 'en') {
      return level.toUpperCase();
    }
    
    const translations: Record<string, string> = {
      'critical': 'Критичний',
      'high': 'Високий',
      'medium': 'Середній',
      'low': 'Низький'
    };
    return translations[level] || level;
  }
}