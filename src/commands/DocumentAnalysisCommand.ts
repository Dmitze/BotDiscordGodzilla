/**
 * Document Analysis Command for Discord AI Assistant Bot
 * Provides comprehensive analysis of Google Documents
 */

import type { BotConfig, CommandExecuteOptions } from '@/types';
import { BaseCommand } from './BaseCommand';
import logger from '@/utils/logger';

import type { GoogleService } from '@/services/GoogleService';
import type { DocumentAnalysisService } from '@/services/DocumentAnalysisService';
import type {
  SlashCommandBuilder,
  SlashCommandStringOption,
} from 'discord.js';
import { t } from '@/i18n';

type DocumentAnalysisOptions = {
  file: string | null;
  analysisType: string | null;
};

export class DocumentAnalysisCommand extends BaseCommand {
  private readonly googleService: GoogleService | undefined;
  private readonly documentAnalysisService: DocumentAnalysisService | undefined;

  constructor(
    config: BotConfig, 
    googleService?: GoogleService,
    documentAnalysisService?: DocumentAnalysisService
  ) {
    super('analyze-doc', t('document.analysis.command.description'), config, {
      i18n: { nameKey: 'commands.analyze-doc.name', descriptionKey: 'document.analysis.command.description' }
    }, (builder: SlashCommandBuilder): SlashCommandBuilder => {
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('file')
          .setDescription(t('document.analysis.opt.file.description'))
          .setRequired(true)
          .setMaxLength(100)
      );
      builder.addStringOption((option: SlashCommandStringOption) =>
        option
          .setName('type')
          .setDescription(t('document.analysis.opt.type.description'))
          .setRequired(false)
          .addChoices(
            { name: 'Full Analysis', value: 'full' },
            { name: 'Structure Only', value: 'structure' },
            { name: 'Summary Only', value: 'summary' },
            { name: 'Action Items Only', value: 'actions' },
            { name: 'Compliance Check', value: 'compliance' },
            { name: 'Quality Assessment', value: 'quality' }
          )
      );
      return builder;
    });

    this.googleService = googleService;
    this.documentAnalysisService = documentAnalysisService;
  }

  /**
   * Execute the document analysis command
   */
  protected async onExecute(options: CommandExecuteOptions): Promise<void> {
    const { interaction } = options;
    
    try {
      // Defer reply as this might take some time
      await interaction.deferReply({ ephemeral: false });

      // Get command options
      const opts: DocumentAnalysisOptions = {
        file: interaction.options.getString('file'),
        analysisType: interaction.options.getString('type') || 'full'
      };

      // Validate options
      if (!opts.file) {
        await interaction.editReply({
          content: t('document.analysis.error.no_file')
        });
        return;
      }

      // Validate services
      if (!this.googleService) {
        await interaction.editReply({
          content: t('document.analysis.error.google_service')
        });
        return;
      }

      if (!this.documentAnalysisService) {
        await interaction.editReply({
          content: t('document.analysis.error.analysis_service')
        });
        return;
      }

      // Find the file in Google Drive
      const searchResult = await this.googleService.searchFiles(`name contains '${opts.file}'`);
      
      if (!searchResult || searchResult.length === 0) {
        await interaction.editReply({
          content: t('document.analysis.error.file_not_found', { fileName: opts.file })
        });
        return;
      }

      // Get the first file's ID and convert to DriveFile
      const fileId = searchResult[0]?.id;
      if (!fileId) {
        await interaction.editReply({
          content: t('document.analysis.error.file_no_id', { fileName: opts.file })
        });
        return;
      }

      // Convert Schema$File to DriveFile using the public getDriveFile method
      const file = await this.googleService.getDriveFile(fileId);
      
      // Ensure the file has an id
      if (!file || !file.id) {
        await interaction.editReply({
          content: t('document.analysis.error.file_no_id', { fileName: opts.file })
        });
        return;
      }

      // Prepare analysis options based on type
      const analysisOptions: any = {};
      
      switch (opts.analysisType) {
        case 'structure':
          analysisOptions.includeStructure = true;
          break;
        case 'summary':
          analysisOptions.includeSummary = true;
          break;
        case 'actions':
          analysisOptions.includeActionItems = true;
          break;
        case 'compliance':
          analysisOptions.includeCompliance = true;
          break;
        case 'quality':
          analysisOptions.includeQuality = true;
          break;
        case 'full':
        default:
          // For full analysis, include all options
          analysisOptions.includeStructure = true;
          analysisOptions.includeSummary = true;
          analysisOptions.includeActionItems = true;
          analysisOptions.includeCompliance = true;
          analysisOptions.includeQuality = true;
          break;
      }

      // Perform analysis
      const analysisResult = await this.documentAnalysisService.analyzeDocument(file, analysisOptions);

      // Format the result based on analysis type
      let formattedResult: string;
      
      switch (opts.analysisType) {
        case 'structure':
          formattedResult = this.formatStructureAnalysis(analysisResult);
          break;
        case 'summary':
          formattedResult = this.formatSummaryAnalysis(analysisResult);
          break;
        case 'actions':
          formattedResult = this.formatActionItemsAnalysis(analysisResult);
          break;
        case 'compliance':
          formattedResult = this.formatComplianceAnalysis(analysisResult);
          break;
        case 'quality':
          formattedResult = this.formatQualityAnalysis(analysisResult);
          break;
        case 'full':
        default:
          formattedResult = this.formatFullAnalysis(analysisResult);
          break;
      }

      // Send the analysis result
      await interaction.editReply({
        content: formattedResult
      });

      // Log the command execution
      logger.info('Document analysis command executed', {
        component: 'DocumentAnalysisCommand',
        userId: interaction.user.id,
        fileId: file.id,
        analysisType: opts.analysisType
      });

    } catch (error) {
      logger.error('Error executing document analysis command', {
        component: 'DocumentAnalysisCommand',
        error: error instanceof Error ? error.message : String(error)
      });

      await interaction.editReply({
        content: t('document.analysis.error.analysis_failed')
      });
    }
  }

  /**
   * Format structure analysis result
   */
  private formatStructureAnalysis(analysis: any): string {
    return t('document.analysis.result.structure', {
      fileName: analysis.fileName,
      // In a real implementation, we would format the actual analysis data
      structure: 'Document structure analysis results would be shown here'
    });
  }

  /**
   * Format summary analysis result
   */
  private formatSummaryAnalysis(analysis: any): string {
    return t('document.analysis.result.summary', {
      fileName: analysis.fileName,
      // In a real implementation, we would format the actual analysis data
      summary: 'Document summary would be shown here'
    });
  }

  /**
   * Format action items analysis result
   */
  private formatActionItemsAnalysis(analysis: any): string {
    return t('document.analysis.result.actions', {
      fileName: analysis.fileName,
      // In a real implementation, we would format the actual analysis data
      actions: 'Document action items would be shown here'
    });
  }

  /**
   * Format compliance analysis result
   */
  private formatComplianceAnalysis(analysis: any): string {
    return t('document.analysis.result.compliance', {
      fileName: analysis.fileName,
      // In a real implementation, we would format the actual analysis data
      compliance: 'Document compliance check results would be shown here'
    });
  }

  /**
   * Format quality analysis result
   */
  private formatQualityAnalysis(analysis: any): string {
    return t('document.analysis.result.quality', {
      fileName: analysis.fileName,
      // In a real implementation, we would format the actual analysis data
      quality: 'Document quality assessment would be shown here'
    });
  }

  /**
   * Format full analysis result
   */
  private formatFullAnalysis(analysis: any): string {
    return t('document.analysis.result.full', {
      fileName: analysis.fileName,
      // In a real implementation, we would format the actual analysis data
      analysis: 'Full document analysis would be shown here'
    });
  }
}