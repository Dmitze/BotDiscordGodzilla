/**
 * 🤖 Інтелектуальний оркестратор робочих процесів документів
 * Intelligent Document Workflow Orchestrator
 */

import type { AIService } from './AIService';

import { AdvancedDocumentAnalyzer } from './AdvancedDocumentAnalyzer';
import logger from '@/utils/logger';

interface WorkflowRule {
  id: string;
  name: string;
  description: string;
  condition: string; // AI-evaluated condition
  actions: WorkflowAction[];
  priority: number;
  enabled: boolean;
}

interface WorkflowAction {
  type: 'notify' | 'route' | 'analyze' | 'approve' | 'archive' | 'escalate' | 'custom';
  config: Record<string, any>;
  delay?: number; // milliseconds
}

interface ProcessingContext {
  fileId: string;
  userId: string;
  channelId: string;
  documentType?: string;
  urgency?: string;
  metadata: Record<string, any>;
}

interface WorkflowExecution {
  id: string;
  workflowId: string;
  context: ProcessingContext;
  status: 'pending' | 'running' | 'completed' | 'failed' | 'paused';
  currentStep: number;
  steps: WorkflowStep[];
  startTime: Date;
  endTime?: Date;
  results: Record<string, any>;
  errors: string[];
}

interface WorkflowStep {
  action: WorkflowAction;
  status: 'pending' | 'completed' | 'failed' | 'skipped';
  result?: any;
  error?: string;
  executedAt?: Date;
}

export class IntelligentWorkflowOrchestrator {
  private rules = new Map<string, WorkflowRule>();
  private executions = new Map<string, WorkflowExecution>();

  constructor(
    private aiService: AIService,
    private documentAnalyzer: AdvancedDocumentAnalyzer
  ) {
    this.initializeDefaultRules();
  }

  /**
   * 🚀 Запуск обробки документа
   */
  async processDocument(context: ProcessingContext): Promise<string> {
    const executionId = this.generateExecutionId();
    
    try {
      // Аналіз документа для визначення типу та терміновості
      const analysis = await this.documentAnalyzer.analyzeDocument(context.fileId);
      
      context.documentType = analysis.documentType;
      context.urgency = analysis.urgencyLevel;
      context.metadata['analysis'] = analysis;

      // Знаходження відповідних правил
      const applicableRules = await this.findApplicableRules(context);
      
      // Створення виконання
      const execution: WorkflowExecution = {
        id: executionId,
        workflowId: applicableRules[0]?.id || 'default',
        context,
        status: 'pending',
        currentStep: 0,
        steps: this.createStepsFromRules(applicableRules),
        startTime: new Date(),
        results: {},
        errors: []
      };

      this.executions.set(executionId, execution);
      
      // Запуск виконання
      this.executeWorkflow(executionId);

      logger.info('Документ додано до обробки', {
        component: 'IntelligentWorkflowOrchestrator',
        executionId,
        fileId: context.fileId,
        documentType: context.documentType,
        urgency: context.urgency
      });

      return executionId;

    } catch (error) {
      logger.error('Помилка запуску обробки документа', {
        component: 'IntelligentWorkflowOrchestrator',
        fileId: context.fileId,
        error: error instanceof Error ? error.message : String(error)
      });
      throw error;
    }
  }

  /**
   * 📋 Знаходження застосовних правил
   */
  private async findApplicableRules(context: ProcessingContext): Promise<WorkflowRule[]> {
    const applicableRules: WorkflowRule[] = [];

    for (const rule of this.rules.values()) {
      if (!rule.enabled) continue;

      const isApplicable = await this.evaluateCondition(rule.condition, context);
      if (isApplicable) {
        applicableRules.push(rule);
      }
    }

    // Сортування за пріоритетом
    return applicableRules.sort((a, b) => b.priority - a.priority);
  }

  /**
   * ⚡ Виконання робочого процесу
   */
  private async executeWorkflow(executionId: string): Promise<void> {
    const execution = this.executions.get(executionId);
    if (!execution) return;

    execution.status = 'running';

    try {
      for (let i = execution.currentStep; i < execution.steps.length; i++) {
        const step = execution.steps[i];
        if (!step) continue;
        
        execution.currentStep = i;

        try {
          step.status = 'pending';
          const result = await this.executeAction(step.action, execution.context);
          
          step.status = 'completed';
          step.result = result;
          step.executedAt = new Date();

          // Затримка якщо потрібно
          if (step.action.delay) {
            await new Promise(resolve => setTimeout(resolve, step.action.delay));
          }

        } catch (error) {
          step.status = 'failed';
          step.error = error instanceof Error ? error.message : String(error);
          execution.errors.push(step.error);
          
          logger.warn('Крок виконання провалився', {
            component: 'IntelligentWorkflowOrchestrator',
            executionId,
            step: i,
            action: step.action.type,
            error: step.error
          });
        }
      }

      execution.status = 'completed';
      execution.endTime = new Date();

    } catch (error) {
      execution.status = 'failed';
      execution.errors.push(error instanceof Error ? error.message : String(error));
      
      logger.error('Виконання робочого процесу провалилося', {
        component: 'IntelligentWorkflowOrchestrator',
        executionId,
        error
      });
    }
  }

  /**
   * 🎯 Виконання дії
   */
  private async executeAction(action: WorkflowAction, context: ProcessingContext): Promise<any> {
    switch (action.type) {
      case 'analyze':
        return await this.executeAnalyzeAction(action, context);
      
      case 'notify':
        return await this.executeNotifyAction(action, context);
      
      case 'route':
        return await this.executeRouteAction(action, context);
      
      case 'approve':
        return await this.executeApproveAction(action, context);
      
      case 'escalate':
        return await this.executeEscalateAction(action, context);
      
      case 'custom':
        return await this.executeCustomAction(action, context);
        
      default:
        throw new Error(`Невідомий тип дії: ${action.type}`);
    }
  }

  /**
   * 📊 Виконання аналізу
   */
  private async executeAnalyzeAction(action: WorkflowAction, context: ProcessingContext): Promise<any> {
    const analysisType = action.config['type'] || 'full';
    
    return await this.documentAnalyzer.analyzeDocument(context.fileId, {
      includeEntities: analysisType === 'full' || analysisType === 'entities',
      includeCompliance: analysisType === 'full' || analysisType === 'compliance',
      includeRiskAssessment: analysisType === 'full' || analysisType === 'risk',
      language: action.config['language'] || 'uk'
    });
  }

  /**
   * 📢 Виконання сповіщення
   */
  private async executeNotifyAction(action: WorkflowAction, context: ProcessingContext): Promise<any> {
    // Імплементація сповіщень через Discord
    logger.info('Відправлення сповіщення', {
      component: 'IntelligentWorkflowOrchestrator',
      channelId: context.channelId,
      message: action.config['message']
    });

    return { notified: true, timestamp: new Date() };
  }

  /**
   * 🔄 Виконання маршрутизації
   */
  private async executeRouteAction(action: WorkflowAction, context: ProcessingContext): Promise<any> {
    const targetChannel = action.config['targetChannel'];
    const targetUser = action.config['targetUser'];
    
    logger.info('Маршрутизація документа', {
      component: 'IntelligentWorkflowOrchestrator',
      fileId: context.fileId,
      targetChannel,
      targetUser
    });

    return { routed: true, target: { channel: targetChannel, user: targetUser } };
  }

  /**
   * ✅ Виконання затвердження
   */
  private async executeApproveAction(action: WorkflowAction, _context: ProcessingContext): Promise<any> {
    // Логіка автоматичного або ручного затвердження
    return { approved: false, pending: true, requiredRole: action.config['requiredRole'] };
  }

  /**
   * ⬆️ Виконання ескалації
   */
  private async executeEscalateAction(action: WorkflowAction, context: ProcessingContext): Promise<any> {
    const escalationLevel = action.config['level'] || 'manager';
    
    logger.warn('Ескалація документа', {
      component: 'IntelligentWorkflowOrchestrator',
      fileId: context.fileId,
      level: escalationLevel,
      reason: action.config['reason']
    });

    return { escalated: true, level: escalationLevel };
  }

  /**
   * 🔧 Виконання кастомної дії
   */
  private async executeCustomAction(action: WorkflowAction, context: ProcessingContext): Promise<any> {
    if (action.config['aiPrompt']) {
      const prompt = this.replaceVariables(action.config['aiPrompt'], context);
      
      const response = await this.aiService.generateResponse(prompt, {
        temperature: 0.3,
        maxTokens: 500
      });

      return { customResult: response.content };
    }

    return { executed: true };
  }

  /**
   * 🧠 Оцінка умови за допомогою AI
   */
  private async evaluateCondition(condition: string, context: ProcessingContext): Promise<boolean> {
    try {
      // Простий інтерпретатор умов
      const evaluationPrompt = `
Оціни умову для документа:

Умова: ${condition}

Контекст документа:
- Тип: ${context.documentType}
- Терміновість: ${context.urgency}
- Назва файлу: ${context.metadata['name'] || 'Невідома'}

Поверни лише "true" або "false":
`;

      const response = await this.aiService.generateResponse(evaluationPrompt, {
        temperature: 0.1,
        maxTokens: 10
      });

      return response.content.toLowerCase().includes('true');

    } catch (error) {
      logger.warn('Помилка оцінки умови', {
        component: 'IntelligentWorkflowOrchestrator',
        condition,
        error
      });
      return false;
    }
  }

  /**
   * 🔧 Ініціалізація стандартних правил
   */
  private initializeDefaultRules(): void {
    // Правило для критичних документів
    this.rules.set('critical_documents', {
      id: 'critical_documents',
      name: 'Критичні документи',
      description: 'Обробка документів з критичною терміновістю',
      condition: 'urgency === "critical"',
      actions: [
        {
          type: 'notify',
          config: {
            message: '🚨 Критичний документ потребує негайної уваги!',
            priority: 'high'
          }
        },
        {
          type: 'escalate',
          config: {
            level: 'commander',
            reason: 'Критична терміновість'
          }
        }
      ],
      priority: 100,
      enabled: true
    });

    // Правило для операційних наказів
    this.rules.set('business_contracts', {
      id: 'business_contracts',
      name: 'операційні накази',
      description: 'Обробка операційних наказів',
      condition: 'documentType === "business_contract"',
      actions: [
        {
          type: 'analyze',
          config: {
            type: 'compliance',
            language: 'uk'
          }
        },
        {
          type: 'route',
          config: {
            targetChannel: 'business_contracts',
            reason: 'операційний наказ'
          }
        }
      ],
      priority: 80,
      enabled: true
    });

    // Правило для фінансових документів
    this.rules.set('financial_documents', {
      id: 'financial_documents', 
      name: 'Фінансові документи',
      description: 'Обробка фінансових звітів та документів',
      condition: 'documentType === "financial_report"',
      actions: [
        {
          type: 'analyze',
          config: {
            type: 'risk',
            language: 'uk'
          }
        },
        {
          type: 'approve',
          config: {
            requiredRole: 'financial_officer',
            autoApprove: false
          }
        }
      ],
      priority: 70,
      enabled: true
    });
  }

  /**
   * 🔧 Допоміжні методи
   */
  private createStepsFromRules(rules: WorkflowRule[]): WorkflowStep[] {
    const steps: WorkflowStep[] = [];

    for (const rule of rules) {
      for (const action of rule.actions) {
        steps.push({
          action,
          status: 'pending'
        });
      }
    }

    return steps;
  }

  private replaceVariables(text: string, context: ProcessingContext): string {
    return text
      .replace(/\{fileId\}/g, context.fileId)
      .replace(/\{documentType\}/g, context.documentType || 'невідомий')
      .replace(/\{urgency\}/g, context.urgency || 'невідома');
  }

  private generateExecutionId(): string {
    return `exec_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * 📊 Отримання статусу виконання
   */
  getExecutionStatus(executionId: string): WorkflowExecution | null {
    return this.executions.get(executionId) || null;
  }

  /**
   * ➕ Додавання кастомного правила
   */
  addWorkflowRule(rule: WorkflowRule): void {
    this.rules.set(rule.id, rule);
    
    logger.info('Додано правило робочого процесу', {
      component: 'IntelligentWorkflowOrchestrator',
      ruleId: rule.id,
      name: rule.name
    });
  }

  /**
   * 📋 Отримання всіх активних виконань
   */
  getActiveExecutions(): WorkflowExecution[] {
    return Array.from(this.executions.values())
      .filter(execution => execution.status === 'running' || execution.status === 'pending');
  }
}