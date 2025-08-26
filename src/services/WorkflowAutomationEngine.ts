/**
 * Двигун автоматизації робочих процесів та документообігу
 * Intelligent Workflow Automation Engine
 */

import type { AIService } from './AIService';
import type { GoogleService } from './GoogleService';
import { EnhancedDocumentService } from './EnhancedDocumentService';
import logger from '@/utils/logger';

export interface WorkflowStep {
  id: string;
  name: string;
  type: 'document_analysis' | 'approval' | 'notification' | 'data_extraction' | 'validation' | 'custom';
  condition?: string; // AI-evaluated condition
  aiPrompt?: string;
  requiredRole?: string[];
  timeoutHours?: number;
  nextSteps?: string[];
  metadata?: Record<string, any>;
}

export interface WorkflowDefinition {
  id: string;
  name: string;
  description: string;
  trigger: 'document_upload' | 'command' | 'schedule' | 'manual';
  steps: WorkflowStep[];
  variables?: Record<string, any>;
}

export interface WorkflowInstance {
  id: string;
  workflowId: string;
  status: 'running' | 'completed' | 'failed' | 'paused';
  currentStep: string;
  variables: Record<string, any>;
  history: WorkflowHistoryEntry[];
  createdAt: Date;
  updatedAt: Date;
  completedAt?: Date;
}

export interface WorkflowHistoryEntry {
  stepId: string;
  status: 'started' | 'completed' | 'failed' | 'skipped';
  result?: any;
  error?: string;
  timestamp: Date;
  executedBy?: string;
}

export class WorkflowAutomationEngine {
  private workflows = new Map<string, WorkflowDefinition>();
  private instances = new Map<string, WorkflowInstance>();

  constructor(
    private aiService: AIService,
    private googleService: GoogleService,
    private documentService: EnhancedDocumentService
  ) {
    this.initializeDefaultWorkflows();
  }

  /**
   * Ініціалізація стандартних робочих процесів
   */
  private initializeDefaultWorkflows(): void {
    // Робочий процес для аналізу нових документів
    this.registerWorkflow({
      id: 'document_intake',
      name: 'Обробка нових документів',
      description: 'Автоматичний аналіз та класифікація нових документів',
      trigger: 'document_upload',
      steps: [
        {
          id: 'analyze_document',
          name: 'Аналіз документа',
          type: 'document_analysis',
          nextSteps: ['classify_urgency']
        },
        {
          id: 'classify_urgency',
          name: 'Класифікація терміновості',
          type: 'custom',
          aiPrompt: 'Визнач рівень терміновості документа: {{document_summary}}',
          nextSteps: ['route_document']
        },
        {
          id: 'route_document',
          name: 'Маршрутизація документа',
          type: 'notification',
          condition: 'urgency === "critical" || urgency === "high"',
          nextSteps: ['create_tasks']
        },
        {
          id: 'create_tasks',
          name: 'Створення завдань',
          type: 'custom',
          aiPrompt: 'Створи список завдань на основі аналізу: {{document_analysis}}'
        }
      ]
    });

    // Робочий процес затвердження
    this.registerWorkflow({
      id: 'approval_process',
      name: 'Процес затвердження',
      description: 'Багаторівневе затвердження документів',
      trigger: 'command',
      steps: [
        {
          id: 'initial_review',
          name: 'Первинна перевірка',
          type: 'approval',
          requiredRole: ['reviewer', 'analyst'],
          timeoutHours: 24,
          nextSteps: ['manager_approval']
        },
        {
          id: 'manager_approval',
          name: 'Затвердження керівника',
          type: 'approval',
          requiredRole: ['manager', 'commander'],
          timeoutHours: 48,
          condition: 'initial_review_status === "approved"',
          nextSteps: ['final_processing']
        },
        {
          id: 'final_processing',
          name: 'Фінальна обробка',
          type: 'document_analysis',
          nextSteps: []
        }
      ]
    });

    // Робочий процес планового аналізу
    this.registerWorkflow({
      id: 'scheduled_analysis',
      name: 'Плановий аналіз',
      description: 'Щоденний аналіз документів та генерація звітів',
      trigger: 'schedule',
      steps: [
        {
          id: 'collect_documents',
          name: 'Збір документів',
          type: 'data_extraction',
          nextSteps: ['analyze_trends']
        },
        {
          id: 'analyze_trends',
          name: 'Аналіз трендів',
          type: 'custom',
          aiPrompt: 'Проаналізуй тренди в документах за останній тиждень',
          nextSteps: ['generate_report']
        },
        {
          id: 'generate_report',
          name: 'Генерація звіту',
          type: 'document_analysis',
          nextSteps: ['distribute_report']
        },
        {
          id: 'distribute_report',
          name: 'Розповсюдження звіту',
          type: 'notification',
          nextSteps: []
        }
      ]
    });
  }

  /**
   * Реєстрація нового робочого процесу
   */
  registerWorkflow(workflow: WorkflowDefinition): void {
    this.workflows.set(workflow.id, workflow);
    logger.info('Зареєстровано робочий процес', {
      component: 'WorkflowAutomationEngine',
      workflowId: workflow.id,
      name: workflow.name,
      stepsCount: workflow.steps.length
    });
  }

  /**
   * Запуск робочого процесу
   */
  async startWorkflow(
    workflowId: string, 
    variables: Record<string, any> = {},
    triggeredBy?: string
  ): Promise<string> {
    const workflow = this.workflows.get(workflowId);
    if (!workflow) {
      throw new Error(`Робочий процес не знайдено: ${workflowId}`);
    }

    const instanceId = this.generateInstanceId();
    const instance: WorkflowInstance = {
      id: instanceId,
      workflowId,
      status: 'running',
      currentStep: workflow.steps[0]?.id || '',
      variables: { ...workflow.variables, ...variables },
      history: [],
      createdAt: new Date(),
      updatedAt: new Date()
    };

    this.instances.set(instanceId, instance);

    logger.info('Запущено робочий процес', {
      component: 'WorkflowAutomationEngine',
      instanceId,
      workflowId,
      triggeredBy,
      variables
    });

    // Запускаємо виконання
    this.executeNextStep(instanceId);

    return instanceId;
  }

  /**
   * Виконання наступного кроку
   */
  private async executeNextStep(instanceId: string): Promise<void> {
    const instance = this.instances.get(instanceId);
    if (!instance || instance.status !== 'running') {
      return;
    }

    const workflow = this.workflows.get(instance.workflowId);
    if (!workflow) {
      this.failWorkflow(instanceId, 'Робочий процес не знайдено');
      return;
    }

    const currentStep = workflow.steps.find(step => step.id === instance.currentStep);
    if (!currentStep) {
      this.completeWorkflow(instanceId);
      return;
    }

    try {
      // Перевірка умов виконання
      if (currentStep.condition && !await this.evaluateCondition(currentStep.condition, instance)) {
        this.addHistoryEntry(instanceId, currentStep.id, 'skipped', null);
        await this.moveToNextStep(instanceId, currentStep);
        return;
      }

      this.addHistoryEntry(instanceId, currentStep.id, 'started', null);

      // Виконання кроку
      const result = await this.executeStep(currentStep, instance);
      
      this.addHistoryEntry(instanceId, currentStep.id, 'completed', result);

      // Переходимо до наступного кроку
      await this.moveToNextStep(instanceId, currentStep);

    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : String(error);
      this.addHistoryEntry(instanceId, currentStep.id, 'failed', null, errorMessage);
      this.failWorkflow(instanceId, errorMessage);
    }
  }

  /**
   * Виконання конкретного кроку
   */
  private async executeStep(step: WorkflowStep, instance: WorkflowInstance): Promise<any> {
    logger.info('Виконання кроку робочого процесу', {
      component: 'WorkflowAutomationEngine',
      instanceId: instance.id,
      stepId: step.id,
      stepType: step.type
    });

    switch (step.type) {
      case 'document_analysis':
        return await this.executeDocumentAnalysis(step, instance);

      case 'data_extraction':
        return await this.executeDataExtraction(step, instance);

      case 'validation':
        return await this.executeValidation(step, instance);

      case 'notification':
        return await this.executeNotification(step, instance);

      case 'custom':
        return await this.executeCustomStep(step, instance);

      case 'approval':
        return await this.executeApprovalStep(step, instance);

      default:
        throw new Error(`Невідомий тип кроку: ${step.type}`);
    }
  }

  /**
   * Виконання аналізу документа
   */
  private async executeDocumentAnalysis(_step: WorkflowStep, instance: WorkflowInstance): Promise<any> {
    const fileId = instance.variables['fileId'];
    if (!fileId) {
      throw new Error('Не вказано fileId для аналізу документа');
    }

    const analysis = await this.documentService.analyzeDocument(fileId);
    
    // Оновлюємо змінні інстансу
    instance.variables['document_analysis'] = analysis;
    instance.variables['document_type'] = analysis.documentType;
    instance.variables['urgency'] = analysis.urgency;
    
    return analysis;
  }

  /**
   * Виконання кастомного кроку з AI
   */
  private async executeCustomStep(step: WorkflowStep, instance: WorkflowInstance): Promise<any> {
    if (!step.aiPrompt) {
      throw new Error('Не вказано AI промпт для кастомного кроку');
    }

    // Заміна змінних у промпті
    let prompt = step.aiPrompt;
    for (const [key, value] of Object.entries(instance.variables)) {
      prompt = prompt.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), String(value));
    }

    const response = await this.aiService.generateResponse(prompt, {
      temperature: 0.3,
      maxTokens: 1000,
      useCache: true
    });

    return response.content;
  }

  /**
   * Виконання кроку затвердження
   */
  private async executeApprovalStep(_step: WorkflowStep, _instance: WorkflowInstance): Promise<any> {
    // Тут логіка для кроків затвердження
    // Наприклад, створення завдань для відповідальних осіб
    
    return {
      status: 'pending_approval',
      requiredRole: 'administrator',
      timeout: null
    };
  }

  /**
   * Виконання сповіщення
   */
  private async executeNotification(step: WorkflowStep, instance: WorkflowInstance): Promise<any> {
    // Логіка відправки сповіщень (Discord, email, etc.)
    logger.info('Відправлено сповіщення', {
      component: 'WorkflowAutomationEngine',
      instanceId: instance.id,
      stepId: step.id
    });
    
    return { notified: true, timestamp: new Date() };
  }

  /**
   * Виконання витягу даних
   */
  private async executeDataExtraction(_step: WorkflowStep, instance: WorkflowInstance): Promise<any> {
    // Логіка витягу даних з Google Drive
    const documents = await (this.googleService as any).searchFiles?.({
      q: "mimeType contains 'document' and modifiedTime > '2024-01-01'",
      pageSize: 100
    }) || { files: [] };

    instance.variables['collected_documents'] = documents;
    return documents;
  }

  /**
   * Виконання валідації
   */
  private async executeValidation(_step: WorkflowStep, _instance: WorkflowInstance): Promise<any> {
    // Логіка валідації документів або даних
    return { validated: true, timestamp: new Date() };
  }

  /**
   * Оцінка умов виконання
   */
  private async evaluateCondition(condition: string, instance: WorkflowInstance): Promise<boolean> {
    try {
      // Простий інтерпретатор умов
      // Можна розширити для більш складної логіки
      const context = instance.variables;
      
      // Заміна змінних
      let evaluationExpression = condition;
      for (const [key, value] of Object.entries(context)) {
        evaluationExpression = evaluationExpression.replace(
          new RegExp(`\\b${key}\\b`, 'g'),
          JSON.stringify(value)
        );
      }

      // Оцінка простих умов
      if (evaluationExpression.includes('===')) {
        const [left, right] = evaluationExpression.split('===').map(s => s.trim());
        return JSON.parse(left || '{}') === JSON.parse(right || '{}');
      }

      return true;
    } catch (error) {
      logger.warn('Помилка оцінки умови', {
        component: 'WorkflowAutomationEngine',
        condition,
        error
      });
      return false;
    }
  }

  /**
   * Перехід до наступного кроку
   */
  private async moveToNextStep(instanceId: string, currentStep: WorkflowStep): Promise<void> {
    const instance = this.instances.get(instanceId);
    if (!instance) return;

    if (!currentStep.nextSteps || currentStep.nextSteps.length === 0) {
      this.completeWorkflow(instanceId);
      return;
    }

    // Поки що беремо перший наступний крок
    // В майбутньому можна додати логіку вибору кроку на основі умов
    const nextStepId = currentStep.nextSteps[0];
    instance.currentStep = nextStepId || '';
    instance.updatedAt = new Date();

    // Продовжуємо виконання
    await this.executeNextStep(instanceId);
  }

  /**
   * Завершення робочого процесу
   */
  private completeWorkflow(instanceId: string): void {
    const instance = this.instances.get(instanceId);
    if (!instance) return;

    instance.status = 'completed';
    instance.completedAt = new Date();
    instance.updatedAt = new Date();

    logger.info('Робочий процес завершено', {
      component: 'WorkflowAutomationEngine',
      instanceId,
      workflowId: instance.workflowId,
      duration: instance.completedAt.getTime() - instance.createdAt.getTime(),
      stepsExecuted: instance.history.length
    });
  }

  /**
   * Провал робочого процесу
   */
  private failWorkflow(instanceId: string, error: string): void {
    const instance = this.instances.get(instanceId);
    if (!instance) return;

    instance.status = 'failed';
    instance.updatedAt = new Date();

    logger.error('Робочий процес провалено', {
      component: 'WorkflowAutomationEngine',
      instanceId,
      workflowId: instance.workflowId,
      error,
      currentStep: instance.currentStep
    });
  }

  /**
   * Додавання запису в історію
   */
  private addHistoryEntry(
    instanceId: string,
    stepId: string,
    status: 'started' | 'completed' | 'failed' | 'skipped',
    result?: any,
    error?: string
  ): void {
    const instance = this.instances.get(instanceId);
    if (!instance) return;

    instance.history.push({
      stepId,
      status,
      result,
      error: error || undefined,
      timestamp: new Date()
    } as any);

    instance.updatedAt = new Date();
  }

  /**
   * Генерація ID інстансу
   */
  private generateInstanceId(): string {
    return `wf_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  }

  /**
   * Отримання статусу робочого процесу
   */
  getWorkflowStatus(instanceId: string): WorkflowInstance | null {
    return this.instances.get(instanceId) || null;
  }

  /**
   * Отримання списку активних робочих процесів
   */
  getActiveWorkflows(): WorkflowInstance[] {
    return Array.from(this.instances.values()).filter(
      instance => instance.status === 'running'
    );
  }
}