/**
 * 🔧 Конфігурація правил робочих процесів
 * Workflow Rules Configuration
 */

export interface WorkflowRuleConfig {
  id: string;
  name: string;
  description: string;
  domain: 'business' | 'administrative' | 'legal' | 'financial' | 'technical' | 'general';
  priority: number;
  enabled: boolean;
  conditions: {
    documentType?: string[];
    urgency?: string[];
    keywords?: string[];
    fileSize?: { min?: number; max?: number };
    author?: string[];
    customCondition?: string;
  };
  actions: Array<{
    type: 'notify' | 'route' | 'analyze' | 'approve' | 'escalate' | 'archive' | 'transform' | 'validate';
    config: Record<string, any>;
    delay?: number;
    condition?: string;
  }>;
  schedule?: {
    enabled: boolean;
    cron?: string;
    timezone?: string;
  };
}

export const WORKFLOW_RULES_CONFIG: WorkflowRuleConfig[] = [
  
  // 🚨 КРИТИЧНІ ДОКУМЕНТИ
  {
    id: 'critical_business_documents',
    name: 'Критичні бізнес документи',
    description: 'Негайна обробка критичних бізнес контрактів та звітів',
    domain: 'business',
    priority: 100,
    enabled: true,
    conditions: {
      documentType: ['business_contract', 'business_report'],
      urgency: ['critical'],
      keywords: ['терміново', 'негайно', 'критично', 'надзвичайна ситуація']
    },
    actions: [
      {
        type: 'notify',
        config: {
          channels: ['alerts', 'management'],
          message: '🚨 КРИТИЧНИЙ БІЗНЕС ДОКУМЕНТ потребує негайної уваги!',
          priority: 'high',
          mentions: ['@manager', '@duty_manager']
        }
      },
      {
        type: 'analyze',
        config: {
          analysisType: 'full',
          includeRiskAssessment: true,
          language: 'uk'
        }
      },
      {
        type: 'escalate',
        config: {
          level: 'command',
          timeout: 300000, // 5 хвилин
          reason: 'Критична терміновість'
        }
      }
    ]
  },

  // 📋 АДМІНІСТРАТИВНІ ДОКУМЕНТИ
  {
    id: 'administrative_processing',
    name: 'Обробка адміністративних документів',
    description: 'Стандартна обробка адміністративних документів та наказів',
    domain: 'administrative',
    priority: 70,
    enabled: true,
    conditions: {
      documentType: ['administrative_doc', 'order', 'instruction'],
      urgency: ['high', 'medium']
    },
    actions: [
      {
        type: 'analyze',
        config: {
          analysisType: 'compliance',
          language: 'uk',
          checkStandards: ['DSTU', 'legal_requirements']
        }
      },
      {
        type: 'validate',
        config: {
          validationType: 'format_check',
          requiredFields: ['number', 'date', 'signature', 'title']
        }
      },
      {
        type: 'route',
        config: {
          targetChannel: 'administrative_docs',
          targetRole: 'admin_officer',
          includeAnalysis: true
        },
        delay: 5000 // 5 секунд після аналізу
      }
    ]
  },

  // ⚖️ ЮРИДИЧНІ ДОКУМЕНТИ
  {
    id: 'legal_document_review',
    name: 'Правова експертиза документів',
    description: 'Автоматична правова перевірка договорів та юридичних документів',
    domain: 'legal',
    priority: 80,
    enabled: true,
    conditions: {
      documentType: ['legal_contract', 'agreement', 'legal_opinion'],
      keywords: ['договір', 'угода', 'контракт', 'правовий висновок']
    },
    actions: [
      {
        type: 'analyze',
        config: {
          analysisType: 'legal_compliance',
          checkCompliance: true,
          language: 'uk',
          legalStandards: ['civil_code', 'commercial_code', 'labor_code']
        }
      },
      {
        type: 'notify',
        config: {
          targetRole: 'legal_counsel',
          message: '⚖️ Новий документ для правової експертизи',
          includePreview: true
        }
      },
      {
        type: 'approve',
        config: {
          requiredRole: 'legal_counsel',
          timeoutHours: 48,
          escalateAfterTimeout: true
        },
        delay: 60000 // 1 хвилина після аналізу
      }
    ]
  },

  // 💰 ФІНАНСОВІ ДОКУМЕНТИ
  {
    id: 'financial_document_processing',
    name: 'Обробка фінансових документів',
    description: 'Перевірка та обробка фінансових звітів і кошторисів',
    domain: 'financial',
    priority: 75,
    enabled: true,
    conditions: {
      documentType: ['financial_report', 'budget', 'invoice'],
      keywords: ['кошторис', 'бюджет', 'фінансовий звіт', 'рахунок']
    },
    actions: [
      {
        type: 'analyze',
        config: {
          analysisType: 'financial',
          includeRiskAssessment: true,
          checkBudgetCompliance: true,
          language: 'uk'
        }
      },
      {
        type: 'validate',
        config: {
          validationType: 'financial_check',
          checkSums: true,
          verifyCalculations: true,
          checkBudgetLimits: true
        }
      },
      {
        type: 'route',
        config: {
          targetChannel: 'finance',
          targetRole: 'financial_officer',
          requireApproval: true
        }
      }
    ]
  },

  // 📊 ЗВІТИ ТА АНАЛІТИКА
  {
    id: 'report_processing',
    name: 'Обробка звітів та аналітичних документів',
    description: 'Автоматична обробка та розповсюдження звітів',
    domain: 'general',
    priority: 60,
    enabled: true,
    conditions: {
      documentType: ['report', 'analysis', 'statistics'],
      keywords: ['звіт', 'аналіз', 'статистика', 'моніторинг']
    },
    actions: [
      {
        type: 'analyze',
        config: {
          analysisType: 'general',
          extractKeyMetrics: true,
          generateSummary: true,
          language: 'uk'
        }
      },
      {
        type: 'transform',
        config: {
          createVisualizations: true,
          generateExecSummary: true,
          format: 'presentation_ready'
        }
      },
      {
        type: 'route',
        config: {
          targetChannel: 'reports',
          distributionList: ['management', 'analysts'],
          scheduleDistribution: true
        }
      }
    ]
  },

  // 🔍 ВЕЛИКІ ФАЙЛИ
  {
    id: 'large_file_processing',
    name: 'Обробка великих файлів',
    description: 'Спеціальна обробка великих документів та архівів',
    domain: 'technical',
    priority: 50,
    enabled: true,
    conditions: {
      fileSize: { min: 10485760 }, // більше 10MB
      customCondition: 'file.size > 10MB OR file.type === "archive"'
    },
    actions: [
      {
        type: 'notify',
        config: {
          message: '📦 Великий файл завантажено, розпочинається обробка...',
          showProgress: true
        }
      },
      {
        type: 'analyze',
        config: {
          analysisType: 'basic',
          timeoutMinutes: 30,
          lowPriority: true
        }
      },
      {
        type: 'archive',
        config: {
          compression: true,
          retention: '5_years',
          indexForSearch: true
        },
        delay: 300000 // 5 хвилин після аналізу
      }
    ]
  },

  // 📅 ПЛАНОВІ ПЕРЕВІРКИ
  {
    id: 'scheduled_document_review',
    name: 'Планові перевірки документів',
    description: 'Регулярна перевірка та аналіз документів за розкладом',
    domain: 'general',
    priority: 30,
    enabled: true,
    conditions: {
      customCondition: 'scheduled_task === true'
    },
    actions: [
      {
        type: 'analyze',
        config: {
          analysisType: 'compliance_audit',
          generateReport: true,
          includeRecommendations: true
        }
      },
      {
        type: 'notify',
        config: {
          targetChannel: 'audit_reports',
          message: '📋 Завершено планову перевірку документів',
          attachReport: true
        }
      }
    ],
    schedule: {
      enabled: true,
      cron: '0 9 * * 1', // Щопонеділка о 9:00
      timezone: 'Europe/Kiev'
    }
  },

  // 🚫 ЗАБОРОНЕНІ ТИПИ
  {
    id: 'rejected_documents',
    name: 'Відхилення неприпустимих документів',
    description: 'Автоматичне відхилення документів неприпустимих типів',
    domain: 'general',
    priority: 90,
    enabled: true,
    conditions: {
      documentType: ['spam', 'malicious', 'inappropriate'],
      keywords: ['заборонено', 'порушення', 'неприпустимо']
    },
    actions: [
      {
        type: 'notify',
        config: {
          message: '🚫 Документ відхилено через порушення політики',
          alertSecurity: true
        }
      },
      {
        type: 'archive',
        config: {
          quarantine: true,
          retentionDays: 30,
          accessRestricted: true
        }
      }
    ]
  }
];

/**
 * Отримання правил за доменом
 */
export function getRulesByDomain(domain: string): WorkflowRuleConfig[] {
  return WORKFLOW_RULES_CONFIG.filter(rule => 
    rule.domain === domain && rule.enabled
  ).sort((a, b) => b.priority - a.priority);
}

/**
 * Отримання правила за ID
 */
export function getRuleById(id: string): WorkflowRuleConfig | undefined {
  return WORKFLOW_RULES_CONFIG.find(rule => rule.id === id);
}

/**
 * Отримання всіх активних правил
 */
export function getActiveRules(): WorkflowRuleConfig[] {
  return WORKFLOW_RULES_CONFIG.filter(rule => rule.enabled)
    .sort((a, b) => b.priority - a.priority);
}

/**
 * Перевірка чи документ відповідає умовам правила
 */
export function matchesRule(
  rule: WorkflowRuleConfig, 
  documentContext: {
    documentType?: string;
    urgency?: string;
    content?: string;
    fileSize?: number;
    author?: string;
  }
): boolean {
  const { conditions } = rule;
  
  // Перевірка типу документа
  if (conditions.documentType && documentContext.documentType) {
    if (!conditions.documentType.includes(documentContext.documentType)) {
      return false;
    }
  }
  
  // Перевірка терміновості
  if (conditions.urgency && documentContext.urgency) {
    if (!conditions.urgency.includes(documentContext.urgency)) {
      return false;
    }
  }
  
  // Перевірка ключових слів
  if (conditions.keywords && documentContext.content) {
    const hasKeyword = conditions.keywords.some(keyword => 
      documentContext.content!.toLowerCase().includes(keyword.toLowerCase())
    );
    if (!hasKeyword) {
      return false;
    }
  }
  
  // Перевірка розміру файлу
  if (conditions.fileSize && documentContext.fileSize) {
    if (conditions.fileSize.min && documentContext.fileSize < conditions.fileSize.min) {
      return false;
    }
    if (conditions.fileSize.max && documentContext.fileSize > conditions.fileSize.max) {
      return false;
    }
  }
  
  // Перевірка автора
  if (conditions.author && documentContext.author) {
    if (!conditions.author.includes(documentContext.author)) {
      return false;
    }
  }
  
  return true;
}

/**
 * Отримання відповідних правил для документа
 */
export function getMatchingRules(documentContext: {
  documentType?: string;
  urgency?: string;
  content?: string;
  fileSize?: number;
  author?: string;
}): WorkflowRuleConfig[] {
  return getActiveRules().filter(rule => matchesRule(rule, documentContext));
}