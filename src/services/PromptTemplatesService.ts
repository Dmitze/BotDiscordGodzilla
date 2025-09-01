import type { BotConfig } from '@/types';

/**
 * Prompt template with versioning and localization support
 */
export interface PromptTemplate {
  id: string;
  version: number;
  template: string;
  variables: string[];
  locale?: string;
  description?: string;
}

/**
 * Service for managing prompt templates with versioning and localization
 */
export class PromptTemplatesService {
  private templates: Map<string, PromptTemplate[]>;
  
  constructor(_config: BotConfig) {
    this.templates = new Map();
    this.initializeDefaultTemplates();
  }

  /**
   * Initialize default prompt templates
   */
  private initializeDefaultTemplates(): void {
    // Document QA template with citations
    this.registerTemplate({
      id: 'document_qa',
      version: 1,
      template: `Ти — помічник, який відповідає стисло, українською, з посиланнями на джерела.

Питання:
{question}

Контекст (релевантні уривки):
{context}

Відповідь: наведи коротку відповідь та в кінці перелік джерел у форматі [{source_index}] з короткими назвами.`,
      variables: ['question', 'context'],
      locale: 'uk',
      description: 'Template for document-based QA with citations'
    });

    // Document summary template
    this.registerTemplate({
      id: 'document_summary',
      version: 1,
      template: `Створи короткий зміст українською мовою для наступного документа:

Назва документа: {document_name}

Текст документа:
{document_text}

Зміст повинен включати:
1. Основну тему документа
2. Ключові ідеї та факти
3. Висновки (якщо є)

Зміст:`,
      variables: ['document_name', 'document_text'],
      locale: 'uk',
      description: 'Template for document summarization'
    });

    // Key points extraction template
    this.registerTemplate({
      id: 'key_points',
      version: 1,
      template: `Видобудь ключові моменти з наступного тексту. Українською мовою:

Текст:
{document_text}

Ключові моменти:
-`,
      variables: ['document_text'],
      locale: 'uk',
      description: 'Template for extracting key points from text'
    });

    // Fact extraction template
    this.registerTemplate({
      id: 'fact_extraction',
      version: 1,
      template: `Видобудь конкретні факти з наступного тексту. Українською мовою:

Текст:
{document_text}

Факти:
1.`,
      variables: ['document_text'],
      locale: 'uk',
      description: 'Template for extracting facts from text'
    });
  }

  /**
   * Register a new prompt template
   * @param template The prompt template to register
   */
  public registerTemplate(template: PromptTemplate): void {
    const key = this.getTemplateKey(template.id, template.locale ?? 'uk');
    if (!this.templates.has(key)) {
      this.templates.set(key, []);
    }
    
    const versions = this.templates.get(key)!;
    versions.push(template);
    // Sort by version number (newest first)
    versions.sort((a, b) => b.version - a.version);
  }

  /**
   * Get a prompt template by ID and locale
   * @param id Template ID
   * @param locale Locale (optional, defaults to 'uk')
   * @param version Specific version (optional, defaults to latest)
   * @returns The prompt template or undefined if not found
   */
  public getTemplate(id: string, locale: string = 'uk', version?: number): PromptTemplate | undefined {
    const key = this.getTemplateKey(id, locale);
    const versions = this.templates.get(key);
    
    if (!versions || versions.length === 0) {
      return undefined;
    }
    
    if (version !== undefined) {
      return versions.find(t => t.version === version);
    }
    
    // Return the latest version
    return versions[0];
  }

  /**
   * Render a prompt template with variables
   * @param id Template ID
   * @param variables Template variables
   * @param locale Locale (optional, defaults to 'uk')
   * @param version Specific version (optional, defaults to latest)
   * @returns Rendered prompt or undefined if template not found
   */
  public renderPrompt(
    id: string, 
    variables: Record<string, string | number>, 
    locale: string = 'uk', 
    version?: number
  ): string | undefined {
    const template = this.getTemplate(id, locale, version);
    
    if (!template) {
      return undefined;
    }
    
    let prompt = template.template;
    
    // Replace variables in the template
    for (const [key, value] of Object.entries(variables)) {
      const placeholder = `{${key}}`;
      prompt = prompt.replace(new RegExp(placeholder, 'g'), String(value));
    }
    
    return prompt;
  }

  /**
   * Get all available template IDs
   * @returns Array of template IDs
   */
  public getTemplateIds(): string[] {
    const ids = new Set<string>();
    for (const key of this.templates.keys()) {
      const id = key.split(':')[0];
      ids.add(id);
    }
    return Array.from(ids);
  }

  /**
   * Get available locales for a template
   * @param id Template ID
   * @returns Array of available locales
   */
  public getAvailableLocales(id: string): string[] {
    const locales = new Set<string>();
    for (const key of this.templates.keys()) {
      const [templateId, locale] = key.split(':');
      if (templateId === id && locale) {
        locales.add(locale);
      }
    }
    return Array.from(locales);
  }

  /**
   * Generate template key
   * @param id Template ID
   * @param locale Locale
   * @returns Template key
   */
  private getTemplateKey(id: string, locale?: string): string {
    return locale ? `${id}:${locale}` : `${id}:uk`;
  }
}