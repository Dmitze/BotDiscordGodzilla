import { PromptTemplatesService } from '../PromptTemplatesService';
import type { BotConfig } from '@/types';

// Mock config
const mockConfig = {} as BotConfig;

describe('PromptTemplatesService', () => {
  let service: PromptTemplatesService;

  beforeEach(() => {
    service = new PromptTemplatesService(mockConfig);
  });

  describe('initializeDefaultTemplates', () => {
    it('should initialize default templates', () => {
      const templateIds = service.getTemplateIds();
      expect(templateIds).toContain('document_qa');
      expect(templateIds).toContain('document_summary');
      expect(templateIds).toContain('key_points');
      expect(templateIds).toContain('fact_extraction');
    });
  });

  describe('registerTemplate', () => {
    it('should register a new template', () => {
      const template = {
        id: 'test_template',
        version: 1,
        template: 'This is a test template with {variable}',
        variables: ['variable'],
        locale: 'uk'
      };

      service.registerTemplate(template);
      const retrieved = service.getTemplate('test_template', 'uk');

      expect(retrieved).toEqual(template);
    });
  });

  describe('getTemplate', () => {
    it('should return the latest version when no version specified', () => {
      const templateV1 = {
        id: 'versioned_template',
        version: 1,
        template: 'Version 1',
        variables: [],
        locale: 'uk'
      };

      const templateV2 = {
        id: 'versioned_template',
        version: 2,
        template: 'Version 2',
        variables: [],
        locale: 'uk'
      };

      service.registerTemplate(templateV1);
      service.registerTemplate(templateV2);

      const retrieved = service.getTemplate('versioned_template', 'uk');
      expect(retrieved?.version).toBe(2);
    });

    it('should return specific version when requested', () => {
      const templateV1 = {
        id: 'specific_version_template',
        version: 1,
        template: 'Version 1',
        variables: [],
        locale: 'uk'
      };

      const templateV2 = {
        id: 'specific_version_template',
        version: 2,
        template: 'Version 2',
        variables: [],
        locale: 'uk'
      };

      service.registerTemplate(templateV1);
      service.registerTemplate(templateV2);

      const retrieved = service.getTemplate('specific_version_template', 'uk', 1);
      expect(retrieved?.version).toBe(1);
      expect(retrieved?.template).toBe('Version 1');
    });
  });

  describe('renderPrompt', () => {
    it('should render prompt with variables', () => {
      const template = {
        id: 'render_test',
        version: 1,
        template: 'Hello {name}, you have {count} messages.',
        variables: ['name', 'count'],
        locale: 'uk'
      };

      service.registerTemplate(template);

      const rendered = service.renderPrompt('render_test', {
        name: 'John',
        count: 5
      }, 'uk');

      expect(rendered).toBe('Hello John, you have 5 messages.');
    });

    it('should return undefined for non-existent template', () => {
      const rendered = service.renderPrompt('non_existent', {}, 'uk');
      expect(rendered).toBeUndefined();
    });
  });

  describe('getAvailableLocales', () => {
    it('should return available locales for a template', () => {
      const templateUk = {
        id: 'localized_template',
        version: 1,
        template: 'Ukrainian version',
        variables: [],
        locale: 'uk'
      };

      const templateEn = {
        id: 'localized_template',
        version: 1,
        template: 'English version',
        variables: [],
        locale: 'en'
      };

      service.registerTemplate(templateUk);
      service.registerTemplate(templateEn);

      const locales = service.getAvailableLocales('localized_template');
      expect(locales).toContain('uk');
      expect(locales).toContain('en');
      expect(locales).toHaveLength(2);
    });
  });
});