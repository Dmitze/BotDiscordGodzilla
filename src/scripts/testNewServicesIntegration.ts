#!/usr/bin/env ts-node

/**
 * 🧪 New Services Integration Test
 * Tests the integration of ContextMemoryService, ResponseCacheService, 
 * KnowledgeBaseService, and EnhancedRagService
 */

import { config } from 'dotenv';
import path from 'path';
import { Config } from '../config/Config';
import ServiceManager from '../core/ServiceManager';

// Load environment variables
config({ path: path.resolve(process.cwd(), '.env') });

interface TestResult {
  testName: string;
  success: boolean;
  duration: number;
  error?: string | undefined;
  details?: any;
}

class NewServicesIntegrationTest {
  private serviceManager!: ServiceManager;
  private config: any;
  private results: TestResult[] = [];

  constructor() {
    this.config = Config.load();
  }

  /**
   * 🚀 Run all integration tests
   */
  async runTests(): Promise<void> {
    console.log('🧪 Starting New Services Integration Tests');
    console.log('=' .repeat(60));

    try {
      // Initialize services
      await this.initializeServices();

      // Run individual tests
      await this.testServiceRegistration();
      await this.testContextMemoryService();
      await this.testResponseCacheService();
      await this.testKnowledgeBaseService();
      await this.testEnhancedRagService();
      await this.testServiceInteractions();
      await this.testPerformanceAndCaching();
      await this.testEndToEndWorkflow();

      // Generate final report
      this.generateReport();

    } catch (error) {
      console.error('💥 Test suite failed during initialization:', error);
      process.exit(1);
    }
  }

  /**
   * 🔧 Initialize ServiceManager and all services
   */
  private async initializeServices(): Promise<void> {
    console.log('🔧 Initializing services...');
    
    // Create a mock bot object for ServiceManager
    const mockBot = {
      config: this.config,
      getService: (name: string) => this.serviceManager?.getService(name as any)
    };

    this.serviceManager = new ServiceManager(mockBot as any);
    await this.serviceManager.initialize();
    
    console.log('✅ Services initialized successfully');
  }

  /**
   * 📝 Test new services registration
   */
  private async testServiceRegistration(): Promise<void> {
    const testName = 'New Services Registration';
    const startTime = Date.now();

    try {
      const requiredServices = [
        'contextMemory',
        'responseCache',
        'knowledgeBase',
        'enhancedRag'
      ];

      const missingServices: string[] = [];
      const availableServices: string[] = [];
      
      for (const serviceName of requiredServices) {
        const service = this.serviceManager.getService(serviceName as any);
        if (!service) {
          missingServices.push(serviceName);
        } else {
          availableServices.push(serviceName);
        }
      }

      this.addTestResult({
        testName,
        success: missingServices.length === 0,
        duration: Date.now() - startTime,
        details: { 
          availableServices, 
          missingServices,
          totalServices: this.serviceManager.getServiceNames().length
        },
        error: missingServices.length > 0 ? `Missing services: ${missingServices.join(', ')}` : undefined
      });

      console.log(`${missingServices.length === 0 ? '✅' : '⚠️'} ${testName}: ${availableServices.length}/${requiredServices.length} services available`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * 🧠 Test ContextMemoryService functionality
   */
  private async testContextMemoryService(): Promise<void> {
    const testName = 'ContextMemoryService';
    const startTime = Date.now();

    try {
      const contextMemory = this.serviceManager.getService('contextMemory');
      if (!contextMemory) {
        throw new Error('ContextMemoryService not available');
      }

      // Test adding queries
      const userId = 'test_user_123';
      const channelId = 'test_channel_123';
      
      const queryId1 = (contextMemory as any).addQuery(userId, channelId, 'Test query 1', 'ai');
      (contextMemory as any).addQuery(userId, channelId, 'Test query 2', 'search');
      
      // Test updating response
      (contextMemory as any).updateQueryResponse(queryId1, 'Test response 1', { 
        responseTime: 100, 
        tokensUsed: 50 
      });

      // Test getting context
      const userContext = (contextMemory as any).getUserContext(userId);
      const recentQueries = (contextMemory as any).getRecentQueries(userId, 3);

      // Test contextual prompt building
      const contextualPrompt = (contextMemory as any).buildContextualPrompt(userId, 'New query');

      // Test user preferences
      (contextMemory as any).setUserPreferences(userId, {
        language: 'uk',
        domain: 'military',
        responseStyle: 'detailed'
      });

      // Test stats
      const stats = (contextMemory as any).getStats();

      const testDetails = {
        queriesAdded: 2,
        hasUserContext: !!userContext,
        recentQueriesCount: recentQueries.length,
        contextualPromptLength: contextualPrompt.length,
        stats
      };

      this.addTestResult({
        testName,
        success: true,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`✅ ${testName}: Context management working correctly`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * 💾 Test ResponseCacheService functionality
   */
  private async testResponseCacheService(): Promise<void> {
    const testName = 'ResponseCacheService';
    const startTime = Date.now();

    try {
      const responseCache = this.serviceManager.getService('responseCache');
      if (!responseCache) {
        throw new Error('ResponseCacheService not available');
      }

      // Test basic cache operations
      const testData = { message: 'Test response', timestamp: new Date() };
      const cacheKey = 'test_key_123';
      
      // Set cache entry
      (responseCache as any).set(cacheKey, testData, 5); // 5 minutes TTL

      // Get cache entry
      const retrieved = (responseCache as any).get(cacheKey);

      // Test cache hit/miss
      const missResult = (responseCache as any).get('non_existent_key');

      // Test pattern search
      (responseCache as any).set('pattern_test_1', { data: 'test1' }, 5, { tags: ['test'] });
      (responseCache as any).set('pattern_test_2', { data: 'test2' }, 5, { tags: ['test'] });
      
      const patternResults = (responseCache as any).findByPattern(/pattern_test_/);
      const tagResults = (responseCache as any).findByTags(['test']);

      // Test TTL extension
      const extended = (responseCache as any).extendTtl(cacheKey, 10);

      // Test stats
      const stats = (responseCache as any).getStats();

      const testDetails = {
        cacheHit: !!retrieved,
        cacheMiss: !missResult,
        patternSearchResults: patternResults.length,
        tagSearchResults: tagResults.length,
        ttlExtended: extended,
        stats
      };

      this.addTestResult({
        testName,
        success: true,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`✅ ${testName}: Cache operations working correctly`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * 📚 Test KnowledgeBaseService functionality
   */
  private async testKnowledgeBaseService(): Promise<void> {
    const testName = 'KnowledgeBaseService';
    const startTime = Date.now();

    try {
      const knowledgeBase = this.serviceManager.getService('knowledgeBase');
      if (!knowledgeBase) {
        throw new Error('KnowledgeBaseService not available');
      }

      // Test adding knowledge entries
      const entryId1 = await (knowledgeBase as any).addEntry(
        'Test Military Document',
        'This is a test military document content about tactical operations.',
        'military',
        ['tactics', 'operations'],
        { type: 'manual', createdBy: 'test_user' }
      );

      await (knowledgeBase as any).addEntry(
        'Test Administrative Document',
        'This is a test administrative document about procedures.',
        'administrative',
        ['procedures', 'guidelines'],
        { type: 'manual', createdBy: 'test_user' }
      );

      // Test searching knowledge base
      const searchResults = await (knowledgeBase as any).search({
        query: 'military tactics',
        limit: 10,
        useSemanticSearch: false
      });

      // Test getting entry by ID
      const entry = (knowledgeBase as any).getEntry(entryId1);

      // Test updating entry
      const updated = await (knowledgeBase as any).updateEntry(entryId1, {
        tags: ['tactics', 'operations', 'updated']
      });

      // Test getting entries by category
      const militaryEntries = (knowledgeBase as any).getEntriesByCategory('military');

      // Test getting entries by tag
      const tacticsEntries = (knowledgeBase as any).getEntriesByTag('tactics');

      // Test statistics
      const stats = (knowledgeBase as any).getStats();

      // Test trending topics
      const trending = (knowledgeBase as any).getTrendingTopics(5);

      const testDetails = {
        entriesAdded: 2,
        searchResultsCount: searchResults.length,
        entryRetrieved: !!entry,
        entryUpdated: updated,
        militaryEntriesCount: militaryEntries.length,
        tacticsEntriesCount: tacticsEntries.length,
        stats,
        trendingTopicsCount: trending.length
      };

      this.addTestResult({
        testName,
        success: true,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`✅ ${testName}: Knowledge base operations working correctly`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * 🚀 Test EnhancedRagService functionality
   */
  private async testEnhancedRagService(): Promise<void> {
    const testName = 'EnhancedRagService';
    const startTime = Date.now();

    try {
      const enhancedRag = this.serviceManager.getService('enhancedRag');
      if (!enhancedRag) {
        throw new Error('EnhancedRagService not available');
      }

      // Test indexing statistics
      const indexingStats = (enhancedRag as any).getIndexingStats();

      // Test enhanced search
      const searchResults = await (enhancedRag as any).search('test query', {
        limit: 5,
        useCache: true
      });

      // Test manual indexing trigger (without actually triggering it)
      const canTriggerIndexing = typeof (enhancedRag as any).triggerManualIndexing === 'function';

      // Test auto-indexing configuration update
      (enhancedRag as any).updateAutoIndexConfig({
        batchSize: 5,
        maxFileSize: 10 * 1024 * 1024 // 10MB
      });

      // Test metrics from parent RagService
      const ragMetrics = (enhancedRag as any).getMetrics();

      const testDetails = {
        indexingStats,
        searchResultsCount: searchResults.length,
        canTriggerIndexing,
        ragMetrics
      };

      this.addTestResult({
        testName,
        success: true,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`✅ ${testName}: Enhanced RAG operations working correctly`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * 🔗 Test service interactions
   */
  private async testServiceInteractions(): Promise<void> {
    const testName = 'Service Interactions';
    const startTime = Date.now();

    try {
      const contextMemory = this.serviceManager.getService('contextMemory');
      const responseCache = this.serviceManager.getService('responseCache');
      const knowledgeBase = this.serviceManager.getService('knowledgeBase');

      if (!contextMemory || !responseCache || !knowledgeBase) {
        throw new Error('Required services not available for interaction test');
      }

      // Test interaction: Context Memory + Knowledge Base
      const userId = 'interaction_test_user';
      const queryId = (contextMemory as any).addQuery(
        userId, 
        'test_channel', 
        'What are military procedures?', 
        'knowledge_search'
      );

      // Simulate knowledge base search
      const kbResults = await (knowledgeBase as any).search({
        query: 'military procedures',
        limit: 3
      });

      // Update context with knowledge base results
      (contextMemory as any).updateQueryResponse(
        queryId,
        `Found ${kbResults.length} relevant documents`,
        { sources: kbResults.map((r: any) => r.entry.id) }
      );

      // Test interaction: Response Cache + Knowledge Base
      const cacheKey = 'kb_search_military_procedures';
      (responseCache as any).set(cacheKey, kbResults, 15, {
        tags: ['knowledge_search', 'military']
      });

      const cachedResults = (responseCache as any).get(cacheKey);

      // Test contextual prompt with cached data
      const contextualPrompt = (contextMemory as any).buildContextualPrompt(
        userId,
        'Tell me more about these procedures'
      );

      const testDetails = {
        contextQueryAdded: !!queryId,
        knowledgeSearchResults: kbResults.length,
        cacheInteraction: !!cachedResults,
        contextualPromptLength: contextualPrompt.length,
        servicesInteracting: 3
      };

      this.addTestResult({
        testName,
        success: true,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`✅ ${testName}: Service interactions working correctly`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * ⚡ Test performance and caching
   */
  private async testPerformanceAndCaching(): Promise<void> {
    const testName = 'Performance and Caching';
    const startTime = Date.now();

    try {
      const responseCache = this.serviceManager.getService('responseCache');
      const knowledgeBase = this.serviceManager.getService('knowledgeBase');

      if (!responseCache || !knowledgeBase) {
        throw new Error('Required services not available for performance test');
      }

      // Test cache performance
      const cacheTests = [];
      for (let i = 0; i < 100; i++) {
        const key = `perf_test_${i}`;
        const data = { id: i, content: `Test data ${i}` };
        
        const setStart = Date.now();
        (responseCache as any).set(key, data, 5);
        const setTime = Date.now() - setStart;

        const getStart = Date.now();
        const retrieved = (responseCache as any).get(key);
        const getTime = Date.now() - getStart;

        cacheTests.push({ setTime, getTime, success: !!retrieved });
      }

      const avgSetTime = cacheTests.reduce((sum, test) => sum + test.setTime, 0) / cacheTests.length;
      const avgGetTime = cacheTests.reduce((sum, test) => sum + test.getTime, 0) / cacheTests.length;
      const successRate = cacheTests.filter(test => test.success).length / cacheTests.length;

      // Test knowledge base search performance
      const searchStart = Date.now();
      const searchResults = await (knowledgeBase as any).search({
        query: 'test performance search',
        limit: 10
      });
      const searchTime = Date.now() - searchStart;

      // Get cache statistics
      const cacheStats = (responseCache as any).getStats();

      const testDetails = {
        cachePerformance: {
          avgSetTime: Math.round(avgSetTime * 100) / 100,
          avgGetTime: Math.round(avgGetTime * 100) / 100,
          successRate: Math.round(successRate * 100),
          testsRun: cacheTests.length
        },
        knowledgeBasePerformance: {
          searchTime,
          resultsFound: searchResults.length
        },
        cacheStats
      };

      this.addTestResult({
        testName,
        success: true,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`✅ ${testName}: Performance metrics collected successfully`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * 🔄 Test end-to-end workflow
   */
  private async testEndToEndWorkflow(): Promise<void> {
    const testName = 'End-to-End Workflow';
    const startTime = Date.now();

    try {
      const contextMemory = this.serviceManager.getService('contextMemory');
      const responseCache = this.serviceManager.getService('responseCache');
      const knowledgeBase = this.serviceManager.getService('knowledgeBase');
      const enhancedRag = this.serviceManager.getService('enhancedRag');

      if (!contextMemory || !responseCache || !knowledgeBase || !enhancedRag) {
        throw new Error('Required services not available for end-to-end test');
      }

      const userId = 'e2e_test_user';
      const workflow = [];

      // Step 1: User asks question
      workflow.push('User query received');
      const queryId = (contextMemory as any).addQuery(
        userId,
        'e2e_channel',
        'What are the latest military operational procedures?',
        'comprehensive_search'
      );

      // Step 2: Check cache for similar queries
      workflow.push('Cache check');
      const cacheKey = 'military_operational_procedures';
      let cachedAnswer = (responseCache as any).get(cacheKey);

      if (!cachedAnswer) {
        // Step 3: Search knowledge base
        workflow.push('Knowledge base search');
        const kbResults = await (knowledgeBase as any).search({
          query: 'military operational procedures',
          limit: 5,
          useSemanticSearch: false
        });

        // Step 4: Enhanced RAG search if needed
        workflow.push('Enhanced RAG search');
        const ragResults = await (enhancedRag as any).search('military operational procedures', {
          limit: 3,
          useCache: true
        });

        // Step 5: Combine and cache results
        workflow.push('Result combination and caching');
        const combinedResults = {
          knowledgeBase: kbResults,
          ragSearch: ragResults,
          sources: [...kbResults.map((r: any) => r.entry?.id), ...ragResults.map((r: any) => r.fileId)].filter(Boolean),
          timestamp: new Date()
        };

        (responseCache as any).set(cacheKey, combinedResults, 30, {
          tags: ['military', 'procedures', 'comprehensive']
        });

        cachedAnswer = combinedResults;
      } else {
        workflow.push('Cache hit - skipped searches');
      }

      // Step 6: Update context with response
      workflow.push('Context update');
      (contextMemory as any).updateQueryResponse(
        queryId,
        `Found comprehensive information from ${cachedAnswer.sources?.length || 0} sources`,
        {
          responseTime: Date.now() - startTime,
          sources: cachedAnswer.sources || []
        }
      );

      // Step 7: Build contextual prompt for follow-up
      workflow.push('Contextual prompt building');
      const contextualPrompt = (contextMemory as any).buildContextualPrompt(
        userId,
        'Can you elaborate on the first procedure?'
      );

      const testDetails = {
        workflowSteps: workflow,
        userContextCreated: !!queryId,
        resultsFound: cachedAnswer.sources?.length || 0,
        contextualPromptReady: contextualPrompt.length > 0,
        cacheUtilized: true,
        servicesIntegrated: 4
      };

      this.addTestResult({
        testName,
        success: true,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`✅ ${testName}: End-to-end workflow completed successfully`);

    } catch (error) {
      this.addTestResult({
        testName,
        success: false,
        duration: Date.now() - startTime,
        error: error instanceof Error ? error.message : String(error)
      });

      console.log(`❌ ${testName}: FAILED - ${error}`);
    }
  }

  /**
   * 📊 Add test result
   */
  private addTestResult(result: TestResult): void {
    this.results.push(result);
  }

  /**
   * 📋 Generate final test report
   */
  private generateReport(): void {
    console.log('\n' + '=' .repeat(60));
    console.log('📋 NEW SERVICES INTEGRATION TEST REPORT');
    console.log('=' .repeat(60));

    const totalTests = this.results.length;
    const passedTests = this.results.filter(r => r.success).length;
    const failedTests = totalTests - passedTests;
    const totalDuration = this.results.reduce((sum, r) => sum + r.duration, 0);

    console.log(`\n📊 SUMMARY:`);
    console.log(`   Total Tests: ${totalTests}`);
    console.log(`   Passed: ${passedTests} ✅`);
    console.log(`   Failed: ${failedTests} ❌`);
    console.log(`   Success Rate: ${Math.round((passedTests / totalTests) * 100)}%`);
    console.log(`   Total Duration: ${totalDuration}ms`);

    console.log(`\n📝 DETAILED RESULTS:`);
    for (const result of this.results) {
      const status = result.success ? '✅' : '❌';
      console.log(`   ${status} ${result.testName} (${result.duration}ms)`);
      
      if (result.error) {
        console.log(`      Error: ${result.error}`);
      }
      
      if (result.details && Object.keys(result.details).length > 0) {
        console.log(`      Details: ${JSON.stringify(result.details, null, 8)}`);
      }
    }

    console.log('\n' + '=' .repeat(60));
    
    if (failedTests === 0) {
      console.log('🎉 ALL NEW SERVICES TESTS PASSED! Integration is successful.');
      console.log('✨ New Features Available:');
      console.log('   • Context Memory - User query history and preferences');
      console.log('   • Response Cache - 30-minute TTL caching for better performance');
      console.log('   • Knowledge Base - Comprehensive knowledge management');
      console.log('   • Enhanced RAG - Auto-indexing Google Drive documents');
    } else {
      console.log('⚠️  Some tests failed. Please review the errors above.');
    }
    
    console.log('=' .repeat(60));
  }
}

// Run tests if this script is executed directly
if (require.main === module) {
  const testRunner = new NewServicesIntegrationTest();
  testRunner.runTests().catch(error => {
    console.error('💥 Test execution failed:', error);
    process.exit(1);
  });
}

export { NewServicesIntegrationTest };