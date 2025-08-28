#!/usr/bin/env ts-node

/**
 * 🧪 Document Analysis Service Test
 * Tests the integration of DocumentAnalysisService and related commands
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
  error?: string;
  details?: any;
}

class DocumentAnalysisTest {
  private serviceManager!: ServiceManager;
  private config: any;
  private results: TestResult[] = [];

  constructor() {
    this.config = Config.load();
  }

  /**
   * 🚀 Run all document analysis tests
   */
  async runTests(): Promise<void> {
    console.log('🧪 Starting Document Analysis Tests');
    console.log('=' .repeat(60));

    try {
      // Initialize services
      await this.initializeServices();

      // Run individual tests
      await this.testServiceRegistration();
      await this.testDocumentAnalysisService();
      await this.testAnalysisTypes();

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
   * 📝 Test document analysis service registration
   */
  private async testServiceRegistration(): Promise<void> {
    const testName = 'Document Analysis Service Registration';
    const startTime = Date.now();

    try {
      const requiredServices = [
        'documentAnalysis',
        'documentAnalyzer',
        'google'
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

      const testResult: TestResult = {
        testName,
        success: missingServices.length === 0,
        duration: Date.now() - startTime,
        details: { 
          availableServices, 
          missingServices,
          totalServices: this.serviceManager.getServiceNames().length
        }
      };
      
      if (missingServices.length > 0) {
        testResult.error = `Missing services: ${missingServices.join(', ')}`;
      }
      
      this.addTestResult(testResult);

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
   * 📄 Test DocumentAnalysisService functionality
   */
  private async testDocumentAnalysisService(): Promise<void> {
    const testName = 'DocumentAnalysisService Functionality';
    const startTime = Date.now();

    try {
      const documentAnalysisService = this.serviceManager.getService('documentAnalysis');
      if (!documentAnalysisService) {
        throw new Error('DocumentAnalysisService not available');
      }

      // Test service methods
      const serviceMethods = [
        'analyzeDocument',
        'analyzeDocumentStructure',
        'summarizeDocumentContent',
        'extractActionItems',
        'checkDocumentCompliance',
        'assessDocumentQuality'
      ];

      const missingMethods: string[] = [];
      const availableMethods: string[] = [];
      
      for (const methodName of serviceMethods) {
        if (typeof (documentAnalysisService as any)[methodName] === 'function') {
          availableMethods.push(methodName);
        } else {
          missingMethods.push(methodName);
        }
      }

      const testDetails = {
        availableMethods,
        missingMethods,
        totalMethods: serviceMethods.length
      };

      const success = missingMethods.length === 0;
      
      this.addTestResult({
        testName,
        success,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`${success ? '✅' : '⚠️'} ${testName}: ${availableMethods.length}/${serviceMethods.length} methods available`);

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
   * 📊 Test different analysis types
   */
  private async testAnalysisTypes(): Promise<void> {
    const testName = 'Document Analysis Types';
    const startTime = Date.now();

    try {
      const documentAnalysisService = this.serviceManager.getService('documentAnalysis');
      if (!documentAnalysisService) {
        throw new Error('DocumentAnalysisService not available');
      }

      // Test different analysis types that should be available
      const analysisTypes = [
        'structure',
        'summary',
        'actions',
        'compliance',
        'quality'
      ];

      const supportedTypes: string[] = [];
      const unsupportedTypes: string[] = [];
      
      // We can't actually run the analysis without real documents,
      // but we can check if the service has the methods for these types
      const methodMap: Record<string, string> = {
        'structure': 'analyzeDocumentStructure',
        'summary': 'summarizeDocumentContent',
        'actions': 'extractActionItems',
        'compliance': 'checkDocumentCompliance',
        'quality': 'assessDocumentQuality'
      };

      for (const type of analysisTypes) {
        const methodName = methodMap[type];
        if (methodName && typeof (documentAnalysisService as any)[methodName] === 'function') {
          supportedTypes.push(type);
        } else {
          unsupportedTypes.push(type);
        }
      }

      const testDetails = {
        supportedTypes,
        unsupportedTypes,
        totalTypes: analysisTypes.length
      };

      const success = unsupportedTypes.length === 0;
      
      this.addTestResult({
        testName,
        success,
        duration: Date.now() - startTime,
        details: testDetails
      });

      console.log(`${success ? '✅' : '⚠️'} ${testName}: ${supportedTypes.length}/${analysisTypes.length} analysis types supported`);

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
   * 📊 Add test result to collection
   */
  private addTestResult(result: TestResult): void {
    this.results.push(result);
  }

  /**
   * 📋 Generate final test report
   */
  private generateReport(): void {
    console.log('\n' + '=' .repeat(60));
    console.log('📊 TEST RESULTS SUMMARY');
    console.log('=' .repeat(60));

    let passedTests = 0;
    let totalTests = this.results.length;

    for (const result of this.results) {
      const status = result.success ? '✅ PASS' : '❌ FAIL';
      console.log(`${status} ${result.testName} (${result.duration}ms)`);
      
      if (result.details) {
        console.log(`   Details: ${JSON.stringify(result.details)}`);
      }
      
      if (result.error) {
        console.log(`   Error: ${result.error}`);
      }
      
      if (result.success) passedTests++;
    }

    console.log('-' .repeat(60));
    console.log(`📈 Overall: ${passedTests}/${totalTests} tests passed`);
    console.log(`🎯 Success Rate: ${((passedTests / totalTests) * 100).toFixed(1)}%`);
    
    if (passedTests === totalTests) {
      console.log('🎉 All tests passed!');
      process.exit(0);
    } else {
      console.log('⚠️  Some tests failed.');
      process.exit(1);
    }
  }
}

// Run the tests if this file is executed directly
if (require.main === module) {
  const testSuite = new DocumentAnalysisTest();
  testSuite.runTests().catch(console.error);
}

export default DocumentAnalysisTest;