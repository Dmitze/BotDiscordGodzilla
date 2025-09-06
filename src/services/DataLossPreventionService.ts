import { BaseService } from '@/core/BaseService';
import type { BotConfig, HealthStatus, ServiceStats } from '@/types';
import type { DriveFile } from '@/types/drive';
import logger from '@/utils/logger';

export interface SensitiveDataPattern {
  id: string;
  name: string;
  pattern: RegExp;
  severity: 'low' | 'medium' | 'high' | 'critical';
  category: string;
  description: string;
}

export interface SensitiveDataFinding {
  id: string;
  patternId: string;
  patternName: string;
  severity: 'low' | 'medium' | 'high' | 'critical';
  category: string;
  matchedText: string;
  position: { start: number; end: number };
  context: string;
  confidence: number; // 0-100
}

export interface DlpScanResult {
  fileId: string;
  fileName: string;
  scannedAt: Date;
  totalFindings: number;
  findingsBySeverity: Record<string, number>;
  findingsByCategory: Record<string, number>;
  findings: SensitiveDataFinding[];
  riskScore: number; // 0-100
  recommendedActions: string[];
}

export interface DlpPolicy {
  id: string;
  name: string;
  description: string;
  enabled: boolean;
  patterns: string[]; // Pattern IDs
  severityThreshold: 'low' | 'medium' | 'high' | 'critical';
  actions: ('log' | 'alert' | 'block' | 'quarantine')[];
}

export class DataLossPreventionService extends BaseService {
  private patterns: SensitiveDataPattern[] = [];
  private policies: DlpPolicy[] = [];
  private scanResults: Map<string, DlpScanResult> = new Map();
  private readonly MAX_CACHE_RESULTS = 1000;
  
  constructor(config: BotConfig) {
    super('DataLossPreventionService', config);
    this.initializeDefaultPatterns();
    this.initializeDefaultPolicies();
  }

  /**
   * Initialize service
   */
  protected async onInitialize(): Promise<void> {
    // Implementation for initialization if needed
    logger.info('DataLossPreventionService initialized', {
      component: 'DataLossPreventionService'
    });
  }

  /**
   * Shutdown service
   */
  protected async onShutdown(): Promise<void> {
    // Implementation for shutdown if needed
    logger.info('DataLossPreventionService shutdown', {
      component: 'DataLossPreventionService'
    });
  }

  /**
   * Health check
   */
  protected async onHealthCheck(): Promise<HealthStatus> {
    return {
      healthy: true,
      service: 'DataLossPreventionService'
    };
  }

  /**
   * Get service stats
   */
  protected onGetStats(): Partial<ServiceStats> {
    return {
      totalPatterns: this.patterns.length,
      activePolicies: this.policies.filter(p => p.enabled).length,
      cachedResults: this.scanResults.size
    };
  }

  /**
   * Initialize default sensitive data patterns
   */
  private initializeDefaultPatterns(): void {
    this.patterns = [
      // Credit card numbers
      {
        id: 'cc-visa',
        name: 'Visa Credit Card',
        pattern: /\b4[0-9]{12}(?:[0-9]{3})?\b/g,
        severity: 'high',
        category: 'Financial',
        description: 'Visa credit card number'
      },
      {
        id: 'cc-mastercard',
        name: 'Mastercard Credit Card',
        pattern: /\b5[1-5][0-9]{14}\b/g,
        severity: 'high',
        category: 'Financial',
        description: 'Mastercard credit card number'
      },
      {
        id: 'cc-amex',
        name: 'American Express Credit Card',
        pattern: /\b3[47][0-9]{13}\b/g,
        severity: 'high',
        category: 'Financial',
        description: 'American Express credit card number'
      },
      
      // Social Security Numbers
      {
        id: 'ssn',
        name: 'Social Security Number',
        pattern: /\b\d{3}-\d{2}-\d{4}\b/g,
        severity: 'critical',
        category: 'Personal Identification',
        description: 'US Social Security Number'
      },
      
      // Email addresses
      {
        id: 'email',
        name: 'Email Address',
        pattern: /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b/g,
        severity: 'low',
        category: 'Personal Information',
        description: 'Email address'
      },
      
      // Phone numbers
      {
        id: 'phone-us',
        name: 'US Phone Number',
        pattern: /\b(\+1[-.\s]?)?\(?([0-9]{3})\)?[-.\s]?([0-9]{3})[-.\s]?([0-9]{4})\b/g,
        severity: 'medium',
        category: 'Personal Information',
        description: 'US phone number'
      },
      
      // API Keys and secrets
      {
        id: 'api-key',
        name: 'API Key',
        pattern: /\b[A-Za-z0-9_]{32}\b/g,
        severity: 'high',
        category: 'Credentials',
        description: 'Generic API key pattern'
      },
      {
        id: 'aws-access-key',
        name: 'AWS Access Key',
        pattern: /\bAKIA[0-9A-Z]{16}\b/g,
        severity: 'critical',
        category: 'Credentials',
        description: 'AWS Access Key ID'
      },
      
      // Passwords in text
      {
        id: 'password',
        name: 'Password in Text',
        pattern: /\b(password|passwd|pwd)\s*[:=]\s*['"][^'"]{4,}['"]/gi,
        severity: 'critical',
        category: 'Credentials',
        description: 'Password in plain text'
      },
      
      // Passport numbers
      {
        id: 'passport-us',
        name: 'US Passport Number',
        pattern: /\b[CNP]\d{8}\b/g,
        severity: 'high',
        category: 'Personal Identification',
        description: 'US Passport number'
      },
      
      // Driver's license
      {
        id: 'drivers-license',
        name: 'Driver\'s License Number',
        pattern: /\b[A-Z]{1,2}\d{4,8}\b/g,
        severity: 'high',
        category: 'Personal Identification',
        description: 'Driver\'s license number'
      }
    ];
  }

  /**
   * Initialize default DLP policies
   */
  private initializeDefaultPolicies(): void {
    this.policies = [
      {
        id: 'default-policy',
        name: 'Default DLP Policy',
        description: 'Standard policy for detecting sensitive information',
        enabled: true,
        patterns: this.patterns.map(p => p.id),
        severityThreshold: 'medium',
        actions: ['log', 'alert']
      },
      {
        id: 'strict-policy',
        name: 'Strict DLP Policy',
        description: 'Strict policy that blocks critical findings',
        enabled: false, // Disabled by default
        patterns: this.patterns.map(p => p.id),
        severityThreshold: 'high',
        actions: ['log', 'alert', 'block']
      }
    ];
  }

  /**
   * Scan document content for sensitive data
   */
  async scanDocument(file: DriveFile, content: string): Promise<DlpScanResult> {
    try {
      // Check if we have a cached result
      const cacheKey = `${file.id}-${file.modifiedTime}`;
      const cachedResult = this.scanResults.get(cacheKey);
      
      if (cachedResult) {
        logger.debug('Returning cached DLP scan result', {
          component: 'DataLossPreventionService',
          fileId: file.id
        });
        return cachedResult;
      }

      // Get active policies
      const activePolicies = this.policies.filter(policy => policy.enabled);
      
      // Get all patterns from active policies
      const policyPatternIds = new Set(
        activePolicies.flatMap(policy => policy.patterns)
      );
      
      const applicablePatterns = this.patterns.filter(
        pattern => policyPatternIds.has(pattern.id)
      );

      // Scan content for sensitive data
      const findings: SensitiveDataFinding[] = [];
      
      for (const pattern of applicablePatterns) {
        const matches = content.matchAll(pattern.pattern);
        
        for (const match of matches) {
          if (match.index !== undefined) {
            // Extract context (100 characters before and after)
            const start = Math.max(0, match.index - 100);
            const end = Math.min(content.length, match.index + match[0].length + 100);
            const context = content.substring(start, end);
            
            const finding: SensitiveDataFinding = {
              id: this.generateId(),
              patternId: pattern.id,
              patternName: pattern.name,
              severity: pattern.severity,
              category: pattern.category,
              matchedText: match[0],
              position: {
                start: match.index,
                end: match.index + match[0].length
              },
              context,
              confidence: this.calculateConfidence(pattern, match[0])
            };
            
            findings.push(finding);
          }
        }
      }

      // Generate statistics
      const findingsBySeverity: Record<string, number> = {};
      const findingsByCategory: Record<string, number> = {};
      
      findings.forEach(finding => {
        findingsBySeverity[finding.severity] = (findingsBySeverity[finding.severity] || 0) + 1;
        findingsByCategory[finding.category] = (findingsByCategory[finding.category] || 0) + 1;
      });

      // Calculate risk score
      const riskScore = this.calculateRiskScore(findings);
      
      // Determine recommended actions
      const recommendedActions = this.determineRecommendedActions(findings, activePolicies);

      // Create scan result
      const scanResult: DlpScanResult = {
        fileId: file.id,
        fileName: file.name || 'Untitled',
        scannedAt: new Date(),
        totalFindings: findings.length,
        findingsBySeverity,
        findingsByCategory,
        findings,
        riskScore,
        recommendedActions
      };

      // Cache the result
      this.cacheScanResult(cacheKey, scanResult);
      
      logger.info('DLP scan completed', {
        component: 'DataLossPreventionService',
        fileId: file.id,
        fileName: file.name,
        totalFindings: findings.length,
        riskScore
      });

      return scanResult;
    } catch (error) {
      logger.error('Error scanning document for sensitive data', {
        component: 'DataLossPreventionService',
        fileId: file.id,
        error: error instanceof Error ? error.message : String(error)
      });
      
      throw error;
    }
  }

  /**
   * Calculate confidence level for a finding
   */
  private calculateConfidence(pattern: SensitiveDataPattern, matchedText: string): number {
    // Base confidence from pattern
    const baseConfidence: Record<string, number> = {
      'low': 30,
      'medium': 50,
      'high': 70,
      'critical': 90
    };
    
    let confidence = baseConfidence[pattern.severity] || 50;
    
    // Adjust based on match characteristics
    if (matchedText.length > 10) {
      confidence += 10; // Longer matches are more likely to be real
    }
    
    // Adjust for specific patterns
    if (pattern.id.includes('credit') || pattern.id.includes('card')) {
      // Credit cards have specific checksums
      if (this.isValidCreditCard(matchedText)) {
        confidence += 20;
      } else {
        confidence -= 30;
      }
    }
    
    // Ensure confidence is within bounds
    return Math.max(0, Math.min(100, confidence));
  }

  /**
   * Basic credit card validation using Luhn algorithm
   */
  private isValidCreditCard(cardNumber: string): boolean {
    // Remove non-digit characters
    const cleaned = cardNumber.replace(/\D/g, '');
    
    // Check length
    if (cleaned.length < 13 || cleaned.length > 19) {
      return false;
    }
    
    // Luhn algorithm
    let sum = 0;
    let isEven = false;
    
    for (let i = cleaned.length - 1; i >= 0; i--) {
      // Check if the character exists before parsing
      const char = cleaned[i];
      if (char !== undefined) {
        const digit = parseInt(char, 10);
        
        // Check if parsing was successful
        if (!isNaN(digit)) {
          let adjustedDigit = digit;
          
          if (isEven) {
            adjustedDigit *= 2;
            if (adjustedDigit > 9) {
              adjustedDigit -= 9;
            }
          }
          
          sum += adjustedDigit;
          isEven = !isEven;
        }
      }
    }
    
    return sum % 10 === 0;
  }

  /**
   * Calculate overall risk score for a document
   */
  private calculateRiskScore(findings: SensitiveDataFinding[]): number {
    if (findings.length === 0) {
      return 0;
    }
    
    // Weight findings by severity
    const severityWeights: Record<string, number> = {
      'low': 1,
      'medium': 3,
      'high': 7,
      'critical': 15
    };
    
    let weightedScore = 0;
    let maxPossibleScore = 0;
    
    findings.forEach(finding => {
      const weight = severityWeights[finding.severity] || 1;
      weightedScore += weight * (finding.confidence / 100);
      maxPossibleScore += weight;
    });
    
    // Normalize to 0-100 scale
    if (maxPossibleScore === 0) {
      return 0;
    }
    
    return Math.round((weightedScore / maxPossibleScore) * 100);
  }

  /**
   * Determine recommended actions based on findings and policies
   */
  private determineRecommendedActions(
    findings: SensitiveDataFinding[],
    policies: DlpPolicy[]
  ): string[] {
    const actions = new Set<string>();
    
    // Collect actions from all applicable policies
    for (const policy of policies) {
      // Check if policy threshold is met
      const hasFindingsAboveThreshold = findings.some(finding => 
        this.isSeverityAtLeast(finding.severity, policy.severityThreshold)
      );
      
      if (hasFindingsAboveThreshold) {
        policy.actions.forEach(action => actions.add(action));
      }
    }
    
    return Array.from(actions);
  }

  /**
   * Check if a severity meets or exceeds a threshold
   */
  private isSeverityAtLeast(severity: string, threshold: string): boolean {
    const severityLevels: Record<string, number> = {
      'low': 1,
      'medium': 2,
      'high': 3,
      'critical': 4
    };
    
    return (severityLevels[severity] || 0) >= (severityLevels[threshold] || 0);
  }

  /**
   * Cache scan result with size management
   */
  private cacheScanResult(key: string, result: DlpScanResult): void {
    // Remove oldest entries if we're at capacity
    if (this.scanResults.size >= this.MAX_CACHE_RESULTS) {
      const firstKey = this.scanResults.keys().next().value;
      if (firstKey) {
        this.scanResults.delete(firstKey);
      }
    }
    
    this.scanResults.set(key, result);
  }

  /**
   * Get cached scan result
   */
  getScanResult(fileId: string, modifiedTime?: string): DlpScanResult | null {
    const cacheKey = modifiedTime ? `${fileId}-${modifiedTime}` : fileId;
    return this.scanResults.get(cacheKey) || null;
  }

  /**
   * Clear cached scan results for a document
   */
  clearScanResults(fileId: string): boolean {
    let deleted = false;
    for (const key of this.scanResults.keys()) {
      if (key.startsWith(fileId)) {
        this.scanResults.delete(key);
        deleted = true;
      }
    }
    return deleted;
  }

  /**
   * Add a custom sensitive data pattern
   */
  addPattern(pattern: SensitiveDataPattern): void {
    this.patterns.push(pattern);
    logger.info('Custom DLP pattern added', {
      component: 'DataLossPreventionService',
      patternId: pattern.id,
      patternName: pattern.name
    });
  }

  /**
   * Remove a sensitive data pattern
   */
  removePattern(patternId: string): boolean {
    const initialLength = this.patterns.length;
    this.patterns = this.patterns.filter(pattern => pattern.id !== patternId);
    
    // Also remove from policies
    this.policies.forEach(policy => {
      policy.patterns = policy.patterns.filter(id => id !== patternId);
    });
    
    const removed = this.patterns.length < initialLength;
    
    if (removed) {
      logger.info('DLP pattern removed', {
        component: 'DataLossPreventionService',
        patternId
      });
    }
    
    return removed;
  }

  /**
   * Add a DLP policy
   */
  addPolicy(policy: DlpPolicy): void {
    this.policies.push(policy);
    logger.info('DLP policy added', {
      component: 'DataLossPreventionService',
      policyId: policy.id,
      policyName: policy.name
    });
  }

  /**
   * Update a DLP policy
   */
  updatePolicy(policyId: string, updates: Partial<DlpPolicy>): DlpPolicy | null {
    const policy = this.policies.find(p => p.id === policyId);
    
    if (policy) {
      Object.assign(policy, updates);
      logger.info('DLP policy updated', {
        component: 'DataLossPreventionService',
        policyId
      });
      return policy;
    }
    
    return null;
  }

  /**
   * Remove a DLP policy
   */
  removePolicy(policyId: string): boolean {
    const initialLength = this.policies.length;
    this.policies = this.policies.filter(policy => policy.id !== policyId);
    
    const removed = this.policies.length < initialLength;
    
    if (removed) {
      logger.info('DLP policy removed', {
        component: 'DataLossPreventionService',
        policyId
      });
    }
    
    return removed;
  }

  /**
   * Enable or disable a policy
   */
  setPolicyEnabled(policyId: string, enabled: boolean): boolean {
    const policy = this.policies.find(p => p.id === policyId);
    
    if (policy) {
      policy.enabled = enabled;
      logger.info('DLP policy enabled status updated', {
        component: 'DataLossPreventionService',
        policyId,
        enabled
      });
      return true;
    }
    
    return false;
  }

  /**
   * Get service statistics
   */
  public override getStats(): ServiceStats {
    // Get base stats from parent class
    const baseStats = super.getStats();
    
    const activePolicies = this.policies.filter(policy => policy.enabled).length;
    
    const results = Array.from(this.scanResults.values());
    const totalRiskScore = results.reduce((sum, result) => sum + result.riskScore, 0);
    const averageRiskScore = results.length > 0 ? Math.round(totalRiskScore / results.length) : 0;
    
    return {
      ...baseStats,
      totalPatterns: this.patterns.length,
      activePolicies,
      cachedResults: this.scanResults.size,
      averageRiskScore
    };
  }

  /**
   * Generate a unique ID
   */
  private generateId(): string {
    return `dlp-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
  }
}