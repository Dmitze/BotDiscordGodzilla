/**
 * MobileOptimizationService
 * - Provides utilities for optimizing Discord UI for mobile users
 * - Adjusts embed layouts, component counts, and formatting for smaller screens
 * - Manages mobile-specific preferences and settings
 */

import type { EmbedBuilder, ActionRowBuilder, ButtonBuilder, StringSelectMenuBuilder } from 'discord.js';
import logger from '@/utils/logger';

export interface MobileOptimizationConfig {
  userId: string;
  enabled: boolean;
  compactMode: boolean;
  maxComponentsPerRow: number;
  maxActionRows: number;
  fontSize: 'small' | 'normal' | 'large';
  contrastMode: boolean;
}

// In-memory store for user preferences
const userMobilePrefs = new Map<string, MobileOptimizationConfig>();

export class MobileOptimizationService {
  /**
   * Get user mobile optimization preferences
   */
  getUserPreferences(userId: string): MobileOptimizationConfig {
    const existing = userMobilePrefs.get(userId);
    if (existing) return existing;
    
    // Create default preferences optimized for mobile
    const prefs: MobileOptimizationConfig = {
      userId,
      enabled: true, // Enabled by default for better UX
      compactMode: true, // Compact mode for smaller screens
      maxComponentsPerRow: 3, // Limit components per row for touch targets
      maxActionRows: 3, // Limit action rows to avoid scrolling
      fontSize: 'normal', // Normal font size
      contrastMode: false // Disabled by default
    };
    
    userMobilePrefs.set(userId, prefs);
    return prefs;
  }

  /**
   * Set user mobile optimization preferences
   */
  setUserPreferences(userId: string, preferences: MobileOptimizationConfig): void {
    userMobilePrefs.set(userId, preferences);
    logger.debug('User mobile optimization preferences updated', {
      component: 'MobileOptimizationService',
      userId,
      enabled: preferences.enabled
    });
  }

  /**
   * Enable or disable mobile optimization for a user
   */
  setEnabled(userId: string, enabled: boolean): void {
    const prefs = this.getUserPreferences(userId);
    prefs.enabled = enabled;
    this.setUserPreferences(userId, prefs);
  }

  /**
   * Optimize embed for mobile viewing
   */
  optimizeEmbed(embed: EmbedBuilder, userId: string): EmbedBuilder {
    const prefs = this.getUserPreferences(userId);
    
    if (!prefs.enabled) {
      return embed;
    }
    
    // Apply compact mode formatting
    if (prefs.compactMode) {
      const description = embed.data.description;
      if (description) {
        // Remove extra whitespace and newlines for compact display
        const compactDescription = description
          .replace(/\n\s*\n/g, '\n') // Remove extra blank lines
          .replace(/\s+/g, ' ') // Reduce multiple spaces
          .trim();
        
        embed.setDescription(compactDescription);
      }
      
      // Reduce field count if needed
      const fields = embed.data.fields || [];
      if (fields.length > 5) {
        // Keep first 5 fields and add a "more" indicator
        const trimmedFields = fields.slice(0, 5);
        trimmedFields.push({
          name: '...',
          value: `*${fields.length - 5} more items available*`,
          inline: false
        });
        embed.setFields(trimmedFields);
      }
    }
    
    // Apply contrast mode if enabled
    if (prefs.contrastMode) {
      // Use high-contrast colors
      embed.setColor(0xFFFFFF); // White background for high contrast
    }
    
    return embed;
  }

  /**
   * Optimize action rows for mobile touch interaction
   */
  optimizeActionRows(
    components: ActionRowBuilder<ButtonBuilder | StringSelectMenuBuilder>[],
    userId: string
  ): ActionRowBuilder<ButtonBuilder | StringSelectMenuBuilder>[] {
    const prefs = this.getUserPreferences(userId);
    
    if (!prefs.enabled) {
      return components;
    }
    
    // Limit the number of action rows
    let optimizedComponents = [...components];
    if (optimizedComponents.length > prefs.maxActionRows) {
      optimizedComponents = optimizedComponents.slice(0, prefs.maxActionRows);
    }
    
    // Optimize each action row
    for (const row of optimizedComponents) {
      // Limit components per row for better touch targets
      if (row.components.length > prefs.maxComponentsPerRow) {
        // Keep only the most important components
        row.components = row.components.slice(0, prefs.maxComponentsPerRow);
      }
      
      // For buttons, ensure they have appropriate labels for mobile
      const buttons = row.components.filter(component => 
        component.type === 2 // Button type
      ) as ButtonBuilder[];
      
      for (const button of buttons) {
        const label = button.data.label || '';
        
        // Truncate long labels for mobile screens
        if (label.length > 15) {
          button.setLabel(`${label.substring(0, 12)}...`);
        }
      }
    }
    
    return optimizedComponents;
  }

  /**
   * Create a mobile-friendly pagination system
   */
  createMobilePagination(
    currentPage: number,
    totalPages: number,
    customIds: Record<string, string> = {}
  ): ActionRowBuilder<ButtonBuilder> {
    // For mobile, use a simplified pagination with fewer buttons
    const row = new ActionRowBuilder<ButtonBuilder>();
    
    // Previous button
    row.addComponents(
      new ButtonBuilder()
        .setCustomId(customIds['prev'] ?? 'prev_page')
        .setLabel('⬅️')
        .setStyle(1) // Primary
        .setDisabled(currentPage === 1)
    );
    
    // Page info button (disabled, shows current page)
    row.addComponents(
      new ButtonBuilder()
        .setCustomId('page_info')
        .setLabel(`${currentPage}/${totalPages}`)
        .setStyle(2) // Secondary
        .setDisabled(true)
    );
    
    // Next button
    row.addComponents(
      new ButtonBuilder()
        .setCustomId(customIds['next'] ?? 'next_page')
        .setLabel('➡️')
        .setStyle(1) // Primary
        .setDisabled(currentPage >= totalPages)
    );
    
    return row;
  }

  /**
   * Optimize text content for mobile screens
   */
  optimizeTextContent(text: string, userId: string): string {
    const prefs = this.getUserPreferences(userId);
    
    if (!prefs.enabled) {
      return text;
    }
    
    // For mobile, we want to ensure text isn't too wide
    const lines = text.split('\n');
    const optimizedLines = lines.map(line => {
      // Truncate very long lines
      if (line.length > 80) {
        return line.substring(0, 77) + '...';
      }
      return line;
    });
    
    return optimizedLines.join('\n');
  }

  /**
   * Get mobile optimization status for a user
   */
  getStatus(userId: string): { enabled: boolean; mode: string } {
    const prefs = this.getUserPreferences(userId);
    
    if (!prefs.enabled) {
      return { enabled: false, mode: 'disabled' };
    }
    
    const mode = prefs.compactMode ? 'compact' : 'normal';
    return { enabled: true, mode };
  }

  /**
   * Apply all mobile optimizations to a message
   */
  applyAllOptimizations(
    embed: EmbedBuilder,
    components: ActionRowBuilder<ButtonBuilder | StringSelectMenuBuilder>[],
    userId: string
  ): {
    embed: EmbedBuilder;
    components: ActionRowBuilder<ButtonBuilder | StringSelectMenuBuilder>[];
  } {
    const optimizedEmbed = this.optimizeEmbed(embed, userId);
    const optimizedComponents = this.optimizeActionRows(components, userId);
    
    return {
      embed: optimizedEmbed,
      components: optimizedComponents
    };
  }

  /**
   * For tests: reset in-memory store
   */
  __reset(): void {
    userMobilePrefs.clear();
  }
}

export const mobileOptimizationService = new MobileOptimizationService();