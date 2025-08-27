/**
 * KeyboardNavigationService
 * - Provides utilities for implementing keyboard navigation shortcuts in Discord UI
 * - Adds numbered labels to buttons for easy keyboard access
 * - Manages keyboard shortcut hints and documentation
 */

import type { ActionRowBuilder, ButtonBuilder, StringSelectMenuBuilder } from 'discord.js';
import logger from '@/utils/logger';

export interface KeyboardShortcut {
  key: string;
  description: string;
  action: string;
}

export interface KeyboardNavigationConfig {
  userId: string;
  enabled: boolean;
  shortcuts: KeyboardShortcut[];
  showHints: boolean;
}

// In-memory store for user preferences
const userKeyboardPrefs = new Map<string, KeyboardNavigationConfig>();

export class KeyboardNavigationService {
  /**
   * Get user keyboard navigation preferences
   */
  getUserPreferences(userId: string): KeyboardNavigationConfig {
    const existing = userKeyboardPrefs.get(userId);
    if (existing) return existing;
    
    // Create default preferences
    const prefs: KeyboardNavigationConfig = {
      userId,
      enabled: true,
      showHints: true,
      shortcuts: [
        { key: '1-9', description: 'Select item by number', action: 'select_item' },
        { key: '↑/↓', description: 'Navigate between items', action: 'navigate_items' },
        { key: 'Enter', description: 'Confirm selection', action: 'confirm' },
        { key: 'Esc', description: 'Cancel/close', action: 'cancel' },
        { key: 'R', description: 'Refresh view', action: 'refresh' },
        { key: 'S', description: 'Search', action: 'search' },
        { key: 'N', description: 'Next page', action: 'next_page' },
        { key: 'P', description: 'Previous page', action: 'prev_page' }
      ]
    };
    
    userKeyboardPrefs.set(userId, prefs);
    return prefs;
  }

  /**
   * Set user keyboard navigation preferences
   */
  setUserPreferences(userId: string, preferences: KeyboardNavigationConfig): void {
    userKeyboardPrefs.set(userId, preferences);
    logger.debug('User keyboard navigation preferences updated', {
      component: 'KeyboardNavigationService',
      userId,
      enabled: preferences.enabled
    });
  }

  /**
   * Enable or disable keyboard navigation for a user
   */
  setEnabled(userId: string, enabled: boolean): void {
    const prefs = this.getUserPreferences(userId);
    prefs.enabled = enabled;
    this.setUserPreferences(userId, prefs);
  }

  /**
   * Add numbered labels to buttons for keyboard navigation
   */
  addNumberedLabels(
    components: ActionRowBuilder<ButtonBuilder | StringSelectMenuBuilder>[]
  ): ActionRowBuilder<ButtonBuilder | StringSelectMenuBuilder>[] {
    // Create a copy of the components to avoid modifying the original
    const labeledComponents = [...components];
    
    // Add numbers to button labels
    let buttonIndex = 1;
    for (const row of labeledComponents) {
      // Skip select menus, only process buttons
      const buttons = row.components.filter(component => 
        component.type === 2 // Button type
      ) as ButtonBuilder[];
      
      for (const button of buttons) {
        if (buttonIndex <= 9) {
          const currentLabel = button.data.label || '';
          // Only add number if it's not already numbered
          if (!/^[1-9]\./.test(currentLabel)) {
            button.setLabel(`${buttonIndex}. ${currentLabel}`);
          }
          buttonIndex++;
        }
      }
    }
    
    return labeledComponents;
  }

  /**
   * Add keyboard shortcut hints to an embed description
   */
  addShortcutHints(description: string, userId: string): string {
    const prefs = this.getUserPreferences(userId);
    
    if (!prefs.showHints || !prefs.enabled) {
      return description;
    }
    
    const shortcutLines = prefs.shortcuts.map(shortcut => 
      `**${shortcut.key}** — ${shortcut.description}`
    );
    
    const hintsSection = `\n\n**⌨️ Keyboard Shortcuts:**\n${shortcutLines.join('\n')}`;
    
    // Limit the total length to avoid Discord limits
    const maxLength = 4096 - hintsSection.length - 100; // Leave some buffer
    const trimmedDescription = description.length > maxLength 
      ? description.substring(0, maxLength) + '...' 
      : description;
    
    return trimmedDescription + hintsSection;
  }

  /**
   * Get available keyboard shortcuts for a user
   */
  getShortcuts(userId: string): KeyboardShortcut[] {
    const prefs = this.getUserPreferences(userId);
    return prefs.shortcuts;
  }

  /**
   * Add a custom shortcut for a user
   */
  addShortcut(userId: string, shortcut: KeyboardShortcut): void {
    const prefs = this.getUserPreferences(userId);
    
    // Check if shortcut already exists
    const existingIndex = prefs.shortcuts.findIndex(s => s.key === shortcut.key);
    
    if (existingIndex >= 0) {
      // Update existing shortcut
      prefs.shortcuts[existingIndex] = shortcut;
    } else {
      // Add new shortcut
      prefs.shortcuts.push(shortcut);
    }
    
    this.setUserPreferences(userId, prefs);
  }

  /**
   * Remove a shortcut for a user
   */
  removeShortcut(userId: string, key: string): boolean {
    const prefs = this.getUserPreferences(userId);
    const initialLength = prefs.shortcuts.length;
    
    prefs.shortcuts = prefs.shortcuts.filter(shortcut => shortcut.key !== key);
    
    this.setUserPreferences(userId, prefs);
    
    return prefs.shortcuts.length < initialLength;
  }

  /**
   * Generate help text for keyboard navigation
   */
  generateHelpText(userId: string): string {
    const prefs = this.getUserPreferences(userId);
    
    if (!prefs.enabled) {
      return 'Keyboard navigation is currently disabled. Enable it with `/keyboard enable`';
    }
    
    const shortcutLines = prefs.shortcuts.map(shortcut => 
      `**${shortcut.key}** — ${shortcut.description}`
    );
    
    return `**⌨️ Keyboard Navigation Help**\n\n${shortcutLines.join('\n')}\n\nUse these shortcuts to navigate the interface more efficiently.`;
  }

  /**
   * For tests: reset in-memory store
   */
  __reset(): void {
    userKeyboardPrefs.clear();
  }
}

export const keyboardNavigationService = new KeyboardNavigationService();