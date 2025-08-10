/**
 * Статистика бота
 * TypeScript версія
 */

import fs from 'fs';
import path from 'path';

interface CommandStats {
  total: number;
  successful: number;
  failed: number;
  users: Set<string>;
  lastUsed: string | null;
}

interface UserStats {
  commands: Record<string, number>;
  firstSeen: string;
  lastSeen: string;
  totalCommands: number;
}

interface DailyStats {
  commands: number;
  users: Set<string>;
  errors: number;
}

interface ErrorEntry {
  message: string;
  stack?: string;
  timestamp: string;
  commandName?: string;
  userId?: string;
}

interface BotStatsData {
  startTime: string;
  commands: Record<string, CommandStats>;
  users: Record<string, UserStats>;
  errors: ErrorEntry[];
  totalCommands: number;
  totalUsers: number;
  dailyStats: Record<string, DailyStats>;
}

class BotStats {
  private statsFile: string;
  private stats: BotStatsData;
  private saveTimer: NodeJS.Timeout | null = null;

  constructor() {
    this.statsFile = path.join('data', 'logs', 'stats.json');
    this.stats = this.loadStats();
  }

  private loadStats(): BotStatsData {
    try {
      if (fs.existsSync(this.statsFile)) {
        const data = fs.readFileSync(this.statsFile, 'utf8');
        const parsed = JSON.parse(data);
        
        // Конвертуємо Set з JSON назад в Set
        for (const commandName in parsed.commands) {
          if (parsed.commands[commandName].users && Array.isArray(parsed.commands[commandName].users)) {
            parsed.commands[commandName].users = new Set(parsed.commands[commandName].users);
          }
        }
        
        for (const date in parsed.dailyStats) {
          if (parsed.dailyStats[date].users && Array.isArray(parsed.dailyStats[date].users)) {
            parsed.dailyStats[date].users = new Set(parsed.dailyStats[date].users);
          }
        }
        
        return parsed;
      }
    } catch (error) {
      console.error('Помилка завантаження статистики:', error);
    }

    // Створюємо нову статистику
    return {
      startTime: new Date().toISOString(),
      commands: {},
      users: {},
      errors: [],
      totalCommands: 0,
      totalUsers: 0,
      dailyStats: {}
    };
  }

  private saveStats(): void {
    try {
      const statsDir = path.dirname(this.statsFile);
      if (!fs.existsSync(statsDir)) {
        fs.mkdirSync(statsDir, { recursive: true });
      }

      // Глубокая копия и конверсия Set → Array
      const copy: any = JSON.parse(JSON.stringify(this.stats));
      for (const name in copy.commands) {
        const users = this.stats.commands[name]?.users;
        if (users instanceof Set) copy.commands[name].users = Array.from(users);
      }
      for (const date in copy.dailyStats) {
        const users = this.stats.dailyStats[date]?.users;
        if (users instanceof Set) copy.dailyStats[date].users = Array.from(users);
      }

      // Атомарная запись
      const tmp = this.statsFile + '.tmp';
      fs.writeFileSync(tmp, JSON.stringify(copy, null, 2));
      fs.renameSync(tmp, this.statsFile);
    } catch (error) {
      console.error('Помилка збереження статистики:', error);
    }
  }

  private scheduleSave(): void {
    if (this.saveTimer) clearTimeout(this.saveTimer);
    this.saveTimer = setTimeout(() => {
      this.saveTimer = null;
      this.saveStats();
    }, 250);
  }

  trackCommand(commandName: string, userId: string, _guildId?: string, success: boolean = true): void {
    const [today] = new Date().toISOString().split('T') as [string, string];

    // Оновлюємо статистику команд
    if (!this.stats.commands[commandName]) {
      this.stats.commands[commandName] = {
        total: 0,
        successful: 0,
        failed: 0,
        users: new Set(),
        lastUsed: null
      };
    }

    this.stats.commands[commandName].total++;
    this.stats.commands[commandName].users.add(userId);
    this.stats.commands[commandName].lastUsed = new Date().toISOString();

    if (success) {
      this.stats.commands[commandName].successful++;
    } else {
      this.stats.commands[commandName].failed++;
    }

    // Оновлюємо статистику користувачів
    if (!this.stats.users[userId]) {
      this.stats.users[userId] = {
        commands: {},
        firstSeen: new Date().toISOString(),
        lastSeen: new Date().toISOString(),
        totalCommands: 0
      };
    }

    if (!this.stats.users[userId].commands[commandName]) {
      this.stats.users[userId].commands[commandName] = 0;
    }

    this.stats.users[userId].commands[commandName]++;
    this.stats.users[userId].lastSeen = new Date().toISOString();
    this.stats.users[userId].totalCommands++;

    // Оновлюємо загальну статистику
    this.stats.totalCommands++;
    this.stats.totalUsers = Object.keys(this.stats.users).length;

    // Оновлюємо денну статистику
    if (!this.stats.dailyStats[today]) {
      this.stats.dailyStats[today] = {
        commands: 0,
        users: new Set(),
        errors: 0
      };
    }

    this.stats.dailyStats[today].commands++;
    this.stats.dailyStats[today].users.add(userId);

    // Зберігаємо статистику
    this.scheduleSave();
  }

  trackError(error: Error | string, commandName?: string, userId?: string): void {
    const errorEntry = {
      message: error instanceof Error ? error.message : String(error),
      timestamp: new Date().toISOString(),
      ...(commandName ? { commandName } : {}),
      ...(userId ? { userId } : {}),
      ...(error instanceof Error && error.stack ? { stack: error.stack } : {}),
    } satisfies ErrorEntry;

    this.stats.errors.push(errorEntry);

    // Обмежуємо кількість помилок в пам'яті
    if (this.stats.errors.length > 1000) {
      this.stats.errors = this.stats.errors.slice(-1000);
    }

    // Оновлюємо денну статистику помилок
    const [today] = new Date().toISOString().split('T') as [string, string];
    if (!this.stats.dailyStats[today]) {
      this.stats.dailyStats[today] = { commands: 0, users: new Set(), errors: 0 };
    }
    this.stats.dailyStats[today].errors++;

    this.scheduleSave();
  }

  getStats(): BotStatsData {
    return { ...this.stats };
  }

  getDailyStats(): Record<string, DailyStats> {
    const dailyStatsCopy: Record<string, DailyStats> = {};
    
    for (const [date, stats] of Object.entries(this.stats.dailyStats)) {
      dailyStatsCopy[date] = {
        ...stats,
        users: new Set(stats.users)
      };
    }
    
    return dailyStatsCopy;
  }

  getTopCommands(limit: number = 5): Array<{ name: string; stats: CommandStats }> {
    return Object.entries(this.stats.commands)
      .map(([name, stats]) => ({ name, stats }))
      .sort((a, b) => b.stats.total - a.stats.total)
      .slice(0, limit);
  }

  getTopUsers(limit: number = 5): Array<{ userId: string; stats: UserStats }> {
    return Object.entries(this.stats.users)
      .map(([userId, stats]) => ({ userId, stats }))
      .sort((a, b) => b.stats.totalCommands - a.stats.totalCommands)
      .slice(0, limit);
  }

  formatUptime(ms: number): string {
    const days = Math.floor(ms / (1000 * 60 * 60 * 24));
    const hours = Math.floor((ms % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((ms % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((ms % (1000 * 60)) / 1000);

    const parts: string[] = [];
    if (days > 0) parts.push(`${days}д`);
    if (hours > 0) parts.push(`${hours}г`);
    if (minutes > 0) parts.push(`${minutes}хв`);
    if (seconds > 0) parts.push(`${seconds}с`);

    return parts.join(' ') || '0с';
  }

  resetStats(): void {
    this.stats = {
      startTime: new Date().toISOString(),
      commands: {},
      users: {},
      errors: [],
      totalCommands: 0,
      totalUsers: 0,
      dailyStats: {}
    };
    this.saveStats();
  }

  getCommandSuccessRate(commandName: string): number {
    const command = this.stats.commands[commandName];
    if (!command || command.total === 0) return 0;
    return (command.successful / command.total) * 100;
  }

  getOverallSuccessRate(): number {
    if (this.stats.totalCommands === 0) return 0;
    
    const totalSuccessful = Object.values(this.stats.commands)
      .reduce((sum, command) => sum + command.successful, 0);
    
    return (totalSuccessful / this.stats.totalCommands) * 100;
  }

  getActiveUsers(days: number = 7): number {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - days);
    
    return Object.values(this.stats.users)
      .filter(user => new Date(user.lastSeen) > cutoffDate)
      .length;
  }
}

export default BotStats;
export { BotStats }; 