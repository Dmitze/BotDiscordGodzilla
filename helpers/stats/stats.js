const fs = require('fs');
const path = require('path');

class BotStats {
  constructor() {
    this.statsFile = './logs/stats.json';
    this.stats = this.loadStats();
  }

  loadStats() {
    try {
      if (fs.existsSync(this.statsFile)) {
        const data = fs.readFileSync(this.statsFile, 'utf8');
        return JSON.parse(data);
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

  saveStats() {
    try {
      const statsDir = path.dirname(this.statsFile);
      if (!fs.existsSync(statsDir)) {
        fs.mkdirSync(statsDir, { recursive: true });
      }
      fs.writeFileSync(this.statsFile, JSON.stringify(this.stats, null, 2));
    } catch (error) {
      console.error('Помилка збереження статистики:', error);
    }
  }

  trackCommand(commandName, userId, guildId, success = true) {
    const today = new Date().toISOString().split('T')[0];
    
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

    this.saveStats();
  }

  trackError(error, commandName = null, userId = null) {
    const errorEntry = {
      timestamp: new Date().toISOString(),
      error: error.message || error,
      command: commandName,
      userId: userId,
      stack: error.stack
    };

    this.stats.errors.push(errorEntry);
    
    // Зберігаємо тільки останні 100 помилок
    if (this.stats.errors.length > 100) {
      this.stats.errors = this.stats.errors.slice(-100);
    }

    if (commandName) {
      this.trackCommand(commandName, userId, null, false);
    }

    this.saveStats();
  }

  getStats() {
    const now = new Date();
    const uptime = now - new Date(this.stats.startTime);

    return {
      uptime: this.formatUptime(uptime),
      totalCommands: this.stats.totalCommands,
      totalUsers: this.stats.totalUsers,
      commands: this.stats.commands,
      recentErrors: this.stats.errors.slice(-10),
      dailyStats: this.getDailyStats()
    };
  }

  getDailyStats() {
    const today = new Date().toISOString().split('T')[0];
    const yesterday = new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString().split('T')[0];

    return {
      today: this.stats.dailyStats[today] || { commands: 0, users: 0, errors: 0 },
      yesterday: this.stats.dailyStats[yesterday] || { commands: 0, users: 0, errors: 0 }
    };
  }

  getTopCommands(limit = 5) {
    return Object.entries(this.stats.commands)
      .sort(([,a], [,b]) => b.total - a.total)
      .slice(0, limit)
      .map(([name, stats]) => ({
        name,
        total: stats.total,
        successful: stats.successful,
        failed: stats.failed,
        uniqueUsers: stats.users.size
      }));
  }

  getTopUsers(limit = 5) {
    return Object.entries(this.stats.users)
      .sort(([,a], [,b]) => b.totalCommands - a.totalCommands)
      .slice(0, limit)
      .map(([userId, stats]) => ({
        userId,
        totalCommands: stats.totalCommands,
        commands: stats.commands,
        firstSeen: stats.firstSeen,
        lastSeen: stats.lastSeen
      }));
  }

  formatUptime(ms) {
    const days = Math.floor(ms / (1000 * 60 * 60 * 24));
    const hours = Math.floor((ms % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
    const minutes = Math.floor((ms % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((ms % (1000 * 60)) / 1000);

    const parts = [];
    if (days > 0) parts.push(`${days}д`);
    if (hours > 0) parts.push(`${hours}г`);
    if (minutes > 0) parts.push(`${minutes}хв`);
    if (seconds > 0) parts.push(`${seconds}с`);

    return parts.join(' ') || '0с';
  }

  resetStats() {
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
}

module.exports = BotStats; 