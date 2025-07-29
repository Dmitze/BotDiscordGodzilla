/**
 * 🔧 Скрипт для виправлення помилок лінтера
 * Версія: 2.3.0
 */

const fs = require('fs-extra');
const path = require('path');
const { execSync } = require('child_process');

// Кольори для консолі
const colors = {
  reset: '\x1b[0m',
  bright: '\x1b[1m',
  red: '\x1b[31m',
  green: '\x1b[32m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  magenta: '\x1b[35m',
  cyan: '\x1b[36m',
};

// Функції логування
function log(message, color = 'reset') {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

function logInfo(message) {
  log(`[INFO] ${message}`, 'blue');
}

function logSuccess(message) {
  log(`[SUCCESS] ${message}`, 'green');
}

function logWarning(message) {
  log(`[WARNING] ${message}`, 'yellow');
}

function logError(message) {
  log(`[ERROR] ${message}`, 'red');
}

// Функція виправлення line endings
function fixLineEndings(filePath) {
  try {
    let content = fs.readFileSync(filePath, 'utf8');

    // Заміна CRLF на LF
    const originalContent = content;
    content = content.replace(/\r\n/g, '\n');

    if (content !== originalContent) {
      fs.writeFileSync(filePath, content, 'utf8');
      return true;
    }

    return false;
  } catch (error) {
    logError(`Помилка виправлення line endings для ${filePath}: ${error.message}`);
    return false;
  }
}

// Функція виправлення trailing commas
function fixTrailingCommas(filePath) {
  try {
    let content = fs.readFileSync(filePath, 'utf8');
    let modified = false;

    // Видалення trailing commas в об'єктах
    content = content.replace(/,(\s*[}\]])/g, '$1');

    // Видалення trailing commas в масивах
    content = content.replace(/,(\s*])/g, '$1');

    if (content !== fs.readFileSync(filePath, 'utf8')) {
      fs.writeFileSync(filePath, content, 'utf8');
      modified = true;
    }

    return modified;
  } catch (error) {
    logError(`Помилка виправлення trailing commas для ${filePath}: ${error.message}`);
    return false;
  }
}

// Функція виправлення відступів
function fixIndentation(filePath) {
  try {
    let content = fs.readFileSync(filePath, 'utf8');
    const lines = content.split('\n');
    const fixedLines = [];

    for (const line of lines) {
      // Заміна табуляції на пробіли
      let fixedLine = line.replace(/\t/g, '  ');

      // Видалення зайвих пробілів в кінці рядка
      fixedLine = fixedLine.replace(/\s+$/, '');

      fixedLines.push(fixedLine);
    }

    const fixedContent = fixedLines.join('\n');

    if (fixedContent !== content) {
      fs.writeFileSync(filePath, fixedContent, 'utf8');
      return true;
    }

    return false;
  } catch (error) {
    logError(`Помилка виправлення відступів для ${filePath}: ${error.message}`);
    return false;
  }
}

// Функція виправлення кодування
function fixEncoding(filePath) {
  try {
    const content = fs.readFileSync(filePath, 'utf8');

    // Перевірка на наявність BOM
    if (content.charCodeAt(0) === 0xfeff) {
      const contentWithoutBOM = content.slice(1);
      fs.writeFileSync(filePath, contentWithoutBOM, 'utf8');
      return true;
    }

    return false;
  } catch (error) {
    logError(`Помилка виправлення кодування для ${filePath}: ${error.message}`);
    return false;
  }
}

// Функція виправлення конкретних помилок
function fixSpecificErrors(filePath) {
  try {
    let content = fs.readFileSync(filePath, 'utf8');
    let modified = false;

    // Виправлення поширених помилок
    const fixes = [
      // Видалення зайвих пробілів
      { pattern: /\s+$/gm, replacement: '' },

      // Видалення порожніх рядків в кінці файлу
      { pattern: /\n+$/, replacement: '\n' },

      // Виправлення подвійних пробілів
      { pattern: /[ ]{2,}/g, replacement: ' ' },

      // Виправлення подвійних рядків
      { pattern: /\n{3,}/g, replacement: '\n\n' },
    ];

    for (const fix of fixes) {
      const newContent = content.replace(fix.pattern, fix.replacement);
      if (newContent !== content) {
        content = newContent;
        modified = true;
      }
    }

    if (modified) {
      fs.writeFileSync(filePath, content, 'utf8');
    }

    return modified;
  } catch (error) {
    logError(`Помилка виправлення специфічних помилок для ${filePath}: ${error.message}`);
    return false;
  }
}

// Функція перевірки та виправлення файлу
function fixFile(filePath) {
  const stats = {
    lineEndings: false,
    trailingCommas: false,
    indentation: false,
    encoding: false,
    specificErrors: false,
  };

  try {
    logInfo(`Виправлення файлу: ${filePath}`);

    stats.lineEndings = fixLineEndings(filePath);
    stats.trailingCommas = fixTrailingCommas(filePath);
    stats.indentation = fixIndentation(filePath);
    stats.encoding = fixEncoding(filePath);
    stats.specificErrors = fixSpecificErrors(filePath);

    const hasChanges = Object.values(stats).some(Boolean);

    if (hasChanges) {
      logSuccess(`Файл виправлено: ${filePath}`);
      logInfo(`  - Line endings: ${stats.lineEndings ? 'виправлено' : 'OK'}`);
      logInfo(`  - Trailing commas: ${stats.trailingCommas ? 'виправлено' : 'OK'}`);
      logInfo(`  - Indentation: ${stats.indentation ? 'виправлено' : 'OK'}`);
      logInfo(`  - Encoding: ${stats.encoding ? 'виправлено' : 'OK'}`);
      logInfo(`  - Specific errors: ${stats.specificErrors ? 'виправлено' : 'OK'}`);
    } else {
      logInfo(`Файл не потребує виправлень: ${filePath}`);
    }

    return hasChanges;
  } catch (error) {
    logError(`Помилка обробки файлу ${filePath}: ${error.message}`);
    return false;
  }
}

// Функція пошуку файлів для виправлення
function findFilesToFix() {
  const extensions = ['.js', '.json', '.md', '.yml', '.yaml'];
  const excludeDirs = ['node_modules', '.git', 'dist', 'build', 'coverage'];
  const files = [];

  function scanDirectory(dir) {
    try {
      const items = fs.readdirSync(dir);

      for (const item of items) {
        const fullPath = path.join(dir, item);
        const stat = fs.statSync(fullPath);

        if (stat.isDirectory()) {
          if (!excludeDirs.includes(item)) {
            scanDirectory(fullPath);
          }
        } else if (stat.isFile()) {
          const ext = path.extname(item).toLowerCase();
          if (extensions.includes(ext)) {
            files.push(fullPath);
          }
        }
      }
    } catch (error) {
      logWarning(`Помилка сканування директорії ${dir}: ${error.message}`);
    }
  }

  scanDirectory('.');
  return files;
}

// Функція запуску ESLint
function runESLint() {
  try {
    logInfo('Запуск ESLint для перевірки...');
    execSync('npx eslint . --ext .js,.json --fix', { stdio: 'inherit' });
    logSuccess('ESLint виконано успішно');
    return true;
  } catch (error) {
    logWarning('ESLint виявив помилки, які не можуть бути виправлені автоматично');
    return false;
  }
}

// Функція запуску Prettier
function runPrettier() {
  try {
    logInfo('Запуск Prettier для форматування...');
    execSync('npx prettier --write "**/*.{js,json,md,yml,yaml}"', { stdio: 'inherit' });
    logSuccess('Prettier виконано успішно');
    return true;
  } catch (error) {
    logWarning('Prettier виявив помилки');
    return false;
  }
}

// Головна функція
function main() {
  console.log(
    colors.cyan + '🔧 Виправлення помилок лінтера Discord AI Assistant Bot v2.3.0' + colors.reset
  );
  console.log(colors.cyan + '='.repeat(70) + colors.reset);

  let totalFiles = 0;
  let fixedFiles = 0;

  try {
    // Пошук файлів для виправлення
    logInfo('Пошук файлів для виправлення...');
    const files = findFilesToFix();
    totalFiles = files.length;

    logInfo(`Знайдено ${totalFiles} файлів для перевірки`);

    // Виправлення файлів
    for (const file of files) {
      if (fixFile(file)) {
        fixedFiles++;
      }
    }

    // Запуск ESLint
    runESLint();

    // Запуск Prettier
    runPrettier();

    // Підсумок
    console.log('\n' + '='.repeat(70));
    log('РЕЗУЛЬТАТИ ВИПРАВЛЕННЯ:', 'bright');
    console.log('='.repeat(70));

    log(`📁 Всього файлів: ${totalFiles}`, 'blue');
    log(`🔧 Виправлено файлів: ${fixedFiles}`, 'green');
    log(
      `✅ Успішність: ${totalFiles > 0 ? Math.round((fixedFiles / totalFiles) * 100) : 0}%`,
      'cyan'
    );

    if (fixedFiles > 0) {
      logSuccess('🎉 Виправлення завершено успішно!');
    } else {
      logInfo('ℹ️ Всі файли вже відповідають стандартам');
    }

    console.log('='.repeat(70));
  } catch (error) {
    logError(`Критична помилка: ${error.message}`);
    process.exit(1);
  }
}

// Запуск головної функції
if (require.main === module) {
  main();
}

module.exports = {
  fixLineEndings,
  fixTrailingCommas,
  fixIndentation,
  fixEncoding,
  fixSpecificErrors,
  fixFile,
  findFilesToFix,
  runESLint,
  runPrettier,
};
