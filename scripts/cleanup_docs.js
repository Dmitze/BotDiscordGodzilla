const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Configuration
const DOCS_DIR = path.join(__dirname, '..', 'docs');
const ARCHIVE_DIR = path.join(DOCS_DIR, 'archive');
const DOCUMENTATION_DIR = path.join(DOCS_DIR, 'documentation');

// Files to keep (relative to docs/)
const KEEP_FILES = [
  'archive/README.md',
  'documentation/README.md'
];

// Directories to keep (relative to docs/)
const KEEP_DIRS = [
  'archive',
  'documentation'
];

// Files to delete (patterns)
const DELETE_PATTERNS = [
  // Match any file with these words in the name (case insensitive)
  /FINAL_/i,
  /REPORT/i,
  /CHECKLIST/i,
  /PLAN/i,
  /ANALYSIS/i,
  /MIGRATION/i,
  /OPTIMIZATION/i,
  /TESTING/i,
  /COVERAGE/i,
  /REFACTOR/i,
  /DEPLOYMENT/i,
  /COMMANDS?/i,
  /SERVICES?/i,
  /PROGRESS/i,
  /COMPLETE/i,
  /COMPREHENSIVE/i,
  /EXPANSION/i,
  /TYPESCRIPT/i,
  /STRUCTURE/i,
  /PHASE[0-9]/i
];

// Function to check if a file should be kept
function shouldKeepFile(filePath) {
  const relativePath = path.relative(DOCS_DIR, filePath).replace(/\\/g, '/');
  
  // Check if file is in the keep list
  if (KEEP_FILES.includes(relativePath)) {
    return true;
  }
  
  // Check if file is in a directory that should be kept
  for (const dir of KEEP_DIRS) {
    if (relativePath.startsWith(dir + '/') && !relativePath.endsWith('/README.md')) {
      return false;
    }
  }
  
  return true;
}

// Function to check if a file matches any delete pattern
function shouldDeleteFile(filePath) {
  const fileName = path.basename(filePath);
  
  // Always keep README.md files
  if (fileName.toLowerCase() === 'readme.md') {
    return false;
  }
  
  // Check if file matches any delete pattern
  return DELETE_PATTERNS.some(pattern => {
    if (typeof pattern === 'string') {
      const regex = new RegExp(pattern.replace(/\*/g, '.*'), 'i');
      return regex.test(fileName);
    } else if (pattern instanceof RegExp) {
      return pattern.test(fileName);
    }
    return false;
  });
}

// Function to delete files
function deleteFiles() {
  console.log('🔍 Пошук застарілих документів...\n');
  
  let deletedCount = 0;
  let keptCount = 0;
  
  // Process archive directory - delete everything except README.md
  if (fs.existsSync(ARCHIVE_DIR)) {
    const files = fs.readdirSync(ARCHIVE_DIR);
    
    console.log('📂 Папка archive/:');
    for (const file of files) {
      const filePath = path.join(ARCHIVE_DIR, file);
      const relativePath = path.relative(process.cwd(), filePath);
      
      if (fs.statSync(filePath).isDirectory()) {
        // Delete entire directory recursively
        console.log(`   🗑️  Видалено директорію: ${file}/`);
        fs.rmSync(filePath, { recursive: true, force: true });
        deletedCount++;
        continue;
      }
      
      // Keep only README.md in archive
      if (file.toLowerCase() === 'readme.md') {
        console.log(`   ✅ Залишено: ${file}`);
        keptCount++;
      } else {
        console.log(`   ❌ Видалено: ${file}`);
        fs.unlinkSync(filePath);
        deletedCount++;
      }
    }
  }
  
  // Process documentation directory - delete everything
  if (fs.existsSync(DOCUMENTATION_DIR)) {
    const files = fs.readdirSync(DOCUMENTATION_DIR);
    
    console.log('\n📂 Папка documentation/:');
    for (const file of files) {
      const filePath = path.join(DOCUMENTATION_DIR, file);
      
      if (fs.statSync(filePath).isDirectory()) {
        // Delete entire directory recursively
        console.log(`   🗑️  Видалено директорію: ${file}/`);
        fs.rmSync(filePath, { recursive: true, force: true });
        deletedCount++;
      } else {
        console.log(`   ❌ Видалено: ${file}`);
        fs.unlinkSync(filePath);
        deletedCount++;
      }
    }
  }
  
  // Update README.md to remove references to deleted files
  updateMainReadme();
  
  console.log(`\n📊 Підсумок:`);
  console.log(`- Видалено файлів: ${deletedCount}`);
  console.log(`- Залишено файлів: ${keptCount}`);
  console.log('\n✅ Очищення завершено. Рекомендується запустити `node scripts/check_links.js` для перевірки залишених посилань.');
}

// Function to update main README.md
function updateMainReadme() {
  const readmePath = path.join(DOCS_DIR, 'README.md');
  if (!fs.existsSync(readmePath)) return;
  
  let content = fs.readFileSync(readmePath, 'utf8');
  
  // Remove sections that reference deleted files
  const patternsToRemove = [
    /## 📚 Документація[\s\S]*?(?=## )/,
    /## 📚 Archive[\s\S]*?(?=## )/,
    /## 📊 Звіти[\s\S]*?(?=## )/,
    /## 🚀 Deployment[\s\S]*?(?=## )/
  ];
  
  patternsToRemove.forEach(pattern => {
    content = content.replace(pattern, '');
  });
  
  // Clean up multiple consecutive empty lines
  content = content.replace(/\n{3,}/g, '\n\n');
  
  fs.writeFileSync(readmePath, content, 'utf8');
  console.log('\n📝 Оновлено головний README.md - видалено посилання на видалені файли');
}

// Run the cleanup
if (require.main === module) {
  console.log('🔄 Запуск очищення застарілої документації...\n');
  
  try {
    deleteFiles();
  } catch (error) {
    console.error('❌ Помилка під час очищення документації:', error);
    process.exit(1);
  }
}

module.exports = { deleteFiles };
