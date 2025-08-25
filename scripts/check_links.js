const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// Configuration
const DOCS_DIR = path.join(__dirname, '../docs');
const IGNORE_DIRS = ['node_modules', '.git', 'build', 'dist'];
const EXTERNAL_URLS = new Set([
  'https://github.com/Dmitze',
  'https://t.me/Dmitry_Shiva',
  'https://github.com/Dmitze/BotDiscordGodzilla',
  'https://github.com/Dmitze/BotDiscordGodzilla/issues',
  'https://github.com/Dmitze/BotDiscordGodzilla/discussions'
]);

// Stats
const stats = {
  totalFiles: 0,
  checkedLinks: 0,
  brokenLinks: 0,
  brokenFiles: new Set(),
  externalLinks: 0,
  externalBroken: 0
};

// Get all markdown files
function getMarkdownFiles(dir, fileList = []) {
  const files = fs.readdirSync(dir);
  
  files.forEach(file => {
    const filePath = path.join(dir, file);
    const stat = fs.statSync(filePath);
    
    if (stat.isDirectory()) {
      if (!IGNORE_DIRS.includes(file)) {
        getMarkdownFiles(filePath, fileList);
      }
    } else if (file.endsWith('.md')) {
      fileList.push(filePath);
      stats.totalFiles++;
    }
  });
  
  return fileList;
}

// Check if file exists
function fileExists(filePath) {
  try {
    // Handle relative paths
    const fullPath = path.isAbsolute(filePath) 
      ? filePath 
      : path.join(DOCS_DIR, filePath);
    
    // Check if file exists
    if (fs.existsSync(fullPath)) {
      return true;
    }
    
    // Check if it's a directory with README.md
    if (fs.existsSync(path.join(fullPath, 'README.md'))) {
      return true;
    }
    
    return false;
  } catch (err) {
    return false;
  }
}

// Check if URL is accessible
async function checkUrl(url) {
  if (EXTERNAL_URLS.has(url)) {
    stats.externalLinks++;
    return true; // Skip checking known external URLs
  }
  
  if (url.startsWith('http')) {
    stats.externalLinks++;
    try {
      // Use curl to check URL
      execSync(`curl -I -s -o /dev/null -w "%{http_code}" "${url}"`, { stdio: 'pipe' });
      return true;
    } catch (error) {
      stats.externalBroken++;
      return false;
    }
  }
  
  return true; // Skip other non-file URLs
}

// Process a single markdown file
async function processFile(filePath) {
  const content = fs.readFileSync(filePath, 'utf8');
  const linkRegex = /\[([^\]]+)\]\(([^)]+)\)/g;
  let match;
  const brokenInFile = [];
  
  while ((match = linkRegex.exec(content)) !== null) {
    const [fullMatch, text, link] = match;
    
    // Skip anchor links
    if (link.startsWith('#')) continue;
    
    // Skip mailto and other protocols
    if (link.includes('://') || link.startsWith('mailto:')) {
      if (!(await checkUrl(link))) {
        brokenInFile.push({
          text,
          link,
          line: content.substring(0, match.index).split('\n').length
        });
      }
      continue;
    }
    
    // Handle local file links
    const cleanLink = link.split('#')[0]; // Remove anchor
    if (cleanLink) {
      stats.checkedLinks++;
      
      // Handle relative paths
      const targetPath = path.isAbsolute(cleanLink) 
        ? cleanLink 
        : path.join(path.dirname(filePath), cleanLink);
      
      if (!fileExists(targetPath)) {
        brokenInFile.push({
          text,
          link,
          line: content.substring(0, match.index).split('\n').length
        });
      }
    }
  }
  
  if (brokenInFile.length > 0) {
    stats.brokenLinks += brokenInFile.length;
    stats.brokenFiles.add(filePath);
    
    console.log(`\n🔴 ${path.relative(DOCS_DIR, filePath)}`);
    brokenInFile.forEach(({ text, link, line }) => {
      console.log(`   Line ${line}: [${text}](${link})`);
    });
  }
}

// Main function
async function main() {
  console.log('🔍 Перевірка посилань у документації...\n');
  
  const files = getMarkdownFiles(DOCS_DIR);
  
  for (const file of files) {
    await processFile(file);
  }
  
  // Print summary
  console.log('\n📊 Результати перевірки:');
  console.log(`- Перевірено файлів: ${stats.totalFiles}`);
  console.log(`- Перевірено посилань: ${stats.checkedLinks + stats.externalLinks}`);
  console.log(`  - Зовнішніх: ${stats.externalLinks} (${stats.externalBroken} зламаних)`);
  console.log(`  - Внутрішніх: ${stats.checkedLinks}`);
  console.log(`- Знайдено зламаних посилань: ${stats.brokenLinks + stats.externalBroken}`);
  console.log(`- Файлів з помилками: ${stats.brokenFiles.size}`);
  
  if (stats.brokenLinks + stats.externalBroken > 0) {
    process.exit(1);
  }
}

main().catch(console.error);
