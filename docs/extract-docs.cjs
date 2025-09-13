const fs = require('fs');
const path = require('path');

// Function to get all HTML files recursively
function getAllHtmlFiles(dir, fileList = []) {
    const files = fs.readdirSync(dir);
    
    files.forEach(file => {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);
        
        if (stat.isDirectory()) {
            // Skip special directories
            if (!file.startsWith('.') && file !== 'node_modules') {
                getAllHtmlFiles(filePath, fileList);
            }
        } else if (file.endsWith('.html') && file !== 'index.html') {
            fileList.push(filePath);
        }
    });
    
    return fileList;
}

// Function to extract title from HTML file
function extractTitle(filePath) {
    try {
        const content = fs.readFileSync(filePath, 'utf8');
        const titleMatch = content.match(/<title>(.*?)<\/title>/i);
        return titleMatch ? titleMatch[1].replace(' - Godzilla Bot Документація', '') : path.basename(filePath, '.html');
    } catch (error) {
        console.error(`Error reading ${filePath}:`, error.message);
        return path.basename(filePath, '.html');
    }
}

// Function to determine category from file path
function getCategory(filePath) {
    const relativePath = path.relative(path.join(__dirname, 'pages'), filePath);
    const dirName = path.dirname(relativePath);
    
    if (dirName === '.') {
        return 'root';
    }
    
    const parts = dirName.split(path.sep);
    return parts[0] || 'root';
}

// Function to get file stats (size and last modified date)
function getFileStats(filePath) {
    try {
        const stats = fs.statSync(filePath);
        return {
            size: stats.size,
            lastModified: stats.mtime
        };
    } catch (error) {
        console.error(`Error getting stats for ${filePath}:`, error.message);
        return {
            size: 0,
            lastModified: new Date()
        };
    }
}

// Main function
function generateDocsData() {
    const pagesDir = path.join(__dirname, 'pages');
    const htmlFiles = getAllHtmlFiles(pagesDir);
    
    const docsData = htmlFiles.map(filePath => {
        const relativePath = path.relative(__dirname, filePath);
        const title = extractTitle(filePath);
        const category = getCategory(filePath);
        const fileStats = getFileStats(filePath);
        
        return {
            path: relativePath,
            title: title,
            category: category,
            size: fileStats.size,
            lastModified: fileStats.lastModified
        };
    });
    
    // Write to JSON file
    fs.writeFileSync(path.join(__dirname, 'docs-data.json'), JSON.stringify(docsData, null, 2));
    console.log(`Processed ${docsData.length} documentation files`);
}

generateDocsData();