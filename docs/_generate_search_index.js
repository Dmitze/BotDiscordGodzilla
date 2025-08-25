const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const { NodeHtmlMarkdown } = require('node-html-markdown');

// Configuration
const DOCS_DIR = __dirname;
const OUTPUT_FILE = path.join(DOCS_DIR, 'search-index.json');
const EXCLUDE_DIRS = ['node_modules', '.git', 'assets', 'img'];
const INCLUDE_EXTENSIONS = ['.html', '.md'];

// Track processed files
const searchIndex = [];

/**
 * Extract text content from HTML or Markdown file
 */
async function processFile(filePath) {
    const relativePath = path.relative(DOCS_DIR, filePath);
    const content = fs.readFileSync(filePath, 'utf8');
    let text = '';
    let title = '';

    try {
        if (filePath.endsWith('.html')) {
            const dom = new JSDOM(content);
            const doc = dom.window.document;
            
            // Extract title
            title = doc.querySelector('h1')?.textContent || 
                    doc.title || 
                    path.basename(filePath, '.html');
            
            // Remove unwanted elements
            const unwantedElements = doc.querySelectorAll('nav, header, footer, script, style, .nav, .header, .footer');
            unwantedElements.forEach(el => el.remove());
            
            // Convert HTML to markdown for cleaner text
            text = NodeHtmlMarkdown.translate(doc.body.innerHTML)
                .replace(/\s+/g, ' ')
                .trim();
        } else if (filePath.endsWith('.md')) {
            // For markdown files, use the first # heading as title
            const lines = content.split('\n');
            const titleLine = lines.find(line => line.startsWith('# '));
            title = titleLine ? titleLine.replace(/^#+\s*/, '') : path.basename(filePath, '.md');
            
            // Remove code blocks and other markdown syntax
            text = content
                .replace(/```[\s\S]*?```/g, '')  // Code blocks
                .replace(/`[^`]+`/g, '')          // Inline code
                .replace(/[#*_\-|>~=]/g, '')      // Markdown syntax
                .replace(/\[([^\]]+)\]\([^)]+\)/g, '$1')  // Links
                .replace(/\s+/g, ' ')
                .trim();
        }
        
        if (title && text) {
            searchIndex.push({
                title: title,
                url: relativePath.replace(/\\/g, '/'), // Ensure forward slashes for URLs
                content: text
            });
        }
    } catch (error) {
        console.error(`Error processing ${filePath}:`, error.message);
    }
}

/**
 * Recursively process all files in a directory
 */
function processDirectory(directory) {
    const files = fs.readdirSync(directory);
    
    files.forEach(file => {
        const fullPath = path.join(directory, file);
        const stat = fs.statSync(fullPath);
        
        // Skip excluded directories
        if (stat.isDirectory()) {
            if (!EXCLUDE_DIRS.includes(file) && !file.startsWith('.')) {
                processDirectory(fullPath);
            }
            return;
        }
        
        // Process supported file types
        const ext = path.extname(file).toLowerCase();
        if (INCLUDE_EXTENSIONS.includes(ext)) {
            processFile(fullPath);
        }
    });
}

// Main function
async function generateSearchIndex() {
    console.log('Generating search index...');
    
    // Process all documentation files
    processDirectory(DOCS_DIR);
    
    // Save the search index
    fs.writeFileSync(OUTPUT_FILE, JSON.stringify(searchIndex, null, 2), 'utf8');
    
    console.log(`Search index generated with ${searchIndex.length} entries`);
    console.log(`Output file: ${OUTPUT_FILE}`);
}

// Run the generator
generateSearchIndex().catch(console.error);
