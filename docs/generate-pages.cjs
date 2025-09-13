const fs = require('fs');
const path = require('path');

// Configuration
const DOCS_DIR = __dirname;
const OUTPUT_DIR = path.join(__dirname, 'pages');

// Create output directory if it doesn't exist
if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// HTML template for documentation pages
const PAGE_TEMPLATE = `<!DOCTYPE html>
<html lang="uk">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{TITLE}} - Godzilla Bot Документація</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/animate.css/4.1.1/animate.min.css">
    <link rel="stylesheet" href="../assets/styles.css">
</head>
<body>
    <div class="floating-elements">
        <div class="floating-element"></div>
        <div class="floating-element"></div>
        <div class="floating-element"></div>
    </div>

    <div class="container">
        <header>
            <h1><i class="fas fa-book"></i> {{TITLE}}</h1>
            <p>Документація Discord AI Assistant Bot - Godzilla</p>
        </header>
        
        <a href="../index.html" class="back-link">
            <i class="fas fa-arrow-left"></i> Повернутися до головної
        </a>
        
        <div class="content">
            {{CONTENT}}
        </div>
        
        <div class="navigation">
            {{PREV_BUTTON}}
            {{NEXT_BUTTON}}
        </div>
        
        <footer>
            <p>© 2025 Godzilla Bot | Створено Дмитром Шивачовим (Dmitze)</p>
            <p>Ліцензія MIT | Версія: v3.0.0</p>
        </footer>
    </div>
    
    <script src="../assets/scripts.js"></script>
</body>
</html>`;

// Function to convert markdown to HTML (simplified)
function markdownToHtml(markdown) {
    // Convert headers
    let html = markdown
        .replace(/^# (.*$)/gm, '<h1>$1</h1>')
        .replace(/^## (.*$)/gm, '<h2>$1</h2>')
        .replace(/^### (.*$)/gm, '<h3>$1</h3>')
        .replace(/^#### (.*$)/gm, '<h4>$1</h4>');
    
    // Convert bold and italic
    html = html
        .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
        .replace(/\*(.*?)\*/g, '<em>$1</em>');
    
    // Convert code blocks
    html = html
        .replace(/```([\s\S]*?)```/g, '<pre><code>$1</code></pre>')
        .replace(/`(.*?)`/g, '<code>$1</code>');
    
    // Convert links
    html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
    
    // Convert lists
    html = html
        .replace(/^\s*-\s+(.*$)/gm, '<li>$1</li>')
        .replace(/(<li>.*<\/li>)+/gs, '<ul>$&</ul>');
    
    // Convert paragraphs
    html = html
        .replace(/^\s*$(.*?)^\s*$/gms, '<p>$1</p>')
        .replace(/<p>\s*<\/p>/g, '');
    
    return html;
}

// Function to get all markdown files recursively
function getAllMarkdownFiles(dir, fileList = []) {
    const files = fs.readdirSync(dir);
    
    files.forEach(file => {
        const filePath = path.join(dir, file);
        const stat = fs.statSync(filePath);
        
        if (stat.isDirectory()) {
            // Skip node_modules and other special directories
            if (!file.startsWith('.') && file !== 'node_modules' && file !== 'pages') {
                getAllMarkdownFiles(filePath, fileList);
            }
        } else if (file.endsWith('.md')) {
            fileList.push(filePath);
        }
    });
    
    return fileList;
}

// Function to generate navigation buttons
function generateNavigationButtons(filePath, allFiles) {
    const currentIndex = allFiles.indexOf(filePath);
    const prevFile = currentIndex > 0 ? allFiles[currentIndex - 1] : null;
    const nextFile = currentIndex < allFiles.length - 1 ? allFiles[currentIndex + 1] : null;
    
    let prevButton = '<a href="#" class="nav-button disabled" style="opacity: 0.5; cursor: not-allowed;"><i class="fas fa-arrow-left"></i> Попередня</a>';
    let nextButton = '<a href="#" class="nav-button disabled" style="opacity: 0.5; cursor: not-allowed;">Наступна <i class="fas fa-arrow-right"></i></a>';
    
    if (prevFile) {
        const prevRelativePath = path.relative(DOCS_DIR, prevFile);
        const prevHtmlPath = path.join('pages', path.dirname(prevRelativePath), path.basename(prevFile, '.md') + '.html');
        const prevTitleMatch = fs.readFileSync(prevFile, 'utf8').match(/^# (.*)$/m);
        const prevTitle = prevTitleMatch ? prevTitleMatch[1] : path.basename(prevFile, '.md');
        prevButton = `<a href="${prevHtmlPath}" class="nav-button"><i class="fas fa-arrow-left"></i> ${prevTitle}</a>`;
    }
    
    if (nextFile) {
        const nextRelativePath = path.relative(DOCS_DIR, nextFile);
        const nextHtmlPath = path.join('pages', path.dirname(nextRelativePath), path.basename(nextFile, '.md') + '.html');
        const nextTitleMatch = fs.readFileSync(nextFile, 'utf8').match(/^# (.*)$/m);
        const nextTitle = nextTitleMatch ? nextTitleMatch[1] : path.basename(nextFile, '.md');
        nextButton = `<a href="${nextHtmlPath}" class="nav-button">${nextTitle} <i class="fas fa-arrow-right"></i></a>`;
    }
    
    return { prevButton, nextButton };
}

// Function to generate HTML page for a markdown file
function generatePage(filePath, allFiles) {
    try {
        // Read markdown content
        const content = fs.readFileSync(filePath, 'utf8');
        
        // Extract title (first line with #)
        const titleMatch = content.match(/^# (.*)$/m);
        const title = titleMatch ? titleMatch[1] : path.basename(filePath, '.md');
        
        // Convert markdown to HTML
        const htmlContent = markdownToHtml(content);
        
        // Generate navigation buttons
        const { prevButton, nextButton } = generateNavigationButtons(filePath, allFiles);
        
        // Generate page HTML
        const pageHtml = PAGE_TEMPLATE
            .replace(/{{TITLE}}/g, title)
            .replace('{{CONTENT}}', htmlContent)
            .replace('{{PREV_BUTTON}}', prevButton)
            .replace('{{NEXT_BUTTON}}', nextButton);
        
        // Create output directory structure
        const relativePath = path.relative(DOCS_DIR, filePath);
        const outputDir = path.join(OUTPUT_DIR, path.dirname(relativePath));
        
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
        }
        
        // Write HTML file
        const outputPath = path.join(outputDir, path.basename(filePath, '.md') + '.html');
        fs.writeFileSync(outputPath, pageHtml);
        
        console.log(`Generated: ${outputPath}`);
    } catch (error) {
        console.error(`Error processing ${filePath}:`, error.message);
    }
}

// Main function
function generateAllPages() {
    console.log('Generating documentation pages...');
    
    // Get all markdown files
    const markdownFiles = getAllMarkdownFiles(DOCS_DIR);
    
    console.log(`Found ${markdownFiles.length} markdown files`);
    
    // Generate page for each markdown file
    markdownFiles.forEach(filePath => {
        generatePage(filePath, markdownFiles);
    });
    
    console.log('Documentation pages generation completed!');
}

// Run the generator
generateAllPages();