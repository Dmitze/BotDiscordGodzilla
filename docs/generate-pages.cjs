const fs = require('fs');
const path = require('path');

// Configuration
const DOCS_DIR = __dirname;
const OUTPUT_DIR = path.join(__dirname, 'pages');

// Create output directory if it doesn't exist
if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// Category-specific color schemes and fonts
const CATEGORY_STYLES = {
    'api': {
        primary: '#FF6B6B',
        secondary: '#4ECDC4',
        accent: '#45B7D1',
        background: '#1a1a2e',
        surface: '#16213e',
        font: "'Source Code Pro', 'Fira Code', monospace"
    },
    'architecture': {
        primary: '#9B5DE5',
        secondary: '#F15BB5',
        accent: '#00BBF9',
        background: '#0a0a0a',
        surface: '#1a1a1a',
        font: "'Roboto Mono', 'Ubuntu Mono', monospace"
    },
    'guides': {
        primary: '#00F5D4',
        secondary: '#90E0EF',
        accent: '#0077B6',
        background: '#03071e',
        surface: '#1d3557',
        font: "'Lato', 'Open Sans', sans-serif"
    },
    'security': {
        primary: '#FF9E00',
        secondary: '#FF5400',
        accent: '#7209B7',
        background: '#000000',
        surface: '#1a1a1a',
        font: "'Oxygen', 'Montserrat', sans-serif"
    },
    'deployment': {
        primary: '#06D6A0',
        secondary: '#118AB2',
        accent: '#073B4C',
        background: '#001219',
        surface: '#003049',
        font: "'Nunito', 'Poppins', sans-serif"
    },
    'changelog': {
        primary: '#8AC926',
        secondary: '#FFCA3A',
        accent: '#FF595E',
        background: '#1f1f1f',
        surface: '#2d2d2d',
        font: "'Space Mono', 'Roboto Mono', monospace"
    },
    'examples': {
        primary: '#FFB703',
        secondary: '#FB8500',
        accent: '#8ECAE6',
        background: '#023047',
        surface: '#219EBC',
        font: "'Quicksand', 'Rubik', sans-serif"
    },
    'root': {
        primary: '#5865F2',
        secondary: '#a1a1aa',
        accent: '#7289da',
        background: '#1e1e2d',
        surface: '#2b2b40',
        font: "'Inter', 'Segoe UI', sans-serif"
    },
    'default': {
        primary: '#5865F2',
        secondary: '#a1a1aa',
        accent: '#7289da',
        background: '#1e1e2d',
        surface: '#2b2b40',
        font: "'Inter', 'Segoe UI', sans-serif"
    }
};

// HTML template for documentation pages with category-specific styles
const PAGE_TEMPLATE = `<!DOCTYPE html>
<html lang="uk">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{TITLE}} - Godzilla Bot Документація</title>
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/animate.css/4.1.1/animate.min.css">
    <!-- Category-specific Google Fonts -->
    {{CATEGORY_FONTS}}
    <style>
        :root {
            --primary: {{PRIMARY}};
            --secondary: {{SECONDARY}};
            --accent: {{ACCENT}};
            --background: {{BACKGROUND}};
            --surface: {{SURFACE}};
            --text: #e6e6e6;
            --text-secondary: #a1a1aa;
            --success: #43b581;
            --warning: #faa61a;
            --error: #f04747;
        }
        
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: {{FONT_FAMILY}};
            line-height: 1.6;
            color: var(--text);
            background: linear-gradient(135deg, var(--background) 0%, #000000 100%);
            background-attachment: fixed;
            margin: 0;
            padding: 0;
            min-height: 100vh;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            padding: 2rem 1rem;
        }
        
        header {
            text-align: center;
            margin-bottom: 2rem;
            padding: 2rem;
            background: var(--surface);
            border-radius: 12px;
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.3);
            position: relative;
            overflow: hidden;
            border: 1px solid rgba({{PRIMARY_R}}, {{PRIMARY_G}}, {{PRIMARY_B}}, 0.3);
        }
        
        h1 {
            font-size: 2rem;
            margin-bottom: 1rem;
            color: var(--primary);
            text-shadow: 0 2px 4px rgba(0, 0, 0, 0.3);
            font-weight: 700;
        }
        
        .content {
            background: var(--surface);
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 8px 20px rgba(0, 0, 0, 0.3);
            margin-bottom: 2rem;
            border: 1px solid rgba({{PRIMARY_R}}, {{PRIMARY_G}}, {{PRIMARY_B}}, 0.2);
        }
        
        .content h1 {
            font-size: 1.8rem;
            margin-top: 0;
            color: var(--primary);
            padding-bottom: 0.5rem;
            border-bottom: 3px solid var(--accent);
        }
        
        .content h2 {
            font-size: 1.5rem;
            margin-top: 1.5rem;
            color: var(--accent);
            border-bottom: 2px solid var(--primary);
            padding-bottom: 0.4rem;
        }
        
        .content h3 {
            font-size: 1.3rem;
            margin-top: 1.2rem;
            color: var(--text);
            position: relative;
            padding-left: 0.8rem;
        }
        
        .content h3:before {
            content: "▶";
            color: var(--primary);
            position: absolute;
            left: 0;
            top: 0;
            font-size: 0.8rem;
        }
        
        .content h4 {
            font-size: 1.1rem;
            margin-top: 1rem;
            color: var(--text-secondary);
        }
        
        .content p {
            margin-bottom: 1rem;
            line-height: 1.6;
        }
        
        .content ul, .content ol {
            margin-left: 1.5rem;
            margin-bottom: 1rem;
        }
        
        .content li {
            margin-bottom: 0.4rem;
            position: relative;
            padding-left: 1.2rem;
        }
        
        .content ul li:before {
            content: "•";
            color: var(--primary);
            position: absolute;
            left: 0;
            top: 0;
        }
        
        .content code {
            background: rgba(0, 0, 0, 0.3);
            padding: 0.2rem 0.4rem;
            border-radius: 4px;
            font-family: 'Courier New', monospace;
            color: var(--accent);
            font-size: 0.9rem;
        }
        
        .content pre {
            background: rgba(0, 0, 0, 0.3);
            padding: 0.8rem;
            border-radius: 8px;
            overflow-x: auto;
            margin: 0.8rem 0;
            border-left: 3px solid var(--primary);
            font-size: 0.9rem;
            line-height: 1.4;
        }
        
        .content pre code {
            background: none;
            padding: 0;
            color: inherit;
        }
        
        .content a {
            color: var(--primary);
            text-decoration: none;
            transition: all 0.3s ease;
            border-bottom: 1px dotted var(--primary);
        }
        
        .content a:hover {
            color: var(--accent);
            text-decoration: underline;
            border-bottom: 1px solid var(--accent);
        }
        
        .content table {
            width: 100%;
            border-collapse: collapse;
            margin: 0.8rem 0;
            background: rgba(0, 0, 0, 0.2);
            border-radius: 8px;
            overflow: hidden;
            font-size: 0.9rem;
        }
        
        .content table th, .content table td {
            padding: 0.6rem;
            text-align: left;
            border: 1px solid var(--text-secondary);
        }
        
        .content table th {
            background: rgba({{PRIMARY_R}}, {{PRIMARY_G}}, {{PRIMARY_B}}, 0.2);
            color: var(--primary);
            font-weight: 600;
        }
        
        .content blockquote {
            border-left: 3px solid var(--primary);
            padding: 0.5rem 0.8rem;
            margin: 0.8rem 0;
            background: rgba({{PRIMARY_R}}, {{PRIMARY_G}}, {{PRIMARY_B}}, 0.1);
            border-radius: 0 8px 8px 0;
        }
        
        .navigation {
            display: flex;
            justify-content: space-between;
            margin-top: 1.5rem;
        }
        
        .nav-button {
            display: inline-block;
            background: var(--surface);
            color: var(--text);
            padding: 0.6rem 1.2rem;
            border-radius: 8px;
            text-decoration: none;
            transition: all 0.3s ease;
            border: 1px solid var(--primary);
            font-weight: 600;
            font-size: 0.9rem;
        }
        
        .nav-button:hover {
            background: var(--primary);
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba({{PRIMARY_R}}, {{PRIMARY_G}}, {{PRIMARY_B}}, 0.4);
        }
        
        .nav-button.disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        
        .nav-button.disabled:hover {
            background: var(--surface);
            transform: none;
            box-shadow: none;
        }
        
        .back-link {
            display: inline-block;
            margin-bottom: 1.2rem;
            color: var(--primary);
            text-decoration: none;
            font-weight: 600;
            transition: all 0.3s ease;
            padding: 0.4rem 0.8rem;
            border-radius: 6px;
            border: 1px solid var(--primary);
            font-size: 0.9rem;
        }
        
        .back-link:hover {
            color: var(--accent);
            transform: translateX(-3px);
            background: rgba({{PRIMARY_R}}, {{PRIMARY_G}}, {{PRIMARY_B}}, 0.1);
        }
        
        footer {
            text-align: center;
            padding: 1.5rem;
            color: var(--text-secondary);
            font-size: 0.85rem;
            margin-top: 1.5rem;
            border-top: 1px solid var(--surface);
        }
        
        /* Floating elements */
        .floating-elements {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            pointer-events: none;
            z-index: -1;
        }
        
        .floating-element {
            position: absolute;
            border-radius: 50%;
            background: radial-gradient(circle, rgba({{PRIMARY_R}}, {{PRIMARY_G}}, {{PRIMARY_B}}, 0.1) 0%, transparent 70%);
            animation: float 6s ease-in-out infinite;
        }
        
        .floating-element:nth-child(1) {
            width: 200px;
            height: 200px;
            top: 10%;
            left: 5%;
            animation-delay: 0s;
        }
        
        .floating-element:nth-child(2) {
            width: 150px;
            height: 150px;
            bottom: 15%;
            right: 10%;
            animation-delay: 2s;
        }
        
        .floating-element:nth-child(3) {
            width: 100px;
            height: 100px;
            top: 40%;
            right: 20%;
            animation-delay: 4s;
        }
        
        @keyframes float {
            0% { transform: translateY(0px); }
            50% { transform: translateY(-10px); }
            100% { transform: translateY(0px); }
        }
        
        /* Category-specific animations */
        .category-animation {
            animation: {{CATEGORY_ANIMATION}} 2s ease-in-out infinite alternate;
        }
        
        @keyframes pulseGlow {
            0% { box-shadow: 0 0 5px var(--primary); }
            100% { box-shadow: 0 0 15px var(--primary), 0 0 20px var(--accent); }
        }
        
        @keyframes borderGlow {
            0% { border-color: var(--primary); }
            100% { border-color: var(--accent); }
        }
        
        @keyframes textGlow {
            0% { text-shadow: 0 0 5px var(--primary); }
            100% { text-shadow: 0 0 8px var(--primary), 0 0 15px var(--accent); }
        }
        
        /* Category-specific header styles */
        header {
            {{CATEGORY_HEADER_STYLE}}
        }
        
        @media (max-width: 768px) {
            .container {
                padding: 1rem;
            }
            
            h1 {
                font-size: 1.6rem;
            }
            
            .content {
                padding: 1rem;
            }
            
            .content h1 {
                font-size: 1.5rem;
            }
            
            .content h2 {
                font-size: 1.3rem;
            }
            
            .content h3 {
                font-size: 1.1rem;
            }
            
            .navigation {
                flex-direction: column;
                gap: 1rem;
            }
            
            .nav-button {
                text-align: center;
            }
            
            .floating-element:nth-child(1) {
                width: 150px;
                height: 150px;
            }
            
            .floating-element:nth-child(2) {
                width: 100px;
                height: 100px;
            }
            
            .floating-element:nth-child(3) {
                width: 70px;
                height: 70px;
            }
        }
        
        /* Bookmark styles */
        .bookmark-btn {
            position: absolute;
            top: 1rem;
            right: 1rem;
            background: none;
            border: none;
            color: var(--text-secondary);
            cursor: pointer;
            font-size: 1.2rem;
            transition: all 0.3s ease;
            width: 30px;
            height: 30px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 50%;
        }
        
        .bookmark-btn:hover {
            color: var(--warning);
            background: rgba(250, 166, 26, 0.1);
        }
        
        .bookmark-btn.bookmarked {
            color: var(--warning);
        }
        
        /* Rating styles */
        .rating-container {
            display: flex;
            align-items: center;
            gap: 0.5rem;
            margin-top: 2rem;
            padding-top: 1rem;
            border-top: 1px solid var(--text-secondary);
        }
        
        .rating-container p {
            margin: 0;
            color: var(--text);
        }
        
        .rating-stars {
            display: flex;
            gap: 0.2rem;
        }
        
        .rating-star {
            color: var(--text-secondary);
            cursor: pointer;
            transition: color 0.2s ease;
            font-size: 1.2rem;
        }
        
        .rating-star.active {
            color: var(--warning);
        }
        
        .rating-value {
            font-size: 0.9rem;
            color: var(--text-secondary);
        }
        
        /* Floating TOC */
        .floating-toc {
            position: fixed;
            right: 20px;
            top: 50%;
            transform: translateY(-50%);
            background: var(--surface);
            border: 1px solid var(--primary);
            border-radius: 8px;
            padding: 1rem;
            max-height: 70vh;
            overflow-y: auto;
            z-index: 100;
            box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2);
            max-width: 250px;
            display: none;
        }
        
        .floating-toc.visible {
            display: block;
        }
        
        .toc-title {
            color: var(--primary);
            margin-bottom: 1rem;
            font-size: 1.1rem;
            text-align: center;
        }
        
        .toc-list {
            list-style: none;
            padding: 0;
            margin: 0;
        }
        
        .toc-item {
            margin-bottom: 0.5rem;
        }
        
        .toc-link {
            color: var(--text-secondary);
            text-decoration: none;
            font-size: 0.9rem;
            transition: color 0.2s ease;
            display: block;
            padding: 0.3rem 0;
        }
        
        .toc-link:hover {
            color: var(--primary);
        }
        
        .toc-link.indent-1 {
            padding-left: 1rem;
        }
        
        .toc-link.indent-2 {
            padding-left: 2rem;
        }
        
        .toc-link.indent-3 {
            padding-left: 3rem;
        }
        
        /* Toggle TOC button */
        .toc-toggle {
            position: fixed;
            right: 20px;
            bottom: 20px;
            background: var(--primary);
            color: white;
            width: 50px;
            height: 50px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.2);
            z-index: 99;
        }
    </style>
</head>
<body>
    <div class="floating-elements">
        <div class="floating-element"></div>
        <div class="floating-element"></div>
        <div class="floating-element"></div>
    </div>

    <div class="container">
        <header class="category-animation">
            <h1><i class="{{CATEGORY_ICON}}"></i> {{TITLE}}</h1>
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

// Category-specific configurations
const CATEGORY_CONFIGS = {
    'api': {
        icon: 'fas fa-plug',
        animation: 'pulseGlow',
        headerStyle: 'background: linear-gradient(135deg, var(--surface) 0%, var(--background) 100%); border: 2px solid var(--primary);'
    },
    'architecture': {
        icon: 'fas fa-project-diagram',
        animation: 'borderGlow',
        headerStyle: 'background: linear-gradient(90deg, var(--surface) 0%, var(--background) 100%); border-left: 5px solid var(--primary);'
    },
    'guides': {
        icon: 'fas fa-book',
        animation: 'textGlow',
        headerStyle: 'background: linear-gradient(45deg, var(--surface) 0%, var(--background) 100%); border-top: 3px solid var(--primary); border-bottom: 3px solid var(--accent);'
    },
    'security': {
        icon: 'fas fa-shield-alt',
        animation: 'pulseGlow',
        headerStyle: 'background: linear-gradient(180deg, var(--surface) 0%, var(--background) 100%); border: 2px dashed var(--primary);'
    },
    'deployment': {
        icon: 'fas fa-server',
        animation: 'borderGlow',
        headerStyle: 'background: linear-gradient(225deg, var(--surface) 0%, var(--background) 100%); border-radius: 20px;'
    },
    'changelog': {
        icon: 'fas fa-history',
        animation: 'textGlow',
        headerStyle: 'background: linear-gradient(90deg, var(--surface) 0%, var(--background) 100%); border: 1px solid var(--primary); box-shadow: 0 0 15px var(--primary);'
    },
    'examples': {
        icon: 'fas fa-lightbulb',
        animation: 'pulseGlow',
        headerStyle: 'background: linear-gradient(135deg, var(--surface) 0%, var(--background) 100%); border: 3px double var(--primary);'
    },
    'root': {
        icon: 'fas fa-home',
        animation: 'borderGlow',
        headerStyle: 'background: linear-gradient(135deg, var(--surface) 0%, var(--background) 100%); border: 1px solid var(--primary);'
    },
    'default': {
        icon: 'fas fa-book',
        animation: 'textGlow',
        headerStyle: 'background: linear-gradient(135deg, var(--surface) 0%, var(--background) 100%); border: 1px solid var(--primary);'
    }
};

// Category-specific Google Fonts
const CATEGORY_FONTS = {
    'api': '<link href="https://fonts.googleapis.com/css2?family=Source+Code+Pro:wght@400;600;700&family=Fira+Code:wght@400;500;600&display=swap" rel="stylesheet">',
    'architecture': '<link href="https://fonts.googleapis.com/css2?family=Roboto+Mono:wght@400;500;600&family=Ubuntu+Mono:wght@400;700&display=swap" rel="stylesheet">',
    'guides': '<link href="https://fonts.googleapis.com/css2?family=Lato:wght@400;700&family=Open+Sans:wght@400;600;700&display=swap" rel="stylesheet">',
    'security': '<link href="https://fonts.googleapis.com/css2?family=Oxygen:wght@400;700&family=Montserrat:wght@400;600;700&display=swap" rel="stylesheet">',
    'deployment': '<link href="https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700&family=Poppins:wght@400;500;600&display=swap" rel="stylesheet">',
    'changelog': '<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Roboto+Mono:wght@400;500;600&display=swap" rel="stylesheet">',
    'examples': '<link href="https://fonts.googleapis.com/css2?family=Quicksand:wght@400;500;600&family=Rubik:wght@400;500;600&display=swap" rel="stylesheet">',
    'root': '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Segoe+UI:wght@400;500;600;700&display=swap" rel="stylesheet">',
    'default': '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Segoe+UI:wght@400;500;600;700&display=swap" rel="stylesheet">'
};

// Function to convert hex color to RGB
function hexToRgb(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? `${parseInt(result[1], 16)}, ${parseInt(result[2], 16)}, ${parseInt(result[3], 16)}` : '88, 101, 242';
}

// Function to convert hex color to RGB object for CSS use
function hexToRgbValues(hex) {
    const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
    return result ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16)
    } : { r: 88, g: 101, b: 242 };
}

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

// Function to get category from file path
function getCategory(filePath) {
    const relativePath = path.relative(DOCS_DIR, filePath);
    const dirName = path.dirname(relativePath);
    
    // If in root docs directory
    if (dirName === '.') {
        return 'root';
    }
    
    // Get the first directory name as category
    const parts = dirName.split(path.sep);
    return parts[0] || 'root';
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
        
        // Get category for styling
        const category = getCategory(filePath);
        const styleConfig = CATEGORY_STYLES[category] || CATEGORY_STYLES['default'];
        const categoryConfig = CATEGORY_CONFIGS[category] || CATEGORY_CONFIGS['default'];
        const fontLinks = CATEGORY_FONTS[category] || CATEGORY_FONTS['default'];
        
        // Convert hex colors to RGB for animations
        const primaryRgb = hexToRgb(styleConfig.primary);
        const rgbValues = hexToRgbValues(styleConfig.primary);
        
        // Generate navigation buttons
        const { prevButton, nextButton } = generateNavigationButtons(filePath, allFiles);
        
        // Generate page HTML with category-specific styles
        let pageHtml = PAGE_TEMPLATE
            .replace(/{{TITLE}}/g, title)
            .replace('{{CONTENT}}', htmlContent)
            .replace('{{PREV_BUTTON}}', prevButton)
            .replace('{{NEXT_BUTTON}}', nextButton)
            .replace('{{PRIMARY}}', styleConfig.primary)
            .replace('{{SECONDARY}}', styleConfig.secondary)
            .replace('{{ACCENT}}', styleConfig.accent)
            .replace('{{BACKGROUND}}', styleConfig.background)
            .replace('{{SURFACE}}', styleConfig.surface)
            .replace('{{FONT_FAMILY}}', styleConfig.font)
            .replace('{{PRIMARY_RGB}}', primaryRgb)
            .replace('{{CATEGORY_FONTS}}', fontLinks)
            .replace('{{CATEGORY_ICON}}', categoryConfig.icon)
            .replace('{{CATEGORY_ANIMATION}}', categoryConfig.animation)
            .replace('{{CATEGORY_HEADER_STYLE}}', categoryConfig.headerStyle)
            .replace(/{{PRIMARY_R}}/g, rgbValues.r.toString())
            .replace(/{{PRIMARY_G}}/g, rgbValues.g.toString())
            .replace(/{{PRIMARY_B}}/g, rgbValues.b.toString());
        
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

// Run the generation
generateAllPages();