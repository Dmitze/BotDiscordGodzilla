// Floating TOC functionality for documentation pages
document.addEventListener('DOMContentLoaded', function() {
    // Create floating TOC toggle button
    const tocToggle = document.createElement('div');
    tocToggle.className = 'toc-toggle';
    tocToggle.id = 'tocToggle';
    tocToggle.title = 'Переключити зміст';
    tocToggle.innerHTML = '<i class="fas fa-list"></i>';
    document.body.appendChild(tocToggle);
    
    // Create floating TOC container
    const floatingToc = document.createElement('div');
    floatingToc.className = 'floating-toc';
    floatingToc.id = 'floatingToc';
    floatingToc.innerHTML = `
        <h3 class="toc-title">Зміст документа</h3>
        <ul class="toc-list" id="tocList"></ul>
    `;
    document.body.appendChild(floatingToc);
    
    // Generate TOC from headings
    generateTOC();
    
    // Toggle TOC visibility
    tocToggle.addEventListener('click', function() {
        floatingToc.classList.toggle('visible');
    });
    
    // Close TOC when clicking outside
    document.addEventListener('click', function(e) {
        if (!floatingToc.contains(e.target) && e.target !== tocToggle) {
            floatingToc.classList.remove('visible');
        }
    });
    
    // Smooth scrolling for TOC links
    document.addEventListener('click', function(e) {
        if (e.target.closest('.toc-link')) {
            e.preventDefault();
            const targetId = e.target.closest('.toc-link').getAttribute('href');
            const targetElement = document.querySelector(targetId);
            if (targetElement) {
                targetElement.scrollIntoView({
                    behavior: 'smooth',
                    block: 'center'
                });
                floatingToc.classList.remove('visible');
            }
        }
    });
});

function generateTOC() {
    const tocList = document.getElementById('tocList');
    if (!tocList) return;
    
    // Clear existing TOC
    tocList.innerHTML = '';
    
    // Find all headings
    const headings = document.querySelectorAll('h1, h2, h3, h4');
    
    if (headings.length === 0) return;
    
    // Create TOC items
    headings.forEach((heading, index) => {
        // Ensure heading has an ID
        if (!heading.id) {
            heading.id = 'heading-' + index;
        }
        
        // Create TOC item
        const tocItem = document.createElement('li');
        tocItem.className = 'toc-item';
        
        const tocLink = document.createElement('a');
        tocLink.className = 'toc-link indent-' + (heading.tagName.charAt(1) - 1);
        tocLink.href = '#' + heading.id;
        tocLink.textContent = heading.textContent;
        
        tocItem.appendChild(tocLink);
        tocList.appendChild(tocItem);
    });
}

// Theme toggle functionality
document.addEventListener('DOMContentLoaded', function() {
    const themeToggle = document.getElementById('themeToggle');
    if (themeToggle) {
        const themeIcon = themeToggle.querySelector('i');
        
        // Check for saved theme preference or respect OS preference
        const savedTheme = localStorage.getItem('theme');
        const osPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        const currentTheme = savedTheme || (osPrefersDark ? 'dark' : 'light');
        
        // Apply theme
        document.body.classList.toggle('light-theme', currentTheme === 'light');
        if (themeIcon) {
            themeIcon.className = currentTheme === 'light' ? 'fas fa-moon' : 'fas fa-sun';
        }
        
        // Toggle theme
        themeToggle.addEventListener('click', () => {
            document.body.classList.toggle('light-theme');
            const isLight = document.body.classList.contains('light-theme');
            if (themeIcon) {
                themeIcon.className = isLight ? 'fas fa-moon' : 'fas fa-sun';
            }
            localStorage.setItem('theme', isLight ? 'light' : 'dark');
        });
    }
});

// Bookmark functionality for individual pages
function togglePageBookmark() {
    const currentPage = {
        title: document.title.replace(' - Godzilla Bot Документація', ''),
        path: window.location.pathname,
        category: getCategoryFromPath()
    };
    
    // Get existing bookmarks from localStorage
    let bookmarks = JSON.parse(localStorage.getItem('bookmarks')) || [];
    
    // Check if page is already bookmarked
    const isBookmarked = bookmarks.some(bookmark => bookmark.path === currentPage.path);
    
    if (isBookmarked) {
        // Remove bookmark
        bookmarks = bookmarks.filter(bookmark => bookmark.path !== currentPage.path);
    } else {
        // Add bookmark
        bookmarks.push(currentPage);
    }
    
    // Save to localStorage
    localStorage.setItem('bookmarks', JSON.stringify(bookmarks));
    
    // Update bookmark button UI if it exists
    const bookmarkBtn = document.getElementById('pageBookmarkBtn');
    if (bookmarkBtn) {
        const icon = bookmarkBtn.querySelector('i');
        if (icon) {
            icon.className = isBookmarked ? 'fas fa-bookmark' : 'fas fa-bookmark';
        }
        bookmarkBtn.title = isBookmarked ? 'Додати в закладки' : 'Видалити з закладок';
    }
    
    return !isBookmarked; // Return new bookmark state
}

function getCategoryFromPath() {
    const path = window.location.pathname;
    if (path.includes('/api/')) return 'api';
    if (path.includes('/guides/')) return 'guides';
    if (path.includes('/architecture/')) return 'architecture';
    if (path.includes('/security/')) return 'security';
    if (path.includes('/deployment/')) return 'deployment';
    if (path.includes('/changelog/')) return 'changelog';
    if (path.includes('/en/')) return 'en';
    return 'root';
}

// Initialize bookmark button on page load
document.addEventListener('DOMContentLoaded', function() {
    // Add bookmark button to header if it doesn't exist
    const header = document.querySelector('header');
    if (header && !document.getElementById('pageBookmarkBtn')) {
        const bookmarkBtn = document.createElement('button');
        bookmarkBtn.id = 'pageBookmarkBtn';
        bookmarkBtn.className = 'bookmark-btn';
        bookmarkBtn.title = 'Додати в закладки';
        bookmarkBtn.innerHTML = '<i class="fas fa-bookmark"></i>';
        bookmarkBtn.addEventListener('click', togglePageBookmark);
        header.style.position = 'relative';
        header.appendChild(bookmarkBtn);
    }
});

// Rating functionality for individual pages
function rateDocument(rating) {
    const currentPagePath = window.location.pathname;
    
    // Get existing ratings from localStorage
    let ratings = JSON.parse(localStorage.getItem('documentRatings')) || {};
    
    // Save rating for this document
    ratings[currentPagePath] = rating;
    
    // Save to localStorage
    localStorage.setItem('documentRatings', JSON.stringify(ratings));
    
    // Update rating UI
    updateRatingUI(rating);
    
    return rating;
}

function updateRatingUI(rating) {
    const ratingContainer = document.getElementById('pageRating');
    if (ratingContainer) {
        const stars = ratingContainer.querySelectorAll('.rating-star');
        stars.forEach((star, index) => {
            star.classList.toggle('active', index < rating);
        });
        
        const ratingValue = ratingContainer.querySelector('.rating-value');
        if (ratingValue) {
            ratingValue.textContent = rating > 0 ? rating.toFixed(1) : '';
        }
    }
}

// Initialize rating functionality
document.addEventListener('DOMContentLoaded', function() {
    // Add rating section to content if it doesn't exist
    const content = document.querySelector('.content');
    if (content && !document.getElementById('pageRating')) {
        const ratingSection = document.createElement('div');
        ratingSection.id = 'pageRating';
        ratingSection.className = 'rating-container';
        ratingSection.innerHTML = `
            <p>Оцініть цей документ:</p>
            <div class="rating-stars">
                ${[1, 2, 3, 4, 5].map(star => 
                    `<span class="rating-star" data-rating="${star}">&#9733;</span>`
                ).join('')}
            </div>
            <span class="rating-value"></span>
        `;
        
        // Add event listeners to stars
        const stars = ratingSection.querySelectorAll('.rating-star');
        stars.forEach(star => {
            star.addEventListener('click', function() {
                const rating = parseInt(this.dataset.rating);
                rateDocument(rating);
            });
        });
        
        content.appendChild(ratingSection);
    }
    
    // Load and display existing rating
    const currentPagePath = window.location.pathname;
    const ratings = JSON.parse(localStorage.getItem('documentRatings')) || {};
    const existingRating = ratings[currentPagePath];
    if (existingRating) {
        updateRatingUI(existingRating);
    }
});