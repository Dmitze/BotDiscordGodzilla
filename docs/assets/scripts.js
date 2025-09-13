// Simple script to enhance documentation interactivity

document.addEventListener('DOMContentLoaded', function() {
    // Add smooth scrolling for anchor links
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });
    
    // Add copy button to code blocks
    document.querySelectorAll('pre').forEach(pre => {
        const code = pre.querySelector('code');
        if (code) {
            const button = document.createElement('button');
            button.className = 'copy-button';
            button.innerHTML = '<i class="fas fa-copy"></i> Копіювати';
            button.style.position = 'absolute';
            button.style.top = '10px';
            button.style.right = '10px';
            button.style.background = 'rgba(88, 101, 242, 0.2)';
            button.style.color = '#e6e6e6';
            button.style.border = '1px solid var(--primary)';
            button.style.borderRadius = '4px';
            button.style.padding = '5px 10px';
            button.style.cursor = 'pointer';
            button.style.fontSize = '0.8rem';
            
            pre.style.position = 'relative';
            pre.appendChild(button);
            
            button.addEventListener('click', function() {
                navigator.clipboard.writeText(code.textContent).then(() => {
                    const originalText = button.innerHTML;
                    button.innerHTML = '<i class="fas fa-check"></i> Скопійовано!';
                    setTimeout(() => {
                        button.innerHTML = originalText;
                    }, 2000);
                });
            });
        }
    });
    
    // Add table of contents for pages with many headings
    const headings = document.querySelectorAll('.content h2, .content h3');
    if (headings.length > 3) {
        const toc = document.createElement('div');
        toc.className = 'table-of-contents';
        toc.innerHTML = '<h3>Зміст</h3><ul></ul>';
        toc.style.background = 'rgba(88, 101, 242, 0.1)';
        toc.style.padding = '1rem';
        toc.style.borderRadius = '8px';
        toc.style.marginBottom = '2rem';
        
        const tocList = toc.querySelector('ul');
        tocList.style.listStyle = 'none';
        tocList.style.paddingLeft = '0';
        
        headings.forEach((heading, index) => {
            // Add ID to heading if it doesn't have one
            if (!heading.id) {
                heading.id = 'heading-' + index;
            }
            
            const listItem = document.createElement('li');
            listItem.style.marginBottom = '0.5rem';
            listItem.style.paddingLeft = heading.tagName === 'H3' ? '1rem' : '0';
            
            const link = document.createElement('a');
            link.href = '#' + heading.id;
            link.textContent = heading.textContent;
            link.style.color = '#a1a1aa';
            link.style.textDecoration = 'none';
            link.style.transition = 'color 0.3s ease';
            
            link.addEventListener('hover', function() {
                this.style.color = '#5865F2';
            });
            
            listItem.appendChild(link);
            tocList.appendChild(listItem);
        });
        
        // Insert TOC after the first heading
        const content = document.querySelector('.content');
        if (content) {
            content.insertBefore(toc, content.firstChild);
        }
    }
});