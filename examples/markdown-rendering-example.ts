/**
 * Example script demonstrating markdown rendering capabilities
 * 
 * This script shows how to use the MarkdownRenderingService to render
 * markdown content in different formats.
 */

import { MarkdownRenderingService } from '../src/services/MarkdownRenderingService';

async function demonstrateMarkdownRendering() {
  // Initialize the service
  const markdownService = MarkdownRenderingService.getInstance({
    // Minimal config for demonstration
  } as any);
  
  // Example 1: Simple text rendering
  console.log('=== Text Rendering Example ===');
  const simpleMarkdown = `
# Welcome to BotDiscordGodzilla

This is a **bold** statement and this is *italic* text.

## Code Example

Here's some JavaScript code:

\`\`\`javascript
function helloWorld() {
  console.log('Hello, World!');
  return true;
}
\`\`\`

## List Example

1. First item
2. Second item
3. Third item

- Bullet point 1
- Bullet point 2
- Bullet point 3
  `;
  
  try {
    const renderedText = await markdownService.renderToText(simpleMarkdown);
    console.log('Rendered Text:');
    console.log(renderedText);
  } catch (error) {
    console.error('Error rendering text:', error);
  }
  
  // Example 2: Image rendering
  console.log('\n=== Image Rendering Example ===');
  const complexMarkdown = `
# Complex Document

## Table Example

| Name | Age | City |
|------|-----|------|
| John | 30  | New York |
| Jane | 25  | London |
| Bob  | 35  | Paris |

## Mathematical Expressions

E = mc²

## Ukrainian Text Example

# Привіт Світ

Це приклад українського тексту з **жирним** та *курсивним* форматуванням.

\`\`\`python
def привіт_світ():
    print("Привіт, Світ!")
    return True
\`\`\`
  `;
  
  try {
    const attachment = await markdownService.renderToImage(complexMarkdown);
    console.log('Image rendered successfully:', attachment.name);
  } catch (error) {
    console.error('Error rendering image:', error);
  }
  
  // Example 3: Code block extraction
  console.log('\n=== Code Block Extraction Example ===');
  const codeBlocks = markdownService.extractCodeBlocks(complexMarkdown);
  console.log('Extracted code blocks:', codeBlocks.length);
  codeBlocks.forEach((block, index) => {
    console.log(`Block ${index + 1}:`);
    console.log(`  Language: ${block.language}`);
    console.log(`  Content: ${block.content.substring(0, 50)}...`);
  });
  
  // Example 4: Validation
  console.log('\n=== Validation Example ===');
  const validResult = markdownService.validateMarkdown(simpleMarkdown);
  console.log('Valid markdown result:', validResult);
  
  const invalidResult = markdownService.validateMarkdown('This is invalid markdown **bold*');
  console.log('Invalid markdown result:', invalidResult);
  
  // Example 5: Metrics
  console.log('\n=== Metrics Example ===');
  const metrics = markdownService.getMetrics();
  console.log('Rendering metrics:', metrics);
}

// Run the demonstration
demonstrateMarkdownRendering().catch(console.error);