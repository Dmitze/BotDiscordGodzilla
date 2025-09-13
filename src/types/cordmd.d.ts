declare module 'cordmd' {
  export function renderMarkdown(markdown: string, options?: any): Promise<Buffer>;
  export function validateMarkdown(markdown: string): { isValid: boolean; errors: string[]; warnings: string[] };
}