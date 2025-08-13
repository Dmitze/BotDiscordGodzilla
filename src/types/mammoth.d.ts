declare module 'mammoth' {
  export interface MammothResult {
    value: string;
    messages: Array<{ type: string; message: string }>
  }
  export function extractRawText(input: { buffer: Buffer }): Promise<MammothResult>;
  const _default: {
    extractRawText: typeof extractRawText;
  };
  export default _default;
}
