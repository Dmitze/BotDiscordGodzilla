declare module 'pdf-parse' {
  interface PdfParseResult {
    numpages: number;
    numrender: number;
    info: Record<string, any>;
    metadata: any;
    text: string;
    version: string;
  }
  function pdfParse(data: Buffer | Uint8Array | ArrayBuffer): Promise<PdfParseResult>;
  export default pdfParse;
}
