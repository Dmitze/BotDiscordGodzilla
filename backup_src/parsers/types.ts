export interface ParsedDoc {
  id: string; // stable document id (e.g., pdf:<fileId> or sheet:<fileId>#<sheetName>:<range>)
  text?: string; // full normalized text
  segments?: string[]; // optional pre-split segments
  lang?: string; // e.g. 'uk' | 'en' | 'unknown'
  labels?: string[]; // custom labels
  updatedAt: number; // unix seconds
  meta: { path: string; type: string; range?: string; sheet?: string };
}

export interface Parser {
  supports(mime: string): boolean;
  // meta may include Google file metadata, drive ids, ranges etc.
  parse(input: Buffer | NodeJS.ReadableStream, meta: any): Promise<ParsedDoc>;
}
