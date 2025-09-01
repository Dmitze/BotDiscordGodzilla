import type { DriveFile } from '@/types/drive';

export const fileDoc: DriveFile = {
  id: 'doc1',
  name: 'Doc One',
  mimeType: 'application/vnd.google-apps.document',
  modifiedTime: '2025-08-10T10:00:00Z',
};

export const filePdf: DriveFile = {
  id: 'pdf1',
  name: 'PDF One',
  mimeType: 'application/pdf',
  modifiedTime: '2025-08-11T11:00:00Z',
};

export const fileWord: DriveFile = {
  id: 'docx1',
  name: 'Word One',
  mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  modifiedTime: '2025-08-12T12:00:00Z',
};

export const nonIndexable: DriveFile = {
  id: 'img1',
  name: 'Image',
  mimeType: 'image/png',
  modifiedTime: '2025-08-13T13:00:00Z',
};

export function clone<T>(obj: T): T { return JSON.parse(JSON.stringify(obj)); }
