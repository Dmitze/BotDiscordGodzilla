PRAGMA journal_mode=WAL;
PRAGMA synchronous=NORMAL;

CREATE TABLE IF NOT EXISTS documents (
  fileId TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  mimeType TEXT,
  ownerEmail TEXT,
  sizeBytes INTEGER,
  modifiedTime INTEGER,
  createdTime INTEGER,
  contentHash TEXT NOT NULL,
  textLen INTEGER NOT NULL,
  lastIndexedAt INTEGER NOT NULL,
  tags TEXT,
  meta TEXT
);

CREATE VIRTUAL TABLE IF NOT EXISTS documents_fts USING fts5(
  fileId UNINDEXED,
  name,
  content,
  tokenize='porter'
);

CREATE TABLE IF NOT EXISTS document_versions (
  fileId TEXT NOT NULL,
  version INTEGER NOT NULL,
  contentHash TEXT NOT NULL,
  textLen INTEGER NOT NULL,
  modifiedTime INTEGER,
  createdAt INTEGER NOT NULL,
  meta TEXT,
  PRIMARY KEY(fileId, version)
);

CREATE INDEX IF NOT EXISTS idx_documents_mime ON documents(mimeType);
CREATE INDEX IF NOT EXISTS idx_documents_owner ON documents(ownerEmail);
CREATE INDEX IF NOT EXISTS idx_documents_modified ON documents(modifiedTime);
