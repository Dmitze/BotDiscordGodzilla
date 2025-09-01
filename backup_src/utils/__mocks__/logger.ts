const mockLogger = {
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn(),
  debug: jest.fn(),
  security: jest.fn(),
  performance: jest.fn(),
  commands: jest.fn(),
  api: jest.fn(),
  system: jest.fn(),
  cleanup: jest.fn(),
};

// ESM default export
export default mockLogger;

// Also provide named export for compatibility
export { mockLogger };

// Ensure CommonJS interop (for require())
// eslint-disable-next-line @typescript-eslint/no-explicit-any
(module as any).exports = mockLogger;
