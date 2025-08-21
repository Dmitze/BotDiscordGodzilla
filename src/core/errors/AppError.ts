export class AppError extends Error {
  public readonly code: string;
  public readonly userMessageKey: string;
  public readonly cause?: unknown;
  public readonly meta?: Record<string, unknown>;

  constructor(code: string, userMessageKey: string, cause?: unknown, meta?: Record<string, unknown>) {
    super(typeof cause === 'string' ? cause : (cause instanceof Error ? cause.message : code));
    this.name = 'AppError';
    this.code = code;
    this.userMessageKey = userMessageKey;
    this.cause = cause;
    if (meta !== undefined) {
      this.meta = meta;
    }
  }

  toLog(): { code: string; message: string; meta?: Record<string, unknown> | undefined } {
    const causeMsg = this.cause instanceof Error ? this.cause.message : (typeof this.cause === 'string' ? this.cause : '');
    return {
      code: this.code,
      message: causeMsg || this.message || this.code,
      ...(this.meta ? { meta: this.meta } : {}),
    };
  }
}
