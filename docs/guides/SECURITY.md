# Security Guide

This project implements multiple layers of security to protect users and reduce abuse surface. This document explains the key mechanisms and how to configure them.

## Security Config (`src/config/security.ts`)

All security-related settings are centralized and validated with Zod:

- HMAC key and TTL for component IDs
- Allowlist of MIME types and maximum file size
- PII masking flags (apply to logs and user-visible replies)

Environment variables are validated at startup. See `.env.example` for the minimal required envs and values.

## Signed Component IDs (HMAC + TTL)

Discord message components (buttons/selects) can be interacted with outside the original context. To prevent replay and tampering, we sign `customId` values with HMAC and embed a short-lived timestamp.

- Sign: `signComponentId(payload)` -> returns a string customId
- Verify: `verifyComponentId(customId)` -> returns `{ valid, reason?: 'expired'|'invalid', payload? }`
- TTL: configured via `securityConfig.components.ttlSec`

### Conventions

- Payloads include `kind` to route handling, e.g. `kind: 'srch'`, `kind: 'filetxt'`, `kind: 'docblk'`
- Signed IDs are backward compatible: handlers fall back to legacy parsers if no valid signature is present

### Base Handling

`BaseCommand.handleComponent()` automatically:

1. Detects signed IDs and verifies them via `verifyComponentId()`
2. On failure, responds with i18n messages:
   - `security.component.invalidId`
   - `security.component.expiredId`
3. On success, injects `context.componentPayload` to downstream handlers

Command overrides should prefer `options.context.componentPayload` when available and only parse legacy `customId` when not.

## PII Masking (`src/utils/pii.ts`)

All user-facing replies and logs pass through masking helpers to redact emails and phone numbers:

- `maskPII(text)` – returns redacted string
- `maskReplyOptions(options)` – applies masking to `content` and embed fields
- `replyWithPrivacy()` from `src/ui/reply.ts` wraps Discord replies to ensure masking is applied consistently (ephemeral by default, can be made public via share flag where applicable)

## i18n for Security Errors (`src/i18n/*.json`)

Security error messages are localized and selected per user automatically:

- `security.component.invalidId`
- `security.component.expiredId`

Use `tUser(key, interaction)` to render messages in the user's locale. Default locale is Ukrainian (`uk`).

## Testing

- Unit/e2e tests should avoid real timeouts by either:
  - Using short TTL in test env
  - Stubbing `verifyComponentId()` to return deterministic results
- In test mode (`NODE_ENV=test`), health checks are forced to healthy to avoid external dependencies affecting test stability.

## Operational Notes

- Rotate HMAC keys periodically. Rotation strategy: accept old signatures for a grace window if needed.
- Keep TTL short (e.g., 5–10 minutes) to reduce replay windows.
- Ensure ephemeral responses for invalid/expired interactions to avoid leaking context in channels.
