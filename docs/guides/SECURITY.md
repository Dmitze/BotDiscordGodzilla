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
- TTL: configured via `securityConfig.components.ttlMs`

### Conventions

- Payloads include `kind` to route handling, e.g. `kind: 'srch'`, `kind: 'filetxt'`, `kind: 'docblk'`
- Compact format prefix `c.` is used to keep IDs short; fields are shortened (e.g., `kind -> k`, `page -> p`).
- Signed IDs are backward compatible: handlers fall back to legacy parsers if no valid signature is present

#### Drive buttons (`kind: 'drive'`)

UI‑картки файлів (`src/ui/FileCardBuilder.ts`) створюють підписані `customId` для дій Google Drive:

- Дії: `action ∈ {'open'|'download'|'summary'|'question'}`
- Формат payload: `{ kind: 'drive', action, id: <driveFileId> }`
- Обробка: у `FileManagerCommand.onComponent()` спочатку читається `options.context.componentPayload` (після валідації HMAC/TTL в `BaseCommand`), далі виконується дія:
  - `open` → відповідає посиланням `https://drive.google.com/file/d/<id>/view`
  - `download` → відповідає посиланням `https://drive.google.com/uc?export=download&id=<id>`
  - `summary` → викликає пайплайн аналізу (модуль `analyzers.ts`) з `analysisType: 'summary'`
  - `question` → інтеграція з RAG/AI: якщо доступний `RagService`, використовується `rag.answer(...)`, інакше — `AIService.generateResponse(...)`

Тестовий режим (NODE_ENV=test): для стабільності юніт‑тестів `FileCardBuilder` зберігає fallback‑формат без підпису — рядок `drive:<action>:<base64({id})>`. Обробник має легасі‑парсер, який активується лише якщо підписаний payload відсутній.

### Base Handling

`BaseCommand.handleComponent()` automatically:

1. Detects signed IDs and verifies them via `verifyComponentId()`
2. On failure, responds with i18n messages:
   - `security.component.invalidId`
   - `security.component.expiredId`
3. On success, injects `context.componentPayload` to downstream handlers

Command overrides should prefer `options.context.componentPayload` when available and only parse legacy `customId` when not.

## Environment Variables

Configure component signing and security using the following env vars (see `env.example`):

```dotenv
# HMAC key (>=16 chars) and TTL for signed component IDs (milliseconds)
COMPONENT_HMAC_KEY=change_me_please_min16chars
COMPONENT_TTL_MS=900000

# File and PII policies
SECURITY_MIME_ALLOWLIST=application/pdf,application/vnd.openxmlformats-officedocument.wordprocessingml.document,text/plain
SECURITY_MAX_BYTES=26214400
SECURITY_PII_MASTER=true
SECURITY_PII_EMAIL=true
SECURITY_PII_PHONE=true
```

Notes:

- Keep TTL short (e.g., 5–15 minutes) to minimize replay windows.
- Rotate `COMPONENT_HMAC_KEY` periodically. Consider a short grace period for key rotation if needed.

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
