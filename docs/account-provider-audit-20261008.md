# Account and provider audit — 2026-10-08

## Scope and fixed defects

- Browser agent/session caches now use an account namespace; ownerless legacy caches are not imported.
- Logout/new-login invalidates pending model, chat-list, history and stream responses, including signing back into the same account.
- Visible model selection is captured in each JSON/file request. Preferences are per account, not process-global, and survive refresh.
- Normal chat retries authentication/capacity errors at the failed LLM call, not by replaying the whole agent graph and tools. A failed authentication Key/URL pair is skipped for the rest of that call.
- Provider routing binds model, key and URL together. The legacy Zhipu/Anthropic chat backup pair is no longer mixed with an OpenAI-compatible Ark model. Ark Coding backup uses `glm-5.3` and a Coding key.
- No fallback after response/tool-call content was emitted. Network timeouts and invalid-request errors are not reasons to switch models.
- Doubao Pro migrates to `doubao-seed-2.1-pro`; legacy saved IDs are accepted as aliases. Text GLM is `glm-5.3`; image understanding is Ark `glm-5.3-flash` (not text-only GLM).
- Existing V4.1 schedule is preserved: Beijing 09:00–12:00 / 14:00–18:00 uses original V4 Flash with its original dedicated API configuration. Other hours retain the separate V4.1 key. UI selection does not change.
- Embedding API, dependencies, prompt/templates, long-manual editing algorithm/timeouts, survey field layout and progress-card styling were not changed.

## Provider prerequisites

Reuse the server's existing `DEEPSEEK_API_KEY` / `DEEPSEEK_BASE_URL` for Ark Coding, or set `ARK_CODING_API_KEY` / `ARK_CODING_BASE_URL` explicitly. Coding base URL is `https://ark.cn-beijing.volces.com/api/coding/v3`. Do not substitute a standard Ark `/api/v3` key. No credential is committed in these fixes; the server `.env` is not uploaded or replaced.

Configured credentials, account balance and server-side model entitlement have **not** been verified by paid/live LLM requests. No login window is used for GitHub publication. Server deployment is separate.

Official provider references:

- https://docs.volcengine.com/docs/ark/coding-plan-enterprise-ai-opencode?lang=zh
- https://docs.volcengine.com/docs/ark/coding-plan-personal-ai-zcode?lang=zh
- https://docs.volcengine.com/docs/ark/coding-plan-personal-model-deprecation?lang=zh

## Regression commands

Run each Python check in a separate process to avoid old unit-test fake-module fixtures contaminating real integration imports:

```text
python integration/check_model_routing.py
node --test tests/frontend_account_isolation.test.cjs
```

These tests use fake providers and temporary local storage; they do not call paid model APIs.

## DFAGENT3 evidence and boundaries

- Chat/list/history/clear/delete/rename/export/file-chat/model endpoints now require authentication and exact stored chat ownership. Admin versus adminquanzhi is not distinguished by prefix matching.
- Request model uses a ContextVar and model-keyed graph cache, not a global first-user model. Generator construction/next/close each binds and resets independently.
- Decoded file-upload names reject path separators, traversal components and NUL before disk operations.
- New-chat creation no longer deletes a previous conversation when a third chat is created for the same agent. Explicit deletion remains.
- New model checks: 21 passed, 4 QZ-only cases skipped. New browser isolation checks: 11 passed.
- Production JWT/handler/storage regression: 22 foreign/default-session denials, 40 parallel own-history reads, invalid/missing JWT and path/body ID mismatch checks passed. Creating additional same-agent chats preserves old history.
- Real ASGI import and construction of all 5 catalog clients passed without provider invocation.
- Remaining **high security risk**: legacy document/knowledge-base download/export routes (`download_document`, `export_document_api`, `export_xlsx_api`, `download_export_document`) still do not declare authentication. Chat ownership fixes are not an authorization audit of all document endpoints; do not expose these legacy endpoints as though the entire application were secured.
- Remaining conditional **high security risk**: without a random explicit `JWT_SECRET`, the signing secret is a predictable hash of `DATA_DIR`. Configure a random server-only secret; changing it invalidates existing login tokens. The deployment secret was not silently changed.
- Remaining **multi-worker storage risk**: legacy conversation/user-list JSON writes and process-local caches are not a cross-process transaction store. Parallel read checks are not concurrent write/load certification.

Additional offline check:

```text
python integration/check_account_api.py
```
