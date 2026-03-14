# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

A Python CLI tool for Outlook 365 that uses OWA bearer token authentication via Playwright browser interception — no Azure app registration, admin consent, or API keys required. Entry point: `outlook` command.

## Build & Run

```sh
pip install -e .              # editable install (hatchling build system)
playwright install chromium   # required for auth
outlook login                 # first-time: opens browser, captures OWA bearer token
outlook inbox                 # verify it works
```

```sh
pytest               # run all 83 unit tests
pytest -m smoke      # run only smoke tests (require live token)
```

## Architecture

### Two API layers

1. **Outlook REST v2** (`outlook.office.com/api/v2.0/me`) — standard mail, calendar, contacts, folders, per-message categories. Used by `OutlookClient` in `client.py`.
2. **OWA service.svc** (`outlook.cloud.microsoft/owa/service.svc`) — reverse-engineered endpoint for master category list operations (create/delete/rename/recolor) and message pinning (`UpdateItem` with `RenewTime`). Uses a non-standard pattern: JSON payload goes in the `x-owa-urlpostdata` header, body is empty. Used by `category_manager.py` and `client.py` (`pin_message`).

### Module responsibilities

- **`cli.py`** — Click group definition + command registration hub only (~92 lines). Imports from `commands/` modules.
- **`commands/`** — All CLI commands split into modules:
  - `_common.py` — shared helpers: `_get_client`, `_handle_api_error`, `_wants_json`, `cfg`
  - `auth.py` — `login`, `whoami`
  - `mail.py` — `inbox`, `read`, `thread`, `send`, `draft`, `draft-send`, `reply`, `reply-draft`, `forward`
  - `schedule.py` — `schedule`, `schedule-list`, `schedule-cancel`, `schedule-draft`
  - `search.py` — `search`
  - `folders.py` — `folders`, `folder`
  - `categories.py` — `categories`, `categorize`, `uncategorize`, `category-create/rename/clear/delete`
  - `signatures.py` — `signature-pull`, `signature-list`, `signature-show`, `signature-delete`
  - `manage.py` — `mark-read`, `move`, `delete`, `flag`, `pin`
  - `attachments.py` — `attachments`
  - `calendar.py` — `calendar`, `event`, `event-create/update/delete/instances/respond`, `calendars`, `free-busy`, `people-search`
  - `contacts.py` — `contacts`
- **`exceptions.py`** — Structured exception hierarchy: `OutlookCliError` → `TokenExpiredError`, `RateLimitError`, `ResourceNotFoundError`, `AuthRequiredError`. Includes `error_code_for_exception()` mapping.
- **`client.py`** — `OutlookClient` wraps httpx for REST v2 API. Manages display-number-to-real-ID mapping (short `#1, #2` numbers → long Outlook IDs). Handles rate limiting (429 retry) and token expiry (401). `get_thread()` fetches conversation chains.
- **`auth.py`** — Playwright-based token capture. Intercepts bearer tokens from OWA network requests. Picks the best token by testing against multiple endpoints. Caches token + browser SSO state.
- **`category_manager.py`** — Standalone module for OWA master category operations. Has its own `_owa_request` helper (separate from `client.py`'s `_owa_action`). `rename_category` and `clear_category` do bulk message propagation via REST v2.
- **`signature_manager.py`** — Signature management: pull from SentItems, save as HTML files in `~/.config/outlook-cli/signatures/`, append to outgoing emails. Handles plain text → HTML conversion when signature is used.
- **`models.py`** — Dataclasses (`Email`, `Folder`, `Attachment`, `Event`, `Attendee`, `Contact`, `EmailAddress`) with `from_api()` class methods that parse Outlook REST v2 JSON. `Email` includes `categories: list[str]`, `flag_status` ("notFlagged"/"flagged"/"complete"), `flag_due: datetime | None`. `Event` includes `attendees: list[Attendee]`, `recurrence`, `event_type` (SingleInstance/Occurrence/Exception/SeriesMaster), `series_master_id`, `display_num`.
- **`formatter.py`** — Rich table output. `Console(stderr=True)` so JSON piping stays clean on stdout. `print_thread()` for conversation view. Inbox flags column shows `*` (unread), `@` (attachment), `!` (flagged), `v` (flag complete). Email detail view shows flag status with due date.
- **`serialization.py`** — `to_json_envelope()` wraps data in `{ok, schema_version, data}` for stdout. `error_json()` for structured errors. `to_json()` / `save_json()` for raw file export.
- **`config.py`** — YAML config loader with deep-merge defaults.
- **`constants.py`** — URLs, cache/config paths.

### Key patterns

- **Display number ID mapping**: Messages and events get short `#1, #2...` numbers stored in `~/.cache/outlook-cli/id_map.json`. Users reference items by these numbers. The map is capped at 500 entries with LRU eviction. Events share the same ID map as messages.
- **Multi-ID commands**: `delete`, `move`, `mark-read`, `categorize`, `uncategorize`, `flag`, `pin` accept multiple message IDs via Click's `nargs=-1`. The variadic argument comes first, fixed argument (destination/category) last.
- **Send confirmation**: `send`, `reply`, `forward`, `draft-send`, `schedule`, `schedule-draft`, `event-create` show details and require confirmation before action. All accept `-y` to skip. Draft-creation commands (`draft`, `reply-draft`) do NOT require confirmation since nothing is sent. `event-delete` also confirms unless `-y`.
- **Draft reply**: `reply-draft` uses `createReply` / `createReplyAll` REST v2 endpoints to create reply drafts with original recipients pre-filled. Body argument is optional (default empty).
- **Scheduled send**: Uses `PidTagDeferredSendTime` (0x3FEF) extended property. `schedule` uses `/sendmail` with the property inline. `schedule-draft` PATCHes an existing draft then sends it. Tracked locally in `scheduled.json` (REST v2 doesn't support `$filter`/`$expand` on extended properties). `schedule-list` cross-references local tracking with Drafts folder by subject to find matching draft IDs. `schedule-cancel` deletes both local tracking and the server draft when found. Time formats: `+30m`, `+1h`, `tomorrow 09:00`, `2024-03-15T10:00`.
- **`$filter` vs `$search` split**: REST v2 can't combine `$filter` and `$search`. Text filters (from/subject/hasattachments) use KQL `$search` (no `$orderby`). Date/read/category filters use `$filter` (supports `$orderby`). See `_build_query_params` in `client.py`.
- **`--no-category` client-side filtering**: REST v2 can't filter for empty `Categories` array. `get_messages` over-fetches in pages (3x batch, max 5 pages) and filters locally to guarantee `--max` count.
- **Signature extraction**: `signature_manager.py` parses SentItems HTML to find the outermost `<table>` containing `mailto:` links. Signatures are stored as plain HTML files — no API dependency.
- **Conversation thread**: `thread` command fetches all messages with the same `ConversationId`. REST v2 doesn't support `$filter` on `ConversationId`, so `get_thread()` searches by base subject (strips Re:/Fwd:/İlt:/Ynt: prefixes) then filters client-side by `ConversationId`. Results sorted oldest-first.
- **Structured JSON envelope**: All `--json` output wraps data in `{ok: true, schema_version: "1", data: [...]}`. Errors return `{ok: false, error: {code, message}}`. Error codes: `session_expired`, `rate_limited`, `not_found`, `not_authenticated`, `unknown_error`. File export (`-o` flag) stays raw (no envelope).
- **Auto-JSON on pipe**: When stdout is not a TTY (piped to `jq`, `grep`, etc.), commands automatically output JSON envelope — no `--json` flag needed. Controlled by `_is_piped()` / `_wants_json()` in `commands/_common.py`.
- **Token flow**: env var `OUTLOOK_TOKEN` → cached `token.json` → interactive Playwright login. Auto re-login on 401 via `_handle_api_error` decorator in `commands/_common.py`.
- **Pin messages**: `pin` uses OWA `service.svc` `UpdateItem` action with `RenewTime` field (not REST v2). Pin sets `RenewTime` to far-future date (`4500-09-01`), unpin deletes the field. Message IDs must be converted from URL-safe base64 (`-`, `_`) to standard base64 (`/`, `+`) for OWA compatibility.
- **File attachments**: `send`, `draft`, `reply`, `reply-draft`, `forward`, and `schedule` accept `--attach`/`-a` (repeatable). When attachments are present, commands use a draft flow: create draft → attach files → send. Small files (<3 MB) use inline base64 via `POST /messages/{id}/attachments`. Large files (>=3 MB) use upload sessions via `createuploadsession` + chunked PUT. `create_forward_draft` uses `POST /messages/{id}/createforward`. Click's `type=click.Path(exists=True)` validates files before execution.
- **Dual OWA helpers**: `client.py` has `_owa_action` and `category_manager.py` has `_owa_request` — both call OWA service.svc with slightly different base URLs (`outlook.office365.com` vs `outlook.cloud.microsoft`).
- **Calendar CRUD**: Full event lifecycle via REST v2: `POST /events` (create), `GET /events/{id}` (read), `PATCH /events/{id}` (update), `DELETE /events/{id}` (delete). Attendee management via `add_event_attendees`/`remove_event_attendees` (GET existing + PATCH merged list). Meeting responses via `POST /events/{id}/{accept|decline|tentativelyaccept}`.
- **Shared calendars**: `--calendar "Name"` resolves display name → ID via `_resolve_calendar` (exact match first, then partial). Queries `/me/calendars/{id}/calendarview` instead of `/me/calendarview`.
- **Recurrence**: `event-create --repeat daily|weekly|monthly` builds `Recurrence` payload with Pattern (Type, Interval, DaysOfWeek, DayOfMonth) + Range (Numbered/EndDate). `event-instances` lists occurrences via `/events/{master_id}/instances` — auto-resolves occurrence → series master via `SeriesMasterId`. `event-delete --series` deletes via series master ID.
- **Free/busy**: `findMeetingTimes` endpoint with attendees, time constraints, duration. Returns MeetingTimeSuggestions with confidence scores.
- **People search**: `/me/people?$search=query` for attendee autocomplete. Returns `ScoredEmailAddresses`.

### Cache & config locations

- Cache: `~/.cache/outlook-cli/` (token.json, browser-state.json, id_map.json, scheduled.json)
- Config: `~/.config/outlook-cli/config.yaml`
- Signatures: `~/.config/outlook-cli/signatures/*.html`
- Overridable via `OUTLOOK_CLI_CACHE` and `OUTLOOK_CLI_CONFIG` env vars

### Dependencies

click, rich, httpx, playwright, PyYAML, beautifulsoup4. Python >=3.10. Build: hatchling.
