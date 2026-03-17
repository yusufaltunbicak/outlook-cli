---
name: outlook-cli
description: CLI skill for Outlook 365 to read, send, search, and manage emails, calendar events, categories, and contacts from the terminal without API keys or admin consent
author: yusufaltunbicak
version: "0.1.2"
tags:
  - outlook
  - email
  - office365
  - calendar
  - events
  - meetings
  - categories
  - attachments
  - terminal
  - cli
---

# outlook-cli Skill

Use this skill when the user wants to read, send, search, or manage Outlook 365 emails and calendar events from the terminal. Also supports file attachments, follow-up flags, message pinning, recurring events, shared calendars, free/busy scheduling, categories, contacts, and signatures.

## Prerequisites

```bash
# Install (requires Python 3.10+)
cd ~/outlook-cli && pip install -e .
playwright install chromium
```

## Authentication

- First run: `outlook login` opens Chromium, user logs in, bearer token is auto-captured from OWA requests.
- Named profiles are supported via `outlook account add|list|current|switch|remove`.
- Tokens, browser state, ID maps, scheduled-send tracking, signatures, and per-profile config are isolated per account profile.
- Profile-scoped cache lives under `~/.cache/outlook-cli/accounts/<profile>/`; config lives under `~/.config/outlook-cli/accounts/<profile>/`.
- Existing single-account installs continue to work through the implicit `default` profile and legacy cache paths.
- Auto re-login on 401 is profile-aware.
- `OUTLOOK_TOKEN` is still supported, but if a profile is bound it must match that profile's mailbox.

```bash
outlook login              # Interactive browser login
outlook login --force      # Force re-login, ignore saved session
outlook whoami             # Verify current user
outlook account add work   # Create and bind a named profile
outlook account list
outlook account current
outlook account switch work
outlook whoami --account work
```

### Account Selection

- Every non-account command accepts `--account NAME`.
- Selection precedence is: `--account NAME` → `OUTLOOK_ACCOUNT` → persisted current account → implicit `default`.
- `outlook whoami` shows the active profile in both human and JSON output.

```bash
outlook inbox --account work
outlook calendar --account personal --days 3
outlook schedule-list --account work
```

## Command Reference

### Inbox

```bash
outlook inbox                                    # List inbox (shows unread/total count)
outlook inbox --max 50                           # Limit count
outlook inbox --unread                           # Unread only
outlook inbox --from "alice.smith"                # Filter by sender
outlook inbox --subject "Q4 Report"              # Filter by subject
outlook inbox --from "acme" --subject "Project"  # Combined filters
outlook inbox --after 2026-03-01                 # After date
outlook inbox --before 2026-03-08               # Before date
outlook inbox --has-attachments                  # Only with attachments
outlook inbox --unread --after 2026-03-09        # Combine any filters
outlook inbox --json                             # JSON output
outlook inbox --json -o emails.json              # Save to file
```

### Read Email

```bash
outlook read 3             # Read message by display number
outlook read 3 --raw       # Show raw HTML body
outlook read 3 --json      # JSON output
```

### Conversation Thread

```bash
outlook thread 3           # Show full conversation for message #3
outlook thread 3 --json    # JSON output
```

### Send / Reply / Forward

```bash
outlook send "to@email.com" "Subject" "Body"                     # Shows confirmation prompt
outlook send "to@email.com" "Subject" "Body" -y                  # Skip confirmation
outlook send "a@b.com,c@d.com" "Subject" "Body" --cc e@f.com
outlook send "to@email.com" "Subject" "<h1>Hi</h1>" --html
outlook send "to@email.com" "Subject" "Body" --signature default  # Append saved signature
outlook send "to@email.com" "Report" "See attached" -a report.pdf         # With attachment
outlook send "to@email.com" "Files" "Here" -a file1.pdf -a file2.xlsx     # Multiple attachments

outlook reply 3 "Thanks!"                       # Shows confirmation prompt
outlook reply 3 "Thanks!" -y                    # Skip confirmation
outlook reply 3 "Noted, will fix." --all         # Reply all
outlook reply 3 "Here it is" -a requested.pdf    # Reply with attachment
outlook reply-draft 3                            # Create reply draft (empty body, edit in Outlook)
outlook reply-draft 3 "Will review tomorrow"     # Create reply draft with body
outlook reply-draft 3 "<p>HTML reply</p>" --html # HTML body (preserves quoted original)
outlook reply-draft 3 "Noted" --all              # Reply-all draft
outlook reply-draft 3 "Body" --signature default # Reply draft with signature
outlook reply-draft 3 --json                     # JSON output

outlook forward 3 "to@email.com"                # Shows confirmation prompt
outlook forward 3 "to@email.com" -y              # Skip confirmation
outlook forward 3 "to@email.com" --comment "FYI"
outlook forward 3 "to@email.com" -a extra.pdf    # Forward with additional attachment
```

### Drafts

```bash
outlook draft "to@email.com" "Subject" "Body"                    # Create draft
outlook draft "a@b.com,c@d.com" "Subject" "Body" --cc e@f.com   # Draft with CC
outlook draft "to@email.com" "Subject" "<h1>Hi</h1>" --html     # HTML draft
outlook draft "to@email.com" "Subject" "Body" -a doc.pdf         # Draft with attachment
outlook draft "to@email.com" "Subject" "Body" --signature default # Draft with signature
outlook draft "to@email.com" "Subject" "Body" --json             # JSON output
outlook draft-send 3                                              # Send draft (shows confirmation)
outlook draft-send 3 -y                                           # Send draft, skip confirmation
```

### Search

```bash
outlook search "keyword"
outlook search "from:alice acme" --max 10
outlook search "subject:Q4 Report" --json
outlook search "keyword" --json -o results.json
```

### Folders

```bash
outlook folders                                  # List all folders with counts
outlook folders --json -o folders.json           # Export to file
outlook folder "Archive" --max 20                 # Messages in a folder
outlook folder "Sent Items" --from "john" --max 10
```

### Categories

```bash
outlook categories                               # List categories with unread/total counts
outlook categories --json                        # JSON output
outlook categorize 3 "FYI"                       # Add category to message
outlook categorize 1 2 3 "FYI"                   # Add category to multiple messages
outlook uncategorize 3 "FYI"                     # Remove category from message
outlook uncategorize 1 2 3 "FYI"                 # Remove from multiple messages
outlook category-create "New Category"           # Create master category
outlook category-create "Urgent" --color 0       # Create with color (0=red, 7=blue, etc.)
outlook category-rename "FYI" "Info"             # Rename + update all messages
outlook category-rename "FYI" "Info" --no-propagate  # Master list only
outlook category-clear "FYI"                     # Remove label from all messages
outlook category-clear "FYI" --folder "Inbox"    # Limit to a folder
outlook category-clear "FYI" --max 50            # Limit to N messages
outlook category-clear "FYI" -y                  # Skip confirmation
outlook category-delete "Old Category"           # Delete (with confirmation)
outlook category-delete "Old Category" -y        # Delete without confirmation
```

### Signatures

```bash
outlook signature-pull                       # Extract signature from recent sent email
outlook signature-pull --name work           # Save with custom name
outlook signature-list                       # List saved signatures
outlook signature-show default               # Preview a signature
outlook signature-delete old-sig             # Delete a signature
outlook signature-delete old-sig -y          # Delete without confirmation
```

Signatures are stored per profile in `~/.config/outlook-cli/accounts/<profile>/signatures/`.

### Scheduled Send

```bash
outlook schedule "to@email.com" "Subject" "Body" "+1h"              # Schedule 1 hour from now
outlook schedule "to@email.com" "Subject" "Body" "+30m" -y          # Schedule 30 min, skip confirm
outlook schedule "to@email.com" "Subject" "Body" "tomorrow 09:00"   # Schedule for tomorrow
outlook schedule "to@email.com" "Subject" "Body" "2026-03-15T10:00" # Exact datetime
outlook schedule "to@email.com" "Subject" "Body" "+2h30m"           # Relative offset
outlook schedule "to@email.com" "Subject" "Body" "+1h" --html       # HTML body
outlook schedule "to@email.com" "Subject" "Body" "+1h" -s default   # With signature
outlook schedule "to@email.com" "Report" "See attached" "+1h" -a report.pdf  # With attachment
outlook schedule "to@email.com" "Subject" "Body" "+1h" --json       # JSON output

outlook schedule-draft 3 "+1h"                                      # Schedule existing draft
outlook schedule-draft 3 "tomorrow 09:00" -y                        # Skip confirmation

outlook schedule-list                                                # List all scheduled emails
outlook schedule-list --json                                         # JSON output

outlook schedule-cancel 1                                            # Cancel + delete draft from server
outlook schedule-cancel 1 -y                                         # Skip confirmation
```

**Time formats:** `+30m`, `+1h`, `+2h30m` (relative), `today 17:00`, `tomorrow 09:00` (day-relative), `2026-03-15T10:00` or `2026-03-15 10:00` (absolute ISO).

**How it works:** `schedule-list` cross-references local tracking with Drafts folder to find matching drafts. `schedule-cancel` deletes the draft from server (preventing delivery) and removes local tracking. Status shows `draft` when a server match is found, `queued` when only locally tracked.

**Workflow: Schedule a reply draft:**
```bash
outlook reply-draft 3 "Will review tomorrow"   # Create reply draft
outlook folder Drafts -n 1                      # Find draft number
outlook schedule-draft 42 "tomorrow 09:00"      # Schedule the reply
outlook schedule-list                            # Verify it's scheduled
```

### Message Management

```bash
outlook mark-read 3                # Mark as read
outlook mark-read 3 --unread       # Mark as unread
outlook mark-read 1 2 3            # Mark multiple as read
outlook move 3 "Archive"            # Move to folder (accepts display name)
outlook move 1 2 3 "Archive"        # Move multiple messages
outlook delete 3                   # Delete (with confirmation)
outlook delete 1 2 3 -y            # Delete multiple without confirmation
outlook flag 3                     # Flag for follow-up
outlook flag 3 4 5                 # Flag multiple messages
outlook flag 3 --due tomorrow      # Flag with due date
outlook flag 3 --due 2026-03-20    # Flag with specific date
outlook flag 3 --due +3d           # Flag due in 3 days
outlook flag 3 --complete          # Mark flag as complete
outlook flag 3 --clear             # Remove flag
outlook pin 3                     # Pin to top of inbox
outlook pin 3 4 5                 # Pin multiple messages
outlook pin 3 --unpin             # Unpin message
```

### Attachments

```bash
outlook attachments 3              # List attachments
outlook attachments 3 -d           # Download all
outlook attachments 3 -d --save-to ~/Downloads
outlook attachments 3 --json
```

### Calendar

```bash
outlook calendar                                    # Next 7 days
outlook calendar --days 14                          # Next 14 days
outlook calendar --calendar "John Smith"             # View a shared/other calendar
outlook calendar --calendar "John" --days 5         # Partial name match works
outlook calendar --json -o events.json
```

### Events

```bash
outlook event 42                                    # View event details (attendees, recurrence, etc.)

# Create
outlook event-create "Meeting" "2026-03-16 10:00" "2026-03-16 11:00"
outlook event-create "Meeting" "tomorrow 14:00" "tomorrow 15:00" \
  -a john@example.com -a jane@example.com \
  -l "Room A" -b "Agenda: Q1 review" -y
outlook event-create "Standup" "tomorrow 09:00" "tomorrow 09:30" \
  --teams -a team@example.com                       # Teams online meeting

# Recurring events
outlook event-create "Weekly Sync" "2026-03-16 10:00" "2026-03-16 11:00" \
  --repeat weekly --repeat-count 8 -a team@example.com
outlook event-create "Daily Standup" "2026-03-16 09:00" "2026-03-16 09:15" \
  --repeat daily --repeat-until 2026-04-30
outlook event-create "Sprint Review" "2026-03-16 14:00" "2026-03-16 15:00" \
  --repeat weekly --repeat-days Monday,Wednesday --repeat-count 12
outlook event-create "Monthly Report" "2026-03-16 10:00" "2026-03-16 11:00" \
  --repeat monthly --repeat-count 6

# Update
outlook event-update 42 --subject "New Title"
outlook event-update 42 --start "2026-03-16 14:00" --end "2026-03-16 15:00"
outlook event-update 42 --location "Room B"
outlook event-update 42 --add-attendee new@example.com
outlook event-update 42 --remove-attendee old@example.com

# Delete
outlook event-delete 42                             # Delete single event/occurrence
outlook event-delete 42 --series                    # Delete entire recurring series
outlook event-delete 42 43 44 -y                    # Delete multiple

# Respond to meeting invitations
outlook event-respond 42 accept
outlook event-respond 42 decline --comment "Can't make it"
outlook event-respond 42 tentative --silent          # Don't notify organizer

# Recurring event instances
outlook event-instances 42                           # List all occurrences (90 days)
outlook event-instances 42 --days 180                # Look further ahead
```

**Time formats for events:** `+1h`, `+30m`, `+2h30m` (relative), `today 17:00`, `tomorrow 09:00` (day-relative), `2026-03-15T10:00` or `2026-03-15 10:00` (absolute ISO).

**Recurrence options:** `--repeat daily|weekly|monthly`, `--repeat-interval N` (default 1), `--repeat-count N` (number of occurrences), `--repeat-until YYYY-MM-DD`, `--repeat-days Monday,Wednesday` (for weekly).

### Calendars / Free-Busy / People

```bash
outlook calendars                                    # List all calendars (own + shared)
outlook calendars --json

outlook free-busy "john@example.com" tomorrow         # Find free slots
outlook free-busy "a@b.com,c@d.com" 2026-03-16 -d 30 # 30-min duration slots
outlook free-busy "team@example.com" today --start-hour 14 --end-hour 18

outlook people-search "john"                          # Find people for attendee autocomplete
outlook people-search "john" --max 5 --json
```

### Contacts

```bash
outlook contacts                   # List contacts
outlook contacts --max 100
outlook contacts --json -o contacts.json
```

## JSON / Scripting

**Auto-JSON on pipe:** When stdout is piped (not a terminal), all commands automatically output JSON — no `--json` flag needed.

```bash
outlook inbox | jq '.data[0].subject'           # auto-JSON when piped
outlook inbox --json                              # explicit JSON in terminal
outlook inbox --json -o emails.json               # save raw JSON to file
outlook search "keyword" | jq '.data | length'
outlook categories | jq '.data[].Category'
```

**Structured envelope:** All JSON output is wrapped in a standard envelope:

```json
{"ok": true, "schema_version": "1", "data": [...]}
```

Errors also return structured JSON (when in JSON mode):

```json
{"ok": false, "schema_version": "1", "error": {"code": "not_found", "message": "..."}}
```

Error codes: `session_expired`, `rate_limited`, `not_found`, `not_authenticated`, `unknown_error`.

### JSON Field Names

Email objects (`inbox`, `search`, `folder` with `--json`):

| Field | Type | Example |
|-------|------|---------|
| `id` | string | Outlook message ID |
| `display_num` | int | `3` |
| `subject` | string | `"Re: Meeting"` |
| `sender` | object | `{"name": "John", "address": "john@x.com"}` |
| `to` | list[object] | `[{"name": "Jane", "address": "jane@x.com"}]` |
| `cc` | list[object] | same as `to` |
| `received` | string | `"2026-03-09T14:30:00Z"` |
| `preview` | string | first ~255 chars of body |
| `body` | string | full body text |
| `body_type` | string | `"Text"` or `"HTML"` |
| `is_read` | bool | `true` |
| `has_attachments` | bool | `false` |
| `importance` | string | `"Normal"` |
| `conversation_id` | string | Outlook conversation ID |
| `categories` | list[string] | `["Spam", "FYI"]` |
| `flag_status` | string | `"notFlagged"`, `"flagged"`, `"complete"` |
| `flag_due` | string\|null | `"2026-03-20T23:59:59"` or `null` |
| `scheduled_send` | string\|null | `"2026-03-15T10:00:00Z"` or `null` |

Event objects (`calendar`, `event`, `event-instances` with `--json`):

| Field | Type | Example |
|-------|------|---------|
| `id` | string | Outlook event ID |
| `display_num` | int | `42` |
| `subject` | string | `"Weekly Sync"` |
| `start` | string | `"2026-03-16T10:00:00"` |
| `end` | string | `"2026-03-16T11:00:00"` |
| `location` | string | `"Room A"` |
| `organizer` | object | `{"name": "John", "address": "john@x.com"}` |
| `attendees` | list[object] | `[{"email": {...}, "type": "Required", "response": "Accepted"}]` |
| `is_all_day` | bool | `false` |
| `show_as` | string | `"Busy"` |
| `response_status` | string | `"Accepted"`, `"NotResponded"`, `"Organizer"` |
| `recurrence` | object\|null | `{"Pattern": {...}, "Range": {...}}` |
| `event_type` | string | `"SingleInstance"`, `"Occurrence"`, `"SeriesMaster"` |
| `is_online_meeting` | bool | `true` |
| `online_meeting_url` | string | Teams join URL |

Folder objects (`folders --json`):
`name` (string), `unread_count` (int), `total_count` (int)

Category objects (`categories --json`):
`Category` (string), `Color` (string), `Unread` (int), `Total` (int)

## Common Patterns for AI Agents

```bash
# Quick inbox check with unread count
outlook inbox --max 10

# Find emails from a specific person
outlook inbox --from "bob.wilson"

# Find emails by subject
outlook inbox --subject "deployment" --unread

# Read the latest email from someone
outlook inbox --from "alice" --max 1 --json

# Check unread count without fetching emails
outlook folders --json | jq '.[] | select(.name == "Inbox") | .unread_count'

# Search across all folders
outlook search "deployment failed" --max 5

# Today's calendar
outlook calendar --days 1

# Monday's meetings on a shared calendar
outlook calendar --calendar "John Smith" --days 5

# Create a recurring weekly standup
outlook event-create "Standup" "2026-03-16 09:00" "2026-03-16 09:15" \
  --repeat weekly --repeat-count 8 -a team@example.com -y

# Check someone's availability for a meeting
outlook free-busy "colleague@company.com" tomorrow -d 30

# Find someone's email for invitation
outlook people-search "john"

# Accept a meeting invitation
outlook event-respond 42 accept

# View event with attendees
outlook event 42

# View full conversation thread
outlook thread 3

# Send a quick reply
outlook reply 3 "Received, will review today."

# Send email with attachments
outlook send "to@email.com" "Report" "See attached" -a report.pdf -a data.xlsx -y

# Download all attachments from a message
outlook attachments 5 -d --save-to ~/Downloads

# Flag message for follow-up with due date
outlook flag 3 --due tomorrow

# Pin important messages to top of inbox
outlook pin 3 4

# Categorize a batch of messages
outlook categorize 1 2 3 "FYI"

# List categories to see what's available
outlook categories

# Create a new category for a project
outlook category-create "Project Alpha" --color 7

# Schedule an email for later
outlook schedule "to@email.com" "Meeting notes" "Attached." "+1h"

# Schedule a reply for tomorrow morning
outlook reply-draft 3 "Will review this"
outlook schedule-draft 42 "tomorrow 09:00"

# Check scheduled emails
outlook schedule-list
```

## ID System

Messages and events get short display numbers (#1, #2, #3...) mapped to real Outlook IDs. Numbers are assigned when listing and persist across commands. Messages and events share the same ID map (capped at 500 entries).
The ID map is profile-local, so `#3` in one account is unrelated to `#3` in another.

```bash
outlook inbox --max 5      # Shows #1-#5
outlook read 3             # Read message #3
outlook reply 3 "OK"       # Reply to #3
outlook calendar --days 7  # Shows event #42, #43...
outlook event 42           # View event details
outlook event-respond 42 accept
```

## Error Handling

- Token expired → auto re-login attempted via cached SSO state.
- `Account profile 'X' not found` → run `outlook account add X` first, or use `outlook account list`.
- `Unknown message #N` → run `outlook inbox` or `outlook calendar` first to populate the ID map.
- `Folder 'X' not found` → run `outlook folders` to see available folder names.
- `Calendar 'X' not found` → run `outlook calendars` to see available calendar names.
- `Category 'X' not found` → run `outlook categories` to see available categories.
- HTTP 429 → automatic exponential backoff (3 retries).

## Safety Notes

- Token is cached with `chmod 600` (owner-only read/write).
- Browser state saved for SSO — avoids repeated logins.
- Tokens, browser state, signatures, scheduled-send tracking, and ID maps are scoped per account profile.
- `send`, `reply`, `forward`, `draft-send`, `schedule`, `schedule-draft`, `event-create`, `event-delete`, `delete`, and `category-delete` ask for confirmation by default (use `-y` to skip).
- `flag`, `pin`, `mark-read`, `categorize` do NOT require confirmation (safe, reversible operations).
- Do not share or log bearer tokens — they grant full mailbox access.
- Prefer `outlook login` over manually copying tokens.
