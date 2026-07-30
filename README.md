# Internal Ticketing App

This Google Apps Script web app is deployed for the Shadowfax Workspace domain. It runs as the deployment owner; regular users must never be asked to grant Spreadsheet, Drive, or external-request permissions.

## Deployment-owner OAuth authorization

GitHub Actions updates the existing DEV or Production deployment, but it cannot approve Google account permissions. After deploying either environment, its deployment owner must complete authorization once:

1. Open the corresponding Apps Script project while signed in as the account that owns or deployed that web app.
2. Select `authorizeApplication` from the function dropdown.
3. Click **Run**.
4. Click **Review permissions**.
5. Select the Shadowfax deployment-owner account.
6. Approve all required permissions.
7. Run `authorizeApplication` again.
8. Confirm it returns `authorized: true`.
9. Reopen that environment's `/exec` URL.

Repeat these steps separately for **DEV** (the Apps Script project/deployment configured for the `develop` branch) and **Production** (the project/deployment configured for `main`). Do not make either web app public; both deployments must remain domain-restricted and execute as the deploying user.

For a safe status check, run `getAuthorizationDiagnostic` manually in the editor. `REQUIRED` means the deployment owner still needs to complete the steps above; the diagnostic never returns the authorization URL.

### If Google's OAuth page is blank or broken

- Allow third-party cookies for `script.google.com`, `accounts.google.com`, and `googleusercontent.com`.
- Allow pop-ups and redirects for `script.google.com`.
- Retry in a clean Chrome profile signed in only with the intended Shadowfax deployment-owner account.

Do not change web-app access to public or ask regular users to authorize application services.

## Automatic, versioned spreadsheet schema

Apps Script deployments update source files but do not update an existing Google Sheet. Schema version **6** extends that deployment-safe repair path with the client-size SLA source columns. `getInitialAppState()` now calls `ensureRuntimeSchema_()` before users, settings, categories, client sizes, SLA cycles, or the ticket index can be read. The guard checks the durable `APP_SCHEMA_VERSION` property and all required sheet names; therefore a manually deleted tab is repaired even when the version property is current. A 300-second cache is used only as a hint, never as the source of truth.

The required tabs are `Tickets`, `Clients`, `Categories`, `Users`, `TicketEvents`, `Settings`, `TicketSLACycles`, `ClientSizePriority`, and `TicketIndex`. Migrations only create missing tabs, append missing headers, and seed absent Client Size codes. They never reorder headers, clear source rows, delete tabs, overwrite configuration edits, or change historical ticket priority. Timestamp formatting is resolved by header name so legacy column positions remain valid.

A short Script Lock serializes schema writes. The first opener repairs the installation; concurrent openers receive `SCHEMA_UPGRADE_IN_PROGRESS`, see **Updating the application structure…**, and retry once. Persistent failures show a controlled instruction to run `repairApplicationSchema()` instead of a raw missing-sheet or stack-trace error. The spreadsheet menu **Internal Ticketing Admin** also exposes safe repair, validation, and derived-index rebuild commands. These are editor/spreadsheet administration actions and are not exposed in the web UI.

`ClientSizePriority` seeds only absent `GOLD_PLATINUM` (HIGH/8 hours), `MEDIUM_SIZED` (MEDIUM/24 hours), and `SMALL_SIZED` (LOW/48 hours) codes. Administrator edits to existing codes are preserved. For 360 tickets, priority and SLA hours are resolved server-side from this configuration; for Regular tickets, both remain category-based.

When a new/empty `TicketIndex` is found, existing `Tickets` are read without mutation and copied with batch `setValues()`. Up to 1,000 rows are processed per continuation, progress is saved in Script Properties, and a temporary time trigger continues a large backfill. The schema version is committed only after the index and complete schema validate. `rebuildTicketIndex()` remains the manual admin operation that clears/rebuilds only this derived tab.

SLA history remains additive: new tickets create one `INITIAL` cycle; resolving closes the one open cycle; reopening preserves prior cycles and starts one `REOPEN` cycle from the reopen timestamp. Legacy tickets without cycle rows return an empty history on read and are derived only when a write needs a cycle. Working time remains Monday-Friday, 11:30-19:30, Asia/Kolkata.

## Deployment and validation runbook

### DEV (`develop`)

1. Merge the feature pull request into `develop` and let the existing workflow update the DEV Apps Script deployment.
2. Open the existing DEV `/exec` URL. The first request automatically upgrades the DEV spreadsheet; no tab creation or one-time schema function is required.
3. Run `validateApplicationSchema()` in the editor or use the spreadsheet admin menu if an audit is desired.
4. Smoke-test registration, role navigation, 360 and Regular ticket creation, pagination, resolve/reopen, attachments, Slack, and duplicate detection.

### Production (`main`, after DEV approval)

1. After DEV validation, merge `develop` into `main`; do not edit `main` directly.
2. Let the existing workflow update the Production deployment and open its `/exec` URL. The first request automatically upgrades the Production spreadsheet.
3. Confirm `validateApplicationSchema().schemaReady` and perform the approved smoke tests.

No administrator must manually create `ClientSizePriority`, `TicketSLACycles`, `TicketIndex`, or any base tab. The only unavoidable manual step remains deployment-owner OAuth consent through `authorizeApplication()` when Google reports that new or renewed scopes require it; regular users must not authorize the deployment.

### Capacity note

The lock, idempotency keys, caches, batched index, and pagination are designed for bursts of roughly 30–40 users. Google Apps Script and Sheets quotas still apply; monitor execution logs and archive operational data deliberately as volume grows.

## Asynchronous Slack notifications

Slack delivery is operational queue data: ticket creation and reopening commit the ticket, SLA cycle, event, and `TicketIndex` first, then append a deduplicated `PENDING` record to `SlackNotifications`. The user-facing request never calls Slack. `dispatchSlackNotifications` claims at most the configured batch every minute, releases `ScriptLock`, and then calls the Incoming Webhook. `monitorSlackAlerts` reads only `TicketIndex` every 15 minutes to enqueue working-time SLA warnings, breaches, and the weekday end-of-day summary. Failed Slack delivery never changes a ticket.

The queue columns are `Notification_ID`, `Dedupe_Key`, `Notification_Type`, `Ticket_ID`, `SLA_Cycle_Number`, `Priority`, `Payload_JSON`, `Status`, `Attempts`, `Next_Attempt_At`, `Created_At`, `Processing_Started_At`, `Sent_At`, `Last_HTTP_Code`, `Last_Error`, and `Updated_At`. Sent and permanently failed operational rows older than the configured retention may be removed; ticket, event, SLA-cycle, and index history is never removed.

Durable keys are `NEW_HIGH:<Ticket_ID>`, `REOPEN_HIGH:<Ticket_ID>:<SLA_Cycle_Number>`, `SLA_WARNING:<Ticket_ID>:<SLA_Cycle_Number>`, `SLA_BREACHED:<Ticket_ID>:<SLA_Cycle_Number>`, `EOD:<yyyy-MM-dd>`, and `TEST:<UUID>`. Temporary network errors, HTTP 429, and HTTP 5xx use 1-, 5-, and 15-minute retry delays (a longer Slack `Retry-After` wins). Other HTTP 4xx responses and exhausted retries are retained as `FAILED`.

### One-time setup for DEV

1. Create a Slack app, enable Incoming Webhooks, and add a webhook to a dedicated **test Slack channel**.
2. Open the DEV Apps Script project (DEV and Production are separate projects).
3. In **Project Settings → Script Properties**, add `SLACK_WEBHOOK_URL` with the webhook URL. Never put it in source or a Sheet.
4. Deploy the DEV code.
5. From the Apps Script editor, run `authorizeApplication()` and accept the deployment-owner scopes.
6. Run `repairApplicationSchema()`; it creates/repairs `SlackNotifications` and seeds absent Slack Settings without overwriting edits.
7. Run `setupSlackAutomationTriggers()`.
8. Run `testSlackConnection()` and confirm the test-channel message.
9. Run `validateSlackAutomation()` and resolve all warnings.
10. Configure the `SLACK_*` rows in the Settings sheet as needed (the existing Settings cache can take up to 300 seconds to refresh).

### One-time setup for Production

Repeat all ten DEV steps in the **Production Apps Script project**, using a webhook for the production internal-ticket channel. Script Properties, authorization, deployments, and installable triggers are not shared between DEV and Production. Installable triggers run as the account that creates them, so the Production deployment owner should install them. Trigger failures and authorization problems are investigated under **Apps Script → Executions**. Do not copy a DEV webhook into Production.

The spreadsheet **Internal Ticketing Admin** menu provides safe actions to validate automation, install/repair triggers, enqueue and dispatch a test, and remove only the two Slack triggers. Diagnostics never return the webhook, trigger IDs, spreadsheet ID, tickets, or users.

### Slack Settings

The schema repair seeds these absent keys while preserving every administrator-configured value: `SLACK_NOTIFICATIONS_ENABLED`, `SLACK_ALERT_PRIORITIES`, `SLACK_MENTION_PRIORITIES`, `SLACK_BREACH_WARNING_ENABLED`, `SLACK_BREACH_WARNING_MINUTES`, `SLACK_BREACH_ALERT_ENABLED`, `SLACK_EOD_SUMMARY_ENABLED`, `SLACK_EOD_HOUR`, `SLACK_EOD_MINUTE`, `SLACK_DISPATCH_BATCH_SIZE`, `SLACK_MAX_RETRIES`, `SLACK_PROCESSING_TIMEOUT_MINUTES`, and `SLACK_NOTIFICATION_RETENTION_DAYS`.
