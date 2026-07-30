# Final adversarial end-to-end reliability audit

**Audit date:** 2026-07-30  
**Target:** `anantgarg-cmd/internal-ticketing-app` → `develop`  
**Live DEV execution:** **NOT RUN**. This checkout has no usable GitHub tunnel, `.clasp.json`, DEV credentials, deployment ID, or access to the DEV spreadsheet. No production or DEV data was read or changed.

## Executive assessment

The repository is a domain-restricted Apps Script V8 web application with one HTML page, two production JavaScript files, ten Sheets tables, a Drive attachment store, an asynchronous durable Slack queue, Script Properties, Script Cache and time-based triggers. The local suite passes. Four defects were fixed: stale derived index rows could hide authoritative tickets (P1), Drive deletion could occur while ScriptLock was held (P1), the health check omitted required cross-sheet diagnostics (P1), and formula-prefixed user text/root-cause manipulation was not rejected safely (P1). No known P0/P1 remains in locally executable logic. Production promotion remains blocked on all live DEV checks.

## Branch safety and comparison

The provided checkout began on `work` at `ad170a1`. It contained no local `main`/`develop` refs and no configured remote. A read-only fetch from GitHub was attempted before coding but the environment's CONNECT tunnel returned HTTP 403. The current commit is the newest supplied merge snapshot and includes PRs 14–20. Work was isolated on `audit/final-adversarial-reliability`; `main` was not modified. A maintainer must verify `origin/develop...origin/main` again when network access is available.

## Internal inventory

### Frontend and contracts

* Page: `Index.html`, with entry/registration, Raise Ticket, My Tickets, Work Queue, SLA Dashboard, ticket detail, duplicate confirmation, resolve and reopen views.
* Forms: `registrationForm`, `ticketForm`; modal actions construct explicit status/reopen payloads.
* Server calls: `getInitialAppState()`, `registerFirstTimeUser(payload)`, `getSlaDuePreview(categoryId, clientSizeCode)`, `checkDuplicate(payload)`, `submitTicket(form)`, `getMyTickets(options)`, `getQueueTickets(filters)`, `getTicketDetail(ticketId)`, `getReopenPreview(ticketId)`, `updateTicketStatus(payload)`, `reopenTicket(payload)`, `getNumbers()`.
* Submission fields: `clientMode`, `clientId`, `clientName`, `clientType`, `clientSize`, `categoryId`, `emailSubject`, `issueDescription`, category-defined controls, `attachment`, `duplicateOverride`, `duplicateIds`, `submissionRequestId`.
* Loader/button cleanup is paired on success and failure. `IN_FLIGHT` prevents repeated submit in one tab; durable request IDs provide server idempotency across tabs/timeouts.

### Apps Script surface

Public/editor or trigger entry points are `doGet`, `authorizeApplication`, `getAuthorizationDiagnostic`, `getBootstrap`, `checkDuplicate`, `getSlaDuePreview`, `getReopenPreview`, `setWebAppUrlFromEditor`, `getEntryState`, `getRuntimeDiagnostics`, `registerFirstTimeUser`, `backfillTicketSlaCycles`, `recalculateOpenTicketSlaDueDates`, `diagnoseWorkingSlaCalculation`, `getInitialAppState`, `rebuildTicketIndex`, `getMyTickets`, `getQueueTickets`, `getTicketDetail`, `getNumbers`, `submitTicket`, `updateTicketStatus`, `reopenTicket`, `runTicketIntegrityDiagnostic`, `runEndToEndHealthCheck`, `testTicketSubmissionPayload`, `monitorSlackAlerts`, `dispatchSlackNotifications`, `setupSlackAutomationTriggers`, `removeSlackAutomationTriggers`, `validateSlackAutomation`, `testSlackConnection`, `setupSystem`, `validateApplicationSchema`, `repairApplicationSchema`, `continueTicketIndexBackfill`, `onOpen`, and compatibility upgrade functions.

### Data/configuration

* Sheets: Tickets, Clients, Categories, Users, TicketEvents, Settings, TicketSLACycles, ClientSizePriority, TicketIndex, SlackNotifications. Canonical headers are in `APP.HEADERS`.
* Script Properties: `SPREADSHEET_ID`, `ATTACHMENT_FOLDER_ID`, `APP_SCHEMA_VERSION`, `SCHEMA_UPGRADE_IN_PROGRESS`, `TICKET_INDEX_BACKFILL_NEXT_ROW`, `WEB_APP_URL`, `SLACK_WEBHOOK_URL`, `SLACK_CLEANUP_DATE`; deployment stamping uses `APP_COMMIT` in source. Secret values are neither returned by the health check nor committed.
* Cache keys: `app:settings:v2`, `app:categories:v2`, `app:client-size-priority:v1`, `app:numbers:v2`, `app:schema-ready:v7`.
* Triggers: `continueTicketIndexBackfill` (one-shot), `dispatchSlackNotifications` (1 minute), `monitorSlackAlerts` (15 minutes). Trigger installation remains manual/editor-controlled.
* Roles: SALES, POC, ADMIN. Self-registration accepts only SALES/POC.
* Statuses/transitions: Raised → Investigating; Reopened → Investigating; Investigating → Resolved; Resolved → Reopened. Every other transition is rejected server-side.
* Idempotency: Tickets `Submission_Request_ID`; TicketEvents `Request_ID` / action request ID; Slack `Dedupe_Key`; SLA reopen cycles are guarded by action event lookup and open-cycle checks.
* ScriptLock: registration; ID allocation; ticket commit; lifecycle/reopen commits; schema migration/repair/backfill; Slack enqueue/claim/row update/cleanup. Drive upload and webhook requests occur outside ticket/dispatcher locks.
* Drive: manual authorization, one-time setup folder creation, attachment folder lookup/upload, add viewer, and uncommitted-or-duplicate file cleanup.
* Slack: durable enqueue, monitoring, EOD enqueue, dispatcher claim, webhook delivery, retries, cleanup, validation and test enqueue.

## Architecture and high-risk flows

Tickets is authoritative. TicketIndex is derived. Ticket creation uploads a validated attachment before the short commit lock, then rechecks the submission request ID under lock and appends Tickets, INITIAL cycle, event and index. Network Slack delivery is always asynchronous. A failure after the Tickets append intentionally retains the attachment and is exposed by the count-only integrity diagnostic; retrying the request ID returns the authoritative ticket rather than creating another.

Schema repair is versioned/additive: missing sheets/columns and missing seed keys are added, existing values/rows are preserved, and setup is not invoked by normal startup. `setupSystem()` remains an explicit one-time editor operation and was not run.

## Defects

| ID | Severity | Reproduction | Root cause | Fix / evidence | Status |
|---|---|---|---|---|---|
| AUD-001 | P1 | Remove a TicketIndex row, then open My Tickets/Queue. | Both lists treated derived TicketIndex as authoritative. | Lists now read Tickets; index remains optimization/diagnostic data. | Fixed, local regression pass |
| AUD-002 | P1 | Race identical submission IDs after both uploads. | Duplicate loser deleted Drive file before releasing ScriptLock; precommit cleanup in catch also ran under lock. | Lock records cleanup decision only; Drive trash call runs after release; committed attachments are retained. | Fixed, local regression pass |
| AUD-003 | P1 | Run health check with duplicate IDs/orphan index/open-cycle corruption. | Health check compared only row totals and omitted duplicate headers/cross-sheet invariants. | Added pure ten-invariant diagnostic, ADMIN wrapper, safe counts and readiness gating. | Fixed, local regression pass |
| AUD-004 | P1 | Submit `=...` user text or an arbitrary nonempty root cause. | `setValues` can interpret formula prefixes; resolution accepted unconfigured root causes. | Prefix formula markers as sheet text and require configured root cause. | Fixed, local regression pass |

P0: 0 found / 0 fixed / 0 open. P1: 4 / 4 / 0. P2: 0 / 0 / 0. P3: 0 / 0 / 0. Live-only findings cannot be ruled out before DEV execution.

## Area conclusions

* **Functions/contracts:** no duplicate top-level production declarations. All 11 frontend server contracts exist; required submit controls stay enabled through dispatch.
* **Data integrity/concurrency:** request/action/dedupe IDs are mandatory; idempotency is rechecked under ScriptLock. Multi-system Sheets writes cannot be truly transactional in Apps Script, so partial post-Tickets writes are detected rather than destructively rolled back.
* **Attachments:** upload is outside lock; only definitely uncommitted/duplicate uploads are cleanup candidates; committed ticket attachments are never trashed by submit error handling.
* **Lifecycle/permissions:** ticket read ownership and POC/ADMIN work operations are server enforced. Reopen ownership is checked, reason/action ID are mandatory, old cycles/events are retained.
* **SLA:** Asia/Kolkata, weekdays 11:30–19:30, weekend rollover and fractional hours are locally covered. Holidays are intentionally unsupported.
* **Schema:** additive repair preserves rows and administrator-edited values. `rebuildTicketIndex` is explicit/admin-only and destructive only to derived TicketIndex, not source sheets.
* **Slack:** webhook calls occur only in dispatcher/test dispatch, outside claim lock; durable dedupe, stale-processing reclaim, 429/5xx/network retries and redacted errors are present.
* **Deployment:** branch-to-secret mapping is develop→DEV and main→Production; workflow updates an existing deployment, validates selected files and stamps commits. Manifest remains DOMAIN / USER_DEPLOYING.

## Remaining limitations and risks

1. **LIVE-GATE (launch blocker):** authorization, Sheets/Drive scopes, actual trigger ownership, webhook delivery, 30/40-user load, real browser file serialization, actual Apps Script quotas and DEV/production project mappings cannot be established locally. Follow `DEV_SMOKE_TEST.md`.
2. Apps Script/Sheets do not offer a cross-sheet transaction. The authoritative ticket can survive while its cycle/event/index append fails. This is safer than deleting a committed ticket/file and is now explicitly diagnosed; administrator repair/backfill is the workaround.
3. Public holidays are not excluded from SLA calculations.
4. Ticket list authority now favors correctness over TicketIndex speed. Validate 40 concurrent reads and quota/latency in DEV with production-scale rows.
5. This environment prevented branch comparison and PR API access via a 403 CONNECT tunnel. Re-run comparison and PR creation from a network-enabled environment.

## Rollback considerations

Revert the audit commit to restore code. No data migration or setup is required. The new diagnostic is read-only. Formula protection only affects new user-authored values. Do not roll back by clearing sheets. If list performance is unacceptable, retain authoritative fallback semantics rather than reverting to index-only reads. Never delete attachments based only on a frontend timeout.

## Safety confirmations

No historical data was accessed or deleted. No source sheet was cleared/recreated. No administrator configuration was changed or seeded during this audit. `setupSystem()` and all live mutation functions were not executed. No secret-like webhook, spreadsheet ID, folder ID, OAuth token, clasp credential, or deployment credential was committed.
