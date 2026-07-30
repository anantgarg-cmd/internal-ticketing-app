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

Apps Script deployments update source files but do not update an existing Google Sheet. Schema version **5** fixes that deployment gap generically. `getInitialAppState()` now calls `ensureRuntimeSchema_()` before users, settings, categories, client sizes, SLA cycles, or the ticket index can be read. The guard checks the durable `APP_SCHEMA_VERSION` property and all required sheet names; therefore a manually deleted tab is repaired even when the version property is current. A 300-second cache is used only as a hint, never as the source of truth.

The required tabs are `Tickets`, `Clients`, `Categories`, `Users`, `TicketEvents`, `Settings`, `TicketSLACycles`, `ClientSizePriority`, and `TicketIndex`. Migrations only create missing tabs, append missing headers, and seed absent Client Size codes. They never reorder headers, clear source rows, delete tabs, overwrite configuration edits, or change historical ticket priority. Timestamp formatting is resolved by header name so legacy column positions remain valid.

A short Script Lock serializes schema writes. The first opener repairs the installation; concurrent openers receive `SCHEMA_UPGRADE_IN_PROGRESS`, see **Updating the application structure…**, and retry once. Persistent failures show a controlled instruction to run `repairApplicationSchema()` instead of a raw missing-sheet or stack-trace error. The spreadsheet menu **Internal Ticketing Admin** also exposes safe repair, validation, and derived-index rebuild commands. These are editor/spreadsheet administration actions and are not exposed in the web UI.

`ClientSizePriority` seeds only absent `GOLD_PLATINUM`, `MEDIUM_SIZED`, and `SMALL_SIZED` codes. Administrator edits to existing codes are preserved. 360 ticket priority is resolved server-side from this configuration; Regular priority remains category-based.

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
