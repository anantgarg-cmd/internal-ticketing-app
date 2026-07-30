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

## SLA-cycle schema recovery

Existing DEV and Production spreadsheets may predate ticket reopening and therefore may not contain the `TicketSLACycles` tab. Deploying Apps Script source does not migrate spreadsheet tabs. The application now repairs this additive schema automatically: `getInitialAppState()` checks the `APP_SCHEMA_VERSION` Script Property for `sla-cycles-v1`; only when the marker is absent or stale does it take a Script Lock, create or validate the tab, append missing columns, and then write the marker. It never invokes `setupSystem()`, creates attachment folders, seeds users, clears rows, reorders columns, or rebuilds source sheets.

The exact `TicketSLACycles` columns are:

`SLA_Cycle_ID`, `Ticket_ID`, `Cycle_Number`, `Cycle_Type`, `Started_At`, `Due_At`, `Ended_At`, `SLA_Result`, `Started_By`, `Ended_By`, `Reopen_Reason`, `Created_At`, `Updated_At`.

`Cycle_Type` is `INITIAL` or `REOPEN`; `SLA_Result` is `OPEN`, `MET`, or `BREACHED`. New and reopened cycles continue to calculate their due date inside Monday-Friday, 11:30-19:30 Asia/Kolkata working hours. A legacy ticket gets a single initial cycle when it is first resolved or reopened. Read-only ticket details degrade to an empty SLA history if recovery is temporarily unavailable; ticket lists do not depend on that history. Writes fail with an administrator-facing upgrade instruction rather than omitting SLA history.

`upgradeSlaCycleSchema()` remains available as a safe, repeatable diagnostic/repair function, but it is not normally a required deployment step. It appends only absent headers and preserves every existing row. `upgradeClientSizeAndPerformanceSchema()` includes the same additive validation and does not clear or rebuild any source sheet.

## Deployment and validation runbook

### DEV (`develop`)

1. Merge the pull request into `develop`; do not merge it into `main` yet.
2. In GitHub Actions, open **Deploy Apps Script** for the merge commit and verify every job step succeeds. The workflow selects `CLASP_JSON_DEV` and `DEPLOYMENT_ID_DEV`, pushes the tracked Apps Script files, creates a version, and updates the existing DEV deployment.
3. Open the existing DEV `/exec` URL as an authorized Shadowfax user. This first startup performs the locked schema recovery and writes `APP_SCHEMA_VERSION=sla-cycles-v1` only after success.
4. In the DEV spreadsheet, verify `TicketSLACycles` exists, row 1 is frozen, the headers above are present, and any pre-existing rows and column order are unchanged.
5. Validate **My Tickets** and **Shared Queue**, open a legacy ticket, raise one ticket, resolve it, reopen it with a reason, and resolve it again. Confirm cycle numbers increment, only the current cycle is `OPEN`, closed history remains unchanged, and working-hours due dates are correct.
6. If startup reports a schema preparation problem, run `upgradeSlaCycleSchema()` once in the DEV Apps Script editor and inspect its non-sensitive summary; then reload the DEV URL. This is fallback remediation, not a normal step.

### Production (`main`, after DEV approval)

1. After DEV validation, merge the already-reviewed `develop` changes into `main`; do not edit `main` directly.
2. Verify the **Deploy Apps Script** workflow for that `main` merge selects `CLASP_JSON_PROD` and `DEPLOYMENT_ID_PROD` and successfully updates the existing Production deployment.
3. Open the existing Production `/exec` URL as an authorized Shadowfax user. The first startup safely creates/validates `TicketSLACycles`; concurrent users are serialized by Script Lock, and later startups skip validation via the version marker.
4. Verify the Production tab, exact headers, frozen header row, timestamp formatting, and preservation of existing tickets/events/users/settings and any SLA-cycle rows.
5. Smoke-test **My Tickets**, **Shared Queue**, and one ticket detail. Then use an approved test ticket to validate initial, resolution, and reopen cycles without altering historical records.
6. Only if automatic recovery explicitly reports failure, run `upgradeSlaCycleSchema()` once in the Production Apps Script editor as the deployment owner, review its summary, and reload the web app.

No one-time manual sheet creation is required. Existing OAuth authorization requirements are unchanged; `authorizeApplication()` is needed only when Google requires deployment-owner consent, as described above.
