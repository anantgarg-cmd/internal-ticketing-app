# DEV live smoke and production gate

> All boxes are **NOT RUN** as of 2026-07-30. Run only in the DEV Apps Script project and DEV spreadsheet. Never run `setupSystem()` as part of this gate. Record ticket IDs only in the private test record, not in source control.

## Preconditions

- [ ] Confirm PR targets `develop`, CI passes, and `develop` contains everything newer in `main`.
- [ ] Confirm develop workflow selects `CLASP_JSON_DEV` and `DEPLOYMENT_ID_DEV`; confirm DEV script/deployment IDs belong together without printing either value.
- [ ] Confirm the deployment is DOMAIN / USER_DEPLOYING and deployment-owner authorization will be manual.

## Exact DEV validation

1. [ ] Merge/deploy the feature branch through `develop` to DEV; verify the existing DEV deployment URL is updated.
2. [ ] In the DEV editor run `authorizeApplication()` as deployment owner.
3. [ ] Run `repairApplicationSchema()`; confirm additive summary only and no existing configuration value changed.
4. [ ] Run `runEndToEndHealthCheck()`.
5. [ ] Confirm `ready` is true, every integrity count is zero, and no sensitive identifier/content is returned.
6. [ ] Test first-time SALES registration.
7. [ ] Test first-time POC registration.
8. [ ] Deactivate a DEV test user, refresh an old session, and confirm access is blocked.
9. [ ] Raise an existing 360 Gold / Platinum ticket with a PDF.
10. [ ] Confirm configured Priority, `CLIENT_SIZE` sources, SLA hours and due time.
11. [ ] Confirm attachment opens for the permitted test user.
12. [ ] Confirm exactly one Tickets row/request ID.
13. [ ] Confirm exactly one INITIAL OPEN SLA cycle.
14. [ ] Confirm exactly one TICKET_RAISED TicketEvents row/request ID.
15. [ ] Confirm exactly one TicketIndex row.
16. [ ] Confirm exactly one SlackNotifications row when priority/settings apply.
17. [ ] Raise a Medium 360 ticket.
18. [ ] Confirm the administrator-configured Medium priority/SLA was preserved and used.
19. [ ] Raise a Small 360 ticket and confirm configured values.
20. [ ] Raise a Regular ticket; confirm category priority/SLA and blank Client Size.
21. [ ] Test a category with optional attachment omitted.
22. [ ] Test required attachment missing; confirm no Drive or sheet write.
23. [ ] Trigger duplicate and choose Go back; confirm no write/request ID change.
24. [ ] Trigger duplicate and choose Raise anyway; confirm one ticket and stored duplicate fields.
25. [ ] Double-click ticket submit; confirm one ticket/file.
26. [ ] Retry the exact `Submission_Request_ID`; confirm the original ticket/file is returned.
27. [ ] As POC move Raised/Reopened to Investigating.
28. [ ] Resolve with configured root cause and note.
29. [ ] Reopen with reason.
30. [ ] Double-click reopen using the same action request ID; confirm one reopen event/cycle.
31. [ ] Confirm exactly one fresh OPEN REOPEN SLA cycle using stored SLA hours.
32. [ ] Confirm historical SLA cycle rows and values remain unchanged.
33. [ ] Directly attempt `getTicketDetail`/reopen as SALES for another user's ticket; confirm rejection.
34. [ ] Test Work Queue combined filters, >100 records, pagination clamping, rapid refresh and index missing/orphan diagnostics.
35. [ ] Test empty/populated SLA Dashboard and cache expiry.
36. [ ] Run `testSlackConnection()`; confirm one queue row and SENT result.
37. [ ] Run `validateSlackAutomation()`; confirm one dispatcher and one monitor.
38. [ ] Run `runTicketIntegrityDiagnostic()`; confirm all ten counts are zero.
39. [ ] Inspect DEV attachment folder for no orphan from deliberate failed/duplicate submissions.
40. [ ] Confirm no red error toast, stuck overlay, disabled form or raw internal error after success/failure paths.

## Concurrency/load extension

- [ ] Send 30 unique submissions and two concurrent calls sharing one request ID; verify unique IDs and one row for shared request.
- [ ] Send 40 simultaneous queue reads; record p50/p95 latency, lock errors and Apps Script quota errors.
- [ ] Race two investigate/resolve/reopen actions; verify one allowed transition/event per action request.
- [ ] Overlap schema repair, Slack monitor and two dispatchers; verify no duplicate Slack dedupe keys or partial corruption.

## Exact Production promotion

1. [ ] Attach DEV evidence for every launch-critical box; all P0/P1 counts remain zero.
2. [ ] Rebase/update the PR on latest `develop`; compare `develop...main` and reconcile newer main-only commits.
3. [ ] Merge the audit PR to `develop`; repeat DEV deployment and health check on the merge commit.
4. [ ] Open the normal promotion PR from `develop` to `main`; require review and green local/CI deployment-safety tests.
5. [ ] Confirm production secrets are `CLASP_JSON_PROD` and `DEPLOYMENT_ID_PROD`, and the deployment belongs to that script without exposing values.
6. [ ] Merge to `main`; verify workflow stamps the short commit and redeploys the existing Production deployment (no new URL).
7. [ ] Manually run `authorizeApplication()` only if scopes changed/authorization is required. Do **not** run `setupSystem()`.
8. [ ] Run read-only `runEndToEndHealthCheck()` first; if not ready, stop and roll back the deployment version without modifying data.
9. [ ] Perform a minimal authorized Production smoke with a designated test client only if organizational policy permits; otherwise rely on DEV evidence.
10. [ ] Confirm UI/backend release and short commit match, triggers are singular, monitoring is healthy, and preserve rollback version/evidence.

## Final production live smoke test (2026-07-31)

> **Deployment status:** Every item below is **NOT RUN**. These steps require the live Apps Script deployment, configured Sheets/Drive, real role accounts, and (where noted) Slack. Local static tests do not count as execution.

### SALES — NOT RUN

- [ ] **NOT RUN** — Raise a 360 ticket and confirm server-returned ClientSizePriority priority/SLA and `CLIENT_SIZE` sources.
- [ ] **NOT RUN** — Raise a Regular ticket and confirm server-returned Subcategory priority/SLA and `CATEGORY` sources.
- [ ] **NOT RUN** — Raise API, Shopify, and WooCommerce tickets; satisfy each practical evidence rule without entering credentials or secrets.
- [ ] **NOT RUN** — Upload each supported attachment type and use a safe HTTP/HTTPS supporting or video link separately (neither should require the other).
- [ ] **NOT RUN** — Trigger a duplicate warning, then use **Raise Anyway** once; verify one ticket/event despite repeat clicks or a retry.
- [ ] **NOT RUN** — View another user's ticket in All Tickets; confirm Work Queue and status actions are absent.
- [ ] **NOT RUN** — Reopen only a personally raised resolved ticket; confirm another Sales user's ticket is denied.
- [ ] **NOT RUN** — Submit a Feature Request using only Product Area, Feature Title, Business Case, and Expected Benefit.

### POC / ADMIN — NOT RUN

- [ ] **NOT RUN** — Open Work Queue, investigate a Raised/Reopened ticket, and resolve it with a configured RCA and resolution note.
- [ ] **NOT RUN** — Reopen a resolved ticket and verify the new independent SLA cycle and timeline event.
- [ ] **NOT RUN** — Manage a Feature Request through a valid transition with product priority, owner, response, and target timeline.
- [ ] **NOT RUN** — Validate dashboard counts/adherence and compare them with representative resolved/open tickets.
- [ ] **NOT RUN** — As ADMIN, run `runEndToEndHealthCheck()` and `validateSlackAutomation()`; then run `testSlackConnection()` only when an intentional Slack test message is acceptable.

### Regression — NOT RUN

- [ ] **NOT RUN** — Open a historical ticket and a legacy-category ticket; verify category fallback, dynamic fields, attachment, resolution, events, and SLA history.
- [ ] **NOT RUN** — Search before pagination in All Tickets and Work Queue; move between pages and open a result.
- [ ] **NOT RUN** — Open a valid attachment and verify a missing/deleted Drive file fails without blanking or trapping the page.
- [ ] **NOT RUN** — Test phone and desktop layouts, table overflow, drawer close/focus return, navigation, empty states, and Error with Retry.
- [ ] **NOT RUN** — Force backend failure and timeout paths on each page; confirm no blank page and no stuck local or global loader.
