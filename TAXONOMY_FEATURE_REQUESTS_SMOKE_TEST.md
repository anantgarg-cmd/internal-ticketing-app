# Taxonomy, All Tickets and Feature Requests smoke test

Release: `taxonomy-feature-requests-all-tickets-v1`

## Before running

1. Back up the spreadsheet and record row counts and administrator-edited Category, ClientSizePriority and ROOT_CAUSES values.
2. As an ADMIN, run `upgradeTaxonomyAndFeatureRequests()` once, record its summary, then run it again and confirm no duplicate rows or columns are added.
3. Do **not** run `setupSystem()` on an existing production spreadsheet.

## Role and ticket regression

- SALES: default Raise Ticket; navigation has Raise Ticket, All Tickets, SLA Dashboard and Feature Requests; direct Work Queue and ticket mutation calls are rejected.
- POC/ADMIN: default Work Queue; all five pages are available; investigate, resolve and reopen a test ticket.
- Raise both 360 and Regular tickets with evidence. Verify duplicate warning/Raise Anyway and repeated Submission_Request_ID behavior.
- Verify Shopify and WooCommerce exist only for 360. Change client type/category and confirm stale subcategory, dynamic fields and evidence clear.
- Verify 360 Priority/SLA sources are CLIENT_SIZE and Regular sources are CATEGORY; verify working-hours due dates and reopen cycles.
- Resolve with each configured RCA; verify Other needs an explanation and Resolution Note remains mandatory.
- Search All Tickets by ticket/client/category/subcategory/subject/raiser/priority/status/SLA. Verify inactive historical raisers and legacy Category_Name display.
- Confirm Sales can reopen only a resolved ticket they raised; POC/ADMIN retain existing reopen rights.

## Feature Requests

- Submit with every mandatory field, valid optional HTTP/HTTPS link and attachment; repeat Submission_Request_ID and confirm one request.
- Search and open details as SALES; confirm status, Product response, attachment and timeline are visible and management controls are absent.
- As POC/ADMIN, update status, priority, owner, response and timeline; repeat Request_ID and confirm one event.
- Attempt invalid transitions and non-HTTP links and confirm clear errors.
- Confirm no Ticket, TicketSLACycles, TicketIndex or SlackNotifications row is created by a feature request.

## Data safety and live status

- Compare pre/post row counts and configuration snapshots: no historical ticket, Category_ID, existing column or administrator-edited setting may be removed or overwritten.
- Deployed Apps Script tests: **NOT RUN** in this repository-only implementation.
- Production deployment verification: **NOT RUN**; deployment is performed by the unchanged GitHub workflow after the main commit is pushed.
