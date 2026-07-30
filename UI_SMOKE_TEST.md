# Production UI smoke test

Status: **NOT RUN** — this checklist requires a deployed Apps Script web app and live role accounts.

## ADMIN
- [ ] Open app; Work Queue is the default and its heading precedes data.
- [ ] Queue shows skeleton then table/cards or empty state; search, filters, and pagination work.
- [ ] SLA Dashboard shows metrics, including valid zero states.
- [ ] Tickets Raised by Me and Raise a Ticket load.
- [ ] Repeated switching among all pages never creates a blank page.

## POC
- [ ] Work Queue loads by default; Dashboard, personal tickets, and raise form load.

## SALES
- [ ] Raise a Ticket loads by default; Sales Tickets loads and search works.
- [ ] Work Queue and Dashboard are absent and cannot be navigated to.

## Ticket detail
- [ ] Open, close, and open another ticket; Investigate, Resolve, and Reopen work.
- [ ] Detail, duplicate, and reopen modals render useful failure states; loader always closes.

## Responsive
- [ ] Test 1440px, 1024px, 768px, and 390px.
- [ ] No clipped navigation text or horizontal overflow; header remains aligned.
- [ ] Desktop collapse and mobile navigation work.

## Failure and latency
- [ ] Queue backend failure shows page error; Retry succeeds.
- [ ] Dashboard and activity failures never make pages blank.
- [ ] Slow requests show the loader; timeout and late response close it without stale rendering.
