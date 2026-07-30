# Frontend/backend contract audit

Release: `shadowfax-ui-stability-v2`. Production remains a single `getInitialAppState` startup followed by permission-gated, lazy reads. All failures and timeouts close the global loader; read failures replace only the dynamic host with a retryable page error. Cache keys include the frontend release and cached values pass through the same normalisers as network values.

| Method | Request and required fields | Return shape / frontend handling | Permission and states |
|---|---|---|---|
| `getInitialAppState` | none | `{state, release, bootstrap}`; REGISTER/BLOCKED/ACTIVE are explicit; 30s timeout is fatal | Domain/auth and active-user checks remain server-owned |
| `registerFirstTimeUser` | `{name, role}` | entry-state response; button restored on every outcome | SALES or POC registration rules unchanged |
| `checkDuplicate` | client/category/subject summary | `{hasDuplicate,matches[]}`; modal or submission | Server visibility rules determine viewability |
| `getSlaDuePreview` | `categoryId`, optional client-size code | `{dueAt}`; inline unavailable state on failure | Existing category/client-size SLA rules unchanged |
| `submitTicket` | real HTML form, file, `submissionRequestId` | ticket/detail object; caches invalidated | Attachments/idempotency/priority/SLA remain server-owned |
| `getMyTickets` | `{search,page,pageSize}` | paginated rows; malformed envelope errors, optional row fields default | SALES scope vs personal scope remains server-owned |
| `getQueueTickets` | search/status/priority/category/clientSize/SLA/page | paginated rows plus categories; skeleton/table/empty/error | POC/ADMIN only |
| `getTicketDetail` | ticket ID | detail, events, SLA cycles; drawer loading/error | Backend visibility is authoritative |
| `updateTicketStatus` | ticket/status/action ID; resolution fields when resolving | updated detail; loader always closes | Existing lifecycle and POC/ADMIN rules unchanged |
| `getReopenPreview` | ticket ID | `{dueAt}`; modal inline state | Backend ownership/reopen checks unchanged |
| `reopenTicket` | ticket ID, reason, action ID | updated detail/new SLA cycle | Existing ownership and lifecycle checks unchanged |
| `getNumbers` | none | dashboard counters and grouping arrays; missing optional groups become empty, invalid envelope errors | POC/ADMIN only |

## Reliability rules

Paginated envelopes require an object with a `rows` array; empty arrays are valid. Optional ticket text fields default to empty strings and pagination values are bounded. Dashboard envelopes must be objects; grouping arrays default to empty and malformed rows are ignored. Requests settle once, ignore late callbacks after timeout, run cleanup and `always` exactly once, and expose only normalised user-safe error text. No contract, Sheets read, Drive call, polling, or business rule was added.
