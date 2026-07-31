const assert = require('assert');
const fs = require('fs');
const vm = require('vm');
const code = fs.readFileSync('Code.js', 'utf8');
const setup = fs.readFileSync('Setup.js', 'utf8');
const html = fs.readFileSync('Index.html', 'utf8');
const workflow = fs.readFileSync('.github/workflows/deploy-apps-script.yml', 'utf8');
const ignore = fs.readFileSync('.claspignore', 'utf8');

// Evaluate the production source itself; Apps Script services are only resolved
// when service-facing functions execute, so pure reliability helpers are testable.
const context = { console, Date, JSON, Math, Object, String, Number, Boolean, Array, RegExp, Set, Map };
vm.createContext(context);
vm.runInContext(setup + '\n' + code, context);

assert.strictEqual(context.safeSheetText_('=IMPORTXML("x")', 100), "'=IMPORTXML(\"x\")");
assert.strictEqual(context.safeSheetText_('+1', 100), "'+1");
assert.strictEqual(context.safeSheetText_('ordinary', 100), 'ordinary');

const tickets = [
  {Ticket_ID:'T1',Submission_Request_ID:'R1',Status:'Raised'},
  {Ticket_ID:'T2',Submission_Request_ID:'R1',Status:'Resolved'},
  {Ticket_ID:'T2',Submission_Request_ID:'R3',Status:'Raised'}
];
const index = [{Ticket_ID:'T1'},{Ticket_ID:'ORPHAN'}];
const cycles = [
  {Ticket_ID:'T1',Cycle_Type:'INITIAL',SLA_Result:'OPEN'},
  {Ticket_ID:'T1',Cycle_Type:'REOPEN',SLA_Result:'OPEN'},
  {Ticket_ID:'T2',Cycle_Type:'INITIAL',SLA_Result:'OPEN'}
];
const events = [{Request_ID:'A1'},{Request_ID:'A1'}];
const slack = [{Dedupe_Key:'D1'},{Dedupe_Key:'D1'}];
const result = context.diagnoseTicketIntegrityRows_(tickets,index,cycles,events,slack);
assert.deepStrictEqual(JSON.parse(JSON.stringify(result)), {
  ticketsWithoutIndex:2,indexRowsWithoutTicket:1,ticketsWithoutInitialSlaCycle:0,ticketsWithMultipleOpenSlaCycles:1,
  duplicateSubmissionRequestIds:1,duplicateTicketIds:1,duplicateEventRequestIds:1,duplicateSlackDedupeKeys:1,
  resolvedTicketsWithOpenSlaCycle:1,openTicketsWithoutOpenSlaCycle:0
});

// Frontend contract: required controls stay enabled and the same form is sent.
for (const field of ['clientMode','clientId','clientName','clientType','clientSize','categoryId','emailSubject','issueDescription','duplicateOverride','duplicateIds','submissionRequestId']) {
  assert(new RegExp(`name=["']${field}["']`).test(html), `missing serializable control ${field}`);
}
assert(html.includes("method:'submitTicket',payload:form"));
assert(!/form\.elements[^\n]*disabled\s*=\s*true/.test(html));
assert(html.includes("const restore=()=>{form.removeAttribute('aria-busy')"));
assert(html.includes("always:restore"));
assert(html.includes("if(IN_FLIGHT.has(requestKey))return"));

// Authoritative list reads must not silently hide Tickets when TicketIndex is stale.
assert(/function getAllTickets[\s\S]*?getSheetObjects_\(APP\.SHEETS\.TICKET_INDEX\)/.test(code));
assert(/function getQueueTickets[\s\S]*?getSheetObjects_\(APP\.SHEETS\.TICKETS\)/.test(code));

// Network/Drive cleanup is after lock release in the ticket commit flow.
const submit = code.slice(code.indexOf('function submitTicket'), code.indexOf('function updateTicketStatus'));
assert(submit.indexOf('finally{lock.releaseLock();}') < submit.indexOf('DriveApp.getFileById(attachment.id).setTrashed(true)'));
assert(submit.indexOf('finally{lock.releaseLock();}') < submit.indexOf('sendSlackAlert_(detail)'));
assert(code.includes("Select a valid root cause from the configured list."));

// Deployment and secret safety.
for (const path of ['test','tests','node_modules','package.json','package-lock.json']) assert(ignore.includes(path));
assert(!/hooks\.slack\.com\/services\/[A-Za-z0-9/_-]+/.test(code + setup + html + workflow));
assert(workflow.includes('CLASP_JSON_DEV') && workflow.includes('DEPLOYMENT_ID_DEV'));
assert(workflow.includes('CLASP_JSON_PROD') && workflow.includes('DEPLOYMENT_ID_PROD'));
assert(workflow.includes('npx clasp redeploy "$DEPLOYMENT_ID"'));
assert(!/\brequire\s*\(/.test(code + setup));
assert(!/module\.exports/.test(code + setup));

console.log('reliability behavior, integrity, form, authority, lock-boundary, and deployment checks passed');
