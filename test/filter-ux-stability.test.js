const assert = require('assert');
const fs = require('fs');
const vm = require('vm');
const code = fs.readFileSync('Code.js', 'utf8');
const setup = fs.readFileSync('Setup.js', 'utf8');
const html = fs.readFileSync('Index.html', 'utf8');
const script = html.split('<script>')[1].split('</script>')[0];

// Search controls and request semantics.
assert(!script.includes('activityTicketClear'), 'no stale missing-element reference');
assert.match(html, /id="activitySearchButton"[^>]*btn btn-primary/);
assert.match(script, /function applyActivitySearch\(\)\{MY_STATE\.search=document\.getElementById\('activityTicketSearch'\)\.value\.trim\(\);MY_STATE\.page=1;loadMyTickets\(true\);\}/);
assert.match(script, /event\.key==='Enter'[\s\S]{0,80}applyActivitySearch\(\)/);
assert.match(script, /activitySearch\.addEventListener\('input',\(\)=>\{\}\)/, 'typing is safe and does not create a duplicate request');
assert(code.indexOf('ticketMatchesActivitySearch_(t,search)') < code.indexOf('paginate_(filtered'), 'search/filter precede pagination');

// All Tickets controls and configured option mapping.
const statusTag = html.match(/<select id="activityStatus"[^>]*>/)[0];
assert(!statusTag.includes('multiple'));
assert(html.includes('<option value="">All statuses</option>'));
assert(!html.includes('id="activityResolvedBy"'));
assert.match(script, /populateClientSizes\('activityClientSize',o\.clientSizes\)/);
assert.match(script, /value="\$\{escAttr\(x\.code\)\}"[^`]*\$\{esc\(x\.label\|\|x\.code\)\}/);
assert.match(script, /size\.disabled=type==='Regular';if\(size\.disabled\)size\.value=''/);
assert.match(script, /MY_STATE=\{\.\.\.MY_STATE,page:r\.page,pageSize:r\.pageSize,totalRows:r\.totalRows,totalPages:r\.totalPages\};applyActivityOptions/);
assert.match(script, /MY_STATE\.totalRows===1\?'ticket':'tickets'/);
assert.match(script, /const labels=\{search:'Search',raisedBy:'Raised by'/);
assert.match(script, /function cascadeActivityCategory/);
assert.match(script, /Created from must not be after Created to/);
assert.match(html, /id="activityMore" class="more-filters" hidden/);
assert.match(html, /Apply Filters/);

// Work Queue consumes the nested filterOptions contract and keeps server authorization.
assert(html.includes('id="qSubcategory"'));
assert.match(script, /QUEUE_STATE\.filterOptions\?\.categories/);
assert.match(script, /populateClientSizes\('qClientSize',QUEUE_STATE\.filterOptions\?\.clientSizes\)/);
assert.match(script, /function cascadeQueueCategory/);
assert(code.includes('requireRole_([APP.ROLES.POC,APP.ROLES.ADMIN])'));
assert.match(script, /QUEUE_STATE\.page=1;loadQueue\(true\)/);

// Server-side matching is case-insensitive, code-based, ordered, and date-key based in Asia/Kolkata.
const context = {console, Date, JSON, Math, Object, String, Number, Boolean, Array, RegExp, Set, Map};
vm.createContext(context);
vm.runInContext(setup+'\n'+code, context);
const ticket={Ticket_ID:'TICKET-42',Client_Name:'Gold Client',Email_Subject:'API Issue',Raiser_Name:'Anant Garg',Raiser_Email:'anant@shadowfax.in',SLA_Due_At:new Date(Date.now()+3600000)};
assert(context.ticketMatchesActivitySearch_(ticket,'ticket-42'));
assert(context.ticketMatchesActivitySearch_(ticket,'gold cli'));
assert(context.ticketMatchesActivitySearch_(ticket,'ANANT'.toLowerCase()));
assert.match(code, /String\(t\.Client_Size\)===configured\.clientSize/);
assert.match(code, /const created=appDateKey_\(t\.Created_At\)/);
assert.match(code, /Utilities\.formatDate\(toDate_\(value\),APP\.TZ,'yyyy-MM-dd'\)/);
assert.strictEqual(context.filterDate_('2026-07-31','created date'),'2026-07-31');
assert.throws(()=>context.validateRange_('2026-08-01','2026-07-31'),/Date from/);
assert(!/getScriptLock|DriveApp/.test(code.slice(code.indexOf('function getAllTickets'),code.indexOf('function getMatchingRows_'))));
console.log('filter UX stability regression checks passed');
