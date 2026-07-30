const assert = require('assert');
const fs = require('fs');
const code = fs.readFileSync('Code.js', 'utf8');
const html = fs.readFileSync('Index.html', 'utf8');
const setup = fs.readFileSync('Setup.js', 'utf8');

// Every Apps Script top-level declaration must be canonical. This deliberately
// scans both production files because Apps Script loads them into one namespace.
const declarations = [];
for (const source of [code, setup]) {
  for (const match of source.matchAll(/^function\s+([A-Za-z_$][\w$]*)\s*\(/gm)) declarations.push(match[1]);
}
const duplicates = [...new Set(declarations.filter((name, i) => declarations.indexOf(name) !== i))];
assert.deepStrictEqual(duplicates, [], `duplicate top-level production functions: ${duplicates.join(', ')}`);

assert(code.includes("const APP_RELEASE = 'shadowfax-ui-stability-v2'"));
assert(html.includes("const EXPECTED_RELEASE = 'shadowfax-ui-stability-v2'"));
assert(html.includes("if(IN_FLIGHT.has(requestKey))return"));
assert(html.includes("form.setAttribute('aria-busy','true')"));
assert(html.includes("form.classList.add('form-submitting')"));
assert(html.includes("method:'submitTicket',payload:form"));
assert(html.includes("rid.value=SUBMISSION_REQUEST_ID"));
assert(html.includes("SUBMISSION_REQUEST_ID=''"));
assert(!html.includes('Array.from(form.elements).forEach(x=>x.disabled=true)'));
assert(!html.includes('Array.from(form.elements).forEach(x=>x.disabled=false)'));
assert(html.indexOf('rid.value=SUBMISSION_REQUEST_ID') < html.indexOf("method:'submitTicket',payload:form"));

for (const field of ['clientMode','clientName','clientType','categoryId','emailSubject','issueDescription','submissionRequestId']) {
  assert(code.includes(`form.${field}`), `submit contract must validate/read ${field}`);
}
assert(code.includes('The category was not submitted. Refresh the page and select the category again.'));
assert(code.includes('The selected category is no longer active. Refresh the page and choose an active category.'));
assert(code.includes('The client size was not submitted. Select a client size and try again.'));
assert(code.includes('getActiveClientSizePriorities_(true)'));
assert(code.includes('removeCachedKeys_([CACHE_KEYS_.CATEGORIES])'));
assert.strictEqual((code.match(/function runEndToEndHealthCheck\s*\(/g)||[]).length, 1);
assert.strictEqual((code.match(/function testTicketSubmissionPayload\s*\(/g)||[]).length, 1);
assert(!/hooks\.slack\.com\/services\//.test(code + setup + html));
assert(!/\brequire\s*\(/.test(code + setup));
console.log('stability, serialization, cache refresh, diagnostics, and duplicate declaration checks passed');
