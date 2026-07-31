const assert = require('node:assert/strict');
const fs = require('node:fs');
const html = fs.readFileSync('Index.html','utf8');
const code = fs.readFileSync('Code.js','utf8');
const manifest = JSON.parse(fs.readFileSync('appsscript.json','utf8'));
const workflow = fs.readFileSync('.github/workflows/deploy-apps-script.yml','utf8');
const ignore = fs.readFileSync('.claspignore','utf8');

// Brand and release.
assert(html.includes('https://www.shadowfax.in/logo_header_tm.svg'));
for (const token of ['--sf-green:#008A71','--sf-yellow:#D5D226','--sf-black:#231F20']) assert(html.includes(token));
assert(html.includes('object-fit:contain') && html.includes('filter:none') && html.includes('box-shadow:none'));
assert(html.includes('"Montserrat",Arial,sans-serif'));
assert(code.includes("const APP_RELEASE = 'simplified-ticket-forms-v1'"));
assert(html.includes("const EXPECTED_RELEASE = 'simplified-ticket-forms-v1'"));

// Performance and Apps Script contracts.
for (const forbidden of ['React','Vue','Angular','jQuery','setInterval(','requestAnimationFrame(function']) assert(!html.includes(forbidden), forbidden);
assert.equal((html.match(/method:'getInitialAppState'/g)||[]).length,1);
assert(html.includes('},300)'));
assert(html.includes("method:'submitTicket',payload:form"));
assert.match(html,/type="file" name="attachment"/);
assert(!/form\.elements[^\n]*disabled\s*=/.test(html));
assert(!/\brequire\s*\(/.test(code));
assert(ignore.includes('node_modules') && ignore.includes('test'));

// Responsive shell, tables/cards, and drawer.
for (const marker of ['sidebar-collapsed','mobile-menu','desktop-table','ticket-cards','detail-drawer','@media(max-width:1024px)','@media(max-width:768px)','@media(max-width:390px)']) assert(html.includes(marker), marker);
assert(html.includes('role="dialog" aria-modal="true"'));
assert(html.includes("e.key==='Escape'") && html.includes('LAST_DRAWER_TRIGGER?.focus()'));
assert(html.includes('data-copy-id') && html.includes('copyTicketId'));
assert(html.includes('aria-current'));
assert(html.includes('role="status" aria-live="polite" aria-busy="true"'));
assert(html.includes('aria-busy="true"'));
assert(html.includes('@media(prefers-reduced-motion:reduce)'));
assert(html.includes('tabindex="0"'));

// Local-only preferences and safe current-tab draft.
assert(html.includes("localStorage.setItem('sf-sidebar-collapsed'"));
assert(html.includes("sessionStorage.setItem('sf-ticket-draft-v2'"));
assert(html.includes('setTimeout(saveDraft,500)'));
assert(!/sessionStorage\.setItem[^\n]*(attachment:|fileName|fileSize|userEmail)/i.test(html));
assert(html.includes('clearDraft();form.reset()'));

assert.equal(manifest.webapp.access,'DOMAIN');
assert.equal(manifest.webapp.executeAs,'USER_DEPLOYING');
assert(workflow.includes('- main'));
console.log('69-point design contract representative static checks passed');
