const assert = require('assert');
const fs = require('fs');

const code = fs.readFileSync('Code.js', 'utf8');
const html = fs.readFileSync('Index.html', 'utf8');

assert(code.includes("const APP_RELEASE = 'production-audit-v1'"));
assert(html.includes("const EXPECTED_RELEASE = 'production-audit-v1'"));
assert(code.includes("function getRuntimeDiagnostics() {\n  requireRole_([APP.ROLES.ADMIN]);"));
assert(!/scriptId:\s*ScriptApp\.getScriptId|serviceUrl:\s*ScriptApp\.getService|companyDomain:/.test(code));
assert(code.includes("function getAuthorizationDiagnostic() {\n  requireRole_([APP.ROLES.ADMIN]);"));
assert(code.includes("throw new Error('The selected Subcategory has an invalid configured priority.')"));
assert(code.includes("throw new Error('The selected Subcategory has invalid configured SLA hours.')"));
assert(!code.includes("number_(category.SLA_Hours, 24)"));
assert(code.includes("const all=getSheetObjects_(APP.SHEETS.TICKET_INDEX), filtered=all.filter"));
assert(code.includes("function cleanupSlackNotificationHistory_(){return{cleaned:0,preserved:true};}"));
assert(!/function cleanupSlackNotificationHistory_[\s\S]{0,900}clearContent/.test(code));
assert(code.includes("throw new Error('You are not allowed to perform this action.');"));

console.log('final production security, configuration, index, retention, and release checks passed');
