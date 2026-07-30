const assert = require('assert');
const fs = require('fs');
const code = fs.readFileSync('Code.js','utf8');
const setup = fs.readFileSync('Setup.js','utf8');
const manifest = JSON.parse(fs.readFileSync('appsscript.json','utf8'));
const queueHeaders=['Notification_ID','Dedupe_Key','Notification_Type','Ticket_ID','SLA_Cycle_Number','Priority','Payload_JSON','Status','Attempts','Next_Attempt_At','Created_At','Processing_Started_At','Sent_At','Last_HTTP_Code','Last_Error','Updated_At'];
queueHeaders.forEach(h=>assert(setup.includes("'"+h+"'"),`missing ${h}`));
['SLACK_NOTIFICATIONS_ENABLED','SLACK_ALERT_PRIORITIES','SLACK_MENTION_PRIORITIES','SLACK_BREACH_WARNING_ENABLED','SLACK_BREACH_WARNING_MINUTES','SLACK_BREACH_ALERT_ENABLED','SLACK_EOD_SUMMARY_ENABLED','SLACK_EOD_HOUR','SLACK_EOD_MINUTE','SLACK_DISPATCH_BATCH_SIZE','SLACK_MAX_RETRIES','SLACK_PROCESSING_TIMEOUT_MINUTES','SLACK_NOTIFICATION_RETENTION_DAYS'].forEach(k=>assert(setup.includes(k),`missing ${k}`));
['enqueueSlackNotification_','calculateWorkingMinutesBetween_','monitorSlackAlerts','enqueueEndOfDaySlackSummary_','dispatchSlackNotifications','cleanupSlackNotificationHistory_','setupSlackAutomationTriggers','removeSlackAutomationTriggers','validateSlackAutomation','testSlackConnection'].forEach(fn=>assert.strictEqual((code.match(new RegExp('function '+fn+'\\s*\\(','g'))||[]).length,1,`${fn} not canonical`));
assert(manifest.oauthScopes.includes('https://www.googleapis.com/auth/script.scriptapp'));
assert.strictEqual(manifest.webapp.access,'DOMAIN');assert.strictEqual(manifest.webapp.executeAs,'USER_DEPLOYING');assert.strictEqual(manifest.timeZone,'Asia/Kolkata');
assert(!/hooks\.slack\.com\/services\//.test(code+setup+fs.readFileSync('README.md','utf8')),'webhook secret committed');
const fetches=[...code.matchAll(/UrlFetchApp\.fetch/g)].map(m=>code.slice(0,m.index).split('\n').length);
assert.strictEqual(fetches.length,2,'only authorization diagnostic and low-level Slack sender may fetch');
const sender=code.slice(code.indexOf('function sendSlackWebhookPayload_'),code.indexOf('function retryDelayMinutes_'));
assert(sender.includes('UrlFetchApp.fetch'));
const submit=code.slice(code.lastIndexOf('function submitTicket'),code.lastIndexOf('function updateTicketStatus'));
const reopen=code.slice(code.lastIndexOf('function reopenTicket'),code.indexOf('function getRecentTicketObjects_',code.lastIndexOf('function reopenTicket')));
assert(!submit.includes('UrlFetchApp.fetch'));assert(!reopen.includes('UrlFetchApp.fetch'));
assert(code.includes("getSheetObjects_(APP.SHEETS.TICKET_INDEX)"));
assert(code.includes("'NEW_HIGH:' + ticket.ticketId"));assert(code.includes("'REOPEN_HIGH:' + ticket.ticketId + ':' + cycle"));
console.log('Slack notification static/security tests passed');
// Execute the production chunked working-minute function with an IST calendar stub.
const vm=require('vm');
const match=code.match(/function calculateWorkingMinutesBetween_\([\s\S]*?\n}/);
const ctx={
  toDate_:v=>new Date(v),
  isWorkingDay_:d=>{const day=new Date(d.getTime()+330*60000).getUTCDay();return day!==0&&day!==6;},
  getWorkingDayStart_:d=>{const x=new Date(d.getTime()+330*60000);return new Date(Date.UTC(x.getUTCFullYear(),x.getUTCMonth(),x.getUTCDate(),6,0));},
  getWorkingDayEnd_:d=>{const x=new Date(d.getTime()+330*60000);return new Date(Date.UTC(x.getUTCFullYear(),x.getUTCMonth(),x.getUTCDate(),14,0));}
};
ctx.moveToNextWorkingStart_=d=>{let x=new Date(d.getTime()+330*60000);do{x=new Date(Date.UTC(x.getUTCFullYear(),x.getUTCMonth(),x.getUTCDate()+1,6,0));}while(!ctx.isWorkingDay_(x));return x;};
vm.createContext(ctx);vm.runInContext(match[0],ctx);
assert.strictEqual(ctx.calculateWorkingMinutesBetween_('2026-08-03T17:30:00+05:30','2026-08-04T11:30:00+05:30'),120);
assert.strictEqual(ctx.calculateWorkingMinutesBetween_('2026-08-07T18:30:00+05:30','2026-08-10T12:30:00+05:30'),120);
assert.strictEqual(ctx.calculateWorkingMinutesBetween_('2026-08-08T10:00:00+05:30','2026-08-10T13:30:00+05:30'),120);
assert.strictEqual(ctx.calculateWorkingMinutesBetween_('2026-08-04T12:00:00+05:30','2026-08-04T11:00:00+05:30'),0);
console.log('Slack working-minute examples passed');
