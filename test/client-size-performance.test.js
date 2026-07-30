const assert = require('assert');
const fs = require('fs');
const code = fs.readFileSync('Code.js','utf8');
const setup = fs.readFileSync('Setup.js','utf8');
const html = fs.readFileSync('Index.html','utf8');

for (const value of [
  "['GOLD_PLATINUM','Gold / Platinum','ADL 300 and above',300,'','HIGH',8,true,1]",
  "['MEDIUM_SIZED','Medium-sized','ADL between 50 and 299',50,299,'MEDIUM',24,true,2]",
  "['SMALL_SIZED','Small-sized','ADL below 50',0,49,'LOW',48,true,3]"
]) assert(setup.includes(value), `missing seed ${value}`);
assert(code.includes("prioritySource: 'CLIENT_SIZE'"));
assert(code.includes("prioritySource: 'CATEGORY'"));
assert(code.includes("slaSource: 'CLIENT_SIZE'"));
assert(code.includes("slaSource: 'CATEGORY'"));
assert(code.includes('slaHours=resolution.slaHours'));
assert(code.includes('SLA_Source:resolution.slaSource'));
assert(code.includes("Choose an active Client Size for a 360 client."));
assert(code.includes("clientSize: ''"));
assert(code.includes("APP.SHEETS.TICKET_INDEX"));
assert(code.includes("Math.min(100"));
assert(code.includes("Submission_Request_ID is required"));
assert(code.includes("Action_Request_ID is required"));
assert(code.includes("getMatchingRows_(APP.SHEETS.EVENTS"));
assert(code.includes("saveAttachment_(form.attachment,ticketId,user.email)"));
assert(code.indexOf("saveAttachment_(form.attachment,ticketId,user.email)") < code.lastIndexOf("lock=LockService.getScriptLock()"));
assert(html.includes("Sales workspace"));
assert(html.includes("Tech/Product workspace"));
assert(html.includes("ADMIN</span>"));
assert(html.includes("showPage(data.user.role==='SALES'?'raisePage':'queuePage')"));
assert(html.includes("setTimeout(()=>{QUEUE_STATE.page=1;loadQueue(true);},300)"));
assert(html.includes("Loading your workspace…"));
assert(html.includes("crypto?.randomUUID"));
assert(html.includes("Uploading evidence and creating your ticket…"));
assert(html.includes("Client Size</th>"));
assert(!html.includes('localStorage'));
assert(html.includes('For 360 clients, priority and SLA are based on client size.'));
assert(html.includes('size.slaHours'));
console.log('client-size/performance static checks passed');
