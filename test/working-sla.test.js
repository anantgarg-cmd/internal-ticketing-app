const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const formatter = (date, options) => new Intl.DateTimeFormat('en-US', { timeZone: 'Asia/Kolkata', ...options }).formatToParts(date);
const part = (date, type, options) => formatter(date, options).find(item => item.type === type).value;
const Utilities = {
  formatDate(date, timezone, pattern) {
    assert.equal(timezone, 'Asia/Kolkata');
    if (pattern === 'Z') return '+0530';
    if (pattern === 'EEE') return part(date, 'weekday', { weekday: 'short' });
    if (pattern === 'yyyy,M,d,H,m,s,S') {
      const options = { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', second: 'numeric', hourCycle: 'h23' };
      return ['year', 'month', 'day', 'hour', 'minute', 'second'].map(type => part(date, type, options)).concat(date.getUTCMilliseconds()).join(',');
    }
    throw new Error(`Unexpected format: ${pattern}`);
  }
};

const context = { console, Date, Intl, Utilities };
vm.createContext(context);
vm.runInContext("const APP={TZ:'Asia/Kolkata'}; function number_(v,d){const n=Number(v);return Number.isFinite(n)?n:d;} function toDate_(v){return v instanceof Date?v:new Date(v);}" +
  fs.readFileSync('Code.js', 'utf8').slice(fs.readFileSync('Code.js', 'utf8').indexOf('const SLA_WORK_START_HOUR_'), fs.readFileSync('Code.js', 'utf8').indexOf('/**\n * Optional editor/admin migration.')), context);

const cases = [
  ['2026-08-03T10:00:00+05:30', 4, '2026-08-03T15:30:00+05:30'],
  ['2026-08-03T11:30:00+05:30', 4, '2026-08-03T15:30:00+05:30'],
  ['2026-08-03T18:30:00+05:30', 4, '2026-08-04T14:30:00+05:30'],
  ['2026-08-07T18:30:00+05:30', 4, '2026-08-10T14:30:00+05:30'],
  ['2026-08-07T19:29:00+05:30', 4, '2026-08-10T15:29:00+05:30'],
  ['2026-08-07T19:30:00+05:30', 4, '2026-08-10T15:30:00+05:30'],
  ['2026-08-08T13:00:00+05:30', 4, '2026-08-10T15:30:00+05:30'],
  ['2026-08-03T11:30:00+05:30', 16, '2026-08-05T11:30:00+05:30']
];
for (const [created, hours, expected] of cases) {
  const actual = context.calculateWorkingSlaDueAt_(new Date(created), hours);
  assert.equal(actual.toISOString(), new Date(expected).toISOString(), `${created} + ${hours} working hours`);
}
console.log(`Passed ${cases.length} working SLA cases.`);
