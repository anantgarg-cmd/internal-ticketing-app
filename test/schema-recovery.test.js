const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

class Range {
  constructor(sheet, row, col, rows, cols) { Object.assign(this, { sheet, row, col, rows, cols }); }
  getDisplayValues() { return this.getValues().map(row => row.map(value => String(value ?? ''))); }
  getValues() { return Array.from({ length: this.rows }, (_, r) => Array.from({ length: this.cols }, (_, c) => this.sheet.data[this.row - 1 + r]?.[this.col - 1 + c] ?? '')); }
  setValues(values) { values.forEach((row, r) => row.forEach((value, c) => { const y=this.row-1+r,x=this.col-1+c; this.sheet.data[y] ||= []; this.sheet.data[y][x]=value; })); return this; }
  setNumberFormat() { return this; } setFontWeight() { return this; } setBackground() { return this; }
  setFontColor() { return this; } setWrap() { return this; }
}
class Sheet {
  constructor(name, data=[]) { this.name=name; this.data=data.map(row=>row.slice()); this.frozen=0; }
  getLastColumn() { return this.data.reduce((max,row)=>Math.max(max,row.length),0); }
  getLastRow() { let n=this.data.length; while(n && !(this.data[n-1]||[]).some(v=>v!==''&&v!=null)) n--; return n; }
  getMaxRows() { return Math.max(1000,this.data.length); }
  getRange(row,col,rows=1,cols=1) { assert.equal(typeof row,'number'); return new Range(this,row,col,rows,cols); }
  setFrozenRows(n) { this.frozen=n; } setRowHeight() {} autoResizeColumns() {}
}
class Spreadsheet {
  constructor() { this.sheets=new Map(); this.created=0; }
  getSheetByName(name) { return this.sheets.get(name)||null; }
  insertSheet(name) { assert(!this.sheets.has(name)); const s=new Sheet(name); this.sheets.set(name,s); this.created++; return s; }
}
const ss = new Spreadsheet();
const props = new Map([['SPREADSHEET_ID','test-sheet']]);
let lockDepth=0;
const context = {
  console, Date, Math, JSON, Object, Array, String, Boolean, Number, RegExp, Error,
  SpreadsheetApp: { openById: id => { assert.equal(id,'test-sheet'); return ss; }, getActiveSpreadsheet:()=>null },
  PropertiesService: { getScriptProperties:()=>({getProperty:k=>props.get(k)||null,setProperty:(k,v)=>props.set(k,v)}) },
  LockService: { getScriptLock:()=>({waitLock:()=>{assert.equal(lockDepth,0);lockDepth++;},releaseLock:()=>{assert.equal(lockDepth,1);lockDepth--;}}) },
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('Setup.js','utf8'),context);
vm.runInContext(fs.readFileSync('Code.js','utf8'),context);

const expected=['SLA_Cycle_ID','Ticket_ID','Cycle_Number','Cycle_Type','Started_At','Due_At','Ended_At','SLA_Result','Started_By','Ended_By','Reopen_Reason','Created_At','Updated_At'];
let result=context.ensureTicketSlaCyclesSchema_();
assert.deepEqual(JSON.parse(JSON.stringify(result)),{sheetCreated:true,columnsAdded:expected,ready:true});
assert.deepEqual(ss.getSheetByName('TicketSLACycles').data[0],expected);
assert.equal(ss.getSheetByName('TicketSLACycles').frozen,1);
assert.equal(ss.created,1);

const sheet=ss.getSheetByName('TicketSLACycles');
sheet.data=[['Ticket_ID','SLA_Cycle_ID'],['TKT-LEGACY','cycle-1']];
result=context.ensureTicketSlaCyclesSchema_();
assert.equal(result.sheetCreated,false);
assert.deepEqual(sheet.data[0].slice(0,2),['Ticket_ID','SLA_Cycle_ID']);
assert.deepEqual(sheet.data[1],['TKT-LEGACY','cycle-1']);
assert.deepEqual(JSON.parse(JSON.stringify(result.columnsAdded)),expected.filter(h=>!['Ticket_ID','SLA_Cycle_ID'].includes(h)));
context.ensureTicketSlaCyclesSchema_();
assert.equal(ss.created,1);
assert.deepEqual(sheet.data[1],['TKT-LEGACY','cycle-1']);

props.delete('APP_SCHEMA_VERSION');
const runtime=context.ensureRuntimeSchema_();
assert.equal(runtime.ready,true);
assert.equal(props.get('APP_SCHEMA_VERSION'),'sla-cycles-v1');
const createdBefore=ss.created;
assert.deepEqual(JSON.parse(JSON.stringify(context.ensureRuntimeSchema_())),{ready:true,checked:false});
assert.equal(ss.created,createdBefore);
assert.equal(lockDepth,0);

// Read-only detail helper degrades to an empty list when recovery unexpectedly fails.
const originalEnsure=context.ensureTicketSlaCyclesSchema_;
ss.sheets.delete('TicketSLACycles');
context.ensureTicketSlaCyclesSchema_=()=>{throw new Error('simulated failure');};
assert.equal(context.getTicketSlaCycles_('TKT-1').length,0);
context.ensureTicketSlaCyclesSchema_=originalEnsure;

console.log('Passed additive SLA-cycle schema and runtime-marker recovery tests.');
