/**
 * ONE-TIME SETUP FILE
 * Run setupSystem() once from the Apps Script editor.
 */

const APP = Object.freeze({
  TITLE: 'Internal Issue Ticketing System',
  TZ: 'Asia/Kolkata',
  SHEETS: Object.freeze({
    TICKETS: 'Tickets',
    CLIENTS: 'Clients',
    CATEGORIES: 'Categories',
    USERS: 'Users',
    EVENTS: 'TicketEvents',
    SETTINGS: 'Settings',
    SLA_CYCLES: 'TicketSLACycles',
    CLIENT_SIZE_PRIORITY: 'ClientSizePriority',
    TICKET_INDEX: 'TicketIndex'
  }),
  ROLES: Object.freeze({ SALES: 'SALES', POC: 'POC', ADMIN: 'ADMIN' }),
  STATUS: Object.freeze({ RAISED: 'Raised', REOPENED: 'Reopened', INVESTIGATING: 'Investigating', RESOLVED: 'Resolved' }),
  HEADERS: Object.freeze({
    Tickets: [
      'Ticket_ID','Created_At','Raiser_Email','Raiser_Name','Client_Mode','Client_ID','Client_Name','Client_Type','Client_Size',
      'Category_ID','Category_Name','Email_Subject','Normalized_Subject','Issue_Description','Priority','SLA_Hours',
      'SLA_Due_At','Status','Picked_Up_By','Picked_Up_At','Investigating_At','Resolution_Note','Root_Cause',
      'Resolved_By','Resolved_At','SLA_Result','Duplicate_Of','Duplicate_Override','Dynamic_Fields_JSON',
      'Attachment_File_ID','Attachment_File_Name','Attachment_URL','Updated_At','Updated_By','Priority_Source','Submission_Request_ID'
    ],
    Clients: ['Client_ID','Client_Name','Client_Type','Active'],
    Categories: ['Category_ID','Client_Type','Category_Name','Priority','SLA_Hours','Fields_JSON','Required_Fields_JSON','Active'],
    Users: ['Email','Name','Role','Active'],
    TicketEvents: ['Event_ID','Ticket_ID','Event_Type','Old_Value','New_Value','Performed_By','Created_At','Note','Request_ID'],
    Settings: ['Key','Value','Description'],
    TicketSLACycles: [
      'SLA_Cycle_ID','Ticket_ID','Cycle_Number','Cycle_Type','Started_At','Due_At','Ended_At','SLA_Result',
      'Started_By','Ended_By','Reopen_Reason','Created_At','Updated_At'
    ],
    ClientSizePriority: ['Client_Size_Code','Display_Label','ADL_Description','Min_ADL','Max_ADL','Priority','Active','Sort_Order'],
    TicketIndex: ['Ticket_ID','Created_At','Raiser_Email','Raiser_Name','Client_ID','Client_Name','Client_Type','Client_Size','Client_Key','Category_ID','Category_Name','Email_Subject','Normalized_Subject','Priority','Priority_Source','SLA_Due_At','Status','Resolved_At','SLA_Result','Current_SLA_Cycle','Submission_Request_ID','Updated_At']
  })
});

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet first, then go to Extensions → Apps Script and run setupSystem().');

  ss.setSpreadsheetTimeZone(APP.TZ);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(APP.SHEETS).forEach(key => {
    const sheetName = APP.SHEETS[key];
    ensureSheet_(ss, sheetName, APP.HEADERS[sheetName]);
  });

  seedSettings_(ss);
  seedCategories_(ss);
  seedClientSizePriority_(ss);
  seedAdminUser_(ss);
  formatSheets_(ss);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ATTACHMENT_FOLDER_ID')) {
    const folder = DriveApp.createFolder('Internal Issue Ticketing - Attachments');
    props.setProperty('ATTACHMENT_FOLDER_ID', folder.getId());
  }

  return {
    success: true,
    spreadsheetId: ss.getId(),
    attachmentFolderId: props.getProperty('ATTACHMENT_FOLDER_ID'),
    adminEmail: Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail(),
    message: 'Setup complete. Now fill the Clients and Users sheets, then deploy the web app.'
  };
}

function ensureSheet_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  const existingFirstRow = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    : [];
  const hasAnyHeader = existingFirstRow.some(v => String(v).trim() !== '');

  if (!hasAnyHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = existingFirstRow.map(value => String(value).trim());
    const missing = headers.filter(header => current.indexOf(header) < 0);
    if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
  }
  sheet.setFrozenRows(1);
}

function seedSettings_(ss) {
  const sheet = ss.getSheetByName(APP.SHEETS.SETTINGS);
  if (sheet.getLastRow() > 1) return;

  const activeEmail = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const domain = activeEmail.includes('@') ? activeEmail.split('@')[1] : 'yourcompany.com';

  const rows = [
    ['COMPANY_DOMAIN', domain, 'Only email addresses from this domain can use the app.'],
    ['DUPLICATE_WINDOW_DAYS', '5', 'Look back this many days for possible duplicate tickets.'],
    ['DUPLICATE_SIMILARITY_THRESHOLD', '0.65', '0 to 1. Higher means stricter subject matching.'],
    ['RESOLVED_VISIBILITY_DAYS', '10', 'Sales can see their resolved tickets for this many days.'],
    ['DASHBOARD_WINDOW_DAYS', '14', 'Numbers dashboard lookback period.'],
    ['ALERT_PRIORITIES', 'HIGH', 'Comma-separated priorities that trigger Slack alerts, e.g. HIGH,CRITICAL.'],
    ['MAX_ATTACHMENT_MB', '5', 'Maximum evidence file size.'],
    ['ROOT_CAUSES', 'Product Bug|Configuration Issue|Data Issue|Integration/API Issue|Access/Permission|User Error|External Dependency|Process Gap|Unable to Reproduce|Other', 'Values shown when resolving a ticket.']
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedAdminUser_(ss) {
  const sheet = ss.getSheetByName(APP.SHEETS.USERS);
  if (sheet.getLastRow() > 1) return;
  const email = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();
  if (!email) throw new Error('Could not read your Google Workspace email. Run setup from your company account.');
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, 4).setValues([[email, 'System Admin', APP.ROLES.ADMIN, true]]);
}

function seedCategories_(ss) {
  const sheet = ss.getSheetByName(APP.SHEETS.CATEGORIES);
  if (sheet.getLastRow() > 1) return;

  const rows = [
    category_('360-CHANNEL', '360', 'Channel Integration', 'HIGH', 8, ['order_id','endpoint','api_log','attachment'], ['order_id']),
    category_('360-PREPAID', '360', 'Pre-paid Wallet', 'HIGH', 8, ['wallet_reference','attachment'], []),
    category_('360-POSTPAID', '360', 'Post-paid Wallet', 'HIGH', 8, ['wallet_reference','attachment'], []),
    category_('360-API', '360', 'API', 'HIGH', 8, ['endpoint','order_id','api_log','attachment'], ['endpoint','api_log']),
    category_('360-LABEL', '360', 'Label', 'MEDIUM', 24, ['awb','label_type','attachment'], ['awb']),
    category_('360-INVOICE', '360', 'Invoice', 'MEDIUM', 24, ['invoice_reference','attachment'], ['invoice_reference']),
    category_('360-COD', '360', 'COD Remittance', 'HIGH', 8, ['awb','amount','expected_date','attachment'], ['awb']),
    category_('360-TOKEN', '360', 'Token', 'HIGH', 8, ['endpoint','api_log','attachment'], ['api_log']),
    category_('360-CPSELLER', '360', 'CP–Seller Linking', 'MEDIUM', 24, ['seller_id','attachment'], ['seller_id']),
    category_('360-LOGIN', '360', 'Login', 'HIGH', 8, ['user_identifier','attachment'], ['user_identifier']),
    category_('360-ONBOARDING', '360', 'Onboarding', 'MEDIUM', 24, ['user_identifier','attachment'], []),
    category_('360-OTHER', '360', 'Other Portal Issues', 'MEDIUM', 24, ['attachment'], []),

    category_('REG-SERVICE', 'Regular', 'Serviceability', 'MEDIUM', 24, ['origin_pincode','destination_pincode','service_type','attachment'], ['origin_pincode','destination_pincode']),
    category_('REG-CREATE', 'Regular', 'Order Creation API', 'HIGH', 8, ['endpoint','order_id','api_log','attachment'], ['endpoint','api_log']),
    category_('REG-CANCEL', 'Regular', 'Cancellation API', 'HIGH', 8, ['endpoint','order_id','api_log','attachment'], ['endpoint','order_id']),
    category_('REG-CALLBACK', 'Regular', 'Callbacks/Share Callbacks', 'HIGH', 8, ['awb','endpoint','api_log','attachment'], ['endpoint']),
    category_('REG-TRACK', 'Regular', 'Track API', 'HIGH', 8, ['awb','endpoint','api_log','attachment'], ['awb']),
    category_('REG-STATUS', 'Regular', 'Status Mapping', 'MEDIUM', 24, ['awb','expected_status','actual_status','attachment'], ['awb','expected_status','actual_status']),
    category_('REG-TOKEN', 'Regular', 'Unable to Fetch Tokens', 'HIGH', 8, ['endpoint','api_log','attachment'], ['api_log']),
    category_('REG-LABEL', 'Regular', 'Label API', 'MEDIUM', 24, ['awb','endpoint','api_log','attachment'], ['awb']),
    category_('REG-COMMS', 'Regular', 'Communication Samples', 'LOW', 48, ['awb','attachment'], ['awb'])
  ];

  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function category_(id, clientType, name, priority, slaHours, fields, required) {
  return [id, clientType, name, priority, slaHours, JSON.stringify(fields), JSON.stringify(required), true];
}

function clientSizePrioritySeedRows_() {
  return [
    ['GOLD_PLATINUM','Gold / Platinum','ADL 300 and above',300,'','HIGH',true,1],
    ['MEDIUM_SIZED','Medium-sized','ADL between 50 and 299',50,299,'MEDIUM',true,2],
    ['SMALL_SIZED','Small-sized','ADL below 50',0,49,'LOW',true,3]
  ];
}

function seedClientSizePriority_(ss) {
  const sheet = ss.getSheetByName(APP.SHEETS.CLIENT_SIZE_PRIORITY);
  const headers = APP.HEADERS.ClientSizePriority;
  const existing = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const codeColumn = headers.indexOf('Client_Size_Code');
  const codes = existing.map(row => String(row[codeColumn] || ''));
  const missing = clientSizePrioritySeedRows_().filter(row => codes.indexOf(row[0]) < 0);
  if (missing.length) sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, headers.length).setValues(missing);
  return missing.map(row => row[0]);
}

function formatSheets_(ss) {
  Object.values(APP.SHEETS).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    const lastCol = sheet.getLastColumn();
    if (!lastCol) return;
    sheet.getRange(1, 1, 1, lastCol)
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#ffffff')
      .setWrap(true);
    sheet.autoResizeColumns(1, lastCol);
    sheet.setRowHeight(1, 32);
  });

  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.TICKETS),
    ['Created_At','SLA_Due_At','Picked_Up_At','Investigating_At','Resolved_At','Updated_At']);
  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.EVENTS), ['Created_At']);
  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.SLA_CYCLES),
    ['Started_At','Due_At','Ended_At','Created_At','Updated_At']);
}

function formatDateTimeColumns_(sheet, headerNames) {
  if (!sheet || !sheet.getLastColumn()) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
  headerNames.forEach(header => {
    const column = headers.indexOf(header) + 1;
    if (column) sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('dd-mmm-yyyy hh:mm');
  });
}

function formatSlaCycleSheet_(sheet) {
  const lastColumn = sheet.getLastColumn();
  sheet.setFrozenRows(1);
  if (lastColumn) {
    sheet.getRange(1, 1, 1, lastColumn).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff').setWrap(true);
    sheet.setRowHeight(1, 32);
    sheet.autoResizeColumns(1, lastColumn);
  }
  formatDateTimeColumns_(sheet, ['Started_At','Due_At','Ended_At','Created_At','Updated_At']);
}

/** Additive, idempotent schema repair. The optional argument is for callers already holding ScriptLock. */
function ensureTicketSlaCyclesSchema_(lockAlreadyHeld) {
  const ss = getSpreadsheet_();
  const headers = APP.HEADERS.TicketSLACycles;
  let sheet = ss.getSheetByName(APP.SHEETS.SLA_CYCLES);
  let lock = null;
  let sheetCreated = false;
  const columnsAdded = [];
  if (!lockAlreadyHeld) {
    lock = LockService.getScriptLock();
    lock.waitLock(30000);
  }
  try {
    sheet = ss.getSheetByName(APP.SHEETS.SLA_CYCLES);
    if (!sheet) {
      sheet = ss.insertSheet(APP.SHEETS.SLA_CYCLES);
      sheetCreated = true;
    }
    const existing = sheet.getLastColumn()
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(value => String(value).trim())
      : [];
    const populated = existing.filter(Boolean);
    const missing = headers.filter(header => populated.indexOf(header) < 0);
    if (missing.length) {
      const startColumn = existing.length + 1;
      sheet.getRange(1, startColumn, 1, missing.length).setValues([missing]);
      Array.prototype.push.apply(columnsAdded, missing);
    }
    formatSlaCycleSheet_(sheet);
    return { sheetCreated, columnsAdded, ready: true };
  } finally {
    if (lock) lock.releaseLock();
  }
}

/** Safe public upgrade for legacy installations; never clears or rebuilds source sheets. */
function upgradeSlaCycleSchema() {
  const result = ensureTicketSlaCyclesSchema_();
  return { success: true, ready: result.ready, sheetCreated: result.sheetCreated, columnsAdded: result.columnsAdded,
    message: result.sheetCreated ? 'TicketSLACycles was created.' : (result.columnsAdded.length ? 'Missing SLA-cycle columns were appended.' : 'TicketSLACycles was already ready.') };
}

/** Complete additive upgrade entry point retained for deployment runbooks. */
function upgradeClientSizeAndPerformanceSchema() {
  const user = requireRole_([APP.ROLES.ADMIN]);
  const ss = getSpreadsheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  const createdSheets = [], columnsAdded = {};
  try {
    [APP.SHEETS.CLIENT_SIZE_PRIORITY, APP.SHEETS.TICKET_INDEX].forEach(name => {
      if (!ss.getSheetByName(name)) createdSheets.push(name);
      ensureSheet_(ss, name, APP.HEADERS[name]);
    });
    [[APP.SHEETS.TICKETS, ['Client_Size','Priority_Source','Submission_Request_ID']], [APP.SHEETS.EVENTS, ['Request_ID']]].forEach(spec => {
      const sheet = ss.getSheetByName(spec[0]);
      const current = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
      const missing = spec[1].filter(header => current.indexOf(header) < 0);
      if (missing.length) sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
      columnsAdded[spec[0]] = missing;
    });
  } finally { lock.releaseLock(); }
  const seededClientSizes = seedClientSizePriority_(ss);
  const slaCycles = ensureTicketSlaCyclesSchema_();
  const index = rebuildTicketIndex_(true);
  invalidateApplicationCaches_();
  formatSheets_(ss);
  return { success: true, runByRole: user.role, createdSheets, columnsAdded, seededClientSizes, indexedTickets: index.indexed, slaCycles,
    message: 'Additive upgrade complete. Source and historical rows were not cleared, moved, or reprioritised.' };
}
