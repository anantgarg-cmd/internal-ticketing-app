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
    SETTINGS: 'Settings'
  }),
  ROLES: Object.freeze({ SALES: 'SALES', POC: 'POC', ADMIN: 'ADMIN' }),
  STATUS: Object.freeze({ RAISED: 'Raised', INVESTIGATING: 'Investigating', RESOLVED: 'Resolved' }),
  HEADERS: Object.freeze({
    Tickets: [
      'Ticket_ID','Created_At','Raiser_Email','Raiser_Name','Client_Mode','Client_ID','Client_Name','Client_Type',
      'Category_ID','Category_Name','Email_Subject','Normalized_Subject','Issue_Description','Priority','SLA_Hours',
      'SLA_Due_At','Status','Picked_Up_By','Picked_Up_At','Investigating_At','Resolution_Note','Root_Cause',
      'Resolved_By','Resolved_At','SLA_Result','Duplicate_Of','Duplicate_Override','Dynamic_Fields_JSON',
      'Attachment_File_ID','Attachment_File_Name','Attachment_URL','Updated_At','Updated_By'
    ],
    Clients: ['Client_ID','Client_Name','Client_Type','Active'],
    Categories: ['Category_ID','Client_Type','Category_Name','Priority','SLA_Hours','Fields_JSON','Required_Fields_JSON','Active'],
    Users: ['Email','Name','Role','Active'],
    TicketEvents: ['Event_ID','Ticket_ID','Event_Type','Old_Value','New_Value','Performed_By','Created_At','Note'],
    Settings: ['Key','Value','Description']
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
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getDisplayValues()[0]
    : [];
  const hasAnyHeader = existingFirstRow.some(v => String(v).trim() !== '');

  if (!hasAnyHeader) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (current.join('|') !== headers.join('|')) {
      throw new Error(`The ${sheetName} sheet already has different columns. Use a new blank spreadsheet or match the supplied headers.`);
    }
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
  sheet.appendRow([email, 'System Admin', APP.ROLES.ADMIN, true]);
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

  const tickets = ss.getSheetByName(APP.SHEETS.TICKETS);
  tickets.getRange('B:B').setNumberFormat('dd-mmm-yyyy hh:mm');
  tickets.getRange('P:P').setNumberFormat('dd-mmm-yyyy hh:mm');
  tickets.getRange('S:T').setNumberFormat('dd-mmm-yyyy hh:mm');
  tickets.getRange('X:X').setNumberFormat('dd-mmm-yyyy hh:mm');
  tickets.getRange('AF:AF').setNumberFormat('dd-mmm-yyyy hh:mm');

  const events = ss.getSheetByName(APP.SHEETS.EVENTS);
  events.getRange('G:G').setNumberFormat('dd-mmm-yyyy hh:mm');
}
