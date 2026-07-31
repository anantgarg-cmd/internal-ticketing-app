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
    TICKET_INDEX: 'TicketIndex',
    SLACK_NOTIFICATIONS: 'SlackNotifications',
    FEATURE_REQUESTS: 'FeatureRequests',
    FEATURE_REQUEST_EVENTS: 'FeatureRequestEvents'
  }),
  ROLES: Object.freeze({ SALES: 'SALES', POC: 'POC', ADMIN: 'ADMIN' }),
  STATUS: Object.freeze({ RAISED: 'Raised', REOPENED: 'Reopened', INVESTIGATING: 'Investigating', RESOLVED: 'Resolved' }),
  HEADERS: Object.freeze({
    Tickets: [
      'Ticket_ID','Created_At','Raiser_Email','Raiser_Name','Client_Mode','Client_ID','Client_Name','Client_Type',
      'Category_ID','Category_Name','Email_Subject','Normalized_Subject','Issue_Description','Priority','SLA_Hours',
      'SLA_Due_At','Status','Picked_Up_By','Picked_Up_At','Investigating_At','Resolution_Note','Root_Cause',
      'Resolved_By','Resolved_At','SLA_Result','Duplicate_Of','Duplicate_Override','Dynamic_Fields_JSON',
      'Attachment_File_ID','Attachment_File_Name','Attachment_URL','Updated_At','Updated_By','Client_Size','Priority_Source','SLA_Source','Submission_Request_ID','Category_Group','Subcategory_Name'
    ],
    Clients: ['Client_ID','Client_Name','Client_Type','Active'],
    Categories: ['Category_ID','Client_Type','Category_Name','Priority','SLA_Hours','Fields_JSON','Required_Fields_JSON','Active','Category_Group','Subcategory_Name','Visible_In_Form','Sort_Order'],
    Users: ['Email','Name','Role','Active'],
    TicketEvents: ['Event_ID','Ticket_ID','Event_Type','Old_Value','New_Value','Performed_By','Created_At','Note','Request_ID'],
    Settings: ['Key','Value','Description'],
    TicketSLACycles: [
      'SLA_Cycle_ID','Ticket_ID','Cycle_Number','Cycle_Type','Started_At','Due_At','Ended_At','SLA_Result',
      'Started_By','Ended_By','Reopen_Reason','Created_At','Updated_At'
    ],
    ClientSizePriority: ['Client_Size_Code','Display_Label','ADL_Description','Min_ADL','Max_ADL','Priority','SLA_Hours','Active','Sort_Order'],
    TicketIndex: ['Ticket_ID','Created_At','Raiser_Email','Raiser_Name','Client_ID','Client_Name','Client_Type','Client_Size','Client_Key','Category_ID','Category_Name','Email_Subject','Normalized_Subject','Priority','Priority_Source','SLA_Due_At','Status','Resolved_At','SLA_Result','Current_SLA_Cycle','Submission_Request_ID','Updated_At','Category_Group','Subcategory_Name'],
    SlackNotifications: ['Notification_ID','Dedupe_Key','Notification_Type','Ticket_ID','SLA_Cycle_Number','Priority','Payload_JSON','Status','Attempts','Next_Attempt_At','Created_At','Processing_Started_At','Sent_At','Last_HTTP_Code','Last_Error','Updated_At'],
    FeatureRequests: ['Feature_Request_ID','Created_At','Requested_By_Email','Requested_By_Name','Client_ID','Client_Name','Client_Type','Client_Size','Product_Area','Feature_Title','Problem_Statement','Business_Case','Expected_Benefit','Business_Urgency','Affected_Clients_Count','Expected_Order_Volume','Revenue_Impact','Current_Workaround','Supporting_Link','Attachment_File_ID','Attachment_File_Name','Attachment_URL','Status','Product_Priority','Product_Owner','Product_Response','Target_Timeline','Submission_Request_ID','Updated_At','Updated_By'],
    FeatureRequestEvents: ['Feature_Request_Event_ID','Feature_Request_ID','Event_Type','Old_Value','New_Value','Performed_By','Note','Request_ID','Created_At']
  })
});

// Version 6 is already the client-size SLA release; Slack is the next additive migration.
const CURRENT_SCHEMA_VERSION = 8;
const APP_SCHEMA_VERSION_PROPERTY_ = 'APP_SCHEMA_VERSION';
const SCHEMA_READY_CACHE_KEY_ = 'app:schema-ready:v8';
const SCHEMA_UPGRADE_IN_PROGRESS_ = 'SCHEMA_UPGRADE_IN_PROGRESS';
const SCHEMA_TEMPORARY_MESSAGE_ = 'The application structure is being upgraded. Please refresh once after a few seconds.';
const SCHEMA_ADMIN_MESSAGE_ = 'The application structure could not be prepared automatically. Please ask the administrator to run repairApplicationSchema() from Apps Script.';

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the Google Sheet first, then go to Extensions → Apps Script and run setupSystem().');

  ss.setSpreadsheetTimeZone(APP.TZ);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  runSchemaMigrations_(0, CURRENT_SCHEMA_VERSION, ss, newSchemaSummary_());

  seedSettings_(ss);
  seedCategories_(ss);
  seedClientSizePriority_(ss);
  seedAdminUser_(ss);
  formatSheets_(ss);
  PropertiesService.getScriptProperties().setProperty(APP_SCHEMA_VERSION_PROPERTY_, String(CURRENT_SCHEMA_VERSION));
  invalidateSchemaCaches_();

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
  const activeEmail = (Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const domain = activeEmail.includes('@') ? activeEmail.split('@')[1] : 'yourcompany.com';
  const rows = [
    ['COMPANY_DOMAIN', domain, 'Only email addresses from this domain can use the app.'],
    ['DUPLICATE_WINDOW_DAYS', '5', 'Look back this many days for possible duplicate tickets.'],
    ['DUPLICATE_SIMILARITY_THRESHOLD', '0.65', '0 to 1. Higher means stricter subject matching.'],
    ['RESOLVED_VISIBILITY_DAYS', '10', 'Sales can see their resolved tickets for this many days.'],
    ['DASHBOARD_WINDOW_DAYS', '14', 'Numbers dashboard lookback period.'],
    ['ALERT_PRIORITIES', 'HIGH', 'Legacy alert priority setting.'],
    ['MAX_ATTACHMENT_MB', '5', 'Maximum evidence file size.'],
    ['ROOT_CAUSES', 'Product Bug|Configuration Issue|Data Issue|Integration/API Issue|Access/Permission|User Error|External Dependency|Process Gap|Unable to Reproduce|Other', 'Values shown when resolving a ticket.']
  ];
  const existing = sheet.getLastRow()>1 ? sheet.getRange(2,1,sheet.getLastRow()-1,1).getDisplayValues().map(row=>String(row[0])) : [];
  const missing=rows.filter(row=>existing.indexOf(row[0])<0);
  if(missing.length)sheet.getRange(sheet.getLastRow()+1,1,missing.length,3).setValues(missing);
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

const TAXONOMY_GROUPS_ = Object.freeze([
  ['API & Integration',['New integration','API error','Webhook','Credentials','IP whitelisting']],
  ['Orders',['Order not created','Order not visible','Bulk upload','Cancellation','Incorrect order details']],
  ['Tracking & Status',['Tracking not available','Status delayed','Status incorrect','Webhook status missing']],
  ['Labels & Documents',['Label not generated','Incorrect label','Barcode not scanning','Missing label details']],
  ['Serviceability & Routing',['Pincode issue','Incorrect hub or node','Wrong service','Incorrect FPD or EDD']],
  ['Reverse, RTO & QC',['Reverse order issue','RTO issue','Pickup issue','QC issue','Exchange issue']],
  ['Billing, COD & Bank',['COD remittance','Bank verification','Wallet','Invoice','Incorrect charges']],
  ['Configuration Request',['Enable service','Enable webhook','Add pincode','Account mapping','Credentials']],
  ['Other',['Unable to identify','New requirement','General support']]
]);
const TAXONOMY_360_ONLY_ = Object.freeze([
  ['SF360 Portal',['Login or access','Page not loading','Feature not visible','Incorrect data or display','Other portal issue']],
  ['Shopify',['Store connection or onboarding','Store not found or domain issue','Owner email, OTP or authentication','Orders not fetched or imported','Order fulfilment or status sync','COD or payment status sync','Pickup location or warehouse mapping','Shopify label issue','Returns or reverse orders','Other Shopify issue']],
  ['WooCommerce',['Plugin connection or onboarding','Authentication or API credentials','Products not fetched','Orders not fetched or imported','Order creation or manifestation','Tracking, status or webhook sync','Pickup location or store mapping','WooCommerce label issue','Returns, RTO or NDR','Other WooCommerce issue']]
]);

function taxonomySlug_(value) { return String(value).toUpperCase().replace(/&/g,' AND ').replace(/[^A-Z0-9]+/g,'-').replace(/^-|-$/g,''); }
function taxonomySeedRows_() {
  const rows=[]; let order=1;
  ['360','Regular'].forEach(type=>TAXONOMY_GROUPS_.forEach(group=>group[1].forEach(name=>rows.push(taxonomyRow_(type,group[0],name,order++)))));
  TAXONOMY_360_ONLY_.forEach(group=>group[1].forEach(name=>rows.push(taxonomyRow_('360',group[0],name,order++))));
  return rows;
}
function taxonomyRow_(type,group,name,order) {
  const prefix=type==='360'?'360':'REG';
  const aliases={
    '360|Shopify|Store connection or onboarding':'360-SHOPIFY-STORE-CONNECTION','360|Shopify|Orders not fetched or imported':'360-SHOPIFY-ORDERS-NOT-FETCHED','360|Shopify|Order fulfilment or status sync':'360-SHOPIFY-STATUS-SYNC','360|Shopify|COD or payment status sync':'360-SHOPIFY-COD-SYNC','360|Shopify|Pickup location or warehouse mapping':'360-SHOPIFY-PICKUP-LOCATION','360|Shopify|Shopify label issue':'360-SHOPIFY-LABEL','360|Shopify|Returns or reverse orders':'360-SHOPIFY-RETURNS',
    '360|WooCommerce|Plugin connection or onboarding':'360-WOOCOMMERCE-CONNECTION','360|WooCommerce|Products not fetched':'360-WOOCOMMERCE-PRODUCTS-NOT-FETCHED','360|WooCommerce|Orders not fetched or imported':'360-WOOCOMMERCE-ORDERS-NOT-FETCHED','360|WooCommerce|Order creation or manifestation':'360-WOOCOMMERCE-ORDER-CREATION','360|WooCommerce|Tracking, status or webhook sync':'360-WOOCOMMERCE-STATUS-SYNC','360|WooCommerce|WooCommerce label issue':'360-WOOCOMMERCE-LABEL','360|WooCommerce|Returns, RTO or NDR':'360-WOOCOMMERCE-RETURNS'
  };
  const id=aliases[[type,group,name].join('|')]||[prefix,taxonomySlug_(group),taxonomySlug_(name)].join('-');
  const fields=['expected_result','actual_result','first_observed_at','impact_scope','environment','affected_count','affected_ids'];
  if(group==='Shopify') fields.push('store_url','registered_email','shopify_order_id','awb','expected_status','actual_status','pickup_location','error_message','attachment');
  else if(group==='WooCommerce') fields.push('store_url','plugin_version','product_sku','woocommerce_order_id','awb','expected_status','actual_status','error_message','attachment');
  else fields.push('attachment');
  const required=['expected_result','actual_result','first_observed_at','impact_scope','environment'];
  if(name==='API error'||name==='Bulk upload'||/label|barcode|page not loading|incorrect data|incorrect charges/i.test(name)) required.push('attachment');
  return [id,type,name,type==='360'?'MEDIUM':'MEDIUM',24,JSON.stringify(fields),JSON.stringify(required),true,group,name,true,order];
}

function seedMissingTaxonomyRows_(ss,summary) {
  const sheet=ss.getSheetByName(APP.SHEETS.CATEGORIES),headers=sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(String);
  const idCol=headers.indexOf('Category_ID'),existing=sheet.getLastRow()>1?sheet.getRange(2,idCol+1,sheet.getLastRow()-1,1).getDisplayValues().map(r=>String(r[0])):[];
  const canonical=APP.HEADERS.Categories,missing=taxonomySeedRows_().filter(row=>existing.indexOf(row[0])<0);
  if(missing.length)sheet.getRange(sheet.getLastRow()+1,1,missing.length,headers.length).setValues(missing.map(row=>headers.map(h=>row[canonical.indexOf(h)]===undefined?'':row[canonical.indexOf(h)])));
  if(summary)summary.configurationRowsSeeded.Taxonomy=missing.map(row=>row[0]);
  return missing.length;
}

function migrateRootCauses_(ss,summary) {
  const oldDefault='Product Bug|Configuration Issue|Data Issue|Integration/API Issue|Access/Permission|User Error|External Dependency|Process Gap|Unable to Reproduce|Other';
  const next='Client data/input issue|Client configuration issue|Shadowfax configuration issue|Product or technical defect|Third-party integration issue|Authentication or credential issue|Data synchronisation issue|Operations issue|Product limitation|Enhancement request|No issue found|Unable to reproduce|External dependency pending|Other';
  const sheet=ss.getSheetByName(APP.SHEETS.SETTINGS),rows=sheet.getDataRange().getValues(),headers=rows[0].map(String),key=headers.indexOf('Key'),value=headers.indexOf('Value');
  for(let i=1;i<rows.length;i++)if(String(rows[i][key])==='ROOT_CAUSES'&&String(rows[i][value])===oldDefault){sheet.getRange(i+1,value+1).setValue(next);if(summary)summary.configurationRowsSeeded.RootCauses=['replaced-known-default'];return true;}
  return false;
}

function category_(id, clientType, name, priority, slaHours, fields, required) {
  return [id, clientType, name, priority, slaHours, JSON.stringify(fields), JSON.stringify(required), true];
}

function clientSizePrioritySeedRows_() {
  return [
    ['GOLD_PLATINUM','Gold / Platinum','ADL 300 and above',300,'','HIGH',8,true,1],
    ['MEDIUM_SIZED','Medium-sized','ADL between 50 and 299',50,299,'MEDIUM',24,true,2],
    ['SMALL_SIZED','Small-sized','ADL below 50',0,49,'LOW',48,true,3]
  ];
}

function seedClientSizePriority_(ss) {
  const sheet = ss.getSheetByName(APP.SHEETS.CLIENT_SIZE_PRIORITY);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
  const existing = sheet.getLastRow() < 2 ? [] : sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const codeColumn = headers.indexOf('Client_Size_Code');
  const codes = existing.map(row => String(row[codeColumn] || ''));
  const missing = clientSizePrioritySeedRows_().filter(row => codes.indexOf(row[0]) < 0);
  if (missing.length) {
    const canonical = APP.HEADERS.ClientSizePriority;
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, headers.length)
      .setValues(missing.map(seed => headers.map(header => seed[canonical.indexOf(header)] === undefined ? '' : seed[canonical.indexOf(header)])));
  }
  // A deployment upgrading from the pre-SLA schema has existing size rows but
  // a newly appended, blank SLA_Hours column. Populate only those blanks;
  // configured values are never replaced.
  const slaColumn = headers.indexOf('SLA_Hours');
  if (slaColumn >= 0 && existing.length) {
    const defaults = clientSizePrioritySeedRows_().reduce((map, row) => { map[row[0]] = row[6]; return map; }, {});
    existing.forEach((row, index) => {
      const code = String(row[codeColumn] || '');
      if (defaults[code] !== undefined && String(row[slaColumn] == null ? '' : row[slaColumn]).trim() === '') {
        sheet.getRange(index + 2, slaColumn + 1).setValue(defaults[code]);
      }
    });
  }
  return missing.map(row => row[0]);
}

function formatSheets_(ss) {
  Object.values(APP.SHEETS).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || !sheet.getLastColumn()) return;
    const lastCol = sheet.getLastColumn();
    sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff').setWrap(true);
    sheet.setFrozenRows(1); sheet.setRowHeight(1, 32); sheet.autoResizeColumns(1, lastCol);
  });
  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.TICKETS), ['Created_At','SLA_Due_At','Picked_Up_At','Investigating_At','Resolved_At','Updated_At']);
  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.EVENTS), ['Created_At']);
  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.SLA_CYCLES), ['Started_At','Due_At','Ended_At','Created_At','Updated_At']);
  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.TICKET_INDEX), ['Created_At','SLA_Due_At','Resolved_At','Updated_At']);
  formatDateTimeColumns_(ss.getSheetByName(APP.SHEETS.SLACK_NOTIFICATIONS), ['Next_Attempt_At','Created_At','Processing_Started_At','Sent_At','Updated_At']);
}

function formatDateTimeColumns_(sheet, headerNames) {
  if (!sheet || !sheet.getLastColumn()) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(value => String(value).trim());
  headerNames.forEach(header => {
    const column = headers.indexOf(header) + 1;
    if (column) sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('dd-mmm-yyyy hh:mm');
  });
}

function newSchemaSummary_() {
  return { success: true, schemaVersion: CURRENT_SCHEMA_VERSION, sheetsCreated: [], columnsAdded: {}, configurationRowsSeeded: {}, ticketIndexRowsCreated: 0, warnings: [] };
}

/** Create a missing sheet/header and append absent columns without moving data. */
function ensureSheetSchema_(spreadsheet, sheetName, expectedHeaders, summary) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  let created = false;
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    created = true;
    if (summary) summary.sheetsCreated.push(sheetName);
  }
  const added = appendMissingColumns_(sheet, expectedHeaders);
  if (created || added.length) {
    sheet.setFrozenRows(1);
    const start = created ? 1 : Math.max(1, sheet.getLastColumn() - added.length + 1);
    const count = created ? sheet.getLastColumn() : added.length;
    if (count) sheet.getRange(1,start,1,count).setFontWeight('bold').setBackground('#17324d').setFontColor('#ffffff');
  }
  if (summary && added.length) summary.columnsAdded[sheetName] = (summary.columnsAdded[sheetName] || []).concat(added);
  return sheet;
}

function appendMissingColumns_(sheet, expectedHeaders) {
  const width = sheet.getLastColumn();
  const existing = width ? sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(value => String(value).trim()) : [];
  const missing = expectedHeaders.filter(header => existing.indexOf(header) < 0);
  if (missing.length) sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  return missing;
}

function seedMissingConfigurationRows_(spreadsheet, summary) {
  const seeded = seedClientSizePriority_(spreadsheet);
  const slackSeeded = seedMissingSlackSettings_(spreadsheet);
  if (summary) summary.configurationRowsSeeded.ClientSizePriority = seeded;
  if (summary) summary.configurationRowsSeeded.Slack = slackSeeded;
  return seeded;
}

function slackSettingSeedRows_() {
  return [
    ['SLACK_NOTIFICATIONS_ENABLED','TRUE','Enable or disable all Slack notifications.'],
    ['SLACK_ALERT_PRIORITIES','HIGH,CRITICAL','Ticket priorities that generate an immediate raised/reopened alert.'],
    ['SLACK_MENTION_PRIORITIES','HIGH,CRITICAL','Priorities for which Slack messages include <!here>.'],
    ['SLACK_BREACH_WARNING_ENABLED','TRUE','Enable the warning before SLA breach.'],
    ['SLACK_BREACH_WARNING_MINUTES','120','Working minutes before SLA breach at which a warning is generated.'],
    ['SLACK_BREACH_ALERT_ENABLED','TRUE','Enable notification after SLA has been breached.'],
    ['SLACK_EOD_SUMMARY_ENABLED','TRUE','Enable the weekday end-of-day summary.'],
    ['SLACK_EOD_HOUR','19','End-of-day summary hour in Asia/Kolkata.'],
    ['SLACK_EOD_MINUTE','45','End-of-day summary minute in Asia/Kolkata.'],
    ['SLACK_DISPATCH_BATCH_SIZE','10','Maximum pending Slack notifications processed in one dispatcher run.'],
    ['SLACK_MAX_RETRIES','3','Maximum Slack delivery attempts.'],
    ['SLACK_PROCESSING_TIMEOUT_MINUTES','10','Minutes after which an abandoned PROCESSING row can be retried.'],
    ['SLACK_NOTIFICATION_RETENTION_DAYS','30','Days to retain sent and permanently failed queue records.']
  ];
}

function seedMissingSlackSettings_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(APP.SHEETS.SETTINGS);
  if (!sheet) return [];
  const width = sheet.getLastColumn(), headers = sheet.getRange(1,1,1,width).getDisplayValues()[0].map(String);
  const keyColumn = headers.indexOf('Key');
  const existing = sheet.getLastRow() > 1 ? sheet.getRange(2,keyColumn+1,sheet.getLastRow()-1,1).getDisplayValues().map(row => String(row[0])) : [];
  const missing = slackSettingSeedRows_().filter(row => existing.indexOf(row[0]) < 0);
  if (missing.length) sheet.getRange(sheet.getLastRow()+1,1,missing.length,width).setValues(missing.map(seed => headers.map((header) => seed[['Key','Value','Description'].indexOf(header)] || '')));
  return missing.map(row => row[0]);
}

/** Every migration is additive and can safely repair a partially/manual migration. */
function runSchemaMigrations_(fromVersion, toVersion, spreadsheet, summary) {
  const ss = spreadsheet || getSpreadsheet_();
  const report = summary || newSchemaSummary_();
  const ensure = names => names.forEach(name => ensureSheetSchema_(ss, name, APP.HEADERS[name], report));
  if (fromVersion < 1 && toVersion >= 1) ensure([APP.SHEETS.TICKETS, APP.SHEETS.CLIENTS, APP.SHEETS.CATEGORIES, APP.SHEETS.USERS, APP.SHEETS.EVENTS, APP.SHEETS.SETTINGS]);
  if (fromVersion < 2 && toVersion >= 2) ensure([APP.SHEETS.SLA_CYCLES]);
  if (fromVersion < 3 && toVersion >= 3) { ensure([APP.SHEETS.CLIENT_SIZE_PRIORITY, APP.SHEETS.TICKETS]); seedMissingConfigurationRows_(ss, report); }
  if (fromVersion < 4 && toVersion >= 4) ensure([APP.SHEETS.TICKET_INDEX, APP.SHEETS.TICKETS, APP.SHEETS.EVENTS]);
  // Migration 5 deliberately rechecks the complete contract. This also makes
  // repair safe where a tab/column was manually created during an older release.
  if (toVersion >= 5) { ensure(Object.values(APP.SHEETS)); seedMissingConfigurationRows_(ss, report); }
  // Version 6 adds SLA source-of-truth columns. Additive header repair and
  // absent-row-only seeding preserve every administrator-edited value.
  if (toVersion >= 6) { ensure([APP.SHEETS.TICKETS, APP.SHEETS.CLIENT_SIZE_PRIORITY]); seedMissingConfigurationRows_(ss, report); }
  // Version 7 adds only operational Slack queue/configuration data. Existing
  // columns and rows remain in place and missing fields are appended.
  if (toVersion >= 7) { ensure([APP.SHEETS.SLACK_NOTIFICATIONS, APP.SHEETS.SETTINGS]); seedMissingConfigurationRows_(ss, report); }
  if (toVersion >= 8) { ensure([APP.SHEETS.CATEGORIES,APP.SHEETS.TICKETS,APP.SHEETS.TICKET_INDEX,APP.SHEETS.FEATURE_REQUESTS,APP.SHEETS.FEATURE_REQUEST_EVENTS]); seedMissingTaxonomyRows_(ss,report); migrateRootCauses_(ss,report); }
  const indexResult = backfillTicketIndexIfNeeded_(ss);
  report.ticketIndexRowsCreated += indexResult.created;
  if (!indexResult.complete) report.warnings.push('TicketIndex backfill will continue automatically.');
  return { summary: report, indexComplete: indexResult.complete };
}

/** Administrator-run, additive and idempotent taxonomy/feature-request upgrade. */
function upgradeTaxonomyAndFeatureRequests() {
  requireRole_([APP.ROLES.ADMIN]);
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const ss=getSpreadsheet_(),summary=newSchemaSummary_();
    runSchemaMigrations_(Number(PropertiesService.getScriptProperties().getProperty(APP_SCHEMA_VERSION_PROPERTY_)||0),CURRENT_SCHEMA_VERSION,ss,summary);
    PropertiesService.getScriptProperties().setProperty(APP_SCHEMA_VERSION_PROPERTY_,String(CURRENT_SCHEMA_VERSION)); invalidateSchemaCaches_();
    return {success:true,schemaVersion:CURRENT_SCHEMA_VERSION,sheetsCreated:summary.sheetsCreated,columnsAdded:summary.columnsAdded,rowsSeeded:summary.configurationRowsSeeded,warnings:summary.warnings};
  } finally { lock.releaseLock(); }
}

function schemaNamesReady_(spreadsheet) {
  const present = {};
  spreadsheet.getSheets().forEach(sheet => { present[sheet.getName()] = true; });
  return Object.values(APP.SHEETS).every(name => present[name]);
}

/** Fast startup guard: one sheet-name read plus the durable version property. */
function ensureRuntimeSchema_() {
  const startedAt = Date.now(), ss = getSpreadsheet_(), props = PropertiesService.getScriptProperties();
  const version = Number(props.getProperty(APP_SCHEMA_VERSION_PROPERTY_) || 0);
  // Cache is only a hint: required names are still checked so manual deletion is detected.
  let cachedReady = false;
  try { cachedReady = CacheService.getScriptCache().get(SCHEMA_READY_CACHE_KEY_) === '1'; } catch (cacheError) {}
  const namesReady = schemaNamesReady_(ss);
  if (version === CURRENT_SCHEMA_VERSION && namesReady) {
    if (!cachedReady) try { CacheService.getScriptCache().put(SCHEMA_READY_CACHE_KEY_, '1', 300); } catch (cacheError) {}
    return { ready: true, checked: false };
  }
  const lock = LockService.getScriptLock(), waitStarted = Date.now();
  if (!lock.tryLock(3000)) {
    console.log(JSON.stringify({ functionName: 'ensureRuntimeSchema', durationMs: Date.now() - startedAt, lockWaitMs: Date.now() - waitStarted, status: 'busy' }));
    const busy = new Error(SCHEMA_UPGRADE_IN_PROGRESS_ + ': ' + SCHEMA_TEMPORARY_MESSAGE_); busy.code = SCHEMA_UPGRADE_IN_PROGRESS_; throw busy;
  }
  try {
    const current = Number(props.getProperty(APP_SCHEMA_VERSION_PROPERTY_) || 0);
    if (current === CURRENT_SCHEMA_VERSION && schemaNamesReady_(ss)) return { ready: true, checked: false };
    const result = runSchemaMigrations_(current, CURRENT_SCHEMA_VERSION, ss, newSchemaSummary_());
    const validation = validateCompleteSchema_(ss);
    if (!validation.ready || !result.indexComplete) {
      const pending = new Error(SCHEMA_UPGRADE_IN_PROGRESS_ + ': ' + SCHEMA_TEMPORARY_MESSAGE_); pending.code = SCHEMA_UPGRADE_IN_PROGRESS_; throw pending;
    }
    props.setProperty(APP_SCHEMA_VERSION_PROPERTY_, String(CURRENT_SCHEMA_VERSION));
    invalidateSchemaCaches_();
    console.log(JSON.stringify({ functionName: 'ensureRuntimeSchema', durationMs: Date.now() - startedAt, lockWaitMs: Date.now() - waitStarted, status: 'repaired' }));
    return { ready: true, checked: true };
  } catch (error) {
    if (error && error.code === SCHEMA_UPGRADE_IN_PROGRESS_) throw error;
    console.error('Application schema repair failed: ' + String(error && error.message || error));
    throw new Error(SCHEMA_ADMIN_MESSAGE_);
  } finally { lock.releaseLock(); }
}

function getRequiredSheet_(sheetName) {
  let sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (sheet) return sheet;
  try { ensureRuntimeSchema_(); } catch (error) {
    if (error && (error.code === SCHEMA_UPGRADE_IN_PROGRESS_ || String(error.message).indexOf(SCHEMA_TEMPORARY_MESSAGE_) >= 0)) throw error;
    console.error('Required application sheet could not be prepared: ' + String(sheetName));
    throw new Error(SCHEMA_ADMIN_MESSAGE_);
  }
  sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) { console.error('Required application sheet remains unavailable: ' + String(sheetName)); throw new Error(SCHEMA_ADMIN_MESSAGE_); }
  return sheet;
}

function validateCompleteSchema_(spreadsheet) {
  const ss = spreadsheet || getSpreadsheet_(), missingSheets = [], missingColumns = {};
  Object.values(APP.SHEETS).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) { missingSheets.push(name); return; }
    const headers = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(value => String(value).trim()) : [];
    const missing = APP.HEADERS[name].filter(header => headers.indexOf(header) < 0);
    if (missing.length) missingColumns[name] = missing;
  });
  return { ready: !missingSheets.length && !Object.keys(missingColumns).length, missingSheets, missingColumns };
}

function validateApplicationSchema() {
  const ss = getSpreadsheet_(), validation = validateCompleteSchema_(ss), props = PropertiesService.getScriptProperties();
  const sizeSheet = ss.getSheetByName(APP.SHEETS.CLIENT_SIZE_PRIORITY), codes = {};
  if (sizeSheet && sizeSheet.getLastRow() > 1) {
    const headers = sizeSheet.getRange(1, 1, 1, sizeSheet.getLastColumn()).getDisplayValues()[0].map(String), codeIndex = headers.indexOf('Client_Size_Code');
    sizeSheet.getRange(2, 1, sizeSheet.getLastRow() - 1, sizeSheet.getLastColumn()).getValues().forEach(row => { codes[String(row[codeIndex])] = true; });
  }
  const requiredCodes = clientSizePrioritySeedRows_().map(row => row[0]);
  return {
    schemaVersion: Number(props.getProperty(APP_SCHEMA_VERSION_PROPERTY_) || 0), requiredSheetsPresent: Object.values(APP.SHEETS).filter(name => validation.missingSheets.indexOf(name) < 0),
    requiredSheetsMissing: validation.missingSheets, missingColumns: validation.missingColumns,
    clientSizeSeedRowsPresent: requiredCodes.filter(code => codes[code]), clientSizeSeedRowsMissing: requiredCodes.filter(code => !codes[code]),
    ticketIndexRowCount: safeDataRowCount_(ss.getSheetByName(APP.SHEETS.TICKET_INDEX)), ticketsRowCount: safeDataRowCount_(ss.getSheetByName(APP.SHEETS.TICKETS)),
    cachesInvalidated: false, schemaReady: validation.ready && Number(props.getProperty(APP_SCHEMA_VERSION_PROPERTY_)) === CURRENT_SCHEMA_VERSION
  };
}

function safeDataRowCount_(sheet) { return sheet ? Math.max(0, sheet.getLastRow() - 1) : 0; }

function repairApplicationSchema() {
  const ss = getSpreadsheet_(), props = PropertiesService.getScriptProperties(), lock = LockService.getScriptLock(), startedAt = Date.now(), waitStarted = Date.now(), summary = newSchemaSummary_();
  lock.waitLock(10000);
  const lockWaitMs = Date.now() - waitStarted;
  try {
    const from = Number(props.getProperty(APP_SCHEMA_VERSION_PROPERTY_) || 0);
    const result = runSchemaMigrations_(from, CURRENT_SCHEMA_VERSION, ss, summary);
    const validation = validateCompleteSchema_(ss);
    if (!validation.ready || !result.indexComplete) throw new Error('Schema or TicketIndex validation is not complete. Retry after the continuation finishes.');
    props.setProperty(APP_SCHEMA_VERSION_PROPERTY_, String(CURRENT_SCHEMA_VERSION));
    invalidateSchemaCaches_();
    console.log(JSON.stringify({ functionName: 'repairApplicationSchema', durationMs: Date.now() - startedAt, lockWaitMs: lockWaitMs }));
    return summary;
  } catch (error) { console.error('Administrator schema repair failed: ' + String(error && error.message || error)); summary.success = false; summary.warnings.push(SCHEMA_ADMIN_MESSAGE_); return summary; }
  finally { lock.releaseLock(); }
}

function backfillTicketIndexIfNeeded_(ss) {
  const source = ss.getSheetByName(APP.SHEETS.TICKETS), target = ss.getSheetByName(APP.SHEETS.TICKET_INDEX);
  if (!source || !target || target.getLastRow() > 1 || source.getLastRow() < 2) return { created: 0, complete: true };
  return continueTicketIndexBackfill_(ss);
}

function continueTicketIndexBackfill_(spreadsheet) {
  const ss = spreadsheet || getSpreadsheet_(), props = PropertiesService.getScriptProperties(), source = ss.getSheetByName(APP.SHEETS.TICKETS), target = ss.getSheetByName(APP.SHEETS.TICKET_INDEX);
  const sourceHeaders = source.getRange(1, 1, 1, source.getLastColumn()).getDisplayValues()[0].map(String), targetHeaders = target.getRange(1, 1, 1, target.getLastColumn()).getDisplayValues()[0].map(String);
  let nextRow = Number(props.getProperty('TICKET_INDEX_BACKFILL_NEXT_ROW') || 2);
  // If another completed index exists, never duplicate it merely because progress is stale.
  if (target.getLastRow() > 1 && nextRow === 2) return { created: 0, complete: true };
  const count = Math.min(1000, Math.max(0, source.getLastRow() - nextRow + 1));
  if (count) {
    const values = source.getRange(nextRow, 1, count, sourceHeaders.length).getValues();
    const rows = values.map(row => { const ticket = sourceHeaders.reduce((object, header, i) => { object[header] = row[i]; return object; }, {}); const item = ticketToIndex_(ticket); return targetHeaders.map(header => item[header] === undefined ? '' : item[header]); });
    target.getRange(target.getLastRow() + 1, 1, rows.length, targetHeaders.length).setValues(rows); nextRow += count;
  }
  const complete = nextRow > source.getLastRow();
  if (complete) props.deleteProperty('TICKET_INDEX_BACKFILL_NEXT_ROW');
  else { props.setProperty('TICKET_INDEX_BACKFILL_NEXT_ROW', String(nextRow)); scheduleTicketIndexContinuation_(); }
  return { created: count, complete };
}

function scheduleTicketIndexContinuation_() {
  const exists = ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === 'continueTicketIndexBackfill');
  if (!exists) ScriptApp.newTrigger('continueTicketIndexBackfill').timeBased().after(60 * 1000).create();
}

function continueTicketIndexBackfill() {
  const lock = LockService.getScriptLock(); if (!lock.tryLock(5000)) return { complete: false };
  try {
    const result = continueTicketIndexBackfill_(getSpreadsheet_());
    if (result.complete) {
      PropertiesService.getScriptProperties().setProperty(APP_SCHEMA_VERSION_PROPERTY_, String(CURRENT_SCHEMA_VERSION));
      ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction() === 'continueTicketIndexBackfill').forEach(trigger => ScriptApp.deleteTrigger(trigger));
      invalidateSchemaCaches_();
    }
    return result;
  } finally { lock.releaseLock(); }
}

function invalidateSchemaCaches_() {
  try { CacheService.getScriptCache().removeAll([SCHEMA_READY_CACHE_KEY_, 'app:settings:v2', 'app:categories:v2', 'app:client-size-priority:v1', 'app:numbers:v2']); } catch (error) { console.warn('Schema cache invalidation was unavailable.'); }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Internal Ticketing Admin')
    .addItem('Repair / Upgrade Application Structure', 'repairApplicationSchema')
    .addItem('Validate Application Structure', 'validateApplicationSchema')
    .addItem('Rebuild Ticket Index', 'rebuildTicketIndex')
    .addSeparator()
    .addItem('Validate Slack Automation', 'validateSlackAutomation')
    .addItem('Install / Repair Slack Triggers', 'setupSlackAutomationTriggers')
    .addItem('Send Slack Test', 'testSlackConnection')
    .addItem('Remove Slack Triggers', 'removeSlackAutomationTriggers')
    .addToUi();
}

/** Backward-compatible upgrade names now use the complete generic repair. */
function upgradeSlaCycleSchema() { return repairApplicationSchema(); }
function upgradeClientSizeAndPerformanceSchema() { return repairApplicationSchema(); }
