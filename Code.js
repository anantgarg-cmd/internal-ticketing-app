/**
 * MAIN BACKEND FILE
 * Do not put passwords or the Slack webhook directly in this file.
 */

const APP_RELEASE = 'sales-tickets-search-loader-v1';
const APP_COMMIT = '__APP_COMMIT__';
let SPREADSHEET_INSTANCE_ = null;
let SLA_SCHEMA_RECOVERY_FAILED_ = false;
const DEPLOYMENT_AUTHORIZATION_MESSAGE = 'The application deployment has not been authorized by its deploying account. Please ask the application administrator to run authorizeApplication() once from the Apps Script editor.';
const CACHE_KEYS_ = Object.freeze({
  SETTINGS: 'app:settings:v2',
  CATEGORIES: 'app:categories:v2',
  CLIENT_SIZES: 'app:client-size-priority:v1',
  NUMBERS: 'app:numbers:v2'
});

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * One-time deployment-owner authorization. Run manually from the Apps Script
 * editor; this function is intentionally not part of web-app startup.
 */
function authorizeApplication() {
  ScriptApp.requireAllScopes(ScriptApp.AuthMode.FULL);

  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = properties.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID is missing from Script Properties. Configure the application before authorizing it.');
  }

  SpreadsheetApp.openById(spreadsheetId).getName();
  DriveApp.getRootFolder().getId();
  UrlFetchApp.fetch('https://www.google.com/generate_204', { muteHttpExceptions: true });

  // Reading the trigger collection is a safe authorization diagnostic; it does not create a trigger.
  ScriptApp.getProjectTriggers();

  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email.endsWith('@shadowfax.in')) {
    throw new Error('Authorization must be completed by a @shadowfax.in deployment-owner account.');
  }

  return {
    authorized: true,
    spreadsheetAccessible: true,
    driveAccessible: true,
    externalRequestAccessible: true,
    triggerManagementAuthorized: true,
    companyDomainValid: true,
    timestamp: new Date().toISOString()
  };
}

/** Returns non-sensitive authorization state for manual editor diagnostics. */
function getAuthorizationDiagnostic() {
  const info = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  const status = info.getAuthorizationStatus();
  return {
    authorizationStatus: status === ScriptApp.AuthorizationStatus.REQUIRED ? 'REQUIRED' : 'NOT_REQUIRED',
    authorizationUrlPresent: Boolean(info.getAuthorizationUrl()),
    timestamp: new Date().toISOString()
  };
}

function getBootstrap() {
  const user = requireUser_();
  return buildBootstrap_(user);
}



function checkDuplicate(payload) {
  const startedAt = Date.now();
  const user = requireUser_();
  validateDuplicatePayload_(payload);
  const settings = getSettings_();
  const days = number_(settings.DUPLICATE_WINDOW_DAYS, 5);
  const threshold = number_(settings.DUPLICATE_SIMILARITY_THRESHOLD, 0.65);
  const cutoff = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  const clientKey = buildClientKeyFromPayload_(payload);
  const normalizedSubject = normalizeSubject_(payload.emailSubject);

  const recent = getRecentTicketObjects_(cutoff, 200);
  const roleMap = user.role === APP.ROLES.SALES ? getUserRoleMap_() : null;
  const raiserByTicketId = user.role === APP.ROLES.SALES ? getSheetObjects_(APP.SHEETS.TICKETS).reduce((raisers, ticket) => {
    raisers[String(ticket.Ticket_ID)] = lower_(ticket.Raiser_Email);
    return raisers;
  }, {}) : null;
  const matches = recent.rows
    .filter(t => buildClientKeyFromTicket_(t) === clientKey)
    .filter(t => String(t.Category_ID) === String(payload.categoryId))
    .map(t => ({ ticket: t, similarity: subjectSimilarity_(normalizedSubject, String(t.Normalized_Subject || '')) }))
    .filter(x => x.similarity >= threshold)
    .sort((a, b) => toDate_(b.ticket.Created_At) - toDate_(a.ticket.Created_At))
    .slice(0, 3)
    .map(x => {
      const raiserEmail = user.role === APP.ROLES.SALES ? raiserByTicketId[String(x.ticket.Ticket_ID)] : '';
      const canView = user.role !== APP.ROLES.SALES || raiserEmail === user.email || isSalesRaisedTicket_({ Raiser_Email: raiserEmail }, roleMap);
      return {
        ticketId: canView ? String(x.ticket.Ticket_ID) : '',
        createdAt: canView ? formatDateTime_(x.ticket.Created_At) : '',
        status: canView ? String(x.ticket.Status) : 'Existing ticket',
        subject: canView ? String(x.ticket.Email_Subject) : 'A protected similar ticket already exists.',
        similarity: Math.round(x.similarity * 100),
        canView
      };
    });

  logPerformance_('checkDuplicate', startedAt, { rows: recent.processed });
  return { hasDuplicate: matches.length > 0, matches };
}


/** Returns a server-calculated SLA estimate for the Raise Ticket form. */
function getSlaDuePreview(categoryId, clientSizeCode) {
  requireUser_();
  const category = getCategoryById_(categoryId);
  if (!category) throw new Error('The selected category is no longer active. Refresh the page and choose again.');
  const resolution = resolveTicketPriority_(category.Client_Type, clientSizeCode, category);
  const dueAt = calculateWorkingSlaDueAt_(new Date(), resolution.slaHours);
  return { dueAt: formatDateTime_(dueAt), dueAtIso: dueAt.toISOString(), priority: resolution.priority, slaHours: resolution.slaHours };
}





function getReopenPreview(ticketId) {
  const user = requireUser_();
  const found = findObjectRow_(APP.SHEETS.TICKETS, 'Ticket_ID', ticketId);
  if (!found) throw new Error('Ticket not found.');
  assertCanReopen_(user, found.object);
  const dueAt = calculateWorkingSlaDueAt_(new Date(), number_(found.object.SLA_Hours, 0));
  return { dueAt: formatDateTime_(dueAt), dueAtIso: dueAt.toISOString() };
}

/** Reopens one resolved ticket and starts an independent SLA cycle. */

function assertCanReopen_(user, ticket) {
  if (String(ticket.Status) !== APP.STATUS.RESOLVED) throw new Error('Only a Resolved ticket can be reopened.');
  if (![APP.ROLES.POC, APP.ROLES.ADMIN].includes(user.role) && lower_(ticket.Raiser_Email) !== user.email) {
    throw new Error('You are not allowed to reopen another user\'s ticket.');
  }
}


function setWebAppUrlFromEditor() {
  // After deployment, paste the /exec URL here, run once, then remove it from the code.
  const webAppUrl = 'PASTE_DEPLOYED_EXEC_URL_HERE';
  if (!webAppUrl.includes('/exec')) throw new Error('Paste the deployed web app URL ending in /exec.');
  PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', webAppUrl);
  return 'Web app URL saved.';
}

// -------------------- Authentication and access --------------------

function getVerifiedCompanyEmail_() {
  let email;
  try {
    email = lower_(Session.getActiveUser().getEmail());
  } catch (err) {
    throwServiceAuthorizationError_(err);
  }
  if (!email) {
    throw new Error('Your email could not be identified. Open this app using your company Google Workspace account. The deployment must be restricted to your organisation.');
  }

  const settings = getSettings_();
  const domain = lower_(settings.COMPANY_DOMAIN);
  if (!domain) throw new Error('COMPANY_DOMAIN is not configured in the Settings sheet. Please contact the application administrator.');
  if (!email.endsWith('@' + domain)) throw new Error('This application is restricted to the company domain.');
  return email;
}

function getEntryState() {
  const email = getVerifiedCompanyEmail_();
  const row = getSheetObjects_(APP.SHEETS.USERS).find(u => lower_(u.Email) === email);
  if (!row) return { state: 'REGISTER', email: email, release: APP_RELEASE };
  if (!truthy_(row.Active)) {
    return {
      state: 'BLOCKED',
      email,
      message: 'Your access has been disabled. Please contact the application administrator.',
      release: APP_RELEASE
    };
  }
  return {
    state: 'ACTIVE',
    email,
    name: String(row.Name || email.split('@')[0]),
    role: String(row.Role || '').toUpperCase(),
    release: APP_RELEASE
  };
}

function getRuntimeDiagnostics() {
  const settings = getSettings_();
  return {
    scriptId: ScriptApp.getScriptId(),
    serviceUrl: ScriptApp.getService().getUrl(),
    release: APP_RELEASE,
    commit: APP_COMMIT === '__APP_COMMIT__' ? '' : APP_COMMIT,
    hasGetEntryState: typeof getEntryState === 'function',
    timestamp: new Date().toISOString(),
    companyDomain: lower_(settings.COMPANY_DOMAIN),
    usersSheetExists: Boolean(getSpreadsheet_().getSheetByName(APP.SHEETS.USERS))
  };
}

function registerFirstTimeUser(payload) {
  const email = getVerifiedCompanyEmail_();
  if (!payload || typeof payload !== 'object') throw new Error('Registration details were not received. Please try again.');

  const name = cleanText_(payload.name, 200);
  if (name.length < 2) throw new Error('Full Name must contain at least two characters.');

  const role = String(payload.role || '').trim().toUpperCase();
  if (![APP.ROLES.SALES, APP.ROLES.POC].includes(role)) {
    throw new Error('Choose either Sales or Tech/Product. ADMIN cannot be selected during registration.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let registeredRow;
  try {
    const existing = getSheetObjects_(APP.SHEETS.USERS).find(u => lower_(u.Email) === email);
    if (!existing) {
      registeredRow = { Email: email, Name: name, Role: role, Active: true };
      appendObject_(APP.SHEETS.USERS, registeredRow);
    } else {
      registeredRow = existing;
    }
  } finally {
    lock.releaseLock();
  }
  if (!truthy_(registeredRow.Active)) return { state: 'BLOCKED', email, message: 'Your access has been disabled. Please contact the application administrator.', release: APP_RELEASE };
  const user = userFromRow_(registeredRow, email);
  return { state: 'ACTIVE', email, name: user.name, role: user.role, bootstrap: buildBootstrap_(user), release: APP_RELEASE };
}

function userFromRow_(row, email) {
  const role = String(row.Role || '').toUpperCase();
  if (!Object.values(APP.ROLES).includes(role)) throw new Error('Your role in the Users sheet is invalid. Use SALES, POC or ADMIN.');
  return { email, name: String(row.Name || email.split('@')[0]), role };
}


function requireRole_(roles) {
  const user = requireUser_();
  if (!roles.includes(user.role)) throw new Error('You are not allowed to perform this action.');
  return user;
}

// -------------------- Data validation --------------------

function validateDuplicatePayload_(payload) {
  if (!payload) throw new Error('Ticket data is missing.');
  const mode = String(payload.clientMode || '');
  const clientId = String(payload.clientId == null ? '' : payload.clientId).trim();
  const clientName = cleanText_(payload.clientName, 200);
  const clientType = String(payload.clientType || '');
  if (!['existing','new'].includes(mode)) throw new Error('Choose a valid client status.');
  if (!clientName) throw new Error('Client name is mandatory.');
  if (!['360','Regular'].includes(clientType)) throw new Error('Choose 360 or Regular for the client.');
  if (mode === 'existing' && !clientId) throw new Error('Client ID is mandatory for an existing client.');
  if (clientId && (!/^\d+$/.test(clientId) || clientId.length > 20)) throw new Error('Client ID must contain only digits and be no more than 20 characters.');
  if (!cleanText_(payload.emailSubject, 300)) throw new Error('Email subject is mandatory.');
  if (!payload.categoryId) throw new Error('Category is mandatory.');
}

function validateTicketForm_(form, category, client) {
  if (!client.name) throw new Error('Client is mandatory.');
  if (!cleanText_(form.emailSubject, 300)) throw new Error('Email subject is mandatory.');
  if (!cleanText_(form.issueDescription, 5000)) throw new Error('Issue description is mandatory.');

  const visibleFields = safeJsonParse_(category.Fields_JSON, []);
  const required = safeJsonParse_(category.Required_Fields_JSON, []);
  required.forEach(field => {
    if (field === 'attachment') {
      if (!hasFile_(form.attachment)) throw new Error('An attachment is required for this category.');
    } else if (!cleanText_(form[field], 5000)) {
      throw new Error(`${fieldLabel_(field)} is mandatory for this category.`);
    }
  });

  Object.keys(form).forEach(key => {
    if (key.startsWith('dynamic_') && !visibleFields.includes(key.replace('dynamic_', ''))) {
      throw new Error('Unexpected field received. Refresh the page and try again.');
    }
  });
}

function resolveClient_(form, expectedClientType) {
  const mode = String(form.clientMode || '');
  const id = String(form.clientId == null ? '' : form.clientId).trim();
  const name = cleanText_(form.clientName, 200);
  const expectedType = String(expectedClientType == null ? '' : expectedClientType).trim();
  const type = String(form.clientType || '').trim();
  if (!['existing','new'].includes(mode)) throw new Error('Choose a valid client status.');
  if (!name) throw new Error('Client name is mandatory.');
  if (mode === 'existing' && !id) throw new Error('Client ID is mandatory for an existing client.');
  if (id && (!/^\d+$/.test(id) || id.length > 20)) throw new Error('Client ID must contain only digits and be no more than 20 characters.');
  if (!['360','Regular'].includes(type)) throw new Error('Choose 360 or Regular for the client.');
  if (type !== expectedType) throw new Error('The selected category does not match the client type.');
  return { mode: mode === 'existing' ? 'Existing' : 'New', id, name, type };
}

// -------------------- Attachments --------------------

function saveAttachment_(blob, ticketId, raiserEmail) {
  if (!hasFile_(blob)) return { id: '', name: '', url: '' };

  const settings = getSettings_();
  const maxBytes = number_(settings.MAX_ATTACHMENT_MB, 5) * 1024 * 1024;
  const bytes = blob.getBytes();
  if (bytes.length > maxBytes) throw new Error(`Attachment is too large. Maximum size is ${settings.MAX_ATTACHMENT_MB || 5} MB.`);

  const allowed = ['image/png','image/jpeg','application/pdf','text/plain','application/json','application/zip','application/x-zip-compressed'];
  const contentType = String(blob.getContentType() || '').toLowerCase();
  if (!allowed.includes(contentType)) throw new Error('Allowed attachment types: PNG, JPG, PDF, TXT, JSON or ZIP.');

  const folderId = PropertiesService.getScriptProperties().getProperty('ATTACHMENT_FOLDER_ID');
  if (!folderId) throw new Error('Attachment storage is not configured. Please contact the application administrator.');
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (err) {
    throwServiceAuthorizationError_(err);
  }
  const original = cleanFileName_(blob.getName() || 'evidence');
  blob.setName(`${ticketId} - ${original}`);
  let file;
  try {
    file = folder.createFile(blob);
    file.setDescription(`Evidence for ${ticketId}. Uploaded by ${raiserEmail}.`);
    file.addViewer(raiserEmail);
  } catch (err) {
    throwServiceAuthorizationError_(err);
  }
  return { id: file.getId(), name: file.getName(), url: file.getUrl() };
}

function hasFile_(blob) {
  return blob && typeof blob.getBytes === 'function' && blob.getBytes().length > 0;
}

// -------------------- Slack --------------------

function sendSlackAlert_(ticket) {
  if (!slackPriorityConfigured_(ticket.priority, 'SLACK_ALERT_PRIORITIES')) return { enqueued: false, duplicate: false };
  return enqueueSlackNotification_({ notificationType: 'NEW_HIGH_PRIORITY', dedupeKey: 'NEW_HIGH:' + ticket.ticketId,
    ticketId: ticket.ticketId, slaCycleNumber: number_(ticket.slaCycleNumber, 1), priority: ticket.priority,
    payload: buildNewHighPrioritySlackPayload_(ticket) });
}

function sendSlackReopenedAlert_(ticket, reopenedBy, reason) {
  if (!slackPriorityConfigured_(ticket.priority, 'SLACK_ALERT_PRIORITIES')) return { enqueued: false, duplicate: false };
  const cycle = number_(ticket.slaCycleNumber, 1);
  return enqueueSlackNotification_({ notificationType: 'REOPENED_HIGH_PRIORITY', dedupeKey: 'REOPEN_HIGH:' + ticket.ticketId + ':' + cycle,
    ticketId: ticket.ticketId, slaCycleNumber: cycle, priority: ticket.priority,
    payload: buildReopenedHighPrioritySlackPayload_(ticket, reopenedBy, reason) });
}

// -------------------- Sheet helpers --------------------

function getSpreadsheet_() {
  if (SPREADSHEET_INSTANCE_) return SPREADSHEET_INSTANCE_;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Spreadsheet access is not configured. Please contact the application administrator.');
  try {
    SPREADSHEET_INSTANCE_ = SpreadsheetApp.openById(id);
  } catch (err) {
    throwServiceAuthorizationError_(err);
  }
  return SPREADSHEET_INSTANCE_;
}

/** Converts only platform authorization failures; business errors remain intact. */
function throwServiceAuthorizationError_(err) {
  const message = String(err && err.message ? err.message : err || '');
  const isAuthorizationFailure = /authoriz|permission|insufficient|scope|access denied|not have access|credentials|oauth/i.test(message);
  if (isAuthorizationFailure) throw new Error(DEPLOYMENT_AUTHORIZATION_MESSAGE);
  throw err;
}

function getSettings_() {
  const cached = getCachedJson_(CACHE_KEYS_.SETTINGS);
  if (cached !== null) return cached;
  const settings = {};
  getSheetObjects_(APP.SHEETS.SETTINGS).forEach(r => { settings[String(r.Key)] = String(r.Value); });
  putCachedJson_(CACHE_KEYS_.SETTINGS, settings, 300);
  return settings;
}

function getActiveClients_() {
  return getSheetObjects_(APP.SHEETS.CLIENTS)
    .filter(r => truthy_(r.Active))
    .map(r => ({ id: String(r.Client_ID), name: String(r.Client_Name), type: String(r.Client_Type) }))
    .filter(c => c.id && c.name && ['360','Regular'].includes(c.type))
    .sort((a, b) => a.name.localeCompare(b.name));
}

function getActiveCategories_() {
  const cached = getCachedJson_(CACHE_KEYS_.CATEGORIES);
  if (cached !== null) return cached;
  const categories = getSheetObjects_(APP.SHEETS.CATEGORIES)
    .filter(r => truthy_(r.Active))
    .map(r => ({
      id: String(r.Category_ID),
      clientType: String(r.Client_Type),
      name: String(r.Category_Name),
      priority: String(r.Priority).toUpperCase(),
      slaHours: number_(r.SLA_Hours, 24),
      fields: safeJsonParse_(r.Fields_JSON, []),
      requiredFields: safeJsonParse_(r.Required_Fields_JSON, [])
    }));
  putCachedJson_(CACHE_KEYS_.CATEGORIES, categories, 300);
  return categories;
}


/** Write paths must fail clearly rather than silently losing cycle history. */
function requireTicketSlaCyclesSchemaForWrite_(lockAlreadyHeld) {
  try {
    const sheet = getSpreadsheet_().getSheetByName(APP.SHEETS.SLA_CYCLES);
    if (sheet) return { ready: true };
    if (lockAlreadyHeld) throw new Error(SCHEMA_ADMIN_MESSAGE_);
    getRequiredSheet_(APP.SHEETS.SLA_CYCLES);
    return { ready: true };
  } catch (err) {
    console.error('SLA-cycle write schema preparation failed: ' + String(err && err.message || err));
    throw new Error(SCHEMA_ADMIN_MESSAGE_);
  }
}

function getSheetObjects_(sheetName) {
  const sheet = getRequiredSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(String);
  return values.slice(1)
    .filter(row => row.some(v => v !== '' && v !== null))
    .map(row => headers.reduce((obj, header, i) => { obj[header] = row[i]; return obj; }, {}));
}

function appendObject_(sheetName, object) {
  const sheet = getRequiredSheet_(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length)
    .setValues([headers.map(h => Object.prototype.hasOwnProperty.call(object, h) ? object[h] : '')]);
}

function findObjectRow_(sheetName, key, value) {
  const sheet = getRequiredSheet_(sheetName);
  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || lastColumn < 1) return null;
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(String);
  const keyIndex = headers.indexOf(key);
  if (keyIndex < 0) throw new Error(`Column ${key} not found in ${sheetName}.`);
  const keys = sheet.getRange(2, keyIndex + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0]) === String(value)) {
      const row = sheet.getRange(i + 2, 1, 1, lastColumn).getValues()[0];
      const object = headers.reduce((obj, h, j) => { obj[h] = row[j]; return obj; }, {});
      return { rowNumber: i + 2, object };
    }
  }
  return null;
}

function updateObjectRow_(sheetName, rowNumber, changes) {
  const sheet = getRequiredSheet_(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  Object.keys(changes).forEach(key => { if (headers.indexOf(key) < 0) throw new Error(`Column ${key} not found in ${sheetName}.`); });
  const range = sheet.getRange(rowNumber, 1, 1, headers.length);
  const row = range.getValues()[0];
  Object.keys(changes).forEach(key => { row[headers.indexOf(key)] = changes[key]; });
  range.setValues([row]);
}


function appendSlaCycle_(cycle) {
  appendObject_(APP.SHEETS.SLA_CYCLES, Object.assign({ SLA_Cycle_ID: Utilities.getUuid() }, cycle));
}


function ensureLegacyInitialCycle_(ticket) {
  if (getTicketSlaCycles_(ticket.Ticket_ID).length) return false;
  const result = ['MET', 'BREACHED'].includes(String(ticket.SLA_Result))
    ? String(ticket.SLA_Result)
    : (toDate_(ticket.Resolved_At) <= toDate_(ticket.SLA_Due_At) ? 'MET' : 'BREACHED');
  appendSlaCycle_({
    Ticket_ID: ticket.Ticket_ID, Cycle_Number: 1, Cycle_Type: 'INITIAL', Started_At: ticket.Created_At,
    Due_At: ticket.SLA_Due_At, Ended_At: ticket.Resolved_At, SLA_Result: result,
    Started_By: ticket.Raiser_Email, Ended_By: ticket.Resolved_By, Created_At: ticket.Created_At,
    Updated_At: ticket.Resolved_At || ticket.Created_At
  });
  return true;
}

function ensureOpenInitialCycle_(ticket) {
  if (getTicketSlaCycles_(ticket.Ticket_ID).length) return false;
  appendSlaCycle_({
    Ticket_ID: ticket.Ticket_ID, Cycle_Number: 1, Cycle_Type: 'INITIAL', Started_At: ticket.Created_At,
    Due_At: ticket.SLA_Due_At, SLA_Result: 'OPEN', Started_By: ticket.Raiser_Email,
    Created_At: ticket.Created_At, Updated_At: ticket.Created_At
  });
  return true;
}

function closeOpenSlaCycle_(ticketId, endedAt, endedBy, result) {
  const cycles = getTicketSlaCycles_(ticketId).filter(cycle => String(cycle.SLA_Result) === 'OPEN');
  if (cycles.length !== 1) throw new Error(`Expected exactly one OPEN SLA cycle; found ${cycles.length}.`);
  const id = cycles[0].SLA_Cycle_ID;
  const found = findObjectRow_(APP.SHEETS.SLA_CYCLES, 'SLA_Cycle_ID', id);
  updateObjectRow_(APP.SHEETS.SLA_CYCLES, found.rowNumber, { Ended_At: endedAt, Ended_By: endedBy, SLA_Result: result, Updated_At: endedAt });
}

function serializeSlaCycle_(cycle) {
  return {
    cycleNumber: number_(cycle.Cycle_Number, 0), cycleType: String(cycle.Cycle_Type),
    startedAt: formatDateTimeOptional_(cycle.Started_At), dueAt: formatDateTimeOptional_(cycle.Due_At),
    endedAt: formatDateTimeOptional_(cycle.Ended_At), slaResult: String(cycle.SLA_Result),
    startedBy: String(cycle.Started_By || ''), endedBy: String(cycle.Ended_By || ''), reopenReason: String(cycle.Reopen_Reason || '')
  };
}

/** Optional admin-only, on-demand legacy migration. Never called automatically. */
function backfillTicketSlaCycles() {
  requireRole_([APP.ROLES.ADMIN]);
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  let created = 0;
  try {
    getSheetObjects_(APP.SHEETS.TICKETS).forEach(ticket => {
      if (!getTicketSlaCycles_(ticket.Ticket_ID).length) {
        const resolved = String(ticket.Status) === APP.STATUS.RESOLVED;
        appendSlaCycle_({
          Ticket_ID: ticket.Ticket_ID, Cycle_Number: 1, Cycle_Type: 'INITIAL', Started_At: ticket.Created_At,
          Due_At: ticket.SLA_Due_At, Ended_At: resolved ? ticket.Resolved_At : '',
          SLA_Result: resolved ? (String(ticket.SLA_Result) || (toDate_(ticket.Resolved_At) <= toDate_(ticket.SLA_Due_At) ? 'MET' : 'BREACHED')) : 'OPEN',
          Started_By: ticket.Raiser_Email, Ended_By: resolved ? ticket.Resolved_By : '', Created_At: ticket.Created_At,
          Updated_At: resolved ? (ticket.Resolved_At || ticket.Created_At) : ticket.Created_At
        }); created++;
      }
    });
  } finally { lock.releaseLock(); }
  return { created };
}

function nextTicketId_() {
  const dateKey = Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd');
  const props = PropertiesService.getScriptProperties();
  const key = `TICKET_COUNTER_${dateKey}`;
  const next = number_(props.getProperty(key), 0) + 1;
  props.setProperty(key, String(next));
  return `TKT-${dateKey}-${String(next).padStart(4, '0')}`;
}

// -------------------- Serialisation and calculations --------------------

const SLA_WORK_START_HOUR_ = 11;
const SLA_WORK_START_MINUTE_ = 30;
const SLA_WORK_END_HOUR_ = 19;
const SLA_WORK_END_MINUTE_ = 30;

/** Builds an instant from calendar fields in the application's configured timezone. */
function dateInAppTimezone_(year, month, day, hour, minute, second, millisecond) {
  const approximate = new Date(Date.UTC(year, month - 1, day, hour, minute, second || 0, millisecond || 0));
  const offsetText = Utilities.formatDate(approximate, APP.TZ, 'Z');
  const sign = offsetText.charAt(0) === '-' ? -1 : 1;
  const offsetMinutes = sign * (number_(offsetText.slice(1, 3), 0) * 60 + number_(offsetText.slice(3, 5), 0));
  return new Date(approximate.getTime() - offsetMinutes * 60000);
}

function appDateParts_(date) {
  return Utilities.formatDate(toDate_(date), APP.TZ, 'yyyy,M,d,H,m,s,S').split(',').map(Number);
}

function isWorkingDay_(date) {
  return !['Sat', 'Sun'].includes(Utilities.formatDate(toDate_(date), APP.TZ, 'EEE'));
}

function getWorkingDayStart_(date) {
  const p = appDateParts_(date);
  return dateInAppTimezone_(p[0], p[1], p[2], SLA_WORK_START_HOUR_, SLA_WORK_START_MINUTE_, 0, 0);
}

function getWorkingDayEnd_(date) {
  const p = appDateParts_(date);
  return dateInAppTimezone_(p[0], p[1], p[2], SLA_WORK_END_HOUR_, SLA_WORK_END_MINUTE_, 0, 0);
}

/** Moves a time outside the schedule to the next opening of a working window. */
function moveToNextWorkingStart_(date) {
  let candidate = toDate_(date);
  if (isWorkingDay_(candidate) && candidate < getWorkingDayStart_(candidate)) return getWorkingDayStart_(candidate);

  // Noon is safely within the same local calendar day when stepping through dates.
  let p = appDateParts_(candidate);
  candidate = dateInAppTimezone_(p[0], p[1], p[2] + 1, 12, 0, 0, 0);
  while (!isWorkingDay_(candidate)) {
    p = appDateParts_(candidate);
    candidate = dateInAppTimezone_(p[0], p[1], p[2] + 1, 12, 0, 0, 0);
  }
  return getWorkingDayStart_(candidate);
}

/** Adds SLA hours only within Monday-Friday, 11:30-19:30 in APP.TZ. */
function calculateWorkingSlaDueAt_(createdAt, slaHours) {
  let cursor = toDate_(createdAt);
  let remainingMs = Math.max(0, number_(slaHours, 0)) * 60 * 60 * 1000;

  if (!isWorkingDay_(cursor) || cursor >= getWorkingDayEnd_(cursor)) cursor = moveToNextWorkingStart_(cursor);
  else if (cursor < getWorkingDayStart_(cursor)) cursor = getWorkingDayStart_(cursor);
  if (remainingMs === 0) return cursor;

  while (remainingMs > 0) {
    const availableMs = getWorkingDayEnd_(cursor).getTime() - cursor.getTime();
    if (remainingMs < availableMs) return new Date(cursor.getTime() + remainingMs);
    remainingMs -= availableMs;
    cursor = moveToNextWorkingStart_(cursor);
  }
  // An SLA ending exactly at closing is represented by the next working opening.
  return cursor;
}

/**
 * Optional editor/admin migration. Recalculates only currently open tickets.
 * This function is never invoked by application startup or ticket creation.
 */
function recalculateOpenTicketSlaDueDates() {
  const user = requireRole_([APP.ROLES.ADMIN]);
  const tickets = getSheetObjects_(APP.SHEETS.TICKETS);
  let updated = 0;
  tickets.forEach(ticket => {
    if (String(ticket.Status) === APP.STATUS.RESOLVED) return;
    const found = findObjectRow_(APP.SHEETS.TICKETS, 'Ticket_ID', ticket.Ticket_ID);
    updateObjectRow_(APP.SHEETS.TICKETS, found.rowNumber, {
      SLA_Due_At: calculateWorkingSlaDueAt_(ticket.Created_At, number_(ticket.SLA_Hours, 0)),
      Updated_At: new Date(),
      Updated_By: user.email
    });
    updated++;
  });
  removeCachedKeys_([CACHE_KEYS_.NUMBERS]);
  return { updated };
}

/** Logs the required non-mutating SLA examples and returns their results. */
function diagnoseWorkingSlaCalculation() {
  const cases = [
    ['Monday before opening', '2026-08-03T10:00:00+05:30', 4, '2026-08-03T15:30:00+05:30'],
    ['Monday at opening', '2026-08-03T11:30:00+05:30', 4, '2026-08-03T15:30:00+05:30'],
    ['Monday evening', '2026-08-03T18:30:00+05:30', 4, '2026-08-04T14:30:00+05:30'],
    ['Friday evening', '2026-08-07T18:30:00+05:30', 4, '2026-08-10T14:30:00+05:30'],
    ['Friday 19:29', '2026-08-07T19:29:00+05:30', 4, '2026-08-10T15:29:00+05:30'],
    ['Friday closing', '2026-08-07T19:30:00+05:30', 4, '2026-08-10T15:30:00+05:30'],
    ['Saturday', '2026-08-08T13:00:00+05:30', 4, '2026-08-10T15:30:00+05:30'],
    ['Sixteen hours', '2026-08-03T11:30:00+05:30', 16, '2026-08-05T11:30:00+05:30']
  ];
  const results = cases.map(test => {
    const actual = calculateWorkingSlaDueAt_(new Date(test[1]), test[2]);
    const expected = new Date(test[3]);
    return { name: test[0], expected: expected.toISOString(), actual: actual.toISOString(), passed: actual.getTime() === expected.getTime() };
  });
  results.forEach(result => console.log(JSON.stringify(result)));
  return results;
}



function queueSort_(a, b) {
  const statusRank = t => String(t.Status) === APP.STATUS.RESOLVED ? 2 : (toDate_(t.SLA_Due_At) < new Date() ? 0 : 1);
  const sr = statusRank(a) - statusRank(b);
  if (sr) return sr;
  const priorityRank = { CRITICAL: 0, HIGH: 1, MEDIUM: 2, LOW: 3 };
  const pr = (priorityRank[String(a.Priority).toUpperCase()] ?? 9) - (priorityRank[String(b.Priority).toUpperCase()] ?? 9);
  if (pr) return pr;
  return toDate_(a.Created_At) - toDate_(b.Created_At);
}

function countBy_(rows, key) {
  const counts = {};
  rows.forEach(r => {
    const value = String(r[key] || 'Unknown');
    counts[value] = (counts[value] || 0) + 1;
  });
  return Object.keys(counts).map(name => ({ name, count: counts[name] })).sort((a, b) => b.count - a.count);
}

function extractDynamicFields_(form, category) {
  const fields = safeJsonParse_(category.Fields_JSON, []);
  const result = {};
  fields.filter(f => f !== 'attachment').forEach(field => {
    const value = cleanText_(form[field], 10000);
    if (value) result[field] = value;
  });
  return result;
}

// -------------------- Duplicate matching --------------------


function buildClientKeyFromPayload_(payload) {
  const clientId = String(payload.clientId == null ? '' : payload.clientId).trim();
  if (clientId) return `ID:${clientId}`;
  return `NAME:${String(payload.clientType)}:${normalizeSubject_(payload.clientName)}`;
}

function buildClientKeyFromTicket_(ticket) {
  const clientId = String(ticket.Client_ID == null ? '' : ticket.Client_ID).trim();
  if (clientId) return `ID:${clientId}`;
  return `NAME:${String(ticket.Client_Type)}:${normalizeSubject_(ticket.Client_Name)}`;
}

function normalizeSubject_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/^(\s*(re|fw|fwd|urgent)\s*[:\-]\s*)+/gi, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\b(the|a|an|please|issue|problem|error|client|regarding|help)\b/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function subjectSimilarity_(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const setA = new Set(a.split(' ').filter(Boolean));
  const setB = new Set(b.split(' ').filter(Boolean));
  const intersection = [...setA].filter(x => setB.has(x)).length;
  const union = new Set([...setA, ...setB]).size;
  const jaccard = union ? intersection / union : 0;
  const containment = Math.min(setA.size, setB.size) ? intersection / Math.min(setA.size, setB.size) : 0;
  return Math.max(jaccard, containment * 0.9);
}

// -------------------- General helpers --------------------

function fieldLabel_(key) {
  const labels = {
    awb: 'AWB number', order_id: 'Order ID', endpoint: 'API endpoint', api_log: 'API error log',
    label_type: 'Label type', invoice_reference: 'Invoice reference', amount: 'Amount', expected_date: 'Expected date',
    seller_id: 'Seller ID', user_identifier: 'User email/mobile', origin_pincode: 'Origin pincode',
    destination_pincode: 'Destination pincode', service_type: 'Service type', expected_status: 'Expected status',
    actual_status: 'Actual status', wallet_reference: 'Wallet reference', attachment: 'Evidence attachment'
  };
  return labels[key] || key.replace(/_/g, ' ');
}

function safeJsonParse_(value, fallback) {
  try { return JSON.parse(String(value || '')); } catch (err) { return fallback; }
}

function getCachedJson_(key) {
  try {
    const value = CacheService.getScriptCache().get(key);
    return value === null ? null : JSON.parse(value);
  } catch (err) {
    console.warn('Cache read unavailable for key ' + key);
    return null;
  }
}

function putCachedJson_(key, value, seconds) {
  try { CacheService.getScriptCache().put(key, JSON.stringify(value), seconds); }
  catch (err) { console.warn('Cache write unavailable for key ' + key); }
}

function removeCachedKeys_(keys) {
  try { CacheService.getScriptCache().removeAll(keys); }
  catch (err) { console.warn('Cache invalidation unavailable'); }
}

function logPerformance_(functionName, startedAt, metadata) {
  const safe = { functionName, durationMs: Date.now() - startedAt };
  if (metadata && metadata.rows !== undefined) safe.rows = metadata.rows;
  if (metadata && metadata.cache !== undefined) safe.cache = metadata.cache;
  if (metadata && metadata.lockWaitMs !== undefined) safe.lockWaitMs = metadata.lockWaitMs;
  console.log(JSON.stringify(safe));
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength || 5000);
}

// Range.setValues treats a leading formula marker as executable content. Keep
// user-authored text as text while retaining the visible value in Sheets.
function safeSheetText_(value, maxLength) {
  const text = cleanText_(value, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function cleanFileName_(name) {
  return String(name).replace(/[\\/:*?"<>|]+/g, '_').slice(0, 150);
}

function lower_(value) { return String(value || '').trim().toLowerCase(); }
function number_(value, fallback) { const n = Number(value); return Number.isFinite(n) ? n : fallback; }
function truthy_(value) { return value === true || ['true','yes','1','active'].includes(lower_(value)); }
function toDate_(value) { const d = value instanceof Date ? value : new Date(value); return isNaN(d) ? new Date(0) : d; }
function formatDateTime_(value) { return Utilities.formatDate(toDate_(value), APP.TZ, 'dd MMM yyyy, hh:mm a'); }
function formatDateTimeOptional_(value) { return value ? formatDateTime_(value) : ''; }

// -------------------- Client size and performance upgrade --------------------

function getUserRowByEmail_(email) {
  const sheet = getRequiredSheet_(APP.SHEETS.USERS);
  const width = sheet.getLastColumn(), height = sheet.getLastRow();
  if (width < 1 || height < 2) return null;
  const headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0].map(String);
  const emailIndex = headers.indexOf('Email');
  if (emailIndex < 0) throw new Error('Email column not found in Users.');
  const values = sheet.getRange(2, emailIndex + 1, height - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) if (lower_(values[i][0]) === email) {
    const row = sheet.getRange(i + 2, 1, 1, width).getValues()[0];
    return headers.reduce((result, header, column) => { result[header] = row[column]; return result; }, {});
  }
  return null;
}

function requireUser_() {
  const email = getVerifiedCompanyEmail_();
  const row = getUserRowByEmail_(email);
  if (!row || !truthy_(row.Active)) throw new Error('You do not have access yet. Ask the app administrator to add your email to the Users sheet.');
  return userFromRow_(row, email);
}

function getActiveClientSizePriorities_(forceRefresh) {
  if (forceRefresh) removeCachedKeys_([CACHE_KEYS_.CLIENT_SIZES]);
  const cached = getCachedJson_(CACHE_KEYS_.CLIENT_SIZES);
  if (cached !== null) return cached;
  const rows = getSheetObjects_(APP.SHEETS.CLIENT_SIZE_PRIORITY).filter(row => truthy_(row.Active)).map(row => ({
    code: String(row.Client_Size_Code), label: String(row.Display_Label), adlDescription: String(row.ADL_Description),
    minAdl: number_(row.Min_ADL, 0), maxAdl: row.Max_ADL === '' ? null : number_(row.Max_ADL, null),
    priority: String(row.Priority).toUpperCase(), slaHours: number_(row.SLA_Hours, 0), sortOrder: number_(row.Sort_Order, 999)
  })).sort((a, b) => a.sortOrder - b.sortOrder);
  putCachedJson_(CACHE_KEYS_.CLIENT_SIZES, rows, 300);
  return rows;
}

function resolveTicketPriority_(clientType, clientSizeCode, category) {
  if (String(clientType) === '360') {
    const code = String(clientSizeCode || '');
    let size = getActiveClientSizePriorities_().find(row => row.code === code);
    if (!size) size = getActiveClientSizePriorities_(true).find(row => row.code === code);
    if (!size) throw new Error('The selected client size is no longer active. Refresh the page and choose an active client size.');
    if (!['CRITICAL','HIGH','MEDIUM','LOW'].includes(size.priority)) throw new Error('The selected Client Size has an invalid configured priority.');
    if (!(size.slaHours > 0)) throw new Error('The selected Client Size has invalid configured SLA hours.');
    return { clientSize: size.code, priority: size.priority, prioritySource: 'CLIENT_SIZE', slaHours: size.slaHours, slaSource: 'CLIENT_SIZE' };
  }
  return { clientSize: '', priority: String(category.Priority).toUpperCase(), prioritySource: 'CATEGORY', slaHours: number_(category.SLA_Hours, 24), slaSource: 'CATEGORY' };
}

function buildBootstrap_(user) {
  const settings = getSettings_();
  return { user, categories: getActiveCategories_(), clientSizes: getActiveClientSizePriorities_(),
    rootCauses: String(settings.ROOT_CAUSES || '').split('|').filter(Boolean),
    duplicateWindowDays: number_(settings.DUPLICATE_WINDOW_DAYS, 5), resolvedVisibilityDays: number_(settings.RESOLVED_VISIBILITY_DAYS, 10),
    dashboardWindowDays: number_(settings.DASHBOARD_WINDOW_DAYS, 14), release: APP_RELEASE,
    commit: APP_COMMIT === '__APP_COMMIT__' ? '' : APP_COMMIT };
}

function getInitialAppState() {
  const startedAt = Date.now();
  ensureRuntimeSchema_();
  const email = getVerifiedCompanyEmail_(), row = getUserRowByEmail_(email);
  let result;
  const commit = APP_COMMIT === '__APP_COMMIT__' ? '' : APP_COMMIT;
  if (!row) result = { state: 'REGISTER', email, release: APP_RELEASE, commit };
  else if (!truthy_(row.Active)) result = { state: 'BLOCKED', email, message: 'Your access has been disabled. Please contact the application administrator.', release: APP_RELEASE, commit };
  else { const user = userFromRow_(row, email); result = { state: 'ACTIVE', email, name: user.name, role: user.role, user, bootstrap: buildBootstrap_(user), release: APP_RELEASE, commit }; }
  logPerformance_('getInitialAppState', startedAt, { rows: row ? 1 : 0 });
  return result;
}

function ticketToIndex_(ticket) {
  return { Ticket_ID: ticket.Ticket_ID, Created_At: ticket.Created_At, Raiser_Email: ticket.Raiser_Email, Raiser_Name: ticket.Raiser_Name,
    Client_ID: ticket.Client_ID, Client_Name: ticket.Client_Name, Client_Type: ticket.Client_Type, Client_Size: ticket.Client_Size || '',
    Client_Key: buildClientKeyFromTicket_(ticket), Category_ID: ticket.Category_ID, Category_Name: ticket.Category_Name,
    Email_Subject: ticket.Email_Subject, Normalized_Subject: ticket.Normalized_Subject || normalizeSubject_(ticket.Email_Subject),
    Priority: ticket.Priority, Priority_Source: ticket.Priority_Source || '', SLA_Due_At: ticket.SLA_Due_At, Status: ticket.Status,
    Resolved_At: ticket.Resolved_At || '', SLA_Result: ticket.SLA_Result || '', Current_SLA_Cycle: ticket.Current_SLA_Cycle || 1,
    Submission_Request_ID: ticket.Submission_Request_ID || '', Updated_At: ticket.Updated_At || ticket.Created_At };
}

function upsertTicketIndex_(ticket) {
  const found = findObjectRow_(APP.SHEETS.TICKET_INDEX, 'Ticket_ID', ticket.Ticket_ID), data = ticketToIndex_(ticket);
  if (found) updateObjectRow_(APP.SHEETS.TICKET_INDEX, found.rowNumber, data); else appendObject_(APP.SHEETS.TICKET_INDEX, data);
}

function rebuildTicketIndex_(upgradeContext) {
  if (!upgradeContext) requireRole_([APP.ROLES.ADMIN]);
  const tickets = getSheetObjects_(APP.SHEETS.TICKETS), sheet = getSpreadsheet_().getSheetByName(APP.SHEETS.TICKET_INDEX);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(String);
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  if (tickets.length) sheet.getRange(2, 1, tickets.length, headers.length).setValues(tickets.map(ticket => { const item = ticketToIndex_(ticket); return headers.map(h => item[h] === undefined ? '' : item[h]); }));
  return { indexed: tickets.length };
}
function rebuildTicketIndex() { return rebuildTicketIndex_(false); }

function invalidateApplicationCaches_() { removeCachedKeys_([CACHE_KEYS_.SETTINGS, CACHE_KEYS_.CATEGORIES, CACHE_KEYS_.CLIENT_SIZES, CACHE_KEYS_.NUMBERS]); }
function invalidateTicketCaches_() { removeCachedKeys_([CACHE_KEYS_.NUMBERS]); }

function serializeTicket_(t) {
  const due = toDate_(t.SLA_Due_At), resolved = String(t.Status) === APP.STATUS.RESOLVED;
  return { ticketId: String(t.Ticket_ID), createdAt: formatDateTime_(t.Created_At), createdAtIso: toDate_(t.Created_At).toISOString(),
    raiserEmail: String(t.Raiser_Email), raiserName: String(t.Raiser_Name || ''), clientId: String(t.Client_ID || ''), clientName: String(t.Client_Name || ''),
    clientType: String(t.Client_Type || ''), clientSize: String(t.Client_Size || ''), clientMode: String(t.Client_Mode || ''), categoryId: String(t.Category_ID || ''),
    categoryName: String(t.Category_Name || ''), emailSubject: String(t.Email_Subject || ''), issueDescription: String(t.Issue_Description || ''),
    priority: String(t.Priority || ''), prioritySource: String(t.Priority_Source || ''), slaHours: number_(t.SLA_Hours, 0), slaSource: String(t.SLA_Source || ''), slaDueAt: formatDateTime_(t.SLA_Due_At),
    slaDueAtIso: due.toISOString(), slaStatus: resolved ? String(t.SLA_Result || '') : (due < new Date() ? 'OVERDUE' : 'ON TRACK'), status: String(t.Status || ''),
    pickedUpBy: String(t.Picked_Up_By || ''), pickedUpAt: formatDateTimeOptional_(t.Picked_Up_At), investigatingAt: formatDateTimeOptional_(t.Investigating_At),
    resolutionNote: String(t.Resolution_Note || ''), rootCause: String(t.Root_Cause || ''), resolvedBy: String(t.Resolved_By || ''), resolvedAt: formatDateTimeOptional_(t.Resolved_At),
    duplicateOf: String(t.Duplicate_Of || ''), duplicateOverride: truthy_(t.Duplicate_Override), attachmentFileName: String(t.Attachment_File_Name || ''),
    attachmentUrl: String(t.Attachment_URL || ''), submissionRequestId: String(t.Submission_Request_ID || ''), updatedAt: formatDateTimeOptional_(t.Updated_At) };
}

function serializeQueueTicket_(t) { const x = serializeTicket_(t); return { ticketId:x.ticketId,createdAt:x.createdAt,createdAtIso:x.createdAtIso,clientName:x.clientName,clientType:x.clientType,clientSize:x.clientSize,categoryName:x.categoryName,emailSubject:x.emailSubject,priority:x.priority,prioritySource:x.prioritySource,status:x.status,slaStatus:x.slaStatus,slaDueAt:x.slaDueAt,slaDueAtIso:x.slaDueAtIso,raiserEmail:x.raiserEmail }; }

function paginate_(rows, requestedPage, requestedSize) { const pageSize=Math.min(100,Math.max(1,Math.floor(number_(requestedSize,50)))), totalRows=rows.length,totalPages=Math.max(1,Math.ceil(totalRows/pageSize)),page=Math.min(Math.max(1,Math.floor(number_(requestedPage,1))),totalPages); return { rows:rows.slice((page-1)*pageSize,page*pageSize),page,pageSize,totalRows,totalPages }; }

/** Builds one request-local, case-normalized role lookup, including inactive users. */
function getUserRoleMap_() {
  return getSheetObjects_(APP.SHEETS.USERS).reduce((roles, row) => {
    const email = lower_(row.Email);
    if (email) roles[email] = String(row.Role || '').trim().toUpperCase();
    return roles;
  }, {});
}

function isSalesRaisedTicket_(ticket, roleMap) {
  const email = lower_(ticket && ticket.Raiser_Email);
  return Boolean(email && roleMap && String(roleMap[email] || '').toUpperCase() === APP.ROLES.SALES);
}

function ticketMatchesActivitySearch_(ticket, search) {
  if (!search) return true;
  const due = toDate_(ticket.SLA_Due_At);
  const resolved = String(ticket.Status) === APP.STATUS.RESOLVED;
  const slaStatus = resolved ? String(ticket.SLA_Result || '') : (due < new Date() ? 'OVERDUE' : 'ON TRACK');
  return [ticket.Ticket_ID, ticket.Client_ID, ticket.Client_Name, ticket.Client_Type, ticket.Client_Size,
    ticket.Category_Name, ticket.Email_Subject, ticket.Raiser_Name, ticket.Raiser_Email, ticket.Priority,
    ticket.Status, slaStatus].some(value => lower_(value).includes(search));
}

function getMyTickets(options) {
  const startedAt=Date.now(),user=requireUser_(),settings=getSettings_(),cutoff=new Date(Date.now()-number_(settings.RESOLVED_VISIBILITY_DAYS,10)*86400000),search=lower_(options&&options.search).trim();
  const roleMap=user.role===APP.ROLES.SALES?getUserRoleMap_():null,all=getSheetObjects_(APP.SHEETS.TICKETS);
  const filtered=all.filter(t=>user.role===APP.ROLES.SALES?isSalesRaisedTicket_(t,roleMap):lower_(t.Raiser_Email)===user.email).filter(t=>String(t.Status)!==APP.STATUS.RESOLVED||toDate_(t.Resolved_At)>=cutoff).filter(t=>ticketMatchesActivitySearch_(t,search)).sort((a,b)=>toDate_(b.Created_At)-toDate_(a.Created_At));
  const page=paginate_(filtered,options&&options.page,options&&options.pageSize); page.rows=page.rows.map(serializeQueueTicket_); logPerformance_('getMyTickets',startedAt,{rows:all.length}); return page;
}

function getQueueTickets(filters) {
  const startedAt=Date.now(); requireRole_([APP.ROLES.POC,APP.ROLES.ADMIN]); filters=filters||{};
  const search=lower_(filters.search),status=filters.status===undefined?'OPEN':String(filters.status),priority=String(filters.priority||''),category=String(filters.category||''),clientSize=String(filters.clientSize||''),sla=String(filters.sla||'');
  const all=getSheetObjects_(APP.SHEETS.TICKETS), filtered=all.filter(t=>{const x=serializeQueueTicket_(t);return(!search||[x.ticketId,x.clientName,x.emailSubject,x.raiserEmail].join(' ').toLowerCase().includes(search))&&(status==='OPEN'?['Raised','Reopened','Investigating'].includes(x.status):(!status||x.status===status))&&(!priority||x.priority===priority)&&(!category||x.categoryName===category)&&(!clientSize||x.clientSize===clientSize)&&(!sla||x.slaStatus===sla);}).sort(queueSort_);
  const page=paginate_(filtered,filters.page,filters.pageSize); page.rows=page.rows.map(serializeQueueTicket_); page.categories=getActiveCategories_().map(c=>c.name).filter((v,i,a)=>a.indexOf(v)===i).sort(); page.clientSizes=getActiveClientSizePriorities_(); logPerformance_('getQueueTickets',startedAt,{rows:all.length}); return page;
}

function getMatchingRows_(sheetName, key, value) {
  const sheet=getSpreadsheet_().getSheetByName(sheetName),width=sheet.getLastColumn(),height=sheet.getLastRow(); if(width<1||height<2)return[];
  const headers=sheet.getRange(1,1,1,width).getDisplayValues()[0].map(String),index=headers.indexOf(key); if(index<0)throw new Error(`Column ${key} not found in ${sheetName}.`);
  const keys=sheet.getRange(2,index+1,height-1,1).getValues(),numbers=[]; keys.forEach((row,i)=>{if(String(row[0])===String(value))numbers.push(i+2);});
  return numbers.map(n=>{const row=sheet.getRange(n,1,1,width).getValues()[0];return headers.reduce((o,h,i)=>{o[h]=row[i];return o;},{});});
}
function getTicketSlaCycles_(ticketId) { try{return getMatchingRows_(APP.SHEETS.SLA_CYCLES,'Ticket_ID',ticketId).sort((a,b)=>number_(b.Cycle_Number,0)-number_(a.Cycle_Number,0));}catch(err){SLA_SCHEMA_RECOVERY_FAILED_=true;return[];} }

function getTicketDetail(ticketId) {
  const startedAt=Date.now(),user=requireUser_(),found=findObjectRow_(APP.SHEETS.TICKETS,'Ticket_ID',ticketId); if(user.role===APP.ROLES.SALES&&(!found||(lower_(found.object.Raiser_Email)!==user.email&&!isSalesRaisedTicket_(found.object,getUserRoleMap_()))))throw new Error('You are not allowed to view this ticket.'); if(!found)throw new Error('Ticket not found.');
  const result=serializeTicket_(found.object); result.dynamicFields=safeJsonParse_(found.object.Dynamic_Fields_JSON,{}); result.slaCycles=getTicketSlaCycles_(ticketId).map(serializeSlaCycle_); result.slaCycleNumber=result.slaCycles.reduce((m,c)=>Math.max(m,c.cycleNumber),0); result.reopenCount=result.slaCycles.filter(c=>c.cycleType==='REOPEN').length; result.canReopen=String(found.object.Status)===APP.STATUS.RESOLVED&&([APP.ROLES.POC,APP.ROLES.ADMIN].includes(user.role)||lower_(found.object.Raiser_Email)===user.email); result.events=getMatchingRows_(APP.SHEETS.EVENTS,'Ticket_ID',ticketId).sort((a,b)=>toDate_(a.Created_At)-toDate_(b.Created_At)).map(e=>({eventType:String(e.Event_Type),oldValue:String(e.Old_Value||''),newValue:String(e.New_Value||''),performedBy:String(e.Performed_By||''),createdAt:formatDateTime_(e.Created_At),note:String(e.Note||'')})); logPerformance_('getTicketDetail',startedAt,{rows:1}); return result;
}

function getNumbers() {
  const startedAt=Date.now();requireRole_([APP.ROLES.POC,APP.ROLES.ADMIN]);const cached=getCachedJson_(CACHE_KEYS_.NUMBERS);if(cached!==null){logPerformance_('getNumbers',startedAt,{cache:'hit'});return cached;}
  const days=number_(getSettings_().DASHBOARD_WINDOW_DAYS,14),cutoff=new Date(Date.now()-days*86400000),now=new Date(),all=getSheetObjects_(APP.SHEETS.TICKET_INDEX);let raised=0,resolved=0,met=0,overdue=0;const open=[],sizes={};
  all.forEach(t=>{if(toDate_(t.Created_At)>=cutoff)raised++;if(t.Resolved_At&&toDate_(t.Resolved_At)>=cutoff){resolved++;if(toDate_(t.Resolved_At)<=toDate_(t.SLA_Due_At))met++;}if(String(t.Status)!==APP.STATUS.RESOLVED){open.push(t);if(toDate_(t.SLA_Due_At)<now)overdue++;sizes[String(t.Client_Size||'Not specified')]=(sizes[String(t.Client_Size||'Not specified')]||0)+1;}});
  const result={days,raised,resolved,open:open.length,overdue,adherence:resolved?Math.round(met/resolved*1000)/10:null,byStatus:countBy_(open,'Status'),byPriority:countBy_(open,'Priority'),byCategory:countBy_(open,'Category_Name'),byClientSize:Object.keys(sizes).map(name=>({name,count:sizes[name]})).sort((a,b)=>b.count-a.count)};putCachedJson_(CACHE_KEYS_.NUMBERS,result,60);logPerformance_('getNumbers',startedAt,{rows:all.length,cache:'miss'});return result;
}

function findTicketBySubmissionRequestId_(requestId) { return requestId ? findObjectRow_(APP.SHEETS.TICKETS,'Submission_Request_ID',requestId) : null; }
function findEventByRequestId_(requestId) { return requestId ? findObjectRow_(APP.SHEETS.EVENTS,'Request_ID',requestId) : null; }
function appendEvent_(ticketId,type,oldValue,newValue,performedBy,note,createdAt,requestId) { appendObject_(APP.SHEETS.EVENTS,{Event_ID:Utilities.getUuid(),Ticket_ID:ticketId,Event_Type:type,Old_Value:oldValue,New_Value:newValue,Performed_By:performedBy,Created_At:createdAt||new Date(),Note:note||'',Request_ID:requestId||''}); }

function submitTicket(form) {
  const startedAt=Date.now(),user=requireUser_();if(!form||typeof form!=='object')throw new Error('The form data was not received. Please refresh and try again.');
  if (!String(form.clientMode || '').trim()) throw new Error('The client mode was not submitted. Refresh the page and choose the client status again.');
  if (!cleanText_(form.clientName,200)) throw new Error('The client name was not submitted. Enter the client name and try again.');
  if (!String(form.clientType || '').trim()) throw new Error('The client type was not submitted. Select a client type and try again.');
  if (!String(form.categoryId || '').trim()) throw new Error('The category was not submitted. Refresh the page and select the category again.');
  if (String(form.clientType) === '360' && !String(form.clientSize || '').trim()) throw new Error('The client size was not submitted. Select a client size and try again.');
  if (!cleanText_(form.emailSubject,300)) throw new Error('The email subject was not submitted. Enter a subject and try again.');
  if (!cleanText_(form.issueDescription,5000)) throw new Error('The issue description was not submitted. Enter the issue description and try again.');
  const requestId=cleanText_(form.submissionRequestId,100);if(!requestId)throw new Error('The submission request ID was not submitted. Refresh the page and try again.');
  let existing=findTicketBySubmissionRequestId_(requestId);if(existing)return getTicketDetail(existing.object.Ticket_ID);
  const category=getCategoryById_(form.categoryId);if(!category)throw new Error('The selected category is no longer active. Refresh the page and choose an active category.');const client=resolveClient_(form,category.Client_Type);validateTicketForm_(form,category,client);const resolution=resolveTicketPriority_(client.type,form.clientSize,category);
  let ticketId,lock=LockService.getScriptLock(),lockStarted=Date.now();lock.waitLock(10000);try{ticketId=nextTicketId_();}finally{lock.releaseLock();}const lockWaitMs=Date.now()-lockStarted;
  let attachment={id:'',name:'',url:''};try{attachment=saveAttachment_(form.attachment,ticketId,user.email);}catch(err){throw err;}
  const createdAt=new Date(),slaHours=resolution.slaHours,slaDueAt=calculateWorkingSlaDueAt_(createdAt,slaHours),subject=safeSheetText_(form.emailSubject,300);
  const ticket={Ticket_ID:ticketId,Created_At:createdAt,Raiser_Email:user.email,Raiser_Name:user.name,Client_Mode:client.mode,Client_ID:safeSheetText_(client.id,20),Client_Name:safeSheetText_(client.name,200),Client_Type:client.type,Client_Size:resolution.clientSize,Category_ID:category.Category_ID,Category_Name:category.Category_Name,Email_Subject:subject,Normalized_Subject:normalizeSubject_(subject),Issue_Description:safeSheetText_(form.issueDescription,5000),Priority:resolution.priority,Priority_Source:resolution.prioritySource,SLA_Hours:slaHours,SLA_Source:resolution.slaSource,SLA_Due_At:slaDueAt,Status:APP.STATUS.RAISED,Picked_Up_By:'',Picked_Up_At:'',Investigating_At:'',Resolution_Note:'',Root_Cause:'',Resolved_By:'',Resolved_At:'',SLA_Result:'ON TRACK',Duplicate_Of:safeSheetText_(form.duplicateIds||'',500),Duplicate_Override:String(form.duplicateOverride||'').toLowerCase()==='true',Dynamic_Fields_JSON:safeSheetText_(JSON.stringify(extractDynamicFields_(form,category)),50000),Attachment_File_ID:attachment.id,Attachment_File_Name:attachment.name,Attachment_URL:attachment.url,Updated_At:createdAt,Updated_By:user.email,Submission_Request_ID:requestId,Current_SLA_Cycle:1};
  let ticketCommitted=false,cleanupAttachment=false,commitError=null,duplicateTicketId='';lock=LockService.getScriptLock();lockStarted=Date.now();lock.waitLock(10000);try{existing=findTicketBySubmissionRequestId_(requestId);if(existing){cleanupAttachment=Boolean(attachment.id);duplicateTicketId=String(existing.object.Ticket_ID);}else{requireTicketSlaCyclesSchemaForWrite_(true);appendObject_(APP.SHEETS.TICKETS,ticket);ticketCommitted=true;appendSlaCycle_({Ticket_ID:ticketId,Cycle_Number:1,Cycle_Type:'INITIAL',Started_At:createdAt,Due_At:slaDueAt,SLA_Result:'OPEN',Started_By:user.email,Created_At:createdAt,Updated_At:createdAt});appendEvent_(ticketId,'TICKET_RAISED','',APP.STATUS.RAISED,user.email,ticket.Duplicate_Override?'Raised despite duplicate warning.':'',createdAt,requestId);upsertTicketIndex_(ticket);invalidateTicketCaches_();}}catch(err){commitError=err;cleanupAttachment=!ticketCommitted&&Boolean(attachment.id);}finally{lock.releaseLock();}
  // Drive calls are deliberately outside ScriptLock. A file attached to an
  // authoritative Tickets row is never deleted, even if a derived write fails.
  if(cleanupAttachment)try{DriveApp.getFileById(attachment.id).setTrashed(true);}catch(orphanErr){console.error('Uncommitted attachment cleanup failed.');}
  if(commitError)throw commitError;
  if(duplicateTicketId)return getTicketDetail(duplicateTicketId);
  const detail=getTicketDetail(ticketId);try{sendSlackAlert_(detail);}catch(err){console.error('Slack alert failed.');}logPerformance_('submitTicket',startedAt,{rows:1,lockWaitMs:lockWaitMs+Date.now()-lockStarted});return detail;
}

function updateTicketStatus(payload) {
  const startedAt=Date.now(),user=requireRole_([APP.ROLES.POC,APP.ROLES.ADMIN]);if(!payload||!payload.ticketId||!payload.newStatus)throw new Error('Ticket and new status are required.');const requestId=cleanText_(payload.actionRequestId,100);if(!requestId)throw new Error('Action_Request_ID is required.');
  let prior=findEventByRequestId_(requestId);if(prior)return getTicketDetail(prior.object.Ticket_ID);const lock=LockService.getScriptLock(),wait=Date.now();lock.waitLock(10000);try{prior=findEventByRequestId_(requestId);if(prior)return getTicketDetail(prior.object.Ticket_ID);const found=findObjectRow_(APP.SHEETS.TICKETS,'Ticket_ID',payload.ticketId);if(!found)throw new Error('Ticket not found.');const ticket=found.object,oldStatus=String(ticket.Status),newStatus=String(payload.newStatus),allowed={Raised:'Investigating',Reopened:'Investigating',Investigating:'Resolved'};if(allowed[oldStatus]!==newStatus)throw new Error(`Transition from ${oldStatus} to ${newStatus} is not allowed.`);const now=new Date(),changes={Status:newStatus,Updated_At:now,Updated_By:user.email};if(newStatus===APP.STATUS.INVESTIGATING){changes.Picked_Up_By=user.email;changes.Picked_Up_At=now;changes.Investigating_At=now;}if(newStatus===APP.STATUS.RESOLVED){changes.Resolution_Note=safeSheetText_(payload.resolutionNote,5000);changes.Root_Cause=safeSheetText_(payload.rootCause,200);if(!changes.Resolution_Note||!changes.Root_Cause)throw new Error('Resolution note and root cause are mandatory.');const allowedRootCauses=String(getSettings_().ROOT_CAUSES||'').split('|').map(v=>v.trim()).filter(Boolean);if(allowedRootCauses.indexOf(cleanText_(payload.rootCause,200))<0)throw new Error('Select a valid root cause from the configured list.');changes.Resolved_By=user.email;changes.Resolved_At=now;changes.SLA_Result=now<=toDate_(ticket.SLA_Due_At)?'MET':'BREACHED';ensureOpenInitialCycle_(ticket);closeOpenSlaCycle_(payload.ticketId,now,user.email,changes.SLA_Result);}updateObjectRow_(APP.SHEETS.TICKETS,found.rowNumber,changes);const updated=Object.assign({},ticket,changes);appendEvent_(payload.ticketId,'STATUS_CHANGED',oldStatus,newStatus,user.email,newStatus===APP.STATUS.RESOLVED?`${changes.Root_Cause}: ${changes.Resolution_Note}`:'',now,requestId);upsertTicketIndex_(updated);invalidateTicketCaches_();}finally{lock.releaseLock();}const result=getTicketDetail(payload.ticketId);logPerformance_('updateTicketStatus',startedAt,{rows:1,lockWaitMs:Date.now()-wait});return result;
}

function reopenTicket(payload) {
  const startedAt=Date.now(),user=requireUser_();if(!payload||!payload.ticketId)throw new Error('Ticket is required.');const reason=cleanText_(payload.reopenReason,5001),requestId=cleanText_(payload.actionRequestId,100);if(!reason||reason.length>5000)throw new Error('A reopening reason of at most 5,000 characters is mandatory.');if(!requestId)throw new Error('Action_Request_ID is required.');let prior=findEventByRequestId_(requestId);if(prior)return getTicketDetail(prior.object.Ticket_ID);
  const lock=LockService.getScriptLock(),wait=Date.now();lock.waitLock(10000);try{prior=findEventByRequestId_(requestId);if(prior)return getTicketDetail(prior.object.Ticket_ID);const found=findObjectRow_(APP.SHEETS.TICKETS,'Ticket_ID',payload.ticketId);if(!found)throw new Error('Ticket not found.');const ticket=found.object;assertCanReopen_(user,ticket);ensureLegacyInitialCycle_(ticket);const cycles=getTicketSlaCycles_(payload.ticketId);if(cycles.some(c=>String(c.SLA_Result)==='OPEN'))throw new Error('This ticket already has an open SLA cycle.');const now=new Date(),due=calculateWorkingSlaDueAt_(now,number_(ticket.SLA_Hours,0)),cycle=cycles.reduce((m,c)=>Math.max(m,number_(c.Cycle_Number,0)),0)+1,changes={Status:APP.STATUS.REOPENED,SLA_Due_At:due,SLA_Result:'',Updated_At:now,Updated_By:user.email,Picked_Up_By:'',Picked_Up_At:'',Investigating_At:'',Resolution_Note:'',Root_Cause:'',Resolved_By:'',Resolved_At:'',Current_SLA_Cycle:cycle};updateObjectRow_(APP.SHEETS.TICKETS,found.rowNumber,changes);appendSlaCycle_({Ticket_ID:payload.ticketId,Cycle_Number:cycle,Cycle_Type:'REOPEN',Started_At:now,Due_At:due,SLA_Result:'OPEN',Started_By:user.email,Reopen_Reason:reason,Created_At:now,Updated_At:now});appendEvent_(payload.ticketId,'TICKET_REOPENED',APP.STATUS.RESOLVED,APP.STATUS.REOPENED,user.email,reason,now,requestId);upsertTicketIndex_(Object.assign({},ticket,changes));invalidateTicketCaches_();}finally{lock.releaseLock();}const detail=getTicketDetail(payload.ticketId);try{sendSlackReopenedAlert_(detail,user.email,reason);}catch(err){console.error('Slack reopen alert failed.');}logPerformance_('reopenTicket',startedAt,{rows:1,lockWaitMs:Date.now()-wait});return detail;
}

function getRecentTicketObjects_(cutoff,batchSize) { const sheet=getRequiredSheet_(APP.SHEETS.TICKET_INDEX);const lastRow=sheet.getLastRow(),width=sheet.getLastColumn();if(lastRow<2)return{rows:[],processed:0};const headers=sheet.getRange(1,1,1,width).getDisplayValues()[0].map(String),created=headers.indexOf('Created_At'),rows=[];let end=lastRow,processed=0,previous=null,unordered=false;while(end>=2){const count=Math.min(batchSize||200,end-1),start=end-count+1,values=sheet.getRange(start,1,count,width).getValues();processed+=count;for(let i=values.length-1;i>=0;i--){const date=toDate_(values[i][created]);if(previous&&date>previous){unordered=true;break;}previous=date;if(date<cutoff)return{rows,processed};rows.push(headers.reduce((o,h,j)=>{o[h]=values[i][j];return o;},{}));}if(unordered)break;end=start-1;}if(unordered){const all=getSheetObjects_(APP.SHEETS.TICKET_INDEX);return{rows:all.filter(t=>toDate_(t.Created_At)>=cutoff),processed:all.length};}return{rows,processed}; }

/** Category lookup reuses the 300-second active-configuration cache. */
function getCategoryById_(id) {
  let category=getActiveCategories_().find(item=>String(item.id)===String(id));
  if (!category) {
    removeCachedKeys_([CACHE_KEYS_.CATEGORIES]);
    category=getActiveCategories_().find(item=>String(item.id)===String(id));
  }
  return category ? {Category_ID:category.id,Client_Type:category.clientType,Category_Name:category.name,Priority:category.priority,SLA_Hours:category.slaHours,Fields_JSON:JSON.stringify(category.fields),Required_Fields_JSON:JSON.stringify(category.requiredFields),Active:true} : null;
}

function duplicateValueCount_(rows, key) {
  const seen={},duplicates={};
  rows.forEach(row=>{const value=String(row[key]||'').trim();if(!value)return;seen[value]=(seen[value]||0)+1;if(seen[value]===2)duplicates[value]=true;});
  return Object.keys(duplicates).length;
}

/** Pure integrity calculation used by the editor diagnostic and local tests. */
function diagnoseTicketIntegrityRows_(tickets,indexRows,cycles,events,slackRows) {
  const ticketIds={},indexIds={},cycleGroups={};
  tickets.forEach(t=>{ticketIds[String(t.Ticket_ID)]=t;});
  indexRows.forEach(t=>{indexIds[String(t.Ticket_ID)]=t;});
  cycles.forEach(c=>{const id=String(c.Ticket_ID);(cycleGroups[id]=cycleGroups[id]||[]).push(c);});
  const counts={
    ticketsWithoutIndex:0,indexRowsWithoutTicket:0,ticketsWithoutInitialSlaCycle:0,ticketsWithMultipleOpenSlaCycles:0,
    duplicateSubmissionRequestIds:duplicateValueCount_(tickets,'Submission_Request_ID'),duplicateTicketIds:duplicateValueCount_(tickets,'Ticket_ID'),
    duplicateEventRequestIds:duplicateValueCount_(events,'Request_ID'),duplicateSlackDedupeKeys:duplicateValueCount_(slackRows,'Dedupe_Key'),
    resolvedTicketsWithOpenSlaCycle:0,openTicketsWithoutOpenSlaCycle:0
  };
  tickets.forEach(ticket=>{
    const id=String(ticket.Ticket_ID),ticketCycles=cycleGroups[id]||[],open=ticketCycles.filter(c=>String(c.SLA_Result)==='OPEN');
    if(!indexIds[id])counts.ticketsWithoutIndex++;
    if(!ticketCycles.some(c=>String(c.Cycle_Type)==='INITIAL'))counts.ticketsWithoutInitialSlaCycle++;
    if(open.length>1)counts.ticketsWithMultipleOpenSlaCycles++;
    if(String(ticket.Status)===APP.STATUS.RESOLVED&&open.length)counts.resolvedTicketsWithOpenSlaCycle++;
    if(String(ticket.Status)!==APP.STATUS.RESOLVED&&!open.length)counts.openTicketsWithoutOpenSlaCycle++;
  });
  indexRows.forEach(row=>{if(!ticketIds[String(row.Ticket_ID)])counts.indexRowsWithoutTicket++;});
  return counts;
}

/** ADMIN-only, read-only diagnostic. It returns counts, never row/user content. */
function runTicketIntegrityDiagnostic() {
  requireRole_([APP.ROLES.ADMIN]);
  return diagnoseTicketIntegrityRows_(
    getSheetObjects_(APP.SHEETS.TICKETS),getSheetObjects_(APP.SHEETS.TICKET_INDEX),getSheetObjects_(APP.SHEETS.SLA_CYCLES),
    getSheetObjects_(APP.SHEETS.EVENTS),getSheetObjects_(APP.SHEETS.SLACK_NOTIFICATIONS));
}

/** Safe, read-only administrator diagnostic. It never returns row contents or identifiers. */
function runEndToEndHealthCheck() {
  requireRole_([APP.ROLES.ADMIN]);
  const ss=getSpreadsheet_(),props=PropertiesService.getScriptProperties(),warnings=[],missingHeaders={},present={};
  const duplicateHeaders={};
  Object.keys(APP.HEADERS).forEach(name=>{
    const sheet=ss.getSheetByName(name);present[name]=Boolean(sheet);
    if(!sheet){missingHeaders[name]=APP.HEADERS[name].slice();return;}
    const headers=sheet.getLastColumn()?sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(String):[];
    const missing=APP.HEADERS[name].filter(header=>headers.indexOf(header)<0);if(missing.length)missingHeaders[name]=missing;
    const duplicates=headers.filter((header,index)=>header&&headers.indexOf(header)!==index).filter((header,index,array)=>array.indexOf(header)===index);if(duplicates.length)duplicateHeaders[name]=duplicates;
  });
  const categories=present[APP.SHEETS.CATEGORIES]&&!missingHeaders[APP.SHEETS.CATEGORIES]?getActiveCategories_():[],sizes=present[APP.SHEETS.CLIENT_SIZE_PRIORITY]&&!missingHeaders[APP.SHEETS.CLIENT_SIZE_PRIORITY]?getActiveClientSizePriorities_():[],settings=present[APP.SHEETS.SETTINGS]&&!missingHeaders[APP.SHEETS.SETTINGS]?getSettings_():{};
  const tickets=safeDataRowCount_(ss.getSheetByName(APP.SHEETS.TICKETS)),index=safeDataRowCount_(ss.getSheetByName(APP.SHEETS.TICKET_INDEX));
  const triggers=ScriptApp.getProjectTriggers(),dispatcherCount=triggers.filter(t=>t.getHandlerFunction()==='dispatchSlackNotifications').length,monitorCount=triggers.filter(t=>t.getHandlerFunction()==='monitorSlackAlerts').length;
  const requiredSettings=['COMPANY_DOMAIN','MAX_ATTACHMENT_MB','DUPLICATE_WINDOW_DAYS','DUPLICATE_SIMILARITY_THRESHOLD','RESOLVED_VISIBILITY_DAYS','DASHBOARD_WINDOW_DAYS','ROOT_CAUSES'];
  const missingSettings=requiredSettings.filter(key=>String(settings[key]||'').trim()==='');
  const auth=ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationStatus()===ScriptApp.AuthorizationStatus.REQUIRED?'REQUIRED':'NOT_REQUIRED';
  if(tickets!==index)warnings.push('TicketIndex row count does not match Tickets; run repairApplicationSchema().');
  if(dispatcherCount!==1)warnings.push('Slack dispatcher trigger count must be one.');if(monitorCount!==1)warnings.push('Slack monitor trigger count must be one.');
  if(missingSettings.length)warnings.push('Required Settings keys are missing.');if(auth==='REQUIRED')warnings.push('Deployment-owner OAuth authorization is required.');
  let webhook=false;try{webhook=Boolean(getSlackWebhookUrl_());}catch(ignore){}
  const emptyIntegrityCounts=diagnoseTicketIntegrityRows_([],[],[],[],[]),integritySheets=[APP.SHEETS.TICKETS,APP.SHEETS.TICKET_INDEX,APP.SHEETS.SLA_CYCLES,APP.SHEETS.EVENTS,APP.SHEETS.SLACK_NOTIFICATIONS],integrityDiagnosticAvailable=integritySheets.every(name=>present[name]&&!missingHeaders[name]);
  const integrityIssueCounts=integrityDiagnosticAvailable?diagnoseTicketIntegrityRows_(getSheetObjects_(APP.SHEETS.TICKETS),getSheetObjects_(APP.SHEETS.TICKET_INDEX),getSheetObjects_(APP.SHEETS.SLA_CYCLES),getSheetObjects_(APP.SHEETS.EVENTS),getSheetObjects_(APP.SHEETS.SLACK_NOTIFICATIONS)):emptyIntegrityCounts;
  if(!integrityDiagnosticAvailable)warnings.push('Integrity counts are unavailable until required sheets and headers are repaired.');
  if(Object.keys(integrityIssueCounts).some(key=>integrityIssueCounts[key]>0))warnings.push('Ticket integrity issues were detected; run runTicketIntegrityDiagnostic() for count-only details.');
  const result={appRelease:APP_RELEASE,appCommit:APP_COMMIT==='__APP_COMMIT__'?'':APP_COMMIT,currentSchemaVersion:CURRENT_SCHEMA_VERSION,schemaPropertyVersion:Number(props.getProperty(APP_SCHEMA_VERSION_PROPERTY_)||0),requiredSheetsPresent:present,missingSheets:Object.keys(present).filter(name=>!present[name]),missingHeadersBySheet:missingHeaders,duplicateHeadersBySheet:duplicateHeaders,active360CategoryCount:categories.filter(c=>c.clientType==='360').length,activeRegularCategoryCount:categories.filter(c=>c.clientType==='Regular').length,activeClientSizeCount:sizes.length,clientSizes:sizes.map(s=>({code:s.code,priority:s.priority,slaHours:s.slaHours})),ticketsRowCount:tickets,ticketIndexRowCount:index,ticketIndexConsistent:tickets===index,integrityDiagnosticAvailable:integrityDiagnosticAvailable,integrityIssueCounts:integrityIssueCounts,slackNotificationsSheetReady:Boolean(present[APP.SHEETS.SLACK_NOTIFICATIONS]&&!missingHeaders[APP.SHEETS.SLACK_NOTIFICATIONS]),slackWebhookConfigured:webhook,dispatcherTriggerCount:dispatcherCount,monitorTriggerCount:monitorCount,requiredSettingsPresent:missingSettings.length===0,missingSettingsKeys:missingSettings,oauthAuthorizationStatus:auth,timezone:APP.TZ,workingHours:{days:'Monday-Friday',start:'11:30',end:'19:30',overnight:false},warnings:warnings};
  result.ready=Object.keys(present).every(k=>present[k])&&Object.keys(missingHeaders).length===0&&Object.keys(duplicateHeaders).length===0&&integrityDiagnosticAvailable&&Object.keys(integrityIssueCounts).every(key=>integrityIssueCounts[key]===0)&&result.schemaPropertyVersion===CURRENT_SCHEMA_VERSION&&result.ticketIndexConsistent&&result.requiredSettingsPresent&&auth==='NOT_REQUIRED'&&result.slackNotificationsSheetReady&&webhook&&dispatcherCount===1&&monitorCount===1;
  return result;
}

/** Validates and resolves a proposed submission without writing Sheets or Drive. */
function testTicketSubmissionPayload(payload) {
  requireRole_([APP.ROLES.ADMIN]);
  if(!payload||typeof payload!=='object')throw new Error('The form data was not received.');
  if(!String(payload.categoryId||'').trim())throw new Error('The category was not submitted. Refresh the page and select the category again.');
  if(String(payload.clientType)==='360'&&!String(payload.clientSize||'').trim())throw new Error('The client size was not submitted. Select a client size and try again.');
  const category=getCategoryById_(payload.categoryId);if(!category)throw new Error('The selected category is no longer active. Refresh the page and choose an active category.');
  const client=resolveClient_(payload,category.Client_Type);validateTicketForm_(payload,category,client);
  const resolved=resolveTicketPriority_(client.type,payload.clientSize,category),due=calculateWorkingSlaDueAt_(new Date(),resolved.slaHours);
  return {clientType:client.type,clientSize:resolved.clientSize,categoryId:String(category.Category_ID),priority:resolved.priority,slaHours:resolved.slaHours,prioritySource:resolved.prioritySource,slaSource:resolved.slaSource,dueAt:formatDateTime_(due),dueAtIso:due.toISOString()};
}

// -------------------- Asynchronous Slack automation --------------------
const SLACK_NOTIFICATION_TYPES_ = Object.freeze(['NEW_HIGH_PRIORITY','REOPENED_HIGH_PRIORITY','SLA_WARNING','SLA_BREACHED','END_OF_DAY_SUMMARY','TEST']);
const SLACK_QUEUE_STATUSES_ = Object.freeze(['PENDING','PROCESSING','SENT','FAILED']);
const SLACK_OPEN_STATUSES_ = Object.freeze(['Raised','Reopened','Investigating']);

function slackEnabled_() { return truthy_(getSettings_().SLACK_NOTIFICATIONS_ENABLED); }
function slackPriorityConfigured_(priority, key) {
  return String(getSettings_()[key] || '').split(',').map(v => v.trim().toUpperCase()).filter(Boolean).indexOf(String(priority || '').toUpperCase()) >= 0;
}
function getSlackWebhookUrl_() {
  const value = String(PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL') || '').trim();
  if (!value || value.indexOf('https://hooks.slack.com/') !== 0) throw new Error('Slack is not configured. An administrator must set a valid SLACK_WEBHOOK_URL Script Property.');
  return value;
}
function slackAppUrl_() { return String(PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl() || ''); }
function slackEscape_(value, limit) {
  let text = String(value == null ? '' : value).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/[\u0000-\u001f\u007f]/g,' ');
  if (limit && text.length > limit) text = text.slice(0, Math.max(0, limit - 1)) + '…';
  return text;
}
function slackMention_(priority) { return slackPriorityConfigured_(priority, 'SLACK_MENTION_PRIORITIES') ? '<!here> ' : ''; }
function slackTicket_(ticket) {
  return { id: ticket.ticketId || ticket.Ticket_ID, client: ticket.clientName || ticket.Client_Name, type: ticket.clientType || ticket.Client_Type,
    size: ticket.clientSize || ticket.Client_Size, category: ticket.categoryName || ticket.Category_Name, priority: ticket.priority || ticket.Priority,
    subject: ticket.emailSubject || ticket.Email_Subject, slaHours: ticket.slaHours || ticket.SLA_Hours, due: ticket.slaDueAt || ticket.SLA_Due_At,
    status: ticket.status || ticket.Status, raiser: ticket.raiserEmail || ticket.Raiser_Email, cycle: ticket.slaCycleNumber || ticket.Current_SLA_Cycle,
    handler: ticket.currentHandler || ticket.Picked_Up_By };
}
function slackFields_(items) { return items.filter(item => item[1] !== '' && item[1] != null).map(item => ({type:'mrkdwn',text:'*'+item[0]+'*\n'+slackEscape_(item[1],300)})); }
function slackPayload_(fallback, title, fields, context, mention) {
  const blocks = [{type:'section',text:{type:'mrkdwn',text:(mention || '')+'*'+title+'*'}}];
  if (fields && fields.length) blocks.push({type:'section',fields:slackFields_(fields)});
  if (context) blocks.push({type:'context',elements:[{type:'mrkdwn',text:context}]});
  const url = slackAppUrl_(); if (url) blocks.push({type:'section',text:{type:'mrkdwn',text:'<'+slackEscape_(url,1000)+'|Open Internal Ticketing>'}});
  return {text:(mention || '')+fallback,blocks:blocks};
}
function buildNewHighPrioritySlackPayload_(ticket) { const t=slackTicket_(ticket); return slackPayload_('NEW HIGH-PRIORITY TICKET '+slackEscape_(t.id,100),'🚨 NEW HIGH-PRIORITY TICKET',[
  ['Ticket',t.id],['Client',t.client],['Client type',t.type],['Client Size',t.size],['Category',t.category],['Priority',t.priority],['Subject',slackEscape_(t.subject,300)],['SLA hours',t.slaHours],['SLA due',formatDateTimeOptional_(t.due)],['Raiser',t.raiser]
],null,slackMention_(t.priority)); }
function buildReopenedHighPrioritySlackPayload_(ticket,reopenedBy,reason) { const t=slackTicket_(ticket); return slackPayload_('REOPENED HIGH-PRIORITY TICKET '+slackEscape_(t.id,100),'🔄 REOPENED HIGH-PRIORITY TICKET',[
  ['Ticket',t.id],['Client',t.client],['Category',t.category],['Priority',t.priority],['Reopened by',reopenedBy],['Reason',slackEscape_(reason,500)],['SLA cycle',t.cycle],['New SLA due',formatDateTimeOptional_(t.due)]
],null,slackMention_(t.priority)); }
function workingDurationText_(minutes) { minutes=Math.max(0,Math.round(minutes)); return (Math.floor(minutes/60)?Math.floor(minutes/60)+' hour'+(minutes>=120?'s':'')+' ':'')+(minutes%60?minutes%60+' minutes':''); }
function buildSlaWarningSlackPayload_(ticket,minutes) { const t=slackTicket_(ticket); return slackPayload_('SLA BREACH WARNING '+slackEscape_(t.id,100),'⏳ SLA BREACH WARNING',[
  ['Ticket',t.id],['Client',t.client],['Client Size',t.size],['Category',t.category],['Priority',t.priority],['Status',t.status],['Current handler',t.handler],['SLA cycle',t.cycle],['Due time',formatDateTimeOptional_(t.due)]
],'Due within '+workingDurationText_(minutes)+' of working SLA time.',slackMention_(t.priority)); }
function buildSlaBreachedSlackPayload_(ticket) { const t=slackTicket_(ticket); return slackPayload_('SLA BREACHED '+slackEscape_(t.id,100),'🔴 SLA BREACHED',[
  ['Ticket',t.id],['Client',t.client],['Category',t.category],['Priority',t.priority],['Status',t.status],['SLA cycle',t.cycle],['Due time',formatDateTimeOptional_(t.due)]
],null,slackMention_(t.priority)); }
function summaryLines_(label, values) { const keys=Object.keys(values||{}).sort(); return '*'+label+'*\n'+(keys.length?keys.map(k=>'• '+slackEscape_(k,80)+': '+values[k]).join('\n'):'• None'); }
function buildEndOfDaySlackPayload_(s) {
  const adherence=s.resolvedToday ? (Math.round(s.slaMetToday/s.resolvedToday*1000)/10)+'%' : 'N/A';
  const overdue=(s.oldestOverdue||[]).slice(0,5).map(t=>'• '+slackEscape_(t.Ticket_ID,100)+' — '+slackEscape_(t.Client_Name,120)+' — due '+formatDateTimeOptional_(t.SLA_Due_At)).join('\n')||'• None';
  return {text:'Internal Ticketing end-of-day summary '+s.date,blocks:[
    {type:'section',text:{type:'mrkdwn',text:'*📊 Internal Ticketing — End-of-Day Summary*\n'+slackEscape_(s.displayDate,100)}},
    {type:'section',fields:slackFields_([['Raised today',s.raisedToday],['Resolved today',s.resolvedToday],['Currently open',s.open],['Overdue',s.overdue],['Due within next '+s.warningMinutes+' working minutes',s.dueSoon],['SLA adherence today',adherence]])},
    {type:'section',text:{type:'mrkdwn',text:summaryLines_('Open by priority',s.byPriority)+'\n\n'+summaryLines_('Open by status',s.byStatus)+'\n\n'+summaryLines_('Open by Client Size',s.byClientSize)}},
    {type:'section',text:{type:'mrkdwn',text:'*Oldest overdue*\n'+overdue}},
    {type:'section',text:{type:'mrkdwn',text:slackAppUrl_()?'<'+slackEscape_(slackAppUrl_(),1000)+'|Open SLA Dashboard>':'Open SLA Dashboard'}}
  ]};
}
function buildSlackTestPayload_() { const release=String(PropertiesService.getScriptProperties().getProperty('APP_RELEASE') || 'Production'); return slackPayload_('Slack integration test successful','✅ Slack integration test successful',[],slackEscape_(release,100)+' • '+formatDateTime_(new Date()),''); }

function queueHeaders_(sheet) { return sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(v=>String(v).trim()); }
function queueObjects_(sheet) { const headers=queueHeaders_(sheet),values=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues():[]; return values.map((row,i)=>({rowNumber:i+2,object:headers.reduce((o,h,j)=>{o[h]=row[j];return o;},{}),values:row})); }
function queueHasDedupe_(sheet,key) { const headers=queueHeaders_(sheet),index=headers.indexOf('Dedupe_Key'); if(index<0)throw new Error('SlackNotifications is missing Dedupe_Key. Run repairApplicationSchema().'); return sheet.getLastRow()>1&&sheet.getRange(2,index+1,sheet.getLastRow()-1,1).getDisplayValues().some(row=>String(row[0])===key); }
function enqueueSlackNotification_(notification) {
  if (!slackEnabled_()) return {enqueued:false,duplicate:false};
  notification=notification||{}; const key=String(notification.dedupeKey||'').trim(); if(!key)throw new Error('Slack notification Dedupe_Key is required.');
  if(SLACK_NOTIFICATION_TYPES_.indexOf(String(notification.notificationType))<0)throw new Error('Invalid Slack Notification_Type.');
  const sheet=getRequiredSheet_(APP.SHEETS.SLACK_NOTIFICATIONS); if(queueHasDedupe_(sheet,key))return{enqueued:false,duplicate:true};
  const lock=LockService.getScriptLock(); lock.waitLock(5000); try { if(queueHasDedupe_(sheet,key))return{enqueued:false,duplicate:true};
    const now=new Date(),headers=queueHeaders_(sheet),data={Notification_ID:Utilities.getUuid(),Dedupe_Key:key,Notification_Type:notification.notificationType,Ticket_ID:notification.ticketId||'',SLA_Cycle_Number:notification.slaCycleNumber||'',Priority:notification.priority||'',Payload_JSON:JSON.stringify(notification.payload||{}),Status:'PENDING',Attempts:0,Next_Attempt_At:now,Created_At:now,Processing_Started_At:'',Sent_At:'',Last_HTTP_Code:'',Last_Error:'',Updated_At:now};
    sheet.getRange(sheet.getLastRow()+1,1,1,headers.length).setValues([headers.map(h=>data[h]===undefined?'':data[h])]); return{enqueued:true,duplicate:false};
  } finally { lock.releaseLock(); }
}

/** Working time between instants, chunked by business-day windows. */
function calculateWorkingMinutesBetween_(fromDate,toDate) {
  const from=toDate_(fromDate),to=toDate_(toDate); if(!from||!to||isNaN(from.getTime())||isNaN(to.getTime())||to<=from)return 0;
  let cursor=from,total=0; while(cursor<to){ if(!isWorkingDay_(cursor)||cursor>=getWorkingDayEnd_(cursor)){cursor=moveToNextWorkingStart_(cursor);continue;}
    if(cursor<getWorkingDayStart_(cursor))cursor=getWorkingDayStart_(cursor); if(cursor>=to)break;
    const end=new Date(Math.min(to.getTime(),getWorkingDayEnd_(cursor).getTime())); if(end>cursor)total+=end-cursor;
    if(end>=to)break; cursor=moveToNextWorkingStart_(cursor);
  } return Math.round(total/60000);
}
function monitorSlackAlerts() {
  const started=Date.now(),metrics={rowsScanned:0,notificationsEnqueued:0,notificationsSent:0,notificationsFailed:0};
  try { if(!slackEnabled_())return metrics; const settings=getSettings_(),now=new Date(),rows=getSheetObjects_(APP.SHEETS.TICKET_INDEX); metrics.rowsScanned=rows.length;
    rows.forEach(ticket=>{ try { const id=String(ticket.Ticket_ID||''),status=String(ticket.Status||''),due=toDate_(ticket.SLA_Due_At),cycle=number_(ticket.Current_SLA_Cycle,1); if(!id||SLACK_OPEN_STATUSES_.indexOf(status)<0||!due||isNaN(due.getTime()))return;
      let result;if(due<=now&&truthy_(settings.SLACK_BREACH_ALERT_ENABLED))result=enqueueSlackNotification_({notificationType:'SLA_BREACHED',dedupeKey:'SLA_BREACHED:'+id+':'+cycle,ticketId:id,slaCycleNumber:cycle,priority:ticket.Priority,payload:buildSlaBreachedSlackPayload_(ticket)});
      else if(due>now&&truthy_(settings.SLACK_BREACH_WARNING_ENABLED)){const remaining=calculateWorkingMinutesBetween_(now,due),limit=number_(settings.SLACK_BREACH_WARNING_MINUTES,120);if(remaining>0&&remaining<=limit)result=enqueueSlackNotification_({notificationType:'SLA_WARNING',dedupeKey:'SLA_WARNING:'+id+':'+cycle,ticketId:id,slaCycleNumber:cycle,priority:ticket.Priority,payload:buildSlaWarningSlackPayload_(ticket,remaining)});}
      if(result&&result.enqueued)metrics.notificationsEnqueued++;
    } catch(rowError){metrics.notificationsFailed++;console.error('monitorSlackAlerts row failed: '+slackEscape_(rowError.message,500));} });
    if(enqueueEndOfDaySlackSummary_().enqueued)metrics.notificationsEnqueued++;
  } catch(error){console.error('monitorSlackAlerts failed: '+slackEscape_(error.message,500));} finally {console.log(JSON.stringify(Object.assign({functionName:'monitorSlackAlerts',durationMs:Date.now()-started,lockWaitMs:0},metrics)));} return metrics;
}
function localDateKey_(date){return Utilities.formatDate(date,APP.TZ,'yyyy-MM-dd');}
function enqueueEndOfDaySlackSummary_() {
  if(!slackEnabled_())return{enqueued:false,duplicate:false}; const settings=getSettings_();if(!truthy_(settings.SLACK_EOD_SUMMARY_ENABLED))return{enqueued:false,duplicate:false};
  const now=new Date(),day=Number(Utilities.formatDate(now,APP.TZ,'u')),hour=Number(Utilities.formatDate(now,APP.TZ,'H')),minute=Number(Utilities.formatDate(now,APP.TZ,'m')),target=number_(settings.SLACK_EOD_HOUR,19)*60+number_(settings.SLACK_EOD_MINUTE,45); if(day>5||hour*60+minute<target)return{enqueued:false,duplicate:false};
  const date=localDateKey_(now),rows=getSheetObjects_(APP.SHEETS.TICKET_INDEX),warning=number_(settings.SLACK_BREACH_WARNING_MINUTES,120),summary={date:date,displayDate:Utilities.formatDate(now,APP.TZ,'d MMMM yyyy'),raisedToday:0,resolvedToday:0,slaMetToday:0,open:0,overdue:0,dueSoon:0,warningMinutes:warning,byPriority:{},byStatus:{},byClientSize:{},oldestOverdue:[]};
  rows.forEach(t=>{if(t.Created_At&&localDateKey_(toDate_(t.Created_At))===date)summary.raisedToday++;if(t.Resolved_At&&localDateKey_(toDate_(t.Resolved_At))===date){summary.resolvedToday++;if(String(t.SLA_Result).toUpperCase()==='MET')summary.slaMetToday++;}if(SLACK_OPEN_STATUSES_.indexOf(String(t.Status))<0)return;summary.open++;[['byPriority',t.Priority||'Not specified'],['byStatus',t.Status||'Not specified'],['byClientSize',t.Client_Size||'Not specified']].forEach(x=>summary[x[0]][x[1]]=(summary[x[0]][x[1]]||0)+1);const due=toDate_(t.SLA_Due_At);if(!due||isNaN(due.getTime()))return;if(due<=now){summary.overdue++;summary.oldestOverdue.push(t);}else{const remaining=calculateWorkingMinutesBetween_(now,due);if(remaining>0&&remaining<=warning)summary.dueSoon++;}});summary.oldestOverdue.sort((a,b)=>toDate_(a.SLA_Due_At)-toDate_(b.SLA_Due_At));
  const result=enqueueSlackNotification_({notificationType:'END_OF_DAY_SUMMARY',dedupeKey:'EOD:'+date,ticketId:'',slaCycleNumber:'',priority:'',payload:buildEndOfDaySlackPayload_(summary)}); if(result.enqueued)cleanupSlackNotificationHistory_();return result;
}
function sendSlackWebhookPayload_(payload) {
  const response=UrlFetchApp.fetch(getSlackWebhookUrl_(),{method:'post',contentType:'application/json',payload:JSON.stringify(payload),muteHttpExceptions:true});
  return {code:response.getResponseCode(),body:String(response.getContentText()||''),headers:response.getAllHeaders?response.getAllHeaders():{}};
}
function retryDelayMinutes_(attempt){return [1,5,15][Math.min(Math.max(attempt-1,0),2)];}
function updateClaimedSlackRow_(notificationId,changes) { const lock=LockService.getScriptLock();lock.waitLock(5000);try{const sheet=getRequiredSheet_(APP.SHEETS.SLACK_NOTIFICATIONS),headers=queueHeaders_(sheet),id=headers.indexOf('Notification_ID'),values=sheet.getLastRow()>1?sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues():[];let index=-1;values.some((row,i)=>{if(String(row[id])===String(notificationId)){index=i;return true;}return false;});if(index<0)return false;Object.keys(changes).forEach(key=>{const column=headers.indexOf(key);if(column>=0)values[index][column]=changes[key];});sheet.getRange(index+2,1,1,headers.length).setValues([values[index]]);return true;}finally{lock.releaseLock();}}
function dispatchSlackNotifications() {
  const started=Date.now(),metrics={rowsScanned:0,notificationsEnqueued:0,notificationsSent:0,notificationsFailed:0,lockWaitMs:0};if(!slackEnabled_())return metrics;
  let claims=[];const settings=getSettings_(),now=new Date(),batch=Math.max(1,number_(settings.SLACK_DISPATCH_BATCH_SIZE,10)),timeout=number_(settings.SLACK_PROCESSING_TIMEOUT_MINUTES,10)*60000,lock=LockService.getScriptLock(),wait=Date.now();
  try{lock.waitLock(5000);metrics.lockWaitMs=Date.now()-wait;const sheet=getRequiredSheet_(APP.SHEETS.SLACK_NOTIFICATIONS),headers=queueHeaders_(sheet),rows=queueObjects_(sheet);metrics.rowsScanned=rows.length;claims=rows.filter(row=>{const o=row.object,next=toDate_(o.Next_Attempt_At),processing=toDate_(o.Processing_Started_At);return(String(o.Status)==='PENDING'&&(!next||isNaN(next.getTime())||next<=now))||(String(o.Status)==='PROCESSING'&&processing&&!isNaN(processing.getTime())&&processing.getTime()<=now.getTime()-timeout);}).slice(0,batch);const status=headers.indexOf('Status'),processing=headers.indexOf('Processing_Started_At'),updated=headers.indexOf('Updated_At');claims.forEach(row=>{row.values[status]='PROCESSING';row.values[processing]=now;row.values[updated]=now;});if(claims.length)sheet.getRange(2,1,rows.length,headers.length).setValues(rows.map(row=>row.values));}catch(error){console.error('dispatchSlackNotifications claim failed: '+slackEscape_(error.message,500));return metrics;}finally{lock.releaseLock();}
  claims.forEach(row=>{const o=row.object,id=o.Notification_ID,attempt=number_(o.Attempts,0)+1,max=Math.max(1,number_(settings.SLACK_MAX_RETRIES,3));try{let payload;try{payload=JSON.parse(String(o.Payload_JSON||''));if(!payload||typeof payload!=='object')throw new Error('not an object');}catch(parseError){updateClaimedSlackRow_(id,{Status:'FAILED',Attempts:attempt,Last_Error:'Invalid Payload_JSON',Updated_At:new Date()});metrics.notificationsFailed++;return;}const result=sendSlackWebhookPayload_(payload),code=result.code;if(code>=200&&code<300){updateClaimedSlackRow_(id,{Status:'SENT',Attempts:attempt,Sent_At:new Date(),Last_HTTP_Code:code,Last_Error:'',Updated_At:new Date()});metrics.notificationsSent++;return;}const temporary=code===429||(code>=500&&code<=599),retry=temporary&&attempt<max;let delay=retryDelayMinutes_(attempt);if(code===429){const h=result.headers||{},raw=h['Retry-After']||h['retry-after'];delay=Math.max(delay,Math.ceil(number_(Array.isArray(raw)?raw[0]:raw,0)/60));}updateClaimedSlackRow_(id,{Status:retry?'PENDING':'FAILED',Attempts:attempt,Next_Attempt_At:retry?new Date(Date.now()+delay*60000):'',Last_HTTP_Code:code,Last_Error:slackEscape_('Slack HTTP '+code,500),Updated_At:new Date()});if(!retry)metrics.notificationsFailed++;}catch(error){const retry=attempt<max;updateClaimedSlackRow_(id,{Status:retry?'PENDING':'FAILED',Attempts:attempt,Next_Attempt_At:retry?new Date(Date.now()+retryDelayMinutes_(attempt)*60000):'',Last_HTTP_Code:'',Last_Error:slackEscape_('Slack network request failed: '+String(error&&error.message||error).replace(/https:\/\/\S+/g,'[redacted]'),500),Updated_At:new Date()});if(!retry)metrics.notificationsFailed++;}});
  console.log(JSON.stringify(Object.assign({functionName:'dispatchSlackNotifications',durationMs:Date.now()-started},metrics)));return metrics;
}
function cleanupSlackNotificationHistory_(){const props=PropertiesService.getScriptProperties(),today=localDateKey_(new Date());if(props.getProperty('SLACK_CLEANUP_DATE')===today)return{cleaned:0};const days=Math.max(1,number_(getSettings_().SLACK_NOTIFICATION_RETENTION_DAYS,30)),cutoff=Date.now()-days*86400000,lock=LockService.getScriptLock();lock.waitLock(5000);try{const sheet=getRequiredSheet_(APP.SHEETS.SLACK_NOTIFICATIONS),headers=queueHeaders_(sheet),status=headers.indexOf('Status'),updated=headers.indexOf('Updated_At');if(sheet.getLastRow()<2){props.setProperty('SLACK_CLEANUP_DATE',today);return{cleaned:0};}const rows=sheet.getRange(2,1,sheet.getLastRow()-1,headers.length).getValues(),keep=rows.filter(row=>['SENT','FAILED'].indexOf(String(row[status]))<0||!toDate_(row[updated])||toDate_(row[updated]).getTime()>=cutoff),cleaned=rows.length-keep.length;if(cleaned){sheet.getRange(2,1,rows.length,headers.length).clearContent();if(keep.length)sheet.getRange(2,1,keep.length,headers.length).setValues(keep);}props.setProperty('SLACK_CLEANUP_DATE',today);return{cleaned:cleaned};}finally{lock.releaseLock();}}
function setupSlackAutomationTriggers(){const handlers={dispatchSlackNotifications:1,monitorSlackAlerts:15},existing=ScriptApp.getProjectTriggers();Object.keys(handlers).forEach(name=>{const matching=existing.filter(t=>t.getHandlerFunction()===name);matching.slice(1).forEach(t=>ScriptApp.deleteTrigger(t));if(!matching.length)ScriptApp.newTrigger(name).timeBased().everyMinutes(handlers[name]).create();});return{handlers:[{name:'dispatchSlackNotifications',frequency:'every 1 minute'},{name:'monitorSlackAlerts',frequency:'every 15 minutes'}]};}
function removeSlackAutomationTriggers(){let removed={dispatchSlackNotifications:0,monitorSlackAlerts:0};ScriptApp.getProjectTriggers().forEach(t=>{const name=t.getHandlerFunction();if(Object.prototype.hasOwnProperty.call(removed,name)){ScriptApp.deleteTrigger(t);removed[name]++;}});return{removed:removed};}
function validateSlackAutomation(){const settings=getSettings_(),sheet=getSpreadsheet_().getSheetByName(APP.SHEETS.SLACK_NOTIFICATIONS),triggers=ScriptApp.getProjectTriggers(),dc=triggers.filter(t=>t.getHandlerFunction()==='dispatchSlackNotifications').length,mc=triggers.filter(t=>t.getHandlerFunction()==='monitorSlackAlerts').length,rows=sheet?queueObjects_(sheet):[],sent=rows.filter(r=>String(r.object.Status)==='SENT').map(r=>toDate_(r.object.Sent_At)).filter(d=>d&&!isNaN(d.getTime())).sort((a,b)=>b-a),warnings=[];let webhook=false;try{getSlackWebhookUrl_();webhook=true;}catch(e){warnings.push('SLACK_WEBHOOK_URL is not configured.');}if(dc!==1)warnings.push('Dispatcher trigger count must be one.');if(mc!==1)warnings.push('Monitor trigger count must be one.');const result={notificationsEnabled:truthy_(settings.SLACK_NOTIFICATIONS_ENABLED),webhookConfigured:webhook,queueSheetReady:Boolean(sheet),dispatcherTriggerPresent:dc>0,monitorTriggerPresent:mc>0,dispatcherTriggerCount:dc,monitorTriggerCount:mc,pendingCount:rows.filter(r=>String(r.object.Status)==='PENDING').length,failedCount:rows.filter(r=>String(r.object.Status)==='FAILED').length,lastSentAt:sent.length?sent[0].toISOString():'',schemaVersion:CURRENT_SCHEMA_VERSION,ready:false,warnings:warnings};result.ready=result.notificationsEnabled&&result.webhookConfigured&&result.queueSheetReady&&dc===1&&mc===1;return result;}
function testSlackConnection(){getSlackWebhookUrl_();const id=Utilities.getUuid(),key='TEST:'+id,result=enqueueSlackNotification_({notificationType:'TEST',dedupeKey:key,ticketId:'',slaCycleNumber:'',priority:'',payload:buildSlackTestPayload_()});if(result.enqueued)dispatchSlackNotifications();const row=queueObjects_(getRequiredSheet_(APP.SHEETS.SLACK_NOTIFICATIONS)).find(r=>String(r.object.Dedupe_Key)===key),status=row?String(row.object.Status):'FAILED';return{success:status==='SENT',notificationId:row?String(row.object.Notification_ID):'',status:status,httpCode:row?number_(row.object.Last_HTTP_Code,0):0};}
