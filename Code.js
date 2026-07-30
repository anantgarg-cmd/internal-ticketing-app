/**
 * MAIN BACKEND FILE
 * Do not put passwords or the Slack webhook directly in this file.
 */

const APP_RELEASE = 'registration-v2';
const APP_COMMIT = '__APP_COMMIT__';
let SPREADSHEET_INSTANCE_ = null;
const DEPLOYMENT_AUTHORIZATION_MESSAGE = 'The application deployment has not been authorized by its deploying account. Please ask the application administrator to run authorizeApplication() once from the Apps Script editor.';
const CACHE_KEYS_ = Object.freeze({
  SETTINGS: 'app:settings:v1',
  CATEGORIES: 'app:categories:v1',
  NUMBERS: 'app:numbers:v1'
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

  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email.endsWith('@shadowfax.in')) {
    throw new Error('Authorization must be completed by a @shadowfax.in deployment-owner account.');
  }

  return {
    authorized: true,
    spreadsheetAccessible: true,
    driveAccessible: true,
    externalRequestAccessible: true,
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

function buildBootstrap_(user) {
  const settings = getSettings_();
  return {
    user,
    categories: getActiveCategories_(),
    rootCauses: String(settings.ROOT_CAUSES || '').split('|').filter(Boolean),
    duplicateWindowDays: number_(settings.DUPLICATE_WINDOW_DAYS, 5),
    resolvedVisibilityDays: number_(settings.RESOLVED_VISIBILITY_DAYS, 10),
    dashboardWindowDays: number_(settings.DASHBOARD_WINDOW_DAYS, 14)
  };
}

function getInitialAppState() {
  const startedAt = Date.now();
  const email = getVerifiedCompanyEmail_();
  const users = getSheetObjects_(APP.SHEETS.USERS);
  const row = users.find(u => lower_(u.Email) === email);
  let result;
  if (!row) result = { state: 'REGISTER', email, release: APP_RELEASE };
  else if (!truthy_(row.Active)) result = { state: 'BLOCKED', email, message: 'Your access has been disabled. Please contact the application administrator.', release: APP_RELEASE };
  else {
    const user = userFromRow_(row, email);
    result = { state: 'ACTIVE', email, name: user.name, role: user.role, bootstrap: buildBootstrap_(user), release: APP_RELEASE };
  }
  logPerformance_('getInitialAppState', startedAt, { rows: users.length });
  return result;
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
  const matches = recent.rows
    .filter(t => buildClientKeyFromTicket_(t) === clientKey)
    .filter(t => String(t.Category_ID) === String(payload.categoryId))
    .map(t => ({ ticket: t, similarity: subjectSimilarity_(normalizedSubject, String(t.Normalized_Subject || '')) }))
    .filter(x => x.similarity >= threshold)
    .sort((a, b) => toDate_(b.ticket.Created_At) - toDate_(a.ticket.Created_At))
    .slice(0, 3)
    .map(x => ({
      ticketId: String(x.ticket.Ticket_ID),
      createdAt: formatDateTime_(x.ticket.Created_At),
      status: String(x.ticket.Status),
      subject: String(x.ticket.Email_Subject),
      similarity: Math.round(x.similarity * 100),
      canView: user.role !== APP.ROLES.SALES || lower_(x.ticket.Raiser_Email) === user.email
    }));

  logPerformance_('checkDuplicate', startedAt, { rows: recent.processed });
  return { hasDuplicate: matches.length > 0, matches };
}

function submitTicket(form) {
  const startedAt = Date.now();
  const user = requireUser_();
  if (!form || typeof form !== 'object') throw new Error('The form data was not received. Please refresh and try again.');

  const category = getCategoryById_(form.categoryId);
  if (!category) throw new Error('The selected category is no longer active. Refresh the page and choose again.');

  const client = resolveClient_(form, category.Client_Type);
  validateTicketForm_(form, category, client);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let ticketId;
  let attachment = { id: '', name: '', url: '' };
  try {
    ticketId = nextTicketId_();
    attachment = saveAttachment_(form.attachment, ticketId, user.email);

    const createdAt = new Date();
    const slaHours = number_(category.SLA_Hours, 24);
    const slaDueAt = calculateWorkingSlaDueAt_(createdAt, slaHours);
    const subject = cleanText_(form.emailSubject, 300);
    const dynamicFields = extractDynamicFields_(form, category);
    const duplicateIds = cleanText_(form.duplicateIds || '', 500);
    const duplicateOverride = String(form.duplicateOverride || '').toLowerCase() === 'true';

    const ticket = {
      Ticket_ID: ticketId,
      Created_At: createdAt,
      Raiser_Email: user.email,
      Raiser_Name: user.name,
      Client_Mode: client.mode,
      Client_ID: client.id,
      Client_Name: client.name,
      Client_Type: client.type,
      Category_ID: category.Category_ID,
      Category_Name: category.Category_Name,
      Email_Subject: subject,
      Normalized_Subject: normalizeSubject_(subject),
      Issue_Description: cleanText_(form.issueDescription, 5000),
      Priority: String(category.Priority).toUpperCase(),
      SLA_Hours: slaHours,
      SLA_Due_At: slaDueAt,
      Status: APP.STATUS.RAISED,
      Picked_Up_By: '',
      Picked_Up_At: '',
      Investigating_At: '',
      Resolution_Note: '',
      Root_Cause: '',
      Resolved_By: '',
      Resolved_At: '',
      SLA_Result: 'ON TRACK',
      Duplicate_Of: duplicateIds,
      Duplicate_Override: duplicateOverride,
      Dynamic_Fields_JSON: JSON.stringify(dynamicFields),
      Attachment_File_ID: attachment.id,
      Attachment_File_Name: attachment.name,
      Attachment_URL: attachment.url,
      Updated_At: createdAt,
      Updated_By: user.email
    };

    appendObject_(APP.SHEETS.TICKETS, ticket);
    appendSlaCycle_({
      Ticket_ID: ticketId, Cycle_Number: 1, Cycle_Type: 'INITIAL', Started_At: createdAt, Due_At: slaDueAt,
      SLA_Result: 'OPEN', Started_By: user.email, Created_At: createdAt, Updated_At: createdAt
    });
    appendEvent_(ticketId, 'TICKET_RAISED', '', APP.STATUS.RAISED, user.email, duplicateOverride ? 'Raised despite duplicate warning.' : '');
    removeCachedKeys_([CACHE_KEYS_.NUMBERS]);
  } finally {
    lock.releaseLock();
  }

  const detail = getTicketDetail(ticketId);
  try {
    sendSlackAlert_(detail);
  } catch (err) {
    if (String(err && err.message || '') === DEPLOYMENT_AUTHORIZATION_MESSAGE) throw err;
    console.error('Slack alert failed: ' + err.message);
  }
  logPerformance_('submitTicket', startedAt, { rows: 1 });
  return detail;
}

/** Returns a server-calculated SLA estimate for the Raise Ticket form. */
function getSlaDuePreview(categoryId) {
  requireUser_();
  const category = getCategoryById_(categoryId);
  if (!category) throw new Error('The selected category is no longer active. Refresh the page and choose again.');
  const dueAt = calculateWorkingSlaDueAt_(new Date(), number_(category.SLA_Hours, 24));
  return { dueAt: formatDateTime_(dueAt), dueAtIso: dueAt.toISOString() };
}

function getMyTickets() {
  const startedAt = Date.now();
  const user = requireUser_();
  const settings = getSettings_();
  const cutoff = new Date(Date.now() - number_(settings.RESOLVED_VISIBILITY_DAYS, 10) * 24 * 60 * 60 * 1000);

  const tickets = getSheetObjects_(APP.SHEETS.TICKETS);
  const result = tickets
    .filter(t => lower_(t.Raiser_Email) === user.email)
    .filter(t => String(t.Status) !== APP.STATUS.RESOLVED || toDate_(t.Resolved_At) >= cutoff)
    .sort((a, b) => toDate_(b.Created_At) - toDate_(a.Created_At))
    .map(serializeTicket_);
  logPerformance_('getMyTickets', startedAt, { rows: tickets.length });
  return result;
}

function getQueueTickets(filters) {
  const startedAt = Date.now();
  requireRole_([APP.ROLES.POC, APP.ROLES.ADMIN]);
  filters = filters || {};
  const pageSize = Math.min(100, Math.max(1, Math.floor(number_(filters.pageSize, 50))));
  const requestedPage = Math.max(1, Math.floor(number_(filters.page, 1)));
  const search = lower_(filters.search);
  const status = filters.status === undefined ? 'OPEN' : String(filters.status);
  const priority = String(filters.priority || '');
  const category = String(filters.category || '');
  const sla = String(filters.sla || '');
  const tickets = getSheetObjects_(APP.SHEETS.TICKETS);
  const filtered = tickets.filter(t => {
    const serialized = serializeQueueTicket_(t);
    return (!search || [serialized.ticketId, serialized.clientName, serialized.emailSubject, serialized.raiserEmail].join(' ').toLowerCase().includes(search)) &&
      (status === 'OPEN' ? serialized.status !== APP.STATUS.RESOLVED : (!status || serialized.status === status)) &&
      (!priority || serialized.priority === priority) && (!category || serialized.categoryName === category) && (!sla || serialized.slaStatus === sla);
  }).sort(queueSort_);
  const totalRows = filtered.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));
  const page = Math.min(requestedPage, totalPages);
  const rows = filtered.slice((page - 1) * pageSize, page * pageSize).map(serializeQueueTicket_);
  const result = { rows, page, pageSize, totalRows, totalPages, categories: getActiveCategories_().map(c => c.name).filter((v, i, a) => a.indexOf(v) === i).sort() };
  logPerformance_('getQueueTickets', startedAt, { rows: tickets.length });
  return result;
}

function getTicketDetail(ticketId) {
  const startedAt = Date.now();
  const user = requireUser_();
  const found = findObjectRow_(APP.SHEETS.TICKETS, 'Ticket_ID', ticketId);
  if (!found) throw new Error('Ticket not found.');
  if (user.role === APP.ROLES.SALES && lower_(found.object.Raiser_Email) !== user.email) {
    throw new Error('You are not allowed to view this ticket.');
  }
  const result = serializeTicket_(found.object);
  result.dynamicFields = safeJsonParse_(found.object.Dynamic_Fields_JSON, {});
  result.slaCycles = getTicketSlaCycles_(ticketId).map(serializeSlaCycle_);
  result.slaCycleNumber = result.slaCycles.length ? Math.max.apply(null, result.slaCycles.map(c => c.cycleNumber)) : 0;
  result.reopenCount = result.slaCycles.filter(c => c.cycleType === 'REOPEN').length;
  result.canReopen = String(found.object.Status) === APP.STATUS.RESOLVED &&
    (user.role === APP.ROLES.POC || user.role === APP.ROLES.ADMIN || lower_(found.object.Raiser_Email) === user.email);
  result.events = getSheetObjects_(APP.SHEETS.EVENTS)
    .filter(e => String(e.Ticket_ID) === String(ticketId))
    .sort((a, b) => toDate_(a.Created_At) - toDate_(b.Created_At))
    .map(e => ({
      eventType: String(e.Event_Type),
      oldValue: String(e.Old_Value || ''),
      newValue: String(e.New_Value || ''),
      performedBy: String(e.Performed_By || ''),
      createdAt: formatDateTime_(e.Created_At),
      note: String(e.Note || '')
    }));
  logPerformance_('getTicketDetail', startedAt, { rows: 1 });
  return result;
}

function updateTicketStatus(payload) {
  const startedAt = Date.now();
  const user = requireRole_([APP.ROLES.POC, APP.ROLES.ADMIN]);
  if (!payload || !payload.ticketId || !payload.newStatus) throw new Error('Ticket and new status are required.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const found = findObjectRow_(APP.SHEETS.TICKETS, 'Ticket_ID', payload.ticketId);
    if (!found) throw new Error('Ticket not found.');
    const ticket = found.object;
    const oldStatus = String(ticket.Status);
    const newStatus = String(payload.newStatus);

    const allowedTransitions = {};
    allowedTransitions[APP.STATUS.RAISED] = APP.STATUS.INVESTIGATING;
    allowedTransitions[APP.STATUS.REOPENED] = APP.STATUS.INVESTIGATING;
    allowedTransitions[APP.STATUS.INVESTIGATING] = APP.STATUS.RESOLVED;
    if (allowedTransitions[oldStatus] !== newStatus) throw new Error(`Transition from ${oldStatus} to ${newStatus} is not allowed.`);
    if (oldStatus === APP.STATUS.RESOLVED) throw new Error('Resolved tickets can only be reopened.');
    if ([APP.STATUS.RAISED, APP.STATUS.REOPENED].includes(oldStatus) && newStatus !== APP.STATUS.INVESTIGATING) {
      throw new Error(`A ${oldStatus} ticket must first move to Investigating.`);
    }
    if (oldStatus === APP.STATUS.INVESTIGATING && newStatus !== APP.STATUS.RESOLVED) {
      throw new Error('An Investigating ticket can only move to Resolved.');
    }

    const now = new Date();
    const changes = {
      Status: newStatus,
      Updated_At: now,
      Updated_By: user.email
    };

    if (newStatus === APP.STATUS.INVESTIGATING) {
      changes.Picked_Up_By = user.email;
      changes.Picked_Up_At = now;
      changes.Investigating_At = now;
    }

    if (newStatus === APP.STATUS.RESOLVED) {
      const note = cleanText_(payload.resolutionNote, 5000);
      const rootCause = cleanText_(payload.rootCause, 200);
      if (!note) throw new Error('Resolution note is mandatory.');
      if (!rootCause) throw new Error('Root cause is mandatory.');
      changes.Resolution_Note = note;
      changes.Root_Cause = rootCause;
      changes.Resolved_By = user.email;
      changes.Resolved_At = now;
      changes.SLA_Result = now <= toDate_(ticket.SLA_Due_At) ? 'MET' : 'BREACHED';
      ensureOpenInitialCycle_(ticket);
      closeOpenSlaCycle_(payload.ticketId, now, user.email, changes.SLA_Result);
    }

    updateObjectRow_(APP.SHEETS.TICKETS, found.rowNumber, changes);
    appendEvent_(payload.ticketId, 'STATUS_CHANGED', oldStatus, newStatus, user.email,
      newStatus === APP.STATUS.RESOLVED ? `${changes.Root_Cause}: ${changes.Resolution_Note}` : '');
    removeCachedKeys_([CACHE_KEYS_.NUMBERS]);
  } finally {
    lock.releaseLock();
  }
  const result = getTicketDetail(payload.ticketId);
  logPerformance_('updateTicketStatus', startedAt, { rows: 1 });
  return result;
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
function reopenTicket(payload) {
  const startedAt = Date.now();
  const user = requireUser_();
  if (!payload || !payload.ticketId) throw new Error('Ticket is required.');
  const reason = cleanText_(payload.reopenReason, 5001);
  if (!reason) throw new Error('Reason for reopening / latest client response is mandatory.');
  if (reason.length > 5000) throw new Error('Reopening reason must be 5,000 characters or fewer.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let reopened;
  try {
    const found = findObjectRow_(APP.SHEETS.TICKETS, 'Ticket_ID', payload.ticketId);
    if (!found) throw new Error('Ticket not found.');
    const ticket = found.object;
    assertCanReopen_(user, ticket);
    ensureLegacyInitialCycle_(ticket);
    const cycles = getTicketSlaCycles_(payload.ticketId);
    if (cycles.some(cycle => String(cycle.SLA_Result) === 'OPEN')) throw new Error('This ticket already has an open SLA cycle.');
    const reopenedAt = new Date();
    const dueAt = calculateWorkingSlaDueAt_(reopenedAt, number_(ticket.SLA_Hours, 0));
    const cycleNumber = cycles.reduce((max, cycle) => Math.max(max, number_(cycle.Cycle_Number, 0)), 0) + 1;
    updateObjectRow_(APP.SHEETS.TICKETS, found.rowNumber, {
      Status: APP.STATUS.REOPENED, SLA_Due_At: dueAt, SLA_Result: '', Updated_At: reopenedAt, Updated_By: user.email,
      Picked_Up_By: '', Picked_Up_At: '', Investigating_At: '', Resolution_Note: '', Root_Cause: '', Resolved_By: '', Resolved_At: ''
    });
    appendSlaCycle_({
      Ticket_ID: payload.ticketId, Cycle_Number: cycleNumber, Cycle_Type: 'REOPEN', Started_At: reopenedAt,
      Due_At: dueAt, SLA_Result: 'OPEN', Started_By: user.email, Reopen_Reason: reason,
      Created_At: reopenedAt, Updated_At: reopenedAt
    });
    appendEvent_(payload.ticketId, 'TICKET_REOPENED', APP.STATUS.RESOLVED, APP.STATUS.REOPENED, user.email, reason, reopenedAt);
    removeCachedKeys_([CACHE_KEYS_.NUMBERS]);
    reopened = { ticketId: payload.ticketId, reopenedAt };
  } finally { lock.releaseLock(); }
  const detail = getTicketDetail(reopened.ticketId);
  try { sendSlackReopenedAlert_(detail, user.email, reason); } catch (err) {
    if (String(err && err.message || '') === DEPLOYMENT_AUTHORIZATION_MESSAGE) throw err;
    console.error('Slack reopen alert failed: ' + err.message);
  }
  logPerformance_('reopenTicket', startedAt, { rows: 1 });
  return detail;
}

function assertCanReopen_(user, ticket) {
  if (String(ticket.Status) !== APP.STATUS.RESOLVED) throw new Error('Only a Resolved ticket can be reopened.');
  if (![APP.ROLES.POC, APP.ROLES.ADMIN].includes(user.role) && lower_(ticket.Raiser_Email) !== user.email) {
    throw new Error('You are not allowed to reopen another user\'s ticket.');
  }
}

function getNumbers() {
  const startedAt = Date.now();
  requireRole_([APP.ROLES.POC, APP.ROLES.ADMIN]);
  const cached = getCachedJson_(CACHE_KEYS_.NUMBERS);
  if (cached !== null) {
    logPerformance_('getNumbers', startedAt, { cache: 'hit' });
    return cached;
  }
  const settings = getSettings_();
  const days = number_(settings.DASHBOARD_WINDOW_DAYS, 14);
  const cutoff = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  const now = new Date();
  const tickets = getSheetObjects_(APP.SHEETS.TICKETS);

  const raisedInWindow = tickets.filter(t => toDate_(t.Created_At) >= cutoff);
  const resolvedInWindow = tickets.filter(t => t.Resolved_At && toDate_(t.Resolved_At) >= cutoff);
  const open = tickets.filter(t => String(t.Status) !== APP.STATUS.RESOLVED);
  const overdue = open.filter(t => toDate_(t.SLA_Due_At) < now);
  const met = resolvedInWindow.filter(t => toDate_(t.Resolved_At) <= toDate_(t.SLA_Due_At)).length;
  const adherence = resolvedInWindow.length ? Math.round((met / resolvedInWindow.length) * 1000) / 10 : null;

  const result = {
    days,
    raised: raisedInWindow.length,
    resolved: resolvedInWindow.length,
    open: open.length,
    overdue: overdue.length,
    adherence,
    byStatus: countBy_(open, 'Status'),
    byPriority: countBy_(open, 'Priority'),
    byCategory: countBy_(open, 'Category_Name')
  };
  putCachedJson_(CACHE_KEYS_.NUMBERS, result, 60);
  logPerformance_('getNumbers', startedAt, { rows: tickets.length, cache: 'miss' });
  return result;
}

function setSlackWebhookFromEditor() {
  // Replace the placeholder temporarily, run once, and then remove the URL from the code before saving again.
  const webhookUrl = 'PASTE_SLACK_WEBHOOK_URL_HERE';
  if (!webhookUrl.startsWith('https://hooks.slack.com/')) throw new Error('Paste a valid Slack incoming webhook URL first.');
  PropertiesService.getScriptProperties().setProperty('SLACK_WEBHOOK_URL', webhookUrl);
  return 'Slack webhook saved securely in Script Properties. Remove it from this function now.';
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

function requireUser_() {
  const email = getVerifiedCompanyEmail_();

  const row = getSheetObjects_(APP.SHEETS.USERS).find(u => lower_(u.Email) === email && truthy_(u.Active));
  if (!row) throw new Error('You do not have access yet. Ask the app administrator to add your email to the Users sheet.');

  return userFromRow_(row, email);
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
  const settings = getSettings_();
  const alertPriorities = String(settings.ALERT_PRIORITIES || 'HIGH')
    .split(',').map(x => x.trim().toUpperCase()).filter(Boolean);
  if (!alertPriorities.includes(String(ticket.priority).toUpperCase())) return;

  const props = PropertiesService.getScriptProperties();
  const webhook = props.getProperty('SLACK_WEBHOOK_URL');
  if (!webhook) return;
  const appUrl = props.getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
  const due = ticket.slaDueAt;
  const text = [
    `🚨 *${ticket.priority}-Priority Internal Ticket Raised*`,
    `*Ticket:* ${ticket.ticketId}`,
    `*Client:* ${ticket.clientName} (${ticket.clientType})`,
    `*Category:* ${ticket.categoryName}`,
    `*Subject:* ${ticket.emailSubject}`,
    `*Raised by:* ${ticket.raiserEmail}`,
    `*Resolution SLA:* ${ticket.slaHours} hours`,
    `*Due by:* ${due}`,
    appUrl ? `*Open app:* ${appUrl}` : ''
  ].filter(Boolean).join('\n');

  let response;
  try {
    response = UrlFetchApp.fetch(webhook, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ text }),
      muteHttpExceptions: true
    });
  } catch (err) {
    throwServiceAuthorizationError_(err);
  }
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`Slack returned ${response.getResponseCode()}: ${response.getContentText()}`);
  }
}

function sendSlackReopenedAlert_(ticket, reopenedBy, reason) {
  const webhook = PropertiesService.getScriptProperties().getProperty('SLACK_WEBHOOK_URL');
  if (!webhook) return;
  const appUrl = PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
  const text = [
    '🔄 *Internal Ticket Reopened*', `*Ticket:* ${ticket.ticketId}`, `*Client:* ${ticket.clientName}`,
    `*Category:* ${ticket.categoryName}`, `*Priority:* ${ticket.priority}`, `*Reopened by:* ${reopenedBy}`,
    `*Reason:* ${reason}`, `*New SLA due:* ${ticket.slaDueAt}`, appUrl ? `*Open app:* ${appUrl}` : ''
  ].filter(Boolean).join('\n');
  let response;
  try {
    response = UrlFetchApp.fetch(webhook, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text }), muteHttpExceptions: true });
  } catch (err) { throwServiceAuthorizationError_(err); }
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) throw new Error(`Slack returned ${response.getResponseCode()}.`);
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

function getCategoryById_(id) {
  return getSheetObjects_(APP.SHEETS.CATEGORIES)
    .find(r => String(r.Category_ID) === String(id) && truthy_(r.Active));
}

function getSheetObjects_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Missing sheet: ${sheetName}`);
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
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  sheet.appendRow(headers.map(h => Object.prototype.hasOwnProperty.call(object, h) ? object[h] : ''));
}

function findObjectRow_(sheetName, key, value) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
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
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  Object.keys(changes).forEach(key => { if (headers.indexOf(key) < 0) throw new Error(`Column ${key} not found in ${sheetName}.`); });
  const range = sheet.getRange(rowNumber, 1, 1, headers.length);
  const row = range.getValues()[0];
  Object.keys(changes).forEach(key => { row[headers.indexOf(key)] = changes[key]; });
  range.setValues([row]);
}

function appendEvent_(ticketId, type, oldValue, newValue, performedBy, note, createdAt) {
  appendObject_(APP.SHEETS.EVENTS, {
    Event_ID: Utilities.getUuid(),
    Ticket_ID: ticketId,
    Event_Type: type,
    Old_Value: oldValue,
    New_Value: newValue,
    Performed_By: performedBy,
    Created_At: createdAt || new Date(),
    Note: note || ''
  });
}

function appendSlaCycle_(cycle) {
  appendObject_(APP.SHEETS.SLA_CYCLES, Object.assign({ SLA_Cycle_ID: Utilities.getUuid() }, cycle));
}

function getTicketSlaCycles_(ticketId) {
  return getSheetObjects_(APP.SHEETS.SLA_CYCLES)
    .filter(cycle => String(cycle.Ticket_ID) === String(ticketId))
    .sort((a, b) => number_(b.Cycle_Number, 0) - number_(a.Cycle_Number, 0));
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

function serializeTicket_(t) {
  const now = new Date();
  const due = toDate_(t.SLA_Due_At);
  const resolved = String(t.Status) === APP.STATUS.RESOLVED;
  let liveSla = String(t.SLA_Result || '');
  if (!resolved) liveSla = due < now ? 'OVERDUE' : 'ON TRACK';

  return {
    ticketId: String(t.Ticket_ID),
    createdAt: formatDateTime_(t.Created_At),
    createdAtIso: toDate_(t.Created_At).toISOString(),
    raiserEmail: String(t.Raiser_Email),
    raiserName: String(t.Raiser_Name || ''),
    clientId: String(t.Client_ID || ''),
    clientName: String(t.Client_Name),
    clientType: String(t.Client_Type),
    clientMode: String(t.Client_Mode),
    categoryId: String(t.Category_ID),
    categoryName: String(t.Category_Name),
    emailSubject: String(t.Email_Subject),
    issueDescription: String(t.Issue_Description),
    priority: String(t.Priority),
    slaHours: number_(t.SLA_Hours, 0),
    slaDueAt: formatDateTime_(t.SLA_Due_At),
    slaDueAtIso: due.toISOString(),
    slaStatus: liveSla,
    status: String(t.Status),
    pickedUpBy: String(t.Picked_Up_By || ''),
    pickedUpAt: formatDateTimeOptional_(t.Picked_Up_At),
    investigatingAt: formatDateTimeOptional_(t.Investigating_At),
    resolutionNote: String(t.Resolution_Note || ''),
    rootCause: String(t.Root_Cause || ''),
    resolvedBy: String(t.Resolved_By || ''),
    resolvedAt: formatDateTimeOptional_(t.Resolved_At),
    duplicateOf: String(t.Duplicate_Of || ''),
    duplicateOverride: truthy_(t.Duplicate_Override),
    attachmentFileName: String(t.Attachment_File_Name || ''),
    attachmentUrl: String(t.Attachment_URL || ''),
    updatedAt: formatDateTimeOptional_(t.Updated_At)
  };
}

function serializeQueueTicket_(t) {
  const full = serializeTicket_(t);
  return {
    ticketId: full.ticketId, createdAt: full.createdAt, createdAtIso: full.createdAtIso,
    clientName: full.clientName, categoryName: full.categoryName, emailSubject: full.emailSubject,
    priority: full.priority, status: full.status, slaStatus: full.slaStatus,
    slaDueAt: full.slaDueAt, slaDueAtIso: full.slaDueAtIso, raiserEmail: full.raiserEmail
  };
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

function getRecentTicketObjects_(cutoff, batchSize) {
  const sheet = getSpreadsheet_().getSheetByName(APP.SHEETS.TICKETS);
  if (!sheet) throw new Error(`Missing sheet: ${APP.SHEETS.TICKETS}`);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return { rows: [], processed: 0 };
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(String);
  const createdIndex = headers.indexOf('Created_At');
  if (createdIndex < 0) throw new Error(`Column Created_At not found in ${APP.SHEETS.TICKETS}.`);
  const rows = [];
  let endRow = lastRow;
  let processed = 0;
  let previousDate = null;
  let mustFallback = false;
  while (endRow >= 2) {
    const count = Math.min(batchSize || 200, endRow - 1);
    const startRow = endRow - count + 1;
    const values = sheet.getRange(startRow, 1, count, lastColumn).getValues();
    processed += values.length;
    let reachedCutoff = false;
    for (let i = values.length - 1; i >= 0; i--) {
      const rawDate = values[i][createdIndex];
      const date = rawDate instanceof Date ? rawDate : new Date(rawDate);
      if (isNaN(date.getTime()) || (previousDate && date > previousDate)) { mustFallback = true; break; }
      previousDate = date;
      if (date < cutoff) { reachedCutoff = true; break; }
      rows.push(headers.reduce((obj, h, j) => { obj[h] = values[i][j]; return obj; }, {}));
    }
    if (mustFallback) break;
    if (reachedCutoff) return { rows, processed };
    endRow = startRow - 1;
  }
  if (mustFallback) {
    const all = getSheetObjects_(APP.SHEETS.TICKETS);
    return { rows: all.filter(t => toDate_(t.Created_At) >= cutoff), processed: all.length };
  }
  return { rows, processed };
}

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
  console.log(JSON.stringify(safe));
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength || 5000);
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
