/**
 * MAIN BACKEND FILE
 * Do not put passwords or the Slack webhook directly in this file.
 */

const APP_RELEASE = 'registration-v2';
const APP_COMMIT = '__APP_COMMIT__';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getBootstrap() {
  const user = requireUser_();
  const settings = getSettings_();
  return {
    user,
    clients: getActiveClients_(),
    categories: getActiveCategories_(),
    rootCauses: String(settings.ROOT_CAUSES || '').split('|').filter(Boolean),
    duplicateWindowDays: number_(settings.DUPLICATE_WINDOW_DAYS, 5),
    resolvedVisibilityDays: number_(settings.RESOLVED_VISIBILITY_DAYS, 10),
    dashboardWindowDays: number_(settings.DASHBOARD_WINDOW_DAYS, 14)
  };
}

function checkDuplicate(payload) {
  const user = requireUser_();
  validateDuplicatePayload_(payload);
  const settings = getSettings_();
  const days = number_(settings.DUPLICATE_WINDOW_DAYS, 5);
  const threshold = number_(settings.DUPLICATE_SIMILARITY_THRESHOLD, 0.65);
  const cutoff = new Date(Date.now() - days * 24 * 60 * 60 * 1000);
  const clientKey = buildClientKeyFromPayload_(payload);
  const normalizedSubject = normalizeSubject_(payload.emailSubject);

  const matches = getSheetObjects_(APP.SHEETS.TICKETS)
    .filter(t => toDate_(t.Created_At) >= cutoff)
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

  return { hasDuplicate: matches.length > 0, matches };
}

function submitTicket(form) {
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
    const slaDueAt = new Date(createdAt.getTime() + slaHours * 60 * 60 * 1000);
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
    appendEvent_(ticketId, 'TICKET_RAISED', '', APP.STATUS.RAISED, user.email, duplicateOverride ? 'Raised despite duplicate warning.' : '');
  } finally {
    lock.releaseLock();
  }

  const detail = getTicketDetail(ticketId);
  try {
    sendSlackAlert_(detail);
  } catch (err) {
    console.error('Slack alert failed: ' + err.message);
  }
  return detail;
}

function getMyTickets() {
  const user = requireUser_();
  const settings = getSettings_();
  const cutoff = new Date(Date.now() - number_(settings.RESOLVED_VISIBILITY_DAYS, 10) * 24 * 60 * 60 * 1000);

  return getSheetObjects_(APP.SHEETS.TICKETS)
    .filter(t => lower_(t.Raiser_Email) === user.email)
    .filter(t => String(t.Status) !== APP.STATUS.RESOLVED || toDate_(t.Resolved_At) >= cutoff)
    .sort((a, b) => toDate_(b.Created_At) - toDate_(a.Created_At))
    .map(serializeTicket_);
}

function getQueueTickets() {
  requireRole_([APP.ROLES.POC, APP.ROLES.ADMIN]);
  return getSheetObjects_(APP.SHEETS.TICKETS)
    .sort((a, b) => queueSort_(a, b))
    .map(serializeTicket_);
}

function getTicketDetail(ticketId) {
  const user = requireUser_();
  const found = findObjectRow_(APP.SHEETS.TICKETS, 'Ticket_ID', ticketId);
  if (!found) throw new Error('Ticket not found.');
  if (user.role === APP.ROLES.SALES && lower_(found.object.Raiser_Email) !== user.email) {
    throw new Error('You are not allowed to view this ticket.');
  }
  const result = serializeTicket_(found.object);
  result.dynamicFields = safeJsonParse_(found.object.Dynamic_Fields_JSON, {});
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
  return result;
}

function updateTicketStatus(payload) {
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

    if (oldStatus === APP.STATUS.RESOLVED) throw new Error('Resolved tickets are read-only in v1.');
    if (oldStatus === APP.STATUS.RAISED && newStatus !== APP.STATUS.INVESTIGATING) {
      throw new Error('A Raised ticket must first move to Investigating.');
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
    }

    updateObjectRow_(APP.SHEETS.TICKETS, found.rowNumber, changes);
    appendEvent_(payload.ticketId, 'STATUS_CHANGED', oldStatus, newStatus, user.email,
      newStatus === APP.STATUS.RESOLVED ? `${changes.Root_Cause}: ${changes.Resolution_Note}` : '');
  } finally {
    lock.releaseLock();
  }
  return getTicketDetail(payload.ticketId);
}

function getNumbers() {
  requireRole_([APP.ROLES.POC, APP.ROLES.ADMIN]);
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

  return {
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
  const email = lower_(Session.getActiveUser().getEmail());
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
  try {
    const existing = getSheetObjects_(APP.SHEETS.USERS).find(u => lower_(u.Email) === email);
    if (!existing) {
      appendObject_(APP.SHEETS.USERS, { Email: email, Name: name, Role: role, Active: true });
    }
  } finally {
    lock.releaseLock();
  }
  return getEntryState();
}

function requireUser_() {
  const email = getVerifiedCompanyEmail_();

  const row = getSheetObjects_(APP.SHEETS.USERS).find(u => lower_(u.Email) === email && truthy_(u.Active));
  if (!row) throw new Error('You do not have access yet. Ask the app administrator to add your email to the Users sheet.');

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
  if (!folderId) throw new Error('Attachment folder is not configured. Run setupSystem() again.');
  const folder = DriveApp.getFolderById(folderId);
  const original = cleanFileName_(blob.getName() || 'evidence');
  blob.setName(`${ticketId} - ${original}`);
  const file = folder.createFile(blob);
  file.setDescription(`Evidence for ${ticketId}. Uploaded by ${raiserEmail}.`);
  try { file.addViewer(raiserEmail); } catch (err) { console.warn('Could not add raiser as viewer: ' + err.message); }
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

  const response = UrlFetchApp.fetch(webhook, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ text }),
    muteHttpExceptions: true
  });
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error(`Slack returned ${response.getResponseCode()}: ${response.getContentText()}`);
  }
}

// -------------------- Sheet helpers --------------------

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('System is not configured. Run setupSystem() first.');
  return SpreadsheetApp.openById(id);
}

function getSettings_() {
  const settings = {};
  getSheetObjects_(APP.SHEETS.SETTINGS).forEach(r => { settings[String(r.Key)] = String(r.Value); });
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
  return getSheetObjects_(APP.SHEETS.CATEGORIES)
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
}

function getCategoryById_(id) {
  return getSheetObjects_(APP.SHEETS.CATEGORIES)
    .find(r => String(r.Category_ID) === String(id) && truthy_(r.Active));
}

function getSheetObjects_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Missing sheet: ${sheetName}`);
  const values = sheet.getDataRange().getValues();
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
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(String);
  const keyIndex = headers.indexOf(key);
  if (keyIndex < 0) throw new Error(`Column ${key} not found in ${sheetName}.`);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][keyIndex]) === String(value)) {
      const object = headers.reduce((obj, h, j) => { obj[h] = data[i][j]; return obj; }, {});
      return { rowNumber: i + 1, object };
    }
  }
  return null;
}

function updateObjectRow_(sheetName, rowNumber, changes) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  Object.keys(changes).forEach(key => {
    const col = headers.indexOf(key) + 1;
    if (!col) throw new Error(`Column ${key} not found in ${sheetName}.`);
    sheet.getRange(rowNumber, col).setValue(changes[key]);
  });
}

function appendEvent_(ticketId, type, oldValue, newValue, performedBy, note) {
  appendObject_(APP.SHEETS.EVENTS, {
    Event_ID: Utilities.getUuid(),
    Ticket_ID: ticketId,
    Event_Type: type,
    Old_Value: oldValue,
    New_Value: newValue,
    Performed_By: performedBy,
    Created_At: new Date(),
    Note: note || ''
  });
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
