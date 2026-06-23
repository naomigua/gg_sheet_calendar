// Create / Modify / Delete events from Google Sheets using checkbox columns Create / Modify / Delete.
// - Trigger reacts only when Create/Modify/Delete column is edited (or runAllPending is executed).
// - After success, checkbox is unchecked. On error, Processing Log and Last Updated contain the error.
// Usage:
// 1) Coller dans Extensions -> Apps Script, Enregistrer.
// 2) Exécuter setup() (une fois) pour créer le trigger et autoriser.
// 3) Ajouter colonnes Create, Modify, Delete (case à cocher) dans l'en-tête.
// 4) Remplir ligne, cocher Create/Modify/Delete, ou exécuter runAllPending().

const SHEET_NAME = 'Events';
const HEADER_ROW = 1;

const HEADER_ALIASES = {
  'title': ['title','titre','sujet','summary'],
  'start date': ['start date','date debut','date début','date_start','startdate','date'],
  'start time': ['start time','heure debut','heure début','start_time','starttime','time'],
  'end date': ['end date','date fin','date_fin','enddate'],
  'end time': ['end time','heure fin','end_time','endtime'],
  'all day': ['all day','journée entière','journee entiere','allday','all-day','full day'],
  'description': ['description','details','détails'],
  'location': ['location','lieu','emplacement'],
  'guests': ['guests','invités','invites','participants','emails'],
  'calendar id': ['calendar id','id calendrier','id calendar','calendarid'],
  'status': ['status','statut','etat'],
  'event id': ['event id','id événement','id evenement','idevent','eventid'],
  'processing log': ['processing log','journal','journal traitement','processing','log','traitement'],
  'last updated': ['last updated','derniere mise a jour','dernière mise à jour','lastupdated','updated','horodatage','timestamp'],
  'create': ['create','ready','a creer','à créer','a_creer','ready to create','go','créer','creer'],
  'modify': ['modify','update','edit','modifier','mettre à jour','maj'],
  'delete': ['delete','remove','supprimer','effacer']
};

const ALIAS_TO_CANONICAL = buildAliasMap();

function buildAliasMap() {
  const map = {};
  for (const canon of Object.keys(HEADER_ALIASES)) {
    HEADER_ALIASES[canon].forEach(a => map[a.toString().trim().toLowerCase()] = canon);
  }
  return map;
}

// --- Setup trigger ---
function setup() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === 'onEditInstallable' && t.getEventType() === ScriptApp.EventType.ON_EDIT);
  if (!exists) {
    ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(ss).onEdit().create();
    Logger.log('Trigger onEditInstallable créé.');
  } else {
    Logger.log('Trigger onEditInstallable déjà présent.');
  }
  ss.getName(); // force permission prompt si nécessaire
  Logger.log('Setup terminé — autorisez si demandé.');
}

function removeOnEditTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const t of triggers) {
    if (t.getHandlerFunction() === 'onEditInstallable' && t.getEventType() === ScriptApp.EventType.ON_EDIT) {
      ScriptApp.deleteTrigger(t);
      Logger.log('Trigger onEditInstallable supprimé.');
    }
  }
}

// onEdit installable: réagit uniquement aux modifications dans les colonnes Create/Modify/Delete.
// Si appelé manuellement (sans e) -> runAllPending() pour test.
function onEditInstallable(e) {
  if (!e || !e.range) {
    Logger.log('onEditInstallable appelé manuellement — exécution de runAllPending() pour test.');
    try { runAllPending(); } catch (err) { Logger.log('Erreur runAllPending: ' + err); }
    return;
  }

  try {
    const range = e.range;
    const sheet = range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    const headerValues = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerIndex = buildHeaderIndex(headerValues);
    const createCol = headerIndex['create'];
    const modifyCol = headerIndex['modify'];
    const deleteCol = headerIndex['delete'];

    // si aucune des colonnes n'existe, ne rien faire
    if (!createCol && !modifyCol && !deleteCol) return;

    // on n'agit que si l'édition a lieu dans l'une des colonnes concernées
    const editedCol = range.getColumn();
    if (editedCol !== createCol && editedCol !== modifyCol && editedCol !== deleteCol) return;

    const newValue = sheet.getRange(range.getRow(), editedCol).getValue();
    const isChecked = parseBoolean(newValue);
    if (!isChecked) return; // n'agir que lors du passage à true

    // Determine action: if multiple boxes are checked, priority Delete > Modify > Create
    // If user edited a specific checkbox and set it true, handle that one.
    if (editedCol === deleteCol) {
      handleDelete(sheet, range.getRow(), headerIndex);
    } else if (editedCol === modifyCol) {
      handleModify(sheet, range.getRow(), headerIndex);
    } else if (editedCol === createCol) {
      handleCreate(sheet, range.getRow(), headerIndex);
    }
  } catch (err) {
    Logger.log('Erreur onEditInstallable: ' + err);
  }
}

// runAllPending: traite toutes les lignes qui ont Create/Modify/Delete cochés.
// Priorité: Delete → Modify → Create pour chaque ligne.
function runAllPending() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Feuille introuvable: ' + SHEET_NAME);
  ensureLogColumns(sheet);

  const headerValues = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerIndex = buildHeaderIndex(headerValues);
  const createCol = headerIndex['create'];
  const modifyCol = headerIndex['modify'];
  const deleteCol = headerIndex['delete'];

  const lastRow = sheet.getLastRow();
  for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
    try {
      // prioritize delete
      if (deleteCol && parseBoolean(sheet.getRange(r, deleteCol).getValue())) {
        handleDelete(sheet, r, headerIndex);
        continue;
      }
      if (modifyCol && parseBoolean(sheet.getRange(r, modifyCol).getValue())) {
        handleModify(sheet, r, headerIndex);
        continue;
      }
      if (createCol && parseBoolean(sheet.getRange(r, createCol).getValue())) {
        handleCreate(sheet, r, headerIndex);
        continue;
      }
    } catch (err) {
      Logger.log('Erreur runAllPending ligne ' + r + ': ' + err);
    }
  }
  Logger.log('runAllPending terminé.');
}

// Debug : traiter une ligne spécifique
function debugRow(row) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Feuille introuvable: ' + SHEET_NAME);
  ensureLogColumns(sheet);
  const headerValues = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerIndex = buildHeaderIndex(headerValues);
  // Decide what to do based on checked boxes priority
  const createCol = headerIndex['create'];
  const modifyCol = headerIndex['modify'];
  const deleteCol = headerIndex['delete'];
  if (deleteCol && parseBoolean(sheet.getRange(row, deleteCol).getValue())) {
    handleDelete(sheet, row, headerIndex); return;
  }
  if (modifyCol && parseBoolean(sheet.getRange(row, modifyCol).getValue())) {
    handleModify(sheet, row, headerIndex); return;
  }
  if (createCol && parseBoolean(sheet.getRange(row, createCol).getValue())) {
    handleCreate(sheet, row, headerIndex); return;
  }
  Logger.log('debugRow: aucune action détectée pour la ligne ' + row);
}

// Ensure Processing Log / Last Updated columns exist
function ensureLogColumns(sheet) {
  const headerValues = sheet.getRange(HEADER_ROW, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  const index = buildHeaderIndex(headerValues);
  let modified = false;
  const toEnsure = ['processing log','last updated','create','modify','delete'];
  let lastCol = sheet.getLastColumn();
  toEnsure.forEach(name => {
    if (!index[name]) {
      lastCol++;
      sheet.getRange(HEADER_ROW, lastCol).setValue(prettyHeader(name));
      modified = true;
    }
  });
  if (modified) Utilities.sleep(200);
}

function prettyHeader(canon) {
  if (canon === 'processing log') return 'Processing Log';
  if (canon === 'last updated') return 'Last Updated';
  if (canon === 'create') return 'Create';
  if (canon === 'modify') return 'Modify';
  if (canon === 'delete') return 'Delete';
  return canon;
}

function buildHeaderIndex(headerRowValues) {
  const map = {};
  for (let i = 0; i < headerRowValues.length; i++) {
    const v = headerRowValues[i];
    if (!v) continue;
    const norm = v.toString().trim().toLowerCase();
    const canon = ALIAS_TO_CANONICAL[norm];
    if (canon) {
      map[canon] = i + 1;
      continue;
    }
    for (const alias in ALIAS_TO_CANONICAL) {
      if (norm.indexOf(alias) !== -1) {
        map[ALIAS_TO_CANONICAL[alias]] = i + 1;
        break;
      }
    }
  }
  return map;
}

// ---- Handlers for Create / Modify / Delete ----

function handleCreate(sheet, row, headerIndex) {
  // basic validation
  const title = getCell(sheet, row, headerIndex['title']);
  const startDate = getCell(sheet, row, headerIndex['start date']);
  if (!title || !startDate) {
    writeProcessingResult(sheet, row, headerIndex, 'Waiting: missing Title or Start Date', 'warning');
    return;
  }
  try {
    const event = createOrUpdateEventFromRow(sheet, row, headerIndex, /*isCreate=*/true, /*isModify=*/false);
    // uncheck Create
    if (headerIndex['create']) sheet.getRange(row, headerIndex['create']).setValue(false);
    if (headerIndex['status']) sheet.getRange(row, headerIndex['status']).setValue('Created');
    writeProcessingResult(sheet, row, headerIndex, 'Created: ' + event.getId(), 'success');
    Logger.log('Created event for row ' + row + ': ' + event.getId());
  } catch (err) {
    const em = err && err.message ? err.message : String(err);
    if (headerIndex['status']) sheet.getRange(row, headerIndex['status']).setValue('Error: ' + em);
    writeProcessingResult(sheet, row, headerIndex, 'Error: ' + em, 'error');
    Logger.log('Error create row ' + row + ': ' + em);
  }
}

function handleModify(sheet, row, headerIndex) {
  const eventId = getCell(sheet, row, headerIndex['event id']);
  if (!eventId) {
    writeProcessingResult(sheet, row, headerIndex, 'Error: missing Event ID for Modify', 'error');
    return;
  }
  try {
    const event = CalendarApp.getEventById(eventId);
    if (!event) throw new Error('Event not found by id: ' + eventId);
    // Update event fields according to row
    createOrUpdateEventFromRow(sheet, row, headerIndex, /*isCreate=*/false, /*isModify=*/true);
    // uncheck Modify
    if (headerIndex['modify']) sheet.getRange(row, headerIndex['modify']).setValue(false);
    if (headerIndex['status']) sheet.getRange(row, headerIndex['status']).setValue('Modified');
    writeProcessingResult(sheet, row, headerIndex, 'Modified: ' + eventId, 'success');
    Logger.log('Modified event for row ' + row + ': ' + eventId);
  } catch (err) {
    const em = err && err.message ? err.message : String(err);
    if (headerIndex['status']) sheet.getRange(row, headerIndex['status']).setValue('Error: ' + em);
    writeProcessingResult(sheet, row, headerIndex, 'Error: ' + em, 'error');
    Logger.log('Error modify row ' + row + ': ' + em);
  }
}

function handleDelete(sheet, row, headerIndex) {
  const eventId = getCell(sheet, row, headerIndex['event id']);
  if (!eventId) {
    writeProcessingResult(sheet, row, headerIndex, 'Error: missing Event ID for Delete', 'error');
    return;
  }
  try {
    const event = CalendarApp.getEventById(eventId);
    if (!event) throw new Error('Event not found by id: ' + eventId);
    event.deleteEvent();
    // uncheck Delete
    if (headerIndex['delete']) sheet.getRange(row, headerIndex['delete']).setValue(false);
    if (headerIndex['status']) sheet.getRange(row, headerIndex['status']).setValue('Deleted');
    writeProcessingResult(sheet, row, headerIndex, 'Deleted: ' + eventId, 'success');
    Logger.log('Deleted event for row ' + row + ': ' + eventId);
  } catch (err) {
    const em = err && err.message ? err.message : String(err);
    if (headerIndex['status']) sheet.getRange(row, headerIndex['status']).setValue('Error: ' + em);
    writeProcessingResult(sheet, row, headerIndex, 'Error: ' + em, 'error');
    Logger.log('Error delete row ' + row + ': ' + em);
  }
}

// Utility that creates or updates event from row. If isCreate true -> creates event and writes Event ID.
// If isModify true -> updates existing event (uses event id).
function createOrUpdateEventFromRow(sheet, row, headerIndex, isCreate, isModify) {
  const title = getCell(sheet, row, headerIndex['title']);
  const startDate = getCell(sheet, row, headerIndex['start date']);
  const allDay = parseBoolean(getCell(sheet, row, headerIndex['all day']));
  const startTime = getCell(sheet, row, headerIndex['start time']);
  const endDate = getCell(sheet, row, headerIndex['end date']);
  const endTime = getCell(sheet, row, headerIndex['end time']);
  const description = getCell(sheet, row, headerIndex['description']);
  const location = getCell(sheet, row, headerIndex['location']);
  const guests = getCell(sheet, row, headerIndex['guests']);
  const calendarId = getCell(sheet, row, headerIndex['calendar id']);

  const calendar = calendarId ? CalendarApp.getCalendarById(calendarId) : CalendarApp.getDefaultCalendar();
  if (!calendar) throw new Error('Calendrier introuvable : ' + calendarId);

  let event;
  if (isCreate) {
    if (allDay) {
      const s = toDateOnly(startDate);
      const e = endDate ? toDateOnly(endDate) : s;
      const exclusiveEnd = new Date(e); exclusiveEnd.setDate(exclusiveEnd.getDate() + 1);
      event = calendar.createAllDayEvent(title, s, exclusiveEnd, {location: location, description: description});
    } else {
      const sdt = combineDateTime(startDate, startTime) || new Date(startDate);
      let edt = combineDateTime(endDate || startDate, endTime);
      if (!edt) edt = new Date(sdt.getTime() + 60*60*1000);
      event = calendar.createEvent(title, sdt, edt, {location: location, description: description});
    }
    // add guests
    if (guests) {
      const emails = guests.toString().split(/[;,]/).map(s => s.trim()).filter(Boolean);
      emails.forEach(email => { try { event.addGuest(email); } catch(e){Logger.log('addGuest failed: '+e);} });
    }
    // write event id in sheet
    if (headerIndex['event id']) sheet.getRange(row, headerIndex['event id']).setValue(event.getId());
    return event;
  } else if (isModify) {
    const eventId = getCell(sheet, row, headerIndex['event id']);
    if (!eventId) throw new Error('Missing event id for modify');
    event = CalendarApp.getEventById(eventId);
    if (!event) throw new Error('Event not found by id: ' + eventId);
    // update basic fields
    try { event.setTitle(title); } catch(e){ Logger.log('setTitle failed: '+e); }
    try {
      if (allDay) {
        const s = toDateOnly(startDate);
        const e = endDate ? toDateOnly(endDate) : s;
        const exclusiveEnd = new Date(e); exclusiveEnd.setDate(exclusiveEnd.getDate() + 1);
        // setTime can be used; for all-day we set midnight to midnight (best-effort)
        event.setAllDayDates ? event.setAllDayDates(s, exclusiveEnd) : event.setTime(s, exclusiveEnd);
      } else {
        const sdt = combineDateTime(startDate, startTime) || new Date(startDate);
        let edt = combineDateTime(endDate || startDate, endTime);
        if (!edt) edt = new Date(sdt.getTime() + 60*60*1000);
        event.setTime(sdt, edt);
      }
    } catch(e){ Logger.log('setTime/setAllDay failed: '+e); }
    try { event.setDescription(description || ''); } catch(e){ Logger.log('setDescription failed: '+e); }
    try { event.setLocation(location || ''); } catch(e){ Logger.log('setLocation failed: '+e); }

    // Guests: sync guest list (remove those not present, add new)
    if (guests) {
      const newGuests = guests.toString().split(/[;,]/).map(s => s.trim()).filter(Boolean);
      try {
        const currentGuests = (event.getGuestList && typeof event.getGuestList === 'function') ? event.getGuestList() : [];
        const currentEmails = currentGuests.map(g => (g.getEmail ? g.getEmail() : (g.email || ''))).filter(Boolean);
        // remove those not in newGuests
        currentEmails.forEach(email => {
          if (newGuests.indexOf(email) === -1) {
            try { event.removeGuest ? event.removeGuest(email) : CalendarApp.getEventById(event.getId()).removeGuest(email); } catch(e){Logger.log('removeGuest failed: '+e);}
          }
        });
        // add missing
        newGuests.forEach(email => {
          if (currentEmails.indexOf(email) === -1) {
            try { event.addGuest(email); } catch(e){Logger.log('addGuest failed: '+e);}
          }
        });
      } catch(e) {
        Logger.log('Guest sync failed: ' + e);
      }
    }
    return event;
  } else {
    throw new Error('createOrUpdateEventFromRow appelé sans mode valide');
  }
}

// ---- Helpers ----

function getCell(sheet, row, col) {
  if (!col) return '';
  const val = sheet.getRange(row, col).getValue();
  return val === null ? '' : val;
}

function parseBoolean(v) {
  if (v === true || v === false) return v;
  if (v === 0 || v === '0') return false;
  if (!v) return false;
  const s = v.toString().trim().toLowerCase();
  return ['true','yes','y','1','oui','o','checked'].indexOf(s) !== -1;
}

function toDateOnly(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  const d = new Date(v);
  if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  throw new Error('Impossible de parser la date: ' + v);
}

function combineDateTime(dateVal, timeVal) {
  if (!dateVal) return null;
  let dateObj = (dateVal instanceof Date) ? new Date(dateVal) : new Date(dateVal);
  if (isNaN(dateObj)) return null;
  if (!timeVal) return dateObj;
  if (timeVal instanceof Date) {
    dateObj.setHours(timeVal.getHours(), timeVal.getMinutes(), timeVal.getSeconds(), 0);
    return dateObj;
  }
  const s = timeVal.toString().trim();
  const parts = s.match(/^(\d{1,2})[:h](\d{2})(?::(\d{2}))?$/);
  if (parts) {
    const hh = parseInt(parts[1], 10);
    const mm = parseInt(parts[2], 10);
    const ss = parts[3] ? parseInt(parts[3], 10) : 0;
    dateObj.setHours(hh, mm, ss, 0);
    return dateObj;
  }
  const t = new Date(timeVal);
  if (!isNaN(t)) {
    dateObj.setHours(t.getHours(), t.getMinutes(), t.getSeconds(), 0);
    return dateObj;
  }
  return null;
}

// Write Processing Log + Last Updated and color row.
// NOTE: Last Updated will contain the ERROR message when type === 'error' (per your request).
function writeProcessingResult(sheet, row, headerIndex, message, type) {
  if (headerIndex['processing log']) sheet.getRange(row, headerIndex['processing log']).setValue(message);

  // <<< Ligne modifiée pour afficher l'erreur dans Last Updated (si type === 'error') >>>
  if (headerIndex['last updated']) {
    if (type === 'error') {
      // write error message directly in Last Updated
      sheet.getRange(row, headerIndex['last updated']).setValue(message);
    } else {
      // write timestamp for success/warning
      const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      sheet.getRange(row, headerIndex['last updated']).setValue(now);
    }
  }
  // <<< fin de la modification >>>

  const lastCol = sheet.getLastColumn();
  let color = '';
  if (type === 'success') color = '#d9ead3';
  else if (type === 'warning') color = '#fff2cc';
  else if (type === 'error') color = '#f4cccc';
  if (color) {
    try { sheet.getRange(row, 1, 1, lastCol).setBackground(color); } catch (e) { Logger.log('Impossible de colorer la ligne: ' + e); }
  }
}
