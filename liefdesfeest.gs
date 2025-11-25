/***** Liefdesfeest – Sheets ↔ Calendar Sync (primary calendar, no duplicates) *****/

const TIMEZONE = 'Europe/Amsterdam';
const SHEET_NAME = 'Overzicht'; // <- jouw tabnaam

// Exacte kopregels zoals in je sheet:
const HEADER = {
  categorie: 'Categorie',
  omschrijving: 'Omschrijving / Context',
  deadline: 'Deadline',
  toegewezen: 'Toegewezen aan',
  status: 'Status',
  notities: 'Notities',
  eventId: 'CalEventId',
  reminderDays: 'ReminderOffsetDays' // optioneel per rij
};

// Voeg dit bovenin toe:
const GUEST_EMAILS = [
  'jacmar@live.nl',   // vul hier het extra e-mailadres in
  // 'nogiemand@example.com'  // eventueel meer
];

// Primaire agenda (laat leeg). Vul anders een specifieke ID in.
const CALENDAR_ID = '';

// Bij "datum+tijd" maken we een timed event met deze duur (minuten)
const DEFAULT_TIMED_EVENT_DURATION_MIN = 30;

// Status die een event moeten VERWIJDEREN
const DELETE_KEYWORDS = ['afgerond', 'geannuleerd'];

// Standaard reminders (dagen vooraf). We doen er 2 (30d en 14d) + dag zelf (0 min).
const DEFAULT_REMINDERS_DAYS = [30, 14];

// Visuele sync-feedback in kolom I
const SYNC_STATUS_COLUMN = 'I';

/* ========== Menu ========== */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Liefdesfeest')
    .addItem('Sync alle taken', 'syncAll')
    .addItem('Push geselecteerde rij', 'syncActiveRow')
    .addSeparator()
    .addItem('Verwijder gekoppelde event (rij)', 'deleteActiveRowEvent')
    .addSeparator()
    .addItem('Herstel autosync (on-edit trigger)', 'ensureAutoSyncTrigger')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('5-min autosync')
        .addItem('Inschakelen', 'enableFiveMinSync')
        .addItem('Uitschakelen', 'disableFiveMinSync')
    )
    .addToUi();
}

/* ========== INSTALLABLE onEdit (maak trigger in UI of via menu) ========== */
function onEditInstalled(e) {
  try {
    if (!e || !e.range || !e.source) return;
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    // Check: zit er een relevante kolom in het bewerkte bereik?
    const relevantHeaders = {
      [HEADER.deadline]: true,
      [HEADER.status]: true,
      [HEADER.omschrijving]: true,
      [HEADER.categorie]: true,
      [HEADER.reminderDays]: true
    };
    const firstCol = e.range.getColumn();
    const lastCol  = firstCol + e.range.getNumColumns() - 1;

    let hasRelevant = false;
    for (let c = firstCol; c <= lastCol; c++) {
      const h = String(sheet.getRange(1, c).getValue()).trim();
      if (relevantHeaders[h]) { hasRelevant = true; break; }
    }
    if (!hasRelevant) return;

    const headers = getHeaderMap(sheet);
    if (!headers) return;

    // Sync elke data-rij in het bewerkte bereik
    const firstRow = Math.max(2, e.range.getRow()); // vanaf rij 2
    const lastRow  = e.range.getRow() + e.range.getNumRows() - 1;

    for (let r = firstRow; r <= lastRow; r++) {
      const rowVals = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
      const isEmpty = rowVals.every(v => v === '' || v === null);
      if (isEmpty) continue;

      const result = syncRow(sheet, r, headers); // 'synced' | 'deleted' | 'skipped'
      if (result === 'synced') {
        setSyncStatus(sheet, r, '✅ Gesynchroniseerd op ' +
          Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
      } else if (result === 'deleted') {
        setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
          Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
      }
    }
  } catch (err) {
    Logger.log(err);
    try { SpreadsheetApp.getActive().toast('Fout bij onEditInstalled: ' + err); } catch (_) {}
  }
}
/* ========== Handmatige acties ========== */
function syncAll() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet "${SHEET_NAME}" niet gevonden.`);
  const headers = getHeaderMap(sheet);
  if (!headers) throw new Error('Kolomkoppen niet gevonden of onvolledig.');

  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    const rowVals = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (rowVals.every(v => v === '' || v === null)) continue;
    const result = syncRow(sheet, r, headers);
    if (result === 'synced') {
      setSyncStatus(sheet, r, '✅ Gesynchroniseerd op ' +
        Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
    } else if (result === 'deleted') {
      setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
        Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
    }
  }
}

function syncActiveRow() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const headers = getHeaderMap(sheet);
  if (!headers) {
    SpreadsheetApp.getUi().alert('Kolomkoppen niet gevonden of onvolledig.');
    return;
  }
  const r = sheet.getActiveRange().getRow();
  if (r <= 1) {
    SpreadsheetApp.getUi().alert('Selecteer een data-rij (vanaf rij 2).');
    return;
  }
  const result = syncRow(sheet, r, headers);
  if (result === 'synced') {
    setSyncStatus(sheet, r, '✅ Gesynchroniseerd op ' +
      Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
  } else if (result === 'deleted') {
    setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
      Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
  }
}

function deleteActiveRowEvent() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const headers = getHeaderMap(sheet);
  if (!headers) {
    SpreadsheetApp.getUi().alert('Kolomkoppen niet gevonden of onvolledig.');
    return;
  }
  const r = sheet.getActiveRange().getRow();
  if (r <= 1) {
    SpreadsheetApp.getUi().alert('Selecteer een data-rij (vanaf rij 2).');
    return;
  }
  const eventId = getCell(sheet, r, headers[HEADER.eventId]);
  if (eventId) {
    const cal = getCalendar();
    try {
      const ev = cal.getEventById(normalizeEventId(eventId));
      if (ev) ev.deleteEvent();
    } catch (_) {}
    setCell(sheet, r, headers[HEADER.eventId], '');
  }
  setSyncStatus(sheet, r, '❌ Verwijderd uit agenda op ' +
    Utilities.formatDate(new Date(), TIMEZONE, 'dd-MM-yyyy HH:mm'));
  SpreadsheetApp.getUi().alert('Event verwijderd (indien aanwezig) en koppeling gewist.');
}

/* ========== Triggers helpers (menu) ========== */
function ensureAutoSyncTrigger() {
  const f = 'onEditInstalled';
  const triggers = ScriptApp.getProjectTriggers();
  const has = triggers.some(t =>
    t.getHandlerFunction() === f &&
    t.getEventType() === ScriptApp.EventType.ON_EDIT
  );
  if (!has) {
    ScriptApp.newTrigger(f)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
  SpreadsheetApp.getActive().toast('Autosync (on-edit) staat AAN.');
}

function enableFiveMinSync() {
  const f = 'syncAll';
  const has = ScriptApp.getProjectTriggers().some(t =>
    t.getHandlerFunction() === f &&
    t.getEventType() === ScriptApp.EventType.TIME_DRIVEN
  );
  if (!has) {
    ScriptApp.newTrigger(f)
      .timeBased()
      .everyMinutes(5)
      .create();
  }
  SpreadsheetApp.getActive().toast('5-min autosync staat AAN.');
}

function disableFiveMinSync() {
  const f = 'syncAll';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === f &&
        t.getEventType() === ScriptApp.EventType.TIME_DRIVEN) {
      ScriptApp.deleteTrigger(t);
    }
  });
  SpreadsheetApp.getActive().toast('5-min autosync is UIT.');
}

/* ========== Kern-sync per rij ========== */
// return: 'synced' | 'deleted' | 'skipped'
function syncRow(sheet, row, headers) {
  const categorie = getCell(sheet, row, headers[HEADER.categorie]);
  const omschrijving = getCell(sheet, row, headers[HEADER.omschrijving]);
  const deadlineVal = getCell(sheet, row, headers[HEADER.deadline]);
  const toegewezen = getCell(sheet, row, headers[HEADER.toegewezen]);
  const status = getCell(sheet, row, headers[HEADER.status]);
  const notities = getCell(sheet, row, headers[HEADER.notities]);
  const existingEventId = getCell(sheet, row, headers[HEADER.eventId]);
  const reminderDaysOverride = getCell(sheet, row, headers[HEADER.reminderDays]); // optioneel

  const cal = getCalendar();

  // Verwijder-criteria
  const statusStr = (status || '').toString().toLowerCase();
  const shouldDelete =
    !deadlineVal ||
    DELETE_KEYWORDS.some(k => statusStr.includes(k));

  if (shouldDelete) {
    if (existingEventId) {
      try {
        const ev = cal.getEventById(normalizeEventId(existingEventId));
        if (ev) ev.deleteEvent();
      } catch (_) {}
      setCell(sheet, row, headers[HEADER.eventId], '');
    }
    return 'deleted';
  }

  // Titel en description
  const title = buildTitle(categorie, omschrijving);
  const description = buildDescription({
    categorie, omschrijving, toegewezen, status, notities
  });

  // Deadline parsen
  const { isAllDay, start, end } = parseDeadline(deadlineVal);

  // Reminders: 30d en 14d + dag zelf (0 min)
  const customDay = Number(reminderDaysOverride);
  const reminderDays = (isFinite(customDay) && customDay > 0)
    ? [customDay] // per rij override
    : DEFAULT_REMINDERS_DAYS.slice();

  let ev = existingEventId ? getEventSafe(cal, existingEventId) : null;

  if (!ev) {
    // Create
    if (isAllDay) {
      ev = cal.createAllDayEvent(title, start, { description });
    } else {
      ev = cal.createEvent(title, start, end, { description });
    }
    setReminders(ev, reminderDays, /*includeSameDay*/ true);
    ensureGuests(ev);
    setCell(sheet, row, headers[HEADER.eventId], ev.getId());
  } else {
    // Update
    try {
      if (isAllDay) {
        ev.setAllDayDate(start);
      } else {
        ev.setTime(start, end);
      }
      ev.setTitle(title);
      ev.setDescription(description);
      setReminders(ev, reminderDays, /*includeSameDay*/ true);
      ensureGuests(ev);
    } catch (e) {
      // Herstelpad: maak opnieuw aan als update faalt
      try { ev.deleteEvent(); } catch (_) {}
      let newEv;
      if (isAllDay) {
        newEv = cal.createAllDayEvent(title, start, { description });
      } else {
        newEv = cal.createEvent(title, start, end, { description });
      }
      setReminders(newEv, reminderDays, /*includeSameDay*/ true);
      ensureGuests(newEv);
      setCell(sheet, row, headers[HEADER.eventId], newEv.getId());
    }
  }
  return 'synced';
}

/* ========== Helpers ========== */
function getCalendar() {
  return CALENDAR_ID ? CalendarApp.getCalendarById(CALENDAR_ID)
                     : CalendarApp.getDefaultCalendar();
}

function getHeaderMap(sheet) {
  const lastCol = sheet.getLastColumn();
  const headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const map = {};
  Object.values(HEADER).forEach(name => {
    const idx = headerVals.findIndex(h => h.trim() === name);
    if (idx !== -1) map[name] = idx + 1;
  });
  // Vereist minimaal deze kolommen:
  const required = [HEADER.omschrijving, HEADER.deadline, HEADER.eventId];
  const ok = required.every(k => map[k]);
  return ok ? map : null;
}

function getCell(sheet, row, col) {
  if (!col) return '';
  return sheet.getRange(row, col).getValue();
}

function setCell(sheet, row, col, val) {
  if (!col) return;
  sheet.getRange(row, col).setValue(val);
}

function buildTitle(categorie, omschrijving) {
  const c = (categorie || '').toString().trim();
  const o = (omschrijving || '').toString().trim();
  if (c && o) return `${c} – ${o}`;
  return o || c || 'Liefdesfeest taak';
}

function buildDescription({ categorie, omschrijving, toegewezen, status, notities }) {
  const parts = [];
  if (omschrijving) parts.push(`Omschrijving: ${omschrijving}`);
  if (categorie) parts.push(`Categorie: ${categorie}`);
  if (toegewezen) parts.push(`Toegewezen aan: ${toegewezen}`);
  if (status) parts.push(`Status: ${status}`);
  if (notities) parts.push(`Notities: ${notities}`);
  parts.push(`Bron: ${SpreadsheetApp.getActiveSpreadsheet().getName()}`);
  parts.push(`Link: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`);
  return parts.join('\n');
}

// Datum of datum+tijd; all-day als er geen tijdcomponent is.
function parseDeadline(val) {
  const d = (val instanceof Date) ? new Date(val) : new Date(val);
  if (isNaN(d)) throw new Error('Ongeldige deadline-datum.');
  const hasTime = d.getHours() + d.getMinutes() + d.getSeconds() !== 0;

  if (!hasTime) {
    // All-day event: setAllDayDate gebruikt lokale datum (geen tijd)
    const start = new Date(d.getFullYear(), d.getMonth(), d.getDate());
    return { isAllDay: true, start, end: null };
  } else {
    const start = new Date(d);
    const end = new Date(start.getTime() + DEFAULT_TIMED_EVENT_DURATION_MIN * 60000);
    return { isAllDay: false, start, end };
  }
}

function normalizeEventId(eventId) {
  return eventId.includes('@') ? `${eventId}` : `${eventId}@google.com`;
}

function getEventSafe(cal, id) {
  try {
    return cal.getEventById(normalizeEventId(id));
  } catch (_) {
    return null;
  }
}

// Zet 30d + 14d + dag-zelf popup. Bij all-day toont Google de 0-min popup meestal in de ochtend.
function setReminders(ev, reminderDaysArray, includeSameDay) {
  try { ev.removeAllReminders(); } catch (_) {}
  (reminderDaysArray || []).forEach(d => {
    const mins = Math.max(0, Math.round(Number(d) * 24 * 60));
    if (isFinite(mins) && mins > 0) {
      try { ev.addPopupReminder(mins); } catch (_) {}
    }
  });
  if (includeSameDay) {
    try { ev.addPopupReminder(0); } catch (_) {}
  }
}

/* ========== Snelle testfunctie ========== */
function _testCreateEvent() {
  const cal = getCalendar();
  const start = new Date(new Date().getTime() + 60 * 60 * 1000);
  const end   = new Date(start.getTime() + 30 * 60 * 1000);
  const ev = cal.createEvent('Test – Liefdesfeest', start, end, { description: 'Test event' });
  ensureGuests(ev);
  Logger.log('Gemaakt: ' + ev.getId());
  SpreadsheetApp.getActive().toast('Testevent aangemaakt voor over 1 uur.');
}

/* ========== Visuele sync-feedback helper ========== */
function setSyncStatus(sheet, row, message) {
  try {
    const colIndex = sheet.getRange(`${SYNC_STATUS_COLUMN}1`).getColumn();
    sheet.getRange(row, colIndex).setValue(message);
  } catch (err) {
    Logger.log('Kon sync-status niet schrijven: ' + err);
  }
}

function ensureGuests(ev) {
  try {
    const existing = ev.getGuestList().map(g => g.getEmail().toLowerCase());
    GUEST_EMAILS.forEach(email => {
      if (!email) return;
      const lower = email.toLowerCase();
      if (!existing.includes(lower)) {
        ev.addGuest(email);
      }
    });
  } catch (e) {
    Logger.log('Kon gasten niet toevoegen: ' + e);
  }
}
