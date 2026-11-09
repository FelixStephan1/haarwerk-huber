/**
 * Haarwerk Huber – Buchungs-Backend (Google Apps Script)
 *
 * Aktionen (Parameter "action"):
 *   POST book   – neue Buchung: prüfen, Kalender, Sheet, Bestätigung, Info an Inhaber
 *   GET  get    – Buchung zu einem Token anzeigen (für termin.html)
 *   POST cancel – Buchung stornieren (bis 24 h vorher)
 *   POST move   – Buchung verschieben (bis 24 h vorher)
 *
 * Kalender-Sync (sync): Der Salon arbeitet nur im Google Kalender. Wird ein
 * Termin dort verschoben oder gelöscht, gleicht sync() das Sheet an und
 * informiert den Kunden. Läuft automatisch SYNC_DELAY_MINUTES nach der letzten
 * Kalender-Änderung (Wartezeit gegen Verklicken) und zusätzlich jede Nacht
 * (Trigger werden von setup() angelegt).
 * Termine, die der Salon selbst einträgt (Pause, Urlaub, Telefonbuchung),
 * stehen nicht im Sheet und werden ignoriert – blockieren aber Online-Buchungen.
 *
 * Jede Buchung bekommt einen geheimen Token. Nur wer den Link aus der
 * Bestätigungs-E-Mail hat, kann die Buchung ändern.
 *
 * Bei einem unerwarteten Fehler bekommt der Inhaber eine Fehler-E-Mail.
 *
 * Die Kopie in Google (Erweiterungen → Apps Script im Sheet) muss mit
 * dieser Datei übereinstimmen. Diese Datei im Repo ist die Quelle.
 * Nach dem Einfügen: Bereitstellen → Bereitstellungen verwalten → ✏️ → Neue Version.
 */

const SHEET_ID = '1fwVo6hySzABl_tT26xx0Z-fv7NR4JD-N_oASHN6JTg8';
const CALENDAR_NAME = 'Haarwerk Huber';
const TIMEZONE = 'Europe/Vienna';
const SALON = 'Haarwerk Huber';
const ADDRESS = 'Musterstraße 12, 8010 Graz';
const PHONE = '+43 316 000 000';
const SITE_URL = 'https://haarwerk-huber.vercel.app';
const CHANGE_DEADLINE_HOURS = 24; // ändern/stornieren nur bis so viele Stunden vorher
const WEEKDAYS = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'];

// Leistung → Dauer in Minuten (muss zu shared.js auf der Website passen)
const SERVICES = {
  'Damen – Waschen, Schneiden, Föhnen': 60,
  'Herren – Waschen, Schneiden, Styling': 30,
  'Kinder bis 12 Jahre': 30,
  'Färben (Ansatz)': 90,
  'Strähnen': 120,
  'Bart trimmen': 15,
  'Hochsteckfrisur': 60,
};

// Öffnungszeiten: Wochentag (0 = Sonntag) → [Öffnung, Schluss] in Minuten ab Mitternacht
const OPENING_HOURS = {
  2: [9 * 60, 18 * 60], // Dienstag
  3: [9 * 60, 18 * 60], // Mittwoch
  4: [9 * 60, 18 * 60], // Donnerstag
  5: [9 * 60, 18 * 60], // Freitag
  6: [8 * 60, 13 * 60], // Samstag
};

// Spalten im Sheet (1 = A)
const COL = {
  received: 1, date: 2, time: 3, service: 4, duration: 5, name: 6,
  email: 7, phone: 8, message: 9, eventId: 10, token: 11, status: 12,
};
const STATUS = {
  booked: 'gebucht',
  moved: 'verschoben',
  cancelled: 'storniert',
  salonMoved: 'verschoben (Salon)',
  salonCancelled: 'storniert (Salon)',
};
const isCancelled = (status) => String(status).startsWith(STATUS.cancelled);

// Bei Änderungen durch den Salon zusätzlich eine Kopie an den Inhaber (so fällt ein versehentliches Verschieben sofort auf)
const OWNER_COPY_ON_SALON_CHANGES = true;

// Wartezeit nach einer Kalender-Änderung, bevor verglichen wird. Jede weitere Änderung startet die
// Wartezeit neu. So bekommt der Kunde bei Verklicken / mehrfachem Verschieben nur EINE E-Mail
// (oder gar keine, wenn der Termin wieder an seinem alten Platz ist).
const SYNC_DELAY_MINUTES = 2;


// ───────────────────────── Einstieg ─────────────────────────

function doGet(e) {
  return handle(e, (p) => {
    if (p.action !== 'get') fail('Unbekannte Aktion.');
    const b = findBooking(p.t);
    return reply(true, '', { booking: publicView(b) });
  });
}

function doPost(e) {
  return handle(e, (p) => {
    if (p.website) return reply(true, 'OK'); // Spam-Falle: Menschen lassen dieses Feld leer
    const action = p.action || 'book';
    if (action === 'book') return book(p);
    if (action === 'cancel') return cancel(p);
    if (action === 'move') return move(p);
    fail('Unbekannte Aktion.');
  });
}

/** Gemeinsame Fehlerbehandlung für GET und POST */
function handle(e, fn) {
  const p = (e && e.parameter) || {};
  try {
    return fn(p);
  } catch (err) {
    if (err.userMessage) return reply(false, err.userMessage);
    MailApp.sendEmail(ownerEmail(), `[${SALON}] FEHLER im Buchungssystem`,
      `Fehler: ${err}\n${err.stack || ''}\n\nDaten:\n${JSON.stringify(p, null, 2)}`);
    return reply(false, `Da ist leider etwas schiefgegangen. Bitte rufen Sie uns an: ${PHONE}`);
  }
}


// ───────────────────────── Aktionen ─────────────────────────

function book(p) {
  const b = validateCustomer(p);
  Object.assign(b, checkSlot(b.date, b.time, b.duration));
  b.token = Utilities.getUuid();
  b.status = STATUS.booked;

  withLock(() => {
    const calendar = getCalendar();
    if (calendar.getEvents(b.start, b.end).length > 0) fail(SLOT_TAKEN);
    const event = calendar.createEvent(eventTitle(b), b.start, b.end, { description: eventDescription(b) });
    b.eventId = event.getId();

    sheet().appendRow([
      new Date(),
      "'" + b.date,   // ' = als Text speichern, damit Sheets nichts umwandelt
      "'" + b.time,
      b.service,
      b.duration,
      b.name,
      b.email,
      "'" + b.phone,
      b.message,
      b.eventId,
      b.token,
      b.status,
    ]);
  });

  sendConfirmation(b, 'new');
  notifyOwner(b, 'Neue Buchung');
  return reply(true, 'Danke! Ihr Termin ist gebucht. Sie erhalten gleich eine Bestätigung per E-Mail.');
}

function cancel(p) {
  let b;
  withLock(() => {
    b = findBooking(p.t);
    assertChangeable(b);

    const event = getCalendar().getEventById(b.eventId);
    if (event) event.deleteEvent();

    sheet().getRange(b.row, COL.status).setValue(STATUS.cancelled);
  });

  sendCancellation(b);
  notifyOwner(b, 'STORNIERT');
  return reply(true, 'Ihr Termin wurde storniert. Sie erhalten eine Bestätigung per E-Mail.');
}

function move(p) {
  let b, oldWhen;
  withLock(() => {
    b = findBooking(p.t);
    assertChangeable(b);
    oldWhen = formatWhen(b.start);

    const date = String(p.date || '').trim();
    const time = String(p.time || '').trim();
    const slot = checkSlot(date, time, b.duration);

    // frei? (der eigene Termin zählt nicht als Konflikt)
    const calendar = getCalendar();
    const conflicts = calendar.getEvents(slot.start, slot.end).filter((ev) => ev.getId() !== b.eventId);
    if (conflicts.length > 0) fail(SLOT_TAKEN);

    let event = calendar.getEventById(b.eventId);
    if (event) {
      event.setTime(slot.start, slot.end);
    } else { // Termin wurde im Kalender gelöscht → neu anlegen
      event = calendar.createEvent(eventTitle(b), slot.start, slot.end, { description: eventDescription(b) });
      sheet().getRange(b.row, COL.eventId).setValue(event.getId());
    }

    Object.assign(b, { date, time }, slot, { status: STATUS.moved });
    sheet().getRange(b.row, COL.date, 1, 2).setValues([["'" + date, "'" + time]]);
    sheet().getRange(b.row, COL.status).setValue(STATUS.moved);
  });

  sendConfirmation(b, 'moved');
  notifyOwner(b, `VERSCHOBEN (vorher ${oldWhen})`);
  return reply(true, 'Ihr Termin wurde verschoben. Sie erhalten eine neue Bestätigung per E-Mail.');
}


// ───────────────────────── Kalender-Sync ─────────────────────────

/**
 * Wird bei JEDER Kalender-Änderung aufgerufen (Trigger). Synchronisiert nicht sofort,
 * sondern stellt einen "Wecker" auf SYNC_DELAY_MINUTES. Ein schon gestellter Wecker
 * wird gelöscht, die Wartezeit startet also neu.
 */
function onCalendarChange() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'syncDelayed')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('syncDelayed').timeBased().after(SYNC_DELAY_MINUTES * 60000).create();
}

/** Der "Wecker": räumt seinen eigenen Trigger weg und synchronisiert. */
function syncDelayed(e) {
  ScriptApp.getProjectTriggers()
    .filter((t) => e && t.getUniqueId() === e.triggerUid)
    .forEach((t) => ScriptApp.deleteTrigger(t));
  sync();
}

/**
 * Vergleicht alle kommenden Buchungen im Sheet mit dem Kalender.
 * Läuft per Trigger (Kalender geändert + jede Nacht). Sendet nur bei echten Unterschieden E-Mails,
 * kann also beliebig oft laufen. Nutzt dieselbe Sperre wie Buchen/Verschieben/Stornieren,
 * sieht also nie einen halb fertigen Zustand.
 */
function sync() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return; // gerade beschäftigt → der nächste Lauf holt es nach
  const changes = [];
  try {
    const sh = sheet();
    const last = sh.getLastRow();
    if (last < 2) return;
    const calendar = getCalendar();
    const now = Date.now();
    const rows = sh.getRange(2, 1, last - 1, COL.status).getDisplayValues();

    rows.forEach((r, i) => {
      const b = rowToBooking(r, i + 2);
      if (!b.token || isCancelled(b.status)) return; // alte Zeilen ohne Token / schon storniert
      if (b.start.getTime() <= now) return;          // Vergangenes nie anfassen

      const event = findEvent(calendar, b.eventId);
      if (!event) {
        sh.getRange(b.row, COL.status).setValue(STATUS.salonCancelled);
        changes.push({ b, kind: 'cancelled' });
        return;
      }

      const start = event.getStartTime();
      const end = event.getEndTime();
      if (start.getTime() === b.start.getTime() && end.getTime() === b.end.getTime()) return; // alles gleich

      const old = formatWhen(b.start);
      const oldEnd = Utilities.formatDate(b.end, TIMEZONE, 'HH:mm');
      const oldDuration = b.duration;
      const startChanged = start.getTime() !== b.start.getTime();
      b.date = Utilities.formatDate(start, TIMEZONE, 'yyyy-MM-dd');
      b.time = Utilities.formatDate(start, TIMEZONE, 'HH:mm');
      b.duration = Math.round((end - start) / 60000);
      b.start = start;
      b.end = end;
      b.status = STATUS.salonMoved;
      sh.getRange(b.row, COL.date, 1, 4).setValues([["'" + b.date, "'" + b.time, b.service, b.duration]]);
      sh.getRange(b.row, COL.status).setValue(b.status);
      const kind = start.getTime() <= now ? 'movedToPast' : (startChanged ? 'moved' : 'resized');
      changes.push({ b, kind, old, oldEnd, oldDuration });
    });
  } finally {
    lock.releaseLock();
  }

  // E-Mails erst nach dem Freigeben der Sperre (das Sheet ist zu diesem Zeitpunkt schon aktuell)
  changes.forEach(({ b, kind, old, oldEnd, oldDuration }) => {
    if (kind === 'cancelled') {
      sendCancellation(b, true);
      if (OWNER_COPY_ON_SALON_CHANGES) notifyOwner(b, 'Vom Salon ABGESAGT – Kunde wurde informiert');
    } else if (kind === 'moved') {
      sendConfirmation(b, 'salonMoved');
      if (OWNER_COPY_ON_SALON_CHANGES) notifyOwner(b, `Vom Salon VERSCHOBEN (vorher ${old}) – Kunde wurde informiert`);
    } else if (kind === 'resized') {
      sendConfirmation(b, 'salonResized', { oldEnd });
      if (OWNER_COPY_ON_SALON_CHANGES) notifyOwner(b, `Vom Salon ANGEPASST (Dauer ${oldDuration} → ${b.duration} Min) – Kunde wurde informiert`);
    } else { // in die Vergangenheit verschoben: Kunde nicht anschreiben, nur Inhaber
      notifyOwner(b, `In die Vergangenheit verschoben (vorher ${old}) – Kunde wurde NICHT informiert`);
    }
  });
}

/**
 * Termin im Salon-Kalender suchen. Prüft zusätzlich über getEvents(), ob er wirklich
 * noch im Kalender steht (gelöschte Termine kann getEventById manchmal noch liefern).
 */
function findEvent(calendar, eventId) {
  if (!eventId) return null;
  const event = calendar.getEventById(eventId);
  if (!event) return null;
  const stillThere = calendar.getEvents(event.getStartTime(), event.getEndTime())
    .some((ev) => ev.getId() === eventId);
  return stillThere ? event : null;
}


// ───────────────────────── Prüfungen ─────────────────────────

const SLOT_TAKEN = 'Zu dieser Zeit ist leider kein Termin frei. Bitte wählen Sie eine andere Uhrzeit.';

/** Fehler, deren Text der Kunde sehen darf */
function fail(msg) {
  const err = new Error(msg);
  err.userMessage = msg;
  throw err;
}

function validateCustomer(p) {
  const clean = (v, max) => String(v || '').trim().slice(0, max);
  const b = {
    service: clean(p.service, 100),
    date: clean(p.date, 10),
    time: clean(p.time, 5),
    name: clean(p.name, 100),
    email: clean(p.email, 150),
    phone: clean(p.phone, 40),
    message: clean(p.message, 1000),
  };
  if (!(b.service in SERVICES)) fail('Bitte wählen Sie eine Leistung.');
  if (!b.name) fail('Bitte geben Sie Ihren Namen an.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(b.email)) fail('Bitte geben Sie eine gültige E-Mail-Adresse an.');
  if (!isValidPhone(b.phone)) fail('Bitte geben Sie eine gültige Telefonnummer an (z. B. +43 660 1234567).');
  b.duration = SERVICES[b.service];
  return b;
}

/** Prüft Datum/Uhrzeit gegen Öffnungszeiten und Vorlauf. Gibt {start, end} zurück. */
function checkSlot(date, time, duration) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || !/^\d{2}:\d{2}$/.test(time)) fail('Bitte wählen Sie Datum und Uhrzeit.');

  const start = Utilities.parseDate(`${date} ${time}`, TIMEZONE, 'yyyy-MM-dd HH:mm');
  const end = new Date(start.getTime() + duration * 60000);
  if (start.getTime() < Date.now() + 60 * 60000) fail('Bitte buchen Sie mindestens eine Stunde im Voraus.');

  const weekday = Number(Utilities.formatDate(start, TIMEZONE, 'u')) % 7; // u: 1 = Mo … 7 = So
  const hours = OPENING_HOURS[weekday];
  if (!hours) fail('An diesem Tag haben wir geschlossen (Montag und Sonntag).');
  const [h, m] = time.split(':').map(Number);
  const startMin = h * 60 + m;
  if (startMin < hours[0] || startMin + duration > hours[1]) {
    fail('Dieser Termin liegt (mit der Dauer der Leistung) außerhalb unserer Öffnungszeiten.');
  }
  return { start, end };
}

/**
 * Telefonnummer prüfen: nur Ziffern, Leerzeichen und + / - ( ),
 * ein + nur ganz vorne, 6 bis 15 Ziffern (15 = internationales Maximum).
 * Muss mit isValidPhone() in shared.js übereinstimmen.
 */
function isValidPhone(phone) {
  if (!/^\+?[0-9 ()\/-]+$/.test(phone)) return false;
  const digits = phone.replace(/\D/g, '').length;
  return digits >= 6 && digits <= 15;
}

function assertChangeable(b) {
  if (isCancelled(b.status)) fail('Dieser Termin wurde bereits storniert.');
  if (b.start.getTime() < Date.now()) fail('Dieser Termin liegt in der Vergangenheit.');
  if (b.start.getTime() - Date.now() < CHANGE_DEADLINE_HOURS * 3600000) {
    fail(`Änderungen sind nur bis ${CHANGE_DEADLINE_HOURS} Stunden vor dem Termin online möglich. Bitte rufen Sie uns an: ${PHONE}`);
  }
}


// ───────────────────────── Sheet & Kalender ─────────────────────────

function sheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

/** Buchung zum Token suchen. Wirft einen Kundenfehler, wenn es keine gibt. */
function findBooking(token) {
  token = String(token || '').trim();
  if (!/^[0-9a-f-]{36}$/i.test(token)) fail('Dieser Link ist ungültig.');

  const sh = sheet();
  const last = sh.getLastRow();
  if (last < 2) fail('Diese Buchung wurde nicht gefunden.');
  // getDisplayValues: alles als Text, so wie es im Sheet steht (keine Zeitzonen-Überraschungen)
  const rows = sh.getRange(2, 1, last - 1, COL.status).getDisplayValues();
  const i = rows.findIndex((r) => r[COL.token - 1] === token);
  if (i === -1) fail('Diese Buchung wurde nicht gefunden.');

  return rowToBooking(rows[i], i + 2);
}

/** Eine Sheet-Zeile (Anzeigewerte) in ein Buchungs-Objekt umwandeln */
function rowToBooking(r, row) {
  const b = {
    row,
    date: r[COL.date - 1],
    time: r[COL.time - 1],
    service: r[COL.service - 1],
    duration: Number(r[COL.duration - 1]),
    name: r[COL.name - 1],
    email: r[COL.email - 1],
    phone: r[COL.phone - 1],
    message: r[COL.message - 1],
    eventId: r[COL.eventId - 1],
    token: r[COL.token - 1],
    status: r[COL.status - 1],
  };
  b.start = Utilities.parseDate(`${b.date} ${b.time}`, TIMEZONE, 'yyyy-MM-dd HH:mm');
  b.end = new Date(b.start.getTime() + b.duration * 60000);
  return b;
}

/** Was termin.html sehen darf (keine E-Mail, kein Telefon) */
function publicView(b) {
  let canChange = true, reason = '';
  try { assertChangeable(b); } catch (err) { canChange = false; reason = err.userMessage || ''; }
  return {
    name: b.name,
    service: b.service,
    duration: b.duration,
    date: b.date,
    time: b.time,
    dateLong: formatDateLong(b.start),
    timeRange: formatTimeRange(b),
    status: b.status,
    canChange,
    reason,
  };
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // gleichzeitige Anfragen nacheinander abarbeiten
  try { return fn(); } finally { lock.releaseLock(); }
}

function getCalendar() {
  return CalendarApp.getCalendarsByName(CALENDAR_NAME)[0]
    || CalendarApp.createCalendar(CALENDAR_NAME, { timeZone: TIMEZONE });
}

function eventTitle(b) {
  return `${b.service} – ${b.name}`;
}

function eventDescription(b) {
  return `Telefon: ${b.phone}\nE-Mail: ${b.email}\nNachricht: ${b.message || '–'}`;
}


// ───────────────────────── E-Mails ─────────────────────────

function manageUrl(b) {
  return `${SITE_URL}/termin.html?t=${b.token}`;
}

/**
 * kind: 'new' = neue Buchung, 'moved' = vom Kunden verschoben, 'salonMoved' = vom Salon verschoben,
 *       'salonResized' = vom Salon verlängert/verkürzt (Beginn gleich; extra.oldEnd = alte Endzeit)
 */
function sendConfirmation(b, kind, extra) {
  const startTime = Utilities.formatDate(b.start, TIMEZONE, 'HH:mm');
  const endTime = Utilities.formatDate(b.end, TIMEZONE, 'HH:mm');
  const dateLong = formatDateLong(b.start);
  const timeRange = formatTimeRange(b);
  const intro = {
    new: 'vielen Dank für Ihre Buchung – wir freuen uns auf Sie!',
    moved: 'Ihr Termin wurde verschoben. Hier sind die neuen Termindaten:',
    salonMoved: 'wir mussten Ihren Termin leider verschieben – bitte entschuldigen Sie. ' +
      'Hier sind die neuen Termindaten. Passt Ihnen die neue Zeit nicht, können Sie sie unten ändern oder uns anrufen.',
    salonResized: `Ihr Termin beginnt wie geplant um ${startTime} Uhr, dauert jetzt aber bis ${endTime} Uhr ` +
      `(vorher bis ${extra && extra.oldEnd} Uhr). Hier sind die aktuellen Termindaten:`,
  }[kind];

  const plain =
`Hallo ${b.name},

${intro}

Leistung: ${b.service}
Termin:   ${dateLong}, ${timeRange}
Adresse:  ${ADDRESS}

Termin ändern oder stornieren (bis ${CHANGE_DEADLINE_HOURS} Stunden vorher):
${manageUrl(b)}

Danach erreichen Sie uns telefonisch unter ${PHONE}.

Ihr Team von ${SALON}

(Hinweis: Dies ist ein Übungsprojekt – der Salon ist erfunden.)`;

  const details =
    detailRow('Leistung', esc(b.service)) +
    detailRow('Datum', dateLong) +
    detailRow('Uhrzeit', `${timeRange} <span style="color:#6f655c;">(ca. ${b.duration} Min.)</span>`) +
    detailRow('Adresse', ADDRESS) +
    (b.message ? detailRow('Ihre Nachricht', esc(b.message)) : '');

  const body = `
      <p style="margin:0 0 16px;">Hallo ${esc(b.name)},</p>
      <p style="margin:0 0 20px;">${intro}</p>
      ${detailBox(details)}
      <p style="margin:28px 0 0;text-align:center;">
        <a href="${manageUrl(b)}" style="display:inline-block;background:#8a5a3c;color:#ffffff;text-decoration:none;font-family:Arial,sans-serif;font-size:15px;padding:12px 24px;border-radius:2px;">Termin ändern oder stornieren</a>
      </p>
      <p style="margin:12px 0 0;text-align:center;font-family:Arial,sans-serif;font-size:13px;color:#6f655c;">
        Online möglich bis ${CHANGE_DEADLINE_HOURS} Stunden vor dem Termin, danach bitte anrufen:
        <a href="tel:${PHONE.replace(/\s/g, '')}" style="color:#8a5a3c;">${PHONE}</a>
      </p>
      <p style="margin:24px 0 0;">Bis bald!<br>Ihr Team von ${SALON}</p>`;

  const subjectStart = {
    new: 'Terminbestätigung',
    moved: 'Termin verschoben',
    salonMoved: 'Termin verschoben',
    salonResized: 'Termin angepasst',
  }[kind];
  MailApp.sendEmail({
    to: b.email,
    subject: `${subjectStart}: ${dateLong}, ${startTime}${kind === 'salonResized' ? '–' + endTime : ''} Uhr – ${SALON}`,
    body: plain,
    htmlBody: emailLayout(body),
    name: SALON,
  });
}

/** bySalon: true = der Salon hat den Termin im Kalender gelöscht */
function sendCancellation(b, bySalon) {
  const intro = bySalon
    ? 'wir mussten Ihren Termin leider absagen – bitte entschuldigen Sie die Unannehmlichkeiten:'
    : 'Ihr Termin wurde storniert:';
  const dateLong = formatDateLong(b.start);
  const timeRange = formatTimeRange(b);
  const plain =
`Hallo ${b.name},

${intro}

Leistung: ${b.service}
Termin:   ${dateLong}, ${timeRange}

Sie möchten einen neuen Termin? Gerne: ${SITE_URL}/#termin

Ihr Team von ${SALON}

(Hinweis: Dies ist ein Übungsprojekt – der Salon ist erfunden.)`;

  const body = `
      <p style="margin:0 0 16px;">Hallo ${esc(b.name)},</p>
      <p style="margin:0 0 20px;">${intro}</p>
      ${detailBox(detailRow('Leistung', esc(b.service)) + detailRow('Datum', dateLong) + detailRow('Uhrzeit', timeRange))}
      <p style="margin:28px 0 0;text-align:center;">
        <a href="${SITE_URL}/#termin" style="display:inline-block;background:#8a5a3c;color:#ffffff;text-decoration:none;font-family:Arial,sans-serif;font-size:15px;padding:12px 24px;border-radius:2px;">Neuen Termin buchen</a>
      </p>
      <p style="margin:24px 0 0;">Ihr Team von ${SALON}</p>`;

  MailApp.sendEmail({
    to: b.email,
    subject: `${bySalon ? 'Termin abgesagt' : 'Termin storniert'}: ${dateLong}, ${Utilities.formatDate(b.start, TIMEZONE, 'HH:mm')} Uhr – ${SALON}`,
    body: plain,
    htmlBody: emailLayout(body),
    name: SALON,
  });
}

function notifyOwner(b, what) {
  const when = formatWhen(b.start);
  MailApp.sendEmail(ownerEmail(), `[${SALON}] ${what}: ${when} – ${b.name}`,
    `${what}\n\n${b.service} (${b.duration} Min)\n${when}\n\n${b.name}\n${b.email}\n${b.phone}\n\nNachricht: ${b.message || '–'}`);
}

function emailLayout(inner) {
  return `
<div style="background:#faf7f2;padding:32px 16px;font-family:Georgia,'Times New Roman',serif;">
  <table role="presentation" cellpadding="0" cellspacing="0" style="max-width:520px;width:100%;margin:0 auto;background:#ffffff;border:1px solid #e4dcd0;">
    <tr><td style="padding:28px 32px;text-align:center;border-bottom:1px solid #e4dcd0;">
      <div style="font-size:28px;color:#2b2622;">${SALON}</div>
      <div style="font-size:14px;color:#6f655c;margin-top:4px;">Ihr Friseur in Graz</div>
    </td></tr>
    <tr><td style="padding:28px 32px;color:#2b2622;font-size:16px;line-height:1.6;">${inner}
    </td></tr>
    <tr><td style="padding:16px 32px;text-align:center;font-family:Arial,sans-serif;font-size:12px;color:#6f655c;border-top:1px solid #e4dcd0;">
      ${SALON} · ${ADDRESS} · ${PHONE}<br>
      Übungsprojekt – dieser Friseursalon ist erfunden.
    </td></tr>
  </table>
</div>`;
}

function detailBox(rows) {
  return `<div style="background:#faf7f2;border-left:3px solid #8a5a3c;padding:16px 20px;">
        <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;font-size:15px;font-family:Arial,sans-serif;">${rows}</table>
      </div>`;
}

function detailRow(label, value) {
  return `<tr><td style="padding:8px 0;color:#6f655c;width:110px;vertical-align:top;">${label}</td>` +
    `<td style="padding:8px 0;color:#2b2622;">${value}</td></tr>`;
}


// ───────────────────────── Hilfsfunktionen ─────────────────────────

function formatDateLong(date) {
  const weekday = WEEKDAYS[Number(Utilities.formatDate(date, TIMEZONE, 'u')) - 1];
  return `${weekday}, ${Utilities.formatDate(date, TIMEZONE, 'dd.MM.yyyy')}`;
}

function formatTimeRange(b) {
  return `${Utilities.formatDate(b.start, TIMEZONE, 'HH:mm')} – ${Utilities.formatDate(b.end, TIMEZONE, 'HH:mm')} Uhr`;
}

function formatWhen(date) {
  return Utilities.formatDate(date, TIMEZONE, 'EEE dd.MM.yyyy HH:mm');
}

/** Text für HTML unschädlich machen (Kundeneingaben!) */
function esc(text) {
  return String(text).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}

function ownerEmail() {
  return Session.getEffectiveUser().getEmail();
}

function reply(ok, message, extra) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({ ok, message }, extra || {})))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Einmal im Editor ausführen (und nach jeder Änderung an den Triggern):
 * fragt die Berechtigungen ab, legt Kalender und Spalten an und richtet die Sync-Trigger ein.
 */
function setup() {
  const calendar = getCalendar();
  const sh = sheet();
  sh.setFrozenRows(1);
  sh.getRange(1, COL.token, 1, 2).setValues([['Token', 'Status']]);

  // alte Trigger entfernen, damit es keine doppelten gibt
  ScriptApp.getProjectTriggers()
    .filter((t) => ['sync', 'syncDelayed', 'onCalendarChange'].includes(t.getHandlerFunction()))
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onCalendarChange').forUserCalendar(calendar.getId()).onEventUpdated().create();
  ScriptApp.newTrigger('sync').timeBased().everyDays(1).atHour(3).inTimezone(TIMEZONE).create();

  Logger.log('Fertig. Kalender, Sheet und Sync-Trigger sind bereit. Besitzer: ' + ownerEmail());
}
