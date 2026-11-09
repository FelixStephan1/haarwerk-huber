/**
 * Haarwerk Huber – Buchungs-Backend (Google Apps Script)
 *
 * Empfängt Buchungen vom Formular auf der Website und:
 *   1. prüft die Angaben (Öffnungszeiten, Pflichtfelder, Spam)
 *   2. prüft, ob der Termin im Kalender noch frei ist
 *   3. schreibt eine Zeile ins Buchungs-Sheet
 *   4. legt einen Termin im Kalender "Haarwerk Huber" an
 *   5. schickt dem Kunden eine Bestätigung per E-Mail
 *   6. schickt dem Inhaber eine Benachrichtigung
 * Bei einem Fehler bekommt der Inhaber eine Fehler-E-Mail.
 *
 * Die Kopie in Google (Erweiterungen → Apps Script im Sheet) muss mit
 * dieser Datei übereinstimmen. Diese Datei im Repo ist die Quelle.
 */

const SHEET_ID = '1fwVo6hySzABl_tT26xx0Z-fv7NR4JD-N_oASHN6JTg8';
const CALENDAR_NAME = 'Haarwerk Huber';
const TIMEZONE = 'Europe/Vienna';
const SALON = 'Haarwerk Huber';

// Leistung → Dauer in Minuten (muss zu den Optionen im Formular passen)
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

function doPost(e) {
  const p = (e && e.parameter) || {};
  try {
    // Spam-Falle: echte Menschen sehen dieses Feld nicht und lassen es leer
    if (p.website) return reply(true, 'OK');

    const booking = validate(p);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000); // zwei gleichzeitige Buchungen nicht doppelt vergeben
    try {
      const calendar = getCalendar();
      if (calendar.getEvents(booking.start, booking.end).length > 0) {
        return reply(false, 'Dieser Termin ist leider schon vergeben. Bitte wählen Sie eine andere Uhrzeit.');
      }

      const event = calendar.createEvent(
        `${booking.service} – ${booking.name}`,
        booking.start,
        booking.end,
        { description: `Telefon: ${booking.phone}\nE-Mail: ${booking.email}\nNachricht: ${booking.message || '–'}` }
      );

      SpreadsheetApp.openById(SHEET_ID).getSheets()[0].appendRow([
        new Date(),
        booking.date,
        booking.time,
        booking.service,
        booking.duration,
        booking.name,
        booking.email,
        "'" + booking.phone, // ' verhindert, dass Sheets die Nummer als Zahl liest
        booking.message,
        event.getId(),
      ]);
    } finally {
      lock.releaseLock();
    }

    sendConfirmation(booking);
    notifyOwner(booking);
    return reply(true, 'Danke! Ihr Termin ist gebucht. Sie erhalten gleich eine Bestätigung per E-Mail.');

  } catch (err) {
    if (err.userMessage) return reply(false, err.userMessage);
    MailApp.sendEmail(ownerEmail(), `[${SALON}] FEHLER bei einer Buchung`,
      `Fehler: ${err}\n\nDaten:\n${JSON.stringify(p, null, 2)}`);
    return reply(false, 'Da ist leider etwas schiefgegangen. Bitte rufen Sie uns an.');
  }
}

function validate(p) {
  const fail = (msg) => { const err = new Error(msg); err.userMessage = msg; throw err; };
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
  if (!b.phone) fail('Bitte geben Sie Ihre Telefonnummer an.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(b.date) || !/^\d{2}:\d{2}$/.test(b.time)) fail('Bitte wählen Sie Datum und Uhrzeit.');

  b.duration = SERVICES[b.service];
  b.start = Utilities.parseDate(`${b.date} ${b.time}`, TIMEZONE, 'yyyy-MM-dd HH:mm');
  b.end = new Date(b.start.getTime() + b.duration * 60000);

  if (b.start.getTime() < Date.now() + 60 * 60000) fail('Bitte buchen Sie mindestens eine Stunde im Voraus.');

  const weekday = Number(Utilities.formatDate(b.start, TIMEZONE, 'u')) % 7; // u: 1 = Mo … 7 = So
  const hours = OPENING_HOURS[weekday];
  const [h, m] = b.time.split(':').map(Number);
  const startMin = h * 60 + m;
  if (!hours) fail('An diesem Tag haben wir geschlossen (Montag und Sonntag).');
  if (startMin < hours[0] || startMin + b.duration > hours[1]) {
    fail('Dieser Termin liegt (mit der Dauer der Leistung) außerhalb unserer Öffnungszeiten.');
  }
  return b;
}

function sendConfirmation(b) {
  const when = Utilities.formatDate(b.start, TIMEZONE, "dd.MM.yyyy 'um' HH:mm");
  MailApp.sendEmail({
    to: b.email,
    subject: `Ihr Termin bei ${SALON} am ${when} Uhr`,
    body:
`Hallo ${b.name},

vielen Dank für Ihre Buchung! Hier Ihre Termindaten:

Leistung: ${b.service}
Termin:   ${when} Uhr (ca. ${b.duration} Minuten)
Adresse:  Musterstraße 12, 8010 Graz

Sie können den Termin nicht wahrnehmen? Bitte sagen Sie uns bis spätestens
24 Stunden vorher unter +43 316 000 000 Bescheid.

Wir freuen uns auf Sie!
Ihr Team von ${SALON}

(Hinweis: Dies ist ein Übungsprojekt – der Salon ist erfunden.)`,
    name: SALON,
  });
}

function notifyOwner(b) {
  const when = Utilities.formatDate(b.start, TIMEZONE, "EEE dd.MM.yyyy HH:mm");
  MailApp.sendEmail(ownerEmail(), `[${SALON}] Neue Buchung: ${when} – ${b.name}`,
    `${b.service} (${b.duration} Min)\n${when}\n\n${b.name}\n${b.email}\n${b.phone}\n\nNachricht: ${b.message || '–'}`);
}

function getCalendar() {
  return CalendarApp.getCalendarsByName(CALENDAR_NAME)[0]
    || CalendarApp.createCalendar(CALENDAR_NAME, { timeZone: TIMEZONE });
}

function ownerEmail() {
  return Session.getEffectiveUser().getEmail();
}

function reply(ok, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok, message }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Einmal im Editor ausführen: fragt die Berechtigungen ab und legt den Kalender an. */
function setup() {
  getCalendar();
  SpreadsheetApp.openById(SHEET_ID).getSheets()[0].setFrozenRows(1);
  Logger.log('Fertig. Kalender und Sheet sind bereit. Besitzer: ' + ownerEmail());
}
