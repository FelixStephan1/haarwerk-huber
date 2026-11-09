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
const ADDRESS = 'Musterstraße 12, 8010 Graz';
const PHONE = '+43 316 000 000';
const WEEKDAYS = ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'];

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
  const weekday = WEEKDAYS[Number(Utilities.formatDate(b.start, TIMEZONE, 'u')) - 1];
  const dateLong = `${weekday}, ${Utilities.formatDate(b.start, TIMEZONE, 'dd.MM.yyyy')}`;
  const timeRange = `${Utilities.formatDate(b.start, TIMEZONE, 'HH:mm')} – ${Utilities.formatDate(b.end, TIMEZONE, 'HH:mm')} Uhr`;

  // Klartext-Version für E-Mail-Programme ohne HTML
  const plain =
`Hallo ${b.name},

vielen Dank für Ihre Buchung! Hier Ihre Termindaten:

Leistung: ${b.service}
Termin:   ${dateLong}, ${timeRange}
Adresse:  ${ADDRESS}

Sie können den Termin nicht wahrnehmen? Bitte sagen Sie uns bis spätestens
24 Stunden vorher unter ${PHONE} Bescheid.

Wir freuen uns auf Sie!
Ihr Team von ${SALON}

(Hinweis: Dies ist ein Übungsprojekt – der Salon ist erfunden.)`;

  const row = (label, value) =>
    `<tr><td style="padding:8px 0;color:#6f655c;width:110px;vertical-align:top;">${label}</td>` +
    `<td style="padding:8px 0;color:#2b2622;">${value}</td></tr>`;

  const html = `
<div style="background:#faf7f2;padding:32px 16px;font-family:Georgia,'Times New Roman',serif;">
  <table role="presentation" cellpadding="0" cellspacing="0" style="max-width:520px;width:100%;margin:0 auto;background:#ffffff;border:1px solid #e4dcd0;">
    <tr><td style="padding:28px 32px;text-align:center;border-bottom:1px solid #e4dcd0;">
      <div style="font-size:28px;color:#2b2622;">${SALON}</div>
      <div style="font-size:14px;color:#6f655c;margin-top:4px;">Ihr Friseur in Graz</div>
    </td></tr>
    <tr><td style="padding:28px 32px;color:#2b2622;font-size:16px;line-height:1.6;">
      <p style="margin:0 0 16px;">Hallo ${esc(b.name)},</p>
      <p style="margin:0 0 20px;">vielen Dank für Ihre Buchung – wir freuen uns auf Sie!</p>
      <div style="background:#faf7f2;border-left:3px solid #8a5a3c;padding:16px 20px;">
        <table role="presentation" cellpadding="0" cellspacing="0" style="width:100%;font-size:15px;font-family:Arial,sans-serif;">
          ${row('Leistung', esc(b.service))}
          ${row('Datum', dateLong)}
          ${row('Uhrzeit', `${timeRange} <span style="color:#6f655c;">(ca. ${b.duration} Min.)</span>`)}
          ${row('Adresse', ADDRESS)}
          ${b.message ? row('Ihre Nachricht', esc(b.message)) : ''}
        </table>
      </div>
      <p style="margin:24px 0 0;font-size:15px;">
        Sie können den Termin nicht wahrnehmen? Bitte sagen Sie uns bis spätestens
        24&nbsp;Stunden vorher Bescheid:
        <a href="tel:${PHONE.replace(/\s/g, '')}" style="color:#8a5a3c;">${PHONE}</a>
      </p>
      <p style="margin:24px 0 0;">Bis bald!<br>Ihr Team von ${SALON}</p>
    </td></tr>
    <tr><td style="padding:16px 32px;text-align:center;font-family:Arial,sans-serif;font-size:12px;color:#6f655c;border-top:1px solid #e4dcd0;">
      ${SALON} · ${ADDRESS} · ${PHONE}<br>
      Übungsprojekt – dieser Friseursalon ist erfunden.
    </td></tr>
  </table>
</div>`;

  MailApp.sendEmail({
    to: b.email,
    subject: `Terminbestätigung: ${dateLong}, ${Utilities.formatDate(b.start, TIMEZONE, 'HH:mm')} Uhr – ${SALON}`,
    body: plain,
    htmlBody: html,
    name: SALON,
  });
}

/** Text für HTML unschädlich machen (Kundeneingaben!) */
function esc(text) {
  return String(text).replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
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
