// Haarwerk Huber – gemeinsamer Code für index.html und termin.html

// URL der Apps-Script-Web-App (Bereitstellen → Web-App). Ändert sich nur bei einer NEUEN Bereitstellung.
const BOOKING_URL = 'https://script.google.com/macros/s/AKfycbwmrHakHJ1OROm2lAyTSuZxV8D6NL0SzJc1zIDO6RZA3aP7IMScyE-cWyZx6KR5PzStaw/exec';
const PHONE = '+43 316 000 000';

// Leistung → Dauer in Minuten (muss zu SERVICES in apps-script/Code.gs passen)
const SERVICES = {
  'Damen – Waschen, Schneiden, Föhnen': 60,
  'Herren – Waschen, Schneiden, Styling': 30,
  'Kinder bis 12 Jahre': 30,
  'Färben (Ansatz)': 90,
  'Strähnen': 120,
  'Bart trimmen': 15,
  'Hochsteckfrisur': 60,
};

// Öffnungszeiten (Wochentag 0 = Sonntag): [von, bis] in Minuten ab Mitternacht
const HOURS = { 2: [540, 1080], 3: [540, 1080], 4: [540, 1080], 5: [540, 1080], 6: [480, 780] };

/** Heutiges Datum als JJJJ-MM-TT (für das min-Attribut des Datumsfelds) */
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

/** Uhrzeiten-Auswahl füllen: nur Zeiten, bei denen die Leistung vor Ladenschluss fertig ist */
function fillTimes(select, dateStr, duration) {
  select.innerHTML = '';
  if (!dateStr) {
    select.add(new Option('Bitte zuerst Datum wählen', ''));
    return;
  }
  const hours = HOURS[new Date(dateStr + 'T12:00').getDay()];
  if (!hours) {
    select.add(new Option('An diesem Tag geschlossen', ''));
    return;
  }
  select.add(new Option('Bitte wählen …', ''));
  for (let m = hours[0]; m + duration <= hours[1]; m += 15) {
    const t = String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0');
    select.add(new Option(t, t));
  }
}

/** Telefonnummer prüfen (gleiche Regel wie isValidPhone im Apps Script) */
function isValidPhone(phone) {
  if (!/^\+?[0-9 ()\/-]+$/.test(phone)) return false;
  const digits = phone.replace(/\D/g, '').length;
  return digits >= 6 && digits <= 15;
}

/** Anfrage an das Apps Script. GET: params als Objekt, POST: FormData oder Objekt */
async function callBooking(method, params) {
  const body = new URLSearchParams(params);
  const res = method === 'GET'
    ? await fetch(`${BOOKING_URL}?${body}`)
    : await fetch(BOOKING_URL, { method: 'POST', body });
  return res.json();
}

/** Statusmeldung anzeigen. kind: 'ok', 'error' oder '' */
function showStatus(el, kind, text) {
  el.className = kind;
  el.textContent = text; // textContent statt innerHTML: Text wird nie als HTML ausgeführt
}
