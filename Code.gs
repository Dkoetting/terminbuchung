// ====================== KONFIGURATION ======================
const SHEET_NAME    = 'Terminbuchung';
const RESEND_API_KEY = 're_g4crPYd9_BcyoobZMaxCt1dXejioQnTqJ';
const NOTIFY_EMAIL  = 'dr-dirk@dr-dirkinstitute.org';
const FROM_EMAIL    = 'info@edvkonzepte.de';


// ====================== HILFSFUNKTIONEN ======================

function getSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Tabellenblatt "' + SHEET_NAME + '" wurde nicht gefunden.');
  }
  return sheet;
}

// Datum "2026-03-27" + Zeit "16:30" → "20260327T163000"
function toIcsDateTime_(dateStr, timeStr) {
  const d = dateStr.replace(/-/g, '');          // 20260327
  const t = timeStr.replace(':', '') + '00';    // 163000
  return d + 'T' + t;
}

// ICS-Datei generieren (Zeitzone Europe/Berlin)
function generateIcs_(data) {
  const uid    = 'terminbuchung-' + Date.now() + '@edvkonzepte.de';
  const stamp  = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const start  = toIcsDateTime_(data.date, data.time);
  const end    = toIcsDateTime_(data.date, data.endtime);
  const summary = (data.meetingtype || 'Termin') + ' – ' + (data.name || '');
  const desc   = 'Thema: ' + (data.topic || '') +
                 (data.notes ? '\\nAnmerkungen: ' + data.notes : '') +
                 (data.phone ? '\\nTel: ' + data.phone : '');

  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//K&N EDV Konzepte//Terminbuchung//DE',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:' + uid,
    'DTSTAMP:' + stamp,
    'DTSTART;TZID=Europe/Berlin:' + start,
    'DTEND;TZID=Europe/Berlin:' + end,
    'SUMMARY:' + summary,
    'DESCRIPTION:' + desc,
    'ORGANIZER;CN=Dr. Dirk Kötting:mailto:' + FROM_EMAIL,
    'ATTENDEE;CN=' + (data.name || '') + ':mailto:' + (data.email || ''),
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');

  return ics;
}

// E-Mail via Resend senden – optional mit ICS-Anhang
function sendResendMail_(to, subject, textBody, icsContent) {
  const url = 'https://api.resend.com/emails';

  const payload = {
    from:    FROM_EMAIL,
    to:      to,
    subject: subject,
    text:    textBody
  };

  if (icsContent) {
    payload.attachments = [{
      filename: 'termin.ics',
      content:  Utilities.base64Encode(icsContent)
    }];
  }

  const options = {
    method:      'post',
    contentType: 'application/json',
    headers:     { Authorization: 'Bearer ' + RESEND_API_KEY },
    payload:     JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = UrlFetchApp.fetch(url, options);
  Logger.log('Resend status: ' + res.getResponseCode());
  Logger.log('Resend response: ' + res.getContentText());
}


// ====================== API: BUCHUNG ANNEHMEN ======================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Keine POST-Daten empfangen.');
    }

    const data = JSON.parse(e.postData.contents);

    const sheet = getSheet_();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Erstellt', 'Datum', 'Start', 'Ende', 'Terminart',
        'Name', 'E-Mail', 'Telefon', 'Firma', 'Thema', 'Anmerkungen'
      ]);
    }

    sheet.appendRow([
      new Date(),
      data.date     || '',
      data.time     || '',
      data.endtime  || '',
      data.meetingtype || '',
      data.name     || '',
      data.email    || '',
      data.phone    || '',
      data.company  || '',
      data.topic    || '',
      data.notes    || ''
    ]);

    // ICS einmal generieren – für beide Mails verwenden
    const ics = generateIcs_(data);

    // -------- Benachrichtigung an dich (mit ICS) --------
    const notifyText =
      'Neue Terminbuchung\n\n' +
      'Datum: '    + (data.date     || '') + '\n' +
      'Zeit: '     + (data.time     || '') + ' – ' + (data.endtime || '') + '\n' +
      'Terminart: '+ (data.meetingtype || '') + '\n\n' +
      'Name: '     + (data.name     || '') + '\n' +
      'E-Mail: '   + (data.email    || '') + '\n' +
      'Telefon: '  + (data.phone    || '') + '\n' +
      'Firma: '    + (data.company  || '') + '\n' +
      'Thema: '    + (data.topic    || '') + '\n\n' +
      'Anmerkungen:\n' + (data.notes || '');

    sendResendMail_(NOTIFY_EMAIL, 'Neue Terminbuchung', notifyText, ics);

    // -------- Bestätigung an Kund:in (mit ICS) --------
    if (data.email) {
      const confirmText =
        'Guten Tag ' + (data.name || '') + ',\n\n' +
        'vielen Dank für Ihre Terminbuchung.\n\n' +
        'Datum: '    + (data.date     || '') + '\n' +
        'Zeit: '     + (data.time     || '') + ' – ' + (data.endtime || '') + '\n' +
        'Terminart: '+ (data.meetingtype || '') + '\n' +
        'Thema: '    + (data.topic    || '') + '\n\n' +
        'Der Termin wurde als ICS-Datei angehängt – einfach öffnen, um ihn in Ihren Kalender zu importieren.\n\n' +
        'Falls Sie den Termin verschieben oder absagen möchten, antworten Sie einfach auf diese E-Mail.\n\n' +
        'Herzliche Grüße\n' +
        'Dr. Dirk Kötting';

      sendResendMail_(data.email, 'Ihre Terminbestätigung', confirmText, ics);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Fehler in doPost: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ====================== API: GEBUCHTE SLOTS LADEN ======================

function doGet() {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return ContentService
        .createTextOutput(JSON.stringify({ booked: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const range  = sheet.getRange(2, 1, lastRow - 1, 11);
    const values = range.getValues();

    const booked = values.map(function(row) {
      return {
        created:     row[0],
        date:        row[1],
        time:        row[2],
        endtime:     row[3],
        meetingtype: row[4],
        name:        row[5],
        email:       row[6],
        phone:       row[7],
        company:     row[8],
        topic:       row[9],
        notes:       row[10]
      };
    });

    return ContentService
      .createTextOutput(JSON.stringify({ booked: booked }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('Fehler in doGet: ' + err);
    return ContentService
      .createTextOutput(JSON.stringify({ booked: [], error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
