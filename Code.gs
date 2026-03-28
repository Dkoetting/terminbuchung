// ====================== KONFIGURATION ======================
const SHEET_NAME     = 'Terminbuchung';
const RESEND_API_KEY = 're_g4crPYd9_BcyoobZMaxCt1dXejioQnTqJ'; // Dein Resend-API-Key
const NOTIFY_EMAIL   = 'dr-dirk@dr-dirkinstitute.org';
const FROM_EMAIL     = 'info@edvkonzepte.de';
// ====================== HILFSFUNKTIONEN ======================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function normalizeSheetDate(value, tz) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  }
  const s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  const dmy = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
  if (dmy) {
    const y = dmy[3].length === 2 ? '20' + dmy[3] : dmy[3];
    return `${y}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
  }
  return s;
}
function normalizeSheetTime(value, tz) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, tz, 'HH:mm');
  }
  const s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d)) return Utilities.formatDate(d, tz, 'HH:mm');
  }
  const m = s.match(/(\d{1,2}):(\d{2})/);
  if (m) return `${m[1].padStart(2, '0')}:${m[2]}`;
  return s;
}
// ICS-Datei generieren – Zeitzone Europe/Berlin
function generateIcs_(payload) {
  const dateStr = payload.date || '';                     // 2026-03-27
  const timeStr = payload.time || '00:00';               // 16:30
  const endStr  = payload.end_time || payload.time || '00:00'; // 17:00
  // "2026-03-27" + "16:30" → "20260327T163000"
  const toIcsDt = function(d, t) {
    return d.replace(/-/g, '') + 'T' + t.replace(':', '') + '00';
  };
  const stamp   = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const uid     = 'terminbuchung-' + Utilities.getUuid() + '@edvkonzepte.de';
  const summary = (payload.meeting_type || 'Termin') + ' – ' + (payload.name || '');
  const desc    = 'Thema: ' + (payload.topic || '') +
                  (payload.notes ? '\\nAnmerkungen: ' + payload.notes : '') +
                  (payload.phone ? '\\nTel: ' + payload.phone : '');
  return [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//K&N EDV Konzepte//Terminbuchung//DE',
    'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:' + uid,
    'DTSTAMP:' + stamp,
    'DTSTART;TZID=Europe/Berlin:' + toIcsDt(dateStr, timeStr),
    'DTEND;TZID=Europe/Berlin:'   + toIcsDt(dateStr, endStr),
    'SUMMARY:' + summary,
    'DESCRIPTION:' + desc,
    'ORGANIZER;CN=Dr. Dirk Koetting:mailto:' + FROM_EMAIL,
    'ATTENDEE;CN=' + (payload.name || '') + ':mailto:' + (payload.email || ''),
    'END:VEVENT',
    'END:VCALENDAR'
  ].join('\r\n');
}
// Resend-Mail senden – icsContent optional
function sendResendMail(to, subject, textBody, icsContent) {
  const url = 'https://api.resend.com/emails';
  const mailPayload = {
    from: FROM_EMAIL,
    to: to,
    subject: subject,
    text: textBody
  };
  if (icsContent) {
    mailPayload.attachments = [{
      filename: 'termin.ics',
      content:  Utilities.base64Encode(icsContent)
    }];
  }
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + RESEND_API_KEY },
    payload: JSON.stringify(mailPayload),
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  Logger.log('Resend status: ' + res.getResponseCode());
  Logger.log('Resend response: ' + res.getContentText());
}
// ====================== GET: GEBUCHTE SLOTS ======================
function doGet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse({ success: false, booked: [], error: `Sheet "${SHEET_NAME}" nicht gefunden` });
    const data = sheet.getDataRange().getValues();
    const tz = ss.getSpreadsheetTimeZone();
    const booked = [];
    for (let i = 1; i < data.length; i++) {
      const dateStr = normalizeSheetDate(data[i][1], tz);
      const timeStr = normalizeSheetTime(data[i][2], tz);
      if (dateStr && timeStr) {
        booked.push({
          date: dateStr,
          time: timeStr,
          key: `${dateStr} ${timeStr}`
        });
      }
    }
    return jsonResponse({ success: true, booked });
  } catch (err) {
    return jsonResponse({ success: false, booked: [], error: String(err) });
  }
}
// ====================== POST: BUCHUNG ANNEHMEN + MAILS ======================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ success: false, error: 'Keine POST-Daten empfangen' });
    }
    const payload = JSON.parse(e.postData.contents);
    // Pflichtfelder aus dem Frontend
    if (!payload.date || !payload.time || !payload.name || !payload.email) {
      return jsonResponse({ success: false, error: 'Pflichtfelder fehlen' });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return jsonResponse({ success: false, error: `Sheet "${SHEET_NAME}" nicht gefunden` });
    const data = sheet.getDataRange().getValues();
    const tz = ss.getSpreadsheetTimeZone();
    const incomingKey = `${payload.date} ${payload.time}`;
    // Prüfen, ob Slot bereits existiert
    for (let i = 1; i < data.length; i++) {
      const rowDate = normalizeSheetDate(data[i][1], tz);
      const rowTime = normalizeSheetTime(data[i][2], tz);
      const rowKey = `${rowDate} ${rowTime}`;
      if (rowKey === incomingKey) {
        return jsonResponse({ success: false, error: 'Slot bereits gebucht' });
      }
    }
    const id = Utilities.getUuid();
    const now = new Date().toISOString();
    // In Sheet schreiben
    sheet.appendRow([
      id,
      payload.date,
      payload.time,
      payload.end_time || '',
      payload.meeting_type || '',
      payload.name,
      payload.email,
      payload.phone || '',
      payload.company || '',
      payload.topic || '',
      payload.notes || '',
      now
    ]);
    // ICS einmal generieren – für beide Mails
    const ics = generateIcs_(payload);
    // ------- E-MAILS MIT RESEND -------
    // Benachrichtigung an dich (mit ICS)
    const infoText =
      'Neue Terminbuchung\n\n' +
      'Datum: ' + payload.date + '\n' +
      'Zeit: ' + payload.time + (payload.end_time ? ' – ' + payload.end_time : '') + '\n' +
      'Terminart: ' + (payload.meeting_type || '') + '\n\n' +
      'Name: ' + payload.name + '\n' +
      'E-Mail: ' + payload.email + '\n' +
      'Telefon: ' + (payload.phone || '') + '\n' +
      'Firma: ' + (payload.company || '') + '\n' +
      'Thema: ' + (payload.topic || '') + '\n\n' +
      'Anmerkungen:\n' + (payload.notes || '');
    sendResendMail(NOTIFY_EMAIL, 'Neue Terminbuchung', infoText, ics);
    // Bestätigung an Kund:in (mit ICS)
    if (payload.email) {
      const confirmText =
        'Guten Tag ' + payload.name + ',\n\n' +
        'vielen Dank für Ihre Terminbuchung.\n\n' +
        'Datum: ' + payload.date + '\n' +
        'Zeit: ' + payload.time + (payload.end_time ? ' – ' + payload.end_time : '') + '\n' +
        'Terminart: ' + (payload.meeting_type || '') + '\n' +
        'Thema: ' + (payload.topic || '') + '\n\n' +
        'Der Termin ist als ICS-Datei angehängt – einfach öffnen zum Import in Ihren Kalender.\n\n' +
        'Falls Sie den Termin verschieben oder absagen möchten, antworten Sie einfach auf diese E-Mail.\n\n' +
        'Herzliche Grüße\n' +
        'Dr. Dirk Kötting';
      sendResendMail(payload.email, 'Ihre Terminbestätigung', confirmText, ics);
    }
    // Antwort an Frontend
    return jsonResponse({
      success: true,
      id,
      booked: {
        date: payload.date,
        time: payload.time,
        key: incomingKey
      }
    });
  } catch (err) {
    return jsonResponse({ success: false, error: String(err) });
  }
}
