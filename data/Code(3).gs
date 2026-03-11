const SHEET_NAME    = 'Terminbuchung';
const RESEND_API_KEY = 're_XXXXXXXXXXXXXXXX'; // <-- Ihren Resend API-Key eintragen
const NOTIFY_EMAIL  = 'dr-dirk@dr-dirkinstitute.org';
const FROM_EMAIL    = 'info@edvkonzepte.de';

// -------------------------------------------------------
// GET: gebuchte Slots liefern (fuer Slot-Blocking im HTML)
// -------------------------------------------------------
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data  = sheet.getDataRange().getValues();
    const tz    = Session.getScriptTimeZone();
    const booked = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][2]) {
        // Datum: Date-Objekt oder String -> immer YYYY-MM-DD
        let dateStr;
        if (data[i][1] instanceof Date) {
          dateStr = Utilities.formatDate(data[i][1], tz, 'yyyy-MM-dd');
        } else {
          dateStr = String(data[i][1]).trim();
        }
        // Zeit: Date-Objekt oder String -> immer HH:mm
        let timeStr;
        if (data[i][2] instanceof Date) {
          timeStr = Utilities.formatDate(data[i][2], tz, 'HH:mm');
        } else {
          timeStr = String(data[i][2]).trim().substring(0, 5);
        }
        booked.push({ date: dateStr, time: timeStr });
      }
    }
    return ContentService
      .createTextOutput(JSON.stringify({ booked, ts: Date.now() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ booked: [], error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------
// POST: Buchung speichern + E-Mails versenden
// -------------------------------------------------------
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // Validierung
    if (!payload.date || !payload.time || !payload.name || !payload.email) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: 'Pflichtfelder fehlen' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Doppelbuchung pruefen
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === payload.date && data[i][2] === payload.time) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: 'Slot bereits gebucht' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // In Sheet schreiben
    const id  = Utilities.getUuid();
    const now = new Date().toISOString();
    sheet.appendRow([
      id,
      payload.date,
      payload.time,
      payload.end_time    || '',
      payload.meeting_type || '',
      payload.name,
      payload.email,
      payload.phone       || '',
      payload.company     || '',
      payload.topic       || '',
      payload.notes       || '',
      now
    ]);

    // E-Mails versenden
    sendEmails(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, id }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// -------------------------------------------------------
// E-Mail via Resend
// -------------------------------------------------------
function sendEmails(b) {
  const notifyHtml = `
    <div style="font-family:sans-serif;font-size:14px;color:#1a1a1a;max-width:560px;">
      <h2 style="color:#21808d;margin-bottom:24px;">Neue Terminbuchung</h2>
      <table cellpadding="8" style="border-collapse:collapse;width:100%;">
        <tr style="background:#f5f5f5;"><td><strong>Name</strong></td><td>${b.name}</td></tr>
        <tr><td><strong>E-Mail</strong></td><td>${b.email}</td></tr>
        <tr style="background:#f5f5f5;"><td><strong>Telefon</strong></td><td>${b.phone || '-'}</td></tr>
        <tr><td><strong>Firma</strong></td><td>${b.company || '-'}</td></tr>
        <tr style="background:#f5f5f5;"><td><strong>Datum</strong></td><td>${b.date}</td></tr>
        <tr><td><strong>Uhrzeit</strong></td><td>${b.time} - ${b.end_time} Uhr</td></tr>
        <tr style="background:#f5f5f5;"><td><strong>Terminart</strong></td><td>${b.meeting_type}</td></tr>
        <tr><td><strong>Thema</strong></td><td>${b.topic || '-'}</td></tr>
        <tr style="background:#f5f5f5;"><td><strong>Anmerkungen</strong></td><td>${b.notes || '-'}</td></tr>
      </table>
    </div>`;

  const confirmHtml = `
    <div style="font-family:sans-serif;font-size:14px;color:#1a1a1a;max-width:560px;">
      <h2 style="color:#21808d;">Terminbestaetigung</h2>
      <p>Sehr geehrte/r ${b.name},</p>
      <p>Ihr Termin wurde erfolgreich gebucht:</p>
      <table cellpadding="8" style="border-collapse:collapse;width:100%;background:#f5f5f5;border-radius:8px;">
        <tr><td><strong>Datum</strong></td><td>${b.date}</td></tr>
        <tr><td><strong>Uhrzeit</strong></td><td>${b.time} - ${b.end_time} Uhr</td></tr>
        <tr><td><strong>Terminart</strong></td><td>${b.meeting_type}</td></tr>
        <tr><td><strong>Thema</strong></td><td>${b.topic || '-'}</td></tr>
      </table>
      <p style="margin-top:24px;">Bei Rueckfragen antworten Sie direkt auf diese E-Mail.</p>
      <p style="color:#666;font-size:13px;">
        Dr. Dirk Koetting<br>
        IT Governance &amp; AI Policy Beratung<br>
        <a href="mailto:${NOTIFY_EMAIL}" style="color:#21808d;">${NOTIFY_EMAIL}</a>
      </p>
    </div>`;

  const headers = {
    'Authorization': `Bearer ${RESEND_API_KEY}`,
    'Content-Type': 'application/json'
  };

  // Benachrichtigung an Sie
  UrlFetchApp.fetch('https://api.resend.com/emails', {
    method: 'post',
    headers,
    payload: JSON.stringify({
      from: FROM_EMAIL,
      to:   NOTIFY_EMAIL,
      reply_to: b.email,
      subject: `Neue Buchung: ${b.name} | ${b.date} ${b.time}`,
      html: notifyHtml
    }),
    muteHttpExceptions: true
  });

  // Bestaetigung an den Bucher
  UrlFetchApp.fetch('https://api.resend.com/emails', {
    method: 'post',
    headers,
    payload: JSON.stringify({
      from: FROM_EMAIL,
      to:   b.email,
      subject: `Terminbestaetigung: ${b.date}, ${b.time} Uhr`,
      html: confirmHtml
    }),
    muteHttpExceptions: true
  });
}
