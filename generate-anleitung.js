const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, PageNumber, PageBreak, LevelFormat, ImageRun
} = require("docx");

const TEAL = "21808D";
const DARK = "1F2121";
const GRAY = "626C71";
const LIGHT_BG = "F0F7F8";
const WHITE = "FFFFFF";
const BORDER_COLOR = "D0D0D0";

const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

const PAGE_WIDTH = 11906; // A4
const PAGE_HEIGHT = 16838;
const MARGIN = 1440;
const CONTENT_WIDTH = PAGE_WIDTH - 2 * MARGIN; // 9026

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 200 },
    children: [new TextRun({ text, bold: true, size: 32, font: "Arial", color: TEAL })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 160 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: DARK })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...opts.paragraphOpts,
    children: [new TextRun({ text, size: 22, font: "Arial", color: DARK, ...opts })]
  });
}

function bold(text) {
  return new TextRun({ text, size: 22, font: "Arial", color: DARK, bold: true });
}

function normal(text) {
  return new TextRun({ text, size: 22, font: "Arial", color: DARK });
}

function teal(text) {
  return new TextRun({ text, size: 22, font: "Arial", color: TEAL, bold: true });
}

function para(...runs) {
  return new Paragraph({ spacing: { after: 120 }, children: runs });
}

function tipBox(text) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [new TableRow({
      children: [new TableCell({
        borders: {
          top: { style: BorderStyle.SINGLE, size: 1, color: TEAL },
          bottom: { style: BorderStyle.SINGLE, size: 1, color: TEAL },
          left: { style: BorderStyle.SINGLE, size: 6, color: TEAL },
          right: { style: BorderStyle.SINGLE, size: 1, color: TEAL }
        },
        shading: { fill: LIGHT_BG, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        children: [new Paragraph({
          children: [
            new TextRun({ text: "Tipp: ", size: 22, font: "Arial", bold: true, color: TEAL }),
            new TextRun({ text, size: 22, font: "Arial", color: DARK })
          ]
        })]
      })]
    })]
  });
}

function codeBlock(text) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [new TableRow({
      children: [new TableCell({
        borders,
        shading: { fill: "F5F5F5", type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 200, right: 200 },
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        children: [new Paragraph({
          children: [new TextRun({ text, size: 20, font: "Consolas", color: DARK })]
        })]
      })]
    })]
  });
}

// Build numbering config
const numberingConfig = [
  {
    reference: "bullets",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "\u2022",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  },
  {
    reference: "steps",
    levels: [{
      level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  },
  {
    reference: "steps2",
    levels: [{
      level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  },
  {
    reference: "steps3",
    levels: [{
      level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  },
  {
    reference: "steps4",
    levels: [{
      level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  },
  {
    reference: "steps5",
    levels: [{
      level: 0, format: LevelFormat.DECIMAL, text: "%1.",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  },
  {
    reference: "bullets2",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "\u2022",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  },
  {
    reference: "bullets3",
    levels: [{
      level: 0, format: LevelFormat.BULLET, text: "\u2022",
      alignment: AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720, hanging: 360 } } }
    }]
  }
];

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, font: "Arial", color: DARK })]
  });
}

function bulletRuns(ref, ...runs) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 80 },
    children: runs
  });
}

function step(ref, ...runs) {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { after: 100 },
    children: runs
  });
}

// Table helper for settings
function settingsTable(headers, rows) {
  const colCount = headers.length;
  const colWidth = Math.floor(CONTENT_WIDTH / colCount);
  const colWidths = headers.map(() => colWidth);

  const headerRow = new TableRow({
    children: headers.map(h => new TableCell({
      borders,
      shading: { fill: TEAL, type: ShadingType.CLEAR },
      margins: cellMargins,
      width: { size: colWidth, type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text: h, size: 20, font: "Arial", color: WHITE, bold: true })]
      })]
    }))
  });

  const dataRows = rows.map((row, i) => new TableRow({
    children: row.map(cell => new TableCell({
      borders,
      shading: { fill: i % 2 === 0 ? "FAFAFA" : WHITE, type: ShadingType.CLEAR },
      margins: cellMargins,
      width: { size: colWidth, type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text: cell, size: 20, font: "Arial", color: DARK })]
      })]
    }))
  }));

  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

// Load logo
let logoData;
try {
  logoData = fs.readFileSync("assets/Logo.jpg");
} catch(e) {
  logoData = null;
}

const children = [];

// ============================================================
// TITLE PAGE
// ============================================================
children.push(new Paragraph({ spacing: { before: 3000 } }));

if (logoData) {
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 400 },
    children: [new ImageRun({
      type: "jpg",
      data: logoData,
      transformation: { width: 120, height: 120 },
      altText: { title: "Logo", description: "Dr. Dirk Koetting Logo", name: "logo" }
    })]
  }));
}

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 200 },
  children: [new TextRun({ text: "Terminbuchung", size: 52, font: "Arial", bold: true, color: TEAL })]
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 100 },
  children: [new TextRun({ text: "Einrichtungs- und Bedienungsanleitung", size: 28, font: "Arial", color: GRAY })]
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  spacing: { after: 600 },
  children: [new TextRun({ text: "Version 1.0 \u2013 Stand: Maerz 2026", size: 22, font: "Arial", color: GRAY })]
}));

children.push(new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: "Dr. Dirk Koetting \u2013 IT Governance & AI Policy Beratung", size: 22, font: "Arial", color: DARK })]
}));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// INHALTSVERZEICHNIS
// ============================================================
children.push(heading1("Inhalt"));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("1. Was ist die Terminbuchung?")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("2. Was wird benoetigt?")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("3. Google Sheet einrichten")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("4. Google Apps Script einrichten")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("5. E-Mail-Versand mit Resend einrichten")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("6. Frontend anpassen")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("7. Seite online stellen")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("8. Taegliche Nutzung")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("9. Anpassungen & Einstellungen")] }));
children.push(new Paragraph({ spacing: { after: 80 }, children: [normal("10. Haeufige Fragen")] }));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 1. WAS IST DIE TERMINBUCHUNG?
// ============================================================
children.push(heading1("1. Was ist die Terminbuchung?"));

children.push(p("Die Terminbuchung ist eine einfache Webseite, ueber die Interessenten und Kunden online einen Beratungstermin bei Ihnen buchen koennen."));

children.push(para(bold("So funktioniert es:")));

children.push(bullet("Der Besucher waehlt eine Terminart (z.B. Erstgespraech oder Beratungstermin)"));
children.push(bullet("Er sieht einen Kalender mit verfuegbaren Tagen (Montag bis Freitag)"));
children.push(bullet("Er waehlt ein freies Zeitfenster (z.B. 16:00 - 16:30)"));
children.push(bullet("Er fuellt ein kurzes Formular aus (Name, E-Mail, Thema)"));
children.push(bullet("Sie erhalten eine E-Mail-Benachrichtigung"));
children.push(bullet("Der Kunde erhaelt eine Bestaetigungsmail"));
children.push(bullet("Der Termin wird automatisch in Google Sheets gespeichert"));
children.push(bullet("Der Slot ist sofort fuer andere Besucher gesperrt"));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(tipBox("Der Kunde kann sich nach der Buchung eine .ics-Datei herunterladen, um den Termin direkt in Outlook oder einen anderen Kalender zu importieren."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 2. WAS WIRD BENOETIGT?
// ============================================================
children.push(heading1("2. Was wird benoetigt?"));

children.push(p("Bevor Sie starten, brauchen Sie folgende Dinge:"));
children.push(new Paragraph({ spacing: { after: 80 } }));

children.push(settingsTable(
  ["Was", "Wozu", "Kosten"],
  [
    ["Google-Konto", "Fuer Google Sheets und Apps Script", "Kostenlos"],
    ["Resend-Konto", "Fuer den E-Mail-Versand", "Kostenlos (bis 3.000 Mails/Monat)"],
    ["Eigene Domain (optional)", "Damit E-Mails von Ihrer Adresse kommen", "ca. 10-15 EUR/Jahr"],
    ["Webhosting", "Damit die Seite im Internet erreichbar ist", "Kostenlos moeglich (z.B. GitHub Pages, Vercel)"],
  ]
));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 3. GOOGLE SHEET EINRICHTEN
// ============================================================
children.push(heading1("3. Google Sheet einrichten"));

children.push(p("Das Google Sheet ist Ihre Datenbank. Hier werden alle Buchungen gespeichert."));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(heading2("Schritt fuer Schritt"));

children.push(step("steps", normal("Oeffnen Sie "), bold("Google Sheets"), normal(" (sheets.google.com) und erstellen Sie eine neue Tabelle.")));
children.push(step("steps", normal("Benennen Sie das erste Tabellenblatt (Tab unten) in "), bold("Terminbuchung"), normal(" um.")));
children.push(step("steps", normal("Tragen Sie in die erste Zeile folgende Spaltenueberschriften ein:")));

children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(settingsTable(
  ["Spalte", "Ueberschrift", "Inhalt"],
  [
    ["A", "ID", "Wird automatisch generiert"],
    ["B", "Datum", "z.B. 2026-03-25"],
    ["C", "Uhrzeit", "z.B. 16:00"],
    ["D", "Ende", "z.B. 16:30"],
    ["E", "Terminart", "z.B. Erstgespraech"],
    ["F", "Name", "Name des Kunden"],
    ["G", "E-Mail", "E-Mail des Kunden"],
    ["H", "Telefon", "Optional"],
    ["I", "Firma", "Optional"],
    ["J", "Thema", "Gewaehltes Beratungsthema"],
    ["K", "Anmerkungen", "Freitext des Kunden"],
    ["L", "Erstellt", "Zeitstempel der Buchung"],
  ]
));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(tipBox("Die Spaltenueberschriften dienen nur Ihrer Uebersicht. Das Script fuellt die Daten automatisch in der richtigen Reihenfolge ein."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 4. GOOGLE APPS SCRIPT
// ============================================================
children.push(heading1("4. Google Apps Script einrichten"));

children.push(p("Das Google Apps Script ist das Backend \u2013 es nimmt Buchungen entgegen, speichert sie im Sheet und versendet E-Mails."));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(heading2("Script anlegen"));

children.push(step("steps2", normal("Oeffnen Sie Ihr Google Sheet.")));
children.push(step("steps2", normal("Klicken Sie oben auf "), bold("Erweiterungen"), normal(" > "), bold("Apps Script"), normal(".")));
children.push(step("steps2", normal("Es oeffnet sich der Script-Editor. Loeschen Sie den vorhandenen Code.")));
children.push(step("steps2", normal("Kopieren Sie den gesamten Inhalt der Datei "), bold("Code.gs"), normal(" aus dem Projektordner und fuegen Sie ihn ein.")));

children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(heading2("Einstellungen im Script anpassen"));

children.push(p("Ganz oben im Script finden Sie vier Einstellungen:"));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(settingsTable(
  ["Einstellung", "Beschreibung", "Beispiel"],
  [
    ["SHEET_NAME", "Name des Tabellenblatts (muss exakt uebereinstimmen)", "Terminbuchung"],
    ["RESEND_API_KEY", "Ihr Resend API-Schluessel (siehe Kapitel 5)", "re_abc123..."],
    ["NOTIFY_EMAIL", "Ihre E-Mail fuer Buchungsbenachrichtigungen", "ihre@email.de"],
    ["FROM_EMAIL", "Absender-Adresse (muss bei Resend verifiziert sein)", "info@ihredomain.de"],
  ]
));

children.push(new Paragraph({ spacing: { after: 200 } }));

children.push(heading2("Script veroeffentlichen (Deploy)"));

children.push(step("steps3", normal("Klicken Sie im Script-Editor oben rechts auf "), bold("Bereitstellen"), normal(" > "), bold("Neue Bereitstellung"), normal(".")));
children.push(step("steps3", normal("Waehlen Sie als Typ: "), bold("Web-App"), normal(".")));
children.push(step("steps3", normal("Setzen Sie "), bold("Ausfuehren als"), normal(": "), bold("Ich"), normal(".")));
children.push(step("steps3", normal("Setzen Sie "), bold("Zugriff"), normal(": "), bold("Jeder"), normal(".")));
children.push(step("steps3", normal("Klicken Sie auf "), bold("Bereitstellen"), normal(".")));
children.push(step("steps3", normal("Kopieren Sie die angezeigte "), bold("URL"), normal(" \u2013 diese brauchen Sie im naechsten Schritt.")));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(tipBox("Wenn Sie den Code spaeter aendern, muessen Sie eine neue Bereitstellung erstellen oder die bestehende aktualisieren. Die URL bleibt bei einer Aktualisierung gleich."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 5. RESEND
// ============================================================
children.push(heading1("5. E-Mail-Versand mit Resend einrichten"));

children.push(p("Resend ist ein Dienst fuer den E-Mail-Versand. Damit werden die Bestaetigungsmails an Kunden und Benachrichtigungen an Sie verschickt."));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(step("steps4", normal("Gehen Sie zu "), bold("resend.com"), normal(" und erstellen Sie ein kostenloses Konto.")));
children.push(step("steps4", normal("Erstellen Sie unter "), bold("API Keys"), normal(" einen neuen API-Schluessel.")));
children.push(step("steps4", normal("Kopieren Sie den Schluessel und tragen Sie ihn in der Datei "), bold("Code.gs"), normal(" bei "), bold("RESEND_API_KEY"), normal(" ein.")));

children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(heading2("Eigene Domain verifizieren (empfohlen)"));

children.push(p("Damit E-Mails von Ihrer eigenen Adresse gesendet werden (z.B. info@ihrefirma.de), muessen Sie Ihre Domain bei Resend verifizieren:"));
children.push(new Paragraph({ spacing: { after: 80 } }));

children.push(step("steps5", normal("Gehen Sie in Resend auf "), bold("Domains"), normal(" > "), bold("Add Domain"), normal(".")));
children.push(step("steps5", normal("Folgen Sie den Anweisungen, um DNS-Eintraege bei Ihrem Domain-Anbieter zu setzen.")));
children.push(step("steps5", normal("Warten Sie, bis die Verifizierung abgeschlossen ist (kann bis zu 24 Stunden dauern).")));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(tipBox("Ohne eigene Domain koennen Sie die Test-Adresse von Resend verwenden. E-Mails gehen dann aber nur an Ihre eigene Adresse."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 6. FRONTEND ANPASSEN
// ============================================================
children.push(heading1("6. Frontend anpassen"));

children.push(p("Die Datei index.html ist die Buchungsseite, die Ihre Kunden sehen. Hier passen Sie Ihre Daten und Einstellungen an."));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(heading2("Apps Script URL eintragen"));

children.push(p("Oeffnen Sie die Datei index.html in einem Texteditor und suchen Sie die Zeile:"));
children.push(codeBlock("const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/.../exec';"));
children.push(p("Ersetzen Sie die URL durch Ihre eigene (aus Kapitel 4)."));

children.push(new Paragraph({ spacing: { after: 100 } }));
children.push(heading2("Persoenliche Daten aendern"));

children.push(p("Suchen Sie im HTML-Bereich folgende Stellen und passen Sie sie an:"));
children.push(new Paragraph({ spacing: { after: 80 } }));

children.push(settingsTable(
  ["Was", "Wo in der Datei", "Aendern in"],
  [
    ["Name", "<h1>Dr. Dirk Koetting</h1>", "Ihren Namen"],
    ["Untertitel", "<p>IT Governance & AI Policy...</p>", "Ihre Taetigkeit"],
    ["Logo", "assets/Logo.jpg", "Eigenes Logo in assets/ ablegen"],
    ["Favicon", "assets/favicon.png", "Eigenes Icon in assets/ ablegen"],
    ["Seitentitel", "<title>Dr. Dirk Koetting...</title>", "Ihren Seitentitel"],
  ]
));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(heading2("Zeitfenster aendern"));

children.push(p("Die angebotenen Zeitfenster stehen in dieser Zeile:"));
children.push(codeBlock("const FIXED_SLOTS = ['16:00','16:30','17:00'];"));
children.push(para(normal("Aendern Sie die Uhrzeiten nach Ihrem Wunsch, z.B. "), bold("['09:00','09:30','10:00','10:30']"), normal(" fuer vormittags.")));

children.push(new Paragraph({ spacing: { after: 100 } }));
children.push(heading2("Terminarten aendern"));

children.push(p("Die Terminarten finden Sie im Abschnitt meetingTypes:"));
children.push(codeBlock("{ id:'erstgespraech', name:'Erstgespraech', duration:30, color:'#21808d' }"));

children.push(settingsTable(
  ["Eigenschaft", "Bedeutung", "Beispiel"],
  [
    ["id", "Interner Name (keine Leerzeichen, keine Umlaute)", "workshop"],
    ["name", "Anzeigename fuer den Kunden", "Workshop-Termin"],
    ["duration", "Dauer in Minuten", "60"],
    ["color", "Farbcode (Hex)", "#ff6600"],
  ]
));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(heading2("Beratungsthemen anpassen"));

children.push(p("Die Themen im Dropdown-Menue finden Sie im HTML-Bereich:"));
children.push(codeBlock('<option value="AI Governance Beratung">AI Governance Beratung</option>'));
children.push(p("Fuegen Sie eigene Themen hinzu oder entfernen Sie vorhandene. Der value ist der interne Wert, der Text dazwischen ist das, was der Kunde sieht."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 7. SEITE ONLINE STELLEN
// ============================================================
children.push(heading1("7. Seite online stellen"));

children.push(p("Die Terminbuchung ist eine einfache HTML-Seite. Sie brauchen keinen speziellen Server. Es gibt mehrere kostenlose Moeglichkeiten:"));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(heading2("Option A: GitHub Pages (kostenlos)"));

children.push(bulletRuns("bullets2", normal("Laden Sie die Dateien in ein GitHub-Repository hoch.")));
children.push(bulletRuns("bullets2", normal("Gehen Sie zu "), bold("Settings"), normal(" > "), bold("Pages"), normal(" > waehlen Sie den Branch "), bold("main"), normal(".")));
children.push(bulletRuns("bullets2", normal("Die Seite ist dann unter "), bold("ihrname.github.io/reponame"), normal(" erreichbar.")));

children.push(new Paragraph({ spacing: { after: 100 } }));
children.push(heading2("Option B: Vercel oder Netlify (kostenlos)"));

children.push(bulletRuns("bullets3", normal("Erstellen Sie ein Konto bei vercel.com oder netlify.com.")));
children.push(bulletRuns("bullets3", normal("Verbinden Sie Ihr GitHub-Repository.")));
children.push(bulletRuns("bullets3", normal("Die Seite wird automatisch deployed und ist sofort erreichbar.")));
children.push(bulletRuns("bullets3", normal("Eigene Domain kann kostenlos angebunden werden.")));

children.push(new Paragraph({ spacing: { after: 100 } }));
children.push(heading2("Option C: Lokaler Test"));

children.push(p("Zum Testen auf Ihrem eigenen Rechner:"));
children.push(codeBlock("python -m http.server 8001 --bind 127.0.0.1"));
children.push(p("Dann im Browser http://127.0.0.1:8001 oeffnen."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 8. TAEGLICHE NUTZUNG
// ============================================================
children.push(heading1("8. Taegliche Nutzung"));

children.push(heading2("Admin-Panel"));

children.push(para(normal("Unten rechts auf der Buchungsseite finden Sie ein "), bold("Zahnrad-Symbol"), normal(". Klicken Sie darauf, um das Admin-Panel zu oeffnen.")));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(para(bold("Hier koennen Sie:")));
children.push(bullet("Einzelne Wochentage aktivieren oder deaktivieren"));
children.push(bullet("Start- und Endzeit pro Wochentag aendern"));
children.push(bullet("Alle bestehenden Buchungen aus dem Google Sheet einsehen"));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(tipBox("Aenderungen im Admin-Panel gelten nur fuer die aktuelle Browser-Sitzung. Nach einem Seiten-Reload sind die Standard-Einstellungen wieder aktiv. Fuer dauerhafte Aenderungen muessen die Werte direkt in der index.html angepasst werden (siehe Kapitel 9)."));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(heading2("Buchungen verwalten"));

children.push(p("Alle Buchungen werden automatisch im Google Sheet gespeichert. Dort koennen Sie:"));
children.push(bullet("Alle Buchungen sortieren und filtern"));
children.push(bullet("Buchungen loeschen (der Slot wird dann wieder frei)"));
children.push(bullet("Daten exportieren (z.B. als CSV)"));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(heading2("E-Mail-Ablauf"));

children.push(p("Bei jeder Buchung werden automatisch zwei E-Mails versendet:"));
children.push(new Paragraph({ spacing: { after: 80 } }));

children.push(settingsTable(
  ["E-Mail", "Empfaenger", "Inhalt"],
  [
    ["Benachrichtigung", "Sie (NOTIFY_EMAIL)", "Alle Buchungsdetails, Antwort-Adresse ist die des Kunden"],
    ["Bestaetigung", "Der Kunde", "Zusammenfassung mit Datum, Uhrzeit, Thema"],
  ]
));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 9. ANPASSUNGEN & EINSTELLUNGEN
// ============================================================
children.push(heading1("9. Anpassungen & Einstellungen"));

children.push(p("Hier eine Uebersicht aller Stellen, an denen Sie Einstellungen aendern koennen:"));
children.push(new Paragraph({ spacing: { after: 100 } }));

children.push(heading2("In der Datei index.html"));

children.push(settingsTable(
  ["Einstellung", "Standard-Wert", "Beschreibung"],
  [
    ["APPS_SCRIPT_URL", "https://script.google.com/...", "URL Ihres Google Apps Scripts"],
    ["SLOT_DURATION_MIN", "30", "Dauer eines Slots in Minuten"],
    ["FIXED_SLOTS", "['16:00','16:30','17:00']", "Angebotene Zeitfenster"],
    ["meetingTypes", "Erstgespraech, Beratungstermin", "Terminarten mit Name, Dauer, Farbe"],
    ["availability (Mo-Fr)", "16:00 - 17:30, aktiviert", "Welche Tage und Uhrzeiten verfuegbar sind"],
    ["availability (Sa-So)", "16:00 - 17:30, deaktiviert", "Wochenende ist standardmaessig aus"],
  ]
));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(heading2("In der Datei Code.gs"));

children.push(settingsTable(
  ["Einstellung", "Beschreibung"],
  [
    ["SHEET_NAME", "Name des Google-Sheet-Tabs (muss exakt uebereinstimmen)"],
    ["RESEND_API_KEY", "Ihr Resend API-Schluessel"],
    ["NOTIFY_EMAIL", "Ihre E-Mail fuer Benachrichtigungen"],
    ["FROM_EMAIL", "Absender-Adresse (muss bei Resend verifiziert sein)"],
  ]
));

children.push(new Paragraph({ spacing: { after: 200 } }));
children.push(heading2("Design"));

children.push(bullet("Dark Mode wird automatisch unterstuetzt (folgt den Systemeinstellungen des Besuchers)"));
children.push(bullet("Mobile-Ansicht ist eingebaut (responsive Design)"));
children.push(bullet("Farben koennen ueber die CSS-Variablen am Anfang der index.html geaendert werden"));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ============================================================
// 10. HAEUFIGE FRAGEN
// ============================================================
children.push(heading1("10. Haeufige Fragen"));

const faqs = [
  ["Kann ein Termin doppelt gebucht werden?",
   "Nein. Das Backend prueft vor jeder Buchung, ob der Slot bereits belegt ist. Zusaetzlich werden gebuchte Slots im Kalender sofort als belegt angezeigt."],
  ["Was passiert, wenn ich einen Termin im Google Sheet loesche?",
   "Der Slot wird beim naechsten Laden der Seite wieder als frei angezeigt. Die Seite gleicht sich beim Oeffnen automatisch mit dem Sheet ab."],
  ["Muss ich den Code aktualisieren, wenn ich Termine aendere?",
   "Nein. Termine werden komplett ueber das Google Sheet verwaltet. Nur fuer Aenderungen an den Zeitfenstern, Terminarten oder Themen muessen Sie die index.html anpassen."],
  ["Wie aendere ich die Verfuegbarkeit dauerhaft?",
   "Passen Sie die Werte im availability-Objekt in der index.html an. Das Admin-Panel aendert die Werte nur fuer die aktuelle Sitzung."],
  ["Funktioniert die Seite auf dem Handy?",
   "Ja. Die Seite passt sich automatisch an kleine Bildschirme an."],
  ["Kann ich mehrere Terminarten mit unterschiedlicher Dauer anbieten?",
   "Ja. Aendern Sie den duration-Wert bei der jeweiligen Terminart im meetingTypes-Array. Beachten Sie aber, dass die FIXED_SLOTS fuer alle Terminarten gleich sind."],
  ["Wie sicher ist das System?",
   "Die Buchungsdaten liegen in Ihrem persoenlichen Google Sheet. Der API-Schluessel fuer Resend liegt im Apps Script (nicht im Frontend). Das Frontend hat keinen direkten Zugriff auf sensible Daten."],
];

faqs.forEach(([q, a]) => {
  children.push(para(bold(q)));
  children.push(p(a));
  children.push(new Paragraph({ spacing: { after: 80 } }));
});

// ============================================================
// BUILD DOCUMENT
// ============================================================
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: TEAL },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: DARK },
        paragraph: { spacing: { before: 280, after: 160 }, outlineLevel: 1 }
      }
    ]
  },
  numbering: { config: numberingConfig },
  sections: [{
    properties: {
      page: {
        size: { width: PAGE_WIDTH, height: PAGE_HEIGHT },
        margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: "Terminbuchung \u2013 Anleitung", size: 18, font: "Arial", color: GRAY, italics: true })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Seite ", size: 18, font: "Arial", color: GRAY }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: GRAY })
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Terminbuchung_Anleitung.docx", buffer);
  console.log("OK: Terminbuchung_Anleitung.docx erstellt");
});
