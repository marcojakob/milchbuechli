// *****************************************************
// Dieses Script muss in den Scripteditor vom 
// Google Sheet reinkopiert werden! 
// *****************************************************

const timeZone = 'Europe/Zurich';
const locale = 'de_CH';
const cellDateFormat = 'dd.MM.yyyy';

const kategorienSheetName = 'Kategorien';

const einnahmenSheetName = 'Einnahmen';
const einnahmenTableRowStart = 9;
const einnahmenTableNumRows = 1000;
const einnahmenTableColumnStart = 2;
const einnahmenTableNumColumns = 4;

const ausgabenSheetName = 'Ausgaben';
const ausgabenTableRowStart = 9;
const ausgabenTableNumRows = 1000;
const ausgabenTableColumnStart = 2;
const ausgabenTableNumColumns = 4;

/**
 * Wird ausgeführt, wenn das Google Sheet geöffnet wird.
 */
function onOpen() {
  // Menu hinzufügen.
  SpreadsheetApp.getUi().createMenu('🥛 Milchbüechli')
    .addItem('Einnahmen aus Dateinamen auslesen', 'einnahmenAusDateinamen')
    .addItem('Ausgaben aus Dateinamen auslesen', 'ausgabenAusDateinamen')
    .addSeparator()
    .addItem('Einnahmen Tabelle leeren', 'einnahmenTabelleLeeren')
    .addItem('Ausgaben Tabelle leeren', 'ausgabenTabelleLeeren')
    .addToUi();

  // Locale setzen.
  // const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  // spreadsheet.setSpreadsheetLocale(locale);
}

/** Einnahmen-Tabelle leeren. */
function einnahmenTabelleLeeren() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(einnahmenSheetName);
  sheet.getRange(einnahmenTableRowStart, einnahmenTableColumnStart, einnahmenTableNumRows, einnahmenTableNumColumns)
    .clearContent();
}

/** Ausgaben-Tabelle leeren. */
function ausgabenTabelleLeeren() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(ausgabenSheetName);
  sheet.getRange(ausgabenTableRowStart, ausgabenTableColumnStart, ausgabenTableNumRows, ausgabenTableNumColumns)
    .clearContent();
}

/**
 * Liest alle Angaben zu EINNAHMEN aus den Dateinamen im entsprechenden Ordner.
 * 
 * Dateinamen Beispiel: 2021-02-01 KUNDE A - Workshop CHF500 BEZ2021-02-20.pdf
 * Der Teil mit BEZ ist optional.
 */
function einnahmenAusDateinamen() {
  const filenamePattern = /^(\S+)\s+(.+)\sCHF(\S+)(\sBEZ(\S+))?\.\S+$/;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(einnahmenSheetName);

  // Ordner finden.
  const folder = getSubfolderByName(einnahmenSheetName);
  if (folder == null) {
    Browser.msgBox('Kein Unterordner mit dem Namen "' + einnahmenSheetName + '" gefunden.');
    return;
  }

  const files = getFilesAlphabetically(folder);

  // Datumsformat setzen.
  sheet.getRange(einnahmenTableRowStart, einnahmenTableColumnStart, files.length, 2).setNumberFormat(cellDateFormat);

  const linkTextStyle = SpreadsheetApp.newTextStyle()
    .setUnderline(false)
    .setForegroundColor("#000000")
    .build()

  let currentRow = einnahmenTableRowStart;

  for (let file of files) {
    const currentRange = sheet.getRange(currentRow, einnahmenTableColumnStart, 1, einnahmenTableNumColumns);

    const matches = filenamePattern.exec(file.getName());

    if (matches !== null && matches.length === 6) {
      const valueRechnungsdatum = SpreadsheetApp.newRichTextValue()
        .setText(matches[1])
        .setLinkUrl(file.getUrl())
        .setTextStyle(linkTextStyle)
        .build();
      const valueBeschreibung = SpreadsheetApp.newRichTextValue()
        .setText(matches[2])
        .setLinkUrl(file.getUrl())
        .setTextStyle(linkTextStyle)
        .build();
      const valueBetrag = SpreadsheetApp.newRichTextValue()
        .setText(matches[3])
        .build();
      const valueZahlungsdatum = SpreadsheetApp.newRichTextValue()
        .setText(matches[5] ?? '')
        .setLinkUrl(file.getUrl())
        .setTextStyle(linkTextStyle)
        .build();
      currentRange.setRichTextValues([[valueRechnungsdatum, valueZahlungsdatum, valueBeschreibung, valueBetrag]]);
    } else {
      // Filname is not correct.
      const fehlermeldung = SpreadsheetApp.newRichTextValue()
        .setText('Fehler im Dateinamen: ' + file.getName())
        .setLinkUrl(file.getUrl())
        .setTextStyle(linkTextStyle)
        .build();
      currentRange.setValue('');
      currentRange.getCell(1, 3).setRichTextValue(fehlermeldung);
    }

    currentRow++;
  }
}

/**
 * Liest alle Angaben zu AUSGABEN aus den Dateinamen im entsprechenden Ordner.
 * 
 * Dateinamen Beispiel: 2021-02-01 Website Hosting CHF500 KAT9.pdf
 */
function ausgabenAusDateinamen() {
  const filenamePattern = /^(\S+)\s(.+)\sCHF(\S+)(\sKAT(\S+))?\.\S+$/;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(ausgabenSheetName);

  const sheetKategorien = spreadsheet.getSheetByName(kategorienSheetName);
  const kategorien = sheetKategorien.getRange(8, 2, 24, 1).getValues().flat();

  // Ordner finden.
  const folder = getSubfolderByName(ausgabenSheetName);
  if (folder == null) {
    Browser.msgBox('Kein Unterordner mit dem Namen "' + ausgabenSheetName + '" gefunden.');
    return;
  }

  const files = getFilesAlphabetically(folder);

  // Sheet Range 

  // Datumsformat setzen.
  sheet.getRange(ausgabenTableRowStart, ausgabenTableColumnStart, files.length, 1).setNumberFormat(cellDateFormat);

  const linkTextStyle = SpreadsheetApp.newTextStyle()
    .setUnderline(false)
    .setForegroundColor("#000000")
    .build()

  let currentRow = ausgabenTableRowStart;

  for (let file of files) {
    const currentRange = sheet.getRange(currentRow, ausgabenTableColumnStart, 1, ausgabenTableNumColumns);

    const matches = filenamePattern.exec(file.getName());

    if (matches !== null && matches.length === 6) {
      const valueDatum = SpreadsheetApp.newRichTextValue()
        .setText(matches[1])
        .setLinkUrl(file.getUrl())
        .setTextStyle(linkTextStyle)
        .build();
      const valueBeschreibung = SpreadsheetApp.newRichTextValue()
        .setText(matches[2])
        .setLinkUrl(file.getUrl())
        .setTextStyle(linkTextStyle)
        .build();
      const valueBetrag = SpreadsheetApp.newRichTextValue()
        .setText(matches[3])
        .build();

      // Kategorie nachschlagen, wenn vorhanden.
      let kategorieText = '';
      const kategorieNumber = parseInt(matches[5]);
      if (!isNaN(kategorieNumber)) {
        kategorieText = kategorien[kategorieNumber - 1];
      }
      const valueKategorie = SpreadsheetApp.newRichTextValue()
        .setText(kategorieText)
        .setTextStyle(linkTextStyle)
        .build();

      currentRange.setRichTextValues([[valueDatum, valueBeschreibung, valueKategorie, valueBetrag]]);
    } else {
      // Filname is not correct.
      const fehlermeldung = SpreadsheetApp.newRichTextValue()
        .setText('Fehler im Dateinamen: ' + file.getName())
        .setLinkUrl(file.getUrl())
        .setTextStyle(linkTextStyle)
        .build();
      currentRange.setValue('');
      currentRange.getCell(1, 2).setRichTextValue(fehlermeldung);
    }

    currentRow++;
  }
}

/**
 * Sucht den Unterordner des Google Sheets mit entsprechendem Namen.
 * Falls kein solcher Unterordner gefunden werden kann, so wird null zurückgegeben.
 */
function getSubfolderByName(folderName) {
  let spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  let spreadsheetFile = DriveApp.getFileById(spreadsheetId);
  let mainFolder = spreadsheetFile.getParents().next();
  let folderIterator = mainFolder.getFoldersByName(folderName);

  if (folderIterator.hasNext()) {
    return folderIterator.next();
  } else {
    return null;
  }
}

/**
 * Returns a list of files in the specified folder sorted alphabetically.
 */
function getFilesAlphabetically(folder) {
  const filesIterator = folder.getFiles();

  let files = [];

  // creates an array of file objects
  while (filesIterator.hasNext()) {
    files.push(filesIterator.next());
  }

  // sorts the files array by file names alphabetically
  files = files.sort(function (a, b) {
    let aName = a.getName().toLowerCase();
    let bName = b.getName().toLowerCase();

    return aName.localeCompare(bName);
  });

  return files;
}