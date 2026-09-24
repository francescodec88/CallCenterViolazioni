/**
 * CONFIGURAZIONE GLOBALE AZIENDALE
 * Architettura a Template (Senza LockService - Massima Velocità)
 */
const CONFIG = {
  // ⚠️ ID DEL FILE SPREADSHEET (Condiviso con chiunque abbia il link)
  TEMPLATE_FILE_ID: '1zkuXIu1BHHMpmMMboGA0-YfR0VSm-yjailP6ZdFitmg', 
  
  APP_FOLDER_NAME: 'CallCenter_Files', 
  NOME_FILE_DB_FINALE: 'Database_Call_Center',
  
  FOGLIO_DB: 'Master_DB',
  FOGLIO_IMPOSTAZIONI: 'Impostazioni',
  FOGLIO_MANUALE: 'Manuale_Operativo',
  
  MAX_RISULTATI_RICERCA: 300
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaccia')
      .setTitle('Call Center | Violazioni')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * CORE: Gestisce l'intero ecosistema dell'utente.
 * Se l'utente è nuovo, crea la cartella e copia il tuo template.
 */
function getDbFile() {
  const userProps = PropertiesService.getUserProperties();
  let dbId = userProps.getProperty('USER_DB_ID');
  
  if (dbId) {
    try { 
      return SpreadsheetApp.openById(dbId); 
    } catch(e) { 
      userProps.deleteProperty('USER_DB_ID'); 
      dbId = null; 
    }
  }
  
  // 2. L'UTENTE E' NUOVO O MANCA IL FILE: Creiamo l'ambiente
  const folders = DriveApp.getFoldersByName(CONFIG.APP_FOLDER_NAME);
  let userFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.APP_FOLDER_NAME);
  
  const templateDB = DriveApp.getFileById(CONFIG.TEMPLATE_FILE_ID);
  const newDb = templateDB.makeCopy(CONFIG.NOME_FILE_DB_FINALE, userFolder);
  
  dbId = newDb.getId();
  userProps.setProperty('USER_DB_ID', dbId);
  return SpreadsheetApp.openById(dbId);
}

const getSheet = (nomeFoglio) => {
  const ss = getDbFile();
  const sheet = ss.getSheetByName(nomeFoglio);
  if (!sheet) throw new Error(`Errore: Foglio '${nomeFoglio}' mancante nel template.`);
  return sheet;
};

const getOggi = (timeZone) => Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");

const pulisciDatiInIngresso = (dati) => {
  return {
    nominativo: dati.nominativo ? String(dati.nominativo).trim().toUpperCase().replace(/\s+/g, ' ') : "",
    telefono: dati.telefono ? String(dati.telefono).trim() : "",
    targa: dati.targa ? String(dati.targa).trim().toUpperCase().replace(/\s+/g, '') : "",
    motivo: dati.motivo ? String(dati.motivo).trim() : "",
    note: dati.note ? String(dati.note).trim() : ""
  };
};

function verificaEInizializzaOperatore(datiNuovi) {
  const sheet = getSheet('DATI OPERATORE');
  
  if (datiNuovi && datiNuovi.cognome) {
    const c = String(datiNuovi.cognome).trim().toUpperCase();
    const n = String(datiNuovi.nome).trim().toUpperCase();
    const m = String(datiNuovi.matricola).trim().toUpperCase();
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, 1, 3).setValues([[c, n, m]]);
    } else {
      sheet.appendRow([c, n, m]);
    }
  }

  if (sheet.getLastRow() > 1) {
    const dati = sheet.getRange(2, 1, 1, 3).getValues()[0];
    const cognome = dati[0] || "";
    const nomeIntero = dati[1] || "";
    const inizialeNome = nomeIntero ? nomeIntero.charAt(0) + "." : "";
    return { registrato: true, cognome: cognome, matricola: dati[2] || "", nomeFormat: `${cognome} ${inizialeNome}`.trim() };
  }
  return { registrato: false };
}

/**
 * SALVATAGGIO ISTANTANEO (Senza LockService perché ogni utente ha il suo file)
 */
function salvaChiamata(dati) {
  const fDati = pulisciDatiInIngresso(dati);
  const sheet = getSheet(CONFIG.FOGLIO_DB);
  const dataOdierna = new Date();
  
  sheet.appendRow([dataOdierna, fDati.nominativo, fDati.targa, fDati.motivo, fDati.note, fDati.telefono]);
  
  const timeZone = Session.getScriptTimeZone();
  return {
    row: sheet.getLastRow(),
    ora: Utilities.formatDate(dataOdierna, timeZone, "HH:mm"),
    dataCompleta: Utilities.formatDate(dataOdierna, timeZone, "dd/MM/yyyy HH:mm")
  };
}

function aggiornaChiamata(dati) {
  const row = parseInt(dati.row, 10);
  const fDati = pulisciDatiInIngresso(dati);
  const sheet = getSheet(CONFIG.FOGLIO_DB);
  
  sheet.getRange(row, 2, 1, 5).setValues([[fDati.nominativo, fDati.targa, fDati.motivo, fDati.note, fDati.telefono]]);
  return { row, ...fDati };
}

function caricaSessioneOggi() {
  const ss = getDbFile();
  const infoOperatore = verificaEInizializzaOperatore();
  
  let scorciatoieMenu = [], scorciatoieBottoni = [];
  try {
    const datiSc = ss.getSheetByName('Scorciatoie_Web').getDataRange().getValues();
    datiSc.slice(1).forEach(row => {
      const tipo = String(row[0]).trim().toUpperCase();
      const nome = String(row[1]).trim();
      const link = String(row[2]).trim();
      if (tipo === 'PULSANTE' && nome && link) scorciatoieBottoni.push({ nome, link });
      else if (tipo === 'MENU' && nome && link) scorciatoieMenu.push({ nome, link });
    });
  } catch(e) { console.warn("Foglio scorciatoie mancante o errato."); }

  const dati = getSheet(CONFIG.FOGLIO_DB).getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const oggiStr = getOggi(timeZone); 
  const elencoOggi = [];
  
  for (let i = dati.length - 1; i >= 1; i--) {
    const rawDate = dati[i][0];
    if (!rawDate) continue;
    const dataObj = new Date(rawDate);
    if (isNaN(dataObj.getTime())) continue;

    if (Utilities.formatDate(dataObj, timeZone, "yyyy-MM-dd") === oggiStr) { 
      elencoOggi.push({
        row: i + 1, ora: Utilities.formatDate(dataObj, timeZone, "HH:mm"), data: Utilities.formatDate(dataObj, timeZone, "dd/MM/yyyy"),
        nominativo: String(dati[i][1] || ''), targa: String(dati[i][2] || '').toUpperCase(),
        motivo: String(dati[i][3] || ''), note: String(dati[i][4] || ''), telefono: String(dati[i][5] || '')
      });
    }
  }

  let opzioniMotivo = [], tagRapidi = [];
  try {
    const imp = ss.getSheetByName(CONFIG.FOGLIO_IMPOSTAZIONI);
    const datiImp = imp.getDataRange().getValues();
    datiImp.slice(1).forEach(row => {
      if (row[0]) opzioniMotivo.push(String(row[0]).trim());
      if (row[1]) tagRapidi.push(String(row[1]).trim());
    });
  } catch (e) {}

  return { 
    conteggio: elencoOggi.length, 
    lista: elencoOggi, 
    motivi: opzioniMotivo, 
    tags: tagRapidi, 
    operatore: infoOperatore,
    scorciatoieMenu: scorciatoieMenu,
    scorciatoieBottoni: scorciatoieBottoni
  };
}

function caricaManualeOperativo() {
  try {
    const sheet = getSheet(CONFIG.FOGLIO_MANUALE);
    const dati = sheet.getDataRange().getValues();
    const manuale = [];
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][1]) {
        manuale.push({ id: String(dati[i][0]), titolo: String(dati[i][1]), contenuto: String(dati[i][2]) });
      }
    }
    if (manuale.length === 0) return [{ id: 'err', titolo: "Nessun Manuale", contenuto: "<p>Nessun manuale salvato nel database.</p>" }];
    return manuale;
  } catch (e) {
    return [{ id: 'err', titolo: "Errore Lettura", contenuto: `<p>Impossibile leggere il manuale dal database interno.</p>` }];
  }
}

function ricercaStorica(params) {
  const dati = getSheet(CONFIG.FOGLIO_DB).getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const query = (params.query || '').toLowerCase().trim();
  const { dataDa, dataA } = params;
  const risultati = [];

  for (let i = dati.length - 1; i >= 1; i--) { 
    if (!dati[i][0]) continue;
    const dataObj = new Date(dati[i][0]);
    if (isNaN(dataObj.getTime())) continue;
    const rowDateStr = Utilities.formatDate(dataObj, timeZone, "yyyy-MM-dd");

    if (dataDa && rowDateStr < dataDa) continue;
    if (dataA && rowDateStr > dataA) continue;
    
    if (query) {
       const rowString = `${dati[i][2]||''} ${dati[i][1]||''} ${dati[i][5]||''} ${dati[i][3]||''} ${dati[i][4]||''}`.toLowerCase();
       if (!rowString.includes(query)) continue;
    }
    risultati.push({
      row: i + 1, data: Utilities.formatDate(dataObj, timeZone, "dd/MM/yyyy"), ora: Utilities.formatDate(dataObj, timeZone, "HH:mm"),
      nominativo: String(dati[i][1] || ''), targa: String(dati[i][2] || '').toUpperCase(),
      motivo: String(dati[i][3] || ''), note: String(dati[i][4] || ''), telefono: String(dati[i][5] || '')
    });
    if (risultati.length >= CONFIG.MAX_RISULTATI_RICERCA) break;
  }
  return { risultati, troncato: risultati.length >= CONFIG.MAX_RISULTATI_RICERCA };
}

function gestisciReportMailDaBase64(azione, pdfBase64, totale, operatore, matricola) {
  const ss = getDbFile();
  const sheet = ss.getSheetByName('Template_Mail');
  
  let destinatari = Session.getActiveUser().getEmail();
  let oggetto = 'Report Chiamate del {{DATA}} - Operatore: {{OPERATORE}}';
  let corpo = 'Buongiorno,\n\nin allegato il report di fine turno.\nOggi sono state gestite {{TOTALE}} chiamate.\n\nSaluti,\nOperatore: {{OPERATORE}}';

  if (sheet && sheet.getLastRow() > 1) {
    const dati = sheet.getRange(2, 1, 1, 3).getValues()[0];
    if (dati[0]) destinatari = String(dati[0]).trim();
    if (dati[1]) oggetto = String(dati[1]);
    if (dati[2]) corpo = String(dati[2]);
  }
  
  const dataOdierna = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  const sostituisciDati = (testo) => testo
    .replace(/{{DATA}}/g, dataOdierna)
    .replace(/{{OPERATORE}}/g, operatore)
    .replace(/{{TOTALE}}/g, totale)
    .replace(/{{MATRICOLA}}/g, matricola);
/**
 * ==========================================
 * CONFIGURAZIONE AZIENDALE E STARTUP
 * ==========================================
 */
const CONFIG = {
  TEMPLATE_FILE_ID: '1zkuXIu1BHHMpmMMboGA0-YfR0VSm-yjailP6ZdFitmg', 
  APP_FOLDER_NAME: 'CallCenter_Files', 
  NOME_FILE_DB_FINALE: 'Database_Call_Center',
  FOGLIO_DB: 'Master_DB',
  FOGLIO_IMPOSTAZIONI: 'Impostazioni',
  MAX_RISULTATI_RICERCA: 300
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interfaccia')
      .setTitle('Call Center | Violazioni')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * GESTIONE FILE PRIVATO DELL'UTENTE (Senza colli di bottiglia)
 */
function getDbFile() {
  const userProps = PropertiesService.getUserProperties();
  let dbId = userProps.getProperty('USER_DB_ID');
  
  // 1. Controllo se l'operatore ha già un file registrato
  if (dbId) {
    try { 
      return SpreadsheetApp.openById(dbId); 
    } catch(e) { 
      userProps.deleteProperty('USER_DB_ID'); 
      dbId = null; 
    }
  }
  
  // 2. Se è il primo avvio, cerca o crea la cartella su Drive
  const folders = DriveApp.getFoldersByName(CONFIG.APP_FOLDER_NAME);
  let userFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(CONFIG.APP_FOLDER_NAME);
  
  // 3. Clona il Template Vergine in modo istantaneo
  const templateDB = DriveApp.getFileById(CONFIG.TEMPLATE_FILE_ID);
  const newDb = templateDB.makeCopy(CONFIG.NOME_FILE_DB_FINALE, userFolder);
  
  dbId = newDb.getId();
  userProps.setProperty('USER_DB_ID', dbId);
  return SpreadsheetApp.openById(dbId);
}

const getSheet = (nomeFoglio) => {
  const ss = getDbFile();
  const sheet = ss.getSheetByName(nomeFoglio);
  if (!sheet) throw new Error(`Errore: Foglio '${nomeFoglio}' mancante nel template.`);
  return sheet;
};

const pulisciDatiInIngresso = (dati) => {
  return {
    nominativo: dati.nominativo ? String(dati.nominativo).trim().toUpperCase().replace(/\s+/g, ' ') : "",
    telefono: dati.telefono ? String(dati.telefono).trim() : "",
    targa: dati.targa ? String(dati.targa).trim().toUpperCase().replace(/\s+/g, '') : "",
    motivo: dati.motivo ? String(dati.motivo).trim() : "",
    note: dati.note ? String(dati.note).trim() : ""
  };
};

/**
 * ==========================================
 * MODULO AUTOINSTALLAZIONE E DATI OPERATORE
 * ==========================================
 */
function verificaEInizializzaOperatore(datiNuovi) {
  const ss = getDbFile();
  let sheet = ss.getSheetByName('DATI OPERATORE');
  
  if (!sheet) {
    sheet = ss.insertSheet('DATI OPERATORE');
    sheet.hideSheet();
    sheet.appendRow(['COGNOME', 'NOME', 'MATRICOLA']);
    sheet.getRange("A1:C1").setFontWeight("bold");
  }

  if (datiNuovi && datiNuovi.cognome) {
    const c = String(datiNuovi.cognome).trim().toUpperCase();
    const n = String(datiNuovi.nome).trim().toUpperCase();
    const m = String(datiNuovi.matricola).trim().toUpperCase();
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, 1, 3).setValues([[c, n, m]]);
    } else {
      sheet.appendRow([c, n, m]);
    }
  }

  if (sheet.getLastRow() > 1) {
    const dati = sheet.getRange(2, 1, 1, 3).getValues()[0];
    const cognome = dati[0] || "";
    const nomeIntero = dati[1] || "";
    const inizialeNome = nomeIntero ? nomeIntero.charAt(0) + "." : "";
    return { 
      registrato: true, 
      cognome: cognome, 
      matricola: dati[2] || "", 
      nomeFormat: `${cognome} ${inizialeNome}`.trim() 
    };
  }
  return { registrato: false };
}

/**
 * ==========================================
 * GESTIONE CHIAMATE E STORICO
 * ==========================================
 */
function salvaChiamata(dati) {
  const fDati = pulisciDatiInIngresso(dati);
  const sheet = getSheet(CONFIG.FOGLIO_DB);
  const dataOdierna = new Date();
  
  // Inserimento immediato senza LockService
  sheet.appendRow([dataOdierna, fDati.nominativo, fDati.targa, fDati.motivo, fDati.note, fDati.telefono]);
  
  const timeZone = Session.getScriptTimeZone();
  return {
    row: sheet.getLastRow(),
    ora: Utilities.formatDate(dataOdierna, timeZone, "HH:mm"),
    dataCompleta: Utilities.formatDate(dataOdierna, timeZone, "dd/MM/yyyy HH:mm")
  };
}

function aggiornaChiamata(dati) {
  const row = parseInt(dati.row, 10);
  const fDati = pulisciDatiInIngresso(dati);
  const sheet = getSheet(CONFIG.FOGLIO_DB);
  
  sheet.getRange(row, 2, 1, 5).setValues([[fDati.nominativo, fDati.targa, fDati.motivo, fDati.note, fDati.telefono]]);
  return { row, ...fDati };
}

function caricaSessioneOggi() {
  const ss = getDbFile();
  const infoOperatore = verificaEInizializzaOperatore();
  
  let scorciatoieMenu = [], scorciatoieBottoni = [];
  try {
    const datiSc = ss.getSheetByName('Scorciatoie_Web').getDataRange().getValues();
    datiSc.slice(1).forEach(row => {
      const tipo = String(row[0]).trim().toUpperCase();
      const nome = String(row[1]).trim();
      const link = String(row[2]).trim();
      if (tipo === 'PULSANTE' && nome && link) scorciatoieBottoni.push({ nome, link });
      else if (tipo === 'MENU' && nome && link) scorciatoieMenu.push({ nome, link });
    });
  } catch(e) {}

  const dati = getSheet(CONFIG.FOGLIO_DB).getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const oggiStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd"); 
  const elencoOggi = [];
  
  for (let i = dati.length - 1; i >= 1; i--) {
    const rawDate = dati[i][0];
    if (!rawDate) continue;
    const dataObj = new Date(rawDate);
    if (isNaN(dataObj.getTime())) continue;

    if (Utilities.formatDate(dataObj, timeZone, "yyyy-MM-dd") === oggiStr) { 
      elencoOggi.push({
        row: i + 1, ora: Utilities.formatDate(dataObj, timeZone, "HH:mm"), data: Utilities.formatDate(dataObj, timeZone, "dd/MM/yyyy"),
        nominativo: String(dati[i][1] || ''), targa: String(dati[i][2] || '').toUpperCase(),
        motivo: String(dati[i][3] || ''), note: String(dati[i][4] || ''), telefono: String(dati[i][5] || '')
      });
    }
  }

  let opzioniMotivo = [], tagRapidi = [];
  try {
    const imp = ss.getSheetByName(CONFIG.FOGLIO_IMPOSTAZIONI);
    const datiImp = imp.getDataRange().getValues();
    datiImp.slice(1).forEach(row => {
      if (row[0]) opzioniMotivo.push(String(row[0]).trim());
      if (row[1]) tagRapidi.push(String(row[1]).trim());
    });
  } catch (e) {}

  return { 
    conteggio: elencoOggi.length, 
    lista: elencoOggi, 
    motivi: opzioniMotivo, 
    tags: tagRapidi, 
    operatore: infoOperatore,
    scorciatoieMenu: scorciatoieMenu,
    scorciatoieBottoni: scorciatoieBottoni
  };
}

function ricercaStorica(params) {
  const dati = getSheet(CONFIG.FOGLIO_DB).getDataRange().getValues();
  const timeZone = Session.getScriptTimeZone();
  const query = (params.query || '').toLowerCase().trim();
  const { dataDa, dataA } = params;
  const risultati = [];

  for (let i = dati.length - 1; i >= 1; i--) { 
    if (!dati[i][0]) continue;
    const dataObj = new Date(dati[i][0]);
    if (isNaN(dataObj.getTime())) continue;
    const rowDateStr = Utilities.formatDate(dataObj, timeZone, "yyyy-MM-dd");

    if (dataDa && rowDateStr < dataDa) continue;
    if (dataA && rowDateStr > dataA) continue;
    
    if (query) {
       const rowString = `${dati[i][2]||''} ${dati[i][1]||''} ${dati[i][5]||''} ${dati[i][3]||''} ${dati[i][4]||''}`.toLowerCase();
       if (!rowString.includes(query)) continue;
    }
    risultati.push({
      row: i + 1, data: Utilities.formatDate(dataObj, timeZone, "dd/MM/yyyy"), ora: Utilities.formatDate(dataObj, timeZone, "HH:mm"),
      nominativo: String(dati[i][1] || ''), targa: String(dati[i][2] || '').toUpperCase(),
      motivo: String(dati[i][3] || ''), note: String(dati[i][4] || ''), telefono: String(dati[i][5] || '')
    });
    if (risultati.length >= CONFIG.MAX_RISULTATI_RICERCA) break;
  }
  return { risultati, troncato: risultati.length >= CONFIG.MAX_RISULTATI_RICERCA };
}

/**
 * ==========================================
 * MODULO EMAIL (Lettura Template)
 * ==========================================
 */
function gestisciReportMailDaBase64(azione, pdfBase64, totale, operatore, matricola) {
  const ss = getDbFile();
  let sheet = ss.getSheetByName('Template_Mail');
  
  if (!sheet) {
    sheet = ss.insertSheet('Template_Mail');
    sheet.appendRow(['DESTINATARI', 'OGGETTO', 'CORPO EMAIL']);
    sheet.appendRow([
      Session.getActiveUser().getEmail(), 
      'Report Chiamate del {{DATA}} - Operatore: {{OPERATORE}}', 
      'Buongiorno,\n\nin allegato il report di fine turno.\nOggi sono state gestite {{TOTALE}} chiamate.\n\nSaluti,\nOperatore: {{OPERATORE}}'
    ]);
  }
  
  let destinatari = Session.getActiveUser().getEmail();
  let oggetto = 'Report'; let corpo = '';

  if (sheet && sheet.getLastRow() > 1) {
    const dati = sheet.getRange(2, 1, 1, 3).getValues()[0];
    if (dati[0]) destinatari = String(dati[0]).trim();
    if (dati[1]) oggetto = String(dati[1]);
    if (dati[2]) corpo = String(dati[2]);
  }
  
  const dataOdierna = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  const sostituisciDati = (testo) => testo
    .replace(/{{DATA}}/g, dataOdierna)
    .replace(/{{OPERATORE}}/g, operatore)
    .replace(/{{TOTALE}}/g, totale)
    .replace(/{{MATRICOLA}}/g, matricola);

  oggetto = sostituisciDati(oggetto);
  corpo = sostituisciDati(corpo);
  
  const splitBase = pdfBase64.split(',');
  const base64Data = splitBase[1] ? splitBase[1] : splitBase[0];
  const nomeFile = `Report_${operatore.replace(/\s+/g,'_')}_${dataOdierna.replace(/\//g,'-')}.pdf`;
  const blobPDF = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', nomeFile);
  
  if (azione === 'invia') {
    GmailApp.sendEmail(destinatari, oggetto, corpo, { attachments: [blobPDF] });
    return "Email inviata con successo!";
  } else if (azione === 'bozza') {
    GmailApp.createDraft(destinatari, oggetto, corpo, { attachments: [blobPDF] });
    return "Bozza creata con successo in Gmail!";
  } else {
    throw new Error("Azione non riconosciuta.");
  }
}
  oggetto = sostituisciDati(oggetto);
  corpo = sostituisciDati(corpo);
  
  const splitBase = pdfBase64.split(',');
  const base64Data = splitBase[1] ? splitBase[1] : splitBase[0];
  const nomeFile = `Report_${operatore.replace(/\s+/g,'_')}_${dataOdierna.replace(/\//g,'-')}.pdf`;
  const blobPDF = Utilities.newBlob(Utilities.base64Decode(base64Data), 'application/pdf', nomeFile);
  
  if (azione === 'invia') {
    GmailApp.sendEmail(destinatari, oggetto, corpo, { attachments: [blobPDF] });
    return "Email inviata con successo a " + destinatari;
  } else if (azione === 'bozza') {
    GmailApp.createDraft(destinatari, oggetto, corpo, { attachments: [blobPDF] });
    return "Bozza creata con successo in Gmail!";
  } else {
    throw new Error("Azione non riconosciuta (usa 'invia' o 'bozza').");
  }
}
