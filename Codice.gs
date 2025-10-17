/*
 * Restituisce il prossimo ID disponibile incrementando quello più alto presente in una colonna.
 * @param {Sheet} sheet - Il foglio su cui cercare.
 * @param {number} colonna - Numero della colonna (1 = colonna A).
 * @param {string} prefisso - Prefisso dell'ID (es. "EG", "S", "P", "C", "I").
 * @returns {string} - Nuovo ID formattato (es. EG001).
*/

function getNextId(sheet, colonna, prefisso) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return prefisso + "001"; // nessun dato, inizia da 001

  const ids = sheet.getRange(2, colonna, lastRow - 1).getValues().flat().filter(String); // togli celle vuote
  const maxIdNum = ids.length > 0
    ? Math.max(...ids.map(id => parseInt(id.replace(prefisso, ""))))
    : 0;

  return prefisso + Utilities.formatString("%03d", maxIdNum + 1);
}

/**
 * Registra un nuovo anno scout nel foglio "Registro anni".
 * Aggiorna la data di fine dell'anno precedente, se presente.
 * @param {Object} form - Oggetto contenente la data di inizio del nuovo anno.
 * @param {Date} form.inizio - La data di inizio del nuovo anno.
 */
function addAnno(form) {
  console.log("Anno inizio ricevuto: ", form.inizio)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "Registro anni";
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log(`Errore: Foglio "${sheetName}" non trovato.`);
    throw new Error(`Foglio "${sheetName}" non trovato.`);
  }

  const dataInizioNuovoAnno = form.inizio;
  const dataOraInserimento = new Date();

  const lastRow = sheet.getLastRow();  
  if (lastRow > 1) {
    // Ci sono righe di anni precedenti (l'ultima riga è la riga lastRow)
    
    // 2. Calcola la data di fine per l'anno precedente (1 secondo prima dell'inizio del nuovo)
    //const dataFineAnnoPrecedente = new Date(dataInizioNuovoAnno.getTime() - 1000);
    const dataFineAnnoPrecedente = dataInizioNuovoAnno
    
    // Aggiorna la colonna "Fine" (colonna C) dell'ultima riga
    sheet.getRange(lastRow, 3).setValue(dataFineAnnoPrecedente);
    
    Logger.log(`Aggiornata riga ${lastRow}: Fine anno precedente impostata a ${dataFineAnnoPrecedente}`);
  }

  // 3. Aggiunge la nuova riga
  const newId = getNextId(sheet, 1, "A"); // ID_anno è la colonna 1 (A)
  
  // Per la colonna "Fine" (C), inseriamo la stringa "oggi" come richiesto
  const nuovaRiga = [
    newId,                          // ID_anno (Colonna A)
    dataInizioNuovoAnno,            // Inizio (Colonna B)
    "oggi",                         // Fine (Colonna C) - come stringa per indicare l'anno in corso
    dataOraInserimento              // Data_Inserimento (Colonna D)
  ];
  
  sheet.appendRow(nuovaRiga);
  Logger.log(`Aggiunto nuovo anno: ${newId} - Inizio: ${dataInizioNuovoAnno}`);
}

function addRagazzo(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ragazzi");
  const newId = getNextId(sheet, 1, "EG");

  const dataOraInserimento = new Date();
  sheet.appendRow([newId, form.nome, form.cognome, dataOraInserimento]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 4).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function removeRagazzo(form) {
  if (!form || !form.ragazzo) {
    throw new Error("ID ragazzo non ricevuto dal form");
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ragazzi");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("Nessun ragazzo presente");

  // Leggo ID (col A)
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  // Cerco la riga corrispondente all'ID
  const rowIndex = data.findIndex(r => r[0].toString().trim() === form.ragazzo.toString().trim());

  if (rowIndex === -1) {
    throw new Error(`Ragazzo non trovato: ID ${form.ragazzo}`);
  }

  // Scrivo "uscito" nella quinta colonna
  const row = rowIndex + 2; // +2 perché i dati partono dalla riga 2
  sheet.getRange(row, 5).setValue("uscito");
}

function addInfoRagazzo(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Info Ragazzi");

  // Data e ora di inserimento
  const dataOraInserimento = new Date();

  // Aggiungo la riga: 
  // [ID_ragazzo, Anno scout, Tappa, Squadriglia, Ruolo, Incarico, Nota, Data, Data Inserimento]
  sheet.appendRow([
    form.ragazzo,
    form.anno || "",
    form.tappa || "",
    form.squadriglia || "",
    form.ruolo || "",
    form.incarico || "",
    form.nota || "",
    form.data || "",
    dataOraInserimento
  ]);

  // Formatto le date (colonna 8 = Data, colonna 9 = Data Inserimento)
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 8).setNumberFormat("dd/MM/yyyy");
  sheet.getRange(newRow, 9).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function addMeta(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mete");
  const newId = getNextId(sheet, 2, "M");

  const dataOraInserimento = new Date();
  sheet.appendRow([form.ragazzo, newId, form.meta, form.data, dataOraInserimento]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 4).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function addImpegno(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Impegni");
  const newId = getNextId(sheet, 2, "IM");

  const dataOraInserimento = new Date();
  sheet.appendRow([form.ragazzo, newId, form.impegno, form.data, dataOraInserimento]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 4).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function addSpecialita(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Specialità");
  const newId = getNextId(sheet, 2, "S");

  const dataOraInserimento = new Date();
  sheet.appendRow([form.ragazzo, newId, form.specialita, form.data, dataOraInserimento]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 5).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function addProva(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prove");
  const newId = getNextId(sheet, 4, "P");

  const dataOraInserimento = new Date();
  sheet.appendRow([form.ragazzo, form.id_specialita, getSpecialitaById(form.id_specialita).specialita, newId, form.num_prova, form.insieme, form.maestro, form.descrizione, form.nota, form.da_notificare, form.data, dataOraInserimento]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 7).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function addPP(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Progressione Personale");
  const newId = getNextId(sheet, 2, "PP");

  const dataOraInserimento = new Date();
  sheet.appendRow([form.ragazzo, newId, form.PP, form.data, dataOraInserimento]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 4).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function addConsiglio(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Consigli");
  const newId = getNextId(sheet, 1, "C");

  const dataOraInserimento = new Date();
  sheet.appendRow([newId, form.tema, form.data, dataOraInserimento]);

  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 4).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}

function addIntervento(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interventi");
  const newId = getNextId(sheet, 1, "I");

  const dataOraInserimento = new Date();
  sheet.appendRow([newId, form.consiglio, form.ragazzo, form.intervento, form.data || "", dataOraInserimento]);
  
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 6).setNumberFormat("dd/MM/yyyy HH:mm:ss");
}


// =====================================================================
// FUNZIONI PER POPOLARE LE TENDINE
// =====================================================================

function getRagazziList(completa = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ragazzi");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); 
  // colonne: ID_ragazzo, Nome, Cognome

  return data
    .filter(r => {
      if (completa) return true;       // se true non filtra
      return isRagazzoAttivo(r[0]);    // altrimenti applica il filtro
    })
    .map(r => [r[0], r[1] + " " + r[2]]);
}

function getSquadriglie() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Info");
  const values = sheet.getRange("A2:A").getValues(); // Colonna A
  return values.flat().filter(v => v && v.toString().trim() !== "");
}

function getTappe() { 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Info");
  const values = sheet.getRange("I2:I").getValues(); // Colonna I
  return values.flat().filter(v => v && v.toString().trim() !== "");
}

function getRuoli() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Info");
  const values = sheet.getRange("E2:E").getValues(); // Colonna E
  return values.flat().filter(v => v && v.toString().trim() !== "");
}

function getIncarichi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Info");
  const values = sheet.getRange("G2:G").getValues(); // Colonna G
  return values.flat().filter(v => v && v.toString().trim() !== "");
}


function getSpecialitaList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Info");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // Nessuna specialità presente

  const data = sheet.getRange(2, 3, lastRow - 1, 1).getValues(); // colonna C dalla riga 2
  return data
    .map(r => r[0])
    .filter(s => s && s.toString().trim() !== ""); // rimuove celle vuote
}

function getSpecialitaByRagazzo(idRagazzo) {
  if (!idRagazzo) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Specialità");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); 
  return data
    .filter(r => String(r[0] || '').toString().trim() === String(idRagazzo).toString().trim())
    .map(r => [String(r[0] || ''), String(r[1] || ''), String(r[2] || '')]);
}

function getConsigliList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Consigli");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // nessun consiglio

  const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues(); 
  // colonne: [ID_consiglio, Tema]
  return data
    .filter(row => row[0] && row[1]) // esclude righe vuote
    .map(row => [row[0], row[1]]);   // ID e Tema
}

function getInfoRagazzo(idRagazzo, campo, checkPresentAnno = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Info Ragazzi");

  // Recupero dati
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Mappa colonne
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxId = headers.indexOf("ID_ragazzo");
  const idxCampo = headers.indexOf(campo);
  const idxData = headers.indexOf("Data");
  const idxInserimento = headers.indexOf("Data Inserimento");

  if (idxId === -1 || idxCampo === -1 || idxInserimento === -1) {
    throw new Error("Colonne richieste non trovate");
  }

  // Filtra righe con quell'ID
  const righe = data.filter(r => r[idxId] == idRagazzo && r[idxCampo] !== "");

  if (righe.length === 0) return null;

  // Trova la riga più recente
  let ultima = null;
  let ultimaData = null;

  righe.forEach(r => {
    let d = r[idxData] || r[idxInserimento]; // se manca Data usa Data Inserimento
    if (d instanceof Date) {
      if (!ultimaData || d > ultimaData) {
        ultimaData = d;
        ultima = r[idxCampo];
      }
    }
  });

  if (checkPresentAnno && !isDateInPresentAnno(ultimaData)) {
    return "no info per l'anno corrente";
  }

  return ultima;
}

/**
 * Restituisce i valori più recenti di più campi per un dato ragazzo.
 * @param {string|number} idRagazzo - ID del ragazzo da cercare
 * @param {Array<string>} campi - Lista dei campi da restituire (es. ["Anno scout", "Tappa", "Ruolo"])
 * @param {Array<boolean>} checkPresentAnno - Lista parallela di boolean (true se il campo deve appartenere all'anno corrente)
 * @param {Array<Array>} dataOpt - (Opzionale) dati già letti dal foglio "Info Ragazzi" (include intestazioni nella riga 0)
 * @returns {Object} - Oggetto con chiavi = nomi campo e valori = ultimo valore valido
 */
function getInfoRagazzoAll(idRagazzo, campi, checkPresentAnno = [], dataOpt = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Info Ragazzi");

  // --- Lettura dati: se non forniti, leggo solo una volta
  let data;
  let headers;
  if (dataOpt && dataOpt.length > 1) {
    headers = dataOpt[0];
    data = dataOpt.slice(1);
  } else {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    const lastCol = sheet.getLastColumn();
    headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  }

  const idxId = headers.indexOf("ID_ragazzo");
  const idxData = headers.indexOf("Data");
  const idxInserimento = headers.indexOf("Data Inserimento");

  if (idxId === -1 || idxData === -1 || idxInserimento === -1) {
    throw new Error("Colonne richieste non trovate nel foglio Info Ragazzi");
  }

  // Prepara struttura risultati
  const risultati = {};
  campi.forEach(c => (risultati[c] = null));

  // Trova indice di ogni campo solo una volta
  const idxCampi = campi.map(campo => headers.indexOf(campo));

  // Cicla una sola volta su tutte le righe del foglio
  const valoriTrovati = {}; // {campo: {data: Date, valore: any}}
  for (let i = 0; i < data.length; i++) {
    const r = data[i];
    if (r[idxId] != idRagazzo) continue;

    for (let j = 0; j < campi.length; j++) {
      const idxCampo = idxCampi[j];
      if (idxCampo === -1 || !r[idxCampo]) continue;

      let d = r[idxData] instanceof Date ? r[idxData] : r[idxInserimento];
      if (!(d instanceof Date)) continue;

      // se richiesto, scarta se non nell’anno corrente
      if (checkPresentAnno[j] && !isDateInPresentAnno(d)) continue;

      if (!valoriTrovati[campi[j]] || d > valoriTrovati[campi[j]].data) {
        valoriTrovati[campi[j]] = { data: d, valore: r[idxCampo] };
      }
    }
  }

  // Prepara output finale
  campi.forEach(campo => {
    risultati[campo] = valoriTrovati[campo] ? valoriTrovati[campo].valore : null;
  });

  return risultati;
}


// =====================================================================
// FUNZIONI PER reportRagazzo
// =====================================================================

function getMetaById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mete");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues(); // ID_meta, Meta, Data
  const row = data.find(r => r[0] === id);
  if (!row) throw new Error("Meta non trovata");
  return { id: row[0], meta: row[1], data: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "dd/MM/yyyy") : "" };
}

function getImpegnoById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Impegni");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues(); // ID_impegno, Impegno, Data
  const row = data.find(r => r[0] === id);
  if (!row) throw new Error("Impegno non trovato");
  return { id: row[0], impegno: row[1], data: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "dd/MM/yyyy") : "" };
}

function getSpecialitaById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Specialità");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues(); // ID_specialita, Specialità, Data
  const row = data.find(r => r[0] === id);
  if (!row) throw new Error("Specialità non trovata");
  return { id: row[0], specialita: row[1], data: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "dd/MM/yyyy") : "" };
}

/**
 * Recupera i dettagli di una specifica prova tramite il suo ID.
 * @param {string} id L'ID della prova da recuperare (es. "P012").
 * @returns {object} Un oggetto contenente i dettagli della prova.
 */
function getProvaById(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetProve = ss.getSheetByName("Prove");

  if (sheetProve.getLastRow() < 2) return null;

  const proveData = sheetProve.getRange(2, 1, sheetProve.getLastRow() - 1, 12).getValues();

  const provaRow = proveData.find(row => row[3] === id);
  if (!provaRow) return null;
  
  return {
    id: provaRow[3], // Colonna D: Prova
    insieme: provaRow[5], // Colonna F: Insiemea
    maestro: provaRow[6], // Colonna G: MaestrodiSpecialità
    descrizione: provaRow[7], // Colonna H: Prova
    nota: provaRow[8],      // Colonna I: Note
    da_notificare: provaRow[9],
    data: provaRow[10] instanceof Date ? Utilities.formatDate(provaRow[10], Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
    data_inserimento: provaRow[11] instanceof Date ? Utilities.formatDate(provaRow[11], Session.getScriptTimeZone(), "dd/MM/yyyy") : ""
  };
}

function getPPById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Progressione Personale");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 3).getValues(); // ID_PP, PP, Data
  const row = data.find(r => r[0] === id);
  if (!row) throw new Error("Progressione Personale non trovata");
  return { id: row[0], PP: row[1], data: row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "dd/MM/yyyy") : "" };
}

// Funzione per reportConsiglio
function getInterventoById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interventi");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); // ID_intervento, ID_consiglio, ID_ragazzo, Intervento
  const row = data.find(r => r[0] === id);
  if (!row) throw new Error("Intervento non trovato");
  return { id: row[0], intervento: row[3] };
}


function parseItalianDate(dateStr) {
  if (!dateStr) return null;
  const parts = dateStr.split("/");
  if (parts.length !== 3) return null;
  const giorno = parseInt(parts[0], 10);
  const mese = parseInt(parts[1], 10) - 1; // mesi 0-11
  const anno = parseInt(parts[2], 10);
  return new Date(anno, mese, giorno);
}

function editMeta(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Mete");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(r => r[0] === form.id);
  if (rowIndex === -1) throw new Error("Meta non trovata");
  const row = rowIndex + 2;
  sheet.getRange(row, 3).setValue(form.meta);
  if (form.data) {
    const parsedDate = parseItalianDate(form.data); // funzione che converte "dd/MM/yyyy"
    if (parsedDate) {
      const cell = sheet.getRange(row, 4);
      cell.setValue(parsedDate);
      cell.setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  }
}

function editImpegno(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Impegni");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(r => r[0] === form.id);
  if (rowIndex === -1) throw new Error("Impegno non trovato");
  const row = rowIndex + 2;
  sheet.getRange(row, 3).setValue(form.impegno);
  if (form.data) {
    const parsedDate = parseItalianDate(form.data); // funzione che converte "dd/MM/yyyy"
    if (parsedDate) {
      const cell = sheet.getRange(row, 4);
      cell.setValue(parsedDate);
      cell.setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  }
}

function editSpecialita(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Specialità");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(r => r[0] === form.id);
  if (rowIndex === -1) throw new Error("Specialità non trovata");
  const row = rowIndex + 2;
  sheet.getRange(row, 3).setValue(form.specialita);
  if (form.data) {
    const parsedDate = parseItalianDate(form.data); // funzione che converte "dd/MM/yyyy"
    if (parsedDate) {
      const cell = sheet.getRange(row, 4);
      cell.setValue(parsedDate);
      cell.setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  }
}

/**
 * Modifica una prova esistente nel foglio "Prove".
 * @param {object} form Un oggetto contenente i dati della prova da modificare.
 * Deve includere: id, descrizione, insieme, maestro, nota, data.
 */
function editProva(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prove");
  const data = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues();
  
  const rowIndex = data.findIndex(r => r[0] === form.id);
  if (rowIndex === -1) throw new Error("Prova non trovata");
  
  const row = rowIndex + 2; // +2 perché getRange parte da riga 2 e findIndex è 0-based

  // Aggiorna le colonne corrette in base alla nuova struttura
  sheet.getRange(row, 6).setValue(form.insieme);      // Colonna F: Insiemea
  sheet.getRange(row, 7).setValue(form.maestro);      // Colonna G: MaestrodiSpecialità
  sheet.getRange(row, 8).setValue(form.descrizione);    // Colonna H: Prova
  sheet.getRange(row, 9).setValue(form.nota);           // Colonna I: Note
  sheet.getRange(row, 10).setValue(form.da_notificare);    // Colonna J: Prova
  
  if (form.data) {
    try {
      // Prova a parsare la data, potrebbe essere già in formato corretto
      const parts = form.data.split('/');
      const parsedDate = new Date(parts[2], parts[1] - 1, parts[0]);
      if (!isNaN(parsedDate.getTime())) {
         const cell = sheet.getRange(row, 11); // Colonna K: Data
         cell.setValue(parsedDate).setNumberFormat("dd/MM/yyyy");
      }
    } catch (e) {
      Logger.log("Formato data non valido per la modifica: " + form.data);
    }
  } else {
    // Se la data è vuota, cancella il contenuto della cella
    sheet.getRange(row, 11).clearContent();
  }
}


function editPP(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Progressione Personale");
  const data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(r => r[0] === form.id);
  if (rowIndex === -1) throw new Error("Progressione Personale non trovata");
  const row = rowIndex + 2;
  sheet.getRange(row, 3).setValue(form.PP);
  if (form.data) {
    const parsedDate = parseItalianDate(form.data); // funzione che converte "dd/MM/yyyy"
    if (parsedDate) {
      const cell = sheet.getRange(row, 4);
      cell.setValue(parsedDate);
      cell.setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  }
}





// =====================================================================
// FUNZIONI DI CONTROLLO
// =====================================================================

function checkRagazzoExists(nome, cognome) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ragazzi");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // Nessun ragazzo ancora inserito

  const data = sheet.getRange(2, 2, lastRow - 1, 2).getValues(); 
  // Colonne: Nome (colonna B), Cognome (colonna C)

  return data.some(r => 
    r[0].toString().trim().toLowerCase() === nome.toString().trim().toLowerCase() &&
    r[1].toString().trim().toLowerCase() === cognome.toString().trim().toLowerCase()
  );
}

function checkSpecialitaExists(idRagazzo, specialita) {
  if (!idRagazzo || !specialita) return false;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Specialità");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); 
  // colonne: ID_ragazzo, ID_specialità, Specialità
  return data.some(r => r[0] === idRagazzo && r[2].toString().trim() === specialita.toString().trim());
}

function isRagazzoAttivo(idRagazzo) {
  if (!idRagazzo) return false;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Ragazzi");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // Nessun ragazzo presente

  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // Colonne 1-5

  // Cerca il ragazzo per ID
  const ragazzo = data.find(r => r[0].toString().trim() === idRagazzo.toString().trim());
  if (!ragazzo) return false; // ID non trovato

  // Colonna 5 = stato, se "uscito" allora false
  return ragazzo[4].toString().trim().toLowerCase() !== "uscito";
}


/**
 * Parsea stringhe tipo "dd/MM/yyyy" o "dd/MM/yyyy HH:mm:ss"
 * (accetta sia ":" che "." come separatore orario) oppure prova
 * il fallback a new Date(string).
 * @returns {Date|null}
 */
function parseEuropeanDateString(s) {
  if (typeof s !== 'string') return null;
  s = s.trim();
  if (!s) return null;

  const parts = s.split(' ');
  const dateParts = parts[0].split('/');
  if (dateParts.length === 3) {
    const giorno = Number(dateParts[0]);
    const mese = Number(dateParts[1]);
    const anno = Number(dateParts[2]);
    if (isNaN(giorno) || isNaN(mese) || isNaN(anno)) return null;

    let ore = 0, minuti = 0, secondi = 0;
    if (parts.length >= 2) {
      // sostituisco i punti con i due punti per gestire "0.00.00"
      const timeStr = parts.slice(1).join(' ').replace(/\./g, ':');
      const t = timeStr.split(':').map(x => Number(x));
      ore = isNaN(t[0]) ? 0 : t[0];
      minuti = isNaN(t[1]) ? 0 : t[1];
      secondi = isNaN(t[2]) ? 0 : t[2];
    }
    return new Date(anno, mese - 1, giorno, ore, minuti, secondi);
  }

  // fallback generico
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Normalizza molti tipi di input in Date.
 * Accetta: Date, number (timestamp), stringa (formato europeo o ISO).
 * @returns {Date|null}
 */
function toDate(val) {
  if (val instanceof Date) return val;
  if (typeof val === 'number') return new Date(val);
  if (typeof val === 'string') {
    // prima provo il parser europeo, poi fallback a Date()
    const p = parseEuropeanDateString(val);
    if (p) return p;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

/**
 * Controlla se dataInput (stringa/dd/MM/yyyy... oppure Date) è compresa
 * nell'intervallo dato dai parametri in input 'inizio' e 'fine'.
 */
function isDateInInterval(dataInput, rawInizio, rawFine) {
  let fine = null;
  if (typeof rawFine === 'string' && rawFine.toString().trim().toLowerCase() === 'oggi') {
    fine = new Date(2500, 11, 31, 23, 59, 59);  // -> consideralo "aperto" -> far-future
  } else {
    fine = toDate(rawFine);
  }
  
  const inizio = toDate(rawInizio);
  const inputDate = toDate(dataInput);

  // se una delle tre date non è valida -> non posso decidere, ritorno false
  if (!inizio) return false;
  if (!fine) return false;
  if (!inputDate) return false;

  return inputDate >= inizio && inputDate <= fine;
}

/**
 * Controlla se dataInput (stringa/dd/MM/yyyy... oppure Date) è compresa
 * tra Inizio e Fine dell'ultima riga di "Registro anni".
 * Se la colonna Fine contiene la stringa "oggi" (ignorando case/whitespace)
 * viene interpretata come "ora" (new Date()).
 */
function isDateInPresentAnno(dataInput) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registro anni");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false; // nessun dato

  const values = sheet.getRange(lastRow, 2, 1, 2).getValues()[0];
  const rawInizio = values[0];
  const rawFine = values[1];

  return isDateInInterval(dataInput, rawInizio, rawFine);
}


// =====================================================================
// FUNZIONI PER ELIMINARE UNA RIGA IN BASE ALL'ID
// =====================================================================

/**
 * Funzione di utilità generica per eliminare una riga in base all'ID in una colonna specifica.
 * @param {string} sheetName - Il nome del foglio di calcolo.
 * @param {string} id - L'ID univoco dell'elemento da eliminare.
 * @param {number} idColumn - Il numero della colonna (1-based) in cui cercare l'ID.
 */
function deleteRowById(sheetName, id, idColumn) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Foglio di calcolo "${sheetName}" non trovato.`);
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  
  let rowToDelete = -1;
  // Cicla a partire dalla seconda riga (la prima è l'intestazione)
  for (let i = 1; i < values.length; i++) {
    // Il valore della colonna nell'array è idColumn - 1
    if (values[i][idColumn - 1] === id) {
      rowToDelete = i + 1; // Le righe di Apps Script sono 1-based
      break;
    }
  }

  if (rowToDelete !== -1) {
    sheet.deleteRow(rowToDelete);
    return `Elemento con ID ${id} eliminato con successo dal foglio ${sheetName}.`;
  } else {
    throw new Error(`ID ${id} non trovato nel foglio ${sheetName}.`);
  }
}

// Funzioni specifiche per l'eliminazione

function deleteMeta(id) {
  try {
    deleteRowById("Mete", id, 2); // L'ID si trova nella colonna 2 (B)
  } catch (e) {
    console.error("Errore durante l'eliminazione della meta: " + e.message);
    throw new Error("Errore durante l'eliminazione della meta.");
  }
}

function deleteImpegno(id) {
  try {
    deleteRowById("Impegni", id, 2); // L'ID si trova nella colonna 2 (B)
  } catch (e) {
    console.error("Errore durante l'eliminazione dell'impegno: " + e.message);
    throw new Error("Errore durante l'eliminazione dell'impegno.");
  }
}

function deleteSpecialita(id) {
  try {
    // 1. Elimina la specialità principale
    deleteRowById("Specialità", id, 2); // L'ID si trova nella colonna 2 (B)
    
    // 2. Elimina le prove associate
    const proveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Prove");
    if (!proveSheet) {
        throw new Error("Foglio di calcolo 'Prove' non trovato.");
    }
    const proveData = proveSheet.getDataRange().getValues();
    const rowsToDelete = [];
    
    // Assumendo che l'ID della specialità si trovi nella colonna 2 (B) ('ID Specialità') nel foglio "Prove"
    const specialitaIdColumn = 2; 
    
    for (let i = proveData.length - 1; i >= 1; i--) { // Cicla all'indietro dalla fine (per non perdere il riferimento delle righe) e salta l'intestazione
        if (proveData[i][specialitaIdColumn - 1] === id) {
            rowsToDelete.push(i + 1); // Aggiungi il numero di riga (1-based)
        }
    }
    
    // Elimina le righe trovate
    rowsToDelete.forEach(row => proveSheet.deleteRow(row));

    return "Specialità e prove associate eliminate con successo.";
  } catch (e) {
    console.error("Errore durante l'eliminazione della specialità: " + e.message);
    throw new Error("Errore durante l'eliminazione della specialità.");
  }
}

function deleteProva(id) {
  try {
    deleteRowById("Prove", id, 4); // L'ID si trova nella colonna 4 (D)
  } catch (e) {
    console.error("Errore durante l'eliminazione della prova: " + e.message);
    throw new Error("Errore durante l'eliminazione della prova.");
  }
}

function deletePP(id) {
  try {
    deleteRowById("Progressione Personale", id, 2); // L'ID si trova nella colonna 2 (B)
  } catch (e) {
    console.error("Errore durante l'eliminazione della PP: " + e.message);
    throw new Error("Errore durante l'eliminazione della PP.");
  }
}


// =====================================================================
// FUNZIONI DI MODIFICA E CANCELLAZIONE PER reportConsiglio.html
// =====================================================================

function editIntervento(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interventi");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(r => r[0] === form.id);
  if (rowIndex === -1) throw new Error("Intervento non trovato");
  const row = rowIndex + 2;
  sheet.getRange(row, 4).setValue(form.intervento);
}

function deleteIntervento(id) {
  try {
    deleteRowById("Interventi", id, 1); // L'ID si trova nella colonna 1 (A)
  } catch (e) {
    console.error("Errore durante l'eliminazione dell'intervento: " + e.message);
    throw new Error("Errore durante l'eliminazione dell'intervento.");
  }
}


// =====================================================================
// FUNZIONI PER MOSTRARE I POPUP
// =====================================================================

function showAddAnno() {
  const html = HtmlService.createHtmlOutputFromFile('addAnno')
      .setWidth(300)
      .setHeight(300)
  SpreadsheetApp.getUi().showModalDialog(html, 'Registra Nuovo Anno Scout');
}

function showAddRagazzo() {
  const html = HtmlService.createHtmlOutputFromFile('addRagazzo')
    .setWidth(300)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Ragazzo');
}

function showRemoveRagazzo() {
  const html = HtmlService.createHtmlOutputFromFile('removeRagazzo')
    .setWidth(300)
    .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(html, 'Rimuovi Ragazzo');
}

function showAddInfoRagazzo() {
  const html = HtmlService.createHtmlOutputFromFile('addInfoRagazzo')
    .setWidth(350)
    .setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Info Ragazzo');
}

function showAddInfoRagazzo_sub(idRagazzo) {
  const template = HtmlService.createTemplateFromFile('addInfoRagazzo');
  template.ragazzoId = idRagazzo;  // passo l'id al popup
  const html = template.evaluate()
    .setWidth(350)
    .setHeight(450);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Aggiungi Info Ragazzo');
}

function showAddMeta() {
  const html = HtmlService.createHtmlOutputFromFile('addMeta')
    .setWidth(600)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Meta');
}

function showAddImpegno() {
  const html = HtmlService.createHtmlOutputFromFile('addImpegno')
    .setWidth(600)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Impegno');
}

function showAddSpecialita() {
  const html = HtmlService.createHtmlOutputFromFile('addSpecialita')
    .setWidth(300)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Specialità');
}

function showAddProva() {
  const html = HtmlService.createHtmlOutputFromFile('addProva')
    .setWidth(600)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Prova');
}

function showAddPP() {
  const html = HtmlService.createHtmlOutputFromFile('addPP')
    .setWidth(600)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Progressione Personale');
}

function showAddConsiglio() {
  const html = HtmlService.createHtmlOutputFromFile('addConsiglio')
    .setWidth(300)
    .setHeight(250);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Consiglio');
}

function showAddIntervento() {
  const html = HtmlService.createHtmlOutputFromFile('addIntervento')
    .setWidth(300)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, 'Aggiungi Intervento');
}

function showReportRagazzo() {
  const html = HtmlService.createHtmlOutputFromFile('reportRagazzo')
    .setWidth(600)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Report EG');
}

function showReportConsiglio() {
  const html = HtmlService.createHtmlOutputFromFile('reportConsiglio')
    .setWidth(600)
    .setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Report EG');
}

function showReportRagazzoDoc() {
    const html = HtmlService.createHtmlOutputFromFile('reportRagazzoDoc')
        .setWidth(400)
        .setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Crea Documenti Report');
}


// =====================================================================
// FUNZIONI UTILITY E REPORT
// =====================================================================

function formatDateItalian(date, format) {
  if (!(date instanceof Date)) return "";
  
  // mesi italiani
  const mesiIT = [
    "gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno",
    "luglio", "agosto", "settembre", "ottobre", "novembre", "dicembre"
  ];
  const mesiShortIT = [
    "gen", "feb", "mar", "apr", "mag", "giu",
    "lug", "ago", "set", "ott", "nov", "dic"
  ];

  let out = Utilities.formatDate(date, Session.getScriptTimeZone(), format);

  // sostituisci i mesi inglesi con quelli italiani
  const mesiEN = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  const mesiShortEN = [
    "Jan","Feb","Mar","Apr","May","Jun",
    "Jul","Aug","Sep","Oct","Nov","Dec"
  ];

  mesiEN.forEach((m, i) => {
    out = out.replace(m, mesiIT[i]);
  });
  mesiShortEN.forEach((m, i) => {
    out = out.replace(m, mesiShortIT[i]);
  });

  return out;
}

function getReportRagazzoData(idRagazzo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Nome e Cognome
  const sheetRag = ss.getSheetByName("Ragazzi");
  const lastRowRag = sheetRag.getLastRow();
  let dataRag = [];
  if (lastRowRag > 1) {
    dataRag = sheetRag.getRange(2, 1, lastRowRag - 1, 3).getValues();
  }
  const ragazzo = dataRag.find(r => r[0] === idRagazzo);
  const nomeCompleto = ragazzo ? (ragazzo[1] + " " + ragazzo[2]) : "Sconosciuto";

  // --- Dati base del ragazzo: Nome e Cognome, Anno, Squadriglia, Tappa e Ruolo attuale
  const info_ragazzo = getInfoRagazzoAll(
    idRagazzo,
    ["Anno scout", "Tappa", "Squadriglia", "Ruolo", "Incarico"],
    [true, false, false, false, true]
  );
  const anno = info_ragazzo["Anno scout"] + "° anno" || "Nessun anno registrato";
  const tappa = info_ragazzo["Tappa"] || "Nessuna tappa registrata";
  const squadriglia = info_ragazzo["Squadriglia"] || "Nessuna squadriglia registrata";
  const ruolo = info_ragazzo["Ruolo"] || "Nessun ruolo registrato";
  const incarico = info_ragazzo["Incarico"] || "Nessun incarico registrato";

  const statoRagazzo = isRagazzoAttivo(idRagazzo);

  // --- Mete
  const sheetMete = ss.getSheetByName("Mete");
  const lastRowMete = sheetMete.getLastRow();
  const mete = lastRowMete > 1
    ? sheetMete.getRange(2, 1, lastRowMete - 1, 5).getValues()
        .filter(r => r[0] === idRagazzo && isDateInPresentAnno(r[3] || r[4]))
        .map(r => ({ id: r[1], meta: r[2], data: formatDateItalian(r[3], "d MMMM yyyy") }))
    : [];

  // --- Impegni
  const sheetImp = ss.getSheetByName("Impegni");
  const lastRowImp = sheetImp.getLastRow();
  const impegni = lastRowImp > 1
    ? sheetImp.getRange(2, 1, lastRowImp - 1, 5).getValues()
        .filter(r => r[0] === idRagazzo && isDateInPresentAnno(r[3] || r[4]))
        .map(r => ({ id: r[1], impegno: r[2], data: formatDateItalian(r[3], "d MMMM yyyy") }))
    : [];

  // --- Specialità
  const sheetSpec = ss.getSheetByName("Specialità");
  const lastRowSpec = sheetSpec.getLastRow();
  const specData = lastRowSpec > 1
    ? sheetSpec.getRange(2, 1, lastRowSpec - 1, 3).getValues()
        .filter(r => r[0] === idRagazzo)
    : [];

  // --- Prove
  const sheetProve = ss.getSheetByName("Prove");
  const lastRowProve = sheetProve.getLastRow();
  const proveData = lastRowProve > 1
    ? sheetProve.getRange(2, 1, lastRowProve - 1, 12).getValues()
    : [];

  // --- Costruzione array specialità con prove associate
  const specialita = specData.map(s => {
    const proveAssoc = proveData
      .filter(p => p[0] === idRagazzo && p[1] === s[1])
      .sort((a, b) => {
        const numA = parseInt(a[3].replace(/P0*/, ''));
        const numB = parseInt(b[3].replace(/P0*/, ''));
        return numA - numB;
      })
      .map(p => {
        const num = p[4];
        const insieme = p[5];
        const maestro = p[6];
        const descrizione = p[7];
        const nota = p[8];
        const da_notificare = p[9] || "No";
        const data = p[10] ? `${formatDateItalian(p[10], "MMM yyyy")}` : "";
        const data_inserimento = p[11] ? `${formatDateItalian(p[11], "dd MMMM yyyy")}` : "";
        return {
          id: p[3], // Aggiungi l'ID della prova
          insieme: insieme,
          maestro: maestro,
          descrizione: `${num}ª prova: ${descrizione}`,
          nota: nota,
          da_notificare: da_notificare,
          data: data,
          data_inserimento: data_inserimento
        };
      });
    return { id: s[1], nome: s[2], prove: proveAssoc };
  });

  // --- Progressione Personale (PP)
  const sheetPP = ss.getSheetByName("Progressione Personale"); 
  const lastRowPP = sheetPP.getLastRow();
  const ppData = lastRowPP > 1
      ? sheetPP.getRange(2, 1, lastRowPP - 1, 5).getValues()
          .filter(r => r[0] === idRagazzo && r[2])
          .map(r => {
              let timestamp = 0;
              if (r[3] instanceof Date && !isNaN(r[3].getTime())) {
                  timestamp = r[3].getTime();
              }
              else if (r[4] instanceof Date && !isNaN(r[4].getTime())) {
                  timestamp = r[4].getTime();
              }
              return {
                  id: r[1], // Aggiungi l'ID della PP
                  timestamp: timestamp,
                  dataStr: formatDateItalian(r[3], "d MMMM yyyy"),
                  testo: r[2]
              };
          })
      : [];

  // --- Storico
  const sheetStorico = ss.getSheetByName("Info Ragazzi");
  const lastRowStorico = sheetStorico.getLastRow();
  
  const sheetAnni = ss.getSheetByName("Registro anni");
  const lastRowAnni = sheetAnni.getLastRow();
  const AnniData = lastRowAnni > 1 ? sheetAnni.getRange(2, 1, lastRowAnni - 1, sheetAnni.getLastColumn()).getValues() : [];
  const numeri_id_anni = AnniData.map(riga => parseInt(riga[0].slice(1), 10));
  const maxID_anno = numeri_id_anni.length > 0 ? Math.max(...numeri_id_anni) : 0;   // numero di anni registrati
  
  const StoricoData = lastRowStorico > 1
      ? sheetStorico.getRange(2, 1, lastRowStorico - 1, sheetStorico.getLastColumn()).getValues()
          .filter(r => r[0] === idRagazzo)
          .sort((a, b) => a[7] ? (a[7].getTime() - b[7].getTime()) : (a[8].getTime() - b[8].getTime()))
          .map(r => {
              r[1] = (!isNaN(r[1]) && Number(r[1]) !== 0) 
                  ? String(Math.round(Number(r[1]))) + "° anno"
                  : "";
              let timestamp = 0;
              let data_utile = r[7]
              if (r[7] instanceof Date && !isNaN(r[7].getTime())) {
                  timestamp = r[7].getTime();
              }
              else if (r[8] instanceof Date && !isNaN(r[8].getTime())) {
                  timestamp = r[8].getTime();
                  data_utile = r[8];
              }

              let anno_id = null;
              let titolo_anno = null;
              info = r.slice(1, 7);
              for (let i = 0; i < AnniData.length; i++) {
                  const rawInizio = AnniData[i][1];
                  const rawFine   = AnniData[i][2];
                  if (isDateInInterval(data_utile, rawInizio, rawFine)) {
                      anno_id = AnniData[i][0]; // prima colonna della riga
                      anno_id = parseInt(anno_id.replace("A", ""), 10);
                      titolo_anno = formatDateItalian(rawInizio, "MMM yyyy").toString() + " - " + ((rawFine=="oggi") ? "oggi" : formatDateItalian(rawFine, "MMM yyyy").toString());
                      break; // trovato l’intervallo, esco
                  }
              }

              return {
                  timestamp: timestamp,
                  dataStr: formatDateItalian(data_utile, "d MMM yyyy"),
                  info: info,
                  num_id_anno: anno_id,
                  anno_titolo: titolo_anno
              };
          })
      : [];

  // --- Ritorno di tutti i dati raccolti
  return {
    id: idRagazzo, // Aggiungi l'ID del ragazzo
    nomeCompleto: nomeCompleto,
    anno: anno,
    tappa: tappa,
    squadriglia: squadriglia,
    ruolo: ruolo,
    incarico: incarico,
    statoRagazzo: statoRagazzo,
    mete: mete,
    impegni: impegni,
    specialita: specialita,
    PP: ppData,
    storico: StoricoData,
    maxID_anno: maxID_anno
  };
}


function getReportConsiglioData(idConsiglio) {
  if (!idConsiglio) return null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- Dati Consiglio
    const sheetCons = ss.getSheetByName("Consigli");
    const lastRowCons = sheetCons.getLastRow();
    const dataCons = lastRowCons > 1 ? sheetCons.getRange(2, 1, lastRowCons - 1, 3).getValues() : [];
    const consiglio = dataCons.find(c => c[0].toString().trim() === idConsiglio.toString().trim());

    if (!consiglio) throw new Error("Consiglio non trovato");

    const temaConsiglio = consiglio[1];
    const dataDelConsiglio = formatDateItalian(consiglio[2], "dd MMMM yyyy");

    // --- Ragazzi
    const sheetRag = ss.getSheetByName("Ragazzi");
    const lastRowRag = sheetRag.getLastRow();
    const dataRag = lastRowRag > 1 ? sheetRag.getRange(2, 1, lastRowRag - 1, 5).getValues() : [];
    const ragazziAttivi = dataRag.filter(r => r[4] !== 'uscito');   // considera solo i ragazzi presenti in reparto
    
    // --- Interventi
    const sheetInt = ss.getSheetByName("Interventi");
    const lastRowInt = sheetInt.getLastRow();
    const dataInt = lastRowInt > 1 ? sheetInt.getRange(2, 1, lastRowInt - 1, 4).getValues() : [];
    const interventiDelConsiglio = dataInt.filter(int => int[1].toString().trim() === idConsiglio.toString().trim());

    const squadriglieMap = {};
    ragazziAttivi.forEach(rag => {
      const [idRag, nome, cognome] = rag;
      const sqNome = getInfoRagazzo(idRag, "Squadriglia") || "Senza Squadriglia";
      const ruolo = (getInfoRagazzo(idRag, "Ruolo") || "Squadrigliere").trim();
      const anno = Number(getInfoRagazzo(idRag, "Anno scout")) || 99;

      if (!squadriglieMap[sqNome]) {
        squadriglieMap[sqNome] = {};
      }

      squadriglieMap[sqNome][idRag] = {
        nomeCompleto: nome + " " + cognome,
        interventi: [],
        ruolo: ruolo,
        anno: anno
      };
    });

    interventiDelConsiglio.forEach(int => {
      const [idInt, , idRag, testo] = int;
      const squadriglia = getInfoRagazzo(idRag, "Squadriglia") || "Senza Squadriglia";
      if (squadriglieMap[squadriglia] && squadriglieMap[squadriglia][idRag]) {
        squadriglieMap[squadriglia][idRag].interventi.push({ id: idInt, testo: testo });
      }
    });

    // Logica di ordinamento: Capo ordinati per anno, Squadrigliere ordinati per anno, infine Vice ordinati per anno
    const ruoloOrder = { "Capo": 1, "Squadrigliere": 2, "Vice": 3 };
    for (const sqNome in squadriglieMap) {
        const ragazziObj = squadriglieMap[sqNome];
        
        // Converte l'oggetto in un array per poterlo ordinare
        const sortedRagazziArray = Object.entries(ragazziObj).sort(([, a], [, b]) => {
                const ruoloA = ruoloOrder[a.ruolo.trim()] || 4; // Valore alto per ruoli non definiti
                const ruoloB = ruoloOrder[b.ruolo.trim()] || 4;
            
            if (ruoloA !== ruoloB) {
                return ruoloA - ruoloB; // Ordina prima per ruolo
            } else {
                return a.anno - b.anno; // Se il ruolo è lo stesso, ordina per anno
            }
        });

        // Ricostruisce l'oggetto con i ragazzi in ordine
        squadriglieMap[sqNome] = sortedRagazziArray.map(([id, data]) => ({
            id,
            ...data
        }));
    }

    return {
      titoloConsiglio: temaConsiglio,
      dataConsiglio: dataDelConsiglio,
      squadriglie: squadriglieMap
    };

  } catch (e) {
    Logger.log("ERRORE in getReportConsiglioData: " + e.message);
    throw e;
  }
}


// =====================================================================
// FUNZIONI GOOGLE DOCS
// =====================================================================

/**
 * Crea un Documento Google con il report di un ragazzo.
 * @param {string} idRagazzo L'ID del ragazzo.
 * @returns {object} Un oggetto con lo stato dell'operazione.
 */
function creaReportInDoc(idRagazzo) {
  try {
    const data = getReportRagazzoData(idRagazzo);
    if (!data) {
      throw new Error("Dati del ragazzo non trovati.");
    }

    const folderName = "ReportRagazzo";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();

    let reportFolder;
    const folders = parentFolder.getFoldersByName(folderName);
    if (folders.hasNext()) {
      reportFolder = folders.next();
    } else {
      reportFolder = parentFolder.createFolder(folderName);
    }
    
    const docName = data.nomeCompleto;
    
    // Controlla se il file esiste già e lo elimina
    const existingFiles = reportFolder.getFilesByName(docName);
    while (existingFiles.hasNext()) {
      existingFiles.next().setTrashed(true);
    }
    
    const doc = DocumentApp.create(docName);
    const docFile = DriveApp.getFileById(doc.getId());
    reportFolder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile); // Pulisce la root

    const body = doc.getBody();

    // Stili
    const BOLD_STYLE = {};
    BOLD_STYLE[DocumentApp.Attribute.BOLD] = true;

    // Titolo
    body.appendParagraph(data.nomeCompleto).setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph("Report generato il: " + formatDateItalian(new Date(), "d MMMM yyyy HH:mm")).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setFontSize(9).setItalic(true);
    body.appendParagraph("\n");

    // Scheda Riepilogo
    body.appendParagraph("Scheda Riepilogo").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (!data.statoRagazzo) {
       body.appendParagraph("NON È PIÙ IN REPARTO").setBold(true).setForegroundColor("#FF0000");
    }
    let p = body.appendParagraph('');
    p.appendText("Anno di reparto: ").setAttributes(BOLD_STYLE);
    p.appendText(data.anno).setBold(false);
    p = body.appendParagraph('');
    p.appendText("Tappa: ").setAttributes(BOLD_STYLE);
    p.appendText(data.tappa).setBold(false);
    p = body.appendParagraph('');
    p.appendText("Squadriglia: ").setAttributes(BOLD_STYLE);
    p.appendText(data.squadriglia).setBold(false);
    p = body.appendParagraph('');
    p.appendText("Ruolo di sq: ").setAttributes(BOLD_STYLE);
    p.appendText(data.ruolo).setBold(false);
    p = body.appendParagraph('');
    p.appendText("Incarico di sq: ").setAttributes(BOLD_STYLE);
    p.appendText(data.incarico).setBold(false);
    body.appendParagraph("\n");

    // Mete e Impegni
    body.appendParagraph("Mete e Impegni").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (data.mete.length > 0) {
      body.appendParagraph("Mete").setHeading(DocumentApp.ParagraphHeading.HEADING2);
      data.mete.forEach(m => {
        body.appendListItem(`${m.meta}${m.data ? "   [" + m.data + "]" : ""}`).setGlyphType(DocumentApp.GlyphType.BULLET);
      });
    }
    if (data.impegni.length > 0) {
      body.appendParagraph("Impegni").setHeading(DocumentApp.ParagraphHeading.HEADING2);
      data.impegni.forEach(i => {
        body.appendListItem(`${i.impegno}${i.data ? "   [" + i.data + "]" : ""}`).setGlyphType(DocumentApp.GlyphType.BULLET);
      });
    }
    if (data.mete.length === 0 && data.impegni.length === 0) {
       body.appendParagraph("Nessuna meta o impegno registrato.").setItalic(true);
    }
     body.appendParagraph("\n");
    
    // Specialità e Brevetti
    body.appendParagraph("Specialità e Brevetti").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if(data.specialita.length > 0) {
      data.specialita.forEach(s => {
        body.appendParagraph(s.nome).setHeading(DocumentApp.ParagraphHeading.HEADING2);
        if (s.prove.length > 0) {
          s.prove.forEach(p => {
            let provaText = `${p.descrizione}${p.data ? "\u0009[" + p.data + "]" : ""}`;
            let details = [];
            if (p.insieme) details.push(`Prova fatta assieme a: ${p.insieme}`);
            if (p.maestro) details.push(`Maestro di Specialità: ${p.maestro}`);
            if (p.nota) details.push(`Note: ${p.nota}`);

            if (details.length > 0) {
              provaText += `\n   - ${details.join('\n   - ')}`;
            }

            let item = body.appendListItem(provaText).setGlyphType(DocumentApp.GlyphType.BULLET);
            if (p.da_notificare && p.da_notificare.toLowerCase() === 'sì') {
                item.setForegroundColor("#CC0000");
            }
          });
        } else {
           body.appendParagraph("Nessuna prova registrata.").setItalic(true);
        }
      });
    } else {
       body.appendParagraph("Nessuna specialità registrata.").setItalic(true);
    }
     body.appendParagraph("\n");

    // Progressione Personale
    body.appendParagraph("Progressione Personale").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (data.PP && data.PP.length > 0) {
      const sortedPP = data.PP.slice().sort((a, b) => b.timestamp - a.timestamp);
      let lastDate = null; // Variabile per tenere traccia dell'ultima data
      
      sortedPP.forEach(item => {
        // Se la data dell'elemento corrente è diversa dall'ultima registrata...
        if (item.dataStr !== lastDate) {
          // ...allora la aggiungiamo come un nuovo sottotitolo
          body.appendParagraph(item.dataStr).setHeading(DocumentApp.ParagraphHeading.HEADING2);
          lastDate = item.dataStr; // E aggiorniamo l'ultima data vista
        }
        // Aggiungiamo l'elemento della PP come punto elenco
        body.appendListItem(item.testo).setGlyphType(DocumentApp.GlyphType.BULLET);
      });
    } else {
      body.appendParagraph("Nessuna progressione personale registrata.").setItalic(true);
    }
    body.appendParagraph("\n");
    
    // Storico
    body.appendParagraph("Storico").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    if (data.storico && data.storico.length > 0) {
        let frasi = ["Nuovo anno registrato:", "Nuova tappa data:", "Cambio squadriglia:", "Nuovo ruolo di sq:", "Nuovo incarico:", "Nuova nota:"]
        for (let i = 1; i <= data.maxID_anno; i++) {
          const sortedStorico = data.storico.slice().filter(el => el.num_id_anno === i).sort((a, b) => a.timestamp - b.timestamp);
          if (sortedStorico.length > 0) {
            body.appendParagraph(sortedStorico[0].anno_titolo).setHeading(DocumentApp.ParagraphHeading.HEADING2);
            let last_item = new Array(sortedStorico[0].length).fill("");
            sortedStorico.forEach(item => {
              for(let k = 0; k < item.info.length; k++) {
                if(item.info[k] != last_item[k] && item.info[k] != "") {
                   let p = body.appendParagraph('');
                   p.appendText(`${frasi[k]} `);
                   p.appendText(item.info[k]).setAttributes(BOLD_STYLE);
                   p.appendText(`\u0009[${item.dataStr}]`).setBold(false);
                   last_item[k] = item.info[k];
                }
              }
            });
          }
        }
    } else {
       body.appendParagraph("Nessuno storico registrato.").setItalic(true);
    }

    doc.saveAndClose();
    return { success: true, message: "Documento creato/aggiornato!" };
  } catch (e) {
    Logger.log("Errore in creaReportInDoc: " + e.message);
    return { success: false, message: e.message };
  }
}


function creaDocumentoGoogle(titolo, data, contenutoHtml) {
  try {
    // Recupera il file del foglio corrente
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());

    // Recupera la cartella del foglio
    const folder = ssFile.getParents().hasNext() ? ssFile.getParents().next() : DriveApp.getRootFolder();

    // Crea un nuovo documento nella stessa cartella
    const doc = DocumentApp.create(titolo);
    const docFile = DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile); // rimuove dal root se necessario

    const body = doc.getBody();

    // Inserisce titolo e data
    body.appendParagraph(titolo).setHeading(DocumentApp.ParagraphHeading.HEADING1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(data).setFontSize(10).setAlignment(DocumentApp.HorizontalAlignment.CENTER).setItalic(true);
    body.appendParagraph(''); // spazio

    // Converte HTML base in testo formattato
    // Qui gestiamo solo h2, h3, ul/li in modo semplice
    const tempDiv = XmlService.parse('<div>' + contenutoHtml + '</div>').getRootElement();
    parseElement(tempDiv, body);

    doc.saveAndClose();
  } catch (e) {
    Logger.log("Errore: " + e);
    throw e;
  }
}

// Funzione ricorsiva per gestire h2, h3, ul/li
function parseElement(element, body) {
  const name = element.getName().toLowerCase();
  const children = element.getChildren();

  switch(name) {
    case 'h2':
      body.appendParagraph(element.getText()).setHeading(DocumentApp.ParagraphHeading.HEADING2).setForegroundColor('#2E86C1');
      break;
    case 'h3':
      body.appendParagraph(element.getText()).setHeading(DocumentApp.ParagraphHeading.HEADING3).setForegroundColor('#1B4F72');
      break;
    case 'ul':
      children.forEach(li => {
        if (li.getName().toLowerCase() === 'li') {
          body.appendListItem(li.getText()).setGlyphType(DocumentApp.GlyphType.BULLET);
        }
      });
      break;
    default:
      // se ci sono altri elementi, esplora ricorsivamente
      children.forEach(child => parseElement(child, body));
  }
}


// =====================================================================
// TRIGGER ONOPEN - Menu delle opzioni
// =====================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📒 Registro Scout')
    .addItem('➕ Aggiungi Ragazzo', 'showAddRagazzo')
    .addItem('➖ Rimuovi Ragazzo', 'showRemoveRagazzo')
    .addSeparator()
    .addItem('🆕 Nuovo Anno', 'showAddAnno')
    .addSeparator()
    .addSeparator()
    .addItem('ℹ️ Aggiungi Info Ragazzo', 'showAddInfoRagazzo')
    .addSeparator()
    .addItem('📍 Aggiungi Meta', 'showAddMeta')
    .addItem('📝 Aggiungi Impegno', 'showAddImpegno')
    .addSeparator()
    .addItem('⭐ Aggiungi Specialità', 'showAddSpecialita')
    .addItem('🧪 Aggiungi Prova', 'showAddProva')
    .addSeparator()
    .addItem('🌱 Aggiungi Progressione Personale', 'showAddPP')
    .addSeparator()
    .addSeparator()
    .addItem('🗓️ Aggiungi Consiglio', 'showAddConsiglio')
    .addItem('💬 Aggiungi Intervento', 'showAddIntervento')
    .addSeparator()
    .addSeparator()
    .addItem('📋 Report Ragazzo', 'showReportRagazzo')
    .addItem('📋 Report Consiglio', 'showReportConsiglio')
    .addSeparator()
    .addItem('📤 Report Ragazzi Docs', 'showReportRagazzoDoc')
    .addToUi();
}



// =====================================================================
// FUNZIONI PER LA WEB APP
// =====================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Registro Scout')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Questa funzione restituisce l'HTML completo di un modulo, script inclusi.
function loadForm(type) {
  let fileName = '';
  switch (type) {
    case 'addRagazzo': fileName = 'addRagazzo'; break;
    case 'removeRagazzo': fileName = 'removeRagazzo'; break;
    case 'addInfoRagazzo': fileName = 'addInfoRagazzo'; break;
    case 'addAnno': fileName = 'addAnno'; break;
    case 'addMeta': fileName = 'addMeta'; break;
    case 'addImpegno': fileName = 'addImpegno'; break;
    case 'addSpecialita': fileName = 'addSpecialita'; break;
    case 'addProva': fileName = 'addProva'; break;
    case 'addPP': fileName = 'addPP'; break;
    case 'addConsiglio': fileName = 'addConsiglio'; break;
    case 'addIntervento': fileName = 'addIntervento'; break;
    case 'reportRagazzo': fileName = 'reportRagazzo'; break;
    case 'reportConsiglio': fileName = 'reportConsiglio'; break;
    case 'reportRagazzoDoc': fileName = 'reportRagazzoDoc'; break;
    default:
      return { html: '<p>Modulo non trovato</p>' };
  }

  try {
    // Valuta il template e restituisce l'HTML completo, script inclusi.
    // Questo è il modo corretto per assicurarsi che tutto il contenuto venga eseguito.
    const htmlContent = HtmlService.createTemplateFromFile(fileName).evaluate().getContent();
    return { html: htmlContent };

  } catch (e) {
    Logger.log("Errore in loadForm: " + e.message);
    return { html: `<p>Errore caricando ${fileName}: ${e.message}</p>` };
  }
}

