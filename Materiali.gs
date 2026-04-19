// ================= MODULO GESTIONE MATERIALI ================= //

function getOrCreateMaterialiFile(segreteria) {
  let folder = DriveApp.getFolderById(FOLDER_MATERIALI_SHEETS_ID);
  let safeName = "Materiali_" + segreteria.replace(/[^a-zA-Z0-9 ()-]/g, "").trim();
  
  let files = folder.getFilesByName(safeName);
  if (files.hasNext()) {
    let existingFile = files.next();
    let ss = SpreadsheetApp.openById(existingFile.getId());
    if (!ss.getSheetByName('MATERIALI')) {
      let s = ss.insertSheet('MATERIALI');
      s.appendRow(['Timestamp', 'CIP Inseritore', 'Categoria', 'Descrizione', 'Quantità', 'Costo', 'Data Acquisto', 'Acquistato Da', 'In Carico A', 'Custode', 'URL Documento fiscale']);
    }
    return ss.getId();
  } else {
    let newSs = SpreadsheetApp.create(safeName);
    let newFile = DriveApp.getFileById(newSs.getId());
    newFile.moveTo(folder);
    
    let defaultSheet = newSs.getSheetByName('Foglio1') || newSs.getSheetByName('Sheet1');
    if (defaultSheet) {
      defaultSheet.setName('MATERIALI');
      defaultSheet.appendRow(['Timestamp', 'CIP Inseritore', 'Categoria', 'Descrizione', 'Quantità', 'Costo', 'Data Acquisto', 'Acquistato Da', 'In Carico A', 'Custode', 'URL Documento fiscale']);
    } else {
      let s = newSs.insertSheet('MATERIALI');
      s.appendRow(['Timestamp', 'CIP Inseritore', 'Categoria', 'Descrizione', 'Quantità', 'Costo', 'Data Acquisto', 'Acquistato Da', 'In Carico A', 'Custode', 'URL Documento fiscale']);
    }
    return newSs.getId();
  }
}

function getDatiMateriali(utenteLoggato) {
  try {
    let res = { materiali: [], utentiMat: [] };

    let ruoloStrLoggato = String(utenteLoggato.ruolo || "").toUpperCase();
    let isMat = ruoloStrLoggato.includes('MAT') && !ruoloStrLoggato.includes('MAT2');
    let isMat2 = ruoloStrLoggato.includes('MAT2');
    let isAdmin = ruoloStrLoggato.includes('AMMINISTRATORE');
    let provLoggato = utenteLoggato.provincia;

    let foglioCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    let datiCariche = foglioCariche.getDataRange().getValues();
    
    for(let i=1; i<datiCariche.length; i++) {
        let cip = String(datiCariche[i][15]).toUpperCase().trim();
        if(!cip) continue;
        let prov = String(datiCariche[i][6]).toUpperCase().trim();
        let addUtente = false;
        
        if (isAdmin || isMat) { addUtente = true; } 
        else if (isMat2) { if (prov === provLoggato) addUtente = true; }
        
        if(addUtente) {
            res.utentiMat.push({ cip: cip, nome: String(datiCariche[i][4]).toUpperCase().trim(), cognome: String(datiCariche[i][3]).toUpperCase().trim(), provincia: prov, ruolo: String(datiCariche[i][18]).toUpperCase() });
        }
    }

    let folder = DriveApp.getFolderById(FOLDER_MATERIALI_SHEETS_ID);
    let files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    
    while(files.hasNext()) {
        let file = files.next();
        let ss = SpreadsheetApp.openById(file.getId());
        let foglio = ss.getSheetByName('MATERIALI');
        if (!foglio) continue;
        
        let dati = foglio.getDataRange().getValues();

        for(let i=1; i<dati.length; i++) {
          if(!dati[i][0] || i === 0) continue;
          let acquistatoDa = String(dati[i][7]);
          let inCaricoA = String(dati[i][8] || "");
          
          let seesIt = false;
          if (isAdmin) { seesIt = true; }
          else if (isMat && inCaricoA.includes('Regionale')) { seesIt = true; }
          else if (isMat2 && inCaricoA.includes(provLoggato)) { seesIt = true; }

          if(seesIt) {
            let dAcquisto = dati[i][6];
            if (dAcquisto instanceof Date) dAcquisto = Utilities.formatDate(dAcquisto, "GMT+1", "yyyy-MM-dd");
            else if (String(dAcquisto).includes('/')) { let p = String(dAcquisto).split('/'); dAcquisto = `${p[2]}-${p[1]}-${p[0]}`; }

            res.materiali.push({ 
                fileId: file.getId(), riga: i+1, dataReg: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy"), 
                categoria: dati[i][2], descrizione: dati[i][3], quantita: dati[i][4], 
                costo: parseFloat(dati[i][5] || 0), dataAcquistoIso: dAcquisto, 
                acquistatoDa: acquistatoDa, inCaricoA: inCaricoA, custode: dati[i][9], fileUrl: dati[i][10] 
            });
          }
        }
    }
    
    res.materiali.sort((a, b) => new Date(b.dataAcquistoIso).getTime() - new Date(a.dataAcquistoIso).getTime());
    return res;
  } catch(e) { return { materiali: [], utentiMat: [] }; }
}

function salvaNuovoMateriale(dati, fileBase64, mimeType, fileName) {
  try {
    let rootFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
    let targetFolderName = dati.inCaricoA;
    let targetFolder;
    
    let folders = rootFolder.getFoldersByName(targetFolderName);
    if (folders.hasNext()) { targetFolder = folders.next(); } else { targetFolder = rootFolder.createFolder(targetFolderName); }

    let fileUrl = "";
    if (fileBase64 && fileName) {
      let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
      let uploadedFile = targetFolder.createFile(blob);
      fileUrl = uploadedFile.getUrl();
    }

    let p = dati.data.split('-'); 
    let dataCorretta = new Date(p[0], p[1] - 1, p[2], 12, 0, 0);

    let ssId = getOrCreateMaterialiFile(dati.inCaricoA);
    let foglio = SpreadsheetApp.openById(ssId).getSheetByName('MATERIALI');
    
    foglio.appendRow([ new Date(), dati.cip, dati.categoria, dati.descrizione, dati.quantita, dati.costo, dataCorretta, dati.acquistatoDa, dati.inCaricoA, dati.custode, fileUrl ]);
    
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaMateriale(fileId, riga) {
    try { 
        let sheet = SpreadsheetApp.openById(fileId).getSheetByName('MATERIALI');
        let segreteria = sheet.getRange(riga, 9).getValue();
        let fileUrl = sheet.getRange(riga, 11).getValue();
        
        if (fileUrl && String(fileUrl).trim() !== "") {
            try {
                let extractIdMatch = String(fileUrl).match(/[-\w]{25,}/);
                if (extractIdMatch && extractIdMatch[0]) {
                    let file = DriveApp.getFileById(extractIdMatch[0]);
                    let rootDocsFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
                    let segrFolders = rootDocsFolder.getFoldersByName(segreteria);
                    let segrFolder = segrFolders.hasNext() ? segrFolders.next() : rootDocsFolder.createFolder(segreteria);
                    
                    let cestinoName = "CESTINO_DOC_" + segreteria;
                    let cestinoFolders = segrFolder.getFoldersByName(cestinoName);
                    let cestinoFolder = cestinoFolders.hasNext() ? cestinoFolders.next() : segrFolder.createFolder(cestinoName);
                    
                    file.moveTo(cestinoFolder);
                }
            } catch(errFile) { /* Ignora se il file è già stato rimosso */ }
        }
        
        sheet.deleteRow(riga); 
        return { success: true }; 
    } catch(e) { return { success: false, error: e.message }; }
}

function eliminaDocumentoMateriale(fileId, riga, segreteria) {
  try {
    let sheet = SpreadsheetApp.openById(fileId).getSheetByName('MATERIALI');
    let fileUrl = sheet.getRange(riga, 11).getValue();
    
    if (fileUrl && String(fileUrl).trim() !== "") {
      try {
        let extractIdMatch = String(fileUrl).match(/[-\w]{25,}/);
        if (extractIdMatch && extractIdMatch[0]) {
          let file = DriveApp.getFileById(extractIdMatch[0]);
          let rootDocsFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
          let segrFolders = rootDocsFolder.getFoldersByName(segreteria);
          let segrFolder = segrFolders.hasNext() ? segrFolders.next() : rootDocsFolder.createFolder(segreteria);
          let cestinoName = "CESTINO_DOC_" + segreteria;
          let cestinoFolders = segrFolder.getFoldersByName(cestinoName);
          let cestinoFolder = cestinoFolders.hasNext() ? cestinoFolders.next() : segrFolder.createFolder(cestinoName);
          file.moveTo(cestinoFolder);
        }
      } catch(e) {}
    }
    
    sheet.getRange(riga, 11).clearContent();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function sostituisciDocumentoMateriale(fileId, riga, segreteria, fileBase64, mimeType, fileName) {
  try {
    let sheet = SpreadsheetApp.openById(fileId).getSheetByName('MATERIALI');
    let fileUrl = sheet.getRange(riga, 11).getValue();
    
    if (fileUrl && String(fileUrl).trim() !== "") {
      try {
        let extractIdMatch = String(fileUrl).match(/[-\w]{25,}/);
        if (extractIdMatch && extractIdMatch[0]) {
          let file = DriveApp.getFileById(extractIdMatch[0]);
          let rootDocsFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
          let segrFolders = rootDocsFolder.getFoldersByName(segreteria);
          let segrFolder = segrFolders.hasNext() ? segrFolders.next() : rootDocsFolder.createFolder(segreteria);
          let cestinoName = "CESTINO_DOC_" + segreteria;
          let cestinoFolders = segrFolder.getFoldersByName(cestinoName);
          let cestinoFolder = cestinoFolders.hasNext() ? cestinoFolders.next() : segrFolder.createFolder(cestinoName);
          file.moveTo(cestinoFolder);
        }
      } catch(e) {}
    }
    
    let rootDocsFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
    let segrFolders = rootDocsFolder.getFoldersByName(segreteria);
    let targetFolder = segrFolders.hasNext() ? segrFolders.next() : rootDocsFolder.createFolder(segreteria);
    
    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
    let uploadedFile = targetFolder.createFile(blob);
    let newFileUrl = uploadedFile.getUrl();
    
    sheet.getRange(riga, 11).setValue(newFileUrl);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Aggiunto per evitare errori dal JavaScript (salvaModificheMultipleMateriali mancante nel dump originale ma chiamato in JS)
function salvaModificheMultipleMateriali(modifiche, operatore) {
  try {
    modifiche.forEach(m => {
      let sheet = SpreadsheetApp.openById(m.fileId).getSheetByName('MATERIALI');
      sheet.getRange(m.riga, 7).setValue(m.data);
      sheet.getRange(m.riga, 3).setValue(m.categoria);
      sheet.getRange(m.riga, 4).setValue(m.descrizione);
      sheet.getRange(m.riga, 5).setValue(m.quantita);
      sheet.getRange(m.riga, 6).setValue(m.costo);
      sheet.getRange(m.riga, 8).setValue(m.acquistatoDa);
      sheet.getRange(m.riga, 9).setValue(m.inCaricoA);
      sheet.getRange(m.riga, 10).setValue(m.custode);
    });
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}