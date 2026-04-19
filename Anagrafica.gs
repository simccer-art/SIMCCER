// ================= LOGIN E UTENTE ================= //
function effettuaLogin(password) {
  const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE'); 
  const dati = foglio.getDataRange().getValues();
  let utenteLoggato = null;
  for (let i = 1; i < dati.length; i++) {
    if (dati[i][17] && dati[i][17] == password) { utenteLoggato = getDatiProfilo(dati[i]); break; }
  }
  if (!utenteLoggato) return { error: "Password non valida" };

  let menuItems = ["PERMESSI SINDACALI"]; 
  try {
    const foglioMenu = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('MENU');
    if (foglioMenu) {
       const datiMenu = foglioMenu.getRange("B2:B").getValues(); let vociGenerate = [];
       for (let i = 0; i < datiMenu.length; i++) { if (datiMenu[i][0]) vociGenerate.push(String(datiMenu[i][0]).toUpperCase().trim()); }
       if (vociGenerate.length > 0) menuItems = vociGenerate;
    }
  } catch(e) {}
  utenteLoggato.menu = menuItems; return utenteLoggato;
}

function refreshUtente(cip) {
  try {
      const dati = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE').getDataRange().getValues();
      for (let i = 1; i < dati.length; i++) {
        if (String(dati[i][15]).toUpperCase().trim() === cip.toUpperCase().trim()) { 
           let u = getDatiProfilo(dati[i]); let menuItems = ["PERMESSI SINDACALI"]; 
           try {
             const foglioMenu = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('MENU');
             if (foglioMenu) {
                const datiMenu = foglioMenu.getRange("B2:B").getValues(); let vociGenerate = [];
                for (let j = 0; j < datiMenu.length; j++) { if (datiMenu[j][0]) vociGenerate.push(String(datiMenu[j][0]).toUpperCase().trim()); }
                if (vociGenerate.length > 0) menuItems = vociGenerate;
             }
           } catch(e) {}
           u.menu = menuItems; return u;
        }
      }
      return null;
  } catch(e) { return null; }
}

// ================= PROFILO E FOTO ================= //
function getDatiProfilo(riga) {
  function formattaDataSicura(val) { if (!val) return ""; if (val instanceof Date) return Utilities.formatDate(val, "GMT+1", "dd/MM/yyyy"); return String(val); }
  let ruoloCompleto = String(riga[18]).toUpperCase().trim(); let ruoloBase = ruoloCompleto.split(',')[0].trim();
  let extraInfo = {cf: "", ibans: [], vetture: [], foto: ""};
  try { if(riga[20]) extraInfo = JSON.parse(riga[20]); } catch(e) {}

  return {
    success: true, grado: riga[2], cognome: String(riga[3]).toUpperCase().trim(), nome: String(riga[4]).toUpperCase().trim(), 
    carica: riga[5], provincia: riga[6], provServizio: String(riga[7]).toUpperCase(), reparto: riga[8], 
    dataElezione: formattaDataSicura(riga[9]), dataNomina: formattaDataSicura(riga[10]), 
    consigliere: String(riga[11]).toUpperCase().includes('CONSIGLIERE') ? 'SI' : 'NO', 
    segreteria: String(riga[12]).toUpperCase(), email: String(riga[13]).toLowerCase(), telefono: String(riga[14]), 
    cip: String(riga[15]).toUpperCase().trim(), ruolo: ruoloCompleto, ruoloBase: ruoloBase, 
    organizzazione: String(riga[19]).toUpperCase(), cf: extraInfo.cf || "", ibans: extraInfo.ibans || [], vetture: extraInfo.vetture || [],
    fotoUrl: extraInfo.foto || ""
  };
}

function salvaFotoProfiloDrive(cip, segreteria, fileBase64, mimeType, fileName) {
  try {
    let rootFolder = DriveApp.getFolderById('1yePAhWoiUMWgD4SlPYlnw5_UPmj-SDKm');
    let targetFolderName = segreteria || "Non Assegnato";
    let targetFolder;
    
    let folders = rootFolder.getFoldersByName(targetFolderName);
    if (folders.hasNext()) { targetFolder = folders.next(); } 
    else { targetFolder = rootFolder.createFolder(targetFolderName); }
    
    try {
      let oldFiles = targetFolder.getFilesByName(fileName);
      while (oldFiles.hasNext()) { oldFiles.next().setTrashed(true); }
    } catch(e) {}
    
    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
    let uploadedFile = targetFolder.createFile(blob);
    uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    let fileId = uploadedFile.getId();
    let directUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w500";
    
    let foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    let dati = foglio.getDataRange().getValues();
    for (let i = 1; i < dati.length; i++) {
      if (String(dati[i][15]).toUpperCase().trim() === cip.toUpperCase().trim()) {
        let extraInfo = {cf: "", ibans: [], vetture: [], foto: ""};
        try { if(dati[i][20]) extraInfo = JSON.parse(dati[i][20]); } catch(e) {}
        extraInfo.foto = directUrl; 
        foglio.getRange(i + 1, 21).setValue(JSON.stringify(extraInfo));
        break;
      }
    }
    return { success: true, url: directUrl };
  } catch (e) { return { success: false, error: e.message }; }
}

function cancellaVecchiaFoto(folder, fileName) {
  try {
    let files = folder.getFilesByName(fileName);
    while (files.hasNext()) { files.next().setTrashed(true); }
  } catch (e) { /* Silenzioso */ }
}

// ================= GESTIONE SEGRETARI E MODIFICHE ================= //
function getUtentiPerGestione(utenteLoggato) {
  try {
    const dati = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE').getDataRange().getValues();
    let lista = []; let liv = GERARCHIA[utenteLoggato.ruoloBase] || 1;
    let isSegrGenER = (liv === 2 && utenteLoggato.provincia === 'EMILIA ROMAGNA');
    if (liv < 2) return [];

    for (let i = 1; i < dati.length; i++) {
      let u = getDatiProfilo(dati[i]); u.ordineFile = i; 
      if (!u.cip) continue;
      if (isSegrGenER || liv >= 3) lista.push(u); 
      else if (liv === 2 && u.provincia === utenteLoggato.provincia) lista.push(u); 
    }
    return lista;
  } catch (e) { throw new Error(e.message); }
}

function salvaRichiestaModificaProfilo(cip, nomeCompleto, datiJsonStr) {
  try {
    let foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('MODIFICHE_PROFILO');
    if (!foglio) return { success: false, error: "Creare il foglio 'MODIFICHE_PROFILO'." };
    foglio.appendRow([new Date(), cip, nomeCompleto, datiJsonStr, "IN ATTESA", ""]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getRichiesteModificaProfilo(ruoloBase) {
  if (GERARCHIA[ruoloBase] < 5) return []; 
  try {
    let foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('MODIFICHE_PROFILO');
    if (!foglio) return [];
    let dati = foglio.getDataRange().getValues(); let richieste = [];
    for(let i=1; i<dati.length; i++) {
       if (dati[i][4] === "IN ATTESA") richieste.push({ riga: i+1, data: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy HH:mm"), cip: dati[i][1], nome: dati[i][2], datiJSON: dati[i][3] });
    }
    return richieste;
  } catch(e) { return []; }
}

function approvaModificaProfilo(rigaModifica, cip, datiNuoviJSON, azione) {
  try {
    let fMod = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('MODIFICHE_PROFILO');
    if(azione === "APPROVATO") {
       let fCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
       let cariche = fCariche.getDataRange().getValues(); let rigaTarget = -1;
       for(let i=1; i<cariche.length; i++) { if(String(cariche[i][15]).toUpperCase().trim() === cip.toUpperCase().trim()) { rigaTarget = i+1; break; } }
       if(rigaTarget > -1) {
          let dati = JSON.parse(datiNuoviJSON);
          if(dati.grado) fCariche.getRange(rigaTarget, 3).setValue(dati.grado);
          if(dati.provServizio) fCariche.getRange(rigaTarget, 8).setValue(dati.provServizio);
          if(dati.reparto) fCariche.getRange(rigaTarget, 9).setValue(dati.reparto);
          if(dati.email) fCariche.getRange(rigaTarget, 14).setValue(dati.email.toLowerCase());
          if(dati.telefono) fCariche.getRange(rigaTarget, 15).setValue(dati.telefono);
          
          let oldExtra = {cf: "", ibans: [], vetture: []};
          try { if(cariche[rigaTarget-1][20]) oldExtra = JSON.parse(cariche[rigaTarget-1][20]); } catch(e){}
          let newExtra = { cf: dati.cf !== undefined ? dati.cf : oldExtra.cf, ibans: dati.ibans !== undefined ? dati.ibans : oldExtra.ibans, vetture: dati.vetture !== undefined ? dati.vetture : oldExtra.vetture };
          fCariche.getRange(rigaTarget, 21).setValue(JSON.stringify(newExtra));
       } else return { success: false, error: "CIP dell'utente non trovato nel foglio CARICHE." };
    }
    fMod.getRange(rigaModifica, 5).setValue(azione);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}