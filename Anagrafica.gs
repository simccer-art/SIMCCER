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
    
    // Trova o crea la cartella della segreteria
    let folders = rootFolder.getFoldersByName(targetFolderName);
    if (folders.hasNext()) { targetFolder = folders.next(); } 
    else { targetFolder = rootFolder.createFolder(targetFolderName); }
    
    // CANCELLA LA VECCHIA FOTO SE ESISTE
    try {
      let oldFiles = targetFolder.getFilesByName(fileName);
      while (oldFiles.hasNext()) { oldFiles.next().setTrashed(true); }
    } catch(e) {}
    
    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
    let uploadedFile = targetFolder.createFile(blob);
    
    uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // CREA IL LINK THUMBNAIL (Sblocca la visione nel tondo HTML)
    let fileId = uploadedFile.getId();
    let directUrl = "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w500";
    
    // Salva nel database
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
    while (files.hasNext()) {
      let file = files.next();
      file.setTrashed(true); // Sposta nel cestino
    }
  } catch (e) { /* Silenzioso */ }
}
