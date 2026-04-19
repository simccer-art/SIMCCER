





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

// ================= SALVATAGGIO FOTO PROFILO ================= //
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

// TROVA E CANCELLA LA VECCHIA FOTO SE ESISTE PER NON OCCUPARE SPAZIO


// ================= MODULO GESTIONE SEGRETARI ================= //
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

// ================= MODIFICA PROFILO ================= //
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

// ================= MODULO CONTABILITA' ================= //
const FOLDER_CONTABILITA_ID = '1TUG3r9PUuPvwK5xK6lpLmTBE44M8wZhK';

// Questa funzione entra nella cartella, trova il file della segreteria (o lo crea) e lo restituisce
function getFoglioContabilita(segreteria) {
  let folder = DriveApp.getFolderById(FOLDER_CONTABILITA_ID);
  let fileName = "Contabilità_" + segreteria;
  let files = folder.getFilesByName(fileName);
  
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    // Se è la prima volta che questa segreteria usa la contabilità, le crea il file Excel
    ss = SpreadsheetApp.create(fileName);
    let file = DriveApp.getFileById(ss.getId());
    file.moveTo(folder); // Lo sposta nella cartella contabilità
  }
  
  // Inizializza le schede interne se mancano
  if (!ss.getSheetByName('IMPOSTAZIONI_CONTO')) { ss.insertSheet('IMPOSTAZIONI_CONTO').appendRow(['Segreteria', 'Istituto', 'IBAN', 'CF']); }
  if (!ss.getSheetByName('CAUSALI')) { ss.insertSheet('CAUSALI').appendRow(['Segreteria', 'Causale']); }
  if (!ss.getSheetByName('ORDINANTI')) { ss.insertSheet('ORDINANTI').appendRow(['Segreteria', 'Nominativo']); }
  if (!ss.getSheetByName('MOVIMENTI')) { 
    ss.insertSheet('MOVIMENTI').appendRow(['Timestamp', 'Segreteria', 'DataMovimento', 'Tipo', 'Causale', 'Dettaglio', 'Importo', 'Operatore', 'Ordinante', 'URL Documento']); 
  }
  
  return ss;
}

function getDatiContabilita(segreteria) {
  try {
    let ss = getFoglioContabilita(segreteria);
    
    let foglioImp = ss.getSheetByName('IMPOSTAZIONI_CONTO');
    let datiImpostazioni = foglioImp ? foglioImp.getDataRange().getValues() : [];
    let impostazioni = { istituto: "", iban: "", cf: "" };
    for(let i=1; i<datiImpostazioni.length; i++) { 
      if(String(datiImpostazioni[i][0]).trim() === segreteria) { 
        impostazioni = { istituto: String(datiImpostazioni[i][1] || ""), iban: String(datiImpostazioni[i][2] || ""), cf: String(datiImpostazioni[i][3] || "") };
        break; 
      } 
    }

    let foglioCau = ss.getSheetByName('CAUSALI');
    let datiCausali = foglioCau ? foglioCau.getDataRange().getValues() : []; 
    let causali = [];
    for(let i=1; i<datiCausali.length; i++) { 
      if(String(datiCausali[i][0]).trim() === segreteria) causali.push(String(datiCausali[i][1] || ""));
    }

    let foglioOrd = ss.getSheetByName('ORDINANTI');
    let datiOrdinanti = foglioOrd ? foglioOrd.getDataRange().getValues() : []; 
    let ordinanti = [];
    for(let i=1; i<datiOrdinanti.length; i++) { 
      if(String(datiOrdinanti[i][0]).trim() === segreteria) ordinanti.push(String(datiOrdinanti[i][1] || ""));
    }

    let foglioMov = ss.getSheetByName('MOVIMENTI');
    let datiMovimenti = foglioMov ? foglioMov.getDataRange().getValues() : []; 
    let movimenti = [];
    for(let i=1; i<datiMovimenti.length; i++) {
      if(String(datiMovimenti[i][1]).trim() === segreteria) {
        let dStr = datiMovimenti[i][2];
        let dMov = (dStr instanceof Date) ? dStr : new Date(dStr);
        if(isNaN(dMov.getTime())) dMov = new Date();

        let isoDate = dMov.getFullYear() + "-" + String(dMov.getMonth()+1).padStart(2,'0') + "-" + String(dMov.getDate()).padStart(2,'0');
        movimenti.push({ 
            riga: i+1, 
            dataObj: dMov.getTime(), 
            dataMovimento: Utilities.formatDate(dMov, "GMT+1", "dd/MM/yyyy"), 
            dataIso: isoDate,
            mese: dMov.getMonth() + 1, 
            anno: dMov.getFullYear(),
            tipo: String(datiMovimenti[i][3] || ""), 
            causale: String(datiMovimenti[i][4] || ""), 
            dettaglio: String(datiMovimenti[i][5] || ""), 
            importo: parseFloat(datiMovimenti[i][6]) || 0, 
            operatore: String(datiMovimenti[i][7] || ""), 
            ordinante: String(datiMovimenti[i][8] || ""),
            fileUrl: String(datiMovimenti[i][9] || "")
        });
      }
    }
    movimenti.sort((a,b) => b.dataObj - a.dataObj);
    return { error: null, impostazioni: impostazioni, causali: causali, ordinanti: ordinanti, movimenti: movimenti };
  } catch(e) {
    return { error: "Errore Script: " + e.message, impostazioni: {istituto:"",iban:"",cf:""}, causali: [], ordinanti: [], movimenti: [] };
  }
}

// ================= GESTIONE IMPOSTAZIONI E TENDINE CONTABILITA' ================= //
function salvaImpostazioniContoContabilita(segreteria, istituto, iban, cf) {
  try {
    let ss = getFoglioContabilita(segreteria);
    let sheet = ss.getSheetByName('IMPOSTAZIONI_CONTO');
    let data = sheet.getDataRange().getValues(); 
    let found = false;
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === segreteria) { 
        sheet.getRange(i+1, 2).setValue(istituto);
        sheet.getRange(i+1, 3).setValue(iban); 
        sheet.getRange(i+1, 4).setValue(cf); 
        found = true; 
        break; 
      }
    }
    if(!found) sheet.appendRow([segreteria, istituto, iban, cf]);
    SpreadsheetApp.flush(); 
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function aggiungiCausaleContabilita(segreteria, causale) {
  try {
    let ss = getFoglioContabilita(segreteria);
    let sheet = ss.getSheetByName('CAUSALI');
    sheet.appendRow([segreteria, causale]); 
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function aggiungiOrdinanteContabilita(segreteria, nominativo) {
  try {
    let ss = getFoglioContabilita(segreteria);
    let sheet = ss.getSheetByName('ORDINANTI');
    sheet.appendRow([segreteria, nominativo]); 
    SpreadsheetApp.flush();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaCausaleContabilita(segreteria, causale) {
  try {
    let ss = getFoglioContabilita(segreteria);
    let sheet = ss.getSheetByName('CAUSALI');
    let data = sheet.getDataRange().getValues();
    for(let i = data.length - 1; i >= 1; i--) {
      if(data[i][0] === segreteria && data[i][1] === causale) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: "Causale non trovata nel database." };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaOrdinanteContabilita(segreteria, nominativo) {
  try {
    let ss = getFoglioContabilita(segreteria);
    let sheet = ss.getSheetByName('ORDINANTI');
    let data = sheet.getDataRange().getValues();
    for(let i = data.length - 1; i >= 1; i--) {
      if(data[i][0] === segreteria && data[i][1] === nominativo) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true };
      }
    }
    return { success: false, error: "Nominativo non trovato nel database." };
  } catch(e) { return { success: false, error: e.message }; }
}

function salvaMovimentoContabilita(segreteria, dataMov, tipo, causale, dettaglio, importo, operatore, ordinante, fileBase64, mimeType, fileName) {
  try {
    let ss = getFoglioContabilita(segreteria);
    let sheet = ss.getSheetByName('MOVIMENTI');

    let fileUrl = "";
    if (fileBase64 && String(fileBase64).trim() !== "" && fileName && String(fileName).trim() !== "") {
      try {
        let folder;
        try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }
        let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType || 'application/pdf', "Contab_" + new Date().getTime() + "_" + fileName);
        let uploadedFile = folder.createFile(blob);
        fileUrl = uploadedFile.getUrl();
      } catch(driveErr) {
        fileUrl = "Errore salvataggio file: " + driveErr.message;
      }
    }

    let parts = String(dataMov).split('-');
    let d = new Date(parts[0], parts[1]-1, parts[2]);

    sheet.appendRow([new Date(), segreteria, d, tipo, causale, dettaglio, importo, operatore, ordinante || "", fileUrl]);
    SpreadsheetApp.flush(); 
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ATTENZIONE: Aggiunto parametro "segreteria" per dire al server quale file aprire per cancellare
function eliminaMovimentoContabilita(segreteria, riga) {
  try { 
    let ss = getFoglioContabilita(segreteria);
    ss.getSheetByName('MOVIMENTI').deleteRow(riga); 
    return { success: true };
  } 
  catch(e) { return { success: false, error: e.message }; }
}

function salvaModificheMultipleContabilita(segreteria, modifiche, operatore) {
  try {
      let ss = getFoglioContabilita(segreteria);
      let sheet = ss.getSheetByName('MOVIMENTI');
      modifiche.forEach(m => {
          let parts = m.data.split('-'); let d = new Date(parts[0], parts[1]-1, parts[2]);
          sheet.getRange(m.riga, 3).setValue(d);
          sheet.getRange(m.riga, 4).setValue(m.tipo);
          sheet.getRange(m.riga, 5).setValue(m.causale);
          sheet.getRange(m.riga, 6).setValue(m.dettaglio);
          sheet.getRange(m.riga, 7).setValue(m.importo);
          sheet.getRange(m.riga, 8).setValue(operatore); 
          sheet.getRange(m.riga, 9).setValue(m.ordinante || "");
      });
      SpreadsheetApp.flush();
      return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getLogoBase64(segreteria) {
  try {
      let folder = DriveApp.getFolderById('1ewdYZ5F_o-dW5SvOOtNVF978Y9gi5yba');
      let provName = segreteria.replace("Provinciale (", "").replace("Regionale (", "").replace(")", "").trim();
      let files = folder.searchFiles("title contains '" + provName + "'");
      if (files.hasNext()) { let f = files.next(); return `data:${f.getMimeType()};base64,${Utilities.base64Encode(f.getBlob().getBytes())}`; }
      let allFiles = folder.getFiles();
      if (allFiles.hasNext()) { let f = allFiles.next(); return `data:${f.getMimeType()};base64,${Utilities.base64Encode(f.getBlob().getBytes())}`; }
  } catch(e) {}
  return "";
}

function generaPDFContabilita(segreteria, mese, anno, operatore, dataDa, dataA) {
  try {
    let dati = getDatiContabilita(segreteria);
    
    let movimenti = dati.movimenti.filter(m => {
        let matchMese = (mese === 'TUTTI' || String(m.mese) === mese);
        let matchAnno = (anno === 'TUTTI' || String(m.anno) === anno);
        let matchDa = true; let matchA = true;
        if (dataDa) matchDa = m.dataObj >= new Date(dataDa).getTime();
        if (dataA) matchA = m.dataObj <= new Date(dataA).setHours(23,59,59,999);
        return matchMese && matchAnno && matchDa && matchA;
    });
    
    movimenti.sort((a,b) => a.dataObj - b.dataObj);
    
    let htmlRows = ""; let prog = 1; let totIn = 0; let totOut = 0;
    movimenti.forEach(m => {
        if(m.tipo.includes('ENTRATA')) totIn += m.importo; else totOut += m.importo;
        let segno = m.tipo.includes('ENTRATA') ? '+' : '-'; let col = m.tipo.includes('ENTRATA') ? 'green' : 'red';
        htmlRows += `<tr><td style="text-align:center;">${prog++}</td><td>${m.dataMovimento}</td><td style="color:${col};">${m.tipo}</td><td>${m.ordinante || ''}</td><td>${m.causale}</td><td>${m.dettaglio}</td><td style="text-align:right; color:${col};">${segno} € ${m.importo.toFixed(2)}</td></tr>`;
    });

    let base64Logo = getLogoBase64(segreteria);
    let imgTag = base64Logo ? `<div style="text-align:center; margin-bottom:20px;"><img src="${base64Logo}" style="max-height:100px;"></div>` : '';

    let parsedSegr = "";
    if (segreteria.startsWith("Provinciale")) { parsedSegr = "Segreteria Provinciale di " + segreteria.replace("Provinciale (", "").replace(")", ""); } 
    else if (segreteria.startsWith("Regionale")) { parsedSegr = "Segreteria Regionale " + segreteria.replace("Regionale (", "").replace(")", ""); } 
    else { parsedSegr = "Segreteria Nazionale"; }

    function formattaD(dIso) { if(!dIso) return ""; let p = dIso.split('-'); return p[2]+"/"+p[1]+"/"+p[0]; }
    
    // --- TRADUZIONE NUMERO MESE IN TESTO ---
    const nomiMesi = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
    let nomeMeseTesto = (mese === "TUTTI") ? "TUTTI I MESI" : nomiMesi[parseInt(mese)];
    
    let strPeriodo = `Mese: ${nomeMeseTesto} - Anno: ${anno}`;
    if (dataDa || dataA) {
       strPeriodo = `Dal: ${dataDa ? formattaD(dataDa) : 'Inizio'} al: ${dataA ? formattaD(dataA) : 'Fine'}`;
    }

    let html = `
      <style>@page { margin: 1cm; }</style>
      <div style="font-family: Arial, sans-serif; font-size:10px;">
         ${imgTag}
         <div style="text-align:center; margin-bottom:15px;">
            <h3 style="margin:5px 0 0 0; font-size:16px; font-weight:normal;">Sindacato Italiano Militari Carabinieri</h3>
            <h4 style="margin:0; font-size:12px;">${parsedSegr}</h4>
         </div>
         
         <h2 style="text-align:center; color:#1a4b84; margin-top:20px; margin-bottom:5px;">REGISTRO CONTABILE</h2>
         <p style="text-align:center; color:#1a4b84; font-weight:bold; margin-top:0;">Periodo di riferimento: ${strPeriodo}</p>
         
         <div style="margin-bottom:15px; padding:10px; background:#f4f6f8; border:1px solid #ccc; font-size:10px;">
            <p style="margin:0;"><strong>Istituto:</strong> ${dati.impostazioni.istituto || "-"} | <strong>IBAN:</strong> ${dati.impostazioni.iban || "-"} | <strong>C.F.:</strong> ${dati.impostazioni.cf || "-"}</p>
         </div>
         
         <table style="width:100%; border-collapse:collapse; font-size:9px;" border="1" cellpadding="3">
            <tr style="background:#eee;"><th>N.</th><th>Data</th><th>Tipo</th><th>Ordinante/Beneficiario</th><th>Causale</th><th>Dettaglio</th><th>Importo</th></tr>
            ${htmlRows}
         </table>
         <br>
         <table style="width:50%; margin-left:auto; border-collapse:collapse; font-size:10px;" border="1" cellpadding="4">
            <tr><td><strong>Totale Entrate:</strong></td><td style="color:green; text-align:right;">€ ${totIn.toFixed(2)}</td></tr>
            <tr><td><strong>Totale Uscite:</strong></td><td style="color:red; text-align:right;">€ ${totOut.toFixed(2)}</td></tr>
            <tr><td style="background:#fff3cd;"><strong>SALDO PERIODO:</strong></td><td style="text-align:right; background:#fff3cd;"><strong>€ ${(totIn - totOut).toFixed(2)}</strong></td></tr>
         </table>
         
         <div style="text-align:right; margin-top:40px; line-height:0.8;">
            <p style="font-size:12px; margin: 2px 0;">IL TESORIERE SIM CC</p>
            <p style="font-size:12px; margin: 2px 0;"><em>${operatore}</em></p>
            <p style="color:red; font-size:9px; margin: 2px 0; margin-top:10px;"><strong>[ DOCUMENTO VERIFICATO E FIRMATO ]</strong></p>
         </div>
      </div>
    `;
    let outputHtml = HtmlService.createHtmlOutput(html);
    let cartella; try { cartella = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { cartella = DriveApp.getRootFolder(); }
    let pdfFile = cartella.createFile(outputHtml.getAs('application/pdf').setName("Registro_Contabile_" + segreteria.replace(/\s+/g,"_") + "_" + new Date().getTime() + ".pdf"));
    return {success: true, url: pdfFile.getUrl()};
  } catch(e) { return {success: false, error: e.message}; }
}

// Funzione core che genera l'HTML del rimborso con Firma e Timbro opzionali



function salvaNuovoRimborso(dati, fileBase64, mimeType, fileName) {
  try {
    let folder; 
    try { 
      folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); 
    } catch(e) { 
      folder = DriveApp.getRootFolder(); 
    }
    
    // Decodifica e salva il file firmato
    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, "Firmato_" + fileName);
    let uploadedFile = folder.createFile(blob);
    let fileUrl = uploadedFile.getUrl();

    // Salva i dati nel foglio RICHIESTE_RIMBORSI
    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    
    // Se il foglio non esiste, lo crea con le intestazioni
    if (!foglio) {
      foglio = ss.insertSheet('RICHIESTE_RIMBORSI');
      foglio.appendRow(["Timestamp", "CIP", "Richiedente", "Provincia", "Dati JSON", "URL File", "Stato", "Note", "Operatore"]);
    }

    foglio.appendRow([ 
      new Date(), 
      dati.cip, 
      dati.richiedente, 
      dati.provincia, 
      JSON.stringify(dati), 
      fileUrl, 
      "IN ATTESA", 
      "", 
      "" 
    ]);

    return { success: true };
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

function getRimborsiUtente(cip) {
  try {
    let foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_RIMBORSI');
    let dati = foglio.getDataRange().getValues(); let lista = [];
    for(let i=1; i<dati.length; i++) {
        if(String(dati[i][1]).toUpperCase().trim() === cip.toUpperCase().trim()) {
            let json = JSON.parse(dati[i][4]);
            let importo = json.totaleGenerale !== undefined ? json.totaleGenerale : (json.importo || 0);
            let descrizione = json.motivo || json.descrizione || "Rimborso spese multiple";
            
            lista.push({ 
                riga: i+1, 
                dataInvio: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy HH:mm"), 
                dataSpesa: json.dataOdierna || json.dataSpesa, 
                importo: parseFloat(importo).toFixed(2), 
                importoAutorizzato: json.importoAutorizzato !== undefined ? parseFloat(json.importoAutorizzato).toFixed(2) : parseFloat(importo).toFixed(2), 
                descrizione: descrizione, 
                fileUrl: dati[i][5], 
                stato: dati[i][6], 
                note: dati[i][7], 
                operatore: dati[i][8], // <-- Aggiunto l'operatore
                protocollo: json.protocollo || '-', 
                segreteriaPagante: json.segreteriaPagante || '-', 
                jsonStr: dati[i][4] 
            });
        }
    }
    return lista.reverse();
  } catch(e) { return []; }
}

function getRimborsiGestione(utenteLoggato) {
  try {
    let foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_RIMBORSI');
    if(!foglio) return {daGestire: [], storico: []};
    let dati = foglio.getDataRange().getValues();
    
    let liv = GERARCHIA[utenteLoggato.ruoloBase] || 1;
    let ruoloStr = String(utenteLoggato.ruolo || "").toUpperCase();
    
    // Separazione netta dei ruoli Tesoriere
    let isTesoriere1 = ruoloStr.includes('TESORIERE') && !ruoloStr.includes('TESORIERE2'); // Solo Regionale
    let isTesoriere2 = ruoloStr.includes('TESORIERE2'); // Solo Provinciale
    let provLoggato = utenteLoggato.provincia;

    let daGestire = []; let storico = [];
    for(let i=1; i<dati.length; i++) {
        if(!dati[i][0]) continue;
        let cip = String(dati[i][1]).trim();
        let richiedente = dati[i][2];
        let prov = dati[i][3];
        let json = JSON.parse(dati[i][4]);
        let fileUrl = dati[i][5];
        let stato = String(dati[i][6]).toUpperCase().trim();
        let note = dati[i][7];
        let operatore = dati[i][8];

        let importo = json.totaleGenerale !== undefined ? json.totaleGenerale : (json.importo || 0);
        let descrizione = json.motivo || json.descrizione || "Rimborso spese multiple";
        
        let sp = json.segreteriaPagante || "";
        let isProvinciale = sp.startsWith("Provinciale");
        let isRegionale = sp.startsWith("Regionale") || sp === "Nazionale";

        // Estrae la provincia esatta da chi deve pagare (es. "FERRARA" da "Provinciale (FERRARA)")
        let targetProv = prov; 
        if (isProvinciale) {
            let match = sp.match(/\(([^)]+)\)/);
            if (match) targetProv = match[1].toUpperCase().trim();
        }

        // LOGICA DI AUTORIZZAZIONE E PAGAMENTO SEPARATA
        let seesIt = false; let canApprove = false; let canPay = false;
        
        if (liv === 5) { 
            seesIt = true; canApprove = true; canPay = true; 
        } else {
            if (isProvinciale && targetProv === provLoggato) {
                // SEGRGEN (liv >= 2) approva per la SUA provincia
                if (liv >= 2) { seesIt = true; canApprove = true; } 
                // SOLO il TESORIERE2 della SUA provincia Paga
                if (isTesoriere2) { seesIt = true; canPay = true; }  
            } else if (isRegionale) {
                // A livello regionale approva il liv 3 (o il SEGRGEN dell'Emilia Romagna)
                if (liv >= 3 || (liv === 2 && provLoggato === 'EMILIA ROMAGNA')) { seesIt = true; canApprove = true; } 
                // SOLO il TESORIERE (Regionale) Paga
                if (isTesoriere1) { seesIt = true; canPay = true; }  
            }
        }
        
        if(!seesIt) continue;

        let p = { riga: i+1, dataInvio: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy"), cip: cip, richiedente: richiedente, provincia: prov, dataSpesa: json.dataOdierna || json.dataSpesa, importo: parseFloat(importo).toFixed(2), importoAutorizzato: json.importoAutorizzato !== undefined ? parseFloat(json.importoAutorizzato).toFixed(2) : parseFloat(importo).toFixed(2), descrizione: descrizione, fileUrl: fileUrl, stato: stato, note: note, operatore: operatore, iban: json.iban, vettura: json.vettura, segreteriaPagante: sp || "-", protocollo: json.protocollo || "-", canApprove: canApprove, canPay: canPay };
        
        if (stato === 'IN ATTESA' && canApprove) daGestire.push(p);
        else if (stato === 'APPROVATO' && canPay) daGestire.push(p);
        else if (stato !== 'IN ATTESA ALLEGATO') storico.push(p);
    }
    return { daGestire: daGestire.reverse(), storico: storico.reverse() };
  } catch(e) { return {daGestire:[], storico:[]}; }
}

// Costanti aggiornate
const TEMPLATE_RIMBORSO_DOC_ID = '1ADoeXnFhtelnm9ZiPNL_1Xk02iXVRWjEiedFBW9H8s4';
const CARTELLA_LOGHI_ID = '1ewdYZ5F_o-dW5SvOOtNVF978Y9gi5yba';

/**
 * Funzione principale per compilare il Google Doc dei Rimborsi
 */
function compilaTemplateRimborsoDoc(payload, mostraFirma, mostraTimbroPagato) {
  try {
    let folder;
    try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }
    
    let safeProt = payload.protocollo ? payload.protocollo.replace(/\//g, '-') : "SenzaProt";
    let nomeFile = "Richiesta_Rimborso_" + safeProt + "_" + payload.cip;
    
    let copiaId = DriveApp.getFileById(TEMPLATE_RIMBORSO_DOC_ID).makeCopy(nomeFile, folder).getId();
    let doc = DocumentApp.openById(copiaId);
    let body = doc.getBody();

    // 2. Gestione LOGO ({{LOGO}})
    try {
      let logoFolder = DriveApp.getFolderById(CARTELLA_LOGHI_ID);
      let provName = payload.provincia.replace("Provinciale (", "").replace("Regionale (", "").replace(")", "").trim();
      let files = logoFolder.searchFiles("title contains '" + provName + "'");
      let logoBlob = null;
      
      if (files.hasNext()) {
        logoBlob = files.next().getBlob();
      } else {
        let allLogos = logoFolder.getFiles();
        if (allLogos.hasNext()) logoBlob = allLogos.next().getBlob();
      }

      if (logoBlob) {
        let elementoLogo = body.findText("{{LOGO}}");
        if (elementoLogo) {
          let im = elementoLogo.getElement().getParent().asParagraph().appendInlineImage(logoBlob);
          let origW = im.getWidth();
          let origH = im.getHeight();
          let maxW = 200; 
          let maxH = 80;  
          
          if (origW > 0 && origH > 0) {
            let ratio = origW / origH;
            if (origW > maxW || origH > maxH) {
              if (ratio > (maxW / maxH)) {
                im.setWidth(maxW); im.setHeight(Math.round(maxW / ratio));
              } else {
                im.setHeight(maxH); im.setWidth(Math.round(maxH * ratio));
              }
            }
          } else {
            im.setWidth(150).setHeight(75);
          }
          body.replaceText("{{LOGO}}", ""); 
        }
      }
    } catch (e) { console.log("Errore logo: " + e.message); }

    // 3. Gestione TAG {{SEGRETERIA}}
    let sp = payload.segreteriaPagante || "";
    let testoSegreteria = "";
    function toTitleCase(str) { return str.toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' '); }
    
    if (sp.includes("Provinciale")) {
      let prov = sp.match(/\(([^)]+)\)/);
      testoSegreteria = "Segreteria Provinciale di " + (prov ? toTitleCase(prov[1]) : "");
    } else if (sp.includes("Regionale")) {
      let reg = sp.match(/\(([^)]+)\)/);
      testoSegreteria = "Segreteria Regionale " + (reg ? toTitleCase(reg[1]) : "");
    } else { 
      testoSegreteria = toTitleCase(sp); 
    }
    body.replaceText("{{SEGRETERIA}}", "Sindacato Italiano Militari Carabinieri\n" + testoSegreteria);

    // 4. Campi Anagrafici e Mezzo
    body.replaceText("{{NOME_COGNOME}}", payload.richiedente || "");
    body.replaceText("{{PROTOCOLLO}}", payload.protocollo || "");
    body.replaceText("{{DATA_PROT}}", payload.dataOdierna || "");
    body.replaceText("{{CF}}", payload.cf || "");
    body.replaceText("{{MOTIVO}}", (payload.motivo || "").toUpperCase());
    body.replaceText("{{IBAN}}", payload.iban || "");
    
    let mezzo = payload.vettura || "";
    let marca = "", modello = "", targa = "";
    if (mezzo.includes("-")) {
      let partiMezzo = mezzo.split("-");
      targa = partiMezzo[1].trim().toUpperCase();
      let marcaModello = partiMezzo[0].trim().split(" ");
      marca = marcaModello[0].toUpperCase();
      modello = marcaModello.slice(1).join(" ").toUpperCase();
    }
    body.replaceText("{{MARCA}}", marca);
    body.replaceText("{{MODELLO}}", modello);
    body.replaceText("{{TARGA}}", targa);

    // 5. Mappatura TRATTE
    for (let i = 1; i <= 6; i++) {
      let tratta = payload.tratte[i - 1];
      body.replaceText("{{DATA_" + i + "}}", tratta ? tratta.data : "");
      body.replaceText("{{ITIN_" + i + "}}", tratta ? String(tratta.itinerario).toUpperCase() : "");
      body.replaceText("{{KM_" + i + "}}", tratta ? tratta.km : "");
    }
    body.replaceText("{{TOT_KM}}", payload.totaleKm || "0");
    body.replaceText("{{IMP_A}}", parseFloat(payload.importoA).toFixed(2));

    // 6. Mappatura SPESE
    let mappaSpese = { "Treno / Aereo / Nave": 7, "Pedaggio Autostradale": 8, "Taxi / Trasporti Locali": 8, "Vitto (Pasti)": 9, "Alloggio (Hotel)": 10, "Spese di rappresentanza": 11, "Spese telefoniche/postali": 12, "Altro": 13 };
    let speseRaggruppate = {};
    (payload.spese || []).forEach(s => {
      let indice = mappaSpese[s.categoria] || 13;
      if (!speseRaggruppate[indice]) speseRaggruppate[indice] = { data: [], doc: [], importoTotale: 0 };
      speseRaggruppate[indice].data.push(s.data);
      speseRaggruppate[indice].doc.push(s.doc ? String(s.doc).toUpperCase() : "");
      speseRaggruppate[indice].importoTotale += parseFloat(s.importo) || 0;
    });

    for (let j = 7; j <= 13; j++) {
      if (speseRaggruppate[j]) {
        body.replaceText("{{DATA_V" + j + "}}", speseRaggruppate[j].data.join("\n"));
        body.replaceText("{{DOC" + j + "}}", speseRaggruppate[j].doc.join("\n"));
        body.replaceText("{{SPESE" + j + "}}", "€ " + speseRaggruppate[j].importoTotale.toFixed(2));
      } else {
        body.replaceText("{{DATA_V" + j + "}}", ""); body.replaceText("{{DOC" + j + "}}", ""); body.replaceText("{{SPESE" + j + "}}", "");
      }
    }

    body.replaceText("{{IMP_B}}", parseFloat(payload.importoB).toFixed(2));
    body.replaceText("{{TOT_GEN}}", parseFloat(payload.totaleGenerale).toFixed(2));

    // Logica Importo Autorizzato e Timbro Arancione
    if (mostraTimbroPagato) {
      let impAut = (payload.importoAutorizzato !== undefined && payload.importoAutorizzato !== null && payload.importoAutorizzato !== "") 
                   ? parseFloat(payload.importoAutorizzato) 
                   : parseFloat(payload.totaleGenerale);
      body.replaceText("{{IMP_AUTORIZZATO}}", "€ " + impAut.toFixed(2));

      // --- ESTRAZIONE PROVINCIA ---
      let sp = payload.segreteriaPagante || "";
      let luogoStr = "";
      let match = sp.match(/\(([^)]+)\)/);
      if (match) {
          luogoStr = match[1].toUpperCase();
          if (luogoStr === "EMILIA ROMAGNA") luogoStr = "BOLOGNA"; 
      } else {
          luogoStr = "ROMA"; 
      }

      let dataOggiStr = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
      let tesoriereNome = payload.nomeTesoriere ? payload.nomeTesoriere.toUpperCase() : "________________";

      // Testo su 4 righe
      let timbroText = "PAGAMENTO AUTORIZZATO\n" +
                       "EURO " + impAut.toFixed(2) + "\n" +
                       "IL TESORIERE " + tesoriereNome + "\n" +
                       luogoStr + ", " + dataOggiStr;

      // Ricerca o creazione del punto in cui inserire il timbro
      let rangeTimbro = body.findText("{{TIMBRO_ARANCIONE}}");
      let table;
      if (rangeTimbro) {
          let el = rangeTimbro.getElement();
          
          // Cerca se si trova dentro una cella di tabella
          let container = el.getParent();
          while (container && container.getType() !== DocumentApp.ElementType.TABLE_CELL && container.getType() !== DocumentApp.ElementType.BODY_SECTION) {
              container = container.getParent();
          }
          
          // Svuota solo il testo del segnaposto senza eliminare le righe
          el.asText().replaceText("{{TIMBRO_ARANCIONE}}", "");
          
          // Se è dentro la tua cella di destra, crea il timbro LÌ DENTRO
          if (container && container.getType() === DocumentApp.ElementType.TABLE_CELL) {
              table = container.asTableCell().appendTable([ [""] ]);
          } else {
              // Se è libero nel foglio, lo accoda sotto
              let par = el.getParent();
              if (par.getType() === DocumentApp.ElementType.TEXT) par = par.getParent();
              let childIndex = body.getChildIndex(par);
              table = body.insertTable(childIndex + 1, [ [""] ]);
          }
      } else {
          body.appendParagraph(""); 
          table = body.appendTable([ [""] ]);
      }

      // --- APPLICAZIONE DIMENSIONI RIGIDE E STILE ---
      table.setBorderColor("#FF8C00"); 
      table.setBorderWidth(1); 
      
      table.setColumnWidth(0, 155.90); // LARGHEZZA 5,5 cm
      
      let row = table.getRow(0);
      row.setMinimumHeight(42.52); // ALTEZZA 1,5 cm
      
      let cell = row.getCell(0);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      
      // Padding minimo
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(2);
      
      cell.clear(); 
      let par = cell.appendParagraph(timbroText);
      par.setAlignment(DocumentApp.HorizontalAlignment.LEFT); 
      par.setLineSpacing(1.0); 
      par.setSpacingBefore(0); 
      par.setSpacingAfter(0);  
      
      let textEl = par.editAsText();
      textEl.setBold(true);
      textEl.setForegroundColor("#FF8C00"); 
      textEl.setFontSize(7.5); 

    } else {
      // Se non stiamo pagando, si limita a far sparire i testi senza cancellare nulla
      body.replaceText("{{IMP_AUTORIZZATO}}", ""); 
      body.replaceText("{{TIMBRO_ARANCIONE}}", ""); 
    }

    // 7. Firma e Pagamento (Rimossa come richiesto)
    body.replaceText("{{FIRMA_PAGANTE}}", "");

    if (mostraTimbroPagato) {
      body.replaceText("{{DATA_PAGAMENTO}}", Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"));
      body.replaceText("{{LUOGO_PAGAMENTO}}", "Bologna (BO)");
    } else {
      body.replaceText("{{DATA_PAGAMENTO}}", ""); body.replaceText("{{LUOGO_PAGAMENTO}}", "");
    }

    doc.saveAndClose();
    let pdfBlob = DriveApp.getFileById(copiaId).getAs('application/pdf');
    let pdfFile = folder.createFile(pdfBlob);
    DriveApp.getFileById(copiaId).setTrashed(true);
    return { success: true, url: pdfFile.getUrl() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// AGGIORNAMENTO DELLE FUNZIONI ESISTENTI:

function generaPDFRimborsoDaFirmare(payload) {
  return compilaTemplateRimborsoDoc(payload, false, false);
}

function generaTimbroEPagaRimborso(datiRimbInput, operatore, importoAutorizzato) {
  try {
    let datiRimb = typeof datiRimbInput === 'string' ? JSON.parse(datiRimbInput) : datiRimbInput;
    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    let riga = datiRimb.riga;
    
    let datiOriginali = JSON.parse(foglio.getRange(riga, 5).getValue());
    
    // Se il tesoriere indica un importo in fase di pagamento, sovrascriviamo quello precedente
    if (importoAutorizzato !== undefined && importoAutorizzato !== null) {
      datiOriginali.importoAutorizzato = importoAutorizzato;
    }
    
    // --- AGGIUNGI QUESTA RIGA ---
    datiOriginali.nomeTesoriere = operatore;
    
    foglio.getRange(riga, 5).setValue(JSON.stringify(datiOriginali));

    let res = compilaTemplateRimborsoDoc(datiOriginali, true, true);
    if (res.success) {
      foglio.getRange(riga, 7).setValue("PAGATO");
      foglio.getRange(riga, 8).setValue("Pagato dal Tesoriere");
      foglio.getRange(riga, 9).setValue(operatore);
      foglio.getRange(riga, 6).setValue(res.url);
    }
    return res;
  } catch(e) { return { success: false, error: e.message }; }
}

function cambiaStatoRimborso(riga, nuovoStato, note, operatore, importoAutorizzato) {
  try {
    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    
    // Aggiorniamo i dati nel JSON della colonna 5 per includere l'importo autorizzato
    let datiJson = JSON.parse(foglio.getRange(riga, 5).getValue());
    if (importoAutorizzato !== undefined && importoAutorizzato !== null) {
      datiJson.importoAutorizzato = importoAutorizzato;
    }
    foglio.getRange(riga, 5).setValue(JSON.stringify(datiJson));
    
    foglio.getRange(riga, 7).setValue(nuovoStato);
    foglio.getRange(riga, 8).setValue(note || "");
    foglio.getRange(riga, 9).setValue(operatore || "");

    if (nuovoStato === "APPROVATO") {
      let res = compilaTemplateRimborsoDoc(datiJson, true, false);
      if(res.success) foglio.getRange(riga, 6).setValue(res.url);
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// ================= MODULO PERMESSI ================= //
function getUtentiPerDelega(utenteLoggato) {
  try {
    const dati = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE').getDataRange().getValues();
    let liv = GERARCHIA[utenteLoggato.ruoloBase] || 1;
    let isSegrGenER = (liv === 2 && utenteLoggato.provincia === 'EMILIA ROMAGNA');
    let lista = []; 
    for (let i = 1; i < dati.length; i++) {
      if (!String(dati[i][11]).toUpperCase().includes('CONSIGLIERE')) continue; 
      let u = { cognome: String(dati[i][3]).toUpperCase().trim(), nome: String(dati[i][4]).toUpperCase().trim(), grado: dati[i][2], carica: dati[i][5], provincia: String(dati[i][6]).toUpperCase().trim(), segreteria: String(dati[i][12]).toUpperCase().trim(), cip: String(dati[i][15]).toUpperCase().trim() };
      if (!u.cip) continue;
      
      // LOGICA CONTO TERZI:
      if (liv === 1 && u.cip === utenteLoggato.cip) { lista.push(u); } 
      else if (isSegrGenER || liv >= 3) { lista.push(u); } 
      else if (liv === 2 && u.provincia === utenteLoggato.provincia) { lista.push(u); }
    }
    return lista;
  } catch (e) { throw new Error(e.message); }
}

function salvaRichiestaPermesso(datiRichiesta) {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI');
    let rBase = String(datiRichiesta.ruoloLoggato).split(',')[0].trim();
    let liv = GERARCHIA[rBase] || 1;
    let isSegrGenER = (liv === 2 && datiRichiesta.provinciaLoggato === 'EMILIA ROMAGNA');
    
    let stato = "IN ATTESA"; 
    if (isSegrGenER || liv >= 3) { stato = "APPROVATO"; } 
    else if (liv === 2) { stato = "VISTATO"; }

    let inseritoDa = (String(datiRichiesta.cip).trim() !== String(datiRichiesta.cipInseritore).trim()) ? datiRichiesta.cognomeInseritore : "";
    let operatore = (stato !== "IN ATTESA") ? datiRichiesta.cognomeInseritore : "";

    datiRichiesta.periodi.forEach(p => {
      // --- FIX DATA: Scomponiamo la stringa YYYY-MM-DD e creiamo la data a mezzogiorno locale ---
      // Questo impedisce lo scivolamento al giorno precedente dovuto al fuso orario
      let partiInizio = p.inizio.split('-');
      let dateInizio = new Date(partiInizio[0], partiInizio[1] - 1, partiInizio[2], 12, 0, 0);
      
      let partiFine = p.fine.split('-');
      let dateFine = new Date(partiFine[0], partiFine[1] - 1, partiFine[2], 12, 0, 0);

      foglio.appendRow([ 
        new Date(), 
        datiRichiesta.cip, 
        `${datiRichiesta.grado} ${datiRichiesta.nome} ${datiRichiesta.cognome}`, 
        datiRichiesta.provincia, 
        dateInizio, 
        dateFine, 
        p.urgente ? "SI" : "NO", 
        p.supera ? "SI" : "NO", 
        stato, 
        p.motivazione || "", 
        inseritoDa, 
        operatore 
      ]);
    });
    return { success: true, stato: stato };
  } catch (e) { return { success: false, error: e.message }; }
}

function getPraticheDaGestire(provincia, ruolo) {
  try {
    const datiCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE').getDataRange().getValues();
    let cipOrder = {}; for (let i = 1; i < datiCariche.length; i++) { let c = String(datiCariche[i][15]).toUpperCase().trim(); if (c && !cipOrder[c]) cipOrder[c] = i; }

    const dati = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI').getDataRange().getValues();
    let daGestire = []; let storico = []; const dataOggi = new Date(); dataOggi.setHours(0, 0, 0, 0);
    let ruoloBase = String(ruolo).split(',')[0].trim();

    for (let i = 1; i < dati.length; i++) {
      if (!dati[i][0]) continue; 
      let stato = String(dati[i][8]).toUpperCase().trim();          
      let dReq = new Date(dati[i][0]); let dFine = new Date(dati[i][5]);
      let scaduta = (!isNaN(dFine.getTime()) && new Date(dFine.setHours(0,0,0,0)) < dataOggi);

      let p = { riga: i + 1, dataRichiesta: Utilities.formatDate(dReq, "GMT+1", "dd/MM/yyyy"), provincia: dati[i][3], cip: dati[i][1], idGruppo: Utilities.formatDate(dReq, "GMT+1", "ddMMyyyy_HHmm"), richiedente: dati[i][2], inizio: Utilities.formatDate(new Date(dati[i][4]), "GMT+1", "dd/MM/yyyy"), fine: Utilities.formatDate(new Date(dati[i][5]), "GMT+1", "dd/MM/yyyy"), urgente: dati[i][6], limite: dati[i][7], stato: stato, motivazione: dati[i][9] || "", inseritoDa: String(dati[i][10]||"").trim(), operatore: String(dati[i][11]||"").trim() };

      // VISIBILITÀ STORICO CORRETTA PER IL SEGRGEN
      if (ruoloBase === "SEGRGEN" && p.provincia === provincia) { 
          if (stato === "IN ATTESA") daGestire.push(p); 
          else storico.push(p); 
      } 
      else if (ruoloBase === "APPROVATORE") { 
          if (stato === "VISTATO") daGestire.push(p); 
          else storico.push(p); 
      }
      else if (ruoloBase === "GESTORE") { 
          if (stato === "APPROVATO") daGestire.push(p); 
          else storico.push(p); 
      }
      else if (ruoloBase === "AMMINISTRATORE") { 
          if (stato === "IN ATTESA" || stato === "VISTATO" || stato === "APPROVATO") daGestire.push(p); 
          else storico.push(p); 
      }
    }
    return { daGestire: daGestire, storico: storico, cipOrder: cipOrder };
  } catch(e) { throw new Error(e.message); }
}

// ================= AGGIORNAMENTO STATO PERMESSI ================= //
function aggiornaStatoPratica(riga, nuovoStato, note, operatore) {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI');
    
    // Aggiorna lo stato (Colonna I -> Indice 9)
    foglio.getRange(riga, 9).setValue(nuovoStato);
    
    // Aggiorna l'operatore che ha effettuato l'azione (Colonna L -> Indice 12)
    if (operatore) {
      foglio.getRange(riga, 12).setValue(operatore);
    }
    
    // Se c'è una nota di rifiuto, la accoda nella colonna delle motivazioni (Colonna J -> Indice 10)
    if (note) {
      let notaAttuale = foglio.getRange(riga, 10).getValue();
      let nuovaNota = notaAttuale ? notaAttuale + " | Note op: " + note : note;
      foglio.getRange(riga, 10).setValue(nuovaNota);
    }
    
    return { success: true };
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

function aggiornaStatoPraticaMultiplo(righeArray, nuovoStato, note, operatore) {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI');
    righeArray.forEach(riga => {
      foglio.getRange(riga, 9).setValue(nuovoStato);
      if (operatore) foglio.getRange(riga, 12).setValue(operatore);
      if (note) {
        let notaAttuale = foglio.getRange(riga, 10).getValue();
        let nuovaNota = notaAttuale ? notaAttuale + " | Note: " + note : note;
        foglio.getRange(riga, 10).setValue(nuovaNota);
      }
    });
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function generaPDFPratica(datiPratica) {
  try {
    // 1. ID del tuo modello Google Doc
    const TEMPLATE_ID = '1SekjRCrllxkY2N-zLOHgBef_8IAcu_MpLra0cRJ_Qw8';
    
    let folder; 
    try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } 
    catch(e) { folder = DriveApp.getRootFolder(); }

    // 2. RECUPERO DATI DAL FOGLIO "CARICHE"
    const sheetCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    const datiCariche = sheetCariche.getDataRange().getValues();
    
    let grado = ""; let nome = ""; let cognome = ""; 
    let reparto = ""; let dataNomina = ""; let carica = "";
    let organizzazione = "";

    for(let i = 1; i < datiCariche.length; i++) {
        if(String(datiCariche[i][15]).toUpperCase().trim() === String(datiPratica.cip).toUpperCase().trim()) {
            grado = String(datiCariche[i][2]).trim();     
            cognome = String(datiCariche[i][3]).trim();   
            nome = String(datiCariche[i][4]).trim();      
            carica = String(datiCariche[i][5]).trim();    
            reparto = String(datiCariche[i][8]).trim();   
            
            // DATA NOMINA (Colonna J - Indice 9)
            let valDataNomina = datiCariche[i][9];
            if (valDataNomina instanceof Date) {
                dataNomina = Utilities.formatDate(valDataNomina, "GMT+1", "dd/MM/yyyy");
            } else if (valDataNomina && String(valDataNomina).trim() !== "") {
                dataNomina = String(valDataNomina).trim();
            } else {
                dataNomina = "N/D"; 
            }
            
            // ESTRAZIONE ORGANIZZAZIONE (Colonna T - Indice 19)
            organizzazione = String(datiCariche[i][19]).trim();
            
            break;
        }
    }

    // Formattazione Grado, Nome e Cognome
    grado = grado.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    nome = nome.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    cognome = cognome.toUpperCase();
    let richiedenteFormattato = (grado || nome || cognome) ? `${grado} ${nome} ${cognome}` : datiPratica.richiedente;

    // 3. RECUPERO DATI DALLA SCHEDA "INDIRIZZI"
    let comandoCorpo = "";
    let ufficio = "";
    let emailComando = "";
    let pecComando = "";

    if (organizzazione !== "") {
        const sheetIndirizzi = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('INDIRIZZI');
        if (sheetIndirizzi) {
            const datiIndirizzi = sheetIndirizzi.getDataRange().getValues();
            
            for(let j = 1; j < datiIndirizzi.length; j++) {
                if (String(datiIndirizzi[j][0]).toUpperCase().trim() === organizzazione.toUpperCase()) {
                    comandoCorpo = String(datiIndirizzi[j][1]).trim(); 
                    ufficio      = String(datiIndirizzi[j][2]).trim(); 
                    emailComando = String(datiIndirizzi[j][3]).trim(); 
                    pecComando   = String(datiIndirizzi[j][4]).trim(); 
                    break; 
                }
            }
        }
    }

    // 4. CREAZIONE COPIA DEL MODELLO E SOSTITUZIONE TESTI
    
    // ---> MODIFICA QUI: Sostituisce eventuali "/" con "!" nel protocollo
    let safeProt = datiPratica.protocollo ? datiPratica.protocollo.replace(/\//g, '!') : "Senza_Prot";
    let safeCognome = cognome ? cognome : datiPratica.richiedente.split(' ').pop().toUpperCase();
    const nomeFile = "Lettera_permesso_" + safeCognome + "_" + safeProt;
    
    const tempDocFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(nomeFile, folder);
    const doc = DocumentApp.openById(tempDocFile.getId());
    const body = doc.getBody();
    
    // Raggruppa i periodi in blocchi da 5 e va a capo
    let periodiFormattati = datiPratica.periodi.map(p => (p.inizio === p.fine) ? p.inizio : "dal " + p.inizio + " al " + p.fine);
    let chunksPeriodi = [];
    for (let k = 0; k < periodiFormattati.length; k += 5) {
        chunksPeriodi.push(periodiFormattati.slice(k, k + 5).join(" - "));
    }
    // Unisce i blocchi con un trattino e un ritorno a capo (\n)
    let strPeriodi = chunksPeriodi.join(" -\n");
    
    let protocolloStr = datiPratica.protocollo ? datiPratica.protocollo : "___________";
    let dataOdierna = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");

    body.replaceText("{{PROTOCOLLO}}", protocolloStr);
    body.replaceText("{{RICHIEDENTE}}", richiedenteFormattato); 
    body.replaceText("{{GRADO}}", grado);
    body.replaceText("{{NOME}}", nome);
    body.replaceText("{{COGNOME}}", cognome);
    body.replaceText("{{CIP}}", datiPratica.cip || "");
    body.replaceText("{{PROVINCIA}}", datiPratica.provincia || "");
    body.replaceText("{{PERIODI}}", strPeriodi);
    body.replaceText("{{DATA_OGGI}}", dataOdierna);
    body.replaceText("{{REPARTO}}", reparto);
    body.replaceText("{{DATA_NOMINA}}", dataNomina);
    body.replaceText("{{CARICA}}", carica);
    
    body.replaceText("{{COMANDO_CORPO}}", comandoCorpo);
    body.replaceText("{{UFFICIO}}", ufficio);
    body.replaceText("{{EMAIL_COMANDO}}", emailComando);
    body.replaceText("{{PEC_COMANDO}}", pecComando);
    
    // 5. SALVATAGGIO E CONVERSIONE IN PDF
    doc.saveAndClose(); 
    const pdfBlob = tempDocFile.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob);
    
    // Cestina la copia in Word
    tempDocFile.setTrashed(true);
    
    return { success: true, url: pdfFile.getUrl() };
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

function confermaGiornoFruito(riga, targetIsoStr, operatore) {
  try {
      const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI');
      let rowValues = foglio.getRange(riga, 1, 1, Math.max(12, foglio.getLastColumn())).getValues()[0];
      let dataRow = []; for(let i=0; i<12; i++) dataRow.push(rowValues[i] !== undefined ? rowValues[i] : "");
      let dInizio = new Date(dataRow[4]); dInizio.setHours(0,0,0,0);
      let dFine = new Date(dataRow[5]); dFine.setHours(0,0,0,0);
      let parts = targetIsoStr.split('-'); let targetD = new Date(parts[0], parts[1]-1, parts[2]); targetD.setHours(0,0,0,0);
      
      if (dInizio.getTime() === dFine.getTime()) { foglio.getRange(riga, 9).setValue("FRUITO"); if (operatore) foglio.getRange(riga, 12).setValue(operatore); return { success: true }; }
      
      function addDays(date, days) { let r = new Date(date); r.setDate(r.getDate() + days); return Utilities.formatDate(r, "GMT+1", "yyyy-MM-dd"); }
      let targetIso = Utilities.formatDate(targetD, "GMT+1", "yyyy-MM-dd");
      let tInizio = Utilities.formatDate(dInizio, "GMT+1", "yyyy-MM-dd"); let tFine = Utilities.formatDate(dFine, "GMT+1", "yyyy-MM-dd");
      
      let rowFruito = [...dataRow]; rowFruito[4] = targetIso; rowFruito[5] = targetIso; rowFruito[8] = "FRUITO"; if (operatore) rowFruito[11] = operatore;
      
      if (targetIso === tInizio) { foglio.getRange(riga, 5).setValue(addDays(targetD, 1)); foglio.appendRow(rowFruito); } 
      else if (targetIso === tFine) { foglio.getRange(riga, 6).setValue(addDays(targetD, -1)); foglio.appendRow(rowFruito); } 
      else { foglio.getRange(riga, 6).setValue(addDays(targetD, -1)); foglio.appendRow(rowFruito); let rowRestante = [...dataRow]; rowRestante[4] = addDays(targetD, 1); foglio.appendRow(rowRestante); }
      return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function cancellaGiornoPratica(riga) {
  try { SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI').deleteRow(riga); return { success: true }; } catch(e) { return { success: false, error: e.message }; }
}

function getMieRichiesteEStatistiche(cip) {
  try {
    const dati = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI').getDataRange().getValues();
    let mieRichieste = []; let stats = { fruiti: 0, chiesti: 0, inSospeso: 0 };
    const oggi = new Date(); const meseCorrente = oggi.getMonth(); const annoCorrente = oggi.getFullYear();
    const dataOggiBase = new Date(); dataOggiBase.setHours(0,0,0,0);

    for (let i = 1; i < dati.length; i++) {
      if (!dati[i][0]) continue; 
      if (String(dati[i][1]).toUpperCase().trim() !== cip.trim()) continue;

      let dReqObj = new Date(dati[i][0]); let dInizioObj = new Date(dati[i][4]); let dFineObj = new Date(dati[i][5]);
      let stato = String(dati[i][8]).toUpperCase().trim();
      let isScaduta = (!isNaN(dFineObj.getTime()) && new Date(dFineObj.setHours(0,0,0,0)) < dataOggiBase);

      mieRichieste.push({
        riga: i + 1, idGruppo: Utilities.formatDate(dReqObj, "GMT+1", "ddMMyyyy_HHmm"), dataRichiesta: Utilities.formatDate(dReqObj, "GMT+1", "dd/MM/yyyy"), inizio: Utilities.formatDate(dInizioObj, "GMT+1", "dd/MM/yyyy"), fine: Utilities.formatDate(new Date(dati[i][5]), "GMT+1", "dd/MM/yyyy"), urgente: dati[i][6], limite: dati[i][7], stato: stato, motivazione: dati[i][9] || "", isScaduta: isScaduta, inseritoDa: String(dati[i][10]||"").trim(), operatore: String(dati[i][11]||"").trim()
      });
      if (dReqObj.getMonth() === meseCorrente && dReqObj.getFullYear() === annoCorrente) stats.chiesti++;
      if (dInizioObj.getMonth() === meseCorrente && dInizioObj.getFullYear() === annoCorrente && stato === "FRUITO") stats.fruiti++;
      if (stato === "IN ATTESA" || stato === "VISTATO") stats.inSospeso++;
    }
    return { richieste: mieRichieste, statistiche: stats };
  } catch(e) { throw new Error(e.message); }
}

function getDatiStatisticheGlobali(provinciaUtente, ruolo, cipUtente) {
  try {
    const datiCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE').getDataRange().getValues();
    const datiPermessi = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI').getDataRange().getValues();
    let ruoloBase = String(ruolo).split(',')[0].trim();
    let isSegrGenER = (ruoloBase === "SEGRGEN" && provinciaUtente === 'EMILIA ROMAGNA');

    let utenti = [];
    for (let i = 1; i < datiCariche.length; i++) {
      let prov = String(datiCariche[i][6]).toUpperCase().trim(); let cip = String(datiCariche[i][15]).toUpperCase().trim();
      if(!cip) continue;
      if (ruoloBase === "UTENTE" && cip !== cipUtente) continue; 
      if (ruoloBase === "SEGRGEN" && !isSegrGenER && prov !== provinciaUtente) continue; 
      utenti.push({ grado: datiCariche[i][2], cognome: String(datiCariche[i][3]).toUpperCase().trim(), nome: String(datiCariche[i][4]).toUpperCase().trim(), provincia: prov, cip: cip });
    }
    
    let permessi = [];
    for (let i = 1; i < datiPermessi.length; i++) {
      if (!datiPermessi[i][0]) continue;
      let prov = String(datiPermessi[i][3]).toUpperCase().trim(); let rowCip = String(datiPermessi[i][1]).toUpperCase().trim();
      
      if (ruoloBase === "UTENTE" && rowCip !== cipUtente) continue;
      if (ruoloBase === "SEGRGEN" && !isSegrGenER && prov !== provinciaUtente) continue;
      
      let valInizio = datiPermessi[i][4]; if (!valInizio) continue; let dInizio = new Date(valInizio); if (isNaN(dInizio.getTime())) continue;
      let dFineStr = Utilities.formatDate(dInizio, "GMT+1", "dd/MM/yyyy");
      if(datiPermessi[i][5] && !isNaN(new Date(datiPermessi[i][5]).getTime())) dFineStr = Utilities.formatDate(new Date(datiPermessi[i][5]), "GMT+1", "dd/MM/yyyy");

      permessi.push({ cip: rowCip, richiedente: String(datiPermessi[i][2]).toUpperCase().trim(), mese: dInizio.getMonth() + 1, anno: dInizio.getFullYear(), inizio: Utilities.formatDate(dInizio, "GMT+1", "dd/MM/yyyy"), fine: dFineStr, stato: String(datiPermessi[i][8]).toUpperCase().trim() });
    }
    return { utenti: utenti, permessi: permessi };
  } catch(e) { throw new Error(e.message); }
}

function generaReportFruitiPDF(anno, mese, ruolo, provincia) {
  try {
    const datiGlobali = getDatiStatisticheGlobali(provincia, ruolo, "");
    let strMese = mese !== "TUTTI" ? mese : "Annuale";
    const doc = DocumentApp.create(`Report_Permessi_Fruiti_${strMese}_${anno}`); const body = doc.getBody();
    body.insertParagraph(0, `REPORT PERMESSI EFFETTIVAMENTE FRUITI\nMese: ${strMese} - Anno: ${anno}\n`).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    
    let mapFruiti = {}; let trovatoQualcosa = false;
    datiGlobali.utenti.forEach(u => {
        let pUtente = datiGlobali.permessi.filter(p => p.stato === 'FRUITO' && p.cip.trim() === u.cip.trim() && (anno === "TUTTI" || String(p.anno) === anno) && (mese === "TUTTI" || String(p.mese) === mese));
        if(pUtente.length > 0) {
            trovatoQualcosa = true; let totaleGG = 0; let dettagli = [];
            pUtente.forEach(p => {
               let d = new Date(p.fine.split('/')[2], p.fine.split('/')[1] - 1, p.fine.split('/')[0]) - new Date(p.inizio.split('/')[2], p.inizio.split('/')[1] - 1, p.inizio.split('/')[0]);
               let gg = Math.round(d / (1000 * 60 * 60 * 24)) + 1; totaleGG += gg;
               dettagli.push(`${p.inizio === p.fine ? `Il ${p.inizio}` : `Dal ${p.inizio} al ${p.fine}`} (${gg} gg)`);
            });
            mapFruiti[`${u.cognome} ${u.nome}`] = { provincia: u.provincia, totale: totaleGG, dettagli: dettagli };
        }
    });

    if (!trovatoQualcosa) body.appendParagraph("Nessun permesso fruito registrato in questo periodo.");
    else Object.keys(mapFruiti).sort().forEach(nome => { body.appendParagraph(`\nSegretario: ${nome} (Provincia: ${mapFruiti[nome].provincia})`).setBold(true); body.appendParagraph(`Totale giorni fruiti: ${mapFruiti[nome].totale}`).setBold(false); mapFruiti[nome].dettagli.forEach(d => body.appendListItem(d)); });

    doc.saveAndClose();
    let folder; try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }
    const pdfFile = folder.createFile(doc.getAs('application/pdf')); DriveApp.getFileById(doc.getId()).setTrashed(true); 
    return { success: true, url: pdfFile.getUrl() };
  } catch(e) { return { success: false, error: e.message }; }
}

// ================= MODULO VISITE AI REPARTI ================= //
function getUtentiPerVisita(utenteLoggato) {
  try {
    const dati = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE').getDataRange().getValues();
    let liv = GERARCHIA[utenteLoggato.ruoloBase] || 1;
    let isSegrGenER = (liv === 2 && utenteLoggato.provincia === 'EMILIA ROMAGNA');
    let lista = []; 
    for (let i = 1; i < dati.length; i++) {
      let u = { cognome: String(dati[i][3]).toUpperCase().trim(), nome: String(dati[i][4]).toUpperCase().trim(), grado: dati[i][2], carica: dati[i][5], provincia: String(dati[i][6]).toUpperCase().trim(), segreteria: String(dati[i][12]).toUpperCase().trim(), cip: String(dati[i][15]).toUpperCase().trim() };
      if (!u.cip) continue;
      
      if (isSegrGenER || liv >= 3) { lista.push(u); } 
      else if (liv === 2 && u.provincia === utenteLoggato.provincia) { lista.push(u); }
      else if (liv === 1 && u.cip === utenteLoggato.cip) { lista.push(u); } 
    }
    return lista;
  } catch (e) { return []; }
}

function salvaRichiestaVisita(dati) {
  try {
    let foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_VISITE');
    if (!foglio) return { success: false, error: "Creare il foglio 'RICHIESTE_VISITE'." };
    const dataInvio = new Date();
    dati.visite.forEach(v => {
      let stringaPartecipanti = v.partecipanti.map(p => p.nome + " (" + p.cip + ")").join(" | ");
      foglio.appendRow([ dataInvio, dati.cipInseritore, dati.nomeInseritore, v.reparto, v.data, v.ora, stringaPartecipanti, "IN ATTESA", "" ]);
    });
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function getVisite(ruolo, cipLoggato, cognomeLoggato) {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_VISITE');
    if (!foglio) return { mie: [], daGestire: [], storico: [] };
    const dati = foglio.getDataRange().getValues();
    let mie = [], daGestire = [], storico = [];
    let liv = GERARCHIA[String(ruolo).split(',')[0].trim()] || 1;
    
    for (let i = 1; i < dati.length; i++) {
      if (!dati[i][0]) continue;
      let stringaPartecipanti = String(dati[i][6]);
      let v = {
        riga: i + 1, dataInvio: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy HH:mm"),
        cip: String(dati[i][1]).trim(), inseritore: dati[i][2],
        reparto: dati[i][3], dataVisita: dati[i][4], oraVisita: dati[i][5],
        partecipanti: stringaPartecipanti, stato: String(dati[i][7]).toUpperCase().trim(), operatore: dati[i][8] || ""
      };
      
      if (v.dataVisita instanceof Date) v.dataVisita = Utilities.formatDate(v.dataVisita, "GMT+1", "dd/MM/yyyy");
      else if(String(v.dataVisita).includes('-')) { let p = String(v.dataVisita).split('-'); v.dataVisita = `${p[2]}/${p[1]}/${p[0]}`; }

      let isInseritore = (v.cip === cipLoggato);
      let isPartecipante = (!isInseritore) && (stringaPartecipanti.includes("(" + cipLoggato + ")") || (cognomeLoggato && stringaPartecipanti.includes(cognomeLoggato)));

      if (isPartecipante || isInseritore) mie.push(v);
      
      let showInGestione = false;
      let showInStorico = false;

      if (isInseritore || liv >= 2) {
          if (liv <= 2 && isInseritore) {
              if (v.stato === "IN ATTESA") showInGestione = true; else showInStorico = true;
          } else if (liv >= 3) {
              if (v.stato === "IN ATTESA" || v.stato === "APPROVATO") showInGestione = true; else showInStorico = true;
          }
          
          if (showInGestione) daGestire.push(v);
          else if (showInStorico) storico.push(v);
      }
    }
    return { mie: mie.reverse(), daGestire: daGestire.reverse(), storico: storico.reverse() };
  } catch (e) { throw new Error(e.message); }
}

function aggiornaStatoVisita(riga, stato, operatore) {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_VISITE');
    foglio.getRange(riga, 8).setValue(stato);
    foglio.getRange(riga, 9).setValue(operatore);
    return { success: true };
  } catch (e) { return { success: false, error: e.message }; }
}

function aggiornaStatoVisiteMultiplo(righeArray, stato, operatore) {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_VISITE');
    righeArray.forEach(r => {
        foglio.getRange(r, 8).setValue(stato);
        foglio.getRange(r, 9).setValue(operatore);
    });
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function generaPDFVisite(righeArray, protocollo, repartiModificati = {}) {
  try {
    const TEMPLATE_ID_VISITE = '1q442uL0sjWkp0K9DeUXK1izFTeaVwQYD1x72XnKdMP8';
    let folder; 
    try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }

    const ss = SpreadsheetApp.openById(GESTIONALE_ID);
    const sheetCariche = ss.getSheetByName('CARICHE');
    const datiCariche = sheetCariche.getDataRange().getValues();
    const sheetVisite = ss.getSheetByName('RICHIESTE_VISITE');
    const datiVisite = sheetVisite.getDataRange().getValues();

    // 1. RECUPERO INFO INDIRIZZO
    let inseritoreCip = "";
    for(let i=1; i<datiVisite.length; i++) {
      if(righeArray.includes(i+1)) { inseritoreCip = String(datiVisite[i][1]).trim(); break; }
    }

    let organizzazione = "";
    for(let i=1; i<datiCariche.length; i++) {
      if(String(datiCariche[i][15]).toUpperCase().trim() === inseritoreCip.toUpperCase()) {
        organizzazione = String(datiCariche[i][19]).trim(); 
        break;
      }
    }

    let comandoCorpo = "", citta = "", emailComando = "";
    if (organizzazione) {
      const sheetIndirizzi = ss.getSheetByName('INDIRIZZI');
      if (sheetIndirizzi) {
        const dInd = sheetIndirizzi.getDataRange().getValues();
        for(let j=1; j<dInd.length; j++) {
          if (String(dInd[j][0]).toUpperCase().trim() === organizzazione.toUpperCase()) {
            comandoCorpo = String(dInd[j][1]).trim(); 
            citta        = String(dInd[j][5]).trim(); // Colonna F
            emailComando = String(dInd[j][3]).trim(); 
            break;
          }
        }
      }
    }

    // 2. CREAZIONE DOCUMENTO
    let safeProt = protocollo ? protocollo.replace(/\//g, '!') : "Senza_Prot";
    const nomeFile = "Lettera_visite_" + safeProt;
    const tempDocFile = DriveApp.getFileById(TEMPLATE_ID_VISITE).makeCopy(nomeFile, folder);
    const doc = DocumentApp.openById(tempDocFile.getId());
    const body = doc.getBody();

    body.replaceText("{{PROTOCOLLO}}", protocollo || "___________");
    body.replaceText("{{DATA_OGGI}}", Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"));
    body.replaceText("{{COMANDO_CORPO}}", comandoCorpo);
    body.replaceText("{{CITTA’}}", citta);
    body.replaceText("{{EMAIL_COMANDO}}", emailComando);

    // 3. GESTIONE {{ELENCO_VISITE}} DA TABELLA MODELLO
    let rangeElement = body.findText("{{ELENCO_VISITE}}");
    if (!rangeElement) throw new Error("Segnaposto {{ELENCO_VISITE}} non trovato");
    
    let el = rangeElement.getElement();
    let par = el.getParent();
    if (par.getType() === DocumentApp.ElementType.TEXT) {
      par = par.getParent(); 
    }
    
    let cellModello = par.getParent();
    if (cellModello.getType() !== DocumentApp.ElementType.TABLE_CELL) {
      throw new Error("Il segnaposto {{ELENCO_VISITE}} DEVE essere dentro una cella di tabella nel Modello Google Docs.");
    }
    
    let rowModello = cellModello.getParent().asTableRow();
    let table = rowModello.getParent().asTable();
    
    let baseFontSize = el.asText().getFontSize() || 11;
    let newFontSize = baseFontSize - 2; 

    // ESTREZIONE E ORDINAMENTO CRONOLOGICO
    let visiteSelezionate = [];
    
    righeArray.forEach(rigaIndex => {
      let rigaDati = datiVisite[rigaIndex-1];
      let dataV = rigaDati[4];
      let oraV = rigaDati[5];
      
      let ts = 0;
      if (dataV instanceof Date) {
        let dateClone = new Date(dataV.getTime());
        if(oraV) {
           let parts = String(oraV).split(':');
           dateClone.setHours(parseInt(parts[0], 10), parseInt(parts[1]||0, 10), 0);
        }
        ts = dateClone.getTime();
      } else {
        let parts = String(dataV).split('/');
        if(parts.length === 3) {
            let d = new Date(parts[2], parts[1]-1, parts[0]);
            if(oraV) {
               let hparts = String(oraV).split(':');
               d.setHours(parseInt(hparts[0], 10), parseInt(hparts[1]||0, 10), 0);
            }
            ts = d.getTime();
        }
      }

      let repartoFinale = rigaDati[3];
      if (repartiModificati && repartiModificati[rigaIndex]) {
        repartoFinale = repartiModificati[rigaIndex];
      }

      visiteSelezionate.push({
        reparto: repartoFinale,
        dataOrig: dataV,
        ora: oraV,
        timestamp: ts,
        partecipantiGrezzi: String(rigaDati[6])
      });
    });

    visiteSelezionate.sort((a, b) => a.timestamp - b.timestamp);

    let indent = "  "; 
    let contattiNotaPieDiPagina = new Map(); 

    // COSTRUZIONE ELENCO CLONANDO LA RIGA DEL MODELLO
    visiteSelezionate.forEach((visita) => {
      // Clona la riga configurata da te nel Modello (eredita "Non spezzare su due pagine")
      let newRow = rowModello.copy();
      table.appendTableRow(newRow);
      
      let newCell = newRow.getCell(0);
      newCell.clear(); // Svuota il segnaposto clonato
      newCell.setPaddingBottom(3).setPaddingTop(3).setPaddingLeft(5).setPaddingRight(5);

      let righeTesto = [];
      let dataStr = visita.dataOrig;
      if (dataStr instanceof Date) dataStr = Utilities.formatDate(dataStr, "GMT+1", "dd/MM/yyyy");
      
      righeTesto.push({label: "REPARTO/COMANDO: ", value: visita.reparto});
      righeTesto.push({label: "GIORNO: ", value: dataStr + " ore " + visita.ora});
      righeTesto.push({label: "RAPPRESENTANTI:", value: ""});

      let pGrezzi = visita.partecipantiGrezzi.split(' | ');
      pGrezzi.forEach(pStr => {
        let cipMatch = pStr.match(/\(([^)]+)\)/);
        if (cipMatch) {
          let cipP = cipMatch[1].toUpperCase().trim();
          for(let k=1; k<datiCariche.length; k++) {
            if(String(datiCariche[k][15]).toUpperCase().trim() === cipP) {
              let cognome = String(datiCariche[k][3]).toUpperCase().trim();
              let nome = String(datiCariche[k][4]).toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).trim();
              let carica = String(datiCariche[k][5]).toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).trim();
              let prov = String(datiCariche[k][6]).toLowerCase().replace(/\b\w/g, c => c.toUpperCase()).trim();
              let email = String(datiCariche[k][13]).trim();   
              let telefono = String(datiCariche[k][14]).trim(); 
              
              righeTesto.push({text: indent + "- " + cognome + " " + nome});
              righeTesto.push({text: indent + "  " + carica + " SIM CC " + prov, italic: true, isCarica: true});

              if (!contattiNotaPieDiPagina.has(cipP)) {
                  let strContatto = `${cognome} ${nome}`;
                  if (telefono || email) {
                      strContatto += ` -`;
                      if (telefono) strContatto += ` Tel: ${telefono}`;
                      if (telefono && email) strContatto += ` |`;
                      if (email) strContatto += ` Email: ${email}`;
                  }
                  contattiNotaPieDiPagina.set(cipP, strContatto);
              }
              break;
            }
          }
        }
      });

      // Compila la cella clonata
      righeTesto.forEach((riga, i) => {
        let currentP;
        if (i === 0 && newCell.getNumChildren() > 0) {
          currentP = newCell.getChild(0).asParagraph();
          currentP.setText(riga.label ? (riga.label + riga.value) : riga.text);
        } else {
          currentP = newCell.appendParagraph(riga.label ? (riga.label + riga.value) : riga.text);
        }
        
        currentP.setLineSpacing(1.0).setSpacingBefore(0).setSpacingAfter(0);
        let finalSize = riga.isCarica ? (newFontSize - 1) : newFontSize;
        currentP.setFontSize(finalSize);
        
        currentP.setBold(false).setItalic(false);
        if (riga.label) currentP.editAsText().setBold(0, riga.label.length - 1, true);
        if (riga.italic) currentP.setItalic(true);
      });
    });

    // Rimuove la riga originale "template" con il segnaposto
    rowModello.removeFromParent();

    // 4. INSERIMENTO NOTA CON I CONTATTI ALLA FINE ASSOLUTA DEL DOCUMENTO
    if (contattiNotaPieDiPagina.size > 0) {
        let testoNota = "________________________\nContatti:\n" + Array.from(contattiNotaPieDiPagina.values()).join("\n");
        // Qualche accapo vuoto per spingerlo in basso sotto la firma
        body.appendParagraph("\n\n"); 
        let pContatti = body.appendParagraph(testoNota);
        pContatti.setFontSize(9); 
        pContatti.setLineSpacing(1.0);
    }

    doc.saveAndClose(); 
    const pdfBlob = tempDocFile.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob);
    tempDocFile.setTrashed(true);
    
    return { success: true, url: pdfFile.getUrl() };
  } catch(e) { return { success: false, error: e.message }; }
}
// ================= NOTIFICHE UNIFICATE ================= //
function getConteggioNotifiche(utenteLoggato) {
  try {
    let countsP = { mieInAttesa: 0, mieVistate: 0, mieApprovate: 0, mieRifiutate: 0, mieEvase: 0, mieFruite: 0, daVistare: 0, daApprovare: 0, daEvadere: 0, hashMie: "", hashGestione: "" };
    let countsV = { mieInAttesa: 0, mieApprovate: 0, mieRifiutate: 0, mieEvase: 0, daApprovare: 0, daEvadere: 0, hashMie: "", hashGestione: "" };
    let countsR = { mieDaCorreggere: 0, mieInAttesa: 0, miePagate: 0, mieRifiutate: 0, daApprovare: 0, daPagare: 0, hashMie: "", hashGestione: "" };

    let liv = GERARCHIA[utenteLoggato.ruoloBase] || 1;
    let ruoloStr = String(utenteLoggato.ruolo || "").toUpperCase();
    
    let isTesoriere1 = ruoloStr.includes('TESORIERE') && !ruoloStr.includes('TESORIERE2');
    let isTesoriere2 = ruoloStr.includes('TESORIERE2');
    let isTesoriere = ruoloStr.includes('TESORIERE'); // Variabile di sicurezza
    let provLoggato = utenteLoggato.provincia;

    const dataOggi = new Date(); dataOggi.setHours(0, 0, 0, 0);
    const limiteRecenti = new Date(); limiteRecenti.setDate(limiteRecenti.getDate() - 15);

    if(liv === 5) {
       let fMod = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('MODIFICHE_PROFILO');
       if(fMod) {
          let dMod = fMod.getDataRange().getValues();
          for(let i=1; i<dMod.length; i++) { if(dMod[i][4] === 'IN ATTESA') countsP.modificheProfilo++; }
       }
    }

    const foglioP = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI');
    if (foglioP) {
      const datiP = foglioP.getDataRange().getValues();
      for (let i = 1; i < datiP.length; i++) {
        if (!datiP[i][0]) continue;
        let cip = String(datiP[i][1]).toUpperCase().trim(); let prov = String(datiP[i][3]).toUpperCase().trim();
        let stato = String(datiP[i][8]).toUpperCase().trim(); let isRecente = new Date(datiP[i][0]) >= limiteRecenti;
        let isScaduta = false; if(datiP[i][5] && !isNaN(new Date(datiP[i][5]).getTime())) { let df=new Date(datiP[i][5]); df.setHours(0,0,0,0); if (df<dataOggi) isScaduta=true; }

        if (cip === utenteLoggato.cip && isRecente) {
           countsP.hashMie += i + stato;
           if (stato === 'IN ATTESA') countsP.mieInAttesa++; if (stato === 'VISTATO') countsP.mieVistate++;
           if (stato === 'APPROVATO') countsP.mieApprovate++; if (stato === 'RIFIUTATO' || stato === 'RIFIUTATA') countsP.mieRifiutate++;
           if (stato === 'EVASA' || stato === 'EVASO') countsP.mieEvase++; if (stato === 'FRUITO') countsP.mieFruite++;
        }
        
        let isInScopeGestione = false;
        if (liv === 5 || liv >= 3) isInScopeGestione = true;
        else if (liv === 2 && prov === provLoggato) isInScopeGestione = true;
        
        if (isInScopeGestione) {
            if (isRecente) countsP.hashGestione += i + stato;
            if (!isScaduta) {
                if (((liv === 2 && prov === provLoggato) || liv === 5) && stato === 'IN ATTESA') countsP.daVistare++;
                if ((liv === 3 || liv === 5) && stato === 'VISTATO') countsP.daApprovare++;
                if (liv >= 4 && stato === 'APPROVATO') countsP.daEvadere++;
            }
        }
      }
    }

    const foglioV = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_VISITE');
    if (foglioV) {
      const datiV = foglioV.getDataRange().getValues();
      for(let i = 1; i < datiV.length; i++) {
        if(!datiV[i][0]) continue;
        let cip = String(datiV[i][1]).trim(); let partecipantiStr = String(datiV[i][6]);
        let stato = String(datiV[i][7]).trim().toUpperCase(); let isRecente = new Date(datiV[i][0]) >= limiteRecenti;
        
        let isInseritore = (cip === utenteLoggato.cip);
        let isPartecipante = (!isInseritore) && (partecipantiStr.includes("(" + utenteLoggato.cip + ")") || partecipantiStr.includes(utenteLoggato.cognome));

        if ((isPartecipante || isInseritore) && isRecente) {
          countsV.hashMie += i + stato;
          if (stato === 'IN ATTESA') countsV.mieInAttesa++;
          if (stato === 'APPROVATO') countsV.mieApprovate++;
          if (stato === 'RIFIUTATO') countsV.mieRifiutate++;
          if (stato === 'EVASA' || stato === 'EVASO') countsV.mieEvase = (countsV.mieEvase || 0) + 1;
        }
        
        if (liv >= 3 && stato === 'IN ATTESA') { countsV.daApprovare++; if(isRecente) countsV.hashGestione += i + stato; }
        if (liv >= 4 && stato === 'APPROVATO') { countsV.daEvadere++; if(isRecente) countsV.hashGestione += i + stato; }
      }
    }

    const fRimb = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_RIMBORSI');
    if (fRimb) {
       const datiR = fRimb.getDataRange().getValues();
       for(let i=1; i<datiR.length; i++) {
           if(!datiR[i][0]) continue;
           let cip = String(datiR[i][1]).trim(); let prov = String(datiR[i][3]).trim();
           let stato = String(datiR[i][6]).trim().toUpperCase(); let isRecente = new Date(datiR[i][0]) >= limiteRecenti;

           let json = {}; try{ json = JSON.parse(datiR[i][4]); } catch(e){}
           let sp = json.segreteriaPagante || "";
           let isProvinciale = sp.startsWith("Provinciale"); let isRegionale = sp.startsWith("Regionale") || sp === "Nazionale";

           if(cip === utenteLoggato.cip && isRecente) {
               countsR.hashMie += i + stato;
               if(stato === 'IN ATTESA') countsR.mieInAttesa++;
               if(stato === 'DA CORREGGERE') countsR.mieDaCorreggere++;
               if(stato === 'PAGATO') countsR.miePagate++;
               if(stato === 'RIFIUTATO') countsR.mieRifiutate++;
           }

           let canApprove = false; let canPay = false; let seesIt = false;
           
           // Estrae la provincia esatta da chi deve pagare
           let targetProv = prov; 
           if (isProvinciale) {
               let match = sp.match(/\(([^)]+)\)/);
               if (match) targetProv = match[1].toUpperCase().trim();
           }
           
           // LOGICA DI AUTORIZZAZIONE E NOTIFICHE SEPARATA
           if (liv === 5) { 
               seesIt = true; canApprove = true; canPay = true; 
           } else {
               if (isProvinciale && targetProv === provLoggato) {
                   if (liv >= 2) { seesIt = true; canApprove = true; } 
                   if (isTesoriere2) { seesIt = true; canPay = true; }  
               } else if (isRegionale) {
                   if (liv >= 3 || (liv === 2 && provLoggato === 'EMILIA ROMAGNA')) { seesIt = true; canApprove = true; } 
                   if (isTesoriere1) { seesIt = true; canPay = true; }  
               }
           }

           if(seesIt) {
               if(stato === 'IN ATTESA' && canApprove) countsR.daApprovare++;
               if(stato === 'APPROVATO' && canPay) countsR.daPagare++;
               if (isRecente) countsR.hashGestione += i + stato;
           }
       }
    }
    // --- NUOVO BLOCCO NOTIFICHE RUBRICA ---
    let countsRubrica = { daApprovare: 0, hashGestione: "" };
    if (liv === 5) { // L'Amministratore riceve le notifiche
       try {
          let fRich = SpreadsheetApp.openById(RUBRICA_DB_ID).getSheetByName('RICHIESTE_RUBRICA');
          if (fRich) {
             let dRich = fRich.getDataRange().getValues();
             for(let i=1; i<dRich.length; i++) {
                if (dRich[i][5] === 'IN ATTESA') {
                   countsRubrica.daApprovare++;
                   countsRubrica.hashGestione += i + "ATTESA";
                }
             }
          }
       } catch(e) {}
    }
    
    return { permessi: countsP, visite: countsV, rimborsi: countsR, rubrica: countsRubrica };
  } catch (e) { return null; }
}
// ================= MODULO GESTIONE MATERIALI ================= //
const FOLDER_MATERIALI_SHEETS_ID = '1q-I43asRS2BUp58dJIb5gqwil7l-sy5Z'; // Cartella per i Fogli di calcolo
const FOLDER_MATERIALI_DOCS_ID = '1yycGzRe2WeY3_iywhC5y1OTlAzJgSHU8';   // Cartella root per i documenti/scontrini

// Funzione per recuperare o creare il file Fogli della Segreteria
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

    // 1. Carica utenti per le tendine dei Custodi
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

    // 2. Carica i materiali scansionando i file nella cartella Fogli Materiali
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
                fileId: file.getId(), // Fondamentale: traccia in quale file si trova la riga
                riga: i+1, dataReg: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy"), 
                categoria: dati[i][2], descrizione: dati[i][3], quantita: dati[i][4], 
                costo: parseFloat(dati[i][5] || 0), dataAcquistoIso: dAcquisto, 
                acquistatoDa: acquistatoDa, inCaricoA: inCaricoA, custode: dati[i][9], fileUrl: dati[i][10] 
            });
          }
        }
    }
    
    // Ordina tutti i materiali dal più recente al più vecchio
    res.materiali.sort((a, b) => new Date(b.dataAcquistoIso).getTime() - new Date(a.dataAcquistoIso).getTime());
    return res;
  } catch(e) { return { materiali: [], utentiMat: [] }; }
}

function salvaNuovoMateriale(dati, fileBase64, mimeType, fileName) {
  try {
    let rootFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
    let targetFolderName = dati.inCaricoA; // Nome della segreteria (es. "Segreteria Provinciale BOLOGNA")
    let targetFolder;
    
    // 1. GESTIONE CARTELLE SU DRIVE
    let folders = rootFolder.getFoldersByName(targetFolderName);
    if (folders.hasNext()) { 
      targetFolder = folders.next(); 
    } else { 
      targetFolder = rootFolder.createFolder(targetFolderName); 
    }

    // 2. CARICAMENTO EVENTUALE ALLEGATO (Scontrino/Fattura)
    let fileUrl = "";
    if (fileBase64 && fileName) {
      let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
      let uploadedFile = targetFolder.createFile(blob);
      fileUrl = uploadedFile.getUrl();
    }

    // 3. FIX DATA (Evita l'errore del giorno precedente)
    // Trasformiamo la stringa "YYYY-MM-DD" in un oggetto Date impostato a mezzogiorno locale
    let p = dati.data.split('-'); 
    let dataCorretta = new Date(p[0], p[1] - 1, p[2], 12, 0, 0);

    // 4. IDENTIFICAZIONE DEL FOGLIO DI CALCOLO DELLA SEGRETERIA
    // Utilizza la funzione di supporto già presente nel tuo codice
    let ssId = getOrCreateMaterialiFile(dati.inCaricoA);
    let foglio = SpreadsheetApp.openById(ssId).getSheetByName('MATERIALI');
    
    // 5. SALVATAGGIO RIGA NEL DATABASE
    // Ordine colonne: Timestamp, CIP, Categoria, Descrizione, Quantità, Costo, Data Acquisto, Acquistato Da, In Carico A, Custode, URL Documento
    foglio.appendRow([
      new Date(),           // Timestamp attuale
      dati.cip,             // CIP dell'inseritore
      dati.categoria,       // Categoria bene
      dati.descrizione,     // Descrizione bene
      dati.quantita,        // Quantità
      dati.costo,           // Costo totale
      dataCorretta,         // <--- DATA CORRETTA (FIXATA)
      dati.acquistatoDa,    // Segreteria che ha pagato
      dati.inCaricoA,       // Segreteria che lo ha in carico
      dati.custode,         // Nome del custode
      fileUrl               // Link al file su Drive
    ]);
    
    return { success: true };
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

function eliminaMateriale(fileId, riga) {
    try { 
        let sheet = SpreadsheetApp.openById(fileId).getSheetByName('MATERIALI');
        
        // 1. Legge il nome della Segreteria (colonna 9, "I" -> In Carico A)
        let segreteria = sheet.getRange(riga, 9).getValue();
        
        // 2. Legge l'URL del file prima di eliminare la riga (colonna 11, "K")
        let fileUrl = sheet.getRange(riga, 11).getValue();
        
        // 3. Se c'è un link, cerca il file e lo sposta nel cestino specifico
        if (fileUrl && String(fileUrl).trim() !== "") {
            try {
                // Estrae l'ID univoco dal link standard di Google Drive
                let extractIdMatch = String(fileUrl).match(/[-\w]{25,}/);
                if (extractIdMatch && extractIdMatch[0]) {
                    let file = DriveApp.getFileById(extractIdMatch[0]);
                    
                    // Trova la cartella radice dei documenti materiali
                    let rootDocsFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
                    
                    // Trova la cartella della specifica segreteria
                    let segrFolders = rootDocsFolder.getFoldersByName(segreteria);
                    let segrFolder;
                    if (segrFolders.hasNext()) {
                        segrFolder = segrFolders.next();
                    } else {
                        // Se per qualche motivo non c'è, la crea per sicurezza
                        segrFolder = rootDocsFolder.createFolder(segreteria);
                    }
                    
                    // Nome dinamico della cartella cestino
                    let cestinoName = "CESTINO_DOC_" + segreteria;
                    
                    // Trova o crea la cartella cestino all'interno della cartella della segreteria
                    let cestinoFolders = segrFolder.getFoldersByName(cestinoName);
                    let cestinoFolder;
                    if (cestinoFolders.hasNext()) {
                        cestinoFolder = cestinoFolders.next();
                    } else {
                        cestinoFolder = segrFolder.createFolder(cestinoName);
                    }
                    
                    // Sposta fisicamente il file nella cartella Cestino
                    file.moveTo(cestinoFolder);
                }
            } catch(errFile) {
                // Se il file era già stato spostato/cancellato a mano o non c'è, ignora l'errore e procedi
            }
        }
        
        // 4. Infine elimina la riga dal database Fogli
        sheet.deleteRow(riga); 
        
        return { success: true }; 
    } catch(e) { 
        return { success: false, error: e.message }; 
    }
}
// Rimuove il link dal foglio e sposta il file nel cestino
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
    
    // Svuota la cella del documento
    sheet.getRange(riga, 11).clearContent();
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getFirmaMutriBase64() {
  try {
    let folder = DriveApp.getFolderById('1ewdYZ5F_o-dW5SvOOtNVF978Y9gi5yba');
    let files = folder.getFilesByName('Mutri.png');
    if (files.hasNext()) {
      let f = files.next();
      return `data:${f.getMimeType()};base64,${Utilities.base64Encode(f.getBlob().getBytes())}`;
    }
  } catch(e) {}
  return "";
}

// Sposta l'eventuale file vecchio nel cestino e carica il nuovo file aggiornando la cella
function sostituisciDocumentoMateriale(fileId, riga, segreteria, fileBase64, mimeType, fileName) {
  try {
    let sheet = SpreadsheetApp.openById(fileId).getSheetByName('MATERIALI');
    let fileUrl = sheet.getRange(riga, 11).getValue();
    
    // 1. Sposta il vecchio file nel cestino (se esiste)
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
    
    // 2. Crea il nuovo file nella cartella della segreteria
    let rootDocsFolder = DriveApp.getFolderById(FOLDER_MATERIALI_DOCS_ID);
    let segrFolders = rootDocsFolder.getFoldersByName(segreteria);
    let targetFolder = segrFolders.hasNext() ? segrFolders.next() : rootDocsFolder.createFolder(segreteria);
    
    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, fileName);
    let uploadedFile = targetFolder.createFile(blob);
    let newFileUrl = uploadedFile.getUrl();
    
    // 3. Aggiorna il link nel foglio
    sheet.getRange(riga, 11).setValue(newFileUrl);
    
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

// Funzione Helper per recuperare la firma di Mutri da Drive
function getFirmaMutriBlob() {
  try {
    // Sostituisci l'ID con quello della tua cartella loghi se necessario
    // (Attualmente usa la costante CARTELLA_LOGHI_ID se definita, altrimenti metti l'ID diretto)
    let folder = DriveApp.getFolderById('1ewdYZ5F_o-dW5SvOOtNVF978Y9gi5yba');
    let files = folder.getFilesByName('Mutri.png');
    if (files.hasNext()) {
      return files.next().getBlob();
    }
  } catch(e) {
    console.log("Errore recupero firma: " + e.message);
  }
  return null;
}
// ================= GESTIONE ALLEGATO RIMBORSO DIFFERITO ================= //
function salvaRimborsoSenzaFile(dati) {
  try {
    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    if (!foglio) {
      foglio = ss.insertSheet('RICHIESTE_RIMBORSI');
      foglio.appendRow(["Timestamp", "CIP", "Richiedente", "Provincia", "Dati JSON", "URL File", "Stato", "Note", "Operatore"]);
    }
    // Salva con un nuovo stato specifico e senza file allegato
    foglio.appendRow([ new Date(), dati.cip, dati.richiedente, dati.provincia, JSON.stringify(dati), "", "IN ATTESA ALLEGATO", "", "" ]);
    return { success: true };
  } catch(e) { 
    return { success: false, error: e.message }; 
  }
}

function allegaFileRimborso(riga, cip, fileBase64, mimeType, fileName) {
  try {
    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    
    // Controllo sicurezza: verifichiamo che il CIP di chi sta caricando sia quello di chi ha creato la riga
    if(String(foglio.getRange(riga, 2).getValue()).toUpperCase().trim() !== String(cip).toUpperCase().trim()) {
        return {success: false, error: "Accesso negato o riga non corrispondente"};
    }

    let folder; 
    try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } 
    catch(e) { folder = DriveApp.getRootFolder(); }

    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, "Firmato_" + fileName);
    let uploadedFile = folder.createFile(blob);
    let fileUrl = uploadedFile.getUrl();

    // Aggiorna il link del file e lancia la pratica in "IN ATTESA" (Pronta per il Segretario)
    foglio.getRange(riga, 6).setValue(fileUrl);
    foglio.getRange(riga, 7).setValue("IN ATTESA");

    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function eliminaRimborso(riga, cip) {
    try {
        let foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_RIMBORSI');
        if(String(foglio.getRange(riga, 2).getValue()).toUpperCase().trim() === String(cip).toUpperCase().trim()) {
            foglio.deleteRow(riga);
            return {success: true};
        }
        return {success: false, error: "Accesso negato"};
    } catch(e) { return {success: false, error: e.message}; }
}
// ================= MODULO RUBRICA ================= //
function getDatiRubrica() {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    if (!foglio) return { error: "Foglio CARICHE non trovato." };
    
    const dati = foglio.getDataRange().getValues();
    let contatti = [];
    
    for (let i = 1; i < dati.length; i++) {
      if (!dati[i][15]) continue; // Salta la riga se non c'è un CIP valido
      
      let extraInfo = {};
      try { if(dati[i][20]) extraInfo = JSON.parse(dati[i][20]); } catch(e) {}
      
      contatti.push({
        grado: String(dati[i][2]).trim(),
        cognome: String(dati[i][3]).toUpperCase().trim(),
        nome: String(dati[i][4]).toUpperCase().trim(),
        carica: String(dati[i][5]).trim(),
        provincia: String(dati[i][6]).toUpperCase().trim(),
        reparto: String(dati[i][8]).trim(),
        email: String(dati[i][13]).toLowerCase().trim(),
        telefono: String(dati[i][14]).trim(),
        fotoUrl: extraInfo.foto || ""
      });
    }
    
    // Ordine alfabetico (Cognome Nome)
    contatti.sort((a, b) => (a.cognome + " " + a.nome).localeCompare(b.cognome + " " + b.nome));
    return contatti;
  } catch(e) {
    return { error: e.message };
  }
}
// ================= MODULO RUBRICA AVANZATA ================= //

function getContattiRubrica() {
  try {
    let contatti = [];
    // 1. LETTURA DINAMICA INTERNI (Da APP_simccer -> CARICHE)
    let foglioCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    if (foglioCariche) {
      let datiCar = foglioCariche.getDataRange().getValues();
      for (let i = 1; i < datiCar.length; i++) {
        if (!datiCar[i][15]) continue; 
        let provS = String(datiCar[i][6]).toUpperCase().trim();
        let livelloVal = "Provinciale";
        if (provS === "EMILIA ROMAGNA") livelloVal = "Regionale";
        else if (provS === "NAZIONALE") livelloVal = "Nazionale";
        let provUfficio = (provS !== "EMILIA ROMAGNA" && provS !== "NAZIONALE") ? provS : "";

        contatti.push({
          riga: "CAR_" + i,
          categoria: "INTERNI",
          tipo: "Persona",
          nome: String(datiCar[i][3]).toUpperCase().trim() + " " + String(datiCar[i][4]).toUpperCase().trim(),
          ruolo: String(datiCar[i][5]).trim() + " - " + String(datiCar[i][2]).trim(),
          livello: livelloVal,
          provincia: provUfficio,
          compagniaReparto: "", 
          sedeUfficio: String(datiCar[i][8]).trim(),
          telefono: String(datiCar[i][14]).trim(),
          email: String(datiCar[i][13]).toLowerCase().trim()
        });
      }
    }

    // 2. LETTURA ESTERNI DAL DATABASE RUBRICA (Solo quelli già approvati/presenti nelle schede)
    let ssNew = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let fogliEsterni = ['COMANDI', 'LEGALI', 'STAMPA', 'ALTRO'];
    fogliEsterni.forEach(nomeFoglio => {
      let foglio = ssNew.getSheetByName(nomeFoglio);
      if (foglio) {
        let dati = foglio.getDataRange().getValues();
        for (let i = 1; i < dati.length; i++) {
          if (!dati[i][0]) continue;
          let isAltro = (nomeFoglio === 'ALTRO');
          let offset = isAltro ? 1 : 0; 
          contatti.push({
            riga: nomeFoglio + "_" + i,
            categoria: isAltro ? String(dati[i][1]).trim() : nomeFoglio,
            nome: String(dati[i][1 + offset]).toUpperCase().trim(),
            ruolo: String(dati[i][2 + offset]).trim(),
            livello: String(dati[i][3 + offset]).trim(),
            provincia: String(dati[i][4 + offset]).trim(),
            compagniaReparto: String(dati[i][5 + offset]).trim(),
            sedeUfficio: String(dati[i][6 + offset]).trim(),
            telefono: String(dati[i][7 + offset]).trim(),
            email: String(dati[i][8 + offset]).trim()
          });
        }
      }
    });
    return contatti;
  } catch(e) { return { error: e.message }; }
}

// Funzione Unificata per Proposte Utenti e Salvataggi Admin
function salvaRichiestaRubrica(dati, isAdmin) {
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    
    // Se è Admin, salva direttamente nel foglio di destinazione
    if (isAdmin) {
       let targetSheetName = dati.categoria_base; // Es. 'COMANDI'
       let foglio = ss.getSheetByName(targetSheetName);
       if (!foglio) return { success: false, error: "Foglio non trovato nel Database." };
       
       if (targetSheetName === 'ALTRO') {
         foglio.appendRow([new Date(), dati.categoria_specifica, dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
       } else {
         foglio.appendRow([new Date(), dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
       }
       return { success: true };
    } 
    
    // Se è Utente, salva nella scheda di approvazione
    let foglioRichieste = ss.getSheetByName('RICHIESTE_RUBRICA');
    if (!foglioRichieste) {
      foglioRichieste = ss.insertSheet('RICHIESTE_RUBRICA');
      foglioRichieste.appendRow(["Timestamp", "Tipo", "DatiJSON", "Motivazione", "Richiedente", "Stato"]);
    }
    
    foglioRichieste.appendRow([
      new Date(), 
      dati.tipo_azione, // NUOVO, MODIFICA, ELIMINA
      JSON.stringify(dati), 
      dati.motivazione || "", 
      dati.utente_nome, 
      "IN ATTESA"
    ]);
    
    return { success: true, pending: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaContattoRubrica(idRiga, motivazione, utenteNome, isAdmin) {
  try {
     if(String(idRiga).startsWith("CAR_")) return { success: false, error: "Contatti INTERNI non gestibili da qui." };
     
     if (isAdmin) {
       let parts = idRiga.split('_');
       let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
       let foglio = ss.getSheetByName(parts[0]);
       if(foglio) {
          foglio.deleteRow(parseInt(parts[1]) + 1);
          return { success: true };
       } else {
          return { success: false, error: "Scheda non trovata per l'eliminazione." };
       }
     } else {
       // Invia proposta di eliminazione all'admin
       return salvaRichiestaRubrica({
         tipo_azione: "ELIMINA",
         id_riga: idRiga,
         motivazione: motivazione,
         utente_nome: utenteNome
       }, false);
     }
  } catch(e) { return { success: false, error: e.message }; }
}

// ================= FUNZIONI AMMINISTRATORE RUBRICA ================= //

function getRichiesteRubrica(ruoloBase) {
  if (GERARCHIA[ruoloBase] < 5) return []; // Sicurezza: Solo Admin
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RUBRICA');
    if (!foglio) return [];
    
    let dati = foglio.getDataRange().getValues();
    let reqs = [];
    
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][5] === "IN ATTESA") {
        
        // Protezione contro le date vuote o non valide
        let dataFormattata = "Data Sconosciuta";
        try {
            if (dati[i][0]) {
               dataFormattata = (dati[i][0] instanceof Date) 
                   ? Utilities.formatDate(dati[i][0], "GMT+1", "dd/MM/yyyy HH:mm")
                   : String(dati[i][0]);
            }
        } catch(e) {}

        reqs.push({ 
          riga: i + 1, 
          data: dataFormattata, 
          tipo: String(dati[i][1]), 
          datiJSON: String(dati[i][2]), 
          motivo: String(dati[i][3]), 
          richiedente: String(dati[i][4]) 
        });
      }
    }
    return reqs.reverse(); // Le più recenti in alto
  } catch(e) { 
    throw new Error("Errore nel Database Rubrica: " + e.message); 
  }
}

function gestisciRichiestaRubricaAzione(riga, azione, datiAggiornatiJSON) {
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let foglioR = ss.getSheetByName('RICHIESTE_RUBRICA');
    
    if(azione === "APPROVATO") {
      // Se l'admin ha modificato, usiamo il nuovo JSON, altrimenti quello originale salvato
      let jsonStr = datiAggiornatiJSON ? datiAggiornatiJSON : foglioR.getRange(riga, 3).getValue();
      let dati = JSON.parse(jsonStr);
      
      if (dati.tipo_azione === "NUOVO") {
        let targetSheetName = dati.categoria_base;
        if (!['COMANDI', 'LEGALI', 'STAMPA'].includes(targetSheetName)) targetSheetName = 'ALTRO';
        
        let foglioDest = ss.getSheetByName(targetSheetName);
        if (targetSheetName === 'ALTRO') {
          foglioDest.appendRow([new Date(), dati.categoria_specifica, dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
        } else {
          foglioDest.appendRow([new Date(), dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
        }
      } else if (dati.tipo_azione === "ELIMINA") {
         let parts = dati.id_riga.split('_');
         let foglioDest = ss.getSheetByName(parts[0]);
         foglioDest.deleteRow(parseInt(parts[1]) + 1);
      }
      // Se fosse una "MODIFICA" pura (da implementare in futuro), la logica andrebbe qui.
    }
    
    // Aggiorniamo lo stato
    foglioR.getRange(riga, 6).setValue(azione);
    
    // Salviamo lo storico delle modifiche fatte dall'Admin
    if (datiAggiornatiJSON) {
        foglioR.getRange(riga, 3).setValue(datiAggiornatiJSON);
    }
    
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}
// ================= MODULO RUBRICA AVANZATA ================= //

function getContattiRubrica() {
  try {
    let contatti = [];
    // 1. LETTURA DINAMICA INTERNI (Da APP_simccer -> CARICHE)
    let foglioCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    if (foglioCariche) {
      let datiCar = foglioCariche.getDataRange().getValues();
      for (let i = 1; i < datiCar.length; i++) {
        if (!datiCar[i][15]) continue; 
        let provS = String(datiCar[i][6]).toUpperCase().trim();
        let livelloVal = "Provinciale";
        if (provS === "EMILIA ROMAGNA") livelloVal = "Regionale";
        else if (provS === "NAZIONALE") livelloVal = "Nazionale";
        let provUfficio = (provS !== "EMILIA ROMAGNA" && provS !== "NAZIONALE") ? provS : "";

        contatti.push({
          riga: "CAR_" + i,
          categoria: "INTERNI",
          tipo: "Persona",
          nome: String(datiCar[i][3]).toUpperCase().trim() + " " + String(datiCar[i][4]).toUpperCase().trim(),
          ruolo: String(datiCar[i][5]).trim() + " - " + String(datiCar[i][2]).trim(),
          livello: livelloVal,
          provincia: provUfficio,
          compagniaReparto: "", 
          sedeUfficio: String(datiCar[i][8]).trim(),
          telefono: String(datiCar[i][14]).trim(),
          email: String(datiCar[i][13]).toLowerCase().trim()
        });
      }
    }

    // 2. LETTURA ESTERNI DAL DATABASE RUBRICA (Solo quelli già approvati/presenti nelle schede)
    let ssNew = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let fogliEsterni = ['COMANDI', 'LEGALI', 'STAMPA', 'ALTRO'];
    fogliEsterni.forEach(nomeFoglio => {
      let foglio = ssNew.getSheetByName(nomeFoglio);
      if (foglio) {
        let dati = foglio.getDataRange().getValues();
        for (let i = 1; i < dati.length; i++) {
          if (!dati[i][0]) continue;
          let isAltro = (nomeFoglio === 'ALTRO');
          let offset = isAltro ? 1 : 0; 
          contatti.push({
            riga: nomeFoglio + "_" + i,
            categoria: isAltro ? String(dati[i][1]).trim() : nomeFoglio,
            nome: String(dati[i][1 + offset]).toUpperCase().trim(),
            ruolo: String(dati[i][2 + offset]).trim(),
            livello: String(dati[i][3 + offset]).trim(),
            provincia: String(dati[i][4 + offset]).trim(),
            compagniaReparto: String(dati[i][5 + offset]).trim(),
            sedeUfficio: String(dati[i][6 + offset]).trim(),
            telefono: String(dati[i][7 + offset]).trim(),
            email: String(dati[i][8 + offset]).trim()
          });
        }
      }
    });
    return contatti;
  } catch(e) { return { error: e.message }; }
}

// Funzione Unificata per Proposte Utenti e Salvataggi Admin
function salvaRichiestaRubrica(dati, isAdmin) {
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    
    // Se è Admin, salva direttamente nel foglio di destinazione
    if (isAdmin) {
       let targetSheetName = dati.categoria_base; // Es. 'COMANDI'
       let foglio = ss.getSheetByName(targetSheetName);
       if (!foglio) return { success: false, error: "Foglio non trovato nel Database." };
       
       if (targetSheetName === 'ALTRO') {
         foglio.appendRow([new Date(), dati.categoria_specifica, dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
       } else {
         foglio.appendRow([new Date(), dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
       }
       return { success: true };
    } 
    
    // Se è Utente, salva nella scheda di approvazione
    let foglioRichieste = ss.getSheetByName('RICHIESTE_RUBRICA');
    if (!foglioRichieste) {
      foglioRichieste = ss.insertSheet('RICHIESTE_RUBRICA');
      foglioRichieste.appendRow(["Timestamp", "Tipo", "DatiJSON", "Motivazione", "Richiedente", "Stato"]);
    }
    
    foglioRichieste.appendRow([
      new Date(), 
      dati.tipo_azione, // NUOVO, MODIFICA, ELIMINA
      JSON.stringify(dati), 
      dati.motivazione || "", 
      dati.utente_nome, 
      "IN ATTESA"
    ]);
    
    return { success: true, pending: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaContattoRubrica(idRiga, motivazione, utenteNome, isAdmin) {
  try {
     if(String(idRiga).startsWith("CAR_")) return { success: false, error: "Contatti INTERNI non gestibili da qui." };
     
     if (isAdmin) {
       let parts = idRiga.split('_');
       let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
       let foglio = ss.getSheetByName(parts[0]);
       if(foglio) {
          foglio.deleteRow(parseInt(parts[1]) + 1);
          return { success: true };
       } else {
          return { success: false, error: "Scheda non trovata per l'eliminazione." };
       }
     } else {
       // Invia proposta di eliminazione all'admin
       return salvaRichiestaRubrica({
         tipo_azione: "ELIMINA",
         id_riga: idRiga,
         motivazione: motivazione,
         utente_nome: utenteNome
       }, false);
     }
  } catch(e) { return { success: false, error: e.message }; }
}

// ================= FUNZIONI AMMINISTRATORE RUBRICA ================= //

function getRichiesteRubrica(ruoloBase) {
  if (GERARCHIA[ruoloBase] < 5) return []; // Sicurezza: Solo Admin
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RUBRICA');
    if (!foglio) return [];
    
    let dati = foglio.getDataRange().getValues();
    let reqs = [];
    
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][5] === "IN ATTESA") {
        
        // Protezione contro le date vuote o non valide
        let dataFormattata = "Data Sconosciuta";
        try {
            if (dati[i][0]) {
               dataFormattata = (dati[i][0] instanceof Date) 
                   ? Utilities.formatDate(dati[i][0], "GMT+1", "dd/MM/yyyy HH:mm")
                   : String(dati[i][0]);
            }
        } catch(e) {}

        reqs.push({ 
          riga: i + 1, 
          data: dataFormattata, 
          tipo: String(dati[i][1]), 
          datiJSON: String(dati[i][2]), 
          motivo: String(dati[i][3]), 
          richiedente: String(dati[i][4]) 
        });
      }
    }
    return reqs.reverse(); // Le più recenti in alto
  } catch(e) { 
    throw new Error("Errore nel Database Rubrica: " + e.message); 
  }
}

function gestisciRichiestaRubricaAzione(riga, azione, datiAggiornatiJSON) {
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let foglioR = ss.getSheetByName('RICHIESTE_RUBRICA');
    
    if(azione === "APPROVATO") {
      // Se l'admin ha modificato, usiamo il nuovo JSON, altrimenti quello originale salvato
      let jsonStr = datiAggiornatiJSON ? datiAggiornatiJSON : foglioR.getRange(riga, 3).getValue();
      let dati = JSON.parse(jsonStr);
      
      if (dati.tipo_azione === "NUOVO") {
        let targetSheetName = dati.categoria_base;
        if (!['COMANDI', 'LEGALI', 'STAMPA'].includes(targetSheetName)) targetSheetName = 'ALTRO';
        
        let foglioDest = ss.getSheetByName(targetSheetName);
        if (targetSheetName === 'ALTRO') {
          foglioDest.appendRow([new Date(), dati.categoria_specifica, dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
        } else {
          foglioDest.appendRow([new Date(), dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]);
        }
      } else if (dati.tipo_azione === "ELIMINA") {
         let parts = dati.id_riga.split('_');
         let foglioDest = ss.getSheetByName(parts[0]);
         foglioDest.deleteRow(parseInt(parts[1]) + 1);
      }
      // Se fosse una "MODIFICA" pura (da implementare in futuro), la logica andrebbe qui.
    }
    
    // Aggiorniamo lo stato
    foglioR.getRange(riga, 6).setValue(azione);
    
    // Salviamo lo storico delle modifiche fatte dall'Admin
    if (datiAggiornatiJSON) {
        foglioR.getRange(riga, 3).setValue(datiAggiornatiJSON);
    }
    
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}
// Funzione per la modifica diretta di un contatto da parte dell'Admin
function modificaContattoRubrica(idRiga, dati) {
  try {
     if(String(idRiga).startsWith("CAR_")) return { success: false, error: "I contatti INTERNI non possono essere modificati da qui." };
     
     let parts = idRiga.split('_');
     let nomeFoglio = parts[0];
     let rigaReale = parseInt(parts[1]) + 1;
     
     let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
     let foglio = ss.getSheetByName(nomeFoglio);
     
     if (!foglio) return { success: false, error: "Scheda non trovata nel database." };
     
     // Sovrascrive le celle della riga mantenendo il timestamp originale
     if (nomeFoglio === 'ALTRO') {
         // Colonne per la scheda ALTRO: 1:Timestamp, 2:Categoria, 3:Nome, 4:Ruolo, 5:Livello, 6:Provincia, 7:Reparto, 8:Sede, 9:Tel, 10:Email
         foglio.getRange(rigaReale, 2).setValue(dati.categoria_specifica);
         foglio.getRange(rigaReale, 3).setValue(dati.nome);
         foglio.getRange(rigaReale, 4).setValue(dati.ruolo);
         foglio.getRange(rigaReale, 5).setValue(dati.livello);
         foglio.getRange(rigaReale, 6).setValue(dati.provincia);
         foglio.getRange(rigaReale, 7).setValue(dati.compagniaReparto);
         foglio.getRange(rigaReale, 8).setValue(dati.sedeUfficio);
         foglio.getRange(rigaReale, 9).setValue(dati.telefono);
         foglio.getRange(rigaReale, 10).setValue(dati.email);
     } else {
         // Colonne per COMANDI, LEGALI, STAMPA: 1:Timestamp, 2:Nome, 3:Ruolo, 4:Livello, 5:Provincia, 6:Reparto, 7:Sede, 8:Tel, 9:Email
         foglio.getRange(rigaReale, 2).setValue(dati.nome);
         foglio.getRange(rigaReale, 3).setValue(dati.ruolo);
         foglio.getRange(rigaReale, 4).setValue(dati.livello);
         foglio.getRange(rigaReale, 5).setValue(dati.provincia);
         foglio.getRange(rigaReale, 6).setValue(dati.compagniaReparto);
         foglio.getRange(rigaReale, 7).setValue(dati.sedeUfficio);
         foglio.getRange(rigaReale, 8).setValue(dati.telefono);
         foglio.getRange(rigaReale, 9).setValue(dati.email);
     }
     return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}