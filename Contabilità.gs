// ================= MODULO CONTABILITA' ================= //

function getFoglioContabilita(segreteria) {
  let folder = DriveApp.getFolderById(FOLDER_CONTABILITA_ID);
  let fileName = "Contabilità_" + segreteria;
  let files = folder.getFilesByName(fileName);
  
  let ss;
  if (files.hasNext()) {
    ss = SpreadsheetApp.open(files.next());
  } else {
    ss = SpreadsheetApp.create(fileName);
    let file = DriveApp.getFileById(ss.getId());
    file.moveTo(folder); 
  }
  
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
            riga: i+1, dataObj: dMov.getTime(), dataMovimento: Utilities.formatDate(dMov, "GMT+1", "dd/MM/yyyy"), dataIso: isoDate,
            mese: dMov.getMonth() + 1, anno: dMov.getFullYear(), tipo: String(datiMovimenti[i][3] || ""), causale: String(datiMovimenti[i][4] || ""), 
            dettaglio: String(datiMovimenti[i][5] || ""), importo: parseFloat(datiMovimenti[i][6]) || 0, operatore: String(datiMovimenti[i][7] || ""), 
            ordinante: String(datiMovimenti[i][8] || ""), fileUrl: String(datiMovimenti[i][9] || "")
        });
      }
    }
    movimenti.sort((a,b) => b.dataObj - a.dataObj);
    return { error: null, impostazioni: impostazioni, causali: causali, ordinanti: ordinanti, movimenti: movimenti };
  } catch(e) {
    return { error: "Errore Script: " + e.message, impostazioni: {istituto:"",iban:"",cf:""}, causali: [], ordinanti: [], movimenti: [] };
  }
}

function salvaImpostazioniContoContabilita(segreteria, istituto, iban, cf) {
  try {
    let ss = getFoglioContabilita(segreteria);
    let sheet = ss.getSheetByName('IMPOSTAZIONI_CONTO');
    let data = sheet.getDataRange().getValues(); let found = false;
    for(let i=1; i<data.length; i++) {
      if(data[i][0] === segreteria) { 
        sheet.getRange(i+1, 2).setValue(istituto); sheet.getRange(i+1, 3).setValue(iban); sheet.getRange(i+1, 4).setValue(cf); found = true; break; 
      }
    }
    if(!found) sheet.appendRow([segreteria, istituto, iban, cf]);
    SpreadsheetApp.flush(); return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function aggiungiCausaleContabilita(segreteria, causale) { try { let ss = getFoglioContabilita(segreteria); ss.getSheetByName('CAUSALI').appendRow([segreteria, causale]); SpreadsheetApp.flush(); return { success: true }; } catch(e) { return { success: false, error: e.message }; } }
function aggiungiOrdinanteContabilita(segreteria, nominativo) { try { let ss = getFoglioContabilita(segreteria); ss.getSheetByName('ORDINANTI').appendRow([segreteria, nominativo]); SpreadsheetApp.flush(); return { success: true }; } catch(e) { return { success: false, error: e.message }; } }

function eliminaCausaleContabilita(segreteria, causale) {
  try {
    let ss = getFoglioContabilita(segreteria); let sheet = ss.getSheetByName('CAUSALI'); let data = sheet.getDataRange().getValues();
    for(let i = data.length - 1; i >= 1; i--) { if(data[i][0] === segreteria && data[i][1] === causale) { sheet.deleteRow(i + 1); SpreadsheetApp.flush(); return { success: true }; } }
    return { success: false, error: "Causale non trovata." };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaOrdinanteContabilita(segreteria, nominativo) {
  try {
    let ss = getFoglioContabilita(segreteria); let sheet = ss.getSheetByName('ORDINANTI'); let data = sheet.getDataRange().getValues();
    for(let i = data.length - 1; i >= 1; i--) { if(data[i][0] === segreteria && data[i][1] === nominativo) { sheet.deleteRow(i + 1); SpreadsheetApp.flush(); return { success: true }; } }
    return { success: false, error: "Nominativo non trovato." };
  } catch(e) { return { success: false, error: e.message }; }
}

function salvaMovimentoContabilita(segreteria, dataMov, tipo, causale, dettaglio, importo, operatore, ordinante, fileBase64, mimeType, fileName) {
  try {
    let ss = getFoglioContabilita(segreteria); let sheet = ss.getSheetByName('MOVIMENTI'); let fileUrl = "";
    if (fileBase64 && String(fileBase64).trim() !== "" && fileName && String(fileName).trim() !== "") {
      try {
        let folder; try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }
        let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType || 'application/pdf', "Contab_" + new Date().getTime() + "_" + fileName);
        let uploadedFile = folder.createFile(blob); fileUrl = uploadedFile.getUrl();
      } catch(driveErr) { fileUrl = "Errore salvataggio: " + driveErr.message; }
    }
    let parts = String(dataMov).split('-'); let d = new Date(parts[0], parts[1]-1, parts[2]);
    sheet.appendRow([new Date(), segreteria, d, tipo, causale, dettaglio, importo, operatore, ordinante || "", fileUrl]);
    SpreadsheetApp.flush(); return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaMovimentoContabilita(segreteria, riga) { try { getFoglioContabilita(segreteria).getSheetByName('MOVIMENTI').deleteRow(riga); return { success: true }; } catch(e) { return { success: false, error: e.message }; } }

function salvaModificheMultipleContabilita(segreteria, modifiche, operatore) {
  try {
      let ss = getFoglioContabilita(segreteria); let sheet = ss.getSheetByName('MOVIMENTI');
      modifiche.forEach(m => {
          let parts = m.data.split('-'); let d = new Date(parts[0], parts[1]-1, parts[2]);
          sheet.getRange(m.riga, 3).setValue(d); sheet.getRange(m.riga, 4).setValue(m.tipo); sheet.getRange(m.riga, 5).setValue(m.causale);
          sheet.getRange(m.riga, 6).setValue(m.dettaglio); sheet.getRange(m.riga, 7).setValue(m.importo); sheet.getRange(m.riga, 8).setValue(operatore); sheet.getRange(m.riga, 9).setValue(m.ordinante || "");
      });
      SpreadsheetApp.flush(); return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function getLogoBase64(segreteria) {
  try {
      let folder = DriveApp.getFolderById(CARTELLA_LOGHI_ID);
      let provName = segreteria.replace("Provinciale (", "").replace("Regionale (", "").replace(")", "").trim();
      let files = folder.searchFiles("title contains '" + provName + "'");
      if (files.hasNext()) { let f = files.next(); return `data:${f.getMimeType()};base64,${Utilities.base64Encode(f.getBlob().getBytes())}`; }
      let allFiles = folder.getFiles();
      if (allFiles.hasNext()) { let f = allFiles.next(); return `data:${f.getMimeType()};base64,${Utilities.base64Encode(f.getBlob().getBytes())}`; }
  } catch(e) {} return "";
}

function generaPDFContabilita(segreteria, mese, anno, operatore, dataDa, dataA) {
  try {
    let dati = getDatiContabilita(segreteria);
    let movimenti = dati.movimenti.filter(m => {
        let matchMese = (mese === 'TUTTI' || String(m.mese) === mese); let matchAnno = (anno === 'TUTTI' || String(m.anno) === anno);
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
    const nomiMesi = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"];
    let nomeMeseTesto = (mese === "TUTTI") ? "TUTTI I MESI" : nomiMesi[parseInt(mese)];
    let strPeriodo = `Mese: ${nomeMeseTesto} - Anno: ${anno}`;
    if (dataDa || dataA) strPeriodo = `Dal: ${dataDa ? formattaD(dataDa) : 'Inizio'} al: ${dataA ? formattaD(dataA) : 'Fine'}`;

    let html = `<style>@page { margin: 1cm; }</style><div style="font-family: Arial, sans-serif; font-size:10px;">${imgTag}<div style="text-align:center; margin-bottom:15px;"><h3 style="margin:5px 0 0 0; font-size:16px; font-weight:normal;">Sindacato Italiano Militari Carabinieri</h3><h4 style="margin:0; font-size:12px;">${parsedSegr}</h4></div><h2 style="text-align:center; color:#1a4b84; margin-top:20px; margin-bottom:5px;">REGISTRO CONTABILE</h2><p style="text-align:center; color:#1a4b84; font-weight:bold; margin-top:0;">Periodo di riferimento: ${strPeriodo}</p><div style="margin-bottom:15px; padding:10px; background:#f4f6f8; border:1px solid #ccc; font-size:10px;"><p style="margin:0;"><strong>Istituto:</strong> ${dati.impostazioni.istituto || "-"} | <strong>IBAN:</strong> ${dati.impostazioni.iban || "-"} | <strong>C.F.:</strong> ${dati.impostazioni.cf || "-"}</p></div><table style="width:100%; border-collapse:collapse; font-size:9px;" border="1" cellpadding="3"><tr style="background:#eee;"><th>N.</th><th>Data</th><th>Tipo</th><th>Ordinante/Beneficiario</th><th>Causale</th><th>Dettaglio</th><th>Importo</th></tr>${htmlRows}</table><br><table style="width:50%; margin-left:auto; border-collapse:collapse; font-size:10px;" border="1" cellpadding="4"><tr><td><strong>Totale Entrate:</strong></td><td style="color:green; text-align:right;">€ ${totIn.toFixed(2)}</td></tr><tr><td><strong>Totale Uscite:</strong></td><td style="color:red; text-align:right;">€ ${totOut.toFixed(2)}</td></tr><tr><td style="background:#fff3cd;"><strong>SALDO PERIODO:</strong></td><td style="text-align:right; background:#fff3cd;"><strong>€ ${(totIn - totOut).toFixed(2)}</strong></td></tr></table><div style="text-align:right; margin-top:40px; line-height:0.8;"><p style="font-size:12px; margin: 2px 0;">IL TESORIERE SIM CC</p><p style="font-size:12px; margin: 2px 0;"><em>${operatore}</em></p><p style="color:red; font-size:9px; margin: 2px 0; margin-top:10px;"><strong>[ DOCUMENTO VERIFICATO E FIRMATO ]</strong></p></div></div>`;
    let outputHtml = HtmlService.createHtmlOutput(html);
    let cartella; try { cartella = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { cartella = DriveApp.getRootFolder(); }
    let pdfFile = cartella.createFile(outputHtml.getAs('application/pdf').setName("Registro_Contabile_" + segreteria.replace(/\s+/g,"_") + "_" + new Date().getTime() + ".pdf"));
    return {success: true, url: pdfFile.getUrl()};
  } catch(e) { return {success: false, error: e.message}; }
}