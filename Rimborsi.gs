// ================= MODULO RIMBORSI ================= //

function salvaNuovoRimborso(dati, fileBase64, mimeType, fileName) {
  try {
    let folder; 
    try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } 
    catch(e) { folder = DriveApp.getRootFolder(); }
    
    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, "Firmato_" + fileName);
    let uploadedFile = folder.createFile(blob);
    let fileUrl = uploadedFile.getUrl();

    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    
    if (!foglio) {
      foglio = ss.insertSheet('RICHIESTE_RIMBORSI');
      foglio.appendRow(["Timestamp", "CIP", "Richiedente", "Provincia", "Dati JSON", "URL File", "Stato", "Note", "Operatore"]);
    }

    foglio.appendRow([ new Date(), dati.cip, dati.richiedente, dati.provincia, JSON.stringify(dati), fileUrl, "IN ATTESA", "", "" ]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
                riga: i+1, dataInvio: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy HH:mm"), 
                dataSpesa: json.dataOdierna || json.dataSpesa, importo: parseFloat(importo).toFixed(2), 
                importoAutorizzato: json.importoAutorizzato !== undefined ? parseFloat(json.importoAutorizzato).toFixed(2) : parseFloat(importo).toFixed(2), 
                descrizione: descrizione, fileUrl: dati[i][5], stato: dati[i][6], note: dati[i][7], 
                operatore: dati[i][8], protocollo: json.protocollo || '-', segreteriaPagante: json.segreteriaPagante || '-', jsonStr: dati[i][4] 
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
    
    let isTesoriere1 = ruoloStr.includes('TESORIERE') && !ruoloStr.includes('TESORIERE2');
    let isTesoriere2 = ruoloStr.includes('TESORIERE2');
    let provLoggato = utenteLoggato.provincia;

    let daGestire = []; let storico = [];
    for(let i=1; i<dati.length; i++) {
        if(!dati[i][0]) continue;
        let cip = String(dati[i][1]).trim(); let richiedente = dati[i][2]; let prov = dati[i][3];
        let json = JSON.parse(dati[i][4]); let fileUrl = dati[i][5]; let stato = String(dati[i][6]).toUpperCase().trim();
        let note = dati[i][7]; let operatore = dati[i][8];

        let importo = json.totaleGenerale !== undefined ? json.totaleGenerale : (json.importo || 0);
        let descrizione = json.motivo || json.descrizione || "Rimborso spese multiple";
        
        let sp = json.segreteriaPagante || "";
        let isProvinciale = sp.startsWith("Provinciale"); let isRegionale = sp.startsWith("Regionale") || sp === "Nazionale";

        let targetProv = prov; 
        if (isProvinciale) {
            let match = sp.match(/\(([^)]+)\)/);
            if (match) targetProv = match[1].toUpperCase().trim();
        }

        let seesIt = false; let canApprove = false; let canPay = false;
        
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
        
        if(!seesIt) continue;

        let p = { riga: i+1, dataInvio: Utilities.formatDate(new Date(dati[i][0]), "GMT+1", "dd/MM/yyyy"), cip: cip, richiedente: richiedente, provincia: prov, dataSpesa: json.dataOdierna || json.dataSpesa, importo: parseFloat(importo).toFixed(2), importoAutorizzato: json.importoAutorizzato !== undefined ? parseFloat(json.importoAutorizzato).toFixed(2) : parseFloat(importo).toFixed(2), descrizione: descrizione, fileUrl: fileUrl, stato: stato, note: note, operatore: operatore, iban: json.iban, vettura: json.vettura, segreteriaPagante: sp || "-", protocollo: json.protocollo || "-", canApprove: canApprove, canPay: canPay };
        
        if (stato === 'IN ATTESA' && canApprove) daGestire.push(p);
        else if (stato === 'APPROVATO' && canPay) daGestire.push(p);
        else if (stato !== 'IN ATTESA ALLEGATO') storico.push(p);
    }
    return { daGestire: daGestire.reverse(), storico: storico.reverse() };
  } catch(e) { return {daGestire:[], storico:[]}; }
}

function compilaTemplateRimborsoDoc(payload, mostraFirma, mostraTimbroPagato) {
  try {
    let folder;
    try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }
    
    let safeProt = payload.protocollo ? payload.protocollo.replace(/\//g, '-') : "SenzaProt";
    let nomeFile = "Richiesta_Rimborso_" + safeProt + "_" + payload.cip;
    
    let copiaId = DriveApp.getFileById(TEMPLATE_RIMBORSO_DOC_ID).makeCopy(nomeFile, folder).getId();
    let doc = DocumentApp.openById(copiaId);
    let body = doc.getBody();

    try {
      let logoFolder = DriveApp.getFolderById(CARTELLA_LOGHI_ID);
      let provName = payload.provincia.replace("Provinciale (", "").replace("Regionale (", "").replace(")", "").trim();
      let files = logoFolder.searchFiles("title contains '" + provName + "'");
      let logoBlob = null;
      
      if (files.hasNext()) { logoBlob = files.next().getBlob(); } 
      else { let allLogos = logoFolder.getFiles(); if (allLogos.hasNext()) logoBlob = allLogos.next().getBlob(); }

      if (logoBlob) {
        let elementoLogo = body.findText("{{LOGO}}");
        if (elementoLogo) {
          let im = elementoLogo.getElement().getParent().asParagraph().appendInlineImage(logoBlob);
          let origW = im.getWidth(); let origH = im.getHeight();
          let maxW = 200; let maxH = 80;  
          
          if (origW > 0 && origH > 0) {
            let ratio = origW / origH;
            if (origW > maxW || origH > maxH) {
              if (ratio > (maxW / maxH)) { im.setWidth(maxW); im.setHeight(Math.round(maxW / ratio)); } 
              else { im.setHeight(maxH); im.setWidth(Math.round(maxH * ratio)); }
            }
          } else { im.setWidth(150).setHeight(75); }
          body.replaceText("{{LOGO}}", ""); 
        }
      }
    } catch (e) { console.log("Errore logo: " + e.message); }

    let sp = payload.segreteriaPagante || ""; let testoSegreteria = "";
    function toTitleCase(str) { return str.toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' '); }
    
    if (sp.includes("Provinciale")) {
      let prov = sp.match(/\(([^)]+)\)/); testoSegreteria = "Segreteria Provinciale di " + (prov ? toTitleCase(prov[1]) : "");
    } else if (sp.includes("Regionale")) {
      let reg = sp.match(/\(([^)]+)\)/); testoSegreteria = "Segreteria Regionale " + (reg ? toTitleCase(reg[1]) : "");
    } else { testoSegreteria = toTitleCase(sp); }
    body.replaceText("{{SEGRETERIA}}", "Sindacato Italiano Militari Carabinieri\n" + testoSegreteria);

    body.replaceText("{{NOME_COGNOME}}", payload.richiedente || "");
    body.replaceText("{{PROTOCOLLO}}", payload.protocollo || "");
    body.replaceText("{{DATA_PROT}}", payload.dataOdierna || "");
    body.replaceText("{{CF}}", payload.cf || "");
    body.replaceText("{{MOTIVO}}", (payload.motivo || "").toUpperCase());
    body.replaceText("{{IBAN}}", payload.iban || "");
    
    let mezzo = payload.vettura || ""; let marca = "", modello = "", targa = "";
    if (mezzo.includes("-")) {
      let partiMezzo = mezzo.split("-"); targa = partiMezzo[1].trim().toUpperCase();
      let marcaModello = partiMezzo[0].trim().split(" "); marca = marcaModello[0].toUpperCase(); modello = marcaModello.slice(1).join(" ").toUpperCase();
    }
    body.replaceText("{{MARCA}}", marca); body.replaceText("{{MODELLO}}", modello); body.replaceText("{{TARGA}}", targa);

    for (let i = 1; i <= 6; i++) {
      let tratta = payload.tratte[i - 1];
      body.replaceText("{{DATA_" + i + "}}", tratta ? tratta.data : "");
      body.replaceText("{{ITIN_" + i + "}}", tratta ? String(tratta.itinerario).toUpperCase() : "");
      body.replaceText("{{KM_" + i + "}}", tratta ? tratta.km : "");
    }
    body.replaceText("{{TOT_KM}}", payload.totaleKm || "0");
    body.replaceText("{{IMP_A}}", parseFloat(payload.importoA).toFixed(2));

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

    if (mostraTimbroPagato) {
      let impAut = (payload.importoAutorizzato !== undefined && payload.importoAutorizzato !== null && payload.importoAutorizzato !== "") 
                   ? parseFloat(payload.importoAutorizzato) 
                   : parseFloat(payload.totaleGenerale);
      body.replaceText("{{IMP_AUTORIZZATO}}", "€ " + impAut.toFixed(2));

      let sp = payload.segreteriaPagante || ""; let luogoStr = "";
      let match = sp.match(/\(([^)]+)\)/);
      if (match) { luogoStr = match[1].toUpperCase(); if (luogoStr === "EMILIA ROMAGNA") luogoStr = "BOLOGNA"; } 
      else { luogoStr = "ROMA"; }

      let dataOggiStr = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");
      let tesoriereNome = payload.nomeTesoriere ? payload.nomeTesoriere.toUpperCase() : "________________";

      let timbroText = "PAGAMENTO AUTORIZZATO\nEURO " + impAut.toFixed(2) + "\nIL TESORIERE " + tesoriereNome + "\n" + luogoStr + ", " + dataOggiStr;

      let rangeTimbro = body.findText("{{TIMBRO_ARANCIONE}}");
      let table;
      if (rangeTimbro) {
          let el = rangeTimbro.getElement();
          let container = el.getParent();
          while (container && container.getType() !== DocumentApp.ElementType.TABLE_CELL && container.getType() !== DocumentApp.ElementType.BODY_SECTION) {
              container = container.getParent();
          }
          el.asText().replaceText("{{TIMBRO_ARANCIONE}}", "");
          if (container && container.getType() === DocumentApp.ElementType.TABLE_CELL) {
              table = container.asTableCell().appendTable([ [""] ]);
          } else {
              let par = el.getParent();
              if (par.getType() === DocumentApp.ElementType.TEXT) par = par.getParent();
              let childIndex = body.getChildIndex(par);
              table = body.insertTable(childIndex + 1, [ [""] ]);
          }
      } else {
          body.appendParagraph(""); 
          table = body.appendTable([ [""] ]);
      }

      table.setBorderColor("#FF8C00"); table.setBorderWidth(1); table.setColumnWidth(0, 155.90); 
      let row = table.getRow(0); row.setMinimumHeight(42.52); 
      let cell = row.getCell(0); cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      cell.setPaddingTop(2).setPaddingBottom(2).setPaddingLeft(4).setPaddingRight(2);
      cell.clear(); 
      let par = cell.appendParagraph(timbroText);
      par.setAlignment(DocumentApp.HorizontalAlignment.LEFT); par.setLineSpacing(1.0); par.setSpacingBefore(0); par.setSpacingAfter(0);  
      let textEl = par.editAsText();
      textEl.setBold(true); textEl.setForegroundColor("#FF8C00"); textEl.setFontSize(7.5); 
    } else {
      body.replaceText("{{IMP_AUTORIZZATO}}", ""); body.replaceText("{{TIMBRO_ARANCIONE}}", ""); 
    }

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
    if (importoAutorizzato !== undefined && importoAutorizzato !== null) { datiOriginali.importoAutorizzato = importoAutorizzato; }
    
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
    
    let datiJson = JSON.parse(foglio.getRange(riga, 5).getValue());
    if (importoAutorizzato !== undefined && importoAutorizzato !== null) { datiJson.importoAutorizzato = importoAutorizzato; }
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

function salvaRimborsoSenzaFile(dati) {
  try {
    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    if (!foglio) {
      foglio = ss.insertSheet('RICHIESTE_RIMBORSI');
      foglio.appendRow(["Timestamp", "CIP", "Richiedente", "Provincia", "Dati JSON", "URL File", "Stato", "Note", "Operatore"]);
    }
    foglio.appendRow([ new Date(), dati.cip, dati.richiedente, dati.provincia, JSON.stringify(dati), "", "IN ATTESA ALLEGATO", "", "" ]);
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function allegaFileRimborso(riga, cip, fileBase64, mimeType, fileName) {
  try {
    let ss = SpreadsheetApp.openById(GESTIONALE_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RIMBORSI');
    
    if(String(foglio.getRange(riga, 2).getValue()).toUpperCase().trim() !== String(cip).toUpperCase().trim()) {
        return {success: false, error: "Accesso negato o riga non corrispondente"};
    }

    let folder; 
    try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }

    let blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), mimeType, "Firmato_" + fileName);
    let uploadedFile = folder.createFile(blob);
    let fileUrl = uploadedFile.getUrl();

    foglio.getRange(riga, 6).setValue(fileUrl);
    foglio.getRange(riga, 7).setValue("IN ATTESA");

    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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

function getFirmaMutriBase64() {
  try {
    let folder = DriveApp.getFolderById(CARTELLA_LOGHI_ID);
    let files = folder.getFilesByName('Mutri.png');
    if (files.hasNext()) {
      let f = files.next();
      return `data:${f.getMimeType()};base64,${Utilities.base64Encode(f.getBlob().getBytes())}`;
    }
  } catch(e) {}
  return "";
}

function getFirmaMutriBlob() {
  try {
    let folder = DriveApp.getFolderById(CARTELLA_LOGHI_ID);
    let files = folder.getFilesByName('Mutri.png');
    if (files.hasNext()) { return files.next().getBlob(); }
  } catch(e) {}
  return null;
}