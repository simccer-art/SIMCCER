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
      let partiInizio = p.inizio.split('-');
      let dateInizio = new Date(partiInizio[0], partiInizio[1] - 1, partiInizio[2], 12, 0, 0);
      let partiFine = p.fine.split('-');
      let dateFine = new Date(partiFine[0], partiFine[1] - 1, partiFine[2], 12, 0, 0);

      foglio.appendRow([ 
        new Date(), datiRichiesta.cip, `${datiRichiesta.grado} ${datiRichiesta.nome} ${datiRichiesta.cognome}`, 
        datiRichiesta.provincia, dateInizio, dateFine, p.urgente ? "SI" : "NO", p.supera ? "SI" : "NO", 
        stato, p.motivazione || "", inseritoDa, operatore 
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

      if (ruoloBase === "SEGRGEN" && p.provincia === provincia) { 
          if (stato === "IN ATTESA") daGestire.push(p); else storico.push(p); 
      } else if (ruoloBase === "APPROVATORE") { 
          if (stato === "VISTATO") daGestire.push(p); else storico.push(p); 
      } else if (ruoloBase === "GESTORE") { 
          if (stato === "APPROVATO") daGestire.push(p); else storico.push(p); 
      } else if (ruoloBase === "AMMINISTRATORE") { 
          if (stato === "IN ATTESA" || stato === "VISTATO" || stato === "APPROVATO") daGestire.push(p); else storico.push(p); 
      }
    }
    return { daGestire: daGestire, storico: storico, cipOrder: cipOrder };
  } catch(e) { throw new Error(e.message); }
}

function aggiornaStatoPratica(riga, nuovoStato, note, operatore) {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI');
    foglio.getRange(riga, 9).setValue(nuovoStato);
    if (operatore) { foglio.getRange(riga, 12).setValue(operatore); }
    if (note) {
      let notaAttuale = foglio.getRange(riga, 10).getValue();
      let nuovaNota = notaAttuale ? notaAttuale + " | Note op: " + note : note;
      foglio.getRange(riga, 10).setValue(nuovaNota);
    }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
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
  } catch(e) { return { success: false, error: e.message }; }
}

function generaPDFPratica(datiPratica) {
  try {
    const TEMPLATE_ID = '1SekjRCrllxkY2N-zLOHgBef_8IAcu_MpLra0cRJ_Qw8';
    let folder; try { folder = DriveApp.getFolderById(FOLDER_REPORTS_ID); } catch(e) { folder = DriveApp.getRootFolder(); }

    const sheetCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    const datiCariche = sheetCariche.getDataRange().getValues();
    
    let grado = ""; let nome = ""; let cognome = ""; 
    let reparto = ""; let dataNomina = ""; let carica = ""; let organizzazione = "";

    for(let i = 1; i < datiCariche.length; i++) {
        if(String(datiCariche[i][15]).toUpperCase().trim() === String(datiPratica.cip).toUpperCase().trim()) {
            grado = String(datiCariche[i][2]).trim(); cognome = String(datiCariche[i][3]).trim();   
            nome = String(datiCariche[i][4]).trim(); carica = String(datiCariche[i][5]).trim();    
            reparto = String(datiCariche[i][8]).trim();   
            
            let valDataNomina = datiCariche[i][9];
            if (valDataNomina instanceof Date) { dataNomina = Utilities.formatDate(valDataNomina, "GMT+1", "dd/MM/yyyy"); } 
            else if (valDataNomina && String(valDataNomina).trim() !== "") { dataNomina = String(valDataNomina).trim(); } 
            else { dataNomina = "N/D"; }
            
            organizzazione = String(datiCariche[i][19]).trim();
            break;
        }
    }

    grado = grado.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    nome = nome.toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
    cognome = cognome.toUpperCase();
    let richiedenteFormattato = (grado || nome || cognome) ? `${grado} ${nome} ${cognome}` : datiPratica.richiedente;

    let comandoCorpo = ""; let ufficio = ""; let emailComando = ""; let pecComando = "";
    if (organizzazione !== "") {
        const sheetIndirizzi = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('INDIRIZZI');
        if (sheetIndirizzi) {
            const datiIndirizzi = sheetIndirizzi.getDataRange().getValues();
            for(let j = 1; j < datiIndirizzi.length; j++) {
                if (String(datiIndirizzi[j][0]).toUpperCase().trim() === organizzazione.toUpperCase()) {
                    comandoCorpo = String(datiIndirizzi[j][1]).trim(); ufficio = String(datiIndirizzi[j][2]).trim(); 
                    emailComando = String(datiIndirizzi[j][3]).trim(); pecComando = String(datiIndirizzi[j][4]).trim(); 
                    break; 
                }
            }
        }
    }

    let safeProt = datiPratica.protocollo ? datiPratica.protocollo.replace(/\//g, '!') : "Senza_Prot";
    let safeCognome = cognome ? cognome : datiPratica.richiedente.split(' ').pop().toUpperCase();
    const nomeFile = "Lettera_permesso_" + safeCognome + "_" + safeProt;
    
    const tempDocFile = DriveApp.getFileById(TEMPLATE_ID).makeCopy(nomeFile, folder);
    const doc = DocumentApp.openById(tempDocFile.getId());
    const body = doc.getBody();
    
    let periodiFormattati = datiPratica.periodi.map(p => (p.inizio === p.fine) ? p.inizio : "dal " + p.inizio + " al " + p.fine);
    let chunksPeriodi = [];
    for (let k = 0; k < periodiFormattati.length; k += 5) { chunksPeriodi.push(periodiFormattati.slice(k, k + 5).join(" - ")); }
    let strPeriodi = chunksPeriodi.join(" -\n");
    
    let protocolloStr = datiPratica.protocollo ? datiPratica.protocollo : "___________";
    let dataOdierna = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy");

    body.replaceText("{{PROTOCOLLO}}", protocolloStr); body.replaceText("{{RICHIEDENTE}}", richiedenteFormattato); 
    body.replaceText("{{GRADO}}", grado); body.replaceText("{{NOME}}", nome); body.replaceText("{{COGNOME}}", cognome);
    body.replaceText("{{CIP}}", datiPratica.cip || ""); body.replaceText("{{PROVINCIA}}", datiPratica.provincia || "");
    body.replaceText("{{PERIODI}}", strPeriodi); body.replaceText("{{DATA_OGGI}}", dataOdierna);
    body.replaceText("{{REPARTO}}", reparto); body.replaceText("{{DATA_NOMINA}}", dataNomina);
    body.replaceText("{{CARICA}}", carica); body.replaceText("{{COMANDO_CORPO}}", comandoCorpo);
    body.replaceText("{{UFFICIO}}", ufficio); body.replaceText("{{EMAIL_COMANDO}}", emailComando); body.replaceText("{{PEC_COMANDO}}", pecComando);
    
    doc.saveAndClose(); 
    const pdfBlob = tempDocFile.getAs('application/pdf'); const pdfFile = folder.createFile(pdfBlob);
    tempDocFile.setTrashed(true);
    
    return { success: true, url: pdfFile.getUrl() };
  } catch(e) { return { success: false, error: e.message }; }
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

function cancellaGiornoPratica(riga) { try { SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('RICHIESTE_PERMESSI').deleteRow(riga); return { success: true }; } catch(e) { return { success: false, error: e.message }; } }

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

      mieRichieste.push({ riga: i + 1, idGruppo: Utilities.formatDate(dReqObj, "GMT+1", "ddMMyyyy_HHmm"), dataRichiesta: Utilities.formatDate(dReqObj, "GMT+1", "dd/MM/yyyy"), inizio: Utilities.formatDate(dInizioObj, "GMT+1", "dd/MM/yyyy"), fine: Utilities.formatDate(new Date(dati[i][5]), "GMT+1", "dd/MM/yyyy"), urgente: dati[i][6], limite: dati[i][7], stato: stato, motivazione: dati[i][9] || "", isScaduta: isScaduta, inseritoDa: String(dati[i][10]||"").trim(), operatore: String(dati[i][11]||"").trim() });
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