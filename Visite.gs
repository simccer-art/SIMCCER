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
      
      let showInGestione = false; let showInStorico = false;

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

    let rangeElement = body.findText("{{ELENCO_VISITE}}");
    if (!rangeElement) throw new Error("Segnaposto {{ELENCO_VISITE}} non trovato");
    
    let el = rangeElement.getElement();
    let par = el.getParent();
    if (par.getType() === DocumentApp.ElementType.TEXT) { par = par.getParent(); }
    
    let cellModello = par.getParent();
    if (cellModello.getType() !== DocumentApp.ElementType.TABLE_CELL) {
      throw new Error("Il segnaposto {{ELENCO_VISITE}} DEVE essere dentro una cella di tabella nel Modello Google Docs.");
    }
    
    let rowModello = cellModello.getParent().asTableRow();
    let table = rowModello.getParent().asTable();
    
    let baseFontSize = el.asText().getFontSize() || 11;
    let newFontSize = baseFontSize - 2; 

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
      if (repartiModificati && repartiModificati[rigaIndex]) { repartoFinale = repartiModificati[rigaIndex]; }

      visiteSelezionate.push({ reparto: repartoFinale, dataOrig: dataV, ora: oraV, timestamp: ts, partecipantiGrezzi: String(rigaDati[6]) });
    });

    visiteSelezionate.sort((a, b) => a.timestamp - b.timestamp);

    let indent = "  "; 
    let contattiNotaPieDiPagina = new Map(); 

    visiteSelezionate.forEach((visita) => {
      let newRow = rowModello.copy();
      table.appendTableRow(newRow);
      
      let newCell = newRow.getCell(0);
      newCell.clear(); 
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

    rowModello.removeFromParent();

    if (contattiNotaPieDiPagina.size > 0) {
        let testoNota = "________________________\nContatti:\n" + Array.from(contattiNotaPieDiPagina.values()).join("\n");
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