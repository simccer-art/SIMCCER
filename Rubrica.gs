// ================= MODULO RUBRICA ================= //

function getDatiRubrica() {
  try {
    const foglio = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    if (!foglio) return { error: "Foglio CARICHE non trovato." };
    
    const dati = foglio.getDataRange().getValues();
    let contatti = [];
    
    for (let i = 1; i < dati.length; i++) {
      if (!dati[i][15]) continue;
      
      let extraInfo = {};
      try { if(dati[i][20]) extraInfo = JSON.parse(dati[i][20]); } catch(e) {}
      
      contatti.push({
        grado: String(dati[i][2]).trim(), cognome: String(dati[i][3]).toUpperCase().trim(), nome: String(dati[i][4]).toUpperCase().trim(),
        carica: String(dati[i][5]).trim(), provincia: String(dati[i][6]).toUpperCase().trim(), reparto: String(dati[i][8]).trim(),
        email: String(dati[i][13]).toLowerCase().trim(), telefono: String(dati[i][14]).trim(), fotoUrl: extraInfo.foto || ""
      });
    }
    contatti.sort((a, b) => (a.cognome + " " + a.nome).localeCompare(b.cognome + " " + b.nome));
    return contatti;
  } catch(e) { return { error: e.message }; }
}

function getContattiRubrica() {
  try {
    let contatti = [];
    let foglioCariche = SpreadsheetApp.openById(GESTIONALE_ID).getSheetByName('CARICHE');
    if (foglioCariche) {
      let datiCar = foglioCariche.getDataRange().getValues();
      for (let i = 1; i < datiCar.length; i++) {
        if (!datiCar[i][15]) continue; 
        let provS = String(datiCar[i][6]).toUpperCase().trim();
        let livelloVal = "Provinciale";
        if (provS === "EMILIA ROMAGNA") livelloVal = "Regionale"; else if (provS === "NAZIONALE") livelloVal = "Nazionale";
        let provUfficio = (provS !== "EMILIA ROMAGNA" && provS !== "NAZIONALE") ? provS : "";

        contatti.push({
          riga: "CAR_" + i, categoria: "INTERNI", tipo: "Persona", nome: String(datiCar[i][3]).toUpperCase().trim() + " " + String(datiCar[i][4]).toUpperCase().trim(),
          ruolo: String(datiCar[i][5]).trim() + " - " + String(datiCar[i][2]).trim(), livello: livelloVal, provincia: provUfficio,
          compagniaReparto: "", sedeUfficio: String(datiCar[i][8]).trim(), telefono: String(datiCar[i][14]).trim(), email: String(datiCar[i][13]).toLowerCase().trim()
        });
      }
    }

    let ssNew = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let fogliEsterni = ['COMANDI', 'LEGALI', 'STAMPA', 'ALTRO'];
    fogliEsterni.forEach(nomeFoglio => {
      let foglio = ssNew.getSheetByName(nomeFoglio);
      if (foglio) {
        let dati = foglio.getDataRange().getValues();
        for (let i = 1; i < dati.length; i++) {
          if (!dati[i][0]) continue;
          let isAltro = (nomeFoglio === 'ALTRO'); let offset = isAltro ? 1 : 0; 
          contatti.push({
            riga: nomeFoglio + "_" + i, categoria: isAltro ? String(dati[i][1]).trim() : nomeFoglio,
            nome: String(dati[i][1 + offset]).toUpperCase().trim(), ruolo: String(dati[i][2 + offset]).trim(),
            livello: String(dati[i][3 + offset]).trim(), provincia: String(dati[i][4 + offset]).trim(),
            compagniaReparto: String(dati[i][5 + offset]).trim(), sedeUfficio: String(dati[i][6 + offset]).trim(),
            telefono: String(dati[i][7 + offset]).trim(), email: String(dati[i][8 + offset]).trim()
          });
        }
      }
    });
    return contatti;
  } catch(e) { return { error: e.message }; }
}

function salvaRichiestaRubrica(dati, isAdmin) {
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    if (isAdmin) {
       let targetSheetName = dati.categoria_base;
       let foglio = ss.getSheetByName(targetSheetName);
       if (!foglio) return { success: false, error: "Foglio non trovato nel Database." };
       
       if (targetSheetName === 'ALTRO') { foglio.appendRow([new Date(), dati.categoria_specifica, dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]); } 
       else { foglio.appendRow([new Date(), dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]); }
       return { success: true };
    } 
    
    let foglioRichieste = ss.getSheetByName('RICHIESTE_RUBRICA');
    if (!foglioRichieste) {
      foglioRichieste = ss.insertSheet('RICHIESTE_RUBRICA');
      foglioRichieste.appendRow(["Timestamp", "Tipo", "DatiJSON", "Motivazione", "Richiedente", "Stato"]);
    }
    
    foglioRichieste.appendRow([ new Date(), dati.tipo_azione, JSON.stringify(dati), dati.motivazione || "", dati.utente_nome, "IN ATTESA" ]);
    return { success: true, pending: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function eliminaContattoRubrica(idRiga, motivazione, utenteNome, isAdmin) {
  try {
     if(String(idRiga).startsWith("CAR_")) return { success: false, error: "Contatti INTERNI non gestibili da qui." };
     
     if (isAdmin) {
       let parts = idRiga.split('_'); let ss = SpreadsheetApp.openById(RUBRICA_DB_ID); let foglio = ss.getSheetByName(parts[0]);
       if(foglio) { foglio.deleteRow(parseInt(parts[1]) + 1); return { success: true }; } 
       else { return { success: false, error: "Scheda non trovata." }; }
     } else {
       return salvaRichiestaRubrica({ tipo_azione: "ELIMINA", id_riga: idRiga, motivazione: motivazione, utente_nome: utenteNome }, false);
     }
  } catch(e) { return { success: false, error: e.message }; }
}

function getRichiesteRubrica(ruoloBase) {
  if (GERARCHIA[ruoloBase] < 5) return []; 
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let foglio = ss.getSheetByName('RICHIESTE_RUBRICA');
    if (!foglio) return [];
    
    let dati = foglio.getDataRange().getValues(); let reqs = [];
    
    for (let i = 1; i < dati.length; i++) {
      if (dati[i][5] === "IN ATTESA") {
        let dataFormattata = "Data Sconosciuta";
        try { if (dati[i][0]) { dataFormattata = (dati[i][0] instanceof Date) ? Utilities.formatDate(dati[i][0], "GMT+1", "dd/MM/yyyy HH:mm") : String(dati[i][0]); } } catch(e) {}
        reqs.push({ riga: i + 1, data: dataFormattata, tipo: String(dati[i][1]), datiJSON: String(dati[i][2]), motivo: String(dati[i][3]), richiedente: String(dati[i][4]) });
      }
    }
    return reqs.reverse(); 
  } catch(e) { throw new Error("Errore nel Database Rubrica: " + e.message); }
}

function gestisciRichiestaRubricaAzione(riga, azione, datiAggiornatiJSON) {
  try {
    let ss = SpreadsheetApp.openById(RUBRICA_DB_ID);
    let foglioR = ss.getSheetByName('RICHIESTE_RUBRICA');
    
    if(azione === "APPROVATO") {
      let jsonStr = datiAggiornatiJSON ? datiAggiornatiJSON : foglioR.getRange(riga, 3).getValue();
      let dati = JSON.parse(jsonStr);
      
      if (dati.tipo_azione === "NUOVO") {
        let targetSheetName = dati.categoria_base;
        if (!['COMANDI', 'LEGALI', 'STAMPA'].includes(targetSheetName)) targetSheetName = 'ALTRO';
        
        let foglioDest = ss.getSheetByName(targetSheetName);
        if (targetSheetName === 'ALTRO') { foglioDest.appendRow([new Date(), dati.categoria_specifica, dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]); } 
        else { foglioDest.appendRow([new Date(), dati.nome, dati.ruolo, dati.livello, dati.provincia, dati.compagniaReparto, dati.sedeUfficio, dati.telefono, dati.email]); }
      } else if (dati.tipo_azione === "ELIMINA") {
         let parts = dati.id_riga.split('_'); let foglioDest = ss.getSheetByName(parts[0]);
         foglioDest.deleteRow(parseInt(parts[1]) + 1);
      }
    }
    
    foglioR.getRange(riga, 6).setValue(azione);
    if (datiAggiornatiJSON) { foglioR.getRange(riga, 3).setValue(datiAggiornatiJSON); }
    return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}

function modificaContattoRubrica(idRiga, dati) {
  try {
     if(String(idRiga).startsWith("CAR_")) return { success: false, error: "Contatti INTERNI non modificabili da qui." };
     let parts = idRiga.split('_'); let nomeFoglio = parts[0]; let rigaReale = parseInt(parts[1]) + 1;
     let ss = SpreadsheetApp.openById(RUBRICA_DB_ID); let foglio = ss.getSheetByName(nomeFoglio);
     
     if (!foglio) return { success: false, error: "Scheda non trovata." };
     
     if (nomeFoglio === 'ALTRO') {
         foglio.getRange(rigaReale, 2).setValue(dati.categoria_specifica); foglio.getRange(rigaReale, 3).setValue(dati.nome);
         foglio.getRange(rigaReale, 4).setValue(dati.ruolo); foglio.getRange(rigaReale, 5).setValue(dati.livello);
         foglio.getRange(rigaReale, 6).setValue(dati.provincia); foglio.getRange(rigaReale, 7).setValue(dati.compagniaReparto);
         foglio.getRange(rigaReale, 8).setValue(dati.sedeUfficio); foglio.getRange(rigaReale, 9).setValue(dati.telefono); foglio.getRange(rigaReale, 10).setValue(dati.email);
     } else {
         foglio.getRange(rigaReale, 2).setValue(dati.nome); foglio.getRange(rigaReale, 3).setValue(dati.ruolo);
         foglio.getRange(rigaReale, 4).setValue(dati.livello); foglio.getRange(rigaReale, 5).setValue(dati.provincia);
         foglio.getRange(rigaReale, 6).setValue(dati.compagniaReparto); foglio.getRange(rigaReale, 7).setValue(dati.sedeUfficio);
         foglio.getRange(rigaReale, 8).setValue(dati.telefono); foglio.getRange(rigaReale, 9).setValue(dati.email);
     }
     return { success: true };
  } catch(e) { return { success: false, error: e.message }; }
}