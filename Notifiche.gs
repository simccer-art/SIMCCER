// ================= NOTIFICHE UNIFICATE ================= //

function getConteggioNotifiche(utenteLoggato) {
  try {
    let countsP = { mieInAttesa: 0, mieVistate: 0, mieApprovate: 0, mieRifiutate: 0, mieEvase: 0, mieFruite: 0, daVistare: 0, daApprovare: 0, daEvadere: 0, hashMie: "", hashGestione: "", modificheProfilo: 0 };
    let countsV = { mieInAttesa: 0, mieApprovate: 0, mieRifiutate: 0, mieEvase: 0, daApprovare: 0, daEvadere: 0, hashMie: "", hashGestione: "" };
    let countsR = { mieDaCorreggere: 0, mieInAttesa: 0, miePagate: 0, mieRifiutate: 0, daApprovare: 0, daPagare: 0, hashMie: "", hashGestione: "" };

    let liv = GERARCHIA[utenteLoggato.ruoloBase] || 1;
    let ruoloStr = String(utenteLoggato.ruolo || "").toUpperCase();
    
    let isTesoriere1 = ruoloStr.includes('TESORIERE') && !ruoloStr.includes('TESORIERE2');
    let isTesoriere2 = ruoloStr.includes('TESORIERE2');
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
           
           let targetProv = prov; 
           if (isProvinciale) {
               let match = sp.match(/\(([^)]+)\)/);
               if (match) targetProv = match[1].toUpperCase().trim();
           }
           
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
    
    let countsRubrica = { daApprovare: 0, hashGestione: "" };
    if (liv === 5) { 
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