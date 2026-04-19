// ================= COSTANTI GLOBALI ================= //
const GESTIONALE_ID = '1tAJTVYkffUSYMLH2au2liKpQFITBcgCcm9647z9klws';
const FOLDER_REPORTS_ID = '1fhYQU98RAD0WzHq42dfFX6pZsXHkPp78'; 
const CONTABILITA_ID = '1TUG3r9PUuPvwK5xK6lpLmTBE44M8wZhK';
const FOLDER_CONTABILITA_ID = '1TUG3r9PUuPvwK5xK6lpLmTBE44M8wZhK'; // Unificata per sicurezza
const GERARCHIA = { "UTENTE": 1, "SEGRGEN": 2, "APPROVATORE": 3, "GESTORE": 4, "AMMINISTRATORE": 5 };
const RUBRICA_DB_ID = '1WCPU0FYp930U5qXthuq83dqlUqx6SF2ZCKRgNl3INzI';
const TEMPLATE_RIMBORSO_DOC_ID = '1ADoeXnFhtelnm9ZiPNL_1Xk02iXVRWjEiedFBW9H8s4';
const CARTELLA_LOGHI_ID = '1ewdYZ5F_o-dW5SvOOtNVF978Y9gi5yba';
const FOLDER_MATERIALI_SHEETS_ID = '1q-I43asRS2BUp58dJIb5gqwil7l-sy5Z'; 
const FOLDER_MATERIALI_DOCS_ID = '1yycGzRe2WeY3_iywhC5y1OTlAzJgSHU8';

// ================= FUNZIONI DI BASE ================= //
function doGet() { 
  return HtmlService.createTemplateFromFile('Index') // <--- Qui la magia: createTemplate al posto di createHtmlOutput
    .evaluate() // <--- Questo comando dice a Google di eseguire i tag <?!= ?>
    .setTitle('Gestionale SIM CC ER')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

// Funzione magica per unire i file HTML/CSS/JS
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}