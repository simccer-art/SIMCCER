const GESTIONALE_ID = '1tAJTVYkffUSYMLH2au2liKpQFITBcgCcm9647z9klws';
const FOLDER_REPORTS_ID = '1fhYQU98RAD0WzHq42dfFX6pZsXHkPp78'; 
const CONTABILITA_ID = '1TUG3r9PUuPvwK5xK6lpLmTBE44M8wZhK';
const GERARCHIA = { "UTENTE": 1, "SEGRGEN": 2, "APPROVATORE": 3, "GESTORE": 4, "AMMINISTRATORE": 5 };
const RUBRICA_DB_ID = '1WCPU0FYp930U5qXthuq83dqlUqx6SF2ZCKRgNl3INzI';
function doGet() { return HtmlService.createHtmlOutputFromFile('Index').setTitle('Gestionale SIM CC ER').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
// Funzione per includere file HTML (CSS e JS) in altri file HTML
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
