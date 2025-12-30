/**
 * BACKEND: Code.gs
 * Versi: Duration Logic (Weeks/Months) + Payment Progress
 */

// function onOpen() {
//   SpreadsheetApp.getUi()
//       .createMenu('⚡ ADMIN TOOLS')
//       .addItem('Semak Sambungan Data', 'checkConnections')
//       .addToUi();
// }

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Dashboard Projek BPM LKIM (Glass UI)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// function checkConnections() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const sheets = ['syarikat', 'bahagian', 'seksyen_unit', 'Vot', 'Perakuan Siap']; 
//   let msg = "";
//   sheets.forEach(name => {
//     let s = ss.getSheetByName(name);
//     if (!s && name === 'Vot') s = ss.getSheetByName('vot');
//     if (!s && name === 'Perakuan Siap') s = ss.getSheetByName('Perakuan_Siap') || ss.getSheetByName('Perakuan Siap Kerja');
//     if (s) msg += `✅ Sheet '${name}' DIJUMPAI.\n`;
//     else msg += `❌ Sheet '${name}' TIDAK DIJUMPAI.\n`;
//   });
//   SpreadsheetApp.getUi().alert(msg);
// }