/**
 * 🧩 Agrega un menú personalizado al abrir la hoja
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📤 Reportes')
    .addItem('Exportar hoja como PDF y enviar', 'exportarYEnviarPDF')
    .addToUi();
}

/**
 * 📄 exportar_pdf.gs
 * 
 * Exporta una hoja específica de Google Sheets como PDF y la envía automáticamente por correo.
 * Ideal para reportes periódicos, presupuestos u hojas de presentación.
 */
function exportarYEnviarPDF() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet(); // Podés cambiar esto si querés exportar una hoja fija
  const hojaId = sheet.getSheetId();
  const urlBase = `https://docs.google.com/spreadsheets/d/${spreadsheet.getId()}/export`;

  const nombrePDF = `Reporte - ${sheet.getName()} - ${new Date().toISOString().slice(0, 10)}`;
  const destinatario = "tucorreo@ejemplo.com"; // ← reemplazá por tu dirección real

  const opciones = {
    exportFormat: "pdf",
    format: "pdf",
    size: "A4",
    portrait: true,
    fitw: true,
    sheetnames: false,
    printtitle: false,
    pagenumbers: false,
    gridlines: false,
    fzr: false,
    gid: hojaId
  };

  const params = Object.keys(opciones).map(k => `${k}=${opciones[k]}`).join("&");
  const token = ScriptApp.getOAuthToken();
  const respuesta = UrlFetchApp.fetch(`${urlBase}?${params}`, {
    headers: {
      Authorization: `Bearer ${token}`
    }
  });

  GmailApp.sendEmail(destinatario, `📄 Envío automático: ${nombrePDF}`, 
    `Se adjunta el reporte generado automáticamente desde Google Sheets.`, {
    attachments: [{
      fileName: `${nombrePDF}.pdf`,
      content: respuesta.getBlob().getBytes(),
      mimeType: "application/pdf"
    }]
  });
}
