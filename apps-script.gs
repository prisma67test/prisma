function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    let data = {};
    try { data = JSON.parse(e.postData.contents); } catch(err) {}

    const email  = (data.email  || '').toLowerCase().trim();
    const nombre = (data.nombre || '').trim();

    // ---- Filtros anti-test ----
    const emailsFalsos = ['ejemplo.com', 'test.com', 'prueba.com', 'fake.com'];
    const esFalso = emailsFalsos.some(d => email.includes(d))
                 || email === ''
                 || nombre === ''
                 || nombre.toLowerCase().includes('test')
                 || nombre.toLowerCase().includes('prisma');

    if (esFalso) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'filtrado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    // --------------------------

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Fecha', 'Nombre', 'Email', 'Fuente']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      nombre,
      email,
      data.fuente || 'directo'
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'activo ✓' }))
    .setMimeType(ContentService.MimeType.JSON);
}
