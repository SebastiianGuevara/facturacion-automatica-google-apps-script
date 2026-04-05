// Ids de las configuracion de las carpetas en drive

const ID_PLANTILLA        = CONFIG.ID_PLANTILLA;
const ID_CARPETA_PDF      = CONFIG.ID_CARPETA_PDF;
const ID_CARPETA_TEMPORAL = CONFIG.ID_CARPETA_TEMPORAL;

// Crea el menú personalizado al abrir el Sheets
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📄 Facturación')
    .addItem('✉️ Generar factura (fila seleccionada)', 'generarFacturaFilaActiva')
    .addItem('📦 Generar TODAS las facturas', 'generarTodasLasFacturas')
    .addSeparator()
    .addItem('📋 Ver registro de facturas', 'verRegistro')
    .addToUi();
}


// ── Genera factura solo para la fila donde está el cursor
function generarFacturaFilaActiva() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  const fila = hoja.getActiveCell().getRow();

  if (fila <= 1) {
    SpreadsheetApp.getUi().alert('⚠️ Selecciona una fila de datos, no el encabezado.');
    return;
  }

  const datos = obtenerDatosFila(hoja, fila);

  if (!datos.nombre_cliente || !datos.email_cliente) {
    SpreadsheetApp.getUi().alert('⚠️ La fila seleccionada no tiene nombre o email.');
    return;
  }

  procesarFactura(datos);
  SpreadsheetApp.getUi().alert(`✅ Factura enviada a: ${datos.email_cliente}`);
}


// ── Genera facturas para TODAS las filas de datos
function generarTodasLasFacturas() {
  const hoja       = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Clientes');
  const ultimaFila = hoja.getLastRow();

  if (ultimaFila < 2) {
    SpreadsheetApp.getUi().alert('❌ No hay datos en la hoja Clientes.');
    return;
  }

  let enviadas = 0;
  for (let fila = 2; fila <= ultimaFila; fila++) {
    const datos = obtenerDatosFila(hoja, fila);
    if (datos.nombre_cliente && datos.email_cliente) {
      procesarFactura(datos);
      enviadas++;
      Utilities.sleep(1500); // Pausa entre envíos para no superar límites de Gmail
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Proceso terminado. Facturas enviadas: ${enviadas}`);
}


// ── Extrae los datos de una fila del Sheets
function obtenerDatosFila(hoja, fila) {
  const numeroFactura = `FAC-${new Date().getFullYear()}-${String(fila).padStart(4, '0')}`;
  const fecha         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');

  return {
    numero_factura:       numeroFactura,
    fecha:                fecha,
    nombre_cliente:       hoja.getRange(fila, 1).getValue(),
    email_cliente:        hoja.getRange(fila, 2).getValue(),
    ciudad:               hoja.getRange(fila, 3).getValue(),
    descripcion_servicio: hoja.getRange(fila, 4).getValue(),
    cantidad:             hoja.getRange(fila, 5).getValue(),
    precio_unitario:      hoja.getRange(fila, 6).getValue(),
    total:                hoja.getRange(fila, 7).getValue(),
  };
}


// ── Crea el PDF y lo envía por correo
function procesarFactura(datos) {
  const carpetaTemp = DriveApp.getFolderById(ID_CARPETA_TEMPORAL);
  const carpetaPDF  = DriveApp.getFolderById(ID_CARPETA_PDF);
  const plantilla   = DriveApp.getFileById(ID_PLANTILLA);

  // 1. Copia la plantilla en la carpeta Temporal
  const copia = plantilla.makeCopy(`Factura_${datos.numero_factura}`, carpetaTemp);

  // 2. Reemplaza los placeholders con los datos reales
  const doc    = DocumentApp.openById(copia.getId());
  const cuerpo = doc.getBody();

  cuerpo.replaceText('{{numero_factura}}',       datos.numero_factura);
  cuerpo.replaceText('{{fecha}}',                datos.fecha);
  cuerpo.replaceText('{{nombre_cliente}}',       datos.nombre_cliente);
  cuerpo.replaceText('{{email_cliente}}',        datos.email_cliente);
  cuerpo.replaceText('{{ciudad}}',               datos.ciudad);
  cuerpo.replaceText('{{descripcion_servicio}}', datos.descripcion_servicio);
  cuerpo.replaceText('{{cantidad}}',             String(datos.cantidad));
  cuerpo.replaceText('{{precio_unitario}}',      String(datos.precio_unitario));
  cuerpo.replaceText('{{total}}',                String(datos.total));

  // 3. Guarda el Doc y lo convierte a PDF
  doc.saveAndClose();
  const blobPDF = copia.getAs(MimeType.PDF);

  // 4. Guarda el PDF en la carpeta Facturas-PDF
  carpetaPDF.createFile(blobPDF)
    .setName(`Factura_${datos.numero_factura}_${datos.nombre_cliente}.pdf`);

  // 5. Envía el correo con el PDF adjunto
  const asunto  = `Factura ${datos.numero_factura} - ${datos.nombre_cliente}`;
  const mensaje = `Estimado/a ${datos.nombre_cliente},\n\n`
                + `Adjunto encontrará su factura No. ${datos.numero_factura} `
                + `por un valor de $${datos.total}.\n\n`
                + `Gracias por su preferencia.\n\n`
                + `Saludos cordiales.`;

  GmailApp.sendEmail(datos.email_cliente, asunto, mensaje, {
    attachments: [blobPDF],
    name: 'Sistema de Facturación'
  });

  // 6. Borra la copia temporal para no acumular archivos
  copia.setTrashed(true);

  // 7. Registra el envío en la hoja Registro-Facturas
  registrarFactura(datos);
}


// ── Escribe una fila en el log de facturas
function registrarFactura(datos) {
  const hojaRegistro = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Registro-Facturas');

  hojaRegistro.appendRow([
    datos.numero_factura,
    datos.nombre_cliente,
    datos.email_cliente,
    new Date(),
    '✅ Enviada'
  ]);
}


// ── Navega a la hoja de registro
function verRegistro() {
  SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Registro-Facturas')
    .activate();
}
