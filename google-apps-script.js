// ============================================================
//  HR BURGER HOUSE — Google Apps Script (versión corregida)
//  El script crea su propia hoja de cálculo automáticamente
// ============================================================

function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    // Encabezados si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      const headers = [
        '#Orden','Fecha','Hora','Cajera',
        'Productos','Subtotal ($)','Descuento ($)','Total ($)','Método de Pago','Recibido ($)','Cambio ($)'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#e67e22')
        .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // Agregar la venta
    sheet.appendRow([
      data.orden,
      data.fecha,
      data.hora,
      data.cajera,
      data.productos,
      parseFloat(data.subtotal) || parseFloat(data.total),
      parseFloat(data.descuento) || 0,
      parseFloat(data.total),
      data.pago,
      parseFloat(data.recibido),
      parseFloat(data.cambio)
    ]);

    // Ajustar columnas al contenido
    sheet.autoResizeColumns(1, 11);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Obtiene la hoja de ventas — la crea la primera vez
function getSheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId    = props.getProperty('SS_ID');
  let ss;

  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch(e) { ss = null; }
  }

  if (!ss) {
    // Primera vez: crear una hoja nueva en Google Drive
    ss = SpreadsheetApp.create('FastPos · HR BURGER HOUSE · Ventas');
    props.setProperty('SS_ID', ss.getId());
    Logger.log('Hoja creada: ' + ss.getUrl());
  }

  // Usar la primera hoja (o crearla si no existe)
  let sheet = ss.getSheets()[0];
  if (!sheet) sheet = ss.insertSheet('Ventas');
  return sheet;
}

// Ejecuta esto manualmente UNA VEZ para ver la URL de tu hoja
function verURLdeHoja() {
  const props = PropertiesService.getScriptProperties();
  const ssId  = props.getProperty('SS_ID');
  if (ssId) {
    const url = 'https://docs.google.com/spreadsheets/d/' + ssId;
    Logger.log('Tu hoja de ventas: ' + url);
    SpreadsheetApp.openById(ssId); // abre la hoja
  } else {
    Logger.log('La hoja aún no fue creada. Envía una venta primero.');
  }
}

// Prueba manual — simula una venta
function testScript() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        orden: '001', fecha: '15/05/2026', hora: '14:30',
        cajera: 'María', productos: '2x Hamburguesa Doble | 1x Refresco',
        total: '158.00', pago: 'Efectivo', recibido: '200.00', cambio: '42.00'
      })
    }
  };
  doPost(fakeEvent);
  verURLdeHoja();
}


// ============================================================
//  INSTRUCCIONES PARA ACTUALIZAR EL SCRIPT
// ============================================================
//
//  1. Ve a script.google.com
//  2. Abre tu proyecto "FastPos HR Burger House"
//  3. Borra TODO el código actual
//  4. Pega el contenido de ESTE archivo
//  5. Guarda (Ctrl+S)
//  6. Haz clic en "Implementar" → "Administrar implementaciones"
//  7. En tu implementación activa → clic en el lápiz (editar)
//  8. En "Versión" selecciona "Nueva versión"
//  9. Clic en "Implementar"
//
//  La URL que ya tienes pegada en Ajustes NO cambia.
//
//  Para ver tu hoja de cálculo:
//  → En el editor de Apps Script, selecciona la función "verURLdeHoja"
//  → Haz clic en Ejecutar
//  → Ve a "Registros de ejecución" y copia el enlace que aparece
// ============================================================
