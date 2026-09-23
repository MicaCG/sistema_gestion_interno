// =========================================================================
// GESTIÓN DE PEDIDOS - TOMAS OBERTI® TATTOO SUPPLIES
// BACKEND GOOGLE APPS SCRIPT (Code.gs) - VERSIÓN MVP UNIFICADA
// =========================================================================

const SPREADSHEET_ID = "1dr3z0GKSV4H1KIhjnc9jk_j7aMq5OQCP-PuYm-VzDs0";

function getSS() {
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch(e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// =========================================================================
// MAPEO INTELIGENTE DE COLUMNAS POR ENCABEZADO
// =========================================================================
function getColumnMap(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const map = {};
  headers.forEach((h, idx) => {
    if (h !== undefined && h !== null) {
      const cleanKey = String(h).trim().toLowerCase().replace(/[\s_-]/g, '');
      map[cleanKey] = idx + 1;
    }
  });
  return map;
}

function findCol(map, possibleNames, defaultCol) {
  for (let name of possibleNames) {
    const clean = name.toLowerCase().replace(/[\s_-]/g, '');
    if (map[clean] !== undefined) return map[clean];
  }
  return defaultCol;
}

// =========================================================================
// ENDPOINTS WEB APP
// =========================================================================
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify(obtenerDatosTablero()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput("API Tomas Oberti Activa");
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  
  try {
    let data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    
    const ss = getSS();
    const sheetActivos = ss.getSheetByName('Activos') || ss.getSheetByName('Pedidos Activos');
    const sheetHistorico = ss.getSheetByName('Historico') || ss.getSheetByName('Histórico');
    
    if (!sheetActivos) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Hoja Activos no encontrada' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    let res = { success: false };
    
    switch (data.action) {
      case 'crearPedido':
        res = crearPedido(sheetActivos, data.pedido);
        break;

      case 'iniciarPreparacion':
        res = iniciarPreparacion(sheetActivos, data.idVenta, data.armador);
        break;

      case 'finalizarPreparacion':
        res = finalizarPreparacion(sheetActivos, data.idVenta);
        break;

      case 'facturarArchivar':
        res = facturarArchivar(sheetActivos, sheetHistorico, data.idVenta, data.facturador);
        break;
        
      default:
        res = { success: false, error: 'Acción no reconocida' };
    }
    
    return ContentService.createTextOutput(JSON.stringify(res))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =========================================================================
// LÓGICA DE NEGOCIO (FORMATO COMPLETO UNIFICADO: DD-MM-YYYY (HH:mm hs))
// =========================================================================

function obtenerDatosTablero() {
  const ss = getSS();
  const sheet = ss.getSheetByName('Activos') || ss.getSheetByName('Pedidos Activos');
  if (!sheet) return { success: false, error: 'Hoja Activos no encontrada', pedidos: [] };
  
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { success: true, pedidos: [] };
  
  const headers = rows[0].map(h => String(h).trim());
  const pedidos = [];
  const idsVistos = new Set();

  for (let i = 1; i < rows.length; i++) {
    const idRaw = rows[i][0];
    if (!idRaw && String(idRaw).trim() === '') continue;
    const idStr = String(idRaw).trim();
    
    if (idsVistos.has(idStr)) continue;
    idsVistos.add(idStr);
    
    let obj = {};
    headers.forEach((h, idx) => {
      let v = rows[i][idx];
      if (v instanceof Date) {
        v = Utilities.formatDate(v, Session.getScriptTimeZone() || "GMT-3", "dd-MM-yyyy (HH:mm 'hs')");
      }
      obj[h] = v !== undefined && v !== null ? String(v).trim() : '';
    });
    pedidos.push(obj);
  }
  return { success: true, pedidos: pedidos };
}

function crearPedido(sheet, pedido) {
  if (!pedido || !pedido.idVenta) return { success: false };
  
  const colMap = getColumnMap(sheet);
  const colId = findCol(colMap, ['idventa', 'id'], 1);
  const colFecha = findCol(colMap, ['fechaingreso', 'fecha'], 2);
  const colPlat = findCol(colMap, ['plataforma'], 3);
  const colEstado = findCol(colMap, ['estado'], 4);
  const colObs = findCol(colMap, ['observaciones'], 10);
  const colEstadoImp = findCol(colMap, ['estadoimpresion'], 12);
  
  const data = sheet.getDataRange().getValues();
  const idNuevo = String(pedido.idVenta).trim();
  
  // Anti-duplicados en Sheets
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId - 1]).trim() === idNuevo) {
      if (colObs > 0 && pedido.observaciones) {
        sheet.getRange(i + 1, colObs).setValue(pedido.observaciones);
      }
      return { success: true, message: 'Ya existía' };
    }
  }

  const numCols = Math.max(sheet.getLastColumn(), 12);
  const fila = new Array(numCols).fill('');
  const tz = Session.getScriptTimeZone() || "GMT-3";
  const ahoraFechaHora = pedido.fechaIngreso || Utilities.formatDate(new Date(), tz, "dd-MM-yyyy (HH:mm 'hs')");

  fila[colId - 1] = idNuevo;
  fila[colFecha - 1] = ahoraFechaHora;
  if (colPlat > 0) fila[colPlat - 1] = pedido.plataforma || 'Mercado Libre';
  if (colEstado > 0) fila[colEstado - 1] = 'Nuevo';
  if (colObs > 0) fila[colObs - 1] = pedido.observaciones || '';
  if (colEstadoImp > 0) fila[colEstadoImp - 1] = 'No Impreso';

  sheet.appendRow(fila);
  return { success: true };
}

function iniciarPreparacion(sheet, idVenta, armador) {
  const data = sheet.getDataRange().getValues();
  const colMap = getColumnMap(sheet);
  
  const colId = findCol(colMap, ['idventa', 'id'], 1);
  const colEstado = findCol(colMap, ['estado'], 4);
  const colHoraInicio = findCol(colMap, ['horainicioprep', 'horainicio'], 5);
  const colArmador = findCol(colMap, ['armadordepo', 'armador'], 6);
  
  const tz = Session.getScriptTimeZone() || "GMT-3";
  // Mismo formato que la fecha de ingreso
  const ahoraFechaHora = Utilities.formatDate(new Date(), tz, "dd-MM-yyyy (HH:mm 'hs')");
  const idBuscado = String(idVenta).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId - 1]).trim() === idBuscado) {
      if (colEstado > 0) sheet.getRange(i + 1, colEstado).setValue('En Preparación');
      if (colHoraInicio > 0) sheet.getRange(i + 1, colHoraInicio).setValue(ahoraFechaHora);
      if (colArmador > 0) sheet.getRange(i + 1, colArmador).setValue(armador);
      return { success: true };
    }
  }
  return { success: false, error: 'Pedido #' + idVenta + ' no encontrado en Activos' };
}

function finalizarPreparacion(sheet, idVenta) {
  const data = sheet.getDataRange().getValues();
  const colMap = getColumnMap(sheet);
  
  const colId = findCol(colMap, ['idventa', 'id'], 1);
  const colEstado = findCol(colMap, ['estado'], 4);
  const colHoraFin = findCol(colMap, ['horafinprep', 'horafin'], 7);
  
  const tz = Session.getScriptTimeZone() || "GMT-3";
  // Mismo formato que la fecha de ingreso
  const ahoraFechaHora = Utilities.formatDate(new Date(), tz, "dd-MM-yyyy (HH:mm 'hs')");
  const idBuscado = String(idVenta).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId - 1]).trim() === idBuscado) {
      if (colEstado > 0) sheet.getRange(i + 1, colEstado).setValue('Para Facturar');
      if (colHoraFin > 0) sheet.getRange(i + 1, colHoraFin).setValue(ahoraFechaHora);
      return { success: true };
    }
  }
  return { success: false, error: 'Pedido #' + idVenta + ' no encontrado en Activos' };
}

function facturarArchivar(sheetActivos, sheetHistorico, idVenta, facturador) {
  if (!sheetHistorico) return { success: false, error: 'Hoja Historico no encontrada' };
  
  const data = sheetActivos.getDataRange().getValues();
  const colMap = getColumnMap(sheetActivos);
  
  const colId = findCol(colMap, ['idventa', 'id'], 1);
  const colEstado = findCol(colMap, ['estado'], 4);
  const colFacturador = findCol(colMap, ['facturadoradmin', 'facturador'], 8);
  const colHoraFact = findCol(colMap, ['horafacturacion', 'horafact'], 9);
  
  const tz = Session.getScriptTimeZone() || "GMT-3";
  // Mismo formato que la fecha de ingreso
  const ahoraFechaHora = Utilities.formatDate(new Date(), tz, "dd-MM-yyyy (HH:mm 'hs')");
  const idBuscado = String(idVenta).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][colId - 1]).trim() === idBuscado) {
      let fila = [...data[i]];
      if (colEstado > 0) fila[colEstado - 1] = 'Facturado';
      if (colFacturador > 0) fila[colFacturador - 1] = facturador;
      if (colHoraFact > 0) fila[colHoraFact - 1] = ahoraFechaHora;
      
      sheetHistorico.appendRow(fila);
      sheetActivos.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Pedido #' + idVenta + ' no encontrado en Activos' };
}
