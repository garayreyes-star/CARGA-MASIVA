/***************************************************************
 * CARGA MASIVA AG — Lookup SKU → Lote/Stock + Subtotal (PRECIO MANUAL)
 * + Exportación CSV por DMI+ODV usando plantilla "PLANILLA EJEMPLO"
 *
 * ✅ LOTE 7550 ES LA PRIMERA OPCIÓN (PRIORIDAD ABSOLUTA).
 * ✅ Generación en formato .CSV (Mucho más rápido y nativo para ERP).
 * ✅ Conserva ceros a la izquierda en SKU y Lote (Texto puro).
 * ✅ Correlativo de líneas (LineNum) reinicia por DMI/ODV.
 * ✅ Formato de fecha DD-MM-YYYY.
 * ✅ Columna L: Registro de FECHA/HORA ("INGRESADO").
 * 🌐 WEBHOOK: Integrado con Flask. Busca DMI en hoja "DMI" y deja ODV en blanco.
 ***************************************************************/

const CM = {
  SHEET_LIST: 'LISTADO DE PRECIO',
  SHEET_ODV:  'CARGA MASIVA ODV',
  SHEET_TEMPLATE: 'PLANILLA EJEMPLO',
  SHEET_HISTORIAL: 'HISTORIAL CARGAS', 

  HEADER_ROW: 1,

  // ODV columns
  COL_DMI: 1,
  COL_ODV: 2,
  COL_SKU: 3,
  COL_LOTE: 4,
  COL_PRECIO: 5,
  COL_QTY: 6,
  COL_SUBTOTAL: 7,
  COL_STOCK_TOTAL: 8,
  COL_STOCK_LOTE: 9,
  COL_FALTAN: 10,
  COL_ESTADO: 11,
  COL_TIMESTAMP: 12, // NUEVA COLUMNA PARA FECHA Y HORA

  SUMMARY_START_COL: 14,  // N
  SUMMARY_TITLE_ROW: 1,
  SUMMARY_START_ROW: 2,
  SUMMARY_TABLE_HEADER_ROW: 6,

  LOTE_PREFERIDO: '7550',
  CSV_SEPARATOR: ';', // Separador estándar para CSV

  BULK_THRESHOLD_CELLS: 60,

  COLORS: {
    OK: '#C6EFCE',
    PARCIAL: '#FCE5CD',
    SIN_STOCK: '#F4CCCC',
    NO_EXISTE: '#F4CCCC',
    DESDOBLAR: '#EA9999',
    LOTE_PREFERIDO: '#D9D2E9'
  },

  EXPORT_FOLDER_NAME: 'EXPORT_CARGA_MASIVA'
};

/***************************************************************
 * 🌐 API WEBHOOK: ESCUCHA LOS PEDIDOS DESDE TU PÁGINA WEB (FLASK)
 ***************************************************************/
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(CM.SHEET_ODV);
    
    if (!sh) throw new Error("No existe la hoja " + CM.SHEET_ODV);

    const items = data.items || [];
    if (items.length === 0) {
       return ContentService.createTextOutput(JSON.stringify({ error: 'Sin items' }))
         .setMimeType(ContentService.MimeType.JSON);
    }

    let finalDMI = String(data.dmi).trim(); 
    const shDMI = ss.getSheetByName('DMI');
    
    if (shDMI) {
      const dmiData = shDMI.getDataRange().getDisplayValues();
      const rutBuscado = String(data.dmi).trim().toLowerCase(); 
      
      for (let i = 1; i < dmiData.length; i++) {
        let rutUsuario = String(dmiData[i][0]).trim().toLowerCase();
        let dmiAsignado = String(dmiData[i][1]).trim();
        if (rutUsuario === rutBuscado && dmiAsignado !== "") {
          finalDMI = dmiAsignado;
          break; 
        }
      }
    }

    const matrix = [];
    for (let i = 0; i < items.length; i++) {
      let sku = items[i].sku || '';
      let qty = items[i].cantidad || 0;
      let precio = items[i].precio || 0;
      matrix.push([finalDMI, "", "'" + sku, "", precio, qty]); 
    }

    let startRow = 2; 
    const maxRows = sh.getMaxRows();
    const skuValues = sh.getRange(1, CM.COL_SKU, maxRows, 1).getValues();
    
    for (let i = skuValues.length - 1; i >= 0; i--) {
      if (String(skuValues[i][0]).trim() !== '') {
        startRow = i + 2; 
        break;
      }
    }

    // Pegar los datos base
    sh.getRange(startRow, 1, matrix.length, 6).setValues(matrix);
    
    // Pegar la fecha y hora
    const timestampStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
    const tsArray = Array(matrix.length).fill([timestampStr]);
    sh.getRange(startRow, CM.COL_TIMESTAMP, matrix.length, 1).setValues(tsArray);

    CM_recalcularODV();

    return ContentService.createTextOutput(JSON.stringify({ status: 'Exito', lineas_procesadas: matrix.length }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'Error', mensaje: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/***************************************************************
 * MENU
 ***************************************************************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Carga Masiva AG')
    .addItem('0) ✅ Migrar/Agregar columna ODV (si falta)', 'CM_ensureODVColumn')
    .addSeparator()
    .addItem('1) Configurar ODV (formatos + headers + colores)', 'CM_setupODV')
    .addItem('2) Preparar LISTADO (SKU/LOTE como texto)', 'CM_prepararListadoTexto')
    .addSeparator()
    .addItem('🪄 Autocompletar ODV por DMI (Forzar manual)', 'CM_autocompletarODVporDMI')
    .addSeparator()
    .addItem('🔄 Recalcular toda la hoja ODV (Rápido)', 'CM_recalcularODV')
    .addItem('📌 Actualizar Resumen DMI (ODVs a generar)', 'CM_actualizarResumenDMI')
    .addItem('🎨 Reaplicar colores (formato condicional)', 'CM_aplicarFormatoODV')
    .addSeparator()
    .addItem('🧩 Buscar otro lote (fila seleccionada)', 'CM_buscarOtroLoteFilaSeleccionada')
    .addItem('🧩 Buscar otros lotes (todas las filas con faltantes/alertas)', 'CM_buscarOtrosLotesPendientes')
    .addSeparator()
    .addItem('📦 Generar CSV (por cada DMI + ODV)', 'CM_exportarCSV_porDMIyODV')
    .addItem('📄 Generar CSV ÚNICO (todas las líneas)', 'generarXLSXUnico')
    .addSeparator()
    .addItem('🧽 Limpiar ODV completa (A..L)', 'CM_limpiarODVCompleta')
    .addToUi();
}

/***************************************************************
 * 0) Asegura columna ODV
 ***************************************************************/
function CM_ensureODVColumn() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);

  const headers = sh.getRange(1, 1, 1, Math.max(12, sh.getLastColumn()))
    .getDisplayValues()[0]
    .map(x => String(x||'').trim().toUpperCase());

  if (headers.includes('ODV')) {
    ss.toast('✅ La columna ODV ya existe.', 'Carga Masiva AG', 5);
    return;
  }

  sh.insertColumnAfter(1);
  sh.getRange(1, 2).setValue('ODV').setFontWeight('bold');
  ss.toast('✅ Columna ODV creada en B.', 'Carga Masiva AG', 6);
}

/***************************************************************
 * SETUP ODV + LISTADO
 ***************************************************************/
function CM_setupODV() {
  CM_ensureODVColumn();

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);

  const rows = sh.getMaxRows() - 1;

  sh.getRange(2, CM.COL_DMI, rows, 1).setNumberFormat('@');
  sh.getRange(2, CM.COL_ODV, rows, 1).setNumberFormat('@');
  sh.getRange(2, CM.COL_SKU, rows, 1).setNumberFormat('@');
  sh.getRange(2, CM.COL_LOTE, rows, 1).setNumberFormat('@');
  sh.getRange(2, CM.COL_PRECIO, rows, 1).setNumberFormat('@');

  sh.getRange(2, CM.COL_QTY, rows, 1).setNumberFormat('0');
  sh.getRange(2, CM.COL_SUBTOTAL, rows, 1).setNumberFormat('0'); 
  sh.getRange(2, CM.COL_STOCK_TOTAL, rows, 1).setNumberFormat('0');
  sh.getRange(2, CM.COL_STOCK_LOTE, rows, 1).setNumberFormat('0');
  sh.getRange(2, CM.COL_FALTAN, rows, 1).setNumberFormat('0');
  sh.getRange(2, CM.COL_TIMESTAMP, rows, 1).setNumberFormat('@');

  const setIfEmpty = (col, name) => {
    const cur = String(sh.getRange(1, col).getValue() || '').trim();
    if (!cur) sh.getRange(1, col).setValue(name).setFontWeight('bold');
  };

  setIfEmpty(CM.COL_DMI, 'DMI');
  setIfEmpty(CM.COL_ODV, 'ODV');
  setIfEmpty(CM.COL_SKU, 'SKU');
  setIfEmpty(CM.COL_LOTE, 'LOTE');
  setIfEmpty(CM.COL_PRECIO, 'PRECIO');
  setIfEmpty(CM.COL_QTY, 'CANTIDADES');
  setIfEmpty(CM.COL_SUBTOTAL, 'SUBTOTAL');
  setIfEmpty(CM.COL_STOCK_TOTAL, 'STOCK TOTAL');
  setIfEmpty(CM.COL_STOCK_LOTE, 'STOCK LOTE');
  setIfEmpty(CM.COL_FALTAN, 'FALTAN');
  setIfEmpty(CM.COL_ESTADO, 'ESTADO');
  setIfEmpty(CM.COL_TIMESTAMP, 'INGRESADO');

  CM_aplicarFormatoODV();
  CM_actualizarResumenDMI();

  ss.toast('✅ ODV configurado (montos enteros sin formato).', 'Carga Masiva AG', 6);
}

function CM_prepararListadoTexto() {
  const ss = SpreadsheetApp.getActive();
  const shList = ss.getSheetByName(CM.SHEET_LIST);
  if (!shList) throw new Error(`No existe la hoja "${CM.SHEET_LIST}"`);

  const meta = getListMeta_(shList);
  shList.getRange(2, meta.colSKU, shList.getMaxRows() - 1, 1).setNumberFormat('@');
  shList.getRange(2, meta.colLote, shList.getMaxRows() - 1, 1).setNumberFormat('@');

  ss.toast('✅ LISTADO preparado: SKU/LOTE como texto.', 'Carga Masiva AG', 6);
}

/***************************************************************
 * AUTOCOMPLETAR ODV POR DMI (Optimizada para uso automático)
 ***************************************************************/
function CM_autocompletarODVporDMI(mostrarToast = true) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const rangeDMI = sh.getRange(2, CM.COL_DMI, lastRow - 1, 1);
  const rangeODV = sh.getRange(2, CM.COL_ODV, lastRow - 1, 1);
  
  const dmiVals = rangeDMI.getDisplayValues();
  const odvVals = rangeODV.getDisplayValues();

  const dmiToOdvMap = new Map();
  for (let i = 0; i < dmiVals.length; i++) {
    const dmi = String(dmiVals[i][0] || '').trim();
    const odv = String(odvVals[i][0] || '').trim();
    if (dmi && odv) {
      if (!dmiToOdvMap.has(dmi)) {
        dmiToOdvMap.set(dmi, odv); 
      }
    }
  }

  if (dmiToOdvMap.size === 0) {
    if (mostrarToast) ss.toast('No encontré ningún ODV. Ingresa al menos un ODV en algún DMI.', 'Carga Masiva AG', 5);
    return;
  }

  let cambios = 0;
  for (let i = 0; i < dmiVals.length; i++) {
    const dmi = String(dmiVals[i][0] || '').trim();
    const odvActual = String(odvVals[i][0] || '').trim();
    
    if (dmi && !odvActual && dmiToOdvMap.has(dmi)) {
      odvVals[i][0] = dmiToOdvMap.get(dmi);
      cambios++;
    }
  }

  if (cambios > 0) {
    rangeODV.setNumberFormat('@').setValues(odvVals);
    if (mostrarToast) ss.toast(`✅ ¡Listo! Se autocompletaron ${cambios} celdas de ODV.`, 'Carga Masiva AG', 6);
    CM_actualizarResumenDMI(); 
  } else {
    if (mostrarToast) ss.toast('ℹ️ No había celdas de ODV vacías que rellenar.', 'Carga Masiva AG', 5);
  }
}

/***************************************************************
 * onEdit (ODV) - CON MAGIA AUTOMÁTICA
 ***************************************************************/
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (!sh || sh.getName() !== CM.SHEET_ODV) return;

    const r0 = e.range.getRow();
    const c0 = e.range.getColumn();
    const nr = e.range.getNumRows();
    const nc = e.range.getNumColumns();
    if (r0 <= CM.HEADER_ROW) return;

    const touchesSKU = rangesIntersect_(r0, c0, nr, nc, 2, CM.COL_SKU, sh.getMaxRows(), 1);
    const touchesQTY = rangesIntersect_(r0, c0, nr, nc, 2, CM.COL_QTY, sh.getMaxRows(), 1);
    const touchesPrecio = rangesIntersect_(r0, c0, nr, nc, 2, CM.COL_PRECIO, sh.getMaxRows(), 1);
    const touchesDMI = rangesIntersect_(r0, c0, nr, nc, 2, CM.COL_DMI, sh.getMaxRows(), 1);
    const touchesODV = rangesIntersect_(r0, c0, nr, nc, 2, CM.COL_ODV, sh.getMaxRows(), 1);

    if (!touchesSKU && !touchesQTY && !touchesPrecio && !touchesDMI && !touchesODV) return;

    if (touchesODV || touchesDMI) {
      CM_autocompletarODVporDMI(false); 
    }

    if (touchesSKU) {
      const skuRange = sh.getRange(r0, CM.COL_SKU, nr, 1);
      skuRange.setNumberFormat('@');

      const disp = skuRange.getDisplayValues();
      for (let i = 0; i < disp.length; i++) {
        const s = String(disp[i][0] || '').trim();
        if (!s) continue;
        if (/E\+?\d+/i.test(s)) sh.getRange(r0 + i, CM.COL_SKU).setValue("'" + s);
      }
    }

    if (touchesSKU) {
      const skuVals = sh.getRange(r0, CM.COL_SKU, nr, 1).getDisplayValues();
      for (let i = 0; i < nr; i++) {
        const sku = String(skuVals[i][0] || '').trim();
        if (!sku) limpiarFilaCompleta_(sh, r0 + i);
      }
    }

    if ((nr * nc) >= CM.BULK_THRESHOLD_CELLS) {
      CM_recalcularODV();
      return;
    }

    const provider = getInventoryProviderBatch_();
    for (let i = 0; i < nr; i++) CM_updateRow_(sh, r0 + i, provider);

    CM_actualizarResumenDMI();
  } catch (err) {
    console.error(err);
  }
}

/***************************************************************
 * RECALC ODV MASIVO
 ***************************************************************/
function CM_recalcularODV() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);

  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return;

  const provider = getInventoryProviderBatch_();
  
  // Ahora leemos hasta la columna de FECHA/HORA
  const range = sh.getRange(2, 1, lastRow - 1, CM.COL_TIMESTAMP);
  const data = range.getDisplayValues();
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const skuRaw = String(row[CM.COL_SKU - 1] || '').trim();
    const qty = toInt_(row[CM.COL_QTY - 1]);
    const precioDisplay = String(row[CM.COL_PRECIO - 1] || '').trim();
    const precioCLP = parseCLPFromText_(precioDisplay);

    if (!skuRaw) {
      row[CM.COL_LOTE - 1] = '';
      row[CM.COL_SUBTOTAL - 1] = '';
      row[CM.COL_STOCK_TOTAL - 1] = '';
      row[CM.COL_STOCK_LOTE - 1] = '';
      row[CM.COL_FALTAN - 1] = '';
      row[CM.COL_ESTADO - 1] = '';
      row[CM.COL_TIMESTAMP - 1] = '';
      continue;
    }

    const skuKey = skuKey_(skuRaw);
    const lots = provider.getLots(skuKey);
    const subtotalCalc = (qty > 0 && precioCLP > 0) ? (precioCLP * qty) : '';

    if (!row[CM.COL_TIMESTAMP - 1]) {
      row[CM.COL_TIMESTAMP - 1] = nowStr;
    }

    if (!lots.length) {
      row[CM.COL_LOTE - 1] = '';
      row[CM.COL_SUBTOTAL - 1] = subtotalCalc;
      row[CM.COL_STOCK_TOTAL - 1] = 0;
      row[CM.COL_STOCK_LOTE - 1] = 0;
      row[CM.COL_FALTAN - 1] = qty > 0 ? qty : 0;
      
      let estado = 'SKU NO EXISTE';
      if (qty > 0 && precioCLP <= 0) estado += ' | Falta PRECIO';
      row[CM.COL_ESTADO - 1] = estado;
      continue;
    }

    const stockTotal = lots.reduce((a, b) => a + b.stock, 0);
    const chosen = chooseLot_(lots); 

    row[CM.COL_LOTE - 1] = String(chosen.lote);
    row[CM.COL_SUBTOTAL - 1] = subtotalCalc;
    row[CM.COL_STOCK_TOTAL - 1] = stockTotal;
    row[CM.COL_STOCK_LOTE - 1] = chosen.stock;

    const faltan = qty > 0 ? Math.max(0, qty - stockTotal) : 0;
    row[CM.COL_FALTAN - 1] = faltan;

    let estado = '';
    if (stockTotal <= 0) {
      estado = 'SIN STOCK';
    } else if (qty <= 0) {
      estado = `OK (stock ${stockTotal})`;
    } else if (qty > chosen.stock && stockTotal > chosen.stock) {
      estado = `⚠️ DESDOBLAR LOTE | Lote tiene ${chosen.stock}, piden ${qty}`;
    } else if (faltan === 0) {
      estado = 'OK';
    } else {
      estado = `PARCIAL (faltan ${faltan})`;
    }

    if (String(chosen.lote) === CM.LOTE_PREFERIDO) estado += ' | Usa lote 7550';
    if (qty > 0 && precioCLP <= 0) estado += ' | Falta PRECIO';

    row[CM.COL_ESTADO - 1] = estado;
  }

  range.setValues(data);
  sh.getRange(2, CM.COL_SUBTOTAL, lastRow - 1, 1).setNumberFormat('0');

  CM_actualizarResumenDMI();
}

/***************************************************************
 * UPDATE ROW
 ***************************************************************/
function CM_updateRow_(shODV, row, provider) {
  const skuRaw = String(shODV.getRange(row, CM.COL_SKU).getDisplayValue() || '').trim();
  const skuKey = skuKey_(skuRaw);
  const qty = toInt_(shODV.getRange(row, CM.COL_QTY).getValue());

  const precioDisplay = String(shODV.getRange(row, CM.COL_PRECIO).getDisplayValue() || '').trim();
  const precioCLP = parseCLPFromText_(precioDisplay);

  if (!skuRaw) {
    clearRowOutputs_(shODV, row);
    return;
  }

  const lots = provider.getLots(skuKey);

  const currentTs = shODV.getRange(row, CM.COL_TIMESTAMP).getValue();
  if (!currentTs) {
    shODV.getRange(row, CM.COL_TIMESTAMP).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss'));
  }

  if (!lots.length) {
    shODV.getRange(row, CM.COL_LOTE).setValue('');
    shODV.getRange(row, CM.COL_SUBTOTAL).setValue('');
    shODV.getRange(row, CM.COL_STOCK_TOTAL).setValue(0);
    shODV.getRange(row, CM.COL_STOCK_LOTE).setValue(0);
    shODV.getRange(row, CM.COL_FALTAN).setValue(qty > 0 ? qty : 0);
    shODV.getRange(row, CM.COL_ESTADO).setValue('SKU NO EXISTE');
    return;
  }

  const stockTotal = lots.reduce((a, b) => a + b.stock, 0);
  const chosen = chooseLot_(lots);

  shODV.getRange(row, CM.COL_LOTE).setValue(String(chosen.lote));

  const subtotal = (qty > 0 && precioCLP > 0) ? (precioCLP * qty) : '';
  shODV.getRange(row, CM.COL_SUBTOTAL).setValue(subtotal);
  shODV.getRange(row, CM.COL_SUBTOTAL).setNumberFormat('0'); 

  shODV.getRange(row, CM.COL_STOCK_TOTAL).setValue(stockTotal);
  shODV.getRange(row, CM.COL_STOCK_LOTE).setValue(chosen.stock);

  const faltan = qty > 0 ? Math.max(0, qty - stockTotal) : 0;
  shODV.getRange(row, CM.COL_FALTAN).setValue(faltan);

  let estado = '';
  if (stockTotal <= 0) {
    estado = 'SIN STOCK';
  } else if (qty <= 0) {
    estado = `OK (stock ${stockTotal})`;
  } else if (qty > chosen.stock && stockTotal > chosen.stock) {
    estado = `⚠️ DESDOBLAR LOTE | Lote tiene ${chosen.stock}, piden ${qty}`;
  } else if (faltan === 0) {
    estado = 'OK';
  } else {
    estado = `PARCIAL (faltan ${faltan})`;
  }

  if (String(chosen.lote) === CM.LOTE_PREFERIDO) estado += ' | Usa lote 7550';
  if (qty > 0 && precioCLP <= 0) estado += ' | Falta PRECIO';

  shODV.getRange(row, CM.COL_ESTADO).setValue(estado);
}

/***************************************************************
 * 📊 NUEVO: FUNCIÓN PARA REGISTRAR EN EL DASHBOARD
 ***************************************************************/
function CM_registrarHistorial(tipoExportacion, lineas) {
  if (!lineas || lineas.length === 0) return;

  const ss = SpreadsheetApp.getActive();
  let shHist = ss.getSheetByName(CM.SHEET_HISTORIAL);

  if (!shHist) {
    shHist = ss.insertSheet(CM.SHEET_HISTORIAL);
    shHist.getRange("A1:H1").setValues([['FECHA Y HORA', 'TIPO EXPORTACIÓN', 'DMI', 'ODV', 'TOTAL LÍNEAS', 'CANTIDAD TOTAL', 'MONTO TOTAL', 'SKU PRINCIPAL']])
      .setFontWeight('bold')
      .setBackground('#111827')
      .setFontColor('#FFFFFF');
    shHist.setFrozenRows(1);
    shHist.setColumnWidth(1, 150);
    shHist.setColumnWidth(2, 130);
    shHist.setColumnWidth(8, 150);
  }

  const map = new Map();
  for (const r of lineas) {
    const key = `${r.dmi}||${r.odv}`;
    if (!map.has(key)) {
      map.set(key, { dmi: r.dmi, odv: r.odv, lineas: 0, qty: 0, monto: 0, skus: {} });
    }
    const o = map.get(key);
    o.lineas++;
    
    const qty = toInt_(r.qtyTxt);
    o.qty += qty;
    o.monto += parseCLPFromText_(r.subtotalTxt);

    if (!o.skus[r.sku]) o.skus[r.sku] = 0;
    o.skus[r.sku] += qty;
  }

  const rowsToAppend = [];
  const fechaActual = new Date();

  for (const o of map.values()) {
    let topSku = '';
    let maxQty = -1;
    for (const [sku, sq] of Object.entries(o.skus)) {
      if (sq > maxQty) { maxQty = sq; topSku = sku; }
    }

    rowsToAppend.push([
      fechaActual,
      tipoExportacion,
      o.dmi,
      o.odv,
      o.lineas,
      o.qty,
      o.monto,
      topSku
    ]);
  }

  if (rowsToAppend.length > 0) {
    const lr = shHist.getLastRow();
    const range = shHist.getRange(lr + 1, 1, rowsToAppend.length, 8);
    range.setValues(rowsToAppend);
    
    shHist.getRange(lr + 1, 1, rowsToAppend.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    shHist.getRange(lr + 1, 5, rowsToAppend.length, 3).setNumberFormat('0'); 
  }
}


// =========================================================================
// 🚀 PUENTES (WRAPPERS) PARA QUE LOS BOTONES DIBUJADOS NO SE ROMPAN
// =========================================================================
function generarXLSXUnico() {
  CM_exportarCSV_unico();
}
function CM_exportarXLSX_porDMIyODV() {
  CM_exportarCSV_porDMIyODV();
}
// =========================================================================


/***************************************************************
 * EXPORT CSV POR DMI + ODV 
 ***************************************************************/
function CM_exportarCSV_porDMIyODV() {
  const ss = SpreadsheetApp.getActive();
  const shODV = ss.getSheetByName(CM.SHEET_ODV);
  const shTpl = ss.getSheetByName(CM.SHEET_TEMPLATE);

  if (!shODV) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);
  if (!shTpl) throw new Error(`No existe la hoja plantilla "${CM.SHEET_TEMPLATE}"`);

  const lastRow = shODV.getLastRow();
  if (lastRow < 2) {
    ss.toast('No hay líneas en CARGA MASIVA ODV.', 'Export', 6);
    return;
  }

  const data = shODV.getRange(2, 1, lastRow - 1, CM.COL_ESTADO).getDisplayValues();

  const groups = new Map();
  const todasLasLineas = []; 

  for (const r of data) {
    const dmi = String(r[CM.COL_DMI - 1] || '').trim();
    const odv = String(r[CM.COL_ODV - 1] || '').trim();
    const sku = String(r[CM.COL_SKU - 1] || '').trim();
    if (!dmi || !odv || !sku) continue;

    const lote = String(r[CM.COL_LOTE - 1] || '').trim();
    const precioTxt = String(r[CM.COL_PRECIO - 1] || '').trim();
    const qtyTxt = String(r[CM.COL_QTY - 1] || '').trim();
    const subtotalTxt = String(r[CM.COL_SUBTOTAL - 1] || '').trim();

    const lineObj = { dmi, odv, sku, lote, precioTxt, qtyTxt, subtotalTxt };
    todasLasLineas.push(lineObj);

    const key = `${dmi}||${odv}`;
    if (!groups.has(key)) groups.set(key, { dmi, odv, lines: [] });
    groups.get(key).lines.push(lineObj);
  }

  if (!groups.size) {
    ss.toast('No encontré líneas completas (DMI+ODV+SKU).', 'Export', 8);
    return;
  }

  const folder = getOrCreateFolder_(CM.EXPORT_FOLDER_NAME);
  
  // Lee las columnas de la plantilla dinámicamente
  const tplLastCol = shTpl.getLastColumn();
  const tplHeaders = shTpl.getRange(1, 1, 1, tplLastCol).getDisplayValues()[0];
  const tplDefaults = shTpl.getRange(2, 1, 1, tplLastCol).getDisplayValues()[0];

  const tplMap = {};
  tplHeaders.forEach((h, i) => { 
    const k = norm_(h); 
    if (k) tplMap[k] = i; 
  });

  const idxCust = tplMap[norm_('CustAccount')];
  const idxSalesId = tplMap[norm_('SalesId')];
  const idxShip = tplMap[norm_('ShippingDateRequested')];
  const idxLine = tplMap[norm_('LineNum')];
  const idxBatch = tplMap[norm_('inventBatchId')];
  const idxItem = tplMap[norm_('ItemId')];
  const idxPrice = tplMap[norm_('SalesPrice')];
  const idxQty = tplMap[norm_('SalesQty')];
  const idxAmount = tplMap[norm_('LineAmount')];

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
  const links = [];

  for (const g of groups.values()) {
    const outName = `CARGA_${sanitize_(g.dmi)}_${sanitize_(g.odv)}_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm')}.csv`;
    
    // Iniciar las filas del CSV con el encabezado
    const csvRows = [tplHeaders];
    
    for (let i = 0; i < g.lines.length; i++) {
      const l = g.lines[i];
      const newRow = [...tplDefaults];
      
      if (idxCust !== undefined) newRow[idxCust] = l.dmi;
      if (idxSalesId !== undefined) newRow[idxSalesId] = l.odv;
      if (idxShip !== undefined) newRow[idxShip] = today;
      if (idxLine !== undefined) newRow[idxLine] = i + 1; // Correlativo individual
      if (idxBatch !== undefined) newRow[idxBatch] = l.lote;
      if (idxItem !== undefined) newRow[idxItem] = l.sku;
      if (idxPrice !== undefined) newRow[idxPrice] = parseCLPFromText_(l.precioTxt);
      if (idxQty !== undefined) newRow[idxQty] = toInt_(l.qtyTxt);
      if (idxAmount !== undefined) newRow[idxAmount] = parseCLPFromText_(l.subtotalTxt);
      
      csvRows.push(newRow);
    }

    const csvContent = arrayToCsv_(csvRows, CM.CSV_SEPARATOR);
    const blob = Utilities.newBlob(csvContent, MimeType.CSV, outName);
    folder.createFile(blob);
    links.push(outName);
  }

  CM_registrarHistorial('DMI INDIVIDUALES (CSV)', todasLasLineas);

  ss.toast(`✅ CSVs generados exitosamente: ${links.length}`, 'Export', 10);
}

/***************************************************************
 * EXPORT CSV ÚNICO
 ***************************************************************/
function CM_exportarCSV_unico() {
  const ss = SpreadsheetApp.getActive();
  const shODV = ss.getSheetByName(CM.SHEET_ODV);
  const shTpl = ss.getSheetByName(CM.SHEET_TEMPLATE);

  if (!shODV) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);
  if (!shTpl) throw new Error(`No existe la hoja plantilla "${CM.SHEET_TEMPLATE}"`);

  const lastRow = shODV.getLastRow();
  if (lastRow < 2) {
    ss.toast('No hay líneas en CARGA MASIVA ODV.', 'Export', 6);
    return;
  }

  const data = shODV.getRange(2, 1, lastRow - 1, CM.COL_ESTADO).getDisplayValues();
  const lines = [];
  
  for (const r of data) {
    const dmi = String(r[CM.COL_DMI - 1] || '').trim();
    const odv = String(r[CM.COL_ODV - 1] || '').trim();
    const sku = String(r[CM.COL_SKU - 1] || '').trim();
    if (!dmi || !odv || !sku) continue;

    lines.push({
      dmi,
      odv,
      sku,
      lote: String(r[CM.COL_LOTE - 1] || '').trim(),
      precioTxt: String(r[CM.COL_PRECIO - 1] || '').trim(),
      qtyTxt: String(r[CM.COL_QTY - 1] || '').trim(),
      subtotalTxt: String(r[CM.COL_SUBTOTAL - 1] || '').trim(),
    });
  }

  if (!lines.length) {
    ss.toast('No encontré líneas completas (DMI+ODV+SKU).', 'Export', 8);
    return;
  }

  // Ordenar para agrupar por DMI y ODV
  lines.sort((a,b) => a.dmi.localeCompare(b.dmi) || a.odv.localeCompare(b.odv) || a.sku.localeCompare(b.sku));

  const folder = getOrCreateFolder_(CM.EXPORT_FOLDER_NAME);
  const outName = `CARGA_UNICA_${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm')}.csv`;

  const tplLastCol = shTpl.getLastColumn();
  const tplHeaders = shTpl.getRange(1, 1, 1, tplLastCol).getDisplayValues()[0];
  const tplDefaults = shTpl.getRange(2, 1, 1, tplLastCol).getDisplayValues()[0];

  const tplMap = {};
  tplHeaders.forEach((h, i) => { 
    const k = norm_(h); 
    if (k) tplMap[k] = i; 
  });

  const idxCust = tplMap[norm_('CustAccount')];
  const idxSalesId = tplMap[norm_('SalesId')];
  const idxShip = tplMap[norm_('ShippingDateRequested')];
  const idxLine = tplMap[norm_('LineNum')];
  const idxBatch = tplMap[norm_('inventBatchId')];
  const idxItem = tplMap[norm_('ItemId')];
  const idxPrice = tplMap[norm_('SalesPrice')];
  const idxQty = tplMap[norm_('SalesQty')];
  const idxAmount = tplMap[norm_('LineAmount')];

  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');

  const csvRows = [tplHeaders];
  let currentDMI = '';
  let currentODV = '';
  let currentLineNum = 0;

  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    
    // Correlativo inteligente que se reinicia por ODV
    if (l.dmi !== currentDMI || l.odv !== currentODV) {
      currentDMI = l.dmi;
      currentODV = l.odv;
      currentLineNum = 1;
    } else {
      currentLineNum++;
    }

    const newRow = [...tplDefaults];
    
    if (idxCust !== undefined) newRow[idxCust] = l.dmi;
    if (idxSalesId !== undefined) newRow[idxSalesId] = l.odv;
    if (idxShip !== undefined) newRow[idxShip] = today;
    if (idxLine !== undefined) newRow[idxLine] = currentLineNum;
    if (idxBatch !== undefined) newRow[idxBatch] = l.lote;
    if (idxItem !== undefined) newRow[idxItem] = l.sku;
    if (idxPrice !== undefined) newRow[idxPrice] = parseCLPFromText_(l.precioTxt);
    if (idxQty !== undefined) newRow[idxQty] = toInt_(l.qtyTxt);
    if (idxAmount !== undefined) newRow[idxAmount] = parseCLPFromText_(l.subtotalTxt);

    csvRows.push(newRow);
  }

  const csvContent = arrayToCsv_(csvRows, CM.CSV_SEPARATOR);
  const blob = Utilities.newBlob(csvContent, MimeType.CSV, outName);
  const file = folder.createFile(blob);

  CM_registrarHistorial('ARCHIVO ÚNICO CSV', lines);

  ss.toast(`✅ CSV ÚNICO creado en Drive`, 'Export', 10);
  SpreadsheetApp.getUi().alert('CSV ÚNICO creado en Drive:\n\n' + file.getUrl());
}

/***************************************************************
 * BOTONES: Buscar otros lotes (DESDOBLAR OPTIMIZADO)
 ***************************************************************/
function CM_buscarOtroLoteFilaSeleccionada() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);

  const row = sh.getActiveRange().getRow();
  if (row <= 1) {
    ss.toast('Selecciona una fila de datos (>=2).', 'Carga Masiva AG', 5);
    return;
  }

  const provider = getInventoryProviderBatch_();
  const ok = desdoblarSiFaltaStock_(sh, row, provider);

  CM_actualizarResumenDMI();
  ss.toast(ok ? '✅ Lotes desdoblados correctamente.' : 'ℹ️ No aplica desdoblar (ya está OK o falta stock real)', 'Carga Masiva AG', 6);
}

function CM_buscarOtrosLotesPendientes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const provider = getInventoryProviderBatch_();
  const data = sh.getRange(2, CM.COL_ESTADO, lastRow - 1, 1).getValues();

  let changed = 0;
  for (let i = data.length - 1; i >= 0; i--) {
    const estado = String(data[i][0]).toUpperCase();
    if (estado.includes('DESDOBLAR') || estado.includes('PARCIAL')) {
      if (desdoblarSiFaltaStock_(sh, i + 2, provider)) changed++;
    }
  }

  CM_actualizarResumenDMI();
  ss.toast(`✅ Listo. Filas desdobladas: ${changed}`, 'Carga Masiva AG', 6);
}

function desdoblarSiFaltaStock_(shODV, row, provider) {
  const rowRange = shODV.getRange(row, 1, 1, CM.COL_TIMESTAMP);
  const rowData = rowRange.getDisplayValues()[0];

  const dmi = rowData[CM.COL_DMI - 1];
  const odv = rowData[CM.COL_ODV - 1];
  const skuRaw = String(rowData[CM.COL_SKU - 1]).trim();
  const loteActual = String(rowData[CM.COL_LOTE - 1]).trim();
  const precioDisplay = String(rowData[CM.COL_PRECIO - 1]).trim();
  const timestampOriginal = String(rowData[CM.COL_TIMESTAMP - 1] || '').trim();
  let qtyTotal = toInt_(rowData[CM.COL_QTY - 1]);

  if (!skuRaw || qtyTotal <= 0) return false;
  
  const skuKey = skuKey_(skuRaw);
  const precioCLP = parseCLPFromText_(precioDisplay);

  const lotsAll = provider.getLots(skuKey);
  if (!lotsAll.length) return false;

  const lots = lotsAll.filter(x => x.stock > 0);
  if (!lots.length) return false;

  const stockActual = toInt_(rowData[CM.COL_STOCK_LOTE - 1]);
  const stockTotal = lots.reduce((s, x) => s + x.stock, 0);

  if (stockActual >= qtyTotal) return false;
  if (stockTotal <= stockActual) return false;

  lots.sort((a,b)=>{
    const aPref = String(a.lote) === CM.LOTE_PREFERIDO ? 1 : 0;
    const bPref = String(b.lote) === CM.LOTE_PREFERIDO ? 1 : 0;
    if (aPref !== bPref) return bPref - aPref; 
    if (b.stock !== a.stock) return b.stock - a.stock;
    return String(a.lote).localeCompare(String(b.lote));
  });

  const idxActual = lots.findIndex(x => String(x.lote) === loteActual);
  if (idxActual > 0) {
    const [x] = lots.splice(idxActual, 1);
    lots.unshift(x);
  }

  const allocations = [];
  let remain = qtyTotal;

  for (const l of lots) {
    if (remain <= 0) break;
    const take = Math.min(remain, l.stock);
    if (take <= 0) continue;
    allocations.push({ lote: l.lote, stock: l.stock, qty: take });
    remain -= take;
  }

  if (allocations.length <= 1) {
    CM_updateRow_(shODV, row, provider);
    return false;
  }

  const extra = allocations.length - 1;
  shODV.insertRowsAfter(row, extra);

  const outputData = [];

  allocations.forEach((a) => {
    const sub = (precioCLP > 0) ? (precioCLP * a.qty) : '';
    let estado = 'OK (multi-lote)';
    if (String(a.lote) === CM.LOTE_PREFERIDO) estado += ' | Usa lote 7550'; 
    if (precioCLP <= 0) estado += ' | Falta PRECIO';

    const newRow = new Array(CM.COL_TIMESTAMP).fill('');
    newRow[CM.COL_DMI - 1] = dmi;
    newRow[CM.COL_ODV - 1] = odv;
    newRow[CM.COL_SKU - 1] = "'" + skuRaw; 
    newRow[CM.COL_LOTE - 1] = String(a.lote);
    newRow[CM.COL_PRECIO - 1] = precioDisplay;
    newRow[CM.COL_QTY - 1] = a.qty;
    newRow[CM.COL_SUBTOTAL - 1] = sub;
    newRow[CM.COL_STOCK_TOTAL - 1] = stockTotal;
    newRow[CM.COL_STOCK_LOTE - 1] = a.stock;
    newRow[CM.COL_FALTAN - 1] = 0;
    newRow[CM.COL_ESTADO - 1] = estado;
    newRow[CM.COL_TIMESTAMP - 1] = timestampOriginal; 
    outputData.push(newRow);
  });

  if (remain > 0) {
    const lastR = outputData.length - 1;
    outputData[lastR][CM.COL_FALTAN - 1] = remain;
    outputData[lastR][CM.COL_ESTADO - 1] = `PARCIAL (faltan ${remain})` + (precioCLP <= 0 ? ' | Falta PRECIO' : '');
  }

  shODV.getRange(row, 1, allocations.length, CM.COL_TIMESTAMP).setValues(outputData);
  shODV.getRange(row, CM.COL_SKU, allocations.length, 1).setNumberFormat('@');
  shODV.getRange(row, CM.COL_PRECIO, allocations.length, 1).setNumberFormat('@');
  shODV.getRange(row, CM.COL_SUBTOTAL, allocations.length, 1).setNumberFormat('0'); 

  return true;
}

/***************************************************************
 * LIMPIEZA
 ***************************************************************/
function CM_limpiarODVCompleta() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  sh.getRange(2, 1, lastRow - 1, CM.COL_TIMESTAMP).clearContent();

  CM_actualizarResumenDMI();
  ss.toast('🧽 ODV limpiada completa (A..L).', 'Carga Masiva AG', 6);
}

function limpiarFilaCompleta_(shODV, row) {
  shODV.getRange(row, 1, 1, CM.COL_TIMESTAMP).clearContent();
}

function clearRowOutputs_(shODV, row) {
  shODV.getRange(row, CM.COL_LOTE).clearContent();
  shODV.getRange(row, CM.COL_SUBTOTAL).clearContent();
  shODV.getRange(row, CM.COL_STOCK_TOTAL, 1, 4).clearContent();
}

/***************************************************************
 * RESUMEN DMI (ODVs a generar) 
 ***************************************************************/
function CM_actualizarResumenDMI() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) return;

  const lastRow = sh.getLastRow();
  const startCol = CM.SUMMARY_START_COL;

  // ✅ FIX: Limpiamos agresivamente el panel derecho para matar cualquier fantasma
  sh.getRange(1, 13, sh.getMaxRows(), 10).clearContent().clearFormat();

  sh.getRange(CM.SUMMARY_TITLE_ROW, startCol, 1, 7)
    .merge()
    .setValue('RESUMEN DMI/ODV (ODVs a generar)')
    .setFontWeight('bold')
    .setFontSize(12)
    .setBackground('#1F2937')
    .setFontColor('#FFFFFF');

  if (lastRow < 2) {
    sh.getRange(CM.SUMMARY_START_ROW, startCol).setValue('Sin datos');
    return;
  }

  const data = sh.getRange(2, 1, lastRow - 1, CM.COL_ESTADO).getDisplayValues();

  const map = new Map();
  let totalLineas = 0, totalQty = 0, totalSubtotal = 0, totalFaltan = 0;

  for (const row of data) {
    const dmi = String(row[CM.COL_DMI - 1] || '').trim();
    const odv = String(row[CM.COL_ODV - 1] || '').trim();
    const sku = String(row[CM.COL_SKU - 1] || '').trim();
    if (!dmi && !odv && !sku) continue;

    totalLineas++;

    const qty = toInt_(row[CM.COL_QTY - 1]);
    const subtotal = parseCLPFromText_(row[CM.COL_SUBTOTAL - 1]);
    const faltan = toInt_(row[CM.COL_FALTAN - 1]);

    totalQty += qty;
    totalSubtotal += subtotal;
    totalFaltan += faltan;

    const key = `${dmi || '(SIN DMI)'}||${odv || '(SIN ODV)'}`;
    if (!map.has(key)) map.set(key, { dmi: dmi || '(SIN DMI)', odv: odv || '(SIN ODV)', lineas:0, qty:0, subtotal:0, faltan:0 });
    const o = map.get(key);
    o.lineas++;
    o.qty += qty;
    o.subtotal += subtotal;
    o.faltan += faltan;
  }

  const kpi = [
    ['ODVs a generar (DMI+ODV)', map.size],
    ['Líneas totales', totalLineas],
    ['Cantidades totales', totalQty],
    ['Subtotal total (según precio ingresado)', totalSubtotal],
    ['Faltantes totales', totalFaltan],
  ];

  sh.getRange(CM.SUMMARY_START_ROW, startCol, kpi.length, 2).setValues(kpi).setFontWeight('bold');
  sh.getRange(CM.SUMMARY_START_ROW + 3, startCol + 1).setNumberFormat('0');

  const headerRow = CM.SUMMARY_TABLE_HEADER_ROW;
  sh.getRange(headerRow, startCol, 1, 6)
    .setValues([['DMI','ODV','Líneas','Total Cant','Subtotal','Faltan']])
    .setFontWeight('bold')
    .setBackground('#111827')
    .setFontColor('#FFFFFF');

  const rows = Array.from(map.values()).map(o => [o.dmi, o.odv, o.lineas, o.qty, o.subtotal, o.faltan]);
  rows.sort((a,b) => (b[5]-a[5]) || (b[4]-a[4]));

  if (rows.length) {
    sh.getRange(headerRow+1, startCol, rows.length, 6).setValues(rows);
    sh.getRange(headerRow+1, startCol+4, rows.length, 1).setNumberFormat('0');
  }

  sh.setColumnWidth(startCol, 160);
  sh.setColumnWidth(startCol+1, 120);
  sh.setColumnWidth(startCol+2, 80);
  sh.setColumnWidth(startCol+3, 110);
  sh.setColumnWidth(startCol+4, 140);
  sh.setColumnWidth(startCol+5, 80);
}

/***************************************************************
 * FORMATO CONDICIONAL ODV
 ***************************************************************/
function CM_aplicarFormatoODV() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CM.SHEET_ODV);
  if (!sh) throw new Error(`No existe la hoja "${CM.SHEET_ODV}"`);

  const maxCols = CM.COL_ESTADO;
  const rngRows = sh.getRange(2, 1, sh.getMaxRows() - 1, maxCols);

  const keep = sh.getConditionalFormatRules().filter(rule =>
    !rule.getRanges().some(r => r.getA1Notation() === rngRows.getA1Notation())
  );

  const A_estado = colToA1_(CM.COL_ESTADO);
  const A_lote   = colToA1_(CM.COL_LOTE);

  const rules = [];

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=REGEXMATCH(UPPER($${A_estado}2),"DESDOBLAR")`)
    .setBackground(CM.COLORS.DESDOBLAR)
    .setRanges([rngRows])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=REGEXMATCH(UPPER($${A_estado}2),"^OK")`)
    .setBackground(CM.COLORS.OK)
    .setRanges([rngRows])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=REGEXMATCH(UPPER($${A_estado}2),"PARCIAL")`)
    .setBackground(CM.COLORS.PARCIAL)
    .setRanges([rngRows])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR(REGEXMATCH(UPPER($${A_estado}2),"SIN STOCK"),REGEXMATCH(UPPER($${A_estado}2),"NO EXISTE"))`)
    .setBackground(CM.COLORS.SIN_STOCK)
    .setRanges([rngRows])
    .build());

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=$${A_lote}2="${CM.LOTE_PREFERIDO}"`)
    .setBackground(CM.COLORS.LOTE_PREFERIDO)
    .setRanges([rngRows])
    .build());

  sh.setConditionalFormatRules(rules.concat(keep));
}

/***************************************************************
 * PROVIDER (LISTADO)
 ***************************************************************/
function getInventoryProviderBatch_() {
  const ss = SpreadsheetApp.getActive();
  const shList = ss.getSheetByName(CM.SHEET_LIST);
  if (!shList) throw new Error(`No existe la hoja "${CM.SHEET_LIST}"`);

  const meta = getListMeta_(shList);
  const map = buildInventoryMap_(shList, meta);
  return { getLots: (skuKey) => map.get(skuKey) || [] };
}

function getListMeta_(shList) {
  const lc = shList.getLastColumn();
  const headers = shList.getRange(1, 1, 1, lc).getDisplayValues()[0];
  const h = {};
  headers.forEach((v, i) => { const k = norm_(v); if (k) h[k] = i + 1; });

  const colSKU   = findCol_(h, ['CODIGO DE ART', 'CÓDIGO DE ART', 'CODIGO', 'SKU', 'COD ART']);
  const colLote  = findCol_(h, ['NUMERO DE LOTE', 'NÚMERO DE LOTE', 'LOTE']);
  const colStock = findCol_(h, ['FISICA DISPONIBLE', 'FÍSICA DISPONIBLE', 'DISPONIBLE', 'STOCK', 'CANT DISP']);

  if (!colSKU || !colLote || !colStock) {
    throw new Error(`No pude identificar columnas en "${CM.SHEET_LIST}". Necesito SKU, LOTE, STOCK.`);
  }
  return { lc, colSKU, colLote, colStock };
}

function findCol_(hmap, candidates) {
  for (const c of candidates) {
    const k = norm_(c);
    if (hmap[k]) return hmap[k];
  }
  const keys = Object.keys(hmap);
  for (const c of candidates) {
    const ck = norm_(c);
    const found = keys.find(k => k.includes(ck));
    if (found) return hmap[found];
  }
  return 0;
}

function buildInventoryMap_(shList, meta) {
  const lr = effectiveLastRowByColDisplay_(shList, meta.colSKU);
  const map = new Map();
  if (lr < 2) return map;

  const skuDisp  = shList.getRange(2, meta.colSKU, lr - 1, 1).getDisplayValues().flat();
  const loteDisp = shList.getRange(2, meta.colLote, lr - 1, 1).getDisplayValues().flat();
  const stockVal = shList.getRange(2, meta.colStock, lr - 1, 1).getValues().flat();

  for (let i = 0; i < skuDisp.length; i++) {
    const key = skuKey_(skuDisp[i]);
    if (!key) continue;

    const lote = String(loteDisp[i] || '').trim();
    const stock = toInt_(stockVal[i]);
    if (!lote) continue;

    if (!map.has(key)) map.set(key, []);
    map.get(key).push({ lote, stock });
  }
  return map;
}

function chooseLot_(lots) {
  const filtered = lots.filter(x => x.stock > 0);
  const arr = filtered.length ? filtered : lots;

  arr.sort((a, b) => {
    const aPref = String(a.lote) === CM.LOTE_PREFERIDO ? 1 : 0;
    const bPref = String(b.lote) === CM.LOTE_PREFERIDO ? 1 : 0;
    if (aPref !== bPref) return bPref - aPref; 
    if (b.stock !== a.stock) return b.stock - a.stock;
    return String(a.lote).localeCompare(String(b.lote));
  });

  return arr[0];
}

/***************************************************************
 * EXPORT HELPERS
 ***************************************************************/
function headerMap_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    const k = norm_(h);
    if (k) map[k] = i; // Guardamos el índice (0-based)
  });
  return map;
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function sanitize_(s) {
  return String(s || '').trim().replace(/[\\/:*?"<>|]+/g, '_').replace(/\s+/g, '_');
}

/***************************************************************
 * HELPERS
 ***************************************************************/
function skuKey_(s) {
  let t = String(s || '');
  try { t = t.normalize('NFKC'); } catch (e) {}
  t = t.replace(/[\u200B-\u200D\uFEFF]/g, '');
  t = t.replace(/^'+/, '');
  t = t.trim().replace(/\s+/g, '');
  t = t.toUpperCase();
  t = t.replace(/[^0-9A-Z]/g, '');
  return t;
}

function parseCLPFromText_(txt) {
  let s = String(txt || '').trim();
  if (!s) return 0;
  s = s.replace(/\s/g, '').replace(/\$/g, '');
  s = s.replace(/[.,]/g, '');
  const n = Number(s.replace(/[^\d-]/g, ''));
  return isNaN(n) ? 0 : Math.round(n);
}

function toInt_(v) {
  if (typeof v === 'number' && isFinite(v)) return Math.floor(v);
  const s = String(v || '').trim();
  if (!s) return 0;
  const n = Number(s.replace(/[^\d-]/g, ''));
  return isNaN(n) ? 0 : Math.floor(n);
}

function effectiveLastRowByColDisplay_(sh, col) {
  const lr = sh.getLastRow();
  if (lr < 2) return lr;
  const vals = sh.getRange(2, col, lr - 1, 1).getDisplayValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0] || '').trim() !== '') return i + 2;
  }
  return 1;
}

function norm_(t) {
  let s = String(t || '').replace(/[✅✔️☑]/g, '');
  try { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); } catch (e) {}
  return s.toUpperCase().trim().replace(/\s+/g, ' ');
}

function colToA1_(col) {
  let s = '';
  while (col > 0) {
    const m = (col - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

function rangesIntersect_(r1, c1, nr1, nc1, r2, c2, nr2, nc2) {
  const r1b = r1 + nr1 - 1;
  const c1b = c1 + nc1 - 1;
  const r2b = r2 + nr2 - 1;
  const c2b = c2 + nc2 - 1;
  return !(r1b < r2 || r2b < r1 || c1b < c2 || c2b < c1);
}

/***************************************************************
 * MAGIA CSV: Convierte las matrices a formato texto puro delimitado
 ***************************************************************/
function arrayToCsv_(data, separator) {
  return data.map(row => 
    row.map(val => {
      let str = String(val || '');
      // Si el texto tiene saltos de línea, el separador, o comillas, lo encierra en comillas dobles
      if (str.includes(separator) || str.includes('"') || str.includes('\n')) {
        str = '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(separator)
  ).join('\r\n');
}
