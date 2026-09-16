/**
 * Control de Deudas — sincronización con Google Sheets.
 *
 * Instrucciones de instalación:
 * 1. Crea una hoja de cálculo nueva en Google Sheets (el nombre no importa).
 * 2. Extensiones > Apps Script.
 * 3. Borra el contenido de Code.gs y pega TODO este archivo.
 * 4. Arriba, en el selector de funciones, elige "setupSheets" y presiona Ejecutar.
 *    La primera vez pedirá autorización — acéptala (es tu propia hoja).
 *    Esto crea las pestañas: Deudas, Telefonia, PersonalesCategorias,
 *    GastosLog, IngresosPersonales, Historial y Config.
 * 5. Implementar > Nueva implementación > tipo "Aplicación web".
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 6. Copia la URL que termina en /exec — esa es la que pegas en la app,
 *    en el pie de página, en "Conectar con Google Sheets".
 *
 * Nota de seguridad: cualquiera con esa URL puede leer o sobrescribir los
 * datos de esta hoja. No la compartas públicamente.
 */

const SHEETS = {
  debts:       { name: 'Deudas',               headers: ['id','nombre','saldo','tasaAnual','pagoMinimo','pagoTotal','prioridad','diaPago','tipo','notas'] },
  telefonia:   { name: 'Telefonia',             headers: ['id','nombre','monto','diaPago'] },
  categorias:  { name: 'PersonalesCategorias',  headers: ['nombre','monto'] },
  gastosLog:   { name: 'GastosLog',             headers: ['id','fecha','categoria','monto','descripcion'] },
  ingresosPersonales: { name: 'IngresosPersonales', headers: ['id','fecha','monto','descripcion'] },
  historial:   { name: 'Historial',             headers: ['id','fecha','nombre','monto','tipo'] },
  config:      { name: 'Config',                headers: ['clave','valor'] },
};

function setupSheets(){
  Object.values(SHEETS).forEach(cfg => getOrCreateSheet(cfg.name, cfg.headers));
}

function getOrCreateSheet(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readTable(cfg){
  const sh = getOrCreateSheet(cfg.name, cfg.headers);
  const values = sh.getDataRange().getValues();
  if(values.length < 2) return [];
  const hdr = values[0];
  return values.slice(1)
    .filter(row => row.some(c => c !== ''))
    .map(row => {
      const obj = {};
      hdr.forEach((h, i) => {
        let v = row[i];
        if(v instanceof Date){
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        obj[h] = v;
      });
      return obj;
    });
}

function writeTable(cfg, rows){
  const sh = getOrCreateSheet(cfg.name, cfg.headers);
  sh.clearContents();
  sh.appendRow(cfg.headers);
  if(rows && rows.length){
    const data = rows.map(r => cfg.headers.map(h => (r[h] !== undefined && r[h] !== null) ? r[h] : ''));
    sh.getRange(2, 1, data.length, cfg.headers.length).setValues(data);
  }
  sh.setFrozenRows(1);
}

function readConfig(){
  const rows = readTable(SHEETS.config);
  const cfg = {};
  rows.forEach(r => { cfg[r.clave] = r.valor; });
  return cfg;
}

function writeConfig(cfgObj){
  const rows = Object.entries(cfgObj).map(([clave, valor]) => ({
    clave,
    valor: typeof valor === 'object' ? JSON.stringify(valor) : valor
  }));
  writeTable(SHEETS.config, rows);
}

function normalizeDebt(d){
  return {
    id: d.id, nombre: d.nombre,
    saldo: Number(d.saldo) || 0, tasaAnual: Number(d.tasaAnual) || 0,
    pagoMinimo: Number(d.pagoMinimo) || 0, pagoTotal: Number(d.pagoTotal) || 0,
    prioridad: Number(d.prioridad) || 99, diaPago: Number(d.diaPago) || 0,
    tipo: d.tipo || 'avalancha', notas: d.notas || ''
  };
}
function normalizeTelefonia(t){
  return { id: t.id, nombre: t.nombre, monto: Number(t.monto) || 0, diaPago: Number(t.diaPago) || 0 };
}
function normalizeCategoria(c){
  return { nombre: c.nombre, monto: Number(c.monto) || 0 };
}
function normalizeGasto(g){
  return { id: g.id, fecha: g.fecha, categoria: g.categoria, monto: Number(g.monto) || 0, descripcion: g.descripcion || '' };
}
function normalizeIngresoPersonal(i){
  return { id: i.id, fecha: i.fecha, monto: Number(i.monto) || 0, descripcion: i.descripcion || '' };
}
function normalizeHistorial(h){
  return { id: h.id, fecha: h.fecha, nombre: h.nombre, monto: Number(h.monto) || 0, tipo: h.tipo };
}

function doGet(e){
  const cfg = readConfig();
  const state = {
    debts: readTable(SHEETS.debts).map(normalizeDebt),
    telefonia: readTable(SHEETS.telefonia).map(normalizeTelefonia),
    personales: {
      presupuesto: Number(cfg.presupuesto) || 0,
      categorias: readTable(SHEETS.categorias).map(normalizeCategoria)
    },
    gastosLog: readTable(SHEETS.gastosLog).map(normalizeGasto),
    ingresosPersonales: readTable(SHEETS.ingresosPersonales).map(normalizeIngresoPersonal),
    historial: readTable(SHEETS.historial).map(normalizeHistorial),
    ultimoPeriodo: cfg.ultimoPeriodo ? JSON.parse(cfg.ultimoPeriodo) : null
  };
  return ContentService.createTextOutput(JSON.stringify(state)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  const body = JSON.parse(e.postData.contents);

  writeTable(SHEETS.debts, body.debts || []);
  writeTable(SHEETS.telefonia, body.telefonia || []);
  writeTable(SHEETS.categorias, (body.personales && body.personales.categorias) || []);
  writeTable(SHEETS.gastosLog, body.gastosLog || []);
  writeTable(SHEETS.ingresosPersonales, body.ingresosPersonales || []);
  writeTable(SHEETS.historial, body.historial || []);
  writeConfig({
    presupuesto: (body.personales && body.personales.presupuesto) || 0,
    ultimoPeriodo: JSON.stringify(body.ultimoPeriodo || null)
  });

  return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
}
