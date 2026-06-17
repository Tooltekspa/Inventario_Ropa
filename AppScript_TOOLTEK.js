/**
 * TOOLTEK® — Sistema de Inventario EPP & Uniformes
 * Backend Google Apps Script + Google Sheets
 *
 * INSTALACIÓN:
 * 1. Crea un Google Sheets nuevo (o usa uno existente).
 * 2. Extensiones → Apps Script. Pega este archivo completo.
 * 3. Ejecuta una vez la función setupInicial() (autoriza permisos).
 *    Esto crea las hojas "Inventario", "Actas" y "Nomina" con el catálogo completo.
 * 4. Implementar → Nueva implementación → Aplicación web:
 *    - Ejecutar como: Yo
 *    - Acceso: Cualquier usuario
 * 5. Copia la URL /exec y pégala en API_URL del index.html.
 *
 * API (todo por GET para evitar CORS):
 *   ?action=getAll                                  → {ok, inventario[], actas[], nomina[]}
 *   ?action=saveItem&data=<json item>               → upsert por ID
 *   ?action=deleteItem&id=<ID>                      → elimina fila
 *   ?action=updateStock&id=<ID>&qty=<n>&ubicacion=  → actualiza stock/ubicación
 *   ?action=addActa&data=<json acta>                → asigna N° correlativo,
 *                                                     descuenta stock y guarda acta
 *   ?action=saveWorker&data=<json>&rutOriginal=<r>  → upsert trabajador por RUT
 *   ?action=deleteWorker&rut=<rut>                  → elimina trabajador
 */

const SHEET_INV    = 'Inventario';
const SHEET_ACTAS  = 'Actas';
const SHEET_NOMINA = 'Nomina';

const HEADERS_INV    = ['ID','Barcode','Nombre','Categoria','Tipo','Talla','Modelo','Cantidad','Ubicacion'];
const HEADERS_ACTAS  = ['Num','Tipo','Nombre','RUT','Cargo','Area','Ciudad','Fecha','Obs','Emision','Items'];
const HEADERS_NOMINA = ['rut','nombres','ap','am','sexo','cargo','tenida','cc',
                        't_camisas','t_poleras','t_softshell','t_parka','t_buzo','t_pantalon','t_zapatos',
                        'entregas_json'];

// ═══ ENDPOINT ═════════════════════════════════════════════════════════════════
function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'getAll';
  let result;
  try {
    switch (action) {
      case 'getAll':        result = getAll(); break;
      case 'saveItem':      result = withLock(() => saveItem(JSON.parse(p.data))); break;
      case 'deleteItem':    result = withLock(() => deleteItem(p.id)); break;
      case 'updateStock':   result = withLock(() => updateStock(p.id, p.qty, p.ubicacion)); break;
      case 'addActa':       result = withLock(() => addActa(JSON.parse(p.data))); break;
      case 'saveWorker':    result = withLock(() => saveWorker(JSON.parse(p.data), p.rutOriginal)); break;
      case 'deleteWorker':  result = withLock(() => deleteWorker(p.rut)); break;
      default: result = { ok: false, error: 'Acción desconocida: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ═══ HELPERS ══════════════════════════════════════════════════════════════════
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('No existe la hoja "' + name + '". Ejecuta setupInicial().');
  return sh;
}

function fechaStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

function leerInventario() {
  const sh = getSheet(SHEET_INV);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, HEADERS_INV.length).getValues();
  return data
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      id:        String(r[0]).trim(),
      barcode:   String(r[1] || ''),
      nombre:    String(r[2] || ''),
      cat:       String(r[3] || 'Prenda'),
      tipo:      String(r[4] || ''),
      talla:     String(r[5] || ''),
      modelo:    String(r[6] || ''),
      qty:       parseInt(r[7], 10) || 0,
      ubicacion: String(r[8] || '')
    }));
}

// Devuelve el número de fila (1-based) donde col[0] === id, o -1
function filaDeId(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return -1;
}

// Busca por columna de nombre arbitrario (para Nomina)
function filaDeColumna(sh, colIndex, value) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, colIndex, last - 1, 1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).trim() === String(value).trim()) return i + 2;
  }
  return -1;
}

// ═══ ACCIONES — INVENTARIO Y ACTAS ════════════════════════════════════════════
function getAll() {
  // ── Actas ──
  const shAct = getSheet(SHEET_ACTAS);
  const lastAct = shAct.getLastRow();
  let actas = [];
  if (lastAct >= 2) {
    const data = shAct.getRange(2, 1, lastAct - 1, HEADERS_ACTAS.length).getValues();
    actas = data
      .filter(r => String(r[0]).trim() !== '')
      .map(r => {
        let items = [];
        try { items = JSON.parse(r[10] || '[]'); } catch (e) {}
        return {
          num:     String(r[0]).trim(),
          tipo:    String(r[1] || ''),
          nombre:  String(r[2] || ''),
          rut:     String(r[3] || ''),
          cargo:   String(r[4] || ''),
          area:    String(r[5] || ''),
          ciudad:  String(r[6] || ''),
          fecha:   fechaStr(r[7]),
          obs:     String(r[8] || ''),
          emision: String(r[9] || ''),
          items:   items
        };
      })
      .sort((a, b) => b.num.localeCompare(a.num));
  }

  // ── Nómina ──
  const nomina = leerNomina();

  return { ok: true, inventario: leerInventario(), actas: actas, nomina: nomina };
}

function saveItem(item) {
  const sh = getSheet(SHEET_INV);
  const fila = [
    item.id, item.barcode || '', item.nombre || '', item.cat || 'Prenda',
    item.tipo || '', item.talla || '', item.modelo || '',
    parseInt(item.qty, 10) || 0, item.ubicacion || ''
  ];
  const r = filaDeId(sh, item.id);
  if (r > 0) sh.getRange(r, 1, 1, fila.length).setValues([fila]);
  else sh.appendRow(fila);
  return { ok: true };
}

function deleteItem(id) {
  const sh = getSheet(SHEET_INV);
  const r = filaDeId(sh, id);
  if (r > 0) sh.deleteRow(r);
  return { ok: true };
}

function updateStock(id, qty, ubicacion) {
  const sh = getSheet(SHEET_INV);
  const r = filaDeId(sh, id);
  if (r < 0) return { ok: false, error: 'ID no encontrado: ' + id };
  sh.getRange(r, 8).setValue(parseInt(qty, 10) || 0);
  if (ubicacion !== undefined) sh.getRange(r, 9).setValue(ubicacion || '');
  return { ok: true };
}

function addActa(acta) {
  const shInv = getSheet(SHEET_INV);
  const shAct = getSheet(SHEET_ACTAS);

  // 1. Validar stock en servidor (fuente de verdad)
  const errores = [];
  const filas = {};
  (acta.items || []).forEach(it => {
    const r = filaDeId(shInv, it.codigo);
    if (r < 0) { errores.push('Código no encontrado: ' + it.codigo); return; }
    const disp = parseInt(shInv.getRange(r, 8).getValue(), 10) || 0;
    if (it.qty > disp) errores.push('Stock insuficiente: ' + it.nombre + ' (disp: ' + disp + ')');
    filas[it.codigo] = { row: r, disp: disp };
  });
  if (errores.length) return { ok: false, error: errores.join('\n') };

  // 2. Correlativo: máximo existente + 1
  let max = 0;
  const last = shAct.getLastRow();
  if (last >= 2) {
    shAct.getRange(2, 1, last - 1, 1).getValues().forEach(r => {
      const n = parseInt(r[0], 10);
      if (!isNaN(n) && n > max) max = n;
    });
  }
  const num = String(max + 1).padStart(4, '0');

  // 3. Descontar stock
  (acta.items || []).forEach(it => {
    const f = filas[it.codigo];
    shInv.getRange(f.row, 8).setValue(f.disp - it.qty);
  });

  // 4. Guardar acta (Num como texto para conservar ceros)
  shAct.appendRow([
    "'" + num, acta.tipo || '', acta.nombre || '', acta.rut || '',
    acta.cargo || '', acta.area || '', acta.ciudad || '',
    acta.fecha || '', acta.obs || '', acta.emision || '',
    JSON.stringify(acta.items || [])
  ]);

  return { ok: true, num: num, inventario: leerInventario() };
}

// ═══ ACCIONES — NÓMINA ════════════════════════════════════════════════════════
const ENTREGAS_DEFAULT = () => ({
  camisas:  {entregado:false, fecha:'', cantidad:0, obs:''},
  poleras:  {entregado:false, fecha:'', cantidad:0, obs:''},
  softshell:{entregado:false, fecha:'', cantidad:0, obs:''},
  parka:    {entregado:false, fecha:'', cantidad:0, obs:''},
  buzo:     {entregado:false, fecha:'', cantidad:0, obs:''},
  pantalon: {entregado:false, fecha:'', cantidad:0, obs:''},
  zapatos:  {entregado:false, fecha:'', cantidad:0, obs:''}
});

function leerNomina() {
  const sh = getSheet(SHEET_NOMINA);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, HEADERS_NOMINA.length).getValues();
  return data
    .filter(r => String(r[0]).trim() !== '')
    .map(r => {
      let entregas = ENTREGAS_DEFAULT();
      try {
        if (r[15]) entregas = Object.assign(ENTREGAS_DEFAULT(), JSON.parse(r[15]));
      } catch (e) {}
      return {
        rut:    String(r[0]).trim(),
        nombres: String(r[1] || ''),
        ap:     String(r[2] || ''),
        am:     String(r[3] || ''),
        sexo:   String(r[4] || 'M'),
        cargo:  String(r[5] || ''),
        tenida: String(r[6] || ''),
        cc:     String(r[7] || ''),
        t: {
          camisas:   String(r[8]  || ''),
          poleras:   String(r[9]  || ''),
          softshell: String(r[10] || ''),
          parka:     String(r[11] || ''),
          buzo:      String(r[12] || ''),
          pantalon:  String(r[13] || ''),
          zapatos:   String(r[14] || ''),
        },
        e: entregas
      };
    });
}

// upsert por RUT — crea fila si no existe, actualiza si existe
function saveWorker(w, rutOriginal) {
  const sh = getSheet(SHEET_NOMINA);
  const rutBuscar = rutOriginal || w.rut;

  const fila = [
    w.rut, w.nombres || '', w.ap || '', w.am || '',
    w.sexo || 'M', w.cargo || '', w.tenida || '', w.cc || '',
    (w.t && w.t.camisas)   || '',
    (w.t && w.t.poleras)   || '',
    (w.t && w.t.softshell) || '',
    (w.t && w.t.parka)     || '',
    (w.t && w.t.buzo)      || '',
    (w.t && w.t.pantalon)  || '',
    (w.t && w.t.zapatos)   || '',
    JSON.stringify(w.e || ENTREGAS_DEFAULT())
  ];

  const rowN = filaDeColumna(sh, 1, rutBuscar); // col 1 = rut
  if (rowN > 0) {
    sh.getRange(rowN, 1, 1, fila.length).setValues([fila]);
  } else {
    sh.appendRow(fila);
  }
  return { ok: true };
}

function deleteWorker(rut) {
  const sh = getSheet(SHEET_NOMINA);
  const rowN = filaDeColumna(sh, 1, rut);
  if (rowN < 0) return { ok: false, error: 'Trabajador no encontrado: ' + rut };
  sh.deleteRow(rowN);
  return { ok: true };
}

// ═══ SETUP INICIAL (ejecutar una sola vez manualmente) ════════════════════════
function setupInicial() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Hoja Inventario ──
  let shInv = ss.getSheetByName(SHEET_INV);
  if (!shInv) shInv = ss.insertSheet(SHEET_INV);
  if (shInv.getLastRow() === 0) {
    shInv.getRange(1, 1, 1, HEADERS_INV.length).setValues([HEADERS_INV])
      .setFontWeight('bold').setBackground('#25638f').setFontColor('#ffffff');
    shInv.setFrozenRows(1);
    const seed = catalogoInicial();
    shInv.getRange(2, 1, seed.length, HEADERS_INV.length).setValues(seed);
    shInv.getRange(2, 1, seed.length, 1).setNumberFormat('@');
    shInv.autoResizeColumns(1, HEADERS_INV.length);
  }

  // ── Hoja Actas ──
  let shAct = ss.getSheetByName(SHEET_ACTAS);
  if (!shAct) shAct = ss.insertSheet(SHEET_ACTAS);
  if (shAct.getLastRow() === 0) {
    shAct.getRange(1, 1, 1, HEADERS_ACTAS.length).setValues([HEADERS_ACTAS])
      .setFontWeight('bold').setBackground('#0081b0').setFontColor('#ffffff');
    shAct.setFrozenRows(1);
    shAct.getRange('A:A').setNumberFormat('@');
    shAct.getRange('H:H').setNumberFormat('@');
  }

  // ── Hoja Nómina — siempre verifica/corrige cabecera ──
  let shNom = ss.getSheetByName(SHEET_NOMINA);
  if (!shNom) shNom = ss.insertSheet(SHEET_NOMINA);

  // Verificar si la cabecera actual es correcta
  const cabActual = shNom.getLastRow() > 0
    ? shNom.getRange(1, 1, 1, HEADERS_NOMINA.length).getValues()[0].map(String)
    : [];
  const cabCorrecta = HEADERS_NOMINA.every((h, i) => cabActual[i] === h);

  if (!cabCorrecta) {
    // Reescribir solo fila 1; datos desde fila 2 se preservan
    shNom.getRange(1, 1, 1, HEADERS_NOMINA.length).setValues([HEADERS_NOMINA]);
    Logger.log('Cabecera Nomina actualizada: ' + HEADERS_NOMINA.join(', '));
  }

  // Siempre aplicar formato a fila 1
  shNom.getRange(1, 1, 1, HEADERS_NOMINA.length)
    .setFontWeight('bold').setBackground('#1a7a4a').setFontColor('#ffffff');
  shNom.setFrozenRows(1);
  shNom.getRange('A:A').setNumberFormat('@'); // RUT como texto
  shNom.autoResizeColumns(1, HEADERS_NOMINA.length);

  Logger.log('Setup completo: ' + SHEET_INV + ', ' + SHEET_ACTAS + ' y ' + SHEET_NOMINA + ' listas.');
}

// ═══ CATÁLOGO INICIAL ═════════════════════════════════════════════════════════
function catalogoInicial() {
  return [
    ["PPSH","Polera Pique S Hombre","Polera Pique S Hombre","Prenda","Pique","S","Hombre",5,"2"],
    ["PPSM","Polera Pique S Mujer","Polera Pique S Mujer","Prenda","Pique","S","Mujer",6,"2"],
    ["PPMH","Polera Pique M Hombre","Polera Pique M Hombre","Prenda","Pique","M","Hombre",10,""],
    ["PPMM","Polera Pique M Mujer","Polera Pique M Mujer","Prenda","Pique","M","Mujer",0,""],
    ["PPLH","Polera Pique L Hombre","Polera Pique L Hombre","Prenda","Pique","L","Hombre",0,""],
    ["PPLM","Polera Pique L Mujer","Polera Pique L Mujer","Prenda","Pique","L","Mujer",0,""],
    ["PPXH","Polera Pique XL Hombre","Polera Pique XL Hombre","Prenda","Pique","XL","Hombre",0,""],
    ["PPXM","Polera Pique XL Mujer","Polera Pique XL Mujer","Prenda","Pique","XL","Mujer",0,""],
    ["PPXXH","Polera Pique XXL Hombre","Polera Pique XXL Hombre","Prenda","Pique","XXL","Hombre",0,""],
    ["PPXXM","Polera Pique XXL Mujer","Polera Pique XXL Mujer","Prenda","Pique","XXL","Mujer",0,""],
    ["PP3H","Polera Pique 3XL Hombre","Polera Pique 3XL Hombre","Prenda","Pique","3XL","Hombre",0,""],
    ["PP3M","Polera Pique 3XL Mujer","Polera Pique 3XL Mujer","Prenda","Pique","3XL","Mujer",0,""],
    ["PP4H","Polera Pique 4XL Hombre","Polera Pique 4XL Hombre","Prenda","Pique","4XL","Hombre",0,""],
    ["PP4M","Polera Pique 4XL Mujer","Polera Pique 4XL Mujer","Prenda","Pique","4XL","Mujer",0,""],
    ["POLOSH","Polera Polo S Hombre","Polera Polo S Hombre","Prenda","Polo","S","Hombre",0,""],
    ["POLOSM","Polera Polo S Mujer","Polera Polo S Mujer","Prenda","Polo","S","Mujer",0,""],
    ["POLOMH","Polera Polo M Hombre","Polera Polo M Hombre","Prenda","Polo","M","Hombre",0,""],
    ["POLOMM","Polera Polo M Mujer","Polera Polo M Mujer","Prenda","Polo","M","Mujer",0,""],
    ["POLOLH","Polera Polo L Hombre","Polera Polo L Hombre","Prenda","Polo","L","Hombre",0,""],
    ["POLOLM","Polera Polo L Mujer","Polera Polo L Mujer","Prenda","Polo","L","Mujer",0,""],
    ["POLOXH","Polera Polo XL Hombre","Polera Polo XL Hombre","Prenda","Polo","XL","Hombre",0,""],
    ["POLOXM","Polera Polo XL Mujer","Polera Polo XL Mujer","Prenda","Polo","XL","Mujer",0,""],
    ["POLOXXH","Polera Polo XXL Hombre","Polera Polo XXL Hombre","Prenda","Polo","XXL","Hombre",0,""],
    ["POLOXXM","Polera Polo XXL Mujer","Polera Polo XXL Mujer","Prenda","Polo","XXL","Mujer",0,""],
    ["POLO3H","Polera Polo 3XL Hombre","Polera Polo 3XL Hombre","Prenda","Polo","3XL","Hombre",0,""],
    ["POLO3M","Polera Polo 3XL Mujer","Polera Polo 3XL Mujer","Prenda","Polo","3XL","Mujer",0,""],
    ["POLO4H","Polera Polo 4XL Hombre","Polera Polo 4XL Hombre","Prenda","Polo","4XL","Hombre",0,""],
    ["POLO4M","Polera Polo 4XL Mujer","Polera Polo 4XL Mujer","Prenda","Polo","4XL","Mujer",0,""],
    ["PBSH","Pantalon Buzo S Hombre","Pantalon Buzo S Hombre","Prenda","Buzo","S","Hombre",0,""],
    ["PBSM","Pantalon Buzo S Mujer","Pantalon Buzo S Mujer","Prenda","Buzo","S","Mujer",0,""],
    ["PBMH","Pantalon Buzo M Hombre","Pantalon Buzo M Hombre","Prenda","Buzo","M","Hombre",0,""],
    ["PBMM","Pantalon Buzo M Mujer","Pantalon Buzo M Mujer","Prenda","Buzo","M","Mujer",0,""],
    ["PBLH","Pantalon Buzo L Hombre","Pantalon Buzo L Hombre","Prenda","Buzo","L","Hombre",0,""],
    ["PBLM","Pantalon Buzo L Mujer","Pantalon Buzo L Mujer","Prenda","Buzo","L","Mujer",0,""],
    ["PBXH","Pantalon Buzo XL Hombre","Pantalon Buzo XL Hombre","Prenda","Buzo","XL","Hombre",0,""],
    ["PBXM","Pantalon Buzo XL Mujer","Pantalon Buzo XL Mujer","Prenda","Buzo","XL","Mujer",0,""],
    ["PBXXH","Pantalon Buzo XXL Hombre","Pantalon Buzo XXL Hombre","Prenda","Buzo","XXL","Hombre",0,""],
    ["PBXXM","Pantalon Buzo XXL Mujer","Pantalon Buzo XXL Mujer","Prenda","Buzo","XXL","Mujer",0,""],
    ["PB3H","Pantalon Buzo 3XL Hombre","Pantalon Buzo 3XL Hombre","Prenda","Buzo","3XL","Hombre",0,""],
    ["PB3M","Pantalon Buzo 3XL Mujer","Pantalon Buzo 3XL Mujer","Prenda","Buzo","3XL","Mujer",0,""],
    ["PB4H","Pantalon Buzo 4XL Hombre","Pantalon Buzo 4XL Hombre","Prenda","Buzo","4XL","Hombre",0,""],
    ["PB4M","Pantalon Buzo 4XL Mujer","Pantalon Buzo 4XL Mujer","Prenda","Buzo","4XL","Mujer",0,""],
    ["PJSH","Poleron Jogger S Hombre","Poleron Jogger S Hombre","Prenda","Jogger","S","Hombre",0,""],
    ["PJSM","Poleron Jogger S Mujer","Poleron Jogger S Mujer","Prenda","Jogger","S","Mujer",0,""],
    ["PJMH","Poleron Jogger M Hombre","Poleron Jogger M Hombre","Prenda","Jogger","M","Hombre",0,""],
    ["PJMM","Poleron Jogger M Mujer","Poleron Jogger M Mujer","Prenda","Jogger","M","Mujer",0,""],
    ["PJLH","Poleron Jogger L Hombre","Poleron Jogger L Hombre","Prenda","Jogger","L","Hombre",0,""],
    ["PJLM","Poleron Jogger L Mujer","Poleron Jogger L Mujer","Prenda","Jogger","L","Mujer",0,""],
    ["PJXH","Poleron Jogger XL Hombre","Poleron Jogger XL Hombre","Prenda","Jogger","XL","Hombre",0,""],
    ["PJXM","Poleron Jogger XL Mujer","Poleron Jogger XL Mujer","Prenda","Jogger","XL","Mujer",0,""],
    ["PJXXH","Poleron Jogger XXL Hombre","Poleron Jogger XXL Hombre","Prenda","Jogger","XXL","Hombre",0,""],
    ["PJXXM","Poleron Jogger XXL Mujer","Poleron Jogger XXL Mujer","Prenda","Jogger","XXL","Mujer",0,""],
    ["PJ3H","Poleron Jogger 3XL Hombre","Poleron Jogger 3XL Hombre","Prenda","Jogger","3XL","Hombre",0,""],
    ["PJ3M","Poleron Jogger 3XL Mujer","Poleron Jogger 3XL Mujer","Prenda","Jogger","3XL","Mujer",0,""],
    ["PJ4H","Poleron Jogger 4XL Hombre","Poleron Jogger 4XL Hombre","Prenda","Jogger","4XL","Hombre",0,""],
    ["PJ4M","Poleron Jogger 4XL Mujer","Poleron Jogger 4XL Mujer","Prenda","Jogger","4XL","Mujer",0,""],
    ["SSH","Softshell S Hombre","Softshell S Hombre","Prenda","Softshell","S","Hombre",0,""],
    ["SSM","Softshell S Mujer","Softshell S Mujer","Prenda","Softshell","S","Mujer",0,""],
    ["SMH","Softshell M Hombre","Softshell M Hombre","Prenda","Softshell","M","Hombre",0,""],
    ["SMM","Softshell M Mujer","Softshell M Mujer","Prenda","Softshell","M","Mujer",0,""],
    ["SLH","Softshell L Hombre","Softshell L Hombre","Prenda","Softshell","L","Hombre",0,""],
    ["SLM","Softshell L Mujer","Softshell L Mujer","Prenda","Softshell","L","Mujer",0,""],
    ["SXH","Softshell XL Hombre","Softshell XL Hombre","Prenda","Softshell","XL","Hombre",0,""],
    ["SXM","Softshell XL Mujer","Softshell XL Mujer","Prenda","Softshell","XL","Mujer",0,""],
    ["SXXH","Softshell XXL Hombre","Softshell XXL Hombre","Prenda","Softshell","XXL","Hombre",0,""],
    ["SXXM","Softshell XXL Mujer","Softshell XXL Mujer","Prenda","Softshell","XXL","Mujer",0,""],
    ["S3H","Softshell 3XL Hombre","Softshell 3XL Hombre","Prenda","Softshell","3XL","Hombre",0,""],
    ["S3M","Softshell 3XL Mujer","Softshell 3XL Mujer","Prenda","Softshell","3XL","Mujer",0,""],
    ["S4H","Softshell 4XL Hombre","Softshell 4XL Hombre","Prenda","Softshell","4XL","Hombre",0,""],
    ["S4M","Softshell 4XL Mujer","Softshell 4XL Mujer","Prenda","Softshell","4XL","Mujer",0,""],
    ["CCSH","Camisas Celeste S Hombre","Camisa Celeste S Hombre","Prenda","Celeste","S","Hombre",0,""],
    ["CCSM","Camisas Celeste S Mujer","Camisa Celeste S Mujer","Prenda","Celeste","S","Mujer",0,""],
    ["CCMH","Camisas Celeste M Hombre","Camisa Celeste M Hombre","Prenda","Celeste","M","Hombre",0,""],
    ["CCMM","Camisas Celeste M Mujer","Camisa Celeste M Mujer","Prenda","Celeste","M","Mujer",0,""],
    ["CCLH","Camisas Celeste L Hombre","Camisa Celeste L Hombre","Prenda","Celeste","L","Hombre",0,""],
    ["CCLM","Camisas Celeste L Mujer","Camisa Celeste L Mujer","Prenda","Celeste","L","Mujer",0,""],
    ["CCXH","Camisas Celeste XL Hombre","Camisa Celeste XL Hombre","Prenda","Celeste","XL","Hombre",0,""],
    ["CCXM","Camisas Celeste XL Mujer","Camisa Celeste XL Mujer","Prenda","Celeste","XL","Mujer",0,""],
    ["CCXXH","Camisas Celeste XXL Hombre","Camisa Celeste XXL Hombre","Prenda","Celeste","XXL","Hombre",0,""],
    ["CCXXM","Camisas Celeste XXL Mujer","Camisa Celeste XXL Mujer","Prenda","Celeste","XXL","Mujer",0,""],
    ["CC3H","Camisas Celeste 3XL Hombre","Camisa Celeste 3XL Hombre","Prenda","Celeste","3XL","Hombre",0,""],
    ["CC3M","Camisas Celeste 3XL Mujer","Camisa Celeste 3XL Mujer","Prenda","Celeste","3XL","Mujer",0,""],
    ["CC4H","Camisas Celeste 4XL Hombre","Camisa Celeste 4XL Hombre","Prenda","Celeste","4XL","Hombre",0,""],
    ["CC4M","Camisas Celeste 4XL Mujer","Camisa Celeste 4XL Mujer","Prenda","Celeste","4XL","Mujer",0,""],
    ["GHSU","Geologo Hemic S Unisex","Geologo Hemic S Unisex","Prenda","Hemic","S","Unisex",0,""],
    ["GHMU","Geologo Hemic M Unisex","Geologo Hemic M Unisex","Prenda","Hemic","M","Unisex",0,""],
    ["GHLU","Geologo Hemic L Unisex","Geologo Hemic L Unisex","Prenda","Hemic","L","Unisex",0,""],
    ["GHXU","Geologo Hemic XL Unisex","Geologo Hemic XL Unisex","Prenda","Hemic","XL","Unisex",0,""],
    ["GHXXU","Geologo Hemic XXL Unisex","Geologo Hemic XXL Unisex","Prenda","Hemic","XXL","Unisex",0,""],
    ["GH3U","Geologo Hemic 3XL Unisex","Geologo Hemic 3XL Unisex","Prenda","Hemic","3XL","Unisex",0,""],
    ["GH4U","Geologo Hemic 4XL Unisex","Geologo Hemic 4XL Unisex","Prenda","Hemic","4XL","Unisex",0,""],
    ["ZAP40","Zapatos 40","Zapatos Talla 40","Prenda","Zapatos","40","Unisex",0,""],
    ["ZAP41","Zapatos 41","Zapatos Talla 41","Prenda","Zapatos","41","Unisex",0,""],
    ["ZAP42","Zapatos 42","Zapatos Talla 42","Prenda","Zapatos","42","Unisex",0,""],
    ["ZAP43","Zapatos 43","Zapatos Talla 43","Prenda","Zapatos","43","Unisex",0,""],
    ["ZAP44","Zapatos 44","Zapatos Talla 44","Prenda","Zapatos","44","Unisex",0,""],
    ["ZAP45","Zapatos 45","Zapatos Talla 45","Prenda","Zapatos","45","Unisex",0,""],
    ["ZAP46","Zapatos 46","Zapatos Talla 46","Prenda","Zapatos","46","Unisex",0,""]
  ];
}
