/**
 * Puente entre el CRM y Google Sheets.
 *
 * 1. Cambia la CLAVE de abajo por una clave secreta tuya (sin espacios).
 * 2. Implementar → Nueva implementación → Aplicación web
 *    - Ejecutar como: Yo
 *    - Quién tiene acceso: Cualquier persona
 * 3. Copia la URL que termina en /exec y pégala en el CRM junto con tu clave.
 */

const CLAVE = 'CAMBIA-ESTA-CLAVE';
const NOMBRE_HOJA = 'Prospectos';

// Campo del CRM → título de la columna en la hoja
const CAMPOS = [
  ['id', 'ID'],
  ['name', 'Nombre'],
  ['phone', 'WhatsApp'],
  ['instagram', 'Instagram'],
  ['category', 'Categoría'],
  ['value', 'Valor'],
  ['stage', 'Etapa'],
  ['followStep', 'Seguimiento'],
  ['noRespReason', 'Motivo sin respuesta'],
  ['date', 'Fecha'],
  ['notes', 'Notas'],
  ['extra', 'Otros datos'],
  ['updatedAt', 'Última actualización']
];
const CLAVES_CONOCIDAS = CAMPOS.map(c => c[0]).filter(k => k !== 'extra' && k !== 'updatedAt');

function doGet(e) {
  return responder(() => {
    verificarClave(e.parameter.key);
    return { ok: true, leads: leerProspectos() };
  });
}

function doPost(e) {
  return responder(() => {
    const body = JSON.parse(e.postData.contents);
    verificarClave(body.key);

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      guardarCambios(body.upserts || [], body.deletes || []);
    } finally {
      lock.releaseLock();
    }
    return { ok: true };
  });
}

// ---------- Utilidades ----------

function responder(fn) {
  let resultado;
  try {
    resultado = fn();
  } catch (err) {
    resultado = { ok: false, error: String(err.message || err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(resultado))
    .setMimeType(ContentService.MimeType.JSON);
}

function verificarClave(clave) {
  if (CLAVE === 'CAMBIA-ESTA-CLAVE') {
    throw new Error('Falta configurar la clave en Apps Script');
  }
  if (clave !== CLAVE) {
    throw new Error('Clave incorrecta');
  }
}

function obtenerHoja() {
  const libro = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = libro.getSheetByName(NOMBRE_HOJA);
  if (!hoja) {
    hoja = libro.insertSheet(NOMBRE_HOJA);
  }
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, CAMPOS.length)
      .setValues([CAMPOS.map(c => c[1])])
      .setFontWeight('bold')
      .setBackground('#f3e8ff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function leerProspectos() {
  const hoja = obtenerHoja();
  const ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return [];

  const filas = hoja.getRange(2, 1, ultimaFila - 1, CAMPOS.length).getDisplayValues();
  return filas
    .filter(fila => fila[0])
    .map(fila => {
      let lead = {};
      CAMPOS.forEach(([clave], i) => {
        if (clave === 'extra' || clave === 'updatedAt') return;
        lead[clave] = fila[i];
      });
      lead.value = Number(lead.value) || 0;

      const extra = fila[CAMPOS.findIndex(c => c[0] === 'extra')];
      if (extra) {
        try { lead = Object.assign(JSON.parse(extra), lead); } catch (e) {}
      }
      return lead;
    });
}

function aFila(lead) {
  const extra = {};
  Object.keys(lead).forEach(k => {
    if (CLAVES_CONOCIDAS.indexOf(k) === -1) extra[k] = lead[k];
  });

  return CAMPOS.map(([clave]) => {
    if (clave === 'extra') return Object.keys(extra).length ? JSON.stringify(extra) : '';
    if (clave === 'updatedAt') return new Date().toLocaleString('es-VE');
    const v = lead[clave];
    return v === undefined || v === null ? '' : String(v);
  });
}

function guardarCambios(upserts, deletes) {
  const hoja = obtenerHoja();

  const leerIds = () => {
    const ultima = hoja.getLastRow();
    if (ultima < 2) return [];
    return hoja.getRange(2, 1, ultima - 1, 1).getDisplayValues().map(f => f[0]);
  };

  // Actualizar existentes o agregar nuevos
  const ids = leerIds();
  const nuevos = [];
  upserts.forEach(lead => {
    if (!lead || !lead.id) return;
    const fila = aFila(lead);
    const pos = ids.indexOf(String(lead.id));
    if (pos !== -1) {
      hoja.getRange(pos + 2, 1, 1, CAMPOS.length).setNumberFormat('@').setValues([fila]);
    } else {
      nuevos.push(fila);
    }
  });
  if (nuevos.length) {
    const inicio = hoja.getLastRow() + 1;
    hoja.getRange(inicio, 1, nuevos.length, CAMPOS.length).setNumberFormat('@').setValues(nuevos);
  }

  // Eliminar (de abajo hacia arriba para no descuadrar filas)
  if (deletes.length) {
    const idsActuales = leerIds();
    const aBorrar = deletes.map(String);
    for (let i = idsActuales.length - 1; i >= 0; i--) {
      if (aBorrar.indexOf(idsActuales[i]) !== -1) hoja.deleteRow(i + 2);
    }
  }
}
