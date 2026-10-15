/**
 * Backend de leads del simulador "Turno Noche" (ProBatch).
 * Guarda cada visita y cada respuesta en esta planilla y avisa por mail cuando llega un lead.
 */
const AVISO_EMAIL = 'enmanueltolosa@gmail.com'; // cambialo por la casilla de Accusys si corresponde

const COLUMNAS = {
  leads: ['fecha', 'banco', 'institucion', 'nombre', 'cargo', 'email', 'rol', 'core', 'recorrido', 'link',
          'identificacion', 'problemas', 'capacidad', 'diagnostico', 'comentarios', 'informe'],
  visitas: ['fecha', 'evento', 'banco', 'rol', 'core', 'recorrido', 'link'],
};

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents || '{}');
    const hoja = d.tipo === 'respuesta' ? 'leads' : 'visitas';
    const sh = obtenerHoja_(hoja);
    sh.appendRow(COLUMNAS[hoja].map(c => {
      const v = d[c];
      return Array.isArray(v) ? v.join(', ') : (v === undefined ? '' : String(v).slice(0, 5000));
    }));
    if (hoja === 'leads' && AVISO_EMAIL) {
      MailApp.sendEmail(AVISO_EMAIL,
        'Nuevo lead ProBatch · ' + (d.institucion || d.banco || 'sin banco'),
        ['Nombre: ' + (d.nombre || '—'), 'Cargo: ' + (d.cargo || '—'), 'Institución: ' + (d.institucion || d.banco || '—'),
         'Email: ' + (d.email || '—'), 'Rol: ' + d.rol + ' · Core: ' + (d.core || '—'),
         'Diagnóstico 30 min: ' + (d.diagnostico || '—'), 'Comentarios: ' + (d.comentarios || '—'), '', d.informe || ''].join('\n'));
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function obtenerHoja_(nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) { sh = ss.insertSheet(nombre); sh.appendRow(COLUMNAS[nombre]); sh.setFrozenRows(1); }
  return sh;
}

// Ejecutá esta función una vez desde el editor para autorizar permisos y crear las hojas.
function prueba() {
  doPost({ postData: { contents: JSON.stringify({ tipo: 'visita', evento: 'prueba', fecha: new Date().toISOString(), banco: 'Banco Demo', rol: 'cio' }) } });
}
