const ACTIVIDADES = {
  1: 'Visita al Estadio Centenario y Museo del Fútbol · 14:00 a 15:30 h',
  2: 'Visita al Gran Parque Central · 15:30 a 19:00 h',
  3: 'Partido de fútbol mixto · 16:30 a 18:30 h',
  4: 'Merienda en Parque Central · Desde las 19:00 h · Costo del interesado',
  6: 'Bus turístico por Montevideo · 15:00 a 19:00 h · Costo del interesado',
  7: 'Actividad libre · Desde las 14:00 h hasta el día siguiente · Costo del interesado'
};
function configurar() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SPREADSHEET_ID')) {
    const ss = SpreadsheetApp.create('Inscripciones actividades recreativas OBA 2026');
    const hoja = ss.getSheets()[0]; hoja.setName('Inscripciones');
    hoja.appendRow(['Fecha','Comprobante','Nombre','Apellido','Correo','Teléfono','País','Institución','Actividades','Observaciones']);
    hoja.setFrozenRows(1);
    p.setProperty('SPREADSHEET_ID', ss.getId());
    console.log('Planilla de inscripciones: ' + ss.getUrl());
  } else console.log('Planilla: https://docs.google.com/spreadsheets/d/' + p.getProperty('SPREADSHEET_ID'));
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    ['nombre','apellido','correo','telefono','pais','institucion'].forEach(k => {if(typeof d[k] !== 'string' || !d[k].trim() || d[k].length > 150) throw Error('Complete los datos obligatorios.');});
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.correo)) throw Error('Correo inválido.');
    if(typeof d.id !== 'string' || !/^[a-f0-9-]{36}$/.test(d.id)) throw Error('Comprobante inválido.');
    if(!Array.isArray(d.actividades) || !d.actividades.length || d.actividades.some(n=>!ACTIVIDADES[n])) throw Error('Seleccione actividades válidas.');
    if((d.actividades.includes(7)&&d.actividades.length>1)||(d.actividades.includes(6)&&d.actividades.some(n=>n<=3))) throw Error('Las actividades coinciden en horario.');
    if(typeof d.observaciones !== 'string' || d.observaciones.length>1000) throw Error('Revise las observaciones.');
    lock.waitLock(20000);
    const ssid = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if(!ssid) throw Error('La organización debe ejecutar configurar antes de recibir inscripciones.');
    const hoja=SpreadsheetApp.openById(ssid).getSheetByName('Inscripciones');
    const last=hoja.getLastRow();
    const existe=last>1 && hoja.getRange(2,2,last-1,1).createTextFinder(d.id).matchEntireCell(true).findNext();
    const safe=v=> /^[=+@\-]/.test(String(v)) ? "'"+v : String(v);
    if(!existe) hoja.appendRow([new Date(),d.id,...['nombre','apellido','correo','telefono','pais','institucion'].map(k=>safe(d[k].trim())),d.actividades.map(n=>ACTIVIDADES[n]).join(' | '),safe(d.observaciones)]);
    return respuesta({ok:true,id:d.id});
  } catch(err) { return respuesta({ok:false,error:err.message || 'No se pudo registrar la inscripción.'}); }
  finally {if(lock.hasLock()) lock.releaseLock();}
}
function respuesta(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
