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
    hoja.appendRow(['Fecha','Comprobante','Nombre','Apellido','Teléfono','País','Institución','Actividades','Observaciones']);
    hoja.setFrozenRows(1);
    p.setProperty('SPREADSHEET_ID', ss.getId());
    console.log('Planilla de inscripciones: ' + ss.getUrl());
  } else console.log('Planilla: https://docs.google.com/spreadsheets/d/' + p.getProperty('SPREADSHEET_ID'));
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const d = JSON.parse(e.postData.contents);
    if(d.action === 'listar' || d.action === 'eliminar') {
      const clave = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
      if(!clave || typeof d.password !== 'string' || d.password !== clave) throw Error('Contraseña administrativa incorrecta.');
      lock.waitLock(20000);
      const ssid = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if(!ssid) throw Error('Ejecute configurar primero.');
      const hoja = SpreadsheetApp.openById(ssid).getSheetByName('Inscripciones');
      const rows = hoja.getDataRange().getValues();
      const headers = rows[0];
      const pos = name => headers.indexOf(name);
      if(d.action === 'eliminar') {
        if(typeof d.id !== 'string' || !/^[a-f0-9-]{36}$/.test(d.id)) throw Error('Comprobante inválido.');
        const i = rows.findIndex((r,i)=>i>0 && String(r[pos('Comprobante')])===d.id);
        if(i<1) throw Error('El registro ya no existe. Actualice el listado.');
        hoja.deleteRow(i+1);
        return respuesta({ok:true,id:d.id});
      }
      return respuesta({ok:true,registros:rows.slice(1).filter(r=>r[pos('Comprobante')]).map(r=>({
        id:String(r[pos('Comprobante')]),fecha:r[pos('Fecha')],nombre:r[pos('Nombre')],apellido:r[pos('Apellido')],
        telefono:r[pos('Teléfono')],pais:r[pos('País')],institucion:r[pos('Institución')],actividades:r[pos('Actividades')],observaciones:r[pos('Observaciones')]
      }))});
    }
    ['nombre','apellido','telefono','pais','institucion'].forEach(k => {if(typeof d[k] !== 'string' || !d[k].trim() || d[k].length > 150) throw Error('Complete los datos obligatorios.');});
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
    if(!existe) {
      const valores = {'Fecha':new Date(),'Comprobante':d.id,'Nombre':safe(d.nombre.trim()),'Apellido':safe(d.apellido.trim()),
        'Teléfono':safe(d.telefono.trim()),'País':safe(d.pais.trim()),'Institución':safe(d.institucion.trim()),
        'Actividades':d.actividades.map(n=>ACTIVIDADES[n]).join(' | '),'Observaciones':safe(d.observaciones)};
      // Conserva la estructura de planillas anteriores, dejando vacía su columna Correo.
      const cabeceras = hoja.getRange(1,1,1,hoja.getLastColumn()).getValues()[0];
      hoja.appendRow(cabeceras.map(k=>Object.prototype.hasOwnProperty.call(valores,k)?valores[k]:''));
    }
    return respuesta({ok:true,id:d.id});
  } catch(err) { return respuesta({ok:false,error:err.message || 'No se pudo registrar la inscripción.'}); }
  finally {if(lock.hasLock()) lock.releaseLock();}
}
function respuesta(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
