// ═══════════════════════════════════════════════════════════════
// Optimizador de Rutas BAT — Backend Google Apps Script
// ───────────────────────────────────────────────────────────────
// Crea y recuerda su propio Google Sheet "Optimizador de Rutas BAT 2026" con:
//   • Equipo   → promotores/móviles (distritos, casa, días, velocidad, horas)
//   • Config   → reglas del supervisor y parámetros (clave / valor JSON)
//   • Puntos   → padrón de tiendas con coordenadas y frecuencia
//   • Rutas    → última ruta generada (se agrega al guardar desde la web)
// Seguridad: la web envía una CLAVE; cámbiala abajo antes de desplegar.
// ═══════════════════════════════════════════════════════════════
const API_KEY = 'CAMBIA-ESTA-CLAVE';
const SS_NAME = 'Optimizador de Rutas BAT 2026';

const H_EQUIPO = ['Nombre','Tipo','KmH','Horas','Distritos','Vive en','L','M','X','J','V','Supervisor'];
const H_PUNTOS = ['ID','Nombre','Cadena','Distrito','Lat','Lon','Frecuencia','Minutos'];
const H_RUTAS  = ['Guardado','Semana','Recurso','Día','Orden','ID','Nombre','Distrito','Lat','Lon','Llegada est.','Min visita'];

// Ejecuta setup() una vez desde el editor: crea el Sheet y muestra su URL en el registro.
function setup(){
  const ss = ss_();
  Logger.log('Sheet: ' + ss.getUrl());
}

function ss_(){
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty('SS_ID');
  if(id){ try{ ss = SpreadsheetApp.openById(id); }catch(e){ ss = null; } }
  if(!ss){
    ss = SpreadsheetApp.create(SS_NAME);
    props.setProperty('SS_ID', ss.getId());
  }
  sheet_(ss, 'Equipo', H_EQUIPO); sheet_(ss, 'Puntos', H_PUNTOS); sheet_(ss, 'Rutas', H_RUTAS);
  if(!ss.getSheetByName('Config')){ const c = ss.insertSheet('Config'); c.getRange(1,1,1,2).setValues([['Clave','Valor (JSON)']]).setFontWeight('bold'); }
  const first = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if(first && ss.getSheets().length > 1) ss.deleteSheet(first);
  return ss;
}
function sheet_(ss, name, headers){
  let sh = ss.getSheetByName(name);
  if(!sh){ sh = ss.insertSheet(name); sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function table_(ss, name){
  const sh = ss.getSheetByName(name), v = sh.getDataRange().getValues();
  if(v.length < 2) return [];
  const h = v[0];
  return v.slice(1).filter(r => r.some(c => c !== '')).map(r => { const o = {}; h.forEach((k,i) => o[k] = r[i]); return o; });
}
function write_(ss, name, headers, rows, append){
  const sh = ss.getSheetByName(name);
  if(!append && sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,Math.max(sh.getLastColumn(),headers.length)).clearContent();
  if(rows.length) sh.getRange(append ? sh.getLastRow()+1 : 2, 1, rows.length, headers.length).setValues(rows);
}
function json_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ═══════════════════════════════ API ═══════════════════════════════
function doGet(e){
  try{
    const p = (e && e.parameter) || {};
    if(p.key !== API_KEY) return json_({ok:false, error:'Clave incorrecta'});
    if(p.action === 'getAll'){
      const ss = ss_(), cfg = {};
      table_(ss, 'Config').forEach(r => { try{ cfg[r['Clave']] = JSON.parse(r['Valor (JSON)']); }catch(_){} });
      return json_({ok:true, equipo:table_(ss,'Equipo'), config:cfg, puntos:table_(ss,'Puntos'), url:ss.getUrl()});
    }
    return json_({ok:false, error:'Acción no válida'});
  }catch(err){ return json_({ok:false, error:String(err)}); }
}

function doPost(e){
  try{
    const b = JSON.parse(e.postData.contents);
    if(b.key !== API_KEY) return json_({ok:false, error:'Clave incorrecta'});
    const ss = ss_();
    if(b.action === 'saveEquipo'){
      write_(ss, 'Equipo', H_EQUIPO, b.rows.map(o => H_EQUIPO.map(h => o[h] === undefined ? '' : o[h])));
      return json_({ok:true, n:b.rows.length});
    }
    if(b.action === 'saveConfig'){
      const sh = ss.getSheetByName('Config');
      if(sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,2).clearContent();
      const keys = Object.keys(b.config);
      if(keys.length) sh.getRange(2,1,keys.length,2).setValues(keys.map(k => [k, JSON.stringify(b.config[k])]));
      return json_({ok:true});
    }
    if(b.action === 'savePuntos'){
      write_(ss, 'Puntos', H_PUNTOS, b.rows.map(o => H_PUNTOS.map(h => o[h] === undefined ? '' : o[h])));
      return json_({ok:true, n:b.rows.length});
    }
    if(b.action === 'saveRutas'){
      const now = new Date();
      write_(ss, 'Rutas', H_RUTAS, b.rows.map(o => [now].concat(H_RUTAS.slice(1).map(h => o[h] === undefined ? '' : o[h]))), true);
      return json_({ok:true, n:b.rows.length});
    }
    return json_({ok:false, error:'Acción no válida'});
  }catch(err){ return json_({ok:false, error:String(err)}); }
}
