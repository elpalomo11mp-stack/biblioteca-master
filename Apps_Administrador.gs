/** SEGUNDO proyecto Apps Script, vinculado a un Google Sheets.
 * No sustituye el Apps Script de libros. Hoja: Aplicativos.
 * Columnas: Categoria, Nombre, URL, Imagen, Descripcion, Orden, Visible.
 */
const SHEET_NAME = 'Aplicativos';
function doGet(e) {
  let result;
  try {
    const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if(!sh) throw new Error('Falta la hoja Aplicativos. Ejecuta prepararPlantilla().');
    const values=sh.getDataRange().getDisplayValues();
    const apps=values.slice(1).filter(r=>r[1]&&r[2]).map(r=>({
      categoria:r[0]||'Otros aplicativos',nombre:r[1],url:r[2],imagen:r[3]||'',
      descripcion:r[4]||'',orden:Number(r[5])||999,
      visible:!['NO','FALSE','0','OCULTO'].includes(String(r[6]).trim().toUpperCase())
    }));
    result={ok:true,apps:apps,updated:new Date().toISOString()};
  }catch(err){result={ok:false,error:String(err)};}
  const cb=e&&e.parameter&&e.parameter.callback;
  if(cb && /^[a-zA-Z_$][\w$]*$/.test(cb))
    return ContentService.createTextOutput(cb+'('+JSON.stringify(result)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
function prepararPlantilla(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(SHEET_NAME);
  if(sh)throw new Error('Ya existe la hoja Aplicativos. No se modificó para proteger tus datos.');
  sh=ss.insertSheet(SHEET_NAME);
  sh.getRange(1,1,3,7).setValues([
    ['Categoria','Nombre','URL','Imagen','Descripcion','Orden','Visible'],
    ['Aprendizaje-Servicio','Aprendizaje y Servicio Solidario','https://aprendizajeserviciosolidario.netlify.app/','','',1,'SI'],
    ['Aprendizaje-Servicio','Guías de Aprendizaje-Servicio','https://guiasaps.netlify.app/','','',2,'SI']
  ]);
  sh.setFrozenRows(1);sh.getRange(1,1,1,7).setFontWeight('bold').setBackground('#008eae').setFontColor('#ffffff');
  sh.autoResizeColumns(1,7);
  const rule=SpreadsheetApp.newDataValidation().requireValueInList(['SI','NO'],true).build();
  sh.getRange('G2:G').setDataValidation(rule);
}
