/** Biblioteca UNIR: carpetas de primer nivel = estantes/categorías; subcarpetas = secciones. */
const ROOT_FOLDER_ID = '1YTT9VDpOp-y0ZvyRkE8QKyjy5IsFCY8j';
function doGet(e) {
  let result;
  try {
    const root = DriveApp.getFolderById(ROOT_FOLDER_ID), books=[], folders=[], categories=new Set();
    const seen=new Set();
    function scan(folder, parts, depth) {
      if(depth>20 || seen.has(folder.getId()))return;
      seen.add(folder.getId());
      const category=parts[0]||'General';
      const path=parts.join(' / ')||'General';
      categories.add(category);
      folders.push({name:path,category:category,depth:parts.length});
      const files=folder.getFiles();
      while(files.hasNext()){
        const file=files.next();
        if(file.isTrashed() || file.getMimeType()!==MimeType.PDF)continue;
        const id=file.getId();
        books.push({id:id,title:file.getName().replace(/\.pdf$/i,''),folder:path,category:category,
          added:file.getDateCreated().toISOString(),modified:file.getLastUpdated().toISOString(),
          url:'https://drive.google.com/file/d/'+id+'/preview'});
      }
      const sub=folder.getFolders();
      while(sub.hasNext()) {const child=sub.next();if(!child.isTrashed())scan(child,parts.concat(child.getName()),depth+1)}
    }
    scan(root,[],0);
    result={ok:true,books:books,folders:folders,categories:[...categories].sort(),updated:new Date().toISOString()};
  }catch(err){result={ok:false,error:String(err)}}
  const callback=e&&e.parameter&&e.parameter.callback;
  if(callback && /^[a-zA-Z_$][\w$]*$/.test(callback))
    return ContentService.createTextOutput(callback+'('+JSON.stringify(result)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
