/** خادم Google Apps Script — الصق هذا الملف في مشروع مرتبط بجدول Google Sheets
 *  أوراق الجدول: Users (code | nid | name | tawgeeh | role) يملؤها المسؤول مسبقًا. باقي الأوراق تُنشأ تلقائيًا. */
const H = {login, savePlan, getPlan, submitModification, exportMine, uploadPlanXlsx};
const FIREBASE_PROJECT_ID = '';   // معرّف مشروع Firebase (Firestore بالوضع Native) — اتركه فارغًا لتعطيل Firebase
const ORG_FOLDER_ID = '';         // معرّف مجلد Drive الخاص بمنظومة المتابعة الإلكترونية (اختياري؛ وإلا يُنشأ مجلد تلقائيًا)
const DRIVE_FOLDER = 'خطط المتابعة - تصدير';
const SHEETS = {
  Plans: ['code','month','year','date','school','savedAt'],
  PlansAfter: ['code','month','year','date','school','modAt'],
  Modifications: ['code','month','year','date','school','action','modAt'],
  ModFiles: ['code','month','year','modAt','fileUrl'],
  PlanFiles: ['code','month','year','sentAt','fileUrl']
};
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doPost(e){
  try{
    const q = JSON.parse(e.postData.contents), f = H[q.action];
    if(!f) throw new Error('إجراء غير معروف');
    return out(Object.assign({ok:true}, f(q)));
  }catch(err){ return out({ok:false, error: String(err.message || err)}); }
}
function sheet(n){
  const ss = SpreadsheetApp.getActive(); let s = ss.getSheetByName(n);
  if(!s){ s = ss.insertSheet(n); s.appendRow(SHEETS[n]); s.getRange('D:D').setNumberFormat('@'); }
  return s;
}
function rows(n){
  const v = sheet(n).getDataRange().getValues(), h = v.shift();
  return v.map(r => { const o = {}; h.forEach((k,i) => o[k] = r[i]); return o; });
}
function user(q){
  const code = CacheService.getScriptCache().get('t_' + q.token);
  if(!code) throw new Error('انتهت الجلسة، سجّل الدخول من جديد');
  const u = rows('Users').find(r => String(r.code).trim() === code);
  if(!u) throw new Error('المستخدم غير موجود');
  return {code:String(u.code).trim(), name:u.name, tawgeeh:u.tawgeeh, role:u.role};
}
function login(q){
  const c = CacheService.getScriptCache(), k = 'f_' + q.code, n = +c.get(k) || 0;
  if(n >= 5) throw new Error('تم إيقاف المحاولات مؤقتًا (15 دقيقة)');
  const u = rows('Users').find(r => String(r.code).trim() === String(q.code).trim() && String(r.nid).trim() === String(q.nid).trim());
  if(!u){ c.put(k, String(n+1), 900); throw new Error('الرقم القومي والكود غير مطابقين للمسجّل'); }
  c.remove(k); const t = Utilities.getUuid(); c.put('t_' + t, String(u.code).trim(), 21600);
  return {token:t, profile:{code:String(u.code).trim(), name:u.name, tawgeeh:u.tawgeeh, role:u.role}};
}
function mine(n, u, q){ return rows(n).filter(r => String(r.code) === u.code && r.month === q.month && String(r.year) === String(q.year)); }
function replace(n, u, q, data){            // يحذف صفوف نفس (الكود/الشهر/السنة) ثم يكتب الجديد
  const s = sheet(n), v = s.getDataRange().getValues();
  for(let i = v.length - 1; i >= 1; i--) if(String(v[i][0]) === u.code && v[i][1] === q.month && String(v[i][2]) === String(q.year)) s.deleteRow(i + 1);
  if(data.length) s.getRange(s.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
}
function clean(q){ return q.rows.slice(0, 400).map(r => [String(r['التاريخ']).slice(0,10), String(r['المدرسة']).slice(0,120)]); }
function savePlan(q){
  const u = user(q), L = LockService.getScriptLock(); L.waitLock(20000);
  try{
    if(mine('Plans', u, q).length) throw new Error('الخطة الأصلية لهذا الشهر مسجلة مسبقًا؛ استخدم «رفع ملف التعديل».');
    const now = new Date();
    replace('Plans', u, q, clean(q).map(r => [u.code, q.month, q.year, "'" + r[0], r[1], now]));
    return {msg: 'تم تسجيل الخطة الأصلية في Google Sheets'};
  } finally { L.releaseLock(); }
}

/* ===== حفظ ملف التعديل المرفوع في Drive المنظمة + Firebase (Firestore) ===== */
function orgFolder(){
  if(ORG_FOLDER_ID) return DriveApp.getFolderById(ORG_FOLDER_ID);
  const n = 'منظومة المتابعة الإلكترونية - التعديلات', it = DriveApp.getFoldersByName(n);
  return it.hasNext() ? it.next() : DriveApp.createFolder(n);
}
function saveModFile(u, q, now){
  if(!q.fileB64) return '';
  const stamp = Utilities.formatDate(now, 'Africa/Cairo', 'yyyyMMdd_HHmm');
  const blob = Utilities.newBlob(Utilities.base64Decode(q.fileB64), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'تعديل_' + u.name + '_' + u.code + '_' + q.month + '_' + q.year + '_' + stamp + '.xlsx');
  return driveFolder(u.name + '_' + u.code, orgFolder()).createFile(blob).getUrl();
}
function fsAdd(u, q, now, log, after, url){
  if(!FIREBASE_PROJECT_ID) return false;
  const str = v => ({stringValue: String(v)}), row = o => ({mapValue: {fields: o}});
  const doc = {fields: {
    code: str(u.code), name: str(u.name), month: str(q.month), year: {integerValue: String(q.year)},
    modAt: {timestampValue: now.toISOString()}, driveUrl: str(url || ''),
    changes: {arrayValue: {values: log.map(l => row({date: str(strip(l[3])), school: str(l[4]), action: str(l[5])}))}},
    after: {arrayValue: {values: after.map(r => row({date: str(r[0]), school: str(r[1])}))}}
  }};
  UrlFetchApp.fetch('https://firestore.googleapis.com/v1/projects/' + FIREBASE_PROJECT_ID + '/databases/(default)/documents/modifications',
    {method: 'post', contentType: 'application/json', headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, payload: JSON.stringify(doc)});
  return true;
}
function submitModification(q){
  const u = user(q), L = LockService.getScriptLock(); L.waitLock(20000);
  try{
    const orig = mine('Plans', u, q).map(r => [String(r.date).replace(/^'/,''), r.school]);
    if(!orig.length) throw new Error('لا توجد خطة أصلية مسجلة لهذا الشهر');
    const after = clean(q), key = a => a.join('|'), now = new Date();
    const oS = orig.map(key), aS = after.map(key), log = [];
    after.forEach(r => { if(oS.indexOf(key(r)) < 0) log.push([u.code, q.month, q.year, "'" + r[0], r[1], 'إضافة', now]); });
    orig.forEach(r => { if(aS.indexOf(key(r)) < 0) log.push([u.code, q.month, q.year, "'" + r[0], r[1], 'حذف', now]); });
    if(!log.length) throw new Error('لا توجد فروق بين الملف المرفوع والخطة الأصلية');
    replace('PlansAfter', u, q, after.map(r => [u.code, q.month, q.year, "'" + r[0], r[1], now]));
    const s = sheet('Modifications'); s.getRange(s.getLastRow() + 1, 1, log.length, 7).setValues(log);
    let url = '', fsOk = false; const warn = [];
    try{ url = saveModFile(u, q, now); if(url) sheet('ModFiles').appendRow([u.code, q.month, q.year, now.toISOString(), url]); }catch(e){ warn.push('Drive: ' + e.message); }
    try{ fsOk = fsAdd(u, q, now, log, after, url); }catch(e){ warn.push('Firebase: ' + e.message); }
    return {msg: 'تم تسجيل التعديل (' + log.length + ' تغيير) في Google Sheets' + (url ? ' وحفظ الملف في Drive' : '') + (fsOk ? ' وFirebase' : '') + (warn.length ? ' — تنبيه: ' + warn.join(' | ') : '')};
  } finally { L.releaseLock(); }
}
function getPlan(q){
  const u = user(q), f = r => ({'التاريخ': String(r.date).replace(/^'/,''), 'المدرسة': r.school});
  const a = mine('PlansAfter', u, q);
  return {profile: u, original: mine('Plans', u, q).map(f), after: a.length ? a.map(f) : null, modAt: a.length ? a[0].modAt : null};
}

/* ===== التصدير إلى Google Drive (ملف XLSX لكل موجه/متابع لكل شهر) ===== */
function strip(d){ return String(d).replace(/^'/, ''); }
function buildXlsx(u, month, year){
  const tmp = SpreadsheetApp.create('tmp_' + u.code), q = {month: month, year: year};
  try{
    const put = (name, head, data) => {
      const s = tmp.insertSheet(name); s.getRange('A:A').setNumberFormat('@');
      s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      if(data.length) s.getRange(2, 1, data.length, head.length).setValues(data);
      s.setRightToLeft(true);
    };
    put('الخطة الأصلية', ['التاريخ','المدرسة'], mine('Plans', u, q).map(r => [strip(r.date), r.school]));
    put('بعد التعديل', ['التاريخ','المدرسة'], mine('PlansAfter', u, q).map(r => [strip(r.date), r.school]));
    put('سجل التعديلات', ['التاريخ','المدرسة','الإجراء','وقت التعديل'],
      mine('Modifications', u, q).map(r => [strip(r.date), r.school, r.action, String(r.modAt)]));
    tmp.deleteSheet(tmp.getSheets()[0]); SpreadsheetApp.flush();
    const res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx',
      {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}});
    return res.getBlob().setName(u.name + '_' + u.code + '_' + month + '_' + year + '.xlsx');
  } finally { DriveApp.getFileById(tmp.getId()).setTrashed(true); }
}
function driveFolder(name, parent){
  const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name);
}
function saveToDrive(u, blob){
  const root = DriveApp.getFoldersByName(DRIVE_FOLDER), top = root.hasNext() ? root.next() : DriveApp.createFolder(DRIVE_FOLDER);
  const dir = driveFolder(u.name + '_' + u.code, top), old = dir.getFilesByName(blob.getName());
  while(old.hasNext()) old.next().setTrashed(true);   // استبدال نسخة الشهر القديمة
  return dir.createFile(blob).getUrl();
}
/** زر المستخدم: يحفظ ملفه في Drive ويعيده للتنزيل */
function exportMine(q){
  const u = user(q), blob = buildXlsx(u, q.month, q.year);
  saveToDrive(u, blob);
  return {msg: 'تم تصدير ملفك إلى Google Drive', name: blob.getName(), b64: Utilities.base64Encode(blob.getBytes())};
}
/** للمسؤول: شغّلها من المحرر أو بمُشغّل يومي — تصدّر كل الموجهين */
function exportAll(){
  const users = {}; rows('Users').forEach(r => users[String(r.code).trim()] = {code:String(r.code).trim(), name:r.name, tawgeeh:r.tawgeeh, role:r.role});
  const seen = {};
  rows('Plans').forEach(r => {
    const k = r.code + '|' + r.month + '|' + r.year, u = users[String(r.code)];
    if(seen[k] || !u) return; seen[k] = 1;
    saveToDrive(u, buildXlsx(u, r.month, String(r.year)));
  });
}
function setupDailyExport(){   // شغّلها مرة واحدة فقط
  ScriptApp.newTrigger('exportAll').timeBased().everyDays(1).atHour(2).create();
}

/* ===== استلام ملف الخطة الشهرية XLSX من الصفحة وحفظه في Drive ===== */
function plansFolder(){
  if(ORG_FOLDER_ID) return driveFolder('الخطط', DriveApp.getFolderById(ORG_FOLDER_ID));
  const n = 'منظومة المتابعة الإلكترونية - الخطط', it = DriveApp.getFoldersByName(n);
  return it.hasNext() ? it.next() : DriveApp.createFolder(n);
}
function fsAddPlan(u, q, rows, now, url){
  if(!FIREBASE_PROJECT_ID) return false;
  const str = v => ({stringValue: String(v)}), row = o => ({mapValue: {fields: o}});
  const doc = {fields: {
    code: str(u.code), name: str(u.name), month: str(q.month), year: {integerValue: String(q.year)},
    sentAt: {timestampValue: now.toISOString()}, driveUrl: str(url || ''),
    visits: {arrayValue: {values: rows.map(r => row({date: str(r[0]), school: str(r[1])}))}}
  }};
  UrlFetchApp.fetch('https://firestore.googleapis.com/v1/projects/' + FIREBASE_PROJECT_ID + '/databases/(default)/documents/plans',
    {method: 'post', contentType: 'application/json', headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}, payload: JSON.stringify(doc)});
  return true;
}
/** زر «إرسال الخطة الشهرية»: ملف XLSX إلى Drive + سجل في Firebase + صفوف في Google Sheets */
function uploadPlanXlsx(q){
  const u = user(q);
  if(!q.fileB64 || q.fileB64.length > 7e6) throw new Error('الملف غير صالح أو أكبر من 5 ميجابايت');
  const bytes = Utilities.base64Decode(q.fileB64);
  if(bytes[0] !== 80 || bytes[1] !== 75) throw new Error('الملف ليس بصيغة XLSX');   // توقيع ZIP "PK"
  const clean = t => String(t || '').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 40), now = new Date(), done = [], warn = [];
  const name = clean(u.name) + '_' + clean(q.month) + '_' + clean(q.year) + '.xlsx';
  let url = '';
  try{
    const blob = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', name);
    const dir = driveFolder(u.name + '_' + u.code, plansFolder()), old = dir.getFilesByName(name);
    while(old.hasNext()) old.next().setTrashed(true);   // استبدال ملف نفس الشهر
    url = dir.createFile(blob).getUrl(); done.push('Google Drive');
    sheet('PlanFiles').appendRow([u.code, q.month, q.year, now.toISOString(), url]);
  }catch(e){ warn.push('Drive: ' + e.message); }
  const data = clean2(q);
  const L = LockService.getScriptLock(); L.waitLock(20000);
  try{
    if(data.length){
      if(mine('Plans', u, q).length) warn.push('الخطة الأصلية لهذا الشهر مسجلة مسبقًا في Sheets (لم تُستبدل؛ استخدم رفع التعديل)');
      else { replace('Plans', u, q, data.map(r => [u.code, q.month, q.year, "'" + r[0], r[1], now])); done.push('Google Sheets'); }
    }
  }catch(e){ warn.push('Sheets: ' + e.message); } finally { L.releaseLock(); }
  try{ if(fsAddPlan(u, q, data, now, url)) done.push('Firebase'); }catch(e){ warn.push('Firebase: ' + e.message); }
  if(!done.length) throw new Error(warn.join(' | ') || 'لم يُحفظ شيء');
  return {msg: 'تم إرسال الخطة إلى: ' + done.join(' + ') + (warn.length ? ' — تنبيه: ' + warn.join(' | ') : '')};
}
function clean2(q){ return (q.rows || []).slice(0, 400).map(r => [String(r['التاريخ']).slice(0,10), String(r['المدرسة']).slice(0,120)]); }
