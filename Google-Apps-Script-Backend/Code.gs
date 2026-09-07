const CONFIG = {
  MASTER_USER_ID: 'master',
  MASTER_INITIAL_PASSWORD: 'Panchayat@2026!',
  MIN_PHOTO_DISTANCE_METERS: 10,
  SHEET_NAME: 'WorkGPS',
  USER_SHEET_NAME: 'Users',
  SESSION_SHEET_NAME: 'Sessions',
  DRIVE_FOLDER_NAME: 'Work Photo Database',
  MAX_PHOTOS: 4,
  SESSION_HOURS: 24
};

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  try {
    if (action === 'login') return jsonOutput_(loginUser(e.parameter.userId, e.parameter.password));
    if (action === 'me') return jsonOutput_(getCurrentUser_(e.parameter.token));
    if (action === 'works') return jsonOutput_(getWorks_(e.parameter.token, e.parameter.sansadNo || ''));
    if (action === 'sansads') return jsonOutput_(getSansads_(e.parameter.token));
    if (action === 'users') return jsonOutput_(getUsers_(e.parameter.token));
    if (action === 'permissions') return jsonOutput_(getPermissions_(e.parameter.token, e.parameter.userId || ''));
    if (action === 'report') return jsonOutput_(getReport_(e.parameter.token));
    if (action === 'database') return jsonOutput_(getDatabase_(e.parameter.token));
    if (action === 'userDatabase') return jsonOutput_(getDatabase_(e.parameter.token));
    if (action === 'photoDatabase') return jsonOutput_(getPhotoDatabase_(e.parameter.token));
    if (action === 'photoPreview') return jsonOutput_(getPhotoPreview_(e.parameter.token, e.parameter.photoId || ''));
    if (action === 'activity') return jsonOutput_(getActivity_(e.parameter.token));
    if (action === 'settings') return jsonOutput_(getSystemSettings_(e.parameter.token));
    if (action === 'health') return jsonOutput_({success:true,service:'Panchayat Work GPS',time:new Date().toISOString()});
  } catch (error) {
    return jsonOutput_({success:false, message:error.message});
  }
  return HtmlService.createHtmlOutputFromFile('index').setTitle('Panchayat Work GPS');
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = body.action;
    const token = body.token || '';
    const d = body.data || {};

    if (action === 'login') return jsonOutput_(loginUser(d.userId, d.password));
    if (action === 'savePhoto') return jsonOutput_(savePhoto(d, token));
    if (action === 'addWork') return jsonOutput_(addWork(d, token));
    if (action === 'addUser') return jsonOutput_(addUser(d, token));
    if (action === 'updateUser') return jsonOutput_(updateUser_(d, token));
    if (action === 'deleteUser') return jsonOutput_(deleteUser_(d, token));
    if (action === 'updatePermissions') return jsonOutput_(updatePermissions(d, token));
    if (action === 'changeAdminPassword') return jsonOutput_(changeAdminPassword_(d, token));
    if (action === 'logout') return jsonOutput_(logout_(token));
    if (action === 'updateWork') return jsonOutput_(updateWork_(d, token));
    if (action === 'deleteWork') return jsonOutput_(deleteWork_(d, token));
    if (action === 'deleteWorks') return jsonOutput_(deleteWorks_(d, token));
    if (action === 'deleteAllWorks') return jsonOutput_(deleteAllWorks_(token));
    if (action === 'uploadExcel') return jsonOutput_(uploadExcel_(d, token));
    if (action === 'updatePhoto') return jsonOutput_(updatePhoto_(d, token));
    if (action === 'saveSettings') return jsonOutput_(saveSystemSettings_(d, token));
    if (action === 'resetSystem') return jsonOutput_(resetSystem_(token));
    if (action === 'changeMasterPassword') return jsonOutput_(changeMasterPassword_(d, token));
    return jsonOutput_({success:false, message:'Unknown API action'});
  } catch (error) {
    return jsonOutput_({success:false, message:error.message});
  }
}

function jsonOutput_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Google Sheet থেকে Extensions → Apps Script খুলুন।');
  return ss;
}

function getWorkSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  return sheet;
}

function setupSheet() {
  const sheet = getWorkSheet_();
  const headers = ['Sl No','Sansad No','Work Code','Work Name','Location'];
  for (let photo=1; photo<=CONFIG.MAX_PHOTOS; photo++) { headers.push('Photo '+photo+' Latitude'); headers.push('Photo '+photo+' Longitude'); }
  headers.push('Date'); headers.push('User');
  const oldLastColumn=sheet.getLastColumn();
  if(oldLastColumn>=26 && sheet.getRange(1,25).getValue()==='Date'){ const lr=sheet.getLastRow(); if(lr>1){ const old=sheet.getRange(2,25,lr-1,2).getValues(); sheet.getRange(2,14,lr-1,2).setValues(old); } }
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if(sheet.getMaxColumns()>headers.length) sheet.deleteColumns(headers.length+1,sheet.getMaxColumns()-headers.length);
  sheet.setFrozenRows(1); return true;
}

function findWork_(workCode) {
  const sheet = getWorkSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const data = sheet.getRange(2,1,lastRow-1,5).getValues();
  const wanted = String(workCode || '').trim().toLowerCase();
  for (let i=0;i<data.length;i++) {
    if (String(data[i][2] || '').trim().toLowerCase() === wanted) {
      return {sheet:sheet,row:i+2,slNo:data[i][0],sansadNo:data[i][1],workCode:data[i][2],workName:data[i][3],location:data[i][4]};
    }
  }
  return null;
}

function isWorkCompleteRow_(row){
  for(let photo=1;photo<=CONFIG.MAX_PHOTOS;photo++){
    const lat=row[5+(photo-1)*2], lng=row[6+(photo-1)*2];
    if(lat==='' || lat===null || lng==='' || lng===null) return false;
  }
  return true;
}

function getWorks_(token, sansadNo) {
  const user = requireSession_(token);
  const sheet = getWorkSheet_();
  if (sheet.getRange(1,1).getValue() !== 'Sl No') setupSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2,1,lastRow-1,5).getValues();
  const allowed = getAllowedSansads_(user);
  const wantedSansad = String(sansadNo || '').trim();
  const isAdmin = String(user.role||'').toLowerCase()==='admin';
  return data.filter(function(row){
    if (row[2] === '') return false;
    const s = String(row[1] || '').trim();
    if (wantedSansad && normalizeSansad_(s) !== normalizeSansad_(wantedSansad)) return false;
    if (!isSansadAllowed_(s, allowed)) return false;
    return true;
  }).map(function(row){
    let photoCount=0;
    for(let i=5;i<13;i+=2){ if(row[i]!=='' && row[i+1]!=='' && row[i]!=null && row[i+1]!=null) photoCount++; }
    return {slNo:String(row[0]||''), sansadNo:String(row[1]||''), code:String(row[2]||''), name:String(row[3]||''), location:String(row[4]||''), photoCount:photoCount, completed:photoCount>=CONFIG.MAX_PHOTOS};
  });
}

function getSansads_(token) {
  const user = requireSession_(token);
  const sheet = getWorkSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2,1,lastRow-1,5).getValues();
  const allowed = getAllowedSansads_(user);
  const map = {};
  data.forEach(function(row){
    const s = String(row[1] || '').trim();
    if (s && isSansadAllowed_(s, allowed)) map[s] = true;
  });
  return Object.keys(map).sort(function(a,b){return a.localeCompare(b,undefined,{numeric:true});});
}

function addWork(d, token) {
  const user = requirePermission_(token, 'addWork');
  const workCode = String(d.workCode || '').trim();
  const workName = String(d.workName || '').trim();
  const sansadNo = String(d.sansadNo || '').trim();
  if (!workCode) throw new Error('Work Code দিন।');
  if (!workName) throw new Error('Work Name দিন।');
  if (!sansadNo) throw new Error('Sansad No দিন।');
  if (findWork_(workCode)) throw new Error('এই Work Code আগে থেকেই আছে।');
  const allowed = getAllowedSansads_(user);
  if (!isSansadAllowed_(sansadNo, allowed)) throw new Error('এই Sansad-এর Work Add করার permission নেই।');
  getWorkSheet_().appendRow([d.slNo || '', sansadNo, workCode, workName, d.location || '']);
  return {success:true,message:'Work successfully added'};
}

function nextPhoto_(sheet,row) {
  for (let photo=1;photo<=CONFIG.MAX_PHOTOS;photo++) {
    const latitudeColumn=6+((photo-1)*2);
    const longitudeColumn=latitudeColumn+1;
    const latitude=sheet.getRange(row,latitudeColumn).getValue();
    const longitude=sheet.getRange(row,longitudeColumn).getValue();
    if (latitude === '' && longitude === '') return photo;
  }
  return null;
}

function getMainDriveFolder_() {
  const folders=DriveApp.getFoldersByName(CONFIG.DRIVE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(CONFIG.DRIVE_FOLDER_NAME);
}

function savePhoto(data, token) {
  const user=requirePermission_(token,'photoGPS');
  const incomingPhotoId=String(data&&data.photoId||'').trim();
  if(incomingPhotoId){
    const ps=setupPhotoDatabase_(), plr=ps.getLastRow();
    if(plr>=2){
      const ids=ps.getRange(2,1,plr-1,1).getValues();
      for(let i=0;i<ids.length;i++) if(String(ids[i][0]||'')===incomingPhotoId){
        const row=ps.getRange(i+2,1,1,10).getValues()[0];
        return {success:true,duplicate:true,photoId:incomingPhotoId,photoNo:row[2],latitude:Number(row[6]),longitude:Number(row[7]),fileUrl:String(row[4]||'')};
      }
    }
  }
  if(!data||!data.workCode)throw new Error('Work Code is required.');
  if(data.latitude===undefined||data.longitude===undefined)throw new Error('GPS coordinates are required.');
  const latitude=Number(data.latitude),longitude=Number(data.longitude);
  if(!isFinite(latitude)||!isFinite(longitude)||(latitude===0&&longitude===0))throw new Error('Invalid GPS coordinates.');
  const work=findWork_(data.workCode); if(!work)throw new Error('Work Code not found.');
  if(!isSansadAllowed_(String(work.sansadNo||''),getAllowedSansads_(user)))throw new Error('Photo/GPS permission is not available for this Sansad.');
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{ const photoNo=nextPhoto_(work.sheet,work.row); if(!photoNo)throw new Error('This work already has 4 photos.');
    const col=6+((photoNo-1)*2); work.sheet.getRange(work.row,col,1,2).setValues([[latitude,longitude]]);
    work.sheet.getRange(work.row,14).setValue(new Date()); work.sheet.getRange(work.row,15).setValue(user.userName);
    let fileUrl=''; let fileId=''; let fileName=''; if(data.photoBase64){ const main=getMainDriveFolder_(); const safeCode=String(work.workCode).replace(/[\\/:*?"<>|]/g,'_'); const safeName=String(work.workName||'Work').replace(/[\\/:*?"<>|]/g,'_'); const folderName=safeCode+'_'+safeName; const fs=main.getFoldersByName(folderName); const folder=fs.hasNext()?fs.next():main.createFolder(folderName); const bytes=Utilities.base64Decode(String(data.photoBase64).split(',').pop()); fileName=safeCode+'_Photo_'+String(photoNo).padStart(2,'0')+'.jpg'; const file=folder.createFile(Utilities.newBlob(bytes,'image/jpeg',fileName)); fileUrl=file.getUrl(); fileId=file.getId(); }
    if(fileUrl){ const ps=setupPhotoDatabase_(); const photoId=incomingPhotoId||Utilities.getUuid(); ps.appendRow([photoId,String(work.workCode),photoNo,fileName,fileUrl,fileId,latitude,longitude,new Date(),user.userName]); }
    return {success:true,photoNo:photoNo,latitude:latitude,longitude:longitude,fileUrl:fileUrl};
  }finally{lock.releaseLock();}
}

function getUserSheet_(){
  const ss=getSpreadsheet_();
  let sheet=ss.getSheetByName(CONFIG.USER_SHEET_NAME);
  if(!sheet) sheet=ss.insertSheet(CONFIG.USER_SHEET_NAME);
  return sheet;
}

function getPhotoPreview_(token, photoId){
  requireSession_(token);
  if(!photoId) throw new Error('Photo ID is required.');
  const sheet=setupPhotoDatabase_();
  const lr=sheet.getLastRow();
  if(lr<2) throw new Error('Photo not found.');
  const rows=sheet.getRange(2,1,lr-1,10).getValues();
  for(let i=0;i<rows.length;i++){
    if(String(rows[i][0]||'')===String(photoId)){
      const work=findWork_(String(rows[i][1]||''));
      if(!work || !isSansadAllowed_(String(work.sansadNo||''),getAllowedSansads_(requireSession_(token)))) throw new Error('Permission denied.');
      const fileId=String(rows[i][5]||'');
      if(!fileId) throw new Error('Photo preview is unavailable.');
      const file=DriveApp.getFileById(fileId);
      const blob=file.getBlob();
      return {success:true,photoId:photoId,mimeType:blob.getContentType()||'image/jpeg',base64:Utilities.base64Encode(blob.getBytes()),fileName:file.getName()};
    }
  }
  throw new Error('Photo not found.');
}

function setupPhotoDatabase_(){
  const ss=getSpreadsheet_();
  let sheet=ss.getSheetByName('Photo Database');
  if(!sheet) sheet=ss.insertSheet('Photo Database');
  const headers=['Photo ID','Work Code','Photo No','File Name','Drive URL','File ID','Latitude','Longitude','Date','User'];
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function getPhotoDatabase_(token){
  const user=requireSession_(token);
  const sheet=setupPhotoDatabase_();
  const lr=sheet.getLastRow();
  if(lr<2) return [];
  const rows=sheet.getRange(2,1,lr-1,10).getValues();
  const allowed=getAllowedSansads_(user);
  return rows.filter(r=>r[1]).filter(r=>{ const w=findWork_(String(r[1]||'')); return w && isSansadAllowed_(String(w.sansadNo||''),allowed); }).map(r=>({photoId:String(r[0]||''),workCode:String(r[1]||''),photoNo:String(r[2]||''),fileName:String(r[3]||''),fileUrl:String(r[4]||''),fileId:String(r[5]||''),latitude:String(r[6]??''),longitude:String(r[7]??''),date:r[8],user:String(r[9]||'')}));
}

function setupUsers(){
  const sheet=getUserSheet_();
  const headers=['User ID','User Name','Password Hash','Role','Active','Add Work','Photo GPS','Report','User Management','Sansad Access'];
  const current = sheet.getLastColumn() >= headers.length
    ? sheet.getRange(1,1,1,headers.length).getValues()[0]
    : [];
  let same = current.length === headers.length;
  if(same) for(let i=0;i<headers.length;i++) if(String(current[i]||'') !== headers[i]) { same=false; break; }
  if(!same) sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  if(sheet.getFrozenRows() !== 1) sheet.setFrozenRows(1);
  return true;
}

function userCacheKey_(userId){
  return 'pwg_user_' + String(userId||'').trim().toLowerCase().replace(/[^a-z0-9_.-]/g,'_').slice(0,80);
}
function clearUserCache_(userId){
  if(userId) CacheService.getScriptCache().remove(userCacheKey_(userId));
}
function clearAllUserCaches_(){
  // CacheService does not expose key enumeration; login cache entries naturally expire.
}
function recordFromRow_(row,rowNumber){
  const p=ensureUserRowShape_(row);
  return {
    userId:String(row[0]||''), userName:String(row[1]||''),
    role:String(row[3]||'User'), active:String(row[4]||'YES'),
    permissions:p, passwordHash:String(row[2]||''), row:rowNumber
  };
}
function getCachedUser_(userId){
  const raw=CacheService.getScriptCache().get(userCacheKey_(userId));
  if(!raw) return null;
  try{return JSON.parse(raw)}catch(e){return null}
}
function cacheUser_(u){
  CacheService.getScriptCache().put(userCacheKey_(u.userId),JSON.stringify(u),600);
  return u;
}
function findUserFast_(userId){
  const wanted=String(userId||'').trim().toLowerCase();
  if(!wanted) throw new Error('User পাওয়া যায়নি।');
  const cached=getCachedUser_(wanted);
  if(cached) return cached;
  const sheet=getUserSheet_();
  const lastRow=sheet.getLastRow();
  if(lastRow<2) throw new Error('User পাওয়া যায়নি।');
  const data=sheet.getRange(2,1,lastRow-1,10).getValues();
  for(let i=0;i<data.length;i++){
    if(String(data[i][0]||'').trim().toLowerCase()===wanted) return cacheUser_(recordFromRow_(data[i],i+2));
  }
  throw new Error('User ID অথবা Password ভুল।');
}

function hashPassword_(password){
  const raw=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(password));
  return raw.map(function(byte){return (byte<0?byte+256:byte).toString(16).padStart(2,'0');}).join('');
}

function ensureUserRowShape_(row){
  const role=String(row[3]||'User').trim().toLowerCase();
  const active=String(row[4]||'YES').toUpperCase();
  let addWork=String(row[5]||'').toUpperCase();
  let photoGPS=String(row[6]||'').toUpperCase();
  let report=String(row[7]||'').toUpperCase();
  let userMgmt=String(row[8]||'').toUpperCase();
  let sansads=String(row[9]||'').trim();
  // Master controls Admin permissions. Existing Admin rows with blank permissions
  // are upgraded to the legacy full-access defaults, but explicit NO values are preserved.
  if(role==='admin') {
    if(!addWork) addWork='YES';
    if(!photoGPS) photoGPS='YES';
    if(!report) report='YES';
    if(!userMgmt) userMgmt='YES';
    if(!sansads) sansads='ALL';
  } else {
    if(!addWork) addWork='NO';
    if(!photoGPS) photoGPS='YES';
    if(!report) report='NO';
    if(!userMgmt) userMgmt='NO';
    if(!sansads) sansads='ALL';
  }
  return {addWork:addWork==='YES',photoGPS:photoGPS==='YES',report:report==='YES',userManagement:userMgmt==='YES',sansadAccess:sansads};
}

function addUser(d, token){
  requirePermission_(token,'userManagement');
  if(!d.userId || !d.userName || !d.password) throw new Error('User ID, User Name এবং Password প্রয়োজন।');
  setupUsers();
  const sheet=getUserSheet_();
  const lastRow=sheet.getLastRow();
  const data=lastRow>=2?sheet.getRange(2,1,lastRow-1,10).getValues():[];
  const wanted=String(d.userId).trim().toLowerCase();
  for(let i=0;i<data.length;i++) if(String(data[i][0]).trim().toLowerCase()===wanted) throw new Error('এই User ID ইতিমধ্যে আছে।');
  const role=d.role||'User';
  const perms=d.permissions||{};
  if(String(role).toLowerCase()==='admin' && !isMaster_(requireSession_(token))) throw new Error('শুধু Master Login থেকে Admin account তৈরি করা যাবে।');
  const addWork=perms.addWork?'YES':'NO';
  const photoGPS=perms.photoGPS===false?'NO':'YES';
  const report=perms.report?'YES':'NO';
  const userMgmt=perms.userManagement?'YES':'NO';
  const sansadAccess=String(perms.sansadAccess||'ALL').trim()||'ALL';
  sheet.appendRow([d.userId,d.userName,hashPassword_(d.password),role,'YES',addWork,photoGPS,report,userMgmt,sansadAccess]);
  clearUserCache_(d.userId);
  return {success:true,message:'User added'};
}

function getUsers_(token){
  requirePermission_(token,'userManagement');
  setupUsers();
  const sheet=getUserSheet_();
  const lastRow=sheet.getLastRow();
  if(lastRow<2) return [];
  const data=sheet.getRange(2,1,lastRow-1,10).getValues();
  return data.map(function(row){
    const p=ensureUserRowShape_(row);
    return {userId:String(row[0]||''),userName:String(row[1]||''),role:String(row[3]||''),active:String(row[4]||''),addWork:p.addWork,photoGPS:p.photoGPS,report:p.report,userManagement:p.userManagement,sansadAccess:p.sansadAccess};
  });
}


function updateUser_(d, token){
  requirePermission_(token,'userManagement');
  if(!d.userId || !d.userName) throw new Error('User ID এবং User Name প্রয়োজন।');
  const target=findUser_(d.userId);
  const session=requireSession_(token);
  if(String(target.role).toLowerCase()==='admin' && String(target.userId).toLowerCase()!==String(session.userId).toLowerCase()){
    throw new Error('অন্য Admin account edit করা যাবে না।');
  }
  const role=String(d.role||'User');
  const p=d.permissions||{};
  if(role.toLowerCase()==='admin' && !isMaster_(session)) throw new Error('Admin account permissions can be changed only by Master Login.');
  const addWork=p.addWork?'YES':'NO';
  const photoGPS=p.photoGPS===false?'NO':'YES';
  const report=p.report?'YES':'NO';
  const userMgmt=p.userManagement?'YES':'NO';
  const sansadAccess=String(p.sansadAccess||'ALL').trim()||'ALL';
  const sheet=getUserSheet_();
  sheet.getRange(target.row,1,1,10).setValues([[
    String(d.userId).trim(),String(d.userName).trim(),
    d.password ? hashPassword_(String(d.password)) : target.passwordHash,
    role,String(target.active||'YES'),addWork,photoGPS,report,userMgmt,sansadAccess
  ]]);
  clearUserCache_(d.userId);
  return {success:true,message:'User updated'};
}

function deleteUser_(d, token){
  requirePermission_(token,'userManagement');
  if(!d.userId) throw new Error('User ID প্রয়োজন।');
  const target=findUser_(d.userId);
  const session=requireSession_(token);
  if(String(target.userId).toLowerCase()===String(session.userId).toLowerCase()) throw new Error('নিজের account delete করা যাবে না।');
  if(String(target.role).toLowerCase()==='admin') throw new Error('Admin account delete করা যাবে না।');
  getUserSheet_().deleteRow(target.row);
  clearUserCache_(d.userId);
  return {success:true,message:'User deleted'};
}

function getPermissions_(token,targetUserId){
  const session=requireSession_(token);
  const target=targetUserId?findUser_(targetUserId):session;
  if(target.userId!==session.userId && !session.permissions.userManagement) throw new Error('Permission denied');
  return {success:true,userId:target.userId,userName:target.userName,role:target.role,permissions:target.permissions};
}

function updatePermissions(d, token){
  requirePermission_(token,'userManagement');
  if(!d.userId) throw new Error('User ID প্রয়োজন।');
  const sheet=getUserSheet_();
  const lastRow=sheet.getLastRow();
  if(lastRow<2) throw new Error('User পাওয়া যায়নি।');
  const data=sheet.getRange(2,1,lastRow-1,10).getValues();
  const wanted=String(d.userId).trim().toLowerCase();
  for(let i=0;i<data.length;i++){
    if(String(data[i][0]).trim().toLowerCase()===wanted){
      const role=String(data[i][3]||'User').trim().toLowerCase();
      if(role==='admin') throw new Error('Admin-এর permission পরিবর্তন করা যাবে না।');
      const p=d.permissions||{};
      sheet.getRange(i+2,6,1,5).setValues([[p.addWork?'YES':'NO',p.photoGPS===false?'NO':'YES',p.report?'YES':'NO',p.userManagement?'YES':'NO',String(p.sansadAccess||'ALL').trim()||'ALL']]);
      clearUserCache_(d.userId);
      return {success:true,message:'Permission updated'};
    }
  }
  throw new Error('User পাওয়া যায়নি।');
}

function changeAdminPassword_(d, token){
  const admin = requireSession_(token);
  if (String(admin.role || '').toLowerCase() !== 'admin') {
    throw new Error('শুধু Admin নিজের password পরিবর্তন করতে পারবেন।');
  }

  const currentPassword = String(d.currentPassword || '');
  const newPassword = String(d.newPassword || '');
  const confirmPassword = String(d.confirmPassword || '');

  if (!currentPassword || !newPassword || !confirmPassword) {
    throw new Error('Current Password, New Password এবং Confirm Password দিন।');
  }
  if (newPassword.length < 6) {
    throw new Error('নতুন Password কমপক্ষে 6 অক্ষরের হতে হবে।');
  }
  if (newPassword !== confirmPassword) {
    throw new Error('New Password এবং Confirm Password মিলছে না।');
  }
  if (hashPassword_(currentPassword) !== findUser_(admin.userId).passwordHash) {
    throw new Error('Current Password ভুল।');
  }

  const sheet = getUserSheet_();
  const row = findUser_(admin.userId).row;
  sheet.getRange(row, 3).setValue(hashPassword_(newPassword));

  clearUserCache_(admin.userId);
  return {success:true, message:'Admin Password successfully changed।'};
}

function findUser_(userId){
  return findUserFast_(userId);
}

function setupMaster_(){
  const props=PropertiesService.getScriptProperties();
  if(!props.getProperty('PWG_MASTER_PASSWORD_HASH')) props.setProperty('PWG_MASTER_PASSWORD_HASH',hashPassword_(CONFIG.MASTER_INITIAL_PASSWORD));
  if(!props.getProperty('PWG_MASTER_USER_ID')) props.setProperty('PWG_MASTER_USER_ID',CONFIG.MASTER_USER_ID);
}
function isMasterId_(id){setupMaster_();return String(id||'').trim().toLowerCase()===String(PropertiesService.getScriptProperties().getProperty('PWG_MASTER_USER_ID')||CONFIG.MASTER_USER_ID).toLowerCase();}
function loginMaster_(userId,password){
  setupMaster_();
  const props=PropertiesService.getScriptProperties();
  if(!isMasterId_(userId) || hashPassword_(password)!==props.getProperty('PWG_MASTER_PASSWORD_HASH')) throw new Error('User ID অথবা Password ভুল।');
  const token=Utilities.getUuid(), created=new Date(), expires=new Date(created.getTime()+CONFIG.SESSION_HOURS*60*60*1000);
  const u={userId:String(userId),userName:'Master',role:'Master',permissions:{addWork:true,photoGPS:true,report:true,userManagement:true,theme:true,settings:true,systemReset:true,logo:true,sansadAccess:'ALL'}};
  const payload={userId:u.userId,userName:u.userName,role:u.role,permissions:u.permissions,expires:expires.getTime(),user:u};
  CacheService.getScriptCache().put('pwg_session_'+token,JSON.stringify(payload),21600);
  PropertiesService.getScriptProperties().setProperty('pwg_session_'+token,JSON.stringify(payload));
  return {success:true,token,userId:u.userId,userName:u.userName,role:u.role,permissions:u.permissions,master:true};
}
function isMaster_(user){return String(user&&user.role||'').toLowerCase()==='master';}
function requireMaster_(token){const u=requireSession_(token);if(!isMaster_(u))throw new Error('শুধু Master Login থেকে এই কাজ করা যাবে।');return u;}
function getSystemSettings_(token){requireMaster_(token);const p=PropertiesService.getScriptProperties();let raw=p.getProperty('PWG_SYSTEM_SETTINGS')||'{}';try{return {success:true,settings:JSON.parse(raw)}}catch(e){return {success:true,settings:{}}}}
function saveSystemSettings_(d,token){requireMaster_(token);const safe={theme:d.theme||'A',themeConfig:d.themeConfig||{},systemName:String(d.systemName||'Panchayat Work GPS').slice(0,120),logoDataUrl:String(d.logoDataUrl||'').slice(0,120000)};PropertiesService.getScriptProperties().setProperty('PWG_SYSTEM_SETTINGS',JSON.stringify(safe));return {success:true,settings:safe,message:'System theme and branding saved.'};}
function changeMasterPassword_(d,token){requireMaster_(token);const cur=String(d.currentPassword||''),next=String(d.newPassword||''),confirm=String(d.confirmPassword||'');setupMaster_();const props=PropertiesService.getScriptProperties();if(!cur||!next||!confirm)throw new Error('Current, New এবং Confirm password দিন।');if(hashPassword_(cur)!==props.getProperty('PWG_MASTER_PASSWORD_HASH'))throw new Error('Current password ভুল।');if(next.length<8)throw new Error('Master password কমপক্ষে 8 অক্ষরের হতে হবে।');if(next!==confirm)throw new Error('New password এবং Confirm password মিলছে না।');props.setProperty('PWG_MASTER_PASSWORD_HASH',hashPassword_(next));return {success:true,message:'Master password changed.'};}
function resetSystem_(token){requireMaster_(token);const ss=getSpreadsheet_();['WorkGPS','Photo Database'].forEach(n=>{const sh=ss.getSheetByName(n);if(sh&&sh.getLastRow()>1)sh.deleteRows(2,sh.getLastRow()-1)});return {success:true,message:'Work/photo data reset. Users and Master settings were preserved.'};}
function audit_(action,user,details){try{const ss=getSpreadsheet_();let sh=ss.getSheetByName('Audit Log');if(!sh){sh=ss.insertSheet('Audit Log');sh.appendRow(['Date','User','Role','Action','Details']);}sh.appendRow([new Date(),user&&user.userName||'',user&&user.role||'',action,String(details||'')]);}catch(e){}}
function distanceMeters_(lat1,lng1,lat2,lng2){const R=6371000,toRad=x=>x*Math.PI/180,dLat=toRad(lat2-lat1),dLng=toRad(lng2-lng1);const a=Math.sin(dLat/2)**2+Math.cos(toRad(lat1))*Math.cos(toRad(lat2))*Math.sin(dLng/2)**2;return 2*R*Math.asin(Math.sqrt(a));}
function findPhotoRow_(photoId){const sh=setupPhotoDatabase_(),lr=sh.getLastRow();if(lr<2)throw new Error('Photo not found.');const rows=sh.getRange(2,1,lr-1,10).getValues();for(let i=0;i<rows.length;i++)if(String(rows[i][0]||'')===String(photoId||''))return {sheet:sh,row:i+2,data:rows[i]};throw new Error('Photo not found.');}
function updatePhoto_(d,token){const user=requirePermission_(token,'photoGPS'),p=findPhotoRow_(d.photoId);const work=findWork_(String(p.data[1]||''));if(!work||!isSansadAllowed_(String(work.sansadNo||''),getAllowedSansads_(user)))throw new Error('Permission denied.');const photoNo=Number(p.data[2]);const lat=Number(d.latitude),lng=Number(d.longitude);if(!isFinite(lat)||!isFinite(lng))throw new Error('Valid latitude/longitude required.');for(let n=1;n<=CONFIG.MAX_PHOTOS;n++){if(n===photoNo)continue;const c=6+(n-1)*2,olat=Number(work.sheet.getRange(work.row,c).getValue()),olng=Number(work.sheet.getRange(work.row,c+1).getValue());if(isFinite(olat)&&isFinite(olng)&&distanceMeters_(lat,lng,olat,olng)<CONFIG.MIN_PHOTO_DISTANCE_METERS)throw new Error('এই GPS location অন্য photo-র খুব কাছে। কমপক্ষে '+CONFIG.MIN_PHOTO_DISTANCE_METERS+' মিটার দূরে দিন।');}let fileUrl=String(p.data[4]||''),fileId=String(p.data[5]||'');if(d.photoBase64){const main=getMainDriveFolder_(),safeCode=String(work.workCode).replace(/[\\/:*?"<>|]/g,'_'),folderName=safeCode+'_'+String(work.workName||'Work').replace(/[\\/:*?"<>|]/g,'_'),fs=main.getFoldersByName(folderName),folder=fs.hasNext()?fs.next():main.createFolder(folderName),bytes=Utilities.base64Decode(String(d.photoBase64).split(',').pop()),fileName=safeCode+'_Photo_'+String(photoNo).padStart(2,'0')+'_EDITED.jpg',file=folder.createFile(Utilities.newBlob(bytes,'image/jpeg',fileName));if(fileId)try{DriveApp.getFileById(fileId).setTrashed(true)}catch(e){}fileUrl=file.getUrl();fileId=file.getId();p.sheet.getRange(p.row,4,1,3).setValues([[fileName,fileUrl,fileId]]);}const c=6+(photoNo-1)*2;work.sheet.getRange(work.row,c,1,2).setValues([[lat,lng]]);p.sheet.getRange(p.row,7,1,4).setValues([[lat,lng,new Date(),user.userName]]);audit_('PHOTO_EDIT',user,'Work '+work.workCode+', Photo '+photoNo);return {success:true,message:'Photo/GPS updated.'};}

function loginUser(userId,password){
  if(isMasterId_(userId)) return loginMaster_(userId,password);
  if(!userId || !password) throw new Error('User ID এবং Password দিন।');
  const wanted=String(userId).trim();
  const passwordHash=hashPassword_(password);
  let user=getCachedUser_(wanted);
  if(!user){
    const sheet=getUserSheet_();
    const lastRow=sheet.getLastRow();
    if(lastRow<2) throw new Error('কোনো User তৈরি করা হয়নি।');
    const data=sheet.getRange(2,1,lastRow-1,10).getValues();
    const wantedLower=wanted.toLowerCase();
    for(let i=0;i<data.length;i++){
      if(String(data[i][0]||'').trim().toLowerCase()===wantedLower){
        user=cacheUser_(recordFromRow_(data[i],i+2));
        break;
      }
    }
  }
  if(!user || user.passwordHash!==passwordHash) throw new Error('User ID অথবা Password ভুল।');
  if(String(user.active||'YES').toUpperCase()!=='YES') throw new Error('এই User inactive।');

  const token=Utilities.getUuid();
  const created=new Date();
  const expires=new Date(created.getTime()+CONFIG.SESSION_HOURS*60*60*1000);
  // Fast login: persist the active session in Script Properties/Cache.
  // The Sessions sheet is retained only for backward compatibility with old sessions.

  const sessionPayload={
    userId:user.userId,userName:user.userName,role:user.role,
    permissions:user.permissions,expires:expires.getTime(),user:user
  };
  CacheService.getScriptCache().put('pwg_session_'+token,JSON.stringify(sessionPayload),21600);
  PropertiesService.getScriptProperties().setProperty('pwg_session_'+token,JSON.stringify(sessionPayload));
  return {success:true,token:token,userId:user.userId,userName:user.userName,role:user.role,permissions:user.permissions};
}
function getCurrentUser_(token){
  const u=requireSession_(token);
  return {success:true,token:token,userId:u.userId,userName:u.userName,role:u.role,permissions:u.permissions};
}

function requireSession_(token){
  if(!token) throw new Error('Login session পাওয়া যায়নি। আবার Login করুন।');
  const cached=CacheService.getScriptCache().get('pwg_session_'+token);
  if(cached){
    try{
      const s=JSON.parse(cached);
      if(Number(s.expires)>Date.now()) return s.user;
    }catch(e){}
  }
  const stored=PropertiesService.getScriptProperties().getProperty('pwg_session_'+token);
  if(stored){
    try{
      const s=JSON.parse(stored);
      if(Number(s.expires)>Date.now()){
        CacheService.getScriptCache().put('pwg_session_'+token,stored,21600);
        return s.user;
      }
      PropertiesService.getScriptProperties().deleteProperty('pwg_session_'+token);
    }catch(e){}
  }
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName(CONFIG.SESSION_SHEET_NAME);
  if(!sh || sh.getLastRow()<2) throw new Error('Login session expired। আবার Login করুন।');
  const data=sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
  const now=new Date();
  for(let i=data.length-1;i>=0;i--){
    if(String(data[i][0])===String(token)){
      const expires=new Date(data[i][3]);
      if(expires<now) throw new Error('Login session expired। আবার Login করুন।');
      const u=findUser_(data[i][1]);
      const sessionPayload={
        userId:u.userId,userName:u.userName,role:u.role,
        permissions:u.permissions,expires:expires.getTime(),user:u
      };
      CacheService.getScriptCache().put('pwg_session_'+token,JSON.stringify(sessionPayload),21600);
      PropertiesService.getScriptProperties().setProperty('pwg_session_'+token,JSON.stringify(sessionPayload));
      return u;
    }
  }
  throw new Error('Invalid login session। আবার Login করুন।');
}

function requirePermission_(token,key){
  const user=requireSession_(token);
  if(isMaster_(user)) return user;
  if(!user.permissions[key]) throw new Error('এই কাজের permission দেওয়া হয়নি।');
  return user;
}

function getAllowedSansads_(user){ return String(user.permissions.sansadAccess||'ALL').trim() || 'ALL'; }
function normalizeSansad_(value){
  let s=String(value??'').trim().toLowerCase();
  if(!s) return '';
  s=s.replace(/\b(sansad|sangsad|সংসদ)\b/g,'').replace(/[._\-\/]+/g,' ').replace(/\s+/g,' ').trim();
  const m=s.match(/\d+(?:\.\d+)?/);
  if(m) return m[0].replace(/\.0+$/,'');
  const romanMap={i:'1',ii:'2',iii:'3',iv:'4',v:'5',vi:'6',vii:'7',viii:'8',ix:'9',x:'10',xi:'11',xii:'12',xiii:'13',xiv:'14',xv:'15',xvi:'16',xvii:'17',xviii:'18',xix:'19',xx:'20'};
  if(romanMap[s]) return romanMap[s];
  const bengaliDigits={'০':'0','১':'1','২':'2','৩':'3','৪':'4','৫':'5','৬':'6','৭':'7','৮':'8','৯':'9'};
  const bn=s.replace(/[০-৯]/g,function(ch){return bengaliDigits[ch]||ch;});
  const bm=bn.match(/\d+(?:\.\d+)?/);
  return bm ? bm[0].replace(/\.0+$/,'') : s;
}
function isSansadAllowed_(sansad,allowed){
  if(String(allowed).trim().toUpperCase()==='ALL') return true;
  const wanted=normalizeSansad_(sansad);
  if(!wanted) return false;
  const list=String(allowed).split(',').map(function(x){return normalizeSansad_(x);}).filter(Boolean);
  return list.indexOf(wanted)>=0;
}
function getActivity_(token){
  const user=requireSession_(token);
  const isAdminUser=String(user.role||'').toLowerCase()==='admin';
  const us=getUserSheet_(), ulr=us.getLastRow();
  const userRows=ulr>=2?us.getRange(2,1,ulr-1,10).getValues():[];
  const ps=setupPhotoDatabase_(), plr=ps.getLastRow();
  const photoRows=plr>=2?ps.getRange(2,1,plr-1,10).getValues():[];
  const workPhotoCount={};
  photoRows.forEach(function(r){const code=String(r[1]||'').trim();if(code)workPhotoCount[code.toLowerCase()]=(workPhotoCount[code.toLowerCase()]||0)+1;});
  if(!isAdminUser){
    const myName=String(user.userName||'').trim().toLowerCase(), myWorks={}; let myPhotos=0;
    photoRows.forEach(function(r){if(String(r[9]||'').trim().toLowerCase()===myName&&r[1]){myPhotos++;myWorks[String(r[1]).trim().toLowerCase()]=true;}});
    return {success:true,activeUserCount:0,totalUsers:0,workCount:Object.keys(myWorks).length,photoCount:myPhotos,users:[]};
  }
  const stats={};
  userRows.forEach(function(r){const id=String(r[0]||'').trim(),name=String(r[1]||id).trim();if(id)stats[id.toLowerCase()]={userId:id,user:name,active:String(r[4]||'YES').toUpperCase()==='YES',workSet:{},photoCount:0,completedSet:{}};});
  photoRows.forEach(function(r){
    const code=String(r[1]||'').trim(),name=String(r[9]||'').trim(); if(!code||!name)return;
    let st=null; Object.keys(stats).some(function(k){if(String(stats[k].user).toLowerCase()===name.toLowerCase()){st=stats[k];return true;}return false;});
    if(!st)return; st.photoCount++; st.workSet[code.toLowerCase()]=true; if(Number(workPhotoCount[code.toLowerCase()]||0)>=4)st.completedSet[code.toLowerCase()]=true;
  });
  const activeUsers=Object.keys(stats).filter(function(k){return stats[k].active;}).map(function(k){const st=stats[k];return {userId:st.userId,user:st.user,workCount:Object.keys(st.workSet).length,photoCount:st.photoCount,completedWorks:Object.keys(st.completedSet).length};}).sort(function(a,b){return (b.workCount-a.workCount)||(b.photoCount-a.photoCount)||a.user.localeCompare(b.user);});
  return {success:true,activeUserCount:activeUsers.length,totalUsers:userRows.length,workCount:Object.keys(workPhotoCount).length,photoCount:photoRows.length,users:activeUsers};
}

function logout_(token){
  CacheService.getScriptCache().remove('pwg_session_'+token);
  PropertiesService.getScriptProperties().deleteProperty('pwg_session_'+token);
  const ss=getSpreadsheet_();
  const sh=ss.getSheetByName(CONFIG.SESSION_SHEET_NAME);
  if(!sh || sh.getLastRow()<2) return {success:true};
  const data=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
  for(let i=data.length-1;i>=0;i--) if(String(data[i][0])===String(token)){sh.deleteRow(i+2);break;}
  return {success:true};
}

function getReport_(token){
  const user=requirePermission_(token,'report');
  const sheet=getWorkSheet_(),lr=sheet.getLastRow(); if(lr<2)return [];
  const allowed=getAllowedSansads_(user);
  return sheet.getRange(2,1,lr-1,15).getValues().filter(function(r){
    return r[2]!=='' && isSansadAllowed_(String(r[1]||''),allowed);
  }).map(function(r){
    return {slNo:r[0],sansadNo:r[1],workCode:r[2],workName:r[3],location:r[4],photos:r.slice(5,13),date:r[13],user:r[14]};
  });
}
function getDatabase_(token){
  const user=requireSession_(token); const sheet=getWorkSheet_(),lr=sheet.getLastRow(); if(lr<2)return []; const allowed=getAllowedSansads_(user);
  return sheet.getRange(2,1,lr-1,15).getValues().filter(r=>r[2]!==''&&isSansadAllowed_(String(r[1]||''),allowed)).map(r=>({slNo:String(r[0]||''),sansadNo:String(r[1]||''),workCode:String(r[2]||''),workName:String(r[3]||''),location:String(r[4]||''),photos:r.slice(5,13),date:r[13],user:String(r[14]||'')}));
}
function updateWork_(d,token){
  const user=requirePermission_(token,'addWork'); const oldCode=String(d.oldWorkCode||'').trim(),code=String(d.workCode||'').trim(); if(!oldCode||!code||!d.workName||!d.sansadNo)throw new Error('Sl No, Sansad No, Work Code and Work Name are required.'); const work=findWork_(oldCode); if(!work)throw new Error('Work not found.'); if(!isSansadAllowed_(String(d.sansadNo),getAllowedSansads_(user)))throw new Error('You do not have access to this Sansad.'); if(code.toLowerCase()!==oldCode.toLowerCase()&&findWork_(code))throw new Error('This Work Code already exists.'); work.sheet.getRange(work.row,1,1,5).setValues([[d.slNo||'',d.sansadNo,code,d.workName,d.location||'']]);
  if(code.toLowerCase()!==oldCode.toLowerCase()){
    const ps=setupPhotoDatabase_(), plr=ps.getLastRow();
    if(plr>=2){const vals=ps.getRange(2,2,plr-1,1).getValues();for(let i=0;i<vals.length;i++)if(String(vals[i][0]||'').trim().toLowerCase()===oldCode.toLowerCase())ps.getRange(i+2,2).setValue(code);}
  }
  return {success:true,message:'Work updated.'};
}
function deleteWork_(d,token){
  const user=requirePermission_(token,'addWork'),code=String(d.workCode||'').trim(); if(!code)throw new Error('Work Code is required.'); const work=findWork_(code); if(!work)throw new Error('Work not found.'); if(!isSansadAllowed_(String(work.sansadNo),getAllowedSansads_(user)))throw new Error('You do not have access to this Sansad.'); const ps=setupPhotoDatabase_(),plr=ps.getLastRow(); if(plr>=2){const rows=ps.getRange(2,1,plr-1,10).getValues();for(let i=rows.length-1;i>=0;i--)if(String(rows[i][1]||'').toLowerCase()===code.toLowerCase()){if(rows[i][5])try{DriveApp.getFileById(String(rows[i][5])).setTrashed(true)}catch(e){}ps.deleteRow(i+2);}} work.sheet.deleteRow(work.row); audit_('WORK_DELETE',user,code); return {success:true,message:'Work and associated photos deleted.'};
}
function deleteWorks_(d,token){
  const user=requirePermission_(token,'addWork');
  const codes=Array.isArray(d.workCodes)?d.workCodes.map(function(x){return String(x).trim().toLowerCase();}).filter(Boolean):[];
  if(!codes.length) throw new Error('No work selected.');
  const sheet=getWorkSheet_(), lr=sheet.getLastRow();
  if(lr<2) return {success:true,message:'No work records found.',deleted:0};
  const allowed=getAllowedSansads_(user), rows=sheet.getRange(2,1,lr-1,5).getValues();
  let deleted=0;
  for(let i=rows.length-1;i>=0;i--){const code=String(rows[i][2]||'').trim().toLowerCase();const sansad=String(rows[i][1]||'').trim();if(codes.indexOf(code)>=0 && isSansadAllowed_(sansad,allowed)){sheet.deleteRow(i+2);deleted++;}}
  return {success:true,message:deleted+' selected work(s) deleted.',deleted:deleted};
}
function deleteAllWorks_(token){
  const user=requirePermission_(token,'addWork');
  const sheet=getWorkSheet_(), lr=sheet.getLastRow();
  if(lr<2) return {success:true,message:'No work records found.',deleted:0};
  const isAdmin=String(user.role||'').toLowerCase()==='admin';
  if(isAdmin){
    const deleted=lr-1;
    sheet.getRange(2,1,deleted,15).clearContent();
    return {success:true,message:'Deleted '+deleted+' work record(s). Total Works is now 0.',deleted:deleted};
  }
  const allowed=getAllowedSansads_(user), rows=sheet.getRange(2,1,lr-1,5).getValues();
  let deleted=0;
  for(let i=rows.length-1;i>=0;i--){const sansad=String(rows[i][1]||'').trim();if(rows[i][2]!=='' && isSansadAllowed_(sansad,allowed)){sheet.deleteRow(i+2);deleted++;}}
  return {success:true,message:'Deleted '+deleted+' work record(s).',deleted:deleted};
}
function uploadExcel_(d,token){
  const user=requirePermission_(token,'addWork'); if(!d.fileBase64)throw new Error('Excel file data not received.'); const name=String(d.fileName||'').toLowerCase(),base64=String(d.fileBase64).split(',').pop(); let rows=[];
  if(name.endsWith('.csv'))rows=parseCsv_(Utilities.newBlob(Utilities.base64Decode(base64)).getDataAsString('UTF-8')); else if(name.endsWith('.xlsx'))rows=parseXlsx_(Utilities.newBlob(Utilities.base64Decode(base64),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',d.fileName)); else throw new Error('Only .xlsx or .csv files are supported.');
  const sheet=getWorkSheet_(); setupSheet(); const lr=sheet.getLastRow(); const existing=lr>1?sheet.getRange(2,3,lr-1,1).getValues().map(r=>String(r[0]).trim().toLowerCase()):[]; let added=0,skipped=0; const out=[];
  rows.forEach(r=>{const sl=r.slNo||String(lr+added),sansad=String(r.sansadNo||'').trim(),code=String(r.workCode||'').trim(),name=String(r.workName||'').trim(),loc=String(r.location||'').trim(); if(!code||!name||!sansad||!isSansadAllowed_(sansad,getAllowedSansads_(user))||existing.indexOf(code.toLowerCase())>=0){skipped++;return;} out.push([sl,sansad,code,name,loc]); existing.push(code.toLowerCase());added++;});
  if(out.length)sheet.getRange(sheet.getLastRow()+1,1,out.length,5).setValues(out); return {success:true,added:added,skipped:skipped,message:'Excel import completed.'};
}
function parseCsv_(text){const lines=String(text).split(/\r?\n/).filter(x=>x.trim()!=='');return lines.length?mapImportRows_(lines.map(parseCsvLine_)):[];}
function parseCsvLine_(line){let o=[],c='',q=false;for(let i=0;i<line.length;i++){const x=line[i];if(x==='"'){if(q&&line[i+1]==='"'){c+='"';i++;}else q=!q;}else if(x===','&&!q){o.push(c);c='';}else c+=x;}o.push(c);return o;}
function mapImportRows_(rows){if(!rows.length)return [];const h=rows[0].map(x=>String(x).trim().toLowerCase()),idx=names=>{for(const n of names){const i=h.indexOf(n);if(i>=0)return i;}return -1;},a=idx(['sl no','slno','serial','serial no']),b=idx(['sansad no','sansad','sansad number']),c=idx(['work code','workcode','code']),d=idx(['work name','workname','name']),e=idx(['location','site','place']),start=c>=0?1:0;return rows.slice(start).map(r=>({slNo:a>=0?r[a]:'',sansadNo:b>=0?r[b]:r[1]||'',workCode:c>=0?r[c]:r[2]||'',workName:d>=0?r[d]:r[3]||'',location:e>=0?r[e]:r[4]||''}));}
function parseXlsx_(blob){blob=Utilities.newBlob(blob.getBytes(),'application/zip',blob.getName()||'upload.xlsx');const files=Utilities.unzip(blob),m={};files.forEach(f=>m[f.getName()]=f);const sb=m['xl/worksheets/sheet1.xml'];if(!sb)throw new Error('First worksheet not found.');let shared=[];if(m['xl/sharedStrings.xml']){const r=XmlService.parse(m['xl/sharedStrings.xml'].getDataAsString()).getRootElement();const ns=r.getNamespace();shared=r.getChildren('si',ns).map(si=>si.getChildren('t',ns).map(t=>t.getText()).join(''));}const root=XmlService.parse(sb.getDataAsString()).getRootElement(),ns=root.getNamespace(),sd=root.getChild('sheetData',ns);if(!sd)return [];const rows=sd.getChildren('row',ns).map(re=>{const a=[];re.getChildren('c',ns).forEach(c=>{const at=c.getAttribute('r'),t=c.getAttribute('t'),v=c.getChild('v',ns);const col=lettersToCol_((at?at.getValue():'A').match(/[A-Z]+/)[0]);let val=v?v.getText():'';if(t&&t.getValue()==='s')val=shared[Number(val)]||'';a[col-1]=val;});return a;});return mapImportRows_(rows);}
function lettersToCol_(s){let n=0;for(let i=0;i<s.length;i++)n=n*26+s.charCodeAt(i)-64;return n;}

function setupSystem(){ setupSheet(); setupUsers(); setupPhotoDatabase_(); getMainDriveFolder_(); setupMaster_(); return {success:true,message:'System setup completed. Master ID: '+CONFIG.MASTER_USER_ID+' | Initial password: '+CONFIG.MASTER_INITIAL_PASSWORD+' (change immediately).'}; }
