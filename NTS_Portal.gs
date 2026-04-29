// ╔══════════════════════════════════════════════════════════════════╗
// ║  NTS Group Portal — Google Apps Script (versiune completă)     ║
// ║  Citire stiluri text din Plan: bold · italic · dimensiune font  ║
// ╚══════════════════════════════════════════════════════════════════╝

// ══ CONFIGURARE ══════════════════════════════════════════════════════
// ID spreadsheet PORTAL (Rapoarte, Users, Certificari, OTP, Config)
// → URL: docs.google.com/spreadsheets/d/ [ID] /edit
var PORTAL_SS_ID = '';   // completează cu ID-ul tău dacă nu e bound

// ID spreadsheet PLAN NTS Field 2026
var PLAN_SS_ID   = '1VfP_SvzaJBUu4BVLmsxTL-s6vrmDXogueBYSCrg0rAs';

// Emailuri
var ADMIN_EMAIL  = 'networkticketsolution@gmail.com';
var EMAIL_TO     = 'contabilitate@networkticketsolution.com';

// Nume foi în spreadsheet-ul Portal
var SHEET_RAPOARTE = 'Rapoarte';
var SHEET_USERS    = 'Users';
var SHEET_CERTS    = 'Certificari';
var SHEET_CONFIG   = 'Config';
var SHEET_OTP      = 'OTP';

// ══ ENTRY POINT ══════════════════════════════════════════════════════
function doGet(e) {
  var params  = e && e.parameter ? e.parameter : {};
  var payload = {};
  if (params.payload) {
    try { payload = JSON.parse(decodeURIComponent(params.payload)); } catch(ex) {}
  }
  var p = {};
  Object.keys(params).forEach(function(k)  { p[k] = params[k];  });
  Object.keys(payload).forEach(function(k) { p[k] = payload[k]; });
  var action = p.action || '';

  var result;
  try {
    if      (action === 'getData')       result = getData_();
    else if (action === 'addRaport')     result = addRaport_(p);
    else if (action === 'sendEmail')     result = sendEmail_(p);
    else if (action === 'saveCert')      result = saveCert_(p);
    else if (action === 'requestOtp')    result = requestOtp_(p);
    else if (action === 'verifyOtp')     result = verifyOtp_(p);
    else if (action === 'checkStatus')   result = checkStatus_(p);
    else if (action === 'getPending')    result = getPending_(p);
    else if (action === 'approveUser')   result = approveUser_(p);
    else if (action === 'denyUser')      result = denyUser_(p);
    else if (action === 'getPlanSheets') result = getPlanSheets_();
    else if (action === 'getPlanData')   result = getPlanData_(p.sheet);
    else result = { ok: false, error: 'Actiune necunoscuta: ' + action };
  } catch(ex) {
    result = { ok: false, error: ex.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) { return doGet(e); }

// ══ HELPER SPREADSHEET PORTAL ════════════════════════════════════════
function getPortalSS_() {
  if (PORTAL_SS_ID && PORTAL_SS_ID.length > 10)
    return SpreadsheetApp.openById(PORTAL_SS_ID);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  throw new Error('PORTAL_SS_ID negăsit. Completează variabila din script.');
}
function getSheet_(name) {
  var ss    = getPortalSS_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ══ getData ══════════════════════════════════════════════════════════
function getData_() {
  return {
    ok: true,
    data: {
      rapoarte:    getRapoarte_(),
      certificari: getCertificari_(),
      plan_url:    getConfig_().plan_url  || '',
      pontaj_url:  getConfig_().pontaj_url || '',
      planSheets:  getPlanSheetsList_()
    }
  };
}

// ── Citire rapoarte — mapare flexibilă (orice format de header) ───
var RAPORT_ALIASES = {
  timestamp:     ['timestamp','Timestamp','TIMESTAMP','data_ora'],
  data:          ['data','Data','DATA','date','Date','zi','Zi'],
  contractor:    ['contractor','Contractor','CONTRACTOR'],
  technicieni:   ['technicieni','Technicieni','tehnicieni','Tehnicieni'],
  nrSite:        ['nrSite','NrSite','nr_site','Nr Site','site','Site'],
  tipLucrare:    ['tipLucrare','TipLucrare','tip_lucrare','Tip Lucrare'],
  status:        ['status','Status','STATUS'],
  descriere:     ['descriere','Descriere','DESCRIERE','description'],
  noteInlocuire: ['noteInlocuire','NoteInlocuire','note_inlocuire','Note Inlocuire']
};
function resolveHeader_(hRow) {
  var map = {};
  hRow.forEach(function(h,i) {
    var hs = String(h||'').trim();
    Object.keys(RAPORT_ALIASES).forEach(function(key) {
      if (RAPORT_ALIASES[key].indexOf(hs) !== -1) map[i] = key;
    });
  });
  return map;
}
function getRapoarte_() {
  var sheet  = getSheet_(SHEET_RAPOARTE);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var colMap  = resolveHeader_(values[0]);
  var useFb   = Object.keys(colMap).length === 0;
  var posKeys = ['timestamp','data','contractor','technicieni',
                 'nrSite','tipLucrare','status','descriere','noteInlocuire'];
  return values.slice(1).map(function(row) {
    var obj = {};
    if (useFb) { posKeys.forEach(function(k,i){ obj[k]=row[i]!==undefined?row[i]:''; }); }
    else { Object.keys(colMap).forEach(function(i){ obj[colMap[i]]=row[i]!==undefined?row[i]:''; }); }
    return obj;
  }).filter(function(r){ return r.timestamp || r.data; });
}
function getCertificari_() {
  var sheet  = getSheet_(SHEET_CERTS);
  var values = sheet.getDataRange().getValues();
  var certs  = {};
  if (values.length < 2) return certs;
  values.slice(1).forEach(function(row){
    if (row[0]) certs[String(row[0])]={ url:row[1]||'', nota:row[2]||'' };
  });
  return certs;
}
function getConfig_() {
  var sheet  = getSheet_(SHEET_CONFIG);
  var values = sheet.getDataRange().getValues();
  var cfg    = {};
  values.forEach(function(row){ if(row[0]) cfg[String(row[0])]=row[1]||''; });
  return cfg;
}
function getPlanSheetsList_() {
  try {
    var ss = SpreadsheetApp.openById(PLAN_SS_ID);
    return ss.getSheets()
      .filter(function(s){ return /^KW\d{2}$/i.test(s.getName()); })
      .map(function(s){ return { name:s.getName().toUpperCase(), gid:s.getSheetId() }; })
      .sort(function(a,b){ return a.name.localeCompare(b.name); });
  } catch(e){ return []; }
}

// ══ addRaport ════════════════════════════════════════════════════════
function addRaport_(p) {
  var sheet  = getSheet_(SHEET_RAPOARTE);
  var values = sheet.getDataRange().getValues();
  if (values.length===0||(values.length===1&&!values[0][0])) {
    sheet.getRange(1,1,1,9).setValues([[
      'timestamp','data','contractor','technicieni',
      'nrSite','tipLucrare','status','descriere','noteInlocuire'
    ]]);
  }
  sheet.appendRow([
    p.timestamp||new Date().toISOString(), p.data||'', p.contractor||'',
    p.technicieni||'', p.nrSite||'—', p.tipLucrare||'', p.status||'',
    p.descriere||'—', p.noteInlocuire||'—'
  ]);
  return { ok: true };
}

// ══ sendEmail ════════════════════════════════════════════════════════
function sendEmail_(p) {
  var subject = '[NTS Solicitare] '+(p.subiect||'')+' – '+(p.nume||'')+' ('+(p.data||'')+')';
  var body    = 'Bună ziua,\n\nNume: '+(p.nume||'')+'\nData: '+(p.data||'')+
                '\nSubiect: '+(p.subiect||'')+'\n\n'+(p.mesaj||'')+'\n\n---\nNTS Group Portal';
  var opts = { name:'NTS Group Portal' };
  if (p.areSig && p.sigImg) {
    try {
      opts.attachments = [Utilities.newBlob(
        Utilities.base64Decode(p.sigImg.replace(/^data:image\/png;base64,/,'')),
        'image/png','semnatura.png')];
    } catch(e){}
  }
  if (p.areImg && p.imgData) {
    try {
      var ext = p.imgData.match(/^data:image\/(\w+);base64,/);
      opts.attachments = (opts.attachments||[]).concat([Utilities.newBlob(
        Utilities.base64Decode(p.imgData.replace(/^data:image\/\w+;base64,/,'')),
        ext?'image/'+ext[1]:'image/jpeg','imagine.'+(ext?ext[1]:'jpg'))]);
    } catch(e){}
  }
  GmailApp.sendEmail(EMAIL_TO, subject, body, opts);
  return { ok: true };
}

// ══ saveCert ═════════════════════════════════════════════════════════
function saveCert_(p) {
  var ini=p.ini||''; if(!ini) return {ok:false,error:'Lipsește ini'};
  var sheet=getSheet_(SHEET_CERTS), values=sheet.getDataRange().getValues();
  if(values.length===0||!values[0][0]){sheet.getRange(1,1,1,3).setValues([['ini','url','nota']]);values=[['ini','url','nota']];}
  for(var i=1;i<values.length;i++){if(values[i][0]===ini){sheet.getRange(i+1,1,1,3).setValues([[ini,p.url||'',p.nota||'']]);return{ok:true};}}
  sheet.appendRow([ini,p.url||'',p.nota||'']); return {ok:true};
}

// ══ AUTENTIFICARE OTP ════════════════════════════════════════════════
function requestOtp_(p) {
  var email=(p.email||'').trim().toLowerCase(), name=(p.name||'').trim();
  if(!email) return {ok:false,error:'Email lipsă'};
  var sheet=getSheet_(SHEET_USERS), values=sheet.getDataRange().getValues();
  if(values.length===0||!values[0][0]){sheet.getRange(1,1,1,5).setValues([['email','name','status','requestedAt','approvedAt']]);values=[['email','name','status','requestedAt','approvedAt']];}
  var found=null,rowIdx=-1;
  for(var i=1;i<values.length;i++){if(String(values[i][0]).toLowerCase()===email){found=values[i];rowIdx=i+1;break;}}
  if(!found){
    sheet.appendRow([email,name,'pending',new Date().toISOString(),'']);
    try{GmailApp.sendEmail(ADMIN_EMAIL,'[NTS Portal] Cerere acces: '+name,'Utilizatorul '+name+' ('+email+') solicită acces.\nAprobă din panoul Admin.',{name:'NTS Portal'});}catch(e){}
    return {ok:true,status:'pending',accessStatus:'new'};
  }
  var status=String(found[2]||'pending').toLowerCase();
  if(status==='denied')  return {ok:true,status:'denied', accessStatus:'denied'};
  if(status==='pending') return {ok:true,status:'pending',accessStatus:'pending'};
  if(status==='approved'){
    var otp=Math.floor(100000+Math.random()*900000).toString();
    var exp=new Date(Date.now()+10*60*1000).toISOString();
    var os=getSheet_(SHEET_OTP), ov=os.getDataRange().getValues();
    if(ov.length===0||!ov[0][0]) os.getRange(1,1,1,3).setValues([['email','otp','expireAt']]);
    ov=os.getDataRange().getValues();
    for(var j=1;j<ov.length;j++){if(String(ov[j][0]).toLowerCase()===email){os.deleteRow(j+1);break;}}
    os.appendRow([email,otp,exp]);
    try{
      GmailApp.sendEmail(email,'[NTS Portal] Cod acces: '+otp,'Codul tău:\n\n'+otp+'\n\nValabil 10 minute.\n---\nNTS Group Portal',{name:'NTS Group Portal'});
      return {ok:true,status:'approved',accessStatus:'approved',otpSent:true};
    }catch(e){return {ok:false,error:'Nu s-a trimis codul: '+e.message};}
  }
  return {ok:true,status:status,accessStatus:status};
}
function verifyOtp_(p) {
  var email=(p.email||'').trim().toLowerCase(), otp=(p.otp||'').trim();
  if(!email||!otp) return {ok:false,error:'Date lipsă'};
  var sheet=getSheet_(SHEET_OTP), values=sheet.getDataRange().getValues(), now=Date.now();
  for(var i=1;i<values.length;i++){
    if(String(values[i][0]).toLowerCase()===email){
      if(String(values[i][1])===otp&&now<new Date(values[i][2]).getTime()){sheet.deleteRow(i+1);return {ok:true,verified:true};}
      return {ok:false,error:'Cod incorect sau expirat',verified:false};
    }
  }
  return {ok:false,error:'Codul nu a fost găsit',verified:false};
}
function checkStatus_(p) {
  var email=(p.email||'').trim().toLowerCase(), sheet=getSheet_(SHEET_USERS), values=sheet.getDataRange().getValues();
  for(var i=1;i<values.length;i++){if(String(values[i][0]).toLowerCase()===email)return{ok:true,status:String(values[i][2]||'pending').toLowerCase()};}
  return {ok:true,status:'unknown'};
}
function getPending_(p) {
  if((p.adminEmail||'').toLowerCase()!==ADMIN_EMAIL.toLowerCase())return{ok:false,error:'Acces interzis'};
  var sheet=getSheet_(SHEET_USERS),values=sheet.getDataRange().getValues(),list=[];
  for(var i=1;i<values.length;i++){if(String(values[i][2]||'').toLowerCase()==='pending')list.push({email:values[i][0],name:values[i][1],requestedAt:values[i][3]});}
  return {ok:true,pending:list};
}
function approveUser_(p) {
  if((p.adminEmail||'').toLowerCase()!==ADMIN_EMAIL.toLowerCase())return{ok:false,error:'Acces interzis'};
  var target=(p.targetEmail||'').toLowerCase(),sheet=getSheet_(SHEET_USERS),values=sheet.getDataRange().getValues();
  for(var i=1;i<values.length;i++){
    if(String(values[i][0]).toLowerCase()===target){
      sheet.getRange(i+1,3).setValue('approved');sheet.getRange(i+1,5).setValue(new Date().toISOString());
      try{GmailApp.sendEmail(target,'[NTS Portal] Acces aprobat!','Accesul la NTS Group Portal a fost aprobat.\n---\nNTS Group Portal',{name:'NTS Group Portal'});}catch(e){}
      return{ok:true};
    }
  }
  return{ok:false,error:'User negăsit'};
}
function denyUser_(p) {
  if((p.adminEmail||'').toLowerCase()!==ADMIN_EMAIL.toLowerCase())return{ok:false,error:'Acces interzis'};
  var target=(p.targetEmail||'').toLowerCase(),sheet=getSheet_(SHEET_USERS),values=sheet.getDataRange().getValues();
  for(var i=1;i<values.length;i++){if(String(values[i][0]).toLowerCase()===target){sheet.getRange(i+1,3).setValue('denied');return{ok:true};}}
  return{ok:false,error:'User negăsit'};
}

// ══ getPlanSheets ════════════════════════════════════════════════════
function getPlanSheets_() {
  try {
    var ss=SpreadsheetApp.openById(PLAN_SS_ID);
    var sheets=ss.getSheets()
      .filter(function(s){return /^KW\d{2}$/i.test(s.getName());})
      .map(function(s){return{name:s.getName().toUpperCase(),gid:s.getSheetId()};})
      .sort(function(a,b){return a.name.localeCompare(b.name);});
    return {ok:true,planSheets:sheets};
  }catch(e){return{ok:false,error:'Nu pot accesa Plan NTS Field: '+e.message};}
}

// ══ getPlanData — date + stiluri TEXT (bold · italic · font size) ═════
// NU trimite culori de fundal — tema dark a portalului rămâne intactă
function getPlanData_(sheetName) {
  if (!sheetName || !/^KW\d{2}$/i.test(sheetName))
    return { ok:false, error:'Nume foaie invalid: '+sheetName };
  try {
    var ss=SpreadsheetApp.openById(PLAN_SS_ID), name=sheetName.toUpperCase(), sheet=null;
    var all=ss.getSheets();
    for(var i=0;i<all.length;i++){if(all[i].getName().toUpperCase()===name){sheet=all[i];break;}}
    if(!sheet) return {ok:false,error:'Foaia '+name+' nu există'};

    var range = sheet.getDataRange();

    // Date
    var values = range.getValues();

    // ── Stiluri TEXT — fără culori de fundal, fără culori text ────
    var fontWeights  = range.getFontWeights();   // 'bold' | 'normal'
    var fontStyles   = range.getFontStyles();    // 'italic' | 'normal'
    var fontSizes    = range.getFontSizes();     // număr în pt (ex: 10, 11, 12)
    var fontFamilies = range.getFontFamilies();  // 'Arial', 'Calibri' etc.

    // Elimină rânduri goale de la final
    var keepRows = [];
    for(var r=0;r<values.length;r++){
      var empty=values[r].every(function(c){return c===''||c===null||c===undefined;});
      if(!empty) keepRows.push(r);
    }
    if(keepRows.length===0&&values.length>0) keepRows.push(0);
    var sl=function(arr){return keepRows.map(function(i){return arr[i];});};

    return {
      ok:    true,
      sheet: name,
      data:  sl(values),
      styles: {
        fontWeights:  sl(fontWeights),
        fontStyles:   sl(fontStyles),
        fontSizes:    sl(fontSizes),
        fontFamilies: sl(fontFamilies)
      }
    };
  }catch(e){
    return {ok:false,error:'Eroare citire '+sheetName+': '+e.message};
  }
}
