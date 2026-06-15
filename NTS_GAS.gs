// ╔══════════════════════════════════════════════════════════════════╗
// ║  NTS Group Portal — Google Apps Script                          ║
// ║                                                                  ║
// ║  FLUX ACCESS:                                                    ║
// ║  1. Userul cere acces din portal (email + nume)                  ║
// ║  2. Admin vede cererea în foaia AccessControl (status: pending)  ║
// ║  3. Admin schimbă statusul în "approved" sau "verified" în GSheet║
// ║  4. Userul apasă "Verifică acum" → detectează status → intră    ║
// ║  5. Sesiunea se salvează local (localStorage)                    ║
// ║                                                                  ║
// ║  MODIFICARE: user șters din spreadsheet → tratat ca "denied"    ║
// ║  (anterior era "unknown" și accesul era menținut)                ║
// ╚══════════════════════════════════════════════════════════════════╝

// ══ CONFIGURARE ══════════════════════════════════════════════════════
var PORTAL_SS_ID = '1KE-nHKDd6iyABN7sJvlH69ac6JdVklUoF2hd_alMymg';
var PLAN_SS_ID   = '1VfP_SvzaJBUu4BVLmsxTL-s6vrmDXogueBYSCrg0rAs';
var ADMIN_EMAIL  = 'networkticketsolution@gmail.com';
var EMAIL_TO     = 'contabilitate@networkticketsolution.com';
var PORTAL_URL   = 'https://networkticketsolution.com';

var SHEET_RAPOARTE   = 'Rapoarte';
var SHEET_USERS      = 'AccessControl';
var SHEET_CERTS      = 'Certificari';
var SHEET_CONFIG     = 'Config';
var SHEET_TEHNICIENI = 'Tehnicieni';
var SHEET_MESAJE     = 'Mesaje';

// Coloane AccessControl (1-based)
// A=Email  B=Nume  C=Status  D=SolicitatLa  E=AprobatLa
var COL_EMAIL     = 1;
var COL_NUME      = 2;
var COL_STATUS    = 3;
var COL_SOLICITAT = 4;
var COL_APROBAT   = 5;

// ══ ENTRY POINT ══════════════════════════════════════════════════════
function doGet(e) {
  var params = e && e.parameter ? e.parameter : {};
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
    if      (action === 'getData')        result = getData_();
    else if (action === 'addRaport')      result = addRaport_(p);
    else if (action === 'sendEmail')      result = sendEmail_(p);
    else if (action === 'saveCert')       result = saveCert_(p);
    // ── ACCES CONTROL ──────────────────────────────────────────────
    else if (action === 'requestAccess')  result = requestAccess_(p);
    else if (action === 'requestOtp')     result = requestAccess_(p);
    else if (action === 'checkStatus')    result = checkStatus_(p);
    else if (action === 'getPending')     result = getPending_(p);
    else if (action === 'approveUser')    result = approveUser_(p);
    else if (action === 'denyUser')       result = denyUser_(p);
    // ── PLAN ───────────────────────────────────────────────────────
    else if (action === 'getPlanSheets')  result = getPlanSheets_();
    else if (action === 'getPlanData')    result = getPlanData_(p.sheet);
    else result = { ok: false, error: 'Actiune necunoscuta: ' + action };
  } catch(ex) {
    result = { ok: false, error: ex.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      var bodyParams = JSON.parse(e.postData.contents);
      if (!e.parameter) e.parameter = {};
      Object.keys(bodyParams).forEach(function(k) {
        e.parameter[k] = bodyParams[k];
      });
      e.parameter.payload = JSON.stringify(bodyParams);
    } catch(ex) {}
  }
  return doGet(e);
}

// ══ HELPERS SPREADSHEET ══════════════════════════════════════════════
var _cachedSS = null;
function getPortalSS_() {
  if (PORTAL_SS_ID && PORTAL_SS_ID.trim().length > 10)
    return SpreadsheetApp.openById(PORTAL_SS_ID.trim());
  return SpreadsheetApp.getActiveSpreadsheet();
}
function getPortalSSCached_() {
  if (!_cachedSS) _cachedSS = getPortalSS_();
  return _cachedSS;
}
function getSheet_(name) {
  var ss = getPortalSSCached_();
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ══ getData ══════════════════════════════════════════════════════════
function getData_() {
  var cfg = getConfig_();
  return {
    ok: true,
    data: {
      rapoarte:    getRapoarte_(),
      certificari: getCertificari_(),
      tehnicieni:  getTehnicieni_(),
      mesaje:      getAllMessagesByTech_(),
      plan_url:    cfg.plan_url   || '',
      pontaj_url:  cfg.pontaj_url || '',
      planSheets:  getPlanSheetsList_(),
      // ── Credențiale NAS — citite din tab-ul Config ────────────────
      // Modifică valorile direct în foaia Config (rândurile nas_user / nas_pass)
      nas_user:    cfg.nas_user   || '',
      nas_pass:    cfg.nas_pass   || ''
    }
  };
}

var RAPORT_ALIASES = {
  timestamp:     ['timestamp','Timestamp','TIMESTAMP','data_ora','Data Ora'],
  data:          ['data','Data','DATA','date','Date','zi','Zi','Data Raportului'],
  contractor:    ['contractor','Contractor','CONTRACTOR'],
  technicieni:   ['technicieni','Technicieni','tehnicieni','Tehnicieni','tehnician','Tehnician'],
  nrSite:        ['nrSite','NrSite','nr_site','Nr Site','Nr. Site','nr.site','Nr.Site',
                  'site','Site','SITE','numar_site','Numar Site','NumarSite','ID Site'],
  tipLucrare:    ['tipLucrare','TipLucrare','tip_lucrare','Tip Lucrare','Tip lucrare',
                  'lucrare','Lucrare','LUCRARE'],
  status:        ['status','Status','STATUS'],
  descriere:     ['descriere','Descriere','DESCRIERE','description','Description','detalii','Detalii'],
  noteInlocuire: ['noteInlocuire','NoteInlocuire','note_inlocuire','Note Inlocuire',
                  'note','Note','observatii','Observatii']
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

// ══ getTehnicieni ════════════════════════════════════════════════════
function getTehnicieni_() {
  try {
    var ss = SpreadsheetApp.openById(PORTAL_SS_ID);
    var sheet = ss.getSheetByName(SHEET_TEHNICIENI);
    if (!sheet) return [];
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    var header = values[0].map(function(h){ return String(h||'').trim().toLowerCase(); });
    var idxTeam = -1, idxName = -1, idxIni = -1, idxPri = -1;
    header.forEach(function(h, i){
      if      (h === 'team' || h === 'echipa')                           idxTeam = i;
      else if (h === 'name' || h === 'nume')                             idxName = i;
      else if (h === 'initials' || h === 'ini' || h === 'initiale')      idxIni  = i;
      else if (h === 'primary' || h === 'pri' || h === 'titular')        idxPri  = i;
    });
    if (idxTeam === -1) idxTeam = 0;
    if (idxName === -1) idxName = 1;
    if (idxIni  === -1) idxIni  = 2;
    if (idxPri  === -1) idxPri  = 3;

    var techs = [];
    values.slice(1).forEach(function(row){
      var name = String(row[idxName]||'').trim();
      var team = String(row[idxTeam]||'').trim();
      if (!name || !team) return;
      var pv  = row[idxPri];
      var pri = (pv === true || pv === 1 ||
                 String(pv).toLowerCase() === 'true' ||
                 String(pv).toLowerCase() === 'da'   ||
                 String(pv).toLowerCase() === 'yes');
      var ini = String(row[idxIni]||'').trim();
      if (!ini) {
        ini = name.split(/\s+/).map(function(p){ return (p[0]||''); }).join('').substring(0,2).toUpperCase();
      }
      techs.push({ name: name, team: team, ini: ini, pri: pri });
    });
    return techs;
  } catch(e) {
    return [];
  }
}

// ══ getConfig ════════════════════════════════════════════════════════
// Citește toate cheile din tab-ul Config (coloana A = cheie, coloana B = valoare)
//
// ┌─────────────┬──────────────────────┐
// │ A           │ B                    │
// ├─────────────┼──────────────────────┤
// │ plan_url    │ https://...          │
// │ pontaj_url  │ https://...          │
// │ nas_user    │ Field                │  ← modifică aici username-ul NAS
// │ nas_pass    │ NTSTeamWork2026!     │  ← modifică aici parola NAS
// └─────────────┴──────────────────────┘
//
function getConfig_() {
  var sheet  = getSheet_(SHEET_CONFIG);
  var values = sheet.getDataRange().getValues();
  var cfg    = {};
  values.forEach(function(row){ if(row[0]) cfg[String(row[0]).trim()]=String(row[1]||''); });
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
var RAPORT_HEADER = ['timestamp','data','contractor','technicieni',
                     'nrSite','tipLucrare','status','descriere','noteInlocuire'];

function addRaport_(p) {
  var sheet  = getSheet_(SHEET_RAPOARTE);
  var values = sheet.getDataRange().getValues();
  if (values.length===0||(values.length===1&&!values[0][0])) {
    sheet.getRange(1,1,1,9).setValues([RAPORT_HEADER]); values=[RAPORT_HEADER];
  }
  var eh=values[0].map(function(h){return String(h||'').trim();}), fix=false;
  for(var hi=0;hi<RAPORT_HEADER.length;hi++){if(eh[hi]!==RAPORT_HEADER[hi]){fix=true;break;}}
  if(fix){
    var isD=values[0].some(function(c){return c&&!/^(timestamp|data|contractor|technicieni|nrSite|tipLucrare|status|descriere|noteInlocuire)$/i.test(String(c));});
    if(isD)sheet.insertRowBefore(1);
    sheet.getRange(1,1,1,9).setValues([RAPORT_HEADER]);
  }
  sheet.appendRow([
    p.timestamp||new Date().toISOString(),
    (p.data?(String(p.data).length===10?p.data+'T12:00:00':p.data):''),
    p.contractor||'', p.technicieni||'', p.nrSite||'—',
    p.tipLucrare||'', p.status||'', p.descriere||'—', p.noteInlocuire||'—'
  ]);
  return { ok:true };
}

// ══ sendEmail ════════════════════════════════════════════════════════
function sendEmail_(p) {
  var subject = '[NTS Solicitare] ' + (p.subiect||'') + ' – ' + (p.nume||'') + ' (' + (p.data||'') + ')';
  var body    = 'Bună ziua,\n\nNume: '    + (p.nume||'')    +
                '\nData: '     + (p.data||'')    +
                '\nSubiect: '  + (p.subiect||'') +
                '\n\n'         + (p.mesaj||'')   +
                '\n\n---\nNTS Group Portal';
  var opts = { name: 'NTS Group Portal' };
  if (p.areSig && p.sigImg && p.sigImg.length > 100) {
    try {
      var sigData = p.sigImg.replace(/^data:image\/png;base64,/, '');
      opts.attachments = [Utilities.newBlob(
        Utilities.base64Decode(sigData), 'image/png', 'semnatura.png'
      )];
    } catch(e) { body += '\n\n[Notă: Semnătura nu a putut fi atașată]'; }
  }
  if (p.areImg && p.imgData && p.imgData.length > 100) {
    try {
      var match  = p.imgData.match(/^data:image\/(\w+);base64,/);
      var ext    = match ? match[1] : 'jpg';
      var imgRaw = p.imgData.replace(/^data:image\/\w+;base64,/, '');
      opts.attachments = (opts.attachments || []).concat([
        Utilities.newBlob(Utilities.base64Decode(imgRaw), 'image/' + ext, 'imagine.' + ext)
      ]);
    } catch(e) { body += '\n\n[Notă: Imaginea nu a putut fi atașată]'; }
  }
  try {
    GmailApp.sendEmail(EMAIL_TO, subject, body, opts);
    return { ok: true };
  } catch(e) {
    try {
      if (p.areSig) body += '\n\n[Semnătura a fost obținută dar nu a putut fi atașată]';
      if (p.areImg) body += '\n[Imaginea a fost obținută dar nu a putut fi atașată]';
      GmailApp.sendEmail(EMAIL_TO, subject, body, { name: 'NTS Group Portal' });
      return { ok: true, warning: 'Trimis fără atașamente: ' + e.message };
    } catch(e2) {
      return { ok: false, error: 'Nu s-a putut trimite emailul: ' + e2.message };
    }
  }
}

// ══ saveCert ═════════════════════════════════════════════════════════
function saveCert_(p) {
  var ini=p.ini||''; if(!ini) return {ok:false,error:'Lipsește ini'};
  var sheet=getSheet_(SHEET_CERTS),values=sheet.getDataRange().getValues();
  if(values.length===0||!values[0][0]){sheet.getRange(1,1,1,3).setValues([['ini','url','nota']]);values=[['ini','url','nota']];}
  for(var i=1;i<values.length;i++){if(values[i][0]===ini){sheet.getRange(i+1,1,1,3).setValues([[ini,p.url||'',p.nota||'']]);return{ok:true};}}
  sheet.appendRow([ini,p.url||'',p.nota||'']); return{ok:true};
}

// ══ ACCESS CONTROL ═══════════════════════════════════════════════════

function findUser_(email) {
  var sheet  = getSheet_(SHEET_USERS);
  var values = sheet.getDataRange().getValues();
  if (values.length === 0 || !values[0][0]) {
    sheet.getRange(1,1,1,5).setValues([['Email','Nume','Status','SolicitatLa','AprobatLa']]);
    return { sheet:sheet, found:null, rowIdx:-1, values:[] };
  }
  for (var i=1; i<values.length; i++) {
    if (String(values[i][COL_EMAIL-1]).toLowerCase() === email)
      return { sheet:sheet, found:values[i], rowIdx:i+1, values:values };
  }
  return { sheet:sheet, found:null, rowIdx:-1, values:values };
}

// ── 1. Userul cere acces ─────────────────────────────────────────────
function requestAccess_(p) {
  var email = (p.email||'').trim().toLowerCase();
  var name  = (p.name ||'').trim();
  if (!email) return { ok:false, error:'Email lipsă' };

  var r = findUser_(email);

  // Utilizator complet nou → înregistrează ca pending
  if (!r.found) {
    r.sheet.appendRow([email, name, 'pending', new Date().toISOString(), '']);
    return { ok:true, status:'pending', accessStatus:'new' };
  }

  var status = String(r.found[COL_STATUS-1]||'pending').toLowerCase();

  if (status === 'denied') {
    return { ok:true, status:'denied', accessStatus:'denied' };
  }

  if (status === 'pending') {
    r.sheet.getRange(r.rowIdx, COL_SOLICITAT).setValue(new Date().toISOString());
    return { ok:true, status:'pending', accessStatus:'pending' };
  }

  if (status === 'approved' || status === 'verified') {
    return { ok:true, status:'approved', accessStatus:'approved' };
  }

  return { ok:true, status:status, accessStatus:status };
}

// ── 2. Verificare status ──────────────────────────────────────────────
// ⚠️ MODIFICARE: user negăsit (șters din spreadsheet) → returnat ca 'denied'
// Anterior returna 'unknown' ceea ce lăsa accesul activ (fail-open)
// Acum ștergerea din spreadsheet echivalează cu revocarea accesului
function checkStatus_(p) {
  var email = (p.email||'').trim().toLowerCase();
  var r     = findUser_(email);

  // ── USER NEGĂSIT (șters din spreadsheet) → BLOCAT ──────────────────
  if (!r.found) return { ok:true, status:'denied' };

  var status = String(r.found[COL_STATUS-1]||'pending').toLowerCase();

  // Normalizăm 'verified' → 'approved' pentru HTML
  if (status === 'approved' || status === 'verified')
    return { ok:true, status:'approved' };

  return { ok:true, status:status };
}

// ── Admin panel ──────────────────────────────────────────────────────
function getPending_(p) {
  if ((p.adminEmail||'').toLowerCase() !== ADMIN_EMAIL.toLowerCase())
    return { ok:false, error:'Acces interzis' };
  var sheet  = getSheet_(SHEET_USERS);
  var values = sheet.getDataRange().getValues();
  var list   = [];
  for (var i=1; i<values.length; i++) {
    if (String(values[i][COL_STATUS-1]||'').toLowerCase() === 'pending')
      list.push({
        email:       values[i][COL_EMAIL-1],
        name:        values[i][COL_NUME-1],
        requestedAt: values[i][COL_SOLICITAT-1]
      });
  }
  return { ok:true, pending:list };
}

function approveUser_(p) {
  if ((p.adminEmail||'').toLowerCase() !== ADMIN_EMAIL.toLowerCase())
    return { ok:false, error:'Acces interzis' };
  var r = findUser_((p.targetEmail||'').toLowerCase());
  if (!r.found) return { ok:false, error:'User negăsit' };
  r.sheet.getRange(r.rowIdx, COL_STATUS).setValue('approved');
  r.sheet.getRange(r.rowIdx, COL_APROBAT).setValue(new Date().toISOString());
  return { ok:true };
}

function denyUser_(p) {
  if ((p.adminEmail||'').toLowerCase() !== ADMIN_EMAIL.toLowerCase())
    return { ok:false, error:'Acces interzis' };
  var r = findUser_((p.targetEmail||'').toLowerCase());
  if (!r.found) return { ok:false, error:'User negăsit' };
  r.sheet.getRange(r.rowIdx, COL_STATUS).setValue('denied');
  return { ok:true };
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

// ══ getPlanData ══════════════════════════════════════════════════════
function getPlanData_(sheetName) {
  if (!sheetName||!/^KW\d{2}$/i.test(sheetName))
    return {ok:false,error:'Nume foaie invalid: '+sheetName};
  try {
    var ss=SpreadsheetApp.openById(PLAN_SS_ID),name=sheetName.toUpperCase(),sheet=null;
    var all=ss.getSheets();
    for(var i=0;i<all.length;i++){if(all[i].getName().toUpperCase()===name){sheet=all[i];break;}}
    if(!sheet) return {ok:false,error:'Foaia '+name+' nu există'};
    var range=sheet.getDataRange();
    var values=range.getValues(),fw=range.getFontWeights(),fs=range.getFontStyles(),
        fz=range.getFontSizes(),ff=range.getFontFamilies();
    var keepRows=[];
    for(var r=0;r<values.length;r++){
      var empty=values[r].every(function(c){return c===''||c===null||c===undefined;});
      if(!empty)keepRows.push(r);
    }
    if(keepRows.length===0&&values.length>0)keepRows.push(0);
    var sl=function(arr){return keepRows.map(function(i){return arr[i];});};
    return {ok:true,sheet:name,data:sl(values),styles:{fontWeights:sl(fw),fontStyles:sl(fs),fontSizes:sl(fz),fontFamilies:sl(ff)}};
  }catch(e){return{ok:false,error:'Eroare citire '+sheetName+': '+e.message};}
}

// ══════════════════════════════════════════════════════════════════════
// ══ MESAJE — mesagerie unidirecțională Admin → Tehnician ════════════════
//
// FOAIA "Mesaje":
//   - Rândul 1 (HEADER) = numele EXACT al fiecărui tehnician, câte unul
//     pe coloană (trebuie să corespundă cu numele din foaia "Tehnicieni")
//   - Rândul 2 = mesajul curent pentru acel tehnician
//
//   Exemplu:
//     A                  B                  C
//   1 Alexandru Lungu     Adrian Talas       Ramon Stefana
//   2 Mergi la depou...   (gol)              Verifică stocul...
//
//   Nu se ține niciun istoric: mesajul afișat este mereu exact conținutul
//   din rândul 2. Dacă adminul șterge celula, mesajul dispare din portal
//   la următoarea sincronizare (fără a fi salvat nicăieri).
// ══════════════════════════════════════════════════════════════════════

// Citește mesajul curent (rândul 2) pentru fiecare tehnician, din foaia "Mesaje"
function getAllMessagesByTech_() {
  var sheet = getSheet_(SHEET_MESAJE);
  var values = sheet.getDataRange().getValues();
  var result = {}; // { 'Nume Tehnician': 'mesaj curent' }
  if (values.length < 1 || !values[0] || !values[0].length) return result;
  var header = values[0];
  var row2   = values[1] || [];
  header.forEach(function(name, colIdx) {
    var techName = String(name || '').trim();
    if (!techName) return;
    var txt = String(row2[colIdx] || '').trim();
    result[techName] = txt;
  });
  return result;
}


