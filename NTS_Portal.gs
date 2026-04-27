/**
 * ╔══════════════════════════════════════════════════════════════════╗
 * ║   NTS Group Portal — Google Apps Script COMPLET v3.0           ║
 * ╠══════════════════════════════════════════════════════════════════╣
 * ║  INSTRUCȚIUNI (o singură dată):                                 ║
 * ║  1. Deschide script.google.com → proiectul tău NTS             ║
 * ║  2. Șterge tot codul vechi → lipeste TOT ce e mai jos          ║
 * ║  3. Salvează (Ctrl+S)                                           ║
 * ║  4. Deploy → Manage deployments → Edit (creion) pe cel existent ║
 * ║     SAU New deployment dacă nu ai unul:                         ║
 * ║     • Type: Web App                                             ║
 * ║     • Execute as: Me                                            ║
 * ║     • Who has access: Anyone                                    ║
 * ║  5. Copiază URL-ul → pune-l în index.html la: const API_URL    ║
 * ╚══════════════════════════════════════════════════════════════════╝
 */

// ══════════════════════════════════════════════════════════════════
//  ⚙  CONFIGURARE
// ══════════════════════════════════════════════════════════════════
var CFG = {
  adminEmail:    'networkticketsolution@gmail.com',
  emailTo:       'contabilitate@networkticketsolution.com',
  sheetRapoarte: 'Rapoarte',
  sheetCert:     'Certificari',
  sheetAccess:   'AccessControl',
  otpExpireMin:  15,    // codul OTP expira dupa 15 minute
};

// ══════════════════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var p = e.parameter || {};
    if (p.payload) {
      try { merge(p, JSON.parse(decodeURIComponent(p.payload))); } catch(_) {}
    }
    return route(p);
  } catch(err) {
    return ok({ ok:false, error:err.message });
  }
}

function doPost(e) {
  try {
    var p = {};
    if (e.postData && e.postData.contents) {
      try { p = JSON.parse(e.postData.contents); } catch(_) {}
    }
    merge(p, e.parameter || {});
    if (p.payload) {
      try { merge(p, JSON.parse(decodeURIComponent(p.payload))); } catch(_) {}
    }
    return route(p);
  } catch(err) {
    return ok({ ok:false, error:err.message });
  }
}

function route(p) {
  switch (p.action) {
    // ── Date aplicație ──────────────────────────────────────────
    case 'getData':      return ok(getData());
    case 'addRaport':    return ok(addRaport(p));
    case 'saveCert':     return ok(saveCert(p));
    case 'sendEmail':    return ok(sendEmailFn(p));
    // ── Autentificare OTP ────────────────────────────────────────
    case 'requestOtp':   return ok(requestOtp(p.email, p.name));
    case 'verifyOtp':    return ok(verifyOtp(p.email, p.otp));
    case 'checkStatus':  return ok(checkStatus(p.email));
    // ── Admin control acces ──────────────────────────────────────
    case 'getPending':   return ok(getPending(p.adminEmail));
    case 'approveUser':  return ok(approveUser(p.adminEmail, p.targetEmail));
    case 'denyUser':     return ok(denyUser(p.adminEmail, p.targetEmail));
    default:
      return ok({ ok:false, error:'Actiune necunoscuta: ' + p.action });
  }
}

// ══════════════════════════════════════════════════════════════════
//  📊  getData
// ══════════════════════════════════════════════════════════════════
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Rapoarte
  var rapoarte = [];
  var shR = getOrCreate(ss, CFG.sheetRapoarte,
    ['Timestamp','Data','Contractor','Technicieni','NrSite','TipLucrare','Status','Descriere','NoteInlocuire']);
  var dR = shR.getDataRange().getValues();
  for (var i = 1; i < dR.length; i++) {
    if (!dR[i][0] && !dR[i][1]) continue;
    rapoarte.push({
      timestamp: str(dR[i][0]), data: str(dR[i][1]),
      contractor: str(dR[i][2]), technicieni: str(dR[i][3]),
      nrSite: str(dR[i][4]), tipLucrare: str(dR[i][5]),
      status: str(dR[i][6]), descriere: str(dR[i][7]),
      noteInlocuire: str(dR[i][8]),
    });
  }

  // Certificări
  var certificari = {};
  var shC = getOrCreate(ss, CFG.sheetCert, ['INI','URL','Nota']);
  var dC = shC.getDataRange().getValues();
  for (var i = 1; i < dC.length; i++) {
    var ini = str(dC[i][0]).toUpperCase();
    if (ini) certificari[ini] = { url: str(dC[i][1]), nota: str(dC[i][2]) };
  }

  // Plan sheets KW##
  var planSheets = [];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function(sh) {
    if (/^KW\d{2}$/i.test(sh.getName()))
      planSheets.push({ name: sh.getName(), gid: sh.getSheetId() });
  });
  planSheets.sort(function(a,b){ return a.name > b.name ? 1 : -1; });

  return {
    ok: true,
    data: { rapoarte:rapoarte, certificari:certificari, planSheets:planSheets, plan_url:'', pontaj_url:'' }
  };
}

// ══════════════════════════════════════════════════════════════════
//  📝  addRaport
// ══════════════════════════════════════════════════════════════════
function addRaport(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, CFG.sheetRapoarte,
    ['Timestamp','Data','Contractor','Technicieni','NrSite','TipLucrare','Status','Descriere','NoteInlocuire']);
  sh.appendRow([
    p.timestamp || new Date().toISOString(),
    p.data || '', p.contractor || '', p.technicieni || '',
    p.nrSite || '—', p.tipLucrare || '', p.status || '',
    p.descriere || '—', p.noteInlocuire || '—',
  ]);
  return { ok:true };
}

// ══════════════════════════════════════════════════════════════════
//  🏅  saveCert
// ══════════════════════════════════════════════════════════════════
function saveCert(p) {
  var ini = str(p.ini).toUpperCase();
  if (!ini) return { ok:false, error:'INI lipsa' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, CFG.sheetCert, ['INI','URL','Nota']);
  var d  = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (str(d[i][0]).toUpperCase() === ini) {
      sh.getRange(i+1,2).setValue(p.url||'');
      sh.getRange(i+1,3).setValue(p.nota||'');
      return { ok:true };
    }
  }
  sh.appendRow([ini, p.url||'', p.nota||'']);
  return { ok:true };
}

// ══════════════════════════════════════════════════════════════════
//  ✉️  sendEmail — cu semnătură și imagine atașată
// ══════════════════════════════════════════════════════════════════
function sendEmailFn(p) {
  var nume    = p.nume    || 'Necunoscut';
  var data    = p.data    || new Date().toLocaleDateString('ro-RO');
  var subiect = p.subiect || 'Solicitare';
  var mesaj   = p.mesaj   || '';
  var sigImg  = p.sigImg  || '';
  var imgData = p.imgData || '';

  var subject = '[NTS Solicitare] ' + subiect + ' — ' + nume + ' (' + data + ')';

  var sigSection = '';
  if (sigImg && sigImg.indexOf('data:image') === 0) {
    sigSection =
      '<div style="margin-top:20px;padding-top:14px;border-top:1px solid #e8e8e8;">' +
      '<p style="font-size:11px;color:#888;margin-bottom:6px;">Semnat digital de: <b>' + nume + '</b></p>' +
      '<img src="' + sigImg + '" style="max-width:220px;border:1px solid #ddd;' +
      'border-radius:6px;background:#fff;padding:6px;" alt="Semnatură">' +
      '</div>';
  }

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;color:#222;">' +
    '<div style="background:#0a1628;padding:16px 20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#00aaff;font-size:18px;margin:0;">🔧 NTS Group Portal</h2>' +
    '<p style="color:#5a7a9a;font-size:11px;margin:3px 0 0;">Portal Tehnicieni · Solicitare nouă</p></div>' +
    '<div style="background:#f9f9f9;border:1px solid #e0e0e0;border-top:none;' +
    'padding:20px 22px;border-radius:0 0 8px 8px;">' +
    '<table style="width:100%;font-size:13px;border-collapse:collapse;">' +
    '<tr><td style="padding:4px 0;color:#666;width:110px">Tehnician:</td><td style="padding:4px 0;font-weight:600">' + nume + '</td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Data:</td><td style="padding:4px 0">' + data + '</td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Subiect:</td><td style="padding:4px 0;font-weight:600;color:#0066cc">' + subiect + '</td></tr>' +
    '</table>' +
    '<div style="margin-top:14px;padding:12px 14px;background:#fff;border:1px solid #e8e8e8;' +
    'border-radius:6px;font-size:13px;line-height:1.6">' + mesaj.replace(/\n/g,'<br>') + '</div>' +
    sigSection +
    '<p style="font-size:10px;color:#bbb;margin-top:18px;text-align:right">Trimis din NTS Group Portal · ' +
    new Date().toLocaleString('ro-RO') + '</p></div></div>';

  var opts = { htmlBody:html, name:'NTS Group Portal' };

  if (imgData && imgData.indexOf('data:') === 0) {
    try {
      var m = imgData.match(/^data:([^;]+);base64,(.+)$/);
      if (m) opts.attachments = [Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'imagine.jpg')];
    } catch(e) { Logger.log('Eroare atasament: ' + e); }
  }

  GmailApp.sendEmail(CFG.emailTo, subject, '', opts);
  return { ok:true };
}

// ══════════════════════════════════════════════════════════════════
//  🔐  AUTENTIFICARE OTP
//  Fără Google Cloud Console — funcționează direct prin GmailApp
// ══════════════════════════════════════════════════════════════════

/**
 * requestOtp(email, name)
 * Verifică statusul emailului. Dacă e aprobat → generează și trimite OTP.
 * Dacă e nou → adaugă ca pending + emailează admin.
 * Returnează: { status: 'approved'|'pending'|'new'|'denied', otpSent: bool }
 */
function requestOtp(email, name) {
  if (!email) return { ok:false, error:'Email lipsa' };

  var emailLow = email.toLowerCase().trim();

  // Adminul poate intra oricând — primeste direct OTP
  if (emailLow === CFG.adminEmail.toLowerCase()) {
    var code = generateAndStoreOtp(emailLow);
    sendOtpEmail(email, name || 'Administrator', code);
    return { ok:true, status:'approved', otpSent:true };
  }

  // Verifica statusul in AccessControl
  var status = getAccessStatus(emailLow);

  if (status === 'approved') {
    var code = generateAndStoreOtp(emailLow);
    sendOtpEmail(email, name || emailLow.split('@')[0], code);
    return { ok:true, status:'approved', otpSent:true };
  }

  if (status === 'denied') {
    return { ok:true, status:'denied', otpSent:false };
  }

  if (status === 'pending') {
    return { ok:true, status:'pending', otpSent:false };
  }

  // 'unknown' — prima cerere → adauga ca pending + notifica admin
  addAccessRequest(emailLow, name || emailLow.split('@')[0]);
  notifyAdmin(email, name || emailLow.split('@')[0]);
  return { ok:true, status:'new', otpSent:false };
}

/**
 * verifyOtp(email, otp)
 * Verifică codul OTP. Returnează { ok, verified }.
 */
function verifyOtp(email, otp) {
  if (!email || !otp) return { ok:false, error:'Date lipsa', verified:false };

  var emailLow = email.toLowerCase().trim();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, CFG.sheetAccess,
    ['Email','Nume','Status','SolicitatLa','AprobatLa','OTP','OTPExpire']);

  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (str(data[i][0]).toLowerCase() !== emailLow) continue;

    var storedOtp    = str(data[i][5]);
    var otpExpire    = data[i][6] ? new Date(data[i][6]).getTime() : 0;
    var now          = Date.now();

    if (!storedOtp)           return { ok:true, verified:false, error:'Nu a fost generat niciun cod. Solicita din nou.' };
    if (now > otpExpire)      return { ok:true, verified:false, error:'Codul a expirat (15 min). Solicita un cod nou.' };
    if (str(otp) !== storedOtp) return { ok:true, verified:false, error:'Cod incorect. Mai ai ' + Math.ceil((otpExpire-now)/60000) + ' minute.' };

    // Cod corect — sterge OTP-ul
    sh.getRange(i+1, 6).setValue('');
    sh.getRange(i+1, 7).setValue('');
    return { ok:true, verified:true };
  }

  // Adminul — verifica separat (nu e in sheet)
  if (emailLow === CFG.adminEmail.toLowerCase()) {
    var prop = PropertiesService.getScriptProperties();
    var stored  = prop.getProperty('otp_admin');
    var expire  = parseInt(prop.getProperty('otp_admin_exp') || '0');
    if (!stored)                  return { ok:true, verified:false, error:'Nu a fost generat niciun cod.' };
    if (Date.now() > expire)      return { ok:true, verified:false, error:'Codul a expirat. Solicita din nou.' };
    if (str(otp) !== stored)      return { ok:true, verified:false, error:'Cod incorect.' };
    prop.deleteProperty('otp_admin');
    prop.deleteProperty('otp_admin_exp');
    return { ok:true, verified:true };
  }

  return { ok:true, verified:false, error:'Email negasit.' };
}

/**
 * checkStatus(email) — verifica statusul fara a genera OTP
 */
function checkStatus(email) {
  if (!email) return { ok:false, status:'unknown' };
  var emailLow = email.toLowerCase().trim();
  if (emailLow === CFG.adminEmail.toLowerCase()) return { ok:true, status:'approved' };
  return { ok:true, status: getAccessStatus(emailLow) };
}

// ── Helpers OTP ──────────────────────────────────────────────────

function generateAndStoreOtp(emailLow) {
  var code = String(Math.floor(100000 + Math.random() * 900000));
  var expire = new Date(Date.now() + CFG.otpExpireMin * 60 * 1000).toISOString();

  if (emailLow === CFG.adminEmail.toLowerCase()) {
    // Adminul — stocheaza in Script Properties
    var prop = PropertiesService.getScriptProperties();
    prop.setProperty('otp_admin', code);
    prop.setProperty('otp_admin_exp', String(Date.now() + CFG.otpExpireMin * 60 * 1000));
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = getOrCreate(ss, CFG.sheetAccess,
      ['Email','Nume','Status','SolicitatLa','AprobatLa','OTP','OTPExpire']);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (str(data[i][0]).toLowerCase() === emailLow) {
        sh.getRange(i+1, 6).setValue(code);
        sh.getRange(i+1, 7).setValue(expire);
        return code;
      }
    }
    // Nu exista — adauga
    sh.appendRow([emailLow, '', 'approved', '', new Date().toISOString(), code, expire]);
  }
  return code;
}

function sendOtpEmail(email, name, code) {
  try {
    GmailApp.sendEmail(email, '🔐 Codul tău de acces NTS Group Portal', '', {
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:440px;">' +
        '<div style="background:#0a1628;padding:16px 20px;border-radius:8px 8px 0 0;">' +
        '<h2 style="color:#00aaff;margin:0;font-size:17px;">🔐 NTS Group Portal</h2></div>' +
        '<div style="background:#f5f5f5;border:1px solid #ddd;border-top:none;' +
        'padding:20px 22px;border-radius:0 0 8px 8px;text-align:center;">' +
        '<p style="font-size:13px;color:#444;margin-bottom:18px;">Bună, <b>' + name + '</b>!<br>' +
        'Codul tău de acces în NTS Group Portal este:</p>' +
        '<div style="background:#fff;border:2px solid #00aaff;border-radius:10px;' +
        'padding:16px 24px;display:inline-block;margin:0 auto;">' +
        '<span style="font-size:34px;font-weight:900;letter-spacing:10px;' +
        'color:#0a1628;font-family:monospace;">' + code + '</span></div>' +
        '<p style="font-size:12px;color:#888;margin-top:14px;">' +
        'Codul este valabil <b>' + CFG.otpExpireMin + ' minute</b>.<br>' +
        'Nu îl partaja cu nimeni.</p>' +
        '<p style="font-size:10px;color:#bbb;margin-top:14px;">— NTS Group Portal</p>' +
        '</div></div>',
      name: 'NTS Group Portal'
    });
  } catch(e) {
    Logger.log('Eroare trimitere OTP: ' + e.message);
    throw new Error('Nu s-a putut trimite emailul cu codul. Verifica adresa de email.');
  }
}

function getAccessStatus(emailLow) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.sheetAccess);
  if (!sh) return 'unknown';
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (str(data[i][0]).toLowerCase() === emailLow)
      return str(data[i][2]).toLowerCase() || 'pending';
  }
  return 'unknown';
}

function addAccessRequest(emailLow, name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, CFG.sheetAccess,
    ['Email','Nume','Status','SolicitatLa','AprobatLa','OTP','OTPExpire']);
  sh.appendRow([emailLow, name, 'pending', new Date().toISOString(), '', '', '']);
}

function notifyAdmin(email, name) {
  try {
    var approveUrl = 'https://script.google.com/macros/s/AKfycbxAjOgmiqOrj-DdgYfL57cToAuMfIi4pErEyw0kNcXNif4-K_PNhp30Sxp2NgLk0R4v/exec?action=approveUser&adminEmail=' +
      encodeURIComponent(CFG.adminEmail) + '&targetEmail=' + encodeURIComponent(email);
    var denyUrl = 'https://script.google.com/macros/s/AKfycbxAjOgmiqOrj-DdgYfL57cToAuMfIi4pErEyw0kNcXNif4-K_PNhp30Sxp2NgLk0R4v/exec?action=denyUser&adminEmail=' +
      encodeURIComponent(CFG.adminEmail) + '&targetEmail=' + encodeURIComponent(email);

    GmailApp.sendEmail(CFG.adminEmail, '🔐 Cerere acces NTS Portal — ' + name, '', {
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:500px;">' +

        '<div style="background:#0a1628;padding:16px 20px;border-radius:8px 8px 0 0;">' +
        '<h2 style="color:#f0a500;margin:0;font-size:17px;">👑 Cerere acces nou — NTS Portal</h2>' +
        '</div>' +

        '<div style="background:#f9f9f9;border:1px solid #ddd;border-top:none;' +
        'padding:20px 22px;border-radius:0 0 8px 8px;">' +

        '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px;">' +
        '<tr><td style="padding:4px 0;color:#666;width:80px"><b>Nume:</b></td>' +
        '<td style="padding:4px 0;">' + name + '</td></tr>' +
        '<tr><td style="padding:4px 0;color:#666;"><b>Email:</b></td>' +
        '<td style="padding:4px 0;">' + email + '</td></tr>' +
        '<tr><td style="padding:4px 0;color:#666;"><b>Data:</b></td>' +
        '<td style="padding:4px 0;">' + new Date().toLocaleString('ro-RO') + '</td></tr>' +
        '</table>' +

        '<p style="font-size:13px;color:#333;margin-bottom:16px;">' +
        'Apasă unul din butoanele de mai jos pentru a decide accesul:</p>' +

        '<table style="width:100%;border-collapse:collapse;">' +
        '<tr>' +
        '<td style="padding-right:8px;">' +
        '<a href="' + approveUrl + '" ' +
        'style="display:block;text-align:center;padding:13px 0;' +
        'background:linear-gradient(135deg,#005533,#00cc77);' +
        'color:#fff;text-decoration:none;border-radius:8px;' +
        'font-family:Arial,sans-serif;font-size:15px;font-weight:700;' +
        'letter-spacing:1px;">✔ APROBĂ ACCESUL</a>' +
        '</td>' +
        '<td style="padding-left:8px;">' +
        '<a href="' + denyUrl + '" ' +
        'style="display:block;text-align:center;padding:13px 0;' +
        'background:#fff;color:#cc2200;text-decoration:none;' +
        'border:2px solid #cc2200;border-radius:8px;' +
        'font-family:Arial,sans-serif;font-size:15px;font-weight:700;' +
        'letter-spacing:1px;">✕ REFUZĂ</a>' +
        '</td>' +
        '</tr>' +
        '</table>' +

        '<div style="margin-top:18px;padding:11px 14px;' +
        'background:#fff3cd;border:1px solid #ffc107;border-radius:6px;' +
        'font-size:12px;color:#856404;">' +
        '⚠ Aceste linkuri funcționează doar dacă ești autentificat cu contul ' +
        '<b>' + CFG.adminEmail + '</b> în browser.' +
        '</div>' +

        '<p style="font-size:10px;color:#bbb;margin-top:14px;text-align:right;">' +
        '— NTS Group Portal</p>' +
        '</div></div>',
      name: 'NTS Group Portal'
    });
  } catch(e) { Logger.log('Notificare admin esuata: ' + e.message); }
}

// ══════════════════════════════════════════════════════════════════
//  👑  ADMIN — Gestionare acces
// ══════════════════════════════════════════════════════════════════
function getPending(adminEmail) {
  if ((adminEmail||'').toLowerCase() !== CFG.adminEmail.toLowerCase())
    return { ok:false, error:'Acces nepermis' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.sheetAccess);
  if (!sh) return { ok:true, pending:[] };

  var data = sh.getDataRange().getValues();
  var pending = [];
  for (var i = 1; i < data.length; i++) {
    if (str(data[i][2]).toLowerCase() === 'pending') {
      pending.push({ email:str(data[i][0]), name:str(data[i][1]), requestedAt:str(data[i][3]) });
    }
  }
  return { ok:true, pending:pending };
}

function approveUser(adminEmail, targetEmail) {
  if ((adminEmail||'').toLowerCase() !== CFG.adminEmail.toLowerCase())
    return { ok:false, error:'Acces nepermis' };

  if (!setStatus(targetEmail, 'approved'))
    return { ok:false, error:'Email negasit' };

  // Email de confirmare la utilizator
  try {
    GmailApp.sendEmail(targetEmail, '✅ Acces aprobat — NTS Group Portal', '', {
      htmlBody:
        '<div style="font-family:Arial,sans-serif;max-width:460px;">' +
        '<div style="background:#0a1628;padding:14px 18px;border-radius:8px 8px 0 0;">' +
        '<h2 style="color:#00cc77;margin:0;font-size:16px;">✅ Accesul tău a fost aprobat!</h2></div>' +
        '<div style="background:#f5f5f5;border:1px solid #ddd;border-top:none;' +
        'padding:18px 20px;border-radius:0 0 8px 8px;">' +
        '<p style="font-size:13px;">Bun venit în <b>NTS Group Portal</b>!</p>' +
        '<p style="font-size:13px;">Deschide aplicația, introdu emailul tău și vei primi codul de acces.</p>' +
        '<p style="font-size:10px;color:#aaa;margin-top:14px;">— NTS Group Portal</p>' +
        '</div></div>',
      name: 'NTS Group Portal'
    });
  } catch(e) { Logger.log('Email aprobare esuate: ' + e.message); }

  return { ok:true };
}

function denyUser(adminEmail, targetEmail) {
  if ((adminEmail||'').toLowerCase() !== CFG.adminEmail.toLowerCase())
    return { ok:false, error:'Acces nepermis' };

  if (!setStatus(targetEmail, 'denied'))
    return { ok:false, error:'Email negasit' };

  return { ok:true };
}

function setStatus(email, status) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.sheetAccess);
  if (!sh) return false;
  var data = sh.getDataRange().getValues();
  var emailLow = (email||'').toLowerCase().trim();
  for (var i = 1; i < data.length; i++) {
    if (str(data[i][0]).toLowerCase() === emailLow) {
      sh.getRange(i+1, 3).setValue(status);
      sh.getRange(i+1, 5).setValue(new Date().toISOString());
      return true;
    }
  }
  return false;
}

// ══════════════════════════════════════════════════════════════════
//  🔧  HELPERS GENERALI
// ══════════════════════════════════════════════════════════════════
function getOrCreate(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    var r = sh.getRange(1, 1, 1, headers.length);
    r.setValues([headers]);
    r.setFontWeight('bold');
    r.setBackground('#0a1628');
    r.setFontColor('#00aaff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function str(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v).trim();
}

function merge(target, source) {
  Object.keys(source).forEach(function(k) { target[k] = source[k]; });
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
