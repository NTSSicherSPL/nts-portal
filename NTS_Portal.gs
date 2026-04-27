/**
 * NTS Group Portal — Google Apps Script v3.1
 * Execute as: Me | Who has access: Anyone
 */

// ══════════════════════════════════════════════════════════════════
//  CONFIGURARE
// ══════════════════════════════════════════════════════════════════
var CFG = {
  adminEmail:    'networkticketsolution@gmail.com',
  emailTo:       'contabilitate@networkticketsolution.com',
  gasUrl:        'https://script.google.com/macros/s/AKfycbxAjOgmiqOrj-DdgYfL57cToAuMfIi4pErEyw0kNcXNif4-K_PNhp30Sxp2NgLk0R4v/exec',
  sheetRapoarte: 'Rapoarte',
  sheetCert:     'Certificari',
  sheetAccess:   'AccessControl',
  otpExpireMin:  15,
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
    return jsonOut({ ok: false, error: err.message });
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
    return jsonOut({ ok: false, error: err.message });
  }
}

function route(p) {
  switch (p.action) {
    case 'getData':      return jsonOut(getData());
    case 'addRaport':    return jsonOut(addRaport(p));
    case 'saveCert':     return jsonOut(saveCert(p));
    case 'sendEmail':    return jsonOut(sendEmailFn(p));
    case 'requestOtp':   return jsonOut(requestOtp(p.email, p.name));
    case 'verifyOtp':    return jsonOut(verifyOtp(p.email, p.otp));
    case 'checkStatus':  return jsonOut(checkStatus(p.email));
    case 'getPending':   return jsonOut(getPending(p.adminEmail));
    case 'approveUser':  return jsonOut(approveUser(p.adminEmail, p.targetEmail));
    case 'denyUser':     return jsonOut(denyUser(p.adminEmail, p.targetEmail));
    default:
      return jsonOut({ ok: false, error: 'Actiune necunoscuta: ' + p.action });
  }
}

// ══════════════════════════════════════════════════════════════════
//  getData
// ══════════════════════════════════════════════════════════════════
function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

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

  var certificari = {};
  var shC = getOrCreate(ss, CFG.sheetCert, ['INI','URL','Nota']);
  var dC = shC.getDataRange().getValues();
  for (var i = 1; i < dC.length; i++) {
    var ini = str(dC[i][0]).toUpperCase();
    if (ini) certificari[ini] = { url: str(dC[i][1]), nota: str(dC[i][2]) };
  }

  var planSheets = [];
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function(sh) {
    if (/^KW\d{2}$/i.test(sh.getName()))
      planSheets.push({ name: sh.getName(), gid: sh.getSheetId() });
  });
  planSheets.sort(function(a, b) { return a.name > b.name ? 1 : -1; });

  return { ok: true, data: { rapoarte: rapoarte, certificari: certificari, planSheets: planSheets, plan_url: '', pontaj_url: '' } };
}

// ══════════════════════════════════════════════════════════════════
//  addRaport
// ══════════════════════════════════════════════════════════════════
function addRaport(p) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, CFG.sheetRapoarte,
    ['Timestamp','Data','Contractor','Technicieni','NrSite','TipLucrare','Status','Descriere','NoteInlocuire']);
  sh.appendRow([
    p.timestamp || new Date().toISOString(),
    p.data || '', p.contractor || '', p.technicieni || '',
    p.nrSite || '-', p.tipLucrare || '', p.status || '',
    p.descriere || '-', p.noteInlocuire || '-',
  ]);
  return { ok: true };
}

// ══════════════════════════════════════════════════════════════════
//  saveCert
// ══════════════════════════════════════════════════════════════════
function saveCert(p) {
  var ini = str(p.ini).toUpperCase();
  if (!ini) return { ok: false, error: 'INI lipsa' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, CFG.sheetCert, ['INI','URL','Nota']);
  var d  = sh.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (str(d[i][0]).toUpperCase() === ini) {
      sh.getRange(i+1, 2).setValue(p.url || '');
      sh.getRange(i+1, 3).setValue(p.nota || '');
      return { ok: true };
    }
  }
  sh.appendRow([ini, p.url || '', p.nota || '']);
  return { ok: true };
}

// ══════════════════════════════════════════════════════════════════
//  sendEmail — solicitari cu semnatura si imagine
// ══════════════════════════════════════════════════════════════════
function sendEmailFn(p) {
  var nume    = p.nume    || 'Necunoscut';
  var data    = p.data    || fmtDate(new Date());
  var subiect = p.subiect || 'Solicitare';
  var mesaj   = p.mesaj   || '';
  var sigImg  = p.sigImg  || '';
  var imgData = p.imgData || '';

  var subject = '[NTS Solicitare] ' + subiect + ' - ' + nume + ' (' + data + ')';

  var sigSection = '';
  if (sigImg && sigImg.indexOf('data:image') === 0) {
    sigSection =
      '<div style="margin-top:20px;padding-top:14px;border-top:1px solid #e8e8e8;">' +
      '<p style="font-size:11px;color:#888;margin-bottom:6px;">Semnat digital de: <b>' + nume + '</b></p>' +
      '<img src="' + sigImg + '" style="max-width:220px;border:1px solid #ddd;border-radius:6px;background:#fff;padding:6px;" alt="Semnatura">' +
      '</div>';
  }

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;color:#222;">' +
    '<div style="background:#0a1628;padding:16px 20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#00aaff;font-size:18px;margin:0;">NTS Group Portal</h2>' +
    '<p style="color:#5a7a9a;font-size:11px;margin:3px 0 0;">Portal Tehnicieni - Solicitare noua</p></div>' +
    '<div style="background:#f9f9f9;border:1px solid #e0e0e0;border-top:none;padding:20px 22px;border-radius:0 0 8px 8px;">' +
    '<table style="width:100%;font-size:13px;border-collapse:collapse;">' +
    '<tr><td style="padding:4px 0;color:#666;width:110px">Tehnician:</td><td style="padding:4px 0;font-weight:600">' + nume + '</td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Data:</td><td style="padding:4px 0">' + data + '</td></tr>' +
    '<tr><td style="padding:4px 0;color:#666">Subiect:</td><td style="padding:4px 0;font-weight:600;color:#0066cc">' + subiect + '</td></tr>' +
    '</table>' +
    '<div style="margin-top:14px;padding:12px 14px;background:#fff;border:1px solid #e8e8e8;border-radius:6px;font-size:13px;line-height:1.6">' +
    mesaj.replace(/\n/g, '<br>') + '</div>' +
    sigSection +
    '<p style="font-size:10px;color:#bbb;margin-top:18px;text-align:right">Trimis din NTS Group Portal - ' + fmtDate(new Date()) + '</p>' +
    '</div></div>';

  var opts = { htmlBody: html, name: 'NTS Group Portal' };
  if (imgData && imgData.indexOf('data:') === 0) {
    try {
      var m = imgData.match(/^data:([^;]+);base64,(.+)$/);
      if (m) opts.attachments = [Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'imagine.jpg')];
    } catch(e) { Logger.log('Eroare atasament: ' + e); }
  }

  GmailApp.sendEmail(CFG.emailTo, subject, '', opts);
  return { ok: true };
}

// ══════════════════════════════════════════════════════════════════
//  AUTENTIFICARE OTP
// ══════════════════════════════════════════════════════════════════
function requestOtp(email, name) {
  if (!email) return { ok: false, error: 'Email lipsa' };

  var emailLow = email.toLowerCase().trim();
  var safeName = name || emailLow.split('@')[0];

  // Adminul — acces direct, primeste OTP
  if (emailLow === CFG.adminEmail.toLowerCase()) {
    var code = generateAndStoreOtp(emailLow);
    sendOtpEmail(email, safeName, code);
    return { ok: true, status: 'approved', otpSent: true };
  }

  var status = getAccessStatus(emailLow);

  if (status === 'approved') {
    var code = generateAndStoreOtp(emailLow);
    sendOtpEmail(email, safeName, code);
    return { ok: true, status: 'approved', otpSent: true };
  }

  if (status === 'denied') {
    return { ok: true, status: 'denied', otpSent: false };
  }

  if (status === 'pending') {
    return { ok: true, status: 'pending', otpSent: false };
  }

  // Prima cerere — adauga ca pending si trimite email admin
  addAccessRequest(emailLow, safeName);
  notifyAdmin(email, safeName);
  return { ok: true, status: 'new', otpSent: false };
}

function verifyOtp(email, otp) {
  if (!email || !otp) return { ok: false, error: 'Date lipsa', verified: false };

  var emailLow = email.toLowerCase().trim();
  var otpStr   = String(otp).trim();

  // Admin — verifica din Script Properties
  if (emailLow === CFG.adminEmail.toLowerCase()) {
    var prop    = PropertiesService.getScriptProperties();
    var stored  = prop.getProperty('otp_admin') || '';
    var expiry  = parseInt(prop.getProperty('otp_admin_exp') || '0');
    if (!stored)             return { ok: true, verified: false, error: 'Niciun cod generat. Solicita din nou.' };
    if (Date.now() > expiry) return { ok: true, verified: false, error: 'Codul a expirat. Solicita un cod nou.' };
    if (otpStr !== stored)   return { ok: true, verified: false, error: 'Cod incorect.' };
    prop.deleteProperty('otp_admin');
    prop.deleteProperty('otp_admin_exp');
    return { ok: true, verified: true };
  }

  // Utilizator normal — verifica din sheet
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(CFG.sheetAccess);
  if (!sh) return { ok: true, verified: false, error: 'Sheet AccessControl negasit.' };

  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (str(data[i][0]).toLowerCase() !== emailLow) continue;
    var stored = str(data[i][5]);
    var expiry = data[i][6] ? new Date(str(data[i][6])).getTime() : 0;
    if (!stored)             return { ok: true, verified: false, error: 'Niciun cod generat. Solicita din nou.' };
    if (Date.now() > expiry) return { ok: true, verified: false, error: 'Codul a expirat (15 min). Solicita un cod nou.' };
    if (otpStr !== stored)   return { ok: true, verified: false, error: 'Cod incorect.' };
    sh.getRange(i+1, 6).setValue('');
    sh.getRange(i+1, 7).setValue('');
    return { ok: true, verified: true };
  }

  return { ok: true, verified: false, error: 'Email negasit in lista.' };
}

function checkStatus(email) {
  if (!email) return { ok: false, status: 'unknown' };
  var emailLow = email.toLowerCase().trim();
  if (emailLow === CFG.adminEmail.toLowerCase()) return { ok: true, status: 'approved' };
  return { ok: true, status: getAccessStatus(emailLow) };
}

// ── OTP helpers ───────────────────────────────────────────────────
function generateAndStoreOtp(emailLow) {
  var code   = String(Math.floor(100000 + Math.random() * 900000));
  var expIso = new Date(Date.now() + CFG.otpExpireMin * 60 * 1000).toISOString();

  if (emailLow === CFG.adminEmail.toLowerCase()) {
    var prop = PropertiesService.getScriptProperties();
    prop.setProperty('otp_admin',     code);
    prop.setProperty('otp_admin_exp', String(Date.now() + CFG.otpExpireMin * 60 * 1000));
    return code;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = getOrCreate(ss, CFG.sheetAccess,
    ['Email','Nume','Status','SolicitatLa','AprobatLa','OTP','OTPExpire']);
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (str(data[i][0]).toLowerCase() === emailLow) {
      sh.getRange(i+1, 6).setValue(code);
      sh.getRange(i+1, 7).setValue(expIso);
      return code;
    }
  }
  // Nu exista — adauga row nou ca approved cu OTP
  sh.appendRow([emailLow, '', 'approved', '', new Date().toISOString(), code, expIso]);
  return code;
}

function sendOtpEmail(email, name, code) {
  var html =
    '<div style="font-family:Arial,sans-serif;max-width:440px;">' +
    '<div style="background:#0a1628;padding:16px 20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#00aaff;margin:0;font-size:17px;">NTS Group Portal</h2></div>' +
    '<div style="background:#f5f5f5;border:1px solid #ddd;border-top:none;padding:24px 22px;border-radius:0 0 8px 8px;text-align:center;">' +
    '<p style="font-size:14px;color:#333;margin-bottom:20px;">Salut, <b>' + name + '</b>!<br>Codul tau de acces este:</p>' +
    '<div style="background:#fff;border:2px solid #00aaff;border-radius:10px;padding:18px 28px;display:inline-block;">' +
    '<span style="font-size:36px;font-weight:900;letter-spacing:10px;color:#0a1628;font-family:monospace;">' + code + '</span></div>' +
    '<p style="font-size:12px;color:#888;margin-top:16px;">Valabil <b>' + CFG.otpExpireMin + ' minute</b>. Nu il partaja cu nimeni.</p>' +
    '<p style="font-size:10px;color:#bbb;margin-top:12px;">— NTS Group Portal</p>' +
    '</div></div>';

  GmailApp.sendEmail(email, 'Codul tau de acces NTS Group Portal', 'Codul tau: ' + code, {
    htmlBody: html,
    name: 'NTS Group Portal'
  });
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
  // URL-ul GAS — direct din CFG, fara variabile externe
  var approveUrl = CFG.gasUrl
    + '?action=approveUser'
    + '&adminEmail=' + CFG.adminEmail
    + '&targetEmail=' + email;

  var denyUrl = CFG.gasUrl
    + '?action=denyUser'
    + '&adminEmail=' + CFG.adminEmail
    + '&targetEmail=' + email;

  var dateStr = fmtDate(new Date());

  var html =
    '<div style="font-family:Arial,sans-serif;max-width:500px;">' +
    '<div style="background:#0a1628;padding:16px 20px;border-radius:8px 8px 0 0;">' +
    '<h2 style="color:#f0a500;margin:0;font-size:17px;">NTS Group Portal - Cerere acces nou</h2>' +
    '</div>' +
    '<div style="background:#f9f9f9;border:1px solid #ddd;border-top:none;padding:20px 22px;border-radius:0 0 8px 8px;">' +
    '<table style="font-size:13px;border-collapse:collapse;margin-bottom:18px;width:100%;">' +
    '<tr><td style="padding:5px 12px 5px 0;color:#666;white-space:nowrap;"><b>Nume:</b></td><td style="padding:5px 0;">' + name + '</td></tr>' +
    '<tr><td style="padding:5px 12px 5px 0;color:#666;"><b>Email:</b></td><td style="padding:5px 0;">' + email + '</td></tr>' +
    '<tr><td style="padding:5px 12px 5px 0;color:#666;"><b>Data:</b></td><td style="padding:5px 0;">' + dateStr + '</td></tr>' +
    '</table>' +
    '<p style="font-size:13px;color:#333;margin-bottom:16px;font-weight:600;">Apasa un buton pentru a decide:</p>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr>' +
    '<td style="padding-right:8px;">' +
    '<a href="' + approveUrl + '" style="display:block;text-align:center;padding:14px 0;background:#00aa55;color:#fff;text-decoration:none;border-radius:8px;font-size:15px;font-weight:bold;font-family:Arial,sans-serif;">APROBA ACCESUL</a>' +
    '</td>' +
    '<td style="padding-left:8px;">' +
    '<a href="' + denyUrl + '" style="display:block;text-align:center;padding:14px 0;background:#fff;color:#cc2200;text-decoration:none;border:2px solid #cc2200;border-radius:8px;font-size:15px;font-weight:bold;font-family:Arial,sans-serif;">REFUZA</a>' +
    '</td>' +
    '</tr>' +
    '</table>' +
    '<p style="font-size:11px;color:#aaa;margin-top:18px;text-align:center;">NTS Group Portal — ' + dateStr + '</p>' +
    '</div></div>';

  GmailApp.sendEmail(
    CFG.adminEmail,
    'Cerere acces NTS Portal: ' + name + ' (' + email + ')',
    'Cerere noua de acces de la: ' + name + ' - ' + email,
    { htmlBody: html, name: 'NTS Group Portal' }
  );

  Logger.log('Email notificare trimis pentru: ' + email);
}

// ══════════════════════════════════════════════════════════════════
//  ADMIN — Gestionare acces
// ══════════════════════════════════════════════════════════════════
function getPending(adminEmail) {
  if ((adminEmail || '').toLowerCase() !== CFG.adminEmail.toLowerCase())
    return { ok: false, error: 'Acces nepermis' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(CFG.sheetAccess);
  if (!sh) return { ok: true, pending: [] };

  var data    = sh.getDataRange().getValues();
  var pending = [];
  for (var i = 1; i < data.length; i++) {
    if (str(data[i][2]).toLowerCase() === 'pending') {
      pending.push({ email: str(data[i][0]), name: str(data[i][1]), requestedAt: str(data[i][3]) });
    }
  }
  return { ok: true, pending: pending };
}

function approveUser(adminEmail, targetEmail) {
  if ((adminEmail || '').toLowerCase() !== CFG.adminEmail.toLowerCase())
    return { ok: false, error: 'Acces nepermis' };
  if (!setStatus(targetEmail, 'approved'))
    return { ok: false, error: 'Email negasit' };

  try {
    var html =
      '<div style="font-family:Arial,sans-serif;max-width:460px;">' +
      '<div style="background:#0a1628;padding:14px 18px;border-radius:8px 8px 0 0;">' +
      '<h2 style="color:#00cc77;margin:0;font-size:16px;">Accesul tau a fost aprobat!</h2></div>' +
      '<div style="background:#f5f5f5;border:1px solid #ddd;border-top:none;padding:18px 20px;border-radius:0 0 8px 8px;">' +
      '<p style="font-size:13px;">Bun venit in <b>NTS Group Portal</b>!</p>' +
      '<p style="font-size:13px;">Deschide aplicatia, introdu emailul tau si vei primi codul de acces.</p>' +
      '<p style="font-size:10px;color:#aaa;margin-top:14px;">— NTS Group Portal</p>' +
      '</div></div>';
    GmailApp.sendEmail(targetEmail, 'Acces aprobat - NTS Group Portal', 'Accesul tau a fost aprobat!', {
      htmlBody: html, name: 'NTS Group Portal'
    });
  } catch(e) { Logger.log('Email aprobare esuata: ' + e.message); }

  return { ok: true };
}

function denyUser(adminEmail, targetEmail) {
  if ((adminEmail || '').toLowerCase() !== CFG.adminEmail.toLowerCase())
    return { ok: false, error: 'Acces nepermis' };
  if (!setStatus(targetEmail, 'denied'))
    return { ok: false, error: 'Email negasit' };
  return { ok: true };
}

function setStatus(email, status) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var sh       = ss.getSheetByName(CFG.sheetAccess);
  if (!sh) return false;
  var emailLow = (email || '').toLowerCase().trim();
  var data     = sh.getDataRange().getValues();
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
//  HELPERS
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

function fmtDate(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
}

function merge(target, source) {
  Object.keys(source).forEach(function(k) { target[k] = source[k]; });
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
