// ════════════════════════════════════════════════════════════════
//  NTS PORTAL — Google Apps Script Backend
//  Versiune: 2.0 | Network Ticket Solution
//
//  INSTRUCȚIUNI DE INSTALARE:
//  1. Deschide https://script.google.com
//  2. Crează un proiect nou → lipește TOT codul de mai jos
//  3. Salvează (Ctrl+S)
//  4. Deploy → New deployment → Web App
//     - Execute as: Me
//     - Who has access: Anyone
//  5. Copiază URL-ul generat → lipește-l în NTS Portal la configurare
//
//  STRUCTURA GOOGLE SHEETS (se creează automat):
//  - Foaie "Rapoarte"     → toate rapoartele zilnice trimise
//  - Foaie "Config"       → plan_url, pontaj_url
//  - Foaie "Certificari"  → linkuri certificări tehnicieni
// ════════════════════════════════════════════════════════════════

// ── ID-ul spreadsheet-ului tău ──────────────────────────────────
// IMPORTANT: Creează un Google Sheets nou, copiază ID-ul din URL
// (ex: https://docs.google.com/spreadsheets/d/ >>> ID_AICI <<< /edit)
// și înlocuiește mai jos:
const SPREADSHEET_ID = '1KE-nHKDd6iyABN7sJvlH69ac6JdVklUoF2hd_alMymg';
// ───────────────────────────────────────────────────────────────

// Numele foilor
const SHEET_RAPOARTE    = 'Rapoarte';
const SHEET_CONFIG      = 'Config';
const SHEET_CERTIFICARI = 'Certificari';

// ── HEADERS CORS ──────────────────────────────────────────────
function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── INIT: creează foile dacă nu există ───────────────────────
function initSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // Foaia Rapoarte
  if (!ss.getSheetByName(SHEET_RAPOARTE)) {
    const s = ss.insertSheet(SHEET_RAPOARTE);
    s.getRange(1,1,1,10).setValues([[
      'Timestamp','Data','Contractor','Tehnicieni',
      'Nr Site','Tip Lucrare','Status','Descriere','Note Înlocuire','ID'
    ]]);
    s.getRange(1,1,1,10).setFontWeight('bold').setBackground('#0a1628').setFontColor('#00aaff');
    s.setFrozenRows(1);
  }

  // Foaia Config
  if (!ss.getSheetByName(SHEET_CONFIG)) {
    const s = ss.insertSheet(SHEET_CONFIG);
    s.getRange(1,1,1,2).setValues([['Cheie','Valoare']]);
    s.getRange(1,1,1,2).setFontWeight('bold').setBackground('#0a1628').setFontColor('#00aaff');
    s.getRange(2,1,4,2).setValues([
      ['plan_url',         ''],
      ['pontaj_url',       ''],
      ['plan_spreadsheet_id', '1XrK6nKjp5_xA61SYodSBJExP0c7_Y5rT'],
      ['last_sync',        new Date().toISOString()]
    ]);
  }

  // Foaia Certificari
  if (!ss.getSheetByName(SHEET_CERTIFICARI)) {
    const s = ss.insertSheet(SHEET_CERTIFICARI);
    s.getRange(1,1,1,4).setValues([['Initiale','Nume','URL Certificare','Nota']]);
    s.getRange(1,1,1,4).setFontWeight('bold').setBackground('#0a1628').setFontColor('#00aaff');
    // Pre-populează cu toți tehnicienii
    const techs = [
      ['AL','Alexandru Lungu','',''],
      ['AT','Adrian Talas','',''],
      ['RS','Ramon Stefana','',''],
      ['CB','Cosmin Bizineche','',''],
      ['LC','Liviu Chelciov','',''],
      ['NB','Nicolas Blaga','',''],
      ['LG','Liviu Gaitan','',''],
      ['SB','Sergiu Bîlcu','',''],
    ];
    s.getRange(2,1,techs.length,4).setValues(techs);
  }
}

// ── GET: citire date SAU scriere via ?payload= (evita CORS) ──
function doGet(e) {
  try {
    initSheets();
    const params = e && e.parameter ? e.parameter : {};

    // Daca vine un payload JSON (POST-via-GET), proceseaza ca scriere
    if (params.payload) {
      const body = JSON.parse(decodeURIComponent(params.payload));
      const action = body.action;
      if (action === 'addRaport')  return jsonResponse(addRaport(body));
      if (action === 'saveCert')   return jsonResponse(saveCert(body));
      if (action === 'saveConfig') return jsonResponse(saveConfig(body));
      return jsonResponse({ ok: false, error: 'Actiune necunoscuta: ' + action });
    }

    // Altfel: returnare date
    const action = params.action;
    if (action === 'getData' || !action) {
      return jsonResponse({ ok: true, data: getAllData() });
    }

    return jsonResponse({ ok: true, data: getAllData() });

  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function getAllData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // === RAPOARTE ===
  const shR = ss.getSheetByName(SHEET_RAPOARTE);
  const rapoarte = [];
  if (shR && shR.getLastRow() > 1) {
    const rows = shR.getRange(2, 1, shR.getLastRow()-1, 10).getValues();
    rows.forEach(row => {
      if (row[0]) { // dacă are timestamp
        rapoarte.push({
          timestamp:     row[0] ? row[0].toString() : '',
          data:          row[1] ? Utilities.formatDate(new Date(row[1]), Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
          contractor:    row[2] || '',
          technicieni:   row[3] || '',
          nrSite:        row[4] || '—',
          tipLucrare:    row[5] || '',
          status:        row[6] || '',
          descriere:     row[7] || '',
          noteInlocuire: row[8] || '',
          id:            row[9] || '',
        });
      }
    });
  }

  // === CONFIG ===
  const shC = ss.getSheetByName(SHEET_CONFIG);
  const config = {};
  if (shC && shC.getLastRow() > 1) {
    const rows = shC.getRange(2, 1, shC.getLastRow()-1, 2).getValues();
    rows.forEach(r => { if(r[0]) config[r[0]] = r[1] || ''; });
  }

  // === CERTIFICARI ===
  const shCert = ss.getSheetByName(SHEET_CERTIFICARI);
  const certificari = {};
  if (shCert && shCert.getLastRow() > 1) {
    const rows = shCert.getRange(2, 1, shCert.getLastRow()-1, 4).getValues();
    rows.forEach(r => {
      if(r[0]) certificari[r[0]] = { nume: r[1]||'', url: r[2]||'', nota: r[3]||'' };
    });
  }

  // === PLAN SHEETS (foi din spreadsheet-ul de plan) ===
  // Citim foile din spreadsheet-ul de plan dacă e configurat
  let planSheets = [];
  const planSsId = config['plan_spreadsheet_id'] || '';
  if (planSsId) {
    try {
      const planSs = SpreadsheetApp.openById(planSsId);
      planSheets = planSs.getSheets()
        .filter(s => s.isSheetHidden() === false)
        .map(s => ({ name: s.getName(), gid: s.getSheetId() }));
    } catch(e) {
      // plan spreadsheet not accessible
    }
  }

  return {
    rapoarte,
    plan_url:    config['plan_url']   || '',
    pontaj_url:  config['pontaj_url'] || '',
    certificari,
    planSheets,
  };
}

// ── POST: procesează acțiunile ────────────────────────────────
function doPost(e) {
  try {
    initSheets();
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === 'addRaport') {
      return jsonResponse(addRaport(body));
    }
    if (action === 'saveCert') {
      return jsonResponse(saveCert(body));
    }
    if (action === 'saveConfig') {
      return jsonResponse(saveConfig(body));
    }

    return jsonResponse({ ok: false, error: 'Acțiune necunoscută: ' + action });

  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ── ADAUGĂ RAPORT ─────────────────────────────────────────────
function addRaport(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_RAPOARTE);

  const id = 'R-' + new Date().getTime();
  // Procesează data corect
  let dataVal = data.data;
  try { dataVal = new Date(data.data); } catch(e) {}

  sh.appendRow([
    data.timestamp    || new Date().toISOString(),
    dataVal,
    data.contractor   || '',
    data.technicieni  || '',
    data.nrSite       || '—',
    data.tipLucrare   || '',
    data.status       || '',
    data.descriere    || '',
    data.noteInlocuire|| '',
    id,
  ]);

  // Formatare automată rând nou
  const lastRow = sh.getLastRow();
  const isFinalizat = (data.status || '') === 'Finalizat';
  sh.getRange(lastRow, 7).setBackground(isFinalizat ? '#002a1a' : '#2a0005').setFontColor(isFinalizat ? '#00cc77' : '#ff4455');

  return { ok: true, id };
}

// ── SALVEAZĂ CERTIFICARE ──────────────────────────────────────
function saveCert(data) {
  const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh   = ss.getSheetByName(SHEET_CERTIFICARI);
  const ini  = data.ini;
  const url  = data.url  || '';
  const nota = data.nota || '';

  if (sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === ini) {
        sh.getRange(i+2, 3, 1, 2).setValues([[url, nota]]);
        return { ok: true };
      }
    }
  }
  // Nu există — adaugă rând nou
  sh.appendRow([ini, '', url, nota]);
  return { ok: true };
}

// ── SALVEAZĂ CONFIG (plan_url / pontaj_url) ───────────────────
function saveConfig(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_CONFIG);

  const keys = ['plan_url', 'pontaj_url'];
  keys.forEach(key => {
    if (data[key] !== undefined) {
      if (sh.getLastRow() > 1) {
        const vals = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
        for (let i = 0; i < vals.length; i++) {
          if (vals[i][0] === key) {
            sh.getRange(i+2, 2).setValue(data[key]);
            return;
          }
        }
      }
      sh.appendRow([key, data[key]]);
    }
  });

  return { ok: true };
}

// ── TEST manual (opțional, rulează din editor) ────────────────
function testInit() {
  initSheets();
  Logger.log('Foile au fost create cu succes!');
  Logger.log(JSON.stringify(getAllData(), null, 2));
}
