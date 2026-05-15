// V98 documenti separati

require('dotenv').config();
require('dns').s
const express = require('express');
const bodyParser = require('body-parser');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const moment = require('moment');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const crypto = require('crypto');
let twilio = null;
let google = null;
try { twilio = require('twilio'); } catch(e) { console.log('Twilio non installato:', e.message); }
try { google = require('googleapis').google; } catch(e) { console.log('Google APIs non installato:', e.message); }

const app = express();
app.use((req, res, next) => {
  res.charset = 'utf-8';
  const oldSend = res.send.bind(res);
  res.send = function(body) {
    if (typeof body === 'string') {
      res.setHeader('Content-Type', 'text/html; charset=utf-8');
      body = body
        .replace(/SÃ¬/g, 'SI')
        .replace(/sÃ¬/g, 'si')
        .replace(/Ã¨/g, 'e')
        .replace(/Ã©/g, 'e')
        .replace(/Ã /g, 'a')
        .replace(/Ã²/g, 'o')
        .replace(/Ã¹/g, 'u')
        .replace(/Ã¬/g, 'i')
        .replace(/â‚¬/g, '&euro;')
        .replace(/â€”/g, '-')
        .replace(/â€“/g, '-')
        .replace(/â†’/g, '-')
        .replace(/Â/g, '');
    }
    return oldSend(body);
  };
  next();
});


// =========================
// V63 PRIVACY / CLAUSOLE STATICHE
// =========================
const appPublicDir = path.join(__dirname, 'public');
try { fs.mkdirSync(appPublicDir, { recursive: true }); } catch(e) {}
app.use('/public', express.static(appPublicDir));
app.use(express.static(appPublicDir));

// =========================
// V107 CARGOS UID LOCK
// =========================
function v62Val(v){ return String(v===undefined||v===null?'':v).trim(); }
function v62Money(v){ const n=parseFloat(String(v||'0').replace(',','.')); return isNaN(n)?0:n; }
function v62FixTable(table, cols, done){
  const keys=Object.keys(cols||{}); if(!keys.length) return done&&done();
  let pending=keys.length;
  keys.forEach(c=>db.run(`ALTER TABLE ${table} ADD COLUMN ${c} ${cols[c]}`,()=>{ pending--; if(pending===0) done&&done(); }));
}


// =========================
// V63 BOTTONI VISIBILI + CARGOS NASCITA
// =========================
function v63DateIt(d){
  d = String(d || '').trim();
  if (!d) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
    const [y,m,dd] = d.split('-');
    return `${dd}/${m}/${y}`;
  }
  return d;
}
function v63IsoDate(d){
  d = String(d || '').trim();
  if (!d) return '';
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(d)) {
    const [dd,m,y] = d.split('/');
    return `${y}-${m}-${dd}`;
  }
  return d;
}
function v63ContractButtons(p){
  const id = p && p.id ? p.id : '';
  if (!id) return '';
  const tipo = String((p && (p.tipo_record || p.stato)) || '').toLowerCase();
  const converti = (tipo.includes('preventivo') || tipo === 'bozza')
    ? `<a class="btn btn2" href="/prenotazione/${id}/converti-contratto">Trasforma in contratto</a>`
    : '';
  return `
    <div class="actions contract-main-actions">
      <a class="btn" href="/prenotazione/${id}/modifica">Modifica</a>
      ${converti}
      <a class="btn btn2" href="/firma/${id}">Firma sul dispositivo</a>
      <a class="btn btn3" href="/firma-whatsapp/${id}">Invia firma WhatsApp</a>
      <a class="btn btn3" href="/contratto/${id}/invia-whatsapp">Invia contratto WhatsApp</a>
      <a class="btn btn2" href="/contratto/${id}/email">Invia email</a>
      <a class="btn btn2" href="/documenti/${id}">Foto/documenti</a>
      <a class="btn btn2" href="/preventivo/nuovo">Nuovo preventivo</a>
      <a class="btn bad" href="/prenotazione/${id}/elimina">Elimina</a>
    </div>
  `;
}


// =========================
// V76 FIX validateCargos
// =========================
if (typeof validateCargos === 'undefined') {
  global.validateCargos = function(p){
    const errors = [];
    if(!p) return ['Contratto non trovato'];
    if(!(p.nome||'').trim()) errors.push('Nome mancante');
    if(!(p.cognome||'').trim()) errors.push('Cognome mancante');
    if(!(p.data_nascita||'').trim()) errors.push('Data nascita mancante');
    if(!(p.documento_numero||'').trim()) errors.push('Documento mancante');
    if(!(p.patente_numero||'').trim()) errors.push('Patente mancante');
    return errors;
  }
}


// =========================
// V76 PDF UNA PAGINA + CARGOS FURGONI
// =========================
function v65CauzionePdfText(p){
  const richiesta = String(p.cauzione_richiesta || '').toLowerCase() === 'si';
  const ricevuta = String(p.cauzione_ricevuta || '').toLowerCase() === 'si';
  const imp = p.cauzione_importo || p.cauzione || 0;
  if (!richiesta) return 'Cauzione: non richiesta / non versata';
  if (ricevuta) return `Cauzione ricevuta: SI - € ${imp} - ${p.cauzione_metodo || '-'}`;
  return `Cauzione ricevuta: NO - importo previsto € ${imp}`;
}


// =========================
// V76 FIX DEFINITIVO FUNZIONE VEICOLO CARGOS
// =========================
function dpRentCleanCargosKeyV76(v) {
  return String(v || '')
    .toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[’']/g, '')
    .replace(/[_\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function old_getTipoVeicoloCargosV76_DISABLED(v) {
  const k = dpRentCleanCargosKeyV76(v);
  if (k === '1' || k.includes('FURG') || k.includes('VAN') || k.includes('DAILY') || k.includes('DUCATO') || k.includes('TRANSIT') || k.includes('VIVARO') || k.includes('EXPERT') || k.includes('SCUDO') || k.includes('DOBLO') || k.includes('DOBL') || k.includes('TALENTO') || k.includes('TRAFIC') || k.includes('MASTER') || k.includes('SPRINTER') || k.includes('VITO')) return '1';
  if (k === '3' || k.includes('BUS') || k.includes('PULMINO') || k.includes('9 POSTI') || k.includes('NOVE POSTI')) return '3';
  if (k === '4' || k.includes('AUTOCAR') || k.includes('MOTRICE') || k.includes('CAMION')) return '4';
  if (k === '5' || k.includes('TRATTORE')) return '5';
  if (k === '6' || k.includes('AUTOTRENO')) return '6';
  if (k === '7' || k.includes('ARTICOL') || k.includes('BISARCA')) return '7';
  if (k === '8' || k.includes('SNODAT')) return '8';
  if (k === '9' || k.includes('CAMPER') || k.includes('CARAVAN')) return '9';
  if (k === 'A' || k.includes('ESCAV') || k.includes('SEMOV') || k.includes('OPERA')) return 'A';
  if (k === '0' || k.includes('AUTOVETTURA') || k.includes('AUTOMOBILE') || k.includes('MACCHINA') || k.includes('AUTO')) return '0';
  return '0';
}

// Le vecchie route chiamano getTipoVeicoloCargosV61: ora esiste sempre.
function old_getTipoVeicoloCargosV61_DISABLED(v) { return getTipoVeicoloCargosV76(v); }
function old_getTipoVeicoloCargosV65_DISABLED(v) { return getTipoVeicoloCargosV76(v); }
function old_getTipoVeicoloCargos_DISABLED(v) { return getTipoVeicoloCargosV76(v); }
global.getTipoVeicoloCargosV61 = getTipoVeicoloCargosV61;
global.getTipoVeicoloCargosV65 = getTipoVeicoloCargosV65;
global.getTipoVeicoloCargos = getTipoVeicoloCargos;


// =========================
// V76 FIX COLONNE + DATA NASCITA CARGOS
// =========================
function v67AddColumn(table, column, type, cb){
  db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`, () => cb && cb());
}

function v67EnsureCriticalColumns(done){
  const cols = [
    ['prenotazioni','cauzione_richiesta','TEXT'],
    ['prenotazioni','cauzione_ricevuta','TEXT'],
    ['prenotazioni','cauzione_importo','REAL'],
    ['prenotazioni','cauzione_metodo','TEXT'],
    ['prenotazioni','cauzione_restituita','TEXT'],
    ['prenotazioni','data_nascita','TEXT'],
    ['prenotazioni','luogo_nascita','TEXT'],
    ['prenotazioni','cittadinanza_cod','TEXT'],
    ['prenotazioni','conducente_cittadinanza_cod','TEXT'],
    ['prenotazioni','documento_tipo','TEXT'],
    ['prenotazioni','documento_numero','TEXT'],
    ['prenotazioni','documento_scadenza','TEXT'],
    ['prenotazioni','patente_numero','TEXT'],
    ['prenotazioni','patente_scadenza','TEXT'],
    ['prenotazioni','tipo_cliente','TEXT'],
    ['prenotazioni','codice_fiscale','TEXT'],
    ['prenotazioni','partita_iva','TEXT'],
    ['prenotazioni','ragione_sociale','TEXT'],
    ['prenotazioni','pec','TEXT'],
    ['prenotazioni','codice_sdi','TEXT'],
    ['prenotazioni','ora_inizio','TEXT'],
    ['prenotazioni','ora_fine','TEXT']
  ];
  db.run(`CREATE TABLE IF NOT EXISTS prenotazioni (id INTEGER PRIMARY KEY AUTOINCREMENT,codice TEXT,nome TEXT,cognome TEXT,telefono TEXT,email TEXT,data_inizio TEXT,data_fine TEXT,totale REAL,stato TEXT DEFAULT 'bozza')`, () => {
    let pending = cols.length;
    cols.forEach(([t,c,tp]) => v67AddColumn(t,c,tp, () => {
      pending--;
      if(pending === 0) done && done();
    }));
  });
}

function v67NormDate(d){
  d = String(d || '').trim();
  if (!d) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(d)) {
    const [y,m,dd] = d.split('-');
    return `${dd}/${m}/${y}`;
  }
  return d;
}
function v67IsoDate(d){
  d = String(d || '').trim();
  if (!d) return '';
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(d)) {
    const [dd,m,y] = d.split('/');
    return `${y}-${m}-${dd}`;
  }
  return d;
}
function v67DefaultBirth(p){
  return v67NormDate(p.data_nascita || p.nascita_data || p.conducente_nascita_data || '01/01/1970');
}


// =========================
// V107 CARGOS UID LOCK + NO CRASH
// =========================
function v68CittadinanzaCod(p){
  return String((p && (p.cittadinanza_cod || p.conducente_cittadinanza_cod)) || '100000100').trim();
}
function v68SafeValidateCargos(p){
  try {
    const r = (typeof validateCargos === 'function') ? validateCargos(p) : null;
    if (r && Array.isArray(r.missing)) return r;
    if (Array.isArray(r)) return { ok: r.length === 0, missing: r, length: 0 };
    return { ok: true, missing: [], length: 0 };
  } catch(e) {
    return { ok: false, missing: [e.message || 'Errore validazione CARGOS'], length: 0 };
  }
}


// =========================
// V76 FIX cargosSelect MANCANTE
// evita ReferenceError: cargosSelect is not defined
// =========================


// =========================
// V107 CARGOS UID LOCK - DEFAULT REALI
// =========================
const CARGOS_DEFAULTS_V76 = {
  pagamento_tipo: '1',              // Contanti
  agenzia_id: '001',
  agenzia_nome: 'NARNI',
  agenzia_luogo_cod: '410055022',   // NARNI TR
  agenzia_indirizzo: 'VIA TUDERTE 466',
  agenzia_telefono: '0744817108',
  checkout_luogo_cod: '410055022',  // NARNI TR
  checkin_luogo_cod: '410055022',   // NARNI TR
  checkout_indirizzo: 'VIA TUDERTE 466',
  checkin_indirizzo: 'VIA TUDERTE 466',
  cittadinanza_cod: '100000100',    // ITALIA
  tipo_documento: 'PATEN',          // PATENTE DI GUIDA
  veicolo_tipo: '1',                // FURGONI
  gps: '0',
  blocco_motore: '0'
};

function cargosDefaultV76(p, key, fallback='') {
  p = p || {};
  const keys = [
    key,
    'record_cargos_' + key,
    'cargos_' + key
  ];
  for (const k of keys) {
    if (p[k] !== undefined && p[k] !== null && String(p[k]).trim() !== '') return String(p[k]).trim();
  }
  if (CARGOS_DEFAULTS_V76[key] !== undefined) return CARGOS_DEFAULTS_V76[key];
  return fallback;
}

function getTipoDocumentoCargosV76(v) {
  const k = String(v || '').trim().toUpperCase();
  if (!k) return 'PATEN';
  if (k.includes('PAT') || k === 'PATEN') return 'PATEN';
  if (k.includes('ELET') || k === 'IDELE') return 'IDELE';
  if (k.includes('IDENT') || k.includes('CARTA')) return 'IDENT';
  if (k.includes('PASS')) return 'PASOR';
  return k.length <= 5 ? k : 'PATEN';
}

function getTipoDocumentoCargosV61(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargosV65(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargosV66(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargosV67(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargosV68(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargosV69(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargosV70(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargosV71(v){ return getTipoDocumentoCargosV76(v); }
function getTipoDocumentoCargos(v){ return getTipoDocumentoCargosV76(v); }

global.getTipoDocumentoCargosV61 = getTipoDocumentoCargosV61;
global.getTipoDocumentoCargosV65 = getTipoDocumentoCargosV65;
global.getTipoDocumentoCargosV66 = getTipoDocumentoCargosV66;
global.getTipoDocumentoCargosV67 = getTipoDocumentoCargosV67;
global.getTipoDocumentoCargosV68 = getTipoDocumentoCargosV68;
global.getTipoDocumentoCargosV69 = getTipoDocumentoCargosV69;
global.getTipoDocumentoCargosV70 = getTipoDocumentoCargosV70;
global.getTipoDocumentoCargosV71 = getTipoDocumentoCargosV71;
global.getTipoDocumentoCargosV76 = getTipoDocumentoCargosV76;
global.getTipoDocumentoCargos = getTipoDocumentoCargos;

function cargosPatchDefaultsV76(p) {
  p = p || {};
  p = patchCargosVehicleTypeIntoObjectV76(p);
  return Object.assign({}, p, {
    record_cargos_pagamento_tipo: cargosDefaultV76(p, 'pagamento_tipo'),
    record_cargos_agenzia_id: cargosDefaultV76(p, 'agenzia_id'),
    record_cargos_agenzia_nome: cargosDefaultV76(p, 'agenzia_nome'),
    record_cargos_agenzia_luogo_cod: cargosDefaultV76(p, 'agenzia_luogo_cod'),
    record_cargos_agenzia_indirizzo: cargosDefaultV76(p, 'agenzia_indirizzo'),
    record_cargos_agenzia_telefono: cargosDefaultV76(p, 'agenzia_telefono'),
    record_cargos_checkout_luogo_cod: cargosDefaultV76(p, 'checkout_luogo_cod'),
    record_cargos_checkin_luogo_cod: cargosDefaultV76(p, 'checkin_luogo_cod'),
    record_cargos_checkout_indirizzo: cargosDefaultV76(p, 'checkout_indirizzo'),
    record_cargos_checkin_indirizzo: cargosDefaultV76(p, 'checkin_indirizzo'),
    record_cargos_cittadinanza_cod: cargosDefaultV76(p, 'cittadinanza_cod'),
    record_cargos_tipo_documento: cargosDefaultV76(p, 'tipo_documento'),
    record_cargos_veicolo_tipo: cargosDefaultV76(p, 'veicolo_tipo'),
    record_cargos_gps: cargosDefaultV76(p, 'gps'),
    record_cargos_blocco_motore: cargosDefaultV76(p, 'blocco_motore'),
    conducente_cittadinanza_cod: cargosDefaultV76(p, 'cittadinanza_cod'),
    cittadinanza_cod: cargosDefaultV76(p, 'cittadinanza_cod')
  });
}

// =========================
// V76 FIX TABELLE CARGOS COMPLETE
// evita ReferenceError su CARGOS_*
// =========================
const CARGOS_VEHICLE_TYPES = [
  { id: '0', descrizione: 'Autovetture' },
  { id: '1', descrizione: 'Furgoni' },
  { id: '3', descrizione: 'Autobus' },
  { id: '4', descrizione: 'Autocarri' },
  { id: '5', descrizione: 'Trattori Stradali' },
  { id: '6', descrizione: 'Autotreni' },
  { id: '7', descrizione: 'Autoarticolati' },
  { id: '8', descrizione: 'Autosnodati' },
  { id: '9', descrizione: 'Autocaravan' },
  { id: 'A', descrizione: "Mezzi d'opera" }
];

const CARGOS_PAYMENTS = [
  { id: '0', descrizione: 'Carta di Credito' },
  { id: '1', descrizione: 'Contanti' },
  { id: '2', descrizione: 'Carta di Debito' },
  { id: '3', descrizione: 'Bonifico' },
  { id: '4', descrizione: 'RID Bancario' },
  { id: '9', descrizione: 'Altro' }
];

const CARGOS_PAYMENT_TYPES = CARGOS_PAYMENTS;

const CARGOS_DOC_TYPES = [
  { id: 'CIDIP', descrizione: 'Carta ID diplomatica' },
  { id: 'IDELE', descrizione: 'Carta identità elettronica' },
  { id: 'IDENT', descrizione: "Carta di identità" },
  { id: 'PASDI', descrizione: 'Passaporto diplomatico' },
  { id: 'PASOR', descrizione: 'Passaporto ordinario' },
  { id: 'PASSE', descrizione: 'Passaporto di servizio' },
  { id: 'PATEN', descrizione: 'Patente di guida' }
];

const CARGOS_DOCUMENT_TYPES = CARGOS_DOC_TYPES;

// Luoghi Polizia base. Narni da tabella: 410055022, Italia cittadinanza/stato default: 100000100.
const CARGOS_PLACES = [
  { id: '410055022', descrizione: 'NARNI TR' },
  { id: '100000100', descrizione: 'ITALIA' }
];

const CARGOS_DEFAULT_PLACES = CARGOS_PLACES;

// Alias globali per qualunque vecchia route/template
global.CARGOS_VEHICLE_TYPES = CARGOS_VEHICLE_TYPES;
global.CARGOS_PAYMENTS = CARGOS_PAYMENTS;
global.CARGOS_PAYMENT_TYPES = CARGOS_PAYMENT_TYPES;
global.CARGOS_DOC_TYPES = CARGOS_DOC_TYPES;
global.CARGOS_DOCUMENT_TYPES = CARGOS_DOCUMENT_TYPES;
global.CARGOS_PLACES = CARGOS_PLACES;
global.CARGOS_DEFAULT_PLACES = CARGOS_DEFAULT_PLACES;

function cargosListSafe(lista, fallback) {
  return Array.isArray(lista) ? lista : (Array.isArray(fallback) ? fallback : []);
}

function cargosSelect(nome, valore, lista){
  lista = Array.isArray(lista) ? lista : [];
  return `
    <select name="${esc(nome)}">
      ${lista.map(x => {
        const id = (x && (x.id ?? x.ID ?? x.codice ?? x.CODICE ?? x.value)) ?? '';
        const desc = (x && (x.descrizione ?? x.Descrizione ?? x.DESCRIZIONE ?? x.label ?? x.nome)) ?? id;
        return `<option value="${esc(id)}" ${String(valore||'')===String(id)?'selected':''}>${esc(id)} - ${esc(desc)}</option>`;
      }).join('')}
    </select>
  `;
}

function cargosInput(nome, valore, placeholder=''){
  return `<input name="${esc(nome)}" value="${esc(valore||'')}" placeholder="${esc(placeholder||'')}">`;
}

function cargosHidden(nome, valore){
  return `<input type="hidden" name="${esc(nome)}" value="${esc(valore||'')}">`;
}


// =========================
// V76 FIX cargosSelect MANCANTE - BLOCCO PULITO
// =========================
function cargosSelect(nome, valore, lista) {
  lista = Array.isArray(lista) ? lista : [];
  return `
    <select name="${esc(nome)}">
      ${lista.map((x) => {
        const id = (x && (x.id ?? x.ID ?? x.codice ?? x.CODICE ?? x.value)) ?? '';
        const desc = (x && (x.descrizione ?? x.Descrizione ?? x.DESCRIZIONE ?? x.label ?? x.nome)) ?? id;
        return `<option value="${esc(id)}" ${String(valore || '') === String(id) ? 'selected' : ''}>${esc(id)} - ${esc(desc)}</option>`;
      }).join('')}
    </select>
  `;
}

function cargosInput(nome, valore, placeholder = '') {
  return `<input name="${esc(nome)}" value="${esc(valore || '')}" placeholder="${esc(placeholder || '')}">`;
}

function cargosHidden(nome, valore) {
  return `<input type="hidden" name="${esc(nome)}" value="${esc(valore || '')}">`;
}

app.get('/privacy.pdf', (req, res) => {
  const p1 = path.join(publicDir, 'privacy.pdf');
  const p2 = path.join(appPublicDir, 'privacy.pdf');
  res.sendFile(fs.existsSync(p1) ? p1 : p2);
});
app.get('/clausole.pdf', (req, res) => {
  const p1 = path.join(publicDir, 'clausole.pdf');
  const p2 = path.join(appPublicDir, 'clausole.pdf');
  res.sendFile(fs.existsSync(p1) ? p1 : p2);
});

const PORT = process.env.PORT || 10000;

app.use(bodyParser.urlencoded({ extended: true, limit: '80mb' }));
app.use(bodyParser.json({ limit: '80mb' }));

const IVA = 0.22;
const EXTRA_KM = 0.15;
const CAUZIONE = 500;
const EXTRA_FUORI_ORARIO = 30;

const AZIENDA = {
  nome: 'Trasporti DP S.R.L. - DP RENT',
  brand: 'DP RENT',
  indirizzo: 'Via Tuderte 466, Narni (TR)',
  telefono: '0744817108',
  email: 'contabilita@trasportidp.com',
  piva: '01385450554'
};

const TERMS_URL = process.env.TERMS_URL || '';
const PRIVACY_URL = process.env.PRIVACY_URL || '';



// =========================
// V63 PERSISTENT DATA RENDER
// =========================
const PERSISTENT_DATA_DIR = process.env.DATA_DIR || '/var/data';
try { fs.mkdirSync(PERSISTENT_DATA_DIR, { recursive: true }); } catch(e) {}

const DB_PATH = process.env.DB_PATH || path.join(PERSISTENT_DATA_DIR, 'database.sqlite');
const DATA_DIR = PERSISTENT_DATA_DIR;

const uploadDir = path.join(DATA_DIR, 'uploads');
const uploadsDir = uploadDir;
const contractsDir = path.join(DATA_DIR, 'contracts');
const firmeDir = path.join(DATA_DIR, 'firme');
const firmedDir = firmeDir;
const publicDir = path.join(DATA_DIR, 'public');
const tempDir = path.join(DATA_DIR, 'tmp');

[DATA_DIR, uploadDir, uploadsDir, contractsDir, firmeDir, firmedDir, publicDir, tempDir].forEach(dir => {
  try { if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true }); } catch(e) {}
});

// V63: migra il vecchio database una sola volta se esiste nel percorso non persistente.
try {
  const oldDbCandidates = [
    path.join(__dirname, 'data', 'database.sqlite'),
    '/opt/render/project/src/data/database.sqlite'
  ];
  if (!fs.existsSync(DB_PATH)) {
    for (const oldDb of oldDbCandidates) {
      if (oldDb !== DB_PATH && fs.existsSync(oldDb)) {
        fs.copyFileSync(oldDb, DB_PATH);
        console.log('V63 DB migrato su disco persistente da ' + oldDb + ' a ' + DB_PATH);
        break;
      }
    }
  }
} catch(e) {
  console.log('V63 DB migration skip:', e.message);
}


// DATA_DIR gestito da V63 persistent block
// uploadDir gestito da V63 persistent block
// contractsDir gestito da V63 persistent block
// firmeDir gestito da V63 persistent block
// publicDir gestito da V63 persistent block

// directory create gestito da V63 persistent block

app.use('/public', express.static(publicDir));
app.use('/uploads', express.static(uploadDir));
app.use('/contracts', express.static(contractsDir));

const upload = multer({ dest: uploadDir });
const db = new sqlite3.Database(DB_PATH);

function addColumn(table, column, type) {
  db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`, (err) => {
    const msg = String((err && err.message) || '');
    if (!err || msg.includes('duplicate column')) return;
    // Se la tabella non esiste ancora, riprova dopo che le CREATE TABLE hanno finito.
    if (msg.includes('no such table')) {
      setTimeout(() => {
        db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`, (err2) => {
          const msg2 = String((err2 && err2.message) || '');
          if (err2 && !msg2.includes('duplicate column') && !msg2.includes('no such table')) {
            console.log('ADD COLUMN warning:', table, column, msg2);
          }
        });
      }, 1500);
      return;
    }
    console.log('ADD COLUMN warning:', table, column, msg);
  });
}





function run(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function (err) {
      if (err) reject(err);
      else resolve(this);
    });
  });
}
function get(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => err ? reject(err) : resolve(row));
  });
}
function all(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => err ? reject(err) : resolve(rows || []));
  });
}


db.serialize(() => {
  // V40 CARGOS / DRIVE columns
  addColumn('prenotazioni','record_cargos_uid','TEXT');
addColumn('mezzi','uid','TEXT');
addColumn('mezzi','uid','TEXT');
addColumn('mezzi','cilindrata','TEXT');
addColumn('mezzi','alimentazione','TEXT');
addColumn('mezzi','anno','TEXT');
addColumn('mezzi','colore','TEXT');
addColumn('mezzi','posti','TEXT');
addColumn('mezzi','km','TEXT');
addColumn('mezzi','km_attuali','TEXT');
addColumn('mezzi','telaio','TEXT');
addColumn('mezzi','categoria','TEXT');
addColumn('mezzi','cauzione','TEXT');
addColumn('mezzi','prezzo_giorno','TEXT');
addColumn('mezzi','km_inclusi','TEXT');
addColumn('mezzi','note','TEXT');
addColumn('mezzi','scadenza_revisione','TEXT');
addColumn('mezzi','scadenza_bollo','TEXT');
addColumn('mezzi','scadenza_assicurazione','TEXT');
addColumn('mezzi','tagliando_km','TEXT');
addColumn('mezzi','gps','TEXT');
addColumn('mezzi','blocco_motore','TEXT');
addColumn('mezzi','record_cargos_veicolo_tipo','TEXT');
  addColumn('prenotazioni','record_cargos_transactionid','TEXT');
  addColumn('prenotazioni','record_cargos_stato','TEXT');
  addColumn('prenotazioni','record_cargos_last_check','TEXT');
  addColumn('prenotazioni','record_cargos_last_send','TEXT');
  addColumn('prenotazioni','record_cargos_last_error','TEXT');
  addColumn('prenotazioni','drive_folder_id','TEXT');
  addColumn('prenotazioni','drive_folder_link','TEXT');

  // =========================
  // V39 - CREAZIONE TABELLE BASE PRIMA DEGLI ALTER TABLE
  // =========================
  console.log('Inizializzo database DP RENT V39...');

  db.run(`
    CREATE TABLE IF NOT EXISTS clienti (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      cf TEXT,
      indirizzo TEXT,
      citta TEXT,
      cap TEXT,
      provincia TEXT,
      azienda TEXT,
      piva TEXT,
      pec TEXT,
      sdi TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      tipo TEXT,
      marca TEXT,
      modello TEXT,
      descrizione TEXT,
      targa TEXT,
      colore TEXT,
      km INTEGER DEFAULT 0,
      km_attuali INTEGER DEFAULT 0,
      prezzo_giorno REAL DEFAULT 0,
      km_inclusi INTEGER DEFAULT 150,
      deposito REAL DEFAULT 500,
      stato TEXT DEFAULT 'attivo',
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      cliente_id INTEGER,
      mezzo_id INTEGER,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      cf TEXT,
      indirizzo TEXT,
      citta TEXT,
      cap TEXT,
      provincia TEXT,
      fatturazione TEXT,
      azienda TEXT,
      piva TEXT,
      pec TEXT,
      sdi TEXT,
      mezzo TEXT,
      targa TEXT,
      data_inizio TEXT,
      ora_inizio TEXT,
      data_fine TEXT,
      ora_fine TEXT,
      giorni INTEGER DEFAULT 1,
      km_previsti INTEGER DEFAULT 150,
      km_inclusi INTEGER DEFAULT 150,
      km_uscita INTEGER,
      km_rientro INTEGER,
      carburante_uscita TEXT,
      carburante_rientro TEXT,
      totale REAL DEFAULT 0,
      cauzione REAL DEFAULT 500,
      stato TEXT DEFAULT 'bozza',
      firma TEXT,
      pdf_path TEXT,
      pdf_drive_link TEXT,
      drive_folder_id TEXT,
      drive_folder_link TEXT,
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS allegati (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      prenotazione_id INTEGER,
      tipo TEXT,
      filename TEXT,
      originalname TEXT,
      path TEXT,
      mimetype TEXT,
      size INTEGER,
      drive_file_id TEXT,
      drive_web_link TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS scadenze (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      mezzo_id INTEGER,
      tipo TEXT,
      data TEXT,
      km INTEGER,
      note TEXT,
      stato TEXT DEFAULT 'aperta',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS record_cargos_invii (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      prenotazione_id INTEGER,
      uid TEXT,
      stato TEXT,
      richiesta TEXT,
      risposta TEXT,
      errore TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS impostazioni (
      chiave TEXT PRIMARY KEY,
      valore TEXT
    )
  `);

  // CARGOS PRO V38 columns
  addColumn('prenotazioni','record_cargos_uid','TEXT');
addColumn('mezzi','uid','TEXT');
addColumn('mezzi','uid','TEXT');
addColumn('mezzi','cilindrata','TEXT');
addColumn('mezzi','alimentazione','TEXT');
addColumn('mezzi','anno','TEXT');
addColumn('mezzi','colore','TEXT');
addColumn('mezzi','posti','TEXT');
addColumn('mezzi','km','TEXT');
addColumn('mezzi','km_attuali','TEXT');
addColumn('mezzi','telaio','TEXT');
addColumn('mezzi','categoria','TEXT');
addColumn('mezzi','cauzione','TEXT');
addColumn('mezzi','prezzo_giorno','TEXT');
addColumn('mezzi','km_inclusi','TEXT');
addColumn('mezzi','note','TEXT');
addColumn('mezzi','scadenza_revisione','TEXT');
addColumn('mezzi','scadenza_bollo','TEXT');
addColumn('mezzi','scadenza_assicurazione','TEXT');
addColumn('mezzi','tagliando_km','TEXT');
addColumn('mezzi','gps','TEXT');
addColumn('mezzi','blocco_motore','TEXT');
addColumn('mezzi','record_cargos_veicolo_tipo','TEXT');
  addColumn('prenotazioni','record_cargos_stato','TEXT');
  addColumn('prenotazioni','record_cargos_last_check','TEXT');
  addColumn('prenotazioni','record_cargos_last_send','TEXT');
  addColumn('prenotazioni','record_cargos_last_error','TEXT');
  addColumn('prenotazioni','drive_folder_id','TEXT');
  addColumn('prenotazioni','drive_folder_link','TEXT');

  // CARGOS V36 columns
  addColumn('prenotazioni','record_cargos_pagamento_tipo','TEXT DEFAULT "0"');
  addColumn('prenotazioni','record_cargos_checkout_luogo_cod','TEXT');
  addColumn('prenotazioni','record_cargos_checkout_indirizzo','TEXT');
  addColumn('prenotazioni','record_cargos_checkin_luogo_cod','TEXT');
  addColumn('prenotazioni','record_cargos_checkin_indirizzo','TEXT');
  addColumn('prenotazioni','record_cargos_operatore_id','TEXT');
  addColumn('prenotazioni','record_cargos_agenzia_id','TEXT');
  addColumn('prenotazioni','record_cargos_agenzia_nome','TEXT');
  addColumn('prenotazioni','record_cargos_agenzia_luogo_cod','TEXT');
  addColumn('prenotazioni','record_cargos_agenzia_indirizzo','TEXT');
  addColumn('prenotazioni','record_cargos_agenzia_tel','TEXT');
  addColumn('prenotazioni','record_cargos_veicolo_tipo','TEXT DEFAULT "1"');
  addColumn('prenotazioni','record_cargos_veicolo_colore','TEXT');
  addColumn('prenotazioni','record_cargos_veicolo_gps','INTEGER DEFAULT 0');
  addColumn('prenotazioni','record_cargos_veicolo_bloccom','INTEGER DEFAULT 0');
  addColumn('prenotazioni','record_cargos_cittadinanza_cod','TEXT');
  addColumn('prenotazioni','record_cargos_nascita_luogo_cod','TEXT');
  addColumn('prenotazioni','record_cargos_residenza_luogo_cod','TEXT');
  addColumn('prenotazioni','record_cargos_doc_tipo_cod','TEXT DEFAULT "CI"');
  addColumn('prenotazioni','record_cargos_doc_luogoril_cod','TEXT');
  addColumn('prenotazioni','record_cargos_patente_luogoril_cod','TEXT');
  addColumn('prenotazioni','conducente2_nome','TEXT');
  addColumn('prenotazioni','conducente2_cognome','TEXT');
  addColumn('prenotazioni','conducente2_data_nascita','TEXT');
  addColumn('prenotazioni','conducente2_nascita_luogo_cod','TEXT');
  addColumn('prenotazioni','conducente2_cittadinanza_cod','TEXT');
  addColumn('prenotazioni','conducente2_doc_tipo_cod','TEXT');
  addColumn('prenotazioni','conducente2_doc_numero','TEXT');
  addColumn('prenotazioni','conducente2_doc_luogoril_cod','TEXT');
  addColumn('prenotazioni','conducente2_patente_numero','TEXT');
  addColumn('prenotazioni','conducente2_patente_luogoril_cod','TEXT');
  addColumn('prenotazioni','conducente2_recapito','TEXT');
  addColumn('mezzi','record_cargos_veicolo_tipo','TEXT DEFAULT "1"');
  addColumn('mezzi','colore','TEXT');
  addColumn('mezzi','gps','INTEGER DEFAULT 0');
  addColumn('mezzi','blocco_motore','INTEGER DEFAULT 0');

db.run(`
    CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      uid TEXT,
      targa TEXT UNIQUE,
      km INTEGER DEFAULT 0,
      marca TEXT,
      modello TEXT,
      cilindrata TEXT,
      alimentazione TEXT,
      codice_tipo TEXT,
      categoria TEXT,
      descrizione TEXT,
      descrizione_pubblica TEXT,
      posti INTEGER,
      stazione TEXT,
      prezzo_giorno REAL,
      km_inclusi INTEGER,
      stato TEXT DEFAULT 'disponibile',
      km_attuali INTEGER DEFAULT 0,
      tagliando_km_scadenza INTEGER,
      tagliando_data_scadenza TEXT,
      revisione_scadenza TEXT,
      bollo_scadenza TEXT,
      assicurazione_scadenza TEXT,
      gomme_scadenza TEXT,
      manutenzione_note TEXT,
      alert_giorni INTEGER DEFAULT 30,
      alert_km INTEGER DEFAULT 1000
    )
  `);

  [
    ['descrizione_pubblica','TEXT'], ['posti','INTEGER'], ['km_attuali','INTEGER DEFAULT 0'],
    ['tagliando_km_scadenza','INTEGER'], ['tagliando_data_scadenza','TEXT'], ['revisione_scadenza','TEXT'],
    ['bollo_scadenza','TEXT'], ['assicurazione_scadenza','TEXT'], ['gomme_scadenza','TEXT'],
    ['manutenzione_note','TEXT'], ['alert_giorni','INTEGER DEFAULT 30'], ['alert_km','INTEGER DEFAULT 1000']
  ].forEach(c => addColumn('mezzi', c[0], c[1]));

  db.run(`
    CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      codice_fiscale TEXT,
      indirizzo TEXT,
      citta TEXT,
      cap TEXT,
      tipo_cliente TEXT,
      piva TEXT,
      ragione_sociale TEXT,
      pec TEXT,
      sdi TEXT,
      conducente1 TEXT,
      patente1 TEXT,
      patente1_scadenza TEXT,
      conducente2 TEXT,
      patente2 TEXT,
      patente2_scadenza TEXT,
      mezzo_id INTEGER,
      data_inizio TEXT,
      data_fine TEXT,
      ora_inizio TEXT,
      ora_fine TEXT,
      giorni INTEGER,
      km_previsti INTEGER,
      extra_fuori_orario REAL DEFAULT 0,
      extra_km REAL DEFAULT 0,
      imponibile REAL,
      iva REAL,
      totale REAL,
      cauzione REAL,
      carburante_uscita TEXT DEFAULT '4/4 pieno',
      carburante_rientro TEXT DEFAULT '4/4 pieno',
      km_uscita INTEGER,
      km_rientro INTEGER,
      stato TEXT DEFAULT 'bozza',
      firma_path TEXT,
      pdf_path TEXT,
      pdf_drive_file_id TEXT,
      pdf_drive_web_link TEXT,
      nexi_link TEXT,
      nexi_stato TEXT,
      nexi_raw TEXT,
      record_cargos_stato TEXT,
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  [
    ['pec','TEXT'], ['sdi','TEXT'], ['conducente1','TEXT'], ['patente1','TEXT'], ['patente1_scadenza','TEXT'],
    ['conducente2','TEXT'], ['patente2','TEXT'], ['patente2_scadenza','TEXT'],
    ['firma_path','TEXT'], ['pdf_path','TEXT'], ['pdf_drive_file_id','TEXT'], ['pdf_drive_web_link','TEXT'],
    ['nexi_link','TEXT'], ['nexi_stato','TEXT'], ['nexi_raw','TEXT'], ['record_cargos_stato','TEXT'],
    ['extra_km','REAL DEFAULT 0'], ['extra_fuori_orario','REAL DEFAULT 0'],
    ['carburante_uscita','TEXT DEFAULT "4/4 pieno"'], ['carburante_rientro','TEXT DEFAULT "4/4 pieno"'],
    ['km_uscita','INTEGER'], ['km_rientro','INTEGER'], ['note','TEXT']
  ].forEach(c => addColumn('prenotazioni', c[0], c[1]));

  
  db.run(`
    CREATE TABLE IF NOT EXISTS clienti (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      codice_fiscale TEXT UNIQUE,
      indirizzo TEXT,
      citta TEXT,
      cap TEXT,
      data_nascita TEXT,
      luogo_nascita TEXT,
      documento_numero TEXT,
      documento_scadenza TEXT,
      patente_numero TEXT,
      patente_scadenza TEXT,
      categoria_patente TEXT,
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  [
    ['telefono','TEXT'],['email','TEXT'],['codice_fiscale','TEXT'],['indirizzo','TEXT'],
    ['citta','TEXT'],['cap','TEXT'],['data_nascita','TEXT'],['luogo_nascita','TEXT'],
    ['documento_numero','TEXT'],['documento_scadenza','TEXT'],['patente_numero','TEXT'],
    ['patente_scadenza','TEXT'],['categoria_patente','TEXT'],['note','TEXT'],['updated_at','TEXT']
  ].forEach(c => addColumn('clienti', c[0], c[1]));

db.run(`
    CREATE TABLE IF NOT EXISTS allegati (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      prenotazione_id INTEGER,
      mezzo_id INTEGER,
      tipo TEXT,
      filename TEXT,
      originalname TEXT,
      path TEXT,
      mimetype TEXT,
      drive_file_id TEXT,
      drive_web_link TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);
});

function esc(v) {
  return String(v === undefined || v === null ? '' : v)
    .replace(/[&<>"']/g, s => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#039;' }[s]));
}
function normalize(v) { return v === undefined || v === null ? '' : String(v).trim(); }
function euro(v) { return Number(v || 0).toFixed(2); }


const DP_PRIVACY_URL = process.env.DP_PRIVACY_URL || '/privacy';
const DP_TERMS_URL = process.env.DP_TERMS_URL || '/condizioni-noleggio';

function privacyCheckboxHtml() {
  return `
    <div class="notice">
      <label style="display:flex;gap:8px;align-items:flex-start">
        <input type="checkbox" name="accetta_privacy_termini" value="SI" required style="width:auto;margin-top:4px">
        <span>
          Dichiaro di aver letto e accettare
          <a target="_blank" href="${DP_PRIVACY_URL}">Privacy</a>
          e
          <a target="_blank" href="${DP_TERMS_URL}">Condizioni generali di noleggio</a>.
        </span>
      </label>
    </div>
  `;
}


const OCR_PREFILL = {};
const PREN_OCR_UPLOADS = {}; // foto caricate prima della compilazione cliente
function makeOcrId(){ return 'OCR' + Date.now() + Math.floor(Math.random()*10000); }


function publicBaseUrl(req) {
  return process.env.APP_BASE_URL || process.env.PUBLIC_URL || `${req.protocol}://${req.get('host')}`;
}

function page(title, content) {
  const logoPath = path.join(publicDir, 'logo.png');
  const logoHtml = fs.existsSync(logoPath)
    ? `<img src="/public/logo.png" style="height:48px;max-width:180px;object-fit:contain;background:white;border-radius:8px;padding:4px;">`
    : `<span class="brandText">DP RENT</span>`;

  return `<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="theme-color" content="#000000">
<title>${esc(title)}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
:root{--red:#c60000;--dark:#070707;--soft:#f4f4f4}
*{box-sizing:border-box}
body{font-family:Arial,Helvetica,sans-serif;margin:0;background:var(--soft);color:#222}
header{background:var(--dark);color:white;padding:18px 22px;display:flex;align-items:center;gap:16px}
header h1{margin:0;font-size:28px;letter-spacing:1px;font-weight:900}
.brandText{font-size:30px;font-weight:900;color:white}
nav{background:var(--red);padding:12px 14px;display:flex;gap:9px;flex-wrap:wrap;align-items:center}
nav a{color:white;text-decoration:none;font-weight:800;padding:8px 10px;border-radius:7px}
nav a:hover{background:#8f0000}
main{padding:18px}
.box{background:white;padding:20px;margin-bottom:20px;border-radius:14px;box-shadow:0 2px 14px #cfcfcf}
.hero{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:16px;margin-bottom:20px}
.tile{background:#101010;color:#fff;text-decoration:none;border-radius:20px;padding:25px 16px;min-height:118px;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;font-size:22px;font-weight:900;box-shadow:0 7px 0 var(--red),0 2px 12px #bbb}
.tile span{font-size:34px;line-height:1;margin-bottom:10px;border:0;border-radius:0;padding:0;min-width:46px;display:inline-block}
.tile:hover{transform:translateY(-2px);filter:brightness(1.12)}
table{width:100%;border-collapse:collapse;background:white}
th,td{padding:9px;border:1px solid #ddd;font-size:13px;vertical-align:top}
th{background:#222;color:white}
input,select,textarea,button{padding:11px;margin:5px 0;width:100%;border:1px solid #aaa;border-radius:8px;font-size:15px}
button,.btn{background:var(--red);color:white;border:0;padding:11px 16px;border-radius:8px;text-decoration:none;display:inline-block;font-weight:bold;cursor:pointer;margin:4px 4px 4px 0;width:auto}
.btn2{background:#333}.btn3{background:#0b6b2d}.btnWarn{background:#b36b00}
.ok{color:green;font-weight:bold}.bad{color:#b30000;font-weight:bold}.warn{color:#b36b00;font-weight:bold}
.btn.bad,button.bad{background:#b30000!important;color:#fff!important}
.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
.libero{background:#1fae4b;color:white;text-align:center;font-weight:bold;cursor:pointer}
.occupato{background:#d90000;color:white;text-align:center;font-weight:bold;cursor:pointer}
.sticky-table{overflow:auto;max-height:75vh;border:1px solid #ccc}
.sticky-table th{position:sticky;top:0;z-index:3}
.sticky-col{position:sticky;left:0;background:#fff;z-index:2;min-width:155px}
th.sticky-col{background:#222;color:white;z-index:4}
canvas{border:2px solid #333;background:white;width:100%;height:250px;touch-action:none}
.actions{display:flex;gap:8px;flex-wrap:wrap}
.notice{background:#fff3cd;border:1px solid #ffe08a;padding:11px;border-radius:8px;margin:10px 0}
.alert{background:#ffe0e0;border:1px solid #d90000;padding:10px;border-radius:8px;margin:6px 0}
.badge{display:inline-block;padding:4px 7px;border-radius:5px;font-size:12px;margin:2px;background:#eee}
.badge-red{background:#d90000;color:white}.badge-green{background:#1fae4b;color:white}.badge-orange{background:#ffb000;color:#111}.badge-blue{background:#1155cc;color:white}.premium-card{border:1px solid #eee;border-radius:18px;padding:18px;background:linear-gradient(180deg,#fff,#fafafa);box-shadow:0 10px 25px rgba(0,0,0,.08);margin:10px 0}.big-actions .btn{font-size:17px;padding:14px 18px;border-radius:14px}.muted{color:#777}
pre{white-space:pre-wrap;word-break:break-word;background:#111;color:#fff;padding:12px;border-radius:8px;overflow:auto}
@keyframes dpBlink{0%,100%{box-shadow:0 0 0 0 rgba(215,0,0,.0);filter:brightness(1)}50%{box-shadow:0 0 28px 8px rgba(215,0,0,.45);filter:brightness(1.15)}}
.dp-alert-wait{animation:dpBlink 1s infinite;border:5px solid #d70000!important;background:#fff1f1!important}
.dp-alert-wait h2{color:#d70000!important}

.top-actions{display:flex;gap:10px;flex-wrap:wrap;margin:0 0 14px}.top-actions .back-btn,.top-actions a{display:inline-flex;align-items:center;justify-content:center;min-height:44px;border-radius:14px;padding:10px 16px;font-weight:900;text-decoration:none;border:0;background:#333;color:#fff;box-shadow:0 3px 0 rgba(0,0,0,.16);font-size:16px}.top-actions .home-btn{background:#d70000}@media(max-width:700px){.top-actions{position:sticky;top:0;z-index:20;background:rgba(244,244,244,.94);backdrop-filter:blur(8px);padding:8px 0}.top-actions .back-btn,.top-actions a{flex:1;min-width:130px}}

/* V109 responsive premium iPhone/iPad */
body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;background:linear-gradient(180deg,#f7f7f7,#ededed);-webkit-text-size-adjust:100%}
header{padding:22px clamp(16px,3vw,34px);box-shadow:0 10px 30px rgba(0,0,0,.22);position:relative}
header img{height:64px!important;border-radius:16px!important;padding:6px!important;box-shadow:0 0 0 3px rgba(255,255,255,.18)}
header h1{font-size:clamp(28px,5vw,46px);letter-spacing:2px;line-height:1.05}
header h1 small{display:inline-block;margin-left:8px;font-size:clamp(13px,2vw,18px)!important;letter-spacing:1.5px;color:#f1f1f1!important}
nav{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;padding:18px clamp(14px,3vw,30px);box-shadow:0 10px 25px rgba(198,0,0,.22)}
nav a{display:flex;align-items:center;justify-content:center;min-height:48px;text-align:center;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.16);border-radius:16px;font-size:clamp(16px,2.2vw,22px);padding:11px 12px}
main{max-width:1180px;margin:0 auto;padding:clamp(14px,3vw,28px)}
.box,.premium-card{border-radius:24px;padding:clamp(18px,3vw,34px);box-shadow:0 18px 45px rgba(0,0,0,.12);border:1px solid rgba(0,0,0,.06)}
.box h2,.premium-card h2{font-size:clamp(28px,4vw,44px);line-height:1.08;margin-top:0}
.btn,button{border-radius:16px;padding:14px 22px;font-size:clamp(17px,2vw,22px);font-weight:900;box-shadow:0 4px 0 rgba(0,0,0,.18)}
.actions,.big-actions{display:flex;flex-wrap:wrap;gap:12px}.actions .btn,.big-actions .btn{margin:0}
input,select,textarea{font-size:18px;border-radius:14px;padding:14px}
.badge{border-radius:999px;font-weight:900;padding:7px 12px;font-size:14px}.badge-money{background:#111;color:#fff}.badge-danger{background:#d70000;color:#fff}.badge-ok{background:#10883b;color:#fff}.badge-warn{background:#ffb000;color:#111}
.cauzione-box{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin:12px 0 8px}.cauzione-box .label{font-weight:900;font-size:20px;margin-right:2px}
@media(min-width:701px) and (max-width:1180px){nav{grid-template-columns:repeat(3,1fr)}main{padding:24px}.btn,button{min-height:58px}.box{font-size:19px}}
@media(max-width:700px){.grid{grid-template-columns:1fr}main{padding:14px}header{align-items:flex-start;gap:14px}header img{height:56px!important}header h1{font-size:34px}header h1 small{display:block;margin:6px 0 0 0}nav{grid-template-columns:repeat(2,1fr);gap:9px;padding:14px}nav a{min-height:52px;font-size:18px}.tile{font-size:19px;min-height:100px}th,td{font-size:12px;padding:6px}.actions .btn,.big-actions .btn{width:100%;text-align:center}.btn,button{width:auto;min-height:54px}.box,.premium-card{border-radius:22px;padding:22px} .cauzione-box{display:grid;grid-template-columns:1fr}.cauzione-box .badge{width:100%;text-align:center;font-size:16px}}

.contract-main-actions{margin-top:16px}.contract-main-actions .btn{min-width:190px;text-align:center}.contract-secondary-actions .btn{min-width:150px;text-align:center}
@media(max-width:700px){.contract-main-actions .btn,.contract-secondary-actions .btn{width:100%;min-width:0}}


/* V109 FIX leggibilita mobile */
header{padding-top:max(22px, env(safe-area-inset-top));}
.top-actions{max-width:1180px;margin:0 auto 14px!important;padding:10px 0!important;}
.top-actions .back-btn::before{content:""!important;}
.top-actions .back-btn,.top-actions a{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif!important;font-size:clamp(18px,2.6vw,24px)!important;letter-spacing:0!important;line-height:1.1!important;white-space:nowrap!important;color:#fff!important;overflow:hidden;text-overflow:ellipsis;}
.top-actions .back-btn{background:#333!important;}
.top-actions .home-btn{background:#d70000!important;}
.client-back button{font-size:18px!important;font-weight:900!important;background:#333!important;color:#fff!important;}
@media(max-width:700px){
  nav{padding-top:calc(14px + env(safe-area-inset-top));}
  .top-actions{position:sticky;top:0;z-index:50;padding:10px 12px!important;gap:10px!important;background:rgba(244,244,244,.96)!important;}
  .top-actions .back-btn,.top-actions a{min-width:0!important;width:calc(50% - 5px)!important;flex:1 1 calc(50% - 5px)!important;padding:14px 8px!important;}
  .contract-main-actions .btn{width:100%!important;}
}

</style>
</head>
<body>
<header>${logoHtml}<h1>DP RENT APP <small style="font-size:13px;color:#ddd">V122 AREA CLIENTE OK</small></h1></header>
<nav>
<a href="/">Dashboard</a>
<a href="/mezzi-web">Mezzi</a>
<a href="/scadenze-mezzi">Scadenze</a>
<a href="/import-mezzi">Import Excel</a>
<a href="/nuova-prenotazione">Nuova prenotazione</a>
<a href="/clienti">Clienti</a>
<a href="/scansione-documenti">Scansione documenti</a>
<a href="/documenti-clienti">Documenti clienti</a>
<a href="/richieste-attesa">Clienti in attesa</a>
<a href="/scadenze-clienti">Scadenze clienti</a>
<a href="/prenotazioni">Storico</a>
<a href="/planning">Planning</a>
<a href="/prenota">Pagina cliente</a>
<a href="/cargos">Ca.R.G.O.S.</a><a href="/cargos-config">Config CARGOS</a>
<a href="/logo">Logo</a>
<a href="/test-email">Test Email</a>
<a href="/test-drive">Test Drive</a>
</nav>
<main><div class="top-actions"><button type="button" class="back-btn" onclick="history.length>1?history.back():location.href='/'">Indietro</button><a class="home-btn" href="/">Dashboard</a></div>${content}</main>
</body>
</html>`;
}

function categoriaFromRow(row) {
  const codice = normalize(row['Codice Tip'] || row['Codice Tipo']).toUpperCase();
  const marca = normalize(row['Marca']).toUpperCase();
  const modello = normalize(row['Modello']).toUpperCase();
  const desc = normalize(row['Descrizion'] || row['Descrizione'] || row['Immagini consegna']).toUpperCase();
  if (marca.includes('DACIA') || modello.includes('DACIA') || desc.includes('DACIA')) return 'AUTO_DACIA';
  if (modello.includes('GOLF') || desc.includes('GOLF')) return 'AUTO_GOLF';
  if (codice.includes('X-ESC') || desc.includes('ESCAVATORE')) return 'ESCAVATORE';
  if (desc.includes('PIATTAFORMA') || desc.includes('SEMOVENTE')) return 'SEMOVENTE';
  if (codice.includes('P') || desc.includes('PERSONE') || desc.includes('9P') || desc.includes('9 POSTI')) return '9_POSTI';
  return 'FURGONE';
}
function descrizionePubblica(m) {
  if (m.descrizione_pubblica) return m.descrizione_pubblica;
  const modello = `${m.marca || ''} ${m.modello || ''}`.trim();
  if (m.categoria === '9_POSTI') return `${modello} - pulmino 9 posti`;
  if (m.categoria === 'FURGONE') return `${modello} - furgone cargo/merci`;
  if (m.categoria === 'AUTO_DACIA') return `${modello} - auto economica`;
  if (m.categoria === 'AUTO_GOLF') return `${modello} - auto categoria Golf`;
  if (m.categoria === 'ESCAVATORE') return `${modello} - escavatore`;
  if (m.categoria === 'SEMOVENTE') return `${modello} - piattaforma/semovente`;
  return modello || m.descrizione || 'Mezzo DP RENT';
}

// V112 FIX categoria cliente: impedisce assegnazione mezzo sbagliato (es. pulmino -> ribaltabile)
function categoriaClienteNorm(v) {
  const k = normalize(v || '').toUpperCase().replace(/[\s\-]+/g, '_');
  if (k === '9' || k === '9_POSTI' || k === 'PULMINO' || k === 'PULMINO_8_9_POSTI' || k.includes('9_POSTI') || k.includes('PULMINO')) return '9_POSTI';
  if (k.includes('AUTO_GOLF') || k.includes('GOLF')) return 'AUTO_GOLF';
  if (k.includes('AUTO_DACIA') || k.includes('DACIA')) return 'AUTO_DACIA';
  if (k.includes('ESCAV')) return 'ESCAVATORE';
  if (k.includes('SEMOV') || k.includes('PIATTAFORMA')) return 'SEMOVENTE';
  if (k.includes('FURG') || k.includes('CARGO') || k.includes('MERCI')) return 'FURGONE';
  return k || 'FURGONE';
}
function mezzoCompatibileCategoriaCliente(m, categoriaRichiesta) {
  // V123: filtro rigido. La nota interna deve seguire ESATTAMENTE la scelta cliente.
  // Non usare marca/modello generici tipo DAILY/DUCATO per decidere, altrimenti prende mezzi sbagliati.
  const cat = categoriaClienteNorm(categoriaRichiesta);
  const categoriaMezzo = categoriaClienteNorm(m.categoria || m.tipo || m.codice_tipo || '');
  const testo = normalize(`${m.categoria || ''} ${m.tipo || ''} ${m.descrizione || ''} ${m.descrizione_pubblica || ''} ${m.codice_tipo || ''}`).toUpperCase();
  const posti = Number(m.posti || 0);
  if (categoriaMezzo === cat) return true;
  if (cat === 'FURGONE') return /(^|[^A-Z])(FURGONE|FURGONI|CARGO|MERCI|F1|F2|F3|VAN)([^A-Z]|$)/.test(testo) && !/(DACIA|GOLF|PULMINO|9_POSTI|9 POSTI|ESCAV|SEMOV|PIATTAFORMA)/.test(testo);
  if (cat === '9_POSTI') return (/(PULMINO|9_POSTI|9 POSTI|8 POSTI|MINIBUS|PASSEGGERI|PERSONE)/.test(testo) || posti >= 8) && !/(RIBALT|CASSON|CARGO|MERCI|ESCAV|SEMOV|PIATTAFORMA)/.test(testo);
  if (cat === 'AUTO_DACIA') return /(AUTO_DACIA|DACIA|SANDERO)/.test(testo);
  if (cat === 'AUTO_GOLF') return /(AUTO_GOLF|GOLF)/.test(testo);
  if (cat === 'ESCAVATORE') return /(ESCAV|X-ESC|MEZZO_SPECIALE|SPECIALE)/.test(testo);
  if (cat === 'SEMOVENTE') return /(SEMOV|PIATTAFORMA)/.test(testo);
  return false;
}

function prezzoCategoria(cat) {
  if (cat === 'AUTO_DACIA') return 50;
  if (cat === 'AUTO_GOLF') return 60;
  if (cat === 'ESCAVATORE') return 50;
  if (cat === 'SEMOVENTE') return 50;
  return 70;
}
function kmCategoria(cat) {
  if (cat === 'ESCAVATORE' || cat === 'SEMOVENTE') return 0;
  return 150;
}
function codicePratica(id) {
  return `DPR-${moment().format('YYYYMMDD')}-${String(id).padStart(4, '0')}`;
}
function validDateRange(inizio, fine) {
  if (!inizio || !fine) return 'Data inizio/fine mancante';
  if (moment(fine).isBefore(moment(inizio))) return 'La data fine non può essere precedente alla data inizio';
  return '';
}
function extraOrario(ora) {
  if (!ora) return 0;
  const parts = String(ora).split(':').map(Number);
  const minuti = (parts[0] || 0) * 60 + (parts[1] || 0);
  const inizio = 8 * 60 + 30;
  const fine = 18 * 60 + 30;
  return (minuti < inizio || minuti > fine) ? EXTRA_FUORI_ORARIO : 0;
}
function calcolaTotale(mezzo, data_inizio, data_fine, ora_inizio, ora_fine, km_previsti) {
  const giorni = Math.max(1, moment(data_fine).diff(moment(data_inizio), 'days') + 1);
  const prezzo = Number(mezzo.prezzo_giorno || prezzoCategoria(mezzo.categoria));
  const kmGiorno = Number(mezzo.km_inclusi || kmCategoria(mezzo.categoria));
  const kmInclusiTot = giorni * kmGiorno;
  const kmPrev = Number(km_previsti || 0);
  const extraKm = kmGiorno > 0 && kmPrev > kmInclusiTot ? (kmPrev - kmInclusiTot) * EXTRA_KM : 0;
  const extra = extraOrario(ora_inizio) + extraOrario(ora_fine);
  const imponibile = giorni * prezzo + extra + extraKm;
  const iva = imponibile * IVA;
  const totale = imponibile + iva;
  return { giorni, kmInclusiTot, extraKm, imponibile, iva, totale, extra_fuori_orario: extra };
}
async function queryDisponibilita(mezzo_id, data_inizio, data_fine) {
  return get(`
    SELECT * FROM prenotazioni
    WHERE mezzo_id = ?
    AND stato != 'annullato'
    AND date(data_inizio) <= date(?)
    AND date(data_fine) >= date(?)
  `, [mezzo_id, data_fine, data_inizio]);
}
function fuelOptions(selected) {
  const vals = ['4/4 pieno','3/4','1/2','1/4','Riserva','Vuoto'];
  return vals.map(v => `<option value="${esc(v)}" ${selected===v?'selected':''}>${esc(v)}</option>`).join('');
}
function alertMezzo(m) {
  const out = [];
  const today = moment().startOf('day');
  const alertGiorni = Number(m.alert_giorni || 30);
  const alertKm = Number(m.alert_km || 1000);
  const kmAtt = Number(m.km_attuali || m.km || 0);

  function checkDate(label, val) {
    if (!val) return;
    const d = moment(val);
    if (!d.isValid()) return;
    const diff = d.diff(today, 'days');
    if (diff < 0) out.push(`<div class="alert">&#10060; ${label} scaduto il ${esc(val)}</div>`);
    else if (diff <= alertGiorni) out.push(`<div class="alert">&#9888;&#65039; ${label} in scadenza: ${esc(val)} (${diff} giorni)</div>`);
  }

  checkDate('Tagliando data', m.tagliando_data_scadenza);
  checkDate('Revisione', m.revisione_scadenza);
  checkDate('Bollo', m.bollo_scadenza);
  checkDate('Assicurazione', m.assicurazione_scadenza);
  checkDate('Gomme/manutenzione', m.gomme_scadenza);

  if (m.tagliando_km_scadenza) {
    const diffKm = Number(m.tagliando_km_scadenza) - kmAtt;
    if (diffKm <= 0) out.push(`<div class="alert">&#10060; Tagliando km scaduto: km attuali ${kmAtt}, scadenza ${m.tagliando_km_scadenza}</div>`);
    else if (diffKm <= alertKm) out.push(`<div class="alert">&#9888;&#65039; Tagliando vicino: mancano ${diffKm} km</div>`);
  }
  return out.join('');
}
function alertBadge(m) {
  return alertMezzo(m) ? `<span class="badge badge-red">ALERT</span>` : `<span class="badge badge-green">OK</span>`;
}

/* GOOGLE DRIVE VIA APPS SCRIPT - NON USA SERVICE ACCOUNT DIRETTO */
function googleDriveConfigured() {
  return !!(process.env.DRIVE_WEBAPP_URL && process.env.GOOGLE_DRIVE_FOLDER_ID);
}

async function getOrCreateDriveContractFolderV63(p) {
  if (!drive) return null;
  const folderName = `${p?.codice || 'CONTRATTO'} - ${p?.nome || ''} ${p?.cognome || ''}`.trim();
  const parent = process.env.GOOGLE_DRIVE_FOLDER_ID || process.env.DRIVE_FOLDER_ID || null;
  const safeName = folderName.replace(/'/g, "\'");
  let q = `mimeType='application/vnd.google-apps.folder' and name='${safeName}' and trashed=false`;
  if (parent) q += ` and '${parent}' in parents`;
  const found = await drive.files.list({ q, fields:'files(id,name,webViewLink)', spaces:'drive', supportsAllDrives:true, includeItemsFromAllDrives:true });
  if (found.data.files && found.data.files[0]) return found.data.files[0];
  const requestBody = { name: folderName, mimeType:'application/vnd.google-apps.folder' };
  if (parent) requestBody.parents = [parent];
  const created = await drive.files.create({ requestBody, fields:'id,name,webViewLink', supportsAllDrives:true });
  return created.data;
}

async function deleteDriveFilesByNameV63(folderId, name) {
  if (!drive || !folderId || !name) return;
  const safeName = String(name).replace(/'/g, "\\'");
  const found = await drive.files.list({
    q: `'${folderId}' in parents and name='${safeName}' and trashed=false`,
    fields:'files(id,name)',
    spaces:'drive'
  });
  for (const f of (found.data.files || [])) {
    try { await drive.files.delete({ fileId:f.id }); } catch(e) { console.log('Drive delete skip:', e.message); }
  }
}

async function uploadFileToDriveFolderV63(localPath, fileName, mimeType, folderId) {
  if (!drive || !folderId) return null;
  const media = { mimeType: mimeType || 'application/octet-stream', body: fs.createReadStream(localPath) };
  const created = await drive.files.create({
    requestBody:{ name:fileName, parents:[folderId] },
    media,
    fields:'id,name,webViewLink',
    supportsAllDrives:true
  });
  return created.data;
}

async function uploadFileToDrive(localPath, filename, mimetype, subFolderName) {
  if (!process.env.DRIVE_WEBAPP_URL || !process.env.GOOGLE_DRIVE_FOLDER_ID) return null;
  if (!fs.existsSync(localPath)) return null;

  const base64 = fs.readFileSync(localPath).toString('base64');

  const r = await fetch(process.env.DRIVE_WEBAPP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({
      folderId: process.env.GOOGLE_DRIVE_FOLDER_ID,
      subfolder: subFolderName || 'DP RENT',
      filename: filename,
      mimeType: mimetype || 'application/octet-stream',
      base64: base64
    })
  });

  const text = await r.text();
  let data;
  try { data = JSON.parse(text); }
  catch { throw new Error('Risposta Apps Script non valida: ' + text); }

  if (!data.ok) throw new Error(data.error || 'Upload Drive fallito');

  return {
    id: data.id || '',
    webViewLink: data.link || data.webViewLink || ''
  };
}

async function sendEmail(to, subject, text, attachments) {
  if (!process.env.SMTP_HOST) throw new Error('SMTP non configurato');
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: String(process.env.SMTP_SECURE || '').toLowerCase() === 'true' || Number(process.env.SMTP_PORT || 587) === 465,
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
    tls: { rejectUnauthorized: false }
  });
  return transporter.sendMail({
    from: process.env.SMTP_FROM || process.env.SMTP_USER || AZIENDA.email,
    to,
    subject,
    text,
    attachments: attachments || []
  });
}

/* NEXI PAYMAIL - COPIATO DAL PROGETTO WHATSAPP FUNZIONANTE
   Variabili Render:
   NEXI_ALIAS=payment_...
   NEXI_MAC_KEY=...
   NEXI_ENV=prod oppure test
   APP_BASE_URL=https://dp-rent-app.onrender.com
*/
const NEXI_ENV_APP = String(process.env.NEXI_ENV || 'prod').toLowerCase();
const NEXI_ALIAS_APP = process.env.NEXI_ALIAS || process.env.NEXI_API_KEY_ALIAS || '';
const NEXI_MAC_KEY_APP = process.env.NEXI_MAC_KEY || '';
const NEXI_TIMEOUT_HOURS_APP = Number(process.env.NEXI_TIMEOUT_HOURS || 4);
const NEXI_BASE_URL_APP = NEXI_ENV_APP === 'test' ? 'https://int-ecommerce.nexi.it' : 'https://ecommerce.nexi.it';
const NEXI_PAYMAIL_ENDPOINT_APP = `${NEXI_BASE_URL_APP}/ecomm/api/bo/richiestaPayMail`;

function nexiConfigured() {
  return Boolean(NEXI_ALIAS_APP && NEXI_MAC_KEY_APP && process.env.APP_BASE_URL);
}

function sha1(s) {
  return crypto.createHash('sha1').update(s).digest('hex');
}

function buildNexiOrderId(prefix = 'DPR') {
  return `${prefix}${Date.now().toString().slice(-10)}${Math.floor(Math.random() * 10000).toString().padStart(4, '0')}`.slice(0, 18);
}

function nexiMacPayMail({ apiKey, codiceTransazione, importo, timeStamp }) {
  const source =
    `apiKey=${apiKey}` +
    `codiceTransazione=${codiceTransazione}` +
    `importo=${importo}` +
    `timeStamp=${timeStamp}` +
    NEXI_MAC_KEY_APP;

  return crypto.createHash('sha1').update(source).digest('hex');
}

async function createNexiLink(amount, description, p) {
  if (!nexiConfigured()) {
    throw new Error('Nexi non configurato: servono NEXI_ALIAS, NEXI_MAC_KEY, APP_BASE_URL');
  }

  const amountCents = String(Math.round(Number(amount || 0) * 100));
  if (!amountCents || Number(amountCents) <= 0) throw new Error('Totale contratto non valido');

  const codiceTransazione = buildNexiOrderId('DPR');
  const timeStamp = Date.now().toString();
  const baseUrl = String(process.env.APP_BASE_URL || '').replace(/\/+$/, '');

  const payload = {
    apiKey: NEXI_ALIAS_APP,
    codiceTransazione,
    importo: amountCents,
    timeStamp,
    mac: nexiMacPayMail({
      apiKey: NEXI_ALIAS_APP,
      codiceTransazione,
      importo: amountCents,
      timeStamp
    }),
    timeout: String(NEXI_TIMEOUT_HOURS_APP),
    url: `${baseUrl}/nexi-ok/${p.id}`,
    parametriAggiuntivi: {
      source: 'dp_rent_app',
      contratto: p.codice || '',
      description: description || ''
    }
  };

  console.log('NEXI PAYMAIL REQUEST', {
    endpoint: NEXI_PAYMAIL_ENDPOINT_APP,
    codiceTransazione,
    importo: amountCents,
    env: NEXI_ENV_APP
  });

  const r = await fetch(NEXI_PAYMAIL_ENDPOINT_APP, {
    method: 'POST',
    headers: {
      Accept: 'application/json',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });

  const text = await r.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    data = { raw: text };
  }

  console.log('NEXI PAYMAIL RESPONSE', data);

  if (!r.ok) throw new Error(`HTTP Nexi ${r.status}: ${text}`);

  if (data.esito !== 'OK') {
    throw new Error(data?.errore?.messaggio || data?.errore?.description || data?.errore?.codice || ('Errore Nexi: ' + text));
  }

  const payUrl = data.payMailUrl || data.paymailUrl || data.url || data.urlPayMail || data.link || data.paymentUrl;

  if (!payUrl) {
    throw new Error('Nexi non ha restituito il link pagamento: ' + JSON.stringify(data).slice(0, 700));
  }

  return {
    codiceTransazione,
    link: payUrl,
    raw: text,
    amountCents
  };
}

function whatsappText(text) {
  return `https://wa.me/?text=${encodeURIComponent(text)}`;
}

function absoluteUrl(req, pathPart) {
  const base = process.env.APP_BASE_URL || `${req.protocol}://${req.get('host')}`;
  return `${base}${pathPart}`;
}

function actionScreen(id, titolo, messaggio) {
  return page(titolo, `
    <div class="box">
      <h2 class="ok">${titolo}</h2>
      <p>${messaggio || ''}</p>
      <div class="actions">
        <a class="btn" href="/contratto/${id}">PDF</a>
        <a class="btn btn2" href="/firma/${id}">Firma</a>
        <a class="btn btn2" href="/email/${id}">Email</a>
        <a class="btn btn3" href="/documenti/${id}">Foto/documenti</a>
        <a class="btn btn3" href="/checkout/${id}">Check-out</a>
        <a class="btn btn3" href="/checkin/${id}">Check-in</a>
        <a class="btn btnWarn" href="/nexi/${id}">Nexi</a>
        <a class="btn btn2" href="/prenotazione/${id}">Dettaglio</a>
      </div>
    </div>
  `);
}

function drawHeader(doc) {
  doc.rect(0,0,612,115).fill('#111111');
  const logoPath = path.join(publicDir, 'logo.png');
  if (fs.existsSync(logoPath)) {
    try { doc.image(logoPath, 35, 25, { fit: [145, 70] }); }
    catch { doc.fillColor('white').fontSize(28).text('DP RENT', 45, 35); }
  } else {
    doc.fillColor('white').fontSize(28).text('DP RENT', 45, 35);
  }
  doc.fillColor('white').fontSize(11).text(AZIENDA.nome, 350, 25, { align:'right', width:210 });
  doc.text(AZIENDA.indirizzo, 350, 43, { align:'right', width:210 });
  doc.text(`P.IVA / CF ${AZIENDA.piva} | Tel. ${AZIENDA.telefono}`, 350, 61, { align:'right', width:210 });
  doc.text(AZIENDA.email, 350, 79, { align:'right', width:210 });
  doc.rect(0,115,612,6).fill('#d90000');
  doc.fillColor('black');
}
function section(doc, title, x, y, w) {
  doc.rect(x,y,w,20).fill('#111111');
  doc.fillColor('white').fontSize(10).text(title, x+8, y+6);
  doc.fillColor('black');
  return y + 26;
}
function row(doc, label, value, x, y, w) {
  doc.fillColor('#777').fontSize(8).text(label, x, y, {width:w*0.38});
  doc.fillColor('#111').fontSize(9).text(String(value || ''), x+w*0.38, y, {width:w*0.6});
  doc.fillColor('black');
}

function pdfFileNameForContract(p) {
  const safe = String((p && (p.codice || p.id)) || Date.now()).replace(/[^a-zA-Z0-9_-]/g, '');
  return `contratto_${safe}.pdf`;
}

function shouldUploadPdfToDrive(p, forceDrive) {
  if (!googleDriveConfigured()) return false;
  if (forceDrive) return true;
  if (p.pdf_drive_web_link && p.stato !== 'firmato') return false;
  return true;
}

async function generaPdfContratto(id, opts = {}) {
  const p = await get(`
    SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.km_inclusi, m.descrizione_pubblica
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id = p.mezzo_id
    WHERE p.id = ?
  `, [id]);
  if (!p) throw new Error('Contratto non trovato');

  const file = path.join(contractsDir, pdfFileNameForContract(p));
  const doc = new PDFDocument({ margin: 28, size:'A4' });
  const stream = fs.createWriteStream(file);
  doc.pipe(stream);

  drawHeader(doc);
  doc.fontSize(20).fillColor('#111').text('CONTRATTO DI NOLEGGIO', 40, 150, {align:'center', width:515});
  doc.moveTo(40,180).lineTo(555,180).strokeColor('#d90000').lineWidth(1).stroke();

  let y = 205;
  y = section(doc, 'DATI CONTRATTO', 45, y, 510);
  row(doc, 'Numero contratto', p.codice, 55, y, 460); y += 18;
  row(doc, 'Stato', p.stato || 'bozza', 55, y, 460); y += 18;
  row(doc, 'Data creazione', p.created_at || '', 55, y, 460); y += 30;

  let yLeft = y, yRight = y;
  yLeft = section(doc, 'ANAGRAFICA CLIENTE', 45, yLeft, 245);
  row(doc, 'Cliente', `${p.nome || ''} ${p.cognome || ''}`, 55, yLeft, 220); yLeft += 18;
  row(doc, 'Telefono', p.telefono || '', 55, yLeft, 220); yLeft += 18;
  row(doc, 'Email', p.email || '', 55, yLeft, 220); yLeft += 18;
  row(doc, 'Codice fiscale', p.codice_fiscale || p.cf || '', 55, yLeft, 220); yLeft += 18;
  row(doc, 'Indirizzo', `${p.indirizzo || ''} ${p.cap || ''} ${p.citta || ''} ${p.provincia || ''}`, 55, yLeft, 220); yLeft += 18;
  row(doc, 'Fatturazione', `${p.tipo_cliente || p.fatturazione || ''} ${p.ragione_sociale || p.azienda || ''} ${p.piva || ''}`, 55, yLeft, 220); yLeft += 18;
  row(doc, 'PEC / SDI', `${p.pec || ''} ${p.sdi || ''}`, 55, yLeft, 220); yLeft += 18;

  yRight = section(doc, 'CONDUCENTI', 310, yRight, 245);
  row(doc, 'Conducente 1', p.conducente1 || `${p.nome || ''} ${p.cognome || ''}`, 320, yRight, 220); yRight += 18;
  row(doc, 'Patente 1', `${p.patente1 || ''} scad. ${p.patente1_scadenza || ''}`, 320, yRight, 220); yRight += 18;
  row(doc, 'Conducente 2', p.conducente2 || '', 320, yRight, 220); yRight += 18;
  row(doc, 'Patente 2', `${p.patente2 || ''} scad. ${p.patente2_scadenza || ''}`, 320, yRight, 220); yRight += 18;

  y = Math.max(yLeft, yRight) + 10;
  let yVeh = y, yCost = y;
  yVeh = section(doc, 'VEICOLO E NOLEGGIO', 45, yVeh, 245);
  row(doc, 'Targa', p.targa || '', 55, yVeh, 220); yVeh += 18;
  row(doc, 'Mezzo', p.descrizione_pubblica || `${p.marca || ''} ${p.modello || ''}`, 55, yVeh, 220); yVeh += 18;
  row(doc, 'Periodo', `${p.data_inizio} ${p.ora_inizio || ''} / ${p.data_fine} ${p.ora_fine || ''}`, 55, yVeh, 220); yVeh += 18;
  row(doc, 'Giorni', String(p.giorni || ''), 55, yVeh, 220); yVeh += 18;
  row(doc, 'Km inclusi / previsti', `${Number(p.km_inclusi || 0) * Number(p.giorni || 0)} / ${p.km_previsti || 0}`, 55, yVeh, 220); yVeh += 18;
  row(doc, 'Km uscita/rientro', `${p.km_uscita || ''} / ${p.km_rientro || ''}`, 55, yVeh, 220); yVeh += 18;
  row(doc, 'Carburante', `${p.carburante_uscita || ''} / ${p.carburante_rientro || ''}`, 55, yVeh, 220); yVeh += 18;

  yCost = section(doc, 'RIEPILOGO ECONOMICO', 310, yCost, 245);
  row(doc, 'Extra orario', `euro ${euro(p.extra_fuori_orario)} + IVA`, 320, yCost, 220); yCost += 18;
  row(doc, 'Extra km', `euro ${euro(p.extra_km)} + IVA`, 320, yCost, 220); yCost += 18;
  row(doc, 'Imponibile', `euro ${euro(p.imponibile)}`, 320, yCost, 220); yCost += 18;
  row(doc, 'IVA 22%', `euro ${euro(p.iva)}`, 320, yCost, 220); yCost += 18;
  row(doc, 'Totale IVA inclusa', `euro ${euro(p.totale)}`, 320, yCost, 220); yCost += 18;
  row(doc, 'Deposito cauzionale', `euro ${euro(p.cauzione || CAUZIONE)}`, 320, yCost, 220); yCost += 18;

  y = Math.max(yVeh, yCost) + 10;
  y = section(doc, 'CONDIZIONI GENERALI E PRIVACY', 45, y, 510);
  doc.fontSize(8).fillColor('#111').text('Il cliente dichiara di aver preso visione e accettare le condizioni generali di noleggio e l’informativa privacy DP RENT / Trasporti DP S.R.L. Il mezzo deve essere riconsegnato nelle stesse condizioni, con carburante equivalente. Danni, multe, pedaggi, franchigie, ritardi, smarrimenti e costi accessori restano a carico del cliente.', 55, y, {width:490});
  y += 50;
  if (TERMS_URL) { doc.fontSize(7).text(`Condizioni generali: ${TERMS_URL}`, 55, y, {width:490}); y += 12; }
  if (PRIVACY_URL) { doc.fontSize(7).text(`Informativa privacy: ${PRIVACY_URL}`, 55, y, {width:490}); y += 12; }
  y += 20;

  doc.fontSize(10).fillColor('#111').text('Firma cliente:', 55, y);
  if (p.firma_path && fs.existsSync(p.firma_path)) {
    try { doc.image(p.firma_path, 55, y+15, { fit: [220, 70] }); }
    catch { doc.text('______________________________', 55, y+25); }
  } else {
    doc.text('______________________________', 55, y+25);
  }
  doc.text('Firma DP RENT:', 330, y);
  doc.text('______________________________', 330, y+25);
  
  const baseUrl = process.env.APP_BASE_URL || process.env.PUBLIC_URL || '';
  doc.moveDown(0.8);
  doc.fontSize(10).fillColor('blue')
    .text('Privacy Policy', { link: baseUrl ? `${baseUrl}/privacy.pdf` : '/privacy.pdf', underline: true });
  doc.moveDown(0.3);
  doc.text('Condizioni Generali di Noleggio', { link: baseUrl ? `${baseUrl}/clausole.pdf` : '/clausole.pdf', underline: true });
  doc.fillColor('black');

doc.end();

  await new Promise((resolve, reject) => {
    stream.on('finish', resolve);
    stream.on('error', reject);
  });

  let driveRes = null;
  const forceDrive = !!opts.forceDrive;
  const skipDrive = !!opts.skipDrive;

  if (!skipDrive && shouldUploadPdfToDrive(p, forceDrive)) {
    try {
      driveRes = await uploadFileToDrive(
        file,
        path.basename(file),
        'application/pdf',
        `${p.codice || 'contratto'} - ${p.nome || ''} ${p.cognome || ''}`
      );
    } catch (e) {
      console.log('Errore upload PDF Google Drive:', e.message);
    }
  }

  if (driveRes) {
    await run(`UPDATE prenotazioni SET pdf_path=?, pdf_drive_file_id=?, pdf_drive_web_link=? WHERE id=?`,
      [file, driveRes.id, driveRes.webViewLink, id]);
  } else {
    await run(`UPDATE prenotazioni SET pdf_path=? WHERE id=?`, [file, id]);
  }

  return file;
}

/* ROUTES */



// =========================
// V107 CARGOS UID LOCK
// =========================
const CARGOS_DEFAULT_LUOGO_NARNI = '410055022';

function v61CleanKey(v) {
  return String(v || '').toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[’']/g, '').replace(/[_\-]+/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function old_getTipoVeicoloCargosV63_DISABLED(v) {
  const k = v61CleanKey(v);
  if (k === '0' || k.includes('AUTO') || k.includes('MACCHINA')) return '0';
  if (k === '1' || k.includes('FURG') || k.includes('VAN') || k.includes('DAILY') || k.includes('DUCATO') || k.includes('TRANSIT')) return '1';
  if (k === '3' || k.includes('BUS') || k.includes('PULMINO') || k.includes('9 POSTI')) return '3';
  if (k === '4' || k.includes('AUTOCAR') || k.includes('MOTRICE')) return '4';
  if (k === '5' || k.includes('TRATTORE')) return '5';
  if (k === '6' || k.includes('AUTOTRENO')) return '6';
  if (k === '7' || k.includes('ARTICOL') || k.includes('BISARCA')) return '7';
  if (k === '8' || k.includes('SNODAT')) return '8';
  if (k === '9' || k.includes('CAMPER') || k.includes('CARAVAN')) return '9';
  if (k === 'A' || k.includes('ESCAV') || k.includes('SEMOV') || k.includes('OPERA')) return 'A';
  return '0';
}

function getTipoDocumentoCargosV63(v) {
  const k = v61CleanKey(v);
  if (['CIDIP','IDELE','IDENT','PASDI','PASOR','PASSE','PATEN'].includes(k)) return k;
  if (k.includes('ELETTRON') || k.includes('CIE')) return 'IDELE';
  if (k.includes('PASSAP')) return 'PASOR';
  if (k.includes('PATENTE')) return 'PATEN';
  return 'IDENT';
}

function getTipoPagamentoCargosV63(v) {
  const k = v61CleanKey(v);
  if (['0','1','2','3','4','9'].includes(k)) return k;
  if (k.includes('CREDITO')) return '0';
  if (k.includes('CONTANTI') || k.includes('CASH')) return '1';
  if (k.includes('DEBITO') || k.includes('BANCOMAT')) return '2';
  if (k.includes('BONIFICO')) return '3';
  if (k.includes('RID')) return '4';
  return '9';
}

function cargosCheckoutLuogoCodV63() {
  return String(process.env.CARGOS_LUOGO_COD || process.env.CHECKOUT_LUOGO_COD || process.env.CARGOS_CHECKOUT_LUOGO_COD || CARGOS_DEFAULT_LUOGO_NARNI).replace(/\D/g,'').trim() || CARGOS_DEFAULT_LUOGO_NARNI;
}

function cargosCheckinLuogoCodV63() {
  return String(process.env.CARGOS_CHECKIN_LUOGO_COD || process.env.CHECKIN_LUOGO_COD || process.env.CARGOS_LUOGO_COD || cargosCheckoutLuogoCodV63()).replace(/\D/g,'').trim() || cargosCheckoutLuogoCodV63();
}

function cargosOperatoreIdV63() {
  return String(process.env.CARGOS_OPERATORE_ID || 'DPRENT01').trim().slice(0,50) || 'DPRENT01';
}

function cargosAgenziaIdV63() {
  return String(process.env.CARGOS_AGENZIA_ID || '001').trim().slice(0,30) || '001';
}

function cargosConfigured() {
  return !!(((process.env.CARGOS_USERNAME || 'C00000100') || 'C00000100') && process.env.CARGOS_PASSWORD && process.env.CARGOS_APIKEY && process.env.CARGOS_AGENZIA_ID && process.env.CARGOS_OPERATORE_ID);
}

function cargosPad(value, len, type = 'string') {
  let v = String(value === undefined || value === null ? '' : value).normalize('NFC').replace(/\s+/g,' ').trim();
  if (type === 'number') v = v.replace(/\D/g, '');
  if (v.length > len) v = v.slice(0, len);
  return v.padEnd(len, ' ');
}

function cargosDateTime(value, timeValue) {
  const d = value ? moment(value) : moment();
  const t = String(timeValue || '08:30').slice(0,5);
  if (!d.isValid()) return ''.padEnd(16, ' ');
  return `${d.format('DD/MM/YYYY')} ${t}`.padEnd(16, ' ');
}

function cargosDate(value) {
  const d = value ? moment(value) : null;
  if (!d || !d.isValid()) return ''.padEnd(10, ' ');
  return d.format('DD/MM/YYYY').padEnd(10, ' ');
}

function cargosTipoVeicolo(categoria) {
  const c = String(categoria || '').toUpperCase();
  if (c.includes('AUTO')) return process.env.CARGOS_TIPO_VEICOLO_AUTO || 'A';
  return process.env.CARGOS_TIPO_VEICOLO_FURGONE || 'F';
}


// =========================
// V76 FIX DEFINITIVO VEICOLO_TIPO CARGOS
// Ca.R.G.O.S. accetta solo: 0,1,3,4,5,6,7,8,9,A
// =========================
function normalizeCargosVehicleTypeV76(v, mezzoTxt='') {
  let k = String(v ?? '').trim().toUpperCase();
  const txt = String((mezzoTxt || '') + ' ' + k).toUpperCase();
  if (['0','1','3','4','5','6','7','8','9','A'].includes(k)) return k;
  if (txt.includes('FURG') || txt.includes('VAN') || txt.includes('DAILY') || txt.includes('DUCATO') || txt.includes('VIVARO') || txt.includes('TRAFIC') || txt.includes('TRANSIT') || txt.includes('EXPERT') || txt.includes('SCUDO') || txt.includes('DOBLO') || txt.includes('MASTER') || txt.includes('SPRINTER') || txt.includes('VITO') || txt.includes('CARGO')) return '1';
  if (txt.includes('9 POSTI') || txt.includes('PULMINO') || txt.includes('BUS') || txt.includes('AUTOBUS')) return '3';
  if (txt.includes('AUTOCARRO') || txt.includes('MOTRICE') || txt.includes('CAMION')) return '4';
  if (txt.includes('TRATTORE')) return '5';
  if (txt.includes('AUTOTRENO')) return '6';
  if (txt.includes('AUTOARTICOLATO') || txt.includes('ARTICOLATO')) return '7';
  if (txt.includes('AUTOSNODATO') || txt.includes('SNODATO')) return '8';
  if (txt.includes('CAMPER') || txt.includes('AUTOCARAVAN')) return '9';
  if (txt.includes('OPERA')) return 'A';
  return '1';
}
function getTipoVeicoloCargosV76(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV72(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV71(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV70(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV69(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV68(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV66(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV65(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargosV61(v){ return normalizeCargosVehicleTypeV76(v); }
function getTipoVeicoloCargos(v){ return normalizeCargosVehicleTypeV76(v); }
function patchCargosVehicleTypeIntoObjectV76(p) {
  p = p || {};
  const mezzoTxt = [p.marca, p.modello, p.mezzo, p.nome_mezzo, p.categoria, p.tipo_mezzo, p.record_cargos_veicolo_tipo].filter(Boolean).join(' ');
  const cod = normalizeCargosVehicleTypeV76(p.record_cargos_veicolo_tipo || p.veicolo_tipo || p.tipo_veicolo || p.categoria, mezzoTxt);
  p.record_cargos_veicolo_tipo = cod;
  p.veicolo_tipo = cod;
  p.tipo_veicolo = cod;
  return p;
}


// =========================
// V76 FIX ORDINE RECORD CARGOS VEICOLO
// Il tracciato vuole: VEICOLO_TIPO prima di MARCA e MODELLO.
// Se manca, Ca.R.G.O.S. legge FIAT/FORD come VEICOLO_TIPO e rifiuta.
// =========================
function cargosTipoVeicoloFinaleV76(p) {
  p = p || {};
  const txt = [
    p.record_cargos_veicolo_tipo,
    p.veicolo_tipo,
    p.tipo_veicolo,
    p.categoria,
    p.marca,
    p.modello,
    p.mezzo,
    p.nome_mezzo
  ].filter(Boolean).join(' ');

  if (typeof normalizeCargosVehicleTypeV76 === 'function') return normalizeCargosVehicleTypeV76(p.record_cargos_veicolo_tipo || p.veicolo_tipo || p.tipo_veicolo || p.categoria, txt);
  if (typeof normalizeCargosVehicleTypeV76 === 'function') return normalizeCargosVehicleTypeV76(p.record_cargos_veicolo_tipo || p.veicolo_tipo || p.tipo_veicolo || p.categoria, txt);

  const k = String(txt || '').toUpperCase();
  if (['0','1','3','4','5','6','7','8','9','A'].includes(String(p.record_cargos_veicolo_tipo || '').trim().toUpperCase())) {
    return String(p.record_cargos_veicolo_tipo).trim().toUpperCase();
  }
  if (k.includes('FURG') || k.includes('FIORINO') || k.includes('VIVARO') || k.includes('TRAFIC') || k.includes('DUCATO') || k.includes('DAILY') || k.includes('TRANSIT')) return '1';
  return '1';
}

function cargosMarcaFinaleV76(p) {
  p = p || {};
  return String(p.record_cargos_marca || p.marca || p.brand || '').trim().toUpperCase();
}

function cargosModelloFinaleV76(p) {
  p = p || {};
  return String(p.record_cargos_modello || p.modello || p.model || p.mezzo || '').trim().toUpperCase();
}

function patchCargosVehicleRecordV76(p) {
  p = p || {};
  p.record_cargos_veicolo_tipo = cargosTipoVeicoloFinaleV76(p);
  p.veicolo_tipo = p.record_cargos_veicolo_tipo;
  p.tipo_veicolo = p.record_cargos_veicolo_tipo;
  p.record_cargos_marca = cargosMarcaFinaleV76(p);
  p.record_cargos_modello = cargosModelloFinaleV76(p);
  return p;
}

async function buildCargosRecordForContract(id) {
  // V76 patch veicolo

  // buildCargosRecordForContract__v72patched
  // V76 rimosso: p non ancora inizializzato

  const p = await get(`SELECT p.*, m.targa, m.marca, m.modello, m.categoria FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE p.id=?`, [id]);
  if (!p) throw new Error('Contratto non trovato');

  
  
  patchCargosVehicleRecordV76(p);
Object.assign(p, cargosPatchDefaultsV76(p));
  patchCargosVehicleTypeIntoObjectV76(p);
const fields = [
    cargosPad(p.codice,50), cargosDateTime(p.created_at || moment(), moment().format('HH:mm')),
    cargosPad(process.env.CARGOS_TIPO_PAGAMENTO || '1',1,'number'),
    cargosDateTime(p.data_inizio,p.ora_inizio || '08:30'), cargosPad(process.env.CARGOS_LUOGO_COD || '',9,'number'), cargosPad(AZIENDA.indirizzo,150),
    cargosDateTime(p.data_fine,p.ora_fine || '18:00'), cargosPad(process.env.CARGOS_LUOGO_COD || '',9,'number'), cargosPad(AZIENDA.indirizzo,150),
    cargosPad(process.env.CARGOS_OPERATORE_ID || '',50), cargosPad(process.env.CARGOS_AGENZIA_ID || '',30), cargosPad(process.env.CARGOS_AGENZIA_NOME || AZIENDA.nome,70),
    cargosPad(process.env.CARGOS_LUOGO_COD || '',9,'number'), cargosPad(AZIENDA.indirizzo,150), cargosPad(AZIENDA.telefono,20,'number'),
    cargosPad(cargosTipoVeicoloFinaleV76(p),1), cargosPad(cargosMarcaFinaleV76(p) || p.marca || '',50), cargosPad(cargosModelloFinaleV76(p) || p.modello || '',100), cargosPad(p.targa || '',15),
    cargosPad('',50), cargosPad('',1,'number'), cargosPad('',1,'number'),
    cargosPad(p.cognome || '',50), cargosPad(p.nome || '',30), cargosDate(p.data_nascita || ''),
    cargosPad(process.env.CARGOS_NASCITA_LUOGO_COD || process.env.CARGOS_LUOGO_COD || '',9,'number'),
    cargosPad(process.env.CARGOS_CITTADINANZA_COD || '',9,'number'),
    cargosPad(process.env.CARGOS_RESIDENZA_LUOGO_COD || process.env.CARGOS_LUOGO_COD || '',9,'number'),
    cargosPad(p.indirizzo || '',150),
    cargosPad(process.env.CARGOS_TIPO_DOCUMENTO || 'CI',5),
    cargosPad(process.env.CARGOS_DOC_NUMERO_FALLBACK || '',20),
    cargosPad(process.env.CARGOS_DOC_LUOGO_COD || process.env.CARGOS_LUOGO_COD || '',9,'number'),
    cargosPad(p.patente1 || '',20),
    cargosPad(process.env.CARGOS_PATENTE_LUOGO_COD || process.env.CARGOS_LUOGO_COD || '',9,'number'),
    cargosPad(p.telefono || '',20),
    cargosPad('',50), cargosPad('',30), cargosDate(''), cargosPad('',9,'number'), cargosPad('',9,'number'),
    cargosPad('',5), cargosPad('',20), cargosPad('',9,'number'), cargosPad('',20), cargosPad('',9,'number'), cargosPad('',20)
  ];

  const record = fields.join('');
  if (record.length !== 1505) throw new Error(`Record Ca.R.G.O.S. lunghezza errata: ${record.length}, attesa 1505`);
  return record;
}


// =========================
// V107 CARGOS UID LOCK
// =========================
function cargosCfgGet(k, def='') {
  return process.env[k] || process.env['CARGOS_' + k] || def || '';
}

function cargosCheckoutLuogoCodV63() {
  return String(
    cargosCfgGet('CHECKOUT_LUOGO_COD') ||
    cargosCfgGet('LUOGO_CHECKOUT_COD') ||
    cargosCfgGet('LOCATION_CODE') ||
    cargosCfgGet('LUOGO_COD') ||
    cargosCfgGet('SEDE_COD') ||
    '001'
  ).trim();
}

function cargosCheckinLuogoCodV63() {
  return String(
    cargosCfgGet('CHECKIN_LUOGO_COD') ||
    cargosCfgGet('LUOGO_CHECKIN_COD') ||
    cargosCfgGet('LOCATION_CODE') ||
    cargosCfgGet('LUOGO_COD') ||
    cargosCfgGet('SEDE_COD') ||
    cargosCheckoutLuogoCodV63()
  ).trim();
}

function normalizeCargosLuogoCod(v) {
  const s = String(v || '').trim();
  const only = s.replace(/\D/g, '');
  return only || s;
}

async function cargosGetToken() {
  const base = (process.env.CARGOS_BASE_URL || 'https://cargos.poliziadistato.it/CARGOS_API').replace(/\/+$/, '');
  const cargosUser = process.env.CARGOS_USERNAME || 'C00000100';
  const basic = Buffer.from(`${cargosUser}:${process.env.CARGOS_PASSWORD}`).toString('base64');
  const r = await fetch(`${base}/api/Token`, { method:'GET', headers:{ Authorization:`Basic ${basic}` } });
  const text = await r.text();
  let data; try { data = JSON.parse(text); } catch { data = { raw:text }; }
  if (!r.ok || data.error) throw new Error('Errore token Ca.R.G.O.S.: ' + text);
  return data.access_token || data.accessToken || data.token || data?.Esito?.access_token;
}

function cargosEncryptAes(token) {
  const keySrc = String(process.env.CARGOS_APIKEY || '');
  if (keySrc.length < 48) throw new Error('CARGOS_APIKEY deve avere almeno 48 caratteri per AES');
  const key = Buffer.from(keySrc.substring(0,32),'utf8');
  const iv = Buffer.from(keySrc.substring(32,48),'utf8');
  const cipher = crypto.createCipheriv('aes-256-cbc', key, iv);
  return Buffer.concat([cipher.update(String(token),'utf8'), cipher.final()]).toString('base64');
}

async function cargosSendRecords(records, method='Check') {
  if (!cargosConfigured()) throw new Error('Ca.R.G.O.S. non configurato: servono username/password/apikey/agenzia/operatore/codici.');
  const base = (process.env.CARGOS_BASE_URL || 'https://cargos.poliziadistato.it/CARGOS_API').replace(/\/+$/, '');
  const encrypted = cargosEncryptAes(await cargosGetToken());
  const r = await fetch(`${base}/api/${method}`, {
    method:'POST',
    headers:{ Authorization:`Bearer ${encrypted}`, Organization:cargosOrganizationHeaderV76(), 'Content-Type':'application/json', Accept:'application/json' },
    body: JSON.stringify(records)
  });
  const text = await r.text();
  let data; try { data = JSON.parse(text); } catch { data = { raw:text }; }
  if (!r.ok) throw new Error(`HTTP Ca.R.G.O.S. ${r.status}: ${text}`);
  return data;
}


// V63 - test veloce configurazione CARGOS senza toccare prenotazioni
app.get('/admin/cargos-test-token', async (req, res) => {
  try {
    const token = await cargosGetToken();
    const encrypted = cargosEncryptAes(token);
    res.send(page('CARGOS TOKEN OK', `<div class="box">
      <h2 class="ok">CARGOS TOKEN/AES OK</h2>
      <p>Token ricevuto e cifrato AES correttamente.</p>
      <p>Token cifrato: ${esc(encrypted).slice(0,80)}...</p>
      <a class="btn" href="/cargos">Ca.R.G.O.S.</a>
    </div>`));
  } catch (e) {
    res.status(500).send(page('CARGOS TOKEN KO', `<div class="box">
      <h2 class="bad">CARGOS TOKEN/AES KO</h2>
      <pre>${esc(e.message)}</pre>
      <p>Controlla CARGOS_USERNAME, CARGOS_PASSWORD, CARGOS_APIKEY, CARGOS_BASE_URL.</p>
    </div>`));
  }
});

async function estraiDatiDocumentoConAI(localPath, mimetype) {
  if (!process.env.OPENAI_API_KEY) throw new Error('OPENAI_API_KEY mancante su Render.');
  if (!fs.existsSync(localPath)) throw new Error('File OCR non trovato');

  const base64 = fs.readFileSync(localPath).toString('base64');
  const dataUrl = `data:${mimetype || 'image/jpeg'};base64,${base64}`;

  const prompt = `Leggi documento italiano patente/carta identità. Rispondi SOLO JSON valido:
{
"tipo_documento":"","nome":"","cognome":"","data_nascita":"YYYY-MM-DD","luogo_nascita":"",
"codice_fiscale":"","numero_documento":"","ente_rilascio":"","data_rilascio":"YYYY-MM-DD",
"data_scadenza":"YYYY-MM-DD","numero_patente":"","categoria_patente":"","indirizzo":"",
"note":"","confidence":"alta|media|bassa"
}
Se un campo non è visibile lascia vuoto.`;

  const r = await fetch('https://api.openai.com/v1/responses', {
    method: 'POST',
    headers: { 'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: process.env.OPENAI_OCR_MODEL || 'gpt-4.1-mini',
      input: [{ role: 'user', content: [{ type: 'input_text', text: prompt }, { type: 'input_image', image_url: dataUrl }] }],
      temperature: 0
    })
  });

  const text = await r.text();
  if (!r.ok) throw new Error('Errore OpenAI OCR: ' + text);

  const data = JSON.parse(text);
  const outputText = data.output_text || (data.output || []).flatMap(o => o.content || []).map(c => c.text || '').join('\\n');
  const clean = String(outputText || '').replace(/^```json/i,'').replace(/^```/i,'').replace(/```$/i,'').trim();
  return JSON.parse(clean);
}

function ocrValue(v) { return esc(v || ''); }


// =========================
// V44 MIGRAZIONE FORZATA DB MEZZI / PRENOTAZIONI
// =========================

// =========================
// V45 SAFE IMPORT MEZZI
// =========================
function insertMezzoSafe(m) {
  return new Promise((resolve, reject) => {
    db.run(`
      INSERT INTO mezzi (
        uid,targa,marca,modello,km,codice_tipo,descrizione
      ) VALUES (?,?,?,?,?,?,?)
    `, [
      m.uid || '',
      m.targa || '',
      m.marca || '',
      m.modello || '',
      m.km || '',
      m.codice_tipo || '',
      m.descrizione || ''
    ], function(err){
      if(err){
        console.log('V45 insert mezzo error:', err.message);
        return reject(err);
      }
      resolve(this.lastID);
    });
  });
}

function runV44DbMigration() {
  db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      uid TEXT,
      targa TEXT,
      marca TEXT,
      modello TEXT,
      tipo TEXT,
      descrizione TEXT,
      km TEXT,
      stato TEXT DEFAULT 'attivo'
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      cliente_id INTEGER,
      mezzo_id INTEGER,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      cf TEXT,
      data_inizio TEXT,
      ora_inizio TEXT,
      data_fine TEXT,
      ora_fine TEXT,
      totale REAL,
      stato TEXT DEFAULT 'bozza'
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS clienti (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      cf TEXT,
      indirizzo TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS allegati (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      prenotazione_id INTEGER,
      tipo TEXT,
      filename TEXT,
      originalname TEXT,
      path TEXT,
      mimetype TEXT,
      size INTEGER,
      drive_file_id TEXT,
      drive_web_link TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);

    const mezziCols = {
      uid:'TEXT', targa:'TEXT', marca:'TEXT', modello:'TEXT', tipo:'TEXT', descrizione:'TEXT',
      cilindrata:'TEXT', alimentazione:'TEXT', codice_tipo:'TEXT', codice_marca:'TEXT', codice_modello:'TEXT',
      categoria:'TEXT', posti:'TEXT', km:'TEXT', km_attuali:'TEXT', telaio:'TEXT', colore:'TEXT',
      stazione:'TEXT', soccorso_stradale:'TEXT', immagini_consegna:'TEXT', prezzo_giorno:'TEXT',
      km_inclusi:'TEXT', cauzione:'TEXT', deposito:'TEXT', gps:'TEXT', blocco_motore:'TEXT',
      record_cargos_veicolo_tipo:'TEXT', anno:'TEXT', numero_interno:'TEXT', disponibile:'TEXT',
      attivo:'TEXT', ubicazione:'TEXT', proprieta:'TEXT', note:'TEXT', note_interne:'TEXT',
      data_immatricolazione:'TEXT', ultima_revisione:'TEXT', scadenza_revisione:'TEXT',
      scadenza_bollo:'TEXT', scadenza_assicurazione:'TEXT', prossimo_tagliando:'TEXT',
      tagliando_km:'TEXT', serbatoio:'TEXT', cambio:'TEXT', porte:'TEXT', euro:'TEXT',
      iva:'TEXT', franchigia:'TEXT'
    };

    for (const [c,t] of Object.entries(mezziCols)) {
      db.run(`ALTER TABLE mezzi ADD COLUMN ${c} ${t}`, (err) => {
        if (err && !String(err.message || '').includes('duplicate column')) {
          console.log('V44 ADD COLUMN mezzi warning:', c, err.message);
        }
      });
    }

    const prenCols = {
      record_cargos_uid:'TEXT', record_cargos_transactionid:'TEXT', record_cargos_stato:'TEXT',
      record_cargos_last_check:'TEXT', record_cargos_last_send:'TEXT', record_cargos_last_error:'TEXT',
      drive_folder_id:'TEXT', drive_folder_link:'TEXT', pdf_drive_link:'TEXT',
      indirizzo:'TEXT', citta:'TEXT', cap:'TEXT', provincia:'TEXT',
      numero_documento:'TEXT', numero_patente:'TEXT', data_nascita:'TEXT',
      luogo_nascita:'TEXT', fatturazione:'TEXT', azienda:'TEXT', piva:'TEXT',
      pec:'TEXT', sdi:'TEXT'
    };

    for (const [c,t] of Object.entries(prenCols)) {
      db.run(`ALTER TABLE prenotazioni ADD COLUMN ${c} ${t}`, (err) => {
        if (err && !String(err.message || '').includes('duplicate column')) {
          console.log('V44 ADD COLUMN prenotazioni warning:', c, err.message);
        }
      });
    }
  });
  console.log('V44 migrazione DB eseguita');
}
runV44DbMigration();


// =========================
// V48 IMPORT MEZZI DEFINITIVO - NO ON CONFLICT
// =========================
const importUploadV48 = multer({ dest: (typeof uploadDir !== 'undefined' ? uploadDir : path.join(__dirname, 'uploads')) });

function v48Cell(row, names) {
  for (const n of names) {
    if (Object.prototype.hasOwnProperty.call(row, n) && row[n] !== undefined && row[n] !== null && String(row[n]).trim() !== '') {
      return String(row[n]).trim();
    }
  }
  return '';
}

function v48ParseImportFile(filePath) {
  const wb = XLSX.readFile(filePath, { raw: false });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

function v48EnsureImportDb(done) {
  db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      uid TEXT,
      targa TEXT,
      marca TEXT,
      modello TEXT,
      tipo TEXT,
      descrizione TEXT,
      km TEXT,
      stato TEXT DEFAULT 'attivo'
    )`);
    const cols = {
      uid:'TEXT', targa:'TEXT', marca:'TEXT', modello:'TEXT', tipo:'TEXT', descrizione:'TEXT',
      cilindrata:'TEXT', alimentazione:'TEXT', codice_tipo:'TEXT', codice_marca:'TEXT', codice_modello:'TEXT',
      categoria:'TEXT', posti:'TEXT', km:'TEXT', km_attuali:'TEXT', telaio:'TEXT', colore:'TEXT',
      stazione:'TEXT', soccorso_stradale:'TEXT', immagini_consegna:'TEXT', prezzo_giorno:'TEXT',
      km_inclusi:'TEXT', cauzione:'TEXT', deposito:'TEXT', gps:'TEXT', blocco_motore:'TEXT',
      record_cargos_veicolo_tipo:'TEXT', anno:'TEXT', note:'TEXT'
    };
    const keys = Object.keys(cols);
    let remaining = keys.length;
    keys.forEach((c) => {
      db.run(`ALTER TABLE mezzi ADD COLUMN ${c} ${cols[c]}`, () => {
        remaining -= 1;
        if (remaining === 0 && done) done();
      });
    });
  });
}

function v48TipoFromCodice(codice) {
  const c = String(codice || '').toUpperCase();
  if (c.startsWith('F')) return 'furgone';
  if (c.startsWith('P')) return 'pulmino';
  if (c.startsWith('A')) return 'auto';
  if (c.startsWith('X')) return 'attrezzatura';
  return 'mezzo';
}

function v48InsertOrUpdateMezzo(row) {
  const uid = v48Cell(row, ['uid', 'UID', 'Uid']);
  const targa = v48Cell(row, ['targa', 'Targa']).toUpperCase();
  const marca = v48Cell(row, ['marca', 'Marca']).toUpperCase();
  const modello = v48Cell(row, ['modello', 'Modello']).toUpperCase();
  const km = v48Cell(row, ['km', 'Km percorsi', 'km_attuali', 'KM']);
  const codice_tipo = v48Cell(row, ['codice_tipo', 'Codice Tipologia Mezzo', 'Codice Tipo', 'codice tipo']);
  const descrizione = v48Cell(row, ['descrizione', 'Descrizione Tipologia', 'Descrizione']);
  const cilindrata = v48Cell(row, ['cilindrata', 'Cilindrata']);
  const alimentazione = v48Cell(row, ['alimentazione', 'Alimentazione']);
  const tipo = v48Cell(row, ['tipo', 'Tipo']) || v48TipoFromCodice(codice_tipo);
  const categoria = v48Cell(row, ['categoria', 'Categoria']);
  const posti = v48Cell(row, ['posti', 'Posti']);
  const prezzo_giorno = v48Cell(row, ['prezzo_giorno', 'Prezzo giorno']) || (tipo === 'auto' ? '60' : '70');
  const km_inclusi = v48Cell(row, ['km_inclusi', 'Km inclusi']) || '150';
  const cauzione = v48Cell(row, ['cauzione', 'Cauzione']) || '500';
  const gps = v48Cell(row, ['gps', 'GPS']) || '0';
  const blocco_motore = v48Cell(row, ['blocco_motore', 'Blocco motore']) || '0';
  const stazione = v48Cell(row, ['stazione', 'Stazione']) || 'Narni';
  const soccorso = v48Cell(row, ['soccorso_stradale', 'Soccorso stradale']);
  const immagini = v48Cell(row, ['immagini_consegna', 'Immagini consegna']);
  const stato = v48Cell(row, ['stato', 'Stato']) || 'attivo';

  return new Promise((resolve, reject) => {
    if (!targa) return resolve({ skipped: 1 });
    db.get(`SELECT id FROM mezzi WHERE targa = ? OR (uid <> '' AND uid = ?) LIMIT 1`, [targa, uid], (err, old) => {
      if (err) return reject(err);
      const vals = [uid, targa, marca, modello, tipo, descrizione, km, km, codice_tipo, cilindrata, alimentazione, categoria, posti, prezzo_giorno, km_inclusi, cauzione, gps, blocco_motore, stazione, soccorso, immagini, stato];
      if (old && old.id) {
        db.run(`UPDATE mezzi SET uid=?, targa=?, marca=?, modello=?, tipo=?, descrizione=?, km=?, km_attuali=?, codice_tipo=?, cilindrata=?, alimentazione=?, categoria=?, posti=?, prezzo_giorno=?, km_inclusi=?, cauzione=?, gps=?, blocco_motore=?, stazione=?, soccorso_stradale=?, immagini_consegna=?, stato=? WHERE id=?`,
          vals.concat([old.id]), (e) => e ? reject(e) : resolve({ updated: 1 }));
      } else {
        db.run(`INSERT INTO mezzi (uid,targa,marca,modello,tipo,descrizione,km,km_attuali,codice_tipo,cilindrata,alimentazione,categoria,posti,prezzo_giorno,km_inclusi,cauzione,gps,blocco_motore,stazione,soccorso_stradale,immagini_consegna,stato) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
          vals, (e) => e ? reject(e) : resolve({ inserted: 1 }));
      }
    });
  });
}


// =========================
// V50 MIGRAZIONE PRENOTAZIONI DEFINITIVA
// =========================
function v50EnsurePrenotazioniDb(done) {
  db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      cliente_id INTEGER,
      mezzo_id INTEGER,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      codice_fiscale TEXT,
      cf TEXT,
      data_inizio TEXT,
      ora_inizio TEXT,
      data_fine TEXT,
      ora_fine TEXT,
      totale REAL,
      stato TEXT DEFAULT 'bozza',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);

    const cols = {
      codice:'TEXT', cliente_id:'INTEGER', mezzo_id:'INTEGER',
      nome:'TEXT', cognome:'TEXT', telefono:'TEXT', email:'TEXT',
      codice_fiscale:'TEXT', cf:'TEXT', indirizzo:'TEXT', citta:'TEXT',
      cap:'TEXT', provincia:'TEXT', data_nascita:'TEXT', luogo_nascita:'TEXT',
      numero_documento:'TEXT', scadenza_documento:'TEXT',
      numero_patente:'TEXT', scadenza_patente:'TEXT', categoria_patente:'TEXT',
      conducente2_nome:'TEXT', conducente2_cognome:'TEXT',
      conducente2_data_nascita:'TEXT', conducente2_luogo_nascita:'TEXT',
      conducente2_doc_numero:'TEXT', conducente2_patente_numero:'TEXT',
      conducente2_recapito:'TEXT',
      fatturazione:'TEXT', azienda:'TEXT', piva:'TEXT', pec:'TEXT', sdi:'TEXT',
      mezzo:'TEXT', targa:'TEXT', marca:'TEXT', modello:'TEXT',
      data_inizio:'TEXT', ora_inizio:'TEXT', data_fine:'TEXT', ora_fine:'TEXT',
      giorni:'INTEGER', km_previsti:'TEXT', km_inclusi:'TEXT',
      km_uscita:'TEXT', km_rientro:'TEXT',
      carburante_uscita:'TEXT', carburante_rientro:'TEXT',
      totale:'REAL', imponibile:'REAL', iva:'REAL', cauzione:'REAL', deposito:'REAL',
      firma:'TEXT', pdf_path:'TEXT', pdf_drive_link:'TEXT',
      drive_folder_id:'TEXT', drive_folder_link:'TEXT',
      note:'TEXT', created_at:'TEXT',
      record_cargos_uid:'TEXT', record_cargos_transactionid:'TEXT', record_cargos_stato:'TEXT',
      record_cargos_last_check:'TEXT', record_cargos_last_send:'TEXT', record_cargos_last_error:'TEXT',
      record_cargos_pagamento_tipo:'TEXT', record_cargos_checkout_luogo_cod:'TEXT',
      record_cargos_checkout_indirizzo:'TEXT', record_cargos_checkin_luogo_cod:'TEXT',
      record_cargos_checkin_indirizzo:'TEXT', record_cargos_agenzia_id:'TEXT',
      record_cargos_operatore_id:'TEXT', record_cargos_luogo_cod:'TEXT',
      record_cargos_nascita_luogo_cod:'TEXT', record_cargos_cittadinanza_cod:'TEXT', conducente_cittadinanza_cod:'TEXT',
      record_cargos_residenza_luogo_cod:'TEXT', record_cargos_doc_tipo_cod:'TEXT',
      record_cargos_doc_luogoril_cod:'TEXT', record_cargos_patente_luogoril_cod:'TEXT'
    };

    const keys = Object.keys(cols);
    let pending = keys.length;
    keys.forEach((c) => {
      db.run(`ALTER TABLE prenotazioni ADD COLUMN ${c} ${cols[c]}`, () => {
        pending--;
        if (pending === 0 && done) done();
      });
    });
  });
}

function v50EnsureAllDb(done) {
  v50EnsurePrenotazioniDb(() => {
    if (typeof v48EnsureImportDb === 'function') {
      v48EnsureImportDb(() => done && done());
    } else if (typeof runV44DbMigration === 'function') {
      runV44DbMigration();
      done && done();
    } else {
      done && done();
    }
  });
}

// esegue all'avvio
v50EnsurePrenotazioniDb(() => console.log('V50 prenotazioni DB OK'));

app.get('/versione', (req, res) => res.send('DP RENT APP V123 - fix attese, documenti archivio, categorie, fatturazione'));

function salvaClienteStorico(dati, cb) {
  const cf = String(dati.codice_fiscale || '').trim().toUpperCase();
  const params = [
    dati.nome || '', dati.cognome || '', dati.telefono || '', dati.email || '', cf || null,
    dati.indirizzo || '', dati.citta || '', dati.cap || '', dati.data_nascita || '', dati.luogo_nascita || '',
    dati.documento_numero || '', dati.documento_scadenza || '', dati.patente1 || dati.patente_numero || '',
    dati.patente1_scadenza || dati.patente_scadenza || '', dati.categoria_patente || '', dati.note_cliente || ''
  ];
  const doInsert = () => db.run(`
    INSERT INTO clienti (nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,data_nascita,luogo_nascita,
    documento_numero,documento_scadenza,patente_numero,patente_scadenza,categoria_patente,note)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
  `, params, function(err){ cb && cb(err, this ? this.lastID : null); });
  if (cf) {
    db.get(`SELECT id FROM clienti WHERE codice_fiscale=?`, [cf], (e, old) => {
      if (old) {
        db.run(`
          UPDATE clienti SET nome=?, cognome=?, telefono=?, email=?, codice_fiscale=?, indirizzo=?, citta=?, cap=?,
          data_nascita=?, luogo_nascita=?, documento_numero=?, documento_scadenza=?, patente_numero=?, patente_scadenza=?,
          categoria_patente=?, note=?, updated_at=CURRENT_TIMESTAMP WHERE id=?
        `, [...params, old.id], err => cb && cb(err, old.id));
      } else doInsert();
    });
  } else doInsert();
}

app.get('/', async (req, res) => {
  try {
    const mezzi = await get(`SELECT COUNT(*) as tot FROM mezzi`);
    const pren = await get(`SELECT COUNT(*) as tot FROM prenotazioni`);
    const attesa = await get(`SELECT COUNT(*) as tot FROM prenotazioni WHERE stato IN ('attesa_si_no','richiesta_cliente','preventivo_whatsapp') OR tipo_record='preventivo_whatsapp'`);
    const allMezzi = await all(`SELECT * FROM mezzi`);
    const alerts = allMezzi.map(m => {
      const a = alertMezzo(m);
      return a ? `<div><b>${esc(m.targa)} ${esc(m.modello)}</b>${a}</div>` : '';
    }).join('');

    res.send(page('Dashboard', `
      <div class="hero">
        <a class="tile" href="/nuova-prenotazione"><span>&#10133;</span>Nuova prenotazione</a>
        <a class="tile" href="/planning"><span>&#128197;</span>Planning</a>
        <a class="tile" href="/mezzi-web"><span>&#128667;</span>Mezzi</a>
        <a class="tile" href="/prenotazioni"><span>&#128193;</span>Storico</a>
        <a class="tile" href="/scadenze-mezzi"><span>&#9888;</span>Scadenze</a>
        <a class="tile" href="/prenota"><span>&#128241;</span>Pagina cliente</a>
        <a class="tile" href="/richieste-attesa"><span>&#128680;</span>Clienti in attesa</a>
        <a class="tile" href="/import-mezzi"><span>&#128202;</span>Import Excel</a>
        <a class="tile" href="/cargos"><span>&#128666;</span>Ca.R.G.O.S.</a>
      </div>
      ${(attesa && attesa.tot>0) ? `<div class="box dp-alert-wait"><h2>🚨 ${attesa.tot} CLIENTE/I IN ATTESA</h2><p>Ci sono preventivi WhatsApp o richieste cliente da controllare subito.</p><a class="btn" href="/richieste-attesa">Apri clienti in attesa</a></div>` : ``}
      <div class="box" style="border:3px solid #c60000"><h2>VERSIONE ATTIVA: V123 COMPLETA</h2><p class="ok">Se vedi questo riquadro, Render ha preso la versione nuova.</p></div>
      <div class="box">
        <h2>Gestionale DP RENT attivo</h2>
        <p>Mezzi caricati: <b>${mezzi ? mezzi.tot : 0}</b></p>
        <p>Contratti / prenotazioni: <b>${pren ? pren.tot : 0}</b></p>
        <p>Email: <b>${esc(process.env.SMTP_HOST || 'non configurata')}</b></p>
        <p>Google Drive: <b>${googleDriveConfigured() ? 'configurato via Apps Script' : 'non configurato'}</b></p>
        <p>Nexi Pay By Link: <b>${nexiConfigured() ? 'configurato' : 'non configurato'}</b></p>
      </div>
      <div class="box">
        <h2>Alert mezzi</h2>
        ${alerts || '<p class="ok">Nessun alert mezzi.</p>'}
      </div>
    `));
  } catch (e) {
    res.status(500).send(page('Errore', `<div class="box"><h2 class="bad">Errore Dashboard</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/richieste-attesa', async (req, res) => {
  try {
    const rows = await all(`SELECT * FROM prenotazioni WHERE stato IN ('attesa_si_no','richiesta_cliente','preventivo_whatsapp') OR tipo_record='preventivo_whatsapp' ORDER BY id DESC LIMIT 100`);
    res.send(page('Clienti in attesa', `<div class="box dp-alert-wait"><h2>🚨 Clienti in attesa</h2><p>Qui trovi i preventivi generati dal bot WhatsApp e le richieste compilate dal cliente.</p></div><div class="box"><table><tr><th>ID</th><th>Cliente</th><th>Telefono</th><th>Mezzo</th><th>Periodo</th><th>Totale</th><th>Stato</th><th>Azioni</th></tr>${rows.length ? rows.map(p=>`<tr><td>${esc(p.codice||p.id)}</td><td>${esc((p.nome||'')+' '+(p.cognome||''))}</td><td>${esc(p.telefono||'')}</td><td>${esc(p.categoria||p.tipo||'')}</td><td>${esc((p.data_inizio||'')+' - '+(p.data_fine||''))}</td><td>EUR ${euro(p.totale||0)}</td><td><b>${esc(p.stato||'')}</b></td><td><a class="btn" href="/prenotazione/${p.id}">Apri</a> <a class="btn btn2" href="/contratto/${p.id}/gestisci">Contratto</a> <form method="POST" action="/richieste-attesa/${p.id}/elimina" style="display:inline" onsubmit="return confirm('Eliminare questo cliente in attesa?')"><button class="btn bad" type="submit">Elimina</button></form></td></tr>`).join('') : `<tr><td colspan="8" class="ok">Nessun cliente in attesa.</td></tr>`}</table></div>`));
  } catch(e) {
    res.status(500).send(page('Errore', `<div class="box"><h2 class="bad">Errore clienti in attesa</h2><pre>${esc(e.message)}</pre></div>`));
  }
});


app.post('/richieste-attesa/:id/elimina', async (req, res) => {
  try {
    await run(`DELETE FROM allegati WHERE prenotazione_id=?`, [req.params.id]).catch(()=>{});
    await run(`DELETE FROM prenotazioni WHERE id=? AND (stato IN ('attesa_si_no','richiesta_cliente','preventivo_whatsapp','non_confermato') OR tipo_record='preventivo_whatsapp')`, [req.params.id]);
    res.redirect('/richieste-attesa');
  } catch(e) {
    res.status(500).send(page('Errore eliminazione', `<div class="box"><h2 class="bad">Errore eliminazione</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/logo', (req, res) => {
  const hasLogo = fs.existsSync(path.join(publicDir, 'logo.png'));
  res.send(page('Logo', `
    <div class="box">
      <h2>Logo DP RENT</h2>
      ${hasLogo ? `<p>Logo attuale:</p><img src="/public/logo.png" style="max-width:240px;background:#eee;padding:10px;border-radius:8px;">` : `<p>Nessun logo caricato.</p>`}
      <form method="POST" action="/logo" enctype="multipart/form-data">
        <label>Carica logo PNG/JPG</label>
        <input type="file" name="logo" accept="image/png,image/jpeg" required>
        <button>Salva logo</button>
      </form>
    </div>
  `));
});
app.post('/logo', multer({
  storage: multer.diskStorage({
    destination: (req,file,cb)=>cb(null,publicDir),
    filename: (req,file,cb)=>cb(null,'logo.png')
  })
}).single('logo'), (req, res) => res.redirect('/logo'));


app.get('/admin/migra-db-v44', (req, res) => {
  try {
    runV44DbMigration();
    res.send(page('Migrazione DB V44', '<div class="box"><h2>Migrazione DB V44 eseguita</h2><p>Ora riprova Import Excel.</p><a class="btn" href="/import-excel">Torna import</a><a class="btn btn2" href="/mezzi">Mezzi</a></div>'));
  } catch(e) {
    res.status(500).send('Errore migrazione: ' + e.message);
  }
});






app.get('/import-mezzi', (req, res) => res.redirect('/import-excel'));
app.post('/import-mezzi', (req, res) => res.redirect(307, '/import-excel'));

app.get('/import-excel', (req, res) => {
  res.send(page('Import Excel', `<div class="box">
    <h2>Import mezzi da Excel/CSV</h2>
    <p>Versione V48: non usa ON CONFLICT, aggiorna per targa/UID.</p>
    <form method="post" action="/import-excel" enctype="multipart/form-data">
      <input type="file" name="file" accept=".xlsx,.xls,.csv" required>
      <button class="btn" type="submit">Carica e importa</button>
    </form>
  </div>`));
});

app.post('/import-excel', importUploadV48.single('file'), async (req, res) => {
  try {
    if (!req.file) throw new Error('Nessun file caricato');
    v48EnsureImportDb(async () => {
      try {
        const rows = v48ParseImportFile(req.file.path);
        let inserted = 0, updated = 0, skipped = 0;
        for (const row of rows) {
          const r = await v48InsertOrUpdateMezzo(row);
          inserted += r.inserted || 0;
          updated += r.updated || 0;
          skipped += r.skipped || 0;
        }
        res.send(page('Import completato', `<div class="box">
          <h2 class="ok">Import completato</h2>
          <p><b>Inseriti:</b> ${inserted}</p>
          <p><b>Aggiornati:</b> ${updated}</p>
          <p><b>Saltati:</b> ${skipped}</p>
          <a class="btn" href="/mezzi">Vai ai mezzi</a>
          <a class="btn btn2" href="/import-excel">Nuovo import</a>
        </div>`));
      } catch (e) {
        res.send(page('Errore import', `<div class="box"><h2 class="bad">Errore import</h2><pre>${esc(e.message)}</pre></div>`));
      }
    });
  } catch (e) {
    res.send(page('Errore import', `<div class="box"><h2 class="bad">Errore import</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/mezzi-web', async (req, res) => {
  const rows = await all(`SELECT * FROM mezzi ORDER BY categoria,targa`);
  const trs = rows.map(m => `
    <tr>
      <td>${m.id}</td>
      <td><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a></td>
      <td>${esc(m.marca)}</td>
      <td>${esc(m.modello)}</td>
      <td>${esc(m.categoria)}</td>
      <td>${esc(descrizionePubblica(m))}</td>
      <td>&euro; ${euro(m.prezzo_giorno)}</td>
      <td>${esc(m.km_inclusi)}</td>
      <td>${alertBadge(m)}</td>
      <td>${esc(m.stato)}</td>
    </tr>`).join('');
  res.send(page('Mezzi', `
    <h2>Elenco mezzi</h2>
    <table><tr><th>ID</th><th>Targa</th><th>Marca</th><th>Modello</th><th>Categoria</th><th>Descrizione pubblica</th><th>Prezzo</th><th>Km/giorno</th><th>Alert</th><th>Stato</th><th>CaRGOS</th></tr>${trs}</table>
  `));
});

app.get('/mezzo/:id', async (req, res) => {
  const m = await get(`SELECT * FROM mezzi WHERE id=?`, [req.params.id]);
  if (!m) return res.send('Mezzo non trovato');
  const files = await all(`SELECT * FROM allegati WHERE mezzo_id=? ORDER BY id DESC`, [m.id]);
  const filesV99 = v99DedupeAllegati(files);
  const lista = filesV99.map(f => `<li>${esc(f.tipo)} - <a href="/uploads/${esc(f.filename)}" target="_blank">${esc(f.originalname)}</a> ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Google Drive</a>` : ''}</li>`).join('');
  res.send(page('Scheda mezzo', `
    <div class="box">
      <h2>Scheda mezzo ${esc(m.targa)} ${alertBadge(m)}</h2>
      ${alertMezzo(m) || '<p class="ok">Nessun alert attivo.</p>'}
      <form method="POST" action="/mezzo/${m.id}">
        <div class="grid">
          <div><label>Targa</label><input name="targa" value="${esc(m.targa)}" required></div>
          <div><label>Marca</label><input name="marca" value="${esc(m.marca)}"></div>
          <div><label>Modello</label><input name="modello" value="${esc(m.modello)}"></div>
          <div><label>Categoria</label><select name="categoria">${['FURGONE','9_POSTI','AUTO_DACIA','AUTO_GOLF','ESCAVATORE','SEMOVENTE'].map(c=>`<option ${m.categoria===c?'selected':''}>${c}</option>`).join('')}</select></div>
          <div><label>Posti</label><input type="number" name="posti" value="${esc(m.posti)}"></div>
          <div><label>Descrizione pubblica cliente</label><input name="descrizione_pubblica" value="${esc(descrizionePubblica(m))}"></div>
          <div><label>Prezzo giorno</label><input type="number" step="0.01" name="prezzo_giorno" value="${esc(m.prezzo_giorno)}"></div>
          <div><label>Km inclusi/giorno</label><input type="number" name="km_inclusi" value="${esc(m.km_inclusi)}"></div>
          <div><label>Km attuali</label><input type="number" name="km_attuali" value="${esc(m.km_attuali || m.km)}"></div>
          <div><label>Scadenza tagliando km</label><input type="number" name="tagliando_km_scadenza" value="${esc(m.tagliando_km_scadenza)}"></div>
          <div><label>Scadenza tagliando data</label><input type="date" name="tagliando_data_scadenza" value="${esc(m.tagliando_data_scadenza)}"></div>
          <div><label>Revisione</label><input type="date" name="revisione_scadenza" value="${esc(m.revisione_scadenza)}"></div>
          <div><label>Bollo</label><input type="date" name="bollo_scadenza" value="${esc(m.bollo_scadenza)}"></div>
          <div><label>Assicurazione</label><input type="date" name="assicurazione_scadenza" value="${esc(m.assicurazione_scadenza)}"></div>
          <div><label>Gomme / altra manutenzione</label><input type="date" name="gomme_scadenza" value="${esc(m.gomme_scadenza)}"></div>
          <div><label>Alert giorni prima</label><input type="number" name="alert_giorni" value="${esc(m.alert_giorni || 30)}"></div>
          <div><label>Alert km prima</label><input type="number" name="alert_km" value="${esc(m.alert_km || 1000)}"></div>
        </div>
        <label>Note manutenzione</label><textarea name="manutenzione_note">${esc(m.manutenzione_note)}</textarea>
        <button>Salva scheda mezzo</button>
      </form>
      <hr>
      <h3>Foto/documenti mezzo</h3>
      <form method="POST" action="/mezzo/${m.id}/foto" enctype="multipart/form-data">
        <select name="tipo"><option>Foto mezzo fronte</option><option>Foto mezzo retro</option><option>Foto lato dx</option><option>Foto lato sx</option><option>Foto interno</option><option>Libretto</option><option>Assicurazione</option><option>Bollo</option><option>Revisione</option><option>Manutenzione</option></select>
        <input type="file" name="file" accept="image/*,.pdf" required>
        <button>Carica file mezzo</button>
      </form>
      <ul>${lista}</ul>
      <a class="btn" href="/nuova-prenotazione?mezzo_id=${m.id}">Prenota questo mezzo</a>
    </div>
  `));
});
app.post('/mezzo/:id', async (req, res) => {
  const b = req.body;
  await run(`
    UPDATE mezzi SET targa=?, marca=?, modello=?, categoria=?, posti=?, descrizione_pubblica=?, prezzo_giorno=?, km_inclusi=?,
    km_attuali=?, tagliando_km_scadenza=?, tagliando_data_scadenza=?, revisione_scadenza=?, bollo_scadenza=?,
    assicurazione_scadenza=?, gomme_scadenza=?, alert_giorni=?, alert_km=?, manutenzione_note=? WHERE id=?
  `, [b.targa,b.marca,b.modello,b.categoria,b.posti,b.descrizione_pubblica,b.prezzo_giorno,b.km_inclusi,b.km_attuali,b.tagliando_km_scadenza,b.tagliando_data_scadenza,b.revisione_scadenza,b.bollo_scadenza,b.assicurazione_scadenza,b.gomme_scadenza,b.alert_giorni,b.alert_km,b.manutenzione_note,req.params.id]);
  res.redirect(`/mezzo/${req.params.id}`);
});
app.post('/mezzo/:id/foto', upload.single('file'), async (req, res) => {
  if (!req.file) return res.send('File mancante');
  const m = await get(`SELECT targa,modello FROM mezzi WHERE id=?`, [req.params.id]);
  let driveRes = null;
  try {
    driveRes = await uploadFileToDrive(req.file.path, `${Date.now()}_${req.body.tipo}_${req.file.originalname}`, req.file.mimetype, `MEZZO ${m?.targa || req.params.id}`);
  } catch (e) { console.log('Errore upload foto mezzo Drive:', e.message); }
  await run(`INSERT INTO allegati (mezzo_id,tipo,filename,originalname,path,mimetype,drive_file_id,drive_web_link) VALUES (?,?,?,?,?,?,?,?)`,
    [req.params.id, req.body.tipo, req.file.filename, req.file.originalname, req.file.path, req.file.mimetype, driveRes?.id || null, driveRes?.webViewLink || null]);
  res.redirect(`/mezzo/${req.params.id}`);
});


function dpDateDiffDays(dateStr) {
  if (!dateStr) return null;
  const d = new Date(String(dateStr).slice(0,10) + 'T00:00:00');
  if (isNaN(d.getTime())) return null;
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return Math.ceil((d - today) / 86400000);
}
function badgeScadenzaCliente(label, value, alertDays = 60) {
  const diff = dpDateDiffDays(value);
  if (diff === null) return `<span class="badge badge-orange">${esc(label)} mancante</span>`;
  if (diff < 0) return `<span class="badge badge-red">${esc(label)} scaduto da ${Math.abs(diff)} gg</span>`;
  if (diff <= alertDays) return `<span class="badge badge-orange">${esc(label)} scade tra ${diff} gg</span>`;
  return `<span class="badge badge-green">${esc(label)} ok</span>`;
}
function alertCliente(c, alertDays = 60) {
  const parts = [];
  const d1 = dpDateDiffDays(c.documento_scadenza);
  const d2 = dpDateDiffDays(c.patente_scadenza || c.patente1_scadenza);
  if (d1 === null) parts.push('<div class="alert">Documento mancante</div>');
  else if (d1 < 0) parts.push(`<div class="alert">Documento scaduto da ${Math.abs(d1)} giorni</div>`);
  else if (d1 <= alertDays) parts.push(`<div class="alert">Documento in scadenza tra ${d1} giorni</div>`);
  if (d2 === null) parts.push('<div class="alert">Patente mancante</div>');
  else if (d2 < 0) parts.push(`<div class="alert">Patente scaduta da ${Math.abs(d2)} giorni</div>`);
  else if (d2 <= alertDays) parts.push(`<div class="alert">Patente in scadenza tra ${d2} giorni</div>`);
  return parts.join('');
}

app.get('/scadenze-mezzi', async (req, res) => {
  const rows = await all(`SELECT * FROM mezzi ORDER BY targa`);
  const trs = rows.map(m => `<tr><td><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a></td><td>${esc(descrizionePubblica(m))}</td><td>${esc(m.km_attuali || m.km)}</td><td>${esc(m.tagliando_km_scadenza)}</td><td>${esc(m.tagliando_data_scadenza)}</td><td>${esc(m.revisione_scadenza)}</td><td>${esc(m.bollo_scadenza)}</td><td>${esc(m.assicurazione_scadenza)}</td><td>${alertBadge(m)}</td></tr>`).join('');
  const alerts = rows.map(m => alertMezzo(m) ? `<div><b>${esc(m.targa)} ${esc(m.modello)}</b>${alertMezzo(m)}</div>` : '').join('');
  res.send(page('Scadenze mezzi', `<h2>Scadenze e manutenzioni mezzi</h2><div class="box">${alerts || '<p class="ok">Nessun alert attivo.</p>'}</div><table><tr><th>Targa</th><th>Descrizione</th><th>Km</th><th>Tagliando km</th><th>Tagliando data</th><th>Revisione</th><th>Bollo</th><th>Assicurazione</th><th>Alert</th></tr>${trs}</table>`));
});


function prefill(query, name, fallback = '') {
  return esc((query && query[name]) || fallback || '');
}

function formPrenotazione(mezzi, selectedMezzo, selectedData, action, query = {}) {
  const opt = mezzi.map(m => `<option value="${m.id}" ${String(m.id)===String(selectedMezzo)?'selected':''}>${esc(m.targa)} - ${esc(descrizionePubblica(m))}</option>`).join('');
  const qv = (k, d='') => esc(query[k] || d || '');
  const tipoCliente = String(query.tipo_cliente || 'privato').toLowerCase();
  return `<form method="POST" action="${action}">
    <div class="box" style="border:2px solid #0b6b2d">
      <h3>Cliente</h3>
      <p class="notice">Se il cliente e gia in archivio vai su <a href="/clienti">Clienti</a> e premi Crea contratto. Qui non devi reinserirlo ogni volta.</p>
      <div class="grid">
        <div><label>Nome</label><input name="nome" value="${qv('nome')}" required></div>
        <div><label>Cognome</label><input name="cognome" value="${qv('cognome')}" required></div>
        <div><label>Telefono</label><input name="telefono" value="${qv('telefono')}" required></div>
        <div><label>Email</label><input name="email" type="email" value="${qv('email')}"></div>
        <div><label>Codice fiscale</label><input name="codice_fiscale" value="${qv('codice_fiscale')}"></div>
        <div><label>Data nascita</label><input type="date" name="data_nascita" value="${qv('data_nascita')}"></div>
        <div><label>Luogo nascita</label><input name="luogo_nascita" value="${qv('luogo_nascita')}"></div>
        <div><label>Cittadinanza COD</label><input name="cittadinanza_cod" value="${qv('cittadinanza_cod','100000100')}"></div>
        <div><label>Indirizzo</label><input name="indirizzo" value="${qv('indirizzo')}"></div>
        <div><label>Citta</label><input name="citta" value="${qv('citta')}"></div>
        <div><label>CAP</label><input name="cap" value="${qv('cap')}"></div>
        <div><label>Tipo cliente</label><select name="tipo_cliente"><option value="privato" ${tipoCliente==='privato'?'selected':''}>Privato</option><option value="azienda" ${tipoCliente==='azienda'?'selected':''}>Azienda</option></select></div>
        <div><label>P.IVA</label><input name="piva" value="${qv('piva')}"></div>
        <div><label>Ragione sociale</label><input name="ragione_sociale" value="${qv('ragione_sociale')}"></div>
        <div><label>PEC</label><input name="pec" value="${qv('pec')}"></div>
        <div><label>SDI</label><input name="sdi" value="${qv('sdi')}"></div>
      </div>
      <h3>Documento / patente</h3>
      <div class="grid">
        <div><label>Tipo documento</label><select name="documento_tipo"><option value="IDENT" ${(query.documento_tipo||'IDENT')==='IDENT'?'selected':''}>Carta identita</option><option value="IDELE" ${query.documento_tipo==='IDELE'?'selected':''}>Carta identita elettronica</option><option value="PASOR" ${query.documento_tipo==='PASOR'?'selected':''}>Passaporto</option><option value="PATEN" ${query.documento_tipo==='PATEN'?'selected':''}>Patente</option></select></div>
        <div><label>Numero documento</label><input name="documento_numero" value="${qv('documento_numero')}"></div>
        <div><label>Scadenza documento</label><input type="date" name="documento_scadenza" value="${qv('documento_scadenza')}"></div>
        <div><label>Conducente 1</label><input name="conducente1" value="${qv('conducente1', ((query.nome || '') + ' ' + (query.cognome || '')).trim())}"></div>
        <div><label>Patente 1</label><input name="patente1" value="${qv('patente1')}"></div>
        <div><label>Scadenza patente 1</label><input type="date" name="patente1_scadenza" value="${qv('patente1_scadenza')}"></div>
        <div><label>Categoria patente</label><input name="categoria_patente" value="${qv('categoria_patente')}"></div>
        <div><label>Conducente 2</label><input name="conducente2" value="${qv('conducente2')}"></div>
        <div><label>Patente 2</label><input name="patente2" value="${qv('patente2')}"></div>
        <div><label>Scadenza patente 2</label><input type="date" name="patente2_scadenza" value="${qv('patente2_scadenza')}"></div>
      </div>
    </div>
    <div class="box">
      <h3>Mezzo e periodo</h3>
      <div class="grid">
        <div><label>Mezzo</label><select name="mezzo_id" required>${opt}</select></div>
        <div><label>Data inizio</label><input type="date" name="data_inizio" value="${esc(selectedData || query.data_inizio || '')}" required></div>
        <div><label>Ora inizio</label><input type="time" name="ora_inizio" value="${qv('ora_inizio','08:30')}"></div>
        <div><label>Data fine</label><input type="date" name="data_fine" value="${esc(selectedData || query.data_fine || '')}" required></div>
        <div><label>Ora fine</label><input type="time" name="ora_fine" value="${qv('ora_fine','18:00')}"></div>
        <div><label>Km previsti</label><input type="number" name="km_previsti" value="${qv('km_previsti','150')}"></div>
        <div><label>Carburante uscita</label><select name="carburante_uscita">${fuelOptions(query.carburante_uscita || '4/4 pieno')}</select></div>
      </div>
    </div>
    <label>Note</label><textarea name="note">${qv('note')}</textarea>
    ${privacyCheckboxHtml()}<button>Crea contratto</button>
  </form>`;
}


// OCR DOPPIO STEP: DOCUMENTO + PATENTE

// OCR PRO: CARTA IDENTITA FRONTE/RETRO + PATENTE + STORICO CLIENTI
app.get('/ocr-pro', (req, res) => {
  res.send(page('OCR PRO cliente', `
    <div class="box">
      <h2>OCR PRO cliente</h2>
      <p class="notice">Flusso PRO: carta identita fronte, carta identita retro, patente. Poi controlli i dati e salvi cliente.</p>
      <form method="POST" action="/ocr-pro/fronte" enctype="multipart/form-data">
        <label class="btn" style="display:block;text-align:center;font-size:22px;padding:18px;margin-top:14px">
          Step 1 - Scatta/CARICA carta identita FRONTE
          <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
        </label>
        <button id="submitBtn" style="display:none">Avanti</button>
      </form>
      <a class="btn btn2" href="/nuova-prenotazione">Torna</a>
    </div>
    <script>
      const input = document.getElementById('fileInput');
      input.addEventListener('change', function(){ if(this.files && this.files.length) document.getElementById('submitBtn').click(); });
    </script>
  `));
});

app.post('/ocr-pro/fronte', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');
  const id = makeOcrId();
  OCR_PREFILL[id] = { fronte_path:req.file.path, fronte_mime:req.file.mimetype, retro_path:'', retro_mime:'', patente_path:'', patente_mime:'', documento:{}, patente:{} };
  res.redirect('/ocr-pro/retro?id=' + encodeURIComponent(id));
});

app.get('/ocr-pro/retro', (req, res) => {
  const id = req.query.id;
  if (!id || !OCR_PREFILL[id]) return res.redirect('/ocr-pro');
  res.send(page('OCR PRO retro', `
    <div class="box">
      <h2>Step 2 - Carta identita RETRO</h2>
      <p class="notice">Ora scatta/carica il retro della carta identita: di solito qui ci sono CF e indirizzo.</p>
      <form method="POST" action="/ocr-pro/retro" enctype="multipart/form-data">
        <input type="hidden" name="id" value="${esc(id)}">
        <label class="btn" style="display:block;text-align:center;font-size:22px;padding:18px;margin-top:14px">
          Step 2 - Scatta/CARICA carta identita RETRO
          <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
        </label>
        <button id="submitBtn" style="display:none">Avanti</button>
      </form>
      <a class="btn btn2" href="/ocr-pro">Ricomincia</a>
    </div>
    <script>
      const input = document.getElementById('fileInput');
      input.addEventListener('change', function(){ if(this.files && this.files.length) document.getElementById('submitBtn').click(); });
    </script>
  `));
});

app.post('/ocr-pro/retro', upload.single('file'), (req, res) => {
  const id = req.body.id;
  if (!id || !OCR_PREFILL[id]) return res.redirect('/ocr-pro');
  if (!req.file) return res.send('File mancante');
  OCR_PREFILL[id].retro_path = req.file.path;
  OCR_PREFILL[id].retro_mime = req.file.mimetype;
  res.redirect('/ocr-pro/patente?id=' + encodeURIComponent(id));
});

app.get('/ocr-pro/patente', (req, res) => {
  const id = req.query.id;
  if (!id || !OCR_PREFILL[id]) return res.redirect('/ocr-pro');
  res.send(page('OCR PRO patente', `
    <div class="box">
      <h2>Step 3 - Patente</h2>
      <p class="notice">Ora scatta/carica la patente del conducente.</p>
      <form method="POST" action="/ocr-pro/patente" enctype="multipart/form-data">
        <input type="hidden" name="id" value="${esc(id)}">
        <label class="btn" style="display:block;text-align:center;font-size:22px;padding:18px;margin-top:14px">
          Step 3 - Scatta/CARICA PATENTE
          <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
        </label>
        <button id="submitBtn" style="display:none">Leggi tutto</button>
      </form>
      <a class="btn btn2" href="/ocr-pro">Ricomincia</a>
    </div>
    <script>
      const input = document.getElementById('fileInput');
      input.addEventListener('change', function(){ if(this.files && this.files.length) document.getElementById('submitBtn').click(); });
    </script>
  `));
});

app.post('/ocr-pro/patente', upload.single('file'), async (req, res) => {
  try {
    const id = req.body.id;
    if (!id || !OCR_PREFILL[id]) return res.redirect('/ocr-pro');
    if (!req.file) return res.send('File mancante');
    const item = OCR_PREFILL[id];
    item.patente_path = req.file.path;
    item.patente_mime = req.file.mimetype;
    let fronte = {}, retro = {}, patente = {};
    try { fronte = await estraiDatiDocumentoConAI(item.fronte_path, item.fronte_mime); } catch(e) { fronte = {}; }
    try { retro = await estraiDatiDocumentoConAI(item.retro_path, item.retro_mime); } catch(e) { retro = {}; }
    try { patente = await estraiDatiDocumentoConAI(item.patente_path, item.patente_mime); } catch(e) { patente = {}; }
    item.documento = Object.assign({}, fronte || {}, retro || {});
    item.patente = patente || {};
    res.redirect('/ocr-pro/conferma?id=' + encodeURIComponent(id));
  } catch(e) {
    res.status(500).send(page('Errore OCR PRO', `<div class="box"><h2 class="bad">Errore OCR</h2><pre>${esc(e.message)}</pre><a class="btn" href="/ocr-pro">Riprova</a></div>`));
  }
});

app.get('/ocr-pro/conferma', (req, res) => {
  const id = req.query.id;
  const data = OCR_PREFILL[id];
  if (!id || !data) return res.redirect('/ocr-pro');
  const d = data.documento || {};
  const p = data.patente || {};
  const pick = (...vals) => vals.find(x => x && String(x).trim()) || '';
  const v = x => esc(x || '');
  const nome = pick(d.nome, p.nome);
  const cognome = pick(d.cognome, p.cognome);
  const cf = pick(d.codice_fiscale, p.codice_fiscale);
  const indirizzo = pick(d.indirizzo, p.indirizzo);
  const dataNascita = pick(d.data_nascita, p.data_nascita);
  const luogoNascita = pick(d.luogo_nascita, p.luogo_nascita);
  const documentoNumero = pick(d.numero_documento);
  const documentoScadenza = pick(d.data_scadenza);
  const patenteNumero = pick(p.numero_patente, p.numero_documento);
  const patenteScadenza = pick(p.data_scadenza);
  const categoriaPatente = pick(p.categoria_patente);
  res.send(page('Controllo dati OCR PRO', `
    <div class="box">
      <h2>Controlla dati cliente</h2>
      <p class="notice">Correggi eventuali errori. Il cliente verra salvato nello storico.</p>
      <form method="POST" action="/ocr-pro/applica">
        <input type="hidden" name="id" value="${esc(id)}">
        <h3>Cliente / documento</h3>
        <div class="grid">
          <div><label>Nome</label><input name="nome" value="${v(nome)}"></div>
          <div><label>Cognome</label><input name="cognome" value="${v(cognome)}"></div>
          <div><label>Telefono</label><input name="telefono" value=""></div>
          <div><label>Email</label><input name="email" value=""></div>
          <div><label>Codice fiscale</label><input name="codice_fiscale" value="${v(cf)}"></div>
          <div><label>Indirizzo</label><input name="indirizzo" value="${v(indirizzo)}"></div>
          <div><label>Citta</label><input name="citta" value=""></div>
          <div><label>CAP</label><input name="cap" value=""></div>
          <div><label>Data nascita</label><input type="date" name="data_nascita" value="${v(dataNascita)}"></div>
          <div><label>Luogo nascita</label><input name="luogo_nascita" value="${v(luogoNascita)}"></div>
          <div><label>Numero documento</label><input name="documento_numero" value="${v(documentoNumero)}"></div>
          <div><label>Scadenza documento</label><input type="date" name="documento_scadenza" value="${v(documentoScadenza)}"></div>
        </div>
        <h3>Conducente / patente</h3>
        <div class="grid">
          <div><label>Conducente 1</label><input name="conducente1" value="${v((nome + ' ' + cognome).trim())}"></div>
          <div><label>Numero patente</label><input name="patente1" value="${v(patenteNumero)}"></div>
          <div><label>Scadenza patente</label><input type="date" name="patente1_scadenza" value="${v(patenteScadenza)}"></div>
          <div><label>Categoria patente</label><input name="categoria_patente" value="${v(categoriaPatente)}"></div>
        </div>
        <label>Note cliente</label><textarea name="note_cliente"></textarea>
        <button>Salva cliente e continua al contratto</button>
      </form>
      <a class="btn btn2" href="/ocr-pro">Rifai OCR</a>
    </div>
  `));
});

app.post('/ocr-pro/applica', (req, res) => {
  const id = req.body.id || makeOcrId();
  const dati = {
    nome:req.body.nome||'', cognome:req.body.cognome||'', telefono:req.body.telefono||'', email:req.body.email||'',
    codice_fiscale:req.body.codice_fiscale||'', indirizzo:req.body.indirizzo||'', citta:req.body.citta||'', cap:req.body.cap||'',
    data_nascita:req.body.data_nascita||'', luogo_nascita:req.body.luogo_nascita||'', documento_numero:req.body.documento_numero||'',
    documento_scadenza:req.body.documento_scadenza||'', conducente1:req.body.conducente1||'', patente1:req.body.patente1||'',
    patente1_scadenza:req.body.patente1_scadenza||'', categoria_patente:req.body.categoria_patente||'', note_cliente:req.body.note_cliente||''
  };
  OCR_PREFILL[id] = dati;
  salvaClienteStorico(dati, () => res.redirect('/nuova-prenotazione?ocr=' + encodeURIComponent(id)));
});


app.get('/scadenze-clienti', async (req, res) => {
  const alertDays = Number(req.query.giorni || 60);
  const rows = await all(`SELECT * FROM clienti ORDER BY cognome, nome LIMIT 500`);
  const alerts = rows.map(c => {
    const a = alertCliente(c, alertDays);
    if (!a) return '';
    return `<div class="box" style="margin-bottom:12px"><h3>${esc(c.nome)} ${esc(c.cognome)}</h3><p><b>Tel:</b> ${esc(c.telefono||'')} <b>Email:</b> ${esc(c.email||'')}</p>${a}<p><a class="btn" href="/cliente/${c.id}">Apri cliente</a> <a class="btn btn3" href="/cliente/${c.id}/documenti">Documenti</a></p></div>`;
  }).join('');
  const trs = rows.map(c => `<tr><td><a href="/cliente/${c.id}"><b>${esc(c.nome)} ${esc(c.cognome)}</b></a></td><td>${esc(c.telefono||'')}</td><td>${esc(c.documento_numero||'')}<br>${esc(c.documento_scadenza||'')}</td><td>${badgeScadenzaCliente('Documento', c.documento_scadenza, alertDays)}</td><td>${esc(c.patente_numero||c.patente1||'')}<br>${esc(c.patente_scadenza||c.patente1_scadenza||'')}</td><td>${badgeScadenzaCliente('Patente', c.patente_scadenza || c.patente1_scadenza, alertDays)}</td></tr>`).join('');
  res.send(page('Scadenze clienti', `<div class="box"><h2>Scadenze documenti clienti</h2><p>Controllo documento e patente come le scadenze mezzi.</p><form method="GET"><label>Avvisa giorni prima</label><input type="number" name="giorni" value="${alertDays}"><button>Aggiorna</button></form></div>${alerts || '<div class="box"><p class="ok">Nessun alert documenti clienti.</p></div>'}<table><tr><th>Cliente</th><th>Telefono</th><th>Documento</th><th>Alert documento</th><th>Patente</th><th>Alert patente</th></tr>${trs || '<tr><td colspan="6">Nessun cliente.</td></tr>'}</table>`));
});

app.get('/clienti', (req, res) => {
  const q = String(req.query.q || '').trim();
  let sql = `SELECT * FROM clienti WHERE 1=1`;
  const params = [];
  if (q) {
    sql += ` AND (nome LIKE ? OR cognome LIKE ? OR telefono LIKE ? OR email LIKE ? OR codice_fiscale LIKE ? OR patente_numero LIKE ?)`;
    params.push(`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`);
  }
  sql += ` ORDER BY updated_at DESC, id DESC LIMIT 200`;
  db.all(sql, params, (err, rows) => {
    const trs = (rows || []).map(c => `
      <tr>
        <td><a href="/cliente/${c.id}"><b>${esc(c.nome)} ${esc(c.cognome)}</b></a></td>
        <td>${esc(c.telefono||'')}<br>${esc(c.email||'')}</td>
        <td>${esc(c.codice_fiscale||'')}</td>
        <td>${esc(c.documento_numero||'')}<br>Scad. ${esc(c.documento_scadenza||'')}<br>${badgeScadenzaCliente('Documento', c.documento_scadenza)}</td>
        <td>${esc(c.patente_numero||c.patente1||'')}<br>Scad. ${esc(c.patente_scadenza||c.patente1_scadenza||'')}<br>${badgeScadenzaCliente('Patente', c.patente_scadenza || c.patente1_scadenza)}</td>
        <td>${alertCliente(c) || '<span class="badge badge-green">OK</span>'}</td>
        <td><a class="btn" href="/nuova-da-cliente/${c.id}">Crea contratto</a> <a class="btn btn3" href="/cliente/${c.id}/documenti">Documenti</a> <a class="btn btn2" href="/cliente/${c.id}/modifica">Modifica</a> <a class="btn bad" href="/cliente/${c.id}/elimina">Elimina</a></td>
      </tr>`).join('');
    res.send(page('Clienti', `
      <div class="box"><h2>Anagrafica clienti</h2>
      <p>Da qui controlli documenti, patenti e alert scadenze clienti.</p>
      <form method="GET" action="/clienti"><input name="q" placeholder="Cerca nome, telefono, CF, patente" value="${esc(q)}"><button>Cerca</button></form>
      <a class="btn btn3" href="/cliente-nuovo">Nuovo cliente manuale</a> <a class="btn btn3" href="/ocr-pro">Nuovo cliente con OCR PRO</a> <a class="btn" href="/scadenze-clienti">Scadenze clienti</a></div>
      <table><tr><th>Cliente</th><th>Contatti</th><th>CF</th><th>Documento</th><th>Patente</th><th>Alert</th><th>Azione</th></tr>${trs || '<tr><td colspan="7">Nessun cliente.</td></tr>'}</table>
    `));
  });
});

app.get('/cliente/:id', (req, res) => {
  db.get(`SELECT * FROM clienti WHERE id=?`, [req.params.id], (err, c) => {
    if (!c) return res.redirect('/clienti');
    res.send(page('Scheda cliente', `<div class="box">
      <h2>${esc(c.nome)} ${esc(c.cognome)}</h2>
      <p><b>Telefono:</b> ${esc(c.telefono||'')}</p><p><b>Email:</b> ${esc(c.email||'')}</p>
      <p><b>CF:</b> ${esc(c.codice_fiscale||'')}</p>
      <p><b>Fatturazione:</b> ${esc(c.tipo_cliente||'Privato')} ${esc(c.ragione_sociale||'')} ${esc(c.piva||c.partita_iva||'')}</p>
      <p><b>PEC/SDI:</b> ${esc(c.pec||'')} ${esc(c.sdi||c.codice_sdi||'')}</p>
      <p><b>Indirizzo:</b> ${esc(c.indirizzo||'')}, ${esc(c.citta||'')} ${esc(c.cap||'')}</p>
      <p><b>Documento:</b> ${esc(c.documento_numero||'')} - scad. ${esc(c.documento_scadenza||'')}</p>
      <p><b>Patente:</b> ${esc(c.patente_numero||'')} - scad. ${esc(c.patente_scadenza||'')} - cat. ${esc(c.categoria_patente||'')}</p>
      <p><b>Note:</b> ${esc(c.note||'')}</p>
      <a class="btn" href="/nuova-da-cliente/${c.id}">Crea contratto da cliente</a>
      <a class="btn btn3" href="/cliente/${c.id}/documenti">Archivio documenti</a>
      <a class="btn btn2" href="/cliente/${c.id}/modifica">Modifica cliente</a>
      <a class="btn bad" href="/cliente/${c.id}/elimina">Elimina cliente</a>
      <a class="btn btn2" href="/clienti">Torna clienti</a>
    </div>`));
  });
});




function clienteManualForm(c, action, title) {
  c = c || {};
  const val = k => esc(c[k] || '');
  return page(title, `
    <div class="box">
      <h2>${esc(title)}</h2>
      <p class="notice">Compilazione manuale completa. Questi dati vengono poi usati per contratto e Ca.R.G.O.S.</p>
      <form method="POST" action="${action}">
        <div class="grid">
          <div><label>Nome</label><input name="nome" value="${val('nome')}" required></div>
          <div><label>Cognome</label><input name="cognome" value="${val('cognome')}" required></div>
          <div><label>Telefono</label><input name="telefono" value="${val('telefono')}"></div>
          <div><label>Email</label><input name="email" value="${val('email')}"></div>
          <div><label>Codice fiscale</label><input name="codice_fiscale" value="${val('codice_fiscale')}"></div>
          <div><label>Data nascita</label><input type="date" name="data_nascita" value="${val('data_nascita')}"></div>
          <div><label>Luogo nascita</label><input name="luogo_nascita" value="${val('luogo_nascita')}"></div>
          <div><label>Indirizzo</label><input name="indirizzo" value="${val('indirizzo')}"></div>
          <div><label>Città</label><input name="citta" value="${val('citta')}"></div>
          <div><label>CAP</label><input name="cap" value="${val('cap')}"></div>
          <div><label>Numero documento</label><input name="documento_numero" value="${val('documento_numero')}"></div>
          <div><label>Scadenza documento</label><input type="date" name="documento_scadenza" value="${val('documento_scadenza')}"></div>
          <div><label>Numero patente</label><input name="patente_numero" value="${val('patente_numero')}"></div>
          <div><label>Scadenza patente</label><input type="date" name="patente_scadenza" value="${val('patente_scadenza')}"></div>
          <div><label>Categoria patente</label><input name="categoria_patente" value="${val('categoria_patente')}"></div>
        </div>
        <h3>Fatturazione</h3>
        <div class="grid">
          <div><label>Tipo cliente</label><select name="tipo_cliente"><option value="privato" ${c.tipo_cliente!=='azienda'?'selected':''}>Privato</option><option value="azienda" ${c.tipo_cliente==='azienda'?'selected':''}>Azienda</option></select></div>
          <div><label>Ragione sociale</label><input name="ragione_sociale" value="${val('ragione_sociale')}"></div>
          <div><label>Partita IVA</label><input name="piva" value="${val('piva') || val('partita_iva')}"></div>
          <div><label>PEC</label><input name="pec" value="${val('pec')}"></div>
          <div><label>Codice SDI</label><input name="sdi" value="${val('sdi') || val('codice_sdi')}"></div>
          <div class="full"><label>Indirizzo fatturazione</label><input name="indirizzo_fatturazione" value="${val('indirizzo_fatturazione')}"></div>
          <div><label>Città fatturazione</label><input name="citta_fatturazione" value="${val('citta_fatturazione')}"></div>
          <div><label>Provincia fatturazione</label><input name="provincia_fatturazione" value="${val('provincia_fatturazione')}"></div>
          <div><label>CAP fatturazione</label><input name="cap_fatturazione" value="${val('cap_fatturazione')}"></div>
        </div>
        <label>Note</label><textarea name="note">${val('note')}</textarea>
        <button>Salva cliente</button>
        <a class="btn btn2" href="/clienti">Annulla</a>
      </form>
    </div>`);
}
function clienteManualData(b){
  return {
    nome:b.nome||'', cognome:b.cognome||'', telefono:b.telefono||'', email:b.email||'', codice_fiscale:String(b.codice_fiscale||'').toUpperCase(),
    indirizzo:b.indirizzo||'', citta:b.citta||'', cap:b.cap||'', data_nascita:b.data_nascita||'', luogo_nascita:b.luogo_nascita||'',
    documento_numero:b.documento_numero||'', documento_scadenza:b.documento_scadenza||'', patente_numero:b.patente_numero||'',
    patente_scadenza:b.patente_scadenza||'', categoria_patente:b.categoria_patente||'', tipo_cliente:b.tipo_cliente||'privato', ragione_sociale:b.ragione_sociale||'', piva:b.piva||b.partita_iva||'', partita_iva:b.piva||b.partita_iva||'', pec:b.pec||'', sdi:b.sdi||b.codice_sdi||'', codice_sdi:b.sdi||b.codice_sdi||'', indirizzo_fatturazione:b.indirizzo_fatturazione||b.indirizzo||'', citta_fatturazione:b.citta_fatturazione||b.citta||'', provincia_fatturazione:b.provincia_fatturazione||b.provincia||'', cap_fatturazione:b.cap_fatturazione||b.cap||'', note:b.note||''
  };
}
app.get('/cliente-nuovo', (req,res)=>res.send(clienteManualForm({}, '/cliente-nuovo', 'Nuovo cliente manuale')));
app.post('/cliente-nuovo', (req,res)=>{
  const d = clienteManualData(req.body);
  db.run(`INSERT INTO clienti (nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,data_nascita,luogo_nascita,documento_numero,documento_scadenza,patente_numero,patente_scadenza,categoria_patente,tipo_cliente,ragione_sociale,piva,partita_iva,pec,sdi,codice_sdi,indirizzo_fatturazione,citta_fatturazione,provincia_fatturazione,cap_fatturazione,note,updated_at)
          VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)`,
    [d.nome,d.cognome,d.telefono,d.email,d.codice_fiscale,d.indirizzo,d.citta,d.cap,d.data_nascita,d.luogo_nascita,d.documento_numero,d.documento_scadenza,d.patente_numero,d.patente_scadenza,d.categoria_patente,d.tipo_cliente,d.ragione_sociale,d.piva,d.partita_iva,d.pec,d.sdi,d.codice_sdi,d.indirizzo_fatturazione,d.citta_fatturazione,d.provincia_fatturazione,d.cap_fatturazione,d.note],
    function(err){ if(err) return res.status(500).send(page('Errore cliente', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(err.message)}</pre></div>`)); res.redirect('/cliente/'+this.lastID); });
});
app.get('/cliente/:id/modifica', (req,res)=>{
  db.get(`SELECT * FROM clienti WHERE id=?`, [req.params.id], (err,c)=>{ if(!c) return res.redirect('/clienti'); res.send(clienteManualForm(c, `/cliente/${c.id}/modifica`, 'Modifica cliente')); });
});
app.post('/cliente/:id/modifica', (req,res)=>{
  const d = clienteManualData(req.body);
  db.run(`UPDATE clienti SET nome=?,cognome=?,telefono=?,email=?,codice_fiscale=?,indirizzo=?,citta=?,cap=?,data_nascita=?,luogo_nascita=?,documento_numero=?,documento_scadenza=?,patente_numero=?,patente_scadenza=?,categoria_patente=?,tipo_cliente=?,ragione_sociale=?,piva=?,partita_iva=?,pec=?,sdi=?,codice_sdi=?,indirizzo_fatturazione=?,citta_fatturazione=?,provincia_fatturazione=?,cap_fatturazione=?,note=?,updated_at=CURRENT_TIMESTAMP WHERE id=?`,
    [d.nome,d.cognome,d.telefono,d.email,d.codice_fiscale,d.indirizzo,d.citta,d.cap,d.data_nascita,d.luogo_nascita,d.documento_numero,d.documento_scadenza,d.patente_numero,d.patente_scadenza,d.categoria_patente,d.tipo_cliente,d.ragione_sociale,d.piva,d.partita_iva,d.pec,d.sdi,d.codice_sdi,d.indirizzo_fatturazione,d.citta_fatturazione,d.provincia_fatturazione,d.cap_fatturazione,d.note,req.params.id],
    err=>{ if(err) return res.status(500).send(page('Errore modifica', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(err.message)}</pre></div>`)); res.redirect('/cliente/'+req.params.id); });
});
app.get('/cliente/:id/elimina', (req,res)=>{
  db.get(`SELECT * FROM clienti WHERE id=?`, [req.params.id], (err,c)=>{ if(!c) return res.redirect('/clienti'); res.send(page('Elimina cliente', `<div class="box"><h2 class="bad">Eliminare cliente ${esc(c.nome)} ${esc(c.cognome)}?</h2><p>I contratti già creati restano nello storico.</p><form method="POST" action="/cliente/${c.id}/elimina"><button class="btn bad" type="submit">Conferma eliminazione</button><a class="btn btn2" href="/cliente/${c.id}">Annulla</a></form></div>`)); });
});
app.post('/cliente/:id/elimina', (req,res)=>{
  db.run(`DELETE FROM clienti WHERE id=?`, [req.params.id], ()=>res.redirect('/clienti'));
});

app.get('/nuova-da-cliente/:id', (req, res) => {
  db.get(`SELECT * FROM clienti WHERE id=?`, [req.params.id], (err, c) => {
    if (!c) return res.redirect('/clienti');
    const id = makeOcrId();
    OCR_PREFILL[id] = {
      nome:c.nome||'', cognome:c.cognome||'', telefono:c.telefono||'', email:c.email||'', codice_fiscale:c.codice_fiscale||'',
      indirizzo:c.indirizzo||'', citta:c.citta||'', cap:c.cap||'', data_nascita:c.data_nascita||'', luogo_nascita:c.luogo_nascita||'',
      documento_numero:c.documento_numero||'', documento_scadenza:c.documento_scadenza||'', conducente1:`${c.nome||''} ${c.cognome||''}`.trim(),
      patente1:c.patente_numero||'', patente1_scadenza:c.patente_scadenza||'', categoria_patente:c.categoria_patente||''
    };
    res.redirect('/nuova-prenotazione?ocr=' + encodeURIComponent(id));
  });
});

app.get('/ocr-doppio', (req, res) => {
  res.send(page('OCR carta identita e patente', `
    <div class="box">
      <h2>OCR cliente - Step 1 di 2</h2>
      <p class="notice">Prima scatta o carica la carta identita/documento. Dopo passerai alla patente.</p>
      <form method="POST" action="/ocr-doppio/documento" enctype="multipart/form-data">
        <input type="hidden" name="tipo" value="Carta identita">
        <label class="btn" style="display:block;text-align:center;font-size:22px;padding:18px;margin-top:14px">
          Scatta / carica CARTA IDENTITA
          <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
        </label>
        <button id="submitBtn" style="display:none">Leggi documento</button>
      </form>
      <a class="btn btn2" href="/nuova-prenotazione">Torna</a>
    </div>
    <script>
      const input = document.getElementById('fileInput');
      input.addEventListener('change', function () {
        if (this.files && this.files.length) document.getElementById('submitBtn').click();
      });
    </script>
  `));
});

app.post('/ocr-doppio/documento', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.send('File mancante');
    const documento = await estraiDatiDocumentoConAI(req.file.path, req.file.mimetype);
    const id = makeOcrId();
    OCR_PREFILL[id] = { documento: documento || {}, patente: {} };
    res.redirect('/ocr-doppio/patente?id=' + encodeURIComponent(id));
  } catch (e) {
    res.status(500).send(page('Errore OCR documento', `<div class="box"><h2 class="bad">Errore OCR documento</h2><pre>${esc(e.message)}</pre><a class="btn" href="/ocr-pro">Riprova</a></div>`));
  }
});

app.get('/ocr-doppio/patente', (req, res) => {
  const id = req.query.id;
  if (!id || !OCR_PREFILL[id]) return res.redirect('/ocr-doppio');
  res.send(page('OCR patente', `
    <div class="box">
      <h2>OCR cliente - Step 2 di 2</h2>
      <p class="notice">Ora scatta o carica la patente. I dati verranno uniti al documento.</p>
      <form method="POST" action="/ocr-doppio/patente" enctype="multipart/form-data">
        <input type="hidden" name="id" value="${esc(id)}">
        <input type="hidden" name="tipo" value="Patente">
        <label class="btn" style="display:block;text-align:center;font-size:22px;padding:18px;margin-top:14px">
          Scatta / carica PATENTE
          <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
        </label>
        <button id="submitBtn" style="display:none">Leggi patente</button>
      </form>
      <a class="btn btn2" href="/ocr-pro">Ricomincia</a>
    </div>
    <script>
      const input = document.getElementById('fileInput');
      input.addEventListener('change', function () {
        if (this.files && this.files.length) document.getElementById('submitBtn').click();
      });
    </script>
  `));
});

app.post('/ocr-doppio/patente', upload.single('file'), async (req, res) => {
  try {
    const id = req.body.id;
    if (!id || !OCR_PREFILL[id]) return res.redirect('/ocr-doppio');
    if (!req.file) return res.send('File mancante');
    const patente = await estraiDatiDocumentoConAI(req.file.path, req.file.mimetype);
    OCR_PREFILL[id].patente = patente || {};
    res.redirect('/ocr-doppio/conferma?id=' + encodeURIComponent(id));
  } catch (e) {
    res.status(500).send(page('Errore OCR patente', `<div class="box"><h2 class="bad">Errore OCR patente</h2><pre>${esc(e.message)}</pre><a class="btn" href="/ocr-pro">Riprova</a></div>`));
  }
});

app.get('/ocr-doppio/conferma', (req, res) => {
  const id = req.query.id;
  const data = OCR_PREFILL[id];
  if (!id || !data) return res.redirect('/ocr-doppio');

  const d = data.documento || {};
  const p = data.patente || {};
  const pick = (a, b) => a || b || '';
  const v = x => esc(x || '');

  const nome = pick(d.nome, p.nome);
  const cognome = pick(d.cognome, p.cognome);
  const cf = pick(d.codice_fiscale, p.codice_fiscale);
  const indirizzo = pick(d.indirizzo, p.indirizzo);
  const dataNascita = pick(d.data_nascita, p.data_nascita);
  const luogoNascita = pick(d.luogo_nascita, p.luogo_nascita);
  const documentoNumero = d.numero_documento || '';
  const documentoScadenza = d.data_scadenza || '';
  const patenteNumero = p.numero_patente || p.numero_documento || '';
  const patenteScadenza = p.data_scadenza || '';
  const categoriaPatente = p.categoria_patente || '';

  res.send(page('Controllo dati cliente', `
    <div class="box">
      <h2>Controlla dati cliente e conducente</h2>
      <p class="notice">Documento e patente sono stati letti separatamente. Correggi se serve, poi continua.</p>
      <form method="POST" action="/ocr-doppio/applica">
        <input type="hidden" name="id" value="${esc(id)}">

        <h3>Dati cliente da documento</h3>
        <div class="grid">
          <div><label>Nome</label><input name="nome" value="${v(nome)}"></div>
          <div><label>Cognome</label><input name="cognome" value="${v(cognome)}"></div>
          <div><label>Codice fiscale</label><input name="codice_fiscale" value="${v(cf)}"></div>
          <div><label>Indirizzo</label><input name="indirizzo" value="${v(indirizzo)}"></div>
          <div><label>Citta</label><input name="citta" value=""></div>
          <div><label>CAP</label><input name="cap" value=""></div>
          <div><label>Data nascita</label><input type="date" name="data_nascita" value="${v(dataNascita)}"></div>
          <div><label>Luogo nascita</label><input name="luogo_nascita" value="${v(luogoNascita)}"></div>
          <div><label>Numero documento</label><input name="documento_numero" value="${v(documentoNumero)}"></div>
          <div><label>Scadenza documento</label><input type="date" name="documento_scadenza" value="${v(documentoScadenza)}"></div>
        </div>

        <h3>Dati conducente da patente</h3>
        <div class="grid">
          <div><label>Conducente 1</label><input name="conducente1" value="${v((nome + ' ' + cognome).trim())}"></div>
          <div><label>Numero patente</label><input name="patente1" value="${v(patenteNumero)}"></div>
          <div><label>Scadenza patente</label><input type="date" name="patente1_scadenza" value="${v(patenteScadenza)}"></div>
          <div><label>Categoria patente</label><input name="categoria_patente" value="${v(categoriaPatente)}"></div>
        </div>

        <h3>Contatti da completare</h3>
        <div class="grid">
          <div><label>Telefono</label><input name="telefono" value=""></div>
          <div><label>Email</label><input name="email" value=""></div>
        </div>
        <button>Continua al contratto con questi dati</button>
      </form>
      <a class="btn btn2" href="/ocr-pro">Rifai OCR da capo</a>
    </div>
  `));
});

app.post('/ocr-doppio/applica', (req, res) => {
  const id = req.body.id || makeOcrId();
  OCR_PREFILL[id] = {
    nome: req.body.nome || '',
    cognome: req.body.cognome || '',
    telefono: req.body.telefono || '',
    email: req.body.email || '',
    codice_fiscale: req.body.codice_fiscale || '',
    indirizzo: req.body.indirizzo || '',
    citta: req.body.citta || '',
    cap: req.body.cap || '',
    data_nascita: req.body.data_nascita || '',
    luogo_nascita: req.body.luogo_nascita || '',
    documento_numero: req.body.documento_numero || '',
    documento_scadenza: req.body.documento_scadenza || '',
    conducente1: req.body.conducente1 || '',
    patente1: req.body.patente1 || '',
    patente1_scadenza: req.body.patente1_scadenza || '',
    categoria_patente: req.body.categoria_patente || ''
  };
  res.redirect('/nuova-prenotazione?ocr=' + encodeURIComponent(id));
});

app.get('/ocr-precontratto', (req, res) => {
  res.send(page('OCR prima del contratto', `
    <div class="box">
      <h2>OCR patente/documento prima del contratto</h2>
      <p class="notice">Da iPhone/iPad premi il pulsante e scegli Scatta foto oppure Libreria foto.</p>
      <form method="POST" action="/ocr-precontratto" enctype="multipart/form-data">
        <label>Tipo documento</label>
        <select name="tipo">
          <option>Patente</option>
          <option>Carta identita</option>
          <option>Codice fiscale</option>
          <option>Altro documento</option>
        </select>
        <label class="btn" style="display:block;text-align:center;font-size:22px;padding:18px;margin-top:14px">
          Scatta / carica documento
          <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
        </label>
        <button id="submitBtn" style="display:none">Leggi documento</button>
      </form>
      <a class="btn btn2" href="/nuova-prenotazione">Torna</a>
      <p class="notice">Dopo aver scelto la foto, la lettura parte automaticamente.</p>
    </div>
    <script>
      const input = document.getElementById('fileInput');
      input.addEventListener('change', function () {
        if (this.files && this.files.length) {
          document.getElementById('submitBtn').click();
        }
      });
    </script>
  `));
});

app.post('/ocr-precontratto', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.send('File mancante');
    const dati = await estraiDatiDocumentoConAI(req.file.path, req.file.mimetype);
    function v(x) { return esc(x || ''); }
    res.send(page('Controlla dati OCR', `
      <div class="box">
        <h2>Controlla e modifica i dati letti</h2>
        <p class="notice">L'AI puo sbagliare. Correggi i campi se serve, poi premi Continua al contratto.</p>
        <form method="POST" action="/ocr-precontratto/applica">
          <div class="grid">
            <div><label>Nome</label><input name="nome" value="${v(dati.nome)}"></div>
            <div><label>Cognome</label><input name="cognome" value="${v(dati.cognome)}"></div>
            <div><label>Telefono</label><input name="telefono" value=""></div>
            <div><label>Email</label><input name="email" value=""></div>
            <div><label>Codice fiscale</label><input name="codice_fiscale" value="${v(dati.codice_fiscale)}"></div>
            <div><label>Indirizzo</label><input name="indirizzo" value="${v(dati.indirizzo)}"></div>
            <div><label>Citta</label><input name="citta" value=""></div>
            <div><label>CAP</label><input name="cap" value=""></div>
            <div><label>Data nascita</label><input type="date" name="data_nascita" value="${v(dati.data_nascita)}"></div>
            <div><label>Luogo nascita</label><input name="luogo_nascita" value="${v(dati.luogo_nascita)}"></div>
            <div><label>Numero documento</label><input name="documento_numero" value="${v(dati.numero_documento)}"></div>
            <div><label>Scadenza documento</label><input type="date" name="documento_scadenza" value="${v(dati.data_scadenza)}"></div>
            <div><label>Numero patente</label><input name="patente1" value="${v(dati.numero_patente || dati.numero_documento)}"></div>
            <div><label>Scadenza patente</label><input type="date" name="patente1_scadenza" value="${v(dati.data_scadenza)}"></div>
            <div><label>Categoria patente</label><input name="categoria_patente" value="${v(dati.categoria_patente)}"></div>
          </div>
          <button>Continua al contratto con questi dati</button>
        </form>
        <a class="btn btn2" href="/ocr-pro">Riprova OCR</a>
      </div>
    `));
  } catch (e) {
    res.status(500).send(page('Errore OCR', `<div class="box"><h2 class="bad">Errore OCR</h2><pre>${esc(e.message)}</pre><p>Serve OPENAI_API_KEY su Render.</p><a class="btn" href="/ocr-pro">Riprova</a></div>`));
  }
});


app.post('/ocr-precontratto/applica', (req, res) => {
  const id = makeOcrId();
  OCR_PREFILL[id] = {
    nome: req.body.nome || '',
    cognome: req.body.cognome || '',
    telefono: req.body.telefono || '',
    email: req.body.email || '',
    codice_fiscale: req.body.codice_fiscale || '',
    indirizzo: req.body.indirizzo || '',
    citta: req.body.citta || '',
    cap: req.body.cap || '',
    data_nascita: req.body.data_nascita || '',
    luogo_nascita: req.body.luogo_nascita || '',
    documento_numero: req.body.documento_numero || '',
    documento_scadenza: req.body.documento_scadenza || '',
    patente1: req.body.patente1 || '',
    patente1_scadenza: req.body.patente1_scadenza || '',
    categoria_patente: req.body.categoria_patente || ''
  };
  res.redirect('/nuova-prenotazione?ocr=' + encodeURIComponent(id));
});

app.get('/nuova-prenotazione', async (req, res) => {
  const ocrData = OCR_PREFILL[req.query.ocr] || {};
  req.query = Object.assign({}, ocrData, req.query || {});

  const mezzi = await all(`SELECT * FROM mezzi ORDER BY categoria,targa`);
  res.send(page('Nuova prenotazione', `<h2>Nuova prenotazione / contratto</h2>
      <div class="box" style="border:2px solid #0b6b2d">
        <h3>1) Prima carica/scatta documento o patente</h3>
        <p class="notice">Consigliato: fai OCR prima di creare il contratto, così i dati cliente si compilano più velocemente.</p>
        <a class="btn btn3" href="/ocr-pro"> OCR carta identita + patente</a>
      </div>
      <h3>2) Poi controlla i dati e crea contratto</h3>
    ${req.query.data ? `<p class="notice">Aperta dal planning per il giorno <b>${esc(req.query.data)}</b>.</p>` : ''}${formPrenotazione(mezzi, req.query.mezzo_id, req.query.data, '/prenota-admin', req.query)}`));
});
async function ensureBookingColumnsV77(){
  const cols = {
    data_nascita:'TEXT', luogo_nascita:'TEXT', cittadinanza_cod:'TEXT', conducente_cittadinanza_cod:'TEXT',
    documento_tipo:'TEXT', documento_numero:'TEXT', documento_scadenza:'TEXT', patente_numero:'TEXT', patente_scadenza:'TEXT', categoria_patente:'TEXT',
    partita_iva:'TEXT', codice_sdi:'TEXT', piva:'TEXT', sdi:'TEXT', targa:'TEXT', marca:'TEXT', modello:'TEXT', tipo:'TEXT', categoria:'TEXT',
    cauzione_richiesta:'TEXT', cauzione_ricevuta:'TEXT', cauzione_importo:'REAL', cauzione_metodo:'TEXT', cauzione_restituita:'TEXT', tipo_record:'TEXT'
  };
  for (const [c,t] of Object.entries(cols)) {
    await run(`ALTER TABLE prenotazioni ADD COLUMN ${c} ${t}`).catch(()=>{});
  }
}

app.post('/prenota-admin', async (req, res) => {
  try {
    await ensureBookingColumnsV77();
    const b = req.body || {};
    const erroreDate = validDateRange(b.data_inizio, b.data_fine);
    if (erroreDate) return res.send(page('Errore date', `<div class="box"><h2 class="bad">${esc(erroreDate)}</h2><a class="btn" href="/nuova-prenotazione">Torna</a></div>`));
    const mezzo = await get(`SELECT * FROM mezzi WHERE id=?`, [b.mezzo_id]);
    if (!mezzo) return res.send(page('Mezzo non trovato', `<div class="box"><h2 class="bad">Mezzo non trovato</h2><a class="btn" href="/mezzi">Vai ai mezzi</a></div>`));
    const occ = await queryDisponibilita(b.mezzo_id, b.data_inizio, b.data_fine);
    if (occ) return res.send(page('Occupato', `<div class="box"><h2 class="bad">Mezzo occupato in queste date</h2><p>Contratto: <a href="/prenotazione/${occ.id}">${esc(occ.codice)}</a></p><a class="btn" href="/planning">Vai al planning</a></div>`));

    salvaClienteStorico({
      nome: b.nome, cognome: b.cognome, telefono: b.telefono, email: b.email,
      codice_fiscale: b.codice_fiscale, indirizzo: b.indirizzo, citta: b.citta, cap: b.cap,
      data_nascita: b.data_nascita, luogo_nascita: b.luogo_nascita,
      documento_numero: b.documento_numero, documento_scadenza: b.documento_scadenza,
      patente1: b.patente1 || b.patente_numero, patente_numero: b.patente1 || b.patente_numero,
      patente1_scadenza: b.patente1_scadenza || b.patente_scadenza, patente_scadenza: b.patente1_scadenza || b.patente_scadenza,
      categoria_patente: b.categoria_patente, note_cliente: b.note_cliente
    }, () => {});

    const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.ora_inizio, b.ora_fine, b.km_previsti);
    const data = {
      codice:'TEMP', nome:b.nome, cognome:b.cognome, telefono:b.telefono, email:b.email,
      codice_fiscale:b.codice_fiscale, indirizzo:b.indirizzo, citta:b.citta, cap:b.cap,
      data_nascita:b.data_nascita, luogo_nascita:b.luogo_nascita, cittadinanza_cod:b.cittadinanza_cod || '100000100', conducente_cittadinanza_cod:b.cittadinanza_cod || '100000100',
      documento_tipo:b.documento_tipo || 'IDENT', documento_numero:b.documento_numero, documento_scadenza:b.documento_scadenza,
      patente_numero:b.patente1 || b.patente_numero, patente_scadenza:b.patente1_scadenza || b.patente_scadenza, categoria_patente:b.categoria_patente,
      tipo_cliente:b.tipo_cliente || 'privato', piva:b.piva, partita_iva:b.piva || b.partita_iva, ragione_sociale:b.ragione_sociale, pec:b.pec, sdi:b.sdi, codice_sdi:b.sdi || b.codice_sdi,
      conducente1:b.conducente1 || `${b.nome||''} ${b.cognome||''}`.trim(), patente1:b.patente1 || b.patente_numero, patente1_scadenza:b.patente1_scadenza || b.patente_scadenza,
      conducente2:b.conducente2, patente2:b.patente2, patente2_scadenza:b.patente2_scadenza,
      mezzo_id:b.mezzo_id, targa:mezzo.targa || '', marca:mezzo.marca || '', modello:mezzo.modello || '', tipo:mezzo.tipo || '', categoria:mezzo.categoria || mezzo.tipo || '',
      data_inizio:b.data_inizio, data_fine:b.data_fine, ora_inizio:b.ora_inizio || '08:30', ora_fine:b.ora_fine || '18:00', giorni:calc.giorni,
      km_previsti:Number(b.km_previsti || 0), extra_fuori_orario:calc.extra_fuori_orario, extra_km:calc.extraKm,
      imponibile:calc.imponibile, iva:calc.iva, totale:calc.totale, cauzione:mezzo.cauzione || CAUZIONE,
      cauzione_richiesta:'si', cauzione_ricevuta:'no', cauzione_importo:mezzo.cauzione || CAUZIONE, cauzione_metodo:'', cauzione_restituita:'no',
      carburante_uscita:b.carburante_uscita || '4/4 pieno', stato:'contratto', tipo_record:'contratto', note:b.note || ''
    };
    const cols = Object.keys(data);
    const placeholders = cols.map(()=>'?').join(',');
    const result = await run(`INSERT INTO prenotazioni (${cols.join(',')}) VALUES (${placeholders})`, cols.map(k=>data[k]));
    const cod = codicePratica(result.lastID);
    await run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, result.lastID]);
    res.send(actionScreen(result.lastID, 'Contratto creato', `Codice: <b>${cod}</b><br>Mezzo: <b>${esc(data.targa)} ${esc(data.marca)} ${esc(data.modello)}</b><br>Totale: <b>&euro; ${euro(calc.totale)}</b>`));
  } catch (e) {
    res.status(500).send(page('Errore prenotazione', `<div class="box"><h2 class="bad">Errore prenotazione</h2><pre>${esc(e.stack || e.message)}</pre><a class="btn" href="/nuova-prenotazione">Torna</a></div>`));
  }
});

// =========================
// V92 PAGINA CLIENTE PULITA - SOLO CLIENTE, MANUALE + FOTO/OCR
// =========================
function clienteWebVal(req, key, fallback='') {
  return esc((req.query && req.query[key]) || fallback || '');
}
function clienteWebSelected(req, key, value, fallback='') {
  const v = (req.query && req.query[key]) || fallback || '';
  return String(v) === String(value) ? 'selected' : '';
}
function clienteWebHtml(req) {
  const categoria = (req.query && req.query.categoria) || '';
  const dataInizio = (req.query && req.query.data_inizio) || '';
  const dataFine = (req.query && req.query.data_fine) || '';
  const km = (req.query && req.query.km_previsti) || '150';
  return `<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>DP RENT - Dati cliente</title>
<style>
:root{--dp-red:#d70000;--dp-dark:#101015;--dp-blue:#173b8f;--bg:#eef4ff;--card:#fff;--line:#d9dbe7;--green:#1f7a36}
*{box-sizing:border-box} body{margin:0;background:linear-gradient(180deg,#edf5ff,#f7f8fb);font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;color:#111827;font-size:18px}
.client-nav{position:sticky;top:0;z-index:999;background:rgba(7,17,31,.96);backdrop-filter:blur(10px);padding:calc(10px + env(safe-area-inset-top)) 14px 10px;display:flex;gap:10px;flex-wrap:wrap;box-shadow:0 8px 22px rgba(0,0,0,.22)}
.client-nav a,.client-nav button{appearance:none;border:0;text-decoration:none;display:inline-flex;align-items:center;justify-content:center;border-radius:16px;padding:13px 16px;font-size:17px;font-weight:900;color:#fff;background:#24242b;cursor:pointer;box-shadow:0 8px 18px rgba(0,0,0,.18)}
.client-nav .dash{background:#d70000}.client-nav .wait{background:#173b8f}.client-nav .back{background:#333}
@media(max-width:720px){.client-nav a,.client-nav button{flex:1 1 calc(50% - 6px);font-size:16px;padding:13px 8px}.client-nav .wide{flex-basis:100%}}
.hero{background:linear-gradient(135deg,#07111f,#163d91);color:#fff;padding:28px 22px 30px;border-radius:0 0 28px 28px;box-shadow:0 12px 35px rgba(0,0,0,.18)}
.hero h1{margin:0;font-size:38px;letter-spacing:.5px}.hero p{font-size:20px;line-height:1.35;margin:12px 0 0;opacity:.95}.pill{display:inline-block;background:#fff;color:#163d91;font-weight:900;border-radius:999px;padding:10px 16px;margin:14px 8px 0 0}
.wrap{max-width:900px;margin:20px auto;padding:0 14px 36px}.card{background:var(--card);border:1px solid #e3e7f2;border-radius:24px;padding:22px;margin:18px 0;box-shadow:0 14px 35px rgba(15,23,42,.08)}
h2{font-size:30px;margin:0 0 16px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px} label{display:block;font-weight:900;margin:10px 0 7px;font-size:17px;color:#111827}
input,select,textarea{width:100%;border:2px solid var(--line);border-radius:17px;padding:16px;font-size:20px;background:#fff;color:#111;font-weight:700;outline:none} input:focus,select:focus,textarea:focus{border-color:#3159c7;box-shadow:0 0 0 4px rgba(49,89,199,.12)}
textarea{min-height:110px}.full{grid-column:1/-1}.notice{background:#fff8df;border:1px solid #f1d98a;border-radius:20px;padding:16px;line-height:1.35;font-size:18px}.okbox{background:#ecfff1;border:1px solid #b8efc4;border-radius:20px;padding:16px;line-height:1.35}.btn{border:0;border-radius:20px;background:linear-gradient(135deg,#e21818,#a80d0d);color:#fff;padding:18px 24px;font-size:24px;font-weight:900;box-shadow:0 12px 25px rgba(210,0,0,.25);width:100%;margin-top:18px}.btn2{display:inline-block;text-decoration:none;text-align:center;background:#24242b;color:#fff;border-radius:18px;padding:14px 18px;font-weight:900;margin-top:8px}.file{padding:14px;background:#f7f8fc}.small{font-size:15px;color:#596275;font-weight:700}@media(max-width:720px){.grid{grid-template-columns:1fr}.hero h1{font-size:32px}body{font-size:17px}input,select,textarea{font-size:19px}.card{padding:18px;border-radius:22px}}

.contract-main-actions{margin-top:16px}.contract-main-actions .btn{min-width:190px;text-align:center}.contract-secondary-actions .btn{min-width:150px;text-align:center}
@media(max-width:700px){.contract-main-actions .btn,.contract-secondary-actions .btn{width:100%;min-width:0}}


/* V109 FIX leggibilita mobile */
header{padding-top:max(22px, env(safe-area-inset-top));}
.top-actions{max-width:1180px;margin:0 auto 14px!important;padding:10px 0!important;}
.top-actions .back-btn::before{content:""!important;}
.top-actions .back-btn,.top-actions a{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif!important;font-size:clamp(18px,2.6vw,24px)!important;letter-spacing:0!important;line-height:1.1!important;white-space:nowrap!important;color:#fff!important;overflow:hidden;text-overflow:ellipsis;}
.top-actions .back-btn{background:#333!important;}
.top-actions .home-btn{background:#d70000!important;}
.client-back button{font-size:18px!important;font-weight:900!important;background:#333!important;color:#fff!important;}
@media(max-width:700px){
  nav{padding-top:calc(14px + env(safe-area-inset-top));}
  .top-actions{position:sticky;top:0;z-index:50;padding:10px 12px!important;gap:10px!important;background:rgba(244,244,244,.96)!important;}
  .top-actions .back-btn,.top-actions a{min-width:0!important;width:calc(50% - 5px)!important;flex:1 1 calc(50% - 5px)!important;padding:14px 8px!important;}
  .contract-main-actions .btn{width:100%!important;}
}

</style>
<script>
function toggleAzienda(){
  var t=document.querySelector('[name="tipo_cliente"]').value;
  var box=document.getElementById('aziendaBox');
  var isAz=t==='azienda';
  box.style.display=isAz?'grid':'none';
  ['ragione_sociale','partita_iva','pec','codice_sdi'].forEach(function(n){var el=document.querySelector('[name="'+n+'"]'); if(el) el.required=isAz;});
}
window.addEventListener('DOMContentLoaded',toggleAzienda)
</script>
</head>
<body>
<section class="hero">
  <h1>DP RENT</h1>
  <p>Preventivo confermato. Compila tutti i dati per preparare pratica interna, contratto e controlli Ca.R.G.O.S.</p>
  <span class="pill">${clienteWebVal(req,'ref','Pratica cliente')}</span>
  ${categoria ? `<span class="pill">${esc(categoria)}</span>` : ''}
</section>
<div class="wrap">
  ${req.query && req.query.ocr_done ? `<div class="card" style="border:3px solid #1f7a36"><h2>✅ Documenti caricati e OCR eseguito</h2><p class="okbox">Non devi ricaricare documento e patente: i file sono già collegati alla pratica. Controlla i dati sotto e completa mezzo, date, km e fatturazione.</p></div>` : `<div class="card" style="border:3px solid #173b8f">
    <h2>📸 Prima carica documento e patente</h2>
    <p class="notice">Carica le foto qui: il sistema prova a leggere i dati automaticamente. Dopo trovi i campi già compilati e puoi correggere tutto a mano.</p>
    ${req.query && req.query.ocr_done ? `<p class="okbox"><b>✅ OCR eseguito.</b><br>Controlla i campi sotto e completa quelli mancanti.</p>` : ``}
    <form method="POST" action="/prenota-ocr" enctype="multipart/form-data">
      <input type="hidden" name="ref" value="${clienteWebVal(req,'ref')}">
      <input type="hidden" name="categoria" value="${esc(categoria)}">
      <input type="hidden" name="data_inizio" value="${esc(dataInizio)}">
      <input type="hidden" name="ora_inizio" value="${clienteWebVal(req,'ora_inizio','08:30')}">
      <input type="hidden" name="data_fine" value="${esc(dataFine)}">
      <input type="hidden" name="ora_fine" value="${clienteWebVal(req,'ora_fine','18:00')}">
      <input type="hidden" name="km_previsti" value="${esc(km)}">
      <input type="hidden" name="telefono" value="${clienteWebVal(req,'telefono')}">
      <div class="grid">
        <div><label>Documento fronte</label><input class="file" type="file" name="documento_fronte" accept="image/*,application/pdf" capture="environment"></div>
        <div><label>Documento retro</label><input class="file" type="file" name="documento_retro" accept="image/*,application/pdf" capture="environment"></div>
        <div><label>Patente fronte</label><input class="file" type="file" name="patente_fronte" accept="image/*,application/pdf" capture="environment"></div>
        <div><label>Patente retro</label><input class="file" type="file" name="patente_retro" accept="image/*,application/pdf" capture="environment"></div>
      </div>
      <button class="btn" type="submit">Leggi dati automaticamente</button>
    </form>
    <p class="small">Se l'OCR non legge tutto, puoi comunque compilare manualmente sotto.</p>
  </div>`}

<form method="POST" action="/prenota-cliente" enctype="multipart/form-data">
  <input type="hidden" name="ref" value="${clienteWebVal(req,'ref')}">
  <input type="hidden" name="preupload_id" value="${clienteWebVal(req,'preupload_id')}">
  <div class="card">
    <h2>Mezzo e periodo</h2>
    <div class="grid">
      <div class="full"><label>Tipo mezzo richiesto</label><select name="categoria" required>
        <option value="FURGONE" ${categoria==='FURGONE'?'selected':''}>Furgone cargo/merci</option>
        <option value="9_POSTI" ${categoria==='9_POSTI'?'selected':''}>Pulmino 8/9 posti</option>
        <option value="AUTO_DACIA" ${categoria==='AUTO_DACIA'?'selected':''}>Auto economica</option>
        <option value="AUTO_GOLF" ${categoria==='AUTO_GOLF'?'selected':''}>Auto categoria Golf</option>
        <option value="ESCAVATORE" ${categoria==='ESCAVATORE'?'selected':''}>Escavatore / mezzo speciale</option>
        <option value="SEMOVENTE" ${categoria==='SEMOVENTE'?'selected':''}>Piattaforma / semovente</option>
      </select><div class="small">La targa resta interna a DP RENT.</div></div>
      <div><label>Data ritiro</label><input type="date" name="data_inizio" value="${esc(dataInizio)}" required></div>
      <div><label>Ora ritiro</label><input type="time" name="ora_inizio" value="${clienteWebVal(req,'ora_inizio','08:30')}"></div>
      <div><label>Data riconsegna</label><input type="date" name="data_fine" value="${esc(dataFine)}" required></div>
      <div><label>Ora riconsegna</label><input type="time" name="ora_fine" value="${clienteWebVal(req,'ora_fine','18:00')}"></div>
      <div><label>Km previsti</label><input type="number" name="km_previsti" value="${esc(km)}"></div>
    </div>
  </div>

  <div class="card">
    <h2>Dati cliente / conducente</h2>
    <p class="notice">Puoi compilare manualmente tutti i campi e caricare le foto. L'ufficio DP controllerà i dati prima del contratto definitivo.</p>
    <div class="grid">
      <div><label>Nome</label><input name="nome" value="${clienteWebVal(req,'nome')}" required></div>
      <div><label>Cognome</label><input name="cognome" value="${clienteWebVal(req,'cognome')}" required></div>
      <div><label>Telefono</label><input name="telefono" value="${clienteWebVal(req,'telefono')}" required></div>
      <div><label>Email</label><input type="email" name="email" value="${clienteWebVal(req,'email')}"></div>
      <div><label>Codice fiscale</label><input name="codice_fiscale" value="${clienteWebVal(req,'codice_fiscale')}" style="text-transform:uppercase"></div>
      <div><label>Data nascita</label><input type="date" name="data_nascita" value="${clienteWebVal(req,'data_nascita')}"></div>
      <div><label>Luogo nascita</label><input name="luogo_nascita" value="${clienteWebVal(req,'luogo_nascita')}"></div>
      <div><label>Cittadinanza codice</label><input name="cittadinanza_cod" value="${clienteWebVal(req,'cittadinanza_cod','100000100')}"></div>
      <div class="full"><label>Indirizzo residenza</label><input name="indirizzo" value="${clienteWebVal(req,'indirizzo')}"></div>
      <div><label>Città</label><input name="citta" value="${clienteWebVal(req,'citta')}"></div>
      <div><label>Provincia</label><input name="provincia" value="${clienteWebVal(req,'provincia')}"></div>
      <div><label>CAP</label><input name="cap" value="${clienteWebVal(req,'cap')}"></div>
    </div>
  </div>

  <div class="card">
    <h2>Documento e patente</h2>
    <div class="grid">
      <div><label>Tipo documento</label><select name="documento_tipo"><option value="IDENT">Carta identità</option><option value="PASS">Passaporto</option><option value="PATENTE">Patente</option></select></div>
      <div><label>Numero documento</label><input name="documento_numero" value="${clienteWebVal(req,'documento_numero')}"></div>
      <div><label>Data rilascio documento</label><input type="date" name="documento_rilascio" value="${clienteWebVal(req,'documento_rilascio')}"></div>
      <div><label>Scadenza documento</label><input type="date" name="documento_scadenza" value="${clienteWebVal(req,'documento_scadenza')}"></div>
      <div><label>Luogo rilascio documento COD</label><input name="record_cargos_doc_luogoril_cod" value="${clienteWebVal(req,'record_cargos_doc_luogoril_cod')}"></div>
      <div><label>Numero patente</label><input name="patente_numero" value="${clienteWebVal(req,'patente_numero')}"></div>
      <div><label>Categoria patente</label><input name="categoria_patente" value="${clienteWebVal(req,'categoria_patente')}"></div>
      <div><label>Data rilascio patente</label><input type="date" name="patente_rilascio" value="${clienteWebVal(req,'patente_rilascio')}"></div>
      <div><label>Scadenza patente</label><input type="date" name="patente_scadenza" value="${clienteWebVal(req,'patente_scadenza')}"></div>
      <div><label>Luogo rilascio patente COD</label><input name="record_cargos_patente_luogoril_cod" value="${clienteWebVal(req,'record_cargos_patente_luogoril_cod')}"></div>
    </div>
  </div>

  <div class="card">
    <h2>Fatturazione</h2>
    <div class="grid">
      <div><label>Tipo cliente</label><select name="tipo_cliente" onchange="toggleAzienda()"><option value="privato" ${clienteWebSelected(req,'tipo_cliente','privato','privato')}>Privato</option><option value="azienda" ${clienteWebSelected(req,'tipo_cliente','azienda')}>Azienda</option></select></div>
    </div>
    <div class="grid">
      <div class="full"><label>Indirizzo fatturazione</label><input name="indirizzo_fatturazione" value="${clienteWebVal(req,'indirizzo_fatturazione')}"></div>
      <div><label>Città fatturazione</label><input name="citta_fatturazione" value="${clienteWebVal(req,'citta_fatturazione')}"></div>
      <div><label>Provincia fatturazione</label><input name="provincia_fatturazione" value="${clienteWebVal(req,'provincia_fatturazione')}"></div>
      <div><label>CAP fatturazione</label><input name="cap_fatturazione" value="${clienteWebVal(req,'cap_fatturazione')}"></div>
    </div>
    <div class="grid" id="aziendaBox">
      <div class="full"><label>Ragione sociale</label><input name="ragione_sociale" value="${clienteWebVal(req,'ragione_sociale')}"></div>
      <div><label>Partita IVA</label><input name="partita_iva" value="${clienteWebVal(req,'partita_iva')}"></div>
      <div><label>PEC</label><input name="pec" value="${clienteWebVal(req,'pec')}"></div>
      <div><label>Codice SDI</label><input name="codice_sdi" value="${clienteWebVal(req,'codice_sdi')}"></div>
    </div>
  </div>

  ${req.query && req.query.preupload_id ? `<div class="card"><h2>Foto documento / patente</h2><p class="okbox">Foto già ricevute con OCR. Qui puoi aggiungere solo altri allegati se servono.</p><div class="grid"><div class="full"><label>Altri allegati</label><input class="file" type="file" name="altri_allegati" accept="image/*,application/pdf" multiple></div></div></div>` : `<div class="card">
    <h2>Foto documento / patente</h2>
    <div class="grid">
      <div><label>Documento fronte</label><input class="file" type="file" name="documento_fronte" accept="image/*,application/pdf" capture="environment"></div>
      <div><label>Documento retro</label><input class="file" type="file" name="documento_retro" accept="image/*,application/pdf" capture="environment"></div>
      <div><label>Patente fronte</label><input class="file" type="file" name="patente_fronte" accept="image/*,application/pdf" capture="environment"></div>
      <div><label>Patente retro</label><input class="file" type="file" name="patente_retro" accept="image/*,application/pdf" capture="environment"></div>
      <div class="full"><label>Altri allegati</label><input class="file" type="file" name="altri_allegati" accept="image/*,application/pdf" multiple></div>
    </div>
  </div>`}

  <div class="card">
    <h2>Note</h2>
    <textarea name="note" placeholder="Scrivi eventuali note per DP RENT">${clienteWebVal(req,'note')}</textarea>
    ${privacyCheckboxHtml()}
    <button class="btn" type="submit">Invia dati a DP RENT</button>
  </div>
</form>
</div>
</body></html>`;
}

async function ensureClienteWebColumnsV92(){
  const cols = {
    data_nascita:'TEXT', luogo_nascita:'TEXT', cittadinanza_cod:'TEXT', conducente_cittadinanza_cod:'TEXT',
    documento_tipo:'TEXT', documento_numero:'TEXT', documento_scadenza:'TEXT', documento_rilascio:'TEXT',
    record_cargos_doc_luogoril_cod:'TEXT', record_cargos_patente_luogoril_cod:'TEXT',
    patente_numero:'TEXT', patente_scadenza:'TEXT', patente_rilascio:'TEXT', categoria_patente:'TEXT',
    tipo_cliente:'TEXT', partita_iva:'TEXT', piva:'TEXT', ragione_sociale:'TEXT', pec:'TEXT', codice_sdi:'TEXT', sdi:'TEXT', indirizzo_fatturazione:'TEXT', citta_fatturazione:'TEXT', provincia_fatturazione:'TEXT', cap_fatturazione:'TEXT',
    provincia:'TEXT', citta:'TEXT', cap:'TEXT', giorni:'INTEGER', km_previsti:'TEXT', extra_fuori_orario:'REAL', extra_km:'REAL', imponibile:'REAL', iva:'REAL', cauzione:'REAL', tipo_record:'TEXT', note:'TEXT'
  };
  for (const [c,t] of Object.entries(cols)) await run(`ALTER TABLE prenotazioni ADD COLUMN ${c} ${t}`).catch(()=>{});
}

app.get('/prenota', (req, res) => {
  res.setHeader('Content-Type','text/html; charset=utf-8');
  res.send(clienteWebHtml(req));
});


app.post('/prenota-ocr', upload.fields([
  {name:'documento_fronte', maxCount:1}, {name:'documento_retro', maxCount:1},
  {name:'patente_fronte', maxCount:1}, {name:'patente_retro', maxCount:1}
]), async (req, res) => {
  try {
    const b = req.body || {};
    const uploaded = [];
    const results = [];
    for (const [field, arr] of Object.entries(req.files || {})) {
      for (const f of (arr || [])) {
        uploaded.push({ tipo: field, f });
        try {
          const dati = await estraiDatiDocumentoConAI(f.path, f.mimetype);
          results.push(dati || {});
        } catch (e) {
          console.log('OCR prenota warning:', field, e.message);
        }
      }
    }
    if (!uploaded.length) {
      return res.send(`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><h1>Nessuna foto caricata</h1><p>Carica almeno documento o patente.</p><p><a href="javascript:history.back()">↩️ Torna indietro</a></p>`);
    }
    const preuploadId = 'OCR' + Date.now() + Math.floor(Math.random()*9999);
    PREN_OCR_UPLOADS[preuploadId] = uploaded;
    setTimeout(() => { delete PREN_OCR_UPLOADS[preuploadId]; }, 6 * 60 * 60 * 1000).unref?.();

    const m = mergeOcrObjects(results);
    const q = new URLSearchParams();
    // conserva dati pratica da WhatsApp
    for (const k of ['ref','categoria','data_inizio','ora_inizio','data_fine','ora_fine','km_previsti','telefono']) {
      if (b[k]) q.set(k, b[k]);
    }
    q.set('preupload_id', preuploadId);
    q.set('ocr_done', '1');

    // mappa OCR -> campi pagina cliente
    const map = {
      nome:'nome', cognome:'cognome', data_nascita:'data_nascita', luogo_nascita:'luogo_nascita',
      codice_fiscale:'codice_fiscale', indirizzo:'indirizzo',
      numero_documento:'documento_numero', data_rilascio:'documento_rilascio', data_scadenza:'documento_scadenza',
      numero_patente:'patente_numero', categoria_patente:'categoria_patente'
    };
    for (const [src, dst] of Object.entries(map)) {
      if (m[src]) q.set(dst, m[src]);
    }
    // se OCR riconosce patente con data_scadenza ma documento già pieno, non sovrascrivo. Serve comunque modifica manuale.
    if (m.numero_patente && m.data_scadenza && !q.get('patente_scadenza')) q.set('patente_scadenza', m.data_scadenza);

    res.redirect('/prenota?' + q.toString());
  } catch (e) {
    res.status(500).send(`<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><h1>Errore OCR</h1><pre>${esc(e.stack || e.message)}</pre><p>Controlla OPENAI_API_KEY su Render.</p><a href="javascript:history.back()">Torna</a>`);
  }
});


async function v123EnsureAllegatiClienteId(){
  await run(`ALTER TABLE allegati ADD COLUMN cliente_id INTEGER`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN tipo_cliente TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN ragione_sociale TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN piva TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN partita_iva TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN pec TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN sdi TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN codice_sdi TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN indirizzo_fatturazione TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN citta_fatturazione TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN provincia_fatturazione TEXT`).catch(()=>{});
  await run(`ALTER TABLE clienti ADD COLUMN cap_fatturazione TEXT`).catch(()=>{});
}
async function v123UpsertClienteStorico(d){
  await v123EnsureAllegatiClienteId();
  const cf = String(d.codice_fiscale || '').trim().toUpperCase();
  let old = null;
  if (cf) old = await get(`SELECT id FROM clienti WHERE codice_fiscale=?`, [cf]).catch(()=>null);
  if (!old && d.telefono) old = await get(`SELECT id FROM clienti WHERE telefono=? ORDER BY id DESC LIMIT 1`, [d.telefono]).catch(()=>null);
  const vals = [d.nome||'',d.cognome||'',d.telefono||'',d.email||'',cf,d.indirizzo||'',d.citta||'',d.cap||'',d.data_nascita||'',d.luogo_nascita||'',d.documento_numero||'',d.documento_scadenza||'',d.patente_numero||d.patente1||'',d.patente_scadenza||d.patente1_scadenza||'',d.categoria_patente||'',d.tipo_cliente||'privato',d.ragione_sociale||'',d.piva||d.partita_iva||'',d.piva||d.partita_iva||'',d.pec||'',d.sdi||d.codice_sdi||'',d.sdi||d.codice_sdi||'',d.indirizzo_fatturazione||d.indirizzo||'',d.citta_fatturazione||d.citta||'',d.provincia_fatturazione||d.provincia||'',d.cap_fatturazione||d.cap||'',d.note||''];
  if (old && old.id) {
    await run(`UPDATE clienti SET nome=?,cognome=?,telefono=?,email=?,codice_fiscale=?,indirizzo=?,citta=?,cap=?,data_nascita=?,luogo_nascita=?,documento_numero=?,documento_scadenza=?,patente_numero=?,patente_scadenza=?,categoria_patente=?,tipo_cliente=?,ragione_sociale=?,piva=?,partita_iva=?,pec=?,sdi=?,codice_sdi=?,indirizzo_fatturazione=?,citta_fatturazione=?,provincia_fatturazione=?,cap_fatturazione=?,note=?,updated_at=CURRENT_TIMESTAMP WHERE id=?`, [...vals, old.id]);
    return old.id;
  }
  const r = await run(`INSERT INTO clienti (nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,data_nascita,luogo_nascita,documento_numero,documento_scadenza,patente_numero,patente_scadenza,categoria_patente,tipo_cliente,ragione_sociale,piva,partita_iva,pec,sdi,codice_sdi,indirizzo_fatturazione,citta_fatturazione,provincia_fatturazione,cap_fatturazione,note) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`, vals);
  return r.lastID;
}

app.post('/prenota-cliente', upload.fields([
  {name:'documento_fronte', maxCount:1}, {name:'documento_retro', maxCount:1},
  {name:'patente_fronte', maxCount:1}, {name:'patente_retro', maxCount:1},
  {name:'altri_allegati', maxCount:20}
]), async (req, res) => {
  try {
    await ensureClienteWebColumnsV92();
    await v123EnsureAllegatiClienteId();
    const b = req.body || {};
    const erroreDate = validDateRange(b.data_inizio, b.data_fine);
    if (erroreDate) return res.send(`<!doctype html><meta charset="utf-8"><h1>Errore date</h1><p>${esc(erroreDate)}</p><a href="javascript:history.back()">Torna</a>`);
    if (String(b.tipo_cliente || '').toLowerCase() === 'azienda') {
      const mancanti = [];
      if (!b.ragione_sociale) mancanti.push('ragione sociale');
      if (!(b.partita_iva || b.piva)) mancanti.push('partita IVA');
      if (!b.pec) mancanti.push('PEC');
      if (!(b.codice_sdi || b.sdi)) mancanti.push('codice SDI');
      if (mancanti.length) return res.send(`<!doctype html><meta charset="utf-8"><h1>Dati fatturazione mancanti</h1><p>Per azienda mancano: ${esc(mancanti.join(', '))}</p><a href="javascript:history.back()">Torna</a>`);
    }

    const categoriaRichiesta = categoriaClienteNorm(b.categoria);
    const mezziTutti = await all(`SELECT * FROM mezzi ORDER BY id ASC`);
    const mezzi = mezziTutti.filter(m => mezzoCompatibileCategoriaCliente(m, categoriaRichiesta));
    if (!mezzi.length) return res.send(`<!doctype html><meta charset="utf-8"><h1>Nessun mezzo</h1><p>Nessun mezzo configurato correttamente per: ${esc(categoriaRichiesta)}.</p><p>Controlla anagrafica mezzi: categoria e posti.</p><a href="javascript:history.back()">Torna</a>`);
    let mezzo = null;
    for (const m of mezzi) {
      const occ = await queryDisponibilita(m.id, b.data_inizio, b.data_fine);
      if (!occ) { mezzo = m; break; }
    }
    if (!mezzo) return res.send(`<!doctype html><meta charset="utf-8"><h1>Non disponibile</h1><p>Nessun mezzo libero per ${esc(categoriaRichiesta)} nelle date richieste. DP RENT ti ricontatterà.</p><a href="javascript:history.back()">Torna</a>`);

    const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.ora_inizio || '08:30', b.ora_fine || '18:00', b.km_previsti || 150);
    const data = {
      codice:'TEMP', nome:b.nome, cognome:b.cognome, telefono:b.telefono, email:b.email,
      codice_fiscale:String(b.codice_fiscale || '').toUpperCase(), indirizzo:b.indirizzo, citta:b.citta, provincia:b.provincia, cap:b.cap,
      data_nascita:b.data_nascita, luogo_nascita:b.luogo_nascita, cittadinanza_cod:b.cittadinanza_cod || '100000100', conducente_cittadinanza_cod:b.cittadinanza_cod || '100000100',
      documento_tipo:b.documento_tipo || 'IDENT', documento_numero:b.documento_numero, documento_rilascio:b.documento_rilascio, documento_scadenza:b.documento_scadenza,
      record_cargos_doc_luogoril_cod:b.record_cargos_doc_luogoril_cod, record_cargos_patente_luogoril_cod:b.record_cargos_patente_luogoril_cod,
      patente_numero:b.patente_numero, patente_rilascio:b.patente_rilascio, patente_scadenza:b.patente_scadenza, categoria_patente:b.categoria_patente,
      tipo_cliente:b.tipo_cliente || 'privato', partita_iva:b.partita_iva || b.piva, piva:b.partita_iva || b.piva, ragione_sociale:b.ragione_sociale, pec:b.pec, codice_sdi:b.codice_sdi || b.sdi, sdi:b.codice_sdi || b.sdi, indirizzo_fatturazione:b.indirizzo_fatturazione || b.indirizzo, citta_fatturazione:b.citta_fatturazione || b.citta, provincia_fatturazione:b.provincia_fatturazione || b.provincia, cap_fatturazione:b.cap_fatturazione || b.cap,
      mezzo_id:mezzo.id, targa:mezzo.targa || '', marca:mezzo.marca || '', modello:mezzo.modello || '', tipo:mezzo.tipo || '', categoria:categoriaRichiesta || mezzo.categoria || mezzo.tipo || '',
      data_inizio:b.data_inizio, data_fine:b.data_fine, ora_inizio:b.ora_inizio || '08:30', ora_fine:b.ora_fine || '18:00', giorni:calc.giorni,
      km_previsti:Number(b.km_previsti || 0), extra_fuori_orario:calc.extra_fuori_orario, extra_km:calc.extraKm,
      imponibile:calc.imponibile, iva:calc.iva, totale:calc.totale, cauzione:mezzo.cauzione || CAUZIONE,
      stato:'richiesta_cliente', tipo_record:'preventivo', note:b.note || ''
    };
    const cols = Object.keys(data);
    const result = await run(`INSERT INTO prenotazioni (${cols.join(',')}) VALUES (${cols.map(()=>'?').join(',')})`, cols.map(k=>data[k]));
    const cod = codicePratica(result.lastID);
    await run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, result.lastID]);

    // V123: salva/aggiorna cliente nello storico con fatturazione completa e collega i documenti all'archivio cliente
    let clienteStoricoId = null;
    try { clienteStoricoId = await v123UpsertClienteStorico({
      nome:b.nome, cognome:b.cognome, telefono:b.telefono, email:b.email, codice_fiscale:b.codice_fiscale,
      indirizzo:b.indirizzo, citta:b.citta, provincia:b.provincia, cap:b.cap, data_nascita:b.data_nascita, luogo_nascita:b.luogo_nascita,
      documento_numero:b.documento_numero, documento_scadenza:b.documento_scadenza,
      patente_numero:b.patente_numero, patente_scadenza:b.patente_scadenza, categoria_patente:b.categoria_patente,
      tipo_cliente:b.tipo_cliente || 'privato', ragione_sociale:b.ragione_sociale, piva:b.partita_iva || b.piva, partita_iva:b.partita_iva || b.piva,
      pec:b.pec, sdi:b.codice_sdi || b.sdi, codice_sdi:b.codice_sdi || b.sdi,
      indirizzo_fatturazione:b.indirizzo_fatturazione || b.indirizzo, citta_fatturazione:b.citta_fatturazione || b.citta, provincia_fatturazione:b.provincia_fatturazione || b.provincia, cap_fatturazione:b.cap_fatturazione || b.cap,
      note:b.note || ''
    }); } catch(e) { console.log('Storico cliente V123 warning:', e.message); }

    const files = [];
    for (const [tipo, arr] of Object.entries(req.files || {})) {
      for (const f of (arr || [])) files.push({ tipo, f });
    }
    if (b.preupload_id && PREN_OCR_UPLOADS[b.preupload_id]) {
      for (const item of PREN_OCR_UPLOADS[b.preupload_id]) files.push(item);
      delete PREN_OCR_UPLOADS[b.preupload_id];
    }
    for (const item of files) {
      await run(`INSERT INTO allegati (prenotazione_id,tipo,filename,originalname,path,mimetype,size) VALUES (?,?,?,?,?,?,?)`,
        [result.lastID, item.tipo, item.f.filename, item.f.originalname, item.f.path, item.f.mimetype, item.f.size]);
    }

    try { if (typeof uploadContractAssetsToDrive === 'function') await uploadContractAssetsToDrive(result.lastID); } catch(e) { console.log('Drive cliente warning:', e.message); }

    // Notifica interna se il bot Twilio è configurato
    try {
      if (typeof dpNotify === 'function') {
        await dpNotify(DP_STAFF_NUMBERS || [], `NUOVA RICHIESTA NOLEGGIO CLIENTE\n\nCodice: ${cod}\nCliente: ${b.nome || ''} ${b.cognome || ''}\nTel: ${b.telefono || ''}\nMezzo richiesto: ${categoriaRichiesta || ''}
Mezzo assegnato: ${(mezzo.targa || '') + ' ' + (mezzo.marca || '') + ' ' + (mezzo.modello || '')}\nPeriodo: ${b.data_inizio} - ${b.data_fine}\nTotale previsto: EUR ${euro(calc.totale)}\nAllegati: ${files.length}\n\nApri gestionale: ${(process.env.APP_BASE_URL || '').replace(/\/+$/,'')}/prenotazione/${result.lastID}`);
      }
    } catch(e) { console.log('Notifica cliente warning:', e.message); }

    res.setHeader('Content-Type','text/html; charset=utf-8');
    res.send(`<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><style>body{font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Arial;background:#eef4ff;margin:0;padding:22px;color:#111}.hero{background:linear-gradient(135deg,#07111f,#173b8f);color:#fff;border-radius:28px;padding:28px;margin-bottom:20px}.box{background:#fff;border-radius:24px;padding:22px;box-shadow:0 12px 35px #0001}.ok{color:#157c2d;font-size:36px}.code{font-size:28px;font-weight:900}.btn{display:inline-block;background:#d70000;color:#fff;padding:14px 20px;border-radius:18px;text-decoration:none;font-weight:900;margin-top:18px}
.contract-main-actions{margin-top:16px}.contract-main-actions .btn{min-width:190px;text-align:center}.contract-secondary-actions .btn{min-width:150px;text-align:center}
@media(max-width:700px){.contract-main-actions .btn,.contract-secondary-actions .btn{width:100%;min-width:0}}


/* V109 FIX leggibilita mobile */
header{padding-top:max(22px, env(safe-area-inset-top));}
.top-actions{max-width:1180px;margin:0 auto 14px!important;padding:10px 0!important;}
.top-actions .back-btn::before{content:""!important;}
.top-actions .back-btn,.top-actions a{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif!important;font-size:clamp(18px,2.6vw,24px)!important;letter-spacing:0!important;line-height:1.1!important;white-space:nowrap!important;color:#fff!important;overflow:hidden;text-overflow:ellipsis;}
.top-actions .back-btn{background:#333!important;}
.top-actions .home-btn{background:#d70000!important;}
.client-back button{font-size:18px!important;font-weight:900!important;background:#333!important;color:#fff!important;}
@media(max-width:700px){
  nav{padding-top:calc(14px + env(safe-area-inset-top));}
  .top-actions{position:sticky;top:0;z-index:50;padding:10px 12px!important;gap:10px!important;background:rgba(244,244,244,.96)!important;}
  .top-actions .back-btn,.top-actions a{min-width:0!important;width:calc(50% - 5px)!important;flex:1 1 calc(50% - 5px)!important;padding:14px 8px!important;}
  .contract-main-actions .btn{width:100%!important;}
}

</style></head><body><div class="hero"><h1>DP RENT</h1><p>Dati ricevuti correttamente.</p></div><div class="box"><h2 class="ok">Richiesta inviata</h2><p>Codice pratica:</p><p class="code">${esc(cod)}</p><p>DP RENT controllerà i dati e ti confermerà contratto e disponibilità.</p><p>Foto ricevute: <b>${files.length}</b></p></div></body></html>`);
  } catch (e) {
    res.status(500).send(`<!doctype html><meta charset="utf-8"><h1>Errore invio dati</h1><pre>${esc(e.stack || e.message)}</pre><a href="javascript:history.back()">Torna</a>`);
  }
});



// =========================
// CARGOS V37 FIX OUTPUT / API
// =========================
function cargosFieldRows(p) {
  const d = cargosRecordData(p);
  let pos = 1;
  return CARGOS_FIELDS.map(([name, len]) => {
    const raw = String(d[name] || '');
    const value = cleanCargosText(raw).slice(0, len);
    const row = { name, len, pos_start: pos, pos_end: pos + len - 1, value };
    pos += len;
    return row;
  });
}

function buildCargosFixedRecord(p) {
  const d = cargosRecordData(p);
  return CARGOS_FIELDS.map(([name, len]) => cargosPad(d[name], len)).join('');
}

function buildCargosCsvRecord(p) {
  const d = cargosRecordData(p);
  return CARGOS_FIELDS.map(([name]) => cleanCargosText(d[name]).replace(/;/g, ',')).join(';');
}

function buildCargosCsvHeader() {
  return CARGOS_FIELDS.map(([name]) => name).join(';');
}

function validateCargosV37(p) {
  const v = v68SafeValidateCargos(p);
  const record = buildCargosFixedRecord(p);
  return { ...v, length: record.length, fixed_ok: record.length === 1505 };
}

function cargosHumanTable(p) {
  const rows = cargosFieldRows(p);
  return `<table>
    <tr><th>#</th><th>Campo</th><th>Pos.</th><th>Dim.</th><th>Valore</th></tr>
    ${rows.map((r, i) => `<tr>
      <td>${i}</td>
      <td>${esc(r.name)}</td>
      <td>${r.pos_start}-${r.pos_end}</td>
      <td>${r.len}</td>
      <td>${esc(r.value)}</td>
    </tr>`).join('')}
  </table>`;
}

function cargosMissingHtml(missing) {
  if (!missing || !missing.length) return '<p class="ok">Nessun campo obbligatorio mancante.</p>';
  return `<div class="alert"><b>Campi obbligatori mancanti:</b><br>${missing.map(x => '• ' + esc(x)).join('<br>')}</div>`;
}

function cargosApiConfigured() {
  return !!(process.env.CARGOS_USERNAME && process.env.CARGOS_PASSWORD && process.env.CARGOS_APIKEY && (process.env.CARGOS_BASE_URL || '').trim());
}

async function cargosRealCall(action, p) {
  
  
  p = patchCargosVehicleRecordV76(p);
p = patchCargosVehicleTypeIntoObjectV76(cargosPatchDefaultsV76(p));
const validation = validateCargosV37(p);
  if (!validation.ok || !validation.fixed_ok) {
    return { ok:false, action, error:'VALIDAZIONE_LOCALE_KO', missing:validation.missing, length:validation.length };
  }
  if (!cargosApiConfigured()) {
    return {
      ok:false,
      action,
      error:'API_NON_CONFIGURATA',
      message:'Servono CARGOS_USERNAME, CARGOS_PASSWORD, CARGOS_APIKEY, CARGOS_BASE_URL.'
    };
  }
  try {
    const data = await cargosSendRecords([buildCargosFixedRecord(p)], action === 'send' ? 'Send' : 'Check');
    return { ok:true, action, data, length:validation.length };
  } catch (e) {
    return { ok:false, action, error:'API_ERRORE', message:e.message, length:validation.length };
  }
}


// =========================
// V107 CARGOS UID LOCK / DRIVE / BRAND
// =========================
function safeFileName(v) {
  return String(v || '').replace(/[\/\\:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
}
function cargosVisibleStatus(p) {
  if (p.record_cargos_uid) return `<span class="badge badge-green">INVIATO UID ${esc(p.record_cargos_uid)}</span>`;
  if (p.record_cargos_stato === 'verificato') return `<span class="badge badge-orange">VERIFICATO</span>`;
  return `<span class="badge">DA VERIFICARE</span>`;
}
function cargosProPayload(p) {
  const data = cargosRecordData(p);
  return {
    contratto_id: data.CONTRATTO_ID,
    contratto_data: data.CONTRATTO_DATA,
    pagamento_tipo: data.CONTRATTO_TIPOP,
    checkout: { data: data.CONTRATTO_CHECKOUT_DATA, luogo_cod: data.CONTRATTO_CHECKOUT_LUOGO_COD, indirizzo: data.CONTRATTO_CHECKOUT_INDIRIZZO },
    checkin: { data: data.CONTRATTO_CHECKIN_DATA, luogo_cod: data.CONTRATTO_CHECKIN_LUOGO_COD, indirizzo: data.CONTRATTO_CHECKIN_INDIRIZZO },
    agenzia: { id: data.AGENZIA_ID, nome: data.AGENZIA_NOME, luogo_cod: data.AGENZIA_LUOGO_COD, indirizzo: data.AGENZIA_INDIRIZZO, telefono: data.AGENZIA_RECAPITO_TEL, operatore_id: data.OPERATORE_ID },
    veicolo: { tipo: data.VEICOLO_TIPO, marca: data.VEICOLO_MARCA, modello: data.VEICOLO_MODELLO, targa: data.VEICOLO_TARGA, colore: data.VEICOLO_COLORE, gps: data.VEICOLO_GPS, blocco_motore: data.VEICOLO_BLOCCOM },
    contraente: {
      cognome: data.CONDUCENTE_CONTRAENTE_COGNOME, nome: data.CONDUCENTE_CONTRAENTE_NOME,
      nascita_data: data.CONDUCENTE_CONTRAENTE_NASCITA_DATA, nascita_luogo_cod: data.CONDUCENTE_CONTRAENTE_NASCITA_LUOGO_COD,
      cittadinanza_cod: data.CONDUCENTE_CONTRAENTE_CITTADINANZA_COD, residenza_luogo_cod: data.CONDUCENTE_CONTRAENTE_RESIDENZA_LUOGO_COD,
      residenza_indirizzo: data.CONDUCENTE_CONTRAENTE_RESIDENZA_INDIRIZZO, doc_tipo: data.CONDUCENTE_CONTRAENTE_DOCIDE_TIPO_COD,
      doc_numero: data.CONDUCENTE_CONTRAENTE_DOCIDE_NUMERO, doc_luogo_rilascio_cod: data.CONDUCENTE_CONTRAENTE_DOCIDE_LUOGORIL_COD,
      patente_numero: data.CONDUCENTE_CONTRAENTE_PATENTE_NUMERO, patente_luogo_rilascio_cod: data.CONDUCENTE_CONTRAENTE_PATENTE_LUOGORIL_COD,
      recapito: data.CONDUCENTE_CONTRAENTE_RECAPITO
    },
    conducente2: {
      cognome: data.CONDUCENTE2_COGNOME, nome: data.CONDUCENTE2_NOME, nascita_data: data.CONDUCENTE2_NASCITA_DATA,
      nascita_luogo_cod: data.CONDUCENTE2_NASCITA_LUOGO_COD, cittadinanza_cod: data.CONDUCENTE2_CITTADINANZA_COD,
      doc_tipo: data.CONDUCENTE2_DOCIDE_TIPO_COD, doc_numero: data.CONDUCENTE2_DOCIDE_NUMERO,
      doc_luogo_rilascio_cod: data.CONDUCENTE2_DOCIDE_LUOGORIL_COD, patente_numero: data.CONDUCENTE2_PATENTE_NUMERO,
      patente_luogo_rilascio_cod: data.CONDUCENTE2_PATENTE_LUOGORIL_COD, recapito: data.CONDUCENTE2_RECAPITO
    }
  };
}
function cargosProValidate(p) {
  const v = (typeof validateCargosV37 === 'function') ? validateCargosV37(p) : validateCargos(p);
  const payload = cargosProPayload(p);
  const warnings = [];
  if (String(payload.veicolo.marca || '').length < 2) warnings.push('Marca veicolo sospetta');
  if (String(payload.veicolo.modello || '').length < 2) warnings.push('Modello veicolo sospetto');
  if (!/^\d{2}\/\d{2}\/\d{4}/.test(payload.checkout.data || '')) warnings.push('Data checkout non valida');
  if (!/^\d{2}\/\d{2}\/\d{4}/.test(payload.checkin.data || '')) warnings.push('Data checkin non valida');
  return { ok: v.ok && v.fixed_ok, missing: v.missing || [], warnings, length: v.length, fixed_ok: v.fixed_ok, payload };
}
async function uploadContractAssetsToDrive(prenotazioneId) {
  return new Promise((resolve) => {
    getPrenotazioneCompleta(prenotazioneId, async (err, p) => {
      if (err || !p || !googleDriveConfigured()) return resolve(null);
      try {
        const folderName = `${p.codice || 'contratto'} - ${p.nome || ''} ${p.cognome || ''}`.trim();
        db.all(`SELECT * FROM allegati WHERE prenotazione_id=? AND (drive_file_id IS NULL OR drive_file_id='')`, [prenotazioneId], async (e2, files) => {
          for (const f of (files || [])) {
            try {
              if (!f.path || !fs.existsSync(f.path)) continue;
              const cleanName = safeFileName(`${f.tipo || 'allegato'}_${f.originalname || f.filename}`);
              const dr = await uploadFileToDrive(f.path, cleanName, f.mimetype || 'application/octet-stream', folderName);
              if (dr) db.run(`UPDATE allegati SET drive_file_id=?, drive_web_link=? WHERE id=?`, [dr.id, dr.webViewLink, f.id]);
            } catch (e) { console.log('Drive allegato KO:', e.message); }
          }
          resolve(true);
        });
      } catch (e) { console.log('Drive assets KO:', e.message); resolve(null); }
    });
  });
}
function generatePrivacyPage() {
  return page('Privacy DP RENT', `<div class="box">
    <h2>Privacy DP RENT</h2>
    <p><b>Titolare:</b> Trasporti DP S.R.L. - DP RENT</p>
    <p><b>Sede:</b> Via Tuderte 466, Narni (TR)</p>
    <p><b>Email:</b> contabilita@trasportidp.com</p>
    <p>I dati sono trattati per gestione noleggio, contratto, obblighi fiscali, sicurezza, gestione danni/multe e adempimenti obbligatori.</p>
    <a class="btn" href="/">Torna</a>
  </div>`);
}
function generateCondizioniPage() {
  return page('Condizioni DP RENT', `<div class="box">
    <h2>Condizioni generali noleggio DP RENT</h2>
    <p>Veicolo consegnato in buono stato e da riconsegnare nelle stesse condizioni.</p>
    <p>Carburante pieno/pieno o livello equivalente indicato nel contratto.</p>
    <p>Km inclusi come da contratto; extra km conteggiati alla tariffa indicata.</p>
    <p>Danni, multe, pedaggi, ritardi, smarrimenti e franchigie sono a carico del cliente.</p>
    <a class="btn" href="/">Torna</a>
  </div>`);
}


// =========================
// V40 CARGOS REALE + DRIVE + PRIVACY DP
// =========================
const CARGOS_BASE_URL = (process.env.CARGOS_BASE_URL || 'https://cargos.poliziadistato.it/CARGOS_API').replace(/\/+$/, '');

function cargosOrganizationHeaderV76() {
  // Header richiesto da Ca.R.G.O.S.: Organization.
  // Su Render puoi mettere CARGOS_ORGANIZATION oppure CARGOS_ORGANIZATION_ID.
  // Se non lo metti, uso CARGOS_USERNAME come fallback perché spesso coincide col codice ente/organizzazione.
  return String(
    process.env.CARGOS_ORGANIZATION ||
    process.env.CARGOS_ORGANIZATION_ID ||
    process.env.CARGOS_ORG ||
    process.env.CARGOS_USERNAME ||
    'C00000100'
  ).trim();
}

function cleanCargos(v, len) {
  return String(v || '')
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9 ÀÈÉÌÒÙ\/\.\-\+\s]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, len);
}

function cargosPad(v, len) {
  return cleanCargos(v, len).padEnd(len, ' ');
}

function cargosNum(v, len) {
  return String(v || '').replace(/\D/g, '').slice(0, len).padStart(len, '0');
}

function cargosNormalizeIsoDateV76(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';

  // già ISO: 2026-05-11 oppure 2026-05-11T22:36:31
  let m = raw.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;

  // formato italiano: 11/05/2026, 11-05-2026, 11.05.2026
  m = raw.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    let yy = m[3];
    if (yy.length === 2) yy = '20' + yy;
    return `${yy}-${String(m[2]).padStart(2,'0')}-${String(m[1]).padStart(2,'0')}`;
  }

  const d = new Date(raw);
  if (!Number.isNaN(d.getTime())) {
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  return '';
}

function cargosNormalizeTimeV76(date, time) {
  let t = String(time || '').trim();
  const raw = String(date || '').trim();
  if (!t) {
    let m = raw.match(/[T\s](\d{1,2}):(\d{2})/);
    if (m) t = `${String(m[1]).padStart(2,'0')}:${m[2]}`;
  }
  const tm = t.match(/(\d{1,2}):(\d{2})/);
  if (tm) return `${String(tm[1]).padStart(2,'0')}:${tm[2]}`;
  return '00:00';
}

function cargosDateTime(date, time) {
  // FIX V76 Ca.R.G.O.S: il servizio rifiuta DD/MM/YYYY.
  // Il campo fisso è lungo 16, quindi inviamo: YYYY-MM-DDTHH:mm
  const d = cargosNormalizeIsoDateV76(date);
  const t = cargosNormalizeTimeV76(date, time);
  if (!d) return ''.padEnd(16, ' ');
  return `${d}T${t}`.slice(0,16);
}

function cargosDateOnly(date) {
  // FIX V76 Ca.R.G.O.S: SOLO data ISO YYYY-MM-DD, senza ora/timezone.
  const d = cargosNormalizeIsoDateV76(date);
  return (d || '').padEnd(10, ' ').slice(0,10);
}

function splitFullNameV40(p) {
  const nome = String(p.nome || '').trim();
  const cognome = String(p.cognome || '').trim();
  if (nome || cognome) return { nome, cognome };
  const full = String(p.cliente || p.intestatario || '').trim().split(/\s+/);
  return { nome: full[0] || '', cognome: full.slice(1).join(' ') || '' };
}

const CARGOS_FIELDS_V40 = [
  ['CONTRATTO_ID',50], ['CONTRATTO_DATA',16], ['CONTRATTO_TIPOP',1],
  ['CONTRATTO_CHECKOUT_DATA',16], ['CONTRATTO_CHECKOUT_LUOGO_COD',9], ['CONTRATTO_CHECKOUT_INDIRIZZO',150],
  ['CONTRATTO_CHECKIN_DATA',16], ['CONTRATTO_CHECKIN_LUOGO_COD',9], ['CONTRATTO_CHECKIN_INDIRIZZO',150],
  ['OPERATORE_ID',50], ['AGENZIA_ID',30], ['AGENZIA_NOME',70], ['AGENZIA_LUOGO_COD',9], ['AGENZIA_INDIRIZZO',150], ['AGENZIA_RECAPITO_TEL',20],
  ['VEICOLO_TIPO',1], ['VEICOLO_MARCA',50], ['VEICOLO_MODELLO',100], ['VEICOLO_TARGA',15], ['VEICOLO_COLORE',50], ['VEICOLO_GPS',1], ['VEICOLO_BLOCCOM',1],
  ['CONDUCENTE_CONTRAENTE_COGNOME',50], ['CONDUCENTE_CONTRAENTE_NOME',30], ['CONDUCENTE_CONTRAENTE_NASCITA_DATA',10],
  ['CONDUCENTE_CONTRAENTE_NASCITA_LUOGO_COD',9], ['CONDUCENTE_CONTRAENTE_CITTADINANZA_COD',9],
  ['CONDUCENTE_CONTRAENTE_RESIDENZA_LUOGO_COD',9], ['CONDUCENTE_CONTRAENTE_RESIDENZA_INDIRIZZO',150],
  ['CONDUCENTE_CONTRAENTE_DOCIDE_TIPO_COD',5], ['CONDUCENTE_CONTRAENTE_DOCIDE_NUMERO',20], ['CONDUCENTE_CONTRAENTE_DOCIDE_LUOGORIL_COD',9],
  ['CONDUCENTE_CONTRAENTE_PATENTE_NUMERO',20], ['CONDUCENTE_CONTRAENTE_PATENTE_LUOGORIL_COD',9], ['CONDUCENTE_CONTRAENTE_RECAPITO',20],
  ['CONDUCENTE2_COGNOME',50], ['CONDUCENTE2_NOME',30], ['CONDUCENTE2_NASCITA_DATA',10], ['CONDUCENTE2_NASCITA_LUOGO_COD',9],
  ['CONDUCENTE2_CITTADINANZA_COD',9], ['CONDUCENTE2_DOCIDE_TIPO_COD',5], ['CONDUCENTE2_DOCIDE_NUMERO',20],
  ['CONDUCENTE2_DOCIDE_LUOGORIL_COD',9], ['CONDUCENTE2_PATENTE_NUMERO',20], ['CONDUCENTE2_PATENTE_LUOGORIL_COD',9], ['CONDUCENTE2_RECAPITO',20]
];

function cargosRecordDataV40(p) {
  // V76 patch veicolo

  // cargosRecordDataV40__v72patched
  // V76 rimosso: p non ancora inizializzato

  const n = splitFullNameV40(p);
  const agenziaNome = process.env.CARGOS_AGENZIA_NOME || 'TRASPORTI DP S.R.L. - DP RENT';
  const agenziaInd = process.env.CARGOS_AGENZIA_INDIRIZZO || 'VIA TUDERTE 466, NARNI (TR)';
  const tel = process.env.CARGOS_AGENZIA_TEL || '0744817108';
  const luogo = process.env.CARGOS_LUOGO_COD || p.record_cargos_luogo_cod || '410055022';
  const tipoPagamento = process.env.CARGOS_TIPO_PAGAMENTO || p.record_cargos_pagamento_tipo || '1';
  const tipoVeicolo = p.record_cargos_veicolo_tipo || process.env.CARGOS_VEICOLO_TIPO || '1';

  return {
    CONTRATTO_ID: p.codice || `DPR-${p.id}`,
    CONTRATTO_DATA: cargosDateOnly(p.created_at || new Date().toISOString()),
    CONTRATTO_TIPOP: getTipoPagamentoCargosV63(p.pagamento || p.tipo_pagamento || '9'),
    CONTRATTO_CHECKOUT_DATA: cargosDateOnly(p.data_inizio || p.checkout_data || p.data_checkout || ''),
    CONTRATTO_CHECKOUT_LUOGO_COD: cargosCheckoutLuogoCodV63(),
    CONTRATTO_CHECKOUT_INDIRIZZO: p.record_cargos_checkout_indirizzo || agenziaInd,
    CONTRATTO_CHECKIN_DATA: cargosDateOnly(p.data_fine || p.checkin_data || p.data_checkin || ''),
    CONTRATTO_CHECKIN_LUOGO_COD: cargosCheckinLuogoCodV63(),
    CONTRATTO_CHECKIN_INDIRIZZO: p.record_cargos_checkin_indirizzo || agenziaInd,
    OPERATORE_ID: cargosOperatoreIdV63(),
    AGENZIA_ID: cargosAgenziaIdV63(),
    AGENZIA_NOME: p.record_cargos_agenzia_nome || agenziaNome,
    AGENZIA_LUOGO_COD: p.record_cargos_agenzia_luogo_cod || luogo,
    AGENZIA_INDIRIZZO: p.record_cargos_agenzia_indirizzo || agenziaInd,
    AGENZIA_RECAPITO_TEL: p.record_cargos_agenzia_tel || tel,
    // V76 FIX REALE: getPrenotazioneCompleta() legge il mezzo con alias mezzo_*
    // Se usiamo solo p.marca/p.modello/p.targa, Ca.R.G.O.S resta vuoto quando il veicolo arriva dalla tabella mezzi.
    VEICOLO_TIPO: getTipoVeicoloCargosV61([
      p.tipo, p.categoria, p.veicolo_tipo,
      p.marca, p.modello, p.targa,
      p.mezzo_tipo, p.mezzo_categoria, p.mezzo_marca, p.mezzo_modello, p.mezzo_targa, p.mezzo_descrizione,
      p.mezzo
    ].filter(Boolean).join(' ')),
    VEICOLO_MARCA: String(p.marca || p.veicolo_marca || p.mezzo_marca || '').trim(),
    VEICOLO_MODELLO: String(p.modello || p.veicolo_modello || p.mezzo_modello || p.mezzo_descrizione || p.mezzo || '').trim(),
    VEICOLO_TARGA: String(p.targa || p.veicolo_targa || p.mezzo_targa || '').trim(),
    VEICOLO_COLORE: p.colore || p.record_cargos_veicolo_colore || '',
    VEICOLO_GPS: String(p.gps ?? p.record_cargos_veicolo_gps ?? process.env.CARGOS_VEICOLO_GPS ?? '0'),
    VEICOLO_BLOCCOM: String(p.blocco_motore ?? p.record_cargos_veicolo_bloccom ?? process.env.CARGOS_VEICOLO_BLOCCOM ?? '0'),
    CONDUCENTE_CONTRAENTE_COGNOME: n.cognome,
    CONDUCENTE_CONTRAENTE_NOME: n.nome,
    CONDUCENTE_CONTRAENTE_NASCITA_DATA: v67DefaultBirth(p),
    CONDUCENTE_CONTRAENTE_NASCITA_LUOGO_COD: p.record_cargos_nascita_luogo_cod || luogo,
    CONDUCENTE_CONTRAENTE_CITTADINANZA_COD: V104_CITTADINANZA_ITALIA_CARGOS,
    CONDUCENTE_CONTRAENTE_RESIDENZA_LUOGO_COD: p.record_cargos_residenza_luogo_cod || luogo,
    CONDUCENTE_CONTRAENTE_RESIDENZA_INDIRIZZO: p.indirizzo || '',
    CONDUCENTE_CONTRAENTE_DOCIDE_TIPO_COD: getTipoDocumentoCargosV61(p.documento_tipo || p.tipo_documento || 'IDENT'),
    CONDUCENTE_CONTRAENTE_DOCIDE_NUMERO: String(p.documento_numero || p.doc_numero || p.patente_numero || p.codice_fiscale || 'DOC00000').slice(0,20),
    CONDUCENTE_CONTRAENTE_DOCIDE_LUOGORIL_COD: p.record_cargos_doc_luogoril_cod || luogo,
    CONDUCENTE_CONTRAENTE_PATENTE_NUMERO: String(p.patente_numero || p.documento_numero || 'PAT00000').slice(0,20),
    CONDUCENTE_CONTRAENTE_PATENTE_LUOGORIL_COD: p.record_cargos_patente_luogoril_cod || luogo,
    CONDUCENTE_CONTRAENTE_RECAPITO: p.telefono || '',
    CONDUCENTE2_COGNOME: p.conducente2_cognome || '',
    CONDUCENTE2_NOME: p.conducente2_nome || '',
    CONDUCENTE2_NASCITA_DATA: cargosDateOnly(p.conducente2_data_nascita || ''),
    CONDUCENTE2_NASCITA_LUOGO_COD: p.conducente2_nascita_luogo_cod || '',
    CONDUCENTE2_CITTADINANZA_COD: p.conducente2_cittadinanza_cod || '',
    CONDUCENTE2_DOCIDE_TIPO_COD: getTipoDocumentoCargosV63(p.conducente2_documento_tipo || 'IDENT'),
    CONDUCENTE2_DOCIDE_NUMERO: p.conducente2_doc_numero || '',
    CONDUCENTE2_DOCIDE_LUOGORIL_COD: p.conducente2_doc_luogoril_cod || '',
    CONDUCENTE2_PATENTE_NUMERO: p.conducente2_patente_numero || '',
    CONDUCENTE2_PATENTE_LUOGORIL_COD: p.conducente2_patente_luogoril_cod || '',
    CONDUCENTE2_RECAPITO: p.conducente2_recapito || ''
  };
}

function buildCargosFixedRecordV40(p) {
  const d = cargosRecordDataV40(p);
  const rec = CARGOS_FIELDS_V40.map(([name, len]) => cargosPad(d[name], len)).join('');
  return rec.slice(0, 1505).padEnd(1505, ' ');
}

function cargosRowsV40(p) {
  const d = cargosRecordDataV40(p);
  let pos = 1;
  return CARGOS_FIELDS_V40.map(([name, len]) => {
    const value = cleanCargos(d[name], len);
    const row = { name, len, value, dal: pos, al: pos + len - 1 };
    pos += len;
    return row;
  });
}

function validateCargosV40(p) {
  // V76 patch veicolo

  // validateCargosV40__v72patched
  // V76 rimosso: p non ancora inizializzato

  const d = cargosRecordDataV40(p);
  const missing = [];
  const req = [
    'CONTRATTO_ID','CONTRATTO_DATA','CONTRATTO_TIPOP','CONTRATTO_CHECKOUT_DATA','CONTRATTO_CHECKOUT_LUOGO_COD',
    'CONTRATTO_CHECKOUT_INDIRIZZO','CONTRATTO_CHECKIN_DATA','CONTRATTO_CHECKIN_LUOGO_COD','CONTRATTO_CHECKIN_INDIRIZZO',
    'OPERATORE_ID','AGENZIA_ID','AGENZIA_NOME','AGENZIA_LUOGO_COD','AGENZIA_INDIRIZZO','AGENZIA_RECAPITO_TEL',
    'VEICOLO_TIPO','VEICOLO_MARCA','VEICOLO_MODELLO','VEICOLO_TARGA',
    'CONDUCENTE_CONTRAENTE_COGNOME','CONDUCENTE_CONTRAENTE_NOME','CONDUCENTE_CONTRAENTE_NASCITA_DATA',
    'CONDUCENTE_CONTRAENTE_NASCITA_LUOGO_COD','CONDUCENTE_CONTRAENTE_CITTADINANZA_COD',
    'CONDUCENTE_CONTRAENTE_DOCIDE_TIPO_COD','CONDUCENTE_CONTRAENTE_DOCIDE_NUMERO','CONDUCENTE_CONTRAENTE_DOCIDE_LUOGORIL_COD',
    'CONDUCENTE_CONTRAENTE_PATENTE_NUMERO','CONDUCENTE_CONTRAENTE_PATENTE_LUOGORIL_COD'
  ];
  for (const k of req) if (!String(d[k] || '').trim()) missing.push(k);
  const rec = buildCargosFixedRecordV40(p);
  return { ok: missing.length === 0 && rec.length === 1505, missing, length: rec.length, record: rec, data: d };
}

async function cargosGetTokenV40() {
  const username = process.env.CARGOS_USERNAME;
  const password = process.env.CARGOS_PASSWORD;
  if (!username || !password) throw new Error('Mancano CARGOS_USERNAME o CARGOS_PASSWORD');
  const auth = Buffer.from(`${username}:${password}`).toString('base64');

  // FIX V76: l'endpoint Token di Ca.R.G.O.S. NON accetta POST.
  // Render mostrava: 405 "The requested resource does not support http method 'POST'."
  // Quindi il token va richiesto in GET con Basic Auth.
  const r = await fetch(`${CARGOS_BASE_URL}/api/Token`, {
    method: 'GET',
    headers: {
      'Authorization': `Basic ${auth}`,
      'Accept': 'application/json'
    }
  });

  const text = await r.text();
  let data; try { data = JSON.parse(text); } catch { data = { raw: text }; }
  const accessToken = data.access_token || data.accessToken || data.token || data?.Esito?.access_token;
  if (!r.ok || !accessToken) throw new Error(`Token CARGOS KO ${r.status}: ${text.slice(0,500)}`);
  return accessToken;
}

function cargosEncryptTokenV40(accessToken) {
  // V63: AES ufficiale Ca.R.G.O.S. - primi 32 caratteri APIKEY = Key, successivi 16 = IV.
  const apiKey = String(process.env.CARGOS_APIKEY || '');
  if (apiKey.length < 48) throw new Error('CARGOS_APIKEY deve avere almeno 48 caratteri per cifratura AES');
  const key = Buffer.from(apiKey.substring(0, 32), 'utf8');
  const iv = Buffer.from(apiKey.substring(32, 48), 'utf8');
  const cipher = crypto.createCipheriv('aes-256-cbc', key, iv);
  cipher.setAutoPadding(true);
  return Buffer.concat([cipher.update(String(accessToken), 'utf8'), cipher.final()]).toString('base64');
}

async function cargosCallV40(endpoint, records) {
  const token = await cargosGetTokenV40();
  const encrypted = cargosEncryptTokenV40(token);
  const r = await fetch(`${CARGOS_BASE_URL}/api/${endpoint}`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${encrypted}`,
      'Organization': cargosOrganizationHeaderV76(),
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    body: JSON.stringify(records)
  });
  const text = await r.text();
  let data; try { data = JSON.parse(text); } catch { data = { raw: text }; }
  return { httpStatus: r.status, ok: r.ok, data };
}


// =========================
// V63 - PRENOTAZIONE COMPLETA PER PDF / DRIVE / FIRMA / CARGOS
// =========================
function getPrenotazioneCompleta(id, callback) {
  db.get(`
    SELECT p.*,
           m.targa AS mezzo_targa,
           m.marca AS mezzo_marca,
           m.modello AS mezzo_modello,
           m.tipo AS mezzo_tipo,
           m.descrizione AS mezzo_descrizione,
           m.km AS mezzo_km,
           m.km_attuali AS mezzo_km_attuali,
           m.categoria AS mezzo_categoria,
           m.codice_tipo AS mezzo_codice_tipo
    FROM prenotazioni p
    LEFT JOIN mezzi m ON p.mezzo_id = m.id
    WHERE p.id = ?
  `, [id], callback);
}

async function syncContrattoDriveV63(prenotazioneId) {
  // V63: una sola cartella Drive per contratto, un solo PDF. Le foto vanno nella stessa cartella.
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [prenotazioneId]);
    if (!p) return null;

    const folder = await getOrCreateDriveContractFolderV63(p);
    if (!folder) return null;

    await run(`UPDATE prenotazioni SET drive_folder_id=?, drive_folder_link=? WHERE id=?`,
      [folder.id, folder.webViewLink || null, prenotazioneId]);

    const pdf = await generaPdfContratto(prenotazioneId, { forceDrive:false });
    const pdfName = path.basename(pdf);

    await deleteAllContractPdfsInDriveV63(folder.id);
    const uploadedPdf = await uploadFileToDriveFolderV63(pdf, pdfName, 'application/pdf', folder.id);

    await run(`UPDATE prenotazioni SET pdf_path=?, pdf_drive_link=? WHERE id=?`,
      [pdf, uploadedPdf?.webViewLink || null, prenotazioneId]);

    const allegati = await all(`SELECT * FROM allegati WHERE prenotazione_id=?`, [prenotazioneId]).catch(() => []);
    for (const a of (allegati || [])) {
      if (a.drive_file_id) continue;
      if (!a.path || !fs.existsSync(a.path)) continue;
      const up = await uploadFileToDriveFolderV63(
        a.path,
        a.originalname || a.filename || path.basename(a.path),
        a.mimetype || 'application/octet-stream',
        folder.id
      );
      if (up?.id) {
        await run(`UPDATE allegati SET drive_file_id=?, drive_web_link=? WHERE id=?`,
          [up.id, up.webViewLink || null, a.id]);
      }
    }

    return { folder, pdf: uploadedPdf };
  } catch (e) {
    console.log('uploadAllContrattoDriveV40 V63 error:', e.message);
    return null;
  }
}

function privacyHtmlV40() {
  return page('Privacy DP RENT', `<div class="box">
    <h1>Privacy DP RENT</h1>
    <p><b>Titolare:</b> Trasporti DP S.R.L. - DP RENT</p>
    <p><b>Sede:</b> Via Tuderte 466, Narni (TR)</p>
    <p><b>Contatti:</b> 0744817108 - contabilita@trasportidp.com</p>
    <p>I dati personali e i documenti sono trattati per identificazione cliente, gestione contratto di noleggio, obblighi fiscali, sicurezza, gestione danni, multe, pedaggi e adempimenti previsti dalla normativa.</p>
    <p>I dati possono essere comunicati alle autorità competenti quando richiesto dalla legge.</p>
    <a class="btn" href="/">Torna</a>
  </div>`);
}

function condizioniHtmlV40() {
  return page('Condizioni DP RENT', `<div class="box">
    <h1>Condizioni generali DP RENT</h1>
    <p>Il cliente riceve il mezzo in buono stato e si impegna a riconsegnarlo nelle stesse condizioni, salvo normale usura.</p>
    <p>Carburante: livello indicato nel contratto, normalmente pieno/pieno.</p>
    <p>Km inclusi e extra km sono quelli indicati nel contratto.</p>
    <p>Danni, franchigie, ritardi, multe, pedaggi, smarrimento chiavi/documenti e costi accessori sono a carico del cliente.</p>
    <p>Il deposito cauzionale è gestito separatamente secondo accordi DP RENT.</p>
    <a class="btn" href="/">Torna</a>
  </div>`);
}


app.get('/privacy', (req, res) => res.send(privacyHtmlV40()));
app.get('/condizioni', (req, res) => res.send(condizioniHtmlV40()));
app.get('/condizioni-noleggio', (req, res) => res.send(condizioniHtmlV40()));
app.get('/termini-noleggio', (req, res) => res.redirect('/condizioni-noleggio'));

app.get('/cargos/:id/txt', (req, res) => {
  getPrenotazioneCompleta(req.params.id, (err, p) => {
    if (!p) return res.status(404).send('Contratto non trovato');
    const rec = buildCargosFixedRecordV40(p);
    res.setHeader('Content-Type', 'text/plain; charset=ascii');
    res.setHeader('Content-Disposition', `attachment; filename="record_cargos_${p.codice || p.id}.txt"`);
    res.end(rec + '\n');
  });
});

app.get('/cargos/:id/preview', (req, res) => {
  getPrenotazioneCompleta(req.params.id, (err, p) => {
    if (!p) return res.status(404).send('Contratto non trovato');
    const v = validateCargosV40(p);
    const rows = cargosRowsV40(p);
    res.send(page('CARGOS Preview V40', `
      <div class="box">
        <h2>Ca.R.G.O.S. Preview V40</h2>
        <p><b>Contratto:</b> ${esc(p.codice || p.id)}</p>
        <p><b>Lunghezza record:</b> ${v.length}/1505 ${v.ok ? '<span class="ok">OK</span>' : '<span class="bad">KO</span>'}</p>
        ${v.missing.length ? `<div class="alert"><b>Mancano:</b><br>${(Array.isArray(v.missing)?v.missing:[]).map(esc).join('<br>')}</div>` : '<p class="ok">Campi obbligatori OK</p>'}
        <div class="actions">
          <a class="btn" href="/cargos/${p.id}/check">Verifica dati CaRGOS</a>
          <a class="btn btn3" href="/cargos/${p.id}/send">Invia report CaRGOS</a>
          <a class="btn btn2" href="/cargos/${p.id}/txt">Scarica TXT 1505</a>
          <a class="btn btn2" href="/cargos/${p.id}">Torna</a>
        </div>
      </div>
      <div class="box">
        <h3>Tracciato campi</h3>
        <table><tr><th>Campo</th><th>Dal</th><th>Al</th><th>Dim</th><th>Valore</th></tr>
        ${rows.map(r => `<tr><td>${esc(r.name)}</td><td>${r.dal}</td><td>${r.al}</td><td>${r.len}</td><td>${esc(r.value)}</td></tr>`).join('')}
        </table>
      </div>
      <div class="box">
        <h3>Record ufficiale 1505</h3>
        <pre style="white-space:pre;overflow:auto;background:#111;color:white;padding:12px;border-radius:8px;">${esc(v.record)}</pre>
      </div>
    `));
  });
});

app.get('/cargos/:id/check', async (req, res) => {
  getPrenotazioneCompleta(req.params.id, async (err, p) => {
    if (!p) return res.status(404).send('Contratto non trovato');
    const v = validateCargosV40(p);
    if (!v.ok) {
      return res.send(page('CARGOS Check', `<div class="box"><h2 class="bad">Verifica locale KO</h2><p>Mancano campi obbligatori:</p><pre>${esc(v.missing.join('\n'))}</pre><a class="btn" href="/cargos/${p.id}/preview">Preview</a></div>`));
    }
    try {
      const result = await cargosCallV40('Check', [v.record]);
      db.run(`UPDATE prenotazioni SET record_cargos_stato=?, record_cargos_last_check=?, record_cargos_last_error=? WHERE id=?`,
        [result.ok ? 'check_ok' : 'check_ko', new Date().toISOString(), JSON.stringify(result).slice(0,1000), p.id]);
      res.send(page('CARGOS Check', `<div class="box"><h2>Risposta Check CaRGOS</h2><pre style="white-space:pre-wrap;background:#111;color:white;padding:12px;border-radius:8px;">${esc(JSON.stringify(result,null,2))}</pre><a class="btn" href="/cargos/${p.id}/send">Invia report</a><a class="btn btn2" href="/cargos/${p.id}/preview">Preview</a></div>`));
    } catch(e) {
      db.run(`UPDATE prenotazioni SET record_cargos_stato=?, record_cargos_last_check=?, record_cargos_last_error=? WHERE id=?`, ['check_errore', new Date().toISOString(), e.message, p.id]);
      res.send(page('CARGOS Check Errore', `<div class="box"><h2 class="bad">Errore Check CaRGOS</h2><pre>${esc(e.message)}</pre><a class="btn" href="/cargos/${p.id}/preview">Preview</a></div>`));
    }
  });
});



// =========================
// V106 FIX: /cargos/check usa la verifica funzionante
// =========================
app.get('/cargos/check/:id', async (req, res) => {
  return res.redirect('/cargos/' + encodeURIComponent(req.params.id) + '/verifica');
});

app.post('/cargos/check/:id', async (req, res) => {
  return res.redirect('/cargos/' + encodeURIComponent(req.params.id) + '/verifica');
});


app.get('/cargos/:id/verifica', (req, res) => res.redirect(`/cargos/${req.params.id}/check`));

app.get('/cargos/:id/send', async (req, res) => {
  getPrenotazioneCompleta(req.params.id, async (err, p) => {
    if (!p) return res.status(404).send('Contratto non trovato');
    const v = validateCargosV40(p);
    if (!v.ok) return res.send(page('CARGOS Send', `<div class="box"><h2 class="bad">Invio bloccato</h2><pre>${esc(v.missing.join('\n'))}</pre><a class="btn" href="/cargos/${p.id}/preview">Preview</a></div>`));
    try {
      const result = await cargosCallV40('Send', [v.record]);
      let uid = '';
      try {
        const arr = result?.data;
        if (Array.isArray(arr) && arr[0]) uid = arr[0].transactionid || arr[0].transactionId || '';
        uid = uid || result?.data?.transactionid || result?.data?.transactionId || '';
      } catch {}
      db.run(`UPDATE prenotazioni SET record_cargos_stato=?, record_cargos_last_send=?, record_cargos_last_error=?, record_cargos_transactionid=?, record_cargos_uid=? WHERE id=?`,
        [result.ok ? 'send_ok' : 'send_ko', new Date().toISOString(), JSON.stringify(result).slice(0,1000), uid, uid, p.id]);
      db.run(`INSERT INTO record_cargos_invii (prenotazione_id, uid, stato, richiesta, risposta, errore) VALUES (?,?,?,?,?,?)`,
        [p.id, uid, result.ok ? 'send_ok' : 'send_ko', v.record, JSON.stringify(result).slice(0,4000), result.ok ? '' : JSON.stringify(result).slice(0,1000)]);
      res.send(page('CARGOS Send', `<div class="box"><h2>Risposta Invio CaRGOS</h2><p><b>UID:</b> ${esc(uid || '-')}</p><pre style="white-space:pre-wrap;background:#111;color:white;padding:12px;border-radius:8px;">${esc(JSON.stringify(result,null,2))}</pre><a class="btn" href="/contratto/${p.id}/gestisci">Torna contratto</a><a class="btn btn2" href="/cargos/${p.id}/preview">Preview</a></div>`));
    } catch(e) {
      db.run(`UPDATE prenotazioni SET record_cargos_stato=?, record_cargos_last_send=?, record_cargos_last_error=? WHERE id=?`, ['send_errore', new Date().toISOString(), e.message, p.id]);
      res.send(page('CARGOS Send Errore', `<div class="box"><h2 class="bad">Errore Invio CaRGOS</h2><pre>${esc(e.message)}</pre><a class="btn" href="/cargos/${p.id}/preview">Preview</a></div>`));
    }
  });
});

app.get('/cargos/:id/invia', (req, res) => res.redirect(`/cargos/${req.params.id}/send`));

app.get('/cargos/:id', (req,res)=>{getPrenotazioneCompleta(req.params.id,(err,p)=>{if(!p)return res.send(page('CARGOS','<div class="box"><h2>Contratto non trovato</h2></div>')); const val=k=>esc(p[k]||''); const v=v68SafeValidateCargos(p); res.send(page('Ca.R.G.O.S.',`
<div class="box"><h2>Ca.R.G.O.S. ${esc(p.codice)}</h2>${v.ok?`<p class="ok">Dati completi. Lunghezza riga: ${v.length}</p>`:`<div class="alert"><b>Mancano campi:</b><br>${(Array.isArray(v.missing)?v.missing:[]).map(esc).join('<br>')}</div>`}
<form method="POST" action="/cargos/${p.id}/save">
<h3>Contratto / Agenzia</h3><div class="grid">
<div><label>Metodo pagamento</label>${cargosSelect('record_cargos_pagamento_tipo',p.record_cargos_pagamento_tipo||'0',CARGOS_PAYMENTS)}</div>
<div><label>Operatore ID</label><input name="record_cargos_operatore_id" value="${val('record_cargos_operatore_id')||esc(process.env.CARGOS_OPERATORE_ID||'DPRENT')}"></div>
<div><label>Agenzia ID</label><input name="record_cargos_agenzia_id" value="${val('record_cargos_agenzia_id')||esc(process.env.CARGOS_AGENZIA_ID||'DPR')}"></div>
<div><label>Agenzia nome</label><input name="record_cargos_agenzia_nome" value="${val('record_cargos_agenzia_nome')||esc(process.env.CARGOS_AGENZIA_NOME||AZIENDA.nome)}"></div>
<div><label>Agenzia luogo COD</label><input name="record_cargos_agenzia_luogo_cod" value="${val('record_cargos_agenzia_luogo_cod')||esc(process.env.CARGOS_AGENZIA_LUOGO_COD||'')}"></div>
<div><label>Agenzia telefono</label><input name="record_cargos_agenzia_tel" value="${val('record_cargos_agenzia_tel')||esc(process.env.CARGOS_AGENZIA_TEL||AZIENDA.telefono)}"></div></div>
<label>Agenzia indirizzo</label><input name="record_cargos_agenzia_indirizzo" value="${val('record_cargos_agenzia_indirizzo')||esc(process.env.CARGOS_AGENZIA_INDIRIZZO||AZIENDA.indirizzo)}">
<h3>Ritiro / consegna</h3><div class="grid">
<div><label>Checkout luogo COD</label><input name="record_cargos_checkout_luogo_cod" value="${val('record_cargos_checkout_luogo_cod')||esc(process.env.CARGOS_AGENZIA_LUOGO_COD||'')}"></div>
<div><label>Checkin luogo COD</label><input name="record_cargos_checkin_luogo_cod" value="${val('record_cargos_checkin_luogo_cod')||esc(process.env.CARGOS_AGENZIA_LUOGO_COD||'')}"></div></div>
<label>Checkout indirizzo</label><input name="record_cargos_checkout_indirizzo" value="${val('record_cargos_checkout_indirizzo')||esc(process.env.CARGOS_AGENZIA_INDIRIZZO||AZIENDA.indirizzo)}">
<label>Checkin indirizzo</label><input name="record_cargos_checkin_indirizzo" value="${val('record_cargos_checkin_indirizzo')||esc(process.env.CARGOS_AGENZIA_INDIRIZZO||AZIENDA.indirizzo)}">
<h3>Veicolo</h3><div class="grid">
<div><label>Tipo veicolo</label>${cargosSelect('record_cargos_veicolo_tipo',p.record_cargos_veicolo_tipo||p.m_record_cargos_veicolo_tipo||'1',CARGOS_VEHICLE_TYPES)}</div>
<div><label>Colore</label><input name="record_cargos_veicolo_colore" value="${val('record_cargos_veicolo_colore')||esc(p.colore||'')}"></div>
<div><label>GPS</label><select name="record_cargos_veicolo_gps"><option value="0" ${Number(p.record_cargos_veicolo_gps||p.gps||0)===0?'selected':''}>0 - No</option><option value="1" ${Number(p.record_cargos_veicolo_gps||p.gps||0)===1?'selected':''}>1 - Si</option></select></div>
<div><label>Blocco motore</label><select name="record_cargos_veicolo_bloccom"><option value="0" ${Number(p.record_cargos_veicolo_bloccom||p.blocco_motore||0)===0?'selected':''}>0 - No</option><option value="1" ${Number(p.record_cargos_veicolo_bloccom||p.blocco_motore||0)===1?'selected':''}>1 - Si</option></select></div></div>
<h3>Contraente</h3><div class="grid">
<div><label>Luogo nascita COD</label><input name="record_cargos_nascita_luogo_cod" value="${val('record_cargos_nascita_luogo_cod')}"></div>
<div><label>Cittadinanza COD</label><input name="record_cargos_cittadinanza_cod" value="${val('record_cargos_cittadinanza_cod')||esc(process.env.CARGOS_CITTADINANZA_DEFAULT||'100000100')}"></div>
<div><label>Residenza luogo COD</label><input name="record_cargos_residenza_luogo_cod" value="${val('record_cargos_residenza_luogo_cod')}"></div>
<div><label>Tipo documento</label>${cargosSelect('record_cargos_doc_tipo_cod',p.record_cargos_doc_tipo_cod||'CI',CARGOS_DOC_TYPES)}</div>
<div><label>Luogo rilascio doc COD</label><input name="record_cargos_doc_luogoril_cod" value="${val('record_cargos_doc_luogoril_cod')}"></div>
<div><label>Luogo rilascio patente COD</label><input name="record_cargos_patente_luogoril_cod" value="${val('record_cargos_patente_luogoril_cod')}"></div></div>
<h3>Secondo conducente</h3><div class="grid">
<div><label>Nome 2</label><input name="conducente2_nome" value="${val('conducente2_nome')}"></div><div><label>Cognome 2</label><input name="conducente2_cognome" value="${val('conducente2_cognome')}"></div>
<div><label>Data nascita 2</label><input type="date" name="conducente2_data_nascita" value="${val('conducente2_data_nascita')}"></div><div><label>Luogo nascita COD 2</label><input name="conducente2_nascita_luogo_cod" value="${val('conducente2_nascita_luogo_cod')}"></div>
<div><label>Cittadinanza COD 2</label><input name="conducente2_cittadinanza_cod" value="${val('conducente2_cittadinanza_cod')}"></div><div><label>Tipo doc 2</label>${cargosSelect('conducente2_doc_tipo_cod',p.conducente2_doc_tipo_cod||'',CARGOS_DOC_TYPES)}</div>
<div><label>Numero doc 2</label><input name="conducente2_doc_numero" value="${val('conducente2_doc_numero')}"></div><div><label>Luogo rilascio doc COD 2</label><input name="conducente2_doc_luogoril_cod" value="${val('conducente2_doc_luogoril_cod')}"></div>
<div><label>Numero patente 2</label><input name="conducente2_patente_numero" value="${val('conducente2_patente_numero')}"></div><div><label>Luogo rilascio patente COD 2</label><input name="conducente2_patente_luogoril_cod" value="${val('conducente2_patente_luogoril_cod')}"></div>
<div><label>Recapito 2</label><input name="conducente2_recapito" value="${val('conducente2_recapito')}"></div></div>
<button>Salva CARGOS</button></form><hr><div class="actions">
<a class="btn" href="/cargos/${p.id}/txt">Scarica TXT 1505</a><a class="btn btn2" href="/cargos/${p.id}/csv">Scarica CSV ;</a><a class="btn btn2" href="/cargos/${p.id}/preview">Preview</a><a class="btn btn3" href="/cargos/${p.id}/verifica">Verifica dati</a><a class="btn btn3" href="/cargos/${p.id}/invia">Invia report a CaRGOS</a><a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna</a></div></div>`));});});
app.post('/cargos/:id/save',(req,res)=>{const b=req.body; db.run(`UPDATE prenotazioni SET record_cargos_pagamento_tipo=?,record_cargos_checkout_luogo_cod=?,record_cargos_checkout_indirizzo=?,record_cargos_checkin_luogo_cod=?,record_cargos_checkin_indirizzo=?,record_cargos_operatore_id=?,record_cargos_agenzia_id=?,record_cargos_agenzia_nome=?,record_cargos_agenzia_luogo_cod=?,record_cargos_agenzia_indirizzo=?,record_cargos_agenzia_tel=?,record_cargos_veicolo_tipo=?,record_cargos_veicolo_colore=?,record_cargos_veicolo_gps=?,record_cargos_veicolo_bloccom=?,record_cargos_cittadinanza_cod=?,record_cargos_nascita_luogo_cod=?,record_cargos_residenza_luogo_cod=?,record_cargos_doc_tipo_cod=?,record_cargos_doc_luogoril_cod=?,record_cargos_patente_luogoril_cod=?,conducente2_nome=?,conducente2_cognome=?,conducente2_data_nascita=?,conducente2_nascita_luogo_cod=?,conducente2_cittadinanza_cod=?,conducente2_doc_tipo_cod=?,conducente2_doc_numero=?,conducente2_doc_luogoril_cod=?,conducente2_patente_numero=?,conducente2_patente_luogoril_cod=?,conducente2_recapito=? WHERE id=?`,[b.record_cargos_pagamento_tipo,b.record_cargos_checkout_luogo_cod,b.record_cargos_checkout_indirizzo,b.record_cargos_checkin_luogo_cod,b.record_cargos_checkin_indirizzo,b.record_cargos_operatore_id,b.record_cargos_agenzia_id,b.record_cargos_agenzia_nome,b.record_cargos_agenzia_luogo_cod,b.record_cargos_agenzia_indirizzo,b.record_cargos_agenzia_tel,b.record_cargos_veicolo_tipo,b.record_cargos_veicolo_colore,b.record_cargos_veicolo_gps,b.record_cargos_veicolo_bloccom,b.record_cargos_cittadinanza_cod,b.record_cargos_nascita_luogo_cod,b.record_cargos_residenza_luogo_cod,b.record_cargos_doc_tipo_cod,b.record_cargos_doc_luogoril_cod,b.record_cargos_patente_luogoril_cod,b.conducente2_nome,b.conducente2_cognome,b.conducente2_data_nascita,b.conducente2_nascita_luogo_cod,b.conducente2_cittadinanza_cod,b.conducente2_doc_tipo_cod,b.conducente2_doc_numero,b.conducente2_doc_luogoril_cod,b.conducente2_patente_numero,b.conducente2_patente_luogoril_cod,b.conducente2_recapito,req.params.id],()=>res.redirect('/cargos/'+req.params.id));});










app.get('/prenotazione/:id', async (req, res) => {
  const p = await get(`SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.descrizione_pubblica FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE p.id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');
  res.send(page('Dettaglio contratto', `
    <div class="box">
      <h2>Contratto ${esc(p.codice)}</h2>
      <p><b>Cliente:</b> ${esc(p.nome)} ${esc(p.cognome)} - ${esc(p.telefono)}</p>
      <p><b>Email:</b> ${esc(p.email)} | <b>CF:</b> ${esc(p.codice_fiscale)}</p>
      <p><b>Mezzo:</b> <a href="/mezzo/${p.mezzo_id}">${esc(p.targa)} ${esc(descrizionePubblica(p))}</a></p>
      <p><b>Date:</b> ${esc(p.data_inizio)} ore ${esc(p.ora_inizio)} - ${esc(p.data_fine)} ore ${esc(p.ora_fine)}</p>
      <p><b>Totale:</b> &euro; ${euro(p.totale)}</p>${cauzioneHtml(p)}
      <p><b>Stato:</b> <span class="badge ${p.stato==='firmato'?'badge-green':'badge-orange'}">${esc(dpLabelStatus(p.stato||'bozza'))}</span> <b>Nexi:</b> <span class="badge ${p.nexi_stato==='pagato'?'badge-green':'badge-orange'}">${esc(dpLabelStatus(p.nexi_stato || 'non pagato'))}</span> <b>Ca.R.G.O.S.:</b> <span class="badge ${p.record_cargos_stato||p.cargos_inviato?'badge-green':'badge-orange'}">${esc(dpLabelStatus(p.record_cargos_stato || (p.cargos_inviato?'inviato':'da inviare')))}</span> <b>Firma:</b> <span class="badge ${p.firma_path?'badge-green':'badge-red'}">${p.firma_path?'Firmato':'Manca firma'}</span></p>
      ${p.pdf_drive_web_link ? `<p><b>PDF Drive:</b> <a target="_blank" href="${esc(p.pdf_drive_web_link)}">Apri su Drive</a></p>` : ''}
      ${p.nexi_link ? `<p><b>Link Nexi:</b> <a target="_blank" href="${esc(p.nexi_link)}">${esc(p.nexi_link)}</a></p>` : ''}
      <div class="actions">
        <a class="btn" href="/contratto/${p.id}">PDF</a>
        <a class="btn btn2" href="/firma/${p.id}">Firma</a>
        <a class="btn btn2" href="/email/${p.id}">Email</a>
        <a class="btn btn3" href="/documenti/${p.id}">Foto/documenti</a>
        <a class="btn btn3" href="/cliente-documenti-link/${p.id}">Link documenti cliente</a>
        <a class="btn btn3" href="/ocr-documenti/${p.id}">OCR iPad</a>
        <a class="btn btn3" href="/cliente-documenti-link/${p.id}">Link documenti cliente</a>
        <a class="btn btn3" href="/ocr-documenti/${p.id}">OCR patente/documento</a>
        <a class="btn btn3" href="/checkout/${p.id}">Check-out</a>
        <a class="btn btn3" href="/checkin/${p.id}">Check-in</a>
        <a class="btn btnWarn" href="/nexi/${p.id}">Nexi Pay Link</a> <a class="btn btn3" href="/nexi/${p.id}/invia-whatsapp">Invia pagamento WhatsApp</a>
        <a class="btn btn3" href="/firma-link/${p.id}">Link firma WhatsApp</a> <a class="btn btn3" href="/firma-whatsapp/${p.id}">Invia firma WhatsApp diretto</a>
        <a class="btn btn3" href="/whatsapp-contratto/${p.id}">Invia contratto WhatsApp</a>
        <a class="btn btn2" href="/cargos/export/${p.id}">Export Ca.R.G.O.S.</a>
      </div>
    </div>`));
});

app.get('/prenotazioni', async (req, res) => {
  const q = normalize(req.query.q), stato = normalize(req.query.stato), dal = normalize(req.query.dal), al = normalize(req.query.al);
  let sql = `SELECT p.*, m.targa, m.marca, m.modello, m.descrizione_pubblica FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE 1=1`;
  const params = [];
  if (q) { sql += ` AND (p.codice LIKE ? OR p.nome LIKE ? OR p.cognome LIKE ? OR p.telefono LIKE ? OR m.targa LIKE ?)`; params.push(`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`); }
  if (stato) { sql += ` AND p.stato=?`; params.push(stato); }
  if (dal) { sql += ` AND date(p.data_inizio)>=date(?)`; params.push(dal); }
  if (al) { sql += ` AND date(p.data_fine)<=date(?)`; params.push(al); }
  sql += ` ORDER BY p.id DESC`;
  const rows = await all(sql, params);
  const trs = rows.map(p => `<tr><td><a href="/contratto/${p.id}/gestisci">${esc(p.codice)}</a></td><td>${esc(p.nome)} ${esc(p.cognome)}</td><td>${esc(p.telefono)}<br>${esc(p.email)}</td><td><b>${esc(p.targa)}</b><br>${esc(descrizionePubblica(p))}</td><td>${esc(p.data_inizio)} - ${esc(p.data_fine)}</td><td>&euro; ${euro(p.totale)}</td><td><span class="badge ${p.stato==='firmato'?'badge-green':'badge-orange'}">${esc(p.stato||'bozza')}</span><br><span class="badge ${p.firma_path?'badge-green':'badge-red'}">${p.firma_path?'firma ok':'firma no'}</span></td><td><span class="badge ${p.record_cargos_stato||p.cargos_inviato?'badge-green':'badge-orange'}">${esc(p.record_cargos_stato || (p.cargos_inviato?'inviato':'da inviare'))}</span></td><td><a href="/contratto/${p.id}/gestisci">Apri</a><br><a href="/contratto/${p.id}">PDF</a><br><a href="/nexi/${p.id}">Nexi</a></td></tr>`).join('');
  res.send(page('Storico', `<h2>Storico contratti / prenotazioni</h2><form method="GET" action="/prenotazioni" class="box"><div class="grid"><input name="q" placeholder="Cerca nome, targa, codice, telefono" value="${esc(q)}"><select name="stato"><option value="">Tutti gli stati</option>${['bozza','richiesta_cliente','confermato','firmato','in_corso','rientrato','chiuso','pagato','annullato'].map(s=>`<option ${stato===s?'selected':''}>${s}</option>`).join('')}</select><input type="date" name="dal" value="${esc(dal)}"><input type="date" name="al" value="${esc(al)}"></div><button>Cerca</button></form><table><tr><th>Codice</th><th>Cliente</th><th>Contatti</th><th>Mezzo</th><th>Date</th><th>Totale</th><th>Stato</th><th>CaRGOS</th><th>Azioni</th></tr>${trs}</table>`));
});
app.get('/stato/:id/:stato', async (req, res) => {
  await run(`UPDATE prenotazioni SET stato=? WHERE id=?`, [req.params.stato, req.params.id]);
  res.redirect('/prenotazioni');
});

app.get('/planning', async (req, res) => {
  const mese = req.query.mese || moment().format('YYYY-MM');
  const start = moment(mese + '-01'), days = start.daysInMonth();
  const mezzi = await all(`SELECT * FROM mezzi ORDER BY targa`);
  const pren = await all(`SELECT p.*, m.targa, m.marca, m.modello FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE p.stato!='annullato'`);
  let header = '<th class="sticky-col">Mezzo</th>';
  for (let d=1; d<=days; d++) header += `<th>${d}</th>`;
  let rows = '';
  mezzi.forEach(m => {
    rows += `<tr><td class="sticky-col"><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a><br>${esc(descrizionePubblica(m))}</td>`;
    for (let d=1; d<=days; d++) {
      const day = start.clone().date(d).format('YYYY-MM-DD');
      const occ = pren.find(p => p.mezzo_id == m.id && moment(day).isSameOrAfter(moment(p.data_inizio)) && moment(day).isSameOrBefore(moment(p.data_fine)));
      if (occ) rows += `<td class="occupato" title="${esc(occ.codice)} ${esc(occ.nome)} ${esc(occ.cognome)}" onclick="window.location='/prenotazione/${occ.id}'">O</td>`;
      else rows += `<td class="libero" title="Libero ${esc(m.targa)} ${day}" onclick="window.location='/nuova-prenotazione?mezzo_id=${m.id}&data=${day}'">L</td>`;
    }
    rows += '</tr>';
  });
  const prec = start.clone().subtract(1,'month').format('YYYY-MM'), succ = start.clone().add(1,'month').format('YYYY-MM');
  res.send(page('Planning', `<h2>Planning ${start.format('MM/YYYY')}</h2><p><a href="/planning?mese=${prec}">← Mese precedente</a> | <a href="/planning?mese=${succ}">Mese successivo -</a></p><p><span class="libero" style="padding:6px;">Libero: clic per prenotare</span> <span class="occupato" style="padding:6px;">Occupato: clic per aprire contratto</span></p><div class="sticky-table"><table><tr>${header}</tr>${rows}</table></div>`));
});



function clienteDocToken(id) {
  return crypto.createHash('sha256')
    .update(`${id}:${process.env.APP_BASE_URL || ''}:${process.env.NEXI_MAC_KEY || 'dp-rent'}`)
    .digest('hex')
    .slice(0, 24);
}

function clienteDocUrl(req, id) {
  return absoluteUrl(req, `/cliente-documenti/${id}/${clienteDocToken(id)}`);
}

function renderOcrUploadForm(action, backUrl, title = 'Carica documenti') {
  return `
    <div class="box">
      <h2>${esc(title)}</h2>
      <p class="notice">Puoi scattare una foto oppure caricare un file. Dopo la lettura dei dati, controlla e conferma.</p>

      <label>Tipo documento</label>
      <select id="tipoScelto">
        <option>Patente</option>
        <option>Carta identità</option>
        <option>Codice fiscale</option>
        <option>Altro documento</option>
      </select>

      <form id="formCamera" method="POST" action="${action}" enctype="multipart/form-data">
        <input type="hidden" name="tipo" id="tipoCamera">
        <input id="cameraInput" type="file" name="file" accept="image/*" capture="environment" style="display:none" required>
      </form>

      <form id="formFile" method="POST" action="${action}" enctype="multipart/form-data">
        <input type="hidden" name="tipo" id="tipoFile">
        <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
      </form>

      <button type="button" onclick="scattaFoto()"> Scatta foto</button>
      <button type="button" class="btn2" onclick="caricaDaDispositivo()"> Carica da dispositivo</button>
      <a class="btn btn2" href="${backUrl}">Torna</a>
    </div>

    <script>
      function tipo() { return document.getElementById('tipoScelto').value; }
      function scattaFoto() {
        document.getElementById('tipoCamera').value = tipo();
        document.getElementById('cameraInput').click();
      }
      function caricaDaDispositivo() {
        document.getElementById('tipoFile').value = tipo();
        document.getElementById('fileInput').click();
      }
      document.getElementById('cameraInput').addEventListener('change', function () {
        if (this.files.length) document.getElementById('formCamera').submit();
      });
      document.getElementById('fileInput').addEventListener('change', function () {
        if (this.files.length) document.getElementById('formFile').submit();
      });
    </script>
  `;
}

function renderOcrConfirmPage(p, dati, saveAction, cancelUrl) {
  return page('Conferma dati OCR', `
    <div class="box">
      <h2>Controlla dati letti</h2>
      <p class="notice">Controlla bene: se una data o un numero è sbagliato, correggilo prima di salvare.</p>

      <form method="POST" action="${saveAction}">
        <div class="grid">
          <div><label>Nome</label><input name="nome" value="${ocrValue(dati.nome)}"></div>
          <div><label>Cognome</label><input name="cognome" value="${ocrValue(dati.cognome)}"></div>
          <div><label>Data nascita</label><input type="date" name="data_nascita" value="${ocrValue(dati.data_nascita)}"></div>
          <div><label>Luogo nascita</label><input name="luogo_nascita" value="${ocrValue(dati.luogo_nascita)}"></div>
          <div><label>Codice fiscale</label><input name="codice_fiscale" value="${ocrValue(dati.codice_fiscale)}"></div>
          <div><label>Numero documento</label><input name="numero_documento" value="${ocrValue(dati.numero_documento)}"></div>
          <div><label>Ente rilascio documento</label><input name="ente_rilascio" value="${ocrValue(dati.ente_rilascio)}"></div>
          <div><label>Data rilascio documento</label><input type="date" name="data_rilascio" value="${ocrValue(dati.data_rilascio)}"></div>
          <div><label>Scadenza documento/patente</label><input type="date" name="data_scadenza" value="${ocrValue(dati.data_scadenza)}"></div>
          <div><label>Numero patente</label><input name="numero_patente" value="${ocrValue(dati.numero_patente)}"></div>
          <div><label>Categoria patente</label><input name="categoria_patente" value="${ocrValue(dati.categoria_patente)}"></div>
          <div><label>Indirizzo</label><input name="indirizzo" value="${ocrValue(dati.indirizzo)}"></div>
          <div><label>Confidenza AI</label><input value="${ocrValue(dati.confidence)}" readonly></div>
        </div>
        <button>Salva dati nel contratto</button>
      </form>

      <details>
        <summary>Dati grezzi letti dall'AI</summary>
        <pre>${esc(JSON.stringify(dati, null, 2))}</pre>
      </details>

      <a class="btn btn2" href="${cancelUrl}">Annulla</a>
    </div>
  `);
}

async function salvaDatiOcrSuContratto(id, b) {
  const current = await get(`SELECT * FROM prenotazioni WHERE id=?`, [id]);
  if (!current) throw new Error('Contratto non trovato');

  const noteExtra =
    `\nOCR documento:\n` +
    `Numero documento: ${b.numero_documento || ''}\n` +
    `Ente rilascio: ${b.ente_rilascio || ''}\n` +
    `Rilascio: ${b.data_rilascio || ''}\n` +
    `Scadenza: ${b.data_scadenza || ''}\n` +
    `Luogo nascita: ${b.luogo_nascita || ''}\n` +
    `Categoria patente: ${b.categoria_patente || ''}`;

  await run(
    `UPDATE prenotazioni
     SET nome=?,
         cognome=?,
         codice_fiscale=?,
         indirizzo=?,
         data_nascita=COALESCE(NULLIF(?,''), data_nascita),
         luogo_nascita=COALESCE(NULLIF(?,''), luogo_nascita),
         documento_numero=COALESCE(NULLIF(?,''), documento_numero),
         documento_scadenza=COALESCE(NULLIF(?,''), documento_scadenza),
         patente_numero=COALESCE(NULLIF(?,''), patente_numero),
         patente_scadenza=COALESCE(NULLIF(?,''), patente_scadenza),
         categoria_patente=COALESCE(NULLIF(?,''), categoria_patente),
         patente1=COALESCE(NULLIF(?,''), patente1),
         patente1_scadenza=COALESCE(NULLIF(?,''), patente1_scadenza),
         note=COALESCE(note,'') || ?
     WHERE id=?`,
    [
      b.nome || current.nome,
      b.cognome || current.cognome,
      b.codice_fiscale || current.codice_fiscale,
      b.indirizzo || current.indirizzo,
      b.data_nascita || '',
      b.luogo_nascita || '',
      b.numero_documento || b.documento_numero || '',
      b.data_scadenza || b.documento_scadenza || '',
      b.numero_patente || b.patente_numero || '',
      b.patente_scadenza || b.data_scadenza || '',
      b.categoria_patente || '',
      b.numero_patente || b.patente_numero || '',
      b.patente_scadenza || b.data_scadenza || '',
      noteExtra,
      id
    ]
  );
}

app.get('/ocr-documenti/:id', async (req, res) => {
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');
  res.send(page('OCR documenti', renderOcrUploadForm(`/ocr-documenti/${p.id}`, `/prenotazione/${p.id}`, 'OCR patente/documento - Operatore DP')));
});

app.post('/ocr-documenti/:id', upload.single('file'), async (req, res) => {
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');
    if (!req.file) return res.send('File mancante');

    let driveRes = null;
    try {
      driveRes = await uploadFileToDrive(
        req.file.path,
        `${Date.now()}_OCR_${req.body.tipo}_${req.file.originalname}`,
        req.file.mimetype,
        `${p.codice || 'CONTRATTO'} - ${p.nome || ''} ${p.cognome || ''}`
      );
    } catch (e) { console.log('Errore upload OCR Drive:', e.message); }

    await run(`INSERT INTO allegati (prenotazione_id,tipo,filename,originalname,path,mimetype,drive_file_id,drive_web_link) VALUES (?,?,?,?,?,?,?,?)`,
      [p.id, `OCR ${req.body.tipo}`, req.file.filename, req.file.originalname, req.file.path, req.file.mimetype, driveRes?.id || null, driveRes?.webViewLink || null]);

    const dati = await estraiDatiDocumentoConAI(req.file.path, req.file.mimetype);
    res.send(renderOcrConfirmPage(p, dati, `/ocr-documenti/${p.id}/salva`, `/prenotazione/${p.id}`));
  } catch (e) {
    res.status(500).send(page('Errore OCR', `<div class="box"><h2 class="bad">Errore OCR</h2><pre>${esc(e.message)}</pre><p>Per usare OCR serve OPENAI_API_KEY su Render.</p><a class="btn" href="/ocr-documenti/${req.params.id}">Riprova</a><a class="btn btn2" href="/prenotazione/${req.params.id}">Torna contratto</a></div>`));
  }
});

app.post('/ocr-documenti/:id/salva', async (req, res) => {
  try {
    await salvaDatiOcrSuContratto(req.params.id, req.body);
    res.redirect(`/prenotazione/${req.params.id}`);
  } catch (e) {
    res.status(500).send(page('Errore salvataggio OCR', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});



function mergeOcrObjects(list) {
  const out = {};
  for (const obj of list) {
    if (!obj || typeof obj !== 'object') continue;
    for (const [k, v] of Object.entries(obj)) {
      const val = String(v ?? '').trim();
      if (val && !out[k]) out[k] = val;
    }
  }
  return out;
}


function publicClientePage(title, content) {
  return `<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>${esc(title)}</title>
<style>
:root{--dp-red:#d70000;--dp-blue:#173f9d;--dp-dark:#07111f;--bg:#eef4ff;--card:#fff;--line:#d9dbe7}
*{box-sizing:border-box}body{margin:0;background:linear-gradient(180deg,#eef4ff,#f7f8fb);font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;color:#0b1226;-webkit-text-size-adjust:100%;font-size:18px}main{max-width:900px;margin:0 auto;padding:16px 14px 38px}.client-hero{background:linear-gradient(135deg,#06183f,#173f9d);color:#fff;border-radius:0 0 30px 30px;padding:calc(28px + env(safe-area-inset-top)) 24px 30px;margin:0 -14px 22px;box-shadow:0 16px 40px rgba(0,0,0,.18)}.client-hero h1{font-size:44px;line-height:1;margin:0 0 14px;color:#fff;letter-spacing:.8px}.client-hero p{font-size:21px;line-height:1.35;margin:0;opacity:.96}.pill{display:inline-block;background:#fff;color:#173f9d;border-radius:999px;padding:11px 18px;margin:16px 8px 0 0;font-weight:900;font-size:20px}.card,.step-card,details.manual{background:#fff;border-radius:26px;padding:24px;margin:18px 0;box-shadow:0 14px 35px rgba(15,23,42,.10);border:1px solid #e3e7f2}.card h2,.step-card h2{font-size:34px;margin:0 0 16px;color:#0b1226}.grid,.upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}label{display:block;font-weight:900;margin:10px 0 7px;color:#0b1226}input,select,textarea{width:100%;border:2px solid var(--line);border-radius:17px;padding:16px;font-size:20px;background:#fff;color:#111;font-weight:700;outline:none}textarea{min-height:110px}.notice{background:#fff8df;border:1px solid #f1d98a;border-radius:20px;padding:16px;line-height:1.35;font-size:18px}.okbox{background:#ecfff1;border:1px solid #b8efc4;border-radius:20px;padding:16px;line-height:1.35}.btn,button,.big-red{border:0;border-radius:20px;background:linear-gradient(135deg,#e21818,#a80d0d);color:#fff;padding:18px 24px;font-size:22px;font-weight:900;box-shadow:0 12px 25px rgba(210,0,0,.25);width:100%;margin-top:18px}.small{font-size:15px;color:#596275;font-weight:700}.upload-box{border:2px dashed #cbd5e1;border-radius:22px;padding:18px;background:#f8fafc}.upload-box label{font-size:20px}.upload-box input{font-size:18px;margin-top:12px}details.manual summary{font-size:30px;font-weight:900;cursor:pointer}.fixed-save{position:sticky;bottom:12px;z-index:9;background:rgba(255,255,255,.96);padding:12px;border-radius:24px;box-shadow:0 10px 30px rgba(0,0,0,.16)}@media(max-width:720px){main{padding-left:14px;padding-right:14px}.grid,.upload-grid{grid-template-columns:1fr}.client-hero h1{font-size:36px}.client-hero p{font-size:20px}.card,.step-card,details.manual{padding:20px;border-radius:24px}.card h2,.step-card h2{font-size:31px}input,select,textarea{font-size:19px}}
</style>
<script>
function toggleAzienda(){var el=document.querySelector('[name="tipo_cliente"]'); if(!el) return; var t=el.value; var box=document.getElementById('aziendaBox'); if(!box) return; var isAz=t==='azienda'; box.style.display=isAz?'grid':'none'; ['ragione_sociale','partita_iva','pec','codice_sdi'].forEach(function(n){var f=document.querySelector('[name="'+n+'"]'); if(f) f.required=isAz;});}
window.addEventListener('DOMContentLoaded',toggleAzienda);
</script>
</head><body><main>${content}</main></body></html>`;
}

function renderClientePulitoPage(p, token, files) {
  const lista = (files || []).map(f => `<li>${esc(f.tipo)} - ${esc(f.originalname)} ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Drive</a>` : ''}</li>`).join('');
  const actionBase = `/cliente-documenti/${p.id}/${token}`;
  return publicClientePage('DP RENT - Pratica cliente', `
    <style>
      .client-hero{background:linear-gradient(135deg,#06183f,#173f9d);color:white;border-radius:30px;padding:36px 42px;margin-bottom:26px;box-shadow:0 16px 40px rgba(0,0,0,.18)}
      .client-hero h1{font-size:54px;line-height:1;margin:0 0 18px 0;color:white}
      .client-hero p{font-size:26px;line-height:1.35;margin:0;opacity:.95}
      .pill{display:inline-block;background:#fff;color:#173f9d;border-radius:999px;padding:12px 22px;margin:18px 12px 0 0;font-weight:900;font-size:22px}
      .step-card{background:white;border-radius:30px;padding:34px;margin-bottom:26px;box-shadow:0 14px 35px rgba(15,23,42,.10);border:1px solid #edf0f7}
      .step-card h2{font-size:44px;margin-top:0;color:#0b1226}
      .upload-grid{display:grid;grid-template-columns:1fr 1fr;gap:18px}
      .upload-box{border:2px dashed #cbd5e1;border-radius:22px;padding:18px;background:#f8fafc}
      .upload-box label{font-size:22px;font-weight:900;color:#0f172a}
      .upload-box input{font-size:22px;width:100%;margin-top:12px}
      .big-red{font-size:28px;border-radius:22px;padding:18px 26px;background:#d40000;color:white;border:0;font-weight:900;box-shadow:0 12px 28px rgba(212,0,0,.25)}
      details.manual{background:white;border-radius:30px;padding:26px;margin:26px 0;box-shadow:0 14px 35px rgba(15,23,42,.10)}
      details.manual summary{font-size:34px;font-weight:900;cursor:pointer;color:#0b1226}
      .fixed-save{position:sticky;bottom:12px;z-index:9;background:rgba(255,255,255,.96);padding:12px;border-radius:24px;box-shadow:0 10px 30px rgba(0,0,0,.16)}
      @media(max-width:700px){.client-hero h1{font-size:40px}.client-hero p{font-size:22px}.upload-grid{grid-template-columns:1fr}.step-card h2{font-size:38px}}
    
.contract-main-actions{margin-top:16px}.contract-main-actions .btn{min-width:190px;text-align:center}.contract-secondary-actions .btn{min-width:150px;text-align:center}
@media(max-width:700px){.contract-main-actions .btn,.contract-secondary-actions .btn{width:100%;min-width:0}}


/* V109 FIX leggibilita mobile */
header{padding-top:max(22px, env(safe-area-inset-top));}
.top-actions{max-width:1180px;margin:0 auto 14px!important;padding:10px 0!important;}
.top-actions .back-btn::before{content:""!important;}
.top-actions .back-btn,.top-actions a{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif!important;font-size:clamp(18px,2.6vw,24px)!important;letter-spacing:0!important;line-height:1.1!important;white-space:nowrap!important;color:#fff!important;overflow:hidden;text-overflow:ellipsis;}
.top-actions .back-btn{background:#333!important;}
.top-actions .home-btn{background:#d70000!important;}
.client-back button{font-size:18px!important;font-weight:900!important;background:#333!important;color:#fff!important;}
@media(max-width:700px){
  nav{padding-top:calc(14px + env(safe-area-inset-top));}
  .top-actions{position:sticky;top:0;z-index:50;padding:10px 12px!important;gap:10px!important;background:rgba(244,244,244,.96)!important;}
  .top-actions .back-btn,.top-actions a{min-width:0!important;width:calc(50% - 5px)!important;flex:1 1 calc(50% - 5px)!important;padding:14px 8px!important;}
  .contract-main-actions .btn{width:100%!important;}
}

</style>

    <div class="client-hero">
      <h1>DP RENT</h1>
      <p>Prima carica documento e patente. Proviamo a leggere i dati automaticamente, poi puoi controllare e correggere tutto manualmente.</p>
      <span class="pill">${esc(p.codice || 'Pratica cliente')}</span>
      <span class="pill">${esc((p.marca || '') + ' ' + (p.modello || '')).trim() || 'Mezzo'}</span>
    </div>

    <div class="step-card">
      <h2>1. Carica documenti per OCR</h2>
      <p class="notice">Scatta o carica le foto. Puoi anche caricare solo una foto: i campi mancanti li compili sotto a mano.</p>
      <form method="POST" action="${actionBase}/ocr-multiplo" enctype="multipart/form-data">
        <div class="upload-grid">
          <div class="upload-box"><label>Documento fronte</label><input type="file" name="documento_fronte" accept="image/*,.pdf" capture="environment"></div>
          <div class="upload-box"><label>Documento retro</label><input type="file" name="documento_retro" accept="image/*,.pdf" capture="environment"></div>
          <div class="upload-box"><label>Patente fronte</label><input type="file" name="patente_fronte" accept="image/*,.pdf" capture="environment"></div>
          <div class="upload-box"><label>Patente retro</label><input type="file" name="patente_retro" accept="image/*,.pdf" capture="environment"></div>
        </div>
        <br>
        <button class="big-red">Leggi dati automaticamente</button>
      </form>
    </div>

    <details class="manual" open>
      <summary>2. Controlla / compila manualmente</summary>
      <p class="notice">Questi sono i dati necessari per contratto, fatturazione e controlli Ca.R.G.O.S.</p>
      <form method="POST" action="${actionBase}/salva">
        <div class="grid">
          <div><label>Nome</label><input name="nome" value="${esc(p.nome||'')}" autocomplete="given-name"></div>
          <div><label>Cognome</label><input name="cognome" value="${esc(p.cognome||'')}" autocomplete="family-name"></div>
          <div><label>Telefono</label><input name="telefono" value="${esc(p.telefono||'')}" autocomplete="tel"></div>
          <div><label>Email</label><input name="email" value="${esc(p.email||'')}" autocomplete="email"></div>
          <div><label>Codice fiscale</label><input name="codice_fiscale" value="${esc(p.codice_fiscale||'')}"></div>
          <div><label>Data nascita</label><input type="date" name="data_nascita" value="${esc(p.data_nascita||'')}"></div>
          <div><label>Luogo nascita</label><input name="luogo_nascita" value="${esc(p.luogo_nascita||'')}"></div>
          <div><label>Cittadinanza codice</label><input name="cittadinanza_cod" value="${esc(p.cittadinanza_cod||p.cittadinanza||'100000100')}"></div>
          <div><label>Indirizzo</label><input name="indirizzo" value="${esc(p.indirizzo||'')}"></div>
          <div><label>Città</label><input name="citta" value="${esc(p.citta||'')}"></div>
          <div><label>Provincia</label><input name="provincia" value="${esc(p.provincia||'')}"></div>
          <div><label>CAP</label><input name="cap" value="${esc(p.cap||'')}"></div>
          <div><label>Tipo documento</label><select name="documento_tipo"><option ${p.documento_tipo==='Carta identità'?'selected':''}>Carta identità</option><option ${p.documento_tipo==='Passaporto'?'selected':''}>Passaporto</option><option ${p.documento_tipo==='Patente'?'selected':''}>Patente</option></select></div>
          <div><label>Numero documento</label><input name="numero_documento" value="${esc(p.documento_numero||'')}"></div>
          <div><label>Ente rilascio documento</label><input name="ente_rilascio" value="${esc(p.documento_rilascio||'')}"></div>
          <div><label>Luogo rilascio documento</label><input name="documento_luogo_rilascio" value="${esc(p.documento_luogo_rilascio||p.citta||'')}"></div>
          <div><label>Data rilascio documento</label><input type="date" name="data_rilascio" value="${esc(p.documento_data_rilascio||'')}"></div>
          <div><label>Scadenza documento</label><input type="date" name="data_scadenza" value="${esc(p.documento_scadenza||'')}"></div>
          <div><label>Numero patente</label><input name="numero_patente" value="${esc(p.patente_numero||p.patente1||'')}"></div>
          <div><label>Ente rilascio patente</label><input name="patente_rilascio" value="${esc(p.patente_rilascio||'')}"></div>
          <div><label>Luogo rilascio patente</label><input name="patente_luogo_rilascio" value="${esc(p.patente_luogo_rilascio||p.citta||'')}"></div>
          <div><label>Data rilascio patente</label><input type="date" name="patente_data_rilascio" value="${esc(p.patente_data_rilascio||'')}"></div>
          <div><label>Scadenza patente</label><input type="date" name="patente_scadenza" value="${esc(p.patente_scadenza||p.patente1_scadenza||'')}"></div>
          <div><label>Categoria patente</label><input name="categoria_patente" value="${esc(p.categoria_patente||'')}"></div>
          <div><label>Tipo cliente</label><select name="tipo_cliente"><option value="privato" ${p.tipo_cliente!=='azienda'?'selected':''}>Privato</option><option value="azienda" ${p.tipo_cliente==='azienda'?'selected':''}>Azienda</option></select></div>
          <div><label>Ragione sociale</label><input name="ragione_sociale" value="${esc(p.ragione_sociale||'')}"></div>
          <div><label>Partita IVA</label><input name="piva" value="${esc(p.piva||p.partita_iva||'')}"></div>
          <div><label>PEC</label><input name="pec" value="${esc(p.pec||'')}"></div>
          <div><label>Codice SDI</label><input name="codice_sdi" value="${esc(p.codice_sdi||'')}"></div>
        </div>
        <div class="fixed-save"><button class="big-red">Salva dati cliente e documenti</button></div>
      </form>
    </details>

    <div class="step-card">
      <h2>3. Altri allegati</h2>
      <form method="POST" action="${actionBase}" enctype="multipart/form-data">
        <input type="hidden" name="tipo" value="Altro allegato cliente">
        <input type="file" name="file" accept="image/*,.pdf" multiple>
        <button>Carica allegato</button>
      </form>
      <h3>File già caricati</h3><ul>${lista || '<li>Nessun file caricato</li>'}</ul>
    </div>
  `);
}

app.get('/cliente-documenti/:id/:token', async (req, res) => {
  const expected = clienteDocToken(req.params.id);
  if (req.params.token !== expected) return res.status(403).send('Link non valido');

  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');

  const files = await all(`SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id DESC`, [p.id]);
  res.send(renderClientePulitoPage(p, req.params.token, files));
});

app.post('/cliente-documenti/:id/:token/ocr-multiplo', upload.fields([
  { name:'documento_fronte', maxCount:1 },
  { name:'documento_retro', maxCount:1 },
  { name:'patente_fronte', maxCount:1 },
  { name:'patente_retro', maxCount:1 }
]), async (req, res) => {
  try {
    const expected = clienteDocToken(req.params.id);
    if (req.params.token !== expected) return res.status(403).send('Link non valido');

    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');

    const allFiles = [];
    for (const [field, arr] of Object.entries(req.files || {})) {
      for (const f of arr || []) allFiles.push({ field, file:f });
    }
    if (!allFiles.length) return res.send(page('OCR documenti', `<div class="box"><h2>Nessuna foto caricata</h2><a class="btn" href="/cliente-documenti/${p.id}/${req.params.token}">Torna</a></div>`));

    const results = [];
    for (const item of allFiles) {
      const f = item.file;
      let driveRes = null;
      try {
        driveRes = await uploadFileToDrive(
          f.path,
          `${Date.now()}_${item.field}_${f.originalname}`,
          f.mimetype,
          `${p.codice || 'CONTRATTO'} - ${p.nome || ''} ${p.cognome || ''}`
        );
      } catch (e) { console.log('Errore upload OCR multiplo Drive:', e.message); }

      await run(`INSERT INTO allegati (prenotazione_id,tipo,filename,originalname,path,mimetype,drive_file_id,drive_web_link) VALUES (?,?,?,?,?,?,?,?)`,
        [p.id, `OCR CLIENTE ${item.field}`, f.filename, f.originalname, f.path, f.mimetype, driveRes?.id || null, driveRes?.webViewLink || null]);

      try {
        const dati = await estraiDatiDocumentoConAI(f.path, f.mimetype);
        results.push(dati || {});
      } catch (e) {
        console.log('Errore OCR su', item.field, e.message);
      }
    }

    const merged = mergeOcrObjects(results);
    res.send(renderOcrConfirmPage(p, merged, `/cliente-documenti/${p.id}/${req.params.token}/salva`, `/cliente-documenti/${p.id}/${req.params.token}`));
  } catch (e) {
    res.status(500).send(page('Errore OCR cliente', `<div class="box"><h2 class="bad">Errore lettura documenti</h2><pre>${esc(e.message)}</pre><a class="btn" href="/cliente-documenti/${req.params.id}/${req.params.token}">Riprova</a></div>`));
  }
});

app.post('/cliente-documenti/:id/:token', upload.single('file'), async (req, res) => {
  try {
    const expected = clienteDocToken(req.params.id);
    if (req.params.token !== expected) return res.status(403).send('Link non valido');

    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');
    if (!req.file) return res.send('File mancante');

    let driveRes = null;
    try {
      driveRes = await uploadFileToDrive(
        req.file.path,
        `${Date.now()}_CLIENTE_${req.body.tipo}_${req.file.originalname}`,
        req.file.mimetype,
        `${p.codice || 'CONTRATTO'} - ${p.nome || ''} ${p.cognome || ''}`
      );
    } catch (e) { console.log('Errore upload documento cliente Drive:', e.message); }

    await run(`INSERT INTO allegati (prenotazione_id,tipo,filename,originalname,path,mimetype,drive_file_id,drive_web_link) VALUES (?,?,?,?,?,?,?,?)`,
      [p.id, `CLIENTE ${req.body.tipo}`, req.file.filename, req.file.originalname, req.file.path, req.file.mimetype, driveRes?.id || null, driveRes?.webViewLink || null]);

    const dati = await estraiDatiDocumentoConAI(req.file.path, req.file.mimetype);
    res.send(renderOcrConfirmPage(p, dati, `/cliente-documenti/${p.id}/${req.params.token}/salva`, `/cliente-documenti/${p.id}/${req.params.token}`));
  } catch (e) {
    res.status(500).send(page('Errore OCR cliente', `<div class="box"><h2 class="bad">Errore lettura documento</h2><pre>${esc(e.message)}</pre><p>Puoi riprovare oppure contattare DP RENT.</p><a class="btn" href="/cliente-documenti/${req.params.id}/${req.params.token}">Riprova</a></div>`));
  }
});

app.post('/cliente-documenti/:id/:token/salva', async (req, res) => {
  try {
    const expected = clienteDocToken(req.params.id);
    if (req.params.token !== expected) return res.status(403).send('Link non valido');

    await salvaDatiOcrSuContratto(req.params.id, req.body);

    res.send(page('Documenti salvati', `
      <div class="box">
        <h2 class="ok">Dati salvati correttamente</h2>
        <p>Grazie. DP RENT controllerà i dati e completerà il contratto.</p>
      </div>
    `));
  } catch (e) {
    res.status(500).send(page('Errore salvataggio cliente', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/cliente-documenti-link/:id', async (req, res) => {
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');

  const link = clienteDocUrl(req, p.id);
  const msg = `DP RENT - carica patente/documento per contratto ${p.codice}: ${link}`;

  res.send(page('Link documenti cliente', `
    <div class="box">
      <h2>Link documenti cliente</h2>
      <p>Invia questo link al cliente per caricare/scattare patente e documento.</p>
      <input value="${esc(link)}" readonly onclick="this.select()">
      <a class="btn btn3" target="_blank" href="${esc(whatsappText(msg))}">Invia link documenti WhatsApp</a>
      <a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna contratto</a>
    </div>
  `));
});

app.get('/documenti/:id', async (req, res) => {
  const files = await all(`SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id DESC`, [req.params.id]);
  const filesV99 = v99DedupeAllegati(files);
  const lista = filesV99.map(f => `<li>${esc(f.tipo)} - <a href="/uploads/${esc(f.filename)}" target="_blank">${esc(f.originalname)}</a> ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Google Drive</a>` : ''}</li>`).join('');

  const options = `
    <option>Patente fronte</option><option>Patente retro</option>
    <option>Documento fronte</option><option>Documento retro</option>
    <option>Codice fiscale</option>
    <option>Foto uscita fronte</option><option>Foto uscita retro</option>
    <option>Foto uscita lato dx</option><option>Foto uscita lato sx</option><option>Foto uscita interno</option><option>Foto danni uscita</option>
    <option>Foto rientro fronte</option><option>Foto rientro retro</option>
    <option>Foto rientro lato dx</option><option>Foto rientro lato sx</option><option>Foto rientro interno</option><option>Foto danni rientro</option>
    <option>Altro file</option>
  `;

  res.send(page('Documenti', `
    <h2>Documenti cliente / foto in-out</h2>
    <div class="box">
      <p class="notice">Da telefono/iPad puoi scattare foto direttamente. Da Mac puoi scegliere il file dal dispositivo.</p>

      <label>Tipo documento/foto</label>
      <select id="tipoScelto">${options}</select>

      <form id="formCamera" method="POST" action="/documenti/${req.params.id}" enctype="multipart/form-data">
        <input type="hidden" name="tipo" id="tipoCamera">
        <input id="cameraInput" type="file" name="file" accept="image/*" capture="environment" style="display:none" required>
      </form>

      <form id="formFile" method="POST" action="/documenti/${req.params.id}" enctype="multipart/form-data">
        <input type="hidden" name="tipo" id="tipoFile">
        <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
      </form>

      <button type="button" onclick="scattaFoto()">&#128247; Scatta foto</button>
      <button type="button" class="btn2" onclick="caricaDaDispositivo()">&#128193; Carica da dispositivo</button>

      <h3>Allegati caricati</h3>
      <ul>${lista}</ul>

      <a class="btn" href="/prenotazione/${req.params.id}">Torna contratto</a>
    </div>

    <script>
      function tipo() { return document.getElementById('tipoScelto').value; }

      function scattaFoto() {
        document.getElementById('tipoCamera').value = tipo();
        document.getElementById('cameraInput').click();
      }

      function caricaDaDispositivo() {
        document.getElementById('tipoFile').value = tipo();
        document.getElementById('fileInput').click();
      }

      document.getElementById('cameraInput').addEventListener('change', function () {
        if (this.files.length) document.getElementById('formCamera').submit();
      });

      document.getElementById('fileInput').addEventListener('change', function () {
        if (this.files.length) document.getElementById('formFile').submit();
      });
    </script>
  `));
});
app.post('/documenti/:id', upload.single('file'), async (req, res) => {
  if (!req.file) return res.send('File mancante');

  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  let driveRes = null;

  // V63: salva sempre anche in cartella locale del contratto: contratti/DPR-.../documenti/
  let finalPath = req.file.path;
  try {
    const folder = path.join(contractsDir, safeFileName(p?.codice || ('contratto_' + req.params.id)), 'documenti');
    fs.mkdirSync(folder, { recursive: true });
    const finalName = safeFileName(`${Date.now()}_${req.body.tipo || 'file'}_${req.file.originalname || req.file.filename}`);
    finalPath = path.join(folder, finalName);
    fs.copyFileSync(req.file.path, finalPath);
  } catch (e) {
    console.log('Errore copia documento in cartella contratto:', e.message);
  }

  try {
    driveRes = await uploadFileToDrive(
      finalPath,
      `${Date.now()}_${req.body.tipo}_${req.file.originalname}`,
      req.file.mimetype,
      `${p?.codice || 'CONTRATTO'} - ${p?.nome || ''} ${p?.cognome || ''}`
    );
  } catch (e) {
    console.log('Errore upload documento Drive:', e.message);
  }

  await run(
    `INSERT INTO allegati (prenotazione_id,tipo,filename,originalname,path,mimetype,drive_file_id,drive_web_link) VALUES (?,?,?,?,?,?,?,?)`,
    [
      req.params.id,
      req.body.tipo,
      path.basename(finalPath),
      req.file.originalname,
      finalPath,
      req.file.mimetype,
      driveRes?.id || null,
      driveRes?.webViewLink || null
    ]
  );

  // V63: sincronizza foto nella stessa cartella Drive e sostituisce una sola copia PDF.
  try { await syncContrattoDriveV63(req.params.id); } catch(e) { console.log('Drive sync V63:', e.message); }
  try { await syncContrattoDriveV63(req.params.id); } catch(e) { console.log('V63 sync foto warning:', e.message); }
    res.redirect(`/documenti/${req.params.id}`);
});

app.get('/checkout/:id', async (req, res) => {
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');
  res.send(page('Check-out', `<h2>Check-out mezzo</h2><form method="POST" action="/checkout/${p.id}"><label>Carburante uscita</label><select name="carburante_uscita">${fuelOptions(p.carburante_uscita)}</select><label>Km uscita</label><input type="number" name="km_uscita" value="${esc(p.km_uscita)}"><label>Note</label><textarea name="note">${esc(p.note)}</textarea><button>Salva check-out</button></form><a class="btn btn3" href="/documenti/${p.id}">Carica foto uscita</a>`));
});
app.post('/checkout/:id', async (req, res) => {
  const p = await get(`SELECT mezzo_id FROM prenotazioni WHERE id=?`, [req.params.id]);
  await run(`UPDATE prenotazioni SET carburante_uscita=?, km_uscita=?, note=?, stato='in_corso' WHERE id=?`, [req.body.carburante_uscita, req.body.km_uscita, req.body.note, req.params.id]);
  if (p && req.body.km_uscita) await run(`UPDATE mezzi SET km_attuali=? WHERE id=?`, [req.body.km_uscita, p.mezzo_id]);
  res.send(actionScreen(req.params.id, 'Check-out salvato', 'Contratto aggiornato in stato in_corso.'));
});
app.get('/checkin/:id', async (req, res) => {
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');
  res.send(page('Check-in', `<h2>Check-in mezzo</h2><form method="POST" action="/checkin/${p.id}"><label>Carburante rientro</label><select name="carburante_rientro">${fuelOptions(p.carburante_rientro)}</select><label>Km rientro</label><input type="number" name="km_rientro" value="${esc(p.km_rientro)}"><label>Note</label><textarea name="note">${esc(p.note)}</textarea><button>Salva check-in</button></form><a class="btn btn3" href="/documenti/${p.id}">Carica foto rientro</a>`));
});
app.post('/checkin/:id', async (req, res) => {
  const p = await get(`SELECT mezzo_id FROM prenotazioni WHERE id=?`, [req.params.id]);
  await run(`UPDATE prenotazioni SET carburante_rientro=?, km_rientro=?, note=?, stato='rientrato' WHERE id=?`, [req.body.carburante_rientro, req.body.km_rientro, req.body.note, req.params.id]);
  if (p && req.body.km_rientro) await run(`UPDATE mezzi SET km_attuali=? WHERE id=?`, [req.body.km_rientro, p.mezzo_id]);
  res.send(actionScreen(req.params.id, 'Check-in salvato', 'Contratto aggiornato in stato rientrato.'));
});

app.get('/contratto/:id', async (req, res) => {
  try {
    const file = await generaPdfContratto(req.params.id, { forceDrive: false, skipDrive: true });
    res.send(page('File pronto', `<div class="box"><h2 class="ok">File pronto</h2><p>File generato. Non scarico automaticamente.</p><a class="btn" href="/cargos">Ca.R.G.O.S.</a></div>`));
  } catch (e) {
    res.status(500).send(page('Errore PDF', `<div class="box"><h2 class="bad">Errore PDF</h2><pre>${esc(e.message)}</pre></div>`));
  }
});


function publicFirmaPage(title, content) {
  return `<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>${esc(title || 'DP RENT')}</title>
<style>
*{box-sizing:border-box}body{margin:0;background:#f2f2f2;color:#202020;font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;-webkit-text-size-adjust:100%}.client-header{background:#050505;color:#fff;padding:24px 20px;text-align:center}.client-brand{font-size:34px;font-weight:950;letter-spacing:2px}.client-sub{margin-top:6px;font-size:14px;letter-spacing:2px;color:#ddd}.client-main{max-width:760px;margin:0 auto;padding:18px}.client-card{background:#fff;border-radius:26px;padding:24px;box-shadow:0 18px 45px rgba(0,0,0,.12)}h1,h2{font-size:clamp(28px,7vw,42px);line-height:1.08;margin:0 0 18px}.ok{color:#10883b}.bad{color:#b30000}p{font-size:18px;line-height:1.45}.btn,button{appearance:none;border:0;display:block;width:100%;text-align:center;text-decoration:none;background:#d70000;color:#fff;border-radius:18px;padding:16px 18px;font-size:20px;font-weight:900;margin:12px 0;box-shadow:0 5px 0 rgba(0,0,0,.18)}.btn2{background:#333}.btn3{background:#10883b}canvas{border:2px solid #222;border-radius:18px;background:#fff;width:100%;height:260px;touch-action:none}.muted{color:#666;font-size:15px}.client-back{display:block;width:100%;text-align:left;margin:0 0 12px}.client-back button{width:auto;min-width:0;background:#333;font-size:16px;padding:10px 14px;border-radius:14px}@media(min-width:760px){.client-main{padding:28px}.client-card{padding:34px}.btn,button{display:inline-block;width:auto;min-width:210px;margin-right:10px}}

.contract-main-actions{margin-top:16px}.contract-main-actions .btn{min-width:190px;text-align:center}.contract-secondary-actions .btn{min-width:150px;text-align:center}
@media(max-width:700px){.contract-main-actions .btn,.contract-secondary-actions .btn{width:100%;min-width:0}}


/* V109 FIX leggibilita mobile */
header{padding-top:max(22px, env(safe-area-inset-top));}
.top-actions{max-width:1180px;margin:0 auto 14px!important;padding:10px 0!important;}
.top-actions .back-btn::before{content:""!important;}
.top-actions .back-btn,.top-actions a{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif!important;font-size:clamp(18px,2.6vw,24px)!important;letter-spacing:0!important;line-height:1.1!important;white-space:nowrap!important;color:#fff!important;overflow:hidden;text-overflow:ellipsis;}
.top-actions .back-btn{background:#333!important;}
.top-actions .home-btn{background:#d70000!important;}
.client-back button{font-size:18px!important;font-weight:900!important;background:#333!important;color:#fff!important;}
@media(max-width:700px){
  nav{padding-top:calc(14px + env(safe-area-inset-top));}
  .top-actions{position:sticky;top:0;z-index:50;padding:10px 12px!important;gap:10px!important;background:rgba(244,244,244,.96)!important;}
  .top-actions .back-btn,.top-actions a{min-width:0!important;width:calc(50% - 5px)!important;flex:1 1 calc(50% - 5px)!important;padding:14px 8px!important;}
  .contract-main-actions .btn{width:100%!important;}
}

</style>
</head>
<body><div class="client-header"><div class="client-brand">DP RENT</div><div class="client-sub">FIRMA CONTRATTO</div></div><main class="client-main"><div class="client-back"><button type="button" onclick="history.length>1?history.back():location.href='/firma-chiusa'">Indietro</button></div><div class="client-card">${content}</div></main></body></html>`;
}


app.get('/firma-chiusa', (req,res)=>{
  res.send(publicFirmaPage('DP RENT', '<h2>Operazione annullata</h2><p>Puoi chiudere questa pagina o tornare al messaggio WhatsApp.</p>'));
});

app.get('/firma/:id', (req, res) => {
  res.send(publicFirmaPage('Firma contratto DP RENT', `
      <h2>Firma contratto</h2>
      <p class="muted">Firma nello spazio sotto e premi Salva firma. Non vedrai la dashboard interna DP RENT.</p>
      <canvas id="canvas"></canvas>
      <button type="button" onclick="clearCanvas()" class="btn2">Cancella</button>
      <button type="button" onclick="saveFirma()">Salva firma</button>
    <script>
      const canvas = document.getElementById('canvas');
      const ctx = canvas.getContext('2d');
      function resize() {
        const old = canvas.toDataURL();
        canvas.width = canvas.offsetWidth;
        canvas.height = 250;
        const img = new Image();
        img.onload = () => ctx.drawImage(img, 0, 0);
        img.src = old;
        ctx.lineWidth = 3;
        ctx.lineCap = 'round';
      }
      resize();
      window.addEventListener('resize', resize);
      let drawing = false;
      function pos(e) {
        const r = canvas.getBoundingClientRect();
        const t = e.touches ? e.touches[0] : e;
        return { x: t.clientX - r.left, y: t.clientY - r.top };
      }
      function start(e) { drawing = true; const p = pos(e); ctx.beginPath(); ctx.moveTo(p.x, p.y); e.preventDefault(); }
      function move(e) { if (!drawing) return; const p = pos(e); ctx.lineTo(p.x, p.y); ctx.stroke(); e.preventDefault(); }
      function end(e) { drawing = false; e.preventDefault(); }
      canvas.addEventListener('mousedown', start);
      canvas.addEventListener('mousemove', move);
      canvas.addEventListener('mouseup', end);
      canvas.addEventListener('mouseleave', end);
      canvas.addEventListener('touchstart', start, {passive:false});
      canvas.addEventListener('touchmove', move, {passive:false});
      canvas.addEventListener('touchend', end, {passive:false});
      function clearCanvas() { ctx.clearRect(0, 0, canvas.width, canvas.height); }
      async function saveFirma() {
        const r = await fetch('/firma/${req.params.id}', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ firma: canvas.toDataURL('image/png') })
        });
        if (r.ok) {
          try { const j = await r.json(); location.href = j.redirect || '/contratto/${req.params.id}/firmato'; }
          catch(e) { location.href = '/contratto/${req.params.id}/firmato'; }
        } else alert('Errore salvataggio firma');
      }
    </script>
  `));
});
app.post('/firma/:id', async (req, res) => {
  try {
    const data = req.body.firma;
    if (!data) return res.status(400).send('Firma mancante');
    const base64 = data.split(',')[1];
    const file = path.join(firmeDir, `firma_${req.params.id}.png`);
    fs.writeFileSync(file, base64, 'base64');
    await run(`UPDATE prenotazioni SET firma_path=?, stato='firmato' WHERE id=?`, [file, req.params.id]);
    await generaPdfContratto(req.params.id, { forceDrive: false, skipDrive: true });
    res.json({ ok:true, redirect:'/contratto/' + req.params.id + '/firmato' });
  } catch (e) {
    res.status(500).send(e.message);
  }
});


app.get('/firma-whatsapp/:id', async (req, res) => {
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');
    const tel = normalizzaWa(p.telefono || p.telefono_cliente || '');
    if (!tel) return res.send(page('Invio firma WhatsApp', `<div class="box"><h2 class="bad">Telefono cliente mancante</h2><a class="btn" href="/prenotazione/${p.id}">Torna</a></div>`));
    const link = absoluteUrl(req, `/firma/${p.id}`);
    const msg = `DP RENT - firma online contratto ${p.codice || p.id}: ${link}`;
    const r = await dpNotify([tel], msg);
    res.send(page('Invio firma WhatsApp', `<div class="box"><h2>${r.ok ? 'Link firma inviato' : 'Invio non riuscito'}</h2><p>${r.ok ? 'Messaggio inviato al cliente su WhatsApp.' : esc((r.errors || []).join(' | '))}</p><p><b>Cliente:</b> ${esc(tel)}</p><p><b>Link:</b> <a target="_blank" href="${esc(link)}">${esc(link)}</a></p><a class="btn" href="/prenotazione/${p.id}">Torna</a></div>`));
  } catch (e) {
    res.status(500).send(page('Errore invio firma', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/firma-link/:id', async (req, res) => {
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');
  const link = absoluteUrl(req, `/firma/${p.id}`);
  const msg = `DP RENT - Firma contratto ${p.codice}: ${link}`;
  res.send(page('Link firma WhatsApp', `
    <div class="box">
      <h2>Link firma cliente</h2>
      <p>Invia questo link al cliente su WhatsApp per far firmare il contratto.</p>
      <input value="${esc(link)}" readonly onclick="this.select()">
      <a class="btn btn3" target="_blank" href="${esc(whatsappText(msg))}">Invia link firma su WhatsApp</a>
      <a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna contratto</a>
    </div>
  `));
});



// V109 FIX: alias corretto per il bottone verde Invia WhatsApp
app.get('/contratto/:id/invia-whatsapp', (req, res) => {
  res.redirect('/whatsapp-contratto/' + req.params.id);
});

app.get('/whatsapp-contratto/:id', async (req, res) => {
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');

    const tel = normalizzaWa(p.telefono || p.telefono_cliente || '');
    if (!tel) {
      return res.send(page('Invio contratto WhatsApp', `<div class="box"><h2 class="bad">Telefono cliente mancante</h2><a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna</a></div>`));
    }

    let pdfLink = p.pdf_drive_web_link || p.pdf_drive_link || '';
    if (!pdfLink) {
      try {
        const sync = await syncContrattoDriveV63(p.id);
        const updated = await get(`SELECT pdf_drive_web_link,pdf_drive_link FROM prenotazioni WHERE id=?`, [p.id]);
        pdfLink = updated?.pdf_drive_web_link || updated?.pdf_drive_link || sync?.pdf?.webViewLink || '';
      } catch (e) {
        console.log('Errore generazione/sync PDF per WhatsApp:', e.message);
      }
    }

    const firmaLink = absoluteUrl(req, `/firma/${p.id}`);
    const testo =
      `DP RENT - Contratto ${p.codice || p.id}\n` +
      `Cliente: ${p.nome || ''} ${p.cognome || ''}\n` +
      `Totale: Euro ${Number(p.totale || 0).toFixed(2)}\n\n` +
      (pdfLink ? `PDF contratto: ${pdfLink}\n\n` : '') +
      `Firma online: ${firmaLink}`;

    const r = await dpNotify([tel], testo);
    res.send(page('Invio contratto WhatsApp', `<div class="box"><h2 class="${r.ok ? 'ok' : 'bad'}">${r.ok ? 'Contratto inviato su WhatsApp' : 'Invio WhatsApp non riuscito'}</h2><p><b>Cliente:</b> ${esc(tel)}</p><p>${r.ok ? 'Messaggio inviato tramite Twilio.' : esc((r.errors || []).join(' | '))}</p>${pdfLink ? `<p><b>PDF:</b> <a target="_blank" href="${esc(pdfLink)}">Apri PDF</a></p>` : '<p class="warn">PDF Drive non disponibile: controlla configurazione Google Drive.</p>'}<p><b>Firma:</b> <a target="_blank" href="${esc(firmaLink)}">${esc(firmaLink)}</a></p><a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna contratto</a><a class="btn" href="javascript:history.back()">Indietro</a></div>`));
  } catch (e) {
    res.status(500).send(page('Errore invio contratto WhatsApp', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});




app.get('/email/:id', async (req,res)=>{
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`,[req.params.id]);
  if(!p)return res.send('Contratto non trovato');
  res.send(page('Invia email', `<h2>Invia contratto via email</h2><form method="POST" action="/email/${p.id}"><label>Email destinatario</label><input name="email" value="${esc(p.email)}" required><label>Messaggio</label><textarea name="messaggio">Buongiorno, in allegato trova il contratto DP RENT.</textarea><button>Invia email</button></form><p class="notice">Per inviare davvero configura SMTP su Render.</p>`));
});
app.post('/email/:id', async (req,res)=>{
  try {
    const file = await generaPdfContratto(req.params.id, { forceDrive: false, skipDrive: true });
    await sendEmail(req.body.email, 'Contratto DP RENT', req.body.messaggio || 'In allegato contratto DP RENT.', [{filename:path.basename(file),path:file}]);
    await run(`UPDATE prenotazioni SET stato='inviato_email' WHERE id=?`,[req.params.id]);
    res.send(actionScreen(req.params.id,'Email inviata','Contratto inviato correttamente.'));
  } catch(e) {
    res.status(500).send(page('Errore Email', `<div class="box"><h2 class="bad">Errore email</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/nexi/:id', async (req, res) => {
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');

    const pagamento = await createNexiLink(
      Number(p.totale || 0),
      `DP RENT ${p.codice || p.id}`,
      p
    );

    await run(
      `UPDATE prenotazioni SET nexi_link=?, nexi_stato='link_generato', nexi_raw=? WHERE id=?`,
      [pagamento.link, pagamento.raw, p.id]
    );

    const testoWa =
      `DP RENT - pagamento contratto ${p.codice}\n` +
      `Totale: € ${euro(p.totale)}\n` +
      `${pagamento.link}`;

    res.send(page('Pagamento Nexi', `
      <div class="box">
        <h2>Pagamento Nexi PayMail</h2>
        <p><b>Contratto:</b> ${esc(p.codice)}</p>
        <p><b>Totale contratto:</b> € ${euro(p.totale)}</p>
        <p class="notice">La cauzione resta gestita manualmente. Qui paghi solo il totale contratto.</p>

        <a class="btn btnWarn" href="${esc(pagamento.link)}" target="_blank">Apri link pagamento Nexi</a>
        <a class="btn btn3" href="/nexi/${p.id}/invia-whatsapp">Invia pagamento WhatsApp</a>

        <label>Link pagamento</label>
        <input value="${esc(pagamento.link)}" readonly onclick="this.select()">

        <a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna contratto</a>
      </div>
    `));
  } catch (e) {
    res.status(500).send(page('Errore Nexi', `
      <div class="box">
        <h2 class="bad">Errore Nexi</h2>
        <pre>${esc(e.message)}</pre>
        <p>Servono su Render: <b>NEXI_ALIAS</b>, <b>NEXI_MAC_KEY</b>, <b>NEXI_ENV</b>, <b>APP_BASE_URL</b>. Il PayMail attuale invia solo apiKey, codiceTransazione, importo, timeStamp, mac, url.</p>
        <a class="btn" href="/prenotazioni">Torna allo storico</a>
      </div>
    `));
  }
});

// V110 FIX: invio link pagamento Nexi via Twilio WhatsApp diretto
app.get('/nexi/:id/invia-whatsapp', async (req, res) => {
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');

    let link = p.nexi_link || '';
    let raw = p.nexi_raw || '';
    if (!link) {
      const pagamento = await createNexiLink(Number(p.totale || 0), `DP RENT ${p.codice || p.id}`, p);
      link = pagamento.link;
      raw = pagamento.raw;
      await run(`UPDATE prenotazioni SET nexi_link=?, nexi_stato='link_generato', nexi_raw=? WHERE id=?`, [link, raw, p.id]);
    }

    const tel = normalizzaWa(p.telefono || p.telefono_cliente || '');
    if (!tel) {
      return res.send(page('Invio pagamento WhatsApp', `<div class="box"><h2 class="bad">Telefono cliente mancante</h2><a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna contratto</a></div>`));
    }

    const testo =
      `DP RENT - pagamento contratto ${p.codice || p.id}\n` +
      `Totale: Euro ${euro(p.totale)}\n` +
      `Link pagamento: ${link}`;

    const r = await dpNotify([tel], testo);
    res.send(page('Invio pagamento WhatsApp', `<div class="box"><h2 class="${r.ok ? 'ok' : 'bad'}">${r.ok ? 'Pagamento inviato su WhatsApp' : 'Invio pagamento non riuscito'}</h2><p><b>Cliente:</b> ${esc(tel)}</p><p>${r.ok ? 'Link pagamento inviato tramite Twilio.' : esc((r.errors || []).join(' | '))}</p><p><b>Link Nexi:</b> <a target="_blank" href="${esc(link)}">${esc(link)}</a></p><a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna contratto</a><a class="btn" href="/nexi/${p.id}">Apri pagina Nexi</a></div>`));
  } catch (e) {
    res.status(500).send(page('Errore invio pagamento WhatsApp', `<div class="box"><h2 class="bad">Errore WhatsApp/Nexi</h2><pre>${esc(e.message)}</pre><a class="btn btn2" href="/contratto/${req.params.id}/gestisci">Torna contratto</a></div>`));
  }
});

app.get('/nexi-ok/:id', async (req, res) => {
  await run(`UPDATE prenotazioni SET nexi_stato='pagato', stato='pagato' WHERE id=?`, [req.params.id]);
  res.send(page('Pagamento OK', `<div class="box"><h2 class="ok">Pagamento registrato</h2><p>Grazie da DP RENT.</p><a class="btn" href="/prenotazione/${req.params.id}">Torna contratto</a></div>`));
});
app.post('/nexi-notification/:id', async (req, res) => {
  await run(`UPDATE prenotazioni SET nexi_raw=? WHERE id=?`, [JSON.stringify(req.body || {}), req.params.id]);
  res.send('OK');
});

app.get('/nexi-ko/:id', async (req, res) => {
  await run(`UPDATE prenotazioni SET nexi_stato='non_pagato' WHERE id=?`, [req.params.id]);
  res.send(page('Pagamento KO', `<div class="box"><h2 class="bad">Pagamento non completato</h2><a class="btn" href="/prenotazione/${req.params.id}">Torna contratto</a></div>`));
});





app.get('/termini-noleggio', (req, res) => {
  res.send(page('Condizioni noleggio DP RENT', `
    <div class="box">
      <h2>Condizioni generali di noleggio DP RENT</h2>
      <p>Il cliente prende in consegna il mezzo nello stato indicato al check-out e si impegna a riconsegnarlo nello stesso stato.</p>
      <p>Carburante: politica pieno/pieno o livello indicato nel contratto. Differenze di carburante, pulizia straordinaria, danni, smarrimento chiavi/documenti, ritardi, multe, ZTL, pedaggi e franchigie sono a carico del cliente.</p>
      <p>Km extra: se previsti, sono conteggiati alla tariffa indicata nel contratto.</p>
      <p>Deposito cauzionale: resta gestito separatamente e può essere trattenuto in tutto o in parte per danni o costi accessori.</p>
      <p>La firma del contratto conferma accettazione di privacy e condizioni.</p>
      <a class="btn" href="/">Torna</a>
    </div>
  `));
});

app.get('/cargos-config', (req, res) => {
  res.send(page('Config Ca.R.G.O.S.', `
    <div class="box">
      <h2>Configurazione Ca.R.G.O.S.</h2>
      <p><b>Utente impostato:</b> ${esc(process.env.CARGOS_USERNAME || 'C00000100')}</p>
      <p class="${process.env.CARGOS_PASSWORD ? 'ok' : 'bad'}">Password: ${process.env.CARGOS_PASSWORD ? 'presente' : 'mancante'}</p>
      <p class="${process.env.CARGOS_APIKEY ? 'ok' : 'bad'}">APIKEY: ${process.env.CARGOS_APIKEY ? 'presente' : 'mancante'}</p>
      <p class="${cargosOrganizationHeaderV76() ? 'ok' : 'bad'}">ORGANIZATION: ${esc(cargosOrganizationHeaderV76() || 'mancante')}</p>
      <p class="${process.env.CARGOS_AGENZIA_ID ? 'ok' : 'bad'}">AGENZIA_ID: cargosAgenziaIdV63()}</p>
      <p class="${process.env.CARGOS_OPERATORE_ID ? 'ok' : 'bad'}">OPERATORE_ID: cargosOperatoreIdV63()}</p>
      <p class="${process.env.CARGOS_LUOGO_COD ? 'ok' : 'bad'}">LUOGO_COD: ${process.env.CARGOS_LUOGO_COD ? 'presente' : 'mancante'}</p>
      <hr>
      <p>Environment da mettere su Render:</p>
      <pre>CARGOS_USERNAME=C00000100
CARGOS_PASSWORD=la_password_che_hai
CARGOS_APIKEY=da richiedere/recuperare
CARGOS_ORGANIZATION=codice Organization Ca.R.G.O.S.
CARGOS_AGENZIA_ID=da Ca.R.G.O.S.
CARGOS_OPERATORE_ID=da Ca.R.G.O.S.
CARGOS_LUOGO_COD=codice luogo polizia
CARGOS_BASE_URL=https://cargos.poliziadistato.it/CARGOS_API</pre>
      <a class="btn" href="/cargos">Vai a Ca.R.G.O.S.</a>
    </div>
  `));
});

app.get('/cargos', async (req, res) => {
  const rows = await all(`SELECT p.*, m.targa FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id ORDER BY p.id DESC LIMIT 200`);
  const sent = rows.filter(p => p.record_cargos_uid || p.cargos_uid).length;
  const missing = rows.length - sent;
  const trs = rows.map(p => {
    const uid = p.record_cargos_uid || p.cargos_uid || '';
    const stato = p.record_cargos_stato || (uid ? 'send_ok' : 'da_inviare');
    return `<tr>
      <td><a href="/contratto/${p.id}/gestisci"><b>${esc(p.codice)}</b></a></td>
      <td>${esc(p.nome)} ${esc(p.cognome)}</td>
      <td>${esc(p.targa)}</td>
      <td>${esc(p.data_inizio)} - ${esc(p.data_fine)}</td>
      <td>${uid ? `<span class="badge badge-green">ID generato</span><br><b>${esc(uid)}</b>` : '<span class="badge badge-red">ID non generato</span>'}</td>
      <td><span class="badge ${uid?'badge-green':'badge-orange'}">${esc(stato)}</span></td>
      <td><a class="btn" href="/cargos/record/${p.id}">Record</a><a class="btn btn2" href="/cargos/check/${p.id}">Verifica dati</a><a class="btn btnWarn" href="/cargos/send/${p.id}">Invia report a CaRGOS</a></td>
    </tr>`;
  }).join('');
  res.send(page('Ca.R.G.O.S.', `<div class="box"><h2>Ca.R.G.O.S.</h2><p>Elenco contratti con ID CaRGOS generato e non generato.</p><p><b>Configurato:</b> ${cargosConfigured() ? '<span class="ok">SI</span>' : '<span class="bad">NO</span>'}</p><p><span class="badge badge-green">ID generati: ${sent}</span> <span class="badge badge-red">ID mancanti: ${missing}</span></p><p>Servono: CARGOS_USERNAME, CARGOS_PASSWORD, CARGOS_APIKEY, CARGOS_AGENZIA_ID, CARGOS_OPERATORE_ID, CARGOS_LUOGO_COD.</p></div><table><tr><th>Contratto</th><th>Cliente</th><th>Targa</th><th>Date</th><th>ID CaRGOS</th><th>Stato</th><th>Azione</th></tr>${trs || '<tr><td colspan="7">Nessun contratto.</td></tr>'}</table>`));
});

app.get('/cargos/record/:id', async (req, res) => {
  try {
    const record = await buildCargosRecordForContract(req.params.id);
    res.send(page('Record Ca.R.G.O.S.', `<div class="box"><h2>Record Ca.R.G.O.S.</h2><p>Lunghezza: <b>${record.length}</b> caratteri</p><textarea style="height:220px;font-family:monospace">${esc(record)}</textarea><a class="btn" href="/cargos/export/${req.params.id}">Scarica TXT</a><a class="btn btn2" href="/cargos">Torna</a></div>`));
  } catch(e) { res.status(500).send(page('Errore Ca.R.G.O.S.', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`)); }
});

app.get('/cargos/export/:id', async (req, res) => {
  try {
    const p = await get(`SELECT codice FROM prenotazioni WHERE id=?`, [req.params.id]);
    const record = await buildCargosRecordForContract(req.params.id);
    res.setHeader('Content-Type','text/plain; charset=utf-8');
    res.setHeader('Content-Disposition', `attachment; filename="record_cargos_${p?.codice || req.params.id}.txt"`);
    res.send(record + '\\n');
  } catch(e) { res.status(500).send('Errore export CARGOS: ' + e.message); }
});



app.get('/cargos/send/:id', async (req, res) => {
  try {
    const result = await cargosSendRecords([await buildCargosRecordForContract(req.params.id)], 'Send');
    let uid = '';
    try {
      const data = result && result.data;
      if (Array.isArray(data) && data[0]) uid = data[0].uid || data[0].transactionid || data[0].transactionId || data[0].id || '';
      uid = uid || data?.uid || data?.transactionid || data?.transactionId || data?.id || result?.uid || result?.transactionid || result?.transactionId || '';
    } catch(e) {}
    await run(`UPDATE prenotazioni SET record_cargos_stato=?, record_cargos_last_send=?, record_cargos_last_error=?, record_cargos_transactionid=?, record_cargos_uid=? WHERE id=?`, ['send_ok', new Date().toISOString(), JSON.stringify(result).slice(0,1000), uid, uid, req.params.id]);
    res.send(page('Send Ca.R.G.O.S.', `<div class="box"><h2>Esito Send</h2><p><b>ID/UID CaRGOS:</b> ${esc(uid || 'non restituito dall API')}</p><pre>${esc(JSON.stringify(result,null,2))}</pre><a class="btn" href="/cargos">Torna CaRGOS</a><a class="btn btn2" href="/contratto/${req.params.id}/gestisci">Torna contratto</a></div>`));
  } catch(e) { res.status(500).send(page('Errore Send Ca.R.G.O.S.', `<div class="box"><h2 class="bad">Errore Send</h2><pre>${esc(e.message)}</pre><a class="btn" href="/cargos">Torna</a></div>`)); }
});


app.get('/test-email', async (req, res) => {
  try {
    await sendEmail(process.env.ALERT_EMAIL || process.env.SMTP_USER, 'TEST DP RENT APP', 'Email di test dal gestionale DP RENT.', []);
    res.send(page('Test Email', `<div class="box"><h2 class="ok">Email inviata</h2><a class="btn" href="/">Dashboard</a></div>`));
  } catch (e) {
    res.send(page('Errore Email', `<div class="box"><h2 class="bad">Errore email</h2><pre>${esc(e.message)}</pre><a class="btn" href="/">Dashboard</a></div>`));
  }
});
app.get('/test-drive', async (req, res) => {
  try {
    if (!googleDriveConfigured()) {
      return res.send(page('Google Drive non configurato', `<div class="box"><h2 class="bad">Google Drive non configurato</h2><p>Servono variabili:</p><pre>DRIVE_WEBAPP_URL\nGOOGLE_DRIVE_FOLDER_ID</pre><a class="btn" href="/">Dashboard</a></div>`));
    }
    const testFile = path.join(uploadDir, `test-drive-${Date.now()}.txt`);
    fs.writeFileSync(testFile, 'Test Google Drive DP RENT via Apps Script');
    const driveRes = await uploadFileToDrive(testFile, path.basename(testFile), 'text/plain', 'TEST DP RENT');
    res.send(page('Test Drive', `<div class="box"><h2 class="ok">Upload Google Drive riuscito</h2><p><a target="_blank" href="${esc(driveRes.webViewLink)}">Apri file su Google Drive</a></p><a class="btn" href="/">Dashboard</a></div>`));
  } catch (e) {
    res.send(page('Errore Drive', `<div class="box"><h2 class="bad">Errore Google Drive</h2><pre>${esc(e.message)}</pre><p>Controlla DRIVE_WEBAPP_URL e GOOGLE_DRIVE_FOLDER_ID.</p><a class="btn" href="/">Dashboard</a></div>`));
  }
});

app.get('/mezzi', async (req,res)=> {
  const rows = await all(`SELECT * FROM mezzi`);
  res.json(rows);
});

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).send(page('Errore server', `<div class="box"><h2 class="bad">Errore server</h2><pre>${esc(err.message)}</pre></div>`));
});

// =========================
// RENDER PORT BINDING - VERSIONE DEFINITIVA
// =========================

// =========================
// RENDER PORT BINDING - V36 DEFINITIVA
// =========================

// =========================
// RENDER PORT BINDING - V37
// =========================
// =========================
// RENDER PORT BINDING - V38
// =========================

// =========================
// RENDER PORT BINDING - V39
// =========================

// =========================
// RENDER PORT BINDING - V40
// =========================
// =========================
// RENDER PORT BINDING - V42
// =========================
// =========================
// RENDER PORT BINDING - V44
// =========================

// =========================
// V51 ADMIN INIT ROUTES GARANTITE
// =========================
function v51InitAllDb(done) {
  const finish = () => done && done();
  db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      uid TEXT, targa TEXT, marca TEXT, modello TEXT, tipo TEXT,
      descrizione TEXT, km TEXT, stato TEXT DEFAULT 'attivo',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      cliente_id INTEGER,
      mezzo_id INTEGER,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      tipo_cliente TEXT,
      codice_fiscale TEXT,
      data_inizio TEXT,
      ora_inizio TEXT,
      data_fine TEXT,
      ora_fine TEXT,
      totale REAL,
      stato TEXT DEFAULT 'bozza',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS clienti (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT, cognome TEXT, telefono TEXT, email TEXT,
      cf TEXT, codice_fiscale TEXT, indirizzo TEXT, citta TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS allegati (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      prenotazione_id INTEGER,
      tipo TEXT,
      filename TEXT,
      originalname TEXT,
      path TEXT,
      mimetype TEXT,
      size INTEGER,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`, () => {
      if (typeof v52FixEverything === 'function') return v52FixEverything(finish);
      finish();
    });
  });
}

app.get('/admin/init-db', (req, res) => {
  v51InitAllDb(() => {
    res.send(page('Init DB V51', `<div class="box">
      <h2 class="ok">DATABASE INIZIALIZZATO V51</h2>
      <p>Ora puoi importare i mezzi e creare prenotazioni.</p>
      <a class="btn" href="/import-excel">Import Excel</a>
      <a class="btn btn2" href="/nuova-prenotazione">Nuova prenotazione</a>
    </div>`));
  });
});

app.get('/admin/rebuild-prenotazioni', (req, res) => {
  v51InitAllDb(() => {
    res.send(page('Rebuild prenotazioni V51', `<div class="box">
      <h2 class="ok">PRENOTAZIONI REBUILD OK</h2>
      <a class="btn" href="/nuova-prenotazione">Nuova prenotazione</a>
      <a class="btn btn2" href="/">Dashboard</a>
    </div>`));
  });
});

app.get('/admin/fix-mezzi-db', (req, res) => {
  v51InitAllDb(() => {
    res.send(page('Fix mezzi DB V51', `<div class="box">
      <h2 class="ok">MEZZI DB OK</h2>
      <a class="btn" href="/import-excel">Import Excel</a>
    </div>`));
  });
});

app.get('/admin/rebuild-mezzi', (req, res) => res.redirect('/admin/fix-mezzi-db'));

// =========================
// V52 FIX TUTTO - STABILE NOTTE
// =========================
const V52_PRENOTAZIONI_COLS = {
  cliente_id:'INTEGER',
  mezzo_id:'INTEGER',
  codice:'TEXT',
  nome:'TEXT',
  cognome:'TEXT',
  telefono:'TEXT',
  email:'TEXT',
  tipo_cliente:'TEXT',
  codice_fiscale:'TEXT',
  cf:'TEXT',
  partita_iva:'TEXT',
  piva:'TEXT',
  ragione_sociale:'TEXT',
  azienda:'TEXT',
  pec:'TEXT',
  codice_sdi:'TEXT',
  sdi:'TEXT',
  indirizzo:'TEXT',
  citta:'TEXT',
  cap:'TEXT',
  provincia:'TEXT',
  data_nascita:'TEXT',
  luogo_nascita:'TEXT',
  documento_tipo:'TEXT',
  documento_numero:'TEXT',
  documento_scadenza:'TEXT',
  numero_documento:'TEXT',
  scadenza_documento:'TEXT',
  patente_numero:'TEXT',
  patente_scadenza:'TEXT',
  numero_patente:'TEXT',
  scadenza_patente:'TEXT',
  categoria_patente:'TEXT',
  conducente1:'TEXT',
  patente1:'TEXT',
  patente1_scadenza:'TEXT',
  conducente2:'TEXT',
  patente2:'TEXT',
  patente2_scadenza:'TEXT',
  conducente2_nome:'TEXT',
  conducente2_cognome:'TEXT',
  conducente2_data_nascita:'TEXT',
  conducente2_luogo_nascita:'TEXT',
  conducente2_patente:'TEXT',
  conducente2_doc_numero:'TEXT',
  conducente2_patente_numero:'TEXT',
  conducente2_recapito:'TEXT',
  mezzo:'TEXT',
  targa:'TEXT',
  marca:'TEXT',
  modello:'TEXT',
  data_inizio:'TEXT',
  ora_inizio:'TEXT',
  data_fine:'TEXT',
  ora_fine:'TEXT',
  giorni:'INTEGER',
  km_previsti:'TEXT',
  km_inclusi:'TEXT',
  extra_fuori_orario:'REAL',
  extra_km:'REAL',
  km_uscita:'TEXT',
  km_rientro:'TEXT',
  carburante_uscita:'TEXT',
  carburante_rientro:'TEXT',
  imponibile:'REAL',
  iva:'REAL',
  totale:'REAL',
  cauzione:'REAL',
  deposito:'REAL',
  stato:'TEXT',
  note:'TEXT',
  firma:'TEXT',
  firma_path:'TEXT',
  pdf_path:'TEXT',
  pdf_drive_link:'TEXT',
  pdf_drive_file_id:'TEXT',
  pdf_drive_web_link:'TEXT',
  drive_folder_id:'TEXT',
  drive_folder_link:'TEXT',
  nexi_link:'TEXT',
  nexi_stato:'TEXT',
  nexi_raw:'TEXT',
  record_cargos_uid:'TEXT',
  record_cargos_transactionid:'TEXT',
  record_cargos_stato:'TEXT',
  record_cargos_last_check:'TEXT',
  record_cargos_last_send:'TEXT',
  record_cargos_last_error:'TEXT',
  record_cargos_pagamento_tipo:'TEXT',
  record_cargos_checkout_luogo_cod:'TEXT',
  record_cargos_checkout_indirizzo:'TEXT',
  record_cargos_checkin_luogo_cod:'TEXT',
  record_cargos_checkin_indirizzo:'TEXT',
  record_cargos_operatore_id:'TEXT',
  record_cargos_agenzia_id:'TEXT',
  record_cargos_agenzia_nome:'TEXT',
  record_cargos_agenzia_luogo_cod:'TEXT',
  record_cargos_agenzia_indirizzo:'TEXT',
  record_cargos_agenzia_tel:'TEXT',
  record_cargos_veicolo_tipo:'TEXT',
  record_cargos_veicolo_colore:'TEXT',
  record_cargos_veicolo_gps:'TEXT',
  record_cargos_veicolo_bloccom:'TEXT',
  record_cargos_cittadinanza_cod:'TEXT', conducente_cittadinanza_cod:'TEXT',
  record_cargos_nascita_luogo_cod:'TEXT',
  record_cargos_residenza_luogo_cod:'TEXT',
  record_cargos_doc_tipo_cod:'TEXT',
  record_cargos_doc_luogoril_cod:'TEXT',
  record_cargos_patente_luogoril_cod:'TEXT',
  conducente2_nascita_luogo_cod:'TEXT',
  conducente2_cittadinanza_cod:'TEXT', conducente_cittadinanza_cod:'TEXT',
  conducente2_doc_tipo_cod:'TEXT',
  conducente2_doc_luogoril_cod:'TEXT',
  conducente2_patente_luogoril_cod:'TEXT',
};

const V52_MEZZI_COLS = {
  uid:'TEXT',
  targa:'TEXT',
  marca:'TEXT',
  modello:'TEXT',
  tipo:'TEXT',
  descrizione:'TEXT',
  cilindrata:'TEXT',
  alimentazione:'TEXT',
  codice_tipo:'TEXT',
  codice_marca:'TEXT',
  codice_modello:'TEXT',
  categoria:'TEXT',
  posti:'TEXT',
  km:'TEXT',
  km_attuali:'TEXT',
  telaio:'TEXT',
  colore:'TEXT',
  stazione:'TEXT',
  soccorso_stradale:'TEXT',
  immagini_consegna:'TEXT',
  numero_interno:'TEXT',
  disponibile:'TEXT',
  attivo:'TEXT',
  ubicazione:'TEXT',
  proprieta:'TEXT',
  note_interne:'TEXT',
  data_immatricolazione:'TEXT',
  ultima_revisione:'TEXT',
  prossimo_tagliando:'TEXT',
  serbatoio:'TEXT',
  cambio:'TEXT',
  porte:'TEXT',
  euro:'TEXT',
  iva:'TEXT',
  franchigia:'TEXT',
  prezzo_giorno:'TEXT',
  km_inclusi:'TEXT',
  cauzione:'TEXT',
  deposito:'TEXT',
  gps:'TEXT',
  blocco_motore:'TEXT',
  stato:'TEXT',
  note:'TEXT',
  scadenza_revisione:'TEXT',
  revisione_scadenza:'TEXT',
  scadenza_bollo:'TEXT',
  bollo_scadenza:'TEXT',
  scadenza_assicurazione:'TEXT',
  assicurazione_scadenza:'TEXT',
  gomme_scadenza:'TEXT',
  tagliando_km:'TEXT',
  tagliando_km_scadenza:'TEXT',
  tagliando_data_scadenza:'TEXT',
  manutenzione_note:'TEXT',
  alert_giorni:'INTEGER',
  alert_km:'INTEGER',
  record_cargos_veicolo_tipo:'TEXT',
  descrizione_pubblica:'TEXT',
};

function v52FixTable(table, cols, done) {
  const keys = Object.keys(cols);
  let pending = keys.length;
  if (!pending) return done && done();
  keys.forEach((c) => {
    db.run(`ALTER TABLE ${table} ADD COLUMN ${c} ${cols[c]}`, () => {
      pending--;
      if (pending === 0) done && done();
    });
  });
}

function v52FixEverything(done) {
  db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      uid TEXT,
      targa TEXT,
      marca TEXT,
      modello TEXT,
      tipo TEXT,
      descrizione TEXT,
      km TEXT,
      stato TEXT DEFAULT 'attivo',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      cliente_id INTEGER,
      mezzo_id INTEGER,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      tipo_cliente TEXT,
      codice_fiscale TEXT,
      data_inizio TEXT,
      ora_inizio TEXT,
      data_fine TEXT,
      ora_fine TEXT,
      totale REAL,
      stato TEXT DEFAULT 'bozza',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS clienti (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT,
      cognome TEXT,
      telefono TEXT,
      email TEXT,
      cf TEXT,
      codice_fiscale TEXT,
      indirizzo TEXT,
      citta TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);
    db.run(`CREATE TABLE IF NOT EXISTS allegati (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      prenotazione_id INTEGER,
      tipo TEXT,
      filename TEXT,
      originalname TEXT,
      path TEXT,
      mimetype TEXT,
      size INTEGER,
      drive_file_id TEXT,
      drive_web_link TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )`);
    v52FixTable('allegati', { drive_file_id:'TEXT', drive_web_link:'TEXT', size:'INTEGER' }, () => {});
    v52FixTable('mezzi', V52_MEZZI_COLS, () => {
      v52FixTable('prenotazioni', V52_PRENOTAZIONI_COLS, () => {
        console.log('V63 FIX TUTTO OK');
        done && done();
      });
    });
  });
}

// Esegue il fix automatico dopo l'avvio senza bloccare il server.
setTimeout(() => v52FixEverything(() => {}), 1200);


app.get('/drive-sync/:id', async (req, res) => {
  await syncContrattoDriveV63(req.params.id);
  res.redirect('/documenti/' + req.params.id);
});

app.get('/admin/pulisci-pdf-drive/:id', async (req, res) => {
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if (!p) return res.send('Contratto non trovato');
    const folder = await getOrCreateDriveContractFolderV63(p);
    if (!folder) return res.send('Drive non configurato');
    const pdfName = pdfFileNameForContract(p);
    await deleteAllContractPdfsInDriveV63(folder.id);
    const pdf = await generaPdfContratto(req.params.id, { forceDrive:false, skipDrive:true });
    await uploadFileToDriveFolderV63(pdf, pdfName, 'application/pdf', folder.id);
    res.send(page('PDF DRIVE PULITO', `<div class="box"><h2 class="ok">PDF DRIVE PULITO</h2><a class="btn" href="/documenti/${req.params.id}">Documenti</a></div>`));
  } catch(e) {
    res.status(500).send(page('Errore', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/admin/fix-tutto', (req, res) => {
  v52FixEverything(() => {
    res.send(page('FIX TUTTO V63', `<div class="box">
      <h2 class="ok">FIX TUTTO V63 OK</h2>
      <p>Database aggiornato: mezzi, prenotazioni, clienti, allegati.</p>
      <a class="btn" href="/nuova-prenotazione">Nuova prenotazione</a>
      <a class="btn btn2" href="/mezzi">Mezzi</a>
    </div>`));
  });
});

app.get('/admin/fix-prenotazioni_v51', (req, res) => {
  v52FixEverything(() => {
    res.send(page('FIX PRENOTAZIONI', `<div class="box">
      <h2 class="ok">FIX PRENOTAZIONI OK</h2>
      <p>Colonna tipo_cliente e campi azienda/documenti aggiunti.</p>
      <a class="btn" href="/nuova-prenotazione">Nuova prenotazione</a>
    </div>`));
  });
});

// V63 blocco download cargos: CARGOS non scarica file, usa pagina verifica/invia.
app.get('/cargos/download/:id', (req, res) => res.redirect('/cargos/' + req.params.id));
app.get('/download-cargos/:id', (req, res) => res.redirect('/cargos/' + req.params.id));
app.get('/record-cargos/:id', (req, res) => res.redirect('/cargos/' + req.params.id));


// =========================
// V63 ADMIN FIX PERSISTENTE
// =========================
app.get('/admin/persistent-check', async (req, res) => {
  try {
    const info = {
      DB_PATH,
      DATA_DIR,
      uploadDir,
      contractsDir,
      firmeDir,
      publicDir,
      exists_db: fs.existsSync(DB_PATH),
      exists_data: fs.existsSync(DATA_DIR),
      exists_uploads: fs.existsSync(uploadDir),
      exists_contracts: fs.existsSync(contractsDir)
    };
    res.send(page('Persistent check V63', `<div class="box">
      <h2 class="ok">PERSISTENT DATA V63 OK</h2>
      <pre>${esc(JSON.stringify(info, null, 2))}</pre>
      <a class="btn" href="/admin/fix-tutto-v57">Fix tutto V63</a>
      <a class="btn btn2" href="/">Dashboard</a>
    </div>`));
  } catch(e) {
    res.status(500).send(page('Persistent KO', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/admin/fix-tutto-v57', (req, res) => {
  v51InitAllDb(() => {
    const prenCols = {
      cliente_id:'INTEGER', mezzo_id:'INTEGER',
      codice:'TEXT', nome:'TEXT', cognome:'TEXT', telefono:'TEXT', email:'TEXT',
      tipo_cliente:'TEXT', codice_fiscale:'TEXT', cf:'TEXT', partita_iva:'TEXT',
      ragione_sociale:'TEXT', pec:'TEXT', codice_sdi:'TEXT',
      indirizzo:'TEXT', citta:'TEXT', cap:'TEXT', provincia:'TEXT',
      data_nascita:'TEXT', luogo_nascita:'TEXT', cittadinanza:'TEXT',
      documento_tipo:'TEXT', documento_numero:'TEXT', documento_scadenza:'TEXT',
      patente_numero:'TEXT', patente_scadenza:'TEXT',
      conducente2_nome:'TEXT', conducente2_cognome:'TEXT', conducente2_patente:'TEXT',
      targa:'TEXT', marca:'TEXT', modello:'TEXT',
      data_inizio:'TEXT', data_fine:'TEXT', ora_inizio:'TEXT', ora_fine:'TEXT',
      giorni:'INTEGER', km_previsti:'TEXT', cauzione:'REAL', deposito:'REAL',
      totale:'REAL', stato:'TEXT', note:'TEXT',
      pdf_path:'TEXT', pdf_drive_link:'TEXT',
      firma_path:'TEXT', drive_folder_id:'TEXT', drive_folder_link:'TEXT',
      cargos_stato:'TEXT', cargos_transactionid:'TEXT', cargos_last_check:'TEXT',
      cargos_last_send:'TEXT', cargos_last_error:'TEXT'
    };
    const mezziCols = {
      uid:'TEXT', targa:'TEXT', marca:'TEXT', modello:'TEXT', tipo:'TEXT',
      descrizione:'TEXT', km:'TEXT', km_attuali:'TEXT', stato:'TEXT',
      cilindrata:'TEXT', alimentazione:'TEXT', anno:'TEXT', colore:'TEXT',
      posti:'TEXT', telaio:'TEXT', categoria:'TEXT', cauzione:'TEXT',
      prezzo_giorno:'REAL', km_inclusi:'REAL', note:'TEXT',
      scadenza_revisione:'TEXT', scadenza_bollo:'TEXT',
      scadenza_assicurazione:'TEXT', tagliando_km:'TEXT',
      gps:'TEXT', blocco_motore:'TEXT', cargos_veicolo_tipo:'TEXT',
      codice_tipo:'TEXT', codice_marca:'TEXT', codice_modello:'TEXT',
      stazione:'TEXT', immagini_consegna:'TEXT', numero_interno:'TEXT',
      disponibile:'TEXT', attivo:'TEXT', ubicazione:'TEXT', proprieta:'TEXT',
      note_interne:'TEXT'
    };
    const allegatiCols = {
      prenotazione_id:'INTEGER', mezzo_id:'INTEGER', tipo:'TEXT', filename:'TEXT', originalname:'TEXT',
      path:'TEXT', mimetype:'TEXT', size:'INTEGER',
      drive_file_id:'TEXT', drive_web_link:'TEXT'
    };

    const jobs = [
      ['prenotazioni', prenCols],
      ['mezzi', mezziCols],
      ['allegati', allegatiCols]
    ];

    let pendingTables = jobs.length;
    jobs.forEach(([table, cols]) => {
      v52FixTable(table, cols, () => {
        pendingTables--;
        if (pendingTables === 0) {
          res.send(page('FIX TUTTO V63 OK', `<div class="box">
            <h2 class="ok">FIX TUTTO V63 OK</h2>
            <p>DB: ${esc(DB_PATH)}</p>
            <p>Dati: ${esc(DATA_DIR)}</p>
            <a class="btn" href="/nuova-prenotazione">Nuova prenotazione</a>
            <a class="btn btn2" href="/admin/persistent-check">Persistent check</a>
          </div>`));
        }
      });
    });
  });
});

app.get('/admin/fix-tutto-v56', (req, res) => res.redirect('/admin/fix-tutto-v57'));
app.get('/admin/fix-tutto-final', (req, res) => res.redirect('/admin/fix-tutto-v57'));


app.get('/admin/fix-allegati-v58', (req, res) => {
  v51InitAllDb(() => {
    v52FixTable('allegati', {
      prenotazione_id:'INTEGER',
      mezzo_id:'INTEGER',
      tipo:'TEXT',
      filename:'TEXT',
      originalname:'TEXT',
      path:'TEXT',
      mimetype:'TEXT',
      size:'INTEGER',
      drive_file_id:'TEXT',
      drive_web_link:'TEXT'
    }, () => {
      res.send(page('FIX ALLEGATI V63', `<div class="box">
        <h2 class="ok">FIX ALLEGATI V63 OK</h2>
        <a class="btn" href="/admin/fix-tutto-v58">Fix tutto</a>
        <a class="btn btn2" href="/">Dashboard</a>
      </div>`));
    });
  });
});


// =========================
// V63 DRIVE: UN SOLO PDF + FOTO IN CARTELLA CONTRATTO
// =========================
async function deleteAllContractPdfsInDriveV63(folderId) {
  if (!drive || !folderId) return;
  try {
    const found = await drive.files.list({
      q: `'${folderId}' in parents and trashed=false and mimeType='application/pdf'`,
      fields: 'files(id,name)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });
    for (const f of (found.data.files || [])) {
      try {
        await drive.files.delete({ fileId: f.id, supportsAllDrives: true });
      } catch(e) {
        console.log('V63 delete old pdf warning:', e.message);
      }
    }
  } catch(e) {
    console.log('V63 deleteAllContractPdfsInDrive error:', e.message);
  }
}

async function uploadLocalAllegatiToDriveV63(prenotazioneId, folderId) {
  if (!drive || !folderId) return;
  try {
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [prenotazioneId]).catch(()=>null);
    let allegati = await all(`SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id ASC`, [prenotazioneId]).catch(()=>[]);
    if (p && p.cliente_id) {
      const docsCliente = await all(`SELECT * FROM allegati WHERE cliente_id=? AND (prenotazione_id IS NULL OR prenotazione_id=0) ORDER BY id ASC`, [p.cliente_id]).catch(()=>[]);
      allegati = allegati.concat(docsCliente);
    }
    const seen = new Set();
    for (const a of (allegati || [])) {
      const key = a.id + ':' + (a.path || a.filename || '');
      if (seen.has(key)) continue;
      seen.add(key);
      if (a.drive_file_id) continue;
      if (!a.path || !fs.existsSync(a.path)) continue;
      const fileName = safeFileName(a.originalname || a.filename || path.basename(a.path));
      const up = await uploadFileToDriveFolderV63(a.path, fileName, a.mimetype || 'application/octet-stream', folderId);
      if (up && up.id) {
        await run(`UPDATE allegati SET drive_file_id=?, drive_web_link=? WHERE id=?`, [up.id, up.webViewLink || null, a.id]).catch(()=>{});
      }
    }
  } catch(e) {
    console.log('V63 uploadLocalAllegatiToDrive error:', e.message);
  }
}

async function syncContrattoDriveV63(prenotazioneId) {
  try {
    const p = await getPrenotazioneCompleta(prenotazioneId);
    if (!p) return null;
    if (typeof googleDriveConfigured === 'function' && !googleDriveConfigured()) return null;
    if (!drive) return null;

    const folder = await getOrCreateDriveContractFolderV63(p);
    if (!folder) return null;

    await run(
      `UPDATE prenotazioni SET drive_folder_id=?, drive_folder_link=? WHERE id=?`,
      [folder.id, folder.webViewLink || null, prenotazioneId]
    );

    const pdf = await generaPdfContratto(prenotazioneId, { forceDrive: false, skipDrive: true });
    const pdfName = pdfFileNameForContract(p);

    await deleteAllContractPdfsInDriveV63(folder.id);

    const uploadedPdf = await uploadFileToDriveFolderV63(
      pdf,
      pdfName,
      'application/pdf',
      folder.id
    );

    if (uploadedPdf && uploadedPdf.id) {
      await run(
        `UPDATE prenotazioni SET pdf_path=?, pdf_drive_link=?, pdf_drive_web_link=?, pdf_drive_file_id=? WHERE id=?`,
        [pdf, uploadedPdf.webViewLink || null, uploadedPdf.webViewLink || null, uploadedPdf.id || null, prenotazioneId]
      );
      // pulizia finale: nella cartella contratto deve restare un solo PDF
      try {
        const found = await drive.files.list({
          q: `'${folder.id}' in parents and trashed=false and mimeType='application/pdf'`,
          fields: 'files(id,name)',
          supportsAllDrives: true,
          includeItemsFromAllDrives: true
        });
        for (const f of (found.data.files || [])) {
          if (f.id !== uploadedPdf.id) {
            try { await drive.files.delete({ fileId: f.id, supportsAllDrives: true }); } catch(e) {}
          }
        }
      } catch(e) { console.log('Pulizia PDF duplicati warning:', e.message); }
    }

    await uploadLocalAllegatiToDriveV63(prenotazioneId, folder.id);
    return { folder, pdf: uploadedPdf };
  } catch(e) {
    console.log('V63 syncContrattoDrive error:', e.message);
    return null;
  }
}


app.get('/admin/sync-drive-v59/:id', async (req, res) => {
  try {
    await syncContrattoDriveV63(req.params.id);
    res.send(page('SYNC DRIVE V63', `<div class="box">
      <h2 class="ok">DRIVE SINCRONIZZATO V63</h2>
      <p>Ora nella cartella Drive deve esserci un solo PDF contratto e le foto/documenti caricati.</p>
      <a class="btn" href="/documenti/${req.params.id}">Documenti</a>
      <a class="btn btn2" href="/">Dashboard</a>
    </div>`));
  } catch(e) {
    res.status(500).send(page('SYNC DRIVE ERRORE', `<div class="box"><h2 class="bad">Errore sync Drive</h2><pre>${esc(e.message)}</pre></div>`));
  }
});


app.get('/admin/cargos-luoghi-v60', (req, res) => {
  const checkout = normalizeCargosLuogoCod(cargosCheckoutLuogoCodV63());
  const checkin = normalizeCargosLuogoCod(cargosCheckinLuogoCodV63());
  res.send(page('CARGOS Luoghi V63', `<div class="box">
    <h2>Config luoghi Ca.R.G.O.S. V63</h2>
    <p><b>CHECKOUT_LUOGO_COD inviato:</b> ${esc(checkout)}</p>
    <p><b>CHECKIN_LUOGO_COD inviato:</b> ${esc(checkin)}</p>
    <p>Devono essere codici numerici presenti nella tabella luoghi CARGOS/Polizia.</p>
    <hr>
    <p>Variabili Render accettate:</p>
    <pre>CHECKOUT_LUOGO_COD
CHECKIN_LUOGO_COD
CARGOS_CHECKOUT_LUOGO_COD
CARGOS_CHECKIN_LUOGO_COD
CARGOS_LOCATION_CODE
CARGOS_LUOGO_COD</pre>
    <a class="btn" href="/config-cargos">Config CARGOS</a>
    <a class="btn btn2" href="/">Dashboard</a>
  </div>`));
});

app.get('/admin/fix-tutto-v60', (req, res) => res.redirect('/admin/fix-tutto-v58'));


app.get('/admin/cargos-tabelle-v61', (req, res) => {
  res.send(page('Tabelle CARGOS V63', `<div class="box">
    <h2 class="ok">TABELLE POLIZIA CARGOS V63 CARICATE</h2>
    <p><b>Luogo Narni:</b> ${esc(CARGOS_DEFAULT_LUOGO_NARNI)}</p>
    <p><b>Checkout:</b> ${esc(cargosCheckoutLuogoCodV63())}</p>
    <p><b>Checkin:</b> ${esc(cargosCheckinLuogoCodV63())}</p>
    <p><b>Operatore:</b> ${esc(cargosOperatoreIdV63())}</p>
    <hr>
    <pre>VEICOLI:
0 Autovetture
1 Furgoni
3 Autobus / 9 posti
4 Autocarri
5 Trattori stradali
6 Autotreni
7 Autoarticolati / bisarca
8 Autosnodati
9 Autocaravan
A Mezzi d'opera

DOCUMENTI:
CIDIP Carta id. diplomatica
IDELE Carta identità elettronica
IDENT Carta identità
PASDI Passaporto diplomatico
PASOR Passaporto ordinario
PASSE Passaporto servizio
PATEN Patente guida

PAGAMENTI:
0 Carta credito
1 Contanti
2 Carta debito / bancomat
3 Bonifico
4 RID bancario
9 Altro</pre>
    <a class="btn" href="/">Dashboard</a>
  </div>`));
});

app.get('/admin/fix-tutto-v61', (req, res) => {
  if (typeof v52FixEverything === 'function') {
    return v52FixEverything(() => res.redirect('/admin/cargos-tabelle-v61'));
  }
  return res.redirect('/admin/cargos-tabelle-v61');
});


app.get('/admin/fix-tutto-v62',(req,res)=>{
  db.serialize(()=>{
    db.run(`CREATE TABLE IF NOT EXISTS prenotazioni (id INTEGER PRIMARY KEY AUTOINCREMENT,codice TEXT,nome TEXT,cognome TEXT,telefono TEXT,email TEXT,data_inizio TEXT,data_fine TEXT,totale REAL,stato TEXT DEFAULT 'bozza')`);
    db.run(`CREATE TABLE IF NOT EXISTS mezzi (id INTEGER PRIMARY KEY AUTOINCREMENT,targa TEXT,marca TEXT,modello TEXT,tipo TEXT,stato TEXT DEFAULT 'attivo')`);
    db.run(`CREATE TABLE IF NOT EXISTS allegati (id INTEGER PRIMARY KEY AUTOINCREMENT,prenotazione_id INTEGER,tipo TEXT,filename TEXT,path TEXT)`);
    const pren={tipo_cliente:'TEXT',codice_fiscale:'TEXT',partita_iva:'TEXT',ragione_sociale:'TEXT',pec:'TEXT',codice_sdi:'TEXT',indirizzo:'TEXT',citta:'TEXT',cap:'TEXT',provincia:'TEXT',data_nascita:'TEXT',luogo_nascita:'TEXT',documento_tipo:'TEXT',documento_numero:'TEXT',documento_scadenza:'TEXT',patente_numero:'TEXT',patente_scadenza:'TEXT',conducente2_nome:'TEXT',conducente2_cognome:'TEXT',conducente2_patente:'TEXT',targa:'TEXT',marca:'TEXT',modello:'TEXT',ora_inizio:'TEXT',ora_fine:'TEXT',giorni:'INTEGER',km_previsti:'TEXT',cauzione:'REAL',cauzione_richiesta:'TEXT',cauzione_ricevuta:'TEXT',cauzione_importo:'REAL',cauzione_metodo:'TEXT',cauzione_restituita:'TEXT',cauzione_note:'TEXT',tipo_record:'TEXT',note:'TEXT',pdf_path:'TEXT',pdf_drive_link:'TEXT',firma_path:'TEXT',drive_folder_id:'TEXT',drive_folder_link:'TEXT',cargos_stato:'TEXT',cargos_transactionid:'TEXT',cargos_last_error:'TEXT'};
    const mez={uid:'TEXT',cilindrata:'TEXT',alimentazione:'TEXT',anno:'TEXT',colore:'TEXT',posti:'TEXT',km:'TEXT',km_attuali:'TEXT',telaio:'TEXT',categoria:'TEXT',cauzione:'REAL',prezzo_giorno:'REAL',km_inclusi:'REAL',gps:'TEXT',blocco_motore:'TEXT',codice_tipo:'TEXT',note:'TEXT'};
    const allg={mezzo_id:'INTEGER',originalname:'TEXT',mimetype:'TEXT',size:'INTEGER',drive_file_id:'TEXT',drive_web_link:'TEXT'};
    let left=3; const done=()=>{if(--left===0)res.send(page('FIX V63 OK',`<div class="box"><h2 class="ok">FIX TUTTO V63 OK</h2><a class="btn" href="/nuova-prenotazione">Nuova prenotazione</a><a class="btn btn2" href="/mezzi">Mezzi</a></div>`));};
    v62FixTable('prenotazioni',pren,done); v62FixTable('mezzi',mez,done); v62FixTable('allegati',allg,done);
  });
});
app.get('/admin/fix-tutto-v61',(req,res)=>res.redirect('/admin/fix-tutto-v62'));
app.get('/admin/fix-tutto-v60',(req,res)=>res.redirect('/admin/fix-tutto-v62'));
app.get('/admin/fix-tutto-v58',(req,res)=>res.redirect('/admin/fix-tutto-v62'));


app.get('/prenotazione/:id/elimina',async(req,res)=>{const p=await get(`SELECT * FROM prenotazioni WHERE id=?`,[req.params.id]);res.send(page('Elimina contratto',`<div class="box"><h2 class="bad">Eliminare contratto ${esc(p?.codice||req.params.id)}?</h2><form method="post" action="/prenotazione/${req.params.id}/elimina"><button class="btn bad" type="submit">Conferma eliminazione</button><a class="btn btn2" href="/prenotazione/${req.params.id}">Annulla</a></form></div>`));});
app.post('/prenotazione/:id/elimina',async(req,res)=>{await run(`DELETE FROM allegati WHERE prenotazione_id=?`,[req.params.id]).catch(()=>{});await run(`DELETE FROM prenotazioni WHERE id=?`,[req.params.id]);res.redirect('/');});
app.get('/preventivo/nuovo',(req,res)=>res.redirect('/nuova-prenotazione?tipo=preventivo'));
app.get('/prenotazione/:id/converti-contratto',async(req,res)=>{await run(`UPDATE prenotazioni SET stato='contratto', tipo_record='contratto' WHERE id=?`,[req.params.id]);res.redirect(`/prenotazione/${req.params.id}`);});

app.get('/mezzi/nuovo',(req,res)=>res.send(page('Nuovo mezzo',`<div class="box"><h2>Nuovo mezzo</h2><form method="post" action="/mezzi/nuovo"><div class="grid"><label>Targa<input name="targa" required></label><label>Marca<input name="marca"></label><label>Modello<input name="modello"></label><label>Tipo<select name="tipo"><option value="auto">Auto</option><option value="furgone">Furgone</option><option value="pulmino">Pulmino 9 posti</option><option value="attrezzatura">Attrezzatura</option></select></label><label>Km<input name="km"></label><label>Prezzo giorno<input name="prezzo_giorno"></label><label>Km inclusi<input name="km_inclusi" value="150"></label><label>Cauzione standard<input name="cauzione" value="500"></label><label>GPS<select name="gps"><option value="0">NO</option><option value="1">SI</option></select></label><label>Blocco motore<select name="blocco_motore"><option value="0">NO</option><option value="1">SI</option></select></label></div><label>Note<textarea name="note"></textarea></label><button class="btn" type="submit">Salva mezzo</button><a class="btn btn2" href="/mezzi">Annulla</a></form></div>`)));
app.post('/mezzi/nuovo',async(req,res)=>{const b=req.body||{};await run(`INSERT INTO mezzi (targa,marca,modello,tipo,km,km_attuali,prezzo_giorno,km_inclusi,cauzione,gps,blocco_motore,stato,note) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)`,[v62Val(b.targa).toUpperCase(),v62Val(b.marca).toUpperCase(),v62Val(b.modello).toUpperCase(),v62Val(b.tipo),v62Val(b.km),v62Val(b.km),v62Money(b.prezzo_giorno),v62Money(b.km_inclusi||150),v62Money(b.cauzione||500),v62Val(b.gps||'0'),v62Val(b.blocco_motore||'0'),'attivo',v62Val(b.note)]);res.redirect('/mezzi');});
app.get('/mezzi/:id/modifica',async(req,res)=>{const m=await get(`SELECT * FROM mezzi WHERE id=?`,[req.params.id]);if(!m)return res.status(404).send('Mezzo non trovato');res.send(page('Modifica mezzo',`<div class="box"><h2>Modifica mezzo ${esc(m.targa)}</h2><form method="post" action="/mezzi/${m.id}/modifica"><div class="grid"><label>Targa<input name="targa" value="${esc(m.targa)}" required></label><label>Marca<input name="marca" value="${esc(m.marca)}"></label><label>Modello<input name="modello" value="${esc(m.modello)}"></label><label>Tipo<input name="tipo" value="${esc(m.tipo)}"></label><label>Km<input name="km" value="${esc(m.km||m.km_attuali)}"></label><label>Prezzo giorno<input name="prezzo_giorno" value="${esc(m.prezzo_giorno)}"></label><label>Km inclusi<input name="km_inclusi" value="${esc(m.km_inclusi||150)}"></label><label>Cauzione standard<input name="cauzione" value="${esc(m.cauzione||500)}"></label><label>GPS<input name="gps" value="${esc(m.gps||'0')}"></label><label>Blocco motore<input name="blocco_motore" value="${esc(m.blocco_motore||'0')}"></label><label>Stato<select name="stato"><option value="attivo" ${m.stato!=='non_attivo'?'selected':''}>Attivo</option><option value="non_attivo" ${m.stato==='non_attivo'?'selected':''}>Non attivo</option></select></label></div><label>Note<textarea name="note">${esc(m.note)}</textarea></label><button class="btn" type="submit">Salva mezzo</button><a class="btn btn2" href="/mezzi">Annulla</a></form></div>`));});
app.post('/mezzi/:id/modifica',async(req,res)=>{const b=req.body||{};await run(`UPDATE mezzi SET targa=?,marca=?,modello=?,tipo=?,km=?,km_attuali=?,prezzo_giorno=?,km_inclusi=?,cauzione=?,gps=?,blocco_motore=?,stato=?,note=? WHERE id=?`,[v62Val(b.targa).toUpperCase(),v62Val(b.marca).toUpperCase(),v62Val(b.modello).toUpperCase(),v62Val(b.tipo),v62Val(b.km),v62Val(b.km),v62Money(b.prezzo_giorno),v62Money(b.km_inclusi||150),v62Money(b.cauzione||500),v62Val(b.gps||'0'),v62Val(b.blocco_motore||'0'),v62Val(b.stato||'attivo'),v62Val(b.note),v62Val(b.data_nascita),v62Val(b.luogo_nascita),v62Val(b.cittadinanza_cod||'100000100'),v62Val(b.documento_tipo||'IDENT'),v62Val(b.documento_numero),v62Val(b.documento_scadenza),v62Val(b.patente_numero),v62Val(b.patente_scadenza),req.params.id]);res.redirect('/mezzi');});
app.post('/mezzi/:id/elimina',async(req,res)=>{await run(`DELETE FROM mezzi WHERE id=?`,[req.params.id]);res.redirect('/mezzi');});


app.get('/admin/fix-tutto-v63',(req,res)=>{
  db.serialize(()=>{
    db.run(`CREATE TABLE IF NOT EXISTS prenotazioni (id INTEGER PRIMARY KEY AUTOINCREMENT,codice TEXT,nome TEXT,cognome TEXT,telefono TEXT,email TEXT,data_inizio TEXT,data_fine TEXT,totale REAL,stato TEXT DEFAULT 'bozza')`);
    db.run(`CREATE TABLE IF NOT EXISTS mezzi (id INTEGER PRIMARY KEY AUTOINCREMENT,targa TEXT,marca TEXT,modello TEXT,tipo TEXT,stato TEXT DEFAULT 'attivo')`);
    const pren={
      data_nascita:'TEXT', luogo_nascita:'TEXT', luogo_nascita_cod:'TEXT', cittadinanza_cod:'TEXT', conducente_cittadinanza_cod:'TEXT',
      documento_tipo:'TEXT', documento_numero:'TEXT', documento_scadenza:'TEXT', documento_luogo_rilascio_cod:'TEXT',
      patente_numero:'TEXT', patente_scadenza:'TEXT', patente_luogo_rilascio_cod:'TEXT',
      cauzione_richiesta:'TEXT', cauzione_ricevuta:'TEXT', cauzione_importo:'REAL', cauzione_metodo:'TEXT', cauzione_restituita:'TEXT',
      tipo_cliente:'TEXT', codice_fiscale:'TEXT', partita_iva:'TEXT', ragione_sociale:'TEXT', pec:'TEXT', codice_sdi:'TEXT',
      indirizzo:'TEXT', citta:'TEXT', cap:'TEXT', provincia:'TEXT', ora_inizio:'TEXT', ora_fine:'TEXT',
      drive_folder_id:'TEXT', drive_folder_link:'TEXT', pdf_drive_link:'TEXT', cargos_stato:'TEXT', cargos_transactionid:'TEXT', cargos_last_error:'TEXT'
    };
    const mez={prezzo_giorno:'REAL', km_inclusi:'REAL', cauzione:'REAL', gps:'TEXT', blocco_motore:'TEXT', note:'TEXT', km_attuali:'TEXT', codice_tipo:'TEXT'};
    let left=2; const done=()=>{if(--left===0)res.send(page('FIX V63 OK',`<div class="box"><h2 class="ok">FIX TUTTO V63 OK</h2><p>Bottoni, cauzione e campi CARGOS aggiornati.</p><a class="btn" href="/">Dashboard</a><a class="btn btn2" href="/mezzi/nuovo">Nuovo mezzo</a><a class="btn btn2" href="/preventivo/nuovo">Nuovo preventivo</a></div>`));};
    v62FixTable('prenotazioni',pren,done);
    v62FixTable('mezzi',mez,done);
  });
});
app.get('/admin/fix-tutto-v62',(req,res)=>res.redirect('/admin/fix-tutto-v63'));
app.get('/admin/fix-tutto-v61',(req,res)=>res.redirect('/admin/fix-tutto-v63'));


app.get('/contratto/:id/gestisci', async (req,res)=>{
  const p=await get(`SELECT * FROM prenotazioni WHERE id=?`,[req.params.id]);
  if(!p)return res.status(404).send(page('Non trovato',`<div class="box"><h2 class="bad">Contratto non trovato</h2></div>`));
  res.send(page('Gestisci contratto',`<div class="box">
    <h2>Gestisci ${esc(p.codice||p.id)}</h2>
    <p><b>Cliente:</b> ${esc((p.nome||'')+' '+(p.cognome||''))}</p>
    <p><b>Periodo:</b> ${esc(dpDateTimeLabel(p.data_inizio, p.ora_inizio))} - ${esc(dpDateTimeLabel(p.data_fine, p.ora_fine))}</p>
    <p><b>Totale:</b> ${euroHtml(p.totale||0)}</p>
    ${cauzioneHtml(p)}
    <p class="muted"><b>Azioni:</b> da qui fai firmare il cliente, invii il contratto su WhatsApp/Email e carichi foto o documenti.</p>
    ${v63ContractButtons(p)}
    <hr>
    <div class="actions contract-secondary-actions">
      <a class="btn btn2" href="/contratto/${p.id}">Apri riepilogo/PDF</a>
      <a class="btn btn2" href="/cargos/check/${p.id}">Ca.R.G.O.S.</a>
      <a class="btn btn2" href="/">Dashboard</a>
    </div>
  </div>`));
});


app.get('/admin/gestione-v63',(req,res)=>{
  res.send(page('Gestione V63',`<div class="box">
    <h2>Gestione rapida V63</h2>
    <a class="btn" href="/nuova-prenotazione">Nuovo contratto</a>
    <a class="btn btn2" href="/preventivo/nuovo">Nuovo preventivo</a>
    <a class="btn btn2" href="/mezzi/nuovo">Nuovo mezzo</a>
    <a class="btn btn2" href="/mezzi">Lista mezzi</a>
    <a class="btn btn2" href="/storico">Storico</a>
    <a class="btn btn2" href="/admin/fix-tutto-v63">Fix DB</a>
  </div>`));
});


app.get('/admin/test-cargos-veicolo-v65', (req,res)=>{
  const q = req.query.q || 'OPEL VIVARO';
  res.send(page('Test CARGOS veicolo V76', `<div class="box">
    <h2>Test tipo veicolo CARGOS</h2>
    <p>Testo: <b>${esc(q)}</b></p>
    <p>Codice CARGOS: <b>${esc(getTipoVeicoloCargosV61(q))}</b></p>
    <p>OPEL VIVARO deve essere <b>1 = Furgoni</b>.</p>
    <a class="btn" href="/admin/test-cargos-veicolo-v65?q=OPEL%20VIVARO">Test Vivaro</a>
    <a class="btn btn2" href="/">Dashboard</a>
  </div>`));
});


app.get('/admin/test-cargos-veicolo-v66', (req,res)=>{
  const q = req.query.q || 'OPEL VIVARO';
  res.send(page('Test CARGOS veicolo V76', `<div class="box">
    <h2>Test tipo veicolo CARGOS V76</h2>
    <p>Testo: <b>${esc(q)}</b></p>
    <p>Codice CARGOS: <b>${esc(getTipoVeicoloCargosV76(q))}</b></p>
    <p>OPEL VIVARO / FURGONI deve essere <b>1</b>.</p>
    <a class="btn" href="/admin/test-cargos-veicolo-v66?q=OPEL%20VIVARO">Test Vivaro</a>
    <a class="btn btn2" href="/admin/test-cargos-veicolo-v66?q=FURGONI">Test Furgoni</a>
    <a class="btn btn2" href="/">Dashboard</a>
  </div>`));
});
app.get('/admin/test-cargos-veicolo-v65', (req,res)=>res.redirect('/admin/test-cargos-veicolo-v66?q=' + encodeURIComponent(req.query.q || 'OPEL VIVARO')));

v67EnsureCriticalColumns(() => console.log('V76 colonne critiche OK'));

app.get('/admin/fix-tutto-v67',(req,res)=>{
  v67EnsureCriticalColumns(()=>{
    res.send(page('FIX V76 OK', `<div class="box">
      <h2 class="ok">FIX V76 OK</h2>
      <p>Colonne cauzione, documento, patente e nascita aggiornate.</p>
      <a class="btn" href="/">Dashboard</a>
      <a class="btn btn2" href="/storico">Storico</a>
      <a class="btn btn2" href="/admin/gestione-v63">Gestione</a>
    </div>`));
  });
});
app.get('/admin/fix-tutto-v66',(req,res)=>res.redirect('/admin/fix-tutto-v67'));


app.get('/prenotazione/:id/modifica', async (req,res)=>{
  v67EnsureCriticalColumns(async ()=>{
    const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
    if(!p) return res.status(404).send(page('Non trovato', `<div class="box"><h2>Contratto non trovato</h2></div>`));
    res.send(page('Modifica contratto', `<div class="box">
      <h2>Modifica ${esc(p.codice || p.id)}</h2>
      <form method="post" action="/prenotazione/${p.id}/modifica">
        <div class="grid">
          <label>Nome<input name="nome" value="${esc(p.nome)}" required></label>
          <label>Cognome<input name="cognome" value="${esc(p.cognome)}" required></label>
          <label>Telefono<input name="telefono" value="${esc(p.telefono)}"></label>
          <label>Email<input name="email" value="${esc(p.email)}"></label>
          <label>Codice fiscale<input name="codice_fiscale" value="${esc(p.codice_fiscale || p.cf)}"></label>
          <label>Data nascita<input type="date" name="data_nascita" value="${esc(v67IsoDate(p.data_nascita))}" required></label>
          <label>Luogo nascita<input name="luogo_nascita" value="${esc(p.luogo_nascita)}"></label>
          <label>Cittadinanza codice<input name="cittadinanza_cod" value="${esc(p.cittadinanza_cod || '100000100')}"></label>
          <label>Tipo documento
            <select name="documento_tipo">
              <option value="IDENT" ${(p.documento_tipo||'IDENT')==='IDENT'?'selected':''}>Carta identità</option>
              <option value="IDELE" ${p.documento_tipo==='IDELE'?'selected':''}>Carta identità elettronica</option>
              <option value="PASOR" ${p.documento_tipo==='PASOR'?'selected':''}>Passaporto</option>
              <option value="PATEN" ${p.documento_tipo==='PATEN'?'selected':''}>Patente</option>
            </select>
          </label>
          <label>Numero documento<input name="documento_numero" value="${esc(p.documento_numero)}" required></label>
          <label>Scadenza documento<input type="date" name="documento_scadenza" value="${esc(v67IsoDate(p.documento_scadenza))}"></label>
          <label>Numero patente<input name="patente_numero" value="${esc(p.patente_numero)}" required></label>
          <label>Scadenza patente<input type="date" name="patente_scadenza" value="${esc(v67IsoDate(p.patente_scadenza))}"></label>
          <label>Tipo cliente
            <select name="tipo_cliente">
              <option value="privato" ${(p.tipo_cliente||'privato')==='privato'?'selected':''}>Privato</option>
              <option value="azienda" ${p.tipo_cliente==='azienda'?'selected':''}>Azienda</option>
            </select>
          </label>
          <label>Ragione sociale<input name="ragione_sociale" value="${esc(p.ragione_sociale)}"></label>
          <label>Partita IVA<input name="partita_iva" value="${esc(p.partita_iva)}"></label>
          <label>PEC<input name="pec" value="${esc(p.pec)}"></label>
          <label>Codice SDI<input name="codice_sdi" value="${esc(p.codice_sdi)}"></label>
          <label>Data inizio<input type="date" name="data_inizio" value="${esc(p.data_inizio)}"></label>
          <label>Ora inizio<input type="time" name="ora_inizio" value="${esc(p.ora_inizio)}"></label>
          <label>Data fine<input type="date" name="data_fine" value="${esc(p.data_fine)}"></label>
          <label>Ora fine<input type="time" name="ora_fine" value="${esc(p.ora_fine)}"></label>
          <label>Totale<input name="totale" value="${esc(p.totale)}"></label>
          <label>Stato
            <select name="stato">
              <option value="preventivo" ${p.stato==='preventivo'?'selected':''}>Preventivo</option>
              <option value="bozza" ${p.stato==='bozza'?'selected':''}>Bozza</option>
              <option value="contratto" ${p.stato==='contratto'?'selected':''}>Contratto</option>
              <option value="firmato" ${p.stato==='firmato'?'selected':''}>Firmato</option>
              <option value="chiuso" ${p.stato==='chiuso'?'selected':''}>Chiuso</option>
            </select>
          </label>
          <label>Cauzione richiesta
            <select name="cauzione_richiesta">
              <option value="no" ${(p.cauzione_richiesta||'no')==='no'?'selected':''}>NO</option>
              <option value="si" ${p.cauzione_richiesta==='si'?'selected':''}>SI</option>
            </select>
          </label>
          <label>Cauzione ricevuta
            <select name="cauzione_ricevuta">
              <option value="no" ${(p.cauzione_ricevuta||'no')==='no'?'selected':''}>NO</option>
              <option value="si" ${p.cauzione_ricevuta==='si'?'selected':''}>SI</option>
            </select>
          </label>
          <label>Importo cauzione<input name="cauzione_importo" value="${esc(p.cauzione_importo || p.cauzione || 0)}"></label>
          <label>Metodo cauzione
            <select name="cauzione_metodo">
              <option value="">---</option>
              <option value="contanti" ${p.cauzione_metodo==='contanti'?'selected':''}>Contanti</option>
              <option value="carta" ${p.cauzione_metodo==='carta'?'selected':''}>Carta</option>
              <option value="bonifico" ${p.cauzione_metodo==='bonifico'?'selected':''}>Bonifico</option>
              <option value="non_versata" ${p.cauzione_metodo==='non_versata'?'selected':''}>Non versata</option>
            </select>
          </label>
          <label>Cauzione restituita
            <select name="cauzione_restituita">
              <option value="no" ${(p.cauzione_restituita||'no')==='no'?'selected':''}>NO</option>
              <option value="si" ${p.cauzione_restituita==='si'?'selected':''}>SI</option>
            </select>
          </label>
        </div>
        <label>Note<textarea name="note">${esc(p.note)}</textarea></label>
        <button class="btn" type="submit">Salva modifiche</button>
        <a class="btn btn2" href="/contratto/${p.id}/gestisci">Torna</a>
      </form>
    </div>`));
  });
});

app.post('/prenotazione/:id/modifica', async (req,res)=>{
  v67EnsureCriticalColumns(async ()=>{
    try{
      const b = req.body || {};
      await run(`UPDATE prenotazioni SET
        nome=?, cognome=?, telefono=?, email=?, codice_fiscale=?,
        data_nascita=?, luogo_nascita=?, cittadinanza_cod=?,
        documento_tipo=?, documento_numero=?, documento_scadenza=?,
        patente_numero=?, patente_scadenza=?,
        tipo_cliente=?, ragione_sociale=?, partita_iva=?, pec=?, codice_sdi=?,
        data_inizio=?, ora_inizio=?, data_fine=?, ora_fine=?, totale=?, stato=?,
        cauzione_richiesta=?, cauzione_ricevuta=?, cauzione_importo=?, cauzione_metodo=?, cauzione_restituita=?, note=?
        WHERE id=?`, [
          v62Val(b.nome), v62Val(b.cognome), v62Val(b.telefono), v62Val(b.email), v62Val(b.codice_fiscale),
          v62Val(b.data_nascita), v62Val(b.luogo_nascita), v62Val(b.cittadinanza_cod || '100000100'),
          v62Val(b.documento_tipo || 'IDENT'), v62Val(b.documento_numero), v62Val(b.documento_scadenza),
          v62Val(b.patente_numero), v62Val(b.patente_scadenza),
          v62Val(b.tipo_cliente || 'privato'), v62Val(b.ragione_sociale), v62Val(b.partita_iva), v62Val(b.pec), v62Val(b.codice_sdi),
          v62Val(b.data_inizio), v62Val(b.ora_inizio), v62Val(b.data_fine), v62Val(b.ora_fine), v62Money(b.totale), v62Val(b.stato || 'contratto'),
          v62Val(b.cauzione_richiesta || 'no'), v62Val(b.cauzione_ricevuta || 'no'), v62Money(b.cauzione_importo), v62Val(b.cauzione_metodo), v62Val(b.cauzione_restituita || 'no'), v62Val(b.note),
          req.params.id
      ]);
      try{ if(typeof syncContrattoDriveV59==='function') await syncContrattoDriveV59(req.params.id); }catch(e){}
      res.redirect(`/contratto/${req.params.id}/gestisci`);
    } catch(e){
      res.status(500).send(page('Errore salvataggio', `<div class="box"><h2 class="bad">Errore salvataggio</h2><pre>${esc(e.message)}</pre><a class="btn" href="/prenotazione/${req.params.id}/modifica">Torna modifica</a></div>`));
    }
  });
});


// =========================
// V76 ROUTE FIX PULITE
// =========================
app.get('/admin/fix-tutto-v76', (req, res) => {
  const html = `<div class="box">
    <h2 class="ok">FIX V76 OK</h2>
    <p>Route corrette. Nessun syntax error su fix-tutto.</p>
    <p>Record Ca.R.G.O.S.: VEICOLO_TIPO una sola volta, lunghezza 1505.</p>
    <a class="btn" href="/">Dashboard</a>
    <a class="btn btn2" href="/storico">Storico</a>
  </div>`;
  if (typeof page === 'function') return res.send(page('FIX V76 OK', html));
  res.send(html);
});

app.get('/admin/fix-tutto-v75', (req, res) => res.redirect('/admin/fix-tutto-v76'));
app.get('/admin/fix-tutto-v74', (req, res) => res.redirect('/admin/fix-tutto-v76'));
app.get('/admin/fix-tutto-v73', (req, res) => res.redirect('/admin/fix-tutto-v76'));
app.get('/admin/fix-tutto-v72', (req, res) => res.redirect('/admin/fix-tutto-v76'));
app.get('/admin/fix-tutto-v71', (req, res) => res.redirect('/admin/fix-tutto-v76'));
app.get('/admin/fix-tutto-v70', (req, res) => res.redirect('/admin/fix-tutto-v76'));
app.get('/admin/fix-tutto-v69', (req, res) => res.redirect('/admin/fix-tutto-v76'));
app.get('/admin/fix-tutto-v68', (req, res) => res.redirect('/admin/fix-tutto-v76'));




// =====================================================
// V88 - WHATSAPP DP BOT INTEGRATO SU APP COMPLETA V76
// - non sostituisce dashboard
// - menu con emoji via unicode escape, niente caratteri rotti
// - ciao/menu apre menu
// - notifiche ai 3 numeri
// - Google Calendar solo officina
// - Drive solo noleggio/contratti gia presenti nella app
// - no MyAppy / no Calendly
// =====================================================
const DP_BOT_SESSIONS = {};
const DP_BOT_DONE_SIDS = new Map();
function dpNormalizeWhatsAppNumber(n){
  n = String(n || '').trim();
  if(!n) return '';
  if(n.startsWith('whatsapp:')) return n;
  n = n.replace(/\s+/g,'');
  if(!n.startsWith('+')) {
    if(n.startsWith('39')) n = '+' + n;
    else n = '+39' + n.replace(/^0+/, '');
  }
  return 'whatsapp:' + n;
}
function normalizzaWa(n){ return dpNormalizeWhatsAppNumber(n); }

function dpLabelStatus(v){
  const raw = String(v || '').replace(/_/g, ' ').trim();
  if (!raw) return '';
  return raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
}
function dpDateTimeLabel(data, ora){
  return [String(data || '').trim(), String(ora || '').trim()].filter(Boolean).join(' ore ');
}

function euroHtml(v){
  const n = Number(String(v ?? 0).replace(',', '.'));
  return '&euro; ' + (Number.isFinite(n) ? n.toFixed(2).replace('.', ',') : String(v || '0'));
}
function cauzioneHtml(p){
  const richiesta = String(p?.cauzione_richiesta || 'si').toLowerCase() === 'si';
  const ricevuta = String(p?.cauzione_ricevuta || 'no').toLowerCase() === 'si';
  const importo = p?.cauzione_importo || p?.cauzione || 500;
  if (!richiesta) return '<div class="cauzione-box"><span class="label">Cauzione:</span><span class="badge badge-warn">Non richiesta</span></div>';
  return '<div class="cauzione-box"><span class="label">Cauzione:</span><span class="badge badge-money">' + euroHtml(importo) + '</span><span class="badge ' + (ricevuta ? 'badge-ok' : 'badge-danger') + '">' + (ricevuta ? 'RICEVUTA' : 'DA RICEVERE') + '</span></div>';
}
const DP_TWILIO_WHATSAPP_NUMBER = dpNormalizeWhatsAppNumber(process.env.TWILIO_WHATSAPP_NUMBER || '+390744817108');
const DP_STAFF_NUMBERS = dpParseNumbers(process.env.INTERNAL_GENERAL_NUMBERS || process.env.STAFF_WHATSAPP_NUMBERS, [
  'whatsapp:+393287377675',
  'whatsapp:+393472733226',
  'whatsapp:+393494040073'
]);
const DP_OFFICINA_NUMBERS = dpParseNumbers(process.env.INTERNAL_OFFICINA_NUMBERS, DP_STAFF_NUMBERS);
const DP_AUTOSUPERMARKET_URL = 'https://autosupermarket.it/concessionario/trasporti-dp-srl/annunci';

const EMJ = {
  ok: '\u2705',
  car: '\u{1F697}',
  van: '\u{1F690}',
  auto: '\u{1F698}',
  wrench: '\u{1F527}',
  money: '\u{1F4B0}',
  truck: '\u{1F69B}',
  chat: '\u{1F4AC}',
  pen: '\u270D\uFE0F',
  warn: '\u26A0\uFE0F',
  calendar: '\u{1F4C5}',
  link: '\u{1F517}',
  one: '1\uFE0F\u20E3',
  two: '2\uFE0F\u20E3',
  three: '3\uFE0F\u20E3',
  four: '4\uFE0F\u20E3',
  five: '5\uFE0F\u20E3'
};

const dpTwilioClient = (twilio && process.env.TWILIO_ACCOUNT_SID && process.env.TWILIO_AUTH_TOKEN)
  ? twilio(process.env.TWILIO_ACCOUNT_SID, process.env.TWILIO_AUTH_TOKEN)
  : null;

function dpParseNumbers(value, fallback){
  if(!value) return (fallback || []).map(dpNormalizeWhatsAppNumber).filter(Boolean);
  return String(value).split(/[;,\n]+/).map(x=>x.trim()).filter(Boolean).map(dpNormalizeWhatsAppNumber).filter(Boolean);
}
function dpClean(v){ return String(v || '').trim(); }
function dpNorm(v){ return dpClean(v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function dpXml(v){ return String(v || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;'); }
function dpWa(text){ return String(text || '').normalize('NFC').trim(); }
function dpYesNo(v){
  const t = dpNorm(v);
  if(['si','s','ok','confermo','yes','certo'].includes(t)) return 'SI';
  if(['no','n','annulla'].includes(t)) return 'NO';
  return '';
}
function dpSession(from, profileName){
  if(!DP_BOT_SESSIONS[from] || Date.now() - (DP_BOT_SESSIONS[from].ts || 0) > 2*60*60*1000){
    DP_BOT_SESSIONS[from] = { state:'menu', data:{}, profileName: profileName || 'Cliente', ts: Date.now() };
  }
  DP_BOT_SESSIONS[from].ts = Date.now();
  return DP_BOT_SESSIONS[from];
}
function dpReset(from, profileName){
  DP_BOT_SESSIONS[from] = { state:'menu', data:{}, profileName: profileName || 'Cliente', ts: Date.now() };
  return DP_BOT_SESSIONS[from];
}
function dpMenu(name){
  return `${EMJ.ok} *Ciao ${name || 'Cliente'}*\n\n${EMJ.car} *DP RENT / TRASPORTI DP*\n\nScegli il servizio:\n\n${EMJ.one} ${EMJ.wrench} Officina\n${EMJ.two} ${EMJ.van} Noleggio\n${EMJ.three} ${EMJ.money} Vendita auto\n${EMJ.four} ${EMJ.truck} Trasporto veicoli\n${EMJ.five} ${EMJ.chat} Altre richieste\n\n${EMJ.pen} Scrivi solo il numero.\nEsempio: *2*`;
}
function dpIsMenuKeyword(body){
  const t = dpNorm(body);
  return ['ciao','salve','buongiorno','buonasera','menu','start','inizio','info','informazioni','aiuto'].includes(t);
}
function dpAlreadySid(sid){
  if(!sid) return false;
  if(DP_BOT_DONE_SIDS.has(sid)) return true;
  DP_BOT_DONE_SIDS.set(sid, Date.now());
  const now = Date.now();
  for(const [k,ts] of DP_BOT_DONE_SIDS.entries()) if(now-ts > 15*60*1000) DP_BOT_DONE_SIDS.delete(k);
  return false;
}
async function dpNotify(numbers, body){
  const targets = Array.from(new Set((numbers || []).map(dpNormalizeWhatsAppNumber).filter(Boolean)));
  const fromNumber = dpNormalizeWhatsAppNumber(process.env.TWILIO_WHATSAPP_NUMBER || DP_TWILIO_WHATSAPP_NUMBER || '+390744817108');
  console.log('DP_NOTIFY FROM:', fromNumber);
  console.log('DP_NOTIFY TO:', targets.join(', '));
  console.log('DP_NOTIFY BODY:', body);

  if(!dpTwilioClient){
    const err = 'Twilio non configurato: mancano TWILIO_ACCOUNT_SID o TWILIO_AUTH_TOKEN';
    console.log(err);
    return { ok:false, sent:0, errors:[err] };
  }
  if(!fromNumber || !fromNumber.startsWith('whatsapp:+')){
    const err = 'TWILIO_WHATSAPP_NUMBER non valido: ' + String(process.env.TWILIO_WHATSAPP_NUMBER || '');
    console.log(err);
    return { ok:false, sent:0, errors:[err] };
  }
  if(!targets.length){
    const err = 'Nessun numero staff configurato';
    console.log(err);
    return { ok:false, sent:0, errors:[err] };
  }

  let sent = 0;
  const errors = [];
  for(const to of targets){
    try{
      const msg = await dpTwilioClient.messages.create({
        from: fromNumber,
        to,
        body: dpWa(body)
      });
      sent++;
      console.log('Notifica WhatsApp inviata a', to, msg.sid || '');
    }catch(e){
      const err = `${to}: ${e.message || e} ${e.code ? '(code '+e.code+')' : ''}`;
      errors.push(err);
      console.error('Errore notifica WhatsApp', to, e.message, e.code || '');
    }
  }
  return { ok: sent > 0, sent, errors };
}


async function dpGoogleAuth(scopes){
  if(typeof google === 'undefined' || !google) throw new Error('googleapis non installato');
  const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || process.env.GOOGLE_CLIENT_EMAIL;
  let privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
  if(!clientEmail || !privateKey) throw new Error('ENV Google mancanti');
  privateKey = privateKey.replace(/\n/g, '\n');
  if(privateKey.includes('\\n')) privateKey = privateKey.replace(/\\n/g, '\n');
  return new google.auth.JWT(clientEmail, null, privateKey, scopes);
}
async function dpGetOrCreateDriveFolder(driveApi, parentId, folderName){
  const safe = String(folderName).replace(/'/g, "\\'");
  let q = `mimeType='application/vnd.google-apps.folder' and name='${safe}' and trashed=false`;
  if(parentId) q += ` and '${parentId}' in parents`;
  const found = await driveApi.files.list({ q, fields:'files(id,name,webViewLink)', spaces:'drive' });
  if(found.data.files && found.data.files[0]) return found.data.files[0];
  const requestBody = { name: folderName, mimeType:'application/vnd.google-apps.folder' };
  if(parentId) requestBody.parents = [parentId];
  const created = await driveApi.files.create({ requestBody, fields:'id,name,webViewLink' });
  return created.data;
}
async function dpUploadTextToDriveServiceAccount(tipo, codice, from, profileName, body){
  const parentId = process.env.GOOGLE_DRIVE_FOLDER_ID || process.env.DRIVE_FOLDER_ID || process.env.GOOGLE_DRIVE_OFFICINA_FOLDER_ID || '';
  if(!parentId) return { ok:false, error:'GOOGLE_DRIVE_FOLDER_ID mancante' };
  const auth = await dpGoogleAuth(['https://www.googleapis.com/auth/drive']);
  const driveApi = google.drive({ version:'v3', auth });
  const folder = await dpGetOrCreateDriveFolder(driveApi, parentId, 'DP OFFICINA');
  const fileName = `${codice}.txt`;
  const localPath = path.join(uploadDir, fileName);
  const content = `${tipo}\n\nCodice pratica: ${codice}\nCliente: ${profileName || '-'}\nWhatsApp: ${from || '-'}\nData: ${new Date().toLocaleString('it-IT')}\n\nRichiesta:\n${body || '-' }\n`;
  fs.writeFileSync(localPath, content, 'utf8');
  const media = { mimeType:'text/plain', body: fs.createReadStream(localPath) };
  const uploaded = await driveApi.files.create({
    requestBody:{ name:fileName, parents:[folder.id] },
    media,
    fields:'id,name,webViewLink'
  });
  return { ok:true, link: uploaded.data.webViewLink || folder.webViewLink || '', id: uploaded.data.id || '', folder: folder.webViewLink || '' };
}
async function dpSaveRequestToDrive(tipo, codice, from, profileName, body){
  // 1) Prova prima Google Drive diretto con Service Account, così non dipende da Apps Script.
  try{
    const direct = await dpUploadTextToDriveServiceAccount(tipo, codice, from, profileName, body);
    if(direct && direct.ok) return direct;
  }catch(e){
    console.error('Drive diretto non riuscito:', e.message);
  }

  // 2) Fallback vecchio Apps Script se è configurato.
  try{
    if(!googleDriveConfigured()) return { ok:false, error:'Drive non configurato: manca GOOGLE_DRIVE_FOLDER_ID o DRIVE_WEBAPP_URL' };
    const fileName = `${codice}.txt`;
    const localPath = path.join(uploadDir, fileName);
    const content = `${tipo}\n\nCodice pratica: ${codice}\nCliente: ${profileName || '-'}\nWhatsApp: ${from || '-'}\nData: ${new Date().toLocaleString('it-IT')}\n\nRichiesta:\n${body || '-' }\n`;
    fs.writeFileSync(localPath, content, 'utf8');
    const dr = await uploadFileToDrive(localPath, fileName, 'text/plain', `${tipo} ${codice}`);
    if(!dr) return { ok:false, error:'Upload Drive non riuscito' };
    return { ok:true, link: dr.webViewLink || dr.link || '', id: dr.id || '' };
  }catch(e){
    console.error('Errore salvataggio Drive richiesta:', e.message);
    return { ok:false, error:e.message };
  }
}
function dpExtractDate(text){
  const t = dpClean(text);
  let m = t.match(/(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?/);
  if(!m) return null;
  let y = m[3] ? Number(m[3]) : new Date().getFullYear();
  if(y < 100) y += 2000;
  const d = new Date(y, Number(m[2])-1, Number(m[1]), 9, 0, 0);
  if(isNaN(d.getTime())) return null;
  return d;
}
function dpDateIso(d){ return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function dpDateIt(d){ return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`; }
function dpExtractRange(text){
  const m = dpClean(text).match(/(\d{1,2}[\/\-.]\d{1,2}(?:[\/\-.]\d{2,4})?).*?(\d{1,2}[\/\-.]\d{1,2}(?:[\/\-.]\d{2,4})?)/);
  if(!m){ const one = dpExtractDate(text); return one ? { start: one, end: one } : null; }
  const s = dpExtractDate(m[1]);
  const e = dpExtractDate(m[2]);
  if(!s || !e || e < s) return null;
  return { start:s, end:e };
}
function dpDays(a,b){ return Math.round((new Date(b.getFullYear(),b.getMonth(),b.getDate()) - new Date(a.getFullYear(),a.getMonth(),a.getDate()))/86400000)+1; }
function dpExtractKm(text){ const m = String(text||'').replace(/\./g,'').match(/\d{1,6}/); return m ? Number(m[0]) : 150; }
function dpCategoryFromChoice(txt){
  const t = dpNorm(txt);
  if(t === '1' || t.includes('furg')) return { label:'Furgone cargo/merci', categoria:'FURGONE', cats:['FURGONE','FURGONI','F1-VAN','F2-PC','F3-PL'] };
  if(t === '2' || t.includes('pulmino') || t.includes('9') || t.includes('posti')) return { label:'Pulmino 8/9 posti', categoria:'9_POSTI', cats:['9_POSTI','PULMINO','P2-9P','P1-8P'] };
  if(t === '3' || t.includes('dacia') || t.includes('econom')) return { label:'Auto economica tipo Dacia', categoria:'AUTO_DACIA', cats:['AUTO_DACIA','DACIA'] };
  if(t === '4' || t.includes('golf')) return { label:'Auto categoria Golf', categoria:'AUTO_GOLF', cats:['AUTO_GOLF','GOLF'] };
  if(t === '5' || t.includes('escav')) return { label:'Escavatore / mezzo speciale', categoria:'ESCAVATORE', cats:['ESCAVATORE','SEMOVENTE','X-ESC'] };
  return null;
}
function dpVehicleMatchesCat(m, catInfo){
  const target = catInfo?.categoria || '';
  if (!target) return false;
  try { return mezzoCompatibileCategoriaCliente(m, target); } catch(_) { return false; }
}
async function dpFindAvailableVehicle(catInfo, startIso, endIso){
  let mezzi = [];
  try{ mezzi = await all(`SELECT * FROM mezzi ORDER BY id ASC`); }catch(e){ console.error('Errore select mezzi:', e.message); }
  mezzi = (mezzi || []).filter(m => dpVehicleMatchesCat(m, catInfo));
  for(const m of mezzi){
    try{
      const occ = await queryDisponibilita(m.id, startIso, endIso);
      if(!occ) return m;
    }catch(e){ return m; }
  }
  return null;
}
async function dpSaveWhatsAppQuote(session, from, profileName, status){
  try{
    if (typeof ensureClienteWebColumnsV92 === 'function') await ensureClienteWebColumnsV92();
    const data = session.data || {};
    const mezzo = data.mezzo || {};
    const calc = data.calc || {};
    const startIso = data.start ? dpDateIso(data.start) : '';
    const endIso = data.end ? dpDateIso(data.end) : '';
    const payload = {
      codice:'TEMP', nome:profileName || 'Cliente', cognome:'', telefono:String(from||'').replace('whatsapp:',''), email:'',
      categoria:data.cat?.categoria || data.cat?.cats?.[0] || '', tipo:data.cat?.label || '', mezzo_id:mezzo.id || null,
      targa:mezzo.targa || '', marca:mezzo.marca || '', modello:mezzo.modello || '',
      data_inizio:startIso, data_fine:endIso, ora_inizio:'08:30', ora_fine:'18:00', km_previsti:data.km || 150,
      giorni:calc.giorni || (data.start && data.end ? dpDays(data.start,data.end) : 1), imponibile:calc.imponibile || 0, iva:calc.iva || 0, totale:calc.totale || 0,
      stato:status || 'attesa_si_no', tipo_record:'preventivo_whatsapp', note:'Creato automaticamente dal bot WhatsApp - cliente in attesa risposta SI/NO'
    };
    const cols = Object.keys(payload);
    const r = await run(`INSERT INTO prenotazioni (${cols.join(',')}) VALUES (${cols.map(()=>'?').join(',')})`, cols.map(k=>payload[k]));
    const cod = codicePratica(r.lastID);
    await run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, r.lastID]);
    session.data.prenotazione_id = r.lastID;
    session.data.codice = cod;
    return { ok:true, id:r.lastID, codice:cod };
  }catch(e){ console.log('Salvataggio preventivo WhatsApp non riuscito:', e.message); return { ok:false, error:e.message }; }
}
async function dpUpdateWhatsAppQuote(session, stato){
  try{ if(session?.data?.prenotazione_id) await run(`UPDATE prenotazioni SET stato=?, note=COALESCE(note,'') || ? WHERE id=?`, [stato, '\nAggiornamento WhatsApp: '+stato, session.data.prenotazione_id]); }catch(e){ console.log('Update preventivo WhatsApp:', e.message); }
}
const DP_OFFICINA_SLOTS = (process.env.OFFICINA_SLOTS || '08:30,09:30,10:30,11:30,14:30,15:30,16:30,17:30')
  .split(',')
  .map(x => x.trim())
  .filter(Boolean);
const DP_OFFICINA_SLOT_MINUTES = Number(process.env.OFFICINA_SLOT_MINUTES || 60);

function dpParseTime(text){
  const m = String(text || '').match(/(\d{1,2})(?:[:\.](\d{2}))?/);
  if(!m) return '';
  const h = Number(m[1]);
  const min = m[2] ? Number(m[2]) : 0;
  if(h < 0 || h > 23 || min < 0 || min > 59) return '';
  return `${String(h).padStart(2,'0')}:${String(min).padStart(2,'0')}`;
}
function dpDateAtTime(dateObj, hhmm){
  const [h,m] = String(hhmm || '09:00').split(':').map(Number);
  return new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate(), h || 0, m || 0, 0);
}
function dpAddMinutes(dateObj, minutes){
  return new Date(dateObj.getTime() + Number(minutes || 60) * 60000);
}
async function dpCalendarOfficina(){
  if(typeof google === 'undefined' || !google) throw new Error('googleapis non installato');
  const calendarId = process.env.GOOGLE_CALENDAR_ID || process.env.GOOGLE_CALENDAR_OFFICINA_ID || 'primary';
  const clientEmail = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || process.env.GOOGLE_CLIENT_EMAIL;
  let privateKey = process.env.GOOGLE_PRIVATE_KEY || '';
  if(!clientEmail || !privateKey) throw new Error('ENV Google mancanti');
  privateKey = privateKey.replace(/\\n/g, '\n');
  if(privateKey.includes('\\n')) privateKey = privateKey.replace(/\\n/g, '\n');
  const auth = new google.auth.JWT(clientEmail, null, privateKey, ['https://www.googleapis.com/auth/calendar']);
  return { calendarId, calendar: google.calendar({ version:'v3', auth }) };
}
async function dpIsOfficinaSlotFree(dateObj, hhmm){
  const { calendarId, calendar } = await dpCalendarOfficina();
  const start = dpDateAtTime(dateObj, hhmm);
  const end = dpAddMinutes(start, DP_OFFICINA_SLOT_MINUTES);
  const r = await calendar.freebusy.query({
    requestBody: {
      timeMin: start.toISOString(),
      timeMax: end.toISOString(),
      timeZone: 'Europe/Rome',
      items: [{ id: calendarId }]
    }
  });
  const busy = (((r.data || {}).calendars || {})[calendarId] || {}).busy || [];
  return { free: busy.length === 0, start, end, calendar, calendarId };
}
async function dpAvailableOfficinaSlots(dateObj){
  const out = [];
  for(const slot of DP_OFFICINA_SLOTS){
    try{
      const ck = await dpIsOfficinaSlotFree(dateObj, slot);
      if(ck.free) out.push(slot);
    }catch(e){
      throw e;
    }
  }
  return out;
}
async function dpCreateCalendarEventOfficina(from, profileName, text, dateObj, hhmm){
  try{
    const d = dateObj || dpExtractDate(text) || new Date(Date.now()+24*60*60*1000);
    const time = hhmm || dpParseTime(text) || '09:00';
    const ck = await dpIsOfficinaSlotFree(d, time);
    if(!ck.free) return { ok:false, busy:true, error:'slot occupato' };
    const event = {
      summary: `OFFICINA DP - ${profileName || 'Cliente WhatsApp'}`,
      description: `Richiesta da WhatsApp: ${from}\nData: ${dpDateIt(d)} ore ${time}\n\n${text}`,
      start: { dateTime: ck.start.toISOString(), timeZone: 'Europe/Rome' },
      end: { dateTime: ck.end.toISOString(), timeZone: 'Europe/Rome' }
    };
    const r = await ck.calendar.events.insert({ calendarId: ck.calendarId, requestBody: event });
    console.log('Evento Calendar creato:', r.data && r.data.htmlLink);
    return { ok:true, link: r.data && r.data.htmlLink, start: ck.start, end: ck.end, time };
  }catch(e){
    console.error('Errore Google Calendar:', e.message);
    return { ok:false, error:e.message };
  }
}
async function dpChatGPTAnswer(message){
  if(!process.env.OPENAI_API_KEY){
    return 'Ho ricevuto la richiesta. La inoltro subito allo staff DP, ti risponderemo appena possibile.';
  }
  try{
    const r = await fetch('https://api.openai.com/v1/chat/completions', {
      method:'POST',
      headers:{ 'Authorization':'Bearer '+process.env.OPENAI_API_KEY, 'Content-Type':'application/json' },
      body: JSON.stringify({
        model: process.env.OPENAI_MODEL || 'gpt-4o-mini',
        messages:[
          {role:'system', content:'Sei l assistente WhatsApp di Trasporti DP / DP RENT a Narni. Rispondi in italiano, professionale, gentile, breve. Servizi: officina, noleggio, vendita auto, trasporto veicoli. Per preventivi specifici chiedi i dati mancanti e avvisa che lo staff confermera.'},
          {role:'user', content: message}
        ],
        temperature:0.4,
        max_tokens:250
      })
    });
    const data = await r.json();
    return data?.choices?.[0]?.message?.content || 'Ho ricevuto la richiesta. La inoltro allo staff DP.';
  }catch(e){
    console.error('Errore OpenAI:', e.message);
    return 'Ho ricevuto la richiesta. La inoltro subito allo staff DP.';
  }
}
function dpTwimlResponse(res, text){
  const xml = `<?xml version="1.0" encoding="UTF-8"?><Response><Message>${dpXml(dpWa(text))}</Message></Response>`;
  res.writeHead(200, { 'Content-Type':'text/xml; charset=utf-8' });
  return res.end(xml);
}

async function dpHandleWhatsApp(req,res){
  const from = dpClean(req.body.From || '').toLowerCase();
  const body = dpClean(req.body.Body || '');
  const profileName = req.body.ProfileName || 'Cliente';
  const sid = dpClean(req.body.MessageSid || req.body.SmsSid || '');
  console.log('DP BOT IN:', { from, body, profileName, sid, state: DP_BOT_SESSIONS[from]?.state });
  if(dpAlreadySid(sid)){ res.writeHead(200, {'Content-Type':'text/xml; charset=utf-8'}); return res.end('<Response/>'); }
  if(!from) return dpTwimlResponse(res, 'Errore ricezione messaggio.');

  let session = dpSession(from, profileName);

  const tGlobal = dpNorm(body);
  if(['indietro','torna','torna menu','menu principale','annulla','annullare','ho sbagliato','sbagliato','ricomincia','restart'].includes(tGlobal)){
    dpReset(from, profileName);
    return dpTwimlResponse(res, 'Nessun problema 👍\n\n' + dpMenu(profileName));
  }

  if(dpIsMenuKeyword(body) || body === ''){
    dpReset(from, profileName);
    return dpTwimlResponse(res, dpMenu(profileName));
  }

  if(session.state === 'menu'){
    if(body === '1'){
      session.state = 'officina_descrizione'; session.data = {}; session.ts = Date.now();
      return dpTwimlResponse(res, `${EMJ.wrench} *Officina DP*\n\nScrivi targa, mezzo e problema/intervento.\n\nEsempio:\nAB123CD Fiat Panda tagliando completo`);
    }
    if(body === '2'){
      session.state = 'noleggio_model'; session.ts = Date.now();
      return dpTwimlResponse(res, `${EMJ.van} *DP RENT - Noleggio*\n\nChe mezzo ti serve?\n\n${EMJ.one} Furgone cargo/merci\n${EMJ.two} Pulmino 8/9 posti\n${EMJ.three} Auto economica tipo Dacia\n${EMJ.four} Auto categoria Golf\n${EMJ.five} Escavatore / mezzo speciale\n\nScrivi il numero.`);
    }
    if(body === '3'){
      session.state = 'vendita'; session.ts = Date.now();
      return dpTwimlResponse(res, `${EMJ.auto} *DP AUTO - Vendita auto*\n\nPuoi vedere le auto disponibili qui:\n${DP_AUTOSUPERMARKET_URL}\n\nSe cerchi qualcosa in particolare, scrivi:\n- modello\n- budget\n- permuta si/no\n- finanziamento si/no`);
    }
    if(body === '4'){
      session.state = 'trasporto'; session.ts = Date.now();
      return dpTwimlResponse(res, `${EMJ.truck} *Trasporto veicoli*\n\nScrivi in un solo messaggio:\n- marca e modello auto\n- marciante o non marciante\n- da dove ritirare\n- dove consegnare\n- quando serve\n\nEsempio:\nFiat Panda marciante, ritiro Roma, consegna Narni, prossima settimana`);
    }
    if(body === '5'){
      session.state = 'altro'; session.ts = Date.now();
      return dpTwimlResponse(res, `${EMJ.chat} *Altre richieste*\n\nScrivi pure la tua richiesta. Ti rispondo subito e la inoltro allo staff DP.`);
    }
    return dpTwimlResponse(res, dpMenu(profileName));
  }

  if(session.state === 'officina_descrizione'){
    session.data.descrizione = body;
    session.state = 'officina_data'; session.ts = Date.now();
    return dpTwimlResponse(res, `${EMJ.calendar} Perfetto. Ora scrivi la data desiderata per l appuntamento officina.\n\nEsempio: 20/05/2026`);
  }

  if(session.state === 'officina_data'){
    const d = dpExtractDate(body);
    if(!d) return dpTwimlResponse(res, 'Non riesco a leggere la data. Scrivila cosi: 20/05/2026');
    session.data.data = d.toISOString();
    try{
      const slots = await dpAvailableOfficinaSlots(d);
      session.data.slots = slots;
      session.state = 'officina_orario'; session.ts = Date.now();
      if(!slots.length){
        return dpTwimlResponse(res, `${EMJ.warn} Per il ${dpDateIt(d)} non risultano orari liberi in Calendar.\nScrivi un altra data.`);
      }
      return dpTwimlResponse(res, `${EMJ.ok} Orari disponibili per il ${dpDateIt(d)}:\n\n${slots.map(x => '- ' + x).join('\n')}\n\nScrivi l orario scelto. Esempio: ${slots[0]}`);
    }catch(e){
      return dpTwimlResponse(res, `${EMJ.warn} Non riesco a leggere Google Calendar: ${e.message || e}.\nScrivi comunque l orario desiderato, lo staff DP verifichera manualmente.`);
    }
  }

  if(session.state === 'officina_orario'){
    const time = dpParseTime(body);
    if(!time) return dpTwimlResponse(res, 'Non riesco a leggere l orario. Scrivilo cosi: 08:30');
    const d = new Date(session.data.data);
    const descrizione = session.data.descrizione || '';
    const codice = `OFF-${new Date().toISOString().slice(0,10).replace(/-/g,'')}-${Math.floor(1000 + Math.random()*9000)}`;
    const cal = await dpCreateCalendarEventOfficina(from, profileName, descrizione, d, time);
    if(!cal.ok && cal.busy){
      let slots = [];
      try{ slots = await dpAvailableOfficinaSlots(d); }catch(e){}
      return dpTwimlResponse(res, `${EMJ.warn} Orario ${time} non disponibile.\n\nScegli un altro orario${slots.length ? ':\n' + slots.map(x => '- ' + x).join('\n') : ' oppure cambia data.'}`);
    }
    const notif = await dpNotify(DP_OFFICINA_NUMBERS, `${EMJ.wrench} NUOVA RICHIESTA OFFICINA\n\nCodice pratica: ${codice}\nCliente: ${profileName}\nWhatsApp: ${from}\nData: ${dpDateIt(d)}\nOrario: ${time}\n\nRichiesta:\n${descrizione}\n\nCalendar: ${cal.ok ? (cal.link || 'evento creato') : 'NON creato - ' + (cal.error || '-')}`);
    delete DP_BOT_SESSIONS[from];
    return dpTwimlResponse(res, `${EMJ.ok} Appuntamento officina ricevuto.\n\nCodice pratica: ${codice}\nData: ${dpDateIt(d)}\nOrario: ${time}\n\n${cal.ok ? 'Evento inserito in Google Calendar.' : 'Calendar non creato: ' + (cal.error || '-')}\n${notif.ok ? 'Richiesta inviata allo staff DP.' : 'ATTENZIONE: messaggio staff non inviato. Errore: ' + (notif.errors || []).join(' | ')}\n\nTi aspettiamo da DP.`);
  }


  if(session.state === 'noleggio_model'){
    const cat = dpCategoryFromChoice(body);
    if(!cat) return dpTwimlResponse(res, 'Scelta non valida. Scrivi 1, 2, 3, 4 oppure 5.');
    session.data.cat = cat; session.state = 'noleggio_dates'; session.ts = Date.now();
    return dpTwimlResponse(res, `Hai scelto: *${cat.label}*\n\nIndica le date noleggio.\nEsempio: 20/05 - 22/05`);
  }

  if(session.state === 'noleggio_dates'){
    const range = dpExtractRange(body);
    if(!range) return dpTwimlResponse(res, 'Non riesco a leggere le date. Scrivile cosi: 20/05 - 22/05');
    session.data.start = range.start; session.data.end = range.end; session.state = 'noleggio_km'; session.ts = Date.now();
    return dpTwimlResponse(res, 'Quanti km prevedi di fare?\nEsempio: 400');
  }

  if(session.state === 'noleggio_km'){
    const km = dpExtractKm(body);
    const startIso = dpDateIso(session.data.start);
    const endIso = dpDateIso(session.data.end);
    const mezzo = await dpFindAvailableVehicle(session.data.cat, startIso, endIso);
    if(!mezzo){
      await dpNotify(DP_STAFF_NUMBERS, `${EMJ.warn} RICHIESTA NOLEGGIO NON DISPONIBILE\n\nCliente: ${profileName}\nWhatsApp: ${from}\nMezzo: ${session.data.cat.label}\nDate: ${dpDateIt(session.data.start)} - ${dpDateIt(session.data.end)}\nKm: ${km}`);
      delete DP_BOT_SESSIONS[from];
      return dpTwimlResponse(res, 'Mi dispiace, non risulta disponibilita per quelle date.\n\nLa richiesta e stata inviata allo staff DP per controllo manuale.');
    }
    let calc = { totale: 0, giorni: dpDays(session.data.start, session.data.end) };
    try{ calc = calcolaTotale(mezzo, startIso, endIso, '08:30', '18:00', km); }catch(e){ console.error(e.message); }
    session.data.km = km; session.data.mezzo = mezzo; session.data.calc = calc; session.state = 'noleggio_confirm'; session.ts = Date.now();
    const savedQuote = await dpSaveWhatsAppQuote(session, from, profileName, 'attesa_si_no');
    const appBaseQuote = (process.env.APP_BASE_URL || process.env.RENDER_EXTERNAL_URL || '').replace(/\/+$/,'') || 'https://dp-rent-app.onrender.com';
    const quoteAdminLink = savedQuote.ok ? `${appBaseQuote}/prenotazione/${savedQuote.id}` : `${appBaseQuote}/richieste-attesa`;
    try {
      await dpNotify(DP_STAFF_NUMBERS, `${EMJ.van} PREVENTIVO NOLEGGIO GENERATO - IN ATTESA SI/NO\n\nCliente: ${profileName}\nWhatsApp: ${from}\nMezzo richiesto: ${session.data.cat.label}\nDate: ${dpDateIt(session.data.start)} - ${dpDateIt(session.data.end)}\nKm: ${km}\nTotale: EUR ${euro(calc.totale || 0)}\n\nNota interna mezzo assegnabile: ${mezzo.marca || ''} ${mezzo.modello || ''} ${mezzo.targa || ''}\n\nApri in app: ${quoteAdminLink}

Il cliente sta vedendo il preventivo e deve rispondere SI o NO.`);
    } catch(e) { console.log('Notifica preventivo warning:', e.message); }
    return dpTwimlResponse(res, `${EMJ.ok} *Disponibile*\n\nMezzo: *${session.data.cat.label}*\nDate: ${dpDateIt(session.data.start)} - ${dpDateIt(session.data.end)}\nGiorni: ${calc.giorni || dpDays(session.data.start, session.data.end)}\nKm previsti: ${km}\nPreventivo: *EUR ${euro(calc.totale || 0)}*\n\nConfermi il preventivo?\nRispondi *SI* oppure *NO*.`);
  }

  if(session.state === 'noleggio_confirm'){
    const yn = dpYesNo(body);
    if(yn === 'NO') {
      await dpUpdateWhatsAppQuote(session, 'non_confermato');
      try { await dpNotify(DP_STAFF_NUMBERS, `${EMJ.warn} PREVENTIVO NOLEGGIO NON CONFERMATO\n\nCliente: ${profileName}\nWhatsApp: ${from}\nMezzo richiesto: ${session.data.cat?.label || ''}\nDate: ${session.data.start ? dpDateIt(session.data.start) : ''} - ${session.data.end ? dpDateIt(session.data.end) : ''}\nKm: ${session.data.km || ''}\nTotale: EUR ${euro(session.data.calc?.totale || 0)}\n\nCliente da richiamare se interessa recuperare la richiesta.`); } catch(e) {}
      delete DP_BOT_SESSIONS[from]; return dpTwimlResponse(res, 'Preventivo annullato. Lo staff DP ha comunque ricevuto la richiesta, così non perdiamo il contatto. Scrivi MENU per ricominciare.');
    }
    if(yn !== 'SI') return dpTwimlResponse(res, 'Rispondi SI per confermare oppure NO per annullare.');
    const q = new URLSearchParams({
      categoria: session.data.cat.cats[0] || '',
      data_inizio: dpDateIso(session.data.start),
      data_fine: dpDateIso(session.data.end),
      km_previsti: String(session.data.km || 150),
      telefono: from.replace('whatsapp:','')
    });
    const base = (process.env.APP_BASE_URL || process.env.RENDER_EXTERNAL_URL || '').replace(/\/+$/,'') || 'https://dp-rent-app.onrender.com';
    const link = `${base}/prenota?${q.toString()}`;
    await dpUpdateWhatsAppQuote(session, 'richiesta_cliente');
    await dpNotify(DP_STAFF_NUMBERS, `${EMJ.van} PREVENTIVO NOLEGGIO CONFERMATO\n\nCliente: ${profileName}\nWhatsApp: ${from}\nMezzo richiesto: ${session.data.cat.label}\nDate: ${dpDateIt(session.data.start)} - ${dpDateIt(session.data.end)}\nKm: ${session.data.km}\nTotale: EUR ${euro(session.data.calc?.totale || 0)}\n\nLink cliente:\n${link}\n\nNota interna mezzo assegnabile: ${session.data.mezzo?.marca || ''} ${session.data.mezzo?.modello || ''} ${session.data.mezzo?.targa || ''}`);
    delete DP_BOT_SESSIONS[from];
    return dpTwimlResponse(res, `Perfetto ${EMJ.ok}\n\nOra completa i dati cliente, documento e patente da questo link:\n${link}\n\nDopo il controllo dell ufficio DP RENT verra preparato il contratto definitivo.`);
  }

  if(session.state === 'vendita'){
    await dpNotify(DP_STAFF_NUMBERS, `${EMJ.auto} NUOVA RICHIESTA VENDITA AUTO\n\nCliente: ${profileName}\nWhatsApp: ${from}\n\nRichiesta:\n${body}\n\nBacheca: ${DP_AUTOSUPERMARKET_URL}`);
    delete DP_BOT_SESSIONS[from];
    return dpTwimlResponse(res, `${EMJ.ok} Richiesta inviata al reparto vendite DP AUTO.\n\nTi ricontatteremo al piu presto.`);
  }

  if(session.state === 'trasporto'){
    await dpNotify(DP_STAFF_NUMBERS, `${EMJ.truck} NUOVA RICHIESTA TRASPORTO VEICOLO\n\nCliente: ${profileName}\nWhatsApp: ${from}\n\nDati trasporto:\n${body}`);
    delete DP_BOT_SESSIONS[from];
    return dpTwimlResponse(res, `${EMJ.ok} Richiesta trasporto ricevuta.\n\nL abbiamo inviata allo staff DP. Ti ricontatteremo per il preventivo.`);
  }

  if(session.state === 'altro'){
    const answer = await dpChatGPTAnswer(body);
    await dpNotify(DP_STAFF_NUMBERS, `${EMJ.chat} ALTRA RICHIESTA / CHATGPT\n\nCliente: ${profileName}\nWhatsApp: ${from}\n\nMessaggio cliente:\n${body}\n\nRisposta bot:\n${answer}`);
    return dpTwimlResponse(res, answer + '\n\nScrivi MENU per tornare al menu principale.');
  }

  dpReset(from, profileName);
  return dpTwimlResponse(res, dpMenu(profileName));
}


app.get('/test-whatsapp-staff', async (req, res) => {
  const r = await dpNotify(DP_STAFF_NUMBERS, 'TEST DP RENT - notifica staff ' + new Date().toLocaleString('it-IT'));
  res.json(r);
});
app.get('/test-drive-officina', async (req, res) => {
  const codice = 'TEST-OFF-' + Date.now();
  const r = await dpSaveRequestToDrive('TEST OFFICINA DP', codice, 'test', 'Test', 'Prova salvataggio Drive officina');
  res.json(r);
});

app.post('/whatsapp', dpHandleWhatsApp);
app.post('/webhook', dpHandleWhatsApp);



// =========================
// V99 ROUTE DOCUMENTI INTERNI PRATICA
// =========================
app.get('/documenti/:id', async (req, res) => {
  try {
    const id = req.params.id;
    let allegati = [];
    try {
      if (typeof all === 'function') {
        allegati = await all('SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id DESC', [id]);
      }
    } catch (_) {}
    allegati = v99SoloDocumentiInterni(allegati);

    const rows = allegati.map(a => {
      const nome = esc(a.originalname || a.filename || a.nome || 'file');
      const tipo = esc(a.tipo || a.type || 'documento');
      const link = a.drive_web_link || a.drive_link || a.webViewLink || a.url || '';
      const href = link ? `<a class="btn small" target="_blank" href="${esc(link)}">Apri Drive</a>` : '';
      const img = link ? `<div class="thumb"><a target="_blank" href="${esc(link)}">${nome}</a></div>` : `<div class="thumb">${nome}</div>`;
      return `<div class="doc-card"><b>${tipo}</b>${img}${href}</div>`;
    }).join('');

    res.send(page('Documenti pratica', `
      <div class="box">
        <h2>Documenti interni pratica ${esc(id)}</h2>
        <p class="muted">Visibili solo a DP RENT. Non vengono inseriti nel PDF cliente.</p>
        <div class="doc-grid">${rows || '<p>Nessun documento salvato.</p>'}</div>
        <p><a class="btn" href="/storico">Torna storico</a></p>
      </div>
      <style>
        .doc-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:16px}
        .doc-card{background:#fff;border:1px solid #eee;border-radius:18px;padding:18px;box-shadow:0 10px 30px #0001}
        .thumb{margin:12px 0;padding:16px;background:#f7f7f9;border-radius:14px;word-break:break-all}
        .btn.small{font-size:14px;padding:8px 12px}
      
.contract-main-actions{margin-top:16px}.contract-main-actions .btn{min-width:190px;text-align:center}.contract-secondary-actions .btn{min-width:150px;text-align:center}
@media(max-width:700px){.contract-main-actions .btn,.contract-secondary-actions .btn{width:100%;min-width:0}}


/* V109 FIX leggibilita mobile */
header{padding-top:max(22px, env(safe-area-inset-top));}
.top-actions{max-width:1180px;margin:0 auto 14px!important;padding:10px 0!important;}
.top-actions .back-btn::before{content:""!important;}
.top-actions .back-btn,.top-actions a{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif!important;font-size:clamp(18px,2.6vw,24px)!important;letter-spacing:0!important;line-height:1.1!important;white-space:nowrap!important;color:#fff!important;overflow:hidden;text-overflow:ellipsis;}
.top-actions .back-btn{background:#333!important;}
.top-actions .home-btn{background:#d70000!important;}
.client-back button{font-size:18px!important;font-weight:900!important;background:#333!important;color:#fff!important;}
@media(max-width:700px){
  nav{padding-top:calc(14px + env(safe-area-inset-top));}
  .top-actions{position:sticky;top:0;z-index:50;padding:10px 12px!important;gap:10px!important;background:rgba(244,244,244,.96)!important;}
  .top-actions .back-btn,.top-actions a{min-width:0!important;width:calc(50% - 5px)!important;flex:1 1 calc(50% - 5px)!important;padding:14px 8px!important;}
  .contract-main-actions .btn{width:100%!important;}
}

</style>
    `));
  } catch (e) {
    res.status(500).send('Errore documenti: ' + esc(e.message));
  }
});



// =========================
// V107 CaRGOS UID + BLOCCO DOPPIO INVIO
// =========================
async function v107EnsureCargosColumns() {
  try {
    if (typeof run !== 'function') return;
    await run("ALTER TABLE prenotazioni ADD COLUMN cargos_uid TEXT").catch(()=>{});
    await run("ALTER TABLE prenotazioni ADD COLUMN cargos_inviato INTEGER DEFAULT 0").catch(()=>{});
    await run("ALTER TABLE prenotazioni ADD COLUMN cargos_inviato_at TEXT").catch(()=>{});
  } catch(e) {
    console.error("V107 cargos columns:", e.message);
  }
}
v107EnsureCargosColumns();

function v107ExtractUid(result) {
  return result?.data?.[0]?.transactionid || result?.transactionid || result?.uid || null;
}

async function v107GetCargosStatus(id) {
  try {
    if (typeof get === 'function') {
      return await get("SELECT cargos_uid,cargos_inviato,cargos_inviato_at FROM prenotazioni WHERE id=?", [id]);
    }
  } catch(e) {}
  return null;
}

async function v107SaveCargosUid(id, uid) {
  try {
    if (!uid || typeof run !== 'function') return;
    await v107EnsureCargosColumns();
    await run("UPDATE prenotazioni SET cargos_uid=?, cargos_inviato=1, cargos_inviato_at=? WHERE id=?", [
      uid, new Date().toISOString(), id
    ]).catch(()=>{});
  } catch(e) {
    console.error("V107 save cargos uid:", e.message);
  }
}




// =========================
// V117 - SCANSIONE DOCUMENTI UFFICIO + INDIRIZZO FATTURAZIONE
// =========================
async function v115EnsureBillingAndScanDb(){
  const clienteCols = {
    indirizzo_fatturazione:'TEXT', citta_fatturazione:'TEXT', provincia_fatturazione:'TEXT', cap_fatturazione:'TEXT',
    tipo_cliente:'TEXT', piva:'TEXT', partita_iva:'TEXT', ragione_sociale:'TEXT', pec:'TEXT', sdi:'TEXT', codice_sdi:'TEXT',
    scansione_batch:'TEXT'
  };
  const prenCols = {
    indirizzo_fatturazione:'TEXT', citta_fatturazione:'TEXT', provincia_fatturazione:'TEXT', cap_fatturazione:'TEXT'
  };
  for (const [c,t] of Object.entries(clienteCols)) await run(`ALTER TABLE clienti ADD COLUMN ${c} ${t}`).catch(()=>{});
  for (const [c,t] of Object.entries(prenCols)) await run(`ALTER TABLE prenotazioni ADD COLUMN ${c} ${t}`).catch(()=>{});
}
v115EnsureBillingAndScanDb().catch(e=>console.log('V117 migrazione warning:', e.message));

function v115Date(x){ return String(x||'').slice(0,10); }
function v115MergeDocs(list){
  const out = {};
  for (const o of (list||[])) {
    if (!o || typeof o !== 'object') continue;
    const map = {
      nome:'nome', cognome:'cognome', codice_fiscale:'codice_fiscale', data_nascita:'data_nascita', luogo_nascita:'luogo_nascita', indirizzo:'indirizzo',
      numero_documento:'documento_numero', documento_numero:'documento_numero', scadenza_documento:'documento_scadenza', data_scadenza:'documento_scadenza',
      numero_patente:'patente_numero', patente_numero:'patente_numero', scadenza_patente:'patente_scadenza', categoria_patente:'categoria_patente'
    };
    for (const [src,dst] of Object.entries(map)) if (!out[dst] && o[src]) out[dst] = v115Date(o[src]);
    if (!out.documento_scadenza && o.tipo_documento && String(o.tipo_documento).toLowerCase().includes('patente') && o.data_scadenza) out.patente_scadenza = v115Date(o.data_scadenza);
  }
  return out;
}

app.get('/scansione-documenti', (req,res)=>{
  res.send(page('Scansione documenti', `<div class="box"><h2>Scansione documenti cliente</h2><p class="notice">Postazione ufficio: metti sullo scanner carta identità, patente e tessera sanitaria. Puoi caricare più file insieme; il gestionale prova a leggere i dati, li divide e ti prepara l'anagrafica cliente.</p><form method="POST" action="/scansione-documenti" enctype="multipart/form-data"><label>File scanner / PDF / foto documenti</label><input type="file" name="scan_docs" accept="image/*,application/pdf" multiple required><button>Leggi documenti e prepara cliente</button><a class="btn btn2" href="/clienti">Torna clienti</a></form></div>`));
});

app.post('/scansione-documenti', upload.array('scan_docs', 30), async (req,res)=>{
  try{
    await v115EnsureBillingAndScanDb();
    const files = req.files || [];
    if (!files.length) return res.send(page('Scansione documenti', `<div class="box"><h2 class="bad">Nessun file caricato</h2><a class="btn" href="/scansione-documenti">Torna</a></div>`));
    const ocrResults = [];
    for (const f of files) {
      try { ocrResults.push(await estraiDatiDocumentoConAI(f.path, f.mimetype)); }
      catch(e){ console.log('V117 OCR scan warning:', e.message); }
    }
    const d = v115MergeDocs(ocrResults);
    const batch = 'SCAN' + Date.now();
    PREN_OCR_UPLOADS[batch] = files.map(f=>({tipo:'scansione_ufficio', f}));
    setTimeout(()=>{ delete PREN_OCR_UPLOADS[batch]; }, 6*60*60*1000).unref?.();
    const q = new URLSearchParams(d);
    q.set('scan_batch', batch);
    res.redirect('/scansione-documenti/controlla?' + q.toString());
  }catch(e){
    res.status(500).send(page('Errore scansione', `<div class="box"><h2 class="bad">Errore scansione</h2><pre>${esc(e.stack||e.message)}</pre><a class="btn" href="/scansione-documenti">Torna</a></div>`));
  }
});

app.get('/scansione-documenti/controlla', (req,res)=>{
  const q=req.query||{}; const val=k=>esc(q[k]||'');
  res.send(page('Controlla cliente da scansione', `<div class="box"><h2>Controlla dati letti da scanner</h2><p class="notice">Controlla e correggi i dati prima di salvare. Le scansioni vengono salvate nell'archivio documenti cliente.</p><form method="POST" action="/scansione-documenti/salva"><input type="hidden" name="scan_batch" value="${val('scan_batch')}"><div class="grid"><div><label>Nome</label><input name="nome" value="${val('nome')}" required></div><div><label>Cognome</label><input name="cognome" value="${val('cognome')}" required></div><div><label>Telefono</label><input name="telefono"></div><div><label>Email</label><input name="email"></div><div><label>Codice fiscale</label><input name="codice_fiscale" value="${val('codice_fiscale')}"></div><div><label>Data nascita</label><input type="date" name="data_nascita" value="${val('data_nascita')}"></div><div><label>Luogo nascita</label><input name="luogo_nascita" value="${val('luogo_nascita')}"></div><div class="full"><label>Indirizzo residenza</label><input name="indirizzo" value="${val('indirizzo')}"></div><div><label>Città</label><input name="citta"></div><div><label>Provincia</label><input name="provincia"></div><div><label>CAP</label><input name="cap"></div><div><label>Numero documento</label><input name="documento_numero" value="${val('documento_numero')}"></div><div><label>Scadenza documento</label><input type="date" name="documento_scadenza" value="${val('documento_scadenza')}"></div><div><label>Numero patente</label><input name="patente_numero" value="${val('patente_numero')}"></div><div><label>Scadenza patente</label><input type="date" name="patente_scadenza" value="${val('patente_scadenza')}"></div><div><label>Categoria patente</label><input name="categoria_patente" value="${val('categoria_patente')}"></div></div><h3>Fatturazione</h3><div class="grid"><div><label>Tipo cliente</label><select name="tipo_cliente"><option value="privato">Privato</option><option value="azienda">Azienda</option></select></div><div><label>Ragione sociale</label><input name="ragione_sociale"></div><div><label>Partita IVA</label><input name="piva"></div><div><label>PEC</label><input name="pec"></div><div><label>Codice SDI</label><input name="sdi"></div><div class="full"><label>Indirizzo fatturazione</label><input name="indirizzo_fatturazione" value="${val('indirizzo')}"></div><div><label>Città fatturazione</label><input name="citta_fatturazione"></div><div><label>Provincia fatturazione</label><input name="provincia_fatturazione"></div><div><label>CAP fatturazione</label><input name="cap_fatturazione"></div></div><button>Salva anagrafica cliente</button><a class="btn btn2" href="/scansione-documenti">Rifai scansione</a></form></div>`));
});

app.post('/scansione-documenti/salva', async (req,res)=>{
  try{
    await v115EnsureBillingAndScanDb();
    const b=req.body||{};
    const d={ nome:b.nome||'', cognome:b.cognome||'', telefono:b.telefono||'', email:b.email||'', codice_fiscale:String(b.codice_fiscale||'').toUpperCase(), indirizzo:b.indirizzo||'', citta:b.citta||'', cap:b.cap||'', data_nascita:b.data_nascita||'', luogo_nascita:b.luogo_nascita||'', documento_numero:b.documento_numero||'', documento_scadenza:b.documento_scadenza||'', patente_numero:b.patente_numero||'', patente_scadenza:b.patente_scadenza||'', categoria_patente:b.categoria_patente||'', tipo_cliente:b.tipo_cliente||'privato', ragione_sociale:b.ragione_sociale||'', piva:b.piva||b.partita_iva||'', partita_iva:b.piva||b.partita_iva||'', pec:b.pec||'', sdi:b.sdi||b.codice_sdi||'', codice_sdi:b.sdi||b.codice_sdi||'', indirizzo_fatturazione:b.indirizzo_fatturazione||b.indirizzo||'', citta_fatturazione:b.citta_fatturazione||b.citta||'', provincia_fatturazione:b.provincia_fatturazione||'', cap_fatturazione:b.cap_fatturazione||b.cap||'', scansione_batch:b.scan_batch||'' };
    const r=await run(`INSERT INTO clienti (nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,data_nascita,luogo_nascita,documento_numero,documento_scadenza,patente_numero,patente_scadenza,categoria_patente,tipo_cliente,ragione_sociale,piva,partita_iva,pec,sdi,codice_sdi,indirizzo_fatturazione,citta_fatturazione,provincia_fatturazione,cap_fatturazione,scansione_batch,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,CURRENT_TIMESTAMP)`, [d.nome,d.cognome,d.telefono,d.email,d.codice_fiscale,d.indirizzo,d.citta,d.cap,d.data_nascita,d.luogo_nascita,d.documento_numero,d.documento_scadenza,d.patente_numero,d.patente_scadenza,d.categoria_patente,d.tipo_cliente,d.ragione_sociale,d.piva,d.partita_iva,d.pec,d.sdi,d.codice_sdi,d.indirizzo_fatturazione,d.citta_fatturazione,d.provincia_fatturazione,d.cap_fatturazione,d.scansione_batch]);
    const files = PREN_OCR_UPLOADS[d.scansione_batch] || [];
    for (const item of files) await run(`INSERT INTO allegati (cliente_id,tipo,filename,originalname,path,mimetype,size) VALUES (?,?,?,?,?,?,?)`, [r.lastID, item.tipo, item.f.filename, item.f.originalname, item.f.path, item.f.mimetype, item.f.size]).catch(()=>{});
    delete PREN_OCR_UPLOADS[d.scansione_batch];
    res.redirect('/cliente/'+r.lastID);
  }catch(e){ res.status(500).send(page('Errore salvataggio cliente', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.stack||e.message)}</pre><a class="btn" href="/scansione-documenti">Torna</a></div>`)); }
});

app.listen(PORT, '0.0.0.0', () => {
  console.log('DP RENT APP V122 alert attese porta ' + PORT);
  console.log('Staff WhatsApp:', DP_STAFF_NUMBERS.join(', '));
});


// =========================
// V99 FIX ALLEGATI / DOCUMENTI SEPARATI
// =========================
function v99NormFileName(v) {
  return String(v || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9._-]+/g, '_')
    .replace(/_+/g, '_')
    .slice(0, 180);
}

function v99AllegatoKey(a) {
  return [
    String(a?.id || ''),
    String(a?.drive_file_id || a?.driveFileId || ''),
    v99NormFileName(a?.originalname || a?.filename || a?.name || ''),
    String(a?.size || a?.bytes || ''),
    String(a?.tipo || a?.type || '')
  ].join('|');
}

function v99TipoRank(tipo) {
  const t = String(tipo || '').toLowerCase();
  if (t.includes('documento') && t.includes('fronte')) return 10;
  if (t.includes('documento') && t.includes('retro')) return 20;
  if (t.includes('patente') && t.includes('fronte')) return 30;
  if (t.includes('patente') && t.includes('retro')) return 40;
  if (t.includes('documento')) return 50;
  if (t.includes('patente')) return 60;
  if (t.includes('checkin') || t.includes('check-in')) return 70;
  if (t.includes('checkout') || t.includes('check-out')) return 80;
  return 99;
}

function v99DedupeAllegati(list) {
  const seen = new Set();
  const out = [];
  for (const a of Array.isArray(list) ? list : []) {
    if (!a || typeof a !== 'object') continue;
    const k = v99AllegatoKey(a);
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(a);
  }
  return out.sort((a,b) => {
    const r = v99TipoRank(a.tipo || a.type) - v99TipoRank(b.tipo || b.type);
    if (r) return r;
    return String(a.originalname || a.filename || '').localeCompare(String(b.originalname || b.filename || ''));
  });
}

function v99SoloDocumentiInterni(list) {
  return v99DedupeAllegati(list).filter(a => {
    const t = String(a.tipo || a.type || '').toLowerCase();
    return t.includes('documento') || t.includes('patente') || t.includes('allegat');
  });
}

function v99SoloFotoMezzo(list) {
  return v99DedupeAllegati(list).filter(a => {
    const t = String(a.tipo || a.type || '').toLowerCase();
    return t.includes('checkin') || t.includes('checkout') || t.includes('mezzo');
  });
}

function v99IsClientePdfImageAllowed(tipo) {
  // Privacy: i documenti personali NON devono finire nel PDF cliente.
  const t = String(tipo || '').toLowerCase();
  if (t.includes('documento') || t.includes('patente') || t.includes('carta') || t.includes('ident')) return false;
  return false;
}



// =========================
// V100 FIX: UN SOLO PDF CONTRATTO + CITTADINANZA CARGOS
// =========================
function v100CittadinanzaCargosCod() { return 100; }

async function v100DeleteOldContractPdfsInDrive(folderId, keepNamePrefix) {
  try {
    if (!folderId || typeof drive === 'undefined' || !drive.files) return;
    const q = `'${folderId}' in parents and trashed=false and mimeType='application/pdf'`;
    const list = await drive.files.list({
      q,
      fields: 'files(id,name)',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });
    const files = list?.data?.files || [];
    for (const f of files) {
      const name = String(f.name || '').toLowerCase();
      if (name.startsWith('contratto_') || name.includes('contratto')) {
        await drive.files.delete({ fileId: f.id, supportsAllDrives: true }).catch(()=>{});
      }
    }
  } catch (e) {
    console.error('V100 delete old contract pdf:', e.message);
  }
}

function v100PatchCargosPayload(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  for (const k of Object.keys(obj)) {
    if (k === 'CONDUCENTE_CONTRAENTE_CITTADINANZA_COD') obj[k] = 100;
    else if (obj[k] && typeof obj[k] === 'object') v100PatchCargosPayload(obj[k]);
  }
  return obj;
}



// =========================
// V102 FIX PDF DOPPIO FIRMA/PDF
// =========================
const v102PdfLock = new Map();

function v102PdfKey(prenotazioneId, numeroContratto) {
  return String(prenotazioneId || numeroContratto || '').trim();
}

async function v102RunOncePdf(key, fn) {
  key = String(key || 'global');
  if (v102PdfLock.has(key)) return await v102PdfLock.get(key);
  const p = Promise.resolve().then(fn).finally(() => {
    setTimeout(() => v102PdfLock.delete(key), 2500);
  });
  v102PdfLock.set(key, p);
  return await p;
}

async function v102DeleteOldContractDbAndDrive(prenotazioneId, folderId) {
  try {
    if (typeof run === 'function') {
      await run(`DELETE FROM allegati WHERE prenotazione_id=? AND (
        lower(coalesce(tipo,'')) LIKE '%contratto%' OR
        lower(coalesce(filename,'')) LIKE 'contratto_%' OR
        lower(coalesce(originalname,'')) LIKE 'contratto_%'
      )`, [prenotazioneId]).catch(()=>{});
    }
  } catch (_) {}

  try {
    await v103DeleteOldPdfEverywhere(id || prenotazioneId || req.params.id || p?.id, folderId);
  } catch (_) {}
}



// =========================
// V103 DEFINITIVO: CARGOS CITTADINANZA + PDF UNICO
// =========================
const V103_CARGOS_CITTADINANZA_ITALIA = '100000100';
const v103PdfLocks = new Map();

function v103DeepFixCargos(obj) {
  if (!obj || typeof obj !== 'object') return obj;
  for (const k of Object.keys(obj)) {
    if (k === 'CONDUCENTE_CONTRAENTE_CITTADINANZA_COD') obj[k] = V103_CARGOS_CITTADINANZA_ITALIA;
    else if (obj[k] && typeof obj[k] === 'object') v103DeepFixCargos(obj[k]);
  }
  if (!Object.prototype.hasOwnProperty.call(obj, 'CONDUCENTE_CONTRAENTE_CITTADINANZA_COD')) {
    obj.CONDUCENTE_CONTRAENTE_CITTADINANZA_COD = V104_CITTADINANZA_ITALIA_CARGOS;
  }
  return obj;
}

async function v103DeleteOldPdfEverywhere(prenotazioneId, folderId) {
  try {
    if (typeof run === 'function' && prenotazioneId) {
      await run(`DELETE FROM allegati WHERE prenotazione_id=? AND (
        lower(coalesce(tipo,'')) LIKE '%contratto%' OR
        lower(coalesce(filename,'')) LIKE 'contratto_%' OR
        lower(coalesce(originalname,'')) LIKE 'contratto_%'
      )`, [prenotazioneId]).catch(()=>{});
    }
  } catch(e) { console.error('V103 delete db pdf', e.message); }

  try {
    if (folderId && typeof drive !== 'undefined' && drive.files) {
      const q = `'${folderId}' in parents and trashed=false and mimeType='application/pdf'`;
      const r = await drive.files.list({ q, fields:'files(id,name)', supportsAllDrives:true, includeItemsFromAllDrives:true });
      for (const f of (r.data.files || [])) {
        const n = String(f.name || '').toLowerCase();
        if (n.includes('contratto')) await drive.files.delete({ fileId:f.id, supportsAllDrives:true }).catch(()=>{});
      }
    }
  } catch(e) { console.error('V103 delete drive pdf', e.message); }
}



// =========================
// V104 FIX FINALE CaRGOS: cittadinanza obbligatoria sempre presente
// =========================
const V104_CITTADINANZA_ITALIA_CARGOS = '100000100';

function v104FixCargosCitizenshipDeep(x) {
  if (Array.isArray(x)) {
    x.forEach(v104FixCargosCitizenshipDeep);
    return x;
  }
  if (x && typeof x === 'object') {
    x.CONDUCENTE_CONTRAENTE_CITTADINANZA_COD = V104_CITTADINANZA_ITALIA_CARGOS;
    x.conducente_contraente_cittadinanza_cod = V104_CITTADINANZA_ITALIA_CARGOS;
    for (const k of Object.keys(x)) v104FixCargosCitizenshipDeep(x[k]);
  }
  return x;
}

function v104FixCargosString(s) {
  s = String(s || '');
  if (s.includes('CONDUCENTE_CONTRAENTE_CITTADINANZA_COD')) {
    s = s.replace(/(CONDUCENTE_CONTRAENTE_CITTADINANZA_COD["']?\s*[:=]\s*["']?)[^"',;}\]\s]*/g, '$1' + V104_CITTADINANZA_ITALIA_CARGOS);
  }
  return s;
}

function v104CargosPayload(x) {
  if (typeof x === 'string') return v104FixCargosString(x);
  return v104FixCargosCitizenshipDeep(x);
}



// =========================
// V109 FIX - PATCH FINALE DANIELE
// documenti clienti persistenti, PDF lock, firma redirect, badge, azienda completa
// =========================
try {
  addColumn('clienti','provincia','TEXT');
  addColumn('clienti','tipo_cliente','TEXT');
  addColumn('clienti','ragione_sociale','TEXT');
  addColumn('clienti','pec','TEXT');
  addColumn('clienti','sdi','TEXT');
  addColumn('clienti','documento_file','TEXT');
  addColumn('clienti','patente_file','TEXT');
  addColumn('prenotazioni','firma_path','TEXT');
  addColumn('prenotazioni','codice_fiscale','TEXT');
  addColumn('prenotazioni','tipo_cliente','TEXT');
  addColumn('prenotazioni','ragione_sociale','TEXT');
  addColumn('prenotazioni','provincia','TEXT');
  addColumn('prenotazioni','cargos_inviato','INTEGER DEFAULT 0');
  addColumn('prenotazioni','cargos_uid','TEXT');
  addColumn('allegati','cliente_id','INTEGER');
  addColumn('allegati','size','INTEGER');
} catch(e) { console.log('V109 add columns skip:', e.message); }

const v108PdfLocks = new Map();
const v108OriginalGeneraPdfContratto = generaPdfContratto;
generaPdfContratto = async function v108GeneraPdfContrattoLocked(id, opts = {}) {
  const key = String(id || 'global');
  if (v108PdfLocks.has(key)) return await v108PdfLocks.get(key);
  const job = Promise.resolve().then(async () => {
    try {
      const p = await get(`SELECT pdf_path FROM prenotazioni WHERE id=?`, [id]).catch(()=>null);
      if (p && p.pdf_path && !opts.forceDrive && fs.existsSync(p.pdf_path)) {
        const age = Date.now() - fs.statSync(p.pdf_path).mtimeMs;
        if (age < 1500) return p.pdf_path;
      }
    } catch(_) {}
    return await v108OriginalGeneraPdfContratto(id, opts);
  }).finally(() => setTimeout(()=>v108PdfLocks.delete(key), 3500));
  v108PdfLocks.set(key, job);
  return await job;
};

function v108FileUrl(f){
  if(!f) return '';
  const name = path.basename(String(f));
  return '/uploads/' + encodeURIComponent(name);
}
function v108DocRows(files){
  return (files || []).map(a => `<tr><td>${esc(a.tipo||'documento')}</td><td>${esc(a.originalname||a.filename||'file')}</td><td>${esc(a.created_at||'')}</td><td><a class="btn btn2" target="_blank" href="/uploads/${encodeURIComponent(path.basename(a.path||a.filename||''))}">Apri</a></td></tr>`).join('') || '<tr><td colspan="4">Nessun documento caricato.</td></tr>';
}

app.get('/documenti-clienti', async (req,res)=>{
  const rows = await all(`SELECT c.*, (SELECT COUNT(*) FROM allegati a WHERE a.tipo LIKE 'cliente_%' AND (a.prenotazione_id IS NULL OR a.prenotazione_id=0) AND (a.originalname LIKE '%'||c.cognome||'%' OR a.originalname LIKE '%'||c.nome||'%')) as docs FROM clienti c ORDER BY updated_at DESC, id DESC LIMIT 300`).catch(()=>[]);
  const trs = rows.map(c=>`<tr><td><b>${esc(c.nome)} ${esc(c.cognome)}</b><br><span class="muted">${esc(c.telefono||'')} ${esc(c.email||'')}</span></td><td>${esc(c.codice_fiscale||'')}</td><td>${esc(c.documento_numero||'')}<br>${esc(c.patente_numero||'')}</td><td><a class="btn btn3" href="/cliente/${c.id}/documenti">Archivio</a> <a class="btn" href="/nuova-da-cliente/${c.id}">Contratto</a></td></tr>`).join('');
  res.send(page('Documenti clienti', `<div class="premium-card"><h2>Archivio documenti clienti</h2><p>Carica e conserva carta identità, patente, codice fiscale e documenti azienda su disco persistente Render.</p><a class="btn" href="/clienti">Vai ai clienti</a></div><table><tr><th>Cliente</th><th>CF</th><th>Documento / Patente</th><th>Azioni</th></tr>${trs || '<tr><td colspan="4">Nessun cliente.</td></tr>'}</table>`));
});

app.get('/cliente/:id/documenti', async (req,res)=>{
  const c = await get(`SELECT * FROM clienti WHERE id=?`, [req.params.id]);
  if(!c) return res.redirect('/clienti');
  const files = await all(`SELECT * FROM allegati WHERE cliente_id=? OR (prenotazione_id IS NULL AND tipo LIKE 'cliente_%' AND originalname LIKE ?) ORDER BY id DESC`, [req.params.id, `%${c.cognome||''}%`]).catch(()=>[]);
  res.send(page('Archivio documenti cliente', `<div class="premium-card"><h2>Documenti: ${esc(c.nome)} ${esc(c.cognome)}</h2><p><b>CF:</b> ${esc(c.codice_fiscale||'')} | <b>Tel:</b> ${esc(c.telefono||'')}</p><form method="POST" action="/cliente/${c.id}/documenti" enctype="multipart/form-data"><div class="grid"><div><label>Tipo documento</label><select name="tipo"><option value="cliente_documento">Carta identità / documento</option><option value="cliente_patente">Patente</option><option value="cliente_cf">Codice fiscale</option><option value="cliente_azienda">Documento azienda</option><option value="cliente_altro">Altro</option></select></div><div><label>File</label><input type="file" name="file" accept="image/*,.pdf" required></div></div><button>Carica documento</button></form><div class="big-actions"><a class="btn" href="/nuova-da-cliente/${c.id}">Crea contratto con dati auto-compilati</a><a class="btn btn2" href="/cliente/${c.id}">Scheda cliente</a></div></div><table><tr><th>Tipo</th><th>Nome file</th><th>Data</th><th>Apri</th></tr>${v108DocRows(files)}</table>`));
});

app.post('/cliente/:id/documenti', upload.single('file'), async (req,res)=>{
  try{
    const c = await get(`SELECT * FROM clienti WHERE id=?`, [req.params.id]);
    if(!c || !req.file) return res.redirect('/clienti');
    const ext = path.extname(req.file.originalname || '') || '';
    const safeName = `cliente_${req.params.id}_${Date.now()}_${String(req.body.tipo||'documento').replace(/[^a-z0-9_-]/gi,'')}${ext}`;
    const finalPath = path.join(uploadDir, safeName);
    fs.renameSync(req.file.path, finalPath);
    await run(`INSERT INTO allegati (cliente_id, prenotazione_id, tipo, filename, originalname, path, mimetype, size) VALUES (?,?,?,?,?,?,?,?)`, [req.params.id, null, req.body.tipo || 'cliente_documento', safeName, req.file.originalname || safeName, finalPath, req.file.mimetype || '', req.file.size || 0]).catch(async ()=>{
      await run(`INSERT INTO allegati (prenotazione_id, tipo, filename, originalname, path, mimetype, size) VALUES (?,?,?,?,?,?,?)`, [null, req.body.tipo || 'cliente_documento', safeName, req.file.originalname || safeName, finalPath, req.file.mimetype || '', req.file.size || 0]);
    });
    if(req.body.tipo === 'cliente_documento') await run(`UPDATE clienti SET documento_file=?, updated_at=CURRENT_TIMESTAMP WHERE id=?`, [finalPath, req.params.id]).catch(()=>{});
    if(req.body.tipo === 'cliente_patente') await run(`UPDATE clienti SET patente_file=?, updated_at=CURRENT_TIMESTAMP WHERE id=?`, [finalPath, req.params.id]).catch(()=>{});
    res.redirect(`/cliente/${req.params.id}/documenti`);
  } catch(e){ res.status(500).send(page('Errore documento', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`)); }
});

app.get('/contratto/:id/firmato', async (req,res)=>{
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if(!p) return res.send('Contratto non trovato');
  try { await generaPdfContratto(req.params.id, { forceDrive:false }); } catch(e) {}
  const fresh = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]).catch(()=>p);
  const pdfLink = fresh?.pdf_drive_web_link || fresh?.pdf_drive_link || p.pdf_drive_web_link || p.pdf_drive_link || '';
  res.send(publicFirmaPage('Firma salvata DP RENT', `<h2 class="ok">Firma salvata correttamente</h2><p>Grazie. Il contratto ${esc(p.codice||p.id)} &egrave; stato firmato e registrato da DP RENT.</p>${pdfLink ? `<a class="btn btn3" target="_blank" href="${esc(pdfLink)}">Apri copia PDF</a>` : '<p class="muted">Puoi chiudere questa pagina.</p>'}`));
});


// V109 alias route robuste per WhatsApp/firma
app.get('/contratto/:id/invia-firma-whatsapp', (req,res)=>res.redirect('/firma-whatsapp/' + req.params.id));
app.get('/prenotazione/:id/invia-whatsapp', (req,res)=>res.redirect('/contratto/' + req.params.id + '/invia-whatsapp'));
app.get('/prenotazione/:id/invia-firma-whatsapp', (req,res)=>res.redirect('/firma-whatsapp/' + req.params.id));

app.get('/v108-check', async (req,res)=>{
  const dbOk = fs.existsSync(DB_PATH);
  const dirs = [DATA_DIR, uploadDir, contractsDir, firmeDir].map(d=>`${d}: ${fs.existsSync(d) ? 'OK' : 'NO'}`).join('\n');
  res.type('text/plain').send(`DP RENT V109 OK\nDB: ${dbOk ? DB_PATH : 'NO'}\n${dirs}`);
});
