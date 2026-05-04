require('dotenv').config();
require('dns').setDefaultResultOrder('ipv4first');

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

const app = express();
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

const DATA_DIR = path.join(__dirname, 'data');
const uploadDir = path.join(__dirname, 'uploads');
const contractsDir = path.join(__dirname, 'contracts');
const firmeDir = path.join(__dirname, 'firme');
const publicDir = path.join(__dirname, 'public');

[DATA_DIR, uploadDir, contractsDir, firmeDir, publicDir].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

app.use('/public', express.static(publicDir));
app.use('/uploads', express.static(uploadDir));
app.use('/contracts', express.static(contractsDir));

const upload = multer({ dest: uploadDir });
const db = new sqlite3.Database(path.join(DATA_DIR, 'database.sqlite'));

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
function addColumn(table, column, type) {
  db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${type}`, () => {});
}

db.serialize(() => {
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
      cargos_stato TEXT,
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  [
    ['pec','TEXT'], ['sdi','TEXT'], ['conducente1','TEXT'], ['patente1','TEXT'], ['patente1_scadenza','TEXT'],
    ['conducente2','TEXT'], ['patente2','TEXT'], ['patente2_scadenza','TEXT'],
    ['firma_path','TEXT'], ['pdf_path','TEXT'], ['pdf_drive_file_id','TEXT'], ['pdf_drive_web_link','TEXT'],
    ['nexi_link','TEXT'], ['nexi_stato','TEXT'], ['nexi_raw','TEXT'], ['cargos_stato','TEXT'],
    ['extra_km','REAL DEFAULT 0'], ['extra_fuori_orario','REAL DEFAULT 0'],
    ['carburante_uscita','TEXT DEFAULT "4/4 pieno"'], ['carburante_rientro','TEXT DEFAULT "4/4 pieno"'],
    ['km_uscita','INTEGER'], ['km_rientro','INTEGER'], ['note','TEXT']
  ].forEach(c => addColumn('prenotazioni', c[0], c[1]));

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

function page(title, content) {
  const logoPath = path.join(publicDir, 'logo.png');
  const logoHtml = fs.existsSync(logoPath)
    ? `<img src="/public/logo.png" style="height:48px;max-width:180px;object-fit:contain;background:white;border-radius:8px;padding:4px;">`
    : `<span class="brandText">DP RENT</span>`;

  return `<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
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
.badge-red{background:#d90000;color:white}.badge-green{background:#1fae4b;color:white}.badge-orange{background:#ffb000;color:#111}
pre{white-space:pre-wrap;word-break:break-word;background:#111;color:#fff;padding:12px;border-radius:8px;overflow:auto}
@media(max-width:700px){.grid{grid-template-columns:1fr}main{padding:10px}header h1{font-size:24px}.tile{font-size:19px;min-height:100px}th,td{font-size:12px;padding:6px}}
</style>
</head>
<body>
<header>${logoHtml}<h1>DP RENT APP <small style="font-size:13px;color:#ddd">V24 OCR + CARGOS</small></h1></header>
<nav>
<a href="/">Dashboard</a>
<a href="/mezzi-web">Mezzi</a>
<a href="/scadenze-mezzi">Scadenze</a>
<a href="/import-mezzi">Import Excel</a>
<a href="/nuova-prenotazione">Nuova prenotazione</a>
<a href="/prenotazioni">Storico</a>
<a href="/planning">Planning</a>
<a href="/prenota">Pagina cliente</a>
<a href="/cargos">Ca.R.G.O.S.</a><a href="/cargos-config">Config CARGOS</a>
<a href="/logo">Logo</a>
<a href="/test-email">Test Email</a>
<a href="/test-drive">Test Drive</a>
</nav>
<main>${content}</main>
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
  if (moment(fine).isBefore(moment(inizio))) return 'La data fine non puÃ² essere precedente alla data inizio';
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
  const safe = String(p.codice || p.id).replace(/[^a-zA-Z0-9_-]/g, '');
  const signed = p.firma_path && fs.existsSync(p.firma_path) ? '_firmato' : '';
  return `contratto_${safe}${signed}.pdf`;
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
  const doc = new PDFDocument({ margin: 40, size:'A4' });
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
  row(doc, 'Codice fiscale', p.codice_fiscale || '', 55, yLeft, 220); yLeft += 18;
  row(doc, 'Indirizzo', `${p.indirizzo || ''} ${p.citta || ''} ${p.cap || ''}`, 55, yLeft, 220); yLeft += 18;
  row(doc, 'Fatturazione', `${p.tipo_cliente || ''} ${p.ragione_sociale || ''} ${p.piva || ''}`, 55, yLeft, 220); yLeft += 18;
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
  doc.fontSize(8).fillColor('#111').text('Il cliente dichiara di aver preso visione e accettare le condizioni generali di noleggio e lâinformativa privacy DP RENT / Trasporti DP S.R.L. Il mezzo deve essere riconsegnato nelle stesse condizioni, con carburante equivalente. Danni, multe, pedaggi, franchigie, ritardi, smarrimenti e costi accessori restano a carico del cliente.', 55, y, {width:490});
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
  doc.end();

  await new Promise((resolve, reject) => {
    stream.on('finish', resolve);
    stream.on('error', reject);
  });

  let driveRes = null;
  const forceDrive = !!opts.forceDrive;

  if (shouldUploadPdfToDrive(p, forceDrive)) {
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

async function buildCargosRecordForContract(id) {
  const p = await get(`SELECT p.*, m.targa, m.marca, m.modello, m.categoria FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE p.id=?`, [id]);
  if (!p) throw new Error('Contratto non trovato');

  const fields = [
    cargosPad(p.codice,50), cargosDateTime(p.created_at || moment(), moment().format('HH:mm')),
    cargosPad(process.env.CARGOS_TIPO_PAGAMENTO || '1',1,'number'),
    cargosDateTime(p.data_inizio,p.ora_inizio || '08:30'), cargosPad(process.env.CARGOS_LUOGO_COD || '',9,'number'), cargosPad(AZIENDA.indirizzo,150),
    cargosDateTime(p.data_fine,p.ora_fine || '18:00'), cargosPad(process.env.CARGOS_LUOGO_COD || '',9,'number'), cargosPad(AZIENDA.indirizzo,150),
    cargosPad(process.env.CARGOS_OPERATORE_ID || '',50), cargosPad(process.env.CARGOS_AGENZIA_ID || '',30), cargosPad(process.env.CARGOS_AGENZIA_NOME || AZIENDA.nome,70),
    cargosPad(process.env.CARGOS_LUOGO_COD || '',9,'number'), cargosPad(AZIENDA.indirizzo,150), cargosPad(AZIENDA.telefono,20,'number'),
    cargosPad(cargosTipoVeicolo(p.categoria),1), cargosPad(p.marca || '',50), cargosPad(p.modello || '',100), cargosPad(p.targa || '',15),
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
    headers:{ Authorization:`Bearer ${encrypted}`, Organization:(process.env.CARGOS_USERNAME || 'C00000100'), 'Content-Type':'application/json', Accept:'application/json' },
    body: JSON.stringify(records)
  });
  const text = await r.text();
  let data; try { data = JSON.parse(text); } catch { data = { raw:text }; }
  if (!r.ok) throw new Error(`HTTP Ca.R.G.O.S. ${r.status}: ${text}`);
  return data;
}

async function estraiDatiDocumentoConAI(localPath, mimetype) {
  if (!process.env.OPENAI_API_KEY) throw new Error('OPENAI_API_KEY mancante su Render.');
  if (!fs.existsSync(localPath)) throw new Error('File OCR non trovato');

  const base64 = fs.readFileSync(localPath).toString('base64');
  const dataUrl = `data:${mimetype || 'image/jpeg'};base64,${base64}`;

  const prompt = `Leggi documento italiano patente/carta identitÃ . Rispondi SOLO JSON valido:
{
"tipo_documento":"","nome":"","cognome":"","data_nascita":"YYYY-MM-DD","luogo_nascita":"",
"codice_fiscale":"","numero_documento":"","ente_rilascio":"","data_rilascio":"YYYY-MM-DD",
"data_scadenza":"YYYY-MM-DD","numero_patente":"","categoria_patente":"","indirizzo":"",
"note":"","confidence":"alta|media|bassa"
}
Se un campo non Ã¨ visibile lascia vuoto.`;

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

app.get('/', async (req, res) => {
  try {
    const mezzi = await get(`SELECT COUNT(*) as tot FROM mezzi`);
    const pren = await get(`SELECT COUNT(*) as tot FROM prenotazioni`);
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
        <a class="tile" href="/import-mezzi"><span>&#128202;</span>Import Excel</a>
        <a class="tile" href="/cargos"><span>&#128666;</span>Ca.R.G.O.S.</a>
      </div>
      <div class="box" style="border:3px solid #c60000"><h2>VERSIONE ATTIVA: V24 OCR + CARGOS</h2><p class="ok">Se vedi questo riquadro, Render ha preso la versione nuova.</p></div>
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

app.get('/import-mezzi', (req, res) => {
  res.send(page('Import Excel', `
    <div class="box">
      <h2>Import mezzi da Excel</h2>
      <p>Legge colonne: UID, Targa, Marca, Modello, Km percor, Codice Tipo, Descrizione.</p>
      <form method="POST" action="/import-mezzi" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx,.xls,.csv" required>
        <button>Carica e importa</button>
      </form>
    </div>
  `));
});
app.post('/import-mezzi', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.send('File mancante');
    const wb = XLSX.readFile(req.file.path);
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    let imported = 0;

    for (const r of rows) {
      const targa = normalize(r['Targa'] || r['targa']);
      if (!targa) continue;
      const cat = categoriaFromRow(r);
      const marca = normalize(r['Marca']);
      const modello = normalize(r['Modello']);
      const descPub = descrizionePubblica({ marca, modello, categoria: cat });
      const km = Number(r['Km percor'] || r['Km percorsi'] || r['Km'] || 0);

      await run(`
        INSERT INTO mezzi
        (uid,targa,km,km_attuali,marca,modello,cilindrata,alimentazione,codice_tipo,categoria,descrizione,descrizione_pubblica,posti,stazione,prezzo_giorno,km_inclusi,stato)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,COALESCE((SELECT stato FROM mezzi WHERE targa=?),'disponibile'))
        ON CONFLICT(targa) DO UPDATE SET
          uid=excluded.uid, km=excluded.km, km_attuali=excluded.km_attuali, marca=excluded.marca,
          modello=excluded.modello, cilindrata=excluded.cilindrata, alimentazione=excluded.alimentazione,
          codice_tipo=excluded.codice_tipo, categoria=excluded.categoria, descrizione=excluded.descrizione,
          descrizione_pubblica=COALESCE(mezzi.descrizione_pubblica, excluded.descrizione_pubblica),
          posti=excluded.posti, stazione=excluded.stazione, prezzo_giorno=COALESCE(mezzi.prezzo_giorno, excluded.prezzo_giorno),
          km_inclusi=COALESCE(mezzi.km_inclusi, excluded.km_inclusi)
      `, [
        normalize(r['UID']), targa, km, km, marca, modello, normalize(r['Cilindrata']),
        normalize(r['Alimentaz'] || r['Alimentazione']), normalize(r['Codice Tip'] || r['Codice Tipo']),
        cat, normalize(r['Descrizion'] || r['Descrizione'] || r['Immagini consegna']),
        descPub, cat === '9_POSTI' ? 9 : null, normalize(r['Stazione']),
        prezzoCategoria(cat), kmCategoria(cat), targa
      ]);
      imported++;
    }
    try { fs.unlinkSync(req.file.path); } catch {}
    res.send(page('Import completato', `<div class="box"><h2 class="ok">Import completato</h2><p>Mezzi importati/aggiornati: <b>${imported}</b></p><a class="btn" href="/mezzi-web">Vai ai mezzi</a></div>`));
  } catch (e) {
    res.status(500).send(page('Errore import', `<div class="box"><h2 class="bad">Errore import</h2><pre>${esc(e.message)}</pre></div>`));
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
      <td>â¬ ${euro(m.prezzo_giorno)}</td>
      <td>${esc(m.km_inclusi)}</td>
      <td>${alertBadge(m)}</td>
      <td>${esc(m.stato)}</td>
    </tr>`).join('');
  res.send(page('Mezzi', `
    <h2>Elenco mezzi</h2>
    <table><tr><th>ID</th><th>Targa</th><th>Marca</th><th>Modello</th><th>Categoria</th><th>Descrizione pubblica</th><th>Prezzo</th><th>Km/giorno</th><th>Alert</th><th>Stato</th></tr>${trs}</table>
  `));
});

app.get('/mezzo/:id', async (req, res) => {
  const m = await get(`SELECT * FROM mezzi WHERE id=?`, [req.params.id]);
  if (!m) return res.send('Mezzo non trovato');
  const files = await all(`SELECT * FROM allegati WHERE mezzo_id=? ORDER BY id DESC`, [m.id]);
  const lista = files.map(f => `<li>${esc(f.tipo)} - <a href="/uploads/${esc(f.filename)}" target="_blank">${esc(f.originalname)}</a> ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Google Drive</a>` : ''}</li>`).join('');
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

app.get('/scadenze-mezzi', async (req, res) => {
  const rows = await all(`SELECT * FROM mezzi ORDER BY targa`);
  const trs = rows.map(m => `<tr><td><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a></td><td>${esc(descrizionePubblica(m))}</td><td>${esc(m.km_attuali || m.km)}</td><td>${esc(m.tagliando_km_scadenza)}</td><td>${esc(m.tagliando_data_scadenza)}</td><td>${esc(m.revisione_scadenza)}</td><td>${esc(m.bollo_scadenza)}</td><td>${esc(m.assicurazione_scadenza)}</td><td>${alertBadge(m)}</td></tr>`).join('');
  const alerts = rows.map(m => alertMezzo(m) ? `<div><b>${esc(m.targa)} ${esc(m.modello)}</b>${alertMezzo(m)}</div>` : '').join('');
  res.send(page('Scadenze mezzi', `<h2>Scadenze e manutenzioni mezzi</h2><div class="box">${alerts || '<p class="ok">Nessun alert attivo.</p>'}</div><table><tr><th>Targa</th><th>Descrizione</th><th>Km</th><th>Tagliando km</th><th>Tagliando data</th><th>Revisione</th><th>Bollo</th><th>Assicurazione</th><th>Alert</th></tr>${trs}</table>`));
});

function formPrenotazione(mezzi, selectedMezzo, selectedData, action) {
  const opt = mezzi.map(m => `<option value="${m.id}" ${String(m.id)===String(selectedMezzo)?'selected':''}>${esc(m.targa)} - ${esc(descrizionePubblica(m))}</option>`).join('');
  return `<form method="POST" action="${action}">
    <div class="grid">
      <div><label>Nome</label><input name="nome" required></div>
      <div><label>Cognome</label><input name="cognome" required></div>
      <div><label>Telefono</label><input name="telefono" required></div>
      <div><label>Email</label><input name="email" type="email"></div>
      <div><label>Codice fiscale</label><input name="codice_fiscale"></div>
      <div><label>Indirizzo</label><input name="indirizzo"></div>
      <div><label>CittÃ </label><input name="citta"></div>
      <div><label>CAP</label><input name="cap"></div>
      <div><label>Tipo cliente</label><select name="tipo_cliente"><option>privato</option><option>azienda</option></select></div>
      <div><label>P.IVA</label><input name="piva"></div>
      <div><label>Ragione sociale</label><input name="ragione_sociale"></div>
      <div><label>PEC</label><input name="pec"></div>
      <div><label>SDI</label><input name="sdi"></div>
      <div><label>Conducente 1</label><input name="conducente1"></div>
      <div><label>Patente 1</label><input name="patente1"></div>
      <div><label>Scadenza patente 1</label><input type="date" name="patente1_scadenza"></div>
      <div><label>Conducente 2</label><input name="conducente2"></div>
      <div><label>Patente 2</label><input name="patente2"></div>
      <div><label>Scadenza patente 2</label><input type="date" name="patente2_scadenza"></div>
      <div><label>Mezzo</label><select name="mezzo_id">${opt}</select></div>
      <div><label>Data inizio</label><input type="date" name="data_inizio" value="${esc(selectedData || '')}" required></div>
      <div><label>Ora inizio</label><input type="time" name="ora_inizio" value="08:30"></div>
      <div><label>Data fine</label><input type="date" name="data_fine" value="${esc(selectedData || '')}" required></div>
      <div><label>Ora fine</label><input type="time" name="ora_fine" value="18:00"></div>
      <div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
      <div><label>Carburante uscita</label><select name="carburante_uscita">${fuelOptions('4/4 pieno')}</select></div>
    </div>
    <label>Note</label><textarea name="note"></textarea>
    <button>Crea contratto</button>
  </form>`;
}
app.get('/nuova-prenotazione', async (req, res) => {
  const mezzi = await all(`SELECT * FROM mezzi ORDER BY categoria,targa`);
  res.send(page('Nuova prenotazione', `<h2>Nuova prenotazione / contratto</h2>${req.query.data ? `<p class="notice">Aperta dal planning per il giorno <b>${esc(req.query.data)}</b>.</p>` : ''}${formPrenotazione(mezzi, req.query.mezzo_id, req.query.data, '/prenota-admin')}`));
});
app.post('/prenota-admin', async (req, res) => {
  try {
    const b = req.body;
    const erroreDate = validDateRange(b.data_inizio, b.data_fine);
    if (erroreDate) return res.send(page('Errore date', `<div class="box"><h2 class="bad">${esc(erroreDate)}</h2><a class="btn" href="/nuova-prenotazione">Torna</a></div>`));
    const mezzo = await get(`SELECT * FROM mezzi WHERE id=?`, [b.mezzo_id]);
    if (!mezzo) return res.send('Mezzo non trovato');
    const occ = await queryDisponibilita(b.mezzo_id, b.data_inizio, b.data_fine);
    if (occ) return res.send(page('Occupato', `<div class="box"><h2 class="bad">Mezzo occupato in queste date</h2><p>Contratto: <a href="/prenotazione/${occ.id}">${esc(occ.codice)}</a></p><a class="btn" href="/planning">Vai al planning</a></div>`));
    const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.ora_inizio, b.ora_fine, b.km_previsti);
    const result = await run(`
      INSERT INTO prenotazioni
      (codice,nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,tipo_cliente,piva,ragione_sociale,pec,sdi,conducente1,patente1,patente1_scadenza,conducente2,patente2,patente2_scadenza,mezzo_id,data_inizio,data_fine,ora_inizio,ora_fine,giorni,km_previsti,extra_fuori_orario,extra_km,imponibile,iva,totale,cauzione,carburante_uscita,stato,note)
      VALUES ('TEMP',?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `, [b.nome,b.cognome,b.telefono,b.email,b.codice_fiscale,b.indirizzo,b.citta,b.cap,b.tipo_cliente,b.piva,b.ragione_sociale,b.pec,b.sdi,b.conducente1,b.patente1,b.patente1_scadenza,b.conducente2,b.patente2,b.patente2_scadenza,b.mezzo_id,b.data_inizio,b.data_fine,b.ora_inizio,b.ora_fine,calc.giorni,Number(b.km_previsti || 0),calc.extra_fuori_orario,calc.extraKm,calc.imponibile,calc.iva,calc.totale,CAUZIONE,b.carburante_uscita || '4/4 pieno','bozza',b.note]);
    const cod = codicePratica(result.lastID);
    await run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, result.lastID]);
    res.send(actionScreen(result.lastID, 'Contratto creato', `Codice: <b>${cod}</b><br>Totale: <b>â¬ ${euro(calc.totale)}</b>`));
  } catch (e) {
    res.status(500).send(page('Errore prenotazione', `<div class="box"><h2 class="bad">Errore prenotazione</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/prenota', (req, res) => {
  res.send(page('Prenota DP RENT', `
    <h2>Richiesta prenotazione cliente</h2>
    <p class="notice">Il cliente sceglie solo la categoria. La targa resta interna.</p>
    <form method="POST" action="/prenota-cliente">
      <div class="grid">
        <div><label>Tipo mezzo</label><select name="categoria" required><option value="FURGONE">Furgone cargo/merci</option><option value="9_POSTI">Pulmino 9 posti</option><option value="AUTO_DACIA">Auto economica</option><option value="AUTO_GOLF">Auto categoria Golf</option><option value="ESCAVATORE">Escavatore</option><option value="SEMOVENTE">Piattaforma / semovente</option></select></div>
        <div><label>Nome</label><input name="nome" required></div><div><label>Cognome</label><input name="cognome" required></div><div><label>Telefono</label><input name="telefono" required></div><div><label>Email</label><input name="email"></div><div><label>Codice fiscale</label><input name="codice_fiscale"></div><div><label>Indirizzo</label><input name="indirizzo"></div><div><label>Data inizio</label><input type="date" name="data_inizio" required></div><div><label>Ora inizio</label><input type="time" name="ora_inizio" value="08:30"></div><div><label>Data fine</label><input type="date" name="data_fine" required></div><div><label>Ora fine</label><input type="time" name="ora_fine" value="18:00"></div><div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
      </div><button>Invia richiesta</button></form>`));
});
app.post('/prenota-cliente', async (req, res) => {
  try {
    const b = req.body;
    const erroreDate = validDateRange(b.data_inizio, b.data_fine);
    if (erroreDate) return res.send(page('Errore date', `<h2 class="bad">${esc(erroreDate)}</h2>`));
    const mezzi = await all(`SELECT * FROM mezzi WHERE categoria=? ORDER BY id ASC`, [b.categoria]);
    if (!mezzi.length) return res.send(page('Nessun mezzo', '<h2>Nessun mezzo per questa categoria</h2>'));
    let mezzo = null;
    for (const m of mezzi) {
      const occ = await queryDisponibilita(m.id, b.data_inizio, b.data_fine);
      if (!occ) { mezzo = m; break; }
    }
    if (!mezzo) return res.send(page('Non disponibile', '<h2>Nessun mezzo libero nelle date richieste</h2>'));
    const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.ora_inizio, b.ora_fine, b.km_previsti);
    const result = await run(`
      INSERT INTO prenotazioni
      (codice,nome,cognome,telefono,email,codice_fiscale,indirizzo,tipo_cliente,mezzo_id,data_inizio,data_fine,ora_inizio,ora_fine,giorni,km_previsti,extra_fuori_orario,extra_km,imponibile,iva,totale,cauzione,stato)
      VALUES ('TEMP',?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    `, [b.nome,b.cognome,b.telefono,b.email,b.codice_fiscale,b.indirizzo,'privato',mezzo.id,b.data_inizio,b.data_fine,b.ora_inizio || '08:30',b.ora_fine || '18:00',calc.giorni,Number(b.km_previsti || 0),calc.extra_fuori_orario,calc.extraKm,calc.imponibile,calc.iva,calc.totale,CAUZIONE,'richiesta_cliente']);
    const cod = codicePratica(result.lastID);
    await run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, result.lastID]);
    res.send(page('Richiesta inviata', `<div class="box"><h2 class="ok">Richiesta inviata</h2><p>Codice: <b>${cod}</b></p><p>Totale previsto: <b>â¬ ${euro(calc.totale)}</b></p><p>DP RENT confermerÃ  la prenotazione.</p></div>`));
  } catch (e) {
    res.status(500).send(page('Errore', `<pre>${esc(e.message)}</pre>`));
  }
});

app.get('/prenotazione/:id', async (req, res) => {
  const p = await get(`SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.descrizione_pubblica FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE p.id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');
  res.send(page('Dettaglio contratto', `
    <div class="box">
      <h2>Contratto ${esc(p.codice)}</h2>
      <p><b>Cliente:</b> ${esc(p.nome)} ${esc(p.cognome)} - ${esc(p.telefono)}</p>
      <p><b>Email:</b> ${esc(p.email)} | <b>CF:</b> ${esc(p.codice_fiscale)}</p>
      <p><b>Mezzo:</b> <a href="/mezzo/${p.mezzo_id}">${esc(p.targa)} ${esc(descrizionePubblica(p))}</a></p>
      <p><b>Date:</b> ${esc(p.data_inizio)} ore ${esc(p.ora_inizio)} â ${esc(p.data_fine)} ore ${esc(p.ora_fine)}</p>
      <p><b>Totale:</b> â¬ ${euro(p.totale)} | <b>Cauzione:</b> â¬ ${euro(p.cauzione || CAUZIONE)}</p>
      <p><b>Stato:</b> ${esc(p.stato)} | <b>Nexi:</b> ${esc(p.nexi_stato || '')} | <b>Ca.R.G.O.S.:</b> ${esc(p.cargos_stato || '')}</p>
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
        <a class="btn btnWarn" href="/nexi/${p.id}">Nexi Pay Link</a>
        <a class="btn btn3" href="/firma-link/${p.id}">Link firma WhatsApp</a>
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
  const trs = rows.map(p => `<tr><td><a href="/prenotazione/${p.id}">${esc(p.codice)}</a></td><td>${esc(p.nome)} ${esc(p.cognome)}</td><td>${esc(p.telefono)}<br>${esc(p.email)}</td><td><b>${esc(p.targa)}</b><br>${esc(descrizionePubblica(p))}</td><td>${esc(p.data_inizio)} â ${esc(p.data_fine)}</td><td>â¬ ${euro(p.totale)}</td><td>${esc(p.stato)}</td><td><a href="/prenotazione/${p.id}">Apri</a><br><a href="/contratto/${p.id}">PDF</a><br><a href="/nexi/${p.id}">Nexi</a></td></tr>`).join('');
  res.send(page('Storico', `<h2>Storico contratti / prenotazioni</h2><form method="GET" action="/prenotazioni" class="box"><div class="grid"><input name="q" placeholder="Cerca nome, targa, codice, telefono" value="${esc(q)}"><select name="stato"><option value="">Tutti gli stati</option>${['bozza','richiesta_cliente','confermato','firmato','in_corso','rientrato','chiuso','pagato','annullato'].map(s=>`<option ${stato===s?'selected':''}>${s}</option>`).join('')}</select><input type="date" name="dal" value="${esc(dal)}"><input type="date" name="al" value="${esc(al)}"></div><button>Cerca</button></form><table><tr><th>Codice</th><th>Cliente</th><th>Contatti</th><th>Mezzo</th><th>Date</th><th>Totale</th><th>Stato</th><th>Azioni</th></tr>${trs}</table>`));
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
  res.send(page('Planning', `<h2>Planning ${start.format('MM/YYYY')}</h2><p><a href="/planning?mese=${prec}">â Mese precedente</a> | <a href="/planning?mese=${succ}">Mese successivo â</a></p><p><span class="libero" style="padding:6px;">Libero: clic per prenotare</span> <span class="occupato" style="padding:6px;">Occupato: clic per aprire contratto</span></p><div class="sticky-table"><table><tr>${header}</tr>${rows}</table></div>`));
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
        <option>Carta identitÃ </option>
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

      <button type="button" onclick="scattaFoto()">ð¸ Scatta foto</button>
      <button type="button" class="btn2" onclick="caricaDaDispositivo()">ð Carica da dispositivo</button>
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
      <p class="notice">Controlla bene: se una data o un numero Ã¨ sbagliato, correggilo prima di salvare.</p>

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
         patente1=COALESCE(NULLIF(?,''), patente1),
         patente1_scadenza=COALESCE(NULLIF(?,''), patente1_scadenza),
         note=COALESCE(note,'') || ?
     WHERE id=?`,
    [
      b.nome || current.nome,
      b.cognome || current.cognome,
      b.codice_fiscale || current.codice_fiscale,
      b.indirizzo || current.indirizzo,
      b.numero_patente || '',
      b.data_scadenza || '',
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


app.get('/cliente-documenti/:id/:token', async (req, res) => {
  const expected = clienteDocToken(req.params.id);
  if (req.params.token !== expected) return res.status(403).send('Link non valido');

  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');

  const files = await all(`SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id DESC`, [p.id]);
  const lista = files.map(f => `<li>${esc(f.tipo)} - ${esc(f.originalname)} ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Drive</a>` : ''}</li>`).join('');

  res.send(page('Carica documenti DP RENT', `
    <div class="box">
      <h2>DP RENT - Carica documenti</h2>
      <p><b>Contratto:</b> ${esc(p.codice)}</p>
      <p><b>Cliente:</b> ${esc(p.nome)} ${esc(p.cognome)}</p>
      <p class="notice">Puoi caricare patente/documento. Dopo la lettura automatica controlli e confermi i dati.</p>
    </div>
    ${renderOcrUploadForm(`/cliente-documenti/${p.id}/${req.params.token}`, `/cliente-documenti/${p.id}/${req.params.token}`, 'Carica/scatta documento')}
    <div class="box"><h3>File giÃ  caricati</h3><ul>${lista || '<li>Nessun file caricato</li>'}</ul></div>
  `));
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
        <p>Grazie. DP RENT controllerÃ  i dati e completerÃ  il contratto.</p>
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
      <a class="btn btn2" href="/prenotazione/${p.id}">Torna contratto</a>
    </div>
  `));
});

app.get('/documenti/:id', async (req, res) => {
  const files = await all(`SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id DESC`, [req.params.id]);
  const lista = files.map(f => `<li>${esc(f.tipo)} - <a href="/uploads/${esc(f.filename)}" target="_blank">${esc(f.originalname)}</a> ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Google Drive</a>` : ''}</li>`).join('');

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

  try {
    driveRes = await uploadFileToDrive(
      req.file.path,
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
      req.file.filename,
      req.file.originalname,
      req.file.path,
      req.file.mimetype,
      driveRes?.id || null,
      driveRes?.webViewLink || null
    ]
  );

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
    const file = await generaPdfContratto(req.params.id, { forceDrive: true });
    res.download(file);
  } catch (e) {
    res.status(500).send(page('Errore PDF', `<div class="box"><h2 class="bad">Errore PDF</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/firma/:id', (req, res) => {
  res.send(page('Firma', `
    <div class="box">
      <h2>Firma contratto</h2>
      <canvas id="canvas"></canvas>
      <button type="button" onclick="clearCanvas()">Cancella</button>
      <button type="button" onclick="saveFirma()">Salva firma</button>
      <a class="btn btn2" href="/prenotazione/${req.params.id}">Torna contratto</a>
    </div>
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
        if (r.ok) location.href = '/prenotazione/${req.params.id}';
        else alert('Errore salvataggio firma');
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
    await generaPdfContratto(req.params.id, { forceDrive: true });
    res.send('OK');
  } catch (e) {
    res.status(500).send(e.message);
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
      <a class="btn btn2" href="/prenotazione/${p.id}">Torna contratto</a>
    </div>
  `));
});

app.get('/whatsapp-contratto/:id', async (req, res) => {
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id]);
  if (!p) return res.send('Contratto non trovato');

  let pdfLink = p.pdf_drive_web_link || '';
  if (!pdfLink) {
    try {
      await generaPdfContratto(p.id);
      const updated = await get(`SELECT pdf_drive_web_link FROM prenotazioni WHERE id=?`, [p.id]);
      pdfLink = updated?.pdf_drive_web_link || '';
    } catch (e) {
      console.log('Errore generazione PDF per WhatsApp:', e.message);
    }
  }

  const firmaLink = absoluteUrl(req, `/firma/${p.id}`);
  const testo =
    `DP RENT - Contratto ${p.codice}\n` +
    `Cliente: ${p.nome || ''} ${p.cognome || ''}\n` +
    `Totale: â¬ ${Number(p.totale || 0).toFixed(2)}\n\n` +
    (pdfLink ? `PDF contratto: ${pdfLink}\n\n` : '') +
    `Firma online: ${firmaLink}`;

  res.redirect(whatsappText(testo));
});



app.get('/email/:id', async (req,res)=>{
  const p = await get(`SELECT * FROM prenotazioni WHERE id=?`,[req.params.id]);
  if(!p)return res.send('Contratto non trovato');
  res.send(page('Invia email', `<h2>Invia contratto via email</h2><form method="POST" action="/email/${p.id}"><label>Email destinatario</label><input name="email" value="${esc(p.email)}" required><label>Messaggio</label><textarea name="messaggio">Buongiorno, in allegato trova il contratto DP RENT.</textarea><button>Invia email</button></form><p class="notice">Per inviare davvero configura SMTP su Render.</p>`));
});
app.post('/email/:id', async (req,res)=>{
  try {
    const file = await generaPdfContratto(req.params.id, { forceDrive: true });
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
      `Totale: â¬ ${euro(p.totale)}\n` +
      `${pagamento.link}`;

    res.send(page('Pagamento Nexi', `
      <div class="box">
        <h2>Pagamento Nexi PayMail</h2>
        <p><b>Contratto:</b> ${esc(p.codice)}</p>
        <p><b>Totale contratto:</b> â¬ ${euro(p.totale)}</p>
        <p class="notice">La cauzione resta gestita manualmente. Qui paghi solo il totale contratto.</p>

        <a class="btn btnWarn" href="${esc(pagamento.link)}" target="_blank">Apri link pagamento Nexi</a>
        <a class="btn btn3" target="_blank" href="${esc(whatsappText(testoWa))}">Invia pagamento WhatsApp</a>

        <label>Link pagamento</label>
        <input value="${esc(pagamento.link)}" readonly onclick="this.select()">

        <a class="btn btn2" href="/prenotazione/${p.id}">Torna contratto</a>
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



app.get('/cargos-config', (req, res) => {
  res.send(page('Config Ca.R.G.O.S.', `
    <div class="box">
      <h2>Configurazione Ca.R.G.O.S.</h2>
      <p><b>Utente impostato:</b> ${esc(process.env.CARGOS_USERNAME || 'C00000100')}</p>
      <p class="${process.env.CARGOS_PASSWORD ? 'ok' : 'bad'}">Password: ${process.env.CARGOS_PASSWORD ? 'presente' : 'mancante'}</p>
      <p class="${process.env.CARGOS_APIKEY ? 'ok' : 'bad'}">APIKEY: ${process.env.CARGOS_APIKEY ? 'presente' : 'mancante'}</p>
      <p class="${process.env.CARGOS_AGENZIA_ID ? 'ok' : 'bad'}">AGENZIA_ID: ${process.env.CARGOS_AGENZIA_ID ? 'presente' : 'mancante'}</p>
      <p class="${process.env.CARGOS_OPERATORE_ID ? 'ok' : 'bad'}">OPERATORE_ID: ${process.env.CARGOS_OPERATORE_ID ? 'presente' : 'mancante'}</p>
      <p class="${process.env.CARGOS_LUOGO_COD ? 'ok' : 'bad'}">LUOGO_COD: ${process.env.CARGOS_LUOGO_COD ? 'presente' : 'mancante'}</p>
      <hr>
      <p>Environment da mettere su Render:</p>
      <pre>CARGOS_USERNAME=C00000100
CARGOS_PASSWORD=la_password_che_hai
CARGOS_APIKEY=da richiedere/recuperare
CARGOS_AGENZIA_ID=da Ca.R.G.O.S.
CARGOS_OPERATORE_ID=da Ca.R.G.O.S.
CARGOS_LUOGO_COD=codice luogo polizia
CARGOS_BASE_URL=https://cargos.poliziadistato.it/CARGOS_API</pre>
      <a class="btn" href="/cargos">Vai a Ca.R.G.O.S.</a>
    </div>
  `));
});

app.get('/cargos', async (req, res) => {
  const rows = await all(`SELECT p.*, m.targa FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id ORDER BY p.id DESC LIMIT 50`);
  const trs = rows.map(p => `<tr><td><a href="/prenotazione/${p.id}">${esc(p.codice)}</a></td><td>${esc(p.nome)} ${esc(p.cognome)}</td><td>${esc(p.targa)}</td><td>${esc(p.data_inizio)} â ${esc(p.data_fine)}</td><td>${esc(p.cargos_stato || '')}</td><td><a class="btn" href="/cargos/record/${p.id}">Record</a><a class="btn btn2" href="/cargos/check/${p.id}">Check</a><a class="btn btnWarn" href="/cargos/send/${p.id}">Send</a></td></tr>`).join('');
  res.send(page('Ca.R.G.O.S.', `<div class="box"><h2>Ca.R.G.O.S.</h2><p>Modulo pronto. Quando hai username/password/APIKEY e codici tabelle, Check e Send diventano reali.</p><p><b>Configurato:</b> ${cargosConfigured() ? '<span class="ok">SI</span>' : '<span class="bad">NO</span>'}</p><p>Servono: CARGOS_USERNAME, CARGOS_PASSWORD, CARGOS_APIKEY, CARGOS_AGENZIA_ID, CARGOS_OPERATORE_ID, CARGOS_LUOGO_COD.</p></div><table><tr><th>Contratto</th><th>Cliente</th><th>Targa</th><th>Date</th><th>Stato</th><th>Azione</th></tr>${trs}</table>`));
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
    res.setHeader('Content-Disposition', `attachment; filename="cargos_${p?.codice || req.params.id}.txt"`);
    res.send(record + '\\n');
  } catch(e) { res.status(500).send('Errore export CARGOS: ' + e.message); }
});

app.get('/cargos/check/:id', async (req, res) => {
  try {
    const result = await cargosSendRecords([await buildCargosRecordForContract(req.params.id)], 'Check');
    await run(`UPDATE prenotazioni SET cargos_stato=? WHERE id=?`, ['check_ok', req.params.id]);
    res.send(page('Check Ca.R.G.O.S.', `<div class="box"><h2>Esito Check</h2><pre>${esc(JSON.stringify(result,null,2))}</pre><a class="btn" href="/prenotazione/${req.params.id}">Torna contratto</a></div>`));
  } catch(e) { res.status(500).send(page('Errore Check Ca.R.G.O.S.', `<div class="box"><h2 class="bad">Errore Check</h2><pre>${esc(e.message)}</pre><a class="btn" href="/cargos">Torna</a></div>`)); }
});

app.get('/cargos/send/:id', async (req, res) => {
  try {
    const result = await cargosSendRecords([await buildCargosRecordForContract(req.params.id)], 'Send');
    await run(`UPDATE prenotazioni SET cargos_stato=? WHERE id=?`, ['send_ok', req.params.id]);
    res.send(page('Send Ca.R.G.O.S.', `<div class="box"><h2>Esito Send</h2><pre>${esc(JSON.stringify(result,null,2))}</pre><a class="btn" href="/prenotazione/${req.params.id}">Torna contratto</a></div>`));
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

app.listen(PORT, '0.0.0.0', () => console.log('DP RENT APP COMPLETA ONLINE su porta ' + PORT));
