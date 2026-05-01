
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

const app = express();
app.use(bodyParser.urlencoded({ extended: true, limit: '30mb' }));
app.use(bodyParser.json({ limit: '30mb' }));

const PORT = process.env.PORT || 3000;

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

const TERMS_URL = 'https://carrentalsoftware.myappy.it/data/public/user/65996976/terms_file.pdf';
const PRIVACY_URL = 'https://carrentalsoftware.myappy.it/data/public/user/65996976/privacy_file.pdf';

const db = new sqlite3.Database('./database.sqlite');

const uploadDir = path.join(__dirname, 'uploads');
const contractsDir = path.join(__dirname, 'contracts');
const firmeDir = path.join(__dirname, 'firme');
const publicDir = path.join(__dirname, 'public');

[uploadDir, contractsDir, firmeDir, publicDir].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir);
});

app.use('/public', express.static(publicDir));

const upload = multer({ dest: uploadDir });
const logoUpload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, publicDir),
    filename: (req, file, cb) => cb(null, 'logo.png')
  })
});

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
    ['descrizione_pubblica','TEXT'],['posti','INTEGER'],['km_attuali','INTEGER DEFAULT 0'],
    ['tagliando_km_scadenza','INTEGER'],['tagliando_data_scadenza','TEXT'],['revisione_scadenza','TEXT'],
    ['bollo_scadenza','TEXT'],['assicurazione_scadenza','TEXT'],['gomme_scadenza','TEXT'],
    ['manutenzione_note','TEXT'],['alert_giorni','INTEGER DEFAULT 30'],['alert_km','INTEGER DEFAULT 1000']
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
      mezzo_id INTEGER,
      data_inizio TEXT,
      data_fine TEXT,
      ora_inizio TEXT,
      ora_fine TEXT,
      giorni INTEGER,
      km_previsti INTEGER,
      extra_fuori_orario REAL DEFAULT 0,
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
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  [
    ['nome','TEXT'],['cognome','TEXT'],['email','TEXT'],['codice_fiscale','TEXT'],['indirizzo','TEXT'],
    ['citta','TEXT'],['cap','TEXT'],['tipo_cliente','TEXT'],['piva','TEXT'],['ragione_sociale','TEXT'],
    ['ora_inizio','TEXT'],['ora_fine','TEXT'],['firma_path','TEXT'],['pdf_path','TEXT'],['note','TEXT'],
    ['extra_fuori_orario','REAL DEFAULT 0'],['carburante_uscita','TEXT DEFAULT "4/4 pieno"'],
    ['carburante_rientro','TEXT DEFAULT "4/4 pieno"'],['km_uscita','INTEGER'],['km_rientro','INTEGER']
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
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);
  addColumn('allegati','mezzo_id','INTEGER');
});

function esc(v) {
  return String(v === undefined || v === null ? '' : v)
    .replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;');
}

function normalize(v) {
  return v === undefined || v === null ? '' : String(v).trim();
}

function page(title, content) {
  const logoPath = path.join(publicDir, 'logo.png');
  const logoHtml = fs.existsSync(logoPath)
    ? `<img src="/public/logo.png" style="height:42px;max-width:150px;object-fit:contain;background:white;border-radius:6px;padding:4px;">`
    : `<span style="font-size:28px;font-weight:900;color:white;">DP RENT</span>`;

  return `
<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<title>${title}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{font-family:Arial;margin:0;background:#f4f4f4;color:#222;}
header{background:#111;color:white;padding:14px 18px;display:flex;align-items:center;gap:14px;}
header h1{margin:0;font-size:22px;letter-spacing:1px;}
nav{background:#b30000;padding:12px;display:flex;gap:14px;flex-wrap:wrap;}
nav a{color:white;text-decoration:none;font-weight:bold;}
main{padding:20px;}
.box{background:white;padding:20px;margin-bottom:20px;border-radius:10px;box-shadow:0 2px 8px #ccc;}
table{width:100%;border-collapse:collapse;background:white;}
th,td{padding:9px;border:1px solid #ddd;font-size:13px;}
th{background:#222;color:white;}
input,select,textarea,button{padding:10px;margin:5px 0;width:100%;box-sizing:border-box;}
button,.btn{background:#b30000;color:white;border:0;padding:10px;border-radius:6px;text-decoration:none;display:inline-block;font-weight:bold;cursor:pointer;margin:4px 4px 4px 0;}
.btn2{background:#333;}.btn3{background:#0b6b2d;}
.ok{color:green;font-weight:bold;}.bad{color:red;font-weight:bold;}
.warn{color:#b36b00;font-weight:bold;}
.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;}
.libero{background:#1fae4b;color:white;text-align:center;font-weight:bold;cursor:pointer;}
.occupato{background:#d90000;color:white;text-align:center;font-weight:bold;cursor:pointer;}
.libero:hover{outline:3px solid #0b6b2d;filter:brightness(1.1);}
.occupato:hover{outline:3px solid #900;filter:brightness(1.1);}
.sticky-table{overflow:auto;max-height:75vh;border:1px solid #ccc;}
.sticky-table th{position:sticky;top:0;z-index:3;}
.sticky-col{position:sticky;left:0;background:#fff;z-index:2;min-width:150px;}
th.sticky-col{background:#222;color:white;z-index:4;}
canvas{border:2px solid #333;background:white;width:100%;height:220px;}
.actions{display:flex;gap:8px;flex-wrap:wrap;}
.notice{background:#fff3cd;border:1px solid #ffeeba;padding:10px;border-radius:8px;margin:10px 0;}
.alert{background:#ffe0e0;border:1px solid #d90000;padding:10px;border-radius:8px;margin:6px 0;}
.badge{display:inline-block;padding:4px 7px;border-radius:5px;font-size:12px;margin:2px;background:#eee;}
.badge-red{background:#d90000;color:white}.badge-orange{background:#ffb000;color:#111}.badge-green{background:#1fae4b;color:white}
@media(max-width:700px){.grid{grid-template-columns:1fr;}main{padding:10px;}table{font-size:11px;} th,td{padding:6px;}}
</style>
</head>
<body>
<header>${logoHtml}<h1>DP RENT APP</h1></header>
<nav>
<a href="/">Dashboard</a>
<a href="/mezzi-web">Mezzi</a>
<a href="/scadenze-mezzi">Scadenze</a>
<a href="/import-mezzi">Import Excel</a>
<a href="/nuova-prenotazione">Nuova prenotazione</a>
<a href="/prenotazioni">Storico</a>
<a href="/planning">Planning</a>
<a href="/prenota">Pagina cliente</a>
<a href="/logo">Logo</a>
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
  if (codice.includes('P') || desc.includes('PERSONE') || desc.includes('9P')) return '9_POSTI';
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
  return modello;
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

function extraOrario(ora) {
  if (!ora) return 0;
  const parts = String(ora).split(':').map(Number);
  const minuti = (parts[0] || 0) * 60 + (parts[1] || 0);
  const inizio = 8 * 60 + 30;
  const fine = 18 * 60 + 30;
  return (minuti < inizio || minuti > fine) ? EXTRA_FUORI_ORARIO : 0;
}

function calcolaTotale(mezzo, data_inizio, data_fine, ora_inizio, ora_fine, km_previsti) {
  const giorni = moment(data_fine).diff(moment(data_inizio), 'days') + 1;
  let imponibile = giorni * Number(mezzo.prezzo_giorno || 0);
  const kmInclusiTot = giorni * Number(mezzo.km_inclusi || 0);
  const kmPrev = Number(km_previsti || 0);
  if (mezzo.km_inclusi > 0 && kmPrev > kmInclusiTot) imponibile += (kmPrev - kmInclusiTot) * EXTRA_KM;
  const extra = extraOrario(ora_inizio) + extraOrario(ora_fine);
  imponibile += extra;
  const iva = imponibile * IVA;
  const totale = imponibile + iva;
  return { giorni, imponibile, iva, totale, extra_fuori_orario: extra };
}

function queryDisponibilita(mezzo_id, data_inizio, data_fine, cb) {
  db.get(`
    SELECT * FROM prenotazioni
    WHERE mezzo_id = ?
    AND stato != 'annullato'
    AND date(data_inizio) <= date(?)
    AND date(data_fine) >= date(?)
  `, [mezzo_id, data_fine, data_inizio], cb);
}

function fuelOptions(selected) {
  const vals = ['4/4 pieno','3/4','1/2','1/4','Riserva','Vuoto'];
  return vals.map(v => `<option value="${v}" ${selected===v?'selected':''}>${v}</option>`).join('');
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
    const diff = d.diff(today, 'days');
    if (diff < 0) out.push(`<div class="alert">❌ ${label} scaduto il ${esc(val)}</div>`);
    else if (diff <= alertGiorni) out.push(`<div class="alert">⚠️ ${label} in scadenza: ${esc(val)} (${diff} giorni)</div>`);
  }

  checkDate('Tagliando data', m.tagliando_data_scadenza);
  checkDate('Revisione', m.revisione_scadenza);
  checkDate('Bollo', m.bollo_scadenza);
  checkDate('Assicurazione', m.assicurazione_scadenza);
  checkDate('Gomme/manutenzione', m.gomme_scadenza);

  if (m.tagliando_km_scadenza) {
    const diffKm = Number(m.tagliando_km_scadenza) - kmAtt;
    if (diffKm <= 0) out.push(`<div class="alert">❌ Tagliando km scaduto: km attuali ${kmAtt}, scadenza ${m.tagliando_km_scadenza}</div>`);
    else if (diffKm <= alertKm) out.push(`<div class="alert">⚠️ Tagliando vicino: mancano ${diffKm} km</div>`);
  }

  return out.join('');
}

function alertBadge(m) {
  const has = alertMezzo(m);
  return has ? `<span class="badge badge-red">ALERT</span>` : `<span class="badge badge-green">OK</span>`;
}

function actionScreen(id, titolo, messaggio) {
  return page(titolo, `
    <div class="box">
      <h2 class="ok">${titolo}</h2>
      <p>${messaggio || ''}</p>
      <div class="actions">
        <a class="btn" href="/contratto/${id}">Scarica / stampa PDF</a>
        <a class="btn btn2" href="/firma/${id}">Firma su tablet</a>
        <a class="btn btn2" href="/email/${id}">Invia via email</a>
        <a class="btn btn3" href="/documenti/${id}">Carica documenti cliente</a>
        <a class="btn btn3" href="/checkout/${id}">Check-out mezzo</a>
        <a class="btn btn3" href="/checkin/${id}">Check-in mezzo</a>
        <a class="btn btn2" href="/prenotazione/${id}">Dettaglio contratto</a>
        <a class="btn btn2" href="/prenotazioni">Vai allo storico</a>
      </div>
    </div>
  `);
}

function drawHeader(doc) {
  doc.rect(0,0,612,115).fill('#111111');
  const logoPath = path.join(publicDir, 'logo.png');
  if (fs.existsSync(logoPath)) {
    try { doc.image(logoPath, 35, 25, { fit: [145, 70] }); }
    catch(e) { doc.fillColor('white').fontSize(28).text('DP RENT', 45, 35); }
  } else {
    doc.fillColor('white').fontSize(28).text('DP RENT', 45, 35);
  }
  doc.fillColor('white').fontSize(11).text(AZIENDA.nome, 350, 25, { align:'right', width:210 });
  doc.text(AZIENDA.indirizzo, 350, 43, { align:'right', width:210 });
  doc.text(`P.IVA / CF ${AZIENDA.piva}  |  Tel. ${AZIENDA.telefono}`, 350, 61, { align:'right', width:210 });
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
  doc.fillColor('#111').fontSize(9).text(value || '', x+w*0.38, y, {width:w*0.6});
  doc.fillColor('black');
}
function generaPdfContratto(id, callback) {
  db.get(`
    SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.km_inclusi, m.descrizione_pubblica
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id = p.mezzo_id
    WHERE p.id = ?
  `, [id], (err, p) => {
    if (err || !p) return callback(err || new Error('Contratto non trovato'));
    const safe = String(p.codice || id).replace(/[^a-zA-Z0-9_-]/g, '');
    const file = path.join(contractsDir, `contratto_${safe}.pdf`);
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
    row(doc, 'Indirizzo', p.indirizzo || '', 55, yLeft, 220); yLeft += 18;
    row(doc, 'Comune/CAP', `${p.citta || ''} ${p.cap || ''}`, 55, yLeft, 220); yLeft += 18;
    row(doc, 'Tipo / P.IVA', `${p.tipo_cliente || ''} ${p.piva || ''}`, 55, yLeft, 220); yLeft += 18;

    yRight = section(doc, 'DETTAGLI PRENOTAZIONE', 310, yRight, 245);
    row(doc, 'Check-out', `${p.data_inizio} ore ${p.ora_inizio || '08:30'} - Narni`, 320, yRight, 220); yRight += 18;
    row(doc, 'Check-in', `${p.data_fine} ore ${p.ora_fine || '18:00'} - Narni`, 320, yRight, 220); yRight += 18;
    row(doc, 'Giorni tariffati', String(p.giorni || ''), 320, yRight, 220); yRight += 18;
    row(doc, 'Extra fuori orario', Number(p.extra_fuori_orario || 0) > 0 ? `euro ${Number(p.extra_fuori_orario).toFixed(2)} + IVA` : 'NO', 320, yRight, 220); yRight += 18;
    row(doc, 'Carburante', `${p.carburante_uscita || '4/4 pieno'} / ${p.carburante_rientro || '4/4 pieno'}`, 320, yRight, 220); yRight += 18;
    row(doc, 'Km uscita/rientro', `${p.km_uscita || ''} / ${p.km_rientro || ''}`, 320, yRight, 220); yRight += 18;

    y = Math.max(yLeft, yRight) + 10;
    let yVeh = y, yCost = y;

    yVeh = section(doc, 'VEICOLO', 45, yVeh, 245);
    row(doc, 'Targa', p.targa || '', 55, yVeh, 220); yVeh += 18;
    row(doc, 'Descrizione', p.descrizione_pubblica || `${p.marca || ''} ${p.modello || ''}`, 55, yVeh, 220); yVeh += 18;
    row(doc, 'Categoria', p.categoria || '', 55, yVeh, 220); yVeh += 18;
    row(doc, 'Km inclusi totali', String(Number(p.km_inclusi || 0) * Number(p.giorni || 0)), 55, yVeh, 220); yVeh += 18;
    row(doc, 'Km previsti', String(p.km_previsti || 0), 55, yVeh, 220); yVeh += 18;

    yCost = section(doc, 'RIEPILOGO ECONOMICO', 310, yCost, 245);
    row(doc, 'Imponibile', `euro ${Number(p.imponibile || 0).toFixed(2)}`, 320, yCost, 220); yCost += 18;
    row(doc, 'IVA 22%', `euro ${Number(p.iva || 0).toFixed(2)}`, 320, yCost, 220); yCost += 18;
    row(doc, 'Totale IVA inclusa', `euro ${Number(p.totale || 0).toFixed(2)}`, 320, yCost, 220); yCost += 18;
    row(doc, 'Deposito cauzionale', `euro ${Number(p.cauzione || CAUZIONE).toFixed(2)}`, 320, yCost, 220); yCost += 18;

    y = Math.max(yVeh, yCost) + 10;
    y = section(doc, 'CONDIZIONI GENERALI E PRIVACY', 45, y, 510);
    doc.fontSize(8).fillColor('#111').text('Il cliente dichiara di aver preso visione e accettare le condizioni generali di noleggio e l’informativa privacy DP RENT / Trasporti DP S.R.L.', 55, y, {width:490});
    y += 22;
    doc.fontSize(7).fillColor('#333').text(`Condizioni generali: ${TERMS_URL}`, 55, y, {width:490}); y += 12;
    doc.text(`Informativa privacy: ${PRIVACY_URL}`, 55, y, {width:490}); y += 18;
    doc.fontSize(8).fillColor('#111').text('Condizioni principali: veicolo consegnato con il pieno e da riconsegnare con il pieno; extra km euro 0,15/km ove previsto; danni, multe, pedaggi, franchigie, smarrimenti e costi accessori a carico del cliente.', 55, y, {width:490});
    y += 45;

    doc.fontSize(10).fillColor('#111').text('Firma cliente:', 55, y);
    if (p.firma_path && fs.existsSync(p.firma_path)) doc.image(p.firma_path, 55, y+15, { fit: [220, 70] });
    else doc.text('______________________________', 55, y+25);
    doc.text('Firma DP RENT:', 330, y);
    doc.text('______________________________', 330, y+25);
    doc.end();

    stream.on('finish', () => {
      db.run(`UPDATE prenotazioni SET pdf_path = ? WHERE id = ?`, [file, id]);
      callback(null, file);
    });
    stream.on('error', callback);
  });
}

// Dashboard
app.get('/', (req, res) => {
  db.get(`SELECT COUNT(*) as tot FROM mezzi`, [], (e1, mezzi) => {
    db.get(`SELECT COUNT(*) as tot FROM prenotazioni`, [], (e2, pren) => {
      db.all(`SELECT * FROM mezzi`, [], (e3, allMezzi) => {
        const alerts = (allMezzi || []).map(m => {
          const a = alertMezzo(m);
          return a ? `<div><b>${esc(m.targa)} ${esc(m.modello)}</b>${a}</div>` : '';
        }).join('');
        res.send(page('Dashboard', `
          <div class="box">
            <h2>Gestionale DP RENT attivo</h2>
            <p>Mezzi caricati: <b>${mezzi ? mezzi.tot : 0}</b></p>
            <p>Contratti / prenotazioni: <b>${pren ? pren.tot : 0}</b></p>
            <p class="notice">Versione test con SQLite/file locali. Quando è definitiva passiamo a dati permanenti.</p>
          </div>
          <div class="box">
            <h2>Alert mezzi</h2>
            ${alerts || '<p class="ok">Nessun alert mezzi.</p>'}
          </div>
        `));
      });
    });
  });
});

// Logo
app.get('/logo', (req, res) => {
  const hasLogo = fs.existsSync(path.join(publicDir, 'logo.png'));
  res.send(page('Logo', `
    <div class="box">
      <h2>Logo DP RENT</h2>
      ${hasLogo ? `<p>Logo attuale:</p><img src="/public/logo.png" style="max-width:240px;background:#eee;padding:10px;border-radius:8px;">` : `<p>Nessun logo caricato: nel PDF verrà usata la scritta DP RENT.</p>`}
      <form method="POST" action="/logo" enctype="multipart/form-data">
        <label>Carica logo PNG/JPG</label>
        <input type="file" name="logo" accept="image/png,image/jpeg" required>
        <button>Salva logo</button>
      </form>
    </div>
  `));
});
app.post('/logo', multer({storage: multer.diskStorage({destination:(req,file,cb)=>cb(null,publicDir), filename:(req,file,cb)=>cb(null,'logo.png')})}).single('logo'), (req, res) => res.redirect('/logo'));

// Import Excel
app.get('/import-mezzi', (req, res) => {
  res.send(page('Import Excel', `
    <div class="box">
      <h2>Import mezzi da Excel</h2>
      <form method="POST" action="/import-mezzi" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx,.xls,.csv" required>
        <button>Carica e importa</button>
      </form>
    </div>
  `));
});
app.post('/import-mezzi', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');
  const wb = XLSX.readFile(req.file.path);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
  let imported = 0;
  const stmt = db.prepare(`
    INSERT OR REPLACE INTO mezzi
    (uid,targa,km,km_attuali,marca,modello,cilindrata,alimentazione,codice_tipo,categoria,descrizione,descrizione_pubblica,posti,stazione,prezzo_giorno,km_inclusi,stato)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,COALESCE((SELECT stato FROM mezzi WHERE targa=?),'disponibile'))
  `);
  rows.forEach(r => {
    const targa = normalize(r['Targa']);
    if (!targa) return;
    const cat = categoriaFromRow(r);
    const modelloTemp = {marca: normalize(r['Marca']), modello: normalize(r['Modello']), categoria: cat};
    const descPub = descrizionePubblica(modelloTemp);
    let posti = cat === '9_POSTI' ? 9 : null;
    stmt.run([
      normalize(r['UID']), targa,
      Number(r['Km percor'] || r['Km percorsi'] || r['Km'] || 0),
      Number(r['Km percor'] || r['Km percorsi'] || r['Km'] || 0),
      normalize(r['Marca']), normalize(r['Modello']), normalize(r['Cilindrata']),
      normalize(r['Alimentaz'] || r['Alimentazione']),
      normalize(r['Codice Tip'] || r['Codice Tipo']),
      cat,
      normalize(r['Descrizion'] || r['Descrizione'] || r['Immagini consegna']),
      descPub,
      posti,
      normalize(r['Stazione']),
      prezzoCategoria(cat),
      kmCategoria(cat),
      targa
    ]);
    imported++;
  });
  stmt.finalize();
  fs.unlinkSync(req.file.path);
  res.send(page('Import completato', `<h2 class="ok">Import completato</h2><p>Mezzi importati/aggiornati: <b>${imported}</b></p><a class="btn" href="/mezzi-web">Vai ai mezzi</a>`));
});

// Mezzi + scheda
app.get('/mezzi-web', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY categoria,targa`, [], (err, rows) => {
    const trs = rows.map(m => `
      <tr>
        <td>${m.id}</td>
        <td><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a></td>
        <td>${esc(m.marca)}</td>
        <td>${esc(m.modello)}</td>
        <td>${esc(m.categoria)}</td>
        <td>${esc(descrizionePubblica(m))}</td>
        <td>euro ${Number(m.prezzo_giorno || 0).toFixed(2)}</td>
        <td>${m.km_inclusi}</td>
        <td>${alertBadge(m)}</td>
        <td>${esc(m.stato)}</td>
      </tr>`).join('');
    res.send(page('Mezzi', `
      <h2>Elenco mezzi</h2>
      <p>Clicca sulla targa per aprire la scheda mezzo con km, manutenzione, bollo, revisione e alert.</p>
      <table>
        <tr><th>ID</th><th>Targa</th><th>Marca</th><th>Modello</th><th>Categoria</th><th>Descrizione pubblica</th><th>Prezzo</th><th>Km/giorno</th><th>Alert</th><th>Stato</th></tr>
        ${trs}
      </table>
    `));
  });
});

app.get('/mezzo/:id', (req, res) => {
  db.get(`SELECT * FROM mezzi WHERE id=?`, [req.params.id], (err, m) => {
    if (!m) return res.send('Mezzo non trovato');
    db.all(`SELECT * FROM allegati WHERE mezzo_id=? ORDER BY id DESC`, [m.id], (e2, files) => {
      const lista = files.map(f => `<li>${esc(f.tipo)} - ${esc(f.originalname)}</li>`).join('');
      res.send(page('Scheda mezzo', `
        <div class="box">
          <h2>Scheda mezzo ${esc(m.targa)} ${alertBadge(m)}</h2>
          ${alertMezzo(m) || '<p class="ok">Nessun alert attivo.</p>'}
          <form method="POST" action="/mezzo/${m.id}">
            <div class="grid">
              <div><label>Targa</label><input name="targa" value="${esc(m.targa)}" required></div>
              <div><label>Marca</label><input name="marca" value="${esc(m.marca)}"></div>
              <div><label>Modello</label><input name="modello" value="${esc(m.modello)}"></div>
              <div><label>Categoria</label><select name="categoria">
                ${['FURGONE','9_POSTI','AUTO_DACIA','AUTO_GOLF','ESCAVATORE','SEMOVENTE'].map(c=>`<option ${m.categoria===c?'selected':''}>${c}</option>`).join('')}
              </select></div>
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
            <label>Note manutenzione</label>
            <textarea name="manutenzione_note">${esc(m.manutenzione_note)}</textarea>
            <button>Salva scheda mezzo</button>
          </form>
          <hr>
          <h3>Foto/documenti mezzo</h3>
          <form method="POST" action="/mezzo/${m.id}/foto" enctype="multipart/form-data">
            <select name="tipo">
              <option>Foto mezzo fronte</option><option>Foto mezzo retro</option><option>Foto lato dx</option><option>Foto lato sx</option>
              <option>Foto interno</option><option>Libretto</option><option>Assicurazione</option><option>Bollo</option><option>Revisione</option><option>Manutenzione</option>
            </select>
            <input type="file" name="file" accept="image/*,.pdf" required>
            <button>Carica file mezzo</button>
          </form>
          <ul>${lista}</ul>
          <a class="btn" href="/mezzi-web">Torna elenco mezzi</a>
        </div>
      `));
    });
  });
});

app.post('/mezzo/:id', (req, res) => {
  const b = req.body;
  db.run(`
    UPDATE mezzi SET
    targa=?, marca=?, modello=?, categoria=?, posti=?, descrizione_pubblica=?, prezzo_giorno=?, km_inclusi=?,
    km_attuali=?, tagliando_km_scadenza=?, tagliando_data_scadenza=?, revisione_scadenza=?, bollo_scadenza=?,
    assicurazione_scadenza=?, gomme_scadenza=?, alert_giorni=?, alert_km=?, manutenzione_note=?
    WHERE id=?
  `, [
    b.targa,b.marca,b.modello,b.categoria,b.posti,b.descrizione_pubblica,b.prezzo_giorno,b.km_inclusi,
    b.km_attuali,b.tagliando_km_scadenza,b.tagliando_data_scadenza,b.revisione_scadenza,b.bollo_scadenza,
    b.assicurazione_scadenza,b.gomme_scadenza,b.alert_giorni,b.alert_km,b.manutenzione_note,req.params.id
  ], () => res.redirect(`/mezzo/${req.params.id}`));
});

app.post('/mezzo/:id/foto', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');
  db.run(`INSERT INTO allegati (mezzo_id,tipo,filename,originalname,path,mimetype) VALUES (?,?,?,?,?,?)`,
    [req.params.id, req.body.tipo, req.file.filename, req.file.originalname, req.file.path, req.file.mimetype],
    () => res.redirect(`/mezzo/${req.params.id}`));
});

app.get('/scadenze-mezzi', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY targa`, [], (err, rows) => {
    const trs = rows.map(m => `
      <tr>
        <td><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a></td>
        <td>${esc(descrizionePubblica(m))}</td>
        <td>${esc(m.km_attuali || m.km)}</td>
        <td>${esc(m.tagliando_km_scadenza)}</td>
        <td>${esc(m.tagliando_data_scadenza)}</td>
        <td>${esc(m.revisione_scadenza)}</td>
        <td>${esc(m.bollo_scadenza)}</td>
        <td>${esc(m.assicurazione_scadenza)}</td>
        <td>${alertBadge(m)}</td>
      </tr>`).join('');
    const alerts = rows.map(m => alertMezzo(m) ? `<div><b>${esc(m.targa)} ${esc(m.modello)}</b>${alertMezzo(m)}</div>` : '').join('');
    res.send(page('Scadenze mezzi', `
      <h2>Scadenze e manutenzioni mezzi</h2>
      <div class="box">${alerts || '<p class="ok">Nessun alert attivo.</p>'}</div>
      <table>
        <tr><th>Targa</th><th>Descrizione</th><th>Km</th><th>Tagliando km</th><th>Tagliando data</th><th>Revisione</th><th>Bollo</th><th>Assicurazione</th><th>Alert</th></tr>
        ${trs}
      </table>
    `));
  });
});

// Nuova prenotazione
app.get('/nuova-prenotazione', (req, res) => {
  const selectedMezzo = normalize(req.query.mezzo_id);
  const selectedData = normalize(req.query.data);
  db.all(`SELECT * FROM mezzi ORDER BY categoria,targa`, [], (err, mezzi) => {
    const opt = mezzi.map(m => `<option value="${m.id}" ${String(m.id)===selectedMezzo?'selected':''}>${esc(m.targa)} - ${esc(descrizionePubblica(m))}</option>`).join('');
    res.send(page('Nuova prenotazione', `
      <h2>Nuova prenotazione / contratto</h2>
      ${selectedData ? `<p class="notice">Aperta dal planning per il giorno <b>${esc(selectedData)}</b>.</p>` : ''}
      <form method="POST" action="/prenota-admin">
        <div class="grid">
          <div><label>Nome</label><input name="nome" required></div>
          <div><label>Cognome</label><input name="cognome" required></div>
          <div><label>Telefono</label><input name="telefono" required></div>
          <div><label>Email</label><input name="email"></div>
          <div><label>Codice fiscale</label><input name="codice_fiscale"></div>
          <div><label>Indirizzo</label><input name="indirizzo"></div>
          <div><label>Citta</label><input name="citta"></div>
          <div><label>CAP</label><input name="cap"></div>
          <div><label>Tipo cliente</label><select name="tipo_cliente"><option>privato</option><option>azienda</option></select></div>
          <div><label>P.IVA</label><input name="piva"></div>
          <div><label>Ragione sociale</label><input name="ragione_sociale"></div>
          <div><label>Mezzo</label><select name="mezzo_id">${opt}</select></div>
          <div><label>Data inizio</label><input type="date" name="data_inizio" value="${esc(selectedData)}" required></div>
          <div><label>Ora inizio</label><input type="time" name="ora_inizio" value="08:30"></div>
          <div><label>Data fine</label><input type="date" name="data_fine" value="${esc(selectedData)}" required></div>
          <div><label>Ora fine</label><input type="time" name="ora_fine" value="18:00"></div>
          <div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
          <div><label>Carburante uscita</label><select name="carburante_uscita">${fuelOptions('4/4 pieno')}</select></div>
        </div>
        <label>Note</label><textarea name="note"></textarea>
        <button>Crea contratto</button>
      </form>
    `));
  });
});

app.post('/prenota-admin', (req, res) => {
  const b = req.body;
  db.get(`SELECT * FROM mezzi WHERE id=?`, [b.mezzo_id], (err, mezzo) => {
    if (!mezzo) return res.send('Mezzo non trovato');
    queryDisponibilita(b.mezzo_id, b.data_inizio, b.data_fine, (e2, occ) => {
      if (occ) return res.send(page('Occupato', `<h2 class="bad">Mezzo occupato in queste date</h2><a class="btn" href="/planning">Vai al planning</a>`));
      const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.ora_inizio, b.ora_fine, b.km_previsti);
      db.run(`
        INSERT INTO prenotazioni
        (codice,nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,tipo_cliente,piva,ragione_sociale,mezzo_id,data_inizio,data_fine,ora_inizio,ora_fine,giorni,km_previsti,extra_fuori_orario,imponibile,iva,totale,cauzione,carburante_uscita,stato,note)
        VALUES ('TEMP',?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
      `, [
        b.nome,b.cognome,b.telefono,b.email,b.codice_fiscale,b.indirizzo,b.citta,b.cap,b.tipo_cliente,b.piva,b.ragione_sociale,
        b.mezzo_id,b.data_inizio,b.data_fine,b.ora_inizio,b.ora_fine,calc.giorni,Number(b.km_previsti || 0),calc.extra_fuori_orario,
        calc.imponibile,calc.iva,calc.totale,CAUZIONE,b.carburante_uscita || '4/4 pieno','bozza',b.note
      ], function(err3) {
        if (err3) return res.send(String(err3));
        const cod = codicePratica(this.lastID);
        db.run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, this.lastID]);
        res.send(actionScreen(this.lastID, 'Contratto creato', `Codice: <b>${cod}</b><br>Totale: <b>euro ${calc.totale.toFixed(2)}</b>`));
      });
    });
  });
});

// Pagina cliente senza targa
app.get('/prenota', (req, res) => {
  res.send(page('Prenota DP RENT', `
    <h2>Richiesta prenotazione cliente</h2>
    <p class="notice">Il cliente sceglie solo la categoria. La targa non viene mostrata lato cliente.</p>
    <form method="POST" action="/prenota-cliente">
      <div class="grid">
        <div><label>Tipo mezzo</label><select name="categoria" required>
          <option value="FURGONE">Furgone cargo/merci</option>
          <option value="9_POSTI">Pulmino 9 posti</option>
          <option value="AUTO_DACIA">Auto economica</option>
          <option value="AUTO_GOLF">Auto categoria Golf</option>
          <option value="ESCAVATORE">Escavatore</option>
          <option value="SEMOVENTE">Piattaforma / semovente</option>
        </select></div>
        <div><label>Nome</label><input name="nome" required></div>
        <div><label>Cognome</label><input name="cognome" required></div>
        <div><label>Telefono</label><input name="telefono" required></div>
        <div><label>Email</label><input name="email"></div>
        <div><label>Codice fiscale</label><input name="codice_fiscale"></div>
        <div><label>Indirizzo</label><input name="indirizzo"></div>
        <div><label>Citta</label><input name="citta"></div>
        <div><label>CAP</label><input name="cap"></div>
        <div><label>Data inizio</label><input type="date" name="data_inizio" required></div>
        <div><label>Ora inizio</label><input type="time" name="ora_inizio" value="08:30"></div>
        <div><label>Data fine</label><input type="date" name="data_fine" required></div>
        <div><label>Ora fine</label><input type="time" name="ora_fine" value="18:00"></div>
        <div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
      </div>
      <button>Invia richiesta</button>
    </form>
  `));
});
app.post('/prenota-cliente', (req, res) => {
  const b = req.body;
  db.all(`SELECT * FROM mezzi WHERE categoria=? ORDER BY id ASC`, [b.categoria], (err, mezzi) => {
    if (!mezzi || mezzi.length === 0) return res.send(page('Nessun mezzo', '<h2>Nessun mezzo disponibile per questa categoria</h2>'));
    function prova(index) {
      if (index >= mezzi.length) return res.send(page('Non disponibile', '<h2>Nessun mezzo libero nelle date richieste</h2>'));
      const mezzo = mezzi[index];
      queryDisponibilita(mezzo.id, b.data_inizio, b.data_fine, (e2, occ) => {
        if (occ) return prova(index + 1);
        const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.ora_inizio, b.ora_fine, b.km_previsti);
        db.run(`
          INSERT INTO prenotazioni
          (codice,nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,tipo_cliente,mezzo_id,data_inizio,data_fine,ora_inizio,ora_fine,giorni,km_previsti,extra_fuori_orario,imponibile,iva,totale,cauzione,stato)
          VALUES ('TEMP',?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        `, [
          b.nome,b.cognome,b.telefono,b.email,b.codice_fiscale,b.indirizzo,b.citta,b.cap,'privato',
          mezzo.id,b.data_inizio,b.data_fine,b.ora_inizio || '08:30',b.ora_fine || '18:00',calc.giorni,Number(b.km_previsti || 0),calc.extra_fuori_orario,
          calc.imponibile,calc.iva,calc.totale,CAUZIONE,'richiesta_cliente'
        ], function(err3) {
          if (err3) return res.send(String(err3));
          const cod = codicePratica(this.lastID);
          db.run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, this.lastID]);
          res.send(page('Richiesta inviata', `
            <h2 class="ok">Richiesta inviata</h2>
            <p>Codice: <b>${cod}</b></p>
            <p>Totale previsto: <b>euro ${calc.totale.toFixed(2)}</b></p>
            <p>DP RENT confermera la prenotazione.</p>
          `));
        });
      });
    }
    prova(0);
  });
});

// Dettaglio, storico, planning
app.get('/prenotazione/:id', (req, res) => {
  db.get(`
    SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.descrizione_pubblica
    FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE p.id=?
  `, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');
    res.send(page('Dettaglio contratto', `
      <div class="box">
        <h2>Contratto ${esc(p.codice)}</h2>
        <p><b>Cliente:</b> ${esc(p.nome)} ${esc(p.cognome)} - ${esc(p.telefono)}</p>
        <p><b>Mezzo:</b> <a href="/mezzo/${p.mezzo_id}">${esc(p.targa)} ${esc(descrizionePubblica(p))}</a></p>
        <p><b>Date:</b> ${esc(p.data_inizio)} ore ${esc(p.ora_inizio)} → ${esc(p.data_fine)} ore ${esc(p.ora_fine)}</p>
        <p><b>Totale:</b> euro ${Number(p.totale || 0).toFixed(2)}</p>
        <p><b>Stato:</b> ${esc(p.stato)}</p>
        <p><b>Carburante:</b> uscita ${esc(p.carburante_uscita)} / rientro ${esc(p.carburante_rientro)}</p>
        <p><b>Km:</b> uscita ${esc(p.km_uscita)} / rientro ${esc(p.km_rientro)}</p>
        <p><b>Note:</b> ${esc(p.note)}</p>
        <div class="actions">
          <a class="btn" href="/contratto/${p.id}">Scarica / stampa PDF</a>
          <a class="btn btn2" href="/firma/${p.id}">Firma su tablet</a>
          <a class="btn btn2" href="/email/${p.id}">Invia via email</a>
          <a class="btn btn3" href="/documenti/${p.id}">Documenti cliente / foto in-out</a>
          <a class="btn btn3" href="/checkout/${p.id}">Check-out</a>
          <a class="btn btn3" href="/checkin/${p.id}">Check-in</a>
          <a class="btn btn2" href="/prenotazioni">Storico</a>
        </div>
      </div>
    `));
  });
});

app.get('/prenotazioni', (req, res) => {
  const q = normalize(req.query.q), stato = normalize(req.query.stato), dal = normalize(req.query.dal), al = normalize(req.query.al);
  let sql = `SELECT p.*, m.targa, m.marca, m.modello, m.descrizione_pubblica FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE 1=1`;
  const params = [];
  if (q) { sql += ` AND (p.codice LIKE ? OR p.nome LIKE ? OR p.cognome LIKE ? OR p.telefono LIKE ? OR m.targa LIKE ?)`; params.push(`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`); }
  if (stato) { sql += ` AND p.stato=?`; params.push(stato); }
  if (dal) { sql += ` AND date(p.data_inizio)>=date(?)`; params.push(dal); }
  if (al) { sql += ` AND date(p.data_fine)<=date(?)`; params.push(al); }
  sql += ` ORDER BY p.id DESC`;
  db.all(sql, params, (err, rows) => {
    const trs = rows.map(p => `
      <tr>
        <td><a href="/prenotazione/${p.id}">${esc(p.codice)}</a></td>
        <td>${esc(p.nome)} ${esc(p.cognome)}</td>
        <td>${esc(p.telefono)}<br>${esc(p.email)}</td>
        <td><b>${esc(p.targa)}</b><br>${esc(descrizionePubblica(p))}</td>
        <td>${esc(p.data_inizio)} → ${esc(p.data_fine)}</td>
        <td>euro ${Number(p.totale || 0).toFixed(2)}</td>
        <td>${esc(p.stato)}</td>
        <td><a href="/prenotazione/${p.id}">Apri</a><br><a href="/contratto/${p.id}">PDF</a><br><a href="/firma/${p.id}">Firma</a><br><a href="/email/${p.id}">Email</a><br><a href="/stato/${p.id}/confermato">Conferma</a></td>
      </tr>`).join('');
    res.send(page('Storico', `
      <h2>Storico contratti / prenotazioni</h2>
      <form method="GET" action="/prenotazioni" class="box">
        <div class="grid">
          <input name="q" placeholder="Cerca nome, targa, codice, telefono" value="${esc(q)}">
          <select name="stato"><option value="">Tutti gli stati</option>${['bozza','richiesta_cliente','confermato','firmato','in_corso','rientrato','chiuso','annullato'].map(s=>`<option ${stato===s?'selected':''}>${s}</option>`).join('')}</select>
          <input type="date" name="dal" value="${esc(dal)}"><input type="date" name="al" value="${esc(al)}">
        </div><button>Cerca</button>
      </form>
      <table><tr><th>Codice</th><th>Cliente</th><th>Contatti</th><th>Mezzo</th><th>Date</th><th>Totale</th><th>Stato</th><th>Azioni</th></tr>${trs}</table>
    `));
  });
});
app.get('/stato/:id/:stato', (req, res) => db.run(`UPDATE prenotazioni SET stato=? WHERE id=?`, [req.params.stato, req.params.id], () => res.redirect('/prenotazioni')));

app.get('/planning', (req, res) => {
  const mese = req.query.mese || moment().format('YYYY-MM');
  const start = moment(mese + '-01'), days = start.daysInMonth();
  db.all(`SELECT * FROM mezzi ORDER BY targa`, [], (e1, mezzi) => {
    db.all(`SELECT p.*, m.targa, m.marca, m.modello FROM prenotazioni p LEFT JOIN mezzi m ON m.id=p.mezzo_id WHERE p.stato!='annullato'`, [], (e2, pren) => {
      let header = '<th class="sticky-col">Mezzo</th>';
      for (let d=1; d<=days; d++) header += `<th>${d}</th>`;
      let rows = '';
      mezzi.forEach(m => {
        rows += `<tr><td class="sticky-col"><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a><br>${esc(descrizionePubblica(m))}</td>`;
        for (let d=1; d<=days; d++) {
          const day = start.clone().date(d).format('YYYY-MM-DD');
          const occ = pren.find(p => p.mezzo_id == m.id && moment(day).isSameOrAfter(moment(p.data_inizio)) && moment(day).isSameOrBefore(moment(p.data_fine)));
          if (occ) {
            const title = `Contratto: ${esc(occ.codice)} | Cliente: ${esc(occ.nome)} ${esc(occ.cognome)} | Tel: ${esc(occ.telefono)} | Dal ${esc(occ.data_inizio)} al ${esc(occ.data_fine)}`;
            rows += `<td class="occupato" title="${title}" onclick="window.location='/prenotazione/${occ.id}'">O</td>`;
          } else {
            rows += `<td class="libero" title="Libero - clicca per prenotare ${esc(m.targa)} il ${day}" onclick="window.location='/nuova-prenotazione?mezzo_id=${m.id}&data=${day}'">L</td>`;
          }
        }
        rows += '</tr>';
      });
      const prec = start.clone().subtract(1,'month').format('YYYY-MM'), succ = start.clone().add(1,'month').format('YYYY-MM');
      res.send(page('Planning', `
        <h2>Planning ${start.format('MM/YYYY')}</h2>
        <p><a href="/planning?mese=${prec}">← Mese precedente</a> | <a href="/planning?mese=${succ}">Mese successivo →</a></p>
        <p><span class="libero" style="padding:6px;">Libero: clic per prenotare</span> <span class="occupato" style="padding:6px;">Occupato: clic per aprire contratto</span></p>
        <div class="sticky-table"><table><tr>${header}</tr>${rows}</table></div>`));
    });
  });
});

// Documenti cliente/foto IN OUT
app.get('/documenti/:id', (req, res) => {
  db.all(`SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id DESC`, [req.params.id], (err, files) => {
    const lista = files.map(f => `<li>${esc(f.tipo)} - ${esc(f.originalname)}</li>`).join('');
    res.send(page('Documenti', `
      <h2>Documenti cliente / foto in-out</h2>
      <form method="POST" action="/documenti/${req.params.id}" enctype="multipart/form-data">
        <label>Tipo documento/foto</label>
        <select name="tipo">
          <option>Patente fronte</option><option>Patente retro</option><option>Documento fronte</option><option>Documento retro</option><option>Codice fiscale</option>
          <option>Foto uscita fronte</option><option>Foto uscita retro</option><option>Foto uscita lato dx</option><option>Foto uscita lato sx</option><option>Foto uscita interno</option><option>Foto danni uscita</option>
          <option>Foto rientro fronte</option><option>Foto rientro retro</option><option>Foto rientro lato dx</option><option>Foto rientro lato sx</option><option>Foto rientro interno</option><option>Foto danni rientro</option>
        </select>
        <input type="file" name="file" accept="image/*,.pdf" required>
        <button>Carica</button>
      </form><h3>Allegati caricati</h3><ul>${lista}</ul><a class="btn" href="/prenotazione/${req.params.id}">Torna contratto</a>`));
  });
});
app.post('/documenti/:id', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');
  db.run(`INSERT INTO allegati (prenotazione_id,tipo,filename,originalname,path,mimetype) VALUES (?,?,?,?,?,?)`,
    [req.params.id, req.body.tipo, req.file.filename, req.file.originalname, req.file.path, req.file.mimetype],
    () => res.redirect(`/documenti/${req.params.id}`));
});

app.get('/checkout/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');
    res.send(page('Check-out', `<h2>Check-out mezzo</h2><form method="POST" action="/checkout/${p.id}">
      <label>Carburante uscita</label><select name="carburante_uscita">${fuelOptions(p.carburante_uscita)}</select>
      <label>Km uscita</label><input type="number" name="km_uscita" value="${esc(p.km_uscita)}">
      <label>Note</label><textarea name="note">${esc(p.note)}</textarea><button>Salva check-out</button></form>
      <a class="btn btn3" href="/documenti/${p.id}">Carica foto uscita</a>`));
  });
});
app.post('/checkout/:id', (req, res) => {
  db.get(`SELECT mezzo_id FROM prenotazioni WHERE id=?`, [req.params.id], (e,p) => {
    db.run(`UPDATE prenotazioni SET carburante_uscita=?, km_uscita=?, note=?, stato='in_corso' WHERE id=?`,
      [req.body.carburante_uscita, req.body.km_uscita, req.body.note, req.params.id], () => {
        if (p && req.body.km_uscita) db.run(`UPDATE mezzi SET km_attuali=? WHERE id=?`, [req.body.km_uscita, p.mezzo_id]);
        res.send(actionScreen(req.params.id, 'Check-out salvato', 'Contratto aggiornato in stato in_corso.'));
      });
  });
});
app.get('/checkin/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');
    res.send(page('Check-in', `<h2>Check-in mezzo</h2><form method="POST" action="/checkin/${p.id}">
      <label>Carburante rientro</label><select name="carburante_rientro">${fuelOptions(p.carburante_rientro)}</select>
      <label>Km rientro</label><input type="number" name="km_rientro" value="${esc(p.km_rientro)}">
      <label>Note</label><textarea name="note">${esc(p.note)}</textarea><button>Salva check-in</button></form>
      <a class="btn btn3" href="/documenti/${p.id}">Carica foto rientro</a>`));
  });
});
app.post('/checkin/:id', (req, res) => {
  db.get(`SELECT mezzo_id FROM prenotazioni WHERE id=?`, [req.params.id], (e,p) => {
    db.run(`UPDATE prenotazioni SET carburante_rientro=?, km_rientro=?, note=?, stato='rientrato' WHERE id=?`,
      [req.body.carburante_rientro, req.body.km_rientro, req.body.note, req.params.id], () => {
        if (p && req.body.km_rientro) db.run(`UPDATE mezzi SET km_attuali=? WHERE id=?`, [req.body.km_rientro, p.mezzo_id]);
        res.send(actionScreen(req.params.id, 'Check-in salvato', 'Contratto aggiornato in stato rientrato.'));
      });
  });
});

// PDF/Firma/Email
app.get('/contratto/:id', (req, res) => generaPdfContratto(req.params.id, (err,file) => err ? res.send('Errore PDF: '+err.message) : res.download(file)));

app.get('/firma/:id', (req, res) => {
  res.send(page('Firma', `
    <h2>Firma contratto</h2><p>Firma con dito o penna sul tablet.</p><canvas id="canvas"></canvas><br>
    <button onclick="clearCanvas()">Cancella</button><button onclick="saveFirma()">Salva firma</button>
    <script>
      const canvas=document.getElementById('canvas'),ctx=canvas.getContext('2d');canvas.width=canvas.offsetWidth;canvas.height=220;let drawing=false;
      function pos(e){const r=canvas.getBoundingClientRect();const t=e.touches?e.touches[0]:e;return{x:t.clientX-r.left,y:t.clientY-r.top};}
      function start(e){drawing=true;const p=pos(e);ctx.beginPath();ctx.moveTo(p.x,p.y);e.preventDefault();}
      function move(e){if(!drawing)return;const p=pos(e);ctx.lineTo(p.x,p.y);ctx.stroke();e.preventDefault();}
      function end(e){drawing=false;e.preventDefault();}
      canvas.addEventListener('mousedown',start);canvas.addEventListener('mousemove',move);canvas.addEventListener('mouseup',end);
      canvas.addEventListener('touchstart',start);canvas.addEventListener('touchmove',move);canvas.addEventListener('touchend',end);
      function clearCanvas(){ctx.clearRect(0,0,canvas.width,canvas.height);}
      function saveFirma(){fetch('/firma/${req.params.id}',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({firma:canvas.toDataURL('image/png')})}).then(()=>location.href='/prenotazione/${req.params.id}');}
    </script>`));
});
app.post('/firma/:id', (req, res) => {
  const data = req.body.firma;
  if (!data) return res.status(400).send('Firma mancante');
  const base64 = data.split(',')[1];
  const file = path.join(firmeDir, `firma_${req.params.id}.png`);
  fs.writeFileSync(file, base64, 'base64');
  db.run(`UPDATE prenotazioni SET firma_path=?, stato='firmato' WHERE id=?`, [file, req.params.id], () => generaPdfContratto(req.params.id, () => res.send('OK')));
});

app.get('/email/:id', (req,res)=>{
  db.get(`SELECT * FROM prenotazioni WHERE id=?`,[req.params.id],(err,p)=>{
    if(!p)return res.send('Contratto non trovato');
    res.send(page('Invia email', `<h2>Invia contratto via email</h2><form method="POST" action="/email/${p.id}">
      <label>Email destinatario</label><input name="email" value="${esc(p.email)}" required>
      <label>Messaggio</label><textarea name="messaggio">Buongiorno, in allegato trova il contratto DP RENT.</textarea><button>Invia email</button></form>
      <p class="notice">Per inviare davvero configura su Render: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS, SMTP_FROM.</p>`));
  });
});
app.post('/email/:id',(req,res)=>{
  generaPdfContratto(req.params.id, async (err,file)=>{
    if(err)return res.send('Errore PDF: '+err.message);
    if(!process.env.SMTP_HOST){
      return res.send(page('SMTP mancante', `<h2 class="bad">Email non configurata</h2><p>Il PDF è pronto, ma per inviare email devi configurare SMTP su Render.</p><a class="btn" href="/contratto/${req.params.id}">Scarica PDF</a>`));
    }
    const transporter=nodemailer.createTransport({
      host:process.env.SMTP_HOST, port:Number(process.env.SMTP_PORT||587),
      secure:Number(process.env.SMTP_PORT||587)===465,
      auth:{user:process.env.SMTP_USER,pass:process.env.SMTP_PASS}
    });
    await transporter.sendMail({
      from:process.env.SMTP_FROM||AZIENDA.email, to:req.body.email, subject:'Contratto DP RENT',
      text:req.body.messaggio||'In allegato contratto DP RENT.',
      attachments:[{filename:path.basename(file),path:file}]
    });
    db.run(`UPDATE prenotazioni SET stato='inviato_email' WHERE id=?`,[req.params.id]);
    res.send(actionScreen(req.params.id,'Email inviata','Contratto inviato correttamente.'));
  });
});

app.get('/mezzi', (req,res)=> db.all(`SELECT * FROM mezzi`, [], (err,rows)=>res.json(rows||[])));

app.listen(PORT, () => console.log('DP RENT APP V8 ONLINE'));
