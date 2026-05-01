require('dns').setDefaultResultOrder('ipv4first');
require('dotenv').config();
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const PDFDocument = require('pdfkit');
const nodemailer = require('nodemailer');
const { google } = require('googleapis');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.use(express.json({ limit: '50mb' }));

const ROOT = __dirname;
const UPLOADS = path.join(ROOT, 'uploads');
const CONTRACTS = path.join(ROOT, 'contracts');
const PUBLIC = path.join(ROOT, 'public');

[UPLOADS, CONTRACTS, PUBLIC].forEach(d => {
  if (!fs.existsSync(d)) fs.mkdirSync(d);
});

app.use('/uploads', express.static(UPLOADS));
app.use('/contracts', express.static(CONTRACTS));
app.use('/public', express.static(PUBLIC));

const upload = multer({ dest: UPLOADS });

const db = new sqlite3.Database('./database.sqlite');

const AZIENDA = {
  nome: 'Trasporti DP S.R.L. - DP RENT',
  indirizzo: 'Via Tuderte 466, Narni (TR)',
  piva: '01385450554',
  telefono: '0744817108',
  email: 'contabilita@trasportidp.com'
};

const IVA = 0.22;
const EXTRA_FUORI_ORARIO = 30;
const EXTRA_KM = 0.15;
const CAUZIONE_DEFAULT = 500;

const TERMS_URL = 'https://carrentalsoftware.myappy.it/data/public/user/65996976/terms_file.pdf';
const PRIVACY_URL = 'https://carrentalsoftware.myappy.it/data/public/user/65996976/privacy_file.pdf';

function esc(v) {
  return String(v ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function todayISO() {
  return new Date().toISOString().slice(0, 10);
}

function nowCode() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}${mm}${dd}`;
}

function diffDays(a, b) {
  const d1 = new Date(a + 'T00:00:00');
  const d2 = new Date(b + 'T00:00:00');
  return Math.floor((d2 - d1) / 86400000) + 1;
}

function dateInRange(day, start, end) {
  return day >= start && day <= end;
}

function daysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function extraOrario(ora) {
  if (!ora) return 0;
  const [h, m] = ora.split(':').map(Number);
  const min = h * 60 + (m || 0);
  const start = 8 * 60 + 30;
  const end = 18 * 60 + 30;
  return min < start || min > end ? EXTRA_FUORI_ORARIO : 0;
}

function categoriaAuto(marca, modello, descrizione, codice) {
  const s = `${marca || ''} ${modello || ''} ${descrizione || ''} ${codice || ''}`.toUpperCase();
  if (s.includes('GOLF')) return 'AUTO_GOLF';
  if (s.includes('DACIA')) return 'AUTO_DACIA';
  if (s.includes('ESCAVATORE')) return 'ESCAVATORE';
  if (s.includes('PIATTAFORMA') || s.includes('SEMOVENTE')) return 'SEMOVENTE';
  if (s.includes('9P') || s.includes('9 POSTI') || s.includes('TOURNEO') || s.includes('PERSONE')) return '9_POSTI';
  return 'FURGONE';
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

function descrizionePubblica(m) {
  if (m.descrizione_pubblica) return m.descrizione_pubblica;
  const base = `${m.marca || ''} ${m.modello || ''}`.trim();
  if (m.categoria === '9_POSTI') return `${base} - pulmino 9 posti`;
  if (m.categoria === 'FURGONE') return `${base} - furgone cargo/merci`;
  if (m.categoria === 'AUTO_DACIA') return `${base} - auto economica`;
  if (m.categoria === 'AUTO_GOLF') return `${base} - auto categoria Golf`;
  if (m.categoria === 'ESCAVATORE') return `${base} - escavatore`;
  if (m.categoria === 'SEMOVENTE') return `${base} - piattaforma/semovente`;
  return base || 'Mezzo DP RENT';
}

function calcolaTotale(mezzo, dataInizio, dataFine, oraInizio, oraFine, kmPrevisti) {
  const giorni = Math.max(1, diffDays(dataInizio, dataFine));
  const prezzo = Number(mezzo.prezzo_giorno || 0);
  const kmInclusiGg = Number(mezzo.km_inclusi || 0);
  const kmInclusiTot = giorni * kmInclusiGg;
  const kmPrev = Number(kmPrevisti || 0);

  let imponibile = giorni * prezzo;

  if (kmInclusiGg > 0 && kmPrev > kmInclusiTot) {
    imponibile += (kmPrev - kmInclusiTot) * EXTRA_KM;
  }

  const extra = extraOrario(oraInizio) + extraOrario(oraFine);
  imponibile += extra;

  const iva = imponibile * IVA;
  const totale = imponibile + iva;

  return {
    giorni,
    kmInclusiTot,
    extra_fuori_orario: extra,
    imponibile,
    iva,
    totale
  };
}

function codiceContratto(id) {
  return `DPR-${nowCode()}-${String(id).padStart(4, '0')}`;
}

function fuelOptions(selected) {
  const arr = ['4/4 pieno', '3/4', '1/2', '1/4', 'Riserva', 'Vuoto'];
  return arr.map(x => `<option value="${x}" ${x === selected ? 'selected' : ''}>${x}</option>`).join('');
}

function page(title, body) {
  const logo = fs.existsSync(path.join(PUBLIC, 'logo.png'))
    ? `<img src="/public/logo.png" style="height:42px;background:white;border-radius:6px;padding:4px">`
    : `<b style="font-size:28px">DP RENT</b>`;

  return `<!DOCTYPE html>
<html lang="it">
<head>
<meta charset="UTF-8">
<title>${esc(title)}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{margin:0;font-family:Arial;background:#f4f4f4;color:#222}
header{background:#111;color:white;padding:15px;display:flex;gap:15px;align-items:center}
nav{background:#c40000;padding:12px;display:flex;gap:14px;flex-wrap:wrap}
nav a{color:white;text-decoration:none;font-weight:bold}
main{padding:18px}
.box{background:white;border-radius:10px;padding:18px;margin-bottom:18px;box-shadow:0 2px 8px #ccc}
table{width:100%;border-collapse:collapse;background:white}
th,td{border:1px solid #ddd;padding:7px;font-size:13px}
th{background:#111;color:white}
input,select,textarea,button{padding:10px;margin:4px 0;width:100%;box-sizing:border-box}
button,.btn{background:#c40000;color:white;border:0;border-radius:6px;padding:10px;text-decoration:none;display:inline-block;font-weight:bold;cursor:pointer;margin:3px}
.btn2{background:#333}.btn3{background:#0b6b2d}
.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
.ok{color:green;font-weight:bold}.bad{color:red;font-weight:bold}.warn{color:#b36b00;font-weight:bold}
.libero{background:#1fae4b;color:white;text-align:center;font-weight:bold;cursor:pointer}
.occupato{background:#d90000;color:white;text-align:center;font-weight:bold;cursor:pointer}
.libero:hover,.occupato:hover{outline:3px solid #111}
.sticky{overflow:auto;max-height:78vh}
.sticky th{position:sticky;top:0;z-index:2}
.first{position:sticky;left:0;background:white;z-index:1;min-width:170px}
th.first{background:#111;color:white;z-index:3}
.alert{background:#ffe0e0;border:1px solid #d90000;border-radius:8px;padding:8px;margin:5px 0}
.notice{background:#fff3cd;border:1px solid #e0c66a;border-radius:8px;padding:8px;margin:8px 0}
canvas{width:100%;height:230px;border:2px solid #111;background:white}
@media(max-width:750px){.grid{grid-template-columns:1fr}main{padding:8px}th,td{font-size:11px;padding:5px}}
</style>
</head>
<body>
<header>${logo}<h1>DP RENT APP</h1></header>
<nav>
<a href="/">Dashboard</a>
<a href="/mezzi">Mezzi</a>
<a href="/scadenze">Scadenze</a>
<a href="/import">Import Excel</a>
<a href="/nuova">Nuova prenotazione</a>
<a href="/storico">Storico</a>
<a href="/planning">Planning</a>
<a href="/cliente">Pagina cliente</a>
<a href="/logo">Logo</a>
<a href="/test-email">Test Email</a>
<a href="/test-drive">Test Drive</a>
</nav>
<main>${body}</main>
</body>
</html>`;
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
      marca TEXT,
      modello TEXT,
      categoria TEXT,
      descrizione TEXT,
      descrizione_pubblica TEXT,
      posti INTEGER,
      km INTEGER DEFAULT 0,
      km_attuali INTEGER DEFAULT 0,
      prezzo_giorno REAL DEFAULT 70,
      km_inclusi INTEGER DEFAULT 150,
      stato TEXT DEFAULT 'disponibile',
      tagliando_km_scadenza INTEGER,
      tagliando_data_scadenza TEXT,
      revisione_scadenza TEXT,
      bollo_scadenza TEXT,
      assicurazione_scadenza TEXT,
      gomme_scadenza TEXT,
      alert_km INTEGER DEFAULT 1000,
      alert_giorni INTEGER DEFAULT 30,
      note TEXT
    )
  `);

  [
    ['descrizione_pubblica','TEXT'],['posti','INTEGER'],['km_attuali','INTEGER DEFAULT 0'],
    ['tagliando_km_scadenza','INTEGER'],['tagliando_data_scadenza','TEXT'],
    ['revisione_scadenza','TEXT'],['bollo_scadenza','TEXT'],['assicurazione_scadenza','TEXT'],
    ['gomme_scadenza','TEXT'],['alert_km','INTEGER DEFAULT 1000'],['alert_giorni','INTEGER DEFAULT 30'],['note','TEXT']
  ].forEach(c => addColumn('mezzi', c[0], c[1]));

  db.run(`
    CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      cliente_nome TEXT,
      cliente_cognome TEXT,
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
      conducente1_nome TEXT,
      conducente1_cognome TEXT,
      conducente1_telefono TEXT,
      conducente1_email TEXT,
      conducente1_cf TEXT,
      conducente1_patente TEXT,
      conducente1_patente_scadenza TEXT,
      conducente1_documento_scadenza TEXT,
      conducente2_nome TEXT,
      conducente2_cognome TEXT,
      conducente2_telefono TEXT,
      conducente2_email TEXT,
      conducente2_cf TEXT,
      conducente2_patente TEXT,
      conducente2_patente_scadenza TEXT,
      conducente2_documento_scadenza TEXT,
      mezzo_id INTEGER,
      data_inizio TEXT,
      data_fine TEXT,
      ora_inizio TEXT,
      ora_fine TEXT,
      giorni INTEGER,
      km_previsti INTEGER,
      extra_fuori_orario REAL,
      imponibile REAL,
      iva REAL,
      totale REAL,
      cauzione REAL DEFAULT 500,
      carburante_uscita TEXT DEFAULT '4/4 pieno',
      carburante_rientro TEXT DEFAULT '4/4 pieno',
      km_uscita INTEGER,
      km_rientro INTEGER,
      stato TEXT DEFAULT 'bozza',
      firma_path TEXT,
      pdf_path TEXT,
      pdf_drive_file_id TEXT,
      pdf_drive_web_link TEXT,
      nexi_payment_url TEXT,
      nexi_stato TEXT,
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  [
    ['pec','TEXT'],['sdi','TEXT'],['firma_path','TEXT'],['pdf_path','TEXT'],
    ['pdf_drive_file_id','TEXT'],['pdf_drive_web_link','TEXT'],
    ['nexi_payment_url','TEXT'],['nexi_stato','TEXT']
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

  addColumn('allegati', 'drive_file_id', 'TEXT');
  addColumn('allegati', 'drive_web_link', 'TEXT');
});

function googleDriveConfigured() {
  return !!(process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL && process.env.GOOGLE_PRIVATE_KEY && process.env.GOOGLE_DRIVE_FOLDER_ID);
}

function getDrive() {
  const key = process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n');
  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key,
    scopes: ['https://www.googleapis.com/auth/drive.file']
  });
  return google.drive({ version: 'v3', auth });
}

async function driveUpload(localPath, filename, mimetype, folderName) {
  if (!googleDriveConfigured()) return null;
  if (!fs.existsSync(localPath)) return null;

  const drive = getDrive();
  const parent = process.env.GOOGLE_DRIVE_FOLDER_ID;
  let folderId = parent;

  if (folderName) {
    const safe = String(folderName).replace(/[\/\\:*?"<>|]/g, '-');
    const q = `'${parent}' in parents and name='${safe}' and mimeType='application/vnd.google-apps.folder' and trashed=false`;
    const found = await drive.files.list({ q, fields: 'files(id,name)', pageSize: 1 });

    if (found.data.files && found.data.files.length) {
      folderId = found.data.files[0].id;
    } else {
      const created = await drive.files.create({
        requestBody: {
          name: safe,
          mimeType: 'application/vnd.google-apps.folder',
          parents: [parent]
        },
        fields: 'id'
      });
      folderId = created.data.id;
    }
  }

  const uploaded = await drive.files.create({
    requestBody: { name: filename, parents: [folderId] },
    media: {
      mimeType: mimetype || 'application/octet-stream',
      body: fs.createReadStream(localPath)
    },
    fields: 'id,webViewLink'
  });

  return {
    id: uploaded.data.id,
    link: uploaded.data.webViewLink
  };
}

async function sendEmail(to, subject, text, attachments = []) {
  if (!process.env.SMTP_HOST) throw new Error('SMTP non configurato');

  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: String(process.env.SMTP_SECURE || '').toLowerCase() === 'true',
connectionTimeout: 30000,
greetingTimeout: 30000,
socketTimeout: 30000,    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });

  return transporter.sendMail({
    from: process.env.SMTP_FROM || AZIENDA.email,
    to,
    subject,
    text,
    attachments
  });
}

async function sendAlert(subject, text) {
  try {
    const to = process.env.ALERT_EMAIL || process.env.SMTP_USER;
    if (to) await sendEmail(to, subject, text);
  } catch (e) {
    console.log('Errore alert email:', e.message);
  }
}

function alertMezzo(m) {
  const list = [];
  const now = new Date(todayISO() + 'T00:00:00');
  const giorniAlert = Number(m.alert_giorni || 30);
  const kmAlert = Number(m.alert_km || 1000);
  const kmAtt = Number(m.km_attuali || m.km || 0);

  function checkDate(label, date) {
    if (!date) return;
    const d = new Date(date + 'T00:00:00');
    const diff = Math.ceil((d - now) / 86400000);
    if (diff < 0) list.push(`❌ ${label} scaduta il ${date}`);
    else if (diff <= giorniAlert) list.push(`⚠️ ${label} in scadenza il ${date} (${diff} giorni)`);
  }

  checkDate('Tagliando data', m.tagliando_data_scadenza);
  checkDate('Revisione', m.revisione_scadenza);
  checkDate('Bollo', m.bollo_scadenza);
  checkDate('Assicurazione', m.assicurazione_scadenza);
  checkDate('Gomme/manutenzione', m.gomme_scadenza);

  if (m.tagliando_km_scadenza) {
    const diffKm = Number(m.tagliando_km_scadenza) - kmAtt;
    if (diffKm <= 0) list.push(`❌ Tagliando km scaduto: superato di ${Math.abs(diffKm)} km`);
    else if (diffKm <= kmAlert) list.push(`⚠️ Tagliando vicino: mancano ${diffKm} km`);
  }

  return list;
}

function renderActionScreen(id, title, msg) {
  return page(title, `
    <div class="box">
      <h2 class="ok">${esc(title)}</h2>
      <p>${msg || ''}</p>
      <a class="btn" href="/contratto/${id}">Scarica/stampa PDF</a>
      <a class="btn btn2" href="/firma/${id}">Firma tablet</a>
      <a class="btn btn2" href="/email/${id}">Invia email</a>
      <a class="btn btn3" href="/documenti/${id}">Documenti/foto</a>
      <a class="btn btn3" href="/checkout/${id}">Check-out</a>
      <a class="btn btn3" href="/checkin/${id}">Check-in</a>
      <a class="btn btn2" href="/nexi/${id}">Pagamento Nexi</a>
      <a class="btn btn2" href="/prenotazione/${id}">Dettaglio</a>
      <a class="btn btn2" href="/storico">Storico</a>
    </div>
  `);
}

app.get('/', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY targa`, [], (err, mezzi) => {
    db.get(`SELECT COUNT(*) AS n FROM prenotazioni`, [], (e2, cnt) => {
      const alerts = (mezzi || []).flatMap(m => alertMezzo(m).map(a => `<div class="alert"><b>${esc(m.targa)} ${esc(m.modello)}</b><br>${esc(a)}</div>`)).join('');

      res.send(page('Dashboard', `
        <div class="box">
          <h2>Dashboard DP RENT</h2>
          <p>Mezzi caricati: <b>${mezzi.length}</b></p>
          <p>Contratti/prenotazioni: <b>${cnt ? cnt.n : 0}</b></p>
          <p class="notice">Google Drive: <b>${googleDriveConfigured() ? 'configurato' : 'non configurato'}</b></p>
        </div>
        <div class="box">
          <h2>Alert</h2>
          ${alerts || '<p class="ok">Nessun alert attivo.</p>'}
        </div>
      `));
    });
  });
});

app.get('/logo', (req, res) => {
  const has = fs.existsSync(path.join(PUBLIC, 'logo.png'));
  res.send(page('Logo', `
    <div class="box">
      <h2>Logo</h2>
      ${has ? '<img src="/public/logo.png" style="max-width:220px;background:#eee;padding:10px">' : '<p>Nessun logo caricato.</p>'}
      <form method="POST" action="/logo" enctype="multipart/form-data">
        <input type="file" name="logo" accept="image/png,image/jpeg" required>
        <button>Salva logo</button>
      </form>
    </div>
  `));
});

app.post('/logo', multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, PUBLIC),
    filename: (req, file, cb) => cb(null, 'logo.png')
  })
}).single('logo'), (req, res) => res.redirect('/logo'));

app.get('/import', (req, res) => {
  res.send(page('Import Excel', `
    <div class="box">
      <h2>Import mezzi Excel</h2>
      <form method="POST" action="/import" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx,.xls,.csv" required>
        <button>Importa</button>
      </form>
    </div>
  `));
});

app.post('/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');

  const wb = XLSX.readFile(req.file.path);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);
  let count = 0;

  const stmt = db.prepare(`
    INSERT OR REPLACE INTO mezzi
    (uid,targa,marca,modello,categoria,descrizione,descrizione_pubblica,posti,km,km_attuali,prezzo_giorno,km_inclusi,stato)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,COALESCE((SELECT stato FROM mezzi WHERE targa=?),'disponibile'))
  `);

  rows.forEach(r => {
    const targa = String(r.Targa || '').trim();
    if (!targa) return;

    const marca = String(r.Marca || '').trim();
    const modello = String(r.Modello || '').trim();
    const descrizione = String(r.Descrizione || r.Descrizion || r['Immagini consegna'] || '').trim();
    const codice = String(r['Codice Tip'] || r['Codice Tipo'] || '').trim();
    const km = Number(r['Km percor'] || r['Km percorsi'] || r.Km || 0);
    const cat = categoriaAuto(marca, modello, descrizione, codice);
    const mezzoTemp = { marca, modello, categoria: cat };
    const descPub = descrizionePubblica(mezzoTemp);
    const posti = cat === '9_POSTI' ? 9 : null;

    stmt.run([
      String(r.UID || ''),
      targa,
      marca,
      modello,
      cat,
      descrizione,
      descPub,
      posti,
      km,
      km,
      prezzoCategoria(cat),
      kmCategoria(cat),
      targa
    ]);
    count++;
  });

  stmt.finalize();
  fs.unlinkSync(req.file.path);

  res.send(page('Import completato', `
    <div class="box">
      <h2 class="ok">Import completato</h2>
      <p>Mezzi importati/aggiornati: <b>${count}</b></p>
      <a class="btn" href="/mezzi">Vai ai mezzi</a>
    </div>
  `));
});

app.get('/mezzi', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY categoria,targa`, [], (err, rows) => {
    const html = rows.map(m => `
      <tr>
        <td><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a></td>
        <td>${esc(m.marca)}</td>
        <td>${esc(m.modello)}</td>
        <td>${esc(m.categoria)}</td>
        <td>${esc(descrizionePubblica(m))}</td>
        <td>${esc(m.km_attuali || m.km || 0)}</td>
        <td>€ ${Number(m.prezzo_giorno || 0).toFixed(2)}</td>
        <td>${alertMezzo(m).length ? '<span class="bad">ALERT</span>' : '<span class="ok">OK</span>'}</td>
      </tr>
    `).join('');

    res.send(page('Mezzi', `
      <h2>Mezzi</h2>
      <table>
        <tr><th>Targa</th><th>Marca</th><th>Modello</th><th>Categoria</th><th>Descrizione cliente</th><th>Km</th><th>Prezzo</th><th>Alert</th></tr>
        ${html}
      </table>
    `));
  });
});

app.get('/mezzo/:id', (req, res) => {
  db.get(`SELECT * FROM mezzi WHERE id=?`, [req.params.id], (err, m) => {
    if (!m) return res.send('Mezzo non trovato');

    db.all(`SELECT * FROM allegati WHERE mezzo_id=? ORDER BY id DESC`, [m.id], (e2, files) => {
      const alerts = alertMezzo(m).map(a => `<div class="alert">${esc(a)}</div>`).join('');
      const lista = files.map(f => `<li>${esc(f.tipo)} - ${esc(f.originalname)} ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Drive</a>` : ''}</li>`).join('');

      res.send(page('Scheda mezzo', `
        <div class="box">
          <h2>Scheda mezzo ${esc(m.targa)}</h2>
          ${alerts || '<p class="ok">Nessun alert.</p>'}
          <form method="POST" action="/mezzo/${m.id}">
            <div class="grid">
              <div><label>Targa</label><input name="targa" value="${esc(m.targa)}"></div>
              <div><label>Marca</label><input name="marca" value="${esc(m.marca)}"></div>
              <div><label>Modello</label><input name="modello" value="${esc(m.modello)}"></div>
              <div><label>Categoria</label><select name="categoria">
                ${['FURGONE','9_POSTI','AUTO_DACIA','AUTO_GOLF','ESCAVATORE','SEMOVENTE'].map(c => `<option ${m.categoria===c?'selected':''}>${c}</option>`).join('')}
              </select></div>
              <div><label>Descrizione cliente</label><input name="descrizione_pubblica" value="${esc(descrizionePubblica(m))}"></div>
              <div><label>Posti</label><input type="number" name="posti" value="${esc(m.posti)}"></div>
              <div><label>Km attuali</label><input type="number" name="km_attuali" value="${esc(m.km_attuali || m.km || 0)}"></div>
              <div><label>Prezzo giorno</label><input type="number" step="0.01" name="prezzo_giorno" value="${esc(m.prezzo_giorno)}"></div>
              <div><label>Km inclusi/giorno</label><input type="number" name="km_inclusi" value="${esc(m.km_inclusi)}"></div>
              <div><label>Tagliando km scadenza</label><input type="number" name="tagliando_km_scadenza" value="${esc(m.tagliando_km_scadenza)}"></div>
              <div><label>Tagliando data</label><input type="date" name="tagliando_data_scadenza" value="${esc(m.tagliando_data_scadenza)}"></div>
              <div><label>Revisione</label><input type="date" name="revisione_scadenza" value="${esc(m.revisione_scadenza)}"></div>
              <div><label>Bollo</label><input type="date" name="bollo_scadenza" value="${esc(m.bollo_scadenza)}"></div>
              <div><label>Assicurazione</label><input type="date" name="assicurazione_scadenza" value="${esc(m.assicurazione_scadenza)}"></div>
              <div><label>Gomme/manutenzione</label><input type="date" name="gomme_scadenza" value="${esc(m.gomme_scadenza)}"></div>
              <div><label>Alert km prima</label><input type="number" name="alert_km" value="${esc(m.alert_km || 1000)}"></div>
              <div><label>Alert giorni prima</label><input type="number" name="alert_giorni" value="${esc(m.alert_giorni || 30)}"></div>
            </div>
            <label>Note</label><textarea name="note">${esc(m.note)}</textarea>
            <button>Salva mezzo</button>
          </form>
        </div>

        <div class="box">
          <h3>Foto/documenti mezzo</h3>
          <form method="POST" action="/mezzo/${m.id}/foto" enctype="multipart/form-data">
            <select name="tipo">
              <option>Foto mezzo fronte</option>
              <option>Foto mezzo retro</option>
              <option>Foto lato dx</option>
              <option>Foto lato sx</option>
              <option>Foto interno</option>
              <option>Libretto</option>
              <option>Assicurazione</option>
              <option>Revisione</option>
              <option>Bollo</option>
              <option>Manutenzione</option>
            </select>
            <input type="file" name="file" accept="image/*,.pdf" capture="environment" required>
            <button>Carica file mezzo</button>
          </form>
          <ul>${lista}</ul>
        </div>
      `));
    });
  });
});

app.post('/mezzo/:id', (req, res) => {
  const b = req.body;
  db.run(`
    UPDATE mezzi SET
    targa=?, marca=?, modello=?, categoria=?, descrizione_pubblica=?, posti=?, km_attuali=?, prezzo_giorno=?, km_inclusi=?,
    tagliando_km_scadenza=?, tagliando_data_scadenza=?, revisione_scadenza=?, bollo_scadenza=?, assicurazione_scadenza=?,
    gomme_scadenza=?, alert_km=?, alert_giorni=?, note=?
    WHERE id=?
  `, [
    b.targa,b.marca,b.modello,b.categoria,b.descrizione_pubblica,b.posti,b.km_attuali,b.prezzo_giorno,b.km_inclusi,
    b.tagliando_km_scadenza,b.tagliando_data_scadenza,b.revisione_scadenza,b.bollo_scadenza,b.assicurazione_scadenza,
    b.gomme_scadenza,b.alert_km,b.alert_giorni,b.note,req.params.id
  ], () => res.redirect(`/mezzo/${req.params.id}`));
});

app.post('/mezzo/:id/foto', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');

  db.get(`SELECT * FROM mezzi WHERE id=?`, [req.params.id], async (err, m) => {
    let drive = null;

    try {
      drive = await driveUpload(
        req.file.path,
        `${Date.now()}_${req.body.tipo}_${req.file.originalname}`,
        req.file.mimetype,
        `MEZZO ${m.targa} ${m.modello || ''}`
      );
    } catch (e) {
      console.log('Drive mezzo errore:', e.message);
    }

    db.run(`
      INSERT INTO allegati (mezzo_id,tipo,filename,originalname,path,mimetype,drive_file_id,drive_web_link)
      VALUES (?,?,?,?,?,?,?,?)
    `, [
      req.params.id, req.body.tipo, req.file.filename, req.file.originalname, req.file.path, req.file.mimetype,
      drive ? drive.id : null, drive ? drive.link : null
    ], () => res.redirect(`/mezzo/${req.params.id}`));
  });
});

app.get('/scadenze', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY targa`, [], (err, mezzi) => {
    const rows = mezzi.map(m => `
      <tr>
        <td><a href="/mezzo/${m.id}">${esc(m.targa)}</a></td>
        <td>${esc(descrizionePubblica(m))}</td>
        <td>${esc(m.km_attuali || m.km || 0)}</td>
        <td>${esc(m.tagliando_km_scadenza)}</td>
        <td>${esc(m.revisione_scadenza)}</td>
        <td>${esc(m.bollo_scadenza)}</td>
        <td>${esc(m.assicurazione_scadenza)}</td>
        <td>${alertMezzo(m).join('<br>') || '<span class="ok">OK</span>'}</td>
      </tr>
    `).join('');

    res.send(page('Scadenze', `
      <h2>Scadenze mezzi</h2>
      <table>
        <tr><th>Targa</th><th>Descrizione</th><th>Km</th><th>Tagliando km</th><th>Revisione</th><th>Bollo</th><th>Assicurazione</th><th>Alert</th></tr>
        ${rows}
      </table>
    `));
  });
});

function formPrenotazione(action, selectedMezzo, selectedData) {
  return new Promise(resolve => {
    db.all(`SELECT * FROM mezzi ORDER BY categoria,targa`, [], (err, mezzi) => {
      const options = mezzi.map(m => `<option value="${m.id}" ${String(m.id)===String(selectedMezzo)?'selected':''}>${esc(m.targa)} - ${esc(descrizionePubblica(m))}</option>`).join('');

      resolve(`
        <form method="POST" action="${action}">
          <div class="grid">
            <div><label>Nome cliente</label><input name="cliente_nome" required></div>
            <div><label>Cognome cliente</label><input name="cliente_cognome"></div>
            <div><label>Telefono</label><input name="telefono" required></div>
            <div><label>Email</label><input name="email"></div>
            <div><label>Codice fiscale</label><input name="codice_fiscale"></div>
            <div><label>Indirizzo</label><input name="indirizzo"></div>
            <div><label>Città</label><input name="citta"></div>
            <div><label>CAP</label><input name="cap"></div>
            <div><label>Tipo cliente</label><select name="tipo_cliente"><option>privato</option><option>azienda</option></select></div>
            <div><label>P.IVA</label><input name="piva"></div>
            <div><label>Ragione sociale</label><input name="ragione_sociale"></div>
            <div><label>PEC</label><input name="pec"></div>
            <div><label>SDI</label><input name="sdi"></div>

            <div style="grid-column:1/-1"><h3>Conducente 1</h3></div>
            <div><label>Nome</label><input name="conducente1_nome"></div>
            <div><label>Cognome</label><input name="conducente1_cognome"></div>
            <div><label>Telefono</label><input name="conducente1_telefono"></div>
            <div><label>Email</label><input name="conducente1_email"></div>
            <div><label>Codice fiscale</label><input name="conducente1_cf"></div>
            <div><label>Numero patente</label><input name="conducente1_patente"></div>
            <div><label>Scadenza patente</label><input type="date" name="conducente1_patente_scadenza"></div>
            <div><label>Scadenza documento</label><input type="date" name="conducente1_documento_scadenza"></div>

            <div style="grid-column:1/-1"><h3>Conducente 2 opzionale</h3></div>
            <div><label>Nome</label><input name="conducente2_nome"></div>
            <div><label>Cognome</label><input name="conducente2_cognome"></div>
            <div><label>Telefono</label><input name="conducente2_telefono"></div>
            <div><label>Email</label><input name="conducente2_email"></div>
            <div><label>Codice fiscale</label><input name="conducente2_cf"></div>
            <div><label>Numero patente</label><input name="conducente2_patente"></div>
            <div><label>Scadenza patente</label><input type="date" name="conducente2_patente_scadenza"></div>
            <div><label>Scadenza documento</label><input type="date" name="conducente2_documento_scadenza"></div>

            <div style="grid-column:1/-1"><h3>Noleggio</h3></div>
            <div><label>Mezzo</label><select name="mezzo_id">${options}</select></div>
            <div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
            <div><label>Data inizio</label><input type="date" name="data_inizio" value="${esc(selectedData || '')}" required></div>
            <div><label>Ora inizio</label><input type="time" name="ora_inizio" value="08:30"></div>
            <div><label>Data fine</label><input type="date" name="data_fine" value="${esc(selectedData || '')}" required></div>
            <div><label>Ora fine</label><input type="time" name="ora_fine" value="18:00"></div>
            <div><label>Carburante uscita</label><select name="carburante_uscita">${fuelOptions('4/4 pieno')}</select></div>
          </div>
          <label>Note</label><textarea name="note"></textarea>
          <button>Crea contratto</button>
        </form>
      `);
    });
  });
}

async function creaPrenotazione(req, res, statoDefault) {
  const b = req.body;

  db.get(`SELECT * FROM mezzi WHERE id=?`, [b.mezzo_id], (err, mezzo) => {
    if (!mezzo) return res.send('Mezzo non trovato');

    db.get(`
      SELECT * FROM prenotazioni
      WHERE mezzo_id=?
      AND stato!='annullato'
      AND date(data_inizio)<=date(?)
      AND date(data_fine)>=date(?)
    `, [b.mezzo_id, b.data_fine, b.data_inizio], (e2, occupato) => {
      if (occupato) {
        return res.send(page('Occupato', `<div class="box"><h2 class="bad">Mezzo occupato in queste date</h2><a class="btn" href="/planning">Planning</a></div>`));
      }

      const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.ora_inizio, b.ora_fine, b.km_previsti);

      db.run(`
        INSERT INTO prenotazioni (
          codice, cliente_nome, cliente_cognome, telefono, email, codice_fiscale, indirizzo, citta, cap,
          tipo_cliente, piva, ragione_sociale, pec, sdi,
          conducente1_nome, conducente1_cognome, conducente1_telefono, conducente1_email, conducente1_cf, conducente1_patente, conducente1_patente_scadenza, conducente1_documento_scadenza,
          conducente2_nome, conducente2_cognome, conducente2_telefono, conducente2_email, conducente2_cf, conducente2_patente, conducente2_patente_scadenza, conducente2_documento_scadenza,
          mezzo_id, data_inizio, data_fine, ora_inizio, ora_fine, giorni, km_previsti, extra_fuori_orario,
          imponibile, iva, totale, cauzione, carburante_uscita, stato, note
        ) VALUES (
          'TEMP',?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?
        )
      `, [
        b.cliente_nome,b.cliente_cognome,b.telefono,b.email,b.codice_fiscale,b.indirizzo,b.citta,b.cap,
        b.tipo_cliente,b.piva,b.ragione_sociale,b.pec,b.sdi,
        b.conducente1_nome,b.conducente1_cognome,b.conducente1_telefono,b.conducente1_email,b.conducente1_cf,b.conducente1_patente,b.conducente1_patente_scadenza,b.conducente1_documento_scadenza,
        b.conducente2_nome,b.conducente2_cognome,b.conducente2_telefono,b.conducente2_email,b.conducente2_cf,b.conducente2_patente,b.conducente2_patente_scadenza,b.conducente2_documento_scadenza,
        b.mezzo_id,b.data_inizio,b.data_fine,b.ora_inizio,b.ora_fine,calc.giorni,Number(b.km_previsti || 0),calc.extra_fuori_orario,
        calc.imponibile,calc.iva,calc.totale,CAUZIONE_DEFAULT,b.carburante_uscita || '4/4 pieno',statoDefault,b.note
      ], function(insertErr) {
        if (insertErr) return res.send(`<pre>${esc(insertErr.message)}</pre>`);
        const codice = codiceContratto(this.lastID);
        db.run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [codice, this.lastID]);
        res.send(renderActionScreen(this.lastID, 'Contratto creato', `Codice: <b>${codice}</b><br>Totale IVA inclusa: <b>€ ${calc.totale.toFixed(2)}</b>`));
      });
    });
  });
}

app.get('/nuova', async (req, res) => {
  const html = await formPrenotazione('/nuova', req.query.mezzo_id, req.query.data);
  res.send(page('Nuova prenotazione', `<div class="box"><h2>Nuova prenotazione</h2>${html}</div>`));
});

app.post('/nuova', (req, res) => creaPrenotazione(req, res, 'bozza'));

app.get('/cliente', async (req, res) => {
  const html = await formPrenotazione('/cliente', null, null);
  res.send(page('Pagina cliente', `
    <div class="box">
      <h2>Richiesta prenotazione DP RENT</h2>
      <p class="notice">Il cliente non vede la targa. La targa resta interna.</p>
      ${html}
    </div>
  `));
});

app.post('/cliente', (req, res) => creaPrenotazione(req, res, 'richiesta_cliente'));

app.get('/planning', (req, res) => {
  const now = new Date();
  const ym = req.query.mese || `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  const [year, month] = ym.split('-').map(Number);
  const numDays = daysInMonth(year, month);

  db.all(`SELECT * FROM mezzi ORDER BY targa`, [], (err, mezzi) => {
    db.all(`SELECT * FROM prenotazioni WHERE stato!='annullato'`, [], (e2, pren) => {
      let header = `<th class="first">Mezzo</th>`;
      for (let d = 1; d <= numDays; d++) header += `<th>${d}</th>`;

      let rows = '';

      mezzi.forEach(m => {
        rows += `<tr><td class="first"><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a><br>${esc(descrizionePubblica(m))}</td>`;

        for (let d = 1; d <= numDays; d++) {
          const day = `${year}-${String(month).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
          const occ = pren.find(p => p.mezzo_id === m.id && dateInRange(day, p.data_inizio, p.data_fine));

          if (occ) {
            const title = `${occ.codice} - ${occ.cliente_nome || ''} ${occ.cliente_cognome || ''} - ${occ.telefono || ''}`;
            rows += `<td class="occupato" title="${esc(title)}" onclick="location.href='/prenotazione/${occ.id}'">O</td>`;
          } else {
            rows += `<td class="libero" title="Libero - clicca per prenotare" onclick="location.href='/nuova?mezzo_id=${m.id}&data=${day}'">L</td>`;
          }
        }

        rows += `</tr>`;
      });

      const prev = new Date(year, month - 2, 1);
      const next = new Date(year, month, 1);
      const prevYm = `${prev.getFullYear()}-${String(prev.getMonth() + 1).padStart(2, '0')}`;
      const nextYm = `${next.getFullYear()}-${String(next.getMonth() + 1).padStart(2, '0')}`;

      res.send(page('Planning', `
        <h2>Planning ${String(month).padStart(2, '0')}/${year}</h2>
        <p><a href="/planning?mese=${prevYm}">← Mese precedente</a> | <a href="/planning?mese=${nextYm}">Mese successivo →</a></p>
        <p><span class="libero" style="padding:6px">Libero</span> <span class="occupato" style="padding:6px">Occupato</span></p>
        <div class="sticky"><table><tr>${header}</tr>${rows}</table></div>
      `));
    });
  });
});

app.get('/storico', (req, res) => {
  const q = req.query.q || '';
  let sql = `
    SELECT p.*, m.targa, m.marca, m.modello, m.descrizione_pubblica, m.categoria
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id=p.mezzo_id
    WHERE 1=1
  `;
  const params = [];

  if (q) {
    sql += ` AND (p.codice LIKE ? OR p.cliente_nome LIKE ? OR p.cliente_cognome LIKE ? OR p.telefono LIKE ? OR m.targa LIKE ?)`;
    params.push(`%${q}%`, `%${q}%`, `%${q}%`, `%${q}%`, `%${q}%`);
  }

  sql += ` ORDER BY p.id DESC`;

  db.all(sql, params, (err, rows) => {
    const html = rows.map(p => `
      <tr>
        <td><a href="/prenotazione/${p.id}">${esc(p.codice)}</a></td>
        <td>${esc(p.cliente_nome)} ${esc(p.cliente_cognome)}</td>
        <td>${esc(p.telefono)}<br>${esc(p.email)}</td>
        <td>${esc(p.targa)}<br>${esc(descrizionePubblica(p))}</td>
        <td>${esc(p.data_inizio)} → ${esc(p.data_fine)}</td>
        <td>€ ${Number(p.totale || 0).toFixed(2)}</td>
        <td>${esc(p.stato)}</td>
        <td>
          <a href="/prenotazione/${p.id}">Apri</a><br>
          <a href="/contratto/${p.id}">PDF</a><br>
          <a href="/firma/${p.id}">Firma</a><br>
          <a href="/email/${p.id}">Email</a>
        </td>
      </tr>
    `).join('');

    res.send(page('Storico', `
      <h2>Storico</h2>
      <form method="GET">
        <input name="q" placeholder="Cerca nome, targa, codice, telefono" value="${esc(q)}">
        <button>Cerca</button>
      </form>
      <table>
        <tr><th>Codice</th><th>Cliente</th><th>Contatti</th><th>Mezzo</th><th>Date</th><th>Totale</th><th>Stato</th><th>Azioni</th></tr>
        ${html}
      </table>
    `));
  });
});

app.get('/prenotazione/:id', (req, res) => {
  db.get(`
    SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.descrizione_pubblica
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id=p.mezzo_id
    WHERE p.id=?
  `, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');

    res.send(page('Dettaglio contratto', `
      <div class="box">
        <h2>Contratto ${esc(p.codice)}</h2>
        <p><b>Cliente:</b> ${esc(p.cliente_nome)} ${esc(p.cliente_cognome)} - ${esc(p.telefono)}</p>
        <p><b>Conducente 1:</b> ${esc(p.conducente1_nome)} ${esc(p.conducente1_cognome)} - Patente ${esc(p.conducente1_patente)} - scad. ${esc(p.conducente1_patente_scadenza)}</p>
        <p><b>Conducente 2:</b> ${esc(p.conducente2_nome)} ${esc(p.conducente2_cognome)} - Patente ${esc(p.conducente2_patente)} - scad. ${esc(p.conducente2_patente_scadenza)}</p>
        <p><b>Mezzo:</b> ${esc(p.targa)} - ${esc(descrizionePubblica(p))}</p>
        <p><b>Date:</b> ${esc(p.data_inizio)} ${esc(p.ora_inizio)} → ${esc(p.data_fine)} ${esc(p.ora_fine)}</p>
        <p><b>Totale:</b> € ${Number(p.totale || 0).toFixed(2)}</p>
        <p><b>Stato:</b> ${esc(p.stato)}</p>
        <p><b>PDF Drive:</b> ${p.pdf_drive_web_link ? `<a target="_blank" href="${esc(p.pdf_drive_web_link)}">Apri PDF su Drive</a>` : 'non ancora caricato'}</p>

        <a class="btn" href="/contratto/${p.id}">Scarica/stampa PDF</a>
        <a class="btn btn2" href="/firma/${p.id}">Firma tablet</a>
        <a class="btn btn2" href="/email/${p.id}">Invia email</a>
        <a class="btn btn3" href="/documenti/${p.id}">Documenti/foto</a>
        <a class="btn btn3" href="/checkout/${p.id}">Check-out</a>
        <a class="btn btn3" href="/checkin/${p.id}">Check-in</a>
        <a class="btn btn2" href="/nexi/${p.id}">Pagamento Nexi</a>
      </div>
    `));
  });
});

function drawPdfHeader(doc) {
  doc.rect(0, 0, 612, 115).fill('#111');

  const logo = path.join(PUBLIC, 'logo.png');
  if (fs.existsSync(logo)) {
    try {
      doc.image(logo, 35, 25, { fit: [145, 70] });
    } catch {
      doc.fillColor('white').fontSize(28).text('DP RENT', 45, 35);
    }
  } else {
    doc.fillColor('white').fontSize(28).text('DP RENT', 45, 35);
  }

  doc.fillColor('white').fontSize(11);
  doc.text(AZIENDA.nome, 350, 25, { width: 210, align: 'right' });
  doc.text(AZIENDA.indirizzo, 350, 43, { width: 210, align: 'right' });
  doc.text(`P.IVA / CF ${AZIENDA.piva} | Tel. ${AZIENDA.telefono}`, 350, 61, { width: 210, align: 'right' });
  doc.text(AZIENDA.email, 350, 79, { width: 210, align: 'right' });
  doc.rect(0, 115, 612, 6).fill('#d90000');
  doc.fillColor('black');
}

function pdfSection(doc, title, x, y, w) {
  doc.rect(x, y, w, 20).fill('#111');
  doc.fillColor('white').fontSize(10).text(title, x + 8, y + 6);
  doc.fillColor('black');
  return y + 26;
}

function pdfRow(doc, label, value, x, y, w) {
  doc.fillColor('#777').fontSize(8).text(label, x, y, { width: w * 0.38 });
  doc.fillColor('#111').fontSize(9).text(value || '', x + w * 0.38, y, { width: w * 0.6 });
}

function generaPdf(id, cb) {
  db.get(`
    SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.descrizione_pubblica, m.km_inclusi
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id=p.mezzo_id
    WHERE p.id=?
  `, [id], (err, p) => {
    if (err || !p) return cb(err || new Error('Contratto non trovato'));

    const file = path.join(CONTRACTS, `contratto_${String(p.codice || id).replace(/[^a-zA-Z0-9_-]/g, '')}.pdf`);
    const doc = new PDFDocument({ size: 'A4', margin: 40 });
    const stream = fs.createWriteStream(file);

    doc.pipe(stream);
    drawPdfHeader(doc);

    doc.fontSize(20).fillColor('#111').text('CONTRATTO DI NOLEGGIO', 40, 150, { width: 515, align: 'center' });
    doc.moveTo(40, 180).lineTo(555, 180).strokeColor('#d90000').stroke();

    let y = 205;
    y = pdfSection(doc, 'DATI CONTRATTO', 45, y, 510);
    pdfRow(doc, 'Numero contratto', p.codice, 55, y, 460); y += 18;
    pdfRow(doc, 'Stato', p.stato, 55, y, 460); y += 18;
    pdfRow(doc, 'Data creazione', p.created_at, 55, y, 460); y += 30;

    let yl = y, yr = y;

    yl = pdfSection(doc, 'ANAGRAFICA CLIENTE', 45, yl, 245);
    pdfRow(doc, 'Cliente', `${p.cliente_nome || ''} ${p.cliente_cognome || ''}`, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Telefono', p.telefono, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Email', p.email, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Codice fiscale', p.codice_fiscale, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Indirizzo', p.indirizzo, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Tipo/P.IVA', `${p.tipo_cliente || ''} ${p.piva || ''}`, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Conducente 1', `${p.conducente1_nome || ''} ${p.conducente1_cognome || ''}`, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Patente 1', `${p.conducente1_patente || ''} scad. ${p.conducente1_patente_scadenza || ''}`, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Conducente 2', `${p.conducente2_nome || ''} ${p.conducente2_cognome || ''}`, 55, yl, 220); yl += 18;

    yr = pdfSection(doc, 'DETTAGLI PRENOTAZIONE', 310, yr, 245);
    pdfRow(doc, 'Check-out', `${p.data_inizio} ore ${p.ora_inizio || ''}`, 320, yr, 220); yr += 18;
    pdfRow(doc, 'Check-in', `${p.data_fine} ore ${p.ora_fine || ''}`, 320, yr, 220); yr += 18;
    pdfRow(doc, 'Giorni tariffati', String(p.giorni || ''), 320, yr, 220); yr += 18;
    pdfRow(doc, 'Extra fuori orario', Number(p.extra_fuori_orario || 0) ? `€ ${Number(p.extra_fuori_orario).toFixed(2)} + IVA` : 'NO', 320, yr, 220); yr += 18;
    pdfRow(doc, 'Carburante', `${p.carburante_uscita || ''} / ${p.carburante_rientro || ''}`, 320, yr, 220); yr += 18;
    pdfRow(doc, 'Km uscita/rientro', `${p.km_uscita || ''} / ${p.km_rientro || ''}`, 320, yr, 220); yr += 18;

    y = Math.max(yl, yr) + 10;
    yl = y; yr = y;

    yl = pdfSection(doc, 'VEICOLO', 45, yl, 245);
    pdfRow(doc, 'Targa', p.targa, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Descrizione', descrizionePubblica(p), 55, yl, 220); yl += 18;
    pdfRow(doc, 'Categoria', p.categoria, 55, yl, 220); yl += 18;
    pdfRow(doc, 'Km inclusi totali', String(Number(p.km_inclusi || 0) * Number(p.giorni || 0)), 55, yl, 220); yl += 18;
    pdfRow(doc, 'Km previsti', String(p.km_previsti || 0), 55, yl, 220); yl += 18;

    yr = pdfSection(doc, 'RIEPILOGO ECONOMICO', 310, yr, 245);
    pdfRow(doc, 'Imponibile', `€ ${Number(p.imponibile || 0).toFixed(2)}`, 320, yr, 220); yr += 18;
    pdfRow(doc, 'IVA 22%', `€ ${Number(p.iva || 0).toFixed(2)}`, 320, yr, 220); yr += 18;
    pdfRow(doc, 'Totale IVA inclusa', `€ ${Number(p.totale || 0).toFixed(2)}`, 320, yr, 220); yr += 18;
    pdfRow(doc, 'Deposito cauzionale', `€ ${Number(p.cauzione || 0).toFixed(2)}`, 320, yr, 220); yr += 18;

    y = Math.max(yl, yr) + 15;
    y = pdfSection(doc, 'CONDIZIONI GENERALI E PRIVACY', 45, y, 510);

    doc.fontSize(8).fillColor('#111').text(
      'Il cliente dichiara di aver preso visione e accettare le condizioni generali di noleggio e l’informativa privacy DP RENT / Trasporti DP S.R.L.',
      55, y, { width: 490 }
    );
    y += 24;
    doc.fontSize(7).text(`Condizioni generali: ${TERMS_URL}`, 55, y, { width: 490 }); y += 12;
    doc.text(`Informativa privacy: ${PRIVACY_URL}`, 55, y, { width: 490 }); y += 20;

    doc.fontSize(8).text(
      'Condizioni principali: veicolo consegnato con il pieno e da riconsegnare con il pieno; extra km €0,15/km ove previsto; danni, multe, pedaggi, franchigie, smarrimenti e costi accessori a carico del cliente.',
      55, y, { width: 490 }
    );
    y += 45;

    doc.fontSize(10).text('Firma cliente:', 55, y);
    if (p.firma_path && fs.existsSync(p.firma_path)) {
      doc.image(p.firma_path, 55, y + 15, { fit: [220, 70] });
    } else {
      doc.text('______________________________', 55, y + 25);
    }

    doc.text('Firma DP RENT:', 330, y);
    doc.text('______________________________', 330, y + 25);

    doc.end();

    stream.on('finish', async () => {
      let drive = null;

      try {
        drive = await driveUpload(
          file,
          path.basename(file),
          'application/pdf',
          `${p.codice || 'contratto'} - ${p.cliente_nome || ''} ${p.cliente_cognome || ''}`
        );
      } catch (e) {
        console.log('Drive PDF errore:', e.message);
      }

      if (drive) {
        db.run(`UPDATE prenotazioni SET pdf_path=?, pdf_drive_file_id=?, pdf_drive_web_link=? WHERE id=?`, [file, drive.id, drive.link, id]);
      } else {
        db.run(`UPDATE prenotazioni SET pdf_path=? WHERE id=?`, [file, id]);
      }

      cb(null, file);
    });

    stream.on('error', cb);
  });
}

app.get('/contratto/:id', (req, res) => {
  generaPdf(req.params.id, (err, file) => {
    if (err) return res.send('Errore PDF: ' + esc(err.message));
    res.download(file);
  });
});

app.get('/firma/:id', (req, res) => {
  res.send(page('Firma', `
    <div class="box">
      <h2>Firma contratto</h2>
      <canvas id="canvas"></canvas>
      <button onclick="clearCanvas()">Cancella</button>
      <button onclick="saveFirma()">Salva firma</button>
    </div>
    <script>
      const canvas=document.getElementById('canvas');
      const ctx=canvas.getContext('2d');
      canvas.width=canvas.offsetWidth;
      canvas.height=230;
      let drawing=false;
      function pos(e){const r=canvas.getBoundingClientRect();const t=e.touches?e.touches[0]:e;return{x:t.clientX-r.left,y:t.clientY-r.top};}
      function start(e){drawing=true;const p=pos(e);ctx.beginPath();ctx.moveTo(p.x,p.y);e.preventDefault();}
      function move(e){if(!drawing)return;const p=pos(e);ctx.lineTo(p.x,p.y);ctx.stroke();e.preventDefault();}
      function end(e){drawing=false;e.preventDefault();}
      canvas.addEventListener('mousedown',start);
      canvas.addEventListener('mousemove',move);
      canvas.addEventListener('mouseup',end);
      canvas.addEventListener('touchstart',start);
      canvas.addEventListener('touchmove',move);
      canvas.addEventListener('touchend',end);
      function clearCanvas(){ctx.clearRect(0,0,canvas.width,canvas.height);}
      function saveFirma(){
        fetch('/firma/${req.params.id}',{
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body:JSON.stringify({firma:canvas.toDataURL('image/png')})
        }).then(()=>location.href='/prenotazione/${req.params.id}');
      }
    </script>
  `));
});

app.post('/firma/:id', (req, res) => {
  const data = req.body.firma;
  if (!data) return res.status(400).send('Firma mancante');

  const base64 = data.split(',')[1];
  const file = path.join(UPLOADS, `firma_${req.params.id}.png`);
  fs.writeFileSync(file, base64, 'base64');

  db.run(`UPDATE prenotazioni SET firma_path=?, stato='firmato' WHERE id=?`, [file, req.params.id], () => {
    generaPdf(req.params.id, () => res.send('OK'));
  });
});

app.get('/documenti/:id', (req, res) => {
  db.all(`SELECT * FROM allegati WHERE prenotazione_id=? ORDER BY id DESC`, [req.params.id], (err, files) => {
    const lista = files.map(f => `<li>${esc(f.tipo)} - ${esc(f.originalname)} ${f.drive_web_link ? `- <a target="_blank" href="${esc(f.drive_web_link)}">Drive</a>` : ''}</li>`).join('');

    res.send(page('Documenti', `
      <div class="box">
        <h2>Documenti / Foto contratto</h2>
        <form method="POST" action="/documenti/${req.params.id}" enctype="multipart/form-data">
          <select name="tipo">
            <option>Patente conducente 1 fronte</option>
            <option>Patente conducente 1 retro</option>
            <option>Documento conducente 1 fronte</option>
            <option>Documento conducente 1 retro</option>
            <option>Patente conducente 2 fronte</option>
            <option>Patente conducente 2 retro</option>
            <option>Documento conducente 2 fronte</option>
            <option>Documento conducente 2 retro</option>
            <option>Foto uscita fronte</option>
            <option>Foto uscita retro</option>
            <option>Foto uscita lato dx</option>
            <option>Foto uscita lato sx</option>
            <option>Foto uscita interno</option>
            <option>Foto danni uscita</option>
            <option>Foto rientro fronte</option>
            <option>Foto rientro retro</option>
            <option>Foto rientro lato dx</option>
            <option>Foto rientro lato sx</option>
            <option>Foto rientro interno</option>
            <option>Foto danni rientro</option>
          </select>
          <input type="file" name="file" accept="image/*,.pdf" capture="environment" required>
          <button>Carica</button>
        </form>
        <ul>${lista}</ul>
        <a class="btn" href="/prenotazione/${req.params.id}">Torna contratto</a>
      </div>
    `));
  });
});

app.post('/documenti/:id', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');

  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], async (err, p) => {
    let drive = null;

    try {
      drive = await driveUpload(
        req.file.path,
        `${Date.now()}_${req.body.tipo}_${req.file.originalname}`,
        req.file.mimetype,
        `${p.codice || 'contratto'} - ${p.cliente_nome || ''} ${p.cliente_cognome || ''}`
      );
    } catch (e) {
      console.log('Drive documento errore:', e.message);
    }

    db.run(`
      INSERT INTO allegati (prenotazione_id,tipo,filename,originalname,path,mimetype,drive_file_id,drive_web_link)
      VALUES (?,?,?,?,?,?,?,?)
    `, [
      req.params.id, req.body.tipo, req.file.filename, req.file.originalname, req.file.path, req.file.mimetype,
      drive ? drive.id : null, drive ? drive.link : null
    ], () => res.redirect(`/documenti/${req.params.id}`));
  });
});

app.get('/checkout/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');
    res.send(page('Check-out', `
      <div class="box">
        <h2>Check-out</h2>
        <form method="POST">
          <label>Carburante uscita</label>
          <select name="carburante_uscita">${fuelOptions(p.carburante_uscita)}</select>
          <label>Km uscita</label>
          <input type="number" name="km_uscita" value="${esc(p.km_uscita)}">
          <label>Note</label>
          <textarea name="note">${esc(p.note)}</textarea>
          <button>Salva check-out</button>
        </form>
        <a class="btn btn3" href="/documenti/${p.id}">Carica foto uscita</a>
      </div>
    `));
  });
});

app.post('/checkout/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    db.run(`UPDATE prenotazioni SET carburante_uscita=?, km_uscita=?, note=?, stato='in_corso' WHERE id=?`,
      [req.body.carburante_uscita, req.body.km_uscita, req.body.note, req.params.id],
      () => {
        if (req.body.km_uscita) db.run(`UPDATE mezzi SET km_attuali=? WHERE id=?`, [req.body.km_uscita, p.mezzo_id]);
        res.send(renderActionScreen(req.params.id, 'Check-out salvato', 'Contratto aggiornato.'));
      });
  });
});

app.get('/checkin/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');
    res.send(page('Check-in', `
      <div class="box">
        <h2>Check-in</h2>
        <form method="POST">
          <label>Carburante rientro</label>
          <select name="carburante_rientro">${fuelOptions(p.carburante_rientro)}</select>
          <label>Km rientro</label>
          <input type="number" name="km_rientro" value="${esc(p.km_rientro)}">
          <label>Note</label>
          <textarea name="note">${esc(p.note)}</textarea>
          <button>Salva check-in</button>
        </form>
        <a class="btn btn3" href="/documenti/${p.id}">Carica foto rientro</a>
      </div>
    `));
  });
});

app.post('/checkin/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    db.run(`UPDATE prenotazioni SET carburante_rientro=?, km_rientro=?, note=?, stato='rientrato' WHERE id=?`,
      [req.body.carburante_rientro, req.body.km_rientro, req.body.note, req.params.id],
      () => {
        if (req.body.km_rientro) {
          db.run(`UPDATE mezzi SET km_attuali=? WHERE id=?`, [req.body.km_rientro, p.mezzo_id]);
        }

        db.get(`SELECT * FROM mezzi WHERE id=?`, [p.mezzo_id], async (e2, mezzo) => {
          const alerts = alertMezzo({ ...mezzo, km_attuali: req.body.km_rientro || mezzo.km_attuali });
          if (alerts.length) {
            await sendAlert(
              `ALERT DP RENT - ${mezzo.targa}`,
              `Mezzo: ${mezzo.targa} ${mezzo.modello || ''}\nKm rientro: ${req.body.km_rientro}\n\n${alerts.join('\n')}`
            );
          }
          res.send(renderActionScreen(req.params.id, 'Check-in salvato', 'Contratto rientrato e km aggiornati.'));
        });
      });
  });
});

app.get('/email/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');
    res.send(page('Invia email', `
      <div class="box">
        <h2>Invia contratto</h2>
        <form method="POST">
          <label>Email destinatario</label>
          <input name="email" value="${esc(p.email)}" required>
          <label>Messaggio</label>
          <textarea name="messaggio">Buongiorno, in allegato trova il contratto DP RENT.</textarea>
          <button>Invia email</button>
        </form>
      </div>
    `));
  });
});

app.post('/email/:id', (req, res) => {
  generaPdf(req.params.id, async (err, file) => {
    if (err) return res.send('Errore PDF: ' + esc(err.message));

    try {
      await sendEmail(req.body.email, 'Contratto DP RENT', req.body.messaggio || 'Contratto DP RENT in allegato.', [
        { filename: path.basename(file), path: file }
      ]);
      res.send(renderActionScreen(req.params.id, 'Email inviata', 'Contratto inviato correttamente.'));
    } catch (e) {
      res.send(page('Errore email', `<div class="box"><h2 class="bad">Errore email</h2><pre>${esc(e.message)}</pre></div>`));
    }
  });
});

app.get('/nexi/:id', (req, res) => {
  const base = process.env.APP_BASE_URL || '';
  const link = `${base}/pagamento-demo/${req.params.id}`;
  db.run(`UPDATE prenotazioni SET nexi_payment_url=?, nexi_stato='link_generato' WHERE id=?`, [link, req.params.id]);
  res.send(renderActionScreen(req.params.id, 'Pagamento Nexi predisposto', `Link provvisorio: <a href="${esc(link)}">${esc(link)}</a>`));
});

app.get('/pagamento-demo/:id', (req, res) => {
  res.send(page('Pagamento demo', `<div class="box"><h2>Pagamento Nexi demo</h2><p>Qui colleghiamo PayMail reale quando mettiamo Alias e MAC key.</p></div>`));
});

app.get('/test-email', async (req, res) => {
  try {
    await sendEmail(
      process.env.ALERT_EMAIL || process.env.SMTP_USER,
      'TEST DP RENT APP',
      'Se ricevi questa email, SMTP Gmail funziona.',
      []
    );
    res.send(page('Test Email', `<div class="box"><h2 class="ok">Email inviata</h2></div>`));
  } catch (e) {
    res.send(page('Errore Email', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/test-drive', async (req, res) => {
  try {
    if (!googleDriveConfigured()) {
      return res.send(page('Drive non configurato', `<div class="box"><h2 class="bad">Google Drive non configurato</h2><pre>GOOGLE_SERVICE_ACCOUNT_EMAIL
GOOGLE_PRIVATE_KEY
GOOGLE_DRIVE_FOLDER_ID</pre></div>`));
    }

    const file = path.join(UPLOADS, `test-drive-${Date.now()}.txt`);
    fs.writeFileSync(file, 'Test Google Drive DP RENT');
    const drive = await driveUpload(file, path.basename(file), 'text/plain', 'TEST DP RENT');

    res.send(page('Test Drive', `<div class="box"><h2 class="ok">Upload Drive riuscito</h2><a target="_blank" href="${esc(drive.link)}">Apri su Google Drive</a></div>`));
  } catch (e) {
    res.send(page('Errore Drive', `<div class="box"><h2 class="bad">Errore</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.listen(PORT, () => {
  console.log(`DP RENT APP attiva su porta ${PORT}`);
});
