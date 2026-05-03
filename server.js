require('dotenv').config();
require('dns').setDefaultResultOrder('ipv4first');

const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const XLSX = require('xlsx');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 10000;

app.use(bodyParser.json({ limit: '80mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '80mb' }));

const DATA_DIR = path.join(__dirname, 'data');
const UPLOAD_DIR = path.join(__dirname, 'uploads');
const CONTRACT_DIR = path.join(__dirname, 'contracts');
const PUBLIC_DIR = path.join(__dirname, 'public');
const DB_FILE = path.join(DATA_DIR, 'db.json');

[DATA_DIR, UPLOAD_DIR, CONTRACT_DIR, PUBLIC_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

app.use('/uploads', express.static(UPLOAD_DIR));
app.use('/contracts', express.static(CONTRACT_DIR));
app.use('/public', express.static(PUBLIC_DIR));

const upload = multer({ dest: UPLOAD_DIR });

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
const CAUZIONE = 500;

const TERMS_URL = '';
const PRIVACY_URL = '';
function defaultDb() {
  return {
    counters: {
      mezzi: 0,
      prenotazioni: 0,
      allegati: 0
    },
    mezzi: [],
    prenotazioni: [],
    allegati: []
  };
}

function loadDb() {
  if (!fs.existsSync(DB_FILE)) {
    const db = defaultDb();
    fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2));
    return db;
  }

  try {
    const db = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
    db.counters = db.counters || { mezzi: 0, prenotazioni: 0, allegati: 0 };
    db.mezzi = db.mezzi || [];
    db.prenotazioni = db.prenotazioni || [];
    db.allegati = db.allegati || [];
    return db;
  } catch (err) {
    console.log('Errore lettura DB:', err.message);
    return defaultDb();
  }
}

function saveDb(db) {
  fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2));
}

function esc(v) {
  return String(v ?? '').replace(/[&<>"']/g, s => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  }[s]));
}

function euro(v) {
  return Number(v || 0).toFixed(2);
}

function nowIt() {
  return new Date().toLocaleString('it-IT', { timeZone: 'Europe/Rome' });
}

function todayCode() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

function isValidDateString(v) {
  if (!v) return false;
  const d = new Date(v + 'T00:00:00');
  return !isNaN(d.getTime());
}

function giorniNoleggio(dataInizio, dataFine) {
  const d1 = new Date(dataInizio + 'T00:00:00');
  const d2 = new Date(dataFine + 'T00:00:00');
  const diff = Math.floor((d2 - d1) / 86400000) + 1;
  return Math.max(1, diff || 1);
}

function validaDate(dataInizio, dataFine) {
  if (!isValidDateString(dataInizio) || !isValidDateString(dataFine)) {
    return 'Data inizio o data fine non valida';
  }

  const d1 = new Date(dataInizio + 'T00:00:00');
  const d2 = new Date(dataFine + 'T00:00:00');

  if (d2 < d1) {
    return 'Errore: la data fine non può essere precedente alla data inizio';
  }

  return '';
}

function formatDate(v) {
  if (!v) return '';
  const d = new Date(v + 'T00:00:00');
  if (isNaN(d.getTime())) return v;
  return d.toLocaleDateString('it-IT');
}

function extraOrario(ora) {
  if (!ora) return 0;
  const [h, m] = ora.split(':').map(Number);
  const minuti = h * 60 + (m || 0);
  return minuti < 510 || minuti > 1110 ? EXTRA_FUORI_ORARIO : 0;
}
function categoriaAuto(marca, modello, descrizione, codice) {
  const s = `${marca || ''} ${modello || ''} ${descrizione || ''} ${codice || ''}`.toUpperCase();

  if (s.includes('GOLF')) return 'AUTO_GOLF';
  if (s.includes('DACIA')) return 'AUTO_DACIA';
  if (s.includes('ESCAVATORE')) return 'ESCAVATORE';
  if (s.includes('SEMOVENTE') || s.includes('PIATTAFORMA')) return 'SEMOVENTE';
  if (s.includes('9 POSTI') || s.includes('9P') || s.includes('TOURNEO') || s.includes('PULMINO')) return '9_POSTI';

  return 'FURGONE';
}

function prezzoCategoria(categoria) {
  if (categoria === 'AUTO_DACIA') return 50;
  if (categoria === 'AUTO_GOLF') return 60;
  if (categoria === 'ESCAVATORE') return 50;
  if (categoria === 'SEMOVENTE') return 50;
  return 70;
}

function kmCategoria(categoria) {
  if (categoria === 'ESCAVATORE' || categoria === 'SEMOVENTE') return 0;
  return 150;
}

function descrizioneCliente(m) {
  if (!m) return '';
  if (m.descrizione_cliente) return m.descrizione_cliente;

  const base = `${m.marca || ''} ${m.modello || ''}`.trim();

  if (m.categoria === '9_POSTI') return `${base} - pulmino 9 posti`;
  if (m.categoria === 'FURGONE') return `${base} - furgone cargo`;
  if (m.categoria === 'AUTO_DACIA') return `${base} - auto economica`;
  if (m.categoria === 'AUTO_GOLF') return `${base} - auto categoria Golf`;
  if (m.categoria === 'ESCAVATORE') return `${base} - escavatore`;
  if (m.categoria === 'SEMOVENTE') return `${base} - semovente / piattaforma`;

  return base || m.descrizione || 'Mezzo DP RENT';
}

function calcolaTotale(p, mezzo) {
  const giorni = giorniNoleggio(p.data_inizio, p.data_fine);
  const prezzo = Number(mezzo.prezzo_giorno || prezzoCategoria(mezzo.categoria));
  const kmGiorno = Number(mezzo.km_inclusi || kmCategoria(mezzo.categoria));
  const kmInclusi = giorni * kmGiorno;
  const kmPrevisti = Number(p.km_previsti || 0);

  const extraKm = kmGiorno > 0 ? Math.max(0, kmPrevisti - kmInclusi) * EXTRA_KM : 0;
  const extraFuoriOrario = extraOrario(p.ora_inizio) + extraOrario(p.ora_fine);

  const imponibile = giorni * prezzo + extraKm + extraFuoriOrario;
  const iva = imponibile * IVA;
  const totale = imponibile + iva;

  return {
    giorni,
    kmInclusi,
    extraKm,
    extraFuoriOrario,
    imponibile,
    iva,
    totale
  };
}

function fuelOptions(selected) {
  return ['4/4 pieno', '3/4', '1/2', '1/4', 'Riserva', 'Vuoto']
    .map(v => `<option value="${esc(v)}" ${v === selected ? 'selected' : ''}>${esc(v)}</option>`)
    .join('');
}

function checkAlerts(m) {
  const list = [];
  const km = Number(m.km || 0);
  const alertKm = Number(m.alert_km || 1000);
  const alertGiorni = Number(m.alert_giorni || 30);

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (m.tagliando_km) {
    const diffKm = Number(m.tagliando_km) - km;

    if (diffKm <= 0) {
      list.push(`❌ Tagliando scaduto di ${Math.abs(diffKm)} km`);
    } else if (diffKm <= alertKm) {
      list.push(`⚠️ Tagliando vicino: mancano ${diffKm} km`);
    }
  }

  function dataAlert(label, value) {
    if (!value) return;

    const d = new Date(value + 'T00:00:00');
    if (isNaN(d.getTime())) return;

    const diff = Math.ceil((d - today) / 86400000);

    if (diff < 0) {
      list.push(`❌ ${label} scaduto il ${formatDate(value)}`);
    } else if (diff <= alertGiorni) {
      list.push(`⚠️ ${label} in scadenza il ${formatDate(value)} (${diff} giorni)`);
    }
  }

  dataAlert('Revisione', m.revisione);
  dataAlert('Bollo', m.bollo);
  dataAlert('Assicurazione', m.assicurazione);
  dataAlert('Gomme / manutenzione', m.gomme);

  return list;
}

function googleDriveConfigured() {
  return !!(
    process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL &&
    process.env.GOOGLE_PRIVATE_KEY &&
    process.env.GOOGLE_DRIVE_FOLDER_ID
  );
}

function getDrive() {
  const privateKey = String(process.env.GOOGLE_PRIVATE_KEY || '').replace(/\\n/g, '\n');

  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: privateKey,
    scopes: ['https://www.googleapis.com/auth/drive.file']
  });

  return google.drive({ version: 'v3', auth });
}

async function driveUpload(localPath, filename, mimetype, folderName) {
  if (!process.env.DRIVE_WEBAPP_URL) return null;
  if (!process.env.GOOGLE_DRIVE_FOLDER_ID) return null;
  if (!fs.existsSync(localPath)) return null;

  const base64 = fs.readFileSync(localPath).toString('base64');

  const response = await fetch(process.env.DRIVE_WEBAPP_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({
      folderId: process.env.GOOGLE_DRIVE_FOLDER_ID,
      subfolder: folderName || 'DP RENT',
      filename: filename,
      mimeType: mimetype || 'application/octet-stream',
      base64: base64
    })
  });

  const text = await response.text();

  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error('Errore risposta Apps Script: ' + text);
  }

  if (!data.ok) {
    throw new Error(data.error || 'Errore upload Drive');
  }

  return {
    id: data.id || '',
    link: data.link || ''
  };
}
  
  return {
    id: uploaded.data.id,
    link: uploaded.data.webViewLink
  };
}

async function sendEmail(to, subject, text, attachments = []) {
  if (!process.env.SMTP_HOST) {
    throw new Error('SMTP non configurato');
  }

  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 465),
    secure: String(process.env.SMTP_SECURE || '').toLowerCase() === 'true',
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    },
    tls: {
      rejectUnauthorized: false
    }
  });

  await transporter.verify();

  return transporter.sendMail({
    from: process.env.SMTP_FROM || process.env.SMTP_USER,
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
  } catch (err) {
    console.log('Errore alert email:', err.message);
  }
}

function nexiConfigured() {
  return !!(process.env.NEXI_ALIAS && process.env.NEXI_MAC_KEY);
}

function nexiPlaceholderUrl(p) {
  const base = process.env.APP_BASE_URL || '';
  const amountCents = Math.round(Number(p.totale || 0) * 100);
  const transId = `DPR${p.id}${Date.now()}`;

  const macString = `codTrans=${transId}divisa=EURimporto=${amountCents}${process.env.NEXI_MAC_KEY || ''}`;
  const mac = crypto.createHash('sha1').update(macString).digest('hex');

  return {
    ready: nexiConfigured(),
    transId,
    amountCents,
    mac,
    callback: `${base}/nexi-callback`
  };
}

function layout(title, body) {
  const logoPath = path.join(PUBLIC_DIR, 'logo.png');

  const logo = fs.existsSync(logoPath)
    ? `<img src="/public/logo.png" style="height:54px;background:white;border-radius:8px;padding:4px">`
    : `<span>DP RENT APP</span>`;

  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${esc(title)}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{margin:0;font-family:Arial;background:#f4f4f4;color:#111}
header{background:#070707;color:white;padding:22px;font-size:32px;font-weight:900;display:flex;align-items:center;gap:15px}
nav{background:#c40000;padding:12px;display:flex;gap:10px;flex-wrap:wrap}
nav a{color:white;text-decoration:none;font-weight:bold;padding:8px}
main{padding:18px}
.card{background:white;border-radius:14px;padding:18px;margin:12px 0;box-shadow:0 2px 10px #ccc}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px}
.tile{background:#111;color:white;border-radius:20px;padding:32px;text-align:center;text-decoration:none;font-size:23px;font-weight:900;border-bottom:8px solid #c40000}
input,select,textarea,button{width:100%;box-sizing:border-box;padding:11px;margin:5px 0 12px;border:1px solid #aaa;border-radius:7px}
button,.btn{background:#c40000;color:white;border:0;border-radius:8px;padding:11px 16px;text-decoration:none;font-weight:bold;display:inline-block;width:auto;margin:4px}
.btn2{background:#222}.btn3{background:#0b6b2d}.ok{color:green}.bad{color:#c40000}.warn{color:#bd6b00}
table{width:100%;border-collapse:collapse;background:white}
th,td{border:1px solid #ddd;padding:7px;font-size:13px;vertical-align:top}
th{background:#111;color:white}
.free{background:#32b849;color:white;text-align:center;font-weight:bold;cursor:pointer}
.busy{background:#d60000;color:white;text-align:center;font-weight:bold;cursor:pointer}
.alert{background:#ffe2e2;border:1px solid #c40000;border-radius:8px;padding:9px;margin:7px 0}
.notice{background:#fff4ce;border:1px solid #d8bd54;border-radius:8px;padding:9px;margin:7px 0}
.sticky{overflow:auto;max-height:78vh}
.sticky th{position:sticky;top:0;z-index:2}
.sticky .first{position:sticky;left:0;background:white;z-index:1;min-width:190px}
.sticky th.first{background:#111;color:white;z-index:3}
canvas{width:100%;height:260px;border:2px solid #111;background:white;touch-action:none}
@media(max-width:700px){header{font-size:25px}main{padding:8px}nav a{font-size:14px}.tile{font-size:19px;padding:24px}}
</style>
</head>
<body>
<header>${logo}</header>
<nav>
<a href="/">Dashboard</a>
<a href="/mezzi">Mezzi</a>
<a href="/import">Import Excel</a>
<a href="/nuova">Nuova prenotazione</a>
<a href="/planning">Planning</a>
<a href="/storico">Storico</a>
<a href="/rientri">Rientri</a>
<a href="/scadenze">Scadenze</a>
<a href="/cliente">Pagina cliente</a>
<a href="/logo">Logo</a>
<a href="/test-email">Test Email</a>
<a href="/test-drive">Test Drive</a>
</nav>
<main>${body}</main>
</body>
</html>`;
}
app.get('/', (req, res) => {
  const db = loadDb();

  const alerts = db.mezzi
    .flatMap(m => checkAlerts(m).map(a => `
      <div class="alert">
        <b>${esc(m.targa)} ${esc(m.descrizione)}</b><br>${esc(a)}
      </div>
    `))
    .join('');

  res.send(layout('Dashboard', `
    <div class="grid">
      <a class="tile" href="/nuova">➕ Nuova prenotazione</a>
      <a class="tile" href="/planning">📅 Planning</a>
      <a class="tile" href="/mezzi">🚐 Mezzi</a>
      <a class="tile" href="/rientri">🔁 Rientri / Check-in</a>
      <a class="tile" href="/storico">📁 Storico</a>
      <a class="tile" href="/scadenze">⚠️ Scadenze</a>
      <a class="tile" href="/cliente">📲 Pagina cliente</a>
      <a class="tile" href="/import">📊 Import Excel</a>
    </div>

    <div class="card">
      <h2>Situazione</h2>
      <p>Mezzi: <b>${db.mezzi.length}</b></p>
      <p>Contratti: <b>${db.prenotazioni.length}</b></p>
      <p>Email: <b>${esc(process.env.SMTP_HOST || 'non configurata')}</b></p>
      <p>Google Drive: <b>${googleDriveConfigured() ? 'configurato' : 'non configurato'}</b></p>
      <p>Nexi: <b>${nexiConfigured() ? 'predisposto' : 'non configurato'}</b></p>
    </div>

    <div class="card">
      <h2>Alert</h2>
      ${alerts || '<p class="ok">Nessun alert.</p>'}
    </div>
  `));
});

app.get('/test-email', async (req, res) => {
  try {
    await sendEmail(
      process.env.ALERT_EMAIL || process.env.SMTP_USER,
      'TEST DP RENT',
      'Email funzionante DP RENT'
    );
    res.send('EMAIL OK');
  } catch (err) {
    res.send('ERRORE EMAIL: ' + err.message);
  }
});

app.get('/test-drive', async (req, res) => {
  try {
    if (!googleDriveConfigured()) {
      return res.send(layout('Drive non configurato', `
        <div class="card">
          <h2 class="bad">DRIVE NON CONFIGURATO</h2>
          <p>Servono queste variabili Render:</p>
          <pre>GOOGLE_SERVICE_ACCOUNT_EMAIL
GOOGLE_PRIVATE_KEY
GOOGLE_DRIVE_FOLDER_ID</pre>
        </div>
      `));
    }

    const file = path.join(UPLOAD_DIR, `test-drive-${Date.now()}.txt`);
    fs.writeFileSync(file, 'Test Google Drive DP RENT');

    const uploaded = await driveUpload(
      file,
      path.basename(file),
      'text/plain',
      'TEST DP RENT'
    );

    res.send(layout('Drive OK', `
      <div class="card">
        <h2 class="ok">DRIVE OK</h2>
        <a class="btn" href="${esc(uploaded.link)}" target="_blank">Apri file su Drive</a>
      </div>
    `));
  } catch (err) {
    res.send(layout('Errore Drive', `
      <div class="card">
        <h2 class="bad">ERRORE DRIVE</h2>
        <pre>${esc(err.message)}</pre>
      </div>
    `));
  }
});

app.get('/logo', (req, res) => {
  res.send(layout('Logo', `
    <div class="card">
      <h2>Logo DP RENT</h2>
      <form method="POST" action="/logo" enctype="multipart/form-data">
        <input type="file" name="logo" accept="image/*" required>
        <button>Salva logo</button>
      </form>
    </div>
  `));
});

app.post('/logo', multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, PUBLIC_DIR),
    filename: (req, file, cb) => cb(null, 'logo.png')
  })
}).single('logo'), (req, res) => {
  res.redirect('/');
});

app.get('/import', (req, res) => {
  res.send(layout('Import Excel', `
    <div class="card">
      <h2>Import mezzi da Excel</h2>
      <p>Colonne lette automaticamente:</p>
      <p><b>Targa, Marca, Modello, Km, Descrizione, Codice Tipo</b></p>
      <form method="POST" action="/import" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx,.xls,.csv" required>
        <button>Importa mezzi</button>
      </form>
    </div>
  `));
});

app.post('/import', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');

  const db = loadDb();
  const wb = XLSX.readFile(req.file.path);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);

  let count = 0;

  rows.forEach(r => {
    const targa = String(r.Targa || r.targa || '').trim();
    if (!targa) return;

    const marca = String(r.Marca || r.marca || '').trim();
    const modello = String(r.Modello || r.modello || '').trim();
    const descrizione = String(r.Descrizione || r.descrizione || '').trim();
    const codice = String(r['Codice Tipo'] || r['Codice Tip'] || '').trim();
    const km = Number(r.Km || r.km || r['Km percor'] || r['Km percorsi'] || 0);

    const categoria = categoriaAuto(marca, modello, descrizione, codice);

    let mezzo = db.mezzi.find(m => String(m.targa).toUpperCase() === targa.toUpperCase());

    if (!mezzo) {
      db.counters.mezzi++;
      mezzo = { id: db.counters.mezzi };
      db.mezzi.push(mezzo);
    }

    mezzo.targa = targa;
    mezzo.marca = marca;
    mezzo.modello = modello;
    mezzo.categoria = categoria;
    mezzo.descrizione = mezzo.descrizione || descrizioneCliente({ marca, modello, categoria });
    mezzo.descrizione_cliente = mezzo.descrizione_cliente || descrizioneCliente({ marca, modello, categoria });
    mezzo.km = km || mezzo.km || 0;
    mezzo.prezzo_giorno = mezzo.prezzo_giorno || prezzoCategoria(categoria);
    mezzo.km_inclusi = mezzo.km_inclusi || kmCategoria(categoria);
    mezzo.alert_km = mezzo.alert_km || 1000;
    mezzo.alert_giorni = mezzo.alert_giorni || 30;

    count++;
  });

  saveDb(db);

  try {
    fs.unlinkSync(req.file.path);
  } catch {}

  res.send(layout('Import completato', `
    <div class="card">
      <h2 class="ok">Import completato</h2>
      <p>Mezzi importati/aggiornati: <b>${count}</b></p>
      <a class="btn" href="/mezzi">Vai ai mezzi</a>
    </div>
  `));
});

app.get('/mezzi', (req, res) => {
  const db = loadDb();

  const rows = db.mezzi.map(m => `
    <tr>
      <td><a href="/mezzo/${m.id}"><b>${esc(m.targa)}</b></a></td>
      <td>${esc(m.descrizione)}</td>
      <td>${esc(m.categoria)}</td>
      <td>${esc(m.km)}</td>
      <td>€ ${euro(m.prezzo_giorno)}</td>
      <td>${esc(m.km_inclusi)}</td>
      <td>${esc(m.tagliando_km)}</td>
      <td>${esc(m.revisione)}</td>
      <td>${esc(m.bollo)}</td>
      <td>${checkAlerts(m).length ? '<span class="bad">ALERT</span>' : '<span class="ok">OK</span>'}</td>
    </tr>
  `).join('');

  res.send(layout('Mezzi', `
    <div class="card">
      <h2>Nuovo mezzo</h2>
      <form method="POST" action="/mezzi">
        <div class="grid">
          <input name="targa" placeholder="Targa" required>
          <input name="descrizione" placeholder="Descrizione cliente es. Opel Vivaro - pulmino 9 posti" required>

          <select name="categoria">
            <option>FURGONE</option>
            <option>9_POSTI</option>
            <option>AUTO_DACIA</option>
            <option>AUTO_GOLF</option>
            <option>ESCAVATORE</option>
            <option>SEMOVENTE</option>
          </select>

          <input name="km" type="number" placeholder="Km attuali">
          <input name="prezzo_giorno" type="number" step="0.01" placeholder="Prezzo giorno" value="70">
          <input name="km_inclusi" type="number" placeholder="Km inclusi giorno" value="150">
          <input name="tagliando_km" type="number" placeholder="Scadenza tagliando km">
          <input name="revisione" type="date">
          <input name="bollo" type="date">
          <input name="assicurazione" type="date">
          <input name="gomme" type="date">
        </div>
        <button>Salva mezzo</button>
      </form>
    </div>

    <table>
      <tr>
        <th>Targa</th>
        <th>Descrizione</th>
        <th>Categoria</th>
        <th>Km</th>
        <th>Prezzo</th>
        <th>Km/g</th>
        <th>Tagliando</th>
        <th>Revisione</th>
        <th>Bollo</th>
        <th>Alert</th>
      </tr>
      ${rows}
    </table>
  `));
});

app.post('/mezzi', (req, res) => {
  const db = loadDb();

  db.counters.mezzi++;

  db.mezzi.push({
    id: db.counters.mezzi,
    ...req.body,
    alert_km: 1000,
    alert_giorni: 30
  });

  saveDb(db);
  res.redirect('/mezzi');
});

app.get('/mezzo/:id', (req, res) => {
  const db = loadDb();
  const m = db.mezzi.find(x => String(x.id) === String(req.params.id));

  if (!m) return res.send('Mezzo non trovato');

  const files = db.allegati.filter(a => String(a.mezzo_id) === String(m.id));
  const alerts = checkAlerts(m).map(a => `<div class="alert">${esc(a)}</div>`).join('');

  res.send(layout('Scheda mezzo', `
    <div class="card">
      <h2>${esc(m.targa)} - ${esc(m.descrizione)}</h2>
      ${alerts || '<p class="ok">Nessun alert.</p>'}

      <form method="POST" action="/mezzo/${m.id}">
        <div class="grid">
          <input name="targa" value="${esc(m.targa)}">
          <input name="descrizione" value="${esc(m.descrizione)}">
          <input name="categoria" value="${esc(m.categoria)}">
          <input name="km" type="number" value="${esc(m.km)}">
          <input name="prezzo_giorno" type="number" step="0.01" value="${esc(m.prezzo_giorno)}">
          <input name="km_inclusi" type="number" value="${esc(m.km_inclusi)}">
          <input name="tagliando_km" type="number" value="${esc(m.tagliando_km)}">
          <input name="revisione" type="date" value="${esc(m.revisione)}">
          <input name="bollo" type="date" value="${esc(m.bollo)}">
          <input name="assicurazione" type="date" value="${esc(m.assicurazione)}">
          <input name="gomme" type="date" value="${esc(m.gomme)}">
        </div>
        <button>Salva scheda mezzo</button>
      </form>

      <a class="btn" href="/nuova?mezzo=${m.id}">Prenota questo mezzo</a>
    </div>

    <div class="card">
      <h3>Foto/documenti mezzo</h3>
      <form method="POST" action="/mezzo/${m.id}/foto" enctype="multipart/form-data">
        <select name="tipo">
          <option>Foto mezzo</option>
          <option>Libretto</option>
          <option>Assicurazione</option>
          <option>Revisione</option>
          <option>Bollo</option>
          <option>Manutenzione</option>
        </select>
        <input type="file" name="file" accept="image/*,.pdf" required>
        <button>Carica</button>
      </form>

      <ul>
        ${files.map(f => `
          <li>
            ${esc(f.tipo)} -
            <a href="/uploads/${esc(f.filename)}">${esc(f.originalname)}</a>
            ${f.drive_link ? ` - <a target="_blank" href="${esc(f.drive_link)}">Drive</a>` : ''}
          </li>
        `).join('')}
      </ul>
    </div>
  `));
});

app.post('/mezzo/:id', (req, res) => {
  const db = loadDb();
  const m = db.mezzi.find(x => String(x.id) === String(req.params.id));

  if (!m) return res.send('Mezzo non trovato');

  Object.assign(m, req.body);
  saveDb(db);

  res.redirect('/mezzo/' + m.id);
});

app.post('/mezzo/:id/foto', upload.single('file'), async (req, res) => {
  const db = loadDb();
  const m = db.mezzi.find(x => String(x.id) === String(req.params.id));

  if (!m) return res.send('Mezzo non trovato');

  if (!req.file) {
    return res.send(layout('Errore upload', `
      <div class="card">
        <h2 class="bad">Nessun file selezionato</h2>
        <a class="btn" href="/mezzo/${m.id}">Torna al mezzo</a>
      </div>
    `));
  }

  let drive = null;

  try {
    drive = await driveUpload(
      req.file.path,
      `${Date.now()}_${req.body.tipo}_${req.file.originalname}`,
      req.file.mimetype,
      `MEZZO ${m.targa}`
    );
  } catch (err) {
    console.log('Errore Drive mezzo:', err.message);
  }

  db.counters.allegati++;

  db.allegati.push({
    id: db.counters.allegati,
    mezzo_id: m.id,
    tipo: req.body.tipo,
    filename: req.file.filename,
    originalname: req.file.originalname,
    drive_link: drive ? drive.link : '',
    created_at: nowIt()
  });

  saveDb(db);
  res.redirect('/mezzo/' + m.id);
});
function formPrenotazione(action, mezzoId = '', data = '') {
  const db = loadDb();

  const opts = db.mezzi.map(m => `
    <option value="${m.id}" ${String(m.id) === String(mezzoId) ? 'selected' : ''}>
      ${esc(m.descrizione)} (${esc(m.targa)})
    </option>
  `).join('');

  return `
    <form method="POST" action="${action}">
      <div class="grid">
        <input name="cliente" placeholder="Cliente" required>
        <input name="telefono" placeholder="Telefono" required>
        <input name="email" type="email" placeholder="Email" required>
        <input name="codice_fiscale" placeholder="Codice fiscale" required>
        <input name="indirizzo" placeholder="Indirizzo" required>
        <select name="fatturazione">
          <option>Privato</option>
          <option>Azienda</option>
        </select>

        <input name="piva" placeholder="P.IVA">
        <input name="ragione_sociale" placeholder="Ragione sociale">
        <input name="pec" placeholder="PEC">
        <input name="sdi" placeholder="SDI">

        <input name="conducente1" placeholder="Conducente 1" required>
        <input name="patente1" placeholder="Patente 1" required>
        <input type="date" name="patente1_scadenza" required>
        <input name="conducente2" placeholder="Conducente 2">
        <input name="patente2" placeholder="Patente 2">
        <input type="date" name="patente2_scadenza">

        <select name="mezzo_id" required>${opts}</select>

        <input type="date" name="data_inizio" value="${esc(data)}" required>
        <input type="time" name="ora_inizio" value="08:30">

        <input type="date" name="data_fine" value="${esc(data)}" required>
        <input type="time" name="ora_fine" value="18:00">

        <input type="number" name="km_previsti" value="150">

        <select name="carburante_uscita">
          ${fuelOptions('4/4 pieno')}
        </select>
      </div>

      <textarea name="note" placeholder="Note"></textarea>

      <button>Crea contratto</button>
    </form>
  `;
}

app.get('/nuova', (req, res) => {
  res.send(layout('Nuova prenotazione', `
    <div class="card">
      <h2>Nuova prenotazione</h2>
      ${formPrenotazione('/nuova', req.query.mezzo, req.query.data)}
    </div>
  `));
});

app.post('/nuova', (req, res) => {
  const db = loadDb();

  const erroreDate = validaDate(req.body.data_inizio, req.body.data_fine);

  if (erroreDate) {
    return res.send(layout('Errore date', `
      <div class="card">
        <h2 class="bad">${esc(erroreDate)}</h2>
        <a class="btn" href="/nuova">Torna alla prenotazione</a>
      </div>
    `));
  }

  const mezzo = db.mezzi.find(m => String(m.id) === String(req.body.mezzo_id));

  if (!mezzo) {
    return res.send(layout('Errore', `
      <div class="card">
        <h2 class="bad">Mezzo mancante</h2>
        <a class="btn" href="/mezzi">Vai ai mezzi</a>
      </div>
    `));
  }

  const occupato = db.prenotazioni.find(p =>
    p.stato !== 'annullato' &&
    String(p.mezzo_id) === String(req.body.mezzo_id) &&
    p.data_inizio <= req.body.data_fine &&
    p.data_fine >= req.body.data_inizio
  );

  if (occupato) {
    return res.send(layout('Mezzo occupato', `
      <div class="card">
        <h2 class="bad">Mezzo occupato in queste date</h2>
        <p>Contratto: <a href="/prenotazione/${occupato.id}">${esc(occupato.codice)}</a></p>
        <a class="btn" href="/planning">Vai al planning</a>
      </div>
    `));
  }

  const c = calcolaTotale(req.body, mezzo);

  db.counters.prenotazioni++;
  const id = db.counters.prenotazioni;

  db.prenotazioni.push({
    id,
    codice: `DPR-${todayCode()}-${String(id).padStart(4, '0')}`,
    stato: 'bozza',
    created_at: nowIt(),
    ...req.body,
    mezzo_id: Number(req.body.mezzo_id),
    giorni: c.giorni,
    km_inclusi: c.kmInclusi,
    extra_km: c.extraKm,
    extra_fuori_orario: c.extraFuoriOrario,
    imponibile: c.imponibile,
    iva: c.iva,
    totale: c.totale,
    cauzione: CAUZIONE,
    carburante_rientro: '4/4 pieno',
    firma_path: ''
  });

  saveDb(db);

  res.redirect('/prenotazione/' + id);
});

app.get('/cliente', (req, res) => {
  res.send(layout('Pagina cliente', `
    <div class="card">
      <h2>Richiesta cliente DP RENT</h2>
      <p class="notice">La targa resta interna. Il cliente vede solo la descrizione del mezzo.</p>
      ${formPrenotazione('/nuova')}
    </div>
  `));
});

app.get('/prenotazione/:id', (req, res) => {
  const db = loadDb();

  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};

  res.send(layout('Contratto', `
    <div class="card">
      <h2>${esc(p.codice)}</h2>

      <p><b>Cliente:</b> ${esc(p.cliente)} - ${esc(p.telefono)}</p>
      <p><b>Email:</b> ${esc(p.email)}</p>
      <p><b>CF:</b> ${esc(p.codice_fiscale)}</p>
      <p><b>Indirizzo:</b> ${esc(p.indirizzo)}</p>

      <p><b>Fatturazione:</b>
        ${esc(p.fatturazione)}
        ${esc(p.ragione_sociale)}
        ${esc(p.piva)}
        ${esc(p.pec)}
        ${esc(p.sdi)}
      </p>

      <p><b>Conducente 1:</b>
        ${esc(p.conducente1)} - patente ${esc(p.patente1)}
        scad. ${formatDate(p.patente1_scadenza)}
      </p>

      <p><b>Conducente 2:</b>
        ${esc(p.conducente2)} - patente ${esc(p.patente2)}
        scad. ${formatDate(p.patente2_scadenza)}
      </p>

      <p><b>Mezzo:</b> ${esc(m.targa)} - ${esc(m.descrizione)}</p>

      <p><b>Periodo:</b>
        ${formatDate(p.data_inizio)} ${esc(p.ora_inizio)}
        →
        ${formatDate(p.data_fine)} ${esc(p.ora_fine)}
      </p>

      <p><b>Giorni:</b> ${esc(p.giorni)}</p>
      <p><b>Km inclusi:</b> ${esc(p.km_inclusi)}</p>
      <p><b>Km previsti:</b> ${esc(p.km_previsti)}</p>
      <p><b>Extra fuori orario:</b> € ${euro(p.extra_fuori_orario)} + IVA</p>
      <p><b>Extra km:</b> € ${euro(p.extra_km)} + IVA</p>
      <p><b>Totale IVA inclusa:</b> € ${euro(p.totale)}</p>
      <p><b>Stato:</b> ${esc(p.stato)}</p>

      ${p.pdf_drive ? `<p><b>PDF Drive:</b> <a target="_blank" href="${esc(p.pdf_drive)}">Apri PDF</a></p>` : ''}
      ${p.firma_path ? `<p class="ok"><b>Firma cliente salvata</b></p>` : `<p class="warn"><b>Firma cliente non presente</b></p>`}

      <a class="btn" href="/pdf/${p.id}">Scarica PDF</a>
      <a class="btn btn2" href="/email/${p.id}">Invia email</a>
      <a class="btn btn2" href="/foto/${p.id}">Foto/documenti</a>
      <a class="btn btn3" href="/firma/${p.id}">Firma contratto</a>
      <a class="btn btn3" href="/firma-link/${p.id}">Link firma cliente</a>
      <a class="btn btn3" href="/checkout/${p.id}">Check-out</a>
      <a class="btn btn3" href="/checkin/${p.id}">Check-in</a>
      <a class="btn btn2" href="/nexi/${p.id}">Nexi</a>
    </div>
  `));
});

app.get('/planning', (req, res) => {
  const db = loadDb();

  const now = new Date();
  const y = Number(req.query.y || now.getFullYear());
  const mo = Number(req.query.m || now.getMonth() + 1);
  const days = new Date(y, mo, 0).getDate();

  const prev = new Date(y, mo - 2, 1);
  const next = new Date(y, mo, 1);

  let head = '<th class="first">Mezzo</th>';

  for (let d = 1; d <= days; d++) {
    head += `<th>${d}</th>`;
  }

  const rows = db.mezzi.map(m => {
    let r = `<tr>
      <td class="first">
        <b>${esc(m.targa)}</b><br>
        ${esc(m.descrizione)}
      </td>
    `;

    for (let d = 1; d <= days; d++) {
      const date = `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;

      const occ = db.prenotazioni.find(p =>
        String(p.mezzo_id) === String(m.id) &&
        p.stato !== 'annullato' &&
        p.data_inizio <= date &&
        p.data_fine >= date
      );

      if (occ) {
        r += `
          <td class="busy"
              title="${esc(occ.codice)} ${esc(occ.cliente)}"
              onclick="location.href='/prenotazione/${occ.id}'">
            O
          </td>
        `;
      } else {
        r += `
          <td class="free"
              title="Libero"
              onclick="location.href='/nuova?mezzo=${m.id}&data=${date}'">
            L
          </td>
        `;
      }
    }

    return r + '</tr>';
  }).join('');

  res.send(layout('Planning', `
    <h2>Planning ${mo}/${y}</h2>

    <p>
      <a href="/planning?y=${prev.getFullYear()}&m=${prev.getMonth() + 1}">← Mese precedente</a>
      |
      <a href="/planning?y=${next.getFullYear()}&m=${next.getMonth() + 1}">Mese successivo →</a>
    </p>

    <div class="sticky">
      <table>
        <tr>${head}</tr>
        ${rows}
      </table>
    </div>
  `));
});
function generaPdf(p, m) {
  const file = path.join(CONTRACT_DIR, `contratto_${p.codice}.pdf`);
  const doc = new PDFDocument({ size: 'A4', margin: 28 });
  const stream = fs.createWriteStream(file);

  doc.pipe(stream);

  doc.rect(0, 0, 595, 86).fill('#111');
  doc.fillColor('white').fontSize(28).text('DP RENT', 34, 24);
  doc.fontSize(9).text(
    `${AZIENDA.nome}\n${AZIENDA.indirizzo}\nP.IVA ${AZIENDA.piva} | Tel. ${AZIENDA.telefono}\n${AZIENDA.email}`,
    335,
    16,
    { width: 220, align: 'right' }
  );
  doc.rect(0, 86, 595, 5).fill('#c40000');

  doc.fillColor('black').fontSize(17).text('CONTRATTO DI NOLEGGIO', 28, 105, {
    width: 540,
    align: 'center'
  });

  let y = 138;

  function section(title) {
    doc.rect(28, y, 540, 17).fill('#111');
    doc.fillColor('white').fontSize(8).text(title, 35, y + 5);
    doc.fillColor('black');
    y += 22;
  }

  function row(label, value, x = 34, w = 245) {
    doc.fillColor('#666').fontSize(7).text(label, x, y, { width: 78 });
    doc.fillColor('#111').fontSize(8).text(String(value || ''), x + 82, y, { width: w - 82 });
  }

  function row2(label1, value1, label2, value2) {
    row(label1, value1, 34, 255);
    row(label2, value2, 305, 245);
    y += 13;
  }

  section('DATI CONTRATTO');
  row2('Numero', p.codice, 'Stato', p.stato);
  row2('Creato il', p.created_at, 'Cauzione', `€ ${euro(p.cauzione)}`);

  section('CLIENTE E FATTURAZIONE');
  row2('Cliente', p.cliente, 'Telefono', p.telefono);
  row2('Email', p.email, 'Codice fiscale', p.codice_fiscale);
  row2('Indirizzo', p.indirizzo, 'Fatturazione', p.fatturazione);
  row2('Ragione sociale', p.ragione_sociale, 'P.IVA', p.piva);
  row2('PEC', p.pec, 'SDI', p.sdi);

  section('CONDUCENTI');
  row2('Conducente 1', p.conducente1, 'Patente 1', `${p.patente1 || ''} scad. ${formatDate(p.patente1_scadenza)}`);
  row2('Conducente 2', p.conducente2, 'Patente 2', `${p.patente2 || ''} scad. ${formatDate(p.patente2_scadenza)}`);

  section('VEICOLO E NOLEGGIO');
  row2('Targa', m.targa, 'Mezzo', m.descrizione);
  row2('Check-out', `${formatDate(p.data_inizio)} ${p.ora_inizio || ''}`, 'Check-in', `${formatDate(p.data_fine)} ${p.ora_fine || ''}`);
  row2('Giorni', p.giorni, 'Km inclusi', p.km_inclusi);
  row2('Km previsti', p.km_previsti, 'Km uscita/rientro', `${p.km_uscita || ''} / ${p.km_rientro || ''}`);
  row2('Carburante', `${p.carburante_uscita || ''} / ${p.carburante_rientro || ''}`, 'Note', p.note || '');

  section('RIEPILOGO ECONOMICO');
  row2('Extra fuori orario', `€ ${euro(p.extra_fuori_orario)} + IVA`, 'Extra km', `€ ${euro(p.extra_km)} + IVA`);
  row2('Imponibile', `€ ${euro(p.imponibile)}`, 'IVA 22%', `€ ${euro(p.iva)}`);
  row2('Totale IVA inclusa', `€ ${euro(p.totale)}`, 'Deposito cauzionale', `€ ${euro(p.cauzione)}`);

  section('CONDIZIONI E PRIVACY');
  doc.fontSize(7).text(
  'Il cliente dichiara di aver preso visione e accettare le condizioni generali di noleggio DP RENT / Trasporti DP S.R.L. e l’informativa privacy. Il mezzo viene consegnato nello stato indicato al check-out e deve essere riconsegnato nelle stesse condizioni, con livello carburante equivalente. Eventuali danni, multe, pedaggi, franchigie, smarrimento chiavi/documenti, ritardi di riconsegna, pulizia straordinaria e costi accessori sono a carico del cliente. La cauzione resta vincolata fino al controllo finale del mezzo.',
  34,
  y,
  { width: 520 }
);
  y += 9;
  doc.text(`Privacy: ${PRIVACY_URL}`, 34, y, { width: 520 });
  y += 11;
  doc.fontSize(7).text(
    'Il cliente dichiara di aver preso visione e accettare condizioni generali e informativa privacy. Il veicolo viene consegnato con carburante indicato nel contratto e deve essere riconsegnato allo stesso livello. Danni, multe, pedaggi, franchigie, smarrimenti e costi accessori sono a carico del cliente.',
    34,
    y,
    { width: 520 }
  );
  y += 44;

  doc.fontSize(9).fillColor('#111').text('Firma cliente:', 34, y);
  doc.text('Firma DP RENT:', 315, y);

  if (p.firma_path && fs.existsSync(p.firma_path)) {
    try {
      doc.image(p.firma_path, 34, y + 13, { fit: [210, 50] });
    } catch (err) {
      doc.text('____________________________', 34, y + 28);
    }
  } else {
    doc.text('____________________________', 34, y + 28);
  }

  doc.text('____________________________', 315, y + 28);

  doc.end();

  return new Promise((resolve, reject) => {
    stream.on('finish', async () => {
      try {
        const uploaded = await driveUpload(
          file,
          path.basename(file),
          'application/pdf',
          `${p.codice} - ${p.cliente || 'cliente'}`
        );

        if (uploaded) {
          p.pdf_drive = uploaded.link;
        }
      } catch (err) {
        console.log('Errore Drive PDF:', err.message);
      }

      resolve(file);
    });

    stream.on('error', reject);
  });
}

app.get('/pdf/:id', async (req, res) => {
  try {
    const db = loadDb();
    const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

    if (!p) return res.send('Contratto non trovato');

    const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};
    const file = await generaPdf(p, m);

    saveDb(db);

    res.download(file);
  } catch (err) {
    res.send(layout('Errore PDF', `
      <div class="card">
        <h2 class="bad">Errore PDF</h2>
        <pre>${esc(err.message)}</pre>
      </div>
    `));
  }
});

app.get('/email/:id', async (req, res) => {
  try {
    const db = loadDb();
    const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

    if (!p) return res.send('Contratto non trovato');

    const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};
    const file = await generaPdf(p, m);

    saveDb(db);

    await sendEmail(
      p.email || process.env.ALERT_EMAIL,
      `Contratto DP RENT ${p.codice}`,
      `Buongiorno,\nin allegato il contratto DP RENT ${p.codice}.\n\nDP RENT`,
      [{ filename: path.basename(file), path: file }]
    );

    res.send(layout('Email inviata', `
      <div class="card">
        <h2 class="ok">Email inviata</h2>
        <a class="btn" href="/prenotazione/${p.id}">Torna al contratto</a>
      </div>
    `));
  } catch (err) {
    res.send(layout('Errore email', `
      <div class="card">
        <h2 class="bad">Errore email</h2>
        <pre>${esc(err.message)}</pre>
      </div>
    `));
  }
});

app.get('/firma/:id', (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  res.send(layout('Firma contratto', `
    <div class="card">
      <h2>Firma contratto ${esc(p.codice)}</h2>
      <p>Firma con dito, penna tablet o mouse.</p>

      <canvas id="firma"></canvas>

      <button type="button" onclick="pulisci()">Cancella</button>
      <button type="button" onclick="salva()">Salva firma</button>

      <a class="btn btn2" href="/prenotazione/${p.id}">Torna al contratto</a>
    </div>

    <script>
      const canvas = document.getElementById('firma');
      const ctx = canvas.getContext('2d');

      function resizeCanvas() {
        const rect = canvas.getBoundingClientRect();
        const data = canvas.toDataURL();
        canvas.width = rect.width;
        canvas.height = 260;
        const img = new Image();
        img.onload = () => ctx.drawImage(img, 0, 0);
        img.src = data;
        ctx.lineWidth = 3;
        ctx.lineCap = 'round';
      }

      resizeCanvas();
      window.addEventListener('resize', resizeCanvas);

      let drawing = false;

      function getPos(e) {
        const rect = canvas.getBoundingClientRect();
        const t = e.touches ? e.touches[0] : e;
        return {
          x: t.clientX - rect.left,
          y: t.clientY - rect.top
        };
      }

      function start(e) {
        drawing = true;
        const p = getPos(e);
        ctx.beginPath();
        ctx.moveTo(p.x, p.y);
        e.preventDefault();
      }

      function move(e) {
        if (!drawing) return;
        const p = getPos(e);
        ctx.lineTo(p.x, p.y);
        ctx.stroke();
        e.preventDefault();
      }

      function stop(e) {
        drawing = false;
        e.preventDefault();
      }

      canvas.addEventListener('mousedown', start);
      canvas.addEventListener('mousemove', move);
      canvas.addEventListener('mouseup', stop);
      canvas.addEventListener('mouseleave', stop);

      canvas.addEventListener('touchstart', start);
      canvas.addEventListener('touchmove', move);
      canvas.addEventListener('touchend', stop);

      function pulisci() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);
      }

      async function salva() {
        const firma = canvas.toDataURL('image/png');

        const r = await fetch('/firma/${p.id}', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ firma })
        });

        if (r.ok) {
          location.href = '/prenotazione/${p.id}';
        } else {
          alert('Errore salvataggio firma');
        }
      }
    </script>
  `));
});

app.post('/firma/:id', (req, res) => {
  try {
    const db = loadDb();
    const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

    if (!p) return res.status(404).send('Contratto non trovato');
    if (!req.body.firma) return res.status(400).send('Firma mancante');

    const base64 = req.body.firma.replace(/^data:image\/png;base64,/, '');
    const filePath = path.join(UPLOAD_DIR, `firma_${p.id}.png`);

    fs.writeFileSync(filePath, base64, 'base64');

    p.firma_path = filePath;
    p.stato = 'firmato';

    saveDb(db);

    res.send('OK');
  } catch (err) {
    res.status(500).send(err.message);
  }
});

app.get('/firma-link/:id', (req, res) => {
  const link = `${req.protocol}://${req.get('host')}/firma/${req.params.id}`;

  res.send(layout('Link firma cliente', `
    <div class="card">
      <h2>Link firma cliente</h2>
      <input value="${esc(link)}" readonly onclick="this.select()">
      <p>Invia questo link al cliente via WhatsApp o email.</p>
      <a class="btn" target="_blank" href="https://wa.me/?text=${encodeURIComponent('Firma contratto DP RENT: ' + link)}">Invia su WhatsApp</a>
      <a class="btn btn2" href="/prenotazione/${req.params.id}">Torna al contratto</a>
    </div>
  `));
});
app.get('/foto/:id', (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));
  if (!p) return res.send('Contratto non trovato');

  const files = db.allegati.filter(a => String(a.prenotazione_id) === String(p.id));

  const options = `
    <option>Patente conducente 1</option>
    <option>Documento conducente 1</option>
    <option>Patente conducente 2</option>
    <option>Documento conducente 2</option>
    <option>Foto uscita fronte</option>
    <option>Foto uscita retro</option>
    <option>Foto uscita lato destro</option>
    <option>Foto uscita lato sinistro</option>
    <option>Foto uscita interno</option>
    <option>Foto rientro fronte</option>
    <option>Foto rientro retro</option>
    <option>Foto rientro lato destro</option>
    <option>Foto rientro lato sinistro</option>
    <option>Foto rientro interno</option>
    <option>Danno</option>
    <option>File generico</option>
  `;

  res.send(layout('Foto / documenti', `
    <div class="card">
      <h2>Foto / documenti ${esc(p.codice)}</h2>

      <label><b>Tipo documento / foto</b></label>
      <select id="tipoScelto">${options}</select>

      <form id="formFoto" method="POST" enctype="multipart/form-data">
        <input type="hidden" name="tipo" id="tipoFoto">
        <input id="cameraInput" type="file" name="file" accept="image/*" capture="environment" style="display:none" required>
      </form>

      <form id="formFile" method="POST" enctype="multipart/form-data">
        <input type="hidden" name="tipo" id="tipoFile">
        <input id="fileInput" type="file" name="file" accept="image/*,.pdf" style="display:none" required>
      </form>

      <button type="button" onclick="scattaFoto()">📸 Scatta foto</button>
      <button type="button" onclick="caricaFile()">📁 Carica da dispositivo</button>

      <br><br>
      <a class="btn btn2" href="/prenotazione/${p.id}">Torna al contratto</a>

      <h3>File caricati</h3>
      <ul>
        ${files.map(f => `
          <li>
            ${esc(f.tipo)} -
            <a href="/uploads/${esc(f.filename)}" target="_blank">${esc(f.originalname)}</a>
            ${f.drive_link ? ` - <a target="_blank" href="${esc(f.drive_link)}">Drive</a>` : ''}
          </li>
        `).join('')}
      </ul>
    </div>

    <script>
      function tipo() {
        return document.getElementById('tipoScelto').value;
      }

      function scattaFoto() {
        document.getElementById('tipoFoto').value = tipo();
        document.getElementById('cameraInput').click();
      }

      function caricaFile() {
        document.getElementById('tipoFile').value = tipo();
        document.getElementById('fileInput').click();
      }

      document.getElementById('cameraInput').addEventListener('change', function () {
        if (this.files.length) document.getElementById('formFoto').submit();
      });

      document.getElementById('fileInput').addEventListener('change', function () {
        if (this.files.length) document.getElementById('formFile').submit();
      });
    </script>
  `));
});
app.post('/foto/:id', upload.single('file'), async (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  if (!req.file) {
    return res.send(layout('Errore upload', `
      <div class="card">
        <h2 class="bad">Nessun file selezionato</h2>
        <p>Da Mac devi scegliere una foto/file. Da telefono o tablet può aprire la fotocamera.</p>
        <a class="btn" href="/foto/${p.id}">Torna a foto/documenti</a>
        <a class="btn btn2" href="/prenotazione/${p.id}">Torna al contratto</a>
      </div>
    `));
  }

  let uploaded = null;

  try {
    uploaded = await driveUpload(
      req.file.path,
      `${Date.now()}_${req.body.tipo}_${req.file.originalname}`,
      req.file.mimetype,
      `${p.codice} - ${p.cliente || 'cliente'}`
    );
  } catch (err) {
    console.log('Errore Drive foto:', err.message);
  }

  db.counters.allegati++;

  db.allegati.push({
    id: db.counters.allegati,
    prenotazione_id: p.id,
    tipo: req.body.tipo,
    filename: req.file.filename,
    originalname: req.file.originalname,
    drive_link: uploaded ? uploaded.link : '',
    created_at: nowIt()
  });

  saveDb(db);

  res.redirect('/foto/' + p.id);
});

app.get('/checkout/:id', (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  res.send(layout('Check-out', `
    <div class="card">
      <h2>Check-out ${esc(p.codice)}</h2>

      <form method="POST">
        <label>Km uscita</label>
        <input type="number" name="km_uscita" value="${esc(p.km_uscita)}">

        <label>Carburante uscita</label>
        <select name="carburante_uscita">${fuelOptions(p.carburante_uscita)}</select>

        <label>Note uscita</label>
        <textarea name="note_uscita">${esc(p.note_uscita)}</textarea>

        <button>Salva check-out</button>
      </form>

      <a class="btn btn2" href="/foto/${p.id}">Carica foto uscita</a>
      <a class="btn btn2" href="/prenotazione/${p.id}">Torna al contratto</a>
    </div>
  `));
});

app.post('/checkout/:id', (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  p.km_uscita = Number(req.body.km_uscita || 0);
  p.carburante_uscita = req.body.carburante_uscita;
  p.note_uscita = req.body.note_uscita;
  p.stato = 'in_corso';

  const m = db.mezzi.find(x => x.id === p.mezzo_id);
  if (m && p.km_uscita) m.km = p.km_uscita;

  saveDb(db);

  res.redirect('/prenotazione/' + p.id);
});

app.get('/checkin/:id', (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  res.send(layout('Check-in / Rientro', `
    <div class="card">
      <h2>Check-in / Rientro ${esc(p.codice)}</h2>

      <form method="POST">
        <label>Km rientro</label>
        <input type="number" name="km_rientro" value="${esc(p.km_rientro)}" required>

        <label>Carburante rientro</label>
        <select name="carburante_rientro">${fuelOptions(p.carburante_rientro)}</select>

        <label>Note rientro / danni</label>
        <textarea name="note_rientro">${esc(p.note_rientro)}</textarea>

        <button>Chiudi rientro</button>
      </form>

      <a class="btn btn2" href="/foto/${p.id}">Carica foto rientro</a>
      <a class="btn btn2" href="/prenotazione/${p.id}">Torna al contratto</a>
    </div>
  `));
});

app.post('/checkin/:id', async (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  p.km_rientro = Number(req.body.km_rientro || 0);
  p.carburante_rientro = req.body.carburante_rientro;
  p.note_rientro = req.body.note_rientro;
  p.stato = 'rientrato';

  const m = db.mezzi.find(x => x.id === p.mezzo_id);

  if (m && p.km_rientro) {
    m.km = p.km_rientro;
  }

  saveDb(db);

  if (m) {
    const alerts = checkAlerts(m);

    if (alerts.length) {
      await sendAlert(
        `ALERT DP RENT ${m.targa}`,
        `Mezzo: ${m.targa} ${m.descrizione}\nKm rientro: ${p.km_rientro}\n\n${alerts.join('\n')}`
      );
    }
  }

  res.redirect('/prenotazione/' + p.id);
});

app.get('/rientri', (req, res) => {
  const db = loadDb();
  const q = String(req.query.q || '').toLowerCase();

  let list = db.prenotazioni.filter(p =>
    p.stato !== 'rientrato' &&
    p.stato !== 'annullato'
  );

  if (q) {
    list = list.filter(p => {
      const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};
      return `${p.codice} ${p.cliente} ${p.telefono} ${m.targa} ${m.descrizione}`
        .toLowerCase()
        .includes(q);
    });
  }

  const rows = list.map(p => {
    const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};

    return `
      <tr>
        <td><a href="/prenotazione/${p.id}">${esc(p.codice)}</a></td>
        <td>${esc(p.cliente)}<br>${esc(p.telefono)}</td>
        <td>${esc(m.targa)}<br>${esc(m.descrizione)}</td>
        <td>${formatDate(p.data_inizio)} → ${formatDate(p.data_fine)}</td>
        <td>${esc(p.stato)}</td>
        <td><a class="btn" href="/checkin/${p.id}">Fai check-in</a></td>
      </tr>
    `;
  }).join('');

  res.send(layout('Rientri / Check-in', `
    <div class="card">
      <h2>Rientri / Check-in</h2>
      <form>
        <input name="q" placeholder="Cerca targa, cliente, telefono, codice" value="${esc(req.query.q || '')}">
        <button>Cerca</button>
      </form>
    </div>

    <table>
      <tr>
        <th>Contratto</th>
        <th>Cliente</th>
        <th>Mezzo</th>
        <th>Date</th>
        <th>Stato</th>
        <th>Azione</th>
      </tr>
      ${rows}
    </table>
  `));
});

app.get('/storico', (req, res) => {
  const db = loadDb();
  const q = String(req.query.q || '').toLowerCase();

  let list = db.prenotazioni.slice().reverse();

  if (q) {
    list = list.filter(p => {
      const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};
      return `${p.codice} ${p.cliente} ${p.telefono} ${m.targa} ${m.descrizione}`
        .toLowerCase()
        .includes(q);
    });
  }

  const rows = list.map(p => {
    const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};

    return `
      <tr>
        <td><a href="/prenotazione/${p.id}">${esc(p.codice)}</a></td>
        <td>${esc(p.cliente)}</td>
        <td>${esc(p.telefono)}</td>
        <td>${esc(m.targa)}<br>${esc(m.descrizione)}</td>
        <td>${formatDate(p.data_inizio)} → ${formatDate(p.data_fine)}</td>
        <td>€ ${euro(p.totale)}</td>
        <td>${esc(p.stato)}</td>
      </tr>
    `;
  }).join('');

  res.send(layout('Storico', `
    <div class="card">
      <h2>Storico contratti</h2>
      <form>
        <input name="q" placeholder="Cerca nome, targa, telefono, codice" value="${esc(req.query.q || '')}">
        <button>Cerca</button>
      </form>
    </div>

    <table>
      <tr>
        <th>Contratto</th>
        <th>Cliente</th>
        <th>Telefono</th>
        <th>Mezzo</th>
        <th>Date</th>
        <th>Totale</th>
        <th>Stato</th>
      </tr>
      ${rows}
    </table>
  `));
});

app.get('/scadenze', (req, res) => {
  const db = loadDb();

  const rows = db.mezzi.map(m => `
    <tr>
      <td><a href="/mezzo/${m.id}">${esc(m.targa)}</a></td>
      <td>${esc(m.descrizione)}</td>
      <td>${esc(m.km)}</td>
      <td>${esc(m.tagliando_km)}</td>
      <td>${formatDate(m.revisione)}</td>
      <td>${formatDate(m.bollo)}</td>
      <td>${formatDate(m.assicurazione)}</td>
      <td>${checkAlerts(m).join('<br>') || '<span class="ok">OK</span>'}</td>
    </tr>
  `).join('');

  res.send(layout('Scadenze', `
    <h2>Scadenze</h2>

    <table>
      <tr>
        <th>Targa</th>
        <th>Mezzo</th>
        <th>Km</th>
        <th>Tagliando km</th>
        <th>Revisione</th>
        <th>Bollo</th>
        <th>Assicurazione</th>
        <th>Alert</th>
      </tr>
      ${rows}
    </table>
  `));
});

app.get('/nexi/:id', (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));

  if (!p) return res.send('Contratto non trovato');

  const info = nexiPlaceholderUrl(p);

  res.send(layout('Nexi', `
    <div class="card">
      <h2>Nexi predisposto</h2>

      <p><b>Configurazione:</b> ${info.ready ? 'Pronta' : 'Mancano NEXI_ALIAS / NEXI_MAC_KEY'}</p>
      <p><b>Contratto:</b> ${esc(p.codice)}</p>
      <p><b>Importo:</b> € ${euro(p.totale)}</p>
      <p><b>Codice transazione:</b> ${esc(info.transId)}</p>
      <p><b>MAC calcolato:</b> ${esc(info.mac)}</p>

      <p class="notice">
        Qui è predisposta la parte tecnica Nexi. Per pagamento reale va collegata la chiamata PayMail/API Nexi con alias e MAC key definitivi.
      </p>

      <a class="btn" href="/prenotazione/${p.id}">Torna contratto</a>
    </div>
  `));
});

app.get('/nexi-callback', (req, res) => {
  res.send('Callback Nexi ricevuto');
});

app.use((err, req, res, next) => {
  console.error(err);
  res.status(500).send(layout('Errore server', `
    <div class="card">
      <h2 class="bad">Errore server</h2>
      <pre>${esc(err.message)}</pre>
    </div>
  `));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log('DP RENT APP avviata su porta ' + PORT);
});
