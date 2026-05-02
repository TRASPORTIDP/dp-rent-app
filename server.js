require('dotenv').config();

const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 10000;

app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ extended: true, limit: '50mb' }));

const DATA_DIR = path.join(__dirname, 'data');
const UPLOAD_DIR = path.join(__dirname, 'uploads');
const CONTRACT_DIR = path.join(__dirname, 'contracts');
const DB_FILE = path.join(DATA_DIR, 'db.json');

[DATA_DIR, UPLOAD_DIR, CONTRACT_DIR].forEach(d => {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

app.use('/uploads', express.static(UPLOAD_DIR));
app.use('/contracts', express.static(CONTRACT_DIR));

const upload = multer({ dest: UPLOAD_DIR });

const AZIENDA = {
  nome: 'Trasporti DP S.R.L. - DP RENT',
  indirizzo: 'Via Tuderte 466, Narni (TR)',
  piva: '01385450554',
  telefono: '0744817108',
  email: 'contabilita@trasportidp.com'
};

const TERMS_URL = 'https://carrentalsoftware.myappy.it/data/public/user/65996976/terms_file.pdf';
const PRIVACY_URL = 'https://carrentalsoftware.myappy.it/data/public/user/65996976/privacy_file.pdf';

function dbDefault() {
  return { mezzi: [], prenotazioni: [], allegati: [], counters: { mezzi: 0, prenotazioni: 0, allegati: 0 } };
}

function loadDb() {
  if (!fs.existsSync(DB_FILE)) {
    const d = dbDefault();
    fs.writeFileSync(DB_FILE, JSON.stringify(d, null, 2));
    return d;
  }
  return JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
}

function saveDb(db) {
  fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2));
}

function esc(v) {
  return String(v || '').replace(/[&<>"']/g, s => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;'
  }[s]));
}

function euro(v) {
  return Number(v || 0).toFixed(2);
}

function todayCode() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}

function giorniNoleggio(a, b) {
  const d1 = new Date(a + 'T00:00:00');
  const d2 = new Date(b + 'T00:00:00');
  const diff = Math.floor((d2 - d1) / 86400000) + 1;
  return Math.max(1, diff || 1);
}

function extraOrario(ora) {
  if (!ora) return 0;
  const [h, m] = ora.split(':').map(Number);
  const min = h * 60 + (m || 0);
  return (min < 510 || min > 1110) ? 30 : 0; // prima 08:30 o dopo 18:30
}

function calcola(p, mezzo) {
  const giorni = giorniNoleggio(p.data_inizio, p.data_fine);
  const prezzo = Number(mezzo.prezzo_giorno || 70);
  const kmInclusi = Number(mezzo.km_inclusi || 150) * giorni;
  const kmPrevisti = Number(p.km_previsti || 0);
  const extraKm = Math.max(0, kmPrevisti - kmInclusi) * 0.15;
  const extraFuoriOrario = extraOrario(p.ora_inizio) + extraOrario(p.ora_fine);
  const imponibile = giorni * prezzo + extraKm + extraFuoriOrario;
  const iva = imponibile * 0.22;
  return { giorni, kmInclusi, extraKm, extraFuoriOrario, imponibile, iva, totale: imponibile + iva };
}

async function sendEmail(to, subject, text, attachments = []) {
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 465),
    secure: String(process.env.SMTP_SECURE || '').toLowerCase() === 'true',
    auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
    tls: { rejectUnauthorized: false }
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

function layout(title, body) {
  return `<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>${esc(title)}</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{margin:0;font-family:Arial;background:#f4f4f4;color:#111}
header{background:#070707;color:white;padding:22px;font-size:34px;font-weight:800}
nav{background:#c40000;padding:12px;display:flex;gap:10px;flex-wrap:wrap}
nav a{color:white;text-decoration:none;font-weight:bold;padding:8px}
main{padding:18px}
.card{background:white;border-radius:14px;padding:18px;margin:12px 0;box-shadow:0 2px 10px #ccc}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:15px}
.tile{background:#111;color:white;border-radius:18px;padding:30px;text-align:center;text-decoration:none;font-size:22px;font-weight:800;border-bottom:8px solid #c40000}
input,select,textarea,button{width:100%;box-sizing:border-box;padding:11px;margin:5px 0 12px;border:1px solid #aaa;border-radius:6px}
button,.btn{background:#c40000;color:white;border:0;border-radius:8px;padding:11px 16px;text-decoration:none;font-weight:bold;display:inline-block;width:auto}
.btn2{background:#222}.ok{color:green}.bad{color:#c40000}
table{width:100%;border-collapse:collapse;background:white}
th,td{border:1px solid #ddd;padding:7px;font-size:13px}
th{background:#111;color:white}
.free{background:#34b233;color:white;text-align:center;font-weight:bold;cursor:pointer}
.busy{background:#d60000;color:white;text-align:center;font-weight:bold;cursor:pointer}
@media(max-width:700px){header{font-size:26px}main{padding:8px}}
</style>
</head>
<body>
<header>DP RENT APP</header>
<nav>
<a href="/">Dashboard</a>
<a href="/mezzi">Mezzi</a>
<a href="/nuova">Nuova prenotazione</a>
<a href="/planning">Planning</a>
<a href="/storico">Storico</a>
<a href="/scadenze">Scadenze</a>
<a href="/cliente">Pagina cliente</a>
<a href="/test-email">Test Email</a>
</nav>
<main>${body}</main>
</body>
</html>`;
}

app.get('/', (req, res) => {
  const db = loadDb();
  res.send(layout('Dashboard', `
    <div class="grid">
      <a class="tile" href="/nuova">➕ Nuova prenotazione</a>
      <a class="tile" href="/planning">📅 Planning</a>
      <a class="tile" href="/mezzi">🚐 Mezzi</a>
      <a class="tile" href="/storico">📁 Storico</a>
      <a class="tile" href="/scadenze">⚠️ Scadenze</a>
      <a class="tile" href="/cliente">📲 Pagina cliente</a>
    </div>
    <div class="card">
      <h2>Situazione</h2>
      <p>Mezzi: <b>${db.mezzi.length}</b></p>
      <p>Prenotazioni: <b>${db.prenotazioni.length}</b></p>
      <p>Email: <b>${process.env.SMTP_HOST || 'non configurata'}</b></p>
    </div>
  `));
});

app.get('/test-email', async (req, res) => {
  try {
    await sendEmail(process.env.ALERT_EMAIL || process.env.SMTP_USER, 'TEST DP RENT', 'Email funzionante DP RENT');
    res.send('EMAIL OK');
  } catch (e) {
    res.send('ERRORE EMAIL: ' + e.message);
  }
});

app.get('/mezzi', (req, res) => {
  const db = loadDb();
  const rows = db.mezzi.map(m => `
    <tr>
      <td><a href="/mezzo/${m.id}">${esc(m.targa)}</a></td>
      <td>${esc(m.descrizione)}</td>
      <td>${esc(m.categoria)}</td>
      <td>${esc(m.km)}</td>
      <td>${esc(m.tagliando_km)}</td>
      <td>${esc(m.revisione)}</td>
      <td>${esc(m.bollo)}</td>
    </tr>`).join('');

  res.send(layout('Mezzi', `
    <div class="card">
      <h2>Nuovo mezzo</h2>
      <form method="POST" action="/mezzi">
        <input name="targa" placeholder="Targa" required>
        <input name="descrizione" placeholder="Descrizione cliente es. Opel Vivaro - pulmino 9 posti" required>
        <select name="categoria">
          <option>FURGONE</option><option>9_POSTI</option><option>AUTO</option><option>ESCAVATORE</option><option>SEMOVENTE</option>
        </select>
        <input name="km" type="number" placeholder="Km attuali">
        <input name="prezzo_giorno" type="number" step="0.01" placeholder="Prezzo giorno" value="70">
        <input name="km_inclusi" type="number" placeholder="Km inclusi giorno" value="150">
        <input name="tagliando_km" type="number" placeholder="Scadenza tagliando km">
        <input name="revisione" type="date">
        <input name="bollo" type="date">
        <button>Salva mezzo</button>
      </form>
    </div>
    <table><tr><th>Targa</th><th>Descrizione</th><th>Categoria</th><th>Km</th><th>Tagliando km</th><th>Revisione</th><th>Bollo</th></tr>${rows}</table>
  `));
});

app.post('/mezzi', (req, res) => {
  const db = loadDb();
  db.counters.mezzi++;
  db.mezzi.push({ id: db.counters.mezzi, ...req.body });
  saveDb(db);
  res.redirect('/mezzi');
});

app.get('/mezzo/:id', (req, res) => {
  const db = loadDb();
  const m = db.mezzi.find(x => String(x.id) === String(req.params.id));
  if (!m) return res.send('Mezzo non trovato');

  res.send(layout('Scheda mezzo', `
    <div class="card">
      <h2>${esc(m.targa)} - ${esc(m.descrizione)}</h2>
      <p>Km attuali: ${esc(m.km)}</p>
      <p>Tagliando km: ${esc(m.tagliando_km)}</p>
      <p>Revisione: ${esc(m.revisione)}</p>
      <p>Bollo: ${esc(m.bollo)}</p>
      <a class="btn" href="/nuova?mezzo=${m.id}">Prenota questo mezzo</a>
    </div>
  `));
});

function formPrenotazione() {
  const db = loadDb();
  const opts = db.mezzi.map(m => `<option value="${m.id}">${esc(m.descrizione)} (${esc(m.targa)})</option>`).join('');
  return `
  <form method="POST" action="/nuova">
    <div class="grid">
      <div><label>Cliente</label><input name="cliente" required></div>
      <div><label>Telefono</label><input name="telefono" required></div>
      <div><label>Email</label><input name="email"></div>
      <div><label>Codice fiscale</label><input name="codice_fiscale"></div>
      <div><label>Indirizzo</label><input name="indirizzo"></div>
      <div><label>Fatturazione</label><select name="fatturazione"><option>Privato</option><option>Azienda</option></select></div>
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

      <div><label>Mezzo</label><select name="mezzo_id">${opts}</select></div>
      <div><label>Data inizio</label><input type="date" name="data_inizio" required></div>
      <div><label>Ora inizio</label><input type="time" name="ora_inizio" value="08:30"></div>
      <div><label>Data fine</label><input type="date" name="data_fine" required></div>
      <div><label>Ora fine</label><input type="time" name="ora_fine" value="18:00"></div>
      <div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
      <div><label>Carburante uscita</label><select name="carburante_uscita"><option>4/4 pieno</option><option>3/4</option><option>1/2</option><option>1/4</option><option>Vuoto</option></select></div>
    </div>
    <button>Crea contratto</button>
  </form>`;
}

app.get('/nuova', (req, res) => {
  res.send(layout('Nuova prenotazione', `<div class="card"><h2>Nuova prenotazione</h2>${formPrenotazione()}</div>`));
});

app.post('/nuova', (req, res) => {
  const db = loadDb();
  const mezzo = db.mezzi.find(m => String(m.id) === String(req.body.mezzo_id));
  if (!mezzo) return res.send('Mezzo mancante');

  const c = calcola(req.body, mezzo);
  db.counters.prenotazioni++;
  const id = db.counters.prenotazioni;

  const p = {
    id,
    codice: `DPR-${todayCode()}-${String(id).padStart(4,'0')}`,
    stato: 'bozza',
    created_at: new Date().toLocaleString('it-IT'),
    ...req.body,
    mezzo_id: Number(req.body.mezzo_id),
    giorni: c.giorni,
    km_inclusi: c.kmInclusi,
    extra_km: c.extraKm,
    extra_fuori_orario: c.extraFuoriOrario,
    imponibile: c.imponibile,
    iva: c.iva,
    totale: c.totale,
    cauzione: 500
  };

  db.prenotazioni.push(p);
  saveDb(db);
  res.redirect('/prenotazione/' + id);
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
      <p><b>Fatturazione:</b> ${esc(p.fatturazione)} ${esc(p.ragione_sociale)} ${esc(p.piva)}</p>
      <p><b>Conducenti:</b> ${esc(p.conducente1)} / ${esc(p.conducente2)}</p>
      <p><b>Mezzo:</b> ${esc(m.targa)} - ${esc(m.descrizione)}</p>
      <p><b>Periodo:</b> ${esc(p.data_inizio)} ${esc(p.ora_inizio)} → ${esc(p.data_fine)} ${esc(p.ora_fine)}</p>
      <p><b>Extra fuori orario:</b> € ${euro(p.extra_fuori_orario)} + IVA</p>
      <p><b>Totale IVA inclusa:</b> € ${euro(p.totale)}</p>
      <a class="btn" href="/pdf/${p.id}">Scarica PDF</a>
      <a class="btn btn2" href="/email/${p.id}">Invia email</a>
      <a class="btn btn2" href="/foto/${p.id}">Foto/documenti</a>
    </div>
  `));
});

function generaPdf(p, m) {
  const file = path.join(CONTRACT_DIR, `contratto_${p.codice}.pdf`);
  const doc = new PDFDocument({ size: 'A4', margin: 40 });
  doc.pipe(fs.createWriteStream(file));

  doc.rect(0,0,595,105).fill('#111');
  doc.fillColor('white').fontSize(30).text('DP RENT', 45, 32);
  doc.fontSize(11).text(`${AZIENDA.nome}\n${AZIENDA.indirizzo}\nP.IVA ${AZIENDA.piva} | Tel. ${AZIENDA.telefono}\n${AZIENDA.email}`, 320, 20, { width: 220, align: 'right' });
  doc.rect(0,105,595,6).fill('#c40000');

  doc.fillColor('black').fontSize(22).text('CONTRATTO DI NOLEGGIO', 40, 140, { align:'center' });
  let y = 190;

  function section(t){ doc.rect(40,y,515,20).fill('#111'); doc.fillColor('white').fontSize(10).text(t,48,y+6); doc.fillColor('black'); y+=30; }
  function row(a,b){ doc.fontSize(9).text(a,50,y); doc.text(String(b||''),210,y); y+=16; }

  section('DATI CONTRATTO');
  row('Numero contratto', p.codice);
  row('Stato', p.stato);
  row('Data creazione', p.created_at);

  section('ANAGRAFICA E FATTURAZIONE');
  row('Cliente', p.cliente);
  row('Telefono', p.telefono);
  row('Email', p.email);
  row('Codice fiscale', p.codice_fiscale);
  row('Fatturazione', p.fatturazione);
  row('Ragione sociale', p.ragione_sociale);
  row('P.IVA', p.piva);
  row('PEC / SDI', `${p.pec || ''} ${p.sdi || ''}`);

  section('CONDUCENTI');
  row('Conducente 1', `${p.conducente1 || ''} - Patente ${p.patente1 || ''} scad. ${p.patente1_scadenza || ''}`);
  row('Conducente 2', `${p.conducente2 || ''} - Patente ${p.patente2 || ''} scad. ${p.patente2_scadenza || ''}`);

  section('VEICOLO E NOLEGGIO');
  row('Targa', m.targa);
  row('Descrizione', m.descrizione);
  row('Check-out', `${p.data_inizio} ore ${p.ora_inizio}`);
  row('Check-in', `${p.data_fine} ore ${p.ora_fine}`);
  row('Giorni', p.giorni);
  row('Carburante', p.carburante_uscita);
  row('Km previsti', p.km_previsti);

  section('RIEPILOGO ECONOMICO');
  row('Imponibile', `€ ${euro(p.imponibile)}`);
  row('IVA 22%', `€ ${euro(p.iva)}`);
  row('Totale IVA inclusa', `€ ${euro(p.totale)}`);
  row('Deposito cauzionale', `€ ${euro(p.cauzione)}`);

  section('CONDIZIONI GENERALI E PRIVACY');
  doc.fontSize(8).text(`Condizioni generali: ${TERMS_URL}`, 50, y); y+=13;
  doc.text(`Privacy: ${PRIVACY_URL}`, 50, y); y+=25;
  doc.text('Il cliente dichiara di aver preso visione e accettare condizioni generali e privacy. Veicolo da riconsegnare con carburante come consegnato. Danni, multe, pedaggi, franchigie e costi accessori a carico del cliente.', 50, y, { width: 500 });
  y+=60;

  doc.fontSize(10).text('Firma cliente: __________________________', 50, y);
  doc.text('Firma DP RENT: __________________________', 310, y);

  doc.end();
  return file;
}

app.get('/pdf/:id', (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));
  if (!p) return res.send('Contratto non trovato');
  const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};
  const file = generaPdf(p, m);
  setTimeout(() => res.download(file), 500);
});

app.get('/email/:id', async (req, res) => {
  const db = loadDb();
  const p = db.prenotazioni.find(x => String(x.id) === String(req.params.id));
  if (!p) return res.send('Contratto non trovato');
  const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};
  const file = generaPdf(p, m);

  try {
    await sendEmail(p.email || process.env.ALERT_EMAIL, `Contratto DP RENT ${p.codice}`, `Buongiorno,\nin allegato il contratto DP RENT ${p.codice}.\n\nDP RENT`, [
      { filename: path.basename(file), path: file }
    ]);
    res.send(layout('Email inviata', `<div class="card"><h2 class="ok">Email inviata</h2><a class="btn" href="/prenotazione/${p.id}">Torna</a></div>`));
  } catch (e) {
    res.send(layout('Errore email', `<div class="card"><h2 class="bad">Errore email</h2><pre>${esc(e.message)}</pre></div>`));
  }
});

app.get('/foto/:id', (req, res) => {
  const db = loadDb();
  const files = db.allegati.filter(a => String(a.prenotazione_id) === String(req.params.id));
  res.send(layout('Foto', `
    <div class="card">
      <h2>Foto / documenti</h2>
      <form method="POST" enctype="multipart/form-data">
        <select name="tipo">
          <option>Patente conducente 1</option>
          <option>Documento conducente 1</option>
          <option>Patente conducente 2</option>
          <option>Documento conducente 2</option>
          <option>Foto uscita</option>
          <option>Foto rientro</option>
          <option>Danno</option>
        </select>
        <input type="file" name="file" accept="image/*,.pdf" capture="environment" required>
        <button>Carica</button>
      </form>
      <ul>${files.map(f => `<li>${esc(f.tipo)} - <a href="/uploads/${esc(f.filename)}">${esc(f.originalname)}</a></li>`).join('')}</ul>
    </div>
  `));
});

app.post('/foto/:id', upload.single('file'), (req, res) => {
  const db = loadDb();
  db.counters.allegati++;
  db.allegati.push({
    id: db.counters.allegati,
    prenotazione_id: Number(req.params.id),
    tipo: req.body.tipo,
    filename: req.file.filename,
    originalname: req.file.originalname,
    created_at: new Date().toLocaleString('it-IT')
  });
  saveDb(db);
  res.redirect('/foto/' + req.params.id);
});

app.get('/storico', (req, res) => {
  const db = loadDb();
  const rows = db.prenotazioni.slice().reverse().map(p => {
    const m = db.mezzi.find(x => x.id === p.mezzo_id) || {};
    return `<tr><td><a href="/prenotazione/${p.id}">${esc(p.codice)}</a></td><td>${esc(p.cliente)}</td><td>${esc(m.targa)}</td><td>${esc(p.data_inizio)} → ${esc(p.data_fine)}</td><td>€ ${euro(p.totale)}</td></tr>`;
  }).join('');
  res.send(layout('Storico', `<table><tr><th>Contratto</th><th>Cliente</th><th>Mezzo</th><th>Date</th><th>Totale</th></tr>${rows}</table>`));
});

app.get('/planning', (req, res) => {
  const db = loadDb();
  const now = new Date();
  const y = Number(req.query.y || now.getFullYear());
  const mth = Number(req.query.m || now.getMonth() + 1);
  const days = new Date(y, mth, 0).getDate();

  let head = '<th>Mezzo</th>';
  for (let d=1; d<=days; d++) head += `<th>${d}</th>`;

  let rows = db.mezzi.map(m => {
    let r = `<tr><td><b>${esc(m.targa)}</b><br>${esc(m.descrizione)}</td>`;
    for (let d=1; d<=days; d++) {
      const date = `${y}-${String(mth).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
      const occ = db.prenotazioni.find(p => p.mezzo_id === m.id && p.data_inizio <= date && p.data_fine >= date);
      r += occ ? `<td class="busy" onclick="location.href='/prenotazione/${occ.id}'">O</td>` : `<td class="free" onclick="location.href='/nuova'">L</td>`;
    }
    return r + '</tr>';
  }).join('');

  res.send(layout('Planning', `<h2>Planning ${mth}/${y}</h2><table><tr>${head}</tr>${rows}</table>`));
});

app.get('/cliente', (req, res) => {
  res.send(layout('Pagina cliente', `<div class="card"><h2>Richiesta cliente</h2>${formPrenotazione()}</div>`));
});

app.get('/scadenze', (req, res) => {
  const db = loadDb();
  const rows = db.mezzi.map(m => `
    <tr><td>${esc(m.targa)}</td><td>${esc(m.descrizione)}</td><td>${esc(m.km)}</td><td>${esc(m.tagliando_km)}</td><td>${esc(m.revisione)}</td><td>${esc(m.bollo)}</td></tr>
  `).join('');
  res.send(layout('Scadenze', `<table><tr><th>Targa</th><th>Mezzo</th><th>Km</th><th>Tagliando km</th><th>Revisione</th><th>Bollo</th></tr>${rows}</table>`));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log('DP RENT APP avviata su porta ' + PORT);
});