const express = require('express');
const bodyParser = require('body-parser');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');
const moment = require('moment');
const XLSX = require('xlsx');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

const PORT = process.env.PORT || 3000;

// CONFIG
const IVA = 0.22;
const EXTRA_KM = 0.15;
const CAUZIONE = 500;
const EXTRA_SERA = 30;

// DB
const db = new sqlite3.Database('./database.sqlite');

// CARTELLE
const uploadDir = path.join(__dirname, 'uploads');
const contractsDir = path.join(__dirname, 'contracts');

if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
if (!fs.existsSync(contractsDir)) fs.mkdirSync(contractsDir);

const upload = multer({ dest: uploadDir });

// INIT DB
db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      targa TEXT UNIQUE,
      marca TEXT,
      modello TEXT,
      categoria TEXT,
      prezzo_giorno REAL,
      km_inclusi INTEGER,
      stato TEXT DEFAULT 'disponibile'
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      nome TEXT,
      telefono TEXT,
      mezzo_id INTEGER,
      data_inizio TEXT,
      data_fine TEXT,
      giorni INTEGER,
      km_previsti INTEGER,
      ritiro_serale INTEGER,
      imponibile REAL,
      iva REAL,
      totale REAL,
      cauzione REAL,
      stato TEXT DEFAULT 'attivo'
    )
  `);
});

// TEMPLATE HTML
function page(title, content) {
  return `
  <html>
  <head>
    <title>${title}</title>
    <style>
      body{font-family:Arial;margin:0;background:#f4f4f4;}
      header{background:#111;color:#fff;padding:15px;}
      nav{background:#b30000;padding:10px;}
      nav a{color:#fff;margin-right:15px;text-decoration:none;font-weight:bold;}
      .box{background:#fff;padding:20px;margin:20px;border-radius:8px;}
      table{width:100%;border-collapse:collapse;}
      th,td{padding:10px;border-bottom:1px solid #ddd;}
      th{background:#222;color:#fff;}
      button{background:#b30000;color:#fff;padding:10px;border:0;}
      input,select{padding:10px;width:100%;margin-bottom:10px;}
    </style>
  </head>
  <body>
    <header><h2>DP RENT APP</h2></header>
    <nav>
      <a href="/">Dashboard</a>
      <a href="/mezzi-web">Mezzi</a>
      <a href="/import-mezzi">Import Excel</a>
      <a href="/nuova-prenotazione">Nuova prenotazione</a>
      <a href="/prenotazioni">Prenotazioni</a>
      <a href="/planning">Planning</a>
    </nav>
    <div class="box">${content}</div>
  </body>
  </html>`;
}

// DASHBOARD
app.get('/', (req, res) => {
  res.send(page('Dashboard', '<h2>Gestionale attivo 🚀</h2>'));
});

// IMPORT PAGINA
app.get('/import-mezzi', (req, res) => {
  res.send(page('Import', `
    <h2>Import Excel</h2>
    <form method="POST" enctype="multipart/form-data">
      <input type="file" name="file" required>
      <button>Carica</button>
    </form>
  `));
});

// IMPORT
app.post('/import-mezzi', upload.single('file'), (req, res) => {
  const wb = XLSX.readFile(req.file.path);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);

  rows.forEach(r => {
    if (!r.Targa) return;

    let prezzo = 70;
    let km = 150;

    if (String(r.Modello).toUpperCase().includes('DACIA')) prezzo = 50;
    if (String(r.Modello).toUpperCase().includes('GOLF')) prezzo = 60;

    db.run(`
      INSERT OR IGNORE INTO mezzi
      (targa, marca, modello, categoria, prezzo_giorno, km_inclusi)
      VALUES (?,?,?,?,?,?)
    `, [r.Targa, r.Marca, r.Modello, 'GEN', prezzo, km]);
  });

  fs.unlinkSync(req.file.path);
  res.redirect('/mezzi-web');
});

// LISTA MEZZI
app.get('/mezzi-web', (req, res) => {
  db.all(`SELECT * FROM mezzi`, [], (e, rows) => {
    let html = '<h2>Mezzi</h2><table><tr><th>Targa</th><th>Modello</th><th>Prezzo</th></tr>';
    rows.forEach(m => {
      html += `<tr><td>${m.targa}</td><td>${m.modello}</td><td>${m.prezzo_giorno}</td></tr>`;
    });
    html += '</table>';
    res.send(page('Mezzi', html));
  });
});

// NUOVA PRENOTAZIONE
app.get('/nuova-prenotazione', (req, res) => {
  db.all(`SELECT * FROM mezzi`, [], (e, mezzi) => {
    let opt = mezzi.map(m => `<option value="${m.id}">${m.targa} ${m.modello}</option>`).join('');

    res.send(page('Nuova', `
      <form method="POST" action="/prenota">
        Nome<input name="nome" required>
        Telefono<input name="telefono">
        Mezzo<select name="mezzo_id">${opt}</select>
        Dal<input type="date" name="data_inizio" required>
        Al<input type="date" name="data_fine" required>
        Km<input name="km_previsti">
        <label><input type="checkbox" name="ritiro_serale"> Ritiro serale</label>
        <button>Crea</button>
      </form>
    `));
  });
});

// CREA PRENOTAZIONE
app.post('/prenota', (req, res) => {
  const { nome, telefono, mezzo_id, data_inizio, data_fine, km_previsti } = req.body;
  const ritiro = req.body.ritiro_serale ? 1 : 0;

  db.get(`SELECT * FROM mezzi WHERE id=?`, [mezzo_id], (e, m) => {
    const giorni = moment(data_fine).diff(moment(data_inizio), 'days') + 1;
    let imp = giorni * m.prezzo_giorno;

    if (ritiro) imp += EXTRA_SERA;

    const iva = imp * IVA;
    const tot = imp + iva;

    db.run(`
      INSERT INTO prenotazioni
      (codice,nome,telefono,mezzo_id,data_inizio,data_fine,giorni,ritiro_serale,imponibile,iva,totale,cauzione)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
    `, [
      'TEMP', nome, telefono, mezzo_id, data_inizio, data_fine,
      giorni, ritiro, imp, iva, tot, CAUZIONE
    ], function () {
      const cod = `DPR-${moment().format('YYYYMMDD')}-${this.lastID}`;
      db.run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, this.lastID]);

      res.send(page('OK', `
        <h2>Prenotazione creata</h2>
        Codice: ${cod}<br>
        Totale: ${tot.toFixed(2)}<br>
        <a href="/contratto/${this.lastID}">Scarica PDF</a>
      `));
    });
  });
});

// PDF
app.get('/contratto/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (e, p) => {
    const file = path.join(contractsDir, `contratto_${p.id}.pdf`);
    const doc = new PDFDocument();
    const stream = fs.createWriteStream(file);

    doc.pipe(stream);
    doc.text('CONTRATTO DP RENT');
    doc.text(`Cliente: ${p.nome}`);
    doc.text(`Totale: ${p.totale}`);
    doc.end();

    stream.on('finish', () => res.download(file));
  });
});

// PRENOTAZIONI
app.get('/prenotazioni', (req, res) => {
  db.all(`SELECT * FROM prenotazioni`, [], (e, rows) => {
    let html = '<h2>Prenotazioni</h2>';
    rows.forEach(p => html += `<p>${p.codice} - ${p.nome}</p>`);
    res.send(page('Prenotazioni', html));
  });
});

// PLANNING
app.get('/planning', (req, res) => {
  db.all(`SELECT * FROM prenotazioni`, [], (e, rows) => {
    let html = '<h2>Planning</h2>';
    rows.forEach(p => html += `<p>${p.data_inizio} → ${p.data_fine} - ${p.codice}</p>`);
    res.send(page('Planning', html));
  });
});

app.listen(PORT, () => console.log('DP RENT COMPLETO 🔥'));
