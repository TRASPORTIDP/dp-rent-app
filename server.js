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
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

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

// DASHBOARD
app.get('/', (req, res) => {
  res.send('DP RENT APP V3 ATTIVA 🚀');
});

// LISTA MEZZI
app.get('/mezzi', (req, res) => {
  db.all(`SELECT * FROM mezzi`, [], (err, rows) => {
    res.json(rows);
  });
});

// IMPORT EXCEL
const upload = multer({ dest: uploadDir });

app.post('/import-mezzi', upload.single('file'), (req, res) => {
  const workbook = XLSX.readFile(req.file.path);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);

  rows.forEach(row => {
    const targa = row['Targa'];
    if (!targa) return;

    let categoria = 'FURGONE';
    let prezzo = 70;
    let km = 150;

    if (String(row['Modello']).toUpperCase().includes('DACIA')) {
      categoria = 'AUTO_DACIA'; prezzo = 50;
    }

    if (String(row['Modello']).toUpperCase().includes('GOLF')) {
      categoria = 'AUTO_GOLF'; prezzo = 60;
    }

    db.run(`
      INSERT OR IGNORE INTO mezzi 
      (targa, marca, modello, categoria, prezzo_giorno, km_inclusi)
      VALUES (?, ?, ?, ?, ?, ?)
    `, [
      targa,
      row['Marca'],
      row['Modello'],
      categoria,
      prezzo,
      km
    ]);
  });

  fs.unlinkSync(req.file.path);

  res.send('Import completato');
});

// PRENOTAZIONE
app.post('/prenota', (req, res) => {
  const { nome, telefono, mezzo_id, data_inizio, data_fine, km_previsti, ritiro_serale } = req.body;

  db.get(`SELECT * FROM mezzi WHERE id = ?`, [mezzo_id], (err, mezzo) => {

    const giorni = moment(data_fine).diff(moment(data_inizio), 'days') + 1;

    let imponibile = giorni * mezzo.prezzo_giorno;

    const kmInclusi = giorni * mezzo.km_inclusi;

    if (km_previsti > kmInclusi) {
      imponibile += (km_previsti - kmInclusi) * EXTRA_KM;
    }

    if (ritiro_serale) {
      imponibile += EXTRA_SERA;
    }

    const iva = imponibile * IVA;
    const totale = imponibile + iva;

    db.run(`
      INSERT INTO prenotazioni 
      (codice, nome, telefono, mezzo_id, data_inizio, data_fine, giorni, km_previsti, ritiro_serale, imponibile, iva, totale, cauzione)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `, [
      'TEMP',
      nome,
      telefono,
      mezzo_id,
      data_inizio,
      data_fine,
      giorni,
      km_previsti,
      ritiro_serale ? 1 : 0,
      imponibile,
      iva,
      totale,
      CAUZIONE
    ], function () {

      const codice = `DPR-${moment().format('YYYYMMDD')}-${this.lastID}`;
      db.run(`UPDATE prenotazioni SET codice = ? WHERE id = ?`, [codice, this.lastID]);

      res.json({ codice, totale, cauzione: CAUZIONE });
    });

  });
});

// PDF CORRETTO
app.get('/contratto/:id', (req, res) => {

  db.get(`
    SELECT p.*, m.targa, m.marca, m.modello
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id = p.mezzo_id
    WHERE p.id = ?
  `, [req.params.id], (err, p) => {

    const file = path.join(contractsDir, `contratto_${p.id}.pdf`);

    const doc = new PDFDocument();
    const stream = fs.createWriteStream(file);

    doc.pipe(stream);

    doc.fontSize(18).text('DP RENT CONTRATTO NOLEGGIO');
    doc.moveDown();

    doc.text(`Cliente: ${p.nome}`);
    doc.text(`Telefono: ${p.telefono}`);
    doc.text(`Mezzo: ${p.targa} ${p.marca} ${p.modello}`);
    doc.text(`Periodo: ${p.data_inizio} - ${p.data_fine}`);
    doc.text(`Totale: ${p.totale}`);
    doc.text(`Cauzione: ${CAUZIONE}`);

    doc.moveDown();
    doc.text('Carburante pieno/pieno');
    doc.text('Extra km 0.15');
    doc.text('Ritiro serale +30');

    doc.end();

    stream.on('finish', () => {
      res.download(file);
    });

  });

});

// AVVIO
app.listen(PORT, () => {
  console.log('DP RENT V3 ONLINE 🔥');
});
