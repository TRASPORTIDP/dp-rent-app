const express = require('express');
const bodyParser = require('body-parser');
const sqlite3 = require('sqlite3').verbose();
const PDFDocument = require('pdfkit');
const fs = require('fs');
const moment = require('moment');

const app = express();
app.use(bodyParser.json());

const db = new sqlite3.Database('./database.sqlite');

// CONFIG
const IVA = 0.22;
const EXTRA_KM = 0.15;
const CAUZIONE = 500;
const EXTRA_SERA = 30;

// INIT DB
db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      targa TEXT,
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
      nome TEXT,
      telefono TEXT,
      mezzo_id INTEGER,
      data_inizio TEXT,
      data_fine TEXT,
      giorni INTEGER,
      km_previsti INTEGER,
      ritiro_serale INTEGER,
      totale REAL,
      stato TEXT DEFAULT 'attivo'
    )
  `);
});

// TEST
app.get('/', (req, res) => {
  res.send('DP RENT APP ATTIVA 🚀');
});

// LISTA MEZZI
app.get('/mezzi', (req, res) => {
  db.all(`SELECT * FROM mezzi`, [], (err, rows) => {
    if (err) return res.send(err);
    res.json(rows);
  });
});

// AGGIUNGI MEZZO
app.post('/mezzi', (req, res) => {
  const { targa, marca, modello, categoria } = req.body;

  let prezzo = 70;
  let km = 150;

  if (categoria === 'AUTO_DACIA') prezzo = 50;
  if (categoria === 'AUTO_GOLF') prezzo = 60;
  if (categoria === 'ESCAVATORE') prezzo = 50;
  if (categoria === 'SEMOVENTE') prezzo = 50;

  db.run(
    `INSERT INTO mezzi (targa, marca, modello, categoria, prezzo_giorno, km_inclusi)
     VALUES (?, ?, ?, ?, ?, ?)`,
    [targa, marca, modello, categoria, prezzo, km],
    function (err) {
      if (err) return res.send(err);
      res.send({ id: this.lastID });
    }
  );
});

// PRENOTAZIONE
app.post('/prenota', (req, res) => {
  const {
    nome,
    telefono,
    mezzo_id,
    data_inizio,
    data_fine,
    km_previsti,
    ritiro_serale
  } = req.body;

  db.get(`SELECT * FROM mezzi WHERE id = ?`, [mezzo_id], (err, mezzo) => {
    if (!mezzo) return res.send('Mezzo non trovato');

    const giorni = moment(data_fine).diff(moment(data_inizio), 'days') + 1;

    let totale = giorni * mezzo.prezzo_giorno;

    // EXTRA KM
    const kmInclusi = giorni * mezzo.km_inclusi;
    if (km_previsti > kmInclusi) {
      totale += (km_previsti - kmInclusi) * EXTRA_KM;
    }

    // RITIRO SERALE
    if (ritiro_serale) {
      totale += EXTRA_SERA;
    }

    const totaleIVA = totale + (totale * IVA);

    db.run(
      `INSERT INTO prenotazioni 
       (nome, telefono, mezzo_id, data_inizio, data_fine, giorni, km_previsti, ritiro_serale, totale)
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [nome, telefono, mezzo_id, data_inizio, data_fine, giorni, km_previsti, ritiro_serale ? 1 : 0, totaleIVA],
      function (err) {
        if (err) return res.send(err);

        res.send({
          id: this.lastID,
          totale: totaleIVA,
          cauzione: CAUZIONE
        });
      }
    );
  });
});

// CONTRATTO PDF
app.get('/contratto/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id = ?`, [req.params.id], (err, p) => {
    if (!p) return res.send('Non trovato');

    const doc = new PDFDocument();
    const file = `contratto_${p.id}.pdf`;

    doc.pipe(fs.createWriteStream(file));

    doc.fontSize(18).text('DP RENT - CONTRATTO NOLEGGIO');
    doc.moveDown();

    doc.text(`Cliente: ${p.nome}`);
    doc.text(`Telefono: ${p.telefono}`);
    doc.text(`Dal: ${p.data_inizio} al ${p.data_fine}`);
    doc.text(`Totale: € ${p.totale}`);
    doc.text(`Cauzione: € ${CAUZIONE}`);

    doc.moveDown();
    doc.text('Carburante: pieno/pieno');
    doc.text('Extra km: €0,15/km');
    doc.text('Ritiro serale: +30€');
    doc.text('IVA 22% inclusa');

    doc.end();

    res.download(file);
  });
});

// PORTA RENDER
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log('DP RENT APP attiva su porta ' + PORT);
});
