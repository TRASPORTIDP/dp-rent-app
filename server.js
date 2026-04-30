require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const sqlite3 = require('sqlite3').verbose();
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json({ limit: '10mb' }));

// ================= DB =================
const db = new sqlite3.Database('./db.sqlite');

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      codice TEXT,
      cliente TEXT,
      telefono TEXT,
      email TEXT,
      mezzo TEXT,
      targa TEXT,
      dal TEXT,
      al TEXT,
      totale REAL,
      cauzione REAL,
      firma TEXT
    )
  `);
});

// ================= UTILS =================
function generaCodice() {
  return 'DPR-' + new Date().toISOString().slice(0,10).replace(/-/g,'') + '-' + Math.floor(Math.random()*1000);
}

// ================= HOME =================
app.get('/', (req, res) => {
  res.send('<h1>DP RENT APP V4 ATTIVA 🚀</h1>');
});

// ================= NUOVA PRENOTAZIONE =================
app.post('/prenotazione', (req, res) => {

  const codice = generaCodice();

  const {
    cliente,
    telefono,
    email,
    mezzo,
    targa,
    dal,
    al,
    totale,
    firma
  } = req.body;

  const cauzione = 500;

  db.run(`
    INSERT INTO prenotazioni 
    (codice, cliente, telefono, email, mezzo, targa, dal, al, totale, cauzione, firma)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `, [codice, cliente, telefono, email, mezzo, targa, dal, al, totale, cauzione, firma]);

  res.json({ codice });
});

// ================= PDF CONTRATTO =================
app.get('/contratto/:codice', (req, res) => {

  db.get("SELECT * FROM prenotazioni WHERE codice = ?", [req.params.codice], (err, row) => {

    if (!row) return res.send("Contratto non trovato");

    const doc = new PDFDocument({ margin: 40 });
    const filePath = `./contratto_${row.codice}.pdf`;

    doc.pipe(fs.createWriteStream(filePath));

    // ===== HEADER =====
    doc.fontSize(18).text("TRASPORTI DP SRL", { align: "center" });
    doc.fontSize(10).text("Via Tuderte 466 - Narni (TR)", { align: "center" });
    doc.text("P.IVA 01385450554", { align: "center" });

    doc.moveDown();

    doc.fontSize(16).text("CONTRATTO DI NOLEGGIO", { align: "center" });

    doc.moveDown();

    // ===== DATI CLIENTE =====
    doc.fontSize(12).text(`Cliente: ${row.cliente}`);
    doc.text(`Telefono: ${row.telefono}`);
    doc.text(`Email: ${row.email}`);

    doc.moveDown();

    // ===== VEICOLO =====
    doc.text(`Veicolo: ${row.mezzo}`);
    doc.text(`Targa: ${row.targa}`);

    doc.moveDown();

    // ===== DATE =====
    doc.text(`Dal: ${row.dal}`);
    doc.text(`Al: ${row.al}`);

    doc.moveDown();

    // ===== COSTI =====
    doc.text(`Totale: € ${row.totale}`);
    doc.text(`Cauzione: € ${row.cauzione}`);

    doc.moveDown();

    // ===== CLAUSOLE =====
    doc.fontSize(10).text(`
CONDIZIONI:
- Il veicolo viene consegnato con il pieno e deve essere restituito con il pieno.
- 150 km inclusi al giorno.
- Km extra €0.15/km.
- Ritiro serale +30€ + IVA.
`);

    doc.moveDown();

    // ===== LINK =====
    doc.text("Condizioni generali:", { underline: true });
    doc.text("https://carrentalsoftware.myappy.it/data/public/user/65996976/terms_file.pdf");

    doc.moveDown();

    doc.text("Privacy:", { underline: true });
    doc.text("https://carrentalsoftware.myappy.it/data/public/user/65996976/privacy_file.pdf");

    doc.moveDown();

    // ===== FIRMA =====
    doc.text("Firma cliente:");

    if (row.firma) {
      const base64Data = row.firma.split(',')[1];
      const imgBuffer = Buffer.from(base64Data, 'base64');
      doc.image(imgBuffer, { width: 150 });
    }

    doc.end();

    doc.on('finish', () => {
      res.download(filePath);
    });

  });

});

// ================= AVVIO =================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("DP RENT APP ONLINE 🚀"));
