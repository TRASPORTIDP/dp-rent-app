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

const IVA = 0.22;
const EXTRA_KM = 0.15;
const CAUZIONE = 500;
const EXTRA_SERA = 30;

const db = new sqlite3.Database('./database.sqlite');

const uploadDir = path.join(__dirname, 'uploads');
const contractsDir = path.join(__dirname, 'contracts');

if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir);
if (!fs.existsSync(contractsDir)) fs.mkdirSync(contractsDir);

const upload = multer({ dest: uploadDir });

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
      stazione TEXT,
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
      tipo_cliente TEXT,
      codice_fiscale TEXT,
      piva TEXT,
      ragione_sociale TEXT,
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
      carburante_uscita TEXT DEFAULT 'pieno',
      carburante_rientro TEXT DEFAULT 'pieno',
      stato TEXT DEFAULT 'confermata',
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);
});

function htmlPage(title, body) {
  return `
  <!DOCTYPE html>
  <html lang="it">
  <head>
    <meta charset="UTF-8">
    <title>${title}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      body { font-family: Arial, sans-serif; margin: 0; background:#f4f4f4; color:#222; }
      header { background:#111; color:white; padding:18px; }
      header h1 { margin:0; font-size:22px; }
      nav { background:#b30000; padding:10px; }
      nav a { color:white; margin-right:15px; text-decoration:none; font-weight:bold; }
      main { padding:20px; }
      .card { background:white; padding:18px; border-radius:10px; margin-bottom:20px; box-shadow:0 2px 6px #ccc; }
      table { width:100%; border-collapse:collapse; background:white; }
      th, td { padding:10px; border-bottom:1px solid #ddd; text-align:left; font-size:14px; }
      th { background:#222; color:white; }
      input, select, button { padding:10px; margin:5px 0; width:100%; box-sizing:border-box; }
      button { background:#b30000; color:white; border:0; cursor:pointer; font-weight:bold; border-radius:6px; }
      .ok { color:green; font-weight:bold; }
      .bad { color:red; font-weight:bold; }
    </style>
  </head>
  <body>
    <header><h1>DP RENT APP</h1></header>
    <nav>
      <a href="/">Dashboard</a>
      <a href="/mezzi-web">Mezzi</a>
      <a href="/import-mezzi">Import Excel</a>
      <a href="/nuova-prenotazione">Nuova prenotazione</a>
      <a href="/prenotazioni-web">Prenotazioni</a>
      <a href="/planning">Planning</a>
    </nav>
    <main>${body}</main>
  </body>
  </html>`;
}

function normalize(value) {
  return value === undefined || value === null ? '' : String(value).trim();
}

function categoriaFromRow(row) {
  const codice = normalize(row['Codice Tip'] || row['Codice Tipo'] || row['codice_tipo']).toUpperCase();
  const marca = normalize(row['Marca']).toUpperCase();
  const modello = normalize(row['Modello']).toUpperCase();
  const desc = normalize(row['Descrizion'] || row['Descrizione'] || row['Immagini consegna']).toUpperCase();

  if (marca.includes('DACIA') || modello.includes('DACIA') || desc.includes('DACIA')) return 'AUTO_DACIA';
  if (modello.includes('GOLF') || desc.includes('GOLF')) return 'AUTO_GOLF';
  if (codice.includes('X-ESC') || desc.includes('ESCAVATORE')) return 'ESCAVATORE';
  if (desc.includes('PIATTAFORMA') || desc.includes('SEMOVENTE')) return 'SEMOVENTE';
  if (codice.includes('P') || desc.includes('PERSONE') || desc.includes('9P')) return '9_POSTI';
  if (codice.includes('F') || desc.includes('MERCI') || desc.includes('CARGO')) return 'FURGONE';
  return 'FURGONE';
}

function prezzoCategoria(categoria) {
  if (categoria === 'AUTO_DACIA') return 50;
  if (categoria === 'AUTO_GOLF') return 60;
  if (categoria === 'ESCAVATORE') return 50;
  if (categoria === 'SEMOVENTE') return 50;
  return 70;
}

function kmInclusiCategoria(categoria) {
  if (categoria === 'ESCAVATORE' || categoria === 'SEMOVENTE') return 0;
  return 150;
}

function codicePratica(id) {
  return `DPR-${moment().format('YYYYMMDD')}-${String(id).padStart(4, '0')}`;
}

app.get('/', (req, res) => {
  db.get(`SELECT COUNT(*) as tot FROM mezzi`, [], (e1, mezzi) => {
    db.get(`SELECT COUNT(*) as tot FROM prenotazioni`, [], (e2, pren) => {
      res.send(htmlPage('Dashboard', `
        <div class="card">
          <h2>Gestionale DP RENT attivo 🚀</h2>
          <p>Mezzi caricati: <b>${mezzi ? mezzi.tot : 0}</b></p>
          <p>Prenotazioni: <b>${pren ? pren.tot : 0}</b></p>
        </div>
      `));
    });
  });
});

app.get('/import-mezzi', (req, res) => {
  res.send(htmlPage('Import Excel', `
    <div class="card">
      <h2>Import mezzi da Excel</h2>
      <form method="POST" action="/import-mezzi" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx,.xls,.csv" required>
        <button type="submit">Carica e importa</button>
      </form>
      <p>Usa il file Excel con colonne tipo: UID, Targa, Km percorsi, Marca, Modello, Cilindrata, Alimentazione, Codice Tip, Descrizione, Stazione.</p>
    </div>
  `));
});

app.post('/import-mezzi', upload.single('file'), (req, res) => {
  if (!req.file) return res.send('File mancante');

  const workbook = XLSX.readFile(req.file.path);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet);

  let imported = 0;
  let skipped = 0;

  const stmt = db.prepare(`
    INSERT OR REPLACE INTO mezzi
    (uid, targa, km, marca, modello, cilindrata, alimentazione, codice_tipo, categoria, descrizione, stazione, prezzo_giorno, km_inclusi, stato)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, COALESCE((SELECT stato FROM mezzi WHERE targa = ?), 'disponibile'))
  `);

  rows.forEach(row => {
    const targa = normalize(row['Targa']);
    if (!targa) {
      skipped++;
      return;
    }

    const categoria = categoriaFromRow(row);
    const prezzo = prezzoCategoria(categoria);
    const kmInclusi = kmInclusiCategoria(categoria);

    stmt.run([
      normalize(row['UID']),
      targa,
      Number(row['Km percor'] || row['Km percorsi'] || row['Km'] || 0),
      normalize(row['Marca']),
      normalize(row['Modello']),
      normalize(row['Cilindrata']),
      normalize(row['Alimentaz'] || row['Alimentazione']),
      normalize(row['Codice Tip'] || row['Codice Tipo']),
      categoria,
      normalize(row['Descrizion'] || row['Descrizione'] || row['Immagini consegna']),
      normalize(row['Stazione']),
      prezzo,
      kmInclusi,
      targa
    ]);

    imported++;
  });

  stmt.finalize();

  fs.unlinkSync(req.file.path);

  res.send(htmlPage('Import completato', `
    <div class="card">
      <h2 class="ok">Import completato</h2>
      <p>Mezzi importati/aggiornati: <b>${imported}</b></p>
      <p>Righe saltate: <b>${skipped}</b></p>
      <a href="/mezzi-web">Vai a elenco mezzi</a>
    </div>
  `));
});

app.get('/mezzi-web', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY categoria, marca, modello`, [], (err, rows) => {
    const trs = rows.map(m => `
      <tr>
        <td>${m.id}</td>
        <td><b>${m.targa}</b></td>
        <td>${m.marca}</td>
        <td>${m.modello}</td>
        <td>${m.categoria}</td>
        <td>€ ${m.prezzo_giorno}</td>
        <td>${m.km_inclusi}</td>
        <td>${m.stato}</td>
      </tr>
    `).join('');

    res.send(htmlPage('Mezzi', `
      <div class="card">
        <h2>Elenco mezzi</h2>
        <table>
          <tr>
            <th>ID</th><th>Targa</th><th>Marca</th><th>Modello</th><th>Categoria</th><th>Prezzo</th><th>Km/giorno</th><th>Stato</th>
          </tr>
          ${trs}
        </table>
      </div>
    `));
  });
});

app.get('/nuova-prenotazione', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY categoria, marca`, [], (err, mezzi) => {
    const options = mezzi.map(m => `
      <option value="${m.id}">${m.targa} - ${m.marca} ${m.modello} - ${m.categoria}</option>
    `).join('');

    res.send(htmlPage('Nuova prenotazione', `
      <div class="card">
        <h2>Nuova prenotazione</h2>
        <form method="POST" action="/prenota-web">
          <label>Nome cliente</label>
          <input name="nome" required>

          <label>Telefono</label>
          <input name="telefono">

          <label>Tipo cliente</label>
          <select name="tipo_cliente">
            <option value="privato">Privato</option>
            <option value="azienda">Azienda</option>
          </select>

          <label>Codice fiscale</label>
          <input name="codice_fiscale">

          <label>Partita IVA</label>
          <input name="piva">

          <label>Ragione sociale</label>
          <input name="ragione_sociale">

          <label>Mezzo</label>
          <select name="mezzo_id">${options}</select>

          <label>Data inizio</label>
          <input type="date" name="data_inizio" required>

          <label>Data fine</label>
          <input type="date" name="data_fine" required>

          <label>Km previsti</label>
          <input type="number" name="km_previsti" value="150">

          <label>
            <input type="checkbox" name="ritiro_serale" value="1" style="width:auto;">
            Ritiro sera prima +30€ + IVA
          </label>

          <button type="submit">Crea prenotazione</button>
        </form>
      </div>
    `));
  });
});

app.post('/prenota-web', (req, res) => {
  const {
    nome, telefono, tipo_cliente, codice_fiscale, piva, ragione_sociale,
    mezzo_id, data_inizio, data_fine, km_previsti
  } = req.body;

  const ritiro_serale = req.body.ritiro_serale ? 1 : 0;

  db.get(`SELECT * FROM mezzi WHERE id = ?`, [mezzo_id], (err, mezzo) => {
    if (!mezzo) return res.send('Mezzo non trovato');

    db.get(`
      SELECT * FROM prenotazioni
      WHERE mezzo_id = ?
      AND stato != 'annullata'
      AND date(data_inizio) <= date(?)
      AND date(data_fine) >= date(?)
    `, [mezzo_id, data_fine, data_inizio], (err2, occupata) => {
      if (occupata) {
        return res.send(htmlPage('Mezzo occupato', `
          <div class="card">
            <h2 class="bad">Mezzo già occupato in quelle date</h2>
            <p>Controlla il planning o scegli altro mezzo.</p>
            <a href="/nuova-prenotazione">Torna indietro</a>
          </div>
        `));
      }

      const giorni = moment(data_fine).diff(moment(data_inizio), 'days') + 1;
      let imponibile = giorni * mezzo.prezzo_giorno;

      const kmPrev = Number(km_previsti || 0);
      const kmInclusiTot = giorni * Number(mezzo.km_inclusi || 0);

      if (mezzo.km_inclusi > 0 && kmPrev > kmInclusiTot) {
        imponibile += (kmPrev - kmInclusiTot) * EXTRA_KM;
      }

      if (ritiro_serale) imponibile += EXTRA_SERA;

      const iva = imponibile * IVA;
      const totale = imponibile + iva;

      db.run(`
        INSERT INTO prenotazioni
        (codice, nome, telefono, tipo_cliente, codice_fiscale, piva, ragione_sociale, mezzo_id, data_inizio, data_fine, giorni, km_previsti, ritiro_serale, imponibile, iva, totale, cauzione)
        VALUES ('TEMP', ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `, [
        nome, telefono, tipo_cliente, codice_fiscale, piva, ragione_sociale,
        mezzo_id, data_inizio, data_fine, giorni, kmPrev, ritiro_serale,
        imponibile, iva, totale, CAUZIONE
      ], function (err3) {
        if (err3) return res.send(String(err3));

        const codice = codicePratica(this.lastID);

        db.run(`UPDATE prenotazioni SET codice = ? WHERE id = ?`, [codice, this.lastID]);

        res.send(htmlPage('Prenotazione creata', `
          <div class="card">
            <h2 class="ok">Prenotazione creata</h2>
            <p>Codice: <b>${codice}</b></p>
            <p>Totale IVA inclusa: <b>€ ${totale.toFixed(2)}</b></p>
            <p>Cauzione: <b>€ ${CAUZIONE}</b></p>
            <p><a href="/contratto/${this.lastID}">Scarica contratto PDF</a></p>
            <p><a href="/prenotazioni-web">Vai alle prenotazioni</a></p>
          </div>
        `));
      });
    });
  });
});

app.get('/prenotazioni-web', (req, res) => {
  db.all(`
    SELECT p.*, m.targa, m.marca, m.modello
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id = p.mezzo_id
    ORDER BY p.id DESC
  `, [], (err, rows) => {
    const trs = rows.map(p => `
      <tr>
        <td>${p.codice}</td>
        <td>${p.nome}</td>
        <td>${p.targa} ${p.marca || ''} ${p.modello || ''}</td>
        <td>${p.data_inizio}</td>
        <td>${p.data_fine}</td>
        <td>€ ${Number(p.totale || 0).toFixed(2)}</td>
        <td>${p.stato}</td>
        <td><a href="/contratto/${p.id}">PDF</a></td>
      </tr>
    `).join('');

    res.send(htmlPage('Prenotazioni', `
      <div class="card">
        <h2>Prenotazioni</h2>
        <table>
          <tr>
            <th>Codice</th><th>Cliente</th><th>Mezzo</th><th>Dal</th><th>Al</th><th>Totale</th><th>Stato</th><th>Contratto</th>
          </tr>
          ${trs}
        </table>
      </div>
    `));
  });
});

app.get('/planning', (req, res) => {
  db.all(`
    SELECT p.*, m.targa, m.marca, m.modello
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id = p.mezzo_id
    WHERE p.stato != 'annullata'
    ORDER BY p.data_inizio ASC
  `, [], (err, rows) => {
    const trs = rows.map(p => `
      <tr>
        <td>${p.data_inizio}</td>
        <td>${p.data_fine}</td>
        <td><b>${p.targa}</b> ${p.marca || ''} ${p.modello || ''}</td>
        <td>${p.nome}</td>
        <td>${p.codice}</td>
      </tr>
    `).join('');

    res.send(htmlPage('Planning', `
      <div class="card">
        <h2>Planning disponibilità</h2>
        <table>
          <tr><th>Dal</th><th>Al</th><th>Mezzo</th><th>Cliente</th><th>Codice</th></tr>
          ${trs}
        </table>
      </div>
    `));
  });
});

app.get('/contratto/:id', (req, res) => {
  db.get(`
    SELECT p.*, m.targa, m.marca, m.modello, m.categoria
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id = p.mezzo_id
    WHERE p.id = ?
  `, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');

    const file = path.join(contractsDir, `contratto_${p.codice}.pdf`);
    const doc = new PDFDocument({ margin: 50 });

    doc.pipe(fs.createWriteStream(file));

    doc.fontSize(20).text('DP RENT - CONTRATTO DI NOLEGGIO', { align: 'center' });
    doc.moveDown();

    doc.fontSize(12);
    doc.text(`Codice pratica: ${p.codice}`);
    doc.text(`Cliente: ${p.nome}`);
    doc.text(`Telefono: ${p.telefono || ''}`);
    doc.text(`Tipo cliente: ${p.tipo_cliente || ''}`);
    doc.text(`Codice fiscale: ${p.codice_fiscale || ''}`);
    doc.text(`Partita IVA: ${p.piva || ''}`);
    doc.text(`Ragione sociale: ${p.ragione_sociale || ''}`);

    doc.moveDown();
    doc.text(`Mezzo: ${p.targa} - ${p.marca || ''} ${p.modello || ''}`);
    doc.text(`Categoria: ${p.categoria || ''}`);
    doc.text(`Periodo: dal ${p.data_inizio} al ${p.data_fine}`);
    doc.text(`Giorni: ${p.giorni}`);
    doc.text(`Km previsti: ${p.km_previsti || 0}`);
    doc.text(`Ritiro serale: ${p.ritiro_serale ? 'SI (+30€ + IVA)' : 'NO'}`);

    doc.moveDown();
    doc.text(`Imponibile: € ${Number(p.imponibile || 0).toFixed(2)}`);
    doc.text(`IVA 22%: € ${Number(p.iva || 0).toFixed(2)}`);
    doc.text(`Totale: € ${Number(p.totale || 0).toFixed(2)}`);
    doc.text(`Cauzione: € ${Number(p.cauzione || CAUZIONE).toFixed(2)}`);

    doc.moveDown();
    doc.fontSize(14).text('CONDIZIONI PRINCIPALI');
    doc.fontSize(11);
    doc.text('Il veicolo viene consegnato con il pieno e deve essere riconsegnato con il pieno.');
    doc.text('Eventuali differenze di carburante saranno addebitate al cliente.');
    doc.text('Extra km: €0,15/km ove previsto.');
    doc.text('Eventuali danni, multe, franchigie o costi accessori sono a carico del cliente secondo condizioni DP RENT.');
    doc.text('Il cliente dichiara di aver ricevuto il veicolo in buono stato salvo annotazioni e foto di consegna.');

    doc.moveDown(2);
    doc.text('Firma cliente: ______________________________');
    doc.moveDown();
    doc.text('Firma DP RENT: ______________________________');

    doc.end();

    doc.on('finish', () => {
      res.download(file);
    });
  });
});

app.get('/mezzi', (req, res) => {
  db.all(`SELECT * FROM mezzi`, [], (err, rows) => res.json(rows || []));
});

app.listen(PORT, () => {
  console.log('DP RENT APP V2 attiva su porta ' + PORT);
});
