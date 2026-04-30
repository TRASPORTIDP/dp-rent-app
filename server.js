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
app.use(bodyParser.urlencoded({ extended: true, limit: '20mb' }));
app.use(bodyParser.json({ limit: '20mb' }));

const PORT = process.env.PORT || 3000;

const IVA = 0.22;
const EXTRA_KM = 0.15;
const CAUZIONE = 500;
const EXTRA_SERA = 30;

const AZIENDA = {
  nome: 'Trasporti D P S.R.L. - DP RENT',
  indirizzo: 'Via Tuderte 466, Narni (TR)',
  telefono: '0744817108',
  email: 'contabilita@trasportidp.com'
};

const db = new sqlite3.Database('./database.sqlite');

const uploadDir = path.join(__dirname, 'uploads');
const contractsDir = path.join(__dirname, 'contracts');
const firmeDir = path.join(__dirname, 'firme');

[uploadDir, contractsDir, firmeDir].forEach(d => {
  if (!fs.existsSync(d)) fs.mkdirSync(d);
});

const upload = multer({ dest: uploadDir });

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
      ritiro_serale INTEGER,
      imponibile REAL,
      iva REAL,
      totale REAL,
      cauzione REAL,
      carburante_uscita TEXT DEFAULT 'pieno',
      carburante_rientro TEXT DEFAULT 'pieno',
      stato TEXT DEFAULT 'bozza',
      firma_path TEXT,
      pdf_path TEXT,
      note TEXT,
      created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
  `);

  addColumn('prenotazioni', 'nome', 'TEXT');
  addColumn('prenotazioni', 'cognome', 'TEXT');
  addColumn('prenotazioni', 'email', 'TEXT');
  addColumn('prenotazioni', 'codice_fiscale', 'TEXT');
  addColumn('prenotazioni', 'indirizzo', 'TEXT');
  addColumn('prenotazioni', 'citta', 'TEXT');
  addColumn('prenotazioni', 'cap', 'TEXT');
  addColumn('prenotazioni', 'tipo_cliente', 'TEXT');
  addColumn('prenotazioni', 'piva', 'TEXT');
  addColumn('prenotazioni', 'ragione_sociale', 'TEXT');
  addColumn('prenotazioni', 'ora_inizio', 'TEXT');
  addColumn('prenotazioni', 'ora_fine', 'TEXT');
  addColumn('prenotazioni', 'firma_path', 'TEXT');
  addColumn('prenotazioni', 'pdf_path', 'TEXT');
  addColumn('prenotazioni', 'note', 'TEXT');
});

function page(title, content) {
  return `
  <!DOCTYPE html>
  <html lang="it">
  <head>
    <meta charset="UTF-8">
    <title>${title}</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
      body{font-family:Arial;margin:0;background:#f4f4f4;color:#222;}
      header{background:#111;color:#fff;padding:18px;}
      header h1{margin:0;font-size:26px;}
      nav{background:#b30000;padding:12px;display:flex;gap:14px;flex-wrap:wrap;}
      nav a{color:#fff;text-decoration:none;font-weight:bold;}
      main{padding:20px;}
      .box{background:#fff;padding:20px;margin-bottom:20px;border-radius:10px;box-shadow:0 2px 8px #ccc;}
      table{width:100%;border-collapse:collapse;background:#fff;}
      th,td{padding:9px;border:1px solid #ddd;font-size:13px;}
      th{background:#222;color:white;}
      input,select,textarea,button{padding:10px;margin:5px 0;width:100%;box-sizing:border-box;}
      button,.btn{background:#b30000;color:#fff;border:0;padding:10px;border-radius:6px;text-decoration:none;display:inline-block;font-weight:bold;cursor:pointer;}
      .btn2{background:#333;}
      .ok{color:green;font-weight:bold;}
      .bad{color:red;font-weight:bold;}
      .grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;}
      .small{font-size:12px;color:#555;}
      .libero{background:#1fae4b;color:white;text-align:center;font-weight:bold;}
      .occupato{background:#d90000;color:white;text-align:center;font-weight:bold;}
      canvas{border:2px solid #333;background:white;width:100%;height:220px;}
      @media(max-width:700px){.grid{grid-template-columns:1fr;} main{padding:10px;} table{font-size:11px;}}
    </style>
  </head>
  <body>
    <header><h1>DP RENT APP</h1></header>
    <nav>
      <a href="/">Dashboard</a>
      <a href="/mezzi-web">Mezzi</a>
      <a href="/import-mezzi">Import Excel</a>
      <a href="/nuova-prenotazione">Nuova prenotazione</a>
      <a href="/prenotazioni">Storico</a>
      <a href="/planning">Planning</a>
      <a href="/prenota">Pagina cliente</a>
    </nav>
    <main>${content}</main>
  </body>
  </html>`;
}

function normalize(v) {
  return v === undefined || v === null ? '' : String(v).trim();
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

function calcolaTotale(mezzo, data_inizio, data_fine, km_previsti, ritiro_serale) {
  const giorni = moment(data_fine).diff(moment(data_inizio), 'days') + 1;
  let imponibile = giorni * Number(mezzo.prezzo_giorno || 0);

  const kmInclusiTot = giorni * Number(mezzo.km_inclusi || 0);
  const kmPrev = Number(km_previsti || 0);

  if (mezzo.km_inclusi > 0 && kmPrev > kmInclusiTot) {
    imponibile += (kmPrev - kmInclusiTot) * EXTRA_KM;
  }

  if (ritiro_serale) imponibile += EXTRA_SERA;

  const iva = imponibile * IVA;
  const totale = imponibile + iva;

  return { giorni, imponibile, iva, totale };
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

function generaPdfContratto(id, callback) {
  db.get(`
    SELECT p.*, m.targa, m.marca, m.modello, m.categoria, m.km_inclusi
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id = p.mezzo_id
    WHERE p.id = ?
  `, [id], (err, p) => {
    if (err || !p) return callback(err || new Error('Contratto non trovato'));

    const safe = String(p.codice || id).replace(/[^a-zA-Z0-9_-]/g, '');
    const file = path.join(contractsDir, `contratto_${safe}.pdf`);
    const doc = new PDFDocument({ margin: 40 });
    const stream = fs.createWriteStream(file);

    doc.pipe(stream);

    doc.fontSize(20).text('CONTRATTO DI NOLEGGIO - DP RENT', { align: 'center' });
    doc.moveDown(0.5);
    doc.fontSize(10).text(AZIENDA.nome, { align: 'center' });
    doc.text(AZIENDA.indirizzo, { align: 'center' });
    doc.text(`Tel. ${AZIENDA.telefono} - Email: ${AZIENDA.email}`, { align: 'center' });

    doc.moveDown();
    doc.fontSize(12).text(`Codice contratto: ${p.codice}`);
    doc.text(`Stato: ${p.stato || 'bozza'}`);
    doc.text(`Data creazione: ${p.created_at || ''}`);

    doc.moveDown();
    doc.fontSize(14).text('DATI CLIENTE', { underline: true });
    doc.fontSize(11);
    doc.text(`Nome e cognome: ${p.nome || ''} ${p.cognome || ''}`);
    doc.text(`Telefono: ${p.telefono || ''}`);
    doc.text(`Email: ${p.email || ''}`);
    doc.text(`Codice fiscale: ${p.codice_fiscale || ''}`);
    doc.text(`Indirizzo: ${p.indirizzo || ''}`);
    doc.text(`Citta/CAP: ${p.citta || ''} ${p.cap || ''}`);
    doc.text(`Tipo cliente: ${p.tipo_cliente || ''}`);
    doc.text(`P.IVA/Ragione sociale: ${p.piva || ''} ${p.ragione_sociale || ''}`);

    doc.moveDown();
    doc.fontSize(14).text('DATI NOLEGGIO', { underline: true });
    doc.fontSize(11);
    doc.text(`Mezzo: ${p.targa || ''} - ${p.marca || ''} ${p.modello || ''}`);
    doc.text(`Categoria: ${p.categoria || ''}`);
    doc.text(`Periodo: dal ${p.data_inizio} ore ${p.ora_inizio || ''} al ${p.data_fine} ore ${p.ora_fine || ''}`);
    doc.text(`Giorni tariffati: ${p.giorni}`);
    doc.text(`Km inclusi: ${Number(p.km_inclusi || 0) * Number(p.giorni || 0)}`);
    doc.text(`Km previsti: ${p.km_previsti || 0}`);
    doc.text(`Ritiro serale: ${p.ritiro_serale ? 'SI (+30 euro + IVA)' : 'NO'}`);
    doc.text(`Carburante consegna: ${p.carburante_uscita || 'pieno'}`);
    doc.text(`Carburante riconsegna: ${p.carburante_rientro || 'pieno'}`);

    doc.moveDown();
    doc.fontSize(14).text('RIEPILOGO ECONOMICO', { underline: true });
    doc.fontSize(11);
    doc.text(`Imponibile: euro ${Number(p.imponibile || 0).toFixed(2)}`);
    doc.text(`IVA 22%: euro ${Number(p.iva || 0).toFixed(2)}`);
    doc.text(`Totale IVA inclusa: euro ${Number(p.totale || 0).toFixed(2)}`);
    doc.text(`Cauzione: euro ${Number(p.cauzione || CAUZIONE).toFixed(2)}`);

    doc.moveDown();
    doc.fontSize(14).text('CONDIZIONI PRINCIPALI', { underline: true });
    doc.fontSize(9);
    doc.text('1. Il veicolo viene consegnato con il pieno e deve essere riconsegnato con il pieno. Eventuali differenze saranno addebitate al cliente.');
    doc.text('2. I chilometri inclusi sono quelli indicati nel presente contratto. Gli eventuali km eccedenti saranno conteggiati a euro 0,15/km ove previsto.');
    doc.text('3. Eventuali danni, multe, pedaggi, franchigie, ritardi, smarrimento chiavi/documenti e costi accessori sono a carico del cliente.');
    doc.text('4. Il cliente dichiara di ricevere il mezzo in buono stato salvo annotazioni, foto o checklist di consegna.');
    doc.text('5. Il cliente si impegna a usare il mezzo nel rispetto del Codice della Strada e delle condizioni DP RENT.');

    doc.moveDown();
    doc.fontSize(14).text('PRIVACY', { underline: true });
    doc.fontSize(9);
    doc.text('Il cliente autorizza il trattamento dei dati personali ai fini della gestione del noleggio, fatturazione, adempimenti fiscali e amministrativi, secondo normativa GDPR vigente.');

    doc.moveDown(2);

    if (p.firma_path && fs.existsSync(p.firma_path)) {
      doc.text('Firma cliente acquisita digitalmente:');
      doc.image(p.firma_path, { fit: [220, 80] });
    } else {
      doc.text('Firma cliente: ______________________________');
    }

    doc.moveDown();
    doc.text('Firma DP RENT: ______________________________');

    doc.end();

    stream.on('finish', () => {
      db.run(`UPDATE prenotazioni SET pdf_path = ? WHERE id = ?`, [file, id]);
      callback(null, file);
    });

    stream.on('error', callback);
  });
}

// DASHBOARD
app.get('/', (req, res) => {
  db.get(`SELECT COUNT(*) as tot FROM mezzi`, [], (e1, mezzi) => {
    db.get(`SELECT COUNT(*) as tot FROM prenotazioni`, [], (e2, pren) => {
      res.send(page('Dashboard', `
        <div class="box">
          <h2>Gestionale DP RENT attivo</h2>
          <p>Mezzi caricati: <b>${mezzi ? mezzi.tot : 0}</b></p>
          <p>Contratti / prenotazioni: <b>${pren ? pren.tot : 0}</b></p>
        </div>
      `));
    });
  });
});

// IMPORT EXCEL
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
    (uid,targa,km,marca,modello,cilindrata,alimentazione,codice_tipo,categoria,descrizione,stazione,prezzo_giorno,km_inclusi,stato)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,COALESCE((SELECT stato FROM mezzi WHERE targa=?),'disponibile'))
  `);

  rows.forEach(r => {
    const targa = normalize(r['Targa']);
    if (!targa) return;

    const cat = categoriaFromRow(r);
    stmt.run([
      normalize(r['UID']),
      targa,
      Number(r['Km percor'] || r['Km percorsi'] || r['Km'] || 0),
      normalize(r['Marca']),
      normalize(r['Modello']),
      normalize(r['Cilindrata']),
      normalize(r['Alimentaz'] || r['Alimentazione']),
      normalize(r['Codice Tip'] || r['Codice Tipo']),
      cat,
      normalize(r['Descrizion'] || r['Descrizione'] || r['Immagini consegna']),
      normalize(r['Stazione']),
      prezzoCategoria(cat),
      kmCategoria(cat),
      targa
    ]);
    imported++;
  });

  stmt.finalize();
  fs.unlinkSync(req.file.path);

  res.send(page('Import completato', `
    <h2 class="ok">Import completato</h2>
    <p>Mezzi importati/aggiornati: <b>${imported}</b></p>
    <a class="btn" href="/mezzi-web">Vai ai mezzi</a>
  `));
});

// MEZZI
app.get('/mezzi-web', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY categoria,targa`, [], (err, rows) => {
    const trs = rows.map(m => `
      <tr>
        <td>${m.id}</td>
        <td><b>${m.targa}</b></td>
        <td>${m.marca || ''}</td>
        <td>${m.modello || ''}</td>
        <td>${m.categoria || ''}</td>
        <td>euro ${Number(m.prezzo_giorno || 0).toFixed(2)}</td>
        <td>${m.km_inclusi}</td>
        <td>${m.stato}</td>
      </tr>
    `).join('');

    res.send(page('Mezzi', `
      <h2>Elenco mezzi</h2>
      <table>
        <tr><th>ID</th><th>Targa</th><th>Marca</th><th>Modello</th><th>Categoria</th><th>Prezzo/giorno</th><th>Km/giorno</th><th>Stato</th></tr>
        ${trs}
      </table>
    `));
  });
});

// NUOVA PRENOTAZIONE ADMIN
app.get('/nuova-prenotazione', (req, res) => {
  db.all(`SELECT * FROM mezzi ORDER BY categoria,targa`, [], (err, mezzi) => {
    const opt = mezzi.map(m => `<option value="${m.id}">${m.targa} - ${m.marca || ''} ${m.modello || ''} - ${m.categoria}</option>`).join('');

    res.send(page('Nuova prenotazione', `
      <h2>Nuova prenotazione / contratto</h2>
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
          <div><label>Data inizio</label><input type="date" name="data_inizio" required></div>
          <div><label>Ora inizio</label><input type="time" name="ora_inizio" value="08:30"></div>
          <div><label>Data fine</label><input type="date" name="data_fine" required></div>
          <div><label>Ora fine</label><input type="time" name="ora_fine" value="18:00"></div>
          <div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
          <div><label>Ritiro serale</label><select name="ritiro_serale"><option value="0">No</option><option value="1">Si +30 euro + IVA</option></select></div>
        </div>
        <label>Note</label><textarea name="note"></textarea>
        <button>Crea contratto</button>
      </form>
    `));
  });
});

app.post('/prenota-admin', (req, res) => {
  const b = req.body;
  const mezzo_id = b.mezzo_id;

  db.get(`SELECT * FROM mezzi WHERE id=?`, [mezzo_id], (err, mezzo) => {
    if (!mezzo) return res.send('Mezzo non trovato');

    queryDisponibilita(mezzo_id, b.data_inizio, b.data_fine, (e2, occ) => {
      if (occ) return res.send(page('Occupato', `<h2 class="bad">Mezzo occupato in queste date</h2><a href="/planning">Vai al planning</a>`));

      const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.km_previsti, Number(b.ritiro_serale) === 1);

      db.run(`
        INSERT INTO prenotazioni
        (codice,nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,tipo_cliente,piva,ragione_sociale,mezzo_id,data_inizio,data_fine,ora_inizio,ora_fine,giorni,km_previsti,ritiro_serale,imponibile,iva,totale,cauzione,stato,note)
        VALUES ('TEMP',?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
      `, [
        b.nome,b.cognome,b.telefono,b.email,b.codice_fiscale,b.indirizzo,b.citta,b.cap,b.tipo_cliente,b.piva,b.ragione_sociale,
        mezzo_id,b.data_inizio,b.data_fine,b.ora_inizio,b.ora_fine,calc.giorni,Number(b.km_previsti || 0),Number(b.ritiro_serale || 0),
        calc.imponibile,calc.iva,calc.totale,CAUZIONE,'bozza',b.note
      ], function(err3) {
        if (err3) return res.send(String(err3));
        const cod = codicePratica(this.lastID);
        db.run(`UPDATE prenotazioni SET codice=? WHERE id=?`, [cod, this.lastID]);
        res.send(page('Creato', `
          <h2 class="ok">Contratto creato</h2>
          <p>Codice: <b>${cod}</b></p>
          <p>Totale: <b>euro ${calc.totale.toFixed(2)}</b></p>
          <a class="btn" href="/contratto/${this.lastID}">Scarica PDF</a>
          <a class="btn btn2" href="/firma/${this.lastID}">Firma su tablet</a>
          <a class="btn btn2" href="/prenotazioni">Storico</a>
        `));
      });
    });
  });
});

// PAGINA CLIENTE AUTOMATICA
app.get('/prenota', (req, res) => {
  res.send(page('Prenota DP RENT', `
    <h2>Richiesta prenotazione cliente</h2>
    <form method="POST" action="/prenota-cliente">
      <div class="grid">
        <div><label>Tipo mezzo</label><select name="categoria" required>
          <option value="FURGONE">Furgone</option>
          <option value="9_POSTI">9 posti</option>
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
        <div><label>Data fine</label><input type="date" name="data_fine" required></div>
        <div><label>Km previsti</label><input type="number" name="km_previsti" value="150"></div>
        <div><label>Ritiro sera prima</label><select name="ritiro_serale"><option value="0">No</option><option value="1">Si +30 euro + IVA</option></select></div>
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

        const calc = calcolaTotale(mezzo, b.data_inizio, b.data_fine, b.km_previsti, Number(b.ritiro_serale) === 1);

        db.run(`
          INSERT INTO prenotazioni
          (codice,nome,cognome,telefono,email,codice_fiscale,indirizzo,citta,cap,tipo_cliente,mezzo_id,data_inizio,data_fine,ora_inizio,ora_fine,giorni,km_previsti,ritiro_serale,imponibile,iva,totale,cauzione,stato)
          VALUES ('TEMP',?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        `, [
          b.nome,b.cognome,b.telefono,b.email,b.codice_fiscale,b.indirizzo,b.citta,b.cap,'privato',
          mezzo.id,b.data_inizio,b.data_fine,'08:30','18:00',calc.giorni,Number(b.km_previsti || 0),Number(b.ritiro_serale || 0),
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

// STORICO CON FILTRI
app.get('/prenotazioni', (req, res) => {
  const q = normalize(req.query.q);
  const stato = normalize(req.query.stato);
  const dal = normalize(req.query.dal);
  const al = normalize(req.query.al);

  let sql = `
    SELECT p.*, m.targa, m.marca, m.modello
    FROM prenotazioni p
    LEFT JOIN mezzi m ON m.id=p.mezzo_id
    WHERE 1=1
  `;
  const params = [];

  if (q) {
    sql += ` AND (p.codice LIKE ? OR p.nome LIKE ? OR p.cognome LIKE ? OR p.telefono LIKE ? OR m.targa LIKE ?)`;
    params.push(`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`,`%${q}%`);
  }
  if (stato) {
    sql += ` AND p.stato=?`;
    params.push(stato);
  }
  if (dal) {
    sql += ` AND date(p.data_inizio)>=date(?)`;
    params.push(dal);
  }
  if (al) {
    sql += ` AND date(p.data_fine)<=date(?)`;
    params.push(al);
  }

  sql += ` ORDER BY p.id DESC`;

  db.all(sql, params, (err, rows) => {
    const trs = rows.map(p => `
      <tr>
        <td>${p.codice}</td>
        <td>${p.nome || ''} ${p.cognome || ''}</td>
        <td>${p.telefono || ''}<br>${p.email || ''}</td>
        <td><b>${p.targa || ''}</b><br>${p.marca || ''} ${p.modello || ''}</td>
        <td>${p.data_inizio} → ${p.data_fine}</td>
        <td>euro ${Number(p.totale || 0).toFixed(2)}</td>
        <td>${p.stato}</td>
        <td>
          <a href="/contratto/${p.id}">PDF</a><br>
          <a href="/firma/${p.id}">Firma</a><br>
          <a href="/email/${p.id}">Email</a><br>
          <a href="/stato/${p.id}/confermato">Conferma</a>
        </td>
      </tr>
    `).join('');

    res.send(page('Storico', `
      <h2>Storico contratti / prenotazioni</h2>
      <form method="GET" action="/prenotazioni" class="box">
        <div class="grid">
          <input name="q" placeholder="Cerca nome, targa, codice, telefono" value="${q}">
          <select name="stato">
            <option value="">Tutti gli stati</option>
            <option ${stato==='bozza'?'selected':''}>bozza</option>
            <option ${stato==='richiesta_cliente'?'selected':''}>richiesta_cliente</option>
            <option ${stato==='confermato'?'selected':''}>confermato</option>
            <option ${stato==='firmato'?'selected':''}>firmato</option>
            <option ${stato==='chiuso'?'selected':''}>chiuso</option>
            <option ${stato==='annullato'?'selected':''}>annullato</option>
          </select>
          <input type="date" name="dal" value="${dal}">
          <input type="date" name="al" value="${al}">
        </div>
        <button>Cerca</button>
      </form>
      <table>
        <tr><th>Codice</th><th>Cliente</th><th>Contatti</th><th>Mezzo</th><th>Date</th><th>Totale</th><th>Stato</th><th>Azioni</th></tr>
        ${trs}
      </table>
    `));
  });
});

app.get('/stato/:id/:stato', (req, res) => {
  db.run(`UPDATE prenotazioni SET stato=? WHERE id=?`, [req.params.stato, req.params.id], () => {
    res.redirect('/prenotazioni');
  });
});

// PLANNING MENSILE
app.get('/planning', (req, res) => {
  const mese = req.query.mese || moment().format('YYYY-MM');
  const start = moment(mese + '-01');
  const days = start.daysInMonth();

  db.all(`SELECT * FROM mezzi ORDER BY targa`, [], (e1, mezzi) => {
    db.all(`SELECT * FROM prenotazioni WHERE stato!='annullato'`, [], (e2, pren) => {
      let header = '<th>Mezzo</th>';
      for (let d = 1; d <= days; d++) header += `<th>${d}</th>`;

      let rows = '';
      mezzi.forEach(m => {
        rows += `<tr><td><b>${m.targa}</b><br>${m.modello || ''}</td>`;
        for (let d = 1; d <= days; d++) {
          const day = start.clone().date(d).format('YYYY-MM-DD');
          const occ = pren.find(p => p.mezzo_id == m.id && moment(day).isSameOrAfter(p.data_inizio) && moment(day).isSameOrBefore(p.data_fine));
          rows += occ ? `<td class="occupato" title="${occ.codice} ${occ.nome || ''}">O</td>` : `<td class="libero">L</td>`;
        }
        rows += '</tr>';
      });

      const prec = start.clone().subtract(1, 'month').format('YYYY-MM');
      const succ = start.clone().add(1, 'month').format('YYYY-MM');

      res.send(page('Planning', `
        <h2>Planning ${start.format('MM/YYYY')}</h2>
        <p><a href="/planning?mese=${prec}">← Mese precedente</a> | <a href="/planning?mese=${succ}">Mese successivo →</a></p>
        <p><span class="libero" style="padding:6px;">Libero</span> <span class="occupato" style="padding:6px;">Occupato</span></p>
        <div style="overflow-x:auto;"><table><tr>${header}</tr>${rows}</table></div>
      `));
    });
  });
});

// PDF
app.get('/contratto/:id', (req, res) => {
  generaPdfContratto(req.params.id, (err, file) => {
    if (err) return res.send('Errore PDF: ' + err.message);
    res.download(file);
  });
});

// FIRMA SU TABLET
app.get('/firma/:id', (req, res) => {
  res.send(page('Firma', `
    <h2>Firma contratto</h2>
    <p>Firma con dito o penna sul tablet.</p>
    <canvas id="canvas"></canvas>
    <br>
    <button onclick="clearCanvas()">Cancella</button>
    <button onclick="saveFirma()">Salva firma</button>

    <script>
      const canvas = document.getElementById('canvas');
      const ctx = canvas.getContext('2d');
      canvas.width = canvas.offsetWidth;
      canvas.height = 220;
      let drawing = false;

      function pos(e){
        const r = canvas.getBoundingClientRect();
        const t = e.touches ? e.touches[0] : e;
        return {x:t.clientX-r.left, y:t.clientY-r.top};
      }

      function start(e){drawing=true; const p=pos(e); ctx.beginPath(); ctx.moveTo(p.x,p.y); e.preventDefault();}
      function move(e){if(!drawing)return; const p=pos(e); ctx.lineTo(p.x,p.y); ctx.stroke(); e.preventDefault();}
      function end(e){drawing=false; e.preventDefault();}

      canvas.addEventListener('mousedown',start);
      canvas.addEventListener('mousemove',move);
      canvas.addEventListener('mouseup',end);
      canvas.addEventListener('touchstart',start);
      canvas.addEventListener('touchmove',move);
      canvas.addEventListener('touchend',end);

      function clearCanvas(){ctx.clearRect(0,0,canvas.width,canvas.height);}
      function saveFirma(){
        fetch('/firma/${req.params.id}', {
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body:JSON.stringify({firma:canvas.toDataURL('image/png')})
        }).then(()=>location.href='/prenotazioni');
      }
    </script>
  `));
});

app.post('/firma/:id', (req, res) => {
  const data = req.body.firma;
  if (!data) return res.status(400).send('Firma mancante');

  const base64 = data.replace(/^data:image\\/png;base64,/, '');
  const file = path.join(firmeDir, `firma_${req.params.id}.png`);
  fs.writeFileSync(file, base64, 'base64');

  db.run(`UPDATE prenotazioni SET firma_path=?, stato='firmato' WHERE id=?`, [file, req.params.id], () => {
    generaPdfContratto(req.params.id, () => res.send('OK'));
  });
});

// EMAIL
app.get('/email/:id', (req, res) => {
  db.get(`SELECT * FROM prenotazioni WHERE id=?`, [req.params.id], (err, p) => {
    if (!p) return res.send('Contratto non trovato');
    res.send(page('Invia email', `
      <h2>Invia contratto via email</h2>
      <form method="POST" action="/email/${p.id}">
        <label>Email destinatario</label>
        <input name="email" value="${p.email || ''}" required>
        <label>Messaggio</label>
        <textarea name="messaggio">Buongiorno, in allegato trova il contratto DP RENT.</textarea>
        <button>Invia email</button>
      </form>
      <p class="small">Per inviare davvero devi impostare su Render le variabili SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS.</p>
    `));
  });
});

app.post('/email/:id', (req, res) => {
  generaPdfContratto(req.params.id, async (err, file) => {
    if (err) return res.send('Errore PDF: ' + err.message);

    if (!process.env.SMTP_HOST) {
      return res.send(page('SMTP mancante', `
        <h2 class="bad">Email non configurata</h2>
        <p>Il PDF è pronto, ma per inviare email devi configurare SMTP su Render.</p>
        <p>Variabili richieste: SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS</p>
        <a href="/contratto/${req.params.id}">Scarica PDF</a>
      `));
    }

    const transporter = nodemailer.createTransport({
      host: process.env.SMTP_HOST,
      port: Number(process.env.SMTP_PORT || 587),
      secure: false,
      auth: {
        user: process.env.SMTP_USER,
        pass: process.env.SMTP_PASS
      }
    });

    await transporter.sendMail({
      from: process.env.SMTP_FROM || AZIENDA.email,
      to: req.body.email,
      subject: 'Contratto DP RENT',
      text: req.body.messaggio || 'In allegato contratto DP RENT.',
      attachments: [{ filename: path.basename(file), path: file }]
    });

    db.run(`UPDATE prenotazioni SET stato='inviato_email' WHERE id=?`, [req.params.id]);
    res.send(page('Email inviata', '<h2 class="ok">Email inviata correttamente</h2><a href="/prenotazioni">Torna allo storico</a>'));
  });
});

app.get('/mezzi', (req, res) => {
  db.all(`SELECT * FROM mezzi`, [], (err, rows) => res.json(rows || []));
});

app.listen(PORT, () => {
  console.log('DP RENT APP COMPLETA ONLINE');
});
