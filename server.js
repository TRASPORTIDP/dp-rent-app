require('dotenv').config();
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const nodemailer = require('nodemailer');
const PDFDocument = require('pdfkit');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use('/uploads', express.static('uploads'));
app.use('/contratti', express.static('contratti'));

const PORT = process.env.PORT || 3000;

// =======================
// DATABASE
// =======================
const db = new sqlite3.Database('./db.sqlite');

// =======================
// CARTELLE
// =======================
['uploads','contratti'].forEach(dir=>{
  if(!fs.existsSync(dir)) fs.mkdirSync(dir);
});

// =======================
// UPLOAD FOTO
// =======================
const upload = multer({ dest: 'uploads/' });

// =======================
// CREAZIONE TABELLE
// =======================
db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS mezzi (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      nome TEXT,
      categoria TEXT,
      km INTEGER DEFAULT 0,
      km_tagliando INTEGER DEFAULT 30000
    )
  `);

  db.run(`
    CREATE TABLE IF NOT EXISTS prenotazioni (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      cliente TEXT,
      telefono TEXT,
      mezzo_id INTEGER,
      data_inizio TEXT,
      data_fine TEXT,
      km_uscita INTEGER,
      km_rientro INTEGER,
      stato TEXT
    )
  `);
});

// =======================
// HOME
// =======================
app.get('/', (req, res) => {
  res.send(`
    <h1>DP RENT</h1>
    <a href="/mezzi">Mezzi</a><br>
    <a href="/nuovo">Nuova prenotazione</a><br>
    <a href="/planning">Planning</a><br>
    <a href="/storico">Storico</a>
  `);
});

// =======================
// MEZZI
// =======================
app.get('/mezzi', (req, res) => {
  db.all("SELECT * FROM mezzi", (err, rows) => {
    let html = `<h1>Mezzi</h1><a href="/">Home</a><br><br>`;
    rows.forEach(m => {
      html += `
        <div>
          <b>${m.nome}</b> (${m.categoria}) - KM: ${m.km}
        </div>
      `;
    });
    html += `
      <h3>Aggiungi mezzo</h3>
      <form method="POST">
        <input name="nome" placeholder="Nome mezzo"><br>
        <input name="categoria" placeholder="Categoria"><br>
        <button>Aggiungi</button>
      </form>
    `;
    res.send(html);
  });
});

app.post('/mezzi', (req,res)=>{
  db.run("INSERT INTO mezzi(nome,categoria) VALUES(?,?)",
    [req.body.nome, req.body.categoria],
    ()=> res.redirect('/mezzi')
  );
});

// =======================
// NUOVA PRENOTAZIONE
// =======================
app.get('/nuovo', (req, res) => {
  db.all("SELECT * FROM mezzi", (err, mezzi) => {
    let select = mezzi.map(m=>`<option value="${m.id}">${m.nome}</option>`).join('');
    res.send(`
      <h1>Nuova prenotazione</h1>
      <form method="POST">
        <input name="cliente" placeholder="Cliente"><br>
        <input name="telefono" placeholder="Telefono"><br>
        <select name="mezzo_id">${select}</select><br>
        <input type="date" name="data_inizio"><br>
        <input type="date" name="data_fine"><br>
        <button>Salva</button>
      </form>
    `);
  });
});

app.post('/nuovo', (req,res)=>{
  db.run(`
    INSERT INTO prenotazioni
    (cliente,telefono,mezzo_id,data_inizio,data_fine,stato)
    VALUES (?,?,?,?,?,'attivo')
  `,
  [req.body.cliente, req.body.telefono, req.body.mezzo_id, req.body.data_inizio, req.body.data_fine],
  ()=> res.redirect('/storico'));
});

// =======================
// PLANNING BASE
// =======================
app.get('/planning', (req,res)=>{
  db.all("SELECT * FROM mezzi", (err, mezzi)=>{
    db.all("SELECT * FROM prenotazioni WHERE stato='attivo'", (err, prenotazioni)=>{

      let html = `<h1>Planning</h1><a href="/">Home</a><br><br>`;

      mezzi.forEach(m=>{
        html += `<h3>${m.nome}</h3>`;

        for(let i=1;i<=30;i++){
          let occupato = prenotazioni.find(p=>p.mezzo_id==m.id);
          if(occupato){
            html += `<span style="background:red;color:white;padding:5px;margin:2px">O</span>`;
          } else {
            html += `<span style="background:green;color:white;padding:5px;margin:2px">L</span>`;
          }
        }

        html += `<br><br>`;
      });

      res.send(html);
    });
  });
});

// =======================
// STORICO
// =======================
app.get('/storico', (req,res)=>{
  db.all(`
    SELECT p.*, m.nome as mezzo
    FROM prenotazioni p
    JOIN mezzi m ON m.id=p.mezzo_id
  `,(err,rows)=>{
    let html = "<h1>Storico</h1>";
    rows.forEach(p=>{
      html+=`
        <div>
          ${p.cliente} - ${p.mezzo} - ${p.data_inizio}
          <a href="/contratto/${p.id}">PDF</a>
          <a href="/chiudi/${p.id}">Chiudi</a>
        </div>
      `;
    });
    res.send(html);
  });
});

// =======================
// PDF CONTRATTO
// =======================
app.get('/contratto/:id',(req,res)=>{
  db.get(`
    SELECT p.*, m.nome as mezzo
    FROM prenotazioni p
    JOIN mezzi m ON m.id=p.mezzo_id
    WHERE p.id=?
  `,[req.params.id],(err,p)=>{

    const file = `contratti/contratto_${p.id}.pdf`;
    const doc = new PDFDocument();

    doc.pipe(fs.createWriteStream(file));

    doc.fontSize(20).text("DP RENT", {align:'center'});
    doc.moveDown();

    doc.text(`Cliente: ${p.cliente}`);
    doc.text(`Telefono: ${p.telefono}`);
    doc.text(`Mezzo: ${p.mezzo}`);
    doc.text(`Dal: ${p.data_inizio}`);
    doc.text(`Al: ${p.data_fine}`);

    doc.end();

    doc.on('finish', ()=>{
      res.download(file);
    });
  });
});

// =======================
// CHIUSURA + KM
// =======================
app.get('/chiudi/:id',(req,res)=>{
  res.send(`
    <form method="POST">
      KM rientro: <input name="km"><br>
      <button>Chiudi</button>
    </form>
  `);
});

app.post('/chiudi/:id',(req,res)=>{
  const km = parseInt(req.body.km);

  db.get("SELECT * FROM prenotazioni WHERE id=?", [req.params.id], (err,p)=>{
    db.run("UPDATE mezzi SET km=? WHERE id=?",
      [km, p.mezzo_id]);

    db.run("UPDATE prenotazioni SET stato='chiuso', km_rientro=? WHERE id=?",
      [km, req.params.id],
      ()=> res.redirect('/storico')
    );
  });
});

// =======================
// SERVER
// =======================
app.listen(PORT, ()=> console.log("DP RENT attivo su "+PORT));
