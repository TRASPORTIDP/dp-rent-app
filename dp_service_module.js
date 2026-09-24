
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const PDFDocument = require('pdfkit');
let nodemailer = null;
try { nodemailer = require('nodemailer'); } catch(e) { console.log('Nodemailer non disponibile DP SERVICE:',e.message); }

let twilio = null;
try { twilio = require('twilio'); } catch(e) { console.log('Twilio non disponibile in DP SERVICE:', e.message); }

const router = express.Router();

let dataDir = process.env.DATA_DIR || '/var/data';
try { fs.mkdirSync(dataDir,{recursive:true}); } catch(e) { dataDir=path.join(__dirname,'data'); fs.mkdirSync(dataDir,{recursive:true}); }
const db = new sqlite3.Database(path.join(dataDir, 'dp_service.sqlite'));
const mainDbPath = process.env.DB_PATH || path.join(dataDir,'database.sqlite');
const mainDb = new sqlite3.Database(mainDbPath);


router.use(express.urlencoded({ extended: true }));
router.use(express.json());
router.use('/public', express.static(path.join(__dirname, 'public')));

// Quando il modulo e montato sotto /service, mantiene tutti i link/form interni nel modulo.
router.use((req,res,next)=>{
  const originalSend=res.send.bind(res);
  res.send=(body)=>{
    if(typeof body==='string'){
      body=body.replace(/\b(href|action|src)=([\"'])\/(?!service\/)/g,'$1=$2/service/');
      // Link speciali che devono restare FUORI dal prefisso /service.
      body=body.replace(/([\"'])\/service\/clienti-azienda(?=\/|[\"'])/g,'$1/clienti-azienda');
      // Link speciale per tornare al menu principale DP GESTIONALE.
      body=body.replace(/href=([\"'])\/service\/__DP_MAIN__\1/g,'href=$1/$1');
    }
    return originalSend(body);
  };
  const originalRedirect=res.redirect.bind(res);
  res.redirect=(...args)=>{
    if(args.length){
      const i=args.length-1;
      if(typeof args[i]==='string' && args[i].startsWith('/') && !args[i].startsWith('/service/') && !args[i].startsWith('/clienti-azienda')) args[i]='/service'+args[i];
    }
    return originalRedirect(...args);
  };
  next();
});

const run = (sql, params=[]) => new Promise((resolve,reject)=>db.run(sql, params, function(err){ if(err) reject(err); else resolve(this); }));
const all = (sql, params=[]) => new Promise((resolve,reject)=>db.all(sql, params, (err,rows)=>err?reject(err):resolve(rows)));
const get = (sql, params=[]) => new Promise((resolve,reject)=>db.get(sql, params, (err,row)=>err?reject(err):resolve(row)));
const mainRun = (sql, params=[]) => new Promise((resolve,reject)=>mainDb.run(sql, params, function(err){ if(err) reject(err); else resolve(this); }));
const mainAll = (sql, params=[]) => new Promise((resolve,reject)=>mainDb.all(sql, params, (err,rows)=>err?reject(err):resolve(rows)));
const mainGet = (sql, params=[]) => new Promise((resolve,reject)=>mainDb.get(sql, params, (err,row)=>err?reject(err):resolve(row)));

async function dpServiceSyncMaster(){
  const exists=await mainGet(`SELECT name FROM sqlite_master WHERE type='table' AND name='dp_clienti_master'`).catch(()=>null);
  if(!exists) return 0;
  const rows=await mainAll(`SELECT * FROM dp_clienti_master ORDER BY id`).catch(()=>[]);
  let n=0;
  for(const m of rows){
    let ex=await get(`SELECT * FROM clienti WHERE source_main_id=? LIMIT 1`,[m.id]).catch(()=>null);
    if(!ex && String(m.piva||'').trim()) ex=await get(`SELECT * FROM clienti WHERE TRIM(COALESCE(piva,''))<>'' AND UPPER(TRIM(piva))=UPPER(TRIM(?)) LIMIT 1`,[m.piva]).catch(()=>null);
    if(!ex && String(m.codice_fiscale||'').trim()) ex=await get(`SELECT * FROM clienti WHERE TRIM(COALESCE(cf,''))<>'' AND UPPER(TRIM(cf))=UPPER(TRIM(?)) LIMIT 1`,[m.codice_fiscale]).catch(()=>null);
    if(!ex && String(m.ragione_sociale||'').trim() && String(m.telefono||'').trim()) ex=await get(`SELECT * FROM clienti WHERE UPPER(TRIM(COALESCE(ragione_sociale,'')))=UPPER(TRIM(?)) AND REPLACE(COALESCE(telefono,''),' ','')=REPLACE(?,' ','') LIMIT 1`,[m.ragione_sociale,m.telefono]).catch(()=>null);
    const vals=[m.codice,m.ragione_sociale,m.piva,m.codice_fiscale,m.indirizzo,m.citta,m.provincia,m.telefono,m.email,m.pec,m.sdi,
      m.codice_iva,m.codice_pagamento,m.banca_cliente,m.banca_codice,m.banca_conto,m.banca_abi,m.banca_cab,m.banca_paese,m.banca_cin_eur,m.banca_cin_it,m.banca_valuta,m.banca_bic,m.banca_iban,m.banca_bban,m.banca_pec,m.banca_note,m.note,m.id];
    if(ex){
      await run(`UPDATE clienti SET codice=?,ragione_sociale=?,piva=?,cf=?,indirizzo=?,citta=?,provincia=?,telefono=?,email=?,pec=?,sdi=?,codice_iva=?,codice_pagamento=?,banca_cliente=?,banca_codice=?,banca_conto=?,banca_abi=?,banca_cab=?,banca_paese=?,banca_cin_eur=?,banca_cin_it=?,banca_valuta=?,banca_bic=?,banca_iban=?,banca_bban=?,banca_pec=?,banca_note=?,note=?,source_main_id=? WHERE id=?`,[...vals,ex.id]).catch(()=>{});
    }else{
      await run(`INSERT INTO clienti(codice,ragione_sociale,piva,cf,indirizzo,citta,provincia,telefono,email,pec,sdi,codice_iva,codice_pagamento,banca_cliente,banca_codice,banca_conto,banca_abi,banca_cab,banca_paese,banca_cin_eur,banca_cin_it,banca_valuta,banca_bic,banca_iban,banca_bban,banca_pec,banca_note,note,source_main_id) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,vals).catch(()=>{});
    }
    n++;
  }
  return n;
}
async function dpServiceSaveToMaster(localId,b){
  const local=await get(`SELECT * FROM clienti WHERE id=?`,[localId]).catch(()=>null);
  const data={
    codice:String(b.codice||local?.codice||'').trim(),ragione_sociale:String(b.ragione_sociale||local?.ragione_sociale||'').trim(),
    piva:String(b.piva||local?.piva||'').trim(),codice_fiscale:String(b.cf||b.codice_fiscale||local?.cf||'').trim(),
    indirizzo:String(b.indirizzo||local?.indirizzo||'').trim(),citta:String(b.citta||local?.citta||'').trim(),provincia:String(b.provincia||local?.provincia||'').trim(),
    telefono:String(b.telefono||local?.telefono||'').trim(),email:String(b.email||local?.email||'').trim(),pec:String(b.pec||local?.pec||'').trim(),sdi:String(b.sdi||local?.sdi||'').trim(),
    codice_iva:String(b.codice_iva||local?.codice_iva||'').trim(),codice_pagamento:String(b.codice_pagamento||local?.codice_pagamento||'').trim(),
    banca_cliente:String(b.banca_cliente||local?.banca_cliente||'').trim(),banca_codice:String(b.banca_codice||local?.banca_codice||'').trim(),banca_conto:String(b.banca_conto||local?.banca_conto||'').trim(),
    banca_abi:String(b.banca_abi||local?.banca_abi||'').trim(),banca_cab:String(b.banca_cab||local?.banca_cab||'').trim(),banca_paese:String(b.banca_paese||local?.banca_paese||'IT').trim(),
    banca_cin_eur:String(b.banca_cin_eur||local?.banca_cin_eur||'').trim(),banca_cin_it:String(b.banca_cin_it||local?.banca_cin_it||'').trim(),banca_valuta:String(b.banca_valuta||local?.banca_valuta||'EUR').trim(),
    banca_bic:String(b.banca_bic||local?.banca_bic||'').trim(),banca_iban:String(b.banca_iban||local?.banca_iban||'').trim(),banca_bban:String(b.banca_bban||local?.banca_bban||'').trim(),
    banca_pec:String(b.banca_pec||local?.banca_pec||'').trim(),banca_note:String(b.banca_note||local?.banca_note||'').trim(),note:String(b.note||local?.note||'').trim()
  };
  let masterId=Number(local?.source_main_id)||0;
  if(masterId){
    const cols=Object.keys(data); await mainRun(`UPDATE dp_clienti_master SET ${cols.map(k=>`${k}=?`).join(',')},updated_at=CURRENT_TIMESTAMP WHERE id=?`,[...Object.values(data),masterId]);
  }else{
    let m=null;
    if(data.piva)m=await mainGet(`SELECT * FROM dp_clienti_master WHERE UPPER(TRIM(COALESCE(piva,'')))=UPPER(TRIM(?))`,[data.piva]).catch(()=>null);
    if(!m&&data.codice_fiscale)m=await mainGet(`SELECT * FROM dp_clienti_master WHERE UPPER(TRIM(COALESCE(codice_fiscale,'')))=UPPER(TRIM(?))`,[data.codice_fiscale]).catch(()=>null);
    if(m) masterId=m.id;
    else{
      const cols=Object.keys(data);const r=await mainRun(`INSERT INTO dp_clienti_master(${cols.join(',')},updated_at) VALUES(${cols.map(()=>'?').join(',')},CURRENT_TIMESTAMP)`,Object.values(data));masterId=r.lastID;
    }
    await run(`UPDATE clienti SET source_main_id=? WHERE id=?`,[masterId,localId]).catch(()=>{});
  }
  await dpServiceSyncMaster();
  return masterId;
}


async function initDb(){
  await run(`CREATE TABLE IF NOT EXISTS clienti (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ragione_sociale TEXT NOT NULL,
    piva TEXT, cf TEXT, indirizzo TEXT, citta TEXT, provincia TEXT,
    telefono TEXT, email TEXT, pec TEXT, sdi TEXT, note TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  const ccols = await all(`PRAGMA table_info(clienti)`);
  const have = new Set(ccols.map(x => x.name));
  for (const [name, type] of [
    ['cf','TEXT'],['indirizzo','TEXT'],['citta','TEXT'],['provincia','TEXT'],
    ['pec','TEXT'],['sdi','TEXT'],['note','TEXT'],['created_at','TEXT'],['codice','TEXT'],['source_main_id','INTEGER'],['source_key','TEXT'],
    ['codice_iva','TEXT'],['codice_pagamento','TEXT'],['banca_cliente','TEXT'],['banca_codice','TEXT'],['banca_conto','TEXT'],
    ['banca_abi','TEXT'],['banca_cab','TEXT'],['banca_paese','TEXT'],['banca_cin_eur','TEXT'],['banca_cin_it','TEXT'],['banca_valuta','TEXT'],
    ['banca_bic','TEXT'],['banca_iban','TEXT'],['banca_bban','TEXT'],['banca_pec','TEXT'],['banca_note','TEXT']
  ]) {
    if (!have.has(name)) await run(`ALTER TABLE clienti ADD COLUMN ${name} ${type}`);
  }

  await run(`CREATE TABLE IF NOT EXISTS veicoli (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    cliente_id INTEGER,
    targa TEXT NOT NULL UNIQUE,
    marca TEXT, modello TEXT, versione TEXT, telaio TEXT, anno TEXT,
    km INTEGER DEFAULT 0, alimentazione TEXT, note TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  const vcols = await all(`PRAGMA table_info(veicoli)`);
  const vhave = new Set(vcols.map(x => x.name));
  for (const [name, type] of [
    ['versione','TEXT'],['anno','TEXT'],['alimentazione','TEXT'],
    ['note','TEXT'],['created_at','TEXT']
  ]) {
    if (!vhave.has(name)) await run(`ALTER TABLE veicoli ADD COLUMN ${name} ${type}`);
  }

  await run(`CREATE TABLE IF NOT EXISTS ordini_lavoro (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero INTEGER,
    anno INTEGER,
    cliente_id INTEGER,
    veicolo_id INTEGER,
    data_apertura TEXT,
    data_chiusura TEXT,
    km_ingresso INTEGER DEFAULT 0,
    km_uscita INTEGER DEFAULT 0,
    descrizione_lavoro TEXT,
    diagnosi TEXT,
    stato TEXT DEFAULT 'APERTO',
    totale REAL DEFAULT 0,
    note TEXT,
    whatsapp_pronto_at TEXT,
    fatturato INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  const ocols = await all(`PRAGMA table_info(ordini_lavoro)`);
  const ohave = new Set(ocols.map(x => x.name));
  for (const [name, type] of [
    ['diagnosi','TEXT'],['whatsapp_pronto_at','TEXT'],['fatturato','INTEGER DEFAULT 0'],['lavorazioni_eseguite','TEXT']
  ]) {
    if (!ohave.has(name)) await run(`ALTER TABLE ordini_lavoro ADD COLUMN ${name} ${type}`);
  }

  await run(`CREATE TABLE IF NOT EXISTS righe_lavoro (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ordine_id INTEGER NOT NULL,
    tipo TEXT DEFAULT 'RICAMBIO',
    descrizione TEXT NOT NULL,
    quantita REAL DEFAULT 1,
    prezzo_unitario REAL DEFAULT 0,
    iva REAL DEFAULT 22,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  await run(`CREATE TABLE IF NOT EXISTS ricambi (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codice TEXT,
    descrizione TEXT NOT NULL,
    categoria TEXT,
    marca TEXT,
    prezzo_acquisto REAL DEFAULT 0,
    prezzo_vendita REAL DEFAULT 0,
    iva REAL DEFAULT 22,
    giacenza REAL DEFAULT 0,
    note TEXT,
    attivo INTEGER DEFAULT 1,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  await run(`CREATE TABLE IF NOT EXISTS preventivi_service (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero INTEGER NOT NULL,
    anno INTEGER NOT NULL,
    ordine_id INTEGER NOT NULL UNIQUE,
    data TEXT NOT NULL,
    imponibile REAL DEFAULT 0,
    iva REAL DEFAULT 0,
    totale REAL DEFAULT 0,
    righe_json TEXT,
    note TEXT,
    stato TEXT DEFAULT 'EMESSO',
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);


  const pcols = await all(`PRAGMA table_info(preventivi_service)`);
  const phave = new Set(pcols.map(x=>x.name));
  for(const [name,type] of [
    ['token_cliente','TEXT'],['approvato_at','TEXT'],['rifiutato_at','TEXT'],['inviato_wa_at','TEXT'],['inviato_email_at','TEXT']
  ]) if(!phave.has(name)) await run(`ALTER TABLE preventivi_service ADD COLUMN ${name} ${type}`);

  await run(`CREATE TABLE IF NOT EXISTS fatture_service (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    numero INTEGER NOT NULL,
    serie TEXT DEFAULT 'S',
    anno INTEGER NOT NULL,
    ordine_id INTEGER NOT NULL UNIQUE,
    data TEXT NOT NULL,
    imponibile REAL DEFAULT 0,
    iva REAL DEFAULT 0,
    totale REAL DEFAULT 0,
    righe_json TEXT,
    pagamento TEXT DEFAULT 'Bonifico vista fattura',
    note TEXT,
    stato TEXT DEFAULT 'EMESSA',
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    UNIQUE(anno,numero,serie)
  )`);

  await run(`CREATE TABLE IF NOT EXISTS service_alerts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tipo TEXT DEFAULT 'RICHIESTA_CLIENTE',
    titolo TEXT,
    messaggio TEXT,
    ordine_id INTEGER,
    letto INTEGER DEFAULT 0,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  await run(`CREATE TABLE IF NOT EXISTS richieste_service (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ordine_id INTEGER,
    nome TEXT,
    telefono TEXT,
    targa TEXT,
    marca TEXT,
    modello TEXT,
    km INTEGER DEFAULT 0,
    richiesta TEXT,
    stato TEXT DEFAULT 'NUOVA',
    created_at TEXT DEFAULT CURRENT_TIMESTAMP
  )`);

  // V307 - sincronizzazione con ANAGRAFICA MASTER unica del DP Gestionale.
  try{
    const n=await dpServiceSyncMaster();
    console.log('DP SERVICE sync clienti master:',n);
  }catch(e){
    console.log('DP SERVICE sync clienti master non eseguito:',e.message);
  }

}


const DP_SERVICE_LISTINO_SEED = [{"categoria": "Liquidi/Additivi", "descrizione": "AD BLUE", "prezzo": 0.6, "note": ""}, {"categoria": "Liquidi/Additivi", "descrizione": "ADDITIVO CONSUMO OLIO TOTAL STOP OIL 300 ML", "prezzo": 52.0, "note": ""}, {"categoria": "Ricambi", "descrizione": "AMMORTIZZATORI RIMORCHIO FE380HC E FV130SN (PAR9019388)", "prezzo": 0.0, "note": ""}, {"categoria": "Liquidi/Additivi", "descrizione": "ANTIGELO BLU", "prezzo": 7.0, "note": ""}, {"categoria": "Liquidi/Additivi", "descrizione": "ANTIGELO GIALLO", "prezzo": 7.0, "note": ""}, {"categoria": "Liquidi/Additivi", "descrizione": "ANTIGELO ROSSO", "prezzo": 7.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA ECM 60AH 242X175X190 L2 B13", "prezzo": 92.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA EXCELL 100AH 313X175X205 LH4 B", "prezzo": 117.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA EXCELL 74 AH 278X175X190 L03 B1", "prezzo": 95.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA EXIDE 62 AH", "prezzo": 80.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA EXIDE 70AH ECM (START&STOP) 278X175X190 L3 B13", "prezzo": 130.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA EXIDE PREMIUM 53 AH 207X175X190 L01 B", "prezzo": 73.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA MIDAC 55 AH (S555.060.050)(H55)", "prezzo": 70.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA MIDAC 60 AH (560.059.054)(C60)", "prezzo": 77.87, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA MIDAC 60 AH AGM (S560.901.068) T2AGM", "prezzo": 150.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA MIDAC 60 AH EFB (S560.501.057)(IT2 EFB)", "prezzo": 100.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA MIDAC 72 AH EFB (S572.501.072)(IT3 EFB)", "prezzo": 120.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA MIDAC 74 AH (S574.012.068)(H74)", "prezzo": 94.26, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA MIDAC 95AH AGM (S595.901.085)(IT5 AGM)", "prezzo": 200.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA SAF 80 AH 680EN", "prezzo": 78.0, "note": ""}, {"categoria": "Batterie", "descrizione": "BATTERIA SAF 80 AH EFB", "prezzo": 115.0, "note": ""}, {"categoria": "Materiali", "descrizione": "BIADESIVO WURTH (12mm) AL MT.", "prezzo": 5.15, "note": ""}, {"categoria": "Segnaletica", "descrizione": "CONTRASSEGNO ADESIVO LIMITE VELOCITA' 100 KM/H", "prezzo": 5.0, "note": ""}, {"categoria": "Segnaletica", "descrizione": "CONTRASSEGNO ADESIVO LIMITE VELOCITA' 80 KM/H", "prezzo": 5.0, "note": ""}, {"categoria": "Servizi", "descrizione": "CONVERGENZA AUTO", "prezzo": 40.0, "note": ""}, {"categoria": "Servizi", "descrizione": "CONVERGENZA AUTOCARRO", "prezzo": 50.0, "note": ""}, {"categoria": "Ricambi", "descrizione": "FASCETTA SEMIASSE", "prezzo": 1.6, "note": ""}, {"categoria": "Filtri", "descrizione": "FILTRO ABITACOLO", "prezzo": 0.0, "note": "Prezzo da valorizzare"}, {"categoria": "Filtri", "descrizione": "FILTRO ARIA", "prezzo": 0.0, "note": "Prezzo da valorizzare"}, {"categoria": "Filtri", "descrizione": "FILTRO ESSICCATORE", "prezzo": 0.0, "note": "Prezzo da valorizzare"}, {"categoria": "Filtri", "descrizione": "FILTRO GASOLIO", "prezzo": 0.0, "note": "Prezzo da valorizzare"}, {"categoria": "Filtri", "descrizione": "FILTRO OLIO", "prezzo": 0.0, "note": "Prezzo da valorizzare"}, {"categoria": "Servizi", "descrizione": "INTERVENTO CARROATTREZZI", "prezzo": 50.0, "note": ""}, {"categoria": "Segnaletica", "descrizione": "KIT ADESIVO TABELLA MOTRICE (2 pz) A83899207 EUROPARTS", "prezzo": 16.0, "note": ""}, {"categoria": "Segnaletica", "descrizione": "KIT ADESIVO TABELLA RIMORCHIO (2 pz) A83899217 EUROPARTS", "prezzo": 23.0, "note": ""}, {"categoria": "Segnaletica", "descrizione": "KIT METALLICO TABELLA MOTRICE (2 pz) A83890094 EUROPARTS", "prezzo": 18.0, "note": ""}, {"categoria": "Segnaletica", "descrizione": "KIT METALLICO TABELLA RIMORCHIO (2 pz) A83890093 EUROPARTS", "prezzo": 25.0, "note": ""}, {"categoria": "Segnaletica", "descrizione": "KIT PANNELLI RITRORIFLETEX 11131 (LOMA MOR11.001)", "prezzo": 30.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA H1 12V (LOMA+DNG)", "prezzo": 2.75, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA 12V 21W 1 FIL", "prezzo": 2.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA 12V W21/5W 2 FIL", "prezzo": 4.4, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA 5W 24V", "prezzo": 2.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA ATT. VETRO 12V 5W (0720162 1)", "prezzo": 2.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA H1 12V 55W (0720111 1)", "prezzo": 5.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA H4 12V", "prezzo": 5.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA H4 12V/60/55W ALOGENA (0720 110 1)", "prezzo": 5.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA H4 24V (WURTH 07201102)", "prezzo": 6.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA H7 12V 55W (0720 114 1)", "prezzo": 5.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA H7 24V", "prezzo": 6.3, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA INDIC. ATTACCO VETRO ARANCIO WY5W 12V (0720 162 10)", "prezzo": 3.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA LED ARANCIO 12V (PLED) (A800200819) EUROPARTS", "prezzo": 6.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA LED BIANCA TUTTOVETRO 12V (A800200838) EUROPARTS", "prezzo": 2.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA LED ORIGINAL H9 12V 65W (OSRAM 64213)", "prezzo": 19.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA LED R5WLED (A800200832) EUROPARTS", "prezzo": 4.35, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA LED ROSSA 12V (P21LED) (A800200815) EUROPARTS", "prezzo": 5.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA ORIGINAL C5W 12V (LOMA OSR6418)", "prezzo": 0.9, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA ORIGINAL P21W 12V (0720 1342) DIREZIONE", "prezzo": 2.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA ORIGINAL R5W 24V (LOMA OSR 5627)", "prezzo": 1.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA ORIGINAL W5W 12V", "prezzo": 1.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA OSRAM ORIGINAL R5W 12V (OSR 5007 LOMA)", "prezzo": 1.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA P21/5W 12V/21/5W (0720 134 1)", "prezzo": 1.2, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA P21W 12V/21W (0720 132 1)", "prezzo": 1.05, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA POSIZIONE 24V 4W (0720 1502)", "prezzo": 2.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA PY21W 12V 21W (0720 138 3)", "prezzo": 3.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA PY21W 21W 24V (072013223)", "prezzo": 4.7, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA R10W 12V (0720 141 1)", "prezzo": 2.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA R5W 12V/5W (0720 140 1)", "prezzo": 1.2, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA TUTTOVETRO (POSIZIONE) W5W 12V 5W (0720 150 11)", "prezzo": 5.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA W16W 12V 16W (07201601)", "prezzo": 3.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA W16W 12V 16W (072016211)", "prezzo": 3.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA W5W 12V/5W ATTACCO VETRO/NON COLORATO (0720162 1)", "prezzo": 0.9, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA XENON OSRAM XENARC ORIGINAL D1S 35W (66140) LOMA", "prezzo": 71.5, "note": ""}, {"categoria": "Lampade", "descrizione": "LAMPADA LED H7 FAN KIT DA 2 (Azzurra)", "prezzo": 55.0, "note": ""}, {"categoria": "Lampade", "descrizione": "LEDRIVING HL EASY H7/H18 COOL WHITE 67000 (LOMA OSR 64210DWESY-2HB)", "prezzo": 43.0, "note": ""}, {"categoria": "Liquidi/Additivi", "descrizione": "LIQUIDO RADIATORE", "prezzo": 7.0, "note": ""}, {"categoria": "Segnaletica", "descrizione": "LUCE INGOMBRO GIALLA PIANA", "prezzo": 11.2, "note": ""}, {"categoria": "Servizi", "descrizione": "MANODOPERA", "prezzo": 30.0, "note": ""}, {"categoria": "Materiali", "descrizione": "MATERIALE D'USO", "prezzo": 5.0, "note": ""}, {"categoria": "Materiali", "descrizione": "MATERIALE D'USO E SMALTIMENTO", "prezzo": 10.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO 0W30 SHELL", "prezzo": 15.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO CAMBIO E DIFFERENZIALE 75W/90", "prezzo": 12.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO FRENI DOT 4", "prezzo": 7.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO IDRAULICO G 46", "prezzo": 6.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO IDROGUIDA", "prezzo": 11.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO MOTORE 0W20", "prezzo": 15.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO MOTORE 5W40", "prezzo": 8.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO MOTORE SELENIA 5W40 (SHE 48158)", "prezzo": 14.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO MOTORE SHELL 5W30", "prezzo": 12.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO MOTORE VITALTECH 15W40", "prezzo": 10.0, "note": ""}, {"categoria": "Oli/Lubrificanti", "descrizione": "OLIO TRUCK 10W40 SHELL", "prezzo": 10.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "PASTIGLIA TURAFALLE (DNG ARX35731)", "prezzo": 18.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "PASTIGLIA TURAFALLE (ROLIN ALUX)(DNG MAL35731)", "prezzo": 18.0, "note": ""}, {"categoria": "Servizi", "descrizione": "PREVISIONE", "prezzo": 22.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "PROLUNGA PLASTICA 18MM (PNEUMATICI)", "prezzo": 8.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "PROLUNGA VALVOLA PIEGATA OTTONE", "prezzo": 5.0, "note": ""}, {"categoria": "Liquidi/Additivi", "descrizione": "RADIATOR REPAIR HP 300 ML", "prezzo": 24.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "RAPPEZZO PNEUMATICO AUTO", "prezzo": 30.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "RAPPEZZO PNEUMATICO AUTOCARRO", "prezzo": 60.0, "note": ""}, {"categoria": "Servizi", "descrizione": "REVISIONE", "prezzo": 54.95, "note": ""}, {"categoria": "Servizi", "descrizione": "REVISIONE SPESE ART.15", "prezzo": 11.96, "note": ""}, {"categoria": "Pneumatici", "descrizione": "RIPARAZIONE PNEUMATICO CON STRISCIA", "prezzo": 10.0, "note": ""}, {"categoria": "Servizi", "descrizione": "SMALTIMENTO", "prezzo": 5.0, "note": ""}, {"categoria": "Servizi", "descrizione": "SOCCORSO SU STRADA", "prezzo": 50.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "SPILLI X PNEUMATICI", "prezzo": 0.7, "note": ""}, {"categoria": "Segnaletica", "descrizione": "TABELLA POST. MOTRICE (N°2)", "prezzo": 30.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "VALVOLA CERCHI GOMMA", "prezzo": 1.5, "note": ""}, {"categoria": "Pneumatici", "descrizione": "VALVOLA FLESSIBILE PNEUMATICI", "prezzo": 5.5, "note": ""}, {"categoria": "Pneumatici", "descrizione": "VALVOLA GOMMA X CERCHI TPMS LAUNCH (SGL)", "prezzo": 40.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "VALVOLA GOMMA X EZ SENSOR WURTH 0879961001", "prezzo": 13.0, "note": ""}, {"categoria": "Pneumatici", "descrizione": "VALVOLA MET. X CERCHI ALLUMINIO (LAUNCH LTR-01)", "prezzo": 40.0, "note": ""}];

async function dpServiceSeedListino(){
  try{
    const n=await get(`SELECT COUNT(*) n FROM ricambi`);
    if(Number(n?.n||0)>0) return;
    for(const x of DP_SERVICE_LISTINO_SEED){
      await run(`INSERT INTO ricambi(descrizione,categoria,prezzo_vendita,iva,note,attivo) VALUES(?,?,?,?,?,1)`,
        [x.descrizione,x.categoria,Number(x.prezzo)||0,22,x.note||'']);
    }
    console.log('DP SERVICE listino iniziale caricato:', DP_SERVICE_LISTINO_SEED.length);
  }catch(e){ console.log('DP SERVICE seed listino errore:',e.message); }
}

function esc(v=''){ return String(v??'').replace(/[&<>"']/g,s=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[s])); }

function page(title, body){
return `<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${esc(title)} - DP SERVICE</title>
<style>
*{box-sizing:border-box} body{margin:0;font-family:Arial,Helvetica,sans-serif;background:#f1f1f1;color:#171717}
header{background:#0d0d0f;color:#fff;padding:22px 28px;border-bottom:6px solid #e00000}
header{display:flex;align-items:center;gap:14px} header img{width:74px;height:50px;object-fit:cover;border-radius:10px;border:1px solid #333} header b{font-size:34px} header span{margin-left:4px;font-weight:700}
main{max-width:1180px;margin:28px auto;padding:0 18px}
.topnav{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px}
.muted{color:#666;font-size:14px}
.hero{background:linear-gradient(120deg,#111,#7c0000,#e00000);color:white;padding:28px;border-radius:26px}
.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:18px;margin-top:20px}
.card{display:block;text-decoration:none;color:#fff;background:#17181b;border-radius:22px;padding:28px;font-size:26px;font-weight:800;border-bottom:6px solid #e00000}
.card.red{background:linear-gradient(120deg,#ff1a1a,#b40000)} .card.blue{background:#102a59}
.box{background:#fff;border-radius:22px;padding:24px;margin-top:20px;box-shadow:0 5px 18px #0001}
.btn{display:inline-block;background:#e00000;color:white;border:0;border-radius:14px;padding:12px 18px;text-decoration:none;font-weight:800;cursor:pointer}
.btn.dark{background:#202124}.btn.green{background:#159447}
input,select,textarea{width:100%;padding:12px;border:1px solid #bbb;border-radius:10px;font-size:16px}
label{font-weight:700;display:block;margin:10px 0 5px}
table{width:100%;border-collapse:collapse;margin-top:16px;background:#fff}
th{background:#1d1d1f;color:white;padding:10px;text-align:left} td{padding:10px;border-bottom:1px solid #ddd;vertical-align:top}
.filters{display:grid;grid-template-columns:2fr 1fr 1fr auto;gap:10px;align-items:end}
.actions{display:flex;gap:8px;flex-wrap:wrap}
.alert-service{background:#fff3cd;border:2px solid #ffb300;border-radius:16px;padding:14px 16px;margin:14px 0;font-weight:800}
.badge-alert{display:inline-block;background:#e00000;color:#fff;border-radius:999px;padding:5px 10px;font-size:14px}
.public-wrap{max-width:760px;margin:0 auto}
.public-wrap .box{border-top:6px solid #e00000}
@media(max-width:780px){.grid,.filters{grid-template-columns:1fr}header span{display:block;margin:6px 0 0}.card{font-size:22px}}
</style></head><body><header><img src="/public/dp_service_logo.png" alt="DP SERVICE"><b>DP SERVICE</b><span>Officina • Veicoli • Ricambi • Fatturazione</span></header><main><div class="topnav"><a class="btn dark" href="/__DP_MAIN__">🏠 Menu DP Gestionale</a></div>${body}</main></body></html>`;
}


function pagePublic(title, body){
return `<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow,noarchive"><meta name="referrer" content="no-referrer">
<title>${esc(title)} - DP SERVICE</title>
<style>
*{box-sizing:border-box}body{margin:0;font-family:Arial,Helvetica,sans-serif;background:#f3f3f3;color:#171717}
header{background:#0d0d0f;color:#fff;padding:18px;border-bottom:5px solid #e00000;text-align:center}
header img{width:92px;height:62px;object-fit:cover;border-radius:12px;vertical-align:middle;margin-right:12px}
header b{font-size:30px;vertical-align:middle}
main{max-width:760px;margin:24px auto;padding:0 14px}.box{background:#fff;border-radius:22px;padding:22px;box-shadow:0 6px 24px #0002;border-top:6px solid #e00000}
.hero{background:linear-gradient(120deg,#111,#8a0000,#e00000);color:#fff;padding:22px;border-radius:20px;margin-bottom:18px}
label{font-weight:800;display:block;margin:13px 0 6px}input,textarea,select{width:100%;padding:13px;border:1px solid #bbb;border-radius:10px;font-size:16px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.btn{width:100%;background:#e00000;color:#fff;border:0;border-radius:13px;padding:15px;font-size:18px;font-weight:900;margin-top:16px}
.privacy{background:#f6f6f6;border-radius:12px;padding:12px;margin-top:14px;font-size:14px;line-height:1.4}.privacy input{width:auto;margin-right:8px}
.small{font-size:13px;color:#666}@media(max-width:650px){.grid{grid-template-columns:1fr}header b{font-size:25px}}
</style></head><body><header><img src="/service/public/dp_service_logo.png" alt="DP SERVICE"><b>DP SERVICE</b></header><main>${body}</main></body></html>`;
}



const DP_SERVICE_AZIENDA = {
  nome: 'Trasporti DP S.r.l. - DP SERVICE',
  indirizzo: 'Via Tuderte 466 - 05035 Narni (TR)',
  piva: '01385450554',
  cf: '01385450554',
  tel: '0744 817108',
  email: 'manutenzione@trasportidp.com',
  web: 'www.trasportidp.com',
  iban: 'IT78Q0200814413000104798294'
};

function dpEuro(v){ return '€ ' + (Number(v)||0).toFixed(2); }
function dpItDate(v){
  if(!v) return '';
  const s=String(v).slice(0,10).split('-');
  return s.length===3 ? `${s[2]}/${s[1]}/${s[0]}` : String(v);
}
function dpCalcRighe(righe){
  let imponibile=0, iva=0;
  for(const r of (righe||[])){
    const q=Number(r.quantita)||0, p=Number(r.prezzo_unitario)||0, aliq=Number(r.iva)||22;
    const imp=q*p; imponibile+=imp; iva+=imp*aliq/100;
  }
  return {imponibile, iva, totale:imponibile+iva};
}
function dpPhone(v){
  let p=String(v||'').replace(/\D/g,'');
  if(!p) return '';
  if(p.startsWith('0039')) p=p.slice(2);
  if(!p.startsWith('39')) p='39'+p;
  return p;
}
function dpBaseUrl(req){ return (process.env.APP_BASE_URL || `${req.protocol}://${req.get('host')}`).replace(/\/$/,''); }

function dpServiceNormWa(v){
  let p=String(v||'').trim().replace(/^whatsapp:/i,'').replace(/[^\d+]/g,'');
  if(!p) return '';
  if(p.startsWith('0039')) p='+'+p.slice(2);
  else if(!p.startsWith('+')) p=p.startsWith('39')?('+'+p):('+39'+p);
  return 'whatsapp:'+p;
}
function dpServiceStaffNumbers(){
  const raw=process.env.INTERNAL_OFFICINA_NUMBERS || process.env.INTERNAL_GENERAL_NUMBERS || process.env.STAFF_WHATSAPP_NUMBERS || '';
  const fallback=['+393287377675','+393472733226','+393494040073'];
  return (raw?String(raw).split(/[;,\n]+/):fallback).map(dpServiceNormWa).filter(Boolean);
}
const dpServiceTwilioClient=(twilio && process.env.TWILIO_ACCOUNT_SID && process.env.TWILIO_AUTH_TOKEN)
  ? twilio(process.env.TWILIO_ACCOUNT_SID,process.env.TWILIO_AUTH_TOKEN) : null;

async function dpServiceNotifyStaff(body){
  if(!dpServiceTwilioClient) return {ok:false,error:'Twilio non configurato'};
  const from=dpServiceNormWa(process.env.TWILIO_WHATSAPP_NUMBER || '+390744817108');
  let sent=0, errors=[];
  for(const to of [...new Set(dpServiceStaffNumbers())]){
    try{
      await dpServiceTwilioClient.messages.create({from,to,body:String(body||'')});
      sent++;
    }catch(e){ errors.push(e.message||String(e)); }
  }
  return {ok:sent>0,sent,errors};
}


function dpServiceUrl(req, rel=''){
  const base=(process.env.APP_BASE_URL || `${req.protocol}://${req.get('host')}`).replace(/\/$/,'');
  const p=String(rel||'').startsWith('/')?String(rel):('/'+String(rel||''));
  return base + '/service' + p;
}
function dpServiceToken(){
  return require('crypto').randomBytes(24).toString('hex');
}
async function dpServiceSendEmail(to,subject,text,attachments=[]){
  if(!to) throw new Error('Email cliente mancante');
  if(!nodemailer || !process.env.SMTP_HOST) throw new Error('SMTP non configurato');
  const transporter=nodemailer.createTransport({
    host:process.env.SMTP_HOST,
    port:Number(process.env.SMTP_PORT||587),
    secure:String(process.env.SMTP_SECURE||'').toLowerCase()==='true' || Number(process.env.SMTP_PORT||587)===465,
    auth:{user:process.env.SMTP_USER,pass:process.env.SMTP_PASS},
    tls:{rejectUnauthorized:false}
  });
  return transporter.sendMail({
    from:process.env.SMTP_FROM || process.env.SMTP_USER || DP_SERVICE_AZIENDA.email,
    to,subject,text,attachments
  });
}
function dpServicePdfBuffer(tipo,numero,data,d,opts={}){
  return new Promise((resolve,reject)=>{
    try{
      const chunks=[];
      const doc=new PDFDocument({size:'A4',margin:0,bufferPages:true});
      doc.on('data',c=>chunks.push(c));
      doc.on('end',()=>resolve(Buffer.concat(chunks)));
      doc.on('error',reject);
      dpDrawServicePdf(doc,tipo,numero,data,d,opts);
      doc.end();
    }catch(e){ reject(e); }
  });
}

async function dpServiceDocData(ordineId){
  const o=await get(`SELECT o.*,v.targa,v.marca,v.modello,v.versione,v.telaio,
      c.ragione_sociale,c.piva,c.cf,c.indirizzo,c.citta,c.provincia,c.telefono,c.email,c.pec,c.sdi
    FROM ordini_lavoro o
    LEFT JOIN veicoli v ON v.id=o.veicolo_id
    LEFT JOIN clienti c ON c.id=o.cliente_id
    WHERE o.id=?`,[ordineId]);
  if(!o) return null;
  const righe=await all('SELECT * FROM righe_lavoro WHERE ordine_id=? ORDER BY id',[ordineId]);
  return {o,righe,calc:dpCalcRighe(righe)};
}

function dpPdfHeader(doc, tipo, numeroLabel, dataLabel){
  const W=doc.page.width;
  const H=172;
  doc.rect(0,0,W,H).fill('#0b0b0d');

  const logoPath=path.join(__dirname,'public','dp_service_logo.png');
  if(fs.existsSync(logoPath)){
    try{
      // Logo più piccolo e pulito, niente effetto poster enorme.
      doc.image(logoPath,38,26,{fit:[145,105],align:'center',valign:'center'});
    }catch(e){}
  }

  // Titolo centrale.
  doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(25)
    .text('DP SERVICE',205,34,{width:210});
  doc.fillColor('#ff2525').font('Helvetica-Bold').fontSize(10.5)
    .text('OFFICINA MULTIMARCA',205,73,{width:210,characterSpacing:.7});
  doc.fillColor('#d9d9d9').font('Helvetica').fontSize(8.8)
    .text('MANUTENZIONE • RIPARAZIONI • DIAGNOSI\nPNEUMATICI • ELETTRAUTO',205,96,{width:210,lineGap:3});

  // Tricolore elegante, stretto sul bordo destro.
  const stripeX=W-58;
  doc.rect(stripeX,0,19,H).fill('#169b62');
  doc.rect(stripeX+19,0,19,H).fill('#ffffff');
  doc.rect(stripeX+38,0,20,H).fill('#d71920');

  // Box documento separato dalle scritte.
  const isPrev=String(tipo||'').toUpperCase()==='PREVENTIVO';
  const bx=W-188, by=32, bw=112, bh=58;
  doc.roundedRect(bx,by,bw,bh,9).fill(isPrev?'#8DFF00':'#e20d17');
  doc.fillColor(isPrev?'#111111':'#ffffff').font('Helvetica-Bold').fontSize(14)
    .text(isPrev?'PREVENTIVO':'FATTURA',bx+7,by+17,{width:bw-14,align:'center'});

  doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(11.5)
    .text(numeroLabel,bx,106,{width:bw,align:'center'});
  doc.fillColor('#d6d6d6').font('Helvetica').fontSize(8.5)
    .text(dataLabel,bx,125,{width:bw,align:'center'});

  // Linea rossa inferiore.
  doc.rect(0,H-5,W,5).fill('#e20d17');
}

function dpPdfFooter(doc){
  const W=doc.page.width,H=doc.page.height;
  doc.rect(0,H-54,W,54).fill('#0b0c0e');
  doc.rect(0,H-54,7,54).fill('#d40000');
  doc.rect(W-80,H-54,26,54).fill('#159447');
  doc.rect(W-54,H-54,27,54).fill('#ffffff');
  doc.rect(W-27,H-54,27,54).fill('#e31b23');
  doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(8).text('DP SERVICE  •  LA TUA AUTO IN BUONE MANI',42,H-39);
  doc.font('Helvetica').fontSize(7.5).text(`${DP_SERVICE_AZIENDA.tel}  •  ${DP_SERVICE_AZIENDA.email}  •  ${DP_SERVICE_AZIENDA.web}`,42,H-25);
}

function dpDrawServicePdf(doc, tipo, numeroLabel, dataLabel, d, extra={}){
  dpPdfHeader(doc,tipo,numeroLabel,dataLabel);
  const {o,righe,calc}=d;
  let y=188, L=42, W=doc.page.width-84;
  doc.fillColor('#151515').font('Helvetica-Bold').fontSize(11).text(DP_SERVICE_AZIENDA.nome,L,y);
  doc.font('Helvetica').fontSize(8.5).text(`${DP_SERVICE_AZIENDA.indirizzo}  |  P.IVA ${DP_SERVICE_AZIENDA.piva}  |  Tel. ${DP_SERVICE_AZIENDA.tel}`,L,y+17);
  doc.fillColor('#d70b12').font('Helvetica-Bold').fontSize(8).text(`Email: ${DP_SERVICE_AZIENDA.email}`,L,y+30);
  y+=58;
  const bw=(W-12)/2;
  doc.roundedRect(L,y,bw,102,7).lineWidth(1).strokeColor('#d0d0d0').stroke();
  doc.rect(L,y,bw,24).fill('#d70b12');
  doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(9).text('CLIENTE / DESTINATARIO',L+10,y+7);
  doc.fillColor('#111111').font('Helvetica-Bold').fontSize(10).text(o.ragione_sociale||'',L+12,y+34,{width:bw-24});
  doc.font('Helvetica').fontSize(8.5).text(`${o.indirizzo||''} ${o.citta||''} ${o.provincia||''}`,L+12,y+51,{width:bw-24});
  doc.text(`P.IVA: ${o.piva||'-'}   C.F.: ${o.cf||'-'}`,L+12,y+68,{width:bw-24});
  doc.text(`SDI: ${o.sdi||'-'}   PEC: ${o.pec||'-'}`,L+12,y+83,{width:bw-24});

  const R=L+bw+12;
  doc.roundedRect(R,y,bw,102,7).lineWidth(1).strokeColor('#d0d0d0').stroke();
  doc.rect(R,y,bw,24).fill('#17181a');
  doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(9).text('DATI VEICOLO / LAVORAZIONE',R+10,y+7);
  doc.fillColor('#111111').font('Helvetica-Bold').fontSize(10).text(`${o.marca||''} ${o.modello||''}`,R+12,y+34,{width:bw-24});
  doc.font('Helvetica').fontSize(8.5).text(`Targa: ${o.targa||'-'}   Km: ${o.km_ingresso||0}`,R+12,y+51);
  doc.text(`Versione: ${o.versione||'-'}`,R+12,y+68,{width:bw-24});
  doc.text(`ODL: ${o.numero||o.id}/${o.anno||''}   Ingresso: ${dpItDate(o.data_apertura)}`,R+12,y+83,{width:bw-24});
  y+=120;

  if(o.descrizione_lavoro){
    doc.fillColor('#d70b12').font('Helvetica-Bold').fontSize(9).text('LAVORO RICHIESTO',L,y);
    doc.fillColor('#222').font('Helvetica').fontSize(8.5).text(o.descrizione_lavoro,L,y+13,{width:W});
    y+=38;
  }

  const widths=[58,205,46,68,42,92];
  const headers=['TIPO','DESCRIZIONE','Q.TÀ','PREZZO','IVA','TOTALE'];
  function tableHeader(){
    doc.rect(L,y,W,22).fill('#17181a');
    let x=L;
    headers.forEach((h,i)=>{doc.fillColor('#fff').font('Helvetica-Bold').fontSize(7.5).text(h,x+5,y+7,{width:widths[i]-10,align:i>=2?'right':'left'});x+=widths[i];});
    y+=22;
  }
  tableHeader();
  for(const r of righe){
    if(y>684){ dpPdfFooter(doc); doc.addPage(); dpPdfHeader(doc,tipo,numeroLabel,dataLabel); y=184; tableHeader(); }
    const imp=(Number(r.quantita)||0)*(Number(r.prezzo_unitario)||0);
    const vals=[r.tipo||'',r.descrizione||'',String(Number(r.quantita)||0),dpEuro(r.prezzo_unitario),`${Number(r.iva)||22}%`,dpEuro(imp)];
    const rh=30;
    doc.rect(L,y,W,rh).fillAndStroke('#ffffff','#dddddd');
    let x=L;
    vals.forEach((v,i)=>{doc.fillColor('#151515').font(i===1?'Helvetica-Bold':'Helvetica').fontSize(7.8).text(v,x+5,y+8,{width:widths[i]-10,align:i>=2?'right':'left',ellipsis:true});x+=widths[i];});
    y+=rh;
  }
  y+=16;
  if(y>620){ dpPdfFooter(doc); doc.addPage(); dpPdfHeader(doc,tipo,numeroLabel,dataLabel); y=196; }
  const tx=L+W-235;
  doc.roundedRect(tx,y,235,92,6).fillAndStroke('#f5f5f5','#cccccc');
  doc.fillColor('#222').font('Helvetica').fontSize(9).text('Imponibile',tx+12,y+14); doc.text(dpEuro(calc.imponibile),tx+118,y+14,{width:102,align:'right'});
  doc.text('IVA',tx+12,y+34); doc.text(dpEuro(calc.iva),tx+118,y+34,{width:102,align:'right'});
  doc.rect(tx,y+54,235,38).fill('#d70b12');
  doc.fillColor('#fff').font('Helvetica-Bold').fontSize(11).text('TOTALE',tx+12,y+67); doc.fontSize(16).text(dpEuro(calc.totale),tx+105,y+63,{width:115,align:'right'});
  if(extra.pagamento){
    doc.fillColor('#111').font('Helvetica-Bold').fontSize(8.5).text('PAGAMENTO',L,y+8);
    doc.font('Helvetica').text(extra.pagamento,L,y+24,{width:tx-L-15});
    doc.font('Helvetica-Bold').text('IBAN',L,y+44); doc.font('Helvetica').text(DP_SERVICE_AZIENDA.iban,L,y+59,{width:tx-L-15});
  }else{
    doc.fillColor('#d70b12').font('Helvetica-BoldOblique').fontSize(15).text('La tua auto in buone mani.',L,y+30,{width:tx-L-15});
  }
  dpPdfFooter(doc);
}


const DP_MECC_COOKIE='dp_service_meccanico';
const DP_MECC_MAX_AGE=10*60*60*1000;
function dpMeccCookies(req){
  const out={}; String(req.headers.cookie||'').split(';').forEach(p=>{const i=p.indexOf('=');if(i>0)out[p.slice(0,i).trim()]=decodeURIComponent(p.slice(i+1).trim());}); return out;
}
function dpMeccSecret(){ return String(process.env.DP_SERVICE_MECCANICI_SECRET || process.env.DP_ADMIN_SECRET || ''); }
function dpMeccPin(){ return String(process.env.DP_SERVICE_MECCANICI_PIN || ''); }
function dpMeccSign(v){ return crypto.createHmac('sha256',dpMeccSecret()).update(v).digest('hex'); }
function dpMeccToken(){
  const p=Buffer.from(JSON.stringify({role:'meccanico',exp:Date.now()+DP_MECC_MAX_AGE})).toString('base64url');
  return p+'.'+dpMeccSign(p);
}
function dpMeccOk(req){
  if(!dpMeccPin() || !dpMeccSecret()) return false;
  const t=dpMeccCookies(req)[DP_MECC_COOKIE]; if(!t||!t.includes('.')) return false;
  const i=t.lastIndexOf('.'),p=t.slice(0,i),sig=t.slice(i+1);
  try{
    const a=Buffer.from(sig),b=Buffer.from(dpMeccSign(p)); if(a.length!==b.length || !crypto.timingSafeEqual(a,b)) return false;
    const x=JSON.parse(Buffer.from(p,'base64url').toString('utf8')); return x.role==='meccanico' && Number(x.exp)>Date.now();
  }catch(_){ return false; }
}
function dpMeccSet(res){
  const secure=String(process.env.NODE_ENV||'').toLowerCase()==='production';
  res.setHeader('Set-Cookie',`${DP_MECC_COOKIE}=${encodeURIComponent(dpMeccToken())}; Path=/service/meccanici; HttpOnly; SameSite=Strict; Max-Age=${Math.floor(DP_MECC_MAX_AGE/1000)}${secure?'; Secure':''}`);
}
function dpMeccClear(res){
  res.setHeader('Set-Cookie',`${DP_MECC_COOKIE}=; Path=/service/meccanici; HttpOnly; SameSite=Strict; Max-Age=0`);
}
function pageMeccanico(title,body){
return `<!doctype html><html lang="it"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow,noarchive"><title>${esc(title)} - DP SERVICE</title>
<style>*{box-sizing:border-box}body{font-family:Arial;margin:0;background:#111;color:#eee}header{padding:20px;background:#090909;border-bottom:5px solid #e00000}main{max-width:900px;margin:auto;padding:18px}.box{background:#1d1d1f;border-radius:18px;padding:18px;margin:14px 0}.btn{display:inline-block;background:#e00000;color:#fff;border:0;border-radius:12px;padding:12px 16px;font-weight:800;text-decoration:none}.dark{background:#333}input,textarea,select{width:100%;padding:13px;border:1px solid #555;background:#111;color:#fff;border-radius:10px;font-size:16px}label{display:block;font-weight:800;margin:12px 0 5px}.order{border-left:5px solid #e00000}.muted{color:#aaa;font-size:13px}</style></head><body><header><b>🔧 DP SERVICE — AREA MECCANICI</b></header><main>${body}</main></body></html>`;
}

router.get('/meccanici/login',(req,res)=>{
  res.setHeader('Cache-Control','no-store');
  if(dpMeccOk(req)) return res.redirect('/meccanici');
  const enabled=!!(dpMeccPin()&&dpMeccSecret());
  res.status(enabled?200:503).send(pageMeccanico('Accesso meccanici',`<div class="box"><h1>Accesso meccanici</h1>${enabled?`<form method="post" action="/service/meccanici/login"><label>PIN officina</label><input name="pin" type="password" inputmode="numeric" required><button class="btn" style="margin-top:15px">ENTRA</button></form>`:`<p>Portale non ancora configurato. Impostare DP_SERVICE_MECCANICI_PIN e DP_SERVICE_MECCANICI_SECRET su Render.</p>`}<p class="muted">Quest'area non mostra clienti, fatture, preventivi o prezzi.</p></div>`));
});
router.post('/meccanici/login',(req,res)=>{
  if(!dpMeccPin() || !dpMeccSecret()) return res.status(503).send('Portale meccanici non configurato');
  const a=Buffer.from(String(req.body.pin||'')),b=Buffer.from(dpMeccPin());
  if(a.length!==b.length || !crypto.timingSafeEqual(a,b)) return res.status(401).send(pageMeccanico('Accesso negato',`<div class="box"><h1>PIN non corretto</h1><a class="btn" href="/service/meccanici/login">Riprova</a></div>`));
  dpMeccSet(res); res.redirect('/meccanici');
});
router.get('/meccanici/esci',(req,res)=>{dpMeccClear(res);res.redirect('/meccanici/login')});
router.use('/meccanici',(req,res,next)=>{res.setHeader('Cache-Control','no-store');res.setHeader('X-Robots-Tag','noindex,nofollow,noarchive');if(dpMeccOk(req))return next();res.redirect('/service/meccanici/login')});

router.get('/meccanici',async(req,res)=>{
  const rows=await all(`SELECT o.id,o.numero,o.anno,o.stato,o.descrizione_lavoro,o.lavorazioni_eseguite,o.data_apertura,
    v.targa,v.marca,v.modello
    FROM ordini_lavoro o LEFT JOIN veicoli v ON v.id=o.veicolo_id
    WHERE o.stato NOT IN ('CHIUSO') ORDER BY CASE o.stato WHEN 'APERTO' THEN 0 WHEN 'IN_LAVORAZIONE' THEN 1 WHEN 'ATTESA_RICAMBI' THEN 2 WHEN 'PRONTO' THEN 3 ELSE 4 END,o.id DESC LIMIT 200`);
  res.send(pageMeccanico('Lavorazioni',`<div style="display:flex;justify-content:space-between;align-items:center"><h1>Lavorazioni officina</h1><a class="btn dark" href="/service/meccanici/esci">Esci</a></div>
    ${rows.map(o=>`<div class="box order"><h2>${esc(o.targa||'-')} — ${esc(o.marca||'')} ${esc(o.modello||'')}</h2><p><b>ODL ${o.numero||o.id}/${o.anno||''}</b> · ${esc(o.stato||'')}</p><p><b>Richiesta:</b> ${esc(o.descrizione_lavoro||'')}</p><a class="btn" href="/service/meccanici/ordine/${o.id}">Apri lavorazione</a></div>`).join('')||'<div class="box">Nessun ordine aperto.</div>'}`));
});
router.get('/meccanici/ordine/:id',async(req,res)=>{
  const o=await get(`SELECT o.id,o.numero,o.anno,o.stato,o.descrizione_lavoro,o.lavorazioni_eseguite,o.data_apertura,o.km_ingresso,
    v.targa,v.marca,v.modello
    FROM ordini_lavoro o LEFT JOIN veicoli v ON v.id=o.veicolo_id WHERE o.id=?`,[req.params.id]);
  if(!o)return res.status(404).send('Ordine non trovato');
  res.send(pageMeccanico('Lavorazione',`<a class="btn dark" href="/service/meccanici">← Elenco</a><div class="box">
    <h1>${esc(o.targa||'-')} — ${esc(o.marca||'')} ${esc(o.modello||'')}</h1><p><b>ODL ${o.numero||o.id}/${o.anno||''}</b></p>
    <p><b>Lavoro richiesto:</b><br>${esc(o.descrizione_lavoro||'')}</p>
    <form method="post" action="/service/meccanici/ordine/${o.id}">
      <label>Lavorazioni eseguite / note tecniche</label><textarea name="lavorazioni_eseguite" rows="9" required>${esc(o.lavorazioni_eseguite||'')}</textarea>
      <label>Stato lavorazione</label><select name="stato"><option ${o.stato==='APERTO'?'selected':''}>APERTO</option><option ${o.stato==='IN_LAVORAZIONE'?'selected':''}>IN_LAVORAZIONE</option><option ${o.stato==='ATTESA_RICAMBI'?'selected':''}>ATTESA_RICAMBI</option><option ${o.stato==='PRONTO'?'selected':''}>PRONTO</option></select>
      <button class="btn" style="margin-top:15px">💾 SALVA LAVORAZIONE</button>
    </form><p class="muted">Non sono visibili dati cliente, prezzi, preventivi, fatture o altre sezioni del gestionale.</p></div>`));
});
router.post('/meccanici/ordine/:id',async(req,res)=>{
  const stato=['APERTO','IN_LAVORAZIONE','ATTESA_RICAMBI','PRONTO'].includes(String(req.body.stato||''))?String(req.body.stato):'IN_LAVORAZIONE';
  await run(`UPDATE ordini_lavoro SET lavorazioni_eseguite=?,stato=? WHERE id=?`,[String(req.body.lavorazioni_eseguite||'').trim(),stato,req.params.id]);
  if(stato==='PRONTO'){
    await run(`INSERT INTO service_alerts(tipo,titolo,messaggio,ordine_id) VALUES('MECCANICO_PRONTO','Veicolo pronto','Il meccanico ha segnato il veicolo come PRONTO.',?)`,[req.params.id]);
  }
  res.redirect('/meccanici/ordine/'+req.params.id);
});


router.get('/', async (req,res)=>{
  let c=await mainGet('SELECT COUNT(*) n FROM dp_clienti_master').catch(()=>null);
  if(!c) c=await get('SELECT COUNT(*) n FROM clienti');
  const v=await get('SELECT COUNT(*) n FROM veicoli');
  const o=await get("SELECT COUNT(*) n FROM ordini_lavoro WHERE stato IN ('APERTO','IN_LAVORAZIONE','ATTESA_RICAMBI')");
  const r=await get("SELECT COUNT(*) n FROM ricambi WHERE attivo=1");
  const pv=await get('SELECT COUNT(*) n FROM preventivi_service');
  const fsrv=await get('SELECT COUNT(*) n FROM fatture_service');
  const al=await get('SELECT COUNT(*) n FROM service_alerts WHERE letto=0');
  await dpServiceSeedListino();
  res.send(page('Dashboard',`
    <div class="hero"><h1 style="margin:0;font-size:44px">DP SERVICE</h1><div style="font-size:20px;font-weight:700">Gestionale Officina</div></div>
    <div class="grid">
      <a class="card red" href="/clienti-azienda">👥 Clienti<br><small>${c.n} anagrafiche • archivio unico</small></a>
      <a class="card blue" href="/veicoli">🚗 Veicoli clienti<br><small>${v.n} veicoli</small></a>
      <a class="card" href="/ricerca">🔎 Ricerca globale<br><small>Targa • Cliente • Modello</small></a>
      <a class="card" href="/ordini">🧾 Ordini di lavoro<br><small>${o.n} aperti</small></a>
      <a class="card" href="/ricambi">📦 Ricambi / Listino<br><small>${r.n} voci</small></a>
      <a class="card" href="/preventivi">📄 Preventivi<br><small>${pv.n} emessi</small></a>
      <a class="card" href="/fatture">💶 Fatture serie S<br><small>${fsrv.n} emesse</small></a>
      <a class="card ${Number(al.n)>0?'red':''}" href="/alert">🔔 Avvisi officina<br><small>${al.n} da leggere</small></a>
      <a class="card blue" href="/richiesta">📲 Pagina richiesta cliente<br><small>Link pubblico officina</small></a><div class="card" style="grid-column:1/-1;cursor:default"><b>⚡ FLUSSO RAPIDO DP SERVICE</b><br><small>Richiesta cliente → Ordine di lavoro → Ricambi/Manodopera → Preventivo → Approvazione cliente online → Lavorazione → Fattura → WhatsApp/Email</small></div>
    </div>`));
});

router.get('/clienti', (req,res)=>res.redirect('/clienti-azienda'));
router.get('/clienti', async (req,res)=>{
  const q=(req.query.q||'').trim();
  const rows=q
    ? await all(`SELECT * FROM clienti WHERE ragione_sociale LIKE ? OR piva LIKE ? OR cf LIKE ? OR telefono LIKE ? OR citta LIKE ? ORDER BY ragione_sociale LIMIT 500`,
      Array(5).fill(`%${q}%`))
    : await all(`SELECT * FROM clienti ORDER BY ragione_sociale LIMIT 500`);
  res.send(page('Clienti',`
    <div class="actions"><a class="btn dark" href="/__DP_MAIN__">🏠 Menu DP Gestionale</a><a class="btn" href="/clienti/nuovo">+ Nuovo cliente</a></div>
    <div class="box"><h1>👥 Clienti (${rows.length})</h1>
      <form class="filters"><input name="q" placeholder="Cliente, P.IVA, C.F., telefono, città" value="${esc(q)}"><span></span><span></span><button class="btn">Cerca</button></form>
      <p class="muted"><b>Anagrafica condivisa:</b> gli stessi clienti di DP RENT e DP TRASPORTI.</p>
      <table><tr><th>Cliente</th><th>P.IVA / C.F.</th><th>Località</th><th>Telefono</th><th>IVA</th><th>Pagamento</th><th>IBAN</th><th></th></tr>
      ${rows.map(x=>`<tr><td><b>${esc(x.ragione_sociale)}</b></td><td>${esc(x.piva)}<br>${esc(x.cf)}</td><td>${esc(x.indirizzo)}<br>${esc(x.citta)} ${esc(x.provincia)}</td><td>${esc(x.telefono)}</td><td>${esc(x.codice_iva||'-')}</td><td>${esc(x.codice_pagamento||'-')}</td><td>${esc(x.banca_iban||'-')}</td><td><a class="btn dark" href="/clienti/${x.id}">Apri</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/clienti/nuovo',(req,res)=>{const CLIENTE={};res.send(page('Nuovo cliente',`
  <div class="actions"><a class="btn dark" href="/clienti">Indietro</a></div>
  <div class="box"><h1>+ Nuovo cliente</h1><p class="muted">Verrà salvato nell'anagrafica unica DP Gestionale.</p>
  <form method="post" action="/clienti">
    <label>Ragione sociale / Nome</label><input name="ragione_sociale" required>
    <label>P.IVA</label><input name="piva">
    <label>Codice fiscale</label><input name="cf">
    <label>Indirizzo</label><input name="indirizzo">
    <label>Città</label><input name="citta">
    <label>Provincia</label><input name="provincia">
    <label>Telefono</label><input name="telefono">
    <label>Email</label><input name="email">
    <label>PEC</label><input name="pec">
    <label>SDI</label><input name="sdi">

    <h3>🧾 IVA e pagamento</h3>
    <label>Codice IVA</label><select name="codice_iva">
      <option value="">Scegli...</option>
      ${[['22','IVA 22%'],['22SP','IVA 22% SPLIT PAYMENT'],['22RC','IVA 22% REVERSE CHARGE'],['22EX','IVA 22% EXTRA UE'],['10','IVA 10%'],['4','IVA 4%'],['5','IVA 5%'],['NI41','Non imponibile Art. 41 Cessioni UE'],['NI8','Non imponibile Art. 8 c.1 lett. c'],['NI8b','Non imponibile Art. 8 bis'],['FC7','Art. 7 TER operazione non soggetta']].map(x=>`<option value="${x[0]}" ${String(CLIENTE.codice_iva||'')===x[0]?'selected':''}>${x[0]} - ${x[1]}</option>`).join('')}
    </select>
    <label>Condizione di pagamento</label><select name="codice_pagamento">
      <option value="">Scegli...</option>
      ${[['BO30','Bonifico 30 gg'],['BO6','Bonifico 60 gg'],['BO9','Bonifico 90 gg'],['BOIMM','Bonifico vista fattura'],['RB3','Ri.Ba. 30 gg DF FM'],['RB36','Ri.Ba. 30/60 gg DF FM'],['RB6','Ri.Ba. 60 gg DF FM'],['RB9','Ri.Ba. 90 gg DF FM'],['RBF10','Ri.Ba. fine mese +10'],['RBFIN','Ri.Ba. fine mese'],['RID30','Rimessa diretta 30 gg'],['RID60','Rimessa diretta 60 gg'],['RID90','Rimessa diretta 90 gg FM'],['RD','Rimessa diretta a vista']].map(x=>`<option value="${x[0]}" ${String(CLIENTE.codice_pagamento||'')===x[0]?'selected':''}>${x[0]} - ${x[1]}</option>`).join('')}
    </select>
    <h3>🏦 Dati bancari cliente</h3>
    <label>Banca cliente</label><input name="banca_cliente" value="${esc(CLIENTE.banca_cliente||'')}">
    <label>Codice banca</label><input name="banca_codice" value="${esc(CLIENTE.banca_codice||'')}">
    <label>Numero conto corrente</label><input name="banca_conto" value="${esc(CLIENTE.banca_conto||'')}">
    <label>ABI</label><input name="banca_abi" value="${esc(CLIENTE.banca_abi||'')}"><label>CAB</label><input name="banca_cab" value="${esc(CLIENTE.banca_cab||'')}">
    <label>Paese</label><input name="banca_paese" value="${esc(CLIENTE.banca_paese||'IT')}">
    <label>CIN EUR</label><input name="banca_cin_eur" value="${esc(CLIENTE.banca_cin_eur||'')}"><label>CIN Italia</label><input name="banca_cin_it" value="${esc(CLIENTE.banca_cin_it||'')}">
    <label>Valuta</label><input name="banca_valuta" value="${esc(CLIENTE.banca_valuta||'EUR')}"><label>BIC / SWIFT</label><input name="banca_bic" value="${esc(CLIENTE.banca_bic||'')}">
    <label>IBAN</label><input name="banca_iban" value="${esc(CLIENTE.banca_iban||'')}"><label>BBAN</label><input name="banca_bban" value="${esc(CLIENTE.banca_bban||'')}">
    <label>PEC banca</label><input name="banca_pec" value="${esc(CLIENTE.banca_pec||'')}"><label>Note banca</label><input name="banca_note" value="${esc(CLIENTE.banca_note||'')}">

    <label>Note</label><textarea name="note"></textarea>
    <p><button class="btn">Salva cliente</button></p>
  </form></div>`));});

router.post('/clienti', async (req,res)=>{
  const b=req.body;
  const r=await run(`INSERT INTO clienti(ragione_sociale,piva,cf,indirizzo,citta,provincia,telefono,email,pec,sdi,codice_iva,codice_pagamento,banca_cliente,banca_codice,banca_conto,banca_abi,banca_cab,banca_paese,banca_cin_eur,banca_cin_it,banca_valuta,banca_bic,banca_iban,banca_bban,banca_pec,banca_note,note) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`,
    [b.ragione_sociale,b.piva,b.cf,b.indirizzo,b.citta,b.provincia,b.telefono,b.email,b.pec,b.sdi,b.codice_iva,b.codice_pagamento,b.banca_cliente,b.banca_codice,b.banca_conto,b.banca_abi,b.banca_cab,b.banca_paese,b.banca_cin_eur,b.banca_cin_it,b.banca_valuta,b.banca_bic,b.banca_iban,b.banca_bban,b.banca_pec,b.banca_note,b.note]);
  await dpServiceSaveToMaster(r.lastID,b);
  res.redirect('/clienti/'+r.lastID);
});

router.get('/clienti/:id', async (req,res)=>{
  const c=await get('SELECT * FROM clienti WHERE id=?',[req.params.id]);
  if(!c) return res.status(404).send('Cliente non trovato');
  const vs=await all('SELECT * FROM veicoli WHERE cliente_id=? ORDER BY targa',[c.id]);
  res.send(page(c.ragione_sociale,`
    <div class="actions"><a class="btn dark" href="/clienti">Indietro</a><a class="btn" href="/clienti/${c.id}/modifica">Modifica cliente</a><a class="btn green" href="/veicoli/nuovo?cliente_id=${c.id}">+ Aggiungi veicolo</a></div>
    <div class="box"><h1>${esc(c.ragione_sociale)}</h1>
      <p><b>P.IVA:</b> ${esc(c.piva)} &nbsp; <b>C.F.:</b> ${esc(c.cf)}</p>
      <p><b>Indirizzo:</b> ${esc(c.indirizzo)} ${esc(c.citta)} ${esc(c.provincia)}</p>
      <p><b>Telefono:</b> ${esc(c.telefono)} &nbsp; <b>Email:</b> ${esc(c.email)}</p>
      <p><b>PEC:</b> ${esc(c.pec)} &nbsp; <b>SDI:</b> ${esc(c.sdi)}</p>
      <p><b>Codice IVA:</b> ${esc(c.codice_iva||'-')} &nbsp; <b>Pagamento:</b> ${esc(c.codice_pagamento||'-')}</p>
      <p><b>Banca:</b> ${esc(c.banca_cliente||'-')} &nbsp; <b>IBAN:</b> ${esc(c.banca_iban||'-')}</p>
      <p><b>Note:</b> ${esc(c.note)}</p>
    </div>
    <div class="box"><h2>🚗 Veicoli del cliente (${vs.length})</h2>
      <table><tr><th>Targa</th><th>Marca / Modello</th><th>Telaio</th><th>Km</th><th></th></tr>
      ${vs.map(v=>`<tr><td><b>${esc(v.targa)}</b></td><td>${esc(v.marca)} ${esc(v.modello)}<br>${esc(v.versione)}</td><td>${esc(v.telaio)}</td><td>${v.km||0}</td><td><a class="btn dark" href="/veicoli/${v.id}">Apri</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/clienti/:id/modifica', async (req,res)=>{
  await dpServiceSyncMaster().catch(()=>{});
  const c=await get('SELECT * FROM clienti WHERE id=?',[req.params.id]); if(!c) return res.status(404).send('Cliente non trovato');
  const CLIENTE=c;
  res.send(page('Modifica cliente',`
  <div class="actions"><a class="btn dark" href="/clienti/${c.id}">Indietro</a></div>
  <div class="box"><h1>Modifica cliente</h1><form method="post" action="/clienti/${c.id}/modifica">
  <label>Ragione sociale / Nome</label><input name="ragione_sociale" value="${esc(c.ragione_sociale)}" required>
  <label>P.IVA</label><input name="piva" value="${esc(c.piva)}"><label>Codice fiscale</label><input name="cf" value="${esc(c.cf)}">
  <label>Indirizzo</label><input name="indirizzo" value="${esc(c.indirizzo)}"><label>Città</label><input name="citta" value="${esc(c.citta)}"><label>Provincia</label><input name="provincia" value="${esc(c.provincia)}">
  <label>Telefono</label><input name="telefono" value="${esc(c.telefono)}"><label>Email</label><input name="email" value="${esc(c.email)}"><label>PEC</label><input name="pec" value="${esc(c.pec)}"><label>SDI</label><input name="sdi" value="${esc(c.sdi)}">

    <h3>🧾 IVA e pagamento</h3>
    <label>Codice IVA</label><select name="codice_iva">
      <option value="">Scegli...</option>
      ${[['22','IVA 22%'],['22SP','IVA 22% SPLIT PAYMENT'],['22RC','IVA 22% REVERSE CHARGE'],['22EX','IVA 22% EXTRA UE'],['10','IVA 10%'],['4','IVA 4%'],['5','IVA 5%'],['NI41','Non imponibile Art. 41 Cessioni UE'],['NI8','Non imponibile Art. 8 c.1 lett. c'],['NI8b','Non imponibile Art. 8 bis'],['FC7','Art. 7 TER operazione non soggetta']].map(x=>`<option value="${x[0]}" ${String(CLIENTE.codice_iva||'')===x[0]?'selected':''}>${x[0]} - ${x[1]}</option>`).join('')}
    </select>
    <label>Condizione di pagamento</label><select name="codice_pagamento">
      <option value="">Scegli...</option>
      ${[['BO30','Bonifico 30 gg'],['BO6','Bonifico 60 gg'],['BO9','Bonifico 90 gg'],['BOIMM','Bonifico vista fattura'],['RB3','Ri.Ba. 30 gg DF FM'],['RB36','Ri.Ba. 30/60 gg DF FM'],['RB6','Ri.Ba. 60 gg DF FM'],['RB9','Ri.Ba. 90 gg DF FM'],['RBF10','Ri.Ba. fine mese +10'],['RBFIN','Ri.Ba. fine mese'],['RID30','Rimessa diretta 30 gg'],['RID60','Rimessa diretta 60 gg'],['RID90','Rimessa diretta 90 gg FM'],['RD','Rimessa diretta a vista']].map(x=>`<option value="${x[0]}" ${String(CLIENTE.codice_pagamento||'')===x[0]?'selected':''}>${x[0]} - ${x[1]}</option>`).join('')}
    </select>
    <h3>🏦 Dati bancari cliente</h3>
    <label>Banca cliente</label><input name="banca_cliente" value="${esc(CLIENTE.banca_cliente||'')}">
    <label>Codice banca</label><input name="banca_codice" value="${esc(CLIENTE.banca_codice||'')}">
    <label>Numero conto corrente</label><input name="banca_conto" value="${esc(CLIENTE.banca_conto||'')}">
    <label>ABI</label><input name="banca_abi" value="${esc(CLIENTE.banca_abi||'')}"><label>CAB</label><input name="banca_cab" value="${esc(CLIENTE.banca_cab||'')}">
    <label>Paese</label><input name="banca_paese" value="${esc(CLIENTE.banca_paese||'IT')}">
    <label>CIN EUR</label><input name="banca_cin_eur" value="${esc(CLIENTE.banca_cin_eur||'')}"><label>CIN Italia</label><input name="banca_cin_it" value="${esc(CLIENTE.banca_cin_it||'')}">
    <label>Valuta</label><input name="banca_valuta" value="${esc(CLIENTE.banca_valuta||'EUR')}"><label>BIC / SWIFT</label><input name="banca_bic" value="${esc(CLIENTE.banca_bic||'')}">
    <label>IBAN</label><input name="banca_iban" value="${esc(CLIENTE.banca_iban||'')}"><label>BBAN</label><input name="banca_bban" value="${esc(CLIENTE.banca_bban||'')}">
    <label>PEC banca</label><input name="banca_pec" value="${esc(CLIENTE.banca_pec||'')}"><label>Note banca</label><input name="banca_note" value="${esc(CLIENTE.banca_note||'')}">

  <label>Note</label><textarea name="note">${esc(c.note)}</textarea><p><button class="btn">Salva modifiche</button></p></form></div>`));
});

router.post('/clienti/:id/modifica', async (req,res)=>{
  const b=req.body;
  await run(`UPDATE clienti SET ragione_sociale=?,piva=?,cf=?,indirizzo=?,citta=?,provincia=?,telefono=?,email=?,pec=?,sdi=?,codice_iva=?,codice_pagamento=?,banca_cliente=?,banca_codice=?,banca_conto=?,banca_abi=?,banca_cab=?,banca_paese=?,banca_cin_eur=?,banca_cin_it=?,banca_valuta=?,banca_bic=?,banca_iban=?,banca_bban=?,banca_pec=?,banca_note=?,note=? WHERE id=?`,
    [b.ragione_sociale,b.piva,b.cf,b.indirizzo,b.citta,b.provincia,b.telefono,b.email,b.pec,b.sdi,b.codice_iva,b.codice_pagamento,b.banca_cliente,b.banca_codice,b.banca_conto,b.banca_abi,b.banca_cab,b.banca_paese,b.banca_cin_eur,b.banca_cin_it,b.banca_valuta,b.banca_bic,b.banca_iban,b.banca_bban,b.banca_pec,b.banca_note,b.note,req.params.id]);
  await dpServiceSaveToMaster(Number(req.params.id),b);
  res.redirect('/clienti/'+req.params.id);
});

router.get('/veicoli', async (req,res)=>{
  const q=(req.query.q||'').trim();
  const params=[]; let where='';
  if(q){ where='WHERE v.targa LIKE ? OR v.marca LIKE ? OR v.modello LIKE ? OR v.versione LIKE ? OR v.telaio LIKE ? OR c.ragione_sociale LIKE ?'; params.push(...Array(6).fill(`%${q}%`)); }
  const rows=await all(`SELECT v.*,c.ragione_sociale FROM veicoli v LEFT JOIN clienti c ON c.id=v.cliente_id ${where} ORDER BY v.targa LIMIT 500`,params);
  res.send(page('Veicoli',`
  <div class="actions"><a class="btn dark" href="/__DP_MAIN__">🏠 Menu DP Gestionale</a><a class="btn" href="/veicoli/nuovo">+ Nuovo veicolo</a></div>
  <div class="box"><h1>🚗 Veicoli clienti (${rows.length})</h1>
  <form class="filters"><input name="q" placeholder="Targa, cliente, marca, modello, telaio" value="${esc(q)}"><span></span><span></span><button class="btn">Filtra</button></form>
  <table><tr><th>Targa</th><th>Cliente</th><th>Veicolo</th><th>Telaio</th><th>Km</th><th></th></tr>
  ${rows.map(v=>`<tr><td><b>${esc(v.targa)}</b></td><td>${esc(v.ragione_sociale)}</td><td>${esc(v.marca)} ${esc(v.modello)}<br>${esc(v.versione)}</td><td>${esc(v.telaio)}</td><td>${v.km||0}</td><td><a class="btn dark" href="/veicoli/${v.id}">Apri</a></td></tr>`).join('')}
  </table></div>`));
});

router.get('/veicoli/nuovo', async (req,res)=>{
  await dpServiceSyncMaster().catch(()=>{});
  const cs=await all('SELECT id,ragione_sociale FROM clienti ORDER BY ragione_sociale LIMIT 5000');
  const selected=Number(req.query.cliente_id)||0;
  res.send(page('Nuovo veicolo',`
  <div class="actions"><a class="btn dark" href="/veicoli">Indietro</a></div>
  <div class="box"><h1>+ Nuovo veicolo cliente</h1><form method="post" action="/veicoli">
  <label>Cliente</label><select name="cliente_id"><option value="">Scegli...</option>${cs.map(c=>`<option value="${c.id}" ${c.id===selected?'selected':''}>${esc(c.ragione_sociale)}</option>`).join('')}</select>
  <label>Targa</label><input name="targa" required style="text-transform:uppercase">
  <label>Marca</label><input name="marca"><label>Modello</label><input name="modello"><label>Versione / Motore</label><input name="versione">
  <label>Telaio</label><input name="telaio"><label>Anno</label><input name="anno"><label>Km attuali</label><input type="number" name="km"><label>Alimentazione</label><input name="alimentazione">
  <label>Note</label><textarea name="note"></textarea><p><button class="btn">Salva veicolo</button></p></form></div>`));
});

router.post('/veicoli', async (req,res)=>{
  const b=req.body; const targa=String(b.targa||'').toUpperCase().replace(/\s/g,'');
  try{
    const r=await run(`INSERT INTO veicoli(cliente_id,targa,marca,modello,versione,telaio,anno,km,alimentazione,note) VALUES(?,?,?,?,?,?,?,?,?,?)`,
      [b.cliente_id||null,targa,b.marca,b.modello,b.versione,b.telaio,b.anno,Number(b.km)||0,b.alimentazione,b.note]);
    res.redirect('/veicoli/'+r.lastID);
  }catch(e){ res.status(400).send(page('Errore',`<div class="box"><h2>Errore</h2><p>${esc(e.message)}</p><a class="btn" href="/veicoli">Torna</a></div>`)); }
});

router.get('/veicoli/:id', async (req,res)=>{
  const v=await get(`SELECT v.*,c.ragione_sociale FROM veicoli v LEFT JOIN clienti c ON c.id=v.cliente_id WHERE v.id=?`,[req.params.id]);
  if(!v) return res.status(404).send('Veicolo non trovato');
  res.send(page(v.targa,`
    <div class="actions"><a class="btn dark" href="/veicoli">Indietro</a><a class="btn" href="/veicoli/${v.id}/modifica">Modifica veicolo</a>${v.cliente_id?`<a class="btn green" href="/clienti/${v.cliente_id}">Apri cliente</a>`:''}</div>
    <div class="box"><h1>🚗 ${esc(v.targa)} — ${esc(v.marca)} ${esc(v.modello)}</h1>
      <p><b>Cliente:</b> ${esc(v.ragione_sociale)}</p><p><b>Versione/Motore:</b> ${esc(v.versione)}</p><p><b>Telaio:</b> ${esc(v.telaio)}</p>
      <p><b>Anno:</b> ${esc(v.anno)} &nbsp; <b>Km:</b> ${v.km||0} &nbsp; <b>Alimentazione:</b> ${esc(v.alimentazione)}</p>
      <p><b>Note:</b> ${esc(v.note)}</p>
      <p><a class="btn" href="/ordini/nuovo?veicolo_id=${v.id}">+ Nuovo ordine di lavoro</a></p>
    </div>
    <div class="box"><h2>🧾 Storico interventi</h2>
      <table><tr><th>N.</th><th>Data</th><th>Km</th><th>Lavoro</th><th>Stato</th><th>Totale</th><th></th></tr>
      ${(await all('SELECT * FROM ordini_lavoro WHERE veicolo_id=? ORDER BY id DESC',[v.id])).map(o=>`<tr><td>${o.numero||o.id}/S</td><td>${esc(o.data_apertura)}</td><td>${o.km_ingresso||0}</td><td>${esc(o.descrizione_lavoro)}</td><td>${esc(o.stato)}</td><td>€ ${(Number(o.totale_ivato)||0).toFixed(2)}</td><td><a class="btn dark" href="/ordini/${o.id}">Apri</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/veicoli/:id/modifica', async (req,res)=>{
  const v=await get('SELECT * FROM veicoli WHERE id=?',[req.params.id]); if(!v) return res.status(404).send('Veicolo non trovato');
  const cs=await all('SELECT id,ragione_sociale FROM clienti ORDER BY ragione_sociale LIMIT 5000');
  res.send(page('Modifica veicolo',`
  <div class="actions"><a class="btn dark" href="/veicoli/${v.id}">Indietro</a></div>
  <div class="box"><h1>Modifica veicolo</h1><form method="post" action="/veicoli/${v.id}/modifica">
  <label>Cliente</label><select name="cliente_id"><option value="">Scegli...</option>${cs.map(c=>`<option value="${c.id}" ${c.id===v.cliente_id?'selected':''}>${esc(c.ragione_sociale)}</option>`).join('')}</select>
  <label>Targa</label><input name="targa" value="${esc(v.targa)}" required><label>Marca</label><input name="marca" value="${esc(v.marca)}"><label>Modello</label><input name="modello" value="${esc(v.modello)}">
  <label>Versione / Motore</label><input name="versione" value="${esc(v.versione)}"><label>Telaio</label><input name="telaio" value="${esc(v.telaio)}"><label>Anno</label><input name="anno" value="${esc(v.anno)}">
  <label>Km</label><input type="number" name="km" value="${v.km||0}"><label>Alimentazione</label><input name="alimentazione" value="${esc(v.alimentazione)}"><label>Note</label><textarea name="note">${esc(v.note)}</textarea>
  <p><button class="btn">Salva modifiche</button></p></form></div>`));
});

router.post('/veicoli/:id/modifica', async (req,res)=>{
  const b=req.body; const targa=String(b.targa||'').toUpperCase().replace(/\s/g,'');
  await run(`UPDATE veicoli SET cliente_id=?,targa=?,marca=?,modello=?,versione=?,telaio=?,anno=?,km=?,alimentazione=?,note=? WHERE id=?`,
    [b.cliente_id||null,targa,b.marca,b.modello,b.versione,b.telaio,b.anno,Number(b.km)||0,b.alimentazione,b.note,req.params.id]);
  res.redirect('/veicoli/'+req.params.id);
});


router.get('/ordini', async (req,res)=>{
  const q=(req.query.q||'').trim();
  const stato=(req.query.stato||'').trim();
  const params=[]; const cond=[];
  if(q){
    cond.push('(v.targa LIKE ? OR v.marca LIKE ? OR v.modello LIKE ? OR c.ragione_sociale LIKE ? OR o.descrizione_lavoro LIKE ?)');
    params.push(...Array(5).fill(`%${q}%`));
  }
  if(stato){ cond.push('o.stato=?'); params.push(stato); }
  const where=cond.length?'WHERE '+cond.join(' AND '):'';
  const rows=await all(`SELECT o.*,v.targa,v.marca,v.modello,c.ragione_sociale,
      COALESCE((SELECT SUM(rr.quantita*rr.prezzo_unitario*(1+COALESCE(rr.iva,22)/100.0)) FROM righe_lavoro rr WHERE rr.ordine_id=o.id),0) AS totale_ivato
    FROM ordini_lavoro o
    LEFT JOIN veicoli v ON v.id=o.veicolo_id
    LEFT JOIN clienti c ON c.id=o.cliente_id
    ${where}
    ORDER BY o.id DESC LIMIT 1000`,params);

  const aperti=rows.filter(o=>['APERTO','IN_LAVORAZIONE','ATTESA_RICAMBI'].includes(o.stato));
  const chiusi=rows.filter(o=>['PRONTO','CHIUSO'].includes(o.stato));

  const tabella=(arr)=>`
    <table><tr><th>N.</th><th>Data</th><th>Cliente</th><th>Veicolo</th><th>Km</th><th>Stato</th><th>Totale</th><th></th></tr>
    ${arr.length ? arr.map(o=>`<tr><td><b>ODL ${o.numero||o.id}</b></td><td>${esc(o.data_apertura)}</td><td>${esc(o.ragione_sociale)}</td><td><b>${esc(o.targa)}</b><br>${esc(o.marca)} ${esc(o.modello)}</td><td>${o.km_ingresso||0}</td><td><b>${esc(o.stato)}</b></td><td>€ ${(Number(o.totale)||0).toFixed(2)}</td><td><a class="btn dark" href="/ordini/${o.id}">Apri</a></td></tr>`).join('') : `<tr><td colspan="8">Nessun ordine</td></tr>`}
    </table>`;

  res.send(page('Ordini di lavoro',`
    <div class="actions"><a class="btn" href="/ordini/nuovo">+ Nuovo ordine di lavoro</a></div>
    <div class="box"><h1>🧾 Ordini di lavoro</h1>
      <form class="filters">
        <input name="q" placeholder="Targa, cliente, modello, lavorazione" value="${esc(q)}">
        <select name="stato">
          <option value="">Tutti gli stati</option>
          ${['APERTO','IN_LAVORAZIONE','ATTESA_RICAMBI','PRONTO','CHIUSO'].map(s=>`<option ${stato===s?'selected':''}>${s}</option>`).join('')}
        </select>
        <span></span><button class="btn">Filtra</button>
      </form>
    </div>
    <div class="box"><h2>🔧 IN LAVORAZIONE (${aperti.length})</h2>${tabella(aperti)}</div>
    <div class="box" id="chiusi"><h2>✅ PRONTI / CHIUSI (${chiusi.length})</h2>${tabella(chiusi)}</div>
  `));
});

router.get('/ordini/nuovo', async (req,res)=>{
  await dpServiceSyncMaster().catch(()=>{});
  const clienti=await all(`SELECT id,ragione_sociale,piva,telefono FROM clienti ORDER BY ragione_sociale LIMIT 6000`);
  const selectedCliente=Number(req.query.cliente_id)||0;
  res.send(page('Nuovo ordine di lavoro',`
    <div class="actions"><a class="btn dark" href="/ordini">Indietro</a></div>
    <div class="box"><h1>+ Accettazione officina</h1>
      <form method="post" action="/ordini">
        <label>Cliente</label>
        <select name="cliente_id" required>
          <option value="">Scegli cliente...</option>
          ${clienti.map(c=>`<option value="${c.id}" ${c.id===selectedCliente?'selected':''}>${esc(c.ragione_sociale)}${c.piva?` — P.IVA ${esc(c.piva)}`:''}</option>`).join('')}
        </select>

        <label>Targa</label>
        <input name="targa" required style="text-transform:uppercase" placeholder="Es. AB123CD">

        <label>Marca</label>
        <input name="marca" placeholder="Es. FIAT">

        <label>Modello</label>
        <input name="modello" required placeholder="Es. PANDA 1.2">

        <label>Versione / Motore</label>
        <input name="versione" placeholder="Es. 1.2 BENZINA / GPL">

        <label>Km ingresso</label>
        <input type="number" name="km_ingresso" min="0">

        <label>Data ingresso</label>
        <input type="date" name="data_apertura" value="${new Date().toISOString().slice(0,10)}">

        <label>Difetto segnalato / lavoro richiesto</label>
        <textarea name="descrizione_lavoro" rows="4" required placeholder="Es. Rumore avantreno, tagliando, spia motore..."></textarea>

        <label>Diagnosi iniziale</label>
        <textarea name="diagnosi" rows="3"></textarea>

        <label>Note</label>
        <textarea name="note" rows="3"></textarea>

        <p><button class="btn">Apri lavorazione</button></p>
      </form>
      <p class="muted"><b>La vettura si salva automaticamente sul cliente.</b> Se la targa esiste già, viene riutilizzata e aggiornata.</p>
    </div>`));
});

router.post('/ordini', async (req,res)=>{
  const b=req.body;
  const clienteId=Number(b.cliente_id)||0;
  if(!clienteId) return res.status(400).send('Cliente non valido');

  const cliente=await get('SELECT * FROM clienti WHERE id=?',[clienteId]);
  if(!cliente) return res.status(400).send('Cliente non trovato');

  const targa=String(b.targa||'').toUpperCase().replace(/\s/g,'');
  if(!targa) return res.status(400).send('Targa obbligatoria');

  let v=await get('SELECT * FROM veicoli WHERE UPPER(targa)=UPPER(?)',[targa]);

  if(v){
    await run(`UPDATE veicoli SET cliente_id=?,marca=?,modello=?,versione=?,km=? WHERE id=?`,[
      clienteId,
      b.marca || v.marca || '',
      b.modello || v.modello || '',
      b.versione || v.versione || '',
      Number(b.km_ingresso)||v.km||0,
      v.id
    ]);
    v=await get('SELECT * FROM veicoli WHERE id=?',[v.id]);
  } else {
    const vr=await run(`INSERT INTO veicoli(cliente_id,targa,marca,modello,versione,km) VALUES(?,?,?,?,?,?)`,[
      clienteId,targa,b.marca||'',b.modello||'',b.versione||'',Number(b.km_ingresso)||0
    ]);
    v=await get('SELECT * FROM veicoli WHERE id=?',[vr.lastID]);
  }

  const y=new Date().getFullYear();
  const nx=await get('SELECT COALESCE(MAX(numero),0)+1 n FROM ordini_lavoro WHERE anno=?',[y]);
  const r=await run(`INSERT INTO ordini_lavoro(numero,anno,cliente_id,veicolo_id,data_apertura,km_ingresso,descrizione_lavoro,diagnosi,note)
    VALUES(?,?,?,?,?,?,?,?,?)`,[
      nx.n,y,clienteId,v.id,b.data_apertura,Number(b.km_ingresso)||0,
      b.descrizione_lavoro,b.diagnosi,b.note
    ]);

  res.redirect('/ordini/'+r.lastID);
});

router.get('/ordini/:id', async (req,res)=>{
  const o=await get(`SELECT o.*,v.targa,v.marca,v.modello,v.versione,c.ragione_sociale,c.telefono,c.piva,c.cf,c.indirizzo,c.citta,c.provincia,c.email,c.pec,c.sdi
    FROM ordini_lavoro o
    LEFT JOIN veicoli v ON v.id=o.veicolo_id
    LEFT JOIN clienti c ON c.id=o.cliente_id
    WHERE o.id=?`,[req.params.id]);
  if(!o) return res.status(404).send('Ordine non trovato');
  const righe=await all('SELECT * FROM righe_lavoro WHERE ordine_id=? ORDER BY id',[o.id]);
  const calc=dpCalcRighe(righe);
  const preventivo=await get('SELECT * FROM preventivi_service WHERE ordine_id=?',[o.id]);
  const fattura=await get('SELECT * FROM fatture_service WHERE ordine_id=?',[o.id]);
  const ricambi=await all('SELECT * FROM ricambi WHERE attivo=1 ORDER BY descrizione LIMIT 2000');
  const wa=(o.telefono||'').replace(/\D/g,'');
  const phone = wa.startsWith('39') ? wa : (wa ? '39'+wa : '');
  const text=encodeURIComponent(`Buongiorno ${o.ragione_sociale||''}, la sua vettura ${o.marca||''} ${o.modello||''} targa ${o.targa||''} è pronta per il ritiro presso DP SERVICE. Grazie.`);
  res.send(page(`ODL ${o.numero}/${o.anno}`,`
    <div class="actions"><a class="btn dark" href="/ordini">Indietro</a><a class="btn" href="/veicoli/${o.veicolo_id}">Apri veicolo</a></div>
    <div class="box">
      <h1>🧾 Ordine di lavoro ${o.numero}/${o.anno}</h1>
      <p><b>${esc(o.ragione_sociale)}</b><br>🚗 <b>${esc(o.targa)}</b> — ${esc(o.marca)} ${esc(o.modello)} ${esc(o.versione)}</p>
      <p><b>Data ingresso:</b> ${esc(o.data_apertura)} &nbsp; <b>Km:</b> ${o.km_ingresso||0}</p>
      ${o.lavorazioni_eseguite?`<div class="box" style="border-left:6px solid #159447"><h3>🔧 Lavorazioni registrate dai meccanici</h3><div style="white-space:pre-wrap">${esc(o.lavorazioni_eseguite)}</div></div>`:''}
      <form method="post" action="/ordini/${o.id}/dettagli" style="margin:14px 0">
        <label>Lavoro richiesto / descrizione intervento</label>
        <textarea name="descrizione_lavoro" rows="4">${esc(o.descrizione_lavoro||'')}</textarea>
        <label>Diagnosi / note tecniche officina</label>
        <textarea name="diagnosi" rows="4">${esc(o.diagnosi||'')}</textarea>
        <button class="btn">💾 Salva descrizione e diagnosi</button>
      </form>
      <form method="post" action="/ordini/${o.id}/stato" style="margin:14px 0">
        <label>Stato lavorazione</label>
        <div class="filters">
          <select name="stato">${['APERTO','IN_LAVORAZIONE','ATTESA_RICAMBI','PRONTO','CHIUSO'].map(s=>`<option ${o.stato===s?'selected':''}>${s}</option>`).join('')}</select>
          <input type="number" name="km_uscita" placeholder="Km uscita" value="${o.km_uscita||''}">
          <span></span><button class="btn">Aggiorna stato</button>
        </div>
      </form>
      <div class="actions" style="margin-top:12px">
        ${phone?`<a class="btn green" target="_blank" href="https://wa.me/${phone}?text=${text}">📲 AUTO PRONTA - WhatsApp</a>`:''}
        ${preventivo?`<a class="btn dark" href="/preventivi/${preventivo.id}">📄 Apri preventivo</a><form method="post" action="/ordini/${o.id}/preventivo/aggiorna"><button class="btn dark">🔄 Aggiorna preventivo dalle righe</button></form>`:`<form method="post" action="/ordini/${o.id}/preventivo"><button class="btn dark">📄 CREA PREVENTIVO DAL LAVORO</button></form>`}
        ${fattura?`<a class="btn" href="/fatture/${fattura.id}">💶 Apri fattura ${fattura.numero}/S</a>`:`<form method="post" action="/ordini/${o.id}/fattura"><button class="btn">💶 CREA FATTURA SERIE S</button></form>`}
      </div>
    </div>
    <div class="box"><h2>Ricambi / lavorazioni</h2>
      <form method="post" action="/ordini/${o.id}/righe" id="rigaForm">
        <div style="display:grid;grid-template-columns:180px 1fr 120px 150px;gap:10px;align-items:end">
          <div>
            <label>Tipo</label>
            <select name="tipo" id="tipoRiga" onchange="if(this.value==='MANODOPERA'){document.getElementById('cercaRicambio').value='MANODOPERA';this.form.descrizione.value='MANODOPERA';this.form.prezzo_unitario.value='30.00';this.form.iva.value='22';}">
              <option value="RICAMBIO">Ricambio</option>
              <option value="MANODOPERA">Manodopera</option>
              <option value="MATERIALE">Materiale d'uso</option>
              <option value="SERVIZIO">Servizio</option>
            </select>
          </div>
          <div>
            <label>Cerca ricambio / lavorazione</label>
            <input id="cercaRicambio" list="listaRicambi" autocomplete="off" placeholder="Inizia a scrivere: filtro olio, H7, batteria..." oninput="dpScegliRicambio(this.value)">
            <datalist id="listaRicambi">
              ${ricambi.map(r=>`<option value="${esc(r.descrizione)}">${esc(r.codice)} ${esc(r.categoria)} — € ${(Number(r.prezzo_vendita)||0).toFixed(2)}</option>`).join('')}
            </datalist>
          </div>
          <div>
            <label>Q.tà / ore</label>
            <input type="number" step="0.01" min="0.01" name="quantita" value="1">
          </div>
          <div>
            <label>Prezzo €</label>
            <input type="number" step="0.01" min="0" name="prezzo_unitario" placeholder="0,00">
          </div>
        </div>
        <input type="hidden" name="ricambio_id" value="">
        <input type="hidden" name="descrizione" value="">
        <input type="hidden" name="iva" value="22">
        <div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:12px">
          <button class="btn">+ Aggiungi riga</button>
          <button type="button" class="btn green" onclick="document.getElementById('tipoRiga').value='MANODOPERA';document.getElementById('cercaRicambio').value='MANODOPERA';this.form.descrizione.value='MANODOPERA';this.form.prezzo_unitario.value='30.00';this.form.quantita.focus();">⏱ Manodopera 30 €/h</button>
          <a class="btn dark" href="/ricambi/nuovo" target="_blank">+ Nuova voce listino</a>
        </div>
      </form>
      <script>
        const DP_RICAMBI=${JSON.stringify(ricambi.map(r=>({id:r.id,descrizione:r.descrizione,prezzo:Number(r.prezzo_vendita)||0,iva:Number(r.iva)||22,categoria:r.categoria||''})))};
        function dpScegliRicambio(val){
          const f=document.getElementById('rigaForm');
          const x=DP_RICAMBI.find(r=>String(r.descrizione).toLowerCase()===String(val).toLowerCase());
          if(x){
            f.ricambio_id.value=x.id;
            f.descrizione.value=x.descrizione;
            f.prezzo_unitario.value=Number(x.prezzo).toFixed(2);
            f.iva.value=x.iva;
            if(String(x.categoria).toLowerCase()==='servizi' && String(x.descrizione).toUpperCase()==='MANODOPERA') f.tipo.value='MANODOPERA';
          } else {
            f.ricambio_id.value='';
            f.descrizione.value=val;
          }
        }
      </script>
      <table><tr><th>Tipo</th><th>Descrizione</th><th>Q.tà / Ore</th><th>Prezzo</th><th>Totale</th><th></th></tr>
      ${righe.map(r=>`<tr><td>${esc(r.tipo)}</td><td>${esc(r.descrizione)}</td><td>${r.quantita}</td><td>€ ${(Number(r.prezzo_unitario)||0).toFixed(2)}</td><td>€ ${((Number(r.quantita)||0)*(Number(r.prezzo_unitario)||0)).toFixed(2)}</td><td><form method="post" action="/ordini/${o.id}/righe/${r.id}/elimina"><button class="btn dark">Elimina</button></form></td></tr>`).join('')}
      </table>
      <div style="max-width:430px;margin-left:auto;margin-top:18px;background:#f4f4f4;border-radius:14px;padding:16px"><div><b>Imponibile:</b> ${dpEuro(calc.imponibile)}</div><div><b>IVA:</b> ${dpEuro(calc.iva)}</div><div style="font-size:26px;color:#d40000;margin-top:8px"><b>TOTALE: ${dpEuro(calc.totale)}</b></div></div>
    </div>`));
});


router.post('/ordini/:id/dettagli', async (req,res)=>{
  await run(`UPDATE ordini_lavoro SET descrizione_lavoro=?,diagnosi=?,note=COALESCE(?,note) WHERE id=?`,
    [String(req.body.descrizione_lavoro||''),String(req.body.diagnosi||''),req.body.note||null,req.params.id]);
  res.redirect('/ordini/'+req.params.id);
});

router.post('/ordini/:id/righe', async (req,res)=>{
  const tipo=req.body.tipo||'RICAMBIO';
  let prezzo=Number(req.body.prezzo_unitario)||0;
  let descrizione=req.body.descrizione||'';
  let iva=Number(req.body.iva)||22;
  if(req.body.ricambio_id){
    const r=await get('SELECT * FROM ricambi WHERE id=? AND attivo=1',[req.body.ricambio_id]);
    if(r){
      if(!descrizione) descrizione=r.descrizione;
      if(prezzo<=0) prezzo=Number(r.prezzo_vendita)||0;
      iva=Number(r.iva)||22;
    }
  }
  if(tipo==='MANODOPERA' && prezzo<=0) prezzo=30;
  if(tipo==='MANODOPERA' && !descrizione) descrizione='MANODOPERA';
  await run(`INSERT INTO righe_lavoro(ordine_id,tipo,descrizione,quantita,prezzo_unitario,iva) VALUES(?,?,?,?,?,?)`,
    [req.params.id,tipo,descrizione,Number(req.body.quantita)||1,prezzo,iva]);
  const t=await get('SELECT COALESCE(SUM(quantita*prezzo_unitario),0) t FROM righe_lavoro WHERE ordine_id=?',[req.params.id]);
  await run('UPDATE ordini_lavoro SET totale=? WHERE id=?',[t.t,req.params.id]);
  res.redirect('/ordini/'+req.params.id);
});

router.post('/ordini/:id/righe/:rid/elimina', async (req,res)=>{
  await run('DELETE FROM righe_lavoro WHERE id=? AND ordine_id=?',[req.params.rid,req.params.id]);
  const t=await get('SELECT COALESCE(SUM(quantita*prezzo_unitario),0) t FROM righe_lavoro WHERE ordine_id=?',[req.params.id]);
  await run('UPDATE ordini_lavoro SET totale=? WHERE id=?',[t.t,req.params.id]);
  res.redirect('/ordini/'+req.params.id);
});

router.post('/ordini/:id/stato', async (req,res)=>{
  const km=Number(req.body.km_uscita)||0;
  const stato=req.body.stato||'APERTO';
  await run('UPDATE ordini_lavoro SET stato=?,km_uscita=?,data_chiusura=? WHERE id=?',[
    stato,km,stato==='CHIUSO'?new Date().toISOString().slice(0,10):null,req.params.id
  ]);
  const o=await get('SELECT veicolo_id FROM ordini_lavoro WHERE id=?',[req.params.id]);
  if(o && km>0) await run('UPDATE veicoli SET km=? WHERE id=?',[km,o.veicolo_id]);
  if(stato==='PRONTO' || stato==='CHIUSO') return res.redirect('/ordini#chiusi');
  res.redirect('/ordini/'+req.params.id);
});



router.post('/ordini/:id/preventivo', async (req,res)=>{
  const d=await dpServiceDocData(req.params.id);
  if(!d) return res.status(404).send('Ordine non trovato');
  if(!d.righe.length) return res.status(400).send(page('Preventivo',`<div class="box"><h1>⚠️ Preventivo vuoto</h1><p>Prima aggiungi almeno una lavorazione, ricambio o manodopera nell'ordine di lavoro.</p><a class="btn" href="/ordini/${req.params.id}">Torna all'ordine</a></div>`));
  let p=await get('SELECT * FROM preventivi_service WHERE ordine_id=?',[req.params.id]);
  if(!p){
    const anno=new Date().getFullYear();
    const nx=await get('SELECT COALESCE(MAX(numero),0)+1 n FROM preventivi_service WHERE anno=?',[anno]);
    const data=new Date().toISOString().slice(0,10);
    const r=await run(`INSERT INTO preventivi_service(numero,anno,ordine_id,data,imponibile,iva,totale,righe_json,note) VALUES(?,?,?,?,?,?,?,?,?)`,[
      nx.n,anno,req.params.id,data,d.calc.imponibile,d.calc.iva,d.calc.totale,JSON.stringify(d.righe),d.o.note||''
    ]);
    await run('UPDATE preventivi_service SET token_cliente=? WHERE id=?',[dpServiceToken(),r.lastID]);
    p=await get('SELECT * FROM preventivi_service WHERE id=?',[r.lastID]);
  }
  res.redirect('/preventivi/'+p.id);
});

router.post('/ordini/:id/fattura', async (req,res)=>{
  const d=await dpServiceDocData(req.params.id);
  if(!d) return res.status(404).send('Ordine non trovato');
  if(!d.righe.length) return res.status(400).send(page('Fattura',`<div class="box"><h1>⚠️ Fattura vuota</h1><p>Prima aggiungi le lavorazioni/ricambi nell'ordine.</p><a class="btn" href="/ordini/${req.params.id}">Torna all'ordine</a></div>`));
  let f=await get('SELECT * FROM fatture_service WHERE ordine_id=?',[req.params.id]);
  if(!f){
    const anno=new Date().getFullYear();
    const nx=await get("SELECT COALESCE(MAX(numero),0)+1 n FROM fatture_service WHERE anno=? AND serie='S'",[anno]);
    const data=new Date().toISOString().slice(0,10);
    const r=await run(`INSERT INTO fatture_service(numero,serie,anno,ordine_id,data,imponibile,iva,totale,righe_json,pagamento,note) VALUES(?,?,?,?,?,?,?,?,?,?,?)`,[
      nx.n,'S',anno,req.params.id,data,d.calc.imponibile,d.calc.iva,d.calc.totale,JSON.stringify(d.righe),'Bonifico vista fattura',d.o.note||''
    ]);
    await run('UPDATE ordini_lavoro SET fatturato=1 WHERE id=?',[req.params.id]);
    f=await get('SELECT * FROM fatture_service WHERE id=?',[r.lastID]);
  }
  res.redirect('/fatture/'+f.id);
});

router.get('/preventivi', async (req,res)=>{
  const rows=await all(`SELECT p.*,o.numero odl,v.targa,c.ragione_sociale FROM preventivi_service p
    LEFT JOIN ordini_lavoro o ON o.id=p.ordine_id LEFT JOIN veicoli v ON v.id=o.veicolo_id LEFT JOIN clienti c ON c.id=o.cliente_id
    ORDER BY p.id DESC LIMIT 500`);
  res.send(page('Preventivi',`<div class="box"><h1>📄 Preventivi DP SERVICE</h1><p><a class="btn" href="/ordini">+ Crea preventivo da un ordine di lavoro</a></p><table><tr><th>N.</th><th>Data</th><th>Cliente</th><th>Targa</th><th>Totale</th><th></th></tr>${rows.map(x=>`<tr><td><b>${x.numero}/P</b></td><td>${dpItDate(x.data)}</td><td>${esc(x.ragione_sociale)}</td><td>${esc(x.targa)}</td><td>${dpEuro(x.totale)}</td><td><a class="btn dark" href="/preventivi/${x.id}">Apri</a></td></tr>`).join('')}</table></div>`));
});

router.get('/preventivi/:id([0-9]+)', async (req,res)=>{
  const p=await get('SELECT * FROM preventivi_service WHERE id=?',[req.params.id]); if(!p) return res.status(404).send('Preventivo non trovato');
  const d=await dpServiceDocData(p.ordine_id); if(!d) return res.status(404).send('Ordine non trovato');
  if(!p.token_cliente){ await run('UPDATE preventivi_service SET token_cliente=? WHERE id=?',[dpServiceToken(),p.id]); p.token_cliente=(await get('SELECT token_cliente FROM preventivi_service WHERE id=?',[p.id])).token_cliente; }
  const ph=dpPhone(d.o.telefono);
  const pdfUrl=dpServiceUrl(req,`/preventivi/${p.id}.pdf`);
  const approveUrl=dpServiceUrl(req,`/preventivi/${p.id}/cliente/${p.token_cliente}`);
  const msg=encodeURIComponent(`Buongiorno ${d.o.ragione_sociale||''}, ecco il preventivo DP SERVICE n. ${p.numero}/P per ${d.o.targa||''}. Totale ${dpEuro(p.totale)}.\n\nPDF: ${pdfUrl}\n\n✅ Approva o rifiuta qui: ${approveUrl}`);
  res.send(page(`Preventivo ${p.numero}/P`,`<div class="box"><h1>📄 PREVENTIVO ${p.numero}/P</h1>
    <p><b>${esc(d.o.ragione_sociale)}</b> - ${esc(d.o.targa)} - ${esc(d.o.marca)} ${esc(d.o.modello)}</p>
    <p>Stato: <b>${esc(p.stato||'EMESSO')}</b></p><h2>Totale ${dpEuro(p.totale)}</h2>
    <div class="actions">
      <a class="btn" target="_blank" href="/preventivi/${p.id}.pdf">📄 PDF</a>
      ${ph?`<a class="btn green" target="_blank" href="https://wa.me/${ph}?text=${msg}">📲 INVIA WHATSAPP</a>`:''}
      ${d.o.email?`<form method="post" action="/preventivi/${p.id}/email"><button class="btn">✉️ INVIA EMAIL + PDF</button></form>`:''}
      <a class="btn dark" target="_blank" href="/preventivi/${p.id}/cliente/${p.token_cliente}">👤 Anteprima cliente / approvazione</a>
      <a class="btn dark" href="/ordini/${p.ordine_id}">Torna all'ordine</a>
    </div></div>`));
});

router.post('/preventivi/:id/email', async (req,res)=>{
  try{
    const p=await get('SELECT * FROM preventivi_service WHERE id=?',[req.params.id]); if(!p) throw new Error('Preventivo non trovato');
    const d=await dpServiceDocData(p.ordine_id); if(!d || !d.o.email) throw new Error('Email cliente mancante');
    if(!p.token_cliente){ await run('UPDATE preventivi_service SET token_cliente=? WHERE id=?',[dpServiceToken(),p.id]); p.token_cliente=(await get('SELECT token_cliente FROM preventivi_service WHERE id=?',[p.id])).token_cliente; }
    try{ const snap=JSON.parse(p.righe_json||'[]'); if(snap.length){d.righe=snap;d.calc=dpCalcRighe(snap);} }catch(_){}
    const buf=await dpServicePdfBuffer('PREVENTIVO',`N. ${p.numero}/P`,dpItDate(p.data),d);
    const approveUrl=dpServiceUrl(req,`/preventivi/${p.id}/cliente/${p.token_cliente}`);
    await dpServiceSendEmail(d.o.email,`Preventivo DP SERVICE ${p.numero}/P`,
      `Buongiorno ${d.o.ragione_sociale||''},\nin allegato trova il preventivo DP SERVICE per ${d.o.targa||''}.\nTotale ${dpEuro(p.totale)}.\n\nPuò approvare o rifiutare qui: ${approveUrl}`,
      [{filename:`Preventivo_DP_SERVICE_${p.numero}_P.pdf`,content:buf,contentType:'application/pdf'}]);
    await run(`UPDATE preventivi_service SET inviato_email_at=CURRENT_TIMESTAMP WHERE id=?`,[p.id]);
    res.redirect('/preventivi/'+p.id);
  }catch(e){ res.status(500).send(page('Invio email',`<div class="box"><h1>Invio non riuscito</h1><p>${esc(e.message)}</p><a class="btn" href="/preventivi/${req.params.id}">Indietro</a></div>`)); }
});

router.get('/preventivi/:id/cliente/:token', async (req,res)=>{
  const p=await get('SELECT * FROM preventivi_service WHERE id=? AND token_cliente=?',[req.params.id,req.params.token]); if(!p) return res.status(404).send('Link non valido');
  const d=await dpServiceDocData(p.ordine_id); if(!d) return res.status(404).send('Ordine non trovato');
  let righe=[]; try{righe=JSON.parse(p.righe_json||'[]')}catch(_){}
  res.send(page('Conferma preventivo',`<div class="public-wrap"><div class="hero"><h1>DP SERVICE</h1><p>Preventivo ${p.numero}/P — ${esc(d.o.targa)}</p></div><div class="box">
    <h2>${esc(d.o.ragione_sociale)}</h2>
    <p><b>Lavoro richiesto:</b> ${esc(d.o.descrizione_lavoro||'')}</p>
    <table><tr><th>Descrizione</th><th>Q.tà</th><th>Totale</th></tr>${righe.map(r=>`<tr><td>${esc(r.descrizione)}</td><td>${r.quantita}</td><td>${dpEuro((Number(r.quantita)||0)*(Number(r.prezzo_unitario)||0))}</td></tr>`).join('')}</table>
    <h1>Totale ${dpEuro(p.totale)}</h1><p>Stato: <b>${esc(p.stato||'EMESSO')}</b></p>
    ${(p.stato==='APPROVATO'||p.stato==='RIFIUTATO')?'':`<div class="actions"><form method="post" action="/preventivi/${p.id}/cliente/${p.token_cliente}/approva"><button class="btn green">✅ APPROVO IL PREVENTIVO</button></form><form method="post" action="/preventivi/${p.id}/cliente/${p.token_cliente}/rifiuta"><button class="btn dark">❌ NON APPROVO</button></form></div>`}
  </div></div>`));
});
router.post('/preventivi/:id/cliente/:token/approva', async (req,res)=>{
  const p=await get('SELECT * FROM preventivi_service WHERE id=? AND token_cliente=?',[req.params.id,req.params.token]); if(!p) return res.status(404).send('Link non valido');
  await run(`UPDATE preventivi_service SET stato='APPROVATO',approvato_at=CURRENT_TIMESTAMP,rifiutato_at=NULL WHERE id=?`,[p.id]);
  await run(`INSERT INTO service_alerts(tipo,titolo,messaggio,ordine_id) VALUES('PREVENTIVO_APPROVATO',?,?,?)`,[`Preventivo ${p.numero}/P APPROVATO`,`Il cliente ha approvato online il preventivo ${p.numero}/P.`,p.ordine_id]);
  await dpServiceNotifyStaff(`✅ DP SERVICE - Preventivo ${p.numero}/P APPROVATO dal cliente.`);
  res.redirect(`/preventivi/${p.id}/cliente/${p.token_cliente}`);
});
router.post('/preventivi/:id/cliente/:token/rifiuta', async (req,res)=>{
  const p=await get('SELECT * FROM preventivi_service WHERE id=? AND token_cliente=?',[req.params.id,req.params.token]); if(!p) return res.status(404).send('Link non valido');
  await run(`UPDATE preventivi_service SET stato='RIFIUTATO',rifiutato_at=CURRENT_TIMESTAMP,approvato_at=NULL WHERE id=?`,[p.id]);
  await run(`INSERT INTO service_alerts(tipo,titolo,messaggio,ordine_id) VALUES('PREVENTIVO_RIFIUTATO',?,?,?)`,[`Preventivo ${p.numero}/P RIFIUTATO`,`Il cliente non ha approvato il preventivo ${p.numero}/P.`,p.ordine_id]);
  await dpServiceNotifyStaff(`❌ DP SERVICE - Preventivo ${p.numero}/P NON APPROVATO dal cliente.`);
  res.redirect(`/preventivi/${p.id}/cliente/${p.token_cliente}`);
});

router.get('/preventivi/:id.pdf', async (req,res)=>{
  const p=await get('SELECT * FROM preventivi_service WHERE id=?',[req.params.id]); if(!p) return res.status(404).send('Preventivo non trovato');
  const d=await dpServiceDocData(p.ordine_id); if(!d) return res.status(404).send('Ordine non trovato');
  try{ const snap=JSON.parse(p.righe_json||'[]'); if(snap.length){d.righe=snap;d.calc=dpCalcRighe(snap);} }catch(e){}
  res.setHeader('Content-Type','application/pdf'); res.setHeader('Content-Disposition',`inline; filename="Preventivo_DP_SERVICE_${p.numero}_P.pdf"`);
  const doc=new PDFDocument({size:'A4',margin:0,bufferPages:true}); doc.pipe(res); dpDrawServicePdf(doc,'PREVENTIVO',`N. ${p.numero}/P`,dpItDate(p.data),d); doc.end();
});

router.get('/fatture', async (req,res)=>{
  const rows=await all(`SELECT f.*,o.numero odl,v.targa,c.ragione_sociale FROM fatture_service f
    LEFT JOIN ordini_lavoro o ON o.id=f.ordine_id LEFT JOIN veicoli v ON v.id=o.veicolo_id LEFT JOIN clienti c ON c.id=o.cliente_id
    ORDER BY f.id DESC LIMIT 500`);
  res.send(page('Fatture serie S',`<div class="box"><h1>💶 Fatture DP SERVICE - Serie S</h1><table><tr><th>N.</th><th>Data</th><th>Cliente</th><th>Targa</th><th>Totale</th><th></th></tr>${rows.map(x=>`<tr><td><b>${x.numero}/S</b></td><td>${dpItDate(x.data)}</td><td>${esc(x.ragione_sociale)}</td><td>${esc(x.targa)}</td><td>${dpEuro(x.totale)}</td><td><a class="btn dark" href="/fatture/${x.id}">Apri</a></td></tr>`).join('')}</table></div>`));
});

router.get('/fatture/:id([0-9]+)', async (req,res)=>{
  const f=await get('SELECT * FROM fatture_service WHERE id=?',[req.params.id]); if(!f) return res.status(404).send('Fattura non trovata');
  const d=await dpServiceDocData(f.ordine_id); if(!d) return res.status(404).send('Ordine non trovato');
  const ph=dpPhone(d.o.telefono); const pdfUrl=dpServiceUrl(req,`/fatture/${f.id}.pdf`);
  const msg=encodeURIComponent(`Buongiorno ${d.o.ragione_sociale||''}, le inviamo la fattura DP SERVICE n. ${f.numero}/S per la vettura ${d.o.targa||''}. Totale ${dpEuro(f.totale)}. PDF: ${pdfUrl}`);
  res.send(page(`Fattura ${f.numero}/S`,`<div class="box"><h1>💶 FATTURA ${f.numero}/S</h1><p><b>${esc(d.o.ragione_sociale)}</b> - ${esc(d.o.targa)} - ${esc(d.o.marca)} ${esc(d.o.modello)}</p><p><b>Pagamento:</b> ${esc(f.pagamento)}</p><h2>Totale ${dpEuro(f.totale)}</h2><div class="actions"><a class="btn" target="_blank" href="/fatture/${f.id}.pdf">📄 PDF</a>${ph?`<a class="btn green" target="_blank" href="https://wa.me/${ph}?text=${msg}">📲 INVIA WHATSAPP</a>`:''}${d.o.email?`<form method="post" action="/fatture/${f.id}/email"><button class="btn">✉️ INVIA EMAIL + PDF</button></form>`:''}<a class="btn dark" href="/ordini/${f.ordine_id}">Torna all'ordine</a></div></div>`));
});


router.post('/fatture/:id/email', async (req,res)=>{
  try{
    const f=await get('SELECT * FROM fatture_service WHERE id=?',[req.params.id]); if(!f) throw new Error('Fattura non trovata');
    const d=await dpServiceDocData(f.ordine_id); if(!d || !d.o.email) throw new Error('Email cliente mancante');
    try{ const snap=JSON.parse(f.righe_json||'[]'); if(snap.length){d.righe=snap;d.calc=dpCalcRighe(snap);} }catch(_){}
    const buf=await dpServicePdfBuffer('FATTURA',`N. ${f.numero}/S`,dpItDate(f.data),d,{pagamento:f.pagamento});
    await dpServiceSendEmail(d.o.email,`Fattura DP SERVICE ${f.numero}/S`,
      `Buongiorno ${d.o.ragione_sociale||''},\nin allegato trova la fattura DP SERVICE ${f.numero}/S per ${d.o.targa||''}.\nTotale ${dpEuro(f.totale)}.`,
      [{filename:`Fattura_DP_SERVICE_${f.numero}_S.pdf`,content:buf,contentType:'application/pdf'}]);
    res.redirect('/fatture/'+f.id);
  }catch(e){ res.status(500).send(page('Invio email',`<div class="box"><h1>Invio non riuscito</h1><p>${esc(e.message)}</p><a class="btn" href="/fatture/${req.params.id}">Indietro</a></div>`)); }
});

router.get('/fatture/:id.pdf', async (req,res)=>{
  const f=await get('SELECT * FROM fatture_service WHERE id=?',[req.params.id]); if(!f) return res.status(404).send('Fattura non trovata');
  const d=await dpServiceDocData(f.ordine_id); if(!d) return res.status(404).send('Ordine non trovato');
  try{ const snap=JSON.parse(f.righe_json||'[]'); if(snap.length){d.righe=snap;d.calc=dpCalcRighe(snap);} }catch(e){}
  res.setHeader('Content-Type','application/pdf'); res.setHeader('Content-Disposition',`inline; filename="Fattura_DP_SERVICE_${f.numero}_S.pdf"`);
  const doc=new PDFDocument({size:'A4',margin:0,bufferPages:true}); doc.pipe(res); dpDrawServicePdf(doc,'FATTURA',`N. ${f.numero}/S`,dpItDate(f.data),d,{pagamento:f.pagamento}); doc.end();
});

router.get('/ricambi', async (req,res)=>{
  const q=(req.query.q||'').trim();
  const params=[]; let where='WHERE attivo=1';
  if(q){
    where += ' AND (codice LIKE ? OR descrizione LIKE ? OR categoria LIKE ? OR marca LIKE ?)';
    params.push(...Array(4).fill(`%${q}%`));
  }
  const rows=await all(`SELECT * FROM ricambi ${where} ORDER BY descrizione LIMIT 1000`,params);
  res.send(page('Ricambi / Listino',`
    <div class="actions"><a class="btn dark" href="/__DP_MAIN__">🏠 Menu DP Gestionale</a><a class="btn" href="/ricambi/nuovo">+ Nuova voce</a></div>
    <div class="box"><h1>📦 Listino ricambi / lavorazioni</h1>
      <form class="filters"><input name="q" placeholder="Codice, descrizione, categoria, marca" value="${esc(q)}"><span></span><span></span><button class="btn">Cerca</button></form>
      <table><tr><th>Codice</th><th>Descrizione</th><th>Categoria</th><th>Marca</th><th>Acquisto</th><th>Vendita</th><th>IVA</th><th></th></tr>
      ${rows.map(x=>`<tr><td>${esc(x.codice)}</td><td><b>${esc(x.descrizione)}</b></td><td>${esc(x.categoria)}</td><td>${esc(x.marca)}</td><td>€ ${(Number(x.prezzo_acquisto)||0).toFixed(2)}</td><td>€ ${(Number(x.prezzo_vendita)||0).toFixed(2)}</td><td>${Number(x.iva)||22}%</td><td><a class="btn dark" href="/ricambi/${x.id}/modifica">Modifica</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/ricambi/nuovo',(req,res)=>res.send(page('Nuova voce listino',`
  <div class="actions"><a class="btn dark" href="/ricambi">Indietro</a></div>
  <div class="box"><h1>+ Nuova voce listino</h1>
    <form method="post" action="/ricambi">
      <label>Codice</label><input name="codice">
      <label>Descrizione</label><input name="descrizione" required>
      <label>Categoria</label><input name="categoria" placeholder="Es. Freni, Filtri, Oli, Elettrico...">
      <label>Marca</label><input name="marca">
      <label>Prezzo acquisto</label><input type="number" step="0.01" name="prezzo_acquisto">
      <label>Prezzo vendita</label><input type="number" step="0.01" name="prezzo_vendita">
      <label>IVA %</label><input type="number" step="0.01" name="iva" value="22">
      <label>Giacenza</label><input type="number" step="0.01" name="giacenza" value="0">
      <label>Note</label><textarea name="note"></textarea>
      <p><button class="btn">Salva voce</button></p>
    </form>
  </div>`)));

router.post('/ricambi', async (req,res)=>{
  const b=req.body;
  await run(`INSERT INTO ricambi(codice,descrizione,categoria,marca,prezzo_acquisto,prezzo_vendita,iva,giacenza,note)
    VALUES(?,?,?,?,?,?,?,?,?)`,[
      b.codice,b.descrizione,b.categoria,b.marca,Number(b.prezzo_acquisto)||0,Number(b.prezzo_vendita)||0,
      Number(b.iva)||22,Number(b.giacenza)||0,b.note
    ]);
  res.redirect('/ricambi');
});

router.get('/ricambi/:id/modifica', async (req,res)=>{
  const x=await get('SELECT * FROM ricambi WHERE id=?',[req.params.id]);
  if(!x) return res.status(404).send('Voce non trovata');
  res.send(page('Modifica voce listino',`
  <div class="actions"><a class="btn dark" href="/ricambi">Indietro</a></div>
  <div class="box"><h1>Modifica voce listino</h1>
    <form method="post" action="/ricambi/${x.id}/modifica">
      <label>Codice</label><input name="codice" value="${esc(x.codice)}">
      <label>Descrizione</label><input name="descrizione" value="${esc(x.descrizione)}" required>
      <label>Categoria</label><input name="categoria" value="${esc(x.categoria)}">
      <label>Marca</label><input name="marca" value="${esc(x.marca)}">
      <label>Prezzo acquisto</label><input type="number" step="0.01" name="prezzo_acquisto" value="${Number(x.prezzo_acquisto)||0}">
      <label>Prezzo vendita</label><input type="number" step="0.01" name="prezzo_vendita" value="${Number(x.prezzo_vendita)||0}">
      <label>IVA %</label><input type="number" step="0.01" name="iva" value="${Number(x.iva)||22}">
      <label>Giacenza</label><input type="number" step="0.01" name="giacenza" value="${Number(x.giacenza)||0}">
      <label>Note</label><textarea name="note">${esc(x.note)}</textarea>
      <p><button class="btn">Salva modifiche</button></p>
    </form>
    <form method="post" action="/ricambi/${x.id}/elimina" onsubmit="return confirm('Disattivare questa voce dal listino?')">
      <button class="btn dark">Disattiva voce</button>
    </form>
  </div>`));
});

router.post('/ricambi/:id/modifica', async (req,res)=>{
  const b=req.body;
  await run(`UPDATE ricambi SET codice=?,descrizione=?,categoria=?,marca=?,prezzo_acquisto=?,prezzo_vendita=?,iva=?,giacenza=?,note=? WHERE id=?`,[
    b.codice,b.descrizione,b.categoria,b.marca,Number(b.prezzo_acquisto)||0,Number(b.prezzo_vendita)||0,
    Number(b.iva)||22,Number(b.giacenza)||0,b.note,req.params.id
  ]);
  res.redirect('/ricambi');
});

router.post('/ricambi/:id/elimina', async (req,res)=>{
  await run('UPDATE ricambi SET attivo=0 WHERE id=?',[req.params.id]);
  res.redirect('/ricambi');
});


router.get('/richiesta',(req,res)=>{
  res.setHeader('Cache-Control','no-store');
  res.setHeader('X-Robots-Tag','noindex, nofollow, noarchive');
  res.send(pagePublic('Richiesta officina',`
    <div class="hero"><h1 style="margin:0">🔧 Richiesta DP SERVICE</h1><p style="margin-bottom:0">Compila solo i dati necessari per la tua richiesta. Questa pagina non permette di vedere clienti, ordini o altri dati del gestionale.</p></div>
    <div class="box">
      <form method="post" action="/service/richiesta" autocomplete="on">
        <label>Nome e cognome / Azienda *</label><input name="nome" autocomplete="name" required>
        <label>Telefono WhatsApp *</label><input name="telefono" inputmode="tel" autocomplete="tel" required placeholder="Es. 3331234567">
        <label>Targa *</label><input name="targa" required autocapitalize="characters" style="text-transform:uppercase">
        <div class="grid">
          <div><label>Marca</label><input name="marca" placeholder="Es. Fiat"></div>
          <div><label>Modello</label><input name="modello" placeholder="Es. Panda"></div>
        </div>
        <label>Km attuali</label><input type="number" name="km" min="0" inputmode="numeric">
        <label>Problema / intervento richiesto *</label>
        <textarea name="richiesta" rows="6" required placeholder="Descrivi il problema, le spie accese o il lavoro che vuoi richiedere"></textarea>
        <div class="privacy"><label style="margin:0;font-weight:600"><input type="checkbox" name="privacy_ok" value="1" required>
          Ho letto l'<a href="/privacy" target="_blank" rel="noopener noreferrer">informativa privacy</a> e autorizzo l'uso dei dati esclusivamente per gestire questa richiesta di officina.</label></div>
        <button class="btn" type="submit">INVIA RICHIESTA ALL'OFFICINA</button>
        <p class="small">Nessuna password del gestionale è richiesta al cliente.</p>
      </form>
    </div>`));
});

router.post('/richiesta', async (req,res)=>{
  res.setHeader('Cache-Control','no-store');
  const b=req.body||{};
  const nome=String(b.nome||'').trim();
  const telefono=String(b.telefono||'').trim();
  const targa=String(b.targa||'').toUpperCase().replace(/\s/g,'');
  const richiesta=String(b.richiesta||'').trim();
  if(b.privacy_ok!=='1') return res.status(400).send(pagePublic('Privacy',`<div class="box"><h1>Consenso necessario</h1><p>Per inviare la richiesta devi leggere e accettare l'informativa privacy.</p></div>`));
  if(!nome || !telefono || !targa || !richiesta) return res.status(400).send(pagePublic('Dati mancanti',`<div class="box"><h1>Dati mancanti</h1><p>Compila i campi obbligatori e riprova.</p></div>`));

  let c=await get(`SELECT * FROM clienti WHERE telefono=? OR telefono LIKE ? ORDER BY id LIMIT 1`,[telefono,`%${telefono.replace(/\D/g,'').slice(-8)}%`]);
  if(!c){
    const cr=await run(`INSERT INTO clienti(ragione_sociale,telefono) VALUES(?,?)`,[nome,telefono]);
    c=await get('SELECT * FROM clienti WHERE id=?',[cr.lastID]);
  }
  let v=await get(`SELECT * FROM veicoli WHERE UPPER(targa)=UPPER(?)`,[targa]);
  if(v){
    await run(`UPDATE veicoli SET cliente_id=?,marca=COALESCE(NULLIF(?,''),marca),modello=COALESCE(NULLIF(?,''),modello),km=? WHERE id=?`,
      [c.id,b.marca||'',b.modello||'',Number(b.km)||v.km||0,v.id]);
  }else{
    const vr=await run(`INSERT INTO veicoli(cliente_id,targa,marca,modello,km) VALUES(?,?,?,?,?)`,
      [c.id,targa,b.marca||'',b.modello||'',Number(b.km)||0]);
    v=await get('SELECT * FROM veicoli WHERE id=?',[vr.lastID]);
  }
  if(!v) v=await get(`SELECT * FROM veicoli WHERE UPPER(targa)=UPPER(?)`,[targa]);

  const y=new Date().getFullYear();
  const nx=await get('SELECT COALESCE(MAX(numero),0)+1 n FROM ordini_lavoro WHERE anno=?',[y]);
  const od=await run(`INSERT INTO ordini_lavoro(numero,anno,cliente_id,veicolo_id,data_apertura,km_ingresso,descrizione_lavoro,stato,note)
    VALUES(?,?,?,?,date('now','localtime'),?,?,?,?)`,
    [nx.n,y,c.id,v.id,Number(b.km)||0,richiesta,'APERTO','Richiesta cliente da pagina pubblica DP SERVICE']);
  await run(`INSERT INTO richieste_service(ordine_id,nome,telefono,targa,marca,modello,km,richiesta) VALUES(?,?,?,?,?,?,?,?)`,
    [od.lastID,nome,telefono,targa,b.marca||'',b.modello||'',Number(b.km)||0,richiesta]);

  const titolo=`Nuova richiesta officina - ${targa}`;
  const msg=`${nome} • ${telefono}\n${targa} ${b.marca||''} ${b.modello||''}\n${richiesta}`;
  await run(`INSERT INTO service_alerts(tipo,titolo,messaggio,ordine_id) VALUES('RICHIESTA_CLIENTE',?,?,?)`,[titolo,msg,od.lastID]);

  const link=dpServiceUrl(req,'/ordini/' + od.lastID);
  const waText=`🔧 NUOVA RICHIESTA DP SERVICE\n\nCliente: ${nome}\nTel: ${telefono}\nTarga: ${targa}\nVeicolo: ${b.marca||''} ${b.modello||''}\nKm: ${Number(b.km)||0}\n\nRichiesta:\n${richiesta}\n\nApri ordine: ${link}`;
  await dpServiceNotifyStaff(waText);

  res.send(pagePublic('Richiesta inviata',`
    <div class="box" style="text-align:center"><h1>✅ Richiesta inviata</h1>
      <p style="font-size:19px">La richiesta è arrivata a <b>DP SERVICE</b>.</p>
      <p>Targa <b>${esc(targa)}</b> — riferimento <b>${nx.n}/${y}</b></p>
      <p>Ti contatteremo al numero indicato per conferma e disponibilità.</p>
    </div>`));
});

router.get('/alert', async (req,res)=>{
  const rows=await all(`SELECT * FROM service_alerts ORDER BY letto ASC,id DESC LIMIT 300`);
  res.send(page('Avvisi officina',`
    <div class="box"><h1>🔔 Avvisi officina</h1>
      ${rows.length?rows.map(a=>`<div class="alert-service">
        <div><span class="badge-alert">${a.letto?'LETTO':'NUOVO'}</span> ${esc(a.titolo||'Avviso')}</div>
        <div style="white-space:pre-wrap;margin:8px 0">${esc(a.messaggio||'')}</div>
        <div class="actions">${a.ordine_id?`<a class="btn" href="/ordini/${a.ordine_id}">Apri ordine di lavoro</a>`:''}
          ${a.letto?'':`<form method="post" action="/alert/${a.id}/letto"><button class="btn dark">Segna come letto</button></form>`}
        </div>
      </div>`).join(''):'<p>Nessun avviso.</p>'}
    </div>`));
});

router.post('/alert/:id/letto', async (req,res)=>{
  await run('UPDATE service_alerts SET letto=1 WHERE id=?',[req.params.id]);
  res.redirect('/alert');
});

router.get('/ricerca', async (req,res)=>{
  const q=(req.query.q||'').trim(); let rows=[];
  if(q) rows=await all(`SELECT v.*,c.ragione_sociale FROM veicoli v LEFT JOIN clienti c ON c.id=v.cliente_id
    WHERE v.targa LIKE ? OR v.marca LIKE ? OR v.modello LIKE ? OR v.versione LIKE ? OR v.telaio LIKE ? OR c.ragione_sociale LIKE ?
    ORDER BY v.targa LIMIT 300`,Array(6).fill(`%${q}%`));
  res.send(page('Ricerca',`
  <div class="actions"><a class="btn dark" href="/__DP_MAIN__">🏠 Menu DP Gestionale</a></div>
  <div class="box"><h1>🔎 Ricerca globale</h1>
  <form class="filters"><input name="q" placeholder="Targa, cliente, marca, modello, telaio" value="${esc(q)}"><span></span><span></span><button class="btn">Cerca</button></form>
  ${q?`<table><tr><th>Targa</th><th>Cliente</th><th>Veicolo</th><th></th></tr>${rows.map(v=>`<tr><td><b>${esc(v.targa)}</b></td><td>${esc(v.ragione_sociale)}</td><td>${esc(v.marca)} ${esc(v.modello)} ${esc(v.versione)}</td><td><a class="btn dark" href="/veicoli/${v.id}">Apri</a></td></tr>`).join('')}</table>`:''}
  </div>`));
});

const dpServiceReady = initDb().then(()=>dpServiceSeedListino());
setTimeout(()=>dpServiceSyncMaster().catch(e=>console.log('DP SERVICE master sync startup:',e.message)),5000).unref?.();
router.use(async (req,res,next)=>{ try{ await dpServiceReady; next(); }catch(e){ next(e); } });
module.exports = router;

// DP SERVICE V307 - CLIENTI MASTER CONDIVISI CON RENT/TRASPORTI

// DP SERVICE V308 - usa dp_clienti_master del gestionale; nessuna sync pesante sulla pagina CLIENTI

// DP SERVICE V309 - Clienti apre archivio unico /clienti-azienda; nessuna sync all'apertura

// DP SERVICE V310 - fix link/redirect archivio clienti master fuori da /service
