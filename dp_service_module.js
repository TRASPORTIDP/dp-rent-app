// DP SERVICE integrato nel gestionale unico

const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');
const PDFDocument = require('pdfkit');

const router = express.Router();
const PORT = process.env.PORT || 10001;

let dataDir = process.env.DP_SERVICE_DATA_DIR || '/var/data/dp_service';
try { fs.mkdirSync(dataDir,{recursive:true}); }
catch(e){ dataDir=path.join(__dirname,'data','dp_service'); fs.mkdirSync(dataDir,{recursive:true}); }
const dpServiceDbPath=path.join(dataDir,'dp_service.sqlite');
const dpServiceSeed=path.join(__dirname,'data','dp_service_seed.sqlite');
try { if(!fs.existsSync(dpServiceDbPath) && fs.existsSync(dpServiceSeed)) fs.copyFileSync(dpServiceSeed,dpServiceDbPath); }
catch(e){ console.log('DP SERVICE seed warning:',e.message); }
const db = new sqlite3.Database(dpServiceDbPath);

router.use(express.urlencoded({ extended: true }));
router.use(express.json());
router.use('/public', express.static(path.join(__dirname, 'public')));

const run = (sql, params=[]) => new Promise((resolve,reject)=>db.run(sql, params, function(err){ if(err) reject(err); else resolve(this); }));
const all = (sql, params=[]) => new Promise((resolve,reject)=>db.all(sql, params, (err,rows)=>err?reject(err):resolve(rows)));
const get = (sql, params=[]) => new Promise((resolve,reject)=>db.get(sql, params, (err,row)=>err?reject(err):resolve(row)));

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
    ['pec','TEXT'],['sdi','TEXT'],['note','TEXT'],['created_at','TEXT'],['codice','TEXT']
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
    ['diagnosi','TEXT'],['whatsapp_pronto_at','TEXT'],['fatturato','INTEGER DEFAULT 0']
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
@media(max-width:780px){.grid,.filters{grid-template-columns:1fr}header span{display:block;margin:6px 0 0}.card{font-size:22px}}
</style></head><body><header><img src="/service/public/dp_service_logo.png" alt="DP SERVICE"><b>DP SERVICE</b><span>Officina • Veicoli • Ricambi • Fatturazione</span></header><main><div class="topnav"><a class="btn dark" href="/service/">🏠 Dashboard</a></div>${body}</main></body></html>`;
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
  const W=doc.page.width, margin=42;
  doc.rect(0,0,W,168).fill('#08090b');
  doc.polygon([W-180,0],[W,0],[W,168],[W-250,168]).fill('#c90000');
  doc.polygon([W-105,0],[W-58,0],[W-150,168],[W-197,168]).fill('#ffffff');
  doc.polygon([W-65,0],[W-20,0],[W-112,168],[W-157,168]).fill('#139447');
  const logoPath=path.join(__dirname,'public','dp_service_logo.png');
  if(fs.existsSync(logoPath)){
    try{ doc.image(logoPath,margin,22,{fit:[176,118],align:'center',valign:'center'}); }catch(e){}
  }
  doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(26).text('DP SERVICE',230,36);
  doc.fillColor('#ff2020').fontSize(11).text('OFFICINA MULTIMARCA',232,72,{characterSpacing:1.2});
  doc.fillColor('#ffffff').font('Helvetica').fontSize(9).text('MANUTENZIONE  •  RIPARAZIONI  •  DIAGNOSI  •  PNEUMATICI',232,94,{width:280});
  const isPrev=String(tipo||'').toUpperCase()==='PREVENTIVO';
  if(isPrev){
    // PREVENTIVO: fascia verde fluo ad alto contrasto, stile DP SERVICE
    doc.save();
    doc.polygon([W-245,48],[W-28,48],[W-43,91],[W-260,91]).fill('#78ff00');
    doc.restore();
    doc.fillColor('#08090b').font('Helvetica-BoldOblique').fontSize(21).text('PREVENTIVO',W-232,58,{width:185,align:'center'});
    doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(12).text(numeroLabel,W-205,101,{width:155,align:'right'});
    doc.font('Helvetica').fontSize(9).text(dataLabel,W-205,121,{width:155,align:'right'});
  } else {
    // FATTURA: fascia rossa piena ad alto contrasto, stile DP SERVICE
    doc.save();
    doc.polygon([W-245,48],[W-28,48],[W-43,91],[W-260,91]).fill('#f20d12');
    doc.restore();
    doc.fillColor('#ffffff').font('Helvetica-BoldOblique').fontSize(22).text('FATTURA',W-232,58,{width:185,align:'center'});
    doc.fillColor('#ffffff').font('Helvetica-Bold').fontSize(12).text(numeroLabel,W-205,101,{width:155,align:'right'});
    doc.font('Helvetica').fontSize(9).text(dataLabel,W-205,121,{width:155,align:'right'});
  }
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

router.get('/', async (req,res)=>{
  const c=await get('SELECT COUNT(*) n FROM clienti');
  const v=await get('SELECT COUNT(*) n FROM veicoli');
  const o=await get("SELECT COUNT(*) n FROM ordini_lavoro WHERE stato IN ('APERTO','IN_LAVORAZIONE','ATTESA_RICAMBI')");
  const r=await get("SELECT COUNT(*) n FROM ricambi WHERE attivo=1");
  const pv=await get('SELECT COUNT(*) n FROM preventivi_service');
  const fsrv=await get('SELECT COUNT(*) n FROM fatture_service');
  res.send(page('Dashboard',`
    <div class="hero"><h1 style="margin:0;font-size:44px">DP SERVICE</h1><div style="font-size:20px;font-weight:700">Gestionale Officina</div></div>
    <div class="grid">
      <a class="card red" href="/service/clienti">👥 Clienti<br><small>${c.n} anagrafiche</small></a>
      <a class="card blue" href="/service/veicoli">🚗 Veicoli clienti<br><small>${v.n} veicoli</small></a>
      <a class="card" href="/service/ricerca">🔎 Ricerca globale<br><small>Targa • Cliente • Modello</small></a>
      <a class="card" href="/service/ordini">🧾 Ordini di lavoro<br><small>${o.n} aperti</small></a>
      <a class="card" href="/service/ricambi">📦 Ricambi / Listino<br><small>${r.n} voci</small></a>
      <a class="card" href="/service/preventivi">📄 Preventivi<br><small>${pv.n} emessi</small></a>
      <a class="card" href="/service/fatture">💶 Fatture serie S<br><small>${fsrv.n} emesse</small></a>
    </div>`));
});

router.get('/clienti', async (req,res)=>{
  const q=(req.query.q||'').trim();
  const rows=q
    ? await all(`SELECT * FROM clienti WHERE ragione_sociale LIKE ? OR piva LIKE ? OR cf LIKE ? OR telefono LIKE ? OR citta LIKE ? ORDER BY ragione_sociale LIMIT 500`,
      Array(5).fill(`%${q}%`))
    : await all(`SELECT * FROM clienti ORDER BY ragione_sociale LIMIT 500`);
  res.send(page('Clienti',`
    <div class="actions"><a class="btn dark" href="/service/">Dashboard</a><a class="btn" href="/service/clienti/nuovo">+ Nuovo cliente</a></div>
    <div class="box"><h1>👥 Clienti (${rows.length})</h1>
      <form class="filters"><input name="q" placeholder="Cliente, P.IVA, C.F., telefono, città" value="${esc(q)}"><span></span><span></span><button class="btn">Cerca</button></form>
      <table><tr><th>Cliente</th><th>P.IVA / C.F.</th><th>Località</th><th>Telefono</th><th></th></tr>
      ${rows.map(x=>`<tr><td><b>${esc(x.ragione_sociale)}</b></td><td>${esc(x.piva)}<br>${esc(x.cf)}</td><td>${esc(x.indirizzo)}<br>${esc(x.citta)} ${esc(x.provincia)}</td><td>${esc(x.telefono)}</td><td><a class="btn dark" href="/service/clienti/${x.id}">Apri</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/clienti/nuovo',(req,res)=>res.send(page('Nuovo cliente',`
  <div class="actions"><a class="btn dark" href="/service/clienti">Indietro</a></div>
  <div class="box"><h1>+ Nuovo cliente</h1>
  <form method="post" action="/service/clienti">
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
    <label>Note</label><textarea name="note"></textarea>
    <p><button class="btn">Salva cliente</button></p>
  </form></div>`)));

router.post('/clienti', async (req,res)=>{
  const b=req.body;
  const r=await run(`INSERT INTO clienti(ragione_sociale,piva,cf,indirizzo,citta,provincia,telefono,email,pec,sdi,note) VALUES(?,?,?,?,?,?,?,?,?,?,?)`,
    [b.ragione_sociale,b.piva,b.cf,b.indirizzo,b.citta,b.provincia,b.telefono,b.email,b.pec,b.sdi,b.note]);
  res.redirect('/service/clienti/'+r.lastID);
});

router.get('/clienti/:id', async (req,res)=>{
  const c=await get('SELECT * FROM clienti WHERE id=?',[req.params.id]);
  if(!c) return res.status(404).send('Cliente non trovato');
  const vs=await all('SELECT * FROM veicoli WHERE cliente_id=? ORDER BY targa',[c.id]);
  res.send(page(c.ragione_sociale,`
    <div class="actions"><a class="btn dark" href="/service/clienti">Indietro</a><a class="btn" href="/service/clienti/${c.id}/modifica">Modifica cliente</a><a class="btn green" href="/service/veicoli/nuovo?cliente_id=${c.id}">+ Aggiungi veicolo</a></div>
    <div class="box"><h1>${esc(c.ragione_sociale)}</h1>
      <p><b>P.IVA:</b> ${esc(c.piva)} &nbsp; <b>C.F.:</b> ${esc(c.cf)}</p>
      <p><b>Indirizzo:</b> ${esc(c.indirizzo)} ${esc(c.citta)} ${esc(c.provincia)}</p>
      <p><b>Telefono:</b> ${esc(c.telefono)} &nbsp; <b>Email:</b> ${esc(c.email)}</p>
      <p><b>PEC:</b> ${esc(c.pec)} &nbsp; <b>SDI:</b> ${esc(c.sdi)}</p>
      <p><b>Note:</b> ${esc(c.note)}</p>
    </div>
    <div class="box"><h2>🚗 Veicoli del cliente (${vs.length})</h2>
      <table><tr><th>Targa</th><th>Marca / Modello</th><th>Telaio</th><th>Km</th><th></th></tr>
      ${vs.map(v=>`<tr><td><b>${esc(v.targa)}</b></td><td>${esc(v.marca)} ${esc(v.modello)}<br>${esc(v.versione)}</td><td>${esc(v.telaio)}</td><td>${v.km||0}</td><td><a class="btn dark" href="/service/veicoli/${v.id}">Apri</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/clienti/:id/modifica', async (req,res)=>{
  const c=await get('SELECT * FROM clienti WHERE id=?',[req.params.id]); if(!c) return res.status(404).send('Cliente non trovato');
  res.send(page('Modifica cliente',`
  <div class="actions"><a class="btn dark" href="/service/clienti/${c.id}">Indietro</a></div>
  <div class="box"><h1>Modifica cliente</h1><form method="post" action="/service/clienti/${c.id}/modifica">
  <label>Ragione sociale / Nome</label><input name="ragione_sociale" value="${esc(c.ragione_sociale)}" required>
  <label>P.IVA</label><input name="piva" value="${esc(c.piva)}"><label>Codice fiscale</label><input name="cf" value="${esc(c.cf)}">
  <label>Indirizzo</label><input name="indirizzo" value="${esc(c.indirizzo)}"><label>Città</label><input name="citta" value="${esc(c.citta)}"><label>Provincia</label><input name="provincia" value="${esc(c.provincia)}">
  <label>Telefono</label><input name="telefono" value="${esc(c.telefono)}"><label>Email</label><input name="email" value="${esc(c.email)}"><label>PEC</label><input name="pec" value="${esc(c.pec)}"><label>SDI</label><input name="sdi" value="${esc(c.sdi)}">
  <label>Note</label><textarea name="note">${esc(c.note)}</textarea><p><button class="btn">Salva modifiche</button></p></form></div>`));
});

router.post('/clienti/:id/modifica', async (req,res)=>{
  const b=req.body;
  await run(`UPDATE clienti SET ragione_sociale=?,piva=?,cf=?,indirizzo=?,citta=?,provincia=?,telefono=?,email=?,pec=?,sdi=?,note=? WHERE id=?`,
    [b.ragione_sociale,b.piva,b.cf,b.indirizzo,b.citta,b.provincia,b.telefono,b.email,b.pec,b.sdi,b.note,req.params.id]);
  res.redirect('/service/clienti/'+req.params.id);
});

router.get('/veicoli', async (req,res)=>{
  const q=(req.query.q||'').trim();
  const params=[]; let where='';
  if(q){ where='WHERE v.targa LIKE ? OR v.marca LIKE ? OR v.modello LIKE ? OR v.versione LIKE ? OR v.telaio LIKE ? OR c.ragione_sociale LIKE ?'; params.push(...Array(6).fill(`%${q}%`)); }
  const rows=await all(`SELECT v.*,c.ragione_sociale FROM veicoli v LEFT JOIN clienti c ON c.id=v.cliente_id ${where} ORDER BY v.targa LIMIT 500`,params);
  res.send(page('Veicoli',`
  <div class="actions"><a class="btn dark" href="/service/">Dashboard</a><a class="btn" href="/service/veicoli/nuovo">+ Nuovo veicolo</a></div>
  <div class="box"><h1>🚗 Veicoli clienti (${rows.length})</h1>
  <form class="filters"><input name="q" placeholder="Targa, cliente, marca, modello, telaio" value="${esc(q)}"><span></span><span></span><button class="btn">Filtra</button></form>
  <table><tr><th>Targa</th><th>Cliente</th><th>Veicolo</th><th>Telaio</th><th>Km</th><th></th></tr>
  ${rows.map(v=>`<tr><td><b>${esc(v.targa)}</b></td><td>${esc(v.ragione_sociale)}</td><td>${esc(v.marca)} ${esc(v.modello)}<br>${esc(v.versione)}</td><td>${esc(v.telaio)}</td><td>${v.km||0}</td><td><a class="btn dark" href="/service/veicoli/${v.id}">Apri</a></td></tr>`).join('')}
  </table></div>`));
});

router.get('/veicoli/nuovo', async (req,res)=>{
  const cs=await all('SELECT id,ragione_sociale FROM clienti ORDER BY ragione_sociale LIMIT 5000');
  const selected=Number(req.query.cliente_id)||0;
  res.send(page('Nuovo veicolo',`
  <div class="actions"><a class="btn dark" href="/service/veicoli">Indietro</a></div>
  <div class="box"><h1>+ Nuovo veicolo cliente</h1><form method="post" action="/service/veicoli">
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
    res.redirect('/service/veicoli/'+r.lastID);
  }catch(e){ res.status(400).send(page('Errore',`<div class="box"><h2>Errore</h2><p>${esc(e.message)}</p><a class="btn" href="/service/veicoli">Torna</a></div>`)); }
});

router.get('/veicoli/:id', async (req,res)=>{
  const v=await get(`SELECT v.*,c.ragione_sociale FROM veicoli v LEFT JOIN clienti c ON c.id=v.cliente_id WHERE v.id=?`,[req.params.id]);
  if(!v) return res.status(404).send('Veicolo non trovato');
  res.send(page(v.targa,`
    <div class="actions"><a class="btn dark" href="/service/veicoli">Indietro</a><a class="btn" href="/service/veicoli/${v.id}/modifica">Modifica veicolo</a>${v.cliente_id?`<a class="btn green" href="/service/clienti/${v.cliente_id}">Apri cliente</a>`:''}</div>
    <div class="box"><h1>🚗 ${esc(v.targa)} — ${esc(v.marca)} ${esc(v.modello)}</h1>
      <p><b>Cliente:</b> ${esc(v.ragione_sociale)}</p><p><b>Versione/Motore:</b> ${esc(v.versione)}</p><p><b>Telaio:</b> ${esc(v.telaio)}</p>
      <p><b>Anno:</b> ${esc(v.anno)} &nbsp; <b>Km:</b> ${v.km||0} &nbsp; <b>Alimentazione:</b> ${esc(v.alimentazione)}</p>
      <p><b>Note:</b> ${esc(v.note)}</p>
      <p><a class="btn" href="/service/ordini/nuovo?veicolo_id=${v.id}">+ Nuovo ordine di lavoro</a></p>
    </div>
    <div class="box"><h2>🧾 Storico interventi</h2>
      <table><tr><th>N.</th><th>Data</th><th>Km</th><th>Lavoro</th><th>Stato</th><th>Totale</th><th></th></tr>
      ${(await all('SELECT * FROM ordini_lavoro WHERE veicolo_id=? ORDER BY id DESC',[v.id])).map(o=>`<tr><td>${o.numero||o.id}/S</td><td>${esc(o.data_apertura)}</td><td>${o.km_ingresso||0}</td><td>${esc(o.descrizione_lavoro)}</td><td>${esc(o.stato)}</td><td>€ ${(Number(o.totale_ivato)||0).toFixed(2)}</td><td><a class="btn dark" href="/service/ordini/${o.id}">Apri</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/veicoli/:id/modifica', async (req,res)=>{
  const v=await get('SELECT * FROM veicoli WHERE id=?',[req.params.id]); if(!v) return res.status(404).send('Veicolo non trovato');
  const cs=await all('SELECT id,ragione_sociale FROM clienti ORDER BY ragione_sociale LIMIT 5000');
  res.send(page('Modifica veicolo',`
  <div class="actions"><a class="btn dark" href="/service/veicoli/${v.id}">Indietro</a></div>
  <div class="box"><h1>Modifica veicolo</h1><form method="post" action="/service/veicoli/${v.id}/modifica">
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
  res.redirect('/service/veicoli/'+req.params.id);
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
    ${arr.length ? arr.map(o=>`<tr><td><b>ODL ${o.numero||o.id}</b></td><td>${esc(o.data_apertura)}</td><td>${esc(o.ragione_sociale)}</td><td><b>${esc(o.targa)}</b><br>${esc(o.marca)} ${esc(o.modello)}</td><td>${o.km_ingresso||0}</td><td><b>${esc(o.stato)}</b></td><td>€ ${(Number(o.totale)||0).toFixed(2)}</td><td><a class="btn dark" href="/service/ordini/${o.id}">Apri</a></td></tr>`).join('') : `<tr><td colspan="8">Nessun ordine</td></tr>`}
    </table>`;

  res.send(page('Ordini di lavoro',`
    <div class="actions"><a class="btn" href="/service/ordini/nuovo">+ Nuovo ordine di lavoro</a></div>
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
  const clienti=await all(`SELECT id,ragione_sociale,piva,telefono FROM clienti ORDER BY ragione_sociale LIMIT 6000`);
  const selectedCliente=Number(req.query.cliente_id)||0;
  res.send(page('Nuovo ordine di lavoro',`
    <div class="actions"><a class="btn dark" href="/service/ordini">Indietro</a></div>
    <div class="box"><h1>+ Accettazione officina</h1>
      <form method="post" action="/service/ordini">
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

  res.redirect('/service/ordini/'+r.lastID);
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
    <div class="actions"><a class="btn dark" href="/service/ordini">Indietro</a><a class="btn" href="/service/veicoli/${o.veicolo_id}">Apri veicolo</a></div>
    <div class="box">
      <h1>🧾 Ordine di lavoro ${o.numero}/${o.anno}</h1>
      <p><b>${esc(o.ragione_sociale)}</b><br>🚗 <b>${esc(o.targa)}</b> — ${esc(o.marca)} ${esc(o.modello)} ${esc(o.versione)}</p>
      <p><b>Data ingresso:</b> ${esc(o.data_apertura)} &nbsp; <b>Km:</b> ${o.km_ingresso||0}</p>
      <p><b>Lavoro richiesto:</b><br>${esc(o.descrizione_lavoro)}</p>
      <p><b>Diagnosi:</b><br>${esc(o.diagnosi)}</p>
      <form method="post" action="/service/ordini/${o.id}/stato" style="margin:14px 0">
        <label>Stato lavorazione</label>
        <div class="filters">
          <select name="stato">${['APERTO','IN_LAVORAZIONE','ATTESA_RICAMBI','PRONTO','CHIUSO'].map(s=>`<option ${o.stato===s?'selected':''}>${s}</option>`).join('')}</select>
          <input type="number" name="km_uscita" placeholder="Km uscita" value="${o.km_uscita||''}">
          <span></span><button class="btn">Aggiorna stato</button>
        </div>
      </form>
      <div class="actions" style="margin-top:12px">
        ${phone?`<a class="btn green" target="_blank" href="https://wa.me/${phone}?text=${text}">📲 AUTO PRONTA - WhatsApp</a>`:''}
        ${preventivo?`<a class="btn dark" href="/service/preventivi/${preventivo.id}">📄 Apri preventivo</a>`:`<form method="post" action="/service/ordini/${o.id}/preventivo"><button class="btn dark">📄 CREA PREVENTIVO</button></form>`}
        ${fattura?`<a class="btn" href="/service/fatture/${fattura.id}">💶 Apri fattura ${fattura.numero}/S</a>`:`<form method="post" action="/service/ordini/${o.id}/fattura"><button class="btn">💶 CREA FATTURA SERIE S</button></form>`}
      </div>
    </div>
    <div class="box"><h2>Ricambi / lavorazioni</h2>
      <form method="post" action="/service/ordini/${o.id}/righe" id="rigaForm">
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
          <a class="btn dark" href="/service/ricambi/nuovo" target="_blank">+ Nuova voce listino</a>
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
      ${righe.map(r=>`<tr><td>${esc(r.tipo)}</td><td>${esc(r.descrizione)}</td><td>${r.quantita}</td><td>€ ${(Number(r.prezzo_unitario)||0).toFixed(2)}</td><td>€ ${((Number(r.quantita)||0)*(Number(r.prezzo_unitario)||0)).toFixed(2)}</td><td><form method="post" action="/service/ordini/${o.id}/righe/${r.id}/elimina"><button class="btn dark">Elimina</button></form></td></tr>`).join('')}
      </table>
      <div style="max-width:430px;margin-left:auto;margin-top:18px;background:#f4f4f4;border-radius:14px;padding:16px"><div><b>Imponibile:</b> ${dpEuro(calc.imponibile)}</div><div><b>IVA:</b> ${dpEuro(calc.iva)}</div><div style="font-size:26px;color:#d40000;margin-top:8px"><b>TOTALE: ${dpEuro(calc.totale)}</b></div></div>
    </div>`));
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
  res.redirect('/service/ordini/'+req.params.id);
});

router.post('/ordini/:id/righe/:rid/elimina', async (req,res)=>{
  await run('DELETE FROM righe_lavoro WHERE id=? AND ordine_id=?',[req.params.rid,req.params.id]);
  const t=await get('SELECT COALESCE(SUM(quantita*prezzo_unitario),0) t FROM righe_lavoro WHERE ordine_id=?',[req.params.id]);
  await run('UPDATE ordini_lavoro SET totale=? WHERE id=?',[t.t,req.params.id]);
  res.redirect('/service/ordini/'+req.params.id);
});

router.post('/ordini/:id/stato', async (req,res)=>{
  const km=Number(req.body.km_uscita)||0;
  const stato=req.body.stato||'APERTO';
  await run('UPDATE ordini_lavoro SET stato=?,km_uscita=?,data_chiusura=? WHERE id=?',[
    stato,km,stato==='CHIUSO'?new Date().toISOString().slice(0,10):null,req.params.id
  ]);
  const o=await get('SELECT veicolo_id FROM ordini_lavoro WHERE id=?',[req.params.id]);
  if(o && km>0) await run('UPDATE veicoli SET km=? WHERE id=?',[km,o.veicolo_id]);
  if(stato==='PRONTO' || stato==='CHIUSO') return res.redirect('/service/ordini#chiusi');
  res.redirect('/service/ordini/'+req.params.id);
});



router.post('/ordini/:id/preventivo', async (req,res)=>{
  const d=await dpServiceDocData(req.params.id);
  if(!d) return res.status(404).send('Ordine non trovato');
  let p=await get('SELECT * FROM preventivi_service WHERE ordine_id=?',[req.params.id]);
  if(!p){
    const anno=new Date().getFullYear();
    const nx=await get('SELECT COALESCE(MAX(numero),0)+1 n FROM preventivi_service WHERE anno=?',[anno]);
    const data=new Date().toISOString().slice(0,10);
    const r=await run(`INSERT INTO preventivi_service(numero,anno,ordine_id,data,imponibile,iva,totale,righe_json,note) VALUES(?,?,?,?,?,?,?,?,?)`,[
      nx.n,anno,req.params.id,data,d.calc.imponibile,d.calc.iva,d.calc.totale,JSON.stringify(d.righe),d.o.note||''
    ]);
    p=await get('SELECT * FROM preventivi_service WHERE id=?',[r.lastID]);
  }
  res.redirect('/service/preventivi/'+p.id);
});

router.post('/ordini/:id/fattura', async (req,res)=>{
  const d=await dpServiceDocData(req.params.id);
  if(!d) return res.status(404).send('Ordine non trovato');
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
  res.redirect('/service/fatture/'+f.id);
});

router.get('/preventivi', async (req,res)=>{
  const rows=await all(`SELECT p.*,o.numero odl,v.targa,c.ragione_sociale FROM preventivi_service p
    LEFT JOIN ordini_lavoro o ON o.id=p.ordine_id LEFT JOIN veicoli v ON v.id=o.veicolo_id LEFT JOIN clienti c ON c.id=o.cliente_id
    ORDER BY p.id DESC LIMIT 500`);
  res.send(page('Preventivi',`<div class="box"><h1>📄 Preventivi DP SERVICE</h1><table><tr><th>N.</th><th>Data</th><th>Cliente</th><th>Targa</th><th>Totale</th><th></th></tr>${rows.map(x=>`<tr><td><b>${x.numero}/P</b></td><td>${dpItDate(x.data)}</td><td>${esc(x.ragione_sociale)}</td><td>${esc(x.targa)}</td><td>${dpEuro(x.totale)}</td><td><a class="btn dark" href="/service/preventivi/${x.id}">Apri</a></td></tr>`).join('')}</table></div>`));
});

router.get('/preventivi/:id([0-9]+)', async (req,res)=>{
  const p=await get('SELECT * FROM preventivi_service WHERE id=?',[req.params.id]); if(!p) return res.status(404).send('Preventivo non trovato');
  const d=await dpServiceDocData(p.ordine_id); if(!d) return res.status(404).send('Ordine non trovato');
  const ph=dpPhone(d.o.telefono); const pdfUrl=`${dpBaseUrl(req)}/preventivi/${p.id}.pdf`;
  const msg=encodeURIComponent(`Buongiorno ${d.o.ragione_sociale||''}, le inviamo il preventivo DP SERVICE n. ${p.numero}/P per la vettura ${d.o.targa||''}. Totale ${dpEuro(p.totale)}. PDF: ${pdfUrl}`);
  res.send(page(`Preventivo ${p.numero}/P`,`<div class="box"><h1>📄 PREVENTIVO ${p.numero}/P</h1><p><b>${esc(d.o.ragione_sociale)}</b> - ${esc(d.o.targa)} - ${esc(d.o.marca)} ${esc(d.o.modello)}</p><h2>Totale ${dpEuro(p.totale)}</h2><div class="actions"><a class="btn" target="_blank" href="/service/preventivi/${p.id}.pdf">📄 PDF</a>${ph?`<a class="btn green" target="_blank" href="https://wa.me/${ph}?text=${msg}">📲 INVIA WHATSAPP</a>`:''}<a class="btn dark" href="/service/ordini/${p.ordine_id}">Torna all'ordine</a></div></div>`));
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
  res.send(page('Fatture serie S',`<div class="box"><h1>💶 Fatture DP SERVICE - Serie S</h1><table><tr><th>N.</th><th>Data</th><th>Cliente</th><th>Targa</th><th>Totale</th><th></th></tr>${rows.map(x=>`<tr><td><b>${x.numero}/S</b></td><td>${dpItDate(x.data)}</td><td>${esc(x.ragione_sociale)}</td><td>${esc(x.targa)}</td><td>${dpEuro(x.totale)}</td><td><a class="btn dark" href="/service/fatture/${x.id}">Apri</a></td></tr>`).join('')}</table></div>`));
});

router.get('/fatture/:id([0-9]+)', async (req,res)=>{
  const f=await get('SELECT * FROM fatture_service WHERE id=?',[req.params.id]); if(!f) return res.status(404).send('Fattura non trovata');
  const d=await dpServiceDocData(f.ordine_id); if(!d) return res.status(404).send('Ordine non trovato');
  const ph=dpPhone(d.o.telefono); const pdfUrl=`${dpBaseUrl(req)}/fatture/${f.id}.pdf`;
  const msg=encodeURIComponent(`Buongiorno ${d.o.ragione_sociale||''}, le inviamo la fattura DP SERVICE n. ${f.numero}/S per la vettura ${d.o.targa||''}. Totale ${dpEuro(f.totale)}. PDF: ${pdfUrl}`);
  res.send(page(`Fattura ${f.numero}/S`,`<div class="box"><h1>💶 FATTURA ${f.numero}/S</h1><p><b>${esc(d.o.ragione_sociale)}</b> - ${esc(d.o.targa)} - ${esc(d.o.marca)} ${esc(d.o.modello)}</p><p><b>Pagamento:</b> ${esc(f.pagamento)}</p><h2>Totale ${dpEuro(f.totale)}</h2><div class="actions"><a class="btn" target="_blank" href="/service/fatture/${f.id}.pdf">📄 PDF</a>${ph?`<a class="btn green" target="_blank" href="https://wa.me/${ph}?text=${msg}">📲 INVIA WHATSAPP</a>`:''}<a class="btn dark" href="/service/ordini/${f.ordine_id}">Torna all'ordine</a></div></div>`));
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
    <div class="actions"><a class="btn dark" href="/service/">Dashboard</a><a class="btn" href="/service/ricambi/nuovo">+ Nuova voce</a></div>
    <div class="box"><h1>📦 Listino ricambi / lavorazioni</h1>
      <form class="filters"><input name="q" placeholder="Codice, descrizione, categoria, marca" value="${esc(q)}"><span></span><span></span><button class="btn">Cerca</button></form>
      <table><tr><th>Codice</th><th>Descrizione</th><th>Categoria</th><th>Marca</th><th>Acquisto</th><th>Vendita</th><th>IVA</th><th></th></tr>
      ${rows.map(x=>`<tr><td>${esc(x.codice)}</td><td><b>${esc(x.descrizione)}</b></td><td>${esc(x.categoria)}</td><td>${esc(x.marca)}</td><td>€ ${(Number(x.prezzo_acquisto)||0).toFixed(2)}</td><td>€ ${(Number(x.prezzo_vendita)||0).toFixed(2)}</td><td>${Number(x.iva)||22}%</td><td><a class="btn dark" href="/service/ricambi/${x.id}/modifica">Modifica</a></td></tr>`).join('')}
      </table>
    </div>`));
});

router.get('/ricambi/nuovo',(req,res)=>res.send(page('Nuova voce listino',`
  <div class="actions"><a class="btn dark" href="/service/ricambi">Indietro</a></div>
  <div class="box"><h1>+ Nuova voce listino</h1>
    <form method="post" action="/service/ricambi">
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
  res.redirect('/service/ricambi');
});

router.get('/ricambi/:id/modifica', async (req,res)=>{
  const x=await get('SELECT * FROM ricambi WHERE id=?',[req.params.id]);
  if(!x) return res.status(404).send('Voce non trovata');
  res.send(page('Modifica voce listino',`
  <div class="actions"><a class="btn dark" href="/service/ricambi">Indietro</a></div>
  <div class="box"><h1>Modifica voce listino</h1>
    <form method="post" action="/service/ricambi/${x.id}/modifica">
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
    <form method="post" action="/service/ricambi/${x.id}/elimina" onsubmit="return confirm('Disattivare questa voce dal listino?')">
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
  res.redirect('/service/ricambi');
});

router.post('/ricambi/:id/elimina', async (req,res)=>{
  await run('UPDATE ricambi SET attivo=0 WHERE id=?',[req.params.id]);
  res.redirect('/service/ricambi');
});

router.get('/ricerca', async (req,res)=>{
  const q=(req.query.q||'').trim(); let rows=[];
  if(q) rows=await all(`SELECT v.*,c.ragione_sociale FROM veicoli v LEFT JOIN clienti c ON c.id=v.cliente_id
    WHERE v.targa LIKE ? OR v.marca LIKE ? OR v.modello LIKE ? OR v.versione LIKE ? OR v.telaio LIKE ? OR c.ragione_sociale LIKE ?
    ORDER BY v.targa LIMIT 300`,Array(6).fill(`%${q}%`));
  res.send(page('Ricerca',`
  <div class="actions"><a class="btn dark" href="/service/">Dashboard</a></div>
  <div class="box"><h1>🔎 Ricerca globale</h1>
  <form class="filters"><input name="q" placeholder="Targa, cliente, marca, modello, telaio" value="${esc(q)}"><span></span><span></span><button class="btn">Cerca</button></form>
  ${q?`<table><tr><th>Targa</th><th>Cliente</th><th>Veicolo</th><th></th></tr>${rows.map(v=>`<tr><td><b>${esc(v.targa)}</b></td><td>${esc(v.ragione_sociale)}</td><td>${esc(v.marca)} ${esc(v.modello)} ${esc(v.versione)}</td><td><a class="btn dark" href="/service/veicoli/${v.id}">Apri</a></td></tr>`).join('')}</table>`:''}
  </div>`));
});

initDb().then(()=>

module.exports = router;
