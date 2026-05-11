'use strict';

/**
 * DP RENT APP - V1 CLEAN
 * Server unico pronto per Render.
 *
 * Obiettivo: dati salvati sempre in modo stabile.
 * Cliente, mezzo, prenotazione e contratto non vengono più persi quando riapri/modifichi.
 *
 * Dipendenze:
 * npm i express body-parser multer pdfkit nodemailer dotenv
 */

require('dotenv').config();

const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const PDFDocument = require('pdfkit');
const nodemailer = require('nodemailer');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 10000;
const HOST = '0.0.0.0';

// =========================
// CARTELLE
// =========================
const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, 'data');
const UPLOADS_DIR = path.join(ROOT, 'uploads');
const CONTRACTS_DIR = path.join(ROOT, 'contracts');
const DB_FILE = path.join(DATA_DIR, 'db.json');

for (const dir of [DATA_DIR, UPLOADS_DIR, CONTRACTS_DIR]) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

// =========================
// EXPRESS
// =========================
app.use(bodyParser.urlencoded({ extended: true, limit: '25mb' }));
app.use(bodyParser.json({ limit: '25mb' }));
app.use('/uploads', express.static(UPLOADS_DIR));
app.use('/contracts', express.static(CONTRACTS_DIR));

const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOADS_DIR),
    filename: (req, file, cb) => {
      const safe = String(file.originalname || 'file').replace(/[^a-zA-Z0-9._-]/g, '_');
      cb(null, `${Date.now()}-${crypto.randomBytes(4).toString('hex')}-${safe}`);
    }
  }),
  limits: { fileSize: 15 * 1024 * 1024 }
});

// =========================
// DATABASE JSON STABILE
// =========================
function emptyDb() {
  return {
    version: 'DP_RENT_APP_V1_CLEAN',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    counters: {
      contractDailySeq: {},
      bookingSeq: 0,
      customerSeq: 0,
      vehicleSeq: 0
    },
    settings: {
      companyName: 'Trasporti DP S.r.l.',
      brandName: 'DP RENT',
      address: 'Via Tuderte 466, Narni (TR)',
      phone: '0744817108',
      email: 'contabilita@trasportidp.com',
      website: 'www.trasportidp.com',
      vatRate: 22,
      depositDefault: 500,
      kmIncludedPerDay: 150,
      extraKmPrice: 0.15,
      eveningPickupPrice: 30,
      operatorCode: '000000',
      agencyId: '001',
      agencyName: 'Trasporti DP S.r.l.',
      agencyAddress: 'Via Tuderte 466, Narni (TR)',
      agencyPhone: '0744817108'
    },
    customers: [],
    vehicles: [],
    bookings: [],
    audit: []
  };
}

function loadDb() {
  if (!fs.existsSync(DB_FILE)) {
    const db = emptyDb();
    saveDb(db, 'init');
    return db;
  }
  try {
    const raw = fs.readFileSync(DB_FILE, 'utf8');
    const db = JSON.parse(raw);
    return migrateDb(db);
  } catch (err) {
    const backup = path.join(DATA_DIR, `db-broken-${Date.now()}.json`);
    try { fs.copyFileSync(DB_FILE, backup); } catch (_) {}
    const db = emptyDb();
    db.audit.push({ at: new Date().toISOString(), action: 'db_recreated_after_parse_error', error: err.message, backup });
    saveDb(db, 'recreated');
    return db;
  }
}

function saveDb(db, reason = 'save') {
  db.updatedAt = new Date().toISOString();
  const tmp = `${DB_FILE}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(db, null, 2));
  fs.renameSync(tmp, DB_FILE);
  return db;
}

function migrateDb(db) {
  const fresh = emptyDb();
  db = { ...fresh, ...db };
  db.counters = { ...fresh.counters, ...(db.counters || {}) };
  db.settings = { ...fresh.settings, ...(db.settings || {}) };
  db.customers = Array.isArray(db.customers) ? db.customers : [];
  db.vehicles = Array.isArray(db.vehicles) ? db.vehicles : [];
  db.bookings = Array.isArray(db.bookings) ? db.bookings : [];
  db.audit = Array.isArray(db.audit) ? db.audit : [];
  return db;
}

let db = loadDb();

function audit(action, data = {}) {
  db.audit.unshift({ at: new Date().toISOString(), action, data });
  db.audit = db.audit.slice(0, 300);
  saveDb(db, action);
}

// =========================
// UTILS
// =========================
function esc(v) {
  return String(v ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function clean(v) { return String(v ?? '').trim(); }
function upper(v) { return clean(v).toUpperCase(); }
function onlyDigits(v) { return clean(v).replace(/\D/g, ''); }
function todayKey() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}
function newId(prefix) { return `${prefix}-${Date.now()}-${crypto.randomBytes(3).toString('hex')}`; }
function nextContractNumber() {
  const key = todayKey();
  db.counters.contractDailySeq[key] = (db.counters.contractDailySeq[key] || 0) + 1;
  const n = String(db.counters.contractDailySeq[key]).padStart(4, '0');
  return `DPR-${key}-${n}`;
}
function asDateInput(v) {
  if (!v) return '';
  const s = String(v);
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  const m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m) return `${m[3]}-${m[2]}-${m[1]}`;
  return '';
}
function displayDate(v) {
  if (!v) return '';
  const s = String(v);
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return s;
}
function euro(n) { return `€ ${Number(n || 0).toFixed(2).replace('.', ',')}`; }
function parseMoney(v) { return Number(String(v ?? '0').replace(',', '.')) || 0; }

function findCustomer(id) { return db.customers.find(x => x.id === id); }
function findVehicle(id) { return db.vehicles.find(x => x.id === id); }
function findBooking(id) { return db.bookings.find(x => x.id === id || x.contractNumber === id); }

function getBookingFull(id) {
  const booking = findBooking(id);
  if (!booking) return null;
  const customer = findCustomer(booking.customerId) || booking.customerSnapshot || {};
  const vehicle = findVehicle(booking.vehicleId) || booking.vehicleSnapshot || {};
  return { booking, customer, vehicle };
}

function normalizeCustomer(body, existing = {}) {
  return {
    ...existing,
    firstName: upper(body.firstName || body.nome || existing.firstName),
    lastName: upper(body.lastName || body.cognome || existing.lastName),
    phone: onlyDigits(body.phone || body.telefono || existing.phone),
    email: clean(body.email || existing.email),
    fiscalCode: upper(body.fiscalCode || body.codiceFiscale || existing.fiscalCode),
    birthDate: clean(body.birthDate || body.dataNascita || existing.birthDate),
    birthPlace: upper(body.birthPlace || body.luogoNascita || existing.birthPlace),
    birthPlaceCode: clean(body.birthPlaceCode || existing.birthPlaceCode),
    citizenshipCode: clean(body.citizenshipCode || existing.citizenshipCode || '100000100'),
    address: clean(body.address || body.indirizzo || existing.address),
    city: upper(body.city || body.comune || existing.city),
    zip: clean(body.zip || existing.zip),
    province: upper(body.province || existing.province),
    documentType: clean(body.documentType || existing.documentType || 'Carta identità'),
    documentNumber: upper(body.documentNumber || existing.documentNumber),
    documentExpiry: clean(body.documentExpiry || existing.documentExpiry),
    documentIssuePlace: upper(body.documentIssuePlace || existing.documentIssuePlace),
    documentIssuePlaceCode: clean(body.documentIssuePlaceCode || existing.documentIssuePlaceCode),
    licenseNumber: upper(body.licenseNumber || body.numeroPatente || existing.licenseNumber),
    licenseExpiry: clean(body.licenseExpiry || existing.licenseExpiry),
    licenseIssuePlace: upper(body.licenseIssuePlace || existing.licenseIssuePlace),
    licenseIssuePlaceCode: clean(body.licenseIssuePlaceCode || existing.licenseIssuePlaceCode),
    customerType: clean(body.customerType || existing.customerType || 'Privato'),
    businessName: clean(body.businessName || existing.businessName),
    vatNumber: upper(body.vatNumber || existing.vatNumber),
    pec: clean(body.pec || existing.pec),
    sdi: upper(body.sdi || existing.sdi),
    notes: clean(body.customerNotes || existing.notes),
    updatedAt: new Date().toISOString()
  };
}

function customerSnapshot(customer) {
  return JSON.parse(JSON.stringify(customer || {}));
}
function vehicleSnapshot(vehicle) {
  return JSON.parse(JSON.stringify(vehicle || {}));
}

function getOrCreateCustomer(body) {
  const id = clean(body.customerId);
  let c = id ? findCustomer(id) : null;
  if (!c && clean(body.fiscalCode)) c = db.customers.find(x => upper(x.fiscalCode) === upper(body.fiscalCode));
  if (!c && clean(body.phone)) c = db.customers.find(x => onlyDigits(x.phone) === onlyDigits(body.phone));
  if (!c) {
    c = { id: newId('CUS'), createdAt: new Date().toISOString(), files: [] };
    db.customers.push(c);
  }
  Object.assign(c, normalizeCustomer(body, c));
  return c;
}

function requiredCustomerMissing(c) {
  const miss = [];
  if (!c.firstName) miss.push('Nome');
  if (!c.lastName) miss.push('Cognome');
  if (!c.phone) miss.push('Telefono');
  if (!c.fiscalCode) miss.push('Codice fiscale');
  if (!c.documentNumber) miss.push('Numero documento');
  if (!c.licenseNumber) miss.push('Numero patente');
  return miss;
}

function requiredVehicleMissing(v) {
  const miss = [];
  if (!v.brand) miss.push('VEICOLO_MARCA');
  if (!v.model) miss.push('VEICOLO_MODELLO');
  if (!v.plate) miss.push('VEICOLO_TARGA');
  return miss;
}

function calcTotals(booking) {
  const days = Math.max(1, Number(booking.days || 1));
  const daily = Number(booking.dailyPrice || 0);
  const extras = Number(booking.extraPrice || 0);
  const subtotal = (days * daily) + extras;
  const vat = subtotal * ((db.settings.vatRate || 22) / 100);
  const total = subtotal + vat;
  return { days, daily, extras, subtotal, vat, total };
}

// =========================
// HTML LAYOUT
// =========================
function layout(title, body) {
  return `<!doctype html>
<html lang="it">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${esc(title)}</title>
<style>
:root{--red:#b40000;--dark:#1f1f1f;--soft:#f4f4f4;--green:#247a2f;--blue:#0b5cad;}
*{box-sizing:border-box}body{font-family:Arial,Helvetica,sans-serif;margin:0;background:#eee;color:#222}.top{background:#151515;color:white;padding:20px 26px}.brand{display:flex;align-items:center;gap:14px}.logo{width:58px;height:58px;border-radius:14px;background:#fff;color:var(--red);display:flex;align-items:center;justify-content:center;font-weight:900;font-size:24px}.brand h1{margin:0;font-size:30px;letter-spacing:1px}.brand small{font-size:14px;color:#ddd}.nav{background:var(--red);padding:0 18px;display:flex;gap:0;flex-wrap:wrap}.nav a{color:white;text-decoration:none;font-weight:700;padding:18px 16px;display:block}.nav a:hover{background:#830000}.wrap{max-width:1280px;margin:24px auto;padding:0 18px}.card{background:white;border-radius:16px;padding:22px;box-shadow:0 3px 16px rgba(0,0,0,.08);margin-bottom:20px}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:14px}.grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:14px}label{font-weight:700;display:block;margin-bottom:6px}input,select,textarea{width:100%;padding:12px;border:1px solid #bbb;border-radius:9px;font-size:16px;background:white}textarea{min-height:90px}.btn{display:inline-block;border:0;border-radius:10px;padding:12px 16px;font-weight:800;text-decoration:none;cursor:pointer;background:var(--red);color:white;margin:4px}.btn.gray{background:#555}.btn.green{background:var(--green)}.btn.blue{background:var(--blue)}.btn.black{background:#111}.btn.small{font-size:13px;padding:8px 10px}.msg-ok{background:#e7f7e7;color:#146b21;border-left:6px solid var(--green);padding:14px;border-radius:10px;font-weight:700}.msg-ko{background:#fff0f0;color:#9a0000;border-left:6px solid var(--red);padding:14px;border-radius:10px;font-weight:700}table{width:100%;border-collapse:collapse}th,td{padding:10px;border-bottom:1px solid #ddd;text-align:left}th{background:#f7f7f7}.muted{color:#777}.pill{display:inline-block;background:#eee;border-radius:999px;padding:5px 10px;font-weight:700}.pre{background:#201a14;color:#fff;padding:16px;border-radius:10px;white-space:pre-wrap;overflow:auto}.danger{color:#b40000;font-weight:900}.ok{color:#247a2f;font-weight:900}.section-title{margin-top:25px;border-bottom:3px solid #eee;padding-bottom:8px}.actions{display:flex;gap:6px;flex-wrap:wrap}.photos{display:flex;gap:10px;flex-wrap:wrap}.photos img{width:150px;height:110px;object-fit:cover;border-radius:10px;border:1px solid #ccc}@media(max-width:800px){.grid,.grid3{grid-template-columns:1fr}.brand h1{font-size:22px}.nav a{padding:14px 10px}.wrap{padding:0 10px}}
</style>
</head>
<body>
<div class="top"><div class="brand"><div class="logo">DP</div><div><h1>DP RENT APP <small>V1 CLEAN</small></h1><small>Dati stabili: cliente + mezzo + contratto</small></div></div></div>
<div class="nav">
<a href="/">Dashboard</a><a href="/vehicles">Mezzi</a><a href="/customers">Clienti</a><a href="/bookings/new">Nuova prenotazione</a><a href="/bookings">Contratti</a><a href="/test-email">Test Email</a><a href="/backup">Backup</a>
</div>
<div class="wrap">${body}</div>
</body></html>`;
}

function field(name, label, value = '', type = 'text', extra = '') {
  return `<div><label>${esc(label)}</label><input name="${esc(name)}" type="${esc(type)}" value="${esc(value)}" ${extra}></div>`;
}
function selectField(name, label, value, options) {
  return `<div><label>${esc(label)}</label><select name="${esc(name)}">${options.map(o => {
    const val = Array.isArray(o) ? o[0] : o;
    const txt = Array.isArray(o) ? o[1] : o;
    return `<option value="${esc(val)}" ${String(val)===String(value)?'selected':''}>${esc(txt)}</option>`;
  }).join('')}</select></div>`;
}

function customerForm(c = {}) {
  return `
<h2 class="section-title">Cliente / contraente</h2>
<div class="grid">
${field('firstName','Nome',c.firstName)}${field('lastName','Cognome',c.lastName)}
${field('phone','Telefono',c.phone)}${field('email','Email',c.email,'email')}
${field('fiscalCode','Codice fiscale',c.fiscalCode)}${field('birthDate','Data nascita',asDateInput(c.birthDate),'date')}
${field('birthPlace','Luogo nascita',c.birthPlace)}${field('birthPlaceCode','Luogo nascita COD',c.birthPlaceCode)}
${field('citizenshipCode','Cittadinanza COD',c.citizenshipCode || '100000100')}${field('address','Indirizzo residenza',c.address)}
${field('city','Comune residenza',c.city)}${field('province','Provincia',c.province)}
${selectField('documentType','Tipo documento',c.documentType || 'Carta identità', ['Carta identità','Patente','Passaporto'])}${field('documentNumber','Numero documento',c.documentNumber)}
${field('documentExpiry','Scadenza documento',asDateInput(c.documentExpiry),'date')}${field('documentIssuePlace','Luogo rilascio documento',c.documentIssuePlace)}
${field('documentIssuePlaceCode','Luogo rilascio documento COD',c.documentIssuePlaceCode)}${field('licenseNumber','Numero patente',c.licenseNumber)}
${field('licenseExpiry','Scadenza patente',asDateInput(c.licenseExpiry),'date')}${field('licenseIssuePlace','Luogo rilascio patente',c.licenseIssuePlace)}
${field('licenseIssuePlaceCode','Luogo rilascio patente COD',c.licenseIssuePlaceCode)}${selectField('customerType','Tipo cliente',c.customerType || 'Privato', ['Privato','Azienda'])}
${field('businessName','Ragione sociale',c.businessName)}${field('vatNumber','Partita IVA',c.vatNumber)}
${field('pec','PEC',c.pec)}${field('sdi','Codice SDI',c.sdi)}
</div>
<div style="margin-top:12px"><label>Note cliente</label><textarea name="customerNotes">${esc(c.notes)}</textarea></div>`;
}

// =========================
// DASHBOARD
// =========================
app.get('/', (req, res) => {
  const active = db.bookings.filter(b => b.status !== 'Annullato').length;
  const body = `<div class="card"><h2>Dashboard</h2><div class="grid3">
  <div class="card"><h3>Contratti</h3><p style="font-size:34px;font-weight:900">${active}</p></div>
  <div class="card"><h3>Clienti</h3><p style="font-size:34px;font-weight:900">${db.customers.length}</p></div>
  <div class="card"><h3>Mezzi</h3><p style="font-size:34px;font-weight:900">${db.vehicles.length}</p></div>
  </div>
  <p><a class="btn green" href="/bookings/new">+ Nuova prenotazione/contratto</a> <a class="btn blue" href="/vehicles/new">+ Nuovo mezzo</a></p>
  </div>
  <div class="card"><h2>Ultimi contratti</h2>${bookingTable(db.bookings.slice().reverse().slice(0,10))}</div>`;
  res.send(layout('Dashboard', body));
});

// =========================
// VEICOLI
// =========================
app.get('/vehicles', (req, res) => {
  const rows = db.vehicles.map(v => `<tr><td>${esc(v.plate)}</td><td>${esc(v.brand)}</td><td>${esc(v.model)}</td><td>${esc(v.category)}</td><td>${euro(v.dailyPrice)}</td><td><a class="btn small" href="/vehicles/${v.id}/edit">Modifica</a></td></tr>`).join('');
  res.send(layout('Mezzi', `<div class="card"><h2>Mezzi</h2><p><a class="btn green" href="/vehicles/new">+ Nuovo mezzo</a></p><table><tr><th>Targa</th><th>Marca</th><th>Modello</th><th>Categoria</th><th>Prezzo giorno + IVA</th><th></th></tr>${rows || '<tr><td colspan="6">Nessun mezzo inserito.</td></tr>'}</table></div>`));
});
app.get('/vehicles/new', (req, res) => res.send(vehiclePage({}))); 
app.get('/vehicles/:id/edit', (req, res) => {
  const v = findVehicle(req.params.id);
  if (!v) return res.status(404).send(layout('Errore', '<div class="card msg-ko">Mezzo non trovato</div>'));
  res.send(vehiclePage(v));
});
function vehiclePage(v) {
  const isEdit = !!v.id;
  const body = `<div class="card"><h2>${isEdit?'Modifica':'Nuovo'} mezzo</h2><form method="post" action="${isEdit?`/vehicles/${v.id}`:'/vehicles'}">
  <div class="grid">
  ${field('plate','Targa',v.plate)}${field('brand','Marca',v.brand)}
  ${field('model','Modello',v.model)}${field('color','Colore',v.color)}
  ${selectField('category','Categoria',v.category || 'Furgoni', ['Furgoni','9 posti','Auto','Altro'])}${field('dailyPrice','Prezzo giorno + IVA',v.dailyPrice || '70','number','step="0.01"')}
  ${field('kmIncluded','Km inclusi giorno',v.kmIncluded || db.settings.kmIncludedPerDay,'number')}${field('currentKm','Km attuali',v.currentKm,'number')}
  ${field('revisionExpiry','Scadenza revisione',asDateInput(v.revisionExpiry),'date')}${field('insuranceExpiry','Scadenza assicurazione',asDateInput(v.insuranceExpiry),'date')}
  ${field('taxExpiry','Scadenza bollo',asDateInput(v.taxExpiry),'date')}${field('serviceKm','Tagliando a km',v.serviceKm,'number')}
  </div><div style="margin-top:12px"><label>Note mezzo</label><textarea name="notes">${esc(v.notes)}</textarea></div>
  <button class="btn green" type="submit">Salva mezzo</button> <a class="btn gray" href="/vehicles">Indietro</a>
  </form></div>`;
  return layout('Mezzo', body);
}
app.post('/vehicles', (req, res) => {
  const v = normalizeVehicle(req.body, { id: newId('VEH'), createdAt: new Date().toISOString() });
  db.vehicles.push(v); audit('vehicle_created', { id: v.id });
  res.redirect('/vehicles');
});
app.post('/vehicles/:id', (req, res) => {
  const v = findVehicle(req.params.id);
  if (!v) return res.status(404).send('Mezzo non trovato');
  Object.assign(v, normalizeVehicle(req.body, v)); audit('vehicle_updated', { id: v.id });
  res.redirect('/vehicles');
});
function normalizeVehicle(body, existing = {}) {
  return { ...existing,
    plate: upper(body.plate || existing.plate), brand: upper(body.brand || existing.brand), model: upper(body.model || existing.model),
    color: upper(body.color || existing.color), category: clean(body.category || existing.category || 'Furgoni'),
    dailyPrice: parseMoney(body.dailyPrice ?? existing.dailyPrice ?? 70), kmIncluded: Number(body.kmIncluded || existing.kmIncluded || db.settings.kmIncludedPerDay),
    currentKm: Number(body.currentKm || existing.currentKm || 0), revisionExpiry: clean(body.revisionExpiry || existing.revisionExpiry),
    insuranceExpiry: clean(body.insuranceExpiry || existing.insuranceExpiry), taxExpiry: clean(body.taxExpiry || existing.taxExpiry),
    serviceKm: Number(body.serviceKm || existing.serviceKm || 0), notes: clean(body.notes || existing.notes), updatedAt: new Date().toISOString()
  };
}

// =========================
// CLIENTI
// =========================
app.get('/customers', (req, res) => {
  const rows = db.customers.slice().reverse().map(c => `<tr><td>${esc(c.lastName)} ${esc(c.firstName)}</td><td>${esc(c.phone)}</td><td>${esc(c.fiscalCode)}</td><td>${esc(c.documentNumber)}</td><td>${esc(c.licenseNumber)}</td><td><a class="btn small" href="/customers/${c.id}/edit">Modifica</a></td></tr>`).join('');
  res.send(layout('Clienti', `<div class="card"><h2>Clienti</h2><p><a class="btn green" href="/customers/new">+ Nuovo cliente</a></p><table><tr><th>Nome</th><th>Telefono</th><th>CF</th><th>Documento</th><th>Patente</th><th></th></tr>${rows || '<tr><td colspan="6">Nessun cliente.</td></tr>'}</table></div>`));
});
app.get('/customers/new', (req,res)=>res.send(customerPage({})));
app.get('/customers/:id/edit', (req,res)=>{
  const c = findCustomer(req.params.id); if(!c) return res.status(404).send('Cliente non trovato'); res.send(customerPage(c));
});
function customerPage(c){
  const isEdit=!!c.id;
  return layout('Cliente', `<div class="card"><h2>${isEdit?'Modifica':'Nuovo'} cliente</h2><form method="post" action="${isEdit?`/customers/${c.id}`:'/customers'}">${customerForm(c)}<button class="btn green" type="submit">Salva cliente</button></form></div>`);
}
app.post('/customers', (req,res)=>{ const c=getOrCreateCustomer(req.body); audit('customer_saved',{id:c.id}); res.redirect('/customers'); });
app.post('/customers/:id', (req,res)=>{ const c=findCustomer(req.params.id); if(!c)return res.status(404).send('Cliente non trovato'); Object.assign(c,normalizeCustomer(req.body,c)); audit('customer_updated',{id:c.id}); res.redirect('/customers'); });

// =========================
// CONTRATTI / PRENOTAZIONI
// =========================
function bookingTable(list) {
  const rows = list.map(b => {
    const c = findCustomer(b.customerId) || b.customerSnapshot || {};
    const v = findVehicle(b.vehicleId) || b.vehicleSnapshot || {};
    return `<tr><td><b>${esc(b.contractNumber)}</b><br><span class="muted">${esc(b.status)}</span></td><td>${esc(c.lastName)} ${esc(c.firstName)}<br>${esc(c.phone)}</td><td>${esc(v.plate)}<br>${esc(v.brand)} ${esc(v.model)}</td><td>${displayDate(b.startDate)} ${esc(b.startTime||'')}<br>${displayDate(b.endDate)} ${esc(b.endTime||'')}</td><td>${euro((b.totals||calcTotals(b)).total)}</td><td class="actions"><a class="btn small" href="/bookings/${b.id}">Apri</a><a class="btn small blue" href="/bookings/${b.id}/edit">Modifica</a><a class="btn small black" href="/bookings/${b.id}/pdf">PDF</a><a class="btn small gray" href="/bookings/${b.id}/cargos">Ca.R.G.O.S</a></td></tr>`;
  }).join('');
  return `<table><tr><th>Contratto</th><th>Cliente</th><th>Mezzo</th><th>Date</th><th>Totale</th><th></th></tr>${rows || '<tr><td colspan="6">Nessun contratto.</td></tr>'}</table>`;
}
app.get('/bookings', (req,res)=>res.send(layout('Contratti', `<div class="card"><h2>Contratti</h2><p><a class="btn green" href="/bookings/new">+ Nuovo contratto</a></p>${bookingTable(db.bookings.slice().reverse())}</div>`)));
app.get('/bookings/new', (req,res)=>res.send(bookingPage(null)));
app.get('/bookings/:id/edit', (req,res)=>{ const full=getBookingFull(req.params.id); if(!full)return res.status(404).send('Contratto non trovato'); res.send(bookingPage(full)); });

function vehicleOptions(selected) {
  const opts = [['','-- Seleziona mezzo reale --']].concat(db.vehicles.map(v => [v.id, `${v.plate || 'SENZA TARGA'} - ${v.brand || ''} ${v.model || ''} (${v.category || ''})`]));
  return selectField('vehicleId','Mezzo reale obbligatorio',selected,opts);
}

function bookingPage(full) {
  const b = full ? full.booking : {};
  const c = full ? full.customer : {};
  const isEdit = !!b.id;
  const body = `<div class="card"><h2>${isEdit?'Modifica':'Nuovo'} contratto ${esc(b.contractNumber||'')}</h2>
  <form method="post" action="${isEdit?`/bookings/${b.id}`:'/bookings'}" enctype="multipart/form-data">
  <input type="hidden" name="customerId" value="${esc(c.id||'')}">
  ${customerForm(c)}
  <h2 class="section-title">Noleggio</h2>
  <div class="grid">
  ${vehicleOptions(b.vehicleId)}${selectField('paymentMethod','Metodo pagamento',b.paymentMethod || 'Contanti', ['Contanti','Carta di Credito','Bancomat','Bonifico','Nexi PayMail'])}
  ${field('startDate','Data check-out',asDateInput(b.startDate),'date')}${field('startTime','Ora check-out',b.startTime || '09:00','time')}
  ${field('endDate','Data check-in',asDateInput(b.endDate),'date')}${field('endTime','Ora check-in',b.endTime || '18:00','time')}
  ${field('checkoutAddress','Indirizzo check-out',b.checkoutAddress || db.settings.address)}${field('checkinAddress','Indirizzo check-in',b.checkinAddress || db.settings.address)}
  ${field('days','Giorni',b.days || 1,'number','min="1"')}${field('dailyPrice','Prezzo giorno + IVA',b.dailyPrice || 70,'number','step="0.01"')}
  ${field('extraPrice','Extra + IVA',b.extraPrice || 0,'number','step="0.01"')}${field('deposit','Cauzione no IVA',b.deposit || db.settings.depositDefault,'number','step="0.01"')}
  ${field('kmOut','Km uscita',b.kmOut || 0,'number')}${field('fuelOut','Carburante uscita',b.fuelOut || '4/4')}
  ${field('kmIn','Km rientro',b.kmIn || 0,'number')}${field('fuelIn','Carburante rientro',b.fuelIn || '')}
  </div>
  <div style="margin-top:12px"><label>Note contratto</label><textarea name="notes">${esc(b.notes)}</textarea></div>
  <h2 class="section-title">Foto documenti / check-in / check-out</h2>
  <div class="grid3">
  <div><label>Foto documento/patente</label><input type="file" name="customerFiles" multiple accept="image/*,.pdf" capture="environment"></div>
  <div><label>Foto uscita mezzo</label><input type="file" name="checkoutFiles" multiple accept="image/*" capture="environment"></div>
  <div><label>Foto rientro mezzo</label><input type="file" name="checkinFiles" multiple accept="image/*" capture="environment"></div>
  </div>
  <p><button class="btn green" type="submit">Salva contratto</button> <a class="btn gray" href="/bookings">Indietro</a></p>
  </form></div>`;
  return layout('Contratto', body);
}

const bookingUpload = upload.fields([
  { name: 'customerFiles', maxCount: 10 },
  { name: 'checkoutFiles', maxCount: 20 },
  { name: 'checkinFiles', maxCount: 20 }
]);

app.post('/bookings', bookingUpload, (req,res)=>{
  const customer = getOrCreateCustomer(req.body);
  const vehicle = findVehicle(req.body.vehicleId);
  const booking = buildBooking(req.body, customer, vehicle, null, req.files || {});
  booking.id = newId('BKG');
  booking.contractNumber = nextContractNumber();
  booking.createdAt = new Date().toISOString();
  db.bookings.push(booking);
  audit('booking_created', { id: booking.id, contractNumber: booking.contractNumber });
  res.redirect(`/bookings/${booking.id}`);
});
app.post('/bookings/:id', bookingUpload, (req,res)=>{
  const existing = findBooking(req.params.id);
  if(!existing) return res.status(404).send('Contratto non trovato');
  const customer = getOrCreateCustomer({ ...req.body, customerId: existing.customerId || req.body.customerId });
  const vehicle = findVehicle(req.body.vehicleId);
  const updated = buildBooking(req.body, customer, vehicle, existing, req.files || {});
  Object.assign(existing, updated, { id: existing.id, contractNumber: existing.contractNumber, createdAt: existing.createdAt, updatedAt: new Date().toISOString() });
  audit('booking_updated', { id: existing.id, contractNumber: existing.contractNumber });
  res.redirect(`/bookings/${existing.id}`);
});

function filesToStored(files) {
  return (files || []).map(f => ({ filename: f.filename, originalname: f.originalname, mimetype: f.mimetype, size: f.size, url: `/uploads/${f.filename}`, at: new Date().toISOString() }));
}

function buildBooking(body, customer, vehicle, existing = {}, files = {}) {
  const daily = parseMoney(body.dailyPrice || (vehicle && vehicle.dailyPrice) || existing.dailyPrice || 70);
  const booking = {
    ...existing,
    customerId: customer.id,
    vehicleId: vehicle ? vehicle.id : clean(body.vehicleId || existing.vehicleId),
    customerSnapshot: customerSnapshot(customer),
    vehicleSnapshot: vehicleSnapshot(vehicle || findVehicle(existing.vehicleId) || existing.vehicleSnapshot || {}),
    status: clean(body.status || existing.status || 'Aperto'),
    paymentMethod: clean(body.paymentMethod || existing.paymentMethod || 'Contanti'),
    startDate: clean(body.startDate || existing.startDate), startTime: clean(body.startTime || existing.startTime || '09:00'),
    endDate: clean(body.endDate || existing.endDate), endTime: clean(body.endTime || existing.endTime || '18:00'),
    checkoutAddress: clean(body.checkoutAddress || existing.checkoutAddress || db.settings.address),
    checkinAddress: clean(body.checkinAddress || existing.checkinAddress || db.settings.address),
    days: Number(body.days || existing.days || 1), dailyPrice: daily,
    extraPrice: parseMoney(body.extraPrice ?? existing.extraPrice ?? 0), deposit: parseMoney(body.deposit ?? existing.deposit ?? db.settings.depositDefault),
    kmOut: Number(body.kmOut || existing.kmOut || 0), kmIn: Number(body.kmIn || existing.kmIn || 0),
    fuelOut: clean(body.fuelOut || existing.fuelOut || '4/4'), fuelIn: clean(body.fuelIn || existing.fuelIn), notes: clean(body.notes || existing.notes),
    files: { customer: [...(((existing.files||{}).customer)||[]), ...filesToStored(files.customerFiles)], checkout: [...(((existing.files||{}).checkout)||[]), ...filesToStored(files.checkoutFiles)], checkin: [...(((existing.files||{}).checkin)||[]), ...filesToStored(files.checkinFiles)] },
    updatedAt: new Date().toISOString()
  };
  booking.totals = calcTotals(booking);
  return booking;
}

app.get('/bookings/:id', (req,res)=>{
  const full = getBookingFull(req.params.id); if(!full) return res.status(404).send('Contratto non trovato');
  const { booking:b, customer:c, vehicle:v } = full;
  const cmiss = requiredCustomerMissing(c); const vmiss = requiredVehicleMissing(v);
  const photos = (arr=[]) => arr.map(f => f.mimetype && f.mimetype.startsWith('image/') ? `<a href="${esc(f.url)}" target="_blank"><img src="${esc(f.url)}"></a>` : `<a class="btn small gray" href="${esc(f.url)}" target="_blank">${esc(f.originalname)}</a>`).join('');
  const body = `<div class="card"><h2>Contratto ${esc(b.contractNumber)}</h2>${cmiss.length || vmiss.length ? `<div class="msg-ko">Mancano: ${esc([...cmiss,...vmiss].join(', '))}</div>` : '<div class="msg-ok">Dati principali completi.</div>'}
  <p class="actions"><a class="btn blue" href="/bookings/${b.id}/edit">Modifica</a><a class="btn black" href="/bookings/${b.id}/pdf">Scarica PDF</a><a class="btn gray" href="/bookings/${b.id}/cargos">Ca.R.G.O.S</a><a class="btn" href="/bookings">Lista</a></p>
  <div class="grid"><div><h3>Cliente</h3><p><b>${esc(c.lastName)} ${esc(c.firstName)}</b><br>Tel: ${esc(c.phone)}<br>Email: ${esc(c.email)}<br>CF: ${esc(c.fiscalCode)}<br>Documento: ${esc(c.documentType)} ${esc(c.documentNumber)}<br>Patente: ${esc(c.licenseNumber)}</p></div>
  <div><h3>Mezzo</h3><p><b>${esc(v.plate)}</b><br>${esc(v.brand)} ${esc(v.model)}<br>Categoria: ${esc(v.category)}<br>Colore: ${esc(v.color)}</p></div></div>
  <h3>Noleggio</h3><p>Dal ${displayDate(b.startDate)} ${esc(b.startTime)} al ${displayDate(b.endDate)} ${esc(b.endTime)}<br>Uscita: ${esc(b.checkoutAddress)}<br>Rientro: ${esc(b.checkinAddress)}<br>Totale noleggio: <b>${euro(b.totals.total)}</b> - Cauzione: <b>${euro(b.deposit)}</b></p>
  <h3>Foto documenti</h3><div class="photos">${photos((b.files||{}).customer) || '<span class="muted">Nessuna foto.</span>'}</div>
  <h3>Foto uscita</h3><div class="photos">${photos((b.files||{}).checkout) || '<span class="muted">Nessuna foto.</span>'}</div>
  <h3>Foto rientro</h3><div class="photos">${photos((b.files||{}).checkin) || '<span class="muted">Nessuna foto.</span>'}</div>
  </div>`;
  res.send(layout('Contratto', body));
});

// =========================
// PDF CONTRATTO
// =========================
app.get('/bookings/:id/pdf', (req,res)=>{
  const full = getBookingFull(req.params.id); if(!full) return res.status(404).send('Contratto non trovato');
  const { booking:b, customer:c, vehicle:v } = full;
  const file = path.join(CONTRACTS_DIR, `${b.contractNumber}.pdf`);
  createContractPdf(file, b, c, v, (err) => {
    if (err) return res.status(500).send(`Errore PDF: ${err.message}`);
    res.download(file);
  });
});
function createContractPdf(file, b, c, v, cb) {
  const doc = new PDFDocument({ size: 'A4', margin: 45 });
  const stream = fs.createWriteStream(file);
  doc.pipe(stream);
  doc.fontSize(20).fillColor('#b40000').text('DP RENT - CONTRATTO DI NOLEGGIO', { align: 'center' });
  doc.moveDown(0.5).fillColor('#000').fontSize(12).text(`${db.settings.companyName} - ${db.settings.address} - Tel. ${db.settings.phone}`, { align: 'center' });
  doc.moveDown().fontSize(16).text(`Contratto: ${b.contractNumber}`);
  doc.fontSize(11).text(`Data stampa: ${new Date().toLocaleString('it-IT')}`);
  doc.moveDown();
  doc.fontSize(14).fillColor('#b40000').text('Cliente / Contraente'); doc.fillColor('#000').fontSize(11);
  doc.text(`Nome: ${c.lastName || ''} ${c.firstName || ''}`);
  doc.text(`Telefono: ${c.phone || ''}   Email: ${c.email || ''}`);
  doc.text(`Codice fiscale: ${c.fiscalCode || ''}`);
  doc.text(`Nato/a a: ${c.birthPlace || ''} il ${displayDate(c.birthDate)}`);
  doc.text(`Residenza: ${c.address || ''} ${c.city || ''} ${c.province || ''}`);
  doc.text(`Documento: ${c.documentType || ''} n. ${c.documentNumber || ''} scad. ${displayDate(c.documentExpiry)}`);
  doc.text(`Patente: ${c.licenseNumber || ''} scad. ${displayDate(c.licenseExpiry)}`);
  if (c.customerType === 'Azienda') doc.text(`Azienda: ${c.businessName || ''} P.IVA ${c.vatNumber || ''} PEC ${c.pec || ''} SDI ${c.sdi || ''}`);
  doc.moveDown();
  doc.fontSize(14).fillColor('#b40000').text('Veicolo'); doc.fillColor('#000').fontSize(11);
  doc.text(`Targa: ${v.plate || ''}`); doc.text(`Marca/Modello: ${v.brand || ''} ${v.model || ''}`); doc.text(`Categoria: ${v.category || ''} Colore: ${v.color || ''}`);
  doc.moveDown();
  doc.fontSize(14).fillColor('#b40000').text('Periodo e costi'); doc.fillColor('#000').fontSize(11);
  doc.text(`Check-out: ${displayDate(b.startDate)} ${b.startTime || ''} - ${b.checkoutAddress || ''}`);
  doc.text(`Check-in: ${displayDate(b.endDate)} ${b.endTime || ''} - ${b.checkinAddress || ''}`);
  doc.text(`Giorni: ${b.totals.days}  Prezzo giorno + IVA: ${euro(b.dailyPrice)}  Extra + IVA: ${euro(b.extraPrice)}`);
  doc.text(`Imponibile: ${euro(b.totals.subtotal)}  IVA: ${euro(b.totals.vat)}  Totale: ${euro(b.totals.total)}`);
  doc.text(`Cauzione no IVA: ${euro(b.deposit)}  Metodo pagamento: ${b.paymentMethod || ''}`);
  doc.text(`Km uscita: ${b.kmOut || ''}  Carburante uscita: ${b.fuelOut || ''}  Km rientro: ${b.kmIn || ''}  Carburante rientro: ${b.fuelIn || ''}`);
  doc.moveDown();
  doc.fontSize(14).fillColor('#b40000').text('Condizioni principali'); doc.fillColor('#000').fontSize(10);
  doc.text('Il veicolo viene consegnato in buono stato d’uso. Il cliente si impegna alla restituzione nelle stesse condizioni, salvo normale usura. Carburante pieno a pieno se indicato. Eventuali danni, franchigie, multe, pedaggi, ritardi, km extra e costi non saldati restano a carico del cliente.');
  doc.moveDown(2);
  doc.fontSize(11).text('Firma cliente ________________________________', { continued: true }).text('   Firma DP ________________________________');
  doc.end();
  stream.on('finish', () => cb(null));
  stream.on('error', cb);
}

// =========================
// Ca.R.G.O.S EXPORT BASE
// =========================
app.get('/bookings/:id/cargos', (req,res)=>{
  const full = getBookingFull(req.params.id); if(!full) return res.status(404).send('Contratto non trovato');
  const result = buildCargos(full.booking, full.customer, full.vehicle);
  const body = `<div class="card"><h2>Ca.R.G.O.S. ${esc(full.booking.contractNumber)}</h2>${result.missing.length?`<div class="msg-ko"><b>Verifica locale KO</b><br>Mancano campi obbligatori:<div class="pre">${esc(result.missing.join('\n'))}</div></div>`:`<div class="msg-ok">Verifica locale OK</div>`}<p><b>Lunghezza riga:</b> ${result.line.length}</p><div class="pre">${esc(result.line)}</div><p><a class="btn blue" href="/bookings/${full.booking.id}/edit">Correggi dati</a> <a class="btn gray" href="/bookings/${full.booking.id}">Indietro</a></p></div>`;
  res.send(layout('Ca.R.G.O.S', body));
});
function buildCargos(b,c,v){
  const fields = {
    CONTRATTO: b.contractNumber,
    DATA_CONTRATTO: displayDate(b.createdAt || new Date().toISOString()),
    PAGAMENTO: b.paymentMethod || '',
    OPERATORE: db.settings.operatorCode,
    AGENZIA_ID: db.settings.agencyId,
    AGENZIA_NOME: db.settings.agencyName,
    AGENZIA_TEL: db.settings.agencyPhone,
    CHECKOUT_DATA: displayDate(b.startDate), CHECKOUT_ORA: b.startTime || '', CHECKOUT_INDIRIZZO: b.checkoutAddress || '',
    CHECKIN_DATA: displayDate(b.endDate), CHECKIN_ORA: b.endTime || '', CHECKIN_INDIRIZZO: b.checkinAddress || '',
    VEICOLO_TIPO: v.category || '', VEICOLO_MARCA: v.brand || '', VEICOLO_MODELLO: v.model || '', VEICOLO_TARGA: v.plate || '', VEICOLO_COLORE: v.color || '',
    CLIENTE_NOME: c.firstName || '', CLIENTE_COGNOME: c.lastName || '', CLIENTE_CF: c.fiscalCode || '', CLIENTE_TEL: c.phone || '', CLIENTE_EMAIL: c.email || '',
    CLIENTE_NASCITA: displayDate(c.birthDate), CLIENTE_LUOGO_NASCITA: c.birthPlace || '', CLIENTE_LUOGO_NASCITA_COD: c.birthPlaceCode || '', CLIENTE_CITTADINANZA_COD: c.citizenshipCode || '100000100',
    CLIENTE_DOC_TIPO: c.documentType || '', CLIENTE_DOC_NUMERO: c.documentNumber || '', CLIENTE_DOC_SCADENZA: displayDate(c.documentExpiry),
    CLIENTE_PATENTE: c.licenseNumber || '', CLIENTE_PATENTE_SCADENZA: displayDate(c.licenseExpiry)
  };
  const required = ['CONTRATTO','CHECKOUT_DATA','CHECKIN_DATA','VEICOLO_MARCA','VEICOLO_MODELLO','VEICOLO_TARGA','CLIENTE_NOME','CLIENTE_COGNOME','CLIENTE_CF','CLIENTE_DOC_NUMERO','CLIENTE_PATENTE'];
  const missing = required.filter(k => !clean(fields[k]));
  const order = Object.keys(fields);
  const line = order.map(k => String(fields[k] ?? '').replace(/;/g, ',')).join(';');
  return { fields, missing, line };
}

// =========================
// EMAIL / BACKUP
// =========================
app.get('/test-email', (req,res)=>{
  res.send(layout('Test Email', `<div class="card"><h2>Test Email</h2><form method="post" action="/test-email"><div class="grid">${field('to','Invia a',db.settings.email,'email')}</div><button class="btn green" type="submit">Invia test</button></form></div>`));
});
app.post('/test-email', async (req,res)=>{
  try {
    const transporter = createTransporter();
    await transporter.sendMail({ from: process.env.SMTP_FROM || db.settings.email, to: req.body.to, subject: 'Test DP RENT APP', text: 'Email di test inviata da DP RENT APP V1 CLEAN.' });
    res.send(layout('Email OK','<div class="card msg-ok">Email inviata.</div>'));
  } catch(err) { res.send(layout('Email KO',`<div class="card msg-ko">Errore email: ${esc(err.message)}</div>`)); }
});
function createTransporter(){
  return nodemailer.createTransport({ host: process.env.SMTP_HOST, port: Number(process.env.SMTP_PORT || 587), secure: String(process.env.SMTP_SECURE || '').toLowerCase()==='true', auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS } });
}
app.get('/backup', (req,res)=>{
  const file = path.join(DATA_DIR, `backup-${Date.now()}.json`);
  fs.writeFileSync(file, JSON.stringify(db,null,2));
  res.download(file);
});

// =========================
// API DEBUG
// =========================
app.get('/api/db', (req,res)=>res.json(db));
app.get('/health', (req,res)=>res.json({ ok:true, version: db.version, time:new Date().toISOString() }));

app.use((req,res)=>res.status(404).send(layout('404','<div class="card msg-ko">Pagina non trovata.</div>')));

app.listen(PORT, HOST, () => {
  console.log(`DP RENT APP V1 CLEAN avviata su http://${HOST}:${PORT}`);
});