require('dotenv').config();

const express = require('express');
const bodyParser = require('body-parser');
const twilio = require('twilio');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(bodyParser.urlencoded({ extended: false, limit: '50mb' }));
app.use(bodyParser.json({ limit: '50mb' }));

const PORT = process.env.PORT || 3000;
const APP_BASE_URL = (process.env.APP_BASE_URL || process.env.RENDER_EXTERNAL_URL || '').replace(/\/+$/, '');

const TWILIO_ACCOUNT_SID = process.env.TWILIO_ACCOUNT_SID || '';
const TWILIO_AUTH_TOKEN = process.env.TWILIO_AUTH_TOKEN || '';
const TWILIO_WHATSAPP_NUMBER = process.env.TWILIO_WHATSAPP_NUMBER || 'whatsapp:+390744817108';
const client = twilio(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN);

const DEFAULT_INTERNAL_NUMBERS = ['whatsapp:+393287377675', 'whatsapp:+393472733226', 'whatsapp:+393494040073'];
const INTERNAL_NUMBERS = uniqueNumbers([
  ...parseNumbers(process.env.INTERNAL_GENERAL_NUMBERS || process.env.INTERNAL_NUMBERS || '', []),
  ...DEFAULT_INTERNAL_NUMBERS
]);

const DRIVE_PARENT_FOLDER_ID = process.env.GOOGLE_DRIVE_FOLDER_ID || process.env.DRIVE_PARENT_FOLDER_ID || '';
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || '';
const OPENAI_MODEL = process.env.OPENAI_MODEL || 'gpt-4o-mini';

const DATA_DIR = path.join(__dirname, 'data');
const UPLOAD_DIR = path.join(__dirname, 'uploads');
fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const DB_FILE = path.join(DATA_DIR, 'dp-bot-db.json');
let db = loadDb();
const sessions = {};
const processedSids = new Map();

const E = { ok:'OK', warn:'ATTENZIONE', wrench:'OFFICINA', car:'NOLEGGIO', money:'VENDITA', truck:'TRASPORTO', phone:'TEL', robot:'INFO' };

function loadDb() {
  try { return JSON.parse(fs.readFileSync(DB_FILE, 'utf8')); }
  catch { return { richieste: [], clienti: [], lastId: 0 }; }
}
function saveDb() { fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2), 'utf8'); }
function nextId(prefix) {
  db.lastId = Number(db.lastId || 0) + 1;
  saveDb();
  const d = new Date();
  const ymd = `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
  return `${prefix}-${ymd}-${String(db.lastId).padStart(4,'0')}`;
}
function parseNumbers(value, fallback) {
  if (!value) return fallback;
  return String(value).split(/[;,\n]/).map(s => s.trim()).filter(Boolean).map(n => n.startsWith('whatsapp:') ? n : `whatsapp:${n}`);
}
function uniqueNumbers(list) {
  return [...new Set((list || []).filter(Boolean))];
}
function clean(v) { return String(v || '').trim(); }
function norm(v) { return clean(v).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }
function safeMsg(v) {
  // Solo testo semplice: evita emoji/caratteri corrotti su WhatsApp/Twilio.
  return String(v || '')
    .replace(/â¬/g, 'EUR ')
    .replace(/[ââ]/g, '-')
    .replace(/[ââ]/g, '"')
    .replace(/[ââ]/g, "'")
    .replace(/[^\x09\x0A\x0D\x20-\x7E]/g, '')
    .replace(/\r\n/g, '\n')
    .trim();
}
function yesNo(v) {
  const t = norm(v);
  if (['si','sÃ¬','s','ok','confermo','certo','yes'].includes(t)) return 'SI';
  if (['no','n','annulla','stop'].includes(t)) return 'NO';
  return '';
}
function isDuplicateSid(sid) {
  if (!sid) return false;
  if (processedSids.has(sid)) return true;
  processedSids.set(sid, Date.now());
  const now = Date.now();
  for (const [k, ts] of processedSids.entries()) if (now - ts > 15 * 60 * 1000) processedSids.delete(k);
  return false;
}
function euro(n) { return Number(n || 0).toFixed(2).replace('.', ','); }
function dateIt(d) { return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`; }
function parseDateRange(text) {
  const t = norm(text).replace(/\bdal\b/g,'').replace(/\bal\b/g,'-').replace(/\ba\b/g,'-').replace(/\s+/g,' ').replace(/\s*-\s*/g,'-');
  const m = t.match(/(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?-(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?/);
  if (!m) return null;
  const nowY = new Date().getFullYear();
  const y1 = m[3] ? normalizeYear(m[3]) : nowY;
  const y2 = m[6] ? normalizeYear(m[6]) : y1;
  const start = new Date(y1, Number(m[2])-1, Number(m[1]), 12);
  const end = new Date(y2, Number(m[5])-1, Number(m[4]), 12);
  if (end < start) end.setFullYear(end.getFullYear()+1);
  const days = Math.round((end - start) / 86400000) + 1;
  if (days < 1 || days > 60) return null;
  return { start, end, days, startLabel: dateIt(start), endLabel: dateIt(end) };
}
function normalizeYear(y) { y = String(y); return y.length === 2 ? Number('20' + y) : Number(y); }
function extractKm(v) { const m = clean(v).replace(/\./g,'').match(/\d{1,6}/); return m ? Number(m[0]) : null; }

function menuText() {
  return `*DP RENT / TRASPORTI DP*\n\nScegli il servizio:\n\n1) Officina\n2) Noleggio\n3) Vendita auto\n4) Trasporto veicoli\n5) Altre richieste / parla con assistente DP\n\nScrivi solo il numero.`;
}
function getSession(from, profile) {
  if (!sessions[from] || Date.now() - sessions[from].createdAt > 60 * 60 * 1000) {
    sessions[from] = { state: 'menu', profileName: profile || 'Cliente', answers: [], createdAt: Date.now(), pending: {} };
  }
  return sessions[from];
}
function resetSession(from, profile) { sessions[from] = { state: 'menu', profileName: profile || 'Cliente', answers: [], createdAt: Date.now(), pending: {} }; return sessions[from]; }
function clearSession(from) { delete sessions[from]; }
function touch(s) { s.createdAt = Date.now(); }

function vehicleType(text) {
  const t = norm(text);
  if (t.includes('9') || t.includes('pulmino') || t.includes('persone')) return { key:'pulmino', label:'Pulmino 9 posti', daily:70, includedKm:150 };
  if (t.includes('auto') || t.includes('macchina') || t.includes('golf')) return { key:'auto', label:'Auto', daily:60, includedKm:150 };
  return { key:'furgone', label:'Furgone cargo/merci', daily:70, includedKm:150 };
}
function estimateRental(v, days, km) {
  const base = v.daily * days;
  const included = v.includedKm * days;
  const extraKm = Math.max(0, Number(km || 0) - included);
  const extra = extraKm * 0.15;
  const taxable = base + extra;
  const total = taxable * 1.22;
  return { base, included, extraKm, extra, taxable, iva: taxable*0.22, total };
}
function isAvailable(typeKey, start, end) {
  const s = start.getTime(), e = end.getTime();
  const busy = db.richieste.filter(r => r.tipo === 'noleggio' && ['richiesta','preventivo','confermato','contratto'].includes(r.stato || 'richiesta'));
  // Controllo categoria, non targa: al cliente non mostriamo targhe.
  const overlapping = busy.filter(r => r.mezzoKey === typeKey && r.startTs && r.endTs && !(e < r.startTs || s > r.endTs));
  return { ok: overlapping.length === 0, overlapping };
}

async function sendInternal(title, text, extra = {}) {
  const body = safeMsg(`${title}\n\n${text}`);
  console.log('NOTIFICA INTERNA:', title, 'NUMERI:', INTERNAL_NUMBERS.join(', '), 'TESTO:', body.slice(0, 500));

  for (const to of INTERNAL_NUMBERS) {
    try {
      if (!TWILIO_ACCOUNT_SID || !TWILIO_AUTH_TOKEN || !TWILIO_WHATSAPP_NUMBER) {
        console.error('Twilio non configurato: impossibile inviare a', to);
        continue;
      }
      const msg = await client.messages.create({ from: TWILIO_WHATSAPP_NUMBER, to, body });
      console.log('Notifica WhatsApp inviata:', to, msg.sid);
    } catch (e) {
      console.error('Errore notifica interna', to, e.message, e.code || '');
    }
  }

  // Drive SOLO se richiesto esplicitamente: noleggio/contratti/foto.
  if (extra && extra.saveDrive) {
    await saveRequestNotificationToDrive(title, body, extra).catch(e => console.error('Drive notifica KO:', e.message));
  }
}

async function downloadTwilioMedia(url, filename) {
  const r = await fetch(url, {
    headers: { Authorization: 'Basic ' + Buffer.from(`${TWILIO_ACCOUNT_SID}:${TWILIO_AUTH_TOKEN}`).toString('base64') }
  });
  if (!r.ok) throw new Error(`Download media HTTP ${r.status}`);
  const buf = Buffer.from(await r.arrayBuffer());
  fs.writeFileSync(filename, buf);
  return filename;
}
async function handleIncomingMedia(req, praticaId) {
  const n = Number(req.body.NumMedia || 0);
  if (!n) return [];
  const saved = [];
  const dir = path.join(UPLOAD_DIR, praticaId || 'whatsapp');
  fs.mkdirSync(dir, { recursive: true });
  for (let i=0;i<n;i++) {
    const url = req.body[`MediaUrl${i}`];
    const ct = req.body[`MediaContentType${i}`] || 'application/octet-stream';
    if (!url) continue;
    const ext = ct.includes('png') ? 'png' : ct.includes('pdf') ? 'pdf' : 'jpg';
    const file = path.join(dir, `foto_${Date.now()}_${i+1}.${ext}`);
    try {
      await downloadTwilioMedia(url, file);
      saved.push({ file, contentType: ct });
      await uploadToDrive(file, path.basename(file), praticaId).catch(e => console.error('Drive upload media KO:', e.message));
    } catch (e) { console.error('Media KO', e.message); }
  }
  return saved;
}

async function getDriveClient() {
  if (!DRIVE_PARENT_FOLDER_ID) return null;
  let google;
  try { google = require('googleapis').google; } catch { return null; }
  const credentials = await getGoogleCredentials();
  if (!credentials) return null;
  const auth = new google.auth.JWT(credentials.client_email, null, credentials.private_key, ['https://www.googleapis.com/auth/drive']);
  return google.drive({ version: 'v3', auth });
}

async function getGoogleCredentials() {
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64) return JSON.parse(Buffer.from(process.env.GOOGLE_SERVICE_ACCOUNT_JSON_BASE64, 'base64').toString('utf8'));
  if (process.env.GOOGLE_SERVICE_ACCOUNT_JSON) return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
  if (process.env.GOOGLE_APPLICATION_CREDENTIALS && fs.existsSync(process.env.GOOGLE_APPLICATION_CREDENTIALS)) return JSON.parse(fs.readFileSync(process.env.GOOGLE_APPLICATION_CREDENTIALS, 'utf8'));
  return null;
}

async function getCalendarClient() {
  let google;
  try { google = require('googleapis').google; } catch { return null; }
  const credentials = await getGoogleCredentials();
  if (!credentials) return null;
  const auth = new google.auth.JWT(credentials.client_email, null, credentials.private_key, ['https://www.googleapis.com/auth/calendar']);
  return google.calendar({ version: 'v3', auth });
}

function parseFirstDate(text) {
  const m = String(text || '').match(/(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?/);
  if (!m) return null;
  const y = m[3] ? normalizeYear(m[3]) : new Date().getFullYear();
  const d = new Date(y, Number(m[2]) - 1, Number(m[1]), 9, 0, 0);
  if (isNaN(d.getTime())) return null;
  return d;
}

async function createOfficinaCalendarEvent(richiesta) {
  const calendar = await getCalendarClient();
  if (!calendar) return null;
  const calendarId = process.env.GOOGLE_CALENDAR_ID || process.env.OFFICINA_CALENDAR_ID || 'primary';
  const start = parseFirstDate(richiesta.testo) || new Date(Date.now() + 24 * 60 * 60 * 1000);
  start.setHours(9, 0, 0, 0);
  const end = new Date(start.getTime() + 60 * 60 * 1000);
  const event = await calendar.events.insert({
    calendarId,
    requestBody: {
      summary: `DP SERVICE - ${richiesta.cliente || 'Cliente'} - ${richiesta.id}`,
      description: `Richiesta officina da WhatsApp\n\nCliente: ${richiesta.cliente}\nTelefono: ${richiesta.whatsapp}\nPratica: ${richiesta.id}\n\n${richiesta.testo}`,
      start: { dateTime: start.toISOString(), timeZone: 'Europe/Rome' },
      end: { dateTime: end.toISOString(), timeZone: 'Europe/Rome' }
    }
  });
  return event.data;
}

async function ensureDriveFolder(name) {
  const drive = await getDriveClient();
  if (!drive) return null;
  const q = [`mimeType='application/vnd.google-apps.folder'`, `name='${String(name).replace(/'/g,"\\'")}'`, `'${DRIVE_PARENT_FOLDER_ID}' in parents`, `trashed=false`].join(' and ');
  const existing = await drive.files.list({ q, fields: 'files(id,name)', pageSize: 1 });
  if (existing.data.files && existing.data.files[0]) return existing.data.files[0].id;
  const created = await drive.files.create({ requestBody: { name, mimeType: 'application/vnd.google-apps.folder', parents: [DRIVE_PARENT_FOLDER_ID] }, fields: 'id' });
  return created.data.id;
}
async function uploadToDrive(filePath, name, praticaId) {
  const drive = await getDriveClient();
  if (!drive) return null;
  const folderId = await ensureDriveFolder(praticaId || 'DP_RICHIESTE');
  const r = await drive.files.create({ requestBody: { name, parents: [folderId] }, media: { body: fs.createReadStream(filePath) }, fields: 'id,webViewLink' });
  return r.data;
}
async function saveRequestNotificationToDrive(title, body, extra) {
  const praticaId = extra.praticaId || 'DP_RICHIESTE';
  const dir = path.join(UPLOAD_DIR, praticaId);
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, `richiesta_${Date.now()}.txt`);
  fs.writeFileSync(file, `${title}\n\n${body}\n\n${JSON.stringify(extra, null, 2)}`, 'utf8');
  return uploadToDrive(file, path.basename(file), praticaId);
}

function createRichiesta(data) {
  const id = data.id || nextId(data.prefix || 'DPR');
  const record = { id, createdAt: new Date().toISOString(), stato: 'richiesta', ...data, id };
  db.richieste.push(record);
  saveDb();
  return record;
}

async function askGpt(question, profileName) {
  if (!OPENAI_API_KEY) return 'Richiesta ricevuta. Ti risponderemo appena possibile.';
  const prompt = `Sei l'assistente WhatsApp di Trasporti DP / DP RENT a Narni. Rispondi in italiano, professionale, breve e utile. Non inventare prezzi se non sono dati. Se serve intervento umano, di' che la richiesta viene girata allo staff. Domanda cliente: ${question}`;
  const r = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { Authorization: `Bearer ${OPENAI_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: OPENAI_MODEL, messages: [{ role: 'system', content: prompt }, { role: 'user', content: question }], temperature: 0.4, max_tokens: 250 })
  });
  if (!r.ok) return 'Ho ricevuto la richiesta. La giro allo staff DP e ti rispondiamo appena possibile.';
  const data = await r.json();
  return data.choices?.[0]?.message?.content || 'Richiesta ricevuta.';
}


function looksLikeFreeRequest(text) {
  const t = norm(text);
  if (!t) return false;
  if (/^[1-5]$/.test(t)) return false;
  return t.length >= 4;
}

async function handleGenericRequestFromMenu(session, from, profile, body, twiml, res) {
  const r = createRichiesta({ prefix: 'GEN', tipo: 'generale', cliente: profile, whatsapp: from, testo: body });
  await sendInternal(
    'NUOVA RICHIESTA GENERALE',
    `Cliente: ${profile}\nWhatsApp: ${from}\nPratica: ${r.id}\n\n${body}`,
    { praticaId: r.id, tipo: 'generale' }
  );
  const answer = await askGpt(body, profile);
  clearSession(from);
  twiml.message(safeMsg(`${answer}\n\nHo inviato la richiesta allo staff DP.\nCodice pratica: ${r.id}`));
  return respondTwiml(res, twiml);
}

async function handleWhatsApp(req, res) {
  const twiml = new twilio.twiml.MessagingResponse();
  const from = clean(req.body.From).toLowerCase();
  const body = clean(req.body.Body);
  const profile = req.body.ProfileName || 'Cliente';
  const sid = clean(req.body.MessageSid);
  try {
    if (isDuplicateSid(sid)) return respondTwiml(res, twiml);
    let session = getSession(from, profile);

    const media = await handleIncomingMedia(req, session.pending?.praticaId || 'WHATSAPP_MEDIA');
    if (media.length && session.pending?.praticaId) {
      await sendInternal('FOTO RICEVUTE DP', `Cliente: ${profile}\nWhatsApp: ${from}\nPratica: ${session.pending.praticaId}\nFoto salvate: ${media.length}`, { praticaId: session.pending.praticaId });
      twiml.message(safeMsg(`Foto ricevute e salvate. Totale file: ${media.length}`));
      return respondTwiml(res, twiml);
    }

    if (!body && !media.length) return respondTwiml(res, twiml);
    if (['menu','inizio','ciao','reset','start'].includes(norm(body))) {
      session = resetSession(from, profile);
      if (looksLikeFreeRequest(body)) {
        return await handleGenericRequestFromMenu(session, from, profile, body, twiml, res);
      }
      twiml.message(safeMsg(menuText()));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'menu') {
      const n = norm(body);
      if (n === '1' || n.includes('officina')) {
        session.state = 'officina_info'; touch(session);
        twiml.message(safeMsg('DP SERVICE / Officina\n\nScrivici targa, mezzo, problema e giorno preferito.\nTi richiamiamo appena possibile.'));
        return respondTwiml(res, twiml);
      }
      if (n === '2' || n.includes('noleggio')) {
        session.state = 'rent_vehicle'; touch(session);
        twiml.message(safeMsg('DP RENT - Noleggio\n\nChe mezzo ti serve?\nEsempio: furgone, pulmino 9 posti, auto.'));
        return respondTwiml(res, twiml);
      }
      if (n === '3' || n.includes('vendita')) {
        session.state = 'vendita_info'; touch(session);
        twiml.message(safeMsg('DP AUTO - Vendita auto\n\nGuarda le auto disponibili:\nhttps://autosupermarket.it/concessionario/trasporti-dp-srl/annunci\n\nSe cerchi qualcosa in particolare scrivi modello, budget, permuta e finanziamento SI/NO.'));
        return respondTwiml(res, twiml);
      }
      if (n === '4' || n.includes('trasporto')) {
        session.state = 'trasporto_info'; touch(session);
        twiml.message(safeMsg('Trasporto veicoli\n\nScrivi in un unico messaggio:\n- marca e modello\n- marciante o non marciante\n- luogo ritiro\n- luogo consegna\n- contatto\n\nPrepariamo la quotazione.'));
        return respondTwiml(res, twiml);
      }
      if (n === '5' || n.includes('richiesta') || n.includes('altro')) {
        session.state = 'gpt_chat'; touch(session);
        twiml.message(safeMsg('Altre richieste\n\nScrivimi cosa ti serve. Ti rispondo subito e invio la richiesta allo staff DP.'));
        return respondTwiml(res, twiml);
      }
      twiml.message(safeMsg(menuText()));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'officina_info') {
      const r = createRichiesta({ prefix: 'OFF', tipo: 'officina', cliente: profile, whatsapp: from, testo: body });
      let calInfo = '';
      try {
        const ev = await createOfficinaCalendarEvent(r);
        if (ev?.htmlLink) calInfo = `\nGoogle Calendar: ${ev.htmlLink}`;
      } catch (e) { console.error('Calendar officina KO:', e.message); }
      await sendInternal('NUOVA RICHIESTA OFFICINA', `Cliente: ${profile}\nWhatsApp: ${from}\nPratica: ${r.id}\n\n${body}${calInfo}`, { praticaId: r.id, tipo: 'officina' });
      clearSession(from);
      twiml.message(safeMsg(`Richiesta officina inviata.\nCodice pratica: ${r.id}\nTi richiamiamo appena possibile.`));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'vendita_info') {
      const r = createRichiesta({ prefix: 'VEN', tipo: 'vendita', cliente: profile, whatsapp: from, testo: body });
      await sendInternal('NUOVA RICHIESTA VENDITA AUTO', `Cliente: ${profile}\nWhatsApp: ${from}\nPratica: ${r.id}\n\n${body}`, { praticaId: r.id, tipo: 'vendita' });
      clearSession(from);
      twiml.message(safeMsg(`Richiesta vendita inviata.\nCodice pratica: ${r.id}\nUn responsabile DP AUTO ti rispondera appena possibile.`));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'trasporto_info') {
      const r = createRichiesta({ prefix: 'TRA', tipo: 'trasporto', cliente: profile, whatsapp: from, testo: body });
      await sendInternal('NUOVA RICHIESTA TRASPORTO VEICOLO', `Cliente: ${profile}\nWhatsApp: ${from}\nPratica: ${r.id}\n\nDATI TRASPORTO:\n${body}`, { praticaId: r.id, tipo: 'trasporto' });
      clearSession(from);
      twiml.message(safeMsg(`Richiesta trasporto inviata.\nCodice pratica: ${r.id}\nPrepariamo la quotazione e ti ricontattiamo.`));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'gpt_chat') {
      const r = createRichiesta({ prefix: 'GEN', tipo: 'generale', cliente: profile, whatsapp: from, testo: body });
      await sendInternal('NUOVA RICHIESTA GENERALE / CHATGPT', `Cliente: ${profile}\nWhatsApp: ${from}\nPratica: ${r.id}\n\n${body}`, { praticaId: r.id, tipo: 'generale' });
      const answer = await askGpt(body, profile);
      twiml.message(safeMsg(`${answer}\n\nCodice pratica: ${r.id}`));
      clearSession(from);
      return respondTwiml(res, twiml);
    }

    if (session.state === 'rent_vehicle') {
      const v = vehicleType(body);
      session.pending.vehicle = v;
      session.pending.mezzoKey = v.key;
      session.pending.mezzoLabel = v.label;
      session.state = 'rent_dates'; touch(session);
      twiml.message(safeMsg(`Perfetto: ${v.label}.\n\nOra scrivi le date.\nEsempio: 20/05 - 22/05`));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'rent_dates') {
      const range = parseDateRange(body);
      if (!range) { twiml.message(safeMsg('Date non valide. Scrivile cosi: 20/05 - 22/05')); return respondTwiml(res, twiml); }
      session.pending.range = range;
      session.state = 'rent_km'; touch(session);
      twiml.message(safeMsg('Quanti km prevedi di fare?\nEsempio: 400'));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'rent_km') {
      const km = extractKm(body);
      if (km === null) { twiml.message(safeMsg('Scrivi solo i km previsti. Esempio: 400')); return respondTwiml(res, twiml); }
      const v = session.pending.vehicle;
      const range = session.pending.range;
      const avail = isAvailable(v.key, range.start, range.end);
      if (!avail.ok) {
        session.state = 'rent_dates'; touch(session);
        twiml.message(safeMsg(`Mi dispiace, ${v.label} non risulta disponibile per quelle date.\nProva altre date. Esempio: 25/05 - 27/05`));
        return respondTwiml(res, twiml);
      }
      const quote = estimateRental(v, range.days, km);
      session.pending.km = km;
      session.pending.quote = quote;
      session.state = 'rent_confirm'; touch(session);
      twiml.message(safeMsg(`Disponibile\n\nMezzo: ${v.label}\nPeriodo: ${range.startLabel} - ${range.endLabel}\nGiorni: ${range.days}\nKm previsti: ${km}\n\nPreventivo: EUR ${euro(quote.total)} IVA inclusa\nCauzione: EUR 500,00 gestita separatamente.\n\nConfermi il preventivo?\nRispondi SI oppure NO.`));
      return respondTwiml(res, twiml);
    }

    if (session.state === 'rent_confirm') {
      const yn = yesNo(body);
      if (yn === 'NO') { clearSession(from); twiml.message(safeMsg('Preventivo annullato. Scrivi MENU per ricominciare.')); return respondTwiml(res, twiml); }
      if (yn !== 'SI') { twiml.message(safeMsg('Rispondi SI per confermare oppure NO per annullare.')); return respondTwiml(res, twiml); }

      const praticaId = nextId('DPR');
      session.pending.praticaId = praticaId;
      const range = session.pending.range;
      const r = createRichiesta({
        id: praticaId, prefix: 'DPR', tipo: 'noleggio', stato: 'preventivo', cliente: profile, whatsapp: from,
        mezzoKey: session.pending.mezzoKey, mezzoLabel: session.pending.mezzoLabel,
        startTs: range.start.getTime(), endTs: range.end.getTime(), startLabel: range.startLabel, endLabel: range.endLabel,
        giorni: range.days, km: session.pending.km, totale: session.pending.quote.total
      });
      const link = `${APP_BASE_URL || 'https://dp-rent-app.onrender.com'}/cliente-web?ref=${encodeURIComponent(praticaId)}`;
      await sendInternal('NUOVO PREVENTIVO NOLEGGIO CONFERMATO', `Cliente: ${profile}\nWhatsApp: ${from}\nPratica: ${praticaId}\nMezzo: ${r.mezzoLabel}\nPeriodo: ${r.startLabel} - ${r.endLabel}\nKm: ${r.km}\nTotale: EUR ${euro(r.totale)}\n\nLink cliente:\n${link}`, { praticaId, tipo:'noleggio', saveDrive:true });
      clearSession(from);
      twiml.message(safeMsg(`Perfetto.\n\nPer completare la pratica apri questo link e inserisci tutti i dati richiesti.\nPuoi compilare manualmente e caricare foto documento/patente.\n\n${link}\n\nDP RENT controllera i dati prima del contratto definitivo.`));
      return respondTwiml(res, twiml);
    }

    session = resetSession(from, profile);
    twiml.message(safeMsg(menuText()));
    return respondTwiml(res, twiml);
  } catch (e) {
    console.error('ERRORE WHATSAPP:', e);
    twiml.message(safeMsg('Problema tecnico. Scrivi MENU e riprova.'));
    return respondTwiml(res, twiml);
  }
}
function respondTwiml(res, twiml) { res.writeHead(200, { 'Content-Type': 'text/xml; charset=utf-8' }); return res.end(twiml.toString()); }

app.get('/', (req, res) => res.send('DP RENT BOT V84 online'));
app.get('/health', (req, res) => res.json({ ok: true, version: 'V84', time: new Date().toISOString() }));
app.get('/richieste', (req, res) => res.json(db.richieste));

app.get('/cliente-web', (req, res) => {
  const ref = clean(req.query.ref);
  const r = db.richieste.find(x => x.id === ref);
  if (!r) return res.status(404).send('Pratica non trovata');
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.send(`<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>DP RENT - Dati cliente</title><style>
body{margin:0;font-family:Arial;background:#eef4ff;color:#111}.hero{background:linear-gradient(135deg,#071a48,#0d47a1);color:white;padding:28px;border-radius:0 0 28px 28px}.wrap{padding:18px}.card{background:white;border-radius:24px;padding:20px;box-shadow:0 12px 35px #0002;margin-bottom:18px}label{font-weight:800;display:block;margin-top:14px}input,select,textarea{width:100%;box-sizing:border-box;padding:15px;border:1px solid #d7d7df;border-radius:14px;font-size:18px}button{background:#c91515;color:white;border:0;border-radius:18px;padding:16px 22px;font-size:20px;font-weight:900;margin-top:18px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}@media(max-width:700px){.grid{grid-template-columns:1fr}}</style></head><body><div class="hero"><h1>DP RENT</h1><p>Completa i dati cliente</p></div><div class="wrap"><div class="card"><b>Codice pratica:</b> ${ref}<br><b>Mezzo:</b> ${r.mezzoLabel}<br><b>Periodo:</b> ${r.startLabel} - ${r.endLabel}<br><b>Preventivo:</b> EUR ${euro(r.totale)}</div><form class="card" method="post" action="/cliente-web" enctype="multipart/form-data"><input type="hidden" name="ref" value="${ref}"><div class="grid"><div><label>Nome</label><input name="nome" required></div><div><label>Cognome</label><input name="cognome" required></div><div><label>Telefono</label><input name="telefono"></div><div><label>Email</label><input name="email"></div><div><label>Codice fiscale</label><input name="cf"></div><div><label>Data nascita</label><input name="data_nascita" placeholder="gg/mm/aaaa"></div><div><label>Luogo nascita</label><input name="luogo_nascita"></div><div><label>Cittadinanza</label><input name="cittadinanza" value="ITALIANA"></div></div><label>Indirizzo residenza</label><input name="indirizzo"><div class="grid"><div><label>Citta</label><input name="citta"></div><div><label>CAP</label><input name="cap"></div></div><label>Numero documento</label><input name="documento_numero"><label>Scadenza documento</label><input name="documento_scadenza"><label>Numero patente</label><input name="patente_numero"><label>Scadenza patente</label><input name="patente_scadenza"><label>Fatturazione</label><select name="tipo_cliente"><option>Privato</option><option>Azienda</option></select><label>Ragione sociale / P.IVA / PEC / SDI se azienda</label><textarea name="dati_fattura"></textarea><label>Foto documento fronte</label><input type="file" name="foto" accept="image/*,application/pdf"><label>Foto documento retro</label><input type="file" name="foto" accept="image/*,application/pdf"><label>Foto patente fronte</label><input type="file" name="foto" accept="image/*,application/pdf"><label>Foto patente retro</label><input type="file" name="foto" accept="image/*,application/pdf"><button>Invia dati</button></form></div></body></html>`);
});

const multer = require('multer');
const upload = multer({ dest: UPLOAD_DIR });
app.post('/cliente-web', upload.array('foto', 12), async (req, res) => {
  const ref = clean(req.body.ref);
  const r = db.richieste.find(x => x.id === ref);
  if (!r) return res.status(404).send('Pratica non trovata');
  r.clienteDati = req.body;
  r.stato = 'dati_cliente_inviati';
  r.files = r.files || [];
  for (const f of (req.files || [])) {
    r.files.push({ originalname: f.originalname, path: f.path, mimetype: f.mimetype });
    await uploadToDrive(f.path, f.originalname || path.basename(f.path), ref).catch(e => console.error('Drive cliente KO:', e.message));
  }
  saveDb();
  await sendInternal('DATI CLIENTE NOLEGGIO INVIATI', `Pratica: ${ref}\nCliente: ${req.body.nome || ''} ${req.body.cognome || ''}\nTelefono: ${req.body.telefono || ''}\nEmail: ${req.body.email || ''}\nFile ricevuti: ${(req.files || []).length}`, { praticaId: ref, tipo:'cliente', saveDrive:true });
  res.setHeader('Content-Type','text/html; charset=utf-8');
  res.send('<meta name="viewport" content="width=device-width,initial-scale=1"><div style="font-family:Arial;background:#eef4ff;min-height:100vh;padding:24px"><div style="background:#092b75;color:white;border-radius:24px;padding:24px"><h1>Grazie, dati inviati</h1><p>DP RENT controllera i dati e ti confermera il contratto.</p></div></div>');
});

app.post('/whatsapp', handleWhatsApp);
app.post('/webhook', handleWhatsApp);

app.listen(PORT, () => {
  console.log(`DP RENT BOT V84 online porta ${PORT}`);
  console.log('Numeri interni:', INTERNAL_NUMBERS.join(', '));
});
