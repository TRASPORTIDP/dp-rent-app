require('dotenv').config();

const express = require('express');
const bodyParser = require('body-parser');
const nodemailer = require('nodemailer');
const multer = require('multer');

const app = express();

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// =====================
// CONFIG BASE
// =====================

const PORT = process.env.PORT || 10000;

// =====================
// STORAGE FILE
// =====================

const upload = multer({ dest: 'uploads/' });

// =====================
// EMAIL (KELIWEB)
// =====================

async function sendEmail(to, subject, text) {

  if (!process.env.SMTP_HOST) {
    throw new Error('SMTP non configurato');
  }

  const port = Number(process.env.SMTP_PORT || 465);
  const secure = port === 465;

  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: port,
    secure: secure,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    },
    tls: {
      rejectUnauthorized: false
    }
  });

  await transporter.verify();

  return transporter.sendMail({
    from: process.env.SMTP_FROM || process.env.SMTP_USER,
    to,
    subject,
    text
  });
}
// =====================
// TEST EMAIL
// =====================

app.get('/test-email', async (req, res) => {
  try {

    await sendEmail(
      process.env.ALERT_EMAIL || process.env.SMTP_USER,
      'TEST EMAIL DP RENT',
      'Se ricevi questa email funziona tutto'
    );

    res.send('EMAIL OK');

  } catch (err) {
    console.log(err);
    res.send('ERRORE EMAIL: ' + err.message);
  }
});
app.get('/', (req, res) => {
  res.send('DP RENT APP ATTIVA');
});
app.listen(PORT, '0.0.0.0', () => {
  console.log('Server avviato sulla porta ' + PORT);
});
