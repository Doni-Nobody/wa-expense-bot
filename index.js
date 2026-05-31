const express = require('express');
const Anthropic = require('@anthropic-ai/sdk');
const { google } = require('googleapis');
const twilio = require('twilio');
const axios = require('axios');
const multer = require('multer');

const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(express.json());

const upload = multer({ storage: multer.memoryStorage() });

// ─── Startup validation ────────────────────────────────────
const REQUIRED_VARS = [
  'ANTHROPIC_API_KEY',
  'TWILIO_ACCOUNT_SID',
  'TWILIO_AUTH_TOKEN',
  'GOOGLE_SHEET_ID',
  'GOOGLE_SERVICE_ACCOUNT_JSON',
];
const missing = REQUIRED_VARS.filter((v) => !process.env[v]);
if (missing.length > 0) {
  console.error('❌ Environment variables belum diisi:', missing.join(', '));
  console.error('   Isi semua variable di Railway → tab Variables, lalu redeploy.');
  process.exit(1);
}

let serviceAccountCredentials;
try {
  serviceAccountCredentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON);
} catch (e) {
  console.error('❌ GOOGLE_SERVICE_ACCOUNT_JSON bukan JSON yang valid.');
  console.error('   Pastikan paste seluruh isi file .json tanpa perubahan apapun.');
  process.exit(1);
}

// ─── Clients ───────────────────────────────────────────────
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const twilioClient = twilio(process.env.TWILIO_ACCOUNT_SID, process.env.TWILIO_AUTH_TOKEN);

// Google Sheets auth via service account
const auth = new google.auth.GoogleAuth({
  credentials: serviceAccountCredentials,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// ─── Claude AI ─────────────────────────────────────────────
const SYSTEM_PROMPT = `Kamu adalah agen pencatat pengeluaran yang ramah. Tugasmu mengekstrak data pengeluaran dari pesan atau foto struk.

Selalu balas HANYA dalam format JSON yang valid:

Jika ada pengeluaran:
{"type":"expense","name":"Nama item/toko","amount":50000,"category":"Makanan & Minuman","note":"catatan opsional","reply":"✅ Tercatat! Makan siang Rp 50.000 di kategori Makanan & Minuman 🍽️"}

Kategori yang tersedia: Makanan & Minuman, Transportasi, Belanja, Tagihan, Kesehatan, Hiburan, Lainnya

Jika pesan adalah /ringkasan atau /summary:
{"type":"summary","reply":"Minta ringkasan"}

Jika tidak ada pengeluaran / tidak jelas:
{"type":"chat","reply":"Pesan balasan yang ramah dan membantu"}

PENTING: Hanya balas dengan JSON, tidak ada teks lain.`;

async function askClaude(text, imageBase64, imageMime) {
  const content = [];

  if (imageBase64) {
    content.push({
      type: 'image',
      source: { type: 'base64', media_type: imageMime || 'image/jpeg', data: imageBase64 },
    });
  }

  content.push({ type: 'text', text: text || 'Ini struk belanja saya, tolong catat semua pengeluarannya.' });

  const msg = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 800,
    system: SYSTEM_PROMPT,
    messages: [{ role: 'user', content }],
  });

  const raw = msg.content.map((b) => b.text || '').join('');
  const cleaned = raw.replace(/```json|```/g, '').trim();
  return JSON.parse(cleaned);
}

// ─── Google Sheets ──────────────────────────────────────────
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_NAME = 'Pengeluaran';

async function ensureSheetExists() {
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
    const sheetNames = meta.data.sheets.map((s) => s.properties.title);
    if (!sheetNames.includes(SHEET_NAME)) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SHEET_ID,
        requestBody: {
          requests: [{ addSheet: { properties: { title: SHEET_NAME } } }],
        },
      });
      // Add header row
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${SHEET_NAME}!A1:G1`,
        valueInputOption: 'RAW',
        requestBody: {
          values: [['Tanggal', 'Jam', 'Nama', 'Kategori', 'Jumlah (Rp)', 'Catatan', 'Pengirim']],
        },
      });
    }
  } catch (err) {
    console.error('ensureSheetExists error:', err.message);
  }
}

async function appendExpense({ name, amount, category, note, sender }) {
  const now = new Date();
  const tanggal = now.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const jam = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' });

  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:G`,
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      values: [[tanggal, jam, name, category, amount, note || '', sender || '']],
    },
  });
}

async function getSummary(sender) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_NAME}!A:G`,
  });
  const rows = (res.data.values || []).slice(1); // skip header

  const now = new Date();
  const thisMonth = `${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

  const monthRows = rows.filter((r) => {
    const parts = (r[0] || '').split('/');
    return parts.length === 3 && `${parts[1]}/${parts[2]}` === thisMonth;
  });

  if (monthRows.length === 0) {
    return `Belum ada pengeluaran bulan ini 😊`;
  }

  const total = monthRows.reduce((s, r) => s + (parseFloat(r[4]) || 0), 0);
  const byCat = {};
  monthRows.forEach((r) => {
    const cat = r[3] || 'Lainnya';
    byCat[cat] = (byCat[cat] || 0) + (parseFloat(r[4]) || 0);
  });

  const catLines = Object.entries(byCat)
    .sort((a, b) => b[1] - a[1])
    .map(([cat, amt]) => `  • ${cat}: Rp ${amt.toLocaleString('id-ID')}`)
    .join('\n');

  return `📊 *Ringkasan Bulan Ini*\n\n${catLines}\n\n💰 *Total: Rp ${total.toLocaleString('id-ID')}*\n📝 ${monthRows.length} transaksi`;
}

// ─── WhatsApp Webhook ───────────────────────────────────────
app.post('/webhook', async (req, res) => {
  const { Body: body, From: from, MediaUrl0: mediaUrl, MediaContentType0: mediaType } = req.body;

  const sender = from || 'unknown';
  let replyText = '';

  try {
    let imageBase64 = null;
    let imageMime = null;

    // Download image if attached
    if (mediaUrl) {
      const imgRes = await axios.get(mediaUrl, {
        responseType: 'arraybuffer',
        auth: {
          username: process.env.TWILIO_ACCOUNT_SID,
          password: process.env.TWILIO_AUTH_TOKEN,
        },
      });
      imageBase64 = Buffer.from(imgRes.data).toString('base64');
      imageMime = mediaType || 'image/jpeg';
    }

    const parsed = await askClaude(body, imageBase64, imageMime);

    if (parsed.type === 'expense') {
      await appendExpense({
        name: parsed.name,
        amount: parsed.amount,
        category: parsed.category,
        note: parsed.note,
        sender,
      });
      replyText = parsed.reply;
    } else if (parsed.type === 'summary') {
      replyText = await getSummary(sender);
    } else {
      replyText = parsed.reply;
    }
  } catch (err) {
    console.error('Webhook error:', err);
    replyText = '❌ Maaf, terjadi kesalahan. Coba kirim lagi ya!';
  }

  // Reply via Twilio WhatsApp
  const twiml = new twilio.twiml.MessagingResponse();
  twiml.message(replyText);
  res.type('text/xml').send(twiml.toString());
});

// ─── Health check ────────────────────────────────────────────
app.get('/', (req, res) => res.json({ status: 'ok', service: 'WA Expense Bot' }));

// ─── Start ───────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`✅ Server jalan di port ${PORT}`);
  await ensureSheetExists();
  console.log('✅ Google Sheets siap');
});
