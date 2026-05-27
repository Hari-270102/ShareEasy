// =============================================
// SHAREEASY - SERVER
// =============================================
// This is the brain of the app.
// It handles: file uploads, AI bot commands,
// format conversion, and sharing links.
// =============================================

require('dotenv').config();

const express  = require('express');
const multer   = require('multer');
const sharp    = require('sharp');
const XLSX     = require('xlsx');
const mammoth  = require('mammoth');
const PDFDoc   = require('pdfkit');
const pdfParse = require('pdf-parse');
const { Document, Packer, Paragraph, TextRun, ImageRun } = require('docx');
const path     = require('path');
const fs       = require('fs');
const { v4: uuidv4 } = require('uuid');
const { GoogleGenerativeAI } = require('@google/generative-ai');
// NOTE: Gemini is kept for future use (when deploying online).
// Currently using local keyword detection (works without internet).
const cors       = require('cors');

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// ── Make sure folders exist ───────────────────
['uploads', 'converted'].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir);
});

// ── File upload configuration ─────────────────
// Multer saves uploaded files to the "uploads" folder
// Each file gets a unique random name so nothing overwrites
const storage = multer.diskStorage({
  destination: 'uploads/',
  filename: (req, file, cb) => {
    const uniqueName = uuidv4() + path.extname(file.originalname);
    cb(null, uniqueName);
  }
});
const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 } // max 50MB
});

// ── Brevo email helper (HTTP API, works on Render) ───────────────────
async function sendBrevoEmail({ to, toName, subject, html, attachmentPath, attachmentName }) {
  const body = {
    sender   : { name: 'ShareEasy', email: 'harikrishnaunofficial@gmail.com' },
    to       : [{ email: to, name: toName || to }],
    subject,
    htmlContent: html
  };
  if (attachmentPath) {
    const fileBuffer = fs.readFileSync(attachmentPath);
    const fileSizeMB = fileBuffer.length / (1024 * 1024);
    if (fileSizeMB <= 10) {
      body.attachment = [{ name: attachmentName, content: fileBuffer.toString('base64') }];
    }
  }
  const res = await fetch('https://api.brevo.com/v3/smtp/email', {
    method : 'POST',
    headers: { 'api-key': process.env.BREVO_API_KEY, 'Content-Type': 'application/json' },
    body   : JSON.stringify(body)
  });
  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.message || 'Brevo send failed');
  }
  return true;
}

// ── In-memory list of shared files ───────────
// (works for local use; for production use a database)
let sharedFiles = [];


// ==============================================
// LOCAL KEYWORD BOT (fallback)
// ==============================================

const FORMAT_MAP = [
  { keys: ['pdf'],                          format: 'PDF'  },
  { keys: ['word', 'docx', 'doc'],          format: 'DOCX' },
  { keys: ['text', 'txt'],                  format: 'TXT'  },
  { keys: ['png'],                          format: 'PNG'  },
  { keys: ['jpg', 'jpeg', 'image'],         format: 'JPG'  },
  { keys: ['webp'],                         format: 'WEBP' },
  { keys: ['xlsx', 'excel', 'xls'],         format: 'XLSX' },
  { keys: ['csv', 'spreadsheet'],           format: 'CSV'  },
];

function localBot(message, currentFormat, currentRecipient, currentRecipientEmail) {
  const msg   = message.toLowerCase().trim();
  const words = msg.split(/\s+/);

  let detectedFormat = null;
  for (const word of words) {
    for (const entry of FORMAT_MAP) {
      if (entry.keys.includes(word)) { detectedFormat = entry.format; break; }
    }
    if (detectedFormat) break;
  }

  let detectedEmail = null;
  const emailMatch = msg.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/);
  if (emailMatch) detectedEmail = emailMatch[0];

  let detectedRecipient = null;
  if (detectedEmail) {
    detectedRecipient = detectedEmail.split('@')[0].replace(/[._\-]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
  } else {
    const m = msg.match(/(?:to|for|send\s+to|share\s+with)\s+([a-z][a-z\s]{1,20}?)(?:\s+(?:as|in|the)|[.,!?]|$)/i);
    if (m) detectedRecipient = m[1].trim().replace(/\b\w/g, c => c.toUpperCase());
  }

  const format    = detectedFormat    || currentFormat;
  const recipient = detectedRecipient || currentRecipient;
  const recipientEmail = detectedEmail || currentRecipientEmail || null;

  let reply;
  if (format && (recipient || recipientEmail)) {
    reply = `Got it! I'll convert to ${format} and send it to ${recipientEmail || recipient}. Click "Send now" to confirm!`;
  } else if (format) {
    reply = `Sure, converting to ${format}. Who should I send it to?`;
  } else if (recipient || recipientEmail) {
    reply = `I'll send it to ${recipientEmail || recipient}. What format? (PDF, DOCX, TXT, PNG, JPG, XLSX, CSV)`;
  } else {
    reply = `Try: "Send this to hari@gmail.com as PDF"`;
  }
  return { reply, format, recipient, recipientEmail, suggestedFormat: null };
}

// ==============================================
// GROQ AI BOT (primary, with local fallback)
// ==============================================

const Groq = require('groq-sdk');
const groq = new Groq({ apiKey: process.env.GROQ_API_KEY });

async function groqBot(message, currentFormat, currentRecipient, currentRecipientEmail) {
  const prompt = `You are ShareEasy's assistant. Users want to convert and send files.
Your job: extract intent from the user's message and respond with ONLY valid JSON.

Context from previous turns:
- Current format: ${currentFormat || 'not set'}
- Current recipient name: ${currentRecipient || 'not set'}
- Current recipient email: ${currentRecipientEmail || 'not set'}

Supported formats: PDF, DOCX, TXT, PNG, JPG, WEBP, XLSX, CSV

User message: "${message}"

Rules:
1. Extract the target file format if mentioned (map "word" → DOCX, "excel" → XLSX, "image" → JPG, "text" → TXT).
2. Extract recipient name and/or email if mentioned.
3. Keep context from previous turns if the user doesn't change it.
4. If both format and recipient are known, say you'll convert and send — tell them to click "Send now".
5. If only format is known, ask who to send it to.
6. If only recipient is known, ask what format they want.
7. If neither, ask for both in a friendly way.
8. Be conversational and friendly, like a helpful assistant.

Respond with ONLY this JSON (no markdown, no code blocks):
{
  "reply": "your friendly message to the user",
  "format": "PDF" or "DOCX" or "TXT" or "PNG" or "JPG" or "WEBP" or "XLSX" or "CSV" or null,
  "recipient": "recipient display name" or null,
  "recipientEmail": "recipient@email.com" or null
}`;

  const completion = await groq.chat.completions.create({
    model   : 'llama-3.3-70b-versatile',
    messages: [{ role: 'user', content: prompt }],
    temperature: 0.3
  });

  const text    = completion.choices[0].message.content.trim();
  const cleaned = text.replace(/^```(?:json)?\s*/i, '').replace(/\s*```$/, '').trim();
  const data    = JSON.parse(cleaned);

  return {
    reply        : data.reply || "I didn't quite get that. Try: \"Send this to hari@gmail.com as PDF\".",
    format       : data.format || currentFormat || null,
    recipient    : data.recipient || currentRecipient || null,
    recipientEmail: data.recipientEmail || currentRecipientEmail || null,
    suggestedFormat: null
  };
}


// ==============================================
// ROUTES
// ==============================================

// ── 1. Upload a file ──────────────────────────
app.post('/api/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file was uploaded.' });
  }
  res.json({
    success: true,
    fileId: req.file.filename,
    originalName: req.file.originalname,
    size: req.file.size
  });
});

// ── 2. Bot chat (Gemini AI) ──
app.post('/api/bot', async (req, res) => {
  const { message, currentFormat = null, currentRecipient = null, currentRecipientEmail = null } = req.body;
  if (!message || !message.trim()) {
    return res.json({ reply: 'Please type a message first!', format: null, recipient: null, recipientEmail: null, suggestedFormat: null });
  }
  try {
    const result = await groqBot(message, currentFormat, currentRecipient, currentRecipientEmail);
    console.log('🤖 Bot: Groq AI');
    res.json(result);
  } catch (err) {
    console.error('Groq error:', err.message);
    console.log('🤖 Bot: local keyword fallback');
    const result = localBot(message, currentFormat, currentRecipient, currentRecipientEmail);
    res.json(result);
  }
});

// ── 3. Convert file and create share link ─────
app.post('/api/convert-and-share', async (req, res) => {
  const { fileId, originalName, targetFormat, recipientName, recipientEmail } = req.body;

  if (!fileId || !targetFormat) {
    return res.status(400).json({ error: 'Missing file or target format.' });
  }

  const inputPath = path.join('uploads', fileId);
  if (!fs.existsSync(inputPath)) {
    return res.status(404).json({ error: 'Uploaded file not found. Please upload again.' });
  }

  const shareId  = uuidv4();
  const outExt   = targetFormat.toLowerCase();
  const outFile  = shareId + '.' + outExt;
  const outPath  = path.join('converted', outFile);
  const fromExt  = path.extname(originalName).replace('.', '').toLowerCase();
  const baseName = path.basename(originalName, path.extname(originalName));

  try {
    await convertFile(inputPath, outPath, fromExt, outExt);

    const entry = {
      shareId,
      displayName : baseName + '.' + outExt,
      convertedFrom: fromExt.toUpperCase(),
      convertedTo : targetFormat.toUpperCase(),
      recipientName : recipientName || 'Recipient',
      recipientEmail: recipientEmail || '',
      filePath : outPath,
      fileSize : fs.statSync(outPath).size,
      createdAt: new Date().toISOString(),
      status   : 'Delivered'
    };

    sharedFiles.unshift(entry);

    // ── Send email with file attached (if recipient email provided) ──
    let emailStatus = 'no_email';
    if (recipientEmail && process.env.BREVO_API_KEY) {
      try {
        const shareLink = `${req.protocol}://${req.get('host')}/share/${shareId}`;
        await sendBrevoEmail({
          to            : recipientEmail,
          toName        : recipientName || recipientEmail,
          subject       : `${recipientName || 'Someone'} shared a file with you via ShareEasy`,
          html          : `
            <div style="font-family:sans-serif;max-width:520px;margin:0 auto">
              <h2 style="color:#1a1a18">⚡ ShareEasy</h2>
              <p>Hi <strong>${recipientName || 'there'}</strong>,</p>
              <p>Someone has shared a converted file with you.</p>
              <table style="background:#f5f4f0;border-radius:8px;padding:12px 16px;margin:16px 0;width:100%">
                <tr><td><strong>File:</strong></td><td>${entry.displayName}</td></tr>
                <tr><td><strong>Converted:</strong></td><td>${entry.convertedFrom} → ${entry.convertedTo}</td></tr>
              </table>
              <p>📎 The file is attached below (if under 10MB).</p>
              <p>Or <a href="${shareLink}">click here to download it</a> anytime.</p>
              <hr style="border:none;border-top:1px solid #eee;margin:20px 0"/>
              <p style="color:#aaa;font-size:12px">Sent via ShareEasy &mdash; convert &amp; share files instantly</p>
            </div>`,
          attachmentPath: outPath,
          attachmentName: entry.displayName
        });
        emailStatus = 'sent';
        console.log(`  ✉️  Email sent to ${recipientEmail}`);
      } catch (mailErr) {
        console.error('Email error:', mailErr.message);
        emailStatus = 'failed';
      }
    }

    res.json({
      success    : true,
      shareId,
      shareLink  : `${req.protocol}://${req.get('host')}/share/${shareId}`,
      emailStatus,
      message    : `✅ File converted to ${targetFormat} and shared successfully!`
    });
  } catch (err) {
    console.error('Conversion error:', err.message);
    res.status(500).json({ error: 'Conversion failed: ' + err.message });
  }
});

// ── 4. List all shared files ──────────────────
app.get('/api/files', (req, res) => {
  // Strip the internal file path before sending to frontend
  const safe = sharedFiles.map(({ filePath, ...rest }) => rest);
  res.json(safe);
});
// ── 4b. Get a single shared file by shareId ─────
app.get('/api/file/:shareId', (req, res) => {
  const entry = sharedFiles.find(f => f.shareId === req.params.shareId);
  if (!entry) return res.status(404).json({ error: 'File not found or expired.' });
  const { filePath, ...safe } = entry;
  res.json(safe);
});
// ── 5. Download a file ────────────────────────
app.get('/api/download/:shareId', (req, res) => {
  const entry = sharedFiles.find(f => f.shareId === req.params.shareId);
  if (!entry) return res.status(404).send('File not found or expired.');
  res.download(entry.filePath, entry.displayName);
});

// ── Serve dedicated recipient download page ───
app.get('/share/:shareId', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'share.html'));
});

// ── Catch-all: serve index.html for /receive ──
app.get('/receive', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});


// ==============================================
// FILE CONVERSION LOGIC
// ==============================================

async function convertFile(inputPath, outputPath, from, to) {
  const imageTypes = ['jpg', 'jpeg', 'png', 'webp', 'gif', 'tiff', 'bmp'];
  const isImgIn  = imageTypes.includes(from);
  const isImgOut = imageTypes.includes(to);

  // ── Image → Image (e.g. PNG → JPG) ──────────
  if (isImgIn && isImgOut) {
    const fmt = to === 'jpg' ? 'jpeg' : to;
    await sharp(inputPath).toFormat(fmt).toFile(outputPath);
    return;
  }

  // ── Image → PDF ──────────────────────────────
  if (isImgIn && to === 'pdf') {
    await imageToPDF(inputPath, outputPath, from);
    return;
  }

  // ── Image → DOCX ─────────────────────────────
  if (isImgIn && to === 'docx') {
    await imageToDocx(inputPath, outputPath);
    return;
  }

  // ── PDF → TXT ────────────────────────────────
  if (from === 'pdf' && to === 'txt') {
    const pdfBuf = fs.readFileSync(inputPath);
    const data   = await pdfParse(pdfBuf);
    fs.writeFileSync(outputPath, data.text);
    return;
  }

  // ── PDF → DOCX ───────────────────────────────
  if (from === 'pdf' && to === 'docx') {
    const pdfBuf = fs.readFileSync(inputPath);
    const data   = await pdfParse(pdfBuf);
    await textToDocx(data.text, outputPath);
    return;
  }

  // ── CSV → XLSX ───────────────────────────────
  if (from === 'csv' && to === 'xlsx') {
    const csv = fs.readFileSync(inputPath, 'utf8');
    const wb  = XLSX.read(csv, { type: 'string' });
    XLSX.writeFile(wb, outputPath);
    return;
  }

  // ── CSV → PDF ───────────────────────────────
  if (from === 'csv' && to === 'pdf') {
    const csv = fs.readFileSync(inputPath, 'utf8');
    const wb  = XLSX.read(csv, { type: 'string' });
    const ws  = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    await spreadsheetToPDF(rows, outputPath);
    return;
  }

  // ── XLSX → CSV ───────────────────────────────
  if (from === 'xlsx' && to === 'csv') {
    const wb  = XLSX.readFile(inputPath);
    const ws  = wb.Sheets[wb.SheetNames[0]];
    fs.writeFileSync(outputPath, XLSX.utils.sheet_to_csv(ws));
    return;
  }

  // ── XLSX → PDF ───────────────────────────────
  if (from === 'xlsx' && to === 'pdf') {
    const wb   = XLSX.readFile(inputPath);
    const ws   = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
    await spreadsheetToPDF(rows, outputPath);
    return;
  }

  // ── DOCX → TXT ───────────────────────────────
  if (from === 'docx' && to === 'txt') {
    const result = await mammoth.extractRawText({ path: inputPath });
    fs.writeFileSync(outputPath, result.value);
    return;
  }

  // ── DOCX → PDF ───────────────────────────────
  if (from === 'docx' && to === 'pdf') {
    const result = await mammoth.extractRawText({ path: inputPath });
    await createPDF(result.value, outputPath);
    return;
  }

  // ── TXT → PDF ────────────────────────────────
  if (from === 'txt' && to === 'pdf') {
    const text = fs.readFileSync(inputPath, 'utf8');
    await createPDF(text, outputPath);
    return;
  }

  // ── TXT → DOCX ───────────────────────────────
  if (from === 'txt' && to === 'docx') {
    const text = fs.readFileSync(inputPath, 'utf8');
    await textToDocx(text, outputPath);
    return;
  }

  // ── Same format — just copy ──────────────────
  fs.copyFileSync(inputPath, outputPath);
}

// Embeds an image into a real PDF file
async function imageToPDF(inputPath, outputPath, fromExt) {
  return new Promise(async (resolve, reject) => {
    try {
      const meta = await sharp(inputPath).metadata();
      const imgW = meta.width  || 800;
      const imgH = meta.height || 600;

      // Always convert to JPEG for consistent PDFKit embedding
      const imgBuf = await sharp(inputPath).jpeg({ quality: 95 }).toBuffer();

      // autoFirstPage:false + explicit addPage ensures truly zero margin
      const doc    = new PDFDoc({ autoFirstPage: false });
      const stream = fs.createWriteStream(outputPath);
      doc.pipe(stream);
      doc.addPage({ size: [imgW, imgH], margin: 0 });

      // cover: fills the entire page box, no white gaps regardless of rounding
      doc.image(imgBuf, 0, 0, { cover: [imgW, imgH] });

      doc.end();
      stream.on('finish', resolve);
      stream.on('error', reject);
    } catch (err) { reject(err); }
  });
}

// Embeds an image into a real DOCX Word file
async function imageToDocx(inputPath, outputPath) {
  const meta   = await sharp(inputPath).metadata();
  const imgW   = meta.width  || 800;
  const imgH   = meta.height || 600;
  // Convert to PNG for docx embedding
  const pngBuf = await sharp(inputPath).png().toBuffer();

  // EMU = English Metric Units: 1 pixel = 9525 EMU (at 96 DPI)
  const emuPerPx = 9525;
  const pageWemu = imgW * emuPerPx;
  const pageHemu = imgH * emuPerPx;

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: pageWemu, height: pageHemu },
          margin: { top: 0, right: 0, bottom: 0, left: 0 }
        }
      },
      children: [
        new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [
            new ImageRun({ data: pngBuf, transformation: { width: imgW, height: imgH }, type: 'png' })
          ]
        })
      ]
    }]
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
}

// Converts plain text into a real DOCX Word file
async function textToDocx(text, outputPath) {
  const paragraphs = text.split('\n').map(line =>
    new Paragraph({ children: [new TextRun(line)] })
  );
  const doc    = new Document({ sections: [{ children: paragraphs }] });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
}

// Creates a PDF from plain text (used by DOCX→PDF and TXT→PDF)
function createPDF(text, outputPath) {
  return new Promise((resolve, reject) => {
    const doc    = new PDFDoc({ margin: 20 });
    const stream = fs.createWriteStream(outputPath);
    doc.pipe(stream);
    doc.font('Helvetica').fontSize(12).text(text, { lineGap: 5 });
    doc.end();
    stream.on('finish', resolve);
    stream.on('error', reject);
  });
}

// Creates a PDF from spreadsheet rows as a table (used by XLSX→PDF and CSV→PDF)
function spreadsheetToPDF(rows, outputPath) {
  return new Promise((resolve, reject) => {
    const doc    = new PDFDoc({ margin: 40, size: 'A4' });
    const stream = fs.createWriteStream(outputPath);
    doc.pipe(stream);

    const pageWidth  = doc.page.width - 80;  // usable width after margins
    const colCount   = rows[0] ? rows[0].length : 1;
    const colWidth   = Math.min(pageWidth / colCount, 160);
    const rowHeight  = 20;
    const fontSize   = 9;

    rows.forEach((row, rowIdx) => {
      // Start a new page if we're near the bottom
      if (doc.y + rowHeight > doc.page.height - 60) doc.addPage();

      const y = doc.y;
      const isHeader = rowIdx === 0;

      // Draw row background for header
      if (isHeader) {
        doc.rect(40, y, pageWidth, rowHeight).fill('#1a1a18');
        doc.fillColor('#ffffff');
      } else {
        // Alternating row background
        if (rowIdx % 2 === 0) doc.rect(40, y, pageWidth, rowHeight).fill('#f5f4f0');
        doc.fillColor('#1a1a18');
      }

      // Draw each cell
      row.forEach((cell, colIdx) => {
        const x = 40 + colIdx * colWidth;
        doc.font(isHeader ? 'Helvetica-Bold' : 'Helvetica')
           .fontSize(fontSize)
           .text(String(cell ?? ''), x + 4, y + 5, {
             width: colWidth - 8,
             height: rowHeight,
             ellipsis: true,
             lineBreak: false
           });
      });

      // Move down for next row
      doc.moveDown(0);
      doc.y = y + rowHeight;
      doc.fillColor('#1a1a18');
    });

    doc.end();
    stream.on('finish', resolve);
    stream.on('error', reject);
  });
}


// ==============================================
// TELEGRAM BOT
// ==============================================

const TelegramBot = require('node-telegram-bot-api');
const axios       = require('axios');
const archiver    = require('archiver');
const Tesseract   = require('tesseract.js');

const BASE_URL = process.env.RENDER_EXTERNAL_URL || 'https://shareeasy-uvia.onrender.com';

if (process.env.TELEGRAM_TOKEN) {
  const WEBHOOK_URL = `${BASE_URL}/bot${process.env.TELEGRAM_TOKEN}`;
  const tgBot = new TelegramBot(process.env.TELEGRAM_TOKEN, { webHook: true });
  tgBot.setWebHook(WEBHOOK_URL);

  const userState  = {};  // chatId → { files[], step, timer }
  const userHistory = {};  // userId → [{ fileId, fileName, date }]

  // ── 4-per-row format buttons ──
  const FORMAT_BUTTONS = {
    reply_markup: {
      inline_keyboard: [
        [
          { text: '📄 PDF',   callback_data: 'fmt_PDF'  },
          { text: '📝 DOCX',  callback_data: 'fmt_DOCX' },
          { text: '📃 Text',  callback_data: 'fmt_TXT'  },
          { text: '🖼 PNG',   callback_data: 'fmt_PNG'  }
        ],
        [
          { text: '🖼 JPG',   callback_data: 'fmt_JPG'  },
          { text: '🌐 WebP',  callback_data: 'fmt_WEBP' },
          { text: '📊 Excel', callback_data: 'fmt_XLSX' },
          { text: '📊 CSV',   callback_data: 'fmt_CSV'  }
        ]
      ]
    }
  };

  const MULTI_BUTTONS = {
    reply_markup: {
      inline_keyboard: [
        [
          { text: '📦 ZIP all files',       callback_data: 'multi_ZIP'  },
          { text: '📄 Convert all to PDF',  callback_data: 'multi_PDF'  }
        ],
        [
          { text: '📝 Convert all to Word', callback_data: 'multi_DOCX' },
          { text: '📃 Convert all to Text', callback_data: 'multi_TXT'  }
        ]
      ]
    }
  };

  // ── /start ──
  tgBot.onText(/\/start/, (msg) => {
    tgBot.sendMessage(msg.chat.id,
      `⚡ *Welcome to ShareEasy — AI File Converter!*\n\n` +
      `Convert files instantly and share them with anyone — no app needed.\n\n` +
      `*What I can do:*\n` +
      `📎 Convert files: PDF • Word • TXT • PNG • JPG • WebP • Excel • CSV\n` +
      `📦 ZIP multiple files together\n` +
      `🔍 Extract text from images (OCR)\n` +
      `🔗 Get a shareable link after every conversion\n\n` +
      `*Just type naturally:*\n` +
      `_"convert to PDF"_, _"zip all files"_, _"extract text"_\n\n` +
      `Send a file to get started! 🚀`,
      { parse_mode: 'Markdown' }
    );
  });

  // ── /help ──
  tgBot.onText(/\/help/, (msg) => {
    tgBot.sendMessage(msg.chat.id,
      `*ShareEasy Help*\n\n` +
      `*Convert a file:* Send any file → pick a format → get it back + share link\n\n` +
      `*Multiple files:* Send several at once → ZIP or convert all to same format\n\n` +
      `*OCR (image → text):* Send a photo/image → tap 🔍 Extract Text\n\n` +
      `*Share link:* After every conversion you get a 24h link — paste it anywhere\n\n` +
      `*Natural language works:*\n` +
      `• _"convert to PDF"_\n` +
      `• _"make it a Word doc"_\n` +
      `• _"zip all files"_\n` +
      `• _"extract text"_`,
      { parse_mode: 'Markdown' }
    );
  });

  // ── Collect incoming files with 3s window ──
  async function handleTelegramFile(msg, fileId, originalName) {
    const chatId = msg.chat.id;
    const caption = (msg.caption || '').toLowerCase();
    const fmtAlias = { doc:'DOCX', docx:'DOCX', jpeg:'JPG', excel:'XLSX', word:'DOCX' };
    const fmtKeys  = ['pdf','docx','doc','txt','png','jpg','jpeg','webp','xlsx','csv','excel','word'];
    const capFmt   = fmtKeys.find(f => caption.includes(f));

    if (!userState[chatId] || !['collecting'].includes(userState[chatId].step)) {
      userState[chatId] = { files: [], step: 'collecting', timer: null };
    }

    userState[chatId].files.push({ fileId, originalName });
    clearTimeout(userState[chatId].timer);

    userState[chatId].timer = setTimeout(async () => {
      const state = userState[chatId];
      if (!state || state.step !== 'collecting') return;
      const count = state.files.length;

      if (count === 1) {
        userState[chatId].step = 'awaiting_format';
        if (capFmt) {
          const fmt = fmtAlias[capFmt] || capFmt.toUpperCase();
          await convertAndSendTelegram(chatId, fmt);
        } else {
          const f = state.files[0];
          const imgExts = ['jpg','jpeg','png','webp','gif','tiff','bmp'];
          const fExt = path.extname(f.originalName).replace('.','').toLowerCase();
          const keyboard = [...FORMAT_BUTTONS.reply_markup.inline_keyboard];
          if (imgExts.includes(fExt) || f.originalName === 'photo.jpg') {
            keyboard.push([{ text: '🔍 Extract Text (OCR)', callback_data: 'ocr' }]);
          }
          tgBot.sendMessage(chatId,
            `✨ Got *${f.originalName}*!\n\n` +
            `What should I convert it to? Tap a format below or just type —\n` +
            `_"convert to PDF"_, _"extract text"_, anything works!`,
            { parse_mode: 'Markdown', reply_markup: { inline_keyboard: keyboard } }
          );
        }
      } else {
        userState[chatId].step = 'awaiting_multi_action';
        tgBot.sendMessage(chatId,
          `✨ Got *${count} files*! What should I do?\n\n` +
          `You can ZIP them all together or convert everything to the same format.\n` +
          `_Or just type: "zip all", "convert all to PDF"_`,
          { parse_mode: 'Markdown', ...MULTI_BUTTONS }
        );
      }
    }, 3000);
  }

  tgBot.on('document', (msg) => handleTelegramFile(msg, msg.document.file_id, msg.document.file_name || 'file'));
  tgBot.on('photo',    (msg) => handleTelegramFile(msg, msg.photo[msg.photo.length-1].file_id, 'photo.jpg'));

  // ── Button taps ──
  tgBot.on('callback_query', async (query) => {
    const chatId = query.message.chat.id;
    await tgBot.answerCallbackQuery(query.id);
    if (query.data.startsWith('fmt_'))        await convertAndSendTelegram(chatId, query.data.replace('fmt_', ''));
    else if (query.data === 'ocr')            await ocrTelegram(chatId);
    else if (query.data === 'multi_ZIP')      await zipAndSendTelegram(chatId);
    else if (query.data.startsWith('multi_')) await convertAllAndSendTelegram(chatId, query.data.replace('multi_', ''));
  });

  // ── Text messages ──
  tgBot.on('message', async (msg) => {
    if (msg.document || msg.photo || !msg.text) return;
    if (msg.text.startsWith('/')) return;
    const chatId = msg.chat.id;
    const state  = userState[chatId];
    const txt    = msg.text.toLowerCase().trim();

    if (state && state.step === 'awaiting_format') {
      if (txt.includes('ocr') || txt.includes('extract text') || txt.includes('read text')) {
        await ocrTelegram(chatId); return;
      }
      const fmtMap = { pdf:'PDF', docx:'DOCX', doc:'DOCX', word:'DOCX', txt:'TXT', text:'TXT', png:'PNG', jpg:'JPG', jpeg:'JPG', webp:'WEBP', xlsx:'XLSX', excel:'XLSX', csv:'CSV' };
      const m = Object.keys(fmtMap).find(k => txt.includes(k));
      if (m) { await convertAndSendTelegram(chatId, fmtMap[m]); return; }
      tgBot.sendMessage(chatId, 'Pick a format or type it — like _"convert to PDF"_:', { parse_mode: 'Markdown', ...FORMAT_BUTTONS });
      return;
    }

    if (state && state.step === 'awaiting_multi_action') {
      if (txt.includes('zip')) { await zipAndSendTelegram(chatId); return; }
      const fmtMap = { pdf:'PDF', docx:'DOCX', doc:'DOCX', word:'DOCX', txt:'TXT', text:'TXT', png:'PNG', jpg:'JPG', jpeg:'JPG', webp:'WEBP', xlsx:'XLSX', excel:'XLSX', csv:'CSV' };
      const m = Object.keys(fmtMap).find(k => txt.includes(k));
      if (m) { await convertAllAndSendTelegram(chatId, fmtMap[m]); return; }
      tgBot.sendMessage(chatId, 'Choose what to do with your files:', MULTI_BUTTONS);
      return;
    }

    tgBot.sendMessage(chatId, '📎 Please send a file first!');
  });

  // ── Convert single file ──
  async function convertAndSendTelegram(chatId, targetFormat) {
    const state = userState[chatId];
    if (!state || !state.files || !state.files[0]) { tgBot.sendMessage(chatId, 'Please send a file first!'); return; }
    const file = state.files[0];
    tgBot.sendMessage(chatId, `⏳ Converting to ${targetFormat}...`);
    try {
      const fileInfo  = await tgBot.getFile(file.fileId);
      const fileUrl   = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${fileInfo.file_path}`;
      const response  = await axios.get(fileUrl, { responseType: 'arraybuffer' });
      const fromExt   = path.extname(file.originalName).replace('.', '').toLowerCase() || 'jpg';
      const inputPath = path.join('uploads', uuidv4() + '.' + fromExt);
      fs.writeFileSync(inputPath, response.data);
      const outExt    = targetFormat.toLowerCase();
      const shareId   = uuidv4();
      const outPath   = path.join('converted', shareId + '.' + outExt);
      const baseName  = path.basename(file.originalName, path.extname(file.originalName));
      await convertFile(inputPath, outPath, fromExt, outExt);
      fs.unlink(inputPath, () => {});
      const outFileName = baseName + '.' + outExt;
      sharedFiles.unshift({
        shareId, displayName: outFileName,
        convertedFrom: fromExt.toUpperCase(), convertedTo: targetFormat.toUpperCase(),
        filePath: outPath, fileSize: fs.statSync(outPath).size,
        createdAt: new Date().toISOString(), status: 'Ready', recipientName: 'Telegram'
      });
      const shareLink = `${BASE_URL}/share/${shareId}`;
      const sentMsg = await tgBot.sendDocument(chatId, outPath, {}, { filename: outFileName });
      if (sentMsg && sentMsg.document) {
        if (!userHistory[chatId]) userHistory[chatId] = [];
        userHistory[chatId].unshift({ fileId: sentMsg.document.file_id, fileName: outFileName, date: Date.now() });
        if (userHistory[chatId].length > 10) userHistory[chatId].pop();
      }
      tgBot.sendMessage(chatId,
        `✅ *Converted to ${targetFormat}!*\n\n🔗 *Share link (24h):*\n${shareLink}\n\n💡 _In any Telegram chat, type_ @ShareEasyFileConverterBot _to share this file instantly!_`,
        { parse_mode: 'Markdown', disable_web_page_preview: true }
      );
      delete userState[chatId];
    } catch (err) {
      console.error('Telegram conversion error:', err.message);
      tgBot.sendMessage(chatId, `❌ Conversion failed: ${err.message}`);
      delete userState[chatId];
    }
  }

  // ── ZIP multiple files ──
  async function zipAndSendTelegram(chatId) {
    const state = userState[chatId];
    if (!state || !state.files || state.files.length === 0) return;
    tgBot.sendMessage(chatId, `⏳ Zipping ${state.files.length} files...`);
    const zipPath = path.join('converted', uuidv4() + '.zip');
    try {
      const downloaded = [];
      for (const file of state.files) {
        const info = await tgBot.getFile(file.fileId);
        const url  = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${info.file_path}`;
        const res  = await axios.get(url, { responseType: 'arraybuffer' });
        const ext  = path.extname(file.originalName) || '.bin';
        const p    = path.join('uploads', uuidv4() + ext);
        fs.writeFileSync(p, res.data);
        downloaded.push({ p, name: file.originalName });
      }
      await new Promise((resolve, reject) => {
        const output  = fs.createWriteStream(zipPath);
        const archive = archiver('zip', { zlib: { level: 9 } });
        archive.pipe(output);
        downloaded.forEach(({ p, name }) => archive.file(p, { name }));
        archive.finalize();
        output.on('close', resolve);
        archive.on('error', reject);
      });
      downloaded.forEach(({ p }) => fs.unlink(p, () => {}));
      const zipShareId = uuidv4();
      sharedFiles.unshift({
        shareId: zipShareId, displayName: 'shareeasy-files.zip',
        convertedFrom: 'MULTI', convertedTo: 'ZIP',
        filePath: zipPath, fileSize: fs.statSync(zipPath).size,
        createdAt: new Date().toISOString(), status: 'Ready', recipientName: 'Telegram'
      });
      const zipShareLink = `${BASE_URL}/share/${zipShareId}`;
      const zipMsg = await tgBot.sendDocument(chatId, zipPath, {}, { filename: 'shareeasy-files.zip' });
      if (zipMsg && zipMsg.document) {
        if (!userHistory[chatId]) userHistory[chatId] = [];
        userHistory[chatId].unshift({ fileId: zipMsg.document.file_id, fileName: 'shareeasy-files.zip', date: Date.now() });
        if (userHistory[chatId].length > 10) userHistory[chatId].pop();
      }
      tgBot.sendMessage(chatId,
        `✅ *Zipped ${state.files.length} files!*\n\n🔗 *Share link (24h):*\n${zipShareLink}\n\n💡 _In any Telegram chat, type_ @ShareEasyFileConverterBot _to share this file instantly!_`,
        { parse_mode: 'Markdown', disable_web_page_preview: true }
      );
      delete userState[chatId];
    } catch (err) {
      console.error('ZIP error:', err.message);
      tgBot.sendMessage(chatId, `❌ ZIP failed: ${err.message}`);
      delete userState[chatId];
    }
  }

  // ── Convert all files to same format ──
  async function convertAllAndSendTelegram(chatId, targetFormat) {
    const state = userState[chatId];
    if (!state || !state.files || state.files.length === 0) return;
    tgBot.sendMessage(chatId, `⏳ Converting ${state.files.length} files to ${targetFormat}...`);
    try {
      for (const file of state.files) {
        const info      = await tgBot.getFile(file.fileId);
        const url       = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${info.file_path}`;
        const res       = await axios.get(url, { responseType: 'arraybuffer' });
        const fromExt   = path.extname(file.originalName).replace('.', '').toLowerCase() || 'jpg';
        const inputPath = path.join('uploads', uuidv4() + '.' + fromExt);
        fs.writeFileSync(inputPath, res.data);
        const outPath   = path.join('converted', uuidv4() + '.' + targetFormat.toLowerCase());
        const baseName  = path.basename(file.originalName, path.extname(file.originalName));
        await convertFile(inputPath, outPath, fromExt, targetFormat.toLowerCase());
        fs.unlink(inputPath, () => {});
        await tgBot.sendDocument(chatId, outPath, {}, { filename: baseName + '.' + targetFormat.toLowerCase() });
      }
      tgBot.sendMessage(chatId, `✅ All *${state.files.length} files* converted to *${targetFormat}*!\n\nSend more files anytime.`, { parse_mode: 'Markdown' });
      delete userState[chatId];
    } catch (err) {
      console.error('Convert all error:', err.message);
      tgBot.sendMessage(chatId, `❌ Conversion failed: ${err.message}`);
      delete userState[chatId];
    }
  }

  // ── OCR: extract text from image ──
  async function ocrTelegram(chatId) {
    const state = userState[chatId];
    if (!state || !state.files || !state.files[0]) { tgBot.sendMessage(chatId, 'Please send an image first!'); return; }
    const file = state.files[0];
    tgBot.sendMessage(chatId, '🔍 Extracting text from image...');
    try {
      const info    = await tgBot.getFile(file.fileId);
      const url     = `https://api.telegram.org/file/bot${process.env.TELEGRAM_TOKEN}/${info.file_path}`;
      const res     = await axios.get(url, { responseType: 'arraybuffer' });
      const ext     = path.extname(file.originalName).replace('.','').toLowerCase() || 'jpg';
      const imgPath = path.join('uploads', uuidv4() + '.' + ext);
      fs.writeFileSync(imgPath, res.data);
      const { data: { text } } = await Tesseract.recognize(imgPath, 'eng');
      fs.unlink(imgPath, () => {});
      const cleaned = text.trim();
      if (!cleaned) {
        tgBot.sendMessage(chatId, '❌ No text found in this image. Make sure it has clear, readable text.');
        delete userState[chatId];
        return;
      }
      const preview = cleaned.length <= 4000 ? cleaned : cleaned.slice(0, 4000) + '...';
      tgBot.sendMessage(chatId, `📄 *Extracted Text:*\n\n${preview}`, { parse_mode: 'Markdown' });
      // Save full text as .txt file + share link
      const shareId = uuidv4();
      const txtPath = path.join('converted', shareId + '.txt');
      const txtName = path.basename(file.originalName, path.extname(file.originalName)) + '-ocr.txt';
      fs.writeFileSync(txtPath, cleaned, 'utf8');
      sharedFiles.unshift({
        shareId, displayName: txtName,
        convertedFrom: ext.toUpperCase(), convertedTo: 'TXT',
        filePath: txtPath, fileSize: fs.statSync(txtPath).size,
        createdAt: new Date().toISOString(), status: 'Ready', recipientName: 'Telegram'
      });
      const shareLink = `${BASE_URL}/share/${shareId}`;
      const ocrMsg = await tgBot.sendDocument(chatId, txtPath, {}, { filename: txtName });
      if (ocrMsg && ocrMsg.document) {
        if (!userHistory[chatId]) userHistory[chatId] = [];
        userHistory[chatId].unshift({ fileId: ocrMsg.document.file_id, fileName: txtName, date: Date.now() });
        if (userHistory[chatId].length > 10) userHistory[chatId].pop();
      }
      tgBot.sendMessage(chatId,
        `🔗 *Share link (24h):*\n${shareLink}\n\n💡 _In any Telegram chat, type_ @ShareEasyFileConverterBot _to share this file!_`,
        { parse_mode: 'Markdown', disable_web_page_preview: true }
      );
      delete userState[chatId];
    } catch (err) {
      console.error('OCR error:', err.message);
      tgBot.sendMessage(chatId, `❌ OCR failed: ${err.message}`);
      delete userState[chatId];
    }
  }

  // ── Inline mode: share converted files from any Telegram chat ──
  tgBot.on('inline_query', async (query) => {
    const userId  = query.from.id;
    const search  = query.query.toLowerCase().trim();
    const history = userHistory[userId] || [];
    const filtered = search
      ? history.filter(f => f.fileName.toLowerCase().includes(search))
      : history;

    if (filtered.length === 0) {
      await tgBot.answerInlineQuery(query.id, [{
        type: 'article',
        id: 'empty',
        title: history.length === 0 ? '📭 No converted files yet' : `🔍 No files matching "${search}"`,
        description: history.length === 0
          ? 'Convert a file first — send any file to this bot'
          : 'Try a different search term',
        input_message_content: { message_text: '👉 Use @ShareEasyFileConverterBot to convert and share files!' }
      }], { cache_time: 0 });
      return;
    }

    const results = filtered.map((item, idx) => ({
      type: 'document',
      id: String(idx),
      title: item.fileName,
      document_file_id: item.fileId,
      description: `Converted ${new Date(item.date).toLocaleString()}`
    }));

    await tgBot.answerInlineQuery(query.id, results, { cache_time: 0 });
  });

  // ── Webhook endpoint ──
  app.post(`/bot${process.env.TELEGRAM_TOKEN}`, (req, res) => {
    tgBot.processUpdate(req.body);
    res.sendStatus(200);
  });

  console.log('🤖 Telegram bot active (webhook mode)');
}

// ==============================================
// DISCORD BOT
// ==============================================
if (process.env.DISCORD_TOKEN && process.env.DISCORD_CLIENT_ID) {
  const { Client, GatewayIntentBits, REST, Routes, SlashCommandBuilder, AttachmentBuilder } = require('discord.js');

  const discordClient = new Client({ intents: [GatewayIntentBits.Guilds] });

  // ── Register slash commands on startup ──
  const commands = [
    new SlashCommandBuilder()
      .setName('convert')
      .setDescription('Convert a file to a different format')
      .addAttachmentOption(opt =>
        opt.setName('file').setDescription('The file to convert').setRequired(true))
      .addStringOption(opt =>
        opt.setName('format')
          .setDescription('Target format')
          .setRequired(true)
          .addChoices(
            { name: 'PDF',  value: 'pdf'  },
            { name: 'DOCX', value: 'docx' },
            { name: 'TXT',  value: 'txt'  },
            { name: 'PNG',  value: 'png'  },
            { name: 'JPG',  value: 'jpg'  },
            { name: 'WebP', value: 'webp' },
            { name: 'XLSX', value: 'xlsx' },
            { name: 'CSV',  value: 'csv'  }
          ))
      .toJSON(),
    new SlashCommandBuilder()
      .setName('zip')
      .setDescription('Coming soon: ZIP multiple files')
      .toJSON()
  ];

  const rest = new REST({ version: '10' }).setToken(process.env.DISCORD_TOKEN);
  rest.put(Routes.applicationCommands(process.env.DISCORD_CLIENT_ID), { body: commands })
    .then(() => console.log('✅ Discord slash commands registered globally'))
    .catch(err => console.error('Discord command register error:', err.message));

  // ── Handle /convert command ──
  discordClient.on('interactionCreate', async interaction => {
    if (!interaction.isChatInputCommand()) return;

    if (interaction.commandName === 'convert') {
      await interaction.deferReply();

      const attachment   = interaction.options.getAttachment('file');
      const targetFormat = interaction.options.getString('format');
      const originalName = attachment.name || 'file';
      const fromExt      = path.extname(originalName).replace('.', '').toLowerCase() || 'jpg';

      // File size check (Discord free: 8MB)
      if (attachment.size > 8 * 1024 * 1024) {
        await interaction.editReply('❌ File too large. Discord supports up to 8MB.');
        return;
      }

      try {
        const response  = await axios.get(attachment.url, { responseType: 'arraybuffer' });
        const inputPath = path.join('uploads', uuidv4() + '.' + fromExt);
        fs.writeFileSync(inputPath, response.data);

        const shareId   = uuidv4();
        const outPath   = path.join('converted', shareId + '.' + targetFormat);
        const baseName  = path.basename(originalName, path.extname(originalName));
        const outName   = baseName + '.' + targetFormat;

        await convertFile(inputPath, outPath, fromExt, targetFormat);
        fs.unlink(inputPath, () => {});

        // Register share link
        sharedFiles.unshift({
          shareId, displayName: outName,
          convertedFrom: fromExt.toUpperCase(), convertedTo: targetFormat.toUpperCase(),
          filePath: outPath, fileSize: fs.statSync(outPath).size,
          createdAt: new Date().toISOString(), status: 'Ready', recipientName: 'Discord'
        });

        const shareLink  = `${BASE_URL}/share/${shareId}`;
        const discordFile = new AttachmentBuilder(outPath, { name: outName });

        await interaction.editReply({
          content: `✅ **Converted to ${targetFormat.toUpperCase()}!**\n🔗 Share link (24h): ${shareLink}`,
          files: [discordFile]
        });

      } catch (err) {
        console.error('Discord conversion error:', err.message);
        await interaction.editReply(`❌ Conversion failed: ${err.message}`);
      }
    }
  });

  discordClient.once('ready', () => console.log(`🎮 Discord bot ready: ${discordClient.user.tag}`));
  discordClient.login(process.env.DISCORD_TOKEN).catch(err => console.error('Discord login error:', err.message));
}

// ==============================================
// START THE SERVER
// ==============================================
app.listen(PORT, () => {
  console.log('\n========================================');
  console.log('  ✅  ShareEasy is running!');
  console.log('========================================');
  console.log(`  👉  Open browser → http://localhost:${PORT}`);
  console.log('  Press Ctrl + C to stop.\n');
});

