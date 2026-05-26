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
      const meta   = await sharp(inputPath).metadata();
      const imgW   = meta.width  || 800;
      const imgH   = meta.height || 600;
      const pageW  = 595.28;  // A4 width in points
      const pageH  = 841.89;  // A4 height in points
      const margin = 40;
      const ratio  = Math.min((pageW - margin * 2) / imgW, (pageH - margin * 2) / imgH);

      const doc    = new PDFDoc({ size: 'A4', margin: 0 });
      const stream = fs.createWriteStream(outputPath);
      doc.pipe(stream);

      // PDFKit supports PNG and JPEG natively — convert anything else to PNG first
      const needsConvert = ['webp', 'gif', 'tiff', 'bmp'].includes(fromExt);
      if (needsConvert) {
        const pngBuf = await sharp(inputPath).png().toBuffer();
        doc.image(pngBuf, margin, margin, { width: imgW * ratio, height: imgH * ratio });
      } else {
        doc.image(inputPath, margin, margin, { width: imgW * ratio, height: imgH * ratio });
      }

      doc.end();
      stream.on('finish', resolve);
      stream.on('error', reject);
    } catch (err) { reject(err); }
  });
}

// Embeds an image into a real DOCX Word file
async function imageToDocx(inputPath, outputPath) {
  const pngBuf = await sharp(inputPath).png().toBuffer();
  const meta   = await sharp(inputPath).metadata();
  const maxPx  = 600;
  const scale  = (meta.width || maxPx) > maxPx ? maxPx / meta.width : 1;
  const w = Math.round((meta.width  || maxPx) * scale);
  const h = Math.round((meta.height || 400)   * scale);

  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({
          children: [
            new ImageRun({ data: pngBuf, transformation: { width: w, height: h }, type: 'png' })
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
    const doc    = new PDFDoc({ margin: 60 });
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
// START THE SERVER
// ==============================================
app.listen(PORT, () => {
  console.log('\n========================================');
  console.log('  ✅  ShareEasy is running!');
  console.log('========================================');
  console.log(`  👉  Open browser → http://localhost:${PORT}`);
  console.log('  Press Ctrl + C to stop.\n');
});
