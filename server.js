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
const nodemailer = require('nodemailer');
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

// ── Email transporter (Gmail) ───────────────────
const mailer = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.EMAIL_FROM,
    pass: process.env.EMAIL_PASS
  }
});

// ── In-memory list of shared files ───────────
// (works for local use; for production use a database)
let sharedFiles = [];


// ==============================================
// SMART LOCAL BOT — keyword detection
// No internet needed. Works behind any firewall.
// Switch to Gemini when deploying online.
// ==============================================

// Levenshtein edit distance — counts how many single-character
// changes (add/remove/replace) it takes to turn string a into b.
// e.g. "pdd" vs "pdf" = 1 change → close enough to suggest
function editDistance(a, b) {
  const m = a.length, n = b.length;
  const dp = Array.from({ length: m + 1 }, (_, i) =>
    Array.from({ length: n + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0))
  );
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] = a[i-1] === b[j-1]
        ? dp[i-1][j-1]
        : 1 + Math.min(dp[i-1][j], dp[i][j-1], dp[i-1][j-1]);
  return dp[m][n];
}

// Format definitions — only the REAL canonical keywords here.
// Fuzzy matching handles all typos automatically.
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

// Try to match a single word to a format.
// Returns { format, exact } where exact=false means it was a fuzzy guess.
function matchFormat(word) {
  // 1. Exact match first
  for (const entry of FORMAT_MAP) {
    if (entry.keys.includes(word)) return { format: entry.format, exact: true };
  }
  // 2. Fuzzy match — allow 1 edit for short words (<=4 chars), 2 for longer
  let best = null, bestDist = Infinity;
  for (const entry of FORMAT_MAP) {
    for (const key of entry.keys) {
      const dist = editDistance(word, key);
      const maxAllowed = key.length <= 4 ? 1 : 2;
      if (dist <= maxAllowed && dist < bestDist) {
        bestDist = dist;
        best = entry.format;
      }
    }
  }
  if (best) return { format: best, exact: false };
  return null;
}

function localBot(message, currentFormat = null, currentRecipient = null, pendingConfirm = null, currentRecipientEmail = null) {
  const msg   = message.toLowerCase().trim();
  const words = msg.split(/\s+/);

  // ── Handle yes/no confirmation for fuzzy suggestions ──
  if (pendingConfirm) {
    if (/^(yes|yeah|yep|y|correct|right|sure|ok|okay)$/i.test(msg)) {
      const format         = pendingConfirm;
      const recipient      = currentRecipient;
      const recipientEmail = currentRecipientEmail;
      if (format && recipient) {
        return { reply: `Great! I'll convert to ${format} and send it to ${recipientEmail || recipient}. Click "Send now" to confirm!`, format, recipient, recipientEmail, suggestedFormat: null };
      }
      if (format) {
        return { reply: `Got it, ${format} it is! Who should I send it to?`, format, recipient: null, recipientEmail: currentRecipientEmail, suggestedFormat: null };
      }
    }
    if (/^(no|nope|n|wrong|nah)$/i.test(msg)) {
      return { reply: `My bad! What format did you want? (PDF, DOCX, TXT, PNG, JPG, XLSX, CSV)`, format: null, recipient: currentRecipient, recipientEmail: currentRecipientEmail, suggestedFormat: null };
    }
  }

  // ── Detect format from all words in message ────
  let detectedFormat = null;
  let suggestedFormat = null;  // fuzzy guess needing confirmation

  for (const word of words) {
    const result = matchFormat(word);
    if (!result) continue;
    if (result.exact) { detectedFormat = result.format; break; }
    if (!suggestedFormat) suggestedFormat = result.format; // keep first fuzzy guess
  }

  // ── Detect recipient email (if typed in chat) ──
  let detectedEmail = null;
  const emailMatch = msg.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/);
  if (emailMatch) detectedEmail = emailMatch[0];

  // ── Detect recipient name ───────────────────
  let detectedRecipient = null;
  // If an email was found, use the part before @ as the display name
  if (detectedEmail) {
    detectedRecipient = detectedEmail.split('@')[0].replace(/[._\-]/g, ' ').replace(/\b\w/g, c => c.toUpperCase());
  } else {
    const recipientMatch = msg.match(
      /(?:to|for|send\s+to|share\s+with|send\s+it\s+to|give\s+to)\s+([a-z][a-z\s]{1,20}?)(?:\s+(?:as|in|the|a|an|format)|[.,!?]|$)/i
    );
    if (recipientMatch) {
      detectedRecipient = recipientMatch[1].trim().replace(/\b\w/g, c => c.toUpperCase());
    }
  }

  // ── Context-aware fill ────────────────────
  const allFormatKeys = FORMAT_MAP.flatMap(e => e.keys);
  const looksLikeName = /^[a-zA-Z]{2,30}$/.test(msg) && !allFormatKeys.includes(msg) && !matchFormat(msg);
  if (!detectedRecipient && currentFormat && !currentRecipient && looksLikeName) {
    detectedRecipient = msg.replace(/\b\w/g, c => c.toUpperCase());
  }

  // Merge with remembered context
  const format         = detectedFormat    || currentFormat;
  const recipient      = detectedRecipient || currentRecipient;
  const recipientEmail = detectedEmail     || null;

  // ── If only a fuzzy guess was found, ask for confirmation ──
  if (!detectedFormat && suggestedFormat) {
    return {
      reply: `Did you mean ${suggestedFormat}? Reply "yes" to confirm or "no" to pick a different format.`,
      format: currentFormat,
      recipient,
      recipientEmail,      // carry email through so it's not lost on "yes"
      suggestedFormat
    };
  }

  // ── Build friendly reply ───────────────────
  let reply;
  if (format && recipient) {
    const dest = recipientEmail || recipient;
    reply = `Got it! I'll convert the file to ${format} and send it to ${dest}. Click "Send now" to confirm!`;
  } else if (format) {
    reply = `Sure, I'll convert the file to ${format}. Who should I send it to?`;
  } else if (recipient) {
    reply = `I'll send it to ${recipient}. What format would you like? (PDF, DOCX, TXT, PNG, JPG, XLSX, CSV)`;
  } else {
    reply = `I didn't catch the format or recipient. Try: "Send this to hari@gmail.com as PDF".`;
  }

  return { reply, format, recipient, recipientEmail, suggestedFormat: null };
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

// ── 2. Bot chat (smart local keyword detection) ──
app.post('/api/bot', (req, res) => {
  const { message, currentFormat = null, currentRecipient = null, pendingConfirm = null, currentRecipientEmail = null } = req.body;
  if (!message || !message.trim()) {
    return res.json({ reply: 'Please type a message first!', format: null, recipient: null, recipientEmail: null, suggestedFormat: null });
  }
  const result = localBot(message, currentFormat, currentRecipient, pendingConfirm, currentRecipientEmail);
  res.json(result);
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
    if (recipientEmail && process.env.EMAIL_FROM && process.env.EMAIL_PASS) {
      try {
        await mailer.sendMail({
          from   : `"ShareEasy" <${process.env.EMAIL_FROM}>`,
          to     : recipientEmail,
          subject: `${recipientName || 'Someone'} shared a file with you via ShareEasy`,
          html   : `
            <div style="font-family:sans-serif;max-width:520px;margin:0 auto">
              <h2 style="color:#1a1a18">⚡ ShareEasy</h2>
              <p>Hi <strong>${recipientName || 'there'}</strong>,</p>
              <p><strong>${entry.recipientName}</strong> has shared a converted file with you.</p>
              <table style="background:#f5f4f0;border-radius:8px;padding:12px 16px;margin:16px 0;width:100%">
                <tr><td><strong>File:</strong></td><td>${entry.displayName}</td></tr>
                <tr><td><strong>Converted:</strong></td><td>${entry.convertedFrom} → ${entry.convertedTo}</td></tr>
              </table>
              <p>The converted file is attached to this email. Simply open it!</p>
              <hr style="border:none;border-top:1px solid #eee;margin:20px 0"/>
              <p style="color:#aaa;font-size:12px">Sent via ShareEasy &mdash; convert &amp; share files instantly</p>
            </div>`,
          attachments: [{
            filename: entry.displayName,
            path    : outPath
          }]
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
