// =============================================
// SHAREEASY — FRONTEND LOGIC
// =============================================
// This file controls everything you see:
// - Tab switching
// - File upload
// - Talking to the bot
// - Sending files
// - Showing received files
// =============================================

// ── State (what we're keeping track of) ──────
let uploadedFileId   = null;   // ID of uploaded file on server
let uploadedFileName = null;   // Original file name
let selectedFormat   = null;   // Chosen target format
let pendingConfirm   = null;   // Fuzzy format suggestion waiting for yes/no
let detectedEmail    = null;   // Email detected from bot chat

// ── Grab HTML elements we'll need ────────────
const dropZone       = document.getElementById('dropZone');
const fileInput      = document.getElementById('fileInput');
const filePreview    = document.getElementById('filePreview');
const fileNameEl     = document.getElementById('fileName');
const fileSizeEl     = document.getElementById('fileSize');
const fileIconEl     = document.getElementById('fileIcon');
const removeFileBtn  = document.getElementById('removeFile');
const recipientName  = document.getElementById('recipientName');
const recipientEmail = document.getElementById('recipientEmail');
const formatGrid     = document.getElementById('formatGrid');
const botInput       = document.getElementById('botInput');
const chatLog        = document.getElementById('chatLog');
const sendBotMsgBtn  = document.getElementById('sendBotMsg');
const sendBtn        = document.getElementById('sendBtn');
const resetBtn       = document.getElementById('resetBtn');
const progressWrap   = document.getElementById('progressWrap');
const progressBar    = document.getElementById('progressBar');
const progressLabel  = document.getElementById('progressLabel');
const successBox     = document.getElementById('successBox');
const shareLinkInput = document.getElementById('shareLink');
const copyLinkBtn    = document.getElementById('copyLink');
const filesList      = document.getElementById('filesList');
const historyList    = document.getElementById('historyList');


// ==============================================
// TAB SWITCHING
// ==============================================
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    // Remove active from all tabs and panels
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    // Activate clicked tab
    tab.classList.add('active');
    document.getElementById('tab-' + tab.dataset.tab).classList.add('active');
    // Load data when switching to receive/history
    if (tab.dataset.tab === 'receive' || tab.dataset.tab === 'history') {
      loadFiles();
    }
  });
});


// ==============================================
// FILE UPLOAD — DRAG & DROP + CLICK
// ==============================================

// Click on drop zone → open file picker
dropZone.addEventListener('click', () => fileInput.click());

// Drag over effect
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('dragover');
});

// Drop a file
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) handleFileSelected(file);
});

// File picked from dialog
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) handleFileSelected(fileInput.files[0]);
});

// Remove the selected file
removeFileBtn.addEventListener('click', () => {
  resetFileState();
});

// Called when user picks a file
async function handleFileSelected(file) {
  // Show preview
  fileIconEl.textContent = getFileIcon(file.name);
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatSize(file.size);
  filePreview.classList.remove('hidden');

  // Upload to server
  const formData = new FormData();
  formData.append('file', file);

  try {
    showToast('Uploading...');
    const res  = await fetch('/api/upload', { method: 'POST', body: formData });
    const data = await res.json();

    if (data.success) {
      uploadedFileId   = data.fileId;
      uploadedFileName = data.originalName;
      showToast('✅ File uploaded!');
      addBotMessage('Got <strong>' + escapeHtml(file.name) + '</strong>! Tell me what format to convert it to and who to send it to.');
    } else {
      showToast('Upload failed. Try again.');
      resetFileState();
    }
  } catch (err) {
    showToast('Upload failed. Is the server running?');
    console.error(err);
    resetFileState();
  }
}


// ==============================================
// FORMAT SELECTION
// ==============================================
formatGrid.addEventListener('click', (e) => {
  const btn = e.target.closest('.fmt-btn');
  if (!btn) return;

  // Deselect all, select clicked
  document.querySelectorAll('.fmt-btn').forEach(b => b.classList.remove('selected'));
  btn.classList.add('selected');
  selectedFormat = btn.dataset.fmt;
});


// ==============================================
// BOT CHAT
// ==============================================
sendBotMsgBtn.addEventListener('click', sendBotMessage);
botInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter') sendBotMessage();
});

async function sendBotMessage() {
  const text = botInput.value.trim();
  if (!text) return;

  // Show user's message in chat
  addUserMessage(text);
  botInput.value = '';

  try {
    const res  = await fetch('/api/bot', {
      method : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body   : JSON.stringify({
        message              : text,
        fileName             : uploadedFileName,
        currentFormat        : selectedFormat   || null,
        currentRecipient     : recipientName.value.trim() || null,
        currentRecipientEmail: detectedEmail || recipientEmail.value.trim() || null,
        pendingConfirm       : pendingConfirm   || null
      })
    });
    const data = await res.json();

    // Show bot's reply — escape to prevent XSS from API response
    addBotMessage(escapeHtml(data.reply));

    // Track fuzzy suggestion waiting for confirmation
    pendingConfirm = data.suggestedFormat || null;

    // Auto-select format if bot detected one (exact match only)
    if (data.format) {
      selectedFormat = data.format;
      document.querySelectorAll('.fmt-btn').forEach(b => {
        b.classList.toggle('selected', b.dataset.fmt === data.format);
      });
    }

    // Auto-fill recipient if bot detected one
    if (data.recipient && !recipientName.value) {
      recipientName.value = data.recipient;
    }
    // Auto-fill email if bot detected one from chat
    if (data.recipientEmail) {
      detectedEmail = data.recipientEmail;          // store in state variable
      recipientEmail.value = data.recipientEmail;   // fill the input field
    }

    // If bot detected format and recipient (or email) and a file is uploaded,
    // offer to send immediately with a one-click button in the chat
    const hasRecipient = data.recipient || data.recipientEmail;
    if (data.format && hasRecipient && uploadedFileId) {
      const div = document.createElement('div');
      div.className = 'bot-msg';
      div.innerHTML =
        '<span class="avatar">S</span>' +
        '<span>I have everything I need. ' +
        '<button class="btn-auto-send" id="autoSendBtn">Send now ✨</button>' +
        '</span>';
      chatLog.appendChild(div);
      chatLog.scrollTop = chatLog.scrollHeight;
      document.getElementById('autoSendBtn').addEventListener('click', () => {
        div.remove();
        sendFile();
      });
    }

  } catch (err) {
    addBotMessage('Could not reach the bot. Check that the server is running.');
  }
}

function addBotMessage(html) {
  const div  = document.createElement('div');
  div.className = 'bot-msg';
  // html is trusted (comes from our own code or escaped before calling this)
  div.innerHTML = '<span class="avatar">S</span><span>' + html + '</span>';
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}

function addUserMessage(text) {
  const div  = document.createElement('div');
  div.className = 'user-msg';
  const span = document.createElement('span');
  span.textContent = text;  // safe: textContent never executes HTML
  div.appendChild(span);
  chatLog.appendChild(div);
  chatLog.scrollTop = chatLog.scrollHeight;
}


// ==============================================
// SEND FILE
// ==============================================
sendBtn.addEventListener('click', sendFile);

async function sendFile() {
  // Validate
  if (!uploadedFileId) {
    showToast('Please upload a file first.');
    return;
  }
  if (!selectedFormat) {
    showToast('Please choose a target format.');
    return;
  }

  // Show progress
  progressWrap.classList.remove('hidden');
  successBox.classList.add('hidden');
  sendBtn.disabled = true;

  // Animate progress bar
  let pct = 0;
  progressBar.style.setProperty('--pct', '0%');
  progressBar.classList.add('running');
  const timer = setInterval(() => {
    pct += Math.random() * 20;
    if (pct > 90) pct = 90;
    progressBar.style.setProperty('--pct', Math.round(pct) + '%');
    progressLabel.textContent = pct < 40 ? 'Uploading...' : pct < 75 ? 'Converting...' : 'Almost done...';
  }, 300);

  try {
    const res  = await fetch('/api/convert-and-share', {
      method : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body   : JSON.stringify({
        fileId       : uploadedFileId,
        originalName : uploadedFileName,
        targetFormat : selectedFormat,
        recipientName : recipientName.value,
        recipientEmail: detectedEmail || recipientEmail.value  // prefer bot-detected email
      })
    });
    const data = await res.json();

    clearInterval(timer);
    progressBar.style.setProperty('--pct', '100%');
    progressLabel.textContent = 'Done!';

    if (data.success) {
      // Save to this device's local history
      saveToLocalHistory({
        shareId      : data.shareId,
        displayName  : uploadedFileName.replace(/\.[^.]+$/, '') + '.' + selectedFormat.toLowerCase(),
        convertedFrom: uploadedFileName.split('.').pop().toUpperCase(),
        convertedTo  : selectedFormat,
        recipientName: recipientName.value || 'Recipient',
        recipientEmail: detectedEmail || recipientEmail.value || '',
        fileSize     : 0,
        createdAt    : new Date().toISOString(),
        status       : 'Delivered'
      });
      setTimeout(() => {
        progressWrap.classList.add('hidden');
        successBox.classList.remove('hidden');
        shareLinkInput.value = data.shareLink;

        // Show email status in bot chat
        let emailNote = '';
        if (data.emailStatus === 'sent') {
          emailNote = ' An email with the file attached has been sent to <strong>' + escapeHtml(recipientEmail.value) + '</strong>!';
        } else if (data.emailStatus === 'failed') {
          emailNote = ' (Email delivery failed — share the link manually.)';
        } else {
          emailNote = ' Copy the link below to share it.';
        }
        addBotMessage('✅ Done! File converted to <strong>' + escapeHtml(selectedFormat) + '</strong>.' + emailNote);
        sendBtn.disabled = false;
      }, 500);
    } else {
      showToast('Error: ' + (data.error || 'Something went wrong.'));
      progressWrap.classList.add('hidden');
      sendBtn.disabled = false;
    }

  } catch (err) {
    clearInterval(timer);
    showToast('Failed to send. Is the server running?');
    progressWrap.classList.add('hidden');
    sendBtn.disabled = false;
    console.error(err);
  }
}

// Copy share link to clipboard
copyLinkBtn.addEventListener('click', () => {
  shareLinkInput.select();
  navigator.clipboard.writeText(shareLinkInput.value).then(() => {
    showToast('Link copied!');
  }).catch(() => {
    document.execCommand('copy');
    showToast('Link copied!');
  });
});


// ==============================================
// RECEIVE / HISTORY — LOAD FILES
// ==============================================
document.getElementById('refreshFiles').addEventListener('click', loadFiles);
document.getElementById('refreshHistory').addEventListener('click', loadFiles);

// ── Local history stored per device in localStorage ──
function getLocalHistory() {
  try { return JSON.parse(localStorage.getItem('shareeasy_history') || '[]'); } catch { return []; }
}
function saveToLocalHistory(entry) {
  const history = getLocalHistory();
  history.unshift(entry);
  localStorage.setItem('shareeasy_history', JSON.stringify(history.slice(0, 50)));
}

async function loadFiles() {
  // History tab — show only THIS device's sent files from localStorage
  const history = getLocalHistory();
  if (history.length === 0) {
    historyList.innerHTML = '<div class="empty-state">No transfers yet.<br/>Send your first file using the Send tab!</div>';
  } else {
    historyList.innerHTML = history.map(f => buildFileCard(f)).join('');
  }

  // Receive tab — still fetches from server (all files available for download)
  try {
    const res   = await fetch('/api/files');
    const files = await res.json();
    if (files.length === 0) {
      filesList.innerHTML = '<div class="empty-state">No files yet.<br/>Ask someone to send you a file using ShareEasy.</div>';
    } else {
      filesList.innerHTML = files.map(f => buildFileCard(f)).join('');
    }
  } catch (err) {
    console.error('Could not load files:', err);
  }
}

function buildFileCard(f) {
  const icon    = getFileIcon(f.displayName);
  const size    = formatSize(f.fileSize);
  const date    = timeAgo(f.createdAt);
  // Escape all server-supplied strings before inserting into innerHTML
  const safeName      = escapeHtml(f.displayName);
  const safeFrom      = escapeHtml(f.convertedFrom);
  const safeTo        = escapeHtml(f.convertedTo);
  const safeRecipient = escapeHtml(f.recipientName);
  const safeShareId   = escapeHtml(f.shareId);
  const badge   = f.status === 'Delivered'
    ? '<span class="badge badge-green">✓ Delivered</span>'
    : '<span class="badge badge-amber">⏳ Pending</span>';

  return `
    <div class="file-card">
      <div class="fc-thumb">${icon}</div>
      <div class="fc-body">
        <div class="fc-name">${safeName}</div>
        <div class="fc-meta">
          <span>${safeFrom} → ${safeTo}</span>
          <span>${size}</span>
          <span>To: ${safeRecipient}</span>
          <span>${date}</span>
        </div>
        <div class="fc-actions">
          ${badge}
          <a class="btn-dl" href="/api/download/${safeShareId}" download>
            ⬇ Download ${safeTo}
          </a>
        </div>
      </div>
    </div>`;
}


// ==============================================
// RESET
// ==============================================
resetBtn.addEventListener('click', resetAll);

function resetAll() {
  resetFileState();
  recipientName.value  = '';
  recipientEmail.value = '';
  selectedFormat       = null;
  pendingConfirm       = null;
  detectedEmail        = null;
  document.querySelectorAll('.fmt-btn').forEach(b => b.classList.remove('selected'));
  successBox.classList.add('hidden');
  progressWrap.classList.add('hidden');
  chatLog.innerHTML = '<div class="bot-msg"><span class="avatar">B</span><span>Ready! Upload a file to get started.</span></div>';
}

function resetFileState() {
  uploadedFileId   = null;
  uploadedFileName = null;
  filePreview.classList.add('hidden');
  fileInput.value  = '';
}


// ==============================================
// HELPERS
// ==============================================

// Escape HTML to prevent XSS when inserting user/server data into the DOM
function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Get emoji icon for file type
function getFileIcon(name) {
  const ext = name.split('.').pop().toLowerCase();
  const icons = { pdf:'📄', docx:'📝', doc:'📝', txt:'🔤', png:'🖼️', jpg:'📷', jpeg:'📷', webp:'🌐', xlsx:'📊', xls:'📊', csv:'📋' };
  return icons[ext] || '📁';
}

// Format bytes → human readable
function formatSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return Math.round(bytes / 1024) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// How long ago was this?
function timeAgo(iso) {
  const diff = Math.round((Date.now() - new Date(iso).getTime()) / 1000);
  if (diff < 60)   return 'just now';
  if (diff < 3600) return Math.floor(diff / 60) + 'm ago';
  if (diff < 86400)return Math.floor(diff / 3600) + 'h ago';
  return Math.floor(diff / 86400) + 'd ago';
}

// Show a quick toast notification
let toastTimer;
function showToast(msg) {
  let toast = document.getElementById('toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'toast';
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.remove('show'), 2500);
}

// If user opens /receive directly, switch to receive tab
if (window.location.pathname === '/receive' || window.location.hash) {
  document.querySelector('[data-tab="receive"]').click();
}
