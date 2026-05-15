// ===== UDISE+ Bot Frontend Application =====

const API_BASE = '';
let socket = null;
let uploadedData = null;
let currentSessionId = null;

// ===== SOCKET.IO CONNECTION =====
function initSocket() {
  socket = io();
  socket.on('connect', () => {
    document.getElementById('connectionStatus').textContent = 'Connected';
  });
  socket.on('disconnect', () => {
    document.getElementById('connectionStatus').textContent = 'Disconnected';
  });
  socket.on('log', (entry) => addLogEntry(entry));
  socket.on('progress', (data) => updateProgress(data));
  socket.on('captcha', () => {
    document.getElementById('captchaAlert').classList.remove('hidden');
  });
  socket.on('complete', (results) => showResults(results));
  socket.on('error', (data) => {
    addLogEntry({ message: `Error: ${data.message}`, type: 'error', timestamp: new Date().toISOString() });
    document.getElementById('statusText').textContent = 'Automation failed';
    document.getElementById('stopBot').classList.add('hidden');
    document.getElementById('restartBot').classList.remove('hidden');
  });
}

// ===== FILE UPLOAD =====
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', (e) => { if (e.target.files.length) handleFile(e.target.files[0]); });

async function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls', 'csv'].includes(ext)) {
    alert('Please upload an Excel (.xlsx, .xls) or CSV file.');
    return;
  }

  const formData = new FormData();
  formData.append('file', file);

  dropZone.innerHTML = '<div class="upload-zone-content"><h3>⏳ Uploading & parsing...</h3></div>';

  try {
    const res = await fetch(`${API_BASE}/api/upload`, { method: 'POST', body: formData });
    const data = await res.json();

    if (!res.ok) throw new Error(data.error || 'Upload failed');

    uploadedData = data;
    showFileInfo(data);
  } catch (err) {
    alert('Upload failed: ' + err.message);
    resetUploadZone();
  }
}

function showFileInfo(data) {
  dropZone.classList.add('hidden');
  const fileInfo = document.getElementById('fileInfo');
  fileInfo.classList.remove('hidden');

  document.getElementById('fileName').textContent = data.fileName;
  document.getElementById('fileRows').textContent = `${data.totalRows} student records found`;

  // Show warning if any
  if (data.warnings) {
    document.getElementById('fileWarning').classList.remove('hidden');
    document.getElementById('warningText').textContent = data.warnings;
  }

  // Build preview table
  const thead = document.getElementById('previewHead');
  const tbody = document.getElementById('previewBody');
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const fields = [
    { key: 'studentName', label: 'Student Name' },
    { key: 'class', label: 'Class' },
    { key: 'penNo', label: 'PEN No' },
    { key: 'attendance', label: 'Attendance' },
    { key: 'percentage', label: 'Percentage' },
    { key: 'progressionStatus', label: 'Status' },
    { key: 'sameSchool', label: 'Same School' }
  ];

  fields.forEach(f => { const th = document.createElement('th'); th.textContent = f.label; thead.appendChild(th); });

  data.preview.forEach(row => {
    const tr = document.createElement('tr');
    fields.forEach(f => { const td = document.createElement('td'); td.textContent = row[f.key] || '—'; tr.appendChild(td); });
    tbody.appendChild(tr);
  });

  // Show column mapping
  const mappingGrid = document.getElementById('mappingGrid');
  mappingGrid.innerHTML = '';
  const mappingSection = document.getElementById('columnMapping');

  const detected = data.preview[0] || {};
  let mappingCount = 0;
  fields.forEach(f => {
    if (detected[f.key]) {
      mappingCount++;
      const item = document.createElement('div');
      item.className = 'mapping-item';
      item.innerHTML = `<span>${f.label}</span> <span class="arrow">→</span> <strong>✓</strong>`;
      mappingGrid.appendChild(item);
    }
  });

  if (mappingCount > 0) mappingSection.classList.remove('hidden');

  document.getElementById('nextStep1').disabled = false;
}

function resetUploadZone() {
  dropZone.classList.remove('hidden');
  dropZone.innerHTML = `
    <div class="upload-zone-content">
      <div class="upload-animation">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
          <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
          <polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
        </svg>
      </div>
      <h3>Drag & Drop your Excel file here</h3>
      <p>or click to browse</p>
      <div class="upload-formats">
        <span class="format-badge">.xlsx</span>
        <span class="format-badge">.xls</span>
        <span class="format-badge">.csv</span>
      </div>
    </div>`;
  document.getElementById('fileInfo').classList.add('hidden');
  document.getElementById('fileWarning').classList.add('hidden');
  document.getElementById('columnMapping').classList.add('hidden');
  document.getElementById('nextStep1').disabled = true;
  fileInput.value = '';
  uploadedData = null;
}

document.getElementById('removeFile').addEventListener('click', resetUploadZone);

// ===== STEP NAVIGATION =====
function goToStep(num) {
  document.querySelectorAll('.step-section').forEach(s => s.classList.remove('active'));
  document.getElementById(`step${num}`).classList.add('active');

  document.querySelectorAll('.step').forEach(s => {
    const n = parseInt(s.dataset.step);
    s.classList.remove('active', 'done');
    if (n === num) s.classList.add('active');
    else if (n < num) s.classList.add('done');
  });
}

document.getElementById('nextStep1').addEventListener('click', () => goToStep(2));
document.getElementById('backStep2').addEventListener('click', () => goToStep(1));

// ===== PASSWORD TOGGLE =====
document.getElementById('togglePassword').addEventListener('click', () => {
  const pw = document.getElementById('password');
  pw.type = pw.type === 'password' ? 'text' : 'password';
});

// ===== START AUTOMATION =====
document.getElementById('nextStep2').addEventListener('click', async () => {
  const udiseCode = document.getElementById('udiseCode').value.trim();
  const password = document.getElementById('password').value;

  if (!udiseCode || !password) { alert('Please enter both UDISE Code and Password.'); return; }
  if (!uploadedData) { alert('No data uploaded. Please go back and upload a file.'); return; }

  goToStep(3);

  try {
    const res = await fetch(`${API_BASE}/api/start-automation`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        socketId: socket.id,
        udiseCode,
        password,
        data: uploadedData.data,
        filePath: uploadedData.filePath
      })
    });

    const result = await res.json();
    if (!res.ok) throw new Error(result.error);
    currentSessionId = result.sessionId;
  } catch (err) {
    addLogEntry({ message: `Failed to start: ${err.message}`, type: 'error', timestamp: new Date().toISOString() });
  }
});

// ===== STOP BOT =====
document.getElementById('stopBot').addEventListener('click', async () => {
  if (!currentSessionId) return;
  if (!confirm('Are you sure you want to stop the bot?')) return;

  try {
    await fetch(`${API_BASE}/api/stop-automation`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ sessionId: currentSessionId })
    });
    addLogEntry({ message: 'Stop signal sent...', type: 'warn', timestamp: new Date().toISOString() });
  } catch (err) { /* ignore */ }
});

// ===== RESTART =====
document.getElementById('restartBot').addEventListener('click', () => {
  currentSessionId = null;
  uploadedData = null;
  document.getElementById('logConsole').innerHTML = '';
  document.getElementById('progressBar').style.width = '0%';
  document.getElementById('progressPercent').textContent = '0%';
  document.getElementById('progressCount').textContent = '0 / 0 students';
  document.getElementById('currentStudent').classList.add('hidden');
  document.getElementById('captchaAlert').classList.add('hidden');
  document.getElementById('resultsSummary').classList.add('hidden');
  document.getElementById('statusText').textContent = 'Initializing the bot...';
  document.getElementById('stopBot').classList.remove('hidden');
  document.getElementById('restartBot').classList.add('hidden');
  resetUploadZone();
  goToStep(1);
});

// ===== CLEAR LOGS =====
document.getElementById('clearLogs').addEventListener('click', () => {
  document.getElementById('logConsole').innerHTML = '';
});

// ===== PROGRESS UPDATES =====
function updateProgress(data) {
  const bar = document.getElementById('progressBar');
  const pct = document.getElementById('progressPercent');
  const cnt = document.getElementById('progressCount');
  const status = document.getElementById('statusText');
  const curr = document.getElementById('currentStudent');

  bar.style.width = `${data.percent}%`;
  pct.textContent = `${data.percent}%`;

  const statusMap = {
    navigating: 'Navigating to UDISE+ portal...',
    logging_in: 'Logging into UDISE+...',
    waiting_captcha: 'Waiting for CAPTCHA to be solved...',
    logged_in: 'Successfully logged in!',
    navigating_student: 'Navigating to student section...',
    ready: 'Ready to process students',
    processing: `Processing student ${data.current || 0} of ${data.total || 0}`,
    completed: 'Automation completed!'
  };
  status.textContent = statusMap[data.step] || data.step;

  if (data.current && data.total) {
    cnt.textContent = `${data.current} / ${data.total} students`;
  }

  if (data.studentName) {
    curr.classList.remove('hidden');
    document.getElementById('currentStudentName').textContent = data.studentName;
  }

  if (data.step === 'completed') {
    curr.classList.add('hidden');
    document.getElementById('stopBot').classList.add('hidden');
    document.getElementById('restartBot').classList.remove('hidden');
  }
}

// ===== LOG ENTRIES =====
function addLogEntry(entry) {
  const console = document.getElementById('logConsole');
  const div = document.createElement('div');
  div.className = `log-entry ${entry.type || 'info'}`;

  const time = entry.timestamp ? new Date(entry.timestamp).toLocaleTimeString() : '--:--:--';
  div.innerHTML = `<span class="log-time">${time}</span><span class="log-msg">${escapeHtml(entry.message)}</span>`;

  console.appendChild(div);
  console.scrollTop = console.scrollHeight;
}

// ===== RESULTS =====
function showResults(results) {
  const summary = document.getElementById('resultsSummary');
  summary.classList.remove('hidden');
  document.getElementById('successCount').textContent = results.success;
  document.getElementById('failedCount').textContent = results.failed;
  document.getElementById('totalCount').textContent = results.total;

  document.getElementById('stopBot').classList.add('hidden');
  document.getElementById('restartBot').classList.remove('hidden');
}

// ===== UTILS =====
function escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

// ===== INIT =====
initSocket();
