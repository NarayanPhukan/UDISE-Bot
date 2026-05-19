// ===== UDISE+ Bot Frontend =====
const API = '';
let socket, uploadedData, sessionId;
let confirmCountdownTimer = null;
let confirmCountdownSeconds = 0;

// ---- Socket ----
function initSocket() {
  socket = io();
  socket.on('connect', () => { document.getElementById('connectionStatus').textContent = 'Connected'; });
  socket.on('disconnect', () => { document.getElementById('connectionStatus').textContent = 'Disconnected'; });
  socket.on('log', addLog);
  socket.on('progress', updateProgress);
  socket.on('captcha', (data) => {
    const modal = document.getElementById('captchaModal');
    modal.classList.remove('hidden');
    if (data.image) document.getElementById('captchaImg').src = data.image;
    document.getElementById('captchaInput').value = '';
    document.getElementById('captchaInput').focus();
  });
  socket.on('confirm-student', (data) => {
    const modal = document.getElementById('confirmModal');
    modal.classList.remove('hidden');
    
    const warningBanner = document.getElementById('alreadyEnteredWarning');
    const submitBtn = document.getElementById('confirmSubmitBtn');
    
    if (data.alreadyEntered) {
      if (warningBanner) warningBanner.classList.remove('hidden');
      if (submitBtn) {
        submitBtn.textContent = 'UPDATE RECORD';
        submitBtn.className = 'btn btn-warn';
      }
    } else {
      if (warningBanner) warningBanner.classList.add('hidden');
      if (submitBtn) {
        submitBtn.textContent = 'CONFIRM & SUBMIT';
        submitBtn.className = 'btn btn-accent';
      }
    }
    
    document.getElementById('confName').textContent = data.studentName || '—';
    document.getElementById('confClass').textContent = data.class || '—';
    document.getElementById('confPen').textContent = data.penNo || '—';
    
    document.getElementById('confMarks').value = data.marks || '';
    document.getElementById('confPercent').value = data.percentage || '';
    document.getElementById('confDays').value = data.attendance || '155';
    
    const progSelect = document.getElementById('confProgression');
    if (data.progressionStatus) {
      const status = String(data.progressionStatus).toLowerCase();
      if (status.includes('fail') || status.includes('not')) {
        progSelect.value = 'Not Promoted';
      } else if (status.includes('without')) {
        progSelect.value = 'Promoted Without Examination';
      } else if (status.includes('discontinued') || status.includes('before')) {
        progSelect.value = 'Discontinued Before Examination';
      } else {
        progSelect.value = 'Promoted';
      }
    } else {
      progSelect.value = 'Promoted';
    }
    
    const schoolSelect = document.getElementById('confSameSchool');
    if (data.sameSchool) {
      const school = String(data.sameSchool).toLowerCase();
      if (school.includes('left') || school.includes('tc')) {
        schoolSelect.value = 'Left School with TC / Without TC';
      } else {
        schoolSelect.value = 'Studying in Same School';
      }
    } else {
      schoolSelect.value = 'Studying in Same School';
    }
    
    const imgContainer = document.getElementById('confirmScreenshotContainer');
    if (data.screenshot) {
      imgContainer.classList.remove('hidden');
      document.getElementById('confirmImg').src = data.screenshot;
    } else {
      imgContainer.classList.add('hidden');
    }

    // Start 10-second countdown timer for auto-confirm
    startConfirmCountdown(10);

    // Auto-focus and highlight the attendance days input field for seamless keyboard operation
    setTimeout(() => {
      const confDays = document.getElementById('confDays');
      if (confDays) {
        confDays.focus();
        confDays.select();
      }
    }, 100);
  });
  socket.on('auto-confirm-close', () => {
    clearConfirmCountdown();
    document.getElementById('confirmModal').classList.add('hidden');
  });
  socket.on('error-flash', (data) => {
    showErrorFlash(data.studentName, data.message);
  });
  socket.on('complete', showResults);
  socket.on('error', (d) => {
    addLog({ message: 'Error: ' + d.message, type: 'error', timestamp: new Date().toISOString() });
    document.getElementById('statusText').textContent = 'Failed';
    document.getElementById('btnStop').classList.add('hidden');
    document.getElementById('btnRestart').classList.remove('hidden');
  });
}

// ---- File Upload ----
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const browseBtn = document.getElementById('browseBtn');

dropZone.addEventListener('click', (e) => { if (e.target.closest('.btn-browse') || e.target === dropZone || e.target.closest('.upload-content')) fileInput.click(); });
browseBtn?.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', (e) => { e.preventDefault(); dropZone.classList.remove('dragover'); if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]); });
fileInput.addEventListener('change', (e) => { if (e.target.files.length) handleFile(e.target.files[0]); });

async function handleFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx','xls','csv'].includes(ext)) return alert('Please upload .xlsx, .xls or .csv');
  const fd = new FormData(); fd.append('file', file);
  dropZone.querySelector('.upload-content').innerHTML = '<h3>⏳ Uploading & parsing...</h3>';
  try {
    const res = await fetch(API + '/api/upload', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error);
    uploadedData = data;
    showFileData(data);
  } catch (err) { alert('Upload failed: ' + err.message); resetUpload(); }
}

function showFileData(data) {
  dropZone.classList.add('hidden');
  const fl = document.getElementById('fileLoaded'); fl.classList.remove('hidden');
  document.getElementById('fileName').textContent = data.fileName;
  document.getElementById('fileRowCount').textContent = data.totalRows + ' records found';
  if (data.warnings) { document.getElementById('fileWarning').classList.remove('hidden'); document.getElementById('warningText').textContent = data.warnings; }

  const fields = [
    { key:'studentName', label:'Student Name' }, { key:'class', label:'Class' },
    { key:'penNo', label:'PEN No' }, { key:'attendance', label:'Attendance' },
    { key:'percentage', label:'Percentage' }, { key:'progressionStatus', label:'Status' },
    { key:'sameSchool', label:'Same School' }
  ];
  const thead = document.getElementById('previewHead');
  const tbody = document.getElementById('previewBody');
  thead.innerHTML = ''; tbody.innerHTML = '';
  fields.forEach(f => { const th = document.createElement('th'); th.textContent = f.label; thead.appendChild(th); });
  data.preview.forEach(row => {
    const tr = document.createElement('tr');
    fields.forEach(f => { const td = document.createElement('td'); td.textContent = row[f.key] || '—'; tr.appendChild(td); });
    tbody.appendChild(tr);
  });

  // Mapping chips
  const chips = document.getElementById('mappingChips'); chips.innerHTML = '';
  const ms = document.getElementById('mappingSection');
  const det = data.preview[0] || {};
  let cnt = 0;
  fields.forEach(f => {
    if (det[f.key]) { cnt++;
      const c = document.createElement('div'); c.className = 'chip';
      c.innerHTML = `<span>${f.label}</span><span class="arr">→</span><span class="chk">✓</span>`;
      chips.appendChild(c);
    }
  });
  if (cnt > 0) ms.classList.remove('hidden');
  document.getElementById('btnGoStep2').disabled = false;
}

function resetUpload() {
  dropZone.classList.remove('hidden');
  dropZone.querySelector('.upload-content').innerHTML = `
    <div class="upload-icon-wrap"><div class="upload-icon-glow"></div>
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
        <path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"/><path d="M12 12v9"/><path d="m16 16-4-4-4 4"/>
      </svg></div>
    <h3>Drag & Drop UDISE+ CSV or Excel Files</h3><p>or browse your files</p>
    <button class="btn btn-outline btn-browse" id="browseBtn">BROWSE FILES</button>
    <div class="upload-formats"><span>.xlsx</span><span>.xls</span><span>.csv</span></div>`;
  document.getElementById('fileLoaded').classList.add('hidden');
  document.getElementById('fileWarning').classList.add('hidden');
  document.getElementById('mappingSection')?.classList.add('hidden');
  document.getElementById('btnGoStep2').disabled = true;
  fileInput.value = ''; uploadedData = null;
  // Re-attach browse btn
  document.getElementById('browseBtn')?.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
}

document.getElementById('removeFileBtn').addEventListener('click', resetUpload);

// ---- Step Nav ----
function goStep(n) {
  document.querySelectorAll('.step-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('panel' + n).classList.add('active');
  document.querySelectorAll('.wizard-step').forEach(s => {
    const sn = +s.dataset.step;
    s.classList.remove('active','done');
    if (sn === n) s.classList.add('active');
    else if (sn < n) s.classList.add('done');
  });
  // Animate wizard lines
  document.querySelectorAll('.wizard-line-fill').forEach((f, i) => {
    f.style.width = (i + 1 < n) ? '100%' : '0';
  });
  // Hide guide on step 2/3
  const g = document.getElementById('guideCard');
  if (g) g.style.display = n === 1 ? '' : 'none';
}

document.getElementById('btnGoStep2').addEventListener('click', () => goStep(2));
document.getElementById('btnBackStep1').addEventListener('click', () => goStep(1));

// ---- Password Toggle ----
document.getElementById('togglePw').addEventListener('click', () => {
  const pw = document.getElementById('password');
  pw.type = pw.type === 'password' ? 'text' : 'password';
});

// ---- Section Selection Filter Toggles ----
const btnChooseSection = document.getElementById('btnChooseSection');
const sectionFilterContainer = document.getElementById('sectionFilterContainer');
const btnApplySection = document.getElementById('btnApplySection');
const targetSectionInput = document.getElementById('targetSectionInput');
const selectedSectionBadge = document.getElementById('selectedSectionBadge');

if (btnChooseSection && sectionFilterContainer) {
  btnChooseSection.addEventListener('click', () => {
    const isHidden = sectionFilterContainer.style.display === 'none' || sectionFilterContainer.style.display === '';
    sectionFilterContainer.style.display = isHidden ? 'block' : 'none';
    if (isHidden && targetSectionInput) {
      targetSectionInput.focus();
    }
  });
}

if (btnApplySection && targetSectionInput && selectedSectionBadge) {
  btnApplySection.addEventListener('click', () => {
    const sec = targetSectionInput.value.trim();
    if (sec) {
      selectedSectionBadge.textContent = `Section ${sec.toUpperCase()}`;
    } else {
      selectedSectionBadge.textContent = 'All Sections';
    }
    sectionFilterContainer.style.display = 'none';
  });

  targetSectionInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      btnApplySection.click();
    }
  });
}

// ---- Start Bot ----
document.getElementById('btnStartBot').addEventListener('click', async () => {
  const code = document.getElementById('udiseCode').value.trim();
  const pass = document.getElementById('password').value;
  if (!code || !pass) return alert('Enter both UDISE Code and Password.');
  if (!uploadedData) return alert('No data. Go back and upload a file.');
  goStep(3);
  try {
    const res = await fetch(API + '/api/start-automation', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ 
        socketId: socket.id, 
        udiseCode: code, 
        password: pass, 
        data: uploadedData.data,
        filePath: uploadedData.filePath,
        semiAutomatic: document.getElementById('semiAutomatic').checked,
        targetSection: targetSectionInput ? targetSectionInput.value.trim() : ''
      })
    });
    const r = await res.json();
    if (!res.ok) throw new Error(r.message || r.error || 'Failed to start');
    sessionId = r.sessionId;
  } catch (err) { addLog({ message: 'Start failed: ' + err.message, type: 'error', timestamp: new Date().toISOString() }); }
});

// ---- Stop / Restart ----
document.getElementById('btnStop').addEventListener('click', async () => {
  if (!sessionId || !confirm('Stop the bot?')) return;
  await fetch(API + '/api/stop-automation', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ sessionId }) }).catch(() => {});
  addLog({ message: 'Stop signal sent...', type: 'warn', timestamp: new Date().toISOString() });
});

document.getElementById('btnRestart').addEventListener('click', () => {
  sessionId = null; uploadedData = null;
  document.getElementById('consoleBody').innerHTML = '';
  document.getElementById('progressBar').style.width = '0';
  document.getElementById('progressPct').textContent = '0%';
  document.getElementById('progressCount').textContent = '0 / 0 students';
  document.getElementById('currentStudent').classList.add('hidden');
  document.getElementById('captchaModal').classList.add('hidden');
  document.getElementById('resultsGrid').classList.add('hidden');
  const reportSection = document.getElementById('reportSection');
  if (reportSection) reportSection.classList.add('hidden');
  document.getElementById('statusText').textContent = 'Initializing...';
  document.getElementById('btnStop').classList.remove('hidden');
  document.getElementById('btnRestart').classList.add('hidden');
  resetUpload(); goStep(1);
});

document.getElementById('copyLogsBtn').addEventListener('click', () => {
  const body = document.getElementById('consoleBody');
  const lines = body.querySelectorAll('.console-line');
  const text = Array.from(lines).map(l => {
    const time = l.querySelector('.cl-time')?.textContent || '';
    const msg = l.querySelector('.cl-msg')?.textContent || '';
    return `[${time}] ${msg}`;
  }).join('\n');
  navigator.clipboard.writeText(text).then(() => {
    const btn = document.getElementById('copyLogsBtn');
    const orig = btn.title;
    btn.title = '✓ Copied!';
    btn.style.color = '#4ade80';
    setTimeout(() => { btn.title = orig; btn.style.color = ''; }, 2000);
  }).catch(() => alert('Failed to copy logs'));
});
document.getElementById('clearLogsBtn').addEventListener('click', () => { document.getElementById('consoleBody').innerHTML = ''; });

// ---- CAPTCHA Submit ----
function submitCaptcha() {
  const text = document.getElementById('captchaInput').value.trim();
  if (!text) return alert('Please enter the CAPTCHA text.');
  socket.emit('captcha-answer', { text });
  document.getElementById('captchaModal').classList.add('hidden');
  addLog({ message: 'CAPTCHA submitted: ' + text, type: 'info', timestamp: new Date().toISOString() });
}
document.getElementById('captchaSubmitBtn').addEventListener('click', submitCaptcha);
document.getElementById('captchaInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') submitCaptcha(); });

// ---- Student Confirmation Submit ----
function submitStudentConfirmation(action) {
  clearConfirmCountdown();
  if (action === 'skip') {
    socket.emit('confirm-student-response', { action: 'skip' });
  } else {
    const marks = document.getElementById('confMarks').value.trim();
    const percentage = document.getElementById('confPercent').value.trim();
    const attendance = document.getElementById('confDays').value.trim();
    const progressionStatus = document.getElementById('confProgression').value;
    const sameSchool = document.getElementById('confSameSchool').value;
    
    socket.emit('confirm-student-response', {
      action: 'submit',
      marks,
      percentage,
      attendance,
      progressionStatus,
      sameSchool
    });
  }
  document.getElementById('confirmModal').classList.add('hidden');
}

// ---- Countdown Timer for Auto-Confirm ----
function startConfirmCountdown(seconds) {
  clearConfirmCountdown();
  confirmCountdownSeconds = seconds;
  const badge = document.getElementById('confirmCountdownBadge');
  if (badge) {
    badge.classList.remove('hidden');
    badge.textContent = `Auto-confirm in ${confirmCountdownSeconds}s`;
  }
  confirmCountdownTimer = setInterval(() => {
    confirmCountdownSeconds--;
    if (badge) badge.textContent = `Auto-confirm in ${confirmCountdownSeconds}s`;
    if (confirmCountdownSeconds <= 0) {
      clearConfirmCountdown();
    }
  }, 1000);
}

function clearConfirmCountdown() {
  if (confirmCountdownTimer) {
    clearInterval(confirmCountdownTimer);
    confirmCountdownTimer = null;
  }
  const badge = document.getElementById('confirmCountdownBadge');
  if (badge) badge.classList.add('hidden');
}

// ---- Error Flash Toast ----
function showErrorFlash(studentName, message) {
  const existing = document.getElementById('errorFlashToast');
  if (existing) existing.remove();
  
  const toast = document.createElement('div');
  toast.id = 'errorFlashToast';
  toast.className = 'error-flash-toast';
  toast.innerHTML = `
    <div class="flash-icon">⚠️</div>
    <div class="flash-body">
      <strong>${studentName || 'Error'}</strong>
      <span>${message || 'An unknown error occurred'}</span>
    </div>
    <button class="flash-close" onclick="this.parentElement.remove()">&times;</button>
  `;
  document.body.appendChild(toast);
  // Auto-remove after 8 seconds
  setTimeout(() => { if (toast.parentElement) toast.remove(); }, 8000);
}

// ---- Refresh Data File ----
async function refreshFileData() {
  if (!uploadedData || !uploadedData.filePath) {
    alert('No file loaded to refresh.');
    return;
  }
  // Re-fetch the preview from the already uploaded data
  showFileData(uploadedData);
  addLog({ message: 'File data refreshed.', type: 'info', timestamp: new Date().toISOString() });
}

// ---- Clear Memory of Entered Students ----
async function clearEnteredMemory() {
  try {
    const res = await fetch(API + '/api/clear-memory', { method: 'POST' });
    const data = await res.json();
    if (data.success) {
      addLog({ message: `🧹 ${data.message}`, type: 'info', timestamp: new Date().toISOString() });
      alert(data.message);
      updateMemoryCount();
    }
  } catch (err) {
    alert('Failed to clear memory: ' + err.message);
  }
}

async function updateMemoryCount() {
  try {
    const res = await fetch(API + '/api/memory-count');
    const data = await res.json();
    const badge = document.getElementById('memoryCountBadge');
    if (badge) {
      badge.textContent = data.count > 0 ? `${data.count} students remembered` : 'No students in memory';
    }
  } catch (e) { /* ignore */ }
}

document.getElementById('confirmSubmitBtn').addEventListener('click', () => submitStudentConfirmation('submit'));
document.getElementById('confirmSkipBtn').addEventListener('click', () => submitStudentConfirmation('skip'));

// ---- Lightbox Zoom Handler ----
const lightboxOverlay = document.getElementById('lightboxOverlay');
const lightboxImg = document.getElementById('lightboxImg');
const captchaImg = document.getElementById('captchaImg');
const confirmImg = document.getElementById('confirmImg');

function openZoom(src) {
  if (!src) return;
  lightboxImg.src = src;
  lightboxOverlay.classList.remove('hidden');
}

if (captchaImg) {
  captchaImg.addEventListener('click', () => openZoom(captchaImg.src));
}
if (confirmImg) {
  confirmImg.addEventListener('click', () => openZoom(confirmImg.src));
}
if (lightboxOverlay) {
  lightboxOverlay.addEventListener('click', () => {
    lightboxOverlay.classList.add('hidden');
  });
}

// ---- Global Keyboard Shortcuts (Key Functions) ----
window.addEventListener('keydown', (e) => {
  const confirmModal = document.getElementById('confirmModal');
  const captchaModal = document.getElementById('captchaModal');
  const lightboxOverlay = document.getElementById('lightboxOverlay');

  // 1. If Zoom Lightbox is active, Esc key closes it
  if (lightboxOverlay && !lightboxOverlay.classList.contains('hidden')) {
    if (e.key === 'Escape') {
      lightboxOverlay.classList.add('hidden');
      e.preventDefault();
      return;
    }
  }

  // 2. If Student Confirmation modal is active (visible)
  if (confirmModal && !confirmModal.classList.contains('hidden')) {
    // Escape key skips the student
    if (e.key === 'Escape') {
      submitStudentConfirmation('skip');
      e.preventDefault();
      return;
    }
    
    // Enter key submits the student confirmation
    if (e.key === 'Enter') {
      submitStudentConfirmation('submit');
      e.preventDefault();
      return;
    }

    // Space key or "z" zooms/toggles lightbox screenshot (if not focused inside input/select fields)
    if ((e.key === ' ' || e.key.toLowerCase() === 'z') && 
        document.activeElement.tagName !== 'INPUT' && 
        document.activeElement.tagName !== 'SELECT') {
      const confirmImg = document.getElementById('confirmImg');
      if (confirmImg && confirmImg.src) {
        openZoom(confirmImg.src);
      }
      e.preventDefault();
      return;
    }
  }

  // 3. If CAPTCHA modal is active, focus CAPTCHA input on Escape
  if (captchaModal && !captchaModal.classList.contains('hidden')) {
    if (e.key === 'Escape') {
      document.getElementById('captchaInput').focus();
      e.preventDefault();
      return;
    }
  }
});

// Attach keydown listener to input fields to trigger confirmation on Enter
const confirmInputIds = ['confMarks', 'confPercent', 'confDays', 'confProgression', 'confSameSchool'];
confirmInputIds.forEach(id => {
  const el = document.getElementById(id);
  if (el) {
    el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') {
        submitStudentConfirmation('submit');
        e.preventDefault();
      }
    });
  }
});

// ---- Progress ----
function updateProgress(d) {
  document.getElementById('progressBar').style.width = d.percent + '%';
  document.getElementById('progressPct').textContent = d.percent + '%';
  const map = {
    navigating:'Navigating to UDISE+...', logging_in:'Logging in...', waiting_captcha:'Solve CAPTCHA in browser...',
    logged_in:'Logged in!', navigating_student:'Finding student section...', ready:'Ready to process',
    processing:`Processing ${d.current||0} of ${d.total||0}`, completed:'Complete!'
  };
  document.getElementById('statusText').textContent = map[d.step] || d.step;
  if (d.current && d.total) document.getElementById('progressCount').textContent = `${d.current} / ${d.total} students`;
  if (d.studentName) { document.getElementById('currentStudent').classList.remove('hidden'); document.getElementById('currentName').textContent = d.studentName; }
  if (d.step === 'completed') { document.getElementById('currentStudent').classList.add('hidden'); document.getElementById('btnStop').classList.add('hidden'); document.getElementById('btnRestart').classList.remove('hidden'); }
}

// ---- Log ----
function addLog(e) {
  const c = document.getElementById('consoleBody');
  const div = document.createElement('div');
  div.className = 'console-line ' + (e.type || 'info');
  const t = e.timestamp ? new Date(e.timestamp).toLocaleTimeString() : '--:--:--';
  const tmp = document.createElement('span'); tmp.textContent = e.message;
  div.innerHTML = `<span class="cl-time">${t}</span><span class="cl-msg">${tmp.innerHTML}</span>`;
  c.appendChild(div); c.scrollTop = c.scrollHeight;
}

// ---- Results ----
function showResults(r) {
  document.getElementById('resultsGrid').classList.remove('hidden');
  document.getElementById('rSuccess').textContent = r.success;
  document.getElementById('rFailed').textContent = r.failed;
  document.getElementById('rTotal').textContent = r.total;
  document.getElementById('btnStop').classList.add('hidden');
  document.getElementById('btnRestart').classList.remove('hidden');

  // Audit Warning Report compilations
  const reportSection = document.getElementById('reportSection');
  const skippedList = document.getElementById('skippedList');
  const lowPercentList = document.getElementById('lowPercentList');
  
  if (reportSection && r.details && Array.isArray(r.details)) {
    // 1. Compile Left-Out / Skipped / Failed Students
    const leftOut = r.details.filter(d => d.status === 'failed' || d.status === 'skipped');
    document.getElementById('countSkipped').textContent = leftOut.length;
    
    if (leftOut.length > 0) {
      skippedList.innerHTML = leftOut.map(s => {
        const reason = s.error || 'Skipped by user';
        const sectionInfo = s.section ? `Section ${s.section}` : 'N/A';
        return `<li>
          <div>
            <span class="student-name">${s.name}</span>
            <span style="font-size:10px; display:block; color:var(--text-muted); margin-top:2px;">PEN: ${s.pen || 'N/A'} • ${sectionInfo}</span>
          </div>
          <span class="student-meta" style="color:var(--danger); border:1px solid rgba(239,68,68,0.25); background:rgba(239,68,68,0.05);">${reason}</span>
        </li>`;
      }).join('');
    } else {
      skippedList.innerHTML = `<li class="empty-list">No skipped or left-out students. All processed perfectly!</li>`;
    }

    // 2. Compile Students Promoted with Marks/Percentage < 39%
    const lowPercent = r.details.filter(d => {
      if (d.status !== 'success') return false;
      const pct = parseFloat(d.percentage);
      if (isNaN(pct) || pct >= 39) return false;
      
      const prog = String(d.progressionStatus || '').toLowerCase();
      return prog.includes('promote') || prog.includes('pass') || prog === 'promoted';
    });
    
    document.getElementById('countLowPercent').textContent = lowPercent.length;
    
    if (lowPercent.length > 0) {
      lowPercentList.innerHTML = lowPercent.map(s => {
        return `<li>
          <div>
            <span class="student-name">${s.name}</span>
            <span style="font-size:10px; display:block; color:var(--text-muted); margin-top:2px;">PEN: ${s.pen || 'N/A'} • Section ${s.section}</span>
          </div>
          <span class="student-meta" style="color:var(--warn); border:1px solid rgba(245,158,11,0.25); background:rgba(245,158,11,0.05);">
            Marks: <strong>${s.percentage}%</strong> • ${s.progressionStatus}
          </span>
        </li>`;
      }).join('');
    } else {
      lowPercentList.innerHTML = `<li class="empty-list">No students found with &lt;39% marks and Promoted status.</li>`;
    }

    reportSection.classList.remove('hidden');
  }
}

// ---- Init ----
initSocket();
initParticles();
updateMemoryCount();

// ---- Premium Particle Background ----
function initParticles() {
  const canvas = document.getElementById('particlesCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  
  let width = canvas.width = window.innerWidth;
  let height = canvas.height = window.innerHeight;
  
  const particles = [];
  const maxParticles = 60;
  
  class Particle {
    constructor() {
      this.reset();
    }
    
    reset() {
      this.x = Math.random() * width;
      this.y = Math.random() * height + Math.random() * 100;
      this.size = Math.random() * 1.8 + 0.6;
      this.speedY = -(Math.random() * 0.4 + 0.15);
      this.speedX = Math.random() * 0.2 - 0.1;
      this.alpha = Math.random() * 0.4 + 0.1;
      this.hue = Math.random() > 0.5 ? 215 : 275; // Ambient blue or purple
    }
    
    update() {
      this.y += this.speedY;
      this.x += this.speedX;
      if (this.y < -10) {
        this.reset();
        this.y = height + 10;
      }
    }
    
    draw() {
      ctx.save();
      ctx.globalAlpha = this.alpha;
      ctx.beginPath();
      ctx.arc(this.x, this.y, this.size, 0, Math.PI * 2);
      ctx.fillStyle = `hsla(${this.hue}, 100%, 70%, ${this.alpha})`;
      ctx.shadowBlur = 6;
      ctx.shadowColor = `hsla(${this.hue}, 100%, 70%, 0.8)`;
      ctx.fill();
      ctx.restore();
    }
  }
  
  for (let i = 0; i < maxParticles; i++) {
    particles.push(new Particle());
    particles[i].y = Math.random() * height;
  }
  
  window.addEventListener('resize', () => {
    width = canvas.width = window.innerWidth;
    height = canvas.height = window.innerHeight;
  });
  
  function animate() {
    ctx.clearRect(0, 0, width, height);
    for (let p of particles) {
      p.update();
      p.draw();
    }
    requestAnimationFrame(animate);
  }
  
  animate();
}
