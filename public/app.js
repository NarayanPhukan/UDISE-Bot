// ===== UDISE+ Bot Frontend =====
const API = '';
let socket, uploadedData, sessionId;

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
        semiAutomatic: document.getElementById('semiAutomatic').checked
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

document.getElementById('confirmSubmitBtn').addEventListener('click', () => submitStudentConfirmation('submit'));
document.getElementById('confirmSkipBtn').addEventListener('click', () => submitStudentConfirmation('skip'));

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
}

// ---- Init ----
initSocket();
initParticles();

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
