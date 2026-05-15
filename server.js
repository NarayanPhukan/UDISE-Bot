const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const puppeteer = require('puppeteer');
const http = require('http');
const { Server } = require('socket.io');
const path = require('path');
const fs = require('fs');
const cors = require('cors');

const app = express();
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: '*' } });

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Configure multer for file uploads
const upload = multer({
  dest: path.join(__dirname, 'uploads'),
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (['.xlsx', '.xls', '.csv'].includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Only Excel files (.xlsx, .xls) and CSV files are allowed'));
    }
  },
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB limit
});

// Ensure uploads directory exists
if (!fs.existsSync(path.join(__dirname, 'uploads'))) {
  fs.mkdirSync(path.join(__dirname, 'uploads'));
}

// Store active automation sessions
const activeSessions = new Map();

// ========================
// API ROUTES
// ========================

// Upload & parse Excel file
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '' });

    if (data.length === 0) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Excel file is empty' });
    }

    // Normalize column headers (case-insensitive matching)
    const normalizedData = data.map((row, index) => {
      const normalized = {};
      for (const [key, value] of Object.entries(row)) {
        const lowerKey = key.toLowerCase().trim();
        if (lowerKey.includes('student') && lowerKey.includes('name')) {
          normalized.studentName = String(value).trim();
        } else if (lowerKey.includes('class') || lowerKey === 'cls') {
          normalized.class = String(value).trim();
        } else if (lowerKey.includes('pen')) {
          normalized.penNo = String(value).trim();
        } else if (lowerKey.includes('attendance')) {
          normalized.attendance = String(value).trim();
        } else if (lowerKey.includes('percent') || lowerKey.includes('%')) {
          normalized.percentage = String(value).trim();
        } else if (lowerKey.includes('progress') || lowerKey.includes('promotion') || lowerKey.includes('status')) {
          normalized.progressionStatus = String(value).trim();
        } else if (lowerKey.includes('same') && lowerKey.includes('school')) {
          normalized.sameSchool = String(value).trim();
        }
      }
      normalized.rowIndex = index + 1;
      return normalized;
    });

    // Validate required fields
    const requiredFields = ['studentName', 'penNo'];
    const missingFields = [];
    for (const field of requiredFields) {
      if (!normalizedData[0][field]) {
        missingFields.push(field);
      }
    }

    res.json({
      success: true,
      fileName: req.file.originalname,
      filePath: req.file.path,
      totalRows: normalizedData.length,
      columns: Object.keys(data[0] || {}),
      preview: normalizedData.slice(0, 5),
      data: normalizedData,
      warnings: missingFields.length > 0
        ? `Could not auto-detect columns: ${missingFields.join(', ')}. Please check your column headers.`
        : null
    });
  } catch (err) {
    console.error('Upload error:', err);
    res.status(500).json({ error: 'Failed to parse Excel file: ' + err.message });
  }
});

// Start automation
app.post('/api/start-automation', (req, res) => {
  const { socketId, udiseCode, password, data, filePath } = req.body;

  if (!udiseCode || !password || !data || data.length === 0) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  const sessionId = Date.now().toString();
  activeSessions.set(sessionId, { status: 'starting', progress: 0 });

  // Start automation in background
  runAutomation(sessionId, socketId, udiseCode, password, data, filePath);

  res.json({ success: true, sessionId });
});

// Stop automation
app.post('/api/stop-automation', (req, res) => {
  const { sessionId } = req.body;
  const session = activeSessions.get(sessionId);
  if (session) {
    session.status = 'stopping';
    res.json({ success: true });
  } else {
    res.status(404).json({ error: 'Session not found' });
  }
});

// Get automation status
app.get('/api/status/:sessionId', (req, res) => {
  const session = activeSessions.get(req.params.sessionId);
  if (session) {
    res.json(session);
  } else {
    res.status(404).json({ error: 'Session not found' });
  }
});

// ========================
// BROWSER AUTOMATION
// ========================

async function runAutomation(sessionId, socketId, udiseCode, password, students, filePath) {
  const socket = io.sockets.sockets.get(socketId);
  const session = activeSessions.get(sessionId);
  let browser = null;

  const emit = (event, data) => {
    if (socket) socket.emit(event, data);
    io.emit(`session:${sessionId}`, { event, data });
  };

  const log = (message, type = 'info') => {
    const entry = { timestamp: new Date().toISOString(), message, type };
    emit('log', entry);
    console.log(`[${type.toUpperCase()}] ${message}`);
  };

  try {
    session.status = 'launching';
    log('🚀 Launching browser...');

    browser = await puppeteer.launch({
      headless: 'new',
      executablePath: process.env.PUPPETEER_EXECUTABLE_PATH || undefined,
      defaultViewport: { width: 1366, height: 768 },
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-gpu',
        '--disable-software-rasterizer',
        '--single-process',
        '--no-zygote',
        '--window-size=1366,768'
      ],
      slowMo: 50
    });

    const page = await browser.newPage();

    // Set timeout for page operations
    page.setDefaultTimeout(30000);
    page.setDefaultNavigationTimeout(60000);

    // ---- Step 1: Navigate to UDISE+ ----
    session.status = 'navigating';
    log('🌐 Navigating to UDISE+ portal...');
    emit('progress', { step: 'navigating', percent: 5 });

    await page.goto('https://udiseplus.gov.in/#/page/login', {
      waitUntil: 'networkidle2',
      timeout: 60000
    });

    await page.waitForTimeout(3000);
    log('✅ UDISE+ portal loaded');

    // ---- Step 2: Login ----
    session.status = 'logging_in';
    log('🔐 Attempting login...');
    emit('progress', { step: 'logging_in', percent: 10 });

    // Wait for login form to appear
    await page.waitForSelector('input[type="text"], input[placeholder*="UDISE"], input[name*="user"], input[name*="udise"]', { timeout: 15000 })
      .catch(() => log('⚠️ Could not find username field with standard selectors, trying alternatives...', 'warn'));

    // Try multiple selectors for the username field
    const usernameSelectors = [
      'input[placeholder*="UDISE"]',
      'input[placeholder*="udise"]',
      'input[placeholder*="User"]',
      'input[placeholder*="user"]',
      'input[name*="user"]',
      'input[name*="udise"]',
      'input[type="text"]'
    ];

    let usernameField = null;
    for (const sel of usernameSelectors) {
      usernameField = await page.$(sel);
      if (usernameField) {
        log(`Found username field: ${sel}`);
        break;
      }
    }

    if (!usernameField) {
      // Take screenshot for debugging
      await page.screenshot({ path: path.join(__dirname, 'uploads', 'login-debug.png') });
      throw new Error('Could not find the username input field on the login page. The page structure may have changed.');
    }

    await usernameField.click({ clickCount: 3 });
    await usernameField.type(udiseCode, { delay: 50 });

    // Find password field
    const passwordSelectors = [
      'input[type="password"]',
      'input[placeholder*="Password"]',
      'input[placeholder*="password"]',
      'input[name*="pass"]'
    ];

    let passwordField = null;
    for (const sel of passwordSelectors) {
      passwordField = await page.$(sel);
      if (passwordField) break;
    }

    if (!passwordField) {
      throw new Error('Could not find the password input field');
    }

    await passwordField.click({ clickCount: 3 });
    await passwordField.type(password, { delay: 50 });

    log('📝 Credentials entered');

    // Handle CAPTCHA if present
    const captchaExists = await page.$('canvas, img[src*="captcha"], .captcha, input[placeholder*="captcha"], input[placeholder*="Captcha"]');
    if (captchaExists) {
      log('⚠️ CAPTCHA detected! Sending screenshot to you...', 'warn');
      session.status = 'waiting_captcha';
      emit('progress', { step: 'waiting_captcha', percent: 12 });

      // Take screenshot and send to frontend as base64
      const screenshotBuffer = await page.screenshot({ encoding: 'base64' });
      emit('captcha', { image: 'data:image/png;base64,' + screenshotBuffer });

      // Wait for user to send CAPTCHA text back via socket
      const captchaText = await waitForCaptchaAnswer(socket, 180000);
      log(`📝 Received CAPTCHA answer: ${captchaText}`);

      // Find and fill CAPTCHA input
      const captchaInputSels = [
        'input[placeholder*="captcha"]', 'input[placeholder*="Captcha"]',
        'input[placeholder*="CAPTCHA"]', 'input[name*="captcha"]',
        'input[id*="captcha"]', 'input[id*="Captcha"]'
      ];
      let captchaField = null;
      for (const sel of captchaInputSels) {
        captchaField = await page.$(sel);
        if (captchaField) break;
      }
      if (captchaField) {
        await captchaField.click({ clickCount: 3 });
        await captchaField.type(captchaText, { delay: 40 });
      } else {
        log('⚠️ Could not find CAPTCHA input field', 'warn');
      }
    }

    // Click login button
    const loginSelectors = [
      'button[type="submit"]', 'button.btn-primary', 'button.login-btn',
      'input[type="submit"]', 'button'
    ];
    for (const sel of loginSelectors) {
      const btn = await page.$(sel);
      if (btn) {
        const text = await page.evaluate(el => el.textContent, btn);
        if (text && (text.toLowerCase().includes('login') || text.toLowerCase().includes('sign in') || text.toLowerCase().includes('submit'))) {
          await btn.click();
          break;
        }
      }
    }
    await page.waitForTimeout(5000);

    // Check if login was successful
    const currentUrl = page.url();
    log(`📍 Current URL: ${currentUrl}`);

    // Check for login errors
    const errorElement = await page.$('.error-message, .alert-danger, .text-danger, .error');
    if (errorElement) {
      const errorText = await page.evaluate(el => el.textContent, errorElement);
      if (errorText && errorText.trim()) {
        log(`❌ Login error: ${errorText.trim()}`, 'error');
        throw new Error(`Login failed: ${errorText.trim()}`);
      }
    }

    session.status = 'logged_in';
    log('✅ Login successful!');
    emit('progress', { step: 'logged_in', percent: 15 });

    // ---- Step 3: Navigate to Student Section ----
    log('📂 Navigating to student data section...');
    emit('progress', { step: 'navigating_student', percent: 20 });

    // UDISE+ student data URLs - try different possible paths
    const studentPaths = [
      'https://udiseplus.gov.in/#/page/student',
      'https://udiseplus.gov.in/#/page/studentDetails',
      'https://udiseplus.gov.in/#/page/student-data',
      'https://udiseplus.gov.in/#/Student',
    ];

    // First try to find student menu link
    const menuSelectors = [
      'a[href*="student"]',
      'a[href*="Student"]',
      'li:has-text("Student")',
      'span:has-text("Student")',
      '.nav-link:has-text("Student")',
      'a:has-text("Student")'
    ];

    let foundStudentMenu = false;
    for (const sel of menuSelectors) {
      try {
        const link = await page.$(sel);
        if (link) {
          await link.click();
          await page.waitForTimeout(3000);
          foundStudentMenu = true;
          log('✅ Found and clicked Student menu');
          break;
        }
      } catch (e) { /* try next */ }
    }

    if (!foundStudentMenu) {
      // Try navigating directly
      for (const studentUrl of studentPaths) {
        try {
          await page.goto(studentUrl, { waitUntil: 'networkidle2', timeout: 15000 });
          await page.waitForTimeout(2000);
          log(`Tried navigating to: ${studentUrl}`);
          break;
        } catch (e) { /* try next */ }
      }
    }

    emit('progress', { step: 'ready', percent: 25 });
    log('📋 Ready to process student records');

    // ---- Step 4: Process each student ----
    session.status = 'processing';
    const results = {
      total: students.length,
      success: 0,
      failed: 0,
      skipped: 0,
      details: []
    };

    for (let i = 0; i < students.length; i++) {
      // Check if session was stopped
      if (activeSessions.get(sessionId)?.status === 'stopping') {
        log('⏹️ Automation stopped by user', 'warn');
        break;
      }

      const student = students[i];
      const progressPercent = 25 + Math.round((i / students.length) * 70);

      log(`\n📝 Processing student ${i + 1}/${students.length}: ${student.studentName || 'Unknown'} (PEN: ${student.penNo || 'N/A'})`);
      emit('progress', {
        step: 'processing',
        percent: progressPercent,
        current: i + 1,
        total: students.length,
        studentName: student.studentName
      });

      try {
        await processStudent(page, student, log);
        results.success++;
        results.details.push({
          row: student.rowIndex,
          name: student.studentName,
          pen: student.penNo,
          status: 'success'
        });
        log(`✅ Successfully updated: ${student.studentName}`, 'success');
      } catch (err) {
        results.failed++;
        results.details.push({
          row: student.rowIndex,
          name: student.studentName,
          pen: student.penNo,
          status: 'failed',
          error: err.message
        });
        log(`❌ Failed for ${student.studentName}: ${err.message}`, 'error');

        // Take screenshot on failure
        try {
          const screenshotPath = path.join(__dirname, 'uploads', `error-${student.penNo || i}.png`);
          await page.screenshot({ path: screenshotPath });
          log(`📸 Error screenshot saved`, 'info');
        } catch (e) { /* ignore screenshot errors */ }
      }

      // Small delay between students
      await page.waitForTimeout(1500);
    }

    // ---- Step 5: Complete ----
    session.status = 'completed';
    session.results = results;

    log(`\n🏁 Automation Complete!`);
    log(`✅ Success: ${results.success}/${results.total}`);
    log(`❌ Failed: ${results.failed}/${results.total}`);
    if (results.skipped > 0) log(`⏭️ Skipped: ${results.skipped}/${results.total}`);

    emit('progress', { step: 'completed', percent: 100 });
    emit('complete', results);

  } catch (err) {
    session.status = 'error';
    session.error = err.message;
    log(`💥 Automation error: ${err.message}`, 'error');
    emit('error', { message: err.message });
  } finally {
    // Clean up uploaded file
    if (filePath && fs.existsSync(filePath)) {
      try { fs.unlinkSync(filePath); } catch (e) { /* ignore */ }
    }

    // Close browser after a delay (allow user to see final state)
    if (browser) {
      setTimeout(async () => {
        try { await browser.close(); } catch (e) { /* ignore */ }
      }, 10000);
    }
  }
}

async function processStudent(page, student, log) {
  // Search for student by PEN number
  if (student.penNo) {
    // Look for search/filter input
    const searchSelectors = [
      'input[placeholder*="PEN"]',
      'input[placeholder*="pen"]',
      'input[placeholder*="Search"]',
      'input[placeholder*="search"]',
      'input[name*="pen"]',
      'input[name*="search"]',
      'input.search-input',
      '#searchInput',
      'input[type="search"]'
    ];

    let searchField = null;
    for (const sel of searchSelectors) {
      searchField = await page.$(sel);
      if (searchField) break;
    }

    if (searchField) {
      await searchField.click({ clickCount: 3 });
      await searchField.type(student.penNo, { delay: 30 });
      await page.waitForTimeout(2000);

      // Click search button if present
      const searchBtns = await page.$$('button');
      for (const btn of searchBtns) {
        const text = await page.evaluate(el => el.textContent, btn);
        if (text && (text.toLowerCase().includes('search') || text.toLowerCase().includes('find') || text.toLowerCase().includes('go'))) {
          await btn.click();
          await page.waitForTimeout(2000);
          break;
        }
      }
    }
  }

  // Try to find and click on the student record
  const studentRow = await findStudentRow(page, student);
  if (studentRow) {
    await studentRow.click();
    await page.waitForTimeout(2000);
  }

  // Now try to fill in the data fields
  await fillStudentData(page, student, log);

  // Save the record
  await saveRecord(page, log);
}

async function findStudentRow(page, student) {
  // Try to find student in a table by PEN or name
  const rows = await page.$$('tr, .student-row, .record-row');
  for (const row of rows) {
    const text = await page.evaluate(el => el.textContent, row);
    if (student.penNo && text.includes(student.penNo)) {
      return row;
    }
    if (student.studentName && text.toLowerCase().includes(student.studentName.toLowerCase())) {
      return row;
    }
  }

  // Try clicking edit button if visible
  const editBtns = await page.$$('button, a, .edit-btn, .btn-edit');
  for (const btn of editBtns) {
    const text = await page.evaluate(el => el.textContent + ' ' + (el.title || ''), btn);
    if (text.toLowerCase().includes('edit') || text.toLowerCase().includes('update')) {
      return btn;
    }
  }

  return null;
}

async function fillStudentData(page, student, log) {
  // Fill attendance
  if (student.attendance) {
    await fillField(page, [
      'input[name*="attendance"]',
      'input[name*="Attendance"]',
      'input[placeholder*="Attendance"]',
      'input[placeholder*="attendance"]',
      '#attendance',
      'input[id*="attendance"]'
    ], student.attendance, log, 'Attendance');
  }

  // Fill percentage
  if (student.percentage) {
    await fillField(page, [
      'input[name*="percent"]',
      'input[name*="Percent"]',
      'input[placeholder*="Percent"]',
      'input[placeholder*="percent"]',
      '#percentage',
      'input[id*="percent"]',
      'input[name*="marks"]'
    ], student.percentage, log, 'Percentage');
  }

  // Fill progression status (dropdown or input)
  if (student.progressionStatus) {
    const progressionValue = student.progressionStatus.toLowerCase();
    const isPromoted = progressionValue.includes('promot') || progressionValue.includes('pass');

    // Try select/dropdown first
    const selectSelectors = [
      'select[name*="progress"]',
      'select[name*="Progress"]',
      'select[name*="promotion"]',
      'select[name*="status"]',
      'select[id*="progress"]',
      'select[id*="promotion"]',
      '#progressionStatus'
    ];

    let foundSelect = false;
    for (const sel of selectSelectors) {
      const select = await page.$(sel);
      if (select) {
        const options = await page.$$eval(`${sel} option`, opts =>
          opts.map(o => ({ value: o.value, text: o.textContent.trim().toLowerCase() }))
        );

        for (const opt of options) {
          if ((isPromoted && (opt.text.includes('promot') || opt.text.includes('pass'))) ||
              (!isPromoted && (opt.text.includes('fail') || opt.text.includes('retain') || opt.text.includes('repeat')))) {
            await page.select(sel, opt.value);
            log(`  ✓ Progression: ${opt.text}`);
            foundSelect = true;
            break;
          }
        }
        if (foundSelect) break;
      }
    }

    if (!foundSelect) {
      // Try radio buttons
      const radioLabels = await page.$$('label, .radio-label');
      for (const label of radioLabels) {
        const text = await page.evaluate(el => el.textContent.toLowerCase(), label);
        if ((isPromoted && text.includes('promot')) || (!isPromoted && text.includes('fail'))) {
          await label.click();
          log(`  ✓ Progression: clicked ${text.trim()}`);
          break;
        }
      }
    }
  }

  // Fill "Studying in same school"
  if (student.sameSchool) {
    const isYes = student.sameSchool.toLowerCase().includes('yes') || student.sameSchool === '1' || student.sameSchool.toLowerCase() === 'true';

    const sameSchoolSelectors = [
      'select[name*="same"]',
      'select[name*="school"]',
      'select[id*="same"]',
      '#sameSchool'
    ];

    let foundSameSchool = false;
    for (const sel of sameSchoolSelectors) {
      const select = await page.$(sel);
      if (select) {
        const options = await page.$$eval(`${sel} option`, opts =>
          opts.map(o => ({ value: o.value, text: o.textContent.trim().toLowerCase() }))
        );
        for (const opt of options) {
          if ((isYes && opt.text.includes('yes')) || (!isYes && opt.text.includes('no'))) {
            await page.select(sel, opt.value);
            log(`  ✓ Same School: ${opt.text}`);
            foundSameSchool = true;
            break;
          }
        }
        if (foundSameSchool) break;
      }
    }

    if (!foundSameSchool) {
      // Try radio buttons
      const radioLabels = await page.$$('label, .radio-label');
      for (const label of radioLabels) {
        const text = await page.evaluate(el => el.textContent.toLowerCase(), label);
        if ((isYes && text.includes('yes')) || (!isYes && text.includes('no'))) {
          const isNearSchool = await page.evaluate(el => {
            const parent = el.closest('.form-group, .form-row, .field-group, div');
            return parent ? parent.textContent.toLowerCase().includes('same school') : false;
          }, label);
          if (isNearSchool) {
            await label.click();
            log(`  ✓ Same School: clicked ${text.trim()}`);
            break;
          }
        }
      }
    }
  }
}

async function fillField(page, selectors, value, log, fieldName) {
  for (const sel of selectors) {
    const field = await page.$(sel);
    if (field) {
      await field.click({ clickCount: 3 });
      await field.type(String(value), { delay: 30 });
      log(`  ✓ ${fieldName}: ${value}`);
      return true;
    }
  }
  log(`  ⚠️ Could not find ${fieldName} field`, 'warn');
  return false;
}

async function saveRecord(page, log) {
  const saveSelectors = [
    'button[type="submit"]',
    'button:has-text("Save")',
    'button:has-text("Update")',
    'button:has-text("Submit")',
    'button.btn-primary',
    'button.save-btn',
    'input[type="submit"]'
  ];

  for (const sel of saveSelectors) {
    try {
      const btn = await page.$(sel);
      if (btn) {
        const text = await page.evaluate(el => el.textContent, btn);
        if (text && (text.toLowerCase().includes('save') || text.toLowerCase().includes('update') || text.toLowerCase().includes('submit'))) {
          await btn.click();
          await page.waitForTimeout(2000);
          log('  💾 Record saved');
          return true;
        }
      }
    } catch (e) { /* try next */ }
  }

  // Fallback: click any primary button
  const allBtns = await page.$$('button.btn-primary, button.btn-success');
  if (allBtns.length > 0) {
    await allBtns[allBtns.length - 1].click();
    await page.waitForTimeout(2000);
    log('  💾 Clicked save button');
    return true;
  }

  log('  ⚠️ Could not find save button', 'warn');
  return false;
}

// ========================
// CAPTCHA HELPER
// ========================

function waitForCaptchaAnswer(socket, timeoutMs = 180000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error('CAPTCHA answer timeout — no response within 3 minutes'));
    }, timeoutMs);

    const handler = (data) => {
      clearTimeout(timer);
      resolve(data.text || '');
    };

    if (socket) {
      socket.once('captcha-answer', handler);
    } else {
      // Fallback: listen on all sockets
      io.once('connection', (s) => s.once('captcha-answer', handler));
    }
  });
}

// ========================
// SOCKET.IO
// ========================

io.on('connection', (socket) => {
  console.log(`Socket connected: ${socket.id}`);

  socket.on('disconnect', () => {
    console.log(`Socket disconnected: ${socket.id}`);
  });
});

// ========================
// START SERVER
// ========================

const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`\n🤖 UDISE+ Bot Server running on http://localhost:${PORT}`);
  console.log(`   Open this URL in your browser to start.\n`);
});
