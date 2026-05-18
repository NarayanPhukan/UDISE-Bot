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

// Utility: delay helper (replaces deprecated page.waitForTimeout)
const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

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
    // Use raw: false to ensure formatted percentages like "85%" are read as "85%" instead of 0.85
    const data = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });

    if (data.length === 0) {
      fs.unlinkSync(req.file.path);
      return res.status(400).json({ error: 'Excel file is empty' });
    }

    // Normalize column headers (case-insensitive matching)
    const normalizedData = data.map((row, index) => {
      const normalized = {};
      for (const [key, value] of Object.entries(row)) {
        const lowerKey = key.toLowerCase().trim();

        // Student Name: match "Name", "Student Name", "StudentName", etc.
        if (!normalized.studentName && (lowerKey === 'name' || (lowerKey.includes('student') && lowerKey.includes('name')))) {
          normalized.studentName = String(value).trim();
        }
        // Class
        else if (!normalized.class && (lowerKey.includes('class') || lowerKey === 'cls')) {
          normalized.class = String(value).trim();
        }
        // PEN Number
        else if (!normalized.penNo && lowerKey.includes('pen')) {
          normalized.penNo = String(value).trim();
        }
        // Attendance / No. of days school attended
        else if (lowerKey.includes('attend') || lowerKey.includes('present')) {
          // Absolute priority to columns that explicitly say 'attend' or 'present'
          normalized.attendance = String(value).split('.')[0].replace(/[^0-9]/g, '');
        }
        else if (!normalized.attendance && lowerKey.includes('day') && !lowerKey.includes('total') && !lowerKey.includes('working')) {
          // Fallback to 'day' only if it's not 'total days' or 'working days'
          normalized.attendance = String(value).split('.')[0].replace(/[^0-9]/g, '');
        }
        // Percentage
        else if (!normalized.percentage && (lowerKey.includes('percent') || lowerKey.includes('%'))) {
          normalized.percentage = String(value).trim().replace('%', '');
        }
        // Marks (Total Marks / Marks Obtained)
        else if (!normalized.marks && (lowerKey.includes('marks') || lowerKey.includes('total'))) {
          // Make sure it's not the percentage column
          if (!lowerKey.includes('%') && !lowerKey.includes('percent')) {
            normalized.marks = String(value).trim().replace('%', '');
          }
        }
        // Progression Status: match "Progression Status", but NOT "Schooling status"
        else if (!normalized.progressionStatus && (lowerKey.includes('progress') || lowerKey.includes('promotion') || (lowerKey.includes('status') && !lowerKey.includes('school')))) {
          normalized.progressionStatus = String(value).trim();
        }
        // Schooling status / Studying in Same School
        else if (!normalized.sameSchool && (lowerKey.includes('schooling') || (lowerKey.includes('same') && lowerKey.includes('school')))) {
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
  const { socketId, udiseCode, password, data, filePath, semiAutomatic } = req.body;

  if (!udiseCode || !password || !data || data.length === 0) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  const sessionId = Date.now().toString();
  activeSessions.set(sessionId, { status: 'starting', progress: 0 });

  // Start automation in background
  runAutomation(sessionId, socketId, udiseCode, password, data, filePath, semiAutomatic);

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

async function runAutomation(sessionId, socketId, udiseCode, password, students, filePath, semiAutomatic = false) {
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

    // ---- Step 1: Navigate to UDISE+ SDMS Login ----
    session.status = 'navigating';
    log('🌐 Navigating to UDISE+ SDMS login...');
    emit('progress', { step: 'navigating', percent: 5 });

    let loginSuccess = false;
    for (let attempts = 0; attempts < 3; attempts++) {
      try {
        await page.goto('https://sdms.udiseplus.gov.in/p2/v1/login?state-id=118', {
          waitUntil: attempts === 0 ? 'networkidle2' : 'domcontentloaded',
          timeout: 60000
        });
        loginSuccess = true;
        break;
      } catch (err) {
        log(`⚠️ Navigation timeout on attempt ${attempts + 1}/3... retrying in 3s`, 'warn');
        await delay(3000);
      }
    }
    
    if (!loginSuccess) {
      throw new Error('Failed to load login page after 3 attempts. The UDISE+ servers might be down or unreachable.');
    }

    // Wait for page to fully render
    log('⏳ Waiting for login page to render...');
    await delay(3000);

    // Wait for any input field to appear
    try {
      await page.waitForSelector('input', { timeout: 20000 });
      log('✅ Login page loaded');
    } catch (e) {
      log('⚠️ Login page may not have fully rendered, continuing...', 'warn');
    }

    await delay(2000);

    // ---- Step 2: Login ----
    session.status = 'logging_in';
    log('🔐 Attempting login...');
    emit('progress', { step: 'logging_in', percent: 10 });

    // Wait for any input field to appear on the page
    await page.waitForSelector('input', { timeout: 20000 })
      .catch(() => log('⚠️ No input fields found on page, trying alternatives...', 'warn'));

    // Extra wait for Angular form bindings to attach
    await delay(1000);

    // Log all available inputs for debugging
    const allInputs = await page.evaluate(() => {
      const inputs = Array.from(document.querySelectorAll('input'));
      return inputs.map(inp => ({
        type: inp.type,
        name: inp.name || '',
        id: inp.id || '',
        placeholder: inp.placeholder || '',
        formControlName: inp.getAttribute('formcontrolname') || inp.getAttribute('formControlName') || '',
        ngModel: inp.getAttribute('ng-model') || inp.getAttribute('[(ngModel)]') || '',
        className: inp.className || '',
        visible: inp.offsetParent !== null
      }));
    });
    log(`📋 Found ${allInputs.length} input fields on page`);
    allInputs.forEach((inp, i) => {
      log(`  Input ${i}: type=${inp.type} name="${inp.name}" id="${inp.id}" placeholder="${inp.placeholder}" formControlName="${inp.formControlName}" visible=${inp.visible}`);
    });

    // Try multiple selectors for the username field (broader set for Angular apps)
    const usernameSelectors = [
      'input[placeholder*="UDISE"]',
      'input[placeholder*="udise"]',
      'input[placeholder*="User"]',
      'input[placeholder*="user"]',
      'input[placeholder*="Enter"]',
      'input[formcontrolname*="user"]',
      'input[formcontrolname*="udise"]',
      'input[formcontrolname*="User"]',
      'input[formcontrolname*="login"]',
      'input[formcontrolname*="Login"]',
      'input[name*="user"]',
      'input[name*="udise"]',
      'input[name*="login"]',
      'input[id*="user"]',
      'input[id*="udise"]',
      'input[id*="login"]',
      'input[ng-model*="user"]',
      'input[ng-model*="udise"]',
      'input[type="text"]',
      'input[type="tel"]',
      'input[type="number"]',
      'input:not([type="password"]):not([type="hidden"]):not([type="submit"]):not([type="checkbox"]):not([type="radio"])'
    ];

    let usernameField = null;
    for (const sel of usernameSelectors) {
      usernameField = await page.$(sel);
      if (usernameField) {
        log(`✅ Found username field: ${sel}`);
        break;
      }
    }

    if (!usernameField) {
      // Take screenshot for debugging and send to frontend
      const debugScreenshot = await page.screenshot({ encoding: 'base64' });
      emit('captcha', { image: 'data:image/png;base64,' + debugScreenshot, debug: true });
      await page.screenshot({ path: path.join(__dirname, 'uploads', 'login-debug.png') });

      // Log page HTML for debugging
      const bodyHTML = await page.evaluate(() => document.body.innerHTML.substring(0, 2000));
      log(`📄 Page HTML (first 2000 chars): ${bodyHTML}`, 'info');

      throw new Error('Could not find the username input field on the login page. The page structure may have changed.');
    }

    await usernameField.click({ clickCount: 3 });
    // Press backspace to clear any existing content
    await usernameField.press('Backspace');
    await page.evaluate((el, val) => {
      el.value = '';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.value = val;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
    }, usernameField, udiseCode);
    await delay(300);

    // Enter password
    const passwordSelectors = [
      'input[type="password"]', 'input[placeholder*="Password"]',
      'input[name="password"]', 'input[id="password"]', '#password'
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
    await passwordField.press('Backspace');
    await page.evaluate((el, val) => {
      el.value = '';
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.value = val;
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
    }, passwordField, password);
    await delay(300);

    log('📝 Credentials entered');

    // Handle CAPTCHA if present
    const captchaExists = await page.$('canvas, img[src*="captcha"], .captcha, input[placeholder*="captcha"], input[placeholder*="Captcha"]');
    if (captchaExists) {
      log('⚠️ CAPTCHA detected! Sending screenshot to you...', 'warn');
      session.status = 'waiting_captcha';
      emit('progress', { step: 'waiting_captcha', percent: 12 });

      // Ensure captcha image is fully loaded before screenshot
      await delay(1000);

      // Take a cropped screenshot of the login box for a zoomed, highly readable preview, falling back to full page if needed
      let screenshotBuffer;
      try {
        const loginCard = await page.$('.login-box, .login-card, .login-container, .card, form, #login-form, .login-form');
        if (loginCard) {
          log('📷 Capturing zoomed login card for CAPTCHA preview...');
          screenshotBuffer = await loginCard.screenshot({ encoding: 'base64' });
        } else {
          log('📷 Falling back to full page screenshot...');
          screenshotBuffer = await page.screenshot({ encoding: 'base64' });
        }
      } catch (err) {
        log(`📷 Screenshot error: ${err.message}, taking full page...`, 'warn');
        screenshotBuffer = await page.screenshot({ encoding: 'base64' });
      }
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
        await captchaField.press('Backspace');
        await page.evaluate((el, val) => {
          el.value = '';
          el.dispatchEvent(new Event('input', { bubbles: true }));
          el.value = val;
          el.dispatchEvent(new Event('input', { bubbles: true }));
          el.dispatchEvent(new Event('change', { bubbles: true }));
          el.dispatchEvent(new Event('blur', { bubbles: true }));
        }, captchaField, captchaText);
        await delay(500);
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
    await delay(5000);

    // Re-check URL to handle slow navigations
    let currentUrl = page.url();
    log(`📍 Current URL: ${currentUrl}`);

    if (!currentUrl.includes('/home') && !currentUrl.includes('/school') && !currentUrl.includes('/academic-choice')) {
      // If we are still on the login page, check for actual error messages
      const errorElement = await page.$('.error-message, .alert-danger');
      if (errorElement) {
        try {
          const errorText = await page.evaluate(el => el.textContent, errorElement);
          if (errorText && errorText.trim()) {
            log(`❌ Login error: ${errorText.trim()}`, 'error');
            throw new Error(`Login failed: ${errorText.trim()}`);
          }
        } catch (err) {
          if (err.message.includes('Execution context was destroyed')) {
            log('⚠️ Page navigated while checking errors, proceeding...');
            await delay(3000);
            currentUrl = page.url();
          } else {
            throw err;
          }
        }
      }
      
      // Final check after potential slow navigation
      if (!currentUrl.includes('/home') && !currentUrl.includes('/school') && !currentUrl.includes('/academic-choice')) {
        currentUrl = page.url(); // One last check
        if (!currentUrl.includes('/home') && !currentUrl.includes('/school') && !currentUrl.includes('/academic-choice')) {
          throw new Error(`Login failed. Redirected to unexpected URL: ${currentUrl}`);
        }
      }
    }

    session.status = 'logged_in';
    log('✅ Login successful!');
    emit('progress', { step: 'logged_in', percent: 15 });

    // ---- Step 3: Navigate to Progression Module ----
    log('📂 Navigating to Progression Module...');
    emit('progress', { step: 'navigating_module', percent: 20 });

    // Extract school ID from the current URL after login
    const postLoginUrl = page.url();
    log(`📍 Post-login URL: ${postLoginUrl}`);
    let schoolId = null;
    const schoolMatch = postLoginUrl.match(/school\/(\d+)/);
    if (schoolMatch) {
      schoolId = schoolMatch[1];
      log(`🏫 School ID: ${schoolId}`);
    } else {
      // We might be on the /home page. Try to click the academic year card or find a link.
      log('🔍 Looking for school ID on the home page...');
      
      // Look for a link containing the school ID
      schoolId = await page.evaluate(() => {
        const links = document.querySelectorAll('a[href*="/school/"]');
        for (const link of links) {
          const match = link.href.match(/school\/(\d+)/);
          if (match) return match[1];
        }
        return null;
      });

      if (schoolId) {
        log(`🏫 School ID (from link): ${schoolId}`);
      } else {
        // Try clicking on a card to proceed
        const yearCards = await page.$$('a, div.card, button, .academic-year');
        for (const card of yearCards) {
          const text = await page.evaluate(el => el.textContent, card);
          if (text.includes('2026-27') || text.includes('2026 - 27') || text.includes('2025-26') || text.includes('2024-25') || text.includes('2023-24') || text.includes('Academic Year')) {
            await card.click();
            await delay(3000);
            break;
          }
        }

        const dashUrl = page.url();
        const dashMatch = dashUrl.match(/school\/(\d+)/);
        if (dashMatch) {
          schoolId = dashMatch[1];
          log(`🏫 School ID (after click): ${schoolId}`);
        } else {
           // Scrape it again just in case the DOM updated
           schoolId = await page.evaluate(() => {
            const links = document.querySelectorAll('a[href*="/school/"]');
            for (const link of links) {
              const match = link.href.match(/school\/(\d+)/);
              if (match) return match[1];
            }
            
            // Check session storage
            for (let i = 0; i < sessionStorage.length; i++) {
              const key = sessionStorage.key(i);
              const val = sessionStorage.getItem(key);
              if (val && val.includes('schoolId')) {
                try {
                  const parsed = JSON.parse(val);
                  if (parsed && parsed.schoolId) return parsed.schoolId.toString();
                } catch(e) {}
              }
            }
            return null;
          });

          if (schoolId) {
            log(`🏫 School ID (from link/storage after click): ${schoolId}`);
          } else {
            // Take a debug screenshot
            const errScreenshot = await page.screenshot({ encoding: 'base64' });
            emit('screenshot', { image: 'data:image/png;base64,' + errScreenshot, error: true });
            
            log('⚠️ Could not determine school ID dynamically. Falling back to known ID 3318905.', 'warn');
            schoolId = '3318905'; // Fallback to user's provided ID
          }
        }
      }
    }

    // Navigate to module-choice page
    const moduleChoiceUrl = `https://sdms.udiseplus.gov.in/g2/#/school/${schoolId}/module-choice`;
    log(`🔗 Navigating to: ${moduleChoiceUrl}`);
    await page.goto(moduleChoiceUrl, { waitUntil: 'networkidle2', timeout: 30000 });
    await delay(3000);

    // Click "Go" button on the Progression Module card
    log('🔘 Clicking "Go" on Progression Module...');
    const goButtons = await page.$$('button, a');
    let clickedGo = false;
    for (const btn of goButtons) {
      const text = await page.evaluate(el => el.textContent.trim(), btn);
      if (text.toLowerCase() === 'go') {
        await btn.click();
        clickedGo = true;
        log('✅ Clicked Go button');
        break;
      }
    }

    if (!clickedGo) {
      // Fallback: navigate directly to promotion page
      log('⚠️ Could not find Go button, navigating directly...', 'warn');
    }

    await delay(3000);
    emit('progress', { step: 'module_loaded', percent: 25 });

    // ---- Step 4: Process students by class ----
    session.status = 'processing';

    // Group students by class
    const studentsByClass = {};
    for (const student of students) {
      const cls = student.class || 'unknown';
      if (!studentsByClass[cls]) studentsByClass[cls] = [];
      studentsByClass[cls].push(student);
    }

    const classGroups = Object.keys(studentsByClass);
    log(`📋 Found ${classGroups.length} class group(s): ${classGroups.join(', ')}`);
    log(`📋 Total students to process: ${students.length}`);

    const results = {
      total: students.length,
      success: 0,
      failed: 0,
      skipped: 0,
      details: []
    };

    let processedCount = 0;

    for (const className of classGroups) {
      const classStudents = studentsByClass[className];

      // Check if session was stopped
      if (activeSessions.get(sessionId)?.status === 'stopping') {
        log('⏹️ Automation stopped by user', 'warn');
        break;
      }

      log(`\n📚 Processing Class: ${className} (${classStudents.length} students)`);

      // Navigate to the promotion page
      const promotionUrl = `https://sdms.udiseplus.gov.in/g2/#/school/${schoolId}/promotion`;
      await page.goto(promotionUrl, { waitUntil: 'networkidle2', timeout: 30000 });
      await delay(3000);

      // Select Class from dropdown — wait for dropdowns to load first
      log(`🔍 Selecting Class: ${className}...`);
      try {
        await page.waitForSelector('select', { timeout: 10000 });
        await delay(1000);
      } catch (e) {
        log('  ⚠️ Dropdowns slow to load, waiting extra...', 'warn');
        await delay(5000);
      }
      const classSelected = await selectDropdownOption(page, 'Select Class', className, log);
      if (!classSelected) {
        log(`❌ Could not select class "${className}", skipping this group`, 'error');
        for (const s of classStudents) {
          results.skipped++;
          processedCount++;
           results.details.push({
             row: s.rowIndex,
             name: s.studentName,
             pen: s.penNo,
             status: 'skipped',
             error: `Class "${className}" not found in dropdown`,
             percentage: s.percentage || '',
             progressionStatus: s.progressionStatus || ''
           });
        }
        continue;
      }
      await delay(1500);

      // Find available sections for this class
      const availableSections = await page.evaluate(() => {
        const selects = document.querySelectorAll('select');
        for (const select of selects) {
          const firstOpt = select.options[0]?.textContent.toLowerCase() || '';
          const label = select.closest('div')?.querySelector('label')?.textContent.toLowerCase() || '';
          if (firstOpt.includes('select section') || label.includes('select section')) {
            return Array.from(select.options)
              .filter(o => o.value && o.index > 0)
              .map(o => ({ value: o.value, text: o.textContent.trim() }));
          }
        }
        return [];
      });

      if (availableSections.length === 0) {
        log('⚠️ No sections found in dropdown, proceeding with default behavior...', 'warn');
        availableSections.push({ value: null, text: 'Default' });
      } else {
        log(`📋 Found ${availableSections.length} sections for Class ${className}: ${availableSections.map(s => s.text).join(', ')}`);
      }

      let remainingStudents = [...classStudents];

      // Helper: navigate to promotion page, select class and section, click Go
      async function loadSection(sectionText) {
        const promotionUrl = `https://sdms.udiseplus.gov.in/g2/#/school/${schoolId}/promotion`;
        const currentUrl = page.url();
        // Re-navigate if we're no longer on the promotion page
        if (!currentUrl.includes('/promotion')) {
          log('🔄 Re-navigating to promotion page...', 'warn');
          await page.goto(promotionUrl, { waitUntil: 'networkidle2', timeout: 30000 });
          await delay(3000);
        }
        // Wait for dropdowns
        try {
          await page.waitForSelector('select', { timeout: 10000 });
          await delay(1000);
        } catch (e) {
          // Dropdowns not found — force re-navigate
          log('🔄 Dropdowns not found, force re-navigating...', 'warn');
          await page.goto(promotionUrl, { waitUntil: 'networkidle2', timeout: 30000 });
          await delay(3000);
          await page.waitForSelector('select', { timeout: 15000 });
          await delay(1000);
        }
        // Select class
        await selectDropdownOption(page, 'Select Class', className, log);
        await delay(1500);
        // Select section
        if (sectionText) {
          await selectDropdownOption(page, 'Select Section', sectionText, log);
        } else {
          await selectDropdownOption(page, 'Select Section', null, log);
        }
        await delay(1000);
        // Click Go
        const goBtn = await findButtonByText(page, 'Go');
        if (goBtn) {
          await goBtn.click();
          await delay(5000);
          log(`✅ Student list loaded for Section ${sectionText}`);
          return true;
        } else {
          log('❌ Could not find Go button on promotion page', 'error');
          return false;
        }
      }

      for (const section of availableSections) {
        if (remainingStudents.length === 0) {
          log(`✅ All students for Class ${className} processed successfully.`);
          break;
        }

        // Load this section's student list
        log(`🔍 Selecting Section: ${section.text}...`);
        const sectionLoaded = await loadSection(section.text);
        if (!sectionLoaded) continue;

        // Take a screenshot of the loaded list for debugging
        const listScreenshot = await page.screenshot({ encoding: 'base64' });
        emit('screenshot', { image: 'data:image/png;base64,' + listScreenshot });

        // Keep track of students that fail due to "not found on the page"
        const notFoundStudents = [];

        // --- Read students from the page first, then match with Excel ---
        const pageText = await page.evaluate(() => document.body.innerText.toLowerCase());

        const studentsToProcessNow = [];
        for (const student of remainingStudents) {
          const nameVal = student.studentName ? String(student.studentName).trim().toLowerCase() : null;
          if (nameVal && pageText.includes(nameVal)) {
            studentsToProcessNow.push(student);
          }
        }
        
        log(`🔍 Found ${studentsToProcessNow.length} matching students on this page by name.`);

        // Identify which ones from remainingStudents were NOT on this page
        for (const student of remainingStudents) {
          if (!studentsToProcessNow.includes(student)) {
            notFoundStudents.push(student);
          }
        }

        if (studentsToProcessNow.length === 0) {
          log(`⚠️ None of the ${remainingStudents.length} remaining students were found on this page. Proceeding to next section.`, 'warn');
          remainingStudents = notFoundStudents;
          continue;
        }

        // Now process each matching student in serial order
        for (const student of studentsToProcessNow) {
          processedCount++;

          // Check if session was stopped
          if (activeSessions.get(sessionId)?.status === 'stopping') {
            log('⏹️ Automation stopped by user', 'warn');
            break;
          }

          const progressPercent = 25 + Math.round((processedCount / students.length) * 70);
          log(`\n📝 Processing ${processedCount}/${students.length}: ${student.studentName || 'Unknown'} (PEN: ${student.penNo || 'N/A'}) in Section ${section.text}`);
          emit('progress', {
            step: 'processing',
            percent: progressPercent,
            current: processedCount,
            total: students.length,
            studentName: student.studentName
          });

          try {
            const procResult = await processStudentOnPage(page, student, log, socket, semiAutomatic);
            results.success++;
            results.details.push({
              row: student.rowIndex,
              name: student.studentName,
              pen: student.penNo,
              status: 'success',
              section: section.text,
              percentage: procResult ? procResult.percentage : student.percentage,
              progressionStatus: procResult ? procResult.progressionStatus : student.progressionStatus
            });
            log(`✅ Successfully updated: ${student.studentName} in Section ${section.text}`, 'success');
          } catch (err) {
            if (err.message.includes('not found on the page') || err.message.includes('not found') || err.message.includes('not found in DOM')) {
               log(`⚠️ ${student.studentName} not found in Section ${section.text}, will retry in next section.`, 'warn');
               notFoundStudents.push(student);
               // Revert processedCount for next attempt
               processedCount--;
            } else if (err.message.includes('context was destroyed') || err.message.includes('navigation') || err.message.includes('detached') || err.message.includes('Target closed')) {
               // Page navigated away — recover and retry
               log(`⚠️ Page navigated away during ${student.studentName}, recovering...`, 'warn');
               processedCount--;
               const recovered = await loadSection(section.text);
               if (recovered) {
                 try {
                   const procResult = await processStudentOnPage(page, student, log, socket, semiAutomatic);
                   results.success++;
                   results.details.push({
                     row: student.rowIndex,
                     name: student.studentName,
                     pen: student.penNo,
                     status: 'success',
                     section: section.text,
                     percentage: procResult ? procResult.percentage : student.percentage,
                     progressionStatus: procResult ? procResult.progressionStatus : student.progressionStatus
                   });
                   log(`✅ Successfully updated (after recovery): ${student.studentName} in Section ${section.text}`, 'success');
                 } catch (retryErr) {
                   if (retryErr.message.includes('not found on the page') || retryErr.message.includes('not found') || retryErr.message.includes('not found in DOM')) {
                     log(`⚠️ ${student.studentName} not found after recovery, will retry in next section.`, 'warn');
                     notFoundStudents.push(student);
                   } else {
                     results.failed++;
                     results.details.push({
                       row: student.rowIndex,
                       name: student.studentName,
                       pen: student.penNo,
                       status: 'failed',
                       error: retryErr.message,
                       section: section.text,
                       percentage: student.percentage || '',
                       progressionStatus: student.progressionStatus || ''
                     });
                     log(`❌ Failed for ${student.studentName} after recovery: ${retryErr.message}`, 'error');
                   }
                 }
               } else {
                 results.failed++;
                 results.details.push({
                   row: student.rowIndex,
                   name: student.studentName,
                   pen: student.penNo,
                   status: 'failed',
                   error: 'Could not recover page',
                   section: section.text,
                   percentage: student.percentage || '',
                   progressionStatus: student.progressionStatus || ''
                 });
                 log(`❌ Failed for ${student.studentName}: Could not recover page`, 'error');
               }
            } else {
               results.failed++;
               results.details.push({
                 row: student.rowIndex,
                 name: student.studentName,
                 pen: student.penNo,
                 status: 'failed',
                 error: err.message,
                 section: section.text,
                 percentage: student.percentage || '',
                 progressionStatus: student.progressionStatus || ''
               });
               log(`❌ Failed for ${student.studentName}: ${err.message}`, 'error');

               // Take screenshot on failure
               try {
                 const errScreenshot = await page.screenshot({ encoding: 'base64' });
                 emit('screenshot', { image: 'data:image/png;base64,' + errScreenshot, error: true });
               } catch (e) { /* ignore */ }
            }
          }

          await delay(1500);
        }

        // Check if stopped
        if (activeSessions.get(sessionId)?.status === 'stopping') {
          break;
        }

        remainingStudents = notFoundStudents;
      }

      // If there are still remaining students after ALL sections have been checked, mark them as failed
      for (const student of remainingStudents) {
         processedCount++;
         results.failed++;
         results.details.push({
            row: student.rowIndex,
            name: student.studentName,
            pen: student.penNo,
            status: 'failed',
            error: `Student not found in any section of Class ${className}`,
            percentage: student.percentage || '',
            progressionStatus: student.progressionStatus || ''
         });
         log(`❌ Failed for ${student.studentName}: Student not found in any section of Class ${className}`, 'error');
         
         const progressPercent = 25 + Math.round((processedCount / students.length) * 70);
         emit('progress', {
            step: 'processing',
            percent: progressPercent,
            current: processedCount,
            total: students.length,
            studentName: student.studentName
         });
      }
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

// ========================
// SDMS PROGRESSION HELPERS
// ========================

/**
 * Find a button on the page by its text content
 */
async function findButtonByText(page, text) {
  const buttons = await page.$$('button');
  for (const btn of buttons) {
    const btnText = await page.evaluate(el => el.textContent.trim(), btn);
    if (btnText.toLowerCase() === text.toLowerCase()) {
      return btn;
    }
  }
  return null;
}

/**
 * Select an option from a dropdown by finding the <select> near a label.
 * If targetValue is null, selects the first non-placeholder option.
 */
async function selectDropdownOption(page, labelText, targetValue, log) {
  // Find all select elements on the page
  const selects = await page.$$('select');

  for (const select of selects) {
    // Check if this select has the matching placeholder/label
    const selectText = await page.evaluate(el => {
      // Check for a matching option text (like "Select Class")
      const firstOption = el.options[0];
      const label = el.closest('.form-group, .form-row, div')?.querySelector('label');
      return {
        firstOptionText: firstOption ? firstOption.textContent.trim() : '',
        labelText: label ? label.textContent.trim() : '',
        id: el.id || '',
        name: el.name || ''
      };
    }, select);

    const matchesLabel = selectText.firstOptionText.toLowerCase().includes(labelText.toLowerCase()) ||
                          selectText.labelText.toLowerCase().includes(labelText.toLowerCase());

    if (!matchesLabel) continue;

    // Get all options
    const options = await page.evaluate(el => {
      return Array.from(el.options).map((opt, i) => ({
        value: opt.value,
        text: opt.textContent.trim(),
        index: i
      }));
    }, select);

    log(`  📋 Dropdown "${labelText}" options: ${options.map(o => o.text).join(', ')}`);

    if (targetValue) {
      // Find matching option by class name
      // The Excel has class values like "I", "II", etc. The dropdown may have "Class I", "I", "1", etc.
      const classValue = targetValue.toString().trim().toUpperCase();

      for (const opt of options) {
        const optText = opt.text.trim().toUpperCase();
        // Try exact match first, then partial
        if (optText === classValue ||
            optText === `CLASS ${classValue}` ||
            optText === `CLASS-${classValue}` ||
            optText.includes(classValue) ||
            // Handle Roman numeral equivalents
            (classValue === 'I' && (optText === '1' || optText.includes('CLASS 1') || optText.includes('CLASS I'))) ||
            (classValue === 'II' && (optText === '2' || optText.includes('CLASS 2') || optText.includes('CLASS II'))) ||
            (classValue === 'III' && (optText === '3' || optText.includes('CLASS 3') || optText.includes('CLASS III')))) {
          await page.evaluate((el, val) => {
            el.value = val;
            el.dispatchEvent(new Event('change', { bubbles: true }));
          }, select, opt.value);
          log(`  ✅ Selected "${opt.text}" for ${labelText}`);
          return true;
        }
      }

      log(`  ⚠️ Could not find "${targetValue}" in ${labelText} dropdown`, 'warn');
      return false;
    } else {
      // Select the first non-placeholder option
      const validOption = options.find(o => o.index > 0 && o.value);
      if (validOption) {
        await page.evaluate((el, val) => {
          el.value = val;
          el.dispatchEvent(new Event('change', { bubbles: true }));
        }, select, validOption.value);
        log(`  ✅ Selected "${validOption.text}" for ${labelText}`);
        return true;
      }
    }
  }

  log(`  ⚠️ Could not find dropdown for "${labelText}"`, 'warn');
  return false;
}

/**
 * Process a single student on the loaded progression page.
 * The page shows all students in the class — find this student's row by PEN,
 * fill in the form fields within that row, and click Update.
 */
async function processStudentOnPage(page, student, log, socket, semiAutomatic = false) {
  const studentName = student.studentName ? String(student.studentName).trim() : '';
  if (!studentName) throw new Error('No name provided to identify this student');
  
  const rawPen = student.penNo ? String(student.penNo).trim() : '';
  const pen = rawPen ? rawPen : studentName.replace(/[^a-zA-Z0-9]/g, '');

  const elementInfo = await page.evaluate((pen, name, progressionStatus, sameSchool) => {
    let penElement = null;
    if (name) {
      const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
      const searchName = name.toLowerCase();
      while (walker.nextNode()) {
        const nodeText = walker.currentNode.textContent.toLowerCase();
        if (nodeText.trim() === searchName || (nodeText.includes(searchName) && nodeText.length < 100)) {
          const el = walker.currentNode.parentElement;
          const style = window.getComputedStyle(el);
          const rect = el.getBoundingClientRect();
          if (style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0) {
            penElement = el;
            break;
          }
        }
      }
    }

    if (!penElement) return { success: false, error: 'Student element not found in DOM' };

    let studentContainer = penElement.closest('tr');
    if (!studentContainer) {
      studentContainer = penElement.closest('.row, .student-block, [class*="record"]');
    }
    if (!studentContainer) {
      studentContainer = penElement;
      let attempts = 0;
      while (studentContainer && attempts < 15 && studentContainer.tagName !== 'BODY') {
        const selects = studentContainer.querySelectorAll('select');
        const inputs = studentContainer.querySelectorAll('input[type="text"], input[type="number"], input:not([type])');
        // Stop if we find a reasonable number of inputs for ONE student (not the whole page)
        if (selects.length >= 1 && inputs.length >= 1 && inputs.length <= 8) {
          break;
        }
        studentContainer = studentContainer.parentElement;
        attempts++;
      }
    }
    
    if (!studentContainer || studentContainer.tagName === 'BODY') {
      return { success: false, error: 'Could not find student form container' };
    }

    studentContainer.id = 'bot-student-container-' + pen;

    // Now identify the fields within this container
    const selects = studentContainer.querySelectorAll('select');
    const inputs = studentContainer.querySelectorAll('input[type="text"], input[type="number"], input:not([type="hidden"]):not([type="checkbox"]):not([type="radio"]):not([type="submit"]):not([type="password"])');

    // Filter inputs to only those that look like data fields (not search fields etc.)
    const dataInputs = Array.from(inputs).filter(inp => {
      const rect = inp.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0; // Only visible inputs
    });

    const info = { success: true, selectsCount: selects.length, inputsCount: dataInputs.length, filled: [] };

    // Tag elements so Puppeteer can interact with them natively, using PEN to ensure uniqueness
    if (selects.length > 0) selects[0].id = `bot-progress-select-${pen}`;
    if (selects.length > 1) selects[selects.length - 1].id = `bot-school-select-${pen}`;

    let marksInput = null, percentInput = null, daysInput = null;
    info.inputDebug = [];
    
    // First, try matching by attributes (formControlName, placeholder, name, id) and nearest label
    dataInputs.forEach((inp, idx) => {
      // Get label text from closest label element only (not container text which is too broad)
      let labelText = '';
      // Walk up to find closest label - check siblings and parent containers
      let parent = inp.parentElement;
      for (let climb = 0; climb < 5 && parent; climb++) {
        const labels = parent.querySelectorAll('label, span.label, .col-form-label, th');
        for (const lbl of labels) {
          const t = lbl.textContent.trim().toLowerCase();
          if (t.length > 2 && t.length < 80) {
            labelText = t;
            break;
          }
        }
        if (labelText) break;
        parent = parent.parentElement;
      }

      const attrParts = [
        inp.name || '',
        inp.getAttribute('formcontrolname') || '',
        inp.placeholder || '',
        inp.id || ''
      ];
      const attrStr = attrParts.join(' ').toLowerCase();
      const fullMatch = (attrStr + ' ' + labelText).toLowerCase();
      
      info.inputDebug.push({
        index: idx,
        name: inp.name,
        formControlName: inp.getAttribute('formcontrolname'),
        placeholder: inp.placeholder,
        id: inp.id,
        labelText: labelText,
        currentValue: inp.value,
        fullMatch: fullMatch
      });

      // Match by priority: explicit attribute first, then label text
      if (fullMatch.includes('attend') || (fullMatch.includes('day') && !fullMatch.includes('total') && !fullMatch.includes('working'))) {
        daysInput = inp;
      } else if (fullMatch.includes('percent') || fullMatch.includes('%') || fullMatch.includes('pcnt')) {
        percentInput = inp;
      } else if (fullMatch.includes('mark') || fullMatch.includes('total') || fullMatch.includes('obtain')) {
        marksInput = inp;
      }
    });

    // Fallback to visual order if attributes didn't match everything
    if (dataInputs.length === 3) {
      if (!marksInput) marksInput = dataInputs[0];
      if (!percentInput) percentInput = dataInputs[1];
      if (!daysInput) daysInput = dataInputs[2];
    } else if (dataInputs.length === 2) {
      if (!percentInput) percentInput = dataInputs[0];
      if (!daysInput) daysInput = dataInputs[1];
    } else if (dataInputs.length > 0) {
      if (!percentInput) percentInput = dataInputs[0];
    }

    if (marksInput) { marksInput.id = `bot-marks-input-${pen}`; info.hasMarksInput = true; }
    if (percentInput) { percentInput.id = `bot-percent-input-${pen}`; info.hasPercentInput = true; }
    if (daysInput) { daysInput.id = `bot-days-input-${pen}`; info.hasDaysInput = true; }

    // Tag the update and correction buttons
    const btnsInContainer = Array.from(studentContainer.querySelectorAll('button, a.btn, input[type="button"]'));
    const updateBtn = btnsInContainer.find(b => b.textContent.trim().toLowerCase() === 'update');
    if (updateBtn) updateBtn.id = `bot-update-btn-${pen}`;
    info.hasUpdateBtn = !!updateBtn;

    const correctionBtn = btnsInContainer.find(b => b.textContent.trim().toLowerCase() === 'correction');
    if (correctionBtn) correctionBtn.id = `bot-correction-btn-${pen}`;
    info.hasCorrectionBtn = !!correctionBtn;

    // Determine values to select
    if (progressionStatus && selects.length > 0) {
      const progressSelect = selects[0];
      const progressValue = progressionStatus.toLowerCase();
      for (const opt of progressSelect.options) {
        const optText = opt.textContent.trim().toLowerCase();
        if (progressValue.includes('promot') && (optText.includes('promot') || optText.includes('pass'))) {
          info.progressValue = opt.value;
          info.filled.push('Progression Status: ' + opt.textContent.trim());
          break;
        }
        if (optText.includes(progressValue.substring(0, 10)) || progressValue.includes(optText.substring(0, 10))) {
          info.progressValue = opt.value;
          info.filled.push('Progression Status: ' + opt.textContent.trim());
          break;
        }
      }
    }

    if (sameSchool && selects.length > 1) {
      const schoolSelect = selects[selects.length - 1];
      const schoolValue = sameSchool.toLowerCase();
      for (const opt of schoolSelect.options) {
        const optText = opt.textContent.trim().toLowerCase();
        if (optText.includes('studying') && schoolValue.includes('studying')) {
          info.schoolValue = opt.value;
          info.filled.push('Schooling Status: ' + opt.textContent.trim());
          break;
        }
        if (optText.includes('same') && schoolValue.includes('same')) {
          info.schoolValue = opt.value;
          info.filled.push('Schooling Status: ' + opt.textContent.trim());
          break;
        }
        if (optText.length > 3 && schoolValue.includes(optText.substring(0, 8))) {
          info.schoolValue = opt.value;
          info.filled.push('Schooling Status: ' + opt.textContent.trim());
          break;
        }
      }
    }

    return info;
  }, pen, studentName, student.progressionStatus || '', student.sameSchool || '');

  if (!elementInfo.success) {
    throw new Error(elementInfo.error || 'Failed to locate student form elements');
  }

  log(`  📊 Form fields: ${elementInfo.selectsCount} dropdowns, ${elementInfo.inputsCount} inputs`);
  
  // Log debug info about each input field detected
  if (elementInfo.inputDebug) {
    for (const d of elementInfo.inputDebug) {
      log(`  🔎 Input[${d.index}]: formControlName="${d.formControlName}" name="${d.name}" placeholder="${d.placeholder}" label="${d.labelText}" currentVal="${d.currentValue}"`);
    }
  }

  // Capture a screenshot of the student container or the whole viewport
  let screenshot = null;
  try {
    const container = await page.$(`#bot-student-container-${pen}`);
    if (container) {
      screenshot = 'data:image/png;base64,' + (await container.screenshot({ encoding: 'base64' }));
    } else {
      screenshot = 'data:image/png;base64,' + (await page.screenshot({ encoding: 'base64' }));
    }
  } catch (e) {
    log(`  ⚠️ Could not take element screenshot: ${e.message}`, 'warn');
  }

  let finalMarks = student.marks ? student.marks.toString() : '';
  let finalPercent = student.percentage ? Math.round(parseFloat(student.percentage)).toString() : '';
  let finalDays = student.attendance ? student.attendance.toString() : '155';
  let finalProgression = student.progressionStatus || 'Promoted';

  // Performance fallback: if portal has percent input but no marks input, and we only have marks, use marks as percentage!
  if (elementInfo.hasPercentInput && !elementInfo.hasMarksInput && !finalPercent && finalMarks) {
    finalPercent = finalMarks;
  }
  // And vice-versa: if portal has marks input but no percent input, and we only have percentage, use percentage as marks!
  if (elementInfo.hasMarksInput && !elementInfo.hasPercentInput && !finalMarks && finalPercent) {
    finalMarks = finalPercent;
  }

  // Helper to re-tag elements because Angular destroys and re-renders the row upon Correction click
  async function reTagElements() {
    log('  🔄 Re-tagging elements after Correction unlock...');
    const reTagInfo = await page.evaluate((pen, name) => {
      let penElement = null;
      if (name) {
        const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
        const searchName = name.toLowerCase();
        while (walker.nextNode()) {
          const nodeText = walker.currentNode.textContent.toLowerCase();
          if (nodeText.trim() === searchName || (nodeText.includes(searchName) && nodeText.length < 100)) {
            const el = walker.currentNode.parentElement;
            const style = window.getComputedStyle(el);
            const rect = el.getBoundingClientRect();
            if (style.display !== 'none' && style.visibility !== 'hidden' && rect.width > 0 && rect.height > 0) {
              penElement = el;
              break;
            }
          }
        }
      }
      if (!penElement) return { success: false, error: 'Student name text not found after Correction re-render' };

      let studentContainer = penElement.closest('tr');
      if (!studentContainer) {
        studentContainer = penElement.closest('.row, .student-block, [class*="record"]');
      }
      if (!studentContainer) {
        studentContainer = penElement;
        let attempts = 0;
        while (studentContainer && attempts < 15 && studentContainer.tagName !== 'BODY') {
          const selects = studentContainer.querySelectorAll('select');
          const inputs = studentContainer.querySelectorAll('input[type="text"], input[type="number"], input:not([type])');
          if (selects.length >= 1 && inputs.length >= 1 && inputs.length <= 8) {
            break;
          }
          studentContainer = studentContainer.parentElement;
          attempts++;
        }
      }
      if (!studentContainer || studentContainer.tagName === 'BODY') return { success: false, error: 'Student form container not found after Correction re-render' };

      studentContainer.id = 'bot-student-container-' + pen;

      const selects = studentContainer.querySelectorAll('select');
      const inputs = studentContainer.querySelectorAll('input[type="text"], input[type="number"], input:not([type="hidden"]):not([type="checkbox"]):not([type="radio"]):not([type="submit"]):not([type="password"])');
      const dataInputs = Array.from(inputs).filter(inp => {
        const rect = inp.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      });

      if (selects.length > 0) selects[0].id = `bot-progress-select-${pen}`;
      if (selects.length > 1) selects[selects.length - 1].id = `bot-school-select-${pen}`;

      let marksInput = null, percentInput = null, daysInput = null;
      dataInputs.forEach((inp, idx) => {
        let labelText = '';
        let parent = inp.parentElement;
        for (let climb = 0; climb < 5 && parent; climb++) {
          const labels = parent.querySelectorAll('label, span.label, .col-form-label, th');
          for (const lbl of labels) {
            const t = lbl.textContent.trim().toLowerCase();
            if (t.length > 2 && t.length < 80) {
              labelText = t;
              break;
            }
          }
          if (labelText) break;
          parent = parent.parentElement;
        }
        const attrParts = [
          inp.name || '',
          inp.getAttribute('formcontrolname') || '',
          inp.placeholder || '',
          inp.id || ''
        ];
        const attrStr = attrParts.join(' ').toLowerCase();
        const fullMatch = (attrStr + ' ' + labelText).toLowerCase();
        if (fullMatch.includes('attend') || (fullMatch.includes('day') && !fullMatch.includes('total') && !fullMatch.includes('working'))) {
          daysInput = inp;
        } else if (fullMatch.includes('percent') || fullMatch.includes('%') || fullMatch.includes('pcnt')) {
          percentInput = inp;
        } else if (fullMatch.includes('mark') || fullMatch.includes('total') || fullMatch.includes('obtain')) {
          marksInput = inp;
        }
      });

      if (dataInputs.length === 3) {
        if (!marksInput) marksInput = dataInputs[0];
        if (!percentInput) percentInput = dataInputs[1];
        if (!daysInput) daysInput = dataInputs[2];
      } else if (dataInputs.length === 2) {
        if (!percentInput) percentInput = dataInputs[0];
        if (!daysInput) daysInput = dataInputs[1];
      } else if (dataInputs.length > 0) {
        if (!percentInput) percentInput = dataInputs[0];
      }

      if (marksInput) marksInput.id = `bot-marks-input-${pen}`;
      if (percentInput) percentInput.id = `bot-percent-input-${pen}`;
      if (daysInput) daysInput.id = `bot-days-input-${pen}`;

      const btnsInContainer = Array.from(studentContainer.querySelectorAll('button, a.btn, input[type="button"]'));
      const updateBtn = btnsInContainer.find(b => b.textContent.trim().toLowerCase() === 'update');
      if (updateBtn) updateBtn.id = `bot-update-btn-${pen}`;

      return {
        success: true,
        hasMarksInput: !!marksInput,
        hasPercentInput: !!percentInput,
        hasDaysInput: !!daysInput,
        hasUpdateBtn: !!updateBtn
      };
    }, pen, studentName);

    if (reTagInfo.success) {
      log('  ✅ Successfully re-tagged elements after Correction');
      elementInfo.hasMarksInput = reTagInfo.hasMarksInput;
      elementInfo.hasPercentInput = reTagInfo.hasPercentInput;
      elementInfo.hasDaysInput = reTagInfo.hasDaysInput;
      elementInfo.hasUpdateBtn = reTagInfo.hasUpdateBtn;
    } else {
      log(`  ⚠️ Re-tagging failed: ${reTagInfo.error}`, 'warn');
    }
  }

  if (semiAutomatic || elementInfo.hasCorrectionBtn) {
    let confirmData = {
      studentName: studentName,
      class: student.class || '',
      penNo: student.penNo || '',
      marks: finalMarks,
      percentage: finalPercent,
      attendance: finalDays,
      progressionStatus: student.progressionStatus || 'Promoted',
      sameSchool: student.sameSchool || 'Studying in Same School',
      screenshot: screenshot,
      alreadyEntered: !!elementInfo.hasCorrectionBtn
    };

    log(`⏳ Waiting for user confirmation for ${studentName}...`);
    socket.emit('confirm-student', confirmData);
    
    const response = await waitForStudentConfirmation(socket);
    log(`👤 Confirmation response: action=${response.action}`);
    
    if (response.action === 'skip') {
      throw new Error('Skipped by user');
    }
    
    finalMarks = response.marks;
    finalPercent = response.percentage;
    finalDays = response.attendance;
    finalProgression = response.progressionStatus;

    // Click correction button if student is already processed to unlock the form!
    if (elementInfo.hasCorrectionBtn) {
      log('  🛠️ Student already processed (status is Done). Clicking "Correction" button to unlock form...');
      await page.click(`#bot-correction-btn-${pen}`);
      await delay(1000);

      // Handle any standard sweetalert / bootstrap confirm modal for Correction click
      try {
        const confirmBtn = await page.evaluateHandle(() => {
          const btns = Array.from(document.querySelectorAll('button, .swal2-confirm, .btn-primary'));
          return btns.find(b => {
            if (b.offsetParent === null) return false; // Must be visible
            const text = b.textContent.trim().toLowerCase();
            return text.includes('yes') || text.includes('ok') || text.includes('confirm') || text.includes('sure');
          }) || null;
        });
        if (confirmBtn) {
          await confirmBtn.click();
          log('  ✅ Confirmed correction dialog popup');
          await delay(2000);
        }
      } catch (e) { /* no confirm dialog */ }

      // CRITICAL: Re-tag elements because Angular destroys and re-renders the row upon Correction click
      await reTagElements();
    }
    
    // Select dropdown options using confirmed choices
    if (elementInfo.selectsCount > 0) {
      await selectRowDropdownOption(page, `bot-progress-select-${pen}`, response.progressionStatus);
      await delay(300);
    }
    if (elementInfo.selectsCount > 1) {
      await selectRowDropdownOption(page, `bot-school-select-${pen}`, response.sameSchool);
      await delay(300);
    }
  } else {
    // Automatic Mode dropdown selection
    
    // Click correction button if student is already processed to unlock the form!
    if (elementInfo.hasCorrectionBtn) {
      log('  🛠️ Student already processed (status is Done). Clicking "Correction" button to unlock form...');
      await page.click(`#bot-correction-btn-${pen}`);
      await delay(1000);

      // Handle any standard sweetalert / bootstrap confirm modal for Correction click
      try {
        const confirmBtn = await page.evaluateHandle(() => {
          const btns = Array.from(document.querySelectorAll('button, .swal2-confirm, .btn-primary'));
          return btns.find(b => {
            if (b.offsetParent === null) return false; // Must be visible
            const text = b.textContent.trim().toLowerCase();
            return text.includes('yes') || text.includes('ok') || text.includes('confirm') || text.includes('sure');
          }) || null;
        });
        if (confirmBtn) {
          await confirmBtn.click();
          log('  ✅ Confirmed correction dialog popup');
          await delay(2000);
        }
      } catch (e) { /* no confirm dialog */ }

      // CRITICAL: Re-tag elements because Angular destroys and re-renders the row upon Correction click
      await reTagElements();
    }

    if (elementInfo.progressValue) {
      await page.select(`#bot-progress-select-${pen}`, elementInfo.progressValue);
      await delay(300);
    }
    if (elementInfo.schoolValue) {
      await page.select(`#bot-school-select-${pen}`, elementInfo.schoolValue);
      await delay(300);
    }
  }

  // Helper: Set input value
  async function setInputValue(selector, value, fieldName) {
    const el = await page.$(selector);
    if (!el) {
      log(`  ⚠️ Could not find element ${selector} for ${fieldName}`, 'warn');
      return false;
    }

    // Check element state and remove readonly/disabled if present
    const elState = await page.evaluate(sel => {
      const e = document.querySelector(sel);
      if (!e) return { found: false };
      const state = {
        found: true,
        type: e.type,
        readOnly: e.readOnly,
        disabled: e.disabled,
        value: e.value
      };
      e.readOnly = false;
      e.disabled = false;
      e.removeAttribute('readonly');
      e.removeAttribute('disabled');
      return state;
    }, selector);

    if (!elState.found) return false;
    log(`  📝 ${fieldName}: type=${elState.type} readonly=${elState.readOnly} disabled=${elState.disabled} current="${elState.value}" → setting "${value}"`);

    try {
      await page.focus(selector);
      await delay(100);
      
      // Select all text using triple click (highly effective in all input types, including type="number")
      await page.click(selector, { clickCount: 3 });
      await delay(100);
      
      // Also try standard keyboard shortcut (Ctrl+A) to make absolutely sure all text is highlighted
      await page.keyboard.down('Control');
      await page.keyboard.press('a');
      await page.keyboard.up('Control');
      await delay(100);
      
      // Delete the selection
      await page.keyboard.press('Backspace');
      await delay(150);
      
      // Double check if the field is cleared. If not (due to type="number" cursor issues in Chrome),
      // we clear it programmatically and dispatch input/change events to update Angular state
      const isCleared = await page.evaluate(sel => {
        const e = document.querySelector(sel);
        return e ? e.value === '' : true;
      }, selector);
      
      if (!isCleared) {
        await page.evaluate(sel => {
          const e = document.querySelector(sel);
          if (e) {
            e.value = '';
            e.dispatchEvent(new Event('input', { bubbles: true }));
            e.dispatchEvent(new Event('change', { bubbles: true }));
          }
        }, selector);
        await delay(100);
      }
      
      // Type the new value natively so Angular form control validation and dirty checking is triggered
      await page.type(selector, value, { delay: 50 });
      await delay(100);
      
      await page.keyboard.press('Tab');
      await delay(200);
      
      const actual = await page.evaluate(sel => document.querySelector(sel).value, selector);
      if (actual !== value) {
        log(`  ❌ Puppeteer type failed for ${fieldName}: expected "${value}" got "${actual}"`, 'error');
      }
    } catch (err) {
      log(`  ❌ Error typing into ${fieldName}: ${err.message}`, 'error');
    }

    return true;
  }

  if (finalMarks && elementInfo.hasMarksInput) {
    await setInputValue(`#bot-marks-input-${pen}`, finalMarks, 'Marks');
    await delay(300);
  }

  if (finalPercent && elementInfo.hasPercentInput) {
    await setInputValue(`#bot-percent-input-${pen}`, finalPercent, 'Percentage');
    await delay(300);
  }

  if (elementInfo.hasDaysInput) {
    await setInputValue(`#bot-days-input-${pen}`, finalDays, 'Days Attended');
    await delay(300);
  }

  // Pre-submit readback check
  const preSubmitValues = await page.evaluate((pen) => {
    const pctEl = document.querySelector(`#bot-percent-input-${pen}`);
    const daysEl = document.querySelector(`#bot-days-input-${pen}`);
    return {
      pct: pctEl ? pctEl.value : 'NOT_FOUND',
      days: daysEl ? daysEl.value : 'NOT_FOUND'
    };
  }, pen);
  log(`  🔍 PRE-SUBMIT CHECK: Percentage="${preSubmitValues.pct}" Days="${preSubmitValues.days}"`);

  if (!elementInfo.hasUpdateBtn) {
    throw new Error('Could not find Update button for this student');
  }

  log('  💾 Clicking Update button...');
  await page.click(`#bot-update-btn-${pen}`);
  await delay(2000);

  // Check for success confirmation or alert
  const alertText = await page.evaluate(() => {
    const alerts = document.querySelectorAll('.alert, .toast, .modal-body, .swal2-popup, [role="alert"]');
    for (const alert of alerts) {
      if (alert.offsetParent !== null && alert.textContent.trim()) {
        return alert.textContent.trim();
      }
    }
    return null;
  });

  if (alertText) {
    log(`  📢 Response: ${alertText.substring(0, 100)}`);
  }

  // Handle any confirmation dialogs (like OK button on success alert)
  try {
    const okBtn = await page.evaluateHandle(() => {
      const btns = Array.from(document.querySelectorAll('button'));
      return btns.find(b => {
        const text = b.textContent.trim().toLowerCase();
        return text === 'ok' || text === 'close' || text === 'yes';
      }) || null;
    });
    if (okBtn) {
      await okBtn.click();
      await delay(1000);
    }
  } catch (e) { /* no dialog to close */ }
  
  return {
    percentage: finalPercent,
    progressionStatus: finalProgression
  };
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
// STUDENT CONFIRMATION HELPERS
// ========================

function waitForStudentConfirmation(socket, timeoutMs = 300000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      reject(new Error('Student confirmation timeout — no response within 5 minutes'));
    }, timeoutMs);

    const handler = (data) => {
      clearTimeout(timer);
      resolve(data);
    };

    if (socket) {
      socket.once('confirm-student-response', handler);
    } else {
      io.once('connection', (s) => s.once('confirm-student-response', handler));
    }
  });
}

async function selectRowDropdownOption(page, selectId, targetText) {
  await page.evaluate((selId, text) => {
    const select = document.getElementById(selId);
    if (!select) return;
    const lowerTarget = text.toLowerCase();
    for (const opt of select.options) {
      const optText = opt.textContent.trim().toLowerCase();
      if (optText.includes(lowerTarget) || lowerTarget.includes(optText)) {
        select.value = opt.value;
        select.dispatchEvent(new Event('change', { bubbles: true }));
        break;
      }
    }
  }, selectId, targetText);
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
