
/* ────────────────────────────────────────────────────────────────
   CONSTANTS & SHIFT CONFIG
   ──────────────────────────────────────────────────────────────── */
const SHIFT_DEFAULT = {
  start: '09:00',   // Helix Industries shift start
  end:   '18:30',   // Shift end
  lunchMins: 30,    // Mandatory lunch deduction
  expectedHrs: 9    // Net working hours expected per day
};

/* ════════════════════════════════════════════════════════════════════
   GOOGLE SHEETS API LAYER
   All data reads/writes go through GSheet.call() which hits the
   Apps Script web endpoint. Results are also cached in localStorage
   so the UI stays fast and works briefly if connection hiccups.
   ════════════════════════════════════════════════════════════════════ */

// ── Configuration — Admin sets this URL in Settings ──────────────────
// Stored in localStorage under 'helix_apiUrl'
// Falls back to localStorage-only mode if not configured (for testing)

const Cache = {
  get(k, def=null)  { try { return JSON.parse(localStorage.getItem('helix_'+k)) ?? def; } catch { return def; } },
  set(k, v)        { localStorage.setItem('helix_'+k, JSON.stringify(v)); },
  del(k)           { localStorage.removeItem('helix_'+k); }
};

// Store === Cache (same object, defined below)

/* ════════════════════════════════════════════════════════════════════
   REMOTE CONFIG — fetched from config.json on every app load.
   Hosted alongside index.html on Netlify.
   Change config.json once → all devices pick it up automatically.
   ════════════════════════════════════════════════════════════════════ */

/** Holds the parsed config.json contents once loaded */
let APP_CONFIG = null;

/**
 * loadRemoteConfig — fetches config.json from the same server as index.html.
 * Uses cache-busting so Netlify always serves the latest version.
 * Falls back gracefully if the file is missing or unreachable.
 */
async function loadRemoteConfig() {
  try {
    // ?v= cache-bust ensures fresh fetch, not a 304 from browser cache
    const resp = await fetch('./config.json?v=' + Date.now(), {
      cache: 'no-store'
    });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    APP_CONFIG = await resp.json();

    // ── Apply config.json values into local cache ──────────────────

    // 1. Apps Script URL — config.json is the source of truth
    const scriptUrl = APP_CONFIG?.googleSheets?.scriptUrl;
    if (scriptUrl && scriptUrl.includes('script.google.com')) {
      Cache.set('apiUrl', scriptUrl);
    }

    // 2. Shift settings — merge with existing, config.json wins
    if (APP_CONFIG?.shift) {
      const s = APP_CONFIG.shift;
      Cache.set('settings', {
        start:       s.start       || SHIFT_DEFAULT.start,
        end:         s.end         || SHIFT_DEFAULT.end,
        lunchMins:   Number(s.lunchMins)   || SHIFT_DEFAULT.lunchMins,
        expectedHrs: Number(s.expectedHrs) || SHIFT_DEFAULT.expectedHrs
      });
    }

    // 3. Team names — config.json wins over local edits
    if (APP_CONFIG?.teams) {
      const t = APP_CONFIG.teams;
      if (t.teamA || t.teamB || t.teamC) {
        Cache.set('teamNames', {
          teamA: t.teamA || 'Team A',
          teamB: t.teamB || 'Team B',
          teamC: t.teamC || 'Team C'
        });
      }
    }

    // 4. User credentials — config.json can change both username (id) and password
    //    The role key (admin/guard/supervisor) is the stable identifier — never the username.
    //    This means you can rename 'admin' → 'gaurav' and 'guard' → 'gate' freely in config.json.
    if (APP_CONFIG && APP_CONFIG.users) {
      const existing = Cache.get('users', []);
      let changed = false;

      Object.entries(APP_CONFIG.users).forEach(function([roleKey, cfg]) {
        if (roleKey === '_note') return; // skip the comment field
        // Find user by ROLE (stable), not by current id (which may be changing)
        const user = existing.find(function(u) { return u.role === roleKey; });
        if (!user) return;

        // Apply new username if provided and different
        if (cfg.id && cfg.id !== user.id) {
          user.id  = cfg.id;
          changed  = true;
        }
        // Apply new password if provided and different
        if (cfg.pass && cfg.pass !== user.pass) {
          user.pass = cfg.pass;
          changed   = true;
        }
      });

      if (changed) {
        Cache.set('users', existing);
        console.log('User credentials updated from config.json');
      }
    }

    console.log('✅ config.json loaded:', APP_CONFIG.version || '');
  } catch (err) {
    // Config file missing or network error — use whatever is in cache
    console.warn('config.json not loaded:', err.message, '— using local cache');
    APP_CONFIG = {};
  }
}

/* ════════════════════════════════════════════════════════════════════
   GOOGLE SHEETS API LAYER
   ════════════════════════════════════════════════════════════════════ */
const GSheet = {
  /**
   * Returns true if Sheets integration is available.
   * The actual Apps Script URL lives in Netlify's APPS_SCRIPT_URL env var — never in the browser.
   * We check by pinging /api — if the function exists, Sheets is configured.
   * For simplicity, we return true if config.json has a scriptUrl OR if manually set.
   * The proxy itself will return a clear error if APPS_SCRIPT_URL isn't set in Netlify.
   */
  isConfigured() {
    // If config.json has a scriptUrl, Sheets is configured
    const cfgUrl = APP_CONFIG && APP_CONFIG.googleSheets && APP_CONFIG.googleSheets.scriptUrl;
    if (cfgUrl && cfgUrl.includes('script.google.com')) return true;
    // If manually saved
    const manUrl = Cache.get('apiUrl', null);
    if (manUrl && manUrl.includes('script.google.com')) return true;
    return false;
  },

  /**
   * All API calls go through the Netlify proxy at /api.
   * The proxy forwards to Apps Script server-side — zero CORS issues.
   *
   * Browser  →  /api (same origin, no CORS)
   *          →  Apps Script (server-to-server, no CORS)
   *
   * @param {string} action - API action name
   * @param {object} body   - Request payload
   * @returns {Promise<any>}
   */
  async call(action, body = {}) {
    if (!this.isConfigured()) throw new Error('Google Sheets not configured — set scriptUrl in Settings');

    const resp = await fetch('/api?action=' + encodeURIComponent(action), {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(body)
    });

    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const json = await resp.json();
    if (!json.ok) throw new Error(json.error || 'API error');
    return json.data;
  },

  /** GET-style call — still goes through proxy */
  async get(action, params = {}) {
    if (!this.isConfigured()) throw new Error('Google Sheets not configured');

    const qs   = Object.entries(params).map(([k,v]) => '&' + k + '=' + encodeURIComponent(v)).join('');
    const resp = await fetch('/api?action=' + encodeURIComponent(action) + qs);

    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const json = await resp.json();
    if (!json.ok) throw new Error(json.error || 'API error');
    return json.data;
  }
};

// ── App state ─────────────────────────────────────────────────────────
let APP_READY = false;  // true once initial data is loaded from Sheets

/* ────────────────────────────────────────────────────────────────
   STORAGE HELPERS — thin wrapper around localStorage
   ──────────────────────────────────────────────────────────────── */
const Store = {
  get(key, fallback = null) {
    try { return JSON.parse(localStorage.getItem('helix_' + key)) ?? fallback; }
    catch { return fallback; }
  },
  set(key, value) {
    localStorage.setItem('helix_' + key, JSON.stringify(value));
  }
};

/* ────────────────────────────────────────────────────────────────
   INITIALIZE DATABASE with default users and settings
   ──────────────────────────────────────────────────────────────── */
/**
 * initApp — called on page load.
 * Loads data from Google Sheets into local cache if URL is configured.
 * Falls back to localStorage-only if not configured.
 */
async function initApp() {
  // Step 1 — Seed local defaults if first run
  if (!Cache.get('users')) {
    Cache.set('users', [
      { id: 'admin',      pass: 'admin123', role: 'admin',      name: 'Administrator' },
      { id: 'guard',      pass: 'guard123', role: 'guard',      name: 'Gate Guard' },
      { id: 'supervisor', pass: 'sup123',   role: 'supervisor', name: 'Supervisor', label: 'Supervisor' }
    ]);
  }
  if (!Cache.get('settings'))   Cache.set('settings',   SHIFT_DEFAULT);
  if (!Cache.get('teamNames'))  Cache.set('teamNames',  { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  if (!Cache.get('empCounter')) Cache.set('empCounter', 1000);
  if (!Cache.get('employees'))  Cache.set('employees',  []);
  if (!Cache.get('attendance')) Cache.set('attendance', []);

  // Step 2 — Fetch config.json from server (applies URL, shift, teams, passwords centrally)
  // This is what makes one config file control all devices
  showGlobalStatus('⏳ Loading configuration…');
  await loadRemoteConfig();

  APP_READY = true;

  // Step 3 — Load live data from Google Sheets (if configured)
  if (GSheet.isConfigured()) {
    try {
      showGlobalStatus('⏳ Syncing with Google Sheets…');
      // Load employees + settings + team names in one call
      const all = await GSheet.get('getAll');
      // Only replace local employees if Sheet has data — never wipe local with empty Sheet
      const sheetEmps = Array.isArray(all.employees) ? all.employees : [];
      const localEmps = Cache.get('employees', []);
      if (sheetEmps.length > 0) {
        Cache.set('employees', sheetEmps);

        // ── CRITICAL: Sync empCounter to highest existing EMP number ──
        // This prevents new employees getting IDs that already exist in the Sheet.
        // Extract the numeric part of each EMP ID (e.g. EMP1043 → 1043) and take the max.
        const maxId = sheetEmps.reduce(function(max, emp) {
          const num = parseInt((emp.id || '').replace(/[^0-9]/g, ''), 10);
          return (!isNaN(num) && num > max) ? num : max;
        }, 1000);

        const localCounter = Cache.get('empCounter', 1000);
        if (maxId > localCounter) {
          Cache.set('empCounter', maxId);
          console.log('empCounter synced to', maxId);
        }

      } else if (localEmps.length > 0) {
        // Sheet empty but local has data — push local to Sheet silently
        GSheet.call('importEmployees', { employees: localEmps })
          .catch(e => console.warn('Auto-push to Sheet failed:', e.message));
      }
      if (all.settings && Object.keys(all.settings).length) {
        const s = all.settings;
        if (!APP_CONFIG?.shift) {
          Cache.set('settings', {
            start:       s.start       || SHIFT_DEFAULT.start,
            end:         s.end         || SHIFT_DEFAULT.end,
            lunchMins:   Number(s.lunchMins)   || SHIFT_DEFAULT.lunchMins,
            expectedHrs: Number(s.expectedHrs) || SHIFT_DEFAULT.expectedHrs
          });
        }
      }
      if (all.teamNames && !APP_CONFIG?.teams) Cache.set('teamNames', all.teamNames);

      // Also pull today's attendance so reports are current on any device
      const today = new Date().toISOString().split('T')[0];
      try {
        const todayAtt = await GSheet.get('getAttendance', { date: today });
        if (Array.isArray(todayAtt) && todayAtt.length) {
          // Merge today's records from Sheets with any locally-pending records
          const local   = Cache.get('attendance', []);
          const localIds = new Set(local.map(r => r.id));
          const merged  = [...local];
          todayAtt.forEach(r => { if (!localIds.has(r.id)) merged.push(r); });
          Cache.set('attendance', merged);
        }
      } catch(e) { /* non-critical — today's attendance just won't be pre-loaded */ }

      showGlobalStatus('');
    } catch(err) {
      showGlobalStatus('⚠ Sheets offline — using cached data');
      console.warn('Sheets load failed:', err.message);
    }
  } else {
    showGlobalStatus('');
  }
}

/** Dismissible status bar shown at top of app */
function showGlobalStatus(msg) {
  let bar = document.getElementById('globalStatusBar');
  if (!bar) {
    bar = document.createElement('div');
    bar.id = 'globalStatusBar';
    bar.style.cssText = [
      'position:fixed','top:0','left:0','right:0','z-index:9999',
      'padding:9px 16px','font-size:13px','font-weight:600',
      'text-align:center','background:#fef3c7','color:#78350f',
      'border-bottom:2px solid #fde68a'
    ].join(';');
    document.body.prepend(bar);
  }
  if (!msg) { bar.style.display='none'; return; }
  bar.style.display = 'block';
  bar.textContent   = msg;
}

initApp();

/* ────────────────────────────────────────────────────────────────
   SESSION STATE
   ──────────────────────────────────────────────────────────────── */
let currentUser    = null;   // Active logged-in user object
let loginRoleMode  = 'admin'; // Which role tab is selected
let currentQREmp   = null;   // Employee whose QR is being displayed
let scannerFilter  = 'all';  // Report filter: all | guard | sup1 | sup2 | sup3
let reportMode     = 'daily'; // Current report tab
let lastScannedId  = null;   // Last scanned QR data (for debounce)
let lastScanTime   = 0;

// Scanner stream references — separate per role
let guardStream    = null;
let guardInterval  = null;
let supStream      = null;
let supInterval    = null;

/* ────────────────────────────────────────────────────────────────
   LOGIN / AUTH
   ──────────────────────────────────────────────────────────────── */

/**
 * Sets which role tab is selected on the login screen.
 * @param {string} role - 'admin' | 'guard' | 'supervisor'
 */
function setRole(role) {
  loginRoleMode = role;
  document.querySelectorAll('.role-btn').forEach(b => b.classList.remove('active'));
  event.currentTarget.classList.add('active');
}

/** Validates credentials and routes to appropriate screen. */
async function doLogin() {
  const uid   = document.getElementById('loginUser').value.trim();
  const pass  = document.getElementById('loginPass').value;
  const errEl = document.getElementById('loginError');
  errEl.textContent = '';

  // Always ensure default users exist in cache before trying to auth
  // This guards against initApp not having finished yet
  ensureDefaultUsers();

  let user = null;

  // ── Step 1: Always try local cache first (fast, works offline) ──
  const cachedUsers = Cache.get('users', []);
  user = cachedUsers.find(u => u.id === uid && u.pass === pass && u.role === loginRoleMode);

  if (user) {
    // Local auth succeeded — no need to hit Sheets
    proceedAfterLogin(user, errEl);
    return;
  }

  // ── Step 2: If Sheets configured, try Sheets auth (handles password changes via Sheet) ──
  if (GSheet.isConfigured()) {
    try {
      errEl.textContent = '⏳ Verifying…';
      const result = await GSheet.call('login', { id: uid, pass, role: loginRoleMode });
      if (result && result.authenticated) {
        user = { ...result, role: loginRoleMode };
        // Update local cache with this validated user
        const users = Cache.get('users', []);
        const existing = users.find(u => u.id === uid);
        if (existing) { existing.pass = pass; Cache.set('users', users); }
        proceedAfterLogin(user, errEl);
        return;
      }
    } catch(err) {
      console.warn('Sheets auth error:', err.message);
      // Fall through to show error
    }
  }

  // ── Both failed ──
  errEl.textContent = 'Incorrect username or password.';
}

/** Ensures the three default users always exist in local cache. */
function ensureDefaultUsers() {
  let users = Cache.get('users', []);
  const defaults = [
    { id: 'admin',      pass: 'admin123', role: 'admin',      name: 'Administrator' },
    { id: 'guard',      pass: 'guard123', role: 'guard',      name: 'Gate Guard'    },
    { id: 'supervisor', pass: 'sup123',   role: 'supervisor', name: 'Supervisor', label: 'Supervisor' }
  ];
  let changed = false;
  defaults.forEach(def => {
    if (!users.find(u => u.id === def.id)) {
      users.push(def);
      changed = true;
    }
  });
  if (changed) Cache.set('users', users);
}

/** Handles everything after a successful login (routing, UI setup). */
function proceedAfterLogin(user, errEl) {
  currentUser = user;
  if (errEl) errEl.textContent = '';
  document.getElementById('loginUser').value = '';
  document.getElementById('loginPass').value = '';

  // Refresh employees from Sheets in background (non-blocking)
  if (GSheet.isConfigured()) {
    GSheet.get('getEmployees')
      .then(fresh => {
        if (Array.isArray(fresh) && fresh.length) {
          Cache.set('employees', fresh);
          if (typeof renderEmployeeList === 'function') renderEmployeeList();
          if (typeof renderSupTeam      === 'function') renderSupTeam();
        }
      })
      .catch(e => console.warn('Post-login employee refresh failed:', e.message));
  }

  // Route to correct screen based on role
  if (user.role === 'admin') {
    document.getElementById('adminAvatar').textContent = (user.name || 'A')[0].toUpperCase();
    showScreen('adminScreen');
    refreshDashboard();
  } else if (user.role === 'guard') {
    document.getElementById('guardAvatar').textContent = (user.name || 'G')[0].toUpperCase();
    showScreen('guardScreen');
    renderGuardScans();
  } else if (user.role === 'supervisor') {
    document.getElementById('supAvatar').textContent = (user.name || 'S')[0].toUpperCase();
    document.getElementById('supRoleBadge').textContent = user.label || user.name || 'Supervisor';
    document.getElementById('supScanHeading').textContent = 'Scanner — ' + (user.label || user.name || 'Supervisor');
    showScreen('supervisorScreen');
    renderSupScans();
    renderSupTeam();
  }
}

/** Logs out and returns to login screen, stopping any active camera. */
function logout() {
  guardStopScanner();
  supStopScanner();
  currentUser = null;
  showScreen('loginScreen');
}

/* ────────────────────────────────────────────────────────────────
   SCREEN & PAGE NAVIGATION
   ──────────────────────────────────────────────────────────────── */

function showScreen(id) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
}

/**
 * Admin bottom-nav switcher.
 * @param {string} page - 'dashboard' | 'employees' | 'reports' | 'settings'
 */
function adminNav(page) {
  document.querySelectorAll('#adminScreen .page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('#adminScreen .nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('aP-' + page).classList.add('active');
  document.getElementById('aN-' + page).classList.add('active');

  // Trigger data refresh when entering each section
  if (page === 'dashboard') refreshDashboard();
  if (page === 'employees') { renderEmployeeList(); updateTeamDropdown(); }
  if (page === 'reports') {
    const today = todayStr();
    document.getElementById('rDailyDate').value = today;
    document.getElementById('rMonthPicker').value = today.slice(0, 7);
    updateReportFilterLabels();
    renderDailyReport();
  }
  if (page === 'settings') loadSettings();
}

/**
 * Supervisor bottom-nav switcher.
 * @param {string} page - 'scan' | 'team'
 */
function supNav(page) {
  document.querySelectorAll('#supervisorScreen .page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('#supervisorScreen .nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('sP-' + page).classList.add('active');
  document.getElementById('sN-' + page).classList.add('active');

  if (page === 'team') renderSupTeam();
}

/* ────────────────────────────────────────────────────────────────
   EMPLOYEE MANAGEMENT
   ──────────────────────────────────────────────────────────────── */

function toggleAddForm() {
  const card = document.getElementById('addEmpFormCard');
  card.style.display = card.style.display === 'none' ? 'block' : 'none';
}

/**
 * Adds a new employee to the database and generates a unique ID.
 */
function addEmployee() {
  const first  = document.getElementById('efFirst').value.trim();
  const last   = document.getElementById('efLast').value.trim();
  const dept   = document.getElementById('efDept').value.trim();
  const supId  = document.getElementById('efSupervisor').value;
  const phone  = document.getElementById('efPhone').value.trim();
  const errEl  = document.getElementById('empFormError');

  if (!first) { errEl.textContent = 'First name is required.'; return; }

  // Generate next ID — empCounter is kept in sync with Sheet on every app load
  const counter = Cache.get('empCounter', 1000) + 1;
  Cache.set('empCounter', counter);

  const emp = {
    id:         'EMP' + counter,
    firstName:  first,
    lastName:   last,
    name:       last ? first + ' ' + last : first,
    dept:       dept || 'General',
    supervisor: supId || '',   // Which supervisor this employee belongs to
    phone:      phone,
    createdAt:  new Date().toISOString()
  };

  // Save locally first (instant UI)
  const emps = Cache.get('employees', []);
  emps.push(emp);
  Cache.set('employees', emps);

  // Reset form
  ['efFirst','efLast','efDept','efPhone'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('efSupervisor').value = '';
  errEl.textContent = '';
  toggleAddForm();
  renderEmployeeList();
  showToast('✅ ' + emp.name + ' added');

  // Sync to Sheets in background
  if (GSheet.isConfigured()) {
    GSheet.call('addEmployee', emp)
      .catch(err => showToast('⚠ Sheets sync failed: ' + err.message, 4000));
  }
}

/**
 * Renders the employee list with optional search filter.
 * Shows today's status badge next to each employee.
 */
function renderEmployeeList() {
  const q    = (document.getElementById('empSearchInput')?.value || '').toLowerCase();
  const emps = Store.get('employees', []).filter(e =>
    !q || e.name.toLowerCase().includes(q) ||
    e.id.toLowerCase().includes(q) ||
    (e.dept || '').toLowerCase().includes(q)
  );

  const card = document.getElementById('empListCard');
  if (!emps.length) {
    card.innerHTML = '<div class="empty"><div class="empty-icon">👥</div><p>No employees found</p></div>';
    return;
  }

  const today = todayStr();
  const att   = Store.get('attendance', []);

  card.innerHTML = emps.map(emp => {
    const todayRecs = att.filter(r => r.empId === emp.id && r.date === today);
    const lastRec   = todayRecs.at(-1);
    const badge     = getStatusBadge(lastRec);
    const initials  = (emp.firstName[0] + (emp.lastName ? emp.lastName[0] : '')).toUpperCase();
    const supLabel  = emp.supervisor ? getSupervisorLabel(emp.supervisor) : '';

    return `
      <div class="emp-row">
        <div class="emp-avatar">${initials}</div>
        <div class="emp-info">
          <div class="emp-name">${emp.name} ${badge}</div>
          <div class="emp-meta">${emp.id} · ${emp.dept}${supLabel ? ' · ' + supLabel : ''}</div>
        </div>
        <div class="emp-actions">
          <button class="btn btn-gold btn-xs" onclick="openQRPopup('${emp.id}')">QR</button>
          <button class="btn btn-outline btn-xs" onclick="editEmployee('${emp.id}')">✏️</button>
          <button class="btn btn-danger btn-xs" onclick="deleteEmployee('${emp.id}')">✕</button>
        </div>
      </div>`;
  }).join('');
}

/** Deletes employee and their attendance records. */
function deleteEmployee(empId) {
  if (!confirm('Delete this employee and all their attendance data?')) return;

  // Remove from local cache
  Cache.set('employees', Cache.get('employees', []).filter(function(e) { return e.id !== empId; }));
  Cache.set('attendance', Cache.get('attendance', []).filter(function(r) { return r.empId !== empId; }));
  renderEmployeeList();
  showToast('🗑 Employee removed');

  if (GSheet.isConfigured()) {
    GSheet.call('deleteEmployee', { id: empId })
      .then(function() {
        // After deletion confirmed on Sheet, re-pull fresh list to ensure local is in sync
        return GSheet.get('getEmployees');
      })
      .then(function(fresh) {
        if (Array.isArray(fresh)) {
          Cache.set('employees', fresh);
          // Re-sync counter after any deletion
          const maxId = fresh.reduce(function(max, emp) {
            const num = parseInt((emp.id || '').replace(/[^0-9]/g, ''), 10);
            return (!isNaN(num) && num > max) ? num : max;
          }, 1000);
          if (maxId > Cache.get('empCounter', 1000)) Cache.set('empCounter', maxId);
          renderEmployeeList();
        }
      })
      .catch(function(err) { showToast('⚠ Sheets sync failed: ' + err.message, 4000); });
  }
}

/** Returns human-readable supervisor label for a supervisor ID. */
function getSupervisorLabel(supId) {
  if (!supId) return '—';
  // Use custom team names from settings
  const n = Store.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  if (supId === 'teamA') return n.teamA;
  if (supId === 'teamB') return n.teamB;
  if (supId === 'teamC') return n.teamC;
  if (supId === 'supervisor') return 'Supervisor';
  // Legacy fallback
  const legacyLabels = { 'sup1': 'Supervisor 1', 'sup2': 'Supervisor 2', 'sup3': 'Supervisor 3' };
  if (legacyLabels[supId]) return legacyLabels[supId];
  const users = Store.get('users', []);
  const s = users.find(u => u.id === supId);
  return s ? s.label || s.name : supId;
}

/* ────────────────────────────────────────────────────────────────
   QR CODE GENERATION
   ──────────────────────────────────────────────────────────────── */

/**
 * Opens the QR popup for a given employee.
 * Encodes employee ID + name + department into the QR payload.
 * @param {string} empId
 */
function openQRPopup(empId) {
  const emps = Store.get('employees', []);
  const emp  = emps.find(e => e.id === empId);
  if (!emp) return;

  currentQREmp = emp;
  document.getElementById('qrPopupName').textContent = emp.name;
  document.getElementById('qrPopupId').textContent   = emp.id + ' · ' + emp.dept;

  // Use shared generateQRCanvas helper (also used by bulk download)
  const target = document.getElementById('qrRenderTarget');
  target.innerHTML = '';
  const qrCanvas    = generateQRCanvas(emp, 280);
  target.appendChild(qrCanvas);
  target._qrCanvas  = qrCanvas;

  openOverlay('qrOverlay');
}

/** Downloads the generated QR code canvas as PNG. */
function downloadQRCode() {
  if (!currentQREmp) return;
  const target = document.getElementById('qrRenderTarget');
  const canvas = target._qrCanvas || target.querySelector('canvas');
  if (!canvas) { showToast('QR not ready, try again'); return; }
  const a = document.createElement('a');
  a.download = 'QR_' + currentQREmp.id + '_' + currentQREmp.name.replace(/\s+/g, '_') + '.png';
  a.href = canvas.toDataURL();
  a.click();
  showToast('⬇ QR downloaded');
}

/* ────────────────────────────────────────────────────────────────
   QR SCANNER — GUARD
   ──────────────────────────────────────────────────────────────── */

/** Starts the camera for the guard's scanner. */
function guardStartScanner() {
  initCameraScanner(
    'guardVideo', 'guardCanvas', 'guardScanStatus',
    'guardStartBtn', 'guardStopBtn',
    stream => { guardStream = stream; },
    intervalId => { guardInterval = intervalId; },
    qrData => processScan(qrData, 'guard')
  );
}

/** Stops the guard's camera scanner. */
function guardStopScanner() {
  stopCameraScanner(
    guardStream, guardInterval,
    'guardScanStatus', 'guardStartBtn', 'guardStopBtn', 'guardVideo'
  );
  guardStream = null; guardInterval = null;
}

/* ────────────────────────────────────────────────────────────────
   QR SCANNER — SUPERVISOR
   ──────────────────────────────────────────────────────────────── */

/** Starts the camera for the supervisor's scanner. */
function supStartScanner() {
  initCameraScanner(
    'supVideo', 'supCanvas', 'supScanStatus',
    'supStartBtn', 'supStopBtn',
    stream => { supStream = stream; },
    intervalId => { supInterval = intervalId; },
    qrData => processScan(qrData, currentUser ? currentUser.id : 'supervisor')
  );
}

/** Stops the supervisor's camera scanner. */
function supStopScanner() {
  stopCameraScanner(
    supStream, supInterval,
    'supScanStatus', 'supStartBtn', 'supStopBtn', 'supVideo'
  );
  supStream = null; supInterval = null;
}

/* ────────────────────────────────────────────────────────────────
   CAMERA SCANNER — SHARED UTILITIES
   ──────────────────────────────────────────────────────────────── */

/**
 * Generic camera initialiser used by both Guard and Supervisor scanners.
 * @param {string} videoId  - ID of <video> element
 * @param {string} canvasId - ID of <canvas> element (hidden, for frame capture)
 * @param {string} statusId - ID of status display div
 * @param {string} startBtnId - ID of start button
 * @param {string} stopBtnId  - ID of stop button
 * @param {function} onStream - Callback receiving the MediaStream
 * @param {function} onInterval - Callback receiving the setInterval ID
 * @param {function} onQRFound  - Callback receiving decoded QR string
 */
function initCameraScanner(videoId, canvasId, statusId, startBtnId, stopBtnId, onStream, onInterval, onQRFound) {
  if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
    document.getElementById(statusId).textContent = '❌ Camera not supported — use Chrome or Safari';
    return;
  }

  var statusEl   = document.getElementById(statusId);
  var startBtn   = document.getElementById(startBtnId);
  var stopBtn    = document.getElementById(stopBtnId);
  var vid        = document.getElementById(videoId);
  var cvs        = document.getElementById(canvasId);
  var ctx        = cvs.getContext('2d', { willReadFrequently: true });

  statusEl.textContent = '⏳ Starting camera…';

  // Simple single constraint — just ask for video, no facingMode fighting
  navigator.mediaDevices.getUserMedia({
      audio: false,
      video: {
        facingMode: { ideal: 'environment' },  // rear camera on phones, webcam on laptops
        width:  { ideal: 1280 },
        height: { ideal: 720 }
      }
    })
    .then(function(stream) {
      onStream(stream);
      vid.srcObject = stream;
      vid.setAttribute('playsinline', '');
      vid.muted = true;

      vid.play().then(function() {
        startBtn.style.display = 'none';
        stopBtn.style.display  = 'flex';
        statusEl.textContent   = '🟢 Scanning — hold QR code in front of camera';
        statusEl.classList.add('ready');

        // ── Scan loop using the 3-pass adaptive method proven in testing ──
        var scanId = setInterval(function() {
          if (!vid.videoWidth || vid.paused || vid.ended) return;

          var w = vid.videoWidth, h = vid.videoHeight;
          cvs.width = w; cvs.height = h;
          ctx.drawImage(vid, 0, 0, w, h);
          var imgData = ctx.getImageData(0, 0, w, h);
          var data    = imgData.data;

          // Pass 1: raw pixels, both inversion modes
          var r1 = jsQR(new Uint8ClampedArray(data), w, h, { inversionAttempts: 'attemptBoth' });
          if (r1 && r1.data) { onQRFound(r1.data); return; }

          // Pass 2: simple luminance threshold at 128
          var bin = new Uint8ClampedArray(data.length);
          for (var i = 0; i < data.length; i += 4) {
            var lum = (data[i] * 77 + data[i+1] * 150 + data[i+2] * 29) >> 8;
            var v   = lum > 128 ? 255 : 0;
            bin[i] = bin[i+1] = bin[i+2] = v; bin[i+3] = 255;
          }
          var r2 = jsQR(bin, w, h, { inversionAttempts: 'attemptBoth' });
          if (r2 && r2.data) { onQRFound(r2.data); return; }

          // Pass 3: adaptive threshold (handles glare & uneven screen brightness)
          var gray = new Float32Array(w * h);
          for (var i = 0; i < data.length; i += 4) {
            gray[i >> 2] = (data[i] * 77 + data[i+1] * 150 + data[i+2] * 29) >> 8;
          }
          var adapt  = new Uint8ClampedArray(data.length);
          var radius = 15;
          for (var y = 0; y < h; y++) {
            for (var x = 0; x < w; x++) {
              var sum = 0, count = 0;
              for (var dy = -radius; dy <= radius; dy += 5) {
                for (var dx = -radius; dx <= radius; dx += 5) {
                  var ny = y + dy, nx = x + dx;
                  if (ny >= 0 && ny < h && nx >= 0 && nx < w) { sum += gray[ny * w + nx]; count++; }
                }
              }
              var pi = (y * w + x) * 4;
              var av = sum / count - 5;
              var vv = gray[y * w + x] < av ? 0 : 255;
              adapt[pi] = adapt[pi+1] = adapt[pi+2] = vv; adapt[pi+3] = 255;
            }
          }
          var r3 = jsQR(adapt, w, h, { inversionAttempts: 'attemptBoth' });
          if (r3 && r3.data) { onQRFound(r3.data); }

        }, 150);

        onInterval(scanId);
      }).catch(function(e) {
        statusEl.textContent = '❌ Could not start video: ' + e.message;
      });
    })
    .catch(function(e) {
      statusEl.textContent = '❌ Camera blocked: ' + e.name + ' — allow camera permission';
    });
}


function stopCameraScanner(stream, intervalId, statusId, startBtnId, stopBtnId, videoId) {
  if (intervalId) clearInterval(intervalId);
  if (stream) {
    stream.getTracks().forEach(function(t) { try { t.stop(); } catch(e){} });
  }
  var vid = document.getElementById(videoId);
  if (vid) { vid.pause(); vid.srcObject = null; }
  var startBtn = document.getElementById(startBtnId);
  var stopBtn  = document.getElementById(stopBtnId);
  if (startBtn) startBtn.style.display = 'flex';
  if (stopBtn)  stopBtn.style.display  = 'none';
  var statusEl = document.getElementById(statusId);
  if (statusEl) {
    statusEl.textContent = '📷 Tap button to start camera';
    statusEl.classList.remove('ready');
  }
}

/* ────────────────────────────────────────────────────────────────
   ATTENDANCE SCAN PROCESSING
   ──────────────────────────────────────────────────────────────── */

/**
 * Processes a decoded QR string, determines attendance action, and records it.
 * Action sequence: Check In → Break Start → Break End → Check Out
 * If an employee is already checked out, next scan = new Check In.
 *
 * @param {string} raw       - Raw QR string (JSON payload)
 * @param {string} scannedBy - User ID of who scanned (guard | sup1 | sup2 | sup3)
 */
// Pending scan state — holds employee + scannedBy while guard picks action
var pendingScan = null;

/**
 * processScan — called when QR is decoded.
 * Instead of auto-determining action, shows a manual action picker.
 * Guard/Supervisor taps Check In / Check Out / Break Start / Break End.
 */
async function processScan(raw, scannedBy) {
  // Debounce: ignore same QR within 4 seconds
  const now = Date.now();
  if (raw === lastScannedId && now - lastScanTime < 4000) return;
  lastScannedId = raw;
  lastScanTime  = now;

  let payload;
  try { payload = JSON.parse(raw); }
  catch { showToast('⚠ Invalid QR code'); return; }

  let emps = Cache.get('employees', []);
  let emp  = emps.find(e => e.id === payload.id);

  // If not found locally and Sheets is configured, do a live fetch
  // This handles the case where cache is stale or empty (e.g. first load on phone)
  if (!emp && GSheet.isConfigured()) {
    showToast('⏳ Looking up employee…', 3000);
    try {
      const fresh = await GSheet.get('getEmployees');
      if (Array.isArray(fresh) && fresh.length) {
        Cache.set('employees', fresh);           // update local cache
        emps = fresh;
        emp  = fresh.find(e => e.id === payload.id);
      }
    } catch(e) {
      console.warn('Live employee fetch failed:', e.message);
    }
  }

  if (!emp) { showToast('⚠ Employee not found: ' + payload.id); return; }

  // Store pending context
  pendingScan = { emp, scannedBy };

  // Show today's existing records in the picker
  const dateStr   = todayStr();
  const att       = Store.get('attendance', []);
  const todayRecs = att.filter(r => r.empId === emp.id && r.date === dateStr);
  const lastRec   = todayRecs.at(-1);

  // Populate action picker
  document.getElementById('apEmpName').textContent = emp.name;
  document.getElementById('apEmpId').textContent   = emp.id + ' · ' + emp.dept;

  // Show timeline of today's records
  const labels   = { 'check-in': 'Check In', 'check-out': 'Check Out', 'break-start': 'Break Start', 'break-end': 'Break End' };
  const dotClass = { 'check-in': 'in', 'check-out': 'out', 'break-start': 'break', 'break-end': 'in' };
  const tlHtml   = todayRecs.length
    ? '<div class="timeline">' + todayRecs.map(r =>
        '<div class="tl-item"><div class="tl-dot ' + (dotClass[r.type]||'in') + '"></div>' +
        '<span class="tl-time">' + r.time + '</span> <span class="tl-label">' + (labels[r.type]||r.type) + '</span></div>'
      ).join('') + '</div>'
    : '<div style="font-size:12px;color:var(--text-muted);padding:4px 0">No records yet today</div>';
  document.getElementById('apTimeline').innerHTML = tlHtml;

  // Smart button highlighting — suggest the most likely next action
  // but ALL buttons remain enabled so guard can choose freely
  const allBtns = ['apBtnIn','apBtnOut','apBtnBreakStart','apBtnBreakEnd'];
  allBtns.forEach(id => {
    var btn = document.getElementById(id);
    btn.style.opacity = '0.6';
    btn.style.transform = 'scale(0.97)';
  });

  // Determine suggested action based on last record
  let suggested = 'apBtnIn';
  if (!lastRec || lastRec.type === 'check-out')                          suggested = 'apBtnIn';
  else if (lastRec.type === 'check-in' || lastRec.type === 'break-end') suggested = 'apBtnOut';
  else if (lastRec.type === 'break-start')                               suggested = 'apBtnBreakEnd';

  // Highlight suggested button
  const sBtn = document.getElementById(suggested);
  sBtn.style.opacity   = '1';
  sBtn.style.transform = 'scale(1.03)';
  sBtn.style.boxShadow = '0 4px 16px rgba(0,0,0,0.15)';

  // Haptic
  if (navigator.vibrate) navigator.vibrate(60);

  openOverlay('actionPickerOverlay');
}

/**
 * confirmAction — called when guard taps one of the 4 action buttons.
 * Records attendance and shows result sheet.
 * @param {string} type - 'check-in' | 'check-out' | 'break-start' | 'break-end'
 */
async function confirmAction(type) {
  if (!pendingScan) return;
  closeOverlay('actionPickerOverlay');

  const { emp, scannedBy } = pendingScan;
  pendingScan = null;

  const dateStr   = todayStr();
  const timeStr   = new Date().toTimeString().slice(0, 8);

  // Pull latest records for this employee from Sheets before saving
  // Prevents conflicts when multiple devices scan the same employee
  if (GSheet.isConfigured()) {
    try {
      const sheetAtt = await GSheet.get('getAttendance', { date: dateStr });
      if (Array.isArray(sheetAtt)) {
        const local    = Cache.get('attendance', []);
        const localIds = new Set(local.map(r => r.id));
        sheetAtt.forEach(r => { if (!localIds.has(r.id)) local.push(r); });
        Cache.set('attendance', local);
      }
    } catch(e) { /* proceed with cached data if Sheets unreachable */ }
  }

  const att       = Cache.get('attendance', []);
  const todayRecs = att.filter(r => r.empId === emp.id && r.date === dateStr);

  const record = {
    id:        Date.now().toString(),
    empId:     emp.id,
    empName:   emp.name,
    dept:      emp.dept,
    date:      dateStr,
    time:      timeStr,
    timestamp: new Date().toISOString(),
    type:      type,
    scannedBy: scannedBy
  };

  att.push(record);
  Cache.set('attendance', att);

  showScanResult(emp, record, [...todayRecs, record], scannedBy);
  renderGuardScans();
  renderSupScans();

  if (navigator.vibrate) navigator.vibrate([80, 40, 80]);

  // Sync attendance record to Sheets in background (non-blocking)
  if (GSheet.isConfigured()) {
    GSheet.call('addAttendance', record)
      .catch(err => console.warn('Attendance sync failed:', err.message));
  }
}

/** Cancel action picker — go back to scanning without saving anything. */
function cancelAction() {
  pendingScan   = null;
  lastScannedId = null; // allow re-scan of same QR
  closeOverlay('actionPickerOverlay');
}

/**
 * Displays the scan result bottom sheet with timeline.
 * @param {object} emp        - Employee object
 * @param {object} record     - The new attendance record just created
 * @param {Array}  todayRecs  - All records for this employee today (incl. new one)
 * @param {string} scannedBy  - Scanner user ID
 */
function showScanResult(emp, record, todayRecs, scannedBy) {
  const actionColors = {
    'check-in':    { bg: 'var(--pastel-green)', color: '#065f46' },
    'check-out':   { bg: 'var(--pastel-rose)',  color: '#9d174d' },
    'break-start': { bg: 'var(--pastel-amber)', color: '#92400e' },
    'break-end':   { bg: 'var(--pastel-sky)',   color: '#0369a1' }
  };
  const emojis = { 'check-in': '✅', 'check-out': '🚪', 'break-start': '☕', 'break-end': '▶️' };
  const labels = { 'check-in': 'CHECK IN', 'check-out': 'CHECK OUT', 'break-start': 'BREAK START', 'break-end': 'BREAK END' };

  document.getElementById('srEmoji').textContent  = emojis[record.type];
  document.getElementById('srName').textContent   = emp.name;
  document.getElementById('srId').textContent     = emp.id + ' · ' + emp.dept;

  const actionEl = document.getElementById('srAction');
  actionEl.textContent   = labels[record.type];
  const ac = actionColors[record.type];
  actionEl.style.background = ac.bg;
  actionEl.style.color       = ac.color;

  // Show scanned-by supervisor label (hidden for guard)
  const supLabel = document.getElementById('srSupLabel');
  if (scannedBy && scannedBy !== 'guard') {
    supLabel.textContent  = '📋 Marked by ' + getSupervisorLabel(scannedBy);
    supLabel.style.display = 'inline-block';
  } else {
    supLabel.style.display = 'none';
  }

  document.getElementById('srTime').textContent =
    '🕐 ' + new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit', second: '2-digit' }) +
    '  ·  ' + new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });

  // Build timeline
  const dotClass = { 'check-in': 'in', 'check-out': 'out', 'break-start': 'break', 'break-end': 'in' };
  const tlHtml = todayRecs.map(r =>
    `<div class="tl-item">
       <div class="tl-dot ${dotClass[r.type] || 'in'}"></div>
       <div class="tl-time">${r.time}</div>
       <div class="tl-label">${labels[r.type] || r.type}</div>
     </div>`
  ).join('');

  document.getElementById('srTimeline').innerHTML =
    tlHtml ? `<div class="timeline">${tlHtml}</div>` : '';

  openOverlay('scanResultOverlay');
}

function closeScanResult() {
  closeOverlay('scanResultOverlay');
  lastScannedId = null; // Allow re-scan of same person after manual close
}

/* ────────────────────────────────────────────────────────────────
   RECENT SCAN LISTS
   ──────────────────────────────────────────────────────────────── */

const LOG_ICONS  = { 'check-in': '✅', 'check-out': '🚪', 'break-start': '☕', 'break-end': '▶️' };
const LOG_LABELS = { 'check-in': 'Check In', 'check-out': 'Check Out', 'break-start': 'Break Start', 'break-end': 'Break End' };
const LOG_DOTS   = { 'check-in': 'in', 'check-out': 'out', 'break-start': 'break', 'break-end': 'in' };

/**
 * Renders today's scans done by the gate guard.
 */
function renderGuardScans() {
  const today = todayStr();
  const recs  = Store.get('attendance', [])
    .filter(r => r.date === today && r.scannedBy === 'guard')
    .reverse().slice(0, 20);
  renderScanLog('guardRecentScans', recs);
}

/**
 * Renders today's scans done by the current supervisor.
 */
function renderSupScans() {
  if (!currentUser) return;
  const today = todayStr();
  const recs  = Store.get('attendance', [])
    .filter(r => r.date === today && r.scannedBy === currentUser.id)
    .reverse().slice(0, 20);
  renderScanLog('supRecentScans', recs);
}

/**
 * Generic scan log renderer.
 * @param {string} containerId - DOM element ID
 * @param {Array}  recs        - Array of attendance records
 */
function renderScanLog(containerId, recs) {
  const el = document.getElementById(containerId);
  if (!el) return;
  if (!recs.length) {
    el.innerHTML = '<div class="empty"><div class="empty-icon">📭</div><p>No scans yet today</p></div>';
    return;
  }
  el.innerHTML = recs.map(r => `
    <div class="log-item">
      <div class="log-icon ${LOG_DOTS[r.type] || 'in'}">${LOG_ICONS[r.type] || '📌'}</div>
      <div class="log-info">
        <div class="log-main">${r.empName}</div>
        <div class="log-sub">${r.empId || ''} · ${LOG_LABELS[r.type] || r.type}</div>
      </div>
      <div class="log-time">${r.time}</div>
    </div>`).join('');
}

/* ────────────────────────────────────────────────────────────────
   SUPERVISOR TEAM VIEW
   ──────────────────────────────────────────────────────────────── */

/**
 * Renders the supervisor's team page showing their assigned employees
 * and today's attendance status for each.
 */
function renderSupTeam() {
  if (!currentUser) return;
  const allEmps = Store.get('employees', []);
  // Filter employees assigned to this supervisor
  const myEmps  = allEmps.filter(e => e.supervisor === currentUser.id);

  document.getElementById('supTeamSub').textContent =
    `${myEmps.length} employee${myEmps.length !== 1 ? 's' : ''} under your supervision`;

  const today = todayStr();
  const att   = Store.get('attendance', []);
  let inCount = 0, outCount = 0;

  const card = document.getElementById('supTeamCard');
  if (!myEmps.length) {
    card.innerHTML = '<div class="empty"><div class="empty-icon">👥</div><p>No employees assigned to you yet.<br>Ask admin to assign employees.</p></div>';
    document.getElementById('supStatIn').textContent  = 0;
    document.getElementById('supStatOut').textContent = 0;
    return;
  }

  card.innerHTML = myEmps.map(emp => {
    const todayRecs = att.filter(r => r.empId === emp.id && r.date === today);
    const lastRec   = todayRecs.at(-1);
    const badge     = getStatusBadge(lastRec);
    const initials  = (emp.firstName[0] + (emp.lastName ? emp.lastName[0] : '')).toUpperCase();

    // Count for stats
    if (lastRec) {
      if (lastRec.type === 'check-in' || lastRec.type === 'break-end') inCount++;
      else if (lastRec.type === 'check-out') outCount++;
    }

    return `
      <div class="emp-row">
        <div class="emp-avatar">${initials}</div>
        <div class="emp-info">
          <div class="emp-name">${emp.name} ${badge}</div>
          <div class="emp-meta">${emp.id} · ${emp.dept}</div>
        </div>
      </div>`;
  }).join('');

  document.getElementById('supStatIn').textContent  = inCount;
  document.getElementById('supStatOut').textContent = outCount;
}

/* ────────────────────────────────────────────────────────────────
   DASHBOARD REFRESH
   ──────────────────────────────────────────────────────────────── */

/**
 * Updates dashboard stats and activity feed with today's data.
 */
function refreshDashboard() {
  const emps  = Store.get('employees', []);
  const today = todayStr();
  const att   = Store.get('attendance', []).filter(r => r.date === today);

  // Update date label
  document.getElementById('todayDateLabel').textContent =
    new Date().toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

  document.getElementById('sDashTotal').textContent = emps.length;

  // Count status per employee
  let inCount = 0, breakCount = 0, outCount = 0;
  emps.forEach(emp => {
    const recs = att.filter(r => r.empId === emp.id);
    const last = recs.at(-1);
    if (!last) return;
    if (last.type === 'check-in'   || last.type === 'break-end')  inCount++;
    else if (last.type === 'break-start')                          breakCount++;
    else if (last.type === 'check-out')                            outCount++;
  });

  document.getElementById('sDashIn').textContent    = inCount;
  document.getElementById('sDashBreak').textContent = breakCount;
  document.getElementById('sDashOut').textContent   = outCount;

  // Activity feed — most recent 25 records
  const recent = [...att].reverse().slice(0, 25);
  const logEl  = document.getElementById('dashActivityLog');
  if (!recent.length) {
    logEl.innerHTML = '<div class="empty"><div class="empty-icon">📭</div><p>No activity today</p></div>';
    return;
  }
  logEl.innerHTML = recent.map(r => `
    <div class="log-item">
      <div class="log-icon ${LOG_DOTS[r.type] || 'in'}">${LOG_ICONS[r.type] || '📌'}</div>
      <div class="log-info">
        <div class="log-main">${r.empName}</div>
        <div class="log-sub">${LOG_LABELS[r.type] || r.type} · ${r.scannedBy === 'guard' ? 'Gate Guard' : getSupervisorLabel(r.scannedBy)}</div>
      </div>
      <div class="log-time">${r.time}</div>
    </div>`).join('');
}

/* ────────────────────────────────────────────────────────────────
   ATTENDANCE CALCULATION
   ──────────────────────────────────────────────────────────────── */

/**
 * Calculates worked hours for one employee on a specific date.
 * Deducts lunch break if total time >= threshold.
 *
 * @param {string} empId - Employee ID
 * @param {string} date  - Date string YYYY-MM-DD
 * @returns {object} { checkIn, checkOut, breakMins, workedHrs, status, progress }
 */
function calcAttendance(empId, date, scannedByFilter) {
  scannedByFilter = scannedByFilter || scannerFilter || 'all';
  let att = Store.get('attendance', []).filter(r => r.empId === empId && r.date === date);
  // Separate guard vs supervisor records when filter is set
  if (scannedByFilter === 'guard') {
    att = att.filter(r => r.scannedBy === 'guard');
  } else if (scannedByFilter === 'supervisor') {
    att = att.filter(r => r.scannedBy === 'supervisor');
  }
  const cfg  = Store.get('settings', SHIFT_DEFAULT);

  if (!att.length) return { checkIn: '—', checkOut: '—', breakMins: 0, workedHrs: 0, status: 'absent', progress: 0 };

  const sorted   = [...att].sort((a, b) => a.time.localeCompare(b.time));
  const firstIn  = sorted.find(r => r.type === 'check-in');
  const lastOut  = sorted.filter(r => r.type === 'check-out').at(-1);

  if (!firstIn) return { checkIn: '—', checkOut: '—', breakMins: 0, workedHrs: 0, status: 'no-checkin', progress: 0 };

  // Calculate explicit break durations from break-start/break-end pairs
  let breakMins = 0;
  let bStart    = null;
  const hasBreakRecord = sorted.some(r => r.type === 'break-start' || r.type === 'break-end');

  sorted.forEach(r => {
    if (r.type === 'break-start') bStart = r.time;
    if (r.type === 'break-end' && bStart) {
      breakMins += timeToMins(r.time) - timeToMins(bStart);
      bStart = null;
    }
  });

  // If only Check In + Check Out (no break records at all) → auto-deduct lunch
  // This handles the "2 scan day" case: deduct lunchMins from settings (default 30)
  if (!hasBreakRecord && lastOut) {
    breakMins = cfg.lunchMins || 30;
  }

  // Total elapsed time (check-in to check-out, or to now if still working)
  const endTime   = lastOut ? lastOut.time : new Date().toTimeString().slice(0, 8);
  const totalMins = Math.max(0, timeToMins(endTime) - timeToMins(firstIn.time));

  // Net working minutes = total elapsed − breaks
  const rawWorkedMins = Math.max(0, totalMins - breakMins);

  // Round to nearest quarter-hour (15 min)
  // e.g. 8h 30m → 8.5, 8h 15m → 8.25, 8h 7m → 8.0, 8h 8m → 8.25
  const roundedMins = Math.round(rawWorkedMins / 15) * 15;
  const workedHrs   = roundedMins / 60;  // stored as decimal e.g. 8.5, 8.25, 8.75

  // Progress toward expected hours (0–100)
  const progress = Math.min(100, (workedHrs / cfg.expectedHrs) * 100);

  const status = !lastOut ? 'active' : 'complete';

  return {
    checkIn:   firstIn.time,
    checkOut:  lastOut ? lastOut.time : '—',
    breakMins: Math.round(breakMins),
    workedHrs: workedHrs,   // decimal quarters: 8.0, 8.25, 8.5, 8.75, 9.0 …
    status,
    progress:  Math.round(progress)
  };
}

/** Converts "HH:MM:SS" or "HH:MM" to total minutes. */
function timeToMins(t) {
  const [h, m, s] = t.split(':').map(Number);
  return h * 60 + m + (s || 0) / 60;
}

/** Formats minutes as "Xh Ym" string. */
function minsToHM(mins) {
  const h = Math.floor(mins / 60);
  const m = Math.round(mins % 60);
  return h + 'h ' + m + 'm';
}

/* ────────────────────────────────────────────────────────────────
   REPORTS — DAILY
   ──────────────────────────────────────────────────────────────── */

let activeReportDate  = ''; // Persists between tab switches
let activeReportMonth = '';

/** Switches between daily and monthly report views. */
function reportTabSwitch(mode) {
  reportMode = mode;
  document.getElementById('rView-daily').style.display   = mode === 'daily'   ? 'block' : 'none';
  document.getElementById('rView-monthly').style.display = mode === 'monthly' ? 'block' : 'none';
  document.querySelectorAll('[id^="rTab-"]').forEach(b => b.classList.remove('active'));
  document.getElementById('rTab-' + mode).classList.add('active');
  if (mode === 'daily')   renderDailyReport();
  if (mode === 'monthly') renderMonthlyReport();
}

/** Sets which scanner's records to show in reports (all / guard / sup1-3). */
function setScannerFilter(filter) {
  scannerFilter = filter;
  document.querySelectorAll('[id^="sF-"]').forEach(b => b.classList.remove('active'));
  document.getElementById('sF-' + filter).classList.add('active');
  if (reportMode === 'daily')   renderDailyReport();
  if (reportMode === 'monthly') renderMonthlyReport();
}

/**
 * Renders the daily attendance report table.
 * Filters records by the selected scanner (guard / supervisor).
 */
function renderDailyReport() {
  const date = document.getElementById('rDailyDate').value;
  if (!date) return;
  activeReportDate = date;

  const cfg  = Store.get('settings', SHIFT_DEFAULT);
  const emps = getFilteredEmployees();
  const el   = document.getElementById('dailyReportCard');

  if (!emps.length) {
    el.innerHTML = '<div class="empty"><div class="empty-icon">👥</div><p>No employees for this filter</p></div>';
    return;
  }

  const rows = emps.map(emp => {
    const h = calcAttendance(emp.id, date);
    return { emp, h };
  });

  const present    = rows.filter(r => r.h.status !== 'absent').length;
  const absent     = rows.filter(r => r.h.status === 'absent').length;
  const totalHours = rows.reduce((s, r) => s + r.h.workedHrs, 0);
  const avgHours   = present > 0 ? (totalHours / present).toFixed(1) : 0;

  el.innerHTML = `
    <div style="display:flex;gap:14px;margin-bottom:18px;flex-wrap:wrap">
      <div style="flex:1;min-width:70px;text-align:center">
        <div style="font-family:'DM Serif Display',serif;font-size:28px;color:var(--col-in)">${present}</div>
        <div style="font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;color:var(--text-muted)">Present</div>
      </div>
      <div style="flex:1;min-width:70px;text-align:center">
        <div style="font-family:'DM Serif Display',serif;font-size:28px;color:var(--col-out)">${absent}</div>
        <div style="font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;color:var(--text-muted)">Absent</div>
      </div>
      <div style="flex:1;min-width:70px;text-align:center">
        <div style="font-family:'DM Serif Display',serif;font-size:28px;color:var(--helix-navy)">${avgHours}h</div>
        <div style="font-size:10px;font-weight:600;letter-spacing:1px;text-transform:uppercase;color:var(--text-muted)">Avg Hours</div>
      </div>
    </div>
    <div class="table-wrap">
    <table class="report">
      <thead><tr>
        <th>Employee</th>
        <th>In</th>
        <th>Out</th>
        <th>Break</th>
        <th>Net Hours</th>
        <th>Status</th>
      </tr></thead>
      <tbody>
      ${rows.map(({ emp, h }) => {
        const full  = h.workedHrs >= cfg.expectedHrs;
        const hrCls = h.status === 'absent' ? '' : full ? 'hours-full' : 'hours-short';
        return `<tr>
          <td>
            <strong>${emp.name}</strong><br>
            <span style="font-size:11px;color:var(--text-muted)">${emp.id} · ${emp.dept}</span>
          </td>
          <td>${h.checkIn}</td>
          <td>${h.checkOut}</td>
          <td>${h.breakMins > 0 ? minsToHM(h.breakMins) : '—'}</td>
          <td>
            <span class="hours-num ${hrCls}">${h.status === 'absent' ? '—' : h.workedHrs.toFixed(2).replace(/\.00$/, '').replace(/0$/, '') + 'h'}</span>
            ${h.status !== 'absent' ? `
            <div class="progress-wrap">
              <div class="progress-bar"><div class="progress-fill" style="width:${h.progress}%"></div></div>
            </div>` : ''}
          </td>
          <td>${getStatusBadgeHtml(h.status)}</td>
        </tr>`;
      }).join('')}
      </tbody>
    </table></div>`;
}

/* ────────────────────────────────────────────────────────────────
   REPORTS — MONTHLY
   ──────────────────────────────────────────────────────────────── */

/**
 * Renders the monthly attendance summary table.
 * Shows days present, days absent, total & average hours per employee.
 */
function renderMonthlyReport() {
  const month = document.getElementById('rMonthPicker').value;
  if (!month) return;
  activeReportMonth = month;

  const [yr, mo]  = month.split('-').map(Number);
  const daysCount = new Date(yr, mo, 0).getDate();
  const cfg       = Store.get('settings', SHIFT_DEFAULT);
  const emps      = getFilteredEmployees();
  const el        = document.getElementById('monthlyReportCard');
  const monthName = new Date(yr, mo-1).toLocaleString('en-IN', { month: 'long', year: 'numeric' });

  if (!emps.length) {
    el.innerHTML = '<div class="empty"><div class="empty-icon">👥</div><p>No employees for this filter</p></div>';
    return;
  }

  // Build day headers — short day name + date number
  // e.g. "Mon\n1", "Tue\n2" ...
  const dayHeaders = [];
  for (let d = 1; d <= daysCount; d++) {
    const dt      = new Date(yr, mo - 1, d);
    const dayName = dt.toLocaleString('en-IN', { weekday: 'short' });
    const isSun   = dt.getDay() === 0;
    const isSat   = dt.getDay() === 6;
    dayHeaders.push({ d, dayName, isSun, isSat });
  }

  // Build data: one row per employee, one value per day
  const empRows = emps.map(emp => {
    let totalHrs = 0, daysPresent = 0, fullDays = 0;
    const dayCells = dayHeaders.map(({ d, isSun, isSat }) => {
      const date = yr + '-' + String(mo).padStart(2,'0') + '-' + String(d).padStart(2,'0');
      // Skip Sundays — show as holiday
      if (isSun) return { val: '—', cls: 'day-holiday', hrs: 0 };
      const h = calcAttendance(emp.id, date, scannerFilter);
      if (h.status === 'absent') return { val: '—', cls: 'day-absent', hrs: 0 };
      daysPresent++;
      totalHrs += h.workedHrs;
      if (h.workedHrs >= cfg.expectedHrs) fullDays++;
      const cls = h.workedHrs >= cfg.expectedHrs ? 'day-full'
                : h.workedHrs > 0               ? 'day-short'
                :                                  'day-absent';
      // Show as decimal quarters (8.0, 8.25, 8.5, 8.75) — already rounded in calcAttendance
      return { val: h.workedHrs.toFixed(2).replace(/\.00$/, '').replace(/0$/, ''), cls, hrs: h.workedHrs };
    });
    const avgHrs = daysPresent > 0 ? (totalHrs / daysPresent).toFixed(1) : '0';
    return { emp, dayCells, totalHrs: totalHrs.toFixed(1), daysPresent,
             daysAbsent: daysCount - daysPresent, avgHrs, fullDays };
  });

  // Summary totals row
  const totalPresent = empRows.reduce((s, r) => s + r.daysPresent, 0);
  const avgPresence  = empRows.length > 0 ? (totalPresent / empRows.length).toFixed(1) : 0;

  // Build column headers HTML
  const thDays = dayHeaders.map(({ d, dayName, isSun, isSat }) => {
    const wkCls = isSun ? 'style="background:#fecaca;color:#991b1b"'
                : isSat ? 'style="background:#fef3c7;color:#92400e"'
                : '';
    return `<th ${wkCls} style="min-width:36px;text-align:center;padding:6px 4px">
              <div style="font-size:9px;opacity:.7">${dayName}</div>
              <div style="font-size:12px;font-weight:700">${d}</div>
            </th>`;
  }).join('');

  // Build employee rows HTML
  const tbodyHtml = empRows.map(({ emp, dayCells, totalHrs, daysPresent, daysAbsent, avgHrs }) => {
    const cellsHtml = dayCells.map(cell => {
      const bg = cell.cls === 'day-full'    ? 'background:#d1fae5;color:#065f46'
               : cell.cls === 'day-short'   ? 'background:#fef3c7;color:#92400e'
               : cell.cls === 'day-holiday' ? 'background:#f3e8ff;color:#6b21a8'
               :                              'background:#f9fafb;color:#9ca3af';
      return `<td style="text-align:center;padding:5px 2px;font-size:11px;font-weight:600;${bg}">${cell.val}</td>`;
    }).join('');

    const absColor = daysAbsent > 5 ? 'color:#dc2626;font-weight:700' : 'color:#6b7280';
    return `<tr>
      <td style="white-space:nowrap;padding:8px 12px;position:sticky;left:0;background:#fff;z-index:2;border-right:2px solid var(--border)">
        <strong style="font-size:13px">${emp.name}</strong><br>
        <span style="font-size:10px;color:var(--text-muted)">${emp.id} · ${emp.dept}</span>
      </td>
      ${cellsHtml}
      <td style="text-align:center;padding:5px 8px;font-weight:700;color:var(--helix-navy);white-space:nowrap;border-left:2px solid var(--border)">${parseFloat(totalHrs).toFixed(2).replace(/\.00$/, '').replace(/0$/, '')}</td>
      <td style="text-align:center;padding:5px 8px;color:#059669;font-weight:700">${daysPresent}</td>
      <td style="text-align:center;padding:5px 8px;${absColor}">${daysAbsent}</td>
    </tr>`;
  }).join('');

  el.innerHTML = `
    <div style="font-size:12px;font-weight:600;color:var(--text-muted);margin-bottom:12px;text-transform:uppercase;letter-spacing:1px">
      ${monthName} &nbsp;·&nbsp; Avg Attendance: ${avgPresence} days
      &nbsp;·&nbsp; Filter: ${scannerFilter === 'all' ? 'All Scanners' : scannerFilter === 'guard' ? 'Guard Only' : scannerFilter === 'supervisor' ? 'Supervisor Only' : getSupervisorLabel(scannerFilter)}
    </div>

    <!-- Colour legend -->
    <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px;font-size:11px">
      <span style="background:#d1fae5;color:#065f46;padding:3px 8px;border-radius:4px;font-weight:600">Full day (≥${cfg.expectedHrs}h)</span>
      <span style="background:#fef3c7;color:#92400e;padding:3px 8px;border-radius:4px;font-weight:600">Short day</span>
      <span style="background:#f9fafb;color:#9ca3af;padding:3px 8px;border-radius:4px;font-weight:600">Absent</span>
      <span style="background:#f3e8ff;color:#6b21a8;padding:3px 8px;border-radius:4px;font-weight:600">Sunday</span>
    </div>

    <div class="table-wrap" style="overflow-x:auto">
    <table class="report" style="border-collapse:collapse;min-width:800px">
      <thead>
        <tr style="position:sticky;top:0;z-index:3">
          <th style="position:sticky;left:0;z-index:4;min-width:160px;text-align:left;padding:10px 12px;background:var(--helix-navy)">Employee</th>
          ${thDays}
          <th style="min-width:60px;text-align:center;background:var(--helix-navy);border-left:2px solid rgba(255,255,255,.2)">Total Hrs</th>
          <th style="min-width:44px;text-align:center;background:var(--helix-navy)">Present</th>
          <th style="min-width:44px;text-align:center;background:var(--helix-navy)">Absent</th>
        </tr>
      </thead>
      <tbody>${tbodyHtml}</tbody>
    </table>
    </div>`;
}


/* ────────────────────────────────────────────────────────────────
   REPORT FILTERING HELPERS
   ──────────────────────────────────────────────────────────────── */

/**
 * Returns employees filtered by the selected scanner.
 * If filter = 'all', return all employees.
 * If filter = 'guard', return employees checked in by guard.
 * If filter = 'supN', return employees assigned to that supervisor.
 */
/**
 * getFilteredEmployees — returns employee list based on active report filter.
 * For 'guard' / 'supervisor' filters: returns ALL employees but calcAttendance
 * will only count records scanned by that role.
 * For team filters: returns only employees in that team.
 */
function getFilteredEmployees() {
  const emps = Store.get('employees', []);
  if (scannerFilter === 'all')        return emps;
  if (scannerFilter === 'guard')      return emps; // all employees, guard-only records filtered in calc
  if (scannerFilter === 'supervisor') return emps; // all employees, supervisor-only records in calc
  // Team filter — show only employees in that team
  return emps.filter(e => e.supervisor === scannerFilter);
}

/**
 * Get attendance records for an employee on a date, optionally filtered by who scanned.
 * scannedByFilter: 'all' | 'guard' | 'supervisor' | 'teamA' | 'teamB' | 'teamC'
 */
function getAttendanceRecords(empId, date, scannedByFilter) {
  let recs = Store.get('attendance', []).filter(r => r.empId === empId && r.date === date);
  if (scannedByFilter === 'guard') {
    recs = recs.filter(r => r.scannedBy === 'guard');
  } else if (scannedByFilter === 'supervisor') {
    recs = recs.filter(r => r.scannedBy === 'supervisor');
  }
  // For team or 'all' filters, use all records regardless of scanner
  return recs;
}

/* ────────────────────────────────────────────────────────────────
   EXCEL (XLSX) EXPORT
   ──────────────────────────────────────────────────────────────── */

/**
 * Exports the daily report to a well-formatted XLSX file.
 * Uses SheetJS (xlsx library).
 */
function exportDailyXLSX() {
  const date = document.getElementById('rDailyDate').value;
  if (!date) { showToast('⚠ Select a date first'); return; }

  const cfg  = Store.get('settings', SHIFT_DEFAULT);
  const emps = getFilteredEmployees();

  // Build data rows
  const dataRows = emps.map(emp => {
    const h = calcAttendance(emp.id, date);
    return {
      'Employee ID':   emp.id,
      'Name':          emp.name,
      'Department':    emp.dept,
      'Supervisor':    emp.supervisor ? getSupervisorLabel(emp.supervisor) : 'General',
      'Check In':      h.checkIn,
      'Check Out':     h.checkOut,
      'Break (mins)':  h.status !== 'absent' ? h.breakMins : '',
      'Net Hours':     h.status !== 'absent' ? h.workedHrs : 0,
      'Expected Hrs':  cfg.expectedHrs,
      'Full Day?':     h.status !== 'absent' ? (h.workedHrs >= cfg.expectedHrs ? 'Yes' : 'No') : '—',
      'Status':        h.status === 'absent' ? 'Absent' : h.status === 'active' ? 'Active' : 'Complete'
    };
  });

  const wb = XLSX.utils.book_new();

  // ── Summary sheet ──
  const summaryData = [
    ['Helix Industries — Daily Attendance Report'],
    ['Date:', date],
    ['Shift:', cfg.start + ' – ' + cfg.end],
    ['Expected Hours:', cfg.expectedHrs + ' hrs'],
    ['Generated:', new Date().toLocaleString('en-IN')],
    ['Filter:', scannerFilter === 'all' ? 'All Scanners' : scannerFilter === 'guard' ? 'Gate Guard' : getSupervisorLabel(scannerFilter)],
    [],
    ['Present:', emps.filter(e => calcAttendance(e.id, date).status !== 'absent').length],
    ['Absent:',  emps.filter(e => calcAttendance(e.id, date).status === 'absent').length],
    ['Total:',   emps.length]
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary');

  // ── Attendance detail sheet ──
  const wsData = XLSX.utils.json_to_sheet(dataRows);

  // Column widths
  wsData['!cols'] = [
    { wch: 12 }, { wch: 22 }, { wch: 16 }, { wch: 16 },
    { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 10 }, { wch: 12 }
  ];
  XLSX.utils.book_append_sheet(wb, wsData, 'Attendance');

  // Write and download
  XLSX.writeFile(wb, `Helix_Daily_Attendance_${date}_${scannerFilter}.xlsx`);
  showToast('📊 Daily XLSX exported!');
}

/**
 * Exports the monthly report to a multi-sheet XLSX file.
 * Sheet 1: Summary | Sheet 2: Monthly totals | Sheet 3+: Day-by-day for each supervisor
 */
function exportMonthlyXLSX() {
  const month = document.getElementById('rMonthPicker').value;
  if (!month) { showToast('⚠ Select a month first'); return; }

  const [yr, mo]  = month.split('-').map(Number);
  const daysCount = new Date(yr, mo, 0).getDate();
  const cfg       = Store.get('settings', SHIFT_DEFAULT);
  const emps      = getFilteredEmployees();
  const monthName = new Date(yr, mo-1).toLocaleString('en-IN', { month: 'long', year: 'numeric' });
  const filterLabel = scannerFilter === 'all' ? 'All Scanners'
                    : scannerFilter === 'guard' ? 'Guard Only'
                    : scannerFilter === 'supervisor' ? 'Supervisor Only'
                    : getSupervisorLabel(scannerFilter);

  const wb = XLSX.utils.book_new();

  // ── Sheet 1: Pivot — Employee rows × Date columns ──
  // Header row: Employee | Dept | 1 | 2 | ... | 31 | Total Hrs | Present | Absent
  const dayNums = [];
  for (let d = 1; d <= daysCount; d++) dayNums.push(d);

  const headerRow = ['Employee', 'Emp ID', 'Department', 'Team',
    ...dayNums.map(d => {
      const dt = new Date(yr, mo-1, d);
      return dt.toLocaleString('en-IN', { weekday: 'short' }) + ' ' + d;
    }),
    'Total Hrs', 'Days Present', 'Days Absent', 'Full Days', 'Avg Hrs/Day'
  ];

  const dataRows = emps.map(emp => {
    let totalHrs = 0, daysPresent = 0, fullDays = 0;
    const dayCells = dayNums.map(d => {
      const date = yr + '-' + String(mo).padStart(2,'0') + '-' + String(d).padStart(2,'0');
      const dt   = new Date(yr, mo-1, d);
      if (dt.getDay() === 0) return 'Sun'; // Sunday
      const h = calcAttendance(emp.id, date, scannerFilter);
      if (h.status === 'absent') return '';
      daysPresent++;
      totalHrs += h.workedHrs;
      if (h.workedHrs >= cfg.expectedHrs) fullDays++;
      return h.workedHrs;
    });

    const avgHrs = daysPresent > 0 ? parseFloat((totalHrs / daysPresent).toFixed(2)) : 0;
    return [
      emp.name, emp.id, emp.dept, getSupervisorLabel(emp.supervisor || ''),
      ...dayCells,
      parseFloat(totalHrs.toFixed(2)), daysPresent, daysCount - daysPresent, fullDays, avgHrs
    ];
  });

  const ws1 = XLSX.utils.aoa_to_sheet([
    [`Helix Industries — Monthly Attendance — ${monthName} — Filter: ${filterLabel}`],
    [],
    headerRow,
    ...dataRows
  ]);

  // Set column widths: name col wide, day cols narrow
  ws1['!cols'] = [
    { wch: 22 }, { wch: 10 }, { wch: 14 }, { wch: 10 },
    ...dayNums.map(() => ({ wch: 6 })),
    { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }
  ];

  XLSX.utils.book_append_sheet(wb, ws1, 'Monthly Pivot');

  // ── Sheet 2: Summary (same as before) ──
  const summaryRows = emps.map(emp => {
    let totalHrs = 0, daysPresent = 0, fullDays = 0;
    for (let d = 1; d <= daysCount; d++) {
      const date = yr + '-' + String(mo).padStart(2,'0') + '-' + String(d).padStart(2,'0');
      if (new Date(yr, mo-1, d).getDay() === 0) continue;
      const h = calcAttendance(emp.id, date, scannerFilter);
      if (h.status !== 'absent') { daysPresent++; totalHrs += h.workedHrs; if (h.workedHrs >= cfg.expectedHrs) fullDays++; }
    }
    return {
      'Employee': emp.name, 'ID': emp.id, 'Department': emp.dept,
      'Team': getSupervisorLabel(emp.supervisor || ''),
      'Days Present': daysPresent, 'Days Absent': daysCount - daysPresent,
      'Full Days': fullDays, 'Total Hrs': parseFloat(totalHrs.toFixed(2)),
      'Avg Hrs/Day': daysPresent > 0 ? parseFloat((totalHrs/daysPresent).toFixed(2)) : 0
    };
  });
  const ws2 = XLSX.utils.json_to_sheet(summaryRows);
  ws2['!cols'] = [{wch:22},{wch:10},{wch:14},{wch:10},{wch:12},{wch:12},{wch:10},{wch:12},{wch:13}];
  XLSX.utils.book_append_sheet(wb, ws2, 'Summary');

  XLSX.writeFile(wb, `Helix_Monthly_${month}_${scannerFilter}.xlsx`);
  showToast('📊 Monthly XLSX exported!');
}


function printReport() { window.print(); }

/* ────────────────────────────────────────────────────────────────
   SETTINGS
   ──────────────────────────────────────────────────────────────── */

function loadSettings() {
  const s = Cache.get('settings', SHIFT_DEFAULT);
  document.getElementById('sShiftStart').value  = s.start       || SHIFT_DEFAULT.start;
  document.getElementById('sShiftEnd').value    = s.end         || SHIFT_DEFAULT.end;
  document.getElementById('sLunchMins').value   = s.lunchMins   ?? SHIFT_DEFAULT.lunchMins;
  document.getElementById('sExpectedHrs').value = s.expectedHrs ?? SHIFT_DEFAULT.expectedHrs;
  // Load custom team names
  const n = Cache.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  document.getElementById('sTeamA').value = n.teamA || 'Team A';
  document.getElementById('sTeamB').value = n.teamB || 'Team B';
  document.getElementById('sTeamC').value = n.teamC || 'Team C';
  // Sheets URL — show config.json URL if present (read-only hint), else show manual
  const cfgUrl  = (APP_CONFIG && APP_CONFIG.googleSheets && APP_CONFIG.googleSheets.scriptUrl) || '';
  const manUrl  = Cache.get('apiUrl', '');
  const urlEl   = document.getElementById('sheetsApiUrl');
  if (urlEl) urlEl.value = cfgUrl || manUrl || '';
  updateSheetsStatus();

  // Show config.json summary box
  renderConfigSummary();

  // Load current usernames into credentials form
  loadCredentials();
}

function saveSettings() {
  const s = {
    start:       document.getElementById('sShiftStart').value   || '09:00',
    end:         document.getElementById('sShiftEnd').value     || '18:30',
    lunchMins:   parseInt(document.getElementById('sLunchMins').value)   || 30,
    expectedHrs: parseFloat(document.getElementById('sExpectedHrs').value) || 9
  };
  Cache.set('settings', s);
  showToast('✅ Settings saved');
  if (GSheet.isConfigured()) {
    GSheet.call('saveSettings', s).catch(e => console.warn('Settings sync:', e.message));
  }
}

/** Save customized team names and refresh all UI that uses them. */
function saveTeamNames() {
  const names = {
    teamA: document.getElementById('sTeamA').value.trim() || 'Team A',
    teamB: document.getElementById('sTeamB').value.trim() || 'Team B',
    teamC: document.getElementById('sTeamC').value.trim() || 'Team C'
  };
  Cache.set('teamNames', names);
  updateTeamDropdown();
  updateReportFilterLabels();
  showToast('✅ Team names saved');
  if (GSheet.isConfigured()) {
    GSheet.call('saveTeamNames', names).catch(e => console.warn('TeamNames sync:', e.message));
  }
}

/** Returns the display name for a team ID, reading custom names from storage. */
function getTeamName(teamId) {
  const names = Store.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  const map = { teamA: names.teamA, teamB: names.teamB, teamC: names.teamC };
  return map[teamId] || teamId || '—';
}

/** Updates the employee assignment dropdown with current team names. */
function updateTeamDropdown() {
  const sel = document.getElementById('efSupervisor');
  if (!sel) return;
  const n = Store.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  sel.innerHTML =
    '<option value="">— No Team / General —</option>' +
    '<option value="teamA">' + n.teamA + '</option>' +
    '<option value="teamB">' + n.teamB + '</option>' +
    '<option value="teamC">' + n.teamC + '</option>';
}

/** Updates report filter tab labels to show custom team names. */
function updateReportFilterLabels() {
  const n = Store.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  var btnA = document.getElementById('sF-teamA');
  var btnB = document.getElementById('sF-teamB');
  var btnC = document.getElementById('sF-teamC');
  if (btnA) btnA.textContent = n.teamA;
  if (btnB) btnB.textContent = n.teamB;
  if (btnC) btnC.textContent = n.teamC;
}

function loadCredentials() {
  const users = Cache.get('users', []);
  const map = { admin: 'uAdminId', guard: 'uGuardId', supervisor: 'uSupId' };
  Object.entries(map).forEach(function([role, elId]) {
    const user = users.find(function(u) { return u.role === role; });
    const el   = document.getElementById(elId);
    if (el && user) el.value = user.id;
  });
  ['uAdminPass','uGuardPass','uSupPass'].forEach(function(id) {
    var el = document.getElementById(id); if (el) el.value = '';
  });
}

function saveAllCredentials() {
  const errEl = document.getElementById('sPassError');
  errEl.style.color = '#dc2626';
  errEl.textContent = '';

  const fields = [
    { role:'admin',      idEl:'uAdminId', passEl:'uAdminPass', label:'Admin' },
    { role:'guard',      idEl:'uGuardId', passEl:'uGuardPass', label:'Guard' },
    { role:'supervisor', idEl:'uSupId',   passEl:'uSupPass',   label:'Supervisor' }
  ];

  const updates = fields.map(function(f) {
    return {
      role:    f.role,
      label:   f.label,
      newId:   (document.getElementById(f.idEl).value   || '').trim().toLowerCase(),
      newPass: (document.getElementById(f.passEl).value || '').trim()
    };
  });

  // Validate
  for (var i = 0; i < updates.length; i++) {
    var u = updates[i];
    if (!u.newId) { errEl.textContent = u.label + ' username cannot be empty.'; return; }
    if (u.newPass && u.newPass.length < 6) {
      errEl.textContent = u.label + ' password must be at least 6 characters.'; return;
    }
  }
  var ids = updates.map(function(u) { return u.newId; });
  if (new Set(ids).size !== ids.length) {
    errEl.textContent = 'Each role must have a unique username.'; return;
  }

  // Apply
  var users = Cache.get('users', []);
  updates.forEach(function(u) {
    var user = users.find(function(usr) { return usr.role === u.role; });
    if (!user) { users.push({ id: u.newId, pass: u.newPass || 'changeme', role: u.role, name: u.label }); }
    else { user.id = u.newId; if (u.newPass) user.pass = u.newPass; }
  });
  Cache.set('users', users);

  ['uAdminPass','uGuardPass','uSupPass'].forEach(function(id) {
    var el = document.getElementById(id); if (el) el.value = '';
  });
  errEl.style.color = '#059669';
  errEl.textContent = '✅ Saved. Use new credentials on next login.';
  showToast('🔐 Credentials updated', 3000);

  // Sync to Sheets
  if (GSheet.isConfigured()) {
    GSheet.call('updateUsers', { users: users })
      .catch(function(e) { console.warn('User sync to Sheets failed:', e.message); });
  }
}


function clearAllAttendance() {
  if (!confirm('Delete ALL attendance records? This cannot be undone.')) return;
  Cache.set('attendance', []);
  showToast('🗑 All local records cleared');
  refreshDashboard();
}

/* ────────────────────────────────────────────────────────────────
   UI HELPER FUNCTIONS
   ──────────────────────────────────────────────────────────────── */

/** Today's date as YYYY-MM-DD string. */
function todayStr() {
  return new Date().toISOString().split('T')[0];
}

/** Returns HTML badge for an attendance status. */
function getStatusBadge(lastRec) {
  if (!lastRec) return '';
  const map = {
    'check-in':    '<span class="badge badge-in">In</span>',
    'break-start': '<span class="badge badge-break">Break</span>',
    'break-end':   '<span class="badge badge-in">In</span>',
    'check-out':   '<span class="badge badge-out">Out</span>'
  };
  return map[lastRec.type] || '';
}

/** Returns HTML badge for a report status string. */
function getStatusBadgeHtml(status) {
  const map = {
    'absent':     '<span class="badge badge-absent">Absent</span>',
    'active':     '<span class="badge badge-in">Active</span>',
    'complete':   '<span class="badge badge-done">Done</span>',
    'no-checkin': '<span class="badge badge-break">Incomplete</span>'
  };
  return map[status] || status;
}

/** Shows a brief toast notification. */
function showToast(msg, ms = 2600) {
  const el = document.getElementById('toastEl');
  el.textContent = msg;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), ms);
}

/** Opens a modal overlay. */
function openOverlay(id) {
  document.getElementById(id).classList.add('show');
}

/** Closes a modal overlay. */
function closeOverlay(id) {
  document.getElementById(id).classList.remove('show');
}

// Close overlays when clicking backdrop
document.querySelectorAll('.overlay').forEach(ov => {
  ov.addEventListener('click', e => {
    if (e.target === ov) ov.classList.remove('show');
  });
});

/* ────────────────────────────────────────────────────────────────
   PWA SERVICE WORKER REGISTRATION
   ──────────────────────────────────────────────────────────────── */
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

/* ═══════════════════════════════════════════════════════════════
   IMPORT PANEL TOGGLE
   ═══════════════════════════════════════════════════════════════ */
function toggleImportPanel() {
  const panel = document.getElementById('importPanel');
  const isOpen = panel.style.display !== 'none';
  panel.style.display = isOpen ? 'none' : 'block';
  // Close add form when import opens
  if (!isOpen) document.getElementById('addEmpFormCard').style.display = 'none';
  // Reset state
  document.getElementById('csvFileInput').value = '';
  document.getElementById('csvPreview').style.display = 'none';
  document.getElementById('csvImportBtn').disabled = true;
  window._csvRows = [];
}

/* ═══════════════════════════════════════════════════════════════
   CSV TEMPLATE DOWNLOAD
   Generates a sample CSV with headers + 3 example rows
   ═══════════════════════════════════════════════════════════════ */
function downloadCSVTemplate() {
  const n = Store.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  const lines = [
    'First Name,Last Name,Department,Team (A/B/C),Phone',
    'Rajan,Sharma,Production,A,9876543210',
    'Priya,,Packaging,B,',
    'Suresh,Kumar,Security,C,9123456780'
  ];
  const blob = new Blob([lines.join('\n')], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'Helix_Employee_Import_Template.csv';
  a.click();
  URL.revokeObjectURL(a.href);
  showToast('📥 Template downloaded');
}

/* ═══════════════════════════════════════════════════════════════
   CSV PREVIEW — parses file and shows preview table before import
   ═══════════════════════════════════════════════════════════════ */
function previewCSV(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(e) {
    const text = e.target.result;
    const rows = parseCSVText(text);

    if (!rows.length) {
      document.getElementById('csvErrors').textContent = 'No valid rows found in file.';
      return;
    }

    window._csvRows = rows;

    // Build preview table
    const n = Store.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
    const teamMap = { 'a': 'teamA', 'b': 'teamB', 'c': 'teamC', 'teamA': 'teamA', 'teamB': 'teamB', 'teamC': 'teamC' };

    let tableHtml = '<thead><tr>' +
      ['First Name','Last Name','Department','Team','Phone'].map(h =>
        '<th style="padding:6px 10px;background:#1e3a8a;color:#fff;font-size:11px;text-align:left;white-space:nowrap">' + h + '</th>'
      ).join('') + '</tr></thead><tbody>';

    rows.forEach(function(r, i) {
      const teamId  = teamMap[(r.team||'').toLowerCase()] || '';
      const teamLbl = teamId ? getSupervisorLabel(teamId) : '—';
      const bg = i % 2 === 0 ? '#fff' : '#f8fafc';
      tableHtml += '<tr style="background:' + bg + '">' +
        ['<strong>' + (r.firstName||'') + '</strong>', r.lastName||'—', r.dept||'General', teamLbl, r.phone||'—'].map(v =>
          '<td style="padding:6px 10px;font-size:12px;border-bottom:1px solid #e2e8f0">' + v + '</td>'
        ).join('') + '</tr>';
    });
    tableHtml += '</tbody>';

    document.getElementById('csvPreviewTable').innerHTML = tableHtml;
    document.getElementById('csvRowCount').textContent = rows.length;
    document.getElementById('csvPreview').style.display = 'block';
    document.getElementById('csvErrors').textContent = '';
    document.getElementById('csvImportBtn').disabled = false;
  };
  reader.readAsText(file);
}

/* ═══════════════════════════════════════════════════════════════
   CSV PARSER — handles quoted fields, blank cells, header row
   ═══════════════════════════════════════════════════════════════ */
function parseCSVText(text) {
  // Normalize line endings
  const lines = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const rows  = [];
  const teamMap = { 'a': 'teamA', 'b': 'teamB', 'c': 'teamC' };

  lines.forEach(function(line, idx) {
    line = line.trim();
    if (!line) return;

    // Parse CSV fields (handle quoted values with commas inside)
    const fields = [];
    let current = '', inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (ch === '"') { inQuotes = !inQuotes; }
      else if (ch === ',' && !inQuotes) { fields.push(current.trim()); current = ''; }
      else { current += ch; }
    }
    fields.push(current.trim());

    const f0 = (fields[0] || '').replace(/^"+|"+$/g, '');

    // Skip header row — if first cell looks like a heading word (not a name digit)
    if (idx === 0 && /^(first|name|employee|fname)/i.test(f0)) return;
    // Skip empty first column
    if (!f0) return;

    const firstName = f0;
    const lastName  = (fields[1] || '').replace(/^"+|"+$/g, '').trim();
    const dept      = (fields[2] || '').replace(/^"+|"+$/g, '').trim() || 'General';
    const teamRaw   = (fields[3] || '').replace(/^"+|"+$/g, '').trim().toLowerCase();
    const phone     = (fields[4] || '').replace(/^"+|"+$/g, '').trim();

    // Map team: A/B/C or teamA/teamB/teamC
    let team = '';
    if (teamRaw === 'a' || teamRaw === 'teama') team = 'teamA';
    else if (teamRaw === 'b' || teamRaw === 'teamb') team = 'teamB';
    else if (teamRaw === 'c' || teamRaw === 'teamc') team = 'teamC';

    rows.push({ firstName, lastName, dept, team, phone });
  });

  return rows;
}

/* ═══════════════════════════════════════════════════════════════
   CSV IMPORT — saves parsed rows as employees, skips exact duplicates
   ═══════════════════════════════════════════════════════════════ */
function importCSV() {
  const rows = window._csvRows || [];
  if (!rows.length) { showToast('⚠ No rows to import'); return; }

  const existing = Store.get('employees', []);
  let counter    = Store.get('empCounter', 1000);
  let added = 0, skipped = 0;

  rows.forEach(function(r) {
    // Skip if exact same first+last name already exists
    const isDupe = existing.some(function(e) {
      return e.firstName.toLowerCase() === r.firstName.toLowerCase() &&
             (e.lastName||'').toLowerCase() === (r.lastName||'').toLowerCase();
    });

    if (isDupe) { skipped++; return; }

    counter++;
    existing.push({
      id:        'EMP' + counter,
      firstName: r.firstName,
      lastName:  r.lastName,
      name:      r.lastName ? r.firstName + ' ' + r.lastName : r.firstName,
      dept:      r.dept,
      supervisor: r.team,
      phone:     r.phone,
      createdAt: new Date().toISOString()
    });
    added++;
  });

  Store.set('employees', existing);
  Store.set('empCounter', counter);

  toggleImportPanel();
  renderEmployeeList();

  const msg = '✅ Imported ' + added + ' employee' + (added !== 1 ? 's' : '') +
              (skipped ? ' · ' + skipped + ' duplicate(s) skipped' : '');
  showToast(msg, 4000);

  // Bulk sync new employees to Sheets
  if (GSheet.isConfigured() && added > 0) {
    const allEmps = Cache.get('employees', []);
    const newOnes = allEmps.slice(-added); // last N added
    GSheet.call('importEmployees', { employees: newOnes })
      .then(() => showToast('☁ Synced to Google Sheets', 2500))
      .catch(err => showToast('⚠ Sheets bulk sync failed: ' + err.message, 4000));
  }
}

/* ═══════════════════════════════════════════════════════════════
   EDIT EMPLOYEE — opens edit modal pre-filled with employee data
   ═══════════════════════════════════════════════════════════════ */
function editEmployee(empId) {
  const emps = Store.get('employees', []);
  const emp  = emps.find(function(e) { return e.id === empId; });
  if (!emp) return;

  document.getElementById('editEmpId').value    = emp.id;
  document.getElementById('editFirst').value    = emp.firstName || '';
  document.getElementById('editLast').value     = emp.lastName  || '';
  document.getElementById('editDept').value     = emp.dept      || '';
  document.getElementById('editPhone').value    = emp.phone     || '';
  document.getElementById('editEmpError').textContent = '';

  // Populate team dropdown with current custom names
  const n = Store.get('teamNames', { teamA: 'Team A', teamB: 'Team B', teamC: 'Team C' });
  const sel = document.getElementById('editTeam');
  sel.innerHTML =
    '<option value="">— No Team —</option>' +
    '<option value="teamA">' + n.teamA + '</option>' +
    '<option value="teamB">' + n.teamB + '</option>' +
    '<option value="teamC">' + n.teamC + '</option>';
  sel.value = emp.supervisor || '';

  openOverlay('editEmpOverlay');
}

/* ═══════════════════════════════════════════════════════════════
   SAVE EDIT EMPLOYEE — validates and updates employee record
   Does NOT change employee ID or attendance records
   ═══════════════════════════════════════════════════════════════ */
function saveEditEmployee() {
  const empId = document.getElementById('editEmpId').value;
  const first = document.getElementById('editFirst').value.trim();
  const last  = document.getElementById('editLast').value.trim();
  const dept  = document.getElementById('editDept').value.trim();
  const team  = document.getElementById('editTeam').value;
  const phone = document.getElementById('editPhone').value.trim();
  const errEl = document.getElementById('editEmpError');

  if (!first) { errEl.textContent = 'First name is required.'; return; }

  const emps = Store.get('employees', []);
  const idx  = emps.findIndex(function(e) { return e.id === empId; });
  if (idx === -1) { errEl.textContent = 'Employee not found.'; return; }

  // Update fields — preserve id, createdAt, and all attendance data
  emps[idx].firstName  = first;
  emps[idx].lastName   = last;
  emps[idx].name       = last ? first + ' ' + last : first;
  emps[idx].dept       = dept || 'General';
  emps[idx].supervisor = team;
  emps[idx].phone      = phone;

  // Also update empName in attendance records so reports show new name
  let att = Store.get('attendance', []);
  att = att.map(function(r) {
    if (r.empId === empId) { r.empName = emps[idx].name; r.dept = emps[idx].dept; }
    return r;
  });
  Cache.set('attendance', att);
  Cache.set('employees', emps);

  closeOverlay('editEmpOverlay');
  renderEmployeeList();
  showToast('✅ ' + emps[idx].name + ' updated');

  if (GSheet.isConfigured()) {
    GSheet.call('updateEmployee', emps[idx])
      .catch(err => showToast('⚠ Sheets sync failed: ' + err.message, 4000));
  }
}

/* ═══════════════════════════════════════════════════════════════
   GENERATE QR CANVAS — shared helper used by both popup and bulk download
   Returns a canvas element with the QR drawn on it.
   @param {object} emp - employee object
   @param {number} size - pixel size of output canvas (default 300)
   ═══════════════════════════════════════════════════════════════ */
function generateQRCanvas(emp, size) {
  size = size || 300;
  const payload = JSON.stringify({ id: emp.id, name: emp.name, dept: emp.dept });
  const qr      = qrcodegen(0, 'M');
  qr.addData(payload);
  qr.make();

  const modules = qr.getModuleCount();
  // Leave a white quiet zone of 4 modules
  const quietZone = 4;
  const totalMod  = modules + quietZone * 2;
  const scale     = Math.max(2, Math.floor(size / totalMod));
  const canvasSize = totalMod * scale;

  const canvas = document.createElement('canvas');
  canvas.width  = canvasSize;
  canvas.height = canvasSize;
  const ctx = canvas.getContext('2d');

  // White background (includes quiet zone)
  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, canvasSize, canvasSize);

  // QR modules
  ctx.fillStyle = '#2d3b4a';
  for (let r = 0; r < modules; r++) {
    for (let c = 0; c < modules; c++) {
      if (qr.isDark(r, c)) {
        ctx.fillRect((c + quietZone) * scale, (r + quietZone) * scale, scale, scale);
      }
    }
  }

  // Employee label below QR
  ctx.fillStyle = '#2d3b4a';
  ctx.font      = 'bold ' + Math.max(10, scale * 2) + 'px sans-serif';
  ctx.textAlign = 'center';

  // Extend canvas height to include label
  const labelCanvas = document.createElement('canvas');
  const labelH      = scale * 6;
  labelCanvas.width  = canvasSize;
  labelCanvas.height = canvasSize + labelH;
  const lctx = labelCanvas.getContext('2d');

  lctx.fillStyle = '#ffffff';
  lctx.fillRect(0, 0, labelCanvas.width, labelCanvas.height);
  lctx.drawImage(canvas, 0, 0);

  const fontSize = Math.max(11, Math.min(16, scale * 2));
  lctx.font      = 'bold ' + fontSize + 'px sans-serif';
  lctx.fillStyle = '#2d3b4a';
  lctx.textAlign = 'center';
  lctx.fillText(emp.name, canvasSize / 2, canvasSize + fontSize + 4);

  lctx.font      = (fontSize - 2) + 'px sans-serif';
  lctx.fillStyle = '#6b7280';
  lctx.fillText(emp.id + ' · ' + emp.dept, canvasSize / 2, canvasSize + fontSize * 2 + 6);

  return labelCanvas;
}

/* ═══════════════════════════════════════════════════════════════
   BULK QR DOWNLOAD — generates all employee QR codes and zips them
   Uses JSZip to create a downloadable ZIP file in the browser.
   ═══════════════════════════════════════════════════════════════ */
function downloadAllQR() {
  const emps = Store.get('employees', []);
  if (!emps.length) { showToast('⚠ No employees to generate QR for'); return; }

  showToast('⏳ Generating ' + emps.length + ' QR codes…', 5000);

  // Use setTimeout to let the toast render before heavy canvas work
  setTimeout(function() {
    try {
      const zip = new JSZip();
      const folder = zip.folder('Helix_QR_Codes');

      emps.forEach(function(emp) {
        const canvas   = generateQRCanvas(emp, 400);
        // Convert canvas to base64 PNG (strip data:image/png;base64, prefix)
        const dataUrl  = canvas.toDataURL('image/png');
        const base64   = dataUrl.split(',')[1];
        const filename = emp.id + '_' + (emp.name || 'Employee').replace(/\s+/g, '_') + '.png';
        folder.file(filename, base64, { base64: true });
      });

      zip.generateAsync({ type: 'blob' }).then(function(blob) {
        const url  = URL.createObjectURL(blob);
        const a    = document.createElement('a');
        a.href     = url;
        a.download = 'Helix_QR_Codes_' + new Date().toISOString().slice(0,10) + '.zip';
        a.click();
        URL.revokeObjectURL(url);
        showToast('✅ Downloaded ' + emps.length + ' QR codes as ZIP!', 3500);
      });
    } catch(err) {
      showToast('❌ ZIP failed: ' + err.message);
      console.error(err);
    }
  }, 100);
}

/* ═══════════════════════════════════════════════════════════════
   DATA EXPORT — writes all localStorage keys to a JSON file
   ═══════════════════════════════════════════════════════════════ */

/** Exports ALL data: employees, attendance, settings, team names, users */
function exportAllData() {
  const backup = {
    _version:   2,
    _exported:  new Date().toISOString(),
    _device:    'Helix Attend',
    employees:  Store.get('employees',  []),
    attendance: Store.get('attendance', []),
    settings:   Store.get('settings',   {}),
    teamNames:  Store.get('teamNames',  {}),
    empCounter: Store.get('empCounter', 1000),
    // Do NOT export users — passwords stay per-device for security
  };

  const empCount = backup.employees.length;
  const attCount = backup.attendance.length;

  downloadJSON(backup, 'Helix_AllData_' + new Date().toISOString().slice(0,10) + '.json');
  showToast('⬇ Exported ' + empCount + ' employees · ' + attCount + ' attendance records', 3500);
}

/** Exports ONLY employee list — useful for setting up a new device without attendance history */
function exportEmployeesOnly() {
  const backup = {
    _version:    2,
    _exported:   new Date().toISOString(),
    _type:       'employees-only',
    employees:   Store.get('employees',  []),
    settings:    Store.get('settings',   {}),
    teamNames:   Store.get('teamNames',  {}),
    empCounter:  Store.get('empCounter', 1000),
  };

  downloadJSON(backup, 'Helix_Employees_' + new Date().toISOString().slice(0,10) + '.json');
  showToast('⬇ Exported ' + backup.employees.length + ' employees', 3000);
}

/** Helper: triggers browser download of a JSON object as a .json file */
function downloadJSON(obj, filename) {
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

/* ═══════════════════════════════════════════════════════════════
   DATA IMPORT — reads a JSON backup and merges into localStorage
   Merge strategy:
     • Employees: merge by ID — add new ones, skip existing IDs
     • Attendance: merge by record ID — add new, skip existing
     • Settings / teamNames: overwrite only if not already set locally
     • empCounter: take the higher value to avoid ID collisions
   ═══════════════════════════════════════════════════════════════ */
function importAllData() {
  const fileInput = document.getElementById('restoreFileInput');
  const errEl     = document.getElementById('restoreError');
  errEl.textContent = '';

  if (!fileInput.files.length) {
    errEl.textContent = 'Please select a backup file first.';
    return;
  }

  const reader = new FileReader();
  reader.onload = function(e) {
    let backup;
    try {
      backup = JSON.parse(e.target.result);
    } catch(err) {
      errEl.textContent = '❌ Invalid file — could not parse JSON.';
      return;
    }

    // Basic validation
    if (!backup.employees && !backup.attendance) {
      errEl.textContent = '❌ File does not appear to be a Helix Attend backup.';
      return;
    }

    let empAdded = 0, empSkipped = 0, attAdded = 0, attSkipped = 0;

    // ── Merge employees ──
    if (Array.isArray(backup.employees)) {
      const existing   = Store.get('employees', []);
      const existingIds = new Set(existing.map(function(e) { return e.id; }));

      backup.employees.forEach(function(emp) {
        if (existingIds.has(emp.id)) {
          empSkipped++;
        } else {
          existing.push(emp);
          existingIds.add(emp.id);
          empAdded++;
        }
      });
      Store.set('employees', existing);
    }

    // ── Merge attendance ──
    if (Array.isArray(backup.attendance)) {
      const existing    = Store.get('attendance', []);
      const existingIds = new Set(existing.map(function(r) { return r.id; }));

      backup.attendance.forEach(function(rec) {
        if (!existingIds.has(rec.id)) {
          existing.push(rec);
          existingIds.add(rec.id);
          attAdded++;
        } else {
          attSkipped++;
        }
      });
      Store.set('attendance', existing);
    }

    // ── Settings: only import if local is default/empty ──
    if (backup.settings && Object.keys(backup.settings).length) {
      const local = Store.get('settings', null);
      if (!local) Store.set('settings', backup.settings);
    }

    // ── Team names: only import if not yet customised locally ──
    if (backup.teamNames && Object.keys(backup.teamNames).length) {
      const local = Store.get('teamNames', null);
      if (!local) Store.set('teamNames', backup.teamNames);
    }

    // ── empCounter: take highest to prevent ID collisions ──
    if (backup.empCounter) {
      const local = Store.get('empCounter', 1000);
      if (backup.empCounter > local) Store.set('empCounter', backup.empCounter);
    }

    // Reset file input
    fileInput.value = '';

    // Refresh UI
    renderEmployeeList();
    refreshDashboard();
    loadSettings();

    const msg = '✅ Import complete: ' +
      empAdded + ' employees added' +
      (empSkipped  ? ', ' + empSkipped  + ' already existed' : '') +
      (attAdded    ? ' · ' + attAdded + ' attendance records added' : '');
    errEl.style.color = '#059669';
    errEl.textContent  = msg;
    showToast(msg, 4500);
  };

  reader.readAsText(fileInput.files[0]);
}

/* ═══════════════════════════════════════════════════════════════
   GOOGLE SHEETS MANAGEMENT FUNCTIONS
   ═══════════════════════════════════════════════════════════════ */

/** Save the Apps Script URL and show connection status */
function saveSheetsUrl() {
  const url = document.getElementById('sheetsApiUrl').value.trim();
  const el  = document.getElementById('sheetsTestResult');
  if (!url) { el.textContent = 'Please enter a URL.'; el.style.color='#dc2626'; return; }
  if (!url.includes('script.google.com')) {
    el.textContent = '⚠ URL should be a script.google.com address.';
    el.style.color = '#d97706'; 
    // Allow saving anyway in case of custom proxy
  }
  Cache.set('apiUrl', url);
  el.textContent = '✅ URL saved. Press Test Connection to verify.';
  el.style.color = '#059669';
  updateSheetsStatus();
}

/** Clear the Sheets URL (revert to local-only mode) */
function clearSheetsUrl() {
  if (!confirm('Disconnect from Google Sheets? The app will use local storage only.')) return;
  Cache.del('apiUrl');
  document.getElementById('sheetsApiUrl').value = '';
  document.getElementById('sheetsTestResult').textContent = 'Disconnected — using local storage.';
  document.getElementById('sheetsTestResult').style.color = '#6b7280';
  updateSheetsStatus();
}

/** Test the connection by calling getAll and showing result */
async function testSheetsConnection() {
  const el = document.getElementById('sheetsTestResult');
  el.textContent = '⏳ Testing…';
  el.style.color = '#d97706';

  if (!GSheet.isConfigured()) {
    el.textContent = '⚠ No URL saved yet. Enter URL and press Save first.';
    el.style.color = '#dc2626';
    return;
  }

  try {
    const all = await GSheet.get('getAll');
    const sheetEmpCount = Array.isArray(all.employees) ? all.employees.length : 0;
    const localEmpCount = Cache.get('employees', []).length;

    el.textContent = '✅ Connected! Sheet has ' + sheetEmpCount + ' employee(s). Local cache has ' + localEmpCount + '.';
    el.style.color = '#059669';

    // SAFE MERGE: only update local cache if Sheet has MORE records than local.
    // Never overwrite local data with an empty or smaller Sheet.
    // Use Push (⬆) to send local data to Sheet, or Pull (⬇) to replace local with Sheet.
    if (Array.isArray(all.employees) && sheetEmpCount > localEmpCount) {
      Cache.set('employees', all.employees);

      // Sync empCounter to prevent duplicate IDs
      const maxId = all.employees.reduce(function(max, emp) {
        const num = parseInt((emp.id || '').replace(/[^0-9]/g, ''), 10);
        return (!isNaN(num) && num > max) ? num : max;
      }, 1000);
      if (maxId > Cache.get('empCounter', 1000)) Cache.set('empCounter', maxId);

      renderEmployeeList();
      el.textContent += ' — local cache updated from Sheet.';
    } else if (sheetEmpCount === 0 && localEmpCount > 0) {
      el.textContent += ' — Sheet is empty. Use ⬆ Push to upload your local employees.';
      el.style.color = '#d97706';
    } else if (sheetEmpCount > 0 && sheetEmpCount <= localEmpCount) {
      el.textContent += ' — local data kept (same or more records). Use ⬇ Pull to force sync from Sheet.';
    }
  } catch(err) {
    el.textContent = '❌ Connection failed: ' + err.message + ' — check URL and deployment settings.';
    el.style.color = '#dc2626';
  }
}

/** Push everything in local cache up to Sheets (one-time migration) */
async function pushLocalToSheets() {
  if (!GSheet.isConfigured()) {
    showToast('⚠ Configure Sheets URL first'); return;
  }

  const employees = Cache.get('employees', []);
  const settings  = Cache.get('settings',  {});
  const teamNames = Cache.get('teamNames', {});

  if (!confirm('Push ' + employees.length + ' employee(s) and settings to Google Sheets?')) return;

  const el = document.getElementById('sheetsTestResult');
  el.style.color = '#d97706';

  try {
    // ── Step 1: Push settings and team names (small, single call) ──
    el.textContent = '⏳ Pushing settings…';
    await GSheet.call('saveSettings', settings);
    await GSheet.call('saveTeamNames', teamNames);

    // ── Step 2: Push employees in batches of 10 to stay under URL/payload limits ──
    const BATCH = 10;
    let totalAdded = 0, totalSkipped = 0;

    for (let i = 0; i < employees.length; i += BATCH) {
      const batch = employees.slice(i, i + BATCH);
      const batchNum = Math.floor(i / BATCH) + 1;
      const totalBatches = Math.ceil(employees.length / BATCH);
      el.textContent = '⏳ Pushing employees… batch ' + batchNum + ' of ' + totalBatches +
                       ' (' + Math.min(i + BATCH, employees.length) + '/' + employees.length + ')';

      const result = await GSheet.call('importEmployees', { employees: batch });
      totalAdded   += result.added   || 0;
      totalSkipped += result.skipped || 0;
    }

    el.textContent = '✅ Done! ' + totalAdded + ' employee(s) added, ' + totalSkipped + ' already existed in Sheet.';
    el.style.color = '#059669';
    showToast('☁ All data pushed to Google Sheets!', 3500);

  } catch(err) {
    el.textContent = '❌ Push failed: ' + err.message;
    el.style.color = '#dc2626';
  }
}

/** Pull latest data from Sheets into local cache */
async function pullSheetsToLocal() {
  if (!GSheet.isConfigured()) {
    showToast('⚠ Configure Sheets URL first'); return;
  }
  const localCount = Cache.get('employees', []).length;
  if (localCount > 0) {
    if (!confirm('Pull from Sheets? This will REPLACE your ' + localCount + ' local employee(s) with whatever is in the Sheet. Make sure you have pushed your data first.')) return;
  }
  const el = document.getElementById('sheetsTestResult');
  el.textContent = '⏳ Pulling from Sheets…';
  el.style.color = '#d97706';

  try {
    const all = await GSheet.get('getAll');
    if (Array.isArray(all.employees)) {
      Cache.set('employees', all.employees);

      // Sync empCounter to prevent duplicate IDs after pull
      const maxId = all.employees.reduce(function(max, emp) {
        const num = parseInt((emp.id || '').replace(/[^0-9]/g, ''), 10);
        return (!isNaN(num) && num > max) ? num : max;
      }, 1000);
      if (maxId > Cache.get('empCounter', 1000)) Cache.set('empCounter', maxId);
    }
    if (all.settings && Object.keys(all.settings).length) {
      Cache.set('settings', { ...SHIFT_DEFAULT, ...all.settings });
    }
    if (all.teamNames) Cache.set('teamNames', all.teamNames);

    renderEmployeeList();
    loadSettings();
    refreshDashboard();

    el.textContent = '✅ Pulled ' + (all.employees||[]).length + ' employees from Sheets.';
    el.style.color = '#059669';
    showToast('⬇ Data pulled from Google Sheets!', 3000);
  } catch(err) {
    el.textContent = '❌ Pull failed: ' + err.message;
    el.style.color = '#dc2626';
  }
}

/** Update the Sheets status banner in settings */
function updateSheetsStatus() {
  const banner = document.getElementById('sheetsStatusBanner');
  if (!banner) return;

  if (GSheet.isConfigured()) {
    banner.style.background = '#d1fae5';
    banner.style.color      = '#065f46';
    banner.textContent      = '☁ Connected — all data syncs to Google Sheets';
  } else {
    banner.style.background = '#fef3c7';
    banner.style.color      = '#78350f';
    banner.textContent      = '💾 Local mode — data stored on this device only';
  }

  // Show where the URL is coming from
  const srcBanner = document.getElementById('configSourceBanner');
  if (!srcBanner) return;
  const cfgUrl = APP_CONFIG?.googleSheets?.scriptUrl;
  if (cfgUrl && cfgUrl.includes('script.google.com')) {
    srcBanner.style.display = 'block';
    srcBanner.innerHTML     = '📄 <strong>URL loaded from config.json</strong> — edit that file on Netlify to change it for all devices. Manual entry below is overridden by config.json.';
  } else if (Cache.get('apiUrl')) {
    srcBanner.style.display = 'block';
    srcBanner.innerHTML     = '⚙ <strong>URL set manually on this device</strong> — add <code>scriptUrl</code> to config.json to sync across all devices automatically.';
  } else {
    srcBanner.style.display = 'none';
  }
}

/* ═══════════════════════════════════════════════════════════════
   CONFIG.JSON SUMMARY — shown in Settings so admin can see
   exactly what values were loaded from the remote config file.
   ═══════════════════════════════════════════════════════════════ */

/**
 * Renders a read-only summary of the current config.json values
 * inside the Settings page so the admin knows what's loaded.
 */
function renderConfigSummary() {
  // Find or create the summary container inside the Sheets card
  let box = document.getElementById('configJsonSummary');
  if (!box) return; // container added in HTML below

  if (!APP_CONFIG || Object.keys(APP_CONFIG).length === 0) {
    box.innerHTML = '<span style="color:#6b7280;font-size:12px">⚠ config.json not loaded — file may be missing from deployment.</span>';
    box.style.display = 'block';
    return;
  }

  const cfg  = APP_CONFIG;
  const url  = cfg?.googleSheets?.scriptUrl || '(not set)';
  const sid  = cfg?.googleSheets?.sheetId   || '(not set)';
  const sh   = cfg?.shift || {};
  const tm   = cfg?.teams || {};
  const ver  = cfg?.version || '—';

  const urlShort = url.length > 50 ? url.slice(0,48) + '…' : url;

  box.style.display = 'block';
  box.innerHTML = `
    <div style="font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:#1e40af;margin-bottom:10px">
      📄 config.json — currently loaded values
      <span style="float:right;font-weight:400;color:#6b7280">v${ver}</span>
    </div>
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      ${row('Apps Script URL', '<code style="word-break:break-all;font-size:11px">' + urlShort + '</code>', url !== '(not set)' ? '#d1fae5' : '#fef3c7')}
      ${row('Sheet ID',        '<code>' + sid + '</code>')}
      ${row('Shift',           sh.start + ' – ' + sh.end + ' (' + sh.expectedHrs + 'h, ' + sh.lunchMins + 'min lunch)')}
      ${row('Teams',           (tm.teamA||'A') + ' · ' + (tm.teamB||'B') + ' · ' + (tm.teamC||'C'))}
    </table>
    <div style="margin-top:10px;font-size:11px;color:#6b7280;line-height:1.6">
      ✏️ To change any value: edit <strong>config.json</strong> on Netlify → <em>Deploys → your site → edit file</em> or push via Git.
      All devices pick up the new values within seconds on next load.
    </div>`;

  function row(label, val, bg) {
    return '<tr style="background:' + (bg||'transparent') + '">' +
      '<td style="padding:5px 8px;color:#374151;font-weight:600;width:36%;border-bottom:1px solid #e2e8f0">' + label + '</td>' +
      '<td style="padding:5px 8px;color:#1e293b;border-bottom:1px solid #e2e8f0">' + val + '</td></tr>';
  }
}

