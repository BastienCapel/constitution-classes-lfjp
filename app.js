/* ==========================================================================
   LFJP - Constitution des Classes 2026/2027
   Application Logic (SPA)
   ========================================================================== */

// --- Constants & Configuration ---
const NORMALIZED_LEVELS = {
  '6EME': '6ème', '6ÈME': '6ème', '6EME.': '6ème', '6E': '6ème', '6': '6ème', 'SIXIEme': '6ème',
  '5EME': '5ème', '5ÈME': '5ème', '5EME.': '5ème', '5E': '5ème', '5': '5ème', 'CINQUIEME': '5ème',
  '4EME': '4ème', '4ÈME': '4ème', '4EME.': '4ème', '4E': '4ème', '4': '4ème', 'QUATRIEME': '4ème',
  '3EME': '3ème', '3ÈME': '3ème', '3EME.': '3ème', '3E': '3ème', '3': '3ème', 'TROISIEME': '3ème',
  '2NDE': '2nde', '2ND': '2nde', '2E': '2nde', '2': '2nde', 'SECONDE': '2nde',
  '1ERE': '1ère', '1ÈRE': '1ère', '1RE': '1ère', '1': '1ère', 'PREMIERE': '1ère',
  'TERMINALE': 'Terminale', 'TERM': 'Terminale', 'TLE': 'Terminale', 'T': 'Terminale'
};

const SECONDARY_LEVELS = ['6ème', '5ème', '4ème', '3ème', '2nde', '1ère', 'Terminale'];

// Apps Script code for settings preview
const APPS_SCRIPT_SOURCE_CODE = `function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", message: "Le script Apps Script fonctionne. Utilisez POST pour envoyer des données." }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // CORS configuration
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    if (e.parameter && e.parameter.method === "OPTIONS") {
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Horodatage", 
        "Type", 
        "Niveau", 
        "ID Élèves", 
        "Noms Élèves", 
        "Auteur", 
        "Motif", 
        "Statut",
        "Saisie EDT"
      ]);
    }
    
    var action = data.action || "add";
    
    if (action === "add") {
      if (!data.type || !data.niveau || !data.ids || !data.names || !data.author) {
        throw new Error("Champs obligatoires manquants.");
      }
      var timestamp = new Date().toISOString();
      sheet.appendRow([
        timestamp,
        data.type,
        data.niveau,
        data.ids,
        data.names,
        data.author,
        data.motif || "",
        "Actif",
        false
      ]);
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "Demande enregistrée !",
        timestamp: timestamp
      })).setMimeType(ContentService.MimeType.JSON);
      
    } else if (action === "cancel") {
      var targetTimestamp = data.timestamp;
      if (!targetTimestamp) throw new Error("Horodatage manquant.");
      var lastRow = sheet.getLastRow();
      var range = sheet.getRange(2, 1, lastRow - 1, 9);
      var values = range.getValues();
      var found = false;
      
      for (var i = 0; i < values.length; i++) {
        var rowTimestamp = values[i][0];
        var isMatch = false;
        if (rowTimestamp instanceof Date) {
          isMatch = rowTimestamp.toISOString() === targetTimestamp;
        } else {
          isMatch = String(rowTimestamp) === targetTimestamp;
        }
        
        if (isMatch) {
          sheet.deleteRow(i + 2);
          found = true;
          break;
        }
      }
      if (!found) throw new Error("Demande introuvable.");
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === "updateEDT") {
      var targetTimestamp = data.timestamp;
      var edtValue = data.saisieEDT;
      if (!targetTimestamp) throw new Error("Horodatage manquant.");
      var lastRow = sheet.getLastRow();
      var range = sheet.getRange(2, 1, lastRow - 1, 9);
      var values = range.getValues();
      var found = false;
      
      for (var i = 0; i < values.length; i++) {
        var rowTimestamp = values[i][0];
        var isMatch = false;
        if (rowTimestamp instanceof Date) {
          isMatch = rowTimestamp.toISOString() === targetTimestamp;
        } else {
          isMatch = String(rowTimestamp) === targetTimestamp;
        }
        
        if (isMatch) {
          sheet.getRange(i + 2, 9).setValue(edtValue);
          found = true;
          break;
        }
      }
      if (!found) throw new Error("Demande introuvable.");
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;

// --- Application State ---
const state = {
  role: sessionStorage.getItem('lfjp_role') || null, // 'prof', 'admin', or null
  get isAdmin() { return this.role === 'admin'; },
  students: [],          // Parsed student list
  requests: [],          // Existing requests fetched from Sheets
  selectedIds: new Set(),// Currently selected student IDs
  selectedLevel: null,  // Grade level of current selection (forcing same-level restriction)
  activeTab: 'students', // Current active view
  settings: {
    sheetUrl: 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSPXEToKzwJ3Iv5PrF3HX02s7VeMdArVvq4UWrJT4d7-PFZZJUwE4jPadfekj39e4JZP_Na4n7wa_1q/pubhtml',
    scriptUrl: 'https://script.google.com/macros/s/AKfycbzUw3LUJLpelFTSTSg0qXDxVhAdLSdicB8dYv9kFuJVwWEgWgSPkOUa2GldvcjoexE_/exec',
  }
};

// --- DOM Cache ---
const DOM = {
  // Navigation
  tabs: document.querySelectorAll('.nav-tab'),
  tabContents: document.querySelectorAll('.tab-content'),
  themeToggle: document.getElementById('theme-toggle-btn'),
  connectionPill: document.getElementById('global-connection-pill'),
  
  // Students tab
  studentsSection: document.getElementById('students-section-container'),
  studentsEmptyState: document.getElementById('students-empty-state'),
  studentsGrid: document.getElementById('students-grid'),
  studentSearch: document.getElementById('student-search'),
  clearSearchBtn: document.getElementById('clear-search-btn'),
  levelChipsContainer: document.getElementById('level-chips-container'),
  emptyStateSyncBtn: document.getElementById('empty-state-sync-btn'),
  syncStudentsBtn: document.getElementById('sync-students-btn'),
  studentsCountText: document.getElementById('students-count-text'),
  
  // Requests tab
  requestsLoader: document.getElementById('requests-loader'),
  requestsEmptyState: document.getElementById('requests-empty-state'),
  requestsTableContainer: document.getElementById('requests-table-container'),
  requestsTableBody: document.getElementById('requests-table-body'),
  requestSearch: document.getElementById('request-search'),
  requestLevelFilter: document.getElementById('request-level-filter'),
  requestTypeFilter: document.getElementById('request-type-filter'),
  syncRequestsBtn: document.getElementById('sync-requests-btn'),
  statsTotal: document.getElementById('stats-total'),
  statsTogether: document.getElementById('stats-together'),
  statsSeparate: document.getElementById('stats-separate'),
  
  // Settings tab
  settingsForm: document.getElementById('settings-form'),
  inputSheetUrl: document.getElementById('input-sheet-url'),
  inputScriptUrl: document.getElementById('input-script-url'),
  testSettingsBtn: document.getElementById('test-settings-btn'),
  testResultBox: document.getElementById('test-result-box'),
  testResultIcon: document.getElementById('test-result-icon'),
  testResultText: document.getElementById('test-result-text'),
  settingsDbCount: document.getElementById('settings-db-count'),
  clearLocalDbBtn: document.getElementById('clear-local-db-btn'),
  appsScriptCodeContainer: document.getElementById('apps-script-code-container'),
  copyScriptCodeBtn: document.getElementById('copy-script-code-btn'),
  
  // Floating Action Bar
  floatingBar: document.getElementById('floating-selection-bar'),
  floatingSelectedCount: document.getElementById('floating-selected-count'),
  floatingSelectedLevel: document.getElementById('floating-selected-level'),
  floatingBtnTogether: document.getElementById('floating-btn-together'),
  floatingBtnSeparate: document.getElementById('floating-btn-separate'),
  floatingBtnClear: document.getElementById('floating-btn-clear'),
  
  // Modals
  requestModal: document.getElementById('request-modal'),
  requestForm: document.getElementById('request-form'),
  modalTitle: document.getElementById('modal-title'),
  modalTypeBadge: document.getElementById('modal-type-badge'),
  modalLevelText: document.getElementById('modal-level-text'),
  modalStudentsList: document.getElementById('modal-students-list'),
  inputAuthor: document.getElementById('input-author'),
  inputMotif: document.getElementById('input-motif'),
  cancelRequestBtn: document.getElementById('cancel-request-btn'),
  closeRequestModalBtn: document.getElementById('close-request-modal-btn'),
  submitRequestBtn: document.getElementById('submit-request-btn'),
  submitSpinner: document.getElementById('submit-spinner'),
  

  
  // Authentication & Portal
  appContainer: document.getElementById('app-container'),
  loginPortal: document.getElementById('login-portal'),
  portalLoginForm: document.getElementById('portal-login-form'),
  portalUsername: document.getElementById('portal-username'),
  portalPassword: document.getElementById('portal-password'),
  portalLoginError: document.getElementById('portal-login-error'),
  adminLoginBtn: document.getElementById('admin-login-btn'), // Used as general logout button now
  tabBtnSettings: document.getElementById('tab-btn-settings'),
  explainerVideoBtn: document.getElementById('explainer-video-btn'),
  videoModal: document.getElementById('video-modal'),
  closeVideoModalBtn: document.getElementById('close-video-modal-btn'),
  
  toastContainer: document.getElementById('toast-container')
};

// --- Helper Functions ---

/**
 * Display a Toast Notification
 */
function showToast(message, type = 'info', duration = 4000) {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  
  let emoji = 'ℹ️';
  if (type === 'success') emoji = '✅';
  if (type === 'error') emoji = '❌';
  if (type === 'warning') emoji = '⚠️';
  
  toast.innerHTML = `
    <span>${emoji} &nbsp; ${message}</span>
    <button class="toast-close-btn" aria-label="Fermer">✕</button>
  `;
  
  DOM.toastContainer.appendChild(toast);
  
  // Event listener for manual close
  toast.querySelector('.toast-close-btn').addEventListener('click', () => {
    toast.classList.add('fade-out');
    setTimeout(() => toast.remove(), 300);
  });
  
  // Auto remove
  setTimeout(() => {
    if (toast.parentElement) {
      toast.classList.add('fade-out');
      setTimeout(() => toast.remove(), 300);
    }
  }, duration);
}

/**
 * Update Authentication and Authorization UI layout depending on state.role
 */
function updateAuthUI() {
  const role = state.role;
  
  if (!role) {
    // Show login portal, hide app
    DOM.loginPortal.removeAttribute('hidden');
    DOM.appContainer.setAttribute('hidden', '');
    
    // Reset inputs
    DOM.portalUsername.value = '';
    DOM.portalPassword.value = '';
    DOM.portalLoginError.setAttribute('hidden', '');
  } else {
    // Hide login portal, show app
    DOM.loginPortal.setAttribute('hidden', '');
    DOM.appContainer.removeAttribute('hidden');
    
    // Clear/set styling class for role-based display
    if (role === 'admin') {
      DOM.appContainer.className = 'role-admin';
      DOM.tabBtnSettings.removeAttribute('hidden');
      DOM.adminLoginBtn.innerHTML = "🚪 Déconnexion (Admin)";
      DOM.adminLoginBtn.className = "btn btn-danger btn-small";
    } else {
      DOM.appContainer.className = 'role-prof';
      DOM.tabBtnSettings.setAttribute('hidden', '');
      DOM.adminLoginBtn.innerHTML = "🚪 Déconnexion (Prof)";
      DOM.adminLoginBtn.className = "btn btn-secondary btn-small";
      
      // If we are currently on the settings tab, auto-switch to students tab
      if (state.activeTab === 'settings') {
        switchTab('students');
      }
    }
  }
}

/**
 * Handle user authentication
 */
function loginUser(username, password) {
  let matchedRole = null;
  
  if (username === 'prof' && password === 'prof') {
    matchedRole = 'prof';
  } else if (username === 'admin' && password === 'admin') {
    matchedRole = 'admin';
  }
  
  if (matchedRole) {
    state.role = matchedRole;
    sessionStorage.setItem('lfjp_role', matchedRole);
    DOM.portalLoginError.setAttribute('hidden', '');
    updateAuthUI();
    showToast(`Connexion réussie (${matchedRole}) !`, "success");
    
    // Sync data since we are now authorized
    if (state.settings.sheetUrl) {
      fetchStudents();
      fetchExistingRequests();
    }
  } else {
    DOM.portalLoginError.removeAttribute('hidden');
    showToast("Identifiant ou mot de passe incorrect.", "error");
  }
}

/**
 * Log out user
 */
function logoutUser() {
  state.role = null;
  sessionStorage.removeItem('lfjp_role');
  updateAuthUI();
  showToast("Déconnexion réussie.", "info");
}

/**
 * Normalize level names (e.g. "6e" or "6ème" or "SIXieme" to standard "6ème")
 */
function normalizeLevel(levelStr) {
  if (!levelStr) return null;
  const clean = levelStr.trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, ""); // strip accents
  
  return NORMALIZED_LEVELS[clean] || null;
}

/**
 * Convert a Google Sheets URL into its CSV export equivalent
 */
function getSheetsCsvUrl(url, tabName = '') {
  if (!url) return '';
  url = url.trim();
  
  // Case A: Published Google Sheets link (pubhtml or pub)
  if (url.includes('/pubhtml') || url.includes('/pub')) {
    // Extract the publication ID
    const match = url.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      let csvUrl = `https://docs.google.com/spreadsheets/d/e/${match[1]}/pub?output=csv`;
      if (tabName) {
        csvUrl += `&sheet=${encodeURIComponent(tabName)}`;
      }
      return csvUrl;
    }
  }
  
  // Case B: Standard Sheets editor URL
  if (url.includes('/spreadsheets/d/')) {
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (match && match[1]) {
      let csvUrl = `https://docs.google.com/spreadsheets/d/${match[1]}/export?format=csv`;
      if (tabName) {
        csvUrl += `&sheet=${encodeURIComponent(tabName)}`;
      } else {
        // Get sheet ID (gid) if present, otherwise default to first sheet
        const gidMatch = url.match(/gid=([0-9]+)/);
        const gid = gidMatch ? gidMatch[1] : '0';
        csvUrl += `&gid=${gid}`;
      }
      return csvUrl;
    }
  }
  
  return url; // return as-is if no match
}

/**
 * Parse CSV Data safely (supports commas and semicolons, handles double quotes)
 */
function parseCSV(csvText) {
  if (!csvText) return [];
  
  // Auto-detect separator: check first line (up to first newline)
  const firstLineEnd = csvText.indexOf('\n');
  const firstLine = firstLineEnd === -1 ? csvText : csvText.substring(0, firstLineEnd);
  const commas = (firstLine.match(/,/g) || []).length;
  const semicolons = (firstLine.match(/;/g) || []).length;
  const separator = semicolons > commas ? ';' : ',';
  
  const results = [];
  let currentRow = [];
  let currentField = '';
  let insideQuote = false;
  
  for (let i = 0; i < csvText.length; i++) {
    const char = csvText[i];
    const nextChar = csvText[i + 1];
    
    if (char === '"') {
      if (insideQuote && nextChar === '"') {
        // Double quote inside quoted field -> escaped quote
        currentField += '"';
        i++; // skip next quote
      } else {
        // Toggle quote state
        insideQuote = !insideQuote;
      }
    } else if (char === separator && !insideQuote) {
      currentRow.push(currentField.trim());
      currentField = '';
    } else if ((char === '\r' || char === '\n') && !insideQuote) {
      // End of row
      if (char === '\r' && nextChar === '\n') {
        i++; // skip \n
      }
      currentRow.push(currentField.trim());
      // Only add non-empty rows
      if (currentRow.length > 1 || (currentRow.length === 1 && currentRow[0] !== '')) {
        results.push(currentRow);
      }
      currentRow = [];
      currentField = '';
    } else {
      currentField += char;
    }
  }
  
  // Push last field/row if any
  if (currentField !== '' || currentRow.length > 0) {
    currentRow.push(currentField.trim());
    if (currentRow.length > 1 || (currentRow.length === 1 && currentRow[0] !== '')) {
      results.push(currentRow);
    }
  }
  
  return results;
}

// --- Local Storage Management ---
const Storage = {
  load() {
    // Load students
    try {
      const storedStudents = localStorage.getItem('lfjp_students');
      state.students = storedStudents ? JSON.parse(storedStudents) : [];
      // Assurer la présence de uniqueId pour la rétrocompatibilité du cache
      state.students.forEach(s => {
        if (!s.uniqueId) s.uniqueId = `${s.id}#${s.name}`;
      });
    } catch (e) {
      console.error('Error loading students from LocalStorage', e);
      state.students = [];
    }
    
    // Load settings
    try {
      const storedSettings = localStorage.getItem('lfjp_settings');
      if (storedSettings) {
        state.settings = JSON.parse(storedSettings);
      } else {
        // Enregistrer les paramètres par défaut
        Storage.saveSettings();
      }
    } catch (e) {
      console.error('Error loading settings from LocalStorage', e);
    }
  },
  
  saveStudents() {
    localStorage.setItem('lfjp_students', JSON.stringify(state.students));
    DOM.settingsDbCount.textContent = state.students.length;
  },
  
  saveSettings() {
    localStorage.setItem('lfjp_settings', JSON.stringify(state.settings));
  }
};

// --- UI Rendering Controllers ---

/**
 * Handle Tab Navigation switching
 */
function switchTab(tabId) {
  state.activeTab = tabId;
  
  // Update nav buttons
  DOM.tabs.forEach(tab => {
    if (tab.dataset.tab === tabId) {
      tab.classList.add('active');
    } else {
      tab.classList.remove('active');
    }
  });
  
  // Update view contents
  DOM.tabContents.forEach(content => {
    if (content.id === `tab-${tabId}`) {
      content.removeAttribute('hidden');
      content.classList.add('active');
    } else {
      content.setAttribute('hidden', '');
      content.classList.remove('active');
    }
  });

  // Special logic on tab activation
  if (tabId === 'requests') {
    fetchExistingRequests();
  }
}

/**
 * Render level chips filtering bar
 */
let activeLevelFilter = '6ème';

function renderLevelChips() {
  const chips = DOM.levelChipsContainer.querySelectorAll('.level-chip');
  chips.forEach(chip => {
    chip.addEventListener('click', () => {
      chips.forEach(c => c.classList.remove('active'));
      chip.classList.add('active');
      activeLevelFilter = chip.dataset.level;
      renderStudentsGrid();
    });
  });
}

/**
 * Build and Render Students Grid
 */
function renderStudentsGrid() {
  const searchQuery = DOM.studentSearch.value.trim().toLowerCase();
  
  // Filter students based on level and search query
  const filtered = state.students.filter(student => {
    const matchesLevel = student.level === activeLevelFilter;
    const matchesSearch = !searchQuery || 
      student.name.toLowerCase().includes(searchQuery) || 
      student.id.toLowerCase().includes(searchQuery);
    return matchesLevel && matchesSearch;
  });
  
  DOM.studentsCountText.textContent = `${filtered.length} élève${filtered.length > 1 ? 's' : ''} trouvé${filtered.length > 1 ? 's' : ''}`;
  DOM.clearSearchBtn.hidden = searchQuery.length === 0;

  if (state.students.length === 0) {
    DOM.studentsSection.setAttribute('hidden', '');
    DOM.studentsEmptyState.removeAttribute('hidden');
    return;
  }
  
  DOM.studentsEmptyState.setAttribute('hidden', '');
  DOM.studentsSection.removeAttribute('hidden');
  
  // Build grid items HTML
  DOM.studentsGrid.innerHTML = '';
  
  if (filtered.length === 0) {
    DOM.studentsGrid.innerHTML = `
      <div class="empty-state-card" style="grid-column: 1/-1; margin: 20px 0; padding: 40px;">
        <div style="font-size: 2rem;">🔍</div>
        <h3>Aucun élève trouvé</h3>
        <p>Essayez de modifier votre recherche ou le niveau sélectionné.</p>
      </div>
    `;
    return;
  }
  
  filtered.forEach(student => {
    const isSelected = state.selectedIds.has(student.uniqueId);
    const card = document.createElement('div');
    card.className = `student-card ${isSelected ? 'selected' : ''}`;
    card.dataset.id = student.id;
    card.dataset.uniqueId = student.uniqueId;
    card.dataset.level = student.level;
    
    // Dim cards of different levels if a selection is active
    if (state.selectedLevel && state.selectedLevel !== student.level) {
      card.style.opacity = '0.4';
    }
    
    // Determine payment class/badge
    const paymentText = student.payment || 'Payé';
    const isPaid = paymentText.toLowerCase().includes('payé') && !paymentText.toLowerCase().includes('attente');
    const paymentBadgeClass = isPaid ? 'badge-payment' : 'badge-payment unpaid';
    
    card.innerHTML = `
      <div class="card-checkbox-wrapper">
        <div class="custom-checkbox">
          <span class="custom-checkbox-icon">✓</span>
        </div>
      </div>
      <div class="student-info-main">
        <h4 class="student-name" style="margin-top: 0;">${student.name}</h4>
        <div class="student-badges">
          <span class="badge badge-level">${student.level}</span>
        </div>
      </div>
    `;
    
    // Click Handler
    card.addEventListener('click', (e) => {
      // Prevent click triggering twice if clicking the checkbox wrapper
      toggleStudentSelection(student.uniqueId, student.level);
    });
    
    DOM.studentsGrid.appendChild(card);
  });
}

/**
 * Handle student card selection logic (enforcing same-level pairing)
 */
function toggleStudentSelection(studentId, studentLevel) {
  // If there's already a selection and it belongs to a different level, prevent or reset
  if (state.selectedLevel && state.selectedLevel !== studentLevel) {
    showToast(`Vous ne pouvez regrouper ou éloigner que des élèves du même niveau (${state.selectedLevel}).`, 'warning');
    return;
  }
  
  if (state.selectedIds.has(studentId)) {
    state.selectedIds.delete(studentId);
  } else {
    state.selectedIds.add(studentId);
  }
  
  // Update active level of current selection
  if (state.selectedIds.size > 0) {
    state.selectedLevel = studentLevel;
  } else {
    state.selectedLevel = null;
  }
  
  // Refresh UI Grid and Floating bar
  renderStudentsGrid();
  updateFloatingSelectionBar();
}

/**
 * Update the bottom Floating Action Bar
 */
function updateFloatingSelectionBar() {
  const count = state.selectedIds.size;
  
  if (count === 0) {
    DOM.floatingBar.setAttribute('hidden', '');
    return;
  }
  
  DOM.floatingSelectedCount.textContent = count;
  DOM.floatingSelectedLevel.textContent = state.selectedLevel;
  DOM.floatingBar.removeAttribute('hidden');
  
  // Disable "Regrouper / Éloigner" buttons if only 1 student is selected
  // (Pédagogiquement: Un regroupement ou éloignement concerne au moins 2 élèves)
  const isActionDisabled = count < 2;
  DOM.floatingBtnTogether.disabled = isActionDisabled;
  DOM.floatingBtnSeparate.disabled = isActionDisabled;
  
  if (isActionDisabled) {
    DOM.floatingBtnTogether.title = "Sélectionnez au moins 2 élèves";
    DOM.floatingBtnSeparate.title = "Sélectionnez au moins 2 élèves";
    DOM.floatingBtnTogether.style.opacity = '0.5';
    DOM.floatingBtnSeparate.style.opacity = '0.5';
    DOM.floatingBtnTogether.style.cursor = 'not-allowed';
    DOM.floatingBtnSeparate.style.cursor = 'not-allowed';
  } else {
    DOM.floatingBtnTogether.title = "";
    DOM.floatingBtnSeparate.title = "";
    DOM.floatingBtnTogether.style.opacity = '1';
    DOM.floatingBtnSeparate.style.opacity = '1';
    DOM.floatingBtnTogether.style.cursor = 'pointer';
    DOM.floatingBtnSeparate.style.cursor = 'pointer';
  }
}

/**
 * Reset all student selections
 */
function clearSelection() {
  state.selectedIds.clear();
  state.selectedLevel = null;
  renderStudentsGrid();
  updateFloatingSelectionBar();
}



// --- Google Sheets Students Fetching ---

/**
 * Attempt to find the GID of a tab name in a published Google Sheet HTML
 */
async function findPublishedGid(sheetUrl, tabName) {
  if (!sheetUrl) return null;
  
  // Convert pub URL to pubhtml URL if needed
  let pubhtmlUrl = sheetUrl.trim();
  if (pubhtmlUrl.includes('/pub') && !pubhtmlUrl.includes('/pubhtml')) {
    pubhtmlUrl = pubhtmlUrl.replace(/\/pub(\?.*)?$/, '/pubhtml');
  }
  
  if (!pubhtmlUrl.includes('/pubhtml')) {
    return null; // Not a published sheet url
  }
  
  try {
    const response = await fetch(`${pubhtmlUrl}${pubhtmlUrl.includes('?') ? '&' : '?'}_cb=${new Date().getTime()}`);
    if (!response.ok) return null;
    
    const htmlText = await response.text();
    // Escape special characters in tab name for regex
    const escapedTab = tabName.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
    const regex = new RegExp(`name:\\s*['"]${escapedTab}['"]\\s*,.*?gid:\\s*['"]([0-9]+)['"]`, 'i');
    const match = htmlText.match(regex);
    if (match && match[1]) {
      return match[1];
    }
  } catch (e) {
    console.warn("Failed to parse pubhtml for GID", e);
  }
  return null;
}

/**
 * Fetch students from the linked Google Sheets ("ELEVES" tab)
 */
async function fetchStudents() {
  if (!state.settings.sheetUrl) {
    DOM.studentsSection.setAttribute('hidden', '');
    DOM.studentsEmptyState.removeAttribute('hidden');
    updateConnectionStatus(false, "Google Sheets non configuré");
    return;
  }
  
  showToast("Synchronisation des élèves...", "info");
  
  const hadNoStudents = state.students.length === 0;
  if (hadNoStudents) {
    DOM.studentsEmptyState.setAttribute('hidden', '');
    DOM.studentsSection.removeAttribute('hidden');
    DOM.studentsGrid.innerHTML = `
      <div style="grid-column: 1/-1; padding: 40px; text-align: center;">
        <div class="spinner" style="margin: 0 auto 15px;"></div>
        <p>Chargement des élèves depuis Google Sheets...</p>
      </div>
    `;
  }
  
  let csvUrl = getSheetsCsvUrl(state.settings.sheetUrl, 'ELEVES');
  
  try {
    // If it's a published Google Sheet, dynamically resolve the GID for the 'ELEVES' tab
    if (state.settings.sheetUrl.includes('/pubhtml') || state.settings.sheetUrl.includes('/pub')) {
      const gid = await findPublishedGid(state.settings.sheetUrl, 'ELEVES');
      if (gid) {
        const match = state.settings.sheetUrl.match(/\/d\/e\/([a-zA-Z0-9-_]+)/);
        if (match && match[1]) {
          csvUrl = `https://docs.google.com/spreadsheets/d/e/${match[1]}/pub?output=csv&gid=${gid}`;
        }
      }
    }
    
    const response = await fetch(`${csvUrl}&_cb=${new Date().getTime()}`);
    if (!response.ok) {
      throw new Error(`Erreur HTTP : ${response.status}`);
    }
    
    const csvText = await response.text();
    const rows = parseCSV(csvText);
    
    if (rows.length <= 1) {
      throw new Error("L'onglet 'ELEVES' semble vide ou invalide.");
    }
    
    // Find column indexes
    const headers = rows[0];
    const findHeaderIndex = (keywords) => {
      return headers.findIndex(h => {
        if (!h) return false;
        const cleanHeader = h.toLowerCase().trim();
        return keywords.some(k => cleanHeader.includes(k.toLowerCase()));
      });
    };
    
    const idxId = findHeaderIndex(['id dossier', 'id_dossier', 'dossier id', 'identifiant']);
    const idxName = findHeaderIndex(["nom de l'eleve", 'nom de l\'élève', 'nom eleve', 'nom', 'eleve']);
    const idxLevel = findHeaderIndex(['classe / niveau', 'classe/niveau', 'classe', 'niveau']);
    const idxPayment = findHeaderIndex(['paiement', 'paye', 'facturation']);
    const idxStatus = findHeaderIndex(['macro-statut', 'statut', 'macro statut', 'reinscription', 'réinscription']);
    
    if (idxId === -1 || idxName === -1 || idxLevel === -1) {
      throw new Error("Colonnes requises manquantes dans l'onglet 'ELEVES'. Assurez-vous d'avoir : 'ID Dossier', 'Nom de l'élève' et 'Classe / Niveau'.");
    }
    
    const importedStudents = [];
    let skippedCount = 0;
    
    // Parse data rows
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (row.length < 3 || !row[idxId] || !row[idxName]) continue;
      
      const rawLevel = row[idxLevel];
      const normalizedLevel = normalizeLevel(rawLevel);
      
      // Filter secondary level classes only
      if (normalizedLevel && SECONDARY_LEVELS.includes(normalizedLevel)) {
        const studentId = row[idxId].trim();
        const studentName = row[idxName].trim();
        importedStudents.push({
          id: studentId,
          name: studentName,
          uniqueId: `${studentId}#${studentName}`,
          level: normalizedLevel,
          payment: (idxPayment !== -1 && row[idxPayment]) ? row[idxPayment].trim() : 'Payé',
          status: (idxStatus !== -1 && row[idxStatus]) ? row[idxStatus].trim() : 'Réinscrit'
        });
      } else {
        skippedCount++;
      }
    }
    
    if (importedStudents.length === 0) {
      throw new Error("Aucun élève du secondaire (6ème à Terminale) n'a été trouvé dans cet onglet.");
    }
    
    // Save state
    state.students = importedStudents;
    Storage.saveStudents();
    
    showToast(`${importedStudents.length} élèves synchronisés avec succès depuis Google Sheets !`, 'success');
    
    // Refresh UI
    clearSelection();
    renderStudentsGrid();
    
  } catch (err) {
    console.error('Error fetching students:', err);
    showToast(`Échec de synchronisation des élèves : ${err.message}`, 'error');
    if (hadNoStudents) {
      DOM.studentsSection.setAttribute('hidden', '');
      DOM.studentsEmptyState.removeAttribute('hidden');
    } else {
      renderStudentsGrid();
    }
  }
}

// --- Google Sheets Integration API ---

/**
 * Fetch existing requests from Google Sheets (published CSV export)
 */
async function fetchExistingRequests() {
  if (!state.settings.sheetUrl) {
    DOM.requestsLoader.setAttribute('hidden', '');
    DOM.requestsTableContainer.setAttribute('hidden', '');
    DOM.requestsEmptyState.removeAttribute('hidden');
    updateConnectionStatus(false, "Google Sheets non configuré");
    return;
  }
  
  DOM.requestsLoader.removeAttribute('hidden');
  DOM.requestsEmptyState.setAttribute('hidden', '');
  DOM.requestsTableContainer.setAttribute('hidden', '');
  
  const csvUrl = getSheetsCsvUrl(state.settings.sheetUrl);
  
  try {
    // Add cache buster to prevent cached responses
    const response = await fetch(`${csvUrl}&_cb=${new Date().getTime()}`);
    
    if (!response.ok) {
      throw new Error(`Erreur lors du téléchargement : HTTP ${response.status}`);
    }
    
    const csvText = await response.text();
    const rows = parseCSV(csvText);
    
    if (rows.length <= 1) {
      state.requests = [];
    } else {
      const headers = rows[0];
      
      // Map columns indexes
      const getColIndex = (name) => headers.findIndex(h => h.trim().toLowerCase() === name.toLowerCase());
      
      const idxTime = getColIndex("Horodatage");
      const idxType = getColIndex("Type");
      const idxLevel = getColIndex("Niveau");
      const idxIds = getColIndex("ID Élèves");
      const idxNames = getColIndex("Noms Élèves");
      const idxAuthor = getColIndex("Auteur");
      const idxMotif = getColIndex("Motif");
      const idxStatus = getColIndex("Statut");
      let idxSaisieEDT = getColIndex("Saisie EDT");
      if (idxSaisieEDT === -1 && headers.length >= 9) {
        idxSaisieEDT = 8; // Fallback to 9th column if the header is empty/unnamed
      }
      
      // Parse requests
      state.requests = rows.slice(1).map(row => {
        return {
          timestamp: (idxTime !== -1 && row[idxTime]) ? row[idxTime].trim() : '',
          type: (idxType !== -1 && row[idxType]) ? row[idxType].trim() : 'Regroupement',
          niveau: (idxLevel !== -1 && row[idxLevel]) ? row[idxLevel].trim() : '',
          ids: (idxIds !== -1 && row[idxIds]) ? row[idxIds].trim() : '',
          names: (idxNames !== -1 && row[idxNames]) ? row[idxNames].trim() : '',
          author: (idxAuthor !== -1 && row[idxAuthor]) ? row[idxAuthor].trim() : '',
          motif: (idxMotif !== -1 && row[idxMotif]) ? row[idxMotif].trim() : '',
          status: (idxStatus !== -1 && row[idxStatus]) ? row[idxStatus].trim() : 'Actif',
          saisieEDT: (idxSaisieEDT !== -1 && row[idxSaisieEDT]) ? row[idxSaisieEDT].trim() : ''
        };
      }).filter(req => req.timestamp); // keep only rows with timestamps
      
      // Reverse requests list to show newest first
      state.requests.reverse();
    }
    
    renderRequestsUI();
    updateConnectionStatus(true, "Connecté à Google Sheets");
    
  } catch (err) {
    console.error('Error fetching sheet data', err);
    DOM.requestsLoader.setAttribute('hidden', '');
    DOM.requestsEmptyState.removeAttribute('hidden');
    DOM.requestsEmptyState.querySelector('p').innerHTML = `
      Impossible de charger les données depuis le Google Sheet.<br>
      <small style="color:var(--separate-color)">${err.message}</small><br><br>
      Vérifiez que le document est bien publié sur le Web (Fichier > Partager > Publier sur le web) au format Page Web et que l'URL est correcte.
    `;
    updateConnectionStatus(false, "Erreur de synchronisation");
    showToast("Échec de synchronisation avec Google Sheets.", "error");
  }
}

/**
 * Render Requests tab contents (table list & counters)
 */
function renderRequestsUI() {
  DOM.requestsLoader.setAttribute('hidden', '');
  
  const searchVal = DOM.requestSearch.value.trim().toLowerCase();
  const levelVal = DOM.requestLevelFilter.value;
  const typeVal = DOM.requestTypeFilter.value;
  
  // Stats Counters (on all requests, active only)
  const activeRequests = state.requests.filter(r => r.status.toLowerCase() !== 'annulé');
  const countTotal = activeRequests.length;
  const countTogether = activeRequests.filter(r => r.type === 'Regroupement').length;
  const countSeparate = activeRequests.filter(r => r.type === 'Éloignement').length;
  
  DOM.statsTotal.textContent = countTotal;
  DOM.statsTogether.textContent = countTogether;
  DOM.statsSeparate.textContent = countSeparate;
  
  // Filtered requests list
  const filtered = state.requests.filter(req => {
    // Completely hide cancelled requests from the table
    if (req.status.toLowerCase() === 'annulé') return false;

    const matchesLevel = levelVal === 'all' || req.niveau === levelVal;
    const matchesType = typeVal === 'all' || req.type === typeVal;
    
    const matchesSearch = !searchVal || 
      req.names.toLowerCase().includes(searchVal) ||
      req.ids.toLowerCase().includes(searchVal) ||
      req.author.toLowerCase().includes(searchVal) ||
      req.motif.toLowerCase().includes(searchVal);
      
    return matchesLevel && matchesType && matchesSearch;
  });
  
  if (filtered.length === 0) {
    DOM.requestsTableContainer.setAttribute('hidden', '');
    DOM.requestsEmptyState.removeAttribute('hidden');
    return;
  }
  
  DOM.requestsEmptyState.setAttribute('hidden', '');
  DOM.requestsTableContainer.removeAttribute('hidden');
  
  DOM.requestsTableBody.innerHTML = '';
  
  filtered.forEach(req => {
    const isCancelled = req.status.toLowerCase() === 'annulé';
    const row = document.createElement('tr');
    if (isCancelled) row.className = 'cancelled';
    
    // Parse students lists
    const ids = req.ids.split(';');
    const names = req.names.split(';');
    
    let studentsHTML = '<ul class="table-students-cell-list">';
    for (let idx = 0; idx < names.length; idx++) {
      if (!names[idx]) continue;
      studentsHTML += `
        <li class="table-student-item">
          <span class="table-student-id">${ids[idx] || ''}</span>
          <span>${names[idx]}</span>
        </li>
      `;
    }
    studentsHTML += '</ul>';
    
    // Format timestamp
    let formattedDate = req.timestamp;
    try {
      const date = new Date(req.timestamp);
      if (!isNaN(date.getTime())) {
        formattedDate = date.toLocaleString('fr-FR', {
          day: 'numeric',
          month: 'short',
          hour: '2-digit',
          minute: '2-digit'
        });
      }
    } catch(e) {}
    
    const typeBadgeClass = req.type === 'Regroupement' ? 'badge-type-together' : 'badge-type-separate';
    const statusBadgeClass = isCancelled ? 'badge-status-cancelled' : 'badge-status-active';
    const statusLabel = isCancelled ? 'Annulé' : 'Actif';
    
    // Disable action button if cancelled or no script URL configured
    const cancelActionHTML = isCancelled
      ? `<span style="font-size:0.8rem; color:var(--text-muted)">Annulé</span>`
      : `<button class="btn btn-danger btn-small cancel-req-btn" data-timestamp="${req.timestamp}">✕ Annuler</button>`;
      
    const edtChecked = String(req.saisieEDT).toLowerCase() === 'true';
    const edtCheckboxHTML = `
      <td class="admin-only" style="text-align: center;">
        <input type="checkbox" class="edt-checkbox" data-timestamp="${req.timestamp}" ${edtChecked ? 'checked' : ''} ${isCancelled ? 'disabled' : ''}>
      </td>
    `;
      
    row.innerHTML = `
      <td><span style="font-size: 0.8rem; white-space: nowrap;">${formattedDate}</span></td>
      <td><span class="badge ${typeBadgeClass}">${req.type}</span></td>
      <td><span class="badge badge-level">${req.niveau}</span></td>
      <td>${studentsHTML}</td>
      <td><strong>${req.author}</strong></td>
      <td><div style="max-width: 250px; font-size:0.82rem; line-height: 1.4;">${req.motif}</div></td>
      <td><span class="badge ${statusBadgeClass}">${statusLabel}</span></td>
      ${edtCheckboxHTML}
      <td class="col-action">${cancelActionHTML}</td>
    `;
    
    // Bind cancellation click handler
    const cancelBtn = row.querySelector('.cancel-req-btn');
    if (cancelBtn) {
      cancelBtn.addEventListener('click', () => {
        if (confirm(`Voulez-vous vraiment annuler cette demande de ${req.type} faite par ${req.author} ?`)) {
          cancelRequest(req.timestamp);
        }
      });
    }
    
    // Bind EDT checkbox change handler
    const edtCheckbox = row.querySelector('.edt-checkbox');
    if (edtCheckbox) {
      edtCheckbox.addEventListener('change', (e) => {
        updateEdtStatus(req.timestamp, e.target.checked);
      });
    }
    
    DOM.requestsTableBody.appendChild(row);
  });
}

/**
 * Update Header Status Pill
 */
function updateConnectionStatus(isOnline, text) {
  if (isOnline) {
    DOM.connectionPill.classList.remove('offline');
    DOM.connectionPill.classList.add('online');
  } else {
    DOM.connectionPill.classList.remove('online');
    DOM.connectionPill.classList.add('offline');
  }
  DOM.connectionPill.querySelector('.pill-text').textContent = text;
}

/**
 * POST a new request using Google Apps Script Web App
 */
async function submitRequest(type, niveau, ids, names, author, motif) {
  if (!state.settings.scriptUrl) {
    showToast("Veuillez d'abord configurer l'URL du script dans l'onglet Configuration.", "warning");
    switchTab('settings');
    return;
  }
  
  DOM.submitSpinner.removeAttribute('hidden');
  DOM.submitRequestBtn.disabled = true;
  
  const payload = {
    action: 'add',
    type,
    niveau,
    ids: ids.join(';'),
    names: names.join(';'),
    author,
    motif
  };
  
  try {
    // We send payload as a plain text JSON to prevent preflight OPTIONS request triggers (CORS)
    const response = await fetch(state.settings.scriptUrl, {
      method: 'POST',
      mode: 'cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });
    
    const result = await response.json();
    
    if (result.status === 'success') {
      showToast("Vœu enregistré avec succès dans Google Sheets !", "success");
      
      // Reset selection and close modal
      clearSelection();
      DOM.requestModal.close();
      
      // Auto redirect and refresh requests tab
      switchTab('requests');
    } else {
      throw new Error(result.message || "Erreur inconnue renvoyée par le script.");
    }
  } catch (err) {
    console.error('Apps Script write error', err);
    showToast(`Erreur d'enregistrement : ${err.message}. Vérifiez la configuration de votre script.`, "error");
  } finally {
    DOM.submitSpinner.setAttribute('hidden', '');
    DOM.submitRequestBtn.disabled = false;
  }
}

/**
 * POST a cancellation request using Google Apps Script Web App
 */
async function cancelRequest(timestamp) {
  if (!state.settings.scriptUrl) {
    showToast("URL de script non configurée.", "error");
    return;
  }
  
  showToast("Annulation en cours...", "info");
  
  const payload = {
    action: 'cancel',
    timestamp: timestamp
  };
  
  try {
    const response = await fetch(state.settings.scriptUrl, {
      method: 'POST',
      mode: 'cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });
    
    const result = await response.json();
    
    if (result.status === 'success') {
      showToast("Demande annulée avec succès.", "success");
      // Refresh requests
      fetchExistingRequests();
    } else {
      throw new Error(result.message || "Erreur lors de l'annulation.");
    }
  } catch (err) {
    console.error(err);
    showToast(`Erreur d'annulation : ${err.message}`, "error");
  }
}

/**
 * POST an EDT status update using Google Apps Script Web App
 */
async function updateEdtStatus(timestamp, isChecked) {
  if (!state.settings.scriptUrl) {
    showToast("URL de script non configurée.", "error");
    return;
  }
  
  showToast("Mise à jour EDT...", "info");
  
  const payload = {
    action: 'updateEDT',
    timestamp: timestamp,
    saisieEDT: isChecked
  };
  
  try {
    const response = await fetch(state.settings.scriptUrl, {
      method: 'POST',
      mode: 'cors',
      headers: {
        'Content-Type': 'text/plain;charset=utf-8'
      },
      body: JSON.stringify(payload)
    });
    
    const result = await response.json();
    
    if (result.status === 'success') {
      showToast("Statut EDT mis à jour avec succès.", "success");
      // Update local state to maintain correct checkbox state without reloading from CSV (lag)
      const request = state.requests.find(r => r.timestamp === timestamp);
      if (request) {
        request.saisieEDT = isChecked ? 'true' : 'false';
      }
    } else {
      throw new Error(result.message || "Erreur lors de la mise à jour EDT.");
    }
  } catch (err) {
    console.error(err);
    showToast(`Erreur de mise à jour EDT : ${err.message}`, "error");
    // Revert checkbox state by rendering the requests list again
    renderRequestsUI();
  }
}

// --- Event Bindings & Init ---

function initTheme() {
  // Load saved theme or fall back to system preference
  const savedTheme = localStorage.getItem('lfjp_theme');
  if (savedTheme === 'dark') {
    document.body.className = 'theme-dark';
  } else if (savedTheme === 'light') {
    document.body.className = 'theme-light';
  } else {
    // System match
    const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    document.body.className = prefersDark ? 'theme-dark' : 'theme-light';
  }
  
  // Theme Toggle listener
  DOM.themeToggle.addEventListener('click', () => {
    const isDark = document.body.classList.contains('theme-dark');
    if (isDark) {
      document.body.className = 'theme-light';
      localStorage.setItem('lfjp_theme', 'light');
    } else {
      document.body.className = 'theme-dark';
      localStorage.setItem('lfjp_theme', 'dark');
    }
  });
}

function bindNavigation() {
  DOM.tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      switchTab(tab.dataset.tab);
    });
  });
  
  // Connect setup instructions Code copy button
  DOM.appsScriptCodeContainer.textContent = APPS_SCRIPT_SOURCE_CODE;
  DOM.copyScriptCodeBtn.addEventListener('click', () => {
    navigator.clipboard.writeText(APPS_SCRIPT_SOURCE_CODE)
      .then(() => {
        DOM.copyScriptCodeBtn.textContent = "Copié !";
        showToast("Code Apps Script copié dans le presse-papiers !", "success");
        setTimeout(() => DOM.copyScriptCodeBtn.textContent = "Copier le code", 2000);
      })
      .catch(err => {
        showToast("Erreur lors de la copie.", "error");
      });
  });
}

function bindStudentsEvents() {
  // Search Events
  DOM.studentSearch.addEventListener('input', renderStudentsGrid);
  DOM.clearSearchBtn.addEventListener('click', () => {
    DOM.studentSearch.value = '';
    renderStudentsGrid();
  });
  
  // Selection Actions
  DOM.floatingBtnClear.addEventListener('click', clearSelection);
  
  // Floating Actions Buttons -> Open Native Modal
  DOM.floatingBtnTogether.addEventListener('click', () => {
    openRequestModal('Regroupement');
  });
  
  DOM.floatingBtnSeparate.addEventListener('click', () => {
    openRequestModal('Éloignement');
  });
  
  // Sync button triggers
  DOM.emptyStateSyncBtn.addEventListener('click', fetchStudents);
  DOM.syncStudentsBtn.addEventListener('click', fetchStudents);
  
  // Close Modals buttons (Native Dialog Close)
  DOM.closeRequestModalBtn.addEventListener('click', () => DOM.requestModal.close());
  DOM.cancelRequestBtn.addEventListener('click', () => DOM.requestModal.close());
  
  if (DOM.explainerVideoBtn) {
    DOM.explainerVideoBtn.addEventListener('click', () => DOM.videoModal.showModal());
  }
  if (DOM.closeVideoModalBtn) {
    DOM.closeVideoModalBtn.addEventListener('click', () => DOM.videoModal.close());
  }
  if (DOM.videoModal) {
    DOM.videoModal.addEventListener('close', () => {
      const iframe = DOM.videoModal.querySelector('iframe');
      if (iframe) {
        const currentSrc = iframe.src;
        iframe.src = '';
        iframe.src = currentSrc;
      }
    });
  }
  
  // Fallbacks for <dialog> light dismiss in Safari
  [DOM.requestModal, DOM.importModal, DOM.videoModal].filter(Boolean).forEach(dialog => {
    if (!('closedBy' in HTMLDialogElement.prototype)) {
      dialog.addEventListener('click', (event) => {
        if (event.target !== dialog) return;
        const rect = dialog.getBoundingClientRect();
        const isInside = (
          rect.top <= event.clientY && event.clientY <= rect.top + rect.height &&
          rect.left <= event.clientX && event.clientX <= rect.left + rect.width
        );
        if (!isInside) dialog.close();
      });
    }
  });
}

/**
 * Open Request Creation Modal
 */
function openRequestModal(type) {
  const selectedCount = state.selectedIds.size;
  if (selectedCount < 2) return;
  
  // Configure Modal headers
  DOM.modalTypeBadge.className = `badge ${type === 'Regroupement' ? 'badge-type-together' : 'badge-type-separate'}`;
  DOM.modalTypeBadge.textContent = type;
  DOM.modalTitle.textContent = type === 'Regroupement' 
    ? 'Nouveau vœu de regroupement' 
    : 'Nouveau vœu d\'éloignement';
  
  DOM.modalLevelText.textContent = state.selectedLevel;
  
  // Populate student list inside modal
  DOM.modalStudentsList.innerHTML = '';
  const selectedStudents = state.students.filter(s => state.selectedIds.has(s.uniqueId));
  
  selectedStudents.forEach(student => {
    const li = document.createElement('li');
    li.innerHTML = `
      <span>👤 &nbsp; ${student.name}</span>
      <span class="student-id-tag">${student.id}</span>
    `;
    DOM.modalStudentsList.appendChild(li);
  });
  
  // Clear textarea
  DOM.inputMotif.value = '';
  
  // Show modal
  DOM.requestModal.showModal();
}

function bindFormSubmissions() {
  // Request Modal Form submission
  DOM.requestForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const type = DOM.modalTypeBadge.textContent;
    const niveau = state.selectedLevel;
    const selectedStudents = state.students.filter(s => state.selectedIds.has(s.uniqueId));
    const ids = selectedStudents.map(s => s.id);
    const names = selectedStudents.map(s => s.name);
    
    const author = DOM.inputAuthor.value.trim();
    const motif = DOM.inputMotif.value.trim();
    
    submitRequest(type, niveau, ids, names, author, motif);
  });
  
  // Settings Form submission
  DOM.settingsForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    state.settings.sheetUrl = DOM.inputSheetUrl.value.trim();
    state.settings.scriptUrl = DOM.inputScriptUrl.value.trim();
    
    Storage.saveSettings();
    showToast("Configuration sauvegardée !", "success");
    
    // Trigger requests sync to verify Sheet URL
    fetchExistingRequests();
  });
  
  // Test connection button
  DOM.testSettingsBtn.addEventListener('click', testBackendConnection);
  
  // Sync requests button
  DOM.syncRequestsBtn.addEventListener('click', () => {
    fetchStudents();
    fetchExistingRequests();
  });
  DOM.requestLevelFilter.addEventListener('change', renderRequestsUI);
  DOM.requestTypeFilter.addEventListener('change', renderRequestsUI);
  DOM.requestSearch.addEventListener('input', renderRequestsUI);
  
  // Clear database button
  DOM.clearLocalDbBtn.addEventListener('click', () => {
    if (confirm("Voulez-vous vraiment effacer tous les élèves importés de la base de données locale ? Cette opération est irréversible.")) {
      state.students = [];
      Storage.saveStudents();
      clearSelection();
      renderStudentsGrid();
      showToast("Base de données locale réinitialisée.", "info");
    }
  });

  // Connection Portal and Logout Events
  DOM.adminLoginBtn.addEventListener('click', () => {
    if (state.role) {
      if (confirm("Voulez-vous vraiment vous déconnecter ?")) {
        logoutUser();
      }
    }
  });

  DOM.portalLoginForm.addEventListener('submit', (e) => {
    e.preventDefault();
    const username = DOM.portalUsername.value.trim();
    const password = DOM.portalPassword.value;
    loginUser(username, password);
  });
}

/**
 * Test connections to the backend Apps Script URL
 */
async function testBackendConnection() {
  const scriptUrl = DOM.inputScriptUrl.value.trim();
  if (!scriptUrl) {
    showToast("Veuillez saisir une URL de script.", "warning");
    return;
  }
  
  DOM.testResultBox.removeAttribute('hidden');
  DOM.testResultBox.className = "connection-status-box";
  DOM.testResultIcon.textContent = "⏳";
  DOM.testResultText.textContent = "Test de connexion en cours...";
  
  try {
    const response = await fetch(scriptUrl, {
      method: 'GET',
      mode: 'cors'
    });
    
    const result = await response.json();
    
    if (result.status === 'ok') {
      DOM.testResultBox.className = "connection-status-box success";
      DOM.testResultIcon.textContent = "✅";
      DOM.testResultText.textContent = "Connexion réussie ! Apps Script configuré avec succès.";
      showToast("Test de connexion réussi !", "success");
    } else {
      throw new Error(result.message || "Erreur de réponse");
    }
  } catch (err) {
    console.error(err);
    DOM.testResultBox.className = "connection-status-box error";
    DOM.testResultIcon.textContent = "❌";
    DOM.testResultText.innerHTML = `
      <strong>Erreur de connexion :</strong> ${err.message}<br>
      <small>Assurez-vous d'avoir déployé en tant qu'<strong>Application Web</strong> avec l'accès configuré sur <strong>"Tout le monde" (Anyone)</strong>.</small>
    `;
    showToast("Échec du test de connexion.", "error");
  }
}



// --- App Initialization ---
document.addEventListener('DOMContentLoaded', () => {
  Storage.load();
  initTheme();
  updateAuthUI();
  bindNavigation();
  renderLevelChips();
  bindStudentsEvents();
  bindFormSubmissions();
  
  // Fill settings inputs in UI
  if (state.settings.sheetUrl) DOM.inputSheetUrl.value = state.settings.sheetUrl;
  if (state.settings.scriptUrl) DOM.inputScriptUrl.value = state.settings.scriptUrl;
  DOM.settingsDbCount.textContent = state.students.length;
  
  // Render students list
  renderStudentsGrid();
  
  // Auto-connect to Sheets if URLs are saved and authorized
  if (state.role && state.settings.sheetUrl) {
    fetchStudents();
    fetchExistingRequests();
  }
  
  // Let the user know the application has loaded
  showToast("Application LFJP Classes chargée avec succès.", "success");
});
