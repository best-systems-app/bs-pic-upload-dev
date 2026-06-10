/**
 * BS pic upload - App.js
 * Version 18 - SharePoint Upload
 * Sicherheits- und Barrierefreiheits-Verbesserungen
 */

// ═════════════════════════════════════════════════════════════════════
// KONSTANTEN & KONFIGURATION
// ═════════════════════════════════════════════════════════════════════

const CONFIG = {
  // Microsoft OAuth
  CLIENT_ID: '180e1fc6-2c2b-4a77-82e8-e3dbcd219491',
  TENANT_ID: '175d1508-8739-4676-bb66-98aa62745feb',
  REDIRECT: 'https://best-systems-app.github.io/bs-pic-upload-dev/',
  SCOPES: 'Files.ReadWrite.All User.Read offline_access',

  // SharePoint
  SP_HOST: 'bestsystemsvienna.sharepoint.com',
  SP_SITE: 'BestSystems-Europa',
  SP_FOLDER: 'BS-PIC-UPLOADER',

  // Validierung
  AUFTRAG_MIN: 1,
  AUFTRAG_MAX: 20,
  MAX_PHOTOS: 20,
  MAX_FILE_MB: 15,

  // UI
  TOAST_DURATION: 3000
};

// ═════════════════════════════════════════════════════════════════════
// STATE
// ═════════════════════════════════════════════════════════════════════

const STATE = {
  navHistory: ['login'],
  auftragNummer: null,
  photos: [],
  numpadVal: '',
  activeTab: 'scanner',
  numpadExpanded: true,
  prefix: 'AB',
  uploadAbortController: null
};

// ═════════════════════════════════════════════════════════════════════
// LOGGING & ERROR HANDLING
// ═════════════════════════════════════════════════════════════════════

function logError(context, error) {
  const message = `[${context}] ${error.message || error}`;
  console.error(message, error);
  // Optional: An Analytics-Service senden
}

function showToast(msg, type = 'error') {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = 'toast show ' + type;
  setTimeout(() => toast.classList.remove('show'), CONFIG.TOAST_DURATION);
}

// ═════════════════════════════════════════════════════════════════════
// AUTH - OAUTH mit PKCE
// ═════════════════════════════════════════════════════════════════════

function b64url(buf) {
  return btoa(String.fromCharCode(...new Uint8Array(buf)))
    .replace(/\+/g, '-').replace(/\//g, '_').replace(/=/g, '');
}

async function genPKCE() {
  const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
  const hashBuffer = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier));
  return {
    verifier,
    challenge: b64url(hashBuffer)
  };
}

async function doLogin() {
  const btn = document.getElementById('btn-login');
  btn.disabled = true;
  btn.classList.add('loading');

  try {
    const { verifier, challenge } = await genPKCE();
    const state = b64url(crypto.getRandomValues(new Uint8Array(8)));

    sessionStorage.setItem('pkce_v', verifier);
    sessionStorage.setItem('oauth_s', state);

    const params = new URLSearchParams({
      client_id: CONFIG.CLIENT_ID,
      response_type: 'code',
      redirect_uri: CONFIG.REDIRECT,
      scope: CONFIG.SCOPES,
      state,
      code_challenge: challenge,
      code_challenge_method: 'S256',
      prompt: 'select_account'
    });

    window.location.href = `https://login.microsoftonline.com/${CONFIG.TENANT_ID}/oauth2/v2.0/authorize?${params}`;
  } catch (e) {
    logError('doLogin', e);
    showToast('Fehler bei Anmeldung', 'error');
    btn.disabled = false;
    btn.classList.remove('loading');
  }
}

async function handleCallback() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get('code');
  const state = params.get('state');

  if (!code || state !== sessionStorage.getItem('oauth_s')) {
    return false;
  }

  const body = new URLSearchParams({
    client_id: CONFIG.CLIENT_ID,
    grant_type: 'authorization_code',
    code,
    redirect_uri: CONFIG.REDIRECT,
    code_verifier: sessionStorage.getItem('pkce_v')
  });

  try {
    const resp = await fetch(
      `https://login.microsoftonline.com/${CONFIG.TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body
      }
    );

    const data = await resp.json();

    if (data.access_token) {
      // ⚠️ Sicherheitshinweis: Tokens in localStorage sind anfällig für XSS
      // Für Production sollte ein Backend-Proxy oder HttpOnly Cookies verwendet werden
      localStorage.removeItem('bs_sp_drive');
      localStorage.setItem('bs_at', data.access_token);
      localStorage.setItem('bs_rt', data.refresh_token || '');
      localStorage.setItem('bs_exp', Date.now() + (data.expires_in - 60) * 1000);

      window.history.replaceState({}, '', CONFIG.REDIRECT);
      return true;
    }

    showError(data.error_description || 'Unbekannter Fehler');
    return false;
  } catch (e) {
    logError('handleCallback', e);
    showError(e.message);
    return false;
  }
}

function showError(message) {
  const errorEl = document.getElementById('login-error');
  errorEl.textContent = 'Fehler: ' + message;
  errorEl.classList.add('visible');
}

async function getToken() {
  const exp = parseInt(localStorage.getItem('bs_exp') || '0');

  if (Date.now() < exp) {
    return localStorage.getItem('bs_at');
  }

  const rt = localStorage.getItem('bs_rt');
  if (!rt) return null;

  try {
    const body = new URLSearchParams({
      client_id: CONFIG.CLIENT_ID,
      grant_type: 'refresh_token',
      refresh_token: rt,
      scope: CONFIG.SCOPES
    });

    const resp = await fetch(
      `https://login.microsoftonline.com/${CONFIG.TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body
      }
    );

    const data = await resp.json();

    if (data.access_token) {
      localStorage.setItem('bs_at', data.access_token);
      localStorage.setItem('bs_exp', Date.now() + (data.expires_in - 60) * 1000);
      if (data.refresh_token) localStorage.setItem('bs_rt', data.refresh_token);
      return data.access_token;
    }
  } catch (e) {
    logError('getToken', e);
  }

  return null;
}

// ═════════════════════════════════════════════════════════════════════
// NAVIGATION
// ═════════════════════════════════════════════════════════════════════

function showScreen(id) {
  // Scan-Screen verlassen: visibilitychange-Listener aufraemen
  if (id !== 'scan') teardownScanScreen();

  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + id).classList.add('active');

  const back = document.getElementById('btn-back');
  const title = document.getElementById('topbar-title');
  const map = {
    login: ['', false],
    scan: ['', false],
    photos: ['', true],
    upload: ['Upload…', false],
    success: ['', false]
  };

  const [titleText, showBack] = map[id] || ['', false];
  title.textContent = titleText;

  if (showBack) {
    back.classList.add('visible');
  } else {
    back.classList.remove('visible');
  }
}

function navigate(id) {
  STATE.navHistory.push(id);
  showScreen(id);
}

function goBack() {
  STATE.navHistory.pop();
  showScreen(STATE.navHistory[STATE.navHistory.length - 1] || 'login');
}

// ═════════════════════════════════════════════════════════════════════
// TABS
// ═════════════════════════════════════════════════════════════════════

function switchTab(tab) {
  if (STATE.activeTab === tab) return;

  STATE.activeTab = tab;

  document.querySelectorAll('.tab').forEach(t => {
    t.classList.remove('active');
    t.setAttribute('aria-selected', 'false');
  });

  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));

  const activeEl = document.getElementById('tab-' + tab);
  activeEl.classList.add('active');
  activeEl.setAttribute('aria-selected', 'true');

  document.getElementById('tab-content-' + tab).classList.add('active');

  if (tab === 'scanner') {
    focusScannerInput();
  } else {
    const si = document.getElementById('scanner-input');
    if (si) si.blur();
  }
}

// ═════════════════════════════════════════════════════════════════════
// NUMPAD
// ═════════════════════════════════════════════════════════════════════

function numpadKey(k) {
  if (STATE.numpadVal.length >= CONFIG.AUFTRAG_MAX) return;

  STATE.numpadVal += k;
  updateNumpadDisplay();

  document.getElementById('btn-numpad-confirm').disabled = STATE.numpadVal.length < CONFIG.AUFTRAG_MIN;
}

function numpadBack() {
  STATE.numpadVal = STATE.numpadVal.slice(0, -1);
  updateNumpadDisplay();

  document.getElementById('btn-numpad-confirm').disabled = STATE.numpadVal.length < CONFIG.AUFTRAG_MIN;

  if (STATE.numpadVal.length === 0) expandNumpad();
}

function numpadClear() {
  STATE.numpadVal = '';
  updateNumpadDisplay();
  document.getElementById('btn-numpad-confirm').disabled = true;
  expandNumpad();
}

function updateNumpadDisplay() {
  const el = document.getElementById('numpad-display');
  const counter = document.getElementById('numpad-counter');

  if (STATE.numpadVal) {
    el.textContent = STATE.numpadVal;
    el.classList.remove('placeholder');
    if (counter) {
      counter.textContent = STATE.numpadVal.length + ' Stellen';
      counter.style.color = 'var(--muted)';
    }
  } else {
    el.textContent = 'Nummer eingeben…';
    el.classList.add('placeholder');
    if (counter) {
      counter.textContent = '0 Stellen';
      counter.style.color = 'var(--muted)';
    }
  }
}

function collapseNumpad() {
  STATE.numpadExpanded = false;
  document.getElementById('numpad-collapsible').classList.add('collapsed');
  document.getElementById('numpad-edit-hint').classList.add('visible');
  document.getElementById('numpad-display-wrap').style.borderColor = 'var(--accent)';
}

function expandNumpad() {
  STATE.numpadExpanded = true;
  document.getElementById('numpad-collapsible').classList.remove('collapsed');
  document.getElementById('numpad-edit-hint').classList.remove('visible');
  document.getElementById('numpad-display-wrap').style.borderColor = '';
}

function confirmNumpad() {
  if (STATE.numpadVal.length >= CONFIG.AUFTRAG_MIN) {
    setAuftrag(STATE.prefix + STATE.numpadVal);
  }
}

function selectPrefix(p) {
  STATE.prefix = p;
  document.querySelectorAll('.prefix-btn').forEach(function(btn) {
    btn.classList.toggle('active', btn.dataset.prefix === p);
  });
}

// ═════════════════════════════════════════════════════════════════════
// SCANNER INPUT
// ═════════════════════════════════════════════════════════════════════

// Steuerzeichen + Whitespace entfernen: \x00-\x1F = alle ASCII-Steuerzeichen (STX, ETX, CR, LF, GS etc.)
function sanitizeScanValue(raw) {
  return raw.replace(/[\x00-\x1F\x7F\s]/g, '');
}

function focusScannerInput() {
  const input = document.getElementById('scanner-input');
  if (!input) return;
  // requestAnimationFrame stellt sicher dass Android Chrome den Layout-Pass
  // abgeschlossen hat bevor focus() aufgerufen wird
  requestAnimationFrame(() => {
    input.focus({ preventScroll: true });
  });
}

function setupScannerInput() {
  const input = document.getElementById('scanner-input');
  if (!input) return;

  input.addEventListener('focus', () => {
    input.removeAttribute('readonly');
    const status = document.getElementById('scanner-status');
    if (status) status.textContent = 'Bereit – jetzt scannen';
  });

  input.addEventListener('blur', () => {
    input.setAttribute('readonly', '');
    const status = document.getElementById('scanner-status');
    if (status) status.textContent = 'Scanner bereit';
  });

  input.addEventListener('input', () => {
    const val = sanitizeScanValue(input.value);
    input.value = val;
    const chars = document.getElementById('scanner-chars');
    if (chars) chars.textContent = val ? val.length + ' Zeichen' : '';
  });

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      const val = sanitizeScanValue(input.value);
      input.value = '';
      const chars = document.getElementById('scanner-chars');
      if (chars) chars.textContent = '';
      if (val) setAuftrag(val);
    }
  });

  const scanCard = document.getElementById('tab-content-scanner');
  if (scanCard) {
    scanCard.addEventListener('touchend', (e) => {
      const tag = e.target.tagName;
      if (tag === 'BUTTON' || tag === 'A' || tag === 'INPUT') return;
      focusScannerInput();
    });
  }
}

// ═════════════════════════════════════════════════════════════════════
// SET AUFTRAG
// ═════════════════════════════════════════════════════════════════════

function setAuftrag(value) {
  value = value.trim();

  if (!value) return;

  if (value.toUpperCase().startsWith('LI')) {
    showToast('LI-Nummern: kein Foto-Upload möglich', 'error');
    _flashScannerCard('var(--danger)');
    return;
  }

  STATE.auftragNummer = value;

  document.getElementById('result-value').textContent = value;
  document.getElementById('result-box').classList.add('visible');
  document.getElementById('btn-to-photos').disabled = false;
  document.getElementById('btn-to-photos-wrap').style.display = '';

  showToast('Auftrag: ' + value, 'success');

  // Haptisches Feedback (Newland unterstuetzt vibrate via Chrome)
  if ('vibrate' in navigator) navigator.vibrate([80, 40, 80]);

  // Visuelles Scan-Erfolg-Feedback: kurz gruen aufleuchten
  _flashScannerCard('#7FBA00');

  if (STATE.activeTab === 'numpad') collapseNumpad();

  // Tastatur wegblenden damit btn-to-photos sichtbar wird
  const si = document.getElementById('scanner-input');
  if (si) si.blur();
}

function _flashScannerCard(color) {
  const wrap = document.querySelector('.scanner-icon-wrap');
  if (!wrap) return;
  const orig = wrap.style.borderColor;
  wrap.style.borderColor = color;
  wrap.style.boxShadow = `0 0 0 4px ${color}44`;
  setTimeout(() => {
    wrap.style.borderColor = orig;
    wrap.style.boxShadow = '';
  }, 800);
}

function resetScan() {
  STATE.auftragNummer = null;
  document.getElementById('result-box').classList.remove('visible');
  document.getElementById('btn-to-photos').disabled = true;
  document.getElementById('btn-to-photos-wrap').style.display = 'none';
  const si = document.getElementById('scanner-input'); if (si) si.value = '';
  STATE.numpadVal = '';
  updateNumpadDisplay();
  document.getElementById('btn-numpad-confirm').disabled = true;
  expandNumpad();

  if (STATE.activeTab === 'scanner') focusScannerInput();
  selectPrefix('AB');
}

function goToPhotos() {
  if (!STATE.auftragNummer) return;

  document.getElementById('photo-auftrag-label').textContent = STATE.auftragNummer;
  navigate('photos');
}

// ═════════════════════════════════════════════════════════════════════
// PHOTOS
// ═════════════════════════════════════════════════════════════════════

function addPhotos(input) {
  const newFiles = Array.from(input.files);
  let warnedLimit = false;

  newFiles.forEach(file => {
    if (STATE.photos.length >= CONFIG.MAX_PHOTOS) {
      if (!warnedLimit) {
        showToast(`Mehr als ${CONFIG.MAX_PHOTOS} Fotos — Upload kann sehr lange dauern`, 'error');
        warnedLimit = true;
      }
      return;
    }

    if (file.size > CONFIG.MAX_FILE_MB * 1024 * 1024) {
      showToast(`"${file.name}" ist größer als ${CONFIG.MAX_FILE_MB} MB`, 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = e => {
      STATE.photos.push({
        file,
        dataUrl: e.target.result
      });
      renderPhotoGrid();
    };
    reader.onerror = () => {
      logError('FileReader', new Error(`Fehler beim Lesen von ${file.name}`));
      showToast(`Fehler beim Lesen von "${file.name}"`, 'error');
    };
    reader.readAsDataURL(file);
  });

  input.value = '';
}

function removePhoto(idx) {
  STATE.photos.splice(idx, 1);
  renderPhotoGrid();
}

function renderPhotoGrid() {
  const grid = document.getElementById('photo-grid');
  grid.innerHTML = '';

  STATE.photos.forEach((p, i) => {
    const div = document.createElement('div');
    div.className = 'photo-thumb';

    const img = document.createElement('img');
    img.src = p.dataUrl;
    img.alt = `Foto ${i + 1}`;

    const btn = document.createElement('button');
    btn.className = 'photo-remove';
    btn.setAttribute('aria-label', 'Foto entfernen');
    btn.textContent = '×';
    btn.addEventListener('click', () => removePhoto(i));

    div.appendChild(img);
    div.appendChild(btn);
    grid.appendChild(div);
  });

  const n = STATE.photos.length;
  const word = n === 1 ? 'Foto' : 'Fotos';

  document.getElementById('photo-count-text').innerHTML = `<span id="photo-count-val">${n}</span> ${word} ausgewählt`;
  document.getElementById('btn-upload').disabled = n === 0;
}

// ═════════════════════════════════════════════════════════════════════
// UPLOAD
// ═════════════════════════════════════════════════════════════════════

async function startUpload() {
  if (!STATE.photos.length) return;

  navigate('upload');
  STATE.uploadAbortController = new AbortController();

  const token = await getToken();
  if (!token) {
    showToast('Nicht angemeldet', 'error');
    showScreen('login');
    return;
  }

  const folderPath = STATE.auftragNummer;
  const total = STATE.photos.length;
  let done = 0;

  const bar = document.getElementById('upload-progress-bar');
  const label = document.getElementById('upload-progress-label');
  const progressBar = document.querySelector('[role="progressbar"]');

  document.getElementById('upload-desc').textContent = 'SharePoint: ' + CONFIG.SP_FOLDER + '/' + folderPath;
  label.textContent = `0 von ${total}`;

  try {
    await ensureFolder(token, folderPath);

    for (let i = 0; i < STATE.photos.length; i++) {
      if (STATE.uploadAbortController.signal.aborted) {
        throw new DOMException('Abgebrochen', 'AbortError');
      }

      const p = STATE.photos[i];
      const ext = (p.file.name.split('.').pop() || 'jpg').toLowerCase();
      const fn = `foto_${String(i + 1).padStart(3, '0')}_${Date.now()}.${ext}`;

      await uploadFile(token, folderPath, fn, p.file, STATE.uploadAbortController.signal);

      done++;
      const percent = Math.round((done / total) * 100);
      bar.style.width = percent + '%';
      progressBar.setAttribute('aria-valuenow', percent);
      label.textContent = `${done} von ${total}`;
    }

    const photoWord = total === 1 ? 'Foto' : 'Fotos';
    document.getElementById('success-path').textContent = 'SharePoint / ' + CONFIG.SP_FOLDER + '/' + folderPath;
    document.getElementById('success-count').textContent = `${total} ${photoWord} hochgeladen`;

    navigate('success');
  } catch (e) {
    if (e.name === 'AbortError') {
      showToast('Upload abgebrochen', 'error');
      navigate('photos');
    } else {
      showToast(`Upload fehlgeschlagen: ${e.message || 'Fehler'}`, 'error');
      navigate('photos');
      logError('startUpload', e);
    }
  } finally {
    STATE.uploadAbortController = null;
  }
}

function cancelUpload() {
  if (STATE.uploadAbortController) {
    STATE.uploadAbortController.abort();
  }
}

async function resolveDriveId(token) {
  const cached = localStorage.getItem('bs_sp_drive');
  if (cached) return cached;

  const siteResp = await fetch(
    'https://graph.microsoft.com/v1.0/sites/' + CONFIG.SP_HOST + ':/sites/' + CONFIG.SP_SITE,
    { headers: { Authorization: 'Bearer ' + token } }
  );
  const site = await siteResp.json();
  if (!site.id) throw new Error('SharePoint-Site nicht erreichbar: ' + (site.error && site.error.message || JSON.stringify(site)));

  const driveResp = await fetch(
    'https://graph.microsoft.com/v1.0/sites/' + site.id + '/drive',
    { headers: { Authorization: 'Bearer ' + token } }
  );
  const drive = await driveResp.json();
  if (!drive.id) throw new Error('SharePoint-Drive nicht gefunden');

  localStorage.setItem('bs_sp_drive', drive.id);
  return drive.id;
}

async function ensureFolder(token, auftragNr) {
  const driveId = await resolveDriveId(token);
  const paths = [CONFIG.SP_FOLDER, CONFIG.SP_FOLDER + '/' + auftragNr];

  for (const p of paths) {
    const check = await fetch(
      'https://graph.microsoft.com/v1.0/drives/' + driveId + '/root:/' + p,
      { headers: { Authorization: 'Bearer ' + token } }
    );
    if (check.status === 404) {
      const segments = p.split('/');
      const name = segments.pop();
      const parentPath = segments.length
        ? 'root:/' + segments.join('/') + ':/children'
        : 'root/children';

      const createResp = await fetch(
        'https://graph.microsoft.com/v1.0/drives/' + driveId + '/' + parentPath,
        {
          method: 'POST',
          headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: name, folder: {}, '@microsoft.graph.conflictBehavior': 'rename' })
        }
      );
      if (!createResp.ok) {
        const err = await createResp.json().catch(() => ({}));
        throw new Error((err.error && err.error.message) || createResp.statusText);
      }
    }
  }
}

async function uploadFile(token, auftragNr, filename, file, signal) {
  const driveId = await resolveDriveId(token);
  const path = CONFIG.SP_FOLDER + '/' + auftragNr + '/' + filename;
  const url = 'https://graph.microsoft.com/v1.0/drives/' + driveId + '/root:/' + path + ':/content';

  const resp = await fetch(url, {
    method: 'PUT',
    headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': file.type || 'image/jpeg' },
    body: file,
    signal
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error((err.error && err.error.message) || resp.statusText);
  }
}

function doLogout() {
  localStorage.removeItem('bs_at');
  localStorage.removeItem('bs_rt');
  localStorage.removeItem('bs_exp');
  localStorage.removeItem('bs_sp_drive');
  sessionStorage.clear();
  document.getElementById('btn-logout').classList.remove('visible');
  showScreen('login');
}

function resetApp() {
  STATE.auftragNummer = null;
  STATE.photos = [];
  STATE.numpadVal = '';

  document.getElementById('result-box').classList.remove('visible');
  document.getElementById('btn-to-photos').disabled = true;
  document.getElementById('btn-to-photos-wrap').style.display = 'none';
  const si = document.getElementById('scanner-input'); if (si) si.value = '';
  updateNumpadDisplay();
  renderPhotoGrid();
  expandNumpad();

  STATE.navHistory = ['login', 'scan'];

  showScreen('scan');
  focusScannerInput();
}

// ═════════════════════════════════════════════════════════════════════
// EVENT LISTENERS - INIT & SETUP
// ═════════════════════════════════════════════════════════════════════

function setupEventListeners() {
  // Back button
  document.getElementById('btn-back').addEventListener('click', goBack);

  // Login
  document.getElementById('btn-login').addEventListener('click', doLogin);

  // Tabs
  document.getElementById('tab-scanner').addEventListener('click', () => switchTab('scanner'));
  document.getElementById('tab-numpad').addEventListener('click', () => switchTab('numpad'));

  // Scanner input
  setupScannerInput();

  // Result clear
  document.getElementById('btn-result-clear').addEventListener('click', resetScan);

  // Numpad
  document.getElementById('numpad-display-wrap').addEventListener('click', expandNumpad);
  document.getElementById('btn-numpad-back').addEventListener('click', () => {
    numpadBack();
  });
  document.getElementById('btn-numpad-clear').addEventListener('click', numpadClear);
  document.getElementById('btn-numpad-confirm').addEventListener('click', confirmNumpad);

  // Numpad keys
  document.querySelectorAll('.numpad-key[data-key]').forEach(btn => {
    btn.addEventListener('click', () => numpadKey(btn.dataset.key));
  });

  // Prefix buttons
  document.querySelectorAll('.prefix-btn').forEach(function(btn) {
    btn.addEventListener('click', function() { selectPrefix(btn.dataset.prefix); });
  });

  // Navigation
  document.getElementById('btn-to-photos').addEventListener('click', goToPhotos);
  document.getElementById('btn-photos-back').addEventListener('click', goBack);

  // Photos
  document.getElementById('photo-input').addEventListener('change', function() {
    addPhotos(this);
  });

  // Label click für Photo input
  document.querySelector('label[for="photo-input"]').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      document.getElementById('photo-input').click();
    }
  });

  // Upload
  document.getElementById('btn-upload').addEventListener('click', startUpload);
  document.getElementById('btn-upload-cancel').addEventListener('click', cancelUpload);

  // Success
  document.getElementById('btn-reset-app').addEventListener('click', resetApp);
  document.getElementById('btn-logout').addEventListener('click', doLogout);
}

// ═════════════════════════════════════════════════════════════════════
// APP INIT
// ═════════════════════════════════════════════════════════════════════

(async function init() {
  // Setup event listeners
  setupEventListeners();

  // Check OAuth callback
  if (window.location.search.indexOf('code=') !== -1) {
    const ok = await handleCallback();
    if (ok) {
      STATE.navHistory = ['login', 'scan'];
      document.getElementById('btn-logout').classList.add('visible');
      showScreen('scan');
      setupScanScreen();
      return;
    }
  }

  // Check existing token
  const token = await getToken();
  if (token) {
    STATE.navHistory = ['login', 'scan'];
    document.getElementById('btn-logout').classList.add('visible');
    showScreen('scan');
    setupScanScreen();
  } else {
    showScreen('login');
  }
})();

function setupScanScreen() {
  // Initialer Focus nach Navigation
  if (STATE.activeTab === 'scanner') focusScannerInput();

  // Wenn App aus Hintergrund zurueckkehrt (z.B. nach Kamera-App)
  document.addEventListener('visibilitychange', _onVisibilityChange);
}

function teardownScanScreen() {
  document.removeEventListener('visibilitychange', _onVisibilityChange);
}

function _onVisibilityChange() {
  if (!document.hidden && STATE.activeTab === 'scanner') {
    const scanScreen = document.getElementById('screen-scan');
    if (scanScreen && scanScreen.classList.contains('active')) {
      focusScannerInput();
    }
  }
}
