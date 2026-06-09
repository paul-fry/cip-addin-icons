// ============================================================
// cip File Email - Filing dialog logic (dialog.js)
// ============================================================
// Runs inside the dialog window opened by commands.js via
// Office.context.ui.displayDialogAsync. Responsibilities:
//
//   1. Acquire a Graph access token silently from MSAL's localStorage
//      cache (populated by an earlier sign-in via the taskpane). If no
//      session is found, show the "please sign in first" panel - we
//      deliberately do NOT prompt for sign-in inside the dialog,
//      because OAuth popup-in-dialog flows are fragile across Outlook
//      clients.
//   2. Load SharePoint sites + document libraries via Graph for the
//      user to pick from.
//   3. On a button click, post the chosen action and (if "file") the
//      destination + tags back to commands.js via messageParent.
// ============================================================

let msalInstance = null;
let currentAccount = null;
let accessToken = null;
let allSites = [];
let allLibraries = [];
let selectedSite = null;
let selectedLib = null;

// Context handed to us by commands.js (the outgoing message's recipients +
// subject), needed to write a matchable pending-filing record.
let dialogContext = { recipients: [], subject: '' };

// Path of the per-user pending-filing queue in OneDrive. Both this dialog and
// the taskpane poller read/write this single JSON file via Graph - it's the
// bridge that survives the send and crosses the isolated dialog context.
const PENDING_PATH = '/me/drive/root:/cip-file-email-pending.json';

Office.onReady(() => {
  // Receive the context (recipients/subject) that commands.js sends in response
  // to our "ready" handshake below.
  try {
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      try {
        const m = JSON.parse(arg.message || '{}');
        if (m.type === 'context') {
          dialogContext.recipients = Array.isArray(m.recipients) ? m.recipients : [];
          dialogContext.subject = m.subject || '';
        }
      } catch (e) { /* ignore malformed parent messages */ }
    });
  } catch (e) { /* older hosts may not support parent messaging - matching falls back to time only */ }

  init().catch(e => {
    console.error('Dialog init failed:', e);
    showBanner('Could not initialise: ' + e.message, 'error');
  });

  // Announce readiness so commands.js sends us the context.
  try { Office.context.ui.messageParent(JSON.stringify({ type: 'ready' })); }
  catch (e) { /* parent will still work; we just may lack recipients */ }
});

async function init() {
  // Wait briefly for the MSAL library to finish loading. Both the script tag
  // and Office.onReady can race; this is the cheapest way to serialize them.
  let tries = 0;
  while (typeof msal === 'undefined' && tries < 50) {
    await new Promise(r => setTimeout(r, 100));
    tries++;
  }
  if (typeof msal === 'undefined') {
    showBanner('Authentication library failed to load.', 'error');
    return;
  }

  const cfg = window.CIP_CONFIG;
  if (!cfg || !cfg.clientId || !cfg.tenantId) {
    showBanner('Add-in configuration missing.', 'error');
    return;
  }

  // The dialog runs in its OWN storage partition - it CANNOT see the taskpane's
  // MSAL session (Chromium 115+ storage partitioning). So the dialog maintains
  // its own session: the user signs in here once, MSAL caches it in the dialog's
  // partition, and every subsequent send authenticates silently.
  //   - storeAuthStateInCookie:true improves redirect reliability in WebView2
  //     (classic Outlook desktop).
  //   - redirectUri is this dialog page itself; it MUST be registered as a SPA
  //     redirect URI in the Entra app registration.
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: 'https://login.microsoftonline.com/' + cfg.tenantId,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
  });
  await msalInstance.initialize();

  // STEP 1: handle a redirect response - we may be returning from loginRedirect.
  let redirectResponse = null;
  try {
    redirectResponse = await msalInstance.handleRedirectPromise();
  } catch (e) {
    console.warn('handleRedirectPromise error:', e && e.message);
  }
  if (redirectResponse && redirectResponse.account) {
    currentAccount = redirectResponse.account;
    msalInstance.setActiveAccount(currentAccount);
    // loginRedirect with scopes returns a usable access token directly.
    if (redirectResponse.accessToken) accessToken = redirectResponse.accessToken;
  } else {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      currentAccount = accounts[0];
      msalInstance.setActiveAccount(currentAccount);
    }
  }

  // STEP 2: if we have an account, get a token silently (unless we already got
  // one from the redirect response above).
  if (currentAccount) {
    try {
      if (!accessToken) accessToken = await getToken();
      document.getElementById('form-area').style.display = 'block';
      await loadSites();
      return;
    } catch (e) {
      console.warn('Silent token failed, will sign in interactively:', e && e.message);
    }
  }

  // STEP 3: no usable session - sign in interactively INSIDE the dialog.
  // This navigates the dialog window to the Microsoft sign-in page and back.
  // After the round trip, the page reloads and STEP 1 picks up the response.
  try {
    showSigningIn();
    await msalInstance.loginRedirect({ scopes: cfg.scopes });
    // loginRedirect navigates away; nothing after this runs until we return.
  } catch (e) {
    console.error('loginRedirect failed:', e);
    showSignInError('Could not start sign-in: ' + (e && e.message ? e.message : e));
  }
}

async function getToken() {
  const request = { scopes: window.CIP_CONFIG.scopes, account: currentAccount };
  const r = await msalInstance.acquireTokenSilent(request);
  return r.accessToken;
}

async function graph(path, opts) {
  opts = opts || {};
  const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
  const headers = { 'Authorization': 'Bearer ' + accessToken, 'Accept': 'application/json' };
  if (opts.body) headers['Content-Type'] = opts.contentType || 'application/json';
  const res = await fetch(url, {
    method: opts.method || 'GET',
    headers: headers,
    body: opts.body
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    const err = new Error('Graph ' + res.status + (text ? ': ' + text.slice(0, 160) : ''));
    err.status = res.status;
    throw err;
  }
  if (opts.raw) return res;                 // caller wants the raw Response (e.g. file content)
  if (res.status === 204) return null;      // No Content (e.g. DELETE)
  const ct = res.headers.get('content-type') || '';
  return ct.indexOf('application/json') !== -1 ? res.json() : res.text();
}

// Read the OneDrive pending-filing queue (array). Returns [] if the file
// doesn't exist yet.
async function readPendingQueue() {
  try {
    const res = await graph(PENDING_PATH + ':/content', { raw: true });
    const txt = await res.text();
    const arr = JSON.parse(txt);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    if (e && e.status === 404) return [];   // no queue file yet - that's fine
    throw e;
  }
}

// Append one record to the OneDrive pending-filing queue and write it back.
async function appendPendingRecord(record) {
  let queue = [];
  try { queue = await readPendingQueue(); }
  catch (e) { console.warn('Could not read pending queue (will start fresh):', e && e.message); queue = []; }
  queue.push(record);
  await graph(PENDING_PATH + ':/content', {
    method: 'PUT',
    body: JSON.stringify(queue),
    contentType: 'application/json'
  });
}

// ============================================================
// Sites & libraries
// ============================================================
async function loadSites() {
  try {
    const result = await graph('/sites?search=*&$top=100');
    allSites = result.value || [];
    const filter = (window.CIP_CONFIG.siteFilter || []);
    if (filter.length > 0) {
      allSites = allSites.filter(s =>
        filter.some(f => (s.webUrl || '').toLowerCase().includes(f.toLowerCase()))
      );
    }
    allSites.sort((a, b) =>
      (a.displayName || a.name || '').localeCompare(b.displayName || b.name || '')
    );
    renderSites(allSites);
    document.getElementById('site-label').textContent = 'Choose a site…';
  } catch (e) {
    showBanner('Could not load SharePoint sites: ' + e.message, 'error');
  }
}

function renderSites(sites) {
  const c = document.getElementById('site-items');
  if (!sites.length) { c.innerHTML = '<div class="empty">No matching sites</div>'; return; }
  c.innerHTML = sites.map(s => `
    <div class="item" onclick="selectSite('${s.id}')">
      <div class="label">${escapeHtml(s.displayName || s.name || '')}</div>
      <div class="url">${escapeHtml(s.webUrl || '')}</div>
    </div>`).join('');
}

function filterSites() {
  const q = document.getElementById('site-search').value.toLowerCase();
  if (!q) { renderSites(allSites); return; }
  renderSites(allSites.filter(s =>
    (s.displayName || '').toLowerCase().includes(q) ||
    (s.name || '').toLowerCase().includes(q) ||
    (s.webUrl || '').toLowerCase().includes(q)
  ));
}

async function selectSite(siteId) {
  selectedSite = allSites.find(s => s.id === siteId);
  selectedLib = null;
  closePickers();

  document.getElementById('site-picker').classList.remove('empty');
  document.getElementById('site-label').textContent = selectedSite.displayName || selectedSite.name;
  document.getElementById('lib-label').textContent = 'Loading libraries…';
  document.getElementById('lib-picker').classList.add('empty');
  updateFileButton();

  try {
    const r = await graph('/sites/' + siteId + '/drives');
    allLibraries = (r.value || []).filter(d => d.driveType === 'documentLibrary');
    allLibraries.sort((a, b) => a.name.localeCompare(b.name));
    renderLibs(allLibraries);
    // Auto-select the preferred library if one is configured and matches
    const preferred = window.CIP_CONFIG.preferredLibrary;
    if (preferred) {
      const match = allLibraries.find(l => l.name === preferred);
      if (match) selectLib(match.id);
    }
    if (!selectedLib) document.getElementById('lib-label').textContent = 'Pick a library…';
  } catch (e) {
    showBanner('Could not load libraries: ' + e.message, 'error');
    document.getElementById('lib-label').textContent = 'Error loading libraries';
  }
}

function renderLibs(libs) {
  const c = document.getElementById('lib-items');
  if (!libs.length) { c.innerHTML = '<div class="empty">No document libraries</div>'; return; }
  c.innerHTML = libs.map(l => `
    <div class="item" onclick="selectLib('${l.id}')">
      <div class="label">${escapeHtml(l.name)}</div>
      ${l.description ? `<div class="url">${escapeHtml(l.description)}</div>` : ''}
    </div>`).join('');
}

function selectLib(libId) {
  selectedLib = allLibraries.find(l => l.id === libId);
  closePickers();
  document.getElementById('lib-picker').classList.remove('empty');
  document.getElementById('lib-label').textContent = selectedLib.name;
  updateFileButton();
}

// ============================================================
// Picker open/close
// ============================================================
function togglePicker(which) {
  const me = document.getElementById((which === 'site' ? 'site' : 'lib') + '-list');
  const other = document.getElementById((which === 'site' ? 'lib' : 'site') + '-list');
  other.classList.remove('show');
  // The library picker only opens once a site is selected
  if (which === 'lib' && !selectedSite) return;
  me.classList.toggle('show');
}
function closePickers() {
  document.querySelectorAll('.picker-list').forEach(p => p.classList.remove('show'));
}
document.addEventListener('click', (e) => {
  if (!e.target.closest('.picker') && !e.target.closest('.picker-list')) closePickers();
});

// ============================================================
// UI helpers
// ============================================================
function updateFileButton() {
  document.getElementById('btn-file').disabled = !(selectedSite && selectedLib);
}

function showBanner(msg, type) {
  const b = document.getElementById('banner');
  b.className = 'banner ' + (type || 'info');
  b.textContent = msg;
}

// Shown briefly while the dialog redirects to the Microsoft sign-in page.
function showSigningIn() {
  document.getElementById('form-area').style.display = 'none';
  const area = document.getElementById('signin-area');
  area.style.display = 'block';
  area.innerHTML = '<strong>Signing in…</strong>Redirecting you to sign in. This happens once - after that, filing on send is instant.';
  document.getElementById('btn-file').disabled = true;
}

// Shown only if interactive sign-in itself fails to start.
function showSignInError(message) {
  document.getElementById('form-area').style.display = 'none';
  const area = document.getElementById('signin-area');
  area.style.display = 'block';
  area.innerHTML = '<strong>Sign-in problem</strong>' + escapeHtml(message || 'Could not sign in.') +
    ' You can still choose "Send without filing" and file this email manually later.';
  document.getElementById('btn-file').disabled = true;
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g,
    c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ============================================================
// Send the user's choice back to commands.js via messageParent.
// For "file", first write a pending-filing record to OneDrive (the bridge the
// taskpane poller reads), then report success/failure back to commands.js.
// ============================================================
async function onAction(action) {
  if (action !== 'file') {
    // skip / cancel - just tell commands.js; nothing to write.
    sendToParent({ type: 'action', action: action });
    return;
  }

  if (!selectedSite || !selectedLib) {
    showBanner('Pick a site and library first.', 'warn');
    return;
  }

  const btn = document.getElementById('btn-file');
  btn.disabled = true;
  btn.textContent = 'Filing…';

  const record = {
    id: 'p_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
    armedAt: new Date().toISOString(),
    recipients: dialogContext.recipients || [],
    subject: dialogContext.subject || '',
    dest: {
      siteId:   selectedSite.id,
      siteName: selectedSite.displayName || selectedSite.name,
      libId:    selectedLib.id,
      libName:  selectedLib.name,
      folder:   document.getElementById('f-folder').value.trim()
    },
    tags: {
      client:   document.getElementById('f-client').value.trim(),
      project:  document.getElementById('f-project').value.trim(),
      category: document.getElementById('f-category').value,
      notes:    document.getElementById('f-notes').value.trim()
    },
    source: 'dialog'
  };

  let written = false;
  try {
    await appendPendingRecord(record);
    written = true;
  } catch (e) {
    console.error('Could not write pending-filing record to OneDrive:', e && e.message);
    written = false;
  }

  // Report back. commands.js allows the send either way - if the write failed,
  // it logs a warning and the user can file manually from the sent copy.
  sendToParent({ type: 'action', action: 'file', written: written });
}

function sendToParent(obj) {
  try { Office.context.ui.messageParent(JSON.stringify(obj)); }
  catch (e) { console.error('messageParent failed:', e); }
}
