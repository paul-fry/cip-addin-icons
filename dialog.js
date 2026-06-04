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

Office.onReady(() => {
  init().catch(e => {
    console.error('Dialog init failed:', e);
    showBanner('Could not initialise: ' + e.message, 'error');
  });
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

  // cacheLocation:'localStorage' so we read the same session the taskpane
  // wrote on sign-in. Origin is the same for both pages so this just works.
  msalInstance = new msal.PublicClientApplication({
    auth: {
      clientId: cfg.clientId,
      authority: 'https://login.microsoftonline.com/' + cfg.tenantId,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: 'localStorage' }
  });
  await msalInstance.initialize();

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length === 0) {
    showSignInPrompt();
    return;
  }
  currentAccount = accounts[0];

  try {
    accessToken = await getToken();
  } catch (e) {
    // Silent token failed (account expired, password changed, etc.). Don't
    // try to popup-prompt from inside the dialog - show the inline prompt.
    console.warn('Silent token failed:', e);
    showSignInPrompt();
    return;
  }

  document.getElementById('form-area').style.display = 'block';
  await loadSites();
}

async function getToken() {
  const request = { scopes: window.CIP_CONFIG.scopes, account: currentAccount };
  const r = await msalInstance.acquireTokenSilent(request);
  return r.accessToken;
}

async function graph(path) {
  const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
  const res = await fetch(url, {
    headers: { 'Authorization': 'Bearer ' + accessToken, 'Accept': 'application/json' }
  });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error('Graph ' + res.status + (text ? ': ' + text.slice(0, 160) : ''));
  }
  return res.json();
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

function showSignInPrompt() {
  document.getElementById('form-area').style.display = 'none';
  document.getElementById('signin-area').style.display = 'block';
  // Disable File - signing in is required first
  document.getElementById('btn-file').disabled = true;
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g,
    c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ============================================================
// Send the user's choice back to commands.js via messageParent
// ============================================================
function onAction(action) {
  let payload;
  if (action === 'file') {
    if (!selectedSite || !selectedLib) {
      showBanner('Pick a site and library first.', 'warn');
      return;
    }
    payload = {
      action: 'file',
      siteId:   selectedSite.id,
      siteName: selectedSite.displayName || selectedSite.name,
      libId:    selectedLib.id,
      libName:  selectedLib.name,
      folder:   document.getElementById('f-folder').value.trim(),
      client:   document.getElementById('f-client').value.trim(),
      project:  document.getElementById('f-project').value.trim(),
      category: document.getElementById('f-category').value,
      notes:    document.getElementById('f-notes').value.trim()
    };
  } else {
    payload = { action };
  }
  try {
    Office.context.ui.messageParent(JSON.stringify(payload));
  } catch (e) {
    console.error('messageParent failed:', e);
  }
}
