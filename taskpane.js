// ============================================================
// cip File Email - Outlook taskpane logic
// ============================================================
// Supports two modes:
//   READ mode    - viewing a received/sent email, click "File email" to upload it to SharePoint now
//   COMPOSE mode - drafting a new email, configure destination, will prompt to file after send
// ============================================================

let msalInstance = null;
let currentAccount = null;
let accessToken = null;
let officeReady = false;
let currentItem = null;
let mode = 'read';  // 'read' | 'compose'
let allSites = [];
let allLibraries = [];
let selectedSite = null;
let selectedLib = null;
let recentlyFiled = [];
let pendingSendInfo = null;  // populated after Send completes, signals "show send banner"

// ----- Multi-select state -----
let multiSelectSupported = false;   // true if Mailbox requirement set 1.13+ is available
let selectedItems = [];             // populated when more than one message is selected in read mode
let batchInProgress = false;        // suppress duplicate batch runs

const RECENT_KEY = 'cip-file-email-recent';

// ============================================================
// Initialise
// ============================================================
Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    officeReady = true;
    currentItem = Office.context.mailbox.item;
    detectMode();
    if (mode === 'compose') {
      configureComposeMode();
    } else {
      populateEmailPreview();
    }
    initAuth();
    // Subscribe to multi-select events on hosts that support requirement set 1.13+
    initMultiSelect();
  }
});

// Detect requirement-set 1.13 multi-select support, subscribe to selection
// changes, and refresh button state on every change. Safe no-op on older
// hosts (Mailbox < 1.13) which simply lack getSelectedItemsAsync.
function initMultiSelect() {
  try {
    if (mode !== 'read') return;
    const mbox = Office.context.mailbox;
    if (!mbox || typeof mbox.getSelectedItemsAsync !== 'function') return;
    multiSelectSupported = true;

    // Initial selection state (covers the case where the add-in opens with
    // multiple messages already selected).
    refreshSelectedItems();

    // Subscribe to changes so the button label and email-preview area
    // stay in sync as the user adds/removes messages from the selection.
    if (Office.EventType && Office.EventType.SelectedItemsChanged) {
      try {
        mbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, () => refreshSelectedItems());
      } catch (e) { /* host doesn't support this event - silently ignore */ }
    }
  } catch (e) {
    // Anything goes wrong, fall back to single-item behaviour
    multiSelectSupported = false;
  }
}

function refreshSelectedItems() {
  if (!multiSelectSupported) return;
  Office.context.mailbox.getSelectedItemsAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      selectedItems = [];
      updatePrimaryButtonLabel();
      return;
    }
    selectedItems = result.value || [];
    // When more than one message is selected, currentItem is null on most hosts.
    // Repaint the email preview area (will show "N messages selected" banner).
    if (mode === 'read') populateEmailPreview();
    updatePrimaryButtonLabel();
  });
}

function updatePrimaryButtonLabel() {
  if (mode !== 'read') return;
  const btn = document.getElementById('file-btn');
  if (!btn) return;
  const count = selectedItems.length;
  if (count > 1) {
    btn.textContent = `File ${count} selected emails`;
  } else {
    btn.textContent = 'File to SharePoint';
  }
}

function detectMode() {
  // Item types: itemType is 'message' for both; we distinguish by whether `to` is an array
  // (read mode) or a Recipients object with getAsync (compose mode).
  if (!currentItem) { mode = 'read'; return; }
  if (currentItem.to && typeof currentItem.to.getAsync === 'function') {
    mode = 'compose';
  } else {
    mode = 'read';
  }
}

function configureComposeMode() {
  // Swap pane copy for compose context
  document.getElementById('pane-doc-type').textContent = 'File on Send';
  document.getElementById('pane-title').textContent = 'File this email after sending';
  document.getElementById('pane-sub').textContent = 'Pick a destination - the sent copy will be filed automatically.';

  // Show the file-on-send toggle, hide the standard email preview (compose item has no readable content yet)
  document.getElementById('file-on-send-toggle').style.display = 'flex';
  document.getElementById('email-preview').style.display = 'none';

  // Rename the primary button - in compose mode it saves the destination/tags, doesn't upload anything yet
  const btn = document.getElementById('file-btn');
  btn.textContent = 'Save filing settings';

  // Load any previously-saved compose-mode settings from custom properties
  loadComposeState();

  // Listen for recipient changes so we can auto-suggest a destination
  if (currentItem.to && typeof currentItem.to.addHandlerAsync === 'function') {
    try {
      currentItem.to.addHandlerAsync(Office.EventType.RecipientsChanged, () => onRecipientsChanged());
    } catch (e) { /* not supported on all hosts */ }
  }
}

async function initAuth() {
  const cfg = window.CIP_CONFIG;
  if (!cfg || cfg.clientId.startsWith('YOUR-') || cfg.tenantId.startsWith('YOUR-')) {
    document.getElementById('config-missing').style.display = 'block';
    document.getElementById('signin-btn').disabled = true;
    return;
  }
  // Wait for MSAL
  let tries = 0;
  while (typeof msal === 'undefined' && tries < 50) {
    await new Promise(r => setTimeout(r, 100));
    tries++;
  }
  if (typeof msal === 'undefined') {
    toast('Could not load Microsoft authentication library', 'error');
    return;
  }
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
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    await onSignedIn();
  }
}

async function signIn() {
  if (!msalInstance) return;
  try {
    // Use popup; Office add-ins block top-level redirects
    const response = await msalInstance.loginPopup({ scopes: window.CIP_CONFIG.scopes });
    currentAccount = response.account;
    await onSignedIn();
  } catch (e) {
    if (e.errorCode !== 'user_cancelled') {
      toast('Sign-in failed: ' + (e.errorMessage || e.message), 'error');
    }
  }
}

async function onSignedIn() {
  document.getElementById('signin-screen').style.display = 'none';
  document.getElementById('main-pane').classList.add('signed-in');
  document.getElementById('user-tag').textContent = currentAccount.username;
  try {
    accessToken = await getToken();
    loadRecent();
    await loadSites();

    // Compose-mode: try to auto-suggest a destination from the recipient
    if (mode === 'compose') {
      onRecipientsChanged();
    } else {
      // Read-mode: first auto-file any queued sends, then fall back to the
      // manual post-send banner for anything the poller couldn't match.
      await drainSentQueue();
      checkPostSendPrompt();
    }
  } catch (e) {
    toast('Could not connect to SharePoint: ' + e.message, 'error');
  }
}

async function getToken() {
  const request = { scopes: window.CIP_CONFIG.scopes, account: currentAccount };
  try {
    const r = await msalInstance.acquireTokenSilent(request);
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError) {
      const r = await msalInstance.acquireTokenPopup(request);
      return r.accessToken;
    }
    throw e;
  }
}

// ============================================================
// Email preview + filed-state check
// ============================================================
function populateEmailPreview() {
  const previewEl = document.getElementById('email-preview');
  if (!previewEl) return;

  // Multi-select state: show a "N messages selected" banner instead of single-message detail
  if (multiSelectSupported && selectedItems.length > 1) {
    const subjects = selectedItems.slice(0, 3).map(it => escapeHtml(it.subject || '(no subject)'));
    const more = selectedItems.length > 3 ? ` <span class="batch-more">+ ${selectedItems.length - 3} more</span>` : '';
    previewEl.classList.add('batch-mode');
    previewEl.innerHTML =
      `<div class="batch-banner">
         <div class="batch-count">${selectedItems.length} messages selected</div>
         <div class="batch-list">${subjects.map(s => `<div class="batch-row">• ${s}</div>`).join('')}${more}</div>
         <div class="batch-hint">Destination and tags below will be applied to every selected message. Each is filed independently — one failure won't stop the rest.</div>
       </div>`;
    // The filed-state banner doesn't apply across a batch; hide if showing.
    hideFiledBanner();
    return;
  }

  // Single-item path (existing behaviour) — only run if we actually have an item
  previewEl.classList.remove('batch-mode');
  if (!currentItem) {
    previewEl.innerHTML = '<div class="meta-row" style="color:var(--text-muted);">No email selected.</div>';
    return;
  }
  // Restore the original markup if it was replaced by the batch banner on a previous render
  if (!document.getElementById('email-subject')) {
    previewEl.innerHTML =
      `<div class="subject" id="email-subject">Loading email...</div>
       <div class="meta-row"><strong>From:</strong> <span id="email-from">--</span></div>
       <div class="meta-row"><strong>To:</strong> <span id="email-to">--</span></div>
       <div class="meta-row"><strong>Date:</strong> <span id="email-date">--</span></div>
       <div class="meta-row with-attach" id="email-attach-row" style="display:none;">📎 <span id="email-attach-count">0</span> attachment(s)</div>`;
  }
  document.getElementById('email-subject').textContent = currentItem.subject || '(no subject)';
  const fromVal = currentItem.from ? `${currentItem.from.displayName || currentItem.from.emailAddress} <${currentItem.from.emailAddress}>` : '-';
  document.getElementById('email-from').textContent = fromVal;
  const to = (currentItem.to || []).map(r => r.displayName || r.emailAddress).join(', ');
  document.getElementById('email-to').textContent = to || '-';
  document.getElementById('email-date').textContent = currentItem.dateTimeCreated ? new Date(currentItem.dateTimeCreated).toLocaleString() : '-';
  const attachments = currentItem.attachments || [];
  if (attachments.length > 0) {
    document.getElementById('email-attach-row').style.display = 'block';
    document.getElementById('email-attach-count').textContent = attachments.length;
  }
  // After preview, check whether this email has already been filed
  checkFiledState();
}

// ============================================================
// Filed-state marker (custom properties on the mail item)
// ============================================================
// We store these keys on the item via Office.context.mailbox.item.loadCustomPropertiesAsync:
//   cip_filed         - "true" if previously filed
//   cip_filed_at      - ISO timestamp of when filed
//   cip_filed_by      - UPN of the technician who filed it
//   cip_filed_url     - SharePoint URL of the filed .eml
//   cip_filed_site    - SharePoint site display name
//   cip_filed_lib     - Library display name
//   cip_filed_id      - SharePoint driveItem id (used to look up the file later)
//
// Custom properties live in the user's mailbox metadata on this specific message,
// per-add-in. They sync across devices and survive Outlook restarts.

let cachedCustomProps = null;
let pendingRefile = false;

function checkFiledState() {
  if (!Office || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item) return;
  Office.context.mailbox.item.loadCustomPropertiesAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.warn('Could not load custom properties:', result.error);
      return;
    }
    cachedCustomProps = result.value;
    const filed = cachedCustomProps.get('cip_filed');
    if (filed === 'true') {
      showFiledBanner({
        filedAt: cachedCustomProps.get('cip_filed_at'),
        filedBy: cachedCustomProps.get('cip_filed_by'),
        filedUrl: cachedCustomProps.get('cip_filed_url'),
        filedSite: cachedCustomProps.get('cip_filed_site'),
        filedLib: cachedCustomProps.get('cip_filed_lib')
      });
    } else {
      hideFiledBanner();
    }
  });
}

function showFiledBanner(info) {
  const banner = document.getElementById('filed-banner');
  const details = document.getElementById('filed-details');
  const openBtn = document.getElementById('filed-open-btn');
  let html = '';
  if (info.filedSite || info.filedLib) {
    html += `<div class="row"><strong>Location:</strong> ${escapeHtml(info.filedLib || '-')}${info.filedSite ? ' on ' + escapeHtml(info.filedSite) : ''}</div>`;
  }
  if (info.filedBy) {
    html += `<div class="row"><strong>Filed by:</strong> ${escapeHtml(info.filedBy)}</div>`;
  }
  if (info.filedAt) {
    try {
      const d = new Date(info.filedAt);
      html += `<div class="row"><strong>Filed at:</strong> ${d.toLocaleString()}</div>`;
    } catch (e) { /* ignore parse */ }
  }
  details.innerHTML = html || '<div class="row">No additional details recorded.</div>';
  if (info.filedUrl) {
    openBtn.style.display = '';
    openBtn.onclick = () => window.open(info.filedUrl, '_blank');
  } else {
    openBtn.style.display = 'none';
  }
  banner.classList.add('show');
  // Subtle reminder: dim the filing form a touch
  document.querySelectorAll('.form-section').forEach(s => { s.style.opacity = '0.85'; });
}

function hideFiledBanner() {
  document.getElementById('filed-banner').classList.remove('show');
  document.querySelectorAll('.form-section').forEach(s => { s.style.opacity = ''; });
}

function isAlreadyFiled() {
  return cachedCustomProps && cachedCustomProps.get('cip_filed') === 'true';
}

function confirmRefile() {
  if (!isAlreadyFiled()) { hideFiledBanner(); return; }
  // Open the modal
  const prev = document.getElementById('refile-prev');
  const filedAt = cachedCustomProps.get('cip_filed_at');
  const filedBy = cachedCustomProps.get('cip_filed_by');
  const filedLib = cachedCustomProps.get('cip_filed_lib');
  const filedSite = cachedCustomProps.get('cip_filed_site');
  let prevHtml = '<strong>Previously filed:</strong><br>';
  if (filedLib || filedSite) prevHtml += `${escapeHtml(filedLib || '-')}${filedSite ? ' · ' + escapeHtml(filedSite) : ''}<br>`;
  if (filedBy) prevHtml += `by ${escapeHtml(filedBy)}`;
  if (filedAt) {
    try { prevHtml += ` on ${new Date(filedAt).toLocaleString()}`; } catch (e) {}
  }
  prev.innerHTML = prevHtml;
  document.getElementById('refile-modal').classList.add('show');
}

function closeRefileModal() {
  document.getElementById('refile-modal').classList.remove('show');
  pendingRefile = false;
}

function proceedWithFiling() {
  document.getElementById('refile-modal').classList.remove('show');
  pendingRefile = true;
  // Actually perform the filing now
  doFileEmail();
}

// ============================================================
// SharePoint Graph calls
// ============================================================
async function graph(path, opts) {
  const o = opts || {};
  const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
  o.headers = Object.assign({
    'Authorization': 'Bearer ' + accessToken,
    'Content-Type': o.contentType || 'application/json',
    'Accept': 'application/json'
  }, o.headers || {});
  if (o.contentType) delete o.contentType;
  const res = await fetch(url, o);
  if (res.status === 401) {
    accessToken = await getToken();
    o.headers['Authorization'] = 'Bearer ' + accessToken;
    const retry = await fetch(url, o);
    if (!retry.ok) throw new Error('Graph ' + retry.status + ': ' + await retry.text());
    return retry.status === 204 ? null : (retry.headers.get('content-type')?.includes('json') ? retry.json() : retry);
  }
  if (!res.ok) throw new Error('Graph ' + res.status + ': ' + await res.text());
  if (res.status === 204) return null;
  return res.headers.get('content-type')?.includes('json') ? res.json() : res;
}

async function loadSites() {
  try {
    const result = await graph('/sites?search=*&$top=100');
    allSites = result.value || [];
    // Apply optional site filter
    const filter = window.CIP_CONFIG.siteFilter || [];
    if (filter.length > 0) {
      allSites = allSites.filter(s => filter.some(f => s.webUrl.toLowerCase().includes(f.toLowerCase())));
    }
    allSites.sort((a, b) => (a.displayName || a.name).localeCompare(b.displayName || b.name));
    renderSites(allSites);
  } catch (e) {
    document.getElementById('site-list-items').innerHTML = '<div class="picker-empty">Could not load sites: ' + e.message + '</div>';
  }
}

function renderSites(sites) {
  const container = document.getElementById('site-list-items');
  if (sites.length === 0) {
    container.innerHTML = '<div class="picker-empty">No matching sites</div>';
    return;
  }
  container.innerHTML = sites.map(s => {
    const url = s.webUrl || '';
    return `<div class="picker-item" onclick="selectSite('${s.id}')">
      <div class="label">${escapeHtml(s.displayName || s.name || 'Unnamed site')}</div>
      <div class="url">${escapeHtml(url)}</div>
    </div>`;
  }).join('');
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

function openSitePicker() {
  closeAllPickers();
  document.getElementById('site-list').classList.add('show');
  document.getElementById('site-search').focus();
}
function openLibPicker() {
  if (!selectedSite) { toast('Pick a site first', 'info'); return; }
  closeAllPickers();
  document.getElementById('lib-list').classList.add('show');
}
function closeAllPickers() {
  document.querySelectorAll('.picker-list').forEach(p => p.classList.remove('show'));
}

document.addEventListener('click', e => {
  if (!e.target.closest('.picker-row')) closeAllPickers();
});

async function selectSite(siteId) {
  selectedSite = allSites.find(s => s.id === siteId);
  selectedLib = null;
  document.getElementById('site-picker').classList.remove('empty');
  document.getElementById('site-picker-label').textContent = selectedSite.displayName || selectedSite.name;
  document.getElementById('lib-picker-label').textContent = 'Loading libraries...';
  document.getElementById('lib-picker').classList.add('empty');
  closeAllPickers();
  // Load doc libraries for this site
  try {
    const result = await graph('/sites/' + siteId + '/drives');
    allLibraries = (result.value || []).filter(d => d.driveType === 'documentLibrary');
    allLibraries.sort((a, b) => a.name.localeCompare(b.name));
    renderLibraries(allLibraries);
    // Auto-select preferred library if it exists
    const preferred = window.CIP_CONFIG.preferredLibrary;
    if (preferred) {
      const match = allLibraries.find(l => l.name === preferred);
      if (match) selectLibrary(match.id);
    }
    if (!selectedLib) {
      document.getElementById('lib-picker-label').textContent = 'Pick a library...';
    }
  } catch (e) {
    document.getElementById('lib-list-items').innerHTML = '<div class="picker-empty">Could not load libraries</div>';
  }
}

function renderLibraries(libs) {
  const container = document.getElementById('lib-list-items');
  if (libs.length === 0) {
    container.innerHTML = '<div class="picker-empty">No document libraries</div>';
    return;
  }
  container.innerHTML = libs.map(l => `<div class="picker-item" onclick="selectLibrary('${l.id}')">
    <div class="label">${escapeHtml(l.name)}</div>
    <div class="url">${l.driveType} · ${escapeHtml(l.description || '')}</div>
  </div>`).join('');
}

function selectLibrary(libId) {
  selectedLib = allLibraries.find(l => l.id === libId);
  document.getElementById('lib-picker').classList.remove('empty');
  document.getElementById('lib-picker-label').textContent = selectedLib.name;
  closeAllPickers();
  // Refresh compose-mode hint if applicable
  if (mode === 'compose') updateFileOnSendState();
}

// ============================================================
// Primary button dispatcher - branches by mode
// ============================================================
function primaryAction() {
  if (mode === 'compose') {
    saveComposeSettings();
  } else if (multiSelectSupported && selectedItems.length > 1) {
    fileSelectedItems();
  } else {
    fileEmail();
  }
}

// ============================================================
// COMPOSE MODE - configure destination before sending
// ============================================================
// Reads the recipient list, looks up the clientMap from config, and pre-fills
// the site/library pickers. Saves the chosen destination as custom properties
// on the compose item so we can find it again when the email opens in read mode
// after sending (post-send prompt flow).
// ============================================================

let lastRecipientHash = '';

async function onRecipientsChanged() {
  if (mode !== 'compose' || !currentItem || !currentItem.to) return;
  if (typeof currentItem.to.getAsync !== 'function') return;
  currentItem.to.getAsync(async result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;
    const recipients = result.value || [];
    const emails = recipients.map(r => (r.emailAddress || '').toLowerCase()).filter(Boolean);
    const hash = emails.join(',');
    if (hash === lastRecipientHash) return;
    lastRecipientHash = hash;
    if (emails.length === 0) return;

    const match = findClientMatch(emails);
    if (match) {
      await applySuggestion(match, emails);
    }
  });
}

function findClientMatch(emails) {
  const clientMap = (window.CIP_CONFIG && window.CIP_CONFIG.clientMap) || [];
  for (const email of emails) {
    const at = email.indexOf('@');
    if (at < 0) continue;
    const domain = email.slice(at + 1);
    for (const entry of clientMap) {
      if (matchesPattern(domain, entry.match)) {
        return entry;
      }
    }
  }
  return null;
}

function matchesPattern(domain, pattern) {
  if (!pattern) return false;
  pattern = pattern.toLowerCase();
  domain = domain.toLowerCase();
  if (pattern === domain) return true;
  if (pattern.startsWith('*.')) {
    return domain.endsWith(pattern.slice(1)) || domain === pattern.slice(2);
  }
  return domain.includes(pattern);
}

async function applySuggestion(match, emails) {
  // Find a matching site in allSites by hint
  const hint = (match.siteHint || '').toLowerCase();
  const site = allSites.find(s =>
    (s.webUrl || '').toLowerCase().includes(hint) ||
    (s.displayName || '').toLowerCase().includes(hint) ||
    (s.name || '').toLowerCase().includes(hint)
  );
  if (!site) return;

  // Show "suggested" chip on the picker
  showSuggestedChip(emails[0]);

  await selectSite(site.id);
  // Auto-select library by hint (otherwise preferredLibrary kicks in inside selectSite)
  if (match.libraryHint) {
    const libHint = match.libraryHint.toLowerCase();
    const matchedLib = allLibraries.find(l => l.name.toLowerCase() === libHint || l.name.toLowerCase().includes(libHint));
    if (matchedLib) selectLibrary(matchedLib.id);
  }
  // Auto-fill client field
  if (match.client && !document.getElementById('f-client').value) {
    document.getElementById('f-client').value = match.client;
  }
}

function showSuggestedChip(forEmail) {
  const sitePicker = document.getElementById('site-picker');
  // Remove existing chip first
  const existing = sitePicker.parentElement.querySelector('.suggested-chip');
  if (existing) existing.remove();
  // Add small chip above the picker
  const lbl = sitePicker.parentElement.parentElement.querySelector('label');
  if (lbl && !lbl.querySelector('.suggested-chip')) {
    const chip = document.createElement('span');
    chip.className = 'suggested-chip';
    const domain = (forEmail.split('@')[1] || forEmail);
    chip.textContent = 'matched ' + domain;
    lbl.appendChild(chip);
  }
}

function updateFileOnSendState() {
  const cb = document.getElementById('file-on-send-cb');
  const hint = document.getElementById('file-on-send-hint');
  if (cb.checked) {
    hint.textContent = selectedSite && selectedLib
      ? 'Will file to ' + selectedLib.name + ' on ' + (selectedSite.displayName || selectedSite.name)
      : 'Pick a destination below.';
  } else {
    hint.textContent = 'Filing disabled. Sent copy stays in Sent Items only.';
  }
}

async function saveComposeSettings() {
  const cb = document.getElementById('file-on-send-cb');
  const fileOnSend = cb.checked;
  if (fileOnSend && (!selectedSite || !selectedLib)) {
    toast('Pick a site and library before enabling File on send', 'error');
    return;
  }
  const settings = {
    cip_file_on_send: fileOnSend ? 'true' : 'false'
  };
  if (fileOnSend && selectedSite && selectedLib) {
    settings.cip_pending_site_id = selectedSite.id;
    settings.cip_pending_site_name = selectedSite.displayName || selectedSite.name;
    settings.cip_pending_lib_id = selectedLib.id;
    settings.cip_pending_lib_name = selectedLib.name;
    settings.cip_pending_folder = expandFolderTemplate(document.getElementById('f-folder').value);
    settings.cip_pending_client = document.getElementById('f-client').value;
    settings.cip_pending_project = document.getElementById('f-project').value;
    settings.cip_pending_category = document.getElementById('f-category').value;
    settings.cip_pending_notes = document.getElementById('f-notes').value;
  }

  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.loadCustomPropertiesAsync(result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        toast('Could not save settings: ' + (result.error && result.error.message), 'error');
        return reject();
      }
      const props = result.value;
      for (const [k, v] of Object.entries(settings)) props.set(k, v);
      props.saveAsync(saveResult => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          if (fileOnSend) {
            // Stage the filing intent to the localStorage queue so the Sent
            // Items poller can file the message automatically after it's sent -
            // no need for the user to open the sent copy.
            stageSendQueueEntry(settings);
            toast('Done. This email will be filed automatically after you send it.', 'success');
          } else {
            toast('File-on-send disabled for this email.', 'success');
          }
          resolve();
        } else {
          toast('Could not save: ' + (saveResult.error && saveResult.error.message), 'error');
          reject();
        }
      });
    });
  });
}

function loadComposeState() {
  if (!currentItem || !currentItem.loadCustomPropertiesAsync) return;
  currentItem.loadCustomPropertiesAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;
    const props = result.value;
    if (props.get('cip_file_on_send') === 'true') {
      document.getElementById('file-on-send-cb').checked = true;
      // Restore form fields if present
      if (props.get('cip_pending_client')) document.getElementById('f-client').value = props.get('cip_pending_client');
      if (props.get('cip_pending_project')) document.getElementById('f-project').value = props.get('cip_pending_project');
      if (props.get('cip_pending_category')) document.getElementById('f-category').value = props.get('cip_pending_category');
      if (props.get('cip_pending_notes')) document.getElementById('f-notes').value = props.get('cip_pending_notes');
      if (props.get('cip_pending_folder')) document.getElementById('f-folder').value = props.get('cip_pending_folder');
      updateFileOnSendState();
    }
  });
}

// ============================================================
// AUTOMATIC FILE-ON-SEND (Sent Items poller)
// ============================================================
// When the user arms file-on-send in compose mode, we stage the destination
// and tags to a localStorage queue keyed by recipient + arm time. After the
// message is sent, the next time the taskpane loads (read mode) we poll Sent
// Items via Graph, match the queued intent to the actual sent message, and
// file it automatically - no banner, no opening the sent copy.
//
// Limitations (documented, acceptable for single-device use):
//   - The queue lives in localStorage, so arming on one device and filing on
//     another won't match. The manual post-send banner remains as a fallback.
//   - Matching is by recipient + sent-after-arm-time. Sending two near-identical
//     messages to the same recipient in quick succession could mis-match; the
//     earliest unfiled match wins.
// ============================================================

const SEND_QUEUE_KEY = 'cip-file-on-send-queue';
const FILED_IDS_KEY  = 'cip-filed-message-ids';

function readSendQueue() {
  try { return JSON.parse(localStorage.getItem(SEND_QUEUE_KEY) || '[]'); }
  catch (e) { return []; }
}
function writeSendQueue(q) {
  try { localStorage.setItem(SEND_QUEUE_KEY, JSON.stringify(q)); } catch (e) {}
}
function readFiledSet() {
  try { return JSON.parse(localStorage.getItem(FILED_IDS_KEY) || '[]'); }
  catch (e) { return []; }
}
function addToFiledSet(internetMessageId) {
  if (!internetMessageId) return;
  try {
    const s = readFiledSet();
    if (!s.includes(internetMessageId)) {
      s.push(internetMessageId);
      // Cap the set so it doesn't grow forever
      while (s.length > 500) s.shift();
      localStorage.setItem(FILED_IDS_KEY, JSON.stringify(s));
    }
  } catch (e) {}
}

// Capture recipients + subject from the compose item and push a queue entry.
function stageSendQueueEntry(settings) {
  const item = Office.context.mailbox.item;
  if (!item || !item.to || typeof item.to.getAsync !== 'function') return;

  item.to.getAsync(res => {
    const recips = (res.status === Office.AsyncResultStatus.Succeeded && Array.isArray(res.value))
      ? res.value.map(r => (r.emailAddress || '').toLowerCase()).filter(Boolean)
      : [];

    const finalise = (subject) => {
      const q = readSendQueue();
      q.push({
        id: 'q_' + Date.now() + '_' + Math.random().toString(36).slice(2, 7),
        armedAt: new Date().toISOString(),
        recipients: recips,
        subject: subject || '',
        dest: {
          siteId: settings.cip_pending_site_id,
          siteName: settings.cip_pending_site_name,
          libId: settings.cip_pending_lib_id,
          libName: settings.cip_pending_lib_name,
          folder: settings.cip_pending_folder || ''
        },
        tags: {
          client: settings.cip_pending_client || '',
          project: settings.cip_pending_project || '',
          category: settings.cip_pending_category || '',
          notes: settings.cip_pending_notes || ''
        },
        attempts: 0
      });
      writeSendQueue(q);
    };

    if (item.subject && typeof item.subject.getAsync === 'function') {
      item.subject.getAsync(sres =>
        finalise(sres.status === Office.AsyncResultStatus.Succeeded ? sres.value : ''));
    } else {
      finalise('');
    }
  });
}

// Poll Sent Items and file any queued intents whose message has now been sent.
// Path of the per-user OneDrive pending-filing queue (written by the dialog
// flow in dialog.js). This is the bridge that lets dialog-armed sends reach
// this poller across the isolated dialog storage context.
const ONEDRIVE_PENDING_PATH = '/me/drive/root:/cip-file-email-pending.json';

async function readOneDriveQueue() {
  // Read the file's metadata first to get a pre-authed download URL. This is
  // more reliable than the /content redirect, which can fail CORS when fetched
  // with an Authorization header from an add-in origin.
  let meta;
  try {
    meta = await graph(ONEDRIVE_PENDING_PATH);   // DriveItem JSON (incl. downloadUrl)
  } catch (e) {
    if (/404/.test(e && e.message || '')) { console.log('[cip] OneDrive queue: no file yet'); return []; }
    throw e;
  }

  let txt = null;
  const dl = meta && meta['@microsoft.graph.downloadUrl'];
  if (dl) {
    try {
      const r = await fetch(dl);              // pre-authed short-lived URL - no auth header
      if (r.ok) txt = await r.text();
      else console.warn('[cip] OneDrive downloadUrl returned', r.status);
    } catch (e) {
      console.warn('[cip] OneDrive downloadUrl fetch failed, trying /content:', e && e.message);
    }
  }

  if (txt === null) {
    // Fallback to the /content endpoint.
    const res = await graph(ONEDRIVE_PENDING_PATH + ':/content');
    if (Array.isArray(res)) { console.log('[cip] OneDrive queue entries:', res.length); return res; }
    if (res && typeof res.text === 'function') txt = await res.text();
    else if (typeof res === 'string') txt = res;
    else return [];
  }

  try {
    const arr = JSON.parse(txt);
    console.log('[cip] OneDrive queue entries:', Array.isArray(arr) ? arr.length : 0);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    console.warn('[cip] could not parse OneDrive queue JSON:', e && e.message);
    return [];
  }
}

async function writeOneDriveQueue(arr) {
  await graph(ONEDRIVE_PENDING_PATH + ':/content', {
    method: 'PUT',
    contentType: 'application/json',
    body: JSON.stringify(arr || [])
  });
}

function stripSource(e) { const c = Object.assign({}, e); delete c._source; return c; }

// Drain BOTH queues: the localStorage queue (compose-taskpane flow) and the
// OneDrive queue (dialog flow). Match each pending entry to a sent message and
// file it, then write back whatever didn't match yet.
async function drainSentQueue() {
  const now = Date.now();
  const MAX_AGE_MS = 24 * 60 * 60 * 1000;   // expire after 24h
  const MAX_ATTEMPTS = 8;
  const fresh = e => (now - new Date(e.armedAt).getTime()) < MAX_AGE_MS && (e.attempts || 0) < MAX_ATTEMPTS;

  const localQ = readSendQueue().filter(fresh);

  let driveQ = [], driveReadOk = false;
  try { driveQ = (await readOneDriveQueue()).filter(fresh); driveReadOk = true; }
  catch (e) { console.warn('Could not read OneDrive pending queue:', e && e.message); }

  const entries = localQ.map(e => { e._source = 'local'; return e; })
    .concat(driveQ.map(e => { e._source = 'onedrive'; return e; }));

  console.log('[cip] drainSentQueue: local=' + localQ.length + ' onedrive=' + driveQ.length);

  if (!entries.length) {
    writeSendQueue(localQ.map(stripSource));
    if (driveReadOk) { try { await writeOneDriveQueue(driveQ.map(stripSource)); } catch (e) {} }
    return;
  }

  let sent = [];
  try {
    const res = await graph('/me/mailFolders/sentitems/messages?$top=25&$orderby=sentDateTime desc&$select=id,subject,toRecipients,sentDateTime,internetMessageId');
    sent = res.value || [];
    console.log('[cip] Sent Items fetched:', sent.length);
  } catch (e) {
    console.warn('Sent Items poll failed:', e && e.message);
    return;   // leave both queues untouched; retry next time the taskpane opens
  }

  const filedSet = readFiledSet();
  const keepLocal = [], keepDrive = [];

  for (const entry of entries) {
    const armedMs = new Date(entry.armedAt).getTime();
    const candidates = sent.filter(m => {
      if (filedSet.includes(m.internetMessageId)) return false;
      const sentMs = new Date(m.sentDateTime).getTime();
      if (sentMs < armedMs - 60000) return false;   // 1 min clock-skew tolerance
      const toAddrs = (m.toRecipients || []).map(r =>
        (r.emailAddress && r.emailAddress.address || '').toLowerCase());
      return !entry.recipients || entry.recipients.length === 0 ||
        entry.recipients.some(r => toAddrs.includes(r));
    }).sort((a, b) => new Date(a.sentDateTime) - new Date(b.sentDateTime));

    const match = candidates[0];
    if (!match) {
      console.log('[cip] no match yet for entry armed at', entry.armedAt, 'to', (entry.recipients || []).join(','));
      (entry._source === 'onedrive' ? keepDrive : keepLocal).push(entry);
      continue;
    }

    console.log('[cip] match found, filing:', match.subject);
    try {
      await fileQueuedMessage(match, entry);
      addToFiledSet(match.internetMessageId);
      console.log('[cip] filed OK:', match.subject);
      // filed - entry consumed
    } catch (e) {
      console.warn('[cip] Auto-file failed (will retry):', e && e.message);
      entry.attempts = (entry.attempts || 0) + 1;
      (entry._source === 'onedrive' ? keepDrive : keepLocal).push(entry);
    }
  }

  writeSendQueue(keepLocal.map(stripSource));
  if (driveReadOk) { try { await writeOneDriveQueue(keepDrive.map(stripSource)); } catch (e) {} }
}

// Upload one sent message to its queued destination. Reuses the existing
// upload helpers by temporarily pointing the destination globals at the
// queued site/library, then restoring them.
async function fileQueuedMessage(sentMsg, entry) {
  const prevSite = selectedSite, prevLib = selectedLib;
  selectedSite = { id: entry.dest.siteId, displayName: entry.dest.siteName, name: entry.dest.siteName };
  selectedLib  = { id: entry.dest.libId, name: entry.dest.libName };

  try {
    accessToken = await getToken();

    const mime = await getEmailMime(sentMsg.id);   // sentMsg.id is already a Graph id
    const folderPath = entry.dest.folder ? await ensureFolder(entry.dest.folder) : '';
    const fauxItem = {
      subject: sentMsg.subject || 'No subject',
      dateTimeCreated: sentMsg.sentDateTime || new Date().toISOString()
    };
    const filename = makeFilename(fauxItem);
    const uploaded = await uploadFile(folderPath, filename, mime);

    const itemShape = {
      subject: sentMsg.subject || '',
      fromName: currentAccount ? currentAccount.name : '',
      fromAddress: currentAccount ? currentAccount.username : '',
      toList: (sentMsg.toRecipients || []).map(r => r.emailAddress && r.emailAddress.address).filter(Boolean).join('; '),
      date: sentMsg.sentDateTime || null
    };
    await setMetadataForItem(uploaded, itemShape, entry.tags);
    await markEmailInOutlook(sentMsg.id);

    addToRecent({
      subject: sentMsg.subject || '(no subject)',
      siteName: entry.dest.siteName,
      libName: entry.dest.libName,
      filename: filename,
      url: uploaded.webUrl,
      date: new Date().toISOString()
    });

    toast('Filed sent email: ' + (sentMsg.subject || '(no subject)'), 'success');
  } finally {
    selectedSite = prevSite;
    selectedLib = prevLib;
  }
}

// ============================================================
// POST-SEND PROMPT - shown in read mode when opening a just-sent email that was marked for filing
// ============================================================

function checkPostSendPrompt() {
  if (!currentItem || mode !== 'read') return;
  // If the poller already auto-filed this message, don't offer to file it again.
  if (currentItem.internetMessageId && readFiledSet().includes(currentItem.internetMessageId)) return;
  currentItem.loadCustomPropertiesAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;
    const props = result.value;
    if (props.get('cip_file_on_send') !== 'true') return;
    // Already filed? Don't prompt again
    if (props.get('cip_filed') === 'true') return;

    pendingSendInfo = {
      siteId: props.get('cip_pending_site_id'),
      siteName: props.get('cip_pending_site_name'),
      libId: props.get('cip_pending_lib_id'),
      libName: props.get('cip_pending_lib_name'),
      folder: props.get('cip_pending_folder'),
      client: props.get('cip_pending_client'),
      project: props.get('cip_pending_project'),
      category: props.get('cip_pending_category'),
      notes: props.get('cip_pending_notes')
    };

    document.getElementById('send-banner').style.display = 'block';
    document.getElementById('send-banner-msg').textContent =
      'You marked this email for filing to ' + (pendingSendInfo.libName || '-') +
      ' on ' + (pendingSendInfo.siteName || '-') + '. Confirm to file it now.';
  });
}

async function fileSentEmail() {
  if (!pendingSendInfo) return;
  // Apply pending settings to the form state so doFileEmail uses them
  // We need to populate selectedSite/selectedLib by looking them up if they're not already loaded
  selectedSite = allSites.find(s => s.id === pendingSendInfo.siteId);
  if (!selectedSite) {
    // Site not in the current list - synthesize a minimal record so the upload still works
    selectedSite = { id: pendingSendInfo.siteId, displayName: pendingSendInfo.siteName, name: pendingSendInfo.siteName };
  }
  // For the library, we need to fetch drives for this site to get the id
  try {
    if (!selectedLib || selectedLib.id !== pendingSendInfo.libId) {
      const result = await graph('/sites/' + selectedSite.id + '/drives');
      allLibraries = (result.value || []).filter(d => d.driveType === 'documentLibrary');
      selectedLib = allLibraries.find(l => l.id === pendingSendInfo.libId);
      if (!selectedLib) {
        // Try matching by name as a fallback
        selectedLib = allLibraries.find(l => l.name === pendingSendInfo.libName);
      }
    }
  } catch (e) {
    toast('Could not resolve filing destination: ' + e.message, 'error');
    return;
  }
  if (!selectedLib) {
    toast('Filing library not found - it may have been renamed', 'error');
    return;
  }

  // Populate form fields so metadata is set correctly
  if (pendingSendInfo.client) document.getElementById('f-client').value = pendingSendInfo.client;
  if (pendingSendInfo.project) document.getElementById('f-project').value = pendingSendInfo.project;
  if (pendingSendInfo.category) document.getElementById('f-category').value = pendingSendInfo.category;
  if (pendingSendInfo.notes) document.getElementById('f-notes').value = pendingSendInfo.notes;
  if (pendingSendInfo.folder) document.getElementById('f-folder').value = pendingSendInfo.folder;

  // Hide the prompt and run the standard filing flow
  document.getElementById('send-banner').style.display = 'none';
  await doFileEmail();
  pendingSendInfo = null;
}

function dismissSendPrompt() {
  // Clear the pending flag so we don't keep prompting
  if (!currentItem) return;
  currentItem.loadCustomPropertiesAsync(result => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) return;
    const props = result.value;
    props.set('cip_file_on_send', 'dismissed');
    props.saveAsync(() => {
      document.getElementById('send-banner').style.display = 'none';
      toast('Prompt dismissed. You can still file manually from this email at any time.', 'info');
    });
  });
}

// ============================================================
// File the email
// ============================================================
async function fileEmail() {
  // Gateway: if already filed and the user didn't explicitly confirm via the modal, prompt them
  if (isAlreadyFiled() && !pendingRefile) {
    confirmRefile();
    return;
  }
  await doFileEmail();
}

async function doFileEmail() {
  pendingRefile = false;  // reset for next time
  if (!selectedSite || !selectedLib) { toast('Pick a SharePoint site and library', 'error'); return; }
  if (!currentItem) { toast('No email selected', 'error'); return; }

  const fileBtn = document.getElementById('file-btn');
  fileBtn.disabled = true;
  fileBtn.textContent = 'Filing...';

  // Compute the Graph id once. Used by both getEmailMime (to fetch MIME) and
  // markEmailInOutlook (to apply the category + subject prefix at the end).
  let graphId;
  try {
    graphId = Office.context.mailbox.convertToRestId(currentItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
  } catch (e) {
    graphId = currentItem.itemId;
  }

  showProgress([
    { id: 'fetch', label: 'Fetching email content from Microsoft Graph', status: 'active' },
    { id: 'folder', label: 'Preparing destination folder', status: 'pending' },
    { id: 'upload', label: 'Uploading .eml file to SharePoint', status: 'pending' },
    { id: 'meta', label: 'Tagging with metadata', status: 'pending' },
    { id: 'mark', label: 'Marking email as filed', status: 'pending' }
  ]);

  try {
    // 1. Get the email MIME content via EWS (Outlook desktop & web)
    setStep('fetch', 'active');
    const mime = await getEmailMime(graphId);
    setStep('fetch', 'done');

    // 2. Resolve / create the destination folder
    setStep('folder', 'active');
    const folder = expandFolderTemplate(document.getElementById('f-folder').value);
    const folderPath = folder ? await ensureFolder(folder) : '';
    setStep('folder', 'done');

    // 3. Upload .eml
    setStep('upload', 'active');
    const filename = makeFilename(currentItem);
    const uploadedItem = await uploadFile(folderPath, filename, mime);
    setStep('upload', 'done');

    // 4. Set metadata on the list item
    setStep('meta', 'active');
    await setMetadata(uploadedItem);
    setStep('meta', 'done');

    // 5. Mark the email as filed (custom props for the add-in's own logic,
    //    plus an Outlook category and subject prefix that are visible to the
    //    user in their mail list without opening the taskpane).
    setStep('mark', 'active');
    await saveFiledMarkers({
      url: uploadedItem.webUrl,
      driveItemId: uploadedItem.id,
      siteName: selectedSite.displayName || selectedSite.name,
      libName: selectedLib.name,
      filedBy: currentAccount.username
    });
    await markEmailInOutlook(graphId);
    setStep('mark', 'done');

    // Done
    showResult({
      success: true,
      filename: filename,
      url: uploadedItem.webUrl,
      libName: selectedLib.name,
      siteName: selectedSite.displayName || selectedSite.name
    });

    // Save to recent
    addToRecent({
      subject: currentItem.subject || '(no subject)',
      siteName: selectedSite.displayName || selectedSite.name,
      libName: selectedLib.name,
      filename: filename,
      url: uploadedItem.webUrl,
      date: new Date().toISOString()
    });

    // Refresh the banner with the new filing info
    checkFiledState();
  } catch (e) {
    console.error(e);
    showResult({ success: false, error: e.message });
  } finally {
    fileBtn.disabled = false;
    fileBtn.textContent = 'File to SharePoint';
  }
}

// Save filed-state markers to the mail item's custom properties (per-add-in storage)
function saveFiledMarkers(info) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.loadCustomPropertiesAsync(result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        return reject(new Error('loadCustomProperties failed: ' + (result.error && result.error.message)));
      }
      const props = result.value;
      props.set('cip_filed', 'true');
      props.set('cip_filed_at', new Date().toISOString());
      props.set('cip_filed_by', info.filedBy || '');
      props.set('cip_filed_url', info.url || '');
      props.set('cip_filed_site', info.siteName || '');
      props.set('cip_filed_lib', info.libName || '');
      if (info.driveItemId) props.set('cip_filed_id', info.driveItemId);
      props.saveAsync(saveResult => {
        if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
          cachedCustomProps = props;
          resolve();
        } else {
          // The upload succeeded but we couldn't write the marker - log and continue
          console.warn('Could not save filed markers: ' + (saveResult.error && saveResult.error.message));
          resolve();  // don't fail the whole flow over the marker
        }
      });
    });
  });
}

// Mark the original message in Outlook itself so the filing is visible
// in the user's mail list without opening the add-in. Adds an Outlook
// category (a coloured pill) and prepends a subject prefix. Both signals
// are configurable in config.js. Requires Graph Mail.ReadWrite.
//
// graphId is the message's REST/Graph id (already converted from EWS).
// Returns nothing - failures are logged and swallowed, because the
// SharePoint-side filing has already succeeded and we don't want a
// labelling glitch to look like a filing failure.
async function markEmailInOutlook(graphId) {
  const cfg = window.CIP_CONFIG || {};
  if (!cfg.markFiledInOutlook) return;

  const wantCategory = !!(cfg.filedCategory && cfg.filedCategory.trim());
  const wantPrefix = !!(cfg.filedSubjectPrefix && cfg.filedSubjectPrefix.trim());
  if (!wantCategory && !wantPrefix) return;

  // Read current state once (to avoid clobbering existing categories / re-prefixing).
  let msg;
  try {
    msg = await graph('/me/messages/' + graphId + '?$select=categories,subject');
  } catch (e) {
    console.warn('Could not read message to mark as filed:', e && e.message);
    return;
  }

  // 1. CATEGORY - applied in its OWN request. Categories are writable on any
  //    message (sent or received) with Mail.ReadWrite. This is the reliable
  //    "filed" marker. It must NOT be combined with the subject update below,
  //    because a subject failure would otherwise take the category down with it.
  if (wantCategory) {
    try {
      const existing = Array.isArray(msg.categories) ? msg.categories.slice() : [];
      if (!existing.includes(cfg.filedCategory)) {
        existing.push(cfg.filedCategory);
        await graph('/me/messages/' + graphId, {
          method: 'PATCH',
          body: JSON.stringify({ categories: existing })
        });
      }
    } catch (e) {
      // 403 here usually means Mail.ReadWrite hasn't been consented yet.
      console.warn('Could not set filed category:', e && e.message);
    }
  }

  // 2. SUBJECT PREFIX - SEPARATE, best-effort request. Microsoft Graph only
  //    permits editing a message's subject while it is a draft (isDraft = true).
  //    Filed messages are always already sent/received, so this normally fails
  //    with a 400 - that is EXPECTED and harmless. It's kept isolated so it can
  //    never block the category above. (Desktop COM add-ins like CloudFiler can
  //    edit subjects via MAPI; a web add-in using Graph cannot.)
  if (wantPrefix) {
    try {
      const today = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
      const newPrefix = cfg.filedSubjectPrefix.replace(/\{date\}/g, today);
      const currentSubject = msg.subject || '';
      const existingPrefixRegex = /^\[Filed on \d{4}-\d{2}-\d{2}\]\s*/;
      const newSubject = existingPrefixRegex.test(currentSubject)
        ? currentSubject.replace(existingPrefixRegex, newPrefix)
        : newPrefix + currentSubject;
      if (newSubject !== currentSubject) {
        await graph('/me/messages/' + graphId, {
          method: 'PATCH',
          body: JSON.stringify({ subject: newSubject })
        });
      }
    } catch (e) {
      // Expected on sent/received messages - Graph allows subject edits on drafts only.
      console.info('Subject prefix not applied (Graph allows subject edits on drafts only):', e && e.message);
    }
  }
}

// Get the email as MIME using Microsoft Graph
// Note: Microsoft began disabling legacy Exchange tokens globally in Feb 2025,
// which broke Office.context.mailbox.getCallbackTokenAsync({ isRest: true }).
// We now use the Graph access token (already acquired via MSAL) to fetch the MIME content
// from https://graph.microsoft.com/v1.0/me/messages/{id}/$value
//
// graphIdArg: optional. When provided (e.g. from the multi-select batch loop) it's used
// directly. When omitted, falls back to the currently-bound mail item's id.
function getEmailMime(graphIdArg) {
  return new Promise(async (resolve, reject) => {
    try {
      // Ensure we have a fresh Graph access token (the one acquired at sign-in might be stale)
      const token = await getToken();

      let graphId = graphIdArg;
      if (!graphId) {
        // The Office itemId is in EWS format; Graph needs its REST/Graph-compatible form.
        // convertToRestId still works for the ID translation even though the REST endpoint itself is deprecated.
        try {
          graphId = Office.context.mailbox.convertToRestId(currentItem.itemId, Office.MailboxEnums.RestVersion.v2_0);
        } catch (e) {
          // In newer Outlook builds the itemId is already in REST format
          graphId = currentItem.itemId;
        }
      }

      const url = 'https://graph.microsoft.com/v1.0/me/messages/' + graphId + '/$value';
      const r = await fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
      if (!r.ok) {
        const errText = await r.text().catch(() => '');
        throw new Error('Graph ' + r.status + (errText ? ': ' + errText.slice(0, 200) : ''));
      }
      const blob = await r.blob();
      resolve(blob);
    } catch (e) {
      reject(new Error('Failed to fetch email from Microsoft Graph: ' + (e.message || e)));
    }
  });
}

function makeFilename(item) {
  const d = item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date();
  const stamp = d.getFullYear() + ('0' + (d.getMonth()+1)).slice(-2) + ('0' + d.getDate()).slice(-2);
  const subj = (item.subject || 'No subject').replace(/[\/\\:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim().slice(0, 100);
  return `${stamp} - ${subj}.eml`;
}

function expandFolderTemplate(folderInput) {
  const tpl = folderInput || window.CIP_CONFIG.defaultFolderTemplate || '';
  if (!tpl) return '';
  const now = new Date();
  const client = document.getElementById('f-client').value.trim() || 'Unknown';
  const project = document.getElementById('f-project').value.trim() || '';
  return tpl
    .replace(/{YYYY}/g, now.getFullYear())
    .replace(/{MM}/g, ('0' + (now.getMonth()+1)).slice(-2))
    .replace(/{DD}/g, ('0' + now.getDate()).slice(-2))
    .replace(/{client}/g, sanitiseSegment(client))
    .replace(/{project}/g, sanitiseSegment(project))
    .replace(/\/+/g, '/')
    .replace(/^\/|\/$/g, '');
}

function sanitiseSegment(s) {
  return String(s).replace(/[\/\\:*?"<>|#%&{}]/g, '_').trim();
}

async function ensureFolder(folderPath) {
  // Create folders recursively under the drive root
  const segments = folderPath.split('/').filter(Boolean);
  let currentPath = '';
  for (const seg of segments) {
    const parent = currentPath || 'root';
    try {
      // Create with conflictBehavior:replace would overwrite - use 'fail' then ignore conflicts
      await graph('/drives/' + selectedLib.id + '/items/' + parent + '/children', {
        method: 'POST',
        body: JSON.stringify({
          name: seg,
          folder: {},
          '@microsoft.graph.conflictBehavior': 'fail'
        })
      });
    } catch (e) {
      // Folder already exists - ignore
    }
    currentPath = currentPath ? currentPath + '/' + seg : seg;
  }
  return currentPath;
}

async function uploadFile(folderPath, filename, blob) {
  const uploadPath = folderPath ? folderPath + '/' + filename : filename;
  // Use simple PUT for files up to ~4 MB; use upload session for larger
  if (blob.size < 4 * 1024 * 1024) {
    const url = '/drives/' + selectedLib.id + '/root:/' + encodeURIComponent(uploadPath).replace(/%2F/gi, '/') + ':/content';
    const result = await graph(url, {
      method: 'PUT',
      body: blob,
      contentType: 'message/rfc822'
    });
    return result;
  } else {
    // Upload session for larger files
    const session = await graph('/drives/' + selectedLib.id + '/root:/' + encodeURIComponent(uploadPath).replace(/%2F/gi, '/') + ':/createUploadSession', {
      method: 'POST',
      body: JSON.stringify({ item: { '@microsoft.graph.conflictBehavior': 'rename' } })
    });
    const chunkSize = 2 * 1024 * 1024;
    let pos = 0;
    let lastResult = null;
    while (pos < blob.size) {
      const end = Math.min(pos + chunkSize, blob.size);
      const chunk = blob.slice(pos, end);
      const r = await fetch(session.uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': (end - pos).toString(),
          'Content-Range': `bytes ${pos}-${end-1}/${blob.size}`
        },
        body: chunk
      });
      if (!r.ok && r.status !== 202) throw new Error('Chunk upload failed: ' + r.status);
      if (r.status === 200 || r.status === 201) lastResult = await r.json();
      pos = end;
    }
    return lastResult;
  }
}

async function setMetadata(driveItem) {
  // Read the listItem id for this drive item, then update fields
  const listItem = await graph('/drives/' + selectedLib.id + '/items/' + driveItem.id + '/listItem?$expand=fields');
  const fields = {};
  const subj = currentItem.subject || '';
  const fromVal = currentItem.from ? (currentItem.from.emailAddress || '') : '';
  const fromName = currentItem.from ? (currentItem.from.displayName || '') : '';
  const toVal = (currentItem.to || []).map(r => r.emailAddress).join('; ');
  const date = currentItem.dateTimeCreated ? new Date(currentItem.dateTimeCreated).toISOString() : null;
  const client = document.getElementById('f-client').value.trim();
  const project = document.getElementById('f-project').value.trim();
  const category = document.getElementById('f-category').value;
  const notes = document.getElementById('f-notes').value.trim();

  // Only include fields if they exist on the library (try common ones; failures on individual fields are ignored)
  const candidateFields = {
    EmailSubject: subj,
    EmailFrom: fromVal,
    EmailFromName: fromName,
    EmailTo: toVal,
    EmailDate: date,
    FiledClient: client,
    FiledProject: project,
    FiledCategory: category,
    FiledNotes: notes,
    FiledBy: currentAccount.username,
    FiledAt: new Date().toISOString()
  };
  for (const [k, v] of Object.entries(candidateFields)) {
    if (v !== null && v !== '') fields[k] = v;
  }

  try {
    await graph('/drives/' + selectedLib.id + '/items/' + driveItem.id + '/listItem/fields', {
      method: 'PATCH',
      body: JSON.stringify(fields)
    });
  } catch (e) {
    // If specific fields don't exist on the library, retry without them - prefer logging a warning
    console.warn('Some metadata fields could not be set (library schema may differ from recommended):', e.message);
  }
}

// ============================================================
// MULTI-SELECT BATCH FILING
// ============================================================
// Uses Office.context.mailbox.getSelectedItemsAsync (Mailbox requirement set 1.13+)
// to file every currently-selected message in one batch. Destination + tags are
// chosen once and applied to all. Each message is filed independently so a single
// failure won't abort the rest.
//
// Constraints we inherit from Office.js:
//   - Max 100 messages per activation.
//   - Selection must be within a single Exchange mailbox folder (unless Conversations view is on).
//   - Reading Pane must be enabled.
//
// Limitation: when multiple messages are selected, Office.context.mailbox.item is null,
// so we cannot write the per-message custom-property markers (cip_filed, etc.) for
// items that aren't the currently-bound item. The SharePoint record (the .eml + metadata
// columns) is the durable filing record; the in-Outlook "filed" banner won't appear on
// batch-filed messages until they're individually opened. That's a worthwhile follow-up.
// ============================================================

async function fileSelectedItems() {
  if (batchInProgress) return;
  if (!selectedItems || selectedItems.length === 0) {
    toast('No messages selected', 'error');
    return;
  }
  if (!selectedSite || !selectedLib) {
    toast('Pick a SharePoint site and library first', 'error');
    return;
  }

  batchInProgress = true;
  const fileBtn = document.getElementById('file-btn');
  fileBtn.disabled = true;

  const total = selectedItems.length;
  const results = [];

  // Resolve the destination folder once (created on first call, no-op thereafter)
  let folderPath = '';
  try {
    const folder = expandFolderTemplate(document.getElementById('f-folder').value);
    folderPath = folder ? await ensureFolder(folder) : '';
  } catch (e) {
    showResult({ success: false, error: 'Could not prepare destination folder: ' + e.message });
    batchInProgress = false;
    fileBtn.disabled = false;
    updatePrimaryButtonLabel();
    return;
  }

  // Pull the shared tag values from the form once - applied to every item
  const sharedTags = {
    client: document.getElementById('f-client').value.trim(),
    project: document.getElementById('f-project').value.trim(),
    category: document.getElementById('f-category').value,
    notes: document.getElementById('f-notes').value.trim()
  };

  for (let i = 0; i < total; i++) {
    const it = selectedItems[i];
    showBatchProgress(i + 1, total, it.subject || '(no subject)');
    try {
      const filed = await fileOneSelectedMessage(it, folderPath, sharedTags);
      results.push({ subject: it.subject || '(no subject)', ok: true, url: filed.url, filename: filed.filename });
    } catch (err) {
      console.error('Batch item failed:', it.subject, err);
      results.push({ subject: it.subject || '(no subject)', ok: false, error: err.message });
    }
  }

  showBatchSummary(results);
  batchInProgress = false;
  fileBtn.disabled = false;
  updatePrimaryButtonLabel();
}

// File one message identified by a selected-items entry. Performs:
//   1. Resolve the Graph/REST id (selected items return EWS-format ids).
//   2. Fetch message metadata (subject, from, to, date) for the SharePoint columns.
//   3. Fetch MIME via Graph.
//   4. Upload to the pre-resolved folder.
//   5. Write list-item metadata.
// Returns { url, filename }.
async function fileOneSelectedMessage(item, folderPath, sharedTags) {
  // 1. Resolve Graph id
  let graphId = item.itemId;
  try {
    graphId = Office.context.mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0);
  } catch (e) {
    // newer hosts already return REST-format ids
  }

  const token = await getToken();

  // 2. Fetch message metadata for filename + SharePoint columns
  const metaUrl = 'https://graph.microsoft.com/v1.0/me/messages/' + graphId +
    '?$select=subject,from,toRecipients,sentDateTime,receivedDateTime,hasAttachments';
  let meta = {};
  try {
    const r = await fetch(metaUrl, { headers: { 'Authorization': 'Bearer ' + token } });
    if (r.ok) meta = await r.json();
  } catch (e) {
    // metadata is best-effort; we can still file with partial info
  }

  // 3. Fetch MIME content
  const mime = await getEmailMime(graphId);

  // 4. Filename - prefer the message's own date over today's
  const fauxItem = {
    subject: meta.subject || item.subject || 'No subject',
    dateTimeCreated: meta.sentDateTime || meta.receivedDateTime || new Date().toISOString()
  };
  const filename = makeFilename(fauxItem);

  // 5. Upload
  const uploaded = await uploadFile(folderPath, filename, mime);

  // 6. Metadata fields. Build a synthetic "item" shape so setMetadataForItem can
  // pull subject/from/to/date from a known-good source rather than currentItem.
  const fromName = meta.from && meta.from.emailAddress ? (meta.from.emailAddress.name || '') : '';
  const fromAddr = meta.from && meta.from.emailAddress ? (meta.from.emailAddress.address || '') : '';
  const toList = (meta.toRecipients || []).map(r => r.emailAddress && r.emailAddress.address).filter(Boolean).join('; ');
  const itemShape = {
    subject: meta.subject || item.subject || '',
    fromName: fromName,
    fromAddress: fromAddr,
    toList: toList,
    date: meta.sentDateTime || meta.receivedDateTime || null
  };
  await setMetadataForItem(uploaded, itemShape, sharedTags);

  // Outlook-visible filed marker (category + subject prefix). Best-effort -
  // failure here doesn't undo the SharePoint upload that just succeeded.
  await markEmailInOutlook(graphId);

  return { url: uploaded.webUrl, filename: filename };
}

// Write list-item metadata from an explicit shape rather than currentItem.
// Mirrors setMetadata() but doesn't depend on the global currentItem.
async function setMetadataForItem(driveItem, shape, sharedTags) {
  const candidateFields = {
    EmailSubject: shape.subject || '',
    EmailFrom: shape.fromAddress || '',
    EmailFromName: shape.fromName || '',
    EmailTo: shape.toList || '',
    EmailDate: shape.date || null,
    FiledClient: sharedTags.client || '',
    FiledProject: sharedTags.project || '',
    FiledCategory: sharedTags.category || '',
    FiledNotes: sharedTags.notes || '',
    FiledBy: currentAccount ? currentAccount.username : '',
    FiledAt: new Date().toISOString()
  };
  const fields = {};
  for (const [k, v] of Object.entries(candidateFields)) {
    if (v !== null && v !== '') fields[k] = v;
  }
  try {
    await graph('/drives/' + selectedLib.id + '/items/' + driveItem.id + '/listItem/fields', {
      method: 'PATCH',
      body: JSON.stringify(fields)
    });
  } catch (e) {
    console.warn('Some metadata fields could not be set on batch item:', e.message);
  }
}

// Per-iteration progress line. Replaces the regular five-step indicator with a
// batch-level "Filing X of N" status so the user can see steady progress through
// the batch even when individual messages are quick.
function showBatchProgress(current, total, subject) {
  const pct = Math.round((current - 1) / total * 100);
  document.getElementById('status-area').innerHTML = `
    <div class="batch-progress">
      <div class="bp-head">
        <strong>Filing ${current} of ${total}</strong>
        <span class="bp-pct">${pct}%</span>
      </div>
      <div class="bp-bar"><div class="bp-fill" style="width:${pct}%"></div></div>
      <div class="bp-subject">${escapeHtml(subject)}</div>
    </div>`;
}

// Final summary card. Lists every item with a ✓ or ✕, with the error for any failures.
function showBatchSummary(results) {
  const ok = results.filter(r => r.ok);
  const fail = results.filter(r => !r.ok);
  const titleCls = fail.length === 0 ? '' : (ok.length === 0 ? 'error' : 'partial');
  const titleText = fail.length === 0
    ? `✓ Filed ${ok.length} of ${results.length}`
    : (ok.length === 0
        ? `⚠ Filing failed — 0 of ${results.length}`
        : `Filed ${ok.length} of ${results.length} — ${fail.length} failed`);

  const rows = results.map(r => {
    if (r.ok) {
      return `<div class="bs-row ok">
        <span class="bs-mark">✓</span>
        <span class="bs-subj">${escapeHtml(r.subject)}</span>
        <a class="bs-link" href="${escapeHtml(r.url)}" target="_blank" rel="noopener">Open</a>
      </div>`;
    }
    return `<div class="bs-row fail">
      <span class="bs-mark">✕</span>
      <span class="bs-subj">${escapeHtml(r.subject)}</span>
      <span class="bs-err">${escapeHtml(r.error || 'unknown error')}</span>
    </div>`;
  }).join('');

  document.getElementById('status-area').innerHTML = `
    <div class="result-card batch-summary ${titleCls}">
      <div class="ttl">${titleText}</div>
      <div class="bs-list">${rows}</div>
    </div>`;

  // Add successful ones to the "recently filed" list
  for (const r of ok) {
    addToRecent({
      subject: r.subject,
      siteName: selectedSite.displayName || selectedSite.name,
      libName: selectedLib.name,
      filename: r.filename,
      url: r.url,
      date: new Date().toISOString()
    });
  }
}

// ============================================================
// UI helpers
// ============================================================
function showProgress(steps) {
  const html = '<div class="step-progress">' + steps.map(s =>
    `<div class="step ${s.status}" data-step="${s.id}">
      <span class="icon"><span class="state"></span></span>
      <span>${s.label}</span>
    </div>`
  ).join('') + '</div>';
  document.getElementById('status-area').innerHTML = html;
  updateStepIcons();
}
function setStep(id, status) {
  const el = document.querySelector(`[data-step="${id}"]`);
  if (!el) return;
  el.classList.remove('pending', 'active', 'done');
  el.classList.add(status);
  updateStepIcons();
}
function updateStepIcons() {
  document.querySelectorAll('.step-progress .step').forEach(s => {
    const icon = s.querySelector('.icon .state');
    if (s.classList.contains('done')) icon.innerHTML = '✓';
    else if (s.classList.contains('active')) icon.innerHTML = '<span class="spinner"></span>';
    else icon.innerHTML = '';
  });
}

function showResult(r) {
  if (r.success) {
    document.getElementById('status-area').innerHTML = `
      <div class="result-card">
        <div class="ttl">✓ Filed</div>
        <p><strong>${escapeHtml(r.filename)}</strong> saved to <strong>${escapeHtml(r.libName)}</strong> on ${escapeHtml(r.siteName)}.</p>
        <p><a href="${escapeHtml(r.url)}" target="_blank" rel="noopener">Open in SharePoint →</a></p>
      </div>`;
  } else {
    document.getElementById('status-area').innerHTML = `
      <div class="result-card error">
        <div class="ttl">⚠ Filing failed</div>
        <p>${escapeHtml(r.error)}</p>
      </div>`;
  }
}

function resetFiling() {
  document.getElementById('f-folder').value = '';
  document.getElementById('f-client').value = '';
  document.getElementById('f-project').value = '';
  document.getElementById('f-category').value = '';
  document.getElementById('f-notes').value = '';
  document.getElementById('status-area').innerHTML = '';
}

function toast(msg, type) {
  // For taskpane, just stuff into the status area briefly
  const cls = type === 'error' ? 'error' : '';
  document.getElementById('status-area').innerHTML = `<div class="result-card ${cls}"><p>${escapeHtml(msg)}</p></div>`;
  setTimeout(() => {
    if (document.getElementById('status-area').textContent.includes(msg)) {
      document.getElementById('status-area').innerHTML = '';
    }
  }, 3500);
}

function escapeHtml(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// Recently filed (local cache)
function loadRecent() {
  try {
    const raw = localStorage.getItem(RECENT_KEY);
    recentlyFiled = raw ? JSON.parse(raw) : [];
    renderRecent();
  } catch (e) { recentlyFiled = []; }
}
function addToRecent(item) {
  recentlyFiled.unshift(item);
  recentlyFiled = recentlyFiled.slice(0, 5);
  try { localStorage.setItem(RECENT_KEY, JSON.stringify(recentlyFiled)); } catch (e) {}
  renderRecent();
}
function renderRecent() {
  if (recentlyFiled.length === 0) {
    document.getElementById('recent-section').style.display = 'none';
    return;
  }
  document.getElementById('recent-section').style.display = 'block';
  document.getElementById('recent-list').innerHTML = recentlyFiled.map(r =>
    `<div class="recent-item">
      <span class="check">✓</span>
      <div class="info">
        <div class="s">${escapeHtml(r.subject)}</div>
        <div class="l">${escapeHtml(r.libName)} · ${escapeHtml(r.siteName)}</div>
      </div>
    </div>`
  ).join('');
}

// ============================================================
// SEARCH - global search across all filed emails in SharePoint
// ============================================================
// Uses Microsoft Graph's /search/query endpoint to do a tenant-wide
// full-text search of indexed SharePoint content. Results are filtered
// to .eml files (the format we file emails as) and hydrated with the
// metadata columns we wrote at filing time.
// ============================================================

let searchResults = [];
let activeFilters = { mine: false, thisMonth: false, hasAttachments: false };
let searchDebounceTimer = null;
let lastSearchQuery = '';

function switchTab(name) {
  document.querySelectorAll('.pane-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === name));
  document.querySelectorAll('.view-pane').forEach(p => p.classList.toggle('active', p.id === 'view-' + name));
  if (name === 'search') {
    // Focus the search box when the tab is opened
    setTimeout(() => document.getElementById('search-input').focus(), 50);
  }
}

function onSearchInput() {
  const v = document.getElementById('search-input').value;
  document.getElementById('search-clear').classList.toggle('show', v.length > 0);
  if (searchDebounceTimer) clearTimeout(searchDebounceTimer);
  // Debounce - wait 350ms after the user stops typing before firing the request
  searchDebounceTimer = setTimeout(() => doSearch(v.trim()), 350);
}

function onSearchKey(ev) {
  if (ev.key === 'Enter') {
    if (searchDebounceTimer) clearTimeout(searchDebounceTimer);
    doSearch(document.getElementById('search-input').value.trim());
  } else if (ev.key === 'Escape') {
    clearSearch();
  }
}

function clearSearch() {
  document.getElementById('search-input').value = '';
  document.getElementById('search-clear').classList.remove('show');
  searchResults = [];
  lastSearchQuery = '';
  renderEmptyState();
}

function cycleFilter(name) {
  activeFilters[name] = !activeFilters[name];
  document.getElementById('filter-' + name).classList.toggle('active', activeFilters[name]);
  // Re-render if we have results
  if (searchResults.length > 0 || lastSearchQuery) {
    renderSearchResults();
  }
}

async function doSearch(query) {
  if (!query || query.length < 2) {
    renderEmptyState();
    return;
  }
  if (query === lastSearchQuery) return;
  lastSearchQuery = query;

  document.getElementById('search-results').innerHTML = '<div class="search-loading"><span class="search-spinner"></span>Searching SharePoint...</div>';
  document.getElementById('search-meta').style.display = 'none';

  try {
    // Restrict results to .eml files (this is what the add-in produces when filing)
    // The 'fileExtension' KQL property limits scope. Could also add 'IsDocument:true'.
    const queryString = `(${query}) AND fileExtension:eml`;
    const body = {
      requests: [{
        entityTypes: ['driveItem'],
        query: { queryString: queryString },
        from: 0, size: 50,
        fields: ['id', 'name', 'webUrl', 'createdDateTime', 'lastModifiedDateTime', 'size', 'parentReference', 'fileSystemInfo']
      }]
    };
    const result = await graph('/search/query', {
      method: 'POST',
      body: JSON.stringify(body)
    });

    const hits = (((result.value || [])[0] || {}).hitsContainers || [])[0] || {};
    const items = (hits.hits || []).map(h => ({
      raw: h.resource,
      hitId: h.hitId,
      rank: h.rank,
      snippet: h.summary || ''
    }));

    // Hydrate each result with its SharePoint listItem fields (metadata columns)
    // We do this in parallel with a small concurrency limit to avoid throttling
    searchResults = await hydrateSearchResults(items);
    renderSearchResults();
  } catch (e) {
    console.error(e);
    document.getElementById('search-results').innerHTML =
      '<div class="search-empty"><div class="icon">⚠</div><div><strong>Search failed</strong></div><p>' + escapeHtml(e.message) + '</p></div>';
  }
}

async function hydrateSearchResults(items) {
  // Cap to first 25 to keep round trips reasonable - rest will hydrate lazily on expand
  const toHydrate = items.slice(0, 25);
  const concurrency = 5;
  const results = items.map(it => ({ ...it, fields: null, hydrated: false }));
  let idx = 0;
  async function worker() {
    while (true) {
      const myIdx = idx++;
      if (myIdx >= toHydrate.length) return;
      const it = toHydrate[myIdx];
      try {
        const driveId = it.raw.parentReference && it.raw.parentReference.driveId;
        const itemId = it.raw.id;
        if (driveId && itemId) {
          const li = await graph('/drives/' + driveId + '/items/' + itemId + '/listItem?$expand=fields');
          results[myIdx].fields = li.fields || {};
          results[myIdx].hydrated = true;
        }
      } catch (e) {
        // Ignore individual hydration failures - the card will just lack metadata
      }
    }
  }
  await Promise.all(Array.from({length: concurrency}, () => worker()));
  return results;
}

function renderEmptyState() {
  document.getElementById('search-meta').style.display = 'none';
  document.getElementById('search-results').innerHTML =
    '<div class="search-empty"><div class="icon">🔍</div><div><strong>Search filed emails</strong></div>' +
    '<p>Type a keyword above to search subject, body, sender, or filed metadata across all the SharePoint libraries you have access to.</p></div>';
}

function applyFilters(results) {
  let out = results;
  if (activeFilters.mine && currentAccount) {
    const me = (currentAccount.username || '').toLowerCase();
    out = out.filter(r => {
      const filedBy = (r.fields && (r.fields.FiledBy || '')).toLowerCase();
      return filedBy && filedBy.includes(me);
    });
  }
  if (activeFilters.thisMonth) {
    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    out = out.filter(r => {
      const created = r.raw.createdDateTime ? new Date(r.raw.createdDateTime) : null;
      return created && created >= startOfMonth;
    });
  }
  if (activeFilters.hasAttachments) {
    // Best-effort: emails with attachments are usually larger than ~50KB MIME
    out = out.filter(r => (r.raw.size || 0) > 50 * 1024);
  }
  return out;
}

function sortResults(results) {
  const sort = document.getElementById('search-sort').value;
  if (sort === 'date_desc') {
    return [...results].sort((a, b) => new Date(b.raw.createdDateTime || 0) - new Date(a.raw.createdDateTime || 0));
  }
  if (sort === 'date_asc') {
    return [...results].sort((a, b) => new Date(a.raw.createdDateTime || 0) - new Date(b.raw.createdDateTime || 0));
  }
  // relevance = natural order from Graph (already ranked)
  return results;
}

function renderSearchResults() {
  const filtered = applyFilters(searchResults);
  const sorted = sortResults(filtered);
  const meta = document.getElementById('search-meta');
  const container = document.getElementById('search-results');

  if (sorted.length === 0) {
    meta.style.display = 'none';
    container.innerHTML =
      '<div class="search-empty"><div class="icon">·</div><div><strong>No matches</strong></div>' +
      '<p>Try a different keyword or remove filters. Note that newly-filed emails can take a few minutes to appear in the index.</p></div>';
    return;
  }

  meta.style.display = 'flex';
  document.getElementById('search-count').textContent = sorted.length + (searchResults.length > sorted.length ? ' / ' + searchResults.length : '');

  const html = '<div class="result-card-wrap">' + sorted.map((r, i) => renderResultCard(r, i)).join('') + '</div>';
  container.innerHTML = html;
}

function renderResultCard(r, idx) {
  const f = r.fields || {};
  const filename = (r.raw.name || '').replace(/\.eml$/i, '');
  const subject = f.EmailSubject || filename || '(no subject)';
  const fromName = f.EmailFromName || '';
  const fromEmail = f.EmailFrom || '';
  const date = f.EmailDate || r.raw.createdDateTime;
  const client = f.FiledClient;
  const project = f.FiledProject;
  const category = f.FiledCategory;
  const filedBy = f.FiledBy;
  const libName = (r.raw.parentReference && r.raw.parentReference.driveType === 'documentLibrary') ? '' : '';
  // Render snippet: Graph includes a summary which we'll trust as the highlighted matching context
  const snippet = (r.snippet || '').replace(/<c0>/g, '<mark>').replace(/<\/c0>/g, '</mark>').replace(/<ddd\/>/g, '...');

  return `<div class="result-card" data-idx="${idx}" data-driveid="${escapeHtml(r.raw.parentReference?.driveId || '')}" data-itemid="${escapeHtml(r.raw.id || '')}" data-url="${escapeHtml(r.raw.webUrl || '')}" onclick="toggleResult(${idx})">
    <div class="r-top-row">
      <div class="r-subject">${escapeHtml(subject)}</div>
      <div class="r-date">${formatShortDate(date)}</div>
    </div>
    <div class="r-from-row">
      <span class="name">${escapeHtml(fromName || fromEmail || 'Unknown')}</span>
      ${fromEmail && fromName ? `<span class="email"> &lt;${escapeHtml(fromEmail)}&gt;</span>` : ''}
    </div>
    <div class="r-tags">
      ${client ? `<span class="r-tag">${escapeHtml(client)}</span>` : ''}
      ${project ? `<span class="r-tag">${escapeHtml(project)}</span>` : ''}
      ${category ? `<span class="r-tag">${escapeHtml(category)}</span>` : ''}
      ${filedBy ? `<span class="r-tag location">by ${escapeHtml(filedBy.split('@')[0])}</span>` : ''}
    </div>
    ${snippet ? `<div class="r-snippet">${snippet}</div>` : ''}
    <div class="r-preview" id="preview-${idx}">
      <div class="preview-body loading">Click to load preview...</div>
    </div>
  </div>`;
}

function formatShortDate(d) {
  if (!d) return '';
  const date = new Date(d);
  if (isNaN(date.getTime())) return '';
  const now = new Date();
  const sameYear = date.getFullYear() === now.getFullYear();
  const opts = sameYear ? { day: 'numeric', month: 'short' } : { day: 'numeric', month: 'short', year: 'numeric' };
  return date.toLocaleDateString('en-GB', opts);
}

async function toggleResult(idx) {
  const card = document.querySelector('.result-card[data-idx="' + idx + '"]');
  if (!card) return;
  const wasExpanded = card.classList.contains('expanded');
  // Collapse other cards
  document.querySelectorAll('.result-card.expanded').forEach(c => c.classList.remove('expanded'));
  if (wasExpanded) return;  // toggled off

  card.classList.add('expanded');
  const driveId = card.dataset.driveid;
  const itemId = card.dataset.itemid;
  const previewEl = card.querySelector('.r-preview');
  if (!driveId || !itemId) {
    previewEl.innerHTML = '<div class="preview-body">Cannot load preview - missing item reference.</div>';
    return;
  }
  // Render meta block from already-hydrated fields, then fetch + parse the .eml
  const result = searchResults[parseInt(card.dataset.idx, 10)];
  const f = result.fields || {};
  const metaHtml = `
    <div class="preview-meta">
      ${f.EmailSubject ? `<div class="meta-line"><strong>Subject:</strong> ${escapeHtml(f.EmailSubject)}</div>` : ''}
      ${(f.EmailFromName || f.EmailFrom) ? `<div class="meta-line"><strong>From:</strong> ${escapeHtml(f.EmailFromName || '')} ${f.EmailFrom ? '&lt;' + escapeHtml(f.EmailFrom) + '&gt;' : ''}</div>` : ''}
      ${f.EmailTo ? `<div class="meta-line"><strong>To:</strong> ${escapeHtml(f.EmailTo)}</div>` : ''}
      ${f.EmailDate ? `<div class="meta-line"><strong>Date:</strong> ${escapeHtml(new Date(f.EmailDate).toLocaleString())}</div>` : ''}
      ${f.FiledBy ? `<div class="meta-line"><strong>Filed by:</strong> ${escapeHtml(f.FiledBy)} ${f.FiledAt ? 'on ' + escapeHtml(new Date(f.FiledAt).toLocaleDateString()) : ''}</div>` : ''}
    </div>
    <div class="preview-body loading" id="preview-body-${idx}"><span class="search-spinner"></span>Loading email body...</div>
    <div class="preview-actions">
      <button class="primary" onclick="event.stopPropagation(); openInSharePoint('${escapeHtml(card.dataset.url)}')">Open in SharePoint</button>
      <button onclick="event.stopPropagation(); downloadEml('${driveId}', '${itemId}', '${escapeHtml(result.raw.name)}')">Download .eml</button>
    </div>
  `;
  previewEl.innerHTML = metaHtml;

  // Fetch + parse the .eml in the background
  try {
    const blob = await graph('/drives/' + driveId + '/items/' + itemId + '/content');
    // graph() returns the Response when content-type isn't JSON. Read it as blob.
    const realBlob = (blob instanceof Response) ? await blob.blob() : blob;
    const text = await realBlob.text();
    const parsed = parseEml(text);
    renderEmlPreview(idx, parsed);
  } catch (e) {
    console.error('Preview load failed', e);
    const bodyEl = document.getElementById('preview-body-' + idx);
    if (bodyEl) bodyEl.innerHTML = '<em>Could not load email body. Use "Open in SharePoint" to view the full content.</em>';
  }
}

function openInSharePoint(url) {
  if (!url) return;
  window.open(url, '_blank');
}

async function downloadEml(driveId, itemId, filename) {
  try {
    const blob = await graph('/drives/' + driveId + '/items/' + itemId + '/content');
    const realBlob = (blob instanceof Response) ? await blob.blob() : blob;
    const a = document.createElement('a');
    a.href = URL.createObjectURL(realBlob);
    a.download = filename || 'email.eml';
    document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 100);
  } catch (e) {
    toast('Download failed: ' + e.message, 'error');
  }
}

function renderEmlPreview(idx, parsed) {
  const bodyEl = document.getElementById('preview-body-' + idx);
  if (!bodyEl) return;
  bodyEl.classList.remove('loading');
  // Prefer plain text for safety; fall back to HTML if plain not available
  let content = '';
  let isHtml = false;
  if (parsed.text) {
    content = escapeHtml(parsed.text);
  } else if (parsed.html) {
    content = sanitiseHtmlForPreview(parsed.html);
    isHtml = true;
  } else {
    content = '<em>This email contains no readable body content. Use "Open in SharePoint" to view the original.</em>';
  }
  bodyEl.classList.toggle('html-body', isHtml);
  bodyEl.innerHTML = content;
  // Render attachments list if any
  if (parsed.attachments && parsed.attachments.length > 0) {
    const attHtml = '<div class="preview-attach"><strong>📎 Attachments:</strong> ' +
      parsed.attachments.map(a => `<span class="att-pill">${escapeHtml(a)}</span>`).join('') + '</div>';
    bodyEl.insertAdjacentHTML('afterend', attHtml);
  }
}

// ============================================================
// Minimal .eml (RFC822) parser - handles the 95% case
// ============================================================
// Returns { text, html, attachments[], headers{} }. Falls back gracefully
// if it hits something it can't handle - the result is best-effort, not
// guaranteed to be perfect for every edge case.
// ============================================================
function parseEml(raw) {
  if (!raw || typeof raw !== 'string') return { text: '', html: '', attachments: [], headers: {} };

  // Normalise line endings
  const src = raw.replace(/\r\n/g, '\n');
  const sep = src.indexOf('\n\n');
  if (sep < 0) return { text: src.slice(0, 5000), html: '', attachments: [], headers: {} };

  const headerText = src.slice(0, sep);
  const bodyText = src.slice(sep + 2);
  const headers = parseHeaders(headerText);

  const result = { text: '', html: '', attachments: [], headers: headers };
  const ct = (headers['content-type'] || 'text/plain').toLowerCase();

  if (ct.startsWith('multipart/')) {
    const boundary = extractBoundary(ct);
    if (boundary) {
      parseMultipart(bodyText, boundary, result);
    } else {
      result.text = bodyText.slice(0, 5000);
    }
  } else if (ct.startsWith('text/html')) {
    result.html = decodeBody(bodyText, headers['content-transfer-encoding'], getCharset(ct));
  } else if (ct.startsWith('text/')) {
    result.text = decodeBody(bodyText, headers['content-transfer-encoding'], getCharset(ct));
  } else {
    // unknown content type at root - try to show what we have
    result.text = bodyText.slice(0, 5000);
  }
  return result;
}

function parseHeaders(text) {
  // Unfold continuation lines (RFC 5322 - lines starting with whitespace are continuations)
  const unfolded = text.replace(/\n[ \t]/g, ' ');
  const lines = unfolded.split('\n');
  const headers = {};
  for (const line of lines) {
    const colonIdx = line.indexOf(':');
    if (colonIdx < 0) continue;
    const name = line.slice(0, colonIdx).trim().toLowerCase();
    const value = line.slice(colonIdx + 1).trim();
    headers[name] = value;
  }
  return headers;
}

function extractBoundary(ct) {
  const m = ct.match(/boundary\s*=\s*"?([^";\s]+)"?/i);
  return m ? m[1] : null;
}

function getCharset(ct) {
  const m = ct.match(/charset\s*=\s*"?([^";\s]+)"?/i);
  return m ? m[1].toLowerCase() : 'utf-8';
}

function parseMultipart(body, boundary, result) {
  const delim = '--' + boundary;
  // Split on the boundary - first chunk is preamble (ignored), last contains closing delim
  const parts = body.split(delim);
  for (const part of parts) {
    const trimmed = part.replace(/^\n/, '').replace(/^--$/, '').replace(/\n--\s*$/, '');
    if (!trimmed || trimmed.trim() === '' || trimmed.trim() === '--') continue;
    const headerEnd = trimmed.indexOf('\n\n');
    if (headerEnd < 0) continue;
    const partHeaderText = trimmed.slice(0, headerEnd);
    const partBody = trimmed.slice(headerEnd + 2);
    const partHeaders = parseHeaders(partHeaderText);
    const partCt = (partHeaders['content-type'] || 'text/plain').toLowerCase();
    const cd = (partHeaders['content-disposition'] || '').toLowerCase();

    if (cd.startsWith('attachment') || cd.includes('filename=')) {
      // Extract filename for the attachments list (we don't decode the binary content)
      const fnMatch = cd.match(/filename\s*=\s*"?([^";]+)"?/i) || partCt.match(/name\s*=\s*"?([^";]+)"?/i);
      if (fnMatch) result.attachments.push(fnMatch[1].trim());
      continue;
    }

    if (partCt.startsWith('multipart/')) {
      const inner = extractBoundary(partCt);
      if (inner) parseMultipart(partBody, inner, result);
    } else if (partCt.startsWith('text/html') && !result.html) {
      result.html = decodeBody(partBody, partHeaders['content-transfer-encoding'], getCharset(partCt));
    } else if (partCt.startsWith('text/plain') && !result.text) {
      result.text = decodeBody(partBody, partHeaders['content-transfer-encoding'], getCharset(partCt));
    }
  }
}

function decodeBody(body, encoding, charset) {
  encoding = (encoding || '').toLowerCase();
  if (encoding === 'base64') {
    try {
      const cleaned = body.replace(/\s+/g, '');
      const binary = atob(cleaned);
      // Best effort: assume UTF-8 if charset is utf-8 (most common); otherwise just return raw
      if (charset === 'utf-8' || charset === 'utf8') {
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        return new TextDecoder('utf-8').decode(bytes);
      }
      return binary;
    } catch (e) { return body; }
  }
  if (encoding === 'quoted-printable') {
    return body
      .replace(/=\r?\n/g, '')
      .replace(/=([0-9A-F]{2})/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
  }
  return body;
}

function sanitiseHtmlForPreview(html) {
  // Conservative HTML sanitisation:
  //  - Remove <script>, <iframe>, <object>, <embed>, <link>, <meta> tags entirely
  //  - Strip on* event handlers
  //  - Disallow javascript: URLs
  //  - Wrap in a div so styles don't leak
  let safe = html
    .replace(/<\?xml[^>]*>/gi, '')
    .replace(/<!DOCTYPE[^>]*>/gi, '')
    .replace(/<\/?html[^>]*>/gi, '')
    .replace(/<\/?head[^>]*>/gi, '')
    .replace(/<\/?body[^>]*>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<link\b[^>]*>/gi, '')
    .replace(/<meta\b[^>]*>/gi, '')
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, '')
    .replace(/<object[\s\S]*?<\/object>/gi, '')
    .replace(/<embed\b[^>]*>/gi, '')
    .replace(/\son\w+\s*=\s*"[^"]*"/gi, '')
    .replace(/\son\w+\s*=\s*'[^']*'/gi, '')
    .replace(/javascript:/gi, '')
    .replace(/<img([^>]*)\ssrc\s*=\s*"cid:[^"]*"/gi, '<img$1');  // strip inline cid: images we can't resolve
  return safe;
}
