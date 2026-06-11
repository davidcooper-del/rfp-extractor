/* ============================================================
   RFP Extractor v2 — Taskpane Logic
   • Hourly auto-timer with countdown ring
   • Microsoft Graph API → appends rows to SharePoint Excel
   • Duplicate detection (by OppName + DueDate)
   • Settings persisted in localStorage
   • All original EWS folder scan + Claude AI logic preserved
   ============================================================ */

'use strict';

// ── Constants ─────────────────────────────────────────────────
const TIMER_CIRCUMFERENCE = 2 * Math.PI * 22; // r=22 → 138.23
const DEFAULT_INTERVAL_MINS = 60;
const GRAPH_SCOPES = ['Files.ReadWrite', 'Sites.ReadWrite.All', 'offline_access'];

// ── State ─────────────────────────────────────────────────────
const state = {
  results: [],
  scanning: false,
  autoEnabled: false,
  timerInterval: null,       // setInterval handle for countdown tick
  scheduledTimeout: null,    // setTimeout handle for next scan trigger
  nextRunAt: null,           // Date of next scheduled run
  graphToken: null,          // Microsoft Graph access token
  settings: {
    apiKey: '',
    siteUrl: '',
    filePath: '',
    sheetName: 'Sheet1',
    intervalMins: DEFAULT_INTERVAL_MINS,
  },
};

// ── Office init ───────────────────────────────────────────────
Office.onReady(function (info) {
  log('info', `Office ready. Host: ${info.host}`);
  loadSettings();
  updateAuthUI();
  log('info', 'Configure settings then click Run Now or enable Auto-scan.');
});

// ── Tab switching ─────────────────────────────────────────────
function switchTab(name) {
  ['main', 'log', 'settings'].forEach(t => {
    document.getElementById('tab-' + t).classList.toggle('active', t === name);
    document.getElementById('panel-' + t).classList.toggle('active', t === name);
  });
}

// ── Settings ──────────────────────────────────────────────────
function loadSettings() {
  try {
    const saved = JSON.parse(localStorage.getItem('rfpExtractorSettings') || '{}');
    state.settings = { ...state.settings, ...saved };
    document.getElementById('settingApiKey').value = state.settings.apiKey || '';
    document.getElementById('settingSiteUrl').value = state.settings.siteUrl || '';
    document.getElementById('settingFilePath').value = state.settings.filePath || '';
    document.getElementById('settingSheetName').value = state.settings.sheetName || 'Sheet1';
    document.getElementById('settingInterval').value = state.settings.intervalMins || DEFAULT_INTERVAL_MINS;
    // Restore graph token if stored
    const tok = localStorage.getItem('rfpGraphToken');
    if (tok) state.graphToken = tok;
  } catch (e) {
    log('warn', 'Could not load saved settings: ' + e.message);
  }
}

function saveSettings() {
  state.settings.apiKey = document.getElementById('settingApiKey').value.trim();
  state.settings.siteUrl = document.getElementById('settingSiteUrl').value.trim().replace(/\/$/, '');
  state.settings.filePath = document.getElementById('settingFilePath').value.trim();
  state.settings.sheetName = document.getElementById('settingSheetName').value.trim() || 'Sheet1';
  state.settings.intervalMins = Math.max(10, parseInt(document.getElementById('settingInterval').value) || DEFAULT_INTERVAL_MINS);
  localStorage.setItem('rfpExtractorSettings', JSON.stringify(state.settings));
  const banner = document.getElementById('saveBanner');
  banner.style.display = 'block';
  setTimeout(() => { banner.style.display = 'none'; }, 2500);
  log('ok', 'Settings saved.');
  // Restart timer with new interval if running
  if (state.autoEnabled) {
    stopAutoTimer();
    startAutoTimer();
  }
}

function markDirty() {
  // nothing — just let user save explicitly
}

// ── Microsoft Graph Auth (MSAL popup flow via Office dialog) ──
async function authenticateGraph() {
  log('info', 'Opening Microsoft sign-in…');
  try {
    // Use Office dialog API to open auth popup
    const authUrl = buildGraphAuthUrl();
    Office.context.ui.displayDialogAsync(
      authUrl,
      { height: 60, width: 40, promptBeforeOpen: false },
      function (result) {
        if (result.status === Office.AsyncResultStatus.Failed) {
          log('err', 'Dialog failed: ' + result.error.message);
          return;
        }
        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (msg) {
          dialog.close();
          try {
            const data = JSON.parse(msg.message);
            if (data.access_token) {
              state.graphToken = data.access_token;
              localStorage.setItem('rfpGraphToken', data.access_token);
              log('ok', 'Microsoft Graph authenticated successfully.');
              updateAuthUI();
            } else if (data.error) {
              log('err', 'Auth error: ' + data.error_description);
            }
          } catch (e) {
            log('err', 'Could not parse auth response: ' + e.message);
          }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (evt) {
          if (evt.error === 12006) log('info', 'Auth dialog closed by user.');
        });
      }
    );
  } catch (e) {
    log('err', 'Auth init failed: ' + e.message);
  }
}

function buildGraphAuthUrl() {
  // This URL must be hosted on YOUR domain alongside the add-in files
  // See auth-callback.html included in this package
  const redirectUri = encodeURIComponent(window.location.origin + '/auth-callback.html');
  const scope = encodeURIComponent(GRAPH_SCOPES.join(' '));
  // Replace YOUR_TENANT_ID and YOUR_CLIENT_ID with your Azure App Registration values
  const tenantId = localStorage.getItem('rfpTenantId') || 'common';
  const clientId = localStorage.getItem('rfpClientId') || 'YOUR_CLIENT_ID';
  return `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize?` +
    `client_id=${clientId}&response_type=token&redirect_uri=${redirectUri}` +
    `&scope=${scope}&response_mode=fragment`;
}

function updateAuthUI() {
  const dot = document.getElementById('authDot');
  const txt = document.getElementById('authStatusText');
  if (state.graphToken) {
    dot.className = 'auth-dot ok';
    txt.textContent = 'Authenticated with Microsoft Graph — SharePoint sync enabled.';
  } else {
    dot.className = 'auth-dot';
    txt.textContent = 'Not authenticated — sign in to enable SharePoint sync.';
  }
}

// ── Auto-timer ─────────────────────────────────────────────────
function toggleAuto(enabled) {
  state.autoEnabled = enabled;
  document.getElementById('autoLabel').textContent = enabled ? 'Auto-scan on' : 'Auto-scan off';
  if (enabled) {
    log('info', `Auto-scan enabled. Runs every ${state.settings.intervalMins} min.`);
    startAutoTimer();
  } else {
    log('info', 'Auto-scan disabled.');
    stopAutoTimer();
    setTimerDisplay(0, '--:--');
    document.getElementById('nextRunTime').textContent = '—';
  }
}

function startAutoTimer() {
  const intervalMs = state.settings.intervalMins * 60 * 1000;
  state.nextRunAt = new Date(Date.now() + intervalMs);
  document.getElementById('nextRunTime').textContent = formatTime(state.nextRunAt);

  // Schedule the actual scan
  state.scheduledTimeout = setTimeout(async function scheduledRun() {
    if (!state.autoEnabled) return;
    log('info', 'Auto-scan triggered.');
    await runScan();
    // Schedule next run
    if (state.autoEnabled) {
      state.nextRunAt = new Date(Date.now() + intervalMs);
      document.getElementById('nextRunTime').textContent = formatTime(state.nextRunAt);
      state.scheduledTimeout = setTimeout(scheduledRun, intervalMs);
    }
  }, intervalMs);

  // Tick the countdown every second
  state.timerInterval = setInterval(() => {
    if (!state.nextRunAt) return;
    const remaining = Math.max(0, state.nextRunAt - Date.now());
    const totalMs = state.settings.intervalMins * 60 * 1000;
    const pct = 1 - remaining / totalMs;
    const mins = Math.floor(remaining / 60000);
    const secs = Math.floor((remaining % 60000) / 1000);
    const label = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
    setTimerDisplay(pct, label);
  }, 1000);
}

function stopAutoTimer() {
  if (state.timerInterval) { clearInterval(state.timerInterval); state.timerInterval = null; }
  if (state.scheduledTimeout) { clearTimeout(state.scheduledTimeout); state.scheduledTimeout = null; }
  state.nextRunAt = null;
}

function setTimerDisplay(pct, label) {
  const ring = document.getElementById('timerRing');
  const lbl = document.getElementById('timerLabel');
  ring.style.strokeDashoffset = TIMER_CIRCUMFERENCE * (1 - pct);
  lbl.textContent = label;
}

function formatTime(date) {
  return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: true });
}

// ── Main scan pipeline ─────────────────────────────────────────
async function runScan() {
  if (state.scanning) return;

  state.results = [];
  setBtnLoading(true);
  setStatus('Scanning…', 'warn');
  setProgress(5);
  document.getElementById('resultsSection').style.display = 'none';
  document.getElementById('emptyState').style.display = 'none';
  renderTable([]);
  log('info', '─── Scan started ───');

  try {
    const mailbox = Office.context.mailbox;

    // Step 1: Find SalesRFP folder
    log('info', 'Locating SalesRFP folder…');
    const folders = await getFolders(mailbox);
    const targetFolder = folders.find(f => f.displayName.toLowerCase() === 'salesrfp');
    if (!targetFolder) throw new Error('Folder "SalesRFP" not found in mailbox.');
    log('ok', `Found: "${targetFolder.displayName}"`);
    setProgress(15);

    // Step 2: Get emails
    log('info', 'Fetching emails from SalesRFP…');
    const emails = await getEmailsFromFolder(targetFolder.id, mailbox);
    if (!emails || emails.length === 0) {
      setStatus('No emails found', 'warn');
      log('warn', 'No emails in SalesRFP folder.');
      document.getElementById('emailCount').textContent = 0;
      document.getElementById('emptyState').style.display = 'block';
      document.getElementById('emptyState').innerHTML = '<div class="empty-icon">📭</div>No emails found in <strong>SalesRFP</strong>.';
      setBtnLoading(false);
      return;
    }
    document.getElementById('emailCount').textContent = emails.length;
    log('ok', `${emails.length} email(s) found. Extracting with Claude AI…`);
    setProgress(20);

    // Step 3: Extract via Claude AI
    const total = emails.length;
    let done = 0;
    for (const email of emails) {
      try {
        log('info', `AI extracting: "${truncate(email.subject, 45)}"`);
        const result = await extractRFPDetails(email);
        result.subject = email.subject;
        state.results.push(result);
      } catch (e) {
        log('warn', `Extract failed: ${e.message}`);
        state.results.push({
          OppName: truncate(email.subject, 50) || 'N/A',
          State: 'N/A', OwnerOrganization: 'N/A', DueDate: 'N/A',
          subject: email.subject, error: e.message,
        });
      }
      done++;
      setProgress(20 + Math.round((done / total) * 50));
      renderTable(state.results);
    }

    log('ok', `AI extraction done. ${state.results.length} records.`);
    setProgress(75);

    // Step 4: Push new rows to SharePoint Excel
    let newCount = 0;
    if (state.graphToken && state.settings.siteUrl && state.settings.filePath) {
      log('info', 'Connecting to SharePoint Excel…');
      try {
        newCount = await appendToSharePointExcel(state.results);
        log('ok', `${newCount} new row(s) appended to SharePoint Excel.`);
      } catch (e) {
        log('err', `SharePoint sync failed: ${e.message}`);
      }
    } else {
      log('warn', 'SharePoint sync skipped — configure settings and sign in first.');
    }

    setProgress(100);
    const now = new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    document.getElementById('lastSync').textContent = now;
    document.getElementById('newCount').textContent = newCount;
    setStatus(`Done — ${newCount} new rows added`, 'ok');
    log('ok', `─── Scan complete at ${now} ───`);
    document.getElementById('resultsSection').style.display = 'flex';
    document.getElementById('resultsSection').style.flexDirection = 'column';

  } catch (err) {
    setStatus('Error', 'err');
    log('err', err.message);
    setProgress(0);
  } finally {
    setBtnLoading(false);
  }
}

// ── Graph API: append rows to SharePoint Excel ────────────────
async function appendToSharePointExcel(records) {
  const { siteUrl, filePath, sheetName } = state.settings;
  const token = state.graphToken;

  // Build the Graph drive item URL for the file
  // Using /sites/{host}:/{site-path}:/drive/root:/{file-path}:/workbook
  const host = new URL(siteUrl).host;
  const sitePath = new URL(siteUrl).pathname;  // e.g. /sites/RADGov
  const encodedFile = filePath.replace(/^\//, ''); // remove leading slash

  const workbookBase =
    `https://graph.microsoft.com/v1.0/sites/${host}:${sitePath}:/drive/root:/${encodedFile}:/workbook`;

  const headers = {
    'Authorization': `Bearer ${token}`,
    'Content-Type': 'application/json',
  };

  // 1. Read existing rows to detect duplicates
  log('info', 'Reading existing rows from Excel for duplicate check…');
  let existingKeys = new Set();
  try {
    const usedRangeRes = await fetch(
      `${workbookBase}/worksheets/${encodeURIComponent(sheetName)}/usedRange?$select=values`,
      { headers }
    );
    if (usedRangeRes.ok) {
      const usedData = await usedRangeRes.json();
      const rows = usedData.values || [];
      // Skip header row; key = OppName+DueDate (cols 0 and 3 assumed)
      for (let i = 1; i < rows.length; i++) {
        const key = `${(rows[i][0] || '').toString().trim().toLowerCase()}||${(rows[i][3] || '').toString().trim()}`;
        existingKeys.add(key);
      }
      log('info', `${existingKeys.size} existing row(s) found for duplicate check.`);
    } else {
      log('warn', 'Could not read existing rows — will append all without dedup.');
    }
  } catch (e) {
    log('warn', 'Dedup read failed: ' + e.message);
  }

  // 2. Filter out duplicates and errors
  const newRows = records.filter(r => {
    if (r.error) return false;
    const key = `${(r.OppName || '').trim().toLowerCase()}||${(r.DueDate || '').trim()}`;
    return !existingKeys.has(key);
  });

  // Mark results with dup/new status for UI
  records.forEach(r => {
    if (r.error) { r._status = 'err'; return; }
    const key = `${(r.OppName || '').trim().toLowerCase()}||${(r.DueDate || '').trim()}`;
    r._status = existingKeys.has(key) ? 'dup' : 'new';
  });
  renderTable(records);

  if (newRows.length === 0) {
    log('info', 'No new rows to append — all already exist in Excel.');
    return 0;
  }

  // 3. Append new rows
  log('info', `Appending ${newRows.length} new row(s)…`);
  const values = newRows.map(r => [
    r.OppName || 'N/A',
    r.State || 'N/A',
    r.OwnerOrganization || 'N/A',
    r.DueDate || 'N/A',
  ]);

  // Get used range to find next empty row
  const tableRes = await fetch(
    `${workbookBase}/worksheets/${encodeURIComponent(sheetName)}/usedRange?$select=rowCount`,
    { headers }
  );

  let startRow = 1; // 0-indexed; row 0 is header
  if (tableRes.ok) {
    const tableData = await tableRes.json();
    startRow = tableData.rowCount || 1;
  }

  // Write rows using range address e.g. A5:D7
  const endRow = startRow + newRows.length - 1;
  const rangeAddress = `A${startRow + 1}:D${endRow + 1}`; // 1-indexed for Excel address

  const patchRes = await fetch(
    `${workbookBase}/worksheets/${encodeURIComponent(sheetName)}/range(address='${rangeAddress}')`,
    {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ values }),
    }
  );

  if (!patchRes.ok) {
    const errBody = await patchRes.text();
    throw new Error(`Graph PATCH failed ${patchRes.status}: ${errBody.slice(0, 200)}`);
  }

  return newRows.length;
}

// ── EWS: Get folders ──────────────────────────────────────────
function getFolders(mailbox) {
  return new Promise((resolve, reject) => {
    mailbox.makeEwsRequestAsync(buildFindFoldersEWS(), function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error('EWS request failed: ' + result.error.message));
        return;
      }
      try { resolve(parseFoldersFromEWSResponse(result.value)); }
      catch (e) { reject(e); }
    });
  });
}

function buildFindFoldersEWS() {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>
  <soap:Body>
    <m:FindFolder Traversal="Shallow">
      <m:FolderShape><t:BaseShape>AllProperties</t:BaseShape></m:FolderShape>
      <m:ParentFolderIds><t:DistinguishedFolderId Id="msgfolderroot"/></m:ParentFolderIds>
    </m:FindFolder>
  </soap:Body>
</soap:Envelope>`;
}

function parseFoldersFromEWSResponse(xmlString) {
  const NS = 'http://schemas.microsoft.com/exchange/services/2006/types';
  const doc = new DOMParser().parseFromString(xmlString, 'text/xml');
  const folders = [];
  for (const node of doc.getElementsByTagNameNS(NS, 'Folder')) {
    const id = node.getElementsByTagNameNS(NS, 'FolderId')[0];
    const name = node.getElementsByTagNameNS(NS, 'DisplayName')[0];
    if (id && name) folders.push({ id: id.getAttribute('Id'), displayName: name.textContent });
  }
  return folders;
}

// ── EWS: Get emails from folder ───────────────────────────────
function getEmailsFromFolder(folderId, mailbox) {
  return new Promise((resolve, reject) => {
    mailbox.makeEwsRequestAsync(buildFindItemsEWS(folderId), function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error('EWS FindItem failed: ' + result.error.message));
        return;
      }
      try { resolve(parseEmailsFromEWSResponse(result.value)); }
      catch (e) { reject(e); }
    });
  });
}

function buildFindItemsEWS(folderId) {
  return `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header><t:RequestServerVersion Version="Exchange2013"/></soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>AllProperties</t:BaseShape>
        <t:AdditionalProperties><t:FieldURI FieldURI="item:Body"/></t:AdditionalProperties>
      </m:ItemShape>
      <m:IndexedPageItemView MaxEntriesReturned="100" Offset="0" BasePoint="Beginning"/>
      <m:ParentFolderIds><t:FolderId Id="${folderId}"/></m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;
}

function parseEmailsFromEWSResponse(xmlString) {
  const NS = 'http://schemas.microsoft.com/exchange/services/2006/types';
  const doc = new DOMParser().parseFromString(xmlString, 'text/xml');
  return Array.from(doc.getElementsByTagNameNS(NS, 'Message')).map(node => ({
    subject: (node.getElementsByTagNameNS(NS, 'Subject')[0] || {}).textContent || '(no subject)',
    body: (node.getElementsByTagNameNS(NS, 'Body')[0] || {}).textContent || '',
  }));
}

// ── Claude AI extraction ──────────────────────────────────────
async function extractRFPDetails(email) {
  const apiKey = state.settings.apiKey;
  if (!apiKey) throw new Error('No API key set — go to Settings.');

  const prompt = `You are an RFP data extraction assistant for a government staffing company.

Extract the following details from this email:

OppName - clean RFP name (remove Lead, QA, ADDEN, state codes from the name)
State - 2-letter US state code (like CA, TX, NY)
OwnerOrganization - the company or government agency issuing the RFP
DueDate - final submission due date in YYYY-MM-DD format

Rules:
- For DueDate: Look for "Due Date", "Deadline", "Proposal Due", "Submission Deadline". Pick the FINAL/LATEST date if multiple appear.
- For State: Extract from subject line pattern like "Lead - CA -" or "TX RFP" or state name in body.
- For OppName: Remove "Lead", "QA", "ADDEN", state codes, and extra punctuation. Keep it short and clean.
- For OwnerOrganization: Look for the issuing agency/department/company name.
- If any field is missing, return N/A for that field only.

EMAIL SUBJECT: ${email.subject}

EMAIL BODY:
${email.body.slice(0, 4000)}

Return ONLY valid JSON (no markdown, no explanation):
{"OppName":"...","State":"...","OwnerOrganization":"...","DueDate":"..."}`;

  const response = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 300,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Claude API ${response.status}: ${errText.slice(0, 120)}`);
  }

  const data = await response.json();
  const raw = data.content.filter(b => b.type === 'text').map(b => b.text).join('');
  const parsed = JSON.parse(raw.replace(/```json|```/g, '').trim());

  return {
    OppName: parsed.OppName || 'N/A',
    State: parsed.State || 'N/A',
    OwnerOrganization: parsed.OwnerOrganization || 'N/A',
    DueDate: parsed.DueDate || 'N/A',
  };
}

// ── Render results table ──────────────────────────────────────
function renderTable(rows) {
  document.getElementById('resultsBadge').textContent = rows.length;
  const tbody = document.getElementById('resultsBody');
  tbody.innerHTML = '';
  for (const row of rows) {
    const tr = document.createElement('tr');
    const statusCell = document.createElement('td');
    if (row._status === 'new') { statusCell.className = 'td-new'; statusCell.textContent = 'NEW'; }
    else if (row._status === 'dup') { statusCell.className = 'td-dup'; statusCell.textContent = 'DUP'; }
    else if (row.error) { statusCell.className = 'td-err'; statusCell.title = row.error; statusCell.textContent = 'ERR'; }
    else { statusCell.className = 'td-na'; statusCell.textContent = '…'; }

    const tdName = document.createElement('td'); tdName.className = 'td-name'; tdName.title = row.OppName; tdName.textContent = row.OppName || 'N/A';
    const tdState = document.createElement('td'); tdState.className = row.State !== 'N/A' ? 'td-state' : 'td-na'; tdState.textContent = row.State || 'N/A';
    const tdOrg = document.createElement('td'); tdOrg.style.fontSize = '10px'; tdOrg.title = row.OwnerOrganization; tdOrg.textContent = truncate(row.OwnerOrganization, 22);
    const tdDate = document.createElement('td'); tdDate.className = row.DueDate !== 'N/A' ? 'td-date' : 'td-na'; tdDate.textContent = row.DueDate || 'N/A';

    tr.appendChild(tdName); tr.appendChild(tdState); tr.appendChild(tdOrg); tr.appendChild(tdDate); tr.appendChild(statusCell);
    tbody.appendChild(tr);
  }
}

// ── UI helpers ────────────────────────────────────────────────
function setStatus(text, cls = 'neutral') {
  const el = document.getElementById('statusText');
  el.textContent = text;
  el.className = `status-val ${cls}`;
}
function setProgress(pct) {
  document.getElementById('progressBar').style.width = `${pct}%`;
}
function setBtnLoading(loading) {
  const btn = document.getElementById('btnScan');
  btn.disabled = loading;
  btn.classList.toggle('loading', loading);
  state.scanning = loading;
}

// ── Logging ───────────────────────────────────────────────────
function log(level, msg) {
  const box = document.getElementById('logBox');
  const line = document.createElement('div');
  const ts = new Date().toLocaleTimeString('en-US', { hour12: false });
  line.className = `log-${level}`;
  line.textContent = `[${ts}] ${msg}`;
  box.appendChild(line);
  box.scrollTop = box.scrollHeight;
}
function clearLog() {
  document.getElementById('logBox').innerHTML = '<span class="log-info">[cleared]</span>';
}

// ── Clear results ─────────────────────────────────────────────
function clearResults() {
  state.results = [];
  renderTable([]);
  document.getElementById('resultsSection').style.display = 'none';
  document.getElementById('emptyState').style.display = 'block';
  document.getElementById('emptyState').innerHTML =
    '<div class="empty-icon">📬</div>Enable <strong>Auto-scan</strong> or click <strong>Run Now</strong> to scan SalesRFP and sync new rows to SharePoint.';
  setStatus('Ready', 'neutral');
  setProgress(0);
  document.getElementById('emailCount').textContent = '—';
  document.getElementById('newCount').textContent = '—';
}

// ── Utility ───────────────────────────────────────────────────
function truncate(str, len) {
  if (!str) return '';
  return str.length > len ? str.slice(0, len) + '…' : str;
}
