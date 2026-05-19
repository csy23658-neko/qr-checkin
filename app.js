'use strict';

// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  accessToken:  null,
  user:         null,   // { email, name }
  role:         null,   // 'admin' | 'volunteer'
  sheetId:      null,
  rows:         [],     // [{ rowIndex, id, name, address, notes, checkinStamp, checkinBy }]
  selectedRow:  null,
  scanner:      null,
  scanCooldown: false,
  clockTimer:   null,
  currentScreen: 'screen-login',
};

// ── Boot ──────────────────────────────────────────────────────────────────────

window.addEventListener('load', () => {
  const poll = setInterval(() => {
    if (typeof google !== 'undefined' && google.accounts) {
      clearInterval(poll);
      initAuth();
    }
  }, 100);
});

// ── Auth ──────────────────────────────────────────────────────────────────────

let tokenClient;

function initAuth() {
  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id:      CONFIG.GOOGLE_CLIENT_ID,
    scope:          'openid email profile https://www.googleapis.com/auth/spreadsheets',
    callback:       onTokenReceived,
    error_callback: () => showLoginError('登入失敗，請重試。'),
  });
}

function signIn() {
  tokenClient.requestAccessToken({ prompt: 'select_account' });
}

async function onTokenReceived(resp) {
  if (resp.error) { showLoginError('登入失敗：' + resp.error); return; }

  state.accessToken = resp.access_token;

  // Silent refresh 2 min before expiry
  setTimeout(() => tokenClient.requestAccessToken({ prompt: '' }),
    (resp.expires_in - 120) * 1000);

  // Get user info from Google
  let info;
  try {
    const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo',
      { headers: { Authorization: `Bearer ${state.accessToken}` } });
    info = await r.json();
  } catch { showLoginError('無法取得使用者資訊，請重試。'); return; }

  const email = (info.email || '').toLowerCase();
  const admins     = CONFIG.ADMINS.map(e => e.toLowerCase());
  const volunteers = CONFIG.VOLUNTEERS.map(e => e.toLowerCase());

  if (admins.includes(email)) {
    state.role = 'admin';
  } else if (volunteers.length === 0 || volunteers.includes(email)) {
    state.role = 'volunteer';
  } else {
    showLoginError('您的帳號（' + email + '）沒有使用此系統的權限。');
    return;
  }

  state.user = { email, name: info.name || email };

  // Show report nav tab for admins
  if (state.role === 'admin') {
    document.getElementById('nav-report').style.display = '';
  }

  // Determine starting screen
  const sid = new URLSearchParams(location.search).get('sheet');
  if (sid) {
    state.sheetId = sid;
    enterApp();
  } else if (state.role === 'admin') {
    showScreen('screen-setup');
  } else {
    showLoginError('請向管理員索取活動連結。');
  }
}

// ── Event Setup ───────────────────────────────────────────────────────────────

async function confirmSetup() {
  hideError('setup-error');
  const url = document.getElementById('input-sheet-url').value.trim();
  const m   = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!m) { showError('setup-error', '網址格式不正確，請重新貼上 Google 試算表網址。'); return; }

  state.sheetId = m[1];
  history.replaceState(null, '', '?sheet=' + state.sheetId);

  const shareUrl = location.href;
  document.getElementById('share-url').textContent = shareUrl;
  document.getElementById('share-box').style.display = '';

  await enterApp();
}

function copyUrl() {
  navigator.clipboard.writeText(location.href).then(() => {
    const b = document.getElementById('btn-copy');
    b.textContent = '✓ 已複製';
    setTimeout(() => { b.textContent = '複製網址'; }, 2000);
  });
}

function goToSetup() {
  // Stop scanner if needed
  if (state.currentScreen === 'screen-scan') stopScanner();
  showScreen('screen-setup');
}

// ── Enter App ─────────────────────────────────────────────────────────────────

async function enterApp() {
  setLoading(true);
  try {
    await loadRows();
    document.getElementById('app-nav').style.display = 'flex';
    goTo('screen-scan');
  } catch (e) {
    if (e.status === 403 || e.status === 404) {
      showLoginError('無法存取試算表。\n請確認您已將試算表共用給您的 Google 帳號，且擁有編輯權限。');
    } else {
      showLoginError('載入資料失敗：' + (e.message || e));
    }
  } finally {
    setLoading(false);
  }
}

// ── Sheets API ────────────────────────────────────────────────────────────────

async function apiGet(range) {
  const r = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${state.sheetId}/values/${encodeURIComponent(range)}`,
    { headers: { Authorization: `Bearer ${state.accessToken}` } }
  );
  if (!r.ok) {
    const j  = await r.json().catch(() => ({}));
    const e  = new Error(j.error?.message || `HTTP ${r.status}`);
    e.status = r.status;
    throw e;
  }
  return r.json();
}

async function apiPut(range, values) {
  const r = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${state.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${state.accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({ range, majorDimension: 'ROWS', values }),
    }
  );
  if (!r.ok) {
    const j  = await r.json().catch(() => ({}));
    throw new Error(j.error?.message || `HTTP ${r.status}`);
  }
}

// ── Data ──────────────────────────────────────────────────────────────────────

async function loadRows() {
  const result = await apiGet('A:F');
  const all    = result.values || [];
  // Row 1 is headers; data starts at row 2
  state.rows = all.slice(1)
    .map((r, i) => ({
      rowIndex:     i + 2,
      id:           (r[0] || '').trim(),
      name:          r[1] || '',
      address:       r[2] || '',
      notes:         r[3] || '',
      checkinStamp:  r[4] || '',
      checkinBy:     r[5] || '',
    }))
    .filter(r => r.id);
}

async function refreshData() {
  setLoading(true);
  try {
    await loadRows();
    if (state.currentScreen === 'screen-list')   renderList();
    if (state.currentScreen === 'screen-report') renderReport();
  } catch { alert('重新整理失敗，請檢查網路。'); }
  finally  { setLoading(false); }
}

// ── Check-in Logic ────────────────────────────────────────────────────────────

async function checkin(row, manual = false) {
  const stamp = nowString();
  const by    = state.user.name + (manual ? ' 手動' : '');
  await apiPut(`E${row.rowIndex}:F${row.rowIndex}`, [[stamp, by]]);
  row.checkinStamp = stamp;
  row.checkinBy    = by;
}

async function undoCheckin(row) {
  await apiPut(`E${row.rowIndex}:F${row.rowIndex}`, [['', '']]);
  row.checkinStamp = '';
  row.checkinBy    = '';
}

// ── Scan Screen ───────────────────────────────────────────────────────────────

async function startScanner() {
  updateClock();
  state.clockTimer = setInterval(updateClock, 1000);

  const container = document.getElementById('qr-reader');
  container.innerHTML = '';
  state.scanner = new Html5Qrcode('qr-reader');

  try {
    await state.scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 230, height: 230 } },
      onScanSuccess,
      () => {}  // per-frame decode errors are normal
    );
  } catch (err) {
    container.innerHTML = `<p class="msg-error" style="padding:20px">相機啟動失敗：${esc(String(err))}<br><br>請確認已允許瀏覽器使用相機。</p>`;
    state.scanner = null;
  }
}

async function stopScanner() {
  clearInterval(state.clockTimer);
  state.clockTimer = null;
  if (state.scanner) {
    try { await state.scanner.stop(); } catch {}
    state.scanner = null;
  }
}

async function onScanSuccess(rawCode) {
  if (state.scanCooldown) return;
  state.scanCooldown = true;

  const code = rawCode.trim();
  const row  = state.rows.find(r => r.id === code);

  if (!row) {
    flashResult('not-found', '未找到', code, null);
  } else if (row.checkinStamp) {
    flashResult('already', '已掃過', code, row);
  } else {
    try {
      await checkin(row);
      flashResult('success', '報到成功', code, row);
    } catch {
      flashResult('err', '寫入失敗', code, null);
    }
  }

  setTimeout(() => {
    state.scanCooldown = false;
    resetBanner();
  }, 3000);
}

function flashResult(type, text, code, row) {
  const banner = document.getElementById('result-banner');
  banner.className = 'result-banner result-' + type;
  document.getElementById('result-main').textContent   = text;
  document.getElementById('result-detail').textContent = row ? (row.id + '  ' + row.name) : code;

  const info = document.getElementById('scan-info');
  if (row) {
    document.getElementById('scan-info-id').textContent      = row.id;
    document.getElementById('scan-info-name').textContent    = row.name;
    document.getElementById('scan-info-address').textContent = row.address;

    const notesEl = document.getElementById('scan-info-notes');
    notesEl.textContent    = row.notes;
    notesEl.style.display  = row.notes ? '' : 'none';

    const stampEl = document.getElementById('scan-info-stamp');
    stampEl.textContent   = row.checkinStamp ? '報到時間：' + row.checkinStamp : '';
    stampEl.style.display = row.checkinStamp ? '' : 'none';

    info.style.display = '';
  } else {
    info.style.display = 'none';
  }
}

function resetBanner() {
  const banner = document.getElementById('result-banner');
  banner.className = 'result-banner result-idle';
  document.getElementById('result-main').textContent   = '準備掃描';
  document.getElementById('result-detail').textContent = '';
  document.getElementById('scan-info').style.display   = 'none';
}

function updateClock() {
  const el = document.getElementById('scan-clock');
  if (el) el.textContent = nowString();
}

// ── List Screen ───────────────────────────────────────────────────────────────

function renderList() {
  const q    = (document.getElementById('search-input')?.value || '').trim().toLowerCase();
  let   rows = state.rows;

  if (q) {
    rows = rows.filter(r =>
      r.name.toLowerCase().includes(q)    ||
      r.id.toLowerCase().endsWith(q)      ||
      r.address.toLowerCase().includes(q) ||
      r.checkinStamp.startsWith(q)
    );
  }

  const total   = rows.length;
  const checked = rows.filter(r => r.checkinStamp).length;
  document.getElementById('list-stats').textContent =
    `已報到：${checked}　未報到：${total - checked}　共 ${total} 人`;

  const box = document.getElementById('list-container');
  if (total === 0) {
    box.innerHTML = '<p class="msg-empty">無符合結果</p>';
    return;
  }

  box.innerHTML = rows.map(r => `
    <div class="list-row${r.checkinStamp ? ' checked-in' : ''}"
         onclick="openManual(${r.rowIndex})">
      <div class="list-row-main">
        <span class="list-id">${esc(r.id)}</span>
        <span class="list-name">${esc(r.name)}</span>
        ${r.notes ? `<span class="list-notes">${esc(r.notes)}</span>` : ''}
      </div>
      <div class="list-row-sub">
        <span class="list-address">${esc(r.address)}</span>
      </div>
      ${r.checkinStamp
        ? `<div class="list-stamp">${esc(r.checkinStamp)}<br>${esc(r.checkinBy)}</div>`
        : ''}
    </div>
  `).join('');
}

// ── Manual Check-in Screen ────────────────────────────────────────────────────

function openManual(rowIndex) {
  const row = state.rows.find(r => r.rowIndex === rowIndex);
  if (!row) return;
  state.selectedRow = row;
  renderManual();
  goTo('screen-manual');
}

function renderManual() {
  const r = state.selectedRow;

  document.getElementById('manual-id').textContent      = r.id;
  document.getElementById('manual-name').textContent    = r.name;
  document.getElementById('manual-address').textContent = r.address;

  const notesEl = document.getElementById('manual-notes');
  notesEl.textContent   = r.notes;
  notesEl.style.display = r.notes ? '' : 'none';

  const stampEl = document.getElementById('manual-stamp');
  if (r.checkinStamp) {
    stampEl.textContent = `報到時間：${r.checkinStamp}\n報到人員：${r.checkinBy}`;
  } else {
    stampEl.textContent = '尚未報到';
  }

  document.getElementById('btn-manual-in').disabled   = !!r.checkinStamp;
  document.getElementById('btn-manual-undo').disabled = !r.checkinStamp;
}

async function manualCheckin() {
  setLoading(true);
  try {
    await checkin(state.selectedRow, true);
    renderManual();
    renderList();
  } catch { alert('操作失敗，請檢查網路後重試。'); }
  finally  { setLoading(false); }
}

async function manualUndo() {
  setLoading(true);
  try {
    await undoCheckin(state.selectedRow);
    renderManual();
    renderList();
  } catch { alert('操作失敗，請檢查網路後重試。'); }
  finally  { setLoading(false); }
}

// ── Report Screen ─────────────────────────────────────────────────────────────

function renderReport() {
  const total   = state.rows.length;
  const checked = state.rows.filter(r => r.checkinStamp).length;

  document.getElementById('report-stats').innerHTML = `
    <div class="stat-card blue">
      <div class="stat-num">${total}</div>
      <div class="stat-label">總人數</div>
    </div>
    <div class="stat-card green">
      <div class="stat-num">${checked}</div>
      <div class="stat-label">已報到</div>
    </div>
    <div class="stat-card orange">
      <div class="stat-num">${total - checked}</div>
      <div class="stat-label">未報到</div>
    </div>
  `;
}

function downloadReport() {
  const header = ['標題', '姓名', '通訊地址', '備註', '報到時間', '報到人員'];
  const csvRows = [
    header,
    ...state.rows.map(r => [r.id, r.name, r.address, r.notes, r.checkinStamp, r.checkinBy]),
  ];
  const csv  = csvRows.map(row =>
    row.map(cell => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(',')
  ).join('\r\n');

  // BOM so Excel opens Chinese text correctly
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `出席報告_${dateString()}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ── Navigation ────────────────────────────────────────────────────────────────

function goTo(screenId) {
  // Tear down scan screen if we're leaving it
  if (state.currentScreen === 'screen-scan' && screenId !== 'screen-scan') {
    stopScanner();
  }

  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(screenId).classList.add('active');

  document.querySelectorAll('.nav-btn').forEach(b => {
    b.classList.toggle('active', b.dataset.screen === screenId);
  });

  state.currentScreen = screenId;

  if (screenId === 'screen-scan')   startScanner();
  if (screenId === 'screen-list')   renderList();
  if (screenId === 'screen-report') renderReport();
}

function showScreen(id) {
  // Used before nav bar appears (login → setup → app)
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  state.currentScreen = id;
}

function goBack() {
  goTo('screen-list');
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function nowString() {
  const d   = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${pad(d.getMonth()+1)}/${pad(d.getDate())} ` +
         `${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
}

function dateString() {
  const d   = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}`;
}

function esc(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function setLoading(on) {
  document.getElementById('loading-overlay').style.display = on ? 'flex' : 'none';
}

function showLoginError(msg) {
  setLoading(false);
  const el = document.getElementById('login-error');
  el.textContent   = msg;
  el.style.display = '';
  showScreen('screen-login');
}

function showError(id, msg) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent   = msg;
  el.style.display = '';
}

function hideError(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = 'none';
}
