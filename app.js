'use strict';

// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  accessToken:      null,
  user:             null,   // { email, name }
  role:             null,   // 'admin' | 'volunteer'
  sheetId:          null,
  spreadsheetTitle: null,
  allowListSheetId: CONFIG.ALLOW_LIST_SHEET_ID || null,
  rows:             [],     // [{ rowIndex, id, name, address, notes, checkinStamp, checkinBy }]
  selectedRow:      null,
  scanner:          null,
  scanCooldown:     false,
  clockTimer:       null,
  currentScreen:    'screen-login',
};

// ── Boot ──────────────────────────────────────────────────────────────────────

window.addEventListener('load', () => {
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
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
    scope:          'openid email profile https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file',
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

  // Get user info
  let info;
  try {
    const r = await fetch('https://www.googleapis.com/oauth2/v3/userinfo',
      { headers: { Authorization: `Bearer ${state.accessToken}` } });
    info = await r.json();
  } catch { showLoginError('無法取得使用者資訊，請重試。'); return; }

  const email  = (info.email || '').toLowerCase();
  const admins = CONFIG.ADMINS.map(e => e.toLowerCase());

  if (admins.includes(email)) {
    state.role = 'admin';
  } else {
    // Check allow list
    setLoading(true);
    let permitted = false;
    try {
      permitted = await checkAllowList(email);
    } catch {
      setLoading(false);
      showLoginError('無法驗證使用者權限，請稍後再試。');
      return;
    }
    setLoading(false);
    if (!permitted) {
      showLoginError(`您的帳號（${email}）沒有使用此系統的權限。`);
      return;
    }
    state.role = 'volunteer';
  }

  state.user = { email, name: info.name || email };

  // Show admin nav tab for admins
  if (state.role === 'admin') {
    document.getElementById('nav-admin').style.display = '';
  }
  document.getElementById('app-nav').style.display = 'flex';

  // Navigate to first screen
  const sid = new URLSearchParams(location.search).get('sheet');
  if (sid) {
    state.sheetId = sid;
    await enterApp();
  } else if (state.role === 'admin') {
    goTo('screen-admin');
  } else {
    // Volunteer without a sheet URL
    document.getElementById('app-nav').style.display = 'none';
    showLoginError('請向管理員索取活動連結。');
  }
}

async function checkAllowList(email) {
  if (!state.allowListSheetId) return false;
  const emails = await loadAllowListEmails();
  return emails.some(e => e.email === email);
}

// ── Event Setup ───────────────────────────────────────────────────────────────

const EVENT_HEADERS = ['標題', '姓名', '通訊地址', '備註', '報到時間', '報到人員'];

// One-click event creation: makes a new Google Sheet with the right columns,
// then shares it with everyone on the volunteer allow list.
async function createEvent() {
  hideError('create-error');
  const name = document.getElementById('input-event-name').value.trim();
  if (!name) { showError('create-error', '請輸入活動名稱。'); return; }

  setLoading(true);
  try {
    const created = await sheetsRequest('POST', '', {
      properties: { title: name },
      sheets: [{ properties: { title: '名單', gridProperties: { frozenRowCount: 1 } } }],
    });
    state.sheetId          = created.spreadsheetId;
    state.spreadsheetTitle = name;
    state.rows             = [];

    await apiPut('A1:F1', [EVENT_HEADERS]);

    // Bold header row (best-effort)
    try {
      const gid = created.sheets?.[0]?.properties?.sheetId ?? 0;
      await sheetsRequest('POST', `/${state.sheetId}:batchUpdate`, {
        requests: [{
          repeatCell: {
            range:  { sheetId: gid, startRowIndex: 0, endRowIndex: 1 },
            cell:   { userEnteredFormat: { textFormat: { bold: true } } },
            fields: 'userEnteredFormat.textFormat.bold',
          },
        }],
      });
    } catch {}

    document.getElementById('input-event-name').value = '';
    afterEventSelected();
    await shareWithVolunteers(true);
  } catch (err) {
    showError('create-error', '建立失敗：' + err.message);
  } finally {
    setLoading(false);
  }
}

// Advanced: load an existing spreadsheet by URL
async function confirmSetup() {
  hideError('setup-error');
  const url = document.getElementById('input-sheet-url').value.trim();
  const m   = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!m) { showError('setup-error', '網址格式不正確，請重新貼上 Google 試算表網址。'); return; }

  state.sheetId = m[1];

  setLoading(true);
  try {
    await loadRows();
    state.spreadsheetTitle = await fetchSpreadsheetTitle(state.sheetId);
    document.getElementById('input-sheet-url').value = '';
    afterEventSelected();
  } catch (e) {
    state.sheetId          = null;
    state.spreadsheetTitle = null;
    const msg = (e.status === 403 || e.status === 404)
      ? '無法存取試算表，請確認共用設定。'
      : '載入失敗：' + (e.message || e);
    showError('setup-error', msg);
  } finally {
    setLoading(false);
  }
}

function afterEventSelected() {
  history.replaceState(null, '', '?sheet=' + state.sheetId);
  saveRecentEvent(state.sheetId, state.spreadsheetTitle || '');
  renderAdmin();
  startAutoRefresh();
}

// ── Recent events (localStorage) ──────────────────────────────────────────────

function getRecentEvents() {
  try { return JSON.parse(localStorage.getItem('qr-recent-events')) || []; }
  catch { return []; }
}

function saveRecentEvent(id, title) {
  const list = getRecentEvents().filter(e => e.id !== id);
  list.unshift({ id, title, ts: Date.now() });
  try { localStorage.setItem('qr-recent-events', JSON.stringify(list.slice(0, 8))); } catch {}
}

function renderRecentEvents() {
  const el = document.getElementById('recent-events');
  if (!el) return;
  const list = getRecentEvents().filter(e => e.id !== state.sheetId);
  if (!list.length) { el.innerHTML = ''; return; }
  el.innerHTML = '<p class="label" style="margin-top:6px">或選擇最近的活動：</p>' +
    list.map(e => `
      <button class="recent-event-row" onclick="switchEvent('${e.id}')">
        <span class="recent-event-title">${esc(e.title || '（未命名活動）')}</span>
        <span class="recent-event-date">${new Date(e.ts).toLocaleDateString('zh-TW')}</span>
      </button>`).join('');
}

async function switchEvent(id) {
  setLoading(true);
  try {
    state.sheetId = id;
    await loadRows();
    state.spreadsheetTitle = await fetchSpreadsheetTitle(id);
    afterEventSelected();
  } catch (e) {
    state.sheetId = null;
    state.rows    = [];
    alert('無法載入此活動：' + (e.message || e));
    renderAdmin();
  } finally {
    setLoading(false);
  }
}

// ── Share event sheet with volunteers (Drive API) ─────────────────────────────

async function shareWithVolunteers(silent = false) {
  if (!state.sheetId) return;
  const msgEl = document.getElementById('share-vols-msg');
  if (msgEl) msgEl.style.display = 'none';

  setLoading(true);
  let ok = 0, fail = 0;
  try {
    const emails = await loadAllowListEmails();
    for (const e of emails) {
      try {
        const r = await fetch(
          `https://www.googleapis.com/drive/v3/files/${state.sheetId}/permissions?sendNotificationEmail=false`,
          {
            method:  'POST',
            headers: { Authorization: `Bearer ${state.accessToken}`, 'Content-Type': 'application/json' },
            body:    JSON.stringify({ role: 'writer', type: 'user', emailAddress: e.email }),
          });
        if (r.ok) ok++; else fail++;
      } catch { fail++; }
    }
    if (msgEl) {
      msgEl.textContent = fail === 0
        ? `✓ 已共用給 ${ok} 位志工`
        : `已共用 ${ok} 位，${fail} 位失敗。若此試算表不是由本系統建立，請直接在 Google 試算表按「共用」加入志工。`;
      msgEl.className     = fail === 0 ? 'msg-success' : 'msg-error';
      msgEl.style.display = '';
    }
  } catch (err) {
    if (!silent && msgEl) {
      msgEl.textContent   = '共用失敗：' + err.message;
      msgEl.className     = 'msg-error';
      msgEl.style.display = '';
    }
  } finally {
    setLoading(false);
  }
}

function copyUrl() {
  navigator.clipboard.writeText(location.href).then(() => {
    const b = document.getElementById('btn-copy');
    b.textContent = '✓ 已複製';
    setTimeout(() => { b.textContent = '複製網址'; }, 2000);
  });
}

// ── Enter App (sheet URL pre-set via URL param) ───────────────────────────────

async function enterApp() {
  setLoading(true);
  try {
    await loadRows();
    state.spreadsheetTitle = await fetchSpreadsheetTitle(state.sheetId);
    if (state.role === 'admin') saveRecentEvent(state.sheetId, state.spreadsheetTitle || '');
    startAutoRefresh();
    goTo('screen-scan');
  } catch (e) {
    document.getElementById('app-nav').style.display = 'none';
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

async function sheetsRequest(method, path, body) {
  const opts = {
    method,
    headers: { Authorization: `Bearer ${state.accessToken}` },
  };
  if (body !== undefined) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  }
  const r = await fetch('https://sheets.googleapis.com/v4/spreadsheets' + path, opts);
  if (!r.ok) {
    const j  = await r.json().catch(() => ({}));
    const e  = new Error(j.error?.message || `HTTP ${r.status}`);
    e.status = r.status;
    throw e;
  }
  return r.json();
}

// Event sheet helpers (use state.sheetId)
function apiGet(range) {
  return sheetsRequest('GET',
    `/${state.sheetId}/values/${encodeURIComponent(range)}`);
}
function apiPut(range, values) {
  return sheetsRequest('PUT',
    `/${state.sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`,
    { range, majorDimension: 'ROWS', values });
}
function apiAppend(range, values) {
  return sheetsRequest('POST',
    `/${state.sheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    { majorDimension: 'ROWS', values });
}

// Generic sheet helpers (any sheetId)
function sheetsGet(sheetId, range) {
  return sheetsRequest('GET',
    `/${sheetId}/values/${encodeURIComponent(range)}`);
}
function sheetsPut(sheetId, range, values) {
  return sheetsRequest('PUT',
    `/${sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`,
    { range, majorDimension: 'ROWS', values });
}
function sheetsAppend(sheetId, range, values) {
  return sheetsRequest('POST',
    `/${sheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`,
    { majorDimension: 'ROWS', values });
}

async function fetchSpreadsheetTitle(sheetId) {
  try {
    const r = await fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=properties.title`,
      { headers: { Authorization: `Bearer ${state.accessToken}` } }
    );
    if (!r.ok) return null;
    const j = await r.json();
    return j.properties?.title || null;
  } catch { return null; }
}

// ── Data ──────────────────────────────────────────────────────────────────────

async function loadRows() {
  const result = await apiGet('A:F');
  const all    = result.values || [];
  state.rows = all.slice(1)
    .map((r, i) => ({
      rowIndex:    i + 2,
      id:          (r[0] || '').trim(),
      name:         r[1] || '',
      address:      r[2] || '',
      notes:        r[3] || '',
      checkinStamp: r[4] || '',
      checkinBy:    r[5] || '',
    }))
    .filter(r => r.id);
}

async function refreshData() {
  setLoading(true);
  try {
    await loadRows();
    if (state.currentScreen === 'screen-list')  renderList();
    if (state.currentScreen === 'screen-admin') renderAdminReport();
  } catch { alert('重新整理失敗，請檢查網路。'); }
  finally  { setLoading(false); }
}

// ── Check-in Logic ────────────────────────────────────────────────────────────

// Re-read the row before writing. Guards against:
//  1. rows inserted/deleted/re-sorted in the sheet (stale rowIndex → wrong person)
//  2. check-ins made from another device since our last load
async function verifyRow(row) {
  const res = await apiGet(`A${row.rowIndex}:F${row.rowIndex}`);
  const cur = (res.values && res.values[0]) || [];

  if ((cur[0] || '').trim() !== row.id) {
    // Sheet rows have shifted — reload and re-locate this person by ID
    await loadRows();
    const fresh = state.rows.find(r => r.id === row.id);
    if (!fresh) {
      const e = new Error('找不到此筆資料，名單可能已被修改。');
      e.code = 'gone';
      throw e;
    }
    return fresh;
  }
  // Sync latest check-in state from the sheet
  row.checkinStamp = cur[4] || '';
  row.checkinBy    = cur[5] || '';
  return row;
}

async function checkin(row, manual = false) {
  row = await verifyRow(row);
  if (row.checkinStamp) {
    const e = new Error('already');
    e.code = 'already';
    e.row  = row;
    throw e;
  }
  const stamp = nowString();
  const by    = state.user.name + (manual ? ' 手動' : '');
  await apiPut(`E${row.rowIndex}:F${row.rowIndex}`, [[stamp, by]]);
  row.checkinStamp = stamp;
  row.checkinBy    = by;
  return row;
}

async function undoCheckin(row) {
  row = await verifyRow(row);
  await apiPut(`E${row.rowIndex}:F${row.rowIndex}`, [['', '']]);
  row.checkinStamp = '';
  row.checkinBy    = '';
  return row;
}

// ── Auto-refresh (keeps multi-device data in sync) ────────────────────────────

let autoRefreshTimer = null;

function startAutoRefresh() {
  if (autoRefreshTimer) return;
  autoRefreshTimer = setInterval(async () => {
    if (!state.sheetId || document.hidden) return;
    try {
      await loadRows();
      if (state.currentScreen === 'screen-list')  renderList();
      if (state.currentScreen === 'screen-admin') renderAdminReport();
    } catch { /* silent — next tick will retry */ }
  }, 45000);
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
      () => {}
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
  let row    = state.rows.find(r => r.id === code);

  if (!row) {
    // Might be a person added from another device — reload once and retry
    try {
      await loadRows();
      row = state.rows.find(r => r.id === code);
    } catch {}
  }

  if (!row) {
    scanFeedback('not-found');
    flashResult('not-found', '未找到', code, null);
  } else if (row.checkinStamp) {
    scanFeedback('already');
    flashResult('already', '已掃過', code, row);
  } else {
    try {
      row = await checkin(row);
      scanFeedback('success');
      flashResult('success', '報到成功', code, row);
    } catch (e) {
      if (e.code === 'already') {
        scanFeedback('already');
        flashResult('already', '已掃過', code, e.row);
      } else {
        scanFeedback('error');
        flashResult('err', '寫入失敗', code, null);
      }
    }
  }

  state.cooldownTimer = setTimeout(dismissScanResult, 3000);
}

// Tap the banner (or wait 3 s) to scan the next person
function dismissScanResult() {
  clearTimeout(state.cooldownTimer);
  state.cooldownTimer  = null;
  state.scanCooldown   = false;
  resetBanner();
}

// ── Scan feedback (beep + vibration) ─────────────────────────────────────────

let audioCtx = null;

function scanFeedback(type) {
  try {
    if (navigator.vibrate) {
      navigator.vibrate(
        type === 'success' ? 80 :
        type === 'already' ? [60, 50, 60] : [220]
      );
    }
    audioCtx = audioCtx || new (window.AudioContext || window.webkitAudioContext)();
    const beep = (freq, t0, dur) => {
      const o = audioCtx.createOscillator();
      const g = audioCtx.createGain();
      o.type = 'sine';
      o.frequency.value = freq;
      g.gain.setValueAtTime(0.25, audioCtx.currentTime + t0);
      g.gain.exponentialRampToValueAtTime(0.001, audioCtx.currentTime + t0 + dur);
      o.connect(g); g.connect(audioCtx.destination);
      o.start(audioCtx.currentTime + t0);
      o.stop(audioCtx.currentTime + t0 + dur);
    };
    if      (type === 'success') beep(1200, 0, .15);
    else if (type === 'already') { beep(800, 0, .12); beep(800, .18, .12); }
    else                         beep(300, 0, .4);
  } catch { /* audio unavailable — vibration/visual still works */ }
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
    notesEl.textContent   = row.notes;
    notesEl.style.display = row.notes ? '' : 'none';

    const stampEl = document.getElementById('scan-info-stamp');
    stampEl.textContent   = row.checkinStamp ? '報到時間：' + row.checkinStamp : '';
    stampEl.style.display = row.checkinStamp ? '' : 'none';

    info.style.display = '';
  } else {
    info.style.display = 'none';
  }

  // Offer walk-in registration when the code isn't in the list
  const walkinBtn = document.getElementById('scan-walkin-btn');
  if (walkinBtn) walkinBtn.style.display = (type === 'not-found') ? '' : 'none';
}

function resetBanner() {
  const banner = document.getElementById('result-banner');
  banner.className = 'result-banner result-idle';
  document.getElementById('result-main').textContent   = '準備掃描';
  document.getElementById('result-detail').textContent = '';
  document.getElementById('scan-info').style.display   = 'none';
  const walkinBtn = document.getElementById('scan-walkin-btn');
  if (walkinBtn) walkinBtn.style.display = 'none';
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
    state.selectedRow = await checkin(state.selectedRow, true);
    renderManual();
    renderList();
  } catch (e) {
    if (e.code === 'already') {
      if (e.row) state.selectedRow = e.row;
      renderManual();
      renderList();
      alert('此人剛剛已在其他裝置完成報到。');
    } else if (e.code === 'gone') {
      alert(e.message);
      goTo('screen-list');
    } else {
      alert('操作失敗，請檢查網路後重試。');
    }
  }
  finally { setLoading(false); }
}

async function manualUndo() {
  setLoading(true);
  try {
    state.selectedRow = await undoCheckin(state.selectedRow);
    renderManual();
    renderList();
  } catch (e) {
    if (e.code === 'gone') { alert(e.message); goTo('screen-list'); }
    else alert('操作失敗，請檢查網路後重試。');
  }
  finally { setLoading(false); }
}

// ── Admin Screen ──────────────────────────────────────────────────────────────

function renderAdmin() {
  renderAdminReport();
  renderAllowListSection(); // async, self-managing
}

// ── Report (within Admin) ─────────────────────────────────────────────────────

function renderAdminReport() {
  const hasEvent = !!state.sheetId;

  document.getElementById('event-none').style.display = hasEvent ? 'none' : '';
  document.getElementById('event-info').style.display = hasEvent ? '' : 'none';
  document.getElementById('share-box').style.display  = hasEvent ? '' : 'none';
  renderRecentEvents();

  const statsEl = document.getElementById('report-stats');

  if (!hasEvent) {
    statsEl.innerHTML = '<p class="msg-empty" style="padding:8px 0">尚未選擇活動</p>';
    return;
  }

  const total   = state.rows.length;
  const checked = state.rows.filter(r => r.checkinStamp).length;
  const title   = state.spreadsheetTitle || '';

  document.getElementById('report-event-title').textContent = title || '（未命名活動）';
  document.getElementById('report-summary').textContent =
    `總數 ${total}　已報到 ${checked}　未報到 ${total - checked}`;
  document.getElementById('event-open-link').href =
    `https://docs.google.com/spreadsheets/d/${state.sheetId}/edit`;
  document.getElementById('share-url').textContent = location.href;
  renderShareQr();

  statsEl.innerHTML = `
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

function renderShareQr() {
  const el = document.getElementById('share-qr');
  if (!el) return;
  el.innerHTML = '';
  try {
    const qr = qrcode(0, 'M');
    qr.addData(location.href);
    qr.make();
    el.innerHTML = qr.createSvgTag({ cellSize: 4, margin: 2 });
  } catch {}
}

// ── Attendee QR codes ─────────────────────────────────────────────────────────

// Fill in IDs for attendees that have a name but no QR value in column A
async function ensureIds() {
  const res  = await apiGet('A:F');
  const all  = res.values || [];
  const used = new Set(all.slice(1).map(r => (r[0] || '').trim()).filter(Boolean));

  let n = 1;
  const nextId = () => {
    let id;
    do { id = 'A' + String(n++).padStart(3, '0'); } while (used.has(id));
    used.add(id);
    return id;
  };

  const data = [];
  all.slice(1).forEach((r, i) => {
    const hasName = (r[1] || '').trim();
    const hasId   = (r[0] || '').trim();
    if (hasName && !hasId) {
      data.push({ range: `A${i + 2}`, majorDimension: 'ROWS', values: [[nextId()]] });
    }
  });

  if (data.length) {
    await sheetsRequest('POST', `/${state.sheetId}/values:batchUpdate`,
      { valueInputOption: 'RAW', data });
  }
  return data.length;
}

async function generateQrCards() {
  hideError('qrgen-error');
  if (!state.sheetId) { showError('qrgen-error', '請先建立或選擇活動。'); return; }

  setLoading(true);
  try {
    await ensureIds();
    await loadRows();
    if (!state.rows.length) {
      showError('qrgen-error', '名單是空的。請先按「在 Google 試算表開啟名單」，在 B 欄填入參加者姓名。');
      return;
    }
    openQrPrintWindow();
  } catch (err) {
    showError('qrgen-error', '產生失敗：' + err.message);
  } finally {
    setLoading(false);
  }
}

function openQrPrintWindow() {
  const cards = state.rows.map(r => {
    let svg = '';
    try {
      const q = qrcode(0, 'M');
      q.addData(r.id);
      q.make();
      svg = q.createSvgTag({ cellSize: 4, margin: 0 });
    } catch {}
    return `<div class="qr-card">${svg}<div class="qr-id">${esc(r.id)}</div><div class="qr-name">${esc(r.name)}</div></div>`;
  }).join('');

  const html = `<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">
<title>${esc(state.spreadsheetTitle || '')} — 參加者 QR 碼</title><style>
  body{font-family:'Segoe UI','PingFang TC','Microsoft JhengHei',sans-serif;margin:20px}
  h1{font-size:18px;text-align:center;margin-bottom:12px}
  .toolbar{text-align:center;margin-bottom:16px}
  .toolbar button{font-size:16px;padding:10px 28px;cursor:pointer;border-radius:8px;border:1px solid #ccc;background:#3860B2;color:#fff}
  .grid{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
  .qr-card{border:1px dashed #999;border-radius:8px;padding:14px 8px;text-align:center;page-break-inside:avoid}
  .qr-card svg{width:110px;height:110px}
  .qr-id{font-size:12px;color:#666;margin-top:6px}
  .qr-name{font-size:17px;font-weight:700}
  @media print{.toolbar{display:none}}
</style></head><body>
<h1>${esc(state.spreadsheetTitle || '')} — 參加者 QR 碼</h1>
<div class="toolbar"><button onclick="print()">🖨 列印</button></div>
<div class="grid">${cards}</div>
</body></html>`;

  const w = window.open('', '_blank');
  if (!w) { alert('瀏覽器封鎖了彈出視窗，請允許彈出視窗後再試一次。'); return; }
  w.document.write(html);
  w.document.close();
}

// ── Printable attendance report ───────────────────────────────────────────────

function printReport() {
  if (!state.sheetId) { alert('請先建立或選擇活動。'); return; }

  const total   = state.rows.length;
  const checked = state.rows.filter(r => r.checkinStamp).length;
  const rowsHtml = state.rows.map((r, i) => `
    <tr${r.checkinStamp ? ' class="in"' : ''}>
      <td>${i + 1}</td><td>${esc(r.id)}</td><td>${esc(r.name)}</td>
      <td>${esc(r.address)}</td><td>${esc(r.notes)}</td>
      <td>${esc(r.checkinStamp)}</td><td>${esc(r.checkinBy)}</td>
    </tr>`).join('');

  const html = `<!DOCTYPE html><html lang="zh-TW"><head><meta charset="UTF-8">
<title>${esc(state.spreadsheetTitle || '')} — 出席報告</title><style>
  body{font-family:'Segoe UI','PingFang TC','Microsoft JhengHei',sans-serif;margin:24px;color:#1a1a1a}
  h1{font-size:20px;margin-bottom:4px}
  .meta{font-size:13px;color:#666;margin-bottom:12px}
  .summary{font-size:15px;font-weight:600;margin-bottom:16px}
  .toolbar{margin-bottom:16px}
  .toolbar button{font-size:16px;padding:10px 28px;cursor:pointer;border-radius:8px;border:1px solid #ccc;background:#3860B2;color:#fff}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th,td{border:1px solid #ccc;padding:6px 8px;text-align:left}
  th{background:#f0f2f5}
  tr.in td{background:#f2fbf3}
  @media print{.toolbar{display:none}}
</style></head><body>
<h1>${esc(state.spreadsheetTitle || '')} 出席報告</h1>
<div class="meta">產生時間：${nowString()}</div>
<div class="summary">總數 ${total}　已報到 ${checked}　未報到 ${total - checked}　出席率 ${total ? Math.round(checked / total * 100) : 0}%</div>
<div class="toolbar"><button onclick="print()">🖨 列印 / 存成 PDF</button></div>
<table>
  <tr><th>#</th><th>編號</th><th>姓名</th><th>通訊地址</th><th>備註</th><th>報到時間</th><th>報到人員</th></tr>
  ${rowsHtml}
</table>
</body></html>`;

  const w = window.open('', '_blank');
  if (!w) { alert('瀏覽器封鎖了彈出視窗，請允許彈出視窗後再試一次。'); return; }
  w.document.write(html);
  w.document.close();
}

function downloadReport() {
  if (!state.sheetId) { alert('請先建立或選擇活動。'); return; }
  const header  = ['標題', '姓名', '通訊地址', '備註', '報到時間', '報到人員'];
  const csvRows = [
    header,
    ...state.rows.map(r => [r.id, r.name, r.address, r.notes, r.checkinStamp, r.checkinBy]),
  ];
  const csv  = csvRows.map(row =>
    row.map(cell => `"${String(cell ?? '').replace(/"/g, '""')}"`).join(',')
  ).join('\r\n');

  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href     = url;
  a.download = `${state.spreadsheetTitle || '出席報告'}_${dateString()}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ── Walk-in ───────────────────────────────────────────────────────────────────

function generateWalkInId() {
  const d   = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `W-${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

async function addWalkIn() {
  hideError('walkin-error');
  document.getElementById('walkin-success').style.display = 'none';

  const name    = document.getElementById('walkin-name').value.trim();
  const address = document.getElementById('walkin-address').value.trim();
  const notes   = document.getElementById('walkin-notes').value.trim() || '現場報到，已核實身分';

  if (!name)          { showError('walkin-error', '請填寫姓名。'); return; }
  if (!state.sheetId) { showError('walkin-error', '請先在「活動設定」中載入試算表。'); return; }

  setLoading(true);
  try {
    const id     = generateWalkInId();
    const stamp  = nowString();
    const by     = state.user.name + ' 手動';

    const result = await apiAppend('A:F', [[id, name, address, notes, stamp, by]]);

    // Parse actual row index from the API response
    const rangeMatch = (result.updates?.updatedRange || '').match(/!A(\d+)/);
    const rowIndex   = rangeMatch ? parseInt(rangeMatch[1]) : (state.rows.length + 2);

    state.rows.push({ rowIndex, id, name, address, notes, checkinStamp: stamp, checkinBy: by });

    // Reset form (keep notes default)
    document.getElementById('walkin-name').value    = '';
    document.getElementById('walkin-address').value = '';
    document.getElementById('walkin-notes').value   = '現場報到，已核實身分';

    const successEl = document.getElementById('walkin-success');
    successEl.textContent   = `✓ ${name} 已新增並報到`;
    successEl.style.display = '';
    setTimeout(() => { successEl.style.display = 'none'; }, 4000);

    renderAdminReport();
  } catch (err) {
    showError('walkin-error', '新增失敗：' + err.message);
  } finally {
    setLoading(false);
  }
}

// ── Allow List ────────────────────────────────────────────────────────────────

async function loadAllowListEmails() {
  const result = await sheetsGet(state.allowListSheetId, 'A:A');
  const all    = result.values || [];
  return all.slice(1)
    .map((r, i) => ({ rowIndex: i + 2, email: (r[0] || '').trim().toLowerCase() }))
    .filter(e => e.email);
}


async function renderAllowListSection() {
  const container = document.getElementById('allowlist-container');
  if (!container) return;

  if (!state.allowListSheetId) {
    container.innerHTML = '<p class="msg-error">未設定允許名單（請更新 config.js）</p>';
    return;
  }

  container.innerHTML = '<p class="allowlist-loading">載入中…</p>';

  try {
    const emails = await loadAllowListEmails();
    hideError('allowlist-add-error');

    if (emails.length === 0) {
      container.innerHTML = '<p class="msg-empty" style="padding:12px 0">名單為空</p>';
      return;
    }

    container.innerHTML = `
      <div class="allowlist-list">
        ${emails.map(e => `
          <div class="allowlist-row">
            <span class="allowlist-email">${esc(e.email)}</span>
            <button class="btn-icon-danger" onclick="removeAllowListEmail(${e.rowIndex})">✕</button>
          </div>
        `).join('')}
      </div>
    `;
  } catch (err) {
    container.innerHTML = `<p class="msg-error">無法讀取名單：${esc(err.message)}</p>`;
  }
}

async function addAllowListEmail() {
  hideError('allowlist-add-error');
  const input = document.getElementById('new-email');
  const email = input.value.trim().toLowerCase();
  if (!email || !email.includes('@')) {
    showError('allowlist-add-error', '請輸入有效的電子郵件地址。');
    return;
  }

  setLoading(true);
  try {
    await sheetsAppend(state.allowListSheetId, 'A:A', [[email]]);
    input.value = '';
    await renderAllowListSection();
  } catch (err) {
    showError('allowlist-add-error', '新增失敗：' + err.message);
  } finally {
    setLoading(false);
  }
}

async function removeAllowListEmail(rowIndex) {
  if (!confirm('確定要移除此帳號？')) return;
  setLoading(true);
  try {
    await sheetsPut(state.allowListSheetId, `A${rowIndex}`, [['']]);
    await renderAllowListSection();
  } catch (err) {
    alert('移除失敗：' + err.message);
  } finally {
    setLoading(false);
  }
}

function resetWalkInForm() {
  hideError('walkin-error');
  document.getElementById('walkin-success').style.display = 'none';
  document.getElementById('walkin-name').value    = '';
  document.getElementById('walkin-address').value = '';
  document.getElementById('walkin-notes').value   = '現場報到，已核實身分';
}

// ── Navigation ────────────────────────────────────────────────────────────────

function goTo(screenId) {
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
  if (screenId === 'screen-walkin') resetWalkInForm();
  if (screenId === 'screen-admin')  renderAdmin();
}

function showScreen(id) {
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
