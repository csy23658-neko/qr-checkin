'use strict';

// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  idToken:          null,
  user:             null,   // { email, name }
  role:             null,   // 'admin' | 'volunteer'
  sheetId:          null,
  spreadsheetTitle: null,
  allowListSheetId: null,
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

function initAuth() {
  google.accounts.id.initialize({
    client_id: CONFIG.GOOGLE_CLIENT_ID,
    callback: onGoogleCredential,
    auto_select: false,
  });
  google.accounts.id.renderButton(document.getElementById('google-signin-button'), {
    type: 'standard', theme: 'outline', size: 'large', text: 'signin_with',
    shape: 'rectangular', logo_alignment: 'left', width: 304, locale: 'zh_TW',
  });
}

async function onGoogleCredential(resp) {
  state.idToken = resp.credential;
  try {
    setLoading(true);
    const session = await apiAction('session');
    state.user = session.user;
    state.role = session.user.role;
  } catch (error) { showLoginError(error.message || '登入失敗。'); return; }
  finally { setLoading(false); }

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

async function apiAction(action, payload = {}) {
  const response = await fetch((CONFIG.API_BASE_URL || '') + '/api', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${state.idToken}` },
    body: JSON.stringify({ action, ...payload }),
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(result.error || `API error (${response.status})`);
  return result;
}

async function checkAllowList(email) {
  if (!state.allowListSheetId) return false;
  const emails = await loadAllowListEmails();
  return emails.some(e => e.email === email);
}

// ── Event Setup ───────────────────────────────────────────────────────────────

const EVENT_HEADERS = ['編號', '姓名', '通訊地址', '備註', '報到時間', '報到人員'];
const IMPORT_HEADERS = EVENT_HEADERS.slice(0, 4);

function downloadActivityTemplate() {
  const a = document.createElement('a');
  a.href = '活動名單範本.xlsx';
  a.download = '活動名單範本.xlsx';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

function downloadMergeList() {
  if (!state.sheetId) { alert('請先建立或選擇活動。'); return; }
  if (typeof XLSX === 'undefined') { alert('匯出工具尚未載入，請重新整理後再試。'); return; }
  const ids = state.rows.map(r => r.id);
  if (new Set(ids).size !== ids.length) { alert('名單中有重複編號，請修正後再匯出。'); return; }
  const rows = [IMPORT_HEADERS, ...state.rows.map(r => [r.id, r.name, r.address, r.notes])];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = [{ wch: 18 }, { wch: 16 }, { wch: 34 }, { wch: 26 }];
  for (let row = 1; row <= rows.length; row++) {
    for (let col = 0; col < IMPORT_HEADERS.length; col++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: col })];
      if (cell) cell.z = '@';
    }
  }
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '名單');
  XLSX.writeFile(book, `${state.spreadsheetTitle || '活動'}_Word合併列印名單.xlsx`);
}

async function importActivityTemplate() {
  hideError('create-error');
  const name = document.getElementById('input-event-name').value.trim();
  const file = document.getElementById('input-activity-file').files[0];
  if (!name) { showError('create-error', '請輸入活動名稱。'); return; }
  if (!file) { showError('create-error', '請選擇已填寫的活動名單範本。'); return; }
  if (typeof XLSX === 'undefined') { showError('create-error', '匯入工具尚未載入，請重新整理後再試。'); return; }

  setLoading(true);
  try {
    const book = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const sheet = book.Sheets['名單'] || book.Sheets[book.SheetNames[0]];
    if (!sheet) throw new Error('找不到名為「名單」的工作表。');
    const values = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
    const headers = (values[0] || []).slice(0, 4).map(v => String(v).trim());
    if (headers.join('|') !== IMPORT_HEADERS.join('|')) {
      throw new Error('欄位必須依序為：編號、姓名、通訊地址、備註。');
    }

    const seen = new Set();
    const attendees = values.slice(1).map((row, i) => {
      const item = row.slice(0, 4).map(value => String(value ?? '').trim());
      if (item.every(value => !value)) return null;
      if (!item[0]) throw new Error(`第 ${i + 2} 列缺少編號。`);
      if (!item[1]) throw new Error(`第 ${i + 2} 列缺少姓名。`);
      if (seen.has(item[0])) throw new Error(`編號「${item[0]}」重複。`);
      seen.add(item[0]);
      return item;
    }).filter(Boolean);
    if (!attendees.length) throw new Error('名單沒有任何參與者資料。');

    await createEventWithAttendees(name, attendees);
    document.getElementById('input-activity-file').value = '';
  } catch (err) {
    showError('create-error', '匯入失敗：' + err.message);
  } finally {
    setLoading(false);
  }
}

// One-click event creation: makes a new Google Sheet with the right columns,
// then shares it with everyone on the volunteer allow list.
async function createEvent() {
  hideError('create-error');
  setCreateStatus('');
  const name = document.getElementById('input-event-name').value.trim();
  if (!name) { showError('create-error', '請輸入活動名稱。'); return; }

  const button = document.getElementById('btn-create-event');
  button.disabled = true;
  button.textContent = '建立活動中…';
  setCreateStatus(`正在建立「${name}」的活動名單…`);
  setLoading(true);
  try {
    const created = await createEventWithAttendees(name, []);
    setCreateStatus(`已建立「${created.title || name}」。您現在可以下載範本、填寫後再上傳名單。`);
  } catch (err) {
    setCreateStatus('');
    showError('create-error', '建立失敗：' + err.message);
  } finally {
    button.disabled = false;
    button.textContent = '＋ 建立新活動';
    setLoading(false);
  }
}

async function createEventWithAttendees(name, attendees) {
    const created = await apiAction('event.create', { name, attendees: attendees.map(row => ({ id: row[0], name: row[1], address: row[2], notes: row[3] })) });
    state.sheetId = created.sheetId;
    state.spreadsheetTitle = created.title;
    state.rows = created.rows;
    document.getElementById('input-event-name').value = '';
    afterEventSelected();
    await shareWithVolunteers(true);
    return created;
}

// Advanced: load an existing spreadsheet by URL
async function confirmSetup() {
  hideError('setup-error');
  const url = document.getElementById('input-sheet-url').value.trim();
  const m   = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (!m) { showError('setup-error', '網址格式不正確，請重新貼上 Google 試算表網址。'); return; }

  setLoading(true);
  try {
    const imported = await apiAction('event.import', { sheetId: m[1] });
    state.sheetId = imported.sheetId;
    state.spreadsheetTitle = imported.title;
    state.rows = imported.rows;
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
  try {
    const shared = await apiAction('event.share', { sheetId: state.sheetId });
    if (msgEl) {
      msgEl.textContent = `✓ 已共用給 ${shared.shared} 位志工`;
      msgEl.className = 'msg-success';
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

async function fetchSpreadsheetTitle(sheetId) {
  try { return (await apiAction('event.get', { sheetId })).title || null; } catch { return null; }
}

// ── Data ──────────────────────────────────────────────────────────────────────

async function loadRows() {
  const event = await apiAction('event.get', { sheetId: state.sheetId });
  state.spreadsheetTitle = event.title;
  state.rows = event.rows;
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

async function checkin(row, manual = false) {
  const result = await apiAction('checkin', { sheetId: state.sheetId, id: row.id, manual });
  if (result.already) {
    const error = new Error('already');
    error.code = 'already';
    error.row = result.row;
    throw error;
  }
  return result.row;
}

async function undoCheckin(row) {
  return (await apiAction('undo', { sheetId: state.sheetId, id: row.id })).row;
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
  const title = document.getElementById('dashboard-title');
  const subtitle = document.getElementById('dashboard-subtitle');
  if (title) title.textContent = state.spreadsheetTitle || '建立並管理您的活動';
  if (subtitle) subtitle.textContent = state.sheetId
    ? '活動名單、列印、志工授權與出席報告皆集中在這裡。'
    : '先建立新活動或匯入填好的名單範本，即可開始管理。';
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

// A participant number is assigned by the event organizer and must never change.
async function ensureIds() {
  const used = new Set(state.rows.map(row => row.id).filter(Boolean));
  const missing = state.rows.filter(row => row.name && !row.id).map(row => row.rowIndex);
  if (missing.length) throw new Error(`第 ${missing.join('、')} 列缺少編號；編號必須由活動方編入且不可自動產生。`);
  if (used.size !== state.rows.length) {
    throw new Error('名單中有重複編號，請修正後再列印。');
  }
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
  const header  = EVENT_HEADERS;
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
    const result = await apiAction('walkin.add', { sheetId: state.sheetId, name, address, notes });
    state.rows = result.rows;

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
  return (await apiAction('allowlist', { method: 'GET' })).entries;
}


async function renderAllowListSection() {
  const container = document.getElementById('allowlist-container');
  if (!container) return;

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
            <button class="btn-icon-danger" onclick="removeAllowListEmail('${esc(e.email)}')">✕</button>
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
    await apiAction('allowlist', { method: 'POST', email });
    input.value = '';
    await renderAllowListSection();
  } catch (err) {
    showError('allowlist-add-error', '新增失敗：' + err.message);
  } finally {
    setLoading(false);
  }
}

async function removeAllowListEmail(email) {
  if (!confirm('確定要移除此帳號？')) return;
  setLoading(true);
  try {
    await apiAction('allowlist', { method: 'DELETE', email });
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

function setCreateStatus(message) {
  const el = document.getElementById('create-status');
  if (!el) return;
  el.textContent = message;
  el.style.display = message ? '' : 'none';
}

function hideError(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = 'none';
}
