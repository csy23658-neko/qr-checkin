'use strict';

// ── State ─────────────────────────────────────────────────────────────────────

const state = {
  idToken:          null,
  driveAccessToken: null,
  user:             null,   // { email, name }
  role:             null,   // 'admin' | 'volunteer'
  sheetId:          null,
  spreadsheetTitle: null,
  rosterReady:      false,
  pendingRosterEventId: null,
  allowListSheetId: null,
  rows:             [],     // [{ rowIndex, id, name, address, notes, checkinStamp, checkinBy }]
  eventMeta:        { startDate: '', endDate: '', startTime: '', endTime: '', location: '', notice: '', templateId: '', navUrl: '' },
  templates:        [],     // 活動通知範本 list from the API
  templatesError:   '',
  templatesLoaded:  false,
  templateFiles:    {},     // id -> { updatedAt, bytes } cache of downloaded templates
  driveReject:      null,
  batchDigits:      3,
  batches:          [],
  noticeRow:        null,
  selectedRow:      null,
  scanner:          null,
  scanCooldown:     false,
  clockTimer:       null,
  currentScreen:    'screen-login',
};

// ── Boot ──────────────────────────────────────────────────────────────────────

window.addEventListener('load', () => {
  state.pendingRosterEventId = loadPendingRosterEventId();
  const createButton = document.getElementById('btn-create-event');
  if (createButton) createButton.addEventListener('click', createEvent);
  const downloadButton = document.getElementById('btn-download-template');
  if (downloadButton) downloadButton.addEventListener('click', downloadActivityTemplate);
  const importButton = document.getElementById('btn-import-template');
  if (importButton) importButton.addEventListener('click', importActivityTemplate);
  const batchDigits = document.getElementById('batch-digits');
  if (batchDigits) batchDigits.addEventListener('change', () => {
    const current = readBatchForm();
    state.batchDigits = current.digits;
    state.batches = current.batches;
    renderBatchEditor();
  });
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  updateActivityPreviewMeta();
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
  state.driveTokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.GOOGLE_CLIENT_ID,
    // `drive.file` limits access to files this app creates or that the user
    // explicitly opens with it.  It is sufficient for creating an activity
    // Sheet and avoids asking for access to every spreadsheet in the account.
    scope: 'https://www.googleapis.com/auth/drive.file',
    callback: () => {},
    error_callback: error => {
      const reject = state.driveReject;
      state.driveReject = null;
      if (!reject) return;
      reject(new Error(error?.type === 'popup_failed_to_open'
        ? 'Google 授權視窗被瀏覽器封鎖。請允許此網站開啟彈出視窗後再試一次。'
        : '未完成 Google Drive 授權（授權視窗已關閉）。'));
    },
  });
}

function ensureDriveAccess() {
  if (state.driveAccessToken) return Promise.resolve(state.driveAccessToken);
  return new Promise((resolve, reject) => {
    state.driveReject = reject;
    state.driveTokenClient.callback = response => {
      state.driveReject = null;
      if (response.error || !response.access_token) return reject(new Error('未取得建立活動所需的 Google Drive 授權。'));
      state.driveAccessToken = response.access_token;
      resolve(response.access_token);
    };
    state.driveTokenClient.requestAccessToken({ prompt: 'consent' });
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
  } else {
    // Keep the first post-login screen consistent for both roles.  The scan
    // screen will explain how to load an activity when no sheet URL exists.
    goTo('screen-scan');
  }
}

async function apiAction(action, payload = {}) {
  const response = await fetch((CONFIG.API_BASE_URL || '') + '/api', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${state.idToken}` },
    body: JSON.stringify({ action, ...payload }),
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(result.error || `API error (${response.status})`);
    error.status = response.status;
    if (Array.isArray(result.details)) error.details = result.details;
    throw error;
  }
  return result;
}

async function checkAllowList(email) {
  if (!state.allowListSheetId) return false;
  const emails = await loadAllowListEmails();
  return emails.some(e => e.email === email);
}

// ── Event Setup ───────────────────────────────────────────────────────────────

const EVENT_HEADERS = ['編號', '姓名', '通訊地址', '備註', '報到時間', '報到人員', '通知報到時段'];
const IMPORT_HEADERS = EVENT_HEADERS.slice(0, 4);
const PENDING_ROSTER_EVENT_KEY = 'qr-pending-roster-event';

function eventMetadataFromForm() {
  const value = id => document.getElementById(id)?.value.trim() || '';
  return {
    startDate: value('input-event-start-date'),
    endDate: value('input-event-end-date'),
    startTime: value('input-event-start-time'),
    endTime: value('input-event-end-time'),
    location: value('input-event-location'),
    notice: document.getElementById('input-event-notice')?.value.replace(/\s+$/, '') || '',
    templateId: value('input-event-template'),
    navUrl: value('input-event-nav-url'),
  };
}

// syncForm: false is used by the 5-second background refresh.  It keeps the
// saved event data current without overwriting fields the administrator is
// still editing (re-rendering the form also dropped keyboard focus).
function applyEventMetadata(event, { syncForm = true } = {}) {
  state.eventMeta = {
    startDate: event?.startDate || '',
    endDate: event?.endDate || '',
    startTime: event?.startTime || '',
    endTime: event?.endTime || '',
    location: event?.location || '',
    notice: event?.notice || '',
    templateId: event?.templateId || '',
    navUrl: event?.navUrl || '',
  };
  if (!syncForm) {
    renderNoticeTemplateInfo();
    return;
  }
  state.batchDigits = Number(event?.batchDigits) || 3;
  state.batches = Array.isArray(event?.batches) ? event.batches.map(batch => ({
    min: String(batch.min || ''), max: String(batch.max || ''), time: String(batch.time || ''),
  })) : [];
  const values = {
    'input-event-start-date': state.eventMeta.startDate,
    'input-event-end-date': state.eventMeta.endDate,
    'input-event-start-time': state.eventMeta.startTime,
    'input-event-end-time': state.eventMeta.endTime,
    'input-event-location': state.eventMeta.location,
    'input-event-notice': state.eventMeta.notice,
    'input-event-nav-url': state.eventMeta.navUrl,
    'batch-digits': state.batchDigits,
  };
  Object.entries(values).forEach(([id, current]) => {
    const el = document.getElementById(id);
    if (el) el.value = current;
  });
  renderTemplateSelect();
  renderNoticeTemplateInfo();
  renderBatchEditor();
}

function clearEventMetadataForm() {
  applyEventMetadata({});
}

function loadPendingRosterEventId() {
  try { return localStorage.getItem(PENDING_ROSTER_EVENT_KEY); } catch { return null; }
}

function setPendingRosterEventId(sheetId) {
  state.pendingRosterEventId = sheetId || null;
  try {
    if (state.pendingRosterEventId) localStorage.setItem(PENDING_ROSTER_EVENT_KEY, state.pendingRosterEventId);
    else localStorage.removeItem(PENDING_ROSTER_EVENT_KEY);
  } catch {}
}

function canUploadRoster() {
  return !!state.sheetId && !state.rosterReady && state.pendingRosterEventId === state.sheetId;
}

function clearImportWorkflow() {
  hideError('import-error');
  setImportStatus('');
  const file = document.getElementById('input-activity-file');
  if (file) file.value = '';
}

function clearActivityWorkflowMessages() {
  hideError('create-error');
  setCreateStatus('');
  clearImportWorkflow();
}

function updateRosterUploadControls() {
  const allowed = canUploadRoster();
  const file = document.getElementById('input-activity-file');
  const button = document.getElementById('btn-import-template');
  const hint = document.getElementById('import-availability-hint');
  if (file) file.disabled = !allowed;
  if (button) button.disabled = !allowed;
  if (hint) {
    hint.textContent = allowed
      ? `可上傳名單至剛建立的「${state.spreadsheetTitle || '新活動'}」。上傳成功後即鎖定名單。`
      : state.sheetId
      ? '目前載入的是既有活動，名單僅供查看與報告下載；若要匯入另一份名單，請先建立新活動。'
      : '請先按「＋ 建立新活動」，再上傳該活動的名單。';
  }
}

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
  const mergeHeaders = [...IMPORT_HEADERS, '報到時段', 'QR內容'];
  const rows = [mergeHeaders, ...state.rows.map(r => [r.id, r.name, r.address, r.notes, r.batchTime || '', qrPayloadFor(r.id)])];
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = [{ wch: 18 }, { wch: 16 }, { wch: 34 }, { wch: 26 }, { wch: 18 }, { wch: 42 }];
  for (let row = 1; row <= rows.length; row++) {
    for (let col = 0; col < mergeHeaders.length; col++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row - 1, c: col })];
      if (cell) cell.z = '@';
    }
  }
  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '名單');
  XLSX.writeFile(book, `${state.spreadsheetTitle || '活動'}_名單匯出.xlsx`);
}

async function importActivityTemplate() {
  hideError('create-error');
  hideError('import-error');
  setImportStatus('');
  const file = document.getElementById('input-activity-file').files[0];
  if (!canUploadRoster()) {
    showError('import-error', state.sheetId
      ? '為避免覆寫既有活動，無法上傳名單至目前活動。請先建立新活動。'
      : '請先按「＋ 建立新活動」，再上傳該活動的名單。');
    return;
  }
  if (!file) { showError('import-error', '請選擇已填寫的活動名單範本。'); return; }
  if (typeof XLSX === 'undefined') { showError('import-error', '匯入工具尚未載入，請重新整理後再試。'); return; }

  const button = document.getElementById('btn-import-template');
  button.disabled = true;
  button.textContent = '正在上傳活動名單…';
  setImportStatus(`正在讀取「${file.name}」…`);
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

    const currentActivityTitle = state.spreadsheetTitle || '目前活動';
    setImportStatus(`正在將 ${attendees.length} 位參與者匯入「${currentActivityTitle}」…`);
    const updated = await apiAction('event.roster.replace', {
      sheetId: state.sheetId,
      attendees: attendees.map(row => ({ id: row[0], name: row[1], address: row[2], notes: row[3] })),
    });
    state.spreadsheetTitle = state.spreadsheetTitle || updated.title;
    state.rows = updated.rows;
    state.rosterReady = true;
    setPendingRosterEventId(null);
    afterEventSelected();
    setImportStatus(`已成功匯入 ${attendees.length} 位參與者至「${state.spreadsheetTitle || currentActivityTitle}」。`, 8000);
    document.getElementById('input-activity-file').value = '';
  } catch (err) {
    setImportStatus('');
    showError('import-error', '上傳失敗：' + err.message);
  } finally {
    button.textContent = '⬆ 上傳活動名單';
    updateRosterUploadControls();
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
  const metadata = eventMetadataFromForm();

  clearCurrentEvent();

  const button = document.getElementById('btn-create-event');
  button.disabled = true;
  button.textContent = '建立活動中…';
  setCreateStatus(`正在建立「${name}」的空白活動資料表…`);
  setLoading(true);
  try {
    const created = await createEventWithAttendees(name, [], metadata);
    hideError('create-error');
    setCreateStatus(`已建立「${created.title || name}」。請下載範本、填妥參與者資料後上傳匯入。`);
  } catch (err) {
    setCreateStatus('');
    showError('create-error', '建立失敗：' + err.message);
  } finally {
    button.disabled = false;
    button.textContent = '＋ 建立新活動';
    setLoading(false);
  }
}

async function createEventWithAttendees(name, attendees, metadata = {}) {
    const driveAccessToken = await ensureDriveAccess();
    const created = await apiAction('event.create', {
      name, driveAccessToken, ...metadata,
      attendees: attendees.map(row => ({ id: row[0], name: row[1], address: row[2], notes: row[3] })),
    });
    state.sheetId = created.sheetId;
    state.spreadsheetTitle = created.title;
    state.rows = created.rows;
    state.rosterReady = attendees.length > 0;
    applyEventMetadata(created);
    setPendingRosterEventId(state.rosterReady ? null : state.sheetId);
    document.getElementById('input-event-name').value = '';
    if (state.rosterReady) {
      afterEventSelected();
    } else {
      syncActivityUrl();
      saveRecentEvent(state.sheetId, state.spreadsheetTitle || '');
      renderAdmin();
    }
    return created;
}

function afterEventSelected() {
  syncActivityUrl();
  saveRecentEvent(state.sheetId, state.spreadsheetTitle || '');
  renderAdmin();
  startAutoRefresh();
}

function clearCurrentEvent() {
  setPendingRosterEventId(null);
  state.sheetId = null;
  state.spreadsheetTitle = null;
  state.rows = [];
  state.rosterReady = false;
  state.eventMeta = { startDate: '', endDate: '', startTime: '', endTime: '', location: '', notice: '' };
  state.batchDigits = 3;
  state.batches = [];
  state.noticeRow = null;
  clearEventMetadataForm();
  state.selectedRow = null;
  clearActivityWorkflowMessages();
  syncActivityUrl();
  renderAdmin();
}

function activityShareUrl() {
  const url = new URL(location.href);
  const sheetId = state.sheetId || url.searchParams.get('sheet');
  const title = state.spreadsheetTitle || url.searchParams.get('name');
  url.search = '';
  if (sheetId) url.searchParams.set('sheet', sheetId);
  if (title) url.searchParams.set('name', title);
  return url.toString();
}

function syncActivityUrl() {
  const url = new URL(location.href);
  url.search = '';
  if (state.sheetId) url.searchParams.set('sheet', state.sheetId);
  if (state.spreadsheetTitle) url.searchParams.set('name', state.spreadsheetTitle);
  history.replaceState(null, '', url.pathname + url.search + url.hash);
  updateActivityPreviewMeta();
}

function updateActivityPreviewMeta() {
  const queryTitle = new URLSearchParams(location.search).get('name');
  const activityTitle = (state.spreadsheetTitle || queryTitle || '').trim();
  const title = activityTitle ? `QR 報到系統（${activityTitle}）` : 'QR 報到系統';
  const description = activityTitle
    ? `${title}，請使用 Google 帳號登入。`
    : 'QR 報到系統，請使用 Google 帳號登入。';
  document.title = title;
  const setMeta = (selector, content) => {
    const el = document.querySelector(selector);
    if (el) el.setAttribute('content', content);
  };
  setMeta('meta[name="description"]', description);
  setMeta('meta[property="og:title"]', title);
  setMeta('meta[property="og:description"]', description);
  setMeta('meta[property="og:url"]', activityShareUrl());
}

// ── Recent events (localStorage) ──────────────────────────────────────────────

const RECENT_EVENT_LIMIT = 2;

function getRecentEvents() {
  try {
    const list = JSON.parse(localStorage.getItem('qr-recent-events')) || [];
    return Array.isArray(list)
      ? list.filter(event => event && event.id)
      : [];
  }
  catch { return []; }
}

function saveRecentEvent(id, title) {
  const list = getRecentEvents().filter(e => e.id !== id);
  list.unshift({ id, title, ts: Date.now() });
  try { localStorage.setItem('qr-recent-events', JSON.stringify(list.slice(0, RECENT_EVENT_LIMIT))); } catch {}
}

function removeRecentEvent(id) {
  const list = getRecentEvents().filter(event => event.id !== id);
  try { localStorage.setItem('qr-recent-events', JSON.stringify(list.slice(0, RECENT_EVENT_LIMIT))); } catch {}
}

function isMissingActivityError(error) {
  const message = String(error?.message || error || '');
  return error?.status === 404 || message.includes('此活動尚未由管理者註冊');
}

function renderRecentEvents() {
  const el = document.getElementById('recent-events');
  if (!el) return;
  const list = getRecentEvents().filter(e => e.id !== state.sheetId).slice(0, RECENT_EVENT_LIMIT);
  if (!list.length) { el.innerHTML = ''; return; }
  el.innerHTML = '<p class="label" style="margin-top:6px">或選擇最近的活動：</p>' +
    list.map(e => `
      <button class="recent-event-row" onclick="switchEvent('${e.id}')">
        <span class="recent-event-title">${esc(e.title || '（未命名活動）')}</span>
        <span class="recent-event-date">${new Date(e.ts).toLocaleDateString('zh-TW')}</span>
      </button>`).join('');
}

async function switchEvent(id) {
  clearActivityWorkflowMessages();
  const previous = {
    sheetId: state.sheetId,
    spreadsheetTitle: state.spreadsheetTitle,
    rows: state.rows,
    rosterReady: state.rosterReady,
    selectedRow: state.selectedRow,
  };
  const cachedEvent = getRecentEvents().find(event => event.id === id);
  setLoading(true);
  try {
    state.sheetId = id;
    state.spreadsheetTitle = null;
    state.rows = [];
    state.rosterReady = false;
    state.selectedRow = null;
    await loadRows();
    state.rosterReady = state.rows.length > 0;
    afterEventSelected();
  } catch (e) {
    state.sheetId = previous.sheetId;
    state.spreadsheetTitle = previous.spreadsheetTitle;
    state.rows = previous.rows;
    state.rosterReady = previous.rosterReady;
    state.selectedRow = previous.selectedRow;
    if (isMissingActivityError(e)) {
      removeRecentEvent(id);
      alert(`無法載入「${cachedEvent?.title || '此活動'}」。活動已不存在或不再可用，已從最近活動移除。`);
    } else {
      alert('無法載入此活動：' + (e.message || e));
    }
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
  navigator.clipboard.writeText(activityShareUrl()).then(() => {
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
    state.rosterReady = state.rows.length > 0;
    syncActivityUrl();
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

async function loadRows({ syncForm = true } = {}) {
  const event = await apiAction('event.get', { sheetId: state.sheetId });
  state.spreadsheetTitle = event.title;
  state.rows = event.rows;
  applyEventMetadata(event, { syncForm });
}

async function refreshData() {
  if (!state.sheetId) {
    if (state.currentScreen === 'screen-list')  renderList();
    if (state.currentScreen === 'screen-admin') renderAdmin();
    return;
  }

  setLoading(true);
  try {
    await loadRows();
    state.rosterReady = state.rows.length > 0;
    if (state.currentScreen === 'screen-list')  renderList();
    if (state.currentScreen === 'screen-admin') renderAdminReport();
  } catch { alert('重新整理失敗，請檢查網路。'); }
  finally  { setLoading(false); }
}

// ── Check-in Logic ────────────────────────────────────────────────────────────

async function checkin(row, manual = false, qrPayload = '') {
  const result = await apiAction('checkin', { sheetId: state.sheetId, id: row.id, manual, ...(qrPayload ? { qrPayload } : {}) });
  if (result.already) {
    const error = new Error('already');
    error.code = 'already';
    error.row = replaceStateRow(result.row);
    throw error;
  }
  return replaceStateRow(result.row);
}

async function undoCheckin(row) {
  return replaceStateRow((await apiAction('undo', { sheetId: state.sheetId, id: row.id })).row);
}

function replaceStateRow(updatedRow) {
  if (!updatedRow?.id) return updatedRow;
  const index = state.rows.findIndex(row => row.id === updatedRow.id);
  if (index >= 0) state.rows[index] = updatedRow;
  else state.rows.push(updatedRow);
  return updatedRow;
}

// ── Auto-refresh (keeps multi-device data in sync) ────────────────────────────

let autoRefreshTimer = null;
let autoRefreshBusy = false;

function startAutoRefresh() {
  if (autoRefreshTimer) return;
  autoRefreshTimer = setInterval(async () => {
    if (!state.sheetId || document.hidden || autoRefreshBusy) return;
    if (state.currentScreen !== 'screen-list' && state.currentScreen !== 'screen-admin') return;
    autoRefreshBusy = true;
    try {
      await loadRows({ syncForm: false });
      if (state.currentScreen === 'screen-list')  renderList();
      if (state.currentScreen === 'screen-admin') renderAdminReport();
    } catch { /* silent — next tick will retry */ }
    finally { autoRefreshBusy = false; }
  }, 5000);
}

// ── Scan Screen ───────────────────────────────────────────────────────────────

async function startScanner() {
  updateClock();
  state.clockTimer = setInterval(updateClock, 1000);

  const container = document.getElementById('qr-reader');
  container.innerHTML = '';
  if (!state.sheetId) {
    const message = state.role === 'admin'
      ? '尚未選擇活動，請到「行政」頁建立或選擇活動。'
      : '尚未載入活動，請使用管理者提供的活動連結。';
    container.innerHTML = `<p class="msg-error" style="padding:20px">${message}</p>`;
    state.scanner = null;
    return;
  }
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
  const parsed = parseClientQrPayload(code);
  if (!parsed) {
    scanFeedback('error');
    flashResult('err', 'QR 格式不正確或已過期', code, null);
    state.cooldownTimer = setTimeout(dismissScanResult, 3000);
    return;
  }
  if (parsed.eventId !== state.sheetId) {
    scanFeedback('error');
    flashResult('err', '此 QR 不屬於目前活動', parsed.id, null);
    state.cooldownTimer = setTimeout(dismissScanResult, 3000);
    return;
  }
  let row    = state.rows.find(r => r.id === parsed.id);

  if (!row) {
    // Might be a person added from another device — reload once and retry
    try {
      await loadRows({ syncForm: false });
      row = state.rows.find(r => r.id === parsed.id);
    } catch {}
  }

  if (!row) {
    scanFeedback('not-found');
    flashResult('not-found', '未找到', parsed.id, null);
  } else if (row.checkinStamp) {
    scanFeedback('already');
    flashResult('already', '已掃過', code, row);
  } else {
    try {
      row = await checkin(row, false, code);
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

function formatEventMeta() {
  const meta = state.eventMeta || {};
  const date = meta.startDate && meta.endDate && meta.startDate !== meta.endDate
    ? `${meta.startDate} 至 ${meta.endDate}`
    : (meta.startDate || meta.endDate || '尚未設定日期');
  const time = meta.startTime || meta.endTime
    ? `${meta.startTime || '—'}${meta.endTime ? `–${meta.endTime}` : ''}`
    : '尚未設定時間';
  return [date, time, meta.location || '尚未設定地點'].join('　｜　');
}

function eventStatus() {
  if (!state.sheetId) return '尚未建立';
  if (state.rows.some(row => row.checkinStamp)) return '報到中';
  if (state.eventMeta?.endDate) {
    const end = Date.parse(`${state.eventMeta.endDate}T23:59:59+08:00`);
    if (Number.isFinite(end) && Date.now() > end) return '已結束';
  }
  if (!state.rosterReady) return '設定中';
  if (!state.batches.length) return '待設定分批';
  return '可準備列印';
}

function renderEventContext() {
  const hasEvent = !!state.sheetId;
  const none = document.getElementById('event-none');
  const info = document.getElementById('event-info');
  if (none) none.style.display = hasEvent ? 'none' : '';
  if (info) info.style.display = hasEvent ? '' : 'none';
  if (!hasEvent) return;
  const title = document.getElementById('report-event-title');
  const meta = document.getElementById('event-meta-summary');
  const status = document.getElementById('event-status');
  if (title) title.textContent = state.spreadsheetTitle || '（未命名活動）';
  if (meta) meta.textContent = formatEventMeta();
  const summary = document.getElementById('report-summary');
  if (summary && !state.rosterReady) summary.textContent = '名單尚未匯入';
  if (status) status.textContent = eventStatus();
  const save = document.getElementById('btn-save-event-settings');
  if (save) save.disabled = !hasEvent;
}

async function saveEventSettings() {
  if (!state.sheetId) return;
  hideError('create-error');
  setCreateStatus('正在儲存活動資料…');
  setLoading(true);
  try {
    const updated = await apiAction('event.update', {
      sheetId: state.sheetId,
      ...eventMetadataFromForm(),
    });
    state.spreadsheetTitle = updated.title;
    state.rows = updated.rows;
    applyEventMetadata(updated);
    setCreateStatus('✓ 活動資料已儲存', 5000);
    renderAdmin();
  } catch (err) {
    setCreateStatus('');
    showError('create-error', '儲存失敗：' + err.message);
  } finally { setLoading(false); }
}

async function deleteCurrentEvent() {
  if (!state.sheetId) return;
  const title = state.spreadsheetTitle || '（未命名活動）';
  const typed = window.prompt(`這會刪除「${title}」及其活動名單。\n若確定，請輸入完整活動名稱：`);
  if (typed === null) return;
  if (typed.trim() !== title.trim()) {
    const msg = document.getElementById('delete-event-msg');
    if (msg) { msg.textContent = '活動名稱不一致，未刪除。'; msg.style.display = ''; }
    return;
  }
  const msg = document.getElementById('delete-event-msg');
  if (msg) msg.style.display = 'none';
  setLoading(true);
  try {
    const driveAccessToken = await ensureDriveAccess();
    await apiAction('event.delete', { sheetId: state.sheetId, driveAccessToken });
    removeRecentEvent(state.sheetId);
    clearCurrentEvent();
    setCreateStatus('活動已刪除；現在可以重新建立活動。', 8000);
  } catch (err) {
    if (msg) { msg.textContent = '刪除失敗：' + err.message; msg.style.display = ''; }
  } finally { setLoading(false); }
}

function syncBatchStateFromForm() {
  const rows = [...document.querySelectorAll('#batch-rows .batch-row')];
  if (!rows.length) return;
  state.batches = rows.map(row => ({
    min: row.querySelector('.batch-min')?.value.trim() || '',
    max: row.querySelector('.batch-max')?.value.trim() || '',
    time: row.querySelector('.batch-time')?.value.trim() || '',
  }));
}

function addBatchRow() {
  syncBatchStateFromForm();
  state.batches.push({ min: '', max: '', time: '' });
  renderBatchEditor();
}

function removeBatchRow(index) {
  syncBatchStateFromForm();
  state.batches.splice(index, 1);
  renderBatchEditor();
}

function renderBatchEditor() {
  const el = document.getElementById('batch-rows');
  if (!el) return;
  const hasEvent = !!state.sheetId;
  const save = document.getElementById('btn-save-batches');
  if (save) save.disabled = !hasEvent;
  if (!hasEvent) {
    el.innerHTML = '<p class="msg-empty">請先建立或切換活動。</p>';
    return;
  }
  const batches = state.batches.length ? state.batches : [{ min: '', max: '', time: '' }];
  el.innerHTML = batches.map((batch, index) => `
    <div class="batch-row">
      <label class="field-label">起始尾數<input class="text-input batch-min" data-index="${index}" value="${esc(batch.min)}" placeholder="001"></label>
      <label class="field-label">結束尾數<input class="text-input batch-max" data-index="${index}" value="${esc(batch.max)}" placeholder="430"></label>
      <label class="field-label">報到時段<input class="text-input batch-time" data-index="${index}" value="${esc(batch.time)}" placeholder="10:00–10:45"></label>
      <button type="button" class="btn-icon-danger" onclick="removeBatchRow(${index})">移除</button>
    </div>`).join('');
}

function readBatchForm() {
  const digits = Number(document.getElementById('batch-digits')?.value || 3);
  const rows = [...document.querySelectorAll('#batch-rows .batch-row')];
  return {
    digits: Number.isInteger(digits) && digits >= 1 && digits <= 12 ? digits : 3,
    batches: rows.map(row => ({
      min: row.querySelector('.batch-min')?.value.trim() || '',
      max: row.querySelector('.batch-max')?.value.trim() || '',
      time: row.querySelector('.batch-time')?.value.trim() || '',
    })).filter(batch => batch.min || batch.max || batch.time),
  };
}

async function saveBatches() {
  if (!state.sheetId) return;
  hideError('batch-error');
  const config = readBatchForm();
  const status = document.getElementById('batch-status');
  if (status) status.style.display = 'none';
  setLoading(true);
  try {
    const updated = await apiAction('event.update', {
      sheetId: state.sheetId,
      ...eventMetadataFromForm(),
      batches: config.batches,
      batchDigits: config.digits,
    });
    state.rows = updated.rows;
    state.spreadsheetTitle = updated.title;
    applyEventMetadata(updated);
    if (status) { status.textContent = '✓ 分批時段已儲存並套用到名單。'; status.style.display = ''; }
    renderAdmin();
  } catch (err) {
    const detail = Array.isArray(err.details) ? `\n${err.details.join('\n')}` : '';
    showError('batch-error', '儲存失敗：' + err.message + detail);
  } finally { setLoading(false); }
}

function qrPayloadFor(id) {
  return `qr-checkin:v2:${encodeURIComponent(state.sheetId || '')}:${encodeURIComponent(String(id || ''))}`;
}

function parseClientQrPayload(raw) {
  const prefix = 'qr-checkin:v2:';
  const code = String(raw || '').trim();
  if (!code.startsWith(prefix)) return null;
  const rest = code.slice(prefix.length);
  const separator = rest.indexOf(':');
  if (separator < 1) return null;
  try {
    return {
      eventId: decodeURIComponent(rest.slice(0, separator)),
      id: decodeURIComponent(rest.slice(separator + 1)),
    };
  } catch { return null; }
}

function renderNoticeResults() {
  const el = document.getElementById('notice-results');
  if (!el) return;
  if (!state.rosterReady) { el.innerHTML = '<p class="msg-empty">名單匯入完成後才能查找參與者。</p>'; return; }
  const q = (document.getElementById('notice-search')?.value || '').trim().toLowerCase();
  const rows = (q ? state.rows.filter(row => String(row.id || '').toLowerCase().includes(q) || String(row.name || '').toLowerCase().includes(q)) : []).slice(0, 12);
  if (!q) { el.innerHTML = '<p class="label">請輸入完整編號或姓名；編號是主要查找鍵。</p>'; return; }
  if (!rows.length) { el.innerHTML = '<p class="msg-empty">找不到符合的參與者。</p>'; return; }
  el.innerHTML = rows.map(row => `
    <button type="button" class="notice-result-row${state.noticeRow?.rowIndex === row.rowIndex ? ' selected' : ''}" onclick="selectNoticeRow(${row.rowIndex})">
      <span class="notice-result-main"><span class="notice-result-id">${esc(row.id)}</span><span class="notice-result-name">${esc(row.name)}</span><span class="notice-result-address">${esc(row.address || '（無地址）')}</span></span>
      <span class="notice-result-time">${esc(row.batchTime || '未設定時段')}</span>
    </button>`).join('');
}

function selectNoticeRow(rowIndex) {
  const row = state.rows.find(item => item.rowIndex === rowIndex);
  if (!row) return;
  state.noticeRow = row;
  const selected = document.getElementById('notice-selected');
  if (selected) selected.style.display = '';
  const name = document.getElementById('notice-selected-name');
  const id = document.getElementById('notice-selected-id');
  const address = document.getElementById('notice-address');
  if (name) name.textContent = row.name;
  if (id) id.textContent = row.id;
  if (address) address.value = row.address || '';
  renderNoticePreview();
  renderNoticeResults();
}

function renderNoticePreview() {
  const el = document.getElementById('notice-preview');
  if (!el) return;
  const row = state.noticeRow;
  if (!row) { el.innerHTML = '<p class="msg-empty">查找並選擇參與者後，這裡會顯示通知單預覽。</p>'; return; }
  let svg = '';
  try { const qr = qrcode(0, 'M'); qr.addData(qrPayloadFor(row.id)); qr.make(); svg = qr.createSvgTag({ cellSize: 3, margin: 1 }); } catch {}
  const meta = [row.address || '（無地址）', row.batchTime || '尚未設定報到時段', state.eventMeta?.location || '（未設定地點）'].join('\n');
  el.innerHTML = `<div class="notice-preview-card"><div><div class="notice-preview-title">${esc(state.spreadsheetTitle || '活動通知單')}</div><div class="notice-preview-meta">${esc(row.id)}　${esc(row.name)}\n${esc(meta)}</div></div><div>${svg}</div></div>`;
}

async function saveNoticeAddress() {
  if (!state.sheetId || !state.noticeRow) return;
  const address = document.getElementById('notice-address')?.value.trim() || '';
  hideError('notice-error');
  setLoading(true);
  try {
    const updated = await apiAction('attendee.update', { sheetId: state.sheetId, id: state.noticeRow.id, address });
    state.rows = updated.rows;
    state.spreadsheetTitle = updated.title;
    applyEventMetadata(updated);
    state.noticeRow = state.rows.find(row => row.id === state.noticeRow.id) || null;
    renderNoticeResults();
    selectNoticeRow(state.noticeRow?.rowIndex);
  } catch (err) {
    showError('notice-error', '地址儲存失敗：' + err.message);
  } finally { setLoading(false); }
}

function noticeRowsForOutput(rows) {
  const selected = (rows || []).filter(row => row && row.id && !/^W-/.test(String(row.id)));
  if (!selected.length) throw new Error('目前沒有可產出的參與者。');
  ensureIds();
  const missingTimes = selected.filter(row => !row.batchTime);
  if (missingTimes.length) {
    throw new Error(`尚有 ${missingTimes.length} 位參與者未分配報到時段，請先完成分批設定。`);
  }
  return selected;
}

function noticeDocumentHtml(rows) {
  const meta = state.eventMeta || {};
  const date = meta.startDate && meta.endDate && meta.startDate !== meta.endDate
    ? `${meta.startDate} 至 ${meta.endDate}` : (meta.startDate || meta.endDate || '');
  const time = meta.startTime || meta.endTime ? `${meta.startTime || ''}${meta.endTime ? `–${meta.endTime}` : ''}` : '';
  const pages = rows.map(row => {
    let qr = '';
    try { const code = qrcode(0, 'M'); code.addData(qrPayloadFor(row.id)); code.make(); qr = code.createSvgTag({ cellSize: 4, margin: 1 }); } catch {}
    return `<section class="notice-page"><div class="notice-copy"><h1>${esc(state.spreadsheetTitle || '活動通知單')}</h1><p class="notice-date">${esc(date)}${time ? `　${esc(time)}` : ''}</p><p><b>編號：</b>${esc(row.id)}</p><p><b>姓名：</b>${esc(row.name)}</p><p><b>通訊地址：</b>${esc(row.address || '（無地址）')}</p><p><b>報到時段：</b>${esc(row.batchTime)}</p><p><b>活動地點：</b>${esc(meta.location || '（未設定）')}</p><div class="notice-note">${esc(meta.notice || '請依通知單所列時段報到，並攜帶本通知單。')}</div></div><div class="notice-qr">${qr}<div>報到 QR Code</div><small>${esc(row.id)}</small></div></section>`;
  }).join('');
  return `<!doctype html><html lang="zh-TW"><head><meta charset="utf-8"><title>${esc(state.spreadsheetTitle || '活動通知單')}</title><style>
  @page{size:A4;margin:0}*{box-sizing:border-box}body{margin:0;background:#eee;font-family:"Microsoft JhengHei","Noto Sans CJK TC",sans-serif;color:#222}.notice-page{width:210mm;min-height:297mm;padding:24mm 22mm;display:flex;justify-content:space-between;gap:16mm;page-break-after:always;background:#fff}.notice-page:last-child{page-break-after:auto}.notice-copy{flex:1}.notice-copy h1{text-align:center;font-size:24px;margin:0 0 10mm}.notice-date{text-align:center;font-size:15px;margin-bottom:15mm}.notice-copy p{font-size:17px;line-height:1.75;margin:5mm 0}.notice-note{margin-top:18mm;border-top:1px solid #888;padding-top:8mm;font-size:15px;line-height:1.8;white-space:pre-line}.notice-qr{width:48mm;text-align:center;font-size:12px;padding-top:28mm}.notice-qr svg{width:42mm;height:42mm}.notice-qr small{display:block;margin-top:3mm;color:#666}@media screen{.notice-page{margin:12px auto;box-shadow:0 1px 5px #aaa}}@media print{body{background:#fff}.notice-page{box-shadow:none;margin:0}}
  </style></head><body>${pages}</body></html>`;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

async function downloadNoticeWord(rows) {
  const html = noticeDocumentHtml(rows);
  downloadBlob(new Blob(['\ufeff', html], { type: 'application/msword;charset=utf-8' }), `${state.spreadsheetTitle || '活動'}_通知單_已合併.doc`);
}

function wrapCanvasText(ctx, text, maxWidth) {
  const lines = [];
  String(text || '').split('\n').forEach(segment => {
    let line = '';
    [...segment].forEach(char => {
      const candidate = line + char;
      if (line && ctx.measureText(candidate).width > maxWidth) { lines.push(line); line = char; }
      else line = candidate;
    });
    lines.push(line);
  });
  return lines;
}

function qrSvgDataUrl(id) {
  const qr = qrcode(0, 'M');
  qr.addData(qrPayloadFor(id)); qr.make();
  return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(qr.createSvgTag({ cellSize: 5, margin: 1 }))}`;
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image); image.onerror = reject; image.src = src;
  });
}

async function renderNoticeJpeg(row) {
  const canvas = document.createElement('canvas');
  canvas.width = 794; canvas.height = 1123;
  const ctx = canvas.getContext('2d');
  ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = '#222'; ctx.textAlign = 'left'; ctx.font = 'bold 28px "Microsoft JhengHei", sans-serif';
  ctx.fillText(state.spreadsheetTitle || '活動通知單', 100, 135);
  ctx.font = '18px "Microsoft JhengHei", sans-serif';
  const meta = state.eventMeta || {};
  const date = meta.startDate && meta.endDate && meta.startDate !== meta.endDate ? `${meta.startDate} 至 ${meta.endDate}` : (meta.startDate || meta.endDate || '');
  const time = meta.startTime || meta.endTime ? `${meta.startTime || ''}${meta.endTime ? `–${meta.endTime}` : ''}` : '';
  ctx.fillText(`${date}${time ? `　${time}` : ''}`, 100, 175);
  ctx.font = '22px "Microsoft JhengHei", sans-serif';
  const lines = [
    `編號：${row.id}`,
    `姓名：${row.name}`,
    `通訊地址：${row.address || '（無地址）'}`,
    `報到時段：${row.batchTime}`,
    `活動地點：${meta.location || '（未設定）'}`,
  ];
  let y = 280;
  lines.forEach(line => { wrapCanvasText(ctx, line, 470).forEach(part => { ctx.fillText(part, 100, y); y += 42; }); y += 6; });
  ctx.strokeStyle = '#888'; ctx.beginPath(); ctx.moveTo(100, y + 35); ctx.lineTo(694, y + 35); ctx.stroke();
  ctx.font = '18px "Microsoft JhengHei", sans-serif';
  y += 85;
  wrapCanvasText(ctx, meta.notice || '請依通知單所列時段報到，並攜帶本通知單。', 594).forEach(line => { ctx.fillText(line, 100, y); y += 32; });
  const qr = await loadImage(qrSvgDataUrl(row.id));
  ctx.drawImage(qr, 520, 245, 205, 205);
  ctx.font = '14px "Microsoft JhengHei", sans-serif'; ctx.textAlign = 'center'; ctx.fillText('報到 QR Code', 622, 475); ctx.fillText(row.id, 622, 498);
  return canvas.toDataURL('image/jpeg', 0.92);
}

function base64Bytes(dataUrl) {
  const binary = atob(dataUrl.split(',')[1]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

function asciiBytes(value) { return new TextEncoder().encode(value); }

function buildImagePdf(jpegs) {
  const pageWidth = 595; const pageHeight = 842; const imageWidth = 794; const imageHeight = 1123;
  const pageObjects = jpegs.map((_, index) => ({ page: 3 + index * 3, content: 4 + index * 3, image: 5 + index * 3 }));
  const objects = [];
  objects[1] = asciiBytes('<< /Type /Catalog /Pages 2 0 R >>');
  objects[2] = asciiBytes(`<< /Type /Pages /Kids [${pageObjects.map(item => `${item.page} 0 R`).join(' ')}] /Count ${pageObjects.length} >>`);
  pageObjects.forEach((item, index) => {
    objects[item.page] = asciiBytes(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /XObject << /Im0 ${item.image} 0 R >> >> /Contents ${item.content} 0 R >>`);
    const content = `q\n${pageWidth} 0 0 ${pageHeight} 0 0 cm\n/Im0 Do\nQ\n`;
    objects[item.content] = asciiBytes(`<< /Length ${asciiBytes(content).length} >>\nstream\n${content}endstream`);
    const bytes = base64Bytes(jpegs[index]);
    objects[item.image] = { bytes, header: `<< /Type /XObject /Subtype /Image /Width ${imageWidth} /Height ${imageHeight} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${bytes.length} >>\nstream\n`, footer: '\nendstream' };
  });
  const chunks = [asciiBytes('%PDF-1.4\n%\xff\xff\xff\xff\n')];
  const offsets = [0]; let length = chunks[0].length;
  for (let index = 1; index < objects.length; index += 1) {
    offsets[index] = length;
    const object = objects[index];
    const start = asciiBytes(`${index} 0 obj\n`);
    const end = asciiBytes('\nendobj\n');
    chunks.push(start); length += start.length;
    if (object.bytes) { const header = asciiBytes(object.header); const footer = asciiBytes(object.footer); chunks.push(header, object.bytes, footer); length += header.length + object.bytes.length + footer.length; }
    else { chunks.push(object); length += object.length; }
    chunks.push(end); length += end.length;
  }
  const xrefOffset = length;
  const xref = [`xref\n0 ${objects.length}\n0000000000 65535 f \n`];
  for (let index = 1; index < objects.length; index += 1) xref.push(`${String(offsets[index]).padStart(10, '0')} 00000 n \n`);
  xref.push(`trailer\n<< /Size ${objects.length} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`);
  chunks.push(asciiBytes(xref.join('')));
  return new Blob(chunks, { type: 'application/pdf' });
}

async function downloadNoticePdf(rows) {
  const jpegs = [];
  for (const row of rows) jpegs.push(await renderNoticeJpeg(row));
  downloadBlob(buildImagePdf(jpegs), `${state.spreadsheetTitle || '活動'}_通知單_已合併.pdf`);
}

async function generateNoticeDocuments(rows) {
  hideError('notice-error');
  try {
    const outputRows = noticeRowsForOutput(rows);
    setLoading(true);
    const templateId = state.eventMeta?.templateId || '';
    if (templateId && templateById(templateId)) {
      await downloadNoticeFromTemplate(outputRows, templateId);
      return;
    }
    await downloadNoticeWord(outputRows);
    await downloadNoticePdf(outputRows);
  } catch (err) {
    showError('notice-error', '產出失敗：' + err.message);
  } finally { setLoading(false); }
}

function generateAllNotices() { generateNoticeDocuments(state.rows); }
function generateSelectedNotice() { if (!state.noticeRow) { showError('notice-error', '請先查找並選擇參與者。'); return; } generateNoticeDocuments([state.noticeRow]); }

// ── Notice Templates (活動通知範本) ─────────────────────────────────────────────

const ROC_WEEKDAYS = '日一二三四五六';

function rocDate(isoDate) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(isoDate || ''));
  if (!match) return '';
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return `${Number(match[1]) - 1911}年${match[2]}月${match[3]}日星期${ROC_WEEKDAYS[date.getDay()]}`;
}

function chineseClock(time) {
  const match = /^(\d{1,2}):(\d{2})/.exec(String(time || ''));
  if (!match) return { period: '', text: '' };
  const hour = Number(match[1]);
  const period = hour < 12 ? '上午' : hour === 12 ? '中午' : '下午';
  const shown = hour > 12 ? hour - 12 : hour;
  return { period, text: `${shown}：${match[2]}` };
}

function eventTimeText(meta) {
  const start = chineseClock(meta.startTime);
  const end = chineseClock(meta.endTime);
  if (!start.text && !end.text) return '';
  if (!end.text) return `${start.period}${start.text}`;
  if (!start.text) return `${end.period}${end.text}`;
  return `${start.period}${start.text}～${end.period === start.period ? '' : end.period}${end.text}`;
}

function eventDateText(meta) {
  const start = rocDate(meta.startDate);
  const end = rocDate(meta.endDate);
  if (start && end && meta.endDate !== meta.startDate) return `${start}至${end}`;
  return start || end;
}

function noticeTemplateValues(row) {
  const meta = state.eventMeta || {};
  return {
    編號: row.id,
    姓名: row.name,
    通訊地址: row.address || '',
    報到時段: row.batchTime || '',
    報到日期: rocDate(meta.startDate),
    活動名稱: state.spreadsheetTitle || '',
    活動日期: eventDateText(meta),
    活動時間: eventTimeText(meta),
    活動說明: meta.notice || '',
    活動地點: meta.location || '',
    發放地點: meta.location || '',
  };
}

function qrGifBytes(text) {
  const qr = qrcode(0, 'M');
  qr.addData(text); qr.make();
  return { bytes: base64Bytes(qr.createDataURL(10, 2)), ext: 'gif' };
}

function bytesToBase64(bytes) {
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

function templateById(id) {
  return state.templates.find(item => item.id === id) || null;
}

async function loadTemplates() {
  try {
    const result = await apiAction('template.list');
    state.templates = result.templates || [];
    state.templatesError = '';
  } catch (err) {
    state.templates = [];
    state.templatesError = err.message;
  }
  renderTemplateSelect();
  renderTemplateManager();
  renderNoticeTemplateInfo();
}

function renderTemplateSelect() {
  const select = document.getElementById('input-event-template');
  if (!select) return;
  const current = state.eventMeta?.templateId || '';
  const options = ['<option value="">系統預設版面（不使用範本）</option>']
    .concat(state.templates.map(item => `<option value="${esc(item.id)}">${esc(item.name)}</option>`));
  if (current && !templateById(current)) options.push(`<option value="${esc(current)}">（範本已刪除）</option>`);
  select.innerHTML = options.join('');
  select.value = current;
}

// Choosing a venue template fills in the wording the venue used last time.
function onEventTemplateChange() {
  const select = document.getElementById('input-event-template');
  const template = templateById(select?.value || '');
  const hint = document.getElementById('event-template-hint');
  if (!template) {
    if (hint) hint.textContent = '未使用範本時，通知單使用系統預設版面。';
    return;
  }
  const source = template.lastUsed || template.defaults || {};
  const set = (id, value) => { const el = document.getElementById(id); if (el) el.value = value || ''; };
  set('input-event-location', source.location);
  set('input-event-notice', source.notice);
  set('input-event-nav-url', source.navUrl);
  if (hint) {
    hint.textContent = template.lastUsed
      ? `已帶入「${template.lastUsed.title || template.name}」使用的說明、地點與導航連結，可再修改；記得儲存活動資料。`
      : `已帶入「${template.name}」範本的預設文字，可再修改；記得儲存活動資料。`;
  }
}

function renderNoticeTemplateInfo() {
  const el = document.getElementById('notice-template-info');
  const button = document.getElementById('btn-generate-all-notices');
  const template = templateById(state.eventMeta?.templateId || '');
  if (el) {
    el.textContent = !state.sheetId ? ''
      : template ? `通知單將套用「${template.name}」範本，產出一份合併的 Word 檔（需要 PDF 時，可用 Word 另存新檔）。`
      : state.eventMeta?.templateId ? '此活動選用的範本已被刪除，將改用系統預設版面。'
      : '此活動未選用通知範本，將使用系統預設版面。可在「1. 活動設定」選擇範本。';
  }
  if (button) button.textContent = template ? '產生全部通知單（Word）' : '產生全部 Word／PDF';
}

function renderTemplateManager() {
  const el = document.getElementById('template-list');
  if (!el) return;
  if (state.templatesError) { el.innerHTML = `<p class="msg-error">無法讀取範本：${esc(state.templatesError)}</p>`; return; }
  if (!state.templates.length) { el.innerHTML = '<p class="msg-empty">尚未建立任何範本。</p>'; return; }
  el.innerHTML = `<div class="template-list">${state.templates.map(item => `
    <div class="template-row">
      <div class="template-row-main">
        <b>${esc(item.name)}</b>
        <span class="template-row-meta">${esc(item.fileName || '')}${item.size ? `・${Math.ceil(item.size / 1024)} KB` : ''}・更新：${esc(item.updatedAt || '')}</span>
        ${item.lastUsed ? `<span class="template-row-meta">最近使用：${esc(item.lastUsed.title || '')}</span>` : ''}
      </div>
      <div class="template-row-actions">
        <button type="button" class="btn btn-secondary btn-small" onclick="editTemplate('${esc(item.id)}')">編輯</button>
        <button type="button" class="btn btn-secondary btn-small" onclick="downloadTemplateFile('${esc(item.id)}')">下載</button>
        <button type="button" class="btn-icon-danger" onclick="deleteTemplate('${esc(item.id)}')">刪除</button>
      </div>
    </div>`).join('')}</div>`;
}

function resetTemplateForm() {
  ['template-id', 'template-name', 'template-location', 'template-notice', 'template-nav-url', 'template-file'].forEach(id => {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  const title = document.getElementById('template-form-title');
  if (title) title.textContent = '新增範本';
  const hint = document.getElementById('template-file-hint');
  if (hint) hint.textContent = '';
  hideError('template-error');
}

function editTemplate(id) {
  const item = templateById(id);
  if (!item) return;
  resetTemplateForm();
  const set = (field, value) => { const el = document.getElementById(field); if (el) el.value = value || ''; };
  set('template-id', item.id);
  set('template-name', item.name);
  set('template-location', item.defaults?.location);
  set('template-notice', item.defaults?.notice);
  set('template-nav-url', item.defaults?.navUrl);
  const title = document.getElementById('template-form-title');
  if (title) title.textContent = `編輯範本：${item.name}`;
  const hint = document.getElementById('template-file-hint');
  if (hint) hint.textContent = `目前檔案：${item.fileName || '（未命名）'}。不選新檔案則保留原檔。`;
  document.getElementById('template-form')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function showTemplateStatus(message) {
  const el = document.getElementById('template-status');
  if (!el) return;
  el.textContent = message;
  el.style.display = message ? '' : 'none';
  if (message) setTimeout(() => { if (el.textContent === message) el.style.display = 'none'; }, 8000);
}

async function saveTemplateForm() {
  hideError('template-error');
  showTemplateStatus('');
  const value = id => document.getElementById(id)?.value || '';
  const id = value('template-id');
  const name = value('template-name').trim();
  const file = document.getElementById('template-file')?.files?.[0];
  if (!name) { showError('template-error', '請輸入範本名稱。'); return; }
  if (!id && !file) { showError('template-error', '請選擇 Word 範本檔（.docx）。'); return; }
  const payload = {
    id, name,
    location: value('template-location').trim(),
    notice: value('template-notice'),
    navUrl: value('template-nav-url').trim(),
  };
  setLoading(true);
  try {
    let warnings = [];
    if (file) {
      if (!/\.docx$/i.test(file.name)) throw new Error('請選擇 Word（.docx）檔；舊版 .doc 請先在 Word 另存為 .docx。');
      const bytes = new Uint8Array(await file.arrayBuffer());
      const info = await NoticeTemplate.inspectTemplate(JSZip, bytes);
      if (!info.images.includes('報到QR')) {
        throw new Error('範本中找不到報到 QR 圖片。請在 Word 中對 QR 圖片按右鍵 →「檢視替代文字」，描述填入 {{報到QR}}。');
      }
      if (!info.text.includes('編號') && !info.text.includes('姓名')) {
        throw new Error('範本中找不到 {{編號}} 或 {{姓名}} 標記，請確認範本已加入標記。');
      }
      const known = ['編號', '姓名', '通訊地址', '報到時段', '報到日期', '活動名稱', '活動日期', '活動時間', '活動說明', '活動地點', '發放地點'];
      const unknown = info.text.filter(tag => !known.includes(tag));
      if (unknown.length) warnings.push(`無法辨識的標記會原樣保留：${unknown.map(tag => `{{${tag}}}`).join('、')}`);
      if (info.hasMultiplePages) warnings.push('範本中有分頁符號；系統會以整份範本為每人一份，請確認範本只有一頁。');
      payload.fileName = file.name;
      payload.data = bytesToBase64(bytes);
    }
    const result = await apiAction('template.save', payload);
    state.templates = result.templates || [];
    delete state.templateFiles[id];
    resetTemplateForm();
    renderTemplateSelect();
    renderTemplateManager();
    renderNoticeTemplateInfo();
    showTemplateStatus(`✓ 已儲存範本「${name}」。${warnings.length ? `\n${warnings.join('\n')}` : ''}`);
  } catch (err) {
    showError('template-error', '儲存失敗：' + err.message);
  } finally { setLoading(false); }
}

async function fetchTemplateBytes(id) {
  const item = templateById(id);
  const cached = state.templateFiles[id];
  if (cached && item && cached.updatedAt === item.updatedAt) return cached.bytes;
  const result = await apiAction('template.get', { id });
  const bytes = base64Bytes(`data:;base64,${result.data}`);
  state.templateFiles[id] = { updatedAt: result.updatedAt, bytes, fileName: result.fileName };
  return bytes;
}

async function downloadTemplateFile(id) {
  const item = templateById(id);
  setLoading(true);
  try {
    const bytes = await fetchTemplateBytes(id);
    downloadBlob(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), item?.fileName || `${item?.name || '通知範本'}.docx`);
  } catch (err) {
    showError('template-error', '下載失敗：' + err.message);
  } finally { setLoading(false); }
}

async function deleteTemplate(id) {
  const item = templateById(id);
  if (!item) return;
  if (!confirm(`確定要刪除範本「${item.name}」？已使用此範本的活動將改用系統預設版面。`)) return;
  setLoading(true);
  try {
    const result = await apiAction('template.delete', { id });
    state.templates = result.templates || [];
    delete state.templateFiles[id];
    renderTemplateSelect();
    renderTemplateManager();
    renderNoticeTemplateInfo();
    showTemplateStatus(`已刪除範本「${item.name}」。`);
  } catch (err) {
    showError('template-error', '刪除失敗：' + err.message);
  } finally { setLoading(false); }
}

async function downloadNoticeFromTemplate(rows, templateId) {
  const template = templateById(templateId);
  const bytes = await fetchTemplateBytes(templateId);
  const pages = rows.map(row => ({
    values: noticeTemplateValues(row),
    images: { 報到QR: qrGifBytes(qrPayloadFor(row.id)) },
  }));
  const navUrl = (state.eventMeta?.navUrl || '').trim();
  const zip = await NoticeTemplate.mergeTemplate(JSZip, bytes, pages, {
    sharedImages: navUrl ? { 導航QR: qrGifBytes(navUrl) } : {},
  });
  const blob = await zip.generateAsync({
    type: 'blob', compression: 'DEFLATE',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const suffix = rows.length === 1 ? `_${rows[0].id}` : '';
  downloadBlob(blob, `${state.spreadsheetTitle || '活動'}_通知單_${template?.name || '範本'}${suffix}.docx`);
}

function renderAdmin() {
  renderEventContext();
  renderAdminReport();
  renderBatchEditor();
  renderNoticeResults();
  renderNoticePreview();
  renderAllowListSection(); // async, self-managing
  renderNoticeTemplateInfo();
  if (!state.templatesLoaded) { state.templatesLoaded = true; loadTemplates(); }
  const title = document.getElementById('dashboard-title');
  const subtitle = document.getElementById('dashboard-subtitle');
  if (title) title.textContent = state.sheetId && state.spreadsheetTitle
    ? state.spreadsheetTitle
    : '建立並管理您的活動';
  if (subtitle) subtitle.textContent = canUploadRoster()
    ? '活動已建立，請完成名單匯入；接著設定分批、產出通知單並分享給志工。'
    : state.sheetId && !state.rosterReady
    ? '目前活動尚無名單；為避免誤寫，請建立新活動後再匯入名單。'
    : state.sheetId
    ? '依序完成分批、通知單與分享；活動當天可直接前往掃描報到。'
    : '先建立活動，再匯入名單、設定報到時段，即可產出通知單。';
  updateRosterUploadControls();
}

// ── Report (within Admin) ─────────────────────────────────────────────────────

function renderAdminReport() {
  const hasEvent = !!state.sheetId;
  const hasRoster = hasEvent && state.rosterReady;

  document.getElementById('event-none').style.display = hasEvent ? 'none' : '';
  document.getElementById('event-info').style.display = hasEvent ? '' : 'none';
  document.getElementById('share-box').style.display  = hasRoster ? '' : 'none';
  const shareEmpty = document.getElementById('share-empty');
  if (shareEmpty) shareEmpty.style.display = hasRoster ? 'none' : '';
  renderRecentEvents();

  const statsEl = document.getElementById('report-stats');

  if (!hasRoster) {
    statsEl.innerHTML = `<p class="msg-empty" style="padding:8px 0">${hasEvent ? '名單尚未匯入' : '尚未選擇活動'}</p>`;
    return;
  }

  const total   = state.rows.length;
  const checked = state.rows.filter(r => r.checkinStamp).length;
  const title   = state.spreadsheetTitle || '';

  document.getElementById('report-event-title').textContent = title || '（未命名活動）';
  document.getElementById('report-summary').textContent =
    `總數 ${total}　已報到 ${checked}　未報到 ${total - checked}`;
  const shareUrl = activityShareUrl();
  document.getElementById('share-url').textContent = shareUrl;
  renderShareQr(shareUrl);

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

function renderShareQr(shareUrl = activityShareUrl()) {
  const el = document.getElementById('share-qr');
  if (!el) return;
  el.innerHTML = '';
  try {
    const qr = qrcode(0, 'M');
    qr.addData(shareUrl);
    qr.make();
    el.innerHTML = qr.createSvgTag({ cellSize: 4, margin: 2 });
  } catch {}
}

// ── Attendee QR codes ─────────────────────────────────────────────────────────

// A participant number is assigned by the event organizer and must never change.
function ensureIds() {
  const used = new Set(state.rows.map(row => row.id).filter(Boolean));
  const missing = state.rows.filter(row => row.name && !row.id).map(row => row.rowIndex);
  if (missing.length) throw new Error(`第 ${missing.join('、')} 列缺少編號；編號必須由活動方編入且不可自動產生。`);
  if (used.size !== state.rows.length) {
    throw new Error('名單中有重複編號，請修正後再列印。');
  }
}

async function generateQrCards() {
  hideError('qrgen-error');
  if (!state.sheetId || !state.rosterReady) {
    showError('qrgen-error', '名單尚未上傳。請下載活動名單範本，填妥後由「建立或切換活動」區塊上傳。');
    return;
  }

  setLoading(true);
  try {
    await ensureIds();
    await loadRows();
    if (!state.rows.length) {
      showError('qrgen-error', '名單是空的。請下載活動名單範本，填妥後由「建立或切換活動」區塊上傳。');
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
      q.addData(qrPayloadFor(r.id));
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

function reportExportRows() {
  return [
    EVENT_HEADERS,
    ...state.rows.map(r => [r.id, r.name, r.address, r.notes, r.checkinStamp, r.checkinBy, r.batchTime || '']),
  ];
}

function downloadReportXlsx() {
  if (!state.sheetId) { alert('請先建立或選擇活動。'); return; }
  if (typeof XLSX === 'undefined') { alert('匯出工具尚未載入，請重新整理後再試。'); return; }

  const rows = reportExportRows();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet['!cols'] = [
    { wch: 16 },
    { wch: 16 },
    { wch: 34 },
    { wch: 26 },
    { wch: 22 },
    { wch: 24 },
    { wch: 22 },
  ];

  // Keep identifiers, timestamps, and user-entered values as text so Excel
  // does not remove leading zeroes or reinterpret them using local settings.
  for (let row = 0; row < rows.length; row++) {
    for (let col = 0; col < EVENT_HEADERS.length; col++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: row, c: col })];
      if (cell) cell.z = '@';
    }
  }

  const book = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(book, sheet, '出席報告');
  XLSX.writeFile(book, `${state.spreadsheetTitle || '出席報告'}_${dateString()}.xlsx`);
}

function downloadReport() {
  if (!state.sheetId) { alert('請先建立或選擇活動。'); return; }
  const csvRows = reportExportRows();
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
    state.rosterReady = true;

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
    const active = b.dataset.screen === screenId;
    b.classList.toggle('active', active);
    if (active) b.setAttribute('aria-current', 'page');
    else b.removeAttribute('aria-current');
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
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hour12: false,
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.year}/${values.month}/${values.day} ` +
    `${values.hour}:${values.minute}:${values.second}`;
}

function dateString() {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
  }).formatToParts(new Date());
  const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
  return `${values.year}${values.month}${values.day}`;
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

let importStatusTimer = null;

function setImportStatus(message, autoDismissMs = 0) {
  const el = document.getElementById('import-status');
  if (!el) return;
  if (importStatusTimer) clearTimeout(importStatusTimer);
  importStatusTimer = null;
  el.textContent = message;
  el.style.display = message ? '' : 'none';
  if (message && autoDismissMs > 0) {
    importStatusTimer = setTimeout(() => {
      if (el.textContent === message) {
        el.textContent = '';
        el.style.display = 'none';
      }
      importStatusTimer = null;
    }, autoDismissMs);
  }
}

function hideError(id) {
  const el = document.getElementById(id);
  if (el) el.style.display = 'none';
}
