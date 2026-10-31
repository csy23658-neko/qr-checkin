let cachedToken = { value: '', expiresAt: 0 };

const encode = value => new TextEncoder().encode(value);
const asJson = (body, status = 200, headers = {}) => new Response(JSON.stringify(body), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', ...headers },
});
const TAIPEI_TIME_ZONE = 'Asia/Taipei';
const taipeiTimestamp = () => new Date().toLocaleString('zh-TW', {
  hour12: false,
  timeZone: TAIPEI_TIME_ZONE,
});

const EVENT_REGISTRY_COLUMNS = 11;
const DEFAULT_BATCH_DIGITS = 3;

function text(value) { return String(value ?? '').trim(); }

function parseBatches(value) {
  if (!value) return [];
  try {
    const parsed = typeof value === 'string' ? JSON.parse(value) : value;
    return Array.isArray(parsed) ? parsed.map(batch => ({
      min: text(batch.min),
      max: text(batch.max),
      time: text(batch.time),
    })) : [];
  } catch { return []; }
}

function eventMetaFromPayload(payload = {}, existing = {}) {
  const digits = Number(payload.batchDigits ?? existing.batchDigits ?? DEFAULT_BATCH_DIGITS);
  return {
    startDate: text(payload.startDate ?? existing.startDate),
    endDate: text(payload.endDate ?? existing.endDate),
    startTime: text(payload.startTime ?? existing.startTime),
    endTime: text(payload.endTime ?? existing.endTime),
    location: text(payload.location ?? existing.location),
    notice: text(payload.notice ?? existing.notice),
    batchDigits: Number.isInteger(digits) && digits >= 1 && digits <= 12 ? digits : DEFAULT_BATCH_DIGITS,
    batches: payload.batches === undefined ? parseBatches(existing.batches) : parseBatches(payload.batches),
  };
}

function registryRow(entry) {
  return [[
    entry.sheetId,
    entry.title,
    entry.createdAt || new Date().toISOString(),
    entry.startDate || '',
    entry.endDate || '',
    entry.startTime || '',
    entry.endTime || '',
    entry.location || '',
    entry.notice || '',
    JSON.stringify(entry.batches || []),
    String(entry.batchDigits || DEFAULT_BATCH_DIGITS),
  ]];
}

function registryEntry(row, rowIndex) {
  if (!row || !text(row[0])) return null;
  return {
    rowIndex,
    sheetId: text(row[0]),
    title: text(row[1]),
    createdAt: text(row[2]),
    startDate: text(row[3]),
    endDate: text(row[4]),
    startTime: text(row[5]),
    endTime: text(row[6]),
    location: text(row[7]),
    notice: text(row[8]),
    batches: parseBatches(row[9]),
    batchDigits: Number(row[10]) || DEFAULT_BATCH_DIGITS,
  };
}

function allowedCors(request, env) {
  const origin = request.headers.get('Origin') || '';
  const origins = (env.ALLOWED_ORIGINS || 'https://qr-checkin-3iv.pages.dev').split(',').map(value => value.trim());
  return origins.includes(origin) ? { 'access-control-allow-origin': origin, vary: 'Origin' } : {};
}

function base64Url(bytes) {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');
}

function pemToBytes(pem) {
  const binary = atob(pem.replace(/-----(BEGIN|END) PRIVATE KEY-----/g, '').replace(/\s/g, ''));
  return Uint8Array.from(binary, char => char.charCodeAt(0));
}

async function serviceToken(env) {
  if (cachedToken.expiresAt > Date.now() + 60_000) return cachedToken.value;
  const service = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  const now = Math.floor(Date.now() / 1000);
  const header = base64Url(encode(JSON.stringify({ alg: 'RS256', typ: 'JWT' })));
  const claim = base64Url(encode(JSON.stringify({
    iss: service.client_email,
    scope: 'https://www.googleapis.com/auth/spreadsheets https://www.googleapis.com/auth/drive.file',
    aud: 'https://oauth2.googleapis.com/token', iat: now, exp: now + 3600,
  })));
  const key = await crypto.subtle.importKey('pkcs8', pemToBytes(service.private_key), { name: 'RSASSA-PKCS1-v1_5', hash: 'SHA-256' }, false, ['sign']);
  const signature = await crypto.subtle.sign('RSASSA-PKCS1-v1_5', key, encode(`${header}.${claim}`));
  const assertion = `${header}.${claim}.${base64Url(new Uint8Array(signature))}`;
  const response = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST', headers: { 'content-type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer', assertion }),
  });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error_description || 'Unable to obtain Google service token.');
  cachedToken = { value: result.access_token, expiresAt: Date.now() + (result.expires_in * 1000) };
  return cachedToken.value;
}

async function google(env, url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: { authorization: `Bearer ${await serviceToken(env)}`, ...(options.headers || {}) },
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(body.error?.message || `Google API error (${response.status})`);
    error.status = response.status;
    throw error;
  }
  return body;
}

async function googleAsAdmin(accessToken, url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: { authorization: `Bearer ${accessToken}`, ...(options.headers || {}) },
  });
  const body = await response.json().catch(() => ({}));
  if (!response.ok) {
    const error = new Error(body.error?.message || `Google API error (${response.status})`);
    error.status = response.status;
    throw error;
  }
  return body;
}

async function values(env, sheetId, range) {
  return google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}`);
}

async function put(env, sheetId, range, data) {
  return google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}?valueInputOption=RAW`, {
    method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ range, majorDimension: 'ROWS', values: data }),
  });
}

async function append(env, sheetId, range, data) {
  return google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`, {
    method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ majorDimension: 'ROWS', values: data }),
  });
}

async function clearValues(env, sheetId, range) {
  return google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}/values/${encodeURIComponent(range)}:clear`, {
    method: 'POST', headers: { 'content-type': 'application/json' }, body: '{}',
  });
}

async function userFrom(request, env) {
  const credential = request.headers.get('authorization')?.replace(/^Bearer\s+/i, '');
  if (!credential) throw new Error('請先登入。');
  const response = await fetch(`https://oauth2.googleapis.com/tokeninfo?id_token=${encodeURIComponent(credential)}`);
  const identity = await response.json();
  const clientId = env.GOOGLE_CLIENT_ID || '234893329401-s9cnhole5d3q885do2b5349btlifk46r.apps.googleusercontent.com';
  if (!response.ok || identity.aud !== clientId || identity.email_verified !== 'true') throw new Error('Google 身分驗證失敗。');
  const email = String(identity.email || '').toLowerCase();
  const admins = (env.ADMIN_EMAILS || '').split(',').map(value => value.trim().toLowerCase());
  if (admins.includes(email)) return { email, name: identity.name || email, role: 'admin' };
  const list = await values(env, env.ALLOW_LIST_SHEET_ID, 'A:A');
  const permitted = (list.values || []).slice(1).some(row => String(row[0] || '').trim().toLowerCase() === email);
  if (!permitted) throw new Error('此帳號未獲授權使用報到系統。');
  return { email, name: identity.name || email, role: 'volunteer' };
}

function admin(user) { if (user.role !== 'admin') throw new Error('只有管理者可以執行此操作。'); }

async function assertEvent(env, sheetId) {
  if (!env.EVENT_REGISTRY_SHEET_ID) {
    const error = new Error('系統活動登錄設定不完整，請聯絡系統管理者。');
    error.status = 503;
    throw error;
  }
  const registry = await values(env, env.EVENT_REGISTRY_SHEET_ID, 'A:K');
  // The registry may be newly created without a header row.  Accept a matching
  // entry anywhere so the first activity is usable as soon as it is registered.
  const rowIndex = (registry.values || []).findIndex(row => text(row[0]) === sheetId);
  const entry = rowIndex >= 0 ? registryEntry(registry.values[rowIndex], rowIndex + 1) : null;
  if (!entry) {
    const error = new Error('此活動尚未由管理者註冊。');
    error.status = 404;
    throw error;
  }
  return entry;
}

function rowsFrom(data) {
  return (data || []).slice(1).map((row, index) => ({
    rowIndex: index + 2, id: String(row[0] || '').trim(), name: row[1] || '', address: row[2] || '', notes: row[3] || '',
    checkinStamp: row[4] || '', checkinBy: row[5] || '', batchTime: row[6] || '',
  })).filter(row => row.id);
}

function validateRows(rows) {
  const seen = new Set();
  rows.forEach((row, index) => {
    if (!row.id || !row.name) throw new Error(`第 ${index + 2} 列必須包含編號與姓名。`);
    if (seen.has(row.id)) throw new Error(`編號「${row.id}」重複。`);
    seen.add(row.id);
  });
}

async function event(env, sheetId) {
  const registered = await assertEvent(env, sheetId);
  const [data, meta] = await Promise.all([
    values(env, sheetId, 'A:G'),
    google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=properties.title`),
  ]);
  return {
    sheetId,
    title: registered.title || meta.properties?.title || '',
    createdAt: registered.createdAt,
    startDate: registered.startDate,
    endDate: registered.endDate,
    startTime: registered.startTime,
    endTime: registered.endTime,
    location: registered.location,
    notice: registered.notice,
    batchDigits: registered.batchDigits,
    batches: registered.batches,
    rows: rowsFrom(data.values),
  };
}

async function createEvent(env, user, payload) {
  admin(user);
  const name = String(payload.name || '').trim();
  const attendees = payload.attendees || [];
  const driveAccessToken = String(payload.driveAccessToken || '');
  if (!driveAccessToken) throw new Error('Google Drive authorization is required to create an activity.');
  if (!name) throw new Error('請輸入活動名稱。');
  validateRows(attendees);
  const metadata = eventMetaFromPayload(payload);
  const created = await googleAsAdmin(driveAccessToken, 'https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST', headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ properties: { title: name }, sheets: [{ properties: { title: '名單', gridProperties: { frozenRowCount: 1 } } }] }),
  });
  const data = [['編號', '姓名', '通訊地址', '備註', '報到時間', '報到人員', '通知報到時段'], ...attendees.map(row => [row.id, row.name, row.address || '', row.notes || '', '', '', ''])];
  const service = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON);
  await googleAsAdmin(driveAccessToken, `https://www.googleapis.com/drive/v3/files/${created.spreadsheetId}/permissions?sendNotificationEmail=false`, {
    method: 'POST', headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ type: 'user', role: 'writer', emailAddress: service.client_email }),
  });
  await put(env, created.spreadsheetId, `A1:G${data.length}`, data);
  await append(env, env.EVENT_REGISTRY_SHEET_ID, 'A:K', registryRow({
    sheetId: created.spreadsheetId,
    title: name,
    createdAt: new Date().toISOString(),
    ...metadata,
  }));
  return event(env, created.spreadsheetId);
}

async function replaceRoster(env, user, payload) {
  admin(user);
  const sheetId = String(payload.sheetId || '').trim();
  if (!sheetId) throw new Error('An activity must be selected before uploading a roster.');
  const entry = await assertEvent(env, sheetId);
  const existing = await values(env, sheetId, 'A:G');
  const hasExistingRosterData = (existing.values || []).slice(1).some(row =>
    row.some(value => String(value ?? '').trim())
  );
  if (hasExistingRosterData) {
    const error = new Error('此活動已有名單，為避免覆寫既有資料，系統已拒絕上傳。請先建立新活動。');
    error.status = 409;
    throw error;
  }
  const attendees = payload.attendees || [];
  validateRows(attendees);
  const data = [['編號', '姓名', '通訊地址', '備註', '報到時間', '報到人員', '通知報到時段'],
    ...attendees.map(row => [row.id, row.name, row.address || '', row.notes || '', '', '', ''])];
  let batchValues = null;
  if (entry.batches.length) {
    const result = validateAndAssignBatches(attendees.map((row, index) => ({ ...row, rowIndex: index + 2 })), entry.batches, entry.batchDigits);
    if (result.errors.length) {
      const error = new Error('現有分批設定無法套用到新名單，請先修正分批設定。');
      error.status = 422;
      error.details = result.errors;
      throw error;
    }
    batchValues = attendees.map((row, index) => [result.assignments.get(index + 2) || '']);
  }
  if (batchValues) batchValues.forEach((value, index) => { data[index + 1][6] = value[0]; });
  await clearValues(env, sheetId, 'A:G');
  await put(env, sheetId, `A1:G${data.length}`, data);
  return event(env, sheetId);
}

async function importEvent(env, user, payload) {
  admin(user);
  const sheetId = String(payload.sheetId || '').trim();
  if (!sheetId) throw new Error('缺少活動試算表 ID。');
  const data = await values(env, sheetId, 'A:G');
  if (((data.values || [])[0] || []).slice(0, 4).join('|') !== '編號|姓名|通訊地址|備註') throw new Error('活動名單欄位必須為：編號、姓名、通訊地址、備註。');
  validateRows(rowsFrom(data.values));
  const meta = await google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=properties.title`);
  await append(env, env.EVENT_REGISTRY_SHEET_ID, 'A:K', registryRow({
    sheetId,
    title: meta.properties?.title || '',
    createdAt: new Date().toISOString(),
    ...eventMetaFromPayload(payload),
  }));
  return event(env, sheetId);
}

function normaliseRangeEndpoint(value, digits) {
  const raw = text(value);
  return /^\d+$/.test(raw) && raw.length <= digits ? raw.padStart(digits, '0') : raw;
}

function isWalkinId(id) { return /^W-/.test(text(id)); }

function validateAndAssignBatches(rows, batches, digits) {
  const errors = [];
  const normalised = (batches || []).map((batch, index) => {
    const min = normaliseRangeEndpoint(batch.min, digits);
    const max = normaliseRangeEndpoint(batch.max, digits);
    const time = text(batch.time);
    if (!/^\d+$/.test(min) || min.length !== digits || !/^\d+$/.test(max) || max.length !== digits) {
      errors.push(`第 ${index + 1} 批的編號範圍必須是 ${digits} 位數字。`);
    } else if (Number(min) > Number(max)) {
      errors.push(`第 ${index + 1} 批的起始編號不可大於結束編號。`);
    }
    if (!time) errors.push(`第 ${index + 1} 批尚未填寫報到時段。`);
    return { min, max, time };
  });

  const sorted = normalised
    .map((batch, index) => ({ ...batch, index }))
    .filter(batch => /^\d+$/.test(batch.min) && /^\d+$/.test(batch.max))
    .sort((a, b) => Number(a.min) - Number(b.min));
  for (let i = 1; i < sorted.length; i += 1) {
    if (Number(sorted[i].min) <= Number(sorted[i - 1].max)) {
      errors.push(`第 ${sorted[i - 1].index + 1} 批與第 ${sorted[i].index + 1} 批的編號範圍重疊。`);
    }
  }

  const assignments = new Map();
  (rows || []).forEach(row => {
    // Walk-ins are created at the venue and do not have a pre-assigned
    // notification slot.  They remain in attendance/report data but are not
    // part of the administrator's numbered roster batch rules.
    if (isWalkinId(row.id)) return;
    const suffix = text(row.id).slice(-digits);
    if (!/^\d+$/.test(suffix) || suffix.length !== digits) {
      errors.push(`編號「${row.id}」無法解析最後 ${digits} 位數字。`);
      return;
    }
    const number = Number(suffix);
    const matches = normalised.filter(batch => number >= Number(batch.min) && number <= Number(batch.max));
    if (matches.length === 0) errors.push(`編號「${row.id}」沒有對應的報到時段。`);
    else if (matches.length > 1) errors.push(`編號「${row.id}」同時符合多個報到時段。`);
    else assignments.set(row.rowIndex, matches[0].time);
  });

  return { batches: normalised, assignments, errors: [...new Set(errors)] };
}

async function updateRegistryEntry(env, entry, metadata) {
  await put(env, env.EVENT_REGISTRY_SHEET_ID, `A${entry.rowIndex}:K${entry.rowIndex}`, registryRow({
    ...entry,
    ...metadata,
  }));
}

async function updateEvent(env, user, payload) {
  admin(user);
  const sheetId = text(payload.sheetId);
  if (!sheetId) throw new Error('缺少活動試算表 ID。');
  const entry = await assertEvent(env, sheetId);
  const metadata = eventMetaFromPayload(payload, entry);
  let assignments = null;
  let previousBatchColumn = null;
  if (payload.batches !== undefined) {
    const current = await event(env, sheetId);
    const result = validateAndAssignBatches(current.rows, metadata.batches, metadata.batchDigits);
    if (result.errors.length) {
      const error = new Error('報到時段設定有誤，尚未儲存。');
      error.status = 422;
      error.details = result.errors;
      throw error;
    }
    metadata.batches = result.batches;
    assignments = result.assignments;
    const raw = await values(env, sheetId, 'A:G');
    previousBatchColumn = (raw.values || []).slice(1).map(row => [row[6] || '']);
    const batchColumn = (raw.values || []).slice(1).map((row, index) => [
      assignments.has(index + 2) ? assignments.get(index + 2) : (metadata.batches.length ? (row[6] || '') : ''),
    ]);
    try {
      if (batchColumn.length) await put(env, sheetId, `G2:G${batchColumn.length + 1}`, batchColumn);
      await updateRegistryEntry(env, entry, metadata);
    } catch (error) {
      if (previousBatchColumn?.length) {
        try { await put(env, sheetId, `G2:G${previousBatchColumn.length + 1}`, previousBatchColumn); } catch {}
      }
      throw error;
    }
    return event(env, sheetId);
  }
  await updateRegistryEntry(env, entry, metadata);
  return event(env, sheetId);
}

async function updateAttendee(env, user, payload) {
  admin(user);
  const sheetId = text(payload.sheetId);
  const id = text(payload.id);
  if (!sheetId || !id) throw new Error('缺少活動或參與者編號。');
  const current = await event(env, sheetId);
  const row = current.rows.find(item => item.id === id);
  if (!row) {
    const error = new Error('找不到此參與者。');
    error.status = 404;
    throw error;
  }
  await put(env, sheetId, `C${row.rowIndex}:C${row.rowIndex}`, [[text(payload.address)]]);
  return event(env, sheetId);
}

async function deleteEvent(env, user, payload) {
  admin(user);
  const sheetId = text(payload.sheetId);
  const driveAccessToken = text(payload.driveAccessToken);
  if (!sheetId) throw new Error('缺少活動試算表 ID。');
  if (!driveAccessToken) throw new Error('刪除活動需要重新取得 Google Drive 授權。');
  const entry = await assertEvent(env, sheetId);
  const current = await event(env, sheetId);
  if (current.rows.some(row => row.checkinStamp)) {
    const error = new Error('這場活動已有報到紀錄，為保留資料目前不能刪除。請保留此活動以供查閱。');
    error.status = 409;
    throw error;
  }
  await googleAsAdmin(driveAccessToken, `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(sheetId)}?supportsAllDrives=true`, {
    method: 'PATCH',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ trashed: true }),
  });
  try {
    await clearValues(env, env.EVENT_REGISTRY_SHEET_ID, `A${entry.rowIndex}:K${entry.rowIndex}`);
  } catch (error) {
    try {
      await googleAsAdmin(driveAccessToken, `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(sheetId)}?supportsAllDrives=true`, {
        method: 'PATCH',
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify({ trashed: false }),
      });
    } catch {}
    throw new Error(`活動已移至 Google Drive 垃圾桶，但系統登錄清除失敗：${error.message || error}`);
  }
  return { deleted: true, sheetId, title: entry.title };
}

function assertVolunteerWriteAllowed(user, current) {
  if (user.role !== 'volunteer' || !current.endDate) return;
  const end = Date.parse(`${current.endDate}T23:59:59+08:00`);
  if (!Number.isFinite(end) || Date.now() <= end + (3 * 24 * 60 * 60 * 1000)) return;
  const error = new Error('活動報到已結束，志工目前只能查看資料。');
  error.status = 410;
  throw error;
}

function parseQrPayload(raw) {
  const prefix = 'qr-checkin:v2:';
  const code = text(raw);
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

async function allowList(env, user, payload) {
  admin(user);
  const data = await values(env, env.ALLOW_LIST_SHEET_ID, 'A:A');
  const entries = (data.values || []).slice(1).map((row, index) => ({ rowIndex: index + 2, email: String(row[0] || '').trim().toLowerCase() })).filter(row => row.email);
  if (payload.method === 'GET') return { entries };
  const email = String(payload.email || '').trim().toLowerCase();
  if (!email.includes('@')) throw new Error('請輸入有效的 email。');
  if (payload.method === 'POST') {
    if (entries.some(entry => entry.email === email)) throw new Error('此帳號已在名單中。');
    await append(env, env.ALLOW_LIST_SHEET_ID, 'A:A', [[email]]);
    return { ok: true };
  }
  const existing = entries.find(entry => entry.email === email);
  if (!existing) throw new Error('找不到此帳號。');
  await put(env, env.ALLOW_LIST_SHEET_ID, `A${existing.rowIndex}`, [['']]);
  return { ok: true };
}

async function shareEvent(env, user, payload) {
  admin(user);
  await assertEvent(env, payload.sheetId);
  const entries = (await allowList(env, user, { method: 'GET' })).entries;
  // Volunteers use the app link and Worker API; never grant them direct Sheet access.
  return { shared: entries.length };
}

async function checkin(env, user, payload, undo = false) {
  const current = await event(env, payload.sheetId);
  assertVolunteerWriteAllowed(user, current);
  let id = text(payload.id);
  if (!undo && !payload.manual) {
    const qr = parseQrPayload(payload.qrPayload);
    if (!qr || qr.eventId !== current.sheetId) {
      const error = new Error('此 QR 不屬於目前活動，未寫入報到紀錄。');
      error.status = 409;
      throw error;
    }
    id = qr.id;
  }
  const row = current.rows.find(item => item.id === id);
  if (!row) throw new Error('找不到此編號。');
  if (!undo && row.checkinStamp) return { row, already: true };
  const stamp = undo ? '' : taipeiTimestamp();
  const by = undo ? '' : `${user.name}${payload.manual ? '（手動）' : ''}`;
  await put(env, payload.sheetId, `E${row.rowIndex}:F${row.rowIndex}`, [[stamp, by]]);
  return { row: { ...row, checkinStamp: stamp, checkinBy: by }, already: false };
}

async function addWalkin(env, user, payload) {
  const current = await event(env, payload.sheetId);
  assertVolunteerWriteAllowed(user, current);
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('請輸入姓名。');
  const row = [`W-${crypto.randomUUID()}`, name, String(payload.address || ''), String(payload.notes || '現場報名'), taipeiTimestamp(), `${user.name}（手動）`];
  await append(env, payload.sheetId, 'A:F', [row]);
  return event(env, payload.sheetId);
}

async function handle(request, env) {
  const cors = allowedCors(request, env);
  if (request.method === 'OPTIONS') return new Response(null, { headers: { ...cors, 'access-control-allow-methods': 'POST, OPTIONS', 'access-control-allow-headers': 'authorization, content-type' } });
  if (request.method !== 'POST') return asJson({ error: 'Not found' }, 404, cors);
  try {
    const payload = await request.json();
    const user = await userFrom(request, env);
    let result;
    if (payload.action === 'session') result = { user };
    else if (payload.action === 'event.get') result = await event(env, payload.sheetId);
    else if (payload.action === 'event.create') result = await createEvent(env, user, payload);
    else if (payload.action === 'event.roster.replace') result = await replaceRoster(env, user, payload);
    else if (payload.action === 'event.import') result = await importEvent(env, user, payload);
    else if (payload.action === 'event.update') result = await updateEvent(env, user, payload);
    else if (payload.action === 'attendee.update') result = await updateAttendee(env, user, payload);
    else if (payload.action === 'event.delete') result = await deleteEvent(env, user, payload);
    else if (payload.action === 'checkin') result = await checkin(env, user, payload);
    else if (payload.action === 'undo') result = await checkin(env, user, payload, true);
    else if (payload.action === 'walkin.add') result = await addWalkin(env, user, payload);
    else if (payload.action === 'allowlist') result = await allowList(env, user, payload);
    else if (payload.action === 'event.share') result = await shareEvent(env, user, payload);
    else throw new Error('不支援的操作。');
    return asJson(result, 200, cors);
  } catch (error) {
    return asJson({ error: error.message || 'Request failed.', ...(error.details ? { details: error.details } : {}) }, error.status || 400, cors);
  }
}

export default { fetch: handle };
