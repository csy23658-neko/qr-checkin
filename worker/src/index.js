let cachedToken = { value: '', expiresAt: 0 };

const encode = value => new TextEncoder().encode(value);
const asJson = (body, status = 200, headers = {}) => new Response(JSON.stringify(body), {
  status, headers: { 'content-type': 'application/json; charset=utf-8', ...headers },
});

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
  if (!response.ok) throw new Error(body.error?.message || `Google API error (${response.status})`);
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
  const registry = await values(env, env.EVENT_REGISTRY_SHEET_ID, 'A:A');
  if (!(registry.values || []).slice(1).some(row => row[0] === sheetId)) throw new Error('此活動尚未由管理者註冊。');
}

function rowsFrom(data) {
  return (data || []).slice(1).map((row, index) => ({
    rowIndex: index + 2, id: String(row[0] || '').trim(), name: row[1] || '', address: row[2] || '', notes: row[3] || '',
    checkinStamp: row[4] || '', checkinBy: row[5] || '',
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
  await assertEvent(env, sheetId);
  const [data, meta] = await Promise.all([
    values(env, sheetId, 'A:F'),
    google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=properties.title`),
  ]);
  return { sheetId, title: meta.properties?.title || '', rows: rowsFrom(data.values) };
}

async function createEvent(env, user, payload) {
  admin(user);
  const name = String(payload.name || '').trim();
  const attendees = payload.attendees || [];
  if (!name) throw new Error('請輸入活動名稱。');
  validateRows(attendees);
  const created = await google(env, 'https://sheets.googleapis.com/v4/spreadsheets', {
    method: 'POST', headers: { 'content-type': 'application/json' },
    body: JSON.stringify({ properties: { title: name }, sheets: [{ properties: { title: '名單', gridProperties: { frozenRowCount: 1 } } }] }),
  });
  const data = [['編號', '姓名', '通訊地址', '備註', '報到時間', '報到人員'], ...attendees.map(row => [row.id, row.name, row.address || '', row.notes || '', '', ''])];
  await put(env, created.spreadsheetId, `A1:F${data.length}`, data);
  await append(env, env.EVENT_REGISTRY_SHEET_ID, 'A:C', [[created.spreadsheetId, name, new Date().toISOString()]]);
  return event(env, created.spreadsheetId);
}

async function importEvent(env, user, payload) {
  admin(user);
  const sheetId = String(payload.sheetId || '').trim();
  if (!sheetId) throw new Error('缺少活動試算表 ID。');
  const data = await values(env, sheetId, 'A:F');
  if (((data.values || [])[0] || []).slice(0, 4).join('|') !== '編號|姓名|通訊地址|備註') throw new Error('活動名單欄位必須為：編號、姓名、通訊地址、備註。');
  validateRows(rowsFrom(data.values));
  const meta = await google(env, `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?fields=properties.title`);
  await append(env, env.EVENT_REGISTRY_SHEET_ID, 'A:C', [[sheetId, meta.properties?.title || '', new Date().toISOString()]]);
  return event(env, sheetId);
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
  const row = current.rows.find(item => item.id === String(payload.id || '').trim());
  if (!row) throw new Error('找不到此編號。');
  if (!undo && row.checkinStamp) return { row, already: true };
  const stamp = undo ? '' : new Date().toLocaleString('zh-TW', { hour12: false });
  const by = undo ? '' : `${user.name}${payload.manual ? '（手動）' : ''}`;
  await put(env, payload.sheetId, `E${row.rowIndex}:F${row.rowIndex}`, [[stamp, by]]);
  return { row: { ...row, checkinStamp: stamp, checkinBy: by }, already: false };
}

async function addWalkin(env, user, payload) {
  await assertEvent(env, payload.sheetId);
  const name = String(payload.name || '').trim();
  if (!name) throw new Error('請輸入姓名。');
  const row = [`W-${crypto.randomUUID()}`, name, String(payload.address || ''), String(payload.notes || '現場報名'), new Date().toLocaleString('zh-TW', { hour12: false }), `${user.name}（手動）`];
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
    else if (payload.action === 'event.import') result = await importEvent(env, user, payload);
    else if (payload.action === 'checkin') result = await checkin(env, user, payload);
    else if (payload.action === 'undo') result = await checkin(env, user, payload, true);
    else if (payload.action === 'walkin.add') result = await addWalkin(env, user, payload);
    else if (payload.action === 'allowlist') result = await allowList(env, user, payload);
    else if (payload.action === 'event.share') result = await shareEvent(env, user, payload);
    else throw new Error('不支援的操作。');
    return asJson(result, 200, cors);
  } catch (error) {
    return asJson({ error: error.message || 'Request failed.' }, 400, cors);
  }
}

export default { fetch: handle };
