function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setMeta(html, attribute, key, value) {
  const safeValue = escapeHtml(value);
  const tag = `<meta ${attribute}="${key}" content="${safeValue}">`;
  const pattern = new RegExp(`<meta\\s+${attribute}=["']${key}["'][^>]*>`, 'i');
  if (pattern.test(html)) return html.replace(pattern, () => tag);
  return html.replace(/<\/head>/i, `${tag}\n</head>`);
}

export async function onRequest(context) {
  const response = await context.next();
  const contentType = response.headers.get('content-type') || '';
  const activityName = new URL(context.request.url).searchParams.get('name')?.trim();

  if (!response.ok || !contentType.includes('text/html') || !activityName) return response;

  const title = `QR 報到系統（${activityName}）`;
  const description = `${title}，請使用 Google 帳號登入。`;
  const requestUrl = new URL(context.request.url).toString();
  let html = await response.text();

  html = html.replace(/<title>[\s\S]*?<\/title>/i, `<title>${escapeHtml(title)}</title>`);
  html = setMeta(html, 'name', 'description', description);
  html = setMeta(html, 'property', 'og:title', title);
  html = setMeta(html, 'property', 'og:description', description);
  html = setMeta(html, 'property', 'og:url', requestUrl);

  const headers = new Headers(response.headers);
  headers.delete('content-length');
  headers.delete('etag');
  return new Response(html, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}
