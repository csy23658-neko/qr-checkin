# Cloudflare staging deployment

This project is split into a public Pages site and a private Worker API.  Do not deploy the app until the Worker is configured; the browser must never receive a Sheets access token or a service-account key.

## Current isolated staging deployment

The event-bound branch is deployed to separate Cloudflare resources:

- Pages: `https://qr-checkin-event-bound-staging.pages.dev`
- Worker API: `https://qr-checkin-event-bound-staging.csy23658.workers.dev`
- Branch: `codex/event-bound-qr-notice-output`

These resources are independent of the existing `qr-checkin` Pages project and `qr-checkin-api-staging` Worker. The staging Worker still requires `ADMIN_EMAILS` and `GOOGLE_SERVICE_ACCOUNT_JSON` to be added as secrets from a secure source before authenticated API actions can work.

## 1. Prepare Google resources

1. Create a dedicated **staging** Google Cloud project and enable Google Sheets API and Google Drive API.
2. Create a Web OAuth client. Add the fixed Pages staging URL to **Authorized JavaScript origins**.
3. Create a service account and create a JSON key for it. Store the JSON only as a Cloudflare Worker secret.
4. Create two private Google Sheets owned by the administrator and share them as **Editor** with the service account:
   - Allow list: column A is `email`; volunteer emails start on row 2.
   - Event registry: row 1 is the header; columns A–K are `sheetId`, `title`, `createdAt`, `startDate`, `endDate`, `startTime`, `endTime`, `location`, `notice`, `batches` (JSON), and `batchDigits`.
5. For an existing event Sheet, share it with the service-account email before importing it through the Dashboard.

Activity Sheets created by the Dashboard use columns A–G: `編號`, `姓名`, `通訊地址`, `備註`, `報到時間`, `報到人員`, and `通知報到時段`.  Existing activities registered with only the original A–C registry fields remain readable; their optional metadata defaults to blank until an administrator saves it.

## 2. Deploy the Worker

From the repository root, authenticate Wrangler and add the two staging-only secrets (values must come from your secure source):

```powershell
npx wrangler secret put GOOGLE_SERVICE_ACCOUNT_JSON --name qr-checkin-event-bound-staging
npx wrangler secret put ADMIN_EMAILS --name qr-checkin-event-bound-staging
npx wrangler deploy --config worker/wrangler.event-bound-staging.jsonc
```

The staging OAuth client ID and the two staging Sheet IDs are non-secret variables in `worker/wrangler.event-bound-staging.jsonc`.

Before deploying, edit `ALLOWED_ORIGINS` in `worker/wrangler.jsonc` to the exact Pages staging origin. The API validates that origin rather than returning permissive CORS headers.

Copy the resulting Worker HTTPS URL into `CONFIG.API_BASE_URL` in `config.js`; this value is public and is safe to commit.

## OAuth sign-in checklist

If Google shows `400: invalid_request` with “You can't sign in to this app because it doesn't comply with Google's OAuth 2.0 policy”, the request is being blocked by Google before it reaches the Worker. In the Google Cloud project that owns the client ID:

1. In **Google Auth Platform → Branding**, complete the app name, support email, homepage, privacy policy, and authorized-domain settings required by Google.
2. In **Audience**, add the administrator's Google account as a test user while the app is in Testing, or publish the app when it is ready for production.
3. In **Clients**, use a **Web application** OAuth client and add the exact Pages origin (for this staging deployment: `https://qr-checkin-event-bound-staging.pages.dev`) under **Authorized JavaScript origins**. Do not use a preview URL unless that origin is registered and the Worker CORS allow-list is updated too.
4. Confirm the client ID in `config.js` is the same client ID configured in the Worker's `GOOGLE_CLIENT_ID` secret.

The QR system's own `ADMIN_EMAILS` and volunteer allow-list are checked only after Google successfully returns an ID token; they cannot fix an OAuth `invalid_request` page.

## Local browser testing

Do not open `index.html` by double-clicking it.  A `file://` page has no valid web origin for Google Identity Services and can be rejected with `400: invalid_request`.

From the repository root, serve the files over HTTP instead:

```powershell
python -m http.server 4173 --bind 127.0.0.1
```

Then open `http://127.0.0.1:4173/` (or `http://localhost:4173/`).  If Google login is needed locally, add the exact local origin to the OAuth Web client and use a Worker environment whose CORS allow-list includes that same origin.  The deployed staging Worker allows the fixed staging Pages origin and these local origins.

## 3. Deploy Pages and protect staging

1. Connect this repository to Cloudflare Pages.
2. Use a fixed staging branch alias or custom staging domain. Do not rely on random preview URLs for OAuth because every OAuth origin must be explicitly registered.
3. Configure Cloudflare Access for the staging URL. Initially allow only the developer email; after internal acceptance, add named nonprofit administrators and volunteers.
4. Add the final Pages staging URL to the Google OAuth client and use the same URL in `ALLOWED_ORIGINS`.

The activity-name link preview uses the Pages Function in `functions/index.js`, so deploy the project through a Git-connected Pages project or Wrangler. Cloudflare Pages direct-upload deployments do not include Functions.

## 4. Handover

After acceptance, give nonprofit administrators access to the repository, Cloudflare project, Google Cloud project, service account, and Sheets. Create a separate production Worker and Pages deployment with production-only secrets and origins. The nonprofit can then move the custom domain independently.
