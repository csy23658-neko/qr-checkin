# Cloudflare staging deployment

This project is split into a public Pages site and a private Worker API.  Do not deploy the app until the Worker is configured; the browser must never receive a Sheets access token or a service-account key.

## 1. Prepare Google resources

1. Create a dedicated **staging** Google Cloud project and enable Google Sheets API and Google Drive API.
2. Create a Web OAuth client. Add the fixed Pages staging URL to **Authorized JavaScript origins**.
3. Create a service account and create a JSON key for it. Store the JSON only as a Cloudflare Worker secret.
4. Create two private Google Sheets owned by the service account:
   - Allow list: column A is `email`; volunteer emails start on row 2.
   - Event registry: columns A–C are `sheetId`, `title`, `createdAt`; row 1 is the header.
5. For an existing event Sheet, share it with the service-account email before importing it through the Dashboard.

## 2. Deploy the Worker

From `worker/`, authenticate Wrangler and create the secrets from `.dev.vars.example`:

```powershell
npx wrangler secret put GOOGLE_CLIENT_ID
npx wrangler secret put GOOGLE_SERVICE_ACCOUNT_JSON
npx wrangler secret put ADMIN_EMAILS
npx wrangler secret put ALLOW_LIST_SHEET_ID
npx wrangler secret put EVENT_REGISTRY_SHEET_ID
npx wrangler deploy
```

Before deploying, edit `ALLOWED_ORIGINS` in `worker/wrangler.jsonc` to the exact Pages staging origin. The API validates that origin rather than returning permissive CORS headers.

Copy the resulting Worker HTTPS URL into `CONFIG.API_BASE_URL` in `config.js`; this value is public and is safe to commit.

## 3. Deploy Pages and protect staging

1. Connect this repository to Cloudflare Pages.
2. Use a fixed staging branch alias or custom staging domain. Do not rely on random preview URLs for OAuth because every OAuth origin must be explicitly registered.
3. Configure Cloudflare Access for the staging URL. Initially allow only the developer email; after internal acceptance, add named nonprofit administrators and volunteers.
4. Add the final Pages staging URL to the Google OAuth client and use the same URL in `ALLOWED_ORIGINS`.

## 4. Handover

After acceptance, give nonprofit administrators access to the repository, Cloudflare project, Google Cloud project, service account, and Sheets. Create a separate production Worker and Pages deployment with production-only secrets and origins. The nonprofit can then move the custom domain independently.
