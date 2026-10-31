# Vexora Labs Customer Usage Portal — Cloud Run

- Phone-only lookup: no OTP or PIN, as requested.
- Normalizes Sri Lankan mobile formats: 077..., +9477..., 9477..., 009477...
- Searches all monthly tabs and shows each matching order separately.
- Subscription URLs are fetched by the backend and never returned to the browser.
- Includes a short per-instance cache and basic per-IP request limiting.

## Google Cloud setup
1. Create/select a Google Cloud project; enable Cloud Run and Google Sheets API.
2. Create a service account for the Cloud Run service.
3. Share the Google Sheet with that service-account email as Viewer.
4. Deploy this folder to Cloud Run using that service account (Application Default Credentials; never put credentials in frontend code).
5. Set environment variables: `SPREADSHEET_ID` (sheet ID), `ALLOWED_SUBSCRIPTION_HOSTS=sg3.vpnlanka.store`, `CACHE_TTL_SECONDS=60`.

## Deploy
```bash
gcloud run deploy vexora-usage-api --source . --region asia-southeast1 --allow-unauthenticated --service-account YOUR_SERVICE_ACCOUNT --set-env-vars SPREADSHEET_ID=YOUR_SHEET_ID,ALLOWED_SUBSCRIPTION_HOSTS=sg3.vpnlanka.store,CACHE_TTL_SECONDS=60 --concurrency 40 --max-instances 50
```
Replace placeholders. The included page is served from the same Cloud Run domain as the API. Copy the Vexora logo to `public/assets/vex.jpg`.

## Sheet columns
Expected existing structure: Date | Contact | Name | ISP Package name | GB | price | Configuration | Subscribtion usage(link). Contact is column B; subscription link is column H. Header detection is also supported.

## Limitations
No OTP/PIN is included. Anyone who knows a registered phone number may be able to look up that customer's subscriptions; this is the privacy trade-off. Subscription links stay on the backend. Rate limiting and cache are per instance, not global. This Cloud Run configuration is not a guarantee of 1,000 concurrent successful requests: load-test Cloud Run, Sheets API and provider limits. For sustained high traffic, sync orders to Firestore/Cloud SQL and add shared Redis cache / Cloud Armor.
