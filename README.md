# Henna Cloud CRM

Cloud backend for Click-to-WhatsApp campaign leads.

## What it does
- Receives Meta WhatsApp webhooks 24/7.
- Creates a CRM lead only when the first message contains Click-to-WhatsApp referral data.
- Direct WhatsApp messages do not create leads.
- Later messages from an existing campaign lead update the same lead for 24 hours.
- Stores leads, every WhatsApp message, payments and follow-ups in PostgreSQL.
- Prevents duplicate Meta messages with a unique message ID.
- Exports the database to Excel via `/api/export.xlsx`.
- Includes a protected JSON API for the web CRM we will build next.

## Railway deployment

1. Create a Railway project.
2. Add PostgreSQL to the project.
3. Add this application as a service from GitHub or `railway up`.
4. Link the app's `DATABASE_URL` to the Postgres service.
5. Add all secrets from `.env.example` to the app's Variables tab.
6. Generate a public domain for the app.
7. In Meta Webhooks, replace the ngrok callback with:
   `https://YOUR-RAILWAY-DOMAIN/webhook`
8. Use the same `VERIFY_TOKEN` value and click Verify and Save.
9. Keep the `messages` webhook field subscribed.
10. Test `/health` in a browser.

## Important
Never commit `.env`, Meta tokens, App Secret or ADMIN_API_KEY to GitHub.

## Admin API
Send the header:
`X-Admin-Key: <ADMIN_API_KEY>`

Useful endpoints:
- `GET /api/leads`
- `GET /api/leads/<id>`
- `PATCH /api/leads/<id>`
- `POST /api/leads/<id>/payments`
- `POST /api/leads/<id>/followups`
- `GET /api/export.xlsx`


## Excel sync client
The complete bundle also contains `local_sync/`. This runs on the Windows PC and pulls leads from `/api/leads` into Henna_CRM_Full.xlsx. Do not run the Excel sync process on Railway.
