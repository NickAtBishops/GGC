# GGC Deal Engine — Web

Hosted frontend for the GGC Deal Engine. Next.js 15 (App Router) + TypeScript +
Tailwind CSS + Firebase (Google sign-in, Firestore run history, Storage
downloads). Talks to the Python engine (Cloud Run) over three endpoints:
`POST /api/analyze`, `GET /api/status/{job_id}`, `GET /api/download/{job_id}`,
all bearer-authenticated with the user's Firebase ID token.

## Install

```bash
cd web
npm install
```

## Environment

```bash
cp .env.local.example .env.local
```

Fill in the Firebase web-app config (Firebase Console → Project settings →
General → Your apps → SDK setup) and the engine base URL:

- `NEXT_PUBLIC_FIREBASE_API_KEY` / `_AUTH_DOMAIN` / `_PROJECT_ID` /
  `_STORAGE_BUCKET` / `_MESSAGING_SENDER_ID` / `_APP_ID` — Firebase web config.
- `NEXT_PUBLIC_ENGINE_URL` — engine base URL, no trailing slash. Defaults to
  `http://localhost:5001` (local Flask) when unset; in production point it at
  the Cloud Run service URL.

One-time Firebase setup (covered by the deploy runbook, `../DEPLOYMENT.md`):
enable the **Google** provider under Authentication → Sign-in method, add your
dev and production domains under Authentication → Authorized domains, and
create the Firestore composite index on `deal_runs` (`uid` asc, `createdAt`
desc) — until it exists the History page shows Firestore's index-required
message with a create link.

## Develop

```bash
npm run dev        # http://localhost:3000, engine assumed at :5001
```

## Build

```bash
npm run build && npm start
```

## Deploy to Vercel

1. Import the repo in Vercel (or run `vercel` from this directory) and set the
   project **Root Directory** to `web/`. Framework preset: Next.js; default
   build command (`next build`) works as-is.
2. Add every `NEXT_PUBLIC_*` variable from `.env.local.example` in Project
   Settings → Environment Variables. They are inlined at build time — redeploy
   after changing any of them.
3. Set `NEXT_PUBLIC_ENGINE_URL` to the Cloud Run URL, and make sure the engine
   allows the Vercel domain in its CORS config (`ALLOWED_ORIGINS`).
4. Add the Vercel production (and preview, if used) domain to Firebase
   Authentication → Authorized domains so the Google sign-in popup works.
