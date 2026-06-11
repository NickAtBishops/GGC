# GGC Deal Engine — Deployment Runbook

This takes the project from "runs on localhost:5001" to hosted, authenticated, multi-user:

```
Browser ── Firebase Auth (Google sign-in)
   │
   ├──> Next.js web app ............ Vercel          (web/)
   │       │  Authorization: Bearer <Firebase ID token>
   │       v
   ├──> Analysis engine ............ Cloud Run       (backend.py, Dockerfile)
   │       │  Anthropic Fable 5 · Google Document AI · Google Maps
   │       v
   └──> Run history + downloads .... Firebase        (Firestore `deal_runs`,
                                                      Storage `runs/{uid}/{jobId}.xlsx`)
```

Local development is unchanged: `python3 backend.py` with no new env vars behaves exactly
as before (no auth, no Firebase). All hosted behavior is opt-in via env.

---

## 0. Prerequisites

- `gcloud` CLI (`brew install google-cloud-sdk`), `firebase` CLI (`npm i -g firebase-tools`),
  `vercel` CLI (`npm i -g vercel`), Node 20+.
- Logins are interactive — run each yourself (in Claude Code, prefix with `!`):
  `gcloud auth login`, `firebase login`, `vercel login`.
- **Use ONE Google Cloud project for everything** — the project already used for Document AI
  (`GCP_PROJECT_ID` in `.env`). Adding Firebase to that same project means Cloud Run's default
  service account reaches Document AI, Firestore, and Storage with no key files at all.

```sh
export PROJECT_ID=<your-gcp-project-id>     # same as GCP_PROJECT_ID in .env
export REGION=us-central1
gcloud config set project $PROJECT_ID
```

## 1. Firebase project setup (one-time, ~10 min, mostly console)

1. [console.firebase.google.com](https://console.firebase.google.com) → **Add project** →
   "Add Firebase to a Google Cloud project" → pick `$PROJECT_ID`. Upgrade to **Blaze** plan.
2. **Authentication → Sign-in method** → enable **Google**. (Optionally also Email/Password.)
3. **Firestore Database** → Create database (production mode, `$REGION`).
4. **Storage** → Get started. Note the bucket name — new projects get
   `$PROJECT_ID.firebasestorage.app`.
5. **Project settings → General → Your apps** → add a **Web app**. Copy the config values
   (apiKey, authDomain, projectId, storageBucket, messagingSenderId, appId) — these become the
   `NEXT_PUBLIC_FIREBASE_*` env vars in step 4.

Then deploy the security rules + index from this directory:

```sh
cd ggc-deal-engine
firebase use --add $PROJECT_ID          # writes .firebaserc (gitignored is fine)
firebase deploy --only firestore:rules,firestore:indexes,storage
```

The rules are deny-by-default: clients can only read `deal_runs` docs whose `uid` matches
their sign-in, and only download `runs/{their-uid}/...` files. All writes go through the
engine's Admin SDK.

## 2. Grant the engine's service account access (one-time)

Cloud Run runs as the Compute Engine default service account. Give it Firestore + Storage:

```sh
SA="$(gcloud projects describe $PROJECT_ID --format='value(projectNumber)')-compute@developer.gserviceaccount.com"
gcloud projects add-iam-policy-binding $PROJECT_ID --member=serviceAccount:$SA --role=roles/datastore.user
gcloud projects add-iam-policy-binding $PROJECT_ID --member=serviceAccount:$SA --role=roles/storage.objectAdmin
gcloud projects add-iam-policy-binding $PROJECT_ID --member=serviceAccount:$SA --role=roles/documentai.apiUser
gcloud services enable run.googleapis.com cloudbuild.googleapis.com artifactregistry.googleapis.com
```

No service-account JSON key is needed in the cloud — Application Default Credentials handle
Document AI, Firestore, and Storage. (`gcp-credentials.json` stays a local-dev-only file.)

## 3. Deploy the engine to Cloud Run

From `ggc-deal-engine/` (Cloud Build uses the `Dockerfile`; no local Docker needed):

```sh
gcloud run deploy ggc-deal-engine \
  --source . \
  --region $REGION \
  --allow-unauthenticated \
  --max-instances 1 \
  --no-cpu-throttling \
  --cpu 2 --memory 2Gi \
  --timeout 3600 \
  --concurrency 40 \
  --set-env-vars "REQUIRE_AUTH=1" \
  --set-env-vars "ALLOWED_EMAILS=nicholas.revenco@gmail.com,michael@<ggc-domain>.com" \
  --set-env-vars "FIREBASE_STORAGE_BUCKET=$PROJECT_ID.firebasestorage.app" \
  --set-env-vars "ALLOWED_ORIGINS=http://localhost:3000" \
  --set-env-vars "GCP_PROJECT_ID=$PROJECT_ID,GCP_LOCATION=us,GCP_LAYOUT_PROCESSOR_ID=<from .env>" \
  --set-env-vars "ANTHROPIC_API_KEY=<NEW key — see §6>,GOOGLE_MAPS_API_KEY=<NEW key — see §6>"
```

Note the service URL it prints (e.g. `https://ggc-deal-engine-xxxx-uc.a.run.app`) — that is
`NEXT_PUBLIC_ENGINE_URL` for the web app. After the web app is live (step 4), update
`ALLOWED_ORIGINS` to include its production URL:

```sh
gcloud run services update ggc-deal-engine --region $REGION \
  --update-env-vars "ALLOWED_ORIGINS=https://<your-app>.vercel.app,http://localhost:3000"
```

**Why these flags are not optional**

- `--max-instances 1` and the Dockerfile's single gunicorn worker: job state lives in process
  memory. A second instance/worker would 404 the status polls for jobs it didn't start.
- `--no-cpu-throttling`: analysis continues on background threads between status polls;
  throttled CPU would stall the pipeline mid-run.
- `--allow-unauthenticated` is safe **only because** `REQUIRE_AUTH=1` makes the app itself
  verify a Firebase ID token on `/api/analyze`, `/api/status/*`, `/api/download/*`, with
  per-uid job ownership and the email allowlist. Never deploy without it.
- Keep the browser tab open while a run is in flight: the 4-second status polls are what keep
  the instance alive with `min-instances 0`. If runs must survive a closed laptop, add
  `--min-instances 1` (costs roughly a continuously-running small VM).
- Tidier alternative for the two API keys: Secret Manager + `--set-secrets` instead of
  `--set-env-vars`.

**Scaling later:** to go beyond one instance, move job state out of process memory — read
job status from Firestore (already mirrored there) and serve downloads from Storage (already
uploaded there); then drop `--max-instances 1`. That work is listed in CLAUDE.md §12.2.

## 4. Deploy the web app to Vercel

```sh
cd web
cp .env.local.example .env.local     # fill in the values from Firebase step 1.5
npm install && npm run dev           # verify locally against the deployed engine
vercel                                # link/create the project (root = web/)
# In Vercel → Project → Settings → Environment Variables, add for Production:
#   NEXT_PUBLIC_FIREBASE_API_KEY / _AUTH_DOMAIN / _PROJECT_ID / _STORAGE_BUCKET /
#   _MESSAGING_SENDER_ID / _APP_ID        (from Firebase web-app config)
#   NEXT_PUBLIC_ENGINE_URL                (Cloud Run URL from step 3)
vercel --prod
```

Then:
1. Firebase console → **Authentication → Settings → Authorized domains** → add the Vercel
   production domain (and any custom domain) so Google sign-in works there.
2. Update the engine's `ALLOWED_ORIGINS` (end of step 3).
3. Optional: point `ggcunderwritingdemo.com` at the Vercel project (Vercel → Domains).

## 5. Smoke test

1. Open the Vercel URL → sign in with an allowlisted Google account.
2. Submit a small deal (one rent-roll PDF). Progress should stream; on completion the
   16-tab model downloads.
3. Sign in with a non-allowlisted account → engine calls must return 403.
4. `History` page lists the run; its Download button serves from Firebase Storage.
5. Verify a second browser/account cannot poll the first account's job id (404 expected).

## 6. SECURITY — do this before/with the first deploy

- **Rotate both API keys now.** This repo's history (per CLAUDE.md §12.4) has contained keys,
  and `.env` currently holds live ones. Create a NEW Anthropic key
  (console.anthropic.com → API keys) and a NEW Maps key, use those in Cloud Run, then
  **revoke the old ones**. Never reuse a key that was ever committed.
- **Restrict the Maps key** (Google Cloud console → Credentials): limit it to the Static Maps
  + Street View APIs.
- `.env`, `gcp-credentials.json` are already gitignored and dockerignored — keep them that
  way. The Cloud Run deployment intentionally uses **no** service-account key file.
- The legacy `index.html` UI served at the engine's `/` does not send auth headers, so with
  `REQUIRE_AUTH=1` it can render but its API calls 401. That's correct — the hosted UI is the
  Vercel app. Local dev (`REQUIRE_AUTH` unset) keeps the legacy UI fully working.

## Environment variable matrix

| Var | Where | Purpose |
|---|---|---|
| `ANTHROPIC_API_KEY` | Cloud Run (+ local `.env`) | Claude API (all three stages run `claude-fable-5`) |
| `MODEL_EXTRACTION` / `MODEL_METHODOLOGY` / `MODEL_MARKET` | Cloud Run (optional) | Per-stage model override (default `claude-fable-5`) |
| `THINKING_EFFORT` | Cloud Run (optional) | `high` (default) or `max` for adaptive-thinking stages |
| `GCP_PROJECT_ID` / `GCP_LOCATION` / `GCP_LAYOUT_PROCESSOR_ID` | Cloud Run + local | Document AI parser |
| `GOOGLE_MAPS_API_KEY` | Cloud Run + local | Property imagery |
| `REQUIRE_AUTH` | Cloud Run = `1` | Firebase ID-token verification on the API |
| `ALLOWED_EMAILS` | Cloud Run | Comma-separated sign-in allowlist (empty = any signed-in user) |
| `FIREBASE_STORAGE_BUCKET` | Cloud Run | Enables Firestore mirroring + Storage upload of finished models |
| `FIREBASE_SERVICE_ACCOUNT_JSON` | only if NOT on Google infra | One-line SA JSON; on Cloud Run leave unset (ADC) |
| `ALLOWED_ORIGINS` | Cloud Run | CORS allowlist = the web app origin(s) |
| `PORT`, `JOBS_DIR`, `IMG_CACHE_DIR`, `EXTRACTION_CACHE_DIR` | set by Dockerfile/Cloud Run | Container plumbing |
| `NEXT_PUBLIC_FIREBASE_*` (6 vars) | Vercel + `web/.env.local` | Firebase web config (not secret) |
| `NEXT_PUBLIC_ENGINE_URL` | Vercel + `web/.env.local` | Cloud Run service URL |
