# GGC Deal Engine

AI underwriting tool for Gary Group Capital: seller financials in → parse → extract →
verify → GGC methodology → populated 16-tab Excel underwriting model out.

- `backend.py` — the whole analysis engine (Flask). Local: `python3 backend.py` → http://localhost:5001
- `web/` — hosted UI (Next.js + Firebase Auth, deploys to Vercel)
- `Dockerfile` — engine container for Cloud Run
- `CLAUDE.md` — what this system is and the rules it must follow (read first)
- `DEPLOYMENT.md` — how to host it (Cloud Run + Firebase + Vercel)

Models: every stage runs Claude Fable 5 (`claude-fable-5`); see CLAUDE.md for the
extraction/methodology split and override knobs.
