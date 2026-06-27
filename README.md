# The CU Library — PowerPoint Export API

Small FastAPI service that turns a board-package payload into a branded `.pptx`.
The CU Library app calls `POST /pptx` and downloads the result.

## Endpoints
- `GET /` — health check
- `POST /pptx` — body is the book payload (theme, period, headline, balanceSheet, incomeStatement, lending, ops, ceoHighlights); returns the `.pptx` file.

## Deploy on Railway
1. Push this folder to a new GitHub repo (e.g. `cu-library-pptx-api`).
2. In Railway: New Project -> Deploy from GitHub repo -> pick the repo.
3. Railway auto-installs `requirements.txt` and runs the `Procfile`
   (`uvicorn main:app --host 0.0.0.0 --port $PORT`). No extra config needed.
4. Copy the public URL Railway gives you (e.g. `https://cu-library-pptx-api-production.up.railway.app`).

## Point the app at it
In the Vercel project for the app, add an environment variable:

    NEXT_PUBLIC_PPTX_API_URL = https://<your-railway-url>

Redeploy. "Export to PowerPoint" now generates server-side. If the variable is
missing or the service is unreachable, the app falls back to in-browser export,
so nothing breaks during setup.

## Run locally
    pip install -r requirements.txt
    uvicorn main:app --reload --port 8000
