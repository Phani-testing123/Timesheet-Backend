# backend/app/main.py
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from dotenv import load_dotenv
from pathlib import Path
import os

# --- Load env files ---
# 1) backend/.env (where you keep EXCEL_TEMPLATE_PATH etc.)
backend_env = Path(__file__).resolve().parents[1] / ".env"
if backend_env.exists():
    load_dotenv(dotenv_path=backend_env, override=False)

# 2) repo/.env (optional fallback if you also keep one at the repo root)
repo_env = Path(__file__).resolve().parents[2] / ".env"
if repo_env.exists():
    load_dotenv(dotenv_path=repo_env, override=False)

app = FastAPI(title="Timesheet Tool API", version="0.1.0")

# --- CORS ---
# Default to permissive for local dev; override with CORS_ALLOW_ORIGINS if you want stricter.
# Example to lock to Next dev: CORS_ALLOW_ORIGINS=http://localhost:3000
cors_origins_env = os.getenv("CORS_ALLOW_ORIGINS", "*")
allow_origins = [o.strip() for o in cors_origins_env.split(",")] if cors_origins_env else ["*"]
allow_credentials = os.getenv("CORS_ALLOW_CREDENTIALS", "0").strip() == "1"

app.add_middleware(
    CORSMiddleware,
    allow_origins=allow_origins if allow_origins else ["*"],
    allow_credentials=allow_credentials,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Routers ---
from app.routes import health, exports  # noqa: E402

app.include_router(health.router)
app.include_router(exports.router, prefix="/exports", tags=["exports"])

# --- Simple health endpoint (in addition to your health router) ---
@app.get("/healthz")
def healthz():
    return {"ok": True}
