from fastapi import FastAPI, Depends
from backend.api.auth import router as auth_router, require_role
from backend.api.routes import router as api_router  # seu router principal
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from pathlib import Path

import os
from fastapi import FastAPI
from backend.core.database import SessionLocal, engine
from backend.core.models import Base, User
from backend.core.security import hash_password

from .api.routes import router as api_router
from .core.excel_repo import DataStore

from backend.core.database import engine
from backend.core.models import Base

from backend.api import users as users_api

app = FastAPI(title="Dimensionador Solar - API Local", version="0.1.0")

Base.metadata.create_all(bind=engine)

def seed_admin():
    admin_user = os.getenv("ADMIN_USER", "admin")
    admin_pass = os.getenv("ADMIN_PASS")  # não coloque default aqui
    admin_role = "admin"

    # Se não definiu ADMIN_PASS, não cria ninguém (segurança)
    if not admin_pass:
        return

    db = SessionLocal()
    try:
        # já existe admin?
        exists_admin = db.query(User).filter(User.role == "admin").first()
        if exists_admin:
            return

        # cria admin
        u = User(
            username=admin_user,
            hashed_password=hash_password(admin_pass),
            role=admin_role,
            is_active=True,
        )
        db.add(u)
        db.commit()
    finally:
        db.close()

@app.on_event("startup")
def on_startup():
    seed_admin()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.state.store = DataStore()

# API primeiro (pra não ser "engolida" pelo static)
app.include_router(auth_router, prefix="/api")

app.include_router(
    api_router,
    prefix="/api",
    dependencies=[Depends(require_role("admin", "engenharia"))]
)

app.include_router(users_api.router, prefix="/api")

# Servir frontend no "/"
FRONTEND_DIR = Path(__file__).resolve().parents[1] / "frontend"
app.mount("/", StaticFiles(directory=FRONTEND_DIR, html=True), name="frontend")