from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from sqlalchemy.orm import Session

from backend.core.database import get_db
from backend.core.models import User
from backend.core.security import verify_password, create_access_token, decode_token

router = APIRouter(prefix="/auth", tags=["auth"])
oauth2_scheme = OAuth2PasswordBearer(tokenUrl="/api/auth/token")


@router.post("/token")
def login(form: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == form.username).first()

    if not user or not user.is_active or not verify_password(form.password, user.hashed_password):
        raise HTTPException(status_code=401, detail="Usuário/senha inválidos")

    if user.role == "comercial":
        raise HTTPException(status_code=403, detail="Perfil COMERCIAL ainda não liberado para uso.")

    token = create_access_token(sub=user.username, role=user.role)
    return {"access_token": token, "token_type": "bearer", "role": user.role}


def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)) -> User:
    try:
        payload = decode_token(token)
    except Exception:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="Token inválido")

    username = payload.get("sub")
    if not username:
        raise HTTPException(status_code=401, detail="Token sem subject")

    user = db.query(User).filter(User.username == username).first()
    if not user or not user.is_active:
        raise HTTPException(status_code=401, detail="Usuário inválido/inativo")

    return user


def require_role(*roles: str):
    def _dep(user: User = Depends(get_current_user)) -> User:
        if user.role not in roles:
            raise HTTPException(status_code=403, detail="Sem permissão")
        return user
    return _dep


@router.get("/me")
def me(user: User = Depends(get_current_user)):
    return {"username": user.username, "role": user.role}