import bcrypt
import uuid
import os
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from fastapi import Request, HTTPException, status
from sqlalchemy.orm import Session
from models import User, Operator

SECRET_KEY = os.getenv("SECRET_KEY", "dev-secret-key-change-in-production")
SESSION_MAX_AGE = 8 * 3600  # 8 ore in secondi

_serializer = URLSafeTimedSerializer(SECRET_KEY)


# ── Password ──────────────────────────────────────────────────────────────────

def hash_password(password: str) -> str:
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()


def verify_password(password: str, hashed: str) -> bool:
    return bcrypt.checkpw(password.encode(), hashed.encode())


# ── Session cookie ────────────────────────────────────────────────────────────

def create_session_token(user_id: int) -> str:
    return _serializer.dumps({"user_id": user_id})


def decode_session_token(token: str) -> int | None:
    try:
        data = _serializer.loads(token, max_age=SESSION_MAX_AGE)
        return data["user_id"]
    except (BadSignature, SignatureExpired, KeyError):
        return None


def get_current_user(request: Request, db: Session) -> User | None:
    token = request.cookies.get("session")
    if not token:
        return None
    user_id = decode_session_token(token)
    if user_id is None:
        return None
    return db.query(User).filter(User.id == user_id).first()


def require_admin(request: Request, db: Session) -> User:
    user = get_current_user(request, db)
    if user is None or not user.is_admin:
        raise HTTPException(status_code=status.HTTP_302_FOUND, headers={"Location": "/login"})
    return user


# ── API Key ───────────────────────────────────────────────────────────────────

def generate_api_key() -> str:
    return str(uuid.uuid4()).replace("-", "")


def get_operator_by_api_key(api_key: str, db: Session) -> Operator | None:
    return db.query(Operator).filter(Operator.api_key == api_key).first()


def require_api_key(request: Request, db: Session) -> Operator:
    api_key = request.headers.get("X-API-Key")
    if not api_key:
        raise HTTPException(status_code=401, detail="API key mancante")
    operator = get_operator_by_api_key(api_key, db)
    if operator is None:
        raise HTTPException(status_code=401, detail="API key non valida")
    return operator
