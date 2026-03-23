from sqlalchemy import create_engine
from sqlalchemy.orm import DeclarativeBase, sessionmaker
import os
from dotenv import load_dotenv

load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./timetracking.db")

engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False},  # solo SQLite
)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)


class Base(DeclarativeBase):
    pass


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db():
    """Crea le tabelle e l'utente admin iniziale."""
    from models import User, Operator, DayEntry, ComuneService  # noqa: F401
    Base.metadata.create_all(bind=engine)

    db = SessionLocal()
    try:
        if db.query(User).count() == 0:
            from auth import hash_password
            admin_user = os.getenv("ADMIN_USERNAME", "admin")
            admin_pass = os.getenv("ADMIN_PASSWORD", "ChangeMe123!")
            db.add(User(username=admin_user, hashed_password=hash_password(admin_pass), is_admin=True))
            db.commit()
            print(f"[init] Utente admin creato: {admin_user}")
    finally:
        db.close()
