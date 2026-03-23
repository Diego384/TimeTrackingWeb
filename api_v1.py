from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.responses import StreamingResponse
from fastapi.security import OAuth2PasswordRequestForm
from sqlalchemy.orm import Session
from sqlalchemy import extract
from datetime import datetime, timezone
import calendar
import io

from database import get_db
from models import User, Operator, DayEntry, ComuneService
from auth import verify_password, create_access_token, get_current_user_jwt, generate_api_key
from excel_export import generate_excel
from schemas import (
    TokenResponse, OperatorOut, OperatorDetailOut,
    DayEntryOut, ComuneServiceOut, MonthlyReportOut, ReportTotals,
    OperatorCreate, OperatorUpdate, OperatorDataUpsert,
)

MESI_IT = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
           "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]

router = APIRouter(prefix="/api/v1", tags=["API v1"])


# ── Auth ──────────────────────────────────────────────────────────────────────

@router.post("/token", response_model=TokenResponse)
def login_for_access_token(
    form_data: OAuth2PasswordRequestForm = Depends(),
    db: Session = Depends(get_db),
):
    user = db.query(User).filter(User.username == form_data.username).first()
    if not user or not verify_password(form_data.password, user.hashed_password):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Username o password non corretti",
            headers={"WWW-Authenticate": "Bearer"},
        )
    token = create_access_token({"sub": user.username})
    return TokenResponse(access_token=token)


# ── Operatori ─────────────────────────────────────────────────────────────────

@router.get("/operators", response_model=list[OperatorOut])
def list_operators(
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_jwt),
):
    return db.query(Operator).order_by(Operator.surname).all()


@router.post("/operators", response_model=OperatorOut, status_code=status.HTTP_201_CREATED)
def create_operator(
    body: OperatorCreate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_jwt),
):
    op = Operator(
        name=body.name,
        surname=body.surname,
        cooperative=body.cooperative,
        email=body.email,
        api_key=generate_api_key(),
    )
    db.add(op)
    db.commit()
    db.refresh(op)
    return op


@router.put("/operators/{op_id}", response_model=OperatorOut)
def update_operator(
    op_id: int,
    body: OperatorUpdate,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_jwt),
):
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")
    if body.name is not None:
        op.name = body.name
    if body.surname is not None:
        op.surname = body.surname
    if body.cooperative is not None:
        op.cooperative = body.cooperative
    if body.email is not None:
        op.email = body.email
    db.commit()
    db.refresh(op)
    return op


@router.get("/operators/{op_id}", response_model=OperatorDetailOut)
def get_operator(
    op_id: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_jwt),
):
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")

    rows = (db.query(
                extract("year", DayEntry.date).label("y"),
                extract("month", DayEntry.date).label("m"))
            .filter(DayEntry.operator_id == op_id)
            .distinct().order_by("y", "m").all())

    return OperatorDetailOut(
        id=op.id,
        name=op.name,
        surname=op.surname,
        cooperative=op.cooperative,
        email=op.email,
        last_sync=op.last_sync,
        months_with_data=[(int(r.y), int(r.m)) for r in rows],
    )


# ── Report mensile ────────────────────────────────────────────────────────────

@router.get("/operators/{op_id}/report/{year}/{month}", response_model=MonthlyReportOut)
def get_monthly_report(
    op_id: int, year: int, month: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_jwt),
):
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")

    entries_raw = (db.query(DayEntry)
                   .filter(DayEntry.operator_id == op_id,
                           extract("year", DayEntry.date) == year,
                           extract("month", DayEntry.date) == month)
                   .order_by(DayEntry.date)
                   .all())

    comuni = (db.query(ComuneService)
              .filter(ComuneService.operator_id == op_id,
                      ComuneService.year == year,
                      ComuneService.month == month)
              .all())

    # Calcolo totali (stessa logica di main.py)
    ore_memofast = sum(e.ore_memofast for e in entries_raw)
    ore_pulmino = sum(e.ore_pulmino for e in entries_raw)
    ore_sostituzioni = sum(e.ore_sostituzioni for e in entries_raw)
    ore_legge104 = sum(e.ore_legge104 for e in entries_raw)
    ore_ferie_giorni = sum(1 for e in entries_raw if e.ore_ferie == -1.0)
    ore_ferie_ore = sum(e.ore_ferie for e in entries_raw if e.ore_ferie > 0)
    ore_malattia_giorni = sum(1 for e in entries_raw if e.ore_malattia > 0)
    totale = ore_memofast + ore_pulmino + ore_sostituzioni + ore_ferie_ore + ore_legge104

    totals = ReportTotals(
        ore_memofast=ore_memofast,
        ore_pulmino=ore_pulmino,
        ore_sostituzioni=ore_sostituzioni,
        ore_ferie_ore=ore_ferie_ore,
        ore_ferie_giorni=ore_ferie_giorni,
        ore_malattia_giorni=ore_malattia_giorni,
        ore_legge104=ore_legge104,
        totale_complessivo=totale,
    )

    return MonthlyReportOut(
        operator=OperatorOut.model_validate(op),
        year=year,
        month=month,
        entries=[DayEntryOut.model_validate(e) for e in entries_raw],
        comuni=[ComuneServiceOut.model_validate(c) for c in comuni],
        totals=totals,
    )


@router.get("/operators/{op_id}/report/{year}/{month}/excel")
def download_excel_report(
    op_id: int, year: int, month: int,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_jwt),
):
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")

    entries = (db.query(DayEntry)
               .filter(DayEntry.operator_id == op_id,
                       extract("year", DayEntry.date) == year,
                       extract("month", DayEntry.date) == month)
               .all())
    comuni = (db.query(ComuneService)
              .filter(ComuneService.operator_id == op_id,
                      ComuneService.year == year,
                      ComuneService.month == month)
              .all())

    buf = generate_excel(op, year, month, entries, comuni)
    filename = f"Report_{op.surname}_{MESI_IT[month]}_{year}.xlsx"
    return StreamingResponse(
        io.BytesIO(buf),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── Upsert dati operatore ─────────────────────────────────────────────────────

@router.put("/operators/{op_id}/entries")
def upsert_operator_entries(
    op_id: int,
    body: OperatorDataUpsert,
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user_jwt),
):
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")

    # Upsert day entries
    for de in body.day_entries:
        existing = (db.query(DayEntry)
                    .filter(DayEntry.operator_id == op_id,
                            DayEntry.date == de.date)
                    .first())
        if existing:
            existing.ore_memofast = de.ore_memofast
            existing.ore_pulmino = de.ore_pulmino
            existing.ore_sostituzioni = de.ore_sostituzioni
            existing.ore_ferie = de.ore_ferie
            existing.ore_malattia = de.ore_malattia
            existing.ore_legge104 = de.ore_legge104
            existing.nota = de.nota
            existing.synced_at = datetime.now(timezone.utc)
        else:
            db.add(DayEntry(
                operator_id=op_id,
                date=de.date,
                ore_memofast=de.ore_memofast,
                ore_pulmino=de.ore_pulmino,
                ore_sostituzioni=de.ore_sostituzioni,
                ore_ferie=de.ore_ferie,
                ore_malattia=de.ore_malattia,
                ore_legge104=de.ore_legge104,
                nota=de.nota,
            ))

    # Upsert comune services
    for cs in body.comune_services:
        existing = (db.query(ComuneService)
                    .filter(ComuneService.operator_id == op_id,
                            ComuneService.year == body.year,
                            ComuneService.month == body.month,
                            ComuneService.comune == cs.comune)
                    .first())
        if existing:
            existing.adi = cs.adi
            existing.ada = cs.ada
            existing.adh = cs.adh
            existing.adm = cs.adm
            existing.asia = cs.asia
            existing.asia_istituti = cs.asia_istituti
            existing.cpf = cs.cpf
            existing.synced_at = datetime.now(timezone.utc)
        else:
            db.add(ComuneService(
                operator_id=op_id,
                year=body.year,
                month=body.month,
                comune=cs.comune,
                adi=cs.adi, ada=cs.ada, adh=cs.adh, adm=cs.adm,
                asia=cs.asia, asia_istituti=cs.asia_istituti, cpf=cs.cpf,
            ))

    db.commit()
    return {
        "upserted_entries": len(body.day_entries),
        "upserted_comuni": len(body.comune_services),
    }
