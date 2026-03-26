from fastapi import FastAPI, Request, Depends, Form, HTTPException, Response, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
from sqlalchemy import extract
from datetime import datetime, timezone, date as date_type
import io
import json
import os
import uuid
import calendar
import qrcode
from pathlib import Path
from qrcode.image.pure import PyPNGImage

from database import get_db, init_db
from models import User, Operator, DayEntry, ComuneService, ContractHours, OperatorFile, WeeklySchedule, WeeklyScheduleEntry

UPLOAD_DIR = Path("uploads")
UPLOAD_DIR.mkdir(exist_ok=True)
from schemas import SyncPayload, SyncResponse, OperatorCreate, ContractHoursIn
from auth import (
    verify_password, create_session_token, get_current_user,
    require_admin, require_api_key, generate_api_key, hash_password
)
from excel_export import generate_excel
from schedule_excel import generate_schedule_excel
from api_v1 import router as api_v1_router

app = FastAPI(title="TimeTracking – Cooperativa Oltre i Sogni")
app.mount("/static", StaticFiles(directory="static"), name="static")
app.include_router(api_v1_router)
templates = Jinja2Templates(directory="templates")

MESI_IT = ["", "Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno",
           "Luglio", "Agosto", "Settembre", "Ottobre", "Novembre", "Dicembre"]


@app.on_event("startup")
def startup():
    init_db()


# ── Helpers template ──────────────────────────────────────────────────────────

def _fmt(v: float) -> str:
    if v == 0:
        return "-"
    return str(int(v)) if v % 1 == 0 else f"{v:.1f}"


def _ctx(db: Session, current_user=None, **kwargs):
    return {"current_user": current_user, "fmt": _fmt, "mesi": MESI_IT, **kwargs}


# ── Auth ──────────────────────────────────────────────────────────────────────

@app.get("/", response_class=RedirectResponse)
def root():
    return RedirectResponse("/dashboard")


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request, db: Session = Depends(get_db)):
    if get_current_user(request, db):
        return RedirectResponse("/dashboard")
    return templates.TemplateResponse(request, "login.html", {"error": None})


@app.post("/login")
def login(request: Request, response: Response,
          username: str = Form(...), password: str = Form(...),
          db: Session = Depends(get_db)):
    user = db.query(User).filter(User.username == username).first()
    if not user or not verify_password(password, user.hashed_password):
        return templates.TemplateResponse(request, "login.html",
                                          {"error": "Credenziali non valide"})
    token = create_session_token(user.id)
    resp = RedirectResponse("/dashboard", status_code=302)
    resp.set_cookie("session", token, httponly=True, max_age=8 * 3600)
    return resp


@app.get("/logout")
def logout():
    resp = RedirectResponse("/login", status_code=302)
    resp.delete_cookie("session")
    return resp


# ── Dashboard ─────────────────────────────────────────────────────────────────

@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    operators = db.query(Operator).order_by(Operator.surname).all()
    # Ultimi 10 sync
    recent = (db.query(DayEntry, Operator)
              .join(Operator)
              .order_by(DayEntry.synced_at.desc())
              .limit(10).all())
    recent_syncs = [{"operator": op, "entry": e} for e, op in recent]
    return templates.TemplateResponse(request, "dashboard.html",
                                      _ctx(db, current_user=user, operators=operators, recent_syncs=recent_syncs))


# ── Operatori ─────────────────────────────────────────────────────────────────

@app.get("/operators", response_class=HTMLResponse)
def operators_list(request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    operators = db.query(Operator).order_by(Operator.surname).all()
    return templates.TemplateResponse(request, "operators.html",
                                      _ctx(db, current_user=user, operators=operators))


@app.post("/operators/create")
def create_operator(request: Request,
                    name: str = Form(...), surname: str = Form(...),
                    cooperative: str = Form("Cooperativa Sociale Oltre i sogni"),
                    email: str = Form(""),
                    db: Session = Depends(get_db)):
    require_admin(request, db)
    op = Operator(name=name, surname=surname, cooperative=cooperative,
                  email=email, api_key=generate_api_key())
    db.add(op)
    db.commit()
    return RedirectResponse(f"/operators/{op.id}", status_code=302)


@app.get("/operators/{op_id}", response_class=HTMLResponse)
def operator_detail(op_id: int, request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")

    # Mesi con dati
    rows = (db.query(extract("year", DayEntry.date).label("y"),
                     extract("month", DayEntry.date).label("m"))
            .filter(DayEntry.operator_id == op_id)
            .distinct().order_by("y", "m").all())
    months_with_data = [(int(r.y), int(r.m)) for r in rows]

    return templates.TemplateResponse(request, "operator_detail.html",
                                      _ctx(db, current_user=user, op=op,
                                           months_with_data=months_with_data,
                                           contract_hours=op.contract_hours))


@app.post("/operators/{op_id}/regenerate-key")
def regenerate_key(op_id: int, request: Request, db: Session = Depends(get_db)):
    require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)
    op.api_key = generate_api_key()
    db.commit()
    return RedirectResponse(f"/operators/{op_id}", status_code=302)


@app.get("/operators/{op_id}/qrcode")
def operator_qrcode(op_id: int, request: Request, db: Session = Depends(get_db)):
    """Genera un QR code PNG con URL server + API Key per configurare l'app mobile."""
    require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)

    # Usa PUBLIC_BASE_URL se configurato (es. http://1.2.3.4:8080), altrimenti request.base_url
    base_url = os.getenv("PUBLIC_BASE_URL", "").rstrip("/") or str(request.base_url).rstrip("/")
    payload = json.dumps({"url": base_url, "api_key": op.api_key}, ensure_ascii=False)

    img = qrcode.make(payload)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return StreamingResponse(buf, media_type="image/png")


@app.post("/operators/{op_id}/delete")
def delete_operator(op_id: int, request: Request, db: Session = Depends(get_db)):
    require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if op:
        db.delete(op)
        db.commit()
    return RedirectResponse("/operators", status_code=302)


@app.post("/operators/{op_id}/contract-hours")
def save_contract_hours(
    op_id: int, request: Request,
    lunedi: float = Form(0), martedi: float = Form(0),
    mercoledi: float = Form(0), giovedi: float = Form(0),
    venerdi: float = Form(0), sabato: float = Form(0),
    domenica: float = Form(0),
    db: Session = Depends(get_db),
):
    require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")
    ch = db.query(ContractHours).filter(ContractHours.operator_id == op_id).first()
    if ch:
        ch.lunedi = lunedi
        ch.martedi = martedi
        ch.mercoledi = mercoledi
        ch.giovedi = giovedi
        ch.venerdi = venerdi
        ch.sabato = sabato
        ch.domenica = domenica
    else:
        db.add(ContractHours(
            operator_id=op_id,
            lunedi=lunedi, martedi=martedi, mercoledi=mercoledi,
            giovedi=giovedi, venerdi=venerdi, sabato=sabato, domenica=domenica,
        ))
    db.commit()
    return RedirectResponse(f"/operators/{op_id}", status_code=302)


# ── Report mensile ────────────────────────────────────────────────────────────

@app.get("/operators/{op_id}/report/{year}/{month}", response_class=HTMLResponse)
def monthly_report(op_id: int, year: int, month: int,
                   request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)

    days_in_month = calendar.monthrange(year, month)[1]
    all_dates = [datetime(year, month, d).date() for d in range(1, days_in_month + 1)]

    entries_raw = (db.query(DayEntry)
                   .filter(DayEntry.operator_id == op_id,
                           extract("year", DayEntry.date) == year,
                           extract("month", DayEntry.date) == month)
                   .all())
    entry_map = {e.date: e for e in entries_raw}

    comuni = (db.query(ComuneService)
              .filter(ComuneService.operator_id == op_id,
                      ComuneService.year == year,
                      ComuneService.month == month)
              .all())

    # Totali
    tot = {k: sum(getattr(e, k) for e in entries_raw)
           for k in ["ore_memofast", "ore_pulmino", "ore_sostituzioni",
                     "ore_malattia", "ore_legge104"]}
    # Ferie: separare giornate intere (-1) da ore specifiche
    tot_ferie_giorni = sum(1 for e in entries_raw if e.ore_ferie == -1.0)
    tot_ferie_ore    = sum(e.ore_ferie for e in entries_raw if e.ore_ferie > 0)
    tot["ore_ferie"] = tot_ferie_ore  # solo ore (per calcoli numerici nel template)
    if tot_ferie_giorni > 0 and tot_ferie_ore > 0:
        tot["ore_ferie_label"] = f"{tot_ferie_giorni}G + {_fmt(tot_ferie_ore)}h"
    elif tot_ferie_giorni > 0:
        tot["ore_ferie_label"] = f"{tot_ferie_giorni} G"
    elif tot_ferie_ore > 0:
        tot["ore_ferie_label"] = f"{_fmt(tot_ferie_ore)} h"
    else:
        tot["ore_ferie_label"] = "–"
    # Malattia: sempre giornate
    tot_malattia_giorni = sum(1 for e in entries_raw if e.ore_malattia > 0)
    tot["ore_malattia"] = 0.0  # non in ore
    tot["ore_malattia_label"] = f"{tot_malattia_giorni} G" if tot_malattia_giorni > 0 else "–"

    return templates.TemplateResponse(request, "report.html",
                                      _ctx(db, current_user=user, op=op, year=year, month=month,
                                           month_name=MESI_IT[month],
                                           all_dates=all_dates, entry_map=entry_map,
                                           comuni=comuni, tot=tot))


@app.get("/operators/{op_id}/report/{year}/{month}/excel")
def download_excel(op_id: int, year: int, month: int,
                   request: Request, db: Session = Depends(get_db)):
    require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)

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
    month_name = MESI_IT[month]
    filename = f"Report_{op.surname}_{month_name}_{year}.xlsx"
    return StreamingResponse(
        io.BytesIO(buf),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


# ── API mobile ────────────────────────────────────────────────────────────────

@app.post("/api/sync", response_model=SyncResponse)
def api_sync(payload: SyncPayload, request: Request, db: Session = Depends(get_db)):
    operator = require_api_key(request, db)

    # Aggiorna info operatore
    operator.name = payload.operator.name
    operator.surname = payload.operator.surname
    operator.cooperative = payload.operator.cooperative
    operator.email = payload.operator.email
    operator.last_sync = datetime.now(timezone.utc)

    # Upsert day entries
    for de in payload.day_entries:
        existing = (db.query(DayEntry)
                    .filter(DayEntry.operator_id == operator.id,
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
                operator_id=operator.id,
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
    for cs in payload.comune_services:
        existing = (db.query(ComuneService)
                    .filter(ComuneService.operator_id == operator.id,
                            ComuneService.year == payload.year,
                            ComuneService.month == payload.month,
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
                operator_id=operator.id,
                year=payload.year,
                month=payload.month,
                comune=cs.comune,
                adi=cs.adi, ada=cs.ada, adh=cs.adh, adm=cs.adm,
                asia=cs.asia, asia_istituti=cs.asia_istituti, cpf=cs.cpf,
            ))

    db.commit()
    return SyncResponse(status="ok", operator_id=operator.id,
                        synced_entries=len(payload.day_entries),
                        synced_comuni=len(payload.comune_services))


@app.get("/api/report/{year}/{month}")
def api_download_report(year: int, month: int,
                        request: Request, db: Session = Depends(get_db)):
    """Scarica il proprio report Excel via API key."""
    operator = require_api_key(request, db)

    entries = (db.query(DayEntry)
               .filter(DayEntry.operator_id == operator.id,
                       extract("year", DayEntry.date) == year,
                       extract("month", DayEntry.date) == month)
               .all())
    comuni = (db.query(ComuneService)
              .filter(ComuneService.operator_id == operator.id,
                      ComuneService.year == year,
                      ComuneService.month == month)
              .all())

    buf = generate_excel(operator, year, month, entries, comuni)
    filename = f"Report_{operator.surname}_{MESI_IT[month]}_{year}.xlsx"
    return StreamingResponse(
        io.BytesIO(buf),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.get("/api/contract-hours")
def api_get_contract_hours(request: Request, db: Session = Depends(get_db)):
    """Restituisce le ore ordinarie contrattuali dell'operatore via API key (uso mobile)."""
    operator = require_api_key(request, db)
    ch = db.query(ContractHours).filter(ContractHours.operator_id == operator.id).first()
    if not ch:
        return {
            "operator_id": operator.id,
            "lunedi": 0, "martedi": 0, "mercoledi": 0,
            "giovedi": 0, "venerdi": 0, "sabato": 0, "domenica": 0,
            "updated_at": None,
        }
    return {
        "operator_id": operator.id,
        "lunedi": ch.lunedi,
        "martedi": ch.martedi,
        "mercoledi": ch.mercoledi,
        "giovedi": ch.giovedi,
        "venerdi": ch.venerdi,
        "sabato": ch.sabato,
        "domenica": ch.domenica,
        "updated_at": ch.updated_at,
    }


# ── Admin: gestione utenti admin ──────────────────────────────────────────────

@app.get("/settings", response_class=HTMLResponse)
def settings_page(request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    users = db.query(User).all()
    return templates.TemplateResponse(request, "settings.html",
                                      _ctx(db, current_user=user, users=users))


@app.post("/settings/change-password")
def change_password(request: Request,
                    current_password: str = Form(...),
                    new_password: str = Form(...),
                    db: Session = Depends(get_db)):
    user = require_admin(request, db)
    if not verify_password(current_password, user.hashed_password):
        return templates.TemplateResponse(request, "settings.html",
                                          _ctx(db, current_user=user,
                                               users=db.query(User).all(),
                                               error="Password attuale non corretta"))
    user.hashed_password = hash_password(new_password)
    db.commit()
    return RedirectResponse("/settings?ok=1", status_code=302)


# ── API mobile: file upload/download ──────────────────────────────────────────

def _file_json(f: OperatorFile) -> dict:
    return {
        "id": f.id,
        "filename": f.filename,
        "mime_type": f.mime_type,
        "file_size": f.file_size,
        "uploaded_by": f.uploaded_by,
        "description": f.description,
        "created_at": f.created_at.isoformat() if f.created_at else None,
    }


@app.get("/api/files")
def api_list_files(request: Request, db: Session = Depends(get_db)):
    """Lista file dell'operatore (mobile)."""
    operator = require_api_key(request, db)
    files = (db.query(OperatorFile)
             .filter(OperatorFile.operator_id == operator.id)
             .order_by(OperatorFile.created_at.desc())
             .all())
    return [_file_json(f) for f in files]


@app.post("/api/files/upload")
async def api_upload_file(
    request: Request,
    file: UploadFile = File(...),
    description: str = Form(""),
    db: Session = Depends(get_db),
):
    """Carica un file dal dispositivo mobile."""
    operator = require_api_key(request, db)
    content = await file.read()
    ext = Path(file.filename or "file").suffix
    stored_name = f"{uuid.uuid4().hex}{ext}"
    (UPLOAD_DIR / stored_name).write_bytes(content)
    op_file = OperatorFile(
        operator_id=operator.id,
        filename=file.filename or stored_name,
        stored_name=stored_name,
        mime_type=file.content_type or "",
        file_size=len(content),
        uploaded_by="operator",
        description=description,
    )
    db.add(op_file)
    db.commit()
    db.refresh(op_file)
    return _file_json(op_file)


@app.get("/api/files/{file_id}/download")
def api_download_file(file_id: int, request: Request, db: Session = Depends(get_db)):
    """Scarica un file (mobile)."""
    operator = require_api_key(request, db)
    f = db.query(OperatorFile).filter(
        OperatorFile.id == file_id,
        OperatorFile.operator_id == operator.id,
    ).first()
    if not f:
        raise HTTPException(404, "File non trovato")
    path = UPLOAD_DIR / f.stored_name
    if not path.exists():
        raise HTTPException(404, "File non presente sul server")
    return FileResponse(
        str(path),
        media_type=f.mime_type or "application/octet-stream",
        filename=f.filename,
    )


@app.delete("/api/files/{file_id}")
def api_delete_file(file_id: int, request: Request, db: Session = Depends(get_db)):
    """Elimina un file (mobile, solo quelli caricati dall'operatore)."""
    operator = require_api_key(request, db)
    f = db.query(OperatorFile).filter(
        OperatorFile.id == file_id,
        OperatorFile.operator_id == operator.id,
    ).first()
    if not f:
        raise HTTPException(404, "File non trovato")
    if f.uploaded_by != "operator":
        raise HTTPException(403, "Non puoi eliminare file caricati dall'admin")
    path = UPLOAD_DIR / f.stored_name
    if path.exists():
        path.unlink()
    db.delete(f)
    db.commit()
    return {"status": "deleted"}


# ── Admin web: gestione file operatore ────────────────────────────────────────

@app.get("/operators/{op_id}/files", response_class=HTMLResponse)
def admin_files_list(op_id: int, request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)
    files = (db.query(OperatorFile)
             .filter(OperatorFile.operator_id == op_id)
             .order_by(OperatorFile.created_at.desc())
             .all())
    return templates.TemplateResponse(request, "operator_files.html",
                                      _ctx(db, current_user=user, op=op, files=files))


@app.post("/operators/{op_id}/files/upload")
async def admin_upload_file(
    op_id: int,
    request: Request,
    file: UploadFile = File(...),
    description: str = Form(""),
    db: Session = Depends(get_db),
):
    require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)
    content = await file.read()
    ext = Path(file.filename or "file").suffix
    stored_name = f"{uuid.uuid4().hex}{ext}"
    (UPLOAD_DIR / stored_name).write_bytes(content)
    op_file = OperatorFile(
        operator_id=op_id,
        filename=file.filename or stored_name,
        stored_name=stored_name,
        mime_type=file.content_type or "",
        file_size=len(content),
        uploaded_by="admin",
        description=description,
    )
    db.add(op_file)
    db.commit()
    return RedirectResponse(f"/operators/{op_id}/files", status_code=302)


@app.get("/operators/{op_id}/files/{file_id}/download")
def admin_download_file(op_id: int, file_id: int, request: Request,
                        db: Session = Depends(get_db)):
    require_admin(request, db)
    f = db.query(OperatorFile).filter(
        OperatorFile.id == file_id,
        OperatorFile.operator_id == op_id,
    ).first()
    if not f:
        raise HTTPException(404)
    path = UPLOAD_DIR / f.stored_name
    if not path.exists():
        raise HTTPException(404, "File non presente sul server")
    return FileResponse(
        str(path),
        media_type=f.mime_type or "application/octet-stream",
        filename=f.filename,
    )


@app.post("/operators/{op_id}/files/{file_id}/delete")
def admin_delete_file(op_id: int, file_id: int, request: Request,
                      db: Session = Depends(get_db)):
    require_admin(request, db)
    f = db.query(OperatorFile).filter(
        OperatorFile.id == file_id,
        OperatorFile.operator_id == op_id,
    ).first()
    if not f:
        raise HTTPException(404)
    path = UPLOAD_DIR / f.stored_name
    if path.exists():
        path.unlink()
    db.delete(f)
    db.commit()
    return RedirectResponse(f"/operators/{op_id}/files", status_code=302)


# ── API mobile: griglia oraria settimanale ─────────────────────────────────────

def _schedule_entry_json(e: WeeklyScheduleEntry) -> dict:
    return {
        "id": e.id,
        "day_of_week": e.day_of_week,
        "row_index": e.row_index,
        "ora_inizio": e.ora_inizio,
        "ora_fine": e.ora_fine,
        "ore": e.ore,
        "utente_assistito": e.utente_assistito,
        "servizio": e.servizio,
        "comune": e.comune,
    }


@app.post("/api/weekly-schedule")
def api_upsert_weekly_schedule(request: Request, body: dict,
                                db: Session = Depends(get_db)):
    """Upsert griglia oraria settimanale (mobile)."""
    operator = require_api_key(request, db)

    week_start_str = body.get("week_start")
    if not week_start_str:
        raise HTTPException(400, "week_start obbligatorio")
    try:
        week_start = date_type.fromisoformat(week_start_str)
    except ValueError:
        raise HTTPException(400, "week_start non valido (usa YYYY-MM-DD)")

    periodo = body.get("periodo_riferimento", "")
    entries_data = body.get("entries", [])

    # Upsert schedule
    schedule = (db.query(WeeklySchedule)
                .filter(WeeklySchedule.operator_id == operator.id,
                        WeeklySchedule.week_start == week_start)
                .first())
    if schedule:
        schedule.periodo_riferimento = periodo
        schedule.updated_at = datetime.now(timezone.utc)
        # Elimina entries esistenti
        db.query(WeeklyScheduleEntry).filter(
            WeeklyScheduleEntry.schedule_id == schedule.id
        ).delete()
    else:
        schedule = WeeklySchedule(
            operator_id=operator.id,
            week_start=week_start,
            periodo_riferimento=periodo,
        )
        db.add(schedule)
        db.flush()

    # Inserisce nuove entries
    for ed in entries_data:
        db.add(WeeklyScheduleEntry(
            schedule_id=schedule.id,
            day_of_week=ed.get("day_of_week", 1),
            row_index=ed.get("row_index", 0),
            ora_inizio=ed.get("ora_inizio", ""),
            ora_fine=ed.get("ora_fine", ""),
            ore=ed.get("ore", 0.0),
            utente_assistito=ed.get("utente_assistito", ""),
            servizio=ed.get("servizio", ""),
            comune=ed.get("comune", ""),
        ))

    db.commit()
    db.refresh(schedule)
    return {"status": "ok", "id": schedule.id, "week_start": str(week_start)}


@app.get("/api/weekly-schedule")
def api_list_weekly_schedules(request: Request, db: Session = Depends(get_db)):
    """Lista griglie orarie dell'operatore (mobile)."""
    operator = require_api_key(request, db)
    schedules = (db.query(WeeklySchedule)
                 .filter(WeeklySchedule.operator_id == operator.id)
                 .order_by(WeeklySchedule.week_start.desc())
                 .all())
    return [
        {
            "id": s.id,
            "week_start": str(s.week_start),
            "periodo_riferimento": s.periodo_riferimento,
            "updated_at": s.updated_at.isoformat() if s.updated_at else None,
            "entry_count": len(s.entries),
        }
        for s in schedules
    ]


@app.get("/api/weekly-schedule/{week_start_str}")
def api_get_weekly_schedule(week_start_str: str, request: Request,
                             db: Session = Depends(get_db)):
    """Restituisce griglia oraria per una settimana (mobile)."""
    operator = require_api_key(request, db)
    try:
        week_start = date_type.fromisoformat(week_start_str)
    except ValueError:
        raise HTTPException(400, "week_start non valido (usa YYYY-MM-DD)")

    schedule = (db.query(WeeklySchedule)
                .filter(WeeklySchedule.operator_id == operator.id,
                        WeeklySchedule.week_start == week_start)
                .first())
    if not schedule:
        raise HTTPException(404, "Griglia non trovata")

    return {
        "id": schedule.id,
        "week_start": str(schedule.week_start),
        "periodo_riferimento": schedule.periodo_riferimento,
        "updated_at": schedule.updated_at.isoformat() if schedule.updated_at else None,
        "entries": [_schedule_entry_json(e) for e in schedule.entries],
    }


@app.delete("/api/weekly-schedule/{week_start_str}")
def api_delete_weekly_schedule(week_start_str: str, request: Request,
                                db: Session = Depends(get_db)):
    """Elimina griglia oraria per una settimana (mobile)."""
    operator = require_api_key(request, db)
    try:
        week_start = date_type.fromisoformat(week_start_str)
    except ValueError:
        raise HTTPException(400, "week_start non valido (usa YYYY-MM-DD)")

    schedule = (db.query(WeeklySchedule)
                .filter(WeeklySchedule.operator_id == operator.id,
                        WeeklySchedule.week_start == week_start)
                .first())
    if not schedule:
        raise HTTPException(404, "Griglia non trovata")
    db.delete(schedule)
    db.commit()
    return {"status": "deleted"}


# ── Admin web: griglie orarie operatore ───────────────────────────────────────

@app.get("/operators/{op_id}/schedules", response_class=HTMLResponse)
def admin_schedules_list(op_id: int, request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404, "Operatore non trovato")
    schedules = (db.query(WeeklySchedule)
                 .filter(WeeklySchedule.operator_id == op_id)
                 .order_by(WeeklySchedule.week_start.desc())
                 .all())
    return templates.TemplateResponse(request, "operator_schedules.html",
                                      _ctx(db, current_user=user, op=op, schedules=schedules))


@app.get("/operators/{op_id}/schedules/{week_start_str}", response_class=HTMLResponse)
def admin_schedule_detail(op_id: int, week_start_str: str,
                          request: Request, db: Session = Depends(get_db)):
    user = require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)
    try:
        week_start = date_type.fromisoformat(week_start_str)
    except ValueError:
        raise HTTPException(400, "Data non valida")

    schedule = (db.query(WeeklySchedule)
                .filter(WeeklySchedule.operator_id == op_id,
                        WeeklySchedule.week_start == week_start)
                .first())
    if not schedule:
        raise HTTPException(404, "Griglia non trovata")

    # Raggruppa entries per giorno
    entries_by_day = {}
    for e in schedule.entries:
        entries_by_day.setdefault(e.day_of_week, []).append(e)

    # Totali per giorno e settimana
    totali_giorno = {dow: sum(e.ore for e in lst)
                     for dow, lst in entries_by_day.items()}
    totale_settimana = sum(totali_giorno.values())

    giorni_nomi = {1: "Lunedì", 2: "Martedì", 3: "Mercoledì",
                   4: "Giovedì", 5: "Venerdì", 6: "Sabato"}

    return templates.TemplateResponse(request, "operator_schedule_detail.html",
                                      _ctx(db, current_user=user, op=op,
                                           schedule=schedule,
                                           entries_by_day=entries_by_day,
                                           totali_giorno=totali_giorno,
                                           totale_settimana=totale_settimana,
                                           giorni_nomi=giorni_nomi))


@app.post("/operators/{op_id}/schedules/{week_start_str}/delete")
def admin_delete_schedule(op_id: int, week_start_str: str,
                          request: Request, db: Session = Depends(get_db)):
    require_admin(request, db)
    try:
        week_start = date_type.fromisoformat(week_start_str)
    except ValueError:
        raise HTTPException(400, "Data non valida")

    schedule = (db.query(WeeklySchedule)
                .filter(WeeklySchedule.operator_id == op_id,
                        WeeklySchedule.week_start == week_start)
                .first())
    if schedule:
        db.delete(schedule)
        db.commit()
    return RedirectResponse(f"/operators/{op_id}/schedules", status_code=302)


@app.get("/operators/{op_id}/schedules/{week_start_str}/excel")
def admin_schedule_excel(op_id: int, week_start_str: str,
                         request: Request, db: Session = Depends(get_db)):
    require_admin(request, db)
    op = db.query(Operator).filter(Operator.id == op_id).first()
    if not op:
        raise HTTPException(404)
    try:
        week_start = date_type.fromisoformat(week_start_str)
    except ValueError:
        raise HTTPException(400, "Data non valida")

    schedule = (db.query(WeeklySchedule)
                .filter(WeeklySchedule.operator_id == op_id,
                        WeeklySchedule.week_start == week_start)
                .first())
    if not schedule:
        raise HTTPException(404, "Griglia non trovata")

    entries_by_day = {}
    for e in schedule.entries:
        entries_by_day.setdefault(e.day_of_week, []).append(e)

    buf = generate_schedule_excel(op, schedule, entries_by_day)
    filename = f"GrigliaOraria_{op.surname}_{week_start_str}.xlsx"
    return StreamingResponse(
        io.BytesIO(buf),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
