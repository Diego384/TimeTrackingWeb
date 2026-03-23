from pydantic import BaseModel, Field
from datetime import date
from typing import Optional


# ── Sync API (mobile → web) ───────────────────────────────────────────────────

class OperatorInfo(BaseModel):
    name: str
    surname: str
    cooperative: str = "Cooperativa Sociale Oltre i sogni"
    email: str = ""


class DayEntryIn(BaseModel):
    date: date
    ore_memofast: float = 0
    ore_pulmino: float = 0
    ore_sostituzioni: float = 0
    ore_ferie: float = 0
    ore_malattia: float = 0
    ore_legge104: float = 0
    nota: str = ""


class ComuneServiceIn(BaseModel):
    comune: str
    adi: float = 0
    ada: float = 0
    adh: float = 0
    adm: float = 0
    asia: float = 0
    asia_istituti: float = 0
    cpf: float = 0


class SyncPayload(BaseModel):
    operator: OperatorInfo
    year: int = Field(ge=2020, le=2100)
    month: int = Field(ge=1, le=12)
    day_entries: list[DayEntryIn] = []
    comune_services: list[ComuneServiceIn] = []


class SyncResponse(BaseModel):
    status: str
    operator_id: int
    synced_entries: int
    synced_comuni: int


# ── Operatore manuale (form web) ──────────────────────────────────────────────

class OperatorCreate(BaseModel):
    name: str
    surname: str
    cooperative: str = "Cooperativa Sociale Oltre i sogni"
    email: str = ""
