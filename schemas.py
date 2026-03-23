from pydantic import BaseModel, Field
from datetime import date, datetime
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


class OperatorUpdate(BaseModel):
    name: Optional[str] = None
    surname: Optional[str] = None
    cooperative: Optional[str] = None
    email: Optional[str] = None


# ── API v1 — output/input schemas ────────────────────────────────────────────

class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"


class OperatorOut(BaseModel):
    id: int
    name: str
    surname: str
    cooperative: str
    email: str
    last_sync: Optional[datetime] = None

    model_config = {"from_attributes": True}


class OperatorDetailOut(OperatorOut):
    months_with_data: list[tuple[int, int]] = []


class DayEntryOut(BaseModel):
    date: date
    ore_memofast: float
    ore_pulmino: float
    ore_sostituzioni: float
    ore_ferie: float
    ore_malattia: float
    ore_legge104: float
    nota: str

    model_config = {"from_attributes": True}


class ComuneServiceOut(BaseModel):
    comune: str
    adi: float
    ada: float
    adh: float
    adm: float
    asia: float
    asia_istituti: float
    cpf: float

    model_config = {"from_attributes": True}


class ReportTotals(BaseModel):
    ore_memofast: float
    ore_pulmino: float
    ore_sostituzioni: float
    ore_ferie_ore: float
    ore_ferie_giorni: int
    ore_malattia_giorni: int
    ore_legge104: float
    totale_complessivo: float


class MonthlyReportOut(BaseModel):
    operator: OperatorOut
    year: int
    month: int
    entries: list[DayEntryOut]
    comuni: list[ComuneServiceOut]
    totals: ReportTotals


class OperatorDataUpsert(BaseModel):
    year: int = Field(ge=2020, le=2100)
    month: int = Field(ge=1, le=12)
    day_entries: list[DayEntryIn] = []
    comune_services: list[ComuneServiceIn] = []
