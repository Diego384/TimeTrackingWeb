from sqlalchemy import Column, Integer, String, Float, Boolean, Date, DateTime, ForeignKey, UniqueConstraint
from sqlalchemy.orm import relationship
from datetime import datetime, timezone
from database import Base


class User(Base):
    __tablename__ = "users"

    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, nullable=False)
    hashed_password = Column(String, nullable=False)
    is_admin = Column(Boolean, default=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))


class Operator(Base):
    __tablename__ = "operators"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    surname = Column(String, nullable=False)
    cooperative = Column(String, default="Cooperativa Sociale Oltre i sogni")
    email = Column(String, default="")
    api_key = Column(String, unique=True, nullable=False, index=True)
    created_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))
    last_sync = Column(DateTime, nullable=True)

    day_entries = relationship("DayEntry", back_populates="operator", cascade="all, delete-orphan")
    comune_services = relationship("ComuneService", back_populates="operator", cascade="all, delete-orphan")
    contract_hours = relationship("ContractHours", back_populates="operator", uselist=False, cascade="all, delete-orphan")

    @property
    def full_name(self):
        return f"{self.name} {self.surname}"


class DayEntry(Base):
    __tablename__ = "day_entries"

    id = Column(Integer, primary_key=True, index=True)
    operator_id = Column(Integer, ForeignKey("operators.id"), nullable=False)
    date = Column(Date, nullable=False)
    ore_memofast = Column(Float, default=0)
    ore_pulmino = Column(Float, default=0)
    ore_sostituzioni = Column(Float, default=0)
    ore_ferie = Column(Float, default=0)
    ore_malattia = Column(Float, default=0)
    ore_legge104 = Column(Float, default=0)
    nota = Column(String, default="")
    synced_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

    operator = relationship("Operator", back_populates="day_entries")

    __table_args__ = (
        UniqueConstraint("operator_id", "date", name="uq_operator_date"),
    )


class ComuneService(Base):
    __tablename__ = "comune_services"

    id = Column(Integer, primary_key=True, index=True)
    operator_id = Column(Integer, ForeignKey("operators.id"), nullable=False)
    year = Column(Integer, nullable=False)
    month = Column(Integer, nullable=False)
    comune = Column(String, nullable=False)
    adi = Column(Float, default=0)
    ada = Column(Float, default=0)
    adh = Column(Float, default=0)
    adm = Column(Float, default=0)
    asia = Column(Float, default=0)
    asia_istituti = Column(Float, default=0)
    cpf = Column(Float, default=0)
    synced_at = Column(DateTime, default=lambda: datetime.now(timezone.utc))

    operator = relationship("Operator", back_populates="comune_services")

    __table_args__ = (
        UniqueConstraint("operator_id", "year", "month", "comune", name="uq_operator_comune_month"),
    )

    @property
    def totale(self):
        return self.adi + self.ada + self.adh + self.adm + self.asia + self.asia_istituti + self.cpf


class ContractHours(Base):
    """Ore ordinarie giornaliere stabilite da contratto per ogni giorno della settimana."""
    __tablename__ = "contract_hours"

    id = Column(Integer, primary_key=True, index=True)
    operator_id = Column(Integer, ForeignKey("operators.id"), unique=True, nullable=False)
    lunedi = Column(Float, default=0)
    martedi = Column(Float, default=0)
    mercoledi = Column(Float, default=0)
    giovedi = Column(Float, default=0)
    venerdi = Column(Float, default=0)
    sabato = Column(Float, default=0)
    domenica = Column(Float, default=0)
    updated_at = Column(DateTime, default=lambda: datetime.now(timezone.utc),
                        onupdate=lambda: datetime.now(timezone.utc))

    operator = relationship("Operator", back_populates="contract_hours")
