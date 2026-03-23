FROM python:3.12-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Crea cartella per il DB persistente
RUN mkdir -p /data

# Variabili di ambiente default (sovrascrivibili con -e o docker-compose)
ENV DATABASE_URL=sqlite:////data/timetracking.db
ENV SECRET_KEY=change-this-in-production

EXPOSE 8000

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
