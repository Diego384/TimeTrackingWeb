#!/usr/bin/env python3
"""
Server MCP per TimeTracking – Cooperativa Oltre i Sogni.

Esegui con:
  python mcp_server.py

Configura in %APPDATA%\\Claude\\claude_desktop_config.json:
  {
    "mcpServers": {
      "timetracking": {
        "command": "python",
        "args": ["C:\\\\Users\\\\diego\\\\Documents\\\\CoopOltreiSogniApps\\\\TimeTrackingWebGit\\\\mcp_server.py"],
        "env": {
          "API_BASE_URL": "http://localhost:8000",
          "API_USERNAME": "admin",
          "API_PASSWORD": "ChangeMe123!"
        }
      }
    }
  }
"""

import os
import asyncio
import httpx
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp import types

API_BASE_URL = os.getenv("API_BASE_URL", "http://localhost:8000")
API_USERNAME = os.getenv("API_USERNAME", "admin")
API_PASSWORD = os.getenv("API_PASSWORD", "")

server = Server("timetracking")
_token: str | None = None


async def _get_token(client: httpx.AsyncClient) -> str:
    global _token
    if _token:
        return _token
    resp = await client.post(
        f"{API_BASE_URL}/api/v1/token",
        data={"username": API_USERNAME, "password": API_PASSWORD},
    )
    resp.raise_for_status()
    _token = resp.json()["access_token"]
    return _token


async def _api_get(path: str) -> dict | list:
    global _token
    async with httpx.AsyncClient() as client:
        token = await _get_token(client)
        resp = await client.get(
            f"{API_BASE_URL}{path}",
            headers={"Authorization": f"Bearer {token}"},
        )
        if resp.status_code == 401:
            _token = None
            token = await _get_token(client)
            resp = await client.get(
                f"{API_BASE_URL}{path}",
                headers={"Authorization": f"Bearer {token}"},
            )
        resp.raise_for_status()
        return resp.json()


async def _api_post(path: str, body: dict) -> dict:
    global _token
    async with httpx.AsyncClient() as client:
        token = await _get_token(client)
        resp = await client.post(
            f"{API_BASE_URL}{path}",
            json=body,
            headers={"Authorization": f"Bearer {token}"},
        )
        if resp.status_code == 401:
            _token = None
            token = await _get_token(client)
            resp = await client.post(
                f"{API_BASE_URL}{path}",
                json=body,
                headers={"Authorization": f"Bearer {token}"},
            )
        resp.raise_for_status()
        return resp.json()


async def _api_put(path: str, body: dict) -> dict:
    global _token
    async with httpx.AsyncClient() as client:
        token = await _get_token(client)
        resp = await client.put(
            f"{API_BASE_URL}{path}",
            json=body,
            headers={"Authorization": f"Bearer {token}"},
        )
        if resp.status_code == 401:
            _token = None
            token = await _get_token(client)
            resp = await client.put(
                f"{API_BASE_URL}{path}",
                json=body,
                headers={"Authorization": f"Bearer {token}"},
            )
        resp.raise_for_status()
        return resp.json()


async def _api_get_bytes(path: str) -> bytes:
    global _token
    async with httpx.AsyncClient() as client:
        token = await _get_token(client)
        resp = await client.get(
            f"{API_BASE_URL}{path}",
            headers={"Authorization": f"Bearer {token}"},
        )
        if resp.status_code == 401:
            _token = None
            token = await _get_token(client)
            resp = await client.get(
                f"{API_BASE_URL}{path}",
                headers={"Authorization": f"Bearer {token}"},
            )
        resp.raise_for_status()
        return resp.content


# ── Tool definitions ──────────────────────────────────────────────────────────

@server.list_tools()
async def list_tools() -> list[types.Tool]:
    return [
        types.Tool(
            name="list_operators",
            description="Restituisce la lista di tutti gli operatori con nome, cognome, email e data ultimo sync.",
            inputSchema={"type": "object", "properties": {}, "required": []},
        ),
        types.Tool(
            name="get_operator_detail",
            description="Dettaglio di un operatore: dati anagrafici e lista mesi con dati disponibili.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operator_id": {"type": "integer", "description": "ID dell'operatore"},
                },
                "required": ["operator_id"],
            },
        ),
        types.Tool(
            name="get_monthly_report",
            description="Report mensile di un operatore: ore giornaliere, servizi per comune e totali.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operator_id": {"type": "integer", "description": "ID dell'operatore"},
                    "year": {"type": "integer", "description": "Anno (es. 2024)"},
                    "month": {"type": "integer", "description": "Mese 1-12"},
                },
                "required": ["operator_id", "year", "month"],
            },
        ),
        types.Tool(
            name="create_operator",
            description="Crea un nuovo operatore con nome, cognome, cooperativa ed email.",
            inputSchema={
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "Nome dell'operatore"},
                    "surname": {"type": "string", "description": "Cognome dell'operatore"},
                    "cooperative": {"type": "string", "description": "Nome della cooperativa (default: Cooperativa Sociale Oltre i sogni)"},
                    "email": {"type": "string", "description": "Email dell'operatore (opzionale)"},
                },
                "required": ["name", "surname"],
            },
        ),
        types.Tool(
            name="update_operator",
            description="Modifica i dati anagrafici di un operatore esistente (nome, cognome, cooperativa, email).",
            inputSchema={
                "type": "object",
                "properties": {
                    "operator_id": {"type": "integer", "description": "ID dell'operatore da modificare"},
                    "name": {"type": "string", "description": "Nuovo nome (opzionale)"},
                    "surname": {"type": "string", "description": "Nuovo cognome (opzionale)"},
                    "cooperative": {"type": "string", "description": "Nuova cooperativa (opzionale)"},
                    "email": {"type": "string", "description": "Nuova email (opzionale)"},
                },
                "required": ["operator_id"],
            },
        ),
        types.Tool(
            name="download_excel_report",
            description="Genera il report Excel mensile di un operatore e lo salva localmente.",
            inputSchema={
                "type": "object",
                "properties": {
                    "operator_id": {"type": "integer", "description": "ID dell'operatore"},
                    "year": {"type": "integer", "description": "Anno"},
                    "month": {"type": "integer", "description": "Mese 1-12"},
                    "save_path": {
                        "type": "string",
                        "description": "Percorso locale dove salvare il file .xlsx (opzionale)",
                    },
                },
                "required": ["operator_id", "year", "month"],
            },
        ),
    ]


# ── Tool handlers ─────────────────────────────────────────────────────────────

@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[types.TextContent]:

    if name == "list_operators":
        data = await _api_get("/api/v1/operators")
        if not data:
            return [types.TextContent(type="text", text="Nessun operatore trovato.")]
        lines = []
        for op in data:
            sync = op.get("last_sync") or "mai"
            lines.append(
                f"- ID {op['id']}: {op['surname']} {op['name']} "
                f"({op['email'] or 'nessuna email'}) — ultimo sync: {sync}"
            )
        return [types.TextContent(type="text", text="\n".join(lines))]

    elif name == "create_operator":
        body = {
            "name": arguments["name"],
            "surname": arguments["surname"],
            "cooperative": arguments.get("cooperative", "Cooperativa Sociale Oltre i sogni"),
            "email": arguments.get("email", ""),
        }
        data = await _api_post("/api/v1/operators", body)
        text = (
            f"Operatore creato con successo!\n"
            f"ID: {data['id']}\n"
            f"Nome: {data['surname']} {data['name']}\n"
            f"Cooperativa: {data['cooperative']}\n"
            f"Email: {data['email'] or 'nessuna'}"
        )
        return [types.TextContent(type="text", text=text)]

    elif name == "update_operator":
        op_id = arguments["operator_id"]
        body = {k: v for k, v in arguments.items() if k != "operator_id"}
        data = await _api_put(f"/api/v1/operators/{op_id}", body)
        text = (
            f"Operatore aggiornato con successo!\n"
            f"ID: {data['id']}\n"
            f"Nome: {data['surname']} {data['name']}\n"
            f"Cooperativa: {data['cooperative']}\n"
            f"Email: {data['email'] or 'nessuna'}"
        )
        return [types.TextContent(type="text", text=text)]

    elif name == "get_operator_detail":
        op_id = arguments["operator_id"]
        data = await _api_get(f"/api/v1/operators/{op_id}")
        months = data.get("months_with_data", [])
        month_list = ", ".join(f"{m[1]}/{m[0]}" for m in months) if months else "nessun dato"
        text = (
            f"Operatore: {data['surname']} {data['name']}\n"
            f"Email: {data['email'] or 'nessuna'}\n"
            f"Cooperativa: {data['cooperative']}\n"
            f"Ultimo sync: {data.get('last_sync') or 'mai'}\n"
            f"Mesi con dati: {month_list}"
        )
        return [types.TextContent(type="text", text=text)]

    elif name == "get_monthly_report":
        op_id = arguments["operator_id"]
        year = arguments["year"]
        month = arguments["month"]
        data = await _api_get(f"/api/v1/operators/{op_id}/report/{year}/{month}")

        op = data["operator"]
        tot = data["totals"]
        entries = data["entries"]
        comuni = data["comuni"]

        lines = [
            f"Report {month}/{year} — {op['surname']} {op['name']}",
            "",
            "TOTALI:",
            f"  Ore Memofast:       {tot['ore_memofast']}",
            f"  Ore Pulmino:        {tot['ore_pulmino']}",
            f"  Ore Sostituzioni:   {tot['ore_sostituzioni']}",
            f"  Ferie (ore):        {tot['ore_ferie_ore']}",
            f"  Ferie (giorni):     {tot['ore_ferie_giorni']}",
            f"  Malattia (giorni):  {tot['ore_malattia_giorni']}",
            f"  Legge 104:          {tot['ore_legge104']}",
            f"  TOTALE COMPLESSIVO: {tot['totale_complessivo']}",
            "",
        ]

        giorni_con_dati = [
            e for e in entries
            if any(e.get(k, 0) != 0 for k in
                   ["ore_memofast", "ore_pulmino", "ore_sostituzioni",
                    "ore_ferie", "ore_malattia", "ore_legge104"])
        ]
        if giorni_con_dati:
            lines.append(f"DETTAGLIO GIORNI ({len(giorni_con_dati)} giorni con attività):")
            for e in giorni_con_dati:
                parts = []
                if e["ore_memofast"]:   parts.append(f"memofast={e['ore_memofast']}")
                if e["ore_pulmino"]:    parts.append(f"pulmino={e['ore_pulmino']}")
                if e["ore_sostituzioni"]: parts.append(f"sost={e['ore_sostituzioni']}")
                if e["ore_ferie"]:      parts.append(f"ferie={e['ore_ferie']}")
                if e["ore_malattia"]:   parts.append(f"malattia={e['ore_malattia']}")
                if e["ore_legge104"]:   parts.append(f"l104={e['ore_legge104']}")
                if e.get("nota"):       parts.append(f"nota={e['nota']}")
                lines.append(f"  {e['date']}: {', '.join(parts)}")
        else:
            lines.append("Nessun giorno con attività registrata.")

        if comuni:
            lines.append("")
            lines.append("SERVIZI PER COMUNE:")
            for c in comuni:
                parts = []
                for campo in ["adi", "ada", "adh", "adm", "asia", "asia_istituti", "cpf"]:
                    if c.get(campo):
                        parts.append(f"{campo.upper()}={c[campo]}")
                lines.append(f"  {c['comune']}: {', '.join(parts) if parts else '-'}")

        return [types.TextContent(type="text", text="\n".join(lines))]

    elif name == "download_excel_report":
        op_id = arguments["operator_id"]
        year = arguments["year"]
        month = arguments["month"]
        save_path = arguments.get("save_path", f"report_{op_id}_{year}_{month:02d}.xlsx")

        content = await _api_get_bytes(
            f"/api/v1/operators/{op_id}/report/{year}/{month}/excel"
        )
        with open(save_path, "wb") as f:
            f.write(content)

        return [types.TextContent(type="text", text=f"File Excel salvato in: {save_path}")]

    else:
        raise ValueError(f"Tool sconosciuto: {name}")


# ── Entry point ───────────────────────────────────────────────────────────────

async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


if __name__ == "__main__":
    asyncio.run(main())
