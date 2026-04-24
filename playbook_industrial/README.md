# Computo XLSX Generator

Internal web app for generating branded Excel (XLSX) files from a typed-rows
JSON bundle, using a fully data-driven Python engine (no fixed template).

## Stack

- **Backend**: Python 3.11 + FastAPI + Uvicorn
- **Frontend**: Single-page HTML + vanilla JS (no framework)
- **Engine**: `genera_computo.py` — `build(data)` builds the workbook from scratch

## Directory layout

```
playbook_industrial/
├── main.py               # FastAPI app entry point
├── genera_computo.py     # XLSX generation engine (data-driven, typed rows)
├── requirements.txt
└── static/
    └── index.html        # Single-page UI
```

## JSON input shape

```json
{
  "metadata": {
    "cliente": "ROSSI LOGISTICA S.r.l.",
    "indirizzo": "Via della Meccanica 12, Bentivoglio (BO)",
    "data": "Aprile 2026",
    "titolo": "RIQUALIFICAZIONE CENTRALE TERMICA",
    "note_header": "Prezzi al cliente finale — IVA esclusa"
  },
  "rows": [
    { "type": "sezione",          "label": "CAPITOLO 1 — DEMOLIZIONI" },
    { "type": "sottosezione",     "label": "SMANTELLAMENTO IMPIANTO" },
    { "type": "voce_demolizione", "desc": "...", "qty": 1, "pu": 2800, ... },
    { "type": "voce",             "desc": "...", "qty": 1, "pu": 1850, ... },
    { "type": "voce_trasporto",   "desc": "...", "qty": 1, "pu": 380, "ric": 0.1 },
    { "type": "sottotot_sezione", "label": "CAPITOLO 1" },
    { "type": "totale" },
    { "type": "esclusioni",       "label": "ESCLUSIONI" },
    { "type": "voce_esterna",     "desc": "...", "qty": 1, "pu": 3500, ... }
  ]
}
```

### Supported row types

| `type`              | Purpose                                   |
|---------------------|-------------------------------------------|
| `sezione`           | Section header (navy)                     |
| `sottosezione`      | Subsection header (light blue, optional `bg`) |
| `voce`              | Standard line item with full formulas (alternating white/gray) |
| `voce_highlight`    | Highlighted line item (green)             |
| `voce_demolizione`  | Demolition line item (pink)               |
| `voce_trasporto`    | Transport line, simplified (alt rows)     |
| `voce_lumpsum`      | Lump-sum price, hardcoded total in col R  |
| `voce_esterna`      | External / out-of-total line (orange)     |
| `sottotot_sezione`  | Auto-summing section subtotal (blue)      |
| `totale`            | Grand total of all subtotals (navy)       |
| `esclusioni`        | "Excluded items" header (navy)            |
| `riga_vuota`        | Spacer (6 pt)                             |

## Local run

```bash
cd playbook_industrial
pip install -r requirements.txt
python main.py
# or:
uvicorn main:app --host 0.0.0.0 --port 8000
```

Open [http://localhost:8000](http://localhost:8000).

## Environment variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT`   | `8000`  | HTTP port the server listens on |

## API endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET`  | `/`         | Single-page UI |
| `POST` | `/generate` | Accepts `multipart/form-data` with `json_file` (`.json`), returns the XLSX as download |

## Behavior

1. Upload the JSON bundle.
2. Click **Genera XLSX**.
3. The engine builds the Excel from scratch and streams it back as a download.

No data is persisted. JSON is processed in memory.

## Upload limits

- JSON: 5 MB max
- File type: `.json` only (enforced server-side)

## Deploy notes (private / internal)

This app is intended for **internal use only**. The UI carries `noindex, nofollow`
meta tags. No authentication is built in — deploy behind a VPN, reverse proxy
with IP allowlist, or basic-auth before exposing externally.

### Docker (example)

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY playbook_industrial/ .
RUN pip install --no-cache-dir -r requirements.txt
ENV PORT=8000
CMD ["python", "main.py"]
```

### Reverse proxy (nginx snippet)

```nginx
location /computo/ {
    proxy_pass http://127.0.0.1:8000/;
    proxy_set_header Host $host;
    proxy_set_header X-Real-IP $remote_addr;
    # auth_basic + auth_basic_user_file for internal access control
}
```
