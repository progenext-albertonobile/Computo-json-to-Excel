# Computo XLSX Generator

Internal web app for generating branded Excel (XLSX) files from a JSON bundle,
using an existing Python generation engine.

## Stack

- **Backend**: Python 3.11 + FastAPI + Uvicorn
- **Frontend**: Single-page HTML + vanilla JS (no framework)
- **Engine**: `genera_computo.py` (XLSX builder), `validate_bundle.py` (capacity validator)

## Directory layout

```
playbook_industrial/
├── main.py               # FastAPI app entry point
├── genera_computo.py     # XLSX generation engine (do not modify)
├── validate_bundle.py    # Capacity/overflow validator (do not modify)
├── requirements.txt
├── static/
│   └── index.html        # Single-page UI
└── templates/
    └── Computo preliminare-V3.xlsx   # Default branded template
```

## Local run

### 1. Install Python 3.11

```bash
python3.11 --version   # confirm
```

### 2. Install dependencies

```bash
cd playbook_industrial
pip install -r requirements.txt
```

### 3. Start the server

```bash
python main.py
# or:
uvicorn main:app --host 0.0.0.0 --port 8000
```

Open [http://localhost:8000](http://localhost:8000) in your browser.

## Environment variables

| Variable | Default | Description |
|----------|---------|-------------|
| `PORT`   | `8000`  | HTTP port the server listens on |

## API endpoints

| Method | Path | Description |
|--------|------|-------------|
| `GET`  | `/`  | Serves the single-page UI |
| `POST` | `/validate` | Validates JSON + optional template, returns section/overflow report |
| `POST` | `/generate` | Validates then generates and streams XLSX back as download |

Both POST endpoints accept `multipart/form-data` with:
- `json_file` (required) — `.json` input bundle
- `template_file` (optional) — `.xlsx` template override (falls back to server default)

## Behavior

1. Upload JSON. Optionally upload an alternative XLSX template.
2. Click **Validate & Generate**.
3. The app runs the capacity check first and displays:
   - Per-section: items / capacity / overflow
   - External items: count / capacity / overflow
4. If **any** section or external overflow > 0, generation is **blocked** with a clear error.
5. If validation passes, the XLSX is generated and a download button appears.

## Upload limits

- JSON: 5 MB max
- Template XLSX: 20 MB max
- File type: `.json` / `.xlsx` only (enforced server-side)

## Deploy notes (private / internal)

This app is intended for **internal use only**. All pages carry `noindex, nofollow`
meta tags. No authentication is built in — deploy behind a VPN, reverse proxy with
IP allowlist, or basic-auth (e.g. nginx, Caddy) before exposing externally.

### Replit

Set the `PORT` secret if you need a non-default port. The app reads it automatically.

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
    # Add auth_basic here for internal access control
}
```

### Data persistence

No data is persisted by default. All uploaded files are processed in memory or
temporary directories that are deleted immediately after the response.
