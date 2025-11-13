# ASCUE Kompas - Python port (full package)

This repo contains a cross-platform Python port skeleton of the ascue_kompas project.

Structure:
- core/: core logic and DB helpers
- desktop/: PySide6 desktop application and settings dialog
- web/: FastAPI application and config endpoints
- config.example.json: example DB config (kept with original addresses as requested)

Configuration:
Copy `config.example.json` to `config.json` in the repo root (or set env VOVKA21_CONFIG to point to a file).

Desktop quick start:
1. python -m venv venv
2. source venv/bin/activate  # or venv\Scripts\activate on Windows
3. pip install -r requirements.txt
4. python desktop/main.py

Web quick start:
1. pip install -r requirements.txt
2. uvicorn web.app:app --reload
3. GET http://127.0.0.1:8000/api/config/db to view current DB settings

Security note:
The config example includes real addresses per your request. Ensure you protect `config.json` and server secrets appropriately.
