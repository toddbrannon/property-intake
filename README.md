# property-intake

AI property-PDF intake pipeline demo. A Japanese property document goes in; a standardized record,
an English summary and a row in `master.xlsx` come out. Fields the model wasn't sure about go to a
review queue for a person.

- `app.py` — web page (upload, review queue, download) + background watcher. What Render runs.
- `extract.py` — sends a PDF to Claude against the field list, validates, writes outputs.
- `watch.py` — processes anything in `inbox/`, moves it to `done/` or `failed/`.
- `fields.csv` — the list of fields to extract. Edit this to change what comes out.
- `render.yaml` — Render Blueprint: one Starter web service + 1 GB disk. New → Blueprint → this repo → Apply.
- `samples/` — three fictional broker PDFs. `demo_extractions/` — recorded answers so they work with no API key.

Settings are environment variables (see `.env.example`). `ANTHROPIC_API_KEY` is optional: without it the
samples work and other PDFs go to Failed. Full setup: see the Property Intake SOP and Render Deployment SOP.
