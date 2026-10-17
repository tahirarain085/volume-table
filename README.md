# Tree enumeration register

Flask app to record trees by beat, km / RD / burji no, kind and girth. Dia, cft and
shi unit come from the volume tables in `volume_data.py`. Data is stored in MySQL
(Aiven). Excel export has one sheet per beat plus abstracts.

## Files
- `app.py` – the app (routes, database, Excel export)
- `volume_data.py` – volume tables (edit numbers here)
- `templates/index.html` – the page
- `.env.example` – copy to `.env` and fill in
- `requirements.txt`, `Procfile` – for installing and deploying

## 1. Run it locally (no database needed)
```
python -m venv .venv
.venv\Scripts\activate          # Windows   (Mac/Linux: source .venv/bin/activate)
pip install -r requirements.txt
python app.py
```
Open http://127.0.0.1:5000. With no database settings it uses a local SQLite file,
so you can test the whole app first.

## 2. Connect to Aiven MySQL
1. In Aiven open your MySQL service, Overview page.
2. Download the **CA certificate** and save it as `ca.pem` next to `app.py`.
3. Copy `.env.example` to `.env` and fill it in. Easiest: paste the **Service URI**
   into `DATABASE_URL`. Or fill `DB_HOST`, `DB_PORT`, `DB_USERNAME`, `DB_PASSWORD`, `DB_NAME`.
   Copy the host exactly. It starts with `mysql-` (one m).
4. If you restricted **Allowed IP addresses** in Aiven, add your own IP.
5. Run `python app.py` and open http://127.0.0.1:5000/health/db.
   You should see `{"database": "connected", "engine": "mysql"}`.

Tables are created automatically on first start.

## 3. Deploy (Render, Railway, etc.)
- Start command: `gunicorn app:app`
- Set the same variables from `.env` in the host's environment settings.
- No file upload for `ca.pem`? Paste its text into `DB_SSL_CA_PEM`.

## Rules in the app
- The same tree no can be reused in another km, but not twice in the same km of the same beat.
- Girth must exist in the volume table for the chosen kind.

## Troubleshooting
| Message | Fix |
|---|---|
| Name or service not known | Host is mistyped. Copy it again from Aiven. |
| Access denied | Wrong user or password. |
| Can't connect / timed out | Wrong port, or your IP is not allowed in Aiven. |
| SSL / certificate error | `ca.pem` is missing or from another service. |
# tree_enumberaion
