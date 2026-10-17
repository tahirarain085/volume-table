"""Tree enumeration register (Flask + SQLAlchemy + MySQL on Aiven).

Run locally:   python app.py
Production:    gunicorn app:app
"""
import os
import re
import ssl
import tempfile
from collections import defaultdict
from datetime import datetime
from io import BytesIO
from urllib.parse import parse_qsl, quote_plus, urlencode, urlsplit, urlunsplit

from dotenv import load_dotenv
from flask import (
    Flask, flash, redirect, render_template, request, send_file, session, url_for,
)
from flask_sqlalchemy import SQLAlchemy
from openpyxl import Workbook
from openpyxl.styles import Border, Font, Side
from sqlalchemy import func, inspect, text
from sqlalchemy.exc import IntegrityError, OperationalError

from volume_data import RANGES

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(BASE_DIR, ".env"))  # must run BEFORE any os.getenv


# ---------------------------------------------------------------------------
# Database configuration
#   1) DATABASE_URL  (paste the "Service URI" from Aiven), or
#   2) DB_USERNAME / DB_PASSWORD / DB_HOST / DB_PORT / DB_NAME
#   3) nothing set -> a local SQLite file, so the app still runs for testing
# ---------------------------------------------------------------------------
def _mysql_ssl_args():
    """SSL settings for Aiven. Uses ca.pem if present, else encrypts without verifying."""
    if os.getenv("DB_SSL", "true").strip().lower() in ("0", "false", "no", "off"):
        return {}
    ca_path = (os.getenv("DB_SSL_CA") or "").strip() or os.path.join(BASE_DIR, "ca.pem")
    pem_text = os.getenv("DB_SSL_CA_PEM")
    if pem_text and not os.path.exists(ca_path):
        ca_path = os.path.join(tempfile.gettempdir(), "aiven_ca.pem")
        with open(ca_path, "w") as f:
            f.write(pem_text.replace("\\n", "\n"))
    if os.path.exists(ca_path):
        return {"ssl": {"ca": ca_path}}
    print("[database] WARNING: ca.pem not found. Connection is encrypted but the server "
          "certificate is NOT verified. Download ca.pem from Aiven and put it next to app.py.")
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return {"ssl": ctx}


def build_database_config():
    """Returns (uri, engine_options)."""
    url = (os.getenv("DATABASE_URL") or "").strip()
    if url:
        url = re.sub(r"^mysql://", "mysql+pymysql://", url)
        parts = urlsplit(url)
        query = [(k, v) for k, v in parse_qsl(parts.query) if k.lower().replace("_", "-") != "ssl-mode"]
        if not any(k == "charset" for k, _ in query):
            query.append(("charset", "utf8mb4"))
        url = urlunsplit(parts._replace(query=urlencode(query)))
        mysql = True
    else:
        host = (os.getenv("DB_HOST") or "").strip()
        if not host:
            print("[database] No DB_HOST or DATABASE_URL set: using local SQLite file trees_local.db")
            return f"sqlite:///{os.path.join(BASE_DIR, 'trees_local.db')}", {}
        user = (os.getenv("DB_USERNAME") or "").strip()
        password = quote_plus(os.getenv("DB_PASSWORD") or "")
        port = (os.getenv("DB_PORT") or "3306").strip()
        name = (os.getenv("DB_NAME") or "").strip()
        url = f"mysql+pymysql://{user}:{password}@{host}:{port}/{name}?charset=utf8mb4"
        mysql = True

    options = {"pool_pre_ping": True, "pool_recycle": 280, "connect_args": _mysql_ssl_args()}
    return url, options


app = Flask(__name__)
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY", "change-me-in-production")
app.config["SQLALCHEMY_DATABASE_URI"], app.config["SQLALCHEMY_ENGINE_OPTIONS"] = build_database_config()
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)


# ---------------------------------------------------------------------------
# Beat names for the dropdown. Replace these with your real beat names.
# The user can pick one of them OR type a new name. Any new name that is saved
# is remembered and shown in the list next time automatically.
# ---------------------------------------------------------------------------
BEATS = [
    "Rajana Kamalia Road",
    "Toba Rajana Road",
    "Moongi Beat",
    "Rajana Pirmahal Road",
    "Rajana Samundri Road",
    "Bhagat Canal",
    "Pirmahal Beat 1"
]


def build_volume_table(ranges):
    """Expand RANGES into {(kind_tree, girth): {...}} and check for mistakes."""
    table = {}
    for i, row in enumerate(ranges, start=1):
        if len(row) != 6:
            raise ValueError(f"RANGES row {i} must have 6 values, found {len(row)}: {row}")

        kind, g_from, g_to, dia, cft, shi_unit = row
        if g_from > g_to:
            raise ValueError(f"RANGES row {i}: girth_from is bigger than girth_to: {row}")

        for girth in range(g_from, g_to + 1):
            key = (kind.strip().lower(), girth)
            if key in table:
                raise ValueError(f"RANGES row {i}: girth {girth} of {kind} is defined twice")
            table[key] = {"dia": dia, "cft": cft, "shi_unit": shi_unit}
    return table


VOLUME_TABLE = build_volume_table(RANGES)
KINDS = sorted({kind for kind, _ in VOLUME_TABLE})


# ---------------------------------------------------------------------------
# Tree no rule
# The same tree no CAN be used again in another km / RD no (or another beat),
# but it can NOT be saved twice in the same km / RD no of the same beat.
# Set TREE_NO_UNIQUE_PER_KIND = True if the same tree no may also be used once
# for each kind (for example shisham 14 and kikar 14) in the same km / RD no.
# ---------------------------------------------------------------------------
TREE_NO_UNIQUE_PER_KIND = False
UNIQUE_COLUMNS = (
    ["beat_name", "km_burji_no"]
    + (["kind_tree"] if TREE_NO_UNIQUE_PER_KIND else [])
    + ["tree_no"]
)
UNIQUE_NAME = "uq_tree_location"


# ---------------------------------------------------------------------------
# Database model
# ---------------------------------------------------------------------------
class Shisham(db.Model):
    __tablename__ = "shisham"

    id = db.Column(db.Integer, primary_key=True)
    tree_no = db.Column(db.String(50), nullable=False)
    beat_name = db.Column(db.String(100), nullable=False)
    km_burji_no = db.Column(db.String(50), nullable=False)
    kind_tree = db.Column(db.String(100), nullable=False)
    girth = db.Column(db.Integer, nullable=False)
    dia = db.Column(db.Integer)
    cft = db.Column(db.Float)
    shi_unit = db.Column(db.Float)
    remarks = db.Column(db.String(255))

    # Duplicate rule: see UNIQUE_COLUMNS above.
    __table_args__ = (db.UniqueConstraint(*UNIQUE_COLUMNS, name=UNIQUE_NAME),)

    def to_dict(self):
        return {
            "id": self.id,
            "tree_no": self.tree_no,
            "beat_name": self.beat_name,
            "km_burji_no": self.km_burji_no,
            "kind_tree": self.kind_tree,
            "girth": self.girth,
            "dia": self.dia,
            "cft": self.cft,
            "shi_unit": self.shi_unit,
            "remarks": self.remarks,
        }


def duplicate_message(tree_no, kind_tree, km_burji_no, beat):
    where = f"km / RD / burji no {km_burji_no} of {beat}"
    if TREE_NO_UNIQUE_PER_KIND:
        return (
            f"Tree no {tree_no} ({kind_tree.capitalize()}) is already saved in {where}. "
            "Use a different tree no."
        )
    return (
        f"Tree no {tree_no} is already saved in {where}. "
        "The same number can be used in another km, but not twice in the same km."
    )


def get_beat_options():
    """Beats from the BEATS list plus every beat name already saved."""
    saved = db.session.query(Shisham.beat_name).distinct().all()
    names = set(BEATS) | {name for (name,) in saved if name}
    return sorted(names, key=str.lower)


# Volume table as plain JSON, used by the page to preview dia / cft / shi unit live
VOLUME_JSON = {f"{kind}|{girth}": v for (kind, girth), v in VOLUME_TABLE.items()}


# ---------------------------------------------------------------------------
# Abstract: number of trees, total cft and total shi unit for each kind
# ---------------------------------------------------------------------------
def _new_abstract():
    return {"rows": {}, "trees": 0, "cft": 0.0, "shi_unit": 0.0}


def _add_to_abstract(a, kind, trees, cft, shi_unit):
    row = a["rows"].setdefault(kind, {"kind": kind, "trees": 0, "cft": 0.0, "shi_unit": 0.0})
    row["trees"] += trees
    row["cft"] += cft
    row["shi_unit"] += shi_unit
    a["trees"] += trees
    a["cft"] += cft
    a["shi_unit"] += shi_unit


def _finish_abstract(a):
    rows = sorted(a["rows"].values(), key=lambda r: r["kind"])
    for r in rows:
        r["cft"] = round(r["cft"], 3)
        r["shi_unit"] = round(r["shi_unit"], 3)
    return {
        "rows": rows,
        "trees": a["trees"],
        "cft": round(a["cft"], 3),
        "shi_unit": round(a["shi_unit"], 3),
    }


def get_abstracts(only_beat=None):
    """Returns ({beat: abstract}, grand_abstract). Uses ALL saved entries."""
    query = db.session.query(
        Shisham.beat_name,
        Shisham.kind_tree,
        func.count(Shisham.id),
        func.coalesce(func.sum(Shisham.cft), 0),
        func.coalesce(func.sum(Shisham.shi_unit), 0),
    )
    if only_beat:
        query = query.filter(Shisham.beat_name == only_beat)
    query = query.group_by(Shisham.beat_name, Shisham.kind_tree)

    per_beat, grand = {}, _new_abstract()
    for beat, kind, trees, cft, shi_unit in query.all():
        cft, shi_unit = float(cft), float(shi_unit)
        _add_to_abstract(per_beat.setdefault(beat, _new_abstract()), kind, trees, cft, shi_unit)
        _add_to_abstract(grand, kind, trees, cft, shi_unit)

    return {b: _finish_abstract(a) for b, a in per_beat.items()}, _finish_abstract(grand)


# ---------------------------------------------------------------------------
# Helpers for the tree form
# ---------------------------------------------------------------------------
def normalize_beat(name):
    """Trim spaces and reuse the spelling of an existing beat (case-insensitive)."""
    beat = " ".join((name or "").split())
    for known in get_beat_options():
        if known.lower() == beat.lower():
            return known
    return beat


def read_tree_form(form, beat, exclude_id=None):
    """Check the form. Returns (error, data); data is ready to save when error is None.

    exclude_id is the tree being edited, so it is not counted as its own duplicate.
    """
    tree_no = " ".join(form.get("tree_no", "").split())
    km_burji_no = " ".join(form.get("km_burji_no", "").split())
    kind_tree = form.get("kind_tree", "").strip().lower()
    girth_raw = form.get("girth", "").strip()
    remarks = form.get("remarks", "").strip()

    if not beat:
        return "Please set the beat name first.", None
    if not (tree_no and km_burji_no and kind_tree and girth_raw):
        return "Please enter the tree no, km / RD / burji no, select a tree kind and enter the girth.", None

    try:
        girth = int(girth_raw)
    except ValueError:
        return "Girth must be a whole number.", None

    row = VOLUME_TABLE.get((kind_tree, girth))
    if row is None:
        return f"No data found for {kind_tree} with girth {girth}.", None

    key = {
        "beat_name": beat,
        "km_burji_no": km_burji_no,
        "kind_tree": kind_tree,
        "tree_no": tree_no,
    }
    duplicate = Shisham.query.filter_by(**{col: key[col] for col in UNIQUE_COLUMNS})
    if exclude_id is not None:
        duplicate = duplicate.filter(Shisham.id != exclude_id)
    if duplicate.first():
        return duplicate_message(tree_no, kind_tree, km_burji_no, beat), None

    return None, {
        "tree_no": tree_no,
        "beat_name": beat,
        "km_burji_no": km_burji_no,
        "kind_tree": kind_tree,
        "girth": girth,
        "dia": row["dia"],
        "cft": row["cft"],
        "shi_unit": row["shi_unit"],
        "remarks": remarks or None,
    }


def render_index(error=None, form=None, edit_tree=None):
    """Build the page: one section per beat (current beat first, then A-Z)."""
    current_beat = session.get("beat_name")

    abstracts, grand = get_abstracts()
    by_beat = defaultdict(list)
    for t in Shisham.query.order_by(Shisham.id.asc()).all():
        by_beat[t.beat_name].append(t)

    beat_order = sorted(by_beat, key=lambda b: (b != current_beat, b.lower()))
    sections = [
        {"beat": b, "trees": by_beat[b], "abstract": abstracts[b]} for b in beat_order
    ]

    return render_template(
        "index.html",
        kinds=KINDS,
        beats=get_beat_options(),
        current_beat=current_beat,
        sections=sections,
        grand=grand,
        volume=VOLUME_JSON,
        error=error,
        form=form or {},
        edit_tree=edit_tree,
    )


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/set-beat", methods=["POST"])
def set_beat():
    """Choose the beat once. All entries that follow belong to this beat."""
    beat = normalize_beat(request.form.get("beat_name"))

    if not beat:
        flash("Please pick or type a beat name.", "error")
        return redirect(url_for("index"))

    session["beat_name"] = beat
    return redirect(url_for("index"))


@app.route("/", methods=["GET", "POST"])
def index():
    current_beat = session.get("beat_name")

    # ----- Add a new tree -----
    if request.method == "POST":
        error, data = read_tree_form(request.form, current_beat)

        if error is None:
            tree = Shisham(**data)
            db.session.add(tree)
            try:
                db.session.commit()
            except IntegrityError:
                db.session.rollback()
                error = duplicate_message(
                    data["tree_no"], data["kind_tree"], data["km_burji_no"], data["beat_name"]
                )
            else:
                flash(
                    f"Saved tree no {tree.tree_no} ({tree.km_burji_no}): "
                    f"{tree.kind_tree.capitalize()}, girth {tree.girth}, dia {tree.dia}, "
                    f"cft {tree.cft}, shi unit {tree.shi_unit}",
                    "success",
                )
                # Redirect so refreshing the page does not save a duplicate
                return redirect(url_for("index"))

        return render_index(error=error, form=request.form)

    # ----- Show the page, optionally with one tree loaded for editing (?edit=ID) -----
    edit_tree = None
    edit_id = request.args.get("edit", type=int)
    if edit_id:
        edit_tree = db.session.get(Shisham, edit_id)
        if edit_tree is None:
            flash("That tree no longer exists.", "error")
            return redirect(url_for("index"))

    return render_index(form=edit_tree.to_dict() if edit_tree else {}, edit_tree=edit_tree)


@app.route("/update/<int:tree_id>", methods=["POST"])
def update_tree(tree_id):
    """Save changes to an existing tree. dia, cft and shi unit are worked out again."""
    tree = db.get_or_404(Shisham, tree_id)
    beat = normalize_beat(request.form.get("beat_name")) or tree.beat_name

    error, data = read_tree_form(request.form, beat, exclude_id=tree.id)

    if error is None:
        for field, value in data.items():
            setattr(tree, field, value)
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            error = duplicate_message(
                data["tree_no"], data["kind_tree"], data["km_burji_no"], data["beat_name"]
            )
        else:
            flash(
                f"Updated tree no {tree.tree_no} ({tree.km_burji_no}) in {tree.beat_name}: "
                f"{tree.kind_tree.capitalize()}, girth {tree.girth}, dia {tree.dia}, "
                f"cft {tree.cft}, shi unit {tree.shi_unit}",
                "success",
            )
            return redirect(url_for("index"))

    db.session.rollback()
    return render_index(error=error, form=request.form, edit_tree=tree)


@app.route("/delete/<int:tree_id>", methods=["POST"])
def delete_tree(tree_id):
    tree = db.get_or_404(Shisham, tree_id)
    label = (
        f"tree no {tree.tree_no} ({tree.kind_tree.capitalize()}, km / burji no "
        f"{tree.km_burji_no}) from {tree.beat_name}"
    )
    db.session.delete(tree)
    db.session.commit()
    flash(f"Deleted {label}.", "success")
    return redirect(url_for("index"))


# ---------------------------------------------------------------------------
# Excel download: one separate sheet per beat, abstract at the bottom of each
# The beat name is written once at the top of its sheet.
# ---------------------------------------------------------------------------
EXCEL_HEADERS = [
    "ID", "Tree No", "Km / RD / Burji No", "Kind",
    "Girth", "Dia", "CFT", "Shi Unit", "Remarks",
]
EXCEL_WIDTHS = [14, 14, 16, 14, 10, 8, 10, 12, 40]
TOTAL_BORDER = Border(top=Side(style="double"))


def safe_sheet_name(name, used):
    """Excel sheet names: max 31 chars, no []:*?/\\ and must be unique."""
    base = re.sub(r"[\[\]:*?/\\]", "", name).strip()[:31] or "Beat"
    candidate, n = base, 2
    while candidate.lower() in used:
        suffix = f" ({n})"
        candidate = base[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def _bold_last_row(ws, border=None):
    for cell in ws[ws.max_row]:
        cell.font = Font(bold=True)
        if border:
            cell.border = border


def write_abstract_block(ws, abstract, title="Abstract"):
    ws.append([])
    ws.append([title])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
    ws.append(["No. of trees", "Kind", "CFT", "Shi unit"])
    _bold_last_row(ws)
    for r in abstract["rows"]:
        ws.append([r["trees"], r["kind"].capitalize(), r["cft"], r["shi_unit"]])
    ws.append([abstract["trees"], "Total", abstract["cft"], abstract["shi_unit"]])
    _bold_last_row(ws, TOTAL_BORDER)


def write_beat_sheet(ws, beat, trees, abstract):
    """Row 1: beat name. Row 2: column names. Then the trees, then the abstract."""
    ws.append([f"Beat: {beat}"])
    ws["A1"].font = Font(bold=True, size=14)

    ws.append(EXCEL_HEADERS)
    _bold_last_row(ws)

    for t in trees:
        ws.append([
            t.id, t.tree_no, t.km_burji_no, t.kind_tree.capitalize(),
            t.girth, t.dia, t.cft, t.shi_unit, t.remarks,
        ])

    write_abstract_block(ws, abstract)

    for col, width in zip("ABCDEFGHI", EXCEL_WIDTHS):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A3"


def write_grand_sheet(ws, grand, per_beat):
    ws.append(["Grand abstract: all beats"])
    ws["A1"].font = Font(bold=True, size=14)
    write_abstract_block(ws, grand, title="By kind")

    ws.append([])
    ws.append(["By beat"])
    ws.cell(row=ws.max_row, column=1).font = Font(bold=True, size=12)
    ws.append(["No. of trees", "Beat", "CFT", "Shi unit"])
    _bold_last_row(ws)
    for beat in sorted(per_beat, key=str.lower):
        a = per_beat[beat]
        ws.append([a["trees"], beat, a["cft"], a["shi_unit"]])
    ws.append([grand["trees"], "Total", grand["cft"], grand["shi_unit"]])
    _bold_last_row(ws, TOTAL_BORDER)

    for col, width in zip("ABCD", [14, 24, 12, 12]):
        ws.column_dimensions[col].width = width


@app.route("/download-excel")
def download_excel():
    """All beats (one sheet each). Use /download-excel?beat=Name for a single beat."""
    only_beat = request.args.get("beat")

    query = Shisham.query
    if only_beat:
        query = query.filter_by(beat_name=only_beat)
    trees = query.order_by(Shisham.beat_name, Shisham.id).all()
    per_beat, grand = get_abstracts(only_beat)

    wb = Workbook()
    first_sheet = wb.active

    beats = sorted(per_beat, key=str.lower)
    if not beats:
        first_sheet.title = "No Data"
    used = set()
    for i, beat in enumerate(beats):
        ws = first_sheet if i == 0 else wb.create_sheet()
        ws.title = safe_sheet_name(beat, used)
        write_beat_sheet(ws, beat, [t for t in trees if t.beat_name == beat], per_beat[beat])

    if len(beats) > 1:
        ws = wb.create_sheet()
        ws.title = safe_sheet_name("Grand Abstract", used)
        write_grand_sheet(ws, grand, per_beat)

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    label = re.sub(r"[^\w\-]+", "_", only_beat) if only_beat else "all_beats"
    filename = f"tree_data_{label}_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Database repair (runs every time the app starts, safe to run again and again)
# - adds beat_name / km_burji_no if an older table does not have them
# - removes the OLD rule "tree no must be unique in the whole table"
# - adds the NEW rule from UNIQUE_COLUMNS
# ---------------------------------------------------------------------------
def upgrade_database():
    if db.engine.dialect.name != "mysql":
        return

    insp = inspect(db.engine)
    if not insp.has_table("shisham"):
        return

    existing_cols = {c["name"] for c in insp.get_columns("shisham")}
    for col, ddl in [("beat_name", "VARCHAR(100) NULL"), ("km_burji_no", "VARCHAR(50) NULL")]:
        if col not in existing_cols:
            with db.engine.begin() as conn:
                conn.execute(text(f"ALTER TABLE shisham ADD COLUMN {col} {ddl}"))
            print(f"[database] added column {col}")

    indexes = insp.get_indexes("shisham")

    # Drop every unique rule that is not the one we want
    # (this is the old rule that blocked the same tree no in another km).
    has_new_rule = False
    for idx in indexes:
        if not idx.get("unique"):
            continue
        cols = list(idx["column_names"])
        if idx["name"] == UNIQUE_NAME and cols == UNIQUE_COLUMNS:
            has_new_rule = True
        elif cols == ["tree_no"] or idx["name"] == UNIQUE_NAME:
            with db.engine.begin() as conn:
                conn.execute(text(f"ALTER TABLE shisham DROP INDEX `{idx['name']}`"))
            print(f"[database] removed old unique rule `{idx['name']}` on {cols}")

    if not has_new_rule:
        cols_sql = ", ".join(f"`{c}`" for c in UNIQUE_COLUMNS)
        try:
            with db.engine.begin() as conn:
                conn.execute(text(f"ALTER TABLE shisham ADD CONSTRAINT {UNIQUE_NAME} UNIQUE ({cols_sql})"))
            print(f"[database] added unique rule on {UNIQUE_COLUMNS}")
        except Exception as exc:  # most likely: duplicates already saved in the same km
            print(
                "[database] WARNING: could not add the unique rule, probably because the same "
                "tree no is already saved twice in one km. Fix those rows, then restart.\n"
                f"           {exc.__class__.__name__}: {str(exc).splitlines()[0][:200]}"
            )


# ---------------------------------------------------------------------------
# Friendly page when the database cannot be reached
# ---------------------------------------------------------------------------
DB_ERROR_PAGE = """<!doctype html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Database not reachable</title>
<body style="font:16px/1.5 system-ui,sans-serif;max-width:640px;margin:40px auto;padding:0 16px">
<h1>Cannot connect to the database</h1>
<p>The app is running, but it could not reach MySQL. Check these:</p>
<ol>
<li><b>Host</b>: copy it exactly from Aiven (it starts with <code>mysql-</code>).</li>
<li><b>Port, user, password, database name</b> in your <code>.env</code> file.</li>
<li><b>Aiven service</b> is running and your IP is in <i>Allowed IP addresses</i>.</li>
<li><b>ca.pem</b> is next to <code>app.py</code>.</li>
</ol>
<p>Open <a href="/health/db">/health/db</a> to see the exact error name.</p>
<pre style="background:#f3f3f3;padding:10px;overflow:auto">{{ detail }}</pre></body>"""


@app.errorhandler(OperationalError)
def database_unreachable(exc):
    db.session.rollback()
    app.logger.error("Database error: %s", exc)
    from flask import render_template_string
    detail = str(exc.orig) if getattr(exc, "orig", None) else str(exc)
    return render_template_string(DB_ERROR_PAGE, detail=detail[:300]), 503


@app.route("/health/db")
def database_health():
    try:
        with db.engine.connect() as connection:
            connection.execute(text("SELECT 1"))
        return {"database": "connected", "engine": db.engine.dialect.name}, 200
    except Exception as exc:
        app.logger.exception("Database connection failed")
        return {"database": "connection failed", "error": exc.__class__.__name__,
                "detail": str(getattr(exc, "orig", exc))[:300]}, 500


# ---------------------------------------------------------------------------
# Create tables / repair old tables on start (also works under gunicorn)
# ---------------------------------------------------------------------------
def init_database():
    with app.app_context():
        try:
            db.create_all()
            upgrade_database()
        except Exception as exc:
            print(f"[database] could not initialise: {exc.__class__.__name__}: "
                  f"{str(exc).splitlines()[0][:200]}")


init_database()


if __name__ == "__main__":
    app.run(debug=os.getenv("FLASK_DEBUG", "0") == "1", port=int(os.getenv("PORT", "5000")))
