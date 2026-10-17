"""Tree enumeration register (Flask + SQLAlchemy + MySQL on Aiven).

Run locally:   python app.py
Production:    gunicorn app:app
"""
import os
import re
import ssl
import tempfile
import uuid
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
from markupsafe import escape
from werkzeug.exceptions import HTTPException
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
app.config["SECRET_KEY"] = os.getenv("SECRET_KEY")
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
# (Every beat has its own table, so the rule is checked inside that table.)
# Set TREE_NO_UNIQUE_PER_KIND = True if the same tree no may also be used once
# for each kind (for example shisham 14 and kikar 14) in the same km / RD no.
# ---------------------------------------------------------------------------
TREE_NO_UNIQUE_PER_KIND = False
UNIQUE_COLUMNS = ["km_burji_no"] + (["kind_tree"] if TREE_NO_UNIQUE_PER_KIND else []) + ["tree_no"]


# ---------------------------------------------------------------------------
# One table per beat
#   "beats"                       : the list of beats (id, name, table_name)
#   "beat_<id>_<name>"            : one table per beat, holds only that beat's trees
# Table names are built from the beat's number and letters a-z / 0-9 only,
# never from raw text, so a beat name can not break the SQL.
# ---------------------------------------------------------------------------
class Beat(db.Model):
    __tablename__ = "beats"

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False, unique=True)
    table_name = db.Column(db.String(80), nullable=False, unique=True)


_tree_models = {}


def tree_model(table_name):
    """Return the model class for one beat's table (made once, then reused)."""
    model = _tree_models.get(table_name)
    if model is None:
        attrs = {
            "__tablename__": table_name,
            "__table_args__": (db.UniqueConstraint(*UNIQUE_COLUMNS, name=f"uq_{table_name}"),),
            "id": db.Column(db.Integer, primary_key=True),
            "tree_no": db.Column(db.String(50), nullable=False),
            "km_burji_no": db.Column(db.String(50), nullable=False),
            "kind_tree": db.Column(db.String(100), nullable=False),
            "girth": db.Column(db.Integer, nullable=False),
            "dia": db.Column(db.Integer),
            "cft": db.Column(db.Float),
            "shi_unit": db.Column(db.Float),
            "remarks": db.Column(db.String(255)),
            "to_dict": lambda self: {c.name: getattr(self, c.name) for c in self.__table__.columns},
        }
        model = type(f"Trees_{table_name}", (db.Model,), attrs)
        _tree_models[table_name] = model
    return model


def _slug(name):
    return re.sub(r"[^a-z0-9]+", "_", name.lower()).strip("_")[:40]


def get_beat(name):
    return Beat.query.filter_by(name=name).first()


def ensure_table(beat):
    try:
        tree_model(beat.table_name).__table__.create(db.engine, checkfirst=True)
    except Exception as exc:  # another worker may have created it a moment earlier
        if not inspect(db.engine).has_table(beat.table_name):
            raise exc


def get_or_create_beat(name):
    """Find the beat, or register it and create its own table."""
    beat = get_beat(name)
    if beat is None:
        try:
            beat = Beat(name=name, table_name="pending_" + uuid.uuid4().hex[:12])
            db.session.add(beat)
            db.session.flush()
            slug = _slug(name)
            beat.table_name = f"beat_{beat.id}" + (f"_{slug}" if slug else "")
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            beat = get_beat(name)
    ensure_table(beat)
    return beat


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
    """Beats from the BEATS list plus every beat already registered."""
    names = set(BEATS) | {name for (name,) in db.session.query(Beat.name).all()}
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
    """Returns ({beat name: abstract}, grand_abstract). Beats with no trees are left out."""
    beats = Beat.query.filter_by(name=only_beat).all() if only_beat else Beat.query.all()
    per_beat, grand = {}, _new_abstract()
    for beat in beats:
        Tree = tree_model(beat.table_name)
        rows = db.session.query(
            Tree.kind_tree,
            func.count(Tree.id),
            func.coalesce(func.sum(Tree.cft), 0),
            func.coalesce(func.sum(Tree.shi_unit), 0),
        ).group_by(Tree.kind_tree).all()
        if not rows:
            continue
        a = _new_abstract()
        for kind, trees, cft, shi_unit in rows:
            _add_to_abstract(a, kind, trees, float(cft), float(shi_unit))
            _add_to_abstract(grand, kind, trees, float(cft), float(shi_unit))
        per_beat[beat.name] = _finish_abstract(a)
    return per_beat, _finish_abstract(grand)


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

    beat is the Beat whose table the tree goes into (or None if no beat is chosen).
    exclude_id is the tree being edited, so it is not counted as its own duplicate.
    """
    tree_no = " ".join(form.get("tree_no", "").split())
    km_burji_no = " ".join(form.get("km_burji_no", "").split())
    kind_tree = form.get("kind_tree", "").strip().lower()
    girth_raw = form.get("girth", "").strip()
    remarks = form.get("remarks", "").strip()

    if beat is None:
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

    Tree = tree_model(beat.table_name)
    key = {"km_burji_no": km_burji_no, "kind_tree": kind_tree, "tree_no": tree_no}
    duplicate = Tree.query.filter_by(**{col: key[col] for col in UNIQUE_COLUMNS})
    if exclude_id is not None:
        duplicate = duplicate.filter(Tree.id != exclude_id)
    if duplicate.first():
        return duplicate_message(tree_no, kind_tree, km_burji_no, beat.name), None

    return None, {
        "tree_no": tree_no,
        "km_burji_no": km_burji_no,
        "kind_tree": kind_tree,
        "girth": girth,
        "dia": row["dia"],
        "cft": row["cft"],
        "shi_unit": row["shi_unit"],
        "remarks": remarks or None,
    }


def render_index(error=None, form=None, edit_tree=None, edit_beat=None):
    """Build the page: one section per beat (current beat first, then A-Z)."""
    current_beat = session.get("beat_name")

    abstracts, grand = get_abstracts()
    sections = []
    for beat in Beat.query.all():
        if beat.name not in abstracts:
            continue
        Tree = tree_model(beat.table_name)
        sections.append({
            "id": beat.id,
            "beat": beat.name,
            "trees": Tree.query.order_by(Tree.id.asc()).all(),
            "abstract": abstracts[beat.name],
        })
    sections.sort(key=lambda s: (s["beat"] != current_beat, s["beat"].lower()))

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
        edit_beat=edit_beat,
    )


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@app.route("/set-beat", methods=["POST"])
def set_beat():
    """Choose the beat once. Its own table is created the first time it is used."""
    name = normalize_beat(request.form.get("beat_name"))

    if not name:
        flash("Please pick or type a beat name.", "error")
        return redirect(url_for("index"))

    get_or_create_beat(name)
    session["beat_name"] = name
    return redirect(url_for("index"))


@app.route("/", methods=["GET", "POST"])
def index():
    current_beat = session.get("beat_name")

    # ----- Add a new tree -----
    if request.method == "POST":
        beat = get_or_create_beat(current_beat) if current_beat else None
        error, data = read_tree_form(request.form, beat)

        if error is None:
            Tree = tree_model(beat.table_name)
            tree = Tree(**data)
            db.session.add(tree)
            try:
                db.session.commit()
            except IntegrityError:
                db.session.rollback()
                error = duplicate_message(
                    data["tree_no"], data["kind_tree"], data["km_burji_no"], beat.name
                )
            else:
                flash(
                    f"Saved tree no {tree.tree_no} ({tree.km_burji_no}) in {beat.name}: "
                    f"{tree.kind_tree.capitalize()}, girth {tree.girth}, dia {tree.dia}, "
                    f"cft {tree.cft}, shi unit {tree.shi_unit}",
                    "success",
                )
                # Redirect so refreshing the page does not save a duplicate
                return redirect(url_for("index"))

        return render_index(error=error, form=request.form)

    # ----- Show the page, optionally with one tree loaded for editing (?edit=ID&beat=BEAT_ID) -----
    edit_tree = edit_beat = None
    edit_id = request.args.get("edit", type=int)
    beat_id = request.args.get("beat", type=int)
    if edit_id and beat_id:
        edit_beat = db.session.get(Beat, beat_id)
        if edit_beat:
            edit_tree = db.session.get(tree_model(edit_beat.table_name), edit_id)
        if edit_tree is None:
            flash("That tree no longer exists.", "error")
            return redirect(url_for("index"))

    return render_index(
        form=edit_tree.to_dict() if edit_tree else {}, edit_tree=edit_tree, edit_beat=edit_beat
    )


@app.route("/update/<int:beat_id>/<int:tree_id>", methods=["POST"])
def update_tree(beat_id, tree_id):
    """Save changes to a tree in its beat's table. dia, cft and shi unit are worked out again."""
    beat = db.get_or_404(Beat, beat_id)
    Tree = tree_model(beat.table_name)
    tree = db.get_or_404(Tree, tree_id)

    error, data = read_tree_form(request.form, beat, exclude_id=tree.id)

    if error is None:
        for field, value in data.items():
            setattr(tree, field, value)
        try:
            db.session.commit()
        except IntegrityError:
            db.session.rollback()
            error = duplicate_message(
                data["tree_no"], data["kind_tree"], data["km_burji_no"], beat.name
            )
        else:
            flash(
                f"Updated tree no {tree.tree_no} ({tree.km_burji_no}) in {beat.name}: "
                f"{tree.kind_tree.capitalize()}, girth {tree.girth}, dia {tree.dia}, "
                f"cft {tree.cft}, shi unit {tree.shi_unit}",
                "success",
            )
            return redirect(url_for("index"))

    db.session.rollback()
    return render_index(error=error, form=request.form, edit_tree=tree, edit_beat=beat)


@app.route("/delete/<int:beat_id>/<int:tree_id>", methods=["POST"])
def delete_tree(beat_id, tree_id):
    beat = db.get_or_404(Beat, beat_id)
    tree = db.get_or_404(tree_model(beat.table_name), tree_id)
    label = (
        f"tree no {tree.tree_no} ({tree.kind_tree.capitalize()}, km / burji no "
        f"{tree.km_burji_no}) from {beat.name}"
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
    per_beat, grand = get_abstracts(only_beat)

    wb = Workbook()
    first_sheet = wb.active

    beats = sorted(per_beat, key=str.lower)
    if not beats:
        first_sheet.title = "No Data"
    used = set()
    for i, name in enumerate(beats):
        Tree = tree_model(get_beat(name).table_name)
        trees = Tree.query.order_by(Tree.id).all()
        ws = first_sheet if i == 0 else wb.create_sheet()
        ws.title = safe_sheet_name(name, used)
        write_beat_sheet(ws, name, trees, per_beat[name])

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
# Move trees from the OLD single table ("shisham" or "trees") into the per-beat
# tables. Runs at start. The old table is renamed to <name>_old_backup, not deleted.
# ---------------------------------------------------------------------------
def migrate_old_tables():
    insp = inspect(db.engine)
    for old in ("shisham", "trees"):
        if not insp.has_table(old):
            continue
        cols = {c["name"] for c in insp.get_columns(old)}
        if not {"tree_no", "kind_tree", "girth"} <= cols:
            continue
        rows = db.session.execute(text(f"SELECT * FROM `{old}`")).mappings().all()
        moved = skipped = 0
        for r in rows:
            name = " ".join(((r.get("beat_name") if "beat_name" in cols else None) or "Unassigned").split())
            beat = get_or_create_beat(normalize_beat(name))
            Tree = tree_model(beat.table_name)
            db.session.add(Tree(
                tree_no=r["tree_no"], km_burji_no=r.get("km_burji_no") or "-",
                kind_tree=r["kind_tree"], girth=r["girth"], dia=r.get("dia"),
                cft=r.get("cft"), shi_unit=r.get("shi_unit"), remarks=r.get("remarks"),
            ))
            try:
                db.session.commit()
                moved += 1
            except IntegrityError:
                db.session.rollback()
                skipped += 1
        backup = f"{old}_old_backup"
        if not insp.has_table(backup):
            with db.engine.begin() as conn:
                conn.execute(text(f"ALTER TABLE `{old}` RENAME TO `{backup}`"))
        print(f"[database] moved {moved} trees from `{old}` into per-beat tables "
              f"({skipped} skipped as duplicates). Old table kept as `{backup}`.")


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
@app.errorhandler(Exception)
def show_error(exc):
    if isinstance(exc, HTTPException):
        return exc
    db.session.rollback()
    app.logger.exception("Unhandled error")
    if os.getenv("SHOW_ERRORS") == "1":
        return f"<h1>Error</h1><pre>{escape(exc.__class__.__name__)}: {escape(str(exc)[:500])}</pre>", 500
    return "Internal Server Error", 500


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
# Create tables on start (also works under gunicorn)
# ---------------------------------------------------------------------------
def init_database():
    with app.app_context():
        try:
            db.create_all()                      # the "beats" list table
            for beat in Beat.query.all():        # beats saved earlier
                ensure_table(beat)
            for name in BEATS:                   # the beats written in BEATS above
                get_or_create_beat(name)
            migrate_old_tables()
        except Exception as exc:
            db.session.rollback()
            print(f"[database] could not initialise: {exc.__class__.__name__}: "
                  f"{str(exc).splitlines()[0][:200]}")


init_database()


if __name__ == "__main__":
    app.run(debug=os.getenv("FLASK_DEBUG", "0") == "1", port=int(os.getenv("PORT", "5000")))