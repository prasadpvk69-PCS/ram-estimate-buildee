import io
import math
import re
import json
import os
from datetime import datetime
from functools import wraps
from flask import Flask, jsonify, request, send_file, render_template, session, redirect, url_for
from werkzeug.security import generate_password_hash, check_password_hash
import psycopg2
import psycopg2.extras
import openpyxl
from openpyxl.styles import (Font, Alignment, Border, Side, PatternFill)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.page import PageMargins

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "RAM_Estimate_Builder_S3cr3t_2024")

# DB: use DATABASE_URL env var (Render) or fall back to local config
_DATABASE_URL = os.environ.get("DATABASE_URL", "")
if _DATABASE_URL.startswith("postgres://"):          # Render uses postgres://, psycopg2 needs postgresql://
    _DATABASE_URL = _DATABASE_URL.replace("postgres://", "postgresql://", 1)

_LOCAL_DB = dict(host="localhost", user="postgres", password="natgrid@2024", dbname="postgres")

def get_conn():
    if _DATABASE_URL:
        return psycopg2.connect(_DATABASE_URL)
    return psycopg2.connect(**_LOCAL_DB)

# ── Auth helpers ─────────────────────────────────────────────────────────────
def require_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return jsonify({"error": "Unauthorized"}), 401
        return f(*args, **kwargs)
    return decorated

def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return jsonify({"error": "Unauthorized"}), 401
        if session.get("role") != "admin":
            return jsonify({"error": "Forbidden"}), 403
        return f(*args, **kwargs)
    return decorated

# ── DB init ───────────────────────────────────────────────────────────────────
def init_descriptions_table(conn):
    """Create work_descriptions table and load from CSV if empty."""
    cur = conn.cursor()
    cur.execute("""
        CREATE EXTENSION IF NOT EXISTS pg_trgm;
        CREATE TABLE IF NOT EXISTS work_descriptions (
            id          SERIAL PRIMARY KEY,
            description TEXT    NOT NULL,
            unit        VARCHAR(50),
            source_file VARCHAR(300),
            search_text TSVECTOR GENERATED ALWAYS AS (to_tsvector('english', description)) STORED
        );
        CREATE INDEX IF NOT EXISTS idx_work_desc_ilike ON work_descriptions
            USING GIN(description gin_trgm_ops);
    """)
    conn.commit()
    # Load from CSV if table is empty
    cur.execute("SELECT COUNT(*) FROM work_descriptions")
    if cur.fetchone()[0] == 0:
        csv_path = os.path.join(os.path.dirname(__file__), "data", "descriptions.csv")
        if os.path.exists(csv_path):
            import csv as _csv
            with open(csv_path, encoding="utf-8") as f:
                reader = _csv.DictReader(f)
                rows = [(r["description"], r["unit"] or None) for r in reader
                        if r["description"].strip()]
            cur.executemany(
                "INSERT INTO work_descriptions (description, unit, source_file) VALUES (%s,%s,'csv')",
                rows
            )
            conn.commit()
            print(f"[DB] Loaded {len(rows)} descriptions from CSV")

def init_users_table():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS users (
            id            SERIAL PRIMARY KEY,
            username      VARCHAR(50) UNIQUE NOT NULL,
            password_hash VARCHAR(200) NOT NULL,
            role          VARCHAR(10) DEFAULT 'user',
            created_at    TIMESTAMP DEFAULT NOW()
        )
    """)
    conn.commit()
    # Create default admin if not exists
    cur.execute("SELECT 1 FROM users WHERE username = 'admin'")
    if not cur.fetchone():
        cur.execute(
            "INSERT INTO users (username, password_hash, role) VALUES (%s, %s, 'admin')",
            ("admin", generate_password_hash("Admin@RAM2024"))
        )
        conn.commit()
    conn.close()

def init_log_table():
    conn = get_conn()
    cur  = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS estimate_logs (
            id          SERIAL PRIMARY KEY,
            name_of_work TEXT,
            item_count  INT,
            total_amount NUMERIC(14,2),
            filename    VARCHAR(300),
            items_json  TEXT,
            created_at  TIMESTAMP DEFAULT NOW()
        )
    """)
    conn.commit()
    conn.close()

_db_ready = False

def _init_db():
    global _db_ready
    if _db_ready:
        return
    try:
        conn = get_conn()
        init_descriptions_table(conn)
        conn.close()
        init_users_table()
        init_log_table()
        _db_ready = True
        print("[DB] All tables ready.")
    except Exception as e:
        print(f"[DB INIT] Will retry on first request: {e}")

# Try at startup — if DB not ready yet (Render cold start), retry on first request
_init_db()

@app.before_request
def ensure_db():
    _init_db()

# ── Login / Logout ────────────────────────────────────────────────────────────
@app.route("/login", methods=["GET"])
def login_page():
    if "user_id" in session:
        return redirect(url_for("index"))
    return render_template("login.html")

@app.route("/login", methods=["POST"])
def login_post():
    data     = request.get_json()
    username = (data.get("username") or "").strip()
    password = data.get("password") or ""
    conn = get_conn()
    cur  = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM users WHERE username = %s", (username,))
    user = cur.fetchone()
    conn.close()
    if not user or not check_password_hash(user["password_hash"], password):
        return jsonify({"error": "Invalid username or password"}), 401
    session["user_id"]  = user["id"]
    session["username"] = user["username"]
    session["role"]     = user["role"]
    return jsonify({"ok": True, "role": user["role"], "username": user["username"]})

@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"ok": True})

# ── User management (admin only) ──────────────────────────────────────────────
@app.route("/api/users", methods=["GET"])
@require_admin
def list_users():
    conn = get_conn()
    cur  = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT id, username, role, created_at FROM users ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    result = []
    for r in rows:
        d = dict(r)
        d["created_at"] = d["created_at"].strftime("%d-%m-%Y %H:%M") if d["created_at"] else ""
        result.append(d)
    return jsonify(result)

@app.route("/api/users", methods=["POST"])
@require_admin
def create_user():
    data     = request.get_json()
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    role     = data.get("role", "user")
    if len(username) < 3:
        return jsonify({"error": "Username must be at least 3 characters"}), 400
    if len(password) < 4:
        return jsonify({"error": "Password must be at least 4 characters"}), 400
    if role not in ("admin", "user"):
        role = "user"
    conn = get_conn()
    cur  = conn.cursor()
    cur.execute("SELECT 1 FROM users WHERE username = %s", (username,))
    if cur.fetchone():
        conn.close()
        return jsonify({"error": "Username already exists"}), 409
    cur.execute(
        "INSERT INTO users (username, password_hash, role) VALUES (%s, %s, %s) RETURNING id",
        (username, generate_password_hash(password), role)
    )
    new_id = cur.fetchone()[0]
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "id": new_id})

@app.route("/api/users/<int:uid>", methods=["DELETE"])
@require_admin
def delete_user(uid):
    if uid == session.get("user_id"):
        return jsonify({"error": "Cannot delete yourself"}), 400
    conn = get_conn()
    cur  = conn.cursor()
    cur.execute("DELETE FROM users WHERE id = %s AND username != 'admin'", (uid,))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

# ── Public descriptions export (for Google Sheets import) ─────────────────────
@app.route("/api/public/descriptions")
def public_descriptions():
    """No auth — used by Google Sheets Apps Script to populate _DB sheet."""
    conn = get_conn()
    cur  = conn.cursor()
    cur.execute("SELECT description, unit FROM work_descriptions ORDER BY id")
    rows = cur.fetchall()
    conn.close()
    return jsonify([{"d": r[0], "u": r[1] or ""} for r in rows])

# ── Search API ────────────────────────────────────────────────────────────────
@app.route("/api/search")
@require_auth
def search():
    q = request.args.get("q", "").strip()
    if len(q) < 2:
        return jsonify([])
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    # Split keywords and match all of them (case-insensitive)
    words = q.split()
    conditions = " AND ".join(["description ILIKE %s"] * len(words))
    params = [f"%{w}%" for w in words]
    cur.execute(
        f"SELECT id, description, unit FROM work_descriptions WHERE {conditions} ORDER BY length(description) LIMIT 20",
        params
    )
    rows = cur.fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

# ── Generate Excel API ────────────────────────────────────────────────────────
def _build_excel(data):
    """Build the Excel workbook from a data dict. Returns (BytesIO, filename)."""
    name_of_work = data.get("name_of_work", "")
    items = data.get("items", [])
    seigniorage = float(data.get("seigniorage", 0) or 0)
    nac_pct = float(data.get("nac_pct", 0.1) or 0.1)
    gst_pct = float(data.get("gst_pct", 18) or 18)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Estimate"

    # ── Column widths — match original file layout ────────────────────────────
    col_widths = [6, 75, 8, 14, 10, 10, 12, 12, 14]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── A4 print setup ───────────────────────────────────────────────────────
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.paperSize  = 9            # 9 = A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.75, bottom=0.75, header=0.3, footer=0.3)
    ws.print_title_rows = "1:6"             # repeat column headers on every printed page

    # ── Styles ────────────────────────────────────────────────────────────────
    thin = Side(style="thin")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_top_bot = Border(top=thin, bottom=thin)

    def hdr_font():  return Font(name="Arial", bold=True, size=10)
    def body_font(): return Font(name="Arial", size=9)
    def center():    return Alignment(horizontal="center", vertical="center", wrap_text=True)
    def left_wrap(): return Alignment(horizontal="left",   vertical="top",    wrap_text=True)
    def right():     return Alignment(horizontal="right",  vertical="center")

    blue_fill  = PatternFill("solid", fgColor="BDD7EE")
    grey_fill  = PatternFill("solid", fgColor="D9D9D9")
    total_fill = PatternFill("solid", fgColor="FFF2CC")

    def write_cell(row, col, value, font=None, align=None, fill=None, border=None, number_format=None):
        c = ws.cell(row=row, column=col, value=value)
        if font:          c.font          = font
        if align:         c.alignment     = align
        if fill:          c.fill          = fill
        if border:        c.border        = border
        if number_format: c.number_format = number_format
        return c

    # ── Header rows — matching original file exactly ──────────────────────────
    # Row 1: APSRTC (thin row, no fill, just centered text like original)
    r = 1
    ws.merge_cells(f"A{r}:I{r}")
    write_cell(r, 1, "RAM", Font(name="Arial", bold=True, size=11), center())
    ws.row_dimensions[r].height = 14

    # Row 2: Detailed Estimate — merged, centered, bold
    r = 2
    ws.merge_cells(f"A{r}:I{r}")
    write_cell(r, 1, "Detailed Estimate", Font(name="Arial", bold=True, size=12), center())
    ws.row_dimensions[r].height = 18

    # Row 3: NAME OF WORK — merged, bold, left-aligned, wraps
    r = 3
    ws.merge_cells(f"A{r}:I{r}")
    write_cell(r, 1, f"NAME OF WORK :- {name_of_work}",
               Font(name="Arial", bold=True, size=10), left_wrap())
    ws.row_dimensions[r].height = 30

    # ── Column headers ────────────────────────────────────────────────────────
    # Merge spans rows 4-5 for SL, Particulars, No, Qty, Rate, Amount
    # D4:F4 spans the "Measurement" label; D5/E5/F5 get L/B/H
    r = 4
    ws.merge_cells(f"A4:A5")
    ws.merge_cells(f"B4:B5")
    ws.merge_cells(f"C4:C5")
    ws.merge_cells(f"D4:F4")
    ws.merge_cells(f"G4:G5")
    ws.merge_cells(f"H4:H5")
    ws.merge_cells(f"I4:I5")
    # Write only to top-left of each merged region
    for col, h in [(1,"SL. NO."),(2,"Particulars"),(3,"No"),(4,"Measurement"),(7,"Quantity"),(8,"Rate"),(9,"Amount")]:
        write_cell(4, col, h, hdr_font(), center(), grey_fill, border_all)
    ws.row_dimensions[4].height = 18

    # Row 5: only D5, E5, F5 are unmerged
    for col, h in [(4,"L"),(5,"B"),(6,"H")]:
        write_cell(5, col, h, hdr_font(), center(), grey_fill, border_all)
    ws.row_dimensions[5].height = 16

    r = 6
    for col in range(1, 10):
        write_cell(r, col, col, hdr_font(), center(), grey_fill, border_all)
    ws.row_dimensions[r].height = 14

    # ── Item rows ─────────────────────────────────────────────────────────────
    current_row = 7
    amount_cells = []   # collect cell refs for grand total formula

    for sl, item in enumerate(items, 1):
        description  = item.get("description", "")
        measurements = item.get("measurements", [])
        rate         = float(item.get("rate") or 0)
        unit         = item.get("unit", "")
        unit_divisor = float(item.get("unit_divisor") or 1)
        item_amount  = _try_float(item.get("amount"))
        item_total_qty = _try_float(item.get("total_qty"))

        # Description row — A=sl.no, B:E merged for description
        dr = current_row
        write_cell(dr, 1, sl, hdr_font(), center(), None, border_all)
        ws.merge_cells(f"B{dr}:E{dr}")
        write_cell(dr, 2, description, body_font(), left_wrap(), None, border_all)
        for col in range(6, 10):   # F(H-dim), G(Qty), H(Rate), I(Amount) empty
            write_cell(dr, col, None, body_font(), center(), None, border_all)
        lines = math.ceil(len(description) / 130)
        ws.row_dimensions[dr].height = max(45, min(14 * max(lines, 3), 200))
        current_row += 1

        # Measurement rows
        qty_cells = []
        for mrow in measurements:
            mr = current_row
            label        = mrow.get("label", "")
            no           = mrow.get("no", "")
            L            = mrow.get("L", "")
            B            = mrow.get("B", "")
            H            = mrow.get("H", "")
            qty_override = mrow.get("qty_override", "")

            write_cell(mr, 1, None,  body_font(), center(), None, border_all)
            write_cell(mr, 2, label, body_font(), left_wrap(), None, border_all)
            write_cell(mr, 3, no,    body_font(), center(), None, border_all)

            # L, B, H — write as-is (text or number)
            for cidx, val in [(4, L), (5, B), (6, H)]:
                num = _try_float(val)
                write_cell(mr, cidx, num if num is not None else (val or None),
                           body_font(), center(), None, border_all)

            # Qty — use manual value entered by user
            qty_val = _try_float(qty_override)
            write_cell(mr, 7, qty_val, body_font(), center(), None, border_all, "#,##0.00")
            qty_cells.append(get_column_letter(7) + str(mr))

            write_cell(mr, 8, None, body_font(), center(), None, border_all)
            write_cell(mr, 9, None, body_font(), center(), None, border_all)
            ws.row_dimensions[mr].height = 16
            current_row += 1

        # Total quantity + rate + amount row
        tr = current_row
        if item_total_qty is not None:
            total_qty_val = item_total_qty
        else:
            total_qty_val = f"=SUM({','.join(qty_cells)})" if qty_cells else 0
        write_cell(tr, 1, None, body_font(), center(), total_fill, border_all)
        write_cell(tr, 2, None, body_font(), center(), total_fill, border_all)
        write_cell(tr, 3, None, body_font(), center(), total_fill, border_all)
        write_cell(tr, 4, None, body_font(), center(), total_fill, border_all)
        write_cell(tr, 5, None, body_font(), center(), total_fill, border_all)
        write_cell(tr, 6, None, body_font(), center(), total_fill, border_all)
        write_cell(tr, 7, total_qty_val, Font(name="Arial", bold=True, size=9), center(), total_fill, border_all, "#,##0.00")
        write_cell(tr, 8, rate if rate else None, Font(name="Arial", bold=True, size=9), right(), total_fill, border_all, "#,##0.00")
        if item_amount is not None:
            amt_val = item_amount
        elif unit_divisor != 1:
            amt_val = f"=G{tr}*H{tr}/{unit_divisor}"
        else:
            amt_val = f"=G{tr}*H{tr}"
        write_cell(tr, 9, amt_val, Font(name="Arial", bold=True, size=9), right(), total_fill, border_all, "#,##0.00")
        amount_cells.append(f"I{tr}")
        ws.row_dimensions[tr].height = 16
        current_row += 1

        # Unit row — unit text placed only under Rate column (H), not full-width merge
        ur = current_row
        for col in range(1, 10):
            write_cell(ur, col, None, body_font(), center(), None, border_all)
        if unit:
            write_cell(ur, 8, unit,   # col H = Rate column
                       Font(name="Arial", size=9, italic=True),
                       center(), PatternFill("solid", fgColor="FFF2CC"), border_all)
        ws.row_dimensions[ur].height = 13
        current_row += 1

    # ── Footer ────────────────────────────────────────────────────────────────
    def footer_row(label, value_or_formula, is_total=False):
        nonlocal current_row
        r = current_row
        ws.merge_cells(f"A{r}:H{r}")
        ff = Font(name="Arial", bold=True, size=10)
        tf = PatternFill("solid", fgColor="FCE4D6") if is_total else None
        write_cell(r, 1, label, ff, right(), tf, border_all)
        write_cell(r, 9, value_or_formula, ff, right(), tf, border_all, "#,##0.00")
        ws.row_dimensions[r].height = 18
        current_row += 1
        return f"I{r}"

    subtotal_formula = f"=SUM({','.join(amount_cells)})" if amount_cells else 0
    sub_ref  = footer_row("Sub Total", subtotal_formula)
    seig_ref = footer_row("Add: Seigniorage Charges", seigniorage)
    tot1_ref = footer_row(f"Total", f"={sub_ref}+I{int(seig_ref[1:])}", is_total=True)

    nac_val  = nac_pct / 100
    nac_ref  = footer_row(f"Add NAC {nac_pct}%", f"=ROUND({tot1_ref}*{nac_val},0)")
    tot2_ref = footer_row("Total", f"={tot1_ref}+I{int(nac_ref[1:])}", is_total=True)

    gst_val  = gst_pct / 100
    gst_ref  = footer_row(f"Add GST {gst_pct}%", f"=ROUND({tot2_ref}*{gst_val},0)")
    tot3_ref = footer_row("Grand Total", f"={tot2_ref}+I{int(gst_ref[1:])}", is_total=True)

    # Say (rounded to nearest 10)
    r = current_row
    ws.merge_cells(f"A{r}:H{r}")
    write_cell(r, 1, "Say", Font(name="Arial", bold=True, size=11), right(),
               PatternFill("solid", fgColor="E2EFDA"), border_all)
    write_cell(r, 9, f"=ROUND({tot3_ref},-2)", Font(name="Arial", bold=True, size=11),
               right(), PatternFill("solid", fgColor="E2EFDA"), border_all, "#,##0.00")
    ws.row_dimensions[r].height = 20

    # ── Save to buffer ────────────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', name_of_work)[:80] or "estimate"
    filename = f"{safe_name}.xlsx"
    return buf, filename, name_of_work, items, seigniorage, nac_pct, gst_pct


@app.route("/api/generate", methods=["POST"])
@require_auth
def generate():
    data = request.get_json()
    buf, filename, name_of_work, items, seigniorage, nac_pct, gst_pct = _build_excel(data)

    # ── Log this estimate ─────────────────────────────────────────────────────
    try:
        grand_approx = 0
        for item in items:
            stored_amt = _try_float(item.get("amount"))
            if stored_amt is not None:
                grand_approx += stored_amt
            else:
                rate = float(item.get("rate") or 0)
                divisor = float(item.get("unit_divisor") or 1)
                tq = _try_float(item.get("total_qty"))
                if tq is not None:
                    grand_approx += tq * rate / divisor
                else:
                    for m in item.get("measurements", []):
                        no_v = _parse_no(m.get("no","")) or 1
                        lv, _ = _eval_dim(m.get("L",""))
                        bv, _ = _eval_dim(m.get("B",""))
                        hv, _ = _eval_dim(m.get("H",""))
                        qty = no_v * (lv or 1) * (bv or 1) * (hv or 1)
                        grand_approx += qty * rate / divisor
        grand_approx += seigniorage
        grand_approx += grand_approx * nac_pct / 100
        grand_approx += grand_approx * gst_pct / 100

        conn = get_conn()
        cur = conn.cursor()
        payload_json = json.dumps({
            "name_of_work": name_of_work,
            "items": items,
            "seigniorage": seigniorage,
            "nac_pct": nac_pct,
            "gst_pct": gst_pct
        })
        cur.execute("""INSERT INTO estimate_logs (name_of_work, item_count, total_amount, filename, items_json)
                       VALUES (%s, %s, %s, %s, %s)""",
                    (name_of_work, len(items), round(grand_approx, 2), filename, payload_json))
        conn.commit()
        conn.close()
    except Exception:
        pass

    return send_file(buf, as_attachment=True,
                     download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── Helpers ───────────────────────────────────────────────────────────────────
def _try_float(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def _eval_dim(s):
    """
    Parse civil estimate dimension expressions into a float.
    Handles:
      - Plain numbers: "3.07" → 3.07
      - Sum/avg: "7.15 + 6.15   2" → (7.15+6.15)/2 = 6.65
      - Simple expressions: "7.15 + 6.15 + 3.0   3" → (7.15+6.15+3.0)/3
    Returns (numeric_value, is_text_expr) where is_text_expr=True means the
    original was a text note so cell should store the text but formula uses numeric.
    """
    if not s and s != 0:
        return None, False
    s = str(s).strip()
    if not s:
        return None, False
    # Plain number
    try:
        return float(s), False
    except ValueError:
        pass
    # Pattern: "a + b [+ c ...] <space> n"  → sum(...)/n
    # The last whitespace-separated token that's a plain number is the divisor
    tokens = s.split()
    if len(tokens) >= 2:
        try:
            divisor = float(tokens[-1])
            expr = ' '.join(tokens[:-1])
            # safe eval: only allow +, -, *, /, digits, dots, spaces, parentheses
            if re.match(r'^[\d\s\.\+\-\*\/\(\)]+$', expr):
                val = eval(expr, {"__builtins__": {}})
                return float(val) / divisor, True
        except Exception:
            pass
    # Try plain eval of the whole string
    try:
        if re.match(r'^[\d\s\.\+\-\*\/\(\)]+$', s):
            return float(eval(s, {"__builtins__": {}})), True
    except Exception:
        pass
    return None, True  # text that can't be parsed

def _parse_no(no):
    """Return numeric multiplier from 'no' field (e.g. '1x1'->1, '2x1'->2, '3'->3)."""
    if not no:
        return None
    s = str(no).strip()
    m = re.match(r'^(\d+)\s*[xX×]\s*(\d+)$', s)
    if m:
        return int(m.group(1)) * int(m.group(2))
    try:
        return float(s)
    except ValueError:
        return None

def _qty_formula(row, no_val, L, B, H, L_num, B_num, H_num):
    """Build Excel formula for quantity = no * L * B * H.
    Uses cell refs for numeric dims, inline values for text-expression dims."""
    parts = []
    if no_val is not None:
        parts.append(str(no_val))
    # L
    L_f = _try_float(L)
    if L_f is not None:
        parts.append(f"D{row}")      # L is plain number → use cell ref
    elif L_num is not None:
        parts.append(str(round(L_num, 6)))  # L was text expr → inline value
    # B
    B_f = _try_float(B)
    if B_f is not None:
        parts.append(f"E{row}")
    elif B_num is not None:
        parts.append(str(round(B_num, 6)))
    # H
    H_f = _try_float(H)
    if H_f is not None:
        parts.append(f"F{row}")
    elif H_num is not None:
        parts.append(str(round(H_num, 6)))
    if not parts:
        return None
    return "=" + "*".join(parts)

# ── Log: list previous estimates ──────────────────────────────────────────────
@app.route("/api/logs")
@require_auth
def get_logs():
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("""
        SELECT id, name_of_work, item_count, total_amount, filename, created_at
        FROM estimate_logs ORDER BY created_at DESC LIMIT 100
    """)
    rows = cur.fetchall()
    conn.close()
    result = []
    for r in rows:
        d = dict(r)
        d["created_at"] = d["created_at"].strftime("%d-%m-%Y %H:%M") if d["created_at"] else ""
        d["total_amount"] = float(d["total_amount"]) if d["total_amount"] else 0
        result.append(d)
    return jsonify(result)

@app.route("/api/logs/<int:log_id>", methods=["GET"])
@require_auth
def get_log(log_id):
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM estimate_logs WHERE id = %s", (log_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return jsonify({"error": "Not found"}), 404
    d = dict(row)
    d["created_at"] = d["created_at"].strftime("%d-%m-%Y %H:%M") if d["created_at"] else ""
    d["total_amount"] = float(d["total_amount"]) if d["total_amount"] else 0
    return jsonify(d)

@app.route("/api/logs/<int:log_id>/excel", methods=["GET"])
@require_auth
def download_log_excel(log_id):
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    cur.execute("SELECT * FROM estimate_logs WHERE id = %s", (log_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return jsonify({"error": "Not found"}), 404
    try:
        payload = json.loads(row["items_json"])
        if isinstance(payload, list):
            payload = {"name_of_work": row["name_of_work"], "items": payload,
                       "seigniorage": 0, "nac_pct": 0.1, "gst_pct": 18}
    except Exception:
        return jsonify({"error": "Corrupt log data"}), 500
    buf, filename, *_ = _build_excel(payload)
    return send_file(buf, as_attachment=True,
                     download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/api/logs/<int:log_id>", methods=["DELETE"])
@require_auth
def delete_log(log_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM estimate_logs WHERE id = %s", (log_id,))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

# ── Descriptions: list / add / delete ─────────────────────────────────────────
@app.route("/api/descriptions")
@require_auth
def list_descriptions():
    q = request.args.get("q", "").strip()
    page = int(request.args.get("page", 1))
    per_page = 30
    offset = (page - 1) * per_page
    conn = get_conn()
    cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
    if q:
        words = q.split()
        cond = " AND ".join(["description ILIKE %s"] * len(words))
        params = [f"%{w}%" for w in words] + [per_page, offset]
        cur.execute(f"SELECT id, description, unit, source_file FROM work_descriptions WHERE {cond} ORDER BY id LIMIT %s OFFSET %s", params)
        cur2 = conn.cursor()
        cur2.execute(f"SELECT COUNT(*) FROM work_descriptions WHERE {cond}", [f"%{w}%" for w in words])
    else:
        cur.execute("SELECT id, description, unit, source_file FROM work_descriptions ORDER BY id LIMIT %s OFFSET %s", (per_page, offset))
        cur2 = conn.cursor()
        cur2.execute("SELECT COUNT(*) FROM work_descriptions")
    total = cur2.fetchone()[0]
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return jsonify({"rows": rows, "total": total, "page": page, "per_page": per_page})

@app.route("/api/descriptions", methods=["POST"])
@require_auth
def add_description():
    data = request.get_json()
    desc = (data.get("description") or "").strip()
    unit = (data.get("unit") or "").strip() or None
    if len(desc) < 10:
        return jsonify({"error": "Description too short"}), 400
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT 1 FROM work_descriptions WHERE description = %s", (desc,))
    if cur.fetchone():
        conn.close()
        return jsonify({"error": "Description already exists"}), 409
    cur.execute("INSERT INTO work_descriptions (description, unit, source_file) VALUES (%s,%s,%s) RETURNING id",
                (desc, unit, "Manual Entry"))
    new_id = cur.fetchone()[0]
    conn.commit()
    conn.close()
    return jsonify({"ok": True, "id": new_id})

@app.route("/api/descriptions/<int:desc_id>", methods=["DELETE"])
@require_auth
def delete_description(desc_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM work_descriptions WHERE id = %s", (desc_id,))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

@app.route("/api/descriptions/<int:desc_id>", methods=["PUT"])
@require_auth
def update_description(desc_id):
    data = request.get_json()
    desc = (data.get("description") or "").strip()
    unit = (data.get("unit") or "").strip() or None
    if len(desc) < 10:
        return jsonify({"error": "Description too short"}), 400
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE work_descriptions SET description=%s, unit=%s WHERE id=%s", (desc, unit, desc_id))
    conn.commit()
    conn.close()
    return jsonify({"ok": True})

# ── Frontend ──────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    if "user_id" not in session:
        return redirect(url_for("login_page"))
    return render_template("index.html",
                           username=session["username"],
                           role=session["role"])

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    app.run(debug=False, host="0.0.0.0", port=port)
