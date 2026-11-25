import io
import os
import json
import queue
import base64
import sqlite3
from functools import wraps
from datetime import datetime
from flask import Flask, render_template, request, jsonify, Response, send_file, session, redirect, url_for
from werkzeug.middleware.proxy_fix import ProxyFix
import qrcode
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=1)
app.secret_key = os.environ.get("SECRET_KEY", "pos-multibranch-secret-2026")

ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "123456")

DB_PATH = "pos_data.db"
EXCEL_PATH = "Sales_Report.xlsx"
PORT = int(os.environ.get("PORT", 8080))

sse_clients = {}

def get_db():
    conn = sqlite3.connect(DB_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()

    # ၁။ ဆိုင်ခွဲများ ဇယား (ဆိုင်ခွဲတစ်ခုချင်းစီအတွက် Password ပါဝင်သည်)
    c.execute('''CREATE TABLE IF NOT EXISTS branches (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT UNIQUE NOT NULL,
                    name TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL DEFAULT '1111'
                )''')

    b_cols = [r[1] for r in c.execute("PRAGMA table_info(branches)").fetchall()]
    if "password" not in b_cols:
        c.execute("ALTER TABLE branches ADD COLUMN password TEXT NOT NULL DEFAULT '1111'")

    if c.execute("SELECT COUNT(*) FROM branches").fetchone()[0] == 0:
        for i in range(1, 6):
            c.execute(
                "INSERT INTO branches (code, name, password) VALUES (?, ?, ?)",
                (f"shop{i}", f"ဆိုင်ခွဲ {i}", f"{i}{i}{i}{i}")
            )

    # ၂။ ပစ္စည်းနှင့် ဈေးနှုန်း ဇယား (Master Price)
    c.execute('''CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    price INTEGER NOT NULL,
                    stock INTEGER NOT NULL DEFAULT 0
                )''')

    # ၃။ ဆိုင်ခွဲအလိုက် လက်ကျန် (Stock) ဇယား
    c.execute('''CREATE TABLE IF NOT EXISTS branch_stocks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    branch_id INTEGER NOT NULL,
                    product_id INTEGER NOT NULL,
                    stock INTEGER NOT NULL DEFAULT 0,
                    UNIQUE(branch_id, product_id)
                )''')

    # ၄။ အရောင်းစာရင်း ဇယား
    c.execute('''CREATE TABLE IF NOT EXISTS sales (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    branch_name TEXT NOT NULL DEFAULT 'ဆိုင်ခွဲ 1',
                    receipt_no TEXT NOT NULL,
                    sale_time TEXT NOT NULL,
                    code TEXT NOT NULL,
                    name TEXT NOT NULL,
                    qty INTEGER NOT NULL,
                    price INTEGER NOT NULL,
                    subtotal INTEGER NOT NULL
                )''')

    # ၅။ Stock အဝင်/အရောင်း/အပျက် မှတ်တမ်း ဇယား
    c.execute('''CREATE TABLE IF NOT EXISTS stock_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    branch_name TEXT NOT NULL DEFAULT 'ဆိုင်ခွဲ 1',
                    log_time TEXT NOT NULL,
                    code TEXT NOT NULL,
                    name TEXT NOT NULL,
                    change_type TEXT NOT NULL,
                    qty_change INTEGER NOT NULL,
                    balance_after INTEGER NOT NULL,
                    note TEXT DEFAULT ''
                )''')

    branches = c.execute("SELECT id FROM branches ORDER BY id ASC").fetchall()
    products = c.execute("SELECT id, stock FROM products").fetchall()
    for idx, b in enumerate(branches):
        for p in products:
            init_st = p["stock"] if idx == 0 else 0
            c.execute(
                "INSERT OR IGNORE INTO branch_stocks (branch_id, product_id, stock) VALUES (?, ?, ?)",
                (b["id"], p["id"], init_st)
            )

    conn.commit()
    conn.close()

init_db()

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("is_admin"):
            if request.path.startswith("/api/"):
                return jsonify({"status": "error", "message": "Admin Login ဝင်ရန် လိုအပ်ပါသည်"}), 401
            return redirect(url_for("admin_login"))
        return f(*args, **kwargs)
    return decorated

def has_branch_access(branch_code):
    if session.get("is_admin"):
        return True
    allowed = session.get("allowed_branches", [])
    return branch_code in allowed

def make_qr_base64(text):
    qr = qrcode.QRCode(version=1, box_size=6, border=2)
    qr.add_data(text)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")

def append_to_excel(rows):
    try:
        if not os.path.exists(EXCEL_PATH):
            wb = Workbook()
            ws = wb.active
            ws.title = "Sales"
            ws.append(["ဆိုင်ခွဲ", "ရက်စွဲနှင့်အချိန်", "ဘောင်ချာနံပါတ်", "ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "အရေအတွက်", "တစ်ခုဈေး", "စုစုပေါင်း"])
        else:
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active
        for r in rows:
            ws.append(r)
        wb.save(EXCEL_PATH)
        return True, "OK"
    except Exception as e:
        return False, str(e)

# --- ပင်မစာမျက်နှာ (ဆိုင်ခွဲ ရွေးချယ်ရန်) ---
@app.route("/")
def index_page():
    conn = get_db()
    branches = conn.execute("SELECT * FROM branches ORDER BY id ASC").fetchall()
    conn.close()
    cards = "".join([
        f'<a href="/pos/{b["code"]}" class="branch-card">🔒 {b["name"]}</a>'
        for b in branches
    ])
    return f'''
    <!DOCTYPE html>
    <html lang="my"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>POS ဆိုင်ခွဲ ရွေးချယ်ရန်</title>
    <style>
      body {{ font-family: 'Segoe UI', Pyidaungsu, sans-serif; background: #0f172a; color: #fff; margin: 0; padding: 30px 15px; text-align: center; }}
      .wrap {{ max-width: 750px; margin: auto; }}
      .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px; margin: 25px 0; }}
      .branch-card {{
        display: block; background: #1e293b; color: #fff; text-decoration: none;
        padding: 24px 16px; border-radius: 12px; font-size: 20px; font-weight: bold;
        border: 2px solid #334155; transition: 0.2s;
      }}
      .branch-card:hover {{ background: #2563eb; border-color: #60a5fa; transform: translateY(-2px); }}
      .admin-link {{ display: inline-block; margin-top: 15px; padding: 12px 24px; background: #0f766e; color: #fff; text-decoration: none; border-radius: 8px; font-weight: bold; }}
    </style></head>
    <body><div class="wrap">
      <h2>POS အရောင်းကောင်တာ (ဆိုင်ခွဲ ရွေးချယ်ပါ)</h2>
      <p style="color:#94a3b8; font-size:14px;">ဆိုင်ခွဲတစ်ခုချင်းစီအတွက် သတ်မှတ်ထားသော Password ဖြင့်သာ ဝင်ရောက်နိုင်ပါသည်</p>
      <div class="grid">{cards}</div>
      <a href="/admin" class="admin-link">Admin စီမံခန့်ခွဲမှု ဝင်ရောက်ရန်</a>
    </div></body></html>
    '''

# --- ဆိုင်ခွဲ Login / Logout ---
@app.route("/pos/<branch_code>/login", methods=["GET", "POST"])
def branch_login(branch_code):
    conn = get_db()
    branch = conn.execute("SELECT * FROM branches WHERE code = ?", (branch_code,)).fetchone()
    conn.close()
    if not branch:
        return redirect(url_for("index_page"))

    error = ""
    if request.method == "POST":
        pw = request.form.get("password", "").strip()
        if pw == branch["password"] or pw == ADMIN_PASSWORD:
            allowed = session.get("allowed_branches", [])
            if branch_code not in allowed:
                allowed.append(branch_code)
            session["allowed_branches"] = allowed
            return redirect(url_for("pos_page", branch_code=branch_code))
        error = "ဆိုင်ခွဲ Password မှားယွင်းနေပါသည်"

    return f'''
    <!DOCTYPE html>
    <html lang="my"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{branch["name"]} - Login</title>
    <style>
      body {{ font-family: 'Segoe UI', Pyidaungsu, sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }}
      .box {{ background: #1e293b; padding: 30px; border-radius: 12px; width: 330px; text-align: center; border: 1px solid #334155; }}
      input {{ width: 100%; padding: 12px; margin: 12px 0; border-radius: 8px; border: none; box-sizing: border-box; font-size: 16px; }}
      button {{ width: 100%; padding: 12px; background: #2563eb; color: #fff; border: none; border-radius: 8px; font-size: 16px; font-weight: bold; cursor: pointer; }}
      .err {{ color: #f87171; font-size: 14px; margin-bottom: 8px; }}
    </style></head>
    <body><div class="box">
      <h3>🏪 {branch["name"]} သို့ ဝင်ရောက်ရန်</h3>
      <div class="err">{error}</div>
      <form method="POST">
        <input type="password" name="password" placeholder="ဆိုင်ခွဲ Password ထည့်ပါ" required autofocus>
        <button type="submit">ကောင်တာ ဖွင့်မည်</button>
      </form>
      <p style="margin-top:15px;"><a href="/" style="color:#94a3b8; font-size:13px;">← ဆိုင်ခွဲရွေးချယ်ရန် စာမျက်နှာသို့</a></p>
    </div></body></html>
    '''

@app.route("/pos/<branch_code>/logout")
def branch_logout(branch_code):
    allowed = session.get("allowed_branches", [])
    if branch_code in allowed:
        allowed.remove(branch_code)
    session["allowed_branches"] = allowed
    return redirect(url_for("branch_login", branch_code=branch_code))

# --- ဆိုင်ခွဲအလိုက် POS ကောင်တာ (Password စစ်ဆေးပြီးမှ ဝင်ခွင့်ပြုမည်) ---
@app.route("/pos/<branch_code>")
def pos_page(branch_code):
    if not has_branch_access(branch_code):
        return redirect(url_for("branch_login", branch_code=branch_code))

    conn = get_db()
    branch = conn.execute("SELECT * FROM branches WHERE code = ?", (branch_code,)).fetchone()
    if not branch:
        conn.close()
        return redirect(url_for("index_page"))

    products = conn.execute('''
        SELECT p.id, p.code, p.name, p.price, COALESCE(bs.stock, 0) AS stock
        FROM products p
        LEFT JOIN branch_stocks bs
            ON bs.product_id = p.id AND bs.branch_id = ?
        ORDER BY p.id DESC
    ''', (branch["id"],)).fetchall()
    conn.close()

    product_list = [dict(row) for row in products]
    base_url = request.host_url.rstrip("/")
    scanner_url = f"{base_url}/scanner/{branch['code']}"
    scanner_qr = make_qr_base64(scanner_url)
    return render_template(
        "pos.html",
        branch=branch,
        products=product_list,
        scanner_url=scanner_url,
        scanner_qr=scanner_qr
    )

# --- ဆိုင်ခွဲအလိုက် သီးသန့် စာရင်း (၄) မျိုး ကြည့်ရှုရန် စာမျက်နှာ ---
@app.route("/pos/<branch_code>/reports")
def branch_reports(branch_code):
    if not has_branch_access(branch_code):
        return redirect(url_for("branch_login", branch_code=branch_code))

    conn = get_db()
    branch = conn.execute("SELECT * FROM branches WHERE code = ?", (branch_code,)).fetchone()
    if not branch:
        conn.close()
        return redirect(url_for("index_page"))

    # ၁။ ဤဆိုင်ခွဲ၏ လက်ကျန်စာရင်း
    stocks = conn.execute('''
        SELECT p.code, p.name, p.price, COALESCE(bs.stock, 0) AS stock, (p.price * COALESCE(bs.stock, 0)) AS total_val
        FROM products p
        LEFT JOIN branch_stocks bs ON bs.product_id = p.id AND bs.branch_id = ?
        ORDER BY p.id DESC
    ''', (branch["id"],)).fetchall()

    # ၂။ ဤဆိုင်ခွဲ၏ အရောင်းစာရင်း
    sales = conn.execute(
        "SELECT * FROM sales WHERE branch_name = ? ORDER BY id DESC LIMIT 300",
        (branch["name"],)
    ).fetchall()
    total_sales_amt = conn.execute(
        "SELECT COALESCE(SUM(subtotal), 0) FROM sales WHERE branch_name = ?",
        (branch["name"],)
    ).fetchone()[0]

    # ၃။ ဤဆိုင်ခွဲ၏ ပစ္စည်းအဝင်စာရင်း
    in_logs = conn.execute(
        "SELECT * FROM stock_logs WHERE branch_name = ? AND change_type = 'ပစ္စည်းအဝင်' ORDER BY id DESC LIMIT 300",
        (branch["name"],)
    ).fetchall()

    # ၄။ ဤဆိုင်ခွဲ၏ အပျက်အစီးစာရင်း
    dmg_logs = conn.execute(
        "SELECT * FROM stock_logs WHERE branch_name = ? AND change_type = 'အပျက်အစီးနှုတ်' ORDER BY id DESC LIMIT 300",
        (branch["name"],)
    ).fetchall()
    conn.close()

    return render_template(
        "branch_reports.html",
        branch=branch,
        stocks=stocks,
        sales=sales,
        total_sales_amt=total_sales_amt,
        in_logs=in_logs,
        dmg_logs=dmg_logs
    )

@app.route("/scanner/<branch_code>")
def scanner_page(branch_code):
    conn = get_db()
    branch = conn.execute("SELECT * FROM branches WHERE code = ?", (branch_code,)).fetchone()
    conn.close()
    if not branch:
        return "ဆိုင်ခွဲ မတွေ့ပါ", 404
    return render_template("scanner.html", branch=branch)

# --- Admin Login / Logout ---
@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    error = ""
    if request.method == "POST":
        if request.form.get("password") == ADMIN_PASSWORD:
            session["is_admin"] = True
            return redirect(url_for("admin_page"))
        error = "Password မှားယွင်းနေပါသည်"
    return f'''
    <!DOCTYPE html>
    <html lang="my"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Admin Login</title>
    <style>
      body {{ font-family: sans-serif; background: #0f172a; color: #fff; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }}
      .box {{ background: #1e293b; padding: 30px; border-radius: 12px; width: 320px; text-align: center; }}
      input {{ width: 100%; padding: 12px; margin: 12px 0; border-radius: 6px; border: none; box-sizing: border-box; font-size: 16px; }}
      button {{ width: 100%; padding: 12px; background: #2563eb; color: #fff; border: none; border-radius: 6px; font-size: 16px; font-weight: bold; cursor: pointer; }}
      .err {{ color: #f87171; font-size: 14px; margin-bottom: 8px; }}
    </style></head>
    <body><div class="box">
      <h3>Admin ဝင်ရောက်ရန်</h3>
      <div class="err">{error}</div>
      <form method="POST">
        <input type="password" name="password" placeholder="Admin Password ထည့်ပါ" required autofocus>
        <button type="submit">ဝင်မည်</button>
      </form>
      <p style="margin-top:15px;"><a href="/" style="color:#94a3b8; font-size:13px;">← ဆိုင်ခွဲရွေးချယ်ရန် စာမျက်နှာသို့</a></p>
    </div></body></html>
    '''

@app.route("/admin/logout")
def admin_logout():
    session.pop("is_admin", None)
    return redirect(url_for("index_page"))

# --- Admin Dashboard ---
@app.route("/admin")
@admin_required
def admin_page():
    conn = get_db()
    branches = conn.execute("SELECT * FROM branches ORDER BY id ASC").fetchall()
    products = conn.execute("SELECT * FROM products ORDER BY id DESC").fetchall()

    stock_rows = conn.execute('''
        SELECT bs.branch_id, b.name AS branch_name, p.id AS product_id, p.code, p.name, p.price, bs.stock
        FROM branch_stocks bs
        JOIN branches b ON bs.branch_id = b.id
        JOIN products p ON bs.product_id = p.id
        ORDER BY b.id ASC, p.id DESC
    ''').fetchall()

    total_by_pid = {}
    for sr in stock_rows:
        total_by_pid[sr["product_id"]] = total_by_pid.get(sr["product_id"], 0) + sr["stock"]

    sales = conn.execute("SELECT * FROM sales ORDER BY id DESC LIMIT 300").fetchall()
    in_logs = conn.execute("SELECT * FROM stock_logs WHERE change_type = 'ပစ္စည်းအဝင်' ORDER BY id DESC LIMIT 300").fetchall()
    dmg_logs = conn.execute("SELECT * FROM stock_logs WHERE change_type = 'အပျက်အစီးနှုတ်' ORDER BY id DESC LIMIT 300").fetchall()
    conn.close()

    items = []
    for p in products:
        items.append({
            "id": p["id"],
            "code": p["code"],
            "name": p["name"],
            "price": p["price"],
            "total_stock": total_by_pid.get(p["id"], 0),
            "qr": make_qr_base64(p["code"])
        })

    base_url = request.host_url.rstrip("/")
    return render_template(
        "admin.html",
        branches=branches,
        products=items,
        stock_rows=stock_rows,
        sales=sales,
        in_logs=in_logs,
        dmg_logs=dmg_logs,
        base_url=base_url
    )

# --- API: ဆိုင်ခွဲ အသစ်ထည့်ခြင်း၊ Password ပြင်ခြင်း၊ ဖျက်ခြင်း ---
@app.route("/api/branches", methods=["POST"])
@admin_required
def add_branch():
    data = request.json
    name = data.get("name", "").strip()
    custom_code = data.get("code", "").strip().lower()
    password = data.get("password", "").strip() or "1111"
    if not name:
        return jsonify({"status": "error", "message": "ဆိုင်ခွဲအမည် ထည့်ပါ"}), 400

    conn = get_db()
    c = conn.cursor()
    if not custom_code:
        max_id = c.execute("SELECT COALESCE(MAX(id), 0) FROM branches").fetchone()[0]
        custom_code = f"shop{max_id + 1}"

    try:
        c.execute("INSERT INTO branches (code, name, password) VALUES (?, ?, ?)", (custom_code, name, password))
        new_bid = c.lastrowid
        products = c.execute("SELECT id FROM products").fetchall()
        for p in products:
            c.execute("INSERT OR IGNORE INTO branch_stocks (branch_id, product_id, stock) VALUES (?, ?, 0)", (new_bid, p["id"]))
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"status": "error", "message": "ဒီဆိုင်ခွဲအမည် သို့မဟုတ် ကုတ် ရှိပြီးသားဖြစ်နေပါတယ်"}), 400
    conn.close()
    return jsonify({"status": "ok"})

@app.route("/api/branches/<int:bid>/password", methods=["PUT"])
@admin_required
def update_branch_password(bid):
    data = request.json
    new_pw = data.get("password", "").strip()
    if not new_pw:
        return jsonify({"status": "error", "message": "Password အသစ် ထည့်ပါ"}), 400
    conn = get_db()
    conn.execute("UPDATE branches SET password = ? WHERE id = ?", (new_pw, bid))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})

@app.route("/api/branches/<int:bid>", methods=["DELETE"])
@admin_required
def delete_branch(bid):
    conn = get_db()
    count = conn.execute("SELECT COUNT(*) FROM branches").fetchone()[0]
    if count <= 1:
        conn.close()
        return jsonify({"status": "error", "message": "အနည်းဆုံး ဆိုင်ခွဲ ၁ ခု ရှိနေရပါမယ်"}), 400
    conn.execute("DELETE FROM branch_stocks WHERE branch_id = ?", (bid,))
    conn.execute("DELETE FROM branches WHERE id = ?", (bid,))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})

# --- API: ပစ္စည်းအသစ်ထည့်ခြင်း၊ ဈေးပြင်ခြင်း၊ ဖျက်ခြင်း ---
@app.route("/api/products", methods=["POST"])
@admin_required
def add_product():
    data = request.json
    name = data.get("name", "").strip()
    price = int(data.get("price", 0) or 0)
    init_stock = int(data.get("stock", 0) or 0)
    custom_code = data.get("code", "").strip()

    if not name or price <= 0 or init_stock < 0:
        return jsonify({"status": "error", "message": "ပစ္စည်းအမည်၊ ဈေးနှုန်းနှင့် လက်ကျန် မှန်ကန်စွာထည့်ပါ"}), 400

    conn = get_db()
    c = conn.cursor()
    if not custom_code:
        max_id = c.execute("SELECT COALESCE(MAX(id), 0) FROM products").fetchone()[0]
        custom_code = f"P{max_id + 1:04d}"

    try:
        c.execute("INSERT INTO products (code, name, price, stock) VALUES (?, ?, ?, ?)", (custom_code, name, price, init_stock))
        pid = c.lastrowid
        branches = c.execute("SELECT id, name FROM branches ORDER BY id ASC").fetchall()
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for b in branches:
            c.execute("INSERT INTO branch_stocks (branch_id, product_id, stock) VALUES (?, ?, ?)", (b["id"], pid, init_stock))
            if init_stock > 0:
                c.execute(
                    "INSERT INTO stock_logs (branch_name, log_time, code, name, change_type, qty_change, balance_after, note) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (b["name"], now_str, custom_code, name, "ပစ္စည်းအဝင်", init_stock, init_stock, "စတင်စာရင်းသွင်း")
                )
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"status": "error", "message": "ဒီပစ္စည်းကုတ် ရှိပြီးသားဖြစ်နေပါတယ်"}), 400
    conn.close()
    return jsonify({"status": "ok", "code": custom_code})

@app.route("/api/products/<int:pid>", methods=["PUT"])
@admin_required
def update_product_price(pid):
    data = request.json
    new_price = int(data.get("price", 0) or 0)
    if new_price <= 0:
        return jsonify({"status": "error", "message": "ဈေးနှုန်း မှန်ကန်စွာထည့်ပါ"}), 400
    conn = get_db()
    conn.execute("UPDATE products SET price = ? WHERE id = ?", (new_price, pid))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})

@app.route("/api/products/<int:pid>", methods=["DELETE"])
@admin_required
def delete_product(pid):
    conn = get_db()
    conn.execute("DELETE FROM branch_stocks WHERE product_id = ?", (pid,))
    conn.execute("DELETE FROM products WHERE id = ?", (pid,))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})

# --- API: ဆိုင်ခွဲအလိုက် Stock အဝင် / အပျက်အစီး နှုတ်ခြင်း ---
@app.route("/api/stock/adjust", methods=["POST"])
@admin_required
def adjust_stock():
    data = request.json
    bid = int(data.get("branch_id"))
    pid = int(data.get("product_id"))
    action = data.get("action")
    qty = int(data.get("qty", 0) or 0)
    note = data.get("note", "").strip()

    if qty <= 0:
        return jsonify({"status": "error", "message": "အရေအတွက် ၁ နှင့်အထက် ထည့်ပါ"}), 400

    conn = get_db()
    c = conn.cursor()
    branch = c.execute("SELECT * FROM branches WHERE id = ?", (bid,)).fetchone()
    product = c.execute("SELECT * FROM products WHERE id = ?", (pid,)).fetchone()
    bs = c.execute("SELECT * FROM branch_stocks WHERE branch_id = ? AND product_id = ?", (bid, pid)).fetchone()

    if not branch or not product or not bs:
        conn.close()
        return jsonify({"status": "error", "message": "အချက်အလက် မတွေ့ပါ"}), 404

    current_stock = bs["stock"]
    if action == "in":
        new_stock = current_stock + qty
        qty_change = qty
        change_type = "ပစ္စည်းအဝင်"
    elif action == "damage":
        if qty > current_stock:
            conn.close()
            return jsonify({"status": "error", "message": f"[{branch['name']}] လက်ကျန် ({current_stock}) ထက် ပိုနှုတ်၍မရပါ"}), 400
        new_stock = current_stock - qty
        qty_change = -qty
        change_type = "အပျက်အစီးနှုတ်"
    else:
        conn.close()
        return jsonify({"status": "error", "message": "လုပ်ဆောင်ချက် မမှန်ကန်ပါ"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("UPDATE branch_stocks SET stock = ? WHERE id = ?", (new_stock, bs["id"]))
    c.execute(
        "INSERT INTO stock_logs (branch_name, log_time, code, name, change_type, qty_change, balance_after, note) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (branch["name"], now_str, product["code"], product["name"], change_type, qty_change, new_stock, note)
    )
    conn.commit()
    conn.close()
    return jsonify({"status": "ok", "new_stock": new_stock})

# --- API: Scan နှင့် Real-time Stream ---
@app.route("/api/scan", methods=["POST"])
def handle_scan():
    data = request.json
    code = data.get("code", "").strip()
    branch_code = data.get("branch_code", "shop1").strip()

    conn = get_db()
    product = conn.execute("SELECT * FROM products WHERE code = ?", (code,)).fetchone()
    conn.close()

    if not product:
        return jsonify({"status": "not_found", "message": f"ကုတ် ({code}) ကို စာရင်းထဲတွင် မတွေ့ပါ"}), 404

    payload = {
        "code": product["code"],
        "name": product["name"],
        "price": product["price"]
    }
    for q in list(sse_clients.get(branch_code, [])):
        q.put(json.dumps(payload, ensure_ascii=False))

    return jsonify({"status": "ok", "product": payload})

@app.route("/stream/<branch_code>")
def stream(branch_code):
    def event_stream():
        q = queue.Queue()
        sse_clients.setdefault(branch_code, []).append(q)
        try:
            while True:
                try:
                    data = q.get(timeout=20)
                    yield f"data: {data}\n\n"
                except queue.Empty:
                    yield ": keepalive\n\n"
        finally:
            if q in sse_clients.get(branch_code, []):
                sse_clients[branch_code].remove(q)

    resp = Response(event_stream(), mimetype="text/event-stream")
    resp.headers["Cache-Control"] = "no-cache"
    resp.headers["X-Accel-Buffering"] = "no"
    return resp

# --- API: ငွေရှင်းခြင်း ---
@app.route("/api/checkout", methods=["POST"])
def checkout():
    data = request.json
    cart = data.get("cart", [])
    branch_code = data.get("branch_code", "shop1").strip()
    if not has_branch_access(branch_code):
        return jsonify({"status": "error", "message": "ဆိုင်ခွဲ Login ဝင်ရန် လိုအပ်ပါသည်"}), 401
    if not cart:
        return jsonify({"status": "error", "message": "ပစ္စည်းမရှိပါ"}), 400

    conn = get_db()
    c = conn.cursor()
    branch = c.execute("SELECT * FROM branches WHERE code = ?", (branch_code,)).fetchone()
    if not branch:
        conn.close()
        return jsonify({"status": "error", "message": "ဆိုင်ခွဲ မတွေ့ပါ"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    receipt_no = f"{branch['code'].upper()}-" + datetime.now().strftime("%Y%m%d%H%M%S")
    excel_rows = []

    for item in cart:
        qty = int(item["qty"])
        price = int(item["price"])
        subtotal = qty * price

        c.execute(
            "INSERT INTO sales (branch_name, receipt_no, sale_time, code, name, qty, price, subtotal) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (branch["name"], receipt_no, now_str, item["code"], item["name"], qty, price, subtotal)
        )

        prod = c.execute("SELECT id FROM products WHERE code = ?", (item["code"],)).fetchone()
        if prod:
            bs = c.execute("SELECT id, stock FROM branch_stocks WHERE branch_id = ? AND product_id = ?", (branch["id"], prod["id"])).fetchone()
            if bs:
                new_stock = bs["stock"] - qty
                c.execute("UPDATE branch_stocks SET stock = ? WHERE id = ?", (new_stock, bs["id"]))
                c.execute(
                    "INSERT INTO stock_logs (branch_name, log_time, code, name, change_type, qty_change, balance_after, note) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                    (branch["name"], now_str, item["code"], item["name"], "အရောင်း", -qty, new_stock, receipt_no)
                )

        excel_rows.append([branch["name"], now_str, receipt_no, item["code"], item["name"], qty, price, subtotal])

    conn.commit()
    conn.close()

    excel_ok, excel_msg = append_to_excel(excel_rows)
    return jsonify({
        "status": "ok",
        "branch_name": branch["name"],
        "receipt_no": receipt_no,
        "sale_time": now_str,
        "excel_ok": excel_ok,
        "excel_msg": excel_msg
    })

@app.route("/api/export_excel")
@admin_required
def export_excel():
    conn = get_db()
    sales_rows = conn.execute("SELECT branch_name, sale_time, receipt_no, code, name, qty, price, subtotal FROM sales ORDER BY id DESC").fetchall()
    stock_rows = conn.execute('''
        SELECT b.name, p.code, p.name, p.price, bs.stock, (p.price * bs.stock)
        FROM branch_stocks bs
        JOIN branches b ON bs.branch_id = b.id
        JOIN products p ON bs.product_id = p.id
        ORDER BY b.id ASC, p.id DESC
    ''').fetchall()
    log_rows = conn.execute("SELECT branch_name, log_time, code, name, change_type, qty_change, balance_after, note FROM stock_logs ORDER BY id DESC").fetchall()
    conn.close()

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sales_All"
    ws1.append(["ဆိုင်ခွဲ", "ရက်စွဲနှင့်အချိန်", "ဘောင်ချာနံပါတ်", "ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "အရေအတွက်", "တစ်ခုဈေး", "စုစုပေါင်း"])
    for r in sales_rows:
        ws1.append(list(r))

    ws2 = wb.create_sheet("Stock_Balance")
    ws2.append(["ဆိုင်ခွဲ", "ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "ရောင်းဈေး", "လက်ကျန်အရေအတွက်", "လက်ကျန်တန်ဖိုး"])
    for r in stock_rows:
        ws2.append(list(r))

    ws3 = wb.create_sheet("Stock_Logs")
    ws3.append(["ဆိုင်ခွဲ", "ရက်စွဲနှင့်အချိန်", "ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "အမျိုးအစား", "အတိုး/အလျော့", "ကျန်ရှိလက်ကျန်", "မှတ်ချက်"])
    for r in log_rows:
        ws3.append(list(r))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, download_name="POS_MultiBranch_Report.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=PORT, threaded=True)
