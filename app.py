import io
import os
import json
import queue
import base64
import socket
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, jsonify, Response, send_file
import qrcode
from openpyxl import Workbook, load_workbook

app = Flask(__name__)
DB_PATH = "pos_data.db"
EXCEL_PATH = "Sales_Report.xlsx"
PORT = 8080

sse_clients = []

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS products (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    price INTEGER NOT NULL,
                    stock INTEGER NOT NULL DEFAULT 0
                )''')

    cols = [row[1] for row in c.execute("PRAGMA table_info(products)").fetchall()]
    if "stock" not in cols:
        c.execute("ALTER TABLE products ADD COLUMN stock INTEGER NOT NULL DEFAULT 0")

    c.execute('''CREATE TABLE IF NOT EXISTS sales (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    receipt_no TEXT NOT NULL,
                    sale_time TEXT NOT NULL,
                    code TEXT NOT NULL,
                    name TEXT NOT NULL,
                    qty INTEGER NOT NULL,
                    price INTEGER NOT NULL,
                    subtotal INTEGER NOT NULL
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS stock_logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    log_time TEXT NOT NULL,
                    code TEXT NOT NULL,
                    name TEXT NOT NULL,
                    change_type TEXT NOT NULL,
                    qty_change INTEGER NOT NULL,
                    balance_after INTEGER NOT NULL,
                    note TEXT DEFAULT ''
                )''')
    conn.commit()
    conn.close()

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

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
            ws.append(["ရက်စွဲနှင့်အချိန်", "ဘောင်ချာနံပါတ်", "ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "အရေအတွက်", "တစ်ခုဈေး", "စုစုပေါင်း"])
        else:
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active
        for r in rows:
            ws.append(r)
        wb.save(EXCEL_PATH)
        return True, "OK"
    except PermissionError:
        return False, "Excel ဖိုင်ကို ဖွင့်ထားသဖြင့် တိုက်ရိုက်မသွင်းနိုင်ပါ (Database ထဲတွင် သိမ်းပြီးပါပြီ)"
    except Exception as e:
        return False, str(e)

@app.route("/")
def pos_page():
    ip = get_local_ip()
    scanner_url = f"http://{ip}:{PORT}/scanner"
    scanner_qr = make_qr_base64(scanner_url)
    return render_template("pos.html", scanner_url=scanner_url, scanner_qr=scanner_qr)

@app.route("/admin")
def admin_page():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    products = conn.execute("SELECT * FROM products ORDER BY id DESC").fetchall()
    sales = conn.execute("SELECT * FROM sales ORDER BY id DESC LIMIT 200").fetchall()
    in_logs = conn.execute("SELECT * FROM stock_logs WHERE change_type = 'ပစ္စည်းအဝင်' ORDER BY id DESC LIMIT 200").fetchall()
    dmg_logs = conn.execute("SELECT * FROM stock_logs WHERE change_type = 'အပျက်အစီးနှုတ်' ORDER BY id DESC LIMIT 200").fetchall()
    total_sales_amt = conn.execute("SELECT COALESCE(SUM(subtotal), 0) FROM sales").fetchone()[0]
    conn.close()

    items = []
    for p in products:
        items.append({
            "id": p["id"],
            "code": p["code"],
            "name": p["name"],
            "price": p["price"],
            "stock": p["stock"],
            "qr": make_qr_base64(p["code"])
        })
    return render_template(
        "admin.html",
        products=items,
        sales=sales,
        in_logs=in_logs,
        dmg_logs=dmg_logs,
        total_sales_amt=total_sales_amt
    )

@app.route("/scanner")
def scanner_page():
    return render_template("scanner.html")

@app.route("/api/products", methods=["POST"])
def add_product():
    data = request.json
    name = data.get("name", "").strip()
    price = int(data.get("price", 0))
    stock = int(data.get("stock", 0) or 0)
    custom_code = data.get("code", "").strip()

    if not name or price <= 0 or stock < 0:
        return jsonify({"status": "error", "message": "ပစ္စည်းအမည်၊ ဈေးနှုန်းနှင့် လက်ကျန် မှန်ကန်စွာထည့်ပါ"}), 400

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    if not custom_code:
        c.execute("SELECT MAX(id) FROM products")
        max_id = c.fetchone()[0] or 0
        custom_code = f"P{max_id + 1:04d}"

    try:
        c.execute("INSERT INTO products (code, name, price, stock) VALUES (?, ?, ?, ?)", (custom_code, name, price, stock))
        if stock > 0:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            c.execute(
                "INSERT INTO stock_logs (log_time, code, name, change_type, qty_change, balance_after, note) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (now_str, custom_code, name, "ပစ္စည်းအဝင်", stock, stock, "စတင်စာရင်းသွင်း")
            )
        conn.commit()
    except sqlite3.IntegrityError:
        conn.close()
        return jsonify({"status": "error", "message": "ဒီပစ္စည်းကုတ် ရှိပြီးသားဖြစ်နေပါတယ်"}), 400
    conn.close()
    return jsonify({"status": "ok", "code": custom_code})

@app.route("/api/products/<int:pid>", methods=["PUT"])
def update_product_price(pid):
    data = request.json
    new_price = int(data.get("price", 0) or 0)
    if new_price <= 0:
        return jsonify({"status": "error", "message": "ဈေးနှုန်း မှန်ကန်စွာထည့်ပါ"}), 400
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE products SET price = ? WHERE id = ?", (new_price, pid))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok", "new_price": new_price})

@app.route("/api/products/<int:pid>", methods=["DELETE"])
def delete_product(pid):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM products WHERE id = ?", (pid,))
    conn.commit()
    conn.close()
    return jsonify({"status": "ok"})

@app.route("/api/stock/adjust", methods=["POST"])
def adjust_stock():
    data = request.json
    pid = data.get("product_id")
    action = data.get("action")
    qty = int(data.get("qty", 0) or 0)
    note = data.get("note", "").strip()

    if qty <= 0:
        return jsonify({"status": "error", "message": "အရေအတွက် ၁ နှင့်အထက် ထည့်ပါ"}), 400

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    product = c.execute("SELECT * FROM products WHERE id = ?", (pid,)).fetchone()
    if not product:
        conn.close()
        return jsonify({"status": "error", "message": "ပစ္စည်းမတွေ့ပါ"}), 404

    current_stock = product["stock"]
    if action == "in":
        new_stock = current_stock + qty
        qty_change = qty
        change_type = "ပစ္စည်းအဝင်"
    elif action == "damage":
        if qty > current_stock:
            conn.close()
            return jsonify({"status": "error", "message": f"လက်ကျန် ({current_stock}) ထက် ပိုနှုတ်၍မရပါ"}), 400
        new_stock = current_stock - qty
        qty_change = -qty
        change_type = "အပျက်အစီးနှုတ်"
    else:
        conn.close()
        return jsonify({"status": "error", "message": "လုပ်ဆောင်ချက် မမှန်ကန်ပါ"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("UPDATE products SET stock = ? WHERE id = ?", (new_stock, pid))
    c.execute(
        "INSERT INTO stock_logs (log_time, code, name, change_type, qty_change, balance_after, note) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (now_str, product["code"], product["name"], change_type, qty_change, new_stock, note)
    )
    conn.commit()
    conn.close()
    return jsonify({"status": "ok", "new_stock": new_stock})

@app.route("/api/scan", methods=["POST"])
def handle_scan():
    code = request.json.get("code", "").strip()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    product = conn.execute("SELECT * FROM products WHERE code = ?", (code,)).fetchone()
    conn.close()

    if not product:
        return jsonify({"status": "not_found", "message": f"ကုတ် ({code}) ကို စာရင်းထဲတွင် မတွေ့ပါ"}), 404

    payload = {
        "code": product["code"],
        "name": product["name"],
        "price": product["price"],
        "stock": product["stock"]
    }
    for q in list(sse_clients):
        q.put(json.dumps(payload, ensure_ascii=False))

    return jsonify({"status": "ok", "product": payload})

@app.route("/stream")
def stream():
    def event_stream():
        q = queue.Queue()
        sse_clients.append(q)
        try:
            while True:
                try:
                    data = q.get(timeout=20)
                    yield f"data: {data}\n\n"
                except queue.Empty:
                    yield ": keepalive\n\n"
        finally:
            if q in sse_clients:
                sse_clients.remove(q)
    return Response(event_stream(), mimetype="text/event-stream")

@app.route("/api/checkout", methods=["POST"])
def checkout():
    data = request.json
    cart = data.get("cart", [])
    if not cart:
        return jsonify({"status": "error", "message": "ပစ္စည်းမရှိပါ"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    receipt_no = "INV-" + datetime.now().strftime("%Y%m%d%H%M%S")

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()
    excel_rows = []

    for item in cart:
        qty = int(item["qty"])
        price = int(item["price"])
        subtotal = qty * price

        c.execute(
            "INSERT INTO sales (receipt_no, sale_time, code, name, qty, price, subtotal) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (receipt_no, now_str, item["code"], item["name"], qty, price, subtotal)
        )

        prod = c.execute("SELECT id, stock FROM products WHERE code = ?", (item["code"],)).fetchone()
        if prod:
            new_stock = prod["stock"] - qty
            c.execute("UPDATE products SET stock = ? WHERE id = ?", (new_stock, prod["id"]))
            c.execute(
                "INSERT INTO stock_logs (log_time, code, name, change_type, qty_change, balance_after, note) VALUES (?, ?, ?, ?, ?, ?, ?)",
                (now_str, item["code"], item["name"], "အရောင်း", -qty, new_stock, receipt_no)
            )

        excel_rows.append([now_str, receipt_no, item["code"], item["name"], qty, price, subtotal])

    conn.commit()
    conn.close()

    excel_ok, excel_msg = append_to_excel(excel_rows)
    return jsonify({
        "status": "ok",
        "receipt_no": receipt_no,
        "sale_time": now_str,
        "excel_ok": excel_ok,
        "excel_msg": excel_msg
    })

@app.route("/api/export_excel")
def export_excel():
    conn = sqlite3.connect(DB_PATH)
    sales_rows = conn.execute("SELECT sale_time, receipt_no, code, name, qty, price, subtotal FROM sales ORDER BY id DESC").fetchall()
    stock_rows = conn.execute("SELECT code, name, price, stock, (price * stock) FROM products ORDER BY id DESC").fetchall()
    log_rows = conn.execute("SELECT log_time, code, name, change_type, qty_change, balance_after, note FROM stock_logs ORDER BY id DESC").fetchall()
    conn.close()

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Sales_All"
    ws1.append(["ရက်စွဲနှင့်အချိန်", "ဘောင်ချာနံပါတ်", "ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "အရေအတွက်", "တစ်ခုဈေး", "စုစုပေါင်း"])
    for r in sales_rows:
        ws1.append(list(r))

    ws2 = wb.create_sheet("Stock_Balance")
    ws2.append(["ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "ရောင်းဈေး", "လက်ကျန်အရေအတွက်", "လက်ကျန်တန်ဖိုး"])
    for r in stock_rows:
        ws2.append(list(r))

    ws3 = wb.create_sheet("Stock_Logs")
    ws3.append(["ရက်စွဲနှင့်အချိန်", "ပစ္စည်းကုတ်", "ပစ္စည်းအမည်", "အမျိုးအစား", "အတိုး/အလျော့", "ကျန်ရှိလက်ကျန်", "မှတ်ချက်"])
    for r in log_rows:
        ws3.append(list(r))

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return send_file(out, download_name="POS_Full_Report.xlsx", as_attachment=True)

if __name__ == "__main__":
    init_db()
    ip = get_local_ip()
    print("=" * 55)
    print(f"PC POS Counter : http://127.0.0.1:{PORT}")
    print(f"Admin & Sticker: http://127.0.0.1:{PORT}/admin")
    print(f"Phone Scanner  : http://{ip}:{PORT}/scanner")
    print("=" * 55)
    app.run(host="0.0.0.0", port=PORT, threaded=True)
