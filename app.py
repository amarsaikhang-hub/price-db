#!/usr/bin/env python3
"""
Price DB — Үнийн саналын сан
- Excel оруулах → AI-аар ангилал, төрөл, спец автомат үүсгэх
- Текст/зургаар хайхад ангилалаас харьцуулж үнэ гаргах
"""

import os
import io
import json
import uuid
import base64
import hashlib
import secrets
from pathlib import Path
from datetime import datetime
from functools import wraps
from flask import Flask, request, jsonify, session, send_file, redirect
import psycopg2
import psycopg2.extras
import openpyxl
import anthropic
from PIL import Image

app = Flask(__name__, static_folder=None)
app.secret_key = os.environ.get("SECRET_KEY", secrets.token_hex(32))

BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

DB_CONFIG = {
    "host": os.environ.get("DB_HOST", "localhost"),
    "port": os.environ.get("DB_PORT", "5432"),
    "dbname": os.environ.get("DB_NAME", "tender_db"),
    "user": os.environ.get("DB_USER", "tender_admin"),
    "password": os.environ.get("DB_PASSWORD", "admin_pass"),
}


def get_db():
    return psycopg2.connect(**DB_CONFIG)


def query(sql, params=None, fetchone=False, fetchall=False, commit=False):
    conn = get_db()
    try:
        cur = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, params)
        result = None
        if fetchone:
            result = cur.fetchone()
        elif fetchall:
            result = cur.fetchall()
        if commit:
            conn.commit()
        return result
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("CREATE EXTENSION IF NOT EXISTS vector")

    # Ангилал (Кабель, Бетон, Трансформатор...)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS product_categories (
            id SERIAL PRIMARY KEY,
            name TEXT UNIQUE NOT NULL,
            description TEXT,
            spec_schema JSONB DEFAULT '{}',
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)

    # Дэд төрөл (Кабель > ВВГ, НУМ...)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS product_types (
            id SERIAL PRIMARY KEY,
            category_id INTEGER REFERENCES product_categories(id) ON DELETE CASCADE,
            name TEXT NOT NULL,
            description TEXT,
            created_at TIMESTAMP DEFAULT NOW(),
            UNIQUE(category_id, name)
        )
    """)

    # Бүтээгдэхүүн (тодорхой нэртэй)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS products (
            id SERIAL PRIMARY KEY,
            category_id INTEGER REFERENCES product_categories(id),
            type_id INTEGER REFERENCES product_types(id),
            name TEXT NOT NULL,
            specs JSONB DEFAULT '{}',
            unit TEXT DEFAULT 'ш',
            embedding vector(1536),
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_products_cat ON products(category_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_products_type ON products(type_id)")

    # Үнэ (нийлүүлэгч эсвэл манайх)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS prices (
            id SERIAL PRIMARY KEY,
            product_id INTEGER REFERENCES products(id) ON DELETE CASCADE,
            source_type TEXT NOT NULL,  -- 'supplier' | 'own'
            source_name TEXT,
            unit_price NUMERIC(18,2),
            quantity NUMERIC(18,2) DEFAULT 1,
            currency TEXT DEFAULT 'MNT',
            price_date DATE,
            notes TEXT,
            file_id INTEGER,
            created_at TIMESTAMP DEFAULT NOW()
        )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_prices_product ON prices(product_id)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_prices_source ON prices(source_type, source_name)")

    # Файлын түүх
    cur.execute("""
        CREATE TABLE IF NOT EXISTS price_files (
            id SERIAL PRIMARY KEY,
            filename TEXT NOT NULL,
            stored_name TEXT NOT NULL,
            source_type TEXT NOT NULL,
            source_name TEXT,
            rows_total INTEGER DEFAULT 0,
            rows_imported INTEGER DEFAULT 0,
            uploaded_by INTEGER,
            uploaded_at TIMESTAMP DEFAULT NOW()
        )
    """)

    conn.commit()
    conn.close()


# ============================================================
# AUTH (tender_db users ашиглана)
# ============================================================

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return jsonify({"error": "Нэвтрэх"}), 401
        return f(*args, **kwargs)
    return decorated


@app.route("/api/login", methods=["POST"])
def login():
    from werkzeug.security import check_password_hash
    data = request.get_json()
    user = query("SELECT id, username, full_name, role, password_hash FROM users WHERE username=%s",
                 (data.get("username", "").strip(),), fetchone=True)
    if not user or not check_password_hash(user["password_hash"], data.get("password", "")):
        return jsonify({"error": "Нэр, нууц үг буруу"}), 401
    session["user_id"] = user["id"]
    session["username"] = user["username"]
    session["full_name"] = user["full_name"]
    session["role"] = user["role"]
    return jsonify({"success": True, "user": dict(user)})


@app.route("/api/me")
def me():
    if "user_id" not in session:
        return jsonify({"logged_in": False}), 401
    return jsonify({"logged_in": True, "user": {
        "username": session["username"], "full_name": session["full_name"], "role": session["role"]
    }})


@app.route("/sso")
def sso_login():
    token = request.args.get("token", "")
    if not token:
        return redirect("/")
    row = query("""
        SELECT * FROM sso_tokens
        WHERE token = %s AND expires_at > NOW()
    """, (token,), fetchone=True)
    if not row:
        return redirect("/")
    session["user_id"]  = row["user_id"]
    session["username"] = row["username"]
    session["full_name"]= row["full_name"]
    session["role"]     = row["role"]
    query("DELETE FROM sso_tokens WHERE token = %s", (token,), commit=True)
    return redirect("/")


@app.route("/api/logout", methods=["POST"])
def logout():
    session.clear()
    return jsonify({"success": True})


# ============================================================
# HELPERS
# ============================================================

def serialize(row):
    from decimal import Decimal
    if not row:
        return row
    r = {}
    for k, v in row.items():
        if isinstance(v, Decimal):
            r[k] = float(v)
        elif hasattr(v, "isoformat"):
            r[k] = v.isoformat()
        else:
            r[k] = v
    return r


def get_claude():
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    return anthropic.Anthropic(api_key=key) if key else None


def get_openai():
    key = os.environ.get("OPENAI_API_KEY", "")
    if not key:
        return None
    import openai
    return openai.OpenAI(api_key=key)


def embed_texts(texts):
    client = get_openai()
    if not client:
        return None
    resp = client.embeddings.create(model="text-embedding-3-small", input=texts)
    return [d.embedding for d in resp.data]


# ============================================================
# ANGILAL / TYPE CRUD
# ============================================================

@app.route("/api/categories")
@login_required
def list_categories():
    rows = query("""
        SELECT c.*, COUNT(p.id) as product_count
        FROM product_categories c
        LEFT JOIN products p ON p.category_id = c.id
        GROUP BY c.id ORDER BY c.name
    """, fetchall=True)
    return jsonify([serialize(r) for r in rows])


@app.route("/api/types")
@login_required
def list_types():
    cat_id = request.args.get("category_id")
    sql = """
        SELECT t.*, c.name as category_name, COUNT(p.id) as product_count
        FROM product_types t
        JOIN product_categories c ON c.id = t.category_id
        LEFT JOIN products p ON p.type_id = t.id
    """
    params = []
    if cat_id:
        sql += " WHERE t.category_id = %s"
        params.append(int(cat_id))
    sql += " GROUP BY t.id, c.name ORDER BY c.name, t.name"
    rows = query(sql, params, fetchall=True)
    return jsonify([serialize(r) for r in rows])


@app.route("/api/products")
@login_required
def list_products():
    cat_id = request.args.get("category_id")
    search = request.args.get("search", "")
    limit = min(int(request.args.get("limit", 100)), 500)
    where = ["1=1"]
    params = []
    if cat_id:
        where.append("p.category_id = %s")
        params.append(int(cat_id))
    if search:
        where.append("p.name ILIKE %s")
        params.append(f"%{search}%")
    rows = query(f"""
        SELECT p.id, p.name, p.unit, p.specs,
               c.name as category_name, t.name as type_name,
               (SELECT COUNT(*) FROM prices WHERE product_id = p.id) as price_count,
               (SELECT MIN(unit_price) FROM prices WHERE product_id = p.id AND unit_price > 0) as min_price,
               (SELECT MAX(unit_price) FROM prices WHERE product_id = p.id AND unit_price > 0) as max_price
        FROM products p
        LEFT JOIN product_categories c ON c.id = p.category_id
        LEFT JOIN product_types t ON t.id = p.type_id
        WHERE {' AND '.join(where)}
        ORDER BY p.name LIMIT %s
    """, params + [limit], fetchall=True)
    return jsonify([serialize(r) for r in rows])


@app.route("/api/products/<int:pid>")
@login_required
def get_product(pid):
    p = query("""
        SELECT p.*, c.name as category_name, t.name as type_name
        FROM products p
        LEFT JOIN product_categories c ON c.id = p.category_id
        LEFT JOIN product_types t ON t.id = p.type_id
        WHERE p.id = %s
    """, (pid,), fetchone=True)
    if not p:
        return jsonify({"error": "Олдсонгүй"}), 404
    prices = query("""
        SELECT * FROM prices WHERE product_id = %s ORDER BY price_date DESC NULLS LAST, id DESC
    """, (pid,), fetchall=True)
    return jsonify({**serialize(p), "prices": [serialize(x) for x in prices]})


# ============================================================
# EXCEL UPLOAD + AI CLASSIFICATION
# ============================================================

def read_excel_rows(path, max_rows=500):
    """Excel файлаас header + data мөрүүдийг авна"""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = []
    headers = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(c).strip() if c is not None else "" for c in row]
            continue
        if i > max_rows:
            break
        vals = [str(c).strip() if c is not None else "" for c in row]
        if any(vals):
            rows.append(vals)
    wb.close()
    return headers, rows


def ai_classify_products(client, headers, sample_rows, source_name):
    """Claude-д Excel-ийн sample өгч ангилал, төрөл, баганын mapping-г санал болгуулна"""
    sample = [headers] + sample_rows[:20]
    sample_text = "\n".join(["\t".join(r) for r in sample])

    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=3000,
        messages=[{
            "role": "user",
            "content": f"""Дараах Excel файл бол "{source_name}"-ээс ирсэн үнийн санал.

ЭХНИЙ 20 МӨР:
{sample_text}

ДААЛГАВАР:
1. Баганын индекс (0-оос эхлэнэ) ямар утга агуулж байгааг илрүүл
2. Мөр бүрт тохирох АНГИЛАЛ (category) болон ДЭД ТӨРӨЛ (type) сана
3. JSON хариул:

{{
  "columns": {{
    "name": 0,          // бүтээгдэхүүний нэр
    "unit": 1,          // нэгж (ш, кг, м, л)
    "quantity": null,   // тоо хэмжээ (байвал)
    "price": 3,         // нэгж үнэ
    "currency": null,   // валют
    "date": null,       // огноо
    "notes": null       // тайлбар
  }},
  "categories": {{
    "Кабель, утас": ["ВВГ кабель", "НУМ кабель"],
    "Трансформатор": ["Хүчний", "Тэжээлийн"]
  }}
}}

Зөвхөн JSON."""
        }]
    )

    raw = msg.content[0].text.strip()
    try:
        return json.loads(raw)
    except:
        s, e = raw.find('{'), raw.rfind('}')
        if s >= 0 and e > s:
            try:
                return json.loads(raw[s:e+1])
            except:
                pass
    return None


def ai_classify_row(client, row_text, categories):
    """Нэг барааг ангилах"""
    cat_list = "\n".join([f"- {c}: {', '.join(ts)}" for c, ts in categories.items()])
    msg = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=300,
        messages=[{
            "role": "user",
            "content": f"""Бараа: "{row_text}"

Ангилалууд:
{cat_list}

JSON хариул: {{"category": "...", "type": "..."}}. Зөвхөн JSON."""
        }]
    )
    raw = msg.content[0].text.strip()
    try:
        return json.loads(raw)
    except:
        s, e = raw.find('{'), raw.rfind('}')
        if s >= 0 and e > s:
            try:
                return json.loads(raw[s:e+1])
            except:
                pass
    return {"category": "Ангилалгүй", "type": "Бусад"}


def get_or_create_category(name, description=""):
    row = query("SELECT id FROM product_categories WHERE name = %s", (name,), fetchone=True)
    if row:
        return row["id"]
    row = query("INSERT INTO product_categories (name, description) VALUES (%s, %s) RETURNING id",
                (name, description), fetchone=True, commit=True)
    return row["id"]


def get_or_create_type(category_id, name):
    row = query("SELECT id FROM product_types WHERE category_id=%s AND name=%s", (category_id, name), fetchone=True)
    if row:
        return row["id"]
    row = query("INSERT INTO product_types (category_id, name) VALUES (%s, %s) RETURNING id",
                (category_id, name), fetchone=True, commit=True)
    return row["id"]


@app.route("/api/upload", methods=["POST"])
@login_required
def upload_excel():
    if "file" not in request.files:
        return jsonify({"error": "Файл сонгоно уу"}), 400
    source_type = request.form.get("source_type", "supplier")
    source_name = request.form.get("source_name", "").strip()
    if source_type not in ("supplier", "own"):
        return jsonify({"error": "source_type буруу"}), 400
    if not source_name:
        return jsonify({"error": "Эх үүсвэр нэр оруулна уу"}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xls")):
        return jsonify({"error": "Зөвхөн Excel файл"}), 400

    stored = f"{uuid.uuid4().hex}_{f.filename}"
    path = UPLOAD_DIR / stored
    f.save(str(path))

    client = get_claude()
    if not client:
        return jsonify({"error": "ANTHROPIC_API_KEY тохируулаагүй"}), 500

    try:
        headers, rows = read_excel_rows(str(path))
        if not rows:
            return jsonify({"error": "Мөр олдсонгүй"}), 400

        # Эхний схем тодорхойлох
        schema = ai_classify_products(client, headers, rows, source_name)
        if not schema:
            return jsonify({"error": "AI файлыг ойлгосонгүй"}), 500

        cols = schema.get("columns", {})
        categories = schema.get("categories", {})

        # Бүх ангилал, төрөл урьдчилж үүсгэх
        cat_map = {}
        type_map = {}
        for cat_name, types in categories.items():
            cat_id = get_or_create_category(cat_name)
            cat_map[cat_name] = cat_id
            for t in types:
                type_map[f"{cat_name}::{t}"] = get_or_create_type(cat_id, t)

        # File record
        file_row = query("""
            INSERT INTO price_files (filename, stored_name, source_type, source_name, rows_total, uploaded_by)
            VALUES (%s,%s,%s,%s,%s,%s) RETURNING id
        """, (f.filename, stored, source_type, source_name, len(rows), session["user_id"]),
        fetchone=True, commit=True)
        file_id = file_row["id"]

        # Мөр тус бүрийг классликашын + insert
        name_col = cols.get("name")
        price_col = cols.get("price")
        unit_col = cols.get("unit")
        qty_col = cols.get("quantity")
        curr_col = cols.get("currency")
        date_col = cols.get("date")
        notes_col = cols.get("notes")

        if name_col is None or price_col is None:
            return jsonify({"error": "Нэр эсвэл үнийн багана олдсонгүй"}), 400

        imported = 0
        for row in rows:
            try:
                name = row[name_col].strip() if name_col < len(row) else ""
                if not name:
                    continue

                # Үнэ тооцоолох
                raw_price = row[price_col] if price_col < len(row) else ""
                price_clean = ''.join(c for c in str(raw_price).replace(',', '') if c.isdigit() or c == '.')
                try:
                    price = float(price_clean) if price_clean else 0
                except:
                    price = 0

                unit = row[unit_col].strip() if unit_col is not None and unit_col < len(row) else "ш"
                qty_str = row[qty_col].strip() if qty_col is not None and qty_col < len(row) else "1"
                try:
                    qty = float(qty_str.replace(',', '')) if qty_str else 1
                except:
                    qty = 1
                currency = row[curr_col].strip() if curr_col is not None and curr_col < len(row) else "MNT"
                if not currency or currency == "":
                    currency = "MNT"
                date_val = row[date_col].strip() if date_col is not None and date_col < len(row) else None
                notes = row[notes_col].strip() if notes_col is not None and notes_col < len(row) else ""

                # AI-аар ангилах (нэгээс нэгээр — удаан, гэхдээ нарийвчлалтай)
                classification = ai_classify_row(client, name, categories)
                cat_name = classification.get("category", "Ангилалгүй")
                type_name = classification.get("type", "Бусад")

                cat_id = cat_map.get(cat_name) or get_or_create_category(cat_name)
                cat_map[cat_name] = cat_id
                tkey = f"{cat_name}::{type_name}"
                type_id = type_map.get(tkey) or get_or_create_type(cat_id, type_name)
                type_map[tkey] = type_id

                # Product байгаа эсэх
                prod = query("SELECT id FROM products WHERE category_id=%s AND type_id=%s AND name=%s",
                             (cat_id, type_id, name), fetchone=True)
                if prod:
                    pid = prod["id"]
                else:
                    prod_row = query("""
                        INSERT INTO products (category_id, type_id, name, unit)
                        VALUES (%s,%s,%s,%s) RETURNING id
                    """, (cat_id, type_id, name, unit), fetchone=True, commit=True)
                    pid = prod_row["id"]

                query("""
                    INSERT INTO prices (product_id, source_type, source_name, unit_price, quantity, currency, price_date, notes, file_id)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                """, (pid, source_type, source_name, price, qty, currency,
                      date_val if date_val else None, notes, file_id), commit=True)
                imported += 1
            except Exception:
                continue

        query("UPDATE price_files SET rows_imported=%s WHERE id=%s", (imported, file_id), commit=True)

        return jsonify({
            "success": True,
            "file_id": file_id,
            "total": len(rows),
            "imported": imported,
            "categories": list(cat_map.keys()),
            "types": list(type_map.keys())
        })
    except Exception as e:
        return jsonify({"error": f"Алдаа: {str(e)}"}), 500


# ============================================================
# EMBED ALL PRODUCTS
# ============================================================

@app.route("/api/embed/batch", methods=["POST"])
@login_required
def embed_batch():
    """Embedding байхгүй бүтээгдэхүүнүүдэд embedding үүсгэх"""
    rows = query("SELECT id, name FROM products WHERE embedding IS NULL LIMIT 100", fetchall=True)
    if not rows:
        return jsonify({"message": "Бүгд embed хийгдсэн", "processed": 0})

    texts = [r["name"] for r in rows]
    embeddings = embed_texts(texts)
    if not embeddings:
        return jsonify({"error": "OPENAI_API_KEY тохируулаагүй"}), 500

    for r, emb in zip(rows, embeddings):
        query("UPDATE products SET embedding = %s::vector WHERE id = %s", (str(emb), r["id"]), commit=True)

    remaining = query("SELECT COUNT(*) as cnt FROM products WHERE embedding IS NULL", fetchone=True)["cnt"]
    return jsonify({"processed": len(rows), "remaining": remaining})


# ============================================================
# SEARCH — text + image
# ============================================================

@app.route("/api/search", methods=["POST"])
@login_required
def search():
    """Text эсвэл зураг өгсний дараа тохирох бүтээгдэхүүн хайх"""
    query_text = ""
    image_b64 = None

    if request.is_json:
        data = request.get_json()
        query_text = data.get("query", "").strip()
    else:
        query_text = request.form.get("query", "").strip()
        if "image" in request.files:
            f = request.files["image"]
            img = Image.open(f.stream)
            img.thumbnail((1024, 1024))
            buf = io.BytesIO()
            img.convert("RGB").save(buf, format="JPEG", quality=85)
            image_b64 = base64.b64encode(buf.getvalue()).decode()

    # Зураг байвал Claude Vision-аар тайлбар гаргах
    if image_b64:
        client = get_claude()
        if not client:
            return jsonify({"error": "ANTHROPIC_API_KEY тохируулаагүй"}), 500
        msg = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=500,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg", "data": image_b64}},
                    {"type": "text", "text": f"""Зураг дээрх бараа, материалыг илрүүлж товч тайлбарла. Техникийн үзүүлэлтэй бол тэмдэглэ.
Жишээ: "ВВГ кабель 3×2.5мм²", "М30 цемент 50кг шуудай", "LED гэрэл 20W".
Нэмэлт асуулт: {query_text}
Зөвхөн товч тайлбар (1-2 өгүүлбэр)."""}
                ]
            }]
        )
        query_text = msg.content[0].text.strip() + (" " + query_text if query_text else "")

    if not query_text:
        return jsonify({"error": "Хайлтын текст эсвэл зураг оруулна уу"}), 400

    # Embedding хайлт
    embeddings = embed_texts([query_text])
    if not embeddings:
        # Embedding боломжгүй — fallback ILIKE
        products = query("""
            SELECT p.id, p.name, c.name as category_name, t.name as type_name,
                   (SELECT COUNT(*) FROM prices WHERE product_id = p.id) as price_count
            FROM products p
            LEFT JOIN product_categories c ON c.id = p.category_id
            LEFT JOIN product_types t ON t.id = p.type_id
            WHERE p.name ILIKE %s LIMIT 20
        """, (f"%{query_text}%",), fetchall=True)
        return jsonify({"query": query_text, "results": [serialize(r) for r in products]})

    # Vector similarity search
    products = query("""
        SELECT p.id, p.name, p.specs, p.unit,
               c.name as category_name, t.name as type_name,
               1 - (p.embedding <=> %s::vector) as similarity
        FROM products p
        LEFT JOIN product_categories c ON c.id = p.category_id
        LEFT JOIN product_types t ON t.id = p.type_id
        WHERE p.embedding IS NOT NULL
        ORDER BY p.embedding <=> %s::vector
        LIMIT 20
    """, (str(embeddings[0]), str(embeddings[0])), fetchall=True)

    # Үнийн мэдээлэл нэмэх
    results = []
    for p in products:
        prices = query("""
            SELECT source_type, source_name, unit_price, quantity, currency, price_date, notes
            FROM prices WHERE product_id = %s
            ORDER BY price_date DESC NULLS LAST, id DESC LIMIT 20
        """, (p["id"],), fetchall=True)

        item = serialize(p)
        item["prices"] = [serialize(x) for x in prices]
        item["price_summary"] = compute_price_summary(prices)
        results.append(item)

    # Хувилбарууд: similarity-оор бүлэглэх
    variants = {"exact": [], "similar": [], "related": []}
    for r in results:
        sim = r.get("similarity", 0)
        if sim >= 0.75:
            variants["exact"].append(r)
        elif sim >= 0.55:
            variants["similar"].append(r)
        else:
            variants["related"].append(r)

    return jsonify({
        "query": query_text,
        "results": results,
        "variants": variants
    })


def compute_price_summary(prices):
    """Үнийн мэдээлэл нэгтгэх"""
    if not prices:
        return None
    from decimal import Decimal
    supplier_prices = []
    own_prices = []
    for p in prices:
        up = p["unit_price"]
        if isinstance(up, Decimal):
            up = float(up)
        if not up or up <= 0:
            continue
        if p["source_type"] == "supplier":
            supplier_prices.append(up)
        else:
            own_prices.append(up)

    summary = {}
    if supplier_prices:
        summary["supplier"] = {
            "min": min(supplier_prices),
            "max": max(supplier_prices),
            "avg": sum(supplier_prices) / len(supplier_prices),
            "count": len(supplier_prices)
        }
    if own_prices:
        summary["own"] = {
            "min": min(own_prices),
            "max": max(own_prices),
            "avg": sum(own_prices) / len(own_prices),
            "count": len(own_prices)
        }
    return summary


# ============================================================
# STATS
# ============================================================

@app.route("/api/stats")
@login_required
def stats():
    cats = query("SELECT COUNT(*) as cnt FROM product_categories", fetchone=True)
    types = query("SELECT COUNT(*) as cnt FROM product_types", fetchone=True)
    prods = query("SELECT COUNT(*) as cnt FROM products", fetchone=True)
    prices = query("SELECT COUNT(*) as cnt, COUNT(CASE WHEN source_type='supplier' THEN 1 END) as sup, COUNT(CASE WHEN source_type='own' THEN 1 END) as own FROM prices", fetchone=True)
    files = query("SELECT COUNT(*) as cnt FROM price_files", fetchone=True)
    emb_ready = query("SELECT COUNT(*) as cnt FROM products WHERE embedding IS NOT NULL", fetchone=True)

    top_cats = query("""
        SELECT c.name, COUNT(DISTINCT p.id) as products, COUNT(pr.id) as prices
        FROM product_categories c
        LEFT JOIN products p ON p.category_id = c.id
        LEFT JOIN prices pr ON pr.product_id = p.id
        GROUP BY c.id, c.name
        ORDER BY prices DESC LIMIT 10
    """, fetchall=True)

    return jsonify({
        "categories": serialize(cats),
        "types": serialize(types),
        "products": serialize(prods),
        "prices": serialize(prices),
        "files": serialize(files),
        "embedded": serialize(emb_ready),
        "top_categories": [serialize(r) for r in top_cats]
    })


# ============================================================
# HTML
# ============================================================

@app.route("/")
def index():
    return send_file(str(BASE_DIR / "index.html"))


@app.route("/<page>")
def page(page):
    if page in ("upload", "search", "categories", "products"):
        return send_file(str(BASE_DIR / f"{page}.html"))
    return "Not found", 404


if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 5002))
    print(f"[price-db] Эхэллээ: http://localhost:{port}")
    app.run(host="0.0.0.0", port=port, debug=True)
