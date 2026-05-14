# Price DB — Үнийн санал, бүтээгдэхүүний сан

Excel файлаас AI шинжилж систематик ангилал үүсгэн үнийн харьцуулалт хийдэг систем.

## Суулгах

```bash
cd ~/Downloads/price-db
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
# .env дотор ANTHROPIC_API_KEY, OPENAI_API_KEY нэмэх
```

## Ажиллуулах

```bash
source venv/bin/activate
export $(cat .env | xargs)
python app.py
```

Browser: `http://localhost:5002`

Нэвтрэх: tender-dashboard-тай ижил (admin / admin123).

## Ашиглах

**1. Excel оруулах** (`/upload`)
- 🏢 Нийлүүлэгчийн үнэ эсвэл 📝 Манай өгсөн үнэ сонгоно
- Эх үүсвэр нэр бичнэ (компани эсвэл тендерийн нэр)
- .xlsx файл оруулна
- AI автомат ангилал үүсгэж бүтээгдэхүүн бүрийг классификласан байдлаар DB-д оруулна

**2. Embedding үүсгэх** (`/products`)
- Хайлт боломжтой болгохын тулд 🧠 Embed товч дарна
- OpenAI text-embedding-3-small ашиглана (шаарадлагатай ~$0.02/10K)

**3. Хайлт** (`/search`)
- Текст (жишээ: "25мм² зэс кабель")
- Зураг (Claude Vision шинжилнэ)
- Хувилбаруудаар гарна:
  - ✓ Яг тохирох (75%+)
  - ~ Ойролцоо (55-75%)
  - … Холбогдох (<55%)

## DB Schema

- `product_categories` — ангилал (Кабель, Бетон...)
- `product_types` — дэд төрөл (ВВГ, М30...)
- `products` — бүтээгдэхүүн бүр (ангилалтай, embedding-тэй)
- `prices` — үнэ бүр (supplier/own эх үүсвэрээр)
- `price_files` — upload түүх

## Порт

- tender-dashboard: 5000
- company-resources: 5001
- **price-db: 5002**
