# مذكرتي Pro — v28

نظام ذكاء اصطناعي يُولّد عروض PowerPoint أكاديمية احترافية لمذكرات التخرج.

## المحركات
| Engine  | الأسلوب | مناسب لـ |
|---------|---------|---------|
| Canva   | عصري بدون شريط جانبي | Licence / Master |
| Premium | شريط جانبي فخم | Master / Doctorat |
| Classic | هيدر أكاديمي راقٍ | جميع المستويات |

## الثيمات (12 ثيم)
`navy_gold` · `dark_teal` · `burgundy` · `forest` · `midnight_purple`
`charcoal_orange` · `ice_blue` · `sand_gold` · `slate_crimson`
`noir` · `atlas` · `sakura`

## الإطلاق على Render

### 1. متغيرات البيئة المطلوبة
```
SECRET_KEY          = (تُولَّد تلقائياً)
ADMIN_PASSWORD_HASH = sha256 لكلمة مرور الأدمن
STORAGE_DIR         = /opt/render/project/src/storage
DB_PATH             = /opt/render/project/src/mathkarati_payments.db
```

### 2. توليد ADMIN_PASSWORD_HASH
```bash
echo -n "كلمةمرورك" | sha256sum
```

### 3. API Endpoints
| Method | Path | Description |
|--------|------|-------------|
| POST | `/generate` | توليد العرض |
| GET  | `/preview/<id>` | جلسة المعاينة |
| POST | `/orders` | إنشاء طلب دفع |
| POST | `/orders/<id>/receipt` | رفع وصل الدفع |
| POST | `/redeem` | استلام الملف بكود |
| GET  | `/health` | فحص الحالة |
| GET  | `/ping` | ping |

### 4. مثال على طلب `/generate`
```json
{
  "titleAr": "دور الذكاء الاصطناعي في التعليم",
  "studentName": "أحمد بن علي",
  "supervisor": "د. محمد السعيد",
  "institution": "جامعة الجزائر",
  "specialization": "علوم التربية",
  "year": "2024-2025",
  "engine": "canva",
  "theme": "navy_gold",
  "mainProblem": "...",
  "objectives": ["هدف 1", "هدف 2"],
  "stats": [{"label": "حجم العينة", "value": "200", "unit": "طالب"}],
  "mainResults": ["نتيجة 1", "نتيجة 2"],
  "generalConclusion": "..."
}
```

## الأداء
- توليد عرض كامل (14 شريحة): **~0.5 ثانية**
- حد أقصى للطلبات: **10 توليدات/دقيقة لكل IP**
- حجم الملف: **65-75 KB لكل عرض**

## بنية المشروع
```
app.py                  ← Flask API
build.sh                ← إعداد Render
render.yaml             ← تكوين Render
requirements.txt        ← المتطلبات
core/
  models.py             ← نماذج البيانات المُحصَّنة
  themes.py             ← 12 ثيم
  payment_models.py     ← قاعدة بيانات الطلبات
  preview.py            ← توليد المعاينة
engine/
  primitives.py         ← مكتبة الرسوميات (v28)
  pipeline.py           ← نقطة الدخول الوحيدة
  slides.py             ← Canva Engine (v28)
  slides_premium.py     ← Premium Engine (v28)
  slides_classic.py     ← Classic Engine (v28)
public/
  index.html            ← الواجهة الأمامية
  admin.html            ← لوحة الأدمن
```
