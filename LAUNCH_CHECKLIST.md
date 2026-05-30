# 🚀 مذكرتي Pro — قائمة التحقق قبل الإطلاق التجاري

## ✅ الخطوات المطلوبة على Render Dashboard

### 1. متغيرات البيئة الإلزامية

| المتغير | القيمة | الوصف |
|---------|--------|-------|
| `ADMIN_PASSWORD_HASH` | sha256 hex | كلمة سر لوحة الإدارة |
| `SECRET_KEY` | auto-generated | مفتاح Flask السري |
| `CCP_NUM` | رقم حسابك | رقم حساب CCP الخاص بك |
| `CCP_KEY` | المفتاح | مفتاح CCP |
| `BARID_NUM` | رقم BaridiMob | رقمك على BaridiMob |
| `PAYMENT_OWNER` | اسمك | الاسم في وصل الدفع |

### 2. متغيرات اختيارية

| المتغير | الوصف |
|---------|-------|
| `PRICE` | السعر (default: 800 دج) |
| `TELEGRAM_BOT_TOKEN` | إشعارات فورية بالطلبات |
| `TELEGRAM_CHAT_ID` | معرف محادثة Telegram |

---

## 🔐 كيف تحسب ADMIN_PASSWORD_HASH

```bash
echo -n "كلمة_السر_هنا" | sha256sum
```

أو في Python:
```python
import hashlib
print(hashlib.sha256("كلمة_السر_هنا".encode()).hexdigest())
```

---

## 📱 إعداد Telegram Bot (اختياري لكن موصى به)

1. افتح @BotFather في Telegram
2. أرسل `/newbot` واتبع التعليمات
3. احفظ الـ Token
4. افتح المحادثة مع البوت
5. اذهب إلى: `https://api.telegram.org/bot{TOKEN}/getUpdates`
6. احفظ `chat_id`

---

## 🌐 بعد الرفع على Render

1. تأكد أن `/health` يُرجع `{"status":"ok"}`
2. جرّب توليد عرض تجريبي
3. تأكد أن `/config` يُرجع بيانات الدفع الصحيحة
4. اختبر لوحة الإدارة: `your-url/admin`
5. جرّب رفع وصل تجريبي وقبوله

---

## ⚠️ ملاحظات هامة

- **Render Free Plan**: الـ storage يُمسح عند كل deploy — استخدم **Starter Plan** مع disk mount
- الـ disk mount مُعيَّن في `render.yaml` على `/opt/render/project/src/storage`
- النسخة الاحتياطية من DB: حمّلها بانتظام من `/admin`
- وقت الاستجابة الأول قد يكون بطيئاً (cold start) — الـ `/warmup` endpoint يساعد

---

## 📊 مراقبة الأداء

- `/health` — حالة كاملة مع إحصائيات
- `/admin` — لوحة الإدارة الكاملة
- Render Logs — في Dashboard

---

## 🔄 Rate Limits المُطبَّقة

| Endpoint | الحد | النافذة |
|----------|------|---------|
| `/generate` | 5 طلبات | 60 ثانية |
| `/orders` | 3 طلبات | 10 دقائق |
| `/redeem` | 10 محاولات | ساعة |

