# ✅ قائمة تدقيق الإطلاق — مذكرتي Pro v28

## قبل الإطلاق

### الإعدادات الأمنية
- [ ] تغيير `ADMIN_PASSWORD_HASH` في Render env vars
- [ ] التحقق أن `SECRET_KEY` مُولَّد تلقائياً من Render
- [ ] اختبار تسجيل دخول الأدمن بكلمة المرور الجديدة

### اختبار البنية
- [ ] `GET /ping` → يرجع `pong`
- [ ] `GET /health` → يرجع `{"status":"ok","font":"Cairo"}`
  - ⚠️ إذا `"font":"Calibri"` فالخط لم يُنصَّب — راجع build logs
- [ ] `POST /generate` بملف اختبار → يرجع `preview_slides`
- [ ] `POST /orders` → يُنشئ طلب
- [ ] رفع وصل على الطلب
- [ ] لوحة الأدمن `/admin` تعمل
- [ ] اعتماد طلب → توليد كود تحميل
- [ ] `POST /redeem` بالكود → يرجع ملف PPTX

### الأداء
- [ ] أول طلب (cold start Render): < 30 ثانية
- [ ] الطلبات اللاحقة (warm): < 2 ثانية
- [ ] `GET /warmup` يُسرّع cold start — اضبطه في frontend

### الخط العربي
للتحقق بعد النشر:
```
GET /health
```
يجب أن يظهر `"font": "Cairo"` وليس `"font": "Calibri"`.
إذا كان Calibri: العروض تعمل لكن بخط أقل جمالاً على بعض أنظمة التشغيل.

## بعد الإطلاق

### مراقبة أولى 24 ساعة
- [ ] راقب Render logs: لا أخطاء 500
- [ ] راقب أحجام ملفات storage/
- [ ] اختبار flow كامل: توليد → معاينة → دفع → استلام

### حدود Render Free Tier
- RAM: 512MB (المشروع يستخدم ~150MB عادةً)
- وقت Spindown: 15 دقيقة خمول → cold start عند أول طلب
- Bandwidth: 100GB/شهر
- Storage: مؤقت — يُحذف عند إعادة النشر ⚠️

### ⚠️ تحذير مهم: التخزين على Render Free
ملفات storage/ (PPTX + صور الوصل + قاعدة البيانات) **تُحذف** عند كل نشر جديد.
**الحل للإنتاج الحقيقي**: استخدم Render Disk أو Supabase Storage أو AWS S3.

## معلومات الدفع (تُضبط من admin.html)
- طريقة الدفع: CCP أو بريدي موب
- السعر: 800 دج (قابل للتغيير من payment_models.py)
- مدة صلاحية كود التحميل: 48 ساعة (افتراضي)

## جهات الاتصال في حالة المشاكل
- Render Status: https://status.render.com
- python-pptx Issues: https://github.com/python-openxml/python-pptx

