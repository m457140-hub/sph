# استضافة أيقونة تطبيق SPH على GitHub Pages

## 📁 الملفات المرفقة
- `icon-72.png` إلى `icon-512.png` (8 مقاسات مختلفة)

## 🚀 خطوات الإعداد

### 1) إنشاء مستودع (Repository)
1. اذهب إلى github.com وسجّل دخول (أو أنشئ حساباً مجانياً)
2. اضغط **"+"** ثم **"New repository"**
3. اسم المستودع مثلاً: `sph-attendance-icon`
4. اختر **Public**
5. اضغط **"Create repository"**

### 2) رفع ملفات الأيقونات
1. داخل المستودع اضغط **"Add file"** → **"Upload files"**
2. اسحب كل ملفات `icon-72.png` ... `icon-512.png` معاً
3. اضغط **"Commit changes"**

### 3) تفعيل GitHub Pages
1. من تبويب **"Settings"** اذهب لـ **"Pages"** في القائمة الجانبية
2. تحت **"Source"** اختر **"Deploy from a branch"**
3. اختر الفرع `main` والمجلد `/ (root)`
4. اضغط **"Save"** وانتظر دقيقتين تقريباً

بعدها سيكون رابطك بهذا الشكل:
```
https://YOUR_USERNAME.github.io/sph-attendance-icon/
```

### 4) تحديث كود Google Apps Script
في أول ملف الكود (`SPH_Attendance_GitHub_Pages.gs`) عدّل السطر التالي فقط:

```javascript
var GITHUB_ICON_BASE = "https://YOUR_GITHUB_USERNAME.github.io/YOUR_REPO_NAME";
```

استبدل:
- `YOUR_GITHUB_USERNAME` باسم حسابك على GitHub
- `YOUR_REPO_NAME` باسم المستودع (مثلاً `sph-attendance-icon`)

لا حاجة لتعديل أي شيء آخر في الكود — كل الأيقونات والمانيفست تُبنى تلقائياً من هذا الرابط.

### 5) الحفظ وإعادة النشر
1. احفظ الكود في Apps Script
2. Deploy → Manage deployments → عدّل النشر الحالي (New version)

### 6) الاختبار
1. تحقق أولاً أن الرابط يعمل مباشرة في المتصفح:
   `https://YOUR_USERNAME.github.io/sph-attendance-icon/icon-512.png`
   (يجب أن تظهر الصورة مباشرة، وليس صفحة 404)
2. افتح رابط التطبيق على الهاتف
3. امسح ذاكرة التخزين المؤقت للمتصفح (Clear cache)
4. أضف التطبيق للشاشة الرئيسية (Add to Home Screen)
5. يجب أن تظهر أيقونة SPH الصحيحة بدلاً من أيقونة Google الافتراضية

## ⚠️ إذا لم تظهر الأيقونة بعد
- تأكد أن GitHub Pages مفعّل فعلاً (علامة ✅ خضراء في Settings → Pages)
- جرّب فتح رابط الأيقونة مباشرة في المتصفح للتأكد أنه يعمل
- امسح الكاش وجرّب في وضع التصفح الخفي (Incognito)
- انتظر 5-10 دقائق إضافية لانتشار الملفات عبر CDN الخاص بـ GitHub
