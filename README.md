# خطة المتابعة الشهرية — نسخة Online

1. أنشئ جدول Google Sheets جديدًا، وأضف ورقة باسم Users بالأعمدة: code | nid | name | tawgeeh | role (الرقم القومي والكود كنص). أبقِ الجدول خاصًا.
2. من الجدول: Extensions ▸ Apps Script، والصق Code.gs، ثم Deploy ▸ New deployment ▸ Web app (Execute as: Me — Who has access: Anyone) وانسخ الرابط.
3. في public/index.html استبدل PASTE_APPS_SCRIPT_WEB_APP_URL_HERE بالرابط (CFG.API).
4. النشر: npm i -g firebase-tools ▸ firebase login ▸ firebase init hosting (اختر public، بدون SPA) ▸ firebase deploy

5. التصدير إلى Drive: شغّل الدالة setupDailyExport مرة واحدة من محرر Apps Script (وافق على صلاحيات Drive) لتصدير كل الموجهين يوميًا الساعة 2 صباحًا، أو شغّل exportAll يدويًا. تُحفظ الملفات في مجلد «خطط المتابعة - تصدير» وفي مجلد فرعي لكل موجه.

6. ملف التعديل المرفوع يُحفظ في Drive وFirebase أيضًا:
   - Drive: ضع معرّف مجلد المنظمة في ORG_FOLDER_ID (من رابط المجلد) أو اتركه فارغًا ليُنشأ مجلد «منظومة المتابعة الإلكترونية - التعديلات» (مجلد فرعي لكل موجه).
   - Firebase: أنشئ Firestore (Native) وضع معرّف المشروع في FIREBASE_PROJECT_ID. في Apps Script: Project Settings ▸ Google Cloud Platform project ▸ Change project، وأدخل رقم مشروع Firebase. انسخ appsscript.json (Show manifest file) ثم نفّذ firebase deploy --only firestore:rules ليُغلق الوصول من المتصفح.
   - بعد أي تعديل أعد النشر: Deploy ▸ Manage deployments ▸ New version.

7. زر «إرسال الخطة الشهرية» يحفظ ملف XLSX في Drive (مجلد الخطط)، ويضيف سجلًا في Firestore (مجموعة plans)، ويكتب الصفوف في ورقة Plans وسجل الملف في ورقة PlanFiles. زر «تنزيل الخطة XLSX على جهازي» يحفظ نسخة محلية بجوار زر «حفظ الخطة الشهرية».
