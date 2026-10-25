<p align="center">
  <img src="assets/icon-128.png" width="96" alt="Nasaq icon">
</p>

<h1 align="center">Nasaq | نَسَق</h1>

<p align="center">
  A PowerPoint add-in that adds a ribbon tab to align, match size and swap shapes in one click.<br>
  Works in PowerPoint for Windows and Mac, in English and Arabic.
</p>

<p align="center">
  <a href="#english">English</a> · <a href="#العربية">العربية</a>
</p>

---

## English

### Features

Every Nasaq command uses the **last shape you select** as the reference, so you decide which shape the others follow.

| Group | Commands | What they do |
|---|---|---|
| **Font** | Font, Font Size | PowerPoint's own font controls, kept at hand |
| **Paragraph** | Text Align ▾, Font Color | Align selected text, or all text in the selected shapes: left, center, right or justify |
| **Shapes** | Insert Shape, Shape Fill | PowerPoint's own shape controls |
| **Align** | Left, Center, Right, Top, Middle, Bottom | Line up the selected shapes with the edge or center of the last selected shape |
| | Distribute Horizontally / Vertically | PowerPoint's own spacing controls |
| **Size** | Match Size, Match Width, Match Height | Resize the selected shapes to match the last selected shape |
| **Position** | Swap | Swap two shapes, center to center |

The ribbon labels, tooltips and messages appear in Arabic when Office is set to Arabic, and in English otherwise.

### Requirements

- PowerPoint for Microsoft 365 on Windows or Mac (the add-in needs PowerPoint API 1.5).
- An internet connection: like all Office web add-ins, Nasaq loads its code from GitHub Pages when you use it.

### Install

The add-in is defined by one file, the manifest:
**https://abumunirah.github.io/nasaq/manifest.xml**

**For an organization (recommended).** A Microsoft 365 admin uploads the manifest in the
[Microsoft 365 admin center](https://admin.microsoft.com) under **Settings → Integrated apps → Upload custom apps**.
The Nasaq tab then appears for the chosen users in every presentation, with nothing to install on their machines.

**For yourself, on Mac**
1. Download `manifest.xml` from the link above.
2. Copy it to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (create the `wef` folder if it doesn't exist).
3. Restart PowerPoint, then choose **Home → Add-ins → My Add-ins → Developer Add-ins → Nasaq**.

**For yourself, on Windows**
1. Put `manifest.xml` in a shared network folder.
2. In PowerPoint, go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**, add the folder's network path and tick **Show in Menu**.
3. Restart PowerPoint, then choose **Home → Add-ins → My Add-ins → Shared Folder → Nasaq**.

### Privacy

Nasaq does not collect, store or send any data. It runs entirely inside PowerPoint and only changes the shapes and text you select.
GitHub Pages serves the add-in's files; it never receives your presentations.

### Development

Built with [Office.js](https://learn.microsoft.com/office/dev/add-ins/) and TypeScript, starting from the Yeoman Office generator.

```bash
npm install
npm start        # starts the dev server on https://localhost:3000 and opens PowerPoint with the add-in
npm stop         # stops the dev server and removes the development add-in
npm run build    # production build in dist/, with URLs pointing to GitHub Pages
npm run validate # checks manifest.xml
```

| Path | Contents |
|---|---|
| `manifest.xml` | Ribbon tab, buttons, icons and English/Arabic labels |
| `src/commands/commands.powerpoint.ts` | The ribbon commands |
| `src/commands/geometry.ts` | Alignment and swap calculations |
| `src/commands/strings.ts` | English and Arabic messages |
| `src/message/` | Dialog used to show messages |
| `assets/` | Icons |

Every push to `main` builds the add-in and publishes it to GitHub Pages
(see [`.github/workflows/pages.yml`](.github/workflows/pages.yml)).

### License

[GPL-3.0](LICENSE)

---

<div dir="rtl" lang="ar">

## العربية

**نَسَق** إضافة لبرنامج PowerPoint تضيف علامة تبويب على الشريط لمحاذاة الأشكال ومطابقة أحجامها وتبديل مواضعها بنقرة واحدة. تعمل على PowerPoint في Windows وMac، باللغتين العربية والإنجليزية.

### المزايا

تستخدم جميع أوامر نَسَق **آخر شكل تحدده** مرجعًا، فأنت من يقرر الشكل الذي تتبعه بقية الأشكال.

| المجموعة | الأوامر | الوظيفة |
|---|---|---|
| **الخط** | الخط، حجم الخط | أدوات الخط الأصلية في PowerPoint في متناول يدك |
| **الفقرة** | محاذاة النص ▾، لون الخط | محاذاة النص المحدد، أو كل النص في الأشكال المحددة: يسار، توسيط، يمين، أو ضبط |
| **الأشكال** | إدراج شكل، تعبئة الشكل | أدوات الأشكال الأصلية في PowerPoint |
| **المحاذاة** | يسار، توسيط أفقي، يمين، أعلى، توسيط عمودي، أسفل | محاذاة الأشكال المحددة مع حافة آخر شكل محدد أو مركزه |
| | التوزيع أفقيًا / عموديًا | أدوات التباعد الأصلية في PowerPoint |
| **الحجم** | مطابقة الحجم، مطابقة العرض، مطابقة الارتفاع | تغيير حجم الأشكال المحددة لتطابق آخر شكل محدد |
| **الموضع** | تبديل | تبديل موضعي شكلين، من المركز إلى المركز |

تظهر تسميات الشريط والتلميحات والرسائل باللغة العربية عندما تكون لغة Office هي العربية، وبالإنجليزية في غير ذلك.

### المتطلبات

- PowerPoint ضمن Microsoft 365 على Windows أو Mac.
- اتصال بالإنترنت: مثل جميع إضافات Office المبنية على الويب، تحمّل نَسَق شيفرتها من GitHub Pages عند استخدامها.

### التثبيت

تُعرَّف الإضافة بملف واحد هو ملف البيان (manifest):
**https://abumunirah.github.io/nasaq/manifest.xml**

**للمؤسسات (موصى به).** يرفع مسؤول Microsoft 365 ملف البيان من [مركز إدارة Microsoft 365](https://admin.microsoft.com) عبر **الإعدادات ← التطبيقات المتكاملة ← رفع تطبيقات مخصصة**. بعدها تظهر علامة التبويب نَسَق للمستخدمين المختارين في كل عرض تقديمي، دون تثبيت أي شيء على أجهزتهم.

**للاستخدام الشخصي على Mac**
1. نزّل ملف `manifest.xml` من الرابط أعلاه.
2. انسخه إلى المجلد `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (أنشئ المجلد `wef` إن لم يكن موجودًا).
3. أعد تشغيل PowerPoint، ثم اختر **الصفحة الرئيسية ← الوظائف الإضافية ← وظائفي الإضافية ← وظائف المطوّر الإضافية ← نَسَق**.

**للاستخدام الشخصي على Windows**
1. ضع ملف `manifest.xml` في مجلد مشترك على الشبكة.
2. في PowerPoint، انتقل إلى **ملف ← خيارات ← مركز التوثيق ← إعدادات مركز التوثيق ← كتالوجات الوظائف الإضافية الموثوقة**، وأضف مسار المجلد على الشبكة وفعّل **إظهار في القائمة**.
3. أعد تشغيل PowerPoint، ثم اختر **الصفحة الرئيسية ← الوظائف الإضافية ← وظائفي الإضافية ← مجلد مشترك ← نَسَق**.

### الخصوصية

لا تجمع نَسَق أي بيانات ولا تخزنها ولا ترسلها. تعمل بالكامل داخل PowerPoint، ولا تغيّر إلا الأشكال والنصوص التي تحددها. تستضيف GitHub Pages ملفات الإضافة فقط، ولا تصلها عروضك التقديمية أبدًا.

### الترخيص

[GPL-3.0](LICENSE)

</div>
