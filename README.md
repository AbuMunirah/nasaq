<p align="center">
  <img src="assets/icon-128.png" width="96" alt="Nasaq icon">
</p>

<h1 align="center">Nasaq | نَسَق</h1>

<p align="center">
  A PowerPoint add-in that adds a ribbon tab to align, match size and swap shapes in one click.<br>
  Works in PowerPoint for Windows and Mac, including Office LTSC, in English and Arabic.
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

### Which edition do I need?

Nasaq comes in two editions with the same tab and buttons:

| | **Nasaq** (web add-in) | **Nasaq Classic** (`.ppam`) |
|---|---|---|
| **For** | PowerPoint for Microsoft 365 (Windows, Mac, web) and Office LTSC 2024 | Office LTSC 2021, 2019 and 2016 on Windows |
| **Macros** | Not used: never asks to enable macros | VBA: the installer puts it in a trusted folder, so it normally loads without a prompt |
| **Internet** | Needed: loads from GitHub Pages | Not needed: works fully offline |
| **Updates** | Automatic | Run the installer again |
| **Arabic labels** | Windows and Mac | Windows (English on Mac) |

**Not sure which PowerPoint you have?** Go to **File → Account → About PowerPoint**.
If it says *Microsoft 365*, or *LTSC* with version 2408 or later, use **Nasaq**. If it says *LTSC* with version 2108, *2019* or *2016*, use **Nasaq Classic**.
(Nasaq needs PowerPoint JavaScript API 1.5, which older perpetual versions don't have.)

### Install

All installers are in the [latest release](https://github.com/AbuMunirah/nasaq/releases/latest) and need no admin rights.

#### Nasaq (web add-in)

**Mac:** download `Nasaq-Installer-Mac.zip`, unzip it and double-click `install-mac.command`.
If macOS says it can't verify the developer, open **System Settings → Privacy & Security** and click **Open Anyway**.
Or skip the download and paste this into Terminal:

```bash
curl -fsSL https://abumunirah.github.io/nasaq/install/install-mac.command | bash
```

**Windows:** download `Nasaq-Installer-Windows.cmd` and double-click it.
If Windows shows "Windows protected your PC", click **More info → Run anyway**.

Then restart PowerPoint and choose **Home → Add-ins → My Add-ins → Developer Add-ins → Nasaq**
(on Office LTSC 2024: **Insert → My Add-ins**).

**For an organization:** a Microsoft 365 admin can deploy it to everyone instead: upload
[`manifest.xml`](https://abumunirah.github.io/nasaq/manifest.xml) in the [Microsoft 365 admin center](https://admin.microsoft.com)
under **Settings → Integrated apps → Upload custom apps**. Nothing needs to be installed on users' machines.

#### Nasaq Classic (Windows, older PowerPoint)

Close PowerPoint, then download `Nasaq-Classic-Installer-Windows.cmd` and double-click it.
If Windows shows "Windows protected your PC", click **More info → Run anyway**.
Open PowerPoint: the Nasaq tab appears and loads automatically from then on.

The installer copies the add-in into PowerPoint's user add-ins folder (`%APPDATA%\Microsoft\AddIns`), which Office trusts by default,
so there is no macro prompt. It doesn't change any security settings. If your organization has disabled trusted locations or VBA,
the installer says so; ask your IT team to allow the add-in.

**Manual install:** download `Nasaq-Classic.ppam`, right-click it → **Properties** → tick **Unblock** → **OK**.
In PowerPoint, go to **File → Options → Add-ins**, set **Manage** to **PowerPoint Add-ins**, click **Go… → Add New…**,
select the file and choose **Enable Macros**.

#### Uninstall

Run the same installer with `--uninstall` (Mac) or `/uninstall` (Windows), for example `install-windows-classic.cmd /uninstall`.

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
| `ppam/` | Nasaq Classic: VBA template, ribbon (`customUI14.xml`) and build script |
| `install/` | Installers for Mac and Windows |

Every push to `main` builds the add-in and publishes it to GitHub Pages
(see [`.github/workflows/pages.yml`](.github/workflows/pages.yml)).

**Nasaq Classic** lives in [`ppam/`](ppam). Its labels are generated from `manifest.xml`, so both editions stay in sync:

1. `python3 ppam/build.py vba` writes `ppam/out/Nasaq.bas`.
2. In PowerPoint, import `Nasaq.bas` into a blank presentation's VBA project (**Debug → Compile** to check it),
   then save it as a macro-enabled presentation (`.pptm`).
3. `python3 ppam/build.py ppam path/to/saved.pptm` writes `ppam/out/Nasaq.ppam` with the ribbon and icons added
   (and the author name cleared). Copy it to `ppam/Nasaq.ppam` to publish it.

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

### أي إصدار أحتاج؟

يتوفر نَسَق بإصدارين بنفس علامة التبويب والأزرار:

| | **نَسَق** (إضافة ويب) | **نَسَق الكلاسيكي** (`.ppam`) |
|---|---|---|
| **مناسب لـ** | PowerPoint ضمن Microsoft 365 (Windows وMac والويب) وOffice LTSC 2024 | Office LTSC 2021 و2019 و2016 على Windows |
| **وحدات الماكرو** | لا يستخدمها: لا يطلب تمكين وحدات الماكرو أبدًا | VBA: يضعه المثبّت في مجلد موثوق، فيُحمَّل عادةً دون أي تنبيه |
| **الإنترنت** | مطلوب: يُحمَّل من GitHub Pages | غير مطلوب: يعمل دون اتصال بالكامل |
| **التحديثات** | تلقائية | شغّل المثبّت مرة أخرى |
| **الواجهة العربية** | Windows وMac | Windows (بالإنجليزية على Mac) |

**لست متأكدًا من إصدار PowerPoint لديك؟** انتقل إلى **ملف ← الحساب ← حول PowerPoint**.
إذا ظهر *Microsoft 365*، أو *LTSC* بالإصدار 2408 أو أحدث، فاستخدم **نَسَق**. وإذا ظهر *LTSC* بالإصدار 2108، أو *2019* أو *2016*، فاستخدم **نَسَق الكلاسيكي**.

### التثبيت

جميع المثبّتات موجودة في [أحدث إصدار](https://github.com/AbuMunirah/nasaq/releases/latest)، ولا تحتاج إلى صلاحيات المسؤول.

#### نَسَق (إضافة ويب)

**Mac:** نزّل `Nasaq-Installer-Mac.zip`، وفك ضغطه، ثم انقر نقرًا مزدوجًا على `install-mac.command`.
إذا ظهرت رسالة بأن macOS لا يمكنه التحقق من المطوّر، افتح **إعدادات النظام ← الخصوصية والأمان** وانقر **فتح على أي حال**.
أو الصق هذا الأمر في تطبيق Terminal بدلًا من التنزيل:

```bash
curl -fsSL https://abumunirah.github.io/nasaq/install/install-mac.command | bash
```

**Windows:** نزّل `Nasaq-Installer-Windows.cmd` وانقر عليه نقرًا مزدوجًا.
إذا ظهرت رسالة "Windows protected your PC"، انقر **More info ← Run anyway**.

بعدها أعد تشغيل PowerPoint واختر **الصفحة الرئيسية ← الوظائف الإضافية ← وظائفي الإضافية ← وظائف المطوّر الإضافية ← نَسَق**
(في Office LTSC 2024: **إدراج ← وظائفي الإضافية**).

**للمؤسسات:** يمكن لمسؤول Microsoft 365 نشره للجميع بدلًا من ذلك، برفع ملف
[`manifest.xml`](https://abumunirah.github.io/nasaq/manifest.xml) في [مركز إدارة Microsoft 365](https://admin.microsoft.com)
عبر **الإعدادات ← التطبيقات المتكاملة ← رفع تطبيقات مخصصة**، دون تثبيت أي شيء على أجهزة المستخدمين.

#### نَسَق الكلاسيكي (Windows، إصدارات PowerPoint الأقدم)

أغلق PowerPoint، ثم نزّل `Nasaq-Classic-Installer-Windows.cmd` وانقر عليه نقرًا مزدوجًا.
إذا ظهرت رسالة "Windows protected your PC"، انقر **More info ← Run anyway**.
افتح PowerPoint: ستظهر علامة التبويب نَسَق، وستُحمَّل تلقائيًا في كل مرة بعد ذلك.

ينسخ المثبّت الإضافة إلى مجلد الوظائف الإضافية للمستخدم (`%APPDATA%\Microsoft\AddIns`)، وهو مجلد يثق به Office افتراضيًا،
لذلك لا يظهر تنبيه وحدات الماكرو. لا يغيّر المثبّت أي إعدادات أمان. إذا كانت مؤسستك قد عطّلت المواقع الموثوقة أو VBA،
فسيُعلمك المثبّت بذلك؛ اطلب حينها من فريق تقنية المعلومات السماح بالإضافة.

**التثبيت اليدوي:** نزّل `Nasaq-Classic.ppam`، وانقر عليه بزر الماوس الأيمن ← **خصائص** ← فعّل **إلغاء الحظر** ← **موافق**.
في PowerPoint، انتقل إلى **ملف ← خيارات ← الوظائف الإضافية**، واختر **PowerPoint Add-ins** من قائمة **إدارة**، وانقر **انتقال… ← إضافة جديد…**،
ثم حدد الملف واختر **تمكين وحدات الماكرو**.

#### إلغاء التثبيت

شغّل المثبّت نفسه مع `--uninstall` (Mac) أو `/uninstall` (Windows)، مثلًا `install-windows-classic.cmd /uninstall`.

### الخصوصية

لا تجمع نَسَق أي بيانات ولا تخزنها ولا ترسلها. تعمل بالكامل داخل PowerPoint، ولا تغيّر إلا الأشكال والنصوص التي تحددها. تستضيف GitHub Pages ملفات الإضافة فقط، ولا تصلها عروضك التقديمية أبدًا.

### الترخيص

[GPL-3.0](LICENSE)

</div>
