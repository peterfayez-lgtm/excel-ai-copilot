/**
 * Taskpane Controller - Excel AI Copilot
 * مسؤول عن واجهة المستخدم وتنسيق التواصل بين الإكسيل والذكاء الاصطناعي
 */

document.addEventListener("DOMContentLoaded", () => {
  // عناصر الواجهة
  const chatContainer = document.getElementById("chatContainer");
  const userInput = document.getElementById("userInput");
  const btnSend = document.getElementById("btnSend");
  const btnClearChat = document.getElementById("btnClearChat");
  const btnSettings = document.getElementById("btnSettings");
  const settingsModal = document.getElementById("settingsModal");
  const btnCloseSettings = document.getElementById("btnCloseSettings");
  const btnCancelSettings = document.getElementById("btnCancelSettings");
  const btnSaveSettings = document.getElementById("btnSaveSettings");
  const apiKeyInput = document.getElementById("apiKeyInput");
  const modelSelect = document.getElementById("modelSelect");
  const customInstructions = document.getElementById("customInstructions");
  const appLogo = document.getElementById("appLogo");
  const appTitle = document.getElementById("appTitle");
  const btnExportWord = document.getElementById("btnExportWord");
  const personaSelect = document.getElementById("personaSelect");
  const personaBadge = document.getElementById("personaBadge");
  const currentSheetName = document.getElementById("currentSheetName");
  const currentSelection = document.getElementById("currentSelection");
  const btnRefreshContext = document.getElementById("btnRefreshContext");
  const chkIncludeContext = document.getElementById("chkIncludeContext");
  const chkNewSheetMode = document.getElementById("chkNewSheetMode");
  const statusIndicator = document.getElementById("statusIndicator");
  const statusText = document.getElementById("statusText");
  const quickActionSelect = document.getElementById("quickActionSelect");
  const quickActionDropdown = document.getElementById("quickActionDropdown");
  const quickActionBtn = document.getElementById("quickActionBtn");
  const quickActionMenu = document.getElementById("quickActionMenu");
  const quickActionLabel = document.getElementById("quickActionLabel");
  const quickActionIcon = document.getElementById("quickActionIcon");
  const actionHostBadge = document.getElementById("actionHostBadge");
  const btnDetectModels = document.getElementById("btnDetectModels");
  const modelDetectStatus = document.getElementById("modelDetectStatus");
  const studioPrimaryGrid = document.getElementById("studioPrimaryGrid");
  const studioSecondaryGrid = document.getElementById("studioSecondaryGrid");
  const studioDrawer = document.getElementById("studioDrawer");
  const btnToggleAllActions = document.getElementById("btnToggleAllActions");
  const toggleAllActionsText = document.getElementById("toggleAllActionsText");
  const toggleAllActionsIcon = document.getElementById("toggleAllActionsIcon");

  // عناصر إرفاق ومعاينة المستندات (صور و PDF)
  const btnAttachImage = document.getElementById("btnAttachImage");
  const imageFileInput = document.getElementById("imageFileInput");
  const imagePreviewContainer = document.getElementById("imagePreviewContainer");
  const imagePreviewImg = document.getElementById("imagePreviewImg");
  const pdfPreviewBadge = document.getElementById("pdfPreviewBadge");
  const imagePreviewName = document.getElementById("imagePreviewName");
  const btnRemoveImage = document.getElementById("btnRemoveImage");

  // عناصر مكتبة القوالب
  const btnTemplates = document.getElementById("btnTemplates");
  const templatesModal = document.getElementById("templatesModal");
  const btnCloseTemplates = document.getElementById("btnCloseTemplates");
  const newTemplateTitle = document.getElementById("newTemplateTitle");
  const newTemplatePrompt = document.getElementById("newTemplatePrompt");
  const btnSaveNewTemplate = document.getElementById("btnSaveNewTemplate");
  const customTemplatesList = document.getElementById("customTemplatesList");

  let attachedFile = null; // { data: string (base64), mimeType: string, dataUrl: string, name: string, isPdf: boolean }
  let chatHistory = []; // ذاكرة المحادثة التراكمية (Multi-Turn Chat History)

  // نصوص توضيحية للشخصيات (Excel & Word)
  const personaDescriptions = {
    cfo: "يركز على: ضبط الحسابات، القوائم المالية، تنسيق العملات، وفحص الميزانيات بدقة محاسبية.",
    analyst: "يركز على: استخراج المؤشرات، الرسوم البيانية الإحصائية، وتحليل الاتجاهات والأنماط.",
    operations: "يركز على: تتبع الكميات، المهل الزمنية، مستويات المخزون، والتنبيهات اللونية.",
    executive: "يركز على: السرعة والمرونة، التنسيق الشامل، وصياغة ملخصات الأعمال الأنيقة.",
    writer: "يركز على: كتابة التقارير الإدارية، خطط الأعمال، والمذكرات التنفيذية للقيادات.",
    legal: "يركز على: صياغة العقود المحكمة، الشروط والالتزامات، والبنود القانونية الرصينة.",
    editor: "يركز على: السلامة اللغوية، تدقيق النحو وعلامات الترقيم، والترجمة الاحترافية للأعمال."
  };

  // وظائف إدارة إرفاق ومعاينة الملفات (صور أو مستندات PDF)
  function showFilePreview(file, dataUrl) {
    const base64Data = dataUrl.split(",")[1];
    const isPdf = file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");
    const mimeType = isPdf ? "application/pdf" : (file.type || "image/png");
    const fileName = file.name || (isPdf ? "مستند.pdf" : "صورة_ملصقة.png");
    
    attachedFile = {
      data: base64Data,
      mimeType: mimeType,
      dataUrl: dataUrl,
      name: fileName,
      isPdf: isPdf
    };

    if (isPdf) {
      if (imagePreviewImg) imagePreviewImg.style.display = "none";
      if (pdfPreviewBadge) pdfPreviewBadge.style.display = "flex";
    } else {
      if (pdfPreviewBadge) pdfPreviewBadge.style.display = "none";
      if (imagePreviewImg) {
        imagePreviewImg.style.display = "block";
        imagePreviewImg.src = dataUrl;
      }
    }

    imagePreviewName.textContent = fileName.length > 25 ? fileName.substring(0, 22) + "..." : fileName;
    imagePreviewContainer.style.display = "flex";
  }

  function clearAttachedFile() {
    attachedFile = null;
    if (imagePreviewImg) {
      imagePreviewImg.src = "";
      imagePreviewImg.style.display = "none";
    }
    if (pdfPreviewBadge) {
      pdfPreviewBadge.style.display = "none";
    }
    imagePreviewContainer.style.display = "none";
    if (imageFileInput) imageFileInput.value = "";
  }

  if (btnRemoveImage) {
    btnRemoveImage.addEventListener("click", clearAttachedFile);
  }

  if (btnAttachImage && imageFileInput) {
    btnAttachImage.addEventListener("click", () => {
      imageFileInput.click();
    });

    imageFileInput.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (!file) return;

      const isImage = file.type.startsWith("image/");
      const isPdf = file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf");

      if (!isImage && !isPdf) {
        alert("يرجى اختيار ملف صورة صالح (PNG, JPG, إلخ) أو مستند PDF.");
        return;
      }

      const reader = new FileReader();
      reader.onload = (evt) => {
        showFilePreview(file, evt.target.result);
      };
      reader.readAsDataURL(file);
    });
  }

  // الاستماع للنسخ واللصق المباشر من الحافظة (Ctrl + V)
  window.addEventListener("paste", (e) => {
    const clipboardData = e.clipboardData || window.clipboardData;
    if (!clipboardData) return;

    const items = clipboardData.items;
    if (!items) return;

    for (let i = 0; i < items.length; i++) {
      if (items[i].type && items[i].type.startsWith("image/")) {
        const file = items[i].getAsFile();
        if (file) {
          const reader = new FileReader();
          reader.onload = (evt) => {
            showFilePreview(file, evt.target.result);
          };
          reader.readAsDataURL(file);
          break;
        }
      }
    }
  });

  let currentHost = "excel"; // البيئة الحالية الافتراضية ("excel" أو "word")

  // مصفوفات الأوامر والإجراءات السريعة المباشرة (Excel و Word)
  const EXCEL_QUICK_ACTIONS = [
    {
      icon: "🔍",
      title: "تدقيق وتطهير الشيت",
      desc: "إصلاح أخطاء الصيغ (#DIV/0!, #N/A) وتظليل القيم المكررة وإزالة الفراغات",
      prompt: "🔍 قم بتدقيق وتطهير البيانات في الشيت: فحص أخطاء المعادلات (#DIV/0!, #N/A, #VALUE!) وتصحيحها بـ IFERROR، وتظليل القيم المكررة، وتنظيف المسافات الزائدة وإصلاح تنسيقات الأرقام."
    },
    {
      icon: "⚖️",
      title: "مطابقة وتسوية الحسابات",
      desc: "تسوية ومطابقة الحسابات وكشوف البنك والفروقات غير المطابقة",
      prompt: "⚖️ قم بمطابقة وتسوية الحسابات بين أوراق العمل: طابق كشف الحساب مع الدفاتر الداخلية، وأنشئ ورقة تقرير مطابقة توضح الحركات المتطابقة والفروقات غير المتطابقة مع بطاقات الملخص والتلوين الشرطي."
    },
    {
      icon: "📊",
      title: "لوحة تحكم تفاعلية (KPI Dashboard)",
      desc: "بطاقات مؤشرات ملونة ورسمين بيانيين جذابين مربوطين بالبيانات",
      prompt: "📊 أنشئ لوحة تحكم إدارية (KPI Dashboard) متكاملة في ورقة جديدة: ضع بطاقات ملونة للمؤشرات الرئيسية، ورسمين بيانيين جذابين مربوطين ببيانات الشيت."
    },
    {
      icon: "📈",
      title: "محاكاة السيناريوهات المالية (What-If)",
      desc: "تحليل حساسية ومقارنة سيناريو متفائل وواقعي ومتشائم مع رسم بياني",
      prompt: "📈 ابنِ نموذج محاكاة سيناريوهات وحساسية القرار المالي (What-If Analysis) في ورقة جديدة: ضع خلايا لنسب المتغيرات، وجدول مقارنة متفائل وواقعي ومتشائم مع رسم بياني توضيحي."
    },
    {
      icon: "📋",
      title: "قائمة منسدلة داخل الخلايا (Data Validation)",
      desc: "إنشاء قيود تحقق وقائمة خيارات بالخلية لمنع الأخطاء البشرية",
      prompt: "📋 أنشئ قائمة منسدلة داخل الخلايا (Data Validation) للعمود المحدد لمنع الأخطاء البشرية وتقييد القيم المدخلة بخيارات محددة مسبقاً."
    },
    {
      icon: "📑",
      title: "تصميم فاتورة ضريبية رسمية (A4 جاهزة)",
      desc: "فاتورة متكاملة مع الضريبة وقائمة طريقة الدفع والتهيئة الفورية للطباعة",
      prompt: "📑 صمم نموذج فاتورة مبيعات ضريبية رسمية (Tax Invoice) متكاملة الأركان في ورقة جديدة: جدول البنود، حساب ضريبة القيمة المضافة، وقائمة منسدلة لطريقة الدفع وتهيئتها للطباعة A4 فوراً."
    },
    {
      icon: "🔮",
      title: "التنبؤ بالاتجاهات المستقبلية (AI Forecasting)",
      desc: "توقع أرقام الفترات القادمة بمعادلات إكسيل مع تظليل صفوف التوقع",
      prompt: "🔮 حلل الاتجاه الزمني وتوقع أرقام الفترات القادمة (Forecasting) باستخدام معادلات الإكسيل مع تمييز صفوف التوقع بلون مخصص."
    },
    {
      icon: "🧭",
      title: "فحص وتشخيص شامل للمصنف (Auto-Audit)",
      desc: "قراءة كافة التابات وتشخيص نقاط القوة والمخاطر وأهم 3 توصيات",
      prompt: "🧭 قم بفحص شامل للمصنف بالكامل: اقرأ كافة أوراق العمل وقدم تقريراً تشخيصياً لمكامن القوة، وأبرز المخاطر المعلقة، وأهم 3 توصيات إدارية قابلة للتنفيذ."
    },
    {
      icon: "⚡",
      title: "توليد كود ماكرو VBA للأتمتة",
      desc: "كود ماكرو نظيف وموثق مع شرح خطوات وضعه في محرر الماكرو (Alt+F11)",
      prompt: "⚡ اكتب لي كود ماكرو VBA نظيف وموثق لأتمتة هذه العملية في إكسيل مع شرح خطوات إضافته في محرر الماكرو (Alt+F11)."
    },
    {
      icon: "🧩",
      title: "شرح وتفكيك وفحص المعادلات",
      desc: "شرح منطق المعادلات في الخلايا واقتراح بدائل أسرع وأكثر كفاءة",
      prompt: "🧩 فكك واشرح المعادلات الموجودة في الخلايا المحددة بالتفصيل، وافحص هل بها أخطاء أو بطء واقترح بدائل عصرية وأكثر كفاءة."
    },
    {
      icon: "🖨️",
      title: "تهيئة الصفحة للطباعة وتصدير PDF",
      desc: "ضبط المقاس على A4 والاتجاه الأفقي Landscape وملاءمة العرض لصفحة واحدة",
      prompt: "🖨️ جهز هذا الشيت للطباعة والتصدير كـ PDF: اضبط المقاس على A4 والاتجاه الأفقي Landscape واجعل العرض ملائم لصفحة واحدة fit-to-1-page."
    },
    {
      icon: "🎨",
      title: "تنسيق مالي احترافي وتلوين شرطي",
      desc: "تنسيق مالي جذاب للعناوين والحدود وتطبيق التدرج الحراري (Heatmap)",
      prompt: "🎨 قم بتنسيق الجدول الحالي بتصميم مالي احترافي مع تمييز العناوين والألوان والحدود والتنسيق الشرطي."
    },
    {
      icon: "📊",
      title: "اقتراح وإنشاء رسم بياني توضيحي",
      desc: "رسم بياني ذكي يتناسب بدقة مع سياق وتوزيع أرقامك",
      prompt: "📈 اقترح وأنشئ رسماً بيانياً مناسباً جداً لهذه البيانات."
    },
    {
      icon: "➕",
      title: "إضافة صف الإجمالي النهائي والمعادلات",
      desc: "دوال SUM و AVERAGE وتنسيق الخط العريض والخلفية المميزة",
      prompt: "➕ أضف صف إجمالي نهائي في الأسفل مع دوال المجموع أو المتوسط المناسبة."
    },
    {
      icon: "💡",
      title: "تقرير تحليلي تنفيذي للإدارة العليا",
      desc: "تحليل معمق للأرقام وأبرز المؤشرات مع أهم 3 توصيات للقيادة",
      prompt: "💡 اكتب لي تقريراً تحليلياً تنفيذياً مفصلاً يفسر هذه الأرقام مع أهم 3 توصيات للإدارة."
    }
  ];

  const WORD_QUICK_ACTIONS = [
    {
      icon: "📝",
      title: "صياغة عقد رسمي / اتفاقية عمل محكمة",
      desc: "عقد متكامل الأركان، التمهيد، التزامات الأطراف، وشروط السداد وفض النزاعات",
      prompt: "📝 صياغة عقد رسمي محكم الأركان (اتفاقية تقديم خدمات وتوريد) يتضمن التمهيد والتزامات الأطراف وشروط السداد والقوة القاهرة وفض النزاعات."
    },
    {
      icon: "📑",
      title: "كتابة تقرير إداري واستراتيجي للقيادة",
      desc: "تقرير شامل: ملخص تنفيذي، محاور العمل، مؤشرات الأداء، والتوصيات",
      prompt: "📑 اكتب تقريراً إدارياً واستراتيجياً شاملاً للعرض على الإدارة العليا: ملخص تنفيذي، محاور العمل، مؤشرات الأداء، والتوصيات الاستراتيجية."
    },
    {
      icon: "🔍",
      title: "تدقيق لغوي ونحوي وإعادة صياغة باحترافية",
      desc: "فحص النص المحدد وتصحيح الأخطاء وإعادة صياغته بأسلوب عربي فصيح ورصين",
      prompt: "🔍 قم بالتدقيق اللغوي والنحوي الشامل للنص المحدد في المستند، وتصحيح الأخطاء وإعادة صياغته بأسلوب عربي فصيح ورصين."
    },
    {
      icon: "📊",
      title: "إنشاء جدول وورد أنيق بألوان الشركات",
      desc: "جدول رسمي منسق بالعناوين والبيانات المنظمة وحدوداً ناعمة في Word",
      prompt: "📊 أنشئ جدولاً رسمياً منسقاً بتنسيق شركات أنيق يتضمن العناوين والبيانات المنظمة وحدوداً ناعمة."
    },
    {
      icon: "💡",
      title: "تلخيص المستند واستخراج أهم النقاط",
      desc: "نقاط تنفيذية مركزة تبرز القرارات والمهام المطلوبة ومواعيد التنفيذ",
      prompt: "💡 لخص لي هذا المستند في نقاط تنفيذية مركزة تبرز القرارات والمهام المطلوبة ومواعيد التنفيذ."
    },
    {
      icon: "🌐",
      title: "ترجمة احترافية للنص المحدد",
      desc: "ترجمة للغة الأعمال الاحترافية مع الحفاظ التام على المصطلحات الدقيقة",
      prompt: "🌐 ترجم النص المحدد إلى لغة الأعمال الاحترافية مع الحفاظ التام على المعنى والمصطلحات المعتمدة."
    },
    {
      icon: "⚡",
      title: "استخراج الجداول وتجهيزها لنقلها إلى Excel",
      desc: "استخراج كافة الأرقام والجداول وتجهيزها لنسخها في شيت إكسيل",
      prompt: "⚡ استخراج كافة البيانات والأرقام والجداول الموجودة في هذا المستند وجهزها بصيغة مرتبة للنسخ في إكسيل."
    },
    {
      icon: "💻",
      title: "كود ماكرو VBA لأتمتة عمليات Word",
      desc: "كود ماكرو لأتمتة مهمة متكررة في وورد مع خطوات تشغيله بـ Alt+F11",
      prompt: "💻 اكتب لي كود ماكرو VBA لأتمتة هذه المهمة المتكررة في Word مع شرح خطوات تشغيله (Alt+F11)."
    }
  ];

  function createActionCard(act) {
    const card = document.createElement("div");
    card.className = "studio-card";
    card.title = act.desc;
    card.innerHTML = `
      <span class="studio-card-icon">${act.icon}</span>
      <div class="studio-card-body">
        <span class="studio-card-title">${act.title}</span>
        <span class="studio-card-desc">${act.desc}</span>
      </div>
    `;
    card.addEventListener("click", () => {
      userInput.value = act.prompt;
      userInput.focus();
      handleSend();
    });
    return card;
  }

  // دالة بناء وتحديث ستوديو الأوامر السريعة المباشرة (Fluent Studio Grid)
  function renderQuickActions(host) {
    const actions = host === "word" ? WORD_QUICK_ACTIONS : EXCEL_QUICK_ACTIONS;
    
    if (actionHostBadge) {
      actionHostBadge.textContent = host === "word" ? "Word" : "Excel";
    }

    // تعبئة الكروت الأساسية الأربعة الأولى
    if (studioPrimaryGrid) {
      studioPrimaryGrid.innerHTML = "";
      actions.slice(0, 4).forEach(act => {
        studioPrimaryGrid.appendChild(createActionCard(act));
      });
    }

    // تعبئة بقية الكروت في الدرج القابل للتوسيع
    if (studioSecondaryGrid) {
      studioSecondaryGrid.innerHTML = "";
      actions.slice(4).forEach(act => {
        studioSecondaryGrid.appendChild(createActionCard(act));
      });
    }

    if (toggleAllActionsText) {
      toggleAllActionsText.textContent = `عرض الكل (${actions.length})`;
    }

    // دعم احتياطي لأي استدعاءات قديمة للقائمة
    if (quickActionMenu) {
      quickActionMenu.innerHTML = "";
      actions.forEach(act => {
        const item = document.createElement("div");
        item.className = "custom-dropdown-item";
        item.innerHTML = `
          <span class="custom-dropdown-item-icon">${act.icon}</span>
          <div class="custom-dropdown-item-body">
            <span class="custom-dropdown-item-title">${act.title}</span>
            <span class="custom-dropdown-item-desc">${act.desc}</span>
          </div>
        `;
        item.addEventListener("click", (e) => {
          e.stopPropagation();
          userInput.value = act.prompt;
          userInput.focus();
          handleSend();
        });
        quickActionMenu.appendChild(item);
      });
    }
  }

  // زر توسيع وطي درج الأوامر الإضافية
  if (btnToggleAllActions && studioDrawer) {
    btnToggleAllActions.addEventListener("click", () => {
      const isClosed = studioDrawer.style.display === "none" || !studioDrawer.style.display;
      if (isClosed) {
        studioDrawer.style.display = "block";
        if (toggleAllActionsText) toggleAllActionsText.textContent = "عرض أقل";
        if (toggleAllActionsIcon) toggleAllActionsIcon.textContent = "▴";
      } else {
        studioDrawer.style.display = "none";
        const actions = currentHost === "word" ? WORD_QUICK_ACTIONS : EXCEL_QUICK_ACTIONS;
        if (toggleAllActionsText) toggleAllActionsText.textContent = `عرض الكل (${actions.length})`;
        if (toggleAllActionsIcon) toggleAllActionsIcon.textContent = "▾";
      }
    });
  }

  function applyHostAdaptation(host) {
    currentHost = host;
    if (host === "word") {
      if (appLogo) appLogo.textContent = "📝";
      if (appTitle) appTitle.textContent = "المساعد الذكي (Word AI)";
      if (btnExportWord) btnExportWord.title = "استخراج الجداول ونقلها لإكسيل";
      
      const currentPersona = personaSelect.value;
      if (!["writer", "legal", "editor"].includes(currentPersona)) {
        personaSelect.value = "writer";
        personaBadge.textContent = personaDescriptions["writer"];
      }
      renderQuickActions("word");
    } else {
      if (appLogo) appLogo.textContent = "📊";
      if (appTitle) appTitle.textContent = "المساعد الذكي (Excel AI)";
      if (btnExportWord) btnExportWord.title = "تصدير التحليل الحالي لتقرير Word منسق (.doc)";
      
      const currentPersona = personaSelect.value;
      if (!["cfo", "analyst", "operations", "executive"].includes(currentPersona)) {
        personaSelect.value = "cfo";
        personaBadge.textContent = personaDescriptions["cfo"];
      }
      renderQuickActions("excel");
    }
    updateContextDisplay();
  }

  // التهيئة المبدئية للأوامر السريعة فور تحميل السكريبت
  renderQuickActions("excel");

  // تهيئة مايكروسوفت أوفيس (كشف هل التطبيق إكسيل أم وورد)
  if (typeof Office !== "undefined") {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Word) {
        console.log("تم الاتصال ببرنامج وورد بنجاح!");
        applyHostAdaptation("word");
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          updateContextDisplay
        );
      } else {
        console.log("تم الاتصال ببرنامج إكسيل بنجاح!");
        applyHostAdaptation("excel");
        Office.context.document.addHandlerAsync(
          Office.EventType.DocumentSelectionChanged,
          updateContextDisplay
        );
      }
    });
  }

  // تحميل الإعدادات المحفوظة
  loadSettings();

  // ----------------------------------------------------
  // تحديث سياق الشيت أو المستند
  // ----------------------------------------------------
  async function updateContextDisplay() {
    try {
      if (currentHost === "word") {
        const ctx = await WordEngine.getContext();
        if (ctx) {
          currentSheetName.textContent = `مستند: نشط`;
          currentSelection.textContent = ctx.hasSelection 
            ? `محدد: "${ctx.selectedText.substring(0, 18)}..."` 
            : `المستند (${ctx.paragraphsCount} فقرة)`;
        }
      } else {
        const ctx = await ExcelEngine.getContext();
        if (ctx) {
          currentSheetName.textContent = `ورقة: ${ctx.sheetName}`;
          currentSelection.textContent = `التحديد: ${ctx.address}`;
        }
      }
    } catch (err) {
      console.error("فشل تحديث السياق:", err);
    }
  }

  btnRefreshContext.addEventListener("click", updateContextDisplay);

  // ----------------------------------------------------
  // تغيير شخصية ومهنة الوكيل
  // ----------------------------------------------------
  personaSelect.addEventListener("change", (e) => {
    const selected = e.target.value;
    personaBadge.textContent = personaDescriptions[selected] || "";
    localStorage.setItem("gemini_persona", selected);
    
    // إضافة رسالة ترحيبية من الشخصية الجديدة
    const personaName = personaSelect.options[personaSelect.selectedIndex].text;
    appendMessage(
      `تم التحويل إلى وضعية: **${personaName}**.\nأنا جاهز لمساعدتك وفق أفضل الممارسات والمعايير المهنية لهذا التخصص!`,
      "assistant"
    );
  });

  // ----------------------------------------------------
  // نافذة الإعدادات
  // ----------------------------------------------------
  function loadSettings() {
    apiKeyInput.value = localStorage.getItem("gemini_api_key") || "";
    let savedModel = localStorage.getItem("gemini_model") || "gemini-2.5-flash";
    if (savedModel === "gemini-3.8-flash") {
      savedModel = "gemini-2.5-flash";
      localStorage.setItem("gemini_model", "gemini-2.5-flash");
    }
    // التأكد من وجود الخيار بالقائمة المنسدلة
    const hasOption = Array.from(modelSelect.options).some(o => o.value === savedModel);
    if (!hasOption) {
      const newOpt = document.createElement("option");
      newOpt.value = savedModel;
      newOpt.textContent = savedModel;
      modelSelect.appendChild(newOpt);
    }
    modelSelect.value = savedModel;
    customInstructions.value = localStorage.getItem("gemini_custom_instructions") || "";
    
    const savedPersona = localStorage.getItem("gemini_persona") || "cfo";
    personaSelect.value = savedPersona;
    personaBadge.textContent = personaDescriptions[savedPersona] || "";

    // تنبيه المستخدم إذا لم يضع المفتاح بعد
    if (!apiKeyInput.value) {
      setTimeout(() => {
        settingsModal.style.display = "flex";
      }, 600);
    }
  }

  // زر استكشاف وفحص النماذج المصرحة لمفتاح المستخدم تلقائياً
  if (btnDetectModels) {
    btnDetectModels.addEventListener("click", async () => {
      const key = apiKeyInput.value.trim() || localStorage.getItem("gemini_api_key");
      if (!key) {
        alert("يرجى إدخال مفتاح API أولاً لتتمكن من فحص النماذج!");
        return;
      }
      btnDetectModels.disabled = true;
      btnDetectModels.textContent = "⏳ جاري الفحص...";
      if (modelDetectStatus) {
        modelDetectStatus.textContent = "جاري الاتصال بخوادم Google AI Studio للاستعلام عن النماذج...";
        modelDetectStatus.style.color = "var(--text-secondary)";
      }

      try {
        const available = await GeminiService.listAvailableModels(key);
        if (available && available.length > 0) {
          modelSelect.innerHTML = "";
          available.forEach(m => {
            const opt = document.createElement("option");
            opt.value = m.id;
            opt.textContent = `${m.name} (${m.id})`;
            modelSelect.appendChild(opt);
          });
          const best = available.find(m => m.id.includes("2.5") && m.id.includes("flash"))
            || available.find(m => m.id.includes("2.0") && m.id.includes("flash"))
            || available.find(m => m.id.includes("flash"))
            || available[0];
          if (best) modelSelect.value = best.id;

          if (modelDetectStatus) {
            modelDetectStatus.innerHTML = `✅ تم اكتشاف <strong>${available.length}</strong> نموذج بنجاح! تم تحديد <strong>${modelSelect.value}</strong> كأفضل خيار.`;
            modelDetectStatus.style.color = "var(--primary-color)";
          }
        } else {
          if (modelDetectStatus) {
            modelDetectStatus.textContent = "⚠️ لم يتم استرجاع نماذج مباشرة، ولكن نظام التبديل الذكي سيعمل تلقائياً عند إرسال أي رسالة.";
            modelDetectStatus.style.color = "var(--warning-color)";
          }
        }
      } catch (err) {
        if (modelDetectStatus) {
          modelDetectStatus.textContent = `⚠️ تعذر الاتصال: ${err.message}`;
          modelDetectStatus.style.color = "var(--danger-color)";
        }
      } finally {
        btnDetectModels.disabled = false;
        btnDetectModels.textContent = "🔄 فحص النماذج";
      }
    });
  }

  btnSettings.addEventListener("click", () => {
    settingsModal.style.display = "flex";
  });

  function closeSettings() {
    settingsModal.style.display = "none";
  }

  btnCloseSettings.addEventListener("click", closeSettings);
  btnCancelSettings.addEventListener("click", closeSettings);

  btnSaveSettings.addEventListener("click", () => {
    const key = apiKeyInput.value.trim();
    if (!key) {
      alert("يرجى إدخال مفتاح API للمتابعة!");
      return;
    }

    localStorage.setItem("gemini_api_key", key);
    localStorage.setItem("gemini_model", modelSelect.value);
    localStorage.setItem("gemini_custom_instructions", customInstructions.value.trim());
    closeSettings();

    appendMessage("✅ تم حفظ مفتاح الذكاء الاصطناعي والإعدادات بنجاح! أنا الآن جاهز تماماً لتنفيذ أوامرك.", "assistant");
  });

  // ----------------------------------------------------
  // نافذة مكتبة القوالب الجاهزة وسير العمل
  // ----------------------------------------------------
  function openTemplates() {
    loadCustomTemplates();
    if (templatesModal) templatesModal.style.display = "flex";
  }

  function closeTemplates() {
    if (templatesModal) templatesModal.style.display = "none";
  }

  if (btnTemplates) btnTemplates.addEventListener("click", openTemplates);
  if (btnCloseTemplates) btnCloseTemplates.addEventListener("click", closeTemplates);

  // تشغيل القوالب الجاهزة بنقرة واحدة
  document.querySelectorAll(".template-card").forEach(card => {
    card.addEventListener("click", () => {
      const prompt = card.getAttribute("data-prompt");
      if (prompt) {
        userInput.value = prompt;
        closeTemplates();
        userInput.focus();
        handleSend();
      }
    });
  });

  // إدارة وحفظ القوالب المخصصة للمستخدم
  function loadCustomTemplates() {
    if (!customTemplatesList) return;
    const raw = localStorage.getItem("excel_custom_templates");
    let templates = [];
    try {
      templates = raw ? JSON.parse(raw) : [];
    } catch (e) {
      templates = [];
    }

    if (templates.length === 0) {
      customTemplatesList.innerHTML = `<div style="font-size:11.5px; color:var(--text-muted); padding:6px 0;">لم تقم بحفظ أي قوالب مخصصة بعد. احفظ أوامرك المتكررة بالأسفل لتصل إليها فوراً!</div>`;
      return;
    }

    customTemplatesList.innerHTML = "";
    templates.forEach((t, idx) => {
      const item = document.createElement("div");
      item.className = "custom-template-item";
      
      const info = document.createElement("div");
      info.className = "custom-template-item-info";
      info.innerHTML = `<strong>⭐ ${t.title}</strong><span>${t.prompt.substring(0, 50)}...</span>`;
      info.addEventListener("click", () => {
        userInput.value = t.prompt;
        closeTemplates();
        userInput.focus();
        handleSend();
      });

      const delBtn = document.createElement("button");
      delBtn.className = "btn-delete-template";
      delBtn.innerHTML = "🗑️";
      delBtn.title = "حذف هذا القالب";
      delBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        if (confirm(`هل تريد حذف قالب "${t.title}"؟`)) {
          templates.splice(idx, 1);
          localStorage.setItem("excel_custom_templates", JSON.stringify(templates));
          loadCustomTemplates();
        }
      });

      item.appendChild(info);
      item.appendChild(delBtn);
      customTemplatesList.appendChild(item);
    });
  }

  if (btnSaveNewTemplate) {
    btnSaveNewTemplate.addEventListener("click", () => {
      const title = newTemplateTitle.value.trim();
      const prompt = newTemplatePrompt.value.trim();
      if (!title || !prompt) {
        alert("يرجى إدخال عنوان القالب ونص الأمر!");
        return;
      }

      const raw = localStorage.getItem("excel_custom_templates");
      let templates = [];
      try {
        templates = raw ? JSON.parse(raw) : [];
      } catch (e) {
        templates = [];
      }

      templates.push({ title, prompt });
      localStorage.setItem("excel_custom_templates", JSON.stringify(templates));
      newTemplateTitle.value = "";
      newTemplatePrompt.value = "";
      loadCustomTemplates();
      alert("✅ تم حفظ القالب بنجاح في مكتبتك!");
    });
  }

  // ----------------------------------------------------
  // إرسال الرسائل والأوامر
  // ----------------------------------------------------
  async function handleSend() {
    const text = userInput.value.trim();
    if (!text && !attachedFile) return;

    const currentAttachment = attachedFile; // احتفاظ بنسخة من الملف المرفق (صورة أو PDF) إن وجد
    let messageToSend = text;
    if (!messageToSend && currentAttachment) {
      if (currentAttachment.isPdf) {
        messageToSend = "قم بقراءة واستخراج كافة الجداول والبيانات من مستند الـ PDF المرفق وتفريغها في الإكسيل بنفس التنسيقات والأرقام بدقة.";
      } else {
        messageToSend = "قم بقراءة الجدول الموجود في الصورة المرفقة وإنشائه في الإكسيل بنفس بياناته وتنسيقاته بدقة.";
      }
    }

    // إضافة رسالة المستخدم للواجهة مع المرفق
    appendMessage(messageToSend, "user", null, currentAttachment);
    userInput.value = "";
    userInput.style.height = "auto";
    clearAttachedFile();
    
    setLoadingState(true);

    try {
      // جلب بيانات الشيت أو المستند إذا كان الخيار مفعلاً
      let contextData = null;
      if (chkIncludeContext.checked) {
        contextData = currentHost === "word" ? await WordEngine.getContext() : await ExcelEngine.getContext();
      }

      // إرسال الطلب لـ Gemini مع المرفق وسياق المستند/الشيت والذاكرة التراكمية
      const response = await GeminiService.sendMessage(messageToSend, contextData, {
        host: currentHost,
        persona: personaSelect.value,
        safeMode: chkNewSheetMode ? chkNewSheetMode.checked : true,
        chatHistory: chatHistory,
        attachment: currentAttachment ? {
          mimeType: currentAttachment.mimeType,
          data: currentAttachment.data,
          isPdf: currentAttachment.isPdf
        } : null
      });

      // حفظ التبادل في ذاكرة المحادثة (Multi-Turn Memory)
      chatHistory.push({
        role: "user",
        parts: [{ text: messageToSend }]
      });
      chatHistory.push({
        role: "model",
        parts: [{ text: response.reply || "" }]
      });
      // الاحتفاظ بآخر 6 أطراف حوار (3 تبادلات) لتوفير الذاكرة والسرعة
      if (chatHistory.length > 6) {
        chatHistory = chatHistory.slice(-6);
      }

      // عرض رد الذكاء الاصطناعي
      appendMessage(response.reply, "assistant", response.suggestions);

      // تنفيذ الأوامر على الشيت أو المستند فورياً إن وجدت
      if (response.actions && response.actions.length > 0) {
        const executionResult = currentHost === "word"
          ? await WordEngine.executeActions(response.actions)
          : await ExcelEngine.executeActions(response.actions);
        renderActionResults(executionResult);
        await updateContextDisplay();
      }

    } catch (error) {
      console.error(error);
      appendMessage(`⚠️ **تنبيه:** ${error.message}`, "assistant error");
    } finally {
      setLoadingState(false);
    }
  }

  btnSend.addEventListener("click", handleSend);

  // إرسال عبر زر Enter (مع Shift+Enter لسطر جديد)
  userInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  });

  // تكبير حقل النص تلقائياً مع الكتابة
  userInput.addEventListener("input", function() {
    this.style.height = "auto";
    this.style.height = (this.scrollHeight) + "px";
  });

  // مسح المحادثة وإعادة تعيين الذاكرة
  btnClearChat.addEventListener("click", () => {
    if (confirm("هل تريد مسح سجل المحادثة الحالي وإعادة ضبط الذاكرة؟")) {
      chatHistory = [];
      chatContainer.innerHTML = `
        <div class="message assistant-message welcome-message">
          <div class="message-avatar">🤖</div>
          <div class="message-content">
            <p>تم مسح المحادثة وتفريغ الذاكرة. أنا جاهز لأوامرك الجديدة على الشيت!</p>
          </div>
        </div>
      `;
    }
  });

  // إدارة قائمة الإجراءات السريعة المنسدلة الذكية (Custom Dropdown)
  if (quickActionBtn && quickActionDropdown) {
    quickActionBtn.addEventListener("click", (e) => {
      e.stopPropagation();
      quickActionDropdown.classList.toggle("open");
    });

    // إغلاق القائمة عند النقر خارجها في أي مكان في الصفحة
    document.addEventListener("click", (e) => {
      if (!quickActionDropdown.contains(e.target)) {
        quickActionDropdown.classList.remove("open");
      }
    });
  }

  // دعم التوافقية مع select القديم
  if (quickActionSelect) {
    quickActionSelect.addEventListener("change", (e) => {
      const prompt = e.target.value;
      if (prompt) {
        userInput.value = prompt;
        quickActionSelect.selectedIndex = 0;
        userInput.focus();
        handleSend();
      }
    });
  }

  // زر تصدير التقرير لوورد أو استخراج الجداول لإكسيل
  if (btnExportWord) {
    btnExportWord.addEventListener("click", async () => {
      if (currentHost === "word") {
        userInput.value = "قم باستخراج كافة الجداول والبيانات الرقمية من هذا المستند وتجهيزها بصيغة مرتبة لنقلها إلى إكسيل.";
        userInput.focus();
        handleSend();
      } else {
        try {
          const ctx = await ExcelEngine.getContext();
          let tableHeaders = [];
          let tableRows = [];
          if (ctx && ctx.values && ctx.values.length > 0) {
            tableHeaders = ctx.values[0] || [];
            tableRows = ctx.values.slice(1) || [];
          }

          const lastReply = chatContainer.querySelector(".assistant-message:last-child .message-content");
          const summaryText = lastReply ? lastReply.innerText.substring(0, 500) : "تقرير تحليلي تم استخراجه من مصنف إكسيل.";

          WordEngine.exportExcelToWordDocument(
            `تقرير_${ctx ? ctx.sheetName : 'التحليل_المالي'}`,
            summaryText,
            tableHeaders,
            tableRows
          );
          appendMessage("📄 **تم تصدير تقرير Word منسق بنجاح!** تم تجهيز المستند وتحميله لتفتحه مباشرة في Microsoft Word.", "assistant");
        } catch (e) {
          alert("حدث خطأ أثناء تصدير ملف وورد: " + e.message);
        }
      }
    });
  }

  // ----------------------------------------------------
  // وظائف مساعدة لعرض الرسائل وتأثيرات الواجهة
  // ----------------------------------------------------
  function appendMessage(text, sender, suggestions = null, attachment = null) {
    const messageDiv = document.createElement("div");
    messageDiv.className = `message ${sender}-message`;

    const avatar = document.createElement("div");
    avatar.className = "message-avatar";
    avatar.textContent = sender === "user" ? "👤" : "🤖";

    const content = document.createElement("div");
    content.className = "message-content";

    // عرض المرفق (صورة أو مستند PDF) إن وجد في فقاعة الرسالة
    if (attachment) {
      if (attachment.isPdf) {
        const pdfBadge = document.createElement("div");
        pdfBadge.className = "message-pdf-badge";
        pdfBadge.innerHTML = `<span class="pdf-badge-icon">📕</span><span>مستند PDF مرفق: ${attachment.name || "ملف.pdf"}</span>`;
        content.appendChild(pdfBadge);
      } else if (attachment.dataUrl) {
        const img = document.createElement("img");
        img.src = attachment.dataUrl;
        img.className = "message-image";
        img.title = "انقر لفتح الصورة بحجم كامل";
        img.addEventListener("click", () => {
          const w = window.open("");
          if (w) {
            w.document.write(`<img src="${attachment.dataUrl}" style="max-width:100%; height:auto;">`);
          }
        });
        content.appendChild(img);
      }
    }

    const textDiv = document.createElement("div");
    textDiv.innerHTML = formatMarkdown(text);
    content.appendChild(textDiv);

    // إضافة أزرار المقترحات التفاعلية للاستكمال بنقرة واحدة
    if (suggestions && Array.isArray(suggestions) && suggestions.length > 0) {
      const suggestionsBox = document.createElement("div");
      suggestionsBox.className = "message-suggestions";

      const title = document.createElement("div");
      title.className = "suggestions-title";
      title.innerHTML = "<span>⚡ خيارات واستكمال مقترحات الوكيل بنقرة واحدة:</span>";
      suggestionsBox.appendChild(title);

      suggestions.forEach(sugText => {
        const btn = document.createElement("button");
        btn.className = "suggestion-btn";
        btn.textContent = sugText;
        btn.title = "اضغط للمتابعة والتنفيذ الفوري";
        btn.addEventListener("click", () => {
          if ((sugText.includes("Word") || sugText.includes("وورد")) && sugText.includes("تصدير") && currentHost === "excel") {
            if (btnExportWord) btnExportWord.click();
          } else {
            userInput.value = sugText;
            handleSend();
          }
        });
        suggestionsBox.appendChild(btn);
      });

      content.appendChild(suggestionsBox);
    }

    // إضافة زر نسخ التقرير التنفيذي لردود الوكيل
    if (sender === "assistant" && text) {
      const actionsBar = document.createElement("div");
      actionsBar.className = "message-actions-bar";

      const btnCopy = document.createElement("button");
      btnCopy.className = "btn-copy-reply";
      btnCopy.innerHTML = "📋 نسخ التقرير";
      btnCopy.title = "نسخ نص التقرير للبريد أو الواتساب";
      btnCopy.addEventListener("click", () => {
        const cleanText = text.replace(/<[^>]*>/g, "");
        navigator.clipboard.writeText(cleanText).then(() => {
          btnCopy.innerHTML = "✅ تم النسخ!";
          setTimeout(() => {
            btnCopy.innerHTML = "📋 نسخ التقرير";
          }, 2000);
        }).catch(() => {
          btnCopy.innerHTML = "✅ تم النسخ!";
          setTimeout(() => {
            btnCopy.innerHTML = "📋 نسخ التقرير";
          }, 2000);
        });
      });

      actionsBar.appendChild(btnCopy);
      content.appendChild(actionsBar);
    }

    messageDiv.appendChild(avatar);
    messageDiv.appendChild(content);
    chatContainer.appendChild(messageDiv);
    chatContainer.scrollTop = chatContainer.scrollHeight;
  }

  function renderActionResults(result) {
    if (!result || !result.details || result.details.length === 0) return;

    const actionCard = document.createElement("div");
    actionCard.className = "action-card";
    
    let html = `
      <div class="action-card-header">
        <span>⚡ تم تنفيذ ${result.details.length} إجراء على الشيت تلقائياً:</span>
      </div>
      <div class="action-card-body">
    `;

    result.details.forEach(detail => {
      html += `<div>• ${detail}</div>`;
    });

    html += `</div>`;

    // زر التراجع الذكي
    if (result.canUndo) {
      html += `
        <div style="margin-top: 6px;">
          <button class="btn-undo-action" id="btnUndoAction">↩️ تراجع عن هذه العملية</button>
        </div>
      `;
    }

    actionCard.innerHTML = html;

    const undoBtn = actionCard.querySelector("#btnUndoAction");
    if (undoBtn) {
      undoBtn.addEventListener("click", async () => {
        undoBtn.disabled = true;
        undoBtn.textContent = "جاري التراجع...";
        try {
          const undoMsg = await ExcelEngine.undoLastAction();
          undoBtn.textContent = "✅ تم التراجع بنجاح!";
          appendMessage(`↩️ **تراجع:** ${undoMsg}`, "assistant");
          await updateContextDisplay();
        } catch (err) {
          alert("تعذر التراجع: " + err.message);
          undoBtn.textContent = "↩️ تراجع عن هذه العملية";
          undoBtn.disabled = false;
        }
      });
    }

    const lastMessage = chatContainer.querySelector(".assistant-message:last-child .message-content");
    if (lastMessage) {
      lastMessage.appendChild(actionCard);
    } else {
      chatContainer.appendChild(actionCard);
    }

    chatContainer.scrollTop = chatContainer.scrollHeight;
  }

  function setLoadingState(isLoading) {
    btnSend.disabled = isLoading;
    if (isLoading) {
      statusIndicator.className = "status-indicator busy";
      statusText.textContent = "جاري التفكير والتنفيذ...";
      
      const typingDiv = document.createElement("div");
      typingDiv.className = "message assistant-message typing-indicator-msg";
      typingDiv.id = "typingIndicator";
      typingDiv.innerHTML = `
        <div class="message-avatar">🤖</div>
        <div class="message-content">
          <div class="typing-dots">
            <span class="typing-dot"></span>
            <span class="typing-dot"></span>
            <span class="typing-dot"></span>
          </div>
        </div>
      `;
      chatContainer.appendChild(typingDiv);
      chatContainer.scrollTop = chatContainer.scrollHeight;
    } else {
      statusIndicator.className = "status-indicator online";
      statusText.textContent = "جاهز للعمل";
      const typing = document.getElementById("typingIndicator");
      if (typing) typing.remove();
    }
  }

  // تحويل مبسط للـ Markdown للعرض الجمالي للنصوص مع دعم كتل الأكواد وزر نسخ الماكرو
  function formatMarkdown(text) {
    if (!text) return "";

    // معالجة كتل الأكواد والماكرو متعددة الأسطر ```lang ... ```
    let formatted = text.replace(/```([a-zA-Z0-9_-]*)\n([\s\S]*?)```/g, (match, lang, code) => {
      const cleanLang = lang.trim() || "VBA";
      const escapedCode = code
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;");
      return `
        <div class="code-block-wrapper">
          <div class="code-block-header">
            <span>💻 ${cleanLang.toUpperCase()}</span>
            <button type="button" class="btn-copy-code" onclick="copyCodeSnippet(this)">📋 نسخ الكود</button>
          </div>
          <pre class="code-block-content"><code>${escapedCode}</code></pre>
        </div>
      `;
    });

    return formatted
      .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
      .replace(/\*(.*?)\*/g, "<em>$1</em>")
      .replace(/`([^`]+)`/g, "<code style='background:#e2e8f0;padding:1px 4px;border-radius:3px;font-family:monospace;'>$1</code>")
      .replace(/\n/g, "<br>");
  }

  // وظيفة عامة لنسخ كود الماكرو بنقرة واحدة
  window.copyCodeSnippet = function(buttonElement) {
    const wrapper = buttonElement.closest(".code-block-wrapper");
    if (!wrapper) return;
    const codeElement = wrapper.querySelector("code");
    if (!codeElement) return;
    const codeText = codeElement.innerText;
    navigator.clipboard.writeText(codeText).then(() => {
      const originalText = buttonElement.innerHTML;
      buttonElement.innerHTML = "✅ تم النسخ!";
      setTimeout(() => {
        buttonElement.innerHTML = originalText;
      }, 2000);
    }).catch(() => {
      buttonElement.innerHTML = "✅ تم النسخ!";
    });
  };
});
