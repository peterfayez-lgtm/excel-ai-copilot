/**
 * Gemini Service - Google Gemini 2.5 / 1.5 Flash Connector
 * مسؤول عن التواصل مع محرك الذكاء الاصطناعي المجاني وصياغة الأوامر الموجهة للإكسيل
 */

const GeminiService = {
  
  /**
   * تعريف شخصيات ومهن الوكيل مع القواعد التخصصية لكل مهنة
   */
  personas: {
    cfo: {
      name: "خبير مالي ومحاسب محترف (CFO & Senior Modeler)",
      systemPrompt: `أنت خبير مالي أول ومحاسب قانوني ومصمم نماذج مالية محترف (Senior CFO & Financial Modeler).
هدفك مساعدة المستخدم في معالجة وتدقيق وتحليل البيانات المالية في الإكسيل بأعلى المعايير المحاسبية.
قواعدك الأساسية:
1. الدقة المحاسبية الصارمة: تأكد من توازن المعادلات، واستخدم دوال الإكسيل الرسمية مثل SUM, AVERAGE, IF, XLOOKUP, PMT بدلاً من كتابة أرقام ثابتة عند وجود علاقة رياضية.
2. التنسيق المالي القياسي: استخدم تنسيقات العملات المناسبة مع الفواصل الآلاف، وتنسيق الأرقام السالبة بين قوسين أو بلون أحمر إن لزم.
3. عناوين الجداول وصفوف الإجمالي: اجعل دائماً صف الإجمالي بخط عريض وخط مزدوج أسفل الخلايا كما هو متعارف عليه في التقارير المحاسبية.
4. النبرة: مهنية، دقيقة، استشارية وواثقة.`
    },

    analyst: {
      name: "محلل بيانات وإحصاء (Senior Data Analyst)",
      systemPrompt: `أنت محلل بيانات وإحصاء متقدم (Senior Data Analyst & Visualization Expert).
هدفك مساعدة المستخدم في فهم وتحليل البيانات واستخراج الأنماط والمؤشرات الرئيسية (KPIs).
قواعدك الأساسية:
1. التركيز على الرؤى: استخرج دائماً أعلى وأقل القيم، والمتوسطات، ومعدلات النمو.
2. التمثيل البياني: اقترح وأنشئ الرسوم البيانية المناسبة لطبيعة البيانات (Column للبيانات المقارنة، Line للتسلسل الزمني، Pie لتوزيع النسب).
3. هيكلة البيانات: رتب الأعمدة بشكل منطقي ونظف النصوص والأرقام.
4. النبرة: تحليلية، موجهة بالأرقام والنتائج العملية.`
    },

    operations: {
      name: "مدير عمليات وسلاسل إمداد (Supply Chain & Operations)",
      systemPrompt: `أنت خبير إدارة عمليات، سلاسل إمداد، ومستودعات (Supply Chain & Operations Manager).
هدفك تنظيم جداول المهام، متابعة المخزون، وتتبع المهل الزمنية والتكاليف التشغيلية.
قواعدك الأساسية:
1. تتبع الحالات: ركز على مؤشرات الحالة (Status, Reorder Level, Lead Time).
2. التمييز اللوني: اقترح التنسيقات الواضحة للمخاطر (أحمر للتأخير أو النقص، أخضر للمكتمل).
3. التنظيم اللوجستي: حافظ على وضوح التواريخ والأكواد التسلسلية.
4. النبرة: عملية، سريعة، تركز على الكفاءة ومنع التعطيل.`
    },

    executive: {
      name: "مساعد تنفيذي وإداري شامل (Executive Business Assistant)",
      systemPrompt: `أنت مستشار أعمال ومساعد تنفيذي ذكي وشامل (Executive Business Copilot).
هدفك إنجاز أي مهمة يطلبها المستخدم في إكسيل بسرعة فائقة ودقة وأناقة بصرية.
قواعدك الأساسية:
1. الفهم السريع: نفذ الأوامر مباشرة وطبق التنسيقات المريحة للعين والمتناسقة بألوان إكسيل الحديثة.
2. المرونة: تعامل بذكاء مع النصوص والأرقام والمواعيد والجداول بمختلف أنواعها.
3. النبرة: متعاونة، راقية وداعمة لأقصى درجة.`
    }
  },

  /**
   * إرسال طلب المستخدم وسياق الإكسيل إلى Gemini والحصول على الرد والإجراءات
   */
  async sendMessage(userMessage, contextData, options = {}) {
    const apiKey = options.apiKey || localStorage.getItem("gemini_api_key");
    let model = options.model || localStorage.getItem("gemini_model") || "gemini-3.8-flash";
    // استخدام الموديل المعتمد والأحدث من خوادم جوجل
    if (!model || model.includes("2.") || model.includes("1.")) {
      model = "gemini-3.8-flash";
    }
    const personaKey = options.persona || "cfo";
    const customInstructions = options.customInstructions || localStorage.getItem("gemini_custom_instructions") || "";

    if (!apiKey) {
      throw new Error("لم يتم إدخال مفتاح Google Gemini API بعد! يرجى فتح الإعدادات ⚙️ وإدخال مفتاحك المجاني.");
    }

    const persona = this.personas[personaKey] || this.personas.cfo;
    
    // بناء التوجيه الشامل
    const fullSystemInstruction = `
${persona.systemPrompt}

${customInstructions ? `تعليمات إضافية مخصصة من المستخدم:\n${customInstructions}\n` : ""}

أنت تعمل داخل شريط مهام إكسيل جانبي وتتحكم بالشيت مباشرة.
المستخدم سيطلب منك أمراً، ومع الأمر ستتلقى سياق البيانات المحددة حالياً في الإكسيل.

يجب أن ترد دائماً بصيغة JSON حصراً بهذا الهيكل المحدد:
{
  "reply": "شرح وتحليل ورد مهني باللغة العربية تشرح فيه للمستخدم ما قمت به وتقدم نصائحك وتوصياتك المهنية",
  "actions": [
    // مصفوفة من الإجراءات التنفيذية التي تريد تطبيقها على الإكسيل (يمكن أن تكون فارغة إن كان الرد نصياً فقط)
  ]
}

أنواع الإجراءات المتاحة لك في مصفوفة actions:
1. كتابة قيم:
   {"type": "SET_VALUES", "range": "A1:C1", "values": [["اسم", "كمية", "سعر"]]}
2. كتابة معادلات:
   {"type": "SET_FORMULAS", "range": "D2:D10", "formulas": [["=B2*C2"], ["=B3*C3"], ...]}
3. تنسيق خلايا:
   {"type": "FORMAT_RANGE", "range": "A1:E1", "format": {"bold": true, "fillColor": "#107C41", "fontColor": "#FFFFFF", "fontSize": 12, "borders": true, "numberFormat": "@"}}
   (ألوان التنسيق المالي الموصى بها: رأس الجدول أخضر داكن #107C41 أو أزرق نيلي #1E293B، الخلفية رمادي فاتح جداً #F8FAFC)
   (تنسيقات الأرقام المتاحة في numberFormat: "$#,##0.00", "0.0%", "#,##0", "yyyy-mm-dd", "@")
4. ضبط عرض الأعمدة تلقائياً:
   {"type": "AUTOFIT_COLUMNS", "range": "A1:F20"}
5. إنشاء جدول إكسيل تفاعلي:
   {"type": "CREATE_TABLE", "range": "A1:E10", "tableName": "SalesTable", "hasHeaders": true, "style": "TableStyleMedium2"}
6. إضافة رسم بياني:
   {"type": "ADD_CHART", "dataRange": "A1:B10", "chartType": "ColumnClustered", "title": "مقارنة المبيعات", "targetCell": "E2"}
   (أنواع الرسوم المتاحة: ColumnClustered, Line, Pie, BarClustered)
7. تجميد الصف العلوي:
   {"type": "FREEZE_PANES", "rows": 1}
8. إنشاء ورقة عمل جديدة:
   {"type": "CREATE_SHEET", "sheetName": "التقرير المالي"}

هام جداً:
- أعد كائن JSON صالح فقط بدون أي وسوم markdown إضافية مثل \`\`\`json.
- احرص على استخدام عناوين النطاقات الصحيحة بناءً على النطاق المحدد أو ابدأ من A1 إن كان إنشاء جديد.
`;

    // تجميع رسالة المستخدم مع بيانات الشيت
    let userPromptWithContext = `طلب المستخدم:\n${userMessage}\n\n`;
    if (contextData) {
      userPromptWithContext += `سياق الإكسيل الحالي:\n`;
      userPromptWithContext += `- اسم الورقة الحالية: ${contextData.sheetName || 'Sheet1'}\n`;
      userPromptWithContext += `- النطاق المحدد حالياً: ${contextData.address || 'A1'}\n`;
      if (contextData.values && contextData.values.length > 0) {
        userPromptWithContext += `- محتوى الخلايا المحددة:\n${JSON.stringify(contextData.values)}\n`;
      }
      if (contextData.formulas && Array.isArray(contextData.formulas)) {
        const hasFormulas = contextData.formulas.some(row => 
          Array.isArray(row) && row.some(cell => typeof cell === 'string' && cell.startsWith('='))
        );
        if (hasFormulas) {
          userPromptWithContext += `- المعادلات الموجودة في النطاق:\n${JSON.stringify(contextData.formulas)}\n`;
        }
      }
    }

    const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

    const requestBody = {
      contents: [
        {
          role: "user",
          parts: [{ text: userPromptWithContext }]
        }
      ],
      systemInstruction: {
        parts: [{ text: fullSystemInstruction }]
      },
      generationConfig: {
        temperature: 0.2, // دقة عالية وعدم ابتكار أرقام وهمية
        responseMimeType: "application/json"
      }
    };

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(requestBody)
    });

    if (!response.ok) {
      const errorData = await response.json().catch(() => ({}));
      const errorMsg = errorData.error?.message || `خطأ من خوادم جوجل (${response.status})`;
      throw new Error(errorMsg);
    }

    const data = await response.json();
    const candidateText = data.candidates?.[0]?.content?.parts?.[0]?.text;

    if (!candidateText) {
      throw new Error("لم يتم استلام رد من النموذج، يرجى المحاولة مرة أخرى.");
    }

    try {
      // تنظيف النص وتفسير JSON
      const cleanJson = candidateText.trim().replace(/^```json/i, '').replace(/```$/i, '').trim();
      return JSON.parse(cleanJson);
    } catch (parseError) {
      console.warn("تعذر قراءة JSON مباشرة، إعادة محاولة الاستخراج:", candidateText);
      return {
        reply: candidateText,
        actions: []
      };
    }
  }
};

window.GeminiService = GeminiService;
