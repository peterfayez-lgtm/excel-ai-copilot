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

  // نصوص توضيحية للشخصيات
  const personaDescriptions = {
    cfo: "يركز على: ضبط الحسابات، القوائم المالية، تنسيق العملات، وفحص الميزانيات بدقة محاسبية.",
    analyst: "يركز على: استخراج المؤشرات، الرسوم البيانية الإحصائية، وتحليل الاتجاهات والأنماط.",
    operations: "يركز على: تتبع الكميات، المهل الزمنية، مستويات المخزون، والتنبيهات اللونية.",
    executive: "يركز على: السرعة والمرونة، التنسيق الشامل، وصياغة ملخصات الأعمال الأنيقة."
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

  // تهيئة مايكروسوفت أوفيس
  if (typeof Office !== "undefined") {
    Office.onReady((info) => {
      if (info.host === Office.HostType.Excel) {
        console.log("تم الاتصال ببرنامج إكسيل بنجاح!");
        updateContextDisplay();
        
        // الاستماع لتغيير تحديد الخلايا لتحديث الشريط فورياً
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
  // تحديث سياق الشيت والتحديد
  // ----------------------------------------------------
  async function updateContextDisplay() {
    try {
      const ctx = await ExcelEngine.getContext();
      if (ctx) {
        currentSheetName.textContent = `ورقة: ${ctx.sheetName}`;
        currentSelection.textContent = `التحديد: ${ctx.address}`;
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
    modelSelect.value = localStorage.getItem("gemini_model") || "gemini-3.8-flash";
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
      // جلب بيانات الشيت إذا كان الخيار مفعلاً
      let contextData = null;
      if (chkIncludeContext.checked) {
        contextData = await ExcelEngine.getContext();
      }

      // إرسال الطلب لـ Gemini مع المرفق وسياق الشيت والذاكرة التراكمية
      const response = await GeminiService.sendMessage(messageToSend, contextData, {
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

      // تنفيذ الأوامر على الشيت فورياً إن وجدت
      if (response.actions && response.actions.length > 0) {
        const executionResult = await ExcelEngine.executeActions(response.actions);
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

  // قائمة الإجراءات السريعة المنسدلة (Quick Actions Dropdown)
  if (quickActionSelect) {
    quickActionSelect.addEventListener("change", (e) => {
      const prompt = e.target.value;
      if (prompt) {
        userInput.value = prompt;
        // إعادة التعيين للخيار الافتراضي ليبقى جاهزاً للاختيار مرة أخرى
        quickActionSelect.selectedIndex = 0;
        userInput.focus();
        handleSend();
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
          userInput.value = sugText;
          handleSend();
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

  // تحويل مبسط للـ Markdown للعرض الجمالي للنصوص
  function formatMarkdown(text) {
    if (!text) return "";
    return text
      .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
      .replace(/\*(.*?)\*/g, "<em>$1</em>")
      .replace(/`([^`]+)`/g, "<code style='background:#e2e8f0;padding:1px 4px;border-radius:3px;font-family:monospace;'>$1</code>")
      .replace(/\n/g, "<br>");
  }
});
