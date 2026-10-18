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
  const actionChips = document.querySelectorAll(".action-chip");

  // عناصر إرفاق ومعاينة الصور
  const btnAttachImage = document.getElementById("btnAttachImage");
  const imageFileInput = document.getElementById("imageFileInput");
  const imagePreviewContainer = document.getElementById("imagePreviewContainer");
  const imagePreviewImg = document.getElementById("imagePreviewImg");
  const imagePreviewName = document.getElementById("imagePreviewName");
  const btnRemoveImage = document.getElementById("btnRemoveImage");

  let attachedImage = null; // { data: string (base64), mimeType: string, dataUrl: string, name: string }

  // نصوص توضيحية للشخصيات
  const personaDescriptions = {
    cfo: "يركز على: ضبط الحسابات، القوائم المالية، تنسيق العملات، وفحص الميزانيات بدقة محاسبية.",
    analyst: "يركز على: استخراج المؤشرات، الرسوم البيانية الإحصائية، وتحليل الاتجاهات والأنماط.",
    operations: "يركز على: تتبع الكميات، المهل الزمنية، مستويات المخزون، والتنبيهات اللونية.",
    executive: "يركز على: السرعة والمرونة، التنسيق الشامل، وصياغة ملخصات الأعمال الأنيقة."
  };

  // وظائف إدارة إرفاق ومعاينة الصور
  function showImagePreview(file, dataUrl) {
    const base64Data = dataUrl.split(",")[1];
    const mimeType = file.type || "image/png";
    const fileName = file.name || "صورة_ملصقة.png";
    
    attachedImage = {
      data: base64Data,
      mimeType: mimeType,
      dataUrl: dataUrl,
      name: fileName
    };

    imagePreviewImg.src = dataUrl;
    imagePreviewName.textContent = fileName.length > 25 ? fileName.substring(0, 22) + "..." : fileName;
    imagePreviewContainer.style.display = "flex";
  }

  function clearAttachedImage() {
    attachedImage = null;
    imagePreviewImg.src = "";
    imagePreviewContainer.style.display = "none";
    if (imageFileInput) imageFileInput.value = "";
  }

  if (btnRemoveImage) {
    btnRemoveImage.addEventListener("click", clearAttachedImage);
  }

  if (btnAttachImage && imageFileInput) {
    btnAttachImage.addEventListener("click", () => {
      imageFileInput.click();
    });

    imageFileInput.addEventListener("change", (e) => {
      const file = e.target.files[0];
      if (!file) return;

      if (!file.type.startsWith("image/")) {
        alert("يرجى اختيار ملف صورة صالح (PNG, JPG, WEBP, إلخ).");
        return;
      }

      const reader = new FileReader();
      reader.onload = (evt) => {
        showImagePreview(file, evt.target.result);
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
            showImagePreview(file, evt.target.result);
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
  // إرسال الرسائل والأوامر
  // ----------------------------------------------------
  async function handleSend() {
    const text = userInput.value.trim();
    if (!text && !attachedImage) return;

    const currentImage = attachedImage; // احتفاظ بنسخة من الصورة المرفقة إن وجدت
    const messageToSend = text || "قم بقراءة الجدول الموجود في الصورة المرفقة وإنشائه في الإكسيل بنفس بياناته وتنسيقاته بدقة.";

    // إضافة رسالة المستخدم للواجهة مع الصورة
    appendMessage(messageToSend, "user", null, currentImage ? currentImage.dataUrl : null);
    userInput.value = "";
    userInput.style.height = "auto";
    clearAttachedImage();
    
    setLoadingState(true);

    try {
      // جلب بيانات الشيت إذا كان الخيار مفعلاً
      let contextData = null;
      if (chkIncludeContext.checked) {
        contextData = await ExcelEngine.getContext();
      }

      // إرسال الطلب لـ Gemini مع الصورة وسياق الشيت
      const response = await GeminiService.sendMessage(messageToSend, contextData, {
        persona: personaSelect.value,
        safeMode: chkNewSheetMode ? chkNewSheetMode.checked : true,
        image: currentImage ? { mimeType: currentImage.mimeType, data: currentImage.data } : null
      });

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

  // مسح المحادثة
  btnClearChat.addEventListener("click", () => {
    if (confirm("هل تريد مسح سجل المحادثة الحالي؟")) {
      chatContainer.innerHTML = `
        <div class="message assistant-message welcome-message">
          <div class="message-avatar">🤖</div>
          <div class="message-content">
            <p>تم مسح المحادثة. أنا جاهز لأوامرك الجديدة على الشيت!</p>
          </div>
        </div>
      `;
    }
  });

  // أزرار الإجراءات السريعة (Quick Action Chips)
  actionChips.forEach(chip => {
    chip.addEventListener("click", () => {
      const prompt = chip.getAttribute("data-prompt");
      userInput.value = prompt;
      userInput.focus();
      handleSend();
    });
  });

  // ----------------------------------------------------
  // وظائف مساعدة لعرض الرسائل وتأثيرات الواجهة
  // ----------------------------------------------------
  function appendMessage(text, sender, suggestions = null, imageUrl = null) {
    const messageDiv = document.createElement("div");
    messageDiv.className = `message ${sender}-message`;

    const avatar = document.createElement("div");
    avatar.className = "message-avatar";
    avatar.textContent = sender === "user" ? "👤" : "🤖";

    const content = document.createElement("div");
    content.className = "message-content";

    // عرض الصورة المرفقة إن وجدت في فقاعة الرسالة
    if (imageUrl) {
      const img = document.createElement("img");
      img.src = imageUrl;
      img.className = "message-image";
      img.title = "انقر لفتح الصورة بحجم كامل";
      img.addEventListener("click", () => {
        const w = window.open("");
        if (w) {
          w.document.write(`<img src="${imageUrl}" style="max-width:100%; height:auto;">`);
        }
      });
      content.appendChild(img);
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
    actionCard.innerHTML = html;

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
