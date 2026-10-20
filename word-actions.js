/**
 * Word Actions Engine - Native Office.js Word Operations
 * مسؤول عن تنفيذ الأوامر والتعديلات وصياغة المستندات والجداول في Microsoft Word
 */

const WordEngine = {
  
  /**
   * قراءة سياق مستند الوورد المفتوح والنص المحدد حالياً
   */
  async getContext() {
    return new Promise((resolve) => {
      // التحقق مما إذا كان Office.js Word محملاً
      if (typeof Word === "undefined") {
        resolve({
          host: "word",
          documentTitle: "مستند تجريبي (خارج وورد)",
          hasSelection: true,
          selectedText: "هذا نص تجريبي محدد داخل المستند لفحص الصياغة والتدقيق.",
          paragraphsCount: 4,
          tablesCount: 1,
          documentSnippet: "المقدمة: تقرير الأعمال السنوي واستعراض الأداء المالي والتشغيلي."
        });
        return;
      }

      Word.run(async (context) => {
        const document = context.document;
        const selection = document.getSelection();
        const body = document.body;
        const paragraphs = body.paragraphs;
        const tables = body.tables;

        selection.load(["text", "isEmpty"]);
        paragraphs.load("text");
        tables.load("items");

        await context.sync();

        const selText = selection.text ? selection.text.trim() : "";
        const pItems = paragraphs.items || [];
        const snippet = pItems
          .map(p => p.text ? p.text.trim() : "")
          .filter(t => t.length > 0)
          .slice(0, 5)
          .join("\n")
          .substring(0, 800);

        resolve({
          host: "word",
          hasSelection: !selection.isEmpty && selText.length > 0,
          selectedText: selText.substring(0, 1500),
          paragraphsCount: pItems.length,
          tablesCount: (tables.items || []).length,
          documentSnippet: snippet
        });
      }).catch(err => {
        console.error("خطأ في قراءة سياق مستند الوورد:", err);
        resolve({
          host: "word",
          hasSelection: false,
          selectedText: "",
          paragraphsCount: 0,
          tablesCount: 0,
          documentSnippet: ""
        });
      });
    });
  },

  /**
   * تنفيذ قائمة الإجراءات المرسلة من الذكاء الاصطناعي مباشرة على مستند الوورد
   */
  async executeActions(actions) {
    if (!actions || !Array.isArray(actions) || actions.length === 0) {
      return { success: true, count: 0, details: [] };
    }

    if (typeof Word === "undefined") {
      console.warn("Word API غير متوفرة في هذه البيئة (تشغيل تجريبي)");
      return {
        success: true,
        count: actions.length,
        details: actions.map(a => `[تجريبي] تنفيذ إجراء ${a.type} على مستند Word`)
      };
    }

    const executedDetails = [];

    await Word.run(async (context) => {
      const document = context.document;
      const body = document.body;

      for (const action of actions) {
        switch (action.type) {
          
          // 1. إدراج فقرة أو عنوان (Insert Paragraph / Heading)
          case "INSERT_PARAGRAPH": {
            if (action.text) {
              const location = action.location || "End"; // "End", "Start", "AfterSelection", "BeforeSelection"
              let targetRange = body;
              let insertLoc = "End";

              if (location === "Start") {
                insertLoc = "Start";
              } else if (location === "AfterSelection" || location === "After") {
                targetRange = document.getSelection();
                insertLoc = "After";
              } else if (location === "BeforeSelection" || location === "Before") {
                targetRange = document.getSelection();
                insertLoc = "Before";
              }

              const p = targetRange.insertParagraph(action.text, insertLoc);

              // تطبيق أنماط العناوين
              if (action.style) {
                try {
                  p.styleBuiltIn = action.style; // e.g. "Heading 1", "Heading 2", "Title", "Normal"
                } catch (e) {
                  // نمط مخصص إن لم يتوفر النمط القياسي
                }
              }

              // تطبيق التنسيقات الإضافية
              if (action.font) {
                if (action.font.bold !== undefined) p.font.bold = action.font.bold;
                if (action.font.italic !== undefined) p.font.italic = action.font.italic;
                if (action.font.color) p.font.color = action.font.color;
                if (action.font.size) p.font.size = action.font.size;
                if (action.font.name) p.font.name = action.font.name;
              }

              if (action.alignment) {
                // "Center", "Right", "Left", "Justified"
                p.alignment = action.alignment;
              }

              const titleText = action.text.length > 30 ? action.text.substring(0, 27) + "..." : action.text;
              executedDetails.push(`إدراج فقرة/عنوان: "${titleText}"`);
            }
            break;
          }

          // 2. استبدال النص المحدد بالنص المحسن أو المترجم (Replace Selection)
          case "REPLACE_SELECTION": {
            if (action.text !== undefined) {
              const selection = document.getSelection();
              selection.insertText(action.text, "Replace");
              executedDetails.push(`استبدال وتحديث النص المحدد بالصياغة الجديدة`);
            }
            break;
          }

          // 3. إنشاء وإدراج جدول منسق في وورد (Insert Table)
          case "INSERT_TABLE": {
            if (action.rows && Array.isArray(action.rows)) {
              let tableValues = [];
              if (action.headers && Array.isArray(action.headers)) {
                tableValues.push(action.headers);
              }
              action.rows.forEach(r => tableValues.push(r));

              const rowCount = tableValues.length;
              const colCount = tableValues[0] ? tableValues[0].length : 1;

              const location = action.location === "AfterSelection" ? "After" : "End";
              const targetRange = action.location === "AfterSelection" ? document.getSelection() : body;

              const table = targetRange.insertTable(rowCount, colCount, location, tableValues);
              table.styleBuiltIn = action.tableStyle || "Grid Table 4 - Accent 1";
              table.styleFirstRow = true;
              table.alignment = "Center";

              executedDetails.push(`إدراج جدول منسق في المستند (${rowCount} صف × ${colCount} عمود)`);
            }
            break;
          }

          // 4. إدراج قائمة نقطية أو مرقمة (Insert Bullet List)
          case "INSERT_BULLET_LIST": {
            if (action.items && Array.isArray(action.items)) {
              action.items.forEach(item => {
                const p = body.insertParagraph("• " + item, "End");
                if (action.boldPrefix && item.includes(":")) {
                  p.font.size = 11;
                }
              });
              executedDetails.push(`إدراج قائمة نقطية من ${action.items.length} بنود`);
            }
            break;
          }

          // 5. البحث والاستبدال التلقائي في كامل المستند (Search and Replace)
          case "SEARCH_AND_REPLACE": {
            if (action.searchText && action.replaceText !== undefined) {
              const searchResults = body.search(action.searchText, {
                matchCase: action.matchCase || false,
                matchWholeWord: action.matchWholeWord || false
              });
              searchResults.load("items");
              await context.sync();

              const count = searchResults.items.length;
              for (let i = 0; i < count; i++) {
                searchResults.items[i].insertText(action.replaceText, "Replace");
              }
              executedDetails.push(`استبدال "${action.searchText}" بـ "${action.replaceText}" (${count} موضع)`);
            }
            break;
          }

          // 6. مسح المستند وتجهيزه لتقرير جديد بالكامل (Clear / New Report Structure)
          case "CLEAR_DOCUMENT": {
            body.clear();
            executedDetails.push(`تفريغ المستند وتجهيزه لكتابة التقرير الجديد`);
            break;
          }
        }
      }

      await context.sync();
    });

    return {
      success: true,
      count: executedDetails.length,
      details: executedDetails
    };
  },

  /**
   * جسر تصدير بيانات الإكسيل إلى مستند Word منسق جاهز للتحميل أو الحفظ
   */
  exportExcelToWordDocument(title, summary, tableHeaders, tableRows) {
    let htmlContent = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${title || "تقرير تنفيذي"}</title>
        <style>
          body { font-family: 'Cairo', 'Segoe UI', Arial, sans-serif; direction: rtl; text-align: right; padding: 20px; }
          h1 { color: #107C41; border-bottom: 2px solid #107C41; padding-bottom: 8px; font-size: 22px; }
          .summary-box { background: #f8fafc; border-right: 4px solid #107C41; padding: 12px; margin: 15px 0; font-size: 14px; line-height: 1.6; }
          table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px; }
          th { background-color: #107C41; color: #ffffff; padding: 10px; border: 1px solid #0d6334; text-align: right; }
          td { padding: 8px 10px; border: 1px solid #cbd5e1; }
          tr:nth-child(even) { background-color: #f1f5f9; }
          .footer { margin-top: 30px; font-size: 11px; color: #64748b; border-top: 1px dashed #cbd5e1; padding-top: 10px; }
        </style>
      </head>
      <body>
        <h1>${title || "تقرير تحليلي تنفيذي"}</h1>
        <div class="summary-box">${summary || ""}</div>
    `;

    if (tableHeaders && tableHeaders.length > 0 && tableRows && tableRows.length > 0) {
      htmlContent += `<table><thead><tr>`;
      tableHeaders.forEach(h => { htmlContent += `<th>${h}</th>`; });
      htmlContent += `</tr></thead><tbody>`;
      tableRows.forEach(row => {
        htmlContent += `<tr>`;
        row.forEach(cell => { htmlContent += `<td>${cell !== null && cell !== undefined ? cell : ''}</td>`; });
        htmlContent += `</tr>`;
      });
      htmlContent += `</tbody></table>`;
    }

    htmlContent += `
        <div class="footer">تم إنشاء هذا المستند تلقائياً بواسطة AI Copilot الموحد لإكسيل ووورد.</div>
      </body>
      </html>
    `;

    const blob = new Blob(['\ufeff' + htmlContent], { type: 'application/msword' });
    const url = URL.createObjectURL(blob);
    const downloadLink = document.createElement("a");
    downloadLink.href = url;
    downloadLink.download = `${title || "تقرير_إداري"}.doc`;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    document.body.removeChild(downloadLink);
    URL.revokeObjectURL(url);
  }
};

window.WordEngine = WordEngine;
