/**
 * Excel Actions Engine - Native Office.js Operations
 * مسؤول عن تنفيذ الأوامر والتعديلات مباشرة على شيت الإكسيل
 */

const ExcelEngine = {
  
  /**
   * فحص وقراءة سياق الشيت والخلايا المحددة حالياً وكافة أوراق المصنف
   */
  async getContext() {
    return new Promise((resolve, reject) => {
      // التحقق مما إذا كان Office.js محملاً
      if (typeof Excel === "undefined") {
        resolve({
          sheetName: "ورقة تجريبية (خارج الإكسيل)",
          address: "A1:C5",
          values: [["المنتج", "الكمية", "السعر"], ["أ", 10, 50], ["ب", 20, 30]],
          formulas: [["", "", ""], ["", "", ""], ["", "", ""]],
          rowCount: 3,
          columnCount: 3,
          allSheets: [
            { name: "ورقة تجريبية", address: "A1:C5", rowCount: 3, headers: ["المنتج", "الكمية", "السعر"] },
            { name: "بيانات_العملاء", address: "A1:D20", rowCount: 20, headers: ["كود_العميل", "اسم_العميل", "المدينة", "الرصيد"] }
          ]
        });
        return;
      }

      Excel.run(async (context) => {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        const selection = context.workbook.getSelectedRange();
        const worksheets = context.workbook.worksheets;
        
        activeSheet.load("name");
        selection.load(["address", "values", "formulas", "rowCount", "columnCount"]);
        worksheets.load("items/name");
        
        await context.sync();

        // 1. فحص النطاقات المستخدمة في جميع أوراق المصنف
        const sheetUsedList = [];
        for (let i = 0; i < worksheets.items.length; i++) {
          const ws = worksheets.items[i];
          const used = ws.getUsedRangeOrNullObject();
          used.load(["address", "rowCount", "columnCount", "isNullObject"]);
          sheetUsedList.push({ name: ws.name, usedObj: used });
        }
        await context.sync();

        // 2. قراءة صف العناوين الأول لكل شيت يحتوي على بيانات
        const headersList = [];
        for (const item of sheetUsedList) {
          if (!item.usedObj.isNullObject) {
            const row0 = item.usedObj.getRow(0);
            row0.load("values");
            headersList.push({
              name: item.name,
              address: item.usedObj.address ? (item.usedObj.address.split("!")[1] || item.usedObj.address) : "A1",
              rowCount: item.usedObj.rowCount || 0,
              rowRange: row0
            });
          } else {
            headersList.push({
              name: item.name,
              address: "فارغة",
              rowCount: 0,
              rowRange: null
            });
          }
        }
        await context.sync();

        const allSheetsReport = headersList.map(item => {
          let headers = [];
          if (item.rowRange && item.rowRange.values && item.rowRange.values[0]) {
            headers = item.rowRange.values[0]
              .map(c => (c !== null && c !== undefined) ? String(c).trim() : "")
              .filter(c => c.length > 0);
          }
          return {
            name: item.name,
            address: item.address,
            rowCount: item.rowCount,
            headers: headers
          };
        });

        resolve({
          sheetName: activeSheet.name,
          address: selection.address.split("!")[1] || selection.address,
          values: selection.values,
          formulas: selection.formulas,
          rowCount: selection.rowCount,
          columnCount: selection.columnCount,
          allSheets: allSheetsReport
        });
      }).catch(err => {
        console.error("خطأ في قراءة سياق الإكسيل:", err);
        resolve(null);
      });
    });
  },

  /**
   * تنفيذ قائمة الإجراءات المرسلة من الذكاء الاصطناعي في معاملة واحدة
   */
  async executeActions(actionsList) {
    if (!actionsList || !Array.isArray(actionsList) || actionsList.length === 0) {
      return { success: true, count: 0, details: [] };
    }

    if (typeof Excel === "undefined") {
      console.warn("تنبيه: بيئة Office.js غير نشطة (معاينة متصفح). تم تسجيل الإجراءات:", actionsList);
      return { success: true, count: actionsList.length, details: actionsList.map(a => a.description || a.type) };
    }

    const executedDetails = [];
    const createdSheetNames = [];

    await Excel.run(async (context) => {
      let currentActive = context.workbook.worksheets.getActiveWorksheet();
      currentActive.load("name");
      await context.sync();
      const initialActiveSheetName = currentActive.name;

      for (const action of actionsList) {
        // 1. نسخ الورقة الحالية بالكامل في ورقة جديدة (لحماية البيانات الأصلية)
        if (action.type === "DUPLICATE_SHEET") {
          const source = currentActive;
          const copy = source.copy(Excel.WorksheetPositionType.after, source);
          if (action.newSheetName) {
            const cleanName = action.newSheetName.replace(/[:\\/?*\[\]]/g, "_").substring(0, 31);
            copy.name = cleanName;
            createdSheetNames.push(cleanName);
          } else {
            copy.load("name");
            await context.sync();
            createdSheetNames.push(copy.name);
          }
          copy.activate();
          currentActive = copy;
          await context.sync();
          executedDetails.push(`نسخ الورقة الأصلية إلى ورقة جديدة باسم: ${copy.name}`);
          continue;
        }

        // 2. إنشاء أو تفعيل ورقة عمل جديدة بأمان تام
        if (action.type === "CREATE_SHEET") {
          if (action.sheetName) {
            const cleanName = action.sheetName.replace(/[:\\/?*\[\]]/g, "_").substring(0, 31);
            const existingSheet = context.workbook.worksheets.getItemOrNullObject(cleanName);
            await context.sync();

            let targetSheet;
            if (existingSheet.isNullObject) {
              targetSheet = context.workbook.worksheets.add(cleanName);
              createdSheetNames.push(cleanName);
            } else {
              targetSheet = existingSheet;
            }
            targetSheet.activate();
            currentActive = targetSheet;
            await context.sync();
            executedDetails.push(`إنشاء/تفعيل الورقة: ${cleanName}`);
          }
          continue;
        }

        // تحديد الورقة المستهدفة للأوامر الأخرى بأمان
        let targetSheet = currentActive;
        if (action.sheetName) {
          const sheetObj = context.workbook.worksheets.getItemOrNullObject(action.sheetName);
          await context.sync();
          if (!sheetObj.isNullObject) {
            targetSheet = sheetObj;
          }
        }

        switch (action.type) {
          
          // 1. كتابة قيم في نطاق محدد
          case "SET_VALUES": {
            if (action.range && action.values) {
              const range = targetSheet.getRange(action.range);
              range.values = action.values;
              executedDetails.push(`تحديث القيم في النطاق ${action.range}`);
            }
            break;
          }

          // 2. كتابة دوال ومعادلات إكسيل
          case "SET_FORMULAS": {
            if (action.range && action.formulas) {
              const range = targetSheet.getRange(action.range);
              // التأكد من أن جميع الصيغ تبدأ بعلامة = وإزالة أي مسافات زائدة
              const cleanFormulas = action.formulas.map(row => {
                if (Array.isArray(row)) {
                  return row.map(formula => {
                    if (typeof formula === "string" && formula.trim() !== "") {
                      const trimmed = formula.trim();
                      return trimmed.startsWith("=") ? trimmed : `=${trimmed}`;
                    }
                    return formula;
                  });
                }
                return row;
              });
              range.formulas = cleanFormulas;
              executedDetails.push(`تطبيق المعادلات في النطاق ${action.range}`);
            }
            break;
          }

          // 3. تنسيق مظهر الخلايا (ألوان، خطوط، حدود، محاذاة)
          case "FORMAT_RANGE": {
            if (action.range) {
              const range = targetSheet.getRange(action.range);
              
              if (action.format) {
                if (action.format.bold !== undefined) range.format.font.bold = action.format.bold;
                if (action.format.fontColor) range.format.font.color = action.format.fontColor;
                if (action.format.fontSize) range.format.font.size = action.format.fontSize;
                if (action.format.fontName) range.format.font.name = action.format.fontName;
                if (action.format.fillColor) range.format.fill.color = action.format.fillColor;
                if (action.format.numberFormat) range.numberFormat = [[action.format.numberFormat]];
                if (action.format.horizontalAlignment) range.format.horizontalAlignment = action.format.horizontalAlignment;
                
                // الحدود (Borders)
                if (action.format.borders) {
                  range.format.borders.getItem('InsideHorizontal').style = 'Continuous';
                  range.format.borders.getItem('InsideVertical').style = 'Continuous';
                  range.format.borders.getItem('EdgeTop').style = 'Continuous';
                  range.format.borders.getItem('EdgeBottom').style = 'Continuous';
                  range.format.borders.getItem('EdgeLeft').style = 'Continuous';
                  range.format.borders.getItem('EdgeRight').style = 'Continuous';
                }
              }
              executedDetails.push(`تنسيق النطاق ${action.range}`);
            }
            break;
          }

          // 4. ضبط العرض التلقائي للأعمدة (منع ظهور علامات ###)
          case "AUTOFIT_COLUMNS": {
            const range = action.range ? targetSheet.getRange(action.range) : targetSheet.getUsedRange();
            range.format.autofitColumns();
            executedDetails.push(`ضبط العرض التلقائي للأعمدة`);
            break;
          }

          // 5. إنشاء جدول إكسيل رسمي (Excel Table)
          case "CREATE_TABLE": {
            if (action.range) {
              const range = targetSheet.getRange(action.range);
              const tableName = action.tableName || `Table_${Date.now()}`;
              const table = targetSheet.tables.add(range, action.hasHeaders !== false);
              table.name = tableName.replace(/\s+/g, '_');
              if (action.style) {
                table.style = action.style; // e.g. "TableStyleMedium2"
              }
              executedDetails.push(`إنشاء جدول إكسيل رسمي في ${action.range}`);
            }
            break;
          }

          // 6. إضافة رسم بياني احترافي (Chart)
          case "ADD_CHART": {
            if (action.dataRange) {
              const dataRange = targetSheet.getRange(action.dataRange);
              const chartType = action.chartType || "ColumnClustered";
              const chart = targetSheet.charts.add(chartType, dataRange, "Auto");
              chart.title.text = action.title || "رسم بياني توضيحي";
              chart.legend.position = "Top";
              
              if (action.targetCell) {
                chart.setPosition(action.targetCell);
              }
              executedDetails.push(`إنشاء رسم بياني (${chartType})`);
            }
            break;
          }

          // 7. تجميد الصف العلوي (Freeze Panes)
          case "FREEZE_PANES": {
            targetSheet.freezePanes.freezeRows(action.rows || 1);
            executedDetails.push(`تجميد الصفوف العلوية`);
            break;
          }

          // 8. التنسيق والترميز الشرطي الذكي (Conditional Formatting & Heatmaps)
          case "ADD_CONDITIONAL_FORMAT": {
            if (action.range) {
              const range = targetSheet.getRange(action.range);
              const formatType = action.formatType || "colorScale";

              if (formatType === "colorScale" || formatType === "heatmap") {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
                cf.colorScale.criteria = {
                  minimum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.lowestValue, color: "#FCA5A5" },
                  midpoint: { formula: null, type: Excel.ConditionalFormatColorCriterionType.percentile, color: "#FEF08A" },
                  maximum: { formula: null, type: Excel.ConditionalFormatColorCriterionType.highestValue, color: "#86EFAC" }
                };
                executedDetails.push(`تطبيق تدرج لوني حراري (Heatmap) على النطاق ${action.range}`);
              } else if (formatType === "highlightNegative") {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.cellValue);
                cf.cellValue.format.font.color = "#B91C1C";
                cf.cellValue.format.fill.color = "#FEE2E2";
                cf.cellValue.rule = { formula1: "0", operator: "LessThan" };
                executedDetails.push(`تمييز القيم السالبة بالأحمر في النطاق ${action.range}`);
              } else if (formatType === "dataBars") {
                const cf = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
                cf.dataBar.barColor = "#3B82F6";
                executedDetails.push(`إضافة أشرطة بيانات بصرية في النطاق ${action.range}`);
              }
            }
            break;
          }

          // 9. تهيئة الصفحة للطباعة والتصدير كـ PDF جاهز (Page Layout Setup)
          case "SET_PAGE_LAYOUT": {
            try {
              const orientation = (action.orientation && action.orientation.toLowerCase() === "portrait") 
                ? Excel.PageOrientation.portrait 
                : Excel.PageOrientation.landscape;
              targetSheet.pageLayout.orientation = orientation;
              targetSheet.pageLayout.paperSize = Excel.PaperType.a4;
              targetSheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 };
              
              if (action.range) {
                targetSheet.pageLayout.setPrintArea(action.range);
              } else {
                const used = targetSheet.getUsedRange();
                targetSheet.pageLayout.setPrintArea(used);
              }
              executedDetails.push(`تهيئة الصفحة للطباعة (A4 أفقي، ملاءمة العرض في صفحة واحدة)`);
            } catch (pageErr) {
              console.warn("تنبيه في إعدادات الطباعة:", pageErr);
            }
            break;
          }

          // 10. قواعد التحقق من صحة البيانات والقوائم المنسدلة داخل خلايا الإكسيل (Data Validation & In-Cell Dropdowns)
          case "SET_DATA_VALIDATION": {
            if (action.range) {
              const range = targetSheet.getRange(action.range);
              try {
                if (action.validationType === "list" && action.values) {
                  const valuesList = Array.isArray(action.values) ? action.values.join(",") : String(action.values);
                  range.dataValidation.rule = {
                    list: {
                      inCellDropDown: true,
                      values: valuesList
                    }
                  };
                  range.dataValidation.prompt = {
                    showPrompt: true,
                    title: action.promptTitle || "اختر من القائمة",
                    message: action.promptMessage || "يرجى تحديد إحدى القيم المعتمدة"
                  };
                  range.dataValidation.errorAlert = {
                    showAlert: true,
                    style: "Stop",
                    title: "قيمة غير مسموح بها",
                    message: "يرجى اختيار قيمة من القائمة المنسدلة فقط."
                  };
                  executedDetails.push(`إنشاء قائمة منسدلة داخل الخلايا في ${action.range} بالقيم [${valuesList}]`);
                } else if (action.validationType === "wholeNumber" || action.validationType === "decimal") {
                  const valType = action.validationType === "wholeNumber" ? "wholeNumber" : "decimal";
                  const ruleObj = {};
                  ruleObj[valType] = {
                    operator: action.operator || "GreaterThan",
                    formula1: action.formula1 !== undefined ? String(action.formula1) : "0"
                  };
                  if (action.formula2 !== undefined) {
                    ruleObj[valType].formula2 = String(action.formula2);
                  }
                  range.dataValidation.rule = ruleObj;
                  range.dataValidation.errorAlert = {
                    showAlert: true,
                    style: "Stop",
                    title: "رقم غير صالح",
                    message: "القيمة المدخلة تخالف شروط التحقق من الأرقام."
                  };
                  executedDetails.push(`تطبيق قيود التحقق الرقمي على ${action.range}`);
                }
              } catch (valErr) {
                console.warn("تنبيه في قواعد التحقق من صحة البيانات:", valErr);
              }
            }
            break;
          }
        }
      }

      await context.sync();

      if (createdSheetNames.length > 0) {
        this.lastUndoAction = {
          type: "CREATED_SHEETS",
          createdSheetNames: createdSheetNames,
          previousSheetName: initialActiveSheetName
        };
      }
    });

    return {
      success: true,
      count: actionsList.length,
      details: executedDetails,
      canUndo: !!this.lastUndoAction
    };
  },

  // بيانات آخر عملية قابلة للتراجع
  lastUndoAction: null,

  /**
   * التراجع عن آخر عملية قام بها الوكيل
   */
  async undoLastAction() {
    if (!this.lastUndoAction) {
      throw new Error("لا توجد عملية سابقة للتراجع عنها حالياً.");
    }

    if (typeof Excel === "undefined") {
      this.lastUndoAction = null;
      return "تم التراجع في وضع المعاينة.";
    }

    const undoData = this.lastUndoAction;
    let message = "";

    await Excel.run(async (context) => {
      if (undoData.type === "CREATED_SHEETS") {
        for (const sName of undoData.createdSheetNames) {
          const sheet = context.workbook.worksheets.getItemOrNullObject(sName);
          await context.sync();
          if (!sheet.isNullObject) {
            sheet.delete();
          }
        }
        if (undoData.previousSheetName) {
          const prev = context.workbook.worksheets.getItemOrNullObject(undoData.previousSheetName);
          await context.sync();
          if (!prev.isNullObject) {
            prev.activate();
          }
        }
        await context.sync();
        message = `تم التراجع بنجاح: تم حذف الورقة المنشأة (${undoData.createdSheetNames.join(", ")}) والعودة للورقة السابقة.`;
      } else if (undoData.type === "RANGE_MODIFICATIONS") {
        for (const mod of undoData.modifications) {
          const sheet = context.workbook.worksheets.getItemOrNullObject(mod.sheetName);
          await context.sync();
          if (!sheet.isNullObject && mod.range) {
            const range = sheet.getRange(mod.range);
            if (mod.hasFormulas && mod.formulas) {
              range.formulas = mod.formulas;
            } else if (mod.values) {
              range.values = mod.values;
            }
          }
        }
        await context.sync();
        message = `تم التراجع بنجاح واسترجاع بيانات وقيم الشيت الأصلية قبل التعديل.`;
      }
    });

    this.lastUndoAction = null;
    return message;
  }
};

window.ExcelEngine = ExcelEngine;
