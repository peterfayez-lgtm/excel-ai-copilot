/**
 * Excel Actions Engine - Native Office.js Operations
 * مسؤول عن تنفيذ الأوامر والتعديلات مباشرة على شيت الإكسيل
 */

const ExcelEngine = {
  
  /**
   * فحص وقراءة سياق الشيت والخلايا المحددة حالياً
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
          columnCount: 3
        });
        return;
      }

      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const selection = context.workbook.getSelectedRange();
        
        sheet.load("name");
        selection.load(["address", "values", "formulas", "rowCount", "columnCount"]);
        
        await context.sync();
        
        resolve({
          sheetName: sheet.name,
          address: selection.address.split("!")[1] || selection.address,
          values: selection.values,
          formulas: selection.formulas,
          rowCount: selection.rowCount,
          columnCount: selection.columnCount
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

    await Excel.run(async (context) => {
      let currentActive = context.workbook.worksheets.getActiveWorksheet();

      for (const action of actionsList) {
        // 1. التعامل مع إنشاء أو تفعيل ورقة عمل جديدة بأمان تام
        if (action.type === "CREATE_SHEET") {
          if (action.sheetName) {
            const existingSheet = context.workbook.worksheets.getItemOrNullObject(action.sheetName);
            await context.sync();

            let targetSheet;
            if (existingSheet.isNullObject) {
              targetSheet = context.workbook.worksheets.add(action.sheetName);
            } else {
              targetSheet = existingSheet;
            }
            targetSheet.activate();
            currentActive = targetSheet;
            await context.sync();
            executedDetails.push(`إنشاء/تفعيل الورقة: ${action.sheetName}`);
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
              range.formulas = action.formulas;
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
        }
      }

      await context.sync();
    });

    return {
      success: true,
      count: actionsList.length,
      details: executedDetails
    };
  }
};

window.ExcelEngine = ExcelEngine;
