/**
 * WPS 智能表格 API 1.0 通用工具函数库
 * 基于 https://airsheet.wps.cn/docs/api/excel/workbook/overview.html
 */

// ==================== HTTP API 调用入口 ====================

/**
 * HTTP API 调用的主入口函数
 * 当通过 Python HTTP 请求调用时，会自动执行此函数
 *
 * 重要：WPS AirScript 需要脚本最后一个表达式作为返回值
 */

// 定义全局结果变量
var globalResult = [];

// 检查是否是 HTTP API 调用（存在 Context 对象）
if (typeof Context !== "undefined" && Context.argv) {
  try {
    console.log("接收到 HTTP API 调用");
    console.log("Context:", JSON.stringify(Context));

    var argv = Context.argv;
    var sheetName = Context.active_sheet;

    // 如果有 items 数据，使用 setRangeValues 批量写入
    if (argv.items && Array.isArray(argv.items)) {
      try {
        const data = argv.items;
        const rows = data.length;
        const cols = data[0] ? data[0].length : 0;

        if (rows > 0 && cols > 0) {
          // 计算范围 (从 A1 开始)
          const endCol = columnNumberToLetter(cols);
          const address = `A1:${endCol}${rows}`;
          setRangeValues(address, data, sheetName);

          globalResult.push({
            success: true,
            message: "数据写入成功",
            rowsWritten: rows,
            range: address,
          });
        } else {
          globalResult.push({
            success: false,
            message: "数据为空",
          });
        }
        console.log("返回结果:", JSON.stringify(globalResult));
      } catch (error) {
        globalResult.push({
          success: false,
          error: error.message,
        });
      }
    }
    // 如果有 function 参数，执行指定函数
    else if (argv.function) {
      globalResult = executeFunction(argv.function, argv, sheetName);
      console.log("返回结果:", JSON.stringify(globalResult));
    }
    // 未指定操作
    else {
      globalResult.push({
        success: false,
        message: "未指定操作",
      });
    }
  } catch (error) {
    console.error("HTTP API 调用出错:", error.message);
    globalResult = [];
    globalResult.push({
      success: false,
      error: error.message,
    });
  }
}

globalResult;

// ==================== HTTP API 辅助函数 ====================

/**
 * 执行指定函数（HTTP API 专用）
 * @param {string} functionName - 函数名
 * @param {Object} params - 参数对象
 * @param {string} sheetName - 工作表名称
 * @returns {Array} 执行结果数组
 */
function executeFunction(functionName, params, sheetName) {
  const result = [];
  console.log("执行函数:", functionName);
  console.log("目标工作表:", sheetName || "当前工作表");

  try {
    switch (functionName) {
      case "getCellValue":
        result.push({
          success: true,
          value: getCellValue(params.address, sheetName),
        });
        break;

      case "setCellValue":
        setCellValue(params.address, params.value, sheetName);
        result.push({ success: true, message: "设置成功" });
        break;

      case "getRangeValues":
        result.push({
          success: true,
          values: getRangeValues(params.address, sheetName),
        });
        break;

      case "setRangeValues":
        setRangeValues(params.address, params.values, sheetName);
        result.push({ success: true, message: "设置成功" });
        break;

      case "setCellFont":
        setCellFont(params.address, params.fontOptions, sheetName);
        result.push({ success: true, message: "字体设置成功" });
        break;

      case "setCellBackgroundColor":
        setCellBackgroundColor(params.address, params.color, sheetName);
        result.push({ success: true, message: "背景色设置成功" });
        break;

      case "setCellAlignment":
        setCellAlignment(params.address, params.alignOptions, sheetName);
        result.push({ success: true, message: "对齐方式设置成功" });
        break;

      case "setCellBorder":
        setCellBorder(params.address, params.borderOptions, sheetName);
        result.push({ success: true, message: "边框设置成功" });
        break;

      case "mergeCells":
        mergeCells(params.address, sheetName);
        result.push({ success: true, message: "合并成功" });
        break;

      case "autoFitColumns":
        autoFitColumns(params.address, sheetName);
        result.push({ success: true, message: "列宽调整成功" });
        break;

      case "insertRows":
        insertRows(params.rowIndex, params.count, sheetName);
        result.push({ success: true, message: "插入行成功" });
        break;

      case "setRowHeight":
        setRowHeight(params.rowIndex, params.height, sheetName);
        result.push({ success: true, message: "行高设置成功" });
        break;

      case "setColumnWidth":
        setColumnWidth(params.columnIndex, params.width, sheetName);
        result.push({ success: true, message: "列宽设置成功" });
        break;

      case "findCell":
        const cells = findCell(
          params.searchText,
          params.searchRange,
          sheetName
        );
        result.push({
          success: true,
          found: cells.length > 0,
          cells: cells,
        });
        break;

      case "replaceInRangeWithCount":
        const count = replaceInRangeWithCount(
          params.searchText,
          params.replaceText,
          params.searchRange,
          sheetName
        );
        result.push({ success: true, count: count });
        break;

      case "sortRange":
        sortRange(params.address, params.sortOptions, sheetName);
        result.push({ success: true, message: "排序成功" });
        break;

      case "copyPasteRange":
        copyPasteRange(
          params.sourceAddress,
          params.targetAddress,
          sheetName,
          sheetName
        );
        result.push({ success: true, message: "复制粘贴成功" });
        break;

      case "clearRange":
        clearRange(params.address, sheetName);
        result.push({ success: true, message: "清除成功" });
        break;

      case "clearRangeContents":
        clearRangeContents(params.address, sheetName);
        result.push({ success: true, message: "清除内容成功" });
        break;

      case "getCellFormula":
        result.push({
          success: true,
          formula: getCellFormula(params.address, sheetName),
        });
        break;

      case "setCellFormula":
        setCellFormula(params.address, params.formula, sheetName);
        result.push({ success: true, message: "设置公式成功" });
        break;

      case "setCellNumberFormat":
        setCellNumberFormat(params.address, params.format, sheetName);
        result.push({ success: true, message: "设置数字格式成功" });
        break;

      case "unmergeCells":
        unmergeCells(params.address, sheetName);
        result.push({ success: true, message: "取消合并成功" });
        break;

      case "deleteRows":
        deleteRows(params.rowIndex, params.count, sheetName);
        result.push({ success: true, message: "删除行成功" });
        break;

      case "insertColumns":
        insertColumns(params.columnIndex, params.count, sheetName);
        result.push({ success: true, message: "插入列成功" });
        break;

      case "deleteColumns":
        deleteColumns(params.columnIndex, params.count, sheetName);
        result.push({ success: true, message: "删除列成功" });
        break;

      case "findAllCells":
        const allCells = findAllCells(
          params.searchText,
          params.searchRange,
          sheetName
        );
        // 转换为标准格式
        const cellsInfo = allCells.map((cell) => ({
          address: cell.Address,
          value: cell.Value,
          row: cell.Row,
          column: cell.Column,
        }));
        result.push({
          success: true,
          cells: cellsInfo,
          count: cellsInfo.length,
        });
        break;

      case "copyRange":
        copyRange(params.sourceAddress, sheetName);
        result.push({ success: true, message: "复制成功" });
        break;

      case "pasteToRange":
        pasteToRange(params.targetAddress, sheetName);
        result.push({ success: true, message: "粘贴成功" });
        break;

      case "getUsedRangeData":
        result.push({
          success: true,
          data: getUsedRangeData(sheetName),
        });
        break;

      case "addWorksheet":
        const newSheet = addWorksheet(params.sheetName);
        result.push({
          success: true,
          message: "添加工作表成功",
          sheetName: newSheet.Name,
        });
        break;

      case "deleteWorksheet":
        deleteWorksheet(params.sheetIdentifier);
        result.push({ success: true, message: "删除工作表成功" });
        break;

      case "worksheetExists":
        result.push({
          success: true,
          exists: worksheetExists(params.sheetName),
        });
        break;

      case "getWorksheetCount":
        result.push({ success: true, count: getWorksheetCount() });
        break;

      case "getWorkbookName":
        result.push({ success: true, sheets: getWorkbookName() });
        break;

      default:
        result.push({
          success: false,
          message: "未知函数: " + functionName,
        });
    }
  } catch (error) {
    result.push({
      success: false,
      error: error.message,
    });
  }
  return result;
}

// ==================== 工作簿 (Workbook) 相关操作 ====================

/**
 * 获取当前活动的工作簿对象
 * @returns {Object} 工作簿对象
 */
function getActiveWorkbook() {
  return Application.ActiveWorkbook;
}

/**
 * 获取工作簿名称
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {string} 工作簿名称
 */
function getWorkbookName(workbook) {
  try {
    const wb = workbook || Application.ActiveWorkbook;

    // WPS AirScript 可能不支持获取工作簿名称
    // 返回所有工作表名称作为替代
    if (wb && wb.Sheets) {
      const sheets = wb.Sheets;
      const sheetNames = [];

      for (let i = 1; i <= sheets.Count; i++) {
        sheetNames.push(sheets.Item(i).Name);
      }

      return sheetNames;
    }

    return [];
  } catch (error) {
    console.error("getWorkbookName 错误:", error.message);
    return [];
  }
}

/**
 * 保存工作簿
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 */
function saveWorkbook(workbook) {
  const wb = workbook || getActiveWorkbook();
  wb.Save();
}

/**
 * 关闭工作簿
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @param {boolean} saveChanges - 是否保存更改，默认 false
 */
function closeWorkbook(workbook, saveChanges = false) {
  const wb = workbook || getActiveWorkbook();
  wb.Close(saveChanges);
}

// ==================== 工作表 (Worksheet) 相关操作 ====================

/**
 * 获取当前活动的工作表对象
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {Object} 工作表对象
 */
function getActiveWorksheet(workbook) {
  const wb = workbook || getActiveWorkbook();
  return wb.ActiveSheet;
}

/**
 * 根据名称获取工作表（支持模糊匹配）
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Object} 工作表对象
 */
function getWorksheetByName(sheetName) {
  // 如果没有传入工作表名称，返回当前活动工作表
  if (!sheetName) {
    return Application.ActiveSheet;
  }

  const workbook = Application.ActiveWorkbook;
  const sheetCount = workbook.Sheets.Count;

  // 精确匹配
  for (let i = 1; i <= sheetCount; i++) {
    const sheet = workbook.Sheets(i);
    if (sheet.Name === sheetName) {
      return sheet;
    }
  }

  // 模糊匹配（包含）
  for (let i = 1; i <= sheetCount; i++) {
    const sheet = workbook.Sheets(i);
    if (sheet.Name.includes(sheetName)) {
      console.log("找到匹配的工作表:", sheet.Name);
      return sheet;
    }
  }

  // 未找到，返回 null
  console.error("未找到工作表:", sheetName);
  return null;
}

/**
 * 根据索引获取工作表
 * @param {number} index - 工作表索引（从1开始）
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {Object} 工作表对象
 */
function getWorksheetByIndex(index, workbook) {
  const wb = workbook || getActiveWorkbook();
  return wb.Worksheets.Item(index);
}

/**
 * 检查工作表是否存在（支持模糊匹配）
 * @param {string} sheetName - 工作表名称
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {boolean} 是否存在
 */
function worksheetExists(sheetName, workbook) {
  const wb = workbook || getActiveWorkbook();
  const sheetCount = wb.Sheets.Count;

  // 精确匹配
  for (let i = 1; i <= sheetCount; i++) {
    const sheet = wb.Sheets(i);
    if (sheet.Name === sheetName) {
      return true;
    }
  }

  // 模糊匹配（包含）
  for (let i = 1; i <= sheetCount; i++) {
    const sheet = wb.Sheets(i);
    if (sheet.Name.includes(sheetName)) {
      return true;
    }
  }

  return false;
}

/**
 * 添加新工作表
 * @param {string} sheetName - 工作表名称（可选）
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {Object} 新创建的工作表对象
 */
function addWorksheet(sheetName, workbook) {
  const wb = workbook || getActiveWorkbook();
  const newSheet = wb.Worksheets.Add();
  if (sheetName) {
    newSheet.Name = sheetName;
  }
  return newSheet;
}

/**
 * 添加新工作表（如果已存在则返回现有工作表）
 * @param {string} sheetName - 工作表名称
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {Object} 工作表对象
 */
function addWorksheetIfNotExists(sheetName, workbook) {
  const wb = workbook || getActiveWorkbook();
  if (worksheetExists(sheetName, wb)) {
    return getWorksheetByName(sheetName, wb);
  }
  return addWorksheet(sheetName, wb);
}

/**
 * 删除工作表
 * @param {string|number} sheetIdentifier - 工作表名称或索引
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 */
function deleteWorksheet(sheetIdentifier, workbook) {
  const wb = workbook || getActiveWorkbook();
  const sheet =
    typeof sheetIdentifier === "string"
      ? getWorksheetByName(sheetIdentifier, wb)
      : getWorksheetByIndex(sheetIdentifier, wb);
  sheet.Delete();
}

/**
 * 获取工作表数量
 * @param {Object} workbook - 工作簿对象，不传则使用当前活动工作簿
 * @returns {number} 工作表数量
 */
function getWorksheetCount(workbook) {
  const wb = workbook || getActiveWorkbook();
  return wb.Worksheets.Count;
}

// ==================== 单元格 (Range) 相关操作 ====================

/**
 * 获取单元格区域对象
 * @param {string} address - 单元格地址，如 "A1" 或 "A1:B10"
 * @param {string|Object} worksheetOrName - 工作表对象或工作表名称，不传则使用当前活动工作表
 * @returns {Object} 单元格区域对象
 */
function getRange(address, worksheetOrName) {
  let ws;

  if (!worksheetOrName) {
    // 没有传入参数，使用当前活动工作表
    ws = Application.ActiveSheet;
  } else if (typeof worksheetOrName === "string") {
    // 传入的是工作表名称
    ws = getWorksheetByName(worksheetOrName);
    if (!ws) {
      throw new Error("未找到工作表: " + worksheetOrName);
    }
  } else {
    // 传入的是工作表对象
    ws = worksheetOrName;
  }

  return ws.Range(address);
}

/**
 * 获取单元格的值
 * @param {string} address - 单元格地址，如 "A1"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {*} 单元格的值
 */
function getCellValue(address, sheetName) {
  const range = getRange(address, sheetName);
  return range.Value;
}

/**
 * 设置单元格的值
 * @param {string} address - 单元格地址，如 "A1"
 * @param {*} value - 要设置的值
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellValue(address, value, sheetName) {
  const range = getRange(address, sheetName);
  range.Value = value;
}

/**
 * 获取单元格区域的值（二维数组）
 * @param {string} address - 单元格区域地址，如 "A1:B10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 二维数组
 */
function getRangeValues(address, sheetName) {
  const range = getRange(address, sheetName);
  return range.Value;
}

/**
 * 设置单元格区域的值（二维数组）
 * @param {string} address - 单元格区域地址，如 "A1:B10"
 * @param {Array} values - 二维数组
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setRangeValues(address, values, sheetName) {
  const range = getRange(address, sheetName);
  range.Value = values;
}

/**
 * 清除单元格内容
 * @param {string} address - 单元格地址，如 "A1" 或 "A1:B10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function clearRange(address, sheetName) {
  const range = getRange(address, sheetName);
  range.Clear();
}

/**
 * 清除单元格内容（保留格式）
 * @param {string} address - 单元格地址，如 "A1" 或 "A1:B10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function clearRangeContents(address, sheetName) {
  const range = getRange(address, sheetName);
  range.ClearContents();
}

/**
 * 获取单元格公式
 * @param {string} address - 单元格地址，如 "A1"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {string} 单元格公式
 */
function getCellFormula(address, sheetName) {
  const range = getRange(address, sheetName);
  return range.Formula;
}

/**
 * 设置单元格公式
 * @param {string} address - 单元格地址，如 "A1"
 * @param {string} formula - 公式字符串，如 "=SUM(A1:A10)"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellFormula(address, formula, sheetName) {
  const range = getRange(address, sheetName);
  range.Formula = formula;
}

// ==================== 单元格格式化操作 ====================

/**
 * 设置单元格字体样式
 * @param {string} address - 单元格地址
 * @param {Object} fontOptions - 字体选项 { name, size, bold, italic, color }
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellFont(address, fontOptions, sheetName) {
  const range = getRange(address, sheetName);
  const font = range.Font;

  if (fontOptions.name) font.Name = fontOptions.name;
  if (fontOptions.size) font.Size = fontOptions.size;
  if (fontOptions.bold !== undefined) font.Bold = fontOptions.bold;
  if (fontOptions.italic !== undefined) font.Italic = fontOptions.italic;
  if (fontOptions.color) font.Color = fontOptions.color;
}

/**
 * 设置单元格背景色
 * @param {string} address - 单元格地址
 * @param {number} color - 颜色值（RGB）
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellBackgroundColor(address, color, sheetName) {
  const range = getRange(address, sheetName);
  range.Interior.Color = color;
}

/**
 * 设置单元格对齐方式
 * @param {string} address - 单元格地址
 * @param {Object} alignOptions - 对齐选项 { horizontal, vertical }
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellAlignment(address, alignOptions, sheetName) {
  const range = getRange(address, sheetName);

  if (alignOptions.horizontal) {
    range.HorizontalAlignment = alignOptions.horizontal;
  }
  if (alignOptions.vertical) {
    range.VerticalAlignment = alignOptions.vertical;
  }
}

/**
 * 设置单元格边框
 * @param {string} address - 单元格地址
 * @param {Object} borderOptions - 边框选项 { lineStyle, weight, color }
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellBorder(address, borderOptions, sheetName) {
  const range = getRange(address, sheetName);
  const borders = range.Borders;

  if (borderOptions.lineStyle) borders.LineStyle = borderOptions.lineStyle;
  if (borderOptions.weight) borders.Weight = borderOptions.weight;
  if (borderOptions.color) borders.Color = borderOptions.color;
}

/**
 * 设置单元格数字格式
 * @param {string} address - 单元格地址
 * @param {string} format - 数字格式，如 "0.00", "#,##0", "yyyy-mm-dd"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setCellNumberFormat(address, format, sheetName) {
  const range = getRange(address, sheetName);
  range.NumberFormat = format;
}

// ==================== 行列操作 ====================

/**
 * 插入行
 * @param {number} rowIndex - 行索引（从1开始）
 * @param {number} count - 插入行数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function insertRows(rowIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  const range = ws.Rows(rowIndex);
  for (let i = 0; i < count; i++) {
    range.Insert();
  }
}

/**
 * 删除行
 * @param {number} rowIndex - 行索引（从1开始）
 * @param {number} count - 删除行数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function deleteRows(rowIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  for (let i = 0; i < count; i++) {
    const range = ws.Rows(rowIndex);
    range.Delete();
  }
}

/**
 * 插入列
 * @param {number} columnIndex - 列索引（从1开始）
 * @param {number} count - 插入列数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function insertColumns(columnIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  const range = ws.Columns(columnIndex);
  for (let i = 0; i < count; i++) {
    range.Insert();
  }
}

/**
 * 删除列
 * @param {number} columnIndex - 列索引（从1开始）
 * @param {number} count - 删除列数，默认1
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function deleteColumns(columnIndex, count = 1, sheetName) {
  const ws = getWorksheetByName(sheetName);
  for (let i = 0; i < count; i++) {
    const range = ws.Columns(columnIndex);
    range.Delete();
  }
}

/**
 * 设置行高
 * @param {number} rowIndex - 行索引（从1开始）
 * @param {number} height - 行高
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setRowHeight(rowIndex, height, sheetName) {
  const ws = getWorksheetByName(sheetName);
  ws.Rows(rowIndex).RowHeight = height;
}

/**
 * 设置列宽
 * @param {number} columnIndex - 列索引（从1开始）
 * @param {number} width - 列宽
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function setColumnWidth(columnIndex, width, sheetName) {
  const ws = getWorksheetByName(sheetName);
  ws.Columns(columnIndex).ColumnWidth = width;
}

/**
 * 自动调整列宽
 * @param {string} address - 单元格区域地址，如 "A:A" 或 "A1:C10"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function autoFitColumns(address, sheetName) {
  const range = getRange(address, sheetName);
  range.Columns.AutoFit();
}

// ==================== 查找和筛选操作 ====================

/**
 * 查找单元格（返回所有匹配项）
 * @param {string} searchText - 要查找的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 找到的所有单元格信息数组 [{address, value, row, column}]，未找到返回空数组
 */
function findCell(searchText, searchRange, sheetName) {
  const range = getRange(searchRange, sheetName);
  const result = [];

  // 查找第一个匹配项
  const firstCell = range.Find(searchText);

  if (!firstCell) {
    return result;
  }

  // 记录第一个单元格的行列，用于判断是否循环回到起点
  const firstRow = firstCell.Row;
  const firstCol = firstCell.Column;
  let currentCell = firstCell;
  let count = 0;
  const maxIterations = 10000; // 防止无限循环

  // 循环查找所有匹配项
  do {
    result.push({
      address: currentCell.Address,
      value: currentCell.Value,
      row: currentCell.Row,
      column: currentCell.Column,
    });

    // 查找下一个匹配项
    currentCell = range.FindNext(currentCell);
    count++;

    // 安全检查：防止无限循环
    if (count > maxIterations) {
      console.error("查找循环次数超过限制，可能存在问题");
      break;
    }

    // 如果找不到或者回到第一个单元格（通过行列判断），则退出循环
  } while (
    currentCell &&
    !(currentCell.Row === firstRow && currentCell.Column === firstCol)
  );

  return result;
}

/**
 * 查找所有匹配的单元格
 * @param {string} searchText - 要查找的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 找到的单元格对象数组
 */
function findAllCells(searchText, searchRange, sheetName) {
  const range = getRange(searchRange, sheetName);
  const results = [];
  const firstCell = range.Find(searchText);

  if (!firstCell) return results;

  // 记录第一个单元格的行列
  const firstRow = firstCell.Row;
  const firstCol = firstCell.Column;
  let currentCell = firstCell;
  let count = 0;
  const maxIterations = 10000; // 防止无限循环

  do {
    results.push(currentCell);
    currentCell = range.FindNext(currentCell);
    count++;

    // 安全检查：防止无限循环
    if (count > maxIterations) {
      console.error("查找循环次数超过限制，可能存在问题");
      break;
    }
  } while (
    currentCell &&
    !(currentCell.Row === firstRow && currentCell.Column === firstCol)
  );

  return results;
}

/**
 * 替换单元格内容
 * @param {string} searchText - 要查找的文本
 * @param {string} replaceText - 替换的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {boolean} 是否成功替换（true=成功，false=未找到）
 */
function replaceInRange(searchText, replaceText, searchRange, sheetName) {
  const range = getRange(searchRange, sheetName);
  return range.Replace(searchText, replaceText);
}

/**
 * 替换单元格内容并返回替换数量
 * @param {string} searchText - 要查找的文本
 * @param {string} replaceText - 替换的文本
 * @param {string} searchRange - 查找范围，如 "A1:Z100"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {number} 替换的数量
 */
function replaceInRangeWithCount(
  searchText,
  replaceText,
  searchRange,
  sheetName
) {
  // 先查找所有匹配项（用于计数）
  const cells = findAllCells(searchText, searchRange, sheetName);
  const count = cells.length;

  // 如果找到匹配项，执行替换
  if (count > 0) {
    const range = getRange(searchRange, sheetName);
    range.Replace(searchText, replaceText);
  }

  return count;
}

// ==================== 排序操作 ====================

/**
 * 对区域进行排序
 * @param {string} address - 要排序的区域地址
 * @param {Object} sortOptions - 排序选项 { key, order, hasHeader }
 *   - key: 排序关键列地址，如 "A1"
 *   - order: 排序顺序，1=升序，2=降序
 *   - hasHeader: 是否包含标题行，默认 false
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function sortRange(address, sortOptions, sheetName) {
  const range = getRange(address, sheetName);
  const key = getRange(sortOptions.key, sheetName);
  const order = sortOptions.order || 1;
  const header = sortOptions.hasHeader ? 1 : 2;

  range.Sort(key, order, null, null, null, null, null, header);
}

// ==================== 复制粘贴操作 ====================

/**
 * 复制单元格区域
 * @param {string} sourceAddress - 源区域地址
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function copyRange(sourceAddress, sheetName) {
  const range = getRange(sourceAddress, sheetName);
  range.Copy();
}

/**
 * 粘贴到指定位置
 * @param {string} targetAddress - 目标区域地址
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function pasteToRange(targetAddress, sheetName) {
  const range = getRange(targetAddress, sheetName);
  range.Select();
  const ws = getWorksheetByName(sheetName);
  ws.Paste();
}

/**
 * 复制并粘贴单元格区域
 * @param {string} sourceAddress - 源区域地址
 * @param {string} targetAddress - 目标区域地址
 * @param {Object} sourceWorksheet - 源工作表对象
 * @param {Object} targetWorksheet - 目标工作表对象
 */
function copyPasteRange(
  sourceAddress,
  targetAddress,
  sourceWorksheet,
  targetWorksheet
) {
  const sourceRange = getRange(sourceAddress, sourceWorksheet);
  const targetRange = getRange(targetAddress, targetWorksheet);
  sourceRange.Copy(targetRange);
}

// ==================== 合并单元格操作 ====================

/**
 * 合并单元格
 * @param {string} address - 要合并的区域地址，如 "A1:B2"
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function mergeCells(address, sheetName) {
  const range = getRange(address, sheetName);
  range.Merge();
}

/**
 * 取消合并单元格
 * @param {string} address - 要取消合并的区域地址
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 */
function unmergeCells(address, sheetName) {
  const range = getRange(address, sheetName);
  range.UnMerge();
}

// ==================== 批量数据操作 ====================

/**
 * 获取已使用区域的数据
 * @param {string} sheetName - 工作表名称，不传则使用当前活动工作表
 * @returns {Array} 二维数组数据
 */
function getUsedRangeData(sheetName) {
  const ws = getWorksheetByName(sheetName);
  const usedRange = ws.UsedRange;
  return usedRange.Value;
}

// ==================== 工具函数 ====================

/**
 * 列字母转数字索引
 * @param {string} column - 列字母，如 "A", "AB"
 * @returns {number} 列索引（从1开始）
 */
function columnLetterToNumber(column) {
  let result = 0;
  for (let i = 0; i < column.length; i++) {
    result = result * 26 + (column.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * 列数字索引转字母
 * @param {number} columnNumber - 列索引（从1开始）
 * @returns {string} 列字母
 */
function columnNumberToLetter(columnNumber) {
  let letter = "";
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

/**
 * RGB 颜色转换为 Excel 颜色值
 * @param {number} r - 红色值 (0-255)
 * @param {number} g - 绿色值 (0-255)
 * @param {number} b - 蓝色值 (0-255)
 * @returns {number} Excel 颜色值
 */
function rgbToExcelColor(r, g, b) {
  return r + g * 256 + b * 256 * 256;
}

// ==================== 返回结果 ====================
return globalResult;
