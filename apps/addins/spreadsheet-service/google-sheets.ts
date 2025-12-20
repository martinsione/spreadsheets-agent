import type { InferToolInput, InferToolOutput } from "ai";
import type { tools } from "@/server/ai/tools";
import type { Sheet, SpreadsheetService } from "@/spreadsheet-service";

// `initializeApp` and `onOpen` are necessary for the app to be loaded in the sidebar
export function initializeApp() {
  const html =
    HtmlService.createHtmlOutputFromFile("index").setTitle("OpenSheets");
  SpreadsheetApp.getUi().showSidebar(html);
}

export function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("OpenSheets")
    .addItem("App", "initializeApp")
    .addToUi();
}

export async function getSheets(): Promise<Sheet[]> {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  return sheets.map((sheet) => {
    const dataRange = sheet.getDataRange();
    return {
      id: sheet.getSheetId(),
      name: sheet.getName(),
      maxRows: dataRange.getNumRows(),
      maxColumns: dataRange.getNumColumns(),
    };
  });
}

export async function getCellRanges(
  input: InferToolInput<typeof tools.getCellRanges>,
): Promise<InferToolOutput<typeof tools.getCellRanges>> {
  const { sheetId, ranges, includeStyles = true, cellLimit = 10000 } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  const dataRange = sheet.getDataRange();
  const dimension = dataRange.getA1Notation();

  const cells: Record<string, unknown> = {};
  const styles: Record<string, Record<string, unknown>> = {};
  const borders: Record<string, string | Record<string, string>> = {};
  let totalCellCount = 0;
  let hasMore = false;
  const nextRanges: string[] = [];

  for (let i = 0; i < ranges.length; i++) {
    const rangeAddress = ranges[i];
    if (!rangeAddress) continue;

    if (totalCellCount >= cellLimit) {
      nextRanges.push(...ranges.slice(i));
      hasMore = true;
      break;
    }

    const range = sheet.getRange(rangeAddress);
    const values = range.getValues();
    const formulas = range.getFormulas();
    const notes = range.getNotes();
    const startRow = range.getRow();
    const startCol = range.getColumn();
    const rowCount = range.getNumRows();
    const colCount = range.getNumColumns();

    for (let row = 0; row < rowCount; row++) {
      if (totalCellCount >= cellLimit) {
        hasMore = true;
        const remainingRows = rowCount - row;
        if (remainingRows > 0) {
          const startCellRow = startRow + row;
          const endCellRow = startRow + rowCount - 1;
          const startColLetter = columnToLetter(startCol - 1);
          const endColLetter = columnToLetter(startCol + colCount - 2);
          nextRanges.push(
            `${startColLetter}${startCellRow}:${endColLetter}${endCellRow}`,
          );
        }
        nextRanges.push(...ranges.slice(i + 1));
        break;
      }

      for (let col = 0; col < colCount; col++) {
        const value = values[row]?.[col];
        const formula = formulas[row]?.[col];
        const note = notes[row]?.[col];

        if (value === "" && (!formula || formula === "")) continue;

        const cellAddress = `${columnToLetter(startCol + col - 1)}${startRow + row}`;

        if (formula && typeof formula === "string" && formula.startsWith("=")) {
          if (note) {
            cells[cellAddress] = [value, formula, note];
          } else {
            cells[cellAddress] = [value, formula];
          }
        } else if (note) {
          cells[cellAddress] = [value, null, note];
        } else {
          cells[cellAddress] = value;
        }

        totalCellCount++;
      }
    }

    // Include styles if requested
    if (includeStyles && rowCount > 0 && colCount > 0) {
      const backgrounds = range.getBackgrounds();
      const fontColors = range.getFontColors();
      const fontWeights = range.getFontWeights();
      const fontStyles = range.getFontStyles();
      const fontSizes = range.getFontSizes();
      const fontFamilies = range.getFontFamilies();
      const horizontalAlignments = range.getHorizontalAlignments();
      const numberFormats = range.getNumberFormats();

      for (let row = 0; row < rowCount; row++) {
        for (let col = 0; col < colCount; col++) {
          const cellAddress = `${columnToLetter(startCol + col - 1)}${startRow + row}`;
          const cellStyles: Record<string, unknown> = {};

          const bg = backgrounds[row]?.[col];
          if (bg && bg !== "#ffffff") {
            cellStyles.fgColor = bg;
          }

          const fontColor = fontColors[row]?.[col];
          if (fontColor && fontColor !== "#000000") {
            cellStyles.color = fontColor;
          }

          const fontWeight = fontWeights[row]?.[col];
          if (fontWeight === "bold") {
            cellStyles.b = true;
          }

          const fontStyle = fontStyles[row]?.[col];
          if (fontStyle === "italic") {
            cellStyles.i = true;
          }

          const fontSize = fontSizes[row]?.[col];
          if (fontSize && fontSize !== 10) {
            cellStyles.sz = fontSize;
          }

          const fontFamily = fontFamilies[row]?.[col];
          if (fontFamily && fontFamily !== "Arial") {
            cellStyles.family = fontFamily;
          }

          const alignment = horizontalAlignments[row]?.[col];
          if (alignment && alignment !== "general") {
            cellStyles.alignment = alignment;
          }

          const numFmt = numberFormats[row]?.[col];
          if (numFmt && numFmt !== "General" && numFmt !== "@") {
            cellStyles.numFmt = numFmt;
          }

          if (Object.keys(cellStyles).length > 0) {
            styles[cellAddress] = cellStyles;
          }
        }
      }
    }
  }

  const result: InferToolOutput<typeof tools.getCellRanges> = {
    worksheet: { name: sheet.getName(), sheetId, dimension, cells },
    hasMore,
  };

  if (includeStyles && Object.keys(styles).length > 0) {
    result.worksheet.styles = styles;
  }

  if (Object.keys(borders).length > 0) {
    result.worksheet.borders = borders;
  }

  if (hasMore && nextRanges.length > 0) {
    result.nextRanges = nextRanges;
  }

  return result;
}

export async function searchData(
  input: InferToolInput<typeof tools.searchData>,
): Promise<InferToolOutput<typeof tools.searchData>> {
  const { searchTerm, sheetId, range, offset = 0, options = {} } = input;
  const {
    matchCase = false,
    matchEntireCell = false,
    maxResults = 500,
  } = options;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  const sheetsToSearch =
    sheetId !== undefined
      ? sheets.filter((s) => s.getSheetId() === sheetId)
      : sheets;

  if (sheetsToSearch.length === 0) {
    return {
      matches: [],
      totalFound: 0,
      returned: 0,
      offset,
      hasMore: false,
      searchTerm,
      searchScope: sheetId !== undefined ? `Sheet ID ${sheetId}` : "All sheets",
      nextOffset: null,
      message:
        sheetId !== undefined
          ? `Sheet with ID ${sheetId} not found`
          : undefined,
    };
  }

  const matches: Array<{
    sheetName: string;
    sheetId: number;
    a1: string;
    value: unknown;
    formula: string | null;
    row: number;
    column: number;
  }> = [];

  const searchTermLower = matchCase ? searchTerm : searchTerm.toLowerCase();
  let matchCount = 0; // Track total matches found (including skipped ones)
  let foundMore = false; // Track if there are more matches beyond what we collected

  for (const sheet of sheetsToSearch) {
    const searchRange = range ? sheet.getRange(range) : sheet.getDataRange();
    const values = searchRange.getValues();
    const formulas = searchRange.getFormulas();
    const startRow = searchRange.getRow();
    const startCol = searchRange.getColumn();

    for (let row = 0; row < values.length; row++) {
      const rowValues = values[row];
      if (!rowValues) continue;

      for (let col = 0; col < rowValues.length; col++) {
        const cellValue = rowValues[col];
        const cellValueStr = String(cellValue ?? "");
        const compareValue = matchCase
          ? cellValueStr
          : cellValueStr.toLowerCase();

        let isMatch = false;
        if (matchEntireCell) {
          isMatch = compareValue === searchTermLower;
        } else {
          isMatch = compareValue.includes(searchTermLower);
        }

        if (isMatch) {
          // Only collect matches after we've passed the offset and before hitting maxResults
          if (matchCount >= offset && matches.length < maxResults) {
            const cellAddress = `${columnToLetter(startCol + col - 1)}${startRow + row}`;
            const formula = formulas[row]?.[col];

            matches.push({
              sheetName: sheet.getName(),
              sheetId: sheet.getSheetId(),
              a1: cellAddress,
              value: cellValue,
              formula: formula?.startsWith("=") ? formula : null,
              row: startRow + row,
              column: startCol + col,
            });
          } else if (matches.length >= maxResults) {
            // Found at least one more match beyond our limit
            foundMore = true;
          }
          matchCount++;
        }

        // Early exit if we've filled results and confirmed there's more
        if (matches.length >= maxResults && foundMore) break;
      }

      if (matches.length >= maxResults && foundMore) break;
    }

    if (matches.length >= maxResults && foundMore) break;
  }

  const totalFound = matchCount;
  const hasMore = foundMore;

  return {
    matches,
    totalFound,
    returned: matches.length,
    offset,
    hasMore,
    searchTerm,
    searchScope: sheetId !== undefined ? `Sheet ID ${sheetId}` : "All sheets",
    nextOffset: hasMore ? offset + matches.length : null,
  };
}

export async function setCellRange(
  input: InferToolInput<typeof tools.setCellRange>,
): Promise<InferToolOutput<typeof tools.setCellRange>> {
  const { sheetId, range, cells, copyToRange, resizeWidth, resizeHeight } =
    input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  const targetRange = sheet.getRange(range);
  const rowCount = targetRange.getNumRows();
  const colCount = targetRange.getNumColumns();

  // Validate dimensions
  if (cells.length !== rowCount) {
    throw new Error(
      `Cell array row count (${cells.length}) does not match range row count (${rowCount})`,
    );
  }
  for (let r = 0; r < cells.length; r++) {
    const cellRow = cells[r];
    if (!cellRow || cellRow.length !== colCount) {
      throw new Error(
        `Cell array column count at row ${r} (${cellRow?.length ?? 0}) does not match range column count (${colCount})`,
      );
    }
  }

  // Build values and formulas arrays
  const values: unknown[][] = [];
  const formulas: (string | null)[][] = [];
  const hasFormulas = cells.some((row) => row.some((cell) => cell?.formula));

  for (let r = 0; r < rowCount; r++) {
    const valueRow: unknown[] = [];
    const formulaRow: (string | null)[] = [];
    for (let c = 0; c < colCount; c++) {
      const cellRow = cells[r];
      const cell = cellRow?.[c];
      if (cell?.formula) {
        formulaRow.push(cell.formula);
        valueRow.push(null);
      } else {
        valueRow.push(cell?.value ?? null);
        formulaRow.push(null);
      }
    }
    values.push(valueRow);
    formulas.push(formulaRow);
  }

  // Set values first (for cells without formulas)
  const valuesForWrite = values.map((row, r) => {
    const formulaRow = formulas[r];
    return row.map((val, c) => (formulaRow?.[c] ? "" : val));
  });
  targetRange.setValues(valuesForWrite);

  // Set formulas where applicable
  if (hasFormulas) {
    for (let r = 0; r < rowCount; r++) {
      const formulaRow = formulas[r];
      if (!formulaRow) continue;
      for (let c = 0; c < colCount; c++) {
        const formula = formulaRow[c];
        if (formula) {
          targetRange.getCell(r + 1, c + 1).setFormula(formula);
        }
      }
    }
  }

  // Apply cell styles
  for (let r = 0; r < rowCount; r++) {
    const cellRow = cells[r];
    if (!cellRow) continue;
    for (let c = 0; c < colCount; c++) {
      const cell = cellRow[c];
      if (!cell) continue;

      const cellRange = targetRange.getCell(r + 1, c + 1);

      if (cell.cellStyles) {
        const styles = cell.cellStyles;

        if (styles.fontColor) {
          cellRange.setFontColor(styles.fontColor);
        }
        if (styles.fontSize) {
          cellRange.setFontSize(styles.fontSize);
        }
        if (styles.fontFamily) {
          cellRange.setFontFamily(styles.fontFamily);
        }
        if (styles.fontWeight === "bold") {
          cellRange.setFontWeight("bold");
        } else if (styles.fontWeight === "normal") {
          cellRange.setFontWeight("normal");
        }
        if (styles.fontStyle === "italic") {
          cellRange.setFontStyle("italic");
        } else if (styles.fontStyle === "normal") {
          cellRange.setFontStyle("normal");
        }
        if (styles.fontLine === "underline") {
          cellRange.setFontLine("underline");
        } else if (styles.fontLine === "line-through") {
          cellRange.setFontLine("line-through");
        } else if (styles.fontLine === "none") {
          cellRange.setFontLine("none");
        }
        if (styles.backgroundColor) {
          cellRange.setBackground(styles.backgroundColor);
        }
        if (styles.horizontalAlignment) {
          cellRange.setHorizontalAlignment(styles.horizontalAlignment);
        }
        if (styles.numberFormat) {
          cellRange.setNumberFormat(styles.numberFormat);
        }
      }

      // Apply border styles
      if (cell.borderStyles) {
        const borderStyles = cell.borderStyles;

        const getBorderStyle = (
          style?: string,
        ): GoogleAppsScript.Spreadsheet.BorderStyle | null => {
          if (!style) return null;
          switch (style) {
            case "solid":
              return SpreadsheetApp.BorderStyle.SOLID;
            case "dashed":
              return SpreadsheetApp.BorderStyle.DASHED;
            case "dotted":
              return SpreadsheetApp.BorderStyle.DOTTED;
            case "double":
              return SpreadsheetApp.BorderStyle.DOUBLE;
            default:
              return SpreadsheetApp.BorderStyle.SOLID;
          }
        };

        if (borderStyles.top) {
          cellRange.setBorder(
            true,
            null,
            null,
            null,
            null,
            null,
            borderStyles.top.color || "#000000",
            getBorderStyle(borderStyles.top.style),
          );
        }
        if (borderStyles.bottom) {
          cellRange.setBorder(
            null,
            null,
            true,
            null,
            null,
            null,
            borderStyles.bottom.color || "#000000",
            getBorderStyle(borderStyles.bottom.style),
          );
        }
        if (borderStyles.left) {
          cellRange.setBorder(
            null,
            true,
            null,
            null,
            null,
            null,
            borderStyles.left.color || "#000000",
            getBorderStyle(borderStyles.left.style),
          );
        }
        if (borderStyles.right) {
          cellRange.setBorder(
            null,
            null,
            null,
            true,
            null,
            null,
            borderStyles.right.color || "#000000",
            getBorderStyle(borderStyles.right.style),
          );
        }
      }

      // Apply notes
      if (cell.note) {
        cellRange.setNote(cell.note);
      }
    }
  }

  // Handle copyToRange if specified
  if (copyToRange) {
    const destRange = sheet.getRange(copyToRange);
    targetRange.copyTo(destRange);
  }

  // Handle resizeWidth
  if (resizeWidth) {
    const startCol = targetRange.getColumn();
    if (resizeWidth.type === "autofit") {
      for (let c = 0; c < colCount; c++) {
        sheet.autoResizeColumn(startCol + c);
      }
    } else if (
      resizeWidth.type === "points" &&
      resizeWidth.value !== undefined
    ) {
      for (let c = 0; c < colCount; c++) {
        sheet.setColumnWidth(startCol + c, resizeWidth.value);
      }
    } else if (resizeWidth.type === "standard") {
      for (let c = 0; c < colCount; c++) {
        sheet.setColumnWidth(startCol + c, 100); // Default Google Sheets column width
      }
    }
  }

  // Handle resizeHeight
  if (resizeHeight) {
    const startRow = targetRange.getRow();
    if (resizeHeight.type === "autofit") {
      for (let r = 0; r < rowCount; r++) {
        sheet.autoResizeRows(startRow + r, 1);
      }
    } else if (
      resizeHeight.type === "points" &&
      resizeHeight.value !== undefined
    ) {
      for (let r = 0; r < rowCount; r++) {
        sheet.setRowHeight(startRow + r, resizeHeight.value);
      }
    } else if (resizeHeight.type === "standard") {
      for (let r = 0; r < rowCount; r++) {
        sheet.setRowHeight(startRow + r, 21); // Default Google Sheets row height
      }
    }
  }

  // Get formula results for cells with formulas
  const formulaResults: Record<string, number | string> = {};
  if (hasFormulas) {
    SpreadsheetApp.flush(); // Force calculation
    const resultValues = targetRange.getValues();
    const startRow = targetRange.getRow();
    const startCol = targetRange.getColumn();

    for (let r = 0; r < rowCount; r++) {
      const formulaRow = formulas[r];
      if (!formulaRow) continue;
      for (let c = 0; c < colCount; c++) {
        if (formulaRow[c]) {
          const cellAddress = `${columnToLetter(startCol + c - 1)}${startRow + r}`;
          const value = resultValues[r]?.[c];
          formulaResults[cellAddress] =
            typeof value === "number" || typeof value === "string"
              ? value
              : String(value);
        }
      }
    }
  }

  const result: InferToolOutput<typeof tools.setCellRange> = {};
  if (Object.keys(formulaResults).length > 0) {
    result.formula_results = formulaResults;
  }

  return result;
}

export async function modifySheetStructure(
  input: InferToolInput<typeof tools.modifySheetStructure>,
): Promise<InferToolOutput<typeof tools.modifySheetStructure>> {
  const {
    sheetId,
    operation,
    dimension,
    reference,
    position = "before",
    count = 1,
  } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  switch (operation) {
    case "insert": {
      if (!dimension || !reference) {
        throw new Error(
          "dimension and reference are required for insert operation",
        );
      }

      if (dimension === "rows") {
        const rowNum = Number.parseInt(reference, 10);
        if (position === "before") {
          sheet.insertRowsBefore(rowNum, count);
        } else {
          sheet.insertRowsAfter(rowNum, count);
        }
      } else {
        const colLetter = reference.toUpperCase();
        const colIndex = letterToColumn(colLetter); // 0-based
        const colNum = colIndex + 1; // Convert to 1-based for Google Apps Script
        if (position === "before") {
          sheet.insertColumnsBefore(colNum, count);
        } else {
          sheet.insertColumnsAfter(colNum, count);
        }
      }
      break;
    }

    case "delete": {
      if (!dimension || !reference) {
        throw new Error(
          "dimension and reference are required for delete operation",
        );
      }

      if (dimension === "rows") {
        const rowNum = Number.parseInt(reference, 10);
        sheet.deleteRows(rowNum, count);
      } else {
        const colLetter = reference.toUpperCase();
        const colIndex = letterToColumn(colLetter);
        sheet.deleteColumns(colIndex + 1, count);
      }
      break;
    }

    case "hide": {
      if (!dimension || !reference) {
        throw new Error(
          "dimension and reference are required for hide operation",
        );
      }

      if (dimension === "rows") {
        const rowNum = Number.parseInt(reference, 10);
        sheet.hideRows(rowNum, count);
      } else {
        const colLetter = reference.toUpperCase();
        const colIndex = letterToColumn(colLetter);
        sheet.hideColumns(colIndex + 1, count);
      }
      break;
    }

    case "unhide": {
      if (!dimension || !reference) {
        throw new Error(
          "dimension and reference are required for unhide operation",
        );
      }

      if (dimension === "rows") {
        const rowNum = Number.parseInt(reference, 10);
        sheet.showRows(rowNum, count);
      } else {
        const colLetter = reference.toUpperCase();
        const colIndex = letterToColumn(colLetter);
        sheet.showColumns(colIndex + 1, count);
      }
      break;
    }

    case "freeze": {
      if (!dimension || count === undefined) {
        throw new Error(
          "dimension and count are required for freeze operation",
        );
      }

      if (dimension === "rows") {
        sheet.setFrozenRows(count);
      } else {
        sheet.setFrozenColumns(count);
      }
      break;
    }

    case "unfreeze": {
      sheet.setFrozenRows(0);
      sheet.setFrozenColumns(0);
      break;
    }
  }

  return {};
}

export async function modifyWorkbookStructure(
  input: InferToolInput<typeof tools.modifyWorkbookStructure>,
): Promise<InferToolOutput<typeof tools.modifyWorkbookStructure>> {
  const { operation, sheetName, sheetId, newName, tabColor } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  switch (operation) {
    case "create": {
      if (!sheetName) {
        throw new Error("sheetName is required for create operation");
      }

      const newSheet = spreadsheet.insertSheet(sheetName);

      if (tabColor) {
        newSheet.setTabColor(tabColor);
      }

      spreadsheet.setActiveSheet(newSheet);

      return {
        sheetId: newSheet.getSheetId(),
        sheetName: newSheet.getName(),
        message: `Sheet "${sheetName}" created successfully`,
      };
    }

    case "delete": {
      if (sheetId === undefined) {
        throw new Error("sheetId is required for delete operation");
      }

      const sheet = sheets.find((s) => s.getSheetId() === sheetId);
      if (!sheet) {
        throw new Error(`Sheet with ID ${sheetId} not found`);
      }

      const deletedName = sheet.getName();
      spreadsheet.deleteSheet(sheet);

      return {
        message: `Sheet "${deletedName}" deleted successfully`,
      };
    }

    case "rename": {
      if (sheetId === undefined || !newName) {
        throw new Error(
          "sheetId and newName are required for rename operation",
        );
      }

      const sheet = sheets.find((s) => s.getSheetId() === sheetId);
      if (!sheet) {
        throw new Error(`Sheet with ID ${sheetId} not found`);
      }

      const oldName = sheet.getName();
      sheet.setName(newName);

      return {
        sheetId,
        sheetName: newName,
        message: `Sheet renamed from "${oldName}" to "${newName}"`,
      };
    }

    case "duplicate": {
      if (sheetId === undefined) {
        throw new Error("sheetId is required for duplicate operation");
      }

      const sheet = sheets.find((s) => s.getSheetId() === sheetId);
      if (!sheet) {
        throw new Error(`Sheet with ID ${sheetId} not found`);
      }

      const copiedSheet = sheet.copyTo(spreadsheet);

      if (newName) {
        copiedSheet.setName(newName);
      }

      spreadsheet.setActiveSheet(copiedSheet);

      return {
        sheetId: copiedSheet.getSheetId(),
        sheetName: copiedSheet.getName(),
        message: "Sheet duplicated successfully",
      };
    }

    default:
      throw new Error(`Unknown operation: ${operation}`);
  }
}

export async function copyTo(
  input: InferToolInput<typeof tools.copyTo>,
): Promise<InferToolOutput<typeof tools.copyTo>> {
  const { sheetId, sourceRange, destinationRange } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  const source = sheet.getRange(sourceRange);
  const destination = sheet.getRange(destinationRange);

  source.copyTo(destination);

  return {};
}

export async function getAllObjects(
  input: InferToolInput<typeof tools.getAllObjects>,
): Promise<InferToolOutput<typeof tools.getAllObjects>> {
  const { sheetId, id } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  const sheetsToSearch =
    sheetId !== undefined
      ? sheets.filter((s) => s.getSheetId() === sheetId)
      : sheets;

  const objects: InferToolOutput<typeof tools.getAllObjects>["objects"] = [];

  for (const sheet of sheetsToSearch) {
    // Get pivot tables (Google Sheets doesn't have direct API for pivot tables in Apps Script)
    // Pivot tables are embedded in sheets and can't be easily enumerated via Apps Script
    // This is a limitation of Google Apps Script

    // Get charts
    const charts = sheet.getCharts();

    for (const chart of charts) {
      const chartId = String(chart.getChartId());
      if (id && chartId !== id) continue;

      const chartTypeMap: Record<string, string> = {
        COLUMN: "columnClustered",
        BAR: "barClustered",
        LINE: "line",
        AREA: "area",
        PIE: "pie",
        SCATTER: "scatter",
        COMBO: "columnClustered",
        STEPPED_AREA: "areaStacked",
        TABLE: "columnClustered",
        HISTOGRAM: "columnClustered",
        BUBBLE: "bubble",
        TREEMAP: "columnClustered",
        WATERFALL: "columnClustered",
        CANDLESTICK: "stockOHLC",
        ORG: "columnClustered",
        RADAR: "radar",
      };

      const options = chart.getOptions();
      // EmbeddedChart doesn't expose getChartType directly, use options
      const chartTypeStr = String(options.get("chartType") || "COLUMN");
      const chartType = chartTypeMap[chartTypeStr] || "columnClustered";
      const title = options.get("title") || "";

      // Get chart position
      const position = chart.getContainerInfo();

      objects.push({
        id: chartId,
        type: "chart",
        sheetId: sheet.getSheetId(),
        chartType: chartType as "columnClustered",
        title: String(title),
        position: {
          top: position.getOffsetY(),
          left: position.getOffsetX(),
        },
      });
    }
  }

  return { objects };
}

export async function modifyObject(
  input: InferToolInput<typeof tools.modifyObject>,
): Promise<InferToolOutput<typeof tools.modifyObject>> {
  const { operation, sheetId, id, objectType, properties } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  if (objectType === "pivotTable") {
    // Google Apps Script has limited support for pivot tables
    // They can be created via the Sheets API but not easily via Apps Script
    throw new Error(
      "Pivot table operations are not fully supported in Google Apps Script. Use the Google Sheets API instead.",
    );
  }

  if (objectType === "chart") {
    switch (operation) {
      case "create": {
        const source =
          properties &&
          "source" in properties &&
          typeof properties.source === "string"
            ? properties.source
            : undefined;
        const chartTypeProp =
          properties &&
          "chartType" in properties &&
          typeof properties.chartType === "string"
            ? properties.chartType
            : undefined;
        const title =
          properties &&
          "title" in properties &&
          typeof properties.title === "string"
            ? properties.title
            : undefined;
        const anchor =
          properties &&
          "anchor" in properties &&
          typeof properties.anchor === "string"
            ? properties.anchor
            : undefined;

        if (!source || !chartTypeProp) {
          throw new Error(
            "source and chartType are required for chart creation",
          );
        }

        const chartTypeMap: Record<string, GoogleAppsScript.Charts.ChartType> =
          {
            columnClustered: Charts.ChartType.COLUMN,
            columnStacked: Charts.ChartType.COLUMN,
            barClustered: Charts.ChartType.BAR,
            barStacked: Charts.ChartType.BAR,
            line: Charts.ChartType.LINE,
            lineMarkers: Charts.ChartType.LINE,
            area: Charts.ChartType.AREA,
            areaStacked: Charts.ChartType.AREA,
            pie: Charts.ChartType.PIE,
            scatter: Charts.ChartType.SCATTER,
            radar: Charts.ChartType.RADAR,
            bubble: Charts.ChartType.BUBBLE,
          };

        const sourceRange = sheet.getRange(source);
        const chartType =
          chartTypeMap[chartTypeProp] || Charts.ChartType.COLUMN;

        const chartBuilder = sheet
          .newChart()
          .setChartType(chartType)
          .addRange(sourceRange)
          .setOption("title", title || "");

        if (anchor) {
          const anchorRange = sheet.getRange(anchor);
          chartBuilder.setPosition(
            anchorRange.getRow(),
            anchorRange.getColumn(),
            0,
            0,
          );
        }

        const chart = chartBuilder.build();
        sheet.insertChart(chart);

        // Get the newly inserted chart's ID
        const insertedCharts = sheet.getCharts();
        const insertedChart = insertedCharts[insertedCharts.length - 1];
        return { id: insertedChart ? String(insertedChart.getChartId()) : "" };
      }

      case "delete": {
        if (!id) {
          throw new Error("id is required for delete operation");
        }

        const charts = sheet.getCharts();
        const chart = charts.find((c) => String(c.getChartId()) === id);

        if (!chart) {
          throw new Error(`Chart with ID ${id} not found`);
        }

        sheet.removeChart(chart);
        return {};
      }

      case "update": {
        if (!id) {
          throw new Error("id is required for update operation");
        }

        const charts = sheet.getCharts();
        const chart = charts.find((c) => String(c.getChartId()) === id);

        if (!chart) {
          throw new Error(`Chart with ID ${id} not found`);
        }

        const title =
          properties && "title" in properties ? properties.title : undefined;
        const anchor =
          properties && "anchor" in properties ? properties.anchor : undefined;

        const chartBuilder = chart.modify();

        if (title && typeof title === "string") {
          chartBuilder.setOption("title", title);
        }

        if (anchor && typeof anchor === "string") {
          const anchorRange = sheet.getRange(anchor);
          chartBuilder.setPosition(
            anchorRange.getRow(),
            anchorRange.getColumn(),
            0,
            0,
          );
        }

        const updatedChart = chartBuilder.build();
        sheet.updateChart(updatedChart);

        return { id };
      }
    }
  }

  throw new Error(`Unknown object type: ${objectType}`);
}

export async function resizeRange(
  input: InferToolInput<typeof tools.resizeRange>,
): Promise<InferToolOutput<typeof tools.resizeRange>> {
  const { sheetId, range, width, height } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  const targetRange = range ? sheet.getRange(range) : sheet.getDataRange();

  const startCol = targetRange.getColumn();
  const colCount = targetRange.getNumColumns();
  const startRow = targetRange.getRow();
  const rowCount = targetRange.getNumRows();

  if (width) {
    if (width.type === "autofit") {
      for (let c = 0; c < colCount; c++) {
        sheet.autoResizeColumn(startCol + c);
      }
    } else if (width.type === "points" && width.value !== undefined) {
      for (let c = 0; c < colCount; c++) {
        sheet.setColumnWidth(startCol + c, width.value);
      }
    } else if (width.type === "standard") {
      for (let c = 0; c < colCount; c++) {
        sheet.setColumnWidth(startCol + c, 100); // Default Google Sheets column width
      }
    }
  }

  if (height) {
    if (height.type === "autofit") {
      sheet.autoResizeRows(startRow, rowCount);
    } else if (height.type === "points" && height.value !== undefined) {
      for (let r = 0; r < rowCount; r++) {
        sheet.setRowHeight(startRow + r, height.value);
      }
    } else if (height.type === "standard") {
      for (let r = 0; r < rowCount; r++) {
        sheet.setRowHeight(startRow + r, 21); // Default Google Sheets row height
      }
    }
  }

  return {};
}

export async function clearCellRange(
  input: InferToolInput<typeof tools.clearCellRange>,
): Promise<InferToolOutput<typeof tools.clearCellRange>> {
  const { sheetId, range, clearType = "contents" } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    throw new Error(`Sheet with ID ${sheetId} not found`);
  }

  const targetRange = sheet.getRange(range);

  switch (clearType) {
    case "all":
      targetRange.clear();
      break;
    case "formats":
      targetRange.clearFormat();
      break;
    case "contents":
    default:
      targetRange.clearContent();
      break;
  }

  return {};
}

/**
 * Activates the specified sheet (switches to it if not already active).
 */
export async function activateSheet(sheetId: number): Promise<void> {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    return;
  }

  spreadsheet.setActiveSheet(sheet);
}

/**
 * Clears the current selection by selecting the active cell.
 */
export async function clearSelection(): Promise<void> {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeCell = spreadsheet.getActiveCell();
  if (activeCell) {
    spreadsheet.setActiveRange(activeCell);
  }
}

/**
 * Activates the specified sheet and selects a range.
 * Use this to show the user which cells will be modified before a write operation.
 */
export async function selectRange(input: {
  sheetId: number;
  range: string;
}): Promise<void> {
  const { sheetId, range } = input;

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const sheet = sheets.find((s) => s.getSheetId() === sheetId);

  if (!sheet) {
    return;
  }

  spreadsheet.setActiveSheet(sheet);
  const targetRange = sheet.getRange(range);
  spreadsheet.setActiveRange(targetRange);
}

// Helper functions

function columnToLetter(columnIndex: number): string {
  let letter = "";
  let temp = columnIndex;
  while (temp >= 0) {
    letter = String.fromCharCode((temp % 26) + 65) + letter;
    temp = Math.floor(temp / 26) - 1;
  }
  return letter;
}

function letterToColumn(letter: string): number {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column = column * 26 + (letter.charCodeAt(i) - 64);
  }
  return column - 1;
}

// Typechecking to ensure the service is implemented correctly
const __googleSheetsService: SpreadsheetService = {
  getSheets,
  getCellRanges,
  searchData,
  getAllObjects,
  setCellRange,
  copyTo,
  clearCellRange,
  resizeRange,
  modifySheetStructure,
  modifyWorkbookStructure,
  modifyObject,
  activateSheet,
  clearSelection,
  selectRange,
};
