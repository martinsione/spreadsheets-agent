import GC from "@mescius/spread-sheets";
import type { tools } from "@repo/core/ai/tools";
import type { Sheet, SpreadsheetService } from "@repo/core/spreadsheet-service";
import type { InferToolInput, InferToolOutput } from "ai";

type Workbook = GC.Spread.Sheets.Workbook;
type Worksheet = GC.Spread.Sheets.Worksheet;

export function createWebSpreadsheetService(
  getWorkbook: () => Workbook | null,
): SpreadsheetService {
  return {
    async getSheets(): Promise<Sheet[]> {
      const workbook = getWorkbook();
      if (!workbook) return [];

      const sheets: Sheet[] = [];
      for (let i = 0; i < workbook.getSheetCount(); i++) {
        const sheet = workbook.getSheet(i);
        const usedRange = sheet.getUsedRange(
          GC.Spread.Sheets.UsedRangeType.all,
        );
        sheets.push({
          id: i,
          name: sheet.name(),
          maxRows: usedRange ? usedRange.rowCount : 0,
          maxColumns: usedRange ? usedRange.colCount : 0,
        });
      }
      return sheets;
    },

    async getCellRanges(
      input: InferToolInput<typeof tools.getCellRanges>,
    ): Promise<InferToolOutput<typeof tools.getCellRanges>> {
      const {
        sheetId,
        ranges,
        includeStyles = true,
        cellLimit = 10000,
      } = input;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const sheet = workbook.getSheet(sheetId);
      if (!sheet) throw new Error(`Sheet ${sheetId} not found`);

      const cells: Record<string, unknown> = {};
      const styles: Record<string, Record<string, unknown>> = {};
      let totalCellCount = 0;
      let hasMore = false;
      const nextRanges: string[] = [];

      const usedRange = sheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.all);
      const dimension = usedRange
        ? `A1:${columnToLetter(usedRange.col + usedRange.colCount - 1)}${usedRange.row + usedRange.rowCount}`
        : undefined;

      for (let i = 0; i < ranges.length; i++) {
        const rangeStr = ranges[i];
        if (!rangeStr) continue;

        if (totalCellCount >= cellLimit) {
          nextRanges.push(...ranges.slice(i));
          hasMore = true;
          break;
        }

        const rangeInfo = parseRangeAddress(rangeStr);
        if (!rangeInfo) continue;

        for (let row = rangeInfo.startRow; row <= rangeInfo.endRow; row++) {
          if (totalCellCount >= cellLimit) {
            hasMore = true;
            break;
          }

          for (let col = rangeInfo.startCol; col <= rangeInfo.endCol; col++) {
            const value = sheet.getValue(row, col);
            const formula = sheet.getFormula(row, col);
            const comment = sheet.comments.get(row, col);
            const cellAddr = `${columnToLetter(col)}${row + 1}`;

            // Build cell value
            if (formula) {
              if (comment) {
                cells[cellAddr] = [value, formula, comment.text()];
              } else {
                cells[cellAddr] = [value, formula];
              }
            } else if (comment) {
              cells[cellAddr] = [value, null, comment.text()];
            } else if (value !== null && value !== "" && value !== undefined) {
              cells[cellAddr] = value;
            }

            // Extract styles if requested
            if (includeStyles) {
              const style = sheet.getStyle(row, col);
              if (style) {
                const cellStyles: Record<string, unknown> = {};

                if (style.backColor && style.backColor !== "#ffffff") {
                  cellStyles.fgColor = style.backColor;
                }
                if (style.foreColor && style.foreColor !== "#000000") {
                  cellStyles.color = style.foreColor;
                }
                if (style.font) {
                  const fontParts = parseFontString(style.font);
                  if (fontParts.bold) cellStyles.b = true;
                  if (fontParts.italic) cellStyles.i = true;
                  if (fontParts.size && fontParts.size !== 11) {
                    cellStyles.sz = fontParts.size;
                  }
                  if (fontParts.family && fontParts.family !== "Calibri") {
                    cellStyles.family = fontParts.family;
                  }
                }
                if (style.hAlign !== undefined && style.hAlign !== 3) {
                  // 3 = general
                  const alignMap: Record<number, string> = {
                    0: "left",
                    1: "center",
                    2: "right",
                  };
                  cellStyles.alignment = alignMap[style.hAlign] || "general";
                }
                if (style.formatter) {
                  cellStyles.numFmt = style.formatter;
                }
                if (
                  style.textDecoration !== undefined &&
                  style.textDecoration !== 0
                ) {
                  if (style.textDecoration & 1) cellStyles.u = "single";
                  if (style.textDecoration & 2) cellStyles.strike = true;
                }

                if (Object.keys(cellStyles).length > 0) {
                  styles[cellAddr] = cellStyles;
                }
              }
            }

            totalCellCount++;
          }
        }
      }

      const result: InferToolOutput<typeof tools.getCellRanges> = {
        worksheet: {
          name: sheet.name(),
          sheetId,
          dimension,
          cells,
        },
        hasMore,
      };

      if (includeStyles && Object.keys(styles).length > 0) {
        result.worksheet.styles = styles;
      }

      if (hasMore && nextRanges.length > 0) {
        result.nextRanges = nextRanges;
      }

      return result;
    },

    async setCellRange(
      input: InferToolInput<typeof tools.setCellRange>,
    ): Promise<InferToolOutput<typeof tools.setCellRange>> {
      const { sheetId, range, cells, copyToRange, resizeWidth, resizeHeight } =
        input;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const sheet = workbook.getSheet(sheetId);
      if (!sheet) throw new Error(`Sheet ${sheetId} not found`);

      const rangeInfo = parseRangeAddress(range);
      if (!rangeInfo) throw new Error(`Invalid range: ${range}`);

      sheet.suspendPaint();
      sheet.suspendCalcService(false);
      try {
        const formulaResults: Record<string, number | string> = {};

        for (let r = 0; r < cells.length; r++) {
          const row = cells[r];
          if (!row) continue;
          for (let c = 0; c < row.length; c++) {
            const cell = row[c];
            if (!cell) continue;
            const targetRow = rangeInfo.startRow + r;
            const targetCol = rangeInfo.startCol + c;

            // Set value or formula
            if (cell.formula) {
              sheet.setFormula(targetRow, targetCol, cell.formula);
            } else if (cell.value !== undefined) {
              sheet.setValue(targetRow, targetCol, cell.value);
            }

            // Set note/comment
            if (cell.note) {
              const existingComment = sheet.comments.get(targetRow, targetCol);
              if (existingComment) {
                existingComment.text(cell.note);
              } else {
                sheet.comments.add(targetRow, targetCol, cell.note);
              }
            }

            // Apply cell styles
            if (cell.cellStyles) {
              const currentStyle = sheet.getStyle(targetRow, targetCol);
              const newStyle = new GC.Spread.Sheets.Style();

              // Copy existing style properties if available
              if (currentStyle) {
                Object.assign(newStyle, currentStyle);
              }

              if (cell.cellStyles.fontColor) {
                newStyle.foreColor = cell.cellStyles.fontColor;
              }
              if (cell.cellStyles.backgroundColor) {
                newStyle.backColor = cell.cellStyles.backgroundColor;
              }

              // Build font string
              const fontParts: string[] = [];
              if (cell.cellStyles.fontWeight === "bold") {
                fontParts.push("bold");
              }
              if (cell.cellStyles.fontStyle === "italic") {
                fontParts.push("italic");
              }
              const fontSize = cell.cellStyles.fontSize || 11;
              const fontFamily = cell.cellStyles.fontFamily || "Calibri";
              fontParts.push(`${fontSize}pt`);
              fontParts.push(fontFamily);

              if (
                cell.cellStyles.fontWeight ||
                cell.cellStyles.fontStyle ||
                cell.cellStyles.fontSize ||
                cell.cellStyles.fontFamily
              ) {
                newStyle.font = fontParts.join(" ");
              }

              // Text decoration
              if (cell.cellStyles.fontLine) {
                let decoration = 0;
                if (cell.cellStyles.fontLine === "underline") {
                  decoration |= 1;
                } else if (cell.cellStyles.fontLine === "line-through") {
                  decoration |= 2;
                }
                newStyle.textDecoration = decoration;
              }

              // Horizontal alignment
              if (cell.cellStyles.horizontalAlignment) {
                const alignMap: Record<string, number> = {
                  left: 0,
                  center: 1,
                  right: 2,
                };
                newStyle.hAlign =
                  alignMap[cell.cellStyles.horizontalAlignment] ?? 3;
              }

              // Number format
              if (cell.cellStyles.numberFormat) {
                newStyle.formatter = cell.cellStyles.numberFormat;
              }

              sheet.setStyle(targetRow, targetCol, newStyle);
            }

            // Apply border styles
            if (cell.borderStyles) {
              applyBorderStyles(sheet, targetRow, targetCol, cell.borderStyles);
            }
          }
        }

        // Resume calc to compute formulas
        sheet.resumeCalcService(true);

        // Collect formula results
        for (let r = 0; r < cells.length; r++) {
          const row = cells[r];
          if (!row) continue;
          for (let c = 0; c < row.length; c++) {
            const cell = row[c];
            if (!cell?.formula) continue;
            const targetRow = rangeInfo.startRow + r;
            const targetCol = rangeInfo.startCol + c;
            const cellAddr = `${columnToLetter(targetCol)}${targetRow + 1}`;
            const computed = sheet.getValue(targetRow, targetCol);
            if (computed !== null && computed !== undefined) {
              formulaResults[cellAddr] =
                typeof computed === "number" || typeof computed === "string"
                  ? computed
                  : String(computed);
            }
          }
        }

        // Copy to range if specified
        if (copyToRange) {
          const destInfo = parseRangeAddress(copyToRange);
          if (destInfo) {
            const srcRowCount = rangeInfo.endRow - rangeInfo.startRow + 1;
            const srcColCount = rangeInfo.endCol - rangeInfo.startCol + 1;
            const destRowCount = destInfo.endRow - destInfo.startRow + 1;
            const destColCount = destInfo.endCol - destInfo.startCol + 1;

            // Repeat the pattern to fill destination
            for (let r = 0; r < destRowCount; r++) {
              for (let c = 0; c < destColCount; c++) {
                const srcRow = rangeInfo.startRow + (r % srcRowCount);
                const srcCol = rangeInfo.startCol + (c % srcColCount);
                const dstRow = destInfo.startRow + r;
                const dstCol = destInfo.startCol + c;

                // Skip if same cell
                if (srcRow === dstRow && srcCol === dstCol) continue;

                const formula = sheet.getFormula(srcRow, srcCol);
                if (formula) {
                  // Use copyTo for proper formula translation
                  sheet.copyTo(
                    srcRow,
                    srcCol,
                    dstRow,
                    dstCol,
                    1,
                    1,
                    GC.Spread.Sheets.CopyToOptions.all,
                  );
                } else {
                  const value = sheet.getValue(srcRow, srcCol);
                  sheet.setValue(dstRow, dstCol, value);
                  const style = sheet.getStyle(srcRow, srcCol);
                  if (style) {
                    sheet.setStyle(dstRow, dstCol, style);
                  }
                }
              }
            }
          }
        }

        // Resize columns
        if (resizeWidth) {
          for (let col = rangeInfo.startCol; col <= rangeInfo.endCol; col++) {
            if (resizeWidth.type === "autofit") {
              sheet.autoFitColumn(col);
            } else if (
              resizeWidth.type === "points" &&
              resizeWidth.value !== undefined
            ) {
              sheet.setColumnWidth(col, resizeWidth.value);
            } else if (resizeWidth.type === "standard") {
              sheet.setColumnWidth(col, 64);
            }
          }
        }

        // Resize rows
        if (resizeHeight) {
          for (let row = rangeInfo.startRow; row <= rangeInfo.endRow; row++) {
            if (resizeHeight.type === "autofit") {
              sheet.autoFitRow(row);
            } else if (
              resizeHeight.type === "points" &&
              resizeHeight.value !== undefined
            ) {
              sheet.setRowHeight(row, resizeHeight.value);
            } else if (resizeHeight.type === "standard") {
              sheet.setRowHeight(row, 20);
            }
          }
        }

        const result: InferToolOutput<typeof tools.setCellRange> = {};
        if (Object.keys(formulaResults).length > 0) {
          result.formula_results = formulaResults;
        }
        return result;
      } finally {
        sheet.resumePaint();
      }
    },

    async searchData(
      input: InferToolInput<typeof tools.searchData>,
    ): Promise<InferToolOutput<typeof tools.searchData>> {
      const { searchTerm, sheetId, range, options = {} } = input;
      const {
        matchCase = false,
        matchEntireCell = false,
        maxResults = 500,
      } = options;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const matches: Array<{
        sheetName: string;
        sheetId: number;
        a1: string;
        value: unknown;
        formula: string | null;
        row: number;
        column: number;
      }> = [];

      const sheetsToSearch =
        sheetId !== undefined
          ? [workbook.getSheet(sheetId)]
          : Array.from({ length: workbook.getSheetCount() }, (_, i) =>
              workbook.getSheet(i),
            );

      const searchTermLower = matchCase ? searchTerm : searchTerm.toLowerCase();

      // Parse optional range parameter to limit search scope
      const rangeLimit = range ? parseRangeAddress(range) : null;

      for (const sheet of sheetsToSearch) {
        if (!sheet) continue;
        const usedRange = sheet.getUsedRange(
          GC.Spread.Sheets.UsedRangeType.all,
        );
        if (!usedRange) continue;

        // Determine search bounds (use rangeLimit if provided, otherwise use usedRange)
        const startRow = rangeLimit ? rangeLimit.startRow : usedRange.row;
        const endRow = rangeLimit
          ? rangeLimit.endRow
          : usedRange.row + usedRange.rowCount - 1;
        const startCol = rangeLimit ? rangeLimit.startCol : usedRange.col;
        const endCol = rangeLimit
          ? rangeLimit.endCol
          : usedRange.col + usedRange.colCount - 1;

        for (let row = startRow; row <= endRow; row++) {
          for (let col = startCol; col <= endCol; col++) {
            if (matches.length >= maxResults) break;

            const value = sheet.getValue(row, col);
            if (value === null || value === undefined) continue;

            const valueStr = String(value);
            const compareValue = matchCase ? valueStr : valueStr.toLowerCase();

            let isMatch = false;
            if (matchEntireCell) {
              isMatch = compareValue === searchTermLower;
            } else {
              isMatch = compareValue.includes(searchTermLower);
            }

            if (isMatch) {
              const formula = sheet.getFormula(row, col);
              matches.push({
                sheetName: sheet.name(),
                sheetId: workbook.getSheetIndex(sheet.name()),
                a1: `${columnToLetter(col)}${row + 1}`,
                value,
                formula: formula || null,
                row: row + 1,
                column: col + 1,
              });
            }
          }
          if (matches.length >= maxResults) break;
        }
        if (matches.length >= maxResults) break;
      }

      return {
        matches,
        totalFound: matches.length,
        returned: matches.length,
        offset: 0,
        hasMore: false,
        searchTerm,
        searchScope: sheetId !== undefined ? `Sheet ${sheetId}` : "All sheets",
        nextOffset: null,
      };
    },

    async modifySheetStructure(
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
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const sheet = workbook.getSheet(sheetId);
      if (!sheet) throw new Error(`Sheet ${sheetId} not found`);

      sheet.suspendPaint();
      try {
        switch (operation) {
          case "insert":
            if (dimension === "rows" && reference) {
              const rowNum = parseInt(reference, 10);
              const rowIndex = position === "after" ? rowNum : rowNum - 1;
              sheet.addRows(rowIndex, count);
            } else if (dimension === "columns" && reference) {
              const colIndex = letterToColumn(reference);
              const startCol = position === "after" ? colIndex + 1 : colIndex;
              sheet.addColumns(startCol, count);
            }
            break;
          case "delete":
            if (dimension === "rows" && reference) {
              const rowIndex = parseInt(reference, 10) - 1;
              sheet.deleteRows(rowIndex, count);
            } else if (dimension === "columns" && reference) {
              const colIndex = letterToColumn(reference);
              sheet.deleteColumns(colIndex, count);
            }
            break;
          case "freeze":
            if (dimension === "rows") {
              sheet.frozenRowCount(count);
            } else if (dimension === "columns") {
              sheet.frozenColumnCount(count);
            }
            break;
          case "unfreeze":
            sheet.frozenRowCount(0);
            sheet.frozenColumnCount(0);
            break;
          case "hide":
            if (dimension === "rows" && reference) {
              const rowStart = parseInt(reference, 10) - 1;
              for (let i = 0; i < count; i++) {
                sheet.setRowVisible(rowStart + i, false);
              }
            } else if (dimension === "columns" && reference) {
              const colStart = letterToColumn(reference);
              for (let i = 0; i < count; i++) {
                sheet.setColumnVisible(colStart + i, false);
              }
            }
            break;
          case "unhide":
            if (dimension === "rows" && reference) {
              const rowStart = parseInt(reference, 10) - 1;
              for (let i = 0; i < count; i++) {
                sheet.setRowVisible(rowStart + i, true);
              }
            } else if (dimension === "columns" && reference) {
              const colStart = letterToColumn(reference);
              for (let i = 0; i < count; i++) {
                sheet.setColumnVisible(colStart + i, true);
              }
            }
            break;
        }
      } finally {
        sheet.resumePaint();
      }

      return {};
    },

    async modifyWorkbookStructure(
      input: InferToolInput<typeof tools.modifyWorkbookStructure>,
    ): Promise<InferToolOutput<typeof tools.modifyWorkbookStructure>> {
      const { operation, sheetName, sheetId, newName, tabColor } = input;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      workbook.suspendPaint();
      try {
        switch (operation) {
          case "create": {
            if (!sheetName)
              throw new Error("sheetName is required for create operation");
            const newIndex = workbook.getSheetCount();
            const newSheet = new GC.Spread.Sheets.Worksheet(sheetName);
            workbook.addSheet(newIndex, newSheet);

            if (tabColor) {
              newSheet.options.sheetTabColor = tabColor;
            }

            workbook.setActiveSheetIndex(newIndex);
            return {
              sheetId: newIndex,
              sheetName,
              message: `Sheet "${sheetName}" created successfully`,
            };
          }
          case "delete": {
            if (sheetId === undefined)
              throw new Error("sheetId is required for delete operation");
            const sheetToDelete = workbook.getSheet(sheetId);
            if (!sheetToDelete)
              throw new Error(`Sheet with ID ${sheetId} not found`);

            const deletedName = sheetToDelete.name();
            workbook.removeSheet(sheetId);
            return { message: `Sheet "${deletedName}" deleted successfully` };
          }
          case "rename": {
            if (sheetId === undefined || !newName) {
              throw new Error(
                "sheetId and newName are required for rename operation",
              );
            }
            const sheetToRename = workbook.getSheet(sheetId);
            if (!sheetToRename)
              throw new Error(`Sheet with ID ${sheetId} not found`);

            const oldName = sheetToRename.name();
            sheetToRename.name(newName);
            return {
              sheetId,
              sheetName: newName,
              message: `Sheet renamed from "${oldName}" to "${newName}"`,
            };
          }
          case "duplicate": {
            if (sheetId === undefined)
              throw new Error("sheetId is required for duplicate operation");
            const source = workbook.getSheet(sheetId);
            if (!source) throw new Error(`Sheet with ID ${sheetId} not found`);

            const copyName = newName || `${source.name()} (Copy)`;
            const newIndex = workbook.getSheetCount();

            // Create new sheet
            const copiedSheet = new GC.Spread.Sheets.Worksheet(copyName);
            workbook.addSheet(newIndex, copiedSheet);

            // Copy all data and formatting using fromJSON/toJSON for complete copy
            const sourceData = source.toJSON();
            copiedSheet.fromJSON(sourceData);
            copiedSheet.name(copyName);

            workbook.setActiveSheetIndex(newIndex);
            return {
              sheetId: newIndex,
              sheetName: copyName,
              message: "Sheet duplicated successfully",
            };
          }
          default:
            throw new Error(`Unknown operation: ${operation}`);
        }
      } finally {
        workbook.resumePaint();
      }
    },

    async activateSheet(sheetId: number) {
      const workbook = getWorkbook();
      if (!workbook) return;

      if (sheetId >= 0 && sheetId < workbook.getSheetCount()) {
        workbook.setActiveSheetIndex(sheetId);
      }
    },

    async clearSelection() {
      const workbook = getWorkbook();
      if (!workbook) return;

      const sheet = workbook.getActiveSheet();
      if (sheet) {
        // Clear selection by selecting just the active cell
        const selection = sheet.getSelections();
        if (selection && selection.length > 0) {
          const firstSel = selection[0];
          if (firstSel) {
            sheet.setSelection(firstSel.row, firstSel.col, 1, 1);
          }
        }
      }
    },

    async selectRange(input) {
      const { sheetId, range } = input;
      const workbook = getWorkbook();
      if (!workbook) return;

      workbook.setActiveSheetIndex(sheetId);
      const sheet = workbook.getActiveSheet();
      const rangeInfo = parseRangeAddress(range);
      if (rangeInfo && sheet) {
        sheet.setSelection(
          rangeInfo.startRow,
          rangeInfo.startCol,
          rangeInfo.endRow - rangeInfo.startRow + 1,
          rangeInfo.endCol - rangeInfo.startCol + 1,
        );
      }
    },

    async getAllObjects(input) {
      const { sheetId, id } = input;
      const workbook = getWorkbook();
      if (!workbook) {
        return { objects: [] };
      }

      const objects: InferToolOutput<typeof tools.getAllObjects>["objects"] =
        [];

      const sheetsToSearch =
        sheetId !== undefined
          ? [workbook.getSheet(sheetId)]
          : Array.from({ length: workbook.getSheetCount() }, (_, i) =>
              workbook.getSheet(i),
            );

      for (const sheet of sheetsToSearch) {
        if (!sheet) continue;
        const currentSheetId = workbook.getSheetIndex(sheet.name());

        // Get charts if the charts API is available
        if (sheet.charts) {
          try {
            const charts = sheet.charts.all();
            for (const chart of charts) {
              const chartId = chart.name();
              if (id && chartId !== id) continue;

              const chartTypeMap: Record<number, string> = {
                0: "columnClustered",
                1: "columnStacked",
                2: "columnStacked100",
                3: "barClustered",
                4: "barStacked",
                5: "barStacked100",
                6: "line",
                7: "lineStacked",
                8: "lineStacked100",
                9: "lineMarkers",
                10: "area",
                11: "areaStacked",
                12: "areaStacked100",
                13: "pie",
                14: "doughnut",
                15: "scatter",
                16: "bubble",
                17: "radar",
                18: "radarFilled",
              };

              const chartType =
                chartTypeMap[chart.chartType()] || "columnClustered";
              const titleObj = chart.title();
              const title =
                titleObj && typeof titleObj === "object" && "text" in titleObj
                  ? String(titleObj.text)
                  : chartId;

              objects.push({
                id: chartId,
                type: "chart",
                sheetId: currentSheetId,
                chartType: chartType as "columnClustered",
                title,
                position: {
                  top: chart.y(),
                  left: chart.x(),
                },
              });
            }
          } catch {
            // Charts addon not available
          }
        }

        // Get pivot tables if the pivot API is available
        if (sheet.pivotTables) {
          try {
            const pivotTables = sheet.pivotTables.all();
            for (const pt of pivotTables) {
              const ptId = pt.name() as string | undefined;
              if (!ptId) continue;
              if (id && ptId !== id) continue;

              objects.push({
                id: ptId,
                type: "pivotTable",
                sheetId: currentSheetId,
                name: ptId,
                range: "",
                source: "",
                values: [{ field: "", summarizeBy: "sum" }],
              });
            }
          } catch {
            // Pivot addon not available
          }
        }
      }

      return { objects };
    },

    async modifyObject(input) {
      const { operation, sheetId, id, objectType, properties } = input;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const sheet = workbook.getSheet(sheetId);
      if (!sheet) throw new Error(`Sheet with ID ${sheetId} not found`);

      sheet.suspendPaint();
      try {
        if (objectType === "chart") {
          if (!sheet.charts) {
            throw new Error(
              "Charts addon is not loaded. Please ensure @mescius/spread-sheets-charts is imported.",
            );
          }

          switch (operation) {
            case "create": {
              const source = properties?.source as string | undefined;
              const chartTypeProp = properties?.chartType as string | undefined;
              const title = properties?.title as string | undefined;
              const anchor = properties?.anchor as string | undefined;

              if (!source || !chartTypeProp) {
                throw new Error(
                  "source and chartType are required for chart creation",
                );
              }

              const chartTypeMap: Record<string, number> = {
                columnClustered: 0,
                columnStacked: 1,
                columnStacked100: 2,
                barClustered: 3,
                barStacked: 4,
                barStacked100: 5,
                line: 6,
                lineStacked: 7,
                lineStacked100: 8,
                lineMarkers: 9,
                area: 10,
                areaStacked: 11,
                areaStacked100: 12,
                pie: 13,
                doughnut: 14,
                scatter: 15,
                bubble: 16,
                radar: 17,
                radarFilled: 18,
              };

              const chartType = chartTypeMap[chartTypeProp] ?? 0;
              const chartName = `Chart${Date.now()}`;

              // Parse anchor position
              let x = 100;
              let y = 100;
              if (anchor) {
                const anchorInfo = parseRangeAddress(anchor);
                if (anchorInfo) {
                  // Convert cell position to pixel coordinates (approximate)
                  x = anchorInfo.startCol * 64;
                  y = anchorInfo.startRow * 20;
                }
              }

              // Create chart with SpreadJS Charts API
              const chart = sheet.charts.add(
                chartName,
                chartType,
                x,
                y,
                400,
                300,
                source,
              );

              if (title) {
                chart.title({ text: title });
              }

              return { id: chartName };
            }

            case "delete": {
              if (!id) throw new Error("id is required for delete operation");
              sheet.charts.remove(id);
              return {};
            }

            case "update": {
              if (!id) throw new Error("id is required for update operation");
              const chart = sheet.charts.get(id);
              if (!chart) throw new Error(`Chart with ID ${id} not found`);

              const title = properties?.title as string | undefined;
              const anchor = properties?.anchor as string | undefined;

              if (title) {
                chart.title({ text: title });
              }

              if (anchor) {
                const anchorInfo = parseRangeAddress(anchor);
                if (anchorInfo) {
                  chart.x(anchorInfo.startCol * 64);
                  chart.y(anchorInfo.startRow * 20);
                }
              }

              return { id };
            }
          }
        } else if (objectType === "pivotTable") {
          if (!sheet.pivotTables) {
            throw new Error(
              "Pivot Table addon is not loaded. Please ensure @mescius/spread-sheets-pivot-addon is imported.",
            );
          }

          switch (operation) {
            case "create": {
              const source = properties?.source as string | undefined;
              const name = properties?.name as string | undefined;
              const range = properties?.range as string | undefined;
              const rows = properties?.rows as
                | Array<{ field: string }>
                | undefined;
              const columns = properties?.columns as
                | Array<{ field: string }>
                | undefined;
              const values = properties?.values as
                | Array<{ field: string; summarizeBy?: string }>
                | undefined;

              if (!source || !name) {
                throw new Error(
                  "source and name are required for pivot table creation",
                );
              }

              // Parse destination range
              let destRow = 0;
              let destCol = 0;
              if (range) {
                const rangeInfo = parseRangeAddress(range);
                if (rangeInfo) {
                  destRow = rangeInfo.startRow;
                  destCol = rangeInfo.startCol;
                }
              }

              // Create pivot table using SpreadJS Pivot API
              const pt = sheet.pivotTables.add(
                name,
                source,
                destRow,
                destCol,
                GC.Spread.Pivot.PivotTableLayoutType.outline,
                GC.Spread.Pivot.PivotTableThemes.light1,
              );

              if (pt) {
                // Add row fields
                if (rows) {
                  for (const row of rows) {
                    pt.add(
                      row.field,
                      name,
                      GC.Spread.Pivot.PivotTableFieldType.rowField,
                    );
                  }
                }

                // Add column fields
                if (columns) {
                  for (const col of columns) {
                    pt.add(
                      col.field,
                      name,
                      GC.Spread.Pivot.PivotTableFieldType.columnField,
                    );
                  }
                }

                // Add value fields
                if (values) {
                  for (const val of values) {
                    pt.add(
                      val.field,
                      name,
                      GC.Spread.Pivot.PivotTableFieldType.valueField,
                    );
                  }
                }
              }

              return { id: name };
            }

            case "delete": {
              if (!id) throw new Error("id is required for delete operation");
              sheet.pivotTables.remove(id);
              return {};
            }

            case "update": {
              if (!id) throw new Error("id is required for update operation");
              const pt = sheet.pivotTables.get(id);
              if (!pt) throw new Error(`PivotTable with ID ${id} not found`);

              // SpreadJS pivot tables don't support rename after creation
              return { id };
            }
          }
        }
      } finally {
        sheet.resumePaint();
      }

      return {};
    },

    async copyTo(input) {
      const { sheetId, sourceRange, destinationRange } = input;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const sheet = workbook.getSheet(sheetId);
      if (!sheet) throw new Error(`Sheet ${sheetId} not found`);

      const sourceInfo = parseRangeAddress(sourceRange);
      const destInfo = parseRangeAddress(destinationRange);

      if (!sourceInfo || !destInfo) {
        throw new Error("Invalid range format");
      }

      sheet.suspendPaint();
      try {
        // Use SpreadJS copyTo for proper formula translation
        sheet.copyTo(
          sourceInfo.startRow,
          sourceInfo.startCol,
          destInfo.startRow,
          destInfo.startCol,
          destInfo.endRow - destInfo.startRow + 1,
          destInfo.endCol - destInfo.startCol + 1,
          GC.Spread.Sheets.CopyToOptions.all,
        );
      } finally {
        sheet.resumePaint();
      }

      return {};
    },

    async clearCellRange(input) {
      const { sheetId, range, clearType = "contents" } = input;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const sheet = workbook.getSheet(sheetId);
      if (!sheet) throw new Error(`Sheet ${sheetId} not found`);

      const rangeInfo = parseRangeAddress(range);
      if (!rangeInfo) throw new Error(`Invalid range: ${range}`);

      sheet.suspendPaint();
      try {
        const cellRange = sheet.getRange(
          rangeInfo.startRow,
          rangeInfo.startCol,
          rangeInfo.endRow - rangeInfo.startRow + 1,
          rangeInfo.endCol - rangeInfo.startCol + 1,
        );

        if (clearType === "contents" || clearType === "all") {
          cellRange.clear(GC.Spread.Sheets.StorageType.data);
        }
        if (clearType === "formats" || clearType === "all") {
          cellRange.clear(GC.Spread.Sheets.StorageType.style);
        }
      } finally {
        sheet.resumePaint();
      }

      return {};
    },

    async resizeRange(input) {
      const { sheetId, range, width, height } = input;
      const workbook = getWorkbook();
      if (!workbook) throw new Error("Workbook not initialized");

      const sheet = workbook.getSheet(sheetId);
      if (!sheet) throw new Error(`Sheet ${sheetId} not found`);

      const rangeInfo = range ? parseRangeAddress(range) : null;
      const startCol = rangeInfo?.startCol ?? 0;
      const endCol = rangeInfo?.endCol ?? sheet.getColumnCount() - 1;
      const startRow = rangeInfo?.startRow ?? 0;
      const endRow = rangeInfo?.endRow ?? sheet.getRowCount() - 1;

      sheet.suspendPaint();
      try {
        if (width) {
          for (let col = startCol; col <= endCol; col++) {
            if (width.type === "autofit") {
              sheet.autoFitColumn(col);
            } else if (width.type === "points" && width.value !== undefined) {
              sheet.setColumnWidth(col, width.value);
            } else if (width.type === "standard") {
              sheet.setColumnWidth(col, 64);
            }
          }
        }

        if (height) {
          for (let row = startRow; row <= endRow; row++) {
            if (height.type === "autofit") {
              sheet.autoFitRow(row);
            } else if (height.type === "points" && height.value !== undefined) {
              sheet.setRowHeight(row, height.value);
            } else if (height.type === "standard") {
              sheet.setRowHeight(row, 20);
            }
          }
        }
      } finally {
        sheet.resumePaint();
      }

      return {};
    },
  };
}

// Helper Functions

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

function parseRangeAddress(range: string): {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
} | null {
  const match = range.match(/^([A-Z]+)?(\d+)?(?::([A-Z]+)?(\d+)?)?$/i);
  if (!match) return null;

  const [, startColStr, startRowStr, endColStr, endRowStr] = match;

  const startCol = startColStr ? letterToColumn(startColStr.toUpperCase()) : 0;
  const startRow = startRowStr ? parseInt(startRowStr, 10) - 1 : 0;
  const endCol = endColStr
    ? letterToColumn(endColStr.toUpperCase())
    : startColStr
      ? startCol
      : 25;
  const endRow = endRowStr
    ? parseInt(endRowStr, 10) - 1
    : startRowStr
      ? startRow
      : 999;

  return { startRow, startCol, endRow, endCol };
}

function parseFontString(font: string): {
  bold: boolean;
  italic: boolean;
  size: number | null;
  family: string | null;
} {
  const result = {
    bold: false,
    italic: false,
    size: null as number | null,
    family: null as string | null,
  };

  if (!font) return result;

  const parts = font.toLowerCase().split(" ");
  for (const part of parts) {
    if (part === "bold") {
      result.bold = true;
    } else if (part === "italic") {
      result.italic = true;
    } else if (part.endsWith("pt")) {
      result.size = parseFloat(part);
    }
  }

  // Extract font family (usually the last part)
  const familyMatch = font.match(/(\d+pt\s+)?(.+)$/i);
  if (familyMatch?.[2]) {
    result.family = familyMatch[2].trim();
  }

  return result;
}

function applyBorderStyles(
  sheet: Worksheet,
  row: number,
  col: number,
  borderStyles: {
    top?: { style?: string; weight?: string; color?: string };
    bottom?: { style?: string; weight?: string; color?: string };
    left?: { style?: string; weight?: string; color?: string };
    right?: { style?: string; weight?: string; color?: string };
  },
) {
  const getBorderLineStyle = (style?: string): GC.Spread.Sheets.LineStyle => {
    switch (style) {
      case "dashed":
        return GC.Spread.Sheets.LineStyle.dashed;
      case "dotted":
        return GC.Spread.Sheets.LineStyle.dotted;
      case "double":
        return GC.Spread.Sheets.LineStyle.double;
      case "solid":
      default:
        return GC.Spread.Sheets.LineStyle.thin;
    }
  };

  const cell = sheet.getCell(row, col);

  if (borderStyles.top) {
    cell.borderTop(
      new GC.Spread.Sheets.LineBorder(
        borderStyles.top.color || "#000000",
        getBorderLineStyle(borderStyles.top.style),
      ),
    );
  }
  if (borderStyles.bottom) {
    cell.borderBottom(
      new GC.Spread.Sheets.LineBorder(
        borderStyles.bottom.color || "#000000",
        getBorderLineStyle(borderStyles.bottom.style),
      ),
    );
  }
  if (borderStyles.left) {
    cell.borderLeft(
      new GC.Spread.Sheets.LineBorder(
        borderStyles.left.color || "#000000",
        getBorderLineStyle(borderStyles.left.style),
      ),
    );
  }
  if (borderStyles.right) {
    cell.borderRight(
      new GC.Spread.Sheets.LineBorder(
        borderStyles.right.color || "#000000",
        getBorderLineStyle(borderStyles.right.style),
      ),
    );
  }
}
