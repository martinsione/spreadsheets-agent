import type { InferToolInput, InferToolOutput } from "ai";
import type { tools } from "@/server/ai/tools";
import type { Sheet } from "@/spreadsheet-service";

export async function getSheets(): Promise<Sheet[]> {
  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const usedRanges = worksheets.items.map((ws) => {
      const range = ws.getUsedRangeOrNullObject(true);
      range.load(["rowCount", "columnCount"]);
      return range;
    });

    await context.sync();

    return worksheets.items.map((ws, i) => {
      const range = usedRanges[i];
      return {
        id: ws.position,
        name: ws.name,
        maxRows: range?.isNullObject === false ? range.rowCount : 0,
        maxColumns: range?.isNullObject === false ? range.columnCount : 0,
      };
    });
  });
}

export async function getCellRanges(
  input: InferToolInput<typeof tools.getCellRanges>,
): Promise<InferToolOutput<typeof tools.getCellRanges>> {
  const { sheetId, ranges, includeStyles = true, cellLimit = 10000 } = input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const worksheet = worksheets.items.find((ws) => ws.position === sheetId);
    if (!worksheet) {
      throw new Error(`Sheet with ID ${sheetId} not found`);
    }

    const usedRange = worksheet.getUsedRangeOrNullObject(true);
    usedRange.load("address");
    await context.sync();

    const dimension = usedRange.isNullObject
      ? undefined
      : usedRange.address?.split("!")[1];

    const cells: Record<string, unknown> = {};
    const styles: Record<string, Record<string, unknown>> = {};
    const borders: Record<string, string | Record<string, string>> = {};
    let totalCellCount = 0;
    let hasMore = false;
    const nextRanges: string[] = [];

    for (let i = 0; i < ranges.length; i++) {
      const rangeAddress = ranges[i];

      if (totalCellCount >= cellLimit) {
        nextRanges.push(...ranges.slice(i));
        hasMore = true;
        break;
      }

      const range = worksheet.getRange(rangeAddress);
      range.load([
        "values",
        "formulas",
        "address",
        "rowCount",
        "columnCount",
        "rowIndex",
        "columnIndex",
      ]);

      if (includeStyles) {
        range.format.load(["horizontalAlignment"]);
        range.format.font.load([
          "bold",
          "italic",
          "underline",
          "strikethrough",
          "size",
          "color",
          "name",
        ]);
        range.format.fill.load(["color"]);
      }

      await context.sync();

      const values = range.values;
      const formulas = range.formulas;
      const startRow = range.rowIndex;
      const startCol = range.columnIndex;

      for (let row = 0; row < range.rowCount; row++) {
        if (totalCellCount >= cellLimit) {
          hasMore = true;
          const remainingRows = range.rowCount - row;
          if (remainingRows > 0) {
            const startCellRow = startRow + row + 1;
            const endCellRow = startRow + range.rowCount;
            const startColLetter = columnToLetter(startCol);
            const endColLetter = columnToLetter(
              startCol + range.columnCount - 1,
            );
            nextRanges.push(
              `${startColLetter}${startCellRow}:${endColLetter}${endCellRow}`,
            );
          }
          nextRanges.push(...ranges.slice(i + 1));
          break;
        }

        for (let col = 0; col < range.columnCount; col++) {
          const value = values[row]?.[col];
          const formula = formulas[row]?.[col];

          if (value === "" && (!formula || formula === "")) continue;

          const cellAddress = `${columnToLetter(startCol + col)}${startRow + row + 1}`;

          if (
            formula &&
            typeof formula === "string" &&
            formula.startsWith("=")
          ) {
            cells[cellAddress] = [value, formula];
          } else {
            cells[cellAddress] = value;
          }

          totalCellCount++;
        }
      }
    }

    const result: InferToolOutput<typeof tools.getCellRanges> = {
      worksheet: { name: worksheet.name, sheetId, dimension, cells },
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
  });
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

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const sheetsToSearch =
      sheetId !== undefined
        ? worksheets.items.filter((ws) => ws.position === sheetId)
        : worksheets.items;

    if (sheetsToSearch.length === 0) {
      return {
        matches: [],
        totalFound: 0,
        returned: 0,
        offset,
        hasMore: false,
        searchTerm,
        searchScope:
          sheetId !== undefined ? `Sheet ID ${sheetId}` : "All sheets",
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

    for (const worksheet of sheetsToSearch) {
      const searchRange = range
        ? worksheet.getRange(range)
        : worksheet.getUsedRangeOrNullObject(true);

      searchRange.load("isNullObject");
      await context.sync();

      if (searchRange.isNullObject) continue;

      let foundRange = searchRange.findOrNullObject(searchTerm, {
        completeMatch: matchEntireCell,
        matchCase,
        searchDirection: Excel.SearchDirection.forward,
      });

      foundRange.load([
        "address",
        "values",
        "formulas",
        "rowIndex",
        "columnIndex",
      ]);
      await context.sync();

      const visitedAddresses = new Set<string>();

      const firstAddress = foundRange.isNullObject ? null : foundRange.address;

      while (!foundRange.isNullObject) {
        const address = foundRange.address;
        if (visitedAddresses.has(address)) break;
        visitedAddresses.add(address);

        if (matches.length >= offset + maxResults) break;

        if (matches.length >= offset) {
          const cellAddress = address.split("!")[1] || address;
          matches.push({
            sheetName: worksheet.name,
            sheetId: worksheet.position,
            a1: cellAddress,
            value: foundRange.values[0]?.[0] ?? null,
            formula: foundRange.formulas[0]?.[0]?.startsWith("=")
              ? foundRange.formulas[0][0]
              : null,
            row: foundRange.rowIndex + 1,
            column: foundRange.columnIndex + 1,
          });
        }

        // Find next match by searching again from the beginning and skipping visited
        foundRange = searchRange.findOrNullObject(searchTerm, {
          completeMatch: matchEntireCell,
          matchCase,
          searchDirection: Excel.SearchDirection.forward,
        });

        foundRange.load([
          "address",
          "values",
          "formulas",
          "rowIndex",
          "columnIndex",
          "isNullObject",
        ]);
        await context.sync();

        // If we've cycled back to the first match, we're done
        if (!foundRange.isNullObject && foundRange.address === firstAddress) {
          break;
        }
      }
    }

    const totalFound = matches.length + offset;
    const returnedMatches = matches.slice(0, maxResults);
    const hasMore = matches.length > maxResults;

    return {
      matches: returnedMatches,
      totalFound,
      returned: returnedMatches.length,
      offset,
      hasMore,
      searchTerm,
      searchScope: sheetId !== undefined ? `Sheet ID ${sheetId}` : "All sheets",
      nextOffset: hasMore ? offset + maxResults : null,
    };
  });
}

export async function setCellRange(
  input: InferToolInput<typeof tools.setCellRange>,
): Promise<InferToolOutput<typeof tools.setCellRange>> {
  const { sheetId, range, cells, copyToRange, resizeWidth, resizeHeight } =
    input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const worksheet = worksheets.items.find((ws) => ws.position === sheetId);
    if (!worksheet) {
      throw new Error(`Sheet with ID ${sheetId} not found`);
    }

    const targetRange = worksheet.getRange(range);
    targetRange.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await context.sync();

    const rowCount = targetRange.rowCount;
    const colCount = targetRange.columnCount;

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
    targetRange.values = values.map((row, r) => {
      const formulaRow = formulas[r];
      return row.map((val, c) => (formulaRow?.[c] ? "" : val));
    });

    // Set formulas where applicable
    if (hasFormulas) {
      for (let r = 0; r < rowCount; r++) {
        const formulaRow = formulas[r];
        if (!formulaRow) continue;
        for (let c = 0; c < colCount; c++) {
          const formula = formulaRow[c];
          if (formula) {
            const cellRange = targetRange.getCell(r, c);
            cellRange.formulas = [[formula]];
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
        if (cell.cellStyles) {
          const cellRange = targetRange.getCell(r, c);
          const styles = cell.cellStyles;

          if (styles.fontColor) {
            cellRange.format.font.color = styles.fontColor;
          }
          if (styles.fontSize) {
            cellRange.format.font.size = styles.fontSize;
          }
          if (styles.fontFamily) {
            cellRange.format.font.name = styles.fontFamily;
          }
          if (styles.fontWeight === "bold") {
            cellRange.format.font.bold = true;
          } else if (styles.fontWeight === "normal") {
            cellRange.format.font.bold = false;
          }
          if (styles.fontStyle === "italic") {
            cellRange.format.font.italic = true;
          } else if (styles.fontStyle === "normal") {
            cellRange.format.font.italic = false;
          }
          if (styles.fontLine === "underline") {
            cellRange.format.font.underline = Excel.RangeUnderlineStyle.single;
          } else if (styles.fontLine === "line-through") {
            cellRange.format.font.strikethrough = true;
          } else if (styles.fontLine === "none") {
            cellRange.format.font.underline = Excel.RangeUnderlineStyle.none;
            cellRange.format.font.strikethrough = false;
          }
          if (styles.backgroundColor) {
            cellRange.format.fill.color = styles.backgroundColor;
          }
          if (styles.horizontalAlignment) {
            cellRange.format.horizontalAlignment =
              styles.horizontalAlignment === "left"
                ? Excel.HorizontalAlignment.left
                : styles.horizontalAlignment === "center"
                  ? Excel.HorizontalAlignment.center
                  : Excel.HorizontalAlignment.right;
          }
          if (styles.numberFormat) {
            cellRange.numberFormat = [[styles.numberFormat]];
          }
        }

        // Apply border styles
        if (cell.borderStyles) {
          const cellRange = targetRange.getCell(r, c);
          const borders = cell.borderStyles;

          const applyBorder = (
            border: Excel.RangeBorder,
            style: { style?: string; weight?: string; color?: string },
          ) => {
            if (style.color) {
              border.color = style.color;
            }
            if (style.style) {
              border.style =
                style.style === "solid"
                  ? Excel.BorderLineStyle.continuous
                  : style.style === "dashed"
                    ? Excel.BorderLineStyle.dash
                    : style.style === "dotted"
                      ? Excel.BorderLineStyle.dot
                      : Excel.BorderLineStyle.double;
            }
            if (style.weight) {
              border.weight =
                style.weight === "thin"
                  ? Excel.BorderWeight.thin
                  : style.weight === "medium"
                    ? Excel.BorderWeight.medium
                    : Excel.BorderWeight.thick;
            }
          };

          if (borders.top) {
            applyBorder(
              cellRange.format.borders.getItem(Excel.BorderIndex.edgeTop),
              borders.top,
            );
          }
          if (borders.bottom) {
            applyBorder(
              cellRange.format.borders.getItem(Excel.BorderIndex.edgeBottom),
              borders.bottom,
            );
          }
          if (borders.left) {
            applyBorder(
              cellRange.format.borders.getItem(Excel.BorderIndex.edgeLeft),
              borders.left,
            );
          }
          if (borders.right) {
            applyBorder(
              cellRange.format.borders.getItem(Excel.BorderIndex.edgeRight),
              borders.right,
            );
          }
        }

        // Apply notes
        if (cell.note) {
          const noteCellRange = targetRange.getCell(r, c);
          noteCellRange.load("address");
          await context.sync();
          const comment = worksheet.comments.add(noteCellRange, cell.note);
          comment.load("id");
        }
      }
    }

    await context.sync();

    // Handle copyToRange if specified
    if (copyToRange) {
      const destRange = worksheet.getRange(copyToRange);
      destRange.copyFrom(targetRange, Excel.RangeCopyType.all);
      await context.sync();
    }

    // Handle resizeWidth
    if (resizeWidth) {
      if (resizeWidth.type === "autofit") {
        targetRange.format.autofitColumns();
      } else if (
        resizeWidth.type === "points" &&
        resizeWidth.value !== undefined
      ) {
        targetRange.format.columnWidth = resizeWidth.value;
      } else if (resizeWidth.type === "standard") {
        targetRange.format.columnWidth = 64; // Default Excel column width
      }
    }

    // Handle resizeHeight
    if (resizeHeight) {
      if (resizeHeight.type === "autofit") {
        targetRange.format.autofitRows();
      } else if (
        resizeHeight.type === "points" &&
        resizeHeight.value !== undefined
      ) {
        targetRange.format.rowHeight = resizeHeight.value;
      } else if (resizeHeight.type === "standard") {
        targetRange.format.rowHeight = 15; // Default Excel row height
      }
    }

    await context.sync();

    // Get formula results for cells with formulas
    const formulaResults: Record<string, number | string> = {};
    if (hasFormulas) {
      targetRange.load(["values"]);
      await context.sync();

      const startRow = targetRange.rowIndex;
      const startCol = targetRange.columnIndex;

      for (let r = 0; r < rowCount; r++) {
        const formulaRow = formulas[r];
        if (!formulaRow) continue;
        for (let c = 0; c < colCount; c++) {
          if (formulaRow[c]) {
            const cellAddress = `${columnToLetter(startCol + c)}${startRow + r + 1}`;
            const valueRow = targetRange.values[r];
            const value = valueRow?.[c];
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
  });
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

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const worksheet = worksheets.items.find((ws) => ws.position === sheetId);
    if (!worksheet) {
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
          const startRow = position === "after" ? rowNum + 1 : rowNum;
          const rangeAddress = `${startRow}:${startRow + count - 1}`;
          const range = worksheet.getRange(rangeAddress);
          range.insert(Excel.InsertShiftDirection.down);
        } else {
          const colLetter = reference.toUpperCase();
          const colIndex = letterToColumn(colLetter);
          const startCol = position === "after" ? colIndex + 1 : colIndex;
          const endCol = startCol + count - 1;
          const rangeAddress = `${columnToLetter(startCol)}:${columnToLetter(endCol)}`;
          const range = worksheet.getRange(rangeAddress);
          range.insert(Excel.InsertShiftDirection.right);
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
          const rangeAddress = `${rowNum}:${rowNum + count - 1}`;
          const range = worksheet.getRange(rangeAddress);
          range.delete(Excel.DeleteShiftDirection.up);
        } else {
          const colLetter = reference.toUpperCase();
          const colIndex = letterToColumn(colLetter);
          const endCol = colIndex + count - 1;
          const rangeAddress = `${colLetter}:${columnToLetter(endCol)}`;
          const range = worksheet.getRange(rangeAddress);
          range.delete(Excel.DeleteShiftDirection.left);
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
          const rangeAddress = `${rowNum}:${rowNum + count - 1}`;
          const range = worksheet.getRange(rangeAddress);
          range.rowHidden = true;
        } else {
          const colLetter = reference.toUpperCase();
          const colIndex = letterToColumn(colLetter);
          const endCol = colIndex + count - 1;
          const rangeAddress = `${colLetter}:${columnToLetter(endCol)}`;
          const range = worksheet.getRange(rangeAddress);
          range.columnHidden = true;
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
          const rangeAddress = `${rowNum}:${rowNum + count - 1}`;
          const range = worksheet.getRange(rangeAddress);
          range.rowHidden = false;
        } else {
          const colLetter = reference.toUpperCase();
          const colIndex = letterToColumn(colLetter);
          const endCol = colIndex + count - 1;
          const rangeAddress = `${colLetter}:${columnToLetter(endCol)}`;
          const range = worksheet.getRange(rangeAddress);
          range.columnHidden = false;
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
          worksheet.freezePanes.freezeRows(count);
        } else {
          worksheet.freezePanes.freezeColumns(count);
        }
        break;
      }

      case "unfreeze": {
        worksheet.freezePanes.unfreeze();
        break;
      }
    }

    await context.sync();
    return {};
  });
}

export async function modifyWorkbookStructure(
  input: InferToolInput<typeof tools.modifyWorkbookStructure>,
): Promise<InferToolOutput<typeof tools.modifyWorkbookStructure>> {
  const { operation, sheetName, sheetId, newName, tabColor } = input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    switch (operation) {
      case "create": {
        if (!sheetName) {
          throw new Error("sheetName is required for create operation");
        }

        const newSheet = worksheets.add(sheetName);
        newSheet.load(["name", "position"]);

        if (tabColor) {
          newSheet.tabColor = tabColor;
        }

        await context.sync();

        return {
          sheetId: newSheet.position,
          sheetName: newSheet.name,
          message: `Sheet "${sheetName}" created successfully`,
        };
      }

      case "delete": {
        if (sheetId === undefined) {
          throw new Error("sheetId is required for delete operation");
        }

        const worksheet = worksheets.items.find(
          (ws) => ws.position === sheetId,
        );
        if (!worksheet) {
          throw new Error(`Sheet with ID ${sheetId} not found`);
        }

        const deletedName = worksheet.name;
        worksheet.delete();
        await context.sync();

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

        const worksheet = worksheets.items.find(
          (ws) => ws.position === sheetId,
        );
        if (!worksheet) {
          throw new Error(`Sheet with ID ${sheetId} not found`);
        }

        const oldName = worksheet.name;
        worksheet.name = newName;
        await context.sync();

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

        const worksheet = worksheets.items.find(
          (ws) => ws.position === sheetId,
        );
        if (!worksheet) {
          throw new Error(`Sheet with ID ${sheetId} not found`);
        }

        const copiedSheet = worksheet.copy();
        copiedSheet.load(["name", "position"]);
        await context.sync();

        if (newName) {
          copiedSheet.name = newName;
          await context.sync();
        }

        return {
          sheetId: copiedSheet.position,
          sheetName: copiedSheet.name,
          message: `Sheet duplicated successfully`,
        };
      }

      default:
        throw new Error(`Unknown operation: ${operation}`);
    }
  });
}

export async function copyTo(
  input: InferToolInput<typeof tools.copyTo>,
): Promise<InferToolOutput<typeof tools.copyTo>> {
  const { sheetId, sourceRange, destinationRange } = input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const worksheet = worksheets.items.find((ws) => ws.position === sheetId);
    if (!worksheet) {
      throw new Error(`Sheet with ID ${sheetId} not found`);
    }

    const source = worksheet.getRange(sourceRange);
    const destination = worksheet.getRange(destinationRange);

    destination.copyFrom(source, Excel.RangeCopyType.all);
    await context.sync();

    return {};
  });
}

export async function getAllObjects(
  input: InferToolInput<typeof tools.getAllObjects>,
): Promise<InferToolOutput<typeof tools.getAllObjects>> {
  const { sheetId, id } = input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const sheetsToSearch =
      sheetId !== undefined
        ? worksheets.items.filter((ws) => ws.position === sheetId)
        : worksheets.items;

    const objects: InferToolOutput<typeof tools.getAllObjects>["objects"] = [];

    for (const worksheet of sheetsToSearch) {
      // Load pivot tables
      const pivotTables = worksheet.pivotTables;
      pivotTables.load("items/name,items/id");
      await context.sync();

      for (const pivotTable of pivotTables.items) {
        if (id && pivotTable.id !== id) continue;

        pivotTable.load(["name", "id"]);
        const layout = pivotTable.layout;
        const rangeObj = layout.getRange();
        rangeObj.load("address");

        await context.sync();

        objects.push({
          id: pivotTable.id,
          type: "pivotTable",
          sheetId: worksheet.position,
          name: pivotTable.name,
          range: rangeObj.address?.split("!")[1] || "",
          source: "", // Source range is not directly accessible after creation
          values: [{ field: "", summarizeBy: "sum" }], // Placeholder - PivotTable field structure is complex
        });
      }

      // Load charts
      const charts = worksheet.charts;
      charts.load("items/name,items/id");
      await context.sync();

      for (const chart of charts.items) {
        if (id && chart.id !== id) continue;

        chart.load(["name", "id", "chartType", "top", "left", "title/text"]);
        await context.sync();

        const chartTypeMap: Record<string, string> = {
          ColumnClustered: "columnClustered",
          ColumnStacked: "columnStacked",
          ColumnStacked100: "columnStacked100",
          "3DColumn": "column3D",
          BarClustered: "barClustered",
          BarStacked: "barStacked",
          BarStacked100: "barStacked100",
          "3DBar": "bar3D",
          Line: "line",
          LineMarkers: "lineMarkers",
          LineStacked: "lineStacked",
          LineStacked100: "lineStacked100",
          "3DLine": "line3D",
          Area: "area",
          AreaStacked: "areaStacked",
          AreaStacked100: "areaStacked100",
          "3DArea": "area3D",
          Pie: "pie",
          PieExploded: "pieExploded",
          "3DPie": "pie3D",
          Doughnut: "doughnut",
          DoughnutExploded: "doughnutExploded",
          XYScatter: "scatter",
          XYScatterLines: "scatterLines",
          XYScatterLinesNoMarkers: "scatterLinesMarkers",
          Radar: "radar",
          RadarMarkers: "radarMarkers",
          RadarFilled: "radarFilled",
          Bubble: "bubble",
        };

        const chartType = chartTypeMap[chart.chartType] || "columnClustered";

        // Load series data
        const seriesCollection = chart.series;
        seriesCollection.load("items");
        await context.sync();

        const readOnlySeries: Array<{
          name: string;
          values: string;
          categories?: string;
        }> = [];

        for (const series of seriesCollection.items) {
          series.load(["name"]);
          await context.sync();

          readOnlySeries.push({
            name: series.name || "",
            values: "", // Values range not directly accessible
          });
        }

        objects.push({
          id: chart.id,
          type: "chart",
          sheetId: worksheet.position,
          chartType: chartType as "columnClustered",
          title: chart.title?.text || chart.name,
          position: {
            top: chart.top,
            left: chart.left,
          },
          readOnlySeries:
            readOnlySeries.length > 0 ? readOnlySeries : undefined,
        });
      }
    }

    return { objects };
  });
}

export async function modifyObject(
  input: InferToolInput<typeof tools.modifyObject>,
): Promise<InferToolOutput<typeof tools.modifyObject>> {
  const { operation, sheetId, id, objectType, properties } = input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const worksheet = worksheets.items.find((ws) => ws.position === sheetId);
    if (!worksheet) {
      throw new Error(`Sheet with ID ${sheetId} not found`);
    }

    if (objectType === "pivotTable") {
      switch (operation) {
        case "create": {
          const source =
            properties &&
            "source" in properties &&
            typeof properties.source === "string"
              ? properties.source
              : undefined;
          const name =
            properties &&
            "name" in properties &&
            typeof properties.name === "string"
              ? properties.name
              : undefined;
          const rangeStr =
            properties &&
            "range" in properties &&
            typeof properties.range === "string"
              ? properties.range
              : undefined;
          const rows =
            properties && "rows" in properties ? properties.rows : undefined;
          const columns =
            properties && "columns" in properties
              ? properties.columns
              : undefined;
          const values =
            properties && "values" in properties
              ? properties.values
              : undefined;

          if (!source || !name) {
            throw new Error(
              "source and name are required for pivot table creation",
            );
          }

          const sourceRange = worksheet.getRange(source);
          const destCell = rangeStr
            ? worksheet.getRange(rangeStr)
            : worksheet.getRange("A1");

          const pivotTable = worksheet.pivotTables.add(
            name,
            sourceRange,
            destCell,
          );

          pivotTable.load("id");
          await context.sync();

          // Add row fields
          if (rows && Array.isArray(rows)) {
            for (const row of rows) {
              if (row && typeof row === "object" && "field" in row) {
                const field = pivotTable.hierarchies.getItemOrNullObject(
                  row.field,
                );
                field.load("isNullObject");
                await context.sync();
                if (!field.isNullObject) {
                  pivotTable.rowHierarchies.add(field);
                }
              }
            }
          }

          // Add column fields
          if (columns && Array.isArray(columns)) {
            for (const col of columns) {
              if (col && typeof col === "object" && "field" in col) {
                const field = pivotTable.hierarchies.getItemOrNullObject(
                  col.field,
                );
                field.load("isNullObject");
                await context.sync();
                if (!field.isNullObject) {
                  pivotTable.columnHierarchies.add(field);
                }
              }
            }
          }

          // Add value fields
          if (values && Array.isArray(values)) {
            for (const val of values) {
              if (val && typeof val === "object" && "field" in val) {
                const field = pivotTable.hierarchies.getItemOrNullObject(
                  val.field,
                );
                field.load("isNullObject");
                await context.sync();
                if (!field.isNullObject) {
                  const dataField = pivotTable.dataHierarchies.add(field);
                  const summarizeBy =
                    "summarizeBy" in val ? val.summarizeBy : undefined;
                  if (summarizeBy && typeof summarizeBy === "string") {
                    const summarizeFunctions: Record<
                      string,
                      Excel.AggregationFunction
                    > = {
                      sum: Excel.AggregationFunction.sum,
                      count: Excel.AggregationFunction.count,
                      average: Excel.AggregationFunction.average,
                      max: Excel.AggregationFunction.max,
                      min: Excel.AggregationFunction.min,
                      product: Excel.AggregationFunction.product,
                      countNums: Excel.AggregationFunction.countNumbers,
                      stdDev: Excel.AggregationFunction.standardDeviation,
                      stdDevp: Excel.AggregationFunction.standardDeviationP,
                      var: Excel.AggregationFunction.variance,
                      varp: Excel.AggregationFunction.varianceP,
                    };
                    dataField.summarizeBy =
                      summarizeFunctions[summarizeBy] ||
                      Excel.AggregationFunction.sum;
                  }
                }
              }
            }
          }

          await context.sync();
          return { id: pivotTable.id };
        }

        case "delete": {
          if (!id) {
            throw new Error("id is required for delete operation");
          }

          const pivotTable = worksheet.pivotTables.getItemOrNullObject(id);
          pivotTable.load("isNullObject");
          await context.sync();

          if (pivotTable.isNullObject) {
            throw new Error(`PivotTable with ID ${id} not found`);
          }

          pivotTable.delete();
          await context.sync();
          return {};
        }

        case "update": {
          if (!id) {
            throw new Error("id is required for update operation");
          }

          const pivotTable = worksheet.pivotTables.getItemOrNullObject(id);
          pivotTable.load("isNullObject");
          await context.sync();

          if (pivotTable.isNullObject) {
            throw new Error(`PivotTable with ID ${id} not found`);
          }

          const name =
            properties && "name" in properties ? properties.name : undefined;
          if (name && typeof name === "string") {
            pivotTable.name = name;
          }

          pivotTable.refresh();
          await context.sync();
          return { id };
        }
      }
    } else if (objectType === "chart") {
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

          const chartTypeMap: Record<string, Excel.ChartType> = {
            columnClustered: Excel.ChartType.columnClustered,
            columnStacked: Excel.ChartType.columnStacked,
            columnStacked100: Excel.ChartType.columnStacked100,
            column3D: Excel.ChartType._3DColumnClustered,
            barClustered: Excel.ChartType.barClustered,
            barStacked: Excel.ChartType.barStacked,
            barStacked100: Excel.ChartType.barStacked100,
            bar3D: Excel.ChartType._3DBarClustered,
            line: Excel.ChartType.line,
            lineMarkers: Excel.ChartType.lineMarkers,
            lineStacked: Excel.ChartType.lineStacked,
            lineStacked100: Excel.ChartType.lineStacked100,
            line3D: Excel.ChartType._3DLine,
            area: Excel.ChartType.area,
            areaStacked: Excel.ChartType.areaStacked,
            areaStacked100: Excel.ChartType.areaStacked100,
            area3D: Excel.ChartType._3DArea,
            pie: Excel.ChartType.pie,
            pieExploded: Excel.ChartType.pieExploded,
            pie3D: Excel.ChartType._3DPie,
            doughnut: Excel.ChartType.doughnut,
            doughnutExploded: Excel.ChartType.doughnutExploded,
            scatter: Excel.ChartType.xyscatter,
            scatterLines: Excel.ChartType.xyscatterLines,
            scatterLinesMarkers: Excel.ChartType.xyscatterLinesNoMarkers,
            radar: Excel.ChartType.radar,
            radarMarkers: Excel.ChartType.radarMarkers,
            radarFilled: Excel.ChartType.radarFilled,
            bubble: Excel.ChartType.bubble,
          };

          const sourceRange = worksheet.getRange(source);
          const excelChartType: Excel.ChartType =
            chartTypeMap[chartTypeProp] ?? Excel.ChartType.columnClustered;

          const chart = worksheet.charts.add(excelChartType, sourceRange);
          chart.load("id");

          if (title && typeof title === "string") {
            chart.title.text = title;
          }

          if (anchor && typeof anchor === "string") {
            const anchorRange = worksheet.getRange(anchor);
            chart.setPosition(anchorRange);
          }

          await context.sync();
          return { id: chart.id };
        }

        case "delete": {
          if (!id) {
            throw new Error("id is required for delete operation");
          }

          const chart = worksheet.charts.getItemOrNullObject(id);
          chart.load("isNullObject");
          await context.sync();

          if (chart.isNullObject) {
            throw new Error(`Chart with ID ${id} not found`);
          }

          chart.delete();
          await context.sync();
          return {};
        }

        case "update": {
          if (!id) {
            throw new Error("id is required for update operation");
          }

          const chart = worksheet.charts.getItemOrNullObject(id);
          chart.load("isNullObject");
          await context.sync();

          if (chart.isNullObject) {
            throw new Error(`Chart with ID ${id} not found`);
          }

          const title =
            properties && "title" in properties ? properties.title : undefined;
          const anchor =
            properties && "anchor" in properties
              ? properties.anchor
              : undefined;
          const chartTypeProp =
            properties && "chartType" in properties
              ? properties.chartType
              : undefined;

          if (title && typeof title === "string") {
            chart.title.text = title;
          }

          if (anchor && typeof anchor === "string") {
            const anchorRange = worksheet.getRange(anchor);
            chart.setPosition(anchorRange);
          }

          if (chartTypeProp && typeof chartTypeProp === "string") {
            const chartTypeMap: Record<string, Excel.ChartType> = {
              columnClustered: Excel.ChartType.columnClustered,
              columnStacked: Excel.ChartType.columnStacked,
              columnStacked100: Excel.ChartType.columnStacked100,
              line: Excel.ChartType.line,
              pie: Excel.ChartType.pie,
              barClustered: Excel.ChartType.barClustered,
            };
            chart.chartType =
              chartTypeMap[chartTypeProp] || Excel.ChartType.columnClustered;
          }

          await context.sync();
          return { id };
        }
      }
    }

    throw new Error(`Unknown object type: ${objectType}`);
  });
}

export async function resizeRange(
  input: InferToolInput<typeof tools.resizeRange>,
): Promise<InferToolOutput<typeof tools.resizeRange>> {
  const { sheetId, range, width, height } = input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const worksheet = worksheets.items.find((ws) => ws.position === sheetId);
    if (!worksheet) {
      throw new Error(`Sheet with ID ${sheetId} not found`);
    }

    const targetRange = range
      ? worksheet.getRange(range)
      : worksheet.getUsedRangeOrNullObject(true);

    targetRange.load("isNullObject");
    await context.sync();

    if (targetRange.isNullObject) {
      return {};
    }

    if (width) {
      if (width.type === "autofit") {
        targetRange.format.autofitColumns();
      } else if (width.type === "points" && width.value !== undefined) {
        targetRange.format.columnWidth = width.value;
      } else if (width.type === "standard") {
        targetRange.format.columnWidth = 64; // Default Excel column width
      }
    }

    if (height) {
      if (height.type === "autofit") {
        targetRange.format.autofitRows();
      } else if (height.type === "points" && height.value !== undefined) {
        targetRange.format.rowHeight = height.value;
      } else if (height.type === "standard") {
        targetRange.format.rowHeight = 15; // Default Excel row height
      }
    }

    await context.sync();
    return {};
  });
}

export async function clearCellRange(
  input: InferToolInput<typeof tools.clearCellRange>,
): Promise<InferToolOutput<typeof tools.clearCellRange>> {
  const { sheetId, range, clearType = "contents" } = input;

  return await Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name,items/position");
    await context.sync();

    const worksheet = worksheets.items.find((ws) => ws.position === sheetId);
    if (!worksheet) {
      throw new Error(`Sheet with ID ${sheetId} not found`);
    }

    const targetRange = worksheet.getRange(range);

    const clearApplyTo =
      clearType === "all"
        ? Excel.ClearApplyTo.all
        : clearType === "formats"
          ? Excel.ClearApplyTo.formats
          : Excel.ClearApplyTo.contents;

    targetRange.clear(clearApplyTo);
    await context.sync();

    return {};
  });
}

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
