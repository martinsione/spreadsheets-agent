import { anthropic } from "@ai-sdk/anthropic";
import { tool } from "ai";
import * as z from "zod";

const DEFAULT_MAX_CELLS = 10000;

/**
 * Validates whether a string is a valid Excel/spreadsheet cell range notation.
 * Supports various formats: A1, A1:B2, A:C, 1:10, etc.
 */
function isValidCellRange(input: string): boolean {
  if (!input || typeof input !== "string") return false;

  const normalized = input.replace(/\s/g, "");
  const cellOrCellRange = /^[A-Z]+\d+(:[A-Z]+\d+)?$/; // Cell or cell-to-cell range: "A1" or "A1:B2"
  const columnRange = /^[A-Z]+:[A-Z]+$/; // Column-to-column range: "A:C"
  const rowRange = /^\d+:\d+$/; // Row-to-row range: "1:10"
  const cellToColumnRange = /^[A-Z]+\d+:[A-Z]+$/; // Cell-to-column range: "A1:B"
  const columnToRowRange = /^[A-Z]+:\d+$/; // Column-to-row range: "A:1"
  const rowToCellRange = /^\d+:[A-Z]+\d+$/; // Row-to-cell range: "1:A1"

  return (
    cellOrCellRange.test(normalized) ||
    columnRange.test(normalized) ||
    rowRange.test(normalized) ||
    cellToColumnRange.test(normalized) ||
    columnToRowRange.test(normalized) ||
    rowToCellRange.test(normalized)
  );
}

const ExplanationSchema = z
  .string()
  .max(50)
  .describe(
    "A very brief description of the action (max 50 chars). Shown next to the tool icon.",
  );

const BaseSpreadsheetObjectSchema = z.object({
  id: z.string().describe("Unique object identifier"),
  type: z.enum(["pivotTable", "chart", "table"]).describe("Type of object"),
  sheetId: z.number().describe("ID of the sheet where the object is located"),
});

const PivotFieldReferenceSchema = z.object({
  field: z
    .string()
    .describe("Field name from source data (OOXML field reference)"),
});

const PivotTableSchema = BaseSpreadsheetObjectSchema.extend({
  type: z.literal("pivotTable"),
  name: z.string().describe("PivotTable name"),
  range: z
    .string()
    .describe(
      "A1 notation of top-left cell where pivot table will be placed (e.g., 'F1')",
    ),
  source: z
    .string()
    .describe(
      "Range address to source data in format 'SheetName!StartCell:EndCell' (e.g., 'Sheet1!A1:D100')",
    ),
  rows: z
    .array(PivotFieldReferenceSchema)
    .optional()
    .describe("Fields to group by vertically"),
  columns: z
    .array(PivotFieldReferenceSchema)
    .optional()
    .describe("Fields to group by horizontally"),
  values: z
    .array(
      z.object({
        field: z.string().describe("Field name from source data to aggregate"),
        summarizeBy: z
          .enum([
            "sum",
            "count",
            "average",
            "max",
            "min",
            "product",
            "countNums",
            "stdDev",
            "stdDevp",
            "var",
            "varp",
          ])
          .optional()
          .default("sum")
          .describe("Aggregation function (OOXML ST_DataConsolidateFunction)"),
      }),
    )
    .min(1)
    .describe("Fields to aggregate (at least one required)"),
});

const ChartSchema = BaseSpreadsheetObjectSchema.extend({
  type: z.literal("chart"),
  chartType: z
    .enum([
      "columnClustered",
      "columnStacked",
      "columnStacked100",
      "column3D",
      "barClustered",
      "barStacked",
      "barStacked100",
      "bar3D",
      "line",
      "lineMarkers",
      "lineStacked",
      "lineStacked100",
      "line3D",
      "area",
      "areaStacked",
      "areaStacked100",
      "area3D",
      "pie",
      "pieExploded",
      "pie3D",
      "doughnut",
      "doughnutExploded",
      "scatter",
      "scatterLines",
      "scatterLinesMarkers",
      "radar",
      "radarMarkers",
      "radarFilled",
      "bubble",
      "stockHLC",
      "stockOHLC",
      "stockVHLC",
      "stockVOHLC",
    ])
    .describe("Type of chart to create"),
  title: z.string().describe("Chart UI display title"),
  source: z
    .string()
    .optional()
    .describe(
      "REQUIRED for create. OPTIONAL for update. Contiguous range address containing all chart data in format 'SheetName!A1:D100'. The range should have series organized in columns with headers in the first row, and category labels in the first column if applicable.",
    ),
  anchor: z
    .string()
    .optional()
    .describe(
      "A1 notation of top-left cell where chart will be anchored (e.g., 'F1'). Only used during create/update.",
    ),
  position: z
    .object({
      top: z.number().describe("Distance from top of worksheet in pt"),
      left: z.number().describe("Distance from left of worksheet in pt"),
    })
    .optional()
    .describe(
      "Chart position in pt. Only returned when fetching existing charts.",
    ),
  readOnlySeries: z
    .array(
      z.object({
        name: z.string().describe("Series name/title"),
        values: z
          .string()
          .describe(
            "A1 notation range for series data values (e.g., 'Sheet1!B2:B10')",
          ),
        categories: z
          .string()
          .optional()
          .describe(
            "A1 notation range for category labels (e.g., 'Sheet1!A2:A10')",
          ),
      }),
    )
    .optional()
    .describe(
      "Series data extracted from the chart. READ-ONLY field populated when fetching charts for inspection.",
    ),
});

const BorderStyleSchema = z.object({
  style: z
    .enum(["solid", "dashed", "dotted", "double"])
    .optional()
    .describe("Border line style"),
  weight: z
    .enum(["thin", "medium", "thick"])
    .optional()
    .describe("Border weight/thickness"),
  color: z
    .string()
    .optional()
    .describe('Border color (hex format like "#000000")'),
});

const SizeConfigSchema = z
  .object({
    type: z
      .enum(["autofit", "points", "standard"])
      .describe(
        "Size method: autofit (fit to content), points (numeric value), standard (reset to default)",
      ),
    value: z
      .number()
      .optional()
      .describe(
        "Size value in points (required for points type, ignored for autofit/standard)",
      ),
  })
  .refine((e) => !(e.type === "points" && e.value === void 0), {
    message: "value is required when type is 'points'",
    path: ["value"],
  });

const getCellRanges = tool({
  // allowedCallers: ["direct", "code_execution_20250825"],
  description:
    "READ. Get detailed information about cells in specified ranges, including values, formulas, and key formatting. Accepts multiple ranges for efficient batch reading.",
  inputSchema: z
    .object({
      sheetId: z.number().describe("The ID of the sheet to read from"),
      ranges: z
        .array(z.string())
        .describe('Array of A1 notation ranges (e.g., ["A1:C10", "E1:F5"])'),
      includeStyles: z
        .boolean()
        .optional()
        .default(!0)
        .describe(
          "Whether to include cell styles in the response (default: true)",
        ),
      cellLimit: z
        .number()
        .optional()
        .default(DEFAULT_MAX_CELLS)
        .describe(
          `Maximum number of non-empty cells to return (default: ${DEFAULT_MAX_CELLS}). Ranges will be truncated at row boundaries to fit within this limit.`,
        ),
    })
    .extend({ explanation: ExplanationSchema })
    .refine((e) => e.ranges && Array.isArray(e.ranges) && e.ranges.length > 0, {
      message: "At least one range must be provided",
      path: ["ranges"],
    })
    .refine(
      (data) =>
        data.ranges.every(
          (range) => typeof range === "string" && isValidCellRange(range),
        ),
      {
        message:
          "All ranges must be in valid A1 notation format (e.g., 'A1:C10', 'A1', 'D:D', '5:5')",
        path: ["ranges"],
      },
    ),
  outputSchema: z.object({
    worksheet: z.object({
      name: z.string().describe("Sheet name"),
      sheetId: z.number().describe("Unique sheet identifier"),
      dimension: z
        .string()
        .optional()
        .describe("Used range of the sheet (e.g., 'A1:Z100')"),
      sheetPr: z
        .object({
          tabColor: z.string().optional().describe("Tab color in hex format"),
        })
        .optional()
        .describe("Sheet properties"),
      cells: z
        .record(z.string(), z.unknown())
        .describe(
          "Map of A1 notation to cell values. Format: direct value (string/number/boolean) for simple cells, or array [value, formula, note?] for cells with formulas/notes. Examples: 'A1': 'Product', 'B2': 19.99, 'C3': [300, '=B2*2'], 'D4': [200, '=SUM(A:A)', 'Check this']. Cells are ordered column-by-column (A1, A2, A3, B1, B2...). Empty cells are omitted.",
        ),
      styles: z
        .record(z.string(), z.record(z.string(), z.unknown()))
        .optional()
        .describe(
          "Map of A1 ranges to style objects with OOXML properties. Supported properties: b (bold: true/false), i (italic: true/false), u (underline: none/single/double), strike (strikethrough: true/false), sz (font size in points), color (font color hex), family (font family), fgColor (background color hex), alignment (left/center/right), numFmt (number format string). Example: 'A1:D1': { b: true, fgColor: '#f0f0f0', alignment: 'center' }",
        ),
      borders: z
        .record(
          z.string(),
          z.union([z.string(), z.record(z.string(), z.string())]),
        )
        .optional()
        .describe(
          "Map of A1 ranges to border definitions. Format: '{width} {style} {color}' where width is thin/medium/thick, style is solid/dashed/dotted/double, color is hex. Use string for all sides or object with top/bottom/left/right keys for specific sides. Examples: 'A1:D1': { bottom: 'medium solid #000' }, 'B2:B4': 'thin solid #F00'",
        ),
    }),
    hasMore: z
      .boolean()
      .describe("True if ranges were truncated to fit within cell limit"),
    nextRanges: z
      .array(z.string())
      .optional()
      .describe(
        "Remaining ranges to fetch in a follow-up request (only present if hasMore=true)",
      ),
  }),
});

const searchData = tool({
  description:
    "READ. Search for text across the spreadsheet and return matching cell locations. Results can be used with getCellRanges for detailed analysis.",
  inputSchema: z
    .object({
      searchTerm: z
        .string()
        .describe("The text to search for in the spreadsheet"),
      sheetId: z
        .number()
        .optional()
        .describe("Optional: Limit search to a specific sheet by its ID"),
      range: z
        .string()
        .optional()
        .describe(
          'Optional: A1 notation range to limit search scope (e.g., "A1:Z100")',
        ),
      offset: z
        .number()
        .optional()
        .describe(
          "Number of results to skip for pagination (default: 0). Use nextOffset from previous response to get next page.",
        ),
      options: z
        .object({
          matchCase: z
            .boolean()
            .optional()
            .describe("Case-sensitive search (default: false)"),
          matchEntireCell: z
            .boolean()
            .optional()
            .describe("Match entire cell contents only (default: false)"),
          useRegex: z
            .boolean()
            .optional()
            .describe(
              "Treat searchTerm as regular expression (default: false)",
            ),
          matchFormulas: z
            .boolean()
            .optional()
            .describe("Search in formula text (default: false)"),
          ignoreDiacritics: z
            .boolean()
            .optional()
            .describe("Ignore accent marks (default: true)"),
          maxResults: z
            .number()
            .optional()
            .describe(
              "Maximum number of results to return per page (default: 500)",
            ),
        })
        .optional()
        .describe("Optional search configuration"),
    })
    .extend({ explanation: ExplanationSchema })
    .refine((e) => e.searchTerm && typeof e.searchTerm === "string", {
      message: "searchTerm is required and must be a string",
      path: ["searchTerm"],
    }),
  outputSchema: z.object({
    matches: z
      .array(
        z.object({
          sheetName: z.string(),
          sheetId: z.number(),
          a1: z.string(),
          value: z.any(),
          formula: z.string().nullable(),
          row: z.number(),
          column: z.number(),
        }),
      )
      .optional(),
    totalFound: z.number().optional(),
    returned: z.number().optional(),
    offset: z.number().optional(),
    hasMore: z.boolean().optional(),
    searchTerm: z.string().optional(),
    searchScope: z.string().optional(),
    nextOffset: z.number().nullable().optional(),
    message: z.string().optional(),
  }),
});

const setCellRange = tool({
  // allowedCallers: ["direct", "code_execution_20250825"],
  description:
    "WRITE. Set values, formulas, notes, and/or formatting for a range of cells. CRITICAL ARRAY DIMENSIONS: The cells array must EXACTLY match range dimensions to avoid InvalidCellRangeError. Calculate: range 'A1:D3' = 3 rows × 4 columns = [[r1c1,r1c2,r1c3,r1c4],[r2c1,r2c2,r2c3,r2c4],[r3c1,r3c2,r3c3,r3c4]]. Range 'A41:N48' = 8 rows × 14 columns = 8 arrays with 14 cells each. Single cell 'A1' = [[cell]]. Column range 'B5:B7' = [[cell1],[cell2],[cell3]]. Use {} for empty cells. Always provide a clear explanation parameter.",
  inputSchema: z
    .object({
      sheetId: z
        .number()
        .describe(
          "The ID of the sheet to modify. You can get this from get_sheets_metadata. This parameter MUST be correct to avoid errors.",
        ),
      range: z
        .string()
        .describe(
          'A1 notation range (e.g., "A1:C10"). Both start and end of the range are inclusive. The cells array must match the dimensions of this range.',
        ),
      cells: z
        .pipe(
          z.transform((data) => {
            if (typeof data === "string") {
              console.debug(
                "[tools::setCellRange] received cells as JSON string, attempting to parse",
              );
              try {
                return JSON.parse(data);
              } catch {}
            }
            return data;
          }),
          z.array(
            z.array(
              z
                .object({
                  value: z.any().optional().describe("Cell value"),
                  formula: z
                    .string()
                    .optional()
                    .describe("Cell formula if any"),
                  note: z.string().optional().describe("Cell note/comment"),
                })
                .extend({
                  cellStyles: z
                    .object({
                      fontColor: z
                        .string()
                        .optional()
                        .describe('Font color (hex format like "#000000")'),
                      fontSize: z
                        .number()
                        .optional()
                        .describe("Font size in points (e.g., 10, 12, 14)"),
                      fontFamily: z
                        .string()
                        .optional()
                        .describe(
                          'Font family (e.g., "Arial", "Roboto", "Times New Roman")',
                        ),
                      fontWeight: z
                        .enum(["normal", "bold"])
                        .optional()
                        .describe("Font weight"),
                      fontStyle: z
                        .enum(["normal", "italic"])
                        .optional()
                        .describe("Font style"),
                      fontLine: z
                        .enum(["none", "underline", "line-through"])
                        .optional()
                        .describe("Font line style"),
                      backgroundColor: z
                        .string()
                        .optional()
                        .describe(
                          'Background color (hex format like "#ffffff")',
                        ),
                      horizontalAlignment: z
                        .enum(["left", "center", "right"])
                        .optional()
                        .describe("Horizontal alignment"),
                      numberFormat: z
                        .string()
                        .optional()
                        .describe(
                          'Number format (e.g., "@" for text, "0.00" for numbers, "$#,##0.00" for currency, "mm/dd/yyyy" for dates)',
                        ),
                    })
                    .describe(
                      "Cell styling properties including fonts, colors, alignment, and number formatting",
                    )
                    .optional(),
                  borderStyles: z
                    .object({
                      top: BorderStyleSchema,
                      bottom: BorderStyleSchema,
                      left: BorderStyleSchema,
                      right: BorderStyleSchema,
                    })
                    .partial()
                    .describe("Border configuration for cells")
                    .optional(),
                }),
            ),
          ),
        )
        .describe(
          "2D array that MUST exactly match range dimensions. EXAMPLES: Range 'A1:C2' (2×3) needs [[cell1,cell2,cell3],[cell4,cell5,cell6]]. Range 'A41:N48' (8×14) needs 8 arrays with 14 cells each. Range 'B5:B7' (3×1) needs [[cell1],[cell2],[cell3]]. Single cell 'A1' needs [[cell]]. COMMON ERRORS: Range 'A1:D3' with [[val1,val2]] is WRONG - needs 3 rows × 4 columns. Range 'A1:N8' with 8×1 array is WRONG - needs 8×14. Use {} for empty cells. Each cell can contain value, formula, note, cellStyles, and borderStyles.",
        ),
      copyToRange: z
        .string()
        .optional()
        .describe(
          "Optional A1 notation range to copy the cells to. The tool will first apply the cells to the input range, then copy that pattern to this destination range. AUTOMATIC EXPANSION: Source A1:G1 can copy to A1:G100 - Excel will repeat the pattern and translate formulas (C2=B2 becomes C3=B3, C4=B4, etc). Perfect for loan schedules, financial models, data series. Example: Set A12:G12 with formulas, then copyToRange:'A13:G370' to create 358 more rows with auto-adjusted references.",
        ),
      resizeWidth: SizeConfigSchema.optional().describe(
        "Optional column width adjustment after setting values. Automatically applied to all columns in the range.",
      ),
      resizeHeight: SizeConfigSchema.optional().describe(
        "Optional row height adjustment after setting values. Automatically applied to all rows in the range.",
      ),
    })
    .extend({ explanation: ExplanationSchema })
    .refine((e) => e.sheetId !== void 0 && e.sheetId !== null, {
      message: "sheetId is required",
      path: ["sheetId"],
    })
    .refine((e) => Number.isInteger(e.sheetId) && !Number.isNaN(e.sheetId), {
      message: "sheetId must be a valid integer",
      path: ["sheetId"],
    })
    .refine((e) => e.range && typeof e.range === "string", {
      message: "range must be a string in A1 notation",
      path: ["range"],
    })
    .refine((e) => e.cells && Array.isArray(e.cells), {
      message: "cells must be an array",
      path: ["cells"],
    })
    .refine((e) => e.cells && e.cells.length > 0, {
      message: "At least one cell must be provided",
      path: ["cells"],
    }),
  outputSchema: z.object({
    formula_results: z
      .record(z.string(), z.union([z.number(), z.string()]))
      .optional()
      .describe(
        "Map of A1 notation to computed values for formula cells. Values are either numbers or error strings.",
      ),
    messages: z
      .array(z.string())
      .optional()
      .describe(
        "Informational messages about adjustments or changes made during execution",
      ),
  }),
});

const modifySheetStructure = tool({
  description:
    "WRITE. Modify sheet structure with various operations. See operation parameter for required fields.",
  inputSchema: z
    .object({
      sheetId: z.number().describe("ID of the sheet to modify"),
      operation: z
        .enum(["insert", "delete", "hide", "unhide", "freeze", "unfreeze"])
        .describe(
          "Operation to perform: insert (requires dimension, reference, count, optional position), delete/hide/unhide (requires dimension, reference, count), freeze (requires dimension, count), unfreeze (no other parameters)",
        ),
      dimension: z
        .enum(["rows", "columns"])
        .optional()
        .describe(
          "What to modify: rows or columns. Required for all operations except unfreeze (which unfreezes both)",
        ),
      reference: z
        .string()
        .optional()
        .describe(
          "Starting row number (e.g., '5') or column letter (e.g., 'C'). Required for insert/delete/hide/unhide operations. Not used for freeze/unfreeze.",
        ),
      position: z
        .enum(["before", "after"])
        .optional()
        .describe(
          "Position relative to reference (default: 'before'). Only used for insert operations",
        ),
      count: z
        .number()
        .min(1)
        .optional()
        .describe(
          "Number of rows/columns to operate on (default: 1). For delete 'A' count=3 means delete A:C. Required for freeze operations.",
        ),
    })
    .extend({ explanation: ExplanationSchema })
    .refine(
      (e) =>
        !(
          (["freeze", "unfreeze"].includes(e.operation) && e.reference) ||
          (!["freeze", "unfreeze"].includes(e.operation) && !e.reference)
        ),
      {
        message:
          "reference is required for insert/delete/hide/unhide operations but not allowed for freeze/unfreeze",
        path: ["reference"],
      },
    )
    .refine(
      (e) =>
        !(
          (e.operation === "unfreeze" && (e.count || e.dimension)) ||
          (e.operation === "freeze" && (!e.count || !e.dimension)) ||
          (!["freeze", "unfreeze"].includes(e.operation) && !e.dimension)
        ),
      {
        message:
          "dimension and count are required for freeze, not allowed for unfreeze, dimension required for other operations",
        path: ["dimension", "count"],
      },
    )
    .refine((e) => e.sheetId !== void 0 && e.sheetId !== null, {
      message: "sheetId is required",
      path: ["sheetId"],
    })
    .refine(
      (e) =>
        e.operation &&
        ["insert", "delete", "hide", "unhide", "freeze", "unfreeze"].includes(
          e.operation,
        ),
      {
        message:
          "operation must be 'insert', 'delete', 'hide', 'unhide', 'freeze', or 'unfreeze'",
        path: ["operation"],
      },
    )
    .refine(
      (e) =>
        !(
          (["freeze", "unfreeze"].includes(e.operation) && e.reference) ||
          (!["freeze", "unfreeze"].includes(e.operation) && !e.reference)
        ),
      {
        message:
          "reference is required for insert/delete/hide/unhide operations but not allowed for freeze/unfreeze",
        path: ["reference"],
      },
    )
    .refine(
      (e) =>
        !(
          (e.operation === "unfreeze" && (e.count || e.dimension)) ||
          (e.operation === "freeze" && (!e.count || !e.dimension)) ||
          (!["freeze", "unfreeze"].includes(e.operation) && !e.dimension)
        ),
      {
        message:
          "dimension and count are required for freeze, not allowed for unfreeze, dimension required for other operations",
        path: ["dimension", "count"],
      },
    ),
  outputSchema: z.object({}),
});

const modifyWorkbookStructure = tool({
  description:
    "WRITE. Create, delete, rename, or duplicate sheets in the spreadsheet workbook",
  inputSchema: z
    .object({
      operation: z
        .enum(["create", "delete", "rename", "duplicate"])
        .describe("Operation to perform"),
      sheetName: z
        .string()
        .optional()
        .describe("Name for the new sheet (required for create)"),
      rows: z
        .number()
        .optional()
        .describe("Number of rows (default: 1000, only for create)"),
      columns: z
        .number()
        .optional()
        .describe("Number of columns (default: 26, only for create)"),
      tabColor: z
        .string()
        .optional()
        .describe("Tab color in hex format (only for create)"),
      sheetId: z
        .number()
        .optional()
        .describe("ID of the sheet (required for delete, rename, duplicate)"),
      newName: z
        .string()
        .optional()
        .describe(
          "New name for the sheet (required for rename, optional for duplicate)",
        ),
    })
    .extend({ explanation: ExplanationSchema })
    .refine(
      (e) =>
        e.operation &&
        ["create", "delete", "rename", "duplicate"].includes(e.operation),
      {
        message: "operation must be one of: create, delete, rename, duplicate",
        path: ["operation"],
      },
    )
    .refine(
      (e) =>
        !(
          e.operation === "create" &&
          (!e.sheetName || typeof e.sheetName !== "string")
        ),
      {
        message:
          "sheetName is required for create operation and must be a string",
        path: ["sheetName"],
      },
    )
    .refine((e) => !(e.operation === "delete" && e.sheetId == null), {
      message: "sheetId is required for delete operation",
      path: ["sheetId"],
    })
    .refine(
      (e) => !(e.operation === "rename" && (e.sheetId == null || !e.newName)),
      {
        message: "sheetId and newName are required for rename operation",
        path: ["sheetId", "newName"],
      },
    ),
  outputSchema: z.object({
    sheetId: z.number().optional(),
    sheetName: z.string().optional(),
    rows: z.number().optional(),
    columns: z.number().optional(),
    message: z.string().optional(),
  }),
});

const copyTo = tool({
  description:
    "WRITE. Copy a range to another location with automatic formula translation and pattern expansion. Excel will repeat source patterns to fill larger destinations and adjust relative references (C2=B2 becomes C3=B3). Perfect for financial models, loan schedules, data series. Example: Copy A12:G12 to A13:G370 to create 358 rows of formulas with auto-adjusted references.",
  inputSchema: z
    .object({
      sheetId: z.number().describe("The ID of the sheet containing the ranges"),
      sourceRange: z
        .string()
        .describe(
          "A1 notation range containing the source data/formulas to copy (e.g., 'A12:G12')",
        ),
      destinationRange: z
        .string()
        .describe(
          "A1 notation range to copy to (e.g., 'A13:G370'). Can be larger than source - pattern will repeat and formulas will auto-adjust (C12=B12 becomes C13=B13, C14=B14, etc)",
        ),
    })
    .extend({ explanation: ExplanationSchema }),
  outputSchema: z.object({}),
});

const getAllObjects = tool({
  description:
    "READ. Get all spreadsheet objects (pivot tables, charts, tables) from specified sheet or all sheets with their current configuration.",
  inputSchema: z
    .object({
      sheetId: z
        .number()
        .optional()
        .describe(
          "Optional sheet ID. If not provided, gets objects from all sheets",
        ),
      id: z
        .string()
        .optional()
        .describe(
          "Optional object ID to filter by (returns single object if found)",
        ),
    })
    .extend({ explanation: ExplanationSchema }),
  outputSchema: z.object({
    objects: z
      .array(z.union([PivotTableSchema, ChartSchema]))
      .describe("Array of spreadsheet objects"),
  }),
});

const modifyObject = tool({
  description:
    "WRITE. Create, update, or delete spreadsheet objects. Create: omit id, provide objectType and properties. Update: provide id and partial properties to change (LIMITATION: cannot update source range or destination location - you must delete and recreate for those changes). Delete: provide id only. IMPORTANT: When recreating objects, always delete the existing object FIRST to avoid range conflicts, then create the new one.",
  inputSchema: z
    .object({
      operation: z
        .enum(["create", "update", "delete"])
        .describe("Operation to perform on the object"),
      sheetId: z.number().describe("Sheet ID where object is/will be located"),
      id: z
        .string()
        .optional()
        .describe("Object ID (required for update/delete, omit for create)"),
      objectType: z
        .enum(["pivotTable", "chart"])
        .describe("Type of object to modify. You must specify this."),
      properties: z
        .union([
          PivotTableSchema.omit({ id: !0, type: !0, sheetId: !0 })
            .partial()
            .loose(),
          ChartSchema.omit({
            id: !0,
            type: !0,
            sheetId: !0,
            readOnlySeries: !0,
          })
            .partial()
            .loose(),
        ])
        .optional()
        .describe(
          "Object properties to create/update. Provide fields relevant to the objectType.",
        ),
    })
    .extend({ explanation: ExplanationSchema })
    .refine((e) => e.operation === "create" || e.id !== void 0, {
      message: "id is required for update and delete operations",
      path: ["id"],
    })
    .refine(
      (e) =>
        e.operation !== "update" || Object.keys(e.properties ?? {}).length > 0,
      {
        message: "at least one property must be provided for update operations",
        path: ["object", "properties"],
      },
    ),
  outputSchema: z.object({
    id: z.string().optional().describe("ID of the object (returned on create)"),
    messages: z
      .array(z.string())
      .optional()
      .describe("Optional messages to return to the model"),
  }),
});

const resizeRange = tool({
  description:
    "WRITE. Resize columns and/or rows in a sheet. Supports autofit (fit to content), specific sizes in character/point units, or standard size reset. Can target specific ranges or entire sheet.",
  inputSchema: z
    .object({
      sheetId: z.number().describe("The ID of the sheet to modify"),
      range: z
        .string()
        .optional()
        .describe(
          "A1 notation range to resize (e.g., 'A:D' for columns, '1:5' for rows). If omitted, applies to entire sheet",
        ),
      width: SizeConfigSchema.optional().describe("Column width settings"),
      height: SizeConfigSchema.optional().describe("Row height settings"),
    })
    .extend({ explanation: ExplanationSchema })
    .refine((e) => e.width || e.height, {
      message: "At least one of width or height must be specified",
      path: ["width", "height"],
    })
    .refine((e) => e.sheetId !== void 0 && e.sheetId !== null, {
      message: "sheetId is required",
      path: ["sheetId"],
    }),
  outputSchema: z.object({}),
});

const clearCellRange = tool({
  description:
    "WRITE. Clear cells in a range. By default clears only content (values/formulas) while preserving formatting. Use clearType='all' to remove both content and formatting, or clearType='formats' to remove only formatting.",
  inputSchema: z
    .object({
      sheetId: z.number().describe("ID of the sheet to clear cells from"),
      range: z
        .string()
        .describe(
          'A1 notation range to clear (e.g., "A1:C10", "B5", "D:D", "3:3")',
        ),
      clearType: z
        .enum(["all", "contents", "formats"])
        .optional()
        .default("contents")
        .describe(
          "What to clear: 'all' (content + formatting), 'contents' (values/formulas only), 'formats' (formatting only)",
        ),
    })
    .extend({ explanation: ExplanationSchema }),
  outputSchema: z.object({}),
});

export const tools = {
  bashCodeExecution: anthropic.tools.bash_20250124({}),
  codeExecution: anthropic.tools.codeExecution_20250825({}),
  webSearch: anthropic.tools.webSearch_20250305({}),
  clearCellRange,
  copyTo,
  getAllObjects,
  getCellRanges,
  modifyObject,
  modifySheetStructure,
  modifyWorkbookStructure,
  resizeRange,
  searchData,
  setCellRange,
};

export const writeTools: Array<keyof typeof tools> = [
  "clearCellRange",
  "copyTo",
  "modifyObject",
  "modifySheetStructure",
  "modifyWorkbookStructure",
  "setCellRange",
];
