import type { InferToolInput, InferToolOutput } from "ai";
import * as z from "zod";
import type { tools } from "@/server/ai/tools";

// ============================================================================
// Core Types
// ============================================================================

export const Sheet = z.object({
  id: z.int(),
  name: z.string(),
  maxRows: z.int(),
  maxColumns: z.int(),
});
export type Sheet = z.infer<typeof Sheet>;

// ============================================================================
// Spreadsheet Service Interface
// ============================================================================

/**
 * Unified interface for spreadsheet operations.
 * Implemented by Excel (Office.js) and Google Sheets (Apps Script) services.
 */
export type SpreadsheetService = {
  // -------------------------------------------------------------------------
  // Read Operations
  // -------------------------------------------------------------------------

  /**
   * Get metadata for all sheets in the workbook.
   */
  getSheets(): Promise<Sheet[]>;

  /**
   * Get cell data from specified ranges including values, formulas, and styles.
   */
  getCellRanges(
    input: InferToolInput<typeof tools.getCellRanges>,
  ): Promise<InferToolOutput<typeof tools.getCellRanges>>;

  /**
   * Search for text across sheets and return matching cell locations.
   */
  searchData(
    input: InferToolInput<typeof tools.searchData>,
  ): Promise<InferToolOutput<typeof tools.searchData>>;

  /**
   * Get all spreadsheet objects (pivot tables, charts) from sheets.
   */
  getAllObjects(
    input: InferToolInput<typeof tools.getAllObjects>,
  ): Promise<InferToolOutput<typeof tools.getAllObjects>>;

  // -------------------------------------------------------------------------
  // Write Operations
  // -------------------------------------------------------------------------

  /**
   * Set values, formulas, notes, and formatting for a range of cells.
   */
  setCellRange(
    input: InferToolInput<typeof tools.setCellRange>,
  ): Promise<InferToolOutput<typeof tools.setCellRange>>;

  /**
   * Copy a range to another location with formula translation.
   */
  copyTo(
    input: InferToolInput<typeof tools.copyTo>,
  ): Promise<InferToolOutput<typeof tools.copyTo>>;

  /**
   * Clear cells in a range (contents, formats, or both).
   */
  clearCellRange(
    input: InferToolInput<typeof tools.clearCellRange>,
  ): Promise<InferToolOutput<typeof tools.clearCellRange>>;

  /**
   * Resize column widths and/or row heights.
   */
  resizeRange(
    input: InferToolInput<typeof tools.resizeRange>,
  ): Promise<InferToolOutput<typeof tools.resizeRange>>;

  // -------------------------------------------------------------------------
  // Structure Operations
  // -------------------------------------------------------------------------

  /**
   * Modify sheet structure: insert/delete/hide/freeze rows and columns.
   */
  modifySheetStructure(
    input: InferToolInput<typeof tools.modifySheetStructure>,
  ): Promise<InferToolOutput<typeof tools.modifySheetStructure>>;

  /**
   * Modify workbook structure: create/delete/rename/duplicate sheets.
   */
  modifyWorkbookStructure(
    input: InferToolInput<typeof tools.modifyWorkbookStructure>,
  ): Promise<InferToolOutput<typeof tools.modifyWorkbookStructure>>;

  // -------------------------------------------------------------------------
  // Object Operations
  // -------------------------------------------------------------------------

  /**
   * Create, update, or delete spreadsheet objects (charts, pivot tables).
   */
  modifyObject(
    input: InferToolInput<typeof tools.modifyObject>,
  ): Promise<InferToolOutput<typeof tools.modifyObject>>;

  // -------------------------------------------------------------------------
  // UI Operations
  // -------------------------------------------------------------------------

  /**
   * Activate (switch to) a specific sheet.
   */
  activateSheet(sheetId: number): Promise<void>;

  /**
   * Clear the current selection.
   */
  clearSelection(): Promise<void>;

  /**
   * Select a range on a sheet (for preview before modifications).
   */
  selectRange(input: { sheetId: number; range: string }): Promise<void>;
};
