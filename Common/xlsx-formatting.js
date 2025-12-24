/**
 * XLSX Formatting Utilities
 * 
 * Utilities for formatting SheetJS (xlsx) worksheets for archive exports.
 * Provides consistent formatting for headers, columns, and data types.
 * 
 * Usage:
 *   import { formatXlsxWorksheet, applyXlsxHeaderStyle, XLSX_FORMATS } from "../../Common/xlsx-formatting.js";
 */

import * as XLSX from "xlsx";

// Standard number format codes for SheetJS
export const XLSX_FORMATS = {
    currency: "$#,##0.00",
    currencyNegative: "$#,##0.00;[Red]($#,##0.00)",
    number: "#,##0.00",
    integer: "#,##0",
    percent: "0.00%",
    date: "yyyy-mm-dd",
    dateShort: "mm/dd/yyyy",
    text: "@"
};

// Standard column widths (in characters)
export const XLSX_COLUMN_WIDTHS = {
    narrow: 10,
    standard: 15,
    wide: 20,
    extraWide: 30,
    description: 40
};

/**
 * Apply header formatting to first row of a SheetJS worksheet
 * Note: SheetJS free version has limited styling support
 * This is a placeholder for future enhancement with SheetJS Pro
 * 
 * @param {Object} worksheet - SheetJS worksheet object
 * @param {number} columnCount - Number of columns in the header
 */
export function applyXlsxHeaderStyle(worksheet, columnCount) {
    // SheetJS free version doesn't support cell styling (fonts, fills, etc.)
    // Column widths and number formats are supported and applied separately
    // For full styling support, SheetJS Pro would be required
    return;
}

/**
 * Set column widths for a worksheet
 * 
 * @param {Object} worksheet - SheetJS worksheet object
 * @param {Array<number>} widths - Array of column widths in characters
 */
export function setXlsxColumnWidths(worksheet, widths) {
    if (!worksheet || !widths || widths.length === 0) return;
    
    worksheet['!cols'] = widths.map(w => ({ wch: w }));
}

/**
 * Apply number format to specific columns
 * 
 * @param {Object} worksheet - SheetJS worksheet object
 * @param {number} rowCount - Total number of rows (including header)
 * @param {Object} columnFormats - Map of column index to format string
 *   Example: { 3: XLSX_FORMATS.currency, 5: XLSX_FORMATS.date }
 */
export function applyXlsxColumnFormats(worksheet, rowCount, columnFormats) {
    if (!worksheet || !columnFormats || rowCount <= 1) return;
    
    // Apply formats to data rows (skip header row 0)
    for (let row = 1; row < rowCount; row++) {
        for (const [colIndex, format] of Object.entries(columnFormats)) {
            const col = parseInt(colIndex, 10);
            const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
            
            if (!worksheet[cellRef]) continue;
            
            // Set number format directly on the cell
            // Note: This sets the format code but SheetJS free version has limited support
            worksheet[cellRef].z = format;
        }
    }
}

/**
 * Auto-detect and apply appropriate formats based on header names
 * 
 * @param {Object} worksheet - SheetJS worksheet object
 * @param {Array<string>} headers - Array of header names
 * @param {number} rowCount - Total number of rows (including header)
 * @returns {Object} Map of column indices to applied formats
 */
export function autoFormatXlsxColumns(worksheet, headers, rowCount) {
    if (!worksheet || !headers || rowCount <= 1) return {};
    
    const appliedFormats = {};
    
    headers.forEach((header, colIndex) => {
        const headerLower = String(header || "").toLowerCase().trim();
        let format = null;
        
        // Currency columns
        if (headerLower.includes("amount") || 
            headerLower.includes("total") || 
            headerLower.includes("liability") ||
            headerLower.includes("pay") ||
            headerLower.includes("wage") ||
            headerLower.includes("salary") ||
            headerLower.includes("rate") && !headerLower.includes("accrual rate") ||
            headerLower.includes("debit") ||
            headerLower.includes("credit") ||
            headerLower.includes("change") ||
            headerLower.includes("fixed") ||
            headerLower.includes("variable") ||
            headerLower.includes("burden") ||
            headerLower.includes("gross")) {
            format = XLSX_FORMATS.currency;
        }
        // Date columns
        else if (headerLower.includes("date") || 
                 headerLower.includes("period")) {
            format = XLSX_FORMATS.dateShort;
        }
        // Percent columns
        else if (headerLower.includes("percent") || 
                 headerLower.includes("rate") && headerLower.includes("burden") ||
                 headerLower === "% of total") {
            format = XLSX_FORMATS.percent;
        }
        // Integer columns
        else if (headerLower.includes("count") || 
                 headerLower.includes("headcount") ||
                 headerLower === "id" ||
                 headerLower.includes("employee_id")) {
            format = XLSX_FORMATS.integer;
        }
        
        if (format) {
            appliedFormats[colIndex] = format;
        }
    });
    
    // Apply the detected formats
    applyXlsxColumnFormats(worksheet, rowCount, appliedFormats);
    
    return appliedFormats;
}

/**
 * Auto-detect and set appropriate column widths based on header names
 * 
 * @param {Array<string>} headers - Array of header names
 * @returns {Array<number>} Array of column widths
 */
export function autoSizeXlsxColumns(headers) {
    if (!headers || headers.length === 0) return [];
    
    return headers.map(header => {
        const headerLower = String(header || "").toLowerCase().trim();
        const headerLength = header.length;
        
        // Extra wide for descriptions and notes
        if (headerLower.includes("description") || 
            headerLower.includes("note") ||
            headerLower.includes("name") && headerLower.includes("account")) {
            return XLSX_COLUMN_WIDTHS.description;
        }
        // Wide for names
        else if (headerLower.includes("name") || 
                 headerLower.includes("department")) {
            return XLSX_COLUMN_WIDTHS.extraWide;
        }
        // Standard for most columns
        else if (headerLength > 12) {
            return XLSX_COLUMN_WIDTHS.wide;
        }
        // Narrow for short columns
        else if (headerLength < 8) {
            return XLSX_COLUMN_WIDTHS.narrow;
        }
        
        return XLSX_COLUMN_WIDTHS.standard;
    });
}

/**
 * Complete formatting for a SheetJS worksheet
 * Applies headers, column widths, and number formats
 * 
 * @param {Object} worksheet - SheetJS worksheet object
 * @param {Array<string>} headers - Array of header names (first row)
 * @param {number} rowCount - Total number of rows (including header)
 * @param {Object} [options] - Optional formatting options
 * @param {Object} [options.columnFormats] - Manual column format overrides
 * @param {Array<number>} [options.columnWidths] - Manual column width overrides
 * @param {boolean} [options.autoFormat=true] - Auto-detect formats from headers
 * @param {boolean} [options.autoSize=true] - Auto-size columns from headers
 */
export function formatXlsxWorksheet(worksheet, headers, rowCount, options = {}) {
    if (!worksheet || !headers || headers.length === 0) return;
    
    const {
        columnFormats = {},
        columnWidths = null,
        autoFormat = true,
        autoSize = true
    } = options;
    
    // Apply header styling
    applyXlsxHeaderStyle(worksheet, headers.length);
    
    // Auto-detect and apply column formats
    if (autoFormat) {
        autoFormatXlsxColumns(worksheet, headers, rowCount);
    }
    
    // Apply manual format overrides
    if (columnFormats && Object.keys(columnFormats).length > 0) {
        applyXlsxColumnFormats(worksheet, rowCount, columnFormats);
    }
    
    // Set column widths
    if (columnWidths) {
        setXlsxColumnWidths(worksheet, columnWidths);
    } else if (autoSize) {
        const widths = autoSizeXlsxColumns(headers);
        setXlsxColumnWidths(worksheet, widths);
    }
}
