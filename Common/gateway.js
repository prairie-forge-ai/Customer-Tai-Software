/**
 * Gateway - Shared Office/Excel initialization and helpers for all modules
 */

// Excel runtime check
export function hasExcelRuntime() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

// Initialize Office and call the provided callback when ready
export function initializeOffice(onReady) {
    try {
        Office.onReady((info) => {
            console.log("Office.onReady fired:", info);
            if (info.host === Office.HostType.Excel) {
                onReady(info);
            } else {
                console.warn("Not running in Excel, host:", info.host);
                onReady(info); // Still init for testing
            }
        });
    } catch (error) {
        console.warn("Office.onReady failed:", error);
        onReady(null);
    }
}

// Shared config table name - single source of truth for all modules
export const SHARED_CONFIG_TABLE = "SS_PF_Config";

// Find a config table by name candidates (prioritizes SS_PF_Config)
export async function getConfigTable(context, tableCandidates = [SHARED_CONFIG_TABLE]) {
    const tables = context.workbook.tables;
    tables.load("items/name");
    await context.sync();

    const match = tables.items?.find((t) => tableCandidates.includes(t.name));
    if (!match) {
        console.warn("Config table not found. Looking for:", tableCandidates);
        return null;
    }

    return context.workbook.tables.getItem(match.name);
}

// Get column indices from table headers
export function getColumnIndices(headers) {
    const normalizedHeaders = headers.map(h => String(h || "").trim().toLowerCase());
    return {
        field: normalizedHeaders.findIndex(h => h === "field" || h === "field name" || h === "setting"),
        value: normalizedHeaders.findIndex(h => h === "value" || h === "setting value"),
        type: normalizedHeaders.findIndex(h => h === "type" || h === "category"),
        title: normalizedHeaders.findIndex(h => h === "title" || h === "display name"),
        permanent: normalizedHeaders.findIndex(h => h === "permanent" || h === "persist")
    };
}

// Load configuration from a config table
export async function loadConfigFromTable(tableCandidates = [SHARED_CONFIG_TABLE]) {
    if (!hasExcelRuntime()) {
        return {};
    }

    try {
        return await Excel.run(async (context) => {
            const table = await getConfigTable(context, tableCandidates);
            if (!table) {
                return {};
            }

            const body = table.getDataBodyRange();
            const headerRange = table.getHeaderRowRange();
            body.load("values");
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0] || [];
            const cols = getColumnIndices(headers);

            if (cols.field === -1 || cols.value === -1) {
                console.warn("Config table missing FIELD or VALUE columns. Headers:", headers);
                return {};
            }

            const values = {};
            const rows = body.values || [];
            rows.forEach((row) => {
                const field = String(row[cols.field] || "").trim();
                if (field) {
                    values[field] = row[cols.value] ?? "";
                }
            });

            console.log("Configuration loaded:", Object.keys(values).length, "fields");
            return values;
        });
    } catch (error) {
        console.error("Failed to load configuration:", error);
        return {};
    }
}

// Save a config value to a config table
// Static fields that should be marked as Permanent (never cleared on archive)
const PERMANENT_FIELDS = new Set([
    "SS_Installation_Key",
    "SS_Company_ID",
    "SS_Company_Name",
    "SS_Accounting_Software",
    "PTO_Payroll_Provider",
    "PR_Payroll_Provider"
]);

export async function saveConfigValue(fieldName, value, tableCandidates = [SHARED_CONFIG_TABLE]) {
    if (!hasExcelRuntime()) return false;

    try {
        await Excel.run(async (context) => {
            const table = await getConfigTable(context, tableCandidates);
            if (!table) {
                console.warn("Config table not found for write");
                return;
            }

            const body = table.getDataBodyRange();
            const headerRange = table.getHeaderRowRange();
            body.load("values");
            headerRange.load("values");
            await context.sync();

            const headers = headerRange.values[0] || [];
            const cols = getColumnIndices(headers);

            if (cols.field === -1 || cols.value === -1) {
                console.error("Config table missing FIELD or VALUE columns");
                return;
            }

            const rows = body.values || [];
            const targetIndex = rows.findIndex(
                (row) => String(row[cols.field] || "").trim() === fieldName
            );

            // Determine if this field should be permanent
            const shouldBePermanent = PERMANENT_FIELDS.has(fieldName);

            if (targetIndex >= 0) {
                body.getCell(targetIndex, cols.value).values = [[value]];
                // Also ensure permanent flag is correct for permanent fields
                if (shouldBePermanent && cols.permanent >= 0) {
                    body.getCell(targetIndex, cols.permanent).values = [["Y"]];
                }
            } else {
                // Add new row with Category, Field, Value, Permanent structure
                const newRow = new Array(headers.length).fill("");
                if (cols.type >= 0) newRow[cols.type] = shouldBePermanent ? "Shared" : "Run Settings";
                newRow[cols.field] = fieldName;
                newRow[cols.value] = value;
                if (cols.permanent >= 0) newRow[cols.permanent] = shouldBePermanent ? "Y" : "N";
                if (cols.title >= 0) newRow[cols.title] = "";
                table.rows.add(null, [newRow]);
                console.log("Added new config row:", fieldName, "=", value, shouldBePermanent ? "(permanent)" : "");
            }

            await context.sync();
            console.log("Saved config:", fieldName, "=", value);
        });
        return true;
    } catch (error) {
        console.error("Failed to save config:", fieldName, error);
        return false;
    }
}

// Activate a worksheet by name
export async function activateWorksheet(sheetName) {
    if (!hasExcelRuntime()) return;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            await context.sync();
            if (!sheet.isNullObject) {
                sheet.activate();
                await context.sync();
            }
        });
    } catch (error) {
        console.error("Failed to activate worksheet:", sheetName, error);
    }
}

// Get used range values from a worksheet
export async function getSheetData(sheetName) {
    if (!hasExcelRuntime()) return [];

    try {
        return await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();

            if (range.isNullObject) {
                return [];
            }
            return range.values || [];
        });
    } catch (error) {
        console.error("Failed to get sheet data:", sheetName, error);
        return [];
    }
}

// Write data to a worksheet (clears existing data first)
export async function writeSheetData(sheetName, data) {
    if (!hasExcelRuntime() || !data.length) return false;

    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            const existingRange = sheet.getUsedRangeOrNullObject();
            await context.sync();

            if (!existingRange.isNullObject) {
                existingRange.clear();
            }

            const targetRange = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
            targetRange.values = data;
            await context.sync();
        });
        return true;
    } catch (error) {
        console.error("Failed to write sheet data:", sheetName, error);
        return false;
    }
}

// =============================================================================
// COLUMN ALIAS SYSTEM
// =============================================================================
// Allows users to define custom column names that map to system-expected columns.
// Stored in SS_PF_Config with Category="column-alias"
// Field pattern: {MODULE}_{systemColumn}_alias (e.g., PR_amount_alias)
// Value: pipe-separated list of aliases (e.g., "gross pay|total pay|earnings")

/**
 * Default column aliases - used when no custom aliases are configured.
 * These are the built-in fuzzy matching patterns that currently exist in the code.
 */
export const DEFAULT_COLUMN_ALIASES = {
    // Payroll Recorder columns
    PR: {
        amount: ["amount", "gross pay", "grosspay", "total pay", "earnings", "gross", "total"],
        employee: ["employee", "employee name", "employee-name", "name", "worker", "emp"],
        department: ["department", "dept", "division", "cost center", "cc", "costcenter"],
        date: ["date", "pay date", "payroll date", "pay period", "period", "paydate"]
    },
    // PTO Accrual columns
    PTO: {
        employee: ["employee", "employee name", "name", "worker"],
        balance: ["balance", "pto balance", "current balance", "accrual balance"],
        accrued: ["accrued", "ytd accrued", "pay period accrued", "earned"],
        used: ["used", "ytd used", "pay period used", "taken"],
        plan: ["plan", "plan description", "pto plan", "plan type"]
    }
};

// Cache for loaded column aliases
let columnAliasCache = null;

/**
 * Normalize a column header for comparison.
 * Converts to lowercase, trims whitespace, removes special characters.
 * @param {string} value - The column header to normalize
 * @returns {string} Normalized string
 */
export function normalizeColumnHeader(value) {
    if (!value) return "";
    return String(value)
        .toLowerCase()
        .trim()
        .replace(/[^a-z0-9\s]/g, "")
        .replace(/\s+/g, " ");
}

/**
 * Load column aliases from SS_PF_Config.
 * Merges custom aliases with defaults.
 * @param {boolean} forceReload - Force reload from Excel even if cached
 * @returns {Promise<Object>} Column aliases by module and column
 */
export async function loadColumnAliases(forceReload = false) {
    if (columnAliasCache && !forceReload) {
        return columnAliasCache;
    }

    // Start with defaults
    const aliases = JSON.parse(JSON.stringify(DEFAULT_COLUMN_ALIASES));

    if (!hasExcelRuntime()) {
        columnAliasCache = aliases;
        return aliases;
    }

    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();

            if (configSheet.isNullObject) {
                console.log("[ColumnAlias] SS_PF_Config not found, using defaults");
                return;
            }

            const range = configSheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();

            if (range.isNullObject) return;

            const values = range.values || [];
            if (values.length < 2) return;

            // Find columns
            const headers = values[0].map(h => String(h || "").toLowerCase().trim());
            const catIdx = headers.findIndex(h => h === "category");
            const fieldIdx = headers.findIndex(h => h === "field");
            const valueIdx = headers.findIndex(h => h === "value");

            if (catIdx === -1 || fieldIdx === -1 || valueIdx === -1) return;

            // Parse column-alias rows
            for (let i = 1; i < values.length; i++) {
                const category = String(values[i][catIdx] || "").toLowerCase().trim();
                const field = String(values[i][fieldIdx] || "").trim();
                const value = String(values[i][valueIdx] || "").trim();

                if (category !== "column-alias" || !field || !value) continue;

                // Parse field: {MODULE}_{column}_alias
                const match = field.match(/^([A-Z]+)_(.+)_alias$/i);
                if (!match) continue;

                const module = match[1].toUpperCase();
                const column = match[2].toLowerCase();
                const customAliases = value.split("|").map(a => normalizeColumnHeader(a)).filter(a => a);

                if (!aliases[module]) {
                    aliases[module] = {};
                }

                // Merge custom aliases with defaults (custom takes priority)
                const existingAliases = aliases[module][column] || [];
                const mergedAliases = [...new Set([...customAliases, ...existingAliases])];
                aliases[module][column] = mergedAliases;

                console.log(`[ColumnAlias] Loaded ${module}_${column}: ${mergedAliases.join(", ")}`);
            }
        });
    } catch (error) {
        console.error("[ColumnAlias] Failed to load aliases:", error);
    }

    columnAliasCache = aliases;
    return aliases;
}

/**
 * Find a column index in headers using aliases.
 * Checks the header against all known aliases for the specified column.
 * @param {Array<string>} headers - Array of column headers
 * @param {string} module - Module key (e.g., "PR", "PTO")
 * @param {string} systemColumn - The system column name (e.g., "amount", "employee")
 * @param {Object} [aliasConfig] - Optional pre-loaded alias config
 * @returns {number} Column index, or -1 if not found
 */
export function findColumnByAlias(headers, module, systemColumn, aliasConfig = null) {
    const config = aliasConfig || columnAliasCache || DEFAULT_COLUMN_ALIASES;
    const moduleConfig = config[module.toUpperCase()];
    
    if (!moduleConfig || !moduleConfig[systemColumn]) {
        console.warn(`[ColumnAlias] No aliases defined for ${module}.${systemColumn}`);
        return -1;
    }

    const aliases = moduleConfig[systemColumn];
    const normalizedHeaders = headers.map(h => normalizeColumnHeader(h));

    // First try exact match
    for (const alias of aliases) {
        const idx = normalizedHeaders.findIndex(h => h === alias);
        if (idx >= 0) {
            console.log(`[ColumnAlias] Found ${module}.${systemColumn} at column ${idx} (exact: "${alias}")`);
            return idx;
        }
    }

    // Then try "includes" match for partial matching
    for (const alias of aliases) {
        const idx = normalizedHeaders.findIndex(h => h.includes(alias) || alias.includes(h));
        if (idx >= 0) {
            console.log(`[ColumnAlias] Found ${module}.${systemColumn} at column ${idx} (partial: "${alias}")`);
            return idx;
        }
    }

    console.warn(`[ColumnAlias] Column not found: ${module}.${systemColumn}. Headers: ${headers.join(", ")}`);
    return -1;
}

/**
 * Build a column index map for a module using aliases.
 * @param {Array<string>} headers - Array of column headers from the data
 * @param {string} module - Module key (e.g., "PR", "PTO")
 * @param {Array<string>} requiredColumns - List of system columns to find
 * @param {Object} [aliasConfig] - Optional pre-loaded alias config
 * @returns {Object} Map of systemColumn -> columnIndex
 */
export function buildColumnIndexMap(headers, module, requiredColumns, aliasConfig = null) {
    const indexMap = {};
    const missing = [];

    for (const col of requiredColumns) {
        const idx = findColumnByAlias(headers, module, col, aliasConfig);
        if (idx >= 0) {
            indexMap[col] = idx;
        } else {
            missing.push(col);
        }
    }

    if (missing.length > 0) {
        console.warn(`[ColumnAlias] Missing columns for ${module}:`, missing);
    }

    return { indexMap, missing };
}

/**
 * Save a column alias to SS_PF_Config.
 * @param {string} module - Module key (e.g., "PR", "PTO")
 * @param {string} systemColumn - The system column name (e.g., "amount")
 * @param {Array<string>} aliases - Array of alias strings
 * @returns {Promise<boolean>} Success flag
 */
export async function saveColumnAlias(module, systemColumn, aliases) {
    if (!hasExcelRuntime()) return false;

    const fieldName = `${module.toUpperCase()}_${systemColumn}_alias`;
    const value = aliases.map(a => a.trim()).filter(a => a).join("|");

    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();

            if (configSheet.isNullObject) {
                console.error("[ColumnAlias] SS_PF_Config sheet not found");
                return;
            }

            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values, rowCount");
            await context.sync();

            if (usedRange.isNullObject) return;

            const values = usedRange.values || [];
            const headers = values[0].map(h => String(h || "").toLowerCase().trim());
            const catIdx = headers.findIndex(h => h === "category");
            const fieldIdx = headers.findIndex(h => h === "field");
            const valueIdx = headers.findIndex(h => h === "value");
            const permIdx = headers.findIndex(h => h === "permanent");

            if (fieldIdx === -1 || valueIdx === -1) {
                console.error("[ColumnAlias] SS_PF_Config missing required columns");
                return;
            }

            // Find existing row
            let existingRow = -1;
            for (let i = 1; i < values.length; i++) {
                if (String(values[i][fieldIdx] || "").trim() === fieldName) {
                    existingRow = i;
                    break;
                }
            }

            if (existingRow >= 0) {
                // Update existing row
                configSheet.getRange(`${String.fromCharCode(65 + valueIdx)}${existingRow + 1}`).values = [[value]];
            } else {
                // Add new row
                const newRowNum = values.length + 1;
                if (catIdx >= 0) {
                    configSheet.getRange(`${String.fromCharCode(65 + catIdx)}${newRowNum}`).values = [["column-alias"]];
                }
                configSheet.getRange(`${String.fromCharCode(65 + fieldIdx)}${newRowNum}`).values = [[fieldName]];
                configSheet.getRange(`${String.fromCharCode(65 + valueIdx)}${newRowNum}`).values = [[value]];
                if (permIdx >= 0) {
                    configSheet.getRange(`${String.fromCharCode(65 + permIdx)}${newRowNum}`).values = [["Y"]];
                }
            }

            await context.sync();
            console.log(`[ColumnAlias] Saved ${fieldName}: ${value}`);

            // Invalidate cache
            columnAliasCache = null;
        });
        return true;
    } catch (error) {
        console.error("[ColumnAlias] Failed to save alias:", error);
        return false;
    }
}

/**
 * Get all configured aliases for a module.
 * Useful for displaying in admin UI.
 * @param {string} module - Module key (e.g., "PR", "PTO")
 * @returns {Promise<Object>} Map of systemColumn -> aliases array
 */
export async function getModuleColumnAliases(module) {
    const aliases = await loadColumnAliases();
    return aliases[module.toUpperCase()] || {};
}

/**
 * Clear the column alias cache.
 * Call this when you know aliases have changed externally.
 */
export function clearColumnAliasCache() {
    columnAliasCache = null;
}
