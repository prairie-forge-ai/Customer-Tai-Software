import { hasExcelRuntime } from "./gateway.js";

const CONFIG_SHEET_NAME = "SS_PF_Config";
const PREFIX_CATEGORY = "module-prefix";
const SYSTEM_MODULE = "system"; // Special module key for SS_ prefix

const MODULE_SELECTOR_KEY = "module-selector";
const MODULE_SELECTOR_HOMEPAGE_SHEET = "SS_Homepage";
const QUICK_ACCESS_SHEET_NAMES = [
    "SS_Employee_Roster",
    "SS_Chart_of_Accounts",
    "SS_PF_Config"
];

// ═══════════════════════════════════════════════════════════════════════════════
// PREFIX-BASED TAB VISIBILITY
// ═══════════════════════════════════════════════════════════════════════════════
// 
// Tab visibility is driven by prefixes defined in SS_PF_Config:
// 
// | Category      | Field (Prefix) | Value (Module)    |
// |---------------|----------------|-------------------|
// | module-prefix | PR_            | payroll-recorder  |
// | module-prefix | PTO_           | pto-accrual       |
// | module-prefix | SS_            | system            |
// 
// When entering a module:
// 1. Show tabs with that module's prefix
// 2. Hide tabs with other module prefixes  
// 3. Always hide "system" prefix tabs (SS_*)
// 4. Leave non-prefixed tabs as-is
//
// ═══════════════════════════════════════════════════════════════════════════════

// Fallback prefix config if SS_PF_Config doesn't have module-prefix rows
const DEFAULT_PREFIX_CONFIG = {
    "PR_": "payroll-recorder",
    "PTO_": "pto-accrual",
    "CC_": "credit-card-expense",
    "COM_": "commission-calc",
    "SS_": "system"
};

// Valid categories for SS_PF_Config
// Structure: Category | Field | Value | Permanent (4 columns)
export const VALID_CATEGORIES = {
    "module-prefix": {
        label: "Module Prefix",
        description: "Maps tab prefixes to modules (e.g., PR_ → payroll-recorder)"
    },
    "run-settings": {
        label: "Run Settings",
        description: "Per-period configuration values (payroll date, accounting period, etc.)"
    },
    "step-notes": {
        label: "Step Notes",
        description: "Notes and sign-off data for workflow steps"
    },
    "shared": {
        label: "Shared",
        description: "Global settings shared across all modules"
    },
    "column-mapping": {
        label: "Column Mapping",
        description: "Maps source columns to target columns for data import"
    },
    "tab-structure": {
        label: "Tab Structure",
        description: "Maps tabs to modules for visibility control"
    }
};

// Get list of valid category values
export function getValidCategoryKeys() {
    return Object.keys(VALID_CATEGORIES);
}

/**
 * Read prefix → module mappings from SS_PF_Config
 * Falls back to DEFAULT_PREFIX_CONFIG if not found
 * @returns {Promise<Object>} Map of prefix → moduleKey
 */
async function getPrefixConfig() {
    if (!hasExcelRuntime()) return { ...DEFAULT_PREFIX_CONFIG };
    
    try {
        return await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject(CONFIG_SHEET_NAME);
            await context.sync();
            
            if (configSheet.isNullObject) {
                console.log("[Tab Visibility] Config sheet not found, using defaults");
                return { ...DEFAULT_PREFIX_CONFIG };
            }
            
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (usedRange.isNullObject || !usedRange.values?.length) {
                return { ...DEFAULT_PREFIX_CONFIG };
            }
            
            const values = usedRange.values;
            const headerMap = buildHeaderMap(values[0]);
            const categoryIdx = headerMap.get("category");
            const fieldIdx = headerMap.get("field");
            const valueIdx = headerMap.get("value");
            
            if (categoryIdx === undefined || fieldIdx === undefined || valueIdx === undefined) {
                console.warn("[Tab Visibility] Missing required columns, using defaults");
                return { ...DEFAULT_PREFIX_CONFIG };
            }
            
            const prefixConfig = {};
            let foundPrefixRows = false;
            
            for (let i = 1; i < values.length; i++) {
                const row = values[i];
                const category = normalizeToken(row[categoryIdx]);
                
                if (category === PREFIX_CATEGORY) {
                    const prefix = String(row[fieldIdx] ?? "").trim().toUpperCase();
                    const moduleKey = normalizeToken(row[valueIdx]);
                    
                    if (prefix && moduleKey) {
                        prefixConfig[prefix] = moduleKey;
                        foundPrefixRows = true;
                    }
                }
            }
            
            if (!foundPrefixRows) {
                console.log("[Tab Visibility] No module-prefix rows found, using defaults");
                return { ...DEFAULT_PREFIX_CONFIG };
            }
            
            console.log("[Tab Visibility] Loaded prefix config:", prefixConfig);
            return prefixConfig;
        });
    } catch (error) {
        console.warn("[Tab Visibility] Error reading prefix config:", error);
        return { ...DEFAULT_PREFIX_CONFIG };
    }
}

/**
 * Apply tab visibility based on module
 * 
 * RULES:
 * - Module Selector: Show SS_Homepage only, hide all PR_*, PTO_*, other SS_*
 * - Payroll-Recorder: Show all PR_* tabs, hide all SS_*, PTO_*
 * - PTO-Accrual: Show all PTO_* tabs, hide all SS_*, PR_*
 * 
 * SS_* tabs can be opened manually from any module (via buttons),
 * but auto-hide when leaving that context.
 * 
 * @param {string} moduleKey - The module being activated
 */
export async function applyModuleTabVisibility(moduleKey) {
    if (!hasExcelRuntime()) return;
    
    const normalizedModuleKey = normalizeToken(moduleKey);
    console.log(`[Tab Visibility] Applying visibility for module: ${normalizedModuleKey}`);
    
    try {
        const prefixConfig = await getPrefixConfig();
        const allPrefixes = Object.keys(prefixConfig)
            .map((p) => String(p ?? "").trim().toUpperCase())
            .filter(Boolean)
            .sort((a, b) => b.length - a.length);

        const moduleKeyToPrefix = buildModuleKeyToPrefix(prefixConfig);
        const activePrefix = normalizedModuleKey === normalizeToken(MODULE_SELECTOR_KEY)
            ? null
            : (moduleKeyToPrefix[normalizedModuleKey] ?? null);

        if (normalizedModuleKey !== normalizeToken(MODULE_SELECTOR_KEY) && !activePrefix) {
            console.warn(`[Tab Visibility] No active prefix found for moduleKey="${normalizedModuleKey}". Skipping visibility changes.`);
            return;
        }

        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            const toShow = [];
            const toHide = [];
            const shownSheetNames = [];
            const hiddenSheetNames = [];
            
            worksheets.items.forEach((sheet) => {
                const sheetName = sheet.name;
                const upperName = sheetName.toUpperCase();
                
                // Determine visibility based on module rules
                let shouldShow = false;
                let shouldHide = false;
                
                const isQuickAccessSheet = QUICK_ACCESS_SHEET_NAMES.some((n) => upperName === n.toUpperCase());

                if (normalizedModuleKey === normalizeToken(MODULE_SELECTOR_KEY)) {
                    // Module Selector: Only SS_Homepage visible, hide ALL other sheets
                    if (upperName === MODULE_SELECTOR_HOMEPAGE_SHEET.toUpperCase()) {
                        shouldShow = true;
                    } else {
                        shouldHide = true;
                    }
                } else {
                    // Module state: show ONLY sheets with the active module prefix; hide sheets with other known prefixes
                    const matchedPrefix = findMatchingPrefix(upperName, allPrefixes);
                    
                    if (matchedPrefix) {
                        if (matchedPrefix === activePrefix) {
                            shouldShow = true;
                        } else {
                            shouldHide = true;
                        }
                    }
                    
                    // Always hide Quick Access sheets on module entry/switch (unless explicitly opened later)
                    if (isQuickAccessSheet) {
                        shouldShow = false;
                        shouldHide = true;
                    }
                }
                
                if (shouldShow) {
                    toShow.push(sheet);
                    console.log(`[Tab Visibility] SHOW: ${sheetName}`);
                    shownSheetNames.push(sheetName);
                } else if (shouldHide) {
                    toHide.push(sheet);
                    console.log(`[Tab Visibility] HIDE: ${sheetName}`);
                    hiddenSheetNames.push(sheetName);
                } else {
                    // Non-prefixed sheet - leave as-is
                    console.log(`[Tab Visibility] SKIP: ${sheetName} (no prefix match)`);
                }
            });
            
            // First, show all tabs that should be visible
            for (const sheet of toShow) {
                sheet.visibility = Excel.SheetVisibility.visible;
            }
            await context.sync();
            
            // Count how many sheets will remain visible after hiding
            // We need at least 1 visible sheet (Excel requirement)
            const currentlyVisible = worksheets.items.filter(
                s => s.visibility === Excel.SheetVisibility.visible
            );
            const sheetsToRemainVisible = currentlyVisible.filter(
                s => !toHide.includes(s)
            );
            
            // Only hide if at least one sheet will remain visible
            if (sheetsToRemainVisible.length >= 1) {
                for (const sheet of toHide) {
                    try {
                        sheet.visibility = Excel.SheetVisibility.hidden;
                    } catch (e) {
                        console.warn(`[Tab Visibility] Could not hide "${sheet.name}":`, e.message);
                    }
                }
                await context.sync();
            } else {
                console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");
            }
            
            console.log(`[Tab Visibility] Done! Showed ${toShow.length}, hid ${toHide.length} tabs`);
            console.log("[Tab Visibility] Summary:", {
                moduleKey: normalizedModuleKey,
                activePrefix,
                shownSheets: shownSheetNames,
                hiddenSheets: hiddenSheetNames
            });
        });
    } catch (error) {
        console.warn(`[Tab Visibility] Error applying visibility:`, error);
    }
}

/**
 * Hide system sheets (SS_* prefix)
 * Called on workbook open
 */
export async function hideSystemSheets() {
    if (!hasExcelRuntime()) return;
    
    try {
        const prefixConfig = await getPrefixConfig();
        const systemPrefixes = [];
        
        for (const [prefix, module] of Object.entries(prefixConfig)) {
            if (module === SYSTEM_MODULE) {
                systemPrefixes.push(prefix);
            }
        }
        
        if (!systemPrefixes.length) {
            systemPrefixes.push("SS_"); // Fallback
        }
        
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            const visibleSheets = worksheets.items.filter(
                sheet => sheet.visibility === Excel.SheetVisibility.visible
            );
            
            let hiddenCount = 0;
            
            worksheets.items.forEach((sheet) => {
                const upperName = sheet.name.toUpperCase();
                const isSystem = systemPrefixes.some(p => upperName.startsWith(p));
                
                if (isSystem && visibleSheets.length - hiddenCount > 1) {
                    sheet.visibility = Excel.SheetVisibility.hidden;
                    hiddenCount++;
                    console.log(`[Tab Visibility] Hidden system sheet: ${sheet.name}`);
                }
            });
            
            await context.sync();
        });
    } catch (error) {
        console.warn("[Tab Visibility] Error hiding system sheets:", error);
    }
}

/**
 * Force ALL sheets to be visible (emergency recovery / debugging)
 * Can be called from console: window.PrairieForge.showAllSheets()
 */
export async function showAllSheets() {
    if (!hasExcelRuntime()) {
        console.log("Excel not available");
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            let unhiddenCount = 0;
            worksheets.items.forEach((sheet) => {
                if (sheet.visibility !== Excel.SheetVisibility.visible) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                    console.log(`[ShowAll] Made visible: ${sheet.name}`);
                    unhiddenCount++;
                }
            });
            
            await context.sync();
            console.log(`[ShowAll] Done! Made ${unhiddenCount} sheets visible. Total: ${worksheets.items.length}`);
        });
    } catch (error) {
        console.error("[Tab Visibility] Unable to show all sheets:", error);
    }
}

/**
 * Force unhide system sheets
 * Can be called from console: window.PrairieForge.unhideSystemSheets()
 */
export async function unhideSystemSheets() {
    if (!hasExcelRuntime()) {
        console.log("Excel not available");
        return;
    }
    
    try {
        const prefixConfig = await getPrefixConfig();
        const systemPrefixes = [];
        
        for (const [prefix, module] of Object.entries(prefixConfig)) {
            if (module === SYSTEM_MODULE) {
                systemPrefixes.push(prefix);
            }
        }
        
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            worksheets.items.forEach((sheet) => {
                const upperName = sheet.name.toUpperCase();
                if (systemPrefixes.some(p => upperName.startsWith(p))) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                    console.log(`[Unhide] Made visible: ${sheet.name}`);
                }
            });
            
            await context.sync();
            console.log("[Unhide] System sheets are now visible!");
        });
    } catch (error) {
        console.error("[Tab Visibility] Unable to unhide system sheets:", error);
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════════

function buildHeaderMap(headers = []) {
    const map = new Map();
    headers.forEach((header, index) => {
        const normalized = normalizeToken(header);
        if (normalized) {
            map.set(normalized, index);
        }
    });
    return map;
}

function normalizeToken(value) {
    return String(value ?? "")
        .trim()
        .toLowerCase()
        .replace(/[\s_]+/g, "-");
}

function buildModuleKeyToPrefix(prefixConfig) {
    const result = {};
    for (const [prefixRaw, moduleRaw] of Object.entries(prefixConfig ?? {})) {
        const prefix = String(prefixRaw ?? "").trim().toUpperCase();
        const moduleKey = normalizeToken(moduleRaw);
        if (!prefix || !moduleKey) continue;
        if (!result[moduleKey]) {
            result[moduleKey] = prefix;
        }
    }
    return result;
}

function findMatchingPrefix(upperSheetName, allPrefixes) {
    for (const prefix of allPrefixes) {
        if (upperSheetName.startsWith(prefix)) return prefix;
    }
    return null;
}

// ═══════════════════════════════════════════════════════════════════════════════
// VALIDATION (for config sheet validation tools)
// ═══════════════════════════════════════════════════════════════════════════════

export function validateConfigRow(row, rowIndex) {
    const errors = [];
    const warnings = [];
    
    const category = normalizeToken(row.category || row[0] || "");
    const field = String(row.field || row[1] || "").trim();
    const value = String(row.value || row[2] || "").trim();
    
    if (!category) {
        errors.push(`Row ${rowIndex}: Missing Category`);
    } else if (!VALID_CATEGORIES[category]) {
        errors.push(`Row ${rowIndex}: Invalid Category "${row.category || row[0]}". Valid: ${Object.keys(VALID_CATEGORIES).join(", ")}`);
    }
    
    if (!field) {
        errors.push(`Row ${rowIndex}: Missing Field name`);
    }
    
    if (!value && category !== "step-notes") {
        warnings.push(`Row ${rowIndex}: Value is empty for "${field}"`);
    }
    
    // Prefix validation
    if (category === PREFIX_CATEGORY) {
        if (!field.endsWith("_")) {
            warnings.push(`Row ${rowIndex}: Prefix "${field}" should end with underscore (e.g., "PR_")`);
        }
    }
    
    return { errors, warnings, isValid: errors.length === 0 };
}

export async function validateConfigSheet() {
    if (!hasExcelRuntime()) return { errors: [], warnings: [], isValid: true };
    
    const results = { errors: [], warnings: [], isValid: true, rowCount: 0 };
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject(CONFIG_SHEET_NAME);
            await context.sync();
            
            if (configSheet.isNullObject) {
                results.errors.push("SS_PF_Config sheet not found");
                results.isValid = false;
                return;
            }
            
            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (usedRange.isNullObject || !usedRange.values || usedRange.values.length < 2) {
                results.warnings.push("SS_PF_Config appears empty or has no data rows");
                return;
            }
            
            const rows = usedRange.values;
            results.rowCount = rows.length - 1;
            
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const validation = validateConfigRow({
                    category: row[0],
                    field: row[1],
                    value: row[2]
                }, i + 1);
                
                results.errors.push(...validation.errors);
                results.warnings.push(...validation.warnings);
                if (!validation.isValid) results.isValid = false;
            }
        });
    } catch (error) {
        results.errors.push(`Validation error: ${error.message}`);
        results.isValid = false;
    }
    
    return results;
}

/**
 * Temporarily show and activate a specific sheet
 * Used by Quick Access to open system sheets like SS_Employee_Roster
 * @param {string} sheetName - The exact name of the sheet to show
 * @returns {Promise<boolean>} True if successful
 */
export async function showAndActivateSheet(sheetName) {
    if (!hasExcelRuntime()) return false;
    
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            await context.sync();
            
            if (sheet.isNullObject) {
                console.warn(`[Tab Visibility] Sheet "${sheetName}" not found`);
                return false;
            }
            
            // Make visible and activate
            sheet.visibility = Excel.SheetVisibility.visible;
            sheet.activate();
            await context.sync();
            
            console.log(`[Tab Visibility] Showed and activated: ${sheetName}`);
            return true;
        });
        return true;
    } catch (error) {
        console.error(`[Tab Visibility] Error showing sheet "${sheetName}":`, error);
        return false;
    }
}

// ═══════════════════════════════════════════════════════════════════════════════
// GLOBAL EXPORTS (for console access)
// ═══════════════════════════════════════════════════════════════════════════════

if (typeof window !== "undefined") {
    window.PrairieForge = window.PrairieForge || {};
    window.PrairieForge.showAllSheets = showAllSheets;
    window.PrairieForge.unhideSystemSheets = unhideSystemSheets;
    window.PrairieForge.applyModuleTabVisibility = applyModuleTabVisibility;
    window.PrairieForge.showAndActivateSheet = showAndActivateSheet;
}
