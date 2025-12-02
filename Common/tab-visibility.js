import { hasExcelRuntime } from "./gateway.js";

const CONFIG_SHEET_NAME = "SS_PF_Config";
const TAB_STRUCTURE_CATEGORY = "tab-structure";
const ALWAYS_SHOW_TOKENS = new Set(["all", "shared", "global", "common", "any", "*"]);

// System sheets that should be hidden by default when workbook opens
const SYSTEM_SHEETS = [
    "SS_PF_Config",
    "SS_Employee_Roster", 
    "SS_Chart_of_Accounts"
];

// ═══════════════════════════════════════════════════════════════════════════════
// MODULE-SPECIFIC TAB VISIBILITY CONFIGURATION
// ═══════════════════════════════════════════════════════════════════════════════
// Defines which tabs should be VISIBLE vs HIDDEN for each module.
// Hidden tabs are still available (user can unhide during session).
// On module exit, hidden tabs revert to hidden state.

const MODULE_TAB_CONFIG = {
    "payroll-recorder": {
        // Tabs that should be VISIBLE when opening this module
        visible: [
            "PR_Data",
            "PR_Data_Clean",
            "PR_Expense_Review",
            "PR_JE_Draft"
        ],
        // Tabs that should be HIDDEN but available (user can manually unhide)
        hidden: [
            "SS_PF_Config",
            "SS_Employee_Roster",
            "SS_Chart_of_Accounts",
            "PR_Expense_Mapping",
            "PR_Archive_Summary",
            "PR_Homepage",
            "SS_Homepage"
        ]
    },
    "pto-accrual": {
        // Tabs that should be VISIBLE when opening this module
        visible: [
            "PTO_Data",
            "PTO_Analysis",
            "PTO_JE_Draft"
        ],
        // Tabs that should be HIDDEN but available (user can manually unhide)
        hidden: [
            "SS_PF_Config",
            "SS_Employee_Roster",
            "SS_Chart_of_Accounts",
            "PTO_Archive_Summary",
            "PTO_Homepage",
            "SS_Homepage"
        ]
    }
};

// Valid categories for SS_PF_Config
export const VALID_CATEGORIES = {
    "tab-structure": {
        label: "Tab Structure",
        description: "Defines which tabs belong to which module",
        requiresValue2: true,  // Module key required
        value2Options: ["payroll-recorder", "pto-accrual", "employee-roster", "shared", "all"]
    },
    "run-settings": {
        label: "Run Settings",
        description: "Per-period configuration values (payroll date, accounting period, etc.)",
        requiresValue2: false
    },
    "step-notes": {
        label: "Step Notes",
        description: "Notes and sign-off data for workflow steps",
        requiresValue2: false
    },
    "shared": {
        label: "Shared",
        description: "Global settings shared across all modules",
        requiresValue2: false
    },
    "column-mapping": {
        label: "Column Mapping",
        description: "Maps source columns to target columns for data import",
        requiresValue2: false
    }
};

// Get list of valid category values (normalized)
export function getValidCategoryKeys() {
    return Object.keys(VALID_CATEGORIES);
}

// Validate a single config row
export function validateConfigRow(row, rowIndex) {
    const errors = [];
    const warnings = [];
    
    const category = normalizeModuleToken(row.category || row[0] || "");
    const field = String(row.field || row[1] || "").trim();
    const value = String(row.value || row[2] || "").trim();
    const value2 = String(row.value2 || row[3] || "").trim();
    
    // Check category is valid
    if (!category) {
        errors.push(`Row ${rowIndex}: Missing Category`);
    } else if (!VALID_CATEGORIES[category]) {
        errors.push(`Row ${rowIndex}: Invalid Category "${row.category || row[0]}". Valid options: ${Object.values(VALID_CATEGORIES).map(c => c.label).join(", ")}`);
    }
    
    // Check Field is not empty
    if (!field) {
        errors.push(`Row ${rowIndex}: Missing Field name`);
    }
    
    // Check Value is not empty (except for certain cases)
    if (!value && category !== "step-notes") {
        warnings.push(`Row ${rowIndex}: Value is empty for "${field}"`);
    }
    
    // Check Value2 rules
    const catConfig = VALID_CATEGORIES[category];
    if (catConfig) {
        if (catConfig.requiresValue2 && !value2) {
            errors.push(`Row ${rowIndex}: "${catConfig.label}" requires Value2 (module key)`);
        }
        if (!catConfig.requiresValue2 && value2) {
            warnings.push(`Row ${rowIndex}: Value2 should be empty for "${catConfig.label}" category (found: "${value2}")`);
        }
        if (catConfig.requiresValue2 && value2 && catConfig.value2Options) {
            const normalizedValue2 = normalizeModuleToken(value2);
            if (!catConfig.value2Options.includes(normalizedValue2)) {
                warnings.push(`Row ${rowIndex}: Value2 "${value2}" may not be recognized. Expected: ${catConfig.value2Options.join(", ")}`);
            }
        }
    }
    
    return { errors, warnings, isValid: errors.length === 0 };
}

// Validate entire config sheet
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
            results.rowCount = rows.length - 1; // Exclude header
            
            // Validate each data row
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const validation = validateConfigRow({
                    category: row[0],
                    field: row[1],
                    value: row[2],
                    value2: row[3]
                }, i + 1); // +1 for Excel row number
                
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

const MODULE_ALIAS_MAP = {
    "payroll-recorder": ["payroll-recorder", "payroll", "payroll recorder", "payroll review", "pr"],
    "employee-roster": ["employee-roster", "employee roster", "headcount", "headcount review", "roster"],
    "pto-accrual": ["pto-accrual", "pto", "pto accrual", "pto review"]
};

export async function applyModuleTabVisibility(moduleKey, { aliasTokens = [] } = {}) {
    if (!hasExcelRuntime()) return;
    
    const normalizedModuleKey = normalizeModuleToken(moduleKey);
    console.log(`[Tab Visibility] Applying visibility for module: ${normalizedModuleKey}`);
    
    // Check if we have explicit configuration for this module
    const moduleConfig = MODULE_TAB_CONFIG[normalizedModuleKey];
    
    if (moduleConfig) {
        // Use explicit module configuration
        await applyExplicitModuleVisibility(moduleConfig, normalizedModuleKey);
    } else {
        // Fall back to SS_PF_Config-based visibility
        await applyConfigBasedVisibility(moduleKey, aliasTokens);
    }
}

/**
 * Apply explicit module visibility based on MODULE_TAB_CONFIG
 * Shows specific tabs and hides others for the selected module
 */
async function applyExplicitModuleVisibility(moduleConfig, moduleKey) {
    const visibleTabs = (moduleConfig.visible || []).map(s => normalizeSheetName(s));
    const hiddenTabs = (moduleConfig.hidden || []).map(s => normalizeSheetName(s));
    
    console.log(`[Tab Visibility] Explicit config for ${moduleKey}:`);
    console.log(`  - Visible: ${moduleConfig.visible.join(", ")}`);
    console.log(`  - Hidden: ${moduleConfig.hidden.join(", ")}`);
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            // Track visibility changes
            const toShow = [];
            const toHide = [];
            
            worksheets.items.forEach((sheet) => {
                const normalizedName = normalizeSheetName(sheet.name);
                
                if (visibleTabs.includes(normalizedName)) {
                    toShow.push(sheet);
                } else if (hiddenTabs.includes(normalizedName)) {
                    toHide.push(sheet);
                }
                // Other tabs not in either list are left as-is
            });
            
            // First, show all tabs that should be visible
            for (const sheet of toShow) {
                sheet.visibility = Excel.SheetVisibility.visible;
                console.log(`[Tab Visibility] SHOW: ${sheet.name}`);
            }
            await context.sync();
            
            // Then hide tabs that should be hidden (ensure at least one visible)
            const visibleCount = worksheets.items.filter(
                s => s.visibility === Excel.SheetVisibility.visible
            ).length;
            
            if (visibleCount > toHide.length) {
                for (const sheet of toHide) {
                    try {
                        sheet.visibility = Excel.SheetVisibility.hidden;
                        console.log(`[Tab Visibility] HIDE: ${sheet.name}`);
                    } catch (e) {
                        console.warn(`[Tab Visibility] Could not hide "${sheet.name}":`, e.message);
                    }
                }
                await context.sync();
            } else {
                console.warn("[Tab Visibility] Skipping hide - would leave no visible sheets");
            }
            
            console.log(`[Tab Visibility] Applied visibility for ${moduleKey}`);
        });
    } catch (error) {
        console.warn(`[Tab Visibility] Error applying visibility for ${moduleKey}:`, error);
    }
}

/**
 * Fall back to SS_PF_Config-based visibility for modules without explicit config
 */
async function applyConfigBasedVisibility(moduleKey, aliasTokens = []) {
    const aliasSet = buildAliasSet([...getAliasTokens(moduleKey), ...aliasTokens]);
    console.log(`[Tab Visibility] Using config-based visibility. Module: ${moduleKey}, Aliases:`, [...aliasSet]);
    
    try {
        await Excel.run(async (context) => {
            const configSheet = context.workbook.worksheets.getItemOrNullObject(CONFIG_SHEET_NAME);
            await context.sync();
            if (configSheet.isNullObject) {
                console.warn(`Config sheet ${CONFIG_SHEET_NAME} is missing; skipping tab visibility.`);
                return;
            }

            const usedRange = configSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();

            if (usedRange.isNullObject) {
                console.warn(`${CONFIG_SHEET_NAME} does not contain any values yet.`);
                return;
            }

            const values = usedRange.values || [];
            if (!values.length) return;

            const headerMap = buildHeaderMap(values[0]);
            const categoryIdx = headerMap.get("category");
            const fieldIdx = headerMap.get("field");
            const valueIdx = headerMap.get("value");
            const value2Idx = headerMap.get("value2"); // Module key column

            console.log(`[Tab Visibility] Headers - Category: ${categoryIdx}, Field: ${fieldIdx}, Value: ${valueIdx}, Value2: ${value2Idx}`);

            if (categoryIdx === undefined || fieldIdx === undefined || valueIdx === undefined) {
                console.warn("SS_PF_Config needs Category, Field, and Value columns to drive tab visibility.");
                return;
            }

            // Structure based on actual SS_PF_Config:
            // Category = "Tab Structure"
            // Field = Module description (e.g., "Payroll Recorder", "PTO Accrual")
            // Value = TAB NAME (e.g., "PR_Data", "PTO_Analysis")
            // Value2 = MODULE KEY (e.g., "payroll-recorder", "pto-accrual")
            const tabRules = values
                .slice(1)
                .map((row) => {
                    const category = normalizeCategoryValue(row[categoryIdx]);
                    // Value column contains the TAB NAME
                    const tabName = String(row[valueIdx] ?? "").trim();
                    // Value2 column contains the MODULE KEY
                    const moduleValue = value2Idx !== undefined 
                        ? String(row[value2Idx] ?? "").trim()
                        : "";
                    return {
                        category,
                        tabName,
                        normalizedTabName: normalizeSheetName(tabName),
                        moduleValue
                    };
                })
                .filter((row) => row.tabName && row.category === TAB_STRUCTURE_CATEGORY);

            console.log(`[Tab Visibility] Found ${tabRules.length} tab-structure rules:`, 
                tabRules.map(r => `${r.tabName} → ${r.moduleValue}`));

            if (!tabRules.length) {
                console.warn("No rows found in SS_PF_Config for Tab Structure.");
                return;
            }

            const ruleMap = new Map();
            tabRules.forEach((rule) => {
                if (!rule.normalizedTabName) return;
                ruleMap.set(rule.normalizedTabName, rule);
            });

            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();

            // Get normalized system sheet names for comparison
            const systemSheetNames = SYSTEM_SHEETS.map(s => normalizeSheetName(s));
            
            // First pass: determine what to show/hide
            const visibilityChanges = [];
            let willBeVisibleCount = 0;
            
            worksheets.items.forEach((sheet) => {
                const normalizedSheetName = normalizeSheetName(sheet.name);
                if (!normalizedSheetName) return;
                
                // Skip system sheets - they are handled by hideSystemSheets()
                if (systemSheetNames.includes(normalizedSheetName)) {
                    console.log(`[Tab Visibility] Skipping system sheet: "${sheet.name}"`);
                    return;
                }
                
                const rule = ruleMap.get(normalizedSheetName);
                if (!rule) {
                    // No rule for this sheet - leave it as-is
                    console.log(`[Tab Visibility] No rule for "${sheet.name}" - leaving as-is`);
                    if (sheet.visibility === Excel.SheetVisibility.visible) {
                        willBeVisibleCount++;
                    }
                    return;
                }
                
                const shouldShow = moduleValueMatches(rule.moduleValue, aliasSet);
                console.log(`[Tab Visibility] "${sheet.name}" (module: ${rule.moduleValue}) → ${shouldShow ? "SHOW" : "HIDE"}`);
                
                if (shouldShow) {
                    willBeVisibleCount++;
                }
                visibilityChanges.push({ sheet, shouldShow });
            });
            
            // Second pass: apply changes, ensuring at least one sheet stays visible
            console.log(`[Tab Visibility] ${willBeVisibleCount} sheets will be visible after changes`);
            
            // First show all sheets that should be visible
            for (const change of visibilityChanges) {
                if (change.shouldShow) {
                    change.sheet.visibility = Excel.SheetVisibility.visible;
                }
            }
            await context.sync();
            
            // Then hide sheets that shouldn't be visible (only if there will be visible sheets)
            if (willBeVisibleCount > 0) {
                for (const change of visibilityChanges) {
                    if (!change.shouldShow) {
                        try {
                            change.sheet.visibility = Excel.SheetVisibility.hidden;
                        } catch (e) {
                            console.warn(`[Tab Visibility] Could not hide "${change.sheet.name}":`, e.message);
                        }
                    }
                }
                await context.sync();
            } else {
                console.warn("[Tab Visibility] No sheets would be visible - skipping hide operations");
            }
        });
    } catch (error) {
        console.warn(`Unable to toggle worksheet visibility for ${moduleKey}:`, error);
    }
}

function buildHeaderMap(headers = []) {
    const map = new Map();
    headers.forEach((header, index) => {
        const normalized = normalizeCategoryValue(header);
        if (normalized) {
            map.set(normalized, index);
        }
    });
    return map;
}

function normalizeCategoryValue(value) {
    return normalizeModuleToken(value);
}

function normalizeSheetName(name) {
    return String(name ?? "").trim().toLowerCase();
}

function splitModuleTokens(value) {
    return String(value ?? "")
        .split(/[,;|/&]+/)
        .map((token) => normalizeModuleToken(token))
        .filter(Boolean);
}

function normalizeModuleToken(value) {
    return String(value ?? "")
        .trim()
        .toLowerCase()
        .replace(/[\s_]+/g, "-");
}

function buildAliasSet(values) {
    const normalizedValues = (values || []).map((value) => normalizeModuleToken(value)).filter(Boolean);
    if (!normalizedValues.length) return new Set();
    return new Set(normalizedValues);
}

function getAliasTokens(moduleKey) {
    const normalizedKey = normalizeModuleToken(moduleKey);
    return MODULE_ALIAS_MAP[normalizedKey] ?? [normalizedKey];
}

function moduleValueMatches(value, aliasSet) {
    const tokens = splitModuleTokens(value);
    if (!tokens.length) return true;
    return tokens.some((token) => ALWAYS_SHOW_TOKENS.has(token) || aliasSet.has(token));
}

// hasExcelRuntime imported from gateway.js

/**
 * Hide system/configuration sheets when workbook opens.
 * These sheets can be manually unhidden by the user during their session,
 * but will be hidden again on next workbook open.
 * 
 * System sheets: SS_PF_Config, SS_Employee_Roster, SS_Chart_of_Accounts
 * 
 * @param {string[]} additionalSheets - Optional additional sheet names to hide
 * @returns {Promise<void>}
 */
export async function hideSystemSheets(additionalSheets = []) {
    if (!hasExcelRuntime()) return;
    
    const sheetsToHide = [...SYSTEM_SHEETS, ...additionalSheets].map(s => normalizeSheetName(s));
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            // Count visible sheets to ensure we don't hide all of them
            const visibleSheets = worksheets.items.filter(
                sheet => sheet.visibility === Excel.SheetVisibility.visible
            );
            
            let hiddenCount = 0;
            
            worksheets.items.forEach((sheet) => {
                const normalizedName = normalizeSheetName(sheet.name);
                
                // Only hide if this is a system sheet AND there will still be visible sheets
                if (sheetsToHide.includes(normalizedName)) {
                    // Ensure at least one sheet stays visible
                    if (visibleSheets.length - hiddenCount > 1) {
                        sheet.visibility = Excel.SheetVisibility.hidden;
                        hiddenCount++;
                        console.log(`[Tab Visibility] Hidden system sheet: ${sheet.name}`);
                    }
                }
            });
            
            await context.sync();
        });
    } catch (error) {
        console.warn("Unable to hide system sheets:", error);
    }
}

/**
 * Get the list of system sheet names
 * @returns {string[]}
 */
export function getSystemSheetNames() {
    return [...SYSTEM_SHEETS];
}

/**
 * Force unhide system sheets (useful when they become inaccessible)
 * Can be called from console: window.PrairieForge.unhideSystemSheets()
 * @returns {Promise<void>}
 */
export async function unhideSystemSheets() {
    if (!hasExcelRuntime()) {
        console.log("Excel not available");
        return;
    }
    
    const sheetsToUnhide = [...SYSTEM_SHEETS];
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            
            worksheets.items.forEach((sheet) => {
                const normalizedName = normalizeSheetName(sheet.name);
                if (sheetsToUnhide.map(s => normalizeSheetName(s)).includes(normalizedName)) {
                    sheet.visibility = Excel.SheetVisibility.visible;
                    console.log(`[Unhide] Made visible: ${sheet.name}`);
                }
            });
            
            await context.sync();
            console.log("[Unhide] System sheets are now visible!");
        });
    } catch (error) {
        console.error("Unable to unhide system sheets:", error);
    }
}

/**
 * Force ALL sheets to be visible (emergency recovery)
 * Can be called from console: window.PrairieForge.showAllSheets()
 * @returns {Promise<void>}
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
            console.log(`[ShowAll] Done! Made ${unhiddenCount} sheets visible. Total sheets: ${worksheets.items.length}`);
        });
    } catch (error) {
        console.error("Unable to show all sheets:", error);
    }
}

// Expose to global scope for console access
if (typeof window !== "undefined") {
    window.PrairieForge = window.PrairieForge || {};
    window.PrairieForge.unhideSystemSheets = unhideSystemSheets;
    window.PrairieForge.showAllSheets = showAllSheets;
}
