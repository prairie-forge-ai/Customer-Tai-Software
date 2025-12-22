/**
 * TaiTools Workbook Initialization
 * Creates and formats all required worksheets for a fresh workbook
 * 
 * © 2025 Prairie Forge LLC
 */

// =============================================================================
// WORKBOOK BLUEPRINT
// =============================================================================

/**
 * Global formatting standards applied to all data sheets
 */
export const FORMATTING_STANDARDS = {
    headerRow: 1,
    headerBackground: "#000000",
    headerTextColor: "#FFFFFF",
    headerFontBold: true,
    freezeRow: 1,
    autoFitColumns: true
};

/**
 * Homepage tab definitions (landing pages with title + subtitle)
 */
export const HOMEPAGE_TABS = [
    {
        name: "SS_Homepage",
        title: "TaiTools",
        subtitle: "Select a module from the side panel to get started.",
        module: "shared"
    },
    {
        name: "PR_Homepage",
        title: "Payroll Recorder",
        subtitle: "Normalize payroll exports, enforce controls, and prep journal entries without leaving Excel.",
        module: "payroll"
    },
    {
        name: "PTO_Homepage",
        title: "PTO Accrual",
        subtitle: "Calculate employee PTO liabilities, compare period-over-period changes, and prepare accrual journal entries.",
        module: "pto"
    }
];

/**
 * Data tab definitions with headers
 */
export const DATA_TABS = [
    // =========================================================================
    // SHARED (SS_) TABS
    // =========================================================================
    {
        name: "SS_PF_Config",
        tableName: "SS_PF_Config",
        module: "shared",
        description: "Central configuration store for all modules",
        headers: ["Category", "Field", "Value", "Permanent"]
    },
    {
        name: "SS_Chart_of_Accounts",
        tableName: "SS_Chart_of_Accounts",
        module: "shared",
        description: "Chart of accounts reference data",
        headers: ["Account_Number", "Account_Name", "Type", "Category"]
    },

    // =========================================================================
    // PTO ACCRUAL (PTO_) TABS
    // =========================================================================
    {
        name: "PTO_Data",
        tableName: "PTO_Data",
        module: "pto",
        description: "Raw PTO data import from provider",
        headers: [
            "Company Name",
            "Form Name",
            "Selection Criteria",
            "Employee Name",
            "Accrue Thru Date",
            "Year Ending",
            "Plan Description",
            "Accrual Rate",
            "Carry Over",
            "Pay Period Accrued",
            "Pay Period Used",
            "YTD Accrued",
            "YTD Used",
            "Balance"
        ]
    },
    {
        name: "PTO_Analysis",
        tableName: "PTO_Analysis",
        module: "pto",
        description: "PTO liability analysis by employee",
        headers: [
            "Analysis Date",
            "Employee Name",
            "Department",
            "Pay Rate",
            "Accrual Rate",
            "Carry Over",
            "YTD Accrued",
            "YTD Used",
            "Balance",
            "Liability Amount",
            "Accrued PTO $ [Prior Period]",
            "Change"
        ]
    },
    {
        name: "PTO_JE_Draft",
        tableName: "PTO_JE_Draft",
        module: "pto",
        description: "PTO journal entry output (QuickBooks format)",
        headers: [
            "RefNumber",
            "TxnDate",
            "Account Number",
            "Account Name",
            "LineAmount",
            "Debit",
            "Credit",
            "LineDesc",
            "Department"
        ]
    },
    {
        name: "PTO_Archive_Summary",
        tableName: "PTO_Archive_Summary",
        module: "pto",
        description: "Historical PTO snapshots per period",
        headers: [
            "Analysis Date",
            "Employee Name",
            "Department",
            "Accrual Rate",
            "Carry Over",
            "Balance",
            "Pay Rate"
        ]
    },

    // =========================================================================
    // PAYROLL RECORDER (PR_) TABS
    // =========================================================================
    {
        name: "PR_Data",
        tableName: "PR_Data",
        module: "payroll",
        description: "Raw payroll data import from provider",
        headers: [
            "Company Name",
            "Report Title",
            "Select Criteria",
            "Pay Date",
            "Department",
            "Department Description",
            "Employee",
            "Regular Earns",
            "OT Earns",
            "Bonus Earns",
            "Commission Earns",
            "PTO Earns",
            "Expense Reimb",
            "Gross Pay",
            "401(k) Match",
            "WC",
            "ER Taxes",
            "Benefits",
            "Admin Fees",
            "Other"
        ]
    },
    {
        name: "PR_Data_Clean",
        tableName: "PR_Data_Clean",
        module: "payroll",
        description: "Normalized payroll data (one row per employee per category)",
        headers: [
            "Payroll Date",
            "Employee",
            "Department",
            "Payroll Category",
            "Account Number",
            "Account Name",
            "Amount",
            "Expense Review"
        ]
    },
    {
        name: "PR_Expense_Review",
        tableName: null, // Script-generated at runtime
        module: "payroll",
        description: "Expense analysis dashboard (generated by workflow)",
        headers: null // No predefined headers - built dynamically
    },
    {
        name: "PR_JE_Draft",
        tableName: "PR_JE_Draft",
        module: "payroll",
        description: "Payroll journal entry output (QuickBooks format)",
        headers: [
            "RefNumber",
            "TxnDate",
            "Account Number",
            "Account Name",
            "LineAmount",
            "Debit",
            "Credit",
            "LineDesc",
            "Department"
        ]
    },
    {
        name: "PR_Archive_Summary",
        tableName: "PR_Archive_Summary",
        module: "payroll",
        description: "Historical payroll snapshots per period",
        headers: [
            "Payroll Date",
            "Employee",
            "Department",
            "Payroll Category",
            "Account Number",
            "Account Name",
            "Amount",
            "Expense Review"
        ]
    },
    {
        name: "PR_Expense_Mapping",
        tableName: "PR_Expense_Mapping",
        module: "payroll",
        description: "Maps Department + Category to Account + Expense Review",
        headers: [
            "Department Name",
            "Account Name",
            "Payroll Category",
            "Account Number",
            "Expense Review"
        ]
    }
];

/**
 * Get all tabs for a specific module
 */
export function getTabsForModule(moduleKey) {
    const homepages = HOMEPAGE_TABS.filter(t => t.module === moduleKey || t.module === "shared");
    const dataTabs = DATA_TABS.filter(t => t.module === moduleKey || t.module === "shared");
    return [...homepages, ...dataTabs];
}

/**
 * Get all tab names
 */
export function getAllTabNames() {
    const homepageNames = HOMEPAGE_TABS.map(t => t.name);
    const dataTabNames = DATA_TABS.map(t => t.name);
    return [...homepageNames, ...dataTabNames];
}

// =============================================================================
// INITIALIZATION FUNCTIONS
// =============================================================================

/**
 * Check if Excel runtime is available
 */
function hasExcelRuntime() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

/**
 * Convert column index to Excel column letter (0 = A, 1 = B, etc.)
 */
function columnIndexToLetter(index) {
    let letter = "";
    let temp = index;
    while (temp >= 0) {
        letter = String.fromCharCode((temp % 26) + 65) + letter;
        temp = Math.floor(temp / 26) - 1;
    }
    return letter;
}

/**
 * Apply standard header formatting to a range
 */
async function formatHeaderRange(sheet, headerCount, context) {
    const lastColumn = columnIndexToLetter(headerCount - 1);
    const headerRange = sheet.getRange(`A1:${lastColumn}1`);
    
    // Apply formatting
    headerRange.format.fill.color = FORMATTING_STANDARDS.headerBackground;
    headerRange.format.font.color = FORMATTING_STANDARDS.headerTextColor;
    headerRange.format.font.bold = FORMATTING_STANDARDS.headerFontBold;
    
    // Auto-fit columns
    headerRange.format.autofitColumns();
    
    // Freeze header row
    sheet.freezePanes.freezeRows(FORMATTING_STANDARDS.freezeRow);
    
    await context.sync();
}

/**
 * Create a homepage tab with title and subtitle
 */
async function createHomepageTab(tabDef, context) {
    // Check if sheet already exists
    let sheet = context.workbook.worksheets.getItemOrNullObject(tabDef.name);
    sheet.load("isNullObject");
    await context.sync();
    
    if (!sheet.isNullObject) {
        console.log(`[Init] Sheet ${tabDef.name} already exists, skipping`);
        return { created: false, name: tabDef.name };
    }
    
    // Create the sheet
    sheet = context.workbook.worksheets.add(tabDef.name);
    
    // Set title in A1 (merged across columns A-H for visibility)
    const titleRange = sheet.getRange("A1:H1");
    titleRange.merge();
    titleRange.values = [[tabDef.title]];
    titleRange.format.font.size = 36;
    titleRange.format.font.color = "#FFFFFF";
    titleRange.format.font.bold = true;
    titleRange.format.verticalAlignment = "Center";
    titleRange.format.rowHeight = 60;
    
    // Set subtitle in A2
    const subtitleRange = sheet.getRange("A2:H2");
    subtitleRange.merge();
    subtitleRange.values = [[tabDef.subtitle]];
    subtitleRange.format.font.size = 12;
    subtitleRange.format.font.color = "#CCCCCC";
    subtitleRange.format.verticalAlignment = "Top";
    subtitleRange.format.rowHeight = 30;
    
    // Black background for entire visible area
    const bgRange = sheet.getRange("A1:Z50");
    bgRange.format.fill.color = "#000000";
    
    await context.sync();
    
    console.log(`[Init] Created homepage: ${tabDef.name}`);
    return { created: true, name: tabDef.name };
}

/**
 * Create a data tab with headers and optional table formatting
 */
async function createDataTab(tabDef, context) {
    // Check if sheet already exists
    let sheet = context.workbook.worksheets.getItemOrNullObject(tabDef.name);
    sheet.load("isNullObject");
    await context.sync();
    
    if (!sheet.isNullObject) {
        console.log(`[Init] Sheet ${tabDef.name} already exists, skipping`);
        return { created: false, name: tabDef.name };
    }
    
    // Create the sheet
    sheet = context.workbook.worksheets.add(tabDef.name);
    
    // If no headers defined (script-generated), just create empty sheet with black row 1
    if (!tabDef.headers || tabDef.headers.length === 0) {
        const emptyHeaderRange = sheet.getRange("A1:Z1");
        emptyHeaderRange.format.fill.color = FORMATTING_STANDARDS.headerBackground;
        sheet.freezePanes.freezeRows(1);
        await context.sync();
        console.log(`[Init] Created empty tab: ${tabDef.name} (headers generated at runtime)`);
        return { created: true, name: tabDef.name };
    }
    
    // Set headers
    const lastColumn = columnIndexToLetter(tabDef.headers.length - 1);
    const headerRange = sheet.getRange(`A1:${lastColumn}1`);
    headerRange.values = [tabDef.headers];
    
    // Apply standard formatting
    await formatHeaderRange(sheet, tabDef.headers.length, context);
    
    // Create formal Excel Table if tableName is specified
    if (tabDef.tableName) {
        try {
            // Define table range (headers + 1 empty data row to start)
            const tableRange = sheet.getRange(`A1:${lastColumn}2`);
            const table = sheet.tables.add(tableRange, true /* hasHeaders */);
            table.name = tabDef.tableName;
            
            // Style the table
            table.style = "TableStyleMedium2"; // Dark header style
            
            await context.sync();
            console.log(`[Init] Created table: ${tabDef.tableName}`);
        } catch (tableError) {
            console.warn(`[Init] Could not create table ${tabDef.tableName}:`, tableError);
        }
    }
    
    console.log(`[Init] Created data tab: ${tabDef.name}`);
    return { created: true, name: tabDef.name };
}

/**
 * Initialize the complete workbook with all TaiTools tabs
 * 
 * @param {Object} options - Configuration options
 * @param {string} options.module - "all", "payroll", "pto", or "shared"
 * @param {boolean} options.skipExisting - If true, skip tabs that already exist
 * @param {Function} options.onProgress - Progress callback (step, total, message)
 * @returns {Promise<Object>} Results summary
 */
export async function initializeWorkbook(options = {}) {
    const {
        module = "all",
        skipExisting = true,
        onProgress = null
    } = options;
    
    if (!hasExcelRuntime()) {
        throw new Error("Excel runtime is not available");
    }
    
    const results = {
        created: [],
        skipped: [],
        errors: []
    };
    
    // Determine which tabs to create
    let homepagesToCreate = HOMEPAGE_TABS;
    let dataTabsToCreate = DATA_TABS;
    
    if (module !== "all") {
        homepagesToCreate = HOMEPAGE_TABS.filter(t => t.module === module || t.module === "shared");
        dataTabsToCreate = DATA_TABS.filter(t => t.module === module || t.module === "shared");
    }
    
    const totalTabs = homepagesToCreate.length + dataTabsToCreate.length;
    let currentStep = 0;
    
    try {
        await Excel.run(async (context) => {
            // Create homepage tabs first
            for (const tabDef of homepagesToCreate) {
                currentStep++;
                if (onProgress) {
                    onProgress(currentStep, totalTabs, `Creating ${tabDef.name}...`);
                }
                
                try {
                    const result = await createHomepageTab(tabDef, context);
                    if (result.created) {
                        results.created.push(result.name);
                    } else {
                        results.skipped.push(result.name);
                    }
                } catch (error) {
                    console.error(`[Init] Error creating ${tabDef.name}:`, error);
                    results.errors.push({ name: tabDef.name, error: error.message });
                }
            }
            
            // Create data tabs
            for (const tabDef of dataTabsToCreate) {
                currentStep++;
                if (onProgress) {
                    onProgress(currentStep, totalTabs, `Creating ${tabDef.name}...`);
                }
                
                try {
                    const result = await createDataTab(tabDef, context);
                    if (result.created) {
                        results.created.push(result.name);
                    } else {
                        results.skipped.push(result.name);
                    }
                } catch (error) {
                    console.error(`[Init] Error creating ${tabDef.name}:`, error);
                    results.errors.push({ name: tabDef.name, error: error.message });
                }
            }
            
            // Activate the homepage
            const homepage = module === "pto" ? "PTO_Homepage" 
                           : module === "payroll" ? "PR_Homepage" 
                           : "SS_Homepage";
            
            const homeSheet = context.workbook.worksheets.getItemOrNullObject(homepage);
            homeSheet.load("isNullObject");
            await context.sync();
            
            if (!homeSheet.isNullObject) {
                homeSheet.activate();
                await context.sync();
            }
        });
        
        if (onProgress) {
            onProgress(totalTabs, totalTabs, "Initialization complete!");
        }
        
    } catch (error) {
        console.error("[Init] Workbook initialization failed:", error);
        throw error;
    }
    
    return results;
}

/**
 * Validate workbook structure - check which tabs exist/are missing
 * 
 * @returns {Promise<Object>} Validation results
 */
export async function validateWorkbookStructure() {
    if (!hasExcelRuntime()) {
        throw new Error("Excel runtime is not available");
    }
    
    const allTabs = getAllTabNames();
    const validation = {
        present: [],
        missing: [],
        extra: []
    };
    
    try {
        await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();
            
            const existingNames = sheets.items.map(s => s.name);
            
            // Check which expected tabs exist
            for (const tabName of allTabs) {
                if (existingNames.includes(tabName)) {
                    validation.present.push(tabName);
                } else {
                    validation.missing.push(tabName);
                }
            }
            
            // Find extra tabs not in our blueprint
            for (const sheetName of existingNames) {
                if (!allTabs.includes(sheetName)) {
                    validation.extra.push(sheetName);
                }
            }
        });
    } catch (error) {
        console.error("[Validate] Structure validation failed:", error);
        throw error;
    }
    
    return validation;
}

/**
 * Repair workbook - create only missing tabs
 * 
 * @param {Function} onProgress - Progress callback
 * @returns {Promise<Object>} Results summary
 */
export async function repairWorkbook(onProgress = null) {
    const validation = await validateWorkbookStructure();
    
    if (validation.missing.length === 0) {
        console.log("[Repair] All tabs present, nothing to repair");
        return { created: [], skipped: validation.present, errors: [] };
    }
    
    console.log(`[Repair] Missing tabs: ${validation.missing.join(", ")}`);
    
    // Initialize with all modules but skipExisting will handle it
    return await initializeWorkbook({
        module: "all",
        skipExisting: true,
        onProgress
    });
}





