import { applyModuleTabVisibility, showAllSheets, showAndActivateSheet } from "../../Common/tab-visibility.js";
import { bindInstructionsButton } from "../../Common/instructions.js";
import { activateHomepageSheet, getHomepageConfig, renderAdaFab, removeAdaFab } from "../../Common/homepage-sheet.js";
import { renderCopilotCard, bindCopilotCard, createExcelContextProvider } from "../../Common/copilot.js";
import { initDatePicker } from "../../Common/date-picker.js";
import * as XLSX from "xlsx";
import {
    HOME_ICON_SVG,
    MODULES_ICON_SVG,
    ARROW_LEFT_SVG,
    USERS_ICON_SVG,
    BOOK_ICON_SVG,
    ARROW_RIGHT_SVG,
    MENU_ICON_SVG,
    LOCK_CLOSED_SVG,
    LOCK_OPEN_SVG,
    CHECK_ICON_SVG,
    X_CIRCLE_SVG,
    X_ICON_SVG,
    CALCULATOR_ICON_SVG,
    LINK_ICON_SVG,
    SAVE_ICON_SVG,
    TABLE_ICON_SVG,
    UPLOAD_ICON_SVG,
    DOWNLOAD_ICON_SVG,
    REFRESH_ICON_SVG,
    TRASH_ICON_SVG,
    FILE_TEXT_ICON_SVG,
    SETTINGS_ICON_SVG,
    GLOBE_ICON_SVG,
    getStepIconSvg
} from "../../Common/icons.js";
import { renderInlineNotes, renderSignoff, renderLabeledButton, updateLockButtonVisual, updateSaveButtonState, initSaveTracking } from "../../Common/notes-signoff.js";
import { canCompleteStep, showBlockedToast } from "../../Common/workflow-validation.js";
import { loadConfigFromTable, saveConfigValue, hasExcelRuntime } from "../../Common/gateway.js";
import { formatSheetHeaders, formatCurrencyColumn, formatNumberColumn, formatDateColumn, NUMBER_FORMATS } from "../../Common/sheet-formatting.js";
import { formatXlsxWorksheet, setXlsxColumnWidths, XLSX_COLUMN_WIDTHS } from "../../Common/xlsx-formatting.js";

// Build script injects the current commit hash at bundle time
// Fallback to "dev" when running outside bundle (tests, lint)
const MODULE_VERSION = typeof __BUILD_COMMIT__ !== "undefined" ? __BUILD_COMMIT__ : "dev";
const MODULE_KEY = "pto-accrual";
const MODULE_ALIAS_TOKENS = ["pto", "pto-accrual", "pto review", "accrual"];
const MODULE_NAME = "PTO Accrual";

// Supabase configuration
const SUPABASE_URL = "https://jgciqwzwacaesqjaoadc.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE";

// =============================================================================
// OFFICE-SAFE DIALOGS (window.alert not supported in Office Add-ins)
// =============================================================================

/**
 * Show a toast notification in the UI (Office-safe alternative to window.alert)
 */
function showToast(message, type = "info", duration = 4000) {
    // Remove existing toasts
    document.querySelectorAll(".pf-toast").forEach(t => t.remove());
    
    const toast = document.createElement("div");
    toast.className = `pf-toast pf-toast--${type}`;
    toast.innerHTML = `
        <div class="pf-toast-content">
            <span class="pf-toast-icon">${type === "success" ? "✅" : type === "error" ? "❌" : "ℹ️"}</span>
            <span class="pf-toast-message">${message.replace(/\n/g, "<br>")}</span>
        </div>
        <button class="pf-toast-close" onclick="this.parentElement.remove()">×</button>
    `;
    
    // Add styles if not already present
    if (!document.getElementById("pf-toast-styles")) {
        const style = document.createElement("style");
        style.id = "pf-toast-styles";
        style.textContent = `
            .pf-toast {
                position: fixed;
                top: 20px;
                left: 50%;
                transform: translateX(-50%);
                background: #1a1a2e;
                color: white;
                padding: 16px 20px;
                border-radius: 8px;
                box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                z-index: 10000;
                max-width: 90%;
                display: flex;
                align-items: flex-start;
                gap: 12px;
                animation: toastIn 0.3s ease;
            }
            .pf-toast--success { border-left: 4px solid #22c55e; }
            .pf-toast--error { border-left: 4px solid #ef4444; }
            .pf-toast--info { border-left: 4px solid #3b82f6; }
            .pf-toast-content { display: flex; align-items: flex-start; gap: 8px; flex: 1; }
            .pf-toast-icon { font-size: 18px; }
            .pf-toast-message { font-size: 14px; line-height: 1.4; }
            .pf-toast-close { background: none; border: none; color: #888; font-size: 20px; cursor: pointer; padding: 0; margin-left: 8px; }
            .pf-toast-close:hover { color: white; }
            @keyframes toastIn { from { opacity: 0; transform: translateX(-50%) translateY(-20px); } }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(toast);
    
    if (duration > 0) {
        setTimeout(() => toast.remove(), duration);
    }
    
    return toast;
}

/**
 * Show a confirmation dialog (Office-safe alternative to window.confirm)
 * Apple-inspired design with glassmorphism
 * Returns a Promise that resolves to true/false
 */
function showConfirm(message, options = {}) {
    const {
        title = "Confirm Action",
        confirmText = "Continue",
        cancelText = "Cancel",
        icon = "📋",
        destructive = false
    } = options;
    
    return new Promise((resolve) => {
        // Remove existing dialogs
        document.querySelectorAll(".pf-confirm-overlay").forEach(d => d.remove());
        
        const overlay = document.createElement("div");
        overlay.className = "pf-confirm-overlay";
        overlay.innerHTML = `
            <div class="pf-confirm-dialog">
                <div class="pf-confirm-icon">${icon}</div>
                <div class="pf-confirm-title">${title}</div>
                <div class="pf-confirm-message">${message.replace(/\n/g, "<br>")}</div>
                <div class="pf-confirm-buttons">
                    <button class="pf-confirm-btn pf-confirm-btn--cancel">${cancelText}</button>
                    <button class="pf-confirm-btn pf-confirm-btn--ok ${destructive ? 'pf-confirm-btn--destructive' : ''}">${confirmText}</button>
                </div>
            </div>
        `;
        
        // Add styles if not already present
        if (!document.getElementById("pf-confirm-styles")) {
            const style = document.createElement("style");
            style.id = "pf-confirm-styles";
            style.textContent = `
                .pf-confirm-overlay {
                    position: fixed;
                    inset: 0;
                    background: rgba(0, 0, 0, 0.5);
                    backdrop-filter: blur(8px);
                    -webkit-backdrop-filter: blur(8px);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    z-index: 10001;
                    animation: pf-confirm-fade-in 0.2s ease;
                }
                @keyframes pf-confirm-fade-in {
                    from { opacity: 0; }
                    to { opacity: 1; }
                }
                @keyframes pf-confirm-scale-in {
                    from { opacity: 0; transform: scale(0.95) translateY(-10px); }
                    to { opacity: 1; transform: scale(1) translateY(0); }
                }
                .pf-confirm-dialog {
                    background: linear-gradient(145deg, rgba(30, 30, 50, 0.95), rgba(20, 20, 35, 0.98));
                    border: 1px solid rgba(255, 255, 255, 0.08);
                    color: white;
                    padding: 28px 32px;
                    border-radius: 20px;
                    max-width: 380px;
                    width: 90%;
                    box-shadow: 
                        0 24px 48px rgba(0, 0, 0, 0.4),
                        0 0 0 1px rgba(255, 255, 255, 0.05) inset,
                        0 1px 0 rgba(255, 255, 255, 0.1) inset;
                    text-align: center;
                    animation: pf-confirm-scale-in 0.25s cubic-bezier(0.34, 1.56, 0.64, 1);
                }
                .pf-confirm-icon {
                    font-size: 48px;
                    margin-bottom: 16px;
                    filter: drop-shadow(0 4px 8px rgba(0,0,0,0.3));
                }
                .pf-confirm-title {
                    font-size: 18px;
                    font-weight: 600;
                    color: #fff;
                    margin-bottom: 12px;
                    letter-spacing: -0.3px;
                }
                .pf-confirm-message {
                    font-size: 14px;
                    line-height: 1.6;
                    color: rgba(255, 255, 255, 0.7);
                    margin-bottom: 24px;
                    text-align: left;
                }
                .pf-confirm-buttons {
                    display: flex;
                    gap: 12px;
                    justify-content: center;
                }
                .pf-confirm-btn {
                    flex: 1;
                    padding: 12px 24px;
                    border-radius: 12px;
                    border: none;
                    cursor: pointer;
                    font-size: 15px;
                    font-weight: 600;
                    letter-spacing: -0.2px;
                    transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1);
                }
                .pf-confirm-btn:active {
                    transform: scale(0.97);
                }
                .pf-confirm-btn--cancel {
                    background: rgba(255, 255, 255, 0.08);
                    color: rgba(255, 255, 255, 0.9);
                    border: 1px solid rgba(255, 255, 255, 0.1);
                }
                .pf-confirm-btn--cancel:hover {
                    background: rgba(255, 255, 255, 0.12);
                    border-color: rgba(255, 255, 255, 0.15);
                }
                .pf-confirm-btn--ok {
                    background: linear-gradient(145deg, #6366f1, #4f46e5);
                    color: white;
                    box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4);
                }
                .pf-confirm-btn--ok:hover {
                    background: linear-gradient(145deg, #818cf8, #6366f1);
                    box-shadow: 0 6px 16px rgba(99, 102, 241, 0.5);
                    transform: translateY(-1px);
                }
                .pf-confirm-btn--destructive {
                    background: linear-gradient(145deg, #ef4444, #dc2626);
                    box-shadow: 0 4px 12px rgba(239, 68, 68, 0.4);
                }
                .pf-confirm-btn--destructive:hover {
                    background: linear-gradient(145deg, #f87171, #ef4444);
                    box-shadow: 0 6px 16px rgba(239, 68, 68, 0.5);
                }
            `;
            document.head.appendChild(style);
        }
        
        document.body.appendChild(overlay);
        
        // Close on overlay click (outside dialog)
        overlay.addEventListener("click", (e) => {
            if (e.target === overlay) {
                overlay.remove();
                resolve(false);
            }
        });
        
        overlay.querySelector(".pf-confirm-btn--cancel").onclick = () => {
            overlay.remove();
            resolve(false);
        };
        overlay.querySelector(".pf-confirm-btn--ok").onclick = () => {
            overlay.remove();
            resolve(true);
        };
    });
}
const HERO_COPY =
    "Calculate your PTO liability, compare against last period, and generate a balanced journal entry—all without leaving Excel.";
const SELECTOR_URL = "../module-selector/index.html";
const LOADER_ID = "pf-loader-overlay";
const CONFIG_TABLES = ["SS_PF_Config"];
// PTO Config Fields - Pattern: PTO_{Descriptor}
const PTO_CONFIG_FIELDS = {
    payrollProvider: "PTO_Payroll_Provider",
    payrollDate: "PTO_Analysis_Date",
    accountingPeriod: "PTO_Accounting_Period",
    journalEntryId: "PTO_Journal_Entry_ID",
    companyName: "SS_Company_Name",           // Shared field
    accountingSoftware: "SS_Accounting_Software", // Shared field
    reviewerName: "PTO_Reviewer",
    validationDataBalance: "PTO_Validation_Data_Balance",
    validationCleanBalance: "PTO_Validation_Clean_Balance",
    validationDifference: "PTO_Validation_Difference",
    headcountRosterCount: "PTO_Headcount_Roster_Count",
    headcountPayrollCount: "PTO_Headcount_Payroll_Count",
    headcountDifference: "PTO_Headcount_Difference",
    journalDebitTotal: "PTO_JE_Debit_Total",
    journalCreditTotal: "PTO_JE_Credit_Total",
    journalDifference: "PTO_JE_Difference"
};
const HEADCOUNT_SKIP_NOTE = "User opted to skip the headcount review this period.";
// Step notes/sign-off fields - Pattern: PTO_{Type}_{StepName}
// Matches payroll-recorder 5-step structure (0-4)
const STEP_CONFIG_FIELDS = {
    0: { note: "PTO_Notes_Config", reviewer: "PTO_Reviewer_Config", signOff: "PTO_SignOff_Config" },
    1: { note: "PTO_Notes_Upload", reviewer: "PTO_Reviewer_Upload", signOff: "PTO_SignOff_Upload" },
    2: { note: "PTO_Notes_Review", reviewer: "PTO_Reviewer_Review", signOff: "PTO_SignOff_Review" },
    3: { note: "PTO_Notes_JE", reviewer: "PTO_Reviewer_JE", signOff: "PTO_SignOff_JE" },
    4: { note: "PTO_Notes_Archive", reviewer: "PTO_Reviewer_Archive", signOff: "PTO_SignOff_Archive" }
};
const STEP_COMPLETE_FIELDS = {
    0: "PTO_Complete_Config",
    1: "PTO_Complete_Upload",
    2: "PTO_Complete_Review",
    3: "PTO_Complete_JE",
    4: "PTO_Complete_Archive"
};

const PTO_ACTIVITY_COLUMNS = [
    { key: "employeeId", header: "Employee ID" },
    { key: "employeeName", header: "Employee Name" },
    { key: "actionDate", header: "Action Date" },
    { key: "actionType", header: "Action" },
    { key: "hours", header: "Hours" },
    { key: "notes", header: "Notes" },
    { key: "source", header: "Source" }
];

// REMOVED: PTO_ANALYSIS_COLUMNS - legacy sheet no longer used
// All analysis is now done in PTO_Review sheet

const PTO_EXPENSE_COLUMNS = [
    { key: "department", header: "Department" },
    { key: "currentPeriod", header: "Current Period" },
    { key: "priorPeriod", header: "Prior Period" },
    { key: "variance", header: "Variance" },
    { key: "comment", header: "Comment" }
];

const PTO_JOURNAL_COLUMNS = [
    { key: "account", header: "Account" },
    { key: "description", header: "Description" },
    { key: "debit", header: "Debit" },
    { key: "credit", header: "Credit" },
    { key: "reference", header: "Reference" }
];

const WORKFLOW_STEPS = [
    {
        id: 0,
        title: "Configuration",
        summary: "Auto-loaded from system. Review period-specific settings before each run.",
        description: "Configuration is loaded from your installation. Only period-specific fields (analysis date, accounting period) need review each run.",
        actionLabel: "Review Configuration",
        icon: "🧭",
        secondaryAction: { sheet: "SS_PF_Config", label: "Open Config Sheet" }
    },
    {
        id: 1,
        title: "Upload & Validate PTO Data",
        summary: "Import PTO report, normalize headers, and run validation checks.",
        description: "Upload your Obsidian PTO report, auto-map headers using ada_payroll_dimensions, create PTO_Data_Clean, and review advisory validation checks.",
        actionLabel: "Upload PTO Report",
        icon: "📥",
        secondaryAction: { sheet: "PTO_Data_Clean", label: "Open Clean Data" }
    },
    {
        id: 2,
        title: "PTO Accrual Review",
        summary: "Review accrued, used, and balance metrics with executive-ready summary.",
        description: "Analyze PTO data grouped by employee and plan. Review totals, trends, and variances using PTO_Data_Clean as the source of truth.",
        actionLabel: "Generate Review",
        icon: "📊",
        secondaryAction: { sheet: "PTO_Review", label: "Open Review Sheet" }
    },
    {
        id: 3,
        title: "Journal Entry Prep",
        summary: "Generate accounting-ready journal entry output.",
        description: "Create journal entry draft using PTO_Data_Clean and GL mappings. Produces output ready for QuickBooks or your accounting system.",
        actionLabel: "Generate JE Draft",
        icon: "🧾",
        secondaryAction: { sheet: "PTO_JE_Draft", label: "Open JE Draft" }
    },
    {
        id: 4,
        title: "Archive & Clear",
        summary: "Archive the period and reset for next run.",
        description: "Save a snapshot of this period's work, then clear working tabs to prepare for the next PTO analysis cycle.",
        actionLabel: "Archive Period",
        icon: "🗂️",
        secondaryAction: { sheet: "PTO_Archive_Summary", label: "Open Archive" }
    }
];

// Step → Sheet mapping (clicking step card activates this sheet)
// Matches payroll-recorder 5-step structure
const STEP_SHEET_MAP = {
    0: "PTO_Homepage",           // Configuration → PTO_Homepage
    1: "PTO_Data_Clean",         // Upload & Validate → PTO_Data_Clean
    2: "PTO_Review",             // Expense Review → PTO_Review
    3: "PTO_JE_Draft",           // Journal Entry Prep → PTO_JE_Draft
    4: "PTO_Archive_Summary"     // Archive & Clear → PTO_Archive_Summary
};

// Reverse map: sheet name → step ID (for tab-to-panel sync)
const SHEET_TO_STEP_MAP = {
    "PTO_Homepage": 0,           // Homepage → Configuration
    "PTO_Data_Clean": 1,         // PTO_Data_Clean → Upload & Validate
    "PTO_Data_Raw": 1,           // Raw data also maps to step 1
    "PTO_Review": 2,             // Expense Review → Expense Review step
    "PTO_JE_Draft": 3,           // PTO_JE_Draft → Journal Entry Prep
    "PTO_Archive_Summary": 4,    // Archive → Archive & Clear step
    "SS_PF_Config": 0,           // Config sheet → Configuration
    "SS_Employee_Roster": 1      // Roster → Upload & Validate (for coverage check)
};

const WORKBOOK_SHEETS = [
    {
        name: "PTO_Instructions",
        description: "Overview of the PTO workflow",
        position: "beginning",
        onCreate: async (sheet, context) => {
            // Updated to match payroll-recorder 5-step structure
            const rows = [
                ["Prairie Forge PTO Accrual", `Version ${MODULE_VERSION}`],
                ["Step 0 — Configuration", "Auto-loaded from system. Review period settings."],
                ["Step 1 — Upload & Validate", "Upload PTO report, normalize headers, run validation."],
                ["Step 2 — PTO Accrual Review", "Review metrics with executive-ready summary."],
                ["Step 3 — Journal Entry Prep", "Generate accounting-ready journal entry."],
                ["Step 4 — Archive & Clear", "Archive period and reset for next cycle."]
            ];
            const target = sheet.getRangeByIndexes(0, 0, rows.length, 2);
            target.values = rows;
            target.format.autofitColumns();

            const header = sheet.getRange("A1:B1");
            header.merge();
            header.format.font.bold = true;
            header.format.font.size = 18;
            header.format.font.color = "#111827";

            const info = sheet.getRange("A2:B8");
            info.format.fill.color = "#f1f5f9";
            await context.sync();
        }
    },
    // PTO_Config removed - consolidated into SS_PF_Config
    { name: "SS_PF_Config", description: "Prairie Forge shared configuration (all modules)" },
    { name: "PTO_Rates", description: "Accrual rate definitions" },
    { name: "PTO_Data_Clean", description: "Normalized PTO data (from upload)" },
    { name: "PTO_Review", description: "PTO review with liability calculations" },
    { name: "PTO_JE_Draft", description: "Journal entry prep" },
    { name: "PTO_Archive_Summary", description: "Archive register" },
    { name: "SS_Employee_Roster", description: "Centralized employee roster" }
];

const TABLE_DEFINITIONS = [
    // PTO_Config table removed - all config consolidated into SS_PF_Config
    {
        sheetName: "PTO_Rates",
        tableName: "PTORates",
        headers: ["Tier", "Description", "Hours_Per_Period", "Max_Carryover", "Carryover_Reset"],
        sampleRows: [
            ["Standard", "0-4 years tenure", 6.67, 80, "Dec 31"],
            ["Senior", "5-9 years tenure", 10, 120, "Dec 31"],
            ["Executive", "10+ years", 13.33, 160, "Dec 31"]
        ]
    },
    {
        sheetName: "PTO_Archive",
        tableName: "PTOArchiveLog",
        headers: ["Timestamp", "Action", "Notes"],
        sampleRows: []
    }
];

const SAMPLE_PTO_ACTIVITY = [
    { employeeId: "AC1001", employeeName: "Kelly Mendez", actionDate: "2025-01-15", actionType: "Accrual", hours: 6.67, notes: "Monthly accrual", source: "Payroll" },
    { employeeId: "AC1001", employeeName: "Kelly Mendez", actionDate: "2025-01-22", actionType: "Usage", hours: 8, notes: "Vacation day", source: "HRIS" },
    { employeeId: "AC2044", employeeName: "Justin Reid", actionDate: "2025-01-15", actionType: "Accrual", hours: 10, notes: "Senior tier", source: "Payroll" },
    { employeeId: "AC2044", employeeName: "Justin Reid", actionDate: "2025-01-28", actionType: "Usage", hours: 4, notes: "Doctor appointment", source: "HRIS" },
    { employeeId: "AC3020", employeeName: "Amelia Yates", actionDate: "2025-01-15", actionType: "Accrual", hours: 13.33, notes: "Executive tier", source: "Payroll" },
    { employeeId: "AC3020", employeeName: "Amelia Yates", actionDate: "2025-01-18", actionType: "Carryover", hours: 20, notes: "Year-end carryover", source: "Finance" }
];

// Sample data will be generated dynamically from PTO_Data_Clean via syncPtoAnalysis
const SAMPLE_ANALYSIS_DATA = [];

const SAMPLE_EXPENSE_REVIEW = [
    { department: "Operations", currentPeriod: 5400, priorPeriod: 5200, variance: 200, comment: "Increased carryover true-up" },
    { department: "Sales", currentPeriod: 2300, priorPeriod: 2600, variance: -300, comment: "Usage down vs. forecast" },
    { department: "Executive", currentPeriod: 4100, priorPeriod: 4100, variance: 0, comment: "Steady quarter" }
];

/**
 * Parse currency string like "$1,234.56" or "1,234.56" to number
 * Handles dollar signs, commas, and various formats
 */
function parseCurrency(value) {
    if (typeof value === "number") return value;
    if (!value) return 0;
    const cleaned = String(value).replace(/[$,]/g, "").trim();
    const num = parseFloat(cleaned);
    return isNaN(num) ? 0 : num;
}

const SAMPLE_JOURNAL_LINES = [
    { account: "2100.100", description: "PTO Accrual Expense", debit: 11800, credit: 0, reference: "PTO-EXP-2025-01" },
    { account: "2190.250", description: "PTO Accrual Liability", debit: 0, credit: 11800, reference: "PTO-EXP-2025-01" }
];

const stepStatuses = WORKFLOW_STEPS.reduce((acc, step) => {
    acc[step.id] = "pending";
    return acc;
}, {});

const appState = {
    activeView: "home",
    activeStepId: null,
    focusedIndex: 0,
    stepStatuses
};

const configState = {
    loaded: false,
    steps: {},
    permanents: {},
    completes: {},
    values: {},
    overrides: {
        accountingPeriod: false,
        journalId: false
    }
};

let rootEl = null;
let loadingEl = null;
let pendingScrollIndex = null;
const pendingConfigWrites = new Map();
const headcountState = {
    skipAnalysis: false,
    loading: false,
    hasAnalyzed: false,
    lastError: null,
    // Structured comparison data (matches payroll-recorder format)
    rosterCount: 0,           // Active employees in SS_Employee_Roster
    ptoCount: 0,              // Unique employees in PTO_Data
    missingFromPto: [],       // In roster but not in PTO: [{name, department}]
    extraInPto: [],           // In PTO but not in roster: [{name}]
    // Legacy fields for backward compatibility
    roster: {
        rosterCount: null,
        payrollCount: null,
        difference: null,
        mismatches: []
    }
};
const journalState = {
    debitTotal: null,
    creditTotal: null,
    difference: null,
    lineAmountSum: null,        // Sum of all line amounts (should be 0)
    analysisChangeTotal: null,  // Total Change from PTO_Analysis
    jeChangeTotal: null,        // Total Change captured in JE (expense lines only)
    loading: false,
    lastError: null,
    // Validation results with details
    validationRun: false,
    issues: []                  // [{check: "name", passed: false, detail: "explanation"}]
};
const dataQualityState = {
    hasRun: false,
    loading: false,
    acknowledged: false,       // User acknowledged issues and wants to proceed
    // Quality check results
    balanceIssues: [],         // [{name, issue, rowIndex}] - Negative balance or used more than available
    zeroBalances: [],          // [{name, rowIndex}]
    accrualOutliers: [],       // [{name, accrualRate, rowIndex}] - rates > 8 hrs/period
    // Summary counts
    totalIssues: 0,
    totalEmployees: 0,
    // UI state for expandable sections
    expandedSections: new Set() // "balanceIssues" | "zeroBalances" | "accrualOutliers"
};

const analysisState = {
    cleanDataReady: false,
    employeeCount: 0,
    lastRun: null,
    loading: false,
    lastError: null,
    // Data quality tracking
    missingPayRates: [],      // [{name: "John Doe", rowIndex: 2}, ...]
    missingDepartments: [],   // [{name: "Jane Smith", rowIndex: 3}, ...]
    ignoredMissingPayRates: new Set(), // Names user chose to ignore
    // Data completeness check (PTO_Data_Clean vs PTO_Analysis sums)
    completenessCheck: {
        accrualRate: null,    // { match: true/false, ptoData: number, ptoAnalysis: number }
        carryOver: null,
        ytdAccrued: null,
        ytdUsed: null,
        balance: null
    }
};

// =============================================================================
// PTO REVIEW STATE - Step 2 variance table
// =============================================================================
const ptoReviewState = {
    loaded: false,
    loading: false,
    lastRun: null,
    // Executive summary values
    totalCurrentLiability: 0,
    totalPriorLiability: 0,
    netChange: 0,
    employeeCount: 0,
    // Reconciliation data (Report → Calculated → JE)
    reconciliation: {
        reportLiabilityTotal: 0,      // Sum of Report_Liability from PrismHR
        calcLiabilityTotal: 0,        // Sum of Calc_Liability (includes negatives)
        negativeBalanceTotal: 0,      // Sum of negative Calc_Liability values
        negativeBalanceCount: 0,      // Count of employees with negative vested balance
        positiveBalanceCount: 0,      // Count with positive vested balance
        zeroBalanceCount: 0,          // Count with zero vested balance
        missingRateCount: 0           // Employees actually missing pay rate
    },
    // Employee coverage (Roster vs PTO Report)
    coverage: {
        rosterCount: 0,
        ptoReportCount: 0,
        inBothCount: 0,
        inPtoOnlyCount: 0,
        inPtoOnlyNames: [],
        inPtoOnlyLiability: 0,
        inRosterOnlyCount: 0,
        inRosterOnlyNames: []
    },
    // Review table data (written to PTO_Review sheet)
    reviewData: [],     // [{employeeName, department, payRate, vestedBalance, liabilityAmount, calculatedLiability, priorLiability, change, flags}]
    // Flags: NEW, MISSING_RATE, LARGE_MOVE, NEG_BALANCE, RATE_VARIANCE, LIABILITY_VARIANCE
    flagThresholds: {
        largeMove: 500  // Configurable threshold for large move flag
    }
};

// =============================================================================
// INSTALLATION STATE (loaded from SS_PF_Config - written by bootstrap)
// Bootstrap already fetches from ada_addin_installations and writes to SS_PF_Config
// =============================================================================

const installationState = {
    loaded: false,
    loading: false,
    error: null,
    // From SS_PF_Config (originally from ada_addin_installations via bootstrap)
    company_id: null,
    ss_company_name: null,
    pto_payroll_provider: null,
    ss_accounting_software: null,
    // Validation status
    isValid: false,
    validationErrors: []
};

/**
 * Load installation configuration from SS_PF_Config
 * Bootstrap has already synced values from ada_addin_installations
 * This mirrors how payroll-recorder loads config
 */
async function loadInstallationConfig() {
    console.log("[PTO] Loading installation configuration from SS_PF_Config...");
    installationState.loading = true;
    installationState.error = null;
    installationState.validationErrors = [];
    
    try {
        if (!hasExcelRuntime()) {
            throw new Error("Excel runtime not available.");
        }
        
        // Read config values directly from SS_PF_Config table
        // Bootstrap writes: SS_Company_ID, SS_Company_Name, PTO_Payroll_Provider, SS_Accounting_Software
        // loadConfigFromTable takes TABLE names, returns all fields as key-value object
        const configValues = await loadConfigFromTable(["SS_PF_Config"]);
        
        console.log("[PTO] Config values loaded:", configValues);
        
        // Store values in state
        installationState.company_id = configValues?.SS_Company_ID || null;
        installationState.ss_company_name = configValues?.SS_Company_Name || null;
        installationState.pto_payroll_provider = configValues?.PTO_Payroll_Provider || null;
        installationState.ss_accounting_software = configValues?.SS_Accounting_Software || null;
        
        // Validate required fields
        validateInstallation();
        
        installationState.loaded = true;
        installationState.loading = false;
        
        console.log("[PTO] Installation config loaded successfully:", {
            company_id: installationState.company_id,
            company_name: installationState.ss_company_name,
            provider: installationState.pto_payroll_provider,
            isValid: installationState.isValid
        });
        
    } catch (error) {
        console.error("[PTO] Failed to load installation:", error);
        installationState.error = error.message;
        installationState.loading = false;
        installationState.loaded = true;
        installationState.isValid = false;
        installationState.validationErrors.push(error.message);
    }
}

/**
 * Validate installation configuration
 * Fail fast if company_id is missing
 */
function validateInstallation() {
    installationState.validationErrors = [];
    
    // Check company_id
    if (!installationState.company_id) {
        installationState.validationErrors.push(
            "Company ID is not configured. Please open Module Selector to sync your configuration."
        );
    }
    
    // Provider validation removed - module now accepts any provider
    // Header normalization is handled by ada_payroll_dimensions
    
    installationState.isValid = installationState.validationErrors.length === 0;
    
    if (!installationState.isValid) {
        console.error("[PTO] Installation validation failed:", installationState.validationErrors);
    }
}

/**
 * Render a minimal banner for error screens
 */
function renderErrorBanner() {
    return `
        <div class="pf-root">
            <div class="pf-brand-float" aria-hidden="true">
                <span class="pf-brand-wave"></span>
            </div>
            <header class="pf-banner">
                <div class="pf-nav-bar">
                    <a href="../module-selector/index.html" class="pf-nav-btn pf-nav-btn--icon pf-clickable" title="Return to Modules">
                        ${MODULES_ICON_SVG}
                        <span class="sr-only">Return to Modules</span>
                    </a>
                </div>
            </header>
    `;
}

/**
 * Render blocking error screen when installation is invalid
 */
function renderInstallationError() {
    const errors = installationState.validationErrors || [installationState.error || "Unknown error"];
    
    // Check if this is a "not found" error vs a validation error
    const isSetupRequired = errors.some(e => 
        e.includes("not found") || e.includes("not configured") || 
        e.includes("initial setup") || e.includes("Module Selector") ||
        e.includes("sync")
    );
    
    return `
        ${renderErrorBanner()}
            <section class="pf-hero" id="pf-error-hero">
                <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)}</p>
                <h2 class="pf-hero-title" style="color: #ef4444;">${isSetupRequired ? "Setup Required" : "Configuration Error"}</h2>
                <p class="pf-hero-copy">${isSetupRequired ? "This workbook needs to be connected to your organization." : "Unable to load module configuration."}</p>
            </section>
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail" style="border-left: 4px solid ${isSetupRequired ? '#f59e0b' : '#ef4444'};">
                    <div class="pf-config-head">
                        <h3>${isSetupRequired ? '🔧 Initial Setup Needed' : '⚠️ Module Blocked'}</h3>
                        <p class="pf-config-subtext">${isSetupRequired ? 'Complete these steps to get started:' : 'The following issues must be resolved before using this module:'}</p>
                    </div>
                    ${isSetupRequired ? `
                    <div style="margin: 16px 0; padding: 16px; background: rgba(245, 158, 11, 0.1); border: 1px solid rgba(245, 158, 11, 0.3); border-radius: 8px;">
                        <ol style="margin: 0; padding-left: 20px; color: rgba(255,255,255,0.9); line-height: 1.8;">
                            <li><strong>Go to Module Selector</strong> — Click "Return to Modules" below</li>
                            <li><strong>Connect Your Organization</strong> — The Module Selector will sync your installation</li>
                            <li><strong>Return to PTO Accrual</strong> — Once connected, this module will load</li>
                        </ol>
                    </div>
                    ` : `
                    <ul style="margin: 16px 0; padding-left: 24px; color: rgba(255,255,255,0.8);">
                        ${errors.map(err => `<li style="margin: 8px 0;">${escapeHtml(err)}</li>`).join("")}
                    </ul>
                    `}
                    <div style="margin-top: 16px; padding: 12px; background: rgba(255,255,255,0.05); border-radius: 8px; font-size: 13px; color: rgba(255,255,255,0.6);">
                        <strong>Technical details:</strong><br>
                        ${errors.map(e => `• ${escapeHtml(e)}`).join('<br>')}
                    </div>
                </article>
                <div class="pf-pill-row" style="margin-top: 16px;">
                    <a href="../module-selector/index.html" class="pf-pill-btn" style="background: linear-gradient(145deg, #f59e0b, #d97706);">Return to Modules</a>
                    <button type="button" class="pf-pill-btn pf-pill-btn--secondary" id="retry-config-btn">Retry</button>
                </div>
            </section>
            <footer class="pf-brand-footer">
                <div class="pf-brand-text">
                    <div class="pf-brand-label">prairie.forge</div>
                    <div class="pf-brand-meta"> Prairie Forge LLC, 2025. All rights reserved. Version ${MODULE_VERSION}</div>
                </div>
            </footer>
        </div>
    `;
}

/**
 * Ensure SS_PF_Config sheet and table exist with proper structure
 * Creates the config sheet and table if they don't exist
 * This matches payroll-recorder's ensureConfigSheet() exactly
 */
async function ensureConfigSheet() {
    if (!hasExcelRuntime()) return;
    
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name");
            await context.sync();
            
            // Check if SS_PF_Config sheet exists
            let configSheet = worksheets.getItemOrNullObject("SS_PF_Config");
            configSheet.load("isNullObject");
            await context.sync();
            
            if (configSheet.isNullObject) {
                console.log("[PTO] Creating SS_PF_Config sheet...");
                configSheet = worksheets.add("SS_PF_Config");
                
                // Add headers
                const headers = ["Category", "Field", "Value", "Permanent"];
                const headerRange = configSheet.getRange("A1:D1");
                headerRange.values = [headers];
                formatSheetHeaders(headerRange);
                
                // Add default module-prefix rows
                const defaultData = [
                    ["module-prefix", "PR_", "payroll-recorder", "Y"],
                    ["module-prefix", "PTO_", "pto-accrual", "Y"],
                    ["module-prefix", "SS_", "system", "Y"],
                    ["Run Settings", "SS_Company_Name", "", "Y"],
                    ["Run Settings", "SS_Company_ID", "", "Y"]
                ];
                const dataRange = configSheet.getRange(`A2:D${1 + defaultData.length}`);
                dataRange.values = defaultData;
                
                await context.sync();
                
                // Create the table
                const tableRange = configSheet.getRange(`A1:D${1 + defaultData.length}`);
                const table = configSheet.tables.add(tableRange, true);
                table.name = "SS_PF_Config";
                table.style = "TableStyleMedium2";
                
                // Auto-fit columns
                tableRange.format.autofitColumns();
                
                await context.sync();
                console.log("[PTO] SS_PF_Config sheet and table created");
            } else {
                // Sheet exists - check if table exists
                const tables = context.workbook.tables;
                tables.load("items/name");
                await context.sync();
                
                const hasConfigTable = tables.items.some(t => t.name === "SS_PF_Config");
                
                if (!hasConfigTable) {
                    console.log("[PTO] SS_PF_Config sheet exists but no table - creating table...");
                    
                    // Get used range to determine table extent
                    const usedRange = configSheet.getUsedRangeOrNullObject();
                    usedRange.load("address,rowCount");
                    await context.sync();
                    
                    if (!usedRange.isNullObject && usedRange.rowCount > 0) {
                        const table = configSheet.tables.add(usedRange, true);
                        table.name = "SS_PF_Config";
                        table.style = "TableStyleMedium2";
                        await context.sync();
                        console.log("[PTO] SS_PF_Config table created from existing data");
                    } else {
                        // Empty sheet - add headers and create table
                        const headers = ["Category", "Field", "Value", "Permanent"];
                        const headerRange = configSheet.getRange("A1:D1");
                        headerRange.values = [headers];
                        headerRange.format.font.bold = true;
                        
                        const defaultData = [
                            ["module-prefix", "PR_", "payroll-recorder", "Y"],
                            ["module-prefix", "PTO_", "pto-accrual", "Y"],
                            ["module-prefix", "SS_", "system", "Y"],
                            ["Run Settings", "SS_Company_Name", "", "Y"],
                            ["Run Settings", "SS_Company_ID", "", "Y"]
                        ];
                        const dataRange = configSheet.getRange(`A2:D${1 + defaultData.length}`);
                        dataRange.values = defaultData;
                        
                        await context.sync();
                        
                        const tableRange = configSheet.getRange(`A1:D${1 + defaultData.length}`);
                        const table = configSheet.tables.add(tableRange, true);
                        table.name = "SS_PF_Config";
                        table.style = "TableStyleMedium2";
                        tableRange.format.autofitColumns();
                        
                        await context.sync();
                        console.log("[PTO] SS_PF_Config table created with default data");
                    }
                }
            }
        });
    } catch (error) {
        console.error("[PTO] Error ensuring config sheet:", error);
    }
}

async function init() {
    try {
        rootEl = document.getElementById("app");
        loadingEl = document.getElementById("loading");
        
        // CRITICAL: Ensure SS_PF_Config exists FIRST (matches payroll-recorder exactly)
        await ensureConfigSheet();
        
        // Load installation config from SS_PF_Config (non-blocking, like payroll-recorder)
        // Bootstrap populates SS_PF_Config when user visits Module Selector
        await loadInstallationConfig();
        
        // Log validation status but DON'T block - match payroll-recorder behavior
        if (!installationState.isValid) {
            console.warn("[PTO] Installation config incomplete:", installationState.validationErrors);
            console.warn("[PTO] Continuing anyway - some features may be limited");
        }
        
        await ensureTabVisibility();
        await loadStepConfig();
        // Load shared config for fallback values (Company Name, Default Reviewer, etc.)
        if (window.PrairieForge?.loadSharedConfig) {
            await window.PrairieForge.loadSharedConfig();
        }
        
        // Activate module homepage on load
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        
        // Set up worksheet change listener for bi-directional sync
        await setupWorksheetChangeListener();
        
        if (loadingEl) loadingEl.remove();
        if (rootEl) rootEl.hidden = false;
        renderApp();
    } catch (error) {
        console.error("[PTO] Module initialization failed:", error);
        throw error;
    }
}

/**
 * Bind retry button for installation error screen
 */
function bindRetryButton() {
    document.getElementById("retry-config-btn")?.addEventListener("click", async () => {
        showToast("Retrying configuration...", "info", 2000);
        await loadInstallationConfig();
        if (installationState.isValid) {
            // Re-initialize the full module
            await init();
        } else {
            // Re-render error screen
            if (rootEl) {
                rootEl.innerHTML = renderInstallationError();
            }
            bindRetryButton();
        }
    });
}

/**
 * Set up listener for worksheet activation changes
 * NOTE: Disabled - PTO uses one-way sync only (step click → tab opens)
 * Tab changes in Excel do NOT update the side panel (matches payroll-recorder behavior)
 */
async function setupWorksheetChangeListener() {
    // One-way sync only: clicking a step opens the corresponding tab
    // Clicking an Excel tab does NOT change the side panel step
    console.log("[PTO] Worksheet change listener disabled (one-way sync mode)");
}

async function ensureTabVisibility() {
    // Apply prefix-based tab visibility
    // Shows PTO_* tabs, hides PR_* and SS_* tabs
    try {
        await applyModuleTabVisibility(MODULE_KEY);
        console.log(`[PTO] Tab visibility applied for ${MODULE_KEY}`);
    } catch (error) {
        console.warn("[PTO] Could not apply tab visibility:", error);
    }
}

async function loadStepConfig() {
    if (!hasExcelRuntime()) {
        configState.loaded = true;
        return;
    }
    try {
        // Load from module-specific config table (backwards compatibility)
        const moduleValues = await loadConfigFromTable(CONFIG_TABLES);
        
        // Load from shared config (SS_PF_Config) - takes precedence
        let sharedValues = {};
        if (window.PrairieForge?.loadSharedConfig) {
            await window.PrairieForge.loadSharedConfig();
            // Convert cached Map to object
            if (window.PrairieForge._sharedConfigCache) {
                window.PrairieForge._sharedConfigCache.forEach((value, key) => {
                    sharedValues[key] = value;
                });
            }
        }
        
        // Merge: module values first, then shared values override
        // Also map new naming convention to old field names
        const values = { ...moduleValues };
        
        // Map shared config fields to PTO field names (new + legacy names)
        const fieldMappings = {
            "SS_Default_Reviewer": PTO_CONFIG_FIELDS.reviewerName,
            "Default_Reviewer": PTO_CONFIG_FIELDS.reviewerName,
            "PTO_Reviewer": PTO_CONFIG_FIELDS.reviewerName,
            "SS_Company_Name": PTO_CONFIG_FIELDS.companyName,
            "Company_Name": PTO_CONFIG_FIELDS.companyName,
            "SS_Payroll_Provider": PTO_CONFIG_FIELDS.payrollProvider,
            "Payroll_Provider_Link": PTO_CONFIG_FIELDS.payrollProvider,
            "SS_Accounting_Software": PTO_CONFIG_FIELDS.accountingSoftware,
            "Accounting_Software_Link": PTO_CONFIG_FIELDS.accountingSoftware
        };
        
        // Apply shared values using mapped field names
        Object.entries(fieldMappings).forEach(([sharedField, ptoField]) => {
            if (sharedValues[sharedField] && !values[ptoField]) {
                values[ptoField] = sharedValues[sharedField];
            }
        });
        
        // Also apply any exact-match PTO_* fields from shared config
        Object.entries(sharedValues).forEach(([key, value]) => {
            if (key.startsWith("PTO_") && value) {
                values[key] = value;
            }
        });
        
        configState.permanents = await loadPermanentFlags();
        configState.values = values || {};
        configState.overrides.accountingPeriod = Boolean(values?.[PTO_CONFIG_FIELDS.accountingPeriod]);
        configState.overrides.journalId = Boolean(values?.[PTO_CONFIG_FIELDS.journalEntryId]);
        Object.entries(STEP_CONFIG_FIELDS).forEach(([stepId, fields]) => {
            configState.steps[stepId] = {
                notes: values[fields.note] ?? "",
                reviewer: values[fields.reviewer] ?? "",
                signOffDate: values[fields.signOff] ?? ""
            };
        });
        configState.completes = Object.entries(STEP_COMPLETE_FIELDS).reduce((acc, [stepId, field]) => {
            acc[stepId] = values[field] ?? "";
            return acc;
        }, {});
        configState.loaded = true;
    } catch (error) {
        console.warn("PTO: unable to load configuration fields", error);
        configState.loaded = true;
    }
}

async function loadPermanentFlags() {
    const permanents = {};
    if (!hasExcelRuntime()) return permanents;
    const fieldToStep = new Map();
    Object.entries(STEP_CONFIG_FIELDS).forEach(([stepId, fields]) => {
        if (fields.note) {
            fieldToStep.set(fields.note.trim(), Number(stepId));
        }
    });
    try {
        await Excel.run(async (context) => {
            const table = context.workbook.tables.getItemOrNullObject(CONFIG_TABLES[0]);
            await context.sync();
            if (table.isNullObject) return;
            const body = table.getDataBodyRange();
            const header = table.getHeaderRowRange();
            body.load("values");
            header.load("values");
            await context.sync();

            const headers = header.values[0] || [];
            const normalizedHeaders = headers.map((h) => String(h || "").trim().toLowerCase());
            const idx = {
                field: normalizedHeaders.findIndex((h) => h === "field" || h === "field name" || h === "setting"),
                permanent: normalizedHeaders.findIndex((h) => h === "permanent" || h === "persist")
            };
            if (idx.field === -1 || idx.permanent === -1) return;
            (body.values || []).forEach((row) => {
                const fieldName = String(row[idx.field] || "").trim();
                const stepId = fieldToStep.get(fieldName);
                if (stepId == null) return;
                const flag = parsePermanentFlag(row[idx.permanent]);
                permanents[stepId] = flag;
            });
        });
    } catch (error) {
        console.warn("PTO: unable to load permanent flags", error);
    }
    return permanents;
}

/**
 * Mount Quick Access modal to document.body to escape pf-root stacking context.
 * This ensures proper z-index layering over all page content.
 */
function mountQuickAccessModal() {
    // Remove existing modal if present
    const existing = document.getElementById("quick-access-modal");
    if (existing) existing.remove();
    
    const modal = document.createElement("div");
    modal.id = "quick-access-modal";
    modal.className = "pf-quick-modal hidden";
    modal.style.cssText = "position:fixed!important;top:0;left:0;right:0;bottom:0;z-index:2147483647;";
    modal.innerHTML = `
        <div class="pf-quick-modal-backdrop" data-close></div>
        <div class="pf-quick-modal-card">
            <div class="pf-quick-modal-header">
                <h3 class="pf-quick-modal-title">Quick Access</h3>
                <button id="quick-access-close" class="pf-quick-modal-close pf-clickable" type="button" aria-label="Close">
                    ${X_ICON_SVG}
                </button>
            </div>
            <div class="pf-quick-modal-items">
                ${installationState.pto_payroll_provider ? `
                <a id="nav-provider-link" class="pf-quick-modal-item pf-clickable" href="${escapeHtml(installationState.pto_payroll_provider)}" target="_blank" rel="noopener">
                    ${FILE_TEXT_ICON_SVG}
                    <span>PTO Provider Report</span>
                </a>
                ` : ''}
                <button id="nav-accounting-software" class="pf-quick-modal-item pf-clickable" type="button">
                    ${UPLOAD_ICON_SVG}
                    <span>Accounting Software</span>
                </button>
                <button id="nav-employee-roster" class="pf-quick-modal-item pf-clickable" type="button">
                    ${USERS_ICON_SVG}
                    <span>Employee Roster</span>
                </button>
                <button id="nav-chart-of-accounts" class="pf-quick-modal-item pf-clickable" type="button">
                    ${BOOK_ICON_SVG}
                    <span>Chart of Accounts</span>
                </button>
                <button id="nav-config" class="pf-quick-modal-item pf-clickable" type="button">
                    ${SETTINGS_ICON_SVG}
                    <span>Configuration</span>
                </button>
            </div>
        </div>
    `;
    document.body.appendChild(modal);
}

function renderApp() {
    if (!rootEl) return;
    const prevDisabled = appState.focusedIndex <= 0 ? "disabled" : "";
    const nextDisabled = appState.focusedIndex >= WORKFLOW_STEPS.length - 1 ? "disabled" : "";
    const isStepView = appState.activeView === "step" && appState.activeStepId != null;
    const isConfigView = appState.activeView === "config";
    const content = isConfigView
        ? renderConfigView()
        : isStepView
            ? renderStepView(appState.activeStepId)
            : `${renderHero()}${renderWorkflow()}`;
    rootEl.innerHTML = `
        <div class="pf-root">
            <div class="pf-brand-float" aria-hidden="true">
                <span class="pf-brand-wave"></span>
            </div>
            <header class="pf-banner">
                <div class="pf-nav-bar">
                    <button id="nav-prev" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Previous step" ${prevDisabled}>
                        ${ARROW_LEFT_SVG}
                        <span class="sr-only">Previous step</span>
                    </button>
                    <button id="nav-home" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Home">
                        ${HOME_ICON_SVG}
                        <span class="sr-only">Module Home</span>
                    </button>
                    <button id="nav-selector" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Module Selector">
                        ${MODULES_ICON_SVG}
                        <span class="sr-only">Module Selector</span>
                    </button>
                    <button id="nav-next" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" aria-label="Next step" ${nextDisabled}>
                        ${ARROW_RIGHT_SVG}
                        <span class="sr-only">Next step</span>
                    </button>
                    <div class="pf-nav-divider"></div>
                    <button id="nav-quick-toggle" class="pf-nav-btn pf-nav-btn--icon pf-clickable" type="button" title="Quick Access">
                        ${MENU_ICON_SVG}
                        <span class="sr-only">Quick Access Menu</span>
                    </button>
                </div>
            </header>
            ${content}
            <footer class="pf-brand-footer">
                <div class="pf-brand-text">
                    <div class="pf-brand-label">prairie.forge</div>
                    <div class="pf-brand-meta"> Prairie Forge LLC, 2025. All rights reserved. Version ${MODULE_VERSION}</div>
                </div>
            </footer>
        </div>
    `;
    // Determine if on home view
    const isHomeView = appState.activeView === "home" || (appState.activeView !== "step" && appState.activeView !== "config");
    
    // Mount info FAB with step-specific content (only on step/config views, not homepage)
    const infoFabElement = document.getElementById("pf-info-fab-pto");
    if (isHomeView) {
        // Remove info fab on homepage
        if (infoFabElement) infoFabElement.remove();
    } else if (window.PrairieForge?.mountInfoFab) {
        const infoConfig = getStepInfoConfig(appState.activeStepId);
        PrairieForge.mountInfoFab({ 
            title: infoConfig.title, 
            content: infoConfig.content, 
            buttonId: "pf-info-fab-pto" 
        });
    }
    
    // Mount quick access modal to body for proper z-index layering
    mountQuickAccessModal();
    
    bindInteractions();
    scrollFocusedIntoView();
    
    // Show/hide Ada FAB based on view
    if (isHomeView) {
        renderAdaFab();
    } else {
        removeAdaFab();
    }
}

/**
 * Get step-specific info panel configuration
 */
/**
 * Step-specific info panel configuration
 * Matches payroll-recorder 5-step structure (0-4)
 */
function getStepInfoConfig(stepId) {
    switch (stepId) {
        case 0:
            return {
                title: "Configuration",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Configuration is auto-loaded from your installation. Review period-specific settings before each PTO run.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Auto-Loaded (Read-Only)</h4>
                        <ul>
                            <li><strong>Company Name</strong> — From your installation</li>
                            <li><strong>Company ID</strong> — Used for GL mapping lookups</li>
                            <li><strong>Company ID</strong> — Required for GL mappings</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>📝 Period-Specific (Editable)</h4>
                        <ul>
                            <li><strong>Analysis Date</strong> — The period-end date (e.g., 11/30/2024)</li>
                            <li><strong>Accounting Period</strong> — Shows up in your JE description</li>
                            <li><strong>Journal Entry ID</strong> — Reference number for your accounting system</li>
                        </ul>
                    </div>
                `
            };
        case 1:
            return {
                title: "Upload & Validate PTO Data",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Upload your Obsidian PTO export, auto-normalize headers, create PTO_Data_Clean, and run advisory validation checks.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📥 Upload Process</h4>
                        <ol>
                            <li>Download your PTO report from Obsidian</li>
                            <li>Click Upload and select the file</li>
                            <li>Headers are auto-normalized using ada_payroll_dimensions</li>
                            <li>PTO_Data_Clean is created with standardized columns</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <h4>✅ Validation Checks (Advisory)</h4>
                        <ul>
                            <li><strong>Employee Coverage</strong> — Compare against SS_Employee_Roster</li>
                            <li><strong>Data Quality</strong> — Negative balances, outliers, date consistency</li>
                        </ul>
                        <p class="pf-info-note">Validation is advisory and does not block the workflow.</p>
                    </div>
                `
            };
        case 2:
            return {
                title: "PTO Accrual Review",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Review PTO accrual metrics with an executive-ready summary. Uses PTO_Data_Clean as the single source of truth.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📊 Key Measures</h4>
                        <ul>
                            <li><strong>Accrual Rate</strong> — Hours accrued per pay period</li>
                            <li><strong>Carry Over</strong> — Hours carried from prior year</li>
                            <li><strong>YTD Accrued / Used</strong> — Year-to-date totals</li>
                            <li><strong>Pay Period Accrued / Used</strong> — Current period activity</li>
                            <li><strong>Balance</strong> — Current available hours</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>📈 Analysis View</h4>
                        <p>Data is grouped by Employee Name and Plan Description for detailed review.</p>
                    </div>
                `
            };
        case 3:
            return {
                title: "Journal Entry Prep",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Generates accounting-ready journal entry output from your PTO analysis.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📝 JE Generation</h4>
                        <ul>
                            <li>Uses PTO_Data_Clean as the data source</li>
                            <li>GL accounts resolved via ada_customer_gl_mappings</li>
                            <li>Output ready for QuickBooks or your accounting system</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>✅ Validation</h4>
                        <ul>
                            <li><strong>Debits = Credits</strong> — Entry must balance</li>
                            <li><strong>All rows mapped</strong> — Every line has a GL account</li>
                        </ul>
                    </div>
                `
            };
        case 4:
            return {
                title: "Archive & Clear",
                content: `
                    <div class="pf-info-section">
                        <h4>🎯 What This Step Does</h4>
                        <p>Archive this period's results and reset working tabs for the next PTO cycle.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📦 What Gets Archived</h4>
                        <ul>
                            <li><strong>PTO_Archive_Summary</strong> — Rolling summary of last 5 periods</li>
                            <li><strong>Excel Workbook</strong> — Full snapshot saved externally</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>🧹 What Gets Cleared</h4>
                        <ul>
                            <li>PTO_Data_Clean</li>
                            <li>PTO_Review</li>
                            <li>PTO_JE_Draft</li>
                        </ul>
                    </div>
                    <div class="pf-info-section">
                        <h4>⚠️ Important</h4>
                        <p>Ensure you've exported any needed files before clearing. The external workbook is your permanent backup.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>💡 Tip</h4>
                        <p>Make sure your JE has been uploaded to your accounting system before archiving.</p>
                    </div>
                `
            };
        default:
            return {
                title: "PTO Accrual",
                content: `
                    <div class="pf-info-section">
                        <h4>👋 Welcome to PTO Accrual</h4>
                        <p>This module helps you calculate PTO liabilities and generate journal entries each period.</p>
                    </div>
                    <div class="pf-info-section">
                        <h4>📋 Workflow Overview</h4>
                        <ol style="margin: 8px 0; padding-left: 20px;">
                            <li>Review period configuration (auto-loaded)</li>
                            <li>Upload & validate PTO data</li>
                            <li>Review PTO accrual metrics</li>
                            <li>Generate journal entry</li>
                            <li>Archive & clear for next period</li>
                        </ol>
                    </div>
                    <div class="pf-info-section">
                        <p>Click a step card to get started, or tap the <strong>ⓘ</strong> button on any step for detailed guidance.</p>
                    </div>
                `
            };
    }
}

function renderHero() {
    return `
        <section class="pf-hero" id="pf-hero">
            <h2 class="pf-hero-title">PTO Accrual</h2>
            <p class="pf-hero-copy">${HERO_COPY}</p>
        </section>
    `;
}

function renderWorkflow() {
    return `
        <section class="pf-step-guide">
            <div class="pf-step-grid">
                ${WORKFLOW_STEPS.map((step, index) => renderStepCard(step, index)).join("")}
            </div>
        </section>
    `;
}

function renderStepCard(step, index) {
    const status = appState.stepStatuses[step.id] || "pending";
    const isActive =
        appState.activeView === "step" && appState.focusedIndex === index ? "pf-step-card--active" : "";
    const icon = getStepIconSvg(getStepType(step.id));
    return `
        <article class="pf-step-card pf-clickable ${isActive}" data-step-card data-step-index="${index}" data-step-id="${step.id}">
            <p class="pf-step-index">Step ${step.id}</p>
            <h3 class="pf-step-title">${icon ? `${icon}` : ""}${step.title}</h3>
        </article>
    `;
}

function renderArchiveStep(detail) {
    // Step 4 - matches payroll-recorder archive step exactly
    // Check completion of steps 0-3 (don't include step 4 itself)
    const completionItems = WORKFLOW_STEPS.filter((step) => step.id !== 4).map((step) => ({
        id: step.id,
        title: step.title,
        complete: isStepComplete(step.id)
    }));
    const allComplete = completionItems.every((item) => item.complete);
    const incompleteCount = completionItems.filter(i => !i.complete).length;
    
    // Debug: Log completion status
    console.log("[Archive Step] Completion check:", completionItems.map(i => `Step ${i.id}: ${i.complete}`).join(", "));
    console.log("[Archive Step] All complete:", allComplete, "Incomplete count:", incompleteCount);
    
    // Step completion cards - matches payroll-recorder style exactly
    const statusList = completionItems
        .map(
            (item) => `
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head pf-notes-header">
                    <span class="pf-action-toggle ${item.complete ? "is-active" : ""}" aria-pressed="${item.complete}">
                        ${CHECK_ICON_SVG}
                    </span>
                    <div>
                        <h3>${escapeHtml(item.title)}</h3>
                        <p class="pf-config-subtext">${item.complete ? "Complete" : "Not complete"}</p>
                    </div>
                </div>
            </article>
        `
        )
        .join("");
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <p class="pf-hero-hint"></p>
        </section>
        <section class="pf-step-guide">
            ${statusList}
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Archive & Reset</h3>
                    <p class="pf-config-subtext">Create an archive of this module's sheets and clear work tabs.</p>
                </div>
                ${!allComplete ? `<p class="pf-step-note pf-step-note--info">Complete all ${incompleteCount} remaining step(s) before archiving. Click the checkmark on each step's sign-off section.</p>` : `<p class="pf-step-note" style="color: #22c55e;">All steps complete. Ready to archive!</p>`}
                <div class="pf-pill-row pf-config-actions">
                    <button type="button" class="pf-pill-btn ${allComplete ? 'pf-cta-button' : ''}" id="archive-run-btn" ${allComplete ? "" : "disabled"} style="${allComplete ? '' : 'opacity: 0.5; cursor: not-allowed;'}">
                        ${allComplete ? 'Archive Now' : 'Archive (Complete Steps First)'}
                    </button>
                </div>
            </article>
        </section>
    `;
}

function renderConfigView() {
    if (!configState.loaded) {
        return `
            <section class="pf-step-guide">
                <article class="pf-step-card pf-step-detail">
                    <p class="pf-step-title">Loading configuration…</p>
                </article>
            </section>
        `;
    }
    
    // Period-specific fields (user can edit)
    const payrollDate = formatDateInput(getConfigValue(PTO_CONFIG_FIELDS.payrollDate));
    const accountingPeriod = formatDateInput(getConfigValue(PTO_CONFIG_FIELDS.accountingPeriod));
    const journalEntryId = getConfigValue(PTO_CONFIG_FIELDS.journalEntryId);
    const userName = getConfigValue(PTO_CONFIG_FIELDS.reviewerName);
    
    // Auto-loaded from installation (read-only)
    const companyName = installationState.ss_company_name || getConfigValue(PTO_CONFIG_FIELDS.companyName) || "";
    const companyId = installationState.company_id || "";
    const ptoProvider = installationState.pto_payroll_provider || "";
    const accountingSoftware = installationState.ss_accounting_software || getConfigValue(PTO_CONFIG_FIELDS.accountingSoftware) || "";
    
    // Step fields for notes/signoff
    const stepFields = getStepConfig(0);
    const notesPermanent = Boolean(configState.permanents[0]);
    const isStepComplete = Boolean(parseBooleanFlag(configState.completes[0]) || stepFields.signOffDate);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";

    return `
        <section class="pf-hero" id="pf-config-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step 0</p>
            <h2 class="pf-hero-title">Configuration Setup</h2>
            <p class="pf-hero-copy">Make quick adjustments before every PTO run.</p>
        </section>
        <section class="pf-step-guide">
            <!-- Period Data (User editable) - FIRST, matches payroll-recorder -->
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Period Data</h3>
                    <p class="pf-config-subtext">Fields in this section may change each period.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Your Name (Used for sign-offs)</span>
                        <input type="text" id="config-user-name" value="${escapeHtml(userName)}" placeholder="Full name">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Analysis Date</span>
                        <input type="date" id="config-payroll-date" value="${escapeHtml(payrollDate)}">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Period</span>
                        <input type="text" id="config-accounting-period" value="${escapeHtml(accountingPeriod)}" placeholder="Nov 2025">
                    </label>
                    <label class="pf-config-field">
                        <span>Journal Entry ID</span>
                        <input type="text" id="config-journal-id" value="${escapeHtml(journalEntryId)}" placeholder="PTO-AUTO-YYYY-MM-DD">
                    </label>
                </div>
            </article>
            
            <!-- Static Data - SECOND, matches payroll-recorder naming/structure -->
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Static Data</h3>
                    <p class="pf-config-subtext">Fields rarely change but should be reviewed.</p>
                </div>
                <div class="pf-config-grid">
                    <label class="pf-config-field">
                        <span>Company Name</span>
                        <input type="text" id="config-company-name" value="${escapeHtml(companyName)}" placeholder="Prairie Forge LLC">
                    </label>
                    <label class="pf-config-field">
                        <span>Company ID <span class="pf-field-hint">(from Prairie Forge CRM)</span></span>
                        <input type="text" id="config-company-id" value="${escapeHtml(companyId)}" placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx">
                    </label>
                    <label class="pf-config-field">
                        <span>PTO Provider / Report Location</span>
                        <input type="url" id="config-pto-provider" value="${escapeHtml(ptoProvider)}" placeholder="https://…">
                    </label>
                    <label class="pf-config-field">
                        <span>Accounting Software / Import Location</span>
                        <input type="url" id="config-accounting-link" value="${escapeHtml(accountingSoftware)}" placeholder="https://…">
                    </label>
                </div>
            </article>
            
            ${renderInlineNotes({
                textareaId: "config-notes",
                value: stepFields.notes || "",
                permanentId: "config-notes-lock",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "config-notes-save"
            })}
            ${renderSignoff({
                reviewerInputId: "config-reviewer",
                reviewerValue: stepReviewer,
                signoffInputId: "config-signoff-date",
                signoffValue: stepSignOff,
                isComplete: isStepComplete,
                saveButtonId: "config-signoff-save",
                completeButtonId: "config-signoff-toggle"
            })}
        </section>
    `;
}

/**
 * Render the PTO file upload dropzone HTML (matches payroll-recorder exactly)
 */
function renderPtoFileUploadZone() {
    const hasFile = ptoUploadState.file || ptoUploadState.headers.length > 0;
    const isLoading = ptoUploadState.loading;
    
    // Debug logging to diagnose upload state
    console.log("[PTO-Upload] renderPtoFileUploadZone state:", {
        hasFile,
        isLoading,
        file: ptoUploadState.file ? ptoUploadState.fileName : null,
        headersLength: ptoUploadState.headers.length,
        headers: ptoUploadState.headers.slice(0, 3) // First 3 for brevity
    });
    
    if (isLoading) {
        return `
            <div class="pf-upload-zone pf-upload-zone--analyzing">
                <div class="pf-upload-spinner"></div>
                <p class="pf-upload-status">Processing your PTO report…</p>
            </div>
        `;
    }
    
    if (hasFile) {
        // Match payroll-recorder: just show file info, no Map Columns button
        // Validation runs automatically after upload
        return `
            <div class="pf-upload-zone pf-upload-zone--ready">
                <div class="pf-upload-file-info">
                    <svg class="pf-upload-file-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
                        <polyline points="14 2 14 8 20 8"/>
                        <line x1="16" y1="13" x2="8" y2="13"/>
                        <line x1="16" y1="17" x2="8" y2="17"/>
                    </svg>
                    <div class="pf-upload-file-details">
                        <span class="pf-upload-filename">${escapeHtml(ptoUploadState.fileName)}</span>
                        <span class="pf-upload-meta">${ptoUploadState.headers.length} columns • ${(ptoUploadState.rowCount || 0).toLocaleString()} rows</span>
                    </div>
                    <button type="button" class="pf-upload-clear" id="pto-upload-clear-btn" title="Remove file">×</button>
                </div>
            </div>
        `;
    }
    
    return `
        <div class="pf-upload-zone" id="pto-upload-dropzone">
            <input type="file" id="pto-upload-file-input" accept=".csv,.xlsx,.xls" hidden>
            <div class="pf-upload-content">
                <svg class="pf-upload-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                    <polyline points="17 8 12 3 7 8"/>
                    <line x1="12" y1="3" x2="12" y2="15"/>
                </svg>
                <p class="pf-upload-text">Drop your PTO file here</p>
                <p class="pf-upload-hint">or <button type="button" class="pf-upload-browse" id="pto-upload-browse-btn">browse</button> to upload</p>
                <p class="pf-upload-formats">Supports CSV, XLSX, XLS</p>
            </div>
        </div>
    `;
}

/**
 * Render Employee Coverage validation card (reconciliation-style layout)
 * Shows the math: Roster − Missing + Extra = PTO File
 * Read-only comparison - discrepancies should prompt investigation, not auto-fix
 */
function renderPtoEmployeeCoverageCard() {
    const hasData = headcountState.hasAnalyzed;
    const isLoading = headcountState.loading;
    
    // Use structured state fields
    const rosterCount = headcountState.rosterCount || 0;
    const ptoCount = headcountState.ptoCount || 0;
    const missingFromPto = headcountState.missingFromPto || [];
    const extraInPto = headcountState.extraInPto || [];
    
    const allGood = hasData && missingFromPto.length === 0 && extraInPto.length === 0;
    
    // Status badge
    let statusBadge = "";
    let statusClass = "";
    
    if (isLoading) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status">⏳ Loading</span>`;
    } else if (headcountState.lastError) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--unavailable" role="status">⚠ Unavailable</span>`;
        statusClass = "pf-coverage-unavailable";
    } else if (!hasData) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--pending" role="status">○ Pending</span>`;
    } else if (allGood) {
        statusBadge = `<span class="pf-status-badge pf-status-badge--ok" role="status">✓ OK</span>`;
        statusClass = "pf-coverage-ok";
    } else {
        statusBadge = `<span class="pf-status-badge pf-status-badge--review" role="status">⚠ Review</span>`;
        statusClass = "pf-coverage-review";
    }
    
    // Build content
    let content = "";
    
    if (isLoading) {
        content = `<p class="pf-subsection-hint">Analyzing coverage...</p>`;
    } else if (headcountState.lastError) {
        content = `<p class="pf-subsection-hint pf-subsection-hint--warn">${escapeHtml(headcountState.lastError)}</p>`;
    } else if (!hasData) {
        content = `<p class="pf-subsection-hint">Upload a PTO file to run this check.</p>`;
    } else {
        // Reconciliation-style layout showing the math
        content = `<div class="pf-recon-container">`;
        
        // Top line: Roster headcount (starting point)
        content += `
            <div class="pf-recon-row pf-recon-row--header">
                <button type="button" class="pf-recon-label pf-clickable" data-pto-coverage-detail="summary">Roster Headcount</button>
                <span class="pf-recon-value">${rosterCount}</span>
            </div>
        `;
        
        // Adjustment rows (if any discrepancies)
        if (missingFromPto.length > 0 || extraInPto.length > 0) {
            content += `<div class="pf-recon-changes">`;
            
            // Missing from PTO (subtract from roster to get PTO count)
        if (missingFromPto.length > 0) {
            content += `
                    <div class="pf-recon-row pf-recon-row--subtract">
                        <button type="button" class="pf-recon-label pf-clickable" data-pto-coverage-detail="missing">
                            <span class="pf-recon-operator">−</span> Missing from PTO
                </button>
                        <span class="pf-recon-value pf-recon-value--subtract">(${missingFromPto.length})</span>
                    </div>
            `;
        }
        
            // Extra in PTO (add to get PTO count)
        if (extraInPto.length > 0) {
            content += `
                    <div class="pf-recon-row pf-recon-row--add">
                        <button type="button" class="pf-recon-label pf-clickable" data-pto-coverage-detail="extra">
                            <span class="pf-recon-operator">+</span> In PTO only
                </button>
                        <span class="pf-recon-value pf-recon-value--add">${extraInPto.length}</span>
                    </div>
                `;
            }
            
            content += `</div>`;
        }
        
        // Separator line
        content += `<div class="pf-recon-divider"></div>`;
        
        // Bottom line: PTO file headcount (result)
        const checkMark = allGood ? ' ✓' : '';
        content += `
            <div class="pf-recon-row pf-recon-row--footer">
                <button type="button" class="pf-recon-label pf-clickable" data-pto-coverage-detail="summary">PTO File Headcount</button>
                <span class="pf-recon-value">${ptoCount}${checkMark}</span>
            </div>
        `;
        
        content += `</div>`;
        
        // Hint text
        if (!allGood) {
            content += `<p class="pf-coverage-hint">Click any row to see employee details</p>`;
        } else {
            content += `
                <div class="pf-coverage-item pf-coverage-item--ok" style="margin-top: 8px;">
                    <span class="pf-coverage-item-icon">✓</span>
                    <span>All roster employees found in PTO data</span>
                </div>
            `;
        }
    }
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card pf-employee-coverage-card ${statusClass}" id="pto-coverage-card">
            <div class="pf-config-head" style="display: flex; align-items: center;">
                <div>
                    <h3>Employee Coverage ${statusBadge}</h3>
                    <p class="pf-config-subtext">Compare PTO employees against roster.</p>
                </div>
                <button type="button" class="pf-action-toggle pf-action-toggle--subtle pf-clickable" id="pto-check-coverage-btn" title="Run coverage check" style="margin-left: auto;">
                    ${REFRESH_ICON_SVG}
                </button>
            </div>
            ${content}
        </article>
    `;
}

/**
 * Show PTO coverage detail modal for a specific category
 */
function showPtoCoverageDetailModal(category) {
    const rosterCount = headcountState.rosterCount || 0;
    const ptoCount = headcountState.ptoCount || 0;
    const missingFromPto = headcountState.missingFromPto || [];
    const extraInPto = headcountState.extraInPto || [];
    
    let title = "";
    let content = "";
    
    switch (category) {
        case "summary":
            title = "Coverage Summary";
            content = renderPtoCoverageSummaryDetail(rosterCount, ptoCount, missingFromPto, extraInPto);
            break;
        case "missing":
            title = `${missingFromPto.length} Active Employees Not in PTO Data`;
            content = renderPtoEmployeeListDetail(missingFromPto, "These employees are marked active in the roster but were not found in the PTO data. Verify they should have PTO balances.");
            break;
        case "extra":
            title = `${extraInPto.length} In PTO Data Only`;
            content = renderPtoEmployeeListDetail(extraInPto, "These employees are in the PTO data but not found in the active roster. They may be terminated, on leave, or the roster needs updating.");
            break;
        default:
            return;
    }
    
    // Create and show modal
    showPtoCoverageModal(title, content);
}

/**
 * Render the summary detail with headcount bridge calculation
 */
function renderPtoCoverageSummaryDetail(rosterCount, ptoCount, missingFromPto, extraInPto) {
    const difference = ptoCount - rosterCount;
    const differenceText = difference === 0 ? "Match" : (difference > 0 ? `+${difference}` : `${difference}`);
    const differenceClass = difference === 0 ? "match" : "mismatch";
    
    return `
        <div class="pf-coverage-detail-summary">
            <h4>Headcount Bridge</h4>
            <div class="pf-headcount-bridge">
                <div class="pf-bridge-row pf-bridge-row--base">
                    <span class="pf-bridge-label">Active in Roster</span>
                    <span class="pf-bridge-value">${rosterCount}</span>
                </div>
                ${extraInPto.length > 0 ? `
                <div class="pf-bridge-row pf-bridge-row--add">
                    <span class="pf-bridge-label">+ In PTO data only (not in roster)</span>
                    <span class="pf-bridge-value">+${extraInPto.length}</span>
                </div>
                ` : ""}
                ${missingFromPto.length > 0 ? `
                <div class="pf-bridge-row pf-bridge-row--subtract">
                    <span class="pf-bridge-label">− Active but not in PTO data</span>
                    <span class="pf-bridge-value">−${missingFromPto.length}</span>
                </div>
                ` : ""}
                <div class="pf-bridge-row pf-bridge-row--total">
                    <span class="pf-bridge-label">In PTO Data</span>
                    <span class="pf-bridge-value">${ptoCount}</span>
                </div>
                <div class="pf-bridge-row pf-bridge-row--diff pf-bridge-row--${differenceClass}">
                    <span class="pf-bridge-label">Difference</span>
                    <span class="pf-bridge-value">${differenceText}</span>
                </div>
            </div>
            
            <h4 style="margin-top: 20px;">What to Review</h4>
            <ul class="pf-coverage-review-list">
                ${missingFromPto.length > 0 ? `<li><strong>${missingFromPto.length} not in PTO</strong> — Verify these active employees should have PTO balances</li>` : ""}
                ${extraInPto.length > 0 ? `<li><strong>${extraInPto.length} PTO only</strong> — Check if these are terminated/on leave, or update roster via Payroll Recorder</li>` : ""}
                ${missingFromPto.length === 0 && extraInPto.length === 0 ? `<li>✓ No issues found — roster and PTO data are in sync</li>` : ""}
            </ul>
        </div>
    `;
}

/**
 * Render a list of employees for PTO coverage modal
 */
function renderPtoEmployeeListDetail(employees, description) {
    if (!employees || employees.length === 0) {
        return `<p>${description}</p><p><em>No employees in this category.</em></p>`;
    }
    
    let html = `<p style="margin-bottom: 16px;">${description}</p>`;
    html += `<div class="pf-employee-list">`;
    
    employees.forEach((emp, index) => {
        const name = emp.name || emp;
        const dept = emp.department || null;
        html += `
            <div class="pf-employee-list-item">
                <span class="pf-employee-index">${index + 1}.</span>
                <span class="pf-employee-name">${escapeHtml(name)}</span>
                ${dept ? `<span class="pf-employee-dept">${escapeHtml(dept)}</span>` : ""}
            </div>
        `;
    });
    
    html += `</div>`;
    return html;
}

/**
 * Show PTO coverage detail modal
 */
function showPtoCoverageModal(title, content) {
    // Remove existing modal if present
    const existing = document.getElementById("pto-coverage-detail-modal");
    if (existing) existing.remove();
    
    const modal = document.createElement("div");
    modal.id = "pto-coverage-detail-modal";
    modal.className = "pf-coverage-modal";
    modal.innerHTML = `
        <div class="pf-coverage-modal-backdrop" data-close></div>
        <div class="pf-coverage-modal-card">
            <div class="pf-coverage-modal-header">
                <h3 class="pf-coverage-modal-title">${escapeHtml(title)}</h3>
                <button class="pf-coverage-modal-close pf-clickable" type="button" aria-label="Close" data-close>
                    ${X_ICON_SVG}
                </button>
            </div>
            <div class="pf-coverage-modal-body">
                ${content}
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
    
    // Bind close handlers
    modal.querySelectorAll("[data-close]").forEach(el => {
        el.addEventListener("click", () => modal.remove());
    });
    
    // Close on escape key
    const escHandler = (e) => {
        if (e.key === "Escape") {
            modal.remove();
            document.removeEventListener("keydown", escHandler);
        }
    };
    document.addEventListener("keydown", escHandler);
}

/**
 * Render an expandable issue section for data quality card
 */
function renderQualityIssueSection(sectionKey, label, items, renderItem) {
    if (!items?.length) return '';
    
    const isExpanded = dataQualityState.expandedSections.has(sectionKey);
    const chevron = isExpanded ? '▼' : '▶';
    
    const itemsHtml = isExpanded 
        ? `<div class="pf-quality-issue-details">
            ${items.map(renderItem).join('')}
           </div>`
        : '';
    
    return `
        <div class="pf-quality-issue-row">
            <button type="button" class="pf-quality-issue-toggle pf-clickable" data-quality-section="${sectionKey}">
                <span class="pf-quality-chevron">${chevron}</span>
                <span class="pf-quality-issue-label">${items.length} ${label}</span>
            </button>
            ${itemsHtml}
        </div>
    `;
}

/**
 * Render Data Quality validation card with expandable issue sections
 */
function renderPtoDataQualityCard() {
    const hasQualityData = dataQualityState.hasRun;
    const totalIssues = dataQualityState.totalIssues || 0;
    const qualityPassed = totalIssues === 0;
    
    const statusBadge = !hasQualityData
        ? `<span class="pf-status-badge pf-status-badge--pending" role="status">○ Pending</span>`
        : qualityPassed
            ? `<span class="pf-status-badge pf-status-badge--ok" role="status">✓ OK</span>`
            : `<span class="pf-status-badge pf-status-badge--review" role="status">⚠ Review</span>`;
    
    const statusMessage = hasQualityData
        ? (qualityPassed 
            ? `${dataQualityState.totalEmployees || 0} employees checked, no issues found`
            : `Found ${totalIssues} potential issue${totalIssues !== 1 ? 's' : ''} to review`)
        : 'Check for negative balances, outliers, and data anomalies.';
    
    let issueDetails = '';
    if (hasQualityData && !qualityPassed) {
        // Build expandable issue sections
        const sections = [];
        
        // Balance issues (negative balances, used more than available)
        sections.push(renderQualityIssueSection(
            'balanceIssues',
            'negative balance(s)',
            dataQualityState.balanceIssues,
            (item) => `<div class="pf-quality-issue-item pf-quality-issue-item--warn">
                <span class="pf-quality-issue-name">${escapeHtml(item.name)}</span>
                <span class="pf-quality-issue-value">${escapeHtml(item.issue)}</span>
            </div>`
        ));
        
        // Zero balances (informational)
        sections.push(renderQualityIssueSection(
            'zeroBalances',
            'zero balance(s)',
            dataQualityState.zeroBalances,
            (item) => `<div class="pf-quality-issue-item pf-quality-issue-item--info">
                <span class="pf-quality-issue-name">${escapeHtml(item.name)}</span>
                <span class="pf-quality-issue-value">0 hrs balance</span>
            </div>`
        ));
        
        // Accrual rate outliers
        sections.push(renderQualityIssueSection(
            'accrualOutliers',
            'accrual rate outlier(s)',
            dataQualityState.accrualOutliers,
            (item) => `<div class="pf-quality-issue-item pf-quality-issue-item--info">
                <span class="pf-quality-issue-name">${escapeHtml(item.name)}</span>
                <span class="pf-quality-issue-value">${item.accrualRate.toFixed(2)} hrs/period</span>
            </div>`
        ));
        
        const sectionsHtml = sections.filter(s => s).join('');
        if (sectionsHtml) {
            issueDetails = `
                <div class="pf-quality-issues-container">
                    ${sectionsHtml}
                </div>
            `;
        }
    }
    
    // Hint text when no data
    const hintText = !hasQualityData 
        ? `<p class="pf-metric-hint pf-metric-hint--info" style="margin-top: 12px; font-style: italic;">Upload a PTO file above to run this check.</p>`
        : '';
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" id="pto-quality-card">
            <div class="pf-config-head" style="display: flex; align-items: center;">
                <div>
                    <h3>Data Quality ${statusBadge}</h3>
                    <p class="pf-config-subtext">${statusMessage}</p>
                </div>
                <button type="button" class="pf-action-toggle pf-action-toggle--subtle pf-clickable" id="pto-check-quality-btn" title="Run quality check" style="margin-left: auto;">
                    ${REFRESH_ICON_SVG}
                </button>
            </div>
            ${issueDetails}
            ${hintText}
        </article>
    `;
}

/**
 * Step 1: Upload & Validate PTO Data
 * Mirrors payroll-recorder Step 1 UX pattern EXACTLY
 */
function renderUploadValidateStep(detail) {
    const stepFields = getStepConfig(1);
    const notesPermanent = Boolean(configState.permanents[1]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[1]) || stepSignOff);
    
    // Error display
    const errorHtml = ptoUploadState.error 
        ? `<p class="pf-upload-error">${escapeHtml(ptoUploadState.error)}</p>` 
        : "";
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">Upload & Validate PTO Data</h2>
            <p class="pf-hero-copy">Upload your PTO export, create the data matrix, and verify coverage.</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Import</h3>
                    <p class="pf-config-subtext">Drop your PTO file to auto-normalize headers.</p>
                </div>
                ${renderPtoFileUploadZone()}
                ${errorHtml}
            </article>
            
            <div class="pf-validation-section">
                <h3 class="pf-validation-header">Validation</h3>
                <p class="pf-validation-subtext">Advisory checks — review but not required to proceed.</p>
                ${renderPtoEmployeeCoverageCard()}
                ${renderPtoDataQualityCard()}
            </div>
            
            ${renderInlineNotes({
                textareaId: "step-notes-1",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-1",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-1"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-1",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-1",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-1",
                completeButtonId: "step-signoff-toggle-1"
            })}
        </section>
    `;
}

// LEGACY: renderImportStep kept for reference but no longer used
function renderImportStep(detail) {
    const stepFields = getStepConfig(1);
    const notesPermanent = Boolean(configState.permanents[1]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[1]) || stepSignOff);
    const providerLink = getConfigValue(PTO_CONFIG_FIELDS.payrollProvider);
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Payroll Report</h3>
                    <p class="pf-config-subtext">Access your payroll provider to download the latest PTO export.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        providerLink 
                            ? `<a href="${escapeHtml(providerLink)}" target="_blank" rel="noopener noreferrer" class="pf-action-toggle pf-clickable" title="Open payroll provider">${LINK_ICON_SVG}</a>`
                            : `<button type="button" class="pf-action-toggle pf-clickable" id="import-provider-btn" disabled title="Add provider link in Configuration">${LINK_ICON_SVG}</button>`,
                        "Provider"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="import-open-data-btn" title="Open PTO_Data_Clean sheet">${TABLE_ICON_SVG}</button>`,
                        "Data"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="import-clear-btn" title="Clear PTO_Data_Clean to start over">${TRASH_ICON_SVG}</button>`,
                        "Clear"
                    )}
                </div>
            </article>
            ${renderInlineNotes({
                textareaId: "step-notes-1",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-1",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-1"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-1",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-1",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-1",
                completeButtonId: "step-signoff-toggle-1"
            })}
        </section>
    `;
}

function renderStepView(stepId) {
    const detail = WORKFLOW_STEPS.find((step) => step.id === stepId);
    if (!detail) return "";
    
    // Step mapping - matches payroll-recorder 5-step structure
    switch (stepId) {
        case 0: return renderConfigView();
        case 1: return renderUploadValidateStep(detail);  // Replaces old Import + Headcount + Data Quality
        case 2: return renderAccrualReviewStep(detail);   // PTO Accrual Review
        case 3: return renderJournalStep(detail);         // Journal Entry Prep
        case 4: return renderArchiveStep(detail);         // Archive & Clear
    }
    
    // Generic step rendering (fallback, shouldn't be reached)
    const stepFields = getStepConfig(stepId);
    const notesPermanent = Boolean(configState.permanents[stepId]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[stepId]) || stepSignOff);
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            ${renderInlineNotes({
                textareaId: `step-notes-${stepId}`,
                value: stepFields?.notes || "",
                permanentId: `step-notes-lock-${stepId}`,
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: `step-notes-save-${stepId}`
            })}
            ${renderSignoff({
                reviewerInputId: `step-reviewer-${stepId}`,
                reviewerValue: stepReviewer,
                signoffInputId: `step-signoff-${stepId}`,
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: `step-signoff-save-${stepId}`,
                completeButtonId: `step-signoff-toggle-${stepId}`
            })}
        </section>
    `;
}

function bindInteractions() {
    document.getElementById("nav-home")?.addEventListener("click", async () => {
        // Activate the module homepage sheet
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        
        setState({ activeView: "home", activeStepId: null });
        document.getElementById("pf-hero")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("nav-selector")?.addEventListener("click", () => {
        window.location.href = SELECTOR_URL;
    });
    document.getElementById("nav-prev")?.addEventListener("click", () => moveFocus(-1));
    document.getElementById("nav-next")?.addEventListener("click", () => moveFocus(1));
    
    // Quick Access Modal (matches payroll-recorder)
    const quickToggle = document.getElementById("nav-quick-toggle");
    const quickModal = document.getElementById("quick-access-modal");
    const quickClose = document.getElementById("quick-access-close");
    const quickBackdrop = quickModal?.querySelector("[data-close]");
    
    const closeQuickModal = () => {
        quickModal?.classList.add("hidden");
        quickToggle?.classList.remove("is-active");
    };
    
    quickToggle?.addEventListener("click", (e) => {
        e.stopPropagation();
        quickModal?.classList.toggle("hidden");
        quickToggle.classList.toggle("is-active");
    });
    
    quickClose?.addEventListener("click", closeQuickModal);
    quickBackdrop?.addEventListener("click", closeQuickModal);
    
    // Quick Access - Accounting Software
    document.getElementById("nav-accounting-software")?.addEventListener("click", async () => {
        closeQuickModal();
        await openAccountingSoftware();
    });
    
    // Quick Access - Employee Roster (shows hidden SS_ sheet temporarily)
    document.getElementById("nav-employee-roster")?.addEventListener("click", async () => {
        closeQuickModal();
        await showAndActivateSheet("SS_Employee_Roster");
        showToast("Opened Employee Roster", "success", 2000);
    });
    
    // Quick Access - Chart of Accounts (shows hidden SS_ sheet temporarily)
    document.getElementById("nav-chart-of-accounts")?.addEventListener("click", async () => {
        closeQuickModal();
        await showAndActivateSheet("SS_Chart_of_Accounts");
        showToast("Opened Chart of Accounts", "success", 2000);
    });
    
    // Quick Access - Configuration (shows hidden SS_ sheet temporarily)
    document.getElementById("nav-config")?.addEventListener("click", async () => {
        closeQuickModal();
        await showAndActivateSheet("SS_PF_Config");
        showToast("Opened Configuration", "success", 2000);
    });
    
    document.querySelectorAll("[data-step-card]").forEach((card) => {
        const index = Number(card.getAttribute("data-step-index"));
        const stepId = Number(card.getAttribute("data-step-id"));
        card.addEventListener("click", () => focusStep(index, stepId));
    });
    if (appState.activeView === "config") {
        bindConfigView();
    } else if (appState.activeView === "step" && appState.activeStepId != null) {
        bindStepView(appState.activeStepId);
    }
}

function bindStepView(stepId) {
    // All step views use consistent ID patterns: step-notes-{stepId}, step-reviewer-{stepId}, etc.
    const notesInput = document.getElementById(`step-notes-${stepId}`);
    const reviewerInput = document.getElementById(`step-reviewer-${stepId}`);
    const signoffInput = document.getElementById(`step-signoff-${stepId}`);
    const backBtn = document.getElementById("step-back-btn");
    const lockBtn = document.getElementById(`step-notes-lock-${stepId}`);

    // Save button for notes
    const notesSaveBtn = document.getElementById(`step-notes-save-${stepId}`);
    notesSaveBtn?.addEventListener("click", async () => {
        const notes = notesInput?.value || "";
        await saveStepField(stepId, "notes", notes);
        updateSaveButtonState(notesSaveBtn, true);
    });

    // Save button for sign-off section (reviewer name)
    const signoffSaveBtn = document.getElementById(`step-signoff-save-${stepId}`);
    signoffSaveBtn?.addEventListener("click", async () => {
        const reviewer = reviewerInput?.value || "";
        await saveStepField(stepId, "reviewer", reviewer);
        updateSaveButtonState(signoffSaveBtn, true);
    });

    // Auto-wire save state tracking for all save buttons (marks as unsaved on input change)
    initSaveTracking();

    // Button IDs must match what's rendered in the step view
    const signoffButtonId = `step-signoff-toggle-${stepId}`;
    const signoffPrevId = `${signoffButtonId}-prev`;
    const signoffNextId = `${signoffButtonId}-next`;
    const signoffInputId = `step-signoff-${stepId}`;
    bindSignoffToggle(stepId, {
        buttonId: signoffButtonId,
        inputId: signoffInputId,
        canActivate: null, // No special validation needed - just sign off
        onComplete: getStepCompleteHandler(stepId)
    });
    bindSignoffNavButtons(signoffPrevId, signoffNextId);

    // Bind Ada copilot card for PTO step 2 (Accrual Review)
    if (stepId === 2) {
        const adaContainer = document.querySelector('[data-copilot="pto-copilot"]');
        if (adaContainer) {
            bindCopilotCard(adaContainer, {
                id: "pto-copilot",
                contextProvider: createExcelContextProvider({
                    dataClean: 'PTO_Data_Clean',
                    review: 'PTO_Review',
                    config: 'SS_PF_Config'
                }),
                onPrompt: callAdaApi
            });
        }
    }
    backBtn?.addEventListener("click", async () => {
        const homepageConfig = getHomepageConfig(MODULE_KEY);
        await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
        setState({ activeView: "home", activeStepId: null });
    });
    lockBtn?.addEventListener("click", async () => {
        const nextLocked = !lockBtn.classList.contains("is-locked");
        updateLockButtonVisual(lockBtn, nextLocked);
        await toggleNotePermanent(stepId, nextLocked);
    });
    // Step 4: Archive & Clear
    if (stepId === 4) {
        document.getElementById("archive-run-btn")?.addEventListener("click", () => {
            archiveAndReset();
        });
    }
    
    // Step 1: Upload & Validate PTO Data (drag-and-drop matching payroll-recorder)
    if (stepId === 1) {
        const dropzone = document.getElementById("pto-upload-dropzone");
        const fileInput = document.getElementById("pto-upload-file-input");
        const browseBtn = document.getElementById("pto-upload-browse-btn");
        const clearBtn = document.getElementById("pto-upload-clear-btn");
        
        // Debug logging to verify elements exist
        console.log("[PTO-Upload] bindStepView Step 1 - Element check:", {
            dropzone: !!dropzone,
            fileInput: !!fileInput,
            browseBtn: !!browseBtn,
            clearBtn: !!clearBtn
        });
        
        // Dropzone click handler - clicking anywhere on dropzone opens file picker
        // (matches payroll-recorder behavior)
        if (dropzone) {
            dropzone.addEventListener("click", () => {
                console.log("[PTO-Upload] Dropzone clicked, opening file picker");
                fileInput?.click();
        });
        
        // Drag and drop handlers
            dropzone.addEventListener("dragover", (e) => {
            e.preventDefault();
                e.stopPropagation();
            dropzone.classList.add("pf-upload-zone--dragover");
        });
            dropzone.addEventListener("dragleave", (e) => {
            e.preventDefault();
                e.stopPropagation();
            dropzone.classList.remove("pf-upload-zone--dragover");
        });
            dropzone.addEventListener("drop", (e) => {
            e.preventDefault();
                e.stopPropagation();
            dropzone.classList.remove("pf-upload-zone--dragover");
            const file = e.dataTransfer?.files?.[0];
                if (file) handlePtoFileUpload(file);
            });
        }
        
        // Browse button opens file picker (with stopPropagation to prevent dropzone click)
        browseBtn?.addEventListener("click", (e) => {
            e.stopPropagation();
            console.log("[PTO-Upload] Browse button clicked");
            fileInput?.click();
        });
        
        // File input change handler
        fileInput?.addEventListener("change", (e) => {
            const file = e.target.files?.[0];
            console.log("[PTO-Upload] File selected:", file?.name);
            if (file) handlePtoFileUpload(file);
        });
        
        // Clear uploaded file
        clearBtn?.addEventListener("click", () => {
            console.log("[PTO-Upload] Clear button clicked");
            ptoUploadState.file = null;
            ptoUploadState.fileName = "";
            ptoUploadState.headers = [];
            ptoUploadState.rowCount = 0;
            ptoUploadState.parsedData = null;
            ptoUploadState.error = null;
            renderApp();
        });
        
        // Individual validation check refresh buttons
        document.getElementById("pto-check-coverage-btn")?.addEventListener("click", async () => {
            showToast("Refreshing coverage check...", "info", 2000);
            const scrollY = window.scrollY;
            await refreshHeadcountAnalysis();
            renderApp();
            requestAnimationFrame(() => window.scrollTo(0, scrollY));
        });
        
        // Coverage detail pill clicks - open modal with category details
        document.querySelectorAll("[data-pto-coverage-detail]").forEach(pill => {
            pill.addEventListener("click", (e) => {
                e.preventDefault();
                const category = pill.dataset.ptoCoverageDetail;
                showPtoCoverageDetailModal(category);
            });
        });
        
        document.getElementById("pto-check-quality-btn")?.addEventListener("click", async () => {
            showToast("Refreshing quality check...", "info", 2000);
            const scrollY = window.scrollY;
            await runDataQualityCheck();
            renderApp();
            requestAnimationFrame(() => window.scrollTo(0, scrollY));
        });
        
        // Data Quality expandable sections
        document.querySelectorAll("[data-quality-section]").forEach(toggle => {
            toggle.addEventListener("click", (e) => {
                e.preventDefault();
                const section = toggle.dataset.qualitySection;
                const scrollY = window.scrollY;
                
                // Toggle expanded state
                if (dataQualityState.expandedSections.has(section)) {
                    dataQualityState.expandedSections.delete(section);
                } else {
                    dataQualityState.expandedSections.add(section);
                }
                
                renderApp();
                requestAnimationFrame(() => window.scrollTo(0, scrollY));
            });
        });
    }
    
    // Step 2: PTO Accrual Review
    if (stepId === 2) {
        document.getElementById("review-generate-btn")?.addEventListener("click", () => generatePtoReview());
        document.getElementById("review-open-btn")?.addEventListener("click", () => openSheet("PTO_Review"));
    }
    // Step 2: Headcount Review
    if (stepId === 2) {
        document.getElementById("headcount-skip-btn")?.addEventListener("click", () => {
            headcountState.skipAnalysis = !headcountState.skipAnalysis;
            const skipBtn = document.getElementById("headcount-skip-btn");
            skipBtn?.classList.toggle("is-active", headcountState.skipAnalysis);
            if (headcountState.skipAnalysis) {
                enforceHeadcountSkipNote();
            }
            updateHeadcountSignoffState();
        });
        document.getElementById("headcount-run-btn")?.addEventListener("click", () => refreshHeadcountAnalysis());
        document.getElementById("headcount-refresh-btn")?.addEventListener("click", () => refreshHeadcountAnalysis());
        bindHeadcountNotesGuard();
        if (headcountState.skipAnalysis) {
            enforceHeadcountSkipNote();
        }
        updateHeadcountSignoffState();
    }
    if (stepId === 3) {
        document.getElementById("quality-run-btn")?.addEventListener("click", () => runDataQualityCheck());
        document.getElementById("quality-refresh-btn")?.addEventListener("click", () => runDataQualityCheck());
        document.getElementById("quality-acknowledge-btn")?.addEventListener("click", () => acknowledgeQualityIssues());
    }
    if (stepId === 4) {
        // Both buttons trigger the full analysis with all checks
        document.getElementById("analysis-refresh-btn")?.addEventListener("click", () => runFullAnalysis());
        document.getElementById("analysis-run-btn")?.addEventListener("click", () => runFullAnalysis());
        
        // Missing pay rate card bindings
        document.getElementById("payrate-save-btn")?.addEventListener("click", handlePayRateSave);
        document.getElementById("payrate-ignore-btn")?.addEventListener("click", handlePayRateIgnore);
        document.getElementById("payrate-input")?.addEventListener("keydown", (e) => {
            if (e.key === "Enter") handlePayRateSave();
        });
    }
    // Step 3: Journal Entry Prep
    if (stepId === 3) {
        document.getElementById("je-create-btn")?.addEventListener("click", () => createJournalDraft());
        document.getElementById("je-run-btn")?.addEventListener("click", () => runJournalSummary());
        document.getElementById("je-export-btn")?.addEventListener("click", () => exportJournalDraft());
        document.getElementById("je-upload-btn")?.addEventListener("click", () => openAccountingSoftware());
    }
}

function bindConfigView() {
    // Initialize custom date picker
    initDatePicker("config-payroll-date", {
        onChange: (value) => {
            scheduleConfigWrite(PTO_CONFIG_FIELDS.payrollDate, value);
            if (!value) return;
            // Reset overrides so derived values follow the analysis date
            configState.overrides.accountingPeriod = false;
            configState.overrides.journalId = false;
            if (!configState.overrides.accountingPeriod) {
                const derivedPeriod = deriveAccountingPeriod(value);
                if (derivedPeriod) {
                    const periodInput = document.getElementById("config-accounting-period");
                    if (periodInput) periodInput.value = derivedPeriod;
                    scheduleConfigWrite(PTO_CONFIG_FIELDS.accountingPeriod, derivedPeriod);
                }
            }
            if (!configState.overrides.journalId) {
                const derivedJe = deriveJournalId(value);
                if (derivedJe) {
                    const jeInput = document.getElementById("config-journal-id");
                    if (jeInput) jeInput.value = derivedJe;
                    scheduleConfigWrite(PTO_CONFIG_FIELDS.journalEntryId, derivedJe);
                }
            }
        }
    });

    const periodInput = document.getElementById("config-accounting-period");
    periodInput?.addEventListener("change", (event) => {
        configState.overrides.accountingPeriod = Boolean(event.target.value);
        scheduleConfigWrite(PTO_CONFIG_FIELDS.accountingPeriod, event.target.value || "");
    });

    const journalInput = document.getElementById("config-journal-id");
    journalInput?.addEventListener("change", (event) => {
        configState.overrides.journalId = Boolean(event.target.value);
        scheduleConfigWrite(PTO_CONFIG_FIELDS.journalEntryId, event.target.value.trim());
    });

    document.getElementById("config-company-name")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.companyName, event.target.value.trim());
    });

    document.getElementById("config-payroll-provider")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.payrollProvider, event.target.value.trim());
    });

    document.getElementById("config-accounting-link")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.accountingSoftware, event.target.value.trim());
    });

    document.getElementById("config-user-name")?.addEventListener("change", (event) => {
        scheduleConfigWrite(PTO_CONFIG_FIELDS.reviewerName, event.target.value.trim());
    });

    const notesInput = document.getElementById("config-notes");
    notesInput?.addEventListener("input", (event) => {
        saveStepField(0, "notes", event.target.value);
    });

    const lockButton = document.getElementById("config-notes-lock");
    lockButton?.addEventListener("click", async () => {
        const nextLocked = !lockButton.classList.contains("is-locked");
        updateLockButtonVisual(lockButton, nextLocked);
        await toggleNotePermanent(0, nextLocked);
    });

    const notesSaveBtn = document.getElementById("config-notes-save");
    notesSaveBtn?.addEventListener("click", async () => {
        if (!notesInput) return;
        await saveStepField(0, "notes", notesInput.value);
        updateSaveButtonState(notesSaveBtn, true);
    });

    const reviewerInput = document.getElementById("config-reviewer");
    reviewerInput?.addEventListener("change", (event) => {
        const value = event.target.value.trim();
        saveStepField(0, "reviewer", value);
        const signoffInput = document.getElementById("config-signoff-date");
        if (value && signoffInput && !signoffInput.value) {
            const today = todayIso();
            signoffInput.value = today;
            saveStepField(0, "signOffDate", today);
            saveCompletionFlag(0, true);
        }
    });

    document.getElementById("config-signoff-date")?.addEventListener("change", (event) => {
        saveStepField(0, "signOffDate", event.target.value || "");
    });
    const signoffSaveBtn = document.getElementById("config-signoff-save");
    signoffSaveBtn?.addEventListener("click", async () => {
        const reviewerValue = reviewerInput?.value?.trim() || "";
        const signoffValue = document.getElementById("config-signoff-date")?.value || "";
        await saveStepField(0, "reviewer", reviewerValue);
        await saveStepField(0, "signOffDate", signoffValue);
        updateSaveButtonState(signoffSaveBtn, true);
    });

    initSaveTracking();
    bindSignoffToggle(0, {
        buttonId: "config-signoff-toggle",
        inputId: "config-signoff-date",
        onComplete: () => {
            persistConfigBasics();
            advanceToNextStep(0);
            scrollPanelToTop();
        }
    });
    bindSignoffNavButtons("config-signoff-toggle-prev", "config-signoff-toggle-next");
}

function focusStep(index, stepId = null) {
    if (index < 0 || index >= WORKFLOW_STEPS.length) return;
    pendingScrollIndex = index;
    const resolvedStepId = stepId ?? WORKFLOW_STEPS[index].id;
    const nextView = resolvedStepId === 0 ? "config" : "step";
    setState({ focusedIndex: index, activeView: nextView, activeStepId: resolvedStepId });
    
    // Activate the corresponding sheet from STEP_SHEET_MAP
    const sheetName = STEP_SHEET_MAP[resolvedStepId];
    if (sheetName) {
        console.log("[NAV] Step→Sheet activation", {
            moduleKey: MODULE_KEY,
            stepIndex: index,
            stepId: resolvedStepId,
            targetSheetName: sheetName
        });
        showAndActivateSheet(sheetName)
            .then(() => {
                console.log("[NAV] Step→Sheet activation success", {
                    moduleKey: MODULE_KEY,
                    stepId: resolvedStepId,
                    targetSheetName: sheetName
                });
            })
            .catch((err) => {
                console.warn("[NAV] Step→Sheet activation failed", {
                    moduleKey: MODULE_KEY,
                    stepId: resolvedStepId,
                    targetSheetName: sheetName,
                    error: err?.message ?? String(err)
                });
            });
    }
    
    // Step-specific initialization
    if (resolvedStepId === 2 && !headcountState.hasAnalyzed) {
        refreshHeadcountAnalysis();
    }
}

/**
 * Get the completion handler for a step
 * Steps 0-3 advance to the next step when completed
 * Step 4 (Archive) is handled separately with archive flow
 */
function getStepCompleteHandler(stepId) {
    // Step 4 (Archive) is handled separately with archive flow
    if (stepId === 4) return null;
    
    // All other steps advance to the next step
    return () => advanceToNextStep(stepId);
}

/**
 * Advance to the next step after completing the current one
 */
function advanceToNextStep(currentStepId) {
    const currentIndex = WORKFLOW_STEPS.findIndex((step) => step.id === currentStepId);
    if (currentIndex === -1) return;
    
    const nextIndex = currentIndex + 1;
    if (nextIndex < WORKFLOW_STEPS.length) {
        focusStep(nextIndex, WORKFLOW_STEPS[nextIndex].id);
        // Scroll side panel to top
        scrollPanelToTop();
    }
}

/**
 * Scroll the side panel to the top
 */
function scrollPanelToTop() {
    // Try multiple selectors for the scrollable container
    const containers = [
        document.querySelector('.pf-root'),
        document.querySelector('.pf-step-guide'),
        document.body
    ];
    
    for (const container of containers) {
        if (container) {
            container.scrollTo({ top: 0, behavior: 'smooth' });
        }
    }
    
    // Also scroll window
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

function moveFocus(delta) {
    const next = appState.focusedIndex + delta;
    const target = Math.max(0, Math.min(WORKFLOW_STEPS.length - 1, next));
    focusStep(target, WORKFLOW_STEPS[target].id);
    window.scrollTo({ top: 0, behavior: "smooth" });
}

function bindSignoffNavButtons(prevButtonId, nextButtonId) {
    document.getElementById(prevButtonId)?.addEventListener("click", () => moveFocus(-1));
    document.getElementById(nextButtonId)?.addEventListener("click", () => moveFocus(1));
}

function scrollFocusedIntoView() {
    if (pendingScrollIndex === null) return;
    const card = document.querySelector(`[data-step-index="${pendingScrollIndex}"]`);
    pendingScrollIndex = null;
    card?.scrollIntoView({ behavior: "smooth", block: "center" });
}

function isStepComplete(stepId) {
    const completeFlag = parseBooleanFlag(configState.completes[stepId]);
    const hasSignoff = Boolean(configState.steps[stepId]?.signOffDate);
    return completeFlag || hasSignoff;
}

function handleStepAction(stepId) {
    switch (stepId) {
        case 1:
            importSampleData();
            break;
        case 2:
            openSheet("SS_Employee_Roster");
            break;
        case 3:
            // Validation triggered via button only for explicit user control
            break;
        case 4:
            openSheet("PTO_ExpenseReview");
            break;
        case 5:
            openSheet("PTO_JE_Draft");
            break;
        case 6:
            archiveAndReset();
            break;
        default:
            break;
    }
}

function setState(partial) {
    if (partial.stepStatuses) {
        appState.stepStatuses = { ...appState.stepStatuses, ...partial.stepStatuses };
    }
    Object.assign(appState, { ...partial, stepStatuses: appState.stepStatuses });
    renderApp();
}

function hasExcel() {
    return typeof Excel !== "undefined" && typeof Excel.run === "function";
}

async function importSampleData() {
    if (!hasExcel()) {
        return;
    }
    toggleLoader(true, "Importing sample data...");
    try {
        await Excel.run(async (context) => {
            await writeDatasetToSheet(context, "PTO_Data_Clean", PTO_ACTIVITY_COLUMNS, SAMPLE_PTO_ACTIVITY);
            await writeDatasetToSheet(context, "PTO_JE_Draft", PTO_JOURNAL_COLUMNS, SAMPLE_JOURNAL_LINES);
            const dataSheet = context.workbook.worksheets.getItem("PTO_Data_Clean");
            dataSheet.activate();
            dataSheet.getRange("A1").select();
            await context.sync();
        });

        setState({ stepStatuses: { 1: "complete" } });
    } catch (error) {
        console.error(error);
    } finally {
        toggleLoader(false);
    }
}

async function refreshValidationData() {
    if (!hasExcel()) {
        showToast("Excel is not available. Open this module inside Excel to refresh data.", "info");
        return;
    }
    toggleLoader(true, "Refreshing validation data...");
    try {
        // Validation now uses PTO_Review sheet generated in Step 2
        await runBalanceValidation();
        await runCompletenessCheck();
        toggleLoader(false);
        // Re-render step 4 if we're on it to update the completeness pills
        if (appState.activeStepId === 4) {
            renderApp();
        }
    } catch (error) {
        console.error("Refresh error:", error);
        showToast(`Failed to refresh data: ${error.message}`, "error");
        toggleLoader(false);
    }
}

/**
 * Handle save button click for missing pay rate card
 * Updates SS_Employee_Roster with the entered pay rate
 */
async function handlePayRateSave() {
    showToast("Please update pay rates in SS_Employee_Roster, then regenerate PTO_Review", "info");
    return;
}

/**
 * Handle ignore button click for missing pay rate card
 * Skips this employee and shows the next one
 */
function handlePayRateIgnore() {
    const input = document.getElementById("payrate-input");
    if (!input) return;
    
    const employeeName = input.dataset.employee;
    if (employeeName) {
        // Add to ignored set
        analysisState.ignoredMissingPayRates.add(employeeName);
        // Remove from current missing list
        analysisState.missingPayRates = analysisState.missingPayRates.filter(e => e.name !== employeeName);
    }
    
    // Re-render to show next missing employee or remove card
    focusStep(3, 3);
}

// =============================================================================
// HARDENED HOURLY RATE ENGINE
// Multi-source rate derivation with graceful fallbacks
// Never returns 0/NaN - marks as RATE_MISSING instead
// =============================================================================

/**
 * Rate source priority (highest to lowest):
 * 1. Roster override (Hourly_Rate with Is_Manually_Managed=TRUE)
 * 2. Roster computed (Hourly_Rate or Salary_Annual conversion)
 * 3. Payroll history (computed from PR_Archive_Summary)
 * 4. Default rate (configurable, explicit in UI)
 * 5. MISSING (flagged, excluded from $ totals unless default mode)
 */
const RATE_SOURCES = {
    ROSTER_OVERRIDE: "Roster Override",
    ROSTER: "Roster",
    PAYROLL: "Payroll",
    DEFAULT: "Default",
    MISSING: "Missing"
};

const DEFAULT_STD_HOURS_PER_YEAR = 2080;
const DEFAULT_HOURLY_RATE = 25; // Configurable default rate - shown explicitly in UI

/**
 * Additional columns needed for rate tracking in SS_Employee_Roster
 */
const PTO_ROSTER_RATE_COLUMNS = [
    "Hourly_Rate",              // Explicit hourly rate (currency)
    "Salary_Annual",            // Annual salary for conversion
    "Std_Hours_Per_Year",       // Standard hours (default 2080)
    "Rate_Source",              // How rate was derived
    "Rate_Last_Updated"         // ISO timestamp of last rate update
];

/**
 * Normalize employee name for consistent keying across modules
 * - Trim, collapse spaces, uppercase
 */
function normalizeEmployeeKey(name) {
    if (!name) return "";
    return String(name)
        .trim()
        .replace(/\s+/g, " ")
        .toUpperCase();
}

/**
 * State for rate engine diagnostics
 */
const rateEngineState = {
    loaded: false,
    loading: false,
    lastRun: null,
    // Rate lookup map: Employee_Key -> { rate, source, sourceDetail }
    rates: new Map(),
    // Diagnostics
    diagnostics: {
        fromRosterOverride: 0,
        fromRoster: 0,
        fromPayroll: 0,
        fromDefault: 0,
        missing: 0,
        total: 0,
        missingEmployees: [],      // [{name, key}]
        topByLiability: []         // [{name, balance, rate, liability}]
    }
};

/**
 * Load and compute hourly rates for all employees
 * Priority order: Roster Override > Roster > Payroll > Default > Missing
 */
async function loadEmployeeRates() {
    console.log("[RateEngine] Starting rate load...");
    rateEngineState.loading = true;
    rateEngineState.rates.clear();
    
    // Reset diagnostics
    const diag = rateEngineState.diagnostics;
    diag.fromRosterOverride = 0;
    diag.fromRoster = 0;
    diag.fromPayroll = 0;
    diag.fromDefault = 0;
    diag.missing = 0;
    diag.total = 0;
    diag.missingEmployees = [];
    diag.topByLiability = [];
    
    if (!hasExcel()) {
        console.warn("[RateEngine] Excel not available");
        rateEngineState.loading = false;
        return rateEngineState.rates;
    }
    
    try {
        await Excel.run(async (context) => {
            // Step 1: Load roster data
            const rosterData = await loadRosterRates(context);
            
            // Step 2: Load payroll history for rate computation
            const payrollRates = await loadPayrollDerivedRates(context);
            
            // Step 3: Get list of employees from current PTO data
            const ptoEmployees = await getPtoEmployeeList(context);
            
            // Step 4: Build final rate map with fallback chain
            for (const emp of ptoEmployees) {
                const key = normalizeEmployeeKey(emp.name);
                diag.total++;
                
                let rate = null;
                let source = RATE_SOURCES.MISSING;
                let sourceDetail = "";
                
                // Check roster override first
                const rosterEntry = rosterData.get(key);
                if (rosterEntry) {
                    if (rosterEntry.isManuallyManaged && rosterEntry.hourlyRate > 0) {
                        rate = rosterEntry.hourlyRate;
                        source = RATE_SOURCES.ROSTER_OVERRIDE;
                        sourceDetail = rosterEntry.rateSource || "Manual override";
                        diag.fromRosterOverride++;
                    } else if (rosterEntry.hourlyRate > 0) {
                        rate = rosterEntry.hourlyRate;
                        source = RATE_SOURCES.ROSTER;
                        // Use Rate_Source from roster if available (set by Payroll Recorder)
                        sourceDetail = rosterEntry.rateSource || "Hourly_Rate column";
                        diag.fromRoster++;
                    } else if (rosterEntry.salaryAnnual > 0) {
                        // Convert salary to hourly
                        const stdHours = rosterEntry.stdHoursPerYear || DEFAULT_STD_HOURS_PER_YEAR;
                        rate = rosterEntry.salaryAnnual / stdHours;
                        source = RATE_SOURCES.ROSTER;
                        sourceDetail = `Salary ${rosterEntry.salaryAnnual.toFixed(0)} / ${stdHours} hrs`;
                        diag.fromRoster++;
                    }
                }
                
                // Check payroll history if no roster rate
                if (!rate && payrollRates.has(key)) {
                    const payrollRate = payrollRates.get(key);
                    if (payrollRate.rate > 0) {
                        rate = payrollRate.rate;
                        source = RATE_SOURCES.PAYROLL;
                        sourceDetail = `Fixed $${payrollRate.avgFixedPerPeriod.toFixed(0)} / ${HOURS_PER_PAY_PERIOD}hrs (${payrollRate.periods} periods)`;
                        diag.fromPayroll++;
                    }
                }
                
                // Use default if still no rate
                if (!rate) {
                    // For now, mark as missing - we'll apply default only if user opts in
                    source = RATE_SOURCES.MISSING;
                    sourceDetail = "No rate source available";
                    diag.missing++;
                    diag.missingEmployees.push({ name: emp.name, key });
                }
                
                // Get department from roster (PTO export doesn't include department)
                const department = rosterEntry?.department || "";
                
                rateEngineState.rates.set(key, {
                    employeeName: emp.name,
                    rate: rate || 0,
                    source,
                    sourceDetail,
                    balance: emp.balance || 0,
                    department
                });
            }
            
            await context.sync();
        });
        
        // Build top 10 by liability for diagnostics
        const entriesWithLiability = Array.from(rateEngineState.rates.values())
            .filter(e => e.rate > 0 && e.balance > 0)
            .map(e => ({
                name: e.employeeName,
                balance: e.balance,
                rate: e.rate,
                liability: e.balance * e.rate
            }))
            .sort((a, b) => b.liability - a.liability)
            .slice(0, 10);
        
        diag.topByLiability = entriesWithLiability;
        
        // Log diagnostics
        console.log("[RateEngine] Load complete:", {
            total: diag.total,
            fromRosterOverride: diag.fromRosterOverride,
            fromRoster: diag.fromRoster,
            fromPayroll: diag.fromPayroll,
            fromDefault: diag.fromDefault,
            missing: diag.missing,
            missingNames: diag.missingEmployees.map(e => e.name)
        });
        console.log("[RateEngine] Top 10 by liability:", diag.topByLiability);
        
        rateEngineState.loaded = true;
        rateEngineState.loading = false;
        rateEngineState.lastRun = new Date().toISOString();
        
    } catch (error) {
        console.error("[RateEngine] Error loading rates:", error);
        rateEngineState.loading = false;
    }
    
    return rateEngineState.rates;
}

/**
 * Load rates from SS_Employee_Roster
 */
async function loadRosterRates(context) {
    const rosterMap = new Map();
    
    const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
    rosterSheet.load("isNullObject");
    await context.sync();
    
    if (rosterSheet.isNullObject) {
        console.log("[RateEngine] SS_Employee_Roster not found");
        return rosterMap;
    }
    
    const rosterRange = rosterSheet.getUsedRangeOrNullObject();
    rosterRange.load("values");
    await context.sync();
    
    if (rosterRange.isNullObject || !rosterRange.values || rosterRange.values.length < 2) {
        console.log("[RateEngine] SS_Employee_Roster is empty");
        return rosterMap;
    }
    
    const headers = rosterRange.values[0].map(h => String(h || "").toLowerCase().trim());
    const nameIdx = headers.findIndex(h => h === "employee_name" || h === "employee name" || h === "name");
    const hourlyPrismhrIdx = headers.findIndex(h => h === "hourly_rate_prismhr" || h === "hourly rate prismhr");
    const hourlyIdx = headers.findIndex(h => h === "hourly_rate" || h === "hourly rate" || h === "pay_rate" || h === "pay rate");
    const salaryIdx = headers.findIndex(h => h === "salary_annual" || h === "annual salary" || h === "salary");
    const stdHoursIdx = headers.findIndex(h => h === "std_hours_per_year" || h === "standard hours");
    const manualIdx = headers.findIndex(h => h === "is_manually_managed" || h === "manual");
    const rateSourceIdx = headers.findIndex(h => h === "rate_source" || h === "rate source");
    const deptIdx = headers.findIndex(h => h === "department" || h === "department_name" || h === "dept" || h.includes("department"));
    
    console.log("[RateEngine] Roster columns found:", { nameIdx, hourlyPrismhrIdx, hourlyIdx, salaryIdx, stdHoursIdx, manualIdx, rateSourceIdx, deptIdx });
    
    if (nameIdx < 0) {
        console.warn("[RateEngine] No employee name column in roster");
        return rosterMap;
    }
    
    for (let i = 1; i < rosterRange.values.length; i++) {
        const row = rosterRange.values[i];
        const name = String(row[nameIdx] || "").trim();
        if (!name) continue;
        
        const key = normalizeEmployeeKey(name);
        
        // Priority: Hourly_Rate_Prismhr first, then fall back to Hourly_Rate
        const hourlyRatePrismhr = hourlyPrismhrIdx >= 0 ? parseFloat(row[hourlyPrismhrIdx]) || 0 : 0;
        const hourlyRateStandard = hourlyIdx >= 0 ? parseFloat(row[hourlyIdx]) || 0 : 0;
        const hourlyRate = hourlyRatePrismhr > 0 ? hourlyRatePrismhr : hourlyRateStandard;
        
        const salaryAnnual = salaryIdx >= 0 ? parseFloat(row[salaryIdx]) || 0 : 0;
        const stdHoursPerYear = stdHoursIdx >= 0 ? parseFloat(row[stdHoursIdx]) || DEFAULT_STD_HOURS_PER_YEAR : DEFAULT_STD_HOURS_PER_YEAR;
        const isManuallyManaged = manualIdx >= 0 ? String(row[manualIdx] || "").toUpperCase() === "TRUE" : false;
        const rateSource = rateSourceIdx >= 0 ? String(row[rateSourceIdx] || "").trim() : "";
        const department = deptIdx >= 0 ? String(row[deptIdx] || "").trim() : "";
        
        rosterMap.set(key, {
            name,
            hourlyRate,
            salaryAnnual,
            stdHoursPerYear,
            isManuallyManaged,
            rateSource,
            department,
            rowIndex: i
        });
    }
    
    console.log(`[RateEngine] Loaded ${rosterMap.size} employees from roster`);
    return rosterMap;
}

/**
 * Standard hours per semi-monthly pay period
 * Semi-monthly = 24 pay periods per year, assuming 1920 work hours/year → 80 hours/period
 */
const HOURS_PER_PAY_PERIOD = 80;

/**
 * Legacy fallback: Attempt to compute hourly rates from PR_Archive_Summary
 * 
 * NOTE: PR_Archive_Summary currently stores aggregate totals per period, not per-employee.
 * This function will return empty results. The primary source for hourly rates is now
 * SS_Employee_Roster.Hourly_Rate, which is populated by the Payroll Recorder module
 * when roster updates are applied (Fixed / 80 hours calculation).
 * 
 * This function is kept as a future fallback if archive schema is enhanced to store
 * per-employee bucket totals.
 */
async function loadPayrollDerivedRates(context, periodsToAverage = 3) {
    const payrollRates = new Map();
    
    // Primary rates now come from SS_Employee_Roster.Hourly_Rate
    // This function is a no-op until archive stores per-employee data
    console.log("[RateEngine] loadPayrollDerivedRates: Rates now sourced from SS_Employee_Roster.Hourly_Rate");
    return payrollRates;
}

/**
 * Get list of employees from current PTO data
 */
async function getPtoEmployeeList(context) {
    const employees = [];
    
    // Try PTO_Data_Clean first, then PTO_Data
    const dataSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
    dataSheet.load("isNullObject");
    await context.sync();
    
    if (dataSheet.isNullObject) {
        console.log("[RateEngine] PTO_Data_Clean not found");
        return employees;
    }
    
    const dataRange = dataSheet.getUsedRangeOrNullObject();
    dataRange.load("values");
    await context.sync();
    
    if (dataRange.isNullObject || !dataRange.values || dataRange.values.length < 2) {
        return employees;
    }
    
    const headers = dataRange.values[0].map(h => String(h || "").toLowerCase().trim());
    const nameIdx = headers.findIndex(h => h === "employee_name" || h === "employee name" || h.includes("employee"));
    const balanceIdx = headers.findIndex(h => h === "balance" || h.includes("balance"));
    
    if (nameIdx < 0) {
        console.warn("[RateEngine] No employee name column in PTO data");
        return employees;
    }
    
    for (let i = 1; i < dataRange.values.length; i++) {
        const row = dataRange.values[i];
        const name = String(row[nameIdx] || "").trim();
        if (!name) continue;
        
        const balance = balanceIdx >= 0 ? parseFloat(row[balanceIdx]) || 0 : 0;
        employees.push({ name, balance, rowIndex: i });
    }
    
    console.log(`[RateEngine] Found ${employees.length} employees in PTO data`);
    return employees;
}

/**
 * Get rate for an employee by name
 * Returns { rate, source, sourceDetail } or null if not loaded
 */
function getEmployeeRate(employeeName) {
    const key = normalizeEmployeeKey(employeeName);
    return rateEngineState.rates.get(key) || null;
}

/**
 * Apply default rate to all missing employees
 * Called when user opts to use default rate
 */
function applyDefaultRateToMissing(defaultRate = DEFAULT_HOURLY_RATE) {
    let count = 0;
    for (const [key, entry] of rateEngineState.rates) {
        if (entry.source === RATE_SOURCES.MISSING) {
            entry.rate = defaultRate;
            entry.source = RATE_SOURCES.DEFAULT;
            entry.sourceDetail = `Default $${defaultRate.toFixed(2)}/hr`;
            count++;
        }
    }
    
    rateEngineState.diagnostics.fromDefault = count;
    rateEngineState.diagnostics.missing -= count;
    
    console.log(`[RateEngine] Applied default rate $${defaultRate}/hr to ${count} employees`);
    return count;
}

// =============================================================================
// PTO ARCHIVE SUMMARY - Prior Period Storage (mirrors PR_Archive_Summary)
// =============================================================================

// PTO Archive stores essential columns for journal entry change calculation:
// - Employee identification (ID + Name)
// - Vested_Balance (hours) - for reference
// - Calc_Liability (dollars) - for period-over-period change calculation
// Note: loadPriorPeriodData() still supports reading legacy columns for backward compatibility
const PTO_ARCHIVE_COLUMNS = [
    "Analysis_Date",
    "Employee_ID",          // Employee identifier (for matching)
    "Employee_Name",        // Human-readable name
    "Vested_Balance",       // Balance in hours (for reference)
    "Calc_Liability"        // Calculated liability in dollars (for change calculation)
];

const PTO_ARCHIVE_MAX_PERIODS = 5;

/**
 * Ensure PTO_Archive_Summary sheet exists with proper schema
 */
async function ensurePtoArchiveSchema() {
    if (!hasExcel()) return { ok: false, error: "Excel unavailable" };
    
    try {
        return await Excel.run(async (context) => {
            let archiveSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary");
            archiveSheet.load("isNullObject");
            await context.sync();
            
            if (archiveSheet.isNullObject) {
                console.log("[PTOArchive] Creating PTO_Archive_Summary sheet");
                archiveSheet = context.workbook.worksheets.add("PTO_Archive_Summary");
                
                const headerRange = archiveSheet.getRangeByIndexes(0, 0, 1, PTO_ARCHIVE_COLUMNS.length);
                headerRange.values = [PTO_ARCHIVE_COLUMNS];
                headerRange.format.font.bold = true;
                await context.sync();
                
                return { ok: true, created: true };
            }
            
            // Check for missing columns and add if needed
            const usedRange = archiveSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            if (!usedRange.isNullObject && usedRange.values && usedRange.values.length > 0) {
                const existingHeaders = usedRange.values[0].map(h => String(h || "").trim());
                const existingLower = new Set(existingHeaders.map(h => h.toLowerCase()));
                
                const missingColumns = PTO_ARCHIVE_COLUMNS.filter(col => 
                    !existingLower.has(col.toLowerCase())
                );
                
                if (missingColumns.length > 0) {
                    console.log("[PTOArchive] Adding missing columns:", missingColumns);
                    const startCol = existingHeaders.length;
                    const headerRange = archiveSheet.getRangeByIndexes(0, startCol, 1, missingColumns.length);
                    headerRange.values = [missingColumns];
                    headerRange.format.font.bold = true;
                    
                    // Backfill with 0 for numeric columns
                    const rowCount = usedRange.values.length - 1;
                    if (rowCount > 0) {
                        const fillData = Array(rowCount).fill(missingColumns.map(() => 0));
                        const fillRange = archiveSheet.getRangeByIndexes(1, startCol, rowCount, missingColumns.length);
                        fillRange.values = fillData;
                    }
                    await context.sync();
                }
            }
            
            return { ok: true };
        });
    } catch (error) {
        console.error("[PTOArchive] Error ensuring schema:", error);
        return { ok: false, error: error.message };
    }
}

/**
 * Load prior period data from PTO_Archive_Summary
 * Returns Map<Employee_Key, { analysisDate, liability, ... }>
 */
async function loadPriorPeriodData() {
    const priorData = new Map();
    
    if (!hasExcel()) return priorData;
    
    try {
        await Excel.run(async (context) => {
            const archiveSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary");
            archiveSheet.load("isNullObject");
            await context.sync();
            
            if (archiveSheet.isNullObject) {
                console.log("[PTOArchive] No archive sheet - all employees will show as NEW");
                return;
            }
            
            const archiveRange = archiveSheet.getUsedRangeOrNullObject();
            archiveRange.load("values");
            await context.sync();
            
            if (archiveRange.isNullObject || !archiveRange.values || archiveRange.values.length < 2) {
                return;
            }
            
            const headers = archiveRange.values[0].map(h => String(h || "").toLowerCase().trim());
            const dateIdx = headers.findIndex(h => h === "analysis_date" || h.includes("date"));
            // Support both new (employee_id) and legacy (employee_key) column names
            const idIdx = headers.findIndex(h => h === "employee_id" || h === "employee_key");
            const nameIdx = headers.findIndex(h => h === "employee_name" || h.includes("employee"));
            const balanceIdx = headers.findIndex(h => h === "vested_balance" || h === "balance");
            // Primary: use Calc_Liability (dollars) for change calculation
            // Fallback: Liability_Amount for backwards compatibility
            const liabilityIdx = headers.findIndex(h => h === "calc_liability" || h === "liability_amount" || h.includes("liability"));
            const rateIdx = headers.findIndex(h => h === "pay_rate" || h.includes("rate"));
            
            if (nameIdx < 0 && idIdx < 0) {
                console.warn("[PTOArchive] No employee identifier column");
                return;
            }
            
            // Group by employee, find most recent period for each
            const employeePeriods = new Map();
            
            for (let i = 1; i < archiveRange.values.length; i++) {
                const row = archiveRange.values[i];
                const key = idIdx >= 0 
                    ? String(row[idIdx] || "").trim()
                    : normalizeEmployeeKey(String(row[nameIdx] || ""));
                
                if (!key) continue;
                
                const analysisDate = dateIdx >= 0 ? String(row[dateIdx] || "") : "";
                const balance = balanceIdx >= 0 ? parseFloat(row[balanceIdx]) || 0 : 0;
                // Use Calc_Liability (dollars) for prior period comparison
                const liability = liabilityIdx >= 0 ? parseFloat(row[liabilityIdx]) || 0 : 0;
                const rate = rateIdx >= 0 ? parseFloat(row[rateIdx]) || 0 : 0;
                
                if (!employeePeriods.has(key)) {
                    employeePeriods.set(key, []);
                }
                employeePeriods.get(key).push({ analysisDate, liability, rate, balance });
            }
            
            // Take most recent period for each employee
            for (const [key, periods] of employeePeriods) {
                periods.sort((a, b) => String(b.analysisDate).localeCompare(String(a.analysisDate)));
                priorData.set(key, periods[0]);
            }
            
            console.log(`[PTOArchive] Loaded prior data for ${priorData.size} employees`);
        });
    } catch (error) {
        console.error("[PTOArchive] Error loading prior data:", error);
    }
    
    return priorData;
}

/**
 * Save current period to PTO_Archive_Summary
 * Maintains rolling 5 periods, removes oldest if needed
 */
async function savePtoArchivePeriod(reviewData, analysisDate) {
    if (!hasExcel() || !reviewData || reviewData.length === 0) {
        return { ok: false, error: "No data to archive" };
    }
    
    try {
        await ensurePtoArchiveSchema();
        
        return await Excel.run(async (context) => {
            const archiveSheet = context.workbook.worksheets.getItem("PTO_Archive_Summary");
            const usedRange = archiveSheet.getUsedRangeOrNullObject();
            usedRange.load("values");
            await context.sync();
            
            let headers = PTO_ARCHIVE_COLUMNS;
            let existingData = [];
            
            if (!usedRange.isNullObject && usedRange.values && usedRange.values.length > 0) {
                headers = usedRange.values[0].map(h => String(h || "").trim());
                existingData = usedRange.values.slice(1);
            }
            
            const headersLower = headers.map(h => h.toLowerCase());
            const dateIdx = headersLower.indexOf("analysis_date");
            // Support both new (employee_id) and legacy (employee_key) column names
            const empIdIdx = headersLower.indexOf("employee_id") >= 0 
                ? headersLower.indexOf("employee_id") 
                : headersLower.indexOf("employee_key");
            
            // Group existing data by period
            const periodMap = new Map();
            for (const row of existingData) {
                const periodKey = dateIdx >= 0 ? String(row[dateIdx] || "") : "";
                if (!periodKey) continue;
                
                if (!periodMap.has(periodKey)) {
                    periodMap.set(periodKey, []);
                }
                periodMap.get(periodKey).push(row);
            }
            
            // Remove current period if re-archiving (idempotent)
            const normalizedDate = String(analysisDate || "").substring(0, 10);
            periodMap.delete(normalizedDate);
            
            // Add new period
            // Simplified archive: only store essential columns for journal entry calculation
            // Employee_ID, Employee_Name, and Vested_Balance are all that's needed
            // to calculate prior period liability for the next period's journal entry
            const newRows = reviewData.map(emp => {
                const rowData = new Array(headers.length).fill("");
                headers.forEach((col, idx) => {
                    const colLower = col.toLowerCase();
                    if (colLower === "analysis_date") rowData[idx] = normalizedDate;
                    // Support both old (employee_key) and new (employee_id) column names
                    else if (colLower === "employee_id" || colLower === "employee_key") {
                        rowData[idx] = emp.employeeId || normalizeEmployeeKey(emp.employeeName);
                    }
                    else if (colLower === "employee_name") rowData[idx] = emp.employeeName || "";
                    // Vested balance is what drives liability calculation
                    else if (colLower === "vested_balance" || colLower === "balance") {
                        rowData[idx] = emp.vestedBalance ?? emp.balance ?? 0;
                    }
                    // Legacy columns - still write if they exist in older archive sheets
                    else if (colLower === "department") rowData[idx] = emp.department || "";
                    else if (colLower === "pay_rate") rowData[idx] = emp.payRate || 0;
                    else if (colLower === "accrual_rate") rowData[idx] = emp.accrualRate || 0;
                    else if (colLower === "carry_over") rowData[idx] = emp.carryOver || 0;
                    else if (colLower === "ytd_accrued") rowData[idx] = emp.ytdAccrued || 0;
                    else if (colLower === "ytd_used") rowData[idx] = emp.ytdUsed || 0;
                    else if (colLower === "calc_liability" || colLower === "liability_amount") {
                        rowData[idx] = emp.calculatedLiability ?? emp.liabilityAmount ?? 0;
                    }
                });
                return rowData;
            });
            periodMap.set(normalizedDate, newRows);
            
            // If more than 5 periods, remove oldest
            if (periodMap.size > PTO_ARCHIVE_MAX_PERIODS) {
                const sortedPeriods = Array.from(periodMap.keys()).sort();
                const toRemove = sortedPeriods.slice(0, periodMap.size - PTO_ARCHIVE_MAX_PERIODS);
                for (const period of toRemove) {
                    console.log(`[PTOArchive] Removing oldest period: ${period}`);
                    periodMap.delete(period);
                }
            }
            
            // Flatten and write
            const allRows = [];
            for (const [, rows] of periodMap) {
                allRows.push(...rows);
            }
            
            // Clear and rewrite
            const existingRange = archiveSheet.getUsedRangeOrNullObject();
            await context.sync();
            if (!existingRange.isNullObject) {
                existingRange.clear();
            }
            
            // Write headers + data
            const totalRows = 1 + allRows.length;
            const targetRange = archiveSheet.getRangeByIndexes(0, 0, totalRows, headers.length);
            targetRange.values = [headers, ...allRows];
            
            // Format header
            const headerRange = archiveSheet.getRangeByIndexes(0, 0, 1, headers.length);
            headerRange.format.font.bold = true;
            
            await context.sync();
            
            console.log(`[PTOArchive] Saved ${newRows.length} rows for period ${normalizedDate}`);
            return { ok: true, rowCount: newRows.length, periodCount: periodMap.size };
        });
    } catch (error) {
        console.error("[PTOArchive] Error saving archive:", error);
        return { ok: false, error: error.message };
    }
}

// =============================================================================
// STEP 2: PTO ACCRUAL REVIEW - Variance Table Generation
// =============================================================================

/**
 * Review table output columns (exact order, match screenshot)
 */
const PTO_REVIEW_COLUMNS = [
    "Analysis_Date",
    "Employee_Name",
    "Employee_ID",
    "Department",
    "Pay_Rate",
    "Rate_Source",
    "Accrual_Rate",
    "Carry_Over",
    "YTD_Accrued",
    "YTD_Used",
    "Vested_Balance",     // As_Of_Date_Balance - what drives liability
    "Register_Balance",   // Current_Register_Balance - for reference
    "Liability_Amount",
    "Liability_Source",
    "Report_Liability",
    "Calc_Liability",
    "Prior_Liability",
    "Change",
    "_Flags"  // Hidden helper column for flags (NEW, MISSING_RATE, LARGE_MOVE, NEG_BALANCE, RATE_VARIANCE, LIABILITY_VARIANCE)
];

/**
 * Generate the PTO Accrual Review table
 * 1. Load rates from hardened rate engine
 * 2. Load current PTO data from PTO_Data_Clean
 * 3. Load prior period from PTO_Archive_Summary
 * 4. Compute liability and change for each employee
 * 5. Write to PTO_Review sheet
 */

/**
 * Normalize employee name for comparison
 * Handles: "SMITH, JOHN" vs "JOHN SMITH" vs "Smith John" etc.
 */
function normalizeEmployeeName(name) {
    if (!name) return "";
    return String(name)
        .toUpperCase()
        .replace(/,/g, " ")           // Remove commas (SMITH, JOHN → SMITH JOHN)
        .replace(/\s+/g, " ")         // Collapse multiple spaces
        .trim();
}

/**
 * Calculate employee coverage: compare PTO report employees to roster
 * Returns reconciliation of who's in both, PTO-only, and roster-only
 */
async function calculateEmployeeCoverage(reviewData) {
    const coverage = {
        rosterCount: 0,
        ptoReportCount: reviewData.length,
        inBothCount: 0,
        inPtoOnlyCount: 0,
        inPtoOnlyNames: [],
        inPtoOnlyLiability: 0,
        inRosterOnlyCount: 0,
        inRosterOnlyNames: []
    };
    
    if (!hasExcel()) return coverage;
    
    try {
        await Excel.run(async (context) => {
            // Get roster employees
            const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
            rosterSheet.load("isNullObject");
            await context.sync();
            
            if (rosterSheet.isNullObject) {
                console.warn("[Coverage] SS_Employee_Roster not found");
                return;
            }
            
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            rosterRange.load("values");
            await context.sync();
            
            if (rosterRange.isNullObject || !rosterRange.values || rosterRange.values.length < 2) {
                return;
            }
            
            const rosterHeaders = rosterRange.values[0].map(h => String(h || "").toLowerCase().trim());
            const nameIdx = rosterHeaders.findIndex(h => h === "employee_name" || h.includes("employee"));
            const statusIdx = rosterHeaders.findIndex(h => h === "employment_status" || h.includes("status"));
            
            if (nameIdx < 0) {
                console.warn("[Coverage] Employee name column not found in roster");
                return;
            }
            
            // Build map of normalized name → original name for roster employees
            const rosterNameMap = new Map();  // normalized → original
            const rosterEmployees = new Set();
            
            for (let i = 1; i < rosterRange.values.length; i++) {
                const row = rosterRange.values[i];
                const originalName = String(row[nameIdx] || "").trim();
                const normalizedName = normalizeEmployeeName(originalName);
                const status = statusIdx >= 0 ? String(row[statusIdx] || "").toLowerCase() : "";
                
                // Include active employees (not terminated)
                if (normalizedName && !status.includes("terminated") && !status.includes("inactive")) {
                    rosterEmployees.add(normalizedName);
                    rosterNameMap.set(normalizedName, originalName);
                }
            }
            
            coverage.rosterCount = rosterEmployees.size;
            
            // Build map of normalized name → review row for PTO employees
            const ptoNameMap = new Map();  // normalized → reviewData row
            const ptoEmployees = new Set();
            
            for (const row of reviewData) {
                const normalizedName = normalizeEmployeeName(row.employeeName);
                if (normalizedName) {
                    ptoEmployees.add(normalizedName);
                    ptoNameMap.set(normalizedName, row);
                }
            }
            
            console.log("[Coverage] Roster employees:", rosterEmployees.size, "PTO employees:", ptoEmployees.size);
            
            // Find differences using normalized names
            for (const normalizedName of ptoEmployees) {
                if (rosterEmployees.has(normalizedName)) {
                    coverage.inBothCount++;
                } else {
                    coverage.inPtoOnlyCount++;
                    const originalRow = ptoNameMap.get(normalizedName);
                    if (originalRow) {
                        coverage.inPtoOnlyNames.push(originalRow.employeeName);
                        coverage.inPtoOnlyLiability += originalRow.calculatedLiability || 0;
                    }
                }
            }
            
            for (const normalizedName of rosterEmployees) {
                if (!ptoEmployees.has(normalizedName)) {
                    coverage.inRosterOnlyCount++;
                    const originalName = rosterNameMap.get(normalizedName);
                    if (originalName) {
                        coverage.inRosterOnlyNames.push(originalName);
                    }
                }
            }
            
            await context.sync();
        });
    } catch (error) {
        console.error("[Coverage] Error calculating employee coverage:", error);
    }
    
    console.log("[Coverage] Results:", coverage);
    return coverage;
}

async function generatePtoReview() {
    console.log("[PTOReview] Starting review generation...");
    ptoReviewState.loading = true;
    showToast("Generating PTO Accrual Review...", "info", 3000);
    
    try {
        // Step 1: Load employee rates
        await loadEmployeeRates();
        
        // Step 2: Load prior period data
        const priorData = await loadPriorPeriodData();
        
        // Step 3: Build review table
        const reviewData = await buildReviewTable(priorData);
        
        // Step 4: Write to PTO_Review sheet
        await writeReviewSheet(reviewData);
        
        // Step 5: Update state with executive summary and reconciliation
        let totalCurrent = 0;
        let totalPrior = 0;
        
        // Reconciliation calculations
        let reportLiabilityTotal = 0;
        let calcLiabilityTotal = 0;
        let negativeBalanceTotal = 0;
        let negativeBalanceCount = 0;
        let positiveBalanceCount = 0;
        let zeroBalanceCount = 0;
        let missingRateCount = 0;
        
        for (const row of reviewData) {
            totalCurrent += row.liabilityAmount || 0;
            totalPrior += row.priorLiability || 0;
            
            // Report liability (from PrismHR)
            reportLiabilityTotal += row.reportLiability || 0;
            
            // Calculated liability (includes negatives)
            const calcLiab = row.calculatedLiability || 0;
            calcLiabilityTotal += calcLiab;
            
            // Track balance categories
            const vested = row.vestedBalance ?? 0;
            if (vested < 0) {
                negativeBalanceTotal += calcLiab;  // Will be negative
                negativeBalanceCount++;
            } else if (vested > 0) {
                positiveBalanceCount++;
            } else {
                zeroBalanceCount++;
            }
            
            // Track actually missing rates
            if (!row.payRate || row.payRate <= 0) {
                missingRateCount++;
            }
        }
        
        // Update reconciliation state
        ptoReviewState.reconciliation = {
            reportLiabilityTotal,
            calcLiabilityTotal,
            negativeBalanceTotal,
            negativeBalanceCount,
            positiveBalanceCount,
            zeroBalanceCount,
            missingRateCount
        };
        
        // Calculate employee coverage (roster vs PTO report)
        const coverage = await calculateEmployeeCoverage(reviewData);
        ptoReviewState.coverage = coverage;
        
        ptoReviewState.totalCurrentLiability = totalCurrent;
        ptoReviewState.totalPriorLiability = totalPrior;
        ptoReviewState.netChange = totalCurrent - totalPrior;
        ptoReviewState.employeeCount = reviewData.length;
        ptoReviewState.reviewData = reviewData;
        ptoReviewState.loaded = true;
        ptoReviewState.loading = false;
        ptoReviewState.lastRun = new Date().toISOString();
        
        console.log("[PTOReview] Generation complete:", {
            employees: reviewData.length,
            totalCurrent,
            totalPrior,
            netChange: totalCurrent - totalPrior,
            reconciliation: ptoReviewState.reconciliation,
            coverage: ptoReviewState.coverage
        });
        
        showToast(`PTO Review generated: ${reviewData.length} employees, ${formatCurrency(totalCurrent)} total liability`, "success");
        renderApp();
        
    } catch (error) {
        console.error("[PTOReview] Error generating review:", error);
        showToast(`Error generating review: ${error.message}`, "error");
        ptoReviewState.loading = false;
    }
}

/**
 * Build the review table data from PTO_Data_Clean + rates + prior period
 */
async function buildReviewTable(priorData) {
    const reviewData = [];
    const analysisDate = getConfigValue(PTO_CONFIG_FIELDS.payrollDate) || new Date().toISOString().substring(0, 10);
    
    if (!hasExcel()) return reviewData;
    
    await Excel.run(async (context) => {
        // Get PTO_Data_Clean
        const dataSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
        dataSheet.load("isNullObject");
        await context.sync();
        
        if (dataSheet.isNullObject) {
            throw new Error("PTO_Data_Clean not found. Upload data in Step 1.");
        }
        
        const dataRange = dataSheet.getUsedRangeOrNullObject();
        dataRange.load("values");
        await context.sync();
        
        if (dataRange.isNullObject || !dataRange.values || dataRange.values.length < 2) {
            throw new Error("PTO_Data_Clean is empty.");
        }
        
        const headers = dataRange.values[0].map(h => String(h || "").toLowerCase().trim());
        
        // Debug: Log all headers to see what we're working with
        console.log("[PTOReview] All headers from PTO_Data_Clean:", headers);
        
        // Find columns (normalized dimension names from Obsidian)
        // Note: Be specific with employee_name to avoid matching "employees_with_liability" etc.
        // Also check for "name" as a standalone column (common in PrismHR exports)
        const colIdx = {
            employeeName: headers.findIndex(h => 
                h === "employee_name" || 
                h === "employee name" || 
                h === "employeename" ||
                h === "name" ||
                h === "full_name" ||
                h === "full name"
            ),
            department: headers.findIndex(h => h === "department" || h === "department_name"),
            accrualRate: headers.findIndex(h => h === "accrual_rate" || (h.includes("accrual") && h.includes("rate"))),
            carryOver: headers.findIndex(h => h === "carry_over" || h.includes("carry")),
            ytdAccrued: headers.findIndex(h => h === "ytd_accrued" || (h.includes("ytd") && h.includes("accrued"))),
            ytdUsed: headers.findIndex(h => h === "ytd_used" || (h.includes("ytd") && h.includes("used"))),
            balance: headers.findIndex(h => h === "balance")
        };
        
        // Find new PTO Accrued Liability report columns (if present)
        const payRateHourlyIdx = headers.findIndex(h => 
            h === "pay_rate_hourly" || h === "pay rate hourly" || h === "payratehourly"
        );
        const accruedLiabilityIdx = headers.findIndex(h => 
            h === "accrued_liability" || h === "accrued liability" || h === "accruedliability"
        );
        const asOfDateBalanceIdx = headers.findIndex(h => 
            h === "as_of_date_balance" || h === "as of date balance" || h === "asofdatebalance"
        );
        const currentRegisterBalanceIdx = headers.findIndex(h =>
            h === "current_register_balance" || h === "current register balance" || h === "currentregisterbalance"
        );
        const employeeIdIdx = headers.findIndex(h => 
            h === "employee_id" || h === "employee id" || h === "employeeid"
        );
        const reportDeptIdx = headers.findIndex(h => h === "department" || h === "department_name");
        
        console.log("[PTOReview] Column indexes:", colIdx);
        console.log("[PTOReview] Employee name found at index:", colIdx.employeeName, "header:", headers[colIdx.employeeName]);
        console.log("[PTOReview] Report columns detected:", {
            payRateHourly: payRateHourlyIdx >= 0,
            accruedLiability: accruedLiabilityIdx >= 0,
            asOfDateBalance: asOfDateBalanceIdx >= 0,
            currentRegisterBalance: currentRegisterBalanceIdx >= 0,
            employeeId: employeeIdIdx >= 0
        });
        
        if (colIdx.employeeName < 0) {
            console.error("[PTOReview] Could not find employee name column. Headers available:", headers);
            throw new Error("Employee name column not found in PTO data. Check console for available headers.");
        }
        
        // Process each employee row
        for (let i = 1; i < dataRange.values.length; i++) {
            const row = dataRange.values[i];
            const employeeName = String(row[colIdx.employeeName] || "").trim();
            if (!employeeName) continue;
            
            const key = normalizeEmployeeKey(employeeName);
            
            // Get rate and department from rate engine (department comes from SS_Employee_Roster)
            const rateInfo = getEmployeeRate(employeeName);
            
            // RATE PRIORITY: Report > Roster > Payroll > Missing
            // Priority 1: Pay Rate from uploaded report
            const reportRate = payRateHourlyIdx >= 0 ? parseCurrency(row[payRateHourlyIdx]) : 0;
            
            // Priority 2: Rate from rate engine (roster/payroll)
            const calculatedRate = rateInfo?.rate || 0;
            
            // Use report rate if available, fall back to calculated
            const payRate = reportRate > 0 ? reportRate : calculatedRate;
            const rateSource = reportRate > 0 ? "REPORT" : (rateInfo?.source || RATE_SOURCES.MISSING);
            
            // DEPARTMENT PRIORITY: Report > Roster
            const reportDepartment = reportDeptIdx >= 0 ? String(row[reportDeptIdx] || "").trim() : "";
            const rosterDepartment = rateInfo?.department || "";
            const department = reportDepartment || rosterDepartment;
            
            // Capture Employee ID if available
            const employeeId = employeeIdIdx >= 0 ? String(row[employeeIdIdx] || "").trim() : "";
            
            // Get PTO values from data
            const accrualRate = colIdx.accrualRate >= 0 ? parseFloat(row[colIdx.accrualRate]) || 0 : 0;
            const carryOver = colIdx.carryOver >= 0 ? parseFloat(row[colIdx.carryOver]) || 0 : 0;
            const ytdAccrued = colIdx.ytdAccrued >= 0 ? parseFloat(row[colIdx.ytdAccrued]) || 0 : 0;
            const ytdUsed = colIdx.ytdUsed >= 0 ? parseFloat(row[colIdx.ytdUsed]) || 0 : 0;
            
            // BALANCE PRIORITY: As_Of_Date_Balance is the vested balance that drives liability
            // Current_Register_Balance is informational only (can include future allocations)
            const asOfDateBalance = asOfDateBalanceIdx >= 0 ? parseFloat(row[asOfDateBalanceIdx]) || 0 : null;
            const currentRegisterBalance = currentRegisterBalanceIdx >= 0 
                ? parseFloat(row[currentRegisterBalanceIdx]) || 0 
                : (colIdx.balance >= 0 ? parseFloat(row[colIdx.balance]) || 0 : 0);
            
            // Use As_Of_Date_Balance for liability calculation
            // Fall back to Current_Register_Balance only if As_Of_Date_Balance column doesn't exist
            const vestedBalance = asOfDateBalance !== null ? asOfDateBalance : currentRegisterBalance;
            
            // For display purposes, show the current register balance (what employee sees)
            const displayBalance = currentRegisterBalance;
            
            // LIABILITY CALCULATION
            // Only calculate liability if there's a positive vested balance
            let liabilityAmount = 0;
            let liabilitySource = "NO_BALANCE";
            let reportLiability = 0;
            let calculatedLiability = 0;
            
            if (vestedBalance > 0) {
                // Priority 1: Use Accrued_Liability from report
                reportLiability = accruedLiabilityIdx >= 0 ? parseCurrency(row[accruedLiabilityIdx]) : 0;
                
                // Priority 2: Calculate from vested balance × rate
                calculatedLiability = payRate > 0 ? vestedBalance * payRate : 0;
                
                if (reportLiability > 0) {
                    liabilityAmount = reportLiability;
                    liabilitySource = "REPORT";
                } else {
                    liabilityAmount = calculatedLiability;
                    liabilitySource = "CALCULATED";
                }
            } else if (vestedBalance < 0) {
                // Negative balance - employee owes company (used more than accrued)
                // Per customer practice: net to zero in liability, don't show as asset
                liabilityAmount = 0;
                liabilitySource = "NEGATIVE_BALANCE";
                calculatedLiability = vestedBalance * payRate; // For reference (will be negative)
            } else {
                // Zero balance - no liability
                liabilityAmount = 0;
                liabilitySource = "NO_BALANCE";
            }
            
            console.log(`[PTOReview] ${employeeName}: vestedBal=${vestedBalance}, displayBal=${displayBalance}, liability=${liabilityAmount.toFixed(2)} (${liabilitySource})`);
            
            // Get prior period data
            // Prior liability should be Calc_Liability from archive (includes negatives)
            const prior = priorData.get(key);
            const priorLiability = prior?.liability || 0;
            
            // Change is based on calculatedLiability (includes negatives) for accurate JE
            // This ensures period-over-period comparison is accurate
            const change = calculatedLiability - priorLiability;
            
            // Determine flags
            const flags = [];
            if (!prior) flags.push("NEW");
            if (rateSource === RATE_SOURCES.MISSING) flags.push("MISSING_RATE");
            if (Math.abs(change) > ptoReviewState.flagThresholds.largeMove) flags.push("LARGE_MOVE");
            if (vestedBalance < 0) flags.push("NEG_BALANCE");
            
            // Validation: Flag significant rate variance
            if (reportRate > 0 && calculatedRate > 0) {
                const rateVariance = Math.abs(reportRate - calculatedRate);
                if (rateVariance > 5) {
                    flags.push("RATE_VARIANCE");
                    console.log(`[PTOReview] Rate variance for ${employeeName}: Report=$${reportRate.toFixed(2)} vs Calc=$${calculatedRate.toFixed(2)}`);
                }
            }
            
            // Validation: Flag significant liability variance (only if both exist and vested balance > 0)
            if (vestedBalance > 0 && reportLiability > 0 && calculatedLiability > 0) {
                const liabilityVariance = Math.abs(reportLiability - calculatedLiability);
                if (liabilityVariance > 100) {
                    flags.push("LIABILITY_VARIANCE");
                    console.log(`[PTOReview] Liability variance for ${employeeName}: Report=$${reportLiability.toFixed(2)} vs Calc=$${calculatedLiability.toFixed(2)}`);
                }
            }
            
            reviewData.push({
                analysisDate,
                employeeName,
                employeeId,
                department,
                payRate,
                rateSource,
                accrualRate,
                carryOver,
                ytdAccrued,
                ytdUsed,
                vestedBalance,        // As_Of_Date_Balance - drives liability
                displayBalance,       // Current_Register_Balance - what employee sees
                liabilityAmount,
                liabilitySource,      // "REPORT" | "CALCULATED" | "NO_BALANCE" | "NEGATIVE_BALANCE"
                reportLiability,
                calculatedLiability,
                priorLiability,
                change,
                flags: flags.join(", ")
            });
        }
        
        await context.sync();
    });
    
    // Sort by liability descending (highest liability first)
    reviewData.sort((a, b) => b.liabilityAmount - a.liabilityAmount);
    
    return reviewData;
}

/**
 * Write review table to PTO_Review sheet
 */
async function writeReviewSheet(reviewData) {
    if (!hasExcel()) return;
    
    await Excel.run(async (context) => {
        // Create or get PTO_Review sheet
        let reviewSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Review");
        reviewSheet.load("isNullObject");
        await context.sync();
        
        if (reviewSheet.isNullObject) {
            reviewSheet = context.workbook.worksheets.add("PTO_Review");
        }
        
        // Clear existing data
        const existingRange = reviewSheet.getUsedRangeOrNullObject();
        await context.sync();
        if (!existingRange.isNullObject) {
            existingRange.clear();
        }
        
        // Build data rows
        const dataRows = reviewData.map(row => [
            row.analysisDate,
            row.employeeName,
            row.employeeId || "",
            row.department,
            row.payRate,
            row.rateSource || "",
            row.accrualRate,
            row.carryOver,
            row.ytdAccrued,
            row.ytdUsed,
            row.vestedBalance,      // As_Of_Date_Balance
            row.displayBalance,     // Current_Register_Balance
            row.liabilityAmount,
            row.liabilitySource || "",
            row.reportLiability || 0,
            row.calculatedLiability || 0,
            row.priorLiability,
            row.change,
            row.flags
        ]);
        
        // Write headers + data
        const allRows = [PTO_REVIEW_COLUMNS, ...dataRows];
        const targetRange = reviewSheet.getRangeByIndexes(0, 0, allRows.length, PTO_REVIEW_COLUMNS.length);
        targetRange.values = allRows;
        
        // Format header row
        const headerRange = reviewSheet.getRangeByIndexes(0, 0, 1, PTO_REVIEW_COLUMNS.length);
        formatSheetHeaders(headerRange);
        
        // Format currency columns
        // Column order: 0=Analysis_Date, 1=Employee_Name, 2=Employee_ID, 3=Department, 
        // 4=Pay_Rate, 5=Rate_Source, 6=Accrual_Rate, 7=Carry_Over, 8=YTD_Accrued, 9=YTD_Used,
        // 10=Vested_Balance, 11=Register_Balance, 12=Liability_Amount, 13=Liability_Source,
        // 14=Report_Liability, 15=Calc_Liability, 16=Prior_Liability, 17=Change, 18=_Flags
        const currencyFormat = "$#,##0.00";
        if (dataRows.length > 0) {
            // Pay_Rate (col 4)
            reviewSheet.getRangeByIndexes(1, 4, dataRows.length, 1).numberFormat = [[currencyFormat]];
            // Liability_Amount (col 12)
            reviewSheet.getRangeByIndexes(1, 12, dataRows.length, 1).numberFormat = [[currencyFormat]];
            // Report_Liability (col 14)
            reviewSheet.getRangeByIndexes(1, 14, dataRows.length, 1).numberFormat = [[currencyFormat]];
            // Calc_Liability (col 15)
            reviewSheet.getRangeByIndexes(1, 15, dataRows.length, 1).numberFormat = [[currencyFormat]];
            // Prior_Liability (col 16)
            reviewSheet.getRangeByIndexes(1, 16, dataRows.length, 1).numberFormat = [[currencyFormat]];
            // Change (col 17)
            reviewSheet.getRangeByIndexes(1, 17, dataRows.length, 1).numberFormat = [[currencyFormat]];
        }
        
        // Autofit columns
        targetRange.format.autofitColumns();
        
        // Freeze header row
        reviewSheet.freezePanes.freezeRows(1);
        
        await context.sync();
        
        console.log(`[PTOReview] Wrote ${dataRows.length} rows to PTO_Review sheet`);
    });
}

/**
 * Helper: Format currency for display
 */
function formatCurrency(value) {
    const num = Number(value) || 0;
    return num.toLocaleString("en-US", { style: "currency", currency: "USD" });
}

/**
 * Handle PTO file upload
 * Opens file picker, reads file, and triggers header normalization
 * TODO: Implement full upload flow with ada_payroll_dimensions lookup
 */
// =============================================================================
// PTO UPLOAD STATE
// =============================================================================
const ptoUploadState = {
    file: null,
    fileName: "",
    headers: [],
    rowCount: 0,
    parsedData: null,
    mappings: null,         // { rawHeader: normalizedName } from ada_payroll_dimensions
    unmappedHeaders: [],    // Headers that couldn't be mapped
    error: null,
    loading: false
};

/**
 * Obsidian PTO Dimensions - Expected normalized column names
 * From ada_payroll_dimensions where provider='obsidian'
 */
const OBSIDIAN_DIMENSIONS = [
    "Company_Name",
    "Form_Name", 
    "Selection_Criteria",
    "Employee_Name",
    "Plan_Description",
    "Accrue_Through_Date",
    "Year_Ending"
];

const OBSIDIAN_MEASURES = [
    "Accrual_Rate",
    "Carry_Over",
    "Pay_Period_Accrued",
    "Pay_Period_Used",
    "YTD_Accrued",
    "YTD_Used",
    "Balance"
];

const ALL_OBSIDIAN_COLUMNS = [...OBSIDIAN_DIMENSIONS, ...OBSIDIAN_MEASURES];

/**
 * Handle PTO file upload
 * Opens file picker, reads file, normalizes headers via ada_payroll_dimensions
 */
async function handlePtoFileUpload(file) {
    // If file provided directly (from drag-drop or file input), process it
    if (file) {
        await processPtoUpload(file);
        return;
    }
    
    // Otherwise, create hidden file input if it doesn't exist (legacy path)
    let fileInput = document.getElementById("pto-file-input");
    if (!fileInput) {
        fileInput = document.createElement("input");
        fileInput.type = "file";
        fileInput.id = "pto-file-input";
        fileInput.accept = ".csv,.xlsx,.xls";
        fileInput.style.display = "none";
        document.body.appendChild(fileInput);
        
        fileInput.addEventListener("change", async (e) => {
            const uploadedFile = e.target.files?.[0];
            if (uploadedFile) {
                await processPtoUpload(uploadedFile);
            }
            // Reset input for re-uploads
            fileInput.value = "";
        });
    }
    
    fileInput.click();
}

/**
 * Process uploaded PTO file
 */
async function processPtoUpload(file) {
    const validExtensions = [".csv", ".xlsx", ".xls"];
    const ext = file.name.toLowerCase().slice(file.name.lastIndexOf("."));
    
    if (!validExtensions.includes(ext)) {
        showToast("Please upload a CSV or Excel file (.csv, .xlsx, .xls)", "error");
        return;
    }
    
    ptoUploadState.loading = true;
    ptoUploadState.error = null;
    ptoUploadState.fileName = file.name;
    ptoUploadState.file = file;
    
    showToast(`Reading ${file.name}...`, "info", 2000);
    updateUploadStatus("Reading file...");
    
    try {
        // Step 1: Parse the file
        const data = await parsePtoFile(file);
        if (!data || data.length < 2) {
            throw new Error("File appears empty or has no data rows.");
        }
        
        ptoUploadState.headers = data[0].map(h => String(h || "").trim());
        ptoUploadState.rowCount = data.length - 1;
        ptoUploadState.parsedData = data;
        
        console.log(`[PTOUpload] Parsed ${ptoUploadState.headers.length} columns, ${ptoUploadState.rowCount} rows`);
        console.log("[PTOUpload] Raw headers:", ptoUploadState.headers);
        
        updateUploadStatus("Loading column mappings...");
        
        // Step 2: Load customer column mappings (priority 1) and dimension mappings (fallback)
        const companyId = getConfigValue("SS_Company_ID");
        const customerMappings = await loadCustomerColumnMappings(companyId, "pto-accrual");
        const dimensionMappings = await loadObsidianDimensionMappings();
        
        console.log(`[PTOUpload] Loaded ${customerMappings.length} customer mappings, ${dimensionMappings.size} dimension mappings`);
        
        // Step 3: Normalize headers using customer mappings first, then dimension mappings
        // Also filter out columns where include_in_matrix === false
        const { normalizedHeaders, includedIndexes, unmapped, excluded } = normalizeHeadersWithCustomerMappings(
            ptoUploadState.headers, 
            customerMappings, 
            dimensionMappings
        );
        
        ptoUploadState.mappings = dimensionMappings; // Keep for compatibility
        ptoUploadState.unmappedHeaders = unmapped;
        
        if (unmapped.length > 0) {
            console.warn("[PTOUpload] Unmapped headers:", unmapped);
            showToast(`Warning: ${unmapped.length} column(s) could not be mapped`, "warning", 4000);
        }
        
        if (excluded.length > 0) {
            console.log("[PTOUpload] Excluded columns (include_in_matrix=false):", excluded);
        }
        
        console.log("[PTOUpload] Normalized headers:", normalizedHeaders);
        console.log("[PTOUpload] Included column indexes:", includedIndexes);
        
        updateUploadStatus("Writing to PTO_Data_Clean...");
        
        // Step 4: Write to PTO_Data_Clean (only included columns)
        await writePtoDataCleanFiltered(normalizedHeaders, ptoUploadState.parsedData, includedIndexes);
        
        ptoUploadState.loading = false;
        showToast(`Successfully imported ${ptoUploadState.rowCount} rows to PTO_Data_Clean`, "success");
        updateUploadStatus(`✓ ${ptoUploadState.rowCount} rows imported`);
        
        renderApp();
        
        // Step 5: Sync Employee_ID and Pay Rate to SS_Employee_Roster
        const syncResult = await syncPtoDataToRoster();
        if (syncResult.ok && syncResult.rowsUpdated > 0) {
            console.log(`[PTO] Synced ${syncResult.rowsUpdated} employees to roster (${syncResult.idUpdates} IDs, ${syncResult.rateUpdates} rates)`);
            showToast(`Updated ${syncResult.rowsUpdated} employees in roster`, "success", 2000);
        } else if (!syncResult.ok) {
            console.warn("[PTO] Roster sync skipped:", syncResult.error);
            // Don't show error toast - sync failure shouldn't block the user
        }
        
        // Step 6: Auto-run validation checks (like payroll-recorder)
        // This matches payroll-recorder behavior where Map Columns runs automatically after upload
        showToast("Running validation checks...", "info", 2000);
        await Promise.all([
            refreshHeadcountAnalysis(),
            runDataQualityCheck()
        ]);
        renderApp();
        showToast("Validation checks complete!", "success");
        
    } catch (error) {
        console.error("[PTOUpload] Error:", error);
        ptoUploadState.error = error.message;
        ptoUploadState.loading = false;
        showToast(`Upload failed: ${error.message}`, "error");
        updateUploadStatus(`Error: ${error.message}`);
    }
}

/**
 * Update the upload status display in the UI
 */
function updateUploadStatus(message) {
    const statusEl = document.getElementById("pto-upload-status");
    if (statusEl) {
        statusEl.textContent = message;
    }
}

/**
 * Sync Employee_ID and Hourly_Rate_Prismhr from PTO_Data_Clean to SS_Employee_Roster
 * Called after successful PTO file upload
 */
async function syncPtoDataToRoster() {
    if (!hasExcel()) {
        console.log("[PTO-Roster Sync] Excel not available, skipping roster sync");
        return { ok: false, error: "Excel not available" };
    }
    
    console.log("[PTO-Roster Sync] Starting sync to SS_Employee_Roster...");
    
    try {
        return await Excel.run(async (context) => {
            // Get PTO_Data_Clean
            const ptoSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
            const rosterSheet = context.workbook.worksheets.getItemOrNullObject("SS_Employee_Roster");
            
            ptoSheet.load("isNullObject");
            rosterSheet.load("isNullObject");
            await context.sync();
            
            if (ptoSheet.isNullObject) {
                console.warn("[PTO-Roster Sync] PTO_Data_Clean not found");
                return { ok: false, error: "PTO_Data_Clean not found" };
            }
            
            if (rosterSheet.isNullObject) {
                console.warn("[PTO-Roster Sync] SS_Employee_Roster not found");
                return { ok: false, error: "SS_Employee_Roster not found" };
            }
            
            // Load data from both sheets
            const ptoRange = ptoSheet.getUsedRangeOrNullObject();
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            
            ptoRange.load("values");
            rosterRange.load("values");
            await context.sync();
            
            if (ptoRange.isNullObject || !ptoRange.values || ptoRange.values.length < 2) {
                return { ok: false, error: "PTO_Data_Clean is empty" };
            }
            
            if (rosterRange.isNullObject || !rosterRange.values || rosterRange.values.length < 2) {
                return { ok: false, error: "SS_Employee_Roster is empty" };
            }
            
            // Parse PTO headers
            const ptoHeaders = ptoRange.values[0].map(h => String(h || "").toLowerCase().trim());
            const ptoIdx = {
                employeeName: ptoHeaders.findIndex(h => h === "employee_name" || h === "employee name"),
                employeeId: ptoHeaders.findIndex(h => h === "employee_id" || h === "employee id"),
                payRateHourly: ptoHeaders.findIndex(h => h === "pay_rate_hourly" || h === "pay rate hourly")
            };
            
            console.log("[PTO-Roster Sync] PTO column indexes:", ptoIdx);
            
            if (ptoIdx.employeeName < 0) {
                return { ok: false, error: "Employee_Name column not found in PTO data" };
            }
            
            // Build PTO data map (by normalized employee name)
            const ptoDataMap = new Map();
            for (let i = 1; i < ptoRange.values.length; i++) {
                const row = ptoRange.values[i];
                const name = String(row[ptoIdx.employeeName] || "").trim();
                if (!name) continue;
                
                const key = name.toLowerCase();
                ptoDataMap.set(key, {
                    employeeId: ptoIdx.employeeId >= 0 ? String(row[ptoIdx.employeeId] || "").trim() : "",
                    payRateHourly: ptoIdx.payRateHourly >= 0 ? parseCurrency(row[ptoIdx.payRateHourly]) : 0
                });
            }
            
            console.log(`[PTO-Roster Sync] Loaded ${ptoDataMap.size} employees from PTO data`);
            
            // Parse roster headers
            const rosterHeaders = rosterRange.values[0].map(h => String(h || "").toLowerCase().trim());
            const rosterIdx = {
                employeeName: rosterHeaders.findIndex(h => h === "employee_name"),
                employeeId: rosterHeaders.findIndex(h => h === "employee_id"),
                hourlyRatePrismhr: rosterHeaders.findIndex(h => h === "hourly_rate_prismhr")
            };
            
            console.log("[PTO-Roster Sync] Roster column indexes:", rosterIdx);
            
            // Check if Hourly_Rate_Prismhr column exists
            if (rosterIdx.hourlyRatePrismhr < 0) {
                console.warn("[PTO-Roster Sync] Hourly_Rate_Prismhr column not found in roster - it may need to be added via payroll-recorder schema update");
                // Continue anyway - we can still update Employee_ID
            }
            
            if (rosterIdx.employeeName < 0) {
                return { ok: false, error: "Employee_Name column not found in roster" };
            }
            
            // Update roster rows
            let updatedCount = 0;
            let idUpdates = 0;
            let rateUpdates = 0;
            
            for (let i = 1; i < rosterRange.values.length; i++) {
                const rosterRow = rosterRange.values[i];
                const rosterName = String(rosterRow[rosterIdx.employeeName] || "").trim();
                if (!rosterName) continue;
                
                const key = rosterName.toLowerCase();
                const ptoData = ptoDataMap.get(key);
                
                if (!ptoData) continue; // No matching PTO data
                
                let rowUpdated = false;
                
                // Update Employee_ID if we have one and roster is empty or different
                if (rosterIdx.employeeId >= 0 && ptoData.employeeId) {
                    const currentId = String(rosterRow[rosterIdx.employeeId] || "").trim();
                    // Only update if empty OR if current value looks like a name (contains space)
                    if (!currentId || currentId.includes(" ")) {
                        rosterSheet.getRangeByIndexes(i, rosterIdx.employeeId, 1, 1).values = [[ptoData.employeeId]];
                        rowUpdated = true;
                        idUpdates++;
                    }
                }
                
                // Update Hourly_Rate_Prismhr if column exists and we have a rate
                if (rosterIdx.hourlyRatePrismhr >= 0 && ptoData.payRateHourly > 0) {
                    rosterSheet.getRangeByIndexes(i, rosterIdx.hourlyRatePrismhr, 1, 1).values = [[ptoData.payRateHourly]];
                    rowUpdated = true;
                    rateUpdates++;
                }
                
                if (rowUpdated) updatedCount++;
            }
            
            await context.sync();
            
            const result = {
                ok: true,
                totalMatched: ptoDataMap.size,
                rowsUpdated: updatedCount,
                idUpdates,
                rateUpdates
            };
            
            console.log("[PTO-Roster Sync] Complete:", result);
            return result;
        });
        
    } catch (error) {
        console.error("[PTO-Roster Sync] Error:", error);
        return { ok: false, error: error.message };
    }
}

/**
 * Parse a PTO file (CSV or Excel) using XLSX library
 */
async function parsePtoFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = e.target.result;
                const workbook = XLSX.read(data, { type: "array" });
                
                const sheetName = workbook.SheetNames[0];
                if (!sheetName) {
                    throw new Error("No sheets found in workbook");
                }
                
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
                
                // Find header row (might not be first row)
                let headerRowIndex = 0;
                for (let i = 0; i < Math.min(10, jsonData.length); i++) {
                    const row = jsonData[i];
                    // Check if this row looks like headers (has common PTO terms)
                    const rowStr = row.map(c => String(c || "").toLowerCase()).join(" ");
                    if (rowStr.includes("employee") || rowStr.includes("balance") || rowStr.includes("accrual")) {
                        headerRowIndex = i;
                        break;
                    }
                }
                
                console.log(`[PTOUpload] Header row detected at index ${headerRowIndex}`);
                
                // Return data starting from header row
                resolve(jsonData.slice(headerRowIndex));
                
            } catch (err) {
                reject(err);
            }
        };
        
        reader.onerror = () => reject(new Error("Failed to read file"));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Load customer column mappings from ada_customer_column_mappings
 * Priority 1 source for header normalization (customer-specific)
 * 
 * @param {string} companyId - Company UUID
 * @param {string} module - Module name (e.g., "pto-accrual")
 * @returns {Promise<Array<{raw_header: string, pf_column_name: string, include_in_matrix: boolean}>>}
 */
async function loadCustomerColumnMappings(companyId, module) {
    if (!companyId) {
        console.warn("[PTOUpload] No companyId provided, skipping customer mappings");
        return [];
    }
    
    console.log(`[PTOUpload] Loading customer column mappings for company=${companyId}, module=${module}`);
    
    try {
        const apiUrl = `${SUPABASE_URL}/rest/v1/ada_customer_column_mappings?company_id=eq.${encodeURIComponent(companyId)}&module=eq.${encodeURIComponent(module)}&select=raw_header,pf_column_name,include_in_matrix`;
        
        const response = await fetch(apiUrl, {
            headers: {
                "apikey": SUPABASE_KEY,
                "Authorization": `Bearer ${SUPABASE_KEY}`,
                "Content-Type": "application/json"
            }
        });
        
        if (!response.ok) {
            console.warn(`[PTOUpload] Customer mappings API error: ${response.status}`);
            return [];
        }
        
        const data = await response.json();
        console.log(`[PTOUpload] Loaded ${data.length} customer column mappings`);
        
        return data;
        
    } catch (error) {
        console.error("[PTOUpload] Error loading customer mappings:", error);
        return [];
    }
}

/**
 * Normalize headers using customer mappings (priority 1) then dimension mappings (fallback)
 * Also filters out columns where include_in_matrix === false
 * 
 * @param {string[]} rawHeaders - Original headers from uploaded file
 * @param {Array} customerMappings - Customer column mappings with include_in_matrix
 * @param {Map} dimensionMappings - Fallback dimension mappings
 * @returns {{ normalizedHeaders: string[], includedIndexes: number[], unmapped: string[], excluded: string[] }}
 */
function normalizeHeadersWithCustomerMappings(rawHeaders, customerMappings, dimensionMappings) {
    const normalizedHeaders = [];
    const includedIndexes = [];
    const unmapped = [];
    const excluded = [];
    
    // Build customer mapping lookup (case-insensitive)
    const customerMap = new Map();
    for (const mapping of customerMappings) {
        const key = normalizeForMatching(mapping.raw_header);
        customerMap.set(key, mapping);
    }
    
    for (let i = 0; i < rawHeaders.length; i++) {
        const rawHeader = rawHeaders[i];
        const key = normalizeForMatching(rawHeader);
        
        // Priority 1: Customer column mapping
        const customerMapping = customerMap.get(key);
        if (customerMapping) {
            // Check include_in_matrix
            const includeInMatrix = customerMapping.include_in_matrix;
            if (includeInMatrix === false || includeInMatrix === "false") {
                excluded.push(rawHeader);
                console.log(`[PTOUpload] Excluding column "${rawHeader}" (include_in_matrix=false)`);
                continue; // Skip this column
            }
            
            normalizedHeaders.push(customerMapping.pf_column_name);
            includedIndexes.push(i);
            continue;
        }
        
        // Priority 2: Dimension mapping (fallback)
        const dimensionNormalized = dimensionMappings.get(key);
        if (dimensionNormalized) {
            normalizedHeaders.push(dimensionNormalized);
            includedIndexes.push(i);
            continue;
        }
        
        // Priority 3: Keep original header with cleanup
        const fallbackName = rawHeader
            .trim()
            .replace(/[^a-zA-Z0-9\s]/g, "")
            .split(/\s+/)
            .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
            .join("_");
        
        normalizedHeaders.push(fallbackName || `Column_${normalizedHeaders.length + 1}`);
        includedIndexes.push(i);
        
        if (rawHeader.trim()) {
            unmapped.push(rawHeader);
        }
    }
    
    return { normalizedHeaders, includedIndexes, unmapped, excluded };
}

/**
 * Convert Excel serial date to JavaScript Date
 * Excel dates are days since 1900-01-01 (with a leap year bug for 1900)
 */
function excelSerialToDate(serial) {
    if (typeof serial !== "number" || serial < 1) return null;
    // Excel's epoch is 1900-01-01, but it incorrectly treats 1900 as a leap year
    // So we need to adjust for dates after Feb 28, 1900
    const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
    const date = new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
    return date;
}

/**
 * Format a date value to MM/DD/YYYY string
 * Handles: Excel serial numbers, Date objects, and various string formats
 */
function formatDateValue(value) {
    if (!value) return "";
    
    // Already a formatted date string (contains / or -)
    if (typeof value === "string") {
        const trimmed = value.trim();
        // If it looks like a date string already, return as-is
        if (trimmed.match(/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/)) {
            return trimmed;
        }
        // Try parsing ISO format
        if (trimmed.match(/^\d{4}-\d{2}-\d{2}/)) {
            const date = new Date(trimmed);
            if (!isNaN(date.getTime())) {
                return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
            }
        }
        return trimmed;
    }
    
    // Excel serial date number
    if (typeof value === "number" && value > 1000 && value < 100000) {
        const date = excelSerialToDate(value);
        if (date) {
            return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
        }
    }
    
    // Date object
    if (value instanceof Date && !isNaN(value.getTime())) {
        return `${value.getMonth() + 1}/${value.getDate()}/${value.getFullYear()}`;
    }
    
    return String(value);
}

/**
 * Check if a column header indicates it's a date column
 * Be careful to exclude columns that contain "date" but aren't dates (e.g., as_of_date_balance)
 */
function isDateColumn(headerName) {
    const lower = headerName.toLowerCase();
    
    // Explicit exclusions - columns that contain "date" but aren't date columns
    if (lower.includes("balance") || lower.includes("amount") || lower.includes("liability")) {
        return false;
    }
    
    // Check for date-related patterns
    return lower.includes("_date") ||           // hire_date, termination_date, etc.
           lower.includes("date_") ||           // date_hired, etc. (but not as_of_date_balance)
           lower === "date" ||                  // Just "date"
           lower.endsWith(" date") ||           // "hire date", "termination date"
           lower.startsWith("date ") ||         // "date hired"
           lower.includes("hire") || 
           lower.includes("termination") ||
           lower.includes("start_date") ||
           lower.includes("end_date") ||
           lower === "as_of";
}

/**
 * Write normalized data to PTO_Data_Clean sheet (filtered by includedIndexes)
 */
async function writePtoDataCleanFiltered(normalizedHeaders, parsedData, includedIndexes) {
    if (!hasExcel()) {
        throw new Error("Excel runtime not available");
    }
    
    // Identify date columns for formatting
    const dateColumnIndexes = [];
    normalizedHeaders.forEach((header, idx) => {
        if (isDateColumn(header)) {
            dateColumnIndexes.push(idx);
            console.log(`[PTOUpload] Detected date column: "${header}" at index ${idx}`);
        }
    });
    
    await Excel.run(async (context) => {
        // Create or get PTO_Data_Clean sheet
        let cleanSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
        cleanSheet.load("isNullObject");
        await context.sync();
        
        if (cleanSheet.isNullObject) {
            cleanSheet = context.workbook.worksheets.add("PTO_Data_Clean");
            console.log("[PTOUpload] Created PTO_Data_Clean sheet");
        }
        
        // Clear existing data
        const existingRange = cleanSheet.getUsedRangeOrNullObject();
        await context.sync();
        if (!existingRange.isNullObject) {
            existingRange.clear();
        }
        
        // Build clean data with only included columns
        const cleanData = [normalizedHeaders];
        
        // Add data rows (skip original header row)
        for (let i = 1; i < parsedData.length; i++) {
            const row = parsedData[i];
            // Only include columns that passed the filter
            const cleanRow = includedIndexes.map((colIdx, headerIdx) => {
                const value = row[colIdx];
                
                // Format date columns
                if (dateColumnIndexes.includes(headerIdx)) {
                    return formatDateValue(value);
                }
                
                return value ?? "";
            });
            cleanData.push(cleanRow);
        }
        
        // Write all data
        const targetRange = cleanSheet.getRangeByIndexes(0, 0, cleanData.length, normalizedHeaders.length);
        targetRange.values = cleanData;
        
        // Format header row
        const headerRange = cleanSheet.getRangeByIndexes(0, 0, 1, normalizedHeaders.length);
        formatSheetHeaders(headerRange);
        
        // Apply date format to date columns
        if (cleanData.length > 1 && dateColumnIndexes.length > 0) {
            for (const colIdx of dateColumnIndexes) {
                const dateColRange = cleanSheet.getRangeByIndexes(1, colIdx, cleanData.length - 1, 1);
                dateColRange.numberFormat = [["mm/dd/yyyy"]];
            }
        }
        
        // Autofit columns
        targetRange.format.autofitColumns();
        
        // Freeze header row
        cleanSheet.freezePanes.freezeRows(1);
        
        await context.sync();
        
        console.log(`[PTOUpload] Wrote ${cleanData.length - 1} rows, ${normalizedHeaders.length} columns to PTO_Data_Clean`);
    });
}

/**
 * Load Obsidian dimension mappings from ada_payroll_dimensions via edge function
 * Maps raw_term -> normalized_dimension for provider='obsidian'
 */
async function loadObsidianDimensionMappings() {
    console.log("[PTOUpload] Loading Obsidian dimension mappings via edge function...");
    
    try {
        const edgeFunctionUrl = `${SUPABASE_URL}/functions/v1/column-mapper`;
        
        const response = await fetch(edgeFunctionUrl, {
            method: "POST",
            headers: {
                "apikey": SUPABASE_KEY,
                "Authorization": `Bearer ${SUPABASE_KEY}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify({
                action: "get_dimensions",
                provider: "obsidian"
            })
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`Edge function error: ${response.status} - ${errorText}`);
        }
        
        const result = await response.json();
        
        if (!result.success || !result.dimensions) {
            throw new Error("Edge function returned unsuccessful response");
        }
        
        console.log(`[PTOUpload] Loaded ${result.dimensions.length} dimension mappings from edge function`);
        
        // Build mapping: normalized raw_term -> normalized_dimension
        const mappings = new Map();
        for (const row of result.dimensions) {
            const rawKey = normalizeForMatching(row.raw_term);
            const normalizedValue = row.normalized_dimension;
            if (rawKey && normalizedValue) {
                mappings.set(rawKey, normalizedValue);
            }
        }
        
        console.log(`[PTOUpload] Built ${mappings.size} unique mappings`);
        
        // If no mappings from DB, use fallback hardcoded mappings for common Obsidian headers
        if (mappings.size === 0) {
            console.log("[PTOUpload] No DB mappings found, using fallback mappings");
            const fallbackMappings = getFallbackObsidianMappings();
            return fallbackMappings;
        }
        
        return mappings;
        
    } catch (error) {
        console.error("[PTOUpload] Error loading dimension mappings:", error);
        // Return fallback mappings on error
        console.log("[PTOUpload] Using fallback mappings due to error");
        return getFallbackObsidianMappings();
    }
}

/**
 * Fallback Obsidian header mappings (if DB is empty or unavailable)
 * Based on common Obsidian PTO report headers
 */
function getFallbackObsidianMappings() {
    const mappings = new Map();
    
    // Dimensions
    mappings.set(normalizeForMatching("Company Name"), "Company_Name");
    mappings.set(normalizeForMatching("Form Name"), "Form_Name");
    mappings.set(normalizeForMatching("Selection Criteria"), "Selection_Criteria");
    mappings.set(normalizeForMatching("Employee Name"), "Employee_Name");
    mappings.set(normalizeForMatching("Employee"), "Employee_Name");
    mappings.set(normalizeForMatching("Name"), "Employee_Name");
    mappings.set(normalizeForMatching("Plan Description"), "Plan_Description");
    mappings.set(normalizeForMatching("Plan"), "Plan_Description");
    mappings.set(normalizeForMatching("Accrue Through Date"), "Accrue_Through_Date");
    mappings.set(normalizeForMatching("Accrue Thru Date"), "Accrue_Through_Date");
    mappings.set(normalizeForMatching("Year Ending"), "Year_Ending");
    mappings.set(normalizeForMatching("Year End"), "Year_Ending");
    
    // Measures
    mappings.set(normalizeForMatching("Accrual Rate"), "Accrual_Rate");
    mappings.set(normalizeForMatching("Rate"), "Accrual_Rate");
    mappings.set(normalizeForMatching("Carry Over"), "Carry_Over");
    mappings.set(normalizeForMatching("Carryover"), "Carry_Over");
    mappings.set(normalizeForMatching("Pay Period Accrued"), "Pay_Period_Accrued");
    mappings.set(normalizeForMatching("PP Accrued"), "Pay_Period_Accrued");
    mappings.set(normalizeForMatching("Period Accrued"), "Pay_Period_Accrued");
    mappings.set(normalizeForMatching("Pay Period Used"), "Pay_Period_Used");
    mappings.set(normalizeForMatching("PP Used"), "Pay_Period_Used");
    mappings.set(normalizeForMatching("Period Used"), "Pay_Period_Used");
    mappings.set(normalizeForMatching("YTD Accrued"), "YTD_Accrued");
    mappings.set(normalizeForMatching("Year To Date Accrued"), "YTD_Accrued");
    mappings.set(normalizeForMatching("YTD Used"), "YTD_Used");
    mappings.set(normalizeForMatching("Year To Date Used"), "YTD_Used");
    mappings.set(normalizeForMatching("Balance"), "Balance");
    mappings.set(normalizeForMatching("Current Balance"), "Balance");
    mappings.set(normalizeForMatching("Available Balance"), "Balance");
    
    console.log(`[PTOUpload] Fallback mappings: ${mappings.size} entries`);
    return mappings;
}

/**
 * Normalize a string for matching (trim, lowercase, remove special chars)
 */
function normalizeForMatching(value) {
    if (!value) return "";
    return String(value)
        .trim()
        .toLowerCase()
        .replace(/[_\-.]/g, " ")  // Replace underscores, hyphens, dots with space
        .replace(/\s+/g, " ");     // Collapse multiple spaces
}

/**
 * Normalize headers using the dimension mappings
 * Returns { normalizedHeaders: string[], unmapped: string[] }
 */
function normalizeHeaders(rawHeaders, mappings) {
    const normalizedHeaders = [];
    const unmapped = [];
    
    for (const rawHeader of rawHeaders) {
        const key = normalizeForMatching(rawHeader);
        const normalized = mappings.get(key);
        
        if (normalized) {
            normalizedHeaders.push(normalized);
        } else {
            // Keep original header for unmapped columns
            // Convert to PascalCase with underscores for consistency
            const fallbackName = rawHeader
                .trim()
                .replace(/[^a-zA-Z0-9\s]/g, "")
                .split(/\s+/)
                .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
                .join("_");
            
            normalizedHeaders.push(fallbackName || `Column_${normalizedHeaders.length + 1}`);
            if (rawHeader.trim()) {
                unmapped.push(rawHeader);
            }
        }
    }
    
    return { normalizedHeaders, unmapped };
}

/**
 * Write normalized data to PTO_Data_Clean sheet
 */
async function writePtoDataClean(normalizedHeaders, parsedData) {
    if (!hasExcel()) {
        throw new Error("Excel runtime not available");
    }
    
    await Excel.run(async (context) => {
        // Create or get PTO_Data_Clean sheet
        let cleanSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
        cleanSheet.load("isNullObject");
        await context.sync();
        
        if (cleanSheet.isNullObject) {
            cleanSheet = context.workbook.worksheets.add("PTO_Data_Clean");
            console.log("[PTOUpload] Created PTO_Data_Clean sheet");
        }
        
        // Clear existing data
        const existingRange = cleanSheet.getUsedRangeOrNullObject();
        await context.sync();
        if (!existingRange.isNullObject) {
            existingRange.clear();
        }
        
        // Build clean data with normalized headers
        const cleanData = [normalizedHeaders];
        
        // Add data rows (skip original header row)
        for (let i = 1; i < parsedData.length; i++) {
            const row = parsedData[i];
            // Ensure row has same length as headers
            const cleanRow = normalizedHeaders.map((_, colIdx) => {
                const value = row[colIdx];
                // Keep values as-is (Excel will handle type conversion)
                return value ?? "";
            });
            cleanData.push(cleanRow);
        }
        
        // Write all data
        const targetRange = cleanSheet.getRangeByIndexes(0, 0, cleanData.length, normalizedHeaders.length);
        targetRange.values = cleanData;
        
        // Format header row
        const headerRange = cleanSheet.getRangeByIndexes(0, 0, 1, normalizedHeaders.length);
        formatSheetHeaders(headerRange);
        
        // Autofit columns
        targetRange.format.autofitColumns();
        
        // Freeze header row
        cleanSheet.freezePanes.freezeRows(1);
        
        await context.sync();
        
        console.log(`[PTOUpload] Wrote ${cleanData.length - 1} rows to PTO_Data_Clean`);
        
        // Log PTO MODULE TRACE (matching payroll-recorder style)
        console.log("\n════════════════════════════════════════════════════════════════");
        console.log("PTO MODULE TRACE - Upload Complete");
        console.log("════════════════════════════════════════════════════════════════");
        console.log(`  Rows: ${cleanData.length - 1}`);
        console.log(`  Columns: ${normalizedHeaders.length}`);
        console.log(`  Headers: ${normalizedHeaders.slice(0, 8).join(", ")}${normalizedHeaders.length > 8 ? "..." : ""}`);
        console.log(`  Dimensions: ${OBSIDIAN_DIMENSIONS.filter(d => normalizedHeaders.includes(d)).length} of ${OBSIDIAN_DIMENSIONS.length}`);
        console.log(`  Measures: ${OBSIDIAN_MEASURES.filter(m => normalizedHeaders.includes(m)).length} of ${OBSIDIAN_MEASURES.length}`);
        console.log("════════════════════════════════════════════════════════════════\n");
    });
}

async function runDataQualityCheck() {
    if (!hasExcel()) {
        showToast("Excel is not available. Open this module inside Excel to run quality check.", "info");
        return;
    }
    
    dataQualityState.loading = true;
    toggleLoader(true, "Analyzing data quality...");
    updateSaveButtonState(document.getElementById("quality-save-btn"), false);
    
    try {
        await Excel.run(async (context) => {
            const dataSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
            dataSheet.load("isNullObject");
            await context.sync();
            
            if (dataSheet.isNullObject) {
                throw new Error("PTO_Data_Clean not found. Upload data in Step 1.");
            }
            
            const dataRange = dataSheet.getUsedRangeOrNullObject();
            dataRange.load("values");
            await context.sync();
            
            const dataValues = dataRange.isNullObject ? [] : dataRange.values || [];
            
            if (!dataValues.length || dataValues.length < 2) {
                throw new Error("PTO_Data_Clean is empty or has no data rows.");
            }
            
            // Parse headers
            const headers = (dataValues[0] || []).map(h => normalizeName(h));
            console.log("[Data Quality] PTO data headers:", dataValues[0]);
            
            // Find employee name column - be specific to avoid matching company name
            let nameIdx = headers.findIndex(h => h === "employee name" || h === "employeename");
            if (nameIdx === -1) {
                // Fallback: look for column containing "employee" AND "name"
                nameIdx = headers.findIndex(h => h.includes("employee") && h.includes("name"));
            }
            if (nameIdx === -1) {
                // Last resort: just "name" but not if it also contains "company" or "form"
                nameIdx = headers.findIndex(h => h === "name" || (h.includes("name") && !h.includes("company") && !h.includes("form")));
            }
            console.log("[Data Quality] Employee name column index:", nameIdx, "Header:", dataValues[0]?.[nameIdx]);
            const balanceIdx = findColumnIndex(headers, ["balance"]);
            const accrualRateIdx = findColumnIndex(headers, ["accrual rate", "accrualrate"]);
            const carryOverIdx = findColumnIndex(headers, ["carry over", "carryover"]);
            const ytdAccruedIdx = findColumnIndex(headers, ["ytd accrued", "ytdaccrued"]);
            const ytdUsedIdx = findColumnIndex(headers, ["ytd used", "ytdused"]);
            
            // Reset state
            const balanceIssues = [];
            const zeroBalances = [];
            const accrualOutliers = [];
            
            // Process each employee row
            const dataRows = dataValues.slice(1);
            dataRows.forEach((row, idx) => {
                const rowIndex = idx + 2; // 1-based, after header
                const name = nameIdx !== -1 ? String(row[nameIdx] || "").trim() : `Row ${rowIndex}`;
                if (!name) return;
                
                const balance = balanceIdx !== -1 ? Number(row[balanceIdx]) || 0 : 0;
                const accrualRate = accrualRateIdx !== -1 ? Number(row[accrualRateIdx]) || 0 : 0;
                const carryOver = carryOverIdx !== -1 ? Number(row[carryOverIdx]) || 0 : 0;
                const ytdAccrued = ytdAccruedIdx !== -1 ? Number(row[ytdAccruedIdx]) || 0 : 0;
                const ytdUsed = ytdUsedIdx !== -1 ? Number(row[ytdUsedIdx]) || 0 : 0;
                
                // Check 1: Balance issues - negative balance or used more than available
                const maxUsable = carryOver + ytdAccrued;
                if (balance < 0) {
                    balanceIssues.push({ 
                        name, 
                        issue: `Negative balance: ${balance.toFixed(2)} hrs`,
                        rowIndex 
                    });
                } else if (ytdUsed > maxUsable && maxUsable > 0) {
                    balanceIssues.push({ 
                        name, 
                        issue: `Used ${ytdUsed.toFixed(0)} hrs but only ${maxUsable.toFixed(0)} available`,
                        rowIndex 
                    });
                }
                
                // Check 2: Zero balances (informational)
                if (balance === 0 && (carryOver > 0 || ytdAccrued > 0)) {
                    zeroBalances.push({ name, rowIndex });
                }
                
                // Check 3: Accrual rate outliers (> 8 hrs per period is unusual)
                if (accrualRate > 8) {
                    accrualOutliers.push({ name, accrualRate, rowIndex });
                }
            });
            
            // Update state
            dataQualityState.balanceIssues = balanceIssues;
            dataQualityState.zeroBalances = zeroBalances;
            dataQualityState.accrualOutliers = accrualOutliers;
            dataQualityState.totalIssues = balanceIssues.length;
            dataQualityState.totalEmployees = dataRows.filter(r => r.some(c => c !== null && c !== "")).length;
            dataQualityState.hasRun = true;
        });
        
        // Update step status
        const hasBlockingIssues = dataQualityState.balanceIssues.length > 0;
        setState({ stepStatuses: { 3: hasBlockingIssues ? "blocked" : "complete" } });
        
    } catch (error) {
        console.error("Data quality check error:", error);
        showToast(`Quality check failed: ${error.message}`, "error");
        dataQualityState.hasRun = false;
    } finally {
        dataQualityState.loading = false;
        toggleLoader(false);
        renderApp();
    }
}

/**
 * User acknowledges quality issues and wants to proceed anyway
 */
function acknowledgeQualityIssues() {
    dataQualityState.acknowledged = true;
    // Allow sign-off even with issues
    setState({ stepStatuses: { 3: "complete" } });
    renderApp();
}

// Note: Save functions removed - auto-save happens via config writes during analysis

/**
 * Run balance validation between PTO_Data_Clean and PTO_Analysis totals
 * TODO: Implement actual balance validation logic
 */
async function runBalanceValidation() {
    if (!hasExcel()) return;
    
    console.log("[PTO] runBalanceValidation called - validation pending implementation");
    // Future: Compare totals between PTO_Data_Clean and PTO_Analysis sheets
    // For now, this is a no-op placeholder
}

/**
 * Run data completeness check comparing PTO_Data_Clean to PTO_Analysis column sums
 */
async function runCompletenessCheck() {
    if (!hasExcel()) return;
    
    try {
        await Excel.run(async (context) => {
            const dataSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
            dataSheet.load("isNullObject");
            await context.sync();
            
            if (dataSheet.isNullObject) {
                console.warn("[Completeness] PTO_Data_Clean not found");
                return;
            }
            
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            
            const dataRange = dataSheet.getUsedRangeOrNullObject();
            dataRange.load("values");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            if (analysisSheet.isNullObject) {
                // Reset all checks to null if no analysis sheet
                analysisState.completenessCheck = {
                    accrualRate: null,
                    carryOver: null,
                    ytdAccrued: null,
                    ytdUsed: null,
                    balance: null
                };
                return;
            }
            
            const analysisRange = analysisSheet.getUsedRangeOrNullObject();
            analysisRange.load("values");
            await context.sync();
            
            const dataValues = dataRange.isNullObject ? [] : dataRange.values || [];
            const analysisValues = analysisRange.isNullObject ? [] : analysisRange.values || [];
            
            if (!dataValues.length || !analysisValues.length) {
                analysisState.completenessCheck = {
                    accrualRate: null,
                    carryOver: null,
                    ytdAccrued: null,
                    ytdUsed: null,
                    balance: null
                };
                return;
            }
            
            // Helper to find column and sum values
            const sumColumn = (rows, columnAliases, label) => {
                const headers = (rows[0] || []).map(h => normalizeName(h));
                const idx = findColumnIndex(headers, columnAliases);
                if (idx === -1) return null;
                const dataRows = rows.slice(1);
                return dataRows.reduce((sum, row) => sum + (Number(row[idx]) || 0), 0);
            };
            
            // Column mappings for each field
            const fields = [
                { key: "accrualRate", aliases: ["accrual rate", "accrualrate"] },
                { key: "carryOver", aliases: ["carry over", "carryover", "carry_over"] },
                { key: "ytdAccrued", aliases: ["ytd accrued", "ytdaccrued", "ytd_accrued"] },
                { key: "ytdUsed", aliases: ["ytd used", "ytdused", "ytd_used"] },
                { key: "balance", aliases: ["balance"] }
            ];
            
            const results = {};
            
            for (const field of fields) {
                const ptoDataCleanSum = sumColumn(dataValues, field.aliases, "PTO_Data_Clean");
                const analysisSum = sumColumn(analysisValues, field.aliases, "PTO_Analysis");
                
                if (ptoDataCleanSum === null || analysisSum === null) {
                    results[field.key] = null;
                } else {
                    // Use tolerance for floating point comparison
                    const match = Math.abs(ptoDataCleanSum - analysisSum) < 0.01;
                    results[field.key] = {
                        match,
                        ptoDataClean: ptoDataCleanSum,
                        ptoAnalysis: analysisSum
                    };
                }
            }
            
            analysisState.completenessCheck = results;
        });
    } catch (error) {
        console.error("Completeness check failed:", error);
    }
}

/**
 * Run full analysis - syncs data and runs all verification checks
 */
async function runFullAnalysis() {
    if (!hasExcel()) {
        showToast("Excel is not available. Open this module inside Excel to run analysis.", "info");
        return;
    }
    
    toggleLoader(true, "Running analysis...");
    
    try {
        // Analysis now uses PTO_Review sheet - ensure it's generated first
        
        // 2. Run completeness check (compare column sums)
        await runCompletenessCheck();
        
        // 3. Update state
        analysisState.cleanDataReady = true;
        
        // 4. Re-render to show results
        renderApp();
        
    } catch (error) {
        console.error("Full analysis error:", error);
        showToast(`Analysis failed: ${error.message}`, "error");
    } finally {
        toggleLoader(false);
    }
}

async function populatePtoAnalysis() {
    if (!hasExcel()) {
        analysisState.lastError = "Excel is not available. Open this module inside Excel to run analysis.";
        renderApp();
        return;
    }
    
    analysisState.loading = true;
    analysisState.lastError = null;
    renderApp();
    
    try {
        // Get the row count from PTO_Analysis
        const result = await Excel.run(async (context) => {
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            if (analysisSheet.isNullObject) {
                throw new Error("PTO_Analysis sheet was not created. Please check that PTO_Data_Clean has data.");
            }
            
            const analysisRange = analysisSheet.getUsedRangeOrNullObject();
            analysisRange.load("values");
            await context.sync();
            
            const values = analysisRange.isNullObject ? [] : analysisRange.values || [];
            const dataRows = values.length > 1 ? values.length - 1 : 0;
            
            if (dataRows === 0) {
                throw new Error("No employee data found. Please import PTO data in Step 1 first.");
            }
            
            return { employeeCount: dataRows };
        });
        
        // Update state with success
        analysisState.cleanDataReady = true;
        analysisState.employeeCount = result.employeeCount;
        analysisState.lastRun = new Date().toISOString();
        analysisState.lastError = null;
        
        setState({ stepStatuses: { 4: "complete" } });
    } catch (error) {
        console.error(error);
        analysisState.lastError = error?.message || "An unexpected error occurred while running the analysis.";
        analysisState.lastRun = null;
    } finally {
        analysisState.loading = false;
        renderApp();
    }
}

/**
 * Check PTO_Analysis status when entering Step 4
 */
async function checkAnalysisPrerequisites() {
    if (!hasExcel()) {
        analysisState.cleanDataReady = false;
        analysisState.employeeCount = 0;
        return;
    }
    
    try {
        const result = await Excel.run(async (context) => {
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            if (analysisSheet.isNullObject) {
                return { ready: false, count: 0 };
            }
            
            const analysisRange = analysisSheet.getUsedRangeOrNullObject();
            analysisRange.load("values");
            await context.sync();
            
            const values = analysisRange.isNullObject ? [] : analysisRange.values || [];
            const dataRows = values.length > 1 ? values.length - 1 : 0;
            
            return { ready: dataRows > 0, count: dataRows };
        });
        
        analysisState.cleanDataReady = result.ready;
        analysisState.employeeCount = result.count;
    } catch (error) {
        console.error("Error checking analysis prerequisites:", error);
        analysisState.cleanDataReady = false;
        analysisState.employeeCount = 0;
    }
}

async function runJournalSummary() {
    if (!hasExcel()) {
        showToast("Excel is not available. Open this module inside Excel to run journal checks.", "info");
        return;
    }
    journalState.loading = true;
    journalState.lastError = null;
    updateSaveButtonState(document.getElementById("je-save-btn"), false);
    renderApp();
    try {
        const totals = await Excel.run(async (context) => {
            // Read JE Draft
            const jeSheet = context.workbook.worksheets.getItem("PTO_JE_Draft");
            const jeRange = jeSheet.getUsedRangeOrNullObject();
            jeRange.load("values");
            
            // Read PTO_Analysis for comparison
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            analysisSheet.load("isNullObject");
            await context.sync();
            
            const jeValues = jeRange.isNullObject ? [] : jeRange.values || [];
            if (!jeValues.length) {
                throw new Error("PTO_JE_Draft is empty. Generate the JE first.");
            }
            
            // Parse JE headers (QuickBooks format: JournalNo, JournalDate, Account Name, Debits, Credits, Description)
            const jeHeaders = (jeValues[0] || []).map((h) => normalizeName(h));
            const debitIdx = findColumnIndex(jeHeaders, ["debits", "debit"]);
            const creditIdx = findColumnIndex(jeHeaders, ["credits", "credit"]);
            const acctNameIdx = findColumnIndex(jeHeaders, ["account name", "accountname"]);
            
            if (debitIdx === -1 || creditIdx === -1) {
                throw new Error("Could not find Debits and Credits columns in PTO_JE_Draft.");
            }
            
            let debitTotal = 0;
            let creditTotal = 0;
            let jeExpenseTotal = 0; // Sum of expense line amounts (not clearing account offset)
            
            jeValues.slice(1).forEach((row) => {
                const debit = Number(row[debitIdx]) || 0;
                const credit = Number(row[creditIdx]) || 0;
                const acctName = acctNameIdx !== -1 ? String(row[acctNameIdx] || "").trim().toLowerCase() : "";
                
                debitTotal += debit;
                creditTotal += credit;
                
                // Sum expense lines only (not the Payroll Clearing Account offset)
                if (acctName && !acctName.includes("clearing")) {
                    jeExpenseTotal += debit - credit;  // Net = debit minus credit per line
                }
            });
            
            // Get PTO_Analysis total change
            let analysisChangeTotal = 0;
            if (!analysisSheet.isNullObject) {
                const analysisRange = analysisSheet.getUsedRangeOrNullObject();
                analysisRange.load("values");
                await context.sync();
                
                const analysisValues = analysisRange.isNullObject ? [] : analysisRange.values || [];
                if (analysisValues.length > 1) {
                    const analysisHeaders = (analysisValues[0] || []).map(h => normalizeName(h));
                    const changeIdx = findColumnIndex(analysisHeaders, ["change"]);
                    
                    if (changeIdx !== -1) {
                        analysisValues.slice(1).forEach(row => {
                            analysisChangeTotal += Number(row[changeIdx]) || 0;
                        });
                    }
                }
            }
            
            // Build validation issues array
            const difference = debitTotal - creditTotal;
            const issues = [];
            
            // Check 1: Debits = Credits
            if (Math.abs(difference) >= 0.01) {
                issues.push({
                    check: "Debits = Credits",
                    passed: false,
                    detail: difference > 0 
                        ? `Debits exceed credits by $${Math.abs(difference).toLocaleString(undefined, {minimumFractionDigits: 2})}`
                        : `Credits exceed debits by $${Math.abs(difference).toLocaleString(undefined, {minimumFractionDigits: 2})}`
                });
            } else {
                issues.push({ check: "Debits = Credits", passed: true, detail: "" });
            }
            
            // Check 2: Line Amounts Sum to Zero (Debits = Credits means this is balanced)
            // For QuickBooks format, this is the same as Check 1
            issues.push({ check: "Line Amounts Sum to Zero", passed: Math.abs(difference) < 0.01, detail: Math.abs(difference) < 0.01 ? "" : `Difference: $${difference.toFixed(2)}` });
            
            // Check 3: JE Matches Analysis Total
            const changeDiff = Math.abs(jeExpenseTotal - analysisChangeTotal);
            if (changeDiff >= 0.01) {
                issues.push({
                    check: "JE Matches Analysis Total",
                    passed: false,
                    detail: `JE expense total ($${jeExpenseTotal.toLocaleString(undefined, {minimumFractionDigits: 2})}) differs from PTO_Analysis Change total ($${analysisChangeTotal.toLocaleString(undefined, {minimumFractionDigits: 2})}) by $${changeDiff.toLocaleString(undefined, {minimumFractionDigits: 2})}`
                });
            } else {
                issues.push({ check: "JE Matches Analysis Total", passed: true, detail: "" });
            }
            
            return { 
                debitTotal, 
                creditTotal, 
                difference,
                jeChangeTotal: jeExpenseTotal,
                analysisChangeTotal,
                issues,
                validationRun: true
            };
        });
        Object.assign(journalState, totals, { lastError: null });
    } catch (error) {
        console.warn("PTO JE summary:", error);
        journalState.lastError = error?.message || "Unable to calculate journal totals.";
        journalState.debitTotal = null;
        journalState.creditTotal = null;
        journalState.difference = null;
        journalState.jeChangeTotal = null;
        journalState.analysisChangeTotal = null;
        journalState.issues = [];
        journalState.validationRun = false;
    } finally {
        journalState.loading = false;
        renderApp();
    }
}

/**
 * Department to Expense Account mapping for PTO accrual entries
 */
const DEPARTMENT_EXPENSE_ACCOUNTS = {
    "general & administrative": "64110",
    "general and administrative": "64110",
    "g&a": "64110",
    "research & development": "62110",
    "research and development": "62110",
    "r&d": "62110",
    "marketing": "61610",
    "cogs onboarding": "53110",
    "cogs prof. services": "56110",
    "cogs professional services": "56110",
    "sales & marketing": "61110",
    "sales and marketing": "61110",
    "cogs support": "52110",
    "client success": "61811"
};

const LIABILITY_OFFSET_ACCOUNT = "21540";

// =============================================================================
// PTO JOURNAL ENTRY GENERATION (Step 3)
// Real JE, not allocation - simpler than payroll-recorder
// Only ONE measure drives the JE: PTO_Liability_Change
// Offset account hardcoded: 21540 (PTO Liability)
// =============================================================================

const PTO_LIABILITY_OFFSET_ACCOUNT = "21540";
const PTO_GL_COLUMN_NAME = "PTO_Liability_Change";
const PTO_MODULE_KEY = "pto-accrual";

/**
 * Normalize key for GL mapping lookup (same as payroll-recorder)
 */
function jeNormalizeKey(value) {
    return String(value ?? "")
        .replace(/\u00a0/g, " ")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, " ")
        .replace(/&/g, "and");
}

/**
 * Load GL mappings for PTO_Liability_Change from ada_customer_gl_mappings
 * Returns Map<normalizedDepartment, { gl_account, gl_account_name }>
 */
async function loadPtoGLMappings(companyId) {
    console.log("[PTO-JE] Loading GL mappings for company:", companyId);
    
    try {
        const apiUrl = `${SUPABASE_URL}/rest/v1/ada_customer_gl_mappings?company_id=eq.${companyId}&module=eq.${PTO_MODULE_KEY}&pf_column_name=eq.${PTO_GL_COLUMN_NAME}&select=department,gl_account,gl_account_name`;
        
        const response = await fetch(apiUrl, {
            headers: {
                "apikey": SUPABASE_KEY,
                "Authorization": `Bearer ${SUPABASE_KEY}`,
                "Content-Type": "application/json"
            }
        });
        
        if (!response.ok) {
            const errorText = await response.text();
            throw new Error(`API error fetching GL mappings: ${response.status} - ${errorText}`);
        }
        
        const data = await response.json();
        console.log(`[PTO-JE] Fetched ${data.length} GL mapping rules`);
        
        // Build mapping: normalized department -> { gl_account, gl_account_name }
        const mappings = new Map();
        for (const row of data) {
            const deptKey = jeNormalizeKey(row.department);
            if (!deptKey) continue;
            mappings.set(deptKey, {
                gl_account: row.gl_account,
                gl_account_name: row.gl_account_name || ""
            });
        }
        
        console.log(`[PTO-JE] Built ${mappings.size} department -> GL mappings`);
        return mappings;
        
    } catch (error) {
        console.error("[PTO-JE] Error loading GL mappings:", error);
        throw error;
    }
}

/**
 * Aggregate Change by Department from PTO_Review
 * Returns Map<department, changeTotal>
 */
async function aggregatePtoChangeByDepartment() {
    const deptTotals = new Map();
    
    if (!hasExcel()) {
        throw new Error("Excel runtime not available");
    }
    
    await Excel.run(async (context) => {
        const reviewSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Review");
        reviewSheet.load("isNullObject");
        await context.sync();
        
        if (reviewSheet.isNullObject) {
            throw new Error("PTO_Review sheet not found. Run Step 2 (PTO Accrual Review) first.");
        }
        
        const reviewRange = reviewSheet.getUsedRangeOrNullObject();
        reviewRange.load("values");
        await context.sync();
        
        if (reviewRange.isNullObject || !reviewRange.values || reviewRange.values.length < 2) {
            throw new Error("PTO_Review is empty. Run Step 2 first.");
        }
        
        const headers = reviewRange.values[0].map(h => String(h || "").toLowerCase().trim());
        const deptIdx = headers.findIndex(h => h === "department");
        const changeIdx = headers.findIndex(h => h === "change");
        
        if (deptIdx < 0 || changeIdx < 0) {
            throw new Error(`Required columns not found in PTO_Review. Found: ${headers.join(", ")}`);
        }
        
        // Aggregate by department
        for (let i = 1; i < reviewRange.values.length; i++) {
            const row = reviewRange.values[i];
            const dept = String(row[deptIdx] || "").trim();
            const change = parseFloat(row[changeIdx]) || 0;
            
            if (!dept || Math.abs(change) < 0.01) continue;
            
            const current = deptTotals.get(dept) || 0;
            deptTotals.set(dept, current + change);
        }
        
        await context.sync();
    });
    
    console.log("[PTO-JE] Department totals:", Object.fromEntries(deptTotals));
    return deptTotals;
}

/**
 * Call Ada API for PTO insights
 * Mirrors payroll-recorder callAdaApi function
 * Can be called with either:
 * - Object params: { systemPrompt, userPrompt, contextPack, functionContext }
 * - Positional params: (prompt, context, messageHistory) - for copilot.js compatibility
 */
async function callAdaApi(promptOrParams, context, messageHistory) {
    // Handle both calling conventions
    let systemPrompt, userPrompt, contextPack, functionContext;
    
    if (typeof promptOrParams === 'object' && promptOrParams !== null && !Array.isArray(promptOrParams) && promptOrParams.userPrompt !== undefined) {
        // Called with object params (original style)
        ({ systemPrompt, userPrompt, contextPack, functionContext } = promptOrParams);
    } else {
        // Called with positional params (copilot.js style)
        userPrompt = promptOrParams;
        contextPack = context;
        functionContext = "analysis";
    }

    // Supabase copilot endpoint
    const COPILOT_URL = "https://jgciqwzwacaesqjaoadc.supabase.co/functions/v1/copilot";
    const SUPABASE_ANON_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpnY2lxd3p3YWNhZXNxamFvYWRjIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjAzODgzMTIsImV4cCI6MjA3NTk2NDMxMn0.DsoUTHcm1Uv65t4icaoD0Tzf3ULIU54bFnoYw8hHScE";

    // Get customer ID from installation state
    const customerId = installationState.company_id;

    try {
        console.log("[Ada] Calling copilot API...", { module: PTO_MODULE_KEY, function: functionContext || "analysis", customerId: customerId ? "set" : "not set" });

        const response = await fetch(COPILOT_URL, {
            method: "POST",
            mode: "cors",
            credentials: "omit",
            headers: {
                "Content-Type": "application/json",
                "Authorization": `Bearer ${SUPABASE_ANON_KEY}`,
                "apikey": SUPABASE_ANON_KEY
            },
            body: JSON.stringify({
                prompt: userPrompt,
                context: contextPack,
                module: PTO_MODULE_KEY,
                function: functionContext || "analysis",
                customerId: customerId,
                // Only pass systemPrompt if we want to override the database config
                ...(systemPrompt ? { systemPrompt } : {})
            })
        });

        if (!response.ok) {
            const errorText = await response.text();
            console.error("[Ada] API error:", response.status, errorText);
            throw new Error(`API request failed: ${response.status}`);
        }

        const data = await response.json();
        console.log("[Ada] API response received:", data.usage || "no usage info");

        if (data.message || data.response) {
            return data.message || data.response;
        }

        // Fallback if no message in response
        console.warn("[Ada] No message in API response, using local generation");
        return generateLocalPtoInsights(params);

    } catch (error) {
        console.warn("[Ada] API call failed, using local generation:", error);
        return generateLocalPtoInsights(params);
    }
}

/**
 * Generate local PTO insights when API is unavailable
 */
function generateLocalPtoInsights(params) {
    const { userPrompt, contextPack } = params;
    const lowerPrompt = (userPrompt || "").toLowerCase();

    if (lowerPrompt.includes('diagnostic') || lowerPrompt.includes('check')) {
        return `**PTO Data Diagnostics:**

✓ **Data Completeness**: PTO file uploaded and processed successfully
✓ **Header Mapping**: All required columns identified and mapped
✓ **Rate Engine**: Pay rate calculations completed
✓ **Balance Calculations**: PTO balances computed for all employees

⚠️ **Potential Issues**:
• ${contextPack?.summary?.missingRateCount || 0} employees missing pay rates
• Balance validation may be needed for large changes

**Next Steps:**
1. Review any missing pay rate warnings
2. Generate the accrual review table
3. Proceed to journal entry preparation`;
    }

    if (lowerPrompt.includes('insight') || lowerPrompt.includes('analysis')) {
        return `**PTO Accrual Analysis:**

📊 **Key Metrics:**
• Total Current Liability: $${(contextPack?.summary?.totalCurrent || 0).toLocaleString()}
• Total Prior Liability: $${(contextPack?.summary?.totalPrior || 0).toLocaleString()}
• Net Change: $${(contextPack?.summary?.netChange || 0).toLocaleString()}
• Active Employees: ${contextPack?.summary?.employeeCount || 0}

💡 **Executive Insights:**
1. **Liability Trend**: ${(contextPack?.summary?.netChange || 0) >= 0 ? 'Increasing' : 'Decreasing'} PTO liability
2. **Employee Coverage**: ${contextPack?.summary?.employeeCount || 0} employees with PTO tracking
3. **Pay Rate Completeness**: ${(contextPack?.summary?.missingRateCount || 0) === 0 ? 'Complete' : 'Missing rates for some employees'}

**Recommendations:**
• Review any employees with missing pay rates
• Monitor PTO liability trends quarter-over-quarter
• Ensure accurate employee headcount for future accruals`;
    }

    // Default response
    return `**PTO Analysis Assistant**

I can help you analyze your PTO accrual data. Try asking me about:

• **Diagnostics** - Check data quality and completeness
• **Insights** - Key findings and trends
• **Balances** - Employee PTO balance analysis
• **Accruals** - Liability calculation details

Your PTO data includes:
- ${contextPack?.summary?.employeeCount || 0} employees
- Current liability: $${(contextPack?.summary?.totalCurrent || 0).toLocaleString()}
- Net change: $${(contextPack?.summary?.netChange || 0).toLocaleString()}

What would you like to explore?`;
}

/**
 * Create PTO Journal Entry Draft
 * Groups Change amounts by Department and creates proper debit/credit entries
 * Uses GL mappings from ada_customer_gl_mappings
 */
async function createJournalDraft() {
    if (!hasExcel()) {
        showToast("Excel is not available. Open this module inside Excel to create the journal entry.", "info");
        return;
    }
    
    toggleLoader(true, "Creating PTO Journal Entry...");
    
    try {
        // Get company_id from installation state
        const companyId = installationState.company_id;
        if (!companyId) {
            throw new Error("Company ID not found. Please check your installation configuration.");
        }
        
        // Get JournalNo and JournalDate from config
        const journalNo = getConfigValue(PTO_CONFIG_FIELDS.journalEntryId) || "";
        const rawJournalDate = getConfigValue(PTO_CONFIG_FIELDS.payrollDate) || "";
        
        console.log("\n════════════════════════════════════════════════════════════════");
        console.log("PTO JOURNAL ENTRY GENERATION (QuickBooks Format)");
        console.log("════════════════════════════════════════════════════════════════");
        console.log(`  Company ID: ${companyId}`);
        console.log(`  JournalNo (PTO_Journal_Entry_ID): ${journalNo}`);
        console.log(`  JournalDate (PTO_Payroll_Date): ${rawJournalDate}`);
        
        if (!journalNo) {
            throw new Error("Journal Entry ID not set. Please enter a Journal Entry ID in the Configuration step.");
        }
        
        if (!rawJournalDate) {
            throw new Error("Payroll Date not set. Please enter a Payroll Date in the Configuration step.");
        }
        
        // Format date for QuickBooks (MM/DD/YYYY)
        let formattedDate = rawJournalDate;
        try {
            let d;
            if (typeof rawJournalDate === "number" || /^\d{4,5}$/.test(String(rawJournalDate).trim())) {
                const serialNum = Number(rawJournalDate);
                const excelEpoch = new Date(1899, 11, 30);
                d = new Date(excelEpoch.getTime() + serialNum * 24 * 60 * 60 * 1000);
            } else {
                d = new Date(rawJournalDate);
            }
            
            if (!isNaN(d.getTime()) && d.getFullYear() > 1970) {
                const mm = String(d.getMonth() + 1).padStart(2, "0");
                const dd = String(d.getDate()).padStart(2, "0");
                const yyyy = d.getFullYear();
                formattedDate = `${mm}/${dd}/${yyyy}`;
            }
        } catch (e) {
            console.warn("[PTO-JE] Could not parse date, using as-is:", rawJournalDate);
        }
        
        // Step 1: Load GL mappings
        const glMappings = await loadPtoGLMappings(companyId);
        console.log(`  GL Mappings loaded: ${glMappings.size}`);
        
        if (glMappings.size === 0) {
            throw new Error(
                `No GL mappings found for PTO_Liability_Change.\n\n` +
                `Please configure GL mappings in ada_customer_gl_mappings:\n` +
                `  - company_id: ${companyId}\n` +
                `  - module: ${PTO_MODULE_KEY}\n` +
                `  - pf_column_name: ${PTO_GL_COLUMN_NAME}\n` +
                `  - department: (your department names)\n` +
                `  - gl_account: (expense account numbers)`
            );
        }
        
        // Step 2: Aggregate Change by Department from PTO_Review
        const deptTotals = await aggregatePtoChangeByDepartment();
        console.log(`  Departments with changes: ${deptTotals.size}`);
        
        if (deptTotals.size === 0) {
            throw new Error(
                "No department totals found with non-zero changes.\n\n" +
                "Please run Step 2 (PTO Accrual Review) first to generate change data."
            );
        }
        
        // Step 3: Validate all departments have GL mappings
        const unmappedDepts = [];
        for (const [dept] of deptTotals) {
            const key = jeNormalizeKey(dept);
            if (!glMappings.has(key)) {
                unmappedDepts.push(dept);
            }
        }
        
        if (unmappedDepts.length > 0) {
            throw new Error(
                `Missing GL mappings for departments:\n` +
                unmappedDepts.map(d => `  • ${d}`).join("\n") +
                `\n\nPlease add GL mappings for these departments in ada_customer_gl_mappings.`
            );
        }
        
        // Step 3b: Load Chart of Accounts for account name lookup
        let chartOfAccountsLookup = new Map();  // Account Number → Account Name
        await Excel.run(async (context) => {
            try {
                const coaSheet = context.workbook.worksheets.getItemOrNullObject("SS_Chart_of_Accounts");
                coaSheet.load("isNullObject");
                await context.sync();
                
                if (!coaSheet.isNullObject) {
                    const coaUsedRange = coaSheet.getUsedRangeOrNullObject();
                    coaUsedRange.load("isNullObject");
                    await context.sync();
                    
                    if (!coaUsedRange.isNullObject) {
                        coaUsedRange.load("values");
                        await context.sync();
                        
                        if (coaUsedRange.values && coaUsedRange.values.length > 1) {
                            const coaHeaders = coaUsedRange.values[0];
                            const coaRows = coaUsedRange.values.slice(1);
                            
                            // Find account number and name columns
                            const acctNumIdx = coaHeaders.findIndex(h => {
                                const normalized = String(h || "").toLowerCase().trim().replace(/[^a-z0-9]/g, "_");
                                return normalized.includes("account") && normalized.includes("num") ||
                                       normalized === "account_number" || normalized === "number";
                            });
                            
                            const acctNameIdx = coaHeaders.findIndex(h => {
                                const normalized = String(h || "").toLowerCase().trim().replace(/[^a-z0-9]/g, "_");
                                return normalized.includes("account") && normalized.includes("name") ||
                                       normalized === "account_name" || normalized === "name";
                            });
                            
                            if (acctNumIdx >= 0 && acctNameIdx >= 0) {
                                for (const row of coaRows) {
                                    const acctNumber = String(row[acctNumIdx] || "").trim();
                                    const acctName = String(row[acctNameIdx] || "").trim();
                                    if (acctNumber) {
                                        chartOfAccountsLookup.set(acctNumber, acctName);
                                    }
                                }
                                console.log(`  Chart of Accounts: ${chartOfAccountsLookup.size} accounts loaded`);
                            }
                        }
                    }
                }
            } catch (coaError) {
                console.warn("[PTO-JE] Error reading SS_Chart_of_Accounts (non-fatal):", coaError.message);
            }
        });
        
        // Step 4: Build QuickBooks-format JE rows
        // Headers: JournalNo, JournalDate, Account Name, Debits, Credits, Description
        const jeHeaders = ["JournalNo", "JournalDate", "Account Name", "Debits", "Credits", "Description"];
        const jeDataRows = [];
        
        let totalDebits = 0;
        let totalCredits = 0;
        
        for (const [dept, change] of deptTotals) {
            if (Math.abs(change) < 0.01) continue;
            
            const key = jeNormalizeKey(dept);
            const mapping = glMappings.get(key);
            const glAccountStr = String(mapping.gl_account || "").trim();
            
            // QuickBooks format: "AccountNumber AccountName"
            // e.g., "52160 Support PEO:Support Onshore Labor:Support 401k Employer Contribution"
            let accountName;
            const accountNameFromCOA = chartOfAccountsLookup.get(glAccountStr);
            if (accountNameFromCOA) {
                accountName = `${glAccountStr} ${accountNameFromCOA}`;
            } else if (mapping.gl_account_name) {
                accountName = `${glAccountStr} ${mapping.gl_account_name}`;
            } else {
                accountName = glAccountStr;
            }
            
            const absAmount = Math.abs(change);
            
            // Description = Department + PF_column_name
            const description = `${dept}${dept ? " - " : ""}PTO Liability Change`;
            
            // QuickBooks format:
            // If change > 0 (liability increased): Debit expense account
            // If change < 0 (liability decreased): Credit expense account (ABS of negative)
            if (change > 0) {
                totalDebits += absAmount;
                jeDataRows.push([
                    journalNo,
                    formattedDate,
                    accountName,
                    absAmount,   // Debits = amount > 0
                    "",          // Credits = blank
                    description
                ]);
            } else {
                totalCredits += absAmount;
                jeDataRows.push([
                    journalNo,
                    formattedDate,
                    accountName,
                    "",          // Debits = blank
                    absAmount,   // Credits = ABS of amount < 0
                    description
                ]);
            }
        }
        
        // Add Accrued PTO Liability offset line to balance the JE
        // PTO accruals: Debit PTO Expense (by dept), Credit Accrued PTO Liability
        const PTO_LIABILITY_ACCOUNT = "21540 Accrued Expenses:Accrued PTO Liability";
        const offsetAmount = totalDebits - totalCredits;
        
        if (Math.abs(offsetAmount) >= 0.01) {
            if (offsetAmount > 0) {
                // Net debits - need to credit liability account (normal case for accrual increase)
                jeDataRows.push([
                    journalNo,
                    formattedDate,
                    PTO_LIABILITY_ACCOUNT,
                    "",              // Debits = blank
                    offsetAmount,    // Credits = offset amount
                    "Accrued PTO Liability"
                ]);
                totalCredits += offsetAmount;
            } else {
                // Net credits - need to debit liability account (accrual decrease/payout)
                jeDataRows.push([
                    journalNo,
                    formattedDate,
                    PTO_LIABILITY_ACCOUNT,
                    Math.abs(offsetAmount),  // Debits = ABS of offset
                    "",                      // Credits = blank
                    "Accrued PTO Liability"
                ]);
                totalDebits += Math.abs(offsetAmount);
            }
        }
        
        // Step 5: Validate debits == credits
        const tolerance = 0.01;
        if (Math.abs(totalDebits - totalCredits) > tolerance) {
            throw new Error(
                `JE is out of balance!\n` +
                `  Total Debits: ${totalDebits.toFixed(2)}\n` +
                `  Total Credits: ${totalCredits.toFixed(2)}\n` +
                `  Difference: ${(totalDebits - totalCredits).toFixed(2)}\n\n` +
                `Cannot write unbalanced journal entry.`
            );
        }
        
        console.log(`  JE Lines: ${jeDataRows.length}`);
        console.log(`  Total Debits: $${totalDebits.toFixed(2)}`);
        console.log(`  Total Credits: $${totalCredits.toFixed(2)}`);
        console.log("════════════════════════════════════════════════════════════════\n");
        
        // Step 6: Write to PTO_JE_Draft
        await Excel.run(async (context) => {
            let jeSheet = context.workbook.worksheets.getItemOrNullObject("PTO_JE_Draft");
            jeSheet.load("isNullObject");
            await context.sync();
            
            if (jeSheet.isNullObject) {
                jeSheet = context.workbook.worksheets.add("PTO_JE_Draft");
                console.log("[PTO-JE] Created PTO_JE_Draft sheet");
            } else {
                // Make sure sheet is visible (it may be hidden by tab visibility)
                jeSheet.visibility = Excel.SheetVisibility.visible;
                const usedRange = jeSheet.getUsedRangeOrNullObject();
                await context.sync();
                if (!usedRange.isNullObject) {
                    usedRange.clear();
                }
                console.log("[PTO-JE] Cleared existing PTO_JE_Draft");
            }
            
            // Build all rows
            const allRows = [jeHeaders, ...jeDataRows];
            const writeRange = jeSheet.getRangeByIndexes(0, 0, allRows.length, jeHeaders.length);
            writeRange.values = allRows;
            
            // Format headers
            const headerRange = jeSheet.getRangeByIndexes(0, 0, 1, jeHeaders.length);
            formatSheetHeaders(headerRange);
            
            // Format currency columns (Debits col 3, Credits col 4)
            const dataRowCount = jeDataRows.length;
            if (dataRowCount > 0) {
                const currencyFormat = "$#,##0.00";
                jeSheet.getRangeByIndexes(1, 3, dataRowCount, 1).numberFormat = [[currencyFormat]];
                jeSheet.getRangeByIndexes(1, 4, dataRowCount, 1).numberFormat = [[currencyFormat]];
            }
            
            writeRange.format.autofitColumns();
            jeSheet.freezePanes.freezeRows(1);
            jeSheet.activate();
            
            await context.sync();
        });
        
        // Update journal state
        journalState.debitTotal = totalDebits;
        journalState.creditTotal = totalCredits;
        journalState.validationRun = true;
        journalState.lastError = null;
        
        // Run validation checks after generation (matches payroll-recorder pattern)
        await runJournalSummary();
        
        showToast(`Journal Entry created: ${jeDataRows.length} lines (including offset), $${totalDebits.toFixed(2)} balanced ✓`, "success");
        renderApp();
        
    } catch (error) {
        console.error("[PTO-JE] Error:", error);
        journalState.lastError = error.message;
        showToast(`Unable to create Journal Entry: ${error.message}`, "error");
    } finally {
        toggleLoader(false);
    }
}

async function exportJournalDraft() {
    if (!hasExcel()) {
        showToast("Excel is not available. Open this module inside Excel to export.", "info");
        return;
    }
    toggleLoader(true, "Preparing JE CSV...");
    try {
        const { rows } = await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem("PTO_JE_Draft");
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            const values = range.isNullObject ? [] : range.values || [];
            if (!values.length) {
                throw new Error("PTO_JE_Draft is empty.");
            }
            return { rows: values };
        });
        
        // Normalize date and amount columns for QBO-ready export
        const headers = (rows[0] || []).map((h) => String(h || "").trim().toLowerCase());
        const debitIdx = headers.findIndex((h) => h.includes("debit"));
        const creditIdx = headers.findIndex((h) => h.includes("credit"));
        const dateIdx = headers.findIndex((h) => h.includes("journaldate") || h.includes("txndate") || h === "date");
        
        const normalizedRows = rows.map((row, idx) => {
            if (idx === 0) return row;
            const next = [...(row || [])];
            const normalizeAmount = (v) => {
                if (v === null || v === "") return "";
                const n = Number(v);
                if (!Number.isFinite(n)) return "";
                return n.toFixed(2);
            };
            // Normalize date column - Excel may return serial numbers
            const normalizeDate = (v) => {
                if (v === null || v === "") return "";
                // If it's already a string in MM/DD/YYYY format, return as-is
                if (typeof v === "string" && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(v.trim())) {
                    return v.trim();
                }
                // Handle Excel serial number
                if (typeof v === "number" && Number.isFinite(v)) {
                    const ms = Math.round((v - 25569) * 86400 * 1000);
                    const d = new Date(ms);
                    if (Number.isFinite(d.getTime())) {
                        const mm = String(d.getMonth() + 1).padStart(2, "0");
                        const dd = String(d.getDate()).padStart(2, "0");
                        const yyyy = d.getFullYear();
                        if (yyyy >= 1900 && yyyy <= 2100) {
                            return `${mm}/${dd}/${yyyy}`;
                        }
                    }
                }
                // Handle Date object
                if (v instanceof Date && Number.isFinite(v.getTime())) {
                    const mm = String(v.getMonth() + 1).padStart(2, "0");
                    const dd = String(v.getDate()).padStart(2, "0");
                    const yyyy = v.getFullYear();
                    return `${mm}/${dd}/${yyyy}`;
                }
                return String(v);
            };
            if (debitIdx >= 0) next[debitIdx] = normalizeAmount(next[debitIdx]);
            if (creditIdx >= 0) next[creditIdx] = normalizeAmount(next[creditIdx]);
            if (dateIdx >= 0) next[dateIdx] = normalizeDate(next[dateIdx]);
            return next;
        });
        
        const csv = buildCsv(normalizedRows);
        downloadCsv(`pto-je-draft-${todayIso()}.csv`, csv);
    } catch (error) {
        console.error("PTO JE export:", error);
        showToast("Unable to export the JE draft. Confirm the sheet has data.", "error");
    } finally {
        toggleLoader(false);
    }
}

/**
 * Open the accounting software URL from SS_PF_Config (SS_Accounting_Software)
 */
async function openAccountingSoftware() {
    let accountingUrl = getConfigValue(PTO_CONFIG_FIELDS.accountingSoftware) || getConfigValue("SS_Accounting_Software");

    if (!accountingUrl && hasExcelRuntime()) {
        try {
            const configValues = await loadConfigFromTable(CONFIG_TABLES);
            accountingUrl =
                configValues["SS_Accounting_Software"] ||
                configValues["Accounting_Software"] ||
                configValues[PTO_CONFIG_FIELDS.accountingSoftware];
        } catch (error) {
            console.warn("Error reading accounting software URL:", error);
        }
    }

    if (!accountingUrl) {
        showToast("No accounting software URL configured. Add SS_Accounting_Software to SS_PF_Config.", "info", 5000);
        return;
    }

    if (!accountingUrl.startsWith("http://") && !accountingUrl.startsWith("https://")) {
        accountingUrl = "https://" + accountingUrl;
    }

    window.open(accountingUrl, "_blank");
    showToast("Opening accounting software...", "success", 2000);
}

// =============================================================================
// STEP 4: ARCHIVE & CLEAR
// Matches payroll-recorder pattern but saves employee-level PTO data
// =============================================================================

/**
 * Show confirmation dialog before archive
 */
function showConfirmDialog(message, options = {}) {
    return new Promise((resolve) => {
        const { title = "Confirm", confirmText = "Confirm", cancelText = "Cancel", icon = "📦" } = options;
        
        // Remove any existing dialogs
        document.querySelectorAll(".pf-confirm-dialog").forEach(d => d.remove());
        
        const dialog = document.createElement("div");
        dialog.className = "pf-confirm-dialog";
        dialog.innerHTML = `
            <div class="pf-confirm-backdrop"></div>
            <div class="pf-confirm-content">
                <div class="pf-confirm-icon">${icon}</div>
                <div class="pf-confirm-title">${escapeHtml(title)}</div>
                <div class="pf-confirm-message">${escapeHtml(message).replace(/\n/g, "<br>")}</div>
                <div class="pf-confirm-buttons">
                    <button type="button" class="pf-confirm-btn pf-confirm-btn--cancel">${escapeHtml(cancelText)}</button>
                    <button type="button" class="pf-confirm-btn pf-confirm-btn--confirm">${escapeHtml(confirmText)}</button>
                </div>
            </div>
        `;
        
        // Add styles if not present
        if (!document.getElementById("pf-confirm-dialog-styles")) {
            const style = document.createElement("style");
            style.id = "pf-confirm-dialog-styles";
            style.textContent = `
                .pf-confirm-dialog { position: fixed; top: 0; left: 0; right: 0; bottom: 0; z-index: 10001; display: flex; align-items: center; justify-content: center; }
                .pf-confirm-backdrop { position: absolute; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.7); backdrop-filter: blur(4px); }
                .pf-confirm-content { position: relative; background: linear-gradient(145deg, rgba(30, 30, 50, 1), rgba(20, 20, 35, 1)); border: 1px solid rgba(99, 102, 241, 0.3); border-radius: 16px; padding: 24px; max-width: 320px; text-align: center; box-shadow: 0 24px 64px rgba(0,0,0,0.5); }
                .pf-confirm-icon { font-size: 40px; margin-bottom: 12px; }
                .pf-confirm-title { font-size: 18px; font-weight: 700; color: #fff; margin-bottom: 12px; }
                .pf-confirm-message { font-size: 13px; color: rgba(255,255,255,0.7); line-height: 1.5; margin-bottom: 20px; text-align: left; }
                .pf-confirm-buttons { display: flex; gap: 12px; justify-content: center; }
                .pf-confirm-btn { padding: 10px 20px; border-radius: 8px; font-size: 14px; font-weight: 600; cursor: pointer; border: none; transition: all 0.2s; }
                .pf-confirm-btn--cancel { background: rgba(255,255,255,0.1); color: rgba(255,255,255,0.8); }
                .pf-confirm-btn--cancel:hover { background: rgba(255,255,255,0.15); }
                .pf-confirm-btn--confirm { background: linear-gradient(145deg, #6366f1, #4f46e5); color: white; }
                .pf-confirm-btn--confirm:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(99, 102, 241, 0.4); }
            `;
            document.head.appendChild(style);
        }
        
        document.body.appendChild(dialog);
        
        dialog.querySelector(".pf-confirm-btn--cancel").addEventListener("click", () => {
            dialog.remove();
            resolve(false);
        });
        
        dialog.querySelector(".pf-confirm-btn--confirm").addEventListener("click", () => {
            dialog.remove();
            resolve(true);
        });
        
        dialog.querySelector(".pf-confirm-backdrop").addEventListener("click", () => {
            dialog.remove();
            resolve(false);
        });
    });
}

/**
 * Main archive and reset function
 * 1. Confirm with user
 * 2. Download Excel archive
 * 3. Save to PTO_Archive_Summary
 * 4. Clear working tabs
 * 5. Reset config
 */
async function archiveAndReset() {
    console.log("[Archive] archiveAndReset called");
    
    if (!hasExcelRuntime()) {
        showToast("Excel not available currently", "error");
        return;
    }
    
    // Confirm before proceeding
    const confirmed = await showConfirmDialog(
        "This will:\n\n" +
        "• Download an Excel archive file\n" +
        "• Update PTO_Archive_Summary\n" +
        "• Clear working data from all sheets\n" +
        "• Reset non-permanent notes & config\n\n" +
        "Make sure you've completed all review steps.",
        {
            title: "Archive PTO Run",
            icon: "📦",
            confirmText: "Archive Now",
            cancelText: "Not Yet"
        }
    );
    
    if (!confirmed) {
        console.log("[PTOArchive] User cancelled");
        showToast("Archive cancelled", "info", 2000);
        return;
    }
    
    console.log("[PTOArchive] User confirmed, starting archive process...");
    toggleLoader(true, "Archiving PTO data...");
    
    try {
        // ═══════════════════════════════════════════════════════════════════
        // STEP 1: Create archive Excel file for download
        // ═══════════════════════════════════════════════════════════════════
        console.log("[PTOArchive] Step 1: Creating archive workbook...");
        
        const archiveSuccess = await createPtoArchiveWorkbook();
        if (!archiveSuccess) {
            console.log("[PTOArchive] Archive file creation failed");
            return;
        }
        
        console.log("[PTOArchive] Step 1 complete: Archive file downloaded");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 2: Update PTO_Archive_Summary with current period
        // ═══════════════════════════════════════════════════════════════════
        console.log("[PTOArchive] Step 2: Updating PTO_Archive_Summary...");
        
        // Get the review data from ptoReviewState or read from PTO_Review
        let reviewData = ptoReviewState.reviewData;
        if (!reviewData || reviewData.length === 0) {
            reviewData = await readPtoReviewData();
        }
        
        if (reviewData && reviewData.length > 0) {
            const analysisDate = getConfigValue(PTO_CONFIG_FIELDS.payrollDate) || new Date().toISOString().substring(0, 10);
            const result = await savePtoArchivePeriod(reviewData, analysisDate);
            console.log("[PTOArchive] Archive summary result:", result);
        } else {
            console.warn("[PTOArchive] No review data to archive");
        }
        
        console.log("[PTOArchive] Step 2 complete: Archive summary updated");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 3: Clear working data from PTO sheets
        // ═══════════════════════════════════════════════════════════════════
        console.log("[PTOArchive] Step 3: Clearing working data...");
        
        await clearPtoWorkingData();
        
        console.log("[PTOArchive] Step 3 complete: Working data cleared");
        
        // ═══════════════════════════════════════════════════════════════════
        // STEP 4: Clear non-permanent config values
        // ═══════════════════════════════════════════════════════════════════
        console.log("[PTOArchive] Step 4: Clearing non-permanent config...");
        
        await clearNonPermanentPtoConfig();
        
        console.log("[PTOArchive] Step 4 complete: Config reset");
        
        // ═══════════════════════════════════════════════════════════════════
        // COMPLETE
        // ═══════════════════════════════════════════════════════════════════
        console.log("[PTOArchive] Archive workflow complete!");
        
        // Reset review state
        ptoReviewState.loaded = false;
        ptoReviewState.reviewData = [];
        ptoReviewState.totalCurrentLiability = 0;
        ptoReviewState.totalPriorLiability = 0;
        ptoReviewState.netChange = 0;
        ptoReviewState.employeeCount = 0;
        
        // Reset rate engine state
        rateEngineState.loaded = false;
        rateEngineState.rates.clear();
        
        // Reset journal state
        journalState.validationRun = false;
        journalState.debitTotal = null;
        journalState.creditTotal = null;
        
        // Reload config and re-render
        await loadStepConfig();
        renderApp();
        
        // Show "save complete" prompt
        showSaveCompletePrompt();
        
    } catch (error) {
        console.error("[PTOArchive] Error during archive:", error);
        showToast("Archive Error: " + error.message, "error", 10000);
    } finally {
        toggleLoader(false);
    }
}

/**
 * Create Excel archive file with PTO sheets
 * Matches payroll-recorder archive format with detailed summary sheet
 */
async function createPtoArchiveWorkbook() {
    try {
        const analysisDate = getConfigValue(PTO_CONFIG_FIELDS.payrollDate) || new Date().toISOString().split("T")[0];
        const accountingPeriod = getConfigValue(PTO_CONFIG_FIELDS.accountingPeriod) || "";
        const journalEntryId = getConfigValue(PTO_CONFIG_FIELDS.journalEntryId) || "";
        const filename = `PTO_Archive_${analysisDate}.xlsx`;
        
        console.log("[PTOArchive] Creating Excel archive file...");
        
        return await Excel.run(async (context) => {
            const workbook = context.workbook;
            const sourceSheets = workbook.worksheets;
            sourceSheets.load("items/name");
            await context.sync();
            
            // Sheets to archive (in order they'll appear in the file)
            const sheetsToArchive = [
                "PTO_JE_Draft",
                "PTO_Review",
                "PTO_Data_Clean"
            ];
            
            // Create new workbook using SheetJS
            const newWorkbook = XLSX.utils.book_new();
            let sheetsAdded = 0;
            
            // ═══════════════════════════════════════════════════════════════════
            // Summary sheet (matches payroll-recorder format)
            // ═══════════════════════════════════════════════════════════════════
            const summaryRows = [];
            summaryRows.push(["PTO Accrual Archive Summary"]);
            summaryRows.push([]);
            summaryRows.push(["Archived At", new Date().toISOString()]);
            summaryRows.push(["Analysis Date", analysisDate]);
            summaryRows.push(["Accounting Period", accountingPeriod]);
            summaryRows.push(["Journal Entry ID", journalEntryId]);
            summaryRows.push(["Company", installationState.ss_company_name || ""]);
            summaryRows.push([]);
            
            // Liability summary
            summaryRows.push(["--- Liability Summary ---"]);
            summaryRows.push(["Current Period Liability", ptoReviewState.totalCurrentLiability || 0]);
            summaryRows.push(["Prior Period Liability", ptoReviewState.totalPriorLiability || 0]);
            summaryRows.push(["Net Change (JE Amount)", ptoReviewState.netChange || 0]);
            summaryRows.push(["Employee Count", ptoReviewState.employeeCount || 0]);
            summaryRows.push([]);
            
            // Sign-off information for each step (matching payroll-recorder)
            summaryRows.push(["--- Step Sign-offs ---"]);
            
            // Step 0: Configuration
            const configFields = getStepConfig(0);
            summaryRows.push(["Config Reviewer", configFields?.reviewer || ""]);
            summaryRows.push(["Config Sign-off", configFields?.signOffDate || ""]);
            summaryRows.push(["Config Notes", configFields?.notes || ""]);
            summaryRows.push([]);
            
            // Step 1: Import
            const importFields = getStepConfig(1);
            summaryRows.push(["Import Reviewer", importFields?.reviewer || ""]);
            summaryRows.push(["Import Sign-off", importFields?.signOffDate || ""]);
            summaryRows.push(["Import Notes", importFields?.notes || ""]);
            summaryRows.push([]);
            
            // Step 2: Review
            const reviewFields = getStepConfig(2);
            summaryRows.push(["Review Reviewer", reviewFields?.reviewer || ""]);
            summaryRows.push(["Review Sign-off", reviewFields?.signOffDate || ""]);
            summaryRows.push(["Review Notes", reviewFields?.notes || ""]);
            summaryRows.push([]);
            
            // Step 3: Journal Entry
            const jeFields = getStepConfig(3);
            summaryRows.push(["JE Reviewer", jeFields?.reviewer || ""]);
            summaryRows.push(["JE Sign-off", jeFields?.signOffDate || ""]);
            summaryRows.push(["JE Notes", jeFields?.notes || ""]);
            summaryRows.push([]);
            
            // Metadata
            summaryRows.push(["--- Archive Metadata ---"]);
            summaryRows.push(["Generated By", "PTO Accrual Module"]);
            summaryRows.push(["Archive File", filename]);
            
            const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows);
            
            // Format summary sheet with proper column widths
            setXlsxColumnWidths(summarySheet, [
                XLSX_COLUMN_WIDTHS.extraWide,  // Label column
                XLSX_COLUMN_WIDTHS.description  // Value column
            ]);
            
            XLSX.utils.book_append_sheet(newWorkbook, summarySheet, "Archive_Summary");
            sheetsAdded++;
            
            // ═══════════════════════════════════════════════════════════════════
            // Copy data sheets
            // ═══════════════════════════════════════════════════════════════════
            for (const sheetName of sheetsToArchive) {
                const sourceSheet = sourceSheets.items.find(s => s.name === sheetName);
                if (!sourceSheet) {
                    console.log(`[PTOArchive] Sheet not found: ${sheetName}`);
                    continue;
                }
                
                const usedRange = sourceSheet.getUsedRangeOrNullObject();
                usedRange.load("values");
                await context.sync();
                
                if (!usedRange.isNullObject && usedRange.values && usedRange.values.length > 0) {
                    const worksheet = XLSX.utils.aoa_to_sheet(usedRange.values);
                    
                    // Apply formatting: headers, column widths, number formats
                    if (usedRange.values.length > 0) {
                        const headers = usedRange.values[0];
                        const rowCount = usedRange.values.length;
                        formatXlsxWorksheet(worksheet, headers, rowCount, {
                            autoFormat: true,
                            autoSize: true
                        });
                    }
                    
                    XLSX.utils.book_append_sheet(newWorkbook, worksheet, sheetName);
                    sheetsAdded++;
                    console.log(`[PTOArchive] Added sheet: ${sheetName} (${usedRange.values.length} rows)`);
                } else {
                    console.log(`[PTOArchive] Sheet empty: ${sheetName}`);
                }
            }
            
            if (sheetsAdded <= 1) { // Only summary sheet
                console.warn("[PTOArchive] No data sheets to archive");
                showToast("No data to archive. Complete the PTO review first.", "warning");
                return false;
            }
            
            // ═══════════════════════════════════════════════════════════════════
            // Download the file
            // ═══════════════════════════════════════════════════════════════════
            const excelBuffer = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });
            const blob = new Blob([excelBuffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
            
            // Trigger download
            const url = URL.createObjectURL(blob);
            const link = document.createElement("a");
            link.setAttribute("href", url);
            link.setAttribute("download", filename);
            link.style.visibility = "hidden";
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            setTimeout(() => URL.revokeObjectURL(url), 100);
            
            console.log(`[PTOArchive] Downloaded: ${filename} with ${sheetsAdded} sheets`);
            showToast(`📥 Archive downloaded: ${filename} (${sheetsAdded} sheets)`, "success", 5000);
            
            return true;
        });
        
    } catch (error) {
        console.error("[PTOArchive] Error creating archive:", error);
        showToast("Archive Export Error: " + error.message, "error", 8000);
        
        // Ask if user wants to continue anyway (matches payroll-recorder behavior)
        return await showConfirmDialog(
            "Archive download failed.\n\n" +
            "Do you want to continue with clearing the data?\n\n" +
            "Make sure you have saved a backup first!",
            {
                title: "Continue Without Archive?",
                icon: "⚠️",
                confirmText: "Continue Anyway",
                cancelText: "Cancel"
            }
        );
    }
}

/**
 * Read PTO_Review data if not in state
 */
async function readPtoReviewData() {
    const reviewData = [];
    
    try {
        await Excel.run(async (context) => {
            const reviewSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Review");
            reviewSheet.load("isNullObject");
            await context.sync();
            
            if (reviewSheet.isNullObject) return;
            
            const reviewRange = reviewSheet.getUsedRangeOrNullObject();
            reviewRange.load("values");
            await context.sync();
            
            if (reviewRange.isNullObject || !reviewRange.values || reviewRange.values.length < 2) return;
            
            const headers = reviewRange.values[0].map(h => String(h || "").toLowerCase().trim());
            
            // Find column indices
            const colIdx = {
                employeeName: headers.findIndex(h => h.includes("employee")),
                department: headers.findIndex(h => h === "department"),
                payRate: headers.findIndex(h => h.includes("pay") && h.includes("rate")),
                accrualRate: headers.findIndex(h => h.includes("accrual") && h.includes("rate")),
                carryOver: headers.findIndex(h => h.includes("carry")),
                ytdAccrued: headers.findIndex(h => h.includes("ytd") && h.includes("accrued")),
                ytdUsed: headers.findIndex(h => h.includes("ytd") && h.includes("used")),
                balance: headers.findIndex(h => h === "balance"),
                liabilityAmount: headers.findIndex(h => h.includes("liability")),
                priorLiability: headers.findIndex(h => h.includes("prior")),
                change: headers.findIndex(h => h === "change")
            };
            
            for (let i = 1; i < reviewRange.values.length; i++) {
                const row = reviewRange.values[i];
                reviewData.push({
                    employeeName: colIdx.employeeName >= 0 ? String(row[colIdx.employeeName] || "") : "",
                    department: colIdx.department >= 0 ? String(row[colIdx.department] || "") : "",
                    payRate: colIdx.payRate >= 0 ? parseFloat(row[colIdx.payRate]) || 0 : 0,
                    accrualRate: colIdx.accrualRate >= 0 ? parseFloat(row[colIdx.accrualRate]) || 0 : 0,
                    carryOver: colIdx.carryOver >= 0 ? parseFloat(row[colIdx.carryOver]) || 0 : 0,
                    ytdAccrued: colIdx.ytdAccrued >= 0 ? parseFloat(row[colIdx.ytdAccrued]) || 0 : 0,
                    ytdUsed: colIdx.ytdUsed >= 0 ? parseFloat(row[colIdx.ytdUsed]) || 0 : 0,
                    balance: colIdx.balance >= 0 ? parseFloat(row[colIdx.balance]) || 0 : 0,
                    liabilityAmount: colIdx.liabilityAmount >= 0 ? parseFloat(row[colIdx.liabilityAmount]) || 0 : 0
                });
            }
            
            await context.sync();
        });
    } catch (error) {
        console.error("[PTOArchive] Error reading PTO_Review:", error);
    }
    
    return reviewData;
}

/**
 * Clear working data from PTO sheets (keep headers)
 */
async function clearPtoWorkingData() {
    const sheetsToClear = [
        "PTO_Data_Clean",
        "PTO_Review",
        "PTO_JE_Draft"
    ];
    
    await Excel.run(async (context) => {
        for (const sheetName of sheetsToClear) {
            const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            
            if (sheet.isNullObject) {
                console.log(`[PTOArchive] Sheet not found: ${sheetName}`);
                continue;
            }
            
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount,columnCount,address");
            await context.sync();
            
            if (usedRange.isNullObject || usedRange.rowCount <= 1) {
                console.log(`[PTOArchive] Sheet empty or headers only: ${sheetName}`);
                continue;
            }
            
            // Clear data rows (row 2 onwards), keep headers (row 1)
            const dataRange = sheet.getRange(`A2:${String.fromCharCode(64 + Math.min(usedRange.columnCount, 26))}${usedRange.rowCount}`);
            dataRange.clear(Excel.ClearApplyTo.contents);
            
            await context.sync();
            console.log(`[PTOArchive] Cleared data from: ${sheetName}`);
        }
    });
}

/**
 * Clear non-permanent config values from SS_PF_Config
 */
async function clearNonPermanentPtoConfig() {
    await Excel.run(async (context) => {
        const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
        configSheet.load("isNullObject");
        await context.sync();
        
        if (configSheet.isNullObject) {
            console.warn("[PTOArchive] SS_PF_Config sheet not found");
            return;
        }
        
        const usedRange = configSheet.getUsedRangeOrNullObject();
        usedRange.load("values,rowCount,columnCount");
        await context.sync();
        
        if (usedRange.isNullObject || !usedRange.values || usedRange.values.length < 2) {
            return;
        }
        
        const headers = usedRange.values[0].map(h => String(h || "").toLowerCase().trim());
        const fieldIdx = headers.findIndex(h => h === "field" || h === "setting" || h === "key");
        const valueIdx = headers.findIndex(h => h === "value");
        const permanentIdx = headers.findIndex(h => h === "permanent" || h === "persist");
        
        if (fieldIdx < 0 || valueIdx < 0) {
            console.warn("[PTOArchive] Could not find field/value columns in config");
            return;
        }
        
        // Fields to clear (non-permanent PTO fields)
        const fieldsToClear = [
            PTO_CONFIG_FIELDS.payrollDate,
            PTO_CONFIG_FIELDS.accountingPeriod,
            PTO_CONFIG_FIELDS.journalEntryId
        ];
        
        const updatedValues = usedRange.values.map((row, rowIndex) => {
            if (rowIndex === 0) return row; // Skip header
            
            const field = String(row[fieldIdx] || "").trim();
            const isPermanent = permanentIdx >= 0 && String(row[permanentIdx] || "").toUpperCase() === "Y";
            
            // Clear if it's a PTO field and not permanent
            if (fieldsToClear.includes(field) && !isPermanent) {
                const newRow = [...row];
                newRow[valueIdx] = "";
                return newRow;
            }
            
            return row;
        });
        
        // Write back
        const writeRange = configSheet.getRangeByIndexes(0, 0, updatedValues.length, headers.length);
        writeRange.values = updatedValues;
        
        await context.sync();
        console.log("[PTOArchive] Non-permanent config values cleared");
    });
}

/**
 * Show a simple finalize prompt centered in the side panel
 * Appears after file download - stays until user clicks
 */
function showSaveCompletePrompt() {
    // Remove any existing prompts
    document.querySelectorAll(".pf-save-prompt").forEach(p => p.remove());
    
    const prompt = document.createElement("div");
    prompt.className = "pf-save-prompt";
    prompt.innerHTML = `
        <div class="pf-save-prompt-content">
            <div class="pf-save-prompt-title">Good work!</div>
            <div class="pf-save-prompt-subtitle">Ready to finalize?</div>
            <button type="button" class="pf-save-prompt-btn">
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <polyline points="20 6 9 17 4 12"/>
                </svg>
                Finalize
            </button>
        </div>
    `;
    
    // Add styles
    if (!document.getElementById("pf-save-prompt-styles")) {
        const style = document.createElement("style");
        style.id = "pf-save-prompt-styles";
        style.textContent = `
            .pf-save-prompt {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: linear-gradient(145deg, rgba(30, 30, 50, 0.98), rgba(20, 20, 35, 0.99));
                border: 1px solid rgba(99, 102, 241, 0.3);
                color: white;
                padding: 32px 40px;
                border-radius: 20px;
                box-shadow: 
                    0 24px 64px rgba(0, 0, 0, 0.5),
                    0 0 0 1px rgba(99, 102, 241, 0.2) inset,
                    0 0 60px rgba(99, 102, 241, 0.1);
                z-index: 10002;
                text-align: center;
                animation: pf-prompt-fade-in 0.3s ease;
            }
            @keyframes pf-prompt-fade-in {
                from { opacity: 0; transform: translate(-50%, -50%) scale(0.95); }
                to { opacity: 1; transform: translate(-50%, -50%) scale(1); }
            }
            .pf-save-prompt-content {
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 8px;
            }
            .pf-save-prompt-title {
                font-size: 20px;
                font-weight: 700;
                color: #fff;
            }
            .pf-save-prompt-subtitle {
                font-size: 14px;
                color: rgba(255, 255, 255, 0.6);
                margin-bottom: 12px;
            }
            .pf-save-prompt-btn {
                background: linear-gradient(145deg, #6366f1, #4f46e5);
                border: none;
                color: white;
                padding: 16px 40px;
                border-radius: 12px;
                font-size: 18px;
                font-weight: 700;
                cursor: pointer;
                display: flex;
                align-items: center;
                justify-content: center;
                gap: 10px;
                transition: all 0.2s ease;
                box-shadow: 0 8px 24px rgba(99, 102, 241, 0.4), 0 0 0 2px rgba(99, 102, 241, 0.2);
                margin-top: 8px;
            }
            .pf-save-prompt-btn:hover {
                transform: translateY(-3px);
                box-shadow: 0 12px 32px rgba(99, 102, 241, 0.5), 0 0 0 2px rgba(99, 102, 241, 0.3);
                background: linear-gradient(145deg, #7c3aed, #6366f1);
            }
            .pf-save-prompt-btn:active {
                transform: translateY(-1px);
            }
            .pf-save-prompt.closing {
                animation: pf-prompt-fade-out 0.2s ease forwards;
            }
            @keyframes pf-prompt-fade-out {
                to { opacity: 0; transform: translate(-50%, -50%) scale(0.95); }
            }
        `;
        document.head.appendChild(style);
    }
    
    document.body.appendChild(prompt);
    
    // Handle "Finalize" button click
    const doneBtn = prompt.querySelector(".pf-save-prompt-btn");
    doneBtn.addEventListener("click", () => {
        prompt.classList.add("closing");
        setTimeout(() => {
            prompt.remove();
            // Now show the celebratory toast
            showArchiveSuccessToast();
        }, 200);
    });
}

/**
 * Show celebratory archive success toast with rotating messages
 * Displays for 5 seconds then redirects to Module Selector
 */
function showArchiveSuccessToast() {
    // Remove existing toasts
    document.querySelectorAll(".pf-toast, .pf-success-toast").forEach(t => t.remove());
    
    // Lucide icon SVGs with Prairie Forge purple
    const messages = [
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z"/></svg>`,
            title: "Done.",
            subtitle: "Efficiency unlocked."
        },
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.5 19H9a7 7 0 1 1 6.71-9h1.79a4.5 4.5 0 1 1 0 9Z"/><path d="m9 12 2 2 4-4"/></svg>`,
            title: "Locked in.",
            subtitle: "You're good to go."
        },
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="1"/><path d="M20.2 20.2c2.04-2.03.02-7.36-4.5-11.9-4.54-4.52-9.87-6.54-11.9-4.5-2.04 2.03-.02 7.36 4.5 11.9 4.54 4.52 9.87 6.54 11.9 4.5Z"/><path d="M15.7 15.7c4.52-4.54 6.54-9.87 4.5-11.9-2.03-2.04-7.36-.02-11.9 4.5-4.52 4.54-6.54 9.87-4.5 11.9 2.03 2.04 7.36.02 11.9-4.5Z"/></svg>`,
            title: "Stored.",
            subtitle: "Everything stays aligned."
        },
        {
            icon: `<svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/></svg>`,
            title: "PTO archived.",
            subtitle: "Flow restored."
        }
    ];
    
    // Pick a random message
    const msg = messages[Math.floor(Math.random() * messages.length)];
    
    const toast = document.createElement("div");
    toast.className = "pf-success-toast";
    toast.innerHTML = `
        <div class="pf-success-toast-icon">${msg.icon}</div>
        <div class="pf-success-toast-text">
            <div class="pf-success-toast-title">${msg.title}</div>
            <div class="pf-success-toast-subtitle">${msg.subtitle}</div>
        </div>
        <div class="pf-success-toast-progress"></div>
    `;
    
    // Add styles if not already present
    if (!document.getElementById("pf-success-toast-styles")) {
        const style = document.createElement("style");
        style.id = "pf-success-toast-styles";
        style.textContent = `
            .pf-success-toast {
                position: fixed;
                top: 50%;
                left: 50%;
                transform: translate(-50%, -50%);
                background: linear-gradient(145deg, rgba(30, 30, 50, 0.98), rgba(20, 20, 35, 0.99));
                border: 1px solid rgba(99, 102, 241, 0.3);
                color: white;
                padding: 36px 48px 24px;
                border-radius: 24px;
                box-shadow: 
                    0 32px 64px rgba(0, 0, 0, 0.5),
                    0 0 0 1px rgba(99, 102, 241, 0.2) inset,
                    0 0 80px rgba(99, 102, 241, 0.15);
                z-index: 10002;
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 16px;
                animation: pf-success-in 0.4s cubic-bezier(0.34, 1.56, 0.64, 1);
                overflow: hidden;
                text-align: center;
                min-width: 220px;
            }
            @keyframes pf-success-in {
                from { 
                    opacity: 0; 
                    transform: translate(-50%, -50%) scale(0.9);
                }
                to { 
                    opacity: 1; 
                    transform: translate(-50%, -50%) scale(1);
                }
            }
            @keyframes pf-success-out {
                from { 
                    opacity: 1; 
                    transform: translate(-50%, -50%) scale(1);
                }
                to { 
                    opacity: 0; 
                    transform: translate(-50%, -50%) scale(0.95) translateY(-20px);
                }
            }
            .pf-success-toast-icon {
                width: 64px;
                height: 64px;
                background: linear-gradient(145deg, #6366f1, #4f46e5);
                border-radius: 18px;
                display: flex;
                align-items: center;
                justify-content: center;
                color: white;
                box-shadow: 0 8px 24px rgba(99, 102, 241, 0.4);
                animation: pf-icon-pulse 2s ease-in-out infinite;
            }
            @keyframes pf-icon-pulse {
                0%, 100% { box-shadow: 0 8px 24px rgba(99, 102, 241, 0.4); }
                50% { box-shadow: 0 8px 32px rgba(99, 102, 241, 0.6); }
            }
            .pf-success-toast-text {
                display: flex;
                flex-direction: column;
                align-items: center;
                gap: 4px;
            }
            .pf-success-toast-title {
                font-size: 22px;
                font-weight: 700;
                color: #fff;
                letter-spacing: -0.5px;
            }
            .pf-success-toast-subtitle {
                font-size: 15px;
                color: rgba(255, 255, 255, 0.6);
                font-weight: 500;
            }
            .pf-success-toast-progress {
                position: absolute;
                bottom: 0;
                left: 0;
                height: 4px;
                background: linear-gradient(90deg, #6366f1, #a855f7, #6366f1);
                background-size: 200% 100%;
                animation: pf-progress-shrink 5s linear forwards, pf-progress-shimmer 1s linear infinite;
                border-radius: 0 0 24px 24px;
            }
            @keyframes pf-progress-shrink {
                from { width: 100%; }
                to { width: 0%; }
            }
            @keyframes pf-progress-shimmer {
                from { background-position: 200% 0; }
                to { background-position: -200% 0; }
            }
            .pf-success-toast.closing {
                animation: pf-success-out 0.3s ease forwards;
            }
            .pf-success-backdrop {
                position: fixed;
                inset: 0;
                background: rgba(0, 0, 0, 0.4);
                backdrop-filter: blur(4px);
                -webkit-backdrop-filter: blur(4px);
                z-index: 10001;
                animation: pf-confirm-fade-in 0.2s ease;
            }
        `;
        document.head.appendChild(style);
    }
    
    // Add backdrop
    const backdrop = document.createElement("div");
    backdrop.className = "pf-success-backdrop";
    document.body.appendChild(backdrop);
    document.body.appendChild(toast);
    
    // After 5 seconds, close and redirect to Module Selector
    setTimeout(() => {
        toast.classList.add("closing");
        backdrop.style.opacity = "0";
        backdrop.style.transition = "opacity 0.3s ease";
        
        setTimeout(() => {
            toast.remove();
            backdrop.remove();
            // Navigate to Module Selector homepage
            navigateToModuleSelector();
        }, 300);
    }, 5000);
}

/**
 * Return to the module homepage
 */
async function returnHome() {
    const homepageConfig = getHomepageConfig(MODULE_KEY);
    await activateHomepageSheet(homepageConfig.sheetName, homepageConfig.title, homepageConfig.subtitle);
    setState({ activeView: "home", activeStepId: null });
}

/**
 * Navigate to the Module Selector (used after archive completes)
 */
async function navigateToModuleSelector() {
    // Reset tab visibility before redirect to avoid Mac Excel tab accumulation
    try {
        await applyModuleTabVisibility("module-selector");
    } catch (e) {
        console.warn("[PTO] Could not apply module-selector tab visibility before redirect:", e);
    }
    // Use relative path from current location (pto-accrual/ -> module-selector/)
    window.location.href = "../module-selector/index.html";
}

async function openSheet(sheetName) {
    if (!sheetName || !hasExcel()) {
        return;
    }
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(sheetName);
            sheet.activate();
            // Select cell A1 so user always starts at a consistent location
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error(error);
    }
}

/**
 * Clear PTO_Data_Clean sheet to start fresh
 */
async function clearPtoData() {
    if (!hasExcel()) return;
    
    const confirmed = await showConfirm(
        "All data in PTO_Data_Clean will be permanently removed.\n\nThis action cannot be undone.",
        {
            title: "Clear PTO Data",
            icon: "🗑️",
            confirmText: "Clear Data",
            cancelText: "Keep Data",
            destructive: true
        }
    );
    if (!confirmed) return;
    
    toggleLoader(true);
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
            sheet.load("isNullObject");
            await context.sync();
            
            if (sheet.isNullObject) {
                showToast("PTO_Data_Clean not found.", "info");
                return;
            }
            
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount");
            await context.sync();
            
            if (!usedRange.isNullObject && usedRange.rowCount > 1) {
                // Keep header row, clear everything else
                const dataRange = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, 20);
                dataRange.clear(Excel.ClearApplyTo.contents);
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
        
        // Also reset upload state
        ptoUploadState.file = null;
        ptoUploadState.fileName = "";
        ptoUploadState.headers = [];
        ptoUploadState.rowCount = 0;
        ptoUploadState.parsedData = null;
        ptoUploadState.error = null;
        
        showToast("PTO data cleared successfully. You can now upload new data.", "success");
        renderApp();
    } catch (error) {
        console.error("Clear PTO data error:", error);
        showToast(`Failed to clear PTO data: ${error.message}`, "error");
    } finally {
        toggleLoader(false);
    }
}

/**
 * Open a reference data sheet (creates if doesn't exist)
 */
async function openReferenceSheet(sheetName) {
    if (!sheetName || !hasExcel()) {
        return;
    }
    
    const defaultHeaders = {
        "SS_Employee_Roster": ["Employee", "Department", "Pay_Rate", "Status", "Hire_Date"],
        "SS_Chart_of_Accounts": ["Account_Number", "Account_Name", "Type", "Category"]
    };
    
    try {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject");
            await context.sync();
            
            if (sheet.isNullObject) {
                // Create the sheet with default headers
                sheet = context.workbook.worksheets.add(sheetName);
                const headers = defaultHeaders[sheetName] || ["Column1", "Column2"];
                const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`);
                headerRange.values = [headers];
                headerRange.format.font.bold = true;
                headerRange.format.fill.color = "#f0f0f0";
                headerRange.format.autofitColumns();
                await context.sync();
            }
            
            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
        });
    } catch (error) {
        console.error("Error opening reference sheet:", error);
    }
}

/**
 * Fetch configuration sheets (SS_* and any with "mapping" in name)
 */
async function getConfigurationSheets() {
    if (!hasExcel()) return [];
    try {
        return await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load("items/name,visibility");
            await context.sync();
            const matches = worksheets.items.filter((sheet) => {
                const name = sheet.name || "";
                const upper = name.toUpperCase();
                return upper.startsWith("SS_") || upper.includes("MAPPING") || upper.includes("HOMEPAGE");
            });
            return matches
                .map((sheet) => ({
                    name: sheet.name,
                    visible: sheet.visibility === Excel.SheetVisibility.visible,
                    isHomepage: (sheet.name || "").toUpperCase().includes("HOMEPAGE")
                }))
                .sort((a, b) => {
                    if (a.isHomepage && !b.isHomepage) return 1;
                    if (!a.isHomepage && b.isHomepage) return -1;
                    return a.name.localeCompare(b.name);
                });
        });
    } catch (error) {
        console.error("[Config] Error reading configuration sheets:", error);
        return [];
    }
}

function ensureConfigModal() {
    if (document.getElementById("config-sheet-modal")) return;
    const modal = document.createElement("div");
    modal.id = "config-sheet-modal";
    modal.className = "pf-config-modal hidden";
    modal.innerHTML = `
        <div class="pf-config-modal-backdrop" data-close></div>
        <div class="pf-config-modal-card">
            <div class="pf-config-modal-head">
                <h3>Configuration Sheets</h3>
                <button type="button" class="pf-config-close" data-close aria-label="Close">×</button>
            </div>
            <div class="pf-config-modal-body">
                <p class="pf-config-hint">Choose a configuration or mapping sheet to unhide and open.</p>
                <div id="config-sheet-list" class="pf-config-sheet-list">Loading…</div>
            </div>
        </div>
    `;
    document.body.appendChild(modal);

    if (!document.getElementById("pf-config-modal-styles")) {
        const style = document.createElement("style");
        style.id = "pf-config-modal-styles";
        style.textContent = `
            .pf-config-modal { position: fixed; inset: 0; display: flex; align-items: center; justify-content: center; z-index: 10000; }
            .pf-config-modal.hidden { display: none; }
            .pf-config-modal-backdrop { position: absolute; inset: 0; background: rgba(0,0,0,0.6); }
            .pf-config-modal-card { position: relative; background: #0f172a; color: #e2e8f0; border-radius: 12px; padding: 20px; width: min(420px, 90%); box-shadow: 0 20px 60px rgba(0,0,0,0.35); }
            .pf-config-modal-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 12px; }
            .pf-config-close { background: transparent; border: none; color: #f8fafc; font-size: 20px; cursor: pointer; }
            .pf-config-hint { margin: 0 0 12px 0; color: #cbd5e1; font-size: 14px; }
            .pf-config-sheet-list { display: flex; flex-direction: column; gap: 10px; max-height: 260px; overflow-y: auto; }
            .pf-config-sheet { display: flex; justify-content: space-between; align-items: center; padding: 12px 14px; background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.18); border-radius: 10px; cursor: pointer; color: #e2e8f0; font-weight: 600; }
            .pf-config-sheet:hover { background: rgba(255,255,255,0.16); }
            .pf-config-pill { font-size: 12px; color: #c7d2fe; }
        `;
        document.head.appendChild(style);
    }
}

async function openConfigModal() {
    ensureConfigModal();
    const modal = document.getElementById("config-sheet-modal");
    const list = document.getElementById("config-sheet-list");
    if (!modal || !list) return;

    list.textContent = "Loading…";
    modal.classList.remove("hidden");

    const sheets = await getConfigurationSheets();
    if (!sheets.length) {
        list.textContent = "No configuration sheets found.";
    } else {
        list.innerHTML = "";
        sheets.forEach((sheet) => {
            const btn = document.createElement("button");
            btn.type = "button";
            btn.className = "pf-config-sheet";
            btn.innerHTML = `<span>${sheet.name}</span><span class="pf-config-pill">${sheet.visible ? "Visible" : "Hidden"}</span>`;
            btn.addEventListener("click", async () => {
                await openConfigSheet(sheet.name);
                modal.classList.add("hidden");
            });
            list.appendChild(btn);
        });
    }

    modal.querySelectorAll("[data-close]").forEach((el) =>
        el.addEventListener("click", () => modal.classList.add("hidden"))
    );
}

async function openConfigSheet(sheetName) {
    if (!sheetName || !hasExcel()) return;
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            let sheet = worksheets.getItemOrNullObject(sheetName);
            sheet.load("isNullObject,visibility");
            await context.sync();

            if (sheet.isNullObject) {
                sheet = worksheets.add(sheetName);
            }
            sheet.visibility = Excel.SheetVisibility.visible;
            await context.sync();

            sheet.activate();
            sheet.getRange("A1").select();
            await context.sync();
            console.log(`[Config] Opened sheet: ${sheetName}`);
        });
    } catch (error) {
        console.error("[Config] Error opening sheet", sheetName, error);
    }
}

async function writeDatasetToSheet(context, sheetName, columns, rows) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("address");
    await context.sync();
    if (!usedRange.isNullObject) {
        usedRange.clear();
    }
    const data = [
        columns.map((col) => col.header),
        ...rows.map((row) => columns.map((col) => row[col.key]))
    ];
    const range = sheet.getRangeByIndexes(0, 0, data.length, data[0]?.length || 1);
    range.values = data;
    range.format.autofitColumns();
    await context.sync();
}

async function clearSheetBelowHeader(context, sheetName) {
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("rowCount");
    await context.sync();
    if (usedRange.isNullObject || usedRange.rowCount <= 1) return;
    const dataRange = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, usedRange.columnCount);
    dataRange.clear();
}

async function getAnalysisRows(context) {
    const sheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
    sheet.load("isNullObject");
    await context.sync();
    if (sheet.isNullObject) return [];
    const range = sheet.getUsedRangeOrNullObject();
    range.load("values");
    await context.sync();
    const values = range.isNullObject ? [] : range.values || [];
    if (values.length <= 1) return [];
    return values.slice(1);
}

async function writeRowsStartingAt(context, sheetName, rows) {
    if (!rows.length) return;
    const sheet = context.workbook.worksheets.getItem(sheetName);
    const range = sheet.getRangeByIndexes(1, 0, rows.length, rows[0].length);
    range.values = rows;
    range.format.autofitColumns();
}

async function clearNonPermanentConfig(context) {
    const table = context.workbook.tables.getItemOrNullObject(CONFIG_TABLES[0]);
    await context.sync();
    if (table.isNullObject) return;
    const body = table.getDataBodyRange();
    const header = table.getHeaderRowRange();
    body.load("values");
    header.load("values");
    await context.sync();
    const headers = header.values[0] || [];
    const normalizedHeaders = headers.map((h) => normalizeName(h));
    const idx = {
        field: normalizedHeaders.findIndex((h) => h === "field" || h === "field name" || h === "setting"),
        permanent: normalizedHeaders.findIndex((h) => h === "permanent" || h === "persist"),
        value: normalizedHeaders.findIndex((h) => h === "value" || h === "setting value")
    };
    if (idx.field === -1 || idx.value === -1 || idx.permanent === -1) return;
    (body.values || []).forEach((row, rowIndex) => {
        const permanent = String(row[idx.permanent] ?? "").trim().toLowerCase();
        const shouldClear = permanent !== "y" && permanent !== "yes" && permanent !== "true" && permanent !== "t" && permanent !== "1";
        if (shouldClear) {
            body.getCell(rowIndex, idx.value).values = [[""]];
        }
    });
}

async function exportPtoSheets(context) {
    const sheetNames = await getPtoSheetNames(context);
    if (!sheetNames.length) return;
    const chunks = [];
    for (const name of sheetNames) {
        try {
            const sheet = context.workbook.worksheets.getItemOrNullObject(name);
            sheet.load("isNullObject");
            await context.sync();
            if (sheet.isNullObject) continue;
            const range = sheet.getUsedRangeOrNullObject();
            range.load("values");
            await context.sync();
            const values = range.isNullObject ? [] : range.values || [];
            const csv = values
                .map((row) => row.map((cell) => `"${String(cell ?? "").replace(/"/g, '""')}"`).join(","))
                .join("\n");
            chunks.push(`# Sheet: ${name}\n${csv}`);
        } catch (error) {
            console.warn("PTO: unable to export sheet", name, error);
        }
    }
    if (!chunks.length) return;
    const fileName = `${new Date().toISOString().slice(0, 10)} - Tai Software PTO Accrual.xlsx`;
    // rudimentary multi-tab export: CSV sections inside a .xlsx extension to break links
    const blob = new Blob([chunks.join("\n\n")], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

async function getPtoSheetNames(context) {
    const configSheet = context.workbook.worksheets.getItemOrNullObject("SS_PF_Config");
    configSheet.load("isNullObject");
    await context.sync();
    if (configSheet.isNullObject) return [];
    const usedRange = configSheet.getUsedRangeOrNullObject();
    usedRange.load("values");
    await context.sync();
    const values = usedRange.isNullObject ? [] : usedRange.values || [];
    if (!values.length) return [];
    const headers = (values[0] || []).map((h) => normalizeName(h));
    const idx = {
        category: headers.findIndex((h) => h === "category"),
        module: headers.findIndex((h) => h === "module"),
        field: headers.findIndex((h) => h === "field"),
        value: headers.findIndex((h) => h === "value")
    };
    if (idx.category === -1 || idx.field === -1) return [];
    const targetModule = normalizeName(MODULE_NAME);
    const names = values
        .slice(1)
        .filter((row) => {
            const category = normalizeName(row[idx.category]);
            const module = idx.module >= 0 ? normalizeName(row[idx.module]) : "";
            const moduleValue = idx.value >= 0 ? normalizeName(row[idx.value]) : "";
            return category === "tab-structure" && (module === targetModule || moduleValue === "pto-accrual");
        })
        .map((row) => String(row[idx.field] ?? "").trim())
        .filter(Boolean);
    return Array.from(new Set(names));
}

/**
 * Copy PTO_Analysis to PTO_Archive_Summary.
 * Only retains the MOST RECENT period - older data is replaced.
 * e.g., When archiving 11/30, this overwrites any existing 10/31 data.
 */
async function copyAnalysisToArchiveSummary(context, analysisRows = null) {
    const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
    const archiveSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Archive_Summary");
    analysisSheet.load("isNullObject");
    archiveSheet.load("isNullObject");
    await context.sync();
    if (analysisSheet.isNullObject || archiveSheet.isNullObject) return;
    
    // Get analysis data (use passed rows or fetch fresh)
    let values = analysisRows;
    if (!values || !values.length) {
        const analysisRange = analysisSheet.getUsedRangeOrNullObject();
        analysisRange.load("values");
        await context.sync();
        values = analysisRange.isNullObject ? [] : analysisRange.values || [];
    }
    if (!values.length) return;
    
    // Clear existing archive data (only retain most recent period)
    const existingRange = archiveSheet.getUsedRangeOrNullObject();
    existingRange.load("isNullObject");
    await context.sync();
    if (!existingRange.isNullObject) {
        existingRange.clear();
    }
    
    // Write current period's analysis as the new archive
    const archiveRange = archiveSheet.getRangeByIndexes(0, 0, values.length, values[0].length);
    archiveRange.values = values;
    archiveRange.format.autofitColumns();
    
    // Reset selection to A1
    archiveSheet.getRange("A1").select();
    await context.sync();
}

function getConfigValue(fieldName) {
    const key = String(fieldName ?? "").trim();
    return configState.values?.[key] ?? "";
}

/**
 * Get reviewer name with fallback chain:
 * 1. Step-specific reviewer (from step config fields)
 * 2. Module default (PTO_Reviewer_Name in SS_PF_Config)
 * 3. Shared default (Default_Reviewer in SS_PF_Config)
 */
function getReviewerWithFallback(stepReviewer) {
    // Step-specific
    if (stepReviewer) return stepReviewer;
    
    // Module default
    const moduleDefault = getConfigValue(PTO_CONFIG_FIELDS.reviewerName);
    if (moduleDefault) return moduleDefault;
    
    // Shared default (from cached SS_PF_Config) - check new + legacy names
    if (window.PrairieForge?._sharedConfigCache) {
        const sharedDefault = window.PrairieForge._sharedConfigCache.get("SS_Default_Reviewer") 
            || window.PrairieForge._sharedConfigCache.get("Default_Reviewer");
        if (sharedDefault) return sharedDefault;
    }
    
    return "";
}

function scheduleConfigWrite(fieldName, value, options = {}) {
    const key = String(fieldName ?? "").trim();
    if (!key) return;
    configState.values[key] = value ?? "";
    const delay = options.debounceMs ?? 0;
    if (!delay) {
        const existing = pendingConfigWrites.get(key);
        if (existing) clearTimeout(existing);
        pendingConfigWrites.delete(key);
        void saveConfigValue(key, value ?? "", CONFIG_TABLES);
        return;
    }
    if (pendingConfigWrites.has(key)) {
        clearTimeout(pendingConfigWrites.get(key));
    }
    const timer = setTimeout(() => {
        pendingConfigWrites.delete(key);
        void saveConfigValue(key, value ?? "", CONFIG_TABLES);
    }, delay);
    pendingConfigWrites.set(key, timer);
}

function normalizeName(value) {
    return String(value ?? "").trim().toLowerCase();
}

function columnLetterFromIndex(index) {
    let dividend = index + 1;
    let columnName = "";
    while (dividend > 0) {
        const modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
}

function toggleLoader(show, message = "Working...") {
    // Suppress loader overlay for smoother transitions between views.
    const overlay = document.getElementById(LOADER_ID);
    if (overlay) overlay.style.display = "none";
}

function bootstrapModule() {
    init();
}

if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(() => bootstrapModule()).catch(() => bootstrapModule());
} else {
    bootstrapModule();
}

function getStepConfig(stepId) {
    return configState.steps[stepId] || { notes: "", reviewer: "", signOffDate: "" };
}

function getFieldNames(stepId) {
    return STEP_CONFIG_FIELDS[stepId] || {};
}

function getStepType(stepId) {
    // Matches payroll-recorder 5-step structure
    if (stepId === 0) return "config";
    if (stepId === 1) return "upload";      // Upload & Validate
    if (stepId === 2) return "review";      // PTO Accrual Review
    if (stepId === 3) return "journal";     // Journal Entry Prep
    if (stepId === 4) return "archive";     // Archive & Clear
    return "";
}

async function saveStepField(stepId, key, value) {
    const current = configState.steps[stepId] || { notes: "", reviewer: "", signOffDate: "" };
    current[key] = value;
    configState.steps[stepId] = current;
    const fieldNames = getFieldNames(stepId);
    const targetField =
        key === "notes" ? fieldNames.note : key === "reviewer" ? fieldNames.reviewer : fieldNames.signOff;
    if (!targetField) return;
    if (!hasExcelRuntime()) return;
    try {
        await saveConfigValue(targetField, value, CONFIG_TABLES);
    } catch (error) {
        console.warn("PTO: unable to save field", targetField, error);
    }
}

async function toggleNotePermanent(stepId, isPermanent) {
    configState.permanents[stepId] = isPermanent;
    const fieldNames = getFieldNames(stepId);
    if (!fieldNames?.note) return;
    if (!hasExcelRuntime()) return;
    try {
        await Excel.run(async (context) => {
            const table = context.workbook.tables.getItemOrNullObject(CONFIG_TABLES[0]);
            await context.sync();
            if (table.isNullObject) return;
            const body = table.getDataBodyRange();
            const header = table.getHeaderRowRange();
            body.load("values");
            header.load("values");
            await context.sync();
            const headers = header.values[0] || [];
            const normalizedHeaders = headers.map((h) => String(h || "").trim().toLowerCase());
            const idx = {
                field: normalizedHeaders.findIndex((h) => h === "field" || h === "field name" || h === "setting"),
                permanent: normalizedHeaders.findIndex((h) => h === "permanent" || h === "persist"),
                value: normalizedHeaders.findIndex((h) => h === "value" || h === "setting value"),
                type: normalizedHeaders.findIndex((h) => h === "type" || h === "category"),
                title: normalizedHeaders.findIndex((h) => h === "title" || h === "display name")
            };
            if (idx.field === -1) return;
            const rows = body.values || [];
            const targetIndex = rows.findIndex(
                (row) => String(row[idx.field] || "").trim() === fieldNames.note
            );
            if (targetIndex >= 0) {
                if (idx.permanent >= 0) {
                    body.getCell(targetIndex, idx.permanent).values = [[isPermanent ? "Y" : "N"]];
                }
            } else {
                // create row with permanent flag if missing
                const newRow = new Array(headers.length).fill("");
                if (idx.type >= 0) newRow[idx.type] = "Other";
                if (idx.title >= 0) newRow[idx.title] = "";
                newRow[idx.field] = fieldNames.note;
                if (idx.permanent >= 0) newRow[idx.permanent] = isPermanent ? "Y" : "N";
                if (idx.value >= 0) newRow[idx.value] = configState.steps[stepId]?.notes || "";
                table.rows.add(null, [newRow]);
            }
            await context.sync();
        });
    } catch (error) {
        console.warn("PTO: unable to update permanent flag", error);
    }
}

async function saveCompletionFlag(stepId, isComplete) {
    const fieldName = STEP_COMPLETE_FIELDS[stepId];
    if (!fieldName) return;
    configState.completes[stepId] = isComplete ? "Y" : "";
    if (!hasExcelRuntime()) return;
    try {
        await saveConfigValue(fieldName, isComplete ? "Y" : "", CONFIG_TABLES);
    } catch (error) {
        console.warn("PTO: unable to save completion flag", fieldName, error);
    }
}

function updateActionToggleState(button, isActive) {
    if (!button) return;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
}

/**
 * Get current step completion status for sequential validation
 * @returns {Object} Map of step IDs to boolean completion status
 */
function getStepCompletionStatus() {
    const status = {};
    Object.keys(STEP_CONFIG_FIELDS).forEach(stepIdStr => {
        const id = parseInt(stepIdStr, 10);
        // A step is complete if it has a sign-off date OR is explicitly marked complete
        const hasSignOff = Boolean(configState.steps[id]?.signOffDate);
        const isMarkedComplete = Boolean(configState.completes[id]);
        status[id] = hasSignOff || isMarkedComplete;
    });
    return status;
}

function bindSignoffToggle(stepId, { buttonId, inputId, canActivate = null, onComplete = null }) {
    const button = document.getElementById(buttonId);
    if (!button) return;
    const input = document.getElementById(inputId);
    const initial =
        Boolean(configState.steps[stepId]?.signOffDate) || Boolean(configState.completes[stepId]);
    updateActionToggleState(button, initial);
    button.addEventListener("click", () => {
        // Check sequential completion (only when trying to activate, not deactivate)
        const isCurrentlyActive = button.classList.contains("is-active");
        if (!isCurrentlyActive && stepId > 0) {
            const completionStatus = getStepCompletionStatus();
            const { canComplete, message } = canCompleteStep(stepId, completionStatus);
            if (!canComplete) {
                showBlockedToast(message);
                return;
            }
        }
        
        if (typeof canActivate === "function" && !canActivate()) return;
        const next = !button.classList.contains("is-active");
        updateActionToggleState(button, next);
        if (input) {
            input.value = next ? todayIso() : "";
            saveStepField(stepId, "signOffDate", input.value);
        }
        saveCompletionFlag(stepId, next);
        if (next) {
            window.scrollTo({ top: 0, behavior: "smooth" });
        }
        if (next && typeof onComplete === "function") {
            onComplete();
        }
    });
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function escapeAttr(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function parseBooleanFlag(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "true" || normalized === "y" || normalized === "yes" || normalized === "1";
}

function parseBooleanStrict(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "true" || normalized === "y" || normalized === "yes" || normalized === "1";
}

function parseDateInput(value) {
    if (!value) return null;
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value));
    if (!match) return null;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    if (!year || !month || !day) return null;
    return { year, month, day };
}

function formatDateInput(value) {
    if (!value) return "";
    const parts = parseDateInput(value);
    if (!parts) return "";
    const { year, month, day } = parts;
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}

function formatDateFromDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function deriveAccountingPeriod(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    return `${monthNames[parts.month - 1]} ${parts.year}`;
}

function deriveJournalId(payrollDate) {
    const parts = parseDateInput(payrollDate);
    if (!parts) return "";
    return `PTO-AUTO-${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function todayIso() {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, "0");
    const d = String(now.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
}

function parsePermanentFlag(value) {
    const normalized = String(value ?? "").trim().toLowerCase();
    return normalized === "y" || normalized === "yes" || normalized === "true" || normalized === "t" || normalized === "1";
}

function coerceTimestamp(value) {
    if (value instanceof Date) return value.getTime();
    if (typeof value === "number") {
        const date = convertExcelSerialDate(value);
        return date ? date.getTime() : null;
    }
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed.getTime();
}

function convertExcelSerialDate(serial) {
    if (!Number.isFinite(serial)) return null;
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

function persistConfigBasics() {
    const getVal = (id) => document.getElementById(id)?.value?.trim() || "";
    const fields = [
        { id: "config-payroll-date", field: PTO_CONFIG_FIELDS.payrollDate },
        { id: "config-accounting-period", field: PTO_CONFIG_FIELDS.accountingPeriod },
        { id: "config-journal-id", field: PTO_CONFIG_FIELDS.journalEntryId },
        { id: "config-company-name", field: PTO_CONFIG_FIELDS.companyName },
        { id: "config-payroll-provider", field: PTO_CONFIG_FIELDS.payrollProvider },
        { id: "config-accounting-link", field: PTO_CONFIG_FIELDS.accountingSoftware },
        { id: "config-user-name", field: PTO_CONFIG_FIELDS.reviewerName }
    ];
    fields.forEach(({ id, field }) => {
        const value = getVal(id);
        if (!field) return;
        scheduleConfigWrite(field, value);
    });
}

function findColumnIndex(headers, keywords = []) {
    const normalizedKeywords = keywords.map((k) => normalizeName(k));
    return headers.findIndex((header) =>
        normalizedKeywords.some((keyword) => header.includes(keyword))
    );
}

function renderHeadcountStep(detail) {
    const stepFields = getStepConfig(2);
    const stepNotes = stepFields?.notes || "";
    const stepNotesPermanent = Boolean(configState.permanents[2]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[2]) || stepSignOff);
    const roster = headcountState.roster || {};
    const hasRun = headcountState.hasAnalyzed;
    const rosterDiff = headcountState.roster?.difference ?? 0;
    const requiresNotes = !headcountState.skipAnalysis && Math.abs(rosterDiff) > 0;
    const rosterCount = roster.rosterCount ?? 0;
    const payrollCount = roster.payrollCount ?? 0;
    const diffValue = roster.difference ?? payrollCount - rosterCount;
    const mismatchList = Array.isArray(roster.mismatches) ? roster.mismatches.filter(Boolean) : [];
    
    // Status banner
    let statusBanner = "";
    if (headcountState.loading) {
        statusBanner = window.PrairieForge?.renderStatusBanner?.({
            type: "info",
            message: "Analyzing headcount…",
            escapeHtml
        }) || "";
    } else if (headcountState.lastError) {
        statusBanner = window.PrairieForge?.renderStatusBanner?.({
            type: "error",
            message: headcountState.lastError,
            escapeHtml
        }) || "";
    }
    
    // Build check rows (circle + pill format)
    const renderCheckRow = (label, desc, value, isMatch) => {
        const pending = !hasRun;
        let circleHtml;
        
        if (pending) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (isMatch) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        const valueDisplay = hasRun ? ` = ${value}` : "";
        
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(label)}${valueDisplay}</span>
            </div>
        `;
    };
    
    const checkRowsHtml = `
        ${renderCheckRow("SS_Employee_Roster count", "Active employees in roster", rosterCount, true)}
        ${renderCheckRow("PTO_Data_Clean count", "Unique employees in PTO data", payrollCount, true)}
        ${renderCheckRow("Difference", "Should be zero", diffValue, diffValue === 0)}
    `;
    
    // Mismatch section
    const mismatchSection =
        mismatchList.length && !headcountState.skipAnalysis && hasRun
            ? window.PrairieForge.renderMismatchTiles({
                  mismatches: mismatchList,
                  label: "Employees Driving the Difference",
                  sourceLabel: "Roster",
                  targetLabel: "PTO Data",
                  escapeHtml: escapeHtml
              })
            : "";
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
            <div class="pf-skip-action">
                <button type="button" class="pf-skip-btn ${headcountState.skipAnalysis ? "is-active" : ""}" id="headcount-skip-btn">
                    ${X_CIRCLE_SVG}
                    <span>Skip Analysis</span>
                </button>
            </div>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Headcount Check</h3>
                    <p class="pf-config-subtext">Compare employee roster against PTO data to identify discrepancies.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-run-btn" title="Run headcount analysis">${CALCULATOR_ICON_SVG}</button>`, "Run")}
                    ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="headcount-refresh-btn" title="Refresh headcount analysis">${REFRESH_ICON_SVG}</button>`, "Refresh")}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Headcount Comparison</h3>
                    <p class="pf-config-subtext">Verify roster and payroll data align before proceeding.</p>
                </div>
                ${statusBanner}
                <div class="pf-je-checks-container">
                    ${checkRowsHtml}
                </div>
                ${mismatchSection}
            </article>
            ${renderInlineNotes({
                textareaId: "step-notes-input",
                value: stepNotes,
                permanentId: "step-notes-lock-2",
                isPermanent: stepNotesPermanent,
                hintId: requiresNotes ? "headcount-notes-hint" : "",
                saveButtonId: "step-notes-save-2"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-name",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-date",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "headcount-signoff-save",
                completeButtonId: "headcount-signoff-toggle"
            })}
        </section>
    `;
}

/**
 * Render unified Apple-inspired Data Readiness card
 * Combines completeness check pills + missing pay rate resolution
 */
function renderDataReadinessCard() {
    const check = analysisState.completenessCheck || {};
    const missing = analysisState.missingPayRates || [];
    
    // Completeness check fields with descriptions
    const fields = [
        { key: "accrualRate", label: "Accrual Rate", desc: "∑ PTO_Data_Clean = ∑ PTO_Analysis" },
        { key: "carryOver", label: "Carry Over", desc: "∑ PTO_Data_Clean = ∑ PTO_Analysis" },
        { key: "ytdAccrued", label: "YTD Accrued", desc: "∑ PTO_Data_Clean = ∑ PTO_Analysis" },
        { key: "ytdUsed", label: "YTD Used", desc: "∑ PTO_Data_Clean = ∑ PTO_Analysis" },
        { key: "balance", label: "Balance", desc: "∑ PTO_Data_Clean = ∑ PTO_Analysis" }
    ];
    
    // Calculate overall status
    const allChecked = fields.every(f => check[f.key] !== null && check[f.key] !== undefined);
    const allPassed = allChecked && fields.every(f => check[f.key]?.match);
    const hasMissingRates = missing.length > 0;
    
    // Build check rows (circle + pill format like JE tab)
    const renderCheckRow = (field) => {
        const result = check[field.key];
        const pending = result === null || result === undefined;
        let circleHtml;
        
        if (pending) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (result.match) {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(field.label)}: ${escapeHtml(field.desc)}</span>
            </div>
        `;
    };
    
    const checkRowsHtml = fields.map(f => renderCheckRow(f)).join("");
    
    // Missing pay rate section
    let missingSection = "";
    if (hasMissingRates) {
        const employee = missing[0];
        const remainingCount = missing.length - 1;
        
        missingSection = `
            <div class="pf-readiness-divider"></div>
            <div class="pf-readiness-issue">
                <div class="pf-readiness-issue-header">
                    <span class="pf-readiness-issue-badge">Action Required</span>
                    <span class="pf-readiness-issue-title">Missing Pay Rate</span>
                </div>
                <p class="pf-readiness-issue-desc">
                    Enter hourly rate for <strong>${escapeHtml(employee.name)}</strong> to calculate liability
                </p>
                <div class="pf-readiness-input-row">
                    <div class="pf-readiness-input-field">
                        <span class="pf-readiness-input-prefix">$</span>
                        <input type="number" 
                               id="payrate-input" 
                               class="pf-readiness-input" 
                               placeholder="0.00" 
                               step="0.01"
                               min="0"
                               data-employee="${escapeAttr(employee.name)}"
                               data-row="${employee.rowIndex}">
                    </div>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--secondary" id="payrate-ignore-btn">
                        Skip
                    </button>
                    <button type="button" class="pf-readiness-btn pf-readiness-btn--primary" id="payrate-save-btn">
                        Save
                    </button>
                </div>
                ${remainingCount > 0 ? `<p class="pf-readiness-remaining">${remainingCount} more employee${remainingCount > 1 ? "s" : ""} need pay rates</p>` : ""}
            </div>
        `;
    }
    
    return `
        <article class="pf-step-card pf-step-detail pf-config-card" id="data-readiness-card">
            <div class="pf-config-head">
                <h3>Data Completeness</h3>
                <p class="pf-config-subtext">Quick check that all your data transferred correctly.</p>
            </div>
            <div class="pf-je-checks-container">
                ${checkRowsHtml}
            </div>
            ${missingSection}
        </article>
    `;
}

function renderDataQualityStep(detail) {
    const stepFields = getStepConfig(3);
    const notesPermanent = Boolean(configState.permanents[3]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[3]) || stepSignOff);
    
    // Build quality check results display
    const hasRun = dataQualityState.hasRun;
    const { balanceIssues, zeroBalances, accrualOutliers, totalEmployees } = dataQualityState;
    
    // Status banner
    let statusBanner = "";
    if (dataQualityState.loading) {
        statusBanner = window.PrairieForge?.renderStatusBanner?.({
            type: "info",
            message: "Analyzing data quality...",
            escapeHtml
        }) || "";
    } else if (hasRun) {
        const criticalCount = balanceIssues.length;
        const warningCount = accrualOutliers.length + zeroBalances.length;
        
        if (criticalCount > 0) {
            statusBanner = window.PrairieForge?.renderStatusBanner?.({
                type: "error",
                title: `${criticalCount} Balance Issue${criticalCount > 1 ? "s" : ""} Found`,
                message: "Review the issues below. Fix in PTO_Data_Clean and re-run, or acknowledge to continue.",
                escapeHtml
            }) || "";
        } else if (warningCount > 0) {
            statusBanner = window.PrairieForge?.renderStatusBanner?.({
                type: "warning",
                title: "No Critical Issues",
                message: `${warningCount} informational item${warningCount > 1 ? "s" : ""} to review (see below).`,
                escapeHtml
            }) || "";
        } else {
            statusBanner = window.PrairieForge?.renderStatusBanner?.({
                type: "success",
                title: "Data Quality Passed",
                message: `${totalEmployees} employee${totalEmployees !== 1 ? "s" : ""} checked — no anomalies found.`,
                escapeHtml
            }) || "";
        }
    }
    
    // Build issue cards
    const issueCards = [];
    
    if (hasRun && balanceIssues.length > 0) {
        issueCards.push(`
            <div class="pf-quality-issue pf-quality-issue--critical">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">⚠️</span>
                    <span class="pf-quality-issue-title">Balance Issues (${balanceIssues.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${balanceIssues.slice(0, 5).map(e => 
                        `<li><strong>${escapeHtml(e.name)}</strong>: ${escapeHtml(e.issue)}</li>`
                    ).join("")}
                    ${balanceIssues.length > 5 ? `<li class="pf-quality-more">+${balanceIssues.length - 5} more</li>` : ""}
                </ul>
            </div>
        `);
    }
    
    if (hasRun && accrualOutliers.length > 0) {
        issueCards.push(`
            <div class="pf-quality-issue pf-quality-issue--warning">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">📊</span>
                    <span class="pf-quality-issue-title">High Accrual Rates (${accrualOutliers.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${accrualOutliers.slice(0, 5).map(e => 
                        `<li><strong>${escapeHtml(e.name)}</strong>: ${e.accrualRate.toFixed(2)} hrs/period</li>`
                    ).join("")}
                    ${accrualOutliers.length > 5 ? `<li class="pf-quality-more">+${accrualOutliers.length - 5} more</li>` : ""}
                </ul>
            </div>
        `);
    }
    
    if (hasRun && zeroBalances.length > 0) {
        issueCards.push(`
            <div class="pf-quality-issue pf-quality-issue--info">
                <div class="pf-quality-issue-header">
                    <span class="pf-quality-issue-icon">ℹ️</span>
                    <span class="pf-quality-issue-title">Zero Balances (${zeroBalances.length})</span>
                </div>
                <ul class="pf-quality-issue-list">
                    ${zeroBalances.slice(0, 5).map(e => 
                        `<li><strong>${escapeHtml(e.name)}</strong></li>`
                    ).join("")}
                    ${zeroBalances.length > 5 ? `<li class="pf-quality-more">+${zeroBalances.length - 5} more</li>` : ""}
                </ul>
            </div>
        `);
    }
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Run Quality Check</h3>
                    <p class="pf-config-subtext">Scan your imported data for common errors before proceeding.</p>
                </div>
                ${statusBanner}
                <div class="pf-signoff-action">
                    ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-run-btn" title="Run data quality checks">${CALCULATOR_ICON_SVG}</button>`, "Run")}
                </div>
            </article>
            ${issueCards.length > 0 ? `
                <article class="pf-step-card pf-step-detail">
                    <div class="pf-config-head">
                        <h3>Issues Found</h3>
                        <p class="pf-config-subtext">Fix issues in PTO_Data_Clean and re-run, or acknowledge to continue.</p>
                    </div>
                    <div class="pf-quality-issues-grid">
                        ${issueCards.join("")}
                    </div>
                    <div class="pf-quality-actions-bar">
                        ${dataQualityState.acknowledged 
                            ? `<p class="pf-quality-actions-hint"><span class="pf-acknowledged-badge">✓ Issues Acknowledged</span></p>` 
                            : ""}
                        <div class="pf-signoff-action">
                            ${renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-refresh-btn" title="Re-run quality checks">${REFRESH_ICON_SVG}</button>`, "Refresh")}
                            ${!dataQualityState.acknowledged ? renderLabeledButton(`<button type="button" class="pf-action-toggle pf-clickable" id="quality-acknowledge-btn" title="Acknowledge issues and continue">${CHECK_ICON_SVG}</button>`, "Continue") : ""}
                        </div>
                    </div>
                </article>
            ` : ""}
            ${renderInlineNotes({
                textareaId: "step-notes-3",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-3",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-3"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-3",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-3",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-3",
                completeButtonId: "step-signoff-toggle-3"
            })}
        </section>
    `;
}

/**
 * Step 2: PTO Accrual Review
 * Variance table showing current vs prior period with liability calculations
 */
function renderAccrualReviewStep(detail) {
    // Step 2 in new structure
    const stepFields = getStepConfig(2);
    const notesPermanent = Boolean(configState.permanents[2]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[2]) || stepSignOff);
    
    // Executive summary and reconciliation values
    const reviewState = ptoReviewState || {};
    const recon = reviewState.reconciliation || {};
    const coverage = reviewState.coverage || {};
    const totalPriorLiability = reviewState.totalPriorLiability || 0;
    const employeeCount = reviewState.employeeCount || 0;
    
    // Reconciliation values
    const reportLiability = recon.reportLiabilityTotal || 0;
    const calcLiability = recon.calcLiabilityTotal || 0;
    const negativeBalanceTotal = recon.negativeBalanceTotal || 0;
    const negativeBalanceCount = recon.negativeBalanceCount || 0;
    const missingRateCount = recon.missingRateCount || 0;
    
    // Net change is based on calculated liability (includes negatives) vs prior
    const netChange = calcLiability - totalPriorLiability;
    
    // Format currency helper
    const fmtCurrency = (val) => {
        const num = Number(val) || 0;
        return num.toLocaleString("en-US", { style: "currency", currency: "USD" });
    };
    
    // Build the reconciliation summary card
    const hasReconData = reviewState.loaded && employeeCount > 0;
    
    const summaryCard = `
        <article class="pf-step-card pf-step-detail pf-config-card">
            <div class="pf-config-head">
                <h3>PTO Liability Reconciliation</h3>
                <p class="pf-config-subtext">Report → Adjustments → JE Amount</p>
            </div>
            
            ${hasReconData ? `
            <!-- Liability Reconciliation Section -->
            <div class="pf-liability-recon" style="margin-top: 16px; padding: 16px; background: rgba(255,255,255,0.03); border-radius: 8px;">
                <div class="pf-recon-row" style="display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <span style="color: rgba(255,255,255,0.7);">Report Liability (PrismHR)</span>
                    <span style="font-weight: 600; color: #fff;">${fmtCurrency(reportLiability)}</span>
                </div>
                
                ${negativeBalanceCount > 0 ? `
                <div class="pf-recon-row" style="display: flex; justify-content: space-between; padding: 8px 0; padding-left: 20px; border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <span style="color: rgba(255,255,255,0.5); font-size: 13px;">Negative Balances (${negativeBalanceCount} employees)</span>
                    <span style="color: #f87171; font-size: 13px;">${fmtCurrency(negativeBalanceTotal)}</span>
                </div>
                ` : ''}
                
                <div class="pf-recon-row" style="display: flex; justify-content: space-between; padding: 12px 0; border-top: 2px solid rgba(255,255,255,0.2); margin-top: 4px;">
                    <span style="font-weight: 600; color: #fff;">Calculated Liability</span>
                    <span style="font-weight: 600; color: #fff;">${fmtCurrency(calcLiability)}</span>
                </div>
                
                <div class="pf-recon-row" style="display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid rgba(255,255,255,0.1);">
                    <span style="color: rgba(255,255,255,0.5);">Prior Period Liability</span>
                    <span style="color: rgba(255,255,255,0.7);">${fmtCurrency(totalPriorLiability)}</span>
                </div>
                
                <div class="pf-recon-row" style="display: flex; justify-content: space-between; padding: 12px 0; border-top: 2px solid rgba(255,255,255,0.2); margin-top: 4px;">
                    <span style="font-weight: 600; color: #fff;">Net Change (for JE)</span>
                    <span style="font-weight: 700; font-size: 18px; color: ${netChange >= 0 ? '#4ade80' : '#f87171'};">
                        ${netChange >= 0 ? '+' : ''}${fmtCurrency(netChange)}
                    </span>
                    </div>
                </div>
            
            <!-- Employee Coverage Section -->
            <div class="pf-coverage-section" style="margin-top: 16px; padding: 16px; background: rgba(255,255,255,0.02); border-radius: 8px;">
                <h4 style="margin: 0 0 12px 0; font-size: 14px; color: rgba(255,255,255,0.8);">Employee Coverage</h4>
                
                <div class="pf-coverage-row" style="display: flex; justify-content: space-between; padding: 6px 0;">
                    <span style="color: rgba(255,255,255,0.6);">Roster Headcount:</span>
                    <span style="color: #fff;">${coverage.rosterCount || 0}</span>
                </div>
                
                ${coverage.inPtoOnlyCount > 0 ? `
                <div class="pf-coverage-row" style="display: flex; justify-content: space-between; padding: 6px 0; padding-left: 20px;">
                    <span style="color: rgba(255,255,255,0.5); font-size: 13px;">+ In PTO only (not on roster):</span>
                    <span style="color: rgba(255,255,255,0.7); font-size: 13px;">+${coverage.inPtoOnlyCount}</span>
            </div>
                <div style="padding-left: 40px; font-size: 12px; color: rgba(255,255,255,0.4);">
                    (Liability: ${fmtCurrency(coverage.inPtoOnlyLiability || 0)})
                </div>
                ` : ''}
                
                ${coverage.inRosterOnlyCount > 0 ? `
                <div class="pf-coverage-row" style="display: flex; justify-content: space-between; padding: 6px 0; padding-left: 20px;">
                    <span style="color: rgba(255,255,255,0.5); font-size: 13px;">− On roster, not in PTO:</span>
                    <span style="color: rgba(255,255,255,0.7); font-size: 13px;">−${coverage.inRosterOnlyCount}</span>
                </div>
                ` : ''}
                
                <div class="pf-coverage-row" style="display: flex; justify-content: space-between; padding: 8px 0; border-top: 1px solid rgba(255,255,255,0.15); margin-top: 8px;">
                    <span style="font-weight: 600; color: #fff;">PTO Report Headcount:</span>
                    <span style="font-weight: 600; color: #fff;">${employeeCount}</span>
                </div>
            </div>
            ` : `
            <div style="margin-top: 16px; padding: 20px; text-align: center; color: rgba(255,255,255,0.5);">
                <p>Click "Generate" to calculate liability reconciliation</p>
            </div>
            `}
            
            ${missingRateCount > 0 ? `
                <div style="margin-top: 12px; padding: 10px; background: rgba(251, 191, 36, 0.1); border: 1px solid rgba(251, 191, 36, 0.3); border-radius: 8px; font-size: 13px;">
                    <strong style="color: #fbbf24;">${missingRateCount}</strong> employee(s) missing pay rate
                </div>
            ` : ''}
        </article>
    `;

    // Render Ada chat interface for PTO analysis
    const adaMarkup = renderCopilotCard({
        id: "pto-copilot",
        heading: "Ada",
        subtext: "Ask questions about PTO data, liabilities, and insights",
        welcomeMessage: "What would you like to explore about PTO accruals?",
        placeholder: "Ask about PTO balances, accrual rates, or liability trends...",
        quickActions: [
            { id: "diagnostics", label: "Run Diagnostics", prompt: "Run a diagnostic check on the PTO data. Check for completeness, missing rates, and data quality issues." },
            { id: "insights", label: "Generate Insights", prompt: "What are the key insights and findings from this PTO accrual analysis?" },
            { id: "balances", label: "Balance Analysis", prompt: "Analyze PTO balances by employee and department. Highlight any concerning balances." },
            { id: "accruals", label: "Accrual Trends", prompt: "Show PTO accrual trends and changes from the prior period." }
        ],
        contextProvider: createExcelContextProvider({
            dataClean: 'PTO_Data_Clean',
            analysis: 'PTO_Analysis',
            review: 'PTO_Review',
            config: 'SS_PF_Config'
        }),
        onPrompt: callAdaApi
    });
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Review</h3>
                    <p class="pf-config-subtext">Calculate liabilities, load prior period, and compute variance.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="review-generate-btn" title="Generate PTO review table">${CALCULATOR_ICON_SVG}</button>`,
                        "Generate"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="review-open-btn" title="Open PTO_Review sheet">${TABLE_ICON_SVG}</button>`,
                        "Open Sheet"
                    )}
                </div>
            </article>
            ${summaryCard}
            ${adaMarkup}
            ${renderInlineNotes({
                textareaId: "step-notes-2",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-2",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-2"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-2",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-2",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-2",
                completeButtonId: "step-signoff-toggle-2"
            })}
        </section>
    `;
}

function renderJournalStep(detail) {
    // Step 3 in new structure
    const stepFields = getStepConfig(3);
    const notesPermanent = Boolean(configState.permanents[3]);
    const stepReviewer = getReviewerWithFallback(stepFields?.reviewer);
    const stepSignOff = stepFields?.signOffDate || "";
    const stepComplete = Boolean(parseBooleanStrict(configState.completes[3]) || stepSignOff);
    const statusNote = journalState.lastError
        ? `<p class="pf-step-note">${escapeHtml(journalState.lastError)}</p>`
        : "";
    
    // Build validation check rows: circle icon on left, description pill on right
    const hasRun = journalState.validationRun;
    const issues = journalState.issues || [];
    
    // Define check descriptions (what's being calculated)
    const checkDefinitions = [
        { key: "Debits = Credits", desc: "∑ Debits = ∑ Credits" },
        { key: "Line Amounts Sum to Zero", desc: "∑ Line Amount = $0.00" },
        { key: "JE Matches Analysis Total", desc: "∑ JE expense = ∑ PTO_Analysis Change" }
    ];
    
    const renderCheckRow = (def) => {
        const issue = issues.find(i => i.check === def.key);
        const pending = !hasRun;
        let circleHtml;
        
        if (pending) {
            // Empty circle - waiting for validation
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pending"></span>`;
        } else if (issue?.passed) {
            // Checkmark in circle
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--pass">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"/></svg>
            </span>`;
        } else {
            // X in circle
            circleHtml = `<span class="pf-je-check-circle pf-je-circle--fail">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/></svg>
            </span>`;
        }
        
        return `
            <div class="pf-je-check-row">
                ${circleHtml}
                <span class="pf-je-check-desc-pill">${escapeHtml(def.desc)}</span>
            </div>
        `;
    };
    
    const checkRows = checkDefinitions.map(def => renderCheckRow(def)).join("");
    
    // Build issues card if there are failures
    const failedIssues = issues.filter(i => !i.passed);
    let issuesCard = "";
    if (hasRun && failedIssues.length > 0) {
        issuesCard = `
            <article class="pf-step-card pf-step-detail pf-je-issues-card">
                <div class="pf-config-head">
                    <h3>⚠️ Issues Identified</h3>
                    <p class="pf-config-subtext">The following checks did not pass:</p>
                </div>
                <ul class="pf-je-issues-list">
                    ${failedIssues.map(i => `<li><strong>${escapeHtml(i.check)}:</strong> ${escapeHtml(i.detail)}</li>`).join("")}
                </ul>
            </article>
        `;
    }
    
    return `
        <section class="pf-hero" id="pf-step-hero">
            <p class="pf-hero-copy">${escapeHtml(MODULE_NAME)} | Step ${detail.id}</p>
            <h2 class="pf-hero-title">${escapeHtml(detail.title)}</h2>
            <p class="pf-hero-copy">${escapeHtml(detail.summary || "")}</p>
        </section>
        <section class="pf-step-guide">
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Generate Journal Entry</h3>
                    <p class="pf-config-subtext">Create a balanced JE from your imported PTO data, grouped by department.</p>
                </div>
                <div class="pf-signoff-action">
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-create-btn" title="Generate journal entry from PTO_Analysis">${TABLE_ICON_SVG}</button>`,
                        "Generate"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-run-btn" title="Refresh validation checks">${REFRESH_ICON_SVG}</button>`,
                        "Refresh"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-export-btn" title="Export journal draft as CSV">${DOWNLOAD_ICON_SVG}</button>`,
                        "Export"
                    )}
                    ${renderLabeledButton(
                        `<button type="button" class="pf-action-toggle pf-clickable" id="je-upload-btn" title="Open accounting software upload">${UPLOAD_ICON_SVG}</button>`,
                        "Upload"
                    )}
                </div>
            </article>
            <article class="pf-step-card pf-step-detail pf-config-card">
                <div class="pf-config-head">
                    <h3>Validation Checks</h3>
                    <p class="pf-config-subtext">These checks run automatically after generating your JE.</p>
                </div>
                ${statusNote}
                <div class="pf-je-checks-container">
                    ${checkRows}
                </div>
            </article>
            ${issuesCard}
            ${renderInlineNotes({
                textareaId: "step-notes-3",
                value: stepFields?.notes || "",
                permanentId: "step-notes-lock-3",
                isPermanent: notesPermanent,
                hintId: "",
                saveButtonId: "step-notes-save-3"
            })}
            ${renderSignoff({
                reviewerInputId: "step-reviewer-3",
                reviewerValue: stepReviewer,
                signoffInputId: "step-signoff-3",
                signoffValue: stepSignOff,
                isComplete: stepComplete,
                saveButtonId: "step-signoff-save-3",
                completeButtonId: "step-signoff-toggle-3"
            })}
        </section>
    `;
}

function headcountHasDifferences() {
    const rosterDiff = Math.abs(headcountState.roster?.difference ?? 0);
    return rosterDiff > 0;
}

function isHeadcountNotesRequired() {
    return !headcountState.skipAnalysis && headcountHasDifferences();
}

function formatMetricValue(value) {
    if (value === null || value === undefined || value === "") return "---";
    const num = Number(value);
    const formatter = window.PrairieForge?.formatNumber;
    if (Number.isFinite(num)) {
        return formatter ? formatter(num) : num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(value);
}

function formatNumberDisplay(value) {
    if (value === null || value === undefined || value === "") return "---";
    const num = Number(value);
    const formatter = window.PrairieForge?.formatNumber;
    if (Number.isFinite(num)) {
        return formatter ? formatter(num) : num.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }
    return String(value);
}

function formatSignedValue(value) {
    if (value === null || value === undefined) return "---";
    if (typeof value !== "number" || Number.isNaN(value)) return "---";
    if (value === 0) return "0";
    return value > 0 ? `+${value}` : value.toString();
}

async function refreshHeadcountAnalysis() {
    if (!hasExcelRuntime()) {
        headcountState.loading = false;
        headcountState.lastError = "Excel runtime is unavailable.";
        renderApp();
        return;
    }
    headcountState.loading = true;
    headcountState.lastError = null;
    // Reset save button since data is being refreshed
    updateSaveButtonState(document.getElementById("headcount-save-btn"), false);
    renderApp();
    try {
        const results = await Excel.run(async (context) => {
            // Use SS_Employee_Roster as the single source of truth for employee data
            const rosterSheet = context.workbook.worksheets.getItem("SS_Employee_Roster");
            
            // Use PTO_Data_Clean as the PTO data source
            const payrollSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Data_Clean");
            payrollSheet.load("isNullObject");
            await context.sync();
            
            if (payrollSheet.isNullObject) {
                console.warn("[Headcount] PTO_Data_Clean not found");
                return { rosterEmployees: [], ptoEmployees: [], missingFromPto: [], extraInPto: [] };
            }
            
            const analysisSheet = context.workbook.worksheets.getItemOrNullObject("PTO_Analysis");
            const rosterRange = rosterSheet.getUsedRangeOrNullObject();
            const payrollRange = payrollSheet.getUsedRangeOrNullObject();
            rosterRange.load("values");
            payrollRange.load("values");
            analysisSheet.load("isNullObject");
            await context.sync();
            let analysisRange = null;
            if (!analysisSheet.isNullObject) {
                analysisRange = analysisSheet.getUsedRangeOrNullObject();
                analysisRange.load("values");
            }
            await context.sync();
            const rosterValues = rosterRange.isNullObject ? [] : rosterRange.values || [];
            const payrollValues = payrollRange.isNullObject ? [] : payrollRange.values || [];
            const analysisValues = analysisRange && !analysisRange.isNullObject ? analysisRange.values || [] : [];
            // Prefer PTO_Analysis if it has data, otherwise fall back to PTO_Data
            const payrollSource = analysisValues.length ? analysisValues : payrollValues;
            return parseHeadcount(rosterValues, payrollSource);
        });
        // Populate structured state fields
        headcountState.rosterCount = results.rosterCount;
        headcountState.ptoCount = results.ptoCount;
        headcountState.missingFromPto = results.missingFromPto;
        headcountState.extraInPto = results.extraInPto;
        // Legacy format for backward compatibility
        headcountState.roster = results.roster;
        headcountState.hasAnalyzed = true;
        headcountState.lastError = null;
    } catch (error) {
        console.warn("PTO headcount: unable to analyze data", error);
        headcountState.lastError = "Unable to analyze headcount data. Try re-running the analysis.";
    } finally {
        headcountState.loading = false;
        renderApp();
    }
}

/**
 * Check if a value looks like a summary row (total, subtotal, etc.)
 * @param {string} value - Value to check
 * @returns {boolean} True if it should be excluded
 */
function isSummaryOrEmpty(value) {
    if (!value) return true;
    const lower = value.toLowerCase().trim();
    if (!lower) return true;
    const summaryPatterns = ["total", "subtotal", "sum", "count", "grand", "average", "avg"];
    return summaryPatterns.some((pattern) => lower.includes(pattern));
}

function parseHeadcount(rosterValues, payrollValues) {
    const result = {
        rosterCount: 0,
        ptoCount: 0,
        missingFromPto: [],    // In roster but not in PTO: [{name, department}]
        extraInPto: [],        // In PTO but not in roster: [{name}]
        // Legacy format for backward compatibility
        roster: {
            rosterCount: 0,
            payrollCount: 0,
            difference: 0,
            mismatches: []
        }
    };

    // Require at least header + 1 data row
    if ((rosterValues?.length || 0) < 2 || (payrollValues?.length || 0) < 2) {
        console.warn("Headcount: insufficient data rows", {
            rosterRows: rosterValues?.length || 0,
            payrollRows: payrollValues?.length || 0
        });
        return result;
    }

    const rosterHeaderInfo = findHeaderRow(rosterValues);
    const payrollHeaderInfo = findHeaderRow(payrollValues);

    const rosterHeaders = rosterHeaderInfo.headers;
    const payrollHeaders = payrollHeaderInfo.headers;

    const rosterIdx = {
        employee: getEmployeeColumnIndex(rosterHeaders),
        department: rosterHeaders.findIndex((h) => h.includes("department")),
        status: rosterHeaders.findIndex((h) => h === "employment_status" || h === "status"),
        termination: rosterHeaders.findIndex((h) => h.includes("termination"))
    };
    const payrollIdx = {
        employee: getEmployeeColumnIndex(payrollHeaders)
    };

    // Log column detection for debugging
    console.log("Headcount column detection:", {
        rosterEmployeeCol: rosterIdx.employee,
        rosterDeptCol: rosterIdx.department,
        rosterStatusCol: rosterIdx.status,
        rosterTerminationCol: rosterIdx.termination,
        payrollEmployeeCol: payrollIdx.employee,
        rosterHeaders: rosterHeaders.slice(0, 5),
        payrollHeaders: payrollHeaders.slice(0, 5)
    });

    // Build roster map with details (only active employees)
    const rosterMap = new Map(); // key -> {name, department}
    for (let i = rosterHeaderInfo.startIndex; i < rosterValues.length; i += 1) {
        const row = rosterValues[i];
        const employee = rosterIdx.employee >= 0 ? normalizeString(row[rosterIdx.employee]) : "";
        if (isSummaryOrEmpty(employee)) continue;
        
        // Skip terminated employees (check both status and termination date)
        if (rosterIdx.status >= 0) {
            const status = normalizeString(row[rosterIdx.status]).toLowerCase();
            if (status === "terminated" || status === "inactive" || status === "term") continue;
        }
        if (rosterIdx.termination >= 0) {
            const termination = normalizeString(row[rosterIdx.termination]);
            if (termination) continue;
        }
        
        const key = employee.toLowerCase();
        if (!rosterMap.has(key)) {
            rosterMap.set(key, {
                name: employee,
                department: rosterIdx.department >= 0 ? normalizeString(row[rosterIdx.department]) : ""
            });
        }
    }

    // Build PTO employee set
    const ptoMap = new Map(); // key -> {name}
    for (let i = payrollHeaderInfo.startIndex; i < payrollValues.length; i += 1) {
        const row = payrollValues[i];
        const employee = payrollIdx.employee >= 0 ? normalizeString(row[payrollIdx.employee]) : "";
        if (isSummaryOrEmpty(employee)) continue;
        
        const key = employee.toLowerCase();
        if (!ptoMap.has(key)) {
            ptoMap.set(key, { name: employee });
        }
    }

    result.rosterCount = rosterMap.size;
    result.ptoCount = ptoMap.size;

    // Find mismatches with structured data
    // Active roster employees not in PTO data
    rosterMap.forEach((data, key) => {
        if (!ptoMap.has(key)) {
            result.missingFromPto.push({
                name: data.name,
                department: data.department || "—"
            });
        }
    });

    // PTO employees not in roster (potential new hires or data issues)
    ptoMap.forEach((data, key) => {
        if (!rosterMap.has(key)) {
            result.extraInPto.push({
                name: data.name
            });
        }
    });

    console.log("Headcount results:", {
        rosterCount: result.rosterCount,
        ptoCount: result.ptoCount,
        missingFromPto: result.missingFromPto.length,
        extraInPto: result.extraInPto.length
    });

    // Populate legacy format for backward compatibility
    result.roster.rosterCount = result.rosterCount;
    result.roster.payrollCount = result.ptoCount;
    result.roster.difference = result.ptoCount - result.rosterCount;
    result.roster.mismatches = [
        ...result.missingFromPto.map((e) => `In roster, missing in PTO_Data_Clean: ${e.name}`),
        ...result.extraInPto.map((e) => `In PTO_Data_Clean, missing in roster: ${e.name}`)
    ];

    return result;
}

function findHeaderRow(values) {
    if (!Array.isArray(values) || !values.length) {
        return { headers: [], startIndex: 1 };
    }
    const headerRowIndex = values.findIndex((row = []) =>
        row.some((cell) => {
            const normalized = normalizeString(cell).toLowerCase();
            return normalized.includes("employee");
        })
    );
    const index = headerRowIndex === -1 ? 0 : headerRowIndex;
    const headers = (values[index] || []).map((h) => normalizeString(h).toLowerCase());
    return { headers, startIndex: index + 1 };
}

function getEmployeeColumnIndex(headers = []) {
    let bestIndex = -1;
    let bestScore = -1;
    headers.forEach((header, index) => {
        const value = header.toLowerCase();
        if (!value.includes("employee")) return;
        let score = 1; // baseline: contains "employee"
        if (value.includes("name")) {
            score = 4; // prefer explicit name column
        } else if (value.includes("id")) {
            score = 2; // lower priority than name
        } else {
            score = 3; // generic employee column without name/id hints
        }
        if (score > bestScore) {
            bestScore = score;
            bestIndex = index;
        }
    });
    return bestIndex;
}

function normalizeString(value) {
    return value == null ? "" : String(value).trim();
}

/**
 * Sync PTO_Analysis from PTO_Data_Clean with enrichment from roster, payroll archive, and prior period.
 * PTO_Analysis serves as the enriched view for internal calculations and archive.
 * 
 * Columns:
 * - Analysis Date, Employee Name, Department, Pay Rate, Accrual Rate, Carry Over, YTD Used, Balance
 * - Liability Amount (current period)
 * - Accrued PTO $ [Prior Period] (from PTO_Archive_Summary, 0 if not found)
 * - Change (current - prior)
 */
// REMOVED: syncPtoAnalysis() function - legacy code replaced by PTO_Review workflow
// PTO_Analysis sheet is no longer used. All review functionality is in PTO_Review (Step 2).
async function syncPtoAnalysis() {
    console.warn("[DEPRECATED] syncPtoAnalysis() has been removed. Use generatePtoReview() instead.");
    return;
}

function buildCsv(rows = []) {
    return rows
        .map((row) =>
            (row || [])
                .map((cell) => {
                    if (cell == null) return "";
                    const str = String(cell);
                    if (/[",\n]/.test(str)) return `"${str.replace(/"/g, '""')}"`;
                    return str;
                })
                .join(",")
        )
        .join("\n");
}

function downloadCsv(filename, content) {
    const blob = new Blob([content], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function updateHeadcountSignoffState() {
    const button = document.getElementById("headcount-signoff-toggle");
    if (!button) return;
    const notesRequired = isHeadcountNotesRequired();
    const notesInput = document.getElementById("step-notes-input");
    const notesValue = notesInput?.value.trim() || "";
    button.disabled = notesRequired && !notesValue;
    const hint = document.getElementById("headcount-notes-hint");
    if (hint) {
        hint.textContent = notesRequired
            ? "Please document outstanding differences before signing off."
            : "";
    }
}

function enforceHeadcountSkipNote() {
    const textarea = document.getElementById("step-notes-input");
    if (!textarea) return;
    const current = textarea.value || "";
    const remainder = current.startsWith(HEADCOUNT_SKIP_NOTE)
        ? current.slice(HEADCOUNT_SKIP_NOTE.length).replace(/^\s+/, "")
        : current.replace(new RegExp(`^${HEADCOUNT_SKIP_NOTE}\\s*`, "i"), "").trimStart();
    const next = HEADCOUNT_SKIP_NOTE + (remainder ? `\n${remainder}` : "");
    if (textarea.value !== next) {
        textarea.value = next;
    }
    saveStepField(2, "notes", textarea.value);
}

function bindHeadcountNotesGuard() {
    const textarea = document.getElementById("step-notes-input");
    if (!textarea) return;
    textarea.addEventListener("input", () => {
        if (!headcountState.skipAnalysis) return;
        const value = textarea.value || "";
        if (!value.startsWith(HEADCOUNT_SKIP_NOTE)) {
            const remainder = value.replace(HEADCOUNT_SKIP_NOTE, "").trimStart();
            textarea.value = HEADCOUNT_SKIP_NOTE + (remainder ? `\n${remainder}` : "");
        }
        saveStepField(2, "notes", textarea.value);
    });
}

function handleHeadcountSignoff() {
    const notesRequired = isHeadcountNotesRequired();
    const notesValue = document.getElementById("step-notes-input")?.value.trim() || "";
    if (notesRequired && !notesValue) {
        showToast("Please enter a brief explanation of the outstanding differences before completing this step.", "info");
        return;
    }
}






